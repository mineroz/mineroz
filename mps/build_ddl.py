#!/usr/bin/env python3
"""
build_ddl.py - generate the SQL schema of the Mako Production System (MPS) web phase
from mps_schema.json, the same schema that drives build_workbook.py.

    python3 mps/build_ddl.py

writes, with no arguments and byte-identical output on every run:

    mps/sql/schema_postgres.sql   PostgreSQL 13 or later
    mps/sql/schema_sqlite.sql     SQLite 3.31 or later (generated columns)
    docs/data_dictionary.md       field-level dictionary and the Excel to SQL mapping

Everything about the data model comes from the schema and from the KPI catalogue
in build_workbook.py (KPIS). The only knowledge typed into this file is listed
under MODEL RULES below; each rule is honoured from the schema first when the
schema carries the matching key (sql_table, sql_name, unique, sql_postgres,
sql_sqlite), so it can move into the schema without touching the generator.

Mapping (documented again in the header of each SQL file and in the dictionary):
  * tables: Excel table name without the tbl prefix, CamelCase to snake_case;
    tables prefilled with one row per day get the suffix _daily
    (tblPlant -> plant_daily, tblMovement -> movement, tblMiningDaily -> mining_daily)
  * columns: the Excel field name in lower case (Milled_t -> milled_t)
  * date -> DATE, text -> TEXT, list -> TEXT + FOREIGN KEY, number -> NUMERIC(18,4),
    int -> INTEGER, pct -> NUMERIC(9,6) CHECK 0..1
  * lists -> ref_<list>(code TEXT PRIMARY KEY, sort_order) seeded from the schema;
    reference tables (Sources, Destinations, Stockpiles) -> real tables with seeds
  * calc fields with an "sql" expression -> GENERATED ALWAYS AS ... STORED (PostgreSQL),
    VIRTUAL (SQLite) or a <table>_calc view when a dialect cannot generate them;
    Excel lookups (INDEX/MATCH into a reference table) -> LEFT JOIN in <table>_calc
  * kpi_daily / kpi_monthly views reproduce the workbook KPI catalogue
"""
from __future__ import annotations

import json
import re
import sqlite3
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
SCHEMA_PATH = HERE / "schema" / "mps_schema.json"
OUT_POSTGRES = HERE / "sql" / "schema_postgres.sql"
OUT_SQLITE = HERE / "sql" / "schema_sqlite.sql"
OUT_DICTIONARY = ROOT / "docs" / "data_dictionary.md"

sys.path.insert(0, str(HERE))
try:
    from build_workbook import KPIS  # the KPI catalogue is the single source of KPI definitions
except ImportError as exc:  # pragma: no cover - environment problem, not a schema problem
    sys.exit(f"build_ddl.py needs build_workbook.py (and openpyxl) next to it: {exc}")

SQL_TYPES = {
    "date": "DATE",
    "text": "TEXT",
    "list": "TEXT",
    "number": "NUMERIC(18,4)",
    "int": "INTEGER",
    "pct": "NUMERIC(9,6)",
}
NUMERIC_TYPES = ("number", "pct")

# ----------------------------------------------------------------------------- MODEL RULES
# Rules the schema does not carry yet. Schema keys win when present (see module docstring).

# Natural key of the tables that hold several rows per day (schema key "unique").
NATURAL_KEYS = {
    "tblMovement": ["Date", "Shift", "Source", "Material", "Destination"],
    "tblFleet": ["Date", "Equipment_Class"],
}

# Reference table keys that must also exist as keys of other reference tables
# (schema desc of Stockpiles.Stockpile: "must match a Source and a Destination").
REF_KEY_MATCHES = {"Stockpiles": ["Sources", "Destinations"]}

# Dialect-specific replacements for calc "sql" expressions that are not portable
# (schema keys "sql_postgres" / "sql_sqlite" win when present).
# tblBudget.Days: the schema expression date_trunc('month', month) resolves to the
# timestamptz overload in PostgreSQL, which is only STABLE, so it is refused inside
# a generated column; the cast to timestamp makes it IMMUTABLE. SQLite has no EXTRACT.
SQL_OVERRIDES = {
    ("tblBudget", "Days"): {
        "postgres": "EXTRACT(DAY FROM (date_trunc('month', month::timestamp) + INTERVAL '1 month - 1 day'))::integer",
        "sqlite": "CAST(strftime('%d', date(month, 'start of month', '+1 month', '-1 day')) AS INTEGER)",
    },
}

# KPI catalogue entries whose Excel Day formula is neither a plain SUMIFS over one table
# nor arithmetic over other KPIs. "agg" runs inside the per-table aggregate (GROUP BY date);
# {num:column} is the column cast to REAL in SQLite. "custom" names a handler in Model.custom_kpi
# that runs in the final SELECT over the base row b (one row per date).
KPI_SQL = {
    # COUNTIFS(tblPlant[Date],{d},tblPlant[Mill_Run_h],"<>")*24: 24 h for each day with a mill run entry
    "Mill_Calendar_h": {"agg": ("tblPlant", "COUNT(mill_run_h) * 24")},
    # point value: the day's GIC when reported, else NULL (Excel shows blank)
    "GIC_oz": {"agg": ("tblPlant", "NULLIF(SUM({num:gic_oz}), 0)"), "nullable": True},
    # rolling 12 calendar months plus month to date, per million hours; prior-year months come from
    # safety_history for months that have no safety_daily rows (the workbook uses cfg_YearStart as the cut)
    "TRIFR_12m": {"custom": "rate12", "column": "recordables"},
    "LTIFR_12m": {"custom": "rate12", "column": "lti"},
    # days since the last LTI in safety_daily (the workbook also reads cfg_PriorLTIDate; the web system
    # should carry the last prior LTI into safety_daily or extend safety_history)
    "Days_Since_LTI": {"custom": "days_since_lti"},
}

SQL_WORDS = {
    "case", "when", "then", "else", "end", "and", "or", "not", "like", "is", "null", "in", "between", "cast", "as",
    "integer", "numeric", "real", "text", "interval", "extract", "day", "month", "year", "from", "true", "false",
    "timestamp", "date", "coalesce", "nullif", "abs", "round", "date_trunc", "strftime", "julianday",
}

LOOKUP_RE = re.compile(r"INDEX\((tbl\w+)\[(\w+)\],MATCH\(\{(\w+)\},\1\[(\w+)\],0\)\)")
SUMIFS_RE = re.compile(r"^SUMIFS\((tbl\w+)\[(\w+)\],\1\[Date\],\{d\}((?:,\1\[\w+\],(?:\"[^\"]*\"|-?[\d.]+))*)\)$")
CRIT_RE = re.compile(r",tbl\w+\[(\w+)\],(\"[^\"]*\"|-?[\d.]+)")
REF_RE = re.compile(r"\{([A-Za-z_][A-Za-z0-9_]*)\}")
IDENT_RE = re.compile(r"\b[A-Za-z_][A-Za-z0-9_]*\b")


# ----------------------------------------------------------------------------- naming
def snake(excel_table: str) -> str:
    n = re.sub(r"^tbl", "", excel_table)
    n = re.sub(r"(?<=[a-z0-9])(?=[A-Z])", "_", n)
    return n.lower()


def cname(field: dict) -> str:
    return field.get("sql_name") or field["name"].lower()


def sql_str(value) -> str:
    if value is None:
        return "NULL"
    if isinstance(value, (int, float)):
        return repr(value)
    return "'" + str(value).replace("'", "''") + "'"


def rewrite_idents(expr: str, mapping) -> str:
    """Replace bare identifiers outside string literals; mapping is a callable name -> replacement or None."""
    out = []
    for part in re.split(r"('(?:[^']|'')*')", expr):
        if part.startswith("'"):
            out.append(part)
        else:
            out.append(IDENT_RE.sub(lambda m: mapping(m.group(0)) or m.group(0), part))
    return "".join(out)


# ----------------------------------------------------------------------------- model
class Model:
    def __init__(self, schema: dict, sqlite_generated: bool = True):
        self.s = schema
        self.sqlite_generated = sqlite_generated
        self.troy = schema["constants"]["TROY_OZ_G"]
        self.rate_basis = schema["constants"].get("RATE_BASIS_HOURS", 1000000)
        self.lists = dict(schema["lists"])
        self.lists.update(schema.get("derived_lists", {}))
        self.ref_tables = schema["reference_tables"]
        self.by_excel = {t["table"]: t for t in schema["tables"]}
        for rname, rdef in self.ref_tables.items():
            self.by_excel[rdef["table"]] = self.ref_tdef(rname)
        # list name that is really a reference table key -> (sql table, sql column)
        self.ref_key_lists = {rdef["key"]: (self.tname(self.ref_tdef(r)), rdef["key"].lower()) for r, rdef in self.ref_tables.items()}
        self.notes: list[str] = []  # generation notes for the file headers

    # ---- schema access ----------------------------------------------------------
    def ref_tdef(self, rname: str) -> dict:
        rdef = self.ref_tables[rname]
        return {"table": rdef["table"], "sheet": rdef["sheet"], "title": rname, "fields": rdef["columns"], "key_field": rdef["key"],
                "rows": rdef["rows"], "ref_name": rname, "is_ref": True}

    def tname(self, tdef: dict) -> str:
        if tdef.get("sql_table"):
            return tdef["sql_table"]
        n = snake(tdef["table"])
        if tdef.get("prefill") == "dates" and not n.endswith("_daily"):
            n += "_daily"
        return n

    @staticmethod
    def input_fields(tdef: dict) -> list[dict]:
        return [f for f in tdef["fields"] if not f.get("calc")]

    @staticmethod
    def calc_fields(tdef: dict) -> list[dict]:
        return [f for f in tdef["fields"] if f.get("calc")]

    def key_field(self, tdef: dict) -> dict | None:
        """Field that is the primary key: the key field of prefilled tables, the key of reference tables."""
        if tdef.get("is_ref"):
            return next(f for f in tdef["fields"] if f["name"] == tdef["key_field"])
        if tdef.get("prefill"):
            return next((f for f in tdef["fields"] if f.get("key")), None)
        return None

    def natural_key(self, tdef: dict) -> list[str]:
        return list(tdef.get("unique") or NATURAL_KEYS.get(tdef["table"], []))

    def list_target(self, f: dict) -> tuple[str, str]:
        lst = f["list"]
        if lst in self.ref_key_lists:
            return self.ref_key_lists[lst]
        if lst in self.lists:
            return f"ref_{lst.lower()}", "code"
        raise KeyError(f"field {f['name']} refers to unknown list {lst}")

    def lookup_of(self, f: dict) -> dict | None:
        m = LOOKUP_RE.search(f.get("calc", ""))
        if not m:
            return None
        ref_tdef = self.by_excel[m.group(1)]
        ref_col = next(c for c in ref_tdef["fields"] if c["name"] == m.group(2))
        return {"table": self.tname(ref_tdef), "column": cname(ref_col), "local": m.group(3).lower(), "key": m.group(4).lower()}

    # ---- calculated fields -------------------------------------------------------
    def calc_expr(self, tdef: dict, f: dict, dialect: str) -> str | None:
        expr = f.get(f"sql_{dialect}") or SQL_OVERRIDES.get((tdef["table"], f["name"]), {}).get(dialect) or f.get("sql")
        if not expr:
            return None
        return expr.replace("{TROY}", repr(self.troy))

    def same_row_inputs_only(self, tdef: dict, expr: str) -> bool:
        body = re.sub(r"'(?:[^']|'')*'", "''", expr)
        body = re.sub(r"\b[A-Za-z_][A-Za-z0-9_]*\s*\(", "(", body)  # drop function names
        inputs = {cname(g) for g in self.input_fields(tdef)}
        for ident in set(IDENT_RE.findall(body)):
            if ident.lower() in SQL_WORDS or ident in inputs:
                continue
            return False
        return True

    def dialect_expr(self, tdef: dict, expr: str, dialect: str, prefix: str = "") -> str:
        """Qualify column references and, for SQLite, cast numeric columns to REAL so that a division
        never falls into integer arithmetic (SQLite stores integral NUMERIC values as INTEGER)."""
        cols = {cname(g): g for g in tdef["fields"]}

        def rep(name):
            g = cols.get(name)
            if g is None:
                return None
            ref = f"{prefix}{name}"
            if dialect == "sqlite" and g["type"] in NUMERIC_TYPES:
                return f"CAST({ref} AS REAL)"
            return ref if prefix else None

        return rewrite_idents(expr, rep)

    def sqlite_can_generate(self, tdef: dict, f: dict, expr: str) -> bool:
        con = sqlite3.connect(":memory:")
        cols = ", ".join(f"{cname(g)} {SQL_TYPES[g['type']]}" for g in self.input_fields(tdef))
        try:
            con.execute(f"CREATE TABLE t ({cols}, {cname(f)} {SQL_TYPES[f['type']]} GENERATED ALWAYS AS ({expr}) VIRTUAL)")
            con.execute("INSERT INTO t DEFAULT VALUES")
            con.execute("SELECT * FROM t").fetchall()
            return True
        except sqlite3.Error:
            return False
        finally:
            con.close()

    def plan_calcs(self, tdef: dict, dialect: str) -> dict:
        """Decide for each calc field: generated column, view column, lookup join or Excel-only."""
        plan = {"generated": [], "view": [], "lookup": [], "excel_only": []}
        for f in self.calc_fields(tdef):
            lk = self.lookup_of(f)
            if lk:
                plan["lookup"].append((f, lk))
                continue
            expr = self.calc_expr(tdef, f, dialect)
            if expr is None:
                plan["excel_only"].append(f)
                continue
            if self.same_row_inputs_only(tdef, expr):
                if dialect == "postgres":
                    plan["generated"].append((f, expr))
                    continue
                sq = self.dialect_expr(tdef, expr, "sqlite")
                if self.sqlite_generated and self.sqlite_can_generate(tdef, f, sq):
                    plan["generated"].append((f, sq))
                    continue
            plan["view"].append((f, expr))
        return plan

    # ---- DDL ---------------------------------------------------------------------
    def id_column(self, dialect: str) -> str:
        return "id BIGSERIAL PRIMARY KEY" if dialect == "postgres" else "id INTEGER PRIMARY KEY AUTOINCREMENT"

    def audit_columns(self) -> list[str]:
        return ["created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP",
                "updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP",
                "updated_by TEXT"]

    def first_of_month_check(self, col: str, dialect: str) -> str:
        if dialect == "postgres":
            return f"EXTRACT(DAY FROM {col}) = 1"
        return f"strftime('%d', {col}) = '01'"

    def check_for(self, f: dict) -> str | None:
        col = cname(f)
        t = f["type"]
        if t == "pct":
            lo = max(0, f.get("min", 0))
            hi = min(1, f.get("max", 1))
            return f"{col} BETWEEN {lo} AND {hi}"
        if t in ("number", "int") and ("min" in f or "max" in f):
            if "min" in f and "max" in f:
                return f"{col} BETWEEN {f['min']} AND {f['max']}"
            if "min" in f:
                return f"{col} >= {f['min']}"
            return f"{col} <= {f['max']}"
        return None

    def validation_text(self, tdef: dict, f: dict) -> str:
        if f.get("calc"):
            return ""
        if f["type"] == "list":
            tbl, col = self.list_target(f)
            return f"list {f['list']} ({tbl}.{col})"
        if f["type"] == "pct":
            return f"{max(0, f.get('min', 0))} to {min(1, f.get('max', 1))}"
        if "min" in f or "max" in f:
            return f"{f.get('min', '')} to {f.get('max', '')}".strip()
        if f["type"] == "date" and tdef.get("prefill") in ("months", "prior_months") and f.get("key"):
            return "first day of month"
        return ""

    def table_ddl(self, tdef: dict, dialect: str) -> list[str]:
        name = self.tname(tdef)
        key = self.key_field(tdef)
        cols: list[str] = []
        cons: list[str] = []
        if key is None:
            cols.append(self.id_column(dialect))
        for f in self.input_fields(tdef):
            col = cname(f)
            parts = [col, SQL_TYPES[f["type"]]]
            if f.get("required") or f.get("key") or f is key:
                parts.append("NOT NULL")
            cols.append(" ".join(parts))
            ck = self.check_for(f)
            if ck:
                cons.append(f"CONSTRAINT ck_{name}_{col} CHECK ({ck})")
            if f["type"] == "date" and f is key and tdef.get("prefill") in ("months", "prior_months"):
                cons.append(f"CONSTRAINT ck_{name}_{col}_first_of_month CHECK ({self.first_of_month_check(col, dialect)})")
            if f["type"] == "list":
                rt, rc = self.list_target(f)
                cons.append(f"CONSTRAINT fk_{name}_{col} FOREIGN KEY ({col}) REFERENCES {rt} ({rc})")
        plan = self.plan_calcs(tdef, dialect)
        for f, expr in plan["generated"]:
            storage = "STORED" if dialect == "postgres" else "VIRTUAL"
            cols.append(f"{cname(f)} {SQL_TYPES[f['type']]} GENERATED ALWAYS AS ({expr}) {storage}")
        cols += self.audit_columns()
        if key is not None:
            cons.insert(0, f"CONSTRAINT pk_{name} PRIMARY KEY ({cname(key)})")
        if tdef.get("is_ref"):
            for other in REF_KEY_MATCHES.get(tdef["ref_name"], []):
                ot = self.ref_tdef(other)
                cons.append(f"CONSTRAINT fk_{name}_{self.tname(ot)} FOREIGN KEY ({cname(key)}) REFERENCES {self.tname(ot)} ({self.ref_tables[other]['key'].lower()})")
        nk = self.natural_key(tdef)
        stmts: list[str] = []
        by_name = {f["name"]: f for f in tdef["fields"]}
        if nk:
            nk_fields = [by_name[n] for n in nk]
            nullable = [f for f in nk_fields if not (f.get("required") or f.get("key"))]
            if not nullable:
                cons.append(f"CONSTRAINT uq_{name}_key UNIQUE ({', '.join(cname(f) for f in nk_fields)})")
            else:
                parts = []
                for f in nk_fields:
                    if f in nullable:
                        if f["type"] not in ("text", "list"):
                            raise NotImplementedError(f"nullable non-text column {f['name']} in the natural key of {name}")
                        parts.append(f"COALESCE({cname(f)}, '')")
                    else:
                        parts.append(cname(f))
                stmts.append(f"CREATE UNIQUE INDEX IF NOT EXISTS ux_{name}_key ON {name} ({', '.join(parts)});")
        body = ",\n".join("    " + c for c in cols + cons)
        stmts.insert(0, f"CREATE TABLE IF NOT EXISTS {name} (\n{body}\n);")
        # index on date for tables where the date is neither the key nor the head of the natural key
        date_f = next((f for f in self.input_fields(tdef) if f["type"] == "date" and f["name"] == "Date"), None)
        if date_f is not None and key is None and (not nk or nk[0] != "Date"):
            stmts.append(f"CREATE INDEX IF NOT EXISTS ix_{name}_date ON {name} (date);")
        return stmts

    def comments_ddl(self, tdef: dict) -> list[str]:
        """PostgreSQL COMMENT ON statements from labels and descriptions."""
        name = self.tname(tdef)
        out = [f"COMMENT ON TABLE {name} IS {sql_str(tdef.get('title', name) + (' (Excel ' + tdef['table'] + ')'))};"]
        for f in tdef["fields"]:
            if self.lookup_of(f) or (f.get("calc") and not self.calc_expr(tdef, f, "postgres")):
                continue
            text = f.get("label", f["name"])
            if f.get("unit"):
                text += f" [{f['unit']}]"
            if f.get("desc"):
                text += ": " + f["desc"]
            out.append(f"COMMENT ON COLUMN {name}.{cname(f)} IS {sql_str(text)};")
        return out

    def calc_view(self, tdef: dict, dialect: str) -> list[str]:
        plan = self.plan_calcs(tdef, dialect)
        if not plan["lookup"] and not plan["view"]:
            return []
        name = self.tname(tdef)
        select = ["t.*"]
        joins = []
        for i, (f, lk) in enumerate(plan["lookup"], start=1):
            select.append(f"j{i}.{lk['column']} AS {cname(f)}")
            joins.append(f"LEFT JOIN {lk['table']} j{i} ON j{i}.{lk['key']} = t.{lk['local']}")
        for f, expr in plan["view"]:
            select.append(f"{self.dialect_expr(tdef, expr, dialect, prefix='t.')} AS {cname(f)}")
        sql = f"CREATE VIEW {name}_calc AS\nSELECT " + ",\n       ".join(select) + f"\nFROM {name} t"
        if joins:
            sql += "\n" + "\n".join(joins)
        return [f"DROP VIEW IF EXISTS {name}_calc;", sql + ";"]

    def has_calc_view(self, tdef: dict, dialect: str) -> bool:
        plan = self.plan_calcs(tdef, dialect)
        return bool(plan["lookup"] or plan["view"])

    def touch_trigger(self, tdef: dict, dialect: str) -> list[str]:
        name = self.tname(tdef)
        if dialect == "postgres":
            return [f"DROP TRIGGER IF EXISTS trg_{name}_touch ON {name};",
                    f"CREATE TRIGGER trg_{name}_touch BEFORE UPDATE ON {name} FOR EACH ROW EXECUTE FUNCTION mps_touch_updated_at();"]
        return [f"CREATE TRIGGER IF NOT EXISTS trg_{name}_touch AFTER UPDATE ON {name} FOR EACH ROW\n"
                f"WHEN NEW.updated_at IS OLD.updated_at\n"
                f"BEGIN\n    UPDATE {name} SET updated_at = CURRENT_TIMESTAMP WHERE rowid = NEW.rowid;\nEND;"]

    # ---- seeds -------------------------------------------------------------------
    def seed_lists(self, dialect: str) -> list[str]:
        out = []
        for lname, values in self.lists.items():
            rows = ", ".join(f"({sql_str(v)}, {i})" for i, v in enumerate(values, start=1))
            out.append(self.insert_ignore(f"ref_{lname.lower()}", "code, sort_order", rows, "code", dialect))
        return out

    def seed_refs(self, dialect: str) -> list[str]:
        out = []
        for rname in self.ref_tables:
            tdef = self.ref_tdef(rname)
            cols = ", ".join(cname(f) for f in tdef["fields"])
            rows = ", ".join("(" + ", ".join(sql_str(v) for v in row) + ")" for row in tdef["rows"])
            out.append(self.insert_ignore(self.tname(tdef), cols, rows, cname(self.key_field(tdef)), dialect))
        return out

    @staticmethod
    def insert_ignore(table: str, cols: str, rows: str, key: str, dialect: str) -> str:
        if dialect == "postgres":
            return f"INSERT INTO {table} ({cols}) VALUES {rows} ON CONFLICT ({key}) DO NOTHING;"
        return f"INSERT OR IGNORE INTO {table} ({cols}) VALUES {rows};"

    # ---- KPI views ---------------------------------------------------------------
    def kpi_source(self, excel_table: str, dialect: str) -> str:
        tdef = self.by_excel[excel_table]
        name = self.tname(tdef)
        return f"{name}_calc" if self.has_calc_view(tdef, dialect) else name

    def num(self, excel_table: str, col: str, dialect: str, prefix: str = "") -> str:
        """Column reference for aggregation, cast to REAL in SQLite when the source field is numeric."""
        tdef = self.by_excel[excel_table]
        f = next((g for g in tdef["fields"] if cname(g) == col), None)
        if dialect == "sqlite" and f is not None and f["type"] in NUMERIC_TYPES:
            return f"CAST({prefix}{col} AS REAL)"
        return f"{prefix}{col}"

    def kpi_plan(self, dialect: str) -> tuple[list[dict], list[str]]:
        """Translate the workbook KPI catalogue. Returns (entries, skipped keys)."""
        entries, skipped = [], []
        for k in KPIS:
            key = k["key"]
            e = {"key": key, "col": key.lower(), "label": k["label"], "unit": k["unit"], "kind": k["kind"], "nullable": False}
            ov = KPI_SQL.get(key)
            if ov and "agg" in ov:
                tbl, expr = ov["agg"]
                expr = re.sub(r"\{num:(\w+)\}", lambda m: self.num(tbl, m.group(1), dialect), expr)
                e.update(mode="agg", table=tbl, agg=expr, nullable=ov.get("nullable", False))
            elif ov and "custom" in ov:
                e.update(mode="custom", custom=ov["custom"], column=ov.get("column"), nullable=True)
            elif k["kind"] == "sum":
                m = SUMIFS_RE.match(k["day"])
                if m:
                    tbl, col = m.group(1), m.group(2).lower()
                    crit = []
                    for c, v in CRIT_RE.findall(m.group(3)):
                        lit = sql_str(v[1:-1]) if v.startswith('"') else v
                        crit.append(f"{c.lower()} = {lit}")
                    ref = self.num(tbl, col, dialect)
                    agg = f"SUM({ref})" if not crit else f"SUM(CASE WHEN {' AND '.join(crit)} THEN {ref} ELSE 0 END)"
                    e.update(mode="agg", table=tbl, agg=agg)
                elif re.fullmatch(r"[\s\d.+\-*/()]*(\{\w+\}[\s\d.+\-*/()]*)+", k["day"]):
                    e.update(mode="expr", expr=k["day"])
                else:
                    skipped.append(key)
                    continue
            elif k["kind"] == "ratio":
                e.update(mode="ratio", num=k["num"].replace("cfg_TroyOz", repr(self.troy)), den=k["den"].replace("cfg_TroyOz", repr(self.troy)))
                e["nullable"] = True
            else:
                skipped.append(key)
                continue
            entries.append(e)
        return entries, skipped

    def kpi_daily_view(self, dialect: str) -> tuple[list[str], list[str]]:
        entries, skipped = self.kpi_plan(dialect)
        by_key = {e["key"]: e for e in entries}
        # per-table aggregates
        tables: dict[str, list[dict]] = {}
        for e in entries:
            if e["mode"] == "agg":
                tables.setdefault(e["table"], []).append(e)
        ctes = []
        spine = " UNION ".join(f"SELECT date FROM {self.tname(self.by_excel[t])}" for t in tables)
        ctes.append(f"days AS (\n    {spine}\n)")
        alias = {}
        for i, (tbl, es) in enumerate(tables.items(), start=1):
            a = f"a{i}"
            alias[tbl] = a
            cols = ",\n           ".join(f"{e['agg']} AS {e['col']}" for e in es)
            ctes.append(f"agg_{snake(tbl)} AS (\n    SELECT date,\n           {cols}\n    FROM {self.kpi_source(tbl, dialect)}\n    GROUP BY date\n)")
        base_cols = ["d.date"]
        for tbl, es in tables.items():
            for e in es:
                ref = f"{alias[tbl]}.{e['col']}"
                base_cols.append(f"{ref} AS {e['col']}" if e["nullable"] else f"COALESCE({ref}, 0) AS {e['col']}")
        joins = "\n".join(f"    LEFT JOIN agg_{snake(tbl)} {alias[tbl]} ON {alias[tbl]}.date = d.date" for tbl in tables)
        ctes.append("base AS (\n    SELECT " + ",\n           ".join(base_cols) + "\n    FROM days d\n" + joins + "\n)")

        def resolve(expr: str) -> str:
            def rep(m):
                e = by_key[m.group(1)]
                if e["mode"] == "agg":
                    return f"b.{e['col']}"
                if e["mode"] == "expr":
                    return "(" + resolve(e["expr"]) + ")"
                raise KeyError(f"KPI {m.group(1)} cannot be used inside another KPI definition")
            return REF_RE.sub(rep, expr)

        select = ["b.date"]
        for e in entries:
            if e["mode"] == "agg":
                select.append(f"b.{e['col']}")
            elif e["mode"] == "expr":
                select.append(f"{resolve(e['expr'])} AS {e['col']}")
            elif e["mode"] == "ratio":
                num, den = resolve(e["num"]), resolve(e["den"])
                select.append(f"CASE WHEN ({den}) > 0 THEN ({num}) / ({den}) END AS {e['col']}")
            else:
                select.append(f"{self.custom_kpi(e, dialect)} AS {e['col']}")
        sql = "CREATE VIEW kpi_daily AS\nWITH " + ",\n".join(ctes) + "\nSELECT " + ",\n       ".join(select) + "\nFROM base b;"
        return [sql], skipped

    def custom_kpi(self, e: dict, dialect: str) -> str:
        if dialect == "postgres":
            month_start = "date_trunc('month', b.date::timestamp)::date"
            win_start = "(date_trunc('month', b.date::timestamp) - INTERVAL '12 months')::date"
            month_of_s2 = "date_trunc('month', s2.date::timestamp)::date"
        else:
            month_start = "date(b.date, 'start of month')"
            win_start = "date(b.date, 'start of month', '-12 months')"
            month_of_s2 = "date(s2.date, 'start of month')"
        if e["custom"] == "rate12":
            col = e["column"]
            daily = self.kpi_source("tblSafety", dialect)  # safety_daily, or its _calc view when hours_total lives there
            history = self.kpi_source("tblSafetyHistory", dialect)

            def window(expr_daily: str, expr_hist: str) -> str:
                return (f"((SELECT COALESCE(SUM({expr_daily}), 0) FROM {daily} s WHERE s.date >= {win_start} AND s.date <= b.date)"
                        f" + (SELECT COALESCE(SUM({expr_hist}), 0) FROM {history} h WHERE h.month >= {win_start} AND h.month < {month_start}"
                        f" AND NOT EXISTS (SELECT 1 FROM {daily} s2 WHERE {month_of_s2} = h.month)))")
            events = window(f"s.{col}", f"h.{col}")
            hrs = window(self.num("tblSafety", "hours_total", dialect, prefix="s."),
                         self.num("tblSafetyHistory", "hours_total", dialect, prefix="h."))
            return f"CASE WHEN {hrs} > 0 THEN {events} * {self.rate_basis} / {hrs} END"
        if e["custom"] == "days_since_lti":
            last = f"(SELECT MAX(s.date) FROM {self.kpi_source('tblSafety', dialect)} s WHERE s.lti > 0 AND s.date <= b.date)"
            if dialect == "postgres":
                return f"b.date - {last}"
            return f"CAST(julianday(b.date) - julianday({last}) AS INTEGER)"
        raise KeyError(e["custom"])

    def kpi_monthly_view(self, dialect: str) -> list[str]:
        entries, _ = self.kpi_plan(dialect)
        by_key = {e["key"]: e for e in entries}
        month = "date_trunc('month', k.date)::date" if dialect == "postgres" else "date(k.date, 'start of month')"

        def resolve(expr: str) -> str:
            def rep(m):
                e = by_key[m.group(1)]
                if e["mode"] in ("agg", "expr"):
                    return f"SUM(k.{e['col']})"
                raise KeyError(m.group(1))
            return REF_RE.sub(rep, expr)

        select = [f"{month} AS month", "COUNT(*) AS days_reported"]
        for e in entries:
            if e["mode"] in ("agg", "expr") and not e["nullable"]:
                select.append(f"SUM(k.{e['col']}) AS {e['col']}")
            elif e["mode"] == "ratio":
                num, den = resolve(e["num"]), resolve(e["den"])
                select.append(f"CASE WHEN ({den}) > 0 THEN ({num}) / ({den}) END AS {e['col']}")
        sql = "CREATE VIEW kpi_monthly AS\nSELECT " + ",\n       ".join(select) + f"\nFROM kpi_daily k\nGROUP BY {month};"
        return [sql]

    # ---- whole script --------------------------------------------------------------
    def table_mapping(self) -> list[tuple[str, str, str, str]]:
        """(Excel table, sheet, SQL table, primary key) for input and reference tables."""
        rows = []
        for tdef in self.s["tables"]:
            key = self.key_field(tdef)
            rows.append((tdef["table"], tdef["sheet"], self.tname(tdef), cname(key) if key else "id"))
        for rname in self.ref_tables:
            tdef = self.ref_tdef(rname)
            rows.append((tdef["table"], tdef["sheet"], self.tname(tdef), cname(self.key_field(tdef))))
        return rows

    def excel_only_fields(self) -> list[tuple[str, str]]:
        out = []
        for tdef in self.s["tables"]:
            for f in self.calc_fields(tdef):
                if not self.lookup_of(f) and not self.calc_expr(tdef, f, "postgres"):
                    out.append((tdef["table"], f["name"]))
        return out

    def script(self, dialect: str) -> str:
        s = self.s
        site = s["site"]
        L: list[str] = []
        title = "PostgreSQL" if dialect == "postgres" else "SQLite"
        L.append(f"-- {s['system']} ({site['name']}, {site['company']}) schema for {title}")
        L.append(f"-- Generated by mps/build_ddl.py from mps/schema/mps_schema.json version {s['version']}. Do not edit; regenerate.")
        L.append("-- Safe to re-run on an existing database: tables use IF NOT EXISTS, views are dropped and recreated,")
        L.append("-- seeds are inserted only when missing. Structural changes to existing tables need a migration.")
        if dialect == "postgres":
            L.append("-- Run: psql -v ON_ERROR_STOP=1 -1 -f schema_postgres.sql  (single transaction; needs plpgsql, PostgreSQL 13 or later)")
        else:
            L.append("-- Run: sqlite3 mps.db < schema_sqlite.sql  (SQLite 3.31 or later; every connection must set PRAGMA foreign_keys = ON)")
        L.append("--")
        L.append("-- Naming: Excel table -> SQL table (primary key)")
        for excel, sheet, name, pk in self.table_mapping():
            L.append(f"--   {excel:<18} sheet {sheet:<18} -> {name} ({pk})")
        L.append("--   lists           -> ref_<list>(code, sort_order), one row per drop-down value")
        L.append("-- Columns are the Excel field names in lower case. Units: tonnes dry metric, grades g/t,")
        L.append(f"-- gold troy ounces ({self.troy} g), percentages stored as fractions 0 to 1.")
        L.append("-- Types: date DATE, text TEXT, list TEXT + FOREIGN KEY, number NUMERIC(18,4), int INTEGER, pct NUMERIC(9,6) CHECK 0..1.")
        L.append("-- CHECK constraints come from the schema min/max; NULL means not reported and always passes.")
        if dialect == "postgres":
            L.append("-- Calculated fields are GENERATED ALWAYS AS ... STORED; Excel lookups (source type, destination type)")
            L.append("-- are LEFT JOINs in the <table>_calc views.")
        else:
            L.append("-- Calculated fields are GENERATED ALWAYS AS ... VIRTUAL with numeric operands cast to REAL")
            L.append("-- (SQLite stores integral NUMERIC values as INTEGER and would otherwise divide as integers);")
            L.append("-- fields SQLite cannot generate and Excel lookups live in the <table>_calc views.")
        L.append("-- kpi_daily reproduces the workbook Calc_Daily day columns; kpi_monthly gives the month totals")
        L.append("-- with ratios recomputed from summed numerators and denominators.")
        eo = self.excel_only_fields()
        if eo:
            L.append("-- Excel-only helper fields not carried to SQL: " + ", ".join(f"{t}.{f}" for t, f in eo) + ".")
        L.append("")
        if dialect == "sqlite":
            L.append("PRAGMA foreign_keys = ON;")
            L.append("")
        # 1. list reference tables
        L.append("-- ---------------------------------------------------------------- reference lists (drop-down values)")
        for lname in self.lists:
            L.append(f"CREATE TABLE IF NOT EXISTS ref_{lname.lower()} (\n    code TEXT PRIMARY KEY,\n    sort_order INTEGER NOT NULL\n);")
        L.append("")
        # 2. reference tables
        L.append("-- ---------------------------------------------------------------- reference tables (Lists and Config sheets)")
        ref_tdefs = [self.ref_tdef(r) for r in self.ref_tables]
        ref_tdefs.sort(key=lambda t: 1 if t["ref_name"] in REF_KEY_MATCHES else 0)  # dependants after their targets
        for tdef in ref_tdefs:
            L.extend(self.table_ddl(tdef, dialect))
        L.append("")
        # 3. input tables
        L.append("-- ---------------------------------------------------------------- input tables")
        for tdef in self.s["tables"]:
            L.append(f"-- {tdef['title']} (owner: {tdef.get('owner', '')})")
            L.extend(self.table_ddl(tdef, dialect))
            L.append("")
        # 4. comments (PostgreSQL)
        if dialect == "postgres":
            L.append("-- ---------------------------------------------------------------- comments")
            for tdef in ref_tdefs + list(self.s["tables"]):
                L.extend(self.comments_ddl(tdef))
            L.append("")
        # 5. audit triggers
        L.append("-- ---------------------------------------------------------------- audit: keep updated_at current")
        if dialect == "postgres":
            L.append("CREATE OR REPLACE FUNCTION mps_touch_updated_at() RETURNS trigger AS $$\nBEGIN\n    NEW.updated_at := CURRENT_TIMESTAMP;\n    RETURN NEW;\nEND;\n$$ LANGUAGE plpgsql;")
        for tdef in ref_tdefs + list(self.s["tables"]):
            L.extend(self.touch_trigger(tdef, dialect))
        L.append("")
        # 6. calc views
        L.append("-- ---------------------------------------------------------------- calculated views")
        L.append("DROP VIEW IF EXISTS kpi_monthly;")
        L.append("DROP VIEW IF EXISTS kpi_daily;")
        for tdef in self.s["tables"]:
            L.extend(self.calc_view(tdef, dialect))
        L.append("")
        # 7. KPI views
        L.append("-- ---------------------------------------------------------------- KPI views (workbook KPI catalogue)")
        L.append("-- Ex-pit = source type Pit; ore = material starting with Ore; rehandle = source type Stockpile.")
        L.append("-- Ratios are NULL when the denominator is zero. Sums are 0 on days without rows, as in the workbook.")
        stmts, skipped = self.kpi_daily_view(dialect)
        if skipped:
            L.append("-- KPI catalogue entries without an SQL translation: " + ", ".join(skipped))
        L.extend(stmts)
        L.extend(self.kpi_monthly_view(dialect))
        L.append("")
        # 8. seeds
        L.append("-- ---------------------------------------------------------------- seed data (inserted only when missing)")
        L.extend(self.seed_lists(dialect))
        L.extend(self.seed_refs(dialect))
        L.append("")
        return "\n".join(L)

    # ---- data dictionary -----------------------------------------------------------
    def dictionary(self) -> str:
        s = self.s
        site = s["site"]
        L: list[str] = []
        esc = lambda v: str(v).replace("|", "\\|").replace("\n", " ")  # noqa: E731
        L.append(f"# {s['system']} data dictionary")
        L.append("")
        L.append(f"{site['name']}, {site['company']} (owner {site['owner']}), {site['country']}. "
                 f"Generated from `mps/schema/mps_schema.json` version {s['version']} by `mps/build_ddl.py`; "
                 "the same schema generates the Excel workbook (`mps/build_workbook.py`). Do not edit this file; change the schema and regenerate.")
        L.append("")
        L.append(f"Conventions: metric units, tonnes dry, grades in g/t, gold in troy ounces ({self.troy} g), "
                 "percentages stored as fractions between 0 and 1, dates are production days (day shift start to night shift end), "
                 "site time zone " + site["timezone"] + ". Blank (NULL) means not reported; zero means reported as zero.")
        L.append("")
        L.append("## Contents")
        L.append("")
        L.append("1. [Excel to SQL mapping](#excel-to-sql-mapping)")
        L.append("2. [Input tables](#input-tables)")
        L.append("3. [Reference tables](#reference-tables)")
        L.append("4. [Lists](#lists)")
        L.append("5. [KPI catalogue](#kpi-catalogue)")
        L.append("")
        # mapping
        L.append("## Excel to SQL mapping")
        L.append("")
        L.append("| Excel table | Sheet | SQL table | Primary key | Rows | Owner |")
        L.append("|---|---|---|---|---|---|")
        for tdef in s["tables"]:
            key = self.key_field(tdef)
            rows = {"dates": "one per day", "months": "one per month", "prior_months": "one per prior-year month"}.get(tdef.get("prefill"), "many per day")
            L.append(f"| {tdef['table']} | {tdef['sheet']} | `{self.tname(tdef)}` | `{cname(key) if key else 'id'}` | {rows} | {esc(tdef.get('owner', ''))} |")
        for rname in self.ref_tables:
            tdef = self.ref_tdef(rname)
            L.append(f"| {tdef['table']} | {tdef['sheet']} | `{self.tname(tdef)}` | `{cname(self.key_field(tdef))}` | reference | Configuration |")
        L.append("")
        L.append("| Schema type | Excel | PostgreSQL and SQLite |")
        L.append("|---|---|---|")
        L.append("| date | date cell, dd-mmm-yyyy | `DATE` (SQLite: ISO text YYYY-MM-DD) |")
        L.append("| text | free text | `TEXT` |")
        L.append("| list | drop-down from the Lists sheet | `TEXT` with a `FOREIGN KEY` to `ref_<list>(code)` or to the reference table |")
        L.append("| number | decimal, min to max warning | `NUMERIC(18,4)` with `CHECK (min to max)` |")
        L.append("| int | whole number, min to max warning | `INTEGER` with `CHECK (min to max)` |")
        L.append("| pct | percentage, stored as a fraction | `NUMERIC(9,6)` with `CHECK (0 to 1)` |")
        L.append("")
        L.append("Rules applied by the generator:")
        L.append("")
        L.append("- Table names drop the `tbl` prefix and become snake_case; tables with one row per day get the suffix `_daily`. Column names are the Excel field names in lower case.")
        L.append("- Primary key: the date or month column of the tables prefilled with one row per period; an `id` column (`BIGSERIAL` in PostgreSQL, `INTEGER PRIMARY KEY AUTOINCREMENT` in SQLite) elsewhere. "
                 "Natural keys: " + "; ".join(f"`{self.tname(self.by_excel[t])}` ({', '.join(c.lower() for c in cols)})" for t, cols in NATURAL_KEYS.items()) +
                 ". Where a natural-key column is optional (movement shift, blank for a daily total) the uniqueness is a unique index on `COALESCE(column, '')` so that two daily totals cannot be entered twice.")
        L.append("- `NOT NULL` for required and key fields. `CHECK` constraints from the schema min and max; Excel only warns, SQL rejects. Month columns must be the first day of the month.")
        L.append("- List fields reference `ref_<list>` tables seeded with the drop-down values in order (`sort_order`). `movement.source` and `movement.destination` reference the `sources` and `destinations` tables; `stockpiles.stockpile` must exist in both.")
        L.append("- Calculated fields with an SQL expression are generated columns (`GENERATED ALWAYS AS ... STORED` in PostgreSQL, `VIRTUAL` in SQLite, where numeric operands are cast to `REAL` to avoid integer division). "
                 "Lookups into reference tables (source type, destination type) are `LEFT JOIN`s in the `<table>_calc` views, and SQLite also puts there anything it cannot generate. "
                 "Generated columns carry no CHECK constraint: validation applies to inputs only.")
        eo = self.excel_only_fields()
        if eo:
            L.append("- Excel-only report helpers are not carried to SQL: " + ", ".join(f"`{t}.{f}`" for t, f in eo) + ".")
        L.append("- Every table except the `ref_*` lists carries `created_at`, `updated_at` (kept current by a trigger) and `updated_by TEXT`.")
        L.append("- `kpi_daily` reproduces the workbook Calc_Daily day columns from the KPI catalogue in `build_workbook.py`; `kpi_monthly` groups it by month with ratios recomputed from the summed numerators and denominators. Both views are recreated by every run of the script.")
        L.append("- Both scripts are idempotent (`IF NOT EXISTS`, seeds inserted only when missing, views recreated). Structural changes to a populated database need a migration.")
        L.append("")
        # input tables
        L.append("## Input tables")
        L.append("")
        for tdef in s["tables"]:
            L.extend(self.dict_table(tdef))
        L.append("## Reference tables")
        L.append("")
        for rname in self.ref_tables:
            tdef = self.ref_tdef(rname)
            L.extend(self.dict_table(tdef))
            cols = [cname(f) for f in tdef["fields"]]
            L.append("Seed rows:")
            L.append("")
            L.append("| " + " | ".join(f"`{c}`" for c in cols) + " |")
            L.append("|" + "---|" * len(cols))
            for row in tdef["rows"]:
                L.append("| " + " | ".join("" if v is None else esc(v) for v in row) + " |")
            L.append("")
        # lists
        L.append("## Lists")
        L.append("")
        L.append("Each list is a table `ref_<list>` with `code TEXT PRIMARY KEY` and `sort_order INTEGER`. Add a value by inserting a row; do not rename codes that are already referenced.")
        L.append("")
        L.append("| List | SQL table | Values (in order) | Used by |")
        L.append("|---|---|---|---|")
        for lname, values in self.lists.items():
            users = []
            for tdef in list(s["tables"]) + [self.ref_tdef(r) for r in self.ref_tables]:
                for f in tdef["fields"]:
                    if f.get("list") == lname:
                        users.append(f"`{self.tname(tdef)}.{cname(f)}`")
            L.append(f"| {lname} | `ref_{lname.lower()}` | {', '.join(values)} | {', '.join(users)} |")
        L.append("")
        # KPI catalogue
        L.append("## KPI catalogue")
        L.append("")
        L.append("Columns of the `kpi_daily` view (one row per date with data), in the order of the workbook KPI catalogue. "
                 "Sum KPIs are day totals (0 when no rows); ratio KPIs are numerator over denominator for the day and NULL when the denominator is 0; "
                 "`kpi_monthly` recomputes ratios from monthly sums, which is how the workbook builds month-to-date figures. "
                 "Ex-pit means source type Pit, ore means a material starting with Ore, rehandle means source type Stockpile.")
        L.append("")
        L.append("| Column | Label | Unit | Kind | Definition (PostgreSQL) |")
        L.append("|---|---|---|---|---|")
        entries, skipped = self.kpi_plan("postgres")
        by_key = {e["key"]: e for e in entries}

        def show(expr: str) -> str:
            return REF_RE.sub(lambda m: by_key[m.group(1)]["col"], expr)
        for e in entries:
            if e["mode"] == "agg":
                d = f"{e['agg']} over `{self.tname(self.by_excel[e['table']])}` per date"
            elif e["mode"] == "expr":
                d = show(e["expr"])
            elif e["mode"] == "ratio":
                d = f"{show(e['num'])} / {show(e['den'])}"
            else:
                d = {"rate12": f"({e.get('column')} over the last 12 calendar months plus month to date) * {self.rate_basis} / hours_total over the same window; prior months from safety_history where safety_daily has no rows",
                     "days_since_lti": "date minus the last date in safety_daily with lti > 0"}[e["custom"]]
            L.append(f"| `{e['col']}` | {esc(e['label'])} | {e['unit']} | {e['kind']} | {esc(d)} |")
        if skipped:
            L.append("")
            L.append("Not translated: " + ", ".join(skipped) + ".")
        L.append("")
        return "\n".join(L)

    def dict_table(self, tdef: dict) -> list[str]:
        L: list[str] = []
        esc = lambda v: str(v).replace("|", "\\|").replace("\n", " ")  # noqa: E731
        name = self.tname(tdef)
        key = self.key_field(tdef)
        L.append(f"### {name} ({tdef['table']}, sheet {tdef['sheet']})")
        L.append("")
        head = tdef.get("title", tdef.get("ref_name", name))
        if tdef.get("owner"):
            head += f". Owner: {tdef['owner']}"
        head += f". Primary key: `{cname(key) if key else 'id'}`"
        nk = self.natural_key(tdef)
        if nk:
            head += f"; unique: ({', '.join(c.lower() for c in nk)})"
        for other in REF_KEY_MATCHES.get(tdef.get("ref_name", ""), []):
            head += f"; `{cname(key)}` must exist in `{self.tname(self.ref_tdef(other))}`"
        plan = self.plan_calcs(tdef, "postgres")
        plan_lite = self.plan_calcs(tdef, "sqlite")
        if plan["lookup"] or plan["view"] or plan_lite["view"]:
            head += f". View `{name}_calc` adds the lookup columns" + (" (SQLite: and the fields it cannot generate)" if plan_lite["view"] else "")
        L.append(head + ".")
        L.append("")
        L.append("| Column | Excel field | Label | Type | Unit | Required | Validation | Description | Calculation |")
        L.append("|---|---|---|---|---|---|---|---|---|")
        gen = {f["name"]: expr for f, expr in plan["generated"]}
        view = {f["name"]: expr for f, expr in plan["view"]}
        look = {f["name"]: lk for f, lk in plan["lookup"]}
        for f in tdef["fields"]:
            calc = ""
            typ = SQL_TYPES[f["type"]]
            if f["name"] in gen:
                calc = f"`{esc(gen[f['name']])}`"
                typ = f"generated {typ}"
            elif f["name"] in view:
                calc = f"`{esc(view[f['name']])}` (view)"
                typ = f"view {typ}"
            elif f["name"] in look:
                lk = look[f["name"]]
                calc = f"lookup `{lk['table']}.{lk['column']}` (view)"
                typ = f"view {typ}"
            elif f.get("calc"):
                calc = f"Excel only: `{esc(f['calc'])}`"
                typ = "not in SQL"
            req = "key" if (f is key) else ("yes" if f.get("required") else "")
            L.append(f"| `{cname(f)}` | {f['name']} | {esc(f.get('label', f['name']))} | {f['type']} ({typ}) | {f.get('unit', '')} | {req} | "
                     f"{esc(self.validation_text(tdef, f))} | {esc(f.get('desc', ''))} | {calc} |")
        L.append("| `created_at` | | Created at | TIMESTAMP | | yes | default now | Audit: row creation time | |")
        L.append("| `updated_at` | | Updated at | TIMESTAMP | | yes | default now | Audit: last change, kept current by trigger | |")
        L.append("| `updated_by` | | Updated by | TEXT | | | | Audit: user who made the last change | |")
        L.append("")
        return L


# ----------------------------------------------------------------------------- entry point
def build(schema_path: Path = SCHEMA_PATH) -> dict[str, str]:
    schema = json.loads(Path(schema_path).read_text(encoding="utf-8"))
    m = Model(schema)
    return {
        str(OUT_POSTGRES): m.script("postgres"),
        str(OUT_SQLITE): m.script("sqlite"),
        str(OUT_DICTIONARY): m.dictionary(),
    }


def main(argv=None) -> int:
    outputs = build()
    for path, text in outputs.items():
        p = Path(path)
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(text + ("\n" if not text.endswith("\n") else ""), encoding="utf-8", newline="\n")
        print(f"wrote {p.relative_to(ROOT)} ({len(text.splitlines()):,} lines)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
