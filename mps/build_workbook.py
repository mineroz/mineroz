#!/usr/bin/env python3
"""
build_workbook.py - generate the Mako Production System (MPS) Excel workbook
from mps_schema.json and the KPI catalogue below.

Design rules baked in:
  * one workbook, no external links, no VBA, no hidden sheets
  * every input is an Excel table (ListObject) with data validation
  * all reporting reads from Calc_Daily, one row per day, built with SUMIFS
    over the tables; the report and dashboard never touch the inputs directly
  * the same schema drives the SQL DDL for the web phase (build_ddl.py)

Usage:
    python build_workbook.py --out MPS_2026.xlsx [--year 2026] [--sample-days 74]

--sample-days N fills N days of synthetic data (clearly marked DEMO) so the
report can be reviewed before real data exists.
"""
from __future__ import annotations

import argparse
import calendar
import json
import random
import sys
from datetime import date, timedelta
from pathlib import Path

import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.axis import DateAxis
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.table import Table, TableFormula, TableStyleInfo

HERE = Path(__file__).resolve().parent
DEFAULT_SCHEMA = HERE / "schema" / "mps_schema.json"

# ----------------------------------------------------------------------------- styling
FONT_NAME = "Calibri"
C_HEAD = "1F3864"        # header fill (dark blue), white text
C_HEAD_CALC = "595959"   # calculated column header (dark grey)
C_LABEL = "F2F2F2"       # label row
C_SECTION = "D9E1F2"     # report section band
C_INPUT_CELL = "FFFFFF"
C_KEY = "FFF2CC"         # report date input
TAB = {"report": "2F5597", "input": "548235", "config": "7F7F7F", "calc": "3A3A3A"}

FMT = {
    "t": "#,##0", "oz": "#,##0", "g/t": "0.00", "%": "0.0%", "h": "#,##0.0", "kg": "#,##0", "kWh": "#,##0",
    "m3": "#,##0", "L": "#,##0", "m": "#,##0", "bcm": "#,##0", "t/h": "#,##0", "kg/t": "0.00", "kWh/t": "0.0",
    "m3/t": "0.00", "ratio": "0.00", "rate": "0.00", "days": "0", "int": "0", "date": "dd-mmm-yyyy", "text": "@", "number": "#,##0.00",
}
TROY = "cfg_TroyOz"

thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def font(bold=False, color="000000", size=10, italic=False):
    return Font(name=FONT_NAME, bold=bold, color=color, size=size, italic=italic)


def fill(hex_):
    return PatternFill("solid", start_color=hex_, end_color=hex_)


# ----------------------------------------------------------------------------- KPI catalogue
# kind: sum   -> Day column (formula given), MTD and YTD running totals
#       ratio -> num / den per period, built from other KPI columns
#       custom-> explicit Day formula only (rates, ages)
#       point -> value at the date (not additive)
# budget: name of a tblBudget column; ratio budgets are derived automatically when num/den are budgeted
def K(key, label, unit, kind, **kw):
    d = dict(key=key, label=label, unit=unit, kind=kind)
    d.update(kw)
    return d


def sumifs(tbl, col, *extra):
    parts = [f"{tbl}[{col}]", f"{tbl}[Date]", "{d}"]
    for c, v in extra:
        parts += [f"{tbl}[{c}]", v]
    return "SUMIFS(" + ",".join(parts) + ")"


KPIS = [
    # ---- mining
    K("Ore_Mined_t", "Ore mined (ex-pit)", "t", "sum", day=sumifs("tblMovement", "Tonnes_t", ("Is_Ore", "1"), ("Source_Type", '"Pit"')), budget="Ore_Mined_t"),
    K("Ore_Mined_oz", "Contained gold mined", "oz", "sum", day=sumifs("tblMovement", "Ore_oz", ("Source_Type", '"Pit"')), budget="Ore_oz"),
    K("Ore_Grade_gpt", "Ore grade mined", "g/t", "ratio", num="{Ore_Mined_oz}*" + TROY, den="{Ore_Mined_t}"),
    K("Waste_t", "Waste mined", "t", "sum", day=sumifs("tblMovement", "Tonnes_t", ("Is_Ore", "0"), ("Source_Type", '"Pit"')), budget="Waste_t"),
    K("TMM_t", "Total material moved", "t", "sum", day="{Ore_Mined_t}+{Waste_t}", budget="TMM_t"),
    K("Strip_Ratio", "Strip ratio (waste : ore)", "ratio", "ratio", num="{Waste_t}", den="{Ore_Mined_t}"),
    K("Rehandle_t", "Rehandle from stockpiles", "t", "sum", day=sumifs("tblMovement", "Tonnes_t", ("Source_Type", '"Stockpile"'))),
    K("Drill_m", "Drilled", "m", "sum", day=sumifs("tblMiningDaily", "Drill_m"), budget="Drill_m"),
    K("Blast_t", "Blasted", "t", "sum", day=sumifs("tblMiningDaily", "Blast_t")),
    K("Explosives_kg", "Explosives", "kg", "sum", day=sumifs("tblMiningDaily", "Explosives_kg")),
    K("Powder_Factor_kgpt", "Powder factor", "kg/t", "ratio", num="{Explosives_kg}", den="{Blast_t}"),
    K("Diesel_Mining_L", "Mining diesel", "L", "sum", day=sumifs("tblMiningDaily", "Diesel_L"), budget="Diesel_L"),
    K("Dewatering_m3", "Pit dewatering", "m3", "sum", day=sumifs("tblMiningDaily", "Dewatering_m3")),
    K("Exc_Calendar_h", "Excavator calendar hours", "h", "sum", day=sumifs("tblFleet", "Calendar_h", ("Equipment_Class", '"Excavator"'))),
    K("Exc_Available_h", "Excavator available hours", "h", "sum", day=sumifs("tblFleet", "Available_h", ("Equipment_Class", '"Excavator"'))),
    K("Exc_Operating_h", "Excavator operating hours", "h", "sum", day=sumifs("tblFleet", "Operating_h", ("Equipment_Class", '"Excavator"'))),
    K("Exc_Availability_pct", "Excavator availability", "%", "ratio", num="{Exc_Available_h}", den="{Exc_Calendar_h}"),
    K("Exc_Utilisation_pct", "Excavator utilisation", "%", "ratio", num="{Exc_Operating_h}", den="{Exc_Available_h}"),
    K("Trk_Calendar_h", "Truck calendar hours", "h", "sum", day=sumifs("tblFleet", "Calendar_h", ("Equipment_Class", '"Haul Truck"'))),
    K("Trk_Available_h", "Truck available hours", "h", "sum", day=sumifs("tblFleet", "Available_h", ("Equipment_Class", '"Haul Truck"'))),
    K("Trk_Operating_h", "Truck operating hours", "h", "sum", day=sumifs("tblFleet", "Operating_h", ("Equipment_Class", '"Haul Truck"'))),
    K("Trk_Availability_pct", "Truck availability", "%", "ratio", num="{Trk_Available_h}", den="{Trk_Calendar_h}"),
    K("Trk_Utilisation_pct", "Truck utilisation", "%", "ratio", num="{Trk_Operating_h}", den="{Trk_Available_h}"),
    # ---- processing
    K("Crushed_t", "Crushed", "t", "sum", day=sumifs("tblPlant", "Crushed_t")),
    K("Milled_t", "Milled", "t", "sum", day=sumifs("tblPlant", "Milled_t"), budget="Milled_t"),
    K("Feed_oz", "Contained gold in feed", "oz", "sum", day=sumifs("tblPlant", "Feed_oz"), budget="Feed_oz"),
    K("Tails_oz", "Gold lost to tails", "oz", "sum", day=sumifs("tblPlant", "Tails_oz")),
    K("Recovered_oz", "Gold recovered", "oz", "sum", day=sumifs("tblPlant", "Recovered_oz"), budget="Recovered_oz"),
    K("Head_Grade_gpt", "Head grade", "g/t", "ratio", num="{Feed_oz}*" + TROY, den="{Milled_t}"),
    K("Tails_Grade_gpt", "Tails grade", "g/t", "ratio", num="{Tails_oz}*" + TROY, den="{Milled_t}"),
    K("Recovery_pct", "Recovery", "%", "ratio", num="{Recovered_oz}", den="{Feed_oz}"),
    K("Gravity_Gold_oz", "Gravity gold", "oz", "sum", day=sumifs("tblPlant", "Gravity_Gold_oz")),
    K("Gold_Poured_oz", "Gold poured", "oz", "sum", day=sumifs("tblPlant", "Gold_Poured_oz"), budget="Gold_Poured_oz"),
    K("Mill_Run_h", "Mill run hours", "h", "sum", day=sumifs("tblPlant", "Mill_Run_h")),
    K("Mill_Planned_Maint_h", "Mill planned maintenance", "h", "sum", day=sumifs("tblPlant", "Mill_Planned_Maint_h")),
    K("Mill_Unplanned_Down_h", "Mill unplanned downtime", "h", "sum", day=sumifs("tblPlant", "Mill_Unplanned_Down_h")),
    K("Mill_Calendar_h", "Mill calendar hours (days reported)", "h", "sum", day='COUNTIFS(tblPlant[Date],{d},tblPlant[Mill_Run_h],"<>")*24'),
    K("Crusher_Run_h", "Crusher run hours", "h", "sum", day=sumifs("tblPlant", "Crusher_Run_h")),
    K("Throughput_tph", "Mill throughput", "t/h", "ratio", num="{Milled_t}", den="{Mill_Run_h}"),
    K("Mill_Availability_pct", "Mill availability", "%", "ratio", num="{Mill_Calendar_h}-{Mill_Planned_Maint_h}-{Mill_Unplanned_Down_h}", den="{Mill_Calendar_h}"),
    K("Mill_Utilisation_pct", "Mill utilisation", "%", "ratio", num="{Mill_Run_h}", den="{Mill_Calendar_h}-{Mill_Planned_Maint_h}-{Mill_Unplanned_Down_h}"),
    K("Cyanide_kg", "Cyanide", "kg", "sum", day=sumifs("tblPlant", "Cyanide_kg")),
    K("Lime_kg", "Lime", "kg", "sum", day=sumifs("tblPlant", "Lime_kg")),
    K("Grinding_Media_kg", "Grinding media", "kg", "sum", day=sumifs("tblPlant", "Grinding_Media_kg")),
    K("Power_kWh", "Plant power", "kWh", "sum", day=sumifs("tblPlant", "Power_kWh")),
    K("Raw_Water_m3", "Raw water", "m3", "sum", day=sumifs("tblPlant", "Raw_Water_m3")),
    K("Cyanide_kgpt", "Cyanide consumption", "kg/t", "ratio", num="{Cyanide_kg}", den="{Milled_t}"),
    K("Lime_kgpt", "Lime consumption", "kg/t", "ratio", num="{Lime_kg}", den="{Milled_t}"),
    K("Media_kgpt", "Grinding media consumption", "kg/t", "ratio", num="{Grinding_Media_kg}", den="{Milled_t}"),
    K("Power_kWhpt", "Power consumption", "kWh/t", "ratio", num="{Power_kWh}", den="{Milled_t}"),
    K("Water_m3pt", "Raw water consumption", "m3/t", "ratio", num="{Raw_Water_m3}", den="{Milled_t}"),
    K("GIC_oz", "Gold in circuit (end of day)", "oz", "point", day="IF(SUMIFS(tblPlant[GIC_oz],tblPlant[Date],{d})>0,SUMIFS(tblPlant[GIC_oz],tblPlant[Date],{d}),\"\")"),
    # ---- gold
    K("Gold_Shipped_oz", "Gold shipped", "oz", "sum", day=sumifs("tblGold", "Gold_oz", ("Type", '"Shipment"'))),
    K("Gold_Sold_oz", "Gold sold", "oz", "sum", day=sumifs("tblGold", "Gold_oz", ("Type", '"Sale"'))),
    K("Dore_Shipped_kg", "Dore shipped", "kg", "sum", day=sumifs("tblGold", "Dore_kg", ("Type", '"Shipment"'))),
    # ---- safety
    K("Hours_Worked", "Hours worked", "h", "sum", day=sumifs("tblSafety", "Hours_Total"), budget="Hours_Worked"),
    K("Fatality", "Fatalities", "int", "sum", day=sumifs("tblSafety", "Fatality")),
    K("LTI", "Lost time injuries", "int", "sum", day=sumifs("tblSafety", "LTI")),
    K("RWI", "Restricted work injuries", "int", "sum", day=sumifs("tblSafety", "RWI")),
    K("MTI", "Medical treatment injuries", "int", "sum", day=sumifs("tblSafety", "MTI")),
    K("FAI", "First aid injuries", "int", "sum", day=sumifs("tblSafety", "FAI")),
    K("Recordables", "Recordable injuries", "int", "sum", day=sumifs("tblSafety", "Recordables")),
    K("Near_Miss", "Near misses", "int", "sum", day=sumifs("tblSafety", "Near_Miss")),
    K("Hazard_Reports", "Hazard reports", "int", "sum", day=sumifs("tblSafety", "Hazard_Reports")),
    K("HPI", "High potential incidents", "int", "sum", day=sumifs("tblSafety", "HPI")),
    K("Env_Incidents", "Environmental incidents", "int", "sum", day=sumifs("tblSafety", "Env_Incidents")),
    K("Community_Incidents", "Community incidents", "int", "sum", day=sumifs("tblSafety", "Community_Incidents")),
    K("Vehicle_Incidents", "Vehicle incidents", "int", "sum", day=sumifs("tblSafety", "Vehicle_Incidents")),
    K("Toolbox_Talks", "Toolbox talks", "int", "sum", day=sumifs("tblSafety", "Toolbox_Talks")),
    K("Inspections", "Inspections and audits", "int", "sum", day=sumifs("tblSafety", "Inspections")),
    K("TRIFR_12m", "TRIFR, rolling 12 months + MTD", "rate", "custom",
      day='IFERROR((SUMIFS(tblSafety[Recordables],tblSafety[Date],">="&EDATE({ms},-12),tblSafety[Date],"<="&{d})'
          '+SUMIFS(tblSafetyHistory[Recordables],tblSafetyHistory[Month],">="&EDATE({ms},-12),tblSafetyHistory[Month],"<"&cfg_YearStart))'
          '/(SUMIFS(tblSafety[Hours_Total],tblSafety[Date],">="&EDATE({ms},-12),tblSafety[Date],"<="&{d})'
          '+SUMIFS(tblSafetyHistory[Hours_Total],tblSafetyHistory[Month],">="&EDATE({ms},-12),tblSafetyHistory[Month],"<"&cfg_YearStart))*1000000,"")'),
    K("LTIFR_12m", "LTIFR, rolling 12 months + MTD", "rate", "custom",
      day='IFERROR((SUMIFS(tblSafety[LTI],tblSafety[Date],">="&EDATE({ms},-12),tblSafety[Date],"<="&{d})'
          '+SUMIFS(tblSafetyHistory[LTI],tblSafetyHistory[Month],">="&EDATE({ms},-12),tblSafetyHistory[Month],"<"&cfg_YearStart))'
          '/(SUMIFS(tblSafety[Hours_Total],tblSafety[Date],">="&EDATE({ms},-12),tblSafety[Date],"<="&{d})'
          '+SUMIFS(tblSafetyHistory[Hours_Total],tblSafetyHistory[Month],">="&EDATE({ms},-12),tblSafetyHistory[Month],"<"&cfg_YearStart))*1000000,"")'),
    K("Days_Since_LTI", "Days since last LTI", "days", "custom",
      day='IF(MAX(N(cfg_PriorLTIDate),_xlfn.MAXIFS(tblSafety[Date],tblSafety[LTI],">0",tblSafety[Date],"<="&{d}))=0,"",'
          '{d}-MAX(N(cfg_PriorLTIDate),_xlfn.MAXIFS(tblSafety[Date],tblSafety[LTI],">0",tblSafety[Date],"<="&{d})))'),
]
KPI = {k["key"]: k for k in KPIS}

# Daily report layout: (section, [(kpi, show_budget)])
REPORT_LAYOUT = [
    ("SAFETY", [("Hours_Worked", True), ("Recordables", False), ("LTI", False), ("RWI", False), ("MTI", False), ("FAI", False),
                ("Near_Miss", False), ("Hazard_Reports", False), ("HPI", False), ("Env_Incidents", False), ("Community_Incidents", False),
                ("Vehicle_Incidents", False), ("TRIFR_12m", False), ("LTIFR_12m", False), ("Days_Since_LTI", False)]),
    ("MINING", [("Ore_Mined_t", True), ("Ore_Grade_gpt", True), ("Ore_Mined_oz", True), ("Waste_t", True), ("TMM_t", True), ("Strip_Ratio", True),
                ("Rehandle_t", False), ("Drill_m", True), ("Blast_t", False), ("Powder_Factor_kgpt", False), ("Diesel_Mining_L", True),
                ("Exc_Availability_pct", False), ("Exc_Utilisation_pct", False), ("Trk_Availability_pct", False), ("Trk_Utilisation_pct", False)]),
    ("PROCESSING", [("Crushed_t", False), ("Milled_t", True), ("Head_Grade_gpt", True), ("Feed_oz", True), ("Recovery_pct", True), ("Recovered_oz", True),
                    ("Gravity_Gold_oz", False), ("Gold_Poured_oz", True), ("Throughput_tph", False), ("Mill_Run_h", False), ("Mill_Availability_pct", False),
                    ("Mill_Utilisation_pct", False), ("Cyanide_kgpt", False), ("Lime_kgpt", False), ("Power_kWhpt", False), ("GIC_oz", False)]),
    ("GOLD", [("Gold_Shipped_oz", False), ("Gold_Sold_oz", False), ("Dore_Shipped_kg", False)]),
]
MONTHLY_KPIS = ["Hours_Worked", "Recordables", "LTI", "TRIFR_12m", "Ore_Mined_t", "Ore_Grade_gpt", "Ore_Mined_oz", "Waste_t", "TMM_t", "Strip_Ratio",
                "Drill_m", "Milled_t", "Head_Grade_gpt", "Recovery_pct", "Recovered_oz", "Gold_Poured_oz", "Gold_Sold_oz", "Throughput_tph",
                "Mill_Availability_pct", "Mill_Utilisation_pct", "Cyanide_kgpt", "Power_kWhpt"]
COMMENT_LINES = 15


# ----------------------------------------------------------------------------- builder
class Builder:
    def __init__(self, schema: dict, year: int, sample_days: int = 0, seed: int = 7, report_date: date | None = None):
        self.s = schema
        self.year = year
        self.sample_days = sample_days
        self.report_date = report_date
        self.rng = random.Random(seed)
        self.wb = openpyxl.Workbook()
        self.wb.remove(self.wb.active)
        self.days = [date(year, 1, 1) + timedelta(days=i) for i in range((date(year, 12, 31) - date(year, 1, 1)).days + 1)]
        self.tables: dict[str, dict] = {}     # table name -> {sheet, first_row, last_row, cols{name: letter}}
        self.cd_cols: dict[str, str] = {}     # Calc_Daily header -> column letter
        self.sample = {}                      # synthetic data (for the validator)
        self.formula_count = 0

    # ---- helpers -----------------------------------------------------------------
    def name(self, nm: str, ref: str):
        self.wb.defined_names[nm] = DefinedName(nm, attr_text=ref)

    def dyn_range(self, sheet: str, col_letter: str, first_row: int) -> str:
        c = col_letter
        return f"{sheet}!${c}${first_row}:INDEX({sheet}!${c}:${c},COUNTA({sheet}!${c}:${c}))"

    def set_widths(self, ws, widths: dict):
        for col, w in widths.items():
            ws.column_dimensions[col].width = w

    def title(self, ws, text: str, sub: str = ""):
        ws["A1"] = text
        ws["A1"].font = font(bold=True, size=14, color=C_HEAD)
        if sub:
            ws["A2"] = sub
            ws["A2"].font = font(italic=True, color="595959")

    def fmt_for(self, field: dict) -> str:
        t = field.get("type")
        if t == "date":
            return FMT["date"]
        if t == "pct":
            return FMT["%"]
        if t == "int":
            return FMT["int"]
        if t == "text":
            return FMT["text"]
        return FMT.get(field.get("unit", ""), FMT["number"] if field.get("unit") is None else FMT["number"])

    def tbl_formula(self, tbl: str, expr: str) -> str:
        """Translate {Col} placeholders into table this-row references."""
        out = expr.replace("{TROY}", TROY)
        import re
        return re.sub(r"\{([A-Za-z_][A-Za-z0-9_]*)\}", lambda m: f"{tbl}[[#This Row],[{m.group(1)}]]", out)

    # ---- tables ------------------------------------------------------------------
    def prefill_rows(self, tdef: dict) -> int:
        p = tdef.get("prefill")
        if p == "dates":
            return len(self.days)
        if p in ("months", "prior_months"):
            return 12
        return int(tdef.get("rows", 100))

    def write_table(self, ws, tdef: dict, anchor_col: int = 1, anchor_row: int = 1, data: list[list] | None = None, style="TableStyleMedium2"):
        """Row anchor_row = labels, anchor_row+1 = header (field names), then data rows. Returns table info."""
        tbl = tdef["table"]
        fields = tdef["fields"]
        nrows = self.prefill_rows(tdef)
        hdr_row = anchor_row + 1
        first = hdr_row + 1
        last = first + nrows - 1
        cols = {}
        for j, f in enumerate(fields):
            c = anchor_col + j
            L = get_column_letter(c)
            cols[f["name"]] = L
            lab = f.get("label", f["name"]) + (f" ({f['unit']})" if f.get("unit") else "")
            lc = ws.cell(anchor_row, c, lab)
            lc.font = font(italic=True, color="595959", size=9)
            lc.fill = fill(C_LABEL)
            lc.alignment = Alignment(wrap_text=True, vertical="bottom")
            hc = ws.cell(hdr_row, c, f["name"])
            hc.font = font(bold=True, color="FFFFFF")
            hc.fill = fill(C_HEAD_CALC if f.get("calc") else C_HEAD)
            hc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.column_dimensions[L].width = max(11, min(28, len(lab) * 0.9 + 2)) if f.get("type") != "text" else 32
        ws.row_dimensions[anchor_row].height = 30
        # prefill / data
        p = tdef.get("prefill")
        key = fields[0]["name"]
        for i in range(nrows):
            r = first + i
            if p == "dates":
                ws.cell(r, anchor_col, self.days[i])
            elif p == "months":
                ws.cell(r, anchor_col, date(self.year, i + 1, 1))
            elif p == "prior_months":
                ws.cell(r, anchor_col, date(self.year - 1, i + 1, 1))
        if data:
            by_key = {}
            if p in ("dates", "months", "prior_months"):
                for row in data:
                    by_key[row[0]] = row
                for i in range(nrows):
                    k = ws.cell(first + i, anchor_col).value
                    row = by_key.get(k)
                    if row:
                        for j, v in enumerate(row):
                            if j == 0 or v is None:
                                continue
                            ws.cell(first + i, anchor_col + j, v)
            else:
                for i, row in enumerate(data[:nrows]):
                    for j, v in enumerate(row):
                        if v is not None:
                            ws.cell(first + i, anchor_col + j, v)
        # formats, formulas, validation
        for j, f in enumerate(fields):
            c = anchor_col + j
            L = get_column_letter(c)
            nf = self.fmt_for(f)
            calc = f.get("calc")
            for r in range(first, last + 1):
                cell = ws.cell(r, c)
                cell.number_format = nf
                if calc:
                    cell.value = "=" + self.tbl_formula(tbl, calc)
                    cell.font = font(color="595959")
                    self.formula_count += 1
            if not calc:
                dv = self.dv_for(f)
                if dv is not None:
                    ws.add_data_validation(dv)
                    dv.add(f"{L}{first}:{L}{last}")
        ref = f"{get_column_letter(anchor_col)}{hdr_row}:{get_column_letter(anchor_col + len(fields) - 1)}{last}"
        t = Table(displayName=tbl, ref=ref)
        t.tableStyleInfo = TableStyleInfo(name=style, showRowStripes=True, showFirstColumn=False, showLastColumn=False, showColumnStripes=False)
        t._initialise_columns()
        for col, f in zip(t.tableColumns, fields):
            col.name = f["name"]
            if f.get("calc"):
                col.calculatedColumnFormula = TableFormula(attr_text=self.tbl_formula(tbl, f["calc"]))
        ws.add_table(t)
        info = {"sheet": ws.title, "hdr_row": hdr_row, "first": first, "last": last, "cols": cols, "fields": fields}
        self.tables[tbl] = info
        return info

    def dv_for(self, f: dict):
        t = f.get("type")
        msg = f.get("desc", "")
        if t == "list" and f.get("list"):
            dv = DataValidation(type="list", formula1=f"=lst_{f['list']}", allow_blank=True, showErrorMessage=True, showInputMessage=bool(msg))
            dv.errorTitle = "Not in list"
            dv.error = f"Choose a value from the list (edit the Lists sheet to add one)."
        elif t in ("number", "pct", "int"):
            lo, hi = f.get("min", 0), f.get("max", 1e12)
            dv = DataValidation(type="whole" if t == "int" else "decimal", operator="between", formula1=str(lo), formula2=str(hi),
                                allow_blank=True, showErrorMessage=True, errorStyle="warning", showInputMessage=bool(msg))
            dv.errorTitle = "Outside expected range"
            dv.error = f"Expected between {lo} and {hi}. Click Yes to keep the value if it is correct."
        elif t == "date":
            dv = DataValidation(type="date", operator="between", formula1="=cfg_YearStart-31", formula2="=cfg_YearEnd", allow_blank=True,
                                showErrorMessage=True, errorStyle="warning", showInputMessage=bool(msg))
            dv.errorTitle = "Date outside the reporting year"
            dv.error = "Dates should fall in the reporting year (December of the prior year is tolerated)."
        else:
            return None
        if msg:
            dv.promptTitle = f.get("label", f["name"])[:32]
            dv.prompt = msg[:250]
        return dv

    # ---- sheets ------------------------------------------------------------------
    def build(self):
        site = self.s["site"]
        wb = self.wb
        # order: README, reports, inputs, config/lists, calc
        ws_readme = wb.create_sheet("README")
        ws_rep = wb.create_sheet("Daily_Report")
        ws_dash = wb.create_sheet("Dashboard")
        ws_month = wb.create_sheet("Monthly_Summary")
        input_sheets = {}
        for tdef in self.s["tables"]:
            if tdef["sheet"] in ("Config", "Lists"):
                continue
            if tdef["sheet"] not in input_sheets:
                input_sheets[tdef["sheet"]] = wb.create_sheet(tdef["sheet"])
        ws_cfg = wb.create_sheet("Config")
        ws_lists = wb.create_sheet("Lists")
        ws_cd = wb.create_sheet("Calc_Daily")
        ws_chart = wb.create_sheet("Chart_Data")
        ws_dict = wb.create_sheet("Data_Dictionary")

        sample = self.make_sample() if self.sample_days else {}
        self.sample = sample

        self.sheet_lists(ws_lists)
        self.sheet_config(ws_cfg, sample)
        for tdef in self.s["tables"]:
            if tdef["sheet"] in input_sheets:
                self.sheet_input(input_sheets[tdef["sheet"]], tdef, sample.get(tdef["table"]))
        self.sheet_calc_daily(ws_cd)
        self.sheet_chart_data(ws_chart)
        self.sheet_daily_report(ws_rep)
        self.sheet_monthly(ws_month)
        self.sheet_dashboard(ws_dash, ws_chart)
        self.sheet_dictionary(ws_dict)
        self.sheet_readme(ws_readme)

        for ws in (ws_rep, ws_dash, ws_month):
            ws.sheet_properties.tabColor = TAB["report"]
        for ws in input_sheets.values():
            ws.sheet_properties.tabColor = TAB["input"]
        for ws in (ws_cfg, ws_lists):
            ws.sheet_properties.tabColor = TAB["config"]
        for ws in (ws_cd, ws_chart, ws_dict):
            ws.sheet_properties.tabColor = TAB["calc"]
        wb.active = wb.sheetnames.index("Daily_Report")
        wb.calculation.fullCalcOnLoad = True
        wb.properties.title = f"{site['name']} Production System {self.year}"
        wb.properties.creator = "Mako Production System"
        wb.properties.subject = "Daily production reporting"
        return wb

    # Lists ------------------------------------------------------------------------
    def sheet_lists(self, ws):
        self.title(ws, "Lists", "Drop-down values. Add rows at the bottom of a list; do not rename headers. Every list is referenced by name (lst_...) in data validation.")
        col = 1
        row0 = 4
        all_lists = dict(self.s["lists"])
        all_lists.update(self.s.get("derived_lists", {}))
        for lname, values in all_lists.items():
            tdef = {"table": f"tblList_{lname}", "fields": [{"name": lname, "label": lname.replace("_", " "), "type": "text"}], "rows": len(values)}
            info = self.write_table(ws, tdef, anchor_col=col, anchor_row=row0, data=[[v] for v in values], style="TableStyleLight9")
            L = get_column_letter(col)
            self.name(f"lst_{lname}", self.dyn_range("Lists", L, info["first"]))
            ws.column_dimensions[L].width = max(16, max(len(v) for v in values) + 2)
            col += 2
        col += 1
        for rname, rdef in self.s["reference_tables"].items():
            if rdef["sheet"] != "Lists":
                continue
            tdef = {"table": rdef["table"], "fields": rdef["columns"], "rows": len(rdef["rows"])}
            info = self.write_table(ws, tdef, anchor_col=col, anchor_row=row0, data=rdef["rows"], style="TableStyleLight9")
            keyL = info["cols"][rdef["key"]]
            self.name(f"lst_{rdef['key']}", self.dyn_range("Lists", keyL, info["first"]))
            ws.column_dimensions[keyL].width = 26
            col += len(rdef["columns"]) + 1
        ws.freeze_panes = "A6"

    # Config -----------------------------------------------------------------------
    def sheet_config(self, ws, sample):
        site = self.s["site"]
        self.title(ws, "Config", "Site settings and opening balances. Names in column C are used by the formulas.")
        settings = [
            ("Site", site["name"], "cfg_Site", "text"),
            ("Company", site["company"], "cfg_Company", "text"),
            ("Reporting year", self.year, "cfg_Year", "int"),
            ("Year start", "=DATE(cfg_Year,1,1)", "cfg_YearStart", "date"),
            ("Year end", "=DATE(cfg_Year,12,31)", "cfg_YearEnd", "date"),
            ("Troy ounce (g)", self.s["constants"]["TROY_OZ_G"], "cfg_TroyOz", "number"),
            ("Date of last LTI before this year", sample.get("prior_lti") if sample else None, "cfg_PriorLTIDate", "date"),
            ("Last date with plant data (calculated)", '=SUMPRODUCT(MAX((tblPlant[Milled_t]<>"")*tblPlant[Date]))', "cfg_LastDataDate", "date"),
            ("Report title", f"{site['name']} daily production report", "cfg_ReportTitle", "text"),
        ]
        r = 4
        ws.cell(r, 1, "Setting").font = font(bold=True); ws.cell(r, 2, "Value").font = font(bold=True); ws.cell(r, 3, "Name").font = font(bold=True)
        for lab, val, nm, typ in settings:
            r += 1
            ws.cell(r, 1, lab)
            c = ws.cell(r, 2, val)
            c.number_format = FMT["date"] if typ == "date" else ("0" if typ == "int" else ("0.0000" if typ == "number" else "@"))
            c.font = font(color="1F3864" if not (isinstance(val, str) and val.startswith("=")) else "595959")
            c.border = BORDER
            ws.cell(r, 3, nm).font = font(color="7F7F7F", size=9)
            self.name(nm, f"Config!$B${r}")
        self.set_widths(ws, {"A": 36, "B": 30, "C": 18})
        # reference tables on Config (stockpiles) and prior-year safety history
        col = 5
        for rname, rdef in self.s["reference_tables"].items():
            if rdef["sheet"] != "Config":
                continue
            rows = sample.get(rdef["table"], rdef["rows"]) if sample else rdef["rows"]
            tdef = {"table": rdef["table"], "title": "Stockpile opening balances (survey at start of year)", "fields": rdef["columns"], "rows": len(rows)}
            ws.cell(3, col, tdef["title"]).font = font(bold=True, color=C_HEAD)
            info = self.write_table(ws, tdef, anchor_col=col, anchor_row=4, data=rows, style="TableStyleLight9")
            col += len(rdef["columns"]) + 1
        for tdef in self.s["tables"]:
            if tdef["sheet"] == "Config":
                ws.cell(3, col, tdef["title"]).font = font(bold=True, color=C_HEAD)
                self.write_table(ws, tdef, anchor_col=col, anchor_row=4, data=sample.get(tdef["table"]) if sample else None, style="TableStyleLight9")
                col += len(tdef["fields"]) + 1

    # Input sheets -----------------------------------------------------------------
    def sheet_input(self, ws, tdef, data):
        if ws["A1"].value is None:
            self.title(ws, tdef["title"], f"Owner: {tdef.get('owner', '')}. Yellow-free rule: type values only in the blue-headed columns; grey-headed columns are formulas.")
            anchor_row = 4
        else:
            anchor_row = ws.max_row + 3
        info = self.write_table(ws, tdef, anchor_col=1, anchor_row=anchor_row, data=data)
        ws.freeze_panes = ws.cell(info["first"], 2)

    # Calc_Daily -------------------------------------------------------------------
    def plan_calc_columns(self):
        """Assign Calc_Daily columns. Returns ordered list of (header, kind, kpi, period, is_budget)."""
        cols = [("Date", "base"), ("Month", "base"), ("Day", "base"), ("Days_In_Month", "base"), ("Month_Start", "base")]
        for k in KPIS:
            kind = k["kind"]
            if kind in ("sum", "ratio"):
                cols += [(k["key"], "actual"), (k["key"] + "_MTD", "actual"), (k["key"] + "_YTD", "actual")]
            else:
                cols += [(k["key"], "actual")]
        for k in KPIS:
            if self.has_budget(k):
                cols += [(f"Bud_{k['key']}_Month", "budget"), (f"Bud_{k['key']}_MTD", "budget"), (f"Bud_{k['key']}_YTD", "budget")]
        self.cd_cols = {h: get_column_letter(i + 1) for i, (h, _) in enumerate(cols)}
        return cols

    def has_budget(self, k) -> bool:
        if k.get("budget"):
            return True
        if k["kind"] == "ratio":
            import re
            refs = set(re.findall(r"\{([A-Za-z_][A-Za-z0-9_]*)\}", k["num"] + k["den"]))
            return all(self.has_budget(KPI[x]) for x in refs if x in KPI)
        return False

    def resolve(self, expr: str, r: int, period: str = "", budget: bool = False) -> str:
        import re
        c = self.cd_cols

        def rep(m):
            key = m.group(1)
            if key == "d":
                return f"$A{r}"
            if key == "ms":
                return f"$E{r}"
            if key == "dim":
                return f"$D{r}"
            if key == "r":
                return str(r)
            if key == "TROY":
                return TROY
            if key in KPI:
                if budget:
                    hdr = f"Bud_{key}_{period}"
                else:
                    hdr = key if period in ("", "Day") else f"{key}_{period}"
                return f"${c[hdr]}{r}"
            raise KeyError(f"unknown placeholder {{{key}}} in {expr}")

        return re.sub(r"\{([A-Za-z_][A-Za-z0-9_]*)\}", rep, expr)

    def sheet_calc_daily(self, ws):
        cols = self.plan_calc_columns()
        ws.append([h for h, _ in cols])
        for j, (h, kind) in enumerate(cols, start=1):
            c = ws.cell(1, j)
            c.font = font(bold=True, color="FFFFFF", size=9)
            c.fill = fill(C_HEAD if kind != "budget" else C_HEAD_CALC)
            c.alignment = Alignment(wrap_text=True, vertical="center")
            ws.column_dimensions[get_column_letter(j)].width = 12
        ws.row_dimensions[1].height = 42
        n = len(self.days)
        first, last = 2, n + 1
        c = self.cd_cols
        for i, d in enumerate(self.days):
            r = first + i
            ws.cell(r, 1, d).number_format = FMT["date"]
            ws.cell(r, 2, f"=MONTH($A{r})")
            ws.cell(r, 3, f"=DAY($A{r})")
            ws.cell(r, 4, f"=DAY(EOMONTH($A{r},0))")
            ws.cell(r, 5, f"=DATE(YEAR($A{r}),MONTH($A{r}),1)").number_format = FMT["date"]
            for k in KPIS:
                key, kind = k["key"], k["kind"]
                nf = FMT.get(k["unit"], FMT["number"])
                if kind == "sum":
                    day = "=" + self.resolve(k["day"], r)
                    mtd = f"={key}_ROW" if False else (f"=${c[key]}{r}" if r == first else f"=IF($B{r}=$B{r-1},${c[key + '_MTD']}{r-1},0)+${c[key]}{r}")
                    ytd = f"=${c[key]}{r}" if r == first else f"=${c[key + '_YTD']}{r-1}+${c[key]}{r}"
                    for hdr, f_ in ((key, day), (key + "_MTD", mtd), (key + "_YTD", ytd)):
                        cell = ws[f"{c[hdr]}{r}"]; cell.value = f_; cell.number_format = nf
                elif kind == "ratio":
                    for period in ("Day", "MTD", "YTD"):
                        hdr = key if period == "Day" else f"{key}_{period}"
                        num = self.resolve(k["num"], r, period)
                        den = self.resolve(k["den"], r, period)
                        cell = ws[f"{c[hdr]}{r}"]
                        cell.value = f'=IFERROR(IF(({den})>0,({num})/({den}),""),"")'
                        cell.number_format = nf
                else:
                    cell = ws[f"{c[key]}{r}"]
                    cell.value = "=" + self.resolve(k["day"], r)
                    cell.number_format = nf
            for k in KPIS:
                if not self.has_budget(k):
                    continue
                key = k["key"]
                nf = FMT.get(k["unit"], FMT["number"])
                if k.get("budget"):
                    bcol = k["budget"]
                    month = f"=SUMIFS(tblBudget[{bcol}],tblBudget[Month],$E{r})"
                    mtd = f"=${c[f'Bud_{key}_Month']}{r}*$C{r}/$D{r}"
                    ytd = f'=SUMIFS(tblBudget[{bcol}],tblBudget[Month],"<"&$E{r})+${c[f"Bud_{key}_MTD"]}{r}'
                    for period, f_ in (("Month", month), ("MTD", mtd), ("YTD", ytd)):
                        cell = ws[f"{c[f'Bud_{key}_{period}']}{r}"]; cell.value = f_; cell.number_format = nf
                else:  # ratio budget
                    for period in ("Month", "MTD", "YTD"):
                        num = self.resolve(k["num"], r, period, budget=True)
                        den = self.resolve(k["den"], r, period, budget=True)
                        cell = ws[f"{c[f'Bud_{key}_{period}']}{r}"]
                        cell.value = f'=IFERROR(IF(({den})>0,({num})/({den}),""),"")'
                        cell.number_format = nf
        self.formula_count += (last - first + 1) * (len(cols) - 1)
        for h, _ in cols:
            L = c[h]
            self.name(f"cd_{h}", f"Calc_Daily!${L}${first}:${L}${last}")
        ws.freeze_panes = "B2"
        ws.protection.sheet = True
        self.cd_rows = (first, last)

    # Chart_Data -------------------------------------------------------------------
    def sheet_chart_data(self, ws):
        first, last = self.cd_rows
        series = [
            ("Date", None, False),
            ("Milled_t", "Milled_t", True),
            ("Budget_Milled_daily", "Bud_Milled_t_Month/DIM", False),
            ("Gold_Poured_YTD", "Gold_Poured_oz_YTD", True),
            ("Budget_Gold_Poured_YTD", "Bud_Gold_Poured_oz_YTD", False),
            ("Ore_Mined_t", "Ore_Mined_t", True),
            ("Waste_t", "Waste_t", True),
            ("Recovery_MTD", "Recovery_pct_MTD", True),
            ("Head_Grade_MTD", "Head_Grade_gpt_MTD", True),
            ("TRIFR_12m", "TRIFR_12m", True),
            ("Budget_Recovery_MTD", "Bud_Recovery_pct_MTD", False),
        ]
        ws.append([s[0] for s in series])
        for j in range(1, len(series) + 1):
            ws.cell(1, j).font = font(bold=True, color="FFFFFF", size=9); ws.cell(1, j).fill = fill(C_HEAD)
            ws.column_dimensions[get_column_letter(j)].width = 14
        c = self.cd_cols
        for r in range(first, last + 1):
            ws.cell(r, 1, f"=Calc_Daily!$A{r}").number_format = FMT["date"]
            for j, (hdr, src, guard) in enumerate(series[1:], start=2):
                if src == "Bud_Milled_t_Month/DIM":
                    val = f"Calc_Daily!${c['Bud_Milled_t_Month']}{r}/Calc_Daily!$D{r}"
                else:
                    val = f"Calc_Daily!${c[src]}{r}"
                f_ = f"=IF($A{r}>cfg_LastDataDate,NA(),{val})" if guard else f"={val}"
                ws.cell(r, j, f_)
        self.formula_count += (last - first + 1) * len(series)
        ws.protection.sheet = True
        self.chart_series = series

    # Daily_Report -----------------------------------------------------------------
    def sheet_daily_report(self, ws):
        c = self.cd_cols
        ws.sheet_view.showGridLines = False
        self.set_widths(ws, {"A": 2, "B": 34, "C": 8, "D": 13, "E": 13, "F": 13, "G": 9, "H": 13, "I": 13, "J": 9, "K": 2, "L": 10})
        ws["B2"] = "=cfg_ReportTitle"; ws["B2"].font = font(bold=True, size=16, color=C_HEAD)
        ws["B3"] = "=cfg_Company&\" | \"&cfg_Site"; ws["B3"].font = font(italic=True, color="595959")
        ws["B5"] = "Report date"; ws["B5"].font = font(bold=True)
        ws["C5"] = self.report_date if self.report_date else "=MAX(cfg_YearStart,MIN(TODAY()-1,cfg_YearEnd,IF(N(cfg_LastDataDate)>0,cfg_LastDataDate,cfg_YearEnd)))"
        ws["C5"].number_format = "ddd dd-mmm-yyyy"; ws["C5"].font = font(bold=True, color=C_HEAD); ws["C5"].fill = fill(C_KEY); ws["C5"].border = BORDER
        ws["C5"].protection = Protection(locked=False)
        ws.merge_cells("C5:E5")
        ws["F5"] = "Type a date to change the report day. Enter =TODAY()-1 to restore the default."; ws["F5"].font = font(italic=True, size=9, color="7F7F7F")
        ws["L5"] = "=IFERROR(MATCH($C$5,cd_Date,0),NA())"; ws["L5"].font = font(color="BFBFBF", size=8)
        ws["L4"] = "row"; ws["L4"].font = font(color="BFBFBF", size=8)
        self.name("rpt_Date", "Daily_Report!$C$5")
        self.name("rpt_Row", "Daily_Report!$L$5")
        ws["B6"] = '="Data to "&TEXT(cfg_LastDataDate,"dd mmm yyyy")&"  |  Week "&_xlfn.ISOWEEKNUM($C$5)&"  |  Day "&DAY($C$5)&" of "&DAY(EOMONTH($C$5,0))'
        ws["B6"].font = font(size=9, color="595959")

        headers = ["", "Unit", "Day", "MTD", "MTD budget", "Var %", "YTD", "YTD budget", "Var %"]
        r = 8
        for section, items in REPORT_LAYOUT:
            for j, h in enumerate(headers):
                cell = ws.cell(r, 2 + j, section if j == 0 else h)
                cell.font = font(bold=True, color="FFFFFF"); cell.fill = fill(C_HEAD)
                cell.alignment = Alignment(horizontal="left" if j == 0 else "center")
            r += 1
            for key, show_budget in items:
                k = KPI[key]
                nf = FMT.get(k["unit"], FMT["number"])
                ws.cell(r, 2, k["label"]).font = font()
                ws.cell(r, 3, "" if k["unit"] in ("int", "ratio", "rate", "days") else k["unit"]).font = font(color="7F7F7F", size=9)
                has_periods = k["kind"] in ("sum", "ratio")
                def idx(hdr):
                    return f'=IFERROR(INDEX(cd_{hdr},rpt_Row),"")'
                ws.cell(r, 4, idx(key)).number_format = nf
                if has_periods:
                    ws.cell(r, 5, idx(key + "_MTD")).number_format = nf
                    ws.cell(r, 8, idx(key + "_YTD")).number_format = nf
                if show_budget and self.has_budget(k):
                    ws.cell(r, 6, idx(f"Bud_{key}_MTD")).number_format = nf
                    ws.cell(r, 9, idx(f"Bud_{key}_YTD")).number_format = nf
                    ws.cell(r, 7, f'=IF(AND(ISNUMBER(F{r}),ISNUMBER(E{r}),F{r}<>0),E{r}/F{r}-1,"")').number_format = "+0.0%;-0.0%;0.0%"
                    ws.cell(r, 10, f'=IF(AND(ISNUMBER(I{r}),ISNUMBER(H{r}),I{r}<>0),H{r}/I{r}-1,"")').number_format = "+0.0%;-0.0%;0.0%"
                for j in range(2, 11):
                    ws.cell(r, j).border = BORDER
                    if j >= 4:
                        ws.cell(r, j).alignment = Alignment(horizontal="right")
                r += 1
            r += 1
        # stockpiles
        sp = self.tables["tblStockpiles"]
        n_sp = sp["last"] - sp["first"] + 1
        hdr = ["STOCKPILES (to report date)", "", "Opening t", "Opening g/t", "In t", "Out t", "Plant feed t", "Balance t", "Balance g/t"]
        for j, h in enumerate(hdr):
            cell = ws.cell(r, 2 + j, h); cell.font = font(bold=True, color="FFFFFF"); cell.fill = fill(C_HEAD)
            cell.alignment = Alignment(horizontal="left" if j == 0 else "center")
        r += 1
        for i in range(1, n_sp + 1):
            nm = f"INDEX(tblStockpiles[Stockpile],{i})"
            ws.cell(r, 2, f"={nm}")
            ws.cell(r, 4, f"=INDEX(tblStockpiles[Opening_t],{i})").number_format = FMT["t"]
            ws.cell(r, 5, f"=INDEX(tblStockpiles[Opening_gpt],{i})").number_format = FMT["g/t"]
            ws.cell(r, 6, f'=SUMIFS(tblMovement[Tonnes_t],tblMovement[Destination],{nm},tblMovement[Date],"<="&rpt_Date)').number_format = FMT["t"]
            ws.cell(r, 7, f'=SUMIFS(tblMovement[Tonnes_t],tblMovement[Source],{nm},tblMovement[Date],"<="&rpt_Date)').number_format = FMT["t"]
            ws.cell(r, 8, f'=IF(INDEX(tblStockpiles[Plant_Feed],{i})="Y",SUMIFS(tblPlant[Milled_t],tblPlant[Date],"<="&rpt_Date)-SUMIFS(tblMovement[Tonnes_t],tblMovement[Dest_Type],"Plant",tblMovement[Date],"<="&rpt_Date),0)').number_format = FMT["t"]
            ws.cell(r, 9, f"=D{r}+F{r}-G{r}-H{r}").number_format = FMT["t"]
            # ounces: opening + in - out - plant feed (feed oz less direct tip oz)
            oz = (f'(D{r}*E{r}/{TROY}'
                  f'+SUMIFS(tblMovement[Contained_oz],tblMovement[Destination],{nm},tblMovement[Date],"<="&rpt_Date)'
                  f'-SUMIFS(tblMovement[Contained_oz],tblMovement[Source],{nm},tblMovement[Date],"<="&rpt_Date)'
                  f'-IF(INDEX(tblStockpiles[Plant_Feed],{i})="Y",SUMIFS(tblPlant[Feed_oz],tblPlant[Date],"<="&rpt_Date)-SUMIFS(tblMovement[Contained_oz],tblMovement[Dest_Type],"Plant",tblMovement[Date],"<="&rpt_Date),0))')
            ws.cell(r, 10, f'=IFERROR(IF(I{r}>0,{oz}*{TROY}/I{r},""),"")').number_format = FMT["g/t"]
            for j in range(2, 11):
                ws.cell(r, j).border = BORDER
            r += 1
        r += 1
        # commentary
        cell = ws.cell(r, 2, "COMMENTARY"); cell.font = font(bold=True, color="FFFFFF"); cell.fill = fill(C_HEAD)
        for j in range(3, 11):
            ws.cell(r, j).fill = fill(C_HEAD)
        ws.cell(r, 3, "Area").font = font(bold=True, color="FFFFFF")
        r += 1
        for n in range(1, COMMENT_LINES + 1):
            ws.cell(r, 2, f'=IFERROR(INDEX(tblCommentary[Area],MATCH(rpt_Date&"|"&{n},tblCommentary[Key_Day],0)),"")').font = font(bold=True, size=9)
            ws.cell(r, 3, f'=IFERROR(INDEX(tblCommentary[Comment],MATCH(rpt_Date&"|"&{n},tblCommentary[Key_Day],0)),"")').font = font(size=9)
            ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=10)
            ws.cell(r, 3).alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 24
            r += 1
        r += 1
        ws.cell(r, 2, '="Generated by the Mako Production System workbook. Printed "&TEXT(NOW(),"dd mmm yyyy hh:mm")').font = font(italic=True, size=8, color="7F7F7F")
        # print setup
        ws.print_area = f"B2:J{r}"
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        ws.print_options.horizontalCentered = True
        ws.page_margins.left = ws.page_margins.right = 0.4
        ws.protection.sheet = True
        self.report_last_row = r

    # Monthly_Summary --------------------------------------------------------------
    def sheet_monthly(self, ws):
        ws.sheet_view.showGridLines = False
        self.set_widths(ws, {"A": 2, "B": 34, "C": 10, **{get_column_letter(i): 11 for i in range(4, 17)}})
        ws["B2"] = '=cfg_Site&" monthly summary "&cfg_Year'; ws["B2"].font = font(bold=True, size=16, color=C_HEAD)
        ws["B3"] = '="Actuals to "&TEXT(cfg_LastDataDate,"dd mmm yyyy")&". Budget from the Budget sheet. Ratios are month values, not averages of days."'
        ws["B3"].font = font(italic=True, size=9, color="595959")
        r = 5
        ws.cell(r, 2, "KPI").font = font(bold=True, color="FFFFFF"); ws.cell(r, 2).fill = fill(C_HEAD)
        ws.cell(r, 3, "").fill = fill(C_HEAD)
        for m in range(1, 13):
            cell = ws.cell(r, 3 + m, date(self.year, m, 1)); cell.number_format = "mmm"; cell.font = font(bold=True, color="FFFFFF"); cell.fill = fill(C_HEAD)
            cell.alignment = Alignment(horizontal="center")
        cell = ws.cell(r, 16, "Year"); cell.font = font(bold=True, color="FFFFFF"); cell.fill = fill(C_HEAD); cell.alignment = Alignment(horizontal="center")
        r += 1
        for key in MONTHLY_KPIS:
            k = KPI[key]
            nf = FMT.get(k["unit"], FMT["number"])
            has_b = self.has_budget(k)
            rows = [("Actual", "actual")] + ([("Budget", "budget"), ("Variance", "var")] if has_b else [])
            first_r = r
            for lab, kind in rows:
                ws.cell(r, 2, k["label"] if kind == "actual" else "").font = font(bold=(kind == "actual"))
                ws.cell(r, 3, lab).font = font(size=9, color="7F7F7F")
                for m in range(1, 13):
                    col = 3 + m
                    ms = f"DATE(cfg_Year,{m},1)"
                    me = f"MIN(EOMONTH({ms},0),cfg_LastDataDate)"
                    if kind == "actual":
                        src = f"cd_{key}_MTD" if k["kind"] in ("sum", "ratio") else f"cd_{key}"
                        f_ = f'=IF({ms}>cfg_LastDataDate,"",IFERROR(INDEX({src},MATCH({me},cd_Date,0)),""))'
                    elif kind == "budget":
                        f_ = f'=IFERROR(INDEX(cd_Bud_{key}_Month,MATCH({ms},cd_Date,0)),"")'
                    else:
                        f_ = f'=IF(AND(ISNUMBER({get_column_letter(col)}{first_r}),ISNUMBER({get_column_letter(col)}{first_r+1}),{get_column_letter(col)}{first_r+1}<>0),{get_column_letter(col)}{first_r}/{get_column_letter(col)}{first_r+1}-1,"")'
                    cell = ws.cell(r, col, f_)
                    cell.number_format = "+0.0%;-0.0%;0.0%" if kind == "var" else nf
                    cell.border = BORDER
                    if kind == "var":
                        cell.font = font(size=9, color="595959")
                # year column
                if kind == "actual":
                    src = f"cd_{key}_YTD" if k["kind"] in ("sum", "ratio") else f"cd_{key}"
                    f_ = f'=IFERROR(INDEX({src},MATCH(MIN(cfg_YearEnd,cfg_LastDataDate),cd_Date,0)),"")'
                elif kind == "budget":
                    f_ = '=IFERROR(INDEX(cd_Bud_%s_YTD,MATCH(cfg_YearEnd,cd_Date,0)),"")' % key
                else:
                    f_ = f'=IF(AND(ISNUMBER(P{first_r}),ISNUMBER(P{first_r+1}),P{first_r+1}<>0),P{first_r}/P{first_r+1}-1,"")'
                cell = ws.cell(r, 16, f_); cell.number_format = "+0.0%;-0.0%;0.0%" if kind == "var" else nf; cell.border = BORDER; cell.font = font(bold=True)
                r += 1
            r += 1 if has_b else 0
        ws.freeze_panes = "D6"
        ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        ws.protection.sheet = True

    # Dashboard --------------------------------------------------------------------
    def sheet_dashboard(self, ws, ws_chart):
        ws.sheet_view.showGridLines = False
        ws["B2"] = '=cfg_Site&" dashboard "&cfg_Year'; ws["B2"].font = font(bold=True, size=16, color=C_HEAD)
        ws["B3"] = '="Data to "&TEXT(cfg_LastDataDate,"dd mmm yyyy")'; ws["B3"].font = font(italic=True, size=9, color="595959")
        first, last = self.cd_rows
        hdr = {s[0]: i + 1 for i, s in enumerate(self.chart_series)}

        def ref(col):
            return Reference(ws_chart, min_col=hdr[col], min_row=1, max_row=last)

        dates = Reference(ws_chart, min_col=1, min_row=first, max_row=last)

        def style(ch, title, y_title, height=8.5, width=24):
            ch.title = title; ch.height = height; ch.width = width
            ch.y_axis.title = y_title; ch.legend.position = "b"
            ch.x_axis.number_format = "mmm"; ch.x_axis.majorTimeUnit = "months"
            ch.y_axis.majorGridlines = None
        # 1 gold poured YTD vs budget
        ch1 = LineChart(); ch1.add_data(ref("Gold_Poured_YTD"), titles_from_data=True); ch1.add_data(ref("Budget_Gold_Poured_YTD"), titles_from_data=True)
        ch1.set_categories(dates); ch1.x_axis = DateAxis(crossAx=100); style(ch1, "Gold poured, year to date (oz)", "oz")
        ch1.series[1].graphicalProperties.line.dashStyle = "dash"
        ws.add_chart(ch1, "B5")
        # 2 milled daily vs budget daily
        ch2 = BarChart(); ch2.type = "col"; ch2.add_data(ref("Milled_t"), titles_from_data=True); ch2.set_categories(dates)
        ln = LineChart(); ln.add_data(ref("Budget_Milled_daily"), titles_from_data=True); ln.set_categories(dates)
        ch2 += ln; style(ch2, "Milled tonnes per day vs budget", "t"); ch2.gapWidth = 30
        ws.add_chart(ch2, "N5")
        # 3 ore and waste
        ch3 = BarChart(); ch3.type = "col"; ch3.grouping = "stacked"; ch3.overlap = 100
        ch3.add_data(ref("Ore_Mined_t"), titles_from_data=True); ch3.add_data(ref("Waste_t"), titles_from_data=True); ch3.set_categories(dates)
        style(ch3, "Ore and waste mined per day (t)", "t"); ch3.gapWidth = 30
        ws.add_chart(ch3, "B24")
        # 4 recovery MTD and head grade MTD
        ch4 = LineChart(); ch4.add_data(ref("Recovery_MTD"), titles_from_data=True); ch4.add_data(ref("Budget_Recovery_MTD"), titles_from_data=True); ch4.set_categories(dates)
        ch4.x_axis = DateAxis(crossAx=100); style(ch4, "Recovery, month to date (%)", "%"); ch4.y_axis.number_format = "0%"
        ch4.series[1].graphicalProperties.line.dashStyle = "dash"
        ch5 = LineChart(); ch5.add_data(ref("Head_Grade_MTD"), titles_from_data=True); ch5.set_categories(dates)
        ch5.y_axis.axId = 200; ch5.y_axis.title = "g/t"; ch5.y_axis.crosses = "max"; ch5.y_axis.majorGridlines = None
        ch4 += ch5
        ws.add_chart(ch4, "N24")
        # 5 TRIFR
        ch6 = LineChart(); ch6.add_data(ref("TRIFR_12m"), titles_from_data=True); ch6.set_categories(dates); ch6.x_axis = DateAxis(crossAx=100)
        style(ch6, "TRIFR, rolling 12 months", "per million hours")
        ws.add_chart(ch6, "B43")
        ws.protection.sheet = True

    # Data_Dictionary --------------------------------------------------------------
    def sheet_dictionary(self, ws):
        self.title(ws, "Data dictionary", "Generated from mps_schema.json. The same definitions drive the SQL schema of the web application.")
        hdr = ["Table", "Sheet", "Field", "Label", "Type", "Unit", "Required", "Validation", "Description", "Formula"]
        r = 4
        for j, h in enumerate(hdr, start=1):
            c = ws.cell(r, j, h); c.font = font(bold=True, color="FFFFFF"); c.fill = fill(C_HEAD)
        for tdef in self.s["tables"]:
            for f in tdef["fields"]:
                r += 1
                val = f"list: {f['list']}" if f.get("list") else (f"{f.get('min', '')} to {f.get('max', '')}" if "min" in f or "max" in f else "")
                ws.append([])  # keep row structure simple
                for j, v in enumerate([tdef["table"], tdef["sheet"], f["name"], f.get("label", ""), "calculated" if f.get("calc") else f["type"],
                                       f.get("unit", ""), "yes" if f.get("required") or f.get("key") else "", val, f.get("desc", ""), f.get("calc", "")], start=1):
                    ws.cell(r, j, v).font = font(size=9)
        r += 2
        ws.cell(r, 1, "KPI catalogue (Calc_Daily)").font = font(bold=True, color=C_HEAD)
        r += 1
        for j, h in enumerate(["Key", "Label", "Unit", "Kind", "Budget column", "Day formula / definition"], start=1):
            c = ws.cell(r, j, h); c.font = font(bold=True, color="FFFFFF"); c.fill = fill(C_HEAD)
        for k in KPIS:
            r += 1
            definition = k.get("day") or f"({k.get('num')}) / ({k.get('den')})"
            for j, v in enumerate([k["key"], k["label"], k["unit"], k["kind"], k.get("budget", "derived" if self.has_budget(k) else ""), definition], start=1):
                ws.cell(r, j, v).font = font(size=9)
        self.set_widths(ws, {"A": 18, "B": 16, "C": 24, "D": 30, "E": 11, "F": 7, "G": 8, "H": 18, "I": 60, "J": 60})
        ws.freeze_panes = "A5"
        ws.protection.sheet = True

    # README -----------------------------------------------------------------------
    def sheet_readme(self, ws):
        site = self.s["site"]
        ws.sheet_view.showGridLines = False
        ws.column_dimensions["A"].width = 3; ws.column_dimensions["B"].width = 120
        lines = [
            (f"{site['name']} Production System, {self.year}", "h1"),
            (f"{site['company']}. Version {self.s['version']}. Generated workbook: do not restructure by hand; change the schema and regenerate.", "sub"),
            ("", ""),
            ("How the workbook works", "h2"),
            ("1. Inputs are the green tabs: Plant, Mining_Movements, Mining_Daily, Fleet, Safety, Gold, Commentary, Budget. Each is an Excel table. Type in the blue-headed columns only; grey-headed columns are formulas that fill automatically.", ""),
            ("2. Calc_Daily rebuilds every KPI per day from the tables (day, month to date, year to date, budget). Daily_Report, Monthly_Summary and Dashboard read only from Calc_Daily.", ""),
            ("3. There are no links to other workbooks and no macros. Copying the file anywhere keeps it working.", ""),
            ("4. Lists holds every drop-down. Add a new pit, stockpile or equipment class there; do not type free text into list columns.", ""),
            ("5. Config holds the year, the stockpile opening balances (survey at 1 January) and the prior-year safety history used for rolling 12-month rates.", ""),
            ("", ""),
            ("Daily routine", "h2"),
            ("Plant: enter the row for the day (dates are pre-filled). Mining: add one movement row per source, material and destination; add the Mining_Daily row and one Fleet row per equipment class. HSE: enter the Safety row. Everyone: one Commentary row per topic. Then open Daily_Report, check the date in C5, and print or export to PDF.", ""),
            ("", ""),
            ("Rules", "h2"),
            ("Do not insert columns inside tables, rename headers or move sheets. Add rows at the bottom of a table only. Blank means not reported; zero means reported as zero. Grades in g/t, tonnes dry metric, gold in troy ounces (31.1035 g).", ""),
            ("Ore versus waste follows the Material list: anything starting with Ore counts as ore. Ex-pit movements (Source type Pit) count to TMM and strip ratio; movements from stockpiles are rehandle.", ""),
            ("Recovery is calculated from head and tails grades (1 minus tails over head). Gold recovered is milled tonnes x (head minus tails) / 31.1035. Monthly metallurgical accounting adjustments are handled outside this workbook until the web system takes over.", ""),
            ("TRIFR and LTIFR are per million hours over the last 12 calendar months plus the current month to date, using the Config history for months before this year.", ""),
            ("", ""),
            ("Support", "h2"),
            ("Schema and generator: mps_schema.json and build_workbook.py in the MPS tools folder. Report issues to the Director, Operations and Projects.", ""),
        ]
        r = 2
        for text, kind in lines:
            c = ws.cell(r, 2, text)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            if kind == "h1":
                c.font = font(bold=True, size=16, color=C_HEAD)
            elif kind == "h2":
                c.font = font(bold=True, size=12, color=C_HEAD)
            elif kind == "sub":
                c.font = font(italic=True, color="595959")
            else:
                c.font = font()
            if text and kind == "":
                ws.row_dimensions[r].height = 15 * max(1, len(text) // 115 + 1)
            r += 1
        ws.protection.sheet = True

    # ----------------------------------------------------------------------------- sample data
    def make_sample(self) -> dict:
        rng = self.rng
        n = min(self.sample_days, len(self.days))
        days = self.days[:n]
        sample: dict = {}
        sample["prior_lti"] = date(self.year - 1, 9, 14)
        plant = []
        for i, d in enumerate(days):
            milled = round(rng.uniform(5800, 7000))
            head = round(rng.uniform(1.3, 1.9), 2)
            tails = round(rng.uniform(0.08, 0.14), 3)
            planned = 8 if i % 14 == 13 else 0
            unplanned = round(rng.choice([0, 0, 0.5, 1, 2]), 1)
            run = round(24 - planned - unplanned - rng.uniform(0, 0.5), 1)
            poured = round(milled * (head - tails) / 31.1035 * 3 * rng.uniform(0.9, 1.1)) if i % 3 == 2 else 0
            plant.append([d, milled + rng.randint(-200, 300), milled, head, tails, round(rng.uniform(30, 60)), poured, run, planned, unplanned,
                          round(rng.uniform(14, 18), 1), round(milled * rng.uniform(0.35, 0.45)), round(milled * rng.uniform(1.2, 1.6)),
                          round(milled * rng.uniform(0.8, 1.0)), round(milled * rng.uniform(28, 34)), round(milled * rng.uniform(0.9, 1.2)),
                          round(rng.uniform(3000, 3500)), None])
        sample["tblPlant"] = plant
        moves = []
        for i, d in enumerate(days):
            moves.append([d, None, "Petowal Pit", "Ore HG", "ROM Pad", round(rng.uniform(2200, 2800)), round(rng.uniform(1.6, 2.0), 2), None, None, None])
            moves.append([d, None, "Petowal Pit", "Ore HG", "Crusher Direct Tip", round(rng.uniform(800, 1200)), round(rng.uniform(1.6, 2.0), 2), None, None, None])
            moves.append([d, None, "Petowal Pit", "Ore LG", "Stockpile LG", round(rng.uniform(1200, 1800)), round(rng.uniform(0.5, 0.7), 2), None, None, None])
            moves.append([d, None, "Petowal Pit", "Waste", "Waste Dump", round(rng.uniform(16000, 20000)), None, None, None, None])
            moves.append([d, None, "Petowal Pit", "Mineralised Waste", "Mineralised Waste Dump", round(rng.uniform(600, 1000)), round(rng.uniform(0.3, 0.4), 2), None, None, None])
            moves.append([d, None, "Stockpile LG", "Ore LG", "ROM Pad", round(rng.uniform(2500, 3500)), 0.6, None, None, None])
        sample["tblMovement"] = moves
        sample["tblMiningDaily"] = [[d, round(rng.uniform(1000, 1400)), round(rng.uniform(22000, 28000)), round(rng.uniform(5000, 7000)),
                                     round(rng.uniform(30000, 40000)), round(rng.uniform(3000, 5000)), None] for d in days]
        fleet = []
        for d in days:
            fleet.append([d, "Excavator", 4, round(rng.uniform(6, 14), 1), round(rng.uniform(62, 78), 1), None])
            fleet.append([d, "Haul Truck", 20, round(rng.uniform(40, 80), 1), round(rng.uniform(300, 360), 1), None])
            fleet.append([d, "Drill", 3, round(rng.uniform(4, 10), 1), round(rng.uniform(40, 55), 1), None])
        sample["tblFleet"] = fleet
        safety = []
        for i, d in enumerate(days):
            safety.append([d, round(rng.uniform(3300, 3700)), round(rng.uniform(4300, 4700)), 0, 1 if i == 19 else 0, 0, 1 if i == 32 else 0,
                           rng.choice([0, 0, 0, 1]), rng.choice([0, 1, 2, 3]), rng.randint(8, 20), 1 if i == 40 else 0, 0, 1 if i == 25 else 0,
                           rng.choice([0, 0, 0, 1]), rng.randint(20, 40), rng.randint(2, 6), None])
        sample["tblSafety"] = safety
        sample["tblSafetyHistory"] = [[date(self.year - 1, m, 1), 240000 + m * 1000, 1 if m in (3, 9) else 0, 1 if m == 9 else 0] for m in range(1, 13)]
        gold = []
        for i, d in enumerate(days):
            if i % 10 == 9:
                gold.append([d, "Shipment", round(rng.uniform(55, 70), 1), round(rng.uniform(1100, 1400)), round(rng.uniform(300, 500)), f"SHP-{self.year}-{i // 10 + 1:03d}", None])
            if i % 10 == 4 and i > 10:
                gold.append([d, "Sale", None, round(rng.uniform(1100, 1400)), round(rng.uniform(300, 500)), f"SAL-{self.year}-{i // 10:03d}", None])
        sample["tblGold"] = gold
        sample["tblBudget"] = [[date(self.year, m, 1), 130000, 1.6, 600000, 195000, 1.55, 0.93, 9000, 250000, 36000, 1100000] for m in range(1, 13)]
        areas = self.s["lists"]["Area"]
        comments = []
        texts = ["Mill feed constrained by crusher liner change.", "Blast delayed to afternoon, fumes clearance.", "Two hazards closed out in the workshop.",
                 "TSF lift progressing, 40 percent complete.", "Community meeting held at Tomboronkoto.", "Excavator EX02 back in service after final drive replacement.",
                 "Grade control reconciliation within 3 percent for the week.", "Power interruption 40 minutes, no process upset.", "Cyanide delivery received, stock at 21 days."]
        for i, d in enumerate(days):
            for j in range(rng.randint(2, 3)):
                comments.append([d, areas[(i + j * 3) % len(areas)], texts[(i + j) % len(texts)], "DEMO"])
        sample["tblCommentary"] = comments
        sample["tblStockpiles"] = [["ROM Pad", 20000, 1.2, "Y", date(self.year, 1, 1)], ["Stockpile HG", 5000, 2.1, "N", date(self.year, 1, 1)],
                                   ["Stockpile MG", 30000, 1.0, "N", date(self.year, 1, 1)], ["Stockpile LG", 150000, 0.55, "N", date(self.year, 1, 1)],
                                   ["Mineralised Waste Dump", 200000, 0.3, "N", date(self.year, 1, 1)]]
        return sample


def build(schema_path: Path, out: Path, year: int, sample_days: int = 0, report_date: date | None = None) -> Builder:
    schema = json.loads(Path(schema_path).read_text(encoding="utf-8"))
    b = Builder(schema, year, sample_days, report_date=report_date)
    wb = b.build()
    if sample_days:
        wb.properties.title += " DEMO DATA"
        ref = wb.defined_names["cfg_ReportTitle"].attr_text  # e.g. Config!$B$13
        wb["Config"][ref.split("!")[1].replace("$", "")] = f"{schema['site']['name']} daily production report  [DEMO DATA]"
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return b


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--schema", default=str(DEFAULT_SCHEMA))
    ap.add_argument("--out", required=True)
    ap.add_argument("--year", type=int, default=2026)
    ap.add_argument("--sample-days", type=int, default=0, help="fill N days of synthetic DEMO data")
    ap.add_argument("--report-date", default=None, help="fix the Daily_Report date (YYYY-MM-DD) instead of the TODAY()-1 default; used by tests")
    a = ap.parse_args(argv)
    rd = date.fromisoformat(a.report_date) if a.report_date else None
    b = build(Path(a.schema), Path(a.out), a.year, a.sample_days, rd)
    print(f"wrote {a.out}: {len(b.wb.sheetnames)} sheets, {len(b.tables)} tables, {len(b.wb.defined_names)} names, ~{b.formula_count:,} formulas")


if __name__ == "__main__":
    main()
