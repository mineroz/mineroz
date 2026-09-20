# mps

Generator, validator and SQL builder for the Mako Production System (MPS) workbook. Everything about the data model lives in `schema/mps_schema.json`; the scripts render it. See `docs/MPS_Design.md` for the design and `docs/data_dictionary.md` for the field-level dictionary. On the site PC the same files sit in `C:\MakoPS\01_Tools\mps` and are driven by the PowerShell scripts described in `tools/README.md`.

## Requirements

- Python 3.11 or later with `openpyxl` 3.1 (`pip install openpyxl`).
- For the validator only: LibreOffice 24 Calc (headless recalculation) and `oletools` (`pip install oletools`) for the VBA check in `tools/xl_inspect.py`.

Excel 2016 or later opens the workbook. Functions that Excel introduced after 2010 are written with the `_xlfn.` prefix (MAXIFS, ISOWEEKNUM) so both Excel and LibreOffice evaluate them; XLOOKUP is not used.

## Build the workbook

From the repository root:

```bash
# blank workbook for the site, dates prefilled for the year, no data
python3 mps/build_workbook.py --out dist/MPS_2026.xlsx

# DEMO workbook: 74 days of deterministic synthetic data, title marked [DEMO DATA]
python3 mps/build_workbook.py --out dist/MPS_2026_DEMO.xlsx --sample-days 74

# another year
python3 mps/build_workbook.py --out dist/MPS_2027.xlsx --year 2027

# fix the report date instead of the default (yesterday or last data date); used by tests
python3 mps/build_workbook.py --out dist/demo.xlsx --sample-days 74 --report-date 2026-02-10
```

The last line of the output states the sheet, table, name and formula counts (17 sheets, 20 tables, 282 names, about 146,000 formulas for 2026). The file is saved with full recalculation on load, so Excel computes everything the first time it is opened; save it once from Excel before distributing so cached values exist for readers that do not recalculate.

On site, `C:\MakoPS\01_Tools\Build-MPS.ps1` runs these two builds (blank and DEMO) into `C:\MakoPS\04_MPS`, finding or installing Python first. `dist/` is for local builds and is not committed.

## Run the validator

```bash
python3 mps/validate_workbook.py            # fresh temp folder, deleted afterwards
python3 mps/validate_workbook.py --keep --workdir /tmp/mps_check
python3 mps/validate_workbook.py --soffice "C:\Program Files\LibreOffice\program\soffice.exe"
```

It builds three workbooks (DEMO with the report date fixed to 10 February, DEMO with the default report date, and blank), recalculates them with LibreOffice, recomputes the key figures in plain Python from the synthetic rows without using any workbook formula, and compares them with Calc_Daily on 1 January, 10 February and the last day with data. It then checks Daily_Report, Monthly_Summary and Chart_Data against Calc_Daily, scans every sheet for error values, runs `tools/xl_inspect.py` on the built file (no external links, no VBA, no dangling references) and confirms the blank workbook recalculates without a single error. One line per check, about 270 checks; exit code 1 on any failure. Allow about a minute.

Run it after every change to the schema or the generator.

## Change the schema and regenerate

Edit `schema/mps_schema.json`, then rebuild, validate, regenerate the DDL and repack the bootstrap that carries the schema and generator to site:

```bash
python3 mps/build_workbook.py --out dist/MPS_2026.xlsx
python3 mps/validate_workbook.py
python3 mps/build_ddl.py
python3 tools/build_setup.py
```

What the schema holds:

| Key | Meaning |
|---|---|
| `lists` | Drop-down values. Each becomes a table on the Lists sheet, a dynamic named range `lst_<name>` and a `ref_<name>` table in SQL. |
| `reference_tables` | Sources, Destinations (sheet Lists) and Stockpiles (sheet Config) with seed rows. |
| `derived_lists` | Values used by reference table columns (Source_Type, Destination_Type). |
| `tables` | The nine input tables. `table` is the Excel table name, `sheet` the tab, `owner` the department, `prefill` one of `dates`, `months`, `prior_months`; otherwise `rows` gives the capacity. |
| `tables[].fields` | `name`, `label`, `type` (`date`, `text`, `list`, `number`, `int`, `pct`), `unit`, `min`, `max`, `required`, `key`, `list`, `desc`, `calc`, `sql`. |

A field with `calc` becomes a grey calculated column. Inside `calc`, `{Col}` is a this-row reference to another column of the same table and `{TROY}` is the troy ounce constant. Give the same field an `sql` expression (column names in lower case) so the DDL can generate it; a `calc` without `sql` stays Excel-only and is listed as such in the SQL header.

Rules of thumb:

- Adding a drop-down value does not need a rebuild: type it at the bottom of the list on the Lists sheet. Mirror it in the schema so the next build carries it.
- Adding or renaming a column, a table or a validation range needs a rebuild. The generator writes an empty workbook; move existing data across by pasting values into the matching columns of the new file, column by column if the order changed.
- Bump `version` in the schema for any change that alters a definition. It is written to the README sheet, the SQL headers and the dictionary.
- Never edit the built workbook's structure by hand; the next build would silently drop the change.
- `Setup-MakoPS.ps1` embeds the schema, the three scripts and this README, so commit the regenerated bootstrap together with the change (`python3 tools/build_setup.py --check` verifies it).

KPIs live in the `KPIS` list at the top of `build_workbook.py`, one `K(...)` entry each: `sum` (a SUMIFS over one table, with MTD and YTD), `ratio` (numerator and denominator built from other KPIs, recomputed per period), `custom` (an explicit Day formula, used for rolling rates and days since LTI) and `point` (a value at the date). `budget` names a tblBudget column; ratio budgets are derived when both sides are budgeted. `REPORT_LAYOUT` and `MONTHLY_KPIS` decide what appears on Daily_Report and Monthly_Summary. `build_ddl.py` translates plain SUMIFS and arithmetic automatically; anything else needs an entry in its `KPI_SQL` map.

## Build the DDL

```bash
python3 mps/build_ddl.py
```

No arguments; output is byte-identical on every run. It writes:

| File | Content |
|---|---|
| `mps/sql/schema_postgres.sql` | PostgreSQL 13 or later. Tables, CHECK constraints from min and max, foreign keys to the list tables, generated columns for calculated fields, `<table>_calc` views for lookups, `kpi_daily` and `kpi_monthly` views, seeds, `updated_at` triggers. |
| `mps/sql/schema_sqlite.sql` | The same for SQLite 3.31 or later, for local testing. |
| `docs/data_dictionary.md` | Field-level dictionary and the Excel to SQL mapping. |

Apply with `psql -v ON_ERROR_STOP=1 -1 -f mps/sql/schema_postgres.sql` (single transaction) or `sqlite3 mps.db < mps/sql/schema_sqlite.sql`. Both scripts are safe to re-run on an existing database; structural changes to populated tables still need a migration.

## Folder layout

| Path | Purpose |
|---|---|
| `mps/schema/mps_schema.json` | The data model. Single source for the workbook, the SQL and the dictionary. |
| `mps/build_workbook.py` | Workbook generator and KPI catalogue. |
| `mps/validate_workbook.py` | Build, recalculate, verify. |
| `mps/build_ddl.py` | SQL and dictionary generator. |
| `mps/sql/` | Generated DDL, committed so reviewers can read it without running anything. |
| `docs/MPS_Design.md` | Design note. |
| `docs/data_dictionary.md` | Generated dictionary. |
| `tools/` | Workbook inspector, the site PowerShell scripts and the bootstrap builder. See `tools/README.md`. |
| `Setup-MakoPS.ps1` | Generated bootstrap; carries the `mps` files to the site PC. |
| `dist/` | Local builds, not committed. |

## C:\MakoPS layout on site

`Setup-MakoPS.ps1` creates and fills the layout, then runs the inspection and builds the workbooks.

| Folder | Content |
|---|---|
| `01_Tools` | `PythonEnv.ps1`, `Invoke-XlInspect.ps1`, `Build-MPS.ps1`, `Export-DailyReport.ps1`, `xl_inspect.py`, `README.md`. |
| `01_Tools\mps` | `build_workbook.py`, `build_ddl.py`, `validate_workbook.py`, `schema\mps_schema.json` and this README. Rebuilt from the repository by the bootstrap; do not edit here. |
| `02_Source\2026` | Working copies of the legacy PMC workbooks. Originals on X: are not touched. |
| `03_Inspection` | `inspection_report.md` and one JSON per legacy workbook. |
| `04_MPS` | `MPS_2026.xlsx`, the live workbook and the only file anyone edits, and `MPS_2026_DEMO.xlsx` for training. `Export-DailyReport.ps1` writes the PDFs to `04_MPS\Reports\Mako_Daily_Report_yyyy-MM-dd.pdf`. Keep dated backup copies in `04_MPS\Backup`. |

Re-running `Build-MPS.ps1` overwrites both workbooks, so copy the live file to `Backup` first and paste its table bodies into the new build. Re-running the inspector on `04_MPS` (`python xl_inspect.py C:\MakoPS\04_MPS -o C:\MakoPS\03_Inspection\mps`) is the quickest proof that the live workbook still has no external links, no VBA and no error values.
