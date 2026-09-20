# Mako Production System: design note

Petowal Mining Company S.A., Mako Gold Mine. Schema 0.1.0. Companions: `docs/data_dictionary.md` (generated) and `mps/README.md`.

## 1. Purpose and scope

The MPS replaces the three 2026 workbooks PMC_Operation_Data, PMC_Operation_Report and PMC_Operation_Commentary. The inspection found external links in formula chains, macros nobody maintains, dead file links, and figures that depend on which copy was opened last.

Two phases. The Excel phase, delivered now, is one generated workbook per year that the site team runs without IT support. The web phase moves the same data model onto company servers with a database, browser forms and the same reports. One schema file generates both, so the web system inherits the definitions agreed in Excel.

Scope: daily and monthly production reporting against budget (safety, mining, processing, gold, stockpiles, commentary). Cost, maintenance work orders and resource modelling are excluded for now.

## 2. Design principles

- Single system of record. A figure is typed once, in one table, by the department that owns it. Reports hold no typed numbers.
- Flat tables. Every input is an Excel table with one row per event (movement, shipment, comment) or per day; no matrices with dates across the top.
- One calculation layer. Calc_Daily derives every KPI per day; report, monthly summary and dashboard read only Calc_Daily.
- No external links, no VBA, no hidden sheets. The file works wherever it is copied, and `tools/xl_inspect.py` can prove it.
- Schema-driven generation. `mps_schema.json` declares tables, fields, units, ranges, lists and calculated columns; `build_workbook.py` renders it. Nobody edits workbook structure by hand.
- The same schema drives the SQL. Each calculated field carries an Excel formula and an SQL expression; `build_ddl.py` emits the PostgreSQL and SQLite DDL and the data dictionary.
- One owner per input table, named on the sheet and in the schema.
- Blank means not reported; zero means reported as zero, so ratios skip blank days.

## 3. Workbook architecture

### Sheet map

| Sheet | Role | Owner | Protected |
|---|---|---|---|
| README | Rules and daily routine | generated | yes |
| Daily_Report | One-page report for a chosen date | production engineer | yes, date cell unlocked |
| Dashboard | Five charts, year to date | read only | yes |
| Monthly_Summary | 22 KPIs by month, budget, variance | read only | yes |
| Plant, Mining_Movements, Mining_Daily, Fleet, Safety, Gold, Budget, Commentary | One input table each | owner per table, section 4 | no |
| Config | cfg_* settings, opening stockpile survey, safety history | production engineer | no |
| Lists | Drop-down values, Sources and Destinations | production engineer | no |
| Calc_Daily | One row per day; every KPI as day, MTD, YTD and budget | generated | yes |
| Chart_Data | Chart series from Calc_Daily | generated | yes |
| Data_Dictionary | Fields and KPI catalogue from the schema | generated | yes |

### Data flow

Input tables, then the calculated columns inside each table (Feed_oz, Is_Ore, Source_Type, Available_h), then Calc_Daily (SUMIFS per date), then Daily_Report, Monthly_Summary and Chart_Data, then the Dashboard charts. Lists and Config feed every stage; nothing flows backwards.

### Calc_Daily

Each row is a date. A sum KPI is a SUMIFS over one table filtered on Date plus criteria such as Is_Ore=1 or Source_Type="Pit". MTD is the previous row's MTD plus today, reset when the month changes; YTD is a running total. A ratio KPI is never averaged: its Day, MTD and YTD cells divide the matching numerator and denominator columns, so MTD recovery is MTD recovered ounces over MTD feed ounces. Budgets come from tblBudget: Bud_Month by SUMIFS on the month start, Bud_MTD pro-rata by day over days in month, Bud_YTD as completed months plus that. Ratio budgets are derived from budgeted numerators and denominators, so strip ratio and recovery need no budget lines. Every column is a named range `cd_<column>`.

### Report date

`cfg_LastDataDate` is the last date with milled tonnes. The report date defaults to the earlier of yesterday and that date, clamped to the year; typing a date overrides it. `rpt_Row` is the MATCH of that date in `cd_Date`, and every report cell is `INDEX(cd_<kpi>, rpt_Row)`. Stockpile balances and commentary use the same date.

### Dashboard

Chart_Data returns NA() beyond `cfg_LastDataDate`, so actual lines stop at the last data point instead of falling to zero; budget lines run the full year. Charts: gold poured YTD against budget, milled per day against budget, ore and waste, MTD recovery and head grade, rolling TRIFR.

### Protection

Report and calculation sheets are protected without a password, to stop accidental typing rather than to secure the file. Input, Config and Lists sheets stay open because their tables must grow. Nothing is hidden.

## 4. Data model summary

| Table | Grain | Key | Owner | Purpose |
|---|---|---|---|---|
| tblPlant | day | Date, prefilled | Processing | Crusher, mill, grades, gold poured, downtime, reagents, power, water, GIC |
| tblMovement | source, material, destination per day; optional shift | unique on those five | Mining dispatch | Movements ledger: everything mined, rehandled or tipped |
| tblMiningDaily | day | Date, prefilled | Mining | Drilling, blasting, explosives, diesel, dewatering |
| tblFleet | equipment class per day | Date, Equipment_Class | Mining, maintenance planner | Units, down hours, operating hours |
| tblSafety | day | Date, prefilled | HSE | Hours worked, injuries by class, incidents, leading indicators |
| tblSafetyHistory | prior-year month | Month | HSE | Hours, recordables, LTI for rolling rates |
| tblGold | shipment or sale | none | Finance, gold room | Dore weight, gold and silver ounces, reference |
| tblBudget | month | Month | Technical services | Budget or latest forecast |
| tblCommentary | comment | none | all departments | Narrative by area and date |
| tblSources, tblDestinations | reference | name | production engineer | Type each movement end: Pit or Stockpile; Stockpile, Plant or Waste |
| tblStockpiles | reference | name | production engineer | Opening survey tonnes and grade, plant feed flag |

### KPI catalogue

Mining. Ore mined is ex-pit only: rows whose Source is typed Pit and whose Material starts with Ore (the ore flag). Waste is ex-pit rows without the flag, which today includes mineralised waste and topsoil. TMM is ore plus waste; strip ratio is waste over ore. Rehandle is any row whose Source is a stockpile and counts to nothing else. Contained ounces are tonnes times grade over 31.1035; grade mined is the tonnage-weighted mean. Powder factor is explosives over blasted tonnes.

Fleet. Calendar hours are units times 24. Availability is calendar minus down, over calendar; utilisation is operating over available. Excavators and haul trucks are reported; other classes are captured only.

Processing. Feed ounces are milled tonnes times head grade, tails ounces likewise. Recovered is feed minus tails, so recovery is one minus tails over head, tonnage-weighted for MTD. Gold poured is reported separately and not forced to reconcile daily; GIC is a point value. Mill calendar hours are 24 per day that has a run-hours entry, so unreported days do not dilute the rate; availability is calendar minus planned minus unplanned, over calendar, and utilisation is run hours over that available time. Reagent, power and water intensities are per milled tonne.

Safety. Recordables are fatalities plus LTI, RWI and MTI. TRIFR and LTIFR are per million hours over a window from twelve months before the first of the report month to the report date; months before the reporting year come from tblSafetyHistory. Days since LTI is the report date minus the later of `cfg_PriorLTIDate` and the last daily row with an LTI.

Stockpiles. Balance to the report date is opening survey plus tonnes delivered minus tonnes reclaimed. For the stockpile flagged Plant_Feed, and there must be exactly one, the plant draw is also deducted as milled tonnes minus direct-tip tonnes. Grade follows the same balance in ounces.

## 5. Daily operating procedure and controls

Timings are proposals (question 8). By 07:00 mining dispatch enters the previous day's movement rows and one Fleet row per class, processing the Plant row and HSE the Safety row. By 07:30 each department adds its Commentary rows, one per topic. At 08:00 the production engineer checks the report date, scans Calc_Daily for blanks, exports Daily_Report to PDF (`Export-DailyReport.ps1`) and distributes it.

Controls built in: drop-downs from named lists, range warnings from the schema min and max (warnings, not blocks, because outliers happen), date validation to the reporting year, greyed calculated columns. Month end: reconcile ore mined with grade control, milled tonnes with the metallurgical balance, stockpiles with the survey and hours worked with payroll; book adjustments as dated rows, never by editing Calc_Daily. Backups are dated file copies each evening; the file is self-contained. The schema version is written to README and Config.

## 6. Migration from the legacy files

1. From the inspection, mark which legacy sheets hold source data rather than derived figures.
2. Load the movements ledger by unpivoting the legacy daily mining sheets to one row per date, source, material and destination with tonnes and grade. Where only ore and waste totals exist, load them against a placeholder pit and say so in Comments.
3. Load plant daily rows; recompute recovery from head and tails and compare with the legacy figure.
4. Load twelve months of safety history (hours, recordables, LTI) and the date of the last LTI.
5. Enter the opening stockpile survey and grades.
6. Parallel run for one full month, both systems producing the daily report; differences are resolved in the schema or the data, never by adjusting outputs.
7. Cut-over criteria: month-end totals agree within tolerance (tonnes 0.5 percent, ounces 1 percent, hours exact), every owner has entered their table unaided for ten consecutive days, and the inspector reports no errors on the live file.

## 7. Web phase

Stack: Python with FastAPI or Django, PostgreSQL, server-rendered HTML with minimal JavaScript. It runs on Windows Server or Linux, needs no client install, and the generators are already Python. Authentication starts with local accounts and moves to Active Directory (LDAP or Kerberos) later.

`build_ddl.py` already emits the DDL from the schema JSON, with calculated fields as generated columns and `kpi_daily` and `kpi_monthly` views that reproduce Calc_Daily; the same file drives form fields, labels and validation. During transition a nightly job imports the current workbook table by table, so Excel remains the entry point until each department switches.

API surface: `/api/<table>` (list, get, upsert on the natural key, delete), `/api/kpi/daily` and `/api/kpi/monthly`, `/api/report/daily/{date}` as HTML or PDF, `/api/lists`, `/api/health`. Roles per table mirror the owner column.

Roll-out: database and nightly import; read-only web reports; entry forms per department, Commentary and Safety first, movements and plant last; retire the workbook; Active Directory.

## 8. Open questions

1. Pit and stage names for the Sources list, and whether stages are tracked.
2. Contractor versus owner fleet: one time model or separate classes?
3. Shift granularity: is the movements ledger kept per shift or per day?
4. Grade source for movements: grade control model, blast-hole assays or a blend?
5. Metallurgical accounting: daily head and tails, or head back-calculated monthly from poured gold and GIC?
6. Gold room flows: booking point (pour, dispatch or refinery outturn) and how adjustments are recorded.
7. Budget versus forecast: does tblBudget hold the approved budget, the latest forecast, or both?
8. Reporting deadlines: cut-off time per table and the time the PDF must be out.
9. PDF distribution list, including the owner in Perth?
10. French for workforce-facing screens and lists in the web phase.
11. TSF and water metrics: freeboard, decant volume, return water, rainfall.
12. Cost reporting: unit costs per tonne and per ounce, and from which system.
13. Workbook assumptions to confirm: mineralised waste and topsoil count as waste in the strip ratio, and a single stockpile feeds the plant.
