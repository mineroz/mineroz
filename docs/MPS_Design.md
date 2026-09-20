# Mako Production System: design note

Petowal Mining Company S.A., Mako Gold Mine. Schema 0.3.0. Companions: `docs/data_dictionary.md` (generated) and `mps/README.md`.

## 1. Purpose and scope

The MPS replaces the three 2026 workbooks PMC_Operation_Data, PMC_Operation_Report and PMC_Operation_Commentary. The inspection found external links in formula chains, macros nobody maintains, dead file links, and figures that depend on which copy was opened last.

Two phases. The Excel phase, delivered now, is one generated workbook per year that the site team runs without IT support. The web phase moves the same data model onto company servers with a database, browser forms and the same reports. One schema file generates both, so the web system inherits the definitions agreed in Excel.

Scope: daily and monthly production reporting against budget (safety, mining, processing, gold, stockpiles, commentary). Cost, maintenance work orders, metallurgical accounting adjustments and resource modelling are excluded for now; section 9 lists what was deferred and why.

## 2. Design principles

- Single system of record. A figure is typed once, in one table, by the department that owns it. Reports hold no typed numbers.
- Flat tables. Every input is an Excel table with one row per event (movement, shipment, comment) or per day; no matrices with dates across the top.
- One calculation layer. Calc_Daily derives every KPI per day; report, monthly summary and dashboard read only Calc_Daily.
- No external links, no VBA, no hidden sheets. The file works wherever it is copied, and `tools/xl_inspect.py` can prove it.
- Schema-driven generation. `mps_schema.json` declares tables, fields, units, ranges, lists, calculated columns, conditional highlights and natural keys; `build_workbook.py` renders it and refuses to build when a formula refers to a column or list the schema does not declare. Nobody edits workbook structure by hand.
- The same schema drives the SQL. Each calculated field carries an Excel formula and an SQL expression; `build_ddl.py` emits the PostgreSQL and SQLite DDL and the data dictionary.
- One owner per input table, named on the sheet and in the schema.
- Blank means not reported; zero means reported as zero. A ratio pairs its numerator with the tonnes or hours that actually carry the measurement, so a missing assay or a missing run-hours entry leaves the period figure unchanged instead of pulling it towards zero.
- Problems are shown, not hidden. Missing grades, unknown sources, duplicate dates, a second plant-feed stockpile and a year mismatch turn red where they occur and are counted on the report.

## 3. Workbook architecture

### Sheet map

17 sheets, 20 tables, 332 defined names and about 162,000 formulas for 2026.

| Sheet | Role | Owner | Protected |
|---|---|---|---|
| README | Rules and daily routine, input-tab list built from the schema | generated | yes |
| Daily_Report | Two-page report for a chosen date | production engineer | yes, date cell unlocked, rows resizable |
| Dashboard | Five charts, year to date | read only | yes |
| Monthly_Summary | 22 KPIs by month, budget, variance | read only | yes |
| Plant, Mining_Movements, Mining_Daily, Fleet, Safety, Gold, Budget, Commentary | One input table each | owner per table, section 4 | no |
| Config | Typed settings and constants, calculated settings and checks, stockpile opening balances, 24-month safety history | production engineer | no |
| Lists | Drop-down values, Sources and Destinations | production engineer | no |
| Calc_Daily | One row per day; every KPI as day, MTD, YTD and budget | generated | yes |
| Chart_Data | Chart series from Calc_Daily | generated | yes |
| Data_Dictionary | Fields and KPI catalogue from the schema | generated | yes |

### Data flow

Input tables, then the calculated columns inside each table (Feed_oz, Is_Ore, Source_Type, Available_h, Use), then Calc_Daily (SUMIFS per date), then Daily_Report, Monthly_Summary and Chart_Data, then the Dashboard charts. Lists and Config feed every stage; nothing flows backwards.

### Calc_Daily

Each row is a date. A sum KPI is a SUMIFS over one table filtered on Date plus criteria such as Is_Ore=1, Source_Type="Pit" or Head_Grade_gpt not blank. MTD is the previous row's MTD plus today, reset when the month changes; YTD is a running total. A ratio KPI is never averaged: its Day, MTD and YTD cells divide the matching numerator and denominator columns and return blank when the denominator is zero. Budgets come from tblBudget: Bud_Month by SUMIFS on the month start, Bud_MTD as Bud_Month times day of month over days in month, Bud_YTD as the completed months plus that. Ratio budgets are derived from budgeted numerators and denominators, so strip ratio, grades and recovery need no budget lines; a helper KPI carries a budget alias (Ore_Graded_t reads the Ore_Mined_t budget) so the derived budget survives the pairing rule. Every column is a named range `cd_<column>`.

### Config settings

Typed settings sit at the top of Config: site, company, report title, the date of the last LTI before the year, and the three constants from the schema, each a named cell that formulas read: `cfg_TroyOz` (31.1035), `cfg_HoursPerDay` (24, fleet and mill calendar hours) and `cfg_RateBasisHours` (1,000,000, the TRIFR and LTIFR basis). Below them, under "Calculated (do not type here)", sit `cfg_Year` (read from the first Calc_Daily date, so the year is fixed by the generator and changed with `--year`), `cfg_YearStart`, `cfg_YearEnd`, `cfg_LastDataDate` (the last date with milled tonnes), `cfg_FeedCheck` and `cfg_YearCheck`. The two checks turn red when they are not OK and are repeated on the report. Table calcs and the KPI catalogue refer to the constants as `{HOURS_PER_DAY}` and `{RATE_BASIS_HOURS}` placeholders, and the SQL generator substitutes the same values.

### Report date and the checks line

The report date in C5 defaults to the earlier of yesterday and `cfg_LastDataDate`, clamped to the year; typing a date overrides it, a stop validation refuses dates outside the year, and the input message quotes the default formula so it can be restored. `rpt_Row` is the MATCH of that date in `cd_Date`, and every report cell is `INDEX(cd_<kpi>, rpt_Row)`. F5 shows the year check, or an instruction when the date falls outside the year, and is otherwise blank so nothing but a warning is printed. B6 states the data date, ISO week and day of month. B7, named `rpt_Checks`, is the completeness and integrity line: the rows entered for the report date in Plant, Movements, Fleet, Mining daily, Safety and Comments (a pasted duplicate reads "Plant 2", a missing day reads 0), the movement rows whose source or destination is unknown, and the movement rows dated outside the year. The line sits inside the print area.

### Daily_Report layout

Four KPI bands (SAFETY, MINING, PROCESSING, GOLD) with Day, MTD, MTD budget, variance, YTD, YTD budget and variance. Rate KPIs print their basis in the unit column as "per 1,000,000 h", read from `cfg_RateBasisHours`. MINING carries two check rows that should read zero: "Ore mined without a grade (check)" and "Movements with unknown source type (check)". Then the stockpile block, fifteen commentary rows (36 pt each, resizable on the protected sheet) and a footer with the print time. The sheet prints on two portrait A4 pages, one page wide, with a manual page break before PROCESSING: safety and mining on the first page, processing, gold, stockpiles and commentary on the second.

### Stockpiles

The block has one row per row of tblStockpiles on Config, including the five spare rows declared in the schema (`spare_rows`), so a stockpile opened during the year appears on the report as soon as it is named, without a rebuild. Each balance runs from the later of 1 January and the stockpile's Survey_Date to the report date: Opening_t is the surveyed tonnage at the start of Survey_Date (blank means 1 January), and after a re-survey the user enters the new tonnes, grade and date. Balance is opening plus deliveries minus reclaim; for the stockpile flagged Plant_Feed = Y the plant draw is also deducted as milled tonnes minus direct tip (movements whose destination type is Plant) over the same window. Grade follows the same balance in ounces. Exactly one stockpile may carry the flag: `cfg_FeedCheck` reads OK when the count is 1, otherwise it shows the rule in red on Config and in the line above the STOCKPILES header, because the deduction is wrong until it is fixed. The feed stockpile balance therefore covers the ROM pad plus the crusher surge stock; a separate crushed-ore balance is deferred (section 9).

### Monthly_Summary and the pro-rata budget

Twenty-two KPIs by month with Actual, Budget and Variance rows where a budget exists. A completed month shows the MTD value at month end against the full month budget. The month in progress shows the MTD actual at `cfg_LastDataDate` against `Bud_<key>_MTD` at the same date, that is the budget pro-rata by day; a future month shows the full month budget and no actual. The Year column uses the YTD actual and `Bud_<key>_YTD` at the last data date, so year to date is compared pro-rata as well. B3 states this rule. Ratios are month values, not averages of days.

### Dashboard

Chart_Data returns NA() beyond `cfg_LastDataDate` and wherever the Calc_Daily source is not numeric, so actual lines stop at the last data point instead of falling to zero; budget lines run the full year. Charts: gold poured YTD against budget, milled per day against budget, ore and waste per day, MTD recovery and head grade, rolling TRIFR. All charts use a date axis with month labels and print on one landscape A4 page.

### Protection, Protected View and one writer at a time

Report and calculation sheets are protected without a password, to stop accidental typing rather than to secure the file; Daily_Report leaves row formatting unlocked so a long comment can be given more height. Input, Config and Lists sheets stay open because their tables must grow. Nothing is hidden. A file received by e-mail or download opens behind Excel's yellow Protected View bar with a blank report until the user clicks Enable Editing; the README sheet says so. The workbook is a single file, so one person edits it at a time: open, enter, save, close. If Excel opens it Read-Only someone else has it; the rule is to wait, never to Save As a copy. The README suggests an entry order (Mining 06:30, Plant 06:45, HSE 07:00, Commentary 07:15) and notes that simultaneous entry needs the file on SharePoint or OneDrive.

## 4. Data model summary

| Table | Grain | Key | Owner | Purpose |
|---|---|---|---|---|
| tblPlant | day | Date, prefilled | Processing | Crusher, mill, grades, gold poured, downtime, reagents, power, water, GIC |
| tblMovement | source, material, destination per day; optional shift | unique on those five | Mining dispatch | Movements ledger: everything mined, rehandled or tipped |
| tblMiningDaily | day | Date, prefilled | Mining | Drilling, grade control drilling, blasting, explosives, diesel, dewatering |
| tblFleet | equipment class per day | unique on Date, Equipment_Class | Mining, maintenance planner | Units, down hours, operating hours (SMU engine hours) |
| tblSafety | day | Date, prefilled | HSE | Hours worked, injuries by class, incidents, leading indicators |
| tblSafetyHistory | month; 24 prefilled (prior year and reporting year) | Month | HSE | Hours, recordables, LTI for rolling rates; Use flags months without daily hours |
| tblGold | shipment or sale | none | Finance, gold room | Dore weight, gold and silver ounces, reference |
| tblBudget | month | Month | Technical services | Budget or latest forecast |
| tblCommentary | comment | none | all departments | Narrative by area and date |
| tblSources, tblDestinations | reference | name | production engineer | Type each movement end: Pit or Stockpile; Stockpile, Plant or Waste. Type is required once the name is filled |
| tblStockpiles | reference, five spare rows | name | production engineer | Opening survey tonnes, grade and date, plant feed flag |

### Schema keys

Besides `name`, `type`, `unit`, `min`, `max`, `required`, `key`, `list`, `desc`, `calc` and `sql`, a field may carry `required_when` (an Excel condition; the cell turns red when the condition holds and the cell is blank, and Data_Dictionary shows "when {Is_Ore}=1") with a `required_when_sql` twin, `warn_when` (a red highlight whenever the condition holds, used for Source_Type and Dest_Type = "?"), `values_from` (the derived list a calculated text column draws on) and `sql_postgres` and `sql_sqlite` where one SQL expression cannot serve both dialects (tblBudget.Days). A table may declare `unique`, a list of column groups forming its natural key (tblMovement, tblFleet), which the DDL reads for its constraints. A reference table may declare `spare_rows`. The `constants` block holds the troy ounce, hours per day and rate basis.

### Validation and highlights

Drop-downs come from the dynamic `lst_<list>` names, which count from the first data row so a value added at the bottom of a list is offered at once. Range checks are warnings, not blocks, because outliers happen. Dates in input tables are validated between `cfg_YearStart` and `cfg_YearEnd`, since a row outside the year has no Calc_Daily row and would vanish from every KPI; the safety history accepts twelve months earlier. The key column of every prefilled table (dates, months, prior months) carries a stop validation that refuses a duplicate date, plus a red highlight for a duplicate that arrives by paste, because a duplicated day would be double counted by SUMIFS.

### KPI catalogue

Mining. Ore mined is ex-pit only: rows whose Source is typed Pit and whose Material starts with Ore (the ore flag). Waste is ex-pit rows without the flag, which today includes mineralised waste and topsoil. TMM is ore plus waste; strip ratio is waste over ore. Rehandle is any row whose Source is a stockpile and counts to nothing else. Contained ounces are tonnes times grade over 31.1035. Every ore row needs a grade (Grade_gpt is required when Is_Ore = 1): the helper KPI Ore_Graded_t sums only pit ore tonnes with a grade and is the denominator of Ore_Grade_gpt, so an ore row entered before its grade does not dilute the grade mined, and Ore_Ungraded_t (Ore_Mined_t minus Ore_Graded_t) shows those tonnes as a check line. Moved_Total_t sums every movement row regardless of type, and Unclassified_t (Moved_Total_t minus TMM minus rehandle) exposes rows whose source is not in tblSources; both check lines should read zero. Powder factor is explosives over blasted tonnes; GC_Drill_m reports grade control drill metres beside production drilling.

Fleet. Calendar hours are units times `cfg_HoursPerDay`. Availability is calendar minus down, over calendar; utilisation is operating over available. Exc_Productivity_tph and Trk_Productivity_tph divide TMM plus rehandle by excavator and by truck operating hours. Excavators and haul trucks are reported; other classes are captured only.

Processing. Feed ounces are milled tonnes times head grade, tails ounces likewise, recovered is feed minus tails, and recovery is recovered over feed. The blank-grade guard applies at row level: Feed_oz is blank without milled tonnes and a head grade, Tails_oz without a tails grade, Recovered_oz and Recovery_pct unless milled, head and tails are all present. At period level the grades and recovery are paired with the tonnes that carry the assay: Milled_Graded_t (milled tonnes with a head grade) is the Head_Grade_gpt denominator, Milled_Tails_Graded_t the Tails_Grade_gpt denominator, and Feed_Recon_oz (feed ounces on days with a tails grade) the Recovery_pct denominator, so a day whose assay is outstanding is left out rather than counted as zero grade or 100 percent recovery. Milled_Timed_t (milled tonnes on days with run hours) is the Throughput_tph numerator. Mill calendar hours are `cfg_HoursPerDay` on any day with run, planned or unplanned hours entered and zero otherwise, so unreported days do not dilute the rates; a full-day shutdown is planned maintenance 24 with run hours 0. Availability is calendar minus planned minus unplanned, over calendar; utilisation is run hours over that available time. Gravity_Share_pct is gravity gold over gold recovered. Gold poured is reported separately and not forced to reconcile daily; GIC is a point value. Reagent, power and water intensities are per milled tonne.

Safety. Recordables are fatalities plus LTI, RWI and MTI, blank while all four are blank. TRIFR and LTIFR are injuries per `cfg_RateBasisHours` over a window from twelve months before the first of the report month to the report date. Daily Safety rows supply the window; for any earlier month in the window that has no daily hours, the Use column of tblSafetyHistory is 1 and that month's hours and injuries come from the history instead, which covers the prior year and the months of the reporting year before go-live. Days since LTI is the report date minus the later of `cfg_PriorLTIDate` and the last daily row with an LTI.

## 5. Daily operating procedure and controls

Timings are proposals to be confirmed (section 9). By 07:00 mining dispatch enters the previous day's movement rows (one daily total per source, material and destination, or Day and Night rows, never both), the Mining_Daily row and one Fleet row per class; processing enters the Plant row and HSE the Safety row. By 07:30 each department adds its Commentary rows, one per topic. At 08:00 the production engineer checks the report date, reads the `rpt_Checks` line and the two MINING check rows, scans for red cells, exports Daily_Report to PDF (`Export-DailyReport.ps1`) and distributes it.

Paste only with Paste Special > Values; an ordinary paste removes the drop-downs and range checks. Add rows at the bottom of a table only; never insert columns. Month end: reconcile ore mined with grade control, milled tonnes with the metallurgical balance, stockpiles with the survey (enter the re-survey on Config) and hours worked with payroll; book adjustments as dated rows, never by editing Calc_Daily. Backups are dated file copies each evening; the file is self-contained. The schema version is written to README and Config.

## 6. Generator, validator and site scripts

`build_workbook.py --out FILE [--year 2026] [--sample-days N] [--report-date YYYY-MM-DD]` renders the workbook; `--sample-days` fills deterministic synthetic data and marks the title [DEMO DATA]. At the end of every build `check_catalogue()` verifies each table and column reference, list criterion, budget column and ratio reference in the KPI catalogue, the report, Config, Monthly_Summary, Lists, Calc_Daily, Chart_Data and every table calc against the schema, and raises an error naming the offending KPI or cell.

`validate_workbook.py [--year 2026] [--keep] [--workdir DIR] [--soffice PATH]` builds four workbooks (DEMO with the report date fixed to 10 February, DEMO with the default date, blank, and an edge case with missing grades, missing run hours, a full-day shutdown, an ungraded ore row, an unknown source, a prior-year row, a re-surveyed stockpile and a Probe sheet that measures every drop-down name), recalculates them headless in LibreOffice, recomputes the key figures in plain Python from the synthetic rows and compares them with Calc_Daily, the report, the monthly summary and the charts, checks the completeness line, the feed check, the validations and the highlights, scans every sheet for error values and runs `xl_inspect.py`. 931 checks in about a minute; exit code 1 on any failure. `--year 2028` exercises the leap-year branch.

On the site PC, `Setup-MakoPS.ps1` (generated by `tools/build_setup.py`; switches `-Root`, `-Source`, `-Year`, `-SkipCopy`, `-SkipRun`, `-SkipBuild`) creates `C:\MakoPS` with `01_Tools`, `02_Source\2026`, `03_Inspection` and `04_MPS`, resolves Python once and passes `-NoInstall` to its children. `PythonEnv.ps1` finds a real Python (the Store stub is rejected by its exit code), installs Python 3.12 per user through the Windows proxy when none exists, and installs missing modules with pip. `Build-MPS.ps1` builds the blank and DEMO workbooks into `04_MPS`, taking the generator from the repository when run from it and otherwise from `01_Tools\mps`. `Export-DailyReport.ps1` opens the workbook read-only through Excel COM with an en-US culture, writes the date into C5, exports the sheet to `04_MPS\Reports\Mako_Daily_Report_yyyy-MM-dd.pdf` and can display an Outlook mail with the PDF attached.

## 7. Migration from the legacy files

1. From the inspection, mark which legacy sheets hold source data rather than derived figures.
2. Load the movements ledger by unpivoting the legacy daily mining sheets to one row per date, source, material and destination with tonnes and grade. Where only ore and waste totals exist, load them against a placeholder pit and say so in Comments.
3. Load plant daily rows; recompute recovery from head and tails and compare with the legacy figure.
4. Fill the safety history (hours, recordables, LTI) for the twelve months before the reporting year and for any month of the reporting year before go-live, and enter the date of the last LTI.
5. Enter the opening stockpile survey, grades and survey dates, and flag the plant feed stockpile.
6. Parallel run for one full month, both systems producing the daily report; differences are resolved in the schema or the data, never by adjusting outputs.
7. Cut-over criteria: month-end totals agree within tolerance (tonnes 0.5 percent, ounces 1 percent, hours exact), every owner has entered their table unaided for ten consecutive days, the check rows and the checks line read zero, and the inspector reports no errors on the live file.

## 8. Web phase

Stack: Python with FastAPI or Django, PostgreSQL, server-rendered HTML with minimal JavaScript. It runs on Windows Server or Linux, needs no client install, and the generators are already Python. Authentication starts with local accounts and moves to Active Directory (LDAP or Kerberos) later.

`build_ddl.py` already emits the DDL from the schema JSON: tables with CHECK constraints from min and max, natural keys from `unique`, foreign keys to the list tables, generated columns for calculated fields (COALESCE where Excel treats blank as zero, NULL where Excel shows blank), and `kpi_daily` and `kpi_monthly` views that reproduce the whole KPI catalogue including the helper KPIs. The same file drives form fields, labels and validation. During transition a nightly job imports the current workbook table by table, so Excel remains the entry point until each department switches.

API surface: `/api/<table>` (list, get, upsert on the natural key, delete), `/api/kpi/daily` and `/api/kpi/monthly`, `/api/report/daily/{date}` as HTML or PDF, `/api/lists`, `/api/health`. Roles per table mirror the owner column.

Roll-out: database and nightly import; read-only web reports; entry forms per department, Commentary and Safety first, movements and plant last; retire the workbook; Active Directory.

## 9. Backlog for the web phase and open design decisions

Items deferred by the review, each with the reason and the proposed treatment, merged with the questions still open with the site.

- Movement grain (F014). Blank Shift means a daily total; Day and Night rows are the alternative. The Excel phase keeps this as a procedure (never both for the same source, material and destination). The web phase should add a declared Full Day shift value, enforce the `unique` key and add a mixed-grain check. Open with the site: which grain dispatch will use.
- Stockpile survey adjustments (F015). An Adjustment source or destination type, excluded from ore mined, TMM and rehandle, belongs with the reconciliation adjustments planned for the web phase; until then a survey goes into the Config opening balance with its Survey_Date and a Comment.
- Crushed ore stockpile (F016) and the plant feed balance (F043). With no crushed-ore stockpile in this phase the feed stockpile balance deliberately covers ROM pad plus crusher surge stock (milled less direct tip, cumulative), which conserves tonnes where a daily MAX(0, milled minus tip) would debit the pad twice. A separate crushed balance needs a new opening balance and a redesigned ROM reconciliation, proposed once Processing confirms how the crushed stockpile is surveyed. Confirm with the site that a single stockpile feeds the plant and that mineralised waste and topsoil count as waste in the strip ratio.
- Metallurgical accounting (F017). Monthly adjustments are outside this phase; propose a tblMetAdjust keyed on month with a met balance KPI. Open: daily head and tails, or head back-calculated monthly from poured gold and GIC?
- Mill standby hours (F018). A Mill_Standby_h field changes the plant time model and the meaning of availability; agree the definitions with Processing, then add the field and an unaccounted-hours check in the schema.
- Gold room (F019). Dore on hand, outturn adjustment, shipment link and price are outside this phase. Open: booking point (pour, dispatch or refinery outturn) and how adjustments are recorded.
- Budget and forecast scenarios (F020). Add a Scenario key to tblBudget and a cfg_PlanScenario setting in the web phase. Open: does tblBudget hold the approved budget, the latest forecast, or both?
- Week-to-date and quarter-to-date (F021). Add WTD and QTD running totals in Calc_Daily when the web phase defines its period model.
- French labels and code versus label (F022). Treat in the web phase with a Materials reference table replacing the Ore prefix rule, together with French for workforce-facing screens and lists.
- TSF, water balance, rainfall and WAD cyanide (F023). A new subject area; add to tblPlant or tblSafety once Environment confirms the daily figures (freeboard, decant volume, return water, rainfall).
- Headcount and persons on board (F024). Add point fields to tblSafety when HSE confirms the categories.
- Fleet down hours (F048). Splitting Down_h into planned and unplanned changes the fleet time model; agree the definitions with Mining and Maintenance, then add the two fields with Down_h as a calculated sum. Related open question: contractor versus owner fleet, one time model or separate classes?
- Occupational health (F050). Occupational illness and malaria cases are a new subject area; add integer fields to tblSafety with sum KPIs once HSE confirms the case definitions.
- Still open with the site, not tied to a backlog item: pit and stage names for the Sources list and whether stages are tracked; the grade source for movements (grade control model, blast-hole assays or a blend); the cut-off time per table and the time the PDF must be out; the PDF distribution list, including the owner in Perth; and cost reporting (unit costs per tonne and per ounce, and from which system).
