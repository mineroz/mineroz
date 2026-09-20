# Mako Production System: design note

Petowal Mining Company S.A., Mako Gold Mine. Schema 0.3.0. Companions: `docs/data_dictionary.md` (generated) and `mps/README.md`.

## 1. Purpose and scope

The MPS replaces the three 2026 workbooks PMC_Operation_Data, PMC_Operation_Report and PMC_Operation_Commentary, whose inspection found external links, unmaintained macros and figures that depend on which copy was opened last.

Two phases: the Excel phase, delivered now, is one generated workbook per year that the site team runs without IT support; the web phase moves the same data model onto company servers with a database and browser forms. One schema file generates both.

Scope: daily and monthly production reporting against budget. Cost, maintenance work orders and metallurgical accounting adjustments are excluded; section 9 lists what was deferred.

## 2. Design principles

- Single system of record: a figure is typed once, in one flat table, by the department that owns it; reports hold no typed numbers.
- One calculation layer: Calc_Daily derives every KPI per day; reports read only Calc_Daily.
- No external links, no VBA, no hidden sheets; `tools/xl_inspect.py` can prove it.
- Schema-driven generation: `mps_schema.json` declares tables, fields, units, ranges, lists, calculated columns, highlights and natural keys; `build_workbook.py` renders it and refuses to build when a formula names a column or list the schema lacks; `build_ddl.py` emits the DDL and data dictionary from the same file.
- Blank means not reported; zero means reported as zero. A ratio pairs its numerator with the tonnes or hours that carry the measurement, so a missing assay leaves the period figure unchanged.
- Problems are shown, not hidden: missing grades, unknown sources, duplicate dates, a second feed stockpile and a year mismatch turn red where they occur and are counted on the report.

## 3. Workbook architecture

17 sheets, 20 tables, 332 defined names and about 162,000 formulas for 2026: README, Daily_Report, Dashboard and Monthly_Summary; eight input sheets (Plant, Mining_Movements, Mining_Daily, Fleet, Safety, Gold, Budget, Commentary) with one owner per table; Config and Lists; Calc_Daily, Chart_Data and Data_Dictionary.

### Calc_Daily

Each row is a date. A sum KPI is a SUMIFS on Date plus criteria such as Is_Ore=1 or Head_Grade_gpt not blank; MTD resets monthly, YTD runs on. A ratio KPI divides the matching numerator and denominator columns per period, blank when the denominator is zero, so it is never an average of days. Budgets: Bud_Month by SUMIFS on tblBudget, Bud_MTD pro-rata by day of month, Bud_YTD as completed months plus that. Ratio budgets derive from budgeted numerators and denominators; a helper KPI carries a budget alias (Ore_Graded_t reads the Ore_Mined_t budget) so the derived budget survives the pairing rule.

### Config

Typed settings: site, company, report title, the last LTI date before the year, and the schema constants as named cells `cfg_TroyOz` (31.1035), `cfg_HoursPerDay` (24) and `cfg_RateBasisHours` (1,000,000), referenced in table calcs and the KPI catalogue as `{HOURS_PER_DAY}` and `{RATE_BASIS_HOURS}`. Under "Calculated (do not type here)" sit `cfg_Year` (read from the first Calc_Daily date: the year is fixed by the generator and changed with `--year`), `cfg_YearStart`, `cfg_YearEnd`, `cfg_LastDataDate` (last date with milled tonnes), `cfg_FeedCheck` and `cfg_YearCheck`; the checks turn red when not OK and are repeated on the report. Config also holds the stockpile opening balances and the 24-month safety history.

### Daily_Report

C5 defaults to the earlier of yesterday and `cfg_LastDataDate`; a typed date overrides it and a stop validation keeps it inside the year. `rpt_Row` is the MATCH of the date in `cd_Date`; every cell is `INDEX(cd_<kpi>, rpt_Row)`. B7, named `rpt_Checks` and inside the print area, is the completeness and integrity line: rows for the report date in Plant, Movements, Fleet, Mining daily, Safety and Comments, movement rows with unknown source or destination, and movement rows dated outside the year.

Four bands (SAFETY, MINING, PROCESSING, GOLD) show day, MTD and YTD against budget; rate KPIs print "per 1,000,000 h" from `cfg_RateBasisHours` as their unit. MINING has two check rows that should read zero: "Ore mined without a grade (check)" and "Movements with unknown source type (check)". Then the stockpile block and fifteen commentary rows. The sheet prints on two portrait A4 pages, one page wide, with a manual break before PROCESSING.

### Stockpiles

One report row per row of tblStockpiles, including the five schema `spare_rows`, so a stockpile opened during the year appears once named. Each balance runs from the later of 1 January and the stockpile's Survey_Date (the opening survey; blank means 1 January) to the report date: opening plus deliveries minus reclaim, and for the stockpile flagged Plant_Feed = Y also minus the plant draw, milled less direct tip; grade follows the same balance in ounces. Exactly one stockpile may carry the flag: `cfg_FeedCheck` reads OK when the count is 1, otherwise the rule appears in red on Config and above the STOCKPILES header. The feed balance therefore covers ROM pad plus crusher surge stock; a crushed-ore balance is deferred.

### Monthly_Summary, Dashboard, protection

Monthly_Summary compares a completed month with its full budget; the month in progress shows the MTD actual at `cfg_LastDataDate` against `Bud_<key>_MTD` at that date, the budget pro-rata by day, and the Year column compares YTD with `Bud_<key>_YTD` there. Chart_Data returns NA() beyond `cfg_LastDataDate` or where the source is not numeric, so actual lines stop at the last data point while budget lines run the year; five charts print on one landscape A4 page.

Report and calculation sheets are protected without a password, against accidental typing only; input, Config and Lists stay open; nothing is hidden. A file received by e-mail opens behind the Protected View bar with a blank report until the user clicks Enable Editing. One person edits at a time: open, enter, save, close; if Excel opens the file Read-Only someone else has it, so wait and never Save As a copy; simultaneous entry needs SharePoint or OneDrive.

## 4. Data model

| Table | Grain and key | Owner and content |
|---|---|---|
| tblPlant | day, prefilled | Processing: crusher, mill, grades, gold poured, downtime, reagents, power, water, GIC |
| tblMovement | source, material, destination per day, optional shift; unique on the five | Mining dispatch: everything mined, rehandled or tipped |
| tblMiningDaily | day, prefilled | Mining: drilling, grade control, blasting, explosives, diesel, dewatering |
| tblFleet | unique on Date, Equipment_Class | Mining and maintenance: units, down and operating hours |
| tblSafety | day, prefilled | HSE: hours, injuries, incidents, leading indicators |
| tblSafetyHistory | 24 months, prior and reporting year | HSE: hours, recordables, LTI; Use flags months without daily hours |
| tblGold | shipment or sale | Finance and gold room: dore, gold and silver ounces |
| tblBudget | month | Technical services: budget or latest forecast |
| tblCommentary | comment | All departments: narrative by area and date |
| tblSources, tblDestinations | reference | Production engineer: type of each movement end, required once named |
| tblStockpiles | reference, five spare rows | Production engineer: opening tonnes, grade, survey date, plant feed flag |

### Schema keys and validation

Beyond the basic field keys, a field may carry `required_when` (red when the condition holds and the cell is blank; Data_Dictionary shows "when {Is_Ore}=1") with a `required_when_sql` twin, `warn_when` (red whenever the condition holds, used for Source_Type and Dest_Type = "?"), and `sql_postgres` and `sql_sqlite` where one expression cannot serve both dialects (tblBudget.Days). A table may declare `unique`, the column groups of its natural key (tblMovement, tblFleet); a reference table may declare `spare_rows`; `constants` holds the troy ounce, hours per day and rate basis.

Drop-downs are dynamic `lst_<list>` names; range checks are warnings. Dates are validated between `cfg_YearStart` and `cfg_YearEnd`, since a row outside the year vanishes from every KPI; the safety history accepts twelve months earlier. The key column of every prefilled table has a stop validation refusing a duplicate date plus a red highlight for a duplicate arriving by paste, which SUMIFS would double count.

### KPI catalogue

Mining. Ore mined is ex-pit only: Source typed Pit and Material starting with Ore. Waste is ex-pit rows without the ore flag, today including mineralised waste and topsoil. TMM is ore plus waste; strip ratio is waste over ore; rehandle is any row from a stockpile. Grade_gpt is required when Is_Ore = 1: Ore_Graded_t (pit ore tonnes with a grade) is the Ore_Grade_gpt denominator, so an ungraded row does not dilute the grade mined, and Ore_Ungraded_t (Ore_Mined_t minus Ore_Graded_t) is the check line. Moved_Total_t sums every movement row; Unclassified_t (Moved_Total_t minus TMM minus rehandle) exposes rows whose source is not in tblSources. GC_Drill_m reports grade control metres beside production drilling.

Fleet. Calendar hours are units times `cfg_HoursPerDay`. Exc_Productivity_tph and Trk_Productivity_tph divide TMM plus rehandle by excavator and by truck operating hours.

Processing. Recovered ounces are feed minus tails ounces; recovery is recovered over feed. Row-level guard: Feed_oz is blank without milled tonnes and a head grade, Tails_oz without a tails grade, Recovered_oz and Recovery_pct unless all three are present. Period ratios pair with the tonnes carrying the assay: Milled_Graded_t (milled with a head grade) is the Head_Grade_gpt denominator, Milled_Tails_Graded_t the Tails_Grade_gpt denominator and Feed_Recon_oz (feed ounces on days with a tails grade) the Recovery_pct denominator, so an outstanding assay never counts as zero grade or full recovery; Milled_Timed_t (milled on days with run hours) is the Throughput_tph numerator. Mill calendar hours are `cfg_HoursPerDay` on any day with run, planned or unplanned hours entered, else zero; a full-day shutdown is planned maintenance 24 with run hours 0, and availability and utilisation follow from that calendar. Gravity_Share_pct is gravity gold over gold recovered.

Safety. Recordables (fatalities, LTI, RWI, MTI) are blank while all four are blank. TRIFR and LTIFR are injuries per `cfg_RateBasisHours` over the window from twelve months before the first of the report month to the report date. Daily rows supply the window; an earlier month without daily hours has Use = 1 in tblSafetyHistory and takes hours and injuries from there, covering the prior year and the months before a mid-year go-live.

## 5. Daily operating procedure

The README suggests Mining 06:30 (one daily total or Day and Night movement rows, never both; the Mining_Daily row; one Fleet row per class), Plant 06:45, HSE 07:00 and Commentary 07:15; timings are proposals. The production engineer then checks the report date, the `rpt_Checks` line, the two check rows and any red cells, and exports and distributes the PDF. Paste only with Paste Special > Values; add rows at the bottom only; never insert columns. Month end: reconcile ore mined with grade control, milled tonnes with the metallurgical balance, stockpiles with the survey and hours with payroll, booking adjustments as dated rows.

## 6. Generator, validator and site scripts

`build_workbook.py --out FILE [--year 2026] [--sample-days N] [--report-date YYYY-MM-DD]` renders the workbook; `--sample-days` fills deterministic synthetic data marked [DEMO DATA]. After every build `check_catalogue()` verifies every table, column, list and budget reference in the KPI catalogue and in every formula sheet against the schema and raises an error naming the KPI or cell.

`validate_workbook.py [--year 2026] [--keep] [--workdir DIR] [--soffice PATH]` builds four workbooks (DEMO with a fixed and with the default report date, blank, and an edge case with partial rows, an unknown source, a prior-year row, a re-surveyed stockpile and a Probe sheet for every drop-down name), recalculates them headless in LibreOffice, recomputes the key figures in plain Python and compares them with Calc_Daily, the report, the summary and the charts, checks the completeness line, feed check, validations and highlights, scans for error values and runs `xl_inspect.py`: 931 checks in about a minute, exit code 1 on any failure.

On site, `Setup-MakoPS.ps1` (generated by `tools/build_setup.py`; switches `-Root`, `-Source`, `-Year`, `-SkipCopy`, `-SkipRun`, `-SkipBuild`) creates `C:\MakoPS` with `01_Tools`, `02_Source\2026`, `03_Inspection` and `04_MPS`. `PythonEnv.ps1` finds or installs a per-user Python and its modules. `Build-MPS.ps1` builds the blank and DEMO workbooks into `04_MPS`. `Export-DailyReport.ps1` writes the date into C5 through Excel COM, exports the sheet to `04_MPS\Reports\Mako_Daily_Report_yyyy-MM-dd.pdf` and can attach it to an Outlook mail.

## 7. Migration

Mark which legacy sheets hold source data; unpivot the daily mining sheets into the movements ledger; load plant daily rows and compare recomputed recovery with the legacy figure; fill the safety history for the twelve months before the reporting year and any month before go-live; enter the opening stockpile surveys with their dates and flag the plant feed stockpile. Parallel run for one month, resolving differences in the schema or the data, never in the outputs. Cut over when month-end totals agree within tolerance (tonnes 0.5 percent, ounces 1 percent, hours exact), every owner has entered their table unaided for ten consecutive days, the checks read zero and the inspector reports no errors.

## 8. Web phase

Stack: Python with FastAPI or Django, PostgreSQL, server-rendered HTML with minimal JavaScript, on Windows Server or Linux; local accounts first, Active Directory later. `build_ddl.py` already emits tables with CHECK constraints from min and max, natural keys from `unique`, foreign keys to the list tables, generated columns for calculated fields and `kpi_daily` and `kpi_monthly` views reproducing the whole KPI catalogue, helper KPIs included. A nightly import reads the workbook table by table during transition; roles per table mirror the owner column. Roll-out: database and import, read-only reports, entry forms per department, retire the workbook.

## 9. Backlog for the web phase and open design decisions

- Movement grain (F014). The Excel phase keeps "one daily total or Day and Night rows, never both" as a procedure. The web phase adds a Full Day shift value, enforces the `unique` key and checks for mixed grain. Open: which grain dispatch will use.
- Stockpile survey adjustments (F015). An Adjustment movement type belongs with the web-phase reconciliation adjustments; until then a survey is entered as a Config opening balance with its Survey_Date.
- Crushed ore stockpile (F016) and the feed balance (F043). Without a crushed-ore stockpile the feed balance covers ROM pad plus crusher surge stock (milled less direct tip, cumulative), which conserves tonnes where a daily MAX(0, milled minus tip) would debit the pad twice. A separate crushed balance needs a new opening balance and a redesigned ROM reconciliation once Processing confirms how it is surveyed. Confirm that one stockpile feeds the plant and that mineralised waste and topsoil count as waste.
- Metallurgical accounting (F017). Propose tblMetAdjust keyed on month with a met balance KPI. Open: daily head and tails, or head back-calculated monthly from poured gold and GIC?
- Mill standby hours (F018). Mill_Standby_h changes the plant time model; agree definitions with Processing, then add the field and an unaccounted-hours check.
- Gold room (F019). Dore on hand, outturn adjustment, shipment link and price. Open: booking point (pour, dispatch or outturn) and how adjustments are recorded.
- Budget and forecast scenarios (F020). A Scenario key on tblBudget and cfg_PlanScenario. Open: approved budget, latest forecast, or both?
- Week-to-date and quarter-to-date (F021). WTD and QTD running totals in Calc_Daily once the period model is defined.
- French labels and code versus label (F022). A Materials reference table replacing the Ore prefix rule, with French for workforce-facing screens and lists.
- TSF, water balance, rainfall and WAD cyanide (F023). Add to tblPlant or tblSafety once Environment confirms the daily figures.
- Headcount and persons on board (F024). Point fields in tblSafety when HSE confirms the categories.
- Fleet down hours (F048). Splitting Down_h into planned and unplanned needs definitions agreed with Mining and Maintenance, then both fields with Down_h as a calculated sum. Related: contractor versus owner fleet, one time model or separate classes?
- Occupational health (F050). Occupational illness and malaria cases as integer fields in tblSafety with sum KPIs once HSE confirms the case definitions.
- Also open: pit and stage names for the Sources list; the grade source for movements (grade control model, blast-hole assays or a blend); cut-off time per table and the PDF deadline; the PDF distribution list, including the owner in Perth; cost reporting (unit costs per tonne and per ounce, and from which system).
