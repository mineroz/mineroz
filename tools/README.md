# tools

Scripts that run on the site PCs (Windows PowerShell 5.1, no admin rights) and the
build script that packs them into the single `Setup-MakoPS.ps1` bootstrap.

## Folder layout on the PC

`Setup-MakoPS.ps1` creates and fills this layout (default root `C:\MakoPS`):

| Folder | Content |
|---|---|
| `01_Tools` | the scripts below, `xl_inspect.py`, `README.md` (this file) |
| `01_Tools\mps` | MPS generator: `build_workbook.py`, `build_ddl.py`, `validate_workbook.py`, `schema\mps_schema.json` |
| `02_Source\2026` | working copies of the PMC operation workbooks (originals on X: are never touched) |
| `03_Inspection` | `inspection_report.md` plus one JSON per inspected workbook |
| `04_MPS` | `MPS_<year>.xlsx` (blank, for production use) and `MPS_<year>_DEMO.xlsx` (synthetic data) |

Run everything from a normal PowerShell window. Mapped drives such as `X:` are not visible in
an elevated (Run as administrator) session, so the copy step would find nothing there. If the
execution policy blocks scripts, start them with
`powershell -ExecutionPolicy Bypass -File <script>`.

## PythonEnv.ps1

Shared helper, dot-sourced by the other scripts (`. (Join-Path $here "PythonEnv.ps1")`). No
parameters of its own.

| Function | Purpose |
|---|---|
| `Find-Python` | first real Python 3.8+ on the PC: PATH, the `py` launcher, per-user and Program Files installs, Anaconda and Miniconda. The Microsoft Store stub under `\Microsoft\WindowsApps\` is ignored. |
| `Install-Python` | silent per-user install of Python 3.12 from python.org (`InstallAllUsers=0 PrependPath=1`, no admin rights), winget as fallback, then refreshes PATH for the running session |
| `Test-PyModule` | `$true` when the interpreter can import a module |
| `Ensure-PyModules` | `pip install` of each missing module, retried with `--user`, with a manual command in the error when that fails too |
| `Resolve-Python` | `Find-Python`, then `Install-Python` unless `-NoInstall`, then `Find-Python` again |

The functions run native executables, so they set `$ErrorActionPreference = "Continue"` in
their own scope and test `$LASTEXITCODE` explicitly: Windows PowerShell 5.1 turns anything a
native command writes to stderr into a terminating error under `Stop`, and pip does write
warnings there.

## Invoke-XlInspect.ps1

Runs `xl_inspect.py` over a folder of workbooks.

```powershell
.\Invoke-XlInspect.ps1                       # <root>\02_Source\2026 to <root>\03_Inspection
.\Invoke-XlInspect.ps1 -Source "X:\13_Management\01_GM_Dashboard\01_Operation_Report\01_Excel\2026"
.\Invoke-XlInspect.ps1 -NoInstall            # fail instead of installing Python
```

Needs `openpyxl` and `oletools`; both are installed when missing.

## Build-MPS.ps1

Builds the Mako Production System workbooks from `mps\schema\mps_schema.json`.

```powershell
.\Build-MPS.ps1                              # C:\MakoPS\04_MPS\MPS_2026.xlsx and MPS_2026_DEMO.xlsx
.\Build-MPS.ps1 -Root D:\MakoPS -Year 2027 -SampleDays 30
.\Build-MPS.ps1 -NoOpen                      # do not open 04_MPS in Explorer
```

| Parameter | Default | Meaning |
|---|---|---|
| `-Root` | `C:\MakoPS` | layout root; output goes to `<Root>\04_MPS` |
| `-Year` | 2026 | reporting year of the workbook |
| `-SampleDays` | 74 | days of synthetic data in the DEMO workbook, 0 skips it |
| `-NoInstall` | | never install Python |
| `-NoOpen` | | do not open the output folder |

The generator is taken from `<Root>\01_Tools\mps\build_workbook.py`; when the script runs from
this repository, `<repo>\mps\build_workbook.py` is used. Each workbook takes a few seconds.
Existing files are overwritten, so close them in Excel first (the script checks and stops with
a message if a file is locked). Open the DEMO workbook to see the Dashboard, Daily_Report and
Monthly_Summary populated; enter real data only in the blank workbook.

## Export-DailyReport.ps1

Produces the daily PDF from an MPS workbook with Excel COM automation, so the workbook itself
needs no macros. The script opens the workbook read-only, writes the date into
`Daily_Report!C5`, lets Excel calculate, exports the sheet with `ExportAsFixedFormat` and closes
Excel without saving. Requires Microsoft Excel on the PC; `-Email` additionally requires classic
Outlook (the new Outlook has no COM automation).

```powershell
.\Export-DailyReport.ps1                                   # yesterday, C:\MakoPS\04_MPS\MPS_<year>.xlsx
.\Export-DailyReport.ps1 -Workbook C:\MakoPS\04_MPS\MPS_2026_DEMO.xlsx -Date 2026-03-15 -Open
.\Export-DailyReport.ps1 -Date 2026-03-15 -Email -To gm@example.com,ops@example.com
```

| Parameter | Default | Meaning |
|---|---|---|
| `-Workbook` | `<Root>\04_MPS\MPS_<year of Date>.xlsx` | workbook to export from |
| `-Date` | yesterday | report date, write it as `yyyy-MM-dd` |
| `-OutDir` | `<workbook folder>\Reports` | where the PDF goes, created when missing |
| `-Root` | `C:\MakoPS` | only used for the default workbook path |
| `-Sheet`, `-DateCell` | `Daily_Report`, `C5` | report sheet and its date cell |
| `-Email -To a,b -Cc c -Subject s` | | display an Outlook mail with the PDF attached (nothing is sent) |
| `-Open` | | open the PDF when done |

Output: `<OutDir>\Mako_Daily_Report_yyyy-MM-dd.pdf`. A warning is printed when the requested
date lies after `cfg_LastDataDate`, because the report would then be blank. To run it every
morning, create a Task Scheduler task with the action
`powershell -ExecutionPolicy Bypass -File C:\MakoPS\01_Tools\Export-DailyReport.ps1`.

## xl_inspect.py

Structural audit of Excel workbooks (.xlsx / .xlsm) ahead of the Mako Production System rebuild.
For each workbook it reports:

- sheets with state (visible / hidden / veryHidden), used range, value and formula counts
- external links in formula index order, whether the target exists on disk, and how many formulas use each
- cross-sheet dependency graph and references to sheets that no longer exist
- defined names, flagging `#REF!` and external targets
- VBA modules, procedures, auto-executing events, hard-coded paths and keywords worth reviewing (needs `oletools`)
- data connections, Power Query parts, pivot caches, charts, tables, slicers
- volatile functions (OFFSET, INDIRECT, NOW, TODAY, RAND), lookups, array formulas
- cached error values (`#REF!`, `#N/A`, `#VALUE!`, `#DIV/0!`, `#NAME?`) per sheet

Legacy `.xls` and binary `.xlsb` are listed but not parsed; save them as `.xlsm` first.

The scan streams every sheet (openpyxl read-only) and reads sheet features straight from the
XML, so memory stays flat on large workbooks and a bloated declared range (for example
`A1:Z1048576` with 400 real rows) costs nothing. Such sheets are flagged `BLOAT` in the report.

Any platform:

```bash
pip install openpyxl oletools
python tools/xl_inspect.py <folder-or-files> -o <output-dir>
```

Output: `inspection_report.md` plus one `<workbook>.json` per file. It is also the quickest
check of a freshly built MPS workbook: `python tools/xl_inspect.py C:\MakoPS\04_MPS -o C:\MakoPS\03_Inspection\mps`.

## build_setup.py

Packs the scripts above, `xl_inspect.py`, this README and the `mps` generator files
(`build_workbook.py`, `build_ddl.py`, `validate_workbook.py`, `schema\mps_schema.json`,
`README.md`; the optional ones are skipped with a warning when absent) into
`Setup-MakoPS.ps1` at the repository root. Each file travels in a single-quoted PowerShell
here-string, so nothing is interpolated; a file containing a line that starts with `'@` cannot
be embedded and the build stops with an error.

```bash
python tools/build_setup.py            # regenerate Setup-MakoPS.ps1, then verify the round trip
python tools/build_setup.py --check    # only verify the existing Setup-MakoPS.ps1 against the sources
```

The verification extracts every here-string from the generated file and compares it with its
source, the target path and the BOM flag. Re-run the build after changing any embedded file.
`Setup-MakoPS.ps1` accepts `-Root`, `-Source`, `-Year`, `-SkipCopy`, `-SkipRun` and `-SkipBuild`.

## Conventions

- PowerShell files: Windows PowerShell 5.1 compatible (no ternary operator, no null-coalescing,
  no PowerShell 7 only cmdlets), saved as UTF-8 with BOM and CRLF.
- Python files: Python 3.8+, `openpyxl` only (plus `oletools` for the VBA scan).
- British English, metric units, gold in troy ounces, grades in g/t.
