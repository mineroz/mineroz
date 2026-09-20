# tools

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

`Invoke-XlInspect.ps1` finds a real Python 3 (it ignores the Microsoft Store stub in
`WindowsApps`). If none is installed it installs Python 3.12 for the current user from
python.org without admin rights, falling back to winget. Pass `-NoInstall` to disable this.

### Run

Windows, from the repo root:

```powershell
.\tools\Invoke-XlInspect.ps1
.\tools\Invoke-XlInspect.ps1 -Source "X:\13_Management\01_GM_Dashboard\01_Operation_Report\01_Excel\2026"
```

Any platform:

```bash
pip install openpyxl oletools
python tools/xl_inspect.py <folder-or-files> -o <output-dir>
```

Output: `inspection_report.md` plus one `<workbook>.json` per file.
