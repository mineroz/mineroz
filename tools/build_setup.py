#!/usr/bin/env python3
"""
build_setup.py - assemble Setup-MakoPS.ps1, a single self-contained PowerShell
script that lays out C:\\MakoPS, writes the tooling into it, copies the 2026 PMC
workbooks from the X: drive, runs the audit and builds the MPS workbooks.

Run from the repo root:
    python tools/build_setup.py            regenerate Setup-MakoPS.ps1 and verify the round trip
    python tools/build_setup.py --check    only verify that the existing Setup-MakoPS.ps1 matches the sources

Every embedded file travels inside a single-quoted PowerShell here-string, so nothing is
interpolated. The one thing such a file must not contain is a line starting with '@ (the
here-string terminator); the build refuses to embed such a file.
"""
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "Setup-MakoPS.ps1"

# (source relative to the repo root, target relative to <Root>, write BOM, required)
# Optional files are skipped with a warning when absent at build time.
EMBED = [
    ("tools/xl_inspect.py", r"01_Tools\xl_inspect.py", False, True),
    ("tools/PythonEnv.ps1", r"01_Tools\PythonEnv.ps1", True, True),
    ("tools/Invoke-XlInspect.ps1", r"01_Tools\Invoke-XlInspect.ps1", True, True),
    ("tools/Build-MPS.ps1", r"01_Tools\Build-MPS.ps1", True, True),
    ("tools/Export-DailyReport.ps1", r"01_Tools\Export-DailyReport.ps1", True, True),
    ("tools/README.md", r"01_Tools\README.md", False, True),
    ("mps/build_workbook.py", r"01_Tools\mps\build_workbook.py", False, True),
    ("mps/build_ddl.py", r"01_Tools\mps\build_ddl.py", False, False),
    ("mps/validate_workbook.py", r"01_Tools\mps\validate_workbook.py", False, False),
    ("mps/schema/mps_schema.json", r"01_Tools\mps\schema\mps_schema.json", False, True),
    ("mps/README.md", r"01_Tools\mps\README.md", False, False),
]

LAYOUT = ["01_Tools", r"01_Tools\mps\schema", r"02_Source\2026", "03_Inspection", "04_MPS"]

HEADER = r'''<#
.SYNOPSIS
    One-shot setup for the Mako Production System working folder on C:.
.DESCRIPTION
    Creates C:\MakoPS with this layout:
        01_Tools         tooling: xl_inspect.py, PythonEnv.ps1, Invoke-XlInspect.ps1, Build-MPS.ps1,
                         Export-DailyReport.ps1, README.md
        01_Tools\mps     MPS generator: build_workbook.py, build_ddl.py, validate_workbook.py,
                         schema\mps_schema.json
        02_Source\2026   working copies of the PMC operation workbooks (originals on X: are not touched)
        03_Inspection    inspection_report.md plus one JSON per workbook
        04_MPS           MPS_<Year>.xlsx (blank) and MPS_<Year>_DEMO.xlsx (synthetic data)
    Then runs the audit and builds the MPS workbooks. Re-running is safe: files are overwritten,
    nothing is deleted. Python 3 is resolved once (installed for the current user when it is
    missing); when that fails the audit and the build are skipped with a warning and the setup
    still finishes, so the tooling is in place for a later run.
    Run from a normal PowerShell window: mapped drives such as X: are not visible in an
    elevated (Run as administrator) session.
.PARAMETER Root
    Target folder. Default C:\MakoPS.
.PARAMETER Source
    Folder holding the workbooks to copy. Default is the 2026 operation report folder on X:.
.PARAMETER Year
    Reporting year of the MPS workbooks. Default 2026.
.PARAMETER SkipCopy
    Do not copy workbooks from X: (use when 02_Source\2026 is already populated).
.PARAMETER SkipRun
    Do not run the inspection.
.PARAMETER SkipBuild
    Do not build the MPS workbooks.
.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Setup-MakoPS.ps1
    powershell -ExecutionPolicy Bypass -File .\Setup-MakoPS.ps1 -SkipCopy -SkipRun
#>
[CmdletBinding()]
param(
    [string]$Root = "C:\MakoPS",
    [string]$Source = "X:\13_Management\01_GM_Dashboard\01_Operation_Report\01_Excel\2026",
    [int]$Year = 2026,
    [switch]$SkipCopy,
    [switch]$SkipRun,
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"
# Absolute root: the .NET file writer below resolves relative paths against the process directory, not the PowerShell location
$Root = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Root)
$utf8NoBom = New-Object System.Text.UTF8Encoding $false
$utf8Bom   = New-Object System.Text.UTF8Encoding $true

function Write-TextFile([string]$Path, [string]$Content, [bool]$Bom) {
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    $enc = if ($Bom) { $utf8Bom } else { $utf8NoBom }
    [System.IO.File]::WriteAllText($Path, $Content, $enc)
    Write-Host ("  wrote {0}  ({1:N0} bytes)" -f $Path, (Get-Item $Path).Length)
}

$elevated = $false
try {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $elevated = (New-Object Security.Principal.WindowsPrincipal $identity).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
} catch { }
if ($elevated) {
    Write-Warning "This PowerShell session is elevated (Run as administrator). Mapped drives such as X: are not visible here; if the copy step fails, re-run from a normal PowerShell window."
}

Write-Host "Creating folder layout under $Root"
foreach ($sub in @(__LAYOUT__)) {
    New-Item -ItemType Directory -Force -Path (Join-Path $Root $sub) | Out-Null
}

Write-Host "Writing tooling"
'''

FOOTER = r'''
# ---------------------------------------------------------------- copy workbooks
$sourceCopy = Join-Path $Root "02_Source\2026"
if (-not $SkipCopy) {
    if (Test-Path $Source) {
        $files = Get-ChildItem -Path $Source -File | Where-Object {
            $_.Extension -match '^\.(xlsx|xlsm|xlsb|xls|xltm|xltx)$' -and $_.Name -notlike '~$*'
        }
        if ($files.Count -eq 0) {
            Write-Warning "No Excel files found in $Source"
        } else {
            Write-Host "Copying $($files.Count) workbook(s) from $Source to $sourceCopy"
            foreach ($f in $files) {
                Copy-Item -LiteralPath $f.FullName -Destination (Join-Path $sourceCopy $f.Name) -Force
                Write-Host ("  {0}  ({1:N1} MB)" -f $f.Name, ($f.Length / 1MB))
            }
        }
    } else {
        Write-Warning "Source folder not reachable: $Source. Copy the workbooks into $sourceCopy manually, then run 01_Tools\Invoke-XlInspect.ps1."
        if ($elevated) { Write-Warning "The session is elevated, which hides mapped drives such as X:. Re-run from a normal PowerShell window." }
    }
}

# ---------------------------------------------------------------- python (resolved once for both scripts below)
$pyOk = $false
if (-not ($SkipRun -and $SkipBuild)) {
    Write-Host ""
    . (Join-Path $Root "01_Tools\PythonEnv.ps1")
    try {
        $py = Resolve-Python
        Write-Host "Using Python: $py"
        $pyOk = $true
    } catch {
        Write-Warning $_.Exception.Message
        Write-Warning "Inspection and workbook build skipped. Install Python 3, then run 01_Tools\Invoke-XlInspect.ps1 and 01_Tools\Build-MPS.ps1 from $Root."
    }
}

# ---------------------------------------------------------------- run inspection
if (-not $SkipRun -and $pyOk) {
    $runner = Join-Path $Root "01_Tools\Invoke-XlInspect.ps1"
    $books = @(Get-ChildItem -Path $sourceCopy -File -ErrorAction SilentlyContinue | Where-Object {
        $_.Extension -match '^\.(xlsx|xlsm|xlsb|xls|xltm|xltx)$' -and $_.Name -notlike '~$*'
    })
    if ($books.Count -eq 0) {
        Write-Warning "No workbooks in $sourceCopy, skipping the inspection. Copy them there and run $runner later."
    } else {
        Write-Host ""
        Write-Host "Running inspection"
        try {
            & $runner -Source $sourceCopy -Out (Join-Path $Root "03_Inspection") -NoInstall
        } catch {
            Write-Warning "Inspection failed: $($_.Exception.Message). Fix the cause and run $runner again."
        }
        $report = Join-Path $Root "03_Inspection\inspection_report.md"
        if (Test-Path $report) {
            Write-Host ""
            Write-Host "Inspection report: $report"
        }
    }
}

# ---------------------------------------------------------------- build MPS workbooks
if (-not $SkipBuild -and $pyOk) {
    $builder = Join-Path $Root "01_Tools\Build-MPS.ps1"
    Write-Host ""
    Write-Host "Building the MPS workbooks for $Year"
    try {
        & $builder -Root $Root -Year $Year -NoInstall -NoOpen
    } catch {
        Write-Warning "Build failed: $($_.Exception.Message). Fix the cause and run $builder again."
    }
}

Write-Host ""
Write-Host "Setup finished. Tools are in $(Join-Path $Root '01_Tools'); see README.md there."
# one Explorer window on the root, after everything has been written
try { Start-Process explorer.exe $Root } catch { }
'''


def here_string(text: str, name: str) -> str:
    """Wrap text in a single-quoted here-string (no interpolation)."""
    # The closing '@ must start a line, so any line beginning with '@ would end the here-string early.
    if text.startswith("'@") or "\n'@" in text:
        raise ValueError(f"{name}: a line starting with '@ would terminate the here-string")
    return "@'\n" + text + "\n'@"


def var_name(src: str) -> str:
    return "$content_" + re.sub(r"[^A-Za-z0-9]", "_", src)


def normalise(text: str) -> str:
    return text.replace("\r\n", "\n").replace("\r", "\n").rstrip("\n")


def load_sources():
    """Return [(src, target, bom, text)] for the files present; warn for missing optional ones."""
    items = []
    for src, target, bom, required in EMBED:
        path = ROOT / src
        if not path.is_file():
            if required:
                raise FileNotFoundError(f"required file missing: {path}")
            print(f"warning: optional file not found, skipped: {src}", file=sys.stderr)
            continue
        text = normalise(path.read_text(encoding="utf-8-sig"))
        items.append((src, target, bom, text))
    return items


def build() -> None:
    items = load_sources()
    parts = [HEADER.replace("__LAYOUT__", ", ".join(f'"{s}"' for s in LAYOUT))]
    for src, target, bom, text in items:
        var = var_name(src)
        parts.append(f"{var} = {here_string(text, src)}\n")
        parts.append(f'Write-TextFile (Join-Path $Root "{target}") {var} ${str(bom).lower()}\n')
    parts.append(FOOTER)
    body = "".join(parts).replace("\r\n", "\n").replace("\n", "\r\n")
    OUT.write_bytes(b"\xef\xbb\xbf" + body.encode("utf-8"))
    print(f"wrote {OUT} ({OUT.stat().st_size:,} bytes, {len(items)} embedded files)")


BLOCK_RE = re.compile(
    r"^(\$content_\w+) = @'\r\n(.*?)\r\n'@\r\n"
    r"Write-TextFile \(Join-Path \$Root \"([^\"]+)\"\) \$content_\w+ \$(true|false)\r\n",
    re.S | re.M,
)


def check() -> int:
    """Extract every here-string from Setup-MakoPS.ps1 and compare with the source files."""
    raw = OUT.read_bytes()
    problems = []
    if not raw.startswith(b"\xef\xbb\xbf"):
        problems.append("Setup-MakoPS.ps1 has no UTF-8 BOM")
    text = raw[3:].decode("utf-8")
    if "\n" in text.replace("\r\n", ""):
        problems.append("Setup-MakoPS.ps1 has bare LF line endings")
    found = {m.group(1): (m.group(2), m.group(3), m.group(4)) for m in BLOCK_RE.finditer(text)}
    for src, target, bom, expected in load_sources():
        var = var_name(src)
        if var not in found:
            problems.append(f"{src}: no here-string found for {var}")
            continue
        got, got_target, got_bom = found.pop(var)
        if got.replace("\r\n", "\n") != expected:
            problems.append(f"{src}: embedded content differs from the source")
        if got_target != target:
            problems.append(f"{src}: target is {got_target}, expected {target}")
        if got_bom != str(bom).lower():
            problems.append(f"{src}: BOM flag is {got_bom}, expected {str(bom).lower()}")
    for var in found:
        problems.append(f"{var}: embedded but not in EMBED")
    for sub in LAYOUT:
        if f'"{sub}"' not in text:
            problems.append(f"layout folder {sub} missing from the setup script")
    for p in problems:
        print("check FAILED:", p, file=sys.stderr)
    if not problems:
        print(f"check OK: {OUT.name} matches {len(load_sources())} source files")
    return 1 if problems else 0


def main(argv=None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    if argv and argv[0] == "--check":
        return check()
    if argv:
        print(__doc__)
        return 2
    build()
    return check()


if __name__ == "__main__":
    sys.exit(main())
