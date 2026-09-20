#!/usr/bin/env python3
"""
build_setup.py - assemble Setup-MakoPS.ps1, a single self-contained PowerShell
script that lays out C:\\MakoPS, writes the inspection tooling into it, copies
the 2026 PMC workbooks from the X: drive, and runs the audit.

Run from the repo root:  python tools/build_setup.py
"""
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TOOLS = ROOT / "tools"
OUT = ROOT / "Setup-MakoPS.ps1"

EMBED = [  # (source file, target relative to C:\MakoPS, write BOM)
    ("xl_inspect.py", r"01_Tools\xl_inspect.py", False),
    ("Invoke-XlInspect.ps1", r"01_Tools\Invoke-XlInspect.ps1", True),
    ("README.md", r"01_Tools\README.md", False),
]

HEADER = r'''<#
.SYNOPSIS
    One-shot setup for the Mako Production System working folder on C:.
.DESCRIPTION
    Creates C:\MakoPS with this layout:
        01_Tools         inspection tooling (xl_inspect.py, Invoke-XlInspect.ps1, README.md)
        02_Source\2026   working copies of the PMC operation workbooks (originals on X: are not touched)
        03_Inspection    inspection_report.md plus one JSON per workbook
    Then runs the audit. Re-running is safe: files are overwritten, nothing is deleted.
.PARAMETER Root
    Target folder. Default C:\MakoPS.
.PARAMETER Source
    Folder holding the workbooks to copy. Default is the 2026 operation report folder on X:.
.PARAMETER SkipCopy
    Do not copy workbooks from X: (use when 02_Source\2026 is already populated).
.PARAMETER SkipRun
    Write files only, do not run the inspection.
.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Setup-MakoPS.ps1
#>
[CmdletBinding()]
param(
    [string]$Root = "C:\MakoPS",
    [string]$Source = "X:\13_Management\01_GM_Dashboard\01_Operation_Report\01_Excel\2026",
    [switch]$SkipCopy,
    [switch]$SkipRun
)

$ErrorActionPreference = "Stop"
$utf8NoBom = New-Object System.Text.UTF8Encoding $false
$utf8Bom   = New-Object System.Text.UTF8Encoding $true

function Write-TextFile([string]$Path, [string]$Content, [bool]$Bom) {
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    $enc = if ($Bom) { $utf8Bom } else { $utf8NoBom }
    [System.IO.File]::WriteAllText($Path, $Content, $enc)
    Write-Host ("  wrote {0}  ({1:N0} bytes)" -f $Path, (Get-Item $Path).Length)
}

Write-Host "Creating folder layout under $Root"
foreach ($sub in @("01_Tools", "02_Source\2026", "03_Inspection")) {
    New-Item -ItemType Directory -Force -Path (Join-Path $Root $sub) | Out-Null
}

Write-Host "Writing tooling"
'''

FOOTER = r'''
# ---------------------------------------------------------------- copy workbooks
if (-not $SkipCopy) {
    $dest = Join-Path $Root "02_Source\2026"
    if (Test-Path $Source) {
        $files = Get-ChildItem -Path $Source -File | Where-Object {
            $_.Extension -match '^\.(xlsx|xlsm|xlsb|xls|xltm|xltx)$' -and $_.Name -notlike '~$*'
        }
        if ($files.Count -eq 0) {
            Write-Warning "No Excel files found in $Source"
        } else {
            Write-Host "Copying $($files.Count) workbook(s) from $Source to $dest"
            foreach ($f in $files) {
                Copy-Item -LiteralPath $f.FullName -Destination (Join-Path $dest $f.Name) -Force
                Write-Host ("  {0}  ({1:N1} MB)" -f $f.Name, ($f.Length / 1MB))
            }
        }
    } else {
        Write-Warning "Source folder not reachable: $Source. Copy the workbooks into $dest manually, then run 01_Tools\Invoke-XlInspect.ps1."
    }
}

# ---------------------------------------------------------------- run inspection
if (-not $SkipRun) {
    $runner = Join-Path $Root "01_Tools\Invoke-XlInspect.ps1"
    Write-Host ""
    Write-Host "Running inspection"
    & $runner -Source (Join-Path $Root "02_Source\2026") -Out (Join-Path $Root "03_Inspection")
    $report = Join-Path $Root "03_Inspection\inspection_report.md"
    if (Test-Path $report) {
        Write-Host ""
        Write-Host "Done. Report: $report"
        try { Start-Process explorer.exe (Join-Path $Root "03_Inspection") } catch { }
    }
}
'''


def here_string(text: str) -> str:
    # Single-quoted here-string: no interpolation. Content must not contain a line that is exactly "'@".
    assert "\n'@" not in text and not text.startswith("'@"), "content would terminate the here-string"
    return "@'\n" + text + "\n'@"


def main() -> None:
    parts = [HEADER]
    for src, target, bom in EMBED:
        text = (TOOLS / src).read_text(encoding="utf-8-sig")
        text = text.replace("\r\n", "\n").rstrip("\n")
        var = "$content_" + src.replace(".", "_").replace("-", "_")
        parts.append(f"{var} = {here_string(text)}\n")
        parts.append(f'Write-TextFile (Join-Path $Root "{target}") {var} ${str(bom).lower()}\n')
    parts.append(FOOTER)
    body = "".join(parts).replace("\r\n", "\n").replace("\n", "\r\n")
    OUT.write_bytes(b"\xef\xbb\xbf" + body.encode("utf-8"))
    print(f"wrote {OUT} ({OUT.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()
