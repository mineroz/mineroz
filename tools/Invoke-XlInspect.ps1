<#
.SYNOPSIS
    Runs xl_inspect.py against the PMC operation report workbooks.
.DESCRIPTION
    Installs openpyxl and oletools into the current Python if missing, then
    produces inspection_report.md and one JSON per workbook in the output folder.
.PARAMETER Source
    Folder holding the workbooks. Defaults to the 2026 operation report folder.
.PARAMETER Out
    Output folder for the report. Defaults to <Source>\_inspection.
.EXAMPLE
    .\Invoke-XlInspect.ps1
    .\Invoke-XlInspect.ps1 -Source "X:\13_Management\01_GM_Dashboard\01_Operation_Report\01_Excel\2026"
#>
[CmdletBinding()]
param(
    [string]$Source = "X:\13_Management\01_GM_Dashboard\01_Operation_Report\01_Excel\2026",
    [string]$Out = ""
)

$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$script = Join-Path $here "xl_inspect.py"

if (-not (Test-Path $Source)) { throw "Source folder not found: $Source" }
if ([string]::IsNullOrWhiteSpace($Out)) { $Out = Join-Path $Source "_inspection" }

$python = Get-Command python -ErrorAction SilentlyContinue
if (-not $python) { $python = Get-Command py -ErrorAction SilentlyContinue }
if (-not $python) { throw "Python 3 not found on PATH. Install from python.org or the Microsoft Store." }
$py = $python.Source

& $py -c "import openpyxl" 2>$null
if ($LASTEXITCODE -ne 0) { Write-Host "Installing openpyxl ..."; & $py -m pip install --quiet openpyxl }
& $py -c "import oletools" 2>$null
if ($LASTEXITCODE -ne 0) { Write-Host "Installing oletools ..."; & $py -m pip install --quiet oletools }

Write-Host "Inspecting workbooks in $Source"
& $py $script $Source -o $Out
if ($LASTEXITCODE -ne 0) { throw "xl_inspect.py failed with exit code $LASTEXITCODE" }

$report = Join-Path $Out "inspection_report.md"
Write-Host ""
Write-Host "Report: $report"
Write-Host "Attach the whole '$Out' folder (report + JSON files) to the session, or commit it under data/2026/_inspection."
