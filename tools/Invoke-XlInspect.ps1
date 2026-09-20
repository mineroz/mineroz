<#
.SYNOPSIS
    Runs xl_inspect.py against the PMC operation report workbooks.
.DESCRIPTION
    Installs openpyxl and oletools into the current Python if missing, then
    produces inspection_report.md and one JSON per workbook in the output folder.
.PARAMETER Source
    Folder holding the workbooks. Defaults to <root>\02_Source\2026 when it exists,
    otherwise the 2026 operation report folder on X:.
.PARAMETER Out
    Output folder for the report. Defaults to <root>\03_Inspection in the C:\MakoPS
    layout, otherwise <Source>\_inspection.
.EXAMPLE
    .\Invoke-XlInspect.ps1
    .\Invoke-XlInspect.ps1 -Source "X:\13_Management\01_GM_Dashboard\01_Operation_Report\01_Excel\2026"
#>
[CmdletBinding()]
param(
    [string]$Source = "",
    [string]$Out = ""
)

# Continue, not Stop: Windows PowerShell 5.1 turns native stderr output into terminating
# errors under Stop, and pip / openpyxl legitimately write warnings to stderr.
$ErrorActionPreference = "Continue"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$script = Join-Path $here "xl_inspect.py"
$root = Split-Path -Parent $here

# Default layout: <root>\01_Tools (this script), <root>\02_Source\2026, <root>\03_Inspection
if ([string]::IsNullOrWhiteSpace($Source)) {
    $local = Join-Path $root "02_Source\2026"
    if (Test-Path $local) { $Source = $local }
    else { $Source = "X:\13_Management\01_GM_Dashboard\01_Operation_Report\01_Excel\2026" }
}
if (-not (Test-Path $Source)) { throw "Source folder not found: $Source" }
if ([string]::IsNullOrWhiteSpace($Out)) {
    if (Test-Path (Join-Path $root "02_Source")) { $Out = Join-Path $root "03_Inspection" }
    else { $Out = Join-Path $Source "_inspection" }
}

$python = Get-Command python -ErrorAction SilentlyContinue
if (-not $python) { $python = Get-Command py -ErrorAction SilentlyContinue }
if (-not $python) { throw "Python 3 not found on PATH. Install from python.org or the Microsoft Store." }
$py = $python.Source

function Test-PyModule([string]$Module) {
    # find_spec writes nothing to stderr, so this is safe on PowerShell 5.1
    & $py -c "import importlib.util, sys; sys.exit(0 if importlib.util.find_spec('$Module') else 1)"
    return ($LASTEXITCODE -eq 0)
}
foreach ($m in @("openpyxl", "oletools")) {
    if (-not (Test-PyModule $m)) {
        Write-Host "Installing $m ..."
        & $py -m pip install --quiet --disable-pip-version-check $m
        if (-not (Test-PyModule $m)) { throw "Could not install Python package '$m'. Install it manually:  $py -m pip install $m" }
    }
}

Write-Host "Inspecting workbooks in $Source"
& $py $script $Source -o $Out
if ($LASTEXITCODE -ne 0) { throw "xl_inspect.py failed with exit code $LASTEXITCODE" }

$report = Join-Path $Out "inspection_report.md"
Write-Host ""
Write-Host "Report: $report"
Write-Host "Attach the whole '$Out' folder (report + JSON files) to the session."
