<#
.SYNOPSIS
    Runs xl_inspect.py against the PMC operation report workbooks.
.DESCRIPTION
    Finds a real Python 3 (ignoring the Microsoft Store stub). If none exists,
    installs Python 3.12 for the current user from python.org (no admin rights),
    falling back to winget. Installs openpyxl and oletools if missing, then writes
    inspection_report.md and one JSON per workbook to the output folder.
    Python discovery and installation live in PythonEnv.ps1 next to this script.
.PARAMETER Source
    Folder holding the workbooks. Defaults to <root>\02_Source\2026 when it exists,
    otherwise the 2026 operation report folder on X:.
.PARAMETER Out
    Output folder for the report. Defaults to <root>\03_Inspection in the C:\MakoPS
    layout, otherwise <Source>\_inspection.
.PARAMETER NoInstall
    Never install Python; fail with instructions instead.
.EXAMPLE
    .\Invoke-XlInspect.ps1
    .\Invoke-XlInspect.ps1 -Source "X:\13_Management\01_GM_Dashboard\01_Operation_Report\01_Excel\2026"
.NOTES
    Mapped drives such as X: are not visible in an elevated (Run as administrator)
    PowerShell session. Run this from a normal PowerShell window.
#>
[CmdletBinding()]
param(
    [string]$Source = "",
    [string]$Out = "",
    [switch]$NoInstall
)

# Continue, not Stop: Windows PowerShell 5.1 turns native stderr output into terminating
# errors under Stop, and pip / openpyxl legitimately write warnings to stderr.
$ErrorActionPreference = "Continue"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$script = Join-Path $here "xl_inspect.py"
$root = Split-Path -Parent $here

$envFile = Join-Path $here "PythonEnv.ps1"
if (-not (Test-Path -LiteralPath $envFile)) { throw "PythonEnv.ps1 not found next to this script: $envFile" }
. $envFile
if (-not (Test-Path -LiteralPath $script)) { throw "xl_inspect.py not found next to this script: $script" }

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

# ---------------------------------------------------------------- python and packages
$py = Resolve-Python -NoInstall:$NoInstall
Write-Host "Using Python: $py"
Ensure-PyModules $py @("openpyxl", "oletools")

# ---------------------------------------------------------------- run
Write-Host "Inspecting workbooks in $Source"
& $py $script $Source -o $Out
if ($LASTEXITCODE -ne 0) { throw "xl_inspect.py failed with exit code $LASTEXITCODE" }

$report = Join-Path $Out "inspection_report.md"
Write-Host ""
Write-Host "Report: $report"
Write-Host "Attach the whole '$Out' folder (report + JSON files) to the session."
