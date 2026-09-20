<#
.SYNOPSIS
    Builds the Mako Production System workbooks (blank and DEMO) from the schema.
.DESCRIPTION
    Finds or installs Python 3 (PythonEnv.ps1), makes sure openpyxl is available, then runs
    mps\build_workbook.py twice:

        <Root>\04_MPS\MPS_<Year>.xlsx        blank workbook for production use
        <Root>\04_MPS\MPS_<Year>_DEMO.xlsx   the same workbook with <SampleDays> days of synthetic DEMO data

    Existing files are overwritten, so close them in Excel before re-running. The generator is
    looked up next to this script first (<Root>\01_Tools\mps\build_workbook.py in the C:\MakoPS
    layout, <repo>\mps\build_workbook.py when this script runs from the repository), then under
    <Root>\01_Tools\mps; the path used is printed.
.PARAMETER Root
    Working folder root. Default C:\MakoPS. Output goes to <Root>\04_MPS.
.PARAMETER Year
    Reporting year of the workbook. Default 2026.
.PARAMETER SampleDays
    Days of synthetic data in the DEMO workbook. Default 74. 0 skips the DEMO workbook.
.PARAMETER NoInstall
    Never install Python; fail with instructions instead.
.PARAMETER NoOpen
    Do not open the output folder in Explorer when done.
.EXAMPLE
    .\Build-MPS.ps1
    .\Build-MPS.ps1 -Root D:\MakoPS -Year 2027 -SampleDays 30
#>
[CmdletBinding()]
param(
    [string]$Root = "C:\MakoPS",
    [int]$Year = 2026,
    [int]$SampleDays = 74,
    [switch]$NoInstall,
    [switch]$NoOpen
)

# Continue, not Stop: Windows PowerShell 5.1 turns native stderr output into terminating
# errors under Stop, and pip / openpyxl legitimately write warnings to stderr.
$ErrorActionPreference = "Continue"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
# Absolute root: the .NET file test below resolves relative paths against the process directory, not the PowerShell location
$Root = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Root)

$envFile = Join-Path $here "PythonEnv.ps1"
if (-not (Test-Path -LiteralPath $envFile)) { throw "PythonEnv.ps1 not found next to this script: $envFile" }
. $envFile

# ---------------------------------------------------------------- locate the generator
# The copies next to this script win over a deployed copy under <Root>, so a run from the repository
# builds with the repository schema even when C:\MakoPS exists.
$candidates = @(
    (Join-Path $here "mps\build_workbook.py"),
    (Join-Path (Split-Path -Parent $here) "mps\build_workbook.py"),
    (Join-Path $Root "01_Tools\mps\build_workbook.py")
)
$generator = $null
foreach ($c in $candidates) { if (Test-Path -LiteralPath $c) { $generator = $c; break } }
if (-not $generator) { throw ("build_workbook.py not found. Looked in:`n  " + ($candidates -join "`n  ")) }
$schema = Join-Path (Split-Path -Parent $generator) "schema\mps_schema.json"
if (-not (Test-Path -LiteralPath $schema)) { throw "Schema not found next to the generator: $schema" }
Write-Host "Using generator: $generator"
Write-Host "Using schema:    $schema"

$outDir = Join-Path $Root "04_MPS"
if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Force -Path $outDir | Out-Null }

# ---------------------------------------------------------------- python and packages
$py = Resolve-Python -NoInstall:$NoInstall
Write-Host "Using Python: $py"
Ensure-PyModules $py @("openpyxl")

# ---------------------------------------------------------------- build
$blank = Join-Path $outDir ("MPS_{0}.xlsx" -f $Year)
$demo  = Join-Path $outDir ("MPS_{0}_DEMO.xlsx" -f $Year)

function Invoke-Generator([string[]]$Arguments, [string]$Target) {
    if (Test-Path -LiteralPath $Target) {
        # a workbook that is open in Excel cannot be overwritten; fail before spending time on the build
        try { $fs = [System.IO.File]::Open($Target, "Open", "ReadWrite", "None"); $fs.Close() }
        catch { throw "Cannot overwrite $Target. Close it in Excel and re-run." }
    }
    Write-Host ("Building {0} ..." -f (Split-Path -Leaf $Target))
    & $py $generator @Arguments
    if ($LASTEXITCODE -ne 0) { throw "build_workbook.py failed with exit code $LASTEXITCODE" }
    if (-not (Test-Path -LiteralPath $Target)) { throw "build_workbook.py finished but $Target was not written" }
}

Invoke-Generator @("--out", $blank, "--year", "$Year") $blank
if ($SampleDays -gt 0) {
    Invoke-Generator @("--out", $demo, "--year", "$Year", "--sample-days", "$SampleDays") $demo
}

# ---------------------------------------------------------------- report
Write-Host ""
Write-Host "MPS workbooks written to $outDir"
Write-Host ("  {0}  ({1:N1} MB)  blank, for production use" -f $blank, ((Get-Item -LiteralPath $blank).Length / 1MB))
if ($SampleDays -gt 0) {
    Write-Host ("  {0}  ({1:N1} MB)  {2} days of synthetic DEMO data" -f $demo, ((Get-Item -LiteralPath $demo).Length / 1MB), $SampleDays)
}
Write-Host ""
Write-Host "Open the DEMO workbook first to see the Dashboard, Daily_Report and Monthly_Summary populated."
Write-Host "Enter real data in the blank workbook only. Read the README sheet inside the workbook before the first entry."
if (-not $NoOpen) { try { Start-Process explorer.exe $outDir } catch { } }
