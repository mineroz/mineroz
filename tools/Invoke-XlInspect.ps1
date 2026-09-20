<#
.SYNOPSIS
    Runs xl_inspect.py against the PMC operation report workbooks.
.DESCRIPTION
    Finds a real Python 3 (ignoring the Microsoft Store stub). If none exists,
    installs Python 3.12 for the current user from python.org (no admin rights),
    falling back to winget. Installs openpyxl and oletools if missing, then writes
    inspection_report.md and one JSON per workbook to the output folder.
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

# ---------------------------------------------------------------- Python discovery
function Test-RealPython([string]$Exe) {
    if ([string]::IsNullOrWhiteSpace($Exe)) { return $false }
    if ($Exe -like "*\Microsoft\WindowsApps\*") { return $false }   # Store stub, prints a message and exits 9009
    if (-not (Test-Path -LiteralPath $Exe)) { return $false }
    $v = & $Exe -c "import sys; print(sys.version_info[0]*100 + sys.version_info[1])" 2>$null
    return ($LASTEXITCODE -eq 0 -and [int]"$v" -ge 308)
}

function Find-Python {
    $candidates = New-Object System.Collections.Generic.List[string]
    foreach ($name in @("python.exe", "python3.exe")) {
        Get-Command $name -All -ErrorAction SilentlyContinue | ForEach-Object { $candidates.Add($_.Source) }
    }
    $pyl = Get-Command py.exe -ErrorAction SilentlyContinue
    if ($pyl -and $pyl.Source -notlike "*\Microsoft\WindowsApps\*") {
        $resolved = & $pyl.Source -3 -c "import sys; print(sys.executable)" 2>$null
        if ($LASTEXITCODE -eq 0 -and $resolved) { $candidates.Add("$resolved".Trim()) }
    }
    $globs = @(
        "$env:LOCALAPPDATA\Programs\Python\Python3*\python.exe",
        "$env:ProgramFiles\Python3*\python.exe",
        "${env:ProgramFiles(x86)}\Python3*\python.exe",
        "C:\Python3*\python.exe",
        "$env:LOCALAPPDATA\anaconda3\python.exe",
        "$env:USERPROFILE\anaconda3\python.exe",
        "$env:USERPROFILE\miniconda3\python.exe",
        "$env:ProgramData\anaconda3\python.exe"
    )
    foreach ($g in $globs) {
        Get-ChildItem -Path $g -ErrorAction SilentlyContinue | Sort-Object FullName -Descending | ForEach-Object { $candidates.Add($_.FullName) }
    }
    foreach ($c in $candidates) { if (Test-RealPython $c) { return $c } }
    return $null
}

function Install-Python {
    $ver = "3.12.10"
    $arch = if ($env:PROCESSOR_ARCHITECTURE -eq "ARM64") { "arm64" } else { "amd64" }
    $url = "https://www.python.org/ftp/python/$ver/python-$ver-$arch.exe"
    $installer = Join-Path $env:TEMP "python-$ver-$arch.exe"
    Write-Host "Python 3 not found. Installing Python $ver for the current user (no admin rights, ~100 MB, 1-2 minutes) ..."
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $url -OutFile $installer -UseBasicParsing
        $p = Start-Process -FilePath $installer -ArgumentList "/quiet InstallAllUsers=0 PrependPath=1 Include_launcher=1 Include_test=0 Include_doc=0" -Wait -PassThru
        if ($p.ExitCode -ne 0) { throw "installer exit code $($p.ExitCode)" }
        Write-Host "Python installed."
    } catch {
        Write-Warning "python.org install failed: $($_.Exception.Message). Trying winget ..."
        $wg = Get-Command winget.exe -ErrorAction SilentlyContinue
        if (-not $wg) { throw "Neither python.org nor winget is available. Install Python 3 manually from https://www.python.org/downloads/windows/ (tick 'Add python.exe to PATH'), then re-run this script." }
        & $wg.Source install --id Python.Python.3.12 -e --scope user --silent --accept-package-agreements --accept-source-agreements
        if ($LASTEXITCODE -ne 0) { throw "winget install failed with exit code $LASTEXITCODE. Install Python 3 manually from https://www.python.org/downloads/windows/ and re-run." }
    }
    # refresh PATH for this session so the new interpreter is visible
    $env:Path = [Environment]::GetEnvironmentVariable("Path", "User") + ";" + [Environment]::GetEnvironmentVariable("Path", "Machine")
}

$py = Find-Python
if (-not $py) {
    if ($NoInstall) { throw "Python 3 not found and -NoInstall was given." }
    Install-Python
    $py = Find-Python
    if (-not $py) { throw "Python was installed but could not be located. Open a new PowerShell window and re-run this script." }
}
Write-Host "Using Python: $py"

# ---------------------------------------------------------------- packages
function Test-PyModule([string]$Module) {
    # find_spec writes nothing to stderr, so this is safe on PowerShell 5.1
    & $py -c "import importlib.util, sys; sys.exit(0 if importlib.util.find_spec('$Module') else 1)"
    return ($LASTEXITCODE -eq 0)
}
foreach ($m in @("openpyxl", "oletools")) {
    if (-not (Test-PyModule $m)) {
        Write-Host "Installing $m ..."
        & $py -m pip install --quiet --disable-pip-version-check $m
        if (-not (Test-PyModule $m)) {
            & $py -m pip install --quiet --disable-pip-version-check --user $m
        }
        if (-not (Test-PyModule $m)) { throw "Could not install Python package '$m'. Install it manually:  `"$py`" -m pip install $m" }
    }
}

# ---------------------------------------------------------------- run
Write-Host "Inspecting workbooks in $Source"
& $py $script $Source -o $Out
if ($LASTEXITCODE -ne 0) { throw "xl_inspect.py failed with exit code $LASTEXITCODE" }

$report = Join-Path $Out "inspection_report.md"
Write-Host ""
Write-Host "Report: $report"
Write-Host "Attach the whole '$Out' folder (report + JSON files) to the session."
