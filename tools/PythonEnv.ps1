<#
.SYNOPSIS
    Python discovery and installation helpers shared by the Mako Production System scripts.
.DESCRIPTION
    Dot-source this file from a script in the same folder:

        . (Join-Path $here "PythonEnv.ps1")

    Functions:
        Find-Python       path of the first real Python 3.8+ on this PC, or $null. Ignores the
                          Microsoft Store stub under \Microsoft\WindowsApps\, checks PATH, the
                          py launcher and the usual install folders (per-user, Program Files,
                          C:\Python3x, Anaconda and Miniconda).
        Install-Python    installs Python 3.12 for the current user from python.org, silently
                          (InstallAllUsers=0 PrependPath=1, no admin rights), with winget as
                          fallback, then refreshes PATH for the running session.
        Test-PyModule     $true when the given interpreter can import the module.
        Ensure-PyModules  pip installs each missing module, retrying with --user, and throws
                          with a manual command when a module still cannot be imported.
        Resolve-Python    Find-Python, then Install-Python unless -NoInstall, then Find-Python
                          again. Returns the interpreter path or throws.

    These functions run native executables (python, pip, winget). Windows PowerShell 5.1 turns
    native stderr output into terminating errors when $ErrorActionPreference is Stop, so each
    function sets Continue for its own scope and checks $LASTEXITCODE explicitly. Callers should
    do the same for the native commands they run themselves.
.NOTES
    Windows PowerShell 5.1 and PowerShell 7 compatible. Save as UTF-8 with BOM, CRLF.
#>

function Test-RealPython {
    param([string]$Exe)
    $ErrorActionPreference = "Continue"
    if ([string]::IsNullOrWhiteSpace($Exe)) { return $false }
    if ($Exe -like "*\Microsoft\WindowsApps\*") { return $false }   # Store stub: prints a message and exits 9009
    if (-not (Test-Path -LiteralPath $Exe)) { return $false }
    $v = $null
    try { $v = & $Exe -c "import sys; print(sys.version_info[0]*100 + sys.version_info[1])" 2>$null }
    catch { return $false }
    if ($LASTEXITCODE -ne 0) { return $false }
    $n = 0
    if (-not [int]::TryParse(("$v").Trim(), [ref]$n)) { return $false }
    return ($n -ge 308)
}

function Find-Python {
    $ErrorActionPreference = "Continue"
    $candidates = New-Object System.Collections.Generic.List[string]
    foreach ($name in @("python.exe", "python3.exe")) {
        Get-Command $name -All -ErrorAction SilentlyContinue | ForEach-Object { $candidates.Add($_.Source) }
    }
    $pyl = Get-Command py.exe -ErrorAction SilentlyContinue
    if ($pyl -and $pyl.Source -notlike "*\Microsoft\WindowsApps\*") {
        $resolved = $null
        try { $resolved = & $pyl.Source -3 -c "import sys; print(sys.executable)" 2>$null } catch { }
        if ($LASTEXITCODE -eq 0 -and $resolved) { $candidates.Add(("$resolved").Trim()) }
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
    param([string]$Version = "3.12.10")   # last 3.12 release with a Windows installer
    $ErrorActionPreference = "Continue"
    $arch = if ($env:PROCESSOR_ARCHITECTURE -eq "ARM64") { "arm64" } else { "amd64" }
    $url = "https://www.python.org/ftp/python/$Version/python-$Version-$arch.exe"
    $installer = Join-Path $env:TEMP "python-$Version-$arch.exe"
    Write-Host "Python 3 not found. Installing Python $Version for the current user (no admin rights, about 100 MB, 1-2 minutes) ..."
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $url -OutFile $installer -UseBasicParsing -ErrorAction Stop
        $p = Start-Process -FilePath $installer -ArgumentList "/quiet InstallAllUsers=0 PrependPath=1 Include_launcher=1 Include_test=0 Include_doc=0" -Wait -PassThru
        # 0 = ok, 3010 = ok but a restart is pending
        if ($p.ExitCode -ne 0 -and $p.ExitCode -ne 3010) { throw "installer exit code $($p.ExitCode)" }
        Write-Host "Python installed."
    } catch {
        Write-Warning "python.org install failed: $($_.Exception.Message). Trying winget ..."
        $wg = Get-Command winget.exe -ErrorAction SilentlyContinue
        if (-not $wg) { throw "Neither python.org nor winget is available. Install Python 3 manually from https://www.python.org/downloads/windows/ (tick 'Add python.exe to PATH'), then re-run this script." }
        & $wg.Source install --id Python.Python.3.12 -e --scope user --silent --accept-package-agreements --accept-source-agreements 2>&1 | ForEach-Object { Write-Host "  $_" }
        if ($LASTEXITCODE -ne 0) { throw "winget install failed with exit code $LASTEXITCODE. Install Python 3 manually from https://www.python.org/downloads/windows/ and re-run." }
    }
    # refresh PATH for this session so the new interpreter is visible without opening a new window
    $env:Path = [Environment]::GetEnvironmentVariable("Path", "User") + ";" + [Environment]::GetEnvironmentVariable("Path", "Machine")
}

function Test-PyModule {
    param([string]$Python, [string]$Module)
    $ErrorActionPreference = "Continue"
    if ([string]::IsNullOrWhiteSpace($Python)) { return $false }
    # find_spec writes nothing to stdout or stderr; Out-Null keeps the pipeline clean regardless
    & $Python -c "import importlib.util, sys; sys.exit(0 if importlib.util.find_spec('$Module') else 1)" 2>$null | Out-Null
    return ($LASTEXITCODE -eq 0)
}

function Ensure-PyModules {
    param([string]$Python, [string[]]$Modules)
    $ErrorActionPreference = "Continue"
    foreach ($m in $Modules) {
        if (Test-PyModule $Python $m) { continue }
        Write-Host "Installing Python package $m ..."
        & $Python -m pip install --quiet --disable-pip-version-check $m 2>&1 | ForEach-Object { Write-Host "  $_" }
        if (-not (Test-PyModule $Python $m)) {
            & $Python -m pip install --quiet --disable-pip-version-check --user $m 2>&1 | ForEach-Object { Write-Host "  $_" }
        }
        if (-not (Test-PyModule $Python $m)) {
            throw "Could not install Python package '$m'. Install it manually:  `"$Python`" -m pip install $m"
        }
    }
}

function Resolve-Python {
    param([switch]$NoInstall)
    $ErrorActionPreference = "Continue"
    $py = Find-Python
    if ($py) { return $py }
    if ($NoInstall) { throw "Python 3 not found and -NoInstall was given. Install Python 3 from https://www.python.org/downloads/windows/ (tick 'Add python.exe to PATH') and re-run." }
    Install-Python | Out-Null
    $py = Find-Python
    if (-not $py) { throw "Python was installed but could not be located. Open a new PowerShell window and re-run this script." }
    return $py
}
