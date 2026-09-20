<#
.SYNOPSIS
    Python discovery and installation helpers shared by the Mako Production System scripts.
.DESCRIPTION
    Dot-source this file from a script in the same folder:

        . (Join-Path $here "PythonEnv.ps1")

    Functions:
        Find-Python       path of the first real Python 3.8+ on PATH or behind the py launcher,
                          otherwise the newest one in the usual install folders (per-user,
                          Program Files, C:\Python3x, Anaconda and Miniconda), or $null. Every
                          candidate is probed for its version; the Microsoft Store stub under
                          \Microsoft\WindowsApps\ fails that probe (exit code 9009) and is
                          skipped, while a real Store Python at the same path is accepted.
        Install-Python    installs Python 3.12 for the current user from python.org, silently
                          (InstallAllUsers=0 InstallLauncherAllUsers=0 PrependPath=1, no admin
                          rights, through the Windows proxy when one is configured), with winget
                          as fallback, then refreshes PATH for the running session.
        Test-PyModule     $true when the given interpreter can import the module.
        Ensure-PyModules  pip installs each missing module (passing the Windows proxy when one
                          applies), retrying with --user, and throws with a manual command when
                          a module still cannot be imported.
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
    if (-not (Test-Path -LiteralPath $Exe)) { return $false }
    # The Microsoft Store stub (python.exe under \Microsoft\WindowsApps\) prints a message and exits 9009 when
    # given arguments, so the exit code test below rejects it; a real Store Python at the same path passes.
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
    if ($pyl) {
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
        # newest first by the file version resource, not by name (a name sort puts Python39 above Python312)
        Get-ChildItem -Path $g -ErrorAction SilentlyContinue | Sort-Object { $_.VersionInfo.FileVersionRaw } -Descending | ForEach-Object { $candidates.Add($_.FullName) }
    }
    foreach ($c in $candidates) { if (Test-RealPython $c) { return $c } }
    return $null
}

function Use-SystemProxy {
    # Windows PowerShell 5.1 uses the proxy from Internet Options but sends it no credentials, so an
    # authenticating site proxy answers 407. Attach the current user's credentials to the default proxy.
    try {
        $proxy = [System.Net.WebRequest]::GetSystemWebProxy()
        $proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
        [System.Net.WebRequest]::DefaultWebProxy = $proxy
    } catch { }
}

function Get-PipProxyArgs {
    # pip reads a static proxy from Internet Options but not a proxy script (PAC). Resolve the proxy
    # Windows would use for PyPI and return it as pip arguments, or an empty array for a direct connection.
    try {
        $target = [uri]"https://pypi.org/"
        $p = [System.Net.WebRequest]::GetSystemWebProxy().GetProxy($target)
        if ($p -and $p.Host -ne $target.Host) { return @("--proxy", $p.AbsoluteUri) }
    } catch { }
    return @()
}

function Install-Python {
    param([string]$Version = "3.12.10")   # last 3.12 release with a Windows installer
    $ErrorActionPreference = "Continue"
    $ProgressPreference = "SilentlyContinue"   # the 5.1 progress bar makes Invoke-WebRequest many times slower
    $arch = if ($env:PROCESSOR_ARCHITECTURE -eq "ARM64") { "arm64" } else { "amd64" }
    $url = "https://www.python.org/ftp/python/$Version/python-$Version-$arch.exe"
    $installer = Join-Path $env:TEMP "python-$Version-$arch.exe"
    Write-Host "Python 3 not found. Installing Python $Version for the current user (no admin rights, 25 MB download, about 100 MB installed, 1-2 minutes) ..."
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Use-SystemProxy
        Invoke-WebRequest -Uri $url -OutFile $installer -UseBasicParsing -ErrorAction Stop
        # InstallLauncherAllUsers=0: the py launcher defaults to a per-machine component, which asks for
        # elevation even in a per-user /quiet install.
        $p = Start-Process -FilePath $installer -ArgumentList "/quiet InstallAllUsers=0 InstallLauncherAllUsers=0 PrependPath=1 Include_launcher=1 Include_test=0 Include_doc=0" -Wait -PassThru
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
    # refresh PATH for this session so the new interpreter is visible without opening a new window;
    # Machine then User as in a fresh window, and the entries this process already had are kept
    $env:Path = [Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [Environment]::GetEnvironmentVariable("Path", "User") + ";" + $env:Path
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
    $proxyArgs = @(Get-PipProxyArgs)
    foreach ($m in $Modules) {
        if (Test-PyModule $Python $m) { continue }
        Write-Host "Installing Python package $m ..."
        if ($proxyArgs.Count -gt 0) { Write-Host "  via proxy $($proxyArgs[1])" }
        & $Python -m pip install --quiet --disable-pip-version-check @proxyArgs $m 2>&1 | ForEach-Object { Write-Host "  $_" }
        if (-not (Test-PyModule $Python $m)) {
            & $Python -m pip install --quiet --disable-pip-version-check --user @proxyArgs $m 2>&1 | ForEach-Object { Write-Host "  $_" }
        }
        if (-not (Test-PyModule $Python $m)) {
            $manual = (@("`"$Python`"", "-m", "pip", "install") + $proxyArgs + @($m)) -join " "
            throw "Could not install Python package '$m'. Install it manually:  $manual"
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
