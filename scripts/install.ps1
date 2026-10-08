<#
.SYNOPSIS
    Installs and registers the Move2Folder Outlook Classic Add-In for the current user.
    Requires NO administrator privileges.

.DESCRIPTION
    Checks all prerequisites (.NET 8 Desktop Runtime x64, 64-bit Outlook),
    copies binaries to %LOCALAPPDATA%\Move2Folder, configures COM and Outlook
    registry entries, checks Outlook resiliency/disabled items, and optionally
    collects detailed debug/troubleshooting logs.

.PARAMETER Debug
    Runs in diagnostic/debug mode: collects detailed environment info, registry state,
    Outlook event logs, and generates a debug report file to locate problems.

.PARAMETER Force
    Forces non-interactive operations (e.g. automatically closes Outlook without prompting).
#>

param(
    [Alias("d")]
    [switch]$Debug,

    [Alias("f")]
    [switch]$Force,

    [Alias("h", "?")]
    [switch]$Help
)

$ErrorActionPreference = "Stop"

# Handle command-line arguments like --debug, -debug, /debug, --force, etc.
foreach ($arg in $args) {
    if ($arg -match "^[-/]{1,2}(debug|d)$") { $Debug = $true }
    if ($arg -match "^[-/]{1,2}(force|f)$") { $Force = $true }
    if ($arg -match "^[-/]{1,2}(help|h|\?)$") { $Help = $true }
}

if ($Help) {
    Write-Host @"
Move2Folder Installation Script
Usage:
    install.cmd [options]
    .\install.ps1 [options]

Options:
    --debug, -Debug    Run in debug mode: collects full system diagnostics,
                       registry status, and Outlook event logs for troubleshooting.
    --force, -Force    Force installation: closes running Outlook automatically.
    --help,  -Help     Show this help message.

Example:
    .\install.ps1 --debug
"@
    exit 0
}

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$TargetDir = Join-Path $env:LOCALAPPDATA "Move2Folder"
$Clsid = "{298621CD-332D-4DE0-B3B8-43BA73CA4893}"
$ProgId = "Move2Folder.Addin"

Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "     Move2Folder - Outlook Add-In Installer       " -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
if ($Debug) {
    Write-Host "[MODE] Running with --debug diagnostics enabled`n" -ForegroundColor Magenta
} else {
    Write-Host "Target: $TargetDir`n" -ForegroundColor Gray
}

# ---------------------------------------------------------------------------
# Helper: Find Source Directory
# ---------------------------------------------------------------------------
function Get-SourceDirectory {
    # 1. Same directory as script (distribution package)
    if (Test-Path (Join-Path $ScriptDir "Move2Folder.comhost.dll")) {
        return $ScriptDir
    }
    # 2. Build output directory in repo
    $binDir = Join-Path $ScriptDir "..\src\Move2Folder\bin\x64\Release\net8.0-windows"
    if (Test-Path (Join-Path $binDir "Move2Folder.comhost.dll")) {
        return (Resolve-Path $binDir).Path
    }
    # 3. Target directory itself if already installed
    if (Test-Path (Join-Path $TargetDir "Move2Folder.comhost.dll")) {
        return $TargetDir
    }
    return $null
}

# ---------------------------------------------------------------------------
# Helper: Detect .NET 8 Desktop Runtime (x64)
# ---------------------------------------------------------------------------
function Test-DotNetRuntime {
    $info = [ordered]@{
        Installed = $false
        Version   = "Not installed"
        Details   = ""
        Only32Bit = $false
    }

    # Method A: dotnet CLI
    try {
        $runtimes = dotnet --list-runtimes 2>$null
        $matches = $runtimes | Where-Object { $_ -match "^Microsoft\.WindowsDesktop\.App\s+8\." }
        if ($matches) {
            $info.Installed = $true
            $info.Details = ($matches -join "; ")
            $first = $matches | Select-Object -First 1
            if ($first -match "Microsoft\.WindowsDesktop\.App\s+([0-9\.]+)") {
                $info.Version = $Matches[1]
            }
            return [PSCustomObject]$info
        }
    } catch {}

    # Method B: 64-bit Filesystem check
    $shared64 = Join-Path $env:ProgramFiles "dotnet\shared\Microsoft.WindowsDesktop.App"
    if (Test-Path $shared64) {
        $vDirs = Get-ChildItem -Path $shared64 -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "8.*" }
        if ($vDirs) {
            $info.Installed = $true
            $info.Version = ($vDirs.Name -join ", ")
            $info.Details = "Found in $shared64"
            return [PSCustomObject]$info
        }
    }

    # Method C: Check if 32-bit only is installed
    $p86 = ${env:ProgramFiles(x86)}
    if ($p86) {
        $shared32 = Join-Path $p86 "dotnet\shared\Microsoft.WindowsDesktop.App"
        if (Test-Path $shared32) {
            $vDirs32 = Get-ChildItem -Path $shared32 -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "8.*" }
            if ($vDirs32) {
                $info.Only32Bit = $true
                $info.Version = "x86 only (" + ($vDirs32.Name -join ", ") + ")"
                $info.Details = "32-bit runtime found in $shared32, but 64-bit runtime is required!"
                return [PSCustomObject]$info
            }
        }
    }

    return [PSCustomObject]$info
}

# ---------------------------------------------------------------------------
# Helper: Detect Outlook Installation & Bitness
# ---------------------------------------------------------------------------
function Get-OutlookInfo {
    $info = [ordered]@{
        Installed = $false
        Bitness   = "Unknown"
        Version   = "Unknown"
        Path      = "Unknown"
        IsRunning = $false
        InstallType = "Unknown"
    }

    # 1. Process check
    $proc = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($proc) {
        $info.IsRunning = $true
        try {
            $info.Path = $proc.MainModule.FileName
        } catch {}
    }

    # 2. App Paths
    $appPaths = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE' -ErrorAction SilentlyContinue
    if ($appPaths -and $appPaths.'(default)') {
        $info.Installed = $true
        if ($info.Path -eq "Unknown") {
            $info.Path = $appPaths.'(default)'
        }
    }

    # 3. Click-to-Run configuration
    $c2r = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
    if ($c2r) {
        $info.Installed = $true
        $info.InstallType = "Click-to-Run"
        if ($c2r.VersionToReport) { $info.Version = $c2r.VersionToReport }
        if ($c2r.Platform) { $info.Bitness = $c2r.Platform }
    }

    # 4. MSI / Traditional registry keys
    if ($info.Bitness -eq "Unknown") {
        foreach ($ver in @("16.0", "15.0", "14.0")) {
            $key = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\$ver\Outlook" -ErrorAction SilentlyContinue
            if ($key) {
                $info.Installed = $true
                $info.InstallType = "MSI"
                if ($key.Bitness) { $info.Bitness = $key.Bitness }
                break
            }
            $key32 = Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\$ver\Outlook" -ErrorAction SilentlyContinue
            if ($key32) {
                $info.Installed = $true
                $info.InstallType = "MSI (32-bit)"
                $info.Bitness = "x86"
                break
            }
        }
    }

    # 5. Fallback bitness inspection from Path
    if ($info.Bitness -eq "Unknown" -and $info.Path -ne "Unknown") {
        $p86 = ${env:ProgramFiles(x86)}
        if ($p86 -and $info.Path.StartsWith($p86, [StringComparison]::OrdinalIgnoreCase)) {
            $info.Bitness = "x86"
        } elseif ($info.Path.StartsWith($env:ProgramFiles, [StringComparison]::OrdinalIgnoreCase)) {
            $info.Bitness = "x64"
        }
    }

    return [PSCustomObject]$info
}

# ---------------------------------------------------------------------------
# Helper: Decode Outlook DisabledItems in Resiliency Key
# ---------------------------------------------------------------------------
function Get-OutlookDisabledAddins {
    $disabledList = @()
    Get-ChildItem "HKCU:\Software\Microsoft\Office" -ErrorAction SilentlyContinue | ForEach-Object {
        $disabledKey = Join-Path $_.PSPath "Outlook\Resiliency\DisabledItems"
        if (Test-Path $disabledKey) {
            $ver = $_.PSChildName
            $props = Get-ItemProperty $disabledKey
            foreach ($p in $props.PSObject.Properties) {
                if ($p.TypeNameOfValue -eq "System.Byte[]") {
                    $decoded = [System.Text.Encoding]::Unicode.GetString($p.Value)
                    $isMove2Folder = $decoded -like "*move2folder*"
                    $disabledList += [PSCustomObject]@{
                        OfficeVersion = $ver
                        ItemName      = $p.Name
                        DecodedText   = $decoded
                        IsMove2Folder = $isMove2Folder
                        KeyPath       = $disabledKey
                    }
                }
            }
        }
    }
    return $disabledList
}

# ---------------------------------------------------------------------------
# Helper: Collect Outlook & Move2Folder Event Log entries
# ---------------------------------------------------------------------------
function Get-DiagnosticEventLogs {
    $events = @()

    # Outlook application events
    try {
        $evts = Get-WinEvent -FilterHashtable @{LogName='Application'; ProviderName='Outlook'} -MaxEvents 30 -ErrorAction SilentlyContinue
        if ($evts) { $events += $evts }
    } catch {}

    # Application Error events (Crashes)
    try {
        $evtsCrash = Get-WinEvent -FilterHashtable @{LogName='Application'; ProviderName='Application Error'} -MaxEvents 10 -ErrorAction SilentlyContinue
        if ($evtsCrash) {
            $events += ($evtsCrash | Where-Object { $_.Message -like "*OUTLOOK*" -or $_.Message -like "*Move2Folder*" })
        }
    } catch {}

    # .NET Runtime events
    try {
        $evtsNet = Get-WinEvent -FilterHashtable @{LogName='Application'; ProviderName='.NET Runtime'} -MaxEvents 10 -ErrorAction SilentlyContinue
        if ($evtsNet) {
            $events += ($evtsNet | Where-Object { $_.Message -like "*Move2Folder*" -or $_.Message -like "*OUTLOOK*" })
        }
    } catch {}

    return $events | Sort-Object TimeCreated -Descending
}

# ---------------------------------------------------------------------------
# Pre-Flight Checks
# ---------------------------------------------------------------------------
Write-Host "[1/4] Running prerequisite checks..." -ForegroundColor Yellow

# Check 1: 64-Bit OS
if (-not [Environment]::Is64BitOperatingSystem) {
    Write-Host "  [FAIL] 64-bit Windows operating system is required." -ForegroundColor Red
    Write-Host "         Move2Folder is an x64 COM Add-In and cannot run on 32-bit Windows." -ForegroundColor Red
    exit 1
} else {
    Write-Host "  [OK] 64-bit Windows operating system detected" -ForegroundColor Green
}

# Check 2: Outlook Installation & Bitness
$outlook = Get-OutlookInfo
if (-not $outlook.Installed -and $outlook.Path -eq "Unknown") {
    Write-Host "  [WARN] Microsoft Outlook was not found in standard registry locations." -ForegroundColor Yellow
    Write-Host "         The Add-In will still be registered, but make sure Outlook Classic is installed." -ForegroundColor Yellow
} else {
    Write-Host "  [OK] Outlook detected: Version $($outlook.Version) ($($outlook.Bitness))" -ForegroundColor Green
    if ($outlook.Bitness -eq "x86") {
        Write-Host "`n  [CRITICAL WARNING] Installed Outlook appears to be 32-bit (x86)!" -ForegroundColor Red
        Write-Host "  Move2Folder is built as a 64-bit (x64) COM Add-In." -ForegroundColor Red
        Write-Host "  64-bit DLLs CANNOT be loaded into a 32-bit Outlook process." -ForegroundColor Red
        Write-Host "  Please install 64-bit Microsoft Office / Outlook to use Move2Folder.`n" -ForegroundColor Red
        if (-not $Force) {
            $continue = Read-Host "Do you want to proceed anyway? (y/n)"
            if ($continue -notmatch "^[jyJY]") {
                exit 1
            }
        }
    }
}

# Check 3: .NET 8 Desktop Runtime (x64)
$dotnet = Test-DotNetRuntime
if ($dotnet.Installed) {
    Write-Host "  [OK] .NET 8 Desktop Runtime (x64) detected: $($dotnet.Version)" -ForegroundColor Green
} else {
    if ($dotnet.Only32Bit) {
        Write-Host "  [FAIL] Only 32-bit .NET 8 Desktop Runtime was found." -ForegroundColor Red
        Write-Host "         A 64-bit (x64) .NET 8 Desktop Runtime is required for Outlook x64!" -ForegroundColor Red
    } else {
        Write-Host "  [FAIL] .NET 8 Desktop Runtime (x64) is not installed." -ForegroundColor Red
    }
    Write-Host "`n  Required package: Microsoft .NET 8 Desktop Runtime (Windows x64)" -ForegroundColor Yellow
    Write-Host "  Download link:    https://aka.ms/dotnet/8.0/windowsdesktop-runtime-win-x64.exe" -ForegroundColor Cyan
    Write-Host "  Official page:    https://dotnet.microsoft.com/download/dotnet/8.0`n" -ForegroundColor Cyan

    if (-not $Force) {
        $openUrl = Read-Host "Would you like to open the .NET 8 download page now in your browser? (y/n)"
        if ($openUrl -match "^[jyJY]") {
            Start-Process "https://aka.ms/dotnet/8.0/windowsdesktop-runtime-win-x64.exe"
        }
        $proceed = Read-Host "Continue installation anyway? (y/n)"
        if ($proceed -notmatch "^[jyJY]") {
            Write-Host "Installation aborted. Please install .NET 8 Desktop Runtime (x64) first." -ForegroundColor Yellow
            exit 1
        }
    }
}

# Check 4: Check if Outlook is running and close if necessary
$outlookProc = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
if ($outlookProc) {
    Write-Host "  [NOTICE] Outlook is currently running and locking files." -ForegroundColor Yellow
    if ($Force) {
        Write-Host "  Closing Outlook automatically (--force)..." -ForegroundColor Yellow
        Stop-Process -Name "OUTLOOK" -Force
        Start-Sleep -Seconds 2
    } else {
        $ans = Read-Host "  Would you like to close Outlook automatically now? (y/n)"
        if ($ans -match "^[jyJY]") {
            Stop-Process -Name "OUTLOOK" -Force
            Start-Sleep -Seconds 2
        } else {
            Write-Host "  Please close Outlook before proceeding." -ForegroundColor Yellow
            Pause
        }
    }
}

# ---------------------------------------------------------------------------
# File Installation
# ---------------------------------------------------------------------------
Write-Host "`n[2/4] Installing Move2Folder binaries..." -ForegroundColor Yellow

$sourceDir = Get-SourceDirectory
if (-not $sourceDir) {
    Write-Host "  [FAIL] Source binary files could not be located." -ForegroundColor Red
    Write-Host "         Please run install.ps1 from the extracted setup package or repository." -ForegroundColor Red
    exit 1
}

if (-not (Test-Path $TargetDir)) {
    New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null
}

$requiredFiles = @(
    "Move2Folder.comhost.dll",
    "Move2Folder.dll",
    "Move2Folder.runtimeconfig.json",
    "Move2Folder.deps.json",
    "Office.dll",
    "Microsoft.Office.Interop.Outlook.dll"
)

foreach ($file in $requiredFiles) {
    $src = Join-Path $sourceDir $file
    if (-not (Test-Path $src)) {
        Write-Host "  [FAIL] Required file missing from source: $file" -ForegroundColor Red
        exit 1
    }
    $dst = Join-Path $TargetDir $file
    try {
        Copy-Item -Path $src -Destination $dst -Force
        Write-Host "  Copied: $file" -ForegroundColor Gray
    } catch {
        Write-Host "  [FAIL] Could not copy ${file}: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "         Ensure Outlook is completely closed and try again." -ForegroundColor Yellow
        exit 1
    }
}
Write-Host "  [OK] All binary files installed to $TargetDir" -ForegroundColor Green

# ---------------------------------------------------------------------------
# Registry Configuration & Resiliency Cleanup
# ---------------------------------------------------------------------------
Write-Host "`n[3/4] Configuring COM and Outlook Add-In registry keys..." -ForegroundColor Yellow

$comHostDll = Join-Path $TargetDir "Move2Folder.comhost.dll"

# 1. COM CLSID in HKCU\Software\Classes\CLSID\{...}
$clsidPath = "HKCU:\Software\Classes\CLSID\$Clsid"
$inprocPath = "$clsidPath\InprocServer32"
$progIdSubPath = "$clsidPath\ProgID"

if (-not (Test-Path $clsidPath)) { New-Item -Path $clsidPath -Force | Out-Null }
Set-ItemProperty -Path $clsidPath -Name "(default)" -Value $ProgId

if (-not (Test-Path $inprocPath)) { New-Item -Path $inprocPath -Force | Out-Null }
Set-ItemProperty -Path $inprocPath -Name "(default)" -Value $comHostDll
Set-ItemProperty -Path $inprocPath -Name "ThreadingModel" -Value "Both"

if (-not (Test-Path $progIdSubPath)) { New-Item -Path $progIdSubPath -Force | Out-Null }
Set-ItemProperty -Path $progIdSubPath -Name "(default)" -Value $ProgId

# 2. ProgID in HKCU\Software\Classes\Move2Folder.Addin
$progIdRoot = "HKCU:\Software\Classes\$ProgId"
$progIdClsid = "$progIdRoot\CLSID"

if (-not (Test-Path $progIdRoot)) { New-Item -Path $progIdRoot -Force | Out-Null }
Set-ItemProperty -Path $progIdRoot -Name "(default)" -Value $ProgId

if (-not (Test-Path $progIdClsid)) { New-Item -Path $progIdClsid -Force | Out-Null }
Set-ItemProperty -Path $progIdClsid -Name "(default)" -Value $Clsid

# 3. Outlook Add-In registration
$outlookAddinKey = "HKCU:\Software\Microsoft\Office\Outlook\Addins\$ProgId"
if (-not (Test-Path $outlookAddinKey)) { New-Item -Path $outlookAddinKey -Force | Out-Null }
Set-ItemProperty -Path $outlookAddinKey -Name "FriendlyName" -Value "Move2Folder"
Set-ItemProperty -Path $outlookAddinKey -Name "Description" -Value "Move2Folder Outlook Classic Add-In"
Set-ItemProperty -Path $outlookAddinKey -Name "LoadBehavior" -Value 3 -Type DWord
Set-ItemProperty -Path $outlookAddinKey -Name "CommandLineSafe" -Value 0 -Type DWord

# 4. Check and clean up Outlook Resiliency (DisabledItems / CrashingAddinList)
$disabledItems = Get-OutlookDisabledAddins
foreach ($item in $disabledItems) {
    if ($item.IsMove2Folder) {
        Write-Host "  [FIX] Found Move2Folder in Outlook DisabledItems ($($item.OfficeVersion)). Removing block..." -ForegroundColor Yellow
        Remove-ItemProperty -Path $item.KeyPath -Name $item.ItemName -Force -ErrorAction SilentlyContinue
    }
}

# 5. Add to DoNotDisableAddinList across Office versions
Get-ChildItem "HKCU:\Software\Microsoft\Office" -ErrorAction SilentlyContinue | ForEach-Object {
    if ($_.PSChildName -match "^[0-9]+\.[0-9]+$") {
        $doNotDisablePath = Join-Path $_.PSPath "Outlook\Resiliency\DoNotDisableAddinList"
        if (-not (Test-Path $doNotDisablePath)) {
            New-Item -Path $doNotDisablePath -Force -ErrorAction SilentlyContinue | Out-Null
        }
        if (Test-Path $doNotDisablePath) {
            Set-ItemProperty -Path $doNotDisablePath -Name $ProgId -Value 1 -Type DWord -ErrorAction SilentlyContinue
        }
    }
}

Write-Host "  [OK] COM server registered: $comHostDll" -ForegroundColor Green
Write-Host "  [OK] Outlook Add-in registered with LoadBehavior=3 (load on startup)" -ForegroundColor Green
Write-Host "  [OK] Outlook Resiliency protection configured (DoNotDisableAddinList=1)" -ForegroundColor Green

# ---------------------------------------------------------------------------
# Diagnostics & Debug Report (Triggered with --debug or on demand)
# ---------------------------------------------------------------------------
Write-Host "`n[4/4] Verifying installation..." -ForegroundColor Yellow

$verifyCom = (Get-ItemProperty $inprocPath -ErrorAction SilentlyContinue).'(default)'
$verifyAddin = (Get-ItemProperty $outlookAddinKey -ErrorAction SilentlyContinue).LoadBehavior

if ($verifyCom -eq $comHostDll -and $verifyAddin -eq 3) {
    Write-Host "  [OK] Verification succeeded! Move2Folder is properly registered." -ForegroundColor Green
} else {
    Write-Host "  [WARN] Verification warning: Check registry permissions." -ForegroundColor Yellow
}

if ($Debug) {
    Write-Host "`n=== Generating Comprehensive Debug Report ===" -ForegroundColor Magenta
    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $debugLogPath = Join-Path $TargetDir "move2folder-debug-$timestamp.log"

    $report = New-Object System.Text.StringBuilder
    $null = $report.AppendLine("================================================================================")
    $null = $report.AppendLine("               Move2Folder Outlook Add-In Diagnostic Report                     ")
    $null = $report.AppendLine("               Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')             ")
    $null = $report.AppendLine("================================================================================")
    $null = $report.AppendLine("")

    # OS & User Info
    $null = $report.AppendLine("[1. SYSTEM ENVIRONMENT]")
    $os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    $null = $report.AppendLine("OS Caption:            $($os.Caption)")
    $null = $report.AppendLine("OS Version / Build:    $($os.Version) (Build $($os.BuildNumber))")
    $null = $report.AppendLine("OS Architecture:       $($os.OSArchitecture)")
    $null = $report.AppendLine("PowerShell Version:    $($PSVersionTable.PSVersion)")
    $null = $report.AppendLine("User Profile:          $env:USERPROFILE")
    $null = $report.AppendLine("LocalAppData:          $env:LOCALAPPDATA")
    $null = $report.AppendLine("Elevation / Admin:     $([bool](([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)))")
    $null = $report.AppendLine("")

    # Outlook Info
    $null = $report.AppendLine("[2. OUTLOOK INSTALLATION]")
    $null = $report.AppendLine("Detected Path:         $($outlook.Path)")
    $null = $report.AppendLine("Install Type:          $($outlook.InstallType)")
    $null = $report.AppendLine("Reported Version:      $($outlook.Version)")
    $null = $report.AppendLine("Architecture (Bitness):$($outlook.Bitness)")
    $null = $report.AppendLine("Process Running Now:   $($outlook.IsRunning)")
    if ($outlook.IsRunning) {
        $procs = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
        foreach ($p in $procs) {
            $null = $report.AppendLine("  PID: $($p.Id), WorkingSet: $([math]::Round($p.WorkingSet64 / 1MB, 2)) MB")
        }
    }
    $null = $report.AppendLine("")

    # .NET Runtimes
    $null = $report.AppendLine("[3. .NET RUNTIMES]")
    $null = $report.AppendLine(".NET 8 Status:         $($dotnet.Installed)")
    $null = $report.AppendLine(".NET 8 Details:        $($dotnet.Details)")
    try {
        $allRuntimes = dotnet --list-runtimes 2>$null
        $null = $report.AppendLine("All Installed Runtimes:")
        foreach ($r in $allRuntimes) {
            $null = $report.AppendLine("  $r")
        }
    } catch {
        $null = $report.AppendLine("  (dotnet CLI not available to list all runtimes)")
    }
    $null = $report.AppendLine("")

    # Files on disk
    $null = $report.AppendLine("[4. INSTALLED FILES IN $TargetDir]")
    foreach ($rf in $requiredFiles) {
        $fp = Join-Path $TargetDir $rf
        if (Test-Path $fp) {
            $fItem = Get-Item $fp
            $null = $report.AppendLine("  [PRESENT] $rf ($($fItem.Length) bytes, LastWrite: $($fItem.LastWriteTime))")
        } else {
            $null = $report.AppendLine("  [MISSING] $rf")
        }
    }
    $null = $report.AppendLine("")

    # Registry Status
    $null = $report.AppendLine("[5. REGISTRY CONFIGURATION]")
    $null = $report.AppendLine("COM CLSID ($Clsid):")
    $null = $report.AppendLine("  Path: HKCU:\Software\Classes\CLSID\$Clsid")
    $null = $report.AppendLine("  InprocServer32: $verifyCom")
    $null = $report.AppendLine("  ThreadingModel: $((Get-ItemProperty $inprocPath -ErrorAction SilentlyContinue).ThreadingModel)")
    $null = $report.AppendLine("ProgID ($ProgId):")
    $null = $report.AppendLine("  Path: HKCU:\Software\Classes\$ProgId")
    $null = $report.AppendLine("  CLSID: $((Get-ItemProperty $progIdClsid -ErrorAction SilentlyContinue).'(default)')")
    $null = $report.AppendLine("Outlook Addin Key:")
    $addinReg = Get-ItemProperty $outlookAddinKey -ErrorAction SilentlyContinue
    $null = $report.AppendLine("  FriendlyName:    $($addinReg.FriendlyName)")
    $null = $report.AppendLine("  Description:     $($addinReg.Description)")
    $null = $report.AppendLine("  LoadBehavior:    $($addinReg.LoadBehavior) $(if ($addinReg.LoadBehavior -eq 3) { '(OK: Load at startup)' } else { '(WARNING: Expected 3)' })")
    $null = $report.AppendLine("  CommandLineSafe: $($addinReg.CommandLineSafe)")
    $null = $report.AppendLine("")

    # Resiliency Status
    $null = $report.AppendLine("[6. OUTLOOK RESILIENCY & DISABLED ITEMS]")
    $disabledList = Get-OutlookDisabledAddins
    if ($disabledList.Count -eq 0) {
        $null = $report.AppendLine("  No disabled items found in Outlook Resiliency keys.")
    } else {
        foreach ($d in $disabledList) {
            $null = $report.AppendLine("  Office $($d.OfficeVersion) Disabled Item '$($d.ItemName)': $($d.DecodedText) (IsMove2Folder: $($d.IsMove2Folder))")
        }
    }
    Get-ChildItem "HKCU:\Software\Microsoft\Office" -ErrorAction SilentlyContinue | ForEach-Object {
        $dndKey = Join-Path $_.PSPath "Outlook\Resiliency\DoNotDisableAddinList"
        if (Test-Path $dndKey) {
            $val = (Get-ItemProperty $dndKey -ErrorAction SilentlyContinue).$ProgId
            $null = $report.AppendLine("  Office $($_.PSChildName) DoNotDisableAddinList for ${ProgId}: $val")
        }
    }
    $null = $report.AppendLine("")

    # Move2Folder User Settings
    $null = $report.AppendLine("[7. MOVE2FOLDER USER SETTINGS]")
    $settingsFile = [System.IO.Path]::Combine($env:APPDATA, "Move2Folder", "settings.json")
    if (Test-Path $settingsFile) {
        $null = $report.AppendLine("  Settings file: $settingsFile")
        $null = $report.AppendLine("  Content: $(Get-Content $settingsFile -Raw)")
    } else {
        $null = $report.AppendLine("  Settings file not yet created (will be generated upon first use).")
    }
    $null = $report.AppendLine("")

    # Windows Event Log - Outlook & Move2Folder
    $null = $report.AppendLine("[8. WINDOWS EVENT LOG DIAGNOSTICS]")
    $events = Get-DiagnosticEventLogs
    $m2fEvents = $events | Where-Object { $_.Message -like "*Move2Folder*" }

    $null = $report.AppendLine("Recent Events Mentioning Move2Folder ($($m2fEvents.Count) found):")
    if ($m2fEvents.Count -eq 0) {
        $null = $report.AppendLine("  None found in recent Application event logs.")
    } else {
        foreach ($e in $m2fEvents) {
            $null = $report.AppendLine("  ----------------------------------------------------------------------")
            $null = $report.AppendLine("  Timestamp: $($e.TimeCreated) | ID: $($e.Id) | Provider: $($e.ProviderName)")
            $null = $report.AppendLine("  Message: $($e.Message.Trim())")
        }
    }

    $null = $report.AppendLine("")
    $null = $report.AppendLine("Latest Outlook Event ID 45 (Add-Ins Loaded on Startup):")
    $evt45 = $events | Where-Object { $_.Id -eq 45 } | Select-Object -First 1
    if ($evt45) {
        $null = $report.AppendLine("  Timestamp: $($evt45.TimeCreated)")
        $null = $report.AppendLine("  Details:")
        $lines = $evt45.Message -split "`r?`n"
        foreach ($line in $lines) {
            $null = $report.AppendLine("    $line")
        }
    } else {
        $null = $report.AppendLine("  No recent Event ID 45 found.")
    }
    $null = $report.AppendLine("")

    # Troubleshooting Assessment
    $null = $report.AppendLine("================================================================================")
    $null = $report.AppendLine("                        TROUBLESHOOTING ASSESSMENT                              ")
    $null = $report.AppendLine("================================================================================")
    $issues = @()

    if ($outlook.Bitness -eq "x86") {
        $issues += "CRITICAL: Outlook is 32-bit (x86). Move2Folder requires 64-bit Outlook."
    }
    if (-not $dotnet.Installed) {
        $issues += "CRITICAL: .NET 8 Desktop Runtime (x64) is not installed."
    }
    if ($verifyAddin -ne 3) {
        $issues += "WARNING: Outlook Addin LoadBehavior is $verifyAddin (Expected: 3). Outlook may have set this to inactive."
    }
    if ($verifyCom -ne $comHostDll) {
        $issues += "ERROR: COM InprocServer32 path does not match installed DLL."
    }
    foreach ($rf in $requiredFiles) {
        if (-not (Test-Path (Join-Path $TargetDir $rf))) {
            $issues += "ERROR: Required binary $rf is missing from $TargetDir."
        }
    }

    if ($issues.Count -eq 0) {
        $null = $report.AppendLine("[HEALTHY] No configuration issues detected.")
        $null = $report.AppendLine("Move2Folder is correctly installed, registered, and configured for Outlook.")
    } else {
        $null = $report.AppendLine("[ISSUES FOUND]")
        foreach ($iss in $issues) {
            $null = $report.AppendLine("  * $iss")
        }
    }

    # Write report file
    [System.IO.File]::WriteAllText($debugLogPath, $report.ToString(), [System.Text.Encoding]::UTF8)

    Write-Host "`n  Debug report saved to:" -ForegroundColor Cyan
    Write-Host "  $debugLogPath" -ForegroundColor Yellow

    if ($issues.Count -gt 0) {
        Write-Host "`n  Detected Potential Issues:" -ForegroundColor Red
        foreach ($iss in $issues) {
            Write-Host "    - $iss" -ForegroundColor Yellow
        }
    } else {
        Write-Host "`n  Configuration status: All checks passed (Healthy)." -ForegroundColor Green
    }
}

# ---------------------------------------------------------------------------
# Final Status
# ---------------------------------------------------------------------------
Write-Host "`n==================================================" -ForegroundColor Green
Write-Host "    Move2Folder was installed successfully!       " -ForegroundColor Green
Write-Host "==================================================" -ForegroundColor Green
Write-Host "Location:  $TargetDir" -ForegroundColor Gray
Write-Host "Shortcut:  Ctrl + Shift + V (in Outlook Classic)" -ForegroundColor Cyan
Write-Host "You can now start Microsoft Outlook." -ForegroundColor Gray
if (-not $Debug) {
    Write-Host "(Tip: Run 'install.cmd --debug' anytime to collect diagnostic logs)`n" -ForegroundColor DarkGray
}
