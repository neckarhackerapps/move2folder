<#
.SYNOPSIS
    Registers the Move2Folder COM Add-In in the Windows Registry for the current user (HKCU).
    Requires NO administrator privileges!
#>

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RootDir = Split-Path -Parent $ScriptDir

$ComHostDll = Join-Path $RootDir "src\Move2Folder\bin\x64\Release\net8.0-windows\Move2Folder.comhost.dll"

if (-not (Test-Path $ComHostDll)) {
    Write-Host "[INFO] Compiling Move2Folder before registration..." -ForegroundColor Yellow
    & "$ScriptDir\build.ps1"
    if (-not (Test-Path $ComHostDll)) {
        Write-Error "Move2Folder.comhost.dll not found at: $ComHostDll"
        exit 1
    }
}

$clsid = "{298621CD-332D-4DE0-B3B8-43BA73CA4893}"
$progId = "Move2Folder.Addin"

Write-Host "=== Registering Move2Folder Add-In ===" -ForegroundColor Cyan
Write-Host "DLL Path: $ComHostDll" -ForegroundColor Gray
Write-Host "CLSID:    $clsid" -ForegroundColor Gray
Write-Host "ProgID:   $progId" -ForegroundColor Gray

# 1. COM CLSID in HKCU\Software\Classes
$clsidPath = "HKCU:\Software\Classes\CLSID\$clsid"
$inprocPath = "$clsidPath\InprocServer32"
$progIdSubPath = "$clsidPath\ProgID"

if (-not (Test-Path $clsidPath)) { New-Item -Path $clsidPath -Force | Out-Null }
Set-ItemProperty -Path $clsidPath -Name "(default)" -Value $progId

if (-not (Test-Path $inprocPath)) { New-Item -Path $inprocPath -Force | Out-Null }
Set-ItemProperty -Path $inprocPath -Name "(default)" -Value $ComHostDll
Set-ItemProperty -Path $inprocPath -Name "ThreadingModel" -Value "Both"

if (-not (Test-Path $progIdSubPath)) { New-Item -Path $progIdSubPath -Force | Out-Null }
Set-ItemProperty -Path $progIdSubPath -Name "(default)" -Value $progId

# 2. ProgID in HKCU\Software\Classes
$progIdRoot = "HKCU:\Software\Classes\$progId"
$progIdClsid = "$progIdRoot\CLSID"

if (-not (Test-Path $progIdRoot)) { New-Item -Path $progIdRoot -Force | Out-Null }
Set-ItemProperty -Path $progIdRoot -Name "(default)" -Value $progId

if (-not (Test-Path $progIdClsid)) { New-Item -Path $progIdClsid -Force | Out-Null }
Set-ItemProperty -Path $progIdClsid -Name "(default)" -Value $clsid

# 3. Outlook Add-In Registration in HKCU\Software\Microsoft\Office\Outlook\Addins
$outlookAddinKey = "HKCU:\Software\Microsoft\Office\Outlook\Addins\$progId"
if (-not (Test-Path $outlookAddinKey)) { New-Item -Path $outlookAddinKey -Force | Out-Null }
Set-ItemProperty -Path $outlookAddinKey -Name "FriendlyName" -Value "Move2Folder"
Set-ItemProperty -Path $outlookAddinKey -Name "Description" -Value "Move2Folder Outlook Classic Add-In"
Set-ItemProperty -Path $outlookAddinKey -Name "LoadBehavior" -Value 3 -Type DWord
Set-ItemProperty -Path $outlookAddinKey -Name "CommandLineSafe" -Value 0 -Type DWord

Write-Host "`n[OK] Move2Folder was registered successfully!" -ForegroundColor Green

# Outlook runtime check
$outlookProcess = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
if ($outlookProcess) {
    Write-Host "`n[NOTICE] Outlook is already running. Please restart Outlook to load the updated Add-In." -ForegroundColor Yellow
} else {
    Write-Host "`n[INFO] Start Outlook now to use Move2Folder." -ForegroundColor Gray
}
