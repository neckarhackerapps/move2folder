<#
.SYNOPSIS
    Unregisters the Move2Folder COM Add-In from the Windows Registry (HKCU).
#>

$ErrorActionPreference = "Continue"

$clsid = "{298621CD-332D-4DE0-B3B8-43BA73CA4893}"
$progId = "Move2Folder.Addin"

Write-Host "=== Unregistering Move2Folder Add-In ===" -ForegroundColor Cyan

# 1. Outlook Add-In Key
$outlookAddinKey = "HKCU:\Software\Microsoft\Office\Outlook\Addins\$progId"
if (Test-Path $outlookAddinKey) {
    Remove-Item -Path $outlookAddinKey -Recurse -Force
    Write-Host "Removed: $outlookAddinKey" -ForegroundColor Gray
}

# 2. ProgID
$progIdRoot = "HKCU:\Software\Classes\$progId"
if (Test-Path $progIdRoot) {
    Remove-Item -Path $progIdRoot -Recurse -Force
    Write-Host "Removed: $progIdRoot" -ForegroundColor Gray
}

# 3. CLSID
$clsidPath = "HKCU:\Software\Classes\CLSID\$clsid"
if (Test-Path $clsidPath) {
    Remove-Item -Path $clsidPath -Recurse -Force
    Write-Host "Removed: $clsidPath" -ForegroundColor Gray
}

Write-Host "`n[OK] Move2Folder was unregistered successfully." -ForegroundColor Green
