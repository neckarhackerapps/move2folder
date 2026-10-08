<#
    .SYNOPSIS
        Creates a ready-to-use ZIP package for deployment to other computers.
#>

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RootDir = Split-Path -Parent $ScriptDir
$BinDir = Join-Path $RootDir "src\Move2Folder\bin\x64\Release\net8.0-windows"
$DistDir = Join-Path $RootDir "dist"
$PackageDir = Join-Path $DistDir "Move2Folder"
$ZipFile = Join-Path $DistDir "Move2Folder-Setup.zip"

Write-Host "=== Creating installation package for other computers ===" -ForegroundColor Cyan

# 1. Ensure project is built
& "$ScriptDir\build.ps1"

# 2. Clean and prepare output directory
if (Test-Path $PackageDir) { Remove-Item -Path $PackageDir -Recurse -Force }
if (Test-Path $ZipFile) { Remove-Item -Path $ZipFile -Force }
New-Item -ItemType Directory -Path $PackageDir -Force | Out-Null

# 3. Copy required binary files
$filesToCopy = @(
    "Move2Folder.comhost.dll",
    "Move2Folder.dll",
    "Move2Folder.runtimeconfig.json",
    "Move2Folder.deps.json",
    "Office.dll",
    "Microsoft.Office.Interop.Outlook.dll"
)

foreach ($f in $filesToCopy) {
    $src = Join-Path $BinDir $f
    if (Test-Path $src) {
        Copy-Item -Path $src -Destination (Join-Path $PackageDir $f) -Force
        Write-Host "Copied: $f" -ForegroundColor Gray
    } else {
        Write-Error "Required file not found: $src"
    }
}

# 4. Copy improved install.ps1 to package
$installScriptSrc = Join-Path $ScriptDir "install.ps1"
if (Test-Path $installScriptSrc) {
    Copy-Item -Path $installScriptSrc -Destination (Join-Path $PackageDir "install.ps1") -Force
    Write-Host "Copied: install.ps1" -ForegroundColor Gray
} else {
    Write-Error "install.ps1 not found in scripts directory: $installScriptSrc"
}

# 5. install.cmd (For double-click execution, forwarding any arguments like --debug)
$installCmdContent = @"
@echo off
title Move2Folder Installation
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1" %*
pause
"@
Set-Content -Path (Join-Path $PackageDir "install.cmd") -Value $installCmdContent -Encoding ASCII

# 6. uninstall.ps1
$uninstallPs1Content = @'
<#
.SYNOPSIS
    Uninstalls the Move2Folder Outlook Add-In cleanly.
#>
$ErrorActionPreference = "Continue"

$clsid = "{298621CD-332D-4DE0-B3B8-43BA73CA4893}"
$progId = "Move2Folder.Addin"
$TargetDir = Join-Path $env:LOCALAPPDATA "Move2Folder"

Write-Host "=== Move2Folder Uninstallation ===" -ForegroundColor Cyan

# 1. Outlook Addin Registration
Remove-Item -Path "HKCU:\Software\Microsoft\Office\Outlook\Addins\$progId" -Recurse -Force -ErrorAction SilentlyContinue

# 2. Resiliency keys
Get-ChildItem "HKCU:\Software\Microsoft\Office" -ErrorAction SilentlyContinue | ForEach-Object {
    $dndKey = Join-Path $_.PSPath "Outlook\Resiliency\DoNotDisableAddinList"
    if (Test-Path $dndKey) {
        Remove-ItemProperty -Path $dndKey -Name $progId -Force -ErrorAction SilentlyContinue
    }
}

# 3. COM Classes & ProgID
Remove-Item -Path "HKCU:\Software\Classes\$progId" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path "HKCU:\Software\Classes\CLSID\$clsid" -Recurse -Force -ErrorAction SilentlyContinue

# 4. Remove installation directory
if (Test-Path $TargetDir) {
    Remove-Item -Path $TargetDir -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "`n[OK] Move2Folder was uninstalled successfully." -ForegroundColor Green
'@
Set-Content -Path (Join-Path $PackageDir "uninstall.ps1") -Value $uninstallPs1Content -Encoding UTF8

# 7. uninstall.cmd
$uninstallCmdContent = @"
@echo off
title Move2Folder Uninstallation
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1" %*
pause
"@
Set-Content -Path (Join-Path $PackageDir "uninstall.cmd") -Value $uninstallCmdContent -Encoding ASCII

# 8. Create ZIP archive
Write-Host "Creating ZIP file: $ZipFile" -ForegroundColor Cyan
Compress-Archive -Path "$PackageDir\*" -DestinationPath $ZipFile -Force

Write-Host "`n[DONE] Installation package created successfully!" -ForegroundColor Green
Write-Host "Package folder: $PackageDir" -ForegroundColor Gray
Write-Host "ZIP file:       $ZipFile" -ForegroundColor Yellow
