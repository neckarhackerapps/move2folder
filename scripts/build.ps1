<#
.SYNOPSIS
    Compiles the Move2Folder Outlook Add-In in Release mode (x64).
#>

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RootDir = Split-Path -Parent $ScriptDir

Write-Host "=== Building Move2Folder (Release x64) ===" -ForegroundColor Cyan
Set-Location $RootDir

$outlook = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
if ($outlook) {
    Write-Warning "Outlook is currently running and may lock DLL files in bin\x64\Release. Please close Outlook if you experience file lock errors."
}

dotnet build -c Release

if ($LASTEXITCODE -eq 0) {
    Write-Host "`n[OK] Build completed successfully!" -ForegroundColor Green
    $dllPath = Join-Path $RootDir "src\Move2Folder\bin\x64\Release\net8.0-windows\Move2Folder.comhost.dll"
    if (Test-Path $dllPath) {
        Write-Host "COM host ready: $dllPath" -ForegroundColor Gray
    }
} else {
    Write-Host "`n[ERROR] Build failed with exit code $LASTEXITCODE." -ForegroundColor Red
    exit $LASTEXITCODE
}
