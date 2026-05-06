#Requires -Version 5.1
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) {
    Write-Error "PSScriptRoot is not set. Launch via Start.cmd or: powershell.exe -File `".\Start.ps1`""
    exit 1
}

$_ver = Join-Path $PSScriptRoot 'version.txt'
$appVersion = if (Test-Path $_ver) { (Get-Content $_ver -Raw).Trim() } else { '' }

. (Join-Path $PSScriptRoot 'src\UI\MainWindow.ps1')
Show-MainWindow -AppVersion $appVersion -ScriptRoot $PSScriptRoot
