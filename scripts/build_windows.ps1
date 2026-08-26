$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $projectRoot

py -3 scripts\generate_icons.py
py -3 -m PyInstaller --noconfirm --clean packaging\MultiDocSync-Windows-x64.spec

Write-Host "Built: $projectRoot\dist\MultiDocSync-Windows-x64.exe"
