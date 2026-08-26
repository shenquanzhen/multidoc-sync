$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $projectRoot

python scripts\generate_icons.py
if ($LASTEXITCODE -ne 0) { throw "Icon generation failed with exit code $LASTEXITCODE" }

python -m PyInstaller --noconfirm --clean packaging\MultiDocSync-Windows-x64.spec
if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed with exit code $LASTEXITCODE" }

Write-Host "Built: $projectRoot\dist\MultiDocSync-Windows-x64.exe"
