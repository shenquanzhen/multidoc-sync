$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $projectRoot

$pythonCommand = Get-Command python -ErrorAction SilentlyContinue
$pythonArguments = @()
if (-not $pythonCommand) {
  $pythonCommand = Get-Command py -ErrorAction Stop
  $pythonArguments = @('-3')
}

& $pythonCommand.Source @pythonArguments scripts\generate_icons.py
if ($LASTEXITCODE -ne 0) { throw "Icon generation failed with exit code $LASTEXITCODE" }

& $pythonCommand.Source @pythonArguments -m PyInstaller --noconfirm --clean packaging\MultiDocSync-Windows-x64.spec
if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed with exit code $LASTEXITCODE" }

Write-Host "Built: $projectRoot\dist\MultiDocSync-Windows-x64.exe"
