# Skrypt budujący pakiet dodatku NVDA (.nvda-addon) dla outlookImageHelper
$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$sourceDir = $scriptDir
$addonName = "outlookImageHelper"
$addonFile = Join-Path $scriptDir "$addonName.nvda-addon"
$tempZip = Join-Path $scriptDir "temp_build_$addonName.zip"

if (Test-Path $tempZip) { Remove-Item $tempZip -Force }
if (Test-Path $addonFile) { Remove-Item $addonFile -Force }

# Usuń foldery __pycache__ przed spakowaniem
Get-ChildItem -Path $sourceDir -Recurse -Directory -Filter "__pycache__" | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "Pakowanie dodatku z $sourceDir..."
$filesToPack = Get-ChildItem -Path $sourceDir -Exclude "$addonName.nvda-addon", "temp_build_*.zip", "build.*", "*.log"
Compress-Archive -Path $filesToPack.FullName -DestinationPath $tempZip -CompressionLevel Optimal
Rename-Item -Path $tempZip -NewName "$addonName.nvda-addon"

Write-Host "Gotowe! Utworzono pakiet: $addonFile" -ForegroundColor Green
