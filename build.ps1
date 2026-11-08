# Skrypt budujacy pakiet dodatku NVDA (.nvda-addon) dla outlookImageHelper
$ErrorActionPreference = "Stop"

$scriptDir = $PSScriptRoot
$sourceDir = $scriptDir
$addonName = "outlookImageHelper"
$addonFile = Join-Path $scriptDir "$addonName.nvda-addon"
$tempZip = Join-Path $scriptDir "temp_build_$addonName.zip"
$validatorScript = Join-Path $scriptDir "Test-AddonSyntax.ps1"
if (-not (Test-Path $validatorScript)) {
    $validatorScript = Join-Path (Split-Path -Parent $scriptDir) "Test-AddonSyntax.ps1"
}

# 1. KROK KONTROLNY: Obowiazkowa walidacja skladni Python i AST przed spakowaniem
if (-not (Test-Path $validatorScript)) {
    Write-Host "BLAD KRYTYCZNY: Brak skryptu walidacji Test-AddonSyntax.ps1! Pakowanie zostalo zablokowane." -ForegroundColor Red
    exit 1
}
Write-Host "Uruchamianie automatycznej walidacji skladni Python i AST..." -ForegroundColor Cyan
& powershell.exe -ExecutionPolicy Bypass -File $validatorScript -TargetDir $sourceDir
if ($LASTEXITCODE -ne 0) {
    Write-Host "BLAD: Wykryto usterke skladniowa lub AST! Pakowanie zostalo przerwane." -ForegroundColor Red
    exit 1
}

if (Test-Path $tempZip) { Remove-Item $tempZip -Force }
if (Test-Path $addonFile) { Remove-Item $addonFile -Force }

# Usun foldery __pycache__ przed spakowaniem
Get-ChildItem -Path $sourceDir -Recurse -Directory -Filter "__pycache__" | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "Pakowanie dodatku z $sourceDir..."
$filesToPack = Get-ChildItem -Path $sourceDir -Exclude "$addonName.nvda-addon", "temp_build_*.zip", "build.*", "Test-AddonSyntax.ps1", "*.log"
Compress-Archive -Path $filesToPack.FullName -DestinationPath $tempZip -CompressionLevel Optimal
Rename-Item -Path $tempZip -NewName "$addonName.nvda-addon"

Write-Host "Gotowe! Utworzono pakiet: $addonFile" -ForegroundColor Green
