$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$dist = Join-Path $root "dist"
$outDir = Join-Path $root "release"
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$zipPath = Join-Path $outDir ("managerial-finance-addin-" + $timestamp + ".zip")

if (-not (Test-Path $dist)) {
  throw "dist folder not found. Run npm run build first."
}

if (-not (Test-Path $outDir)) {
  New-Item -ItemType Directory -Path $outDir | Out-Null
}

$staging = Join-Path $outDir "staging"
if (Test-Path $staging) {
  Remove-Item -Recurse -Force $staging
}
New-Item -ItemType Directory -Path $staging | Out-Null

Copy-Item -Recurse -Path $dist -Destination (Join-Path $staging "dist")
Copy-Item -Path (Join-Path $root "manifest.github.xml") -Destination (Join-Path $staging "manifest.xml")

Compress-Archive -Path (Join-Path $staging "*") -DestinationPath $zipPath -Force
Remove-Item -Recurse -Force $staging

Write-Output "Packaged release: $zipPath"
