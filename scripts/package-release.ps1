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
Copy-Item -Path (Join-Path $root "manifest.offline.xml") -Destination (Join-Path $staging "manifest.offline.xml")
Copy-Item -Recurse -Path (Join-Path $root "runtime") -Destination (Join-Path $staging "runtime")
Copy-Item -Path (Join-Path $root "Start Addin.bat") -Destination (Join-Path $staging "Start Addin.bat")
Copy-Item -Path (Join-Path $root "Stop Addin.bat") -Destination (Join-Path $staging "Stop Addin.bat")
Copy-Item -Path (Join-Path $root "Start Addin.command") -Destination (Join-Path $staging "Start Addin.command")
Copy-Item -Path (Join-Path $root "Stop Addin.command") -Destination (Join-Path $staging "Stop Addin.command")
Copy-Item -Path (Join-Path $root "README.md") -Destination (Join-Path $staging "README.md")
Copy-Item -Path (Join-Path $root "README.txt") -Destination (Join-Path $staging "README.txt")
Copy-Item -Path (Join-Path $root "DEPLOY_AND_SHARE.md") -Destination (Join-Path $staging "DEPLOY_AND_SHARE.md")

Compress-Archive -Path (Join-Path $staging "*") -DestinationPath $zipPath -Force
Remove-Item -Recurse -Force $staging

Write-Output "Packaged release: $zipPath"
