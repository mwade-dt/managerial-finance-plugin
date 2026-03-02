$ErrorActionPreference = "Stop"

$catalogId = "{9bdb4db7-65a6-43f3-9f84-6a8829fc20a9}"
$catalogKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$catalogId"

if (Test-Path $catalogKey) {
  Remove-Item -Path $catalogKey -Recurse -Force
  Write-Output "Trusted add-in catalog removed."
} else {
  Write-Output "Trusted add-in catalog not found. Nothing to remove."
}
