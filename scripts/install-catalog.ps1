$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$manifest = Join-Path $root "manifest.offline.xml"
$catalogId = "{9bdb4db7-65a6-43f3-9f84-6a8829fc20a9}"
$catalogKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$catalogId"

if (-not (Test-Path $manifest)) {
  throw "manifest.offline.xml not found at: $manifest"
}

New-Item -Path $catalogKey -Force | Out-Null
New-ItemProperty -Path $catalogKey -Name "Id" -Value $catalogId -PropertyType String -Force | Out-Null
New-ItemProperty -Path $catalogKey -Name "Url" -Value $root -PropertyType String -Force | Out-Null
New-ItemProperty -Path $catalogKey -Name "Flags" -Value 1 -PropertyType DWord -Force | Out-Null

Write-Output "Trusted add-in catalog configured for:"
Write-Output $root
Write-Output "Close and reopen Excel. Then go to Home -> Add-ins -> More Add-ins -> My Add-ins -> SHARED FOLDER."
