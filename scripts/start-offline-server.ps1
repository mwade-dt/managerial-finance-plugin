$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$caddy = Join-Path $root "runtime\bin\windows_amd64\caddy.exe"
$pidFile = Join-Path $root "runtime\server.pid"
$url = "http://127.0.0.1:32123/taskpane.html"

if (-not (Test-Path $caddy)) {
  throw "Caddy runtime not found at: $caddy"
}

if (-not (Test-Path (Join-Path $root "dist\taskpane.html"))) {
  throw "dist/taskpane.html not found. Build artifacts are missing."
}

if (Test-Path $pidFile) {
  $existingPid = Get-Content $pidFile -ErrorAction SilentlyContinue
  if ($existingPid) {
    $existingProc = Get-Process -Id $existingPid -ErrorAction SilentlyContinue
    if ($existingProc) {
      Write-Output "Offline server already running (PID $existingPid)."
      Write-Output "Open Excel and upload manifest.offline.xml."
      exit 0
    }
  }
}

$proc = Start-Process -FilePath $caddy -ArgumentList "file-server", "--listen", "127.0.0.1:32123", "--root", (Join-Path $root "dist") -PassThru -WindowStyle Hidden -WorkingDirectory $root
$proc.Id | Set-Content $pidFile

Start-Sleep -Seconds 1
Write-Output "Offline server started (PID $($proc.Id))."
Write-Output "URL: $url"
Write-Output "Next: Excel -> Insert -> Get Add-ins -> Upload My Add-in -> manifest.offline.xml"
