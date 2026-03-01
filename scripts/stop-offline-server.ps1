$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$pidFile = Join-Path $root "runtime\server.pid"

if (-not (Test-Path $pidFile)) {
  Write-Output "No server PID file found. Nothing to stop."
  exit 0
}

$serverPid = Get-Content $pidFile -ErrorAction SilentlyContinue
if (-not $serverPid) {
  Remove-Item -Force $pidFile -ErrorAction SilentlyContinue
  Write-Output "PID file was empty."
  exit 0
}

$proc = Get-Process -Id $serverPid -ErrorAction SilentlyContinue
if ($proc) {
  Stop-Process -Id $serverPid -Force
  Write-Output "Stopped offline server (PID $serverPid)."
} else {
  Write-Output "Server process not running."
}

Remove-Item -Force $pidFile -ErrorAction SilentlyContinue
