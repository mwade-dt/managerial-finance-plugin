#!/bin/bash
set -euo pipefail

ROOT="$(cd "$(dirname "$0")" && pwd)"
PID_FILE="$ROOT/runtime/server.pid"

if [[ -f "$PID_FILE" ]]; then
  PID="$(cat "$PID_FILE" || true)"
  if [[ -n "${PID:-}" ]] && ps -p "$PID" >/dev/null 2>&1; then
    echo "Offline server already running (PID $PID)."
    echo "Open Excel and upload manifest.offline.xml"
    exit 0
  fi
fi

ARCH="$(uname -m)"
if [[ "$ARCH" == "arm64" ]]; then
  CADDY="$ROOT/runtime/bin/mac_arm64/caddy"
else
  CADDY="$ROOT/runtime/bin/mac_amd64/caddy"
fi

if [[ ! -f "$CADDY" ]]; then
  echo "Caddy runtime not found: $CADDY"
  exit 1
fi

if [[ ! -f "$ROOT/dist/taskpane.html" ]]; then
  echo "dist/taskpane.html not found."
  exit 1
fi

chmod +x "$CADDY"
nohup "$CADDY" file-server --listen 127.0.0.1:32123 --root "$ROOT/dist" >/dev/null 2>&1 &
echo $! > "$PID_FILE"

echo "Offline server started at http://127.0.0.1:32123/taskpane.html"
echo "In Excel: Insert > Get Add-ins > Upload My Add-in > manifest.offline.xml"
