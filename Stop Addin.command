#!/bin/bash
set -euo pipefail

ROOT="$(cd "$(dirname "$0")" && pwd)"
PID_FILE="$ROOT/runtime/server.pid"

if [[ ! -f "$PID_FILE" ]]; then
  echo "No PID file found. Nothing to stop."
  exit 0
fi

PID="$(cat "$PID_FILE" || true)"
if [[ -n "${PID:-}" ]] && ps -p "$PID" >/dev/null 2>&1; then
  kill "$PID"
  echo "Stopped offline server (PID $PID)."
else
  echo "Server process was not running."
fi

rm -f "$PID_FILE"
