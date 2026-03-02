@echo off
setlocal
set SCRIPT_DIR=%~dp0
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%scripts\uninstall-catalog.ps1"
if errorlevel 1 (
  echo Failed to remove trusted add-in catalog.
  pause
  exit /b 1
)
echo.
echo Trusted add-in catalog removed.
pause
