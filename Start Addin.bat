@echo off
setlocal
set SCRIPT_DIR=%~dp0
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%scripts\start-offline-server.ps1"
if errorlevel 1 (
  echo Failed to start offline server.
  pause
  exit /b 1
)
echo.
echo Offline server is running.
echo In Excel: Insert ^> Get Add-ins ^> Upload My Add-in ^> manifest.offline.xml
pause
