@echo off
setlocal
set SCRIPT_DIR=%~dp0
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%scripts\install-catalog.ps1"
if errorlevel 1 (
  echo Failed to configure trusted add-in catalog.
  pause
  exit /b 1
)
echo.
echo Trusted add-in catalog configured.
echo Next in Excel: Home ^> Add-ins ^> More Add-ins ^> My Add-ins ^> SHARED FOLDER
pause
