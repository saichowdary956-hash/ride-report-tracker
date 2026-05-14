@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "RIDE_REPORT_BASE_DIR=%SCRIPT_DIR%app_data"
set "ALLOW_SQLITE_FALLBACK=1"
set "DATABASE_URL="
set "POSTGRES_URL="
set "POSTGRES_PRISMA_URL="
set "POSTGRES_URL_NON_POOLING="

if not exist "%RIDE_REPORT_BASE_DIR%\outputs" mkdir "%RIDE_REPORT_BASE_DIR%\outputs"
if not exist "%RIDE_REPORT_BASE_DIR%\uploads" mkdir "%RIDE_REPORT_BASE_DIR%\uploads"

where python >nul 2>nul
if errorlevel 1 (
  echo Python is required to run this offline app.
  echo Install Python 3.11 or newer, then double-click this file again.
  pause
  exit /b 1
)

python -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
  echo Installing required offline dependency: openpyxl
  python -m pip install -r "%SCRIPT_DIR%offline_requirements.txt"
  if errorlevel 1 (
    echo Could not install dependencies. Connect to internet once or install openpyxl manually.
    pause
    exit /b 1
  )
)

echo Starting RideReport Offline App...
echo Data folder: %RIDE_REPORT_BASE_DIR%
python "%SCRIPT_DIR%ride_report_app.py" --open
pause
