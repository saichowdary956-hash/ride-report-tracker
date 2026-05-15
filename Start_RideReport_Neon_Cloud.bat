@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "RIDE_REPORT_BASE_DIR=%SCRIPT_DIR%app_data_cloud"
set "ALLOW_SQLITE_FALLBACK=0"
set "NEON_URL_FILE=%SCRIPT_DIR%neon_database_url.txt"

if exist "%NEON_URL_FILE%" (
  set /p DATABASE_URL=<"%NEON_URL_FILE%"
) else (
  echo Paste your Neon PostgreSQL connection string below.
  echo It should start with postgresql:// or postgres://
  set /p DATABASE_URL=Neon DATABASE_URL: 
  >"%NEON_URL_FILE%" echo %DATABASE_URL%
)

if "%DATABASE_URL%"=="" (
  echo DATABASE_URL is empty. The app cannot use Neon storage.
  pause
  exit /b 1
)

set "POSTGRES_URL="
set "POSTGRES_PRISMA_URL="
set "POSTGRES_URL_NON_POOLING="

if not exist "%RIDE_REPORT_BASE_DIR%\outputs" mkdir "%RIDE_REPORT_BASE_DIR%\outputs"
if not exist "%RIDE_REPORT_BASE_DIR%\uploads" mkdir "%RIDE_REPORT_BASE_DIR%\uploads"

where python >nul 2>nul
if errorlevel 1 (
  echo Python is required to run this app.
  echo Install Python 3.11 or newer, then double-click this file again.
  pause
  exit /b 1
)

python -c "import openpyxl, psycopg" >nul 2>nul
if errorlevel 1 (
  echo Installing required dependencies for Neon mode...
  python -m pip install -r "%SCRIPT_DIR%requirements.txt"
  if errorlevel 1 (
    echo Could not install dependencies.
    pause
    exit /b 1
  )
)

echo Starting RideReport with Neon Postgres storage...
echo Local cache folder: %RIDE_REPORT_BASE_DIR%
echo Neon URL source: %NEON_URL_FILE%
python "%SCRIPT_DIR%ride_report_app.py" --open
pause
