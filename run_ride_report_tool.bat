@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "INPUT=%~1"
if "%INPUT%"=="" set "INPUT=C:\Users\saich\Downloads"
python "%SCRIPT_DIR%ride_report_tool.py" "%INPUT%" --out "%SCRIPT_DIR%outputs"
pause
