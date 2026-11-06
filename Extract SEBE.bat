@echo off
rem Double-click in File Explorer: makes SEBE files for any new CALM-dashboard-* folder.
rem Needs Python 3 from python.org (tick "Add python.exe to PATH" when installing).
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
  py -3 update-dashboard-data.py %*
) else (
  python update-dashboard-data.py %*
)
echo.
pause
