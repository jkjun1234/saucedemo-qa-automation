@echo off
rem Role: start the GUI with this project's Python and Chromium installation.
cd /d "%~dp0"
set "PLAYWRIGHT_BROWSERS_PATH=%~dp0.browsers"
".venv\Scripts\python.exe" "app.py"
if errorlevel 1 pause
