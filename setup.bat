@echo off
cd /d "%~dp0"
echo Setting up the SOW Agent...
py -3 -m venv .venv 2>nul || python -m venv .venv
if not exist .venv\Scripts\python.exe (
  echo.
  echo Python was not found. Install Python 3.10 or newer from https://www.python.org/downloads/
  echo and tick "Add python.exe to PATH" in the installer. Then run this file again.
  pause
  exit /b 1
)
.venv\Scripts\python.exe -m pip install --upgrade pip
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo Installation failed. Check the internet connection and run this file again.
  pause
  exit /b 1
)
if not exist .env copy .env.example .env >nul
.venv\Scripts\python.exe tools\selftest.py
echo.
echo Setup finished. Next: open the .env file with Notepad and paste your API key.
pause
