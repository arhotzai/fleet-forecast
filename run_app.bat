@echo off
REM ============================================================
REM  Fleet Forecast - one-click launcher for Windows
REM  Just double-click this file. First run sets everything up.
REM ============================================================
cd /d "%~dp0"
title Fleet Forecast

echo.
echo ============================================
echo    FLEET FORECAST  -  starting up...
echo ============================================
echo.

REM --- check Python is installed ---
where python >nul 2>nul
if errorlevel 1 (
  echo  Python is NOT installed on this PC.
  echo.
  echo  1^) Install it from:  https://www.python.org/downloads/
  echo     IMPORTANT: on the first screen, tick "Add Python to PATH".
  echo  2^) Then double-click this file again.
  echo.
  pause
  exit /b
)

echo  Installing the required libraries...
echo  ^(first time only - may take a few minutes^)
echo.
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo  Could not install the libraries - this is usually the office
  echo  network blocking downloads. Ask IT to allow "pip / pypi.org",
  echo  or let me prepare an offline bundle.
  echo.
  pause
  exit /b
)

echo.
echo  Starting Fleet Forecast... a browser tab will open automatically.
echo  Leave this black window OPEN while you use the app.
echo  To stop the app, close this window.
echo.
python -m streamlit run app.py

pause
