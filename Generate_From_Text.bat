@echo off
title Telco Threat Report Generator - Text Input Mode
setlocal

echo.
echo  ================================================
echo   TELCO THREAT REPORT GENERATOR - TEXT MODE
echo  ================================================
echo.

if "%~1"=="" (
  echo  Usage:
  echo    %~n0 ^<input.txt^> [output.pdf]
  echo.
  echo  Example:
  echo    %~n0 C:\Users\havardha\Downloads\report_input.txt C:\Users\havardha\Downloads\report.pdf
  echo.
  pause
  exit /b 1
)

set "INPUT_FILE=%~1"
set "OUTPUT_FILE=%~2"

python --version >nul 2>&1
if errorlevel 1 (
  py --version >nul 2>&1
  if errorlevel 1 (
    echo  [ERROR] Python not found. Install Python 3.8+ and try again.
    pause
    exit /b 1
  )
  set "PY_CMD=py"
) else (
  set "PY_CMD=python"
)

%PY_CMD% -m pip install reportlab pyyaml python-docx --quiet

if "%OUTPUT_FILE%"=="" (
  %PY_CMD% "%~dp0generate_from_text.py" --input "%INPUT_FILE%"
) else (
  %PY_CMD% "%~dp0generate_from_text.py" --input "%INPUT_FILE%" --pdf "%OUTPUT_FILE%"
)

echo.
pause
