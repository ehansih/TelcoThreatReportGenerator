@echo off
title Telco Threat Intelligence Report Generator
setlocal

echo.
echo  ================================================
echo   TELCO THREAT INTEL REPORT GENERATOR
echo  ================================================
echo.

:: ── Try Windows Python ────────────────────────────────────────────────────────
python --version >nul 2>&1
if not errorlevel 1 goto :run_python

:: ── Try Python Launcher (py.exe) ──────────────────────────────────────────────
py --version >nul 2>&1
if not errorlevel 1 goto :run_py

:: ── Try WSL (Windows Subsystem for Linux) ────────────────────────────────────
wsl echo ok >nul 2>&1
if not errorlevel 1 goto :run_wsl

:: ── Nothing found ─────────────────────────────────────────────────────────────
echo.
echo  [ERROR] Python not found on this system.
echo.
echo  To fix this, choose one option:
echo.
echo  OPTION 1 — Install Python for Windows (recommended):
echo    1. Open: https://python.org/downloads
echo    2. Download Python 3.11 or higher
echo    3. Run the installer
echo    4. IMPORTANT: check "Add Python to PATH" during install
echo    5. Open a new Command Prompt and run this .bat again
echo.
echo  OPTION 2 — Use WSL (Windows Subsystem for Linux):
echo    1. Open PowerShell as Administrator
echo    2. Run: wsl --install
echo    3. Restart your PC, then run this .bat again
echo.
pause
exit /b 1

:: ── Windows Python ────────────────────────────────────────────────────────────
:run_python
echo  Found Python. Checking dependencies...
python -c "import reportlab" >nul 2>&1
if errorlevel 1 (
    echo  Installing reportlab...
    pip install reportlab --quiet
)
python -c "import yaml" >nul 2>&1
if errorlevel 1 (
    echo  Installing pyyaml...
    pip install pyyaml --quiet
)
python -c "import docx" >nul 2>&1
if errorlevel 1 (
    echo  Installing python-docx...
    pip install python-docx --quiet
)
echo  Launching...
echo.
python "%~dp0generate_report_gui.py"
goto :done

:: ── Python Launcher ───────────────────────────────────────────────────────────
:run_py
echo  Found Python Launcher (py.exe). Checking dependencies...
py -m pip install reportlab pyyaml python-docx --quiet
echo  Launching...
echo.
py "%~dp0generate_report_gui.py"
goto :done

:: ── WSL ───────────────────────────────────────────────────────────────────────
:run_wsl
echo  Windows Python not found. Using WSL (Linux subsystem)...
echo  Note: GUI requires WSLg support (Windows 11 or Windows 10 with WSLg update).
echo.
for /f "delims=" %%P in ('wsl wslpath "%~dp0generate_report_gui.py"') do set WSL_SCRIPT=%%P
wsl bash -c "pip3 install reportlab pyyaml python-docx -q 2>/dev/null; python3 '%WSL_SCRIPT%'"
goto :done

:done
endlocal
