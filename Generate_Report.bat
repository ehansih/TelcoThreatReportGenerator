@echo off
title Telco Threat Intelligence Report Generator
setlocal

echo.
echo  ================================================
echo   TELCO THREAT INTEL REPORT GENERATOR
echo   Web UI — opens automatically in your browser
echo  ================================================
echo.

:: ── Try Windows Python ────────────────────────────────────────────────────────
python --version >nul 2>&1
if not errorlevel 1 goto :run_python

:: ── Try Python Launcher (py.exe) ──────────────────────────────────────────────
py --version >nul 2>&1
if not errorlevel 1 goto :run_py

:: ── Try WSL ───────────────────────────────────────────────────────────────────
wsl echo ok >nul 2>&1
if not errorlevel 1 goto :run_wsl

:: ── Nothing found ─────────────────────────────────────────────────────────────
echo  [ERROR] Python not found on this system.
echo.
echo  OPTION 1 — Install Python for Windows (recommended):
echo    1. Go to https://python.org/downloads
echo    2. Download Python 3.11 or higher
echo    3. Run the installer — check "Add Python to PATH"
echo    4. Restart Command Prompt and run this .bat again
echo.
echo  OPTION 2 — Use WSL:
echo    1. Open PowerShell as Administrator
echo    2. Run: wsl --install
echo    3. Restart your PC and run this .bat again
echo.
pause
exit /b 1

:run_python
echo  Using Windows Python. Installing dependencies...
pip install flask reportlab pyyaml python-docx --quiet
echo  Starting web server...
echo  Browser will open automatically at http://localhost:5000
echo.
python "%~dp0app.py"
goto :done

:run_py
echo  Using Python Launcher. Installing dependencies...
py -m pip install flask reportlab pyyaml python-docx --quiet
echo  Starting web server...
echo  Browser will open automatically at http://localhost:5000
echo.
py "%~dp0app.py"
goto :done

:run_wsl
echo  Using WSL (Linux subsystem). Installing dependencies...
for /f "delims=" %%P in ('wsl wslpath "%~dp0app.py"') do set WSL_APP=%%P
wsl bash -c "pip3 install flask reportlab pyyaml python-docx -q --break-system-packages 2>/dev/null; python3 '%WSL_APP%'"
goto :done

:done
endlocal
