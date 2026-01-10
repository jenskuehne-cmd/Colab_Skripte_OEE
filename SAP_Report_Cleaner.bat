@echo off
chcp 65001 >nul
REM ============================================================
REM SAP Report Cleaner - Windows Starter
REM Doppelklick zum Starten im Explorer
REM ============================================================

cd /d "%~dp0"

echo ============================================================
echo   SAP Report Cleaner
echo ============================================================
echo.

REM Python finden
set PYTHON_CMD=

REM Option 1: Lokale venv
if exist ".venv\Scripts\python.exe" (
    set PYTHON_CMD=.venv\Scripts\python.exe
    goto :found
)

REM Option 2: System Python
where python >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    REM Prüfe ob Python 3
    python --version 2>&1 | findstr /C:"Python 3" >nul
    if %ERRORLEVEL% EQU 0 (
        set PYTHON_CMD=python
        goto :found
    )
)

where python3 >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    set PYTHON_CMD=python3
    goto :found
)

REM Python nicht gefunden
echo X Fehler: Python 3 nicht gefunden!
echo.
echo Bitte Python 3 installieren:
echo   https://www.python.org/downloads/
echo.
echo Wichtig: Bei der Installation "Add Python to PATH" aktivieren!
echo.
pause
exit /b 1

:found
echo + Python gefunden: %PYTHON_CMD%

REM Prüfe ob Script existiert
if not exist "sap_report_cleaner_gui.py" (
    echo.
    echo X Fehler: sap_report_cleaner_gui.py nicht gefunden!
    echo   Die Datei muss im gleichen Ordner wie diese .bat Datei liegen.
    echo.
    pause
    exit /b 1
)

REM Abhängigkeiten prüfen und installieren
echo + Pruefe Abhaengigkeiten...

%PYTHON_CMD% -c "import pandas" 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo   - Installiere pandas...
    %PYTHON_CMD% -m pip install pandas -q
)

%PYTHON_CMD% -c "import openpyxl" 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo   - Installiere openpyxl...
    %PYTHON_CMD% -m pip install openpyxl -q
)

echo + Abhaengigkeiten OK
echo.
echo Starte SAP Report Cleaner...
echo.

REM Script starten
%PYTHON_CMD% sap_report_cleaner_gui.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo X Ein Fehler ist aufgetreten.
    pause
)
