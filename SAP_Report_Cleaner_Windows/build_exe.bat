@echo off
REM ============================================================
REM  SAP Report Cleaner - EXE Builder
REM  Erstellt eine standalone .exe Datei
REM ============================================================

echo.
echo ========================================
echo   SAP Report Cleaner - EXE Builder
echo ========================================
echo.

REM Prüfe Python
python --version >nul 2>&1
if errorlevel 1 (
    echo FEHLER: Python ist nicht installiert!
    echo Bitte installieren Sie Python von python.org
    pause
    exit /b 1
)

echo [1/3] Installiere PyInstaller...
pip install pyinstaller pandas numpy openpyxl --quiet

echo [2/3] Erstelle EXE-Datei...
echo      (Das kann 1-2 Minuten dauern)
echo.

pyinstaller --onefile --windowed --name "SAP_Report_Cleaner" ^
    --add-data "." ^
    --hidden-import pandas ^
    --hidden-import numpy ^
    --hidden-import openpyxl ^
    SAP_Report_Cleaner.py

echo.
echo [3/3] Aufräumen...
rmdir /s /q build 2>nul
del SAP_Report_Cleaner.spec 2>nul

echo.
echo ========================================
echo   FERTIG!
echo ========================================
echo.
echo Die EXE-Datei befindet sich in:
echo   dist\SAP_Report_Cleaner.exe
echo.
echo Diese Datei können Sie an Kollegen verteilen.
echo.
pause


