#!/bin/bash
# ============================================================
# SAP Report Cleaner - macOS Starter
# Doppelklick zum Starten im Finder
# ============================================================

# Zum Skript-Verzeichnis wechseln (wo diese .command Datei liegt)
cd "$(dirname "$0")"

echo "============================================================"
echo "  SAP Report Cleaner"
echo "============================================================"
echo ""

# Python-Pfad finden (versucht verschiedene Optionen)
PYTHON_CMD=""

# Option 1: Lokale venv (falls vorhanden)
if [ -f ".venv/bin/python3" ]; then
    PYTHON_CMD=".venv/bin/python3"
elif [ -f ".venv/bin/python" ]; then
    PYTHON_CMD=".venv/bin/python"
# Option 2: System Python
elif command -v python3 &> /dev/null; then
    PYTHON_CMD="python3"
elif command -v python &> /dev/null; then
    # Prüfe ob es Python 3 ist
    if python --version 2>&1 | grep -q "Python 3"; then
        PYTHON_CMD="python"
    fi
fi

# Python nicht gefunden?
if [ -z "$PYTHON_CMD" ]; then
    echo "❌ Fehler: Python 3 nicht gefunden!"
    echo ""
    echo "Bitte Python 3 installieren:"
    echo "  → https://www.python.org/downloads/"
    echo ""
    echo "Oder mit Homebrew:"
    echo "  → brew install python3"
    echo ""
    read -p "Drücken Sie Enter zum Beenden..."
    exit 1
fi

echo "✓ Python gefunden: $PYTHON_CMD"

# Prüfe ob das Python-Script existiert
if [ ! -f "sap_report_cleaner_gui.py" ]; then
    echo ""
    echo "❌ Fehler: sap_report_cleaner_gui.py nicht gefunden!"
    echo "   Die Datei muss im gleichen Ordner wie diese .command Datei liegen."
    echo ""
    read -p "Drücken Sie Enter zum Beenden..."
    exit 1
fi

# Prüfe und installiere Abhängigkeiten
echo "✓ Prüfe Abhängigkeiten..."

# pandas prüfen
$PYTHON_CMD -c "import pandas" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "  → Installiere pandas..."
    $PYTHON_CMD -m pip install pandas -q
fi

# openpyxl prüfen
$PYTHON_CMD -c "import openpyxl" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "  → Installiere openpyxl..."
    $PYTHON_CMD -m pip install openpyxl -q
fi

echo "✓ Abhängigkeiten OK"
echo ""
echo "Starte SAP Report Cleaner..."
echo ""

# GUI-Script starten
$PYTHON_CMD sap_report_cleaner_gui.py

# Status prüfen
EXIT_CODE=$?

if [ $EXIT_CODE -ne 0 ]; then
    echo ""
    echo "❌ Ein Fehler ist aufgetreten (Code: $EXIT_CODE)"
    read -p "Drücken Sie Enter zum Beenden..."
fi
