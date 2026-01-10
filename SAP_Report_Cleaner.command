#!/bin/bash
# ============================================================
# SAP Report Cleaner - macOS Starter
# Doppelklick zum Starten im Finder
# ============================================================

# Zum Skript-Verzeichnis wechseln
cd "$(dirname "$0")"

# Python-Pfad finden (versucht verschiedene Optionen)
if command -v python3 &> /dev/null; then
    PYTHON_CMD="python3"
elif command -v python &> /dev/null; then
    PYTHON_CMD="python"
elif [ -f ".venv/bin/python" ]; then
    PYTHON_CMD=".venv/bin/python"
elif [ -f ".venv/bin/python3" ]; then
    PYTHON_CMD=".venv/bin/python3"
else
    echo "❌ Fehler: Python nicht gefunden!"
    echo "Bitte Python 3 installieren: https://www.python.org/downloads/"
    read -p "Drücken Sie Enter zum Beenden..."
    exit 1
fi

echo "============================================================"
echo "  SAP Report Cleaner"
echo "============================================================"
echo ""
echo "Python: $PYTHON_CMD"
echo ""

# GUI-Script starten
$PYTHON_CMD sap_report_cleaner_gui.py

# Warten falls Fehler aufgetreten sind
if [ $? -ne 0 ]; then
    echo ""
    echo "❌ Ein Fehler ist aufgetreten."
    read -p "Drücken Sie Enter zum Beenden..."
fi

