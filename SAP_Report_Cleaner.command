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

# Python-Pfad finden
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
$PYTHON_CMD --version

# Prüfe ob das Python-Script existiert
if [ ! -f "sap_report_cleaner_gui.py" ]; then
    echo ""
    echo "❌ Fehler: sap_report_cleaner_gui.py nicht gefunden!"
    echo "   Die Datei muss im gleichen Ordner wie diese .command Datei liegen."
    echo ""
    read -p "Drücken Sie Enter zum Beenden..."
    exit 1
fi

# ============================================================
# Abhängigkeiten installieren
# ============================================================
echo ""
echo "Prüfe Abhängigkeiten..."

# Funktion zum Prüfen und Installieren eines Pakets
install_if_missing() {
    PACKAGE=$1
    IMPORT_NAME=${2:-$1}  # Falls Import-Name anders als Paketname
    
    $PYTHON_CMD -c "import $IMPORT_NAME" 2>/dev/null
    if [ $? -ne 0 ]; then
        echo "  → Installiere $PACKAGE..."
        $PYTHON_CMD -m pip install $PACKAGE -q --user
        if [ $? -ne 0 ]; then
            echo "    ⚠ Installation fehlgeschlagen, versuche ohne --user..."
            $PYTHON_CMD -m pip install $PACKAGE -q
        fi
    else
        echo "  ✓ $PACKAGE"
    fi
}

# Hauptabhängigkeiten installieren
install_if_missing "pandas"
install_if_missing "openpyxl"
install_if_missing "numpy"

# tkinter prüfen (kann nicht per pip installiert werden)
$PYTHON_CMD -c "import tkinter" 2>/dev/null
if [ $? -ne 0 ]; then
    echo ""
    echo "⚠ Warnung: tkinter nicht gefunden!"
    echo "  tkinter ist normalerweise in Python enthalten."
    echo ""
    echo "  Falls Sie Python über Homebrew installiert haben:"
    echo "    → brew install python-tk"
    echo ""
    echo "  Oder Python von python.org neu installieren"
    echo "  (enthält tkinter automatisch)"
    echo ""
    read -p "Drücken Sie Enter um fortzufahren (oder Ctrl+C zum Abbrechen)..."
else
    echo "  ✓ tkinter"
fi

echo ""
echo "✓ Alle Abhängigkeiten OK"
echo ""
echo "============================================================"
echo "Starte SAP Report Cleaner..."
echo "============================================================"
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
