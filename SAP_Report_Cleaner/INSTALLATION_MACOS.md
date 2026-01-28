# SAP Report Cleaner - macOS Installation

## Für Anfänger: Schritt-für-Schritt Anleitung

---

## Schritt 1: Python installieren

Python ist eine Programmiersprache, die für dieses Tool benötigt wird.

### 1.1 Python herunterladen

1. Öffnen Sie **Safari** (oder einen anderen Browser)
2. Gehen Sie zu: **https://www.python.org/downloads/**
3. Klicken Sie auf den gelben Button **"Download Python 3.x.x"**
4. Die Datei wird in Ihren **Downloads**-Ordner geladen

### 1.2 Python installieren

1. Öffnen Sie den **Finder**
2. Gehen Sie zu **Downloads**
3. **Doppelklick** auf die Datei `python-3.x.x-macos...pkg`
4. Ein Installationsfenster öffnet sich:
   - Klicken Sie auf **"Fortfahren"**
   - Klicken Sie auf **"Fortfahren"**
   - Klicken Sie auf **"Akzeptieren"** (Lizenz)
   - Klicken Sie auf **"Installieren"**
5. Geben Sie Ihr **Mac-Passwort** ein (das Passwort, mit dem Sie sich am Mac anmelden)
6. Klicken Sie auf **"Software installieren"**
7. Warten Sie bis "Die Installation war erfolgreich" erscheint
8. Klicken Sie auf **"Schließen"**

### 1.3 Prüfen ob es funktioniert hat

1. Öffnen Sie das **Terminal**:
   - Klicken Sie auf die **Lupe** oben rechts (Spotlight)
   - Tippen Sie `Terminal`
   - Drücken Sie `Enter`
2. Ein schwarzes/weißes Fenster öffnet sich
3. Tippen Sie: `python3 --version`
4. Drücken Sie `Enter`
5. Es sollte erscheinen: `Python 3.x.x`

✅ **Python ist installiert!**

Sie können das Terminal-Fenster jetzt schließen.

---

## Schritt 2: SAP Report Cleaner installieren

### 2.1 ZIP-Datei finden

Sie haben eine Datei erhalten: `SAP_Report_Cleaner.zip`

Diese befindet sich wahrscheinlich in:
- Ihrem **Downloads**-Ordner (wenn per E-Mail/Browser)
- oder dort wo Sie sie gespeichert haben

### 2.2 ZIP-Datei entpacken

1. Öffnen Sie den **Finder**
2. Gehen Sie zum Ordner mit der ZIP-Datei
3. **Doppelklick** auf `SAP_Report_Cleaner.zip`
4. Ein neuer Ordner `SAP_Report_Cleaner` erscheint

### 2.3 Ordner verschieben (optional)

Sie können den Ordner an einen beliebigen Ort verschieben:
- Auf den **Schreibtisch** (zum schnellen Zugriff)
- In **Dokumente** (zur Aufbewahrung)

### 2.4 Inhalt prüfen

Öffnen Sie den Ordner. Er sollte diese Dateien enthalten:

```
📁 SAP_Report_Cleaner
├── SAP_Report_Cleaner.command  ← Diese Datei starten Sie!
├── sap_report_cleaner_gui.py
├── sap_report_cleaner.py
├── requirements.txt
├── INSTALLATION_MACOS.md       ← Diese Anleitung
├── INSTALLATION_WINDOWS.md
└── SAP_Report_Cleaner_README.md
```

✅ **Installation abgeschlossen!**

---

## Schritt 3: Programm zum ersten Mal starten

### 3.1 Programm öffnen

1. Öffnen Sie den Ordner `SAP_Report_Cleaner` im Finder
2. **Doppelklick** auf `SAP_Report_Cleaner.command`

### 3.2 Sicherheitswarnung (erscheint nur beim ersten Mal)

macOS zeigt eine Warnung: *"SAP_Report_Cleaner.command kann nicht geöffnet werden"*

Das ist normal! So umgehen Sie die Warnung:

**Methode 1 (Einfach):**
1. Klicken Sie auf **"OK"** um die Warnung zu schließen
2. Gehen Sie zurück zum Finder
3. **Rechtsklick** auf `SAP_Report_Cleaner.command`
   - (Falls Sie keine Maus mit Rechtsklick haben: `Ctrl` gedrückt halten und klicken)
4. Wählen Sie **"Öffnen"** aus dem Menü
5. Klicken Sie im neuen Dialog auf **"Öffnen"**

**Methode 2 (Über Systemeinstellungen):**
1. Öffnen Sie **Systemeinstellungen** (Apple-Menü → Systemeinstellungen)
2. Klicken Sie auf **"Datenschutz & Sicherheit"**
3. Scrollen Sie nach unten
4. Sie sehen: *"SAP_Report_Cleaner.command wurde blockiert"*
5. Klicken Sie auf **"Trotzdem öffnen"**
6. Geben Sie Ihr Mac-Passwort ein

### 3.3 Erstes Starten

1. Ein Terminal-Fenster öffnet sich (schwarzer/weißer Hintergrund)
2. Sie sehen Text wie:
   ```
   ============================================================
     SAP Report Cleaner
   ============================================================
   ✓ Python gefunden
   Prüfe Abhängigkeiten...
   ```
3. Beim ersten Start werden automatisch benötigte Komponenten installiert
4. Warten Sie bis "Starte SAP Report Cleaner..." erscheint
5. Dann öffnet sich das Dateiauswahl-Fenster

✅ **Das Programm läuft!**

---

## Schritt 4: Programm benutzen

### 4.1 SAP-Report auswählen

1. Ein Finder-Fenster erscheint: "SAP-Report auswählen"
2. Navigieren Sie zu Ihrer SAP-Datei (`.txt` oder `.xls`)
3. Klicken Sie auf die Datei
4. Klicken Sie auf **"Öffnen"**

**Tipp:** Falls das Fenster nicht sichtbar ist, drücken Sie `Cmd + Tab` um zwischen Fenstern zu wechseln.

### 4.2 Ausgabeformat wählen

Ein Dialog fragt: *"Möchten Sie die Daten als Excel-Datei speichern?"*

| Auswahl | Ergebnis |
|---------|----------|
| **Ja** | Excel-Datei (.xlsx) mit 2 Tabellenblättern |
| **Nein** | CSV-Datei (.csv) - einfaches Textformat |

**Empfehlung:** Wählen Sie **"Ja"** für Excel.

### 4.3 Speicherort wählen

1. Ein Finder-Fenster erscheint: "Bereinigte Datei speichern als"
2. Wählen Sie einen Ordner (z.B. Downloads oder Schreibtisch)
3. Der Dateiname ist vorausgefüllt (z.B. `MeinReport_cleaned.xlsx`)
4. Klicken Sie auf **"Sichern"**

### 4.4 Fertig!

1. Ein Fenster zeigt: *"Verarbeitung abgeschlossen!"*
2. Sie sehen eine Statistik (wie viele Zeilen bereinigt wurden)
3. Klicken Sie auf **"OK"**
4. Die bereinigte Datei ist jetzt am gewählten Ort gespeichert!

---

## Schritt 5: Programm erneut starten (in Zukunft)

Ab jetzt ist es ganz einfach:

1. Öffnen Sie den Ordner `SAP_Report_Cleaner`
2. **Doppelklick** auf `SAP_Report_Cleaner.command`
3. Das Programm startet sofort

---

## Problemlösungen

### Problem: "python3: command not found"

**Ursache:** Python ist nicht installiert.
**Lösung:** Gehen Sie zurück zu Schritt 1 und installieren Sie Python.

### Problem: Das Programm startet nicht / Sicherheitswarnung

**Lösung:** 
1. Rechtsklick auf `SAP_Report_Cleaner.command`
2. "Öffnen" wählen
3. "Öffnen" bestätigen

### Problem: Dateiauswahl-Fenster erscheint nicht

**Lösung:** 
- Das Fenster ist vielleicht hinter anderen Fenstern versteckt
- Drücken Sie `Cmd + Tab` um alle offenen Programme zu sehen
- Oder klicken Sie auf "Python" im Dock (unten am Bildschirm)

### Problem: Fehlermeldung "Permission denied"

**Lösung:**
1. Öffnen Sie Terminal (Spotlight → Terminal)
2. Tippen Sie: `chmod +x ` (mit Leerzeichen am Ende)
3. Ziehen Sie die Datei `SAP_Report_Cleaner.command` ins Terminal-Fenster
4. Drücken Sie Enter
5. Schließen Sie Terminal
6. Starten Sie das Programm erneut

### Problem: Andere Fehlermeldung

**Lösung:** Machen Sie einen Screenshot der Fehlermeldung und senden Sie ihn an den Support.

---

## Kontakt bei Problemen

Bei Fragen wenden Sie sich an:
- [Hier Namen/E-Mail eintragen]

---

*SAP Report Cleaner - Version 1.0 - Januar 2026*
