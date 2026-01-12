# SAP Report Cleaner - Windows Installation

## Für Anfänger: Schritt-für-Schritt Anleitung

---

## Schritt 1: Python installieren

Python ist eine Programmiersprache, die für dieses Tool benötigt wird.

### 1.1 Python herunterladen

1. Öffnen Sie Ihren **Webbrowser** (Edge, Chrome, Firefox)
2. Gehen Sie zu: **https://www.python.org/downloads/**
3. Klicken Sie auf den gelben Button **"Download Python 3.x.x"**
4. Die Datei wird heruntergeladen (z.B. in Ihren Downloads-Ordner)

### 1.2 Python installieren

1. Öffnen Sie den **Downloads**-Ordner
   - Drücken Sie `Windows + E` (öffnet den Explorer)
   - Klicken Sie links auf **"Downloads"**
2. **Doppelklick** auf die Datei `python-3.x.x-amd64.exe`
3. Ein Installationsfenster öffnet sich

**WICHTIG - Diese Option aktivieren:**

Am unteren Rand des Fensters sehen Sie:
```
☐ Add Python to PATH
```

4. **Setzen Sie dort einen Haken!** ☑️ Add Python to PATH
5. Klicken Sie auf **"Install Now"**
6. Warten Sie bis "Setup was successful" erscheint
7. Klicken Sie auf **"Close"**

### 1.3 Prüfen ob es funktioniert hat

1. Drücken Sie `Windows + R`
2. Tippen Sie `cmd`
3. Drücken Sie `Enter`
4. Ein schwarzes Fenster öffnet sich (Eingabeaufforderung)
5. Tippen Sie: `python --version`
6. Drücken Sie `Enter`
7. Es sollte erscheinen: `Python 3.x.x`

✅ **Python ist installiert!**

Sie können das schwarze Fenster jetzt schließen (X klicken oder `exit` tippen).

---

## Schritt 2: SAP Report Cleaner installieren

### 2.1 ZIP-Datei finden

Sie haben eine Datei erhalten: `SAP_Report_Cleaner.zip`

Diese befindet sich wahrscheinlich in:
- Ihrem **Downloads**-Ordner (wenn per E-Mail/Browser)
- oder dort wo Sie sie gespeichert haben

### 2.2 ZIP-Datei entpacken

1. Öffnen Sie den **Explorer** (Windows + E)
2. Gehen Sie zum Ordner mit der ZIP-Datei
3. **Rechtsklick** auf `SAP_Report_Cleaner.zip`
4. Wählen Sie **"Alle extrahieren..."**
5. Klicken Sie auf **"Extrahieren"**
6. Ein neuer Ordner `SAP_Report_Cleaner` wird erstellt

### 2.3 Ordner verschieben (optional)

Sie können den Ordner an einen beliebigen Ort verschieben:
- Auf den **Desktop** (zum schnellen Zugriff)
- In **Dokumente** (zur Aufbewahrung)

### 2.4 Inhalt prüfen

Öffnen Sie den Ordner. Er sollte diese Dateien enthalten:

```
📁 SAP_Report_Cleaner
├── SAP_Report_Cleaner.bat      ← Diese Datei starten Sie!
├── sap_report_cleaner_gui.py
├── sap_report_cleaner.py
├── requirements.txt
├── INSTALLATION_WINDOWS.md     ← Diese Anleitung
├── INSTALLATION_MACOS.md
└── SAP_Report_Cleaner_README.md
```

✅ **Installation abgeschlossen!**

---

## Schritt 3: Programm zum ersten Mal starten

### 3.1 Programm öffnen

1. Öffnen Sie den Ordner `SAP_Report_Cleaner`
2. **Doppelklick** auf `SAP_Report_Cleaner.bat`

### 3.2 Sicherheitswarnung (erscheint möglicherweise)

Windows zeigt möglicherweise: *"Der Computer wurde durch Windows geschützt"*

Das ist normal! So umgehen Sie die Warnung:

1. Klicken Sie auf **"Weitere Informationen"** (kleiner blauer Text)
2. Klicken Sie auf **"Trotzdem ausführen"**

### 3.3 Erstes Starten

1. Ein schwarzes Fenster öffnet sich (Eingabeaufforderung)
2. Sie sehen Text wie:
   ```
   ============================================================
     SAP Report Cleaner
   ============================================================
   + Python gefunden
   Pruefe Abhaengigkeiten...
   ```
3. Beim ersten Start werden automatisch benötigte Komponenten installiert
4. Das kann 1-2 Minuten dauern
5. Warten Sie bis "Starte SAP Report Cleaner..." erscheint
6. Dann öffnet sich das Dateiauswahl-Fenster

✅ **Das Programm läuft!**

---

## Schritt 4: Programm benutzen

### 4.1 SAP-Report auswählen

1. Ein Fenster erscheint: "SAP-Report auswählen"
2. Navigieren Sie zu Ihrer SAP-Datei (`.txt` oder `.xls`)
3. Klicken Sie auf die Datei
4. Klicken Sie auf **"Öffnen"**

**Tipp:** Falls das Fenster nicht sichtbar ist, schauen Sie in der Taskleiste unten.

### 4.2 Ausgabeformat wählen

Ein Dialog fragt: *"Möchten Sie die Daten als Excel-Datei speichern?"*

| Auswahl | Ergebnis |
|---------|----------|
| **Ja** | Excel-Datei (.xlsx) mit 2 Tabellenblättern |
| **Nein** | CSV-Datei (.csv) - einfaches Textformat |

**Empfehlung:** Wählen Sie **"Ja"** für Excel.

### 4.3 Speicherort wählen

1. Ein Fenster erscheint: "Bereinigte Datei speichern als"
2. Wählen Sie einen Ordner (z.B. Downloads oder Desktop)
3. Der Dateiname ist vorausgefüllt (z.B. `MeinReport_cleaned.xlsx`)
4. Klicken Sie auf **"Speichern"**

### 4.4 Fertig!

1. Ein Fenster zeigt: *"Verarbeitung abgeschlossen!"*
2. Sie sehen eine Statistik (wie viele Zeilen bereinigt wurden)
3. Klicken Sie auf **"OK"**
4. Die bereinigte Datei ist jetzt am gewählten Ort gespeichert!

---

## Schritt 5: Programm erneut starten (in Zukunft)

Ab jetzt ist es ganz einfach:

1. Öffnen Sie den Ordner `SAP_Report_Cleaner`
2. **Doppelklick** auf `SAP_Report_Cleaner.bat`
3. Das Programm startet sofort

**Tipp:** Sie können eine Verknüpfung auf dem Desktop erstellen:
1. Rechtsklick auf `SAP_Report_Cleaner.bat`
2. "Senden an" → "Desktop (Verknüpfung erstellen)"

---

## Problemlösungen

### Problem: "Python wurde nicht gefunden" oder "'python' wird nicht erkannt"

**Ursache:** Python wurde ohne "Add to PATH" installiert.

**Lösung:**
1. Python deinstallieren:
   - Windows-Einstellungen → Apps → Python suchen → Deinstallieren
2. Python neu installieren (Schritt 1)
3. **Diesmal unbedingt "Add Python to PATH" aktivieren!**

### Problem: Das Fenster schließt sich sofort wieder

**Lösung:**
1. Rechtsklick auf `SAP_Report_Cleaner.bat`
2. Wählen Sie **"Bearbeiten"** (oder "Mit Editor öffnen")
3. Fügen Sie ganz am Ende eine neue Zeile hinzu: `pause`
4. Speichern Sie die Datei
5. Starten Sie erneut
6. Jetzt sehen Sie die Fehlermeldung

### Problem: "Der Computer wurde durch Windows geschützt"

**Lösung:**
1. Klicken Sie auf **"Weitere Informationen"**
2. Klicken Sie auf **"Trotzdem ausführen"**

### Problem: Dateiauswahl-Fenster erscheint nicht

**Lösung:**
- Das Fenster ist vielleicht hinter anderen Fenstern versteckt
- Schauen Sie in der Taskleiste unten
- Klicken Sie auf das Python-Symbol

### Problem: Andere Fehlermeldung

**Lösung:** Machen Sie einen Screenshot der Fehlermeldung und senden Sie ihn an den Support.

**Screenshot erstellen:**
1. Drücken Sie `Windows + Shift + S`
2. Ziehen Sie ein Rechteck um die Fehlermeldung
3. Das Bild ist jetzt in der Zwischenablage
4. Fügen Sie es in eine E-Mail ein mit `Ctrl + V`

---

## Kontakt bei Problemen

Bei Fragen wenden Sie sich an:
- [Hier Namen/E-Mail eintragen]

---

*SAP Report Cleaner - Version 1.0 - Januar 2026*
