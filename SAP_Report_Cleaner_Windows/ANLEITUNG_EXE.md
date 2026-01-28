# SAP Report Cleaner - Windows EXE Version

## Für Endbenutzer (Kollegen ohne Python)

Die `SAP_Report_Cleaner.exe` ist ein eigenständiges Programm - **keine Installation nötig!**

---

## So benutzen Sie das Programm

### Schritt 1: Programm starten

1. **Doppelklick** auf `SAP_Report_Cleaner.exe`
2. Das Fenster öffnet sich:

```
╔════════════════════════════════════════╗
║  🧹 SAP Report Cleaner                 ║
║  Bereinigt SAP-Reports in 3 Schritten  ║
╠════════════════════════════════════════╣
║  Schritt 1: SAP-Datei auswählen        ║
║  [📂 Datei wählen]                     ║
║                                        ║
║  Schritt 2: Ausgabeformat              ║
║  ◉ Excel (mit gelöschten Zeilen)       ║
║  ○ CSV (nur bereinigte Daten)          ║
║                                        ║
║  Schritt 3: Bereinigen & Speichern     ║
║  [🚀 Bereinigen & Speichern]           ║
╚════════════════════════════════════════╝
```

### Schritt 2: SAP-Datei auswählen

1. Klicken Sie auf **"📂 Datei wählen"**
2. Der Datei-Explorer öffnet sich (automatisch im Downloads-Ordner)
3. Wählen Sie Ihre SAP-Report-Datei (.txt oder .xls)
4. Klicken Sie auf **"Öffnen"**

### Schritt 3: Format wählen

- **Excel:** Enthält 2 Tabellenblätter (Bereinigte Daten + Gelöschte Zeilen)
- **CSV:** Nur die bereinigten Daten

### Schritt 4: Bereinigen & Speichern

1. Klicken Sie auf **"🚀 Bereinigen & Speichern"**
2. Die Datei wird verarbeitet
3. Ein Fenster zeigt die Statistik:
   - Anzahl bereinigte Zeilen
   - Anzahl entfernte Summenzeilen
   - Anzahl entfernte Zeilen ohne Materialnummer
4. Der Downloads-Ordner öffnet sich automatisch

### Fertig!

Die bereinigte Datei liegt in Ihrem **Downloads-Ordner**:
- `[Originalname]_cleaned.xlsx` oder
- `[Originalname]_cleaned.csv`

---

## Für Administratoren (EXE erstellen)

### Voraussetzungen

- Windows 10/11
- Python 3.8 oder neuer
- Internetverbindung (für einmalige Installation)

### EXE erstellen

1. **Python installieren** (falls noch nicht vorhanden):
   - python.org → Download → Windows installer
   - ☑️ "Add Python to PATH" aktivieren!

2. **Ordner öffnen:**
   - Diesen Ordner im Explorer öffnen

3. **Build-Script starten:**
   - Doppelklick auf `build_exe.bat`
   - Warten bis "FERTIG!" erscheint (ca. 1-2 Minuten)

4. **EXE finden:**
   - Die fertige EXE ist in: `dist\SAP_Report_Cleaner.exe`

### EXE verteilen

Die Datei `SAP_Report_Cleaner.exe` kann einzeln an Kollegen verteilt werden:
- Per E-Mail
- Per Netzlaufwerk
- Per USB-Stick

**Keine weitere Installation nötig!**

---

## Problemlösung

| Problem | Lösung |
|---------|--------|
| Windows SmartScreen Warnung | "Weitere Informationen" → "Trotzdem ausführen" |
| Programm startet nicht | Als Administrator ausführen (Rechtsklick) |
| Keine Daten | Prüfen ob SAP-Datei Tab-getrennt ist |
| Excel-Fehler | CSV wählen statt Excel |

### Windows SmartScreen

Beim ersten Start erscheint möglicherweise:
> "Der Computer wurde durch Windows geschützt"

1. Klicken Sie auf **"Weitere Informationen"**
2. Klicken Sie auf **"Trotzdem ausführen"**

Dies erscheint, weil die EXE nicht signiert ist - das Programm ist trotzdem sicher.

---

## Technische Details

| Eigenschaft | Wert |
|-------------|------|
| Python-Version | 3.8+ |
| Abhängigkeiten | pandas, numpy, openpyxl (in EXE enthalten) |
| EXE-Größe | ca. 30-50 MB |
| Betriebssystem | Windows 10/11 |

---

*SAP Report Cleaner v1.0 - Windows EXE - Januar 2026*

