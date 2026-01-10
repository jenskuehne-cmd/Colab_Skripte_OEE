# SAP Report Cleaner

## Übersicht

Dieses Python-Script bereinigt SAP-Reports, die als Tab-getrennte Textdateien (`.txt` oder pseudo-`.xls`) vorliegen und bereitet sie für die Datenanalyse vor.

### Versionen

| Datei | Beschreibung | Starten |
|-------|--------------|---------|
| `SAP_Report_Cleaner.command` | **macOS Starter** | Doppelklick im Finder |
| `SAP_Report_Cleaner.bat` | **Windows Starter** | Doppelklick im Explorer |
| `sap_report_cleaner_gui.py` | GUI-Version | `python3 sap_report_cleaner_gui.py` |
| `sap_report_cleaner.py` | Kommandozeilen-Version | `python3 sap_report_cleaner.py [datei]` |

---

## Voraussetzungen

### Python-Version
- Python 3.7 oder höher

### Abhängigkeiten installieren

```bash
pip3 install pandas openpyxl
```

---

## Verwendung

### Option 1: Doppelklick (empfohlen für Endanwender)

**macOS:**
1. Doppelklick auf `SAP_Report_Cleaner.command` im Finder
2. Falls Sicherheitswarnung: Rechtsklick → "Öffnen" → "Öffnen" bestätigen

**Windows:**
1. Doppelklick auf `SAP_Report_Cleaner.bat` im Explorer
2. Falls Sicherheitswarnung: "Trotzdem ausführen" klicken

**Ablauf nach dem Start:**
1. 📂 **Dateiauswahl-Dialog** → Quelldatei auswählen
2. 💾 **Speichern-Dialog** → Zielort und Dateiname wählen
3. ✅ **Erfolgsmeldung** mit Statistik

### Option 2: Terminal/Kommandozeile

```bash
python3 sap_report_cleaner_gui.py
```

### Option 2: Kommandozeile

```bash
# Mit direktem Dateipfad
python3 sap_report_cleaner.py sourceDateien/L91_Material.txt

# Interaktive Dateiauswahl
python3 sap_report_cleaner.py
```

### Option 3: In Python/Jupyter importieren

```python
from sap_report_cleaner import run

# Mit Dateipfad
df = run("sourceDateien/L91_Material.txt")

# Interaktiv
df = run()

# DataFrame ist direkt für weitere Analyse verfügbar
print(df.describe())
```

---

## Eingabedateien

Das Script verarbeitet SAP-Reports mit folgender Struktur:

| Zeile | Inhalt |
|-------|--------|
| 1 | Datum und Titel (wird ignoriert) |
| 2-3 | Leer |
| 4 | **Header-Zeile** (Spaltenüberschriften) |
| 5+ | Datenzeilen |

### Erwartete Spaltenstruktur

| Spalte | Name | Beschreibung |
|--------|------|--------------|
| A | (leer) | Wird ignoriert |
| B | Marker | `*` = Summenzeile (wird gelöscht) |
| C-Q | Daten | Werden extrahiert |

---

## Ausgabedateien

Das Script erzeugt zwei Ausgabedateien im gleichen Verzeichnis wie die Eingabedatei:

### 1. CSV-Datei: `[name]_cleaned.csv`
- Trennzeichen: Semikolon (`;`)
- Encoding: UTF-8 mit BOM (Excel-kompatibel)
- Enthält nur bereinigte Daten

### 2. Excel-Datei: `[name]_cleaned.xlsx`
- **Sheet "Bereinigte Daten"**: Alle bereinigten Datensätze
- **Sheet "Gelöschte Zeilen"**: Protokoll der entfernten Zeilen mit Löschgrund

---

## Datenbereinigung

Das Script führt folgende Bereinigungen durch:

| Aktion | Beschreibung |
|--------|--------------|
| ✅ Header erkennen | Findet automatisch die Zeile mit "Material" |
| ✅ Spalten filtern | Behält nur Spalten C bis Q |
| ✅ Summenzeilen entfernen | Zeilen mit `*` oder `**` in Spalte B |
| ✅ Leere Zeilen entfernen | Komplett leere Zeilen |
| ✅ Ohne Materialnummer entfernen | Zeilen ohne Wert in Spalte C |
| ✅ Zahlenformate bereinigen | Deutsche Formate (1.234,56) → Standard |
| ✅ Datumsformate konvertieren | DD.MM.YY → DD.MM.YYYY |

---

## Spalten und Datentypen

### Ausgabe-Spalten (C bis Q)

| # | Spaltenname | Datentyp | Beschreibung | Beispiel |
|---|-------------|----------|--------------|----------|
| 1 | **Material** | `int64` | Materialnummer (SAP) | `86008355` |
| 2 | **Functional Loc.** | `str` | Technischer Platz | `232VSTE091-TRP-004` |
| 3 | **Equipment** | `str` | Equipmentnummer | `10007109` |
| 4 | **Material Description** | `str` | Materialbeschreibung | `FOERDERGURT (P-3000)` |
| 5 | **Work Ctr** | `str` | Arbeitsplatz | `MEC02010` |
| 6 | **Withdrawn** | `int64` | Entnommen (Menge) | `-10`, `5` |
| 7 | **W/o resrv.** | `int64` | Ohne Reservierung | `-8` |
| 8 | **Reserved** | `int64` | Reserviert | `3500` |
| 9 | **Reserv.ref** | `int64` | Reservierungsreferenz | `0` |
| 10 | **Pstng Date** | `str` | Buchungsdatum | `28.06.2023` |
| 11 | **Order** | `int64` | Auftragsnummer | `40910756` |
| 12 | **ID** | `str` | ID | `12856365` |
| 13 | **Message** | `int64` | Nachricht | `12856365` |
| 14 | **ICt** | `str` | Indikator | `L`, `1` |
| 15 | **Customer** | `str` | Kunde | (meist leer) |

### Datentypen für Pandas

```python
# Nach dem Import des DataFrames
df.dtypes

# Ergebnis:
# Material                  int64  (oder float64 bei NaN)
# Functional Loc.          object  (string)
# Equipment                object  (string)
# Material Description     object  (string)
# Work Ctr                 object  (string)
# Withdrawn               float64  (kann NaN enthalten)
# W/o resrv.              float64  (kann NaN enthalten)
# Reserved                float64  (kann NaN enthalten)
# Reserv.ref              float64  (kann NaN enthalten)
# Pstng Date               object  (string, Format: DD.MM.YYYY)
# Order                   float64  (kann NaN enthalten)
# ID                       object  (string)
# Message                 float64  (kann NaN enthalten)
# ICt                      object  (string)
# Customer                 object  (string)
```

### Hinweise zu Datentypen

| Typ | Hinweis |
|-----|---------|
| `int64` / `float64` | Numerische Spalten können `NaN` enthalten, daher `float64` |
| `object` (string) | Textspalten, leere Werte = leerer String `""` |
| Datum | Als String gespeichert (`DD.MM.YYYY`) für Excel-Kompatibilität |

---

## Beispiel: Weiterverarbeitung

### Daten laden

```python
import pandas as pd

# Direkt aus Script
from sap_report_cleaner import run
df = run("sourceDateien/L91_Material.txt")

# Oder aus CSV
df = pd.read_csv("sourceDateien/L91_Material_cleaned.csv", sep=';', encoding='utf-8-sig')

# Oder aus Excel
df = pd.read_excel("sourceDateien/L91_Material_cleaned.xlsx", sheet_name='Bereinigte Daten')
```

### Datum konvertieren (für Zeitanalysen)

```python
# String zu datetime konvertieren
df['Pstng Date'] = pd.to_datetime(df['Pstng Date'], format='%d.%m.%Y', errors='coerce')

# Jetzt können Sie Zeitanalysen durchführen
df['Jahr'] = df['Pstng Date'].dt.year
df['Monat'] = df['Pstng Date'].dt.month
```

### Aggregationen

```python
# Summe pro Material
df.groupby('Material')['Withdrawn'].sum()

# Anzahl pro Arbeitsplatz
df.groupby('Work Ctr').size()

# Materialverbrauch pro Monat
df.groupby(df['Pstng Date'].dt.to_period('M'))['Withdrawn'].sum()
```

### Filterung

```python
# Nur Entnahmen (negative Werte)
entnahmen = df[df['Withdrawn'] < 0]

# Bestimmter Technischer Platz
trp_004 = df[df['Functional Loc.'].str.contains('TRP-004', na=False)]

# Bestimmter Zeitraum
from datetime import datetime
df_2024 = df[df['Pstng Date'].dt.year == 2024]
```

---

## Gelöschte Zeilen

Das Sheet "Gelöschte Zeilen" enthält:

| Spalte | Beschreibung |
|--------|--------------|
| **Grund** | Löschgrund: `Summenzeile` oder `Keine Materialnummer` |
| **Original_Zeile** | Zeilennummer in der Originaldatei |
| **Daten** | Original-Dateninhalt (Tab-getrennt) |

---

## Fehlerbehebung

### Problem: "openpyxl nicht gefunden"
```bash
pip3 install openpyxl
```

### Problem: Datei nicht gefunden
- Prüfen Sie den Dateipfad
- Relative Pfade beziehen sich auf das aktuelle Arbeitsverzeichnis

### Problem: Header nicht erkannt
- Das Script sucht nach "Material" in der Datei
- Falls nicht gefunden, wird Standard verwendet (Zeile 4, Spalte C)

### Problem: Zahlen werden als Text angezeigt
- Das Script konvertiert Zahlen automatisch
- Bei Problemen: CSV in Excel öffnen und Spalten manuell formatieren

---

## Konfiguration anpassen

Im Script können Sie die Spaltentypen anpassen:

```python
# Zeile ~65-72 in sap_report_cleaner.py

# Spalten die als Text bleiben sollen
TEXT_COLUMNS = ['Functional Loc.', 'Equipment', 'Material Description', 
                'Work Ctr', 'ID', 'ICt', 'Customer']

# Spalten die als Zahlen konvertiert werden
NUMERIC_COLUMNS = ['Material', 'Withdrawn', 'W/o resrv.', 'Reserved', 
                   'Reserv.ref', 'Order', 'Message']
```

---

## Lizenz

Internes Tool für SAP-Report-Bereinigung.

---

## Änderungshistorie

| Datum | Version | Änderung |
|-------|---------|----------|
| 2026-01-10 | 1.0 | Initiale Version |

