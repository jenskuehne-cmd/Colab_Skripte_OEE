# SAP Report Cleaner - Google Sheets Version

## Für wen ist das?

Kollegen mit Google-Konto, die:
- Daten direkt in Google Sheets weiterverarbeiten möchten
- Keine Software installieren können/wollen
- Ergebnisse einfach teilen möchten

---

## Einrichtung (einmalig, 5 Minuten)

### Schritt 1: Neues Google Sheet erstellen

1. Öffnen Sie [sheets.google.com](https://sheets.google.com)
2. Klicken Sie auf **+ Leere Tabelle**
3. Benennen Sie es: **"SAP Report Cleaner"**

### Schritt 2: Apps Script öffnen

1. Im Menü: **Erweiterungen → Apps Script**
2. Ein neuer Tab öffnet sich

### Schritt 3: Code einfügen

1. **Löschen** Sie den vorhandenen Code (alles in `Code.gs`)
2. **Kopieren** Sie den gesamten Inhalt der Datei `Code.gs` aus diesem Ordner
3. **Fügen Sie ihn ein** (Strg+V)
4. Klicken Sie auf **💾 Speichern** (oder Strg+S)

### Schritt 4: Berechtigungen erteilen

1. Klicken Sie auf **▶ Ausführen** (wählen Sie `onOpen`)
2. Ein Popup erscheint: **"Autorisierung erforderlich"**
3. Klicken Sie auf **Berechtigungen überprüfen**
4. Wählen Sie Ihr Google-Konto
5. Klicken Sie auf **Erweitert** → **Zu SAP Report Cleaner (unsicher)**
6. Klicken Sie auf **Zulassen**

> ⚠️ Die Warnung erscheint, weil das Script nicht von Google verifiziert ist. Es ist sicher.

### Schritt 5: Sheet neu laden

1. Schließen Sie den Apps Script Tab
2. **Laden Sie das Google Sheet neu** (F5 oder Browser-Refresh)
3. Das Menü **"🧹 SAP Cleaner"** erscheint oben

### Fertig! ✅

---

## Verwendung (täglich)

### So bereinigen Sie einen SAP-Report:

```
1️⃣ SAP-Report herunterladen
   └─> Download aus SAP P30 (.txt Datei)

2️⃣ In Google Drive hochladen
   └─> drive.google.com → Ordner "SAP_Reports" → Datei reinziehen

3️⃣ Im Google Sheet
   └─> Menü: 🧹 SAP Cleaner → 📂 Neueste Datei bereinigen

4️⃣ Fertig!
   └─> Bereinigte Daten im Tab "Bereinigte Daten"
```

---

## Schritt-für-Schritt mit Bildern

### 1. SAP-Datei hochladen

1. Öffnen Sie **Google Drive** (drive.google.com)
2. Suchen Sie den Ordner **"SAP_Reports"**
   - Falls er nicht existiert: Beim ersten Klick auf "Bereinigen" wird er automatisch erstellt
3. Ziehen Sie Ihre SAP-Datei in den Ordner

### 2. Im Sheet bereinigen

1. Öffnen Sie das **SAP Report Cleaner** Sheet
2. Klicken Sie auf **🧹 SAP Cleaner** → **📂 Neueste Datei bereinigen**
3. Bestätigen Sie die angezeigte Datei mit **Ja**
4. Warten Sie kurz (5-30 Sekunden)
5. Ein Fenster zeigt die Statistik

### 3. Ergebnis

Die bereinigten Daten sind jetzt im Tab **"Bereinigte Daten"**:
- Mit Filterung
- Formatiert
- Bereit zur Weiterverarbeitung

---

## Das Menü

| Menüpunkt | Funktion |
|-----------|----------|
| **📂 Neueste Datei bereinigen** | Nimmt die neueste .txt/.xls Datei aus dem Ordner |
| **📋 Dateien im Ordner anzeigen** | Zeigt alle Dateien im SAP_Reports Ordner |
| **📁 Ordner in Drive öffnen** | Öffnet den Ordner direkt in Google Drive |
| **❓ Hilfe** | Zeigt Kurzanleitung |

---

## Was wird bereinigt?

| Aktion | Beschreibung |
|--------|--------------|
| Summenzeilen entfernen | Zeilen mit `*` oder `**` in Spalte B |
| Ohne Materialnummer entfernen | Zeilen ohne Wert in Spalte C (Material) |
| Ohne Abbuchung entfernen | Zeilen ohne Zahl in Spalte F (Withdrawn) |
| Zahlenformate korrigieren | Deutsche Formate (1.234,56) → Standard |
| Datumsformate konvertieren | DD.MM.YY → DD.MM.YYYY |
| Spalten filtern | Nur Spalten C bis Q (Material bis Customer) |

---

## Tabs im Sheet

| Tab | Inhalt |
|-----|--------|
| **Bereinigte Daten** | Die sauberen Daten zur Weiterverarbeitung |
| **Gelöschte Zeilen** | Welche Zeilen entfernt wurden und warum (Grund, Zeilennummer, Daten) |

---

## Für Kollegen freigeben

### Das fertige Sheet teilen:

1. **Datei → Freigeben → Für andere freigeben**
2. E-Mail-Adressen eingeben oder Link erstellen
3. Berechtigung: **Bearbeiter** (damit sie das Menü nutzen können)

### Wichtig für Kollegen:

Beim ersten Mal müssen sie:
1. Das Menü **🧹 SAP Cleaner** anklicken
2. Berechtigungen erteilen (wie oben beschrieben)

---

## Problemlösung

| Problem | Lösung |
|---------|--------|
| Menü erscheint nicht | Sheet neu laden (F5) |
| "Autorisierung erforderlich" | Siehe Schritt 4 der Einrichtung |
| "Keine Datei gefunden" | SAP-Datei in den Ordner "SAP_Reports" hochladen |
| Daten fehlen | Prüfen ob SAP-Datei Tab-getrennt ist |
| Langsam | Große Dateien brauchen bis zu 1 Minute |

---

## Technische Details

| Eigenschaft | Wert |
|-------------|------|
| Sprache | Google Apps Script (JavaScript) |
| Max. Laufzeit | 6 Minuten |
| Max. Zeilen | ~50.000 (Sheet-Limit) |
| Speicherort | Google Drive, Ordner "SAP_Reports" |

---

*SAP Report Cleaner v1.0 - Google Sheets - Januar 2026*

