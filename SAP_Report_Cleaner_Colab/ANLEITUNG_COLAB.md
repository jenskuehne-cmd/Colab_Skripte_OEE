# SAP Report Cleaner - Google Colab Version

## Für Kollegen ohne Python-Installation

Diese Version läuft komplett im Browser über Google Colab - **keine Installation nötig!**

---

## Voraussetzungen

✅ Google-Account (haben alle in der Firma)  
✅ Internetzugang  
✅ Browser (Chrome, Firefox, Safari, Edge)  

**Nicht nötig:** Python, Installation, Admin-Rechte

---

## Schritt 1: Notebook öffnen

### Option A: Link vom Kollegen erhalten
1. Klicken Sie auf den Link, den Sie erhalten haben
2. Das Notebook öffnet sich in Google Colab

### Option B: Aus Google Drive öffnen
1. Öffnen Sie **Google Drive** (drive.google.com)
2. Navigieren Sie zum Ordner mit dem Notebook
3. **Doppelklick** auf `SAP_Report_Cleaner.ipynb`
4. Wählen Sie **"Mit Google Colaboratory öffnen"**

### Option C: Notebook hochladen
1. Öffnen Sie **Google Colab** (colab.research.google.com)
2. Klicken Sie auf **Datei → Notebook hochladen**
3. Wählen Sie die Datei `SAP_Report_Cleaner.ipynb`

---

## Schritt 2: Notebook ausführen

### 2.1 Die Code-Zelle starten

1. Suchen Sie die Zelle mit dem Code (beginnt mit "🚀 Ausführen")
2. Klicken Sie auf das **▶️ Play-Symbol** links neben der Zelle
   - Oder drücken Sie **Shift + Enter**

### 2.2 Google Drive Zugriff erlauben

1. Es erscheint ein Popup: "Notebook benötigt Zugriff auf Google Drive"
2. Klicken Sie auf **"Mit Google Drive verbinden"**
3. Wählen Sie Ihr Google-Konto
4. Klicken Sie auf **"Zulassen"**

### 2.3 Warten bis fertig

Sie sehen:
```
✅ Module geladen
✅ Funktionen geladen
📁 Verbinde mit Google Drive...
✅ Google Drive verbunden!
```

Danach erscheint die Benutzeroberfläche.

---

## Schritt 3: SAP-Report bereinigen

### 3.1 Quelldatei wählen (mit Maus-Dialog!)

**Option A: Datei vom Computer hochladen** ⭐ Empfohlen
1. Bei "Quelle" wählen Sie: **"📤 Vom Computer hochladen"**
2. Klicken Sie auf **"📤 Datei laden"**
3. **Ein Datei-Dialog öffnet sich automatisch!**
4. Navigieren Sie zu Ihrem **Downloads-Ordner**
5. Wählen Sie die SAP-Report-Datei (.txt oder .xls)
6. Klicken Sie auf **"Öffnen"**

> 💡 **Kein Pfad eintippen nötig!** Sie können mit der Maus navigieren.

**Option B: Datei aus Google Drive**
1. Bei "Quelle" wählen Sie: **"📁 Aus Google Drive wählen"**
2. Geben Sie den Pfad ein, z.B.:
   ```
   /content/drive/MyDrive/Downloads/L91_Material.txt
   ```
3. Klicken Sie auf **"📤 Datei laden"**

### 3.2 Format wählen

Wählen Sie das Ausgabeformat:
- **📊 Excel** - Enthält 2 Tabellenblätter (Daten + Gelöschte Zeilen)
- **📄 CSV** - Nur die bereinigten Daten

### 3.3 Speicherort wählen

**Option A: Auf Computer herunterladen** ⭐ Einfachste Option
1. Wählen Sie: **"💾 Auf meinen Computer herunterladen"**
2. Die Datei landet automatisch in Ihrem **Downloads-Ordner**

**Option B: In Google Drive speichern**
1. Wählen Sie: **"📁 In Google Drive speichern"**
2. Geben Sie den Drive-Pfad ein, z.B.:
   ```
   /content/drive/MyDrive/SAP_Bereinigt/
   ```
   Der Ordner wird automatisch erstellt.

### 3.4 Verarbeiten

1. Klicken Sie auf **"🚀 Verarbeiten & Speichern"**
2. Warten Sie bis "✅ Fertig!" erscheint
3. **Bei Download:** Ihr Browser lädt die Datei herunter
4. **Bei Drive:** Die Datei ist in Ihrem Google Drive

---

## Schritt 4: Bereinigte Datei finden

### Bei "Auf Computer herunterladen":
- Die Datei wird direkt heruntergeladen
- Schauen Sie in Ihrem **Downloads-Ordner**
- Dateiname: `[Originalname]_cleaned.xlsx` oder `.csv`

### Bei "In Google Drive speichern":
1. Öffnen Sie **Google Drive** (drive.google.com)
2. Navigieren Sie zum Speicherort (z.B. "SAP_Bereinigt")
3. **Doppelklick** auf die Datei zum Öffnen

---

## Häufige Fragen

### Wo finde ich meine SAP-Reports?
Nach dem Download aus SAP P30 sind die Dateien normalerweise in:
- **Windows:** `Downloads`-Ordner
- **Mac:** `Downloads`-Ordner

### Was bedeutet der Drive-Pfad?
`/content/drive/MyDrive/` = Ihr Google Drive Hauptordner

Beispiele:
- `/content/drive/MyDrive/Downloads/report.txt` = Datei im Downloads-Ordner
- `/content/drive/MyDrive/SAP/report.txt` = Datei im SAP-Ordner

### Wie sehe ich meine Drive-Ordner?
1. Klicken Sie links auf das **Ordner-Symbol** 📁
2. Navigieren Sie zu `drive → MyDrive`
3. Rechtsklick auf eine Datei → **"Pfad kopieren"**

---

## Problemlösung

| Problem | Lösung |
|---------|--------|
| Datei-Dialog öffnet sich nicht | Popup-Blocker deaktivieren für colab.google.com |
| Download startet nicht | Browser-Einstellungen prüfen, Popups erlauben |
| "Datei nicht gefunden" | Bei Drive: Pfad prüfen, Groß/Kleinschreibung beachten |
| Drive nicht verbunden | Zelle nochmal ausführen, Zugriff erlauben |
| Keine Daten in Ergebnis | Prüfen ob Datei Tab-getrennt ist (.txt) |

---

## Notebook teilen

### Für Kollegen freigeben:
1. In Colab: **Datei → In Drive speichern**
2. In Drive: **Rechtsklick → Freigeben**
3. E-Mail-Adressen der Kollegen eingeben
4. "Betrachter" oder "Bearbeiter" wählen

### Als Link teilen:
1. **Datei → Freigeben → Link abrufen**
2. "Jeder mit dem Link" wählen
3. Link kopieren und versenden

---

## Kontakt

Bei Fragen wenden Sie sich an:
- [Hier Namen/E-Mail eintragen]

---

*SAP Report Cleaner - Colab Version 1.0 - Januar 2026*

