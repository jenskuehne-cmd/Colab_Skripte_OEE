# Colab Skripte OEE

Sammlung von Notebooks und Tools für OEE-Analysen und SAP-Datenverarbeitung.

---

## SAP Report Cleaner

Tool zum Bereinigen von SAP-Reports (Tab-getrennte Pseudo-XLS-Dateien).

| Version | Ordner | Beschreibung |
|---------|--------|--------------|
| **Lokal (Python)** | [`SAP_Report_Cleaner/`](SAP_Report_Cleaner/) | Python-Script mit GUI für Windows/macOS |
| **Google Colab** | [`SAP_Report_Cleaner_Colab/`](SAP_Report_Cleaner_Colab/) | Browser-basiert, kein Python nötig |
| **Google Sheets** | [`SAP_Report_Cleaner_GoogleSheets/`](SAP_Report_Cleaner_GoogleSheets/) | Direkt in Google Sheets integriert |
| **Windows EXE** | [`SAP_Report_Cleaner_Windows/`](SAP_Report_Cleaner_Windows/) | Standalone-Programm für Windows |

### Welche Version wählen?

- **Python installiert?** → `SAP_Report_Cleaner/` (Doppelklick auf `.bat` oder `.command`)
- **Kein Python, aber Google-Konto?** → `SAP_Report_Cleaner_Colab/` oder `SAP_Report_Cleaner_GoogleSheets/`
- **Windows ohne Python?** → `SAP_Report_Cleaner_Windows/` (EXE erstellen lassen)

### Weitergabe an Kollegen

Fertiges ZIP-Paket zur Weitergabe: **`SAP_Report_Cleaner/dist/SAP_Report_Cleaner.zip`**

Einfach ZIP senden → Empfänger entpackt → Doppelklick auf Starter

---

## Weitere Notebooks

| Ordner | Beschreibung |
|--------|--------------|
| [`Notebooks/`](Notebooks/) | OEE Playground und weitere Analyse-Notebooks |
| [`Notebooks/RoleMappingAspire/`](Notebooks/RoleMappingAspire/) | Role Mapping Daten-Analysen |
| [`Notebooks/SAP_P30/`](Notebooks/SAP_P30/) | SAP P30 spezifische Notebooks |

---

## Schnellstart

```bash
# Lokale Python-Version starten (macOS)
cd SAP_Report_Cleaner
./SAP_Report_Cleaner.command

# Lokale Python-Version starten (Windows)
cd SAP_Report_Cleaner
SAP_Report_Cleaner.bat
```

Detaillierte Anleitungen befinden sich in den jeweiligen Unterordnern.
