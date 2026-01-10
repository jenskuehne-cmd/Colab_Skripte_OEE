#!/usr/bin/env python3
"""
SAP Report Cleaner
==================
Bereinigt SAP-Reports, die als Tab-getrennte Textdateien vorliegen.

Funktionen:
- Liest Tab-getrennte TXT/XLS-Dateien
- Entfernt Datumsinformationen aus Zeile 1
- Filtert nur Spalten C bis Q (Material bis Customer)
- Entfernt Summenzeilen (markiert mit "*" in Spalte B)
- Entfernt leere Zeilen
- Entfernt Zeilen ohne Materialnummer in Spalte C
- Bereinigt Zahlenformate für pandas
- Konvertiert Datumsformate (DD.MM.YY)
- Exportiert als CSV und Excel (mit gelöschten Zeilen)

Verwendung:
    python3 sap_report_cleaner.py [dateipfad]
    
    Ohne Argument: Interaktive Dateiauswahl
    Mit Argument: Direkte Verarbeitung der angegebenen Datei

Voraussetzungen:
    pip3 install pandas openpyxl
"""

import os
import sys

# Tk Deprecation-Warnung unterdrücken (macOS)
os.environ['TK_SILENCE_DEPRECATION'] = '1'

import pandas as pd
import numpy as np
import re


from pathlib import Path
from datetime import datetime

# ============================================================================
# KONFIGURATION
# ============================================================================

# Erwartete Spaltenüberschriften (C bis Q)
EXPECTED_HEADERS = [
    'Material',           # C - Text (Materialnummer)
    'Functional Loc.',    # D - Text
    'Equipment',          # E - Text
    'Material Description', # F - Text
    'Work Ctr',           # G - Text
    'Withdrawn',          # H - Integer
    'W/o resrv.',         # I - Integer
    'Reserved',           # J - Integer
    'Reserv.ref',         # K - Integer
    'Pstng Date',         # L - Datum (DD.MM.YY)
    'Order',              # M - Integer
    'ID',                 # N - Text
    'Message',            # O - Integer
    'ICt',                # P - Text
    'Customer'            # Q - Text
]

# Spaltentypen
# Hinweis: Material als Text = führende Nullen bleiben erhalten
# Ändern Sie hier, wenn Material als Zahl gewünscht ist
TEXT_COLUMNS = ['Functional Loc.', 'Equipment', 'Material Description', 
                'Work Ctr', 'ID', 'ICt', 'Customer']
# Material wird jetzt als Zahl behandelt (wenn rein numerisch)
DATE_COLUMN = 'Pstng Date'
NUMERIC_COLUMNS = ['Material', 'Withdrawn', 'W/o resrv.', 'Reserved', 'Reserv.ref', 'Order', 'Message']


# ============================================================================
# HILFSFUNKTIONEN
# ============================================================================

def install_openpyxl():
    """Installiert openpyxl falls nicht vorhanden."""
    try:
        import openpyxl
        return True
    except ImportError:
        print("⚠ openpyxl nicht gefunden. Installiere...")
        import subprocess
        methods = [
            [sys.executable, "-m", "pip", "install", "openpyxl", "-q"],
            ["pip3", "install", "openpyxl", "-q"],
            ["pip", "install", "openpyxl", "-q"],
        ]
        for method in methods:
            try:
                subprocess.check_call(method, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                import openpyxl
                print("✓ openpyxl installiert")
                return True
            except (subprocess.CalledProcessError, FileNotFoundError, ImportError):
                continue
        print("⚠ openpyxl konnte nicht installiert werden. Excel-Export nicht möglich.")
        return False


def clean_number(value):
    """
    Bereinigt einen Zahlenwert aus SAP-Format.
    - Entfernt Leerzeichen
    - Behandelt Tausenderpunkte
    - Behandelt Dezimalkommas
    - Gibt Integer zurück
    """
    if pd.isna(value) or value is None:
        return None
    
    val_str = str(value).strip()
    if val_str == '' or val_str == '-':
        return None
    
    # Entferne führende/folgende Leerzeichen und Nicht-Zahlen-Zeichen (außer .,-)
    val_str = val_str.replace('\xa0', '').replace(' ', '')
    
    # Prüfe auf deutsches Zahlenformat (1.234,56)
    if ',' in val_str and '.' in val_str:
        # 1.234,56 -> 1234.56
        val_str = val_str.replace('.', '').replace(',', '.')
    elif ',' in val_str:
        # Nur Komma: könnte Dezimalkomma sein (1,5) oder Tausender (1,234)
        parts = val_str.split(',')
        if len(parts) == 2 and len(parts[1]) <= 2:
            # Wahrscheinlich Dezimalkomma
            val_str = val_str.replace(',', '.')
        else:
            # Wahrscheinlich Tausendertrennzeichen
            val_str = val_str.replace(',', '')
    elif '.' in val_str:
        # Nur Punkt: prüfe ob Dezimalpunkt oder Tausenderpunkt
        parts = val_str.split('.')
        if len(parts) == 2 and len(parts[1]) == 3 and len(parts[0]) >= 1:
            # z.B. "3.500" ist wahrscheinlich 3500 (Tausenderpunkt)
            val_str = val_str.replace('.', '')
    
    try:
        # Versuche als Float zu parsen und zu Integer zu konvertieren
        num = float(val_str)
        return int(round(num))
    except ValueError:
        return None


def convert_date(value, output_format='%d.%m.%Y'):
    """
    Konvertiert Datum aus SAP-Format (DD.MM.YY oder DD.MM.YYYY).
    Gibt einen formatierten String zurück für bessere Excel-Kompatibilität.
    """
    if pd.isna(value) or value is None:
        return ''
    
    val_str = str(value).strip()
    if val_str == '':
        return ''
    
    # Versuche verschiedene Datumsformate
    input_formats = ['%d.%m.%y', '%d.%m.%Y', '%Y-%m-%d']
    
    for fmt in input_formats:
        try:
            parsed_date = datetime.strptime(val_str, fmt)
            # Rückgabe als formatierter String (DD.MM.YYYY)
            return parsed_date.strftime(output_format)
        except ValueError:
            continue
    
    return val_str  # Original zurückgeben wenn Parsing fehlschlägt


def read_sap_file(file_path):
    """
    Liest eine SAP-Report-Datei (Tab-getrennt).
    Gibt alle Zeilen als Liste von Listen zurück.
    """
    print(f"\n📂 Lese Datei: {file_path}")
    
    with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
        content = f.read()
    
    lines = content.split('\n')
    print(f"   Gefunden: {len(lines)} Zeilen")
    
    # In Spalten aufteilen
    all_rows = []
    for line in lines:
        columns = line.split('\t')
        all_rows.append(columns)
    
    return all_rows


def find_header_row(rows):
    """
    Findet die Zeile mit den Spaltenüberschriften.
    Sucht nach 'Material' als erstem Header in Spalte C.
    """
    for idx, row in enumerate(rows):
        # Suche in der Zeile nach "Material" (sollte in Spalte C sein, also Index 2)
        for col_idx, cell in enumerate(row):
            cell_clean = str(cell).strip().lower()
            if cell_clean == 'material':
                print(f"   Header gefunden in Zeile {idx + 1}, Spalte {col_idx + 1}")
                return idx, col_idx
    
    return None, None


def process_sap_report(file_path):
    """
    Hauptfunktion: Verarbeitet eine SAP-Report-Datei.
    """
    # Datei einlesen
    all_rows = read_sap_file(file_path)
    
    if not all_rows:
        print("❌ Fehler: Datei ist leer")
        return None, None
    
    # Header-Zeile finden
    header_row_idx, header_start_col = find_header_row(all_rows)
    
    if header_row_idx is None:
        print("⚠ Warnung: Header-Zeile nicht automatisch gefunden")
        print("   Verwende Standard: Zeile 4, Spalte C (Index 2)")
        header_row_idx = 3  # 0-basiert, also Zeile 4
        header_start_col = 2  # Spalte C
    
    # Header extrahieren
    header_row = all_rows[header_row_idx]
    
    # Relevante Spalten: C bis Q = Index 2 bis 16 (15 Spalten)
    # Stelle sicher, dass genug Spalten vorhanden sind
    num_expected_cols = 15  # C bis Q
    
    # Extrahiere Header für Spalten C-Q
    extracted_headers = []
    for i in range(header_start_col, header_start_col + num_expected_cols):
        if i < len(header_row):
            extracted_headers.append(str(header_row[i]).strip())
        else:
            extracted_headers.append(f'Col_{i}')
    
    print(f"\n📋 Extrahierte Header: {extracted_headers}")
    
    # Verwende erwartete Header für Konsistenz
    print(f"   Verwende Standard-Header: {EXPECTED_HEADERS}")
    
    # Daten sammeln (nach Header-Zeile)
    cleaned_data = []
    deleted_rows = []
    
    stats = {
        'total_rows': 0,
        'sum_rows': 0,
        'empty_rows': 0,
        'no_material': 0,
        'kept_rows': 0
    }
    
    for row_idx in range(header_row_idx + 1, len(all_rows)):
        row = all_rows[row_idx]
        stats['total_rows'] += 1
        
        # Prüfe auf komplett leere Zeile
        if all(str(cell).strip() == '' for cell in row):
            stats['empty_rows'] += 1
            continue
        
        # Hole Spalte B (Index 1) für Summenzeilen-Prüfung
        col_b = str(row[1]).strip() if len(row) > 1 else ''
        
        # Prüfe auf Summenzeile (markiert mit * in Spalte B)
        if col_b == '*' or col_b == '**':
            stats['sum_rows'] += 1
            # Speichere mit Grund
            deleted_rows.append({
                'Grund': 'Summenzeile',
                'Original_Zeile': row_idx + 1,
                'Daten': '\t'.join(str(c) for c in row)
            })
            continue
        
        # Extrahiere Spalten C bis Q
        data_row = []
        for i in range(header_start_col, header_start_col + num_expected_cols):
            if i < len(row):
                data_row.append(str(row[i]).strip())
            else:
                data_row.append('')
        
        # Prüfe auf Materialnummer in Spalte C (erstes Element)
        material_nr = data_row[0] if data_row else ''
        if not material_nr:
            stats['no_material'] += 1
            deleted_rows.append({
                'Grund': 'Keine Materialnummer',
                'Original_Zeile': row_idx + 1,
                'Daten': '\t'.join(data_row)
            })
            continue
        
        # Zeile behalten
        cleaned_data.append(data_row)
        stats['kept_rows'] += 1
    
    print(f"\n📊 Statistik:")
    print(f"   Gesamt Zeilen:     {stats['total_rows']}")
    print(f"   Summenzeilen:      {stats['sum_rows']} (gelöscht)")
    print(f"   Leere Zeilen:      {stats['empty_rows']} (gelöscht)")
    print(f"   Ohne Materialnr:   {stats['no_material']} (gelöscht)")
    print(f"   Bereinigte Zeilen: {stats['kept_rows']}")
    
    # DataFrame erstellen
    df = pd.DataFrame(cleaned_data, columns=EXPECTED_HEADERS)
    df_deleted = pd.DataFrame(deleted_rows)
    
    print(f"\n✅ DataFrame erstellt: {df.shape[0]} Zeilen, {df.shape[1]} Spalten")
    
    return df, df_deleted


def convert_data_types(df):
    """
    Konvertiert Spalten in die korrekten Datentypen.
    """
    print("\n🔄 Konvertiere Datentypen...")
    
    # Numerische Spalten
    for col in NUMERIC_COLUMNS:
        if col in df.columns:
            df[col] = df[col].apply(clean_number)
    
    # Datum-Spalte
    if DATE_COLUMN in df.columns:
        df[DATE_COLUMN] = df[DATE_COLUMN].apply(convert_date)
    
    # Text-Spalten bleiben wie sie sind
    for col in TEXT_COLUMNS:
        if col in df.columns:
            df[col] = df[col].astype(str).replace('nan', '').replace('None', '')
    
    print("   ✓ Datentypen konvertiert")
    return df


def export_results(df, df_deleted, input_file):
    """
    Exportiert die Ergebnisse als CSV und Excel.
    """
    input_path = Path(input_file)
    base_name = input_path.stem
    output_dir = input_path.parent
    
    # CSV Export
    csv_path = output_dir / f"{base_name}_cleaned.csv"
    df.to_csv(csv_path, index=False, sep=';', encoding='utf-8-sig')
    print(f"\n💾 CSV exportiert: {csv_path}")
    
    # Excel Export
    if install_openpyxl():
        excel_path = output_dir / f"{base_name}_cleaned.xlsx"
        try:
            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Bereinigte Daten', index=False)
                if not df_deleted.empty:
                    df_deleted.to_excel(writer, sheet_name='Gelöschte Zeilen', index=False)
            print(f"💾 Excel exportiert: {excel_path}")
        except Exception as e:
            print(f"⚠ Excel-Export fehlgeschlagen: {e}")
            # Fallback: Gelöschte Zeilen als separate CSV
            deleted_csv = output_dir / f"{base_name}_deleted.csv"
            df_deleted.to_csv(deleted_csv, index=False, sep=';', encoding='utf-8-sig')
            print(f"💾 Gelöschte Zeilen als CSV: {deleted_csv}")
    else:
        # Fallback: Gelöschte Zeilen als separate CSV
        deleted_csv = output_dir / f"{base_name}_deleted.csv"
        df_deleted.to_csv(deleted_csv, index=False, sep=';', encoding='utf-8-sig')
        print(f"💾 Gelöschte Zeilen als CSV: {deleted_csv}")
    
    return csv_path


def select_file():
    """
    Dateiauswahl (interaktiv oder per Argument).
    """
    # Prüfe Kommandozeilenargument
    if len(sys.argv) > 1:
        file_path = Path(sys.argv[1])
        if file_path.exists():
            return str(file_path.resolve())
        else:
            print(f"❌ Datei nicht gefunden: {sys.argv[1]}")
            raise FileNotFoundError(f"Datei nicht gefunden: {sys.argv[1]}")
    
    # Versuche tkinter Dialog
    try:
        # Tk Deprecation-Warnung unterdrücken (macOS)
        os.environ['TK_SILENCE_DEPRECATION'] = '1'
        
        # Stderr temporär unterdrücken während tkinter-Import
        import io
        import contextlib
        
        stderr_backup = sys.stderr
        sys.stderr = io.StringIO()
        
        try:
            import tkinter as tk
            from tkinter import filedialog
        finally:
            sys.stderr = stderr_backup
        
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        
        file_path = filedialog.askopenfilename(
            title="SAP-Report auswählen",
            filetypes=[
                ("Text/Excel Dateien", "*.txt *.xls"),
                ("Textdateien", "*.txt"),
                ("Excel-Dateien", "*.xls *.xlsx"),
                ("Alle Dateien", "*.*")
            ]
        )
        
        root.destroy()
        
        if file_path:
            return file_path
    except ImportError:
        pass
    
    # Fallback: Manuelle Eingabe
    print("\n📂 Bitte Dateipfad eingeben (oder 'q' zum Abbrechen):")
    file_path = input("   Pfad: ").strip()
    
    if file_path.lower() == 'q':
        print("⚠ Abgebrochen.")
        return None
    
    if file_path.startswith('"') and file_path.endswith('"'):
        file_path = file_path[1:-1]
    
    if not Path(file_path).exists():
        print(f"❌ Datei nicht gefunden: {file_path}")
        return None
    
    return file_path


def run(file_path=None):
    """
    Hauptfunktion - kann auch direkt mit Dateipfad aufgerufen werden.
    
    Beispiel:
        from sap_report_cleaner import run
        df = run("sourceDateien/L91_Material.txt")
    """
    print("=" * 60)
    print("  SAP Report Cleaner")
    print("=" * 60)
    
    # Datei auswählen
    if file_path is None:
        file_path = select_file()
        if file_path is None:
            return None
    else:
        # Prüfe ob Datei existiert
        if not Path(file_path).exists():
            print(f"❌ Datei nicht gefunden: {file_path}")
            return None
        file_path = str(Path(file_path).resolve())
    
    # Verarbeiten
    df, df_deleted = process_sap_report(file_path)
    
    if df is None:
        print("❌ Verarbeitung fehlgeschlagen")
        return None
    
    # Datentypen konvertieren
    df = convert_data_types(df)
    
    # Vorschau
    print("\n📋 Vorschau (erste 5 Zeilen):")
    print(df.head().to_string())
    
    # Exportieren
    export_results(df, df_deleted, file_path)
    
    print("\n" + "=" * 60)
    print("  ✅ Fertig!")
    print("=" * 60)
    
    return df


def main():
    """Kommandozeilen-Einstiegspunkt."""
    try:
        df = run()
        if df is None:
            sys.exit(1)
    except KeyboardInterrupt:
        print("\n⚠ Abgebrochen.")
        sys.exit(0)


if __name__ == "__main__":
    main()

