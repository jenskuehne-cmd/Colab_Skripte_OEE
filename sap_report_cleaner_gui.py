#!/usr/bin/env python3
"""
SAP Report Cleaner (GUI Version)
================================
Bereinigt SAP-Reports mit grafischer Dateiauswahl.
Funktioniert auf macOS und Windows.

Verwendung:
    python3 sap_report_cleaner_gui.py
    
    1. Quelldatei auswählen (Finder/Explorer Dialog)
    2. Zieldatei wählen (Speicherort und Name)
    3. Fertig!

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
# GUI-FUNKTIONEN
# ============================================================================

def init_tkinter():
    """Initialisiert tkinter und gibt root zurück."""
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox
        
        root = tk.Tk()
        root.withdraw()  # Hauptfenster verstecken
        
        # Fenster in den Vordergrund bringen (wichtig für macOS)
        root.attributes('-topmost', True)
        root.update()
        
        return root, filedialog, messagebox
    except ImportError:
        print("❌ Fehler: tkinter nicht verfügbar.")
        print("   Bitte installieren oder sap_report_cleaner.py verwenden.")
        sys.exit(1)


def select_source_file(filedialog):
    """Öffnet Dialog zur Auswahl der Quelldatei."""
    print("\n📂 Bitte Quelldatei auswählen...")
    
    file_path = filedialog.askopenfilename(
        title="SAP-Report auswählen (Quelldatei)",
        filetypes=[
            ("SAP Reports", "*.txt *.xls"),
            ("Textdateien", "*.txt"),
            ("Excel-Dateien", "*.xls *.xlsx"),
            ("Alle Dateien", "*.*")
        ],
        initialdir=os.getcwd()
    )
    
    if not file_path:
        return None
    
    print(f"   ✓ Ausgewählt: {file_path}")
    return file_path


def select_target_file(filedialog, source_path):
    """Öffnet Dialog zur Auswahl des Speicherorts."""
    print("\n💾 Bitte Speicherort wählen...")
    
    # Vorgeschlagener Dateiname basierend auf Quelldatei
    source_name = Path(source_path).stem
    default_name = f"{source_name}_cleaned"
    source_dir = str(Path(source_path).parent)
    
    file_path = filedialog.asksaveasfilename(
        title="Bereinigte Datei speichern als",
        filetypes=[
            ("Excel-Datei", "*.xlsx"),
            ("CSV-Datei", "*.csv"),
        ],
        defaultextension=".xlsx",
        initialfile=default_name,
        initialdir=source_dir
    )
    
    if not file_path:
        return None
    
    print(f"   ✓ Speichern unter: {file_path}")
    return file_path


def show_success_message(messagebox, output_path, stats):
    """Zeigt Erfolgsmeldung an."""
    message = (
        f"Verarbeitung abgeschlossen!\n\n"
        f"Bereinigte Zeilen: {stats['kept_rows']}\n"
        f"Gelöschte Zeilen: {stats['sum_rows'] + stats['no_material']}\n"
        f"  - Summenzeilen: {stats['sum_rows']}\n"
        f"  - Ohne Material: {stats['no_material']}\n\n"
        f"Gespeichert unter:\n{output_path}"
    )
    messagebox.showinfo("SAP Report Cleaner", message)


def show_error_message(messagebox, error):
    """Zeigt Fehlermeldung an."""
    messagebox.showerror("Fehler", str(error))


# ============================================================================
# KONFIGURATION
# ============================================================================

EXPECTED_HEADERS = [
    'Material', 'Functional Loc.', 'Equipment', 'Material Description',
    'Work Ctr', 'Withdrawn', 'W/o resrv.', 'Reserved', 'Reserv.ref',
    'Pstng Date', 'Order', 'ID', 'Message', 'ICt', 'Customer'
]

TEXT_COLUMNS = ['Functional Loc.', 'Equipment', 'Material Description', 
                'Work Ctr', 'ID', 'ICt', 'Customer']
DATE_COLUMN = 'Pstng Date'
NUMERIC_COLUMNS = ['Material', 'Withdrawn', 'W/o resrv.', 'Reserved', 
                   'Reserv.ref', 'Order', 'Message']


# ============================================================================
# DATENVERARBEITUNGS-FUNKTIONEN
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
        return False


def clean_number(value):
    """Bereinigt einen Zahlenwert aus SAP-Format."""
    if pd.isna(value) or value is None:
        return None
    
    val_str = str(value).strip()
    if val_str == '' or val_str == '-':
        return None
    
    val_str = val_str.replace('\xa0', '').replace(' ', '')
    
    if ',' in val_str and '.' in val_str:
        val_str = val_str.replace('.', '').replace(',', '.')
    elif ',' in val_str:
        parts = val_str.split(',')
        if len(parts) == 2 and len(parts[1]) <= 2:
            val_str = val_str.replace(',', '.')
        else:
            val_str = val_str.replace(',', '')
    elif '.' in val_str:
        parts = val_str.split('.')
        if len(parts) == 2 and len(parts[1]) == 3 and len(parts[0]) >= 1:
            val_str = val_str.replace('.', '')
    
    try:
        num = float(val_str)
        return int(round(num))
    except ValueError:
        return None


def convert_date(value, output_format='%d.%m.%Y'):
    """Konvertiert Datum aus SAP-Format."""
    if pd.isna(value) or value is None:
        return ''
    
    val_str = str(value).strip()
    if val_str == '':
        return ''
    
    input_formats = ['%d.%m.%y', '%d.%m.%Y', '%Y-%m-%d']
    
    for fmt in input_formats:
        try:
            parsed_date = datetime.strptime(val_str, fmt)
            return parsed_date.strftime(output_format)
        except ValueError:
            continue
    
    return val_str


def read_sap_file(file_path):
    """Liest eine SAP-Report-Datei."""
    print(f"\n📂 Lese Datei: {Path(file_path).name}")
    
    with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
        content = f.read()
    
    lines = content.split('\n')
    print(f"   Gefunden: {len(lines)} Zeilen")
    
    all_rows = []
    for line in lines:
        columns = line.split('\t')
        all_rows.append(columns)
    
    return all_rows


def find_header_row(rows):
    """Findet die Header-Zeile."""
    for idx, row in enumerate(rows):
        for col_idx, cell in enumerate(row):
            cell_clean = str(cell).strip().lower()
            if cell_clean == 'material':
                print(f"   Header gefunden in Zeile {idx + 1}")
                return idx, col_idx
    return None, None


def process_sap_report(file_path):
    """Verarbeitet eine SAP-Report-Datei."""
    all_rows = read_sap_file(file_path)
    
    if not all_rows:
        raise ValueError("Datei ist leer")
    
    header_row_idx, header_start_col = find_header_row(all_rows)
    
    if header_row_idx is None:
        print("⚠ Header nicht gefunden, verwende Standard")
        header_row_idx = 3
        header_start_col = 2
    
    num_expected_cols = 15
    
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
        
        if all(str(cell).strip() == '' for cell in row):
            stats['empty_rows'] += 1
            continue
        
        col_b = str(row[1]).strip() if len(row) > 1 else ''
        
        if col_b == '*' or col_b == '**':
            stats['sum_rows'] += 1
            deleted_rows.append({
                'Grund': 'Summenzeile',
                'Original_Zeile': row_idx + 1,
                'Daten': '\t'.join(str(c) for c in row)
            })
            continue
        
        data_row = []
        for i in range(header_start_col, header_start_col + num_expected_cols):
            if i < len(row):
                data_row.append(str(row[i]).strip())
            else:
                data_row.append('')
        
        material_nr = data_row[0] if data_row else ''
        if not material_nr:
            stats['no_material'] += 1
            deleted_rows.append({
                'Grund': 'Keine Materialnummer',
                'Original_Zeile': row_idx + 1,
                'Daten': '\t'.join(data_row)
            })
            continue
        
        cleaned_data.append(data_row)
        stats['kept_rows'] += 1
    
    print(f"\n📊 Statistik:")
    print(f"   Bereinigte Zeilen: {stats['kept_rows']}")
    print(f"   Summenzeilen:      {stats['sum_rows']} (gelöscht)")
    print(f"   Ohne Materialnr:   {stats['no_material']} (gelöscht)")
    
    df = pd.DataFrame(cleaned_data, columns=EXPECTED_HEADERS)
    df_deleted = pd.DataFrame(deleted_rows)
    
    return df, df_deleted, stats


def convert_data_types(df):
    """Konvertiert Spalten in korrekte Datentypen."""
    print("\n🔄 Konvertiere Datentypen...")
    
    for col in NUMERIC_COLUMNS:
        if col in df.columns:
            df[col] = df[col].apply(clean_number)
    
    if DATE_COLUMN in df.columns:
        df[DATE_COLUMN] = df[DATE_COLUMN].apply(convert_date)
    
    for col in TEXT_COLUMNS:
        if col in df.columns:
            df[col] = df[col].astype(str).replace('nan', '').replace('None', '')
    
    print("   ✓ Datentypen konvertiert")
    return df


def export_results(df, df_deleted, output_path):
    """Exportiert die Ergebnisse."""
    output_path = Path(output_path)
    
    # Bestimme Format basierend auf Dateiendung
    if output_path.suffix.lower() == '.csv':
        # CSV Export
        df.to_csv(output_path, index=False, sep=';', encoding='utf-8-sig')
        print(f"\n💾 CSV exportiert: {output_path}")
        
        # Gelöschte Zeilen als separate CSV
        if not df_deleted.empty:
            deleted_path = output_path.parent / f"{output_path.stem}_deleted.csv"
            df_deleted.to_csv(deleted_path, index=False, sep=';', encoding='utf-8-sig')
            print(f"💾 Gelöschte Zeilen: {deleted_path}")
    else:
        # Excel Export
        if install_openpyxl():
            try:
                with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Bereinigte Daten', index=False)
                    if not df_deleted.empty:
                        df_deleted.to_excel(writer, sheet_name='Gelöschte Zeilen', index=False)
                print(f"\n💾 Excel exportiert: {output_path}")
            except Exception as e:
                # Fallback zu CSV
                csv_path = output_path.with_suffix('.csv')
                df.to_csv(csv_path, index=False, sep=';', encoding='utf-8-sig')
                print(f"\n💾 CSV exportiert (Excel-Export fehlgeschlagen): {csv_path}")
                return csv_path
        else:
            # Fallback zu CSV
            csv_path = output_path.with_suffix('.csv')
            df.to_csv(csv_path, index=False, sep=';', encoding='utf-8-sig')
            print(f"\n💾 CSV exportiert (openpyxl nicht verfügbar): {csv_path}")
            return csv_path
    
    return output_path


# ============================================================================
# HAUPTPROGRAMM
# ============================================================================

def main():
    """Hauptprogramm mit GUI-Dialogen."""
    print("=" * 60)
    print("  SAP Report Cleaner (GUI Version)")
    print("=" * 60)
    
    # tkinter initialisieren
    root, filedialog, messagebox = init_tkinter()
    
    try:
        # 1. Quelldatei auswählen
        source_path = select_source_file(filedialog)
        if not source_path:
            print("\n⚠ Keine Datei ausgewählt. Abbruch.")
            return
        
        # 2. Zieldatei wählen
        target_path = select_target_file(filedialog, source_path)
        if not target_path:
            print("\n⚠ Kein Speicherort gewählt. Abbruch.")
            return
        
        # 3. Verarbeiten
        df, df_deleted, stats = process_sap_report(source_path)
        
        # 4. Datentypen konvertieren
        df = convert_data_types(df)
        
        # 5. Exportieren
        output_path = export_results(df, df_deleted, target_path)
        
        # 6. Erfolgsmeldung
        print("\n" + "=" * 60)
        print("  ✅ Fertig!")
        print("=" * 60)
        
        show_success_message(messagebox, str(output_path), stats)
        
    except Exception as e:
        print(f"\n❌ Fehler: {e}")
        show_error_message(messagebox, e)
    
    finally:
        root.destroy()


if __name__ == "__main__":
    main()

