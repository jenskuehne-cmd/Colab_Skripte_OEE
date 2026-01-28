#!/usr/bin/env python3
"""
SAP Report Cleaner - Windows Version
=====================================
Bereinigt SAP-Reports in 3 einfachen Schritten.
Speichert automatisch in den Downloads-Ordner.
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from datetime import datetime

# Pandas und openpyxl beim Start prüfen
try:
    import pandas as pd
    import numpy as np
except ImportError:
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("Fehler", 
        "pandas ist nicht installiert.\n\n"
        "Bitte installieren Sie es mit:\n"
        "pip install pandas numpy openpyxl")
    sys.exit(1)

# ============================================================
# KONFIGURATION
# ============================================================

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

# ============================================================
# HILFSFUNKTIONEN
# ============================================================

def get_downloads_folder():
    """Gibt den Downloads-Ordner des Benutzers zurück."""
    if sys.platform == 'win32':
        import winreg
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders")
            downloads = winreg.QueryValueEx(key, "{374DE290-123F-4565-9164-39C4925E467B}")[0]
            return Path(downloads)
        except:
            pass
    # Fallback
    return Path.home() / "Downloads"

def clean_number(value):
    """Bereinigt Zahlenwerte aus SAP-Format."""
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
        return int(round(float(val_str)))
    except ValueError:
        return None

def convert_date(value):
    """Konvertiert Datum aus SAP-Format."""
    if pd.isna(value) or value is None:
        return ''
    val_str = str(value).strip()
    if val_str == '':
        return ''
    for fmt in ['%d.%m.%y', '%d.%m.%Y', '%Y-%m-%d']:
        try:
            return datetime.strptime(val_str, fmt).strftime('%d.%m.%Y')
        except ValueError:
            continue
    return val_str

def process_sap_report(file_path):
    """Verarbeitet eine SAP-Report-Datei."""
    # Datei lesen
    with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
        content = f.read()
    
    lines = content.split('\n')
    all_rows = [line.split('\t') for line in lines]

    # Header finden
    header_row_idx, header_start_col = None, None
    for idx, row in enumerate(all_rows):
        for col_idx, cell in enumerate(row):
            if str(cell).strip().lower() == 'material':
                header_row_idx, header_start_col = idx, col_idx
                break
        if header_row_idx is not None:
            break

    if header_row_idx is None:
        header_row_idx, header_start_col = 3, 2

    # Daten verarbeiten
    cleaned_data, deleted_rows = [], []
    stats = {'total': 0, 'sum_rows': 0, 'empty': 0, 'no_material': 0, 'kept': 0}

    for row_idx in range(header_row_idx + 1, len(all_rows)):
        row = all_rows[row_idx]
        stats['total'] += 1

        if all(str(cell).strip() == '' for cell in row):
            stats['empty'] += 1
            continue

        col_b = str(row[1]).strip() if len(row) > 1 else ''
        if col_b in ['*', '**']:
            stats['sum_rows'] += 1
            deleted_rows.append({'Grund': 'Summenzeile', 'Zeile': row_idx + 1,
                                 'Daten': '\t'.join(str(c) for c in row)})
            continue

        data_row = [str(row[i]).strip() if i < len(row) else ''
                    for i in range(header_start_col, header_start_col + 15)]

        if not data_row[0]:
            stats['no_material'] += 1
            deleted_rows.append({'Grund': 'Keine Materialnummer', 'Zeile': row_idx + 1,
                                 'Daten': '\t'.join(data_row)})
            continue

        cleaned_data.append(data_row)
        stats['kept'] += 1

    # DataFrame erstellen
    df = pd.DataFrame(cleaned_data, columns=EXPECTED_HEADERS)
    df_deleted = pd.DataFrame(deleted_rows)

    # Datentypen konvertieren
    for col in NUMERIC_COLUMNS:
        if col in df.columns:
            df[col] = df[col].apply(clean_number)
    if DATE_COLUMN in df.columns:
        df[DATE_COLUMN] = df[DATE_COLUMN].apply(convert_date)

    return df, df_deleted, stats

# ============================================================
# HAUPTFENSTER
# ============================================================

class SAPReportCleanerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("SAP Report Cleaner")
        self.root.geometry("500x400")
        self.root.resizable(False, False)
        
        # Variablen
        self.source_file = None
        self.result_df = None
        self.result_deleted = None
        self.format_var = tk.StringVar(value="excel")
        
        self.create_widgets()
    
    def create_widgets(self):
        # Hauptframe
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Titel
        title = ttk.Label(main_frame, text="🧹 SAP Report Cleaner", 
                         font=('Segoe UI', 18, 'bold'))
        title.pack(pady=(0, 5))
        
        subtitle = ttk.Label(main_frame, text="Bereinigt SAP-Reports in 3 Schritten",
                            font=('Segoe UI', 10))
        subtitle.pack(pady=(0, 20))
        
        # Separator
        ttk.Separator(main_frame).pack(fill=tk.X, pady=10)
        
        # Schritt 1
        step1_frame = ttk.LabelFrame(main_frame, text="Schritt 1: SAP-Datei auswählen", 
                                     padding="10")
        step1_frame.pack(fill=tk.X, pady=5)
        
        self.file_label = ttk.Label(step1_frame, text="Keine Datei ausgewählt",
                                    foreground='gray')
        self.file_label.pack(side=tk.LEFT, expand=True, fill=tk.X)
        
        self.select_btn = ttk.Button(step1_frame, text="📂 Datei wählen",
                                     command=self.select_file)
        self.select_btn.pack(side=tk.RIGHT)
        
        # Schritt 2
        step2_frame = ttk.LabelFrame(main_frame, text="Schritt 2: Ausgabeformat", 
                                     padding="10")
        step2_frame.pack(fill=tk.X, pady=5)
        
        ttk.Radiobutton(step2_frame, text="📊 Excel (mit gelöschten Zeilen)", 
                       variable=self.format_var, value="excel").pack(anchor=tk.W)
        ttk.Radiobutton(step2_frame, text="📄 CSV (nur bereinigte Daten)", 
                       variable=self.format_var, value="csv").pack(anchor=tk.W)
        
        # Schritt 3
        step3_frame = ttk.LabelFrame(main_frame, text="Schritt 3: Bereinigen & Speichern", 
                                     padding="10")
        step3_frame.pack(fill=tk.X, pady=5)
        
        info_label = ttk.Label(step3_frame, 
                              text="Die bereinigte Datei wird in Ihrem Downloads-Ordner gespeichert",
                              foreground='gray')
        info_label.pack(pady=(0, 10))
        
        self.process_btn = ttk.Button(step3_frame, text="🚀 Bereinigen & Speichern",
                                      command=self.process_file, state=tk.DISABLED)
        self.process_btn.pack()
        
        # Status
        ttk.Separator(main_frame).pack(fill=tk.X, pady=15)
        
        self.status_label = ttk.Label(main_frame, text="⏳ Warte auf Dateiauswahl...",
                                      font=('Segoe UI', 10))
        self.status_label.pack()
    
    def select_file(self):
        """Datei-Dialog öffnen."""
        file_path = filedialog.askopenfilename(
            title="SAP-Report auswählen",
            initialdir=get_downloads_folder(),
            filetypes=[
                ("Text/Excel Dateien", "*.txt *.xls"),
                ("Textdateien", "*.txt"),
                ("Alle Dateien", "*.*")
            ]
        )
        
        if file_path:
            self.source_file = file_path
            filename = Path(file_path).name
            self.file_label.config(text=filename, foreground='black')
            self.process_btn.config(state=tk.NORMAL)
            self.status_label.config(text=f"✅ Datei geladen: {filename}")
    
    def process_file(self):
        """Datei verarbeiten und speichern."""
        if not self.source_file:
            messagebox.showerror("Fehler", "Bitte zuerst eine Datei auswählen!")
            return
        
        self.status_label.config(text="⏳ Verarbeite...")
        self.root.update()
        
        try:
            # Verarbeiten
            df, df_deleted, stats = process_sap_report(self.source_file)
            
            # Ausgabepfad
            downloads = get_downloads_folder()
            base_name = Path(self.source_file).stem
            
            if self.format_var.get() == "excel":
                output_path = downloads / f"{base_name}_cleaned.xlsx"
                try:
                    import openpyxl
                    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                        df.to_excel(writer, sheet_name='Bereinigte Daten', index=False)
                        if not df_deleted.empty:
                            df_deleted.to_excel(writer, sheet_name='Gelöschte Zeilen', index=False)
                except ImportError:
                    messagebox.showwarning("Hinweis", 
                        "openpyxl nicht installiert. Speichere als CSV statt.")
                    output_path = downloads / f"{base_name}_cleaned.csv"
                    df.to_csv(output_path, index=False, sep=';', encoding='utf-8-sig')
            else:
                output_path = downloads / f"{base_name}_cleaned.csv"
                df.to_csv(output_path, index=False, sep=';', encoding='utf-8-sig')
            
            # Erfolg
            self.status_label.config(
                text=f"✅ Fertig! {stats['kept']} Zeilen bereinigt"
            )
            
            messagebox.showinfo("Erfolgreich!", 
                f"Datei wurde gespeichert:\n\n"
                f"📁 {output_path}\n\n"
                f"📊 Statistik:\n"
                f"   ✓ {stats['kept']} bereinigte Zeilen\n"
                f"   ✗ {stats['sum_rows']} Summenzeilen entfernt\n"
                f"   ✗ {stats['no_material']} ohne Materialnr. entfernt")
            
            # Explorer öffnen
            if sys.platform == 'win32':
                os.startfile(downloads)
            
        except Exception as e:
            self.status_label.config(text="❌ Fehler!")
            messagebox.showerror("Fehler", f"Verarbeitung fehlgeschlagen:\n\n{str(e)}")

# ============================================================
# START
# ============================================================

if __name__ == "__main__":
    root = tk.Tk()
    
    # Windows-Stil
    try:
        root.tk.call('tk', 'scaling', 1.5)
    except:
        pass
    
    # Icon setzen (optional)
    try:
        root.iconbitmap(default='')
    except:
        pass
    
    app = SAPReportCleanerApp(root)
    root.mainloop()

