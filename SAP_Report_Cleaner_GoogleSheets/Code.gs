/**
 * SAP Report Cleaner - Google Sheets Version
 * ============================================
 * Bereinigt SAP-Reports automatisch aus einem Google Drive Ordner.
 * 
 * Ablauf:
 * 1. SAP-Datei in den Ordner "SAP_Reports" in Google Drive hochladen
 * 2. In diesem Sheet auf "SAP Cleaner" → "Neueste Datei bereinigen" klicken
 * 3. Die bereinigten Daten erscheinen im Tab "Bereinigte Daten"
 */

// ============================================================
// KONFIGURATION
// ============================================================

const CONFIG = {
  // Name des Ordners in Google Drive (wird automatisch erstellt)
  FOLDER_NAME: "SAP_Reports",
  
  // Erwartete Spaltenüberschriften
  HEADERS: [
    'Material', 'Functional Loc.', 'Equipment', 'Material Description',
    'Work Ctr', 'Withdrawn', 'W/o resrv.', 'Reserved', 'Reserv.ref',
    'Pstng Date', 'Order', 'ID', 'Message', 'ICt', 'Customer'
  ],
  
  // Numerische Spalten (werden als Zahlen formatiert)
  NUMERIC_COLUMNS: ['Material', 'Withdrawn', 'W/o resrv.', 'Reserved', 
                    'Reserv.ref', 'Order', 'Message']
};

// ============================================================
// MENÜ ERSTELLEN
// ============================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🧹 SAP Cleaner')
    .addItem('📂 Neueste Datei bereinigen', 'processNewestFile')
    .addSeparator()
    .addItem('📋 Dateien im Ordner anzeigen', 'showFilesInFolder')
    .addItem('📁 Ordner in Drive öffnen', 'openDriveFolder')
    .addSeparator()
    .addItem('❓ Hilfe', 'showHelp')
    .addToUi();
}

// ============================================================
// HAUPTFUNKTION: Neueste Datei verarbeiten
// ============================================================

function processNewestFile() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // Ordner finden oder erstellen
    const folder = getOrCreateFolder();
    
    // Neueste .txt oder .xls Datei finden
    const file = getNewestFile(folder);
    
    if (!file) {
      ui.alert('Keine Datei gefunden', 
        'Es wurde keine .txt oder .xls Datei im Ordner "' + CONFIG.FOLDER_NAME + '" gefunden.\n\n' +
        'Bitte laden Sie zuerst eine SAP-Report-Datei in diesen Ordner hoch.',
        ui.ButtonSet.OK);
      return;
    }
    
    // Bestätigung
    const response = ui.alert('Datei gefunden',
      'Neueste Datei: ' + file.getName() + '\n' +
      'Erstellt: ' + file.getDateCreated().toLocaleString() + '\n\n' +
      'Möchten Sie diese Datei bereinigen?',
      ui.ButtonSet.YES_NO);
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    // Status anzeigen
    ss.toast('Verarbeite ' + file.getName() + '...', 'SAP Cleaner', 30);
    
    // Datei verarbeiten
    const result = processSAPFile(file);
    
    // Ergebnisse in Sheets schreiben
    writeToSheet(ss, 'Bereinigte Daten', result.cleanedData, CONFIG.HEADERS);
    writeToSheet(ss, 'Gelöschte Zeilen', result.deletedData, ['Grund', 'Zeile', 'Daten']);
    
    // Erfolg
    ss.toast('Fertig!', 'SAP Cleaner', 5);
    
    ui.alert('✅ Erfolgreich!',
      'Datei: ' + file.getName() + '\n\n' +
      '📊 Statistik:\n' +
      '   ✓ ' + result.stats.kept + ' bereinigte Zeilen\n' +
      '   ✗ ' + result.stats.sumRows + ' Summenzeilen entfernt\n' +
      '   ✗ ' + result.stats.noMaterial + ' ohne Materialnr. entfernt\n\n' +
      'Die Daten sind jetzt im Tab "Bereinigte Daten".',
      ui.ButtonSet.OK);
    
    // Zum Tab wechseln
    ss.getSheetByName('Bereinigte Daten').activate();
    
  } catch (error) {
    ui.alert('❌ Fehler', 'Verarbeitung fehlgeschlagen:\n\n' + error.message, ui.ButtonSet.OK);
    console.error(error);
  }
}

// ============================================================
// SAP-DATEI VERARBEITEN
// ============================================================

function processSAPFile(file) {
  // Datei lesen
  const content = file.getBlob().getDataAsString('UTF-8');
  const lines = content.split('\n');
  const allRows = lines.map(line => line.split('\t'));
  
  // Header-Zeile finden
  let headerRowIdx = null;
  let headerStartCol = null;
  
  for (let idx = 0; idx < allRows.length; idx++) {
    const row = allRows[idx];
    for (let colIdx = 0; colIdx < row.length; colIdx++) {
      if (row[colIdx].trim().toLowerCase() === 'material') {
        headerRowIdx = idx;
        headerStartCol = colIdx;
        break;
      }
    }
    if (headerRowIdx !== null) break;
  }
  
  if (headerRowIdx === null) {
    headerRowIdx = 3;
    headerStartCol = 2;
  }
  
  // Daten verarbeiten
  const cleanedData = [];
  const deletedData = [];
  const stats = { total: 0, sumRows: 0, empty: 0, noMaterial: 0, kept: 0 };
  
  for (let rowIdx = headerRowIdx + 1; rowIdx < allRows.length; rowIdx++) {
    const row = allRows[rowIdx];
    stats.total++;
    
    // Leere Zeile?
    if (row.every(cell => cell.trim() === '')) {
      stats.empty++;
      continue;
    }
    
    // Summenzeile?
    const colB = row[1] ? row[1].trim() : '';
    if (colB === '*' || colB === '**') {
      stats.sumRows++;
      deletedData.push(['Summenzeile', rowIdx + 1, row.join('\t')]);
      continue;
    }
    
    // Daten extrahieren (Spalten C bis Q)
    const dataRow = [];
    for (let i = headerStartCol; i < headerStartCol + 15; i++) {
      dataRow.push(row[i] ? row[i].trim() : '');
    }
    
    // Keine Materialnummer?
    if (!dataRow[0]) {
      stats.noMaterial++;
      deletedData.push(['Keine Materialnummer', rowIdx + 1, dataRow.join('\t')]);
      continue;
    }
    
    // Zahlen bereinigen
    for (let i = 0; i < dataRow.length; i++) {
      const header = CONFIG.HEADERS[i];
      if (CONFIG.NUMERIC_COLUMNS.includes(header)) {
        dataRow[i] = cleanNumber(dataRow[i]);
      }
    }
    
    // Datum konvertieren
    const dateIdx = CONFIG.HEADERS.indexOf('Pstng Date');
    if (dateIdx >= 0 && dataRow[dateIdx]) {
      dataRow[dateIdx] = convertDate(dataRow[dateIdx]);
    }
    
    cleanedData.push(dataRow);
    stats.kept++;
  }
  
  return { cleanedData, deletedData, stats };
}

// ============================================================
// HILFSFUNKTIONEN
// ============================================================

function cleanNumber(value) {
  if (!value || value === '-') return '';
  
  let str = value.toString().trim()
    .replace(/\s/g, '')
    .replace(/\u00A0/g, '');
  
  // Deutsches Format: 1.234,56 → 1234.56
  if (str.includes(',') && str.includes('.')) {
    str = str.replace(/\./g, '').replace(',', '.');
  } else if (str.includes(',')) {
    const parts = str.split(',');
    if (parts.length === 2 && parts[1].length <= 2) {
      str = str.replace(',', '.');
    } else {
      str = str.replace(/,/g, '');
    }
  } else if (str.includes('.')) {
    const parts = str.split('.');
    if (parts.length === 2 && parts[1].length === 3) {
      str = str.replace('.', '');
    }
  }
  
  const num = parseFloat(str);
  return isNaN(num) ? '' : Math.round(num);
}

function convertDate(value) {
  if (!value) return '';
  
  const str = value.toString().trim();
  const patterns = [
    /^(\d{2})\.(\d{2})\.(\d{2})$/,   // DD.MM.YY
    /^(\d{2})\.(\d{2})\.(\d{4})$/    // DD.MM.YYYY
  ];
  
  for (const pattern of patterns) {
    const match = str.match(pattern);
    if (match) {
      let year = parseInt(match[3]);
      if (year < 100) {
        year += year < 50 ? 2000 : 1900;
      }
      return match[1] + '.' + match[2] + '.' + year;
    }
  }
  
  return str;
}

function getOrCreateFolder() {
  const folders = DriveApp.getFoldersByName(CONFIG.FOLDER_NAME);
  
  if (folders.hasNext()) {
    return folders.next();
  }
  
  // Ordner erstellen
  const folder = DriveApp.createFolder(CONFIG.FOLDER_NAME);
  SpreadsheetApp.getUi().alert('Ordner erstellt',
    'Der Ordner "' + CONFIG.FOLDER_NAME + '" wurde in Ihrem Google Drive erstellt.\n\n' +
    'Bitte laden Sie Ihre SAP-Reports dort hoch.',
    SpreadsheetApp.getUi().ButtonSet.OK);
  
  return folder;
}

function getNewestFile(folder) {
  const files = folder.getFiles();
  let newestFile = null;
  let newestDate = new Date(0);
  
  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName().toLowerCase();
    
    // Nur .txt und .xls Dateien
    if (name.endsWith('.txt') || name.endsWith('.xls')) {
      const created = file.getDateCreated();
      if (created > newestDate) {
        newestDate = created;
        newestFile = file;
      }
    }
  }
  
  return newestFile;
}

function writeToSheet(ss, sheetName, data, headers) {
  // Sheet finden oder erstellen
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  // Alte Daten löschen
  sheet.clear();
  
  if (data.length === 0) {
    sheet.getRange(1, 1).setValue('Keine Daten');
    return;
  }
  
  // Header schreiben
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('white');
  
  // Daten schreiben
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  }
  
  // Spaltenbreite anpassen
  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }
  
  // Filter hinzufügen
  sheet.getRange(1, 1, data.length + 1, headers.length).createFilter();
}

// ============================================================
// ZUSÄTZLICHE MENÜFUNKTIONEN
// ============================================================

function showFilesInFolder() {
  const ui = SpreadsheetApp.getUi();
  const folder = getOrCreateFolder();
  const files = folder.getFiles();
  
  let fileList = 'Dateien im Ordner "' + CONFIG.FOLDER_NAME + '":\n\n';
  let count = 0;
  
  while (files.hasNext()) {
    const file = files.next();
    fileList += '📄 ' + file.getName() + '\n';
    fileList += '   Erstellt: ' + file.getDateCreated().toLocaleString() + '\n\n';
    count++;
  }
  
  if (count === 0) {
    fileList += '(keine Dateien)';
  }
  
  ui.alert('📋 Dateien', fileList, ui.ButtonSet.OK);
}

function openDriveFolder() {
  const folder = getOrCreateFolder();
  const url = folder.getUrl();
  
  const html = '<script>window.open("' + url + '", "_blank");google.script.host.close();</script>';
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(1).setHeight(1),
    'Öffne Drive...'
  );
}

function showHelp() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('❓ Hilfe - SAP Report Cleaner',
    'So funktioniert es:\n\n' +
    '1️⃣ SAP-Report aus P30 herunterladen (.txt Datei)\n\n' +
    '2️⃣ Datei in Google Drive hochladen:\n' +
    '   • Öffnen Sie Google Drive\n' +
    '   • Gehen Sie in den Ordner "SAP_Reports"\n' +
    '   • Ziehen Sie die Datei hinein\n\n' +
    '3️⃣ In diesem Sheet:\n' +
    '   • Menü "SAP Cleaner" → "Neueste Datei bereinigen"\n' +
    '   • Die bereinigten Daten erscheinen im Tab\n\n' +
    '📊 Was wird bereinigt:\n' +
    '   • Summenzeilen (mit * markiert) entfernt\n' +
    '   • Zeilen ohne Materialnummer entfernt\n' +
    '   • Zahlenformate korrigiert\n' +
    '   • Nur relevante Spalten (C-Q)',
    ui.ButtonSet.OK);
}

