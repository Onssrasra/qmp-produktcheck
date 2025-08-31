const express = require('express');
const cors = require('cors');
const helmet = require('helmet');
const path = require('path');
const multer = require('multer');
const ExcelJS = require('exceljs');

const {
  toNumber,
  parseWeight,
  weightToKg,
  parseDimensionsToLBH,
  normPartNo,
  mapMaterialClassificationToExcel,
  normalizeNCode
} = require('./utils');
const { SiemensProductScraper, a2vUrl } = require('./scraper');
const { checkCompleteness } = require('./completeness-checker');

const app = express();
const PORT = process.env.PORT || 3000;
const SCRAPE_CONCURRENCY = Number(process.env.SCRAPE_CONCURRENCY || 4);
const WEIGHT_TOL_PCT = Number(process.env.WEIGHT_TOL_PCT || 0); // 0 = strikt

// Ursprüngliche Spalten-Definition (für die Input-Erkennung)
const ORIGINAL_COLS = { Z:'Z', E:'E', C:'C', S:'S', T:'T', U:'U', V:'V', W:'W', P:'P', N:'N' };

// DB/Web-Spaltenpaare – hier definieren wir, nach welchen Originalspalten wir eine Web-Nachbarspalte einfügen
const DB_WEB_PAIRS = [
  { original: 'C', dbCol: null, webCol: null, label: 'Material-Kurztext' },
  { original: 'E', dbCol: null, webCol: null, label: 'Herstellartikelnummer' },
  { original: 'N', dbCol: null, webCol: null, label: 'Fert./Prüfhinweis' },
  { original: 'P', dbCol: null, webCol: null, label: 'Werkstoff' },
  { original: 'S', dbCol: null, webCol: null, label: 'Nettogewicht' },
  { original: 'U', dbCol: null, webCol: null, label: 'Länge' },
  { original: 'V', dbCol: null, webCol: null, label: 'Breite' },
  { original: 'W', dbCol: null, webCol: null, label: 'Höhe' }
];

const HEADER_ROW = 3;      // Spaltennamen
const LABEL_ROW = 4;       // "DB-Wert" / "Web-Wert"
const FIRST_DATA_ROW = 5;  // erste Datenzeile

app.use(helmet({ contentSecurityPolicy: false }));
app.use(cors());
app.use(express.json({ limit: '2mb' }));
app.use(express.static(__dirname));

const scraper = new SiemensProductScraper();

// -------- Helpers: Spalten / Adressen ----------
function getColumnLetter(index) {
  let result = '';
  while (index > 0) {
    index--;
    result = String.fromCharCode(65 + (index % 26)) + result;
    index = Math.floor(index / 26);
  }
  return result;
}
function getColumnIndex(letter) {
  let index = 0;
  for (let i = 0; i < letter.length; i++) {
    index = index * 26 + (letter.charCodeAt(i) - 64);
  }
  return index;
}

// -------- Layout-Berechnung / Struktur ----------
function calculateNewColumnStructure(ws) {
  const newStructure = { pairs: [], otherCols: new Map(), totalInsertedCols: 0 };
  let insertedCols = 0;

  // Für jedes DB/Web-Paar fügen wir rechts daneben 1 Spalte ein
  for (const pair of DB_WEB_PAIRS) {
    const originalIndex = getColumnIndex(pair.original);
    const adjustedOriginalIndex = originalIndex + insertedCols;
    pair.dbCol  = getColumnLetter(adjustedOriginalIndex);
    pair.webCol = getColumnLetter(adjustedOriginalIndex + 1);
    newStructure.pairs.push({ ...pair });
    insertedCols++;
  }
  newStructure.totalInsertedCols = insertedCols;

  // Andere Spalten passend verschieben (Mapping alt → neu)
  const lastCol = ws.lastColumn?.number || ws.columnCount || ws.getRow(HEADER_ROW).cellCount || 0;
  for (let colIndex = 1; colIndex <= lastCol; colIndex++) {
    const originalLetter = getColumnLetter(colIndex);
    const isPairColumn = DB_WEB_PAIRS.some(p => p.original === originalLetter);
    if (!isPairColumn) {
      let insertedBefore = 0;
      for (const p of DB_WEB_PAIRS) {
        if (getColumnIndex(p.original) < colIndex) insertedBefore++;
      }
      const newLetter = getColumnLetter(colIndex + insertedBefore);
      newStructure.otherCols.set(originalLetter, newLetter);
    }
  }
  return newStructure;
}

// -------- Formatierungen ----------
function fillColor(ws, addr, color) {
  if (!color) return;
  const map = {
    green:  'FFD5F4E6', // hellgrün
    red:    'FFFDEAEA', // hellrot
    orange: 'FFFFEAA7', // hellorange
    dbBlue: 'FFE6F3FF', // hellblau (Label DB)
    webBlue:'FFCCE7FF'  // noch helleres Blau (Label Web)
  };
  ws.getCell(addr).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: map[color] || map.green } };
}
function copyColumnFormatting(ws, fromCol, toCol, rowStart, rowEnd) {
  for (let row = rowStart; row <= rowEnd; row++) {
    const fromCell = ws.getCell(`${fromCol}${row}`);
    const toCell   = ws.getCell(`${toCol}${row}`);
    if (fromCell.fill)      toCell.fill = fromCell.fill;
    if (fromCell.font)      toCell.font = fromCell.font;
    if (fromCell.border)    toCell.border = fromCell.border;
    if (fromCell.alignment) toCell.alignment = fromCell.alignment;
    if (fromCell.style)     Object.assign(toCell.style, fromCell.style);
  }
}
function applyLabelCellFormatting(ws, addr, isWebCell = false) {
  const cell = ws.getCell(addr);
  fillColor(ws, addr, isWebCell ? 'webBlue' : 'dbBlue');
  cell.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
  cell.font = { bold: true, size: 10 };
  cell.alignment = { horizontal: 'center', vertical: 'middle' };
}

// -------- Vergleichslogik ----------
function hasValue(v){ return v!==null && v!==undefined && v!=='' && String(v).trim()!==''; }
function eqText(a,b){
  if (a==null||b==null) return false;
  const A=String(a).trim().toLowerCase().replace(/\s+/g,' ');
  const B=String(b).trim().toLowerCase().replace(/\s+/g,' ');
  return A===B;
}
function eqPart(a,b){ return normPartNo(a)===normPartNo(b); }
function eqN(a,b){ return normalizeNCode(a)===normalizeNCode(b); }
function eqWeight(exS, webVal){
  const { value: wv } = parseWeight(webVal);
  if (wv==null) return false;
  const exNum = toNumber(exS); if (exNum==null) return false;
  return Math.abs(exNum - wv) < 1e-9;
}
function eqDimension(exVal, webDimText, dimType){
  const exNum = toNumber(exVal); if (exNum==null) return false;
  const d = parseDimensionsToLBH(webDimText);
  const webVal = (dimType==='L')?d.L:(dimType==='B')?d.B:d.H;
  if (webVal==null) return false;
  return exNum===webVal;
}

// -------- Top-Header (Zeile 1) --------
function applyTopHeader(ws) {
  // Fills (Hintergründe) sichern
  const b1Fill  = ws.getCell('B1').fill;
  const ag1Fill = ws.getCell('AG1').fill;
  const ah1Fill = ws.getCell('AH1').fill;

  // evtl. vorhandene Merges lösen
  try { ws.unMergeCells('B1:AF1'); } catch {}
  try { ws.unMergeCells('AH1:AJ1'); } catch {}

  // B1:AF1
  ws.mergeCells('B1:AF1');
  const b1 = ws.getCell('B1');
  b1.value = 'DB AG SAP R/3 K MARA Stammdaten Stand 20.Mai 2025';
  if (b1Fill) b1.fill = b1Fill;
  b1.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

  // AG1 (einzeln)
  const ag1 = ws.getCell('AG1');
  ag1.value = 'SAP Klassifizierung aus Okt24';
  if (ag1Fill) ag1.fill = ag1Fill;
  ag1.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

  // AH1:AJ1
  ws.mergeCells('AH1:AJ1');
  const ah1 = ws.getCell('AH1');
  ah1.value = 'Zusatz Herstellerdaten aus Abfragen in 2024';
  if (ah1Fill) ah1.fill = ah1Fill;
  ah1.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
}

// -------- NEU: Header-Zeilen 2 & 3 pro DB/Web-Paar zusammenfassen --------
function mergePairHeaders(ws, pairs) {
  for (const pair of pairs) {
    const dbCol  = pair.dbCol;   // Buchstabe, z.B. "C"
    const webCol = pair.webCol;  // Buchstabe, z.B. "D"
    if (!dbCol || !webCol) continue;

    // Vorhandene Merges lösen (falls schon gemergt)
    try { ws.unMergeCells(`${dbCol}2:${webCol}2`); } catch {}
    try { ws.unMergeCells(`${dbCol}3:${webCol}3`); } catch {}

    // Werte aus DB-Header holen (wir verwenden bewusst die DB-Seite als Quelle)
    const v2 = ws.getCell(`${dbCol}2`).value; // technischer Code
    const v3 = ws.getCell(`${dbCol}3`).value; // Klartext Spaltenname

    // Merge durchführen
    ws.mergeCells(`${dbCol}2:${webCol}2`);
    ws.mergeCells(`${dbCol}3:${webCol}3`);

    // Werte und Optik setzen (oben links der Merge-Range)
    const top2 = ws.getCell(`${dbCol}2`);
    const top3 = ws.getCell(`${dbCol}3`);
    top2.value = v2;
    top3.value = v3;

    // Ausrichtung mittig
    top2.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    top3.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

    // Hintergrund/Fonts der DB-Zelle übernehmen (falls vorhanden)
    const src2 = ws.getCell(`${dbCol}2`);
    const src3 = ws.getCell(`${dbCol}3`);
    if (src2.fill)  top2.fill  = src2.fill;
    if (src2.font)  top2.font  = src2.font;
    if (src2.border)top2.border= src2.border;

    if (src3.fill)  top3.fill  = src3.fill;
    if (src3.font)  top3.font  = src3.font;
    if (src3.border)top3.border= src3.border;
  }
}

// -------- Routes ----------
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'index.html')));
app.get('/api/health', (req, res) => res.json({ ok: true, time: new Date().toISOString() }));

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

app.post('/api/process-excel', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Bitte Excel-Datei hochladen (file).' });

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(req.file.buffer);

    // 1) A2V-Nummern aus Spalte Z (ursprünglich) einsammeln, bevor wir umbauen
    const tasks = [];
    const rowsPerSheet = new Map(); // ws -> [rowIndex,...]
    for (const ws of wb.worksheets) {
      const indices = [];
      const last = ws.lastRow?.number || 0;
      for (let r = FIRST_DATA_ROW - 1; r <= last; r++) { // -1, weil wir gleich eine Zeile 4 einfügen
        const a2v = (ws.getCell(`${ORIGINAL_COLS.Z}${r}`).value || '').toString().trim().toUpperCase();
        if (a2v.startsWith('A2V')) { indices.push(r); tasks.push(a2v); }
      }
      rowsPerSheet.set(ws, indices);
    }

    // 2) Scrapen
    const resultsMap = await scraper.scrapeMany(tasks, SCRAPE_CONCURRENCY);

    // 3) Umbau pro Worksheet
    for (const ws of wb.worksheets) {
      // 3.1 Spaltenstruktur berechnen
      const structure = calculateNewColumnStructure(ws);

      // 3.2 Spalten einfügen (von rechts nach links)
      for (const pair of [...structure.pairs].reverse()) {
        const insertPos = getColumnIndex(pair.original) + 1; // rechts neben der Originalspalte
        ws.spliceColumns(insertPos, 0, [null]);
      }

      // 3.3 Zeile 4 (Labels) einfügen
      ws.spliceRows(LABEL_ROW, 0, [null]);

      // 3.4 Zeilen 2 & 3 Inhalte in Web-Spalten spiegeln + Labels schreiben
      for (const pair of structure.pairs) {
        // Inhalte 2/3 spiegeln
        const dbTech = ws.getCell(`${pair.dbCol}2`).value;
        const dbName = ws.getCell(`${pair.dbCol}3`).value;
        ws.getCell(`${pair.webCol}2`).value = dbTech;
        ws.getCell(`${pair.webCol}3`).value = dbName;
        copyColumnFormatting(ws, pair.dbCol, pair.webCol, 1, 3);

        // Labels Zeile 4
        ws.getCell(`${pair.dbCol}${LABEL_ROW}`).value  = 'DB-Wert';
        ws.getCell(`${pair.webCol}${LABEL_ROW}`).value = 'Web-Wert';
        applyLabelCellFormatting(ws, `${pair.dbCol}${LABEL_ROW}`, false);
        applyLabelCellFormatting(ws, `${pair.webCol}${LABEL_ROW}`, true);
      }

      // 3.5 Top-Header (Zeile 1) setzen
      applyTopHeader(ws);

      // 3.6 NEU: Header in Zeile 2 und 3 pro Paar zusammenfassen (C2:D2, C3:D3, F2:G2, F3:G3, ...)
      mergePairHeaders(ws, structure.pairs);

      // 3.7 Web-Daten eintragen / vergleichen
      const prodRows = rowsPerSheet.get(ws) || [];
      for (const originalRow of prodRows) {
        const currentRow = originalRow + 1; // wegen eingefügter Label-Zeile

        // neue Z-Spalte (A2V) bestimmen
        let zCol = ORIGINAL_COLS.Z;
        if (structure.otherCols.has(ORIGINAL_COLS.Z)) zCol = structure.otherCols.get(ORIGINAL_COLS.Z);
        const a2v = (ws.getCell(`${zCol}${currentRow}`).value || '').toString().trim().toUpperCase();
        const web = resultsMap.get(a2v) || {};

        // je Paar
        for (const pair of structure.pairs) {
          const dbValue = ws.getCell(`${pair.dbCol}${currentRow}`).value;
          let webValue = null;
          let isEqual = false;

          switch (pair.original) {
            case 'C': // Material-Kurztext
              webValue = (web.Produkttitel && web.Produkttitel !== 'Nicht gefunden') ? web.Produkttitel : null;
              isEqual  = webValue ? eqText(dbValue || '', webValue) : false;
              break;
            case 'E': // Herstellartikelnummer
              webValue = (web['Weitere Artikelnummer'] && web['Weitere Artikelnummer'] !== 'Nicht gefunden')
                        ? web['Weitere Artikelnummer']
                        : a2v;
              isEqual  = eqPart(dbValue || a2v, webValue);
              break;
            case 'N': // Fert./Prüfhinweis
              if (web.Materialklassifizierung && web.Materialklassifizierung !== 'Nicht gefunden') {
                const code = normalizeNCode(mapMaterialClassificationToExcel(web.Materialklassifizierung));
                if (code) { webValue = code; isEqual = eqN(dbValue || '', code); }
              }
              break;
            case 'P': // Werkstoff
              webValue = (web.Werkstoff && web.Werkstoff !== 'Nicht gefunden') ? web.Werkstoff : null;
              isEqual  = webValue ? eqText(dbValue || '', webValue) : false;
              break;
            case 'S': // Nettogewicht
              if (web.Gewicht && web.Gewicht !== 'Nicht gefunden') {
                const { value } = parseWeight(web.Gewicht);
                if (value != null) { webValue = value; isEqual = eqWeight(dbValue, web.Gewicht); }
              }
              break;
            case 'U': // Länge
              if (web.Abmessung && web.Abmessung !== 'Nicht gefunden') {
                const d = parseDimensionsToLBH(web.Abmessung);
                if (d.L != null) { webValue = d.L; isEqual = eqDimension(dbValue, web.Abmessung, 'L'); }
              }
              break;
            case 'V': // Breite
              if (web.Abmessung && web.Abmessung !== 'Nicht gefunden') {
                const d = parseDimensionsToLBH(web.Abmessung);
                if (d.B != null) { webValue = d.B; isEqual = eqDimension(dbValue, web.Abmessung, 'B'); }
              }
              break;
            case 'W': // Höhe
              if (web.Abmessung && web.Abmessung !== 'Nicht gefunden') {
                const d = parseDimensionsToLBH(web.Abmessung);
                if (d.H != null) { webValue = d.H; isEqual = eqDimension(dbValue, web.Abmessung, 'H'); }
              }
              break;
          }

          const hasDb = hasValue(dbValue);
          const hasWeb = webValue !== null;

          if (hasWeb) {
            ws.getCell(`${pair.webCol}${currentRow}`).value = webValue;
            // Nur markieren wenn DB-Wert vorhanden ist
            if (hasDb) {
              fillColor(ws, `${pair.webCol}${currentRow}`, isEqual ? 'green' : 'red');
            }
            // Wenn DB-Wert fehlt, aber Web-Wert vorhanden → keine Markierung
          } else {
            // Web-Wert fehlt, aber DB-Wert vorhanden → orange
            if (hasDb) fillColor(ws, `${pair.webCol}${currentRow}`, 'orange');
          }
        }
      }
    }

    const out = await wb.xlsx.writeBuffer();
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition','attachment; filename="Web_Vergleich_Ergebnis.xlsx"');
    res.send(Buffer.from(out));

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// Neue Route für Vollständigkeitsprüfung
app.post('/api/check-completeness', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Bitte Excel-Datei hochladen (file).' });

    const result = await checkCompleteness(req.file.buffer);
    
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition','attachment; filename="Qualitätsbericht.xlsx"');
    res.send(Buffer.from(result));

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// Neue Route für Qualitätsbericht Statistiken
app.post('/api/quality-stats', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Bitte Excel-Datei hochladen (file).' });

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(req.file.buffer);
    const ws = wb.worksheets[0];
    
    let greenCount = 0;  // Vollständige und richtige Datensätze
    let redCount = 0;    // Unvollständige oder unplausible Datensätze
    
    // Debug: Log all found colors
    const foundColors = new Set();
    
    // Count colored cells in data rows (from row 4)
    const lastRow = ws.lastRow ? ws.lastRow.number : 0;
    for (let r = 4; r <= lastRow; r++) {
      const row = ws.getRow(r);
      let rowHasRed = false;
      let rowHasGreen = false;
      
      for (let c = 1; c <= row.cellCount; c++) {
        const cell = row.getCell(c);
        if (cell.fill && cell.fill.fgColor) {
          const color = cell.fill.fgColor.argb;
          foundColors.add(color);
          
          if (color === 'FFCCFFCC') { // Green - vollständig und richtig
            rowHasGreen = true;
          } else if (color === 'FFFFCCCC') { // Red - unvollständig oder unplausibel
            rowHasRed = true;
          }
        }
      }
      
      // Count complete rows (all green) vs incomplete rows (any red)
      if (rowHasGreen && !rowHasRed) {
        greenCount++;
      } else if (rowHasRed) {
        redCount++;
      }
    }
    
    console.log('Found colors in Qualitätsbericht:', Array.from(foundColors));
    console.log('Quality counts:', { complete: greenCount, incomplete: redCount });
    
    const totalRows = lastRow - 3; // Abzüglich Header-Zeilen
    
    res.json({
      total: totalRows,
      complete: greenCount,
      incomplete: redCount,
      completePercentage: Math.round((greenCount / totalRows) * 100),
      incompletePercentage: Math.round((redCount / totalRows) * 100),
      debug: {
        foundColors: Array.from(foundColors),
        totalRows: lastRow
      }
    });

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// Neue Route für Web-Suche Statistiken
app.post('/api/web-search-stats', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Bitte Excel-Datei hochladen (file).' });

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(req.file.buffer);
    const ws = wb.worksheets[0];
    
    let greenCount = 0;  // Übereinstimmungen (grün)
    let redCount = 0;    // Abweichungen (rot)
    let orangeCount = 0; // Fehlende Web-Werte (orange)
    
    // Debug: Log all found colors
    const foundColors = new Set();
    
    // Count colored cells in web value columns
    const lastRow = ws.lastRow ? ws.lastRow.number : 0;
    for (let r = 5; r <= lastRow; r++) { // Start from row 5 (after labels)
      for (let c = 1; c <= ws.columnCount; c++) {
        const cell = ws.getCell(r, c);
        if (cell.fill && cell.fill.fgColor) {
          const color = cell.fill.fgColor.argb;
          foundColors.add(color);
          
          // Check only the specific colors used in the system
          if (color === 'FFD5F4E6') { // Green - Übereinstimmungen
            greenCount++;
          } else if (color === 'FFFDEAEA') { // Red - Abweichungen
            redCount++;
          } else if (color === 'FFFFEAA7') { // Orange - Fehlende Web-Werte
            orangeCount++;
          }
        }
      }
    }
    
    console.log('Found colors in Excel:', Array.from(foundColors));
    console.log('Counts:', { green: greenCount, red: redCount, orange: orangeCount });
    
    let siemensRows = 0;        // Anzahl Zeilen mit A2V-Nummern
    let searchedValues = 0;     // Gesuchte Werte (alle Zeilen die im Web gesucht wurden)
    let totalWebValues = 0;     // Gefundene Web-Werte (alle Zellen)
    
    // Verwende Spalte Z für Siemens Mobility Materialnummer
    const siemensColumn = 26; // Spalte Z
    console.log(`Verwende Spalte Z (${siemensColumn}) für Siemens Mobility Materialnummer`);
    
    // Prüfe ob Spalte Z den richtigen Header hat
    const headerCell = ws.getCell(3, siemensColumn);
    const headerValue = headerCell.value;
    console.log(`Spalte Z Header (Zeile 3): "${headerValue}"`);
    
    if (!headerValue || !headerValue.toString().includes('Siemens Mobility Materialnummer')) {
      console.log('Warnung: Spalte Z hat nicht den erwarteten Header "Siemens Mobility Materialnummer"');
    }
    
    // Zähle Siemens-Zeilen und Web-Werte
    for (let r = 5; r <= lastRow; r++) { // Start from row 5 (after labels)
      const siemensCell = ws.getCell(r, siemensColumn);
      const siemensValue = (siemensCell.value || '').toString().trim().toUpperCase();
      
      if (siemensValue.startsWith('A2V')) {
        siemensRows++;
        console.log(`Siemens-Zeile gefunden: ${r}, A2V: ${siemensValue}`);
        
        // Zähle Web-Werte in dieser Zeile (nur Zellen mit Farbmarkierungen)
        for (let c = 1; c <= ws.columnCount; c++) {
          const cell = ws.getCell(r, c);
          if (cell.fill && cell.fill.fgColor) {
            totalWebValues++;
          }
        }
      }
    }
    
    // Gesuchte Werte = Anzahl Siemens-Zeilen × 8 (8 Datenfelder pro Produkt)
    searchedValues = siemensRows * 8;
    
    console.log('Berechnung:', {
      siemensRows,
      searchedValues,
      totalWebValues,
      a2vColumn
    });
    
    // Berechne zusätzliche Prozentangaben
    const foundWebValuesPercentage = searchedValues > 0 ? Math.round((totalWebValues / searchedValues) * 100) : 0;
    
    res.json({
      totalSiemens: siemensRows,
      searchedValues: searchedValues,
      foundWebValues: totalWebValues,
      foundWebValuesPercentage: foundWebValuesPercentage,
      green: greenCount,
      red: redCount,
      orange: orangeCount,
      greenPercentage: totalWebValues > 0 ? Math.round((greenCount / totalWebValues) * 100) : 0,
      redPercentage: totalWebValues > 0 ? Math.round((redCount / totalWebValues) * 100) : 0,
      orangePercentage: totalWebValues > 0 ? Math.round((orangeCount / totalWebValues) * 100) : 0,
      debug: {
        foundColors: Array.from(foundColors),
        totalRows: lastRow,
        totalColumns: ws.columnCount,
        siemensRows: siemensRows,
        searchedValues: searchedValues,
        totalWebValues: totalWebValues,
        siemensColumn: siemensColumn
      }
    });

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// Neue Route für optimierte Web-Suche (nur Siemens-Produkte)
app.post('/api/process-excel-siemens', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Bitte Excel-Datei hochladen (file).' });

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(req.file.buffer);

    // 1) Nur Siemens-Produkte (A2V-Nummern) extrahieren
    const tasks = [];
    const siemensRowsPerSheet = new Map(); // ws -> [rowIndex,...]
    
    for (const ws of wb.worksheets) {
      const siemensIndices = [];
      const last = ws.lastRow?.number || 0;
      
      for (let r = FIRST_DATA_ROW - 1; r <= last; r++) {
        const a2v = (ws.getCell(`${ORIGINAL_COLS.Z}${r}`).value || '').toString().trim().toUpperCase();
        if (a2v.startsWith('A2V')) { 
          siemensIndices.push(r); 
          tasks.push(a2v); 
        }
      }
      siemensRowsPerSheet.set(ws, siemensIndices);
    }
    
    // Nur fortfahren wenn Siemens-Produkte gefunden wurden
    if (tasks.length === 0) {
      return res.status(400).json({ error: 'Keine Siemens-Produkte (A2V-Nummern) in der Datei gefunden.' });
    }

    // 2) Scrapen
    const resultsMap = await scraper.scrapeMany(tasks, SCRAPE_CONCURRENCY);

    // 3) Erstelle neue Siemens-Datei und verarbeite nur Siemens-Produkte
    const siemensWb = new ExcelJS.Workbook();
    
    for (const ws of wb.worksheets) {
      const siemensWs = siemensWb.addWorksheet(ws.name);
      
      // Kopiere Header-Zeilen (1-3)
      for (let r = 1; r <= 3; r++) {
        const row = ws.getRow(r);
        for (let c = 1; c <= row.cellCount; c++) {
          const cell = row.getCell(c);
          siemensWs.getCell(r, c).value = cell.value;
          if (cell.fill) siemensWs.getCell(r, c).fill = cell.fill;
          if (cell.font) siemensWs.getCell(r, c).font = cell.font;
          if (cell.border) siemensWs.getCell(r, c).border = cell.border;
          if (cell.alignment) siemensWs.getCell(r, c).alignment = cell.alignment;
        }
      }
      
      // Kopiere nur Siemens-Zeilen (ab Zeile 4)
      const siemensRows = siemensRowsPerSheet.get(ws) || [];
      let newRowIndex = 4;
      
      for (const originalRow of siemensRows) {
        const row = ws.getRow(originalRow);
        for (let c = 1; c <= row.cellCount; c++) {
          const cell = row.getCell(c);
          siemensWs.getCell(newRowIndex, c).value = cell.value;
          if (cell.fill) siemensWs.getCell(newRowIndex, c).fill = cell.fill;
          if (cell.font) siemensWs.getCell(newRowIndex, c).font = cell.font;
          if (cell.border) siemensWs.getCell(newRowIndex, c).border = cell.border;
          if (cell.alignment) siemensWs.getCell(newRowIndex, c).alignment = cell.alignment;
        }
        newRowIndex++;
      }
    }
    
    // 4) Verarbeite die Siemens-Datei
    for (const ws of siemensWb.worksheets) {
      // 4.1 Spaltenstruktur berechnen
      const structure = calculateNewColumnStructure(ws);

      // 4.2 Spalten einfügen (von rechts nach links)
      for (const pair of [...structure.pairs].reverse()) {
        const insertPos = getColumnIndex(pair.original) + 1; // rechts neben der Originalspalte
        ws.spliceColumns(insertPos, 0, [null]);
      }

      // 4.3 Zeile 4 (Labels) einfügen
      ws.spliceRows(LABEL_ROW, 0, [null]);

      // 4.4 Zeilen 2 & 3 Inhalte in Web-Spalten spiegeln + Labels schreiben
      for (const pair of structure.pairs) {
        // Inhalte 2/3 spiegeln
        const dbTech = ws.getCell(`${pair.dbCol}2`).value;
        const dbName = ws.getCell(`${pair.dbCol}3`).value;
        ws.getCell(`${pair.webCol}2`).value = dbTech;
        ws.getCell(`${pair.webCol}3`).value = dbName;
        copyColumnFormatting(ws, pair.dbCol, pair.webCol, 1, 3);

        // Labels Zeile 4
        ws.getCell(`${pair.dbCol}${LABEL_ROW}`).value  = 'DB-Wert';
        ws.getCell(`${pair.webCol}${LABEL_ROW}`).value = 'Web-Wert';
        applyLabelCellFormatting(ws, `${pair.dbCol}${LABEL_ROW}`, false);
        applyLabelCellFormatting(ws, `${pair.webCol}${LABEL_ROW}`, true);
      }

      // 4.5 Top-Header (Zeile 1) setzen
      applyTopHeader(ws);

      // 4.6 Header in Zeile 2 und 3 pro Paar zusammenfassen
      mergePairHeaders(ws, structure.pairs);

      // 4.7 Web-Daten eintragen / vergleichen
      const siemensRows = siemensRowsPerSheet.get(wb.worksheets.find(w => w.name === ws.name)) || [];
      for (let i = 0; i < siemensRows.length; i++) {
        const currentRow = 5 + i; // Start ab Zeile 5 (nach Labels)

        // neue Z-Spalte (A2V) bestimmen
        let zCol = ORIGINAL_COLS.Z;
        if (structure.otherCols.has(ORIGINAL_COLS.Z)) zCol = structure.otherCols.get(ORIGINAL_COLS.Z);
        const a2v = (ws.getCell(`${zCol}${currentRow}`).value || '').toString().trim().toUpperCase();
        const web = resultsMap.get(a2v) || {};

        // je Paar
        for (const pair of structure.pairs) {
          const dbValue = ws.getCell(`${pair.dbCol}${currentRow}`).value;
          let webValue = null;
          let isEqual = false;

          switch (pair.original) {
            case 'C': // Material-Kurztext
              webValue = (web.Produkttitel && web.Produkttitel !== 'Nicht gefunden') ? web.Produkttitel : null;
              isEqual  = webValue ? eqText(dbValue || '', webValue) : false;
              break;
            case 'E': // Herstellartikelnummer
              webValue = (web['Weitere Artikelnummer'] && web['Weitere Artikelnummer'] !== 'Nicht gefunden')
                        ? web['Weitere Artikelnummer']
                        : a2v;
              isEqual  = eqPart(dbValue || a2v, webValue);
              break;
            case 'N': // Fert./Prüfhinweis
              if (web.Materialklassifizierung && web.Materialklassifizierung !== 'Nicht gefunden') {
                const code = normalizeNCode(mapMaterialClassificationToExcel(web.Materialklassifizierung));
                if (code) { webValue = code; isEqual = eqN(dbValue || '', code); }
              }
              break;
            case 'P': // Werkstoff
              webValue = (web.Werkstoff && web.Werkstoff !== 'Nicht gefunden') ? web.Werkstoff : null;
              isEqual  = webValue ? eqText(dbValue || '', webValue) : false;
              break;
            case 'S': // Nettogewicht
              if (web.Gewicht && web.Gewicht !== 'Nicht gefunden') {
                const { value } = parseWeight(web.Gewicht);
                if (value != null) { webValue = value; isEqual = eqWeight(dbValue, web.Gewicht); }
              }
              break;
            case 'U': // Länge
              if (web.Abmessung && web.Abmessung !== 'Nicht gefunden') {
                const d = parseDimensionsToLBH(web.Abmessung);
                if (d.L != null) { webValue = d.L; isEqual = eqDimension(dbValue, web.Abmessung, 'L'); }
              }
              break;
            case 'V': // Breite
              if (web.Abmessung && web.Abmessung !== 'Nicht gefunden') {
                const d = parseDimensionsToLBH(web.Abmessung);
                if (d.B != null) { webValue = d.B; isEqual = eqDimension(dbValue, web.Abmessung, 'B'); }
              }
              break;
            case 'W': // Höhe
              if (web.Abmessung && web.Abmessung !== 'Nicht gefunden') {
                const d = parseDimensionsToLBH(web.Abmessung);
                if (d.H != null) { webValue = d.H; isEqual = eqDimension(dbValue, web.Abmessung, 'H'); }
              }
              break;
          }

          const hasDb = hasValue(dbValue);
          const hasWeb = webValue !== null;

          if (hasWeb) {
            ws.getCell(`${pair.webCol}${currentRow}`).value = webValue;
            // Nur markieren wenn DB-Wert vorhanden ist
            if (hasDb) {
              fillColor(ws, `${pair.webCol}${currentRow}`, isEqual ? 'green' : 'red');
            }
            // Wenn DB-Wert fehlt, aber Web-Wert vorhanden → keine Markierung
          } else {
            // Web-Wert fehlt, aber DB-Wert vorhanden → orange
            if (hasDb) fillColor(ws, `${pair.webCol}${currentRow}`, 'orange');
          }
        }
      }
    }

    const out = await siemensWb.xlsx.writeBuffer();
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition','attachment; filename="Web_Vergleich_Siemens.xlsx"');
    res.send(Buffer.from(out));

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => console.log(`Server running at http://0.0.0.0:${PORT}`));
