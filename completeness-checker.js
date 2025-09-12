/* completeness-checker.js */
const ExcelJS = require('exceljs');

// Constants for completeness check
const HEADER_ROW = 3;       // Header in row 3
const FIRST_DATA_ROW = 4;   // Data start in row 4

// Colors
const FILL_RED    = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFCCCC' } }; // Pflicht fehlt
const FILL_ORANGE = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE0B2' } }; // unplausibel/ungültig
const FILL_GREEN  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCCFFCC' } }; // Zeile OK
const FILL_AMPEL_GREEN = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00F26D' } }; // Ampel grün
const FILL_AMPEL_RED   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } }; // Ampel rot

// Allowed values for Fert./Prüfhinweis segments
const POS1 = new Set(['OHNE','1','2','3']);
const POS2 = new Set(['N','3.2','3.1','2.2','2.1']);
const POS3 = new Set(['N','CL1','CL2','CL3']);
const POS4 = new Set(['N','J']);
const POS5 = new Set(['N','A1','A2','A3','A5','A+']);

// Pflichtfelder: B–J, N, R–W (1-based Excel columns)
const MUST_COL_RANGES = [
  { start: 2, end: 10 },  // B..J
  { start: 14, end: 14 }, // N
  { start: 18, end: 23 }, // R..W
];

function isEmpty(v){ return v == null || String(v).trim() === ''; }

function toNum(v){
  if (v == null || String(v).trim() === '') return null;
  const n = Number(String(v).replace(',', '.'));
  return Number.isFinite(n) ? n : null;
}

const TEXT_MASS_RE = /\d{1,4}[\s×xX*/]{1,3}\d{1,4}/;

function hasTextMeasure(s){ return s != null && TEXT_MASS_RE.test(String(s)); }

function validFertPruef(v){
  if (v == null) return false;
  const parts = String(v).split('/').map(t => String(t).trim());
  if (parts.length !== 5) return false;
  return POS1.has(parts[0]) && POS2.has(parts[1]) && POS3.has(parts[2]) && POS4.has(parts[3]) && POS5.has(parts[4]);
}

/** Add Ampelbewertung column for quality report */
function addQualityAmpelColumn(ws) {
  // Header für Ampelbewertung-Spalte hinzufügen
  ws.getCell('A3').value = 'Ampelbewertung';
  
  // Formatierung für Header
  ws.getCell('A3').font = { bold: true };
  ws.getCell('A3').alignment = { horizontal: 'center', vertical: 'middle' };
}

/** Calculate Ampel status for quality report row - Optimized */
function calculateQualityAmpelStatus(hasRed) {
  // Simple logic based on hasRed flag from main loop
  return hasRed ? FILL_AMPEL_RED : FILL_AMPEL_GREEN;
}

/** Utility: copy entire worksheet values and formatting for first 3 rows */
function cloneWorksheetValues(src, dst){
  // Copy column widths
  for (let c=1; c<=src.columnCount; c++){
    const w = src.getColumn(c).width;
    if (w) dst.getColumn(c).width = w;
  }
  
  // Copy all rows' values 1:1
  const last = src.lastRow ? src.lastRow.number : src.rowCount;
  for (let r=1; r<=last; r++){
    const sRow = src.getRow(r);
    const dRow = dst.getRow(r);
    dRow.values = sRow.values;
    
    // Preserve ALL formatting for the first 3 rows (Zeile 1, 2, 3)
    if (r <= 3) {
      for (let c=1; c<=sRow.cellCount; c++){
        const srcCell = sRow.getCell(c);
        const dstCell = dRow.getCell(c);
        
        // Copy ALL formatting properties
        if (srcCell.fill) dstCell.fill = srcCell.fill;
        if (srcCell.font) dstCell.font = srcCell.font;
        if (srcCell.border) dstCell.border = srcCell.border;
        if (srcCell.alignment) dstCell.alignment = srcCell.alignment;
        if (srcCell.numFmt) dstCell.numFmt = srcCell.numFmt;
        if (srcCell.style) dstCell.style = srcCell.style;
        
        // Copy merged cells if they exist
        if (srcCell.master && srcCell.master.address) {
          try {
            const masterAddr = srcCell.master.address;
            dst.mergeCells(masterAddr);
          } catch (e) {
            // Ignore merge errors
          }
        }
      }
    }
  }
}

/** Apply correct header structure for Qualitätsbericht */
function applyQualitaetsberichtHeaders(ws) {
  // Zeile 1: B1:X1 - "DB AG SAP R/3 K MARA Stammdaten Stand 20.Mai 2025"
  try {
    ws.unMergeCells('B1:X1');
  } catch (e) {}
  ws.mergeCells('B1:X1');
  const b1 = ws.getCell('B1');
  b1.value = 'DB AG SAP R/3 K MARA Stammdaten Stand 20.Mai 2025';
  b1.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  
  // Zeile 1: Y1 - "SAP Klassifizierung aus Okt24"
  const y1 = ws.getCell('Y1');
  y1.value = 'SAP Klassifizierung aus Okt24';
  y1.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  
  // Zeile 1: Z1:AB1 - "Zusatz Herstellerdaten aus Abfragen in 2024"
  try {
    ws.unMergeCells('Z1:AB1');
  } catch (e) {}
  ws.mergeCells('Z1:AB1');
  const z1 = ws.getCell('Z1');
  z1.value = 'Zusatz Herstellerdaten aus Abfragen in 2024';
  z1.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
}

/** Find column index by header name in HEADER_ROW (exact match after trim) */
function colByHeader(ws, name){
  const hdr = ws.getRow(HEADER_ROW);
  for (let c=1; c<=ws.columnCount; c++){
    const v = hdr.getCell(c).value;
    if (v != null && String(v).trim() === name) return c;
  }
  return null;
}

/**
 * Main function to check completeness of Excel data
 * @param {Buffer} fileBuffer - Excel file buffer
 * @returns {Buffer} - Processed Excel file buffer
 */
async function checkCompleteness(fileBuffer) {
  const inWb = new ExcelJS.Workbook();
  await inWb.xlsx.load(fileBuffer);

  // Use first worksheet of the uploaded workbook
  const src = inWb.worksheets[0];
  if (!src) throw new Error('Keine Tabelle im Workbook gefunden.');

  // Prepare output workbook with only one sheet: "Qualitätsbericht"
  const outWb = new ExcelJS.Workbook();
  const wsQ = outWb.addWorksheet('Qualitätsbericht');

  // Clone original values to Qualitätsbericht to preserve structure (no subheaders/structure changes)
  cloneWorksheetValues(src, wsQ);
  
  // Apply correct header structure for Qualitätsbericht
  applyQualitaetsberichtHeaders(wsQ);

  // Add Ampelbewertung column
  addQualityAmpelColumn(wsQ);

  // Map column indexes by header name (from row 3)
  const cFert = colByHeader(src, 'Fert./Prüfhinweis');
  const cL    = colByHeader(src, 'Länge');
  const cB    = colByHeader(src, 'Breite');
  const cH    = colByHeader(src, 'Höhe');
  
  // Gewichtsspalten nach Header-Namen (aus Zeile 3)
  const cNetto = colByHeader(src, 'Nettogewicht');
  const cBrutto = colByHeader(src, 'Bruttogewicht');

  // Iterate data rows (from row 4) - Optimized for large files
  const last = src.lastRow ? src.lastRow.number : FIRST_DATA_ROW - 1;
  const totalRows = last - FIRST_DATA_ROW + 1;
  console.log(`Processing ${totalRows} rows in batches...`);
  
  // Process in ultra-small batches for maximum stability
  const BATCH_SIZE = 20;
  
  for (let batchStart = FIRST_DATA_ROW; batchStart <= last; batchStart += BATCH_SIZE) {
    const batchEnd = Math.min(batchStart + BATCH_SIZE - 1, last);
    const progress = Math.round(((batchStart - FIRST_DATA_ROW) / totalRows) * 100);
    console.log(`Processing batch ${batchStart}-${batchEnd} (${progress}%)`);
    
    for (let r = batchStart; r <= batchEnd; r++) {
      const rowQ = wsQ.getRow(r);
      const rowS = src.getRow(r);

      if (!rowS || rowS.cellCount === 0) continue;

    let hasRed = false;

    // 1) Pflichtfelder (B–J, N, R–W): mark red if empty
    for (const {start, end} of MUST_COL_RANGES){
      for (let c = start; c <= end && c <= src.columnCount; c++){
        const v = rowS.getCell(c).value;
        if (isEmpty(v)){
          rowQ.getCell(c).fill = FILL_RED;
          hasRed = true;
        }
      }
    }

    // 2) Fert./Prüfhinweis invalid → rot
    if (cFert){
      const val = rowS.getCell(cFert).value;
      if (!isEmpty(val) && !validFertPruef(val)){
        rowQ.getCell(cFert).fill = FILL_RED;
        hasRed = true;
      }
    }

    // 3) Maße L/B/H: <0 → rot; all 0/empty → rot
    const vL = cL ? toNum(rowS.getCell(cL).value) : null;
    const vB = cB ? toNum(rowS.getCell(cB).value) : null;
    const vH = cH ? toNum(rowS.getCell(cH).value) : null;

    const markRed = (c) => { if (c){ rowQ.getCell(c).fill = FILL_RED; hasRed = true; } };

    if ([vL, vB, vH].some(v => v != null && v < 0)){
      markRed(cL); markRed(cB); markRed(cH);
    } else {
      const allZeroOrNone = [vL, vB, vH].every(v => v == null || v === 0);
      if (allZeroOrNone){
        markRed(cL); markRed(cB); markRed(cH);
      }
    }

    // 4) Nettogewicht <= 0 → rot (nur diese Spalte)
    if (cNetto){
      const g = toNum(rowS.getCell(cNetto).value);
      if (g != null && g <= 0){
        rowQ.getCell(cNetto).fill = FILL_RED;
        hasRed = true;
      }
    }
    
    // 5) Bruttogewicht <= 0 → rot (nur diese Spalte)
    if (cBrutto){
      const bg = toNum(rowS.getCell(cBrutto).value);
      if (bg != null && bg <= 0){
        rowQ.getCell(cBrutto).fill = FILL_RED;
        hasRed = true;
      }
    }
    
    // 6) Bruttogewicht < Nettogewicht → rot (nur die falsche Spalte)
    if (cBrutto && cNetto){
      const bg = toNum(rowS.getCell(cBrutto).value);
      const ng = toNum(rowS.getCell(cNetto).value);
      if (bg != null && ng != null && bg < ng){
        // Nur die Spalte markieren, die den falschen Wert hat
        // Wenn Bruttogewicht kleiner als Nettogewicht, dann ist Bruttogewicht falsch
        rowQ.getCell(cBrutto).fill = FILL_RED;
        hasRed = true;
      }
    }

    // 7) OPTIMIERT: Keine ganze Zeile grün färben - nur Ampelbewertung setzen
    // Das spart massiv Zeit und Speicher!
    
    // 8) Ampelbewertung in Spalte A: Rot wenn Probleme, Grün wenn alles OK
    const ampelFill = calculateQualityAmpelStatus(hasRed);
    rowQ.getCell(1).fill = ampelFill; // Column A
    }
    
    // Aggressives Memory cleanup nach jedem kleinen Batch
    if (global.gc) {
      global.gc();
    }
    
    // Kurze Pause für CPU-Entlastung bei schwacher Hardware
    await new Promise(resolve => setTimeout(resolve, 50));
  }

  console.log(`Completed processing ${totalRows} rows`);
  
  // Return the workbook as buffer
  return await outWb.xlsx.writeBuffer();
}

module.exports = {
  checkCompleteness
};
