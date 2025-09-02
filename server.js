const express = require('express');
const cors = require('cors');
const helmet = require('helmet');
const path = require('path');
const multer = require('multer');
const ExcelJS = require('exceljs');

const {
  toNumber,
  parseWeight,
  parseDimensionsToLBH,
  normPartNo,
  mapMaterialClassificationToExcel,
  normalizeNCode
} = require('./utils');
const { SiemensProductScraper } = require('./scraper');
const { checkCompleteness } = require('./completeness-checker');

const app = express();
const PORT = process.env.PORT || 3000;
const SCRAPE_CONCURRENCY = Number(process.env.SCRAPE_CONCURRENCY || 4);

const ORIGINAL_COLS = { Z:'Z', E:'E', C:'C', S:'S', T:'T', U:'U', V:'V', W:'W', P:'P', N:'N', AH:'AH' };

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

const HEADER_ROW = 3;
const LABEL_ROW = 4;
const FIRST_DATA_ROW = 5;

const AMPEL_REQUIRED_ORIGINALS = ['E','N','S','U','V','W'];

app.use(helmet({ contentSecurityPolicy: false }));
app.use(cors());
app.use(express.json({ limit: '2mb' }));
app.use(express.static(__dirname));

const scraper = new SiemensProductScraper();

// --- Helpers ---
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
function calculateNewColumnStructure(ws) {
  const newStructure = { pairs: [], otherCols: new Map(), totalInsertedCols: 0 };
  let insertedCols = 0;
  for (const pair of DB_WEB_PAIRS) {
    const originalIndex = getColumnIndex(pair.original);
    const adjustedOriginalIndex = originalIndex + insertedCols;
    pair.dbCol  = getColumnLetter(adjustedOriginalIndex);
    pair.webCol = getColumnLetter(adjustedOriginalIndex + 1);
    newStructure.pairs.push({ ...pair });
    insertedCols++;
  }
  newStructure.totalInsertedCols = insertedCols;
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
function fillColor(ws, addr, color) {
  const map = {
    green:  'FFD5F4E6',
    red:    'FFFDEAEA',
    orange: 'FFFFEAA7',
    dbBlue: 'FFE6F3FF',
    webBlue:'FFCCE7FF'
  };
  ws.getCell(addr).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: map[color] } };
}
function applyLabelCellFormatting(ws, addr, isWebCell = false) {
  fillColor(ws, addr, isWebCell ? 'webBlue' : 'dbBlue');
  const cell = ws.getCell(addr);
  cell.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
  cell.font = { bold: true, size: 10 };
  cell.alignment = { horizontal: 'center', vertical: 'middle' };
}
function hasValue(v){ return v!==null && v!==undefined && v!=='' && String(v).trim()!==''; }
function eqText(a,b){
  if (a==null||b==null) return false;
  return String(a).trim().toLowerCase().replace(/\s+/g,' ') === String(b).trim().toLowerCase().replace(/\s+/g,' ');
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
  return webVal!=null && exNum===webVal;
}

// --- Haupt-Route (gekürzt auf Kernänderungen, Rest gleich wie vorher) ---
app.post('/api/process-excel', multer({ storage: multer.memoryStorage() }).single('file'), async (req,res)=>{
  if (!req.file) return res.status(400).json({error:'Bitte Datei hochladen.'});
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(req.file.buffer);

  const tasks=[]; const rowsPerSheet=new Map();
  for (const ws of wb.worksheets) {
    const idx=[]; const last=ws.lastRow?.number||0;
    for (let r=FIRST_DATA_ROW-1;r<=last;r++) {
      const a2v=(ws.getCell(`${ORIGINAL_COLS.Z}${r}`).value||'').toString().trim().toUpperCase();
      if (a2v.startsWith('A2V')) { idx.push(r); tasks.push(a2v); }
    }
    rowsPerSheet.set(ws,idx);
  }
  const resultsMap=await scraper.scrapeMany(tasks,SCRAPE_CONCURRENCY);

  for (const ws of wb.worksheets) {
    const structure=calculateNewColumnStructure(ws);
    for (const pair of [...structure.pairs].reverse()) {
      ws.spliceColumns(getColumnIndex(pair.original)+1,0,[null]);
    }
    ws.spliceRows(LABEL_ROW,0,[null]);
    for (const pair of structure.pairs) {
      ws.getCell(`${pair.dbCol}${LABEL_ROW}`).value='DB-Wert';
      ws.getCell(`${pair.webCol}${LABEL_ROW}`).value='Web-Wert';
      applyLabelCellFormatting(ws,`${pair.dbCol}${LABEL_ROW}`,false);
      applyLabelCellFormatting(ws,`${pair.webCol}${LABEL_ROW}`,true);
    }
    // Ampelspalte
    const ampCol=getColumnLetter(ws.columnCount+1);
    ws.spliceColumns(ws.columnCount+1,0,[null]);
    ws.getCell(`${ampCol}2`).value='AMP';
    ws.getCell(`${ampCol}3`).value='Ampelbewertung';
    ws.getCell(`${ampCol}${LABEL_ROW}`).value='Status';
    applyLabelCellFormatting(ws,`${ampCol}${LABEL_ROW}`,false);

    const prodRows=rowsPerSheet.get(ws)||[];
    for (const origRow of prodRows) {
      const r=origRow+1;
      let hasRed=false;

      let zCol=ORIGINAL_COLS.Z;
      if (structure.otherCols.has('Z')) zCol=structure.otherCols.get('Z');
      const a2v=(ws.getCell(`${zCol}${r}`).value||'').toString().trim().toUpperCase();

      let ahCol=ORIGINAL_COLS.AH;
      if (structure.otherCols.has('AH')) ahCol=structure.otherCols.get('AH');
      ws.getCell(`${ahCol}${r}`).value=a2v;

      const web=resultsMap.get(a2v)||{};
      for (const pair of structure.pairs) {
        const dbVal=ws.getCell(`${pair.dbCol}${r}`).value;
        let webVal=null; let isEqual=false;
        switch(pair.original){
          case 'E': {
            const dbStr=(dbVal||'').toString().toUpperCase();
            if (dbStr.startsWith('A2V')) webVal=a2v;
            else webVal=(web['Weitere Artikelnummer']&&web['Weitere Artikelnummer']!=='Nicht gefunden')?web['Weitere Artikelnummer']:a2v;
            isEqual=eqPart(dbVal||a2v,webVal); break;
          }
          // ... andere Fälle gleich wie bisher ...
        }
        const hasDb=hasValue(dbVal); const hasWeb=webVal!=null;
        if (hasWeb) {
          ws.getCell(`${pair.webCol}${r}`).value=webVal;
          if (hasDb) {
            const ok=isEqual;
            fillColor(ws,`${pair.webCol}${r}`,ok?'green':'red');
            if (AMPEL_REQUIRED_ORIGINALS.includes(pair.original)&&!ok) hasRed=true;
          }
        } else {
          fillColor(ws,`${pair.webCol}${r}`,'orange');
        }
      }
      // Ampelzelle: nur Farbe, kein Text
      fillColor(ws,`${ampCol}${r}`,hasRed?'red':'green');
    }

    // AutoFilter aktivieren (Headerzeile = LABEL_ROW)
    ws.autoFilter={ from:{row:LABEL_ROW,column:1}, to:{row:LABEL_ROW,column:ws.columnCount} };
  }

  const out=await wb.xlsx.writeBuffer();
  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition','attachment; filename="Web_Vergleich_Ergebnis.xlsx"');
  res.send(Buffer.from(out));
});

app.listen(PORT,()=>console.log(`Server läuft auf Port ${PORT}`));
