/***** НАСТРОЙКИ — addNewDayBlock *****/
const VALID_SHEETS = ["POTTERY", "PAINTS", "SUPPLIES"];
const DATA_START_ROW = 4;

const HDR_ROW_1 = 1; // дата (merged)
const HDR_ROW_2 = 2; // пустая merged-строка
const HDR_ROW_3 = 3; // IN/OUT + Total

/***** НАСТРОЙКИ — Shelf_Transfer *****/
const ST_STORAGE_SHEET = 'POTTERY';
const ST_ORDER_SHEET   = 'OUT';
const ST_COLORS_SHEET  = 'PAINTS';

const ST_SHELF_HEADER  = 'Shelf #';
// 'Item #' добавлен — новый заголовок артикула в листе OUT
const ST_ORDER_SKU_HEADERS = ['Item #','מק״ט','מק\'\'ט','מק""ט','מק\'ט','מק"ט'];
// Shelf# вставляется перед первой статусной колонкой
const ST_STATUS_HEADER = 'סטטוס pottery';

const ST_HEADER_SCAN_MAX_ROWS = 10;

const ST_DATA_START_ROW_STORAGE = 4;
const ST_DATA_START_ROW_COLORS  = 4;

const ST_BLOCK_HEADER_ROW = 3;

const ST_NOTE_ANCHOR_PREFIX = 'הערות להזמנה:';

/***** НАСТРОЙКИ — Transfer OUT *****/
// Заголовки колонок листа OUT (строка 1) — менять здесь при переименовании
// Порядок не важен: поиск идёт по имени, не по позиции
const OUT_HEADER_ITEM           = 'Item #';          // артикул (SKU)
const OUT_HEADER_DESC           = 'Description';     // название товара
const OUT_HEADER_QTY            = 'Qty';             // количество
const OUT_HEADER_STATUS_POTTERY = 'סטטוס pottery';   // статус для POTTERY: пусто = не переносить
const OUT_HEADER_STATUS_PAINTS  = 'סטטוס paints';    // статус для PAINTS:  пусто = не переносить

/***** НАСТРОЙКИ — Transfer IN *****/
const IN_SHEET_NAME    = 'IN';
const TARGET_SHEETS_IN = ['POTTERY', 'PAINTS'];

// Заголовки колонок листа IN (строка 1) — менять здесь при переименовании
const IN_HEADER_ITEM           = 'Item #';          // артикул (SKU)
const IN_HEADER_DESC           = 'Description';     // название товара
const IN_HEADER_QTY            = 'Qty';             // количество
const IN_HEADER_STATUS_POTTERY = 'סטטוס pottery';   // статус для POTTERY: пусто = не переносить
const IN_HEADER_STATUS_PAINTS  = 'סטטוס paints';    // статус для PAINTS:  пусто = не переносить


/***********************************************************************
 * УТИЛИТЫ — addNewDayBlock
 ***********************************************************************/
function columnToLetter(column) {
  let letter = "";
  while (column > 0) {
    const temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function _norm(v){ return String(v || "").toLowerCase().trim(); }

function getTotalNowCol(sheet) {
  const header1 = sheet.getRange(HDR_ROW_1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idx = header1.findIndex(h => String(h).trim() === "TOTAL NOW");
  if (idx === -1) throw new Error("Column \"TOTAL NOW\" not found");
  return idx + 1;
}

function getFirstBlockStartCol(sheet) {
  const lastCol = sheet.getLastColumn();
  if (!lastCol) return 6;
  const header3 = sheet.getRange(HDR_ROW_3, 1, 1, lastCol).getValues()[0].map(_norm);
  const firstTotalIdx0 = header3.findIndex(v => v === "total");
  if (firstTotalIdx0 > 0) return firstTotalIdx0;
  return 6;
}

function buildTotalNowFormulaA1(row, firstBlockStartCol, lastCol) {
  const firstColLetter = columnToLetter(firstBlockStartCol);
  const lastColLetter  = columnToLetter(lastCol);
  return `=IFERROR(INDEX(${row}:${row},
    MAX(FILTER(COLUMN(${firstColLetter}$${HDR_ROW_3}:${lastColLetter}$${HDR_ROW_3}),
    ${firstColLetter}$${HDR_ROW_3}:${lastColLetter}$${HDR_ROW_3}="Total"))), "")`;
}

function enforceNumberFormatForColumn(sheet, startRow, col, numRows) {
  if (numRows <= 0) return;
  sheet.getRange(startRow, col, numRows, 1).setNumberFormat("0.############");
}

function applyNumberValidation(range, minZero = true, allowInvalid = true) {
  const b = SpreadsheetApp.newDataValidation();
  const builder = minZero
    ? b.requireNumberGreaterThanOrEqualTo(0)
    : b.requireNumberBetween(-1e12, 1e12);
  range.setDataValidation(builder.setAllowInvalid(allowInvalid).build());
}

function removeProtections(sheet) {
  try { sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove()); } catch(e){}
  try { sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove()); } catch(e){}
}

function ensureThreeRowHeader(sheet) {
  const lastCol = sheet.getLastColumn();
  if (!lastCol) return;
  const r2 = sheet.getRange(2, 1, 1, lastCol).getValues()[0].map(_norm);
  const r3 = sheet.getRange(3, 1, 1, lastCol).getValues()[0].map(_norm);
  const r2Has = r2.some(v => v === "total" || v.startsWith("in") || v.startsWith("out"));
  const r3Has = r3.some(v => v === "total" || v.startsWith("in") || v.startsWith("out"));
  if (r2Has && !r3Has) sheet.insertRowAfter(1);
}

function normalizeHeaderRow3Alignment(sheetName) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sh) return;
  const lastCol = sh.getLastColumn();
  if (!lastCol) return;
  sh.getRange(HDR_ROW_3, 1, 1, lastCol).setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  sh.getRange(HDR_ROW_2, 1, 1, lastCol).setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  try { sh.autoResizeRows(HDR_ROW_2, 2); } catch(e) {}
}

function relaxValidationForInputColumns(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < DATA_START_ROW || !lastCol) return;
  const h3 = sheet.getRange(HDR_ROW_3, 1, 1, lastCol).getValues()[0].map(_norm);
  const rowsCount = lastRow - DATA_START_ROW + 1;
  for (let c = 1; c <= lastCol; c++) {
    if (h3[c - 1].startsWith("in") || h3[c - 1].startsWith("out")) {
      applyNumberValidation(sheet.getRange(DATA_START_ROW, c, rowsCount, 1), true, true);
    }
  }
}

function applyBlockFontWeights(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (!lastCol) return;
  sheet.getRange(HDR_ROW_1, 1, 1, lastCol).setFontWeight("bold");
  sheet.getRange(HDR_ROW_2, 1, 1, lastCol).setFontWeight("bold");
  sheet.getRange(HDR_ROW_3, 1, 1, lastCol).setFontWeight("bold");
  const header1 = sheet.getRange(HDR_ROW_1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  let totalNowCol = header1.indexOf("TOTAL NOW") + 1;
  if (!totalNowCol) return;
  sheet.getRange(HDR_ROW_1, totalNowCol, lastRow, 1).setFontWeight("bold");
  if (lastRow >= DATA_START_ROW) {
    const rowsCount = lastRow - DATA_START_ROW + 1;
    const startColForBlocks = totalNowCol + 1;
    if (startColForBlocks <= lastCol) {
      sheet.getRange(DATA_START_ROW, startColForBlocks, rowsCount, lastCol - startColForBlocks + 1).setFontWeight("normal");
    }
  }
}

function applyTotalNowConditionalFormatting(sheet, totalNowCol) {
  const startRow = DATA_START_ROW;
  const lastRow = sheet.getLastRow();
  const rowsCount = Math.max(lastRow - startRow + 1, 1);
  const totalNowRange = sheet.getRange(startRow, totalNowCol, rowsCount, 1);
  const rules = sheet.getConditionalFormatRules() || [];
  const filtered = rules.filter(r => {
    const rngs = r.getRanges();
    if (!rngs || !rngs.length) return true;
    return !rngs.some(rg => {
      if (rg.getSheet().getName() !== sheet.getName()) return false;
      const col = rg.getColumn();
      const width = rg.getNumColumns();
      return (col <= totalNowCol && (col + width - 1) >= totalNowCol);
    });
  });
  const orangeRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=AND($${columnToLetter(totalNowCol)}${startRow}>=10,$${columnToLetter(totalNowCol)}${startRow}<20)`)
    .setFontColor("#FFA500").setRanges([totalNowRange]).build();
  const redRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(10).setFontColor("#FF0000").setRanges([totalNowRange]).build();
  sheet.setConditionalFormatRules([...filtered, orangeRule, redRule]);
}


/***********************************************************************
 * УТИЛИТЫ — Shelf_Transfer
 ***********************************************************************/
function st_normKey(v){ return (v === null || v === undefined) ? '' : String(v).trim(); }

function st_clean(s){
  return String(s || '')
    .replace(/[\u200E\u200F\u202A-\u202E]/g, '')
    .replace(/[""„"]/g, '"')
    .replace(/[׳׳′""]/g, "''")
    .trim().toLowerCase();
}

function st_findHeaderPosition(sheet, headerTargets, maxRows = ST_HEADER_SCAN_MAX_ROWS){
  const targets = (Array.isArray(headerTargets)?headerTargets:[headerTargets]).map(st_clean);
  const lastCol = sheet.getLastColumn();
  const scanRows = Math.min(maxRows, Math.max(1, sheet.getLastRow()));
  for (let r=1; r<=scanRows; r++){
    const rowVals = sheet.getRange(r,1,1,lastCol).getDisplayValues()[0];
    for (let c=1;c<=lastCol;c++){
      if (targets.includes(st_clean(rowVals[c-1]))) return {row:r,col:c};
    }
  }
  return {row:-1,col:-1};
}

function st_findHeaderInRow(sheet, headerTargets, row){
  if (row<1) return -1;
  const targets = (Array.isArray(headerTargets)?headerTargets:[headerTargets]).map(st_clean);
  const lastCol = sheet.getLastColumn();
  const rowVals = sheet.getRange(row,1,1,lastCol).getDisplayValues()[0].map(st_clean);
  const idx = rowVals.findIndex(v=>targets.includes(v));
  return (idx>=0)?(idx+1):-1;
}

function st_getSafeSortStartRow(sheet, dataStartRow){
  let start = dataStartRow;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < dataStartRow) return start;
  const merges = sheet.getRange(1,1,lastRow,lastCol).getMergedRanges() || [];
  for (const mr of merges){
    const r0 = mr.getRow();
    const r1 = r0 + mr.getNumRows() - 1;
    if (r1 >= dataStartRow && r0 < start) start = r1 + 1;
  }
  return start;
}

function st_findAnchoredRows(sheet, startRow){
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (startRow > lastRow) return [];
  const values = sheet.getRange(startRow,1,lastRow-startRow+1,lastCol).getDisplayValues();
  const rows = [];
  for (let i=0;i<values.length;i++){
    if (values[i].some(v=>String(v||'').trim().startsWith(ST_NOTE_ANCHOR_PREFIX))) rows.push(startRow+i);
  }
  return rows;
}

function st_ensureThreeRowHeader(sheet){
  const lastCol = sheet.getLastColumn();
  if (!lastCol) return;
  const r2 = sheet.getRange(2, 1, 1, lastCol).getDisplayValues()[0].map(st_clean);
  const r3 = sheet.getRange(3, 1, 1, lastCol).getDisplayValues()[0].map(st_clean);
  const r2Has = r2.some(v => v === 'total' || v === 'in' || v === 'out');
  const r3Has = r3.some(v => v === 'total' || v === 'in' || v === 'out');
  if (r2Has && !r3Has) sheet.insertRowAfter(1);
}


/***********************************************************************
 * ОСНОВНЫЕ ФУНКЦИИ — addNewDayBlock
 ***********************************************************************/
function addNewDayBlock(mode, sheetName) {
  if (!VALID_SHEETS.includes(sheetName)) return;
  mode = String(mode || "").toUpperCase().trim();
  if (!["IN", "OUT"].includes(mode)) throw new Error(`Invalid mode: "${mode}".`);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet "${sheetName}" not found`);

  removeProtections(sheet);
  ensureThreeRowHeader(sheet);
  normalizeHeaderRow3Alignment(sheetName);

  const totalRows = sheet.getLastRow();
  const lastColBefore = sheet.getLastColumn();
  const prevTotalCol = lastColBefore;
  const date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

  sheet.insertColumnsAfter(lastColBefore, 2);
  const startCol = lastColBefore + 1;
  const valueCol = startCol;
  const totalCol = startCol + 1;
  sheet.getRange(HDR_ROW_1, startCol, 1, 2).merge().setValue(date).setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.getRange(HDR_ROW_2, startCol, 1, 2).merge().setValue("").setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.getRange(HDR_ROW_3, valueCol).setValue(mode);
  sheet.getRange(HDR_ROW_3, totalCol).setValue("Total");
  if (mode === "IN") sheet.getRange(HDR_ROW_3, valueCol).setBackground("#ffeb3b");

  if (totalRows >= DATA_START_ROW) {
    const rowsCount = totalRows - DATA_START_ROW + 1;
    const inputRange = sheet.getRange(DATA_START_ROW, valueCol, rowsCount, 1);
    inputRange.clearContent();
    enforceNumberFormatForColumn(sheet, DATA_START_ROW, valueCol, rowsCount);
    applyNumberValidation(inputRange, true, true);
    const sign = (mode === "IN") ? "+" : "-";
    const formulas = [];
    for (let r = DATA_START_ROW; r <= totalRows; r++) {
      const prev = sheet.getRange(r, prevTotalCol).getA1Notation();
      const val  = sheet.getRange(r, valueCol).getA1Notation();
      formulas.push([`=IFERROR(VALUE(TRIM(${prev})),0)${sign}IFERROR(VALUE(TRIM(${val})),0)`]);
    }
    sheet.getRange(DATA_START_ROW, totalCol, rowsCount, 1).setFormulas(formulas);
    enforceNumberFormatForColumn(sheet, DATA_START_ROW, totalCol, rowsCount);
    const totalNowCol = getTotalNowCol(sheet);
    const firstBlockStartCol = getFirstBlockStartCol(sheet);
    const lastCol = sheet.getLastColumn();
    const totalNowFormulas = [];
    for (let r = DATA_START_ROW; r <= totalRows; r++) {
      totalNowFormulas.push([buildTotalNowFormulaA1(r, firstBlockStartCol, lastCol)]);
    }
    sheet.getRange(DATA_START_ROW, totalNowCol, rowsCount, 1).setFormulas(totalNowFormulas);
    applyBlockFontWeights(sheet);
    applyTotalNowConditionalFormatting(sheet, totalNowCol);
  }

  const maxRows = sheet.getMaxRows();
  sheet.getRange(1, startCol, maxRows, 1).setBorder(null, true, null, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  SpreadsheetApp.flush();
}

function ensureFormulasInAllTotals(sheet) {
  const totalRows = sheet.getLastRow();
  const totalCols = sheet.getLastColumn();
  if (totalRows < DATA_START_ROW || totalCols === 0) return;
  const h3raw = sheet.getRange(HDR_ROW_3, 1, 1, totalCols).getValues()[0];
  const h3 = h3raw.map(_norm);
  const rowsCount = totalRows - DATA_START_ROW + 1;
  for (let j = 1; j <= totalCols; j++) {
    if (h3[j - 1] !== "total") continue;
    const label = _norm(h3raw[j - 2] || "");
    const sign = label.startsWith("in") ? "+" : label.startsWith("out") ? "-" : null;
    if (!sign) continue;
    const r1c1 = `=IFERROR(VALUE(TRIM(R[0]C[-2])),0)${sign}IFERROR(VALUE(TRIM(R[0]C[-1])),0)`;
    sheet.getRange(DATA_START_ROW, j, rowsCount, 1).setFormulasR1C1(Array.from({ length: rowsCount }, () => [r1c1]));
  }
  applyBlockFontWeights(sheet);
}

function updateDayBlock(sheetName) {
  if (!VALID_SHEETS.includes(sheetName)) return;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet "${sheetName}" not found`);
  removeProtections(sheet);
  ensureThreeRowHeader(sheet);
  normalizeHeaderRow3Alignment(sheetName);
  const totalRows = sheet.getLastRow();
  const totalCols = sheet.getLastColumn();
  if (totalRows < DATA_START_ROW || totalCols === 0) return;
  const h3raw = sheet.getRange(HDR_ROW_3, 1, 1, totalCols).getValues()[0];
  const h3 = h3raw.map(_norm);
  const totalIndices = [];
  for (let i = 0; i < h3.length; i++) if (h3[i] === "total") totalIndices.push(i + 1);
  if (!totalIndices.length) return;
  const newestTotalCol = totalIndices.pop();
  const rowsCount = totalRows - DATA_START_ROW + 1;
  const label = _norm(h3raw[newestTotalCol - 2] || "");
  const sign = label.startsWith("in") ? "+" : label.startsWith("out") ? "-" : null;
  if (sign) {
    const r1c1 = `=IFERROR(VALUE(TRIM(R[0]C[-2])),0)${sign}IFERROR(VALUE(TRIM(R[0]C[-1])),0)`;
    sheet.getRange(DATA_START_ROW, newestTotalCol, rowsCount, 1).setFormulasR1C1(Array.from({ length: rowsCount }, () => [r1c1]));
  }
  ensureFormulasInAllTotals(sheet);
  const totalNowCol = getTotalNowCol(sheet);
  const firstBlockStartCol = getFirstBlockStartCol(sheet);
  const lastCol = sheet.getLastColumn();
  const totalNowFormulas = [];
  for (let r = DATA_START_ROW; r <= totalRows; r++) {
    totalNowFormulas.push([buildTotalNowFormulaA1(r, firstBlockStartCol, lastCol)]);
  }
  sheet.getRange(DATA_START_ROW, totalNowCol, rowsCount, 1).setFormulas(totalNowFormulas);
  applyBlockFontWeights(sheet);
  applyTotalNowConditionalFormatting(sheet, totalNowCol);
  SpreadsheetApp.flush();
}

function ensureActionDropdown(sheet) {
  const name = sheet.getName();
  const actions = name === 'POTTERY'
    ? ["➕ IN", "➕ OUT", "🔁 Update block", "🔤 Sort A→Z"]
    : ["➕ IN", "➕ OUT", "🔁 Update block"];
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(actions, true).setAllowInvalid(false).build();
  sheet.getRange("A1").setDataValidation(rule).setValue("");
}

function sortByDescriptionPottery() {
  const sheetName = 'POTTERY';
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < DATA_START_ROW) return;
  const rangeToSort = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, lastCol);
  rangeToSort.sort({ column: 2, ascending: true });
  SpreadsheetApp.getActiveSpreadsheet().toast('Sheet POTTERY sorted A → Z by DESCRIPTION (column B).');
}

function fixAllTotalsAndTotalNow(sheetName) {
  if (!VALID_SHEETS.includes(sheetName)) return;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return;
  ensureThreeRowHeader(sheet);
  const totalRows = sheet.getLastRow();
  if (totalRows < DATA_START_ROW) return;
  ensureFormulasInAllTotals(sheet);
  const totalNowCol = getTotalNowCol(sheet);
  const firstBlockStartCol = getFirstBlockStartCol(sheet);
  const lastCol = sheet.getLastColumn();
  const rowsCount = totalRows - DATA_START_ROW + 1;
  const formulas = [];
  for (let r = DATA_START_ROW; r <= totalRows; r++) {
    formulas.push([buildTotalNowFormulaA1(r, firstBlockStartCol, lastCol)]);
  }
  sheet.getRange(DATA_START_ROW, totalNowCol, rowsCount, 1).setFormulas(formulas);
  applyBlockFontWeights(sheet);
  applyTotalNowConditionalFormatting(sheet, totalNowCol);
}


/***********************************************************************
 * ЗАДАНИЕ 2 — Проверка одного листа перед переносом
 ***********************************************************************/
function _checkSheetForTransfer(sheetName, expectedMode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet)
    return { status: 'critical', message: `⛔ Sheet "${sheetName}": not found.` };

  const lastCol = sheet.getLastColumn();
  if (!lastCol)
    return { status: 'critical', message: `⛔ Sheet "${sheetName}": no data blocks found.` };

  const row3Header = sheet.getRange(HDR_ROW_3, 1, 1, lastCol).getValues()[0];
  let targetCol = -1;
  let lastBlockLabel = "";

  for (let c = lastCol; c >= 1; c--) {
    if (String(row3Header[c - 1]).trim().toLowerCase() === 'total') {
      const potentialCol = c - 1;
      lastBlockLabel = String(row3Header[potentialCol - 1]).trim().toUpperCase();
      if (lastBlockLabel === expectedMode) targetCol = potentialCol;
      break;
    }
  }

  if (targetCol === -1 && !lastBlockLabel)
    return { status: 'critical', message: `⛔ Sheet "${sheetName}": no data blocks found.` };

  if (targetCol === -1)
    return {
      status: 'skip',
      message: `⚠️ "${sheetName}": last block is "${lastBlockLabel}", expected "${expectedMode}".\n` +
               `   Add a new ${expectedMode} block using the menu in A1.`
    };

  const lastRow = sheet.getLastRow();
  const numRows = lastRow - DATA_START_ROW + 1;
  if (numRows > 0) {
    const checkVals = sheet.getRange(DATA_START_ROW, targetCol, numRows, 1).getValues();
    if (checkVals.some(r => String(r[0]).trim() !== ""))
      return {
        status: 'skip',
        message: `⚠️ "${sheetName}": ${expectedMode} block already has data.\n` +
                 `   Clear it before transferring.`
      };
  }

  return { status: 'ok', targetCol };
}

/***********************************************************************
 * ЗАДАНИЕ 2 — Итоговый план переноса
 ***********************************************************************/
function _resolveTransferPlan(sheetNames, expectedMode) {
  const results = sheetNames.map(name => ({ name, check: _checkSheetForTransfer(name, expectedMode) }));

  const criticals = results.filter(r => r.check.status === 'critical');
  const skips     = results.filter(r => r.check.status === 'skip');
  const oks       = results.filter(r => r.check.status === 'ok');

  if (criticals.length > 0) {
    return {
      action: 'stop',
      message: [...criticals, ...skips].map(r => r.check.message).join("\n\n"),
      sheetsToProcess: []
    };
  }

  if (oks.length === 0) {
    return {
      action: 'stop',
      message: skips.map(r => r.check.message).join("\n\n"),
      sheetsToProcess: []
    };
  }

  if (skips.length > 0) {
    return {
      action: 'partial',
      message: skips.map(r => r.check.message).join("\n\n"),
      sheetsToProcess: oks.map(r => ({ name: r.name, targetCol: r.check.targetCol }))
    };
  }

  return {
    action: 'ask',
    message: '',
    sheetsToProcess: oks.map(r => ({ name: r.name, targetCol: r.check.targetCol }))
  };
}

/***********************************************************************
 * ДИАЛОГ ВЫБОРА НАПРАВЛЕНИЯ ПЕРЕНОСА
 ***********************************************************************/
function _askWhereToTransfer(mode) {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    `Transfer ${mode} — куда переносить?`,
    '1 — только POTTERY\n2 — только PAINTS\n3 — в оба листа',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return null;

  const val = response.getResponseText().trim();
  if (val === '1') return 'POTTERY';
  if (val === '2') return 'PAINTS';
  if (val === '3') return 'BOTH';

  ui.alert('Invalid input. Enter 1, 2 or 3 and click OK.');
  return null;
}

function _applyTargetChoice(sheetsToProcess, choice) {
  if (choice === 'BOTH') return sheetsToProcess;
  return sheetsToProcess.filter(s => s.name === choice);
}


/***********************************************************************
 * ОСНОВНЫЕ ФУНКЦИИ — Transfer IN (из листа 'IN')
 ***********************************************************************/
function transferFromInSheet_ST() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const shIn = ss.getSheetByName(IN_SHEET_NAME);
  if (!shIn) { ui.alert(`Error: Sheet "${IN_SHEET_NAME}" not found.`); return; }

  const lastRowIn = shIn.getLastRow();
  if (lastRowIn < 2) { ss.toast('No data found on the IN sheet.'); return; }

  // ── Находим колонки листа IN по заголовкам строки 1 ──────────────
  const inLastCol          = shIn.getLastColumn();
  const inHeaders          = shIn.getRange(1, 1, 1, inLastCol).getValues()[0];
  const inColItem          = inHeaders.findIndex(h => String(h).trim().toLowerCase() === IN_HEADER_ITEM.toLowerCase());
  const inColDesc          = inHeaders.findIndex(h => String(h).trim().toLowerCase() === IN_HEADER_DESC.toLowerCase());
  const inColQty           = inHeaders.findIndex(h => String(h).trim().toLowerCase() === IN_HEADER_QTY.toLowerCase());
  // ── ИЗМЕНЕНИЕ 1: ищем колонку Shelf # в листе IN ─────────────────
  const inColShelf         = inHeaders.findIndex(h => String(h).trim() === ST_SHELF_HEADER);
  // ─────────────────────────────────────────────────────────────────
  const inColStatusPottery = inHeaders.findIndex(h => String(h).trim() === IN_HEADER_STATUS_POTTERY);
  const inColStatusPaints  = inHeaders.findIndex(h => String(h).trim() === IN_HEADER_STATUS_PAINTS);

  // Обязательные колонки
  if (inColItem < 0 || inColDesc < 0 || inColQty < 0) {
    const missing = [
      inColItem < 0 ? `"${IN_HEADER_ITEM}"` : null,
      inColDesc < 0 ? `"${IN_HEADER_DESC}"` : null,
      inColQty  < 0 ? `"${IN_HEADER_QTY}"` : null
    ].filter(Boolean).join(', ');
    ui.alert(`Sheet "${IN_SHEET_NAME}": required headers ${missing} not found in row 1.`);
    return;
  }

  // Хотя бы одна статусная колонка должна существовать
  if (inColStatusPottery < 0 && inColStatusPaints < 0) {
    ui.alert(
      `Sheet "${IN_SHEET_NAME}": columns\n` +
      `"${IN_HEADER_STATUS_POTTERY}" and "${IN_HEADER_STATUS_PAINTS}" not found.\n` +
      `Check the headers in row 1.`
    );
    return;
  }

  // ── Спрашиваем что переносим ──────────────────────────────────────
  const promptLines = [];
  if (inColStatusPottery >= 0) promptLines.push(`1 — ${IN_HEADER_STATUS_POTTERY}`);
  if (inColStatusPaints  >= 0) promptLines.push(`2 — ${IN_HEADER_STATUS_PAINTS}`);

  const response = ui.prompt(
    'Transfer IN — what to transfer?',
    promptLines.join('\n'),
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) { ss.toast('Transfer cancelled.'); return; }

  const choice = response.getResponseText().trim();
  let selectedStatusCol, targetSheetName, selectedLabel;

  if (choice === '1') {
    if (inColStatusPottery < 0) { ui.alert(`Column "${IN_HEADER_STATUS_POTTERY}" not found in row 1.`); return; }
    selectedStatusCol = inColStatusPottery;
    targetSheetName   = ST_STORAGE_SHEET;       // POTTERY
    selectedLabel     = IN_HEADER_STATUS_POTTERY;
  } else if (choice === '2') {
    if (inColStatusPaints < 0) { ui.alert(`Column "${IN_HEADER_STATUS_PAINTS}" not found in row 1.`); return; }
    selectedStatusCol = inColStatusPaints;
    targetSheetName   = ST_COLORS_SHEET;        // PAINTS
    selectedLabel     = IN_HEADER_STATUS_PAINTS;
  } else {
    ui.alert('Invalid input. Enter 1 or 2 and click OK.');
    return;
  }

  // ── Считываем данные IN, фильтруем по выбранной статусной колонке ─
  const inData = shIn.getRange(2, 1, lastRowIn - 1, inLastCol).getValues();
  const inMap  = new Map();
  inData.forEach(row => {
    const sku    = String(row[inColItem]          || '').trim().toLowerCase();
    const status = String(row[selectedStatusCol]  || '').trim();
    if (!sku || !status) return;
    // ── ИЗМЕНЕНИЕ 2: сохраняем значение shelf в inMap ────────────────
    inMap.set(sku, {
      name:  row[inColDesc],
      qty:   row[inColQty],
      shelf: inColShelf >= 0 ? row[inColShelf] : ''
    });
    // ─────────────────────────────────────────────────────────────────
  });

  if (inMap.size === 0) {
    ui.alert(`Nothing to transfer.\nAll rows in column "${selectedLabel}" are empty.`);
    return;
  }

  // ── Валидируем целевой лист ───────────────────────────────────────
  const check = _checkSheetForTransfer(targetSheetName, 'IN');

  if (check.status === 'critical') {
    ui.alert(`Transfer IN failed:\n\n${check.message}`);
    return;
  }
  if (check.status === 'skip') {
    ui.alert(check.message);
    return;
  }

  // ── Выполняем перенос ─────────────────────────────────────────────
  const targetCol     = check.targetCol;
  const sheet         = ss.getSheetByName(targetSheetName);
  const lastCol       = sheet.getLastColumn();
  const lastRowTarget = sheet.getLastRow();
  const numRows       = lastRowTarget - DATA_START_ROW + 1;

  // ── ИЗМЕНЕНИЕ 3: ищем Shelf # только для POTTERY (в PAINTS её нет) ─
  let shelfColTarget = -1;
  if (targetSheetName === ST_STORAGE_SHEET && inColShelf >= 0) {
    const shelfPos = st_findHeaderPosition(sheet, ST_SHELF_HEADER, ST_HEADER_SCAN_MAX_ROWS);
    shelfColTarget = shelfPos.col; // останется -1, если колонки нет
  }
  // ─────────────────────────────────────────────────────────────────

  // SKU in column A (ITEM)
  const skuTargetVals = numRows > 0
    ? sheet.getRange(DATA_START_ROW, 1, numRows, 1).getValues()
    : [];
  const existingSkus = new Set(skuTargetVals.map(r => String(r[0] || '').trim().toLowerCase()));

  // ── ИЗМЕНЕНИЕ 4: обновляем существующие строки + пишем shelf ──────
  let totalApplied = 0, totalNewRows = 0;
  if (numRows > 0) {
    const output      = [];
    const shelfOutput = shelfColTarget > 0 ? [] : null;
    for (let i = 0; i < numRows; i++) {
      const sku  = String(skuTargetVals[i][0] || '').trim().toLowerCase();
      const data = inMap.get(sku);
      output.push([data ? data.qty : '']);
      if (shelfOutput !== null) shelfOutput.push([data ? data.shelf : '']);
      if (data) totalApplied++;
    }
    sheet.getRange(DATA_START_ROW, targetCol, numRows, 1).setValues(output);
    if (shelfOutput !== null && shelfColTarget > 0) {
      sheet.getRange(DATA_START_ROW, shelfColTarget, numRows, 1).setValues(shelfOutput);
    }
  }
  // ─────────────────────────────────────────────────────────────────

  // ── ИЗМЕНЕНИЕ 5: новые строки — вставляем shelf в нужную колонку ──
  const newRowsData = [];
  inMap.forEach((data, sku) => {
    if (!existingSkus.has(sku)) {
      const newRow          = new Array(lastCol).fill('');
      newRow[0]             = sku.toUpperCase(); // ITEM        (column A)
      newRow[1]             = data.name;          // DESCRIPTION (column B)
      newRow[targetCol - 1] = data.qty;
      if (shelfColTarget > 0 && data.shelf !== '') {
        newRow[shelfColTarget - 1] = data.shelf;
      }
      newRowsData.push(newRow);
      totalApplied++; totalNewRows++;
    }
  });
  // ─────────────────────────────────────────────────────────────────

  if (newRowsData.length > 0) {
    const startNewRow = sheet.getLastRow() + 1;
    sheet.getRange(startNewRow, 1, newRowsData.length, lastCol).setValues(newRowsData);
    sheet.getRange(startNewRow - 1, 1, 1, lastCol).copyTo(
      sheet.getRange(startNewRow, 1, newRowsData.length, lastCol),
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false
    );
  }

  updateDayBlock(targetSheetName);
  ss.toast(
    `Transfer to ${targetSheetName} complete! ` +
    `Updated: ${totalApplied - totalNewRows}, new rows added: ${totalNewRows}`
  );
}


/***********************************************************************
 * ОСНОВНЫЕ ФУНКЦИИ — Shelf_Transfer (Transfer OUT)
 ***********************************************************************/
function updateOrderShelfByStorage_ST(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shStorage = ss.getSheetByName(ST_STORAGE_SHEET);
  const shOrder   = ss.getSheetByName(ST_ORDER_SHEET);
  if (!shStorage || !shOrder) throw new Error('STORAGE or ' + ST_ORDER_SHEET + ' sheet not found');
  st_ensureThreeRowHeader(shStorage);

  const shelfPos = st_findHeaderPosition(shStorage, ST_SHELF_HEADER, ST_HEADER_SCAN_MAX_ROWS);
  if (shelfPos.col < 1) throw new Error(`На STORAGE нет "${ST_SHELF_HEADER}"`);
  const shelfColStorage = shelfPos.col;

  const lastRowStorage = shStorage.getLastRow();
  if (lastRowStorage < ST_DATA_START_ROW_STORAGE)
    throw new Error(`No data found on STORAGE below row ${ST_DATA_START_ROW_STORAGE}.`);

  const numRowsStorage = lastRowStorage - ST_DATA_START_ROW_STORAGE + 1;
  const storageVals = shStorage.getRange(ST_DATA_START_ROW_STORAGE, 1, numRowsStorage, shelfColStorage).getValues();
  const shelfMap = new Map();
  for (let i=0;i<numRowsStorage;i++){
    const key = st_normKey(storageVals[i][0]);
    const val = storageVals[i][shelfColStorage-1];
    if (key) shelfMap.set(key,val);
  }

  const skuPos = st_findHeaderPosition(shOrder, ST_ORDER_SKU_HEADERS);
  if (skuPos.col < 1) throw new Error(`На ${ST_ORDER_SHEET} нет колонки артикулов`);
  const headerRow    = skuPos.row;
  const dataStartRow = headerRow+1;

  let statusCol = st_findHeaderInRow(shOrder, ST_STATUS_HEADER, headerRow);
  if (statusCol<1) throw new Error(`Header "${ST_STATUS_HEADER}" not found in row ${headerRow}`);
  const existing = st_findHeaderPosition(shOrder, ST_SHELF_HEADER);
  if (existing.col > 0) {
    shOrder.deleteColumn(existing.col);
    statusCol = st_findHeaderInRow(shOrder, ST_STATUS_HEADER, headerRow);
  }

  shOrder.insertColumnBefore(statusCol);
  const shelfColOrder = statusCol;
  shOrder.getRange(headerRow, shelfColOrder).setValue(ST_SHELF_HEADER);

  const lastRowOrder = shOrder.getLastRow();
  if (lastRowOrder >= dataStartRow){
    const numRows = lastRowOrder - dataStartRow + 1;
    const skuVals = shOrder.getRange(dataStartRow, skuPos.col, numRows, 1).getValues();
    const out = Array.from({length:numRows},()=>['']);
    for (let i=0;i<numRows;i++){
      const k = st_normKey(skuVals[i][0]);
      if (k && shelfMap.has(k)) out[i][0] = shelfMap.get(k);
    }
    shOrder.getRange(dataStartRow, shelfColOrder, numRows, 1).setValues(out);
    shOrder.autoResizeColumn(shelfColOrder);
  }
}

function sortOrderByShelf_ST(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shOrder = ss.getSheetByName(ST_ORDER_SHEET);
  if (!shOrder) throw new Error(ST_ORDER_SHEET + ' sheet not found');

  const skuPos = st_findHeaderPosition(shOrder, ST_ORDER_SKU_HEADERS);
  if (skuPos.col < 1) throw new Error('SKU column not found');
  const headerRow = skuPos.row;
  const dataStartRow = headerRow+1;

  const shelfPos = st_findHeaderPosition(shOrder, ST_SHELF_HEADER);
  if (shelfPos.col < 1) throw new Error(`"${ST_SHELF_HEADER}" column not found`);

  const lastRow = shOrder.getLastRow();
  const lastCol = shOrder.getLastColumn();
  let sortStart = st_getSafeSortStartRow(shOrder, dataStartRow);
  const anchorsSet = new Set(st_findAnchoredRows(shOrder, dataStartRow));

  let segStart = sortStart;
  for (let r=sortStart; r<=lastRow+1; r++){
    const isAnchor = anchorsSet.has(r);
    const isEnd = (r === lastRow+1);
    if (isAnchor || isEnd){
      const segEnd = r-1;
      if (segEnd >= segStart){
        const segRows = segEnd - segStart + 1;
        if (segRows > 1){
          shOrder.getRange(segStart,1,segRows,lastCol).sort([{column:shelfPos.col,ascending:true}]);
        }
      }
      segStart = r+1;
    }
  }
}

function _transferToSheet_lastBlock(sheet, pairs, dataStartRow, sheetName){
  st_ensureThreeRowHeader(sheet);

  const lastCol = sheet.getLastColumn();
  if (lastCol < 2) throw new Error(`На ${sheetName} слишком мало колонок`);
  const hdr = sheet.getRange(ST_BLOCK_HEADER_ROW, 1, 1, lastCol).getDisplayValues()[0].map(v => String(v||'').trim());
  let lastTotalCol = -1;
  for (let c=lastCol;c>=1;c--){
    if (st_clean(hdr[c-1]) === 'total'){ lastTotalCol=c; break; }
  }
  if (lastTotalCol === -1) throw new Error(`На ${sheetName} нет Total (в строке ${ST_BLOCK_HEADER_ROW})`);

  const lastBlockType = st_clean(hdr[lastTotalCol - 2] || '');
  if (lastBlockType !== 'out') {
    SpreadsheetApp.getUi().alert(
      `Transfer error — "${sheetName}":\n` +
      `Last block is "${lastBlockType.toUpperCase()}", expected "OUT".\n\n` +
      `Create a new OUT block on "${sheetName}" (use the menu in A1).`
    );
    return 0;
  }

  const targetCol = lastTotalCol - 1;
  const lastRow = sheet.getLastRow();
  const count = Math.max(0,lastRow-dataStartRow+1);
  if (count > 0){
    const vals = sheet.getRange(dataStartRow,targetCol,count,1).getValues();
    const notEmpty = vals.some(r=>String(r[0]).trim()!=='');
    if (notEmpty){
      SpreadsheetApp.getUi().alert(`The last block in ${sheetName} is not empty. Please clear IN/OUT before transfer.`);
      return 0;
    }
  }

  const rowsRange = (count>0) ? sheet.getRange(dataStartRow,1,count,1).getDisplayValues() : [];
  const rowMap = new Map();
  for (let i=0;i<rowsRange.length;i++){
    const key = st_normKey(rowsRange[i][0]);
    if (key) rowMap.set(key,dataStartRow+i);
  }

  const out = Array.from({length:count},()=>['']);
  let written = 0;
  for (const {sku,val} of pairs){
    const row = rowMap.get(sku);
    if (!row) continue;
    out[row-dataStartRow][0] = val;
    written++;
  }
  if (count>0){
    sheet.getRange(dataStartRow,targetCol,count,1).setValues(out);
    sheet.getRange(dataStartRow,targetCol,count,1).setNumberFormat('0.############');
  }
  return written;
}

function transferToStorageLastBlock_ST() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const shOut = ss.getSheetByName(ST_ORDER_SHEET);
  if (!shOut) { ui.alert(`Error: Sheet "${ST_ORDER_SHEET}" not found.`); return; }

  const lastRowOut = shOut.getLastRow();
  if (lastRowOut < 2) { ss.toast(`No data found on sheet ${ST_ORDER_SHEET}.`); return; }

  // ── Находим колонки листа OUT по заголовкам строки 1 ─────────────
  const outLastCol          = shOut.getLastColumn();
  const outHeaders          = shOut.getRange(1, 1, 1, outLastCol).getValues()[0];
  const outColItem          = outHeaders.findIndex(h => String(h).trim().toLowerCase() === OUT_HEADER_ITEM.toLowerCase());
  const outColDesc          = outHeaders.findIndex(h => String(h).trim().toLowerCase() === OUT_HEADER_DESC.toLowerCase());
  const outColQty           = outHeaders.findIndex(h => String(h).trim().toLowerCase() === OUT_HEADER_QTY.toLowerCase());
  const outColStatusPottery = outHeaders.findIndex(h => String(h).trim() === OUT_HEADER_STATUS_POTTERY);
  const outColStatusPaints  = outHeaders.findIndex(h => String(h).trim() === OUT_HEADER_STATUS_PAINTS);

  // Обязательные колонки
  if (outColItem < 0 || outColDesc < 0 || outColQty < 0) {
    const missing = [
      outColItem < 0 ? `"${OUT_HEADER_ITEM}"` : null,
      outColDesc < 0 ? `"${OUT_HEADER_DESC}"` : null,
      outColQty  < 0 ? `"${OUT_HEADER_QTY}"` : null
    ].filter(Boolean).join(', ');
    ui.alert(`Sheet "${ST_ORDER_SHEET}": required headers ${missing} not found in row 1.`);
    return;
  }

  // Хотя бы одна статусная колонка должна существовать
  if (outColStatusPottery < 0 && outColStatusPaints < 0) {
    ui.alert(
      `Sheet "${ST_ORDER_SHEET}": columns\n` +
      `"${OUT_HEADER_STATUS_POTTERY}" and "${OUT_HEADER_STATUS_PAINTS}" not found.\n` +
      `Check the headers in row 1.`
    );
    return;
  }

  // ── Спрашиваем что переносим ──────────────────────────────────────
  const promptLines = [];
  if (outColStatusPottery >= 0) promptLines.push(`1 — ${OUT_HEADER_STATUS_POTTERY}`);
  if (outColStatusPaints  >= 0) promptLines.push(`2 — ${OUT_HEADER_STATUS_PAINTS}`);

  const response = ui.prompt(
    'Transfer OUT — what to transfer?',
    promptLines.join('\n'),
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) { ss.toast('Transfer cancelled.'); return; }

  const choice = response.getResponseText().trim();
  let selectedStatusCol, targetSheetName, selectedLabel;

  if (choice === '1') {
    if (outColStatusPottery < 0) { ui.alert(`Column "${OUT_HEADER_STATUS_POTTERY}" not found in row 1.`); return; }
    selectedStatusCol = outColStatusPottery;
    targetSheetName   = ST_STORAGE_SHEET;        // POTTERY
    selectedLabel     = OUT_HEADER_STATUS_POTTERY;
  } else if (choice === '2') {
    if (outColStatusPaints < 0) { ui.alert(`Column "${OUT_HEADER_STATUS_PAINTS}" not found in row 1.`); return; }
    selectedStatusCol = outColStatusPaints;
    targetSheetName   = ST_COLORS_SHEET;         // PAINTS
    selectedLabel     = OUT_HEADER_STATUS_PAINTS;
  } else {
    ui.alert('Invalid input. Enter 1 or 2 and click OK.');
    return;
  }

  // ── Считываем данные OUT, фильтруем по выбранной статусной колонке ─
  const outData = shOut.getRange(2, 1, lastRowOut - 1, outLastCol).getValues();
  const outMap  = new Map();
  outData.forEach(row => {
    const sku    = String(row[outColItem]         || '').trim().toLowerCase();
    const status = String(row[selectedStatusCol]  || '').trim();
    if (!sku || !status) return;
    outMap.set(sku, {
      name: row[outColDesc],
      qty:  row[outColQty]
    });
  });

  if (outMap.size === 0) {
    ui.alert(`Nothing to transfer.\nAll rows in column "${selectedLabel}" are empty.`);
    return;
  }

  // ── Валидируем целевой лист ───────────────────────────────────────
  const check = _checkSheetForTransfer(targetSheetName, 'OUT');

  if (check.status === 'critical') {
    ui.alert(`Transfer OUT failed:\n\n${check.message}`);
    return;
  }
  if (check.status === 'skip') {
    ui.alert(check.message);
    return;
  }

  // ── Выполняем перенос ─────────────────────────────────────────────
  const targetCol     = check.targetCol;
  const sheet         = ss.getSheetByName(targetSheetName);
  const lastCol       = sheet.getLastColumn();
  const lastRowTarget = sheet.getLastRow();
  const numRows       = lastRowTarget - DATA_START_ROW + 1;

  // SKU in column A (ITEM)
  const skuTargetVals = numRows > 0
    ? sheet.getRange(DATA_START_ROW, 1, numRows, 1).getValues()
    : [];
  const existingSkus = new Set(skuTargetVals.map(r => String(r[0] || '').trim().toLowerCase()));

  // Update existing items
  let totalApplied = 0, totalNewRows = 0;
  if (numRows > 0) {
    const output = [];
    for (let i = 0; i < numRows; i++) {
      const sku = String(skuTargetVals[i][0] || '').trim().toLowerCase();
      output.push([outMap.has(sku) ? outMap.get(sku).qty : '']);
      if (outMap.has(sku)) totalApplied++;
    }
    sheet.getRange(DATA_START_ROW, targetCol, numRows, 1).setValues(output);
  }

  // Append missing items to the bottom of the list
  const newRowsData = [];
  outMap.forEach((data, sku) => {
    if (!existingSkus.has(sku)) {
      const newRow          = new Array(lastCol).fill('');
      newRow[0]             = sku.toUpperCase();
      newRow[1]             = data.name;
      newRow[targetCol - 1] = data.qty;
      newRowsData.push(newRow);
      totalApplied++; totalNewRows++;
    }
  });

  if (newRowsData.length > 0) {
    const startNewRow = sheet.getLastRow() + 1;
    sheet.getRange(startNewRow, 1, newRowsData.length, lastCol).setValues(newRowsData);
    sheet.getRange(startNewRow - 1, 1, 1, lastCol).copyTo(
      sheet.getRange(startNewRow, 1, newRowsData.length, lastCol),
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false
    );
  }

  updateDayBlock(targetSheetName);
  ss.toast(
    `Transfer to ${targetSheetName} complete! ` +
    `Updated: ${totalApplied - totalNewRows}, new rows added: ${totalNewRows}`
  );
}


/***********************************************************************
 * ЕДИНОЕ МЕНЮ И ТРИГГЕРЫ
 ***********************************************************************/
function onOpen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Transfer IN')
    .addItem('📥 Transfer from IN to POTTERY/PAINTS', 'transferFromInSheet_ST')
    .addToUi();

  ui.createMenu('Transfer OUT')
    .addItem('Shelf #', 'updateOrderShelfByStorage_ST')
    .addItem('Sort by Shelf #', 'sortOrderByShelf_ST')
    .addItem('Transfer from OUT to POTTERY/PAINTS', 'transferToStorageLastBlock_ST')
    .addToUi();

  VALID_SHEETS.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    try { ensureThreeRowHeader(sh); } catch(e){}
    ensureActionDropdown(sh);
    try { normalizeHeaderRow3Alignment(name); } catch(e){}
    try { relaxValidationForInputColumns(sh); } catch(e){}
    try { applyBlockFontWeights(sh); } catch(e){}
  });
}

function onEdit(e) {
  if (!e || !e.range || !e.range.getSheet) return;
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  if (!VALID_SHEETS.includes(sheetName)) return;
  const a1 = e.range.getA1Notation();

  if (a1 === "A1") {
    const val = String(e.value || "");
    if (!val) return;
    if (val === "➕ IN")          addNewDayBlock("IN", sheetName);
    else if (val === "➕ OUT")    addNewDayBlock("OUT", sheetName);
    else if (val === "🔁 Update block") updateDayBlock(sheetName);
    else if (val === "🔤 Sort A→Z" && sheetName === 'POTTERY') sortByDescriptionPottery();
    sheet.getRange("A1").setValue("");
    return;
  }

  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row < DATA_START_ROW) return;
  const head = _norm(String(sheet.getRange(HDR_ROW_3, col).getValue() || ""));
  if (!head.startsWith("in") && !head.startsWith("out")) return;
  const hasFormula = !!e.range.getFormula();
  if (hasFormula) { e.range.setNumberFormat("0.############"); return; }

  const s = String(e.value || "").trim();
  if (s === "") { e.range.setNumberFormat("0.############"); return; }

  const normalized = s.replace(/\s+/g, "").replace(",", ".");
  const num = Number(normalized);
  if (!Number.isFinite(num)) {
    e.range.setValue("");
    try { SpreadsheetApp.getActive().toast("Please enter a number (use . or , as decimal separator).", "Invalid input", 3); } catch(err){}
    return;
  }
  e.range.setValue(num);
  e.range.setNumberFormat("0.############");
}