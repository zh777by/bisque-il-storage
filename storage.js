/***** НАСТРОЙКИ *****/
const VALID_SHEETS = ["PAINTS", "POTTERY", "SUPPLIES"];

// Было 2 строки шапки (дата / IN-OUT-Total), стало 3 строки — как в первом скрипте:
// HDR_ROW_1 — дата (merged)
// HDR_ROW_2 — пустая merged-строка (разделитель)
// HDR_ROW_3 — IN/OUT/Total, STORAGE/FLOOR/TOTAL ALL и т.д.
const HDR_ROW_1 = 1;
const HDR_ROW_2 = 2;
const HDR_ROW_3 = 3;
const DATA_START_ROW = 4; // было 3

const HEADER_COLOR_HEX = ["#b6d7a8", "#f4cccc", "#b4a7d6", "#5b95f9"]; // Тот самый цвет шапок

// Transfer-настройки
const IN_SHEET_NAME = 'IN';
const TARGET_SHEETS = ['POTTERY', 'PAINTS'];
const DATA_START_ROW_TARGET = DATA_START_ROW; // было отдельной константой = 3, теперь синхронизировано

// Заголовки статусных колонок листа IN (строка 1) — менять здесь при переименовании.
// Поиск идёт по тексту заголовка, а не по позиции колонки —
// можно свободно добавлять/удалять/переставлять колонки на листе IN.
const IN_HEADER_STATUS_POTTERY = 'סטטוס pottery (pcs)';
const IN_HEADER_STATUS_PAINTS  = 'סטטוס paints (pcs)';
const IN_HEADER_STATUS_SUPPLY  = 'סטטוס supply (pcs)'; // зарезервировано на будущее, перенос пока не используется

// Листы, на которых контролируется отрицательный остаток (TOTAL NOW < 0)
const NEGATIVE_STOCK_CHECK_SHEETS = ['POTTERY', 'PAINTS'];

/***** УТИЛИТЫ *****/
function columnToLetter(column) {
  let letter = "";
  while (column > 0) {
    let temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function _norm(v) { return String(v || "").toLowerCase().trim(); }

function getTotalNowCol(sheet) {
  const header1 = sheet.getRange(HDR_ROW_1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idx = header1.findIndex(h => String(h).trim() === "TOTAL NOW");
  if (idx === -1) throw new Error(`Колонка "TOTAL NOW" не найдена`);
  return idx + 1;
}

/** АВТО-МИГРАЦИЯ СО СТАРОЙ 2-СТРОЧНОЙ ШАПКИ НА НОВУЮ 3-СТРОЧНУЮ **/
// Если на листе всё ещё старая шапка (метки IN/OUT/Total/STORAGE/FLOOR/TOTAL ALL
// стоят в строке 2, а не в строке 3) — вставляем пустую строку после первой,
// сдвигая всё на одну строку вниз.
function ensureThreeRowHeader(sheet) {
  const lastCol = sheet.getLastColumn();
  if (!lastCol) return;
  const r2 = sheet.getRange(2, 1, 1, lastCol).getValues()[0].map(_norm);
  const r3 = sheet.getRange(3, 1, 1, lastCol).getValues()[0].map(_norm);
  const isLabel = v => v === "total" || v === "total all" || v === "storage" || v === "floor" ||
                        v.startsWith("in") || v.startsWith("out");
  const r2Has = r2.some(isLabel);
  const r3Has = r3.some(isLabel);
  if (r2Has && !r3Has) sheet.insertRowAfter(1);
}

function normalizeHeaderRow3Alignment(sheetName) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sh) return;
  const lastCol = sh.getLastColumn();
  if (!lastCol) return;
  sh.getRange(HDR_ROW_3, 1, 1, lastCol).setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  sh.getRange(HDR_ROW_2, 1, 1, lastCol).setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  try { sh.autoResizeRows(HDR_ROW_2, 2); } catch (e) {}
}

/** СТИРАНИЕ НУЛЕЙ В СИНИХ ШАПКАХ **/
// Логика разделителей применяется ТОЛЬКО к листу SUPPLIES — на PAINTS и POTTERY
// эта функция ничего не делает и данные строк не трогает.
// Строка на SUPPLIES считается "служебным разделителем" (и поэтому стирается),
// только если у неё И подходящий цвет фона в колонке A, И одновременно ПУСТО
// значение в колонке TOTAL NOW.
function _clearGroupHeaderRows(sheet, sheetName) {
  if (sheetName !== 'SUPPLIES') return; // разделители существуют только на SUPPLIES

  const totalRows = sheet.getLastRow();
  if (totalRows < DATA_START_ROW) return;

  const totalNowCol = getTotalNowCol(sheet);
  const totalCols   = sheet.getLastColumn();
  if (totalCols <= totalNowCol) return;

  const rowsCount    = totalRows - DATA_START_ROW + 1;
  const numDataCols  = totalCols - totalNowCol + 1;

  const bgs          = sheet.getRange(DATA_START_ROW, 1, rowsCount, 1).getBackgrounds();
  const totalNowVals = sheet.getRange(DATA_START_ROW, totalNowCol, rowsCount, 1).getValues();
  let bStart = null, bLen = 0;

  const flush = () => {
    if (bStart !== null && bLen > 0) {
      sheet.getRange(bStart, totalNowCol, bLen, numDataCols).clearContent();
    }
    bStart = null; bLen = 0;
  };

  for (let i = 0; i < rowsCount; i++) {
    const bg = (bgs[i][0] || "").toLowerCase();
    const totalNowEmpty = String(totalNowVals[i][0] || "").trim() === "";
    const isHeader = HEADER_COLOR_HEX.includes(bg) && totalNowEmpty; // <-- проверка по TOTAL NOW
    const r = DATA_START_ROW + i;
    if (isHeader) {
      if (bStart === null) { bStart = r; bLen = 1; }
      else if (r === bStart + bLen) { bLen++; }
      else { flush(); bStart = r; bLen = 1; }
    } else {
      flush();
    }
  }
  flush();
}

function _batchSetR1C1(sheet, startRow, endRow, col, r1c1) {
  const nRows = endRow - startRow + 1;
  if (nRows <= 0) return;
  sheet.getRange(startRow, col, nRows, 1).setFormulasR1C1(Array.from({ length: nRows }, () => [r1c1]));
}

function buildTotalNowFormulaA1(row, firstBlockCol, lastCol, sheetName) {
  const fL = columnToLetter(firstBlockCol);
  const lL = columnToLetter(lastCol);
  return `=IFERROR(INDEX(${row}:${row}, MAX(FILTER(COLUMN(${fL}$${HDR_ROW_3}:${lL}$${HDR_ROW_3}), ${fL}$${HDR_ROW_3}:${lL}$${HDR_ROW_3}="Total"))), "")`;
}

function updateDayBlock(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  ensureThreeRowHeader(sheet);
  normalizeHeaderRow3Alignment(sheetName);

  const totalRows = sheet.getLastRow();
  if (totalRows < DATA_START_ROW) return;

  const rowsCount   = totalRows - DATA_START_ROW + 1;
  const totalNowCol = getTotalNowCol(sheet);
  const lastCol     = sheet.getLastColumn();
  const firstBlockCol = totalNowCol + 1;

  ensureFormulasInAllTotals(sheet, sheetName);

  const formulasNow = [];
  for (let r = DATA_START_ROW; r <= totalRows; r++) {
    formulasNow.push([buildTotalNowFormulaA1(r, firstBlockCol, lastCol, sheetName)]);
  }
  const totalNowRange = sheet.getRange(DATA_START_ROW, totalNowCol, rowsCount, 1);
  totalNowRange.setFormulas(formulasNow);
  totalNowRange.setNumberFormat("0.############");

  _clearGroupHeaderRows(sheet, sheetName);
  SpreadsheetApp.flush();

  // Проверка на отрицательный остаток — выполняется при каждом обновлении блока
  checkNegativeStock();
}

function ensureFormulasInAllTotals(sheet, sheetName) {
  const totalRows = sheet.getLastRow();
  const totalCols = sheet.getLastColumn();
  const hRow = sheet.getRange(HDR_ROW_3, 1, 1, totalCols).getValues()[0].map(_norm);
  const totalNowCol = getTotalNowCol(sheet);

  for (let j = totalNowCol + 1; j <= totalCols; j++) {
    if (hRow[j-1] === "total all") {
      _batchSetR1C1(sheet, DATA_START_ROW, totalRows, j, `=IFERROR(VALUE(TRIM(R[0]C[-2])),0)+IFERROR(VALUE(TRIM(R[0]C[-1])),0)`);
      continue;
    }
    if (hRow[j-1] !== "total") continue;

    const is4col = j >= 4 && hRow[j-4].startsWith("in") && hRow[j-3].startsWith("out") && hRow[j-2].startsWith("out to floor");
    const is3col = !is4col && j >= 3 && hRow[j-3].startsWith("in") && hRow[j-2].startsWith("out to floor");

    let r1c1;
    if (is4col) r1c1 = `=IFERROR(VALUE(TRIM(R[0]C[-4])),0)+IFERROR(VALUE(TRIM(R[0]C[-3])),0)-IFERROR(VALUE(TRIM(R[0]C[-2])),0)-IFERROR(VALUE(TRIM(R[0]C[-1])),0)`;
    else if (is3col) r1c1 = `=IFERROR(VALUE(TRIM(R[0]C[-3])),0)+IFERROR(VALUE(TRIM(R[0]C[-2])),0)-IFERROR(VALUE(TRIM(R[0]C[-1])),0)`;
    else {
      const sign = (hRow[j-2] && hRow[j-2].startsWith("in")) ? "+" : "-";
      const isFirst = (j - 2) <= totalNowCol;
      if (isFirst) r1c1 = `=IF(R[0]C[-1]="",0,${sign === '+' ? '' : sign}IFERROR(VALUE(TRIM(R[0]C[-1])),0))`;
      else {
        const isAfterInv = hRow[j-3] === "total all" && (sheetName === "PAINTS" || sheetName === "POTTERY");
        const prev = isAfterInv ? "R[0]C[-4]" : "R[0]C[-2]";
        r1c1 = `=IF(R[0]C[-1]="",IFERROR(VALUE(TRIM(${prev})),0),IFERROR(VALUE(TRIM(${prev})),0)${sign}IFERROR(VALUE(TRIM(R[0]C[-1])),0))`;
      }
    }
    _batchSetR1C1(sheet, DATA_START_ROW, totalRows, j, r1c1);
    sheet.getRange(DATA_START_ROW, j, totalRows - DATA_START_ROW + 1, 1).setNumberFormat("0.############");
  }
}

function addNewDayBlock(mode, sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  ensureThreeRowHeader(sheet);
  normalizeHeaderRow3Alignment(sheetName);

  const lastColBefore = sheet.getLastColumn();
  const maxRows = sheet.getMaxRows();
  const date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

  sheet.insertColumnsAfter(lastColBefore, 2);
  const label = mode.toLowerCase().includes("in") ? "IN" : mode.toLowerCase().includes("floor") ? "out to floor" : "OUT";

  sheet.getRange(HDR_ROW_1, lastColBefore + 1, 1, 2).merge().setValue(date)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.getRange(HDR_ROW_2, lastColBefore + 1, 1, 2).merge().setValue("")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.getRange(HDR_ROW_3, lastColBefore + 1).setValue(label);
  sheet.getRange(HDR_ROW_3, lastColBefore + 2).setValue("Total");
  if (label === "IN") sheet.getRange(HDR_ROW_3, lastColBefore + 1).setBackground("#ffeb3b");

  sheet.getRange(HDR_ROW_1, lastColBefore + 1, maxRows, 1).setBorder(null, true, null, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);

  updateDayBlock(sheetName);
}

function addNewInventoryBlock(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  ensureThreeRowHeader(sheet);
  normalizeHeaderRow3Alignment(sheetName);

  const lastColBefore = sheet.getLastColumn();
  const maxRows = sheet.getMaxRows();
  const date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

  sheet.insertColumnsAfter(lastColBefore, 3);
  sheet.getRange(HDR_ROW_1, lastColBefore + 1, 1, 3).merge().setValue("Inventory " + date)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.getRange(HDR_ROW_2, lastColBefore + 1, 1, 3).merge().setValue("")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.getRange(HDR_ROW_3, lastColBefore + 1).setValue("STORAGE");
  sheet.getRange(HDR_ROW_3, lastColBefore + 2).setValue("FLOOR");
  sheet.getRange(HDR_ROW_3, lastColBefore + 3).setValue("TOTAL ALL");

  sheet.getRange(HDR_ROW_1, lastColBefore + 1, maxRows, 1).setBorder(null, true, null, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);

  updateDayBlock(sheetName);
}

/***** ПРОВЕРКА ОТРИЦАТЕЛЬНОГО ОСТАТКА *****/
/**
 * Сканирует листы из NEGATIVE_STOCK_CHECK_SHEETS (POTTERY, PAINTS),
 * находит все строки, где TOTAL NOW < 0, и если такие есть —
 * показывает модальное окно со списком проблемных SKU.
 * Вызывается из updateDayBlock() при каждом пересчёте блока,
 * а также из onOpen() при открытии файла.
 */
function checkNegativeStock() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const problemBlocks = [];

  NEGATIVE_STOCK_CHECK_SHEETS.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < DATA_START_ROW) return;

    let totalNowCol;
    try {
      totalNowCol = getTotalNowCol(sheet);
    } catch (e) {
      return; // на этом листе нет колонки TOTAL NOW
    }

    const rowsCount = lastRow - DATA_START_ROW + 1;
    const totalNowVals = sheet.getRange(DATA_START_ROW, totalNowCol, rowsCount, 1).getValues();
    const skuVals = sheet.getRange(DATA_START_ROW, 2, rowsCount, 1).getValues(); // колонка B = SKU

    const negatives = [];
    for (let i = 0; i < rowsCount; i++) {
      const val = totalNowVals[i][0];
      if (typeof val === "number" && val < 0) {
        const sku = String(skuVals[i][0] || "").trim() || "(no SKU)";
        const row = DATA_START_ROW + i;
        negatives.push(`  • ${sku} — row ${row}: ${val}`);
      }
    }

    if (negatives.length > 0) {
      problemBlocks.push(`Sheet "${sheetName}":\n` + negatives.join("\n"));
    }
  });

  if (problemBlocks.length > 0) {
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      "⚠️ Negative stock detected!",
      "Stock quantity cannot be less than zero. Please resolve the following:\n\n" +
        problemBlocks.join("\n\n"),
      ui.ButtonSet.OK
    );
  }
}

/***** UI *****/
function onOpen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  SpreadsheetApp.getUi()
    .createMenu('Transfer')
    .addItem('📥 Transfer from IN to POTTERY/PAINTS', 'transferFromInSheet_ST')
    .addToUi();

  VALID_SHEETS.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    try { ensureThreeRowHeader(sh); } catch (e) {}
    try { normalizeHeaderRow3Alignment(name); } catch (e) {}

    const actions = name === 'POTTERY'
      ? ["➕ IN", "➕ OUT", "➕ Out to floor", "📦 INVENTORY", "🔁 Update block", "🔤 Sort A→Z"]
      : ["➕ IN", "➕ OUT", "➕ Out to floor", "📦 INVENTORY", "🔁 Update block"];
    sh.getRange("A1").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(actions, true).build());
  });

  // Показать предупреждение и при открытии файла, если проблема ещё не решена
  checkNegativeStock();
}

function onEdit(e) {
  if (!e || !e.range || e.range.getA1Notation() !== "A1") return;
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  if (!VALID_SHEETS.includes(sheetName)) return;

  const val = String(e.value || "");
  if (!val) return;

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) return;

  try {
    const lowerVal = val.toLowerCase();
    if      (lowerVal.includes("in") && !lowerVal.includes("inventory")) addNewDayBlock("IN", sheetName);
    else if (lowerVal.includes("out") && !lowerVal.includes("floor"))    addNewDayBlock("OUT", sheetName);
    else if (lowerVal.includes("floor"))     addNewDayBlock("out to floor", sheetName);
    else if (lowerVal.includes("inventory")) addNewInventoryBlock(sheetName);
    else if (lowerVal.includes("update"))    updateDayBlock(sheetName);
    else if (lowerVal.includes("sort")) sortByColumnA(sheetName);

    sheet.getRange("A1").setValue("");
  } finally { lock.releaseLock(); }
}

function sortByColumnA(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < DATA_START_ROW) return;

  const rangeToSort = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, lastCol);
  rangeToSort.sort({ column: 1, ascending: true });

  SpreadsheetApp.getActiveSpreadsheet().toast(`Sheet ${sheetName} sorted A → Z by column A.`);
}

/***** TRANSFER *****/
function transferFromInSheet_ST() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const shIn = ss.getSheetByName(IN_SHEET_NAME);

  if (!shIn) {
    ui.alert("Error: 'IN' sheet not found.");
    return;
  }

  const lastRowIn = shIn.getLastRow();
  if (lastRowIn < 2) {
    ss.toast("No data found on 'IN' sheet.");
    return;
  }

  // --- 1. Находим статусные колонки на листе IN по названию заголовка ---
  // Поиск идёт по тексту в строке 1, а не по фиксированной позиции —
  // колонки можно свободно добавлять, удалять и переставлять.
  const inLastCol = shIn.getLastColumn();
  const inHeaders = shIn.getRange(1, 1, 1, inLastCol).getValues()[0];

  const colStatusPottery = inHeaders.findIndex(h => String(h).trim() === IN_HEADER_STATUS_POTTERY);
  const colStatusPaints  = inHeaders.findIndex(h => String(h).trim() === IN_HEADER_STATUS_PAINTS);
  // const colStatusSupply = inHeaders.findIndex(h => String(h).trim() === IN_HEADER_STATUS_SUPPLY); // зарезервировано

  if (colStatusPottery < 0 && colStatusPaints < 0) {
    ui.alert(
      `Sheet "${IN_SHEET_NAME}": columns\n` +
      `"${IN_HEADER_STATUS_POTTERY}" and "${IN_HEADER_STATUS_PAINTS}" not found in row 1.\n` +
      `Check the headers.`
    );
    return;
  }

  // --- 2. Спрашиваем что переносим (показываем только найденные варианты) ---
  const promptLines = [];
  if (colStatusPottery >= 0) promptLines.push(`1 — ${IN_HEADER_STATUS_POTTERY}`);
  if (colStatusPaints  >= 0) promptLines.push(`2 — ${IN_HEADER_STATUS_PAINTS}`);

  const response = ui.prompt(
    'Transfer',
    'What are we transferring?\n' + promptLines.join('\n'),
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const choice = response.getResponseText().trim();
  let targetSheetName, statusColIndex, statusLabel;

  if (choice === '1') {
    if (colStatusPottery < 0) { ui.alert(`Column "${IN_HEADER_STATUS_POTTERY}" not found in row 1.`); return; }
    targetSheetName = 'POTTERY';
    statusColIndex  = colStatusPottery;
    statusLabel     = IN_HEADER_STATUS_POTTERY;
  } else if (choice === '2') {
    if (colStatusPaints < 0) { ui.alert(`Column "${IN_HEADER_STATUS_PAINTS}" not found in row 1.`); return; }
    targetSheetName = 'PAINTS';
    statusColIndex  = colStatusPaints;
    statusLabel     = IN_HEADER_STATUS_PAINTS;
  } else {
    ui.alert('Invalid input. Please enter 1 or 2.');
    return;
  }

  // --- 3. Читаем данные с листа IN (вся ширина листа, не фиксированные 5 колонок) ---
  const inData = shIn.getRange(2, 1, lastRowIn - 1, inLastCol).getValues();

  const rowsToTransfer = inData
    .map(row => ({
      sku:  String(row[0] || '').trim(),
      desc: row[1],
      qty:  row[statusColIndex]
    }))
    .filter(r => r.sku !== '' && String(r.qty).trim() !== '');

  if (rowsToTransfer.length === 0) {
    ui.alert(`Nothing to transfer: all rows in the "${statusLabel}" column are empty.`);
    return;
  }

  // --- 4. Validate target sheet ---
  const sheet = ss.getSheetByName(targetSheetName);
  if (!sheet) {
    ui.alert(`Error: sheet '${targetSheetName}' not found.`);
    return;
  }
  ensureThreeRowHeader(sheet);

  const lastCol    = sheet.getLastColumn();
  const rowHeader  = sheet.getRange(HDR_ROW_3, 1, 1, lastCol).getValues()[0];

  let targetCol   = -1;
  let lastInLabel = '';
  for (let c = lastCol; c >= 1; c--) {
    const h = String(rowHeader[c - 1]).trim().toLowerCase();
    if (h === 'in' || h === 'out' || h === 'out to floor') {
      targetCol   = c;
      lastInLabel = h;
      break;
    }
  }

  if (targetCol === -1) {
    ui.alert(`No IN/OUT block found on sheet "${targetSheetName}".`);
    return;
  }

  // --- 5. Block validation ---
  if (lastInLabel !== 'in') {
    ui.alert(
      `The active block on sheet "${targetSheetName}" is "${lastInLabel.toUpperCase()}".\n` +
      `The active block must be IN before transferring.`
    );
    return;
  }

  const lastRowTarget = sheet.getLastRow();
  const numRows       = lastRowTarget - DATA_START_ROW_TARGET + 1;

  if (numRows > 0) {
    const blockVals = sheet.getRange(DATA_START_ROW_TARGET, targetCol, numRows, 1).getValues();
    const notEmpty  = blockVals.some(r => String(r[0]).trim() !== '');
    if (notEmpty) {
      ui.alert(
        `The IN block on sheet "${targetSheetName}" is not empty.\n` +
        `Please clear it manually before transferring.`
      );
      return;
    }
  }

  // --- 6. Perform transfer ---
  const skuToRow = new Map();
  if (numRows > 0) {
    const skuVals = sheet.getRange(DATA_START_ROW_TARGET, 2, numRows, 1).getValues();
    skuVals.forEach((r, i) => {
      const sku = String(r[0] || '').trim().toUpperCase();
      if (sku) skuToRow.set(sku, DATA_START_ROW_TARGET + i);
    });
  }

  const newRows = [];
  let totalApplied = 0;
  let totalNew     = 0;

  rowsToTransfer.forEach(({ sku, desc, qty }) => {
    const skuUpper = sku.toUpperCase();
    if (skuToRow.has(skuUpper)) {
      sheet.getRange(skuToRow.get(skuUpper), targetCol).setValue(qty);
      totalApplied++;
    } else {
      const newRow = new Array(lastCol).fill('');
      newRow[1]             = skuUpper;
      newRow[0]             = desc;
      newRow[targetCol - 1] = qty;
      newRows.push(newRow);
      totalNew++;
    }
  });

  if (newRows.length > 0) {
    const startNewRow = sheet.getLastRow() + 1;
    sheet.getRange(startNewRow, 1, newRows.length, lastCol).setValues(newRows);

    const fmtRow = sheet.getRange(startNewRow - 1, 1, 1, lastCol);
    fmtRow.copyTo(
      sheet.getRange(startNewRow, 1, newRows.length, lastCol),
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
      false
    );
  }

  updateDayBlock(targetSheetName);

  ss.toast(
    `Transfer to ${targetSheetName} complete! ` +
    `Matched: ${totalApplied}, new rows added: ${totalNew}`
  );
}
