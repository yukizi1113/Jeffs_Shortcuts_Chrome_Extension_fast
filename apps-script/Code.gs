/**
 * Jeff's Keyboard Shortcuts for Excel 2010 v3.1.19_extended_v.2_distribution
 * Google Sheets port (2026-08-25)
 *
 * This is a faithful/best-effort port of the shortcut routines found in the
 * supplied XLAM.  It is designed to be bound to ONE Google Spreadsheet.
 *
 * IMPORTANT differences from Excel are documented in jeffShowMigrationNotes().
 */

const JEFF = Object.freeze({
  // Current settings read from the supplied XLAM's Settings sheet.
  numberFormats: Object.freeze([
    '#,##0.0;(#,##0.0)',
    '$#,##0.0;($#,##0.0)',
    '#,##0.0%;(#,##0.0%)',
    '#,##0.0"x"',
    'General',
  ]),
  fontColors: Object.freeze(['#0000FF', '#00B050', '#FFFFFF', '#000000']),
  // null = no fill. Sheets cannot reliably distinguish explicit white fill
  // from a default/no-fill cell through the basic Spreadsheet service.
  fillColors: Object.freeze([null, '#FFFF00', '#CCFFFF', '#C0C0C0', '#002960']),
  specifiedFill: '#002960',
  specifiedFont: '#FFFFFF',
  defaultFont: '#000000',

  // XLAM hard-coded dimensions translated approximately to pixels.
  // Excel source assigns RowHeight=13.2, but DefaultRowHeight is declared As Long,
  // so the actual runtime value is coerced to 13pt. Width step=1, height step=2pt,
  // square/default-narrow ColumnWidth=2.11.
  defaultColumnWidthPx: 61,
  defaultRowHeightPx: 17,
  widthStepPx: 7,
  heightStepPx: 3,
  squareWidthPx: 20,

  datePattern: 'yyyyMMdd', // current XLAM setting: 8-digit date, at front, no time
  dateAtFront: true,

  undoPrefix: 'JEFF_SUPERFILL_UNDO_',
  borderPrefix: 'JEFF_BORDER_STATE_',
  dispatcherPrefixKey: 'JEFF_DISPATCH_PREFIX',
  dispatcherPrefixTimeKey: 'JEFF_DISPATCH_PREFIX_TIME',
  dispatcherTimeoutMs: 8000,
});

/** Adds a human-readable menu whenever the spreadsheet is opened. */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Jeff's Shortcuts")
    .addSubMenu(ui.createMenu('Formatting')
      .addItem('J11  Number format toggle', 'jeffNumberFormatToggle')
      .addItem('J12  Font color toggle', 'jeffFontColorToggle')
      .addItem('J13  Fill color toggle', 'jeffFillColorToggle')
      .addItem('J31  Center across (approx.)', 'jeffCenterAcrossSelection')
      .addItem('J32  Remove center across', 'jeffRemoveCenterAcrossSelection')
      .addItem('J33  Clear font color', 'jeffRemoveFontColor')
      .addItem('J34  Clear fill color', 'jeffRemoveFillColor')
      .addItem('J35  Wrap text toggle', 'jeffWrapTextToggle')
      .addItem('J38  Table top toggle (approx.)', 'jeffTableTop')
      .addItem('J39  Navy specified style toggle', 'jeffSpecifiedStyleToggle')
      .addItem('J41  Arial 10 + hide gridlines', 'jeffSetDefaultFont')
      .addItem('J42  Border toggle', 'jeffBorderToggle')
      .addItem('J43  Clear all borders', 'jeffRemoveBorder')
      .addItem('J54  Decrease decimals', 'jeffDecreaseDecimal')
      .addItem('J55  Increase decimals', 'jeffIncreaseDecimal'))
    .addSubMenu(ui.createMenu('Fill / formulas')
      .addItem('J18  SuperFill down to end', 'jeffSuperFillDownEnd')
      .addItem('J19  SuperFill down to first blank', 'jeffSuperFillDownBlank')
      .addItem('J21  Undo last SuperFill', 'jeffUndoSuperFill')
      .addItem('J24  SuperFill right to end', 'jeffSuperFillRightEnd')
      .addItem('J25  SuperFill right to first blank', 'jeffSuperFillRightBlank')
      .addItem('J26  Recalculate selection (approx.)', 'jeffCalcSelection')
      .addItem('J27  Convert references to absolute', 'jeffConvertToAbsolute')
      .addItem('J28  Convert references to relative', 'jeffConvertToRelative'))
    .addSubMenu(ui.createMenu('Size / alignment')
      .addItem('J44  Default column width', 'jeffSetDefaultWidth')
      .addItem('J45  Default row height', 'jeffSetDefaultHeight')
      .addItem('J46  Increase column width', 'jeffIncreaseWidth')
      .addItem('J47  Decrease column width', 'jeffDecreaseWidth')
      .addItem('J48  Increase row height', 'jeffIncreaseHeight')
      .addItem('J49  Decrease row height', 'jeffDecreaseHeight')
      .addItem('J51  Narrow/square column width', 'jeffDefaultSquareWidth')
      .addItem('J52  Vertical alignment toggle', 'jeffVerticalAlignmentToggle')
      .addItem('J53  Horizontal alignment toggle', 'jeffHorizontalAlignmentToggle'))
    .addSubMenu(ui.createMenu('File / diagnostics')
      .addItem('J17  Make dated copy', 'jeffSaveAsWithDate')
      .addItem('Migration notes', 'jeffShowMigrationNotes')
      .addItem('Authorization check', 'jeffAuthorize'))
    .addToUi();

  // Excel Customizations.exportedui 相当のメニュー
  ui.createMenu('Excel Toolbar')
    .addItem('参照元セルのトレース', 'excelTracePrecedents')
    .addItem('参照先セルのトレース', 'excelTraceDependents')
    .addItem('トレース解除', 'excelClearTrace')
    .addSeparator()
    .addItem('バックアップを保存', 'excelSaveBackUp')
    .addItem('添付メールの下書きを作成', 'excelSendAsAttachment')
    .addSeparator()
    .addItem('アウトラインを閉じる', 'excelCollapseOutline')
    .addItem('アウトラインを開く', 'excelExpandOutline')
    .addItem('重複を削除', 'excelRemoveDuplicates')
    .addSeparator()
    .addItem('セルの塗りつぶし', 'excelSetFillColor')
    .addSubMenu(
      ui.createMenu('What-If Analysis')
        .addItem('Goal Seek', 'excelGoalSeek')
    )
    .addToUi();
}

/** Run once from Apps Script editor to request all needed permissions. */
function jeffAuthorize() {
  const ss = SpreadsheetApp.getActive();
  SpreadsheetApp.flush();
  // Drive access is required only for the dated-copy command.
  DriveApp.getFileById(ss.getId()).getName();
  ss.toast("Jeff's Shortcuts: authorization is ready.", 'Jeff', 5);
}

// -----------------------------------------------------------------------------
// 9-key dispatcher used by AutoHotkey.
// Import jeffKey1 ... jeffKey9 as Google Sheets macros and assign
// Ctrl+Alt+Shift+1 ... Ctrl+Alt+Shift+9 respectively.
// AutoHotkey sends a two-digit command such as 11 by invoking Key1 twice.
// -----------------------------------------------------------------------------
function jeffKey1() { _jeffDispatcherDigit_(1); }
function jeffKey2() { _jeffDispatcherDigit_(2); }
function jeffKey3() { _jeffDispatcherDigit_(3); }
function jeffKey4() { _jeffDispatcherDigit_(4); }
function jeffKey5() { _jeffDispatcherDigit_(5); }
function jeffKey6() { _jeffDispatcherDigit_(6); }
function jeffKey7() { _jeffDispatcherDigit_(7); }
function jeffKey8() { _jeffDispatcherDigit_(8); }
function jeffKey9() { _jeffDispatcherDigit_(9); }

function _jeffDispatcherDigit_(digit) {
  const props = PropertiesService.getUserProperties();
  const now = Date.now();
  const prefix = props.getProperty(JEFF.dispatcherPrefixKey);
  const prefixTime = Number(props.getProperty(JEFF.dispatcherPrefixTimeKey) || 0);

  if (!prefix || now - prefixTime > JEFF.dispatcherTimeoutMs) {
    props.setProperty(JEFF.dispatcherPrefixKey, String(digit));
    props.setProperty(JEFF.dispatcherPrefixTimeKey, String(now));
    return;
  }

  props.deleteProperty(JEFF.dispatcherPrefixKey);
  props.deleteProperty(JEFF.dispatcherPrefixTimeKey);
  const code = prefix + String(digit);
  _jeffDispatch_(code);
}

function _jeffDispatch_(code) {
  const actions = {
    '11': jeffNumberFormatToggle,
    '12': jeffFontColorToggle,
    '13': jeffFillColorToggle,
    '17': jeffSaveAsWithDate,
    '18': jeffSuperFillDownEnd,
    '19': jeffSuperFillDownBlank,
    '21': jeffUndoSuperFill,
    '24': jeffSuperFillRightEnd,
    '25': jeffSuperFillRightBlank,
    '26': jeffCalcSelection,
    '27': jeffConvertToAbsolute,
    '28': jeffConvertToRelative,
    '31': jeffCenterAcrossSelection,
    '32': jeffRemoveCenterAcrossSelection,
    '33': jeffRemoveFontColor,
    '34': jeffRemoveFillColor,
    '35': jeffWrapTextToggle,
    '36': jeffIndentIncrease,
    '37': jeffIndentDecrease,
    '38': jeffTableTop,
    '39': jeffSpecifiedStyleToggle,
    '41': jeffSetDefaultFont,
    '42': jeffBorderToggle,
    '43': jeffRemoveBorder,
    '44': jeffSetDefaultWidth,
    '45': jeffSetDefaultHeight,
    '46': jeffIncreaseWidth,
    '47': jeffDecreaseWidth,
    '48': jeffIncreaseHeight,
    '49': jeffDecreaseHeight,
    '51': jeffDefaultSquareWidth,
    '52': jeffVerticalAlignmentToggle,
    '53': jeffHorizontalAlignmentToggle,
    '54': jeffDecreaseDecimal,
    '55': jeffIncreaseDecimal,
    '59': jeffShowMigrationNotes,
  };
  const fn = actions[code];
  if (!fn) {
    SpreadsheetApp.getActive().toast('Unknown Jeff command: ' + code, 'Jeff', 4);
    return;
  }
  fn();
}

// -----------------------------------------------------------------------------
// Formatting cycles
// -----------------------------------------------------------------------------
function jeffNumberFormatToggle() {
  const r = _jeffRange_();
  const current = r.getCell(1, 1).getNumberFormat();
  const next = _jeffNext_(current, JEFF.numberFormats, x => String(x).toLowerCase());
  r.setNumberFormat(next);
}

function jeffFontColorToggle() {
  const r = _jeffRange_();
  let current = r.getCell(1, 1).getFontColor();
  current = current ? current.toUpperCase() : '#000000';
  const next = _jeffNext_(current, JEFF.fontColors, x => String(x).toUpperCase());
  r.setFontColor(next);
}

function jeffFillColorToggle() {
  const r = _jeffRange_();
  let current = r.getCell(1, 1).getBackground();
  current = current ? current.toUpperCase() : null;
  // Basic Apps Script reports the default/no-fill background as white.
  // Treat white as no fill. This is the only intentional ambiguity here.
  if (current === '#FFFFFF') current = null;
  const next = _jeffNext_(current, JEFF.fillColors, x => x === null ? '__NONE__' : String(x).toUpperCase());
  r.setBackground(next);
}

function jeffRemoveFontColor() {
  _jeffRange_().setFontColor('#000000');
}

function jeffRemoveFillColor() {
  _jeffRange_().setBackground(null);
}

function jeffWrapTextToggle() {
  const r = _jeffRange_();
  const current = r.getCell(1, 1).getWrap();
  r.setWrap(!current);
}

function jeffSpecifiedStyleToggle() {
  const r = _jeffRange_();
  const c = r.getCell(1, 1);
  const isOn = c.getFontWeight() === 'bold' &&
    String(c.getFontColor() || '').toUpperCase() === JEFF.specifiedFont &&
    String(c.getBackground() || '').toUpperCase() === JEFF.specifiedFill;
  if (isOn) {
    r.setFontWeight('normal').setFontColor(JEFF.defaultFont).setBackground(null);
  } else {
    r.setFontWeight('bold').setFontColor(JEFF.specifiedFont).setBackground(JEFF.specifiedFill);
  }
}

function jeffSetDefaultFont() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const all = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
  all.setFontFamily('Arial').setFontSize(10);
  sheet.setHiddenGridlines(true);
}

// -----------------------------------------------------------------------------
// Borders
// -----------------------------------------------------------------------------
function jeffBorderToggle() {
  const r = _jeffRange_();
  const props = PropertiesService.getUserProperties();
  const key = JEFF.borderPrefix + SpreadsheetApp.getActive().getId() + '_' +
    r.getSheet().getSheetId() + '_' + r.getA1Notation();
  const previous = Number(props.getProperty(key) || 0);
  const next = previous >= 3 ? 0 : previous + 1;

  if (next === 0) {
    // The XLAM clears ALL borders when it cycles past its double border.
    r.setBorder(false, false, false, false, false, false);
  } else {
    const style = next === 1
      ? SpreadsheetApp.BorderStyle.SOLID
      : next === 2
        ? SpreadsheetApp.BorderStyle.DOTTED
        : SpreadsheetApp.BorderStyle.DOUBLE;
    // Match XLAM: bottom edge + inside horizontal borders only.
    r.setBorder(null, null, true, null, null, true, null, style);
  }
  props.setProperty(key, String(next));
}

function jeffRemoveBorder() {
  const r = _jeffRange_();
  r.setBorder(false, false, false, false, false, false);
  const key = JEFF.borderPrefix + SpreadsheetApp.getActive().getId() + '_' +
    r.getSheet().getSheetId() + '_' + r.getA1Notation();
  PropertiesService.getUserProperties().deleteProperty(key);
}

// -----------------------------------------------------------------------------
// Alignment / "table top"
// -----------------------------------------------------------------------------
function jeffVerticalAlignmentToggle() {
  const r = _jeffRange_();
  const current = r.getCell(1, 1).getVerticalAlignment();
  const next = current === 'middle' ? 'bottom' : current === 'bottom' ? 'top' : 'middle';
  r.setVerticalAlignment(next);
}

function jeffHorizontalAlignmentToggle() {
  const r = _jeffRange_();
  const current = r.getCell(1, 1).getHorizontalAlignment();
  const next = current === 'center' ? 'right' : current === 'right' ? 'left' : 'center';
  r.setHorizontalAlignment(next);
}

function jeffCenterAcrossSelection() {
  // Google Sheets has no xlCenterAcrossSelection equivalent. Do NOT merge cells,
  // because merging would alter the data model. This is the safest approximation.
  _jeffRange_().setHorizontalAlignment('center');
}

function jeffRemoveCenterAcrossSelection() {
  // The XLAM changes CenterAcrossSelection to ordinary center.
  _jeffRange_().setHorizontalAlignment('center');
}

function jeffTableTop() {
  const r = _jeffRange_();
  const c = r.getCell(1, 1);
  const isOn = c.getVerticalAlignment() === 'bottom' &&
    c.getFontLine() === 'underline' && c.getFontWeight() === 'bold';

  if (isOn) {
    r.setHorizontalAlignment(null)
      .setVerticalAlignment('middle')
      .setFontLine('none')
      .setFontWeight('normal')
      .setNumberFormat('General');
  } else {
    // XLAM also unmerges, turns wrapping off, uses CenterAcrossSelection,
    // and applies SINGLE ACCOUNTING underline. Sheets has no accounting-underline
    // or CenterAcrossSelection, so ordinary underline + center is used.
    try { r.breakApart(); } catch (e) { /* no merged cells or partial merge */ }
    r.setNumberFormat('@')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('bottom')
      .setWrap(false)
      .setFontLine('underline')
      .setFontWeight('bold');
  }
}

// Google Sheets has no cell-indent property in Apps Script or the current Range API.
// We intentionally do NOT fake indent by adding spaces, because that changes cell values.
function jeffIndentIncrease() {
  SpreadsheetApp.getActive().toast('Sheets has no non-destructive cell-indent API. No change made.', 'Jeff', 6);
}
function jeffIndentDecrease() {
  SpreadsheetApp.getActive().toast('Sheets has no non-destructive cell-indent API. No change made.', 'Jeff', 6);
}

// -----------------------------------------------------------------------------
// Row / column dimensions
// -----------------------------------------------------------------------------
function jeffSetDefaultWidth() { _jeffSetSelectedColumnWidths_(JEFF.defaultColumnWidthPx); }
function jeffSetDefaultHeight() { _jeffSetSelectedRowHeights_(JEFF.defaultRowHeightPx); }
function jeffDefaultSquareWidth() { _jeffSetSelectedColumnWidths_(JEFF.squareWidthPx); }
function jeffIncreaseWidth() { _jeffAdjustSelectedColumnWidths_(JEFF.widthStepPx); }
function jeffDecreaseWidth() { _jeffAdjustSelectedColumnWidths_(-JEFF.widthStepPx); }
function jeffIncreaseHeight() { _jeffAdjustSelectedRowHeights_(JEFF.heightStepPx); }
function jeffDecreaseHeight() { _jeffAdjustSelectedRowHeights_(-JEFF.heightStepPx); }

function _jeffSetSelectedColumnWidths_(pixels) {
  const r = _jeffRange_();
  const s = r.getSheet();
  for (let col = r.getColumn(); col <= r.getLastColumn(); col++) s.setColumnWidth(col, pixels);
}
function _jeffSetSelectedRowHeights_(pixels) {
  const r = _jeffRange_();
  const s = r.getSheet();
  for (let row = r.getRow(); row <= r.getLastRow(); row++) s.setRowHeight(row, pixels);
}
function _jeffAdjustSelectedColumnWidths_(delta) {
  const r = _jeffRange_();
  const s = r.getSheet();
  for (let col = r.getColumn(); col <= r.getLastColumn(); col++) {
    s.setColumnWidth(col, Math.max(2, s.getColumnWidth(col) + delta));
  }
}
function _jeffAdjustSelectedRowHeights_(delta) {
  const r = _jeffRange_();
  const s = r.getSheet();
  for (let row = r.getRow(); row <= r.getLastRow(); row++) {
    s.setRowHeight(row, Math.max(2, s.getRowHeight(row) + delta));
  }
}

// -----------------------------------------------------------------------------
// Decimal places
// -----------------------------------------------------------------------------
function jeffIncreaseDecimal() { _jeffChangeDecimals_(1); }
function jeffDecreaseDecimal() { _jeffChangeDecimals_(-1); }

function _jeffChangeDecimals_(delta) {
  const r = _jeffRange_();
  const formats = r.getNumberFormats();
  const out = formats.map(row => row.map(fmt => _jeffAdjustDecimalFormat_(fmt, delta)));
  r.setNumberFormats(out);
}

function _jeffAdjustDecimalFormat_(format, delta) {
  if (!format || String(format).toLowerCase() === 'general') return delta > 0 ? '0.0' : '0';
  // Avoid converting obvious date/time formats. Quoted literal text is removed first.
  const probe = String(format).replace(/"[^"]*"/g, '');
  if (/[ymdhHsS]/.test(probe) && !/%/.test(probe)) return format;

  return String(format).split(';').map(section => {
    let m = section.match(/\.([0#]+)/);
    if (m) {
      if (delta > 0) return section.replace(m[0], '.' + m[1] + '0');
      if (m[1].length > 1) return section.replace(m[0], '.' + m[1].slice(0, -1));
      return section.replace(m[0], '');
    }
    if (delta < 0) return section;

    // Add one decimal immediately after the final integer placeholder before
    // percent / suffix text. This covers the XLAM's configured formats.
    let inQuote = false;
    let insertAt = -1;
    for (let i = 0; i < section.length; i++) {
      if (section[i] === '"') inQuote = !inQuote;
      if (!inQuote && (section[i] === '0' || section[i] === '#')) insertAt = i + 1;
    }
    return insertAt >= 0 ? section.slice(0, insertAt) + '.0' + section.slice(insertAt) : section;
  }).join(';');
}

// -----------------------------------------------------------------------------
// Recalculate selection (best available Sheets approximation)
// -----------------------------------------------------------------------------
function jeffCalcSelection() {
  const r = _jeffRange_();
  const formulas = r.getFormulas();
  let touched = 0;
  for (let i = 0; i < formulas.length; i++) {
    for (let j = 0; j < formulas[i].length; j++) {
      if (!formulas[i][j]) continue;
      try {
        r.getCell(i + 1, j + 1).setFormula(formulas[i][j]);
        touched++;
      } catch (e) {
        // Array/spill output or protected cells may reject rewriting; skip them.
      }
    }
  }
  SpreadsheetApp.flush();
  SpreadsheetApp.getActive().toast('Re-triggered ' + touched + ' formula cell(s).', 'Jeff', 3);
}

// -----------------------------------------------------------------------------
// Convert all references in formula cells to absolute / relative A1 references.
// We use Sheets' own R1C1 conversion first, then translate each R1C1 reference.
// -----------------------------------------------------------------------------
function jeffConvertToAbsolute() { _jeffConvertReferences_(true); }
function jeffConvertToRelative() { _jeffConvertReferences_(false); }

function _jeffConvertReferences_(makeAbsolute) {
  const r = _jeffRange_();
  const formulasR1C1 = r.getFormulasR1C1();
  let changed = 0;

  for (let i = 0; i < formulasR1C1.length; i++) {
    for (let j = 0; j < formulasR1C1[i].length; j++) {
      const f = formulasR1C1[i][j];
      if (!f) continue;
      const row = r.getRow() + i;
      const col = r.getColumn() + j;
      const converted = _jeffR1C1ToA1Formula_(f, row, col, makeAbsolute);
      try {
        r.getCell(i + 1, j + 1).setFormula(converted);
        changed++;
      } catch (e) {
        // Leave formula unchanged if an unusual construct cannot be set back.
      }
    }
  }
  SpreadsheetApp.getActive().toast('Converted ' + changed + ' formula cell(s).', 'Jeff', 3);
}

function _jeffR1C1ToA1Formula_(formula, currentRow, currentCol, makeAbsolute) {
  return _jeffMapOutsideDoubleQuotes_(formula, segment =>
    segment.replace(/(^|[^A-Za-z0-9_])R(\[-?\d+\]|\d*)C(\[-?\d+\]|\d*)(?![A-Za-z0-9_])/g, (all, prefix, rPart, cPart) => {
      const targetRow = _jeffR1C1Index_(rPart, currentRow);
      const targetCol = _jeffR1C1Index_(cPart, currentCol);
      if (targetRow < 1 || targetCol < 1) return all;
      const colText = _jeffColumnLetters_(targetCol);
      const ref = makeAbsolute ? ('$' + colText + '$' + targetRow) : (colText + targetRow);
      return prefix + ref;
    })
  );
}

function _jeffR1C1Index_(part, current) {
  if (part === '') return current;
  if (part[0] === '[') return current + Number(part.slice(1, -1));
  return Number(part);
}

function _jeffColumnLetters_(n) {
  let s = '';
  while (n > 0) {
    n--;
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s;
}

function _jeffMapOutsideDoubleQuotes_(text, mapper) {
  // Leave string literals and quoted sheet names untouched.
  const parts = String(text).split(/("(?:""|[^"])*"|'(?:''|[^'])*')/g);
  return parts.map((part, i) => (i % 2 ? part : mapper(part))).join('');
}

// -----------------------------------------------------------------------------
// SuperFill and one-level custom undo
// -----------------------------------------------------------------------------
function jeffSuperFillDownEnd() { _jeffSuperFillDown_('toEnd'); }
function jeffSuperFillDownBlank() { _jeffSuperFillDown_('toBlank'); }
function jeffSuperFillRightEnd() { _jeffSuperFillRight_('toEnd'); }
function jeffSuperFillRightBlank() { _jeffSuperFillRight_('toBlank'); }

function _jeffSuperFillDown_(endDefn) {
  const source = _jeffRange_();
  const sheet = source.getSheet();
  const startRow = source.getRow();
  const startCol = source.getColumn();
  const endRow = source.getLastRow();
  const endCol = source.getLastColumn();
  const leftCol = Math.max(1, startCol - 1);
  const rightCol = Math.min(sheet.getMaxColumns(), endCol + 1);

  let maxRow = 0;
  for (let col = leftCol; col <= rightCol; col++) {
    maxRow = Math.max(maxRow, _jeffFindLastRow_(sheet, col, endRow, endDefn));
  }
  if (maxRow <= endRow) return;

  const overwritten = sheet.getRange(endRow + 1, startCol, maxRow - endRow, source.getNumColumns());
  _jeffStoreUndo_(overwritten);
  const destination = sheet.getRange(startRow, startCol, maxRow - startRow + 1, source.getNumColumns());
  source.autoFill(destination, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  destination.activate();
}

function _jeffSuperFillRight_(endDefn) {
  const source = _jeffRange_();
  const sheet = source.getSheet();
  const startRow = source.getRow();
  const startCol = source.getColumn();
  const endRow = source.getLastRow();
  const endCol = source.getLastColumn();
  const topRow = Math.max(1, startRow - 1);
  const bottomRow = Math.min(sheet.getMaxRows(), endRow + 1);

  let maxCol = 0;
  for (let row = topRow; row <= bottomRow; row++) {
    maxCol = Math.max(maxCol, _jeffFindLastCol_(sheet, row, endCol, endDefn));
  }
  if (maxCol <= endCol) return;

  const overwritten = sheet.getRange(startRow, endCol + 1, source.getNumRows(), maxCol - endCol);
  _jeffStoreUndo_(overwritten);
  const destination = sheet.getRange(startRow, startCol, source.getNumRows(), maxCol - startCol + 1);
  source.autoFill(destination, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  destination.activate();
}

function _jeffFindLastRow_(sheet, col, startRow, endDefn) {
  const maxRows = sheet.getMaxRows();
  if (endDefn === 'toEnd') {
    const bottom = sheet.getRange(maxRows, col);
    if (!_jeffCellIsEmpty_(bottom)) return maxRows;
    return bottom.getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  }
  if (startRow > maxRows) return 0;
  const start = sheet.getRange(startRow, col);
  if (_jeffCellIsEmpty_(start)) return 0;
  if (startRow === maxRows) return startRow;
  const below = sheet.getRange(startRow + 1, col);
  if (_jeffCellIsEmpty_(below)) return startRow;
  return start.getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
}

function _jeffFindLastCol_(sheet, row, startCol, endDefn) {
  const maxCols = sheet.getMaxColumns();
  if (endDefn === 'toEnd') {
    const right = sheet.getRange(row, maxCols);
    if (!_jeffCellIsEmpty_(right)) return maxCols;
    return right.getNextDataCell(SpreadsheetApp.Direction.PREVIOUS).getColumn();
  }
  if (startCol > maxCols) return 0;
  const start = sheet.getRange(row, startCol);
  if (_jeffCellIsEmpty_(start)) return 0;
  if (startCol === maxCols) return startCol;
  const next = sheet.getRange(row, startCol + 1);
  if (_jeffCellIsEmpty_(next)) return startCol;
  return start.getNextDataCell(SpreadsheetApp.Direction.NEXT).getColumn();
}


function _jeffCellIsEmpty_(cell) {
  return cell.getFormula() === '' && cell.getValue() === '';
}

function _jeffStoreUndo_(range) {
  const payload = {
    spreadsheetId: SpreadsheetApp.getActive().getId(),
    sheetId: range.getSheet().getSheetId(),
    a1: range.getA1Notation(),
    values: range.getValues().map(row => row.map(_jeffPackValue_)),
    formulas: range.getFormulas(),
  };
  const json = JSON.stringify(payload);
  const props = PropertiesService.getUserProperties();
  _jeffClearUndoProperties_(props);

  // Chunking avoids the per-property size ceiling. 6000 chars is deliberately
  // conservative for UTF-8 property storage.
  const chunkSize = 6000;
  const count = Math.ceil(json.length / chunkSize);
  props.setProperty(JEFF.undoPrefix + 'COUNT', String(count));
  for (let i = 0; i < count; i++) {
    props.setProperty(JEFF.undoPrefix + i, json.slice(i * chunkSize, (i + 1) * chunkSize));
  }
}

function jeffUndoSuperFill() {
  const props = PropertiesService.getUserProperties();
  const count = Number(props.getProperty(JEFF.undoPrefix + 'COUNT') || 0);
  if (!count) return;
  let json = '';
  for (let i = 0; i < count; i++) json += props.getProperty(JEFF.undoPrefix + i) || '';

  let payload;
  try { payload = JSON.parse(json); } catch (e) { _jeffClearUndoProperties_(props); return; }
  if (payload.spreadsheetId !== SpreadsheetApp.getActive().getId()) {
    SpreadsheetApp.getActive().toast('Last SuperFill belongs to a different spreadsheet.', 'Jeff', 5);
    return;
  }
  const sheet = SpreadsheetApp.getActive().getSheets().find(s => s.getSheetId() === payload.sheetId);
  if (!sheet) { _jeffClearUndoProperties_(props); return; }

  const target = sheet.getRange(payload.a1);
  const values = payload.values.map(row => row.map(_jeffUnpackValue_));
  target.setValues(values);
  // Restore formulas after constants, matching the XLAM's .Formula undo data.
  for (let i = 0; i < payload.formulas.length; i++) {
    for (let j = 0; j < payload.formulas[i].length; j++) {
      if (payload.formulas[i][j]) target.getCell(i + 1, j + 1).setFormula(payload.formulas[i][j]);
    }
  }
  sheet.activate();
  target.activate();
  _jeffClearUndoProperties_(props);
}

function _jeffPackValue_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') return { t: 'date', v: v.toISOString() };
  return { t: 'value', v: v };
}
function _jeffUnpackValue_(p) {
  return p && p.t === 'date' ? new Date(p.v) : (p ? p.v : '');
}
function _jeffClearUndoProperties_(props) {
  const all = props.getProperties();
  Object.keys(all).forEach(k => { if (k.indexOf(JEFF.undoPrefix) === 0) props.deleteProperty(k); });
}

// -----------------------------------------------------------------------------
// Date-stamped "Save As" -> Google Drive dated copy
// -----------------------------------------------------------------------------
function jeffSaveAsWithDate() {
  const ss = SpreadsheetApp.getActive();
  SpreadsheetApp.flush();
  const file = DriveApp.getFileById(ss.getId());
  const tz = ss.getSpreadsheetTimeZone();
  const stamp = Utilities.formatDate(new Date(), tz, JEFF.datePattern);
  let base = file.getName();

  // Remove an existing yymmdd / yyyymmdd, optionally with -hhmm, from front or end.
  base = base.replace(/^(?:\d{6}|\d{8})(?:-\d{4})?\s+/, '');
  base = base.replace(/\s+(?:\d{6}|\d{8})(?:-\d{4})?$/, '');
  const newName = JEFF.dateAtFront ? (stamp + ' ' + base) : (base + ' ' + stamp);

  let copy;
  const parents = file.getParents();
  if (parents.hasNext()) copy = file.makeCopy(newName, parents.next());
  else copy = file.makeCopy(newName);

  SpreadsheetApp.getUi().alert('Dated copy created:\n' + newName + '\n\n' + copy.getUrl());
}

// -----------------------------------------------------------------------------
// Help / diagnostics
// -----------------------------------------------------------------------------
function jeffShowMigrationNotes() {
  const msg = [
    "Jeff's XLAM -> Google Sheets migration notes",
    '',
    'Exact / near-exact: number/font/fill cycles, wrap, row/column sizes, alignment cycles,',
    'border cycle, SuperFill logic, one-level SuperFill undo, formula absolute/relative conversion, dated copy.',
    '',
    'Approximate only:',
    '- Excel worksheet zoom: Apps Script exposes no Sheets viewport-zoom API. AutoHotkey uses browser zoom instead.',
    '- Center Across Selection: Sheets has no equivalent; ordinary center is used without merging.',
    '- Single Accounting Underline: Sheets has no equivalent; ordinary underline is used.',
    '- Calculate Selection: Sheets recalculates automatically; this port re-sets formulas to trigger recalculation.',
    '- No-fill vs explicit white fill cannot be distinguished reliably with the basic Apps Script service.',
    '',
    'Not reproduced destructively:',
    '- Cell indent. Sheets exposes no cell-indent property, so this port does not add spaces to values.',
    '',
    'XLAM defects found:',
    '- Actual VBA: Ctrl+Shift+K = Zoom In, Ctrl+Shift+J = Zoom Out (manual has them reversed).',
    '- Ctrl+Alt+Shift+O points to missing Sub ShowOptions in the XLAM.',
    '- Menu Calculate Selection points to CalculateSelection, but the real routine is CalcSelection.',
    '- DeactivateKeyboardShortcuts forgets to unbind Ctrl+Alt+D and Ctrl+Alt+R.',
    '- DefaultRowHeight is declared Long: assigning 13.2 actually coerces the runtime value to 13.',
    '- Keyboard-layout labels/comments are stale for layouts 2/3; current file is layout 1, so comma/period is unambiguous.',
  ].join('\n');
  SpreadsheetApp.getUi().alert(msg);
}

// -----------------------------------------------------------------------------
// Generic helpers
// -----------------------------------------------------------------------------
function _jeffRange_() {
  const r = SpreadsheetApp.getActiveRange();
  if (!r) throw new Error('Select a cell or range first.');
  return r;
}

function _jeffNext_(current, list, normalizer) {
  const c = normalizer(current);
  for (let i = 0; i < list.length; i++) {
    if (normalizer(list[i]) === c) return list[(i + 1) % list.length];
  }
  return list[0];
}

// ============================================================
// Excel Toolbar
// ============================================================

// =====================================================
// 参照元セル
// =====================================================

function excelTracePrecedents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const cell = ss.getActiveCell();

  const formula = cell.getFormula();

  if (!formula) {
    SpreadsheetApp.getUi().alert(
      '選択したセルには数式がありません。'
    );
    return;
  }

  saveTraceOrigin_(sheet, cell);

  const refs = parseReferences_(formula, sheet.getName());

  if (refs.length === 0) {
    SpreadsheetApp.getUi().alert(
      '直接参照しているセルを検出できませんでした。'
    );
    return;
  }

  const sameSheetRefs = [
    ...new Set(
      refs
        .filter(r => r.sheet === sheet.getName())
        .map(r => r.a1)
    )
  ];

  if (sameSheetRefs.length > 0) {
    const rangeList = sheet.getRangeList(sameSheetRefs);
    sheet.setActiveRangeList(rangeList);
  }

  const text = refs
    .map(r => `${r.sheet}!${r.a1}`)
    .join(', ');

  ss.toast(
    `参照元: ${text}`,
    'Excel Toolbar',
    8
  );
}


// =====================================================
// 参照先セル
// =====================================================

function excelTraceDependents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const targetSheet = ss.getActiveSheet();
  const targetCell = ss.getActiveCell();

  saveTraceOrigin_(targetSheet, targetCell);

  const targetRow = targetCell.getRow();
  const targetColumn = targetCell.getColumn();
  const targetSheetName = targetSheet.getName();

  const dependents = {};

  ss.getSheets().forEach(sheet => {

    const dataRange = sheet.getDataRange();
    const formulas = dataRange.getFormulas();

    for (let r = 0; r < formulas.length; r++) {

      for (let c = 0; c < formulas[r].length; c++) {

        const formula = formulas[r][c];

        if (!formula) continue;

        const refs =
          parseReferences_(formula, sheet.getName());

        const dependsOnTarget = refs.some(ref => {

          if (ref.sheet !== targetSheetName) {
            return false;
          }

          return rangeContainsCell_(
            targetSheet,
            ref.a1,
            targetRow,
            targetColumn
          );
        });

        if (dependsOnTarget) {

          if (!dependents[sheet.getName()]) {
            dependents[sheet.getName()] = [];
          }

          dependents[sheet.getName()].push(
            sheet.getRange(
              dataRange.getRow() + r,
              dataRange.getColumn() + c
            ).getA1Notation()
          );
        }
      }
    }
  });

  const sheets = Object.keys(dependents);

  if (sheets.length === 0) {
    SpreadsheetApp.getUi().alert(
      '直接参照しているセルは見つかりませんでした。'
    );
    return;
  }

  // 現在のシートに参照先があれば優先
  const displaySheetName =
    dependents[targetSheetName]
      ? targetSheetName
      : sheets[0];

  const displaySheet =
    ss.getSheetByName(displaySheetName);

  ss.setActiveSheet(displaySheet);

  const rangeList =
    displaySheet.getRangeList(
      [...new Set(dependents[displaySheetName])]
    );

  displaySheet.setActiveRangeList(rangeList);

  const text = sheets
    .map(name =>
      `${name}!${dependents[name].join(',')}`
    )
    .join(' / ');

  ss.toast(
    `参照先: ${text}`,
    'Excel Toolbar',
    8
  );
}


// =====================================================
// トレース解除
// =====================================================

function excelClearTrace() {

  const prop =
    PropertiesService
      .getUserProperties()
      .getProperty('TRACE_ORIGIN');

  if (!prop) return;

  const origin = JSON.parse(prop);

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(origin.sheet);

  if (!sheet) return;

  ss.setActiveSheet(sheet);

  sheet
    .getRange(origin.a1)
    .activate();

  PropertiesService
    .getUserProperties()
    .deleteProperty('TRACE_ORIGIN');
}


function saveTraceOrigin_(sheet, cell) {

  PropertiesService
    .getUserProperties()
    .setProperty(
      'TRACE_ORIGIN',
      JSON.stringify({
        sheet: sheet.getName(),
        a1: cell.getA1Notation()
      })
    );
}


// =====================================================
// 数式からA1参照を抽出
// =====================================================

function parseReferences_(formula, defaultSheet) {

  const refs = [];

  const regex =
    /(?:(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))!)?(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)/g;

  let match;

  while ((match = regex.exec(formula)) !== null) {

    let sheet =
      match[1] || match[2] || defaultSheet;

    sheet =
      sheet.replace(/''/g, "'");

    const a1 =
      match[3].replace(/\$/g, '');

    refs.push({
      sheet: sheet,
      a1: a1
    });
  }

  return refs;
}


function rangeContainsCell_(
  sheet,
  a1,
  row,
  column
) {

  try {

    const range =
      sheet.getRange(a1);

    return (
      row >= range.getRow() &&
      row <= range.getLastRow() &&
      column >= range.getColumn() &&
      column <= range.getLastColumn()
    );

  } catch (e) {

    return false;
  }
}


// =====================================================
// バックアップ
// =====================================================

function excelSaveBackUp() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const file =
    DriveApp.getFileById(ss.getId());

  const timezone =
    Session.getScriptTimeZone();

  const timestamp =
    Utilities.formatDate(
      new Date(),
      timezone,
      'yyyyMMdd_HHmmss'
    );

  const name =
    `${ss.getName()}_Backup_${timestamp}`;

  const parents =
    file.getParents();

  if (parents.hasNext()) {

    file.makeCopy(
      name,
      parents.next()
    );

  } else {

    file.makeCopy(name);
  }

  ss.toast(
    `バックアップを作成しました: ${name}`,
    'Excel Toolbar',
    5
  );
}


// =====================================================
// XLSXを添付したGmail下書き
// =====================================================

function excelSendAsAttachment() {

  const ui =
    SpreadsheetApp.getUi();

  const response =
    ui.prompt(
      '添付メール',
      '送信先メールアドレスを入力してください',
      ui.ButtonSet.OK_CANCEL
    );

  if (
    response.getSelectedButton()
    !== ui.Button.OK
  ) {
    return;
  }

  const email =
    response.getResponseText().trim();

  if (!email) return;

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const url =
    `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=xlsx`;

  const blob =
    UrlFetchApp.fetch(
      url,
      {
        headers: {
          Authorization:
            'Bearer ' +
            ScriptApp.getOAuthToken()
        }
      }
    )
    .getBlob()
    .setName(
      ss.getName() + '.xlsx'
    );

  GmailApp.createDraft(
    email,
    ss.getName(),
    '',
    {
      attachments: [blob]
    }
  );

  ui.alert(
    'Gmailに添付メールの下書きを作成しました。'
  );
}


// =====================================================
// アウトライン
// =====================================================

function excelCollapseOutline() {

  const range =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getActiveRange();

  range.collapseGroups();
}


function excelExpandOutline() {

  const range =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getActiveRange();

  range.expandGroups();
}


// =====================================================
// 重複削除
// =====================================================

function excelRemoveDuplicates() {

  const ui =
    SpreadsheetApp.getUi();

  const result =
    ui.alert(
      '重複の削除',
      '選択範囲内の重複行を削除します。続行しますか？',
      ui.ButtonSet.YES_NO
    );

  if (result !== ui.Button.YES) {
    return;
  }

  SpreadsheetApp
    .getActiveSpreadsheet()
    .getActiveRange()
    .removeDuplicates();
}


// =====================================================
// 塗りつぶし
// =====================================================

function excelSetFillColor() {

  const ui =
    SpreadsheetApp.getUi();

  const result =
    ui.prompt(
      '塗りつぶし',
      '色をHEX形式で入力してください（例 #FFF2CC）',
      ui.ButtonSet.OK_CANCEL
    );

  if (
    result.getSelectedButton()
    !== ui.Button.OK
  ) {
    return;
  }

  let color =
    result.getResponseText().trim();

  if (!/^#[0-9A-Fa-f]{6}$/.test(color)) {

    ui.alert(
      '#RRGGBB形式で入力してください。'
    );

    return;
  }

  SpreadsheetApp
    .getActiveSpreadsheet()
    .getActiveRange()
    .setBackground(color);
}


// =====================================================
// Goal Seek
// =====================================================

function excelGoalSeek() {

  const ui =
    SpreadsheetApp.getUi();

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getActiveSheet();

  let response =
    ui.prompt(
      'Goal Seek',
      '結果セル（数式セル）を入力してください。例: D10',
      ui.ButtonSet.OK_CANCEL
    );

  if (
    response.getSelectedButton()
    !== ui.Button.OK
  ) return;

  const resultCell =
    sheet.getRange(
      response.getResponseText().trim()
    );

  if (!resultCell.getFormula()) {

    ui.alert(
      '結果セルには数式が必要です。'
    );

    return;
  }

  response =
    ui.prompt(
      'Goal Seek',
      '目標値を入力してください。',
      ui.ButtonSet.OK_CANCEL
    );

  if (
    response.getSelectedButton()
    !== ui.Button.OK
  ) return;

  const target =
    Number(
      response.getResponseText()
    );

  if (!Number.isFinite(target)) {

    ui.alert(
      '目標値は数値で入力してください。'
    );

    return;
  }

  response =
    ui.prompt(
      'Goal Seek',
      '変化させるセルを入力してください。例: B5',
      ui.ButtonSet.OK_CANCEL
    );

  if (
    response.getSelectedButton()
    !== ui.Button.OK
  ) return;

  const changeCell =
    sheet.getRange(
      response.getResponseText().trim()
    );

  const original =
    changeCell.getValue();

  let x0 =
    Number(original);

  if (!Number.isFinite(x0)) {

    ui.alert(
      '変化させるセルには数値が必要です。'
    );

    return;
  }

  function f(x) {

    changeCell.setValue(x);

    SpreadsheetApp.flush();

    return (
      Number(resultCell.getValue())
      - target
    );
  }

  let x1 =
    x0 +
    (Math.abs(x0) > 1
      ? Math.abs(x0) * 0.01
      : 1);

  let f0 = f(x0);
  let f1 = f(x1);

  const tolerance =
    1e-8 *
    Math.max(1, Math.abs(target));

  let solved = false;

  for (let i = 0; i < 100; i++) {

    if (Math.abs(f1) <= tolerance) {

      solved = true;
      break;
    }

    const denominator =
      f1 - f0;

    if (
      Math.abs(denominator)
      < 1e-14
    ) {

      x1 +=
        Math.max(
          1,
          Math.abs(x1) * 0.05
        );

      f1 = f(x1);

      continue;
    }

    const x2 =
      x1 -
      f1 *
      (x1 - x0) /
      denominator;

    if (!Number.isFinite(x2)) {
      break;
    }

    x0 = x1;
    f0 = f1;

    x1 = x2;
    f1 = f(x1);
  }

  if (solved) {

    changeCell.setValue(x1);

    SpreadsheetApp.flush();

    ui.alert(
      `Goal Seek完了\n\n${changeCell.getA1Notation()} = ${x1}\n${resultCell.getA1Notation()} = ${resultCell.getValue()}`
    );

  } else {

    changeCell.setValue(original);

    SpreadsheetApp.flush();

    ui.alert(
      '収束しませんでした。元の値に戻しました。'
    );
  }
}

