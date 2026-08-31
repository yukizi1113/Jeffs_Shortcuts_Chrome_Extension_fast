'use strict';

const JEFF = Object.freeze({
  numberFormats: [
    '#,##0.0;(#,##0.0)',
    '$#,##0.0;($#,##0.0)',
    '#,##0.0%;(#,##0.0%)',
    '#,##0.0"x"',
    'General'
  ],
  fontColors: ['#0000FF', '#00B050', '#FFFFFF', '#000000'],
  fillColors: [null, '#FFFF00', '#CCFFFF', '#C0C0C0', '#002960'],
  specifiedFill: '#002960',
  specifiedFont: '#FFFFFF',
  defaultFont: '#000000',
  defaultColumnWidthPx: 61,
  defaultRowHeightPx: 17,
  widthStepPx: 7,
  heightStepPx: 3,
  squareWidthPx: 20,
  datePattern: 'yyyyMMdd',
  dateAtFront: true
});

const metaCache = new Map();
const stateCache = new Map();
// Short-lived optimistic cache for rapid color toggles on the same selection.
// First press reads the sheet; subsequent presses within 5s need only one write each.
const toggleCache = new Map();

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  (async () => {
    try {
      if (msg?.type === 'JEFF_AUTHORIZE') {
        await getToken(true);
        sendResponse({ok: true});
        return;
      }
      if (msg?.type === 'JEFF_RUN') {
        const result = await runJeff(msg.command, msg.context, sender);
        sendResponse({ok: true, ...(result || {})});
        return;
      }
      sendResponse({ok: false, error: 'Unknown message.'});
    } catch (e) {
      const text = e?.message || String(e);
      sendResponse({
        ok: false,
        code: text.includes('AUTH_REQUIRED') ? 'AUTH_REQUIRED' : undefined,
        error: text.replace(/^AUTH_REQUIRED:\s*/, '')
      });
    }
  })();
  return true;
});

async function getToken(interactive = false) {
  try {
    const result = await chrome.identity.getAuthToken({interactive});
    const token = typeof result === 'string' ? result : result?.token;
    if (!token) throw new Error('No OAuth token returned.');
    return token;
  } catch (e) {
    if (!interactive) throw new Error('AUTH_REQUIRED: Googleアカウントの承認が必要です。');
    throw e;
  }
}

async function apiFetch(url, options = {}, retry = true) {
  const token = await getToken(false);
  const headers = new Headers(options.headers || {});
  headers.set('Authorization', 'Bearer ' + token);
  if (options.body && !headers.has('Content-Type')) headers.set('Content-Type', 'application/json');

  const res = await fetch(url, {...options, headers});
  if (res.status === 401 && retry) {
    await chrome.identity.removeCachedAuthToken({token});
    return apiFetch(url, options, false);
  }
  if (!res.ok) {
    let detail = '';
    try {
      const j = await res.json();
      detail = j?.error?.message || JSON.stringify(j);
    } catch (_) {
      detail = await res.text();
    }
    throw new Error(`Google API ${res.status}: ${detail || res.statusText}`);
  }
  if (res.status === 204) return null;
  const text = await res.text();
  return text ? JSON.parse(text) : null;
}

async function runJeff(command, ctx, sender) {
  if (command === 'ZIN' || command === 'ZOUT' || command === 'Z100') {
    return browserZoom(command, sender?.tab?.id);
  }

  validateContext(ctx);
  const meta = await getMeta(ctx.spreadsheetId);
  const sheet = meta.sheets.find(s => Number(s.sheetId) === Number(ctx.sheetId));
  if (!sheet) throw new Error('現在のシートをSheets APIで特定できませんでした。ページを再読み込みしてください。');
  const grid = parseA1ToGridRange(ctx.a1, Number(ctx.sheetId), sheet.gridProperties);

  const env = {ctx, meta, sheet, grid};
  switch (command) {
    case '11': return numberFormatToggle(env);
    case '12': return fontColorToggle(env);
    case '13': return fillColorToggle(env);
    case '17': return saveAsWithDate(env);
    case '18': return superFill(env, 'ROWS', 'toEnd');
    case '19': return superFill(env, 'ROWS', 'toBlank');
    case '21': return undoSuperFill(env);
    case '24': return superFill(env, 'COLUMNS', 'toEnd');
    case '25': return superFill(env, 'COLUMNS', 'toBlank');
    case '26': return recalcSelection(env);
    case '27': return convertReferences(env, true);
    case '28': return convertReferences(env, false);
    case '31': return setFormat(env, {horizontalAlignment: 'CENTER'}, 'userEnteredFormat.horizontalAlignment');
    case '32': return setFormat(env, {horizontalAlignment: 'CENTER'}, 'userEnteredFormat.horizontalAlignment');
    case '33': return setFormat(env, {textFormat: {foregroundColorStyle: rgbStyle('#000000')}}, 'userEnteredFormat.textFormat.foregroundColorStyle');
    case '34': return clearFormatField(env, 'userEnteredFormat.backgroundColorStyle');
    case '35': return wrapToggle(env);
    case '36': return {toast: 'Sheetsには非破壊のセルインデントAPIがありません。変更しませんでした。'};
    case '37': return {toast: 'Sheetsには非破壊のセルインデントAPIがありません。変更しませんでした。'};
    case '38': return tableTopToggle(env);
    case '39': return specifiedStyleToggle(env);
    case '41': return setDefaultFont(env);
    case '42': return borderToggle(env);
    case '43': return clearBorders(env);
    case '44': return setDimension(env, 'COLUMNS', JEFF.defaultColumnWidthPx);
    case '45': return setDimension(env, 'ROWS', JEFF.defaultRowHeightPx);
    case '46': return adjustDimension(env, 'COLUMNS', JEFF.widthStepPx);
    case '47': return adjustDimension(env, 'COLUMNS', -JEFF.widthStepPx);
    case '48': return adjustDimension(env, 'ROWS', JEFF.heightStepPx);
    case '49': return adjustDimension(env, 'ROWS', -JEFF.heightStepPx);
    case '51': return setDimension(env, 'COLUMNS', JEFF.squareWidthPx);
    case '52': return verticalAlignmentToggle(env);
    case '53': return horizontalAlignmentToggle(env);
    case '54': return changeDecimals(env, -1);
    case '55': return changeDecimals(env, 1);
    case '59': return {toast: 'Jeff Fast: Chrome拡張版。Apps Scriptを介さず、全Google Sheetsで動作します。'};
    case 'PRECEDENT':
      return jumpToPrecedent(env);
    case 'FILTER_TOGGLE':
      return toggleBasicFilter(env);
    case 'EXCEL_NUMBER':
      return excelNumberFormat(env);
    case 'EXCEL_PERCENT':
      return excelPercentFormat(env);
    default:
      throw new Error('未定義のJeffコマンド: ' + command);
  }
}

function validateContext(ctx) {
  if (!ctx?.spreadsheetId) throw new Error('Spreadsheet IDを取得できませんでした。');
  if (!Number.isFinite(Number(ctx?.sheetId))) throw new Error('Sheet IDを取得できませんでした。');
  if (!ctx?.a1) throw new Error('選択範囲を取得できませんでした。');
}

async function getMeta(spreadsheetId, force = false) {
  if (!force && metaCache.has(spreadsheetId)) return metaCache.get(spreadsheetId);
  const fields = encodeURIComponent('properties(title,timeZone),sheets(properties(sheetId,title,gridProperties))');
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}?fields=${fields}`;
  const j = await apiFetch(url);
  const meta = {
    title: j?.properties?.title || '',
    timeZone: j?.properties?.timeZone || 'Etc/UTC',
    sheets: (j?.sheets || []).map(s => s.properties)
  };
  metaCache.set(spreadsheetId, meta);
  return meta;
}

async function getGridData(spreadsheetId, gridRange, fieldMask = null) {
  const fields = fieldMask || 'sheets(properties(sheetId),data(startRow,startColumn,rowData(values(userEnteredValue,userEnteredFormat,effectiveFormat,effectiveValue,formattedValue)),rowMetadata(pixelSize),columnMetadata(pixelSize)))';
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}:getByDataFilter?fields=${encodeURIComponent(fields)}`;
  const j = await apiFetch(url, {
    method: 'POST',
    body: JSON.stringify({dataFilters: [{gridRange}], includeGridData: true})
  });
  const sh = (j?.sheets || []).find(s => Number(s?.properties?.sheetId) === Number(gridRange.sheetId));
  return sh?.data?.[0] || {startRow: gridRange.startRowIndex || 0, startColumn: gridRange.startColumnIndex || 0, rowData: []};
}

async function batchUpdate(spreadsheetId, requests) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}:batchUpdate`;
  return apiFetch(url, {method: 'POST', body: JSON.stringify({requests})});
}

function parseA1ToGridRange(a1, sheetId, gp = {}) {
  let s = String(a1).trim();
  const bang = s.lastIndexOf('!');
  if (bang >= 0) s = s.slice(bang + 1);
  s = s.replace(/\$/g, '');

  const rows = Number(gp.rowCount || 1000);
  const cols = Number(gp.columnCount || 26);

  const cell = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i.exec(s);
  if (cell) {
    const c1 = lettersToCol(cell[1]);
    const r1 = Number(cell[2]) - 1;
    const c2 = cell[3] ? lettersToCol(cell[3]) : c1;
    const r2 = cell[4] ? Number(cell[4]) - 1 : r1;
    return {
      sheetId,
      startRowIndex: Math.min(r1, r2),
      endRowIndex: Math.max(r1, r2) + 1,
      startColumnIndex: Math.min(c1, c2),
      endColumnIndex: Math.max(c1, c2) + 1
    };
  }

  const colRange = /^([A-Z]+):([A-Z]+)$/i.exec(s);
  if (colRange) {
    const c1 = lettersToCol(colRange[1]);
    const c2 = lettersToCol(colRange[2]);
    return {sheetId, startRowIndex: 0, endRowIndex: rows, startColumnIndex: Math.min(c1,c2), endColumnIndex: Math.max(c1,c2)+1};
  }

  const rowRange = /^(\d+):(\d+)$/.exec(s);
  if (rowRange) {
    const r1 = Number(rowRange[1]) - 1;
    const r2 = Number(rowRange[2]) - 1;
    return {sheetId, startRowIndex: Math.min(r1,r2), endRowIndex: Math.max(r1,r2)+1, startColumnIndex: 0, endColumnIndex: cols};
  }

  throw new Error(`選択範囲「${a1}」をA1形式として解釈できませんでした。名前付き範囲ではなく通常セル範囲を選択してください。`);
}

function lettersToCol(s) {
  let n = 0;
  for (const ch of String(s).toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64;
  return n - 1;
}

function colToLetters(n0) {
  let n = n0 + 1, s = '';
  while (n > 0) {
    n--;
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s;
}

function gridToA1(g) {
  const a = colToLetters(g.startColumnIndex) + (g.startRowIndex + 1);
  const b = colToLetters(g.endColumnIndex - 1) + g.endRowIndex;
  return a === b ? a : `${a}:${b}`;
}

function topLeftCell(data) {
  return data?.rowData?.[0]?.values?.[0] || {};
}

function getCellAt(data, absRow, absCol) {
  const r = absRow - Number(data.startRow || 0);
  const c = absCol - Number(data.startColumn || 0);
  if (r < 0 || c < 0) return {};
  return data?.rowData?.[r]?.values?.[c] || {};
}

function isCellEmpty(cell) {
  return !cell || !cell.userEnteredValue || Object.keys(cell.userEnteredValue).length === 0;
}

function normPattern(p) {
  if (!p) return 'general';
  return String(p).replace(/\s+/g, '').toLowerCase();
}

function nextInList(current, list, normalizer = x => x) {
  const c = normalizer(current);
  const i = list.findIndex(x => normalizer(x) === c);
  return i < 0 ? list[0] : list[(i + 1) % list.length];
}

function rgbStyle(hex) {
  const h = hex.replace('#', '');
  return {rgbColor: {
    red: parseInt(h.slice(0,2),16) / 255,
    green: parseInt(h.slice(2,4),16) / 255,
    blue: parseInt(h.slice(4,6),16) / 255
  }};
}

function colorStyleToHex(style) {
  const c = style?.rgbColor;
  if (!c) return null;
  const to = v => Math.max(0, Math.min(255, Math.round((v || 0) * 255))).toString(16).padStart(2, '0').toUpperCase();
  return '#' + to(c.red) + to(c.green) + to(c.blue);
}

function formatCellData(format) {
  return {userEnteredFormat: format};
}

async function setFormat(env, format, fields) {
  await batchUpdate(env.ctx.spreadsheetId, [{repeatCell: {range: env.grid, cell: formatCellData(format), fields}}]);
  return null;
}

async function clearFormatField(env, field) {
  await batchUpdate(env.ctx.spreadsheetId, [{repeatCell: {range: env.grid, cell: {userEnteredFormat: {}}, fields: field}}]);
  return null;
}
// ============================================================
// Google Sheets: filter toggle
// Windows Excel Alt -> A -> T equivalent
// ============================================================

async function toggleBasicFilter(env) {

    const spreadsheetId = env.ctx.spreadsheetId;
    const sheetId = Number(env.grid.sheetId);

    // 現在このシートにBasic Filterが存在するか確認
    const fields = encodeURIComponent(
        'sheets(properties(sheetId),basicFilter)'
    );

    const url =
        `https://sheets.googleapis.com/v4/spreadsheets/` +
        `${encodeURIComponent(spreadsheetId)}?fields=${fields}`;

    const meta = await apiFetch(url);

    const sheet = (meta?.sheets || []).find(
        s => Number(s?.properties?.sheetId) === sheetId
    );

    // すでにフィルターがある場合 → 解除
    if (sheet?.basicFilter) {

        await batchUpdate(
            spreadsheetId,
            [
                {
                    clearBasicFilter: {
                        sheetId: sheetId
                    }
                }
            ]
        );

        return {
            toast: 'フィルターを解除しました。'
        };
    }

    // フィルターがない場合
    // 現在選択している範囲にフィルターを設定
    const range = {
        sheetId: sheetId,
        startRowIndex: env.grid.startRowIndex,
        endRowIndex: env.grid.endRowIndex,
        startColumnIndex: env.grid.startColumnIndex,
        endColumnIndex: env.grid.endColumnIndex
    };

    await batchUpdate(
        spreadsheetId,
        [
            {
                setBasicFilter: {
                    filter: {
                        range: range
                    }
                }
            }
        ]
    );

    return {
        toast: '選択範囲にフィルターを設定しました。'
    };
}

//
// Excel Ctrl+Shift+1
// 桁区切り + 小数2桁
//
async function excelNumberFormat(env) {
  await setFormat(
    env,
    {
      numberFormat: {
        type: 'NUMBER',
        pattern: '#,##0.00'
      }
    },
    'userEnteredFormat.numberFormat'
  );

  return {
    toast: 'Excel形式: 桁区切り（小数2桁）'
  };
}

//
// Excel Ctrl+Shift+5
// パーセンテージ + 小数0桁
//
async function excelPercentFormat(env) {
  await setFormat(
    env,
    {
      numberFormat: {
        type: 'PERCENT',
        pattern: '0%'
      }
    },
    'userEnteredFormat.numberFormat'
  );

  return {
    toast: 'Excel形式: パーセンテージ（小数0桁）'
  };
}

function numberFormatObject(pattern) {
  if (!pattern || normPattern(pattern) === 'general') return null;
  let type = 'NUMBER';
  if (pattern.includes('%')) type = 'PERCENT';
  else if (/[$€£¥]/.test(pattern)) type = 'CURRENCY';
  else if (pattern === '@') type = 'TEXT';
  return {type, pattern};
}

async function numberFormatToggle(env) {
  const data = await getGridData(env.ctx.spreadsheetId, oneCell(env.grid));
  const cell = topLeftCell(data);
  const current = cell?.userEnteredFormat?.numberFormat?.pattern || 'General';
  const next = nextInList(current, JEFF.numberFormats, normPattern);
  const nf = numberFormatObject(next);
  if (!nf) return clearFormatField(env, 'userEnteredFormat.numberFormat');
  return setFormat(env, {numberFormat: nf}, 'userEnteredFormat.numberFormat');
}

function toggleCacheKey(env, kind) {
  return `${kind}:${env.ctx.spreadsheetId}:${env.ctx.sheetId}:${env.ctx.a1}`;
}

function recentToggleState(key) {
  const hit = toggleCache.get(key);
  if (!hit) return null;
  if (Date.now() - hit.at > 5000) {
    toggleCache.delete(key);
    return null;
  }
  return hit.value;
}

async function fontColorToggle(env) {
  const key = toggleCacheKey(env, 'font');
  let current = recentToggleState(key);
  if (current == null) {
    const data = await getGridData(env.ctx.spreadsheetId, oneCell(env.grid));
    const cell = topLeftCell(data);
    current = colorStyleToHex(cell?.userEnteredFormat?.textFormat?.foregroundColorStyle)
      || colorStyleToHex(cell?.effectiveFormat?.textFormat?.foregroundColorStyle)
      || '#000000';
  }
  const next = nextInList(String(current).toUpperCase(), JEFF.fontColors, x => String(x).toUpperCase());
  await setFormat(env, {textFormat: {foregroundColorStyle: rgbStyle(next)}}, 'userEnteredFormat.textFormat.foregroundColorStyle');
  toggleCache.set(key, {value: next, at: Date.now()});
  return null;
}

async function fillColorToggle(env) {
  const key = toggleCacheKey(env, 'fill');
  let current = recentToggleState(key);
  if (current === null) {
    const data = await getGridData(env.ctx.spreadsheetId, oneCell(env.grid));
    const cell = topLeftCell(data);
    current = colorStyleToHex(cell?.userEnteredFormat?.backgroundColorStyle);
  }
  const next = nextInList(current, JEFF.fillColors, x => x == null ? '__NONE__' : String(x).toUpperCase());
  if (next == null) {
    await clearFormatField(env, 'userEnteredFormat.backgroundColorStyle');
  } else {
    await setFormat(env, {backgroundColorStyle: rgbStyle(next)}, 'userEnteredFormat.backgroundColorStyle');
  }
  toggleCache.set(key, {value: next, at: Date.now()});
  return null;
}

async function wrapToggle(env) {
  const data = await getGridData(env.ctx.spreadsheetId, oneCell(env.grid));
  const cell = topLeftCell(data);
  const cur = cell?.userEnteredFormat?.wrapStrategy || cell?.effectiveFormat?.wrapStrategy || 'OVERFLOW_CELL';
  const next = cur === 'WRAP' ? 'OVERFLOW_CELL' : 'WRAP';
  return setFormat(env, {wrapStrategy: next}, 'userEnteredFormat.wrapStrategy');
}

async function specifiedStyleToggle(env) {
  const data = await getGridData(env.ctx.spreadsheetId, oneCell(env.grid));
  const c = topLeftCell(data);
  const f = c?.userEnteredFormat || {};
  const fg = colorStyleToHex(f?.textFormat?.foregroundColorStyle);
  const bg = colorStyleToHex(f?.backgroundColorStyle);
  const isOn = !!f?.textFormat?.bold && fg === JEFF.specifiedFont && bg === JEFF.specifiedFill;
  const format = isOn
    ? {textFormat: {bold: false, foregroundColorStyle: rgbStyle(JEFF.defaultFont)}}
    : {textFormat: {bold: true, foregroundColorStyle: rgbStyle(JEFF.specifiedFont)}, backgroundColorStyle: rgbStyle(JEFF.specifiedFill)};
  const fields = isOn
    ? 'userEnteredFormat.textFormat.bold,userEnteredFormat.textFormat.foregroundColorStyle,userEnteredFormat.backgroundColorStyle'
    : 'userEnteredFormat.textFormat.bold,userEnteredFormat.textFormat.foregroundColorStyle,userEnteredFormat.backgroundColorStyle';
  if (isOn) {
    await batchUpdate(env.ctx.spreadsheetId, [
      {repeatCell: {range: env.grid, cell: formatCellData(format), fields}},
      {repeatCell: {range: env.grid, cell: {userEnteredFormat: {}}, fields: 'userEnteredFormat.backgroundColorStyle'}}
    ]);
    return null;
  }
  return setFormat(env, format, fields);
}

async function verticalAlignmentToggle(env) {
  const data = await getGridData(env.ctx.spreadsheetId, oneCell(env.grid));
  const c = topLeftCell(data);
  const cur = c?.userEnteredFormat?.verticalAlignment || c?.effectiveFormat?.verticalAlignment || 'MIDDLE';
  const next = cur === 'MIDDLE' ? 'BOTTOM' : cur === 'BOTTOM' ? 'TOP' : 'MIDDLE';
  return setFormat(env, {verticalAlignment: next}, 'userEnteredFormat.verticalAlignment');
}

async function horizontalAlignmentToggle(env) {
  const data = await getGridData(env.ctx.spreadsheetId, oneCell(env.grid));
  const c = topLeftCell(data);
  const cur = c?.userEnteredFormat?.horizontalAlignment || c?.effectiveFormat?.horizontalAlignment || 'LEFT';
  const next = cur === 'CENTER' ? 'RIGHT' : cur === 'RIGHT' ? 'LEFT' : 'CENTER';
  return setFormat(env, {horizontalAlignment: next}, 'userEnteredFormat.horizontalAlignment');
}

async function tableTopToggle(env) {
  const data = await getGridData(env.ctx.spreadsheetId, oneCell(env.grid));
  const c = topLeftCell(data);
  const f = c?.userEnteredFormat || {};
  const isOn = (f.verticalAlignment === 'BOTTOM') && !!f?.textFormat?.underline && !!f?.textFormat?.bold;
  if (isOn) {
    const format = {
      verticalAlignment: 'MIDDLE',
      textFormat: {underline: false, bold: false}
    };
    await batchUpdate(env.ctx.spreadsheetId, [
      {repeatCell: {range: env.grid, cell: formatCellData(format), fields: 'userEnteredFormat.verticalAlignment,userEnteredFormat.textFormat.underline,userEnteredFormat.textFormat.bold'}},
      {repeatCell: {range: env.grid, cell: {userEnteredFormat: {}}, fields: 'userEnteredFormat.horizontalAlignment,userEnteredFormat.numberFormat'}}
    ]);
  } else {
    const format = {
      numberFormat: {type: 'TEXT', pattern: '@'},
      horizontalAlignment: 'CENTER',
      verticalAlignment: 'BOTTOM',
      wrapStrategy: 'OVERFLOW_CELL',
      textFormat: {underline: true, bold: true}
    };
    await batchUpdate(env.ctx.spreadsheetId, [{repeatCell: {
      range: env.grid, cell: formatCellData(format),
      fields: 'userEnteredFormat.numberFormat,userEnteredFormat.horizontalAlignment,userEnteredFormat.verticalAlignment,userEnteredFormat.wrapStrategy,userEnteredFormat.textFormat.underline,userEnteredFormat.textFormat.bold'
    }}]);
  }
  return null;
}

async function setDefaultFont(env) {
  const gp = env.sheet.gridProperties;
  const full = {sheetId: env.grid.sheetId, startRowIndex: 0, endRowIndex: gp.rowCount, startColumnIndex: 0, endColumnIndex: gp.columnCount};
  await batchUpdate(env.ctx.spreadsheetId, [
    {repeatCell: {range: full, cell: formatCellData({textFormat: {fontFamily: 'Arial', fontSize: 10}}), fields: 'userEnteredFormat.textFormat.fontFamily,userEnteredFormat.textFormat.fontSize'}},
    {updateSheetProperties: {properties: {sheetId: env.grid.sheetId, gridProperties: {hideGridlines: true}}, fields: 'gridProperties.hideGridlines'}}
  ]);
  return null;
}

function borderStyleObj(style) {
  return style === 'NONE' ? {style: 'NONE'} : {style, colorStyle: rgbStyle('#000000')};
}

async function borderToggle(env) {
  const key = `border:${env.ctx.spreadsheetId}:${env.grid.sheetId}:${gridToA1(env.grid)}`;
  let state = stateCache.get(key);
  if (state == null) {
    const data = await getGridData(env.ctx.spreadsheetId, oneCell(env.grid));
    const c = topLeftCell(data);
    const style = c?.userEnteredFormat?.borders?.bottom?.style || 'NONE';
    state = style === 'SOLID' ? 1 : style === 'DOTTED' ? 2 : style === 'DOUBLE' ? 3 : 0;
  }
  const next = state >= 3 ? 0 : state + 1;
  stateCache.set(key, next);
  if (next === 0) return clearBorders(env);
  const style = next === 1 ? 'SOLID' : next === 2 ? 'DOTTED' : 'DOUBLE';
  await batchUpdate(env.ctx.spreadsheetId, [{updateBorders: {
    range: env.grid,
    bottom: borderStyleObj(style),
    innerHorizontal: borderStyleObj(style)
  }}]);
  return null;
}

async function clearBorders(env) {
  const none = borderStyleObj('NONE');
  await batchUpdate(env.ctx.spreadsheetId, [{updateBorders: {
    range: env.grid,
    top: none, bottom: none, left: none, right: none, innerHorizontal: none, innerVertical: none
  }}]);
  return null;
}

async function setDimension(env, dimension, pixels) {
  const isCols = dimension === 'COLUMNS';
  const range = {
    sheetId: env.grid.sheetId,
    dimension,
    startIndex: isCols ? env.grid.startColumnIndex : env.grid.startRowIndex,
    endIndex: isCols ? env.grid.endColumnIndex : env.grid.endRowIndex
  };
  await batchUpdate(env.ctx.spreadsheetId, [{updateDimensionProperties: {range, properties: {pixelSize: pixels}, fields: 'pixelSize'}}]);
  return null;
}

async function adjustDimension(env, dimension, delta) {
  const data = await getGridData(env.ctx.spreadsheetId, env.grid,
    'sheets(properties(sheetId),data(startRow,startColumn,rowMetadata(pixelSize),columnMetadata(pixelSize)))');
  const isCols = dimension === 'COLUMNS';
  const md = isCols ? (data.columnMetadata || []) : (data.rowMetadata || []);
  const start = isCols ? env.grid.startColumnIndex : env.grid.startRowIndex;
  const end = isCols ? env.grid.endColumnIndex : env.grid.endRowIndex;
  const fallback = isCols ? 100 : 21;
  const requests = [];
  for (let i = start; i < end; i++) {
    const rel = i - start;
    const cur = Number(md?.[rel]?.pixelSize || fallback);
    requests.push({updateDimensionProperties: {
      range: {sheetId: env.grid.sheetId, dimension, startIndex: i, endIndex: i + 1},
      properties: {pixelSize: Math.max(2, cur + delta)}, fields: 'pixelSize'
    }});
  }
  if (requests.length) await batchUpdate(env.ctx.spreadsheetId, requests);
  return null;
}

async function changeDecimals(env, delta) {
  const data = await getGridData(env.ctx.spreadsheetId, env.grid);
  const height = env.grid.endRowIndex - env.grid.startRowIndex;
  const width = env.grid.endColumnIndex - env.grid.startColumnIndex;
  const rows = [];
  for (let r = 0; r < height; r++) {
    const values = [];
    for (let c = 0; c < width; c++) {
      const cell = getCellAt(data, env.grid.startRowIndex + r, env.grid.startColumnIndex + c);
      const fmt = cell?.userEnteredFormat?.numberFormat?.pattern || cell?.effectiveFormat?.numberFormat?.pattern || 'General';
      const pattern = adjustDecimalFormat(fmt, delta);
      const nf = numberFormatObject(pattern);
      values.push(nf ? {userEnteredFormat: {numberFormat: nf}} : {userEnteredFormat: {}});
    }
    rows.push({values});
  }
  await batchUpdate(env.ctx.spreadsheetId, [{updateCells: {
    range: env.grid,
    rows,
    fields: 'userEnteredFormat.numberFormat'
  }}]);
  return null;
}

function adjustDecimalFormat(format, delta) {
  if (!format || normPattern(format) === 'general') return delta > 0 ? '0.0' : '0';
  const probe = String(format).replace(/"[^"]*"/g, '');
  if (/[ymdhHsS]/.test(probe) && !/%/.test(probe)) return format;
  return String(format).split(';').map(section => {
    const m = section.match(/\.([0#]+)/);
    if (m) {
      if (delta > 0) return section.replace(m[0], '.' + m[1] + '0');
      if (m[1].length > 1) return section.replace(m[0], '.' + m[1].slice(0, -1));
      return section.replace(m[0], '');
    }
    if (delta < 0) return section;
    let inQuote = false, insertAt = -1;
    for (let i = 0; i < section.length; i++) {
      if (section[i] === '"') inQuote = !inQuote;
      if (!inQuote && (section[i] === '0' || section[i] === '#')) insertAt = i + 1;
    }
    return insertAt >= 0 ? section.slice(0, insertAt) + '.0' + section.slice(insertAt) : section;
  }).join(';');
}

async function recalcSelection(env) {
  const data = await getGridData(env.ctx.spreadsheetId, env.grid);
  const rows = rowsWithUserEnteredValues(data, env.grid, v => v);
  await batchUpdate(env.ctx.spreadsheetId, [{updateCells: {range: env.grid, rows, fields: 'userEnteredValue'}}]);
  return {toast: '選択範囲の数式を再投入しました。'};
}

async function convertReferences(env, makeAbsolute) {
  const data = await getGridData(env.ctx.spreadsheetId, env.grid);
  const rows = rowsWithUserEnteredValues(data, env.grid, value => {
    if (!value?.formulaValue) return value;
    return {...value, formulaValue: transformFormulaA1(value.formulaValue, makeAbsolute)};
  });
  await batchUpdate(env.ctx.spreadsheetId, [{updateCells: {range: env.grid, rows, fields: 'userEnteredValue'}}]);
  return null;
}

function rowsWithUserEnteredValues(data, grid, mapper) {
  const h = grid.endRowIndex - grid.startRowIndex;
  const w = grid.endColumnIndex - grid.startColumnIndex;
  const rows = [];
  for (let r = 0; r < h; r++) {
    const values = [];
    for (let c = 0; c < w; c++) {
      const cell = getCellAt(data, grid.startRowIndex + r, grid.startColumnIndex + c);
      const v = mapper(cell?.userEnteredValue ? structuredClone(cell.userEnteredValue) : {});
      values.push({userEnteredValue: v || {}});
    }
    rows.push({values});
  }
  return rows;
}

function transformFormulaA1(formula, makeAbsolute) {
  const parts = String(formula).split(/("(?:""|[^"])*")/g);
  return parts.map((part, i) => {
    if (i % 2) return part;
    return part.replace(/(^|[^A-Za-z0-9_])((?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?)([A-Z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_])/g,
      (m, pre, sheet, dc, col, dr, row) => {
        const ref = makeAbsolute ? `$${col}$${row}` : `${col}${row}`;
        return pre + (sheet || '') + ref;
      });
  }).join('');
}

async function superFill(env, dimension, endDefn) {
  const gp = env.sheet.gridProperties;
  const source = env.grid;
  const isRows = dimension === 'ROWS';

  let contextRange;
  if (isRows) {
    const left = Math.max(0, source.startColumnIndex - 1);
    const right = Math.min(gp.columnCount, source.endColumnIndex + 1);
    contextRange = {
      sheetId: source.sheetId,
      startRowIndex: endDefn === 'toEnd' ? 0 : Math.max(0, source.endRowIndex - 1),
      endRowIndex: gp.rowCount,
      startColumnIndex: left,
      endColumnIndex: right
    };
  } else {
    const top = Math.max(0, source.startRowIndex - 1);
    const bottom = Math.min(gp.rowCount, source.endRowIndex + 1);
    contextRange = {
      sheetId: source.sheetId,
      startRowIndex: top,
      endRowIndex: bottom,
      startColumnIndex: endDefn === 'toEnd' ? 0 : Math.max(0, source.endColumnIndex - 1),
      endColumnIndex: gp.columnCount
    };
  }

  const context = await getGridData(env.ctx.spreadsheetId, contextRange,
    'sheets(properties(sheetId),data(startRow,startColumn,rowData(values(userEnteredValue))))');

  let targetEnd;
  if (isRows) {
    let maxRowExclusive = source.endRowIndex;
    for (let c = contextRange.startColumnIndex; c < contextRange.endColumnIndex; c++) {
      const end = endDefn === 'toEnd'
        ? findLastRowExclusive(context, contextRange, c)
        : findContiguousRowEndExclusive(context, contextRange, c, source.endRowIndex - 1);
      maxRowExclusive = Math.max(maxRowExclusive, end);
    }
    targetEnd = maxRowExclusive;
    if (targetEnd <= source.endRowIndex) return null;
  } else {
    let maxColExclusive = source.endColumnIndex;
    for (let r = contextRange.startRowIndex; r < contextRange.endRowIndex; r++) {
      const end = endDefn === 'toEnd'
        ? findLastColExclusive(context, contextRange, r)
        : findContiguousColEndExclusive(context, contextRange, r, source.endColumnIndex - 1);
      maxColExclusive = Math.max(maxColExclusive, end);
    }
    targetEnd = maxColExclusive;
    if (targetEnd <= source.endColumnIndex) return null;
  }

  const overwritten = isRows
    ? {sheetId: source.sheetId, startRowIndex: source.endRowIndex, endRowIndex: targetEnd, startColumnIndex: source.startColumnIndex, endColumnIndex: source.endColumnIndex}
    : {sheetId: source.sheetId, startRowIndex: source.startRowIndex, endRowIndex: source.endRowIndex, startColumnIndex: source.endColumnIndex, endColumnIndex: targetEnd};

  await storeUndo(env.ctx.spreadsheetId, overwritten);

  const fillLength = isRows ? targetEnd - source.endRowIndex : targetEnd - source.endColumnIndex;
  await batchUpdate(env.ctx.spreadsheetId, [{autoFill: {sourceAndDestination: {
    source,
    dimension,
    fillLength
  }}}]);

  const dest = isRows
    ? {...source, endRowIndex: targetEnd}
    : {...source, endColumnIndex: targetEnd};
  return {selectA1: gridToA1(dest)};
}

function findLastRowExclusive(data, range, absCol) {
  let last = 0;
  const rows = data.rowData || [];
  for (let rr = 0; rr < rows.length; rr++) {
    const absRow = (data.startRow || range.startRowIndex) + rr;
    const cell = getCellAt(data, absRow, absCol);
    if (!isCellEmpty(cell)) last = absRow + 1;
  }
  return last;
}

function findLastColExclusive(data, range, absRow) {
  let last = 0;
  const row = data?.rowData?.[absRow - Number(data.startRow || range.startRowIndex)]?.values || [];
  const startCol = Number(data.startColumn || range.startColumnIndex);
  for (let cc = 0; cc < row.length; cc++) {
    if (!isCellEmpty(row[cc])) last = startCol + cc + 1;
  }
  return last;
}

function findContiguousRowEndExclusive(data, range, absCol, startAbsRow) {
  const start = getCellAt(data, startAbsRow, absCol);
  if (isCellEmpty(start)) return 0;
  let end = startAbsRow + 1;
  for (let r = startAbsRow + 1; r < range.endRowIndex; r++) {
    const cell = getCellAt(data, r, absCol);
    if (isCellEmpty(cell)) break;
    end = r + 1;
  }
  return end;
}

function findContiguousColEndExclusive(data, range, absRow, startAbsCol) {
  const start = getCellAt(data, absRow, startAbsCol);
  if (isCellEmpty(start)) return 0;
  let end = startAbsCol + 1;
  for (let c = startAbsCol + 1; c < range.endColumnIndex; c++) {
    const cell = getCellAt(data, absRow, c);
    if (isCellEmpty(cell)) break;
    end = c + 1;
  }
  return end;
}

async function storeUndo(spreadsheetId, range) {
  const data = await getGridData(spreadsheetId, range,
    'sheets(properties(sheetId),data(startRow,startColumn,rowData(values(userEnteredValue))))');
  const payload = {
    spreadsheetId,
    range,
    rowData: data.rowData || []
  };
  await chrome.storage.local.set({JEFF_SUPERFILL_UNDO: payload});
}

async function undoSuperFill(env) {
  const obj = await chrome.storage.local.get('JEFF_SUPERFILL_UNDO');
  const payload = obj?.JEFF_SUPERFILL_UNDO;
  if (!payload) return {toast: '戻せるSuperFillはありません。'};
  if (payload.spreadsheetId !== env.ctx.spreadsheetId) return {toast: '最後のSuperFillは別のSpreadsheetです。'};
  const r = payload.range;
  const h = r.endRowIndex - r.startRowIndex;
  const w = r.endColumnIndex - r.startColumnIndex;
  const rows = [];
  for (let rr = 0; rr < h; rr++) {
    const vals = [];
    for (let cc = 0; cc < w; cc++) {
      const v = payload?.rowData?.[rr]?.values?.[cc]?.userEnteredValue || {};
      vals.push({userEnteredValue: v});
    }
    rows.push({values: vals});
  }
  await batchUpdate(env.ctx.spreadsheetId, [{updateCells: {range: r, rows, fields: 'userEnteredValue'}}]);
  await chrome.storage.local.remove('JEFF_SUPERFILL_UNDO');
  return {selectA1: gridToA1(r)};
}

async function saveAsWithDate(env) {
  const id = env.ctx.spreadsheetId;
  const metaUrl = `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(id)}?fields=name,parents&supportsAllDrives=true`;
  const f = await apiFetch(metaUrl);
  const stamp = formatYmdInTimeZone(new Date(), env.meta.timeZone);
  let base = String(f?.name || env.meta.title || 'Spreadsheet');
  base = base.replace(/^(?:\d{6}|\d{8})(?:-\d{4})?\s+/, '');
  base = base.replace(/\s+(?:\d{6}|\d{8})(?:-\d{4})?$/, '');
  const name = JEFF.dateAtFront ? `${stamp} ${base}` : `${base} ${stamp}`;
  const body = {name};
  if (Array.isArray(f?.parents) && f.parents.length) body.parents = f.parents;
  const copyUrl = `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(id)}/copy?fields=id,name&supportsAllDrives=true`;
  const copy = await apiFetch(copyUrl, {method: 'POST', body: JSON.stringify(body)});
  return {toast: `日付付きコピーを作成しました: ${copy?.name || name}`};
}

function formatYmdInTimeZone(date, timeZone) {
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone,
    year: 'numeric', month: '2-digit', day: '2-digit'
  }).formatToParts(date);
  const o = Object.fromEntries(parts.map(p => [p.type, p.value]));
  return `${o.year}${o.month}${o.day}`;
}

async function browserZoom(command, tabId) {
  if (!tabId) return null;
  if (command === 'Z100') {
    await chrome.tabs.setZoom(tabId, 1.0);
    return null;
  }
  const cur = await chrome.tabs.getZoom(tabId);
  const next = command === 'ZIN' ? Math.min(5, cur + 0.1) : Math.max(0.25, cur - 0.1);
  await chrome.tabs.setZoom(tabId, Math.round(next * 10) / 10);
  return null;
}

//
// Excel Ctrl+[ equivalent:
// 選択中の数式が直接参照しているセルへ移動
//
async function jumpToPrecedent(env) {

  const data = await getGridData(
    env.ctx.spreadsheetId,
    oneCell(env.grid),
    'sheets(properties(sheetId),data(startRow,startColumn,rowData(values(userEnteredValue))))'
  );

  const cell = topLeftCell(data);
  const formula = cell?.userEnteredValue?.formulaValue;

  if (!formula) {
    return {
      toast: '参照元ジャンプ: 選択セルに数式がありません。'
    };
  }

  const refs = extractDirectA1References(formula);

  if (!refs.length) {
    return {
      toast:
        '参照元ジャンプ: 直接のA1形式参照を見つけられませんでした。' +
        'INDIRECT・OFFSET・名前付き範囲などは未対応です。'
    };
  }

  const ref = refs[0];

  let targetSheet = env.sheet.title;

  if (ref.sheetName) {
    const found = env.meta.sheets.find(
      s => s.title === ref.sheetName
    );

    if (!found) {
      return {
        toast:
          `参照元ジャンプ: シート「${ref.sheetName}」が見つかりません。`
      };
    }

    targetSheet = found.title;
  }

  const a1 = ref.range.replace(/\$/g, '');

  const selectA1 = ref.sheetName
    ? `${quoteSheetName(targetSheet)}!${a1}`
    : a1;

  const suffix =
    refs.length > 1
      ? `（直接参照 ${refs.length} 件のうち最初）`
      : '';

  return {
    selectA1,
    toast: `参照元へ移動: ${selectA1}${suffix}`
  };
}


function quoteSheetName(name) {
  return `'${String(name).replace(/'/g, "''")}'`;
}


//
// 数式中の "A1" のような文字列を
// セル参照と誤認しないために文字列部分を隠す
//
function maskDoubleQuotedStrings(formula) {

  let out = '';
  let inString = false;

  for (let i = 0; i < formula.length; i++) {

    const ch = formula[i];

    if (ch === '"') {

      // Google Sheetsで "" は文字列内の "
      if (inString && formula[i + 1] === '"') {
        out += '  ';
        i++;
        continue;
      }

      inString = !inString;
      out += ' ';
      continue;
    }

    out += inString ? ' ' : ch;
  }

  return out;
}


//
// 数式中から直接参照しているA1形式セルを抽出
//
function extractDirectA1References(formula) {

  const s = maskDoubleQuotedStrings(
    String(formula || '')
  );

  const refs = [];

  const re =
    /(^|[^A-Za-z0-9_])(?:(?:'((?:[^']|'')+)'|([A-Za-z0-9_\u3040-\u30FF\u3400-\u9FFF]+))!)?(\$?[A-Z]{1,4}\$?\d+)(?::(\$?[A-Z]{1,4}\$?\d+))?(?![A-Za-z0-9_(])/g;

  let m;

  while ((m = re.exec(s)) !== null) {

    const sheetName =
      m[2]
        ? m[2].replace(/''/g, "'")
        : (m[3] || null);

    const range =
      m[5]
        ? `${m[4]}:${m[5]}`
        : m[4];

    refs.push({
      sheetName,
      range
    });
  }

  return refs;
}


function oneCell(g) {
  return {
    sheetId: g.sheetId,
    startRowIndex: g.startRowIndex,
    endRowIndex: g.startRowIndex + 1,
    startColumnIndex: g.startColumnIndex,
    endColumnIndex: g.startColumnIndex + 1
  };
}
