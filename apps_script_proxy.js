/**
 * ROCKET Dashboard — Google Apps Script Data Proxy
 * -------------------------------------------------
 * Deploy: script.google.com -> New project -> paste code ->
 *         Deploy -> Web app -> Execute as: Me -> Who: Anyone -> Deploy -> Copy URL
 *
 * After deploy: paste URL in index.html, CONFIG.APPS_SCRIPT_URL
 * and set CONFIG.DEMO = false
 *
 * Dashboard Config Sheet ID:
 * 1EKR-czK1UvXZDIJUe5MLb70yZXyEMe1VIYTx15UOY3E
 */

const DASHBOARD_CONFIG_ID = '1EKR-czK1UvXZDIJUe5MLb70yZXyEMe1VIYTx15UOY3E';

function doGet(e) {
  const action = e && e.parameter && e.parameter.action || 'getData';
  let result;
  try {
    if (action === 'getData')        result = getData();
    else if (action === 'getConfig') result = getConfig();
    else result = { error: 'Unknown action' };
  } catch (err) {
    result = { error: err.message, stack: err.stack };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let result;
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'writeNote') {
      result = writeNote(body.key, body.note);
    } else {
      result = { error: 'Unknown POST action' };
    }
  } catch (err) {
    result = { error: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function writeNote(key, note) {
  if (!key) return { error: 'Missing key' };
  const ss = SpreadsheetApp.openById(DASHBOARD_CONFIG_ID);
  let sheet = ss.getSheetByName('Dashboard Notes');
  if (!sheet) {
    sheet = ss.insertSheet('Dashboard Notes');
    sheet.appendRow(['key', 'text', 'date']);
  }
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).toLowerCase());
  const keyCol = headers.indexOf('key');
  const txtCol = headers.indexOf('text');
  const datCol = headers.indexOf('date');
  if (keyCol < 0) return { error: 'Notes sheet missing "key" column' };
  for (let i = 1; i < data.length; i++) {
    if (data[i][keyCol] === key) {
      if (!note) {
        sheet.deleteRow(i + 1);
      } else {
        sheet.getRange(i + 1, txtCol + 1).setValue(note.text || '');
        sheet.getRange(i + 1, datCol + 1).setValue(note.date || '');
      }
      return { ok: true };
    }
  }
  if (note) {
    sheet.appendRow([key, note.text || '', note.date || '']);
  }
  return { ok: true };
}

function getData() {
  const ss = SpreadsheetApp.openById(DASHBOARD_CONFIG_ID);
  const reportsSheet = ss.getSheetByName('Reports');
  if (!reportsSheet) throw new Error('Sheet "Reports" not found in Dashboard Config');

  const reports = getSheetData(reportsSheet);

  const visible = reports.filter(r => {
    const v = String(r.visible || r['видимый'] || '').toLowerCase().trim();
    return v === 'true' || v === '1' || v === 'да';
  });

  const result = {
    report_errors: [],
    months:         [],
    revenue:        [],
    var_costs:      [],
    indirect:       [],
    op_profit:      [],
    interest:       [],
    net_profit:     [],
    margin_pct:     [],
    op_margin_pct:  [],
    np_margin_pct:  [],
    ncf_op:         [],
    ncf_inv:        [],
    ncf_fin:        [],
    dividends:      [],
    cash:           [],
    assets_total:   [],
    equity:         [],
    liabilities:    [],
    current_ratio:  [],
    autonomy:       [],
    de_ratio:       [],
    dfl:            [],
    crm_pct:        [],
    ftd_new:        [],
    ftd_old:        [],
    rev_core:       [],
    rev_influence:  [],
    rev_streams:    [],
    rev_mannequin:  [],
    tr_fb:          [],
    tr_tg:          [],
    payment_cal:    [],
    roadmap_ckp:      null,
    roadmap_tasks:    null,
    dashboard_notes:  {},
  };

  const REPORT_MAP = {
    'opiu':           readPnL,
    'opiu_combo':     readPnL,
    'opiu_consol':    readPnL,
    'opiu_model':     readPnL,
    'dds_common':     readCashflow,
    'dds_operations': readCashflow,
    'balance':        readBalance,
    'traffic':        readTraffic,
    'payment_cal':    readPaymentCal,
    'payment_plan':   readPaymentCal,
    'roadmap_ckp':    readRoadmapCkp,
    'roadmap_tasks':  readRoadmapTasks,
  };

  visible.forEach(report => {
    const reportId = String(
      report['верхний_подвал'] || report['report_type'] || report['id'] || ''
    ).toLowerCase().trim();

    const handler = REPORT_MAP[reportId];
    if (handler && report.sheet_id && report.sheet_name) {
      try {
        handler(report, result);
      } catch (e) {
        Logger.log('Error reading ' + reportId + ': ' + e.message);
        result.report_errors.push(reportId + ' (' + e.message.substring(0,40) + ')');
      }
    }
  });

  if (result.months.length === 0 && result.revenue.length > 0) {
    result.months = result.revenue.map((_, i) => 'Период ' + (i + 1));
  }

  const N = result.months.length;
  if (N > 0) {
    const DATA_KEYS = [
      'revenue','var_costs','indirect','op_profit','interest','net_profit',
      'margin_pct','op_margin_pct','np_margin_pct',
      'ncf_op','ncf_inv','ncf_fin','dividends','cash',
      'assets_total','equity','liabilities','current_ratio','autonomy','de_ratio','dfl',
      'crm_pct','ftd_new','ftd_old','rev_core','rev_influence','rev_streams','rev_mannequin','tr_fb','tr_tg',
    ];
    DATA_KEYS.forEach(k => {
      if (Array.isArray(result[k]) && result[k].length > N) {
        result[k] = result[k].slice(-N);
      }
    });
  }

  result.threshold_rules  = getThresholdRules(ss);
  result.assumptions      = getAssumptions(ss);
  result.notes            = getNotes(ss);
  result.ai_analysis_raw  = getAiAnalysis(ss);
  result.dashboard_notes  = getDashboardNotes(ss);

  return result;
}

function readPnL(report, result) {
  const ss = getSheetById(report.sheet_id);
  if (!ss) return;
  const sheet = getSheetByName(ss, report.sheet_name);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();

  const KEYS = {
    revenue:    ['выручка', 'доходы', 'revenue', 'итого выручка', 'общая выручка', 'total revenue'],
    var_costs:  ['переменные', 'variable costs', 'трафик (переменные)', 'затраты на трафик'],
    indirect:   ['постоянные', 'косвенные расходы', 'косвенные', 'накладные', 'fixed'],
    op_profit:  ['ebitda', 'ebit', 'операционная прибыль', 'опер', 'операционн', 'прибыль от продаж'],
    interest:   ['проценты', 'interest', 'финансовые расходы', 'ниже ebitda'],
    net_profit: ['чистая прибыль за период', 'чистая прибыль', 'чистая', 'net profit'],
  };

  function rowLabel(row) {
    return String(row[1] || row[0] || row[2] || '').toLowerCase().trim();
  }

  const months = [];
  let headerRow = null;

  for (let i = 0; i < Math.min(10, data.length); i++) {
    const row = data[i];
    const monthCount = row.filter(c => isMonth(c)).length;
    if (monthCount >= 2) { headerRow = i; break; }
  }

  if (headerRow !== null) {
    const header = data[headerRow];
    header.forEach((cell, j) => {
      if (j > 0 && isMonth(cell)) months.push({ col: j, label: formatMonth(cell) });
    });
    if (result.months.length === 0) result.months = months.map(m => m.label);
  }

  for (const [key, keywords] of Object.entries(KEYS)) {
    for (let i = headerRow !== null ? headerRow + 1 : 0; i < data.length; i++) {
      const label = rowLabel(data[i]);
      if (keywords.some(k => label.includes(k))) {
        const vals = months.length > 0
          ? months.map(m => toNum(data[i][m.col]))
          : data[i].slice(1).filter(v => typeof v === 'number' && v !== 0).map(toNum);
        if (vals.length > 0 && result[key].length === 0) {
          result[key] = vals;
        }
        break;
      }
    }
  }

  if (result.revenue.length > 0 && result.var_costs.length > 0) {
    const gross = result.revenue.map((v, i) => v - (result.var_costs[i] || 0));
    if (result.margin_pct.length === 0)
      result.margin_pct = result.revenue.map((v, i) => v > 0 ? +(gross[i]/v*100).toFixed(2) : 0);
  }
  if (result.revenue.length > 0 && result.op_profit.length > 0) {
    if (result.op_margin_pct.length === 0)
      result.op_margin_pct = result.revenue.map((v, i) =>
        v > 0 ? +(result.op_profit[i]/v*100).toFixed(2) : 0);
  }
  if (result.revenue.length > 0 && result.net_profit.length > 0) {
    if (result.np_margin_pct.length === 0)
      result.np_margin_pct = result.revenue.map((v, i) =>
        v > 0 ? +(result.net_profit[i]/v*100).toFixed(2) : 0);
  }
}

function readCashflow(report, result) {
  const ss = getSheetById(report.sheet_id);
  if (!ss) return;
  const sheet = getSheetByName(ss, report.sheet_name);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const KEYS = {
    ncf_op:    ['операцион', 'operating', 'от операционн'],
    ncf_inv:   ['инвестиц', 'invest', 'от инвестиц'],
    ncf_fin:   ['финанс', 'financ', 'от финансов'],
    dividends: ['дивид', 'dividend'],
    cash:      ['остаток', 'баланс', 'cash balance', 'остат денег', 'конечный остаток'],
  };
  extractRows(data, KEYS, result);
}

function readBalance(report, result) {
  const ss = getSheetById(report.sheet_id);
  if (!ss) return;
  const sheet = getSheetByName(ss, report.sheet_name);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const KEYS = {
    assets_total:  ['активы', 'assets total', 'баланс итого', 'итого активы', 'валюта баланса'],
    equity:        ['капитал', 'equity', 'собственный капитал'],
    liabilities:   ['обязательства', 'liabilit', 'итого обязательства'],
    current_ratio: ['ликвидность', 'current ratio', 'текущая ликвидность'],
    autonomy:      ['автоном', 'autonomy', 'коэф.*автоном'],
    de_ratio:      ['d/e', 'долг/капит', 'долг к капиталу'],
    dfl:           ['dfl', 'рычаг', 'финансовый рычаг'],
  };
  extractRows(data, KEYS, result);
}

function readTraffic(report, result) {
  const ss = getSheetById(report.sheet_id);
  if (!ss) return;
  const sheet = getSheetByName(ss, report.sheet_name);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const KEYS = {
    crm_pct:       ['crm%', 'crm %', 'crm доля', 'доля crm'],
    ftd_new:       ['ftd new', 'ftd нов', 'ftd новые'],
    ftd_old:       ['ftd old', 'ftd баз', 'ftd crm', 'ftd базовые'],
    rev_core:      ['core', 'кор', 'core revenue'],
    rev_influence: ['influence', 'инфлюенс'],
    rev_streams:   ['stream', 'стрим'],
    rev_mannequin: ['mannequin', 'манекен'],
    tr_fb:         ['facebook', 'fb бюд', 'fb budget'],
    tr_tg:         ['telegram', 'tg бюд', 'tg budget'],
  };
  extractRows(data, KEYS, result);
}

function readPaymentCal(report, result) {
  Logger.log('Payment calendar: ' + report.sheet_name + ' (not merged)');
}

function readRoadmapCkp(report, result) {
  const ss = getSheetById(report.sheet_id);
  if (!ss) return;
  const candidates = [report.sheet_name, '3 ЦКП', 'ЦКП', '3ЦКП'].filter(Boolean);
  let sheet = null;
  for (const n of candidates) { sheet = ss.getSheetByName(n); if (sheet) break; }
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  let periodRow = -1, periodCols = [];
  for (let i = 0; i < Math.min(6, data.length); i++) {
    data[i].forEach((c, j) => {
      if (/^(янв|фев|мар|апр|май|июн|июл|авг|сен|окт|ноя|дек)/i.test(String(c))) {
        if (periodRow < 0) periodRow = i;
        periodCols.push({ idx: j, label: String(c) });
      }
    });
    if (periodRow >= 0) break;
  }
  if (periodRow < 0 || !periodCols.length) return;
  const blocks = [];
  let currentBlock = null;
  for (let i = periodRow + 1; i < data.length; i++) {
    const row = data[i];
    const col0 = String(row[0] || '').trim();
    const col1 = String(row[1] || '').trim();
    if (!col0 && !col1) continue;
    if (/^ЦКПs*d/i.test(col0)) {
      currentBlock = { id: col0, name: col1, statuses: periodCols.map(p => String(row[p.idx] || '')), criteria: [] };
      blocks.push(currentBlock);
    } else if (currentBlock) {
      const label = col1 || col0;
      if (label) currentBlock.criteria.push({ name: label, values: periodCols.map(p => { const v = row[p.idx]; return (v === '' || v === null || v === undefined) ? null : v; }) });
    }
  }
  result.roadmap_ckp = { periods: periodCols.map(p => p.label), blocks: blocks };
}

function readRoadmapTasks(report, result) {
  const ss = getSheetById(report.sheet_id);
  if (!ss) return;
  const candidates = [report.sheet_name, 'АИ', 'АИ (задачи)', 'Tasks'].filter(Boolean);
  let sheet = null;
  for (const n of candidates) { sheet = ss.getSheetByName(n); if (sheet) break; }
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;
  const headers = data[0].map(h => String(h).toLowerCase().trim());
  const col = {};
  headers.forEach((h, i) => {
    if (h.includes('№') || (h.includes('встреч') && !h.includes('дата'))) col.id = i;
    if (h.includes('дата') && h.includes('встреч')) col.date_meeting = i;
    if (h.includes('проект')) col.project = i;
    if (h.includes('задач') && !h.includes('под')) col.task_type = i;
    if (h.includes('подзадач')) col.subtask = i;
    if (h.includes('ответствен')) col.responsible = i;
    if (h.includes('дата') && h.includes('план')) col.date_plan = i;
    if (h.includes('дата') && (h.includes('факт') || h.includes('исполн'))) col.date_fact = i;
    if (h.includes('статус')) col.status = i;
    if (h.includes('результат')) col.result = i;
  });
  const tasks = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const taskType = col.task_type !== undefined ? String(row[col.task_type] || '').trim() : '';
    if (!taskType) continue;
    tasks.push({
      id: col.id !== undefined ? row[col.id] : i,
      date_meeting: formatRoadmapDate(col.date_meeting !== undefined ? row[col.date_meeting] : null),
      project: col.project !== undefined ? String(row[col.project] || '') : '',
      task_type: taskType,
      subtask: col.subtask !== undefined ? String(row[col.subtask] || '') : '',
      responsible: col.responsible !== undefined ? String(row[col.responsible] || '') : '',
      date_plan: formatRoadmapDate(col.date_plan !== undefined ? row[col.date_plan] : null),
      date_fact: formatRoadmapDate(col.date_fact !== undefined ? row[col.date_fact] : null),
      status: col.status !== undefined ? String(row[col.status] || '') : '',
      result: col.result !== undefined ? String(row[col.result] || '') : ''
    });
  }
  result.roadmap_tasks = tasks;
}

function formatRoadmapDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return val.getFullYear() + '-' + String(val.getMonth()+1).padStart(2,'0') + '-' + String(val.getDate()).padStart(2,'0');
  }
  return String(val);
}

function extractRows(data, KEYS, result) {
  for (const [key, keywords] of Object.entries(KEYS)) {
    for (let i = 0; i < data.length; i++) {
      const label = String(data[i][0] || '').toLowerCase();
      if (keywords.some(k => label.includes(k))) {
        const vals = data[i].slice(1).filter(v => v !== '' && v !== null).map(v => toNum(v));
        if (vals.length > 0 && result[key] && result[key].length === 0) result[key] = vals;
        break;
      }
    }
  }
}

function getSheetById(sheetId) {
  try { return SpreadsheetApp.openById(extractSheetId(sheetId)); }
  catch(e) { Logger.log('Cannot open sheet: ' + sheetId + ' - ' + e.message); return null; }
}

function getSheetByName(ss, name) {
  const m = getAnalysisMonth(), y = getAnalysisYear();
  const resolved = name.replace(/{month}/g, String(m)).replace(/{year}/g, String(y)).replace(/{month_name}/g, getMonthName(m));
  return ss.getSheetByName(resolved) || ss.getSheetByName(name);
}

function extractSheetId(urlOrId) {
  const match = String(urlOrId).match(/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : String(urlOrId);
}

function getSheetData(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(h => String(h).toLowerCase().trim().replace(/\s+/g,'_'));
  return values.slice(1).filter(row => row.some(c => c !== '')).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function isMonth(val) {
  if (!val) return false;
  const s = String(val).toLowerCase();
  return /^(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|янв|фев|мар|апр|май|июн|июл|авг|сен|окт|ноя|дек)/.test(s)
    || /^\d{4}-\d{2}$/.test(s)
    || (val instanceof Date);
}

function formatMonth(val) {
  if (val instanceof Date) {
    const names = ['Янв','Фев','Мар','Апр','Май','Июн','Июл','Авг','Сен','Окт','Ноя','Дек'];
    return names[val.getMonth()] + ' ' + val.getFullYear();
  }
  const s = String(val).trim();
  const MAP = {'янв':'Янв','фев':'Фев','мар':'Мар','апр':'Апр','май':'Май','июн':'Июн',
                'июл':'Июл','авг':'Авг','сен':'Сен','окт':'Окт','ноя':'Ноя','дек':'Дек'};
  const lo = s.toLowerCase();
  for (const [prefix, label] of Object.entries(MAP)) {
    if (lo.startsWith(prefix)) {
      const numMatch = s.match(/\d+/);
      if (numMatch) { let yr = parseInt(numMatch[0]); if (yr < 100) yr += 2000; return label + ' ' + yr; }
      return label;
    }
  }
  return s;
}

function toNum(v) {
  if (typeof v === 'number') return v;
  if (typeof v === 'string') { const n = parseFloat(v.replace(/[, ]/g,'').replace('%','')); return isNaN(n) ? 0 : n; }
  return 0;
}

function getAnalysisMonth() {
  try {
    const ss = SpreadsheetApp.openById(DASHBOARD_CONFIG_ID);
    const sheet = ss.getSheetByName('Допущения') || ss.getSheetByName('Assumptions');
    if (!sheet) return new Date().getMonth() + 1;
    const data = getSheetData(sheet);
    const row = data.find(r => { const k = String(r.key || r.parameter || r['параметр'] || r['ключ'] || '').toLowerCase(); return k.includes('analysis_month') || k.includes('месяц'); });
    return row ? parseInt(row.value || row['значение'] || row.val) : new Date().getMonth() + 1;
  } catch(e) { return new Date().getMonth() + 1; }
}

function getAnalysisYear() {
  try {
    const ss = SpreadsheetApp.openById(DASHBOARD_CONFIG_ID);
    const sheet = ss.getSheetByName('Допущения') || ss.getSheetByName('Assumptions');
    if (!sheet) return new Date().getFullYear();
    const data = getSheetData(sheet);
    const row = data.find(r => { const k = String(r.key || r.parameter || r['параметр'] || r['ключ'] || '').toLowerCase(); return k.includes('analysis_year') || k.includes('год'); });
    return row ? parseInt(row.value || row['значение'] || row.val) : new Date().getFullYear();
  } catch(e) { return new Date().getFullYear(); }
}

function getMonthName(m) {
  const names = ['','Январь','Февраль','Март','Апрель','Май','Июнь','Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'];
  return names[m] || String(m);
}

function getThresholdRules(ss) { try { const s = ss.getSheetByName('Threshold Rules'); return s ? getSheetData(s) : []; } catch(e) { return []; } }
function getAssumptions(ss) { try { const s = ss.getSheetByName('Допущения') || ss.getSheetByName('Assumptions'); return s ? getSheetData(s) : []; } catch(e) { return []; } }
function getNotes(ss) { try { const s = ss.getSheetByName('Notes'); return s ? getSheetData(s) : []; } catch(e) { return []; } }

function getDashboardNotes(ss) {
  try {
    const sheet = ss.getSheetByName('Dashboard Notes');
    if (!sheet) return {};
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return {};
    const headers = data[0].map(h => String(h).toLowerCase());
    const keyCol = headers.indexOf('key'), txtCol = headers.indexOf('text'), datCol = headers.indexOf('date');
    if (keyCol < 0) return {};
    const result = {};
    for (let i = 1; i < data.length; i++) {
      const k = String(data[i][keyCol] || '').trim();
      if (!k) continue;
      result[k] = { text: String(data[i][txtCol] || ''), date: String(data[i][datCol] || '') };
    }
    return result;
  } catch(e) { return {}; }
}

function getAiAnalysis(ss) {
  try {
    const sheet = ss.getSheetByName('AI Analysis');
    if (!sheet) return null;
    const data = getSheetData(sheet);
    return data.sort((a,b) => new Date(b.date||b['дата']) - new Date(a.date||a['дата']))[0] || null;
  } catch(e) { return null; }
}

function getConfig() {
  const ss = SpreadsheetApp.openById(DASHBOARD_CONFIG_ID);
  return { reports: getSheetData(ss.getSheetByName('Reports') || ss.getSheets()[0]), rules: getThresholdRules(ss), assumptions: getAssumptions(ss) };
}
