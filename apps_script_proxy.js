/**
 * Rocket Financial Dashboard — Google Apps Script Data Proxy
 * Версия: 4.0 | 01 июня 2026
 *
 * Принцип работы:
 * 1. Читает лист "Reports" из Dashboard Config (visible=TRUE)
 * 2. Для каждого отчёта читает данные из Google Sheets по sheet_id и sheet_name
 * 3. Трансформирует в единый JSON-формат для дашборда
 * 4. Возвращает через doGet() с CORS-заголовками
 *
 * Деплой:
 * - Расширения → Apps Script → вставить код → Развернуть → Новое развёртывание
 * - Тип: Веб-приложение, Доступ: Все
 * - Скопировать URL и вставить в index.html как DATA_PROXY_URL
 */

// ID Dashboard Config (не меняй)
const CONFIG_SHEET_ID = '1EKR-czK1UvXZDIJUe5MLb70yZXyEMe1VIYTx15UOY3E';

// Шаблонные переменные (из листа Допущения)
const ANALYSIS_MONTH = 4;   // апрель (1-indexed)
const ANALYSIS_YEAR = 2026;

// ============================================================
// ТОЧКА ВХОДА — HTTP GET
// ============================================================
function doGet(e) {
  const params = e ? e.parameter : {};
  const action = params.action || 'get_all';

  let result;
  try {
    if (action === 'get_all') {
      result = getAllData();
    } else if (action === 'add_note') {
      result = { ok: false, error: 'Use POST for notes' };
    } else {
      result = { ok: false, error: 'Unknown action' };
    }
  } catch (err) {
    result = { ok: false, error: err.toString(), stack: err.stack };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// ТОЧКА ВХОДА — HTTP POST (примечания)
// ============================================================
function doPost(e) {
  let result;
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'get_all') {
      result = getAllData();
    } else if (body.action === 'add_note') {
      result = addNote(body.note);
    } else {
      result = { ok: false, error: 'Unknown action' };
    }
  } catch (err) {
    result = { ok: false, error: err.toString() };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// ГЛАВНАЯ ФУНКЦИЯ — СБОРКА ВСЕХ ДАННЫХ
// ============================================================
function getAllData() {
  const config = SpreadsheetApp.openById(CONFIG_SHEET_ID);

  // 1. Читаем конфигурацию отчётов
  const reports = getReports(config);

  // 2. Читаем данные по каждому видимому отчёту
  const rawData = {};
  for (const report of reports) {
    if (report.visible !== 'TRUE' && report.visible !== true) continue;
    try {
      rawData[report.id] = readSheetData(report.sheet_id, report.sheet_name);
    } catch (err) {
      rawData[report.id] = { error: err.toString() };
    }
  }

  // 3. Трансформируем в формат дашборда
  const dashboard = transformData(rawData);

  // 4. Добавляем метаданные
  dashboard.threshold_rules = getThresholdRules(config);
  dashboard.notes = getNotes(config);
  dashboard.analysis = getAIAnalysis(config);
  dashboard.assumptions = getAssumptions(config);
  dashboard.updated_at = new Date().toISOString();
  dashboard.ok = true;

  return dashboard;
}

// ============================================================
// ЧТЕНИЕ КОНФИГУРАЦИИ
// ============================================================
function getReports(config) {
  const sheet = config.getSheetByName('Reports');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  }).filter(r => r.id);
}

function getThresholdRules(config) {
  const sheet = config.getSheetByName('Threshold Rules');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  return data.slice(1)
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    })
    .filter(r => r.active === 'TRUE' || r.active === true);
}

function getNotes(config) {
  const sheet = config.getSheetByName('Notes');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  }).filter(r => r.chart_id || r.text);
}

function getAIAnalysis(config) {
  const sheet = config.getSheetByName('AI Analysis');
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  // Берём последнюю запись
  const headers = data[0];
  const lastRow = data[data.length - 1];
  const obj = {};
  headers.forEach((h, i) => { obj[h] = lastRow[i]; });
  return obj;
}

function getAssumptions(config) {
  const sheet = config.getSheetByName('Допущения');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  return data.slice(1)
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    })
    .filter(r => r.active === 'TRUE' || r.active === true);
}

// ============================================================
// ЧТЕНИЕ ДАННЫХ ИЗ GOOGLE SHEETS
// ============================================================
function readSheetData(sheetUrl, sheetName) {
  // Извлекаем ID таблицы из URL
  const match = sheetUrl.toString().match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (!match) throw new Error('Invalid sheet URL: ' + sheetUrl);
  const sheetId = match[1];

  // Подставляем шаблонные переменные в имя листа
  const resolvedName = resolveSheetName(sheetName.toString());

  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheetByName(resolvedName);
  if (!sheet) {
    // Попробуем найти лист по частичному совпадению
    const sheets = ss.getSheets();
    const found = sheets.find(s => s.getName().includes(resolvedName.split(' ')[0]));
    if (!found) throw new Error('Sheet not found: ' + resolvedName + ' in ' + sheetId);
    return found.getDataRange().getValues();
  }

  return sheet.getDataRange().getValues();
}

function resolveSheetName(name) {
  const monthNames = ['Январь','Февраль','Март','Апрель','Май','Июнь',
                      'Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'];
  const monthName = monthNames[ANALYSIS_MONTH - 1];
  return name
    .replace('{month}', ANALYSIS_MONTH.toString().padStart(2, '0'))
    .replace('{year}', ANALYSIS_YEAR.toString())
    .replace('{month_name}', monthName);
}

// ============================================================
// ТРАНСФОРМАЦИЯ ДАННЫХ В ФОРМАТ ДАШБОРДА
// ============================================================
function transformData(raw) {
  const D = {
    months: [],
    revenue: [],
    var_costs: [],
    margin_pct: [],
    gross_margin_pct: [],
    op_margin_pct: [],
    np_margin_pct: [],
    op_profit: [],
    net_profit: [],
    traffic_cost: [],
    indirect: [],
    interest: [],
    ncf_op: [],
    ncf_inv: [],
    ncf_fin: [],
    dividends: [],
    cash: [],
    assets_total: [],
    equity: [],
    liabilities: [],
    lt_debt: [],
    st_debt: [],
    kz: [],
    current_ratio: [],
    abs_liquidity: [],
    fin_stability: [],
    autonomy: [],
    de_ratio: [],
    roe_monthly: [],
    roa_monthly: [],
    dfl: [],
    rev_core: [],
    rev_streams: [],
    rev_influence: [],
    rev_mannequin: [],
    tr_fb: [],
    tr_tg: [],
    tr_teasers: [],
    tr_influence: [],
    tr_pp: [],
    tr_google: [],
    ftd_old: [],
    ftd_new: [],
    ftd_total: [],
    crm_rev: [],
    crm_pct: [],
    // Дополнительные данные для новых вкладок
    pnl_combo: null,
    pnl_consol: null,
    model_data: null,
    payment_data: null,
    tasks_data: null,
    ckp_data: null,
  };

  // ---- ОПиУ (opiu) ----
  if (raw.opiu && !raw.opiu.error) {
    parseOPIU(raw.opiu, D);
  }

  // ---- ДДС (dds_common) ----
  if (raw.dds_common && !raw.dds_common.error) {
    parseDDS(raw.dds_common, D);
  }

  // ---- Баланс (balance) ----
  if (raw.balance && !raw.balance.error) {
    parseBalance(raw.balance, D);
  }

  // ---- Трафик (traffic) ----
  if (raw.traffic && !raw.traffic.error) {
    parseTraffic(raw.traffic, D);
  }

  // ---- ОПиУ блоки (opiu_combo) ----
  if (raw.opiu_combo && !raw.opiu_combo.error) {
    D.pnl_combo = raw.opiu_combo;
  }

  // ---- Консолидация (opiu_consol) ----
  if (raw.opiu_consol && !raw.opiu_consol.error) {
    D.pnl_consol = raw.opiu_consol;
  }

  // ---- Модель (opiu_model) ----
  if (raw.opiu_model && !raw.opiu_model.error) {
    D.model_data = raw.opiu_model;
    parseModel(raw.opiu_model, D);
  }

  // ---- Платёжный календарь (payment_cal) ----
  if (raw.payment_cal && !raw.payment_cal.error) {
    D.payment_data = raw.payment_cal;
  }

  // ---- ЦКП (roadmap_ckp) ----
  if (raw.roadmap_ckp && !raw.roadmap_ckp.error) {
    D.ckp_data = raw.roadmap_ckp;
  }

  // ---- Задачи (roadmap_tasks) ----
  if (raw.roadmap_tasks && !raw.roadmap_tasks.error) {
    D.tasks_data = raw.roadmap_tasks;
  }

  return D;
}

// ============================================================
// ПАРСЕРЫ ПО ТИПАМ ОТЧЁТОВ
// ============================================================

function parseOPIU(rows, D) {
  // Ищем строки по ключевым словам
  // Структура: первая колонка — метрика, остальные — месяцы
  const monthRow = findRow(rows, ['Месяц', 'месяц', 'Период', 'Jan', 'Янв']);
  if (monthRow) {
    D.months = monthRow.slice(1).filter(v => v !== '' && v !== null);
  }

  const revRow = findRow(rows, ['Выручка', 'Revenue', 'Доход']);
  if (revRow) D.revenue = parseNumbers(revRow.slice(1), D.months.length);

  const vcRow = findRow(rows, ['Переменные расходы', 'Перем.расходы', 'Перем. расходы', 'Variable']);
  if (vcRow) D.var_costs = parseNumbers(vcRow.slice(1), D.months.length);

  const opRow = findRow(rows, ['Операционная прибыль', 'Опер. прибыль', 'EBIT', 'Опер.прибыль']);
  if (opRow) D.op_profit = parseNumbers(opRow.slice(1), D.months.length);

  const npRow = findRow(rows, ['Чистая прибыль', 'ЧП', 'Net profit', 'Прибыль чистая']);
  if (npRow) D.net_profit = parseNumbers(npRow.slice(1), D.months.length);

  const trRow = findRow(rows, ['Трафик', 'Затраты на трафик', 'Traffic']);
  if (trRow) D.traffic_cost = parseNumbers(trRow.slice(1), D.months.length);

  const indRow = findRow(rows, ['Косвенные', 'Косвенные расходы', 'Indirect', 'Накладные']);
  if (indRow) D.indirect = parseNumbers(indRow.slice(1), D.months.length);

  const intRow = findRow(rows, ['Проценты', 'Процентные расходы', 'Interest', '% кредит']);
  if (intRow) D.interest = parseNumbers(intRow.slice(1), D.months.length);

  // Рентабельности
  const n = D.months.length;
  for (let i = 0; i < n; i++) {
    const rev = D.revenue[i] || 0;
    const vc = D.var_costs[i] || 0;
    const op = D.op_profit[i] || 0;
    const np = D.net_profit[i] || 0;
    if (rev > 0) {
      D.margin_pct.push(round2((rev - vc) / rev * 100));
      D.gross_margin_pct.push(round2((rev - vc) / rev * 100));
      D.op_margin_pct.push(round2(op / rev * 100));
      D.np_margin_pct.push(round2(np / rev * 100));
    } else {
      D.margin_pct.push(0);
      D.gross_margin_pct.push(0);
      D.op_margin_pct.push(0);
      D.np_margin_pct.push(0);
    }
  }

  // Выручка по направлениям
  const coreRow = findRow(rows, ['Core', 'Ядро', 'Основное']);
  if (coreRow) D.rev_core = parseNumbers(coreRow.slice(1), n);

  const streamsRow = findRow(rows, ['Streams', 'Стримы', 'Rocket Streams']);
  if (streamsRow) D.rev_streams = parseNumbers(streamsRow.slice(1), n);

  const influenceRow = findRow(rows, ['Influence', 'Инфлюенс']);
  if (influenceRow) D.rev_influence = parseNumbers(influenceRow.slice(1), n);

  const mannequinRow = findRow(rows, ['Mannequin', 'Манекен']);
  if (mannequinRow) D.rev_mannequin = parseNumbers(mannequinRow.slice(1), n);
}

function parseDDS(rows, D) {
  const n = D.months.length || 5;

  const ncfOpRow = findRow(rows, ['NCF операционный', 'NCF Операционный', 'Операционный NCF', 'NCF_op']);
  if (ncfOpRow) D.ncf_op = parseNumbers(ncfOpRow.slice(1), n);

  const ncfInvRow = findRow(rows, ['NCF инвестиционный', 'NCF Инвестиционный', 'Инвестиционный NCF']);
  if (ncfInvRow) D.ncf_inv = parseNumbers(ncfInvRow.slice(1), n);

  const ncfFinRow = findRow(rows, ['NCF финансовый', 'NCF Финансовый', 'Финансовый NCF']);
  if (ncfFinRow) D.ncf_fin = parseNumbers(ncfFinRow.slice(1), n);

  const divRow = findRow(rows, ['Дивиденды', 'Dividends', 'Дивид.']);
  if (divRow) D.dividends = parseNumbers(divRow.slice(1), n);

  const cashRow = findRow(rows, ['Остаток ДС', 'Кэш', 'Cash', 'Остаток денежных средств', 'Денежные средства']);
  if (cashRow) D.cash = parseNumbers(cashRow.slice(1), n);
}

function parseBalance(rows, D) {
  const n = D.months.length || 5;

  const assetsRow = findRow(rows, ['Активы всего', 'Итого активы', 'Total assets', 'Баланс', 'Активы итого']);
  if (assetsRow) D.assets_total = parseNumbers(assetsRow.slice(1), n);

  const equityRow = findRow(rows, ['Капитал', 'Собственный капитал', 'Equity', 'СК']);
  if (equityRow) D.equity = parseNumbers(equityRow.slice(1), n);

  const liabRow = findRow(rows, ['Обязательства', 'Liabilities', 'Обяз.', 'Итого обязательства']);
  if (liabRow) D.liabilities = parseNumbers(liabRow.slice(1), n);

  const ltDebtRow = findRow(rows, ['Долгосрочные займы', 'ДЗ долгосрочные', 'LT debt', 'Долгосрочный долг']);
  if (ltDebtRow) D.lt_debt = parseNumbers(ltDebtRow.slice(1), n);

  const stDebtRow = findRow(rows, ['Краткосрочные займы', 'КЗ краткосрочные', 'ST debt', 'Краткосрочный долг']);
  if (stDebtRow) D.st_debt = parseNumbers(stDebtRow.slice(1), n);

  const kzRow = findRow(rows, ['Кредиторская задолженность', 'КЗ', 'Accounts payable']);
  if (kzRow) D.kz = parseNumbers(kzRow.slice(1), n);

  // Коэффициенты
  const crRow = findRow(rows, ['Текущая ликвидность', 'Current ratio', 'Тек. ликвидность']);
  if (crRow) D.current_ratio = parseNumbers(crRow.slice(1), n);

  const alRow = findRow(rows, ['Абсолютная ликвидность', 'Абс. ликвидность', 'Abs liquidity']);
  if (alRow) D.abs_liquidity = parseNumbers(alRow.slice(1), n);

  const fsRow = findRow(rows, ['Финансовая устойчивость', 'Фин. устойчивость', 'Financial stability']);
  if (fsRow) D.fin_stability = parseNumbers(fsRow.slice(1), n);

  const autoRow = findRow(rows, ['Коэффициент автономии', 'Автономия', 'Autonomy', 'Коэф. автономии']);
  if (autoRow) D.autonomy = parseNumbers(autoRow.slice(1), n);

  const deRow = findRow(rows, ['D/E', 'Долг/Капитал', 'Debt to equity', 'D/E ratio']);
  if (deRow) D.de_ratio = parseNumbers(deRow.slice(1), n);

  const dflRow = findRow(rows, ['DFL', 'Финансовый рычаг', 'Эффект фин. рычага']);
  if (dflRow) D.dfl = parseNumbers(dflRow.slice(1), n);

  const roeRow = findRow(rows, ['ROE', 'Рентабельность капитала']);
  if (roeRow) D.roe_monthly = parseNumbers(roeRow.slice(1), n);

  const roaRow = findRow(rows, ['ROA', 'Рентабельность активов']);
  if (roaRow) D.roa_monthly = parseNumbers(roaRow.slice(1), n);
}

function parseTraffic(rows, D) {
  const n = D.months.length || 5;

  const fbRow = findRow(rows, ['Facebook', 'FB', 'Фейсбук']);
  if (fbRow) D.tr_fb = parseNumbers(fbRow.slice(1), n);

  const tgRow = findRow(rows, ['Telegram', 'TG', 'Телеграм', 'TG Ads']);
  if (tgRow) D.tr_tg = parseNumbers(tgRow.slice(1), n);

  const teasRow = findRow(rows, ['Тизерки', 'Тизерные сети', 'Teasers']);
  if (teasRow) D.tr_teasers = parseNumbers(teasRow.slice(1), n);

  const infRow = findRow(rows, ['Influence', 'Инфлюенсеры', 'Блогеры']);
  if (infRow) D.tr_influence = parseNumbers(infRow.slice(1), n);

  const ppRow = findRow(rows, ['ПП', 'Партнёрские программы', 'PP', 'Affiliate']);
  if (ppRow) D.tr_pp = parseNumbers(ppRow.slice(1), n);

  const gRow = findRow(rows, ['Google', 'Google Ads']);
  if (gRow) D.tr_google = parseNumbers(gRow.slice(1), n);
}

function parseModel(rows, D) {
  const n = D.months.length || 5;

  const ftdOldRow = findRow(rows, ['FTD старых', 'FTD old', 'FTD_old', 'Старые FTD']);
  if (ftdOldRow) D.ftd_old = parseNumbers(ftdOldRow.slice(1), n);

  const ftdNewRow = findRow(rows, ['FTD новых', 'FTD new', 'FTD_new', 'Новые FTD']);
  if (ftdNewRow) D.ftd_new = parseNumbers(ftdNewRow.slice(1), n);

  const ftdTotalRow = findRow(rows, ['FTD всего', 'FTD total', 'FTD итого', 'Всего FTD']);
  if (ftdTotalRow) D.ftd_total = parseNumbers(ftdTotalRow.slice(1), n);

  const crmRevRow = findRow(rows, ['CRM выручка', 'CRM Revenue', 'Выручка CRM']);
  if (crmRevRow) D.crm_rev = parseNumbers(crmRevRow.slice(1), n);

  const crmPctRow = findRow(rows, ['CRM %', 'CRM доля', 'CRM-доля', 'Доля CRM']);
  if (crmPctRow) D.crm_pct = parseNumbers(crmPctRow.slice(1), n);
}

// ============================================================
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// ============================================================

/**
 * Ищет строку в массиве данных по ключевым словам в первой колонке
 */
function findRow(rows, keywords) {
  for (const row of rows) {
    const cell = (row[0] || '').toString().trim();
    for (const kw of keywords) {
      if (cell.toLowerCase().includes(kw.toLowerCase())) {
        return row;
      }
    }
  }
  return null;
}

/**
 * Парсит числа из строки, возвращает массив нужной длины
 */
function parseNumbers(arr, maxLen) {
  const result = [];
  for (let i = 0; i < arr.length && result.length < maxLen; i++) {
    const v = arr[i];
    if (v === '' || v === null || v === undefined) continue;
    const n = parseFloat(v.toString().replace(/[^\d.-]/g, ''));
    if (!isNaN(n)) result.push(n);
  }
  // Дополняем нулями если не хватает
  while (result.length < maxLen) result.push(0);
  return result;
}

function round2(n) {
  return Math.round(n * 100) / 100;
}

// ============================================================
// СОХРАНЕНИЕ ПРИМЕЧАНИЙ
// ============================================================
function addNote(note) {
  try {
    const config = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = config.getSheetByName('Notes');
    if (!sheet) return { ok: false, error: 'Notes sheet not found' };

    sheet.appendRow([
      note.report_id || '',
      note.month || new Date().toISOString().slice(0, 7),
      note.chart_id || '',
      note.text || '',
      new Date().toISOString()
    ]);
    return { ok: true };
  } catch (err) {
    return { ok: false, error: err.toString() };
  }
}

// ============================================================
// ТЕСТОВАЯ ФУНКЦИЯ (запускать вручную для отладки)
// ============================================================
function testProxy() {
  const result = getAllData();
  Logger.log(JSON.stringify(result).slice(0, 2000));
  return result;
}
