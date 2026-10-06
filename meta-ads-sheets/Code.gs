/**
 * Meta (Facebook) Ads → Google Таблица
 *
 * Забирает статистику рекламного кабинета через Meta Marketing API (Insights)
 * и складывает в лист «Data»: расход, показы, охват, клики, лиды, CPL, CTR, CPC, CPM
 * по дням и объявлениям. Лист «Сводка» считает итоги по кампаниям за выбранный период.
 *
 * Установка и настройка: README.md
 */

const API_VERSION = 'v23.0';
const GRAPH_URL = 'https://graph.facebook.com/';
const TOKEN_PROPERTY = 'META_TOKEN';
const SCHEDULED_HANDLER = 'refreshDataScheduled';

const SHEET_SETTINGS = 'Настройки';
const SHEET_SUMMARY = 'Сводка';
const SHEET_DATA = 'Data';

const CHUNK_DAYS = 14; // период запрашиваем кусками, иначе Meta отказывает на больших объёмах
const PAGE_LIMIT = 500;
const MAX_RETRIES = 3;

// Колонки листа Data: [ключ, заголовок, формат]. Текстовый формат '@' нужен для ID:
// иначе Sheets превратит 18-значный ID в число и испортит последние цифры.
const COLUMNS = [
  ['date', 'Дата', 'yyyy-mm-dd'],
  ['account_id', 'ID аккаунта', '@'],
  ['account_name', 'Аккаунт', '@'],
  ['currency', 'Валюта', '@'],
  ['campaign_id', 'ID кампании', '@'],
  ['campaign_name', 'Кампания', '@'],
  ['adset_id', 'ID группы', '@'],
  ['adset_name', 'Группа объявлений', '@'],
  ['ad_id', 'ID объявления', '@'],
  ['ad_name', 'Объявление', '@'],
  ['spend', 'Расход', '#,##0.00'],
  ['impressions', 'Показы', '#,##0'],
  ['reach', 'Охват', '#,##0'],
  ['clicks', 'Клики (все)', '#,##0'],
  ['link_clicks', 'Клики по ссылке', '#,##0'],
  ['leads', 'Лиды', '#,##0'],
  ['cpl', 'CPL (цена лида)', '#,##0.00'],
  ['ctr', 'CTR по ссылке', '0.00%'],
  ['cpc', 'CPC по ссылке', '#,##0.00'],
  ['cpm', 'CPM', '#,##0.00']
];
const COL = {};
COLUMNS.forEach(function (c, i) { COL[c[0]] = i; });

const SETTINGS = [
  { key: 'accounts', label: 'Рекламные аккаунты', value: '', text: true,
    hint: 'ID кабинета, лучше с префиксом: act_1234567890. Несколько — через запятую.' },
  { key: 'leadTypes', label: 'Что считать лидом (action_type)', value: 'lead', text: true,
    hint: 'lead — все лиды; onsite_conversion.lead_grouped — формы Meta; offsite_conversion.fb_pixel_lead — сайт (пиксель); ' +
      'onsite_conversion.messaging_conversation_started_7d — переписки. Указывайте один тип, чтобы не удвоить лиды.' },
  { key: 'refreshDays', label: 'Перезаписывать последние N дней', value: 7,
    hint: 'Meta дописывает конверсии задним числом, поэтому последние дни при каждом обновлении берутся заново.' },
  { key: 'initialDays', label: 'Первая загрузка: дней истории', value: 90,
    hint: 'Сколько дней загрузить при самом первом запуске.' },
  { key: 'includeToday', label: 'Включать сегодняшний день', value: 'Нет',
    hint: 'Да / Нет. Сегодня данные неполные. Если часовой пояс таблицы впереди часового пояса кабинета, Meta может отказать.' }
];
const SETTINGS_FIRST_ROW = 2;
const STATUS_LABELS = ['Последнее обновление', 'Статус', 'Период данных'];
const STATUS_FIRST_ROW = SETTINGS_FIRST_ROW + SETTINGS.length + 1;

const SUMMARY_TOTAL_ROW = 6;
const SUMMARY_FIRST_ROW = 7;
const SUMMARY_KEY_COLUMN = 10; // J: служебная колонка с ключом для SUMIFS (скрыта)

// ───────────────────────────── меню ─────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Meta Ads')
    .addItem('1. Подготовить таблицу', 'setupSheets')
    .addItem('2. Ввести токен Meta', 'promptToken')
    .addItem('3. Проверить подключение', 'testConnection')
    .addSeparator()
    .addItem('Обновить данные', 'refreshData')
    .addItem('Загрузить историю с даты…', 'promptBackfill')
    .addSeparator()
    .addItem('Включить ежедневное обновление', 'enableDailyTrigger')
    .addItem('Выключить ежедневное обновление', 'disableDailyTrigger')
    .addToUi();
}

function refreshData() {
  runRefresh({ interactive: true });
}

/** Обработчик триггера: при ошибке бросает исключение, чтобы Google прислал письмо о сбое. */
function refreshDataScheduled() {
  runRefresh({ interactive: false });
}

function promptBackfill() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Загрузить историю',
    'С какой даты загрузить данные? Формат ГГГГ-ММ-ДД, например 2026-01-01.\n' +
    'Данные за этот период будут перезаписаны свежими из Meta.', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const since = resp.getResponseText().trim();
  if (!isValidIsoDate(since)) {
    ui.alert('Не похоже на дату. Нужен формат ГГГГ-ММ-ДД, например 2026-01-01.');
    return;
  }
  runRefresh({ interactive: true, since: since });
}

function promptToken() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Токен Meta',
    'Вставьте токен доступа (нужно разрешение ads_read). Он сохранится в настройках скрипта, а не в таблице.',
    ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const token = resp.getResponseText().trim();
  if (!token) {
    ui.alert('Токен пустой, ничего не сохранено.');
    return;
  }
  PropertiesService.getScriptProperties().setProperty(TOKEN_PROPERTY, token);
  ui.alert('Токен сохранён. Теперь выберите «3. Проверить подключение».');
}

function testConnection() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  try {
    const settings = readSettings(ss);
    const token = getToken();
    const lines = settings.accounts.map(function (id) {
      try {
        const info = getAccountInfo(id, token);
        return '✔ act_' + id + ': ' + info.name + ', валюта ' + info.currency + ', часовой пояс ' + info.timezone_name;
      } catch (e) {
        return '✘ act_' + id + ': ' + e.message;
      }
    });
    ui.alert('Проверка подключения:\n\n' + lines.join('\n') +
      '\n\nЧасовой пояс таблицы: ' + ss.getSpreadsheetTimeZone());
  } catch (e) {
    ui.alert('Не получилось: ' + e.message);
  }
}

function enableDailyTrigger() {
  removeScheduledTriggers();
  ScriptApp.newTrigger(SCHEDULED_HANDLER).timeBased().everyDays(1).atHour(6).create();
  notify('Ежедневное обновление включено: каждый день около 6:00 (часовой пояс проекта скрипта).');
}

function disableDailyTrigger() {
  removeScheduledTriggers();
  notify('Ежедневное обновление выключено.');
}

function removeScheduledTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === SCHEDULED_HANDLER) ScriptApp.deleteTrigger(t);
  });
}

// ───────────────────────── подготовка листов ─────────────────────────

function setupSheets() {
  const ss = SpreadsheetApp.getActive();
  const settings = ss.getSheetByName(SHEET_SETTINGS) || ss.insertSheet(SHEET_SETTINGS, 0);
  const summary = ss.getSheetByName(SHEET_SUMMARY) || ss.insertSheet(SHEET_SUMMARY, 1);
  const data = ss.getSheetByName(SHEET_DATA) || ss.insertSheet(SHEET_DATA, 2);
  setupSettingsSheet(settings);
  setupSummarySheet(summary);
  setupDataSheet(data);
  removeEmptyExtraSheets(ss);
  updateSummary(ss);
  notify('Таблица подготовлена. Заполните лист «Настройки», затем введите токен (меню Meta Ads).');
}

function setupSettingsSheet(sheet) {
  sheet.getRange(1, 1, 1, 3).setValues([['Параметр', 'Значение', 'Подсказка']]).setFontWeight('bold');
  SETTINGS.forEach(function (s, i) {
    const row = SETTINGS_FIRST_ROW + i;
    sheet.getRange(row, 1).setValue(s.label);
    const cell = sheet.getRange(row, 2);
    if (s.text) cell.setNumberFormat('@');
    if (cell.getValue() === '') cell.setValue(s.value);
    sheet.getRange(row, 3).setValue(s.hint);
  });
  const yesNo = SpreadsheetApp.newDataValidation().requireValueInList(['Да', 'Нет'], true).build();
  sheet.getRange(SETTINGS_FIRST_ROW + SETTINGS.length - 1, 2).setDataValidation(yesNo);
  STATUS_LABELS.forEach(function (label, i) {
    sheet.getRange(STATUS_FIRST_ROW + i, 1).setValue(label).setFontWeight('bold');
  });
  sheet.setColumnWidth(1, 260);
  sheet.setColumnWidth(2, 220);
}

function setupSummarySheet(sheet) {
  sheet.getRange(1, 1).setValue('Сводка по кампаниям').setFontWeight('bold');
  sheet.getRange(2, 1).setValue('С даты');
  sheet.getRange(3, 1).setValue('По дату');
  const from = sheet.getRange(2, 2);
  const to = sheet.getRange(3, 2);
  if (from.getValue() === '') from.setFormula('=TODAY()-29');
  if (to.getValue() === '') to.setFormula('=TODAY()');
  from.setNumberFormat('yyyy-mm-dd');
  to.setNumberFormat('yyyy-mm-dd');
  const headers = ['Кампания', 'Расход', 'Показы', 'Клики по ссылке', 'Лиды', 'CPL', 'CTR', 'CPC', 'CPM', 'ключ (служебный)'];
  sheet.getRange(SUMMARY_TOTAL_ROW - 1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.getRange(SUMMARY_TOTAL_ROW, 1).setValue('ИТОГО').setFontWeight('bold');
  sheet.getRange(SUMMARY_TOTAL_ROW, 2, 1, 8).setFormulas([summaryRowFormulas(SUMMARY_TOTAL_ROW, false)]).setFontWeight('bold');
  applySummaryFormats(sheet, SUMMARY_TOTAL_ROW, 1);
  sheet.setFrozenRows(SUMMARY_TOTAL_ROW);
  sheet.setColumnWidth(1, 320);
  sheet.hideColumns(SUMMARY_KEY_COLUMN);
}

function setupDataSheet(sheet) {
  const headers = COLUMNS.map(function (c) { return c[1]; });
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function removeEmptyExtraSheets(ss) {
  const ours = [SHEET_SETTINGS, SHEET_SUMMARY, SHEET_DATA];
  ss.getSheets().forEach(function (sh) {
    if (ours.indexOf(sh.getName()) === -1 && sh.getLastRow() === 0 && sh.getLastColumn() === 0 && ss.getSheets().length > ours.length) {
      ss.deleteSheet(sh);
    }
  });
}

// ───────────────────────────── обновление ─────────────────────────────

function runRefresh(opts) {
  const interactive = !!(opts && opts.interactive);
  const ss = SpreadsheetApp.getActive();
  const lock = LockService.getScriptLock();
  let result = { period: '', rows: 0, errors: [] };

  if (!lock.tryLock(30000)) {
    result.errors.push('Обновление уже выполняется, подождите и повторите.');
  } else {
    try {
      if (interactive) ss.toast('Загружаю данные из Meta…', 'Meta Ads', 5);
      result = doRefresh(ss, opts || {});
    } catch (e) {
      result.errors.push(e.message);
      writeStatus(ss, 'Ошибка: ' + e.message, '');
    } finally {
      lock.releaseLock();
    }
  }

  if (result.errors.length) {
    // Запуск по триггеру должен упасть: так сбой виден в журнале выполнений Apps Script.
    if (!interactive) throw new Error(result.errors.join('\n'));
    notify('Обновление не удалось или прошло с ошибками:\n\n' + result.errors.join('\n'));
  } else if (interactive) {
    notify('Готово. Период ' + result.period + ', строк: ' + result.rows + '.');
  }
}

/** Загружает окно из Meta и обновляет листы. Ошибки отдельных аккаунтов собирает в errors. */
function doRefresh(ss, opts) {
  const settings = readSettings(ss);
  const token = getToken();
  const tz = ss.getSpreadsheetTimeZone();
  const dataSheet = ss.getSheetByName(SHEET_DATA);
  if (!dataSheet) throw new Error('Нет листа «Data». Выберите в меню «Подготовить таблицу».');

  const win = computeWindow(settings, todayIso(tz), dataSheet.getLastRow() <= 1, opts.since);
  const fresh = [];
  const okIds = [];
  const errors = [];
  settings.accounts.forEach(function (id) {
    try {
      fetchAccountRows(id, token, win, settings.leadTypes).forEach(function (r) { fresh.push(r); });
      okIds.push(id);
    } catch (e) {
      errors.push('act_' + id + ': ' + e.message);
    }
  });

  // Аккаунт с ошибкой не трогаем: его старые строки в окне остаются как были.
  if (okIds.length) {
    writeData(dataSheet, fresh, { since: win.since, until: win.until, accountIds: okIds }, tz);
    updateSummary(ss);
  }

  const period = win.since + ' … ' + win.until;
  writeStatus(ss, errors.length ? 'Ошибки: ' + errors.join(' | ') : 'OK: строк загружено ' + fresh.length, period);
  return { period: period, rows: fresh.length, errors: errors };
}

function writeStatus(ss, status, period) {
  const sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!sheet) return;
  const stamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm');
  sheet.getRange(STATUS_FIRST_ROW, 2, 3, 1).setValues([[stamp], [status], [period]]);
}

function notify(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    console.log(message); // запуск по триггеру: интерфейса нет
  }
}

function getToken() {
  const token = PropertiesService.getScriptProperties().getProperty(TOKEN_PROPERTY);
  if (!token) throw new Error('Не задан токен Meta. Выберите в меню «Ввести токен Meta».');
  return token;
}

// ───────────────────────────── настройки ─────────────────────────────

function readSettings(ss) {
  const sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!sheet) throw new Error('Нет листа «Настройки». Выберите в меню «Подготовить таблицу».');
  const values = sheet.getRange(SETTINGS_FIRST_ROW, 2, SETTINGS.length, 1).getValues();
  const raw = {};
  SETTINGS.forEach(function (s, i) { raw[s.key] = values[i][0]; });
  return parseSettings(raw);
}

function parseSettings(raw) {
  const seen = {};
  const accounts = [];
  String(raw.accounts == null ? '' : raw.accounts).split(/[,;\s]+/).forEach(function (part) {
    const id = part.replace(/^act_/i, '');
    if (!id) return;
    if (!/^\d+$/.test(id)) throw new Error('Неверный ID аккаунта «' + part + '». Нужны цифры, например act_1234567890.');
    if (!seen[id]) { seen[id] = true; accounts.push(id); }
  });
  if (!accounts.length) throw new Error('На листе «Настройки» не указан рекламный аккаунт.');

  const leadTypes = String(raw.leadTypes == null ? '' : raw.leadTypes).split(/[,;\s]+/).filter(Boolean);
  return {
    accounts: accounts,
    leadTypes: leadTypes.length ? leadTypes : ['lead'],
    refreshDays: clampInt(raw.refreshDays, 7, 1, 90),
    initialDays: clampInt(raw.initialDays, 90, 1, 1095),
    includeToday: /^(да|yes|true|1)$/i.test(String(raw.includeToday == null ? '' : raw.includeToday).trim())
  };
}

function clampInt(value, fallback, min, max) {
  const n = Math.floor(Number(value));
  if (!isFinite(n) || n < min) return fallback;
  return Math.min(n, max);
}

/** Окно загрузки. since задан вручную при «Загрузить историю». */
function computeWindow(settings, today, isFirstLoad, forcedSince) {
  const until = settings.includeToday ? today : addDays(today, -1);
  const days = isFirstLoad ? settings.initialDays : settings.refreshDays;
  const since = forcedSince || addDays(until, -(days - 1));
  if (since > until) throw new Error('Начальная дата ' + since + ' позже конечной ' + until + '.');
  return { since: since, until: until };
}

// ───────────────────────────── Meta API ─────────────────────────────

function fetchAccountRows(accountId, token, win, leadTypes) {
  const rows = [];
  buildChunks(win.since, win.until, CHUNK_DAYS).forEach(function (chunk) {
    let url = buildInsightsUrl(accountId, token, chunk.since, chunk.until);
    while (url) {
      const body = fetchJson(url, token);
      (body.data || []).forEach(function (item) { rows.push(toRow(item, leadTypes)); });
      const next = body.paging && body.paging.next;
      if (next && next.indexOf(GRAPH_URL) !== 0) throw new Error('Meta вернула неожиданную ссылку для следующей страницы.');
      url = next || null;
    }
  });
  return rows;
}

function insightsFields() {
  return ['date_start', 'account_id', 'account_name', 'account_currency',
    'campaign_id', 'campaign_name', 'adset_id', 'adset_name', 'ad_id', 'ad_name',
    'spend', 'impressions', 'reach', 'clicks', 'inline_link_clicks', 'actions'];
}

function buildInsightsUrl(accountId, token, since, until) {
  return GRAPH_URL + API_VERSION + '/act_' + accountId + '/insights?' + toQuery({
    level: 'ad',
    fields: insightsFields().join(','),
    time_increment: '1',
    time_range: JSON.stringify({ since: since, until: until }),
    limit: String(PAGE_LIMIT),
    access_token: token
  });
}

function getAccountInfo(accountId, token) {
  return fetchJson(GRAPH_URL + API_VERSION + '/act_' + accountId + '?' + toQuery({
    fields: 'name,currency,timezone_name',
    access_token: token
  }), token);
}

function toQuery(params) {
  return Object.keys(params).map(function (k) {
    return encodeURIComponent(k) + '=' + encodeURIComponent(params[k]);
  }).join('&');
}

function fetchJson(url, token) {
  let lastError;
  for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const status = resp.getResponseCode();
    let body = null;
    try { body = JSON.parse(resp.getContentText()); } catch (e) { body = null; }
    if (status >= 200 && status < 300 && body && !body.error) return body;

    const err = body && body.error;
    lastError = new Error(redact(explainMetaError(err, status), token));
    const retryable = status >= 500 || (err && (isRateLimit(err.code) || err.is_transient === true));
    if (!retryable || attempt === MAX_RETRIES) break;
    Utilities.sleep(Math.pow(2, attempt + 1) * 1000);
  }
  throw lastError;
}

function isRateLimit(code) {
  return code === 4 || code === 17 || code === 32 || code === 613 || (code >= 80000 && code <= 80014);
}

function explainMetaError(err, httpStatus) {
  if (!err) return 'Неожиданный ответ Meta (HTTP ' + httpStatus + ').';
  const code = err.code;
  const msg = err.message || '';
  if (code === 190) return 'Токен недействителен или истёк (код 190). Получите новый и введите через меню. ' + msg;
  if (code === 10 || (code >= 200 && code <= 299)) {
    return 'Нет доступа (код ' + code + '). У токена нужно разрешение ads_read и доступ к этому рекламному аккаунту. ' + msg;
  }
  if (code === 100) return 'Meta отклонила запрос (код 100): ' + msg + ' Проверьте ID аккаунта и права токена.';
  if (isRateLimit(code)) return 'Превышен лимит запросов Meta (код ' + code + '). Повторите позже. ' + msg;
  return 'Ошибка Meta API (код ' + code + (err.error_subcode ? '/' + err.error_subcode : '') + '): ' + msg;
}

function redact(text, token) {
  return token ? String(text).split(token).join('***') : String(text);
}

/** Одна строка Insights → строка листа Data (дата в ISO, остальное по порядку COLUMNS). */
function toRow(item, leadTypes) {
  const spend = num(item.spend);
  const impressions = num(item.impressions);
  const linkClicks = num(item.inline_link_clicks);
  const leads = countActions(item.actions, leadTypes);
  const row = new Array(COLUMNS.length);
  row[COL.date] = item.date_start;
  row[COL.account_id] = text(item.account_id);
  row[COL.account_name] = text(item.account_name);
  row[COL.currency] = text(item.account_currency);
  row[COL.campaign_id] = text(item.campaign_id);
  row[COL.campaign_name] = text(item.campaign_name);
  row[COL.adset_id] = text(item.adset_id);
  row[COL.adset_name] = text(item.adset_name);
  row[COL.ad_id] = text(item.ad_id);
  row[COL.ad_name] = text(item.ad_name);
  row[COL.spend] = spend;
  row[COL.impressions] = impressions;
  row[COL.reach] = num(item.reach);
  row[COL.clicks] = num(item.clicks);
  row[COL.link_clicks] = linkClicks;
  row[COL.leads] = leads;
  row[COL.cpl] = leads > 0 ? round(spend / leads, 4) : '';
  row[COL.ctr] = impressions > 0 ? round(linkClicks / impressions, 6) : '';
  row[COL.cpc] = linkClicks > 0 ? round(spend / linkClicks, 4) : '';
  row[COL.cpm] = impressions > 0 ? round(spend / impressions * 1000, 4) : '';
  return row;
}

function countActions(actions, types) {
  let total = 0;
  (actions || []).forEach(function (a) {
    if (types.indexOf(a.action_type) !== -1) total += num(a.value);
  });
  return total;
}

function num(v) {
  const n = Number(v);
  return v === undefined || v === null || v === '' || !isFinite(n) ? 0 : n;
}

function text(v) {
  return v === undefined || v === null ? '' : String(v);
}

function round(x, digits) {
  const k = Math.pow(10, digits);
  return Math.round(x * k) / k;
}

// ───────────────────────────── даты ─────────────────────────────

function isoToUtc(iso) {
  const p = iso.split('-');
  return Date.UTC(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
}

function utcToIso(ms) {
  const d = new Date(ms);
  return d.getUTCFullYear() + '-' + pad2(d.getUTCMonth() + 1) + '-' + pad2(d.getUTCDate());
}

function pad2(n) {
  return n < 10 ? '0' + n : String(n);
}

function addDays(iso, n) {
  return utcToIso(isoToUtc(iso) + n * 86400000);
}

function isValidIsoDate(s) {
  return /^\d{4}-\d{2}-\d{2}$/.test(s) && utcToIso(isoToUtc(s)) === s;
}

function todayIso(tz) {
  return Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
}

/** Делит [since, until] (включительно) на куски по days дней. */
function buildChunks(since, until, days) {
  const chunks = [];
  let cur = since;
  while (cur <= until) {
    let end = addDays(cur, days - 1);
    if (end > until) end = until;
    chunks.push({ since: cur, until: end });
    cur = addDays(end, 1);
  }
  return chunks;
}

function isDate(v) {
  return Object.prototype.toString.call(v) === '[object Date]';
}

function cellToIso(v, tz) {
  if (isDate(v)) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  const s = v === undefined || v === null ? '' : String(v);
  return /^\d{4}-\d{2}-\d{2}/.test(s) ? s.slice(0, 10) : '';
}

// ───────────────────────────── запись в Data ─────────────────────────────

/**
 * Заменяет в листе строки окна [since, until] по успешно загруженным аккаунтам свежими.
 * Лист отсортирован по дате, поэтому всё, что новее since, лежит одним блоком внизу:
 * читаем и перезаписываем только его, а не всю историю.
 */
function writeData(sheet, fresh, win, tz) {
  const ncol = COLUMNS.length;
  const nData = Math.max(0, sheet.getLastRow() - 1);
  const dates = nData ? sheet.getRange(2, 1, nData, 1).getValues() : [];
  // Идём снизу: проверяем только хвост окна, а не всю историю (formatDate на каждую строку дорог).
  let first = dates.length;
  while (first > 0 && cellToIso(dates[first - 1][0], tz) >= win.since) first--;

  const startRow = 2 + first;
  const tailLength = nData - first;
  const tail = tailLength
    ? sheet.getRange(startRow, 1, tailLength, ncol).getValues().map(function (r) { return normalizeRow(r, tz); })
    : [];
  const merged = mergeRows(tail, fresh, win);

  if (merged.length) {
    const lastNeeded = startRow + merged.length - 1;
    const maxRows = sheet.getMaxRows();
    if (lastNeeded > maxRows) sheet.insertRowsAfter(maxRows, lastNeeded - maxRows);
    applyFormats(sheet, startRow, merged.length);
    sheet.getRange(startRow, 1, merged.length, ncol).setValues(merged.map(function (r) { return toSheetRow(r, tz); }));
  }
  // Сначала запись, потом очистка хвоста: при сбое между шагами данные не пропадут.
  const leftover = tail.length - merged.length;
  if (leftover > 0) sheet.getRange(startRow + merged.length, 1, leftover, ncol).clearContent();
}

function mergeRows(tail, fresh, win) {
  const refreshed = {};
  win.accountIds.forEach(function (id) { refreshed[id] = true; });
  const kept = tail.filter(function (r) {
    const inWindow = r[COL.date] >= win.since && r[COL.date] <= win.until;
    return !(inWindow && refreshed[r[COL.account_id]]);
  });
  return kept.concat(fresh).sort(compareRows);
}

const SORT_KEYS = ['date', 'account_id', 'campaign_id', 'adset_id', 'ad_id'];

function compareRows(a, b) {
  for (let i = 0; i < SORT_KEYS.length; i++) {
    const k = COL[SORT_KEYS[i]];
    if (a[k] < b[k]) return -1;
    if (a[k] > b[k]) return 1;
  }
  return 0;
}

/** Строка из листа → каноничный вид (дата ISO, текстовые поля строками). */
function normalizeRow(r, tz) {
  const row = r.slice();
  row[COL.date] = cellToIso(r[COL.date], tz);
  COLUMNS.forEach(function (c, i) {
    if (c[2] === '@') row[i] = text(r[i]);
  });
  return row;
}

function toSheetRow(r, tz) {
  const row = r.slice();
  row[COL.date] = Utilities.parseDate(r[COL.date], tz, 'yyyy-MM-dd');
  return row;
}

function applyFormats(sheet, startRow, count) {
  COLUMNS.forEach(function (c, i) {
    sheet.getRange(startRow, i + 1, count, 1).setNumberFormat(c[2]);
  });
}

// ───────────────────────────── Сводка ─────────────────────────────

function colLetter(index) {
  let n = index + 1;
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function dataRange(key) {
  const c = colLetter(COL[key]);
  return SHEET_DATA + '!$' + c + ':$' + c;
}

function sumifs(key, row, byCampaign) {
  let f = 'SUMIFS(' + dataRange(key);
  if (byCampaign) f += ',' + dataRange('campaign_name') + ',"="&$' + colLetter(SUMMARY_KEY_COLUMN - 1) + row;
  f += ',' + dataRange('date') + ',">="&$B$2,' + dataRange('date') + ',"<="&$B$3)';
  return '=' + f;
}

/** Формулы B..I строки сводки: расход, показы, клики по ссылке, лиды, CPL, CTR, CPC, CPM. */
function summaryRowFormulas(row, byCampaign) {
  return [
    sumifs('spend', row, byCampaign),
    sumifs('impressions', row, byCampaign),
    sumifs('link_clicks', row, byCampaign),
    sumifs('leads', row, byCampaign),
    '=IF(E' + row + '>0,B' + row + '/E' + row + ',"")',
    '=IF(C' + row + '>0,D' + row + '/C' + row + ',"")',
    '=IF(D' + row + '>0,B' + row + '/D' + row + ',"")',
    '=IF(C' + row + '>0,B' + row + '/C' + row + '*1000,"")'
  ];
}

/** Спецсимволы SUMIFS (~ * ?) в названии кампании нужно экранировать. */
function escapeCriterion(name) {
  return String(name).replace(/[~*?]/g, '~$&');
}

function updateSummary(ss) {
  const data = ss.getSheetByName(SHEET_DATA);
  const sheet = ss.getSheetByName(SHEET_SUMMARY);
  if (!data || !sheet) return;

  const last = data.getLastRow();
  const seen = {};
  if (last > 1) {
    data.getRange(2, COL.campaign_name + 1, last - 1, 1).getValues().forEach(function (r) {
      const name = text(r[0]);
      if (name) seen[name] = true;
    });
  }
  const names = Object.keys(seen).sort();

  const oldRows = sheet.getLastRow() - SUMMARY_FIRST_ROW + 1;
  if (oldRows > 0) sheet.getRange(SUMMARY_FIRST_ROW, 1, oldRows, SUMMARY_KEY_COLUMN).clearContent();
  if (!names.length) return;

  const lastNeeded = SUMMARY_FIRST_ROW + names.length - 1;
  if (lastNeeded > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), lastNeeded - sheet.getMaxRows());

  const n = names.length;
  sheet.getRange(SUMMARY_FIRST_ROW, 1, n, 1).setNumberFormat('@').setValues(names.map(function (x) { return [x]; }));
  sheet.getRange(SUMMARY_FIRST_ROW, SUMMARY_KEY_COLUMN, n, 1).setNumberFormat('@')
    .setValues(names.map(function (x) { return [escapeCriterion(x)]; }));
  sheet.getRange(SUMMARY_FIRST_ROW, 2, n, 8).setFormulas(names.map(function (x, i) {
    return summaryRowFormulas(SUMMARY_FIRST_ROW + i, true);
  }));
  applySummaryFormats(sheet, SUMMARY_TOTAL_ROW, n + 1);
}

function applySummaryFormats(sheet, startRow, count) {
  const formats = ['#,##0.00', '#,##0', '#,##0', '#,##0', '#,##0.00', '0.00%', '#,##0.00', '#,##0.00'];
  formats.forEach(function (f, i) {
    sheet.getRange(startRow, 2 + i, count, 1).setNumberFormat(f);
  });
}
