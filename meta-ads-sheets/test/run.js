/**
 * Тесты Code.gs без Google: код грузится в изолированный контекст, а сервисы Google
 * (SpreadsheetApp, UrlFetchApp, ...) заменены имитациями.
 *
 * Имитация таблицы повторяет места, где настоящие Sheets ломают наивный код:
 *  - getRange за пределами листа и setValues неверного размера бросают ошибку;
 *  - строка из цифр в ячейке без текстового формата превращается в число
 *    (так портятся 18-значные ID), строка вида 2026-10-05 превращается в дату;
 *  - getValues возвращает объекты Date для дат.
 *
 * Запуск: node meta-ads-sheets/test/run.js
 */
const vm = require('vm');
const fs = require('fs');
const path = require('path');
const assert = require('assert');

const CODE = fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8');
const plain = (x) => JSON.parse(JSON.stringify(x)); // объекты из контекста vm → обычные

// ───────────── имитация таблицы ─────────────
class FakeSheet {
  constructor(name, maxRows = 1000, maxCols = 26) {
    this.name = name; this.maxRows = maxRows; this.maxCols = maxCols;
    this.values = new Map(); this.formats = new Map(); this.hiddenCols = []; this.frozen = 0;
  }
  key(r, c) { return r + ',' + c; }
  getName() { return this.name; }
  getMaxRows() { return this.maxRows; }
  getMaxColumns() { return this.maxCols; }
  insertRowsAfter(after, n) { assert(after <= this.maxRows); this.maxRows += n; }
  setFrozenRows(n) { this.frozen = n; }
  setColumnWidth() {}
  hideColumns(c) { this.hiddenCols.push(c); }
  getLastRow() { let m = 0; for (const k of this.values.keys()) m = Math.max(m, Number(k.split(',')[0])); return m; }
  getLastColumn() { let m = 0; for (const k of this.values.keys()) m = Math.max(m, Number(k.split(',')[1])); return m; }
  getRange(r, c, nr = 1, nc = 1) {
    if (!(r >= 1 && c >= 1 && nr >= 1 && nc >= 1 && r + nr - 1 <= this.maxRows && c + nc - 1 <= this.maxCols)) {
      throw new Error(`Range вне листа ${this.name}: r=${r} c=${c} nr=${nr} nc=${nc} (maxRows=${this.maxRows})`);
    }
    return new FakeRange(this, r, c, nr, nc);
  }
  get(r, c) { const v = this.values.get(this.key(r, c)); return v === undefined ? '' : v.formula !== undefined ? v.formula : v.value; }
  put(r, c, v) {
    const fmt = this.formats.get(this.key(r, c)) || 'General';
    if (v === '' || v === null || v === undefined) { this.values.delete(this.key(r, c)); return; }
    if (typeof v === 'string' && fmt !== '@') {
      if (v.startsWith('=')) { this.values.set(this.key(r, c), { formula: v }); return; }
      if (/^-?\d+(\.\d+)?$/.test(v)) v = Number(v);                       // как в Sheets
      else if (/^\d{4}-\d{2}-\d{2}$/.test(v)) v = new Date(v + 'T00:00:00Z');
    }
    this.values.set(this.key(r, c), { value: v });
  }
}
class FakeRange {
  constructor(sheet, r, c, nr, nc) { Object.assign(this, { sheet, r, c, nr, nc }); }
  getValues() {
    const out = [];
    for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(this.sheet.get(this.r + i, this.c + j)); out.push(row); }
    return out;
  }
  getValue() { return this.sheet.get(this.r, this.c); }
  setValues(v) {
    if (v.length !== this.nr) throw new Error(`setValues: строк ${v.length}, ожидалось ${this.nr}`);
    v.forEach((row, i) => {
      if (row.length !== this.nc) throw new Error(`setValues: колонок ${row.length}, ожидалось ${this.nc} (строка ${i})`);
      row.forEach((x, j) => this.sheet.put(this.r + i, this.c + j, x));
    });
    return this;
  }
  setValue(v) { this.sheet.put(this.r, this.c, v); return this; }
  setFormula(f) { return this.setValue(f); }
  setFormulas(f) { return this.setValues(f); }
  clearContent() { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.sheet.values.delete(this.sheet.key(this.r + i, this.c + j)); return this; }
  setNumberFormat(f) { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.sheet.formats.set(this.sheet.key(this.r + i, this.c + j), f); return this; }
  setFontWeight() { return this; }
  setDataValidation() { return this; }
}
class FakeSpreadsheet {
  constructor(tz) { this.tz = tz; this.sheets = [new FakeSheet('Лист1')]; }
  getSheets() { return this.sheets.slice(); }
  getSheetByName(n) { return this.sheets.find((s) => s.name === n) || null; }
  insertSheet(n, i) { const s = new FakeSheet(n); this.sheets.splice(i === undefined ? this.sheets.length : i, 0, s); return s; }
  deleteSheet(s) { this.sheets = this.sheets.filter((x) => x !== s); }
  getSpreadsheetTimeZone() { return this.tz; }
  toast() {}
}

// ───────────── имитация часовых поясов (Utilities) ─────────────
function tzOffsetMs(ms, tz) {
  const parts = new Intl.DateTimeFormat('en-US', { timeZone: tz, hourCycle: 'h23', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit' }).formatToParts(new Date(ms));
  const g = (t) => Number(parts.find((p) => p.type === t).value);
  return Date.UTC(g('year'), g('month') - 1, g('day'), g('hour'), g('minute'), g('second')) - Math.floor(ms / 1000) * 1000;
}
const p2 = (n) => String(n).padStart(2, '0');

// ───────────── окружение ─────────────
function makeEnv({ tz = 'UTC', now = '2026-10-06T10:00:00Z', interactive = true } = {}) {
  const env = { alerts: [], sleeps: [], fetched: [], triggers: [], props: {}, fetchHandler: null, interactive, now: new Date(now).getTime() };
  const ss = new FakeSpreadsheet(tz);
  env.ss = ss;
  const RealDate = Date;
  class FakeDate extends RealDate {
    constructor(...a) { if (a.length === 0) super(env.now); else super(...a); }
    static now() { return env.now; }
  }
  const sandbox = {
    console: { log() {} },
    Date: FakeDate,
    SpreadsheetApp: {
      getActive: () => ss,
      getUi: () => {
        if (!env.interactive) throw new Error('Cannot call SpreadsheetApp.getUi() from this context.');
        return { alert: (m) => env.alerts.push(m) };
      },
      newDataValidation: () => ({ requireValueInList() { return this; }, build() { return {}; } })
    },
    Utilities: {
      sleep: (ms) => env.sleeps.push(ms),
      formatDate(d, zone, fmt) {
        const l = new RealDate(d.getTime() + tzOffsetMs(d.getTime(), zone));
        const date = `${l.getUTCFullYear()}-${p2(l.getUTCMonth() + 1)}-${p2(l.getUTCDate())}`;
        return fmt === 'yyyy-MM-dd' ? date : fmt === 'yyyy-MM-dd HH:mm' ? `${date} ${p2(l.getUTCHours())}:${p2(l.getUTCMinutes())}` : assert.fail('формат ' + fmt);
      },
      parseDate(s, zone, fmt) {
        assert.strictEqual(fmt, 'yyyy-MM-dd');
        const [y, m, d] = s.split('-').map(Number);
        const guess = RealDate.UTC(y, m - 1, d);
        return new RealDate(guess - tzOffsetMs(guess - tzOffsetMs(guess, zone), zone));
      }
    },
    UrlFetchApp: {
      fetch(url) {
        env.fetched.push(url);
        const r = env.fetchHandler(new URL(url));
        return { getResponseCode: () => r.status, getContentText: () => (typeof r.body === 'string' ? r.body : JSON.stringify(r.body)) };
      }
    },
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => (k in env.props ? env.props[k] : null), setProperty: (k, v) => { env.props[k] = v; } }) },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) },
    ScriptApp: {
      getProjectTriggers: () => env.triggers.slice(),
      deleteTrigger: (t) => { env.triggers = env.triggers.filter((x) => x !== t); },
      newTrigger(fn) {
        const b = { timeBased: () => b, everyDays: (n) => { b.days = n; return b; }, atHour: (h) => { b.hour = h; return b; },
          create() { const t = { getHandlerFunction: () => fn, days: b.days, hour: b.hour }; env.triggers.push(t); return t; } };
        return b;
      }
    }
  };
  vm.createContext(sandbox);
  vm.runInContext(CODE, sandbox, { filename: 'Code.gs' });
  env.ctx = sandbox;
  env.eval = (expr) => plain(vm.runInContext(expr, sandbox));
  return env;
}

// ───────────── данные Meta ─────────────
const ID = { acc: '111', campA: '120200000000000001', campB: '120200000000000002', setA: '120200000000000011', setB: '120200000000000012',
  adA1: '120200000000000111', adA2: '120200000000000112', adB1: '120200000000000121' };
function item(over) {
  return Object.assign({
    date_start: '2026-10-01', date_stop: '2026-10-01', account_id: ID.acc, account_name: 'Acme', account_currency: 'USD',
    campaign_id: ID.campA, campaign_name: 'Camp A', adset_id: ID.setA, adset_name: 'Set A1', ad_id: ID.adA1, ad_name: 'Ad A1',
    spend: '10.50', impressions: '1000', reach: '800', clicks: '40', inline_link_clicks: '25',
    actions: [{ action_type: 'link_click', value: '25' }, { action_type: 'lead', value: '3' }]
  }, over);
}
const metaError = (code, message, extra) => ({ status: code === 190 ? 401 : 400, body: { error: Object.assign({ message, type: 'OAuthException', code }, extra) } });

function prepare(env, { accounts = 'act_111', refreshDays = 7, initialDays = 5, includeToday = 'Нет', leadTypes = 'lead' } = {}) {
  env.ctx.setupSheets();
  const s = env.ss.getSheetByName('Настройки');
  s.put(2, 2, accounts); s.put(3, 2, leadTypes); s.put(4, 2, refreshDays); s.put(5, 2, initialDays); s.put(6, 2, includeToday);
  env.props.META_TOKEN = 'SECRET_TOKEN_123';
  env.alerts.length = 0;
}
const dataRows = (env) => {
  const s = env.ss.getSheetByName('Data');
  const out = [];
  for (let r = 2; r <= s.getLastRow(); r++) out.push(s.getRange(r, 1, 1, 20).getValues()[0]);
  return out;
};
const iso = (d) => d.toISOString().slice(0, 10);

// ───────────── раннер ─────────────
const tests = [];
const test = (name, fn) => tests.push({ name, fn });

// ───────────── тесты ─────────────
test('даты: addDays через границы месяца, года и високосный день', () => {
  const c = makeEnv().ctx;
  assert.strictEqual(c.addDays('2026-10-31', 1), '2026-11-01');
  assert.strictEqual(c.addDays('2026-01-01', -1), '2025-12-31');
  assert.strictEqual(c.addDays('2028-02-28', 1), '2028-02-29');
  assert.strictEqual(c.addDays('2027-02-28', 1), '2027-03-01');
  assert.ok(c.isValidIsoDate('2026-02-28'));
  assert.ok(!c.isValidIsoDate('2026-02-30'));
  assert.ok(!c.isValidIsoDate('2026-2-3'));
  assert.ok(!c.isValidIsoDate('сегодня'));
});

test('buildChunks: включительно, без дыр и перекрытий', () => {
  const c = makeEnv().ctx;
  assert.deepStrictEqual(plain(c.buildChunks('2026-10-01', '2026-10-01', 14)), [{ since: '2026-10-01', until: '2026-10-01' }]);
  assert.deepStrictEqual(plain(c.buildChunks('2026-10-01', '2026-10-28', 14)),
    [{ since: '2026-10-01', until: '2026-10-14' }, { since: '2026-10-15', until: '2026-10-28' }]);
  assert.deepStrictEqual(plain(c.buildChunks('2026-10-01', '2026-10-30', 14)).map((x) => x.until), ['2026-10-14', '2026-10-28', '2026-10-30']);
  assert.deepStrictEqual(plain(c.buildChunks('2026-10-05', '2026-10-01', 14)), []);
});

test('parseSettings: act_, запятые, дубли, значения по умолчанию, ошибки', () => {
  const c = makeEnv().ctx;
  const s = plain(c.parseSettings({ accounts: 'act_111, 222;act_111', leadTypes: ' lead ', refreshDays: '7', initialDays: 90, includeToday: 'Да' }));
  assert.deepStrictEqual(s, { accounts: ['111', '222'], leadTypes: ['lead'], refreshDays: 7, initialDays: 90, includeToday: true });
  const d = plain(c.parseSettings({ accounts: '5', leadTypes: '', refreshDays: 'abc', initialDays: 0, includeToday: '' }));
  assert.deepStrictEqual(d, { accounts: ['5'], leadTypes: ['lead'], refreshDays: 7, initialDays: 90, includeToday: false });
  assert.strictEqual(c.parseSettings({ accounts: '1', refreshDays: 9999 }).refreshDays, 90);
  assert.throws(() => c.parseSettings({ accounts: '' }), /не указан рекламный аккаунт/);
  assert.throws(() => c.parseSettings({ accounts: 'act_12ab' }), /Неверный ID/);
});

test('computeWindow: вчера по умолчанию, сегодня по настройке, ручная дата', () => {
  const c = makeEnv().ctx;
  const base = { initialDays: 90, refreshDays: 7, includeToday: false };
  assert.deepStrictEqual(plain(c.computeWindow(base, '2026-10-06', false)), { since: '2026-09-29', until: '2026-10-05' });
  assert.deepStrictEqual(plain(c.computeWindow(base, '2026-10-06', true)), { since: '2026-07-08', until: '2026-10-05' });
  assert.deepStrictEqual(plain(c.computeWindow({ ...base, includeToday: true }, '2026-10-06', false)), { since: '2026-09-30', until: '2026-10-06' });
  assert.deepStrictEqual(plain(c.computeWindow(base, '2026-10-06', false, '2026-01-01')), { since: '2026-01-01', until: '2026-10-05' });
  assert.throws(() => c.computeWindow(base, '2026-10-06', false, '2026-12-01'), /позже/);
});

test('buildInsightsUrl: уровень ad, по дням, период, токен, нужные поля', () => {
  const c = makeEnv().ctx;
  const u = new URL(c.buildInsightsUrl('111', 'TOK', '2026-10-01', '2026-10-14'));
  assert.strictEqual(u.origin + u.pathname, 'https://graph.facebook.com/v23.0/act_111/insights');
  assert.strictEqual(u.searchParams.get('level'), 'ad');
  assert.strictEqual(u.searchParams.get('time_increment'), '1');
  assert.deepStrictEqual(JSON.parse(u.searchParams.get('time_range')), { since: '2026-10-01', until: '2026-10-14' });
  assert.strictEqual(u.searchParams.get('access_token'), 'TOK');
  const fields = u.searchParams.get('fields').split(',');
  for (const f of ['spend', 'impressions', 'reach', 'clicks', 'inline_link_clicks', 'actions', 'ad_id', 'campaign_name', 'account_currency']) assert.ok(fields.includes(f), f);
  // У UrlFetchApp лимит на длину URL — 2 КБ; токены системных пользователей бывают длинными.
  const long = c.buildInsightsUrl('1234567890123456', 'EAA' + 'x'.repeat(400), '2026-10-01', '2026-10-14');
  assert.ok(long.length < 1500, 'URL слишком длинный: ' + long.length);
});

test('toRow: сумма типов лидов, CPL/CTR/CPC/CPM, пустые значения без деления на ноль', () => {
  const c = makeEnv().ctx;
  const C = makeEnv().eval('COL');
  const r = plain(c.toRow(item({ actions: [{ action_type: 'lead', value: '3' }, { action_type: 'onsite_conversion.lead_grouped', value: '2' }] }), ['lead']));
  assert.strictEqual(r[C.leads], 3);
  assert.strictEqual(r[C.cpl], 3.5);
  assert.strictEqual(r[C.ctr], 0.025);
  assert.strictEqual(r[C.cpc], 0.42);
  assert.strictEqual(r[C.cpm], 10.5);
  assert.strictEqual(plain(c.toRow(item(), ['lead', 'link_click']))[C.leads], 28, 'несколько типов суммируются');
  const z = plain(c.toRow(item({ spend: '0', impressions: '0', inline_link_clicks: undefined, actions: undefined }), ['lead']));
  assert.strictEqual(z[C.leads], 0);
  assert.deepStrictEqual([z[C.cpl], z[C.ctr], z[C.cpc], z[C.cpm]], ['', '', '', '']);
  assert.strictEqual(typeof r[C.ad_id], 'string');
});

test('fetchJson: повтор при лимите запросов с нарастающей паузой, затем успех', () => {
  const env = makeEnv();
  let n = 0;
  env.fetchHandler = () => (++n < 3 ? metaError(17, 'User request limit reached') : { status: 200, body: { data: [] } });
  assert.deepStrictEqual(plain(env.ctx.fetchJson('https://graph.facebook.com/x', 'TOK')), { data: [] });
  assert.strictEqual(n, 3);
  assert.deepStrictEqual(env.sleeps, [2000, 4000]);
});

test('fetchJson: ошибка токена не повторяется, понятный текст, токен вырезан', () => {
  const env = makeEnv();
  let n = 0;
  env.fetchHandler = () => { n++; return metaError(190, 'Error validating access token: SECRET_TOKEN_123 expired'); };
  assert.throws(() => env.ctx.fetchJson('https://graph.facebook.com/x', 'SECRET_TOKEN_123'), (e) => {
    assert.ok(/Токен недействителен/.test(e.message));
    assert.ok(!e.message.includes('SECRET_TOKEN_123'), 'токен утёк в текст ошибки: ' + e.message);
    return true;
  });
  assert.strictEqual(n, 1);
});

test('fetchJson: лимит исчерпан после MAX_RETRIES, не-JSON ответ даёт понятную ошибку', () => {
  const env = makeEnv();
  let n = 0;
  env.fetchHandler = () => { n++; return metaError(4, 'Application request limit reached'); };
  assert.throws(() => env.ctx.fetchJson('https://graph.facebook.com/x', 'TOKEN_XYZ_123'), /Превышен лимит/);
  assert.strictEqual(n, 4);
  env.fetchHandler = () => ({ status: 502, body: '<html>Bad gateway</html>' });
  assert.throws(() => env.ctx.fetchJson('https://graph.facebook.com/x', 'TOKEN_XYZ_123'), /HTTP 502/);
});

test('fetchAccountRows: куски по 14 дней и страницы по paging.next', () => {
  const env = makeEnv();
  const seen = [];
  env.fetchHandler = (u) => {
    const tr = JSON.parse(u.searchParams.get('time_range'));
    const after = u.searchParams.get('after');
    seen.push(`${tr.since}..${tr.until}${after ? '#' + after : ''}`);
    if (tr.since === '2026-10-01' && !after) {
      return { status: 200, body: { data: [item({ ad_id: 'p1' })], paging: { next: u.toString() + '&after=CURSOR2' } } };
    }
    return { status: 200, body: { data: [item({ ad_id: 'p2', date_start: tr.since })] } };
  };
  const rows = env.ctx.fetchAccountRows('111', 'T', { since: '2026-10-01', until: '2026-10-20' }, ['lead']);
  assert.deepStrictEqual(seen, ['2026-10-01..2026-10-14', '2026-10-01..2026-10-14#CURSOR2', '2026-10-15..2026-10-20']);
  assert.strictEqual(rows.length, 3);
});

test('fetchAccountRows: ссылка пагинации на чужой хост отвергается (токен не уйдёт наружу)', () => {
  const env = makeEnv();
  env.fetchHandler = () => {
    // Страховка: если проверка хоста сломается, тест упадёт, а не зациклится.
    assert.ok(env.fetched.length < 3, 'код пошёл по ссылке на чужой хост: ' + env.fetched[env.fetched.length - 1]);
    return { status: 200, body: { data: [item()], paging: { next: 'https://evil.example.com/steal?access_token=T' } } };
  };
  assert.throws(() => env.ctx.fetchAccountRows('111', 'T', { since: '2026-10-01', until: '2026-10-02' }, ['lead']), /неожиданную ссылку/);
  assert.strictEqual(env.fetched.length, 1, 'на чужой хост запрос не должен уйти');
});

test('setupSheets: листы, заголовки, лишний пустой лист удалён, повторный запуск безопасен', () => {
  const env = makeEnv();
  env.ctx.setupSheets();
  assert.deepStrictEqual(env.ss.getSheets().map((s) => s.name), ['Настройки', 'Сводка', 'Data']);
  const data = env.ss.getSheetByName('Data');
  assert.strictEqual(data.get(1, 1), 'Дата');
  assert.strictEqual(data.get(1, 17), 'CPL (цена лида)');
  assert.strictEqual(data.get(1, 20), 'CPM');
  const st = env.ss.getSheetByName('Настройки');
  st.put(2, 2, 'act_999'); st.put(4, 2, 14);
  env.ctx.setupSheets();
  assert.strictEqual(st.get(2, 2), 'act_999', 'настройки пользователя не затираются');
  assert.strictEqual(st.get(4, 2), 14);
});

test('сквозной сценарий: первая загрузка → перезапись окна → новые дни', () => {
  const env = makeEnv();
  prepare(env, { initialDays: 5 });
  const C = env.eval('COL');

  // ── первая загрузка: окно 10-01..10-05, две страницы ответа
  env.fetchHandler = (u) => {
    if (u.searchParams.get('after')) {
      return { status: 200, body: { data: [item({ date_start: '2026-10-03' })] } };
    }
    return { status: 200, body: { data: [
      item({ date_start: '2026-10-01' }),
      item({ date_start: '2026-10-01', ad_id: ID.adA2, ad_name: 'Ad A2', spend: '5', impressions: '500', inline_link_clicks: '10', actions: [] }),
      item({ date_start: '2026-10-02', campaign_id: ID.campB, campaign_name: 'Camp*B?', adset_id: ID.setB, ad_id: ID.adB1, ad_name: 'Ad B1', spend: '20', actions: [{ action_type: 'lead', value: '4' }] })
    ], paging: { next: u.toString() + '&after=P2' } } };
  };
  env.ctx.refreshData();
  let rows = dataRows(env);
  assert.strictEqual(rows.length, 4);
  assert.deepStrictEqual(rows.map((r) => iso(r[C.date])), ['2026-10-01', '2026-10-01', '2026-10-02', '2026-10-03']);
  assert.strictEqual(rows[0][C.ad_id], ID.adA1, '18-значный ID должен остаться строкой без искажений');
  assert.strictEqual(rows[0][C.account_id], '111');
  assert.strictEqual(rows[0][C.spend], 10.5);
  assert.strictEqual(rows[0][C.cpl], 3.5);
  assert.strictEqual(rows[1][C.cpl], '', 'без лидов CPL пустой');
  assert.ok(env.alerts[0].startsWith('Готово'), env.alerts.join('|'));
  const st = env.ss.getSheetByName('Настройки');
  assert.strictEqual(st.get(9, 2), 'OK: строк загружено 4');
  assert.strictEqual(st.get(10, 2), '2026-10-01 … 2026-10-05');

  // ── сводка: список кампаний и формулы
  const sum = env.ss.getSheetByName('Сводка');
  assert.deepStrictEqual([sum.get(7, 1), sum.get(8, 1)], ['Camp A', 'Camp*B?'], 'кампании по алфавиту');
  assert.strictEqual(sum.get(7, 10), 'Camp A');
  assert.strictEqual(sum.get(8, 10), 'Camp~*B~?', 'спецсимволы SUMIFS экранированы');
  assert.strictEqual(sum.get(8, 2), '=SUMIFS(Data!$K:$K,Data!$F:$F,"="&$J8,Data!$A:$A,">="&$B$2,Data!$A:$A,"<="&$B$3)');
  assert.strictEqual(sum.get(8, 6), '=IF(E8>0,B8/E8,"")');
  assert.strictEqual(sum.get(6, 5), '=SUMIFS(Data!$P:$P,Data!$A:$A,">="&$B$2,Data!$A:$A,"<="&$B$3)');
  assert.strictEqual(sum.get(2, 2), '=TODAY()-29');

  // ── следующий день: окно 4 дня (10-03..10-06). 10-03 пересчитан, adA2 исчез, добавились новые дни
  env.now = new Date('2026-10-07T10:00:00Z').getTime();
  env.ss.getSheetByName('Настройки').put(4, 2, 4);
  env.fetchHandler = (u) => {
    assert.deepStrictEqual(JSON.parse(u.searchParams.get('time_range')), { since: '2026-10-03', until: '2026-10-06' });
    return { status: 200, body: { data: [
      item({ date_start: '2026-10-03', spend: '12', actions: [{ action_type: 'lead', value: '4' }] }),
      item({ date_start: '2026-10-04' }),
      item({ date_start: '2026-10-06', spend: '7' })
    ] } };
  };
  env.ctx.refreshData();
  rows = dataRows(env);
  assert.deepStrictEqual(rows.map((r) => iso(r[C.date])), ['2026-10-01', '2026-10-01', '2026-10-02', '2026-10-03', '2026-10-04', '2026-10-06']);
  assert.strictEqual(rows[0][C.spend], 10.5, 'история до окна не тронута');
  assert.strictEqual(rows[2][C.campaign_name], 'Camp*B?');
  assert.strictEqual(rows[3][C.spend], 12, 'день внутри окна перезаписан свежими данными');
  assert.strictEqual(rows[3][C.leads], 4);
  assert.strictEqual(rows.filter((r) => iso(r[C.date]) === '2026-10-03').length, 1, 'дубликатов нет');

  // ── повторный запуск с теми же данными идемпотентен
  const before = JSON.stringify(dataRows(env));
  env.ctx.refreshData();
  assert.strictEqual(JSON.stringify(dataRows(env)), before);

  // ── история с более ранней даты заменяет всё с этой даты и не дублирует
  env.fetchHandler = () => ({ status: 200, body: { data: [item({ date_start: '2026-09-30' }), item({ date_start: '2026-10-01' })] } });
  env.ctx.runRefresh({ interactive: true, since: '2026-09-30' });
  rows = dataRows(env);
  assert.deepStrictEqual(rows.map((r) => iso(r[C.date])), ['2026-09-30', '2026-10-01']);
});

test('несколько аккаунтов: сбой одного не стирает его старые строки, а соседний обновляется', () => {
  const env = makeEnv({ interactive: false });
  prepare(env, { accounts: 'act_111, act_222', initialDays: 3 });
  const C = env.eval('COL');
  const byAccount = (u) => u.pathname.split('/')[2];

  env.fetchHandler = (u) => ({ status: 200, body: { data: [item({ date_start: '2026-10-04', account_id: byAccount(u).slice(4), spend: byAccount(u) === 'act_111' ? '1' : '2' })] } });
  env.ctx.refreshDataScheduled();
  assert.strictEqual(dataRows(env).length, 2);

  env.fetchHandler = (u) => byAccount(u) === 'act_222'
    ? metaError(190, 'expired')
    : { status: 200, body: { data: [item({ date_start: '2026-10-04', account_id: '111', spend: '99' })] } };
  assert.throws(() => env.ctx.refreshDataScheduled(), /act_222.*Токен недействителен/s);
  const rows = dataRows(env);
  assert.strictEqual(rows.length, 2);
  assert.strictEqual(rows.find((r) => r[C.account_id] === '111')[C.spend], 99, 'исправный аккаунт обновился');
  assert.strictEqual(rows.find((r) => r[C.account_id] === '222')[C.spend], 2, 'старые данные упавшего аккаунта сохранены');
  const status = env.ss.getSheetByName('Настройки').get(9, 2);
  assert.ok(/^Ошибки: act_222/.test(status), status);
});

test('ручной запуск при ошибке показывает окно с причиной и не бросает исключение', () => {
  const env = makeEnv();
  prepare(env);
  env.props.META_TOKEN = 'T';
  env.fetchHandler = () => metaError(200, '(#200) Requires ads_read permission');
  env.ctx.refreshData();
  assert.strictEqual(env.alerts.length, 1);
  assert.ok(/ads_read/.test(env.alerts[0]));
  assert.strictEqual(dataRows(env).length, 0);
});

test('нет токена или аккаунта: понятная ошибка, а не падение', () => {
  const env = makeEnv();
  env.ctx.setupSheets();
  env.alerts.length = 0;
  env.ctx.refreshData();
  assert.ok(/не указан рекламный аккаунт/.test(env.alerts[0]), env.alerts[0]);
  env.ss.getSheetByName('Настройки').put(2, 2, 'act_5');
  env.ctx.refreshData();
  assert.ok(/Не задан токен/.test(env.alerts[1]), env.alerts[1]);
});

test('часовой пояс: даты не съезжают на день при записи и чтении (Asia/Dubai, UTC-5)', () => {
  for (const tz of ['Asia/Dubai', 'America/New_York', 'Pacific/Auckland']) {
    const env = makeEnv({ tz, now: '2026-10-06T10:00:00Z' });
    prepare(env, { initialDays: 3 });
    const C = env.eval('COL');
    env.fetchHandler = () => ({ status: 200, body: { data: [item({ date_start: '2026-10-03' }), item({ date_start: '2026-10-05' })] } });
    env.ctx.refreshData();
    env.ctx.refreshData(); // второй проход читает даты из таблицы и должен их узнать
    const rows = dataRows(env);
    assert.strictEqual(rows.length, 2, tz);
    const seen = rows.map((r) => env.ctx.Utilities.formatDate(r[C.date], tz, 'yyyy-MM-dd'));
    assert.deepStrictEqual(seen, ['2026-10-03', '2026-10-05'], tz);
  }
});

test('writeData: рост листа за пределы 1000 строк и очистка лишнего хвоста', () => {
  const env = makeEnv();
  const C = env.eval('COL');
  const sheet = new FakeSheet('Data', 5);
  env.ctx.setupDataSheet(sheet);
  const mk = (date, ad) => env.ctx.toRow(item({ date_start: date, ad_id: ad }), ['lead']);
  const win = (since, until) => ({ since, until, accountIds: ['111'] });
  env.ctx.writeData(sheet, [1, 2, 3, 4, 5, 6, 7].map((i) => mk('2026-10-01', 'a' + i)), win('2026-10-01', '2026-10-01'), 'UTC');
  assert.strictEqual(sheet.getLastRow(), 8);
  assert.ok(sheet.getMaxRows() >= 8);
  env.ctx.writeData(sheet, [mk('2026-10-01', 'only')], win('2026-10-01', '2026-10-01'), 'UTC');
  assert.strictEqual(sheet.getLastRow(), 2, 'лишние строки стёрты');
  assert.strictEqual(sheet.get(2, C.ad_id + 1), 'only');
});

test('ID и названия из цифр остаются текстом', () => {
  const env = makeEnv();
  prepare(env, { initialDays: 2 });
  const C = env.eval('COL');
  env.fetchHandler = () => ({ status: 200, body: { data: [item({ date_start: '2026-10-04', campaign_name: '100', ad_name: '2026-10-05', ad_id: '120200000000000999' })] } });
  env.ctx.refreshData();
  const r = dataRows(env)[0];
  assert.strictEqual(r[C.campaign_name], '100');
  assert.strictEqual(r[C.ad_name], '2026-10-05');
  assert.strictEqual(r[C.ad_id], '120200000000000999');
});

test('ежедневный триггер: один, не плодится, выключается', () => {
  const env = makeEnv();
  env.ctx.enableDailyTrigger();
  env.ctx.enableDailyTrigger();
  assert.strictEqual(env.triggers.length, 1);
  assert.strictEqual(env.triggers[0].getHandlerFunction(), 'refreshDataScheduled');
  assert.strictEqual(env.triggers[0].days, 1);
  env.triggers.push({ getHandlerFunction: () => 'чужой' });
  env.ctx.disableDailyTrigger();
  assert.deepStrictEqual(env.triggers.map((t) => t.getHandlerFunction()), ['чужой'], 'чужие триггеры не трогаем');
});

test('testConnection показывает аккаунт и ошибку по каждому аккаунту отдельно', () => {
  const env = makeEnv();
  prepare(env, { accounts: 'act_111, act_222' });
  env.fetchHandler = (u) => u.pathname.includes('act_111')
    ? { status: 200, body: { name: 'Acme', currency: 'USD', timezone_name: 'America/Los_Angeles', id: 'act_111' } }
    : metaError(100, 'Unsupported get request');
  env.ctx.testConnection();
  assert.ok(/✔ act_111: Acme, валюта USD, часовой пояс America\/Los_Angeles/.test(env.alerts[0]), env.alerts[0]);
  assert.ok(/✘ act_222/.test(env.alerts[0]), env.alerts[0]);
});

// ───────────── запуск ─────────────
let failed = 0;
for (const t of tests) {
  try { t.fn(); console.log('  ok   ' + t.name); } catch (e) {
    failed++; console.log('  FAIL ' + t.name + '\n       ' + String(e.stack || e).split('\n').slice(0, 6).join('\n       '));
  }
}
console.log(`\n${tests.length - failed}/${tests.length} тестов прошло`);
process.exit(failed ? 1 : 0);
