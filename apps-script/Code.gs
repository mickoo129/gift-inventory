/**
 * 贈品記數 — Google Sheet 後台
 *
 * 用法：
 * 1. 改下面 PASSCODE 做你自己嘅密碼（網頁連接時要輸入）
 * 2. 揀上面「setup」→ 按「執行」（第一次會問授權，按允許）
 * 3. 部署 → 新增部署作業 → 網頁應用程式
 *    執行身分：我　｜　存取權：所有人
 * 4. 複製「網頁應用程式網址」，貼入網頁嘅「設定」
 *
 * 之後改咗呢度嘅程式碼：部署 → 管理部署作業 → 鉛筆 → 版本揀「新版本」→ 部署
 * （網址唔會變）
 */

const PASSCODE = '請改做你自己嘅密碼';

const TABLES = {
  products: ['year', 'code', 'name', 'total', 'sort'],
  people:   ['year', 'name', 'remark', 'sort'],
  orders:   ['year', 'person', 'code', 'qty'],
  events:   ['id', 'batch', 'year', 'type', 'person', 'code', 'qty', 'date', 'remark', 'createdAt'],
};

const SHEET_NAMES = { products: '產品', people: '客人', orders: '訂單', events: '變動記錄' };

// 同一年入面唔可以重複嘅欄（防止重複匯入）
const UNIQUE = { products: ['year', 'code'], people: ['year', 'name'], events: ['id'] };

/* ---------- 一次性設定 ---------- */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(TABLES).forEach(function (t) {
    let sh = ss.getSheetByName(SHEET_NAMES[t]);
    if (!sh) sh = ss.insertSheet(SHEET_NAMES[t]);
    if (sh.getLastRow() === 0) {
      sh.appendRow(TABLES[t]);
      sh.setFrozenRows(1);
    }
    // 全部欄設做純文字，避免 Google 自動將日期/編號變格式
    sh.getRange(1, 1, sh.getMaxRows(), TABLES[t].length).setNumberFormat('@');
  });
}

/* ---------- Google Sheet 選單 ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('贈品記數').addItem('更新全部總表', 'rebuildAllSummaries').addToUi();
}

function rebuildAllSummaries() {
  const ys = {};
  readTable_('products').forEach(function (r) { ys[r.year] = true; });
  Object.keys(ys).forEach(buildSummary_);
}

/* ---------- 網頁入口 ---------- */
function doGet(e) {
  return handle_(e.parameter || {});
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: '資料格式錯誤' }); }
  return handle_(body);
}

function handle_(req) {
  try {
    if (String(req.passcode || '') !== PASSCODE) return json_({ ok: false, error: '密碼唔啱' });
    const year = String(req.year || '');
    if (req.action === 'load') return json_({ ok: true, data: load_(year) });
    if (req.action === 'batch') {
      const lock = LockService.getScriptLock();
      lock.waitLock(20000);
      let touched;
      try { touched = runBatch_(req.ops || []); } finally { lock.releaseLock(); }
      try { touched.forEach(buildSummary_); } catch (err) { /* 總表出錯唔影響儲存 */ }
      return json_({ ok: true, data: load_(year) });
    }
    return json_({ ok: false, error: '未知指令' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/* ---------- 讀 ---------- */
function load_(year) {
  const out = { years: [] };
  const yearSet = {};
  Object.keys(TABLES).forEach(function (t) {
    const rows = readTable_(t);
    rows.forEach(function (r) { if (r.year) yearSet[r.year] = true; });
    out[t] = year ? rows.filter(function (r) { return r.year === year; }) : [];
  });
  out.years = Object.keys(yearSet).sort();
  return out;
}

function sheet_(t) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES[t]);
  if (!sh) throw new Error('搵唔到工作表「' + SHEET_NAMES[t] + '」，請先執行 setup');
  return sh;
}

function readTable_(t) {
  const sh = sheet_(t);
  const values = sh.getDataRange().getValues();
  const cols = TABLES[t];
  const tz = Session.getScriptTimeZone();
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const v = values[i];
    if (v.every(function (c) { return c === '' || c === null; })) continue;
    const o = {};
    cols.forEach(function (c, j) {
      let x = v[j];
      if (x instanceof Date) x = Utilities.formatDate(x, tz, 'yyyy-MM-dd');
      o[c] = x === null || x === undefined ? '' : String(x);
    });
    rows.push(o);
  }
  return rows;
}

/* ---------- 寫 ---------- */
function matches_(row, where) {
  return Object.keys(where).every(function (k) { return String(row[k]) === String(where[k]); });
}

function runBatch_(ops) {
  // 先喺記憶體改好，冇出錯先至每張表一次過寫返落去
  const cache = {};
  const touched = {};
  function get(t) {
    if (!TABLES[t]) throw new Error('未知資料表：' + t);
    if (!cache[t]) cache[t] = readTable_(t);
    return cache[t];
  }
  function key(t, r) { return UNIQUE[t].map(function (c) { return String(r[c]); }).join('|'); }
  ops.forEach(function (op) {
    const rows = get(op.table);
    if (op.where && op.where.year) touched[op.where.year] = true;
    if (op.op === 'insert') {
      const seen = {};
      if (UNIQUE[op.table]) rows.forEach(function (r) { seen[key(op.table, r)] = true; });
      (op.rows || [op.row]).forEach(function (r) {
        const o = {};
        TABLES[op.table].forEach(function (c) { o[c] = r[c] === undefined || r[c] === null ? '' : String(r[c]); });
        if (UNIQUE[op.table]) {
          const k = key(op.table, o);
          if (seen[k]) {
            if (op.table === 'products') throw new Error(o.year + ' 年已經有產品 ' + o.code + '，唔可以重複加入（可能已經匯入過）');
            if (op.table === 'people') throw new Error(o.year + ' 年已經有「' + o.name + '」');
            throw new Error('重複記錄，可能已經儲存咗，請重新整理睇吓');
          }
          seen[k] = true;
        }
        if (o.year) touched[o.year] = true;
        rows.push(o);
      });
    } else if (op.op === 'delete') {
      if (!op.where || !Object.keys(op.where).length) throw new Error('刪除要有條件');
      cache[op.table] = rows.filter(function (r) { return !matches_(r, op.where); });
    } else if (op.op === 'update') {
      if (!op.where || !Object.keys(op.where).length) throw new Error('更新要有條件');
      rows.forEach(function (r) {
        if (!matches_(r, op.where)) return;
        Object.keys(op.set || {}).forEach(function (k) {
          if (TABLES[op.table].indexOf(k) >= 0) r[k] = String(op.set[k]);
        });
      });
    } else {
      throw new Error('未知操作：' + op.op);
    }
  });
  Object.keys(cache).forEach(function (t) { writeTable_(t, cache[t]); });
  return Object.keys(touched);
}

function writeTable_(t, rows) {
  const sh = sheet_(t);
  const cols = TABLES[t];
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, cols.length).clearContent();
  if (!rows.length) return;
  const range = sh.getRange(2, 1, rows.length, cols.length);
  range.setNumberFormat('@');
  range.setValues(rows.map(function (r) { return cols.map(function (c) { return r[c]; }); }));
}

/* ---------- 總表（同以前 Excel 一樣嘅格式） ---------- */
function buildSummary_(year) {
  if (!year) return;
  const n = function (x) { const v = Number(x); return isFinite(v) ? v : 0; };
  const bySort = function (a, b) { return n(a.sort) - n(b.sort); };
  const inYear = function (r) { return r.year === year; };
  const prods = readTable_('products').filter(inYear).sort(bySort);
  const people = readTable_('people').filter(inYear).sort(bySort);
  const orders = readTable_('orders').filter(inYear);
  const events = readTable_('events').filter(inYear);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = '總表 ' + year;
  let sh = ss.getSheetByName(name);
  if (!prods.length) { if (sh) ss.deleteSheet(sh); return; }
  if (!sh) sh = ss.insertSheet(name);

  const ord = {}, pick = {}, last = {}, P = {};
  prods.forEach(function (p) { P[p.code] = { total: n(p.total), ordered: 0, picked: 0, pending: 0, committed: 0, gift: 0, sale: 0 }; });
  orders.forEach(function (o) { (ord[o.person] = ord[o.person] || {})[o.code] = ((ord[o.person] || {})[o.code] || 0) + n(o.qty); });
  events.forEach(function (e) {
    if (e.type === 'pickup') {
      (pick[e.person] = pick[e.person] || {})[e.code] = ((pick[e.person] || {})[e.code] || 0) + n(e.qty);
      if (!last[e.person] || e.date > last[e.person]) last[e.person] = e.date;
    } else if (P[e.code] && (e.type === 'gift' || e.type === 'sale')) P[e.code][e.type] += n(e.qty);
  });
  const names = {};
  Object.keys(ord).concat(Object.keys(pick)).forEach(function (k) { names[k] = true; });
  Object.keys(names).forEach(function (who) {
    prods.forEach(function (p) {
      const o = (ord[who] || {})[p.code] || 0, k = (pick[who] || {})[p.code] || 0;
      const s = P[p.code];
      s.ordered += o; s.picked += k; s.pending += Math.max(0, o - k); s.committed += Math.max(o, k);
    });
  });

  const rows = [];
  rows.push(['名', '狀態', '最後領取', '備註'].concat(prods.map(function (p) { return p.code + ' ' + p.name; })));
  const line = function (label, f) { rows.push([label, '', '', ''].concat(prods.map(function (p) { return String(f(P[p.code])); }))); };
  line('總數', function (s) { return s.total; });
  line('可再分', function (s) { return s.total - s.committed - s.gift - s.sale; });
  line('未攞', function (s) { return s.pending; });
  line('已攞', function (s) { return s.picked; });
  line('賣出', function (s) { return s.sale; });
  line('送出', function (s) { return s.gift; });
  rows.push(['', '', '', ''].concat(prods.map(function () { return ''; })));
  people.forEach(function (pp) {
    const o = ord[pp.name] || {}, k = pick[pp.name] || {};
    let totO = 0, pend = 0, any = false;
    const cells = prods.map(function (p) {
      const a = o[p.code] || 0, b = k[p.code] || 0;
      totO += a; pend += Math.max(0, a - b); if (b) any = true;
      if (!a && !b) return '';
      if (b >= a) return b + ' ✓';
      if (b === 0) return String(a);
      return b + '/' + a;
    });
    const st = !totO ? '冇訂' : pend === 0 ? '攞晒' : any ? '攞咗部分' : '未攞';
    rows.push([pp.name, st, last[pp.name] || '', pp.remark || ''].concat(cells));
  });
  rows.push(['', '', '', ''].concat(prods.map(function () { return ''; })));
  rows.push(['說明：數字＝訂咗未攞；3/7＝訂 7 件攞咗 3 件；7 ✓＝攞晒。呢張表會自動更新，喺度改唔會影響網頁。', '', '', ''].concat(prods.map(function () { return ''; })));

  sh.clear();
  const range = sh.getRange(1, 1, rows.length, rows[0].length);
  range.setNumberFormat('@');
  range.setValues(rows);
  sh.getRange(1, 1, 1, rows[0].length).setFontWeight('bold').setWrap(true).setVerticalAlignment('top');
  sh.getRange(2, 1, 6, rows[0].length).setBackground('#f6ead3');
  sh.getRange(3, 1, 1, rows[0].length).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.setFrozenColumns(1);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
