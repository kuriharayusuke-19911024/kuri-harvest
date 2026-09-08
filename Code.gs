/**
 * 丹波栗 収穫管理アプリ — Google Apps Script API
 * 株式会社丹波農商
 *
 * スプレッドシート「栗収穫管理」に紐づくコンテナバインドのスクリプトとして使うのが基本。
 * （スプレッドシートを開いて 拡張機能 > Apps Script）
 * スタンドアロンで使う場合だけ SPREADSHEET_ID にIDを入れる。
 *
 * デプロイ：デプロイ > 新しいデプロイ > 種類「ウェブアプリ」
 *           実行ユーザー「自分」／アクセスできるユーザー「全員」
 *
 * API
 *   GET  ?action=getAll                                  -> {fields,vars,sizes,sups,prices,harvest,ship,buy}
 *   POST {action:'add',        kind:'harvest'|'ship'|'buy', record:{...}}
 *   POST {action:'delete',     kind:'harvest'|'ship'|'buy', id:'...'}
 *   POST {action:'saveSettings', fields:[], vars:[], sizes:[], sups:[], prices:{}}
 *
 * 応答はすべて {ok:true, data:...} / {ok:false, error:'...'}
 */

var SPREADSHEET_ID = '';                       // 空ならアクティブなスプレッドシート
var SETTING_SHEET  = '設定';
var SHEETS = { harvest: '収穫', ship: '出荷', buy: '買取' };

var DEFAULTS = {
  fields: ['今中', '段宿'],
  vars:   ['銀寄', '筑波', '美栗', '丹沢'],
  sizes:  ['3L以上', '2L', 'L', 'M'],
  sups:   ['荻野運送', '木寺様', '余田様', '高畑様'],
  prices: { '3L以上': 1700, '2L': 1500, 'L': 1100, 'M': 700 }
};

/* ===================== エントリポイント ===================== */

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || 'getAll';
    if (action === 'getAll') return json({ ok: true, data: getAll() });
    return json({ ok: false, error: 'unknown action: ' + action });
  } catch (err) {
    return json({ ok: false, error: errMsg(err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);   // 3人同時操作でも行が混ざらないように直列化
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var action = body.action;

    if (action === 'add')          return json({ ok: true, data: addRecord(body.kind, body.record) });
    if (action === 'delete')       return json({ ok: true, data: deleteRecord(body.kind, body.id) });
    if (action === 'saveSettings') return json({ ok: true, data: saveSettings(body) });
    if (action === 'getAll')       return json({ ok: true, data: getAll() });

    return json({ ok: false, error: 'unknown action: ' + action });
  } catch (err) {
    return json({ ok: false, error: errMsg(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/** 初回だけエディタから手動実行するとシートを一式作る（任意） */
function setup() {
  var ss = ss_();
  var st = readSettings(ss);
  ['harvest', 'ship', 'buy'].forEach(function (kind) {
    sheet_(ss, SHEETS[kind], headersFor(kind, st.sizes));
  });
  return 'OK: ' + ss.getName();
}

/* ===================== 処理本体 ===================== */

function getAll() {
  var ss = ss_();
  var st = readSettings(ss);
  return {
    fields:  st.fields,
    vars:    st.vars,
    sizes:   st.sizes,
    sups:    st.sups,
    prices:  st.prices,
    harvest: readRecords(ss, 'harvest', st.sizes),
    ship:    readRecords(ss, 'ship',    st.sizes),
    buy:     readRecords(ss, 'buy',     st.sizes)
  };
}

function addRecord(kind, rec) {
  if (!SHEETS[kind])   throw new Error('unknown kind: ' + kind);
  if (!rec || !rec.id) throw new Error('record.id が必要です');

  var ss = ss_();
  var st = readSettings(ss);
  // 設定にないサイズがレコードに含まれていても取りこぼさない
  var sizes = unique(st.sizes.concat(Object.keys(rec.q || {})));

  var sh   = sheet_(ss, SHEETS[kind], headersFor(kind, st.sizes));
  var head = ensureHeaders(sh, headersFor(kind, sizes));

  if (findRowById(sh, head, rec.id) > 0) return { id: rec.id, duplicated: true };

  var map = buildRowMap(kind, rec, sizes);
  sh.appendRow(head.map(function (h) { return (h in map) ? map[h] : ''; }));
  return { id: rec.id };
}

function deleteRecord(kind, id) {
  if (!SHEETS[kind]) throw new Error('unknown kind: ' + kind);
  var sh = ss_().getSheetByName(SHEETS[kind]);
  if (!sh) return { deleted: 0 };

  var vals = sh.getDataRange().getValues();
  if (vals.length < 2) return { deleted: 0 };

  var head = headRow(vals[0]);
  var i = head.indexOf('id');
  if (i < 0) return { deleted: 0 };

  var target = String(id).trim();
  for (var r = vals.length - 1; r >= 1; r--) {
    if (String(vals[r][i]).trim() === target) {
      sh.deleteRow(r + 1);
      return { deleted: 1 };
    }
  }
  return { deleted: 0 };   // 既に無い場合もエラーにしない
}

function saveSettings(body) {
  var ss  = ss_();
  var cur = readSettings(ss);
  var st = {
    fields: Array.isArray(body.fields) ? body.fields.map(String) : cur.fields,
    vars:   Array.isArray(body.vars)   ? body.vars.map(String)   : cur.vars,
    sizes:  Array.isArray(body.sizes)  ? body.sizes.map(String)  : cur.sizes,
    sups:   Array.isArray(body.sups)   ? body.sups.map(String)   : cur.sups,
    prices: (body.prices && typeof body.prices === 'object' && !Array.isArray(body.prices))
              ? body.prices : cur.prices
  };
  writeSettings(sheet_(ss, SETTING_SHEET, ['key', 'value']), st);
  return st;
}

/* ===================== シート入出力 ===================== */

function headersFor(kind, sizes) {
  if (kind === 'harvest') {
    return ['id', '日付', '園地', '品種'].concat(sizes).concat(['合計kg', 'メモ', '登録日時']);
  }
  if (kind === 'ship') {
    return ['id', '日付', '出荷先', '品種'].concat(sizes).concat(['合計kg', '単価', 'メモ', '登録日時']);
  }
  return ['id', '日付', '買取先']
    .concat(sizes.map(function (s) { return s + '_kg'; }))
    .concat(sizes.map(function (s) { return s + '_単価'; }))
    .concat(['合計kg', '金額', 'メモ', '登録日時']);
}

function buildRowMap(kind, rec, sizes) {
  var q = rec.q || {}, m = {}, total = 0;

  m['id']       = String(rec.id);
  m['日付']     = String(rec.date || '');
  m['メモ']     = String(rec.memo || '');
  m['登録日時'] = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm:ss');

  if (kind === 'buy') {
    var p = rec.p || {}, amt = 0;
    m['買取先'] = String(rec.sup || '');
    sizes.forEach(function (s) {
      var w = num(q[s]), pr = num(p[s]);
      m[s + '_kg'] = w;
      m[s + '_単価'] = pr;
      total += w; amt += w * pr;
    });
    m['金額'] = Math.round(amt);
  } else {
    sizes.forEach(function (s) { var w = num(q[s]); m[s] = w; total += w; });
    m['品種'] = String(rec['var'] || '');
    if (kind === 'harvest') {
      m['園地'] = String(rec.field || '');
    } else {
      m['出荷先'] = String(rec.dest || '');
      m['単価'] = (rec.price === null || rec.price === undefined || rec.price === '') ? '' : num(rec.price);
    }
  }
  m['合計kg'] = Math.round(total * 10) / 10;
  return m;
}

function readRecords(ss, kind, sizes) {
  var sh = ss.getSheetByName(SHEETS[kind]);
  if (!sh) return [];

  var vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];

  var head = headRow(vals[0]);
  var idx = {};
  head.forEach(function (h, i) { if (h && !(h in idx)) idx[h] = i; });
  if (!('id' in idx)) return [];

  var col = function (row, name) { var i = idx[name]; return (i === undefined) ? '' : row[i]; };
  var out = [];

  for (var r = 1; r < vals.length; r++) {
    var row = vals[r];
    var id  = String(col(row, 'id') || '').trim();
    if (!id) continue;

    var rec = { id: id, date: dstr(col(row, '日付')), memo: String(col(row, 'メモ') || '') };
    var q = {};

    if (kind === 'buy') {
      var p = {};
      sizes.forEach(function (s) { q[s] = num(col(row, s + '_kg')); p[s] = num(col(row, s + '_単価')); });
      rec.q = q; rec.p = p;
      rec.sup = String(col(row, '買取先') || '');
    } else {
      sizes.forEach(function (s) { q[s] = num(col(row, s)); });
      rec.q = q;
      rec['var'] = String(col(row, '品種') || '');
      if (kind === 'harvest') {
        rec.field = String(col(row, '園地') || '');
      } else {
        rec.dest = String(col(row, '出荷先') || '');
        var pv = col(row, '単価');
        rec.price = (pv === '' || pv === null || pv === undefined) ? null : num(pv);
      }
    }
    out.push(rec);
  }
  return out;
}

function readSettings(ss) {
  var sh = sheet_(ss, SETTING_SHEET, ['key', 'value']);
  var vals = sh.getDataRange().getValues();
  var raw = {};

  for (var r = 1; r < vals.length; r++) {
    var k = String(vals[r][0] || '').trim();
    if (!k) continue;
    try { raw[k] = JSON.parse(String(vals[r][1])); } catch (e) { raw[k] = vals[r][1]; }
  }

  var st = {};
  ['fields', 'vars', 'sizes', 'sups'].forEach(function (k) {
    st[k] = Array.isArray(raw[k]) ? raw[k].map(String) : DEFAULTS[k].slice();
  });
  st.prices = (raw.prices && typeof raw.prices === 'object' && !Array.isArray(raw.prices))
                ? raw.prices : copy(DEFAULTS.prices);

  if (vals.length < 2) writeSettings(sh, st);   // 初回は既定値を書き込む
  return st;
}

function writeSettings(sh, st) {
  var rows = [
    ['fields', JSON.stringify(st.fields)],
    ['vars',   JSON.stringify(st.vars)],
    ['sizes',  JSON.stringify(st.sizes)],
    ['sups',   JSON.stringify(st.sups)],
    ['prices', JSON.stringify(st.prices)]
  ];
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, Math.max(sh.getLastColumn(), 2)).clearContent();
  sh.getRange(2, 1, rows.length, 2).setValues(rows);
}

/* ===================== 小物 ===================== */

function ss_() {
  if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);
  var a = SpreadsheetApp.getActiveSpreadsheet();
  if (!a) throw new Error('スプレッドシートが見つかりません。SPREADSHEET_ID を設定してください。');
  return a;
}

function sheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
  return sh;
}

/** 足りない見出しを右端に追記して、確定した見出し配列を返す */
function ensureHeaders(sh, names) {
  var head = headRow(sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0]);
  var add = [];
  names.forEach(function (n) {
    if (head.indexOf(n) < 0 && add.indexOf(n) < 0) add.push(n);
  });
  if (add.length) {
    sh.getRange(1, head.length + 1, 1, add.length).setValues([add]);
    head = head.concat(add);
  }
  return head;
}

function findRowById(sh, head, id) {
  var i = head.indexOf('id');
  if (i < 0 || sh.getLastRow() < 2) return -1;
  var col = sh.getRange(2, i + 1, sh.getLastRow() - 1, 1).getValues();
  var target = String(id).trim();
  for (var r = 0; r < col.length; r++) {
    if (String(col[r][0]).trim() === target) return r + 2;
  }
  return -1;
}

/** 見出し行を文字列化し、右端の空セルを落とす */
function headRow(row) {
  var h = row.map(function (v) { return String(v == null ? '' : v).trim(); });
  while (h.length && h[h.length - 1] === '') h.pop();
  return h;
}

/** 日付セルは Date でも文字列でも 'yyyy-MM-dd' に揃える */
function dstr(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  return String(v == null ? '' : v).trim();
}

function num(v) {
  var n = parseFloat(v);
  return isFinite(n) ? n : 0;
}

function unique(arr) {
  var seen = {}, out = [];
  arr.forEach(function (v) {
    var k = String(v);
    if (k && !seen[k]) { seen[k] = 1; out.push(k); }
  });
  return out;
}

function copy(o) { return JSON.parse(JSON.stringify(o)); }

function tz_() { return Session.getScriptTimeZone() || 'Asia/Tokyo'; }

function errMsg(err) { return String((err && err.message) ? err.message : err); }

function json(o) {
  return ContentService
    .createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
