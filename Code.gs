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
 *   GET  ?action=getAll
 *     -> {fields,sizes,sups,mems,tasks,workers,prices,peelPrice,
 *         harvest,ship,buy,peel,peelOut,peelShip,schedule}
 *   POST {action:'add',          kind:'harvest'|'ship'|'buy'|'peel'|'peelOut'|'peelShip', record:{...}}
 *   POST {action:'delete',       kind:同上, id:'...'}
 *   POST {action:'saveSettings', fields?, sizes?, sups?, mems?, tasks?,
 *                                workers?, prices?, peelPrice?, supPrices?, goodsItems?, orderMap?, schedule?}
 *   POST {action:'orderShip',   orderId, date, dest, lines:[{name,qty}], memo}   ← 注文発送管理アプリから
 *   POST {action:'orderCancel', orderId}
 *   商品名は設定 orderMap で対応づけ。無い名前は guessLine_ で言葉から推定し、
 *   それでも分からないものは「連携ログ」シートに残り、画面上部に注意が出る。
 *
 * 応答はすべて {ok:true, data:...} / {ok:false, error:'...'}
 */

var SPREADSHEET_ID = '';                       // 空ならアクティブなスプレッドシート
var SETTING_SHEET  = '設定';
var SHEETS = { harvest: '収穫', ship: '出荷', buy: '買取', peel: '内職', peelOut: '内職渡し', peelShip: '渋皮出荷', goods: '加工品' };
var LOG_SHEET = '連携ログ';   // 注文発送管理アプリからの自動連携の記録（対応づけできなかった商品名もここに残る）

var DEFAULTS = {
  fields:    ['今中', '段宿'],
  sizes:     ['3L以上', '2L', 'L', 'M'],
  sups:      ['荻野運送', '木寺様', '余田様', '高畑様'],
  mems:      ['栗原直人', '栗原優介', '秋山龍之輔', '栗原亜美', '木村みき', '栗原美智代', '松本琉雅'],
  tasks:     ['栗拾い', '選別', '出荷', '買取受入'],
  workers:   ['栗原直人', '栗原優介', '秋山龍之輔', '栗原亜美', '木村みき', '栗原美智代', '松本琉雅'],
  prices:    { '3L以上': 1700, '2L': 1500, 'L': 1100, 'M': 700 },
  peelPrice: { good: 1500, bad: 300 },
  supPrices: { '荻野運送': { '3L以上': 2300, '2L': 1800, 'L': 1400 } },
  goodsItems: [
    { name: '丹波栗ペースト 500g', unit: 'パック' },
    { name: '丹波栗ペースト 1kg', unit: 'パック' },
    { name: '丹波栗ペースト 2kg', unit: 'パック' },
    { name: '丹波栗 ボイルドマロン 1kg', unit: 'パック' },
    { name: '丹波栗渋皮煮 1号缶（固形物1.9kg）', unit: '缶' },
    { name: '丹波栗 甘露煮 1号缶（固形物1.9kg）', unit: '缶' }
  ],
  // 注文発送管理アプリの商品名 → 栗在庫の対応（kind: ship=生栗 / peelShip=渋皮むき栗 / goods=加工品）
  orderMap: {
    '丹波栗 生栗 3L': { kind: 'ship', size: '3L以上' },
    '丹波栗 生栗 4L': { kind: 'ship', size: '3L以上' },
    '丹波栗 生栗 2L': { kind: 'ship', size: '2L' },
    '丹波栗 生栗 L':  { kind: 'ship', size: 'L' },
    '丹波栗 3L 鬼皮剥き栗（渋皮残し）': { kind: 'peelShip', field: 'good' },
    '丹波栗 3L 鬼皮剥き栗（渋皮傷入り）': { kind: 'peelShip', field: 'bad' },
    '丹波栗ペースト 500g': { kind: 'goods' },
    '丹波栗ペースト 1kg':  { kind: 'goods' },
    '丹波栗ペースト 2kg':  { kind: 'goods' },
    '丹波栗 ボイルドマロン 1kg': { kind: 'goods' },
    '丹波栗渋皮煮 1号缶（固形物1.9kg）': { kind: 'goods' },
    '丹波栗 甘露煮 1号缶（固形物1.9kg）': { kind: 'goods' }
  },
  schedule:  {}
};

var ARR_KEYS = ['fields', 'sizes', 'sups', 'mems', 'tasks', 'workers'];
var OBJ_KEYS = ['prices', 'peelPrice', 'supPrices', 'orderMap', 'schedule'];
var LIST_KEYS = ['goodsItems'];   // オブジェクト配列の設定

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
    if (action === 'orderShip')    return json({ ok: true, data: orderShip(body) });
    if (action === 'orderCancel')  return json({ ok: true, data: orderCancel(body.orderId) });

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
  Object.keys(SHEETS).forEach(function (kind) {
    sheet_(ss, SHEETS[kind], headersFor(kind, st.sizes));
  });
  return 'OK: ' + ss.getName();
}

/* ===================== 処理本体 ===================== */

function getAll() {
  var ss = ss_();
  var st = readSettings(ss);
  return {
    fields:    st.fields,
    sizes:     st.sizes,
    sups:      st.sups,
    mems:      st.mems,
    tasks:     st.tasks,
    workers:   st.workers,
    prices:    st.prices,
    peelPrice: st.peelPrice,
    supPrices: st.supPrices,
    goodsItems: st.goodsItems,
    orderMap:  st.orderMap,
    schedule:  st.schedule,
    harvest:   readRecords(ss, 'harvest', st.sizes),
    ship:      readRecords(ss, 'ship',    st.sizes),
    buy:       readRecords(ss, 'buy',     st.sizes),
    peel:      readRecords(ss, 'peel',    st.sizes),
    peelOut:   readRecords(ss, 'peelOut', st.sizes),
    peelShip:  readRecords(ss, 'peelShip',st.sizes),
    goods:     readRecords(ss, 'goods',   st.sizes),
    syncLog:   readSyncLog(ss)
  };
}

function addRecord(kind, rec) {
  if (!SHEETS[kind])   throw new Error('unknown kind: ' + kind);
  if (!rec || !rec.id) throw new Error('record.id が必要です');

  var ss = ss_();
  var st = readSettings(ss);
  // 設定にないサイズがレコードに含まれていても取りこぼさない（peelは無関係）
  var sizes = (kind === 'peel' || kind === 'peelShip' || kind === 'goods')
    ? st.sizes
    : unique(st.sizes.concat(Object.keys(rec.q || {})));

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
  var st = {};
  ARR_KEYS.forEach(function (k) {
    st[k] = Array.isArray(body[k]) ? body[k].map(String) : cur[k];
  });
  OBJ_KEYS.forEach(function (k) {
    st[k] = (body[k] && typeof body[k] === 'object' && !Array.isArray(body[k]))
              ? body[k] : cur[k];
  });
  LIST_KEYS.forEach(function (k) {
    st[k] = Array.isArray(body[k]) ? body[k] : cur[k];
  });
  st.schedule = pruneSchedule_(st.schedule);
  writeSettings(sheet_(ss, SETTING_SHEET, ['key', 'value']), st);
  return st;
}

/** 作業予定は1セルにJSONで保存するため（上限5万文字）、空の日と2年以上前の日を落として肥大化を防ぐ */
function pruneSchedule_(sc) {
  if (!sc || typeof sc !== 'object') return {};
  var out = {}, limit = Utilities.formatDate(new Date(Date.now() - 2 * 366 * 86400000), tz_(), 'yyyy-MM-dd');
  Object.keys(sc).forEach(function (d) {
    var r = sc[d];
    if (!r || typeof r !== 'object') return;
    if (/^\d{4}-\d{2}-\d{2}$/.test(d) && d < limit) return;
    var a = r.a && typeof r.a === 'object' ? r.a : {};
    var has = !!(r.memo && String(r.memo).trim());
    Object.keys(a).forEach(function (t) { if (Array.isArray(a[t]) && a[t].length) has = true; });
    if (has) out[d] = r;
  });
  return out;
}

/* ===================== 注文発送管理アプリからの連携 ===================== */

/**
 * 注文が「発送済」になったときに呼ばれる。
 * body: { orderId, date, dest, lines:[{name, qty, unit}], memo }
 * 同じ orderId の既存レコードを消してから登録し直すので、何度呼んでも二重にならない。
 */
function orderShip(body) {
  var orderId = String(body.orderId || '').trim();
  if (!orderId) throw new Error('orderId が必要です');
  var st   = readSettings(ss_());
  var map  = st.orderMap || {};
  var date = String(body.date || Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'));
  var dest = String(body.dest || '');
  var memo = '注文No.' + orderId + '（自動）' + (body.memo ? ' ' + body.memo : '');
  var pre  = 'ord_' + orderId + '_';

  var removed = orderCancel(orderId).deleted;

  var q = {}, peel = { good: 0, bad: 0 }, goods = [], unknown = [];
  (body.lines || []).forEach(function (l) {
    var name = String(l.name || '').trim(), qty = num(l.qty);
    if (!name || qty <= 0) return;
    var m = map[name] || guessLine_(name, st);
    if (!m) { if (!/送料|運賃|手数料/.test(name)) unknown.push(name); return; }
    if (m.kind === 'ship')          q[m.size] = (q[m.size] || 0) + qty;
    else if (m.kind === 'peelShip') peel[m.field === 'bad' ? 'bad' : 'good'] += qty;
    else if (m.kind === 'goods')    goods.push({ item: m.item || name, qty: qty });
  });

  var added = [];
  if (Object.keys(q).length) {
    addRecord('ship', { id: pre + 'ship', date: date, dest: dest, q: q, price: null, memo: memo });
    added.push('ship');
  }
  if (peel.good > 0 || peel.bad > 0) {
    addRecord('peelShip', { id: pre + 'peel', date: date, dest: dest, good: peel.good, bad: peel.bad, price: null, memo: memo });
    added.push('peelShip');
  }
  goods.forEach(function (g, i) {
    addRecord('goods', { id: pre + 'g' + i, date: date, item: g.item, dir: 'out', qty: g.qty, memo: memo + ' → ' + dest });
    added.push('goods:' + g.item);
  });
  writeSyncLog_(ss_(), orderId, dest, added, unknown);
  return { orderId: orderId, removed: removed, added: added, unknown: unknown };
}

/**
 * 商品名が orderMap に無いとき、名前の中の言葉から在庫の種類を推定する。
 * 注文アプリでは商品名を自由入力できるため（例：「P420000 生栗 2Lサイズ」「丹波栗　3L」）、
 * 明らかなものだけ拾い、迷うものは unknown にして「連携ログ」と画面の注意表示に残す。
 */
function guessLine_(name, st) {
  var n = String(name).replace(/[\s　]+/g, '').replace(/[Ａ-Ｚａ-ｚ０-９]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); }).toUpperCase();
  n = n.replace(/^P\d{5,}/, '');                    // 商品コード（P420000 など）を外す

  // 加工品：登録済みの商品名と「ペースト＋容量」などで照合
  var items = st.goodsItems || [];
  for (var i = 0; i < items.length; i++) {
    var gn = String(items[i].name).replace(/[\s　]+/g, '').toUpperCase();
    if (n === gn) return { kind: 'goods', item: items[i].name };
  }
  if (/ペースト/.test(n)) {
    var cap = (n.match(/(\d+(?:\.\d+)?)(KG|G)/) || [])[0];
    if (cap) {
      for (var j = 0; j < items.length; j++) {
        var gj = String(items[j].name).replace(/[\s　]+/g, '').toUpperCase();
        if (/ペースト/.test(gj) && gj.indexOf(cap) >= 0) return { kind: 'goods', item: items[j].name };
      }
    }
    return null;   // 容量不明のペーストは判断しない
  }
  if (/渋皮煮|甘露煮|ボイルドマロン/.test(n)) {
    var kw = /渋皮煮/.test(n) ? '渋皮煮' : /甘露煮/.test(n) ? '甘露煮' : 'ボイルドマロン';
    if (/瓶|小瓶|ビン/.test(n)) return null;           // 瓶詰は在庫管理の対象外（判断しない）
    if (kw === 'ボイルドマロン' || /1号|１号|一号/.test(n)) {
      for (var k = 0; k < items.length; k++) {
        var gk = String(items[k].name);
        if (gk.indexOf(kw) >= 0 && (kw === 'ボイルドマロン' || /1号/.test(gk))) return { kind: 'goods', item: gk };
      }
    }
    return null;
  }

  // 渋皮むき栗（鬼皮剥き）
  if (/鬼皮|むき栗|剥き栗|剝き栗/.test(n) && !/渋皮煮/.test(n)) {
    return { kind: 'peelShip', field: /傷/.test(n) ? 'bad' : 'good' };
  }

  // 生栗：サイズ表記があるもの（殻付き・渋皮付・皮付き・生栗・熟成栗 もすべて生栗の在庫から出る）
  var sz = n.match(/(^|[^0-9A-Z])([2-5]L|L|M)(?![A-Z])/);
  if (sz && !/ペースト|煮|瓶|缶/.test(n)) {
    var size = sz[2];
    var sizes = st.sizes || [];
    if (/^[3-5]L$/.test(size)) {                       // 3L・4L・5L は「3L以上」へ
      for (var a = 0; a < sizes.length; a++) if (/^3L/.test(sizes[a])) return { kind: 'ship', size: sizes[a] };
    }
    if (sizes.indexOf(size) >= 0) return { kind: 'ship', size: size };
  }
  return null;
}

function writeSyncLog_(ss, orderId, dest, added, unknown) {
  var sh = sheet_(ss, LOG_SHEET, ['日時', '注文ID', '出荷先', '登録した在庫', '対応づけできなかった商品名']);
  // 同じ注文IDの古い行は消して最新だけ残す（名前を直して再登録したら注意表示も消えるように）
  var vals = sh.getDataRange().getValues();
  for (var r = vals.length - 1; r >= 1; r--) if (String(vals[r][1]) === orderId) sh.deleteRow(r + 1);
  sh.appendRow([Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm:ss'), orderId, dest, added.join(', '), unknown.join(', ')]);
  var last = sh.getLastRow();
  if (last > 501) sh.deleteRows(2, last - 501);        // 直近500件だけ残す
}

/** 対応づけできなかった商品名がある連携だけ、直近90日分を返す（画面の注意表示に使う） */
function readSyncLog(ss) {
  var sh = ss.getSheetByName(LOG_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  var vals = sh.getDataRange().getValues(), out = [];
  var limit = new Date(Date.now() - 90 * 86400000);
  for (var r = 1; r < vals.length; r++) {
    var un = String(vals[r][4] || '').trim();
    if (!un) continue;
    var at = vals[r][0] instanceof Date ? vals[r][0] : new Date(String(vals[r][0]));
    if (isFinite(at) && at < limit) continue;
    out.push({ at: dstr(vals[r][0]), orderId: String(vals[r][1]), dest: String(vals[r][2] || ''), unknown: un.split(/,\s*/) });
  }
  return out;
}

/** 注文の取消・発送済の取り下げ・削除時に呼ばれる。orderId に紐づくレコードを全部消す */
function orderCancel(orderId) {
  orderId = String(orderId || '').trim();
  if (!orderId) throw new Error('orderId が必要です');
  var pre = 'ord_' + orderId + '_', ss = ss_(), deleted = 0;
  ['ship', 'peelShip', 'goods'].forEach(function (kind) {
    var sh = ss.getSheetByName(SHEETS[kind]);
    if (!sh || sh.getLastRow() < 2) return;
    var vals = sh.getDataRange().getValues();
    var i = headRow(vals[0]).indexOf('id');
    if (i < 0) return;
    for (var r = vals.length - 1; r >= 1; r--) {
      if (String(vals[r][i]).indexOf(pre) === 0) { sh.deleteRow(r + 1); deleted++; }
    }
  });
  return { orderId: orderId, deleted: deleted };
}

/* ===================== シート入出力 ===================== */

function headersFor(kind, sizes) {
  if (kind === 'harvest') {
    return ['id', '日付', '園地'].concat(sizes).concat(['廃棄kg', '合計kg', 'メモ', '登録日時']);
  }
  if (kind === 'ship') {
    return ['id', '日付', '出荷先'].concat(sizes).concat(['合計kg', '単価', 'メモ', '登録日時']);
  }
  if (kind === 'peel') {
    return ['id', '日付', '作業者', '良品kg', '良品単価', '傷ありkg', '傷あり単価', '合計kg', '金額', 'メモ', '登録日時'];
  }
  if (kind === 'peelOut') {
    return ['id', '日付', '作業者'].concat(sizes).concat(['kg', 'メモ', '登録日時']);
  }
  if (kind === 'peelShip') {
    return ['id', '日付', '出荷先', '良品kg', '傷ありkg', '合計kg', '単価', '金額', 'メモ', '登録日時'];
  }
  if (kind === 'goods') {
    return ['id', '日付', '商品', '区分', '数量', 'メモ', '登録日時'];
  }
  return ['id', '日付', '買取先']
    .concat(sizes.map(function (s) { return s + '_kg'; }))
    .concat(sizes.map(function (s) { return s + '_単価'; }))
    .concat(['廃棄kg', '合計kg', '金額', 'メモ', '登録日時']);
}

function buildRowMap(kind, rec, sizes) {
  var m = {};
  m['id']       = String(rec.id);
  m['日付']     = String(rec.date || '');
  m['メモ']     = String(rec.memo || '');
  m['登録日時'] = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm:ss');

  if (kind === 'peel') {
    var g = num(rec.good), b = num(rec.bad), gp = num(rec.gp), bp = num(rec.bp);
    m['作業者']     = String(rec.who || '');
    m['良品kg']     = g;
    m['良品単価']   = gp;
    m['傷ありkg']   = b;
    m['傷あり単価'] = bp;
    m['合計kg']     = Math.round((g + b) * 10) / 10;
    m['金額']       = Math.round(g * gp + b * bp);
    return m;
  }
  if (kind === 'peelOut') {
    var oq = rec.q || {}, ot = 0;
    m['作業者'] = String(rec.who || '');
    sizes.forEach(function (s) { var w = num(oq[s]); m[s] = w; ot += w; });
    m['kg']     = ot > 0 ? Math.round(ot * 10) / 10 : num(rec.kg);
    return m;
  }
  if (kind === 'goods') {
    m['商品'] = String(rec.item || '');
    m['区分'] = rec.dir === 'out' ? '出庫' : '入庫';
    m['数量'] = num(rec.qty);
    return m;
  }
  if (kind === 'peelShip') {
    var sg = num(rec.good), sb = num(rec.bad);
    var hasP = !(rec.price === null || rec.price === undefined || rec.price === '');
    m['出荷先']   = String(rec.dest || '');
    m['良品kg']   = sg;
    m['傷ありkg'] = sb;
    m['合計kg']   = Math.round((sg + sb) * 10) / 10;
    m['単価']     = hasP ? num(rec.price) : '';
    m['金額']     = hasP ? Math.round((sg + sb) * num(rec.price)) : '';
    return m;
  }

  var q = rec.q || {}, total = 0;

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
    m['廃棄kg'] = num(rec.waste);   // 買取の廃棄は支払・在庫の対象外
  } else {
    sizes.forEach(function (s) { var w = num(q[s]); m[s] = w; total += w; });
    if (kind === 'harvest') {
      m['園地']   = String(rec.field || '');
      m['廃棄kg'] = num(rec.waste);   // 廃棄は在庫・出荷可能量には含めない
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

    if (kind === 'peel') {
      rec.who  = String(col(row, '作業者') || '');
      rec.good = num(col(row, '良品kg'));
      rec.bad  = num(col(row, '傷ありkg'));
      rec.gp   = num(col(row, '良品単価'));
      rec.bp   = num(col(row, '傷あり単価'));
      out.push(rec);
      continue;
    }
    if (kind === 'peelOut') {
      rec.who = String(col(row, '作業者') || '');
      var oq2 = {};
      sizes.forEach(function (s) { oq2[s] = num(col(row, s)); });
      rec.q  = oq2;
      rec.kg = num(col(row, 'kg'));
      out.push(rec);
      continue;
    }
    if (kind === 'goods') {
      rec.item = String(col(row, '商品') || '');
      rec.dir  = String(col(row, '区分') || '') === '出庫' ? 'out' : 'in';
      rec.qty  = num(col(row, '数量'));
      out.push(rec);
      continue;
    }
    if (kind === 'peelShip') {
      rec.dest = String(col(row, '出荷先') || '');
      rec.good = num(col(row, '良品kg'));
      rec.bad  = num(col(row, '傷ありkg'));
      var spv = col(row, '単価');
      rec.price = (spv === '' || spv === null || spv === undefined) ? null : num(spv);
      out.push(rec);
      continue;
    }

    var q = {};
    if (kind === 'buy') {
      var p = {};
      sizes.forEach(function (s) { q[s] = num(col(row, s + '_kg')); p[s] = num(col(row, s + '_単価')); });
      rec.q = q; rec.p = p;
      rec.sup = String(col(row, '買取先') || '');
      rec.waste = num(col(row, '廃棄kg'));
    } else {
      sizes.forEach(function (s) { q[s] = num(col(row, s)); });
      rec.q = q;
      if (kind === 'harvest') {
        rec.field = String(col(row, '園地') || '');
        rec.waste = num(col(row, '廃棄kg'));
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
  ARR_KEYS.forEach(function (k) {
    st[k] = Array.isArray(raw[k]) ? raw[k].map(String) : DEFAULTS[k].slice();
  });
  OBJ_KEYS.forEach(function (k) {
    st[k] = (raw[k] && typeof raw[k] === 'object' && !Array.isArray(raw[k]))
              ? raw[k] : copy(DEFAULTS[k]);
  });
  LIST_KEYS.forEach(function (k) {
    st[k] = (Array.isArray(raw[k]) && raw[k].length) ? raw[k] : copy(DEFAULTS[k]);
  });

  if (vals.length < 2) writeSettings(sh, st);   // 初回は既定値を書き込む
  return st;
}

function writeSettings(sh, st) {
  var rows = ARR_KEYS.concat(OBJ_KEYS).concat(LIST_KEYS).map(function (k) {
    return [k, JSON.stringify(st[k])];
  });
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
