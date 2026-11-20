/**
 * 菜品銷售紀錄系統 — Google 雲端同步後端（Apps Script，綁定在試算表上）
 *
 * 試算表分頁：菜單、銷售（每天每道菜一列）、備註。可直接在試算表查看；手動修改也會同步回 App。
 * 每列有「更新時間」(毫秒) 與「版本」：App 以更新時間決定新舊，以版本只下載有變動的列。
 * 密碼：第一次從 App 連線（ping）時設定，存成 SHA-256 雜湊在指令碼屬性 PW_HASH。
 */

const TABLES = {
  dishes: {
    sheet: '菜單',
    head: ['ID', '分類', '菜名', '售價', '成本', '停售', '已刪除', '排序', '更新時間', '版本'],
    text: [1, 2, 3],
    key: (r) => r.id,
    valid: (r) => !!r.id,
    toRow: (r, rev) => [r.id, str_(r.cat), str_(r.name), num_(r.price), num_(r.cost), r.off ? '是' : '', r.deleted ? '是' : '', num_(r.ord), num_(r.at), rev],
    fromRow: (a) => ({ id: str_(a[0]), cat: str_(a[1]), name: str_(a[2]), price: num_(a[3]), cost: num_(a[4]), off: yes_(a[5]), deleted: yes_(a[6]), ord: num_(a[7]), at: num_(a[8]) }),
  },
  sales: {
    sheet: '銷售',
    // 售完時間、輸入者（v2.3 起）加在最後面，舊的列不用搬動
    head: ['日期', '菜品ID', '分類', '菜名', '份數', '單價', '成本', '金額', '更新時間', '版本', '售完時間', '輸入者'],
    text: [1, 2, 3, 4, 11, 12],
    at: 8, ver: 9,
    key: (r) => r.d + '|' + r.id,
    valid: (r) => /^\d{4}-\d{2}-\d{2}$/.test(r.d) && !!r.id,
    toRow: (r, rev, dish) => {
      const d = dish[r.id] || {};
      return [r.d, r.id, str_(d.cat), str_(d.name), num_(r.q), num_(r.p), num_(r.c), num_(r.q) * num_(r.p), num_(r.at), rev, str_(r.so), str_(r.by)];
    },
    fromRow: (a) => ({ d: ymd_(a[0]), id: str_(a[1]), q: num_(a[4]), p: num_(a[5]), c: num_(a[6]), at: num_(a[8]), so: hm_(a[10]), by: str_(a[11]) }),
  },
  notes: {
    sheet: '備註',
    head: ['日期', '備註', '更新時間', '版本'],
    text: [1, 2],
    key: (r) => r.d,
    valid: (r) => /^\d{4}-\d{2}-\d{2}$/.test(r.d),
    toRow: (r, rev) => [r.d, str_(r.note), num_(r.at), rev],
    fromRow: (a) => ({ d: ymd_(a[0]), note: str_(a[1]), at: num_(a[2]) }),
  },
};
const ORDER = ['dishes', 'sales', 'notes'];
// 更新時間、版本在第幾欄（從 0 算）；沒寫就是最後兩欄
const atCol_ = (t) => (t.at != null ? t.at : t.head.length - 2);
const verCol_ = (t) => (t.ver != null ? t.ver : t.head.length - 1);

function str_(v) { return v == null ? '' : String(v); }
function num_(v) { const n = Number(v); return isFinite(n) ? n : 0; }
function yes_(v) { return v === true || v === '是' || v === 'TRUE' || v === 'true' || v === 1; }
function ymd_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const s = str_(v).trim().replace(/\//g, '-');
  const m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  return m ? m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2) : s;
}
// 售完時間：試算表可能把 11:30 轉成時間格式
function hm_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm');
  return str_(v).trim();
}
function props_() { return PropertiesService.getScriptProperties(); }
function rev_() { return num_(props_().getProperty('REV')); }
function hash_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'dishsales:' + s, Utilities.Charset.UTF_8)
    .map((b) => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function sheet_(t) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(t.sheet);
  if (!sh) {
    sh = ss.insertSheet(t.sheet);
    sh.getRange(1, 1, 1, t.head.length).setValues([t.head]).setFontWeight('bold');
    sh.setFrozenRows(1);
  } else if (sh.getLastColumn() < t.head.length) {
    // 舊版建立的分頁：補上新增欄位的標題
    sh.getRange(1, 1, 1, t.head.length).setValues([t.head]).setFontWeight('bold');
  }
  return sh;
}
function read_(t) {
  const sh = sheet_(t), n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, 1, n, t.head.length).getValues().filter((a) => str_(a[0]) !== '') : [];
}
function write_(t, rows) {
  const sh = sheet_(t), w = t.head.length, last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, w).clearContent();
  if (!rows.length) return;
  // 文字欄先設成純文字，避免 = + 開頭被當公式、日期被轉成日期格式
  t.text.forEach((c) => sh.getRange(2, c, rows.length, 1).setNumberFormat('@'));
  sh.getRange(2, 1, rows.length, w).setValues(rows);
}

function doGet() {
  return ContentService.createTextOutput('菜品銷售紀錄系統 雲端服務運作中');
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, err: '要求格式錯誤' }); }
  const p = props_();
  const stored = p.getProperty('PW_HASH');
  const pw = str_(req.pw);
  if (!pw) return json_({ ok: false, err: '請輸入密碼' });
  let claimed = false;
  if (!stored) {
    if (req.a !== 'ping') return json_({ ok: false, err: '雲端尚未設定密碼，請重新連線' });
    if (pw.length < 4) return json_({ ok: false, err: '密碼至少 4 個字' });
    p.setProperty('PW_HASH', hash_(pw));
    claimed = true;
  } else if (hash_(pw) !== stored) {
    return json_({ ok: false, err: '密碼錯誤', auth: true });
  }

  if (req.a === 'ping') {
    return json_({
      ok: true, claimed: claimed, rev: rev_(), sheet: SpreadsheetApp.getActive().getUrl(),
      counts: { dishes: read_(TABLES.dishes).length, sales: read_(TABLES.sales).filter((a) => num_(a[4]) > 0).length },
    });
  }
  if (req.a === 'rev') return json_({ ok: true, rev: rev_() });
  if (req.a === 'setpw') {
    const np = str_(req.newpw);
    if (np.length < 4) return json_({ ok: false, err: '新密碼至少 4 個字' });
    p.setProperty('PW_HASH', hash_(np));
    return json_({ ok: true });
  }
  if (req.a === 'sync') {
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try { return json_(sync_(req)); } finally { lock.releaseLock(); }
  }
  return json_({ ok: false, err: '未知的動作' });
}

function sync_(req) {
  const p = props_();
  const ch = req.changes || {};
  const cur = rev_(), next = cur + 1;
  const data = {}, changed = {};
  for (const name of ORDER) {
    const t = TABLES[name];
    const rows = read_(t);
    const idx = {};
    rows.forEach((a, i) => { idx[t.key(t.fromRow(a))] = i; });
    data[name] = { rows: rows, idx: idx };
  }
  const dishMap = () => {
    const m = {};
    data.dishes.rows.forEach((a) => { m[str_(a[0])] = { cat: a[1], name: a[2], ord: num_(a[7]) }; });
    return m;
  };
  let dish = dishMap();
  for (const name of ORDER) {
    const t = TABLES[name], D = data[name], atCol = atCol_(t);
    for (const r of ch[name] || []) {
      if (!r || !t.valid(r)) continue;
      const k = t.key(r), i = D.idx[k];
      if (i === undefined) { D.idx[k] = D.rows.length; D.rows.push(t.toRow(r, next, dish)); changed[name] = true; }
      else if (num_(r.at) > num_(D.rows[i][atCol])) { D.rows[i] = t.toRow(r, next, dish); changed[name] = true; }
    }
    if (name === 'dishes' && changed.dishes) dish = dishMap();
  }
  if (changed.dishes) {
    // 菜名或分類改了，銷售分頁的名稱一起更新（只改顯示欄，不動版本）
    data.sales.rows.forEach((a) => {
      const d = dish[str_(a[1])];
      if (d && (a[2] !== d.cat || a[3] !== d.name)) { a[2] = d.cat; a[3] = d.name; changed.salesNames = true; }
    });
  }
  if (ch.shop && num_(ch.shop.at) > num_(p.getProperty('SHOP_AT'))) {
    p.setProperties({ SHOP: str_(ch.shop.v), SHOP_AT: String(num_(ch.shop.at)), SHOP_REV: String(next) });
    changed.shop = true;
  }
  // 其他設定（鎖定過去日期等）整包存在指令碼屬性
  if (ch.settings && num_(ch.settings.at) > num_(p.getProperty('SETTINGS_AT'))) {
    p.setProperties({ SETTINGS: str_(ch.settings.v), SETTINGS_AT: String(num_(ch.settings.at)), SETTINGS_REV: String(next) });
    changed.settings = true;
  }
  const any = changed.dishes || changed.sales || changed.notes || changed.shop || changed.settings;
  if (changed.dishes) {
    data.dishes.rows.sort((a, b) => num_(a[7]) - num_(b[7]));
    write_(TABLES.dishes, data.dishes.rows);
  }
  if (changed.sales || changed.salesNames) {
    data.sales.rows.sort((a, b) => (ymd_(a[0]) < ymd_(b[0]) ? -1 : ymd_(a[0]) > ymd_(b[0]) ? 1 : num_((dish[a[1]] || {}).ord) - num_((dish[b[1]] || {}).ord)));
    write_(TABLES.sales, data.sales.rows);
  }
  if (changed.notes) {
    data.notes.rows.sort((a, b) => (ymd_(a[0]) < ymd_(b[0]) ? -1 : 1));
    write_(TABLES.notes, data.notes.rows);
  }
  const rev = any ? next : cur;
  if (any) p.setProperty('REV', String(rev));

  // 回傳用戶端還沒有的列（版本 > since）；since 比目前版本還大（試算表被重設過）就全部回傳
  let since = num_(req.since);
  if (since > rev) since = 0;
  const out = { ok: true, rev: rev, full: since === 0, sheet: SpreadsheetApp.getActive().getUrl() };
  for (const name of ORDER) {
    const t = TABLES[name], vc = verCol_(t);
    out[name] = data[name].rows.filter((a) => num_(a[vc]) > since).map(t.fromRow);
  }
  if (num_(p.getProperty('SHOP_REV')) > since) out.shop = { v: str_(p.getProperty('SHOP')), at: num_(p.getProperty('SHOP_AT')) };
  if (num_(p.getProperty('SETTINGS_REV')) > since) out.settings = { v: str_(p.getProperty('SETTINGS')), at: num_(p.getProperty('SETTINGS_AT')) };
  return out;
}

/** 在試算表上手動修改時：更新該列的更新時間與版本，App 下次同步就會拿到 */
function onEdit(e) {
  try {
    const sh = e.range.getSheet();
    let t = null;
    for (const name of ORDER) if (TABLES[name].sheet === sh.getName()) t = TABLES[name];
    if (!t) return;
    const r0 = Math.max(2, e.range.getRow()), r1 = e.range.getRow() + e.range.getNumRows() - 1;
    if (r1 < r0) return;
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const rev = rev_() + 1, w = t.head.length, now = Date.now();
      const vals = sh.getRange(r0, 1, r1 - r0 + 1, w).getValues();
      vals.forEach((a) => {
        if (str_(a[0]) === '') return;
        if (t === TABLES.sales) { a[0] = ymd_(a[0]); a[7] = num_(a[4]) * num_(a[5]); }
        if (t === TABLES.notes) a[0] = ymd_(a[0]);
        a[atCol_(t)] = now; a[verCol_(t)] = rev;
      });
      t.text.forEach((c) => sh.getRange(r0, c, vals.length, 1).setNumberFormat('@'));
      sh.getRange(r0, 1, vals.length, w).setValues(vals);
      props_().setProperty('REV', String(rev));
    } finally { lock.releaseLock(); }
  } catch (err) {
    console.error(err);
  }
}

/** 在編輯器手動執行一次，完成授權並建立分頁 */
function setup() {
  ORDER.forEach((n) => sheet_(TABLES[n]));
  Logger.log('完成，目前版本 ' + rev_() + '，密碼' + (props_().getProperty('PW_HASH') ? '已設定' : '尚未設定（第一次從 App 連線時設定）'));
}

/** 忘記密碼時在編輯器執行：清除密碼，下次從 App 連線時重新設定 */
function resetPassword() {
  props_().deleteProperty('PW_HASH');
  Logger.log('已清除密碼');
}
