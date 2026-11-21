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
    // 售完時間、輸入者（v2.3 起）、備量（v2.5 起）加在最後面，舊的列不用搬動
    head: ['日期', '菜品ID', '分類', '菜名', '份數', '單價', '成本', '金額', '更新時間', '版本', '售完時間', '輸入者', '備量'],
    text: [1, 2, 3, 4, 11, 12],
    at: 8, ver: 9,
    key: (r) => r.d + '|' + r.id,
    valid: (r) => /^\d{4}-\d{2}-\d{2}$/.test(r.d) && !!r.id,
    toRow: (r, rev, dish) => {
      const d = dish[r.id] || {};
      return [r.d, r.id, str_(d.cat), str_(d.name), num_(r.q), num_(r.p), num_(r.c), num_(r.q) * num_(r.p), num_(r.at), rev, str_(r.so), str_(r.by), num_(r.m) || ''];
    },
    fromRow: (a) => ({ d: ymd_(a[0]), id: str_(a[1]), q: num_(a[4]), p: num_(a[5]), c: num_(a[6]), at: num_(a[8]), so: hm_(a[10]), by: str_(a[11]), m: num_(a[12]) }),
  },
  notes: {
    sheet: '備註',
    // 特殊日標註（v2.5 起，例如「雨天,連假」）加在最後面
    head: ['日期', '備註', '更新時間', '版本', '標註'],
    text: [1, 2, 5],
    at: 2, ver: 3,
    key: (r) => r.d,
    valid: (r) => /^\d{4}-\d{2}-\d{2}$/.test(r.d),
    toRow: (r, rev) => [r.d, str_(r.note), num_(r.at), rev, str_(r.tags)],
    fromRow: (a) => ({ d: ymd_(a[0]), note: str_(a[1]), at: num_(a[2]), tags: str_(a[4]) }),
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
  if (req.a === 'report') {
    // 立刻寄一封上週的週報（測試用）
    try { return json_({ ok: true, to: sendWeeklyReport_(true) }); } catch (err) { return json_({ ok: false, err: reportErr_(err) }); }
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
  // 其他設定（鎖定、目標、週報）存在指令碼屬性；每一項各自比新舊（v2.5 起每項有自己的 at，舊版只送整包的 at）
  if (ch.settings && num_(ch.settings.at) > 0) {
    let cur = {}, inc = {};
    try { cur = JSON.parse(p.getProperty('SETTINGS') || '{}') || {}; } catch (e) { cur = {}; }
    try { inc = JSON.parse(str_(ch.settings.v) || '{}') || {}; } catch (e) { inc = {}; }
    const curAt = num_(p.getProperty('SETTINGS_AT'));
    let upd = false;
    Object.keys(inc).forEach((k) => {
      if (!inc[k] || typeof inc[k] !== 'object') return;
      const ia = inc[k].at != null ? num_(inc[k].at) : num_(ch.settings.at), ca = cur[k] ? num_(cur[k].at) || curAt : 0;
      if (ia > ca) { cur[k] = Object.assign({}, inc[k], { at: ia }); upd = true; }
    });
    if (upd) {
      p.setProperties({ SETTINGS: JSON.stringify(cur), SETTINGS_AT: String(Math.max(curAt, num_(ch.settings.at))), SETTINGS_REV: String(next) });
      changed.settings = true;
      if (inc.report) ensureReportTrigger_();
    }
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

// ------------------------------------------------------------------ 自動週報（v2.5）
// App 的「自動週報」設定存在 SETTINGS（JSON 的 report: {on, email}）；每週一早上 8 點由時間觸發寄出上週摘要。
// 第一次使用要在編輯器執行一次 setupReport()，授權寄信和建立觸發條件。
function settings_() { try { return JSON.parse(props_().getProperty('SETTINGS') || '{}'); } catch (e) { return {}; } }
function reportErr_(err) {
  const m = String(err && err.message || err);
  return /permission|authoriz|權限|授權|required/i.test(m) ? '後端還沒授權寄信：請在 Apps Script 編輯器執行一次 setupReport()' : m;
}
function ensureReportTrigger_() {
  try {
    const on = !!(settings_().report || {}).on;
    const has = ScriptApp.getProjectTriggers().filter((t) => t.getHandlerFunction() === 'weeklyReport');
    if (on && !has.length) ScriptApp.newTrigger('weeklyReport').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
    if (!on) has.forEach((t) => ScriptApp.deleteTrigger(t));
  } catch (e) { console.warn('觸發條件：' + e); } // 還沒授權時略過，執行 setupReport() 後就會建立
}
/** 在編輯器手動執行一次：授權寄信、建立每週一的觸發條件 */
function setupReport() {
  // Google 授權畫面可以只勾部分權限；這行會把還沒給的權限再要一次（寄信、每週自動執行）
  ScriptApp.requireAllScopes(ScriptApp.AuthMode.FULL);
  setup();
  MailApp.getRemainingDailyQuota();
  ensureReportTrigger_();
  const r = settings_().report || {};
  Logger.log(r.on ? '完成：每週一 8 點寄到 ' + r.email : '完成授權。在 App 的「菜品與設定 → 自動週報」開啟後就會每週寄出');
}
function weeklyReport() { sendWeeklyReport_(false); }

function sendWeeklyReport_(force) {
  const r = settings_().report || {};
  if (!r.email) throw new Error('App 還沒設定週報要寄到哪個 Email');
  if (!r.on && !force) return '';
  const tz = Session.getScriptTimeZone();
  const day = (d) => Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  const t = new Date(); t.setHours(12, 0, 0, 0);
  const back = (t.getDay() + 6) % 7; // 距離本週一幾天
  const mon = new Date(t); mon.setDate(t.getDate() - back - 7);
  const ks = [], pk = [];
  for (let i = 0; i < 7; i++) {
    const d = new Date(mon); d.setDate(mon.getDate() + i); ks.push(day(d));
    const e = new Date(mon); e.setDate(mon.getDate() + i - 7); pk.push(day(e));
  }
  const dish = {};
  read_(TABLES.dishes).forEach((a) => { dish[str_(a[0])] = { name: str_(a[2]), cat: str_(a[1]) }; });
  const byDay = {}, prevDay = {}, byDish = {}, prevDish = {};
  let so = 0, made = 0, left = 0;
  read_(TABLES.sales).forEach((a) => {
    const k = ymd_(a[0]), id = str_(a[1]), q = num_(a[4]), m = num_(a[12]);
    if (ks.indexOf(k) >= 0) {
      byDay[k] = (byDay[k] || 0) + q; if (q) byDish[id] = (byDish[id] || 0) + q;
      if (q && str_(a[10])) so++;
      if (m) { made += m; left += Math.max(0, m - q); }
    } else if (pk.indexOf(k) >= 0) { prevDay[k] = (prevDay[k] || 0) + q; if (q) prevDish[id] = (prevDish[id] || 0) + q; }
  });
  const sum = (o) => Object.keys(o).reduce((s, k) => s + o[k], 0);
  const tot = sum(byDay), ptot = sum(prevDay), open = ks.filter((k) => byDay[k] > 0).length;
  const pct = (a, b) => (b ? (a >= b ? '↑' : '↓') + Math.abs(Math.round((a / b - 1) * 100)) + '%' : '—');
  const nm = (id) => (dish[id] ? dish[id].name : '（已移除）');
  const top = Object.keys(byDish).sort((a, b) => byDish[b] - byDish[a]);
  const down = Object.keys(prevDish).filter((id) => prevDish[id] >= 10 && (byDish[id] || 0) <= prevDish[id] * 0.7)
    .sort((a, b) => (byDish[a] || 0) / prevDish[a] - (byDish[b] || 0) / prevDish[b]).slice(0, 5);
  const shop = str_(props_().getProperty('SHOP')) || '菜品銷售紀錄';
  const WD = '日一二三四五六';
  const esc = (x) => str_(x).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const td = 'padding:7px 10px;border-bottom:1px solid #eee;';
  const kpi = (l, v, s) => '<td style="padding:10px 14px;border:1px solid #e7e0d5;border-radius:10px;vertical-align:top"><div style="color:#8a8174;font-size:12px;font-weight:700">' + l
    + '</div><div style="font-size:22px;font-weight:800;margin-top:2px">' + v + '</div><div style="color:#666;font-size:12px">' + (s || '') + '</div></td>';
  const wdOf = (k) => WD[new Date(k + 'T12:00:00').getDay()];
  const html = '<div style="font-family:\'Microsoft JhengHei\',\'PingFang TC\',sans-serif;color:#1f1a14;max-width:640px">'
    + '<div style="border-bottom:3px solid #d9480f;padding-bottom:8px;margin-bottom:14px"><div style="color:#8a8174;font-weight:700">' + esc(shop) + '</div>'
    + '<h2 style="margin:2px 0 0">菜品銷售週報</h2><div style="color:#8a8174;font-size:13px">' + ks[0].replace(/-/g, '/') + '（一）～ ' + ks[6].replace(/-/g, '/') + '（日）</div></div>'
    + '<table style="border-collapse:separate;border-spacing:8px 0;margin:0 -8px 16px"><tr>'
    + kpi('上週總份數', tot.toLocaleString() + ' 份', '比前一週 ' + pct(tot, ptot)) + kpi('營業天數', open + ' 天', open ? '平均每天 ' + Math.round(tot / open) + ' 份' : '')
    + kpi('賣完次數', so + ' 次', '') + (made ? kpi('報廢率', Math.round(left / made * 100) + '%', '剩 ' + left + ' 份') : '') + '</tr></table>'
    + '<h3 style="margin:16px 0 6px;color:#d9480f">每天份數</h3><table style="border-collapse:collapse;width:100%">'
    + ks.map((k, i) => '<tr><td style="' + td + '">' + k.slice(5).replace('-', '/') + '（' + wdOf(k) + '）</td><td style="' + td + 'text-align:right;font-weight:700">' + (byDay[k] || '—')
      + '</td><td style="' + td + 'text-align:right;color:#8a8174">前一週 ' + (prevDay[pk[i]] || '—') + '</td></tr>').join('') + '</table>'
    + '<h3 style="margin:16px 0 6px;color:#d9480f">賣最多前 5 名</h3><table style="border-collapse:collapse;width:100%">'
    + (top.slice(0, 5).map((id, i) => '<tr><td style="' + td + '">' + (i + 1) + '. ' + esc(nm(id)) + '</td><td style="' + td + 'text-align:right;font-weight:700">' + byDish[id]
      + ' 份</td><td style="' + td + 'text-align:right;color:#8a8174">' + pct(byDish[id], prevDish[id] || 0) + '</td></tr>').join('') || '<tr><td style="' + td + '">上週沒有紀錄</td></tr>') + '</table>'
    + (down.length ? '<h3 style="margin:16px 0 6px;color:#b91c1c">⚠ 比前一週明顯下滑</h3><ul style="margin:0;padding-left:20px">'
      + down.map((id) => '<li>' + esc(nm(id)) + '：' + (byDish[id] || 0) + ' 份（前一週 ' + prevDish[id] + ' 份）</li>').join('') + '</ul>' : '')
    + '<p style="color:#999;font-size:12px;margin-top:20px">由菜品銷售紀錄系統自動寄出。要停止寄送，請到 App 的「菜品與設定 → 自動週報」關閉。</p></div>';
  MailApp.sendEmail({ to: r.email, subject: shop + ' 菜品銷售週報 ' + ks[0].slice(5).replace('-', '/') + '～' + ks[6].slice(5).replace('-', '/'), htmlBody: html, name: shop });
  return r.email;
}
