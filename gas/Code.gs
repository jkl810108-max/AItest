/**
 * 鍋淬 GUOCUI 訂單後端（Google Apps Script，綁定在 Google 試算表）
 *
 * 流程：網站 POST JSON -> doPost 驗證 -> 以「Menu」工作表重算價格 -> 寫入「Orders」「OrderItems」
 * 價格以試算表為準，瀏覽器送來的金額只用來比對，不會被信任。
 *
 * 首次使用：在編輯器執行一次 setup()，再部署為網頁應用程式（見 gas/README.md）。
 */

var SHEET_ORDERS = 'Orders';
var SHEET_ITEMS = 'OrderItems';
var SHEET_MENU = 'Menu';
var TZ = 'Asia/Taipei';
var FREE_SHIPPING_MIN = 1200;
var SHIPPING_FEE = 160;
var STORES = ['大安旗艦店', '信義概念店', '中山日常店'];
var MODES = { dinein: '內用', takeaway: '外帶', delivery: '冷凍宅配' };

var ORDER_HEADERS = ['訂單編號', '建立時間', '狀態', '取餐方式', '門市', '桌號', '姓名', '電話', 'Email', '地址', '備註', '商品小計', '運費', '訂單總額', 'requestId'];
var ITEM_HEADERS = ['訂單編號', '商品ID', '商品名稱', '單價', '數量', '小計'];
var MENU_HEADERS = ['id', '名稱', '分類', '價格', '單位', '冷凍', '供應中'];
var MENU_SEED = [
  ['p1', '經典鮮肉鍋貼', 'potstickers', 85, '8 入', false, true],
  ['p2', '韭香鮮肉鍋貼', 'potstickers', 90, '8 入', false, true],
  ['p3', '韓式泡菜鍋貼', 'potstickers', 100, '8 入', false, true],
  ['p4', '鮮蝦雙饗鍋貼', 'potstickers', 125, '8 入', false, true],
  ['d1', '高麗菜鮮肉水餃', 'dumplings', 85, '8 入', false, true],
  ['d2', '韭菜鮮肉水餃', 'dumplings', 90, '8 入', false, true],
  ['d3', '鮮蝦水餃', 'dumplings', 130, '8 入', false, true],
  ['d4', '田園蔬食水餃', 'dumplings', 95, '8 入', false, true],
  ['s1', '經典酸辣湯', 'drinks', 50, '1 碗', false, true],
  ['s2', '暖心玉米濃湯', 'drinks', 55, '1 碗', false, true],
  ['s3', '冷泡烏龍茶', 'drinks', 45, '1 杯', false, true],
  ['s4', '桂花紅茶', 'drinks', 50, '1 杯', false, true],
  ['a1', '爽脆小黃瓜', 'sides', 40, '1 盤', false, true],
  ['a2', '黃金泡菜', 'sides', 45, '1 盤', false, true],
  ['a3', '和風海帶芽', 'sides', 40, '1 盤', false, true],
  ['a4', '家常滷豆干', 'sides', 45, '1 盤', false, true],
  ['gift50', '招牌招財鍋貼禮盒', 'gifts', 680, '50 入 / 盒', true, true]
];

/** 首次執行：建立三張工作表、標題列與菜單。可重複執行，不會覆蓋既有資料。 */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, SHEET_ORDERS, ORDER_HEADERS, true);
  ensureSheet_(ss, SHEET_ITEMS, ITEM_HEADERS, true);
  var menu = ensureSheet_(ss, SHEET_MENU, MENU_HEADERS, false);
  if (menu.getLastRow() < 2) menu.getRange(2, 1, MENU_SEED.length, MENU_HEADERS.length).setValues(MENU_SEED);
  var first = ss.getSheetByName('工作表1') || ss.getSheetByName('Sheet1');
  if (first && ss.getSheets().length > 1 && first.getLastRow() === 0) ss.deleteSheet(first);
  Logger.log('完成。請到「專案設定 > 指令碼屬性」可選填 NOTIFY_EMAIL（新訂單通知信箱）。');
}

function ensureSheet_(ss, name, headers, textFormat) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#f3e5d8');
    sh.setFrozenRows(1);
  }
  // 純文字格式：避免姓名、備註等欄位被當成公式執行（試算表公式注入）。
  if (textFormat) sh.getRange(2, 1, 5000, headers.length).setNumberFormat('@');
  return sh;
}

function doGet() {
  return json_({ ok: true, service: 'guocui-orders' });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return json_({ ok: false, code: 'BUSY', message: '系統忙碌中，請稍後再試。' });
  }
  try {
    var body = parseBody_(e);
    var result = processOrder_(body);
    return json_(result);
  } catch (err) {
    console.error(err && err.stack || err);
    return json_({ ok: false, code: 'SERVER', message: '訂單送出失敗，請稍後再試或直接聯絡門市。' });
  } finally {
    lock.releaseLock();
  }
}

function parseBody_(e) {
  var raw = e && e.postData && e.postData.contents;
  if (!raw || raw.length > 20000) throw new Error('bad body');
  return JSON.parse(raw);
}

function processOrder_(body) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var orders = ss.getSheetByName(SHEET_ORDERS);
  var items = ss.getSheetByName(SHEET_ITEMS);
  var menuSheet = ss.getSheetByName(SHEET_MENU);
  if (!orders || !items || !menuSheet) return { ok: false, code: 'NOT_SETUP', message: '後台尚未完成設定。' };

  var menu = readMenu_(menuSheet);
  var checked = validateOrder_(body, menu);
  if (!checked.ok) return checked;
  var o = checked.order;

  // 重送保護：同一個 requestId 只會建立一張訂單。
  var dup = findOrderByRequestId_(orders, o.requestId);
  if (dup) return { ok: true, orderNo: dup.orderNo, total: dup.total, duplicate: true };

  var orderNo = nextOrderNo_(orders);
  var now = new Date();
  orders.appendRow([
    orderNo, Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm:ss'), '新訂單', MODES[o.mode],
    o.mode === 'delivery' ? '' : STORES[o.store], o.table, o.customer, o.phone, o.email, o.address, o.note,
    o.subtotal, o.shipping, o.total, o.requestId
  ]);
  var rows = o.lines.map(function (l) { return [orderNo, l.id, l.name, l.price, l.qty, l.price * l.qty]; });
  items.getRange(items.getLastRow() + 1, 1, rows.length, ITEM_HEADERS.length).setValues(rows);

  notify_(orderNo, o);
  return { ok: true, orderNo: orderNo, total: o.total };
}

function readMenu_(sheet) {
  var map = {};
  var n = sheet.getLastRow() - 1;
  if (n < 1) return map;
  sheet.getRange(2, 1, n, MENU_HEADERS.length).getValues().forEach(function (r) {
    if (!r[0]) return;
    map[String(r[0])] = { id: String(r[0]), name: String(r[1]), price: Number(r[3]), frozen: r[5] === true || String(r[5]).toUpperCase() === 'TRUE', available: r[6] === true || String(r[6]).toUpperCase() === 'TRUE' };
  });
  return map;
}

/** 純函式：驗證輸入並以菜單重算金額（可在本機以 Node 測試）。 */
function validateOrder_(body, menu) {
  var fail = function (code, message) { return { ok: false, code: code, message: message }; };
  if (!body || typeof body !== 'object') return fail('BAD', '資料格式錯誤。');
  if (body.website) return fail('SPAM', '無法送出。'); // 蜜罐欄位：真人看不到也不會填
  var mode = String(body.mode || '');
  if (!MODES[mode]) return fail('BAD', '請選擇取餐方式。');

  var customer = clean_(body.customer, 40), phone = clean_(body.phone, 20);
  if (!customer) return fail('BAD', '請輸入聯絡姓名。');
  if (!/^[0-9+()\s\-]{8,20}$/.test(phone) || phone.replace(/\D/g, '').length < 8) return fail('BAD', '請輸入有效電話。');
  var email = clean_(body.email, 120);
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return fail('BAD', 'Email 格式不正確。');
  var address = clean_(body.address, 160), table = clean_(body.table, 8);
  if (mode === 'delivery' && address.length < 6) return fail('BAD', '請填寫完整地址。');
  if (mode === 'dinein' && !table) return fail('BAD', '請填寫桌號。');
  var store = Number(body.store);
  if (mode !== 'delivery' && !(Number.isInteger(store) && store >= 0 && store < STORES.length)) return fail('BAD', '請選擇門市。');
  var requestId = clean_(body.requestId, 64);
  if (!/^[A-Za-z0-9\-]{8,64}$/.test(requestId)) return fail('BAD', '資料格式錯誤。');

  if (!Array.isArray(body.items) || body.items.length < 1 || body.items.length > 30) return fail('BAD', '購物車是空的。');
  var seen = {}, lines = [], subtotal = 0, allFrozen = true;
  for (var i = 0; i < body.items.length; i++) {
    var it = body.items[i] || {};
    var p = menu[String(it.id)];
    var qty = Number(it.qty);
    if (!p || seen[p.id]) return fail('BAD', '商品資料有誤，請重新整理後再試。');
    if (!p.available) return fail('SOLD_OUT', '「' + p.name + '」目前暫停供應，請回到購物車調整。');
    if (!Number.isInteger(qty) || qty < 1 || qty > 99) return fail('BAD', '商品數量有誤。');
    seen[p.id] = true;
    if (!p.frozen) allFrozen = false;
    subtotal += p.price * qty;
    lines.push({ id: p.id, name: p.name, price: p.price, qty: qty });
  }
  if (mode === 'delivery' && !allFrozen) return fail('BAD', '宅配僅限整筆皆為冷凍禮盒的訂單。');
  var shipping = mode === 'delivery' && subtotal < FREE_SHIPPING_MIN ? SHIPPING_FEE : 0;
  var total = subtotal + shipping;
  if (body.clientTotal !== undefined && Number(body.clientTotal) !== total) {
    return fail('PRICE_CHANGED', '商品價格已更新（目前合計 NT$' + total + '），請重新整理頁面後再下單。');
  }
  return { ok: true, order: { mode: mode, store: store, table: table, customer: customer, phone: phone, email: email, address: mode === 'delivery' ? address : '', note: clean_(body.note, 300), lines: lines, subtotal: subtotal, shipping: shipping, total: total, requestId: requestId } };
}

function clean_(v, max) {
  return String(v == null ? '' : v).replace(/[\u0000-\u001f\u007f]/g, ' ').trim().slice(0, max);
}

function findOrderByRequestId_(sheet, requestId) {
  var last = sheet.getLastRow();
  if (last < 2) return null;
  var col = ORDER_HEADERS.indexOf('requestId') + 1;
  var ids = sheet.getRange(Math.max(2, last - 199), col, Math.min(200, last - 1), 1).getValues();
  var base = Math.max(2, last - 199);
  for (var i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === requestId) {
      var row = sheet.getRange(base + i, 1, 1, ORDER_HEADERS.length).getValues()[0];
      return { orderNo: row[0], total: row[13] };
    }
  }
  return null;
}

/** 訂單編號：GC-YYYYMMDD-001，每日從 001 起算，在鎖內產生所以不會重複。 */
function nextOrderNo_(sheet) {
  var day = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd');
  var prefix = 'GC-' + day + '-';
  var last = sheet.getLastRow(), seq = 0;
  if (last >= 2) {
    var nos = sheet.getRange(Math.max(2, last - 499), 1, Math.min(500, last - 1), 1).getValues();
    nos.forEach(function (r) {
      var s = String(r[0]);
      if (s.indexOf(prefix) === 0) seq = Math.max(seq, Number(s.slice(prefix.length)) || 0);
    });
  }
  return prefix + ('00' + (seq + 1)).slice(-3);
}

function notify_(orderNo, o) {
  var to = PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAIL');
  if (!to) return;
  try {
    var lines = o.lines.map(function (l) { return l.name + ' × ' + l.qty + '　NT$' + l.price * l.qty; }).join('\n');
    MailApp.sendEmail(to, '【新訂單】' + orderNo + '｜' + MODES[o.mode] + '｜NT$' + o.total,
      lines + '\n\n合計：NT$' + o.total + '\n姓名：' + o.customer + '\n電話：' + o.phone +
      (o.mode === 'delivery' ? '\n地址：' + o.address : '\n門市：' + STORES[o.store] + (o.table ? '　桌號：' + o.table : '')) +
      (o.note ? '\n備註：' + o.note : ''));
  } catch (err) {
    console.error('notify failed', err); // 通知失敗不影響訂單
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
