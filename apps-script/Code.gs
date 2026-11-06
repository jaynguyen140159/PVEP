/**
 * PVEP 20 NĂM — HỆ THỐNG QUẢN LÝ SẢN PHẨM & TỒN KHO (ĐỒNG BỘ TEMPLATE WEBSITE)
 * Backend Google Apps Script (file "Mã.gs" / "Code.gs")
 *
 * Cấu trúc Sheet SAN_PHAM (13 cột A -> M):
 * MaSP | TenSanPham | PhanLoai | TonKho | ThongDiep | TieuDeMoTa | MoTaChiTiet | ThongSo_1 | ThongSo_2 | ThongSo_3 | ThongSo_4 | HinhAnh | TrangThai
 */

const SS_ID = '1EetTvh90juJ2LoaTHK1k8pIe0tyZcFsHM6yQaqQ01UQ';
const SHEET_PRODUCTS = 'SAN_PHAM';
const SHEET_USERS = 'TAI_KHOAN';
const SESSION_SECONDS = 6 * 60 * 60;

/* ------------------------------ Web App & API ------------------------------ */
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'listProducts') {
    var list = readProductsFromSheet_();
    return ContentService.createTextOutput(JSON.stringify({ ok: true, products: list }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var back = (e && e.parameter && e.parameter.back) || '';
  if (!/^https?:\/\/[^\s"'<>]+$/i.test(back)) back = '';
  var t = HtmlService.createTemplateFromFile('index');
  t.backUrl = back;
  return t.evaluate()
    .setTitle('PVEP 20 Năm · Quản lý Sản phẩm & Tồn kho')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action === 'syncProducts' && Array.isArray(body.products)) {
      writeProductsToSheet_(body.products);
      return ContentService.createTextOutput(JSON.stringify({ ok: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: false, message: 'Invalid action' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/* ------------------------------ Helpers ------------------------------ */
function sheet_(name) {
  var sh = SpreadsheetApp.openById(SS_ID).getSheetByName(name);
  if (!sh) throw new Error('Không tìm thấy tab ' + name);
  return sh;
}

function rows_(name) {
  var sh = sheet_(name);
  var v = sh.getDataRange().getValues();
  var head = v.shift() || [];
  return v.map(function (r, i) {
    var o = { _row: i + 2 };
    head.forEach(function (h, c) { o[h] = r[c]; });
    return o;
  });
}

function requireSession_(token) {
  var raw = token && CacheService.getScriptCache().get('sess_' + token);
  if (!raw) throw new Error('SESSION_EXPIRED');
  return JSON.parse(raw);
}

function readProductsFromSheet_() {
  return rows_(SHEET_PRODUCTS).filter(function (p) {
    return String(p.MaSP || p.ID || '').trim() !== '';
  }).map(function (p) {
    var code = String(p.MaSP || p.ID || '').trim().replace(/^SP/i, '');
    if (code.length === 1) code = '0' + code;
    return {
      code: code,
      name: String(p.TenSanPham || ''),
      category: String(p.PhanLoai || p.DanhMuc || ''),
      stock: Number(p.TonKho) || 0,
      tagline: String(p.ThongDiep || ''),
      specHeading: String(p.TieuDeMoTa || ''),
      desc: String(p.MoTaChiTiet || ''),
      specs: [
        String(p.ThongSo_1 || ''),
        String(p.ThongSo_2 || ''),
        String(p.ThongSo_3 || ''),
        String(p.ThongSo_4 || '')
      ].filter(Boolean),
      image: String(p.HinhAnh || ''),
      status: String(p.TrangThai || 'Đang hiển thị')
    };
  });
}

function writeProductsToSheet_(products) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet_(SHEET_PRODUCTS);
    var headers = [[
      'MaSP', 'TenSanPham', 'PhanLoai', 'TonKho', 'ThongDiep',
      'TieuDeMoTa', 'MoTaChiTiet', 'ThongSo_1', 'ThongSo_2',
      'ThongSo_3', 'ThongSo_4', 'HinhAnh', 'TrangThai'
    ]];
    var rows = products.map(function (p) {
      var sp = p.specs || [];
      var stk = Math.max(0, Number(p.stock) || 0);
      return [
        String(p.code || ''),
        String(p.name || ''),
        String(p.category || ''),
        stk,
        String(p.tagline || ''),
        String(p.specHeading || ''),
        String(p.desc || ''),
        String(sp[0] || ''),
        String(sp[1] || ''),
        String(sp[2] || ''),
        String(sp[3] || ''),
        String(p.image || ''),
        stk > 0 ? 'Đang hiển thị' : 'Hết hàng'
      ];
    });
    sh.clearContents();
    sh.getRange(1, 1, 1, 13).setValues(headers);
    if (rows.length > 0) {
      sh.getRange(2, 1, rows.length, 13).setValues(rows);
    }
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------ Auth ------------------------------ */
function login(username, password) {
  username = String(username || '').trim().toLowerCase();
  password = String(password || '');
  var user = rows_(SHEET_USERS).filter(function (u) {
    return String(u.Username).trim().toLowerCase() === username &&
           String(u.TrangThai).trim() === 'HoatDong';
  })[0];
  if (!user || String(user.Password) !== password) {
    return { ok: false, message: 'Tên đăng nhập hoặc mật khẩu không đúng.' };
  }
  var token = Utilities.getUuid() + Utilities.getUuid();
  var session = { username: user.Username, name: user.HoTen, role: user.ChucVu };
  CacheService.getScriptCache().put('sess_' + token, JSON.stringify(session), SESSION_SECONDS);
  return { ok: true, token: token, user: session };
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('sess_' + token);
  return true;
}

/* ------------------------------ Data & Actions ------------------------------ */
function getData(token) {
  var s = requireSession_(token);
  return { user: s, products: readProductsFromSheet_() };
}

function updateStock(token, code, newStock) {
  requireSession_(token);
  var list = readProductsFromSheet_();
  var item = list.filter(function (x) { return x.code === String(code); })[0];
  if (!item) throw new Error('Không tìm thấy sản phẩm ' + code);
  item.stock = Math.max(0, Math.floor(Number(newStock) || 0));
  writeProductsToSheet_(list);
  return { ok: true, products: list };
}

function saveProductTemplate(token, prod) {
  requireSession_(token);
  var list = readProductsFromSheet_();
  var code = String(prod.code || '').trim().replace(/^SP/i, '');
  if (code.length === 1) code = '0' + code;
  prod.code = code;
  var idx = -1;
  for (var i = 0; i < list.length; i++) {
    if (list[i].code === code) { idx = i; break; }
  }
  if (idx >= 0) list[idx] = prod;
  else list.push(prod);
  writeProductsToSheet_(list);
  return { ok: true, products: list };
}
