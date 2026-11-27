/**
 * PVEP 20 NĂM — HỆ THỐNG QUẢN LÝ KHO QUÀ TẶNG & TỜ TRÌNH (SERVERLESS APPS SCRIPT BACKEND)
 * Backend Google Apps Script (file "Mã.gs" / "Code.gs")
 *
 * Sheet 1 - SAN_PHAM (13 cột A -> M):
 *   MaSP | TenSanPham | PhanLoai | TonKho | ThongDiep | TieuDeMoTa | MoTaChiTiet | ThongSo_1 | ThongSo_2 | ThongSo_3 | ThongSo_4 | HinhAnh | TrangThai
 * Sheet 2 - TAI_KHOAN:
 *   Username | Password | HoTen | ChucVu | TrangThai
 * Sheet 3 - TO_TRINH (Tự động khởi tạo nếu chưa có):
 *   ID | MaToTrinh | NgayLap | KinhGui | DonVi | NguoiLap | MucDich | DoiTuong | TrangThai | NgayDuyet | NgayTao | DanhSachJSON
 * Sheet 4 - NHAT_KY (Tự động khởi tạo nếu chưa có):
 *   ThoiGian | NhanVien | MaSP | HanhDong | TonCu | TonMoi
 */

const SS_ID = '1EetTvh90juJ2LoaTHK1k8pIe0tyZcFsHM6yQaqQ01UQ';
const SHEET_PRODUCTS = 'SAN_PHAM';
const SHEET_USERS = 'TAI_KHOAN';
const SHEET_PROPOSALS = 'TO_TRINH';
const SHEET_LOGS = 'NHAT_KY';
const SESSION_SECONDS = 6 * 60 * 60;

const HEADERS_PRODUCTS = [
  'MaSP', 'TenSanPham', 'PhanLoai', 'TonKho', 'ThongDiep',
  'TieuDeMoTa', 'MoTaChiTiet', 'ThongSo_1', 'ThongSo_2',
  'ThongSo_3', 'ThongSo_4', 'HinhAnh', 'TrangThai'
];

const HEADERS_PROPOSALS = [
  'ID', 'MaToTrinh', 'NgayLap', 'KinhGui', 'DonVi', 'NguoiLap',
  'MucDich', 'DoiTuong', 'TrangThai', 'NgayDuyet', 'NgayTao', 'DanhSachJSON'
];

const HEADERS_LOGS = [
  'ThoiGian', 'NhanVien', 'MaSP', 'HanhDong', 'TonCu', 'TonMoi'
];

/* ------------------------------ Web App & API ------------------------------ */
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'listProducts') {
    var list = readProductsFromSheet_();
    return jsonOut_({ ok: true, products: list });
  }
  if (e && e.parameter && e.parameter.action === 'bootstrap') {
    return jsonOut_({
      ok: true,
      products: readProductsFromSheet_(),
      serverTime: new Date().toISOString()
    });
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
    var action = String(body.action || '').trim();

    if (action === 'login') {
      return jsonOut_(login(body.username, body.password));
    }

    if (action === 'logout') {
      logout(body.token);
      return jsonOut_({ ok: true });
    }

    if (action === 'getAllData') {
      requireSession_(body.token);
      return jsonOut_({
        ok: true,
        products: readProductsFromSheet_(),
        proposals: readProposalsFromSheet_(),
        logs: readLogsFromSheet_(),
        serverTime: new Date().toISOString()
      });
    }

    if (action === 'syncProducts' && Array.isArray(body.products)) {
      requireSession_(body.token);
      return jsonOut_(syncProductsWithConcurrency_(body.products, body.expectedStocks, !!body.force));
    }

    if (action === 'appendLog' && body.log) {
      var sessLog = requireSession_(body.token);
      appendLogToSheet_(body.log, sessLog);
      return jsonOut_({ ok: true });
    }

    if (action === 'saveProposal' && body.proposal) {
      requireSession_(body.token);
      var proposalsAfterSave = saveProposalToSheet_(body.proposal);
      return jsonOut_({ ok: true, proposals: proposalsAfterSave });
    }

    if (action === 'approveProposal' && body.proposalId) {
      var sessApprove = requireSession_(body.token);
      return jsonOut_(approveProposalAtomic_(body.proposalId, sessApprove));
    }

    return jsonOut_({ ok: false, message: 'Invalid action' });
  } catch (err) {
    var msg = String(err && err.message ? err.message : err);
    return jsonOut_({ ok: false, code: msg === 'SESSION_EXPIRED' ? 'SESSION_EXPIRED' : 'SERVER_ERROR', message: msg });
  }
}

/* ------------------------------ Helpers ------------------------------ */
function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheet_(name) {
  var sh = SpreadsheetApp.openById(SS_ID).getSheetByName(name);
  if (!sh) throw new Error('Không tìm thấy tab ' + name);
  return sh;
}

function ensureSheet_(name, headers) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }
  if (sh.getLastRow() === 0 && headers && headers.length) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sh;
}

function rows_(name, optionalHeaders) {
  var sh = optionalHeaders ? ensureSheet_(name, optionalHeaders) : sheet_(name);
  var v = sh.getDataRange().getValues();
  if (v.length <= 1) return [];
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

function computeInitials_(name, username) {
  var parts = String(name || '').trim().split(/\s+/).filter(Boolean);
  if (parts.length >= 2) {
    return (parts[0].charAt(0) + parts[parts.length - 1].charAt(0)).toUpperCase();
  }
  return String(username || 'NV').slice(0, 2).toUpperCase();
}

function normalizeCode_(raw) {
  var code = String(raw || '').trim().replace(/^sp/i, '');
  if (code.length === 1) code = '0' + code;
  return code;
}

/* ------------------------------ Products CRUD ------------------------------ */
function readProductsFromSheet_() {
  return rows_(SHEET_PRODUCTS).filter(function (p) {
    return String(p.MaSP || p.ID || '').trim() !== '';
  }).map(function (p) {
    var code = normalizeCode_(p.MaSP || p.ID);
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

function writeProductsToSheetLocked_(products) {
  var sh = sheet_(SHEET_PRODUCTS);
  var rows = products.map(function (p) {
    var sp = p.specs || [];
    var stk = Math.max(0, Number(p.stock) || 0);
    return [
      normalizeCode_(p.code),
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
  sh.getRange(1, 1, 1, HEADERS_PRODUCTS.length).setValues([HEADERS_PRODUCTS]);
  if (rows.length > 0) {
    sh.getRange(2, 1, rows.length, HEADERS_PRODUCTS.length).setValues(rows);
  }
}

function writeProductsToSheet_(products) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    writeProductsToSheetLocked_(products);
  } finally {
    lock.releaseLock();
  }
}

function syncProductsWithConcurrency_(products, expectedStocks, force) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var liveProducts = readProductsFromSheet_();
    if (!force && expectedStocks && typeof expectedStocks === 'object') {
      var conflicts = [];
      Object.keys(expectedStocks).forEach(function (codeKey) {
        var code = normalizeCode_(codeKey);
        var expVal = Number(expectedStocks[codeKey]) || 0;
        var liveItem = liveProducts.filter(function (x) { return x.code === code; })[0];
        var stagedItem = products.filter(function (x) { return normalizeCode_(x.code) === code; })[0];
        var liveStock = liveItem ? (Number(liveItem.stock) || 0) : 0;
        var stagedStock = stagedItem ? (Number(stagedItem.stock) || 0) : expVal;
        if (liveStock !== expVal) {
          var delta = stagedStock - expVal;
          conflicts.push({
            code: code,
            name: stagedItem ? stagedItem.name : (liveItem ? liveItem.name : ('Sản phẩm #' + code)),
            initialStock: expVal,
            liveStock: liveStock,
            stagedStock: stagedStock,
            delta: delta,
            mergedStock: Math.max(0, liveStock + delta)
          });
        }
      });
      if (conflicts.length > 0) {
        return {
          ok: false,
          code: 'CONFLICT',
          conflicts: conflicts,
          liveProducts: liveProducts
        };
      }
    }

    writeProductsToSheetLocked_(products);
    return {
      ok: true,
      products: readProductsFromSheet_(),
      serverTime: new Date().toISOString()
    };
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------ Logs CRUD ------------------------------ */
function readLogsFromSheet_() {
  return rows_(SHEET_LOGS, HEADERS_LOGS).filter(function (r) {
    return String(r.ThoiGian || r.HanhDong || '').trim() !== '';
  }).map(function (r) {
    return {
      time: String(r.ThoiGian || ''),
      actor: String(r.NhanVien || ''),
      code: String(r.MaSP || ''),
      action: String(r.HanhDong || ''),
      oldVal: Number(r.TonCu) || 0,
      newVal: Number(r.TonMoi) || 0
    };
  });
}

function appendLogToSheetLocked_(logObj, session) {
  var sh = ensureSheet_(SHEET_LOGS, HEADERS_LOGS);
  var timeStr = String(logObj.time || Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'HH:mm:ss dd/MM/yyyy'));
  var actorStr = String(logObj.actor || (session && session.name) || 'Quản trị viên');
  sh.insertRowAfter(1);
  sh.getRange(2, 1, 1, HEADERS_LOGS.length).setValues([[
    timeStr,
    actorStr,
    String(logObj.code || ''),
    String(logObj.action || ''),
    Number(logObj.oldVal) || 0,
    Number(logObj.newVal) || 0
  ]]);
  var lastRow = sh.getLastRow();
  if (lastRow > 501) {
    sh.deleteRows(502, lastRow - 501);
  }
}

function appendLogToSheet_(logObj, session) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    appendLogToSheetLocked_(logObj, session);
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------ Proposals CRUD ------------------------------ */
function readProposalsFromSheet_() {
  return rows_(SHEET_PROPOSALS, HEADERS_PROPOSALS).filter(function (r) {
    return String(r.ID || r.MaToTrinh || '').trim() !== '';
  }).map(function (r) {
    var items = [];
    try { items = JSON.parse(String(r.DanhSachJSON || '[]')) || []; } catch (e) {}
    return {
      id: String(r.ID || ''),
      code: String(r.MaToTrinh || ''),
      date: String(r.NgayLap || ''),
      recipientBoss: String(r.KinhGui || ''),
      dept: String(r.DonVi || ''),
      requester: String(r.NguoiLap || ''),
      purpose: String(r.MucDich || ''),
      targetGuests: String(r.DoiTuong || ''),
      status: String(r.TrangThai || 'pending'),
      approvedAt: String(r.NgayDuyet || ''),
      createdAt: String(r.NgayTao || ''),
      items: items
    };
  });
}

function proposalToRow_(p) {
  return [
    String(p.id || ''),
    String(p.code || ''),
    String(p.date || ''),
    String(p.recipientBoss || ''),
    String(p.dept || ''),
    String(p.requester || ''),
    String(p.purpose || ''),
    String(p.targetGuests || ''),
    String(p.status || 'pending'),
    String(p.approvedAt || ''),
    String(p.createdAt || ''),
    JSON.stringify(p.items || [])
  ];
}

function saveProposalToSheet_(proposal) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = ensureSheet_(SHEET_PROPOSALS, HEADERS_PROPOSALS);
    var existing = readProposalsFromSheet_();
    var targetId = String(proposal.id || '');
    var foundIndex = -1;
    for (var i = 0; i < existing.length; i++) {
      if (existing[i].id === targetId) { foundIndex = i; break; }
    }
    var rowData = proposalToRow_(proposal);
    if (foundIndex >= 0) {
      sh.getRange(foundIndex + 2, 1, 1, HEADERS_PROPOSALS.length).setValues([rowData]);
    } else {
      sh.insertRowAfter(1);
      sh.getRange(2, 1, 1, HEADERS_PROPOSALS.length).setValues([rowData]);
      var lastRow = sh.getLastRow();
      if (lastRow > 201) {
        sh.deleteRows(202, lastRow - 201);
      }
    }
    return readProposalsFromSheet_();
  } finally {
    lock.releaseLock();
  }
}

function approveProposalAtomic_(proposalId, session) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = ensureSheet_(SHEET_PROPOSALS, HEADERS_PROPOSALS);
    var proposals = readProposalsFromSheet_();
    var propIndex = -1;
    for (var i = 0; i < proposals.length; i++) {
      if (proposals[i].id === String(proposalId)) { propIndex = i; break; }
    }
    if (propIndex < 0) {
      return { ok: false, message: 'Không tìm thấy tờ trình #' + proposalId };
    }

    var prop = proposals[propIndex];
    if (prop.status === 'approved') {
      return {
        ok: true,
        alreadyApproved: true,
        products: readProductsFromSheet_(),
        proposals: proposals,
        logs: readLogsFromSheet_()
      };
    }

    var products = readProductsFromSheet_();
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'HH:mm:ss dd/MM/yyyy');
    (prop.items || []).forEach(function (it) {
      var code = normalizeCode_(it.code);
      var pIdx = -1;
      for (var j = 0; j < products.length; j++) {
        if (products[j].code === code) { pIdx = j; break; }
      }
      if (pIdx >= 0) {
        var oldVal = Number(products[pIdx].stock) || 0;
        var newVal = Math.max(0, oldVal - (Number(it.qty) || 0));
        products[pIdx].stock = newVal;
        appendLogToSheetLocked_({
          time: nowStr,
          actor: (session && session.name) || 'Quản trị viên',
          code: code,
          action: 'Xuất kho theo Tờ trình đã duyệt (' + prop.code + ')',
          oldVal: oldVal,
          newVal: newVal
        }, session);
      }
    });

    prop.status = 'approved';
    prop.approvedAt = nowStr;
    writeProductsToSheetLocked_(products);
    sh.getRange(propIndex + 2, 1, 1, HEADERS_PROPOSALS.length).setValues([proposalToRow_(prop)]);

    return {
      ok: true,
      alreadyApproved: false,
      products: readProductsFromSheet_(),
      proposals: readProposalsFromSheet_(),
      logs: readLogsFromSheet_()
    };
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------ Auth ------------------------------ */
function login(username, password) {
  username = String(username || '').trim().toLowerCase();
  password = String(password || '');
  var user = rows_(SHEET_USERS).filter(function (u) {
    var un = String(u.Username || '').trim().toLowerCase();
    return un !== '' && un === username && String(u.TrangThai || '').trim() === 'HoatDong';
  })[0];
  if (!user || String(user.Password) !== password) {
    return { ok: false, message: 'Tên đăng nhập hoặc mật khẩu không đúng.' };
  }
  var token = Utilities.getUuid() + Utilities.getUuid();
  var fullName = String(user.HoTen || username);
  var session = {
    username: String(user.Username || username).trim().toLowerCase(),
    name: fullName,
    role: String(user.ChucVu || 'Nhân viên'),
    initials: computeInitials_(fullName, username)
  };
  CacheService.getScriptCache().put('sess_' + token, JSON.stringify(session), SESSION_SECONDS);
  return { ok: true, token: token, user: session };
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('sess_' + token);
  return true;
}

/* ------------------------------ Legacy RPC Functions (compat) ------------------------------ */
function getData(token) {
  var s = requireSession_(token);
  return { user: s, products: readProductsFromSheet_() };
}

function updateStock(token, code, newStock) {
  requireSession_(token);
  var list = readProductsFromSheet_();
  var targetCode = normalizeCode_(code);
  var item = list.filter(function (x) { return x.code === targetCode; })[0];
  if (!item) throw new Error('Không tìm thấy sản phẩm ' + targetCode);
  item.stock = Math.max(0, Math.floor(Number(newStock) || 0));
  writeProductsToSheet_(list);
  return { ok: true, products: list };
}

function saveProductTemplate(token, prod) {
  requireSession_(token);
  var list = readProductsFromSheet_();
  var code = normalizeCode_(prod.code);
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
