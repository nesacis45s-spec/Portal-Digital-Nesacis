/**
 * BACKEND "Jurnal Harian Guru Digital" — Google Apps Script + Google Spreadsheet
 *
 * Cara pakai: lihat PANDUAN_SETUP.md
 *  - Skrip ini harus dibuat dari menu Extensions > Apps Script di dalam Spreadsheet Anda.
 *  - Deploy sebagai Web App: Execute as "Me", Who has access "Anyone".
 *
 * Sheet yang dibuat otomatis:
 *  - Users    : email | nama | salt | passHash | createdAt
 *  - Sessions : tokenHash | email | expires
 *  - Data     : email | key | updatedAt | chunk1 | chunk2 | ...   (1 baris per pengguna per jenis data;
 *               data JSON dipecah per 40.000 karakter karena batas 50.000 karakter per sel)
 *  - Resets   : email | codeHash | expires   (kode reset password)
 *  - Periode  : id | label | spreadsheetId   (daftar Tahun Ajaran/Semester; ISI MANUAL di spreadsheet PUSAT ini)
 *
 * Pembagian data:
 *  - Spreadsheet PUSAT (tempat skrip ini): akun (Users, Sessions, Resets) + daftar Periode.
 *  - Spreadsheet PER PERIODE (file terpisah, dibuka lewat spreadsheetId): sheet Data (jurnal, absensi, nilai, dst.)
 *    dan Sheet1 (sumber data siswa). Sheet "Data" dibuat otomatis di file periode saat pertama dipakai.
 */

var SESSION_DAYS = 30;
var CHUNK_SIZE = 40000;
var MAX_FAILS = 5;            // maksimal salah password sebelum dikunci sementara
var LOCK_SECONDS = 300;       // lama kunci (detik)
var SISWA_SHEET = 'Sheet1';   // nama sheet sumber data siswa (untuk fitur "Ambil dari Spreadsheet")
var ALLOWED_KEYS = ['profile', 'students', 'journals', 'attendance', 'scores', 'schedule', 'settings', 'notes'];
var PERIODE_SHEET = 'Periode'; // sheet di spreadsheet pusat: id | label | spreadsheetId

/* ------------------------------ ENTRY POINT ------------------------------ */

function doGet() {
  return json_({ ok: true, msg: 'Jurnal Guru API aktif.' });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    var req = JSON.parse(e.postData.contents);
    // Aksi yang hanya MEMBACA tidak perlu antre di kunci global -> beberapa guru bisa memuat data bersamaan.
    var readOnly = (req.action === 'load' || req.action === 'periodes' || req.action === 'sheetStudents');
    if (!readOnly) { lock.waitLock(20000); locked = true; }
    switch (req.action) {
      case 'register':       return json_(register_(req));
      case 'login':          return json_(login_(req));
      case 'logout':         return json_(logout_(req));
      case 'load':           return json_(load_(req));
      case 'save':           return json_(save_(req));
      case 'sheetStudents':  return json_(sheetStudents_(req));
      case 'periodes':       return json_(periodes_(req));
      case 'forgot':         return json_(forgot_(req));
      case 'reset':          return json_(reset_(req));
      case 'changePassword': return json_(changePassword_(req));
      default:               return json_({ ok: false, error: 'Aksi tidak dikenal.' });
    }
  } catch (err) {
    return json_({ ok: false, error: 'Server sibuk atau terjadi kesalahan: ' + err.message });
  } finally {
    if (locked) { try { lock.releaseLock(); } catch (x) {} }
  }
}

/* -------------------------------- HELPERS -------------------------------- */

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_(name, headers, ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    try {
      sh = ss.insertSheet(name);
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      sh.setFrozenRows(1);
    } catch (err) {
      sh = ss.getSheetByName(name);      // sudah dibuat oleh permintaan lain
      if (!sh) throw err;
    }
  }
  return sh;
}
function users_()    { return sheet_('Users',    ['email', 'nama', 'salt', 'passHash', 'createdAt']); }
function sessions_() { return sheet_('Sessions', ['tokenHash', 'email', 'expires']); }
function data_(ss)   { return sheet_('Data',     ['email', 'key', 'updatedAt', 'chunk1'], ss); }   // ss = spreadsheet periode
function resets_()   { return sheet_('Resets',   ['email', 'codeHash', 'expires']); }

/* ------------------------------- PERIODE --------------------------------- */

var PERIODE_CACHE_KEY = 'periodeList_v1';
var PERIODE_CACHE_SEC = 600;     // perubahan sheet Periode terbaca maksimal 10 menit (atau langsung, lewat onEdit di bawah)

function cacheGet_(k) {
  try { var v = CacheService.getScriptCache().get(k); return v ? JSON.parse(v) : null; } catch (e) { return null; }
}
function cachePut_(k, obj, sec) {
  try { CacheService.getScriptCache().put(k, JSON.stringify(obj), sec); } catch (e) {}
}
function cacheDel_(k) {
  try { CacheService.getScriptCache().remove(k); } catch (e) {}
}

/** Otomatis menghapus cache daftar periode saat sheet "Periode" diedit manual. */
function onEdit(e) {
  try {
    if (e && e.range && e.range.getSheet().getName() === PERIODE_SHEET) cacheDel_(PERIODE_CACHE_KEY);
  } catch (err) {}
}

/** Baca sheet "Periode" di spreadsheet pusat -> [{id, label, ssId}] (di-cache 10 menit) */
function periodeList_() {
  var cached = cacheGet_(PERIODE_CACHE_KEY);
  if (cached) return cached;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PERIODE_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  var vals = sh.getDataRange().getDisplayValues();
  var head = vals[0].map(function (h) { return String(h).trim().toLowerCase().replace(/[\s_]/g, ''); });
  var iId = head.indexOf('id'), iLabel = head.indexOf('label'), iSS = head.indexOf('spreadsheetid');
  if (iId < 0) iId = 0;
  if (iLabel < 0) iLabel = 1;
  if (iSS < 0) iSS = 2;
  var out = [];
  for (var r = 1; r < vals.length; r++) {
    var id = String(vals[r][iId]).trim();
    var label = String(vals[r][iLabel]).trim();
    var raw = String(vals[r][iSS]).trim();
    var m = raw.match(/\/d\/([a-zA-Z0-9_-]+)/);       // boleh tempel URL lengkap spreadsheet
    var ssId = m ? m[1] : raw;
    if (!id || !ssId) continue;
    out.push({ id: id, label: label || id, ssId: ssId });
  }
  if (out.length) cachePut_(PERIODE_CACHE_KEY, out, PERIODE_CACHE_SEC);
  return out;
}

/** Daftar periode untuk ditampilkan di aplikasi (tanpa spreadsheetId). */
function publicPeriodes_() {
  return periodeList_().map(function (p) { return { id: p.id, label: p.label }; });
}

/** Buka spreadsheet periode. Hasil: { ss } atau { error, reauth }. */
function periodeSS_(id) {
  id = String(id || '').trim();
  if (!id) return { error: 'Periode belum dipilih. Silakan masuk kembali.', reauth: true };
  var list = periodeList_();
  for (var i = 0; i < list.length; i++) {
    if (list[i].id !== id) continue;
    try {
      return { ss: SpreadsheetApp.openById(list[i].ssId) };
    } catch (err) {
      return { error: 'Spreadsheet untuk periode "' + list[i].label + '" tidak bisa dibuka. Periksa spreadsheetId dan izin akses pemilik skrip.' };
    }
  }
  return { error: 'Periode tidak dikenal. Silakan masuk kembali dan pilih periode.', reauth: true };
}

function periodes_(req) {
  var email = auth_(req.token);
  if (!email) return { ok: false, error: 'Sesi berakhir. Silakan masuk kembali.', auth: false };
  return { ok: true, periodes: publicPeriodes_() };
}

/** Jalankan SEKALI dari editor Apps Script (tombol Run): memberi izin akses ke file periode + memeriksa semua spreadsheetId. */
function tesPeriode() {
  var list = periodeList_();
  if (!list.length) { Logger.log('Sheet "Periode" kosong / tidak ditemukan.'); return; }
  list.forEach(function (p) {
    try {
      var ss = SpreadsheetApp.openById(p.ssId);
      Logger.log('OK   ' + p.id + ' (' + p.label + ') -> ' + ss.getName());
    } catch (e) {
      Logger.log('GAGAL ' + p.id + ' (' + p.label + ') -> ' + e.message);
    }
  });
}

function hex_(bytes) {
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}
function sha256_(text) {
  return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8));
}
function hashPassword_(password, salt) {
  var h = salt + ':' + password;
  for (var i = 0; i < 50; i++) h = sha256_(h + salt);   // sedikit "stretching"
  return h;
}
function normEmail_(s) { return String(s || '').trim().toLowerCase(); }
function validEmail_(s) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s); }

/** Cari baris (nomor baris sheet) berdasarkan nilai kolom 1, -1 jika tidak ada. */
function findRow_(sh, value, col) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var vals = sh.getRange(2, col || 1, last - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).toLowerCase() === String(value).toLowerCase()) return i + 2;
  }
  return -1;
}

/* ------------------------------- SESSIONS -------------------------------- */

function createSession_(email) {
  var token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
  var sh = sessions_();
  cleanupSessions_(sh);
  var expires = Date.now() + SESSION_DAYS * 86400000;
  sh.appendRow([sha256_(token), email, expires]);
  return { token: token, expires: expires };
}

function cleanupSessions_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return;
  var vals = sh.getRange(2, 1, last - 1, 3).getValues();
  var now = Date.now();
  for (var i = vals.length - 1; i >= 0; i--) {
    if (Number(vals[i][2]) < now) sh.deleteRow(i + 2);
  }
}

/** Kembalikan email jika token valid, null jika tidak. */
function auth_(token) {
  if (!token) return null;
  var h = sha256_(String(token));
  var ck = 's_' + h;
  var c = cacheGet_(ck);
  if (c && Number(c.exp) > Date.now()) return String(c.email);        // cepat: tanpa membaca sheet Sessions

  var sh = sessions_();
  var row = findRow_(sh, h, 1);
  if (row < 0) return null;
  var r = sh.getRange(row, 1, 1, 3).getValues()[0];
  var exp = Number(r[2]);
  if (exp < Date.now()) return null;     // baris kedaluwarsa dibersihkan saat login berikutnya (cleanupSessions_)
  cachePut_(ck, { email: String(r[1]), exp: exp }, Math.max(60, Math.min(21600, Math.floor((exp - Date.now()) / 1000))));
  return String(r[1]);
}

/* ------------------------------ AUTH ACTIONS ----------------------------- */

function register_(req) {
  var email = normEmail_(req.email);
  var nama = String(req.nama || '').trim().slice(0, 100);
  var password = String(req.password || '');
  if (!validEmail_(email)) return { ok: false, error: 'Format email tidak valid.' };
  if (!nama) return { ok: false, error: 'Nama wajib diisi.' };
  if (password.length < 6) return { ok: false, error: 'Password minimal 6 karakter.' };

  var sh = users_();
  if (findRow_(sh, email, 1) > 0) return { ok: false, error: 'Email sudah terdaftar. Silakan masuk.' };

  var salt = Utilities.getUuid();
  sh.appendRow([email, nama, salt, hashPassword_(password, salt), new Date()]);
  var s = createSession_(email);
  return { ok: true, token: s.token, expires: s.expires, email: email, nama: nama, periodes: publicPeriodes_() };
}

function login_(req) {
  var email = normEmail_(req.email);
  var password = String(req.password || '');
  var cache = CacheService.getScriptCache();
  var failKey = 'fail_' + sha256_(email);
  var fails = Number(cache.get(failKey) || 0);
  if (fails >= MAX_FAILS) return { ok: false, error: 'Terlalu banyak percobaan gagal. Coba lagi dalam 5 menit.' };

  var sh = users_();
  var row = findRow_(sh, email, 1);
  if (row < 0) { cache.put(failKey, String(fails + 1), LOCK_SECONDS); return { ok: false, error: 'Email atau password salah.' }; }
  var u = sh.getRange(row, 1, 1, 4).getValues()[0];
  if (hashPassword_(password, String(u[2])) !== String(u[3])) {
    cache.put(failKey, String(fails + 1), LOCK_SECONDS);
    return { ok: false, error: 'Email atau password salah.' };
  }
  cache.remove(failKey);
  var s = createSession_(email);
  return { ok: true, token: s.token, expires: s.expires, email: email, nama: String(u[1]), periodes: publicPeriodes_() };
}

function logout_(req) {
  var sh = sessions_();
  var h = sha256_(String(req.token || ''));
  cacheDel_('s_' + h);
  var row = findRow_(sh, h, 1);
  if (row > 0) sh.deleteRow(row);
  return { ok: true };
}

function changePassword_(req) {
  var email = auth_(req.token);
  if (!email) return { ok: false, error: 'Sesi berakhir. Silakan masuk kembali.', auth: false };
  var oldPw = String(req.oldPassword || ''), newPw = String(req.newPassword || '');
  if (newPw.length < 6) return { ok: false, error: 'Password baru minimal 6 karakter.' };
  var sh = users_();
  var row = findRow_(sh, email, 1);
  var u = sh.getRange(row, 1, 1, 4).getValues()[0];
  if (hashPassword_(oldPw, String(u[2])) !== String(u[3])) return { ok: false, error: 'Password lama salah.' };
  var salt = Utilities.getUuid();
  sh.getRange(row, 3, 1, 2).setValues([[salt, hashPassword_(newPw, salt)]]);
  return { ok: true };
}

/** Lupa password: kirim kode 6 digit ke email (butuh izin MailApp; kuota ±100 email/hari). */
function forgot_(req) {
  var email = normEmail_(req.email);
  var sh = users_();
  if (findRow_(sh, email, 1) > 0) {
    var code = String(Math.floor(100000 + Math.random() * 900000));
    var rs = resets_();
    var old = findRow_(rs, email, 1);
    if (old > 0) rs.deleteRow(old);
    rs.appendRow([email, sha256_(code + email), Date.now() + 15 * 60000]);
    MailApp.sendEmail({
      to: email,
      subject: 'Kode reset password - Jurnal Harian Guru Digital',
      body: 'Kode reset password Anda: ' + code + '\n\nKode berlaku 15 menit. Abaikan email ini jika Anda tidak memintanya.'
    });
  }
  // Selalu balas sukses agar email yang terdaftar tidak bisa ditebak orang lain.
  return { ok: true };
}

function reset_(req) {
  var email = normEmail_(req.email);
  var code = String(req.code || '').trim();
  var newPw = String(req.newPassword || '');
  if (newPw.length < 6) return { ok: false, error: 'Password baru minimal 6 karakter.' };
  var rs = resets_();
  var row = findRow_(rs, email, 1);
  if (row < 0) return { ok: false, error: 'Kode salah atau sudah kedaluwarsa.' };
  var r = rs.getRange(row, 1, 1, 3).getValues()[0];
  if (Number(r[2]) < Date.now() || String(r[1]) !== sha256_(code + email)) {
    return { ok: false, error: 'Kode salah atau sudah kedaluwarsa.' };
  }
  var sh = users_();
  var urow = findRow_(sh, email, 1);
  if (urow < 0) return { ok: false, error: 'Akun tidak ditemukan.' };
  var salt = Utilities.getUuid();
  sh.getRange(urow, 3, 1, 2).setValues([[salt, hashPassword_(newPw, salt)]]);
  rs.deleteRow(row);
  return { ok: true };
}

/* ------------------------------ DATA ACTIONS ----------------------------- */

/** Ambil SEMUA data milik pengguna: { key: "json string" } */
function load_(req) {
  var email = auth_(req.token);
  if (!email) return { ok: false, error: 'Sesi berakhir. Silakan masuk kembali.', auth: false };

  var p = periodeSS_(req.periode);
  if (p.error) return { ok: false, error: p.error, auth: p.reauth ? false : undefined };

  var sh = data_(p.ss);
  var last = sh.getLastRow();
  var out = {}, versions = {};
  if (last >= 2) {
    var lastCol = sh.getLastColumn();
    var vals = sh.getRange(2, 1, last - 1, lastCol).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (String(vals[i][0]).toLowerCase() !== email) continue;
      var parts = [];
      for (var c = 3; c < vals[i].length; c++) {
        if (vals[i][c] === '' || vals[i][c] === null) break;
        parts.push(String(vals[i][c]));
      }
      out[String(vals[i][1])] = parts.join('');
      versions[String(vals[i][1])] = Number(vals[i][2]) || 0;
    }
  }
  return { ok: true, email: email, data: out, versions: versions };
}

/** Simpan SATU jenis data (menggantikan versi sebelumnya). */
function save_(req) {
  var email = auth_(req.token);
  if (!email) return { ok: false, error: 'Sesi berakhir. Silakan masuk kembali.', auth: false };
  var key = String(req.key || '');
  if (ALLOWED_KEYS.indexOf(key) < 0) return { ok: false, error: 'Jenis data tidak valid.' };
  var text = String(req.data == null ? '' : req.data);
  if (text.length > 4000000) return { ok: false, error: 'Data terlalu besar.' };

  var chunks = [];
  for (var i = 0; i < text.length; i += CHUNK_SIZE) chunks.push(text.substr(i, CHUNK_SIZE));
  if (!chunks.length) chunks.push('');

  var p = periodeSS_(req.periode);
  if (p.error) return { ok: false, error: p.error, auth: p.reauth ? false : undefined };

  var sh = data_(p.ss);
  var needCols = 3 + chunks.length;
  if (sh.getMaxColumns() < needCols) sh.insertColumnsAfter(sh.getMaxColumns(), needCols - sh.getMaxColumns());

  // cari baris email+key
  var last = sh.getLastRow(), row = -1;
  if (last >= 2) {
    var ek = sh.getRange(2, 1, last - 1, 2).getValues();
    for (var r = 0; r < ek.length; r++) {
      if (String(ek[r][0]).toLowerCase() === email && String(ek[r][1]) === key) { row = r + 2; break; }
    }
  }
  if (row < 0) row = last + 1 < 2 ? 2 : last + 1;

  var now = Date.now();
  var width = Math.max(needCols, sh.getLastColumn());
  var rowVals = [email, key, now].concat(chunks);
  while (rowVals.length < width) rowVals.push('');   // kosongkan sisa chunk lama
  sh.getRange(row, 1, 1, width).setNumberFormat('@').setValues([rowVals]);
  return { ok: true, updatedAt: now };
}

/* --------------------- SUMBER DATA SISWA: SHEET1 ---------------------- */
/**
 * Membaca daftar siswa dari sheet SISWA_SHEET (default "Sheet1").
 * Baris 1 = judul kolom: No, NIS, NISN, Nama, Jenis Kelamin, Kelas (urutan bebas, huruf besar/kecil bebas).
 * Kolom OPSIONAL "Email Guru": jika diisi, baris itu hanya bisa diambil oleh guru dengan email tersebut
 * (boleh beberapa email, pisahkan dengan koma). Jika kosong, baris bisa diambil oleh semua guru yang login.
 */
function sheetStudents_(req) {
  var email = auth_(req.token);
  if (!email) return { ok: false, error: 'Sesi berakhir. Silakan masuk kembali.', auth: false };

  var p = periodeSS_(req.periode);
  if (p.error) return { ok: false, error: p.error, auth: p.reauth ? false : undefined };

  var sh = p.ss.getSheetByName(SISWA_SHEET);
  if (!sh) return { ok: false, error: 'Sheet "' + SISWA_SHEET + '" tidak ditemukan di spreadsheet periode ini.' };

  var last = sh.getLastRow(), cols = sh.getLastColumn();
  if (last < 2 || cols < 1) return { ok: true, rows: [], total: 0 };

  var vals = sh.getRange(1, 1, last, cols).getDisplayValues();   // teks seperti yang tampil (angka 0 di depan NISN aman)
  var headers = vals[0].map(function (h) { return String(h).trim(); });
  var ownerCol = -1;
  headers.forEach(function (h, i) { if (ownerCol < 0 && h.toLowerCase() === 'email guru') ownerCol = i; });

  var rows = [];
  for (var r = 1; r < vals.length; r++) {
    var row = vals[r];
    if (row.every(function (c) { return String(c).trim() === ''; })) continue;
    if (ownerCol >= 0) {
      var owner = String(row[ownerCol]).trim().toLowerCase();
      if (owner && owner.split(/[,;\s]+/).indexOf(email) < 0) continue;   // milik guru lain
    }
    var o = {};
    for (var c = 0; c < headers.length; c++) {
      if (headers[c] && c !== ownerCol) o[headers[c]] = row[c];
    }
    rows.push(o);
    if (rows.length >= 5000) break;
  }
  return { ok: true, rows: rows, total: rows.length };
}