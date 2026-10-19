/**
 * ULANGAN ONLINE — BACKEND GOOGLE SPREADSHEET (Google Apps Script) · versi multi-guru
 *
 * CARA PAKAI SINGKAT (lengkap di PANDUAN.md):
 * 1. Buat Google Spreadsheet baru → Ekstensi → Apps Script, tempel seluruh isi file ini.
 * 2. GANTI PASSWORD_ADMIN_AWAL di bawah.
 * 3. Jalankan fungsi  setupSpreadsheet  satu kali (beri izin akses).
 * 4. Terapkan → Deployment baru → Aplikasi web (Jalankan sebagai: Saya · Akses: Siapa saja).
 * 5. Salin URL aplikasi web ke  API_URL  di index.html.
 * 6. Login ke halaman guru dengan username  admin  dan password awal tadi, lalu tambah akun guru.
 *
 * Catatan: jangan mengedit sheet secara manual saat ulangan berlangsung; gunakan halaman guru.
 */

const PASSWORD_ADMIN_AWAL = 'admin123*';   // <-- WAJIB diganti, hanya dipakai saat setup
const MAPEL = ['Pend. Agama', 'Pend. Pancasila', 'B. Indonesia', 'Matematika', 'IPA', 'IPS', 'B. Inggris', 'Seni Budaya', 'PJOK', 'Informatika', 'B. Sunda'];
const TINGKAT = ['7', '8', '9'];

const SH = { AKUN: 'Akun', UL: 'Ulangan', SOAL: 'Soal', PES: 'Peserta' };
const HEAD = {
  AKUN: ['Username', 'Nama', 'Password (hash)', 'Peran', 'Aktif', 'Dibuat'],
  UL: ['Kode', 'Judul', 'Pemilik', 'Durasi (menit)', 'Jumlah Soal (0=semua)', 'Maks Pelanggaran', 'Dibuka', 'Tampil Nilai', 'KKM', 'Daftar Kelas', 'Maks No Absen', 'Dibuat', 'Acak Berimbang Level', 'Urut Mudah-Sulit', 'Mata Pelajaran', 'Tingkat', 'Jumlah Bank Soal'],
  SOAL: ['Kode', 'No', 'Soal', 'A', 'B', 'C', 'D', 'Kunci', 'Level', 'URL Gambar'],
  PES: ['Token', 'Kunci Siswa', 'Kode', 'Nama', 'Kelas', 'Absen', 'Status', 'Mulai', 'Soal Terpilih', 'Selesai', 'Aktivitas Terakhir', 'Terjawab', 'Soal Saat Ini', 'Pelanggaran', 'Log Pelanggaran', 'Jawaban', 'Benar', 'Salah', 'Kosong', 'Nilai', 'Durasi (menit)', 'Alasan Selesai']
};
const NCOL = { AKUN: 6, UL: 17, SOAL: 10, PES: 22 };
const KELAS_AWAL = 'A,B,C,D,E,F,G,H,I';   // huruf kelas paralel; tingkat (7/8/9) dipilih per ulangan

/* ====================== SETUP (jalankan sekali) ====================== */
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ensure = (name, header) => {
    let sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#12263A').setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }
    return sh;
  };
  const akun = ensure(SH.AKUN, HEAD.AKUN), ul = ensure(SH.UL, HEAD.UL), soal = ensure(SH.SOAL, HEAD.SOAL),
        pes = ensure(SH.PES, HEAD.PES);
  const teks = (sh, cols) => cols.forEach(c => sh.getRange(c + ':' + c).setNumberFormat('@'));   // teks murni: mencegah rumus & "05" jadi 5
  const waktu = (sh, cols) => cols.forEach(c => sh.getRange(c + ':' + c).setNumberFormat('dd/MM/yyyy HH:mm:ss'));
  teks(akun, ['A', 'B']); teks(ul, ['A', 'B', 'J']); teks(soal, ['A', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']);
  teks(pes, ['A', 'B', 'C', 'D', 'E', 'F', 'I', 'O', 'P', 'V']); waktu(pes, ['H', 'J', 'K']); waktu(akun, ['F']); waktu(ul, ['L']);
  if (!String(pes.getRange(1, 17).getValue())) pes.getRange(1, 17, 1, 6).setValues([HEAD.PES.slice(16)]).setFontWeight('bold').setBackground('#12263A').setFontColor('#ffffff');   // migrasi versi sebelumnya
  soal.setColumnWidth(3, 380);

  for (let c = 12; c < HEAD.UL.length; c++) {   // migrasi versi sebelumnya: lengkapi judul kolom yang belum ada
    if (!String(ul.getRange(1, c + 1).getValue())) ul.getRange(1, c + 1).setValue(HEAD.UL[c]).setFontWeight('bold').setBackground('#12263A').setFontColor('#ffffff');
  }
  teks(ul, ['M', 'N', 'O', 'P']);
  if (akun.getLastRow() <= 1) akun.appendRow(['admin', 'Administrator', buatHash(PASSWORD_ADMIN_AWAL), 'admin', 'YA', new Date()]);
  ['Sheet1', 'Sheet 1', 'Lembar1', 'Lembar 1'].forEach(n => {
    const d = ss.getSheetByName(n);
    if (d && ss.getSheets().length > 1 && d.getLastRow() === 0) ss.deleteSheet(d);
  });
  Logger.log('Setup selesai. Buat Deployment → Aplikasi web, lalu login sebagai admin.');
}

/* ====================== HTTP ====================== */
function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    if (p.action === 'config') return out(aksiConfig({ kode: p.kode }));
    return out({ ok: true, info: 'API ulangan online aktif' });
  } catch (err) { return out(galat(err)); }
}
function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const r = ROUTES[req.action];
    if (!r) return out({ ok: false, error: 'AKSI_TIDAK_DIKENAL' });
    const user = r.auth ? auth(req.t, r.auth) : null;
    return out(r.fn(req, user));
  } catch (err) { return out(galat(err)); }
}
function galat(err) {
  const m = String((err && err.message) || err);
  return { ok: false, error: /^[A-Z_]+(:.*)?$/.test(m) ? m : 'SERVER: ' + m };
}
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

const ROUTES = {
  config: { fn: aksiConfig },
  ul_publik: { fn: aksiUlPublik },
  mulai: { fn: aksiMulai },
  sync: { fn: aksiSync },
  selesai: { fn: aksiSelesai },
  login: { fn: aksiLogin },
  logout: { fn: (r) => { C().remove('tok:' + r.t); return { ok: true }; } },
  akun_ganti_password: { fn: aksiGantiPassword, auth: 1 },
  ul_list: { fn: aksiUlList, auth: 1 },
  ul_buat: { fn: aksiUlBuat, auth: 1 },
  ul_get: { fn: aksiUlGet, auth: 1 },
  ul_simpan: { fn: aksiUlSimpan, auth: 1 },
  ul_hapus: { fn: aksiUlHapus, auth: 1 },
  ul_soal_simpan: { fn: aksiSoalSimpan, auth: 1 },
  ul_monitor: { fn: aksiMonitor, auth: 1 },
  ul_hasil: { fn: aksiHasil, auth: 1 },
  ul_reset_siswa: { fn: aksiResetSiswa, auth: 1 },
  ul_kosongkan: { fn: aksiKosongkanUl, auth: 1 },
  ad_akun_list: { fn: aksiAkunList, auth: 2 },
  ad_akun_tambah: { fn: aksiAkunTambah, auth: 2 },
  ad_akun_reset: { fn: aksiAkunReset, auth: 2 },
  ad_akun_aktif: { fn: aksiAkunAktif, auth: 2 },
  ad_akun_hapus: { fn: aksiAkunHapus, auth: 2 },
  ad_ringkasan: { fn: aksiRingkasan, auth: 2 },
  ad_kosongkan: { fn: aksiKosongkanDb, auth: 2 }
};

/* ====================== UTIL ====================== */
const sheet = n => SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n);
const C = () => CacheService.getScriptCache();
function cget(k) { const v = C().get(k); if (!v) return null; try { return JSON.parse(v); } catch (e) { return null; } }
function cput(k, o, ttl) { try { C().put(k, JSON.stringify(o), ttl || 600); } catch (e) { /* terlalu besar: lewati cache */ } }
function bersihCache(kode) { C().removeAll(['c:' + kode, 'sl:' + kode, 'm:' + kode]); }
function withLock(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) throw new Error('SERVER_SIBUK');
  try { return fn(); } finally { lock.releaseLock(); }
}
const toMs = d => (d instanceof Date ? d.getTime() : Number(d) || 0);
const pad2 = n => { n = String(n); return n.length < 2 ? '0' + n : n; };
const norm = s => String(s || '').toLowerCase().replace(/\s+/g, ' ').trim();
const ya = b => (b ? 'YA' : 'TIDAK');

function cari(sh, kolom, nilai) {   // TextFinder: cepat, tanpa membaca seluruh sheet
  if (nilai === '' || nilai == null) return 0;
  const f = sh.getRange(kolom + ':' + kolom).createTextFinder(String(nilai)).matchEntireCell(true).matchCase(true).findNext();
  return f && f.getRow() > 1 ? f.getRow() : 0;
}
function bacaSemua(name, ncol) {
  const sh = sheet(name), last = sh.getLastRow();
  return { sh, last, rows: last > 1 ? sh.getRange(2, 1, last - 1, ncol).getValues() : [] };
}
function hapusBaris(name, idx, ncol, nilai) {   // tulis ulang tanpa baris yang cocok (tanpa deleteRow)
  const { sh, last, rows } = bacaSemua(name, ncol);
  const sisa = rows.filter(r => String(r[idx]) !== String(nilai));
  if (sisa.length === rows.length) return 0;
  sh.getRange(2, 1, last - 1, ncol).clearContent();
  if (sisa.length) sh.getRange(2, 1, sisa.length, ncol).setValues(sisa);
  return rows.length - sisa.length;
}
function kosongkanSheet(name, ncol) {
  const sh = sheet(name), last = sh.getLastRow();
  if (last < 2) return 0;
  sh.getRange(2, 1, last - 1, ncol).clearContent();
  if (sh.getMaxRows() > 2) sh.deleteRows(3, sh.getMaxRows() - 2);
  return last - 1;
}

/* --- jawaban: "no:H;no:H" --- */
function parseJ(s) { const o = {}; String(s || '').split(';').forEach(p => { const m = p.split(':'); if (m.length === 2 && /^\d+$/.test(m[0]) && /^[A-D]$/.test(m[1])) o[m[0]] = m[1]; }); return o; }
function strJ(o) { return Object.keys(o).map(n => n + ':' + o[n]).join(';'); }
function idsDari(s) { return String(s || '').split(';').map(Number).filter(x => x > 0); }

/* --- password --- */
function hashPw(pw, salt) {
  const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + pw, Utilities.Charset.UTF_8);
  return b.map(x => ('0' + (x & 255).toString(16)).slice(-2)).join('');
}
function buatHash(pw) { const s = Utilities.getUuid().replace(/-/g, '').slice(0, 16); return s + ':' + hashPw(pw, s); }
function cekHash(pw, simpan) { const p = String(simpan).split(':'); return p.length === 2 && hashPw(pw, p[0]) === p[1]; }

/* ====================== ULANGAN & SOAL (dengan cache) ====================== */
function ulDariRow(r) {
  return {
    kode: String(r[0]), judul: String(r[1]), pemilik: String(r[2]), durasi: Number(r[3]) || 30, jumlah: Number(r[4]) || 0,
    maks: Number(r[5]) || 3, aktif: String(r[6]).toUpperCase() === 'YA', tampil_nilai: String(r[7]).toUpperCase() === 'YA',
    kkm: Number(r[8]) || 75, kelas: String(r[9]).split(',').map(s => s.trim().toUpperCase()).filter(Boolean),
    maxAbsen: Number(r[10]) || 32, dibuat: toMs(r[11]),
    acakLevel: String(r[12]).toUpperCase() === 'YA', urutLevel: String(r[13]).toUpperCase() === 'YA',
    mapel: String(r[14] || ''), tingkat: String(r[15] || ''), bank: Number(r[16]) || 0
  };
}
function bacaUlangan(kode) {
  kode = String(kode || '').trim().toUpperCase();
  if (!/^[A-Z0-9]{4,10}$/.test(kode)) return null;
  const c = cget('c:' + kode); if (c) return c;
  const sh = sheet(SH.UL), row = cari(sh, 'A', kode);
  if (!row) return null;
  const u = ulDariRow(sh.getRange(row, 1, 1, NCOL.UL).getValues()[0]);
  cput('c:' + kode, u, 600);
  return u;
}
function soalDariRow(r) {
  return { no: Number(r[1]), soal: String(r[2]), pilihan: [r[3], r[4], r[5], r[6]].map(String).filter(x => x !== ''),
           kunci: String(r[7]).trim().toUpperCase(), level: String(r[8] || ''), gambar: String(r[9] || '') };
}
function soalLengkap(kode) {   // lengkap dengan kunci — hanya dipakai di server
  const c = cget('sl:' + kode); if (c) return c;
  const b = [];
  bacaSemua(SH.SOAL, NCOL.SOAL).rows.forEach(r => { if (String(r[0]) === kode) b.push(soalDariRow(r)); });
  b.sort((x, y) => x.no - y.no);
  cput('sl:' + kode, b, 600);
  return b;
}
function hangatkanSoal(kodes) {   // isi cache soal untuk banyak kode sekaligus (satu kali baca sheet)
  const perlu = kodes.filter(k => !C().get('sl:' + k)); if (!perlu.length) return;
  const b = {}; perlu.forEach(k => b[k] = []);
  bacaSemua(SH.SOAL, NCOL.SOAL).rows.forEach(r => { const k = String(r[0]); if (b[k]) b[k].push(soalDariRow(r)); });
  perlu.forEach(k => { b[k].sort((x, y) => x.no - y.no); cput('sl:' + k, b[k], 600); });
}
function bersihDaftar() { const k = []; MAPEL.forEach(m => TINGKAT.forEach(t => k.push('p:' + m + '|' + t))); C().removeAll(k); }
const jumlahTampil = (u, total) => (u.jumlah > 0 ? Math.min(u.jumlah, total) : total);
function soalDariIds(bank, idsStr) { const ids = idsDari(idsStr); return ids.length ? ids.map(n => bank.find(q => q.no === n)).filter(Boolean) : bank; }
function soalPublik(bank, idsStr) { return soalDariIds(bank, idsStr).map(q => ({ no: q.no, soal: q.soal, pilihan: q.pilihan, gambar: q.gambar })); }
/* Pemilihan soal per siswa.
 *  - berimbang: jatah tiap level proporsional dengan bank soal (metode sisa terbesar, seri dipecah urut level)
 *    sehingga komposisi level SAMA untuk semua siswa; soal di dalam level dipilih acak.
 *  - urut: soal diurutkan dari level terendah ke tertinggi; urutan di dalam level tetap acak.
 *  Tanpa opsi atau tanpa data level: acak biasa. */
function levelKey(l) { l = String(l || '').trim().toUpperCase(); if (!l) return 1e9; const m = l.match(/\d+/); return m ? Number(m[0]) : 1e8; }
function kocok(arr) { const a = arr.slice(); for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); const x = a[i]; a[i] = a[j]; a[j] = x; } return a; }
function kuotaLevel(ukuran, n) {   // ukuran: [jumlah soal tiap grup]; hasil: jatah tiap grup, jumlah = n
  const total = ukuran.reduce((a, b) => a + b, 0), q = ukuran.map(u => Math.floor(n * u / total));
  let sisa = n - q.reduce((a, b) => a + b, 0);
  ukuran.map((u, i) => ({ i, r: n * u / total - q[i] })).sort((x, y) => y.r - x.r || x.i - y.i).forEach(o => { if (sisa > 0 && q[o.i] < ukuran[o.i]) { q[o.i]++; sisa--; } });
  return q;
}
function acak(bank, n, opt) {
  opt = opt || {};
  let pilih;
  if (opt.berimbang && n < bank.length && bank.some(q => String(q.level || '').trim())) {
    const g = {};
    bank.forEach(q => { const k = String(q.level || '').trim().toUpperCase(); (g[k] = g[k] || []).push(q); });
    const kunci = Object.keys(g).sort((a, b) => levelKey(a) - levelKey(b) || (a < b ? -1 : 1));
    const kuota = kuotaLevel(kunci.map(k => g[k].length), n);
    pilih = [];
    kunci.forEach((k, i) => kocok(g[k]).slice(0, kuota[i]).forEach(q => pilih.push(q)));
  } else pilih = kocok(bank).slice(0, n);
  pilih = opt.urut ? pilih.sort((a, b) => levelKey(a.level) - levelKey(b.level)) : kocok(pilih);   // sort stabil: urutan acak di dalam level terjaga
  return pilih.map(q => q.no);
}
function nilaiDari(jw, soal) {
  let benar = 0, salah = 0, kosong = 0;
  soal.forEach(q => { const j = jw[q.no]; if (!j) kosong++; else if (j === q.kunci) benar++; else salah++; });
  return { benar, salah, kosong, nilai: Math.round((benar / soal.length) * 10000) / 100 };
}

/* ====================== SISWA ====================== */
function aksiConfig(req) {
  const u = bacaUlangan(req.kode);
  if (!u) throw new Error('KODE_TIDAK_ADA');
  const total = soalLengkap(u.kode).length;
  return { ok: true, kode: u.kode, judul: u.judul, mapel: u.mapel, tingkat: u.tingkat, durasi: u.durasi, jumlah: jumlahTampil(u, total), maks: u.maks,
           aktif: u.aktif && total > 0, kelas: u.kelas, maxAbsen: u.maxAbsen, serverNow: Date.now() };
}

/* Daftar ulangan yang sedang dibuka untuk satu mata pelajaran + tingkat (dipakai halaman siswa). Di-cache 2 menit; dibersihkan tiap guru mengubah ulangan. */
function aksiUlPublik(req) {
  const mapel = String(req.mapel || ''), tingkat = String(req.tingkat || '');
  if (MAPEL.indexOf(mapel) < 0 || TINGKAT.indexOf(tingkat) < 0) throw new Error('PILIHAN_TIDAK_VALID');
  const ck = 'p:' + mapel + '|' + tingkat, c = cget(ck); if (c) return c;
  const nama = {}; bacaSemua(SH.AKUN, 2).rows.forEach(r => { nama[String(r[0]).toLowerCase()] = String(r[1]); });
  const list = bacaSemua(SH.UL, NCOL.UL).rows.map(ulDariRow).filter(u => u.aktif && u.bank > 0 && u.mapel === mapel && u.tingkat === tingkat)
    .sort((a, b) => b.dibuat - a.dibuat)
    .map(u => ({ kode: u.kode, judul: u.judul, mapel: u.mapel, tingkat: u.tingkat, durasi: u.durasi, jumlah: jumlahTampil(u, u.bank), maks: u.maks,
                 kelas: u.kelas, maxAbsen: u.maxAbsen, aktif: true, guru: nama[u.pemilik.toLowerCase()] || u.pemilik }));
  hangatkanSoal(list.map(u => u.kode));   // pemanasan cache: klik "Mulai" serentak tidak membaca spreadsheet
  const o = { ok: true, ulangan: list, serverNow: Date.now() };
  cput(ck, o, 120);
  return o;
}

function cariToken(token) {
  token = String(token || '');
  if (token.length < 20) return null;
  const sh = sheet(SH.PES);
  let row = Number(C().get('r:' + token)) || 0;
  if (row >= 2) { const v = sh.getRange(row, 1, 1, NCOL.PES).getValues()[0]; if (v[0] === token) return { sh, row, v }; }
  row = cari(sh, 'A', token);
  if (!row) return null;
  C().put('r:' + token, String(row), 21600);
  return { sh, row, v: sh.getRange(row, 1, 1, NCOL.PES).getValues()[0] };
}

/* MULAI tanpa lock: dari 300 siswa yang menekan tombol bersamaan, semuanya diproses paralel.
 * Duplikat (kelas+absen yang sama) diselesaikan sesudah penulisan: baris paling atas menang, yang lain ditandai RESET. */
function aksiMulai(req) {
  const u = bacaUlangan(req.kode);
  if (!u) throw new Error('KODE_TIDAK_ADA');
  const nama = String(req.nama || '').trim().replace(/\s+/g, ' ');
  const kelas = String(req.kelas || '').trim().toUpperCase();
  const absen = pad2(req.absen);
  if (!/^[\p{L}\p{N}][\p{L}\p{M}\p{N}.' -]{2,59}$/u.test(nama)) throw new Error('NAMA_TIDAK_VALID');
  if (!u.tingkat || kelas.indexOf(u.tingkat) !== 0 || u.kelas.indexOf(kelas.slice(u.tingkat.length)) < 0) throw new Error('KELAS_TIDAK_VALID');
  if (!/^\d{1,2}$/.test(String(req.absen)) || Number(req.absen) < 1 || Number(req.absen) > u.maxAbsen) throw new Error('ABSEN_TIDAK_VALID');
  const key = u.kode + '|' + kelas + '|' + absen;
  const bank = soalLengkap(u.kode), sh = sheet(SH.PES), now = Date.now();
  const dasar = { ok: true, judul: u.judul, durasi: u.durasi, maks: u.maks, serverNow: now };
  const lanjut = row => {
    const v = sh.getRange(row, 1, 1, NCOL.PES).getValues()[0];
    if (v[6] === 'SELESAI') throw new Error('SUDAH_SELESAI');
    if (norm(v[3]) !== norm(nama)) throw new Error('ABSEN_DIPAKAI:' + v[3]);
    C().put('r:' + v[0], String(row), 21600);
    return Object.assign({}, dasar, { lanjut: true, token: v[0], soal: soalPublik(bank, v[8]), jawaban: parseJ(v[15]), mulai: toMs(v[7]), pelanggaran: Number(v[13]) || 0 });
  };
  const ada = cari(sh, 'B', key);
  if (ada) return lanjut(ada);
  if (!u.aktif) throw new Error('UJIAN_DITUTUP');
  if (!bank.length) throw new Error('TIDAK_ADA_SOAL');
  const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  const ids = acak(bank, jumlahTampil(u, bank.length), { berimbang: u.acakLevel, urut: u.urutLevel }).join(';');
  sh.appendRow([token, key, u.kode, nama, kelas, absen, 'BERLANGSUNG', new Date(now), ids, '', new Date(now), 0, 1, 0, '', '', '', '', '', '', '', '']);
  const mine = cari(sh, 'A', token);
  if (!mine) throw new Error('SERVER_SIBUK');   // baris tidak ditemukan: klien mencoba lagi
  const semua = sh.getRange('B:B').createTextFinder(key).matchEntireCell(true).matchCase(true).findAll().map(x => x.getRow()).filter(r => r > 1).sort((a, b) => a - b);
  if (semua.length && semua[0] !== mine) {   // kalah balapan dengan permintaan lain untuk siswa yang sama
    sh.getRange(mine, 2).setValue(''); sh.getRange(mine, 7).setValue('RESET');
    return lanjut(semua[0]);
  }
  C().put('r:' + token, String(mine), 21600);
  return Object.assign({}, dasar, { lanjut: false, token, soal: soalPublik(bank, ids), jawaban: {}, mulai: now, pelanggaran: 0 });
}

/* Simpan massal: satu panggilan membawa seluruh jawaban + denyut jantung. Tanpa lock, hanya menulis baris milik siswa itu. */
function aksiSync(req) {
  const p = cariToken(req.token);
  if (!p || p.v[6] === 'RESET') throw new Error('SESI_TIDAK_ADA');
  if (p.v[6] !== 'BERLANGSUNG') throw new Error('SUDAH_SELESAI');
  const jw = parseJ(req.jawaban);
  p.sh.getRange(p.row, 11, 1, 6).setValues([[new Date(), Object.keys(jw).length, Number(req.posisi) || 1,
    Math.max(Number(p.v[13]) || 0, Number(req.pelanggaran) || 0), String(req.log || '').slice(0, 900), strJ(jw)]]);
  return { ok: true, serverNow: Date.now() };
}

/* Hasil ditulis ke baris Peserta milik siswa itu sendiri (tanpa lock). Hasil dihitung deterministik, jadi pengiriman ganda aman. */
function selesaikan(p, jw, alasan, soal) {
  const bersih = {};
  soal.forEach(q => { const j = jw[q.no]; if (j && 'ABCD'.indexOf(j) < q.pilihan.length) bersih[q.no] = j; });
  const hit = nilaiDari(bersih, soal), now = new Date();
  const durasi = Math.max(0, Math.round((now.getTime() - toMs(p.v[7])) / 600) / 100), pel = Number(p.v[13]) || 0;
  p.sh.getRange(p.row, 10, 1, 13).setValues([[now, now, soal.length - hit.kosong, p.v[12], pel, p.v[14], strJ(bersih), hit.benar, hit.salah, hit.kosong, hit.nilai, durasi, alasan]]);
  p.sh.getRange(p.row, 7).setValue('SELESAI');   // terakhir: pembaca yang melihat SELESAI pasti sudah melihat hasilnya
  return Object.assign({ durasi }, hit);
}

function aksiSelesai(req) {
  const p = cariToken(req.token);
  if (!p || p.v[6] === 'RESET') throw new Error('SESI_TIDAK_ADA');
  const u = bacaUlangan(p.v[2]) || { tampil_nilai: false };
  const keluar = r => { const o = { ok: true, tampil_nilai: u.tampil_nilai }; if (u.tampil_nilai && r) { o.nilai = r.nilai; o.benar = r.benar; o.salah = r.salah; o.kosong = r.kosong; } return o; };
  if (p.v[6] === 'SELESAI') return keluar({ nilai: Number(p.v[19]), benar: p.v[16], salah: p.v[17], kosong: p.v[18] });
  const bank = soalLengkap(p.v[2]), soal = soalDariIds(bank, p.v[8]);
  const jw = Object.assign(parseJ(p.v[15]), parseJ(req.jawaban));   // kiriman klien yang terbaru menang
  p.v[13] = Math.max(Number(p.v[13]) || 0, Number(req.pelanggaran) || 0);
  return keluar(selesaikan(p, jw, String(req.alasan || 'SELESAI').slice(0, 60), soal));
}

/* Selesaikan otomatis peserta yang waktunya habis tetapi tidak sempat mengirim (browser tertutup). */
function finalizeExpired(u) {
  const batas = u.durasi * 60000 + 120000, now = Date.now();
  const { sh, rows } = bacaSemua(SH.PES, 16), exp = [];
  rows.forEach((v, i) => { if (v[2] === u.kode && v[6] === 'BERLANGSUNG' && now - toMs(v[7]) > batas) exp.push({ sh, row: i + 2, v }); });
  if (!exp.length) return;
  const bank = soalLengkap(u.kode);
  exp.forEach(p => { if (sh.getRange(p.row, 7).getValue() === 'BERLANGSUNG') selesaikan(p, parseJ(p.v[15]), 'WAKTU HABIS (otomatis)', soalDariIds(bank, p.v[8])); });
}

/* ====================== AKUN & LOGIN ====================== */
function cariAkun(un) {
  const { sh, rows } = bacaSemua(SH.AKUN, NCOL.AKUN);
  for (let i = 0; i < rows.length; i++) if (String(rows[i][0]).toLowerCase() === un) return { sh, row: i + 2, v: rows[i] };
  return null;
}
function aksiLogin(req) {
  const un = String(req.username || '').trim().toLowerCase(), pw = String(req.password || '');
  const kk = 'gagal:' + un, gagal = Number(C().get(kk)) || 0;
  if (gagal >= 5) throw new Error('TERKUNCI');
  const a = cariAkun(un);
  if (!a || String(a.v[4]).toUpperCase() !== 'YA' || !cekHash(pw, a.v[2])) { C().put(kk, String(gagal + 1), 300); throw new Error('LOGIN_SALAH'); }
  C().remove(kk);
  const t = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
  const user = { username: String(a.v[0]), nama: String(a.v[1]), peran: String(a.v[3]) };
  cput('tok:' + t, user, 21600);
  return { ok: true, t, user };
}
function auth(t, level) {
  const u = t ? cget('tok:' + t) : null;
  if (!u) throw new Error('SESI_HABIS');
  if (level === 2 && u.peran !== 'admin') throw new Error('BUKAN_ADMIN');
  return u;
}
function aksiGantiPassword(req, user) {
  const baru = String(req.baru || '');
  if (baru.length < 6) throw new Error('PASSWORD_PENDEK');
  const a = cariAkun(user.username);
  if (!a || !cekHash(String(req.lama || ''), a.v[2])) throw new Error('PASSWORD_LAMA_SALAH');
  a.sh.getRange(a.row, 3).setValue(buatHash(baru));
  return { ok: true };
}
function aksiAkunList() {
  const { rows } = bacaSemua(SH.AKUN, NCOL.AKUN), uls = bacaSemua(SH.UL, NCOL.UL).rows, n = {};
  uls.forEach(r => n[r[2]] = (n[r[2]] || 0) + 1);
  return { ok: true, akun: rows.map(r => ({ username: String(r[0]), nama: String(r[1]), peran: String(r[3]), aktif: String(r[4]).toUpperCase() === 'YA', dibuat: toMs(r[5]), ulangan: n[r[0]] || 0 })) };
}
function aksiAkunTambah(req) {
  const un = String(req.username || '').trim().toLowerCase(), nama = String(req.nama || '').trim().slice(0, 60), pw = String(req.password || '');
  const peran = req.peran === 'admin' ? 'admin' : 'guru';
  if (!/^[a-z0-9._-]{3,30}$/.test(un)) throw new Error('USERNAME_TIDAK_VALID');
  if (nama.length < 2 || /^[=+\-@]/.test(nama)) throw new Error('NAMA_TIDAK_VALID');
  if (pw.length < 6) throw new Error('PASSWORD_PENDEK');
  return withLock(() => {
    if (cariAkun(un)) throw new Error('USERNAME_ADA');
    sheet(SH.AKUN).appendRow([un, nama, buatHash(pw), peran, 'YA', new Date()]);
    return { ok: true };
  });
}
function aksiAkunReset(req) {
  if (String(req.password || '').length < 6) throw new Error('PASSWORD_PENDEK');
  const a = cariAkun(String(req.username || '').toLowerCase());
  if (!a) throw new Error('TIDAK_DITEMUKAN');
  a.sh.getRange(a.row, 3).setValue(buatHash(String(req.password)));
  return { ok: true };
}
function aksiAkunAktif(req, user) {
  const un = String(req.username || '').toLowerCase();
  if (un === user.username) throw new Error('TIDAK_BOLEH_DIRI_SENDIRI');
  const a = cariAkun(un);
  if (!a) throw new Error('TIDAK_DITEMUKAN');
  a.sh.getRange(a.row, 5).setValue(ya(!!req.aktif));
  return { ok: true };
}
function aksiAkunHapus(req, user) {
  const un = String(req.username || '').toLowerCase();
  if (un === user.username) throw new Error('TIDAK_BOLEH_DIRI_SENDIRI');
  return withLock(() => {
    const a = cariAkun(un);
    if (!a) throw new Error('TIDAK_DITEMUKAN');
    const { sh, last, rows } = bacaSemua(SH.UL, NCOL.UL);   // ulangan milik guru ini dipindahkan ke admin yang menghapus
    if (rows.length) { sh.getRange(2, 3, last - 1, 1).setValues(rows.map(r => [String(r[2]) === un ? user.username : r[2]])); rows.forEach(r => bersihCache(String(r[0]))); }
    hapusBaris(SH.AKUN, 0, NCOL.AKUN, a.v[0]);
    return { ok: true };
  });
}

/* ====================== GURU: ULANGAN ====================== */
function akses(kode, user) {
  const u = bacaUlangan(kode);
  if (!u) throw new Error('KODE_TIDAK_ADA');
  if (user.peran !== 'admin' && u.pemilik !== user.username) throw new Error('TIDAK_BERHAK');
  return u;
}
function kodeBaru() {
  const abj = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789', sh = sheet(SH.UL);
  for (let k = 0; k < 20; k++) {
    let s = ''; for (let i = 0; i < 6; i++) s += abj.charAt(Math.floor(Math.random() * abj.length));
    if (!cari(sh, 'A', s)) return s;
  }
  throw new Error('GAGAL_KODE');
}
function aksiUlList(req, user) {
  const { rows } = bacaSemua(SH.UL, NCOL.UL), ps = bacaSemua(SH.PES, 7).rows;
  const jalan = {}, selesai = {};
  ps.forEach(r => { if (r[6] === 'RESET') return; if (r[6] === 'SELESAI') selesai[r[2]] = (selesai[r[2]] || 0) + 1; else jalan[r[2]] = (jalan[r[2]] || 0) + 1; });
  const list = rows.map(ulDariRow).filter(u => user.peran === 'admin' || u.pemilik === user.username)
    .map(u => Object.assign(u, { soal: u.bank, jalan: jalan[u.kode] || 0, selesai: selesai[u.kode] || 0 }));
  list.sort((a, b) => b.dibuat - a.dibuat);
  return { ok: true, ulangan: list };
}
function validUl(s, lama) {
  const judul = String(s.judul || (lama && lama.judul) || '').trim().slice(0, 120);
  if (judul.length < 3 || /^[=+\-@]/.test(judul)) throw new Error('JUDUL_TIDAK_VALID');
  const mapel = String(s.mapel || (lama && lama.mapel) || '').trim();
  if (MAPEL.indexOf(mapel) < 0) throw new Error('MAPEL_TIDAK_VALID');
  const tingkat = String(s.tingkat || (lama && lama.tingkat) || '').trim();
  if (TINGKAT.indexOf(tingkat) < 0) throw new Error('TINGKAT_TIDAK_VALID');
  const kelas = String(s.kelas == null ? KELAS_AWAL : s.kelas).split(',').map(x => x.trim().toUpperCase())
    .map(x => (x.length > 1 && x.indexOf(tingkat) === 0 ? x.slice(tingkat.length) : x)).filter(Boolean);   // "9A" → "A"
  if (!kelas.length || kelas.length > 40 || kelas.some(k => !/^[A-Z0-9]{1,3}$/.test(k))) throw new Error('KELAS_LIST_TIDAK_VALID');
  const n = (v, a, b, d) => { v = Number(v); return isFinite(v) && v >= a && v <= b ? Math.round(v) : d; };
  return { judul, mapel, tingkat, kelas: kelas.filter((k, i) => kelas.indexOf(k) === i).join(','), durasi: n(s.durasi, 5, 300, 30), jumlah: n(s.jumlah, 0, 500, 0),
           maks: n(s.maks, 1, 20, 3), kkm: n(s.kkm, 0, 100, 75), maxAbsen: n(s.maxAbsen, 1, 99, 32), aktif: !!s.aktif, tampil_nilai: !!s.tampil_nilai, acak_level: !!s.acak_level, urut_level: !!s.urut_level };
}
function aksiUlBuat(req, user) {
  const v = validUl({ judul: req.judul, mapel: req.mapel, tingkat: req.tingkat, aktif: false, tampil_nilai: true, durasi: 30, jumlah: 0, maks: 3, kkm: 75, kelas: KELAS_AWAL, maxAbsen: 32 });
  const pilihan = String(req.kode || '').trim().toUpperCase();
  if (pilihan && !/^[A-Z0-9]{4,10}$/.test(pilihan)) throw new Error('KODE_TIDAK_VALID');
  return withLock(() => {
    const sh = sheet(SH.UL);
    if (pilihan && cari(sh, 'A', pilihan)) throw new Error('KODE_ADA');
    const kode = pilihan || kodeBaru();
    sh.appendRow([kode, v.judul, user.username, v.durasi, v.jumlah, v.maks, ya(false), ya(true), v.kkm, v.kelas, v.maxAbsen, new Date(), ya(false), ya(false), v.mapel, v.tingkat, 0]);
    bersihDaftar();
    return { ok: true, kode };
  });
}
function aksiUlGet(req, user) {
  const u = akses(req.kode, user), soal = soalLengkap(u.kode);
  if (soal.length !== u.bank) {   // perbaiki hitungan bank soal bila sheet diedit manual
    const sh = sheet(SH.UL), row = cari(sh, 'A', u.kode); if (row) sh.getRange(row, 17).setValue(soal.length);
    bersihCache(u.kode); bersihDaftar(); u.bank = soal.length;
  }
  return { ok: true, ul: u, soal };
}
function gantiKodeSoal(lama, baru) {
  const { sh, last, rows } = bacaSemua(SH.SOAL, 1);
  if (rows.length) sh.getRange(2, 1, last - 1, 1).setValues(rows.map(r => [String(r[0]) === lama ? baru : r[0]]));
}
function aksiUlSimpan(req, user) {
  const lama = akses(req.kode, user), s = req.setting || {}, v = validUl(s, lama);
  const baru = String(s.kode || '').trim().toUpperCase();
  return withLock(() => {
    const sh = sheet(SH.UL), row = cari(sh, 'A', lama.kode);
    if (!row) throw new Error('KODE_TIDAK_ADA');
    let kode = lama.kode;
    const berisi = (baru && baru !== lama.kode) || (lama.tingkat && v.tingkat !== lama.tingkat) ? adaSiswa(lama.kode) : false;
    if (baru && baru !== lama.kode) {
      if (!/^[A-Z0-9]{4,10}$/.test(baru)) throw new Error('KODE_TIDAK_VALID');
      if (cari(sh, 'A', baru)) throw new Error('KODE_ADA');
      if (berisi) throw new Error('KODE_TERKUNCI');
      kode = baru;
    }
    if (lama.tingkat && v.tingkat !== lama.tingkat && berisi) throw new Error('TINGKAT_TERKUNCI');
    sh.getRange(row, 2, 1, 10).setValues([[v.judul, lama.pemilik, v.durasi, v.jumlah, v.maks, ya(v.aktif), ya(v.tampil_nilai), v.kkm, v.kelas, v.maxAbsen]]);
    sh.getRange(row, 13, 1, 4).setValues([[ya(v.acak_level), ya(v.urut_level), v.mapel, v.tingkat]]);
    if (kode !== lama.kode) { sh.getRange(row, 1).setValue(kode); gantiKodeSoal(lama.kode, kode); }
    bersihCache(lama.kode); bersihCache(kode); bersihDaftar();
    return { ok: true, ul: bacaUlangan(kode) };
  });
}
function adaSiswa(kode) {
  const { rows } = bacaSemua(SH.PES, 7);
  return rows.some(r => r[2] === kode && r[6] !== 'RESET');
}
function tandaiReset(kode) {   // tidak menghapus baris Peserta agar nomor baris tetap stabil saat ada siswa lain sedang menyimpan
  const { sh, last, rows } = bacaSemua(SH.PES, 7);
  let n = 0; const keys = [], sts = [];
  rows.forEach(r => { if (r[2] === kode && r[6] !== 'RESET') { keys.push(['']); sts.push(['RESET']); n++; } else { keys.push([r[1]]); sts.push([r[6]]); } });
  if (n) { sh.getRange(2, 2, last - 1, 1).setValues(keys); sh.getRange(2, 7, last - 1, 1).setValues(sts); }
  return n;
}
function aksiUlHapus(req, user) {
  const u = akses(req.kode, user);
  return withLock(() => {
    tandaiReset(u.kode); hapusBaris(SH.SOAL, 0, NCOL.SOAL, u.kode); hapusBaris(SH.UL, 0, NCOL.UL, u.kode);
    bersihCache(u.kode); bersihDaftar();
    return { ok: true };
  });
}
function aksiKosongkanUl(req, user) {
  const u = akses(req.kode, user);
  return withLock(() => { const n = tandaiReset(u.kode); bersihCache(u.kode); return { ok: true, dihapus: n }; });
}

function bersihSoal(arr) {
  if (!Array.isArray(arr) || !arr.length) throw new Error('SOAL_KOSONG');
  if (arr.length > 200) throw new Error('SOAL_TERLALU_BANYAK');
  return arr.map((q, i) => {
    const s = String(q.soal || '').trim().slice(0, 2000), no = i + 1;
    const op = [q.a, q.b, q.c, q.d].map(x => String(x == null ? '' : x).trim().slice(0, 500));
    const kunci = String(q.kunci || '').trim().toUpperCase(), gambar = String(q.gambar || '').trim().slice(0, 500);
    const n = op.filter(x => x !== '').length;
    if (!s) throw new Error('SOAL_KOSONG_BARIS:' + no);
    if (op[0] === '' || op[1] === '' || (op[2] === '' && op[3] !== '')) throw new Error('PILIHAN_TIDAK_VALID:' + no);
    if (!/^[A-D]$/.test(kunci) || 'ABCD'.indexOf(kunci) >= n) throw new Error('KUNCI_TIDAK_VALID:' + no);
    if (gambar && !/^https?:\/\/\S+$/i.test(gambar)) throw new Error('GAMBAR_TIDAK_VALID:' + no);
    return [no, s, op[0], op[1], op[2], op[3], kunci, String(q.level || '').trim().slice(0, 20), gambar];
  });
}
function aksiSoalSimpan(req, user) {
  const u = akses(req.kode, user), rows = bersihSoal(req.soal);
  return withLock(() => {
    if (adaSiswa(u.kode)) throw new Error('SUDAH_ADA_SISWA');
    hapusBaris(SH.SOAL, 0, NCOL.SOAL, u.kode);
    const sh = sheet(SH.SOAL), awal = sh.getLastRow() + 1;
    sh.getRange(awal, 1, rows.length, NCOL.SOAL).setValues(rows.map(r => [u.kode].concat(r)));
    const su = sheet(SH.UL), ru = cari(su, 'A', u.kode); if (ru) su.getRange(ru, 17).setValue(rows.length);
    bersihCache(u.kode); bersihDaftar();
    return { ok: true, jumlah: rows.length };
  });
}

function aksiMonitor(req, user) {
  const u = akses(req.kode, user);
  const c = cget('m:' + u.kode); if (c) return c;
  finalizeExpired(u);
  const now = Date.now(), nBank = soalLengkap(u.kode).length;
  const list = [];
  bacaSemua(SH.PES, 16).rows.forEach(r => {
    if (r[2] !== u.kode || r[6] === 'RESET') return;
    const mulai = toMs(r[7]);
    list.push({ nama: String(r[3]), kelas: String(r[4]), absen: pad2(r[5]), status: r[6], mulai, aktif: toMs(r[10]), terjawab: Number(r[11]) || 0,
      no: Number(r[12]) || 1, pelanggaran: Number(r[13]) || 0, log: String(r[14] || ''), total: idsDari(r[8]).length || nBank,
      sisa: r[6] === 'BERLANGSUNG' ? Math.max(0, u.durasi * 60000 - (now - mulai)) : 0 });
  });
  const o = { ok: true, serverNow: now, durasi: u.durasi, peserta: list };
  cput('m:' + u.kode, o, 5);
  return o;
}
function aksiHasil(req, user) {
  const u = akses(req.kode, user);
  finalizeExpired(u);
  const hasil = [];
  bacaSemua(SH.PES, NCOL.PES).rows.forEach(r => {
    if (r[2] !== u.kode || r[6] !== 'SELESAI') return;
    hasil.push({ waktu: toMs(r[9]), nama: String(r[3]), kelas: String(r[4]), absen: pad2(r[5]), benar: r[16], salah: r[17], kosong: r[18], nilai: Number(r[19]),
                 durasi: Number(r[20]), pelanggaran: Number(r[13]) || 0, alasan: String(r[21]), soal: idsDari(r[8]), jawaban: parseJ(r[15]) });
  });
  return { ok: true, ul: u, soal: soalLengkap(u.kode), hasil };
}
function aksiResetSiswa(req, user) {
  const u = akses(req.kode, user), key = u.kode + '|' + String(req.kelas).toUpperCase() + '|' + pad2(req.absen);
  const sh = sheet(SH.PES), row = cari(sh, 'B', key);
  if (!row) throw new Error('TIDAK_DITEMUKAN');
  const token = String(sh.getRange(row, 1).getValue());
  sh.getRange(row, 2).setValue(''); sh.getRange(row, 7).setValue('RESET');   // hasil ikut tidak tampil karena status bukan SELESAI
  C().remove('m:' + u.kode);
  return { ok: true };
}

/* ====================== ADMIN: DATABASE ====================== */
function aksiRingkasan() {
  const ps = bacaSemua(SH.PES, 7).rows, n = nm => Math.max(0, sheet(nm).getLastRow() - 1);
  return { ok: true, akun: n(SH.AKUN), ulangan: n(SH.UL), soal: n(SH.SOAL), peserta: ps.filter(r => r[6] !== 'RESET').length, hasil: ps.filter(r => r[6] === 'SELESAI').length };
}
function aksiKosongkanDb(req) {
  if (String(req.konfirmasi) !== 'KOSONGKAN') throw new Error('KONFIRMASI_SALAH');
  const mode = req.mode === 'semua' ? 'semua' : 'siswa';
  return withLock(() => {
    const kodes = bacaSemua(SH.UL, 1).rows.map(r => String(r[0])), ps = bacaSemua(SH.PES, 7).rows;
    const hasil = { ok: true, mode, peserta: ps.filter(r => r[6] !== 'RESET').length, hasil: ps.filter(r => r[6] === 'SELESAI').length, soal: 0, ulangan: 0 };
    kosongkanSheet(SH.PES, NCOL.PES);
    if (mode === 'semua') { hasil.soal = kosongkanSheet(SH.SOAL, NCOL.SOAL); hasil.ulangan = kosongkanSheet(SH.UL, NCOL.UL); }
    const keys = []; kodes.forEach(k => keys.push('c:' + k, 'sl:' + k, 'm:' + k));
    if (keys.length) C().removeAll(keys);
    bersihDaftar();
    return hasil;
  });
}