/* NESACIS — pustaka bersama semua halaman: sesi login terpadu + pemanggil API. */
(function () {
  'use strict';
  var CFG = window.NESACIS_CONFIG || {};
  var KEY = 'nesacis_session';

  var PESAN = {
    SESI_HABIS: 'Sesi berakhir. Silakan masuk kembali.',
    TIDAK_BERHAK: 'Akunmu tidak memiliki akses untuk tindakan ini.',
    SERVER_SIBUK: 'Server sedang sibuk. Coba lagi sebentar lagi.',
    TIDAK_DITEMUKAN: 'Data tidak ditemukan.',
    USERNAME_ADA: 'Username/NISN sudah dipakai.',
    USERNAME_TIDAK_VALID: 'Username 3–30 karakter: huruf kecil, angka, titik, garis bawah, atau strip.',
    NAMA_TIDAK_VALID: 'Nama tidak valid (3–60 karakter; huruf, angka, spasi, titik, strip).',
    PASSWORD_PENDEK: 'Password minimal 6 karakter.',
    TIDAK_BOLEH_DIRI_SENDIRI: 'Tidak bisa dilakukan pada akunmu sendiri.',
    KONFIRMASI_SALAH: 'Ketik KOSONGKAN dengan huruf besar untuk konfirmasi.'
  };

  function readSession() {
    try { return JSON.parse(localStorage.getItem(KEY) || 'null'); } catch (e) { return null; }
  }

  var N = window.NESACIS = {
    API_URL: String(CFG.API_URL || '').trim(),

    /** Sesi yang masih berlaku, atau null. */
    session: function () {
      var s = readSession();
      if (!s || !s.token || !s.user || Number(s.expires) < Date.now()) return null;
      return s;
    },
    save: function (res) {
      var s = { token: res.token, expires: res.expires, user: res.user };
      try { localStorage.setItem(KEY, JSON.stringify(s)); } catch (e) {}
      return s;
    },
    clear: function () { try { localStorage.removeItem(KEY); } catch (e) {} },
    updateUser: function (u) { var s = readSession(); if (s) { s.user = u; try { localStorage.setItem(KEY, JSON.stringify(s)); } catch (e) {} } },

    /** Pesan galat yang ramah. */
    pesan: function (err) {
      err = String(err || '');
      if (!err) return 'Terjadi kesalahan.';
      if (err.indexOf(' ') >= 0) return err;                       // sudah kalimat
      var i = err.indexOf(':'), k = i < 0 ? err : err.slice(0, i);
      if (PESAN[k]) return PESAN[k];
      return 'Gagal: ' + err;
    },

    /** Panggil API. Menyertakan token otomatis. Mengembalikan {ok:false,error} saat gagal jaringan. */
    call: function (action, data, opt) {
      opt = opt || {};
      if (!N.API_URL) return Promise.resolve({ ok: false, error: 'Aplikasi belum dihubungkan ke database. Isi API_URL di file config.js.' });
      var s = N.session(), body = Object.assign({ action: action }, s && !opt.tanpaToken ? { token: s.token } : {}, data || {});
      var ctl = new AbortController(), to = setTimeout(function () { ctl.abort(); }, opt.timeout || 60000);
      return fetch(N.API_URL, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(body), signal: ctl.signal, redirect: 'follow' })
        .then(function (r) { return r.text(); })
        .then(function (t) {
          var j; try { j = JSON.parse(t); } catch (e) { return { ok: false, error: 'Server membalas data yang tidak dikenali. Pastikan Web App sudah di-deploy dengan akses "Siapa saja" dan kode terbaru.' }; }
          if (j && j.ok === false && (j.auth === false || j.error === 'SESI_HABIS') && !opt.tanpaRedirect) {
            N.clear();
            if (!opt.diam) N.keLogin();
          }
          return j;
        })
        .catch(function (e) {
          return { ok: false, error: e && e.name === 'AbortError' ? 'Server terlalu lama merespons. Coba lagi.' : (navigator.onLine ? 'Tidak dapat terhubung ke server. Periksa URL di config.js dan koneksi internet.' : 'Perangkat sedang offline.') };
        })
        .then(function (r) { clearTimeout(to); return r; });
    },

    /** Arahkan ke portal untuk masuk, lalu kembali ke halaman ini. */
    keLogin: function (peran) {
      var next = location.pathname.split('/').pop() + location.search + location.hash;
      location.replace('index.html?masuk=' + encodeURIComponent(peran || 'guru') + '&next=' + encodeURIComponent(next));
    },

    /** Pastikan sudah login dengan salah satu peran; jika tidak, arahkan ke portal. Mengembalikan sesi atau null. */
    wajib: function (peranOk, peranLogin) {
      var s = N.session();
      if (!s) { N.keLogin(peranLogin || peranOk[0]); return null; }
      if (peranOk.indexOf(s.user.peran) < 0) { N.keLogin(peranLogin || peranOk[0]); return null; }
      return s;
    },

    keluar: function (ke) {
      var s = N.session();
      var done = function () { N.clear(); location.href = ke || 'index.html'; };
      if (!s) return done();
      N.call('logout', {}, { timeout: 6000, tanpaRedirect: true }).then(done, done);
    }
  };
})();
