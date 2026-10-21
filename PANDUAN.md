# Panduan Pemasangan NESACIS (versi terpadu)

## 1. Backend (Google Apps Script)
1. Buka script.google.com → **Proyek baru** (proyek mandiri, bukan dari dalam spreadsheet).
2. Buat 6 file skrip dan tempel isi dari folder `backend/`: `Code.gs`, `Akun.gs`, `Data.gs`, `Rekap.gs`, `Ulangan.gs`, `Materi.gs`, `Backup.gs`, `Cadangan.gs`.
3. Di `Code.gs` **ganti `PASSWORD_ADMIN_AWAL`**.
4. Pilih fungsi **`setup`** → Run → setujui izin (Spreadsheet, Drive, dan Mail). Setup otomatis membuat 6 spreadsheet
   (Akun, Jurnal, Nilai, Absensi, Ulangan, Materi), folder Drive untuk berkas materi, akun `admin`, dan satu periode awal.
   Alamat 6 file tampil di Log dan di Portal Admin > Ringkasan.
5. **Deploy → Deployment baru → Aplikasi web**: Execute as **Me**, Who has access **Anyone**. Salin URL `/exec`.
   Setiap mengubah kode skrip, buat **versi deployment baru**.

## 2. Frontend
1. Isi `API_URL` di **`config.js`** (satu tempat untuk semua halaman).
2. Unggah semua file HTML/JS/CSS/ikon ke hosting statis (HTTPS). Jika memakai `sw.js` lama, naikkan nomor versi cache-nya
   agar HP tidak menampilkan halaman lama.

## 3. Mulai dipakai
1. Portal → **Portal Admin** → login `admin`. Buat akun guru (tab Akun Guru), isi Data Siswa (impor Excel) dan cek Periode.
2. Guru: Portal Guru → Jurnal Harian / Ulangan Online. Materi diunggah di menu **Materi Pembelajaran** pada Jurnal Guru.
3. Murid: Portal Murid → Daftar (NISN + password sendiri) → Aku Ingin Belajar / Hasil Belajarku / Ujianku.

## Pembagian database
| File | Isi |
|---|---|
| Akun | Akun (admin/guru/murid), Siswa (daftar NISN), Periode, Sesi, Reset, Pengaturan |
| Jurnal | profil, siswa, jurnal, jadwal, catatan, pengaturan tiap guru |
| Nilai | nilai tiap guru |
| Absensi | absensi tiap guru |
| Ulangan | Ulangan, Soal, Peserta (+NISN) |
| Materi | metadata materi (berkas PDF/HTML ada di folder Google Drive) |

## Backup
- Admin: Portal Admin > Backup & Restore (semua database + berkas materi, bisa dipulihkan penuh).
- Guru: Jurnal Guru > Backup & Restore tetap berfungsi dan memulihkan data ke periode yang sedang aktif.
- Data lama dari aplikasi sebelumnya: guru cukup Export Backup di aplikasi lama, lalu Restore di Jurnal Guru yang baru.

## Fitur tambahan (pembaruan)
**Jika Anda sudah menjalankan `setup()` sebelumnya:** tempel kode terbaru, jalankan `setup()` sekali lagi (aman diulang; menambah sheet Rekap, kolom Perangkat, dan folder cadangan), lalu buat **versi deployment baru**.

| Fitur | Cara memakai |
|---|---|
| Cadangan otomatis | Portal Admin > Backup & Restore > Cadangan otomatis: aktifkan, pilih harian/mingguan, jam, jumlah cadangan disimpan. Pertama kali Google meminta izin "mengelola pemicu (trigger)": jalankan fungsi `pasangJadwalCadangan` dari editor Apps Script sekali dan setujui. Zona waktu mengikuti Setelan Proyek Apps Script (pilih Asia/Jakarta). Daftar cadangan di halaman yang sama punya tombol **Pulihkan**. |
| Hasil Belajarku cepat | Otomatis. Rekap per murid dibuat saat guru menyimpan. Setelah memulihkan backup lama atau migrasi, klik Portal Admin > Sistem > **Bangun ulang rekap**. |
| Antrean ulangan | Tombol Mulai siswa menunggu acak 1-3 detik. |
| Nilai ulangan ke jurnal | Otomatis di Jurnal Guru (menu **Tarik Nilai Ulangan** untuk segera). Masuk sebagai nilai kategori "Ulangan" pada jurnal kelas yang sama (mapel cocok, tanggal jurnal pada/sebelum tanggal ulangan). Siswa harus punya NISN di Data Siswa guru. Nilai yang sudah ada tidak ditimpa. |
| Ekspor hasil + analisis soal | Tab Hasil di Ulangan Online: PDF dan Excel mengikuti filter kelas. Memuat statistik, per kelas, sebaran nilai, dan analisis butir (kesulitan, daya pembeda, sebaran A-D, pengecoh, KR-20, rekomendasi). |
| Satu perangkat per sesi | Otomatis. Perangkat kedua ditolak selama perangkat pertama aktif; bila perangkat pertama diam >2 menit, perangkat lain boleh mengambil alih. Guru dapat menekan **Lepas perangkat** di tab Pantau. |
| Peringatan pindah tab | Layar dijeda, judul tab berkedip, ada hitungan sisa kesempatan dan peringatan terakhir sebelum jawaban dikumpulkan otomatis. |
| Kenaikan kelas | Portal Admin > Kenaikan Kelas: atur tujuan tiap kelas, jalankan sekaligus (file Akun dicadangkan otomatis dulu). Tersedia pembuatan periode baru dan sinkronisasi akun dari Data Siswa. |
| Video panduan | `panduan.html` (folder `video/`), tertaut di portal, Jurnal Guru, Portal Murid, dan Portal Admin. |
| Kode ulangan wajib | Walau siswa memilih ulangan dari daftar (mapel + tingkat), kolom **Kode ulangan** tetap wajib diisi. Daftar pilihan tidak lagi memuat kode (hanya pengenal acak), dan server menolak kode yang bukan milik ulangan yang dipilih (maks. 8 salah per 5 menit per akun). Siswa yang masuk lewat tautan/kode langsung tidak terpengaruh. |
| Menu guru lebih cepat | Server hanya membaca baris milik ulangan yang dibuka (bukan seluruh sheet), bank soal besar kini ikut ter-cache, dan daftar ulangan guru di-cache 15 detik. Di layar: daftar tampil seketika dari data terakhir lalu disegarkan diam-diam, data ulangan diambil saat kursor menyentuh tombol Lihat, dan pantauan diminta bersamaan dengan data soal. |
