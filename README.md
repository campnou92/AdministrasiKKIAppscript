# Sistem Administrasi KKI INSTIPER Yogyakarta

Sistem Informasi Administrasi dan Manajemen Dokumen Kerjasama Internasional & Nasional Kantor Urusan Internasional (KKI) Institut Pertanian STIPER (INSTIPER) Yogyakarta berbasis Google Apps Script dan Web App responsif.

## 📌 Fitur Utama

1. **Dashboard & Statistik Terpadu (Beranda):**
   - Ringkasan surat masuk belum diproses dan surat keluar menunggu tanda tangan.
   - Status tenggat waktu tugas & tindak lanjut (peringatan terlambat dan mendesak).
   - Rekap dokumen kerjasama aktif dan naskah yang akan kedaluwarsa.
   - Akses cepat (*quick actions*) untuk penambahan data langsung.

2. **Pengelolaan Arsip Surat (Surat Masuk & Keluar):**
   - Pencarian real-time dan penyaringan multi-kriteria (Tahun arsip, status proses, status tanda tangan).
   - Pratinjau (*preview*) dokumen scan surat masuk dan berkas PDF surat keluar langsung di aplikasi.
   - Pembuatan surat keluar otomatis dari template Google Docs dengan penomoran dinamis.
   - Unggah hasil scan/PDF tanda tangan digital & integrasi Google Drive.

3. **Pusat Tugas & Pengingat (Reminder Hub):**
   - Manajemen tugas dan disposisi tindak lanjut surat maupun naskah kerjasama.
   - Fitur tambah tugas manual langsung dari web app.
   - Deteksi otomatis tugas yang telah melewati tenggat (*overdue*) atau mendesak (≤ 3 hari).
   - Filter tugas berdasarkan status (*Pending*, *On Progress*, *Done*, *Overdue*, *Urgent*) dan sumber rujukan.

4. **Portal Naskah Kerjasama (MoU, MoA, LoI, IA):**
   - Pencarian dan filter dokumen berdasarkan jenis naskah, wilayah (Internasional / Nasional), status (Aktif, Akan Berakhir ≤ 6 Bulan, Expired, Draf Review), dan tahun mulai.
   - Unggah naskah dan pratinjau dokumen langsung di dalam aplikasi.
   - Alur verifikasi draf pengajuan usulan naskah kerjasama.

5. **Notifikasi Otomatis Multi-Kanal (Email & Telegram Bot):**
   - Pengiriman rekap laporan harian dan pengingat tenggat waktu secara otomatis setiap hari kerja.
   - Integrasi bot Telegram (`@KKIInstiperbot`) dengan deteksi otomatis Chat ID dan fitur uji kirim.
   - Pusat pengaturan jadwal jam pengiriman (*trigger*) dan penentuan hari peringatan H- sebelum tenggat.

6. **Manajemen Pengguna & Hak Akses Berjenjang:**
   - **Admin:** Hak penuh (kelola surat, tugas, konfigurasi penomoran surat, integrasi Telegram/Email, verifikasi MoU, dan manajemen akun pengguna).
   - **Pimpinan:** Hak akses pemantauan arsip surat, disposisi tugas, naskah kerjasama, dan pengajuan usulan.
   - **Kontributor:** Akses katalog kerjasama aktif dan pengajuan naskah kerjasama baru.

---

## 🛠️ Arsitektur & Teknologi

- **Backend / Engine:** Google Apps Script (GAS) — JavaScript ES6+.
- **Database / Penyimpanan:** Google Sheets & Google Drive API.
- **Frontend:** Single Page Application (SPA) responsif HTML5, CSS3 kustom modern, dan Vanilla JavaScript.
- **Integrasi Eksternal:** Telegram Bot API (`sendMessage`, `getUpdates`) & Google Mail (`GmailApp`).
- **Deployment Management:** `@google/clasp`.

---

## 🚀 Penggunaan & Deployment

Proyek ini dikembangkan dan disinkronkan ke Google Apps Script menggunakan `@google/clasp`:

```bash
# Push perubahan ke Google Apps Script
clasp push

# Deploy pembaruan versi web app
clasp deploy -i <DEPLOYMENT_ID> -d "Deskripsi pembaruan"
```

---

&copy; Kantor Urusan Internasional (KKI) — Institut Pertanian STIPER (INSTIPER) Yogyakarta.
