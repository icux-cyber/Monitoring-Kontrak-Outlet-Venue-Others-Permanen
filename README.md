# Arsitektur Sistem Monitoring OOH

Aplikasi ini adalah sistem berbasis Streamlit untuk mengelola aset media **Out of Home (OOH)**, memantau kontrak, dan mencatat pergantian brand atau visual.

## Diagram arsitektur

```text
┌──────────────┐
│   Pengguna   │
└──────┬───────┘
       │ menggunakan aplikasi
       ▼
┌──────────────────────────────────────────┐
│ app.py — Antarmuka Streamlit             │
│                                          │
│ • Dashboard, grafik, peta, dan monitoring│
│ • Form master data OOH                   │
│ • Pergantian brand/visual dan riwayat    │
│ • Laporan kontrak dan ekspor Excel       │
│ • Kontrol backup/load Google Sheets      │
└───────────┬──────────────────┬───────────┘
            │                  │
            ▼                  ▼
┌─────────────────────┐  ┌────────────────────────┐
│ data_manager.py     │  │ Google Apps Script     │
│                     │  │ melalui HTTP POST      │
│ • Baca/tulis Excel  │  │ Backup dan pemulihan   │
│ • Hitung kontrak    │  │ data Google Sheets     │
│ • Kelola riwayat    │  └────────────────────────┘
│ • Rollback perubahan│
└──────────┬──────────┘
           ▼
┌──────────────────────────┐   ┌───────────────────┐
│ data/*.xlsx              │   │ File GeoJSON      │
│ Penyimpanan lokal        │   │ Batas wilayah peta│
│ • Master OOH             │   └───────────────────┘
│ • Riwayat pergantian     │
│ • Master vendor/visual/  │
│   pejabat                │
└──────────────────────────┘
```

## Komponen

- **`app.py`** — Lapisan antarmuka dan alur aplikasi. Menyajikan dashboard, pengelolaan OOH, pergantian brand/visual, laporan kontrak, monitoring kontrak outlet, pengaturan master, monitoring event, dan daftar materi promosi.
- **`data_manager.py`** — Lapisan pengelolaan data. Membaca dan menyimpan berkas Excel, menghitung durasi serta status kontrak, memvalidasi workbook kontrak outlet, mencatat pergantian, dan memuat master materi promosi.
- **`data/*.xlsx`** — Penyimpanan lokal untuk master titik OOH, 707 data kontrak outlet awal, riwayat pergantian, vendor/visual/pejabat, serta daftar materi promosi yang difilter.
- **Google Apps Script / Google Sheets** — Integrasi eksternal untuk membackup dan memulihkan data melalui permintaan HTTP POST dari aplikasi.
- **Berkas GeoJSON** — Sumber batas kecamatan yang dipakai untuk visualisasi peta.

## Alur data

1. Pengguna berinteraksi dengan fitur melalui antarmuka Streamlit di `app.py`.
2. Untuk membaca atau mengubah data lokal, aplikasi memanggil fungsi di `data_manager.py`.
3. Data master, riwayat, dan kontrak outlet disimpan terpisah dalam berkas Excel di direktori `data/`; status kontrak dihitung saat data dimuat.
4. Menu Monitoring Kontrak Outlet menerima template aplikasi atau sheet `COSPKR06`, serta menyediakan edit tabel, unduh template, dan ekspor seluruh data.
5. Daftar materi promosi hanya memuat material Vinyl Billboard, Vinyl Baliho, Vinyl PJU, dan Vinyl Midis; file Excel dapat diunggah untuk mengganti daftar, filter stok tersedia, dan pilihan brand/versi pada penggantian visual disaring sesuai jenis serta ukuran titik OOH.
6. Dashboard menggunakan data tersebut untuk menyajikan ringkasan, grafik, tabel, dan peta. Peta memerlukan berkas GeoJSON.
7. Pengguna dapat memicu backup ke atau pemulihan dari Google Sheets melalui Google Apps Script.

## Teknologi utama

- Python dan Streamlit
- Pandas dan berkas Excel (`openpyxl`)
- Plotly dan Matplotlib untuk visualisasi
- GeoJSON untuk data batas peta
- `requests` untuk integrasi Google Apps Script

## Backup kontrak outlet ke Google Sheets

Menu **Backup ke Google Sheet** mengirim seluruh data kontrak outlet ke tab `Monitoring Kontrak Outlet`. Backup mengganti isi tab tersebut, tetapi tidak mengubah tab lain di spreadsheet.

1. Buka spreadsheet tujuan, pilih **Extensions > Apps Script**, lalu salin isi `google-apps-script/Code.gs` ke project Apps Script yang terikat ke spreadsheet tersebut.
2. Deploy sebagai **Web app**, pilih **Execute as: Me** dan izinkan akses **Anyone**.
3. Simpan URL deployment di Streamlit secrets:

   ```toml
   GOOGLE_APPS_SCRIPT_URL = "URL_WEB_APP"
   ```

   Untuk Streamlit Community Cloud, masukkan nilai tersebut pada bagian **Settings > Secrets**. Untuk lokal, gunakan `.streamlit/secrets.toml`; jangan commit file secrets.
4. Buka menu **Backup ke Google Sheet** dan pilih **Backup data sekarang**.

Web App ini tidak menggunakan token. Siapa pun yang mengetahui URL deployment dapat menjalankan backup dan menimpa isi tab `Monitoring Kontrak Outlet`; simpan URL di secrets dan jangan membagikannya.

Bagian **Upload / Download Excel** tersedia pada halaman Dashboard. Impor dari sana mengganti seluruh daftar customer.
