# 📊 SCJ Dashboard — Panduan Update Data

## Cara Update Data Excel

1. **Siapkan file Excel** terbaru dengan format:
   `Bookings - By Carrier - 2026.xlsx`

2. **Letakkan file** ke folder `data/`:
   ```
   scj-dashboard-flow/
   └── data/
       └── Bookings - By Carrier - 2026.xlsx   ← taruh di sini
   ```

3. **Jalankan server**:
   ```
   npm start
   ```

4. **Buka browser**: http://localhost:2099

5. **Klik tombol "Refresh Data"** di dashboard untuk memuat ulang tanpa restart server.

---

## Auto-Refresh

Server otomatis mendeteksi perubahan file Excel setiap 3 detik. Cukup replace file di `data/` dan klik Refresh.

## Struktur Folder

```
scj-dashboard-flow/
├── data/
│   └── Bookings - By Carrier - 2026.xlsx   ← Excel data source
├── public/
│   ├── index.html                           ← Dashboard frontend
│   └── SCJ Logo.jpg                         ← Logo perusahaan
├── server.js                                ← Express backend (port 2099)
├── package.json
├── start.bat                                ← Double-click to run
└── UPDATE_DATA.md                           ← File ini
```

## Port

Server berjalan di **port 2099**. Ubah di `server.js` bila perlu.
