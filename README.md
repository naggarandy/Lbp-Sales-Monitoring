# 📊 LBP Sales Monitor

Aplikasi web monitoring penjualan — **upload file XLSX, langsung dapat dashboard lengkap**.  
Bisa dibuka dari **HP** (browser Chrome/Safari).

## Fitur

- 📤 Upload file Excel format LBP (atau sejenis)
- 📦 Monitoring **Produk** — Qty, Omset Bruto, Net Sales, % kontribusi
- 🏪 Monitoring **Outlet** — ranking + detail produk per outlet
- 👤 Monitoring **Salesman** — ranking + detail produk
- 📍 Monitoring **Wilayah** (Kabupaten, Kecamatan) & Channel
- 📅 Tren **harian & mingguan**
- 🔄 Analisis **Return / Credit Note**
- 🔎 Filter interaktif (Kabupaten, Salesman, Subbrand, Channel, Tanggal)
- ⬇️ Export hasil ke Excel
- 📱 Tampilan ramah HP

## Cara Jalankan di Komputer

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Jalankan aplikasi
streamlit run app.py
```

Browser akan terbuka di `http://localhost:8501`

## Cara Akses dari HP

### Opsi A — Jaringan WiFi yang sama
```bash
streamlit run app.py --server.address 0.0.0.0 --server.port 8501
```
Lalu buka di HP: `http://IP-KOMPUTER:8501`  
(Cek IP dengan `ipconfig` / `ifconfig`)

### Opsi B — Deploy gratis ke cloud (paling mudah untuk HP)
1. Buat akun di [streamlit.io/cloud](https://streamlit.io/cloud)
2. Upload folder ini ke GitHub
3. Connect repo → Deploy
4. Dapat link publik, buka dari HP kapan saja

### Opsi C — LocalTunnel / Ngrok (sementara)
```bash
# Terminal 1
streamlit run app.py --server.port 8501

# Terminal 2
npx localtunnel --port 8501
# atau: ngrok http 8501
```
Dapat link publik untuk dibuka di HP.

## Format File yang Didukung

File XLSX dengan kolom (minimal):
- No Outlet, Nama Outlet
- Nama Produk, SUBBRANDNAME / Pcode
- QTYPCS, Harga Bruto, Total
- Salesman, Kabupaten, Channel
- Tanggal Faktur

Baris header biasanya di baris ke-2 (format LBP standar).

## Struktur Folder

```
lbp_monitor_app/
├── app.py              # Aplikasi utama
├── requirements.txt    # Dependencies
└── README.md           # Panduan ini
```
