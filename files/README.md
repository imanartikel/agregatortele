# Bot Order Pengiriman → Excel Agregator

## Setup

### 1. Install dependencies
```bash
pip install -r requirements.txt
```

### 2. Isi config di bot.py

Buka `bot.py`, edit bagian ini:

```python
TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "ISI_TOKEN_LU_DISINI")
```
Ganti dengan token dari @BotFather.

Edit data toko pengirim:
```python
PENGIRIM = {
    "nama":       "Nama Toko Lu",
    "alamat":     "Alamat Toko Lengkap",
    "telp":       "08xxxxxxxxxx",
    "kecamatan":  "Kecamatan Toko",
    "kota":       "Kota Toko",
}
```

Default kurir (kalau customer tidak menyebutkan):
```python
DEFAULT_KURIR = "JNE"      # atau "Lion Parcel"
DEFAULT_LAYANAN = "REG"
```

### 3. Taruh template Excel
Pastikan file `template.xlsx` ada di folder yang sama dengan `bot.py`.

### 4. Jalankan bot
```bash
python bot.py
```

Atau pakai environment variable biar token ga hardcode:
```bash
export TELEGRAM_BOT_TOKEN="token_lu_disini"
python bot.py
```

---

## Cara Pakai

### Format pesan yang diterima bot:
```
Nama : Janan Shofiyah A
Alamat lengkap : Bogor raya permai fe 2 no 10
Kec : bogor barat
Kota/kab : kota bogor
Kode Pos : 16113
No hp : 08111199109
Pesanan : sprei 160 x 200 x 35 motif bunga sesuai gambar
```

### Opsional (bisa ditambah):
```
Kurir : Lion Parcel     ← kalau mau ganti dari default
Layanan : REG
Berat : 1.5
Harga Barang : 150000
Instruksi : Jangan dilipat
No Referensi : ORD-001
```

### Deteksi otomatis kurir:
- Kalau di teks ada kata "lion" atau "lion parcel" → otomatis Lion Parcel
- Kalau ada "jne" → otomatis JNE
- Kalau tidak ada → pakai DEFAULT_KURIR

### Perintah bot:
- `/start` — lihat panduan format
- `/list` — lihat semua order pending (belum diexport)
- `/export` — download file Excel siap upload agregator
- `/clear` — hapus semua order pending

---

## Flow Kerja

1. Customer kirim pesan order ke bot
2. Bot auto-parse & simpan di memory
3. Kalau mau export: ketik `/export` atau klik tombol
4. Bot kirim file `.xlsx` yang sudah terisi
5. Upload langsung ke agregator pengiriman

---

## Catatan

- Data order disimpan di **memory** (hilang kalau bot di-restart)
- Kalau mau persistent, perlu tambah SQLite — bisa dikembangin nanti
- Satu chat = satu "bucket" order yang terpisah
- File output disimpan sementara di folder `output/` lalu dihapus setelah dikirim

---

## Deploy di VPS (opsional)

```bash
# Buat systemd service biar auto-start
sudo nano /etc/systemd/system/order-bot.service
```

```ini
[Unit]
Description=Telegram Order Bot
After=network.target

[Service]
WorkingDirectory=/path/ke/folder/bot
ExecStart=/usr/bin/python3 bot.py
Environment=TELEGRAM_BOT_TOKEN=token_lu
Restart=always
User=ubuntu

[Install]
WantedBy=multi-user.target
```

```bash
sudo systemctl enable order-bot
sudo systemctl start order-bot
```
