import os
import re
import logging
import json
import uuid
from datetime import datetime
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, ContextTypes, filters, ApplicationBuilder
)
import shutil
import xlwings as xw
import pythoncom

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)

# ─── CONFIG ───────────────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(BASE_DIR)

# Load .env manually if exists
env_path = os.path.join(PROJECT_ROOT, ".env")
if os.path.exists(env_path):
    with open(env_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                os.environ[key.strip()] = val.strip().strip('"').strip("'")

TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
if not TOKEN:
    raise ValueError("ERROR: TELEGRAM_BOT_TOKEN is not set in environment or .env file!")
TEMPLATE_PATH = os.path.join(BASE_DIR, "template.xlsx")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
DB_PATH = os.path.join(BASE_DIR, "pending_orders.json")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Data pengirim tetap — isi sesuai toko lu
PENGIRIM = {
    "nama":       "Ruma Asri",
    "alamat":     "Perumahan tamansari bukit Damai blok c4 no. 7",
    "telp":       "089636270890",
    "kecamatan":  "Gunung Sindur",
    "kota":       "Bogor",
}

# Default kurir & layanan (bisa dioverride per order)
DEFAULT_KURIR = "ID Express"
DEFAULT_LAYANAN = "Reguler"

# Simpan pending orders per chat (order yang belum diexport)
pending_orders: dict[int, list[dict]] = {}

def save_db():
    try:
        # Convert datetime objects ke strings agar bisa di-JSON
        serializable_data = {}
        for chat_id, orders in pending_orders.items():
            serializable_data[str(chat_id)] = []
            for o in orders:
                o_copy = o.copy()
                if 'timestamp' in o_copy and isinstance(o_copy['timestamp'], datetime):
                    o_copy['timestamp'] = o_copy['timestamp'].isoformat()
                serializable_data[str(chat_id)].append(o_copy)
                
        with open(DB_PATH, "w") as f:
            json.dump(serializable_data, f)
    except Exception as e:
        logging.error(f"Gagal save DB: {e}")

def load_db():
    global pending_orders
    if os.path.exists(DB_PATH):
        try:
            with open(DB_PATH, "r") as f:
                data = json.load(f)
                for chat_id_str, orders in data.items():
                    processed_orders = []
                    for o in orders:
                        if 'timestamp' in o:
                            o['timestamp'] = datetime.fromisoformat(o['timestamp'])
                        processed_orders.append(o)
                    pending_orders[int(chat_id_str)] = processed_orders
            logging.info(f"Database dimuat: {len(pending_orders)} chat(s) ditemukan.")
        except Exception as e:
            logging.error(f"Gagal load DB: {e}")

# ─── PARSER ───────────────────────────────────────────────────────────────────
def parse_order(text: str) -> dict | None:
    text_lower = text.lower()
    result = {}
    
    match_nama = re.search(r"nama\s*:\s*(.+)", text, re.IGNORECASE)
    if match_nama:
        result['nama'] = match_nama.group(1).strip()
    else:
        # Nama di baris pertama tanpa label (format: "Nisa\nAlamat : ...")
        first_line = next((line.strip() for line in text.splitlines() if line.strip()), None)
        if first_line and not re.match(
            r"(alamat|kec|kota|kab|kode|no\.?\s*hp|deskripsi|pesanan|berat|harga|kurir|jne|lion|sicepat)",
            first_line, re.IGNORECASE
        ):
            result['nama'] = first_line
    
    match_alamat = re.search(r"alamat\s*:\s*(.*?)(?=kel(?:urahan)?[\s.:]|kec(?:amatan)?[\s.:]|kota[\s.:]|kab(?:upaten)?[\s.:]|kode\s*pos|no\.?\s*hp|deskripsi|pesanan|jne|lion|berat|$)", text, re.IGNORECASE | re.DOTALL)
    if match_alamat: result['alamat'] = match_alamat.group(1).strip()

    # Kelurahan — opsional, digabung ke alamat saat export
    match_kel = re.search(r"kel(?:urahan)?[\s.:]+(\w[\w\s]*?)(?=\s+kec|\s+kab|\s+kota|,|\n|$)", text, re.IGNORECASE)
    if match_kel: result['kelurahan'] = match_kel.group(1).strip()

    # Support: "kec : Cerme", "kecamatan : Cerme", "kec.Cerme", "kec. Cerme"
    match_kec = re.search(r"kec(?:amatan)?[\s.:]+(\w[\w\s]*?)(?=\s+kab|\s+kota|,|\n|$)", text, re.IGNORECASE)
    if match_kec: result['kecamatan'] = match_kec.group(1).strip()

    # Support: "kab : Gresik", "kabupaten : Gresik", "kab.Gresik", "kota : Gresik"
    match_kota = re.search(r"(?:kota|kab(?:upaten)?)[\s.:]+(\w[\w\s]*?)(?=\s+kec|\s+kode|,|\n|$)", text, re.IGNORECASE)
    if match_kota: result['kota'] = match_kota.group(1).strip()
    
    match_kodepos = re.search(r"kode\s*pos\s*[:]?\s*(\d+)", text, re.IGNORECASE)
    if match_kodepos: result['kode_pos'] = match_kodepos.group(1).strip()

    # No HP: Support label "No HP" atau deteksi otomatis angka 08xx / 62xx
    # Penting: jangan pakai \s di dalam character class karena akan makan newline
    match_hp = re.search(r"no\.?\s*h[p]\s*[:]?\s*([0-9\-\+]+)", text, re.IGNORECASE)
    if not match_hp:
        match_hp = re.search(r"(08[0-9]{8,12}|62[0-9]{9,13})", text)
    if match_hp: result['hp'] = match_hp.group(1).strip()
    
    # Berat: Support "3kg" atau "berat 3kg" atau "berat : 3"
    match_berat = re.search(r"(?:berat\s*[:]?\s*)?([0-9.,]+)\s*kg", text_lower)
    if not match_berat:
        match_berat = re.search(r"berat\s*[:]?\s*([0-9.,]+)", text_lower)
    if match_berat: result['berat'] = match_berat.group(1).strip()

    # Harga — dengan atau tanpa label
    match_harga = re.search(r"(?:harga|nilai)\s*(?:barang)?\s*[:]?\s*([0-9.,]+)", text_lower)
    if match_harga:
        result['nilai_barang'] = match_harga.group(1).replace('.', '').replace(',', '').strip()
    else:
        # Fallback: angka format ribuan Indonesia berdiri sendiri di satu baris (contoh: 1.200.000)
        match_harga_bare = re.search(r"^([0-9]{1,3}(?:\.[0-9]{3})+)$", text, re.MULTILINE)
        if match_harga_bare:
            result['nilai_barang'] = match_harga_bare.group(1).replace('.', '')

    # Pesanan / Deskripsi
    # pesanan\w* → nangkep typo seperti "pesananan", "pesannan", dll
    # \n\n = baris kosong sebagai pemisah antara deskripsi dan berat/harga
    match_pesanan = re.search(r"(?:deskripsi|pesanan\w*)\s*:\s*(.*?)(?=\n\n|\d+\s*kg|jne|lion|sicepat|j&t|berat|asuransi|harga|nilai|$)", text, re.IGNORECASE | re.DOTALL)
    if match_pesanan:
        result['pesanan'] = match_pesanan.group(1).strip()
    else:
        # Fallback: Ambil teks setelah No HP tapi sebelum kurir/berat/harga
        if match_hp:
            start_idx = match_hp.end()
            remainder = text[start_idx:].strip()
            remainder = re.sub(r"^(?::|\s)+", "", remainder).strip()

            next_field = re.search(r"(jne|lion|sicepat|j&t|berat|asuransi|harga|nilai|kg)", remainder, re.IGNORECASE)
            if next_field:
                candidate = remainder[:next_field.start()].strip()
            else:
                candidate = remainder
            # Abaikan kalau yang tersisa hanya angka (berat/harga tanpa label)
            if candidate and not re.fullmatch(r"[0-9\s.,]+", candidate):
                result['pesanan'] = candidate

    if "jne" in text_lower:
        result["kurir"] = "JNE"
        result["layanan"] = "Reguler"
    elif "lion" in text_lower or "lion parcel" in text_lower:
        result["kurir"] = "Lion Parcel"
        result["layanan"] = "Reguler"
    elif "sicepat" in text_lower:
        result["kurir"] = "SiCepat"
        result["layanan"] = "Reguler"
    elif "id express" in text_lower or "idx" in text_lower:
        result["kurir"] = "ID Express"
        result["layanan"] = "Reguler"
    else:
        result["kurir"] = DEFAULT_KURIR
        result["layanan"] = DEFAULT_LAYANAN
        
    if "tanpa asuransi" in text_lower:
        result["asuransi"] = "Tidak"
    else:
        result["asuransi"] = "Ya"
        
    required = ["nama", "alamat", "hp", "kecamatan", "kota"]
    missing = [f for f in required if f not in result or not result[f]]
    result['_missing'] = missing
    return result


# ─── EXCEL EXPORT ─────────────────────────────────────────────────────────────
def export_to_excel(orders: list[dict], chat_id: int) -> str:
    # Inisialisasi COM untuk thread ini (wajib buat xlwings di background thread)
    pythoncom.CoInitialize()
    
    try:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = os.path.join(OUTPUT_DIR, f"orders_{chat_id}_{timestamp}.xlsx")
        
        # Copy template asli
        shutil.copy(TEMPLATE_PATH, out_path)
        
        logging.info(f"Mulai loading template pake Excel Asli: {out_path}")
        
        # Jalankan Excel di background (hidden)
        app = xw.App(visible=False)
        try:
            wb = app.books.open(out_path)
            ws = wb.sheets["Upload"]
            
            start_row = 2
            
            # Kumpulin semua data dulu biar nulisnya sekali jalan (jauh lebih cepet)
            excel_data = []
            for i, order in enumerate(orders):
                row_data = [
                    i + 1,                                          # 1: No
                    "PICKUP",                                       # 2: Jenis Pengiriman
                    order.get("kurir", DEFAULT_KURIR),             # 3: Pilihan Kurir
                    order.get("layanan", DEFAULT_LAYANAN),          # 4: Jenis Layanan
                    PENGIRIM["nama"],                               # 5: Nama Pengirim
                    PENGIRIM["alamat"],                             # 6: Alamat Pengirim
                    PENGIRIM["telp"],                               # 7: No Telp Pengirim
                    PENGIRIM["kecamatan"],                          # 8: Kec Pengirim
                    PENGIRIM["kota"],                               # 9: Kota Pengirim
                    order.get("nama", ""),                          # 10: Nama Penerima
                    order.get("alamat", ""),                        # 11: Alamat Penerima
                    order.get("hp", ""),                            # 12: No Telp Penerima
                    order.get("kecamatan", ""),                     # 13: Kec Penerima
                    order.get("kota", ""),                          # 14: Kota Penerima
                    order.get("berat", ""),                         # 15: Berat (kg)
                    "", "", "",                                     # 16,17,18: P, L, T (kosong)
                    order.get("pesanan", ""),                       # 19: Isi Paket
                    order.get("asuransi", ""),                      # 20: Asuransi
                    order.get("nilai_barang", ""),                  # 21: Harga Barang
                    "",                                             # 22: Nilai COD (kosong)
                    order.get("instruksi", ""),                     # 23: Instruksi
                    order.get("no_ref", "")                         # 24: No Referensi
                ]
                excel_data.append(row_data)
            
            # Tulis SEMUA data sekaligus (ini kuncinya biar ngebut)
            if excel_data:
                ws.range(f"A{start_row}").value = excel_data
            
            # Tambahin Zebra Coloring (warna selang-seling)
            for i in range(len(orders)):
                if i % 2 == 1:
                    row_num = start_row + i
                    ws.range(f"A{row_num}:X{row_num}").color = (240, 244, 255)
            
            wb.save()
            wb.close()
        finally:
            app.quit() # Pastikan Excel-nya mati
            
        logging.info(f"File berhasil disimpan pake Excel Asli: {out_path}")
        return out_path
    except Exception as e:
        logging.error(f"Gagal ekspor pake xlwings: {e}")
        raise e
    finally:
        pythoncom.CoUninitialize()


# ─── HANDLERS ─────────────────────────────────────────────────────────────────
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "🛍️ *Bot Order Pengiriman*\n\n"
        "━━━━━━━━━━━━━━━\n"
        "*✅ FORMAT PALING AMAN:*\n"
        "```\n"
        "Nama : Ririn Rosdiana\n"
        "Alamat : Jl Raya Banjarsari No.117 RT 04 RW 03\n"
        "Kec : Cerme\n"
        "Kab : Gresik\n"
        "Kode Pos : 61171\n"
        "No. HP : 081235253322\n"
        "\n"
        "Deskripsi :\n"
        "Sprei King 200x200\n"
        "\n"
        "Berat 3 kg\n"
        "Harga 250000\n"
        "```\n\n"
        "━━━━━━━━━━━━━━━\n"
        "*✅ DO (boleh):*\n"
        "• Kec/Kab/Kel boleh pakai titik dua, titik, atau spasi sebagai pemisah\n"
        "  `Kec : Cerme` atau `Kec.Cerme` atau `Kecamatan Cerme` → semua jalan\n"
        "• Nama boleh langsung di baris pertama tanpa label\n"
        "• Deskripsi/berat/harga boleh tidak diisi\n"
        "• Tulis `JNE` → pakai JNE\n"
        "• Tulis `Lion` → pakai Lion Parcel\n"
        "• Tulis `SiCepat` → pakai SiCepat\n"
        "• Tulis `ID Express` atau `IDX` → pakai ID Express\n"
        "• Default (tidak tulis apapun) → *ID Express*\n"
        "• Tulis `tanpa asuransi` → asuransi diisi Tidak\n"
        "• Harga boleh tanpa label: `1.200.000` di baris sendiri\n"
        "• Berat: `2 kg` atau `2kg` atau `Berat 2 kg`\n\n"
        "━━━━━━━━━━━━━━━\n"
        "*❌ JANGAN:*\n"
        "• `Nama`, `Alamat`, `No. HP` *WAJIB pakai titik dua* ( : )\n"
        "• Jangan tulis No. HP sambung langsung ke baris berikutnya tanpa enter\n"
        "• Jangan taruh berat/harga *sebelum* No. HP\n"
        "• Jangan gabung kec & kab di baris Alamat (taruh di baris sendiri)\n"
        "• Jangan kirim pesan selain format order di chat ini\n\n"
        "━━━━━━━━━━━━━━━\n"
        "*⚠️ Yang WAJIB ada:* Nama : · Alamat : · No. HP : · Kec · Kab\n"
        "*Yang opsional:* Kelurahan, Kode Pos, Deskripsi, Berat, Harga\n\n"
        "━━━━━━━━━━━━━━━\n"
        "*Perintah:*\n"
        "/list — lihat order pending hari ini\n"
        "/export — download Excel\n"
        "/clear — hapus semua pending\n"
        "/start — tampilkan panduan ini",
        parse_mode="Markdown"
    )

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        chat_id = update.effective_chat.id
        text = update.message.text
        
        order = parse_order(text)
        missing = order.get('_missing', []) if order else []

        label_map = {
            "nama": "Nama", "alamat": "Alamat", "hp": "No. HP",
            "kecamatan": "Kecamatan (Kec)", "kota": "Kabupaten/Kota (Kab/Kota)"
        }

        if missing:
            # Cek apakah ini percobaan order (ada minimal 1 label field atau nomor HP)
            if re.search(r"(nama|alamat|no\.?\s*hp|deskripsi|pesanan)\s*:", text, re.IGNORECASE) or \
               re.search(r"(08[0-9]{8,12}|62[0-9]{9,13})", text):
                missing_labels = "\n".join(f"• {label_map.get(f, f)}" for f in missing)
                await update.message.reply_text(
                    f"⚠️ *Order gagal disimpan!*\n\n"
                    f"Field berikut tidak ditemukan:\n{missing_labels}\n\n"
                    f"Kirim /start untuk lihat format.",
                    parse_mode="Markdown"
                )
            return

        # Bersihkan _missing sebelum disimpan
        order.pop('_missing', None)

        # Kelurahan digabung ke alamat
        if order.get('kelurahan'):
            order['alamat'] = f"{order['alamat']}, Kel. {order['kelurahan']}"

        # Siapkan preview untuk konfirmasi
        current_orders = pending_orders.get(chat_id, [])
        count = len(current_orders) + 1
        kurir_info = f"{order['kurir']} {order.get('layanan', '')}"
        pesanan_preview = order.get('pesanan', '-')

        keyboard = [
            [InlineKeyboardButton("📥 Export Sekarang", callback_data="export"),
             InlineKeyboardButton("📋 Lihat List", callback_data="list")]
        ]
        reply_markup = InlineKeyboardMarkup(keyboard)

        # Escape markdown characters in preview to avoid parsing errors
        safe_pesanan = pesanan_preview.replace('*', 'x')
        safe_nama = order['nama'].replace('*', '')

        await update.message.reply_text(
            f"✅ *Order #{count} tersimpan!*\n\n"
            f"👤 {safe_nama}\n"
            f"📍 {order.get('kecamatan', '-')}, {order.get('kota', '-')}\n"
            f"📦 {safe_pesanan[:60]}{'...' if len(safe_pesanan) > 60 else ''}\n"
            f"🚚 {kurir_info}\n\n"
            f"Total pending: *{count} order*",
            parse_mode="Markdown",
            reply_markup=reply_markup
        )

        # SIMPAN KE DATABASE HANYA JIKA PESAN BERHASIL TERKIRIM
        order['id'] = str(uuid.uuid4())[:8] # Generate ID pendek
        order['timestamp'] = datetime.now()
        if chat_id not in pending_orders:
            pending_orders[chat_id] = []
        pending_orders[chat_id].append(order)
        save_db()
    except Exception as e:
        logging.error(f"Error handling message: {e}", exc_info=True)
        # Jangan diem aja kalau error
        await update.message.reply_text("❌ Terjadi kesalahan internal saat memproses order. Mohon cek format atau hubungi admin.")

async def list_orders(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    today = datetime.now().date()
    # Ambil order hari ini
    orders = [o for o in pending_orders.get(chat_id, []) if o.get('timestamp', datetime.now()).date() == today]
    
    if not orders:
        msg_text = "📭 Tidak ada order pending untuk hari ini."
        if update.callback_query:
            await update.callback_query.answer()
            await update.callback_query.message.reply_text(msg_text)
        else:
            await update.message.reply_text(msg_text)
        return
    
    lines = [f"📋 *{len(orders)} Order Pending (HARI INI):*\n"]
    keyboard = []
    row = []
    
    for i, o in enumerate(orders, 1):
        lines.append(
            f"{i}. *{o['nama']}* — {o.get('kota', '-')}\n"
            f"   {o['pesanan'][:50]}{'...' if len(o['pesanan']) > 50 else ''}"
        )
        # Tombol hapus per nomor
        order_id = o.get('id', str(uuid.uuid4())[:8])
        row.append(InlineKeyboardButton(f"❌ Hapus #{i}", callback_data=f"del_{order_id}"))
        
        if len(row) == 2: # 2 tombol per baris biar rapi
            keyboard.append(row)
            row = []
            
    if row:
        keyboard.append(row)
        
    # Tambah tombol export di paling bawah
    keyboard.append([InlineKeyboardButton("📥 Export ke Excel", callback_data="export")])
    
    final_text = "\n".join(lines)
    
    if update.callback_query:
        await update.callback_query.answer()
        await update.callback_query.message.reply_text(final_text, parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(keyboard))
    else:
        await update.message.reply_text(final_text, parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(keyboard))

async def delete_order_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    chat_id = update.effective_chat.id
    order_id = query.data.replace("del_", "")
    
    if chat_id in pending_orders:
        original_count = len(pending_orders[chat_id])
        pending_orders[chat_id] = [o for o in pending_orders[chat_id] if o.get('id') != order_id]
        
        if len(pending_orders[chat_id]) < original_count:
            save_db()
            await query.answer("✅ Order berhasil dihapus!")
            # Update pesan list (panggil list_orders lagi)
            await query.message.delete() # Hapus pesan lama biar nggak numpuk
            await list_orders(update, context)
        else:
            await query.answer("⚠️ Order tidak ditemukan atau sudah dihapus.", show_alert=True)
    else:
        await query.answer("❌ Data tidak ditemukan.")

async def export_orders(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    today = datetime.now().date()
    orders = [o for o in pending_orders.get(chat_id, []) if o.get('timestamp', datetime.now()).date() == today]
    
    query = update.callback_query
    msg = query.message if query else update.message
    
    if not orders:
        await msg.reply_text("📭 Tidak ada order hari ini untuk diexport.")
        return
    
    if query:
        await query.answer()
    
    await msg.reply_text(f"⏳ Membuat Excel untuk {len(orders)} order...")
    
    try:
        # Jalankan di thread terpisah agar tidak memblokir bot
        import asyncio
        loop = asyncio.get_event_loop()
        path = await loop.run_in_executor(None, export_to_excel, orders, chat_id)
        timestamp = datetime.now().strftime("%d%m%Y_%H%M")
        filename = f"orders_{timestamp}.xlsx"
        
        keyboard = [[InlineKeyboardButton("🗑️ Clear pending orders", callback_data="clear")]]
        
        with open(path, "rb") as f:
            await msg.reply_document(
                document=f,
                filename=filename,
                caption=f"✅ *{len(orders)} order* berhasil diexport!\n\nSiap diupload ke agregator.",
                parse_mode="Markdown",
                reply_markup=InlineKeyboardMarkup(keyboard)
            )
        
        try:
            os.remove(path)
        except Exception as e_rm:
            logging.warning(f"Could not remove temp file {path}: {e_rm}")
    except Exception as e:
        logging.error(f"Export error: {e}")
        await msg.reply_text(f"❌ Error saat export: {str(e)}")

async def clear_orders(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    count = len(pending_orders.get(chat_id, []))
    pending_orders[chat_id] = []
    save_db()  # <--- Simpan perubahan (kosong)
    
    query = update.callback_query
    if query:
        await query.answer()
        await query.message.reply_text(f"🗑️ {count} order pending dihapus.")
    else:
        await update.message.reply_text(f"🗑️ {count} order pending dihapus.")

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    if query.data == "export":
        await export_orders(update, context)
    elif query.data == "list":
        await list_orders(update, context)
    elif query.data == "clear":
        await clear_orders(update, context)
    elif query.data.startswith("del_"):
        await delete_order_callback(update, context)


# ─── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    # Load data lama jika ada
    load_db()
    
    app = ApplicationBuilder().token(TOKEN).build()
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("list", list_orders))
    app.add_handler(CommandHandler("export", export_orders))
    app.add_handler(CommandHandler("clear", clear_orders))
    app.add_handler(CallbackQueryHandler(button_handler))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    
    print("Bot running...")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
