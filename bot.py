import logging
import os
import math
import re
import csv
import asyncio
import openpyxl
from collections import Counter
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.error import Forbidden, BadRequest, RetryAfter
from telegram.ext import (
    Application, CommandHandler, MessageHandler, CallbackQueryHandler,
    filters, ContextTypes, ConversationHandler
)

# --- KONFIGURASI BOT ---
# GANTI PAKAI TOKEN BARU DARI BOTFATHER
BOT_TOKEN = "8800432411:AAELcmFbgzYv1bK1zTLN3ZK2ZTp8VxvWi7k"

# ID Admin (hanya ID ini yang bisa broadcast)
ADMIN_ID = 7585723733

# File penyimpanan user yang pernah start bot
USERS_FILE = "users.txt"

# ID Supergrup (Buat ngecek member)
REQUIRED_GROUP_IDS = [-1004421755620, -1003603279803]

# Link Undangan (Buat tombol Join)
INVITE_LINKS = [
    "https://t.me/+wNHEBfYPBMBjNzdl",
    "https://t.me/kurumimybini"
]

logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)

# --- STATE UNTUK CONVERSATION HANDLER ---
(
    MENU,
    TXT_NAMA_FILE, TXT_NAMA_KONTAK, TXT_LIMIT,
    MANUAL_NAMA, MANUAL_NOMOR,
    MERGE_FILE,
    SPLIT_LIMIT,
    EDIT_NAMA, EDIT_PREFIX,
    TXT_SPLIT_LIMIT,
    # Fitur Admin & Navy
    ADMIN_NAVY_NAMA_GRUP,
    ADMIN_NAVY_NOMOR,
    ADMIN_NAVY_PILIH,
    ADMIN_NAVY_FILE_NAME,
    # Fitur Broadcast
    BROADCAST_MESSAGE,
) = range(16)

# Kamus kode negara (lengkap)
COUNTRY_CODES = {
    # Asia Tenggara
    "62": "Indonesia 🇮🇩",
    "60": "Malaysia 🇲🇾",
    "65": "Singapore 🇸🇬",
    "66": "Thailand 🇹🇭",
    "84": "Vietnam 🇻🇳",
    "63": "Philippines 🇵🇭",
    "855": "Cambodia 🇰🇭",
    "856": "Laos 🇱🇦",
    "95": "Myanmar 🇲🇲",
    "673": "Brunei 🇧🇳",
    "670": "Timor-Leste 🇹🇱",
    # Asia Selatan
    "91": "India 🇮🇳",
    "92": "Pakistan 🇵🇰",
    "880": "Bangladesh 🇧🇩",
    "94": "Sri Lanka 🇱🇰",
    "977": "Nepal 🇳🇵",
    "975": "Bhutan 🇧🇹",
    "960": "Maldives 🇲🇻",
    "93": "Afghanistan 🇦🇫",
    # Asia Timur
    "86": "China 🇨🇳",
    "81": "Japan 🇯🇵",
    "82": "South Korea 🇰🇷",
    "850": "North Korea 🇰🇵",
    "886": "Taiwan 🇹🇼",
    "852": "Hong Kong 🇭🇰",
    "853": "Macau 🇲🇴",
    "976": "Mongolia 🇲🇳",
    # Timur Tengah
    "971": "UAE 🇦🇪",
    "966": "Saudi Arabia 🇸🇦",
    "974": "Qatar 🇶🇦",
    "965": "Kuwait 🇰🇼",
    "968": "Oman 🇴🇲",
    "973": "Bahrain 🇧🇭",
    "962": "Jordan 🇯🇴",
    "961": "Lebanon 🇱🇧",
    "963": "Syria 🇸🇾",
    "964": "Iraq 🇮🇶",
    "98": "Iran 🇮🇷",
    "972": "Israel 🇮🇱",
    "970": "Palestine 🇵🇸",
    "967": "Yemen 🇾🇪",
    "90": "Turkey 🇹🇷",
    "994": "Azerbaijan 🇦🇿",
    "995": "Georgia 🇬🇪",
    "374": "Armenia 🇦🇲",
    # Eropa
    "44": "United Kingdom 🇬🇧",
    "33": "France 🇫🇷",
    "49": "Germany 🇩🇪",
    "39": "Italy 🇮🇹",
    "34": "Spain 🇪🇸",
    "31": "Netherlands 🇳🇱",
    "32": "Belgium 🇧🇪",
    "41": "Switzerland 🇨🇭",
    "43": "Austria 🇦🇹",
    "46": "Sweden 🇸🇪",
    "47": "Norway 🇳🇴",
    "45": "Denmark 🇩🇰",
    "358": "Finland 🇫🇮",
    "48": "Poland 🇵🇱",
    "420": "Czech Republic 🇨🇿",
    "36": "Hungary 🇭🇺",
    "40": "Romania 🇷🇴",
    "359": "Bulgaria 🇧🇬",
    "30": "Greece 🇬🇷",
    "351": "Portugal 🇵🇹",
    "353": "Ireland 🇮🇪",
    "7": "Russia 🇷🇺",
    "380": "Ukraine 🇺🇦",
    "375": "Belarus 🇧🇾",
    "370": "Lithuania 🇱🇹",
    "371": "Latvia 🇱🇻",
    "372": "Estonia 🇪🇪",
    "385": "Croatia 🇭🇷",
    "381": "Serbia 🇷🇸",
    "387": "Bosnia 🇧🇦",
    "382": "Montenegro 🇲🇪",
    "389": "North Macedonia 🇲🇰",
    "383": "Kosovo 🇽🇰",
    "355": "Albania 🇦🇱",
    "386": "Slovenia 🇸🇮",
    "421": "Slovakia 🇸🇰",
    "352": "Luxembourg 🇱🇺",
    "356": "Malta 🇲🇹",
    "357": "Cyprus 🇨🇾",
    "354": "Iceland 🇮🇸",
    # Amerika
    "1": "USA/Canada 🇺🇸",
    "52": "Mexico 🇲🇽",
    "55": "Brazil 🇧🇷",
    "54": "Argentina 🇦🇷",
    "56": "Chile 🇨🇱",
    "57": "Colombia 🇨🇴",
    "51": "Peru 🇵🇪",
    "58": "Venezuela 🇻🇪",
    "593": "Ecuador 🇪🇨",
    "591": "Bolivia 🇧🇴",
    "595": "Paraguay 🇵🇾",
    "598": "Uruguay 🇺🇾",
    "507": "Panama 🇵🇦",
    "506": "Costa Rica 🇨🇷",
    "503": "El Salvador 🇸🇻",
    "502": "Guatemala 🇬🇹",
    "504": "Honduras 🇭🇳",
    "505": "Nicaragua 🇳🇮",
    "53": "Cuba 🇨🇺",
    "1809": "Dominican Republic 🇩🇴",
    "1876": "Jamaica 🇯🇲",
    "1868": "Trinidad & Tobago 🇹🇹",
    # Afrika
    "20": "Egypt 🇪🇬",
    "27": "South Africa 🇿🇦",
    "234": "Nigeria 🇳🇬",
    "254": "Kenya 🇰🇪",
    "233": "Ghana 🇬🇭",
    "212": "Morocco 🇲🇦",
    "213": "Algeria 🇩🇿",
    "216": "Tunisia 🇹🇳",
    "218": "Libya 🇱🇾",
    "251": "Ethiopia 🇪🇹",
    "255": "Tanzania 🇹🇿",
    "256": "Uganda 🇺🇬",
    "250": "Rwanda 🇷🇼",
    "257": "Burundi 🇧🇮",
    "260": "Zambia 🇿🇲",
    "263": "Zimbabwe 🇿🇼",
    "265": "Malawi 🇲🇼",
    "267": "Botswana 🇧🇼",
    "268": "Eswatini 🇸🇿",
    "266": "Lesotho 🇱🇸",
    "258": "Mozambique 🇲🇿",
    "261": "Madagascar 🇲🇬",
    "230": "Mauritius 🇲🇺",
    "248": "Seychelles 🇸🇨",
    "221": "Senegal 🇸🇳",
    "225": "Ivory Coast 🇨🇮",
    "226": "Burkina Faso 🇧🇫",
    "223": "Mali 🇲🇱",
    "227": "Niger 🇳🇪",
    "228": "Togo 🇹🇬",
    "229": "Benin 🇧🇯",
    "241": "Gabon 🇬🇦",
    "242": "Congo 🇨🇬",
    "243": "DR Congo 🇨🇩",
    "237": "Cameroon 🇨🇲",
    "235": "Chad 🇹🇩",
    "236": "Central African Republic 🇨🇫",
    "240": "Equatorial Guinea 🇬🇶",
    "239": "Sao Tome 🇸🇹",
    "244": "Angola 🇦🇴",
    "249": "Sudan 🇸🇩",
    "211": "South Sudan 🇸🇸",
    "252": "Somalia 🇸🇴",
    "253": "Djibouti 🇩🇯",
    "291": "Eritrea 🇪🇷",
    # Oseania
    "61": "Australia 🇦🇺",
    "64": "New Zealand 🇳🇿",
    "679": "Fiji 🇫🇯",
    "675": "Papua New Guinea 🇵🇬",
    "677": "Solomon Islands 🇸🇧",
    "678": "Vanuatu 🇻🇺",
    "685": "Samoa 🇼🇸",
    "676": "Tonga 🇹🇴",
    "680": "Palau 🇵🇼",
    "691": "Micronesia 🇫🇲",
    "692": "Marshall Islands 🇲🇭",
    "674": "Nauru 🇳🇷",
    "686": "Kiribati 🇰🇮",
    "688": "Tuvalu 🇹🇻",
}

# --- FUNGSI CEK FORCE SUBSCRIBE ---
async def is_user_joined(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    belum_join = []

    for group_id in REQUIRED_GROUP_IDS:
        try:
            member = await context.bot.get_chat_member(chat_id=group_id, user_id=user_id)
            if member.status in ['left', 'kicked']:
                belum_join.append(group_id)
        except Exception as e:
            logging.error(f"Gagal cek member di {group_id}: {e}")
            belum_join.append(group_id)

    return belum_join

# --- HELPER USER LIST (untuk broadcast) ---
def load_users():
    if not os.path.exists(USERS_FILE):
        return set()
    with open(USERS_FILE, 'r') as f:
        return set(int(line.strip()) for line in f if line.strip().isdigit())

def save_user(user_id: int):
    users = load_users()
    if user_id not in users:
        with open(USERS_FILE, 'a') as f:
            f.write(str(user_id) + "\n")

# --- MENU UTAMA ---
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()

    # Simpan user yang start bot (untuk broadcast)
    if update.effective_user:
        save_user(update.effective_user.id)

    belum_join = await is_user_joined(update, context)
    if belum_join:
        keyboard = []
        for i, link in enumerate(INVITE_LINKS):
            keyboard.append([InlineKeyboardButton(f"Join Grup Wajib {i+1}", url=link)])

        reply_markup = InlineKeyboardMarkup(keyboard)
        pesan = (
            "⚠️ **Wajib Join Dulu Bray!**\n\n"
            "Sebelum pakai bot **CV Kurumi**, kamu harus join grup/channel di bawah ini.\n"
            "Setelah join, klik /start lagi ya."
        )

        if update.message:
            await update.message.reply_text(pesan, reply_markup=reply_markup, parse_mode='Markdown')
        else:
            await update.callback_query.edit_message_text(pesan, reply_markup=reply_markup, parse_mode='Markdown')
        return MENU

    keyboard = [
        [InlineKeyboardButton("?? Konversi", callback_data='cat_konversi')],
        [InlineKeyboardButton("✏️ Edit & Input", callback_data='cat_edit')],
        [InlineKeyboardButton("🔗 Gabung & Pecah", callback_data='cat_gabung')],
        [InlineKeyboardButton("🛠️ Tools", callback_data='cat_tools')],
        [InlineKeyboardButton("💰 Lainnya", callback_data='cat_lainnya')],
    ]

    # Tombol Broadcast hanya muncul untuk Admin
    if update.effective_user and update.effective_user.id == ADMIN_ID:
        keyboard.append([InlineKeyboardButton("📢 Broadcast (Admin)", callback_data='menu_broadcast')])

    reply_markup = InlineKeyboardMarkup(keyboard)
    pesan = "Halo bray! Selamat datang di **CV Kurumi** 🦄.\nSilakan pilih kategori di bawah ini:"

    if update.message:
        await update.message.reply_text(pesan, reply_markup=reply_markup, parse_mode='Markdown')
    else:
        await update.callback_query.edit_message_text(pesan, reply_markup=reply_markup, parse_mode='Markdown')
    return MENU

# --- HANDLER TOMBOL MENU ---
async def menu_pilihan(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data

    # ===== KATEGORI =====
    if data == 'cat_konversi':
        keyboard = [
            [InlineKeyboardButton("📄 TXT → VCF", callback_data='menu_txt_to_vcf')],
            [InlineKeyboardButton("📇 VCF → TXT", callback_data='menu_vcf_to_txt')],
            [InlineKeyboardButton("📊 XLSX → TXT", callback_data='menu_xlsx_to_txt')],
            [InlineKeyboardButton("🔙 Kembali", callback_data='cat_back')],
        ]
        await query.edit_message_text(
            "📄 **Konversi**\n\nPilih jenis konversi:",
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode='Markdown'
        )
        return MENU

    elif data == 'cat_edit':
        keyboard = [
            [InlineKeyboardButton("✏️ Edit VCF", callback_data='menu_edit_vcf')],
            [InlineKeyboardButton("✍️ Input Manual", callback_data='menu_manual')],
            [InlineKeyboardButton("👮‍♂️ 1 File VCF (Admin & Navy)", callback_data='menu_admin_navy')],
            [InlineKeyboardButton("🔙 Kembali", callback_data='cat_back')],
        ]
        await query.edit_message_text(
            "✏️ **Edit & Input**\n\nPilih fitur:",
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode='Markdown'
        )
        return MENU

    elif data == 'cat_gabung':
        keyboard = [
            [InlineKeyboardButton("🔗 Gabung VCF", callback_data='menu_merge_vcf')],
            [InlineKeyboardButton("✂️ Pecah VCF", callback_data='menu_split_vcf')],
            [InlineKeyboardButton("📄 Pecah TXT", callback_data='menu_split_txt')],
            [InlineKeyboardButton("🔙 Kembali", callback_data='cat_back')],
        ]
        await query.edit_message_text(
            "🔗 **Gabung & Pecah**\n\nPilih fitur:",
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode='Markdown'
        )
        return MENU

    elif data == 'cat_tools':
        keyboard = [
            [InlineKeyboardButton("🔍 Cek File (TXT/XLSX)", callback_data='menu_cek_file')],
            [InlineKeyboardButton("🔙 Kembali", callback_data='cat_back')],
        ]
        await query.edit_message_text(
            "🛠️ **Tools**\n\nPilih fitur:",
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode='Markdown'
        )
        return MENU

    elif data == 'cat_lainnya':
        keyboard = [
            [InlineKeyboardButton("💰 Donate", callback_data='menu_donate')],
            [InlineKeyboardButton("🔙 Kembali", callback_data='cat_back')],
        ]
        await query.edit_message_text(
            "💰 **Lainnya**\n\nPilih menu:",
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode='Markdown'
        )
        return MENU

    elif data == 'cat_back':
        # Kembali ke menu utama
        keyboard = [
            [InlineKeyboardButton("📄 Konversi", callback_data='cat_konversi')],
            [InlineKeyboardButton("✏️ Edit & Input", callback_data='cat_edit')],
            [InlineKeyboardButton("🔗 Gabung & Pecah", callback_data='cat_gabung')],
            [InlineKeyboardButton("🛠️ Tools", callback_data='cat_tools')],
            [InlineKeyboardButton("💰 Lainnya", callback_data='cat_lainnya')],
        ]
        if query.from_user.id == ADMIN_ID:
            keyboard.append([InlineKeyboardButton("📢 Broadcast (Admin)", callback_data='menu_broadcast')])

        await query.edit_message_text(
            "Halo bray! Selamat datang di **CV Kurumi** 🦄.\nSilakan pilih kategori di bawah ini:",
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode='Markdown'
        )
        return MENU

    # ===== FITUR ASLI =====
    if data == 'menu_txt_to_vcf':
        await query.edit_message_text("📄 **TXT ke VCF**\n\nSilakan kirim file `.txt` berisi nomor.")
        context.user_data['waiting_for'] = 'txt'
        return MENU
    elif data == 'menu_vcf_to_txt':
        await query.edit_message_text("📇 **VCF ke TXT**\n\nSilakan kirim file `.vcf`.")
        context.user_data['waiting_for'] = 'vcf_to_txt'
        return MENU
    elif data == 'menu_xlsx_to_txt':
        await query.edit_message_text("📊 **XLSX ke TXT**\n\nSilakan kirim file `.xlsx` atau `.csv`.")
        context.user_data['waiting_for'] = 'xlsx_to_txt'
        return MENU
    elif data == 'menu_edit_vcf':
        await query.edit_message_text("✏️ **Edit VCF**\n\nSilakan kirim file `.vcf`.")
        context.user_data['waiting_for'] = 'edit'
        return MENU
    elif data == 'menu_manual':
        context.user_data['manual_data'] = []
        await query.edit_message_text("✍️ **Input Manual**\n\nKetik **Nama Kontak** (contoh: `NAFF`):")
        return MANUAL_NAMA
    elif data == 'menu_merge_vcf':
        context.user_data['merge_files'] = []
        await query.edit_message_text("🔗 **Gabung VCF**\n\nKirim file `.vcf` pertama.")
        context.user_data['waiting_for'] = 'merge'
        return MERGE_FILE
    elif data == 'menu_split_vcf':
        await query.edit_message_text("✂️ **Pecah VCF**\n\nSilakan kirim file `.vcf` yang mau dipecah.")
        context.user_data['waiting_for'] = 'split'
        return MENU
    elif data == 'menu_split_txt':
        await query.edit_message_text("📄 **Pecah TXT**\n\nSilakan kirim file `.txt` yang mau dipecah.")
        context.user_data['waiting_for'] = 'split_txt'
        return MENU
    elif data == 'menu_cek_file':
        await query.edit_message_text(
            "🔍 **Cek File (TXT / XLSX / CSV)**\n\n"
            "Silakan kirim file `.txt`, `.xlsx`, atau `.csv`.\n"
            "Bot akan hitung Total, Unik, Duplikat + deteksi negara."
        )
        context.user_data['waiting_for'] = 'cek_file'
        return MENU
    elif data == 'menu_admin_navy':
        context.user_data['an_groups'] = []
        context.user_data['an_current_numbers'] = []
        await query.edit_message_text(
            "👮‍♂️ **1 File VCF (Admin & Navy)**\n\n"
            "Ketik **Nama Grup** pertama (contoh: `Admin Gacor`):"
        )
        return ADMIN_NAVY_NAMA_GRUP

    elif data == 'menu_broadcast':
        if query.from_user.id != ADMIN_ID:
            await query.answer("⛔ Kamu bukan admin!", show_alert=True)
            return MENU
        await query.edit_message_text(
            "📢 **Broadcast Message**\n\n"
            "Kirim pesan yang ingin disiarkan ke **semua user** bot ini.\n"
            "Bisa berupa teks biasa.\n\n"
            "Ketik `/cancel` untuk batal."
        )
        return BROADCAST_MESSAGE

    elif data == 'menu_donate':
        qris_path = "qris.png"
        if not os.path.exists(qris_path):
            await query.edit_message_text("❌ File QRIS belum tersedia. Hubungi admin.")
            return MENU

        await query.answer()
        with open(qris_path, 'rb') as photo:
            await context.bot.send_photo(
                chat_id=query.message.chat_id,
                photo=photo,
                caption=(
                    "💰 **Donasi ke Admin CV Kurumi**\n\n"
                    "Scan QRIS di atas untuk mendukung admin.\n"
                    "Terima kasih banyak atas dukungannya! ❤️"
                ),
                parse_mode='Markdown'
            )
        return MENU

    return MENU

# --- HANDLER FILE ---
async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    ext = doc.file_name.split('.')[-1].lower()

    context.user_data['file_id'] = doc.file_id
    context.user_data['file_name'] = doc.file_name
    mode = context.user_data.get('waiting_for')

    if mode == 'txt':
        if ext != 'txt':
            await update.message.reply_text("❌ Harus file .txt ya!")
            return MENU
        await update.message.reply_text("Ketik **Nama File** (contoh: `Kontak_NAFF`):")
        return TXT_NAMA_FILE

    elif mode == 'vcf_to_txt':
        await proses_vcf_to_txt(update, context)
        return MENU

    elif mode == 'xlsx_to_txt':
        await proses_xlsx_to_txt(update, context)
        return MENU

    elif mode == 'edit':
        await update.message.reply_text("Ketik **Nama Kontak Baru** (atau `tidak` untuk skip):")
        return EDIT_NAMA

    elif mode == 'merge':
        context.user_data['merge_files'].append(doc.file_id)
        await update.message.reply_text(
            f"✅ File VCF ke-{len(context.user_data['merge_files'])} ditambahkan.\n"
            "Kirim lagi, atau ketik `selesai` jika sudah."
        )
        return MERGE_FILE

    elif mode == 'split':
        await update.message.reply_text("Ketik **Jumlah kontak per file** (contoh: `500`):")
        return SPLIT_LIMIT

    elif mode == 'split_txt':
        if ext != 'txt':
            await update.message.reply_text("❌ Harus file .txt ya!")
            return MENU
        await update.message.reply_text("Ketik **Jumlah baris per file** (contoh: `500`):")
        return TXT_SPLIT_LIMIT

    elif mode == 'cek_file':
        if ext not in ['txt', 'xlsx', 'csv']:
            await update.message.reply_text("❌ Harus file .txt / .xlsx / .csv ya!")
            return MENU
        await proses_cek_file(update, context)
        return MENU

    return MENU

# ==================== FITUR BARU 1: CEK FILE ====================
def extract_numbers_from_file(path: str, ext: str) -> list:
    nomor_list = []
    if ext == 'txt':
        with open(path, 'r', encoding='utf-8', errors='ignore') as f:
            for line in f:
                n = ''.join(filter(str.isdigit, line.strip()))
                if len(n) >= 8:
                    nomor_list.append(n)
    elif ext == 'xlsx':
        wb = openpyxl.load_workbook(path, data_only=True)
        sheet = wb.active
        for row in sheet.iter_rows(values_only=True):
            for cell in row:
                if cell is not None:
                    n = ''.join(filter(str.isdigit, str(cell)))
                    if len(n) >= 8:
                        nomor_list.append(n)
    elif ext == 'csv':
        with open(path, 'r', encoding='utf-8', errors='ignore') as f:
            reader = csv.reader(f)
            for row in reader:
                for cell in row:
                    n = ''.join(filter(str.isdigit, str(cell)))
                    if len(n) >= 8:
                        nomor_list.append(n)
    return nomor_list

def detect_country(nomor: str) -> str:
    # Coba kode panjang dulu
    for code in sorted(COUNTRY_CODES.keys(), key=len, reverse=True):
        if nomor.startswith(code):
            return COUNTRY_CODES[code]
    return "Tidak diketahui ❓"

async def proses_cek_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    file_id = context.user_data['file_id']
    file_name = context.user_data['file_name']
    ext = file_name.split('.')[-1].lower()

    file = await context.bot.get_file(file_id)
    path = f"temp_cek_{file_id}.{ext}"
    await file.download_to_drive(path)

    nomor_list = extract_numbers_from_file(path, ext)
    os.remove(path)

    if not nomor_list:
        await update.message.reply_text("❌ Tidak ditemukan nomor valid di file.")
        context.user_data.clear()
        return

    total = len(nomor_list)
    counter = Counter(nomor_list)
    unik = len(counter)
    duplikat = total - unik

    # Deteksi negara
    country_count = Counter()
    for n in counter.keys():  # pakai unik saja biar lebih cepat
        country_count[detect_country(n)] += 1

    # Susun pesan
    pesan = (
        f"🔍 **Hasil Cek File**\n"
        f"📁 `{file_name}`\n\n"
        f"📊 **Statistik Nomor:**\n"
        f"• Total      : `{total}`\n"
        f"• Unik       : `{unik}`\n"
        f"• Duplikat   : `{duplikat}`\n\n"
        f"🌍 **Deteksi Negara (berdasarkan nomor unik):**\n"
    )
    for negara, jumlah in country_count.most_common():
        pesan += f"• {negara}: `{jumlah}`\n"

    keyboard = []
    if duplikat > 0:
        # Simpan data sementara untuk hapus duplikat
        context.user_data['cek_nomor_unik'] = list(counter.keys())
        context.user_data['cek_file_name'] = file_name
        keyboard.append([InlineKeyboardButton("🗑️ Hapus Duplikat & Download", callback_data='cek_hapus_duplikat')])

    keyboard.append([InlineKeyboardButton("🔙 Kembali ke Menu", callback_data='cek_kembali')])

    await update.message.reply_text(
        pesan,
        reply_markup=InlineKeyboardMarkup(keyboard),
        parse_mode='Markdown'
    )

async def cek_aksi(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if query.data == 'cek_hapus_duplikat':
        nomor_unik = context.user_data.get('cek_nomor_unik', [])
        if not nomor_unik:
            await query.edit_message_text("❌ Data sudah tidak tersedia. Ketik /start lagi.")
            return MENU

        nama_txt = "Nomor_Unik.txt"
        with open(nama_txt, 'w') as f:
            for n in nomor_unik:
                f.write(n + "\n")

        await context.bot.send_document(
            chat_id=query.message.chat_id,
            document=open(nama_txt, 'rb'),
            caption=f"✅ {len(nomor_unik)} nomor unik (duplikat sudah dihapus)"
        )
        os.remove(nama_txt)

        await query.edit_message_text(
            "🎉 Selesai! File nomor unik sudah dikirim.\nKetik /start untuk kembali ke menu."
        )
        context.user_data.clear()
        return MENU

    elif query.data == 'cek_kembali':
        await query.edit_message_text("Ketik /start untuk kembali ke menu utama.")
        context.user_data.clear()
        return MENU

    return MENU

# ==================== FITUR BARU 2: 1 FILE VCF (ADMIN & NAVY) ====================
async def an_nama_grup(update: Update, context: ContextTypes.DEFAULT_TYPE):
    nama = update.message.text.strip()
    if not nama:
        await update.message.reply_text("❌ Nama grup tidak boleh kosong. Ketik lagi:")
        return ADMIN_NAVY_NAMA_GRUP

    context.user_data['an_current_name'] = nama
    context.user_data['an_current_numbers'] = []
    await update.message.reply_text(
        f"✅ Nama grup: **{nama}**\n\n"
        f"Sekarang kirim nomor.\n"
        f"• Bisa satu nomor\n"
        f"• Bisa banyak nomor sekaligus (pisah spasi / baris baru)\n\n"
        f"Contoh:\n"
        f"`628123456789`\n"
        f"atau\n"
        f"`628111 628222 628333`\n\n"
        f"Nomor akan otomatis dinamai:\n"
        f"`{nama} 1`, `{nama} 2`, dst.\n\n"
        f"Kalau sudah selesai untuk grup ini, ketik: `selesai`",
        parse_mode='Markdown'
    )
    return ADMIN_NAVY_NOMOR

async def an_nomor(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()

    # Cek perintah selesai
    if text.lower() == 'selesai':
        numbers = context.user_data.get('an_current_numbers', [])
        if not numbers:
            await update.message.reply_text("❌ Belum ada nomor. Kirim minimal 1 nomor dulu.")
            return ADMIN_NAVY_NOMOR

        # Simpan grup saat ini
        context.user_data['an_groups'].append({
            "name": context.user_data['an_current_name'],
            "numbers": numbers.copy()
        })

        total_grup = len(context.user_data['an_groups'])
        total_nomor = sum(len(g['numbers']) for g in context.user_data['an_groups'])

        # Ringkasan grup yang sudah ada
        ringkasan = ""
        for g in context.user_data['an_groups']:
            ringkasan += f"• **{g['name']}** → {len(g['numbers'])} nomor\n"

        keyboard = [
            [InlineKeyboardButton("➕ Tambah Grup Lagi", callback_data='an_tambah_grup')],
            [InlineKeyboardButton("✅ Buat 1 File VCF Sekarang", callback_data='an_buat_file')],
        ]
        await update.message.reply_text(
            f"✅ Grup **{context.user_data['an_current_name']}** berhasil disimpan "
            f"({len(numbers)} nomor).\n\n"
            f"📋 **Ringkasan sementara:**\n{ringkasan}\n"
            f"Total: **{total_grup} grup** | **{total_nomor} nomor**\n\n"
            f"Mau tambah grup lain (contoh: Navy) atau langsung buat file?",
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode='Markdown'
        )
        return ADMIN_NAVY_PILIH

    # Ambil semua nomor dari pesan (bisa banyak sekaligus)
    # Pisah berdasarkan spasi, koma, atau baris baru
    potongan = re.split(r'[\s,;]+', text)
    nomor_baru = []
    for p in potongan:
        n = ''.join(filter(str.isdigit, p))
        if len(n) >= 8:
            nomor_baru.append(n)

    if not nomor_baru:
        await update.message.reply_text(
            "❌ Tidak ada nomor valid.\n"
            "Kirim nomor (bisa banyak sekaligus) atau ketik `selesai`."
        )
        return ADMIN_NAVY_NOMOR

    # Tambahkan semua nomor yang valid
    for n in nomor_baru:
        context.user_data['an_current_numbers'].append(n)

    total_di_grup = len(context.user_data['an_current_numbers'])
    nama_grup = context.user_data['an_current_name']

    # Tampilkan ringkasan penambahan
    if len(nomor_baru) == 1:
        urutan = total_di_grup
        pesan = (
            f"✅ Ditambahkan: **{nama_grup} {urutan:03d}** → `+{nomor_baru[0]}`\n"
            f"Total di grup ini: **{total_di_grup}** nomor.\n\n"
            f"Kirim nomor lagi atau ketik `selesai`."
        )
    else:
        pesan = (
            f"✅ Berhasil menambahkan **{len(nomor_baru)}** nomor ke grup **{nama_grup}**.\n"
            f"Total di grup ini sekarang: **{total_di_grup}** nomor.\n\n"
            f"Kirim nomor lagi atau ketik `selesai`."
        )

    await update.message.reply_text(pesan, parse_mode='Markdown')
    return ADMIN_NAVY_NOMOR

async def an_pilih(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if query.data == 'an_tambah_grup':
        await query.edit_message_text(
            "Ketik **Nama Grup** berikutnya (contoh: `Navy Elite` / `Admin 2`):"
        )
        return ADMIN_NAVY_NAMA_GRUP

    elif query.data == 'an_buat_file':
        await query.edit_message_text(
            "📝 Ketik **Nama File** yang diinginkan (tanpa .vcf)\n"
            "Contoh: `Admin_Navy_Final`"
        )
        return ADMIN_NAVY_FILE_NAME

    return ADMIN_NAVY_PILIH

async def an_nama_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    nama_file = update.message.text.strip()
    if not nama_file:
        await update.message.reply_text("❌ Nama file tidak boleh kosong. Ketik lagi:")
        return ADMIN_NAVY_FILE_NAME

    # Bersihkan karakter aneh
    nama_file = re.sub(r'[\\/*?:"<>|]', '', nama_file)
    if not nama_file.endswith('.vcf'):
        nama_file += '.vcf'

    groups = context.user_data.get('an_groups', [])
    if not groups:
        await update.message.reply_text("❌ Tidak ada data. Ketik /start ulang.")
        return MENU

    await update.message.reply_text("⏳ Sedang membuat file VCF...")

    with open(nama_file, 'w', encoding='utf-8') as f:
        for grup in groups:
            for idx, nomor in enumerate(grup['numbers'], start=1):
                # Nomor urut dengan padding 001, 002, dst
                fn = f"{grup['name']} {idx:03d}"
                f.write(
                    f"BEGIN:VCARD\n"
                    f"VERSION:3.0\n"
                    f"FN:{fn}\n"
                    f"TEL;TYPE=CELL:+{nomor}\n"
                    f"END:VCARD\n"
                )

    total_kontak = sum(len(g['numbers']) for g in groups)

    # Buat ringkasan untuk caption
    ringkasan = ""
    for g in groups:
        ringkasan += f"• {g['name']}: {len(g['numbers'])} kontak\n"

    await context.bot.send_document(
        chat_id=update.effective_chat.id,
        document=open(nama_file, 'rb'),
        caption=(
            f"✅ **1 File VCF siap!**\n"
            f"📁 `{nama_file}`\n\n"
            f"{ringkasan}\n"
            f"👥 Total: **{len(groups)} grup** | **{total_kontak} kontak**"
        ),
        parse_mode='Markdown'
    )
    os.remove(nama_file)

    await update.message.reply_text(
        "🎉 Selesai! Terima kasih sudah pakai **CV Kurumi**.\nKetik /start untuk menu lagi."
    )
    context.user_data.clear()
    return MENU

# ==================== FITUR BROADCAST (ADMIN ONLY) ====================
async def broadcast_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # Pastikan hanya admin
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("⛔ Kamu tidak punya akses.")
        return MENU

    pesan = update.message.text
    if not pesan:
        await update.message.reply_text("❌ Pesan kosong. Kirim teks yang ingin di-broadcast.")
        return BROADCAST_MESSAGE

    users = list(load_users())
    if not users:
        await update.message.reply_text("❌ Belum ada user yang tersimpan.\n\nUser harus /start dulu supaya masuk daftar.")
        return MENU

    total = len(users)
    await update.message.reply_text(
        f"⏳ Mengirim broadcast ke **{total}** user...\n"
        f"Mohon tunggu, jangan spam /start."
    )

    sukses = 0
    gagal = 0
    blocked = []

    for i, uid in enumerate(users, start=1):
        try:
            await context.bot.send_message(
                chat_id=uid,
                text=f"📢 **Pengumuman dari Admin CV Kurumi**\n\n{pesan}",
                parse_mode='Markdown'
            )
            sukses += 1

            # Jeda kecil supaya tidak kena flood limit Telegram
            await asyncio.sleep(0.07)

        except Forbidden:
            # User block bot atau belum pernah chat
            gagal += 1
            blocked.append(uid)
        except RetryAfter as e:
            # Kena rate limit, tunggu sesuai saran Telegram
            await asyncio.sleep(e.retry_after + 1)
            try:
                await context.bot.send_message(
                    chat_id=uid,
                    text=f"📢 **Pengumuman dari Admin CV Kurumi**\n\n{pesan}",
                    parse_mode='Markdown'
                )
                sukses += 1
            except Exception:
                gagal += 1
        except BadRequest:
            gagal += 1
        except Exception as e:
            logging.error(f"Gagal kirim ke {uid}: {e}")
            gagal += 1

        # Update progress setiap 25 user
        if i % 25 == 0:
            try:
                await update.message.reply_text(f"⏳ Progress: {i}/{total} ...")
            except Exception:
                pass

    # Hapus user yang sudah block bot dari daftar
    if blocked:
        remaining = [u for u in users if u not in blocked]
        with open(USERS_FILE, 'w') as f:
            for u in remaining:
                f.write(str(u) + "\n")

    await update.message.reply_text(
        f"✅ **Broadcast selesai!**\n\n"
        f"• Total user : `{total}`\n"
        f"• Berhasil   : `{sukses}`\n"
        f"• Gagal      : `{gagal}`\n"
        f"• Dihapus (block) : `{len(blocked)}`\n\n"
        f"Ketik /start untuk kembali ke menu."
    )
    return MENU

# ==================== FITUR LAMA (TIDAK DIUBAH) ====================

# --- TXT TO VCF ---
async def txt_nama_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['nama_file'] = update.message.text
    await update.message.reply_text("Ketik **Nama Kontak** (contoh: `NAFF`):")
    return TXT_NAMA_KONTAK

async def txt_nama_kontak(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['nama_kontak'] = update.message.text
    await update.message.reply_text("Mau dipecah berapa nomor per file? (contoh: `500`):")
    return TXT_LIMIT

async def txt_limit(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        limit = int(update.message.text)
        context.user_data['limit'] = limit
        await update.message.reply_text("⏳ **CV Kurumi** sedang memproses TXT ke VCF...")
        await proses_txt_to_vcf(update, context)
    except ValueError:
        await update.message.reply_text("❌ Harus angka ya bray!")
        return TXT_LIMIT
    return MENU

async def proses_txt_to_vcf(update, context):
    file_id = context.user_data['file_id']
    file = await context.bot.get_file(file_id)
    path = f"temp_{file_id}.txt"
    await file.download_to_drive(path)

    nomor_list = []
    with open(path, 'r') as f:
        for line in f:
            n = ''.join(filter(str.isdigit, line.strip()))
            if len(n) >= 8:
                nomor_list.append(n)
    os.remove(path)

    limit = context.user_data['limit']
    total_f = math.ceil(len(nomor_list) / limit)

    for i in range(total_f):
        chunk = nomor_list[i*limit : (i+1)*limit]
        # Nama file: "Testing 1.vcf" (tanpa kata Part)
        nama_vcf = f"{context.user_data['nama_file']} {i+1}.vcf"
        with open(nama_vcf, 'w', encoding='utf-8') as f:
            for j, n in enumerate(chunk):
                # Nomor urut dengan padding 001, 002, dst
                urutan = i * limit + j + 1
                f.write(
                    f"BEGIN:VCARD\nVERSION:3.0\n"
                    f"FN:{context.user_data['nama_kontak']} {urutan:03d}\n"
                    f"TEL;TYPE=CELL:+{n}\nEND:VCARD\n"
                )
        await context.bot.send_document(
            chat_id=update.effective_chat.id,
            document=open(nama_vcf, 'rb'),
            caption=f"{context.user_data['nama_file']} {i+1} ({len(chunk)} kontak)"
        )
        os.remove(nama_vcf)

    await context.bot.send_message(
        chat_id=update.effective_chat.id,
        text="🎉 Selesai! Terima kasih sudah pakai **CV Kurumi**. Ketik /start buat menu lagi."
    )
    context.user_data.clear()

# --- INPUT MANUAL ---
async def manual_nama(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['temp_nama'] = update.message.text
    await update.message.reply_text("Ketik **Nomor HP** (contoh: `628123456789`):")
    return MANUAL_NOMOR

async def manual_nomor(update: Update, context: ContextTypes.DEFAULT_TYPE):
    nomor = ''.join(filter(str.isdigit, update.message.text))
    nama = context.user_data['temp_nama']

    context.user_data['manual_data'].append((nama, nomor))

    keyboard = [
        [InlineKeyboardButton("➕ Tambah Lagi", callback_data='manual_lagi')],
        [InlineKeyboardButton("✅ Selesai & Simpan", callback_data='manual_selesai')]
    ]
    await update.message.reply_text(
        f"✅ {nama} ({nomor}) ditambahkan.\nTotal: {len(context.user_data['manual_data'])} kontak.",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )
    return MENU

async def manual_aksi(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if query.data == 'manual_lagi':
        await query.edit_message_text("Ketik **Nama Kontak** lagi:")
        return MANUAL_NAMA
    elif query.data == 'manual_selesai':
        await query.edit_message_text("⏳ Memproses file VCF Manual...")
        vcf_name = "Manual.vcf"
        with open(vcf_name, 'w', encoding='utf-8') as f:
            for nama, nomor in context.user_data['manual_data']:
                f.write(
                    f"BEGIN:VCARD\nVERSION:3.0\n"
                    f"FN:{nama}\nTEL;TYPE=CELL:+{nomor}\nEND:VCARD\n"
                )
        await context.bot.send_document(
            chat_id=query.message.chat_id,
            document=open(vcf_name, 'rb'),
            caption="File VCF Manual"
        )
        os.remove(vcf_name)
        await context.bot.send_message(
            chat_id=query.message.chat_id,
            text="🎉 Selesai! Terima kasih sudah pakai **CV Kurumi**. Ketik /start buat menu lagi."
        )
        context.user_data.clear()
        return MENU

# --- VCF TO TXT ---
async def proses_vcf_to_txt(update, context):
    file_id = context.user_data['file_id']
    file = await context.bot.get_file(file_id)
    path = f"temp_{file_id}.vcf"
    await file.download_to_drive(path)

    with open(path, 'r', encoding='utf-8') as f:
        matches = re.findall(r'TEL.*?:(.*)', f.read())
    os.remove(path)

    nama_txt = "Ekstrak.txt"
    with open(nama_txt, 'w') as f:
        for m in matches:
            f.write(''.join(filter(str.isdigit, m)) + "\n")

    await context.bot.send_document(
        chat_id=update.effective_chat.id,
        document=open(nama_txt, 'rb'),
        caption="Hasil Ekstrak"
    )
    os.remove(nama_txt)
    await context.bot.send_message(
        chat_id=update.effective_chat.id,
        text="🎉 Selesai! Terima kasih sudah pakai **CV Kurumi**. Ketik /start buat menu lagi."
    )
    context.user_data.clear()

# --- XLSX / CSV TO TXT ---
async def proses_xlsx_to_txt(update, context):
    file_id = context.user_data['file_id']
    file_name = context.user_data['file_name']
    ext = file_name.split('.')[-1].lower()

    file = await context.bot.get_file(file_id)
    path = f"temp_{file_id}.{ext}"
    await file.download_to_drive(path)

    nomor_list = []
    if ext == 'xlsx':
        wb = openpyxl.load_workbook(path)
        sheet = wb.active
        for row in sheet.iter_rows(values_only=True):
            for cell in row:
                if cell is not None:
                    n = ''.join(filter(str.isdigit, str(cell)))
                    if len(n) >= 8:
                        nomor_list.append(n)
    elif ext == 'csv':
        with open(path, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            for row in reader:
                for cell in row:
                    n = ''.join(filter(str.isdigit, str(cell)))
                    if len(n) >= 8:
                        nomor_list.append(n)

    os.remove(path)

    nomor_unik = list(set(nomor_list))

    nama_txt = "Hasil_XLSX.txt"
    with open(nama_txt, 'w') as f:
        for n in nomor_unik:
            f.write(n + "\n")

    await context.bot.send_document(
        chat_id=update.effective_chat.id,
        document=open(nama_txt, 'rb'),
        caption=f"✅ {len(nomor_unik)} nomor unik berhasil diekstrak!"
    )
    os.remove(nama_txt)
    await context.bot.send_message(
        chat_id=update.effective_chat.id,
        text="🎉 Selesai! Terima kasih sudah pakai **CV Kurumi**. Ketik /start buat menu lagi."
    )
    context.user_data.clear()

# --- PECAH TXT ---
async def proses_split_txt(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        limit = int(update.message.text)
    except ValueError:
        await update.message.reply_text("❌ Harus angka ya bray!")
        return TXT_SPLIT_LIMIT

    file_id = context.user_data['file_id']
    file_name = context.user_data['file_name'].replace('.txt', '')

    file = await context.bot.get_file(file_id)
    path = f"temp_{file_id}.txt"
    await file.download_to_drive(path)

    with open(path, 'r') as f:
        lines = [line.strip() for line in f if line.strip()]
    os.remove(path)

    total_lines = len(lines)
    total_f = math.ceil(total_lines / limit)

    await update.message.reply_text(f"⏳ Memproses {total_lines} baris jadi {total_f} file...")

    for i in range(total_f):
        chunk = lines[i*limit : (i+1)*limit]
        txt_name = f"{file_name} {i+1}.txt"

        with open(txt_name, 'w') as f:
            f.write("\n".join(chunk))

        await context.bot.send_document(
            chat_id=update.effective_chat.id,
            document=open(txt_name, 'rb'),
            caption=f"✅ Part {i+1} ({len(chunk)} baris)"
        )
        os.remove(txt_name)

    await context.bot.send_message(
        chat_id=update.effective_chat.id,
        text="🎉 Selesai! Terima kasih sudah pakai **CV Kurumi**."
    )
    context.user_data.clear()
    return MENU

# --- GABUNG VCF ---
async def merge_selesai(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.message.text.lower() == 'selesai':
        if len(context.user_data['merge_files']) < 2:
            await update.message.reply_text("❌ Minimal kirim 2 file VCF.")
            return MERGE_FILE

        await update.message.reply_text("⏳ Memproses penggabungan...")
        merged = "Merged.vcf"
        with open(merged, 'w', encoding='utf-8') as out_f:
            for fid in context.user_data['merge_files']:
                f = await context.bot.get_file(fid)
                p = f"temp_{fid}.vcf"
                await f.download_to_drive(p)
                with open(p, 'r', encoding='utf-8') as in_f:
                    out_f.write(in_f.read() + "\n")
                os.remove(p)

        await context.bot.send_document(
            chat_id=update.effective_chat.id,
            document=open(merged, 'rb'),
            caption="✅ File VCF Berhasil Digabung!"
        )
        os.remove(merged)
        await context.bot.send_message(
            chat_id=update.effective_chat.id,
            text="🎉 Selesai! Terima kasih sudah pakai **CV Kurumi**. Ketik /start buat menu lagi."
        )
        context.user_data.clear()
        return MENU
    else:
        await update.message.reply_text("Ketik `selesai` jika sudah selesai ngirim file.")
        return MERGE_FILE

# --- EDIT VCF ---
async def edit_nama(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['nama_baru'] = update.message.text
    await update.message.reply_text("Mau ganti prefix nomor? (Ketik `62` atau `tidak`):")
    return EDIT_PREFIX

async def edit_prefix(update: Update, context: ContextTypes.DEFAULT_TYPE):
    prefix = update.message.text.lower()
    file_id = context.user_data['file_id']

    file = await context.bot.get_file(file_id)
    path = f"temp_{file_id}.vcf"
    await file.download_to_drive(path)
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()
    os.remove(path)

    content = re.sub(r'FN:.*', f'FN:{context.user_data["nama_baru"]}', content)

    if prefix != 'tidak':
        def rep(m):
            n = m.group(1)
            if n.startswith('0'):
                n = n[1:]
            return f":+{prefix}{n}"
        content = re.sub(r':(.*)', rep, content)

    nama_baru_vcf = f"Edited_{context.user_data['file_name']}"
    with open(nama_baru_vcf, 'w', encoding='utf-8') as f:
        f.write(content)

    await context.bot.send_document(
        chat_id=update.effective_chat.id,
        document=open(nama_baru_vcf, 'rb'),
        caption="File VCF Diedit"
    )
    os.remove(nama_baru_vcf)

    await context.bot.send_message(
        chat_id=update.effective_chat.id,
        text="🎉 Selesai! Terima kasih sudah pakai **CV Kurumi**. Ketik /start buat menu lagi."
    )
    context.user_data.clear()
    return MENU

# --- PECAH VCF ---
async def split_vcf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        limit = int(update.message.text)
        file_id = context.user_data['file_id']
        file = await context.bot.get_file(file_id)
        path = f"temp_{file_id}.vcf"
        await file.download_to_drive(path)

        with open(path, 'r', encoding='utf-8') as f:
            content = f.read()
        os.remove(path)

        vcards = re.findall(r'BEGIN:VCARD.*?END:VCARD', content, re.DOTALL)
        total_f = math.ceil(len(vcards) / limit)

        for i in range(total_f):
            chunk = vcards[i*limit : (i+1)*limit]
            sp_name = f"Split {i+1}.vcf"
            with open(sp_name, 'w', encoding='utf-8') as f:
                f.write("\n".join(chunk))
            await context.bot.send_document(
                chat_id=update.effective_chat.id,
                document=open(sp_name, 'rb'),
                caption=f"Part {i+1}"
            )
            os.remove(sp_name)

        await context.bot.send_message(
            chat_id=update.effective_chat.id,
            text="🎉 Selesai! Terima kasih sudah pakai **CV Kurumi**. Ketik /start buat menu lagi."
        )
        context.user_data.clear()
    except ValueError:
        await update.message.reply_text("❌ Harus angka ya bray!")
        return SPLIT_LIMIT
    return MENU

# --- CANCEL & FALLBACK ---
async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("❌ Dibatalkan. Ketik /start buat balik ke menu **CV Kurumi**.")
    context.user_data.clear()
    return ConversationHandler.END

async def fallback_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("❓ Perintah tidak dikenali. Ketik /start buat balik ke menu **CV Kurumi**.")
    return ConversationHandler.END

# --- MAIN ---
def main():
    app = Application.builder().token(BOT_TOKEN).build()

    conv_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            MENU: [
                CallbackQueryHandler(menu_pilihan),
                CallbackQueryHandler(manual_aksi, pattern='^manual_'),
                CallbackQueryHandler(cek_aksi, pattern='^cek_'),
                MessageHandler(filters.Document.ALL, handle_file),
            ],
            TXT_NAMA_FILE: [MessageHandler(filters.TEXT & ~filters.COMMAND, txt_nama_file)],
            TXT_NAMA_KONTAK: [MessageHandler(filters.TEXT & ~filters.COMMAND, txt_nama_kontak)],
            TXT_LIMIT: [MessageHandler(filters.TEXT & ~filters.COMMAND, txt_limit)],
            MANUAL_NAMA: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_nama)],
            MANUAL_NOMOR: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_nomor)],
            MERGE_FILE: [
                MessageHandler(filters.Document.ALL, handle_file),
                MessageHandler(filters.TEXT & ~filters.COMMAND, merge_selesai)
            ],
            EDIT_NAMA: [MessageHandler(filters.TEXT & ~filters.COMMAND, edit_nama)],
            EDIT_PREFIX: [MessageHandler(filters.TEXT & ~filters.COMMAND, edit_prefix)],
            SPLIT_LIMIT: [MessageHandler(filters.TEXT & ~filters.COMMAND, split_vcf)],
            TXT_SPLIT_LIMIT: [MessageHandler(filters.TEXT & ~filters.COMMAND, proses_split_txt)],
            # Fitur Admin & Navy
            ADMIN_NAVY_NAMA_GRUP: [MessageHandler(filters.TEXT & ~filters.COMMAND, an_nama_grup)],
            ADMIN_NAVY_NOMOR: [MessageHandler(filters.TEXT & ~filters.COMMAND, an_nomor)],
            ADMIN_NAVY_PILIH: [CallbackQueryHandler(an_pilih, pattern='^an_')],
            ADMIN_NAVY_FILE_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, an_nama_file)],
            # Fitur Broadcast (Admin only)
            BROADCAST_MESSAGE: [MessageHandler(filters.TEXT & ~filters.COMMAND, broadcast_message)],
        },
        fallbacks=[
            CommandHandler("cancel", cancel),
            CommandHandler("start", start),
            MessageHandler(filters.TEXT & ~filters.COMMAND, fallback_text)
        ]
    )

    app.add_handler(conv_handler)
    print("Bot CV Kurumi berjalan...")
    app.run_polling()

if __name__ == '__main__':
    main()
