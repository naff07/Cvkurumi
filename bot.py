import logging
import os
import math
import re
import csv
import json
import asyncio
from datetime import datetime, timedelta
from collections import Counter

import openpyxl
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.error import Forbidden, BadRequest, RetryAfter
from telegram.ext import (
    Application, CommandHandler, MessageHandler, CallbackQueryHandler,
    filters, ContextTypes, ConversationHandler
)

# =========================================================
# KONFIGURASI
# =========================================================
BOT_TOKEN = "8813112237:AAGQhvZ_c-r8QmPH6cLge3jA_mLk4h-av0E"
ADMIN_ID = 7585723733

USERS_FILE = "users.txt"
FORCE_SUB_FILE = "force_sub.txt"
ACCESS_FILE = "access.json"
TRIAL_DAYS = 14

REQUIRED_GROUP_IDS = [-1004421755620, -1003603279803]
INVITE_LINKS = [
    "https://t.me/+wNHEBfYPBMBjNzdl",
    "https://t.me/kurumimybini"
]

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# =========================================================
# STATE
# =========================================================
(
    MENU,
    # /to_vcf
    TO_VCF_FILE, TO_VCF_NAMA_FILE, TO_VCF_NAMA_KONTAK, TO_VCF_LIMIT,
    # /to_txt
    TO_TXT_FILE,
    # /admin (Admin & Navy)
    ADMIN_NAMA_GRUP, ADMIN_NOMOR, ADMIN_PILIH, ADMIN_NAMA_FILE,
    # /manual
    MANUAL_NOMOR, MANUAL_NAMA, MANUAL_NAMA_FILE,
    # /add
    ADD_FILE, ADD_NOMOR,
    # /delete
    DELETE_FILE, DELETE_NOMOR,
    # /renamectc
    RENAME_CTC_FILE, RENAME_CTC_NAMA,
    # /renamefile
    RENAME_FILE, RENAME_FILE_NEW,
    # /merge
    MERGE_FILE,
    # /split
    SPLIT_FILE, SPLIT_LIMIT,
    # /count
    COUNT_FILE,
    # /nodup
    NODUP_FILE,
    # /getname
    GETNAME_FILE,
    # /generate
    GENERATE_NAMA, GENERATE_JUMLAH,
    # /getconten
    GETCONTEN_FILE,
    # broadcast
    BROADCAST_MESSAGE,
) = range(33)

# =========================================================
# KODE NEGARA
# =========================================================
COUNTRY_CODES = {
    "62": "Indonesia 🇮🇩", "60": "Malaysia 🇲🇾", "65": "Singapore 🇸🇬",
    "66": "Thailand 🇹🇭", "84": "Vietnam 🇻🇳", "63": "Philippines 🇵🇭",
    "855": "Cambodia 🇰🇭", "856": "Laos 🇱🇦", "95": "Myanmar 🇲🇲",
    "673": "Brunei 🇧🇳", "670": "Timor-Leste 🇹🇱",
    "91": "India 🇮🇳", "92": "Pakistan 🇵🇰", "880": "Bangladesh 🇧🇩",
    "94": "Sri Lanka 🇱🇰", "977": "Nepal 🇳🇵", "975": "Bhutan 🇧🇹",
    "960": "Maldives 🇲🇻", "93": "Afghanistan 🇦🇫",
    "86": "China 🇨🇳", "81": "Japan 🇯🇵", "82": "South Korea 🇰🇷",
    "850": "North Korea 🇰🇵", "886": "Taiwan 🇹🇼", "852": "Hong Kong 🇭🇰",
    "853": "Macau 🇲🇴", "976": "Mongolia 🇲🇳",
    "971": "UAE 🇦🇪", "966": "Saudi Arabia 🇸🇦", "974": "Qatar 🇶🇦",
    "965": "Kuwait 🇰🇼", "968": "Oman 🇴🇲", "973": "Bahrain 🇧🇭",
    "962": "Jordan 🇯🇴", "961": "Lebanon 🇱🇧", "963": "Syria 🇸🇾",
    "964": "Iraq 🇮🇶", "98": "Iran 🇮🇷", "972": "Israel 🇮🇱",
    "970": "Palestine 🇵🇸", "967": "Yemen 🇾🇪", "90": "Turkey 🇹🇷",
    "994": "Azerbaijan 🇦🇿", "995": "Georgia 🇬🇪", "374": "Armenia 🇦🇲",
    "44": "United Kingdom 🇬🇧", "33": "France 🇫🇷", "49": "Germany 🇩🇪",
    "39": "Italy 🇮🇹", "34": "Spain 🇪🇸", "31": "Netherlands 🇳🇱",
    "32": "Belgium 🇧🇪", "41": "Switzerland 🇨🇭", "43": "Austria 🇦🇹",
    "46": "Sweden 🇸🇪", "47": "Norway 🇳🇴", "45": "Denmark 🇩🇰",
    "358": "Finland 🇫🇮", "48": "Poland 🇵🇱", "420": "Czech Republic 🇨🇿",
    "36": "Hungary 🇭🇺", "40": "Romania 🇷🇴", "359": "Bulgaria 🇧🇬",
    "30": "Greece 🇬🇷", "351": "Portugal 🇵🇹", "353": "Ireland 🇮🇪",
    "7": "Russia 🇷🇺", "380": "Ukraine 🇺🇦", "375": "Belarus 🇧🇾",
    "370": "Lithuania 🇱🇹", "371": "Latvia 🇱🇻", "372": "Estonia 🇪🇪",
    "385": "Croatia 🇭🇷", "381": "Serbia 🇷🇸", "387": "Bosnia 🇧🇦",
    "382": "Montenegro 🇲🇪", "389": "North Macedonia 🇲🇰", "383": "Kosovo 🇽🇰",
    "355": "Albania 🇦🇱", "386": "Slovenia 🇸🇮", "421": "Slovakia 🇸🇰",
    "352": "Luxembourg 🇱🇺", "356": "Malta 🇲🇹", "357": "Cyprus 🇨🇾",
    "354": "Iceland 🇮🇸",
    "1": "USA/Canada 🇺🇸", "52": "Mexico 🇲🇽", "55": "Brazil 🇧🇷",
    "54": "Argentina 🇦🇷", "56": "Chile 🇨🇱", "57": "Colombia 🇨🇴",
    "51": "Peru 🇵🇪", "58": "Venezuela 🇻🇪", "593": "Ecuador 🇪🇨",
    "591": "Bolivia 🇧🇴", "595": "Paraguay 🇵🇾", "598": "Uruguay 🇺🇾",
    "507": "Panama 🇵🇦", "506": "Costa Rica 🇨🇷", "503": "El Salvador 🇸🇻",
    "502": "Guatemala 🇬🇹", "504": "Honduras 🇭🇳", "505": "Nicaragua 🇳🇮",
    "53": "Cuba 🇨🇺", "1809": "Dominican Republic 🇩🇴", "1876": "Jamaica 🇯🇲",
    "1868": "Trinidad & Tobago 🇹🇹",
    "20": "Egypt 🇪🇬", "27": "South Africa 🇿🇦", "234": "Nigeria 🇳🇬",
    "254": "Kenya 🇰🇪", "233": "Ghana 🇬🇭", "212": "Morocco 🇲🇦",
    "213": "Algeria 🇩🇿", "216": "Tunisia 🇹🇳", "218": "Libya 🇱🇾",
    "251": "Ethiopia 🇪🇹", "255": "Tanzania 🇹🇿", "256": "Uganda 🇺🇬",
    "250": "Rwanda 🇷🇼", "257": "Burundi 🇧🇮", "260": "Zambia 🇿🇲",
    "263": "Zimbabwe 🇿🇼", "265": "Malawi 🇲🇼", "267": "Botswana 🇧🇼",
    "268": "Eswatini 🇸🇿", "266": "Lesotho 🇱🇸", "258": "Mozambique 🇲🇿",
    "261": "Madagascar 🇲🇬", "230": "Mauritius 🇲🇺", "248": "Seychelles 🇸🇨",
    "221": "Senegal 🇸🇳", "225": "Ivory Coast 🇨🇮", "226": "Burkina Faso 🇧🇫",
    "223": "Mali 🇲🇱", "227": "Niger 🇳🇪", "228": "Togo 🇹🇬",
    "229": "Benin 🇧🇯", "241": "Gabon 🇬🇦", "242": "Congo 🇨🇬",
    "243": "DR Congo 🇨🇩", "237": "Cameroon 🇨🇲", "235": "Chad 🇹🇩",
    "236": "Central African Republic 🇨🇫", "240": "Equatorial Guinea 🇬🇶",
    "239": "Sao Tome 🇸🇹", "244": "Angola 🇦🇴", "249": "Sudan 🇸🇩",
    "211": "South Sudan 🇸🇸", "252": "Somalia 🇸🇴", "253": "Djibouti 🇩🇯",
    "291": "Eritrea 🇪🇷",
    "61": "Australia 🇦🇺", "64": "New Zealand 🇳🇿", "679": "Fiji 🇫🇯",
    "675": "Papua New Guinea 🇵🇬", "677": "Solomon Islands 🇸🇧",
    "678": "Vanuatu 🇻🇺", "685": "Samoa 🇼🇸", "676": "Tonga 🇹🇴",
    "680": "Palau 🇵🇼", "691": "Micronesia 🇫🇲", "692": "Marshall Islands 🇲🇭",
    "674": "Nauru 🇳🇷", "686": "Kiribati 🇰🇮", "688": "Tuvalu 🇹🇻",
}

# =========================================================
# HELPER
# =========================================================
def load_users() -> dict:
    """Return dict {user_id: username}"""
    result = {}
    if not os.path.exists(USERS_FILE):
        return result
    with open(USERS_FILE, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            parts = line.split("|", 1)
            try:
                uid = int(parts[0])
                uname = parts[1] if len(parts) > 1 else "-"
                result[uid] = uname
            except ValueError:
                continue
    return result

def save_user(user_id: int, username: str | None = None):
    users = load_users()
    uname = (username or "-").lstrip("@") or "-"
    # Update if new or username changed
    if user_id not in users or users[user_id] != uname:
        users[user_id] = uname
        with open(USERS_FILE, "w", encoding="utf-8") as f:
            for uid, name in sorted(users.items()):
                f.write(f"{uid}|{name}\n")

def is_force_sub_enabled() -> bool:
    if not os.path.exists(FORCE_SUB_FILE):
        return True
    with open(FORCE_SUB_FILE, "r") as f:
        return f.read().strip().lower() != "off"

def set_force_sub(enabled: bool):
    with open(FORCE_SUB_FILE, "w") as f:
        f.write("on" if enabled else "off")

def load_access() -> dict:
    if not os.path.exists(ACCESS_FILE):
        return {}
    try:
        with open(ACCESS_FILE, "r") as f:
            return json.load(f)
    except Exception:
        return {}

def save_access(data: dict):
    with open(ACCESS_FILE, "w") as f:
        json.dump(data, f, indent=2)

def ensure_user_access(user_id: int) -> dict:
    data = load_access()
    key = str(user_id)
    now = datetime.now()
    if key not in data:
        data[key] = {
            "joined_at": now.isoformat(),
            "access_until": (now + timedelta(days=TRIAL_DAYS)).isoformat(),
            "type": "trial"
        }
        save_access(data)
    return data[key]

def has_access(user_id: int) -> tuple[bool, str]:
    if user_id == ADMIN_ID:
        return True, "Admin ∞"
    info = ensure_user_access(user_id)
    access_until = info.get("access_until")
    if access_until is None:
        return True, "Permanent ✅"
    try:
        until = datetime.fromisoformat(access_until)
    except Exception:
        return False, "Akses tidak valid"
    now = datetime.now()
    if now <= until:
        sisa = (until - now).days
        tipe = info.get("type", "trial")
        if tipe == "trial":
            return True, f"Trial • sisa {sisa} hari"
        return True, f"Aktif • sisa {sisa} hari"
    return False, "Akses berakhir ❌"

def grant_access(user_id: int, days: int | None = None):
    data = load_access()
    key = str(user_id)
    now = datetime.now()
    if key not in data:
        data[key] = {"joined_at": now.isoformat()}
    if days is None:
        data[key]["access_until"] = None
        data[key]["type"] = "permanent"
    else:
        data[key]["access_until"] = (now + timedelta(days=days)).isoformat()
        data[key]["type"] = "granted"
    save_access(data)

def revoke_access(user_id: int):
    data = load_access()
    key = str(user_id)
    now = datetime.now()
    if key not in data:
        data[key] = {"joined_at": now.isoformat()}
    data[key]["access_until"] = (now - timedelta(days=1)).isoformat()
    data[key]["type"] = "revoked"
    save_access(data)

async def is_user_joined(update: Update, context: ContextTypes.DEFAULT_TYPE) -> list:
    user_id = update.effective_user.id
    belum = []
    for gid in REQUIRED_GROUP_IDS:
        try:
            member = await context.bot.get_chat_member(chat_id=gid, user_id=user_id)
            if member.status in ("left", "kicked"):
                belum.append(gid)
        except Exception as e:
            logger.error(f"Gagal cek member {gid}: {e}")
            belum.append(gid)
    return belum

# Pre-sort country codes sekali saja (panjang dulu) biar deteksi lebih cepat
_COUNTRY_CODES_SORTED = sorted(COUNTRY_CODES.keys(), key=len, reverse=True)

# Regex nomor: tangkap deretan digit (dengan +, spasi, strip, titik opsional di antaranya)
_RE_PHONE_CANDIDATE = re.compile(
    r"(?:\+?\d[\d\s\-().]{6,}\d)"
)
# Hanya digit
_RE_DIGITS = re.compile(r"\d+")


def detect_country(nomor: str) -> str:
    for code in _COUNTRY_CODES_SORTED:
        if nomor.startswith(code):
            return COUNTRY_CODES[code]
    return "Tidak diketahui ❓"


def clean_number(text: str) -> str:
    """Ambil digit saja dari teks. Handle float Excel (1.234e10) juga."""
    if text is None:
        return ""
    s = str(text).strip()
    if not s:
        return ""

    # Excel kadang simpan nomor panjang sebagai scientific notation
    if "e" in s.lower() or "E" in s:
        try:
            # contoh: 6.28123456789e10 → 628123456789
            val = float(s.replace(",", "."))
            if val == int(val) and val > 1e7:
                s = str(int(val))
        except (ValueError, OverflowError):
            pass

    return "".join(ch for ch in s if ch.isdigit())


def normalize_number(n: str) -> str:
    """
    Normalisasi nomor:
    - buang leading 00 → ganti jadi tidak perlu
    - 08xxxx (Indonesia lokal) biarkan; user bisa ganti prefix lewat edit
    - minimal 8 digit, maksimal 15 (standar E.164)
    """
    if not n:
        return ""
    # Buang leading zeros berlebih tapi jangan hapus semua
    # Contoh: 0062... → 62...
    if n.startswith("00") and len(n) > 10:
        n = n[2:]
    if len(n) < 8 or len(n) > 15:
        return ""
    return n


def extract_numbers_from_text(text: str) -> list[str]:
    """Ekstrak semua nomor dari teks bebas (bisa banyak per baris)."""
    result = []
    if not text:
        return result

    # Coba kandidat mirip nomor telepon dulu
    for match in _RE_PHONE_CANDIDATE.finditer(text):
        n = normalize_number(clean_number(match.group(0)))
        if n:
            result.append(n)

    # Fallback: tiap baris digabung digitnya (untuk format 1 nomor per baris)
    if not result:
        for line in text.splitlines():
            n = normalize_number(clean_number(line))
            if n:
                result.append(n)

    return result


def extract_numbers_from_file(path: str, ext: str) -> list[str]:
    """Ekstrak nomor dari txt / xlsx / csv dengan parsing lebih tangguh."""
    numbers: list[str] = []

    if ext == "txt":
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            content = f.read()
        numbers = extract_numbers_from_text(content)

    elif ext == "xlsx":
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        sheet = wb.active
        for row in sheet.iter_rows(values_only=True):
            for cell in row:
                if cell is None:
                    continue
                # Angka murni dari Excel
                if isinstance(cell, (int, float)):
                    try:
                        n = normalize_number(str(int(cell)))
                        if n:
                            numbers.append(n)
                            continue
                    except (ValueError, OverflowError):
                        pass
                # Teks / campuran
                for n in extract_numbers_from_text(str(cell)):
                    numbers.append(n)
        wb.close()

    elif ext == "csv":
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            # Coba deteksi delimiter sederhana
            sample = f.read(4096)
            f.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
            except csv.Error:
                dialect = csv.excel
            reader = csv.reader(f, dialect)
            for row in reader:
                for cell in row:
                    if not cell:
                        continue
                    for n in extract_numbers_from_text(str(cell)):
                        numbers.append(n)

    return numbers


def parse_vcards(content: str) -> list[str]:
    """Parse blok VCARD; lebih toleran terhadap spasi/baris."""
    return re.findall(
        r"BEGIN\s*:\s*VCARD.*?END\s*:\s*VCARD",
        content,
        re.DOTALL | re.IGNORECASE,
    )


def get_fn(vcard: str) -> str:
    m = re.search(r"FN\s*:\s*(.*)", vcard, re.IGNORECASE)
    if not m:
        return "Unknown"
    # Hilangkan parameter vCard aneh di akhir baris
    return m.group(1).split("\n")[0].strip()


def get_tel(vcard: str) -> str:
    """Ambil nomor dari baris TEL (dukung beberapa format vCard)."""
    # TEL;TYPE=CELL:+62...  /  TEL;CELL:62...  /  TEL:+62...
    matches = re.findall(r"TEL(?:;[^:]*)?:(.*)", vcard, re.IGNORECASE)
    for raw in matches:
        n = normalize_number(clean_number(raw))
        if n:
            return n
    return ""

def safe_filename(name: str) -> str:
    return re.sub(r'[\\/*?:"<>|]', "", name).strip()

# =========================================================
# CEK AKSES
# =========================================================
async def require_access(update: Update, context: ContextTypes.DEFAULT_TYPE) -> bool:
    uid = update.effective_user.id
    ok, status = has_access(uid)
    if ok:
        return True
    msg = (
        f"🔒 **Akses Berakhir**\n\n"
        f"Masa trial / akses kamu sudah habis.\n"
        f"Hubungi admin untuk perpanjang.\n\n"
        f"Status: `{status}`"
    )
    if update.message:
        await update.message.reply_text(msg, parse_mode="Markdown")
    elif update.callback_query:
        await update.callback_query.answer("Akses berakhir", show_alert=True)
        await update.callback_query.edit_message_text(msg, parse_mode="Markdown")
    return False

# =========================================================
# MENU UTAMA (/start)
# =========================================================
async def show_menu(update: Update, context: ContextTypes.DEFAULT_TYPE, edit: bool = False):
    user = update.effective_user
    if not user and update.callback_query:
        user = update.callback_query.from_user
    uid = user.id if user else 0

    if uid:
        ensure_user_access(uid)
        _, status = has_access(uid)
    else:
        status = "-"

    nama = user.first_name if user else "User"
    text = (
        f"✨ **CV Kurumi Premium**\n"
        f"─────────────────\n"
        f"Halo, **{nama}**!\n"
        f"Status akses: `{status}`\n"
        f"─────────────────\n"
        f"Pilih menu di bawah untuk mulai.\n\n"
        f"mulai bot → `/start`\n"
        f"konversi file ke .vcf → `/to_vcf`\n"
        f"konversi file ke .txt → `/to_txt`\n"
        f"fitur adminnavy → `/admin`\n"
        f"input kontak manual → `/manual`\n"
        f"tambah kontak ke .vcf → `/add`\n"
        f"hapus kontak dari file → `/delete`\n"
        f"ganti nama kontak → `/renamectc`\n"
        f"ganti nama file → `/renamefile`\n"
        f"gabungkan file → `/merge`\n"
        f"pecah file → `/split`\n"
        f"hitung jumlah kontak → `/count`\n"
        f"hapus kontak duplikat → `/nodup`\n"
        f"extract nama file → `/getname`\n"
        f"generate nama file → `/generate`\n"
        f"ekstrak isi file .txt → `/getconten`\n"
        f"menu pengaturan → `/setting`\n"
        f"cek status akun → `/status`"
    )

    keyboard = []
    if user and user.id == ADMIN_ID:
        fs = "ON ✅" if is_force_sub_enabled() else "OFF ❌"
        keyboard = [
            [InlineKeyboardButton(f"⚙️ Force Sub: {fs}", callback_data="admin_toggle_fs")],
            [InlineKeyboardButton("👥 Cek User Bot", callback_data="admin_cek_user")],
            [InlineKeyboardButton("📢 Broadcast", callback_data="admin_broadcast")],
        ]
    markup = InlineKeyboardMarkup(keyboard) if keyboard else None

    if edit and update.callback_query:
        await update.callback_query.edit_message_text(text, reply_markup=markup, parse_mode="Markdown")
    elif update.message:
        await update.message.reply_text(text, reply_markup=markup, parse_mode="Markdown")
    elif update.callback_query:
        await update.callback_query.edit_message_text(text, reply_markup=markup, parse_mode="Markdown")
    return MENU

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    user = update.effective_user
    if user:
        save_user(user.id, user.username)
        ensure_user_access(user.id)

    if is_force_sub_enabled():
        belum = await is_user_joined(update, context)
        if belum:
            kb = [[InlineKeyboardButton(f"Join Grup Wajib {i+1}", url=link)] for i, link in enumerate(INVITE_LINKS)]
            msg = (
                "⚠️ **Wajib Join Dulu**\n\n"
                "Sebelum memakai **CV Kurumi Premium**, kamu harus join grup/channel di bawah.\n"
                "Setelah join, kirim /start lagi."
            )
            if update.message:
                await update.message.reply_text(msg, reply_markup=InlineKeyboardMarkup(kb), parse_mode="Markdown")
            else:
                await update.callback_query.edit_message_text(msg, reply_markup=InlineKeyboardMarkup(kb), parse_mode="Markdown")
            return MENU

    return await show_menu(update, context, edit=bool(update.callback_query))

# =========================================================
# COMMAND HANDLERS
# =========================================================
async def cmd_to_vcf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await require_access(update, context):
        return MENU
    await update.message.reply_text(
        "📄 **Konversi ke VCF** (`/to_vcf`)\n\n"
        "Kirim file `.txt` / `.xlsx` / `.csv` yang berisi nomor."
    )
    return TO_VCF_FILE

async def cmd_to_txt(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await require_access(update, context):
        return MENU
    await update.message.reply_text(
        "📇 **Konversi ke TXT** (`/to_txt`)\n\n"
        "Kirim file `.vcf` / `.xlsx` / `.csv`."
    )
    return TO_TXT_FILE

async def cmd_admin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await require_access(update, context):
        return MENU
    context.user_data["an_groups"] = []
    context.user_data["an_current_numbers"] = []
    await update.message.reply_text(
        "👮‍♂️ **Admin & Navy VCF** (`/admin`)\n\n"
        "Ketik **Nama Grup** pertama (contoh: `Admin Gacor`):"
    )
    return ADMIN_NAMA_GRUP

async def cmd_manual(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await require_access(update, context):
        return MENU
    context.user_data["manual_data"] = []
    await update.message.reply_text(
        "✍️ **Input Manual**\n\n"
        "Kirim nomor (bisa banyak sekaligus).\n"
        "Contoh: `628111111111` atau `628111 628222 628333`"
    )
    return MANUAL_NOMOR

async def cmd_add(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await require_access(update, context):
        return MENU
    await update.message.reply_text(
        "➕ **Tambah Kontak ke VCF** (`/add`)\n\n"
        "Kirim file `.vcf` yang ingin ditambahkan kontaknya."
    )
    return ADD_FILE

async def cmd_delete(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await require_access(update, context):
        return MENU
    await update.message.reply_text(
        "🗑️ **Hapus Kontak** (`/delete`)\n\n"
        "Kirim file `.vcf` atau `.txt`."
    )
    return DELETE_FILE

async def cmd_renamectc(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await require_access(update, context):
        return MENU
    await update.message.reply_text(
        "✏️ **Ganti Nama Kontak** (`/renamectc`)\n\n"
        "Kirim file `.vcf`."
    )
    return RENAME_CTC_FILE

async def cmd_renamefile(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await require_access(update, context):
        return MENU
    await update.message.reply_text(
        "📝 **Ganti Nama File** (`/renamefile`)\n\n"
        "Kirim file yang ingin diganti namanya."
    )
    return RENAME_FILE

async def cmd_merge(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await require_access(update, context):
        return MENU
    context.user_data["merge_files"] = []
    await update.message.reply_text(
        "🔗 **Gabungkan VCF** (`/merge`)\n\n"
        "Kirim file `.vcf` pertama.\n"
        "Kirim sebanyak yang dibutuhkan, lalu ketik `selesai`."
    )
    return MERGE_FILE

async def cmd_split(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await require_access(update, context):
        return MENU
    await update.message.reply_text(
        "✂️ **Pecah File** (`/split`)\n\n"
        "Kirim file `.vcf` atau `.txt`."
    )
    return SPLIT_FILE

async def cmd_count(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await require_access(update, context):
        return MENU
    await update.message.reply_text(
        "🔢 **Hitung Jumlah Kontak** (`/count`)\n\n"
        "Kirim file `.vcf` / `.txt` / `.xlsx` / `.csv`."
    )
    return COUNT_FILE

async def cmd_nodup(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await require_access(update, context):
        return MENU
    await update.message.reply_text(
        "🧹 **Hapus Duplikat** (`/nodup`)\n\n"
        "Kirim file `.vcf` atau `.txt`."
    )
    return NODUP_FILE

async def cmd_getname(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await require_access(update, context):
        return MENU
    await update.message.reply_text(
        "📛 **Extract Nama Kontak** (`/getname`)\n\n"
        "Kirim file `.vcf`."
    )
    return GETNAME_FILE

async def cmd_generate(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await require_access(update, context):
        return MENU
    await update.message.reply_text(
        "⚙️ **Generate Nama** (`/generate`)\n\n"
        "Ketik pola nama (contoh: `Kontak_shopee` atau `Admin`):"
    )
    return GENERATE_NAMA

async def cmd_getconten(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await require_access(update, context):
        return MENU
    await update.message.reply_text(
        "📄 **Ekstrak Isi TXT** (`/getconten`)\n\n"
        "Kirim file `.txt`."
    )
    return GETCONTEN_FILE

async def cmd_setting(update: Update, context: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    _, status = has_access(uid)
    fs = "ON ✅" if is_force_sub_enabled() else "OFF ❌"
    users = load_users()
    total_user = len(users)
    text = (
        f"⚙️ **Menu Pengaturan** (`/setting`)\n\n"
        f"• Status akses : `{status}`\n"
        f"• Force Subscribe : {fs}\n"
        f"• Trial default : `{TRIAL_DAYS}` hari\n"
        f"• Total user bot : `{total_user}`\n\n"
    )
    if uid == ADMIN_ID:
        text += (
            "**Admin commands:**\n"
            "`/grant <user_id> [hari]` — beri akses\n"
            "`/revoke <user_id>` — cabut akses\n"
            "`/cekakses <user_id>` — cek akses user\n"
            "`/cekuser` — daftar semua user (ID + username)\n"
            "`/broadcast` — kirim pesan ke semua user"
        )
    else:
        text += "Hubungi admin untuk perpanjang akses."
    await update.message.reply_text(text, parse_mode="Markdown")
    return MENU

async def cmd_status(update: Update, context: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    _, status = has_access(uid)
    info = load_access().get(str(uid), {})
    tipe = "Admin" if uid == ADMIN_ID else info.get("type", "-")
    until = "Permanent ∞" if uid == ADMIN_ID else (info.get("access_until") or "Permanent")
    text = (
        f"📊 **Status Akun** (`/status`)\n\n"
        f"• User ID : `{uid}`\n"
        f"• Status  : `{status}`\n"
        f"• Tipe    : `{tipe}`\n"
        f"• Join    : `{str(info.get('joined_at', '-'))[:10]}`\n"
        f"• Until   : `{until}`"
    )
    await update.message.reply_text(text, parse_mode="Markdown")
    return MENU

async def cmd_cekuser(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Admin only: tampilkan daftar user ID + username"""
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("⛔ Hanya admin.")
        return MENU

    users = load_users()  # {id: username}
    total = len(users)
    fs = "Aktif ✅" if is_force_sub_enabled() else "Nonaktif ❌"

    await update.message.reply_text(
        f"👥 **Statistik User Bot**\n\n"
        f"• Total user : `{total}`\n"
        f"• Force Sub  : {fs}",
        parse_mode="Markdown"
    )

    if total == 0:
        await update.message.reply_text("Belum ada user yang tersimpan.")
        return MENU

    # Buat file daftar
    lines = ["user_id | username"]
    for uid, uname in sorted(users.items()):
        lines.append(f"{uid} | @{uname}" if uname != "-" else f"{uid} | -")
    list_file = "daftar_user.txt"
    with open(list_file, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    await context.bot.send_document(
        chat_id=update.effective_chat.id,
        document=open(list_file, "rb"),
        caption=f"📄 Daftar {total} User (ID + Username)"
    )
    os.remove(list_file)
    return MENU

async def cmd_broadcast_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Admin only: mulai broadcast lewat perintah /broadcast"""
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("⛔ Hanya admin.")
        return MENU
    await update.message.reply_text(
        "📢 **Broadcast**\n\n"
        "Kirim pesan yang ingin disiarkan ke **semua user** bot ini.\n"
        "Ketik /cancel untuk batal."
    )
    return BROADCAST_MESSAGE

# =========================================================
# HANDLER FILE & STATE
# =========================================================
async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc or not doc.file_name:
        await update.message.reply_text("❌ File tidak valid.")
        return MENU

    ext = doc.file_name.rsplit(".", 1)[-1].lower()
    context.user_data["file_id"] = doc.file_id
    context.user_data["file_name"] = doc.file_name
    context.user_data["file_ext"] = ext

    state = context.user_data.get("_state")  # kita set di masing-masing command

    # --- TO_VCF ---
    if context.user_data.get("mode") == "to_vcf":
        if ext not in ("txt", "xlsx", "csv"):
            await update.message.reply_text("❌ Harus file .txt / .xlsx / .csv")
            return TO_VCF_FILE
        await update.message.reply_text("Ketik **Nama File** (contoh: `Kontak_shopee`):")
        return TO_VCF_NAMA_FILE

    # --- TO_TXT ---
    if context.user_data.get("mode") == "to_txt":
        if ext not in ("vcf", "xlsx", "csv"):
            await update.message.reply_text("❌ Harus file .vcf / .xlsx / .csv")
            return TO_TXT_FILE
        await proses_to_txt(update, context)
        return MENU

    # --- ADD ---
    if context.user_data.get("mode") == "add":
        if ext != "vcf":
            await update.message.reply_text("❌ Harus file .vcf")
            return ADD_FILE
        await update.message.reply_text(
            "Kirim nomor yang ingin ditambahkan (bisa banyak, pisah spasi/baris):"
        )
        return ADD_NOMOR

    # --- DELETE ---
    if context.user_data.get("mode") == "delete":
        if ext not in ("vcf", "txt"):
            await update.message.reply_text("❌ Harus file .vcf atau .txt")
            return DELETE_FILE
        await update.message.reply_text(
            "Kirim nomor yang ingin **dihapus** (bisa banyak):"
        )
        return DELETE_NOMOR

    # --- RENAME CTC ---
    if context.user_data.get("mode") == "renamectc":
        if ext != "vcf":
            await update.message.reply_text("❌ Harus file .vcf")
            return RENAME_CTC_FILE
        await update.message.reply_text(
            "Ketik **Nama Kontak Baru** (akan jadi Nama 001, Nama 002, ...):"
        )
        return RENAME_CTC_NAMA

    # --- RENAME FILE ---
    if context.user_data.get("mode") == "renamefile":
        await update.message.reply_text(
            "Ketik **Nama File Baru** (tanpa ekstensi):\nContoh: `Hasil_Final`"
        )
        return RENAME_FILE_NEW

    # --- MERGE ---
    if context.user_data.get("mode") == "merge":
        if ext != "vcf":
            await update.message.reply_text("❌ Harus file .vcf")
            return MERGE_FILE
        context.user_data.setdefault("merge_files", []).append(doc.file_id)
        await update.message.reply_text(
            f"✅ File ke-{len(context.user_data['merge_files'])} ditambahkan.\n"
            "Kirim lagi atau ketik `selesai`."
        )
        return MERGE_FILE

    # --- SPLIT ---
    if context.user_data.get("mode") == "split":
        if ext not in ("vcf", "txt"):
            await update.message.reply_text("❌ Harus file .vcf atau .txt")
            return SPLIT_FILE
        await update.message.reply_text("Ketik **jumlah per file** (contoh: `500`):")
        return SPLIT_LIMIT

    # --- COUNT ---
    if context.user_data.get("mode") == "count":
        if ext not in ("vcf", "txt", "xlsx", "csv"):
            await update.message.reply_text("❌ Format tidak didukung")
            return COUNT_FILE
        await proses_count(update, context)
        return MENU

    # --- NODUP ---
    if context.user_data.get("mode") == "nodup":
        if ext not in ("vcf", "txt"):
            await update.message.reply_text("❌ Harus file .vcf atau .txt")
            return NODUP_FILE
        await proses_nodup(update, context)
        return MENU

    # --- GETNAME ---
    if context.user_data.get("mode") == "getname":
        if ext != "vcf":
            await update.message.reply_text("❌ Harus file .vcf")
            return GETNAME_FILE
        await proses_getname(update, context)
        return MENU

    # --- GETCONTEN ---
    if context.user_data.get("mode") == "getconten":
        if ext != "txt":
            await update.message.reply_text("❌ Harus file .txt")
            return GETCONTEN_FILE
        await proses_getconten(update, context)
        return MENU

    await update.message.reply_text("❓ Tidak ada proses aktif. Ketik /start")
    return MENU

# =========================================================
# PROSES FUNGSI
# =========================================================
async def to_vcf_nama_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["nama_file"] = safe_filename(update.message.text)
    await update.message.reply_text("Ketik **Nama Kontak** (contoh: `shopee`):")
    return TO_VCF_NAMA_KONTAK

async def to_vcf_nama_kontak(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["nama_kontak"] = update.message.text.strip()
    await update.message.reply_text("Mau dipecah berapa nomor per file? (contoh: `500`):")
    return TO_VCF_LIMIT

async def to_vcf_limit(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        limit = int(update.message.text.strip())
        if limit < 1:
            raise ValueError
        context.user_data["limit"] = limit
    except ValueError:
        await update.message.reply_text("❌ Harus angka positif.")
        return TO_VCF_LIMIT

    await update.message.reply_text("⏳ Sedang memproses...")
    await proses_to_vcf(update, context)
    return MENU

async def proses_to_vcf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    file_id = context.user_data["file_id"]
    file_name = context.user_data["file_name"]
    ext = context.user_data["file_ext"]
    nama_file = context.user_data["nama_file"]
    nama_kontak = context.user_data["nama_kontak"]
    limit = context.user_data["limit"]

    file = await context.bot.get_file(file_id)
    path = f"temp_{file_id}.{ext}"
    await file.download_to_drive(path)

    numbers = extract_numbers_from_file(path, ext)
    os.remove(path)

    if not numbers:
        await update.message.reply_text("❌ Tidak ditemukan nomor valid.")
        context.user_data.clear()
        return

    total_files = math.ceil(len(numbers) / limit)
    total_kontak = len(numbers)

    # Deteksi negara dari nomor unik
    country_count = Counter()
    for n in set(numbers):
        country_count[detect_country(n)] += 1

    for i in range(total_files):
        chunk = numbers[i * limit : (i + 1) * limit]
        out_name = f"{nama_file} {i + 1}.vcf"
        with open(out_name, "w", encoding="utf-8") as f:
            for j, n in enumerate(chunk):
                urutan = i * limit + j + 1
                f.write(
                    f"BEGIN:VCARD\nVERSION:3.0\n"
                    f"FN:{nama_kontak} {urutan:03d}\n"
                    f"TEL;TYPE=CELL:+{n}\nEND:VCARD\n"
                )
        # Kirim file TANPA caption di bawahnya
        await context.bot.send_document(
            chat_id=update.effective_chat.id,
            document=open(out_name, "rb"),
        )
        os.remove(out_name)

    # Ringkasan singkat setelah semua file terkirim
    negara_lines = []
    for negara, jml in country_count.most_common(8):
        negara_lines.append(f"• {negara}: {jml}")
    negara_text = "\n".join(negara_lines) if negara_lines else "• Tidak terdeteksi"

    ringkasan = (
        f"✅ **Selesai**\n"
        f"📁 `{nama_file}`\n"
        f"📄 {total_files} file • {total_kontak} kontak\n\n"
        f"🌍 **Negara:**\n{negara_text}\n\n"
        f"Ketik /start untuk menu lagi."
    )
    await update.message.reply_text(ringkasan, parse_mode="Markdown")
    context.user_data.clear()

async def proses_to_txt(update: Update, context: ContextTypes.DEFAULT_TYPE):
    file_id = context.user_data["file_id"]
    file_name = context.user_data["file_name"]
    ext = context.user_data["file_ext"]

    file = await context.bot.get_file(file_id)
    path = f"temp_{file_id}.{ext}"
    await file.download_to_drive(path)

    if ext == "vcf":
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            content = f.read()
        numbers = [get_tel(v) for v in parse_vcards(content) if get_tel(v)]
    else:
        numbers = extract_numbers_from_file(path, ext)

    os.remove(path)
    numbers = list(dict.fromkeys(numbers))  # unique preserve order

    out_name = "Hasil_Ekstrak.txt"
    with open(out_name, "w") as f:
        f.write("\n".join(numbers))

    await context.bot.send_document(
        chat_id=update.effective_chat.id,
        document=open(out_name, "rb"),
        caption=f"✅ {len(numbers)} nomor unik diekstrak"
    )
    os.remove(out_name)
    await update.message.reply_text("🎉 Selesai! Ketik /start.")
    context.user_data.clear()

# --- MANUAL ---
# Alur: nomor → langsung nama kontak → nama file → VCF (tanpa "selesai")

async def manual_nomor(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    # Pakai parser optimasi
    baru = extract_numbers_from_text(text)
    if not baru:
        # fallback sederhana
        baru = [
            normalize_number(clean_number(p))
            for p in re.split(r"[\s,;]+", text)
            if normalize_number(clean_number(p))
        ]
    if not baru:
        await update.message.reply_text("❌ Tidak ada nomor valid. Kirim ulang nomor.")
        return MANUAL_NOMOR

    context.user_data["manual_data"] = baru
    await update.message.reply_text(
        f"✅ {len(baru)} nomor tersimpan.\n\n"
        "Ketik **Nama Kontak** (contoh: `shopee`):",
        parse_mode="Markdown"
    )
    return MANUAL_NAMA


async def manual_nama(update: Update, context: ContextTypes.DEFAULT_TYPE):
    nama = update.message.text.strip()
    if not nama:
        await update.message.reply_text("❌ Nama kontak tidak boleh kosong.")
        return MANUAL_NAMA
    context.user_data["temp_nama"] = nama
    await update.message.reply_text(
        "Ketik **Nama File** (tanpa .vcf)\n"
        "Contoh: `Kontak_shopee`"
    )
    return MANUAL_NAMA_FILE


async def manual_nama_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    nama_file = safe_filename(update.message.text.strip())
    if not nama_file:
        await update.message.reply_text("❌ Nama file tidak boleh kosong.")
        return MANUAL_NAMA_FILE

    numbers = context.user_data.get("manual_data", [])
    nama = context.user_data.get("temp_nama", "Kontak")
    if not numbers:
        await update.message.reply_text("❌ Data nomor hilang. Ketik /manual ulang.")
        context.user_data.clear()
        return MENU

    if not nama_file.lower().endswith(".vcf"):
        nama_file += ".vcf"

    await update.message.reply_text("⏳ Membuat VCF...")

    with open(nama_file, "w", encoding="utf-8") as f:
        for i, n in enumerate(numbers, 1):
            f.write(
                f"BEGIN:VCARD\nVERSION:3.0\n"
                f"FN:{nama} {i:03d}\nTEL;TYPE=CELL:+{n}\nEND:VCARD\n"
            )

    # Deteksi negara singkat
    country_count = Counter(detect_country(n) for n in set(numbers))
    negara = ", ".join(f"{k} ({v})" for k, v in country_count.most_common(3))

    await context.bot.send_document(
        chat_id=update.effective_chat.id,
        document=open(nama_file, "rb"),
    )
    os.remove(nama_file)

    await update.message.reply_text(
        f"✅ Selesai\n"
        f"📁 `{nama_file}`\n"
        f"👥 {len(numbers)} kontak\n"
        f"🌍 {negara}\n\n"
        f"Ketik /start untuk menu lagi.",
        parse_mode="Markdown"
    )
    context.user_data.clear()
    return MENU

# --- ADD ---
async def add_nomor(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    baru = [clean_number(p) for p in re.split(r"[\s,;]+", text) if len(clean_number(p)) >= 8]
    if not baru:
        await update.message.reply_text("❌ Tidak ada nomor valid.")
        return ADD_NOMOR

    file_id = context.user_data["file_id"]
    file_name = context.user_data["file_name"]
    file = await context.bot.get_file(file_id)
    path = f"temp_add_{file_id}.vcf"
    await file.download_to_drive(path)

    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        content = f.read()
    existing = parse_vcards(content)
    base = "Kontak"
    if existing:
        fn = get_fn(existing[0])
        base = re.sub(r"\s*\d+$", "", fn).strip() or "Kontak"

    start = len(existing) + 1
    with open(path, "w", encoding="utf-8") as f:
        f.write(content.rstrip() + "\n")
        for i, n in enumerate(baru):
            f.write(
                f"BEGIN:VCARD\nVERSION:3.0\n"
                f"FN:{base} {start + i:03d}\nTEL;TYPE=CELL:+{n}\nEND:VCARD\n"
            )

    out = safe_filename(file_name.replace(".vcf", "") + "_added.vcf")
    os.rename(path, out)
    await context.bot.send_document(
        chat_id=update.effective_chat.id,
        document=open(out, "rb"),
        caption=f"✅ Ditambah {len(baru)} kontak. Total sekarang: {len(existing) + len(baru)}"
    )
    os.remove(out)
    await update.message.reply_text("🎉 Selesai! Ketik /start.")
    context.user_data.clear()
    return MENU

# --- DELETE ---
async def delete_nomor(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    to_del = {clean_number(p) for p in re.split(r"[\s,;]+", text) if len(clean_number(p)) >= 8}
    if not to_del:
        await update.message.reply_text("❌ Tidak ada nomor valid.")
        return DELETE_NOMOR

    file_id = context.user_data["file_id"]
    file_name = context.user_data["file_name"]
    ext = context.user_data["file_ext"]
    file = await context.bot.get_file(file_id)
    path = f"temp_del_{file_id}.{ext}"
    await file.download_to_drive(path)

    if ext == "txt":
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            lines = f.readlines()
        kept = [l for l in lines if clean_number(l) not in to_del]
        out = safe_filename(file_name.replace(".txt", "") + "_deleted.txt")
        with open(out, "w") as f:
            f.writelines(kept)
        await context.bot.send_document(
            chat_id=update.effective_chat.id,
            document=open(out, "rb"),
            caption=f"✅ Dihapus {len(lines) - len(kept)} nomor. Sisa {len(kept)}"
        )
        os.remove(out)
    else:
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            content = f.read()
        vcards = parse_vcards(content)
        kept = [v for v in vcards if get_tel(v) not in to_del]
        out = safe_filename(file_name.replace(".vcf", "") + "_deleted.vcf")
        with open(out, "w", encoding="utf-8") as f:
            f.write("\n".join(kept))
        await context.bot.send_document(
            chat_id=update.effective_chat.id,
            document=open(out, "rb"),
            caption=f"✅ Dihapus {len(vcards) - len(kept)} kontak. Sisa {len(kept)}"
        )
        os.remove(out)

    os.remove(path)
    await update.message.reply_text("🎉 Selesai! Ketik /start.")
    context.user_data.clear()
    return MENU

# --- RENAME CTC ---
async def rename_ctc_nama(update: Update, context: ContextTypes.DEFAULT_TYPE):
    nama = update.message.text.strip()
    if not nama:
        await update.message.reply_text("❌ Nama tidak boleh kosong.")
        return RENAME_CTC_NAMA

    file_id = context.user_data["file_id"]
    file_name = context.user_data["file_name"]
    file = await context.bot.get_file(file_id)
    path = f"temp_ren_{file_id}.vcf"
    await file.download_to_drive(path)

    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        content = f.read()
    os.remove(path)

    vcards = parse_vcards(content)
    new_vcards = []
    for i, v in enumerate(vcards, 1):
        tel = get_tel(v)
        new_vcards.append(
            f"BEGIN:VCARD\nVERSION:3.0\nFN:{nama} {i:03d}\nTEL;TYPE=CELL:+{tel}\nEND:VCARD"
        )

    out = safe_filename(file_name.replace(".vcf", "") + "_renamed.vcf")
    with open(out, "w", encoding="utf-8") as f:
        f.write("\n".join(new_vcards))
    await context.bot.send_document(
        chat_id=update.effective_chat.id,
        document=open(out, "rb"),
        caption=f"✅ Semua nama diganti menjadi `{nama} 001` dst.\nTotal: {len(new_vcards)}",
        parse_mode="Markdown"
    )
    os.remove(out)
    await update.message.reply_text("🎉 Selesai! Ketik /start.")
    context.user_data.clear()
    return MENU

# --- RENAME FILE ---
async def rename_file_new(update: Update, context: ContextTypes.DEFAULT_TYPE):
    nama = safe_filename(update.message.text.strip())
    if not nama:
        await update.message.reply_text("❌ Nama tidak boleh kosong.")
        return RENAME_FILE_NEW

    file_id = context.user_data["file_id"]
    file_name = context.user_data["file_name"]
    ext = context.user_data["file_ext"]
    if not nama.lower().endswith(f".{ext}"):
        nama += f".{ext}"

    file = await context.bot.get_file(file_id)
    await file.download_to_drive(nama)
    await context.bot.send_document(
        chat_id=update.effective_chat.id,
        document=open(nama, "rb"),
        caption=f"✅ File diganti nama menjadi `{nama}`",
        parse_mode="Markdown"
    )
    os.remove(nama)
    await update.message.reply_text("🎉 Selesai! Ketik /start.")
    context.user_data.clear()
    return MENU

# --- MERGE ---
async def merge_selesai(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.message.text.strip().lower() != "selesai":
        await update.message.reply_text("Ketik `selesai` jika sudah selesai mengirim file.")
        return MERGE_FILE

    files = context.user_data.get("merge_files", [])
    if len(files) < 2:
        await update.message.reply_text("❌ Minimal 2 file VCF.")
        return MERGE_FILE

    await update.message.reply_text("⏳ Menggabungkan...")
    out = "Merged.vcf"
    with open(out, "w", encoding="utf-8") as outf:
        for fid in files:
            f = await context.bot.get_file(fid)
            p = f"temp_{fid}.vcf"
            await f.download_to_drive(p)
            with open(p, "r", encoding="utf-8", errors="ignore") as inf:
                outf.write(inf.read().rstrip() + "\n")
            os.remove(p)

    await context.bot.send_document(
        chat_id=update.effective_chat.id,
        document=open(out, "rb"),
        caption="✅ File VCF berhasil digabung!"
    )
    os.remove(out)
    await update.message.reply_text("🎉 Selesai! Ketik /start.")
    context.user_data.clear()
    return MENU

# --- SPLIT ---
async def split_limit(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        limit = int(update.message.text.strip())
        if limit < 1:
            raise ValueError
    except ValueError:
        await update.message.reply_text("❌ Harus angka positif.")
        return SPLIT_LIMIT

    file_id = context.user_data["file_id"]
    file_name = context.user_data["file_name"]
    ext = context.user_data["file_ext"]

    file = await context.bot.get_file(file_id)
    path = f"temp_split_{file_id}.{ext}"
    await file.download_to_drive(path)

    if ext == "vcf":
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            content = f.read()
        items = parse_vcards(content)
        total = math.ceil(len(items) / limit)
        await update.message.reply_text(f"⏳ Memecah {len(items)} kontak menjadi {total} file...")
        for i in range(total):
            chunk = items[i * limit : (i + 1) * limit]
            out = f"Split {i + 1}.vcf"
            with open(out, "w", encoding="utf-8") as f:
                f.write("\n".join(chunk))
            await context.bot.send_document(
                chat_id=update.effective_chat.id,
                document=open(out, "rb"),
                caption=f"Split {i + 1} ({len(chunk)} kontak)"
            )
            os.remove(out)
    else:
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            lines = [l.strip() for l in f if l.strip()]
        total = math.ceil(len(lines) / limit)
        await update.message.reply_text(f"⏳ Memecah {len(lines)} baris menjadi {total} file...")
        base = file_name.rsplit(".", 1)[0]
        for i in range(total):
            chunk = lines[i * limit : (i + 1) * limit]
            out = f"{base} {i + 1}.txt"
            with open(out, "w") as f:
                f.write("\n".join(chunk))
            await context.bot.send_document(
                chat_id=update.effective_chat.id,
                document=open(out, "rb"),
                caption=f"{base} {i + 1} ({len(chunk)} baris)"
            )
            os.remove(out)

    os.remove(path)
    await update.message.reply_text("🎉 Selesai! Ketik /start.")
    context.user_data.clear()
    return MENU

# --- COUNT ---
async def proses_count(update: Update, context: ContextTypes.DEFAULT_TYPE):
    file_id = context.user_data["file_id"]
    file_name = context.user_data["file_name"]
    ext = context.user_data["file_ext"]

    file = await context.bot.get_file(file_id)
    path = f"temp_count_{file_id}.{ext}"
    await file.download_to_drive(path)

    if ext == "vcf":
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            content = f.read()
        vcards = parse_vcards(content)
        total = len(vcards)
        numbers = [get_tel(v) for v in vcards if get_tel(v)]
    else:
        numbers = extract_numbers_from_file(path, ext)
        total = len(numbers)

    os.remove(path)
    unik = len(set(numbers))
    dup = total - unik

    country = Counter(detect_country(n) for n in set(numbers))
    text = (
        f"🔢 **Hasil Hitung**\n📁 `{file_name}`\n\n"
        f"• Total    : `{total}`\n"
        f"• Unik     : `{unik}`\n"
        f"• Duplikat : `{dup}`\n\n"
        f"🌍 **Negara (unik):**\n"
    )
    for negara, jml in country.most_common(12):
        text += f"• {negara}: `{jml}`\n"

    await update.message.reply_text(text, parse_mode="Markdown")
    context.user_data.clear()

# --- NODUP ---
async def proses_nodup(update: Update, context: ContextTypes.DEFAULT_TYPE):
    file_id = context.user_data["file_id"]
    file_name = context.user_data["file_name"]
    ext = context.user_data["file_ext"]

    file = await context.bot.get_file(file_id)
    path = f"temp_nodup_{file_id}.{ext}"
    await file.download_to_drive(path)

    if ext == "txt":
        numbers = extract_numbers_from_file(path, "txt")
        unik = list(dict.fromkeys(numbers))
        out = safe_filename(file_name.replace(".txt", "") + "_nodup.txt")
        with open(out, "w") as f:
            f.write("\n".join(unik))
        await context.bot.send_document(
            chat_id=update.effective_chat.id,
            document=open(out, "rb"),
            caption=f"✅ {len(unik)} nomor unik (dari {len(numbers)})"
        )
        os.remove(out)
    else:
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            content = f.read()
        vcards = parse_vcards(content)
        seen = set()
        unique = []
        for v in vcards:
            tel = get_tel(v)
            if tel and tel not in seen:
                seen.add(tel)
                unique.append(v)
        out = safe_filename(file_name.replace(".vcf", "") + "_nodup.vcf")
        with open(out, "w", encoding="utf-8") as f:
            f.write("\n".join(unique))
        await context.bot.send_document(
            chat_id=update.effective_chat.id,
            document=open(out, "rb"),
            caption=f"✅ {len(unique)} kontak unik (dari {len(vcards)})"
        )
        os.remove(out)

    os.remove(path)
    await update.message.reply_text("🎉 Selesai! Ketik /start.")
    context.user_data.clear()

# --- GETNAME ---
async def proses_getname(update: Update, context: ContextTypes.DEFAULT_TYPE):
    file_id = context.user_data["file_id"]
    file_name = context.user_data["file_name"]
    file = await context.bot.get_file(file_id)
    path = f"temp_getname_{file_id}.vcf"
    await file.download_to_drive(path)

    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        content = f.read()
    os.remove(path)

    names = [get_fn(v) for v in parse_vcards(content)]
    out = "Nama_Kontak.txt"
    with open(out, "w", encoding="utf-8") as f:
        f.write("\n".join(names))
    await context.bot.send_document(
        chat_id=update.effective_chat.id,
        document=open(out, "rb"),
        caption=f"✅ {len(names)} nama diekstrak dari `{file_name}`",
        parse_mode="Markdown"
    )
    os.remove(out)
    await update.message.reply_text("🎉 Selesai! Ketik /start.")
    context.user_data.clear()

# --- GENERATE ---
async def generate_nama(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["gen_base"] = update.message.text.strip()
    await update.message.reply_text(
        f"Nama dasar: **{context.user_data['gen_base']}**\n\n"
        "Ketik **jumlah** yang ingin digenerate (contoh: `100`):",
        parse_mode="Markdown"
    )
    return GENERATE_JUMLAH

async def generate_jumlah(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        jml = int(update.message.text.strip())
        if not (1 <= jml <= 10000):
            raise ValueError
    except ValueError:
        await update.message.reply_text("❌ Jumlah harus 1 - 10000.")
        return GENERATE_JUMLAH

    base = context.user_data.get("gen_base", "Kontak")
    lines = [f"{base} {i:03d}" for i in range(1, jml + 1)]
    out = safe_filename(f"{base}_generated.txt")
    with open(out, "w") as f:
        f.write("\n".join(lines))
    await context.bot.send_document(
        chat_id=update.effective_chat.id,
        document=open(out, "rb"),
        caption=f"✅ {jml} nama digenerate."
    )
    os.remove(out)
    await update.message.reply_text("🎉 Selesai! Ketik /start.")
    context.user_data.clear()
    return MENU

# --- GETCONTEN ---
async def proses_getconten(update: Update, context: ContextTypes.DEFAULT_TYPE):
    file_id = context.user_data["file_id"]
    file_name = context.user_data["file_name"]
    file = await context.bot.get_file(file_id)
    path = f"temp_content_{file_id}.txt"
    await file.download_to_drive(path)

    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        content = f.read()

    await context.bot.send_document(
        chat_id=update.effective_chat.id,
        document=open(path, "rb"),
        caption=f"📄 `{file_name}` • {len(content.splitlines())} baris",
        parse_mode="Markdown"
    )
    os.remove(path)

    lines = content.splitlines()
    preview = "\n".join(lines[:25])
    if len(lines) > 25:
        preview += "\n... (dan seterusnya)"
    await update.message.reply_text(f"Preview:\n```\n{preview}\n```", parse_mode="Markdown")
    context.user_data.clear()

# --- ADMIN & NAVY ---
async def _simpan_grup_dan_tampilkan(update_or_query, context, from_callback=False):
    """Simpan grup saat ini lalu tampilkan ringkasan + tombol lanjut."""
    numbers = context.user_data.get("an_current_numbers", [])
    if not numbers:
        msg = "❌ Belum ada nomor di grup ini."
        if from_callback:
            await update_or_query.answer(msg, show_alert=True)
        else:
            await update_or_query.message.reply_text(msg)
        return ADMIN_NOMOR

    context.user_data.setdefault("an_groups", []).append({
        "name": context.user_data["an_current_name"],
        "numbers": numbers.copy()
    })
    context.user_data["an_current_numbers"] = []

    total_g = len(context.user_data["an_groups"])
    total_n = sum(len(g["numbers"]) for g in context.user_data["an_groups"])
    ringkas = "".join(
        f"• **{g['name']}** → {len(g['numbers'])} nomor\n"
        for g in context.user_data["an_groups"]
    )
    kb = [
        [InlineKeyboardButton("➕ Tambah Grup Lagi", callback_data="an_tambah")],
        [InlineKeyboardButton("✅ Buat 1 File VCF", callback_data="an_buat")],
    ]
    text = (
        f"✅ Grup **{context.user_data['an_current_name']}** disimpan "
        f"({len(numbers)} nomor).\n\n"
        f"📋 **Ringkasan:**\n{ringkas}\n"
        f"Total: **{total_g} grup** | **{total_n} nomor**"
    )
    if from_callback:
        await update_or_query.edit_message_text(
            text, reply_markup=InlineKeyboardMarkup(kb), parse_mode="Markdown"
        )
    else:
        await update_or_query.message.reply_text(
            text, reply_markup=InlineKeyboardMarkup(kb), parse_mode="Markdown"
        )
    return ADMIN_PILIH


async def admin_nama_grup(update: Update, context: ContextTypes.DEFAULT_TYPE):
    nama = update.message.text.strip()
    if not nama:
        await update.message.reply_text("❌ Nama grup tidak boleh kosong.")
        return ADMIN_NAMA_GRUP
    context.user_data["an_current_name"] = nama
    context.user_data["an_current_numbers"] = []
    await update.message.reply_text(
        f"✅ Nama grup: **{nama}**\n\n"
        "Kirim nomor (bisa banyak sekaligus).\n"
        "Setelah selesai, tekan tombol **✅ Selesai Grup Ini**.",
        parse_mode="Markdown"
    )
    return ADMIN_NOMOR


async def admin_nomor(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()

    # Masih izinkan ketik "selesai" sebagai alternatif
    if text.lower() in ("selesai", "done", "finish"):
        return await _simpan_grup_dan_tampilkan(update, context, from_callback=False)

    baru = [clean_number(p) for p in re.split(r"[\s,;]+", text) if len(clean_number(p)) >= 8]
    if not baru:
        await update.message.reply_text(
            "❌ Tidak ada nomor valid.\n"
            "Kirim nomor, atau tekan tombol **✅ Selesai Grup Ini** di bawah."
        )
        # Tetap tampilkan tombol supaya user bisa selesai tanpa ketik
        kb = [[InlineKeyboardButton("✅ Selesai Grup Ini", callback_data="an_selesai_grup")]]
        if context.user_data.get("an_current_numbers"):
            await update.message.reply_text(
                f"Total saat ini: **{len(context.user_data['an_current_numbers'])}** nomor",
                reply_markup=InlineKeyboardMarkup(kb),
                parse_mode="Markdown"
            )
        return ADMIN_NOMOR

    context.user_data.setdefault("an_current_numbers", []).extend(baru)
    total = len(context.user_data["an_current_numbers"])

    kb = [
        [InlineKeyboardButton("➕ Kirim Nomor Lagi", callback_data="an_tambah_nomor")],
        [InlineKeyboardButton("✅ Selesai Grup Ini", callback_data="an_selesai_grup")],
    ]
    await update.message.reply_text(
        f"✅ Ditambah **{len(baru)}** nomor. Total di grup ini: **{total}**\n"
        "Kirim nomor lagi, atau tekan **✅ Selesai Grup Ini**.",
        reply_markup=InlineKeyboardMarkup(kb),
        parse_mode="Markdown"
    )
    return ADMIN_NOMOR


async def admin_pilih(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if query.data == "an_tambah_nomor":
        await query.edit_message_text("Kirim nomor lagi (bisa banyak):")
        return ADMIN_NOMOR

    if query.data == "an_selesai_grup":
        return await _simpan_grup_dan_tampilkan(query, context, from_callback=True)

    if query.data == "an_tambah":
        await query.edit_message_text("Ketik **Nama Grup** berikutnya:")
        return ADMIN_NAMA_GRUP

    if query.data == "an_buat":
        await query.edit_message_text(
            "📝 Ketik **Nama File** (tanpa .vcf)\n"
            "Contoh: `Admin_Navy_Final`"
        )
        return ADMIN_NAMA_FILE

    return ADMIN_PILIH


async def admin_nama_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    raw = update.message.text.strip()
    # Tolak kalau user ketik "selesai" lagi sebagai nama file
    if raw.lower() in ("selesai", "done", "finish", "ya", "ok"):
        await update.message.reply_text(
            "❌ Itu bukan nama file.\n"
            "Ketik nama file yang diinginkan, contoh: `Admin_Navy_Final`"
        )
        return ADMIN_NAMA_FILE

    nama = safe_filename(raw)
    if not nama:
        await update.message.reply_text("❌ Nama file tidak boleh kosong.")
        return ADMIN_NAMA_FILE
    if not nama.lower().endswith(".vcf"):
        nama += ".vcf"

    groups = context.user_data.get("an_groups", [])
    if not groups:
        await update.message.reply_text("❌ Tidak ada data. Ketik /start.")
        return MENU

    await update.message.reply_text("⏳ Membuat file VCF...")
    with open(nama, "w", encoding="utf-8") as f:
        for g in groups:
            for i, n in enumerate(g["numbers"], 1):
                f.write(
                    f"BEGIN:VCARD\nVERSION:3.0\n"
                    f"FN:{g['name']} {i:03d}\nTEL;TYPE=CELL:+{n}\nEND:VCARD\n"
                )

    # Caption sederhana sesuai permintaan
    ringkas_lines = [f"{g['name']} {len(g['numbers'])} kontak" for g in groups]
    ringkas = "\n".join(ringkas_lines)
    total = sum(len(g["numbers"]) for g in groups)

    await context.bot.send_document(
        chat_id=update.effective_chat.id,
        document=open(nama, "rb"),
        caption=(
            f"✅ VCF siap — `{nama}`\n\n"
            f"{ringkas}\n\n"
            f"Total: {len(groups)} grup | {total} kontak"
        ),
        parse_mode="Markdown"
    )
    os.remove(nama)
    await update.message.reply_text("🎉 Selesai! Ketik /start.")
    context.user_data.clear()
    return MENU

# =========================================================
# ADMIN EXTRA
# =========================================================
async def admin_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.from_user.id != ADMIN_ID:
        await query.answer("⛔ Bukan admin!", show_alert=True)
        return MENU

    if query.data == "admin_toggle_fs":
        current = is_force_sub_enabled()
        set_force_sub(not current)
        status = "ON ✅" if not current else "OFF ❌"
        await query.answer(f"Force Sub: {status}", show_alert=True)
        return await show_menu(update, context, edit=True)

    if query.data == "admin_cek_user":
        users = load_users()  # {id: username}
        total = len(users)
        fs = "Aktif ✅" if is_force_sub_enabled() else "Nonaktif ❌"
        await context.bot.send_message(
            chat_id=query.message.chat_id,
            text=f"👥 **Statistik User Bot**\n\n• Total user : `{total}`\n• Force Sub  : {fs}",
            parse_mode="Markdown"
        )
        if total > 0:
            # Buat file daftar id + username
            lines = ["user_id | username"]
            for uid, uname in sorted(users.items()):
                lines.append(f"{uid} | @{uname}" if uname != "-" else f"{uid} | -")
            list_file = "daftar_user.txt"
            with open(list_file, "w", encoding="utf-8") as f:
                f.write("\n".join(lines))
            await context.bot.send_document(
                chat_id=query.message.chat_id,
                document=open(list_file, "rb"),
                caption=f"📄 Daftar {total} User (ID + Username)"
            )
            os.remove(list_file)
        return MENU

    if query.data == "admin_broadcast":
        await query.edit_message_text(
            "📢 **Broadcast**\n\nKirim pesan yang ingin disiarkan ke semua user.\n"
            "Ketik /cancel untuk batal."
        )
        return BROADCAST_MESSAGE

    return MENU

async def broadcast_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("⛔ Akses ditolak.")
        return MENU

    pesan = update.message.text
    users_dict = load_users()  # {id: username}
    users = list(users_dict.keys())
    if not users:
        await update.message.reply_text("❌ Belum ada user.")
        return MENU

    total = len(users)
    await update.message.reply_text(f"⏳ Mengirim ke **{total}** user...")

    sukses = gagal = 0
    blocked = []
    for i, uid in enumerate(users, 1):
        try:
            await context.bot.send_message(
                chat_id=uid,
                text=f"📢 **Pengumuman dari Admin CV Kurumi**\n\n{pesan}",
                parse_mode="Markdown"
            )
            sukses += 1
            await asyncio.sleep(0.07)
        except Forbidden:
            gagal += 1
            blocked.append(uid)
        except RetryAfter as e:
            await asyncio.sleep(e.retry_after + 1)
            try:
                await context.bot.send_message(
                    chat_id=uid,
                    text=f"📢 **Pengumuman dari Admin CV Kurumi**\n\n{pesan}",
                    parse_mode="Markdown"
                )
                sukses += 1
            except Exception:
                gagal += 1
        except Exception:
            gagal += 1

        if i % 25 == 0:
            try:
                await update.message.reply_text(f"⏳ Progress: {i}/{total}")
            except Exception:
                pass

    if blocked:
        # Hapus user yang block dari daftar
        for uid in blocked:
            users_dict.pop(uid, None)
        with open(USERS_FILE, "w", encoding="utf-8") as f:
            for uid, name in sorted(users_dict.items()):
                f.write(f"{uid}|{name}\n")

    await update.message.reply_text(
        f"✅ **Broadcast selesai!**\n\n"
        f"• Total    : `{total}`\n"
        f"• Berhasil : `{sukses}`\n"
        f"• Gagal    : `{gagal}`\n"
        f"• Dihapus  : `{len(blocked)}`"
    )
    return MENU

async def cmd_grant(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("⛔ Hanya admin.")
        return
    args = context.args
    if not args:
        await update.message.reply_text(
            "Format:\n`/grant <user_id>` → permanent\n"
            "`/grant <user_id> <hari>`\n\nContoh: `/grant 123456789 30`",
            parse_mode="Markdown"
        )
        return
    try:
        target = int(args[0])
    except ValueError:
        await update.message.reply_text("❌ User ID harus angka.")
        return
    days = None
    if len(args) >= 2:
        try:
            days = int(args[1])
        except ValueError:
            await update.message.reply_text("❌ Jumlah hari harus angka.")
            return
    grant_access(target, days)
    msg = f"✅ User `{target}` diberi akses **{'permanent' if days is None else f'{days} hari'}**."
    await update.message.reply_text(msg, parse_mode="Markdown")
    try:
        await context.bot.send_message(
            target,
            f"🎉 Akses kamu telah diperpanjang oleh admin.\n"
            f"{'Permanent' if days is None else f'{days} hari'}.\nKetik /start."
        )
    except Exception:
        pass

async def cmd_revoke(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("⛔ Hanya admin.")
        return
    args = context.args
    if not args:
        await update.message.reply_text("Format: `/revoke <user_id>`", parse_mode="Markdown")
        return
    try:
        target = int(args[0])
    except ValueError:
        await update.message.reply_text("❌ User ID harus angka.")
        return
    revoke_access(target)
    await update.message.reply_text(f"✅ Akses user `{target}` dicabut.", parse_mode="Markdown")
    try:
        await context.bot.send_message(target, "🔒 Akses kamu telah dicabut oleh admin.")
    except Exception:
        pass

async def cmd_cekakses(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        _, status = has_access(update.effective_user.id)
        await update.message.reply_text(f"Status akses kamu: `{status}`", parse_mode="Markdown")
        return
    args = context.args
    if not args:
        await update.message.reply_text("Format: `/cekakses <user_id>`", parse_mode="Markdown")
        return
    try:
        target = int(args[0])
    except ValueError:
        await update.message.reply_text("❌ User ID harus angka.")
        return
    ensure_user_access(target)
    _, status = has_access(target)
    info = load_access().get(str(target), {})
    await update.message.reply_text(
        f"👤 User: `{target}`\n"
        f"Status: `{status}`\n"
        f"Tipe: `{info.get('type', '-')}`\n"
        f"Join: `{str(info.get('joined_at', '-'))[:10]}`\n"
        f"Until: `{info.get('access_until') or 'Permanent'}`",
        parse_mode="Markdown"
    )

# =========================================================
# CANCEL & FALLBACK
# =========================================================
async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("❌ Dibatalkan. Ketik /start untuk kembali ke menu.")
    return ConversationHandler.END

async def fallback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("❓ Perintah tidak dikenali. Ketik /start untuk melihat daftar perintah.")
    return MENU

# =========================================================
# MAIN
# =========================================================
def main():
    app = Application.builder().token(BOT_TOKEN).build()

    # Admin commands (luar conversation)
    app.add_handler(CommandHandler("grant", cmd_grant))
    app.add_handler(CommandHandler("revoke", cmd_revoke))
    app.add_handler(CommandHandler("cekakses", cmd_cekakses))
    app.add_handler(CommandHandler("cekuser", cmd_cekuser))

    # Set mode di setiap entry point
    async def entry_to_vcf(u, c):
        c.user_data.clear()
        c.user_data["mode"] = "to_vcf"
        return await cmd_to_vcf(u, c)

    async def entry_to_txt(u, c):
        c.user_data.clear()
        c.user_data["mode"] = "to_txt"
        return await cmd_to_txt(u, c)

    async def entry_add(u, c):
        c.user_data.clear()
        c.user_data["mode"] = "add"
        return await cmd_add(u, c)

    async def entry_delete(u, c):
        c.user_data.clear()
        c.user_data["mode"] = "delete"
        return await cmd_delete(u, c)

    async def entry_renamectc(u, c):
        c.user_data.clear()
        c.user_data["mode"] = "renamectc"
        return await cmd_renamectc(u, c)

    async def entry_renamefile(u, c):
        c.user_data.clear()
        c.user_data["mode"] = "renamefile"
        return await cmd_renamefile(u, c)

    async def entry_merge(u, c):
        c.user_data.clear()
        c.user_data["mode"] = "merge"
        return await cmd_merge(u, c)

    async def entry_split(u, c):
        c.user_data.clear()
        c.user_data["mode"] = "split"
        return await cmd_split(u, c)

    async def entry_count(u, c):
        c.user_data.clear()
        c.user_data["mode"] = "count"
        return await cmd_count(u, c)

    async def entry_nodup(u, c):
        c.user_data.clear()
        c.user_data["mode"] = "nodup"
        return await cmd_nodup(u, c)

    async def entry_getname(u, c):
        c.user_data.clear()
        c.user_data["mode"] = "getname"
        return await cmd_getname(u, c)

    async def entry_getconten(u, c):
        c.user_data.clear()
        c.user_data["mode"] = "getconten"
        return await cmd_getconten(u, c)

    conv = ConversationHandler(
        entry_points=[
            CommandHandler("start", start),
            CommandHandler("to_vcf", entry_to_vcf),
            CommandHandler("to_txt", entry_to_txt),
            CommandHandler("admin", cmd_admin),
            CommandHandler("manual", cmd_manual),
            CommandHandler("add", entry_add),
            CommandHandler("delete", entry_delete),
            CommandHandler("renamectc", entry_renamectc),
            CommandHandler("renamefile", entry_renamefile),
            CommandHandler("merge", entry_merge),
            CommandHandler("split", entry_split),
            CommandHandler("count", entry_count),
            CommandHandler("nodup", entry_nodup),
            CommandHandler("getname", entry_getname),
            CommandHandler("generate", cmd_generate),
            CommandHandler("getconten", entry_getconten),
            CommandHandler("setting", cmd_setting),
            CommandHandler("status", cmd_status),
            CommandHandler("broadcast", cmd_broadcast_start),
        ],
        states={
            MENU: [
                CallbackQueryHandler(admin_callback, pattern="^admin_"),
                CallbackQueryHandler(admin_pilih, pattern="^an_"),
                MessageHandler(filters.Document.ALL, handle_document),
            ],
            # to_vcf
            TO_VCF_FILE: [MessageHandler(filters.Document.ALL, handle_document)],
            TO_VCF_NAMA_FILE: [MessageHandler(filters.TEXT & ~filters.COMMAND, to_vcf_nama_file)],
            TO_VCF_NAMA_KONTAK: [MessageHandler(filters.TEXT & ~filters.COMMAND, to_vcf_nama_kontak)],
            TO_VCF_LIMIT: [MessageHandler(filters.TEXT & ~filters.COMMAND, to_vcf_limit)],
            # to_txt
            TO_TXT_FILE: [MessageHandler(filters.Document.ALL, handle_document)],
            # admin & navy
            ADMIN_NAMA_GRUP: [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_nama_grup)],
            ADMIN_NOMOR: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, admin_nomor),
                CallbackQueryHandler(admin_pilih, pattern="^an_"),
            ],
            ADMIN_PILIH: [CallbackQueryHandler(admin_pilih, pattern="^an_")],
            ADMIN_NAMA_FILE: [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_nama_file)],
            # manual (nomor → nama kontak → nama file)
            MANUAL_NOMOR: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_nomor)],
            MANUAL_NAMA: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_nama)],
            MANUAL_NAMA_FILE: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_nama_file)],
            # add
            ADD_FILE: [MessageHandler(filters.Document.ALL, handle_document)],
            ADD_NOMOR: [MessageHandler(filters.TEXT & ~filters.COMMAND, add_nomor)],
            # delete
            DELETE_FILE: [MessageHandler(filters.Document.ALL, handle_document)],
            DELETE_NOMOR: [MessageHandler(filters.TEXT & ~filters.COMMAND, delete_nomor)],
            # renamectc
            RENAME_CTC_FILE: [MessageHandler(filters.Document.ALL, handle_document)],
            RENAME_CTC_NAMA: [MessageHandler(filters.TEXT & ~filters.COMMAND, rename_ctc_nama)],
            # renamefile
            RENAME_FILE: [MessageHandler(filters.Document.ALL, handle_document)],
            RENAME_FILE_NEW: [MessageHandler(filters.TEXT & ~filters.COMMAND, rename_file_new)],
            # merge
            MERGE_FILE: [
                MessageHandler(filters.Document.ALL, handle_document),
                MessageHandler(filters.TEXT & ~filters.COMMAND, merge_selesai),
            ],
            # split
            SPLIT_FILE: [MessageHandler(filters.Document.ALL, handle_document)],
            SPLIT_LIMIT: [MessageHandler(filters.TEXT & ~filters.COMMAND, split_limit)],
            # count / nodup / getname / getconten
            COUNT_FILE: [MessageHandler(filters.Document.ALL, handle_document)],
            NODUP_FILE: [MessageHandler(filters.Document.ALL, handle_document)],
            GETNAME_FILE: [MessageHandler(filters.Document.ALL, handle_document)],
            GETCONTEN_FILE: [MessageHandler(filters.Document.ALL, handle_document)],
            # generate
            GENERATE_NAMA: [MessageHandler(filters.TEXT & ~filters.COMMAND, generate_nama)],
            GENERATE_JUMLAH: [MessageHandler(filters.TEXT & ~filters.COMMAND, generate_jumlah)],
            # broadcast
            BROADCAST_MESSAGE: [MessageHandler(filters.TEXT & ~filters.COMMAND, broadcast_message)],
        },
        fallbacks=[
            CommandHandler("cancel", cancel),
            CommandHandler("start", start),
            MessageHandler(filters.TEXT & ~filters.COMMAND, fallback),
        ],
        allow_reentry=True,
    )

    app.add_handler(conv)
    print("Bot CV Kurumi (Command Mode) siap jalan...")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
