import os
import io
import csv
import json
import html
import hmac
import base64
import signal
import sqlite3
import hashlib
import asyncio
import logging
from datetime import datetime
from urllib.parse import parse_qsl, unquote
import pytz
import aiohttp
import aiosqlite
import openpyxl
from openpyxl import Workbook

from aiohttp import web
from aiogram import Bot, Dispatcher, F, types
from aiogram.filters import CommandStart, CommandObject
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    ReplyKeyboardMarkup,
    KeyboardButton,
    InlineKeyboardMarkup,
    InlineKeyboardButton,
    WebAppInfo,
    MenuButtonWebApp,
    MenuButtonDefault
)
from aiogram.exceptions import TelegramAPIError, TelegramForbiddenError

# ==========================================
# ⚙️ CONFIGURATION (আপনার তথ্য দিন)
# ==========================================
BOT_TOKEN = "8991869110:AAFLVJJ-oJVN_WGDohERAYypRU3HphjFd40"      # এখানে সরাসরি আপনার বট টোকেন বসান
ADMIN_ID =  7507323720               # এখানে সরাসরি আপনার এডমিন Telegram ID বসান
MINI_APP_URL = os.getenv("MINI_APP_URL", "").strip()
PORT = int(os.getenv("PORT", 8080))
DB_NAME = "database.sqlite3"
DHAKA_TZ = pytz.timezone("Asia/Dhaka")

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(storage=MemoryStorage())

# ==========================================
# ☁️ GITHUB AUTO BACKUP/RESTORE
# Render's filesystem is ephemeral — every restart/redeploy wipes DB_NAME.
# The actual storage stays SQLite (nothing else in this file changes), but
# for GitHub the data is exported as plain JSON — same style as the other
# bot's bot_data.json — so the backup file is human-readable, not a binary
# SQLite blob. On startup this JSON is pulled from GitHub and loaded back
# into the database (after init_db() creates the schema); it's pushed again
# periodically (and on shutdown), so no data is lost across restarts.
# Uses the SAME "bot-backups" repo as the other bot — just a different
# filename (GITHUB_BACKUP_PATH), so both bots' backups can live side by side.
# ==========================================
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN", "").strip()
GITHUB_REPO = os.getenv("GITHUB_REPO", "").strip()               # e.g. "yourname/bot-backups"
GITHUB_BRANCH = os.getenv("GITHUB_BRANCH", "main").strip()
GITHUB_BACKUP_PATH = os.getenv("GITHUB_BACKUP_PATH", "tgz_file_data.json").strip()
GITHUB_BACKUP_INTERVAL = int(os.getenv("GITHUB_BACKUP_INTERVAL", "60"))  # seconds
GITHUB_ENABLED = bool(GITHUB_TOKEN and GITHUB_REPO)
_github_api_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_BACKUP_PATH}"
_github_headers = {"Authorization": f"token {GITHUB_TOKEN}", "Accept": "application/vnd.github+json"}
_github_last_pushed_hash = None


def _export_db_to_json_sync() -> bytes:
    """Read every table in the SQLite database and export it as one JSON
    object: {"table_name": [ {col: val, ...}, ... ], ...}. This is what
    actually gets pushed to GitHub — plain JSON, same style as bot_data.json."""
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")
    tables = [row[0] for row in cur.fetchall()]
    dump = {}
    for table in tables:
        cur.execute(f"SELECT * FROM {table}")
        dump[table] = [dict(row) for row in cur.fetchall()]
    conn.close()
    return json.dumps(dump, ensure_ascii=False, indent=2).encode("utf-8")


def _import_json_into_db_sync(raw_bytes: bytes):
    """Load a JSON backup (produced by _export_db_to_json_sync) back into the
    SQLite database. Must run AFTER init_db() so the table schemas already
    exist; replaces each existing table's rows with the backed-up ones."""
    dump = json.loads(raw_bytes.decode("utf-8"))
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    for table, rows in dump.items():
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name = ?", (table,))
        if not cur.fetchone():
            continue  # এই টেবিল বর্তমান স্কিমাতে নেই, স্কিপ করা হলো
        cur.execute(f"DELETE FROM {table}")
        if rows:
            columns = list(rows[0].keys())
            col_list = ",".join(columns)
            placeholders = ",".join(["?"] * len(columns))
            cur.executemany(
                f"INSERT INTO {table} ({col_list}) VALUES ({placeholders})",
                [tuple(row.get(c) for c in columns) for row in rows]
            )
    conn.commit()
    conn.close()


async def github_pull_backup():
    """Fetch the latest JSON backup from GitHub. Returns the raw JSON bytes
    (or None if there's nothing to restore) — the caller applies it into the
    database AFTER init_db() has created the schema."""
    if not GITHUB_ENABLED:
        logger.info("☁️ GitHub backup not configured (set GITHUB_TOKEN & GITHUB_REPO) — skipping restore.")
        return None
    try:
        async with aiohttp.ClientSession() as session:
            async with session.get(f"{_github_api_url}?ref={GITHUB_BRANCH}", headers=_github_headers, timeout=15) as res:
                if res.status == 200:
                    data = await res.json()
                    raw = base64.b64decode(data.get("content", ""))
                    global _github_last_pushed_hash
                    _github_last_pushed_hash = hashlib.md5(raw).hexdigest()
                    logger.info("☁️ Fetched latest backup JSON from GitHub.")
                    return raw
                elif res.status == 404:
                    logger.info("☁️ No existing GitHub backup found; starting fresh.")
                else:
                    text = await res.text()
                    logger.warning(f"⚠️ GitHub backup pull failed ({res.status}): {text[:200]}")
    except Exception as exc:
        logger.warning(f"⚠️ GitHub backup pull error: {exc}")
    return None


async def github_push_backup(force: bool = False):
    """Export the current database to JSON and push it to GitHub, but only if
    it actually changed since the last push (unless force=True, used on shutdown)."""
    global _github_last_pushed_hash
    if not GITHUB_ENABLED:
        return
    try:
        content_bytes = await asyncio.to_thread(_export_db_to_json_sync)
        current_hash = hashlib.md5(content_bytes).hexdigest()
        if not force and current_hash == _github_last_pushed_hash:
            return  # নতুন কোনো পরিবর্তন নেই, GitHub API-তে অযথা কল করার দরকার নেই

        async with aiohttp.ClientSession() as session:
            sha = None
            async with session.get(f"{_github_api_url}?ref={GITHUB_BRANCH}", headers=_github_headers, timeout=15) as get_res:
                if get_res.status == 200:
                    sha = (await get_res.json()).get("sha")

            payload = {
                "message": f"Auto backup {datetime.now().isoformat()}",
                "content": base64.b64encode(content_bytes).decode("utf-8"),
                "branch": GITHUB_BRANCH,
            }
            if sha:
                payload["sha"] = sha

            async with session.put(_github_api_url, headers=_github_headers, json=payload, timeout=20) as put_res:
                if put_res.status in (200, 201):
                    _github_last_pushed_hash = current_hash
                    logger.info("☁️ Backup JSON pushed to GitHub.")
                else:
                    text = await put_res.text()
                    logger.warning(f"⚠️ GitHub backup push failed ({put_res.status}): {text[:200]}")
    except Exception as exc:
        logger.warning(f"⚠️ GitHub backup push error: {exc}")


async def github_backup_loop():
    """Background loop: checks every GITHUB_BACKUP_INTERVAL seconds and pushes
    to GitHub only when the database actually changed since the last push."""
    if not GITHUB_ENABLED:
        return
    while True:
        await asyncio.sleep(GITHUB_BACKUP_INTERVAL)
        await github_push_backup()

# Telegram premium button icons and button colors, matching the IP Sell UI.
#
# IMPORTANT: Telegram's Bot API has no "style" (color) or "icon_custom_emoji_id"
# field on InlineKeyboardButton / KeyboardButton at all — buttons are always
# plain text. aiogram silently accepts these extra kwargs (it doesn't raise),
# so _button() below never hits its fallback branch, but Telegram itself just
# ignores the unknown fields when it receives them. That's why colored/"premium"
# buttons have never actually shown up on a real device — it isn't a bug in
# this file, it's a hard platform limitation. Premium/custom emoji only render
# for real inside message TEXT (via HTML parse_mode + <tg-emoji>), which is
# what cemoji() below is for. Keep using PREMIUM_ICONS on buttons if you like
# (harmless, and keeps this codebase consistent with your other bots), but use
# cemoji() in message text when you actually want the premium emoji to show.
PREMIUM_ICONS = {
    "back": "5267490665117275176",
    "close": "5420130255174145507",
    "refresh": "5375338737028841420",
    "gear": "5420155432272438703",
    "user": "5352861489541714456",
    "king": "5217822164362739968",
    "gift": "5420396762189831222",
    "link": "5420517437885943844",
    "money": "5190576863226933563",
    "card": "5190899075968441286",
    "chat": "5192704641564974847",
    "folder": "5352721946054268944",
    "bag": "5229064374403998351",
    "send": "5353001161878182134",
    "shield": "5190447043545438788",
    "megaphone": "5424818078833715060",  # 📣 used for the Support/Channel button
    "bubble": "5443038326535759644",     # 💬 used for the Support/Community button
}

# Reference table of every premium/custom emoji ID from the "News Emoji" pack
# supplied by Tarikul, keyed by the normal fallback emoji it renders instead
# of. Use cemoji() to drop these into message TEXT (never into button labels
# — Telegram doesn't support custom emoji there).
NEWS_EMOJI_IDS = {
    '👀': '5210956306952758910', '🙂': '5461117441612462242', '⚡️': '5456140674028019486',
    '☄️': '5224607267797606837', '🛍': '5229064374403998351', '⛔️': '5260293700088511294',
    '🚫': '5240241223632954241', '❗️': '5274099962655816924', '‼️': '5440660757194744323',
    '⁉️': '5314504236132747481', '❓': '5436113877181941026', '⚠️': '5447644880824181073',
    '🌐': '5447410659077661506', '💬': '5443038326535759644', '💭': '5467538555158943525',
    '📊': '5231200819986047254', '🔼': '5449683594425410231', '🔽': '5447183459602669338',
    '🕯': '5451882707875276247', '📈': '5244837092042750681', '📉': '5246762912428603768',
    '✔️': '5206607081334906820', '❌': '5210952531676504517', '🆒': '5222079954421818267',
    '🔔': '5458603043203327669', '🥸': '5391112412445288650', '🤡': '5269531045165816230',
    '🫦': '5395444514028529554', '📌': '5397782960512444700', '💵': '5409048419211682843',
    '💸': '5233326571099534068', '💱': '5402186569006210455', '▶️': '5264919878082509254',
    '🔴': '5411225014148014586', '🟢': '5416081784641168838', '➡️': '5416117059207572332',
    '🔥': '5424972470023104089', '💥': '5276032951342088188', '🎙': '5294339927318739359',
    '🎤': '5224736245665511429', '📣': '5424818078833715060', '🤫': '5431609822288033666',
    '👎': '5449875686837726134', '🗣️': '5460795800101594035', '🔍': '5231012545799666522',
    '🛡': '5251203410396458957', '🔗': '5271604874419647061', '🖥': '5282843764451195532',
    '©': '5323442290708985472', 'ℹ️': '5334544901428229844', '👍': '5337080053119336309',
    '⏸': '5359543311897998264', '💯': '5341498088408234504', '🔄': '5375338737028841420',
    '🔝': '5415655814079723871', '🆕': '5382357040008021292', '🔜': '5440621591387980068',
    '📍': '5391032818111363540', '➕': '5397916757333654639', '💎': '5427168083074628963',
    '⭐️': '5438496463044752972', '✨': '5325547803936572038', '👑': '5217822164362739968',
    '🗑': '5445267414562389170', '🔖': '5222444124698853913', '✉️': '5253742260054409879',
    '🔒': '5296369303661067030', '😮': '5303479226882603449', '📎': '5305265301917549162',
    '⚙️': '5341715473882955310', '🎮': '5361741454685256344', '🔈': '5388632425314140043',
    '⌛': '5386367538735104399', '⬇️': '5406745015365943482', '☀️': '5402477260982731644',
    '🌧': '5399913388845322366', '🌛': '5449569374065152798', '❄️': '5449449325434266744',
    '🌈': '5409109841538994759', '💧': '5393512611968995988', '🗓': '5413879192267805083',
    '💡': '5422439311196834318', '🥇': '5440539497383087970', '🥈': '5447203607294265305',
    '🥉': '5453902265922376865', '🎵': '5463107823946717464', '🆓': '5406756500108501710',
    '✏️': '5395444784611480792', '🚨': '5395695537687123235', '🏠': '5416041192905265756',
    '🚩': '5460755126761312667', '🎉': '5461151367559141950',
}

def cemoji(fallback: str, icon: str | None = None) -> str:
    """Return HTML for a real premium/custom emoji to use inside message TEXT
    (requires parse_mode="HTML"). Looks the id up by icon name in PREMIUM_ICONS
    first, then by the fallback character in NEWS_EMOJI_IDS. Falls back to the
    plain emoji if no id is known. NEVER use this inside a button label —
    Telegram buttons cannot render custom emoji at all."""
    eid = (PREMIUM_ICONS.get(icon) if icon else None) or NEWS_EMOJI_IDS.get(fallback)
    if not eid:
        return fallback
    return f'<tg-emoji emoji-id="{eid}">{fallback}</tg-emoji>'

def _button(button_type, text, icon=None, style="primary", **kwargs):
    """Build a colored/custom-icon button with a safe legacy fallback."""
    button_kwargs = dict(kwargs)
    emoji_id = PREMIUM_ICONS.get(icon) if icon else None
    if emoji_id:
        button_kwargs["icon_custom_emoji_id"] = emoji_id
    if style:
        button_kwargs["style"] = style
    try:
        return button_type(text=text, **button_kwargs)
    except Exception:
        # Older aiogram versions do not know the new Telegram button fields.
        button_kwargs.pop("icon_custom_emoji_id", None)
        button_kwargs.pop("style", None)
        return button_type(text=text, **button_kwargs)

def pbtn(text, icon=None, style="primary", callback_data=None, url=None, web_app=None):
    button_kwargs = {}
    if callback_data is not None:
        button_kwargs["callback_data"] = callback_data
    if url is not None:
        button_kwargs["url"] = url
    if web_app is not None:
        button_kwargs["web_app"] = web_app
    return _button(
        InlineKeyboardButton,
        text,
        icon=icon,
        style=style,
        **button_kwargs,
    )

def rkbtn(text, icon=None, style="primary"):
    return _button(KeyboardButton, text, icon=icon, style=style)

def get_mini_app_url() -> str:
    """Use the deployed public HTTPS URL automatically after deployment."""
    configured = MINI_APP_URL
    candidates = [
        configured,
        os.getenv("REPLIT_DEPLOYMENT_URL", ""),
        os.getenv("REPLIT_DEV_DOMAIN", ""),
        os.getenv("REPLIT_DOMAINS", "").split(",")[0].strip(),
        os.getenv("RENDER_EXTERNAL_URL", ""),
        os.getenv("PUBLIC_URL", ""),
    ]
    for value in candidates:
        value = value.strip().rstrip("/")
        if value:
            if not value.startswith(("http://", "https://")):
                value = "https://" + value
            return value + "/"
    raise RuntimeError(
        "No public URL found. Set MINI_APP_URL or expose REPLIT_DEPLOYMENT_URL/RENDER_EXTERNAL_URL."
    )

# ==========================================
# 🗄️ DATABASE INITIALIZATION
# ==========================================
async def init_db():
    async with aiosqlite.connect(DB_NAME) as db:
        await db.execute("PRAGMA journal_mode=WAL;")
        
        await db.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                telegram_id INTEGER UNIQUE,
                username TEXT,
                first_name TEXT,
                balance REAL DEFAULT 0.0,
                is_blocked INTEGER DEFAULT 0,
                created_at TEXT,
                last_active TEXT
            )
        """)

        await db.execute("""
            CREATE TABLE IF NOT EXISTS admins (
                telegram_id INTEGER PRIMARY KEY,
                created_at TEXT NOT NULL
            )
        """)
        await db.execute(
            "INSERT OR IGNORE INTO admins (telegram_id, created_at) VALUES (?, ?)",
            (int(ADMIN_ID), now_dhaka_str())
        )
        await db.execute("""
            CREATE TABLE IF NOT EXISTS payment_methods (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL,
                is_active INTEGER DEFAULT 1,
                created_at TEXT NOT NULL
            )
        """)
        for method in ("bKash", "Nagad", "Rocket"):
            await db.execute(
                "INSERT OR IGNORE INTO payment_methods (name, is_active, created_at) VALUES (?, 1, ?)",
                (method, now_dhaka_str())
            )
        
        await db.execute("""
            CREATE TABLE IF NOT EXISTS categories (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE,
                is_active INTEGER DEFAULT 1,
                created_at TEXT
            )
        """)
        
        await db.execute("""
            CREATE TABLE IF NOT EXISTS services (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                category_id INTEGER,
                name TEXT,
                price REAL DEFAULT 0.0,
                is_active INTEGER DEFAULT 1,
                created_at TEXT,
                FOREIGN KEY(category_id) REFERENCES categories(id) ON DELETE CASCADE
            )
        """)
        
        await db.execute("""
            CREATE TABLE IF NOT EXISTS submissions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER,
                telegram_id INTEGER,
                category_id INTEGER,
                service_id INTEGER,
                total_ids INTEGER,
                duplicate_ids INTEGER,
                accepted_ids INTEGER,
                status TEXT DEFAULT 'pending',
                file_id TEXT,
                created_at TEXT
            )
        """)
        
        await db.execute("""
            CREATE TABLE IF NOT EXISTS submission_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                submission_id INTEGER,
                user_id INTEGER,
                telegram_id INTEGER,
                category_id INTEGER,
                service_id INTEGER,
                account_id TEXT,
                is_matched INTEGER DEFAULT 0,
                created_at TEXT,
                FOREIGN KEY(submission_id) REFERENCES submissions(id) ON DELETE CASCADE
            )
        """)
        await db.execute("CREATE INDEX IF NOT EXISTS idx_sub_items ON submission_items(service_id, account_id, is_matched);")
        # Safe migrations for databases created by earlier versions.
        try:
            await db.execute("ALTER TABLE submission_items ADD COLUMN row_data TEXT")
        except Exception:
            pass

        await db.execute("""
            CREATE TABLE IF NOT EXISTS matched_records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER,
                telegram_id INTEGER,
                category_id INTEGER,
                service_id INTEGER,
                submission_item_id INTEGER UNIQUE,
                account_id TEXT,
                price REAL,
                reference_code TEXT,
                date_dhaka TEXT,
                created_at TEXT
            )
        """)

        await db.execute("""
            CREATE TABLE IF NOT EXISTS withdrawals (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER,
                telegram_id INTEGER,
                method TEXT,
                account_number TEXT,
                amount REAL,
                status TEXT DEFAULT 'Pending',
                created_at TEXT
            )
        """)

        await db.execute("""
            CREATE TABLE IF NOT EXISTS referrals (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                referrer_id INTEGER,
                referred_user_id INTEGER UNIQUE,
                bonus_paid REAL DEFAULT 0.0,
                created_at TEXT
            )
        """)

        await db.execute("""
            CREATE TABLE IF NOT EXISTS bot_settings (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        """)

        defaults = [
            ("support_link", "https://t.me/"),
            ("force_join_enabled", "0"),
            ("force_join_channel", ""),
            ("force_join_link", "https://t.me/"),
            ("custom_link_1_title", "Channel"),
            ("custom_link_1_url", "https://t.me/"),
            ("custom_link_2_title", "Community"),
            ("custom_link_2_url", "https://t.me/")
        ]
        for k, v in defaults:
            await db.execute("INSERT OR IGNORE INTO bot_settings (key, value) VALUES (?, ?)", (k, v))
            
        await db.commit()

async def get_setting(key: str, default: str = "") -> str:
    async with aiosqlite.connect(DB_NAME) as db:
        async with db.execute("SELECT value FROM bot_settings WHERE key = ?", (key,)) as cursor:
            row = await cursor.fetchone()
            return row[0] if row else default

async def set_setting(key: str, value: str):
    async with aiosqlite.connect(DB_NAME) as db:
        await db.execute("INSERT OR REPLACE INTO bot_settings (key, value) VALUES (?, ?)", (key, value))
        await db.commit()

def now_dhaka_str():
    return datetime.now(DHAKA_TZ).strftime("%Y-%m-%d %H:%M:%S")

def today_dhaka_str():
    return datetime.now(DHAKA_TZ).strftime("%Y-%m-%d")

def normalize_account_id(value) -> str:
    """Keep IDs stable when Excel/Google Sheets adds whitespace or .0."""
    if value is None:
        return ""
    text = str(value).replace("\ufeff", "").strip()
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return text

async def is_admin_id(telegram_id: int) -> bool:
    if int(telegram_id) == int(ADMIN_ID):
        return True
    async with aiosqlite.connect(DB_NAME) as db:
        async with db.execute("SELECT 1 FROM admins WHERE telegram_id = ?", (int(telegram_id),)) as c:
            return await c.fetchone() is not None

# ==========================================
# 🤖 BOT FSM & KEYBOARDS
# ==========================================
class SubmitFileState(StatesGroup):
    waiting_for_category = State()
    waiting_for_service = State()
    waiting_for_file = State()

class WithdrawState(StatesGroup):
    waiting_for_method = State()
    waiting_for_account = State()
    waiting_for_amount = State()

async def get_main_keyboard(user_id: int) -> ReplyKeyboardMarkup:
    kb = [
        [rkbtn("Submit File", icon="send", style="success")],
        [
            rkbtn("Balance", icon="money", style="primary"),
            rkbtn("Withdraw", icon="card", style="success"),
        ],
        [
            rkbtn("Refer", icon="gift", style="primary"),
            rkbtn("Leaderboard", icon="king", style="primary"),
        ],
        [rkbtn("Support", icon="chat", style="primary")],
    ]
    if await is_admin_id(user_id):
        # Keep the reply keyboard button as a normal button. The Web App is
        # opened only after the admin presses the inline button in the reply.
        kb.append([rkbtn("Admin Panel", icon="gear", style="primary")])
    return ReplyKeyboardMarkup(keyboard=kb, resize_keyboard=True)

async def check_force_join(user_id: int) -> bool:
    enabled = await get_setting("force_join_enabled", "0")
    if enabled != "1":
        return True
    channel = await get_setting("force_join_channel", "")
    if not channel:
        return True
    try:
        member = await bot.get_chat_member(chat_id=channel, user_id=user_id)
        if member.status in ["member", "administrator", "creator"]:
            return True
    except Exception as e:
        logger.error(f"Force join check error: {e}")
        return True
    return False

@dp.message(F.text == "Admin Panel")
async def open_admin_panel(message: types.Message):
    if not await is_admin_id(message.from_user.id):
        return await message.answer("You are not authorized to open the Admin Panel.")
    try:
        admin_webapp_url = get_mini_app_url()
    except RuntimeError:
        return await message.answer("Admin Panel URL is not configured yet.")
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [pbtn("Open Admin Panel", icon="gear", style="primary", web_app=WebAppInfo(url=admin_webapp_url))]
    ])
    await message.answer(
        "Press the button below to open the Admin Panel.",
        reply_markup=kb
    )

# ==========================================
# 🤖 USER BOT HANDLERS
# ==========================================
@dp.message(CommandStart())
async def cmd_start(message: types.Message, command: CommandObject, state: FSMContext):
    await state.clear()
    user_id = message.from_user.id
    username = message.from_user.username or ""
    first_name = message.from_user.first_name or "User"
    now_str = now_dhaka_str()

    async with aiosqlite.connect(DB_NAME) as db:
        async with db.execute("SELECT id, is_blocked FROM users WHERE telegram_id = ?", (user_id,)) as cursor:
            user = await cursor.fetchone()
        
        if not user:
            await db.execute(
                "INSERT INTO users (telegram_id, username, first_name, balance, created_at, last_active) VALUES (?, ?, ?, 0.0, ?, ?)",
                (user_id, username, first_name, now_str, now_str)
            )
            await db.commit()
            
            if command.args and command.args.startswith("ref_"):
                try:
                    ref_id = int(command.args.split("ref_")[1])
                    if ref_id != user_id:
                        async with db.execute("SELECT id FROM users WHERE telegram_id = ?", (ref_id,)) as c:
                            if await c.fetchone():
                                await db.execute(
                                    "INSERT OR IGNORE INTO referrals (referrer_id, referred_user_id, created_at) VALUES (?, ?, ?)",
                                    (ref_id, user_id, now_str)
                                )
                                await db.commit()
                except Exception as e:
                    logger.error(f"Referral parsing error: {e}")
        else:
            if user[1] == 1:
                return await message.answer("⛔ Your account is suspended. Contact support.")
            await db.execute(
                "UPDATE users SET username = ?, first_name = ?, last_active = ? WHERE telegram_id = ?",
                (username, first_name, now_str, user_id)
            )
            await db.commit()

    if not await check_force_join(user_id):
        link = await get_setting("force_join_link", "https://t.me/")
        kb = InlineKeyboardMarkup(inline_keyboard=[
            [pbtn("Join Channel", icon="link", style="primary", url=link)],
            [pbtn("Check Membership", icon="refresh", style="success", callback_data="check_force_join")]
        ])
        return await message.answer(
            "⚠️ **You must join our official channel to use this bot!**\n\n"
            "Please join the channel and click **Check Membership** below.",
            reply_markup=kb,
            parse_mode="Markdown"
        )

    welcome_text = (
        f"👋 **Welcome, {first_name}!**\n\n"
        "Earn daily rewards by submitting valid account IDs securely.\n"
        "Select an option from the menu below to get started:"
    )
    await message.answer(welcome_text, reply_markup=await get_main_keyboard(user_id), parse_mode="Markdown")

@dp.callback_query(F.data == "check_force_join")
async def cb_check_force_join(call: types.CallbackQuery, state: FSMContext):
    if await check_force_join(call.from_user.id):
        await call.message.delete()
        await call.message.answer(
            "✅ **Verification successful! Welcome to the bot.**",
            reply_markup=await get_main_keyboard(call.from_user.id),
            parse_mode="Markdown"
        )
    else:
        await call.answer("❌ You have not joined the channel yet!", show_alert=True)

# ----------------- Submit File Flow -----------------
@dp.message(F.text == "Submit File")
async def submit_file_start(message: types.Message, state: FSMContext):
    if not await check_force_join(message.from_user.id):
        return await message.answer("⚠️ Please join our channel first by sending /start.")

    async with aiosqlite.connect(DB_NAME) as db:
        async with db.execute("SELECT id, name FROM categories WHERE is_active = 1") as cursor:
            categories = await cursor.fetchall()

    if not categories:
        return await message.answer("🚫 No active categories are available right now. Please check back later.")

    # Map each exact button label -> category row so the next step can look
    # the selection up directly, without re-querying/re-matching by name.
    category_map = {c[1]: (c[0], c[1]) for c in categories}
    kb = [[rkbtn(c[1], icon="folder", style="primary")] for c in categories]
    kb.append([rkbtn("Cancel", icon="close", style="danger")])
    await state.update_data(category_map=category_map)
    await state.set_state(SubmitFileState.waiting_for_category)
    await message.answer("📁 **Select a Category:**", reply_markup=ReplyKeyboardMarkup(keyboard=kb, resize_keyboard=True), parse_mode="Markdown")

@dp.message(SubmitFileState.waiting_for_category)
async def submit_file_category_selected(message: types.Message, state: FSMContext):
    if message.text == "Cancel":
        await state.clear()
        return await message.answer("Cancelled.", reply_markup=await get_main_keyboard(message.from_user.id))
    if message.text == "Support":
        await state.clear()
        return await show_support(message)

    data = await state.get_data()
    category_map = data.get("category_map") or {}
    category = category_map.get(message.text)

    if not category:
        return await message.answer("⚠️ Please select a valid category from the keyboard buttons.")

    category_id, category_name = category[0], category[1]

    async with aiosqlite.connect(DB_NAME) as db:
        async with db.execute("SELECT id, name, price FROM services WHERE category_id = ? AND is_active = 1", (category_id,)) as cursor:
            services = await cursor.fetchall()

    if not services:
        await state.clear()
        return await message.answer("🚫 No active services found for this category.", reply_markup=await get_main_keyboard(message.from_user.id))

    # Map each exact button label -> service row, so the next step can look
    # the selection up directly instead of re-parsing/re-querying by name
    # (the old "split on em-dash then re-query" approach broke as soon as a
    # category had more than a couple of services, or a name had extra
    # spaces / a dash in it).
    service_map = {}
    kb = []
    for s in services:
        label = f"{s[1]} — ৳{s[2]:.2f}/acc"
        # Guard against two services rendering an identical label (e.g. same
        # name + price) by disambiguating with the service id.
        if label in service_map:
            label = f"{s[1]} (#{s[0]}) — ৳{s[2]:.2f}/acc"
        service_map[label] = (s[0], s[1], s[2])
        kb.append([rkbtn(label, icon="bag", style="primary")])
    kb.append([rkbtn("Cancel", icon="close", style="danger")])

    await state.update_data(category_id=category_id, category_name=category_name, service_map=service_map)
    await state.set_state(SubmitFileState.waiting_for_service)
    await message.answer(f"📂 **Category:** {category_name}\n⚙️ **Select a Service:**", reply_markup=ReplyKeyboardMarkup(keyboard=kb, resize_keyboard=True), parse_mode="Markdown")

@dp.message(SubmitFileState.waiting_for_service)
async def submit_file_service_selected(message: types.Message, state: FSMContext):
    if message.text == "Cancel":
        await state.clear()
        return await message.answer("Cancelled.", reply_markup=await get_main_keyboard(message.from_user.id))
    if message.text == "Support":
        await state.clear()
        return await show_support(message)

    data = await state.get_data()
    service_map = data.get("service_map") or {}
    service = service_map.get(message.text)

    if not service:
        return await message.answer("⚠️ Please select a valid service from the keyboard buttons.")

    service_id, service_name, service_price = service[0], service[1], service[2]
    await state.update_data(service_id=service_id, service_name=service_name, service_price=service_price)
    await state.set_state(SubmitFileState.waiting_for_file)
    
    kb = ReplyKeyboardMarkup(
        keyboard=[[rkbtn("Cancel", icon="close", style="danger")]],
        resize_keyboard=True,
    )
    await message.answer(
        f"📁 **Service Selected:** {service_name} (৳{service_price:.2f}/acc)\n\n"
        "Please send your **Excel (.xlsx)** file now.\n"
        "The system will automatically validate the file and extract submitted IDs.",
        reply_markup=kb,
        parse_mode="Markdown"
    )

@dp.message(SubmitFileState.waiting_for_file, ~F.document)
async def submit_file_waiting_for_file_non_document(message: types.Message, state: FSMContext):
    # Catches everything except documents while we're waiting for the Excel
    # file — most importantly "Cancel", which previously fell through and did
    # nothing because the only handler registered for this state required
    # F.document.
    if message.text == "Cancel":
        await state.clear()
        return await message.answer("Cancelled.", reply_markup=await get_main_keyboard(message.from_user.id))
    if message.text == "Support":
        await state.clear()
        return await show_support(message)

    await message.answer("❌ Please send a valid **.xlsx** Excel file, or press Cancel.", parse_mode="Markdown")

@dp.message(SubmitFileState.waiting_for_file, F.document)
async def handle_excel_file(message: types.Message, state: FSMContext):
    doc = message.document
    if not (doc.file_name and doc.file_name.lower().endswith(".xlsx")):
        return await message.answer("❌ Invalid file type! Please upload a valid **.xlsx** Excel file.")

    processing_msg = await message.answer("⏳ *Downloading and processing your file...*", parse_mode="Markdown")
    data = await state.get_data()
    category_id = data["category_id"]
    service_id = data["service_id"]
    user_id = message.from_user.id
    now_str = now_dhaka_str()

    try:
        file_io = io.BytesIO()
        await bot.download(doc, destination=file_io)
        file_io.seek(0)
        
        workbook = openpyxl.load_workbook(file_io, data_only=True)
        sheet = workbook.active

        raw_ids = []
        source_rows = []
        for row in sheet.iter_rows(values_only=True):
            row_values = [normalize_account_id(cell) for cell in (row or ())]
            if not any(row_values):
                continue
            # Account IDs are in the first populated column, not every cell
            # (cookies/usernames in the other columns must not be counted).
            account = next((cell for cell in row_values if cell), "")
            if account.lower() in {"id", "account", "account_id", "username", "email"}:
                continue
            if account:
                raw_ids.append(account)
                source_rows.append(row_values)

        if not raw_ids:
            await processing_msg.delete()
            return await message.answer("❌ The uploaded Excel file is empty or has no readable account IDs.")

        total_ids = len(raw_ids)
        seen = set()
        unique_ids = []
        unique_rows = []
        duplicate_count = 0
        for item, row_data in zip(raw_ids, source_rows):
            if item not in seen:
                seen.add(item)
                unique_ids.append(item)
                unique_rows.append(row_data)
            else:
                duplicate_count += 1

        accepted_count = len(unique_ids)

        async with aiosqlite.connect(DB_NAME) as db:
            async with db.execute("SELECT id FROM users WHERE telegram_id = ?", (user_id,)) as cursor:
                u_row = await cursor.fetchone()
                db_user_id = u_row[0] if u_row else None

            cursor = await db.execute(
                """INSERT INTO submissions 
                   (user_id, telegram_id, category_id, service_id, total_ids, duplicate_ids, accepted_ids, status, file_id, created_at)
                   VALUES (?, ?, ?, ?, ?, ?, ?, 'pending', ?, ?)""",
                (db_user_id, user_id, category_id, service_id, total_ids, duplicate_count, accepted_count, doc.file_id, now_str)
            )
            submission_id = cursor.lastrowid

            items_to_insert = [
                (submission_id, db_user_id, user_id, category_id, service_id, acc_id, 0, now_str, json.dumps(row_data, ensure_ascii=False))
                for acc_id, row_data in zip(unique_ids, unique_rows)
            ]
            await db.executemany(
                """INSERT INTO submission_items 
                   (submission_id, user_id, telegram_id, category_id, service_id, account_id, is_matched, created_at, row_data)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                items_to_insert
            )
            await db.commit()

        await processing_msg.delete()
        await state.clear()
        confirmation = (
            "✅ **Your file has been submitted successfully!**\n\n"
            f"📊 **Total IDs:** {total_ids}\n"
            f"♻️ **Duplicate IDs:** {duplicate_count}\n"
            f"✅ **Accepted Unique IDs:** {accepted_count}\n\n"
            "💰 *Your payment will be added automatically once submitted IDs are matched & approved in the Reports system.*"
        )
        await message.answer(confirmation, reply_markup=await get_main_keyboard(user_id), parse_mode="Markdown")

    except Exception as e:
        logger.error(f"Error parsing excel: {e}")
        await processing_msg.delete()
        await message.answer("❌ Error processing Excel file. Ensure it is a valid, uncorrupted .xlsx file.", reply_markup=await get_main_keyboard(user_id))

# ----------------- Balance -----------------
@dp.message(F.text == "Balance")
async def show_balance(message: types.Message):
    user_id = message.from_user.id
    today_str = today_dhaka_str()

    async with aiosqlite.connect(DB_NAME) as db:
        async with db.execute("SELECT balance FROM users WHERE telegram_id = ?", (user_id,)) as cursor:
            row = await cursor.fetchone()
            total_balance = row[0] if row else 0.0

        async with db.execute(
            "SELECT COALESCE(SUM(price), 0.0) FROM matched_records WHERE telegram_id = ? AND date_dhaka = ?",
            (user_id, today_str)
        ) as cursor:
            today_earned = (await cursor.fetchone())[0]

        async with db.execute(
            "SELECT COALESCE(SUM(accepted_ids), 0) FROM submissions WHERE telegram_id = ?", (user_id,)
        ) as cursor:
            total_submitted = (await cursor.fetchone())[0]

        async with db.execute(
            "SELECT COUNT(id) FROM submission_items WHERE telegram_id = ? AND is_matched = 0", (user_id,)
        ) as cursor:
            pending_ids = (await cursor.fetchone())[0]

    text = (
        "💼 **Your Financial Summary:**\n\n"
        f"💰 **Total Balance:** ৳{total_balance:.2f}\n"
        f"📅 **Today's Earnings:** ৳{today_earned:.2f}\n"
        f"📤 **Total IDs Submitted:** {total_submitted}\n"
        f"⏳ **Pending Verification:** {pending_ids} IDs\n\n"
        "💡 *Earnings are added instantly when your submissions match admin reports.*"
    )
    await message.answer(text, reply_markup=await get_main_keyboard(user_id), parse_mode="Markdown")

# ----------------- Refer -----------------
@dp.message(F.text == "Refer")
async def show_refer(message: types.Message):
    user_id = message.from_user.id
    bot_info = await bot.get_me()
    ref_link = f"https://t.me/{bot_info.username}?start=ref_{user_id}"

    async with aiosqlite.connect(DB_NAME) as db:
        async with db.execute("SELECT COUNT(id) FROM referrals WHERE referrer_id = ?", (user_id,)) as cursor:
            total_refs = (await cursor.fetchone())[0]

    text = (
        "👥 **Referral Program**\n\n"
        "Invite your friends and earn rewards for every active partner!\n\n"
        f"🔗 **Your Referral Link:**\n`{ref_link}`\n\n"
        f"📊 **Total Referred:** {total_refs} users"
    )
    await message.answer(text, reply_markup=await get_main_keyboard(user_id), parse_mode="Markdown")

# ----------------- Withdraw Flow -----------------
@dp.message(F.text == "Withdraw")
async def withdraw_start(message: types.Message, state: FSMContext):
    user_id = message.from_user.id
    async with aiosqlite.connect(DB_NAME) as db:
        async with db.execute("SELECT balance FROM users WHERE telegram_id = ?", (user_id,)) as cursor:
            row = await cursor.fetchone()
            balance = row[0] if row else 0.0

    if balance < 10.0:
        return await message.answer(f"❌ Minimum withdrawal amount is **৳10.00**. Your current balance is **৳{balance:.2f}**.", parse_mode="Markdown")

    async with aiosqlite.connect(DB_NAME) as db:
        async with db.execute("SELECT name FROM payment_methods WHERE is_active = 1 ORDER BY id") as c:
            methods = [row[0] for row in await c.fetchall()]
    if not methods:
        return await message.answer("Withdraw is currently disabled.")
    method_icons = {
        "bkash": "money",
        "nagad": "money",
        "rocket": "send",
    }
    kb = [[rkbtn(name, icon=method_icons.get(name.lower(), "money"), style="success")] for name in methods]
    kb.append([rkbtn("Cancel", icon="close", style="danger")])
    await state.set_state(WithdrawState.waiting_for_method)
    await state.update_data(current_balance=balance)
    await message.answer(
        f"💰 **Current Balance:** ৳{balance:.2f}\n\nSelect your payout method:",
        reply_markup=ReplyKeyboardMarkup(keyboard=kb, resize_keyboard=True),
        parse_mode="Markdown"
    )

@dp.message(WithdrawState.waiting_for_method)
async def withdraw_method_chosen(message: types.Message, state: FSMContext):
    if message.text == "Cancel":
        await state.clear()
        return await message.answer("Withdrawal cancelled.", reply_markup=await get_main_keyboard(message.from_user.id))
    if message.text == "Support":
        await state.clear()
        return await show_support(message)

    async with aiosqlite.connect(DB_NAME) as db:
        async with db.execute("SELECT name FROM payment_methods WHERE name = ? AND is_active = 1", (message.text,)) as c:
            method_row = await c.fetchone()
    if not method_row:
        return await message.answer("This payment method is disabled or unavailable.")
    method = method_row[0]
    await state.update_data(withdraw_method=method)
    await state.set_state(WithdrawState.waiting_for_account)
    await message.answer(
        f"📱 Enter your **{method}** account number:",
        reply_markup=ReplyKeyboardMarkup(
            keyboard=[[rkbtn("Cancel", icon="close", style="danger")]],
            resize_keyboard=True,
        ),
        parse_mode="Markdown",
    )

@dp.message(WithdrawState.waiting_for_account)
async def withdraw_account_chosen(message: types.Message, state: FSMContext):
    if message.text == "Cancel":
        await state.clear()
        return await message.answer("Withdrawal cancelled.", reply_markup=await get_main_keyboard(message.from_user.id))
    if message.text == "Support":
        await state.clear()
        return await show_support(message)

    account_num = message.text.strip()
    if len(account_num) < 10:
        return await message.answer("⚠️ Please enter a valid account number.")

    await state.update_data(account_number=account_num)
    await state.set_state(WithdrawState.waiting_for_amount)
    data = await state.get_data()
    await message.answer(f"💵 Enter withdrawal amount (Max: ৳{data['current_balance']:.2f}):", parse_mode="Markdown")

@dp.message(WithdrawState.waiting_for_amount)
async def withdraw_amount_chosen(message: types.Message, state: FSMContext):
    if message.text == "Cancel":
        await state.clear()
        return await message.answer("Withdrawal cancelled.", reply_markup=await get_main_keyboard(message.from_user.id))
    if message.text == "Support":
        await state.clear()
        return await show_support(message)

    try:
        amount = float(message.text.strip())
    except ValueError:
        return await message.answer("⚠️ Please enter a valid numerical amount.")

    user_id = message.from_user.id
    data = await state.get_data()
    balance = data["current_balance"]

    if amount < 10.0:
        return await message.answer("❌ Minimum withdrawal amount is ৳10.00.")
    if amount > balance:
        return await message.answer(f"❌ Insufficient balance! Maximum available: ৳{balance:.2f}")

    now_str = now_dhaka_str()
    async with aiosqlite.connect(DB_NAME) as db:
        await db.execute("UPDATE users SET balance = balance - ? WHERE telegram_id = ?", (amount, user_id))
        async with db.execute("SELECT id FROM users WHERE telegram_id = ?", (user_id,)) as c:
            db_u_id = (await c.fetchone())[0]
        await db.execute(
            "INSERT INTO withdrawals (user_id, telegram_id, method, account_number, amount, status, created_at) VALUES (?, ?, ?, ?, ?, 'Pending', ?)",
            (db_u_id, user_id, data["withdraw_method"], data["account_number"], amount, now_str)
        )
        await db.commit()

    await state.clear()
    await message.answer(
        f"✅ **Withdrawal Request Placed!**\n\n"
        f"💳 **Method:** {data['withdraw_method']}\n"
        f"🔢 **Account:** `{data['account_number']}`\n"
        f"💵 **Amount:** ৳{amount:.2f}\n"
        f"⏳ **Status:** Pending Review\n\n"
        "Your request will be processed soon by the admin.",
        reply_markup=await get_main_keyboard(user_id),
        parse_mode="Markdown"
    )

# ----------------- Leaderboard -----------------
@dp.message(F.text == "Leaderboard")
async def show_leaderboard(message: types.Message):
    today_str = today_dhaka_str()
    async with aiosqlite.connect(DB_NAME) as db:
        query = """
            SELECT u.first_name, SUM(m.price) as today_earnings
            FROM matched_records m
            JOIN users u ON m.telegram_id = u.telegram_id
            WHERE m.date_dhaka = ?
            GROUP BY m.telegram_id
            ORDER BY today_earnings DESC
            LIMIT 10
        """
        async with db.execute(query, (today_str,)) as cursor:
            leaders = await cursor.fetchall()

    text = f"🏆 **Today's Leaderboard ({today_str} Dhaka Time)**\n\n"
    if not leaders:
        text += "No earnings recorded today yet. Be the first!"
    else:
        medals = ["🥇", "🥈", "🥉", "4️⃣", "5️⃣", "6️⃣", "7️⃣", "8️⃣", "9️⃣", "🔟"]
        for idx, (name, earnings) in enumerate(leaders):
            badge = medals[idx] if idx < len(medals) else f"#{idx+1}"
            safe_name = name or "User"
            text += f"{badge} **{safe_name}** — ৳{earnings:.2f}\n"

    text += "\n*Resets daily at 12:00 AM Asia/Dhaka.*"
    await message.answer(text, reply_markup=await get_main_keyboard(message.from_user.id), parse_mode="Markdown")

# ----------------- Support -----------------
@dp.message(F.text == "Support")
async def show_support(message: types.Message):
    support_link = await get_setting("support_link", "https://t.me/")
    c1_title = await get_setting("custom_link_1_title", "Channel")
    c1_url = await get_setting("custom_link_1_url", "https://t.me/")
    c2_title = await get_setting("custom_link_2_title", "Community")
    c2_url = await get_setting("custom_link_2_url", "https://t.me/")

    # Button labels stay plain text — Telegram buttons can't render custom
    # emoji at all, so icon= below is cosmetically inert on real devices.
    # The premium emoji actually renders here, in the message text, via
    # HTML parse_mode + <tg-emoji>.
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [pbtn("Contact Support", icon="chat", style="primary", url=support_link)],
        [
            pbtn(c1_title, icon="megaphone", style="primary", url=c1_url),
            pbtn(c2_title, icon="bubble", style="success", url=c2_url),
        ],
    ])
    text = (
        f"{cemoji('💬', 'bubble')} <b>Need help, have questions, or want updates?</b>\n\n"
        f"{cemoji('☎️')} Contact Support — talk to our team directly\n"
        f"{cemoji('📣', 'megaphone')} {html.escape(c1_title)} — news and announcements\n"
        f"{cemoji('👥')} {html.escape(c2_title)} — join the community\n\n"
        "Tap a button below:"
    )
    await message.answer(text, reply_markup=kb, parse_mode="HTML")

# ==========================================
# 🌐 MINI APP AUTH & REST API
# ==========================================
def verify_telegram_init_data(init_data: str) -> dict | None:
    """Robust Telegram WebApp HMAC-SHA256 initData validation."""
    if not init_data:
        return None
    try:
        parsed = dict(parse_qsl(init_data, keep_blank_values=True))
        if "hash" not in parsed:
            return None
        received_hash = parsed.pop("hash")
        
        # Sort key=value pairs alphabetically
        data_check_string = "\n".join(f"{k}={parsed[k]}" for k in sorted(parsed.keys()))
        
        secret_key = hmac.new(b"WebAppData", BOT_TOKEN.encode("utf-8"), hashlib.sha256).digest()
        computed_hash = hmac.new(secret_key, data_check_string.encode("utf-8"), hashlib.sha256).hexdigest()

        if hmac.compare_digest(computed_hash, received_hash):
            user_data = json.loads(parsed.get("user", "{}"))
            return user_data
    except Exception as e:
        logger.error(f"initData verification error: {e}")
    return None

async def authenticate_admin(request: web.Request):
    auth_header = (
        request.headers.get("X-Telegram-Init-Data")
        or request.query.get("initData")
        or request.query.get("tgWebAppData")
    )
    if not auth_header:
        raise web.HTTPUnauthorized(text=json.dumps({"error": "Missing initData"}), content_type="application/json")
    
    user_info = verify_telegram_init_data(auth_header)
    telegram_id = int(user_info.get("id", 0)) if user_info else 0
    if not user_info or not await is_admin_id(telegram_id):
        logger.warning("Admin authorization failed: webapp_user_id=%s configured_admin_id=%s", telegram_id, ADMIN_ID)
        raise web.HTTPForbidden(text=json.dumps({"error": "Unauthorized Access"}), content_type="application/json")
    return user_info

# --- Admin API Routes ---
async def api_dashboard_stats(request: web.Request):
    await authenticate_admin(request)
    today_str = today_dhaka_str()
    async with aiosqlite.connect(DB_NAME) as db:
        async with db.execute("SELECT COUNT(id) FROM users") as c:
            total_users = (await c.fetchone())[0]
        async with db.execute("SELECT COUNT(id) FROM users WHERE created_at LIKE ?", (f"{today_str}%",)) as c:
            today_users = (await c.fetchone())[0]
        async with db.execute("SELECT COALESCE(SUM(total_ids), 0) FROM submissions") as c:
            total_submitted_ids = (await c.fetchone())[0]
        async with db.execute("SELECT COALESCE(SUM(total_ids), 0) FROM submissions WHERE created_at LIKE ?", (f"{today_str}%",)) as c:
            today_submitted_ids = (await c.fetchone())[0]
        async with db.execute("SELECT COALESCE(SUM(amount), 0) FROM withdrawals WHERE status = 'Paid'") as c:
            total_paid = (await c.fetchone())[0]
        async with db.execute("SELECT COUNT(id) FROM withdrawals WHERE status = 'Pending'") as c:
            pending_withdrawals = (await c.fetchone())[0]
        async with db.execute("SELECT COUNT(id) FROM submission_items WHERE is_matched = 0") as c:
            pending_items = (await c.fetchone())[0]
        async with db.execute("SELECT COUNT(id) FROM categories") as c:
            total_categories = (await c.fetchone())[0]
        async with db.execute("SELECT COUNT(id) FROM services") as c:
            total_services = (await c.fetchone())[0]

    return web.json_response({
        "total_users": total_users,
        "today_users": today_users,
        "total_submitted_ids": total_submitted_ids,
        "today_submitted_ids": today_submitted_ids,
        "total_paid": total_paid,
        "pending_withdrawals": pending_withdrawals,
        "pending_items": pending_items,
        "total_categories": total_categories,
        "total_services": total_services
    })

async def api_get_users(request: web.Request):
    await authenticate_admin(request)
    async with aiosqlite.connect(DB_NAME) as db:
        db.row_factory = aiosqlite.Row
        async with db.execute("SELECT id, telegram_id, username, first_name, balance, is_blocked, created_at FROM users ORDER BY id DESC") as c:
            rows = await c.fetchall()
            users = [dict(r) for r in rows]
    return web.json_response(users)

async def api_admins(request: web.Request):
    await authenticate_admin(request)
    async with aiosqlite.connect(DB_NAME) as db:
        if request.method == "GET":
            async with db.execute("SELECT telegram_id, created_at FROM admins ORDER BY telegram_id") as c:
                return web.json_response([{"telegram_id": r[0], "created_at": r[1]} for r in await c.fetchall()])
        data = await request.json()
        admin_id = int(data.get("telegram_id"))
        if admin_id <= 0:
            return web.json_response({"error": "Invalid Telegram ID"}, status=400)
        if request.method == "DELETE":
            if admin_id == int(ADMIN_ID):
                return web.json_response({"error": "The primary admin cannot be removed"}, status=400)
            await db.execute("DELETE FROM admins WHERE telegram_id = ?", (admin_id,))
        else:
            await db.execute("INSERT OR IGNORE INTO admins (telegram_id, created_at) VALUES (?, ?)", (admin_id, now_dhaka_str()))
        await db.commit()
        return web.json_response({"status": "deleted" if request.method == "DELETE" else "created"})

async def api_payment_methods(request: web.Request):
    await authenticate_admin(request)
    async with aiosqlite.connect(DB_NAME) as db:
        if request.method == "GET":
            async with db.execute("SELECT id, name, is_active FROM payment_methods ORDER BY id") as c:
                return web.json_response([
                    {"id": r[0], "name": r[1], "is_active": r[2]} for r in await c.fetchall()
                ])
        data = await request.json()
        if request.method == "POST":
            name = str(data.get("name", "")).strip()
            if not name:
                return web.json_response({"error": "Name is required"}, status=400)
            await db.execute(
                "INSERT OR IGNORE INTO payment_methods (name, is_active, created_at) VALUES (?, 1, ?)",
                (name, now_dhaka_str())
            )
        elif request.method == "PUT":
            await db.execute("UPDATE payment_methods SET is_active = ? WHERE id = ?", (int(data.get("is_active", 0)), int(data["id"])))
        elif request.method == "DELETE":
            await db.execute("DELETE FROM payment_methods WHERE id = ?", (int(data["id"]),))
        await db.commit()
    return web.json_response({"status": "updated"})

async def api_user_action(request: web.Request):
    await authenticate_admin(request)
    data = await request.json()
    action = data.get("action")
    user_id = data.get("user_id")

    async with aiosqlite.connect(DB_NAME) as db:
        if action == "toggle_block":
            await db.execute("UPDATE users SET is_blocked = CASE WHEN is_blocked=1 THEN 0 ELSE 1 END WHERE telegram_id = ?", (user_id,))
        elif action == "adjust_balance":
            amount = float(data.get("amount", 0.0))
            await db.execute("UPDATE users SET balance = balance + ? WHERE telegram_id = ?", (amount, user_id))
        await db.commit()
    return web.json_response({"status": "success"})

async def api_categories(request: web.Request):
    await authenticate_admin(request)
    async with aiosqlite.connect(DB_NAME) as db:
        db.row_factory = aiosqlite.Row
        if request.method == "GET":
            async with db.execute("SELECT * FROM categories ORDER BY id DESC") as c:
                return web.json_response([dict(r) for r in await c.fetchall()])
        elif request.method == "POST":
            data = await request.json()
            name = data.get("name", "").strip()
            if name:
                await db.execute("INSERT OR IGNORE INTO categories (name, is_active, created_at) VALUES (?, 1, ?)", (name, now_dhaka_str()))
                await db.commit()
            return web.json_response({"status": "created"})
        elif request.method == "PUT":
            data = await request.json()
            cat_id = data.get("id")
            if "is_active" in data:
                await db.execute("UPDATE categories SET is_active = ? WHERE id = ?", (data["is_active"], cat_id))
            if "name" in data:
                await db.execute("UPDATE categories SET name = ? WHERE id = ?", (data["name"], cat_id))
            await db.commit()
            return web.json_response({"status": "updated"})
        elif request.method == "DELETE":
            cat_id = request.query.get("id")
            await db.execute("DELETE FROM categories WHERE id = ?", (cat_id,))
            await db.commit()
            return web.json_response({"status": "deleted"})

async def api_services(request: web.Request):
    await authenticate_admin(request)
    async with aiosqlite.connect(DB_NAME) as db:
        db.row_factory = aiosqlite.Row
        if request.method == "GET":
            async with db.execute("""
                SELECT s.*, c.name as category_name 
                FROM services s 
                LEFT JOIN categories c ON s.category_id = c.id 
                ORDER BY s.id DESC
            """) as c:
                return web.json_response([dict(r) for r in await c.fetchall()])
        elif request.method == "POST":
            data = await request.json()
            category_id = data.get("category_id")
            name = data.get("name")
            price = float(data.get("price", 0.0))
            await db.execute(
                "INSERT INTO services (category_id, name, price, is_active, created_at) VALUES (?, ?, ?, 1, ?)",
                (category_id, name, price, now_dhaka_str())
            )
            await db.commit()
            return web.json_response({"status": "created"})
        elif request.method == "PUT":
            data = await request.json()
            srv_id = data.get("id")
            if "is_active" in data:
                await db.execute("UPDATE services SET is_active = ? WHERE id = ?", (data["is_active"], srv_id))
            if "price" in data:
                await db.execute("UPDATE services SET price = ? WHERE id = ?", (float(data["price"]), srv_id))
            if "name" in data:
                await db.execute("UPDATE services SET name = ? WHERE id = ?", (data["name"], srv_id))
            await db.commit()
            return web.json_response({"status": "updated"})
        elif request.method == "DELETE":
            srv_id = request.query.get("id")
            await db.execute("DELETE FROM services WHERE id = ?", (srv_id,))
            await db.commit()
            return web.json_response({"status": "deleted"})

async def api_files_summary(request: web.Request):
    await authenticate_admin(request)
    async with aiosqlite.connect(DB_NAME) as db:
        db.row_factory = aiosqlite.Row
        query = """
            SELECT s.id as service_id, s.name as service_name, c.name as category_name,
                   COUNT(si.id) as total_submitted_count,
                   SUM(CASE WHEN si.is_matched = 0 THEN 1 ELSE 0 END) as unmatched_count
            FROM services s
            JOIN categories c ON s.category_id = c.id
            LEFT JOIN submission_items si ON s.id = si.service_id
            GROUP BY s.id
        """
        async with db.execute(query) as c:
            return web.json_response([dict(r) for r in await c.fetchall()])

async def api_files_ids(request: web.Request):
    """Return every submitted account ID for a service, shown openly in the
    admin panel instead of only being available as a hidden download."""
    await authenticate_admin(request)
    service_id = request.query.get("service_id")
    async with aiosqlite.connect(DB_NAME) as db:
        async with db.execute(
            "SELECT account_id FROM submission_items WHERE service_id = ? ORDER BY id", (service_id,)
        ) as c:
            rows = await c.fetchall()
    return web.json_response([r[0] for r in rows])

async def api_files_download(request: web.Request):
    await authenticate_admin(request)
    service_id = request.query.get("service_id")
    async with aiosqlite.connect(DB_NAME) as db:
        async with db.execute("SELECT name FROM services WHERE id = ?", (service_id,)) as c:
            row = await c.fetchone()
            service_name = row[0] if row else "service"
        
        async with db.execute("SELECT account_id, row_data FROM submission_items WHERE service_id = ? AND is_matched = 0 ORDER BY id", (service_id,)) as c:
            rows = await c.fetchall()

    # Rebuild an actual XLSX workbook and keep every original row column.
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Submitted Accounts"
    for account_id, row_data in rows:
        try:
            values = json.loads(row_data) if row_data else [account_id]
            if not isinstance(values, list):
                values = [account_id]
        except (TypeError, ValueError):
            values = [account_id]
        sheet.append(values)

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)

    filename = f"{service_name.lower().replace(' ', '_')}_unmatched.xlsx"
    return web.Response(
        body=output.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )

async def api_files_delete_all(request: web.Request):
    await authenticate_admin(request)
    async with aiosqlite.connect(DB_NAME) as db:
        # Delete child rows first so this also works on older SQLite databases
        # that were created without foreign_keys=ON.
        await db.execute("DELETE FROM matched_records")
        await db.execute("DELETE FROM submission_items")
        await db.execute("DELETE FROM submissions")
        await db.commit()
    return web.json_response({"status": "deleted"})

async def api_reports_match(request: web.Request):
    await authenticate_admin(request)
    data = await request.json()
    service_id = int(data.get("service_id"))
    raw_results = data.get("results_text", "")
    
    submitted_matches = []
    for line in raw_results.splitlines():
        first_cell = line.split("\t")[0].split(",")[0]
        value = normalize_account_id(first_cell)
        if value and value.lower() not in {"id", "account", "account_id"}:
            submitted_matches.append(value)
    if not submitted_matches:
        return web.json_response({"error": "No valid result IDs provided."}, status=400)

    now_str = now_dhaka_str()
    date_dhaka = today_dhaka_str()

    user_rewards = {}
    total_processed = 0

    async with aiosqlite.connect(DB_NAME) as db:
        async with db.execute("SELECT price, name FROM services WHERE id = ?", (service_id,)) as c:
            s_row = await c.fetchone()
            if not s_row:
                return web.json_response({"error": "Invalid service"}, status=400)
            service_price, service_name = s_row

        for account_id in submitted_matches:
            async with db.execute(
                "SELECT id, user_id, telegram_id, category_id FROM submission_items WHERE service_id = ? AND account_id = ? AND is_matched = 0 LIMIT 1",
                (service_id, account_id)
            ) as c:
                item = await c.fetchone()

            if item:
                item_id, db_u_id, tg_id, cat_id = item
                ref_code = f"MATCH_{item_id}_{int(datetime.now().timestamp())}"

                await db.execute("UPDATE submission_items SET is_matched = 1 WHERE id = ?", (item_id,))
                
                await db.execute(
                    """INSERT OR IGNORE INTO matched_records 
                       (user_id, telegram_id, category_id, service_id, submission_item_id, account_id, price, reference_code, date_dhaka, created_at)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (db_u_id, tg_id, cat_id, service_id, item_id, account_id, service_price, ref_code, date_dhaka, now_str)
                )

                await db.execute("UPDATE users SET balance = balance + ? WHERE telegram_id = ?", (service_price, tg_id))

                if tg_id not in user_rewards:
                    user_rewards[tg_id] = {"count": 0, "amount": 0.0}
                user_rewards[tg_id]["count"] += 1
                user_rewards[tg_id]["amount"] += service_price
                total_processed += 1

        await db.commit()

    for tg_id, stats in user_rewards.items():
        try:
            msg = (
                "🎉 **Payment Added!**\n\n"
                f"📌 **Service:** {service_name}\n"
                f"✅ **Matched:** {stats['count']} IDs\n"
                f"💵 **Rate:** ৳{service_price:.2f}/ID\n"
                f"💰 **Total Credited:** ৳{stats['amount']:.2f}\n\n"
                "Your balance has been updated successfully!"
            )
            await bot.send_message(chat_id=tg_id, text=msg, parse_mode="Markdown")
        except Exception as e:
            logger.error(f"Failed to notify user {tg_id}: {e}")

    return web.json_response({
        "status": "success",
        "matched_count": total_processed,
        "users_credited": len(user_rewards)
    })

async def api_withdrawals(request: web.Request):
    await authenticate_admin(request)
    async with aiosqlite.connect(DB_NAME) as db:
        db.row_factory = aiosqlite.Row
        if request.method == "GET":
            async with db.execute("""
                SELECT w.*, u.first_name, u.username 
                FROM withdrawals w 
                LEFT JOIN users u ON w.telegram_id = u.telegram_id 
                ORDER BY w.id DESC
            """) as c:
                return web.json_response([dict(r) for r in await c.fetchall()])
        elif request.method == "POST":
            data = await request.json()
            w_id = data.get("id")
            new_status = data.get("status")
            
            async with db.execute("SELECT telegram_id, amount, status FROM withdrawals WHERE id = ?", (w_id,)) as c:
                w_row = await c.fetchone()

            if w_row:
                tg_id, amount, old_status = w_row["telegram_id"], w_row["amount"], w_row["status"]
                
                if new_status == "Rejected" and old_status != "Rejected":
                    await db.execute("UPDATE users SET balance = balance + ? WHERE telegram_id = ?", (amount, tg_id))
                
                await db.execute("UPDATE withdrawals SET status = ? WHERE id = ?", (new_status, w_id))
                await db.commit()

                try:
                    await bot.send_message(
                        chat_id=tg_id,
                        text=f"🔔 **Withdrawal Update:**\nYour withdrawal request of **৳{amount:.2f}** has been marked as **{new_status}**.",
                        parse_mode="Markdown"
                    )
                except Exception as e:
                    logger.error(f"Failed to notify user on withdrawal update: {e}")

            return web.json_response({"status": "updated"})

async def api_broadcast(request: web.Request):
    await authenticate_admin(request)
    data = await request.json()
    message_text = data.get("text", "")
    photo_url = data.get("photo_url", "").strip()

    if not message_text and not photo_url:
        return web.json_response({"error": "Empty message"}, status=400)

    async with aiosqlite.connect(DB_NAME) as db:
        async with db.execute("SELECT telegram_id FROM users WHERE is_blocked = 0") as c:
            users = await c.fetchall()

    success, failed = 0, 0
    for (uid,) in users:
        try:
            if photo_url:
                await bot.send_photo(chat_id=uid, photo=photo_url, caption=message_text, parse_mode="Markdown")
            else:
                await bot.send_message(chat_id=uid, text=message_text, parse_mode="Markdown")
            success += 1
            await asyncio.sleep(0.04)
        except (TelegramForbiddenError, TelegramAPIError):
            failed += 1
        except Exception:
            failed += 1

    return web.json_response({"status": "completed", "success": success, "failed": failed})

async def api_settings(request: web.Request):
    await authenticate_admin(request)
    if request.method == "GET":
        async with aiosqlite.connect(DB_NAME) as db:
            db.row_factory = aiosqlite.Row
            async with db.execute("SELECT * FROM bot_settings") as c:
                rows = await c.fetchall()
                settings = {r["key"]: r["value"] for r in rows}
        return web.json_response(settings)
    elif request.method == "POST":
        data = await request.json()
        for k, v in data.items():
            await set_setting(k, str(v))
        return web.json_response({"status": "saved"})

# --- Mini App HTML Interface ---
MINI_APP_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
  <title>Admin Dashboard</title>
  <script src="https://telegram.org/js/telegram-web-app.js"></script>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@500;600;700;800&family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
  <style>
    :root {
      --bg: #eef1f8;
      --surface: #ffffff;
      --surface-alt: #f8fafc;
      --primary: #4f46e5;
      --primary-dark: #4338ca;
      --primary-soft: #eef0ff;
      --accent: #0891b2;
      --accent-soft: #e0f7fa;
      --text: #0f172a;
      --text-muted: #64748b;
      --success: #16a34a;
      --success-soft: #e7f8ed;
      --danger: #dc2626;
      --danger-soft: #fdeaea;
      --warning: #d97706;
      --warning-soft: #fdf3e0;
      --border: #e6eaf2;
      --shadow-sm: 0 1px 2px rgba(15, 23, 42, .05);
      --shadow-md: 0 6px 20px -6px rgba(38, 43, 89, .12);
      --radius: 16px;
      --radius-sm: 10px;
    }

    * { box-sizing: border-box; margin: 0; padding: 0; }
    html { scroll-behavior: smooth; }

    body {
      background: var(--bg);
      background-image:
        radial-gradient(600px 300px at 100% -10%, rgba(79, 70, 229, .07), transparent),
        radial-gradient(500px 260px at -10% 0%, rgba(8, 145, 178, .06), transparent);
      background-attachment: fixed;
      color: var(--text);
      padding-bottom: 90px;
      font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      -webkit-font-smoothing: antialiased;
    }

    h1, h2, h3 { font-family: 'Plus Jakarta Sans', 'Inter', sans-serif; }

    .icon { width: 1em; height: 1em; display: inline-block; vertical-align: -0.15em; flex-shrink: 0; }

    /* ---------- Header ---------- */
    header {
      background: rgba(255, 255, 255, .85);
      backdrop-filter: blur(10px);
      -webkit-backdrop-filter: blur(10px);
      padding: 14px 18px;
      border-bottom: 1px solid var(--border);
      position: sticky;
      top: 0;
      z-index: 100;
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 10px;
    }
    .brand { display: flex; align-items: center; gap: 10px; }
    .brand-icon {
      width: 34px; height: 34px;
      border-radius: 10px;
      background: linear-gradient(135deg, var(--primary), var(--accent));
      display: flex; align-items: center; justify-content: center;
      color: #fff;
      font-size: 18px;
      box-shadow: var(--shadow-md);
      flex-shrink: 0;
    }
    h1 { font-size: 16px; font-weight: 800; color: var(--text); letter-spacing: -0.01em; }
    .brand-sub { font-size: 11px; color: var(--text-muted); font-weight: 500; margin-top: 1px; }

    .badge {
      padding: 5px 10px 5px 8px;
      border-radius: 20px;
      font-size: 11px;
      font-weight: 700;
      display: inline-flex;
      align-items: center;
      gap: 5px;
      white-space: nowrap;
    }
    .badge-success { background: var(--success-soft); color: var(--success); }
    .badge-warning { background: var(--warning-soft); color: var(--warning); }
    .badge-danger { background: var(--danger-soft); color: var(--danger); }
    .pulse-dot {
      width: 7px; height: 7px; border-radius: 50%; background: var(--success);
      animation: pulse 1.8s ease-out infinite;
      flex-shrink: 0;
    }
    @keyframes pulse {
      0% { box-shadow: 0 0 0 0 rgba(22, 163, 74, .55); }
      70% { box-shadow: 0 0 0 6px rgba(22, 163, 74, 0); }
      100% { box-shadow: 0 0 0 0 rgba(22, 163, 74, 0); }
    }

    /* ---------- Nav tabs ---------- */
    .nav-tabs {
      display: flex;
      overflow-x: auto;
      background: rgba(255, 255, 255, .7);
      backdrop-filter: blur(8px);
      border-bottom: 1px solid var(--border);
      padding: 8px 10px;
      gap: 6px;
      position: sticky;
      top: 61px;
      z-index: 90;
      scrollbar-width: none;
    }
    .nav-tabs::-webkit-scrollbar { display: none; }
    .nav-tab {
      padding: 8px 13px;
      font-size: 12.5px;
      font-weight: 600;
      font-family: inherit;
      color: var(--text-muted);
      background: transparent;
      border: 1px solid transparent;
      border-radius: 999px;
      cursor: pointer;
      white-space: nowrap;
      display: inline-flex;
      align-items: center;
      gap: 6px;
      transition: background .2s ease, color .2s ease, transform .15s ease, box-shadow .2s ease;
    }
    .nav-tab .icon { width: 15px; height: 15px; }
    .nav-tab:hover { background: var(--primary-soft); color: var(--primary); }
    .nav-tab:active { transform: scale(.95); }
    .nav-tab.active {
      background: linear-gradient(135deg, var(--primary), var(--primary-dark));
      color: #fff;
      box-shadow: 0 6px 16px -4px rgba(79, 70, 229, .45);
    }

    .container { padding: 16px; max-width: 720px; margin: 0 auto; }

    /* ---------- Cards ---------- */
    .card {
      background: var(--surface);
      border: 1px solid var(--border);
      border-radius: var(--radius);
      padding: 18px;
      margin-bottom: 16px;
      box-shadow: var(--shadow-sm);
      transition: box-shadow .25s ease;
    }
    .card:hover { box-shadow: var(--shadow-md); }
    .card h3 {
      font-size: 14px;
      font-weight: 700;
      margin-bottom: 12px;
      display: flex;
      align-items: center;
      gap: 8px;
      color: var(--text);
    }
    .card h3 .icon { width: 17px; height: 17px; color: var(--primary); }

    .grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 12px; }
    .stat-card {
      background: var(--surface);
      border: 1px solid var(--border);
      border-radius: var(--radius);
      padding: 16px;
      box-shadow: var(--shadow-sm);
      transition: transform .2s ease, box-shadow .2s ease;
      animation: fadeUp .45s ease both;
    }
    .stat-card:hover { transform: translateY(-3px); box-shadow: var(--shadow-md); }
    .stat-card:nth-child(1) { animation-delay: .02s; }
    .stat-card:nth-child(2) { animation-delay: .07s; }
    .stat-card:nth-child(3) { animation-delay: .12s; }
    .stat-card:nth-child(4) { animation-delay: .17s; }
    .stat-card:nth-child(5) { animation-delay: .22s; }
    .stat-card:nth-child(6) { animation-delay: .27s; }
    .stat-icon {
      width: 34px; height: 34px;
      border-radius: 10px;
      display: flex; align-items: center; justify-content: center;
      margin-bottom: 10px;
      font-size: 17px;
    }
    .stat-icon .icon { width: 17px; height: 17px; }
    .stat-icon.indigo { background: var(--primary-soft); color: var(--primary); }
    .stat-icon.cyan { background: var(--accent-soft); color: var(--accent); }
    .stat-icon.amber { background: var(--warning-soft); color: var(--warning); }
    .stat-icon.green { background: var(--success-soft); color: var(--success); }
    .stat-icon.red { background: var(--danger-soft); color: var(--danger); }
    .stat-val { font-size: 21px; font-weight: 800; font-family: 'Plus Jakarta Sans', sans-serif; letter-spacing: -0.01em; }
    .stat-label { font-size: 11px; color: var(--text-muted); text-transform: uppercase; letter-spacing: .03em; font-weight: 600; margin-top: 2px; }
    .stat-card-clickable { cursor: pointer; }
    .stat-card-clickable:active { transform: translateY(-1px) scale(.98); }

    /* ---------- Submitted IDs full page ---------- */
    .page-overlay {
      position: fixed; inset: 0; z-index: 200;
      background: var(--bg);
      display: flex; flex-direction: column;
      animation: slideInPage .22s ease both;
    }
    @keyframes slideInPage { from { transform: translateX(16px); opacity: 0; } to { transform: translateX(0); opacity: 1; } }
    .page-header {
      display: flex; align-items: center; gap: 10px;
      padding: 14px 16px;
      background: rgba(255, 255, 255, .85);
      backdrop-filter: blur(10px);
      -webkit-backdrop-filter: blur(10px);
      border-bottom: 1px solid var(--border);
      position: sticky; top: 0; z-index: 5;
    }
    .page-header h3 { display: flex; align-items: center; gap: 8px; font-size: 15px; }
    .page-header h3 .icon { width: 17px; height: 17px; color: var(--primary); }
    .page-back {
      width: 34px; height: 34px; padding: 0; flex-shrink: 0; border-radius: 50%;
      background: var(--surface-alt); border: 1px solid var(--border);
      display: flex; align-items: center; justify-content: center;
    }
    .page-body { padding: 16px; max-width: 720px; margin: 0 auto; width: 100%; overflow-y: auto; flex: 1; }
    .modal-loading { text-align: center; color: var(--text-muted); font-size: 13px; padding: 30px 0; }

    /* ---------- Developer Info Modal ---------- */
    .dev-modal-overlay {
      position: fixed; inset: 0; z-index: 300;
      background: rgba(15, 23, 42, .55);
      backdrop-filter: blur(3px);
      display: flex; align-items: center; justify-content: center;
      padding: 20px;
      animation: fadeSlideIn .25s ease;
    }
    .dev-modal-box {
      position: relative;
      width: 100%; max-width: 340px;
      background: var(--surface);
      border-radius: var(--radius);
      box-shadow: var(--shadow-md);
      padding: 30px 22px 22px 22px;
      text-align: center;
    }
    .dev-modal-close {
      position: absolute; top: 12px; right: 12px;
      width: 28px; height: 28px;
      border: none; border-radius: 50%;
      background: var(--surface-alt);
      color: var(--text-muted);
      display: flex; align-items: center; justify-content: center;
      cursor: pointer;
    }
    .dev-modal-close .icon { width: 14px; height: 14px; }
    .dev-modal-avatar {
      width: 64px; height: 64px;
      margin: 0 auto 14px auto;
      border-radius: 18px;
      background: linear-gradient(135deg, var(--primary), var(--accent));
      display: flex; align-items: center; justify-content: center;
      color: #fff;
    }
    .dev-modal-avatar .icon { width: 30px; height: 30px; }
    .dev-modal-name { font-family: 'Plus Jakarta Sans', sans-serif; font-weight: 800; font-size: 16px; color: var(--text); margin-bottom: 16px; }
    .dev-modal-rows { display: flex; flex-direction: column; gap: 8px; margin-bottom: 20px; }
    .dev-modal-row {
      display: flex; justify-content: space-between; align-items: center;
      background: var(--surface-alt);
      border: 1px solid var(--border);
      border-radius: var(--radius-sm);
      padding: 9px 14px;
      font-size: 13px;
    }
    .dev-modal-label { color: var(--text-muted); font-weight: 600; display: flex; align-items: center; gap: 6px; }
    .dev-modal-label .icon { width: 14px; height: 14px; }
    .dev-modal-value { color: var(--text); font-weight: 700; }
    .dev-modal-contact {
      width: 100%;
      border: none; border-radius: var(--radius-sm);
      background: linear-gradient(135deg, var(--primary), var(--primary-dark));
      color: #fff; font-weight: 700; font-size: 14px;
      padding: 12px;
      display: flex; align-items: center; justify-content: center; gap: 8px;
      cursor: pointer;
      box-shadow: var(--shadow-md);
    }
    .dev-modal-contact .icon { width: 16px; height: 16px; }

    .svc-block { border: 1px solid var(--border); border-radius: var(--radius-sm); margin-bottom: 10px; overflow: hidden; }
    .svc-head {
      display: flex; align-items: center; justify-content: space-between; gap: 8px;
      padding: 11px 13px; cursor: pointer; background: var(--surface-alt);
    }
    .svc-head-info { display: flex; flex-direction: column; gap: 1px; }
    .svc-head-cat { font-size: 10.5px; color: var(--text-muted); font-weight: 600; text-transform: uppercase; letter-spacing: .02em; }
    .svc-head-name { font-size: 13.5px; font-weight: 700; }
    .svc-head-count { font-size: 12px; font-weight: 700; color: var(--primary); background: var(--primary-soft); padding: 3px 9px; border-radius: 999px; }
    .svc-ids-panel { display: none; padding: 12px 13px; border-top: 1px solid var(--border); }
    .svc-ids-panel.open { display: block; }
    .svc-ids-list { display: flex; flex-wrap: wrap; gap: 6px; max-height: 220px; overflow-y: auto; }
    .svc-id-chip {
      font-size: 12px; font-family: 'Inter', monospace;
      background: var(--surface-alt); border: 1px solid var(--border);
      padding: 4px 8px; border-radius: 8px; color: var(--text);
    }
    .svc-empty { font-size: 12px; color: var(--text-muted); padding: 4px 2px; }

    @keyframes fadeUp {
      from { opacity: 0; transform: translateY(10px); }
      to { opacity: 1; transform: translateY(0); }
    }

    /* ---------- Forms ---------- */
    label { display: block; font-size: 12px; font-weight: 600; color: var(--text-muted); margin-top: 10px; margin-bottom: 4px; }
    input, select, textarea, button {
      width: 100%;
      padding: 11px 12px;
      margin-top: 4px;
      margin-bottom: 12px;
      border-radius: var(--radius-sm);
      border: 1.5px solid var(--border);
      background: var(--surface-alt);
      color: var(--text);
      font-size: 14px;
      font-family: inherit;
      transition: border-color .18s ease, box-shadow .18s ease, background .18s ease;
    }
    input::placeholder, textarea::placeholder { color: #a4adba; }
    input:focus, select:focus, textarea:focus {
      outline: none;
      border-color: var(--primary);
      background: var(--surface);
      box-shadow: 0 0 0 3px rgba(79, 70, 229, .12);
    }

    button {
      cursor: pointer;
      font-weight: 700;
      border: none;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: 7px;
      transition: transform .12s ease, box-shadow .2s ease, filter .2s ease;
    }
    button .icon { width: 15px; height: 15px; }
    button.btn-primary { background: linear-gradient(135deg, var(--primary), var(--primary-dark)); color: #fff; box-shadow: 0 6px 16px -6px rgba(79, 70, 229, .55); }
    button.btn-success { background: linear-gradient(135deg, #22c55e, var(--success)); color: #fff; box-shadow: 0 6px 16px -6px rgba(22, 163, 74, .5); }
    button.btn-danger { background: linear-gradient(135deg, #ef4444, var(--danger)); color: #fff; box-shadow: 0 6px 16px -6px rgba(220, 38, 38, .45); }
    button.btn-icon-only { width: auto; padding: 9px; display: inline-flex; align-items: center; justify-content: center; }
    button.btn-icon-only .icon { margin: 0; }
    .setting-row { display: flex; align-items: center; justify-content: space-between; gap: 12px; margin: 8px 0 14px; }
    .setting-row label { margin: 0; }
    .setting-hint { color: var(--text-muted); font-size: 11px; margin-top: 3px; }
    button.btn-toggle { width: auto; min-width: 124px; margin: 0; padding: 10px 14px; }
    button.btn-toggle.is-on { background: linear-gradient(135deg, #22c55e, var(--success)); color: #fff; }
    button.btn-toggle.is-off { background: var(--surface-alt); color: var(--danger); border: 1px solid #fecaca; }
    button:hover { filter: brightness(1.06); transform: translateY(-1px); }
    button:active { transform: translateY(0) scale(.97); }

    table button {
      width: auto;
      padding: 6px 10px;
      margin: 2px 3px 2px 0;
      font-size: 11.5px;
      border-radius: 8px;
      box-shadow: none;
    }
    table button .icon { width: 13px; height: 13px; }

    /* ---------- Tables ---------- */
    .table-container { overflow-x: auto; border-radius: var(--radius-sm); border: 1px solid var(--border); }
    table { width: 100%; border-collapse: collapse; font-size: 12.5px; background: var(--surface); }
    th, td { padding: 11px 12px; text-align: left; border-bottom: 1px solid var(--border); }
    thead th {
      color: var(--text-muted);
      font-weight: 700;
      text-transform: uppercase;
      font-size: 10.5px;
      letter-spacing: .03em;
      background: var(--surface-alt);
      position: sticky; top: 0;
    }
    tbody tr { transition: background .15s ease; }
    tbody tr:hover { background: var(--primary-soft); }
    tbody tr:last-child td { border-bottom: none; }

    .clickable { cursor: pointer; text-decoration: none; color: var(--primary); font-weight: 600; border-bottom: 1.5px dashed rgba(79, 70, 229, .35); transition: color .15s ease; }
    .clickable:hover { color: var(--primary-dark); }

    /* ---------- Tab content transitions ---------- */
    .tab-content { display: none; }
    .tab-content.active { display: block; animation: fadeSlideIn .35s cubic-bezier(.22, 1, .36, 1); }
    @keyframes fadeSlideIn {
      from { opacity: 0; transform: translateY(8px); }
      to { opacity: 1; transform: translateY(0); }
    }

    code {
      background: var(--surface-alt);
      border: 1px solid var(--border);
      padding: 2px 6px;
      border-radius: 6px;
      font-size: 11.5px;
    }
  </style>
</head>
<body>
  <header>
    <div class="brand">
      <div class="brand-icon" id="brand-icon-slot"></div>
      <div>
        <h1>Admin Panel</h1>
        <div class="brand-sub">Bot control center</div>
      </div>
    </div>
    <span class="badge badge-success"><span class="pulse-dot"></span>Verified Admin</span>
  </header>

  <div class="nav-tabs">
      <button class="nav-tab active" onclick="switchTab('dashboard', this)" data-icon="layout-dashboard">Dashboard</button>
      <button class="nav-tab" onclick="switchTab('users', this)" data-icon="users">Users</button>
      <button class="nav-tab" onclick="switchTab('categories', this)" data-icon="tags">Categories</button>
      <button class="nav-tab" onclick="switchTab('services', this)" data-icon="shopping-bag">Services</button>
      <button class="nav-tab" onclick="switchTab('files', this)" data-icon="folder">Files</button>
      <button class="nav-tab" onclick="switchTab('reports', this)" data-icon="bar-chart">Reports</button>
      <button class="nav-tab" onclick="switchTab('withdrawals', this)" data-icon="wallet">Withdrawals</button>
      <button class="nav-tab" onclick="switchTab('payments', this)" data-icon="credit-card">Payment Methods</button>
      <button class="nav-tab" onclick="switchTab('broadcast', this)" data-icon="megaphone">Broadcast</button>
      <button class="nav-tab" onclick="switchTab('settings', this)" data-icon="settings">Settings</button>
  </div>

  <div class="container">
    <div id="tab-dashboard" class="tab-content active">
      <div class="grid">
        <div class="stat-card">
          <div class="stat-icon indigo" data-icon="users"></div>
          <div class="stat-label">Total Users</div><div id="st-users" class="stat-val">0</div>
        </div>
        <div class="stat-card">
          <div class="stat-icon cyan" data-icon="user-plus"></div>
          <div class="stat-label">Today's Users</div><div id="st-today-users" class="stat-val">0</div>
        </div>
        <div class="stat-card stat-card-clickable" onclick="openSubmittedIds()">
          <div class="stat-icon amber" data-icon="file-text"></div>
          <div class="stat-label">Submitted IDs</div><div id="st-ids" class="stat-val">0</div>
        </div>
        <div class="stat-card">
          <div class="stat-icon green" data-icon="banknote"></div>
          <div class="stat-label">Total Paid</div><div id="st-paid" class="stat-val">৳0</div>
        </div>
        <div class="stat-card">
          <div class="stat-icon red" data-icon="clock"></div>
          <div class="stat-label">Pending Payouts</div><div id="st-pending-w" class="stat-val">0</div>
        </div>
        <div class="stat-card stat-card-clickable" onclick="openDeveloperInfo()">
          <div class="stat-icon indigo" data-icon="code"></div>
          <div class="stat-label">Developer</div><div class="stat-val" style="font-size:13px;">Info</div>
        </div>
      </div>
    </div>

    <div id="ids-page" class="page-overlay" style="display:none;">
      <div class="page-header">
        <button class="page-back" onclick="closeSubmittedIds()" data-icon="arrow-left"></button>
        <h3><span class="icon" data-icon="file-text"></span>Submitted IDs by Service</h3>
      </div>
      <div id="ids-page-body" class="page-body">
        <div class="modal-loading">Loading…</div>
      </div>
    </div>

    <div id="dev-modal-overlay" class="dev-modal-overlay" style="display:none;" onclick="if(event.target===this) closeDeveloperInfo()">
      <div class="dev-modal-box">
        <button class="dev-modal-close" onclick="closeDeveloperInfo()" data-icon="x"></button>
        <div class="dev-modal-avatar" data-icon="code"></div>
        <div class="dev-modal-name">𝕋𝕒𝕣𝕚𝕜𝕦𝕝 𝔼𝕏𝔼 💻</div>
        <div class="dev-modal-rows">
          <div class="dev-modal-row"><span class="dev-modal-label"><span class="icon" data-icon="user-cog"></span>Name</span><span class="dev-modal-value">𝕋𝕒𝕣𝕚𝕜𝕦𝕝 𝔼𝕏𝔼 💻</span></div>
          <div class="dev-modal-row"><span class="dev-modal-label"><span class="icon" data-icon="clock"></span>Age</span><span class="dev-modal-value">14</span></div>
          <div class="dev-modal-row"><span class="dev-modal-label"><span class="icon" data-icon="landmark"></span>Country</span><span class="dev-modal-value">Bangladesh</span></div>
        </div>
        <button class="dev-modal-contact" onclick="contactDeveloper()" data-icon="send">Contact</button>
      </div>
    </div>


    <div id="tab-users" class="tab-content">
      <div class="card">
        <h3><span class="icon" data-icon="user-cog"></span>User Management</h3>
        <label>Grant admin access</label>
        <input type="number" id="new-admin-id" placeholder="Telegram ID to grant admin access">
        <button class="btn-primary" onclick="grantAdmin()" data-icon="shield-plus">Grant Admin Access</button>
        <div class="table-container">
          <table>
            <thead><tr><th>Admin ID</th><th>Action</th></tr></thead>
            <tbody id="admins-table-body"></tbody>
          </table>
        </div>
        <label>Search users</label>
        <input type="text" id="user-search" placeholder="🔍 Search name or Chat ID..." onkeyup="filterUsers()">
        <div class="table-container">
          <table>
            <thead><tr><th>Chat ID</th><th>Name</th><th>Balance</th><th>Actions</th></tr></thead>
            <tbody id="users-table-body"></tbody>
          </table>
        </div>
      </div>
    </div>

    <div id="tab-categories" class="tab-content">
      <div class="card">
        <h3><span class="icon" data-icon="folder-plus"></span>Add Category</h3>
        <input type="text" id="new-cat-name" placeholder="Category Name (e.g. Facebook)">
        <button class="btn-primary" onclick="addCategory()" data-icon="plus">Add Category</button>
      </div>
      <div class="card">
        <h3><span class="icon" data-icon="tags"></span>Categories List</h3>
        <div class="table-container">
          <table>
            <thead><tr><th>Name</th><th>Status</th><th>Action</th></tr></thead>
            <tbody id="cat-table-body"></tbody>
          </table>
        </div>
      </div>
    </div>

    <div id="tab-services" class="tab-content">
      <div class="card">
        <h3><span class="icon" data-icon="package-plus"></span>Add Service</h3>
        <label>Category</label>
        <select id="service-cat-select"></select>
        <input type="text" id="new-srv-name" placeholder="Service Name (e.g. 2FA Accounts)">
        <input type="number" step="0.1" id="new-srv-price" placeholder="Price per account (৳)">
        <button class="btn-primary" onclick="addService()" data-icon="plus">Add Service</button>
      </div>
      <div class="card">
        <h3><span class="icon" data-icon="shopping-bag"></span>Services List</h3>
        <div class="table-container">
          <table>
            <thead><tr><th>Category</th><th>Service</th><th>Price</th><th>Status</th><th>Action</th></tr></thead>
            <tbody id="srv-table-body"></tbody>
          </table>
        </div>
      </div>
    </div>

    <div id="tab-files" class="tab-content">
      <div class="card">
        <h3><span class="icon" data-icon="folder-open"></span>Aggregated Submitted Files</h3>
        <button class="btn-danger" onclick="deleteAllFiles()" data-icon="trash">Delete All Saved Files</button>
        <div class="table-container">
          <table>
            <thead><tr><th>Category</th><th>Service</th><th>Pending IDs</th><th>Download</th></tr></thead>
            <tbody id="files-table-body"></tbody>
          </table>
        </div>
      </div>
    </div>

    <div id="tab-reports" class="tab-content">
      <div class="card">
        <h3><span class="icon" data-icon="target"></span>Result Matching & Auto-Payout</h3>
        <label>1. Select Category:</label>
        <select id="report-cat-select" onchange="loadReportServices()"></select>
        <label>2. Select Service:</label>
        <select id="report-srv-select"></select>
        <label>3. Paste Successful Result IDs (one per line):</label>
        <textarea id="report-ids" rows="6" placeholder="Paste account IDs here..."></textarea>
        <button class="btn-success" onclick="processMatching()" data-icon="zap">Run Matching & Auto Pay</button>
      </div>
    </div>

    <div id="tab-withdrawals" class="tab-content">
      <div class="card">
        <h3><span class="icon" data-icon="wallet"></span>Withdrawal Requests</h3>
        <div class="table-container">
          <table>
            <thead><tr><th>User</th><th>Method</th><th>Account</th><th>Amount</th><th>Status</th><th>Actions</th></tr></thead>
            <tbody id="withdrawals-table-body"></tbody>
          </table>
        </div>
      </div>
    </div>

    <div id="tab-payments" class="tab-content">
      <div class="card">
        <h3><span class="icon" data-icon="credit-card"></span>Add Payment Method</h3>
        <input type="text" id="new-payment-name" placeholder="e.g. bKash, Nagad, Rocket">
        <button class="btn-primary" onclick="addPaymentMethod()" data-icon="plus">Add Payment Method</button>
      </div>
      <div class="card">
        <h3><span class="icon" data-icon="landmark"></span>Payment Methods</h3>
        <div class="table-container"><table>
          <thead><tr><th>Name</th><th>Status</th><th>Action</th></tr></thead>
          <tbody id="payment-table-body"></tbody>
        </table></div>
      </div>
    </div>

    <div id="tab-broadcast" class="tab-content">
      <div class="card">
        <h3><span class="icon" data-icon="megaphone"></span>Send Broadcast</h3>
        <label>Photo URL (optional)</label>
        <input type="text" id="bc-photo" placeholder="Optional Photo URL">
        <label>Message</label>
        <textarea id="bc-text" rows="4" placeholder="Message text..."></textarea>
        <button class="btn-primary" onclick="sendBroadcast()" data-icon="send">Send Broadcast</button>
      </div>
    </div>

    <div id="tab-settings" class="tab-content">
      <div class="card">
        <h3><span class="icon" data-icon="settings"></span>Bot & Force Join Settings</h3>
        <label>Support Link:</label>
        <input type="text" id="set-support-link">
        <button class="btn-primary" onclick="saveSupportLink()" data-icon="save">Save Support Link</button>
        <label>Force Join Channel (Username or ID):</label>
        <input type="text" id="set-fj-channel" placeholder="@channel">
        <label>Force Join Link:</label>
        <input type="text" id="set-fj-link" placeholder="https://t.me/...">
        <div class="setting-row">
          <div>
            <label>Force Join</label>
            <div class="setting-hint">Require channel membership before using the bot.</div>
          </div>
          <button type="button" id="force-join-toggle" class="btn-toggle is-off" onclick="toggleForceJoin()">Force Join: OFF</button>
        </div>
        <label>Custom Link 1 Title & URL:</label>
        <input type="text" id="set-c1-title">
        <input type="text" id="set-c1-url">
        <label>Custom Link 2 Title & URL:</label>
        <input type="text" id="set-c2-title">
        <input type="text" id="set-c2-url">
        <button class="btn-primary" onclick="saveSettings()" data-icon="save">Save Settings</button>
      </div>
    </div>
  </div>

  <script>
    /* ---------- Self-contained inline SVG icon set (no external CDN) ---------- */
    const ICONS = {
      'arrow-left': '<line x1="19" y1="12" x2="5" y2="12"/><polyline points="12 19 5 12 12 5"/>',
      'shield-check': '<path d="M12 3l7 3v6c0 4.5-3 7.5-7 9-4-1.5-7-4.5-7-9V6l7-3z"/><path d="M9 12l2 2 4-4"/>',
      'layout-dashboard': '<rect x="3" y="3" width="7" height="9" rx="1.5"/><rect x="14" y="3" width="7" height="5" rx="1.5"/><rect x="14" y="12" width="7" height="9" rx="1.5"/><rect x="3" y="16" width="7" height="5" rx="1.5"/>',
      'users': '<circle cx="9" cy="8" r="3"/><path d="M3 20c0-3.5 2.7-6 6-6s6 2.5 6 6"/><circle cx="17.5" cy="9" r="2.3"/><path d="M15.7 14.2c2.5.5 4.3 2.5 4.3 5.3"/>',
      'tags': '<path d="M11 3H5a2 2 0 00-2 2v6l9 9 8-8-9-9z"/><circle cx="7.5" cy="7.5" r="1.3" fill="currentColor" stroke="none"/>',
      'shopping-bag': '<path d="M6 8h12l-1 12H7L6 8z"/><path d="M9 8V6a3 3 0 016 0v2"/>',
      'folder': '<path d="M3 6a2 2 0 012-2h4l2 2h8a2 2 0 012 2v9a2 2 0 01-2 2H5a2 2 0 01-2-2V6z"/>',
      'folder-open': '<path d="M3 7a2 2 0 012-2h4l2 2h8a2 2 0 012 2H8l-2 8H3V7z"/><path d="M6 17l2-7h13l-2 7H6z"/>',
      'folder-plus': '<path d="M3 6a2 2 0 012-2h4l2 2h8a2 2 0 012 2v9a2 2 0 01-2 2H5a2 2 0 01-2-2V6z"/><line x1="12" y1="10.5" x2="12" y2="15.5"/><line x1="9.5" y1="13" x2="14.5" y2="13"/>',
      'bar-chart': '<line x1="5" y1="20" x2="5" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="19" y1="20" x2="19" y2="14"/>',
      'wallet': '<path d="M3 7a2 2 0 012-2h12a2 2 0 012 2v10a2 2 0 01-2 2H5a2 2 0 01-2-2V7z"/><path d="M16.5 12h2.5v3h-2.5a1.5 1.5 0 010-3z"/>',
      'credit-card': '<rect x="2" y="5" width="20" height="14" rx="2"/><line x1="2" y1="10" x2="22" y2="10"/>',
      'megaphone': '<path d="M3 10v4h3l6 4V6l-6 4H3z"/><path d="M14 9a4 4 0 010 6"/>',
      'settings': '<circle cx="12" cy="12" r="3"/><path d="M12 3v2.2M12 18.8V21M21 12h-2.2M5.2 12H3M18.4 5.6l-1.5 1.5M7.1 16.9l-1.5 1.5M18.4 18.4l-1.5-1.5M7.1 7.1L5.6 5.6"/>',
      'user-plus': '<circle cx="9" cy="8" r="3"/><path d="M3 20c0-3.5 2.7-6 6-6s6 2.5 6 6"/><line x1="19" y1="8" x2="19" y2="14"/><line x1="16" y1="11" x2="22" y2="11"/>',
      'user-minus': '<circle cx="9" cy="8" r="3"/><path d="M3 20c0-3.5 2.7-6 6-6s6 2.5 6 6"/><line x1="16" y1="11" x2="22" y2="11"/>',
      'user-cog': '<circle cx="8.5" cy="8" r="3"/><path d="M2.5 20c0-3.5 2.7-6 6-6s6 2.5 6 6"/><circle cx="18" cy="15.5" r="2.2"/><path d="M18 12.3v1M18 17.7v1M15.3 15.5h1M19.7 15.5h1"/>',
      'shield-plus': '<path d="M12 3l7 3v6c0 4.5-3 7.5-7 9-4-1.5-7-4.5-7-9V6l7-3z"/><line x1="12" y1="9" x2="12" y2="14"/><line x1="9.5" y1="11.5" x2="14.5" y2="11.5"/>',
      'shield-alert': '<path d="M12 3l7 3v6c0 4.5-3 7.5-7 9-4-1.5-7-4.5-7-9V6l7-3z"/><line x1="12" y1="8" x2="12" y2="12.5"/><circle cx="12" cy="15.6" r="0.7" fill="currentColor" stroke="none"/>',
      'file-text': '<path d="M7 3h7l4 4v14a1 1 0 01-1 1H7a1 1 0 01-1-1V4a1 1 0 011-1z"/><line x1="9" y1="12" x2="15" y2="12"/><line x1="9" y1="16" x2="15" y2="16"/>',
      'banknote': '<rect x="2" y="6" width="20" height="12" rx="2"/><circle cx="12" cy="12" r="3"/><line x1="6" y1="9" x2="6" y2="9.01"/><line x1="18" y1="15" x2="18" y2="15.01"/>',
      'clock': '<circle cx="12" cy="12" r="9"/><polyline points="12 7 12 12 16 14"/>',
      'hourglass': '<path d="M6 3h12M6 21h12M7 3c0 5 4 7 5 8-1 1-5 3-5 8M17 3c0 5-4 7-5 8 1 1 5 3 5 8"/>',
      'package-plus': '<path d="M3 7l9-4 9 4-9 4-9-4z"/><path d="M3 7v10l9 4 9-4V7"/><line x1="12" y1="11" x2="12" y2="21"/><line x1="18" y1="3" x2="18" y2="7"/><line x1="16" y1="5" x2="20" y2="5"/>',
      'plus': '<line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/>',
      'trash': '<line x1="4" y1="7" x2="20" y2="7"/><path d="M6 7l1 13a2 2 0 002 2h6a2 2 0 002-2l1-13"/><path d="M9 7V4a1 1 0 011-1h4a1 1 0 011 1v3"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/>',
      'target': '<circle cx="12" cy="12" r="8"/><circle cx="12" cy="12" r="4.5"/><circle cx="12" cy="12" r="1" fill="currentColor" stroke="none"/>',
      'zap': '<polygon points="13 2 4 14 11 14 10 22 20 9 13 9 13 2" fill="currentColor" stroke="none"/>',
      'send': '<path d="M3 11l18-8-8 18-3-8-7-2z"/><line x1="21" y1="3" x2="11" y2="13"/>',
      'save': '<path d="M5 3h11l4 4v13a1 1 0 01-1 1H5a1 1 0 01-1-1V4a1 1 0 011-1z"/><path d="M8 3v5h7V3"/><rect x="7" y="13" width="10" height="7"/>',
      'unlock': '<rect x="4" y="11" width="16" height="9" rx="2"/><path d="M7 11V7a5 5 0 019-3"/>',
      'lock': '<rect x="4" y="11" width="16" height="9" rx="2"/><path d="M7 11V7a5 5 0 0110 0v4"/>',
      'eye': '<path d="M2 12s4-7 10-7 10 7 10 7-4 7-10 7-10-7-10-7z"/><circle cx="12" cy="12" r="3"/>',
      'eye-off': '<path d="M3 3l18 18"/><path d="M10.6 5.2C11 5.1 11.5 5 12 5c6 0 10 7 10 7a17.7 17.7 0 01-4 4.7M6.5 6.6C4 8.3 2 12 2 12s4 7 10 7c1.4 0 2.7-.3 3.9-.8"/><path d="M9.5 9.7a3 3 0 004.3 4.2"/>',
      'tag': '<path d="M11 3H5a2 2 0 00-2 2v6l9 9 8-8-9-9z"/><circle cx="7.5" cy="7.5" r="1.3" fill="currentColor" stroke="none"/>',
      'download': '<path d="M12 3v12"/><polyline points="7 11 12 16 17 11"/><path d="M5 19h14"/>',
      'check': '<polyline points="4 12 9 17 20 6"/>',
      'x': '<line x1="5" y1="5" x2="19" y2="19"/><line x1="19" y1="5" x2="5" y2="19"/>',
      'landmark': '<line x1="4" y1="21" x2="20" y2="21"/><polygon points="12 3 21 8 3 8"/><line x1="6" y1="10" x2="6" y2="18"/><line x1="10" y1="10" x2="10" y2="18"/><line x1="14" y1="10" x2="14" y2="18"/><line x1="18" y1="10" x2="18" y2="18"/>',
      'code': '<polyline points="8 6 2 12 8 18"/><polyline points="16 6 22 12 16 18"/>'
    };

    function I(name) {
      const inner = ICONS[name];
      if (!inner) return '';
      return '<svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">' + inner + '</svg>';
    }

    // Renders icons into any static element carrying data-icon="name",
    // and injects the button icon before its existing label text.
    function renderIcons(root) {
      (root || document).querySelectorAll('[data-icon]').forEach(el => {
        const name = el.getAttribute('data-icon');
        if (el.dataset.iconDone) return;
        if (el.tagName === 'BUTTON' || el.tagName === 'SPAN') {
          el.insertAdjacentHTML('afterbegin', I(name));
        } else {
          el.innerHTML = I(name);
        }
        el.dataset.iconDone = '1';
      });
    }

    const tg = window.Telegram?.WebApp;
    if (tg) { tg.ready(); tg.expand(); }
     const initData = tg?.initData ||
       new URLSearchParams(window.location.search).get('tgWebAppData') || "";

    // smooth count-up animation for dashboard numbers
    function animateValue(el, endVal, isCurrency) {
      const startVal = parseFloat((el.innerText || '0').replace(/[^\d.-]/g, '')) || 0;
      const duration = 650;
      const startTime = performance.now();
      function step(now) {
        const progress = Math.min((now - startTime) / duration, 1);
        const eased = 1 - Math.pow(1 - progress, 3);
        const current = startVal + (endVal - startVal) * eased;
        el.innerText = isCurrency ? ('৳' + current.toFixed(2)) : Math.round(current);
        if (progress < 1) requestAnimationFrame(step);
      }
      requestAnimationFrame(step);
    }

    async function req(url, options = {}) {
      options.headers = options.headers || {};
      options.headers['X-Telegram-Init-Data'] = initData;
      const res = await fetch(url, options);
      if (res.status === 401 || res.status === 403) {
        document.body.innerHTML = `
          <div style="padding:60px 24px;text-align:center;">
            <div style="width:56px;height:56px;border-radius:16px;background:var(--danger-soft);display:flex;align-items:center;justify-content:center;margin:0 auto 16px;color:var(--danger);font-size:26px;">${I('shield-alert')}</div>
            <h2 style="color:#dc2626;margin-bottom:8px;font-family:'Plus Jakarta Sans',sans-serif;">Unauthorized Access</h2>
            <p style="color:#64748b;font-size:14px;">Only the configured bot admin can access this panel.</p>
          </div>`;
        throw new Error("Unauthorized");
      }
      return res.json();
    }

     function switchTab(name, tabButton) {
      document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
      document.querySelectorAll('.nav-tab').forEach(el => el.classList.remove('active'));
      document.getElementById('tab-' + name).classList.add('active');
       tabButton.classList.add('active');
       tabButton.scrollIntoView({ behavior: 'smooth', inline: 'center', block: 'nearest' });

      if (name === 'dashboard') loadDashboard();
      if (name === 'users') loadUsers();
      if (name === 'categories') loadCategories();
      if (name === 'services') loadServices();
      if (name === 'files') loadFiles();
      if (name === 'reports') loadReportsData();
      if (name === 'withdrawals') loadWithdrawals();
       if (name === 'payments') loadPaymentMethods();
      if (name === 'settings') loadSettings();
    }

    async function loadDashboard() {
      const d = await req('/api/dashboard/stats');
      animateValue(document.getElementById('st-users'), d.total_users, false);
      animateValue(document.getElementById('st-today-users'), d.today_users, false);
      animateValue(document.getElementById('st-ids'), d.total_submitted_ids, false);
      animateValue(document.getElementById('st-paid'), d.total_paid, true);
      animateValue(document.getElementById('st-pending-w'), d.pending_withdrawals, false);
    }

    async function openSubmittedIds() {
      const page = document.getElementById('ids-page');
      const body = document.getElementById('ids-page-body');
      page.style.display = 'flex';
      body.innerHTML = '<div class="modal-loading">Loading…</div>';
      try {
        const list = await req('/api/files/summary');
        if (!list.length) {
          body.innerHTML = '<div class="svc-empty">No services yet. Add a service first.</div>';
          return;
        }
        body.innerHTML = list.map(f => `
          <div class="svc-block">
            <div class="svc-head" onclick="toggleSvcIds(this, ${f.service_id})">
              <div class="svc-head-info">
                <div class="svc-head-cat">${f.category_name}</div>
                <div class="svc-head-name">${f.service_name}</div>
              </div>
              <div class="svc-head-count">${f.total_submitted_count} IDs</div>
            </div>
            <div class="svc-ids-panel" id="svc-ids-${f.service_id}"></div>
          </div>
        `).join('');
      } catch (e) {
        body.innerHTML = '<div class="svc-empty">Failed to load. Try again.</div>';
      }
    }

    function closeSubmittedIds() {
      document.getElementById('ids-page').style.display = 'none';
    }

    const DEVELOPER_CONTACT_URL = 'https://t.me/rciasin';

    function openDeveloperInfo() {
      document.getElementById('dev-modal-overlay').style.display = 'flex';
    }

    function closeDeveloperInfo() {
      document.getElementById('dev-modal-overlay').style.display = 'none';
    }

    function contactDeveloper() {
      // মিনি অ্যাপ বন্ধ করে ডেভেলপারের টেলিগ্রাম আইডিতে নিয়ে যায়
      if (tg && tg.openTelegramLink) {
        tg.openTelegramLink(DEVELOPER_CONTACT_URL);
        if (tg.close) tg.close();
      } else {
        window.open(DEVELOPER_CONTACT_URL, '_blank');
      }
    }

    async function toggleSvcIds(headEl, serviceId) {
      const panel = document.getElementById('svc-ids-' + serviceId);
      const isOpen = panel.classList.contains('open');
      if (isOpen) {
        panel.classList.remove('open');
        return;
      }
      panel.classList.add('open');
      if (!panel.dataset.loaded) {
        panel.innerHTML = '<div class="modal-loading">Loading…</div>';
        try {
          const ids = await req('/api/files/ids?service_id=' + serviceId);
          panel.innerHTML = ids.length
            ? `<div class="svc-ids-list">${ids.map(id => `<span class="svc-id-chip">${id}</span>`).join('')}</div>`
            : '<div class="svc-empty">No open IDs for this service yet.</div>';
          panel.dataset.loaded = '1';
        } catch (e) {
          panel.innerHTML = '<div class="svc-empty">Failed to load IDs.</div>';
        }
      }
    }

    let allUsers = [];
    async function loadUsers() {
      allUsers = await req('/api/users');
      renderUsers(allUsers);
       loadAdmins();
    }
    function renderUsers(list) {
      document.getElementById('users-table-body').innerHTML = list.map(u => `
        <tr>
          <td><span class="clickable" onclick="copyId('${u.telegram_id}')">${u.telegram_id}</span></td>
          <td>${u.first_name || ''} <small>(${u.username ? '@'+u.username : ''})</small></td>
          <td>৳${u.balance.toFixed(2)}</td>
          <td>
            <button class="${u.is_blocked ? 'btn-success' : 'btn-danger'}" onclick="toggleBlock(${u.telegram_id})">${I(u.is_blocked ? 'unlock' : 'lock')}${u.is_blocked ? 'Unblock' : 'Block'}</button>
            <button class="btn-primary" onclick="adjustBalance(${u.telegram_id})">${I('banknote')}±৳</button>
          </td>
        </tr>
      `).join('');
    }
    function filterUsers() {
      const q = document.getElementById('user-search').value.toLowerCase();
      renderUsers(allUsers.filter(u => (u.first_name && u.first_name.toLowerCase().includes(q)) || String(u.telegram_id).includes(q)));
    }
    function copyId(id) {
      navigator.clipboard.writeText(id);
      alert('Copied Chat ID: ' + id);
    }
    async function toggleBlock(id) {
      await req('/api/user/action', { method:'POST', body: JSON.stringify({ action:'toggle_block', user_id: id }) });
      loadUsers();
    }
    async function adjustBalance(id) {
      const amt = prompt("Enter amount to add or deduct (e.g. 50 or -50):");
      if (amt !== null) {
        await req('/api/user/action', { method:'POST', body: JSON.stringify({ action:'adjust_balance', user_id: id, amount: parseFloat(amt) }) });
        loadUsers();
      }
    }
     async function grantAdmin() {
       const telegram_id = document.getElementById('new-admin-id').value;
       if (!telegram_id) return alert('Enter a Telegram ID');
       await req('/api/admins', { method:'POST', body: JSON.stringify({ telegram_id }) });
       document.getElementById('new-admin-id').value = '';
       alert('Admin access granted');
       loadAdmins();
      }
      async function loadAdmins() {
        const admins = await req('/api/admins');
        document.getElementById('admins-table-body').innerHTML = admins.map(a => `
          <tr>
            <td>${a.telegram_id}</td>
            <td><button class="btn-danger"
              onclick="removeAdmin(${a.telegram_id})">${I('user-minus')}Remove Admin</button></td>
          </tr>`).join('');
      }
      async function removeAdmin(telegram_id) {
        if (!confirm('Remove admin access for this Telegram ID?')) return;
        await req('/api/admins', { method:'DELETE', body: JSON.stringify({ telegram_id }) });
        loadAdmins();
      }

    async function loadCategories() {
      const cats = await req('/api/categories');
      document.getElementById('cat-table-body').innerHTML = cats.map(c => `
        <tr>
          <td>${c.name}</td>
          <td><span class="badge ${c.is_active ? 'badge-success' : 'badge-danger'}">${c.is_active ? 'Active' : 'Disabled'}</span></td>
          <td>
            <button class="btn-danger" onclick="deleteCategory(${c.id})">${I('trash')}Delete</button>
          </td>
        </tr>
      `).join('');
    }
    async function addCategory() {
      const name = document.getElementById('new-cat-name').value;
      if (!name) return alert('Enter category name');
      await req('/api/categories', { method:'POST', body: JSON.stringify({ name }) });
      document.getElementById('new-cat-name').value = '';
      loadCategories();
    }
    async function deleteCategory(id) {
      if (confirm('Delete category and all its services?')) {
        await req('/api/categories?id=' + id, { method:'DELETE' });
        loadCategories();
      }
    }

    let globalServices = [];
    async function loadServices() {
      const cats = await req('/api/categories');
      document.getElementById('service-cat-select').innerHTML = cats.map(c => `<option value="${c.id}">${c.name}</option>`).join('');
      globalServices = await req('/api/services');
      document.getElementById('srv-table-body').innerHTML = globalServices.map(s => `
        <tr>
          <td>${s.category_name || '-'}</td>
          <td>${s.name}</td>
          <td>৳${s.price.toFixed(2)}</td>
          <td><span class="badge ${s.is_active ? 'badge-success' : 'badge-danger'}">${s.is_active ? 'Active' : 'Disabled'}</span></td>
          <td>
            <button class="btn-primary" onclick="toggleService(${s.id}, ${s.is_active ? 0 : 1})">${I(s.is_active ? 'eye-off' : 'eye')}${s.is_active ? 'Disable' : 'Enable'}</button>
            <button class="btn-primary" onclick="editServicePrice(${s.id})">${I('tag')}Price</button>
            <button class="btn-danger" onclick="deleteService(${s.id})">${I('trash')}Delete</button>
          </td>
        </tr>
      `).join('');
    }
    async function addService() {
      const category_id = document.getElementById('service-cat-select').value;
      const name = document.getElementById('new-srv-name').value;
      const price = document.getElementById('new-srv-price').value;
      if (!name || !price) return alert('Enter name & price');
      await req('/api/services', { method:'POST', body: JSON.stringify({ category_id, name, price }) });
      document.getElementById('new-srv-name').value = '';
      document.getElementById('new-srv-price').value = '';
      loadServices();
    }
    async function toggleService(id, status) {
      await req('/api/services', { method:'PUT', body: JSON.stringify({ id, is_active: status }) });
      loadServices();
    }
    async function editServicePrice(id) {
      const p = prompt("Enter new price per account (৳):");
      if (p !== null) {
        await req('/api/services', { method:'PUT', body: JSON.stringify({ id, price: parseFloat(p) }) });
        loadServices();
      }
    }
    async function deleteService(id) {
      if (confirm('Delete service?')) {
        await req('/api/services?id=' + id, { method:'DELETE' });
        loadServices();
      }
    }

    async function loadFiles() {
      const list = await req('/api/files/summary');
      document.getElementById('files-table-body').innerHTML = list.map(f => `
        <tr>
          <td>${f.category_name}</td>
          <td>${f.service_name}</td>
          <td>${f.unmatched_count}</td>
          <td><button class="btn-primary btn-icon-only" title="Download CSV" onclick="window.open('/api/files/download?service_id=${f.service_id}&initData='+encodeURIComponent(initData))">${I('download')}</button></td>
        </tr>
      `).join('');
    }
     async function deleteAllFiles() {
       if (!confirm('Delete every saved submission file and all stored account rows? This cannot be undone.')) return;
       await req('/api/files/delete-all', { method:'DELETE' });
       alert('All saved files and account rows were deleted.');
       loadFiles();
       loadDashboard();
     }

    async function loadReportsData() {
      const cats = await req('/api/categories');
      document.getElementById('report-cat-select').innerHTML = '<option value="">-- Select Category --</option>' + cats.map(c => `<option value="${c.id}">${c.name}</option>`).join('');
      globalServices = await req('/api/services');
    }
    function loadReportServices() {
      const catId = document.getElementById('report-cat-select').value;
      const filtered = globalServices.filter(s => String(s.category_id) === String(catId));
      document.getElementById('report-srv-select').innerHTML = filtered.map(s => `<option value="${s.id}">${s.name} (৳${s.price.toFixed(2)})</option>`).join('');
    }
    async function processMatching() {
      const service_id = document.getElementById('report-srv-select').value;
      const results_text = document.getElementById('report-ids').value;
      if (!service_id || !results_text) return alert('Select service and paste result IDs!');
      if (!confirm('Run matching and pay matched users automatically?')) return;

      const res = await req('/api/reports/match', { method:'POST', body: JSON.stringify({ service_id, results_text }) });
      alert(`✅ Complete! Matched: ${res.matched_count} IDs across ${res.users_credited} users.`);
      document.getElementById('report-ids').value = '';
    }

    async function loadWithdrawals() {
      const list = await req('/api/withdrawals');
      document.getElementById('withdrawals-table-body').innerHTML = list.map(w => `
        <tr>
          <td>${w.first_name || 'User'} <small>(${w.telegram_id})</small></td>
          <td>${w.method}</td>
          <td><code>${w.account_number}</code></td>
          <td>৳${w.amount.toFixed(2)}</td>
          <td><span class="badge ${w.status==='Paid'?'badge-success':(w.status==='Pending'?'badge-warning':'badge-danger')}">${w.status}</span></td>
          <td>
            ${w.status==='Pending' ? `
              <button class="btn-success" onclick="updateWithdrawal(${w.id}, 'Paid')">${I('check')}Paid</button>
              <button class="btn-danger" onclick="updateWithdrawal(${w.id}, 'Rejected')">${I('x')}Reject</button>
            ` : '-'}
          </td>
        </tr>
      `).join('');
    }
    async function updateWithdrawal(id, status) {
      await req('/api/withdrawals', { method:'POST', body: JSON.stringify({ id, status }) });
      loadWithdrawals();
    }

     async function loadPaymentMethods() {
       const list = await req('/api/payment-methods');
       document.getElementById('payment-table-body').innerHTML = list.map(p => `
         <tr>
           <td>${p.name}</td>
           <td><span class="badge ${p.is_active ? 'badge-success' : 'badge-danger'}">${p.is_active ? 'Active' : 'Disabled'}</span></td>
           <td>
             <button class="${p.is_active ? 'btn-danger' : 'btn-success'}"
               onclick="togglePaymentMethod(${p.id}, ${p.is_active ? 0 : 1})">${I(p.is_active ? 'eye-off' : 'eye')}${p.is_active ? 'Disable' : 'Enable'}</button>
             <button class="btn-danger" onclick="deletePaymentMethod(${p.id})">${I('trash')}Delete</button>
           </td>
         </tr>`).join('');
     }
     async function addPaymentMethod() {
       const name = document.getElementById('new-payment-name').value.trim();
       if (!name) return alert('Enter a payment method name');
       await req('/api/payment-methods', { method:'POST', body: JSON.stringify({ name }) });
       document.getElementById('new-payment-name').value = '';
       loadPaymentMethods();
     }
     async function togglePaymentMethod(id, is_active) {
       await req('/api/payment-methods', { method:'PUT', body: JSON.stringify({ id, is_active }) });
       loadPaymentMethods();
     }
     async function deletePaymentMethod(id) {
       if (confirm('Delete this payment method?')) {
         await req('/api/payment-methods', { method:'DELETE', body: JSON.stringify({ id }) });
         loadPaymentMethods();
       }
     }

    async function sendBroadcast() {
      const photo_url = document.getElementById('bc-photo').value;
      const text = document.getElementById('bc-text').value;
      if (!text && !photo_url) return alert('Enter message!');
      if (!confirm('Send broadcast to all users?')) return;

      const res = await req('/api/broadcast', { method:'POST', body: JSON.stringify({ photo_url, text }) });
      alert(`Broadcast sent: ${res.success}, Failed/Blocked: ${res.failed}`);
      document.getElementById('bc-text').value = '';
    }

    let forceJoinEnabled = '0';

    function renderForceJoinToggle() {
      const button = document.getElementById('force-join-toggle');
      if (!button) return;
      const enabled = forceJoinEnabled === '1';
      button.textContent = enabled ? 'Force Join: ON' : 'Force Join: OFF';
      button.classList.toggle('is-on', enabled);
      button.classList.toggle('is-off', !enabled);
    }

    async function saveSupportLink() {
      const supportLink = document.getElementById('set-support-link').value.trim();
      if (!supportLink) return alert('Enter a support link');
      await req('/api/settings', {
        method: 'POST',
        body: JSON.stringify({ support_link: supportLink }),
      });
      alert('✅ Support link saved successfully!');
    }

    async function toggleForceJoin() {
      const previous = forceJoinEnabled;
      forceJoinEnabled = previous === '1' ? '0' : '1';
      renderForceJoinToggle();
      try {
        await req('/api/settings', {
          method: 'POST',
          body: JSON.stringify({ force_join_enabled: forceJoinEnabled }),
        });
        alert(forceJoinEnabled === '1' ? '✅ Force Join enabled!' : '✅ Force Join disabled!');
      } catch (error) {
        forceJoinEnabled = previous;
        renderForceJoinToggle();
      }
    }

    async function loadSettings() {
      const s = await req('/api/settings');
      document.getElementById('set-support-link').value = s.support_link || '';
      document.getElementById('set-fj-channel').value = s.force_join_channel || '';
      document.getElementById('set-fj-link').value = s.force_join_link || '';
      forceJoinEnabled = s.force_join_enabled === '1' ? '1' : '0';
      renderForceJoinToggle();
      document.getElementById('set-c1-title').value = s.custom_link_1_title || '';
      document.getElementById('set-c1-url').value = s.custom_link_1_url || '';
      document.getElementById('set-c2-title').value = s.custom_link_2_title || '';
      document.getElementById('set-c2-url').value = s.custom_link_2_url || '';
    }
    async function saveSettings() {
      const payload = {
        support_link: document.getElementById('set-support-link').value,
        force_join_channel: document.getElementById('set-fj-channel').value,
        force_join_link: document.getElementById('set-fj-link').value,
        force_join_enabled: forceJoinEnabled,
        custom_link_1_title: document.getElementById('set-c1-title').value,
        custom_link_1_url: document.getElementById('set-c1-url').value,
        custom_link_2_title: document.getElementById('set-c2-title').value,
        custom_link_2_url: document.getElementById('set-c2-url').value,
      };
      await req('/api/settings', { method:'POST', body: JSON.stringify(payload) });
      alert('✅ Settings saved successfully!');
    }

    document.getElementById('brand-icon-slot').innerHTML = I('shield-check');
    renderIcons();
    loadDashboard();
  </script>
</body>
</html>"""

async def handle_webapp_index(request: web.Request):
    return web.Response(text=MINI_APP_HTML, content_type="text/html")

async def _shutdown_with_final_backup():
    logger.info("🛑 Shutdown signal received — pushing final backup to GitHub...")
    try:
        await github_push_backup(force=True)
    finally:
        os._exit(0)

# ==========================================
# 🚀 MAIN APPLICATION ENTRYPOINT
# ==========================================
async def main():
    pending_backup_json = await github_pull_backup()  # GitHub থেকে JSON আনা হলো, schema তৈরির আগে
    await init_db()
    if pending_backup_json:
        await asyncio.to_thread(_import_json_into_db_sync, pending_backup_json)
        logger.info("☁️ Restored data from GitHub backup into the database.")

    # শাটডাউন/redeploy হওয়ার ঠিক আগে শেষ ব্যাকআপটাও GitHub-এ push করে নেওয়া হয়,
    # যাতে সবশেষ periodic push-এর পরের কোনো ডেটাও হারিয়ে না যায়।
    loop = asyncio.get_running_loop()
    for sig in (signal.SIGTERM, signal.SIGINT):
        try:
            loop.add_signal_handler(
                sig,
                lambda: asyncio.create_task(_shutdown_with_final_backup())
            )
        except NotImplementedError:
            pass  # some platforms (e.g. Windows) don't support add_signal_handler

    asyncio.create_task(github_backup_loop())
    if GITHUB_ENABLED:
        logger.info(f"☁️ GitHub auto-backup enabled → {GITHUB_REPO}@{GITHUB_BRANCH}/{GITHUB_BACKUP_PATH} (every {GITHUB_BACKUP_INTERVAL}s)")
    else:
        logger.info("☁️ GitHub auto-backup disabled (set GITHUB_TOKEN & GITHUB_REPO env vars to enable).")

    # Web Server চালু করা
    app = web.Application()
    app.router.add_get("/", handle_webapp_index)
    app.router.add_get("/api/dashboard/stats", api_dashboard_stats)
    app.router.add_get("/api/users", api_get_users)
    app.router.add_route("*", "/api/admins", api_admins)
    app.router.add_route("*", "/api/payment-methods", api_payment_methods)
    app.router.add_post("/api/user/action", api_user_action)
    app.router.add_route("*", "/api/categories", api_categories)
    app.router.add_route("*", "/api/services", api_services)
    app.router.add_get("/api/files/summary", api_files_summary)
    app.router.add_get("/api/files/ids", api_files_ids)
    app.router.add_get("/api/files/download", api_files_download)
    app.router.add_delete("/api/files/delete-all", api_files_delete_all)
    app.router.add_post("/api/reports/match", api_reports_match)
    app.router.add_route("*", "/api/withdrawals", api_withdrawals)
    app.router.add_post("/api/broadcast", api_broadcast)
    app.router.add_route("*", "/api/settings", api_settings)

    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", PORT)
    await site.start()
    logger.info(f"🌐 Web Server started on port {PORT}")

    # ডুপ্লিকেট সেশন এবং পুরনো ওয়েবহুক ক্লিয়ার করা (TelegramConflictError সমাধান)
    await bot.delete_webhook(drop_pending_updates=True)
    # Remove the old top menu Mini App button. The reply-keyboard Admin Panel
    # button below is the only entry point now.
    try:
        await bot.set_chat_menu_button(chat_id=ADMIN_ID, menu_button=MenuButtonDefault())
    except Exception as e:
        logger.warning(f"Could not reset admin menu button: {e}")

    logger.info("🤖 Starting bot polling...")
    await dp.start_polling(bot, drop_pending_updates=True)

if __name__ == "__main__":
    asyncio.run(main())
