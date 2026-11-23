from keep_alive import keep_alive
keep_alive()


# -*- coding: utf-8 -*-
"""
Gmail Work BD - Telegram Micro Tasking Bot (Ultimate Edition)
Fully Optimized, High-Speed Broadcast, Auto-Delete Credentials, Multi-Channel Force Join, 
Dynamic Withdraw Control (ON/OFF & Min/Max), Instant Referral Alerts & Enhanced UI.
"""

import asyncio
import datetime
import html
import io
import logging
import os
import re
import sqlite3
import sys
import secrets
import random
import warnings
import time
from typing import Dict, List, Optional, Tuple, Union
from contextlib import contextmanager
from queue import Queue
import threading

import pyotp

# Try importing openpyxl for Excel spreadsheet generation
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

from telegram import (
    KeyboardButton,
    ReplyKeyboardMarkup,
    ReplyKeyboardRemove,
    Update,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
)
from telegram.constants import ParseMode
from telegram.warnings import PTBUserWarning
from telegram.error import Forbidden, BadRequest, RetryAfter, TelegramError
from telegram.ext import (
    Application,
    ApplicationBuilder,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    ConversationHandler,
    MessageHandler,
    filters,
)

warnings.filterwarnings("ignore", category=PTBUserWarning)

import inspect as _inspect

_KB_HAS_STYLE = "style" in _inspect.signature(KeyboardButton.__init__).parameters
_IKB_HAS_STYLE = "style" in _inspect.signature(InlineKeyboardButton.__init__).parameters

def kb_btn(text: str, style: str = "primary") -> KeyboardButton:
    if _KB_HAS_STYLE:
        return KeyboardButton(text, style=style)
    return KeyboardButton(text, api_kwargs={"style": style})

def ikb_btn(text: str, style: str = "primary", **kwargs) -> InlineKeyboardButton:
    if _IKB_HAS_STYLE:
        return InlineKeyboardButton(text, style=style, **kwargs)
    return InlineKeyboardButton(text, api_kwargs={"style": style}, **kwargs)

if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except AttributeError:
        pass

# -----------------------------------------------------------------------------
# LOGGING SETUP
# -----------------------------------------------------------------------------
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("bot.log", encoding="utf-8"),
    ],
)
logger = logging.getLogger("Gmail_Work_BD")

# -----------------------------------------------------------------------------
# DATABASE CONNECTION POOL & INIT
# -----------------------------------------------------------------------------
DB_FILE = os.getenv("DB_FILE", "bot_database.db")
BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
ADMIN_ID = (os.getenv("ADMIN_ID") or os.getenv("ADMIN_IDS") or "").strip()

class DatabasePool:
    _instance = None
    _lock = threading.Lock()

    def __new__(cls):
        if cls._instance is None:
            with cls._lock:
                if cls._instance is None:
                    cls._instance = super().__new__(cls)
                    cls._instance._initialized = False
        return cls._instance

    def __init__(self):
        if self._initialized:
            return
        self._initialized = True
        self._pool = Queue()
        self._max_connections = 50
        self._min_connections = 5
        self._create_initial_pool()
        logger.info(f"Database pool initialized with {self._min_connections} connections")

    def _create_initial_pool(self):
        for _ in range(self._min_connections):
            conn = self._create_connection()
            self._pool.put(conn)

    def _create_connection(self):
        conn = sqlite3.connect(DB_FILE, check_same_thread=False, timeout=30.0)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA synchronous=NORMAL")
        conn.execute("PRAGMA cache_size=10000")
        conn.execute("PRAGMA temp_store=MEMORY")
        return conn

    @contextmanager
    def get_connection(self):
        conn = None
        try:
            conn = self._pool.get(timeout=10.0)
            yield conn
        except Exception as e:
            if conn:
                self._pool.put(conn)
            raise e
        finally:
            if conn:
                self._pool.put(conn)

    def reload_pool(self):
        old_conns = []
        while not self._pool.empty():
            try:
                old_conns.append(self._pool.get_nowait())
            except Exception:
                break
        for c in old_conns:
            try:
                c.close()
            except Exception:
                pass
        self._create_initial_pool()
        logger.info("Database pool reloaded after restore.")

_db_pool = DatabasePool()

def get_db():
    return _db_pool.get_connection()

def get_db_direct():
    conn = sqlite3.connect(DB_FILE, check_same_thread=False, timeout=30.0)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    return conn

def init_db():
    with get_db_direct() as conn:
        cursor = conn.cursor()

        cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            user_id INTEGER PRIMARY KEY,
            username TEXT,
            full_name TEXT,
            balance REAL DEFAULT 0.0,
            referral_code TEXT UNIQUE,
            referred_by INTEGER,
            ref_count INTEGER DEFAULT 0,
            ref_earnings REAL DEFAULT 0.0,
            is_banned INTEGER DEFAULT 0,
            joined_at TIMESTAMP,
            language TEXT DEFAULT 'bn',
            custom_ref_fixed REAL DEFAULT NULL
        )
        """)
        try:
            cursor.execute("ALTER TABLE users ADD COLUMN custom_ref_fixed REAL DEFAULT NULL")
        except sqlite3.OperationalError:
            pass
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_users_referral_code ON users(referral_code)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_users_referred_by ON users(referred_by)")

        cursor.execute("""
        CREATE TABLE IF NOT EXISTS task_submissions (
            submission_id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            task_type TEXT,
            submitted_username TEXT,
            submitted_password TEXT,
            secret_key_2fa TEXT,
            cookies_data TEXT DEFAULT '',
            status TEXT DEFAULT 'Pending',
            reward_amount REAL DEFAULT 0.0,
            reject_reason TEXT DEFAULT '',
            is_archived INTEGER DEFAULT 0,
            created_at TIMESTAMP,
            updated_at TIMESTAMP
        )
        """)
        
        try:
            cursor.execute("ALTER TABLE task_submissions ADD COLUMN is_archived INTEGER DEFAULT 0")
        except sqlite3.OperationalError:
            pass
        try:
            cursor.execute("ALTER TABLE task_submissions ADD COLUMN category_label TEXT DEFAULT ''")
        except sqlite3.OperationalError:
            pass

        cursor.execute("""
        CREATE TABLE IF NOT EXISTS old_gmail_categories (
            cat_key TEXT PRIMARY KEY,
            label TEXT NOT NULL,
            rate REAL NOT NULL,
            created_at TEXT
        )
        """)

        cursor.execute("CREATE INDEX IF NOT EXISTS idx_submissions_user_id ON task_submissions(user_id)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_submissions_status ON task_submissions(status)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_submissions_username ON task_submissions(submitted_username)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_submissions_archived ON task_submissions(is_archived)")

        cursor.execute("""
        CREATE TABLE IF NOT EXISTS admin_settings (
            key TEXT PRIMARY KEY,
            value TEXT
        )
        """)

        cursor.execute("""
        CREATE TABLE IF NOT EXISTS withdrawals (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            method TEXT,
            number TEXT,
            amount REAL,
            status TEXT DEFAULT 'Pending',
            created_at TIMESTAMP
        )
        """)
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_withdrawals_user_id ON withdrawals(user_id)")
        # Backward-compatible fields for USDT accounting.
        for col_name, col_type in [
            ("usdt_amount", "REAL"),
            ("usdt_fee", "REAL"),
            ("usdt_rate", "REAL"),
            ("currency", "TEXT"),
        ]:
            try:
                cursor.execute(f"ALTER TABLE withdrawals ADD COLUMN {col_name} {col_type}")
            except sqlite3.OperationalError:
                pass

        cursor.execute("""
        CREATE TABLE IF NOT EXISTS sub_admins (
            user_id INTEGER PRIMARY KEY,
            added_by INTEGER,
            created_at TIMESTAMP
        )
        """)

        default_settings = {
            "gmail_task_price": "22.00",
            "gmail_default_password": "aass1122",
            "wd_bkash_active": "ON",
            "wd_nagad_active": "ON",
            "wd_usdt_active": "ON",
            "min_withdraw_bkash": "50.0",
            "bkash_charge_bdt": "0.0",
            "nagad_charge_bdt": "0.0",
            "max_withdraw_bkash": "5000.0",
            "min_withdraw_nagad": "50.0",
            "max_withdraw_nagad": "5000.0",
            "min_withdraw_usdt": "0.25",
            "max_withdraw_usdt": "100.0",
            "usdt_fee": "0.05",
            "usdt_rate_bdt": "120.0",
            "usdt_charge_enabled": "ON",
            "usdt_charge_usdt": "0.05",
            "admin_id": ADMIN_ID,
            "bot_display_name": "Gmail Work BD",
            "welcome_banner_file_id": "",
            "support_handle": "https://t.me/TrustVaultMailsOwners",
            "official_channel_link": "https://t.me/TrustVaultMailsOfficial",
            "ref_bonus_percent": "10.0",
            "force_join_enabled": "OFF",
            "force_channels_list": "@TrustVaultMailsOfficial|https://t.me/TrustVaultMailsOfficial",
            "gmail_task_active": "ON",
            "bulk_gmail_task_active": "ON",
            "old_gmail_task_active": "ON",
            "old_gmail_2fa_active": "ON",
            "old_gmail_default_password": "oldpass123",
            "old_gmail_recovery_email": "",
            "maintenance_mode": "OFF",
        }

        for key, val in default_settings.items():
            cursor.execute("INSERT OR IGNORE INTO admin_settings (key, value) VALUES (?, ?)", (key, str(val)))

        # One-time migration for the old mismatched Force Join default.
        cursor.execute("SELECT value FROM admin_settings WHERE key = ?", ("force_channels_list",))
        fj_row = cursor.fetchone()
        if fj_row and fj_row["value"] == "@TrustVaultMails_Owners|https://t.me/TrustVaultMailsOfficial":
            cursor.execute(
                "UPDATE admin_settings SET value = ? WHERE key = ?",
                ("@TrustVaultMailsOfficial|https://t.me/TrustVaultMailsOfficial", "force_channels_list")
            )

        if ADMIN_ID:
            cursor.execute(
                "UPDATE admin_settings SET value = ? WHERE key = ?",
                (ADMIN_ID, "admin_id")
            )
            primary_id = next((int(x.strip()) for x in ADMIN_ID.split(",") if x.strip().isdigit()), None)
            if primary_id is not None:
                cursor.execute("DELETE FROM sub_admins WHERE user_id != ?", (primary_id,))
        cursor.execute("SELECT value FROM admin_settings WHERE key = 'old_cats_seeded'")
        if not cursor.fetchone():
            cursor.execute("SELECT COUNT(*) FROM old_gmail_categories")
            if cursor.fetchone()[0] == 0:
                now_ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                cursor.execute(
                    "INSERT OR IGNORE INTO old_gmail_categories (cat_key, label, rate, created_at) VALUES (?, ?, ?, ?)",
                    ("old1", "2018 - 2024 Old", 35.0, now_ts)
                )
            cursor.execute("INSERT OR REPLACE INTO admin_settings (key, value) VALUES ('old_cats_seeded', '1')")
        conn.commit()
    logger.info("Database initialized successfully.")

# -----------------------------------------------------------------------------
# CENTRALIZED MULTILINGUAL TRANSLATION SYSTEM
# -----------------------------------------------------------------------------
MESSAGES = {
    "bn": {
        "btn_task": "Gmail Work",
        "btn_wallet": "Wallet",
        "btn_ref": "My Referrals",
        "btn_leaderboard": "Leaderboard",
        "btn_withdraw": "Withdraw",
        "btn_support": "Support",
        "btn_lang": "Language",
        "btn_admin_panel": "Admin Panel",
        "btn_main_menu": "Main Menu",
        "btn_cancel": "Cancel",
        "btn_gmail": "New Gmail Work",
        "btn_bulk_gmail": "Bulk Gmail Work",
        "btn_new_single": "Single New Gmail",
        "btn_new_bulk": "Bulk New Gmail",
        "btn_old_gmail": "Old Gmail Work",
        "btn_old_single": "Single Old Gmail",
        "btn_old_bulk": "Bulk Old Gmail",
        "btn_back_gmail_work": "Back to Gmail Work",
        "btn_old_cats": "Old Gmail Categories",
        "btn_gen_2fa": "🔑 2FA Code Generate করুন",
        "btn_acct_done": "✅ একাউন্ট খোলা শেষ",
        "btn_bkash": "Bkash",
        "btn_nagad": "Nagad",
        "btn_usdt": "USDT (BEP-20)",
        "btn_binance_uid": "Binance UID",
        "btn_binance_uid": "Binance UID",
        "btn_pending_tasks": "Pending Tasks",
        "btn_held_tasks": "Held Tasks",
        "btn_task_search": "Task Search",
        "btn_user_export": "User Export",
        "btn_pwd_mgr": "Password Manager",
        "btn_price_mgr": "Price Manager",
        "btn_wd_reqs": "Withdraw Requests",
        "btn_sys_settings": "System Settings",
        "btn_live_stats": "Live Stats",
        "btn_broadcast": "Broadcast",
        "btn_backup_db": "Backup DB",
        "btn_restore_db": "Restore DB",

        "btn_bulk_approve": "Approve All",
        "btn_bulk_reject": "Reject All",
        "btn_excel_export": "Spreadsheet UI",
        "btn_user_mgmt": "User Management",
        "btn_admin_mgmt": "Admin Control",

        "maintenance_msg": "⚙️ **বটটি বর্তমানে মেইনটেনেন্স মোডে রয়েছে। দয়া করে কিছুক্ষণ পর চেষ্টা করুন।**",
        "banned_msg": "❌ আপনার অ্যাকাউন্টটি সিস্টেমে নিষিদ্ধ করা হয়েছে।",
        "force_join_msg": "📢 **বটটি ব্যবহার করতে আপনাকে আমাদের অফিশিয়াল চ্যানেলে যুক্ত থাকতে হবে:**\n\n📌 অনুগ্রহ করে নিচের সবকটি চ্যানেলে জয়েন করে **✅ ভেরিফাই করুন** বাটনে চাপ দিন:",
        "btn_join_channel": "Join Channel",
        "btn_verify": "Verify",
        "force_join_success": "✅ **ধন্যবাদ! আপনার চ্যানেল ভেরিফিকেশন সফল হয়েছে।**",
        "force_join_fail": "❌ আপনি এখনো সবগুলো চ্যানেলে জয়েন করেননি! সব চ্যানেলে জয়েন করে পুনরায় চেষ্টা করুন।",
        "welcome_msg": "💎 **স্বাগতম, {name}!** ✨\n━━━━━━━━━━━━━━━━━━━━━━\n**{bot_name}**-এ আপনাকে স্বাগতম।\nএখানে আপনি সহজ **Gmail** অ্যাকাউন্ট সাবমিট করে দ্রুত ও নিশ্চিত ইনকাম করতে পারবেন।\n━━━━━━━━━━━━━━━━━━━━━━\n👉 **কাজ শুরু করতে নিচের Gmail Work বাটনে চাপ দিন।**",
        "ref_new_user_joined": "🎉 **নতুন রেফারেল যোগ হয়েছে!**\n━━━━━━━━━━━━━━━━━━━━━━\n👤 **ইউজার:** {name}\n🆔 **ইউজার আইডি:** <code>{user_id}</code>\n\nকর্মী কাজ সম্পন্ন করলে আপনি পাবেন ১০% সরাসরি কমিশন!",
        "wd_all_disabled": "❌ বর্তমানে সকল পেমেন্ট মেথড বন্ধ রয়েছে। অনুগ্রহ করে পরে চেষ্টা করুন।",
        "task_cat_prompt": "📂 **Gmail Work Menu**\n━━━━━━━━━━━━━━━━━━━━━━\nনিচের যেকোনো একটি কাজের ধরন বেছে নিন:",
        "task_off_gmail": "❌ জিমেইলের কাজ বর্তমানে বন্ধ রয়েছে।",
        "gmail_guidelines": "✨ <b>New Gmail Work</b>\n━━━━━━━━━━━━━━━━━━━━━━\n💵 <b>রেট (প্রতি জিমেইল):</b> ৳{price}\n🔑 <b>পাসওয়ার্ড (ট্যাপ করে কপি করুন):</b> <code>{fixed_pass}</code>\n━━━━━━━━━━━━━━━━━━━━━━\n<b>কাজের নিয়মাবলি:</b>\n১. উপরের পাসওয়ার্ড দিয়ে একটি নতুন Gmail তৈরি করুন।\n২. তৈরি করা Gmail অ্যাড্রেসটি নিচে লিখে পাঠান।\n৩. এরপর নিচের <b>Submit Gmail</b> বাটনে চাপ দিন।",
        "btn_done": "Submit Gmail",
        "file_received_msg": "📄 **ফাইল/ডেটা সফলভাবে গ্রহণ করা হয়েছে!**\n\nকাজটি রিভিউয়ের জন্য জমা দিতে নিচে **'✅ জমা দিন (Done)'** বাটনে চাপ দিন।",
        "no_file_or_data_err": "⚠️ **কোনো ফাইল বা ডেটা পাওয়া যায়নি!**\n\nপ্রথমে এক্সেল ফাইল (.xlsx) বা টেক্সট পাঠান, তারপর **'✅ জমা দিন (Done)'** বাটনে চাপ দিন।",
        "otp_prompt_secret": "🔑 **আপনার 2FA Secret Key-টি প্রদান করুন:**\n\nঅ্যাপ থেকে প্রাপ্ত সিক্রেট কোডটি নিচে মেসেজ করুন:",
        "otp_invalid_secret": "❌ **ভুল সিক্রেট কোড!**\n\nদয়া করে সঠিক ৩২ অক্ষরের বেস৩২ সিক্রেট কি পাঠান:",
        "otp_verified_msg": "✅ **ওটিপি সিক্রেট কি ভেরিফাই হয়েছে!**\n━━━━━━━━━━━━━━━━━━━━━━\n🔐 **সিক্রেট কি:** <code>{secret}</code>\n🔢 **লাইভ OTP:** <code>{otp}</code> (ট্যাপ করে কপি করুন)\n━━━━━━━━━━━━━━━━━━━━━━\nএই কোডটি দিয়ে অ্যাকাউন্ট ভেরিফিকেশন সম্পন্ন করুন।\nসম্পন্ন হলে **'✅ একাউন্ট খোলা শেষ'** বাটনে চাপ দিন।",
        "otp_need_verify_first": "⚠️ **প্রথমে 2FA সিক্রেট কি ভেরিফাই করুন!**\n\n'🔑 2FA Code Generate করুন' বাটনে চাপ দিন।",
        "enter_username_prompt": "👤 **আপনার অ্যাকাউন্টের ইউজারনেম/ইমেইল/UID টাইপ করে পাঠান:**",
        "task_submitted_success": "🎉 <b>টাস্ক সফলভাবে জমা দেওয়া হয়েছে!</b> ✨\n━━━━━━━━━━━━━━━━━━━━━━\n🆔 <b>টাস্ক আইডি:</b> #{sub_id}\n📱 <b>টাইপ:</b> {task_display}\n👤 <b>ইউজারনেম:</b> <code>{username}</code>\n━━━━━━━━━━━━━━━━━━━━━━\n🔍 <b>আপনার তথ্য সফলভাবে পাওয়া গেছে!</b>\n\nআগামী ২৪ থেকে ৪৮ ঘণ্টার মধ্যে আপনার একাউন্ট চেক করে ব্যালেন্স যোগ করে দেওয়া হবে।\n\nএডমিন রিভিউ করে দ্রুত আপনার ওয়ালেটে ব্যালেন্স যুক্ত করে দেবে।\n\nধন্যবাদ!",
        "bulk_gmail_guidelines": "📦 <b>Bulk Gmail Work (একসাথে অনেকগুলো জমা)</b>\n━━━━━━━━━━━━━━━━━━━━━━\n💵 <b>রেট (প্রতি জিমেইল):</b> ৳{price}\n🔑 <b>সবগুলোর জন্য একই পাসওয়ার্ড (ট্যাপ করে কপি করুন):</b>\n<code>{fixed_pass}</code>\n━━━━━━━━━━━━━━━━━━━━━━\n<b>কাজের নিয়ম:</b>\n১. উপরের একই পাসওয়ার্ড দিয়ে আপনার সব নতুন Gmail তৈরি করুন।\n২. এরপর সবগুলো Gmail অ্যাড্রেস <b>প্রতি লাইনে ১টি করে</b> নিচে লিখে পাঠান।\n\n<b>উদাহরণ:</b>\n<code>mail1@gmail.com\nmail2@gmail.com\nmail3@gmail.com</code>",
        "bulk_gmail_submitted_success": "✅ <b>Bulk Gmail সফলভাবে জমা হয়েছে!</b>\n━━━━━━━━━━━━━━━━━━━━━━\n📧 <b>মোট জমা:</b> {count} টি জিমেইল\n⏳ <b>স্ট্যাটাস:</b> রিভিউ পেন্ডিং (২৪–৪৮ ঘণ্টা)\n━━━━━━━━━━━━━━━━━━━━━━\nঅ্যাডমিন চেক করে এপ্রুভ করলেই ব্যালেন্স ওয়ালেটে যোগ হবে।",
        "gmail_submitted_success": "✅ <b>Gmail সফলভাবে জমা হয়েছে!</b>\n━━━━━━━━━━━━━━━━━━━━━━\n🆔 <b>টাস্ক আইডি:</b> #{sub_id}\n⏳ <b>স্ট্যাটাস:</b> রিভিউ পেন্ডিং (২৪–৪৮ ঘণ্টা)\n━━━━━━━━━━━━━━━━━━━━━━\nচেক সম্পন্ন হলেই আপনার ওয়ালেটে ব্যালেন্স যোগ করে দেওয়া হবে।",
        "operation_cancelled": "❌ কাজটি বাতিল করা হয়েছে।",
        "returning_main_menu": "🔙 মূল মেনুতে ফিরে যাচ্ছেন...",

        "wallet_text": "👛 **আপনার অ্যাকাউন্ট ওয়ালেট** 💎\n━━━━━━━━━━━━━━━━━━━━━━\n💰 **বর্তমান ব্যালেন্স:** ৳{balance:.2f}\n💵 **USDT সমমূল্য:** ${usdt_balance:.4f} (রেট: ৳{usdt_rate:.2f}/USDT)\n━━━━━━━━━━━━━━━━━━━━━━\n💸 **পেন্ডিং উইথড্র:** ৳{pending_wd:.2f}\n💰 **মোট আয় (Lifetime):** ৳{lifetime:.2f}\n👥 **রেফারেল আয়:** ৳{ref_earnings:.2f}\n━━━━━━━━━━━━━━━━━━━━━━\n✅ **সম্পন্ন কাজ:** {completed} টি\n⏳ **রিভিউতে আছে:** {pending} টি\n⌛ **হোল্ডে আছে:** {held} টি\n❌ **রিজেক্টেড:** {rejected} টি",
        "wd_gateway_prompt": "💸 **উইথড্রয়াল পেমেন্ট গেটওয়ে নির্বাচন করুন:**\n━━━━━━━━━━━━━━━━━━━━━━\nপছন্দের মেথড নির্বাচন করুন:",
        "wd_invalid_method": "❌ সঠিক উইথড্রয়াল মেথড নির্বাচন করুন।",
        "wd_enter_account_usdt": "📱 **আপনার USDT (BEP-20) ওয়ালেট অ্যাড্রেস লিখুন:**",
        "wd_enter_account_binance_uid": "🆔 **আপনার Binance UID লিখুন:**",
        "wd_enter_account_mfs": "📱 **আপনার {method} অ্যাকাউন্ট নম্বর লিখুন (১১-১৪ ডিজিট):**",
        "wd_invalid_account": "❌ ভুল অ্যাকাউন্ট নম্বর! সঠিক ১১-১৪ ডিজিটের নম্বর দিন।",
        "wd_enter_amount": "💵 **উইথড্র করার পরিমাণ লিখুন:**\n━━━━━━━━━━━━━━━━━━━━━━\n💰 আপনার বর্তমান ব্যালেন্স: ৳{balance:.2f}\n🔻 সর্বনিম্ন সীমা: ৳{min_limit}\n🔺 সর্বোচ্চ সীমা: ৳{max_limit}",
        "wd_invalid_amount_number": "❌ সঠিক সংখ্যা লিখুন (যেমন: ১০০, ২০০.৫):",
        "wd_min_limit_err": "❌ সর্বনিম্ন উইথড্রসীমা ৳{min_limit}! পুনরায় চেষ্টা করুন।",
        "wd_max_limit_err": "❌ সর্বোচ্চ উইথড্রসীমা ৳{max_limit}! পুনরায় চেষ্টা করুন।",
        "wd_insufficient_bal": "❌ আপনার অ্যাকাউন্টে পর্যাপ্ত ব্যালেন্স নেই! পুনরায় চেষ্টা করুন।",
        "wd_usdt_fee_err": "❌ ফি বাদ দেওয়ার পরে পরিমাণ ০ এর কম! বড় পরিমাণ লিখুন।",
        "wd_success_msg": "✅ **পেমেন্ট রিকোয়েস্ট সফলভাবে সাবমিট হয়েছে!** ✨\n━━━━━━━━━━━━━━━━━━━━━━\n🆔 **উইথড্র আইডি:** #{wd_id}\n💵 **পরিমাণ:** ৳{amount:.2f}{fee_text}\n📱 **মেথড:** {method} ({number})\n━━━━━━━━━━━━━━━━━━━━━━\n💰 আপনার অ্যাকাউন্ট থেকে ৳{amount:.2f} কেটে নেওয়া হয়েছে।\n⏳ স্ট্যাটাস: পেন্ডিং (১২-২৪ ঘণ্টার মধ্যে রিভিউ সম্পন্ন হবে)",

        "ref_text": "👥 **আপনার ব্যক্তিগত রেফারেল প্রোগ্রাম** 💎\n━━━━━━━━━━━━━━━━━━━━━━\nআপনার রেফারেল লিঙ্ক ব্যবহার করে বন্ধুদের আমন্ত্রণ জানান। আপনার আমন্ত্রিত কর্মী যতগুলো কাজ সম্পন্ন ও এপ্রুভ করাবে, প্রতিটি কাজের ওপর আপনি পাবেন সরাসরি **{ref_percent}% কমিশন** আপনার ওয়ালেটে।\n\n🔗 **আপনার রেফারেল লিঙ্ক:**\n<code>{ref_link}</code>\n━━━━━━━━━━━━━━━━━━━━━━\n📊 **রেফারেল ট্র্যাকার:**\n• মোট আমন্ত্রিত কর্মী: <code>{ref_count}</code> জন\n• রেফারেল বোনাস থেকে আয়: <code>৳{ref_earnings:.2f}</code>",
        "leaderboard_title": "🏆 **শীর্ষ ১০ উপার্জনকারী কর্মী (Leaderboard):** 👑\n━━━━━━━━━━━━━━━━━━━━━━\n",
        "no_leaderboard_data": "কোনো ডেটা এখনো পাওয়া যায়নি।",
        "support_text": "🎧 **গ্রাহক সেবা কেন্দ্র (Help Zone)**\n━━━━━━━━━━━━━━━━━━━━━━\n👤 সম্মানিত মেম্বার, আপনার যেকোনো প্রশ্নের জন্য আমাদের সাপোর্ট টিমের সাথে যোগাযোগ করুন।\n📌 অযথা মেসেজ দেওয়া থেকে বিরত থাকুন।",
        "btn_admin_support": "Admin Support",
        "btn_official_channel": "Official Channel",
        "lang_select_prompt": "🌐 **ভাষা নির্বাচন করুন / Select Language:**\n\nনিচের বাটন থেকে আপনার পছন্দের ভাষা নির্বাচন করুন:",
        "lang_set_bn": "✅ বটের ভাষা সফলভাবে **বাংলা** করা হয়েছে।",
        "lang_set_en": "✅ Bot language set to **English**.",

        "user_task_approved": "✅ <b>অভিনন্দন!</b> আপনার [{task_display}] কাজ (<code>{username}</code>) এপ্রুভ হয়েছে। ৳{reward:.2f} ব্যালেন্সে যোগ হয়েছে।",
        "user_task_rejected": "❌ <b>আপনার [{task_display}] কাজ ({username}) রিজেক্ট করা হয়েছে।</b>\n📋 <b>কারণ:</b> {reason}",
        "user_task_held": "⏳ <b>আপনার টাস্ক #{sub_id} হোল্ডে রাখা হয়েছে।</b>\n\n<b>কারণ:</b>\nআপনার কাজটি ম্যানুয়াল রিভিউর অধীনে রয়েছে।",
        "user_ref_commission": "🎉 <b>রেফারেল কমিশন যুক্ত হয়েছে!</b>\n\nআপনার রেফারকৃত কর্মী একটি কাজ সম্পন্ন করায় আপনি <b>৳{bonus:.2f}</b> কমিশন পেয়েছেন।",
        "user_wd_approved": "✅ <b>আপনার উইথড্র রিকোয়েস্ট #{wd_id} সফলভাবে সম্পন্ন হয়েছে!</b>\n💵 <b>৳{amount:.2f}</b> আপনার <code>{number}</code> নম্বরে পাঠানো হয়েছে।",
        "user_wd_rejected": "❌ <b>আপনার উইথড্র রিকোয়েস্ট #{wd_id} রিজেক্ট করা হয়েছে।</b>\n💵 <b>৳{amount:.2f}</b> আপনার ওয়ালেটে রিফান্ড (ফেরত) করা হয়েছে।",

        "admin_dash_title": "👑 **{bot_name}: অ্যাডমিন প্যানেল** 👑\n━━━━━━━━━━━━━━━━━━━━━━\n👥 **মোট ইউজার:** {total_users} জন\n⏳ **পেন্ডিং কাজ:** {pending_tasks} টি\n⌛ **হোল্ড কাজ:** {held_tasks} টি\n💸 **পেন্ডিং উইথড্র:** {pending_wd} টি\n━━━━━━━━━━━━━━━━━━━━━━\nযেকোনো বিষয় কনফিগার বা ভেরিফাই করতে নিচের মেনুগুলো ব্যবহার করুন।",
        "admin_no_pending_tasks": "✅ কোনো পেন্ডিং টাস্ক নেই।",
        "admin_pending_tasks_title": "📩 **পেন্ডিং টাস্ক লিস্ট (সর্বশেষ ১০টি):**\n\n",
        "admin_no_held_tasks": "⏳ কোনো হোল্ড টাস্ক নেই।",
        "admin_held_tasks_title": "⌛ <b>হোল্ড টাস্ক লিস্ট:</b>\n\n",
        "admin_search_prompt_msg": "🔍 **টাস্ক সার্চ:**\nইউজারনেম, জিমেইল, অথবা টাস্ক আইডি লিখুন:",
        "admin_export_prompt_msg": "📂 **ইউজার এক্সপোর্ট:**\nযে ইউজারের ডেটা চান তার টেলিগ্রাম ইউজার আইডি লিখুন:",
        "admin_pwd_mgr_title": "🔑 **পাসওয়ার্ড ম্যানেজ**\n━━━━━━━━━━━━━━━━━━━━━━\n✉️ New/Bulk Gmail Pass: <code>{gmail_pass}</code>\n♻️ Old Gmail Pass: <code>{old_pass}</code>\n💌 Old Recovery: <code>{old_rec}</code>\n━━━━━━━━━━━━━━━━━━━━━━\nনতুন পাসওয়ার্ড সেট করতে নিচের বাটন ব্যবহার করুন।",
        "admin_price_mgr_title": "🏷️ **টাস্ক প্রাইস ম্যানেজার (Task Price Management)**\n━━━━━━━━━━━━━━━━━━━━━━\n✉️ Gmail: <code>৳{gmail}</code>\n👥 Ref Bonus: <code>{ref_bonus}%</code>\n━━━━━━━━━━━━━━━━━━━━━━\nনতুন দাম নির্ধারণ করতে নিচের বাটনে চাপ দিন:",
        "admin_no_pending_wds": "✅ কোনো পেন্ডিং উইথড্র রিকোয়েস্ট নেই।",
        "admin_pending_wds_title": "💸 **পেন্ডিং উইথড্র রিকোয়েস্ট:**\n\n",
        "admin_sys_settings_title": "⚙️ **সিস্টেম সেটিংস (ON/OFF Toggles & Controls)**\n━━━━━━━━━━━━━━━━━━━━━━\n🏷️ Bot Name: {bot_name}\n🎧 Support: {support_link}\n📢 Channel: {channel_link}\n🔒 Force Join Status: {force_join}\n🛠️ Maintenance Mode: {maint}\n✉️ New Gmail: {gmail} | Bulk Gmail: {bulk_gmail}\n♻️ Old Gmail: {old_gmail} | Old 2FA: {old_2fa}\n🟢 Bkash WD: {bkash_wd} (Min: {min_bkash} | Max: {max_bkash} | Charge: ৳{bkash_charge})\n🟠 Nagad WD: {nagad_wd} (Min: {min_nagad} | Max: {max_nagad} | Charge: ৳{nagad_charge})\n🔵 USDT WD: {usdt_wd} (Min: {min_usdt} | Max: {max_usdt})\n💱 USDT Rate: ৳{usdt_rate}/USDT\n💳 USDT Charge: {charge_status} ({charge_usdt} USDT)\n━━━━━━━━━━━━━━━━━━━━━━\nটগল বা পরিবর্তন করতে নিচের বাটন ব্যবহার করুন।",
        "admin_live_stats_title": "📊 **লাইভ স্ট্যাটাস** ⚡\n━━━━━━━━━━━━━━━━━━━━━━\n👥 মোট ইউজার: {total_users}\n⏳ পেন্ডিং টাস্ক: {pending_tasks}\n⌛ হোল্ড টাস্ক: {held_tasks}\n✅ এপ্রুভড টাস্ক: {approved_tasks}\n❌ রিজেক্টেড টাস্ক: {rejected_tasks}\n💸 পেন্ডিং উইথড্র: {pending_wd}\n💰 মোট ইউজার ব্যালেন্স: ৳{total_balance:.2f}",
        "admin_broadcast_prompt_msg": "𝚁𝙱𝙺 𝙰𝙿𝙿 𝚂𝚃𝙾𝚁𝙴: 📢 Broadcast Message পাঠান (All Media Supported)\n━━━━━━━━━━━━━━━━━━━━━━\nএখানে পছন্দমতো যেকোনো **Text, Photo 🖼️, Video 🎥, Document 📄, Voice 🎤, Audio 🎵** সেন্ড করুন। ক্যাপশন থাকলে ক্যাপশনসহ ইউজারদের কাছে অরিজিনাল মেসেজ হিসেবে যাবে।",
        "admin_search_no_results": "❌ কোনো টাস্ক পাওয়া যায়নি।",
        "admin_search_results_title": "🔍 **সার্চ রেজাল্ট:**\n\n",
        "admin_export_invalid_id": "❌ সঠিক ইউজার আইডি দিন।",
        "admin_export_no_tasks": "❌ এই ইউজারের কোনো টাস্ক নেই।",
        "admin_broadcast_done": "✅ **Broadcast Completed** ✨\n━━━━━━━━━━━━━━━━━━━━━━\n👥 **Total Users:** {total}\n📨 **Message Type:** {msg_type}\n✅ **Successfully Sent:** {count}\n❌ **Failed:** {failed}\n📈 **Success Rate:** {success_rate}%\n⏱️ **Total Time:** {time_taken}s\n━━━━━━━━━━━━━━━━━━━━━━\n📊 **Failure Breakdown:**\n• 🚫 Bot Blocked: {blocked}\n• 🗑️ Account Deleted / Chat Not Found: {deleted}\n• ⚠️ Other Errors: {other}",
        "admin_rej_custom_prompt": "📝 **রিজেক্ট করার কাস্টম কারণটি লিখে পাঠান:**",
        "btn_admin_menu": "Admin Menu",
        "admin_bulk_app_prompt": "✅ **বাল্ক এপ্রুভ (Bulk Approve):**\n━━━━━━━━━━━━━━━━━━━━━━\nএখানে একসাথে ৫০-২০০টি Username / UID / Gmail পেস্ট করে সেন্ড করুন (প্রতি লাইনে একটি করে অথবা কমা/স্পেস দিয়ে):\n\nবট অটোম্যাটিক্যালি এগুলো চেক করে **শুধুমাত্র ইনপুট দেওয়া টাস্কগুলো এপ্রুভ** করবে ও ইউজারকে ওয়ালেট ব্যালেন্সসহ নোটিফিকেশন পাঠাবে।",
        "admin_bulk_rej_prompt": "❌ **বাল্ক রিজেক্ট (Bulk Reject):**\n━━━━━━━━━━━━━━━━━━━━━━\nএখানে একসাথে ৫০-২০০টি Username / UID / Gmail পেস্ট করে সেন্ড করুন (প্রতি লাইনে একটি করে অথবা কমা/স্পেস দিয়ে):\n\nবট অটোম্যাটিক্যালি এগুলো চেক করে **শুধুমাত্র ইনপুট দেওয়া টাস্কগুলো রিজেক্ট** করবে ও সংশ্লিষ্ট ইউজারকে নোটিফিকেশন পাঠাবে।",
    },
    "en": {
        "btn_task": "Gmail Work",
        "btn_wallet": "Wallet",
        "btn_ref": "My Referrals",
        "btn_leaderboard": "Leaderboard",
        "btn_withdraw": "Withdraw",
        "btn_support": "Support",
        "btn_lang": "Language",
        "btn_admin_panel": "Admin Panel",
        "btn_main_menu": "Main Menu",
        "btn_cancel": "Cancel",
        "btn_gmail": "New Gmail Work",
        "btn_bulk_gmail": "Bulk Gmail Work",
        "btn_new_single": "Single New Gmail",
        "btn_new_bulk": "Bulk New Gmail",
        "btn_old_gmail": "Old Gmail Work",
        "btn_old_single": "Single Old Gmail",
        "btn_old_bulk": "Bulk Old Gmail",
        "btn_back_gmail_work": "Back to Gmail Work",
        "btn_old_cats": "Old Gmail Categories",
        "btn_gen_2fa": "🔑 Generate 2FA Code",
        "btn_acct_done": "✅ Account Creation Done",
        "btn_bkash": "Bkash",
        "btn_nagad": "Nagad",
        "btn_usdt": "USDT (BEP-20)",
        "btn_pending_tasks": "Pending Tasks",
        "btn_held_tasks": "Held Tasks",
        "btn_task_search": "Task Search",
        "btn_user_export": "User Export",
        "btn_pwd_mgr": "Password Manager",
        "btn_price_mgr": "Price Manager",
        "btn_wd_reqs": "Withdraw Requests",
        "btn_sys_settings": "System Settings",
        "btn_live_stats": "Live Stats",
        "btn_broadcast": "Broadcast",
        "btn_backup_db": "Backup DB",
        "btn_restore_db": "Restore DB",

        "btn_bulk_approve": "Approve All",
        "btn_bulk_reject": "Reject All",
        "btn_excel_export": "Spreadsheet UI",
        "btn_user_mgmt": "User Management",
        "btn_admin_mgmt": "Admin Control",

        "maintenance_msg": "⚙️ **The bot is currently under maintenance. Please try again later.**",
        "banned_msg": "❌ Your account has been banned from the system.",
        "force_join_msg": "📢 **To use this bot, you must join our official channels:**\n\nPlease join all channels below and press the **✅ Verify** button:",
        "btn_join_channel": "Join Channel",
        "btn_verify": "Verify",
        "force_join_success": "✅ **Thank you! Your channel verification was successful.**",
        "force_join_fail": "❌ You haven't joined all channels yet! Please join all channels and try again.",
        "welcome_msg": "💎 **Welcome, {name}!** ✨\n━━━━━━━━━━━━━━━━━━━━━━\nWelcome to **{bot_name}**.\nEarn reliably by completing simple **Gmail** submission tasks.\n━━━━━━━━━━━━━━━━━━━━━━\n👉 **Press Gmail Work below to get started.**",
        "ref_new_user_joined": "🎉 **New Referral Joined!**\n━━━━━━━━━━━━━━━━━━━━━━\n👤 **User:** {name}\n🆔 **User ID:** <code>{user_id}</code>\n\nYou will earn 10% lifetime commission on their completed tasks!",
        "wd_all_disabled": "❌ All withdrawal gateways are currently disabled. Please try again later.",
        "task_cat_prompt": "📂 **Gmail Work Menu**\n━━━━━━━━━━━━━━━━━━━━━━\nSelect a task option below:",
        "task_off_gmail": "❌ Gmail tasks are currently disabled.",
        "gmail_guidelines": "✨ <b>New Gmail Work</b>\n━━━━━━━━━━━━━━━━━━━━━━\n💵 <b>Rate (Per Gmail):</b> ৳{price}\n🔑 <b>Password (Tap to copy):</b> <code>{fixed_pass}</code>\n━━━━━━━━━━━━━━━━━━━━━━\n<b>Instructions:</b>\n1. Create a new Gmail account using the password above.\n2. Send your created Gmail address below.\n3. Tap <b>Submit Gmail</b> to complete.",
        "btn_done": "Submit Gmail",
        "file_received_msg": "📄 **File/Data received successfully!**\n\nPress **'✅ Submit (Done)'** below to send it to the admin.",
        "no_file_or_data_err": "⚠️ **No file or data found!**\n\nPlease send an Excel file (.xlsx) or text first, then press **'✅ Submit (Done)'**.",
        "otp_prompt_secret": "🔑 **Provide your 2FA Secret Key:**\n\nSend the secret code obtained from the app here:",
        "otp_invalid_secret": "❌ **Invalid Secret Code!**\n\nPlease send a valid Base32 Secret Key:",
        "otp_verified_msg": "✅ **OTP Secret Key verified!**\n\n🔐 **Secret Key:** <code>{secret}</code>\n🔢 **Live OTP:** <code>{otp}</code> (Tap to copy)\n\nComplete account verification using this code.\nWhen finished, press the '✅ Account Creation Done' button.",
        "otp_need_verify_first": "⚠️ **Please verify your 2FA Secret Key first!**\n\nPress the '🔑 Generate 2FA Code' button.",
        "enter_username_prompt": "👤 **Type and send your account Username/Email/UID:**",
        "task_submitted_success": "🎉 <b>Task submitted successfully!</b> ✨\n━━━━━━━━━━━━━━━━━━━━━━\n🆔 <b>Task ID:</b> #{sub_id}\n📱 <b>Type:</b> {task_display}\n👤 <b>Username:</b> <code>{username}</code>\n━━━━━━━━━━━━━━━━━━━━━━\n🔍 <b>Your information has been received successfully!</b>\n\nYour account will be checked and the balance will be added within 24–48 hours.\n\nThe admin will review your submission and credit your wallet as soon as possible.\n\nThank you!",
        "bulk_gmail_guidelines": "📦 <b>Bulk Gmail Work (Submit Multiple)</b>\n━━━━━━━━━━━━━━━━━━━━━━\n💵 <b>Rate (Per Gmail):</b> ৳{price}\n🔑 <b>Same Password For All (Tap to copy):</b>\n<code>{fixed_pass}</code>\n━━━━━━━━━━━━━━━━━━━━━━\n<b>Instructions:</b>\n1. Create all your Gmail accounts using the SAME password above.\n2. Send all Gmail addresses below, <b>one per line</b>.\n\n<b>Example:</b>\n<code>mail1@gmail.com\nmail2@gmail.com\nmail3@gmail.com</code>",
        "bulk_gmail_submitted_success": "✅ <b>Bulk Gmail Submitted Successfully!</b>\n━━━━━━━━━━━━━━━━━━━━━━\n📧 <b>Total Submitted:</b> {count} Gmail(s)\n⏳ <b>Status:</b> Pending Review (24–48 hrs)\n━━━━━━━━━━━━━━━━━━━━━━\nYour wallet will be credited once reviewed and approved.",
        "gmail_submitted_success": "✅ <b>Gmail Submitted Successfully!</b>\n━━━━━━━━━━━━━━━━━━━━━━\n🆔 <b>Task ID:</b> #{sub_id}\n⏳ <b>Status:</b> Pending Review (24–48 hrs)\n━━━━━━━━━━━━━━━━━━━━━━\nYour wallet will be credited as soon as the review is complete.",
        "operation_cancelled": "❌ Operation cancelled.",
        "returning_main_menu": "🔙 Returning to Main Menu...",

        "wallet_text": "💵 **Your Balance**\n━━━━━━━━━━━━━━━━━━━━━━\n💰 **Current Balance:** ৳{balance:.2f}\n💵 **USDT Equivalent:** ${usdt_balance:.4f} (Rate: ৳{usdt_rate:.2f}/USDT)\n━━━━━━━━━━━━━━━━━━━━━━\n💸 **Pending Withdrawal:** ৳{pending_wd:.2f}\n💰 **Lifetime Earnings:** ৳{lifetime:.2f}\n👥 **Referral Earnings:** ৳{ref_earnings:.2f}\n━━━━━━━━━━━━━━━━━━━━━━\n✅ **Completed Tasks:** {completed}\n⏳ **Under Review:** {pending}\n⌛ **On Hold:** {held}\n❌ **Rejected:** {rejected}",
        "wd_gateway_prompt": "💸 **Select Withdrawal Payment Gateway:**\n\nChoose payment method:",
        "wd_invalid_method": "❌ Please select a valid withdrawal method.",
        "wd_enter_account_usdt": "📱 **Enter your USDT (BEP-20) wallet address:**",
        "wd_enter_account_binance_uid": "🆔 **Enter your Binance UID:**",
        "wd_enter_account_mfs": "📱 **Enter your {method} account number (11-14 digits):**",
        "wd_invalid_account": "❌ Invalid account number! Enter a valid 11-14 digit number.",
        "wd_enter_amount": "💵 **Enter withdrawal amount:**\n\n💰 Current Balance: ৳{balance:.2f}\n🔻 Minimum Limit: ৳{min_limit}\n🔺 Maximum Limit: ৳{max_limit}",
        "wd_invalid_amount_number": "❌ Enter a valid number (e.g. 100, 200.5):",
        "wd_min_limit_err": "❌ Minimum withdrawal limit is ৳{min_limit}! Try again.",
        "wd_max_limit_err": "❌ Maximum withdrawal limit is ৳{max_limit}! Try again.",
        "wd_insufficient_bal": "❌ Insufficient balance in your account! Try again.",
        "wd_usdt_fee_err": "❌ Amount after fee deduction is less than 0! Enter a larger amount.",
        "wd_success_msg": "✅ **Payment request submitted successfully!**\n\n🆔 **Withdrawal ID:** #{wd_id}\n💵 **Amount:** ৳{amount:.2f}{fee_text}\n📱 **Method:** {method} ({number})\n💰 ৳{amount:.2f} has been deducted from your current balance.\n⏳ Status: Pending (Review will be completed within 12-24 hours)",

        "ref_text": "👥 **My Referral Program**\n\nInvite your friends using your referral link. You will earn a direct **{ref_percent}% commission** on EVERY task your referred worker completes and gets approved.\n\n🔗 **Your Personal Referral Link:**\n<code>{ref_link}</code>\n\n📊 **Your Referral Tracker:**\n• Total Invited Workers: <code>{ref_count}</code>\n• Total Referral Earnings: <code>৳{ref_earnings:.2f}</code>",
        "leaderboard_title": "🏆 **Top 10 Earners:**\n\n",
        "no_leaderboard_data": "No data available yet.",
        "support_text": "↗️ **Customer Support Center**\n\n👤 Dear Member, for any issue please contact our support team.\n📌 Avoid sending unnecessary messages.",
        "btn_admin_support": "Admin Support",
        "btn_official_channel": "Official Channel",
        "lang_select_prompt": "🌐 **Select Language:**\n\nChoose your preferred language from the buttons below:",
        "lang_set_bn": "✅ বটের ভাষা সফলভাবে **বাংলা** করা হয়েছে।",
        "lang_set_en": "✅ Bot language set to **English**.",

        "user_task_approved": "✅ <b>Congratulations!</b> Your [{task_display}] task (<code>{username}</code>) has been approved. ৳{reward:.2f} added to your balance.",
        "user_task_rejected": "❌ <b>Your [{task_display}] task ({username}) was rejected.</b>\n📋 <b>Reason:</b> {reason}",
        "user_task_held": "⏳ <b>Your task #{sub_id} has been placed on Hold.</b>\n\n<b>Reason:</b>\nYour task is under manual review.",
        "user_ref_commission": "🎉 <b>Referral Commission Added!</b>\n\nYou received <b>৳{bonus:.2f}</b> commission for a task completed by your referred worker.",
        "user_wd_approved": "✅ <b>Your withdrawal request #{wd_id} has been approved successfully!</b>\n💵 <b>৳{amount:.2f}</b> sent to your <code>{number}</code> account.",
        "user_wd_rejected": "❌ <b>Your withdrawal request #{wd_id} was rejected.</b>\n💵 <b>৳{amount:.2f}</b> refunded back to your wallet.",

        "admin_dash_title": "👑 **{bot_name}: Admin Panel**\n\n👥 **Total Users:** {total_users}\n⏳ **Pending Tasks:** {pending_tasks}\n⌛ **Hold Tasks:** {held_tasks}\n💸 **Pending Withdrawals:** {pending_wd}\n\nUse the menus below to configure or verify anything.",
        "admin_no_pending_tasks": "✅ No pending tasks.",
        "admin_pending_tasks_title": "📩 **Pending Tasks List (Latest 10):**\n\n",
        "admin_no_held_tasks": "⏳ No held tasks.",
        "admin_held_tasks_title": "⌛ <b>Held Tasks List:</b>\n\n",
        "admin_search_prompt_msg": "🔍 **Task Search:**\nType Username, Gmail, or Task ID:",
        "admin_export_prompt_msg": "📂 **User Export:**\nType Telegram User ID of the user you want data for:",
        "admin_pwd_mgr_title": "🔑 **Password Manager**\n\n✉️ New/Bulk Gmail Pass: <code>{gmail_pass}</code>\n♻️ Old Gmail Pass: <code>{old_pass}</code>\n💌 Old Recovery: <code>{old_rec}</code>\n\nUse buttons below to set new passwords.",
        "admin_price_mgr_title": "🏷️ **Task Price Manager**\n\n✉️ Gmail: <code>৳{gmail}</code>\n👥 Ref Bonus: <code>{ref_bonus}%</code>\n\nUse buttons below to set new prices:",
        "admin_no_pending_wds": "✅ No pending withdrawal requests.",
        "admin_pending_wds_title": "💸 **Pending Withdrawal Requests:**\n\n",
        "admin_sys_settings_title": "⚙️ **System Settings (ON/OFF Toggles)**\n\n🏷️ Bot Name: {bot_name}\n🎧 Support: {support_link}\n📢 Channel: {channel_link}\n🔒 Force Join: {force_join}\n🛠️ Maintenance: {maint}\n✉️ New Gmail: {gmail} | Bulk Gmail: {bulk_gmail}\n♻️ Old Gmail: {old_gmail} | Old 2FA: {old_2fa}\n🟢 Bkash WD: {bkash_wd} (Min: {min_bkash} | Max: {max_bkash} | Charge: ৳{bkash_charge})\n🟠 Nagad WD: {nagad_wd} (Min: {min_nagad} | Max: {max_nagad} | Charge: ৳{nagad_charge})\n🔵 USDT WD: {usdt_wd} (Min: {min_usdt} | Max: {max_usdt})\n💱 USDT Rate: ৳{usdt_rate}/USDT\n💳 USDT Charge: {charge_status} ({charge_usdt} USDT)\n\nUse buttons below to toggle or edit settings.",
        "admin_live_stats_title": "📊 **Live Stats**\n\n👥 Total Users: {total_users}\n⏳ Pending Tasks: {pending_tasks}\n⌛ Hold Tasks: {held_tasks}\n✅ Approved Tasks: {approved_tasks}\n❌ Rejected Tasks: {rejected_tasks}\n💸 Pending Withdrawals: {pending_wd}\n💰 Total User Balance: ৳{total_balance:.2f}",
        "admin_broadcast_prompt_msg": "𝚁𝙱𝙺 𝙰𝙿𝙿 𝚂𝚃𝙾𝚁𝙴: 📢 Broadcast Message পাঠান (All Media Supported)\n━━━━━━━━━━━━━━━━━━━━━━\nSend any **Text, Photo 🖼️, Video 🎥, Document 📄, Voice 🎤, Audio 🎵**. Caption will be preserved.",
        "admin_search_no_results": "❌ No tasks found.",
        "admin_search_results_title": "🔍 **Search Results:**\n\n",
        "admin_export_invalid_id": "❌ Enter a valid User ID.",
        "admin_export_no_tasks": "❌ This user has no tasks.",
        "admin_broadcast_done": "✅ **Broadcast Completed**\n\n👥 **Total Users:** {total}\n📨 **Message Type:** {msg_type}\n✅ **Successfully Sent:** {count}\n❌ **Failed:** {failed}\n📈 **Success Rate:** {success_rate}%\n⏱️ **Total Time:** {time_taken}s\n\n📊 **Failure Breakdown:**\n• 🚫 Bot Blocked: {blocked}\n• 🗑️ Account Deleted / Chat Not Found: {deleted}\n• ⚠️ Other Errors: {other}",
        "admin_rej_custom_prompt": "📝 **Type the custom rejection reason:**",
        "btn_admin_menu": "Admin Menu",
        "admin_bulk_app_prompt": "✅ **Bulk Approve:**\n\nPaste 50-200 Usernames / UIDs / Gmails here (one per line or separated by space/comma):\n\nBot will match and **Approve ONLY the matched tasks**.",
        "admin_bulk_rej_prompt": "❌ **Bulk Reject:**\n\nPaste 50-200 Usernames / UIDs / Gmails here (one per line or separated by space/comma):\n\nBot will match and **Reject ONLY the matched tasks**.",
    }
}

def get_user_lang(user_id: int) -> str:
    if not user_id:
        return 'bn'
    try:
        with get_db() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT language FROM users WHERE user_id = ?", (user_id,))
            row = cursor.fetchone()
            if row and row["language"]:
                return row["language"]
    except Exception:
        pass
    return 'bn'

def tr(user_id_or_lang: Union[int, str, None], key: str, **kwargs) -> str:
    if isinstance(user_id_or_lang, int):
        lang = get_user_lang(user_id_or_lang)
    elif isinstance(user_id_or_lang, str):
        lang = user_id_or_lang if user_id_or_lang in ('bn', 'en') else 'bn'
    else:
        lang = 'bn'

    lang_dict = MESSAGES.get(lang, MESSAGES['bn'])
    msg_template = lang_dict.get(key, MESSAGES['bn'].get(key, ''))
    if kwargs and msg_template:
        try:
            return msg_template.format(**kwargs)
        except Exception:
            return msg_template
    return msg_template

# -----------------------------------------------------------------------------
# HELPER FUNCTIONS
# -----------------------------------------------------------------------------
def get_setting_val(key: str, default: str = "") -> str:
    try:
        with get_db() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT value FROM admin_settings WHERE key = ?", (key,))
            row = cursor.fetchone()
            return row["value"] if row else default
    except Exception:
        return default

def set_setting_val(key: str, value: str):
    try:
        with get_db() as conn:
            cursor = conn.cursor()
            cursor.execute("INSERT OR REPLACE INTO admin_settings (key, value) VALUES (?, ?)", (key, str(value)))
            conn.commit()
    except Exception as e:
        logger.error(f"Error setting setting val {key}: {e}")


def get_old_categories():
    try:
        with get_db() as conn:
            return conn.execute("SELECT * FROM old_gmail_categories ORDER BY created_at ASC").fetchall()
    except Exception:
        return []

def get_old_category(cat_key: str):
    try:
        with get_db() as conn:
            return conn.execute("SELECT * FROM old_gmail_categories WHERE cat_key = ?", (cat_key,)).fetchone()
    except Exception:
        return None

def add_old_category(label: str, rate: float) -> str:
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    existing = {r["cat_key"] for r in get_old_categories()}
    i = 1
    while f"old{i}" in existing:
        i += 1
    key = f"old{i}"
    with get_db() as conn:
        conn.execute(
            "INSERT INTO old_gmail_categories (cat_key, label, rate, created_at) VALUES (?, ?, ?, ?)",
            (key, label, rate, now)
        )
        conn.commit()
    return key

def del_old_category(cat_key: str):
    with get_db() as conn:
        conn.execute("DELETE FROM old_gmail_categories WHERE cat_key = ?", (cat_key,))
        conn.commit()

def is_owner(user_id: int) -> bool:
    admin_id_setting = ADMIN_ID or get_setting_val("admin_id", "")
    primary_ids = [int(x.strip()) for x in admin_id_setting.split(",") if x.strip().isdigit()]
    return bool(primary_ids and user_id == primary_ids[0])

def is_admin(user_id: int) -> bool:
    if is_owner(user_id):
        return True

    admin_id_setting = ADMIN_ID or get_setting_val("admin_id", "")
    admins = [int(x.strip()) for x in admin_id_setting.split(",") if x.strip().isdigit()]
    if user_id in admins:
        return True

    try:
        with get_db() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT user_id FROM sub_admins WHERE user_id = ?", (user_id,))
            if cursor.fetchone():
                return True
    except Exception:
        pass
    return False

def is_banned(user_id: int) -> bool:
    try:
        with get_db() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT is_banned FROM users WHERE user_id = ?", (user_id,))
            row = cursor.fetchone()
            return bool(row["is_banned"]) if row else False
    except Exception:
        return False

def get_user_data(user_id: int):
    try:
        with get_db() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM users WHERE user_id = ?", (user_id,))
            return cursor.fetchone()
    except Exception:
        None


def get_referrer_commission(referrer_id: int, reward_amount: float) -> float:
    """Return fixed custom referral bonus if set for referrer, else global % commission."""
    ref_user = get_user_data(referrer_id)
    if ref_user and ref_user["custom_ref_fixed"] is not None:
        return float(ref_user["custom_ref_fixed"])
    commission_percent = float(get_setting_val("ref_bonus_percent", "10.0"))
    return float(reward_amount) * (commission_percent / 100.0)

def register_user(tg_user, referred_by_id=None):
    user_id = tg_user.id
    if get_user_data(user_id):
        return

    ref_code = f"NX{secrets.token_hex(3).upper()}"
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    ref_by = None
    if referred_by_id and referred_by_id != user_id:
        try:
            with get_db() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT user_id FROM users WHERE user_id = ?", (referred_by_id,))
                if cursor.fetchone():
                    ref_by = referred_by_id
                    cursor.execute("UPDATE users SET ref_count = ref_count + 1 WHERE user_id = ?", (ref_by,))
                    conn.commit()
        except Exception as e:
            logger.error(f"Error updating ref count: {e}")

    try:
        with get_db() as conn:
            cursor = conn.cursor()
            cursor.execute(
                """
                INSERT INTO users (user_id, username, full_name, balance, referral_code, referred_by, joined_at, language)
                VALUES (?, ?, ?, 0.0, ?, ?, ?, 'bn')
                """,
                (user_id, tg_user.username, tg_user.full_name, ref_code, ref_by, now)
            )
            conn.commit()
    except Exception as e:
        logger.error(f"Error registering user: {e}")

# MULTI-CHANNEL FORCE JOIN CHECKER
# -----------------------------------------------------------------------------
# Supported admin input formats:
#   @publicchannel
#   @publicchannel|https://t.me/publicchannel
#   --1004397984847|https://t.me/+invite_hash
# IMPORTANT: for private/invite-only channels the left side MUST be the real
# Telegram chat ID (usually -100xxxxxxxxxx) and the bot must be an admin there.
# -----------------------------------------------------------------------------
def parse_force_join_channels(raw_value: str) -> List[Tuple[str, str]]:
    """Parse and normalize the admin-configured force-join channel list."""
    channels: List[Tuple[str, str]] = []
    seen = set()
    raw_value = str(raw_value or "")

    for raw_line in re.split(r"[\r\n]+", raw_value):
        line = raw_line.strip()
        if not line:
            continue

        if "|" in line:
            identifier, join_link = line.split("|", 1)
            identifier = identifier.strip()
            join_link = join_link.strip()
        else:
            identifier = line
            join_link = ""

        # Accept a public t.me URL as the identifier when no explicit ID is given.
        if (identifier.startswith("https://t.me/") or identifier.startswith("http://t.me/")):
            tail = identifier.split("t.me/", 1)[1].strip("/")
            if tail and not tail.startswith("+") and not tail.startswith("joinchat/"):
                identifier = "@" + tail.lstrip("@")

        if not identifier:
            continue

        # Normalize @username but preserve numeric Telegram chat IDs.
        if not identifier.startswith("-") and not identifier.isdigit() and not identifier.startswith("@"):
            identifier = "@" + identifier

        # If no explicit join URL is supplied, generate a public-channel URL.
        if not join_link:
            if identifier.startswith("@"):
                join_link = f"https://t.me/{identifier[1:]}"
            elif identifier.startswith("-") or identifier.isdigit():
                # A numeric private-channel ID cannot be converted into a join URL.
                # Admin must supply the actual invite link after '|'.
                join_link = "https://t.me/"

        # Only allow safe Telegram links in the button.
        if not re.match(r"^https?://t\.me/", join_link, re.IGNORECASE):
            if identifier.startswith("@"):
                join_link = f"https://t.me/{identifier[1:]}"
            else:
                join_link = "https://t.me/"

        key = identifier.lower()
        if key not in seen:
            seen.add(key)
            channels.append((identifier, join_link))

    return channels


def _member_is_joined(member) -> bool:
    """Return True for every Telegram membership state that counts as joined."""
    status = getattr(member, "status", "")
    if status in {"creator", "administrator", "member"}:
        return True
    # Telegram can return 'restricted' while the user is still a member.
    if status == "restricted" and bool(getattr(member, "is_member", False)):
        return True
    return False


async def check_force_join(user_id: int, context: ContextTypes.DEFAULT_TYPE) -> Tuple[bool, List[Tuple[str, str]]]:
    if is_admin(user_id):
        return True, []

    enabled = str(get_setting_val("force_join_enabled", "OFF")).strip().upper()
    if enabled != "ON":
        return True, []

    raw_channels = get_setting_val(
        "force_channels_list",
        "@TrustVaultMailsOfficial|https://t.me/TrustVaultMailsOfficial"
    )
    channels = parse_force_join_channels(raw_channels)

    if not channels:
        # Never lock every user out because the admin accidentally saved an empty list.
        logger.warning("Force Join is ON but force_channels_list is empty; allowing access.")
        return True, []

    unjoined: List[Tuple[str, str]] = []
    for channel_identifier, join_link in channels:
        try:
            chat_identifier = (
                int(channel_identifier)
                if channel_identifier.lstrip("-").isdigit()
                else channel_identifier
            )
            member = await context.bot.get_chat_member(
                chat_id=chat_identifier,
                user_id=user_id
            )
            if not _member_is_joined(member):
                unjoined.append((channel_identifier, join_link))
        except Forbidden as e:
            # Usually means the bot cannot inspect membership (e.g. it is not an
            # admin/member of the channel). Fail closed for security and log clearly.
            logger.error(
                "Force Join cannot inspect %s. Make the bot an administrator in that channel. Error: %s",
                channel_identifier, e
            )
            unjoined.append((channel_identifier, join_link))
        except BadRequest as e:
            logger.error(
                "Force Join BadRequest for %s. Check the chat ID/username and bot permissions. Error: %s",
                channel_identifier, e
            )
            unjoined.append((channel_identifier, join_link))
        except TelegramError as e:
            logger.error("Force Join Telegram error for %s: %s", channel_identifier, e)
            unjoined.append((channel_identifier, join_link))
        except Exception as e:
            logger.exception("Unexpected Force Join error for %s: %s", channel_identifier, e)
            unjoined.append((channel_identifier, join_link))

    return (len(unjoined) == 0), unjoined


async def send_force_join_prompt(update: Update, context: ContextTypes.DEFAULT_TYPE,
                                 unjoined_channels: Optional[List[Tuple[str, str]]] = None):
    """Render the force-join buttons consistently from every entry point."""
    user = update.effective_user
    if not user:
        return
    user_id = user.id

    if unjoined_channels is None:
        _, unjoined_channels = await check_force_join(user_id, context)

    buttons = []
    for idx, (_channel_id, link) in enumerate(unjoined_channels, 1):
        if link == "https://t.me/":
            # Do not render a useless URL button for a numeric ID without an invite.
            continue
        buttons.append([
            ikb_btn(f"🔗 {tr(user_id, 'btn_join_channel')} {idx}", url=link)
        ])

    buttons.append([
        ikb_btn(tr(user_id, "btn_verify"), style="success", callback_data="check_force_join_cb")
    ])

    text = tr(user_id, "force_join_msg")
    if not buttons[:-1]:
        text += "\n\n⚠️ Admin has not configured valid join links yet."

    if update.callback_query:
        await update.callback_query.edit_message_text(
            text,
            reply_markup=InlineKeyboardMarkup(buttons),
            parse_mode=ParseMode.MARKDOWN
        )
    elif update.message:
        await update.message.reply_text(
            text,
            reply_markup=InlineKeyboardMarkup(buttons),
            parse_mode=ParseMode.MARKDOWN
        )


def generate_random_username() -> str:
    prefix = "ig_"
    chars = "abcdefghijklmnopqrstuvwxyz0123456789"
    target_len = random.choice([13, 14])
    random_str = ''.join(random.choices(chars, k=target_len - len(prefix)))
    return f"{prefix}{random_str}"

def is_valid_gmail_address(email: str) -> bool:
    """Validate a single Gmail address with an exact gmail.com domain."""
    email = (email or "").strip()
    pattern = r"^[A-Za-z0-9](?:[A-Za-z0-9._%+\-]{0,62}[A-Za-z0-9])?@gmail\.com$"
    return bool(re.fullmatch(pattern, email, flags=re.IGNORECASE))


def get_gmail_done_keyboard(user_id: int = None):
    """Inline buttons dedicated to the Gmail task flow."""
    return InlineKeyboardMarkup([
        [
            ikb_btn(tr(user_id, "btn_done"), style="success", callback_data="gmail_done"),
            ikb_btn(tr(user_id, "btn_cancel"), style="danger", callback_data="gmail_cancel")
        ]
    ])


def is_menu_or_admin_button(text: str) -> bool:
    if not text:
        return False
    known_buttons = [
        "Gmail Work", "New Gmail Work", "Bulk Gmail Work", "Single New Gmail", "Bulk New Gmail", "Old Gmail Work", "Single Old Gmail", "Bulk Old Gmail", "Back to Gmail Work", "Old Gmail Categories", "Wallet", "My Referrals", "Leaderboard",
        "Withdraw", "Support", "Language", "Admin Panel", "Main Menu", "Cancel",
        "Submit Gmail", "Pending Tasks", "Held Tasks", "Approve All", "Reject All",
        "Password Manager", "Price Manager", "Task Search", "User Export",
        "Spreadsheet UI", "Withdraw Requests", "User Management", "Admin Control",
        "System Settings", "Live Stats", "Broadcast", "Backup DB", "Restore DB"
    ]
    return text in known_buttons

FIRST_NAMES = ["Sohan", "Kamal", "Robi", "Arif", "Nayan", "Sumon", "Rashed", "Alim", "Nusrat", "Riya", "Fariha", "Mim"]
LAST_NAMES = ["Ahmed", "Sarker", "Hasan", "Rahman", "Uddin", "Chowdhury", "Akter", "Khatun", "Islam", "Khan"]

# -----------------------------------------------------------------------------
# CONVERSATION STATES
# -----------------------------------------------------------------------------
(
    STATE_TASK_CATEGORY,
    STATE_TASK_INSTAGRAM,
    STATE_TASK_FACEBOOK,
    STATE_TASK_IG_2FA_INPUT,
    STATE_TASK_IG_SEED_INPUT,
    STATE_TASK_FB_2FA_INPUT,
    STATE_TASK_FB_COOKIES_UID,
    STATE_TASK_FB_COOKIES_STR,
    STATE_TASK_GMAIL_INPUT,
    STATE_TASK_2FA_SECRET,
    STATE_TASK_FINAL_SUBMIT,
    STATE_WD_METHOD,
    STATE_WD_NUMBER,
    STATE_WD_AMOUNT,
    STATE_ADMIN_SEARCH_QUERY,
    STATE_ADMIN_EXPORT_USER,
    STATE_ADMIN_BROADCAST,
    STATE_ADMIN_REJECT_REASON,
    STATE_ADMIN_SETTING_VAL,
    STATE_ADMIN_BULK_APPROVE,
    STATE_ADMIN_BULK_REJECT,
    STATE_ADMIN_USER_MGMT,
    STATE_ADMIN_ADD_SUBADMIN,
    STATE_ADMIN_REM_SUBADMIN,
    STATE_ADMIN_BAN_USER,
    STATE_ADMIN_UNBAN_USER,
    STATE_ADMIN_ADD_BAL,
    STATE_ADMIN_SUB_BAL,
    STATE_ADMIN_USER_SEARCH,
    STATE_ADMIN_FORCE_CHANNELS_EDIT,
) = range(30)
STATE_ADMIN_SET_REF_BONUS = 30
STATE_TASK_BULK_GMAIL_INPUT = 31
STATE_ADMIN_RESTORE_DB = 32
STATE_ADMIN_WELCOME_BANNER = 33
STATE_TASK_OLD_MENU = 34
STATE_TASK_OLD_CAT = 35
STATE_TASK_OLD_SINGLE_INPUT = 36
STATE_TASK_OLD_BULK_INPUT = 37
STATE_ADMIN_OLD_CAT_ADD = 38
STATE_TASK_NEW_MENU = 39

# -----------------------------------------------------------------------------
# DYNAMIC MULTILINGUAL REPLY KEYBOARDS
# -----------------------------------------------------------------------------

def get_main_keyboard(user_id: int = None, lang: str = None) -> ReplyKeyboardMarkup:
    l = lang or get_user_lang(user_id)
    keyboard = [
        [kb_btn(tr(l, "btn_task"), style="success")],
        [kb_btn(tr(l, "btn_wallet"), style="primary"), kb_btn(tr(l, "btn_withdraw"), style="danger")],
        [kb_btn(tr(l, "btn_ref"), style="danger"), kb_btn(tr(l, "btn_leaderboard"), style="success")],
        [kb_btn(tr(l, "btn_support"), style="success"), kb_btn(tr(l, "btn_lang"), style="primary")],
    ]
    if user_id and is_admin(user_id):
        keyboard.append([kb_btn(tr(l, "btn_admin_panel"), style="danger")])
    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True)

def get_task_category_keyboard(user_id: int = None, lang: str = None) -> ReplyKeyboardMarkup:
    l = lang or get_user_lang(user_id)
    keyboard = [
        [kb_btn(tr(l, "btn_gmail"), style="success"), kb_btn(tr(l, "btn_old_gmail"), style="primary")],
        [kb_btn(tr(l, "btn_main_menu"), style="danger")]
    ]
    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True)

def get_new_gmail_menu_keyboard(user_id: int = None, lang: str = None) -> ReplyKeyboardMarkup:
    l = lang or get_user_lang(user_id)
    keyboard = [
        [kb_btn(tr(l, "btn_new_single"), style="success"), kb_btn(tr(l, "btn_new_bulk"), style="primary")],
        [kb_btn(tr(l, "btn_back_gmail_work"), style="danger"), kb_btn(tr(l, "btn_main_menu"), style="primary")]
    ]
    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True)

def get_old_gmail_menu_keyboard(user_id: int = None, lang: str = None) -> ReplyKeyboardMarkup:
    l = lang or get_user_lang(user_id)
    keyboard = [
        [kb_btn(tr(l, "btn_old_single"), style="success"), kb_btn(tr(l, "btn_old_bulk"), style="primary")],
        [kb_btn(tr(l, "btn_back_gmail_work"), style="danger"), kb_btn(tr(l, "btn_main_menu"), style="primary")]
    ]
    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True)

def get_task_done_keyboard(user_id: int = None, lang: str = None) -> ReplyKeyboardMarkup:
    l = lang or get_user_lang(user_id)
    keyboard = [
        [kb_btn(tr(l, "btn_done"))],
        [kb_btn(tr(l, "btn_cancel"), style="danger"), kb_btn(tr(l, "btn_main_menu"))]
    ]
    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True)

get_seed_done_keyboard = get_task_done_keyboard

def get_wallet_keyboard(user_id: int = None, lang: str = None) -> ReplyKeyboardMarkup:
    l = lang or get_user_lang(user_id)
    keyboard = [
        [kb_btn(tr(l, "btn_withdraw"), style="success"), kb_btn(tr(l, "btn_main_menu"), style="primary")]
    ]
    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True)

def get_withdraw_keyboard(user_id: int = None, lang: str = None) -> ReplyKeyboardMarkup:
    l = lang or get_user_lang(user_id)
    bkash_active = get_setting_val("wd_bkash_active", "ON") == "ON"
    nagad_active = get_setting_val("wd_nagad_active", "ON") == "ON"
    usdt_active = get_setting_val("wd_usdt_active", "ON") == "ON"

    min_b = float(get_setting_val("min_withdraw_bkash", "50.0"))
    chg_b = float(get_setting_val("bkash_charge_bdt", "0.0"))
    min_n = float(get_setting_val("min_withdraw_nagad", "50.0"))
    chg_n = float(get_setting_val("nagad_charge_bdt", "0.0"))

    b_label = f"{tr(l, 'btn_bkash')} -> Min: {min_b:g}৳" + (f" (-{chg_b:g}৳)" if chg_b > 0 else "")
    n_label = f"{tr(l, 'btn_nagad')} -> Min: {min_n:g}৳" + (f" (-{chg_n:g}৳)" if chg_n > 0 else "")

    keyboard = []
    if bkash_active:
        keyboard.append([kb_btn(b_label, style="danger")])
    if nagad_active:
        keyboard.append([kb_btn(n_label, style="primary")])

    if usdt_active:
        keyboard.append([
            kb_btn(tr(l, "btn_usdt"), style="success"),
            kb_btn(tr(l, "btn_binance_uid"), style="primary")
        ])

    keyboard.append([kb_btn(tr(l, "btn_main_menu"), style="danger")])
    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True)

def get_cancel_keyboard(user_id: int = None, lang: str = None) -> ReplyKeyboardMarkup:
    l = lang or get_user_lang(user_id)
    keyboard = [
        [kb_btn(tr(l, "btn_cancel"), style="danger"), kb_btn(tr(l, "btn_main_menu"))]
    ]
    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True)

def get_admin_keyboard(user_id: int = None, lang: str = None) -> ReplyKeyboardMarkup:
    l = lang or get_user_lang(user_id)
    keyboard = [
        [kb_btn(tr(l, "btn_pending_tasks"), style="danger"), kb_btn(tr(l, "btn_held_tasks"), style="primary")],
        [kb_btn(tr(l, "btn_bulk_approve"), style="success"), kb_btn(tr(l, "btn_bulk_reject"), style="danger")],
        [kb_btn(tr(l, "btn_pwd_mgr"), style="primary"), kb_btn(tr(l, "btn_price_mgr"), style="success")],
        [kb_btn(tr(l, "btn_task_search"), style="danger"), kb_btn(tr(l, "btn_user_export"), style="primary")],
        [kb_btn(tr(l, "btn_excel_export"), style="success"), kb_btn(tr(l, "btn_wd_reqs"), style="danger")],
        [kb_btn(tr(l, "btn_user_mgmt"), style="primary"), kb_btn(tr(l, "btn_admin_mgmt"), style="success")],
        [kb_btn(tr(l, "btn_sys_settings"), style="danger"), kb_btn(tr(l, "btn_live_stats"), style="primary")],
        [kb_btn(tr(l, "btn_broadcast"), style="success"), kb_btn(tr(l, "btn_backup_db"), style="primary")],
        [kb_btn(tr(l, "btn_old_cats"), style="danger"), kb_btn(tr(l, "btn_restore_db"), style="success")],
        [kb_btn(tr(l, "btn_main_menu"), style="primary")],
    ]
    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True)

def clear_user_state(context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()

async def safe_delete_message(context: ContextTypes.DEFAULT_TYPE, chat_id: int, message_id: int):
    """Safely delete message without throwing errors"""
    try:
        await context.bot.delete_message(chat_id=chat_id, message_id=message_id)
    except Exception:
        pass

# -----------------------------------------------------------------------------
# START COMMAND & FORCE JOIN
# -----------------------------------------------------------------------------

async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    clear_user_state(context)
    user = update.effective_user
    user_id = user.id

    if get_setting_val("maintenance_mode") == "ON" and not is_admin(user_id):
        await update.message.reply_text(
            tr(user_id, "maintenance_msg"),
            parse_mode=ParseMode.MARKDOWN
        )
        return ConversationHandler.END

    referred_by_id = None
    if context.args and len(context.args) > 0 and context.args[0].isdigit():
        referred_by_id = int(context.args[0])

    is_new_user = (get_user_data(user_id) is None)
    register_user(user, referred_by_id)

    # INSTANT REFERRAL NOTIFICATION TO REFERRER
    if is_new_user and referred_by_id and referred_by_id != user_id:
        try:
            ref_user = get_user_data(referred_by_id)
            if ref_user:
                await context.bot.send_message(
                    chat_id=referred_by_id,
                    text=tr(referred_by_id, "ref_new_user_joined", name=html.escape(user.first_name), user_id=user_id),
                    parse_mode=ParseMode.HTML
                )
        except Exception as e:
            logger.error(f"Failed to send referral alert to {referred_by_id}: {e}")

    if is_banned(user_id):
        await update.message.reply_text(tr(user_id, "banned_msg"))
        return ConversationHandler.END

    is_joined, unjoined_channels = await check_force_join(user_id, context)
    if not is_joined:
        await send_force_join_prompt(update, context, unjoined_channels)
        return ConversationHandler.END

    welcome_msg = tr(user_id, "welcome_msg", name=html.escape(user.first_name), bot_name=get_setting_val("bot_display_name", "Gmail Work BD"))

    banner_id = get_setting_val("welcome_banner_file_id", "").strip()
    if banner_id:
        try:
            await update.message.reply_photo(
                photo=banner_id,
                caption=welcome_msg,
                reply_markup=get_main_keyboard(user_id),
                parse_mode=ParseMode.MARKDOWN
            )
        except Exception as e:
            logger.warning(f"Failed to send welcome banner photo: {e}")
            await update.message.reply_text(
                welcome_msg,
                reply_markup=get_main_keyboard(user_id),
                parse_mode=ParseMode.MARKDOWN
            )
    else:
        await update.message.reply_text(
            welcome_msg,
            reply_markup=get_main_keyboard(user_id),
            parse_mode=ParseMode.MARKDOWN
        )
    return ConversationHandler.END

async def check_force_join_cb(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    user_id = query.from_user.id

    is_joined, unjoined_channels = await check_force_join(user_id, context)
    if is_joined:
        await query.answer(tr(user_id, "force_join_success"), show_alert=True)
        await safe_delete_message(context, query.message.chat_id, query.message.message_id)
        await context.bot.send_message(
            chat_id=user_id,
            text=tr(user_id, "force_join_success"),
            reply_markup=get_main_keyboard(user_id),
            parse_mode=ParseMode.MARKDOWN
        )
    else:
        await query.answer(tr(user_id, "force_join_fail"), show_alert=True)
        # Keep the verification UI visible so the user can immediately retry.
        try:
            await send_force_join_prompt(update, context, unjoined_channels)
        except Exception as e:
            logger.error("Failed to refresh Force Join prompt: %s", e)

# -----------------------------------------------------------------------------
# TASK WORKFLOWS
# -----------------------------------------------------------------------------

async def handle_task_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    clear_user_state(context)
    user_id = update.effective_user.id
    if is_banned(user_id):
        await update.message.reply_text(tr(user_id, "banned_msg"))
        return ConversationHandler.END

    is_joined, _ = await check_force_join(user_id, context)
    if not is_joined:
        await start_command(update, context)
        return ConversationHandler.END

    await update.message.reply_text(
        tr(user_id, "task_cat_prompt"),
        reply_markup=get_task_category_keyboard(user_id),
        parse_mode=ParseMode.MARKDOWN
    )
    return STATE_TASK_CATEGORY


async def handle_new_gmail_submenu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    clear_user_state(context)
    user_id = update.effective_user.id
    if is_banned(user_id):
        await update.message.reply_text(tr(user_id, "banned_msg"))
        return ConversationHandler.END

    is_joined, _ = await check_force_join(user_id, context)
    if not is_joined:
        await start_command(update, context)
        return ConversationHandler.END

    if get_setting_val("gmail_task_active", "ON") == "OFF" and get_setting_val("bulk_gmail_task_active", "ON") == "OFF":
        await update.message.reply_text("❌ New Gmail Work বর্তমানে বন্ধ রয়েছে।", reply_markup=get_task_category_keyboard(user_id))
        return STATE_TASK_CATEGORY

    await update.message.reply_text(
        "✉️ <b>New Gmail Work</b>\n━━━━━━━━━━━━━━━━━━━━━━\nSingle অথবা Bulk New Gmail বেছে নিন:",
        reply_markup=get_new_gmail_menu_keyboard(user_id),
        parse_mode=ParseMode.HTML
    )
    return STATE_TASK_NEW_MENU

async def handle_gmail_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    clear_user_state(context)
    user_id = update.effective_user.id
    if is_banned(user_id):
        await update.message.reply_text(tr(user_id, "banned_msg"))
        return ConversationHandler.END

    is_joined, _ = await check_force_join(user_id, context)
    if not is_joined:
        await start_command(update, context)
        return ConversationHandler.END

    if get_setting_val("gmail_task_active", "ON") == "OFF":
        await update.message.reply_text("❌ New Gmail Work বর্তমানে বন্ধ রয়েছে।", reply_markup=get_task_category_keyboard(user_id))
        return STATE_TASK_CATEGORY

    price = get_setting_val("gmail_task_price", "22.00")
    fixed_pass = get_setting_val("gmail_default_password", "aass1122")

    context.user_data["task_type"] = "gmail"
    context.user_data["reward_amount"] = float(price)
    context.user_data["assigned_pass"] = fixed_pass

    context.user_data.pop("gmail_address", None)
    msg = tr(user_id, "gmail_guidelines", price=price, fixed_pass=fixed_pass)
    sent_m = await update.message.reply_text(
        msg,
        reply_markup=get_gmail_done_keyboard(user_id),
        parse_mode=ParseMode.HTML
    )
    context.user_data["cred_msg_id"] = sent_m.message_id
    return STATE_TASK_GMAIL_INPUT

async def process_gmail_submission(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Validate and temporarily store Gmail; submit only from the Done callback."""
    user_id = update.effective_user.id
    gmail_acc = (update.message.text or "").strip()

    if is_menu_or_admin_button(gmail_acc):
        clear_user_state(context)
        return ConversationHandler.END

    if not is_valid_gmail_address(gmail_acc):
        await update.message.reply_text(
            "❌ সঠিক Gmail এড্রেস প্রদান করুন (যেমন: example@gmail.com)!",
            reply_markup=get_gmail_done_keyboard(user_id)
        )
        return STATE_TASK_GMAIL_INPUT

    normalized_gmail = gmail_acc.lower()
    context.user_data["gmail_address"] = normalized_gmail

    await update.message.reply_text(
        f"✅ Gmail এড্রেস গ্রহণ করা হয়েছে: <code>{html.escape(normalized_gmail)}</code>\n\n"
        "এখন কাজটি সাবমিট করতে নিচের 'জমা দিন' বাটনে চাপ দিন।",
        reply_markup=get_gmail_done_keyboard(user_id),
        parse_mode=ParseMode.HTML
    )
    return STATE_TASK_GMAIL_INPUT


async def gmail_done_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Check Gmail, block duplicates, then create the database submission."""
    query = update.callback_query
    user_id = query.from_user.id
    gmail_acc = (context.user_data.get("gmail_address") or "").strip().lower()

    if not gmail_acc or not is_valid_gmail_address(gmail_acc):
        await query.answer(
            "❌ আপনি এখনো Gmail এড্রেস জমা দেননি! আগে সঠিক Gmail এড্রেসটি লিখুন।",
            show_alert=True
        )
        return STATE_TASK_GMAIL_INPUT

    # Prevent the same Gmail from being submitted again while any prior
    # submission is active or already approved.
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT 1
            FROM task_submissions
            WHERE LOWER(TRIM(submitted_username)) = ?
              AND task_type = 'gmail'
              AND status IN ('Pending', 'Hold', 'Approved')
            LIMIT 1
            """,
            (gmail_acc,)
        )
        duplicate = cursor.fetchone()

    if duplicate:
        await query.answer(
            "❌ এই Gmail Address আগে জমা দেওয়া হয়েছে। অন্য Gmail ব্যবহার করুন।",
            show_alert=True
        )
        return STATE_TASK_GMAIL_INPUT

    await query.answer()

    if "cred_msg_id" in context.user_data:
        await safe_delete_message(
            context, query.message.chat_id, context.user_data["cred_msg_id"]
        )

    password = context.user_data.get("assigned_pass", "aass1122")
    reward = context.user_data.get("reward_amount", 22.0)
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute(
            """
            INSERT INTO task_submissions
            (user_id, task_type, submitted_username, submitted_password,
             secret_key_2fa, status, reward_amount, is_archived, created_at, updated_at)
            VALUES (?, 'gmail', ?, ?, 'N/A', 'Pending', ?, 0, ?, ?)
            """,
            (user_id, gmail_acc, password, reward, now, now)
        )
        submission_id = cursor.lastrowid
        conn.commit()

    await query.edit_message_text(
        tr(user_id, "gmail_submitted_success", sub_id=submission_id),
        parse_mode=ParseMode.HTML
    )
    await context.bot.send_message(
        chat_id=user_id,
        text="📋 আপনার Gmail টাস্কটি রিভিউয়ের জন্য পাঠানো হয়েছে।",
        reply_markup=get_main_keyboard(user_id)
    )
    await notify_admins_new_submission(
        context, submission_id, user_id, "gmail",
        gmail_acc, password, "N/A", "", reward
    )

    clear_user_state(context)
    return ConversationHandler.END


async def gmail_cancel_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    clear_user_state(context)
    await query.edit_message_text(tr(query.from_user.id, "operation_cancelled"))
    await context.bot.send_message(
        chat_id=query.message.chat_id,
        text=tr(query.from_user.id, "returning_main_menu"),
        reply_markup=get_main_keyboard(query.from_user.id)
    )
    return ConversationHandler.END


def get_bulk_gmail_done_keyboard(user_id: int = None, count: int = 0):
    submit_label = f"Submit All ({count})" if count > 0 else "Submit All"
    return InlineKeyboardMarkup([
        [
            ikb_btn(submit_label, style="success", callback_data="bulk_gmail_done"),
            ikb_btn(tr(user_id, "btn_cancel"), style="danger", callback_data="bulk_gmail_cancel")
        ]
    ])


async def handle_bulk_gmail_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    clear_user_state(context)
    user_id = update.effective_user.id
    if is_banned(user_id):
        await update.message.reply_text(tr(user_id, "banned_msg"))
        return ConversationHandler.END

    is_joined, _ = await check_force_join(user_id, context)
    if not is_joined:
        await start_command(update, context)
        return ConversationHandler.END

    if get_setting_val("bulk_gmail_task_active", "ON") == "OFF":
        await update.message.reply_text("❌ Bulk Gmail Work বর্তমানে বন্ধ রয়েছে।", reply_markup=get_task_category_keyboard(user_id))
        return STATE_TASK_CATEGORY

    price = get_setting_val("gmail_task_price", "22.00")
    fixed_pass = get_setting_val("gmail_default_password", "aass1122")

    context.user_data["task_type"] = "gmail"
    context.user_data["reward_amount"] = float(price)
    context.user_data["assigned_pass"] = fixed_pass
    context.user_data.pop("bulk_gmails", None)

    msg = tr(user_id, "bulk_gmail_guidelines", price=price, fixed_pass=fixed_pass)
    sent_m = await update.message.reply_text(
        msg,
        reply_markup=get_bulk_gmail_done_keyboard(user_id, 0),
        parse_mode=ParseMode.HTML
    )
    context.user_data["cred_msg_id"] = sent_m.message_id
    return STATE_TASK_BULK_GMAIL_INPUT


async def process_bulk_gmail_submission(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    raw_text = (update.message.text or "").strip()

    if is_menu_or_admin_button(raw_text):
        clear_user_state(context)
        return ConversationHandler.END

    lines = [ln.strip().lower() for ln in re.split(r"[\r\n,\s]+", raw_text) if ln.strip()]
    unique_lines = list(dict.fromkeys(lines))

    valid_gmails = []
    invalid_gmails = []
    dup_gmails = []

    with get_db() as conn:
        cursor = conn.cursor()
        for email in unique_lines:
            if not is_valid_gmail_address(email):
                invalid_gmails.append(email)
                continue
            cursor.execute(
                """
                SELECT 1 FROM task_submissions
                WHERE LOWER(TRIM(submitted_username)) = ?
                  AND task_type = 'gmail'
                  AND status IN ('Pending', 'Hold', 'Approved')
                LIMIT 1
                """,
                (email,)
            )
            if cursor.fetchone():
                dup_gmails.append(email)
            else:
                valid_gmails.append(email)

    if not valid_gmails:
        await update.message.reply_text(
            f"❌ কোনো নতুন বৈধ Gmail পাওয়া যায়নি!\n"
            f"• ভুল ফরম্যাট: {len(invalid_gmails)} টি\n"
            f"• আগে জমা দেওয়া (ডুপ্লিকেট): {len(dup_gmails)} টি\n\n"
            f"সঠিক নতুন Gmail-গুলো প্রতি লাইনে একটি করে পাঠান।",
            reply_markup=get_bulk_gmail_done_keyboard(user_id, 0)
        )
        return STATE_TASK_BULK_GMAIL_INPUT

    context.user_data["bulk_gmails"] = valid_gmails[:50]
    chosen = context.user_data["bulk_gmails"]
    preview = "\n".join(f"• <code>{html.escape(g)}</code>" for g in chosen[:15])
    if len(chosen) > 15:
        preview += f"\n... এবং আরও {len(chosen) - 15} টি"

    extra_note = ""
    if invalid_gmails or dup_gmails:
        extra_note = f"\n⚠️ বাদ গেছে — ভুল: {len(invalid_gmails)} টি | ডুপ্লিকেট: {len(dup_gmails)} টি"

    await update.message.reply_text(
        f"✅ <b>মোট {len(chosen)} টি বৈধ Gmail পাওয়া গেছে:</b>\n{preview}{extra_note}\n\n"
        f"সবগুলো জমা দিতে নিচের <b>Submit All ({len(chosen)})</b> বাটনে চাপ দিন।",
        reply_markup=get_bulk_gmail_done_keyboard(user_id, len(chosen)),
        parse_mode=ParseMode.HTML
    )
    return STATE_TASK_BULK_GMAIL_INPUT


async def bulk_gmail_done_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    user_id = query.from_user.id
    gmails = context.user_data.get("bulk_gmails") or []

    if not gmails:
        await query.answer("❌ আগে প্রতি লাইনে একটি করে Gmail লিখে পাঠান!", show_alert=True)
        return STATE_TASK_BULK_GMAIL_INPUT

    password = context.user_data.get("assigned_pass", "aass1122")
    reward = context.user_data.get("reward_amount", 22.0)
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    submitted_items = []
    with get_db() as conn:
        cursor = conn.cursor()
        for email in gmails:
            cursor.execute(
                """
                SELECT 1 FROM task_submissions
                WHERE LOWER(TRIM(submitted_username)) = ?
                  AND task_type = 'gmail'
                  AND status IN ('Pending', 'Hold', 'Approved')
                LIMIT 1
                """,
                (email,)
            )
            if cursor.fetchone():
                continue
            cursor.execute(
                """
                INSERT INTO task_submissions
                (user_id, task_type, submitted_username, submitted_password,
                 secret_key_2fa, status, reward_amount, is_archived, created_at, updated_at)
                VALUES (?, 'gmail', ?, ?, 'N/A', 'Pending', ?, 0, ?, ?)
                """,
                (user_id, email, password, reward, now, now)
            )
            submitted_items.append((cursor.lastrowid, email))
        conn.commit()

    if not submitted_items:
        await query.answer("❌ এই Gmail-গুলো ইতিমধ্যেই জমা হয়ে গেছে!", show_alert=True)
        return STATE_TASK_BULK_GMAIL_INPUT

    await query.answer()

    if "cred_msg_id" in context.user_data:
        await safe_delete_message(context, query.message.chat_id, context.user_data["cred_msg_id"])

    await query.edit_message_text(
        tr(user_id, "bulk_gmail_submitted_success", count=len(submitted_items)),
        parse_mode=ParseMode.HTML
    )
    await context.bot.send_message(
        chat_id=user_id,
        text="📋 আপনার Bulk Gmail টাস্কগুলো রিভিউয়ের জন্য পাঠানো হয়েছে।",
        reply_markup=get_main_keyboard(user_id)
    )

    admin_id_setting = ADMIN_ID or get_setting_val("admin_id", "")
    admins = [int(x.strip()) for x in admin_id_setting.split(",") if x.strip().isdigit()]
    try:
        with get_db() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT user_id FROM sub_admins")
            for row in cursor.fetchall():
                admins.append(row["user_id"])
    except Exception:
        pass
    admins = list(set(admins))

    lines_str = "\n".join(f"• #{sid} — <code>{html.escape(em)}</code>" for sid, em in submitted_items[:25])
    if len(submitted_items) > 25:
        lines_str += f"\n... (+{len(submitted_items) - 25} more)"

    admin_msg = (
        f"📦 <b>নতুন Bulk Gmail সাবমিশন! ({len(submitted_items)} টি)</b>\n\n"
        f"👤 <b>User ID:</b> <code>{user_id}</code>\n"
        f"🔑 <b>Password:</b> <code>{html.escape(password)}</code>\n"
        f"💵 <b>Rate:</b> ৳{reward:.2f} / each\n\n"
        f"{lines_str}\n\n"
        f"<i>(Pending Tasks বা Spreadsheet UI থেকে এপ্রুভ/রিজেক্ট করুন)</i>"
    )
    for admin in admins:
        try:
            await context.bot.send_message(chat_id=admin, text=admin_msg, parse_mode=ParseMode.HTML)
        except Exception as e:
            logger.error(f"Failed to notify admin {admin} of bulk gmail: {e}")

    clear_user_state(context)
    return ConversationHandler.END


async def bulk_gmail_cancel_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    clear_user_state(context)
    await query.edit_message_text(tr(query.from_user.id, "operation_cancelled"))
    await context.bot.send_message(
        chat_id=query.message.chat_id,
        text=tr(query.from_user.id, "returning_main_menu"),
        reply_markup=get_main_keyboard(query.from_user.id)
    )
    return ConversationHandler.END


# --- OLD GMAIL WORKFLOWS ---
async def handle_old_gmail_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    clear_user_state(context)
    user_id = update.effective_user.id
    if is_banned(user_id):
        await update.message.reply_text(tr(user_id, "banned_msg"))
        return ConversationHandler.END

    is_joined, _ = await check_force_join(user_id, context)
    if not is_joined:
        await start_command(update, context)
        return ConversationHandler.END

    if get_setting_val("old_gmail_task_active", "ON") == "OFF":
        await update.message.reply_text("❌ Old Gmail Work বর্তমানে বন্ধ রয়েছে।", reply_markup=get_task_category_keyboard(user_id))
        return STATE_TASK_CATEGORY

    cats = get_old_categories()
    if not cats:
        await update.message.reply_text(
            "⚠️ অ্যাডমিন এখনো কোনো Old Gmail ক্যাটাগরি যোগ করেননি।",
            reply_markup=get_task_category_keyboard(user_id)
        )
        return STATE_TASK_CATEGORY

    await update.message.reply_text(
        "♻️ <b>Old Gmail Work</b>\n━━━━━━━━━━━━━━━━━━━━━━\nSingle অথবা Bulk Old Gmail বেছে নিন:",
        reply_markup=get_old_gmail_menu_keyboard(user_id),
        parse_mode=ParseMode.HTML
    )
    return STATE_TASK_OLD_MENU


async def handle_old_single_or_bulk_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    btn_text = (update.message.text or "").strip()
    mode = "bulk" if "Bulk" in btn_text else "single"
    context.user_data["old_mode"] = mode

    cats = get_old_categories()
    if not cats:
        await update.message.reply_text("⚠️ কোনো ক্যাটাগরি পাওয়া যায়নি।", reply_markup=get_old_gmail_menu_keyboard(user_id))
        return STATE_TASK_OLD_MENU

    buttons = []
    for c in cats:
        buttons.append([ikb_btn(f"{c['label']} — ৳{float(c['rate']):g}", style="primary", callback_data=f"oldcat_{c['cat_key']}")])
    buttons.append([ikb_btn("Cancel", style="danger", callback_data="oldcat_cancel")])

    title = "📦 <b>Bulk Old Gmail</b>" if mode == "bulk" else "♻️ <b>Single Old Gmail</b>"
    await update.message.reply_text(
        f"{title}\n━━━━━━━━━━━━━━━━━━━━━━\nকোন সাল/ক্যাটাগরির Old Gmail জমা দেবেন? নিচে থেকে বেছে নিন:",
        reply_markup=InlineKeyboardMarkup(buttons),
        parse_mode=ParseMode.HTML
    )
    return STATE_TASK_OLD_CAT


def _old_done_kb(user_id: int, mode: str, count: int = 0):
    if mode == "bulk":
        lbl = f"Submit All ({count})" if count > 0 else "Submit All"
        done_cb = "old_bulk_done"
    else:
        lbl = "Submit Old Gmail"
        done_cb = "old_single_done"
    return InlineKeyboardMarkup([[
        ikb_btn(lbl, style="success", callback_data=done_cb),
        ikb_btn(tr(user_id, "btn_cancel"), style="danger", callback_data="old_gmail_cancel")
    ]])


async def old_category_selected_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    user_id = query.from_user.id
    data = query.data

    if data == "oldcat_cancel":
        await query.answer()
        clear_user_state(context)
        await query.edit_message_text(tr(user_id, "operation_cancelled"))
        await context.bot.send_message(chat_id=user_id, text=tr(user_id, "returning_main_menu"), reply_markup=get_main_keyboard(user_id))
        return ConversationHandler.END

    cat_key = data.replace("oldcat_", "", 1)
    cat = get_old_category(cat_key)
    if not cat:
        await query.answer("❌ ক্যাটাগরি পাওয়া যায়নি!", show_alert=True)
        return STATE_TASK_OLD_CAT

    await query.answer()
    mode = context.user_data.get("old_mode", "single")
    old_pass = get_setting_val("old_gmail_default_password", "oldpass123")
    old_rec = get_setting_val("old_gmail_recovery_email", "").strip()
    twofa_on = get_setting_val("old_gmail_2fa_active", "ON") == "ON"

    context.user_data["old_cat_key"] = cat["cat_key"]
    context.user_data["old_cat_label"] = cat["label"]
    context.user_data["reward_amount"] = float(cat["rate"])
    context.user_data["assigned_pass"] = old_pass
    context.user_data.pop("old_gmail_single", None)
    context.user_data.pop("old_gmail_bulk", None)

    rec_line = f"💌 <b>Recovery Email (ট্যাপ করে কপি):</b>\n<code>{html.escape(old_rec)}</code>\n" if old_rec else ""
    tfa_line = "🔐 <b>2FA নিয়ম:</b> অবশ্যই অ্যাকাউন্টের 2FA অফ করে জমা দিতে হবে।\n" if twofa_on else ""
    step_line = (
        "📝 সবগুলো Gmail <b>প্রতি লাইনে ১টি করে</b> নিচে লিখে পাঠান:\n\n<b>উদাহরণ:</b>\n<code>oldmail1@gmail.com\noldmail2@gmail.com</code>"
        if mode == "bulk" else
        "📝 আপনার Old Gmail অ্যাড্রেসটি নিচে লিখে পাঠান:\n\n<b>উদাহরণ:</b> <code>oldmail2020@gmail.com</code>"
    )

    msg = (
        f"♻️ <b>{html.escape(cat['label'])}</b> ({'Bulk' if mode=='bulk' else 'Single'})\n"
        f"━━━━━━━━━━━━━━━━━━━━━━\n"
        f"💵 <b>রেট:</b> ৳{float(cat['rate']):g} / each\n"
        f"🔑 <b>পাসওয়ার্ড (ট্যাপ করে কপি):</b> <code>{html.escape(old_pass)}</code>\n"
        f"{rec_line}{tfa_line}"
        f"━━━━━━━━━━━━━━━━━━━━━━\n"
        f"অ্যাকাউন্টে ওপরের পাসওয়ার্ড{(' ও রিকভারি' if old_rec else '')} সেট করে:\n{step_line}"
    )

    sent_m = await query.edit_message_text(msg, reply_markup=_old_done_kb(user_id, mode, 0), parse_mode=ParseMode.HTML)
    context.user_data["cred_msg_id"] = sent_m.message_id
    return STATE_TASK_OLD_BULK_INPUT if mode == "bulk" else STATE_TASK_OLD_SINGLE_INPUT


async def process_old_single_submission(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    email = (update.message.text or "").strip()

    if is_menu_or_admin_button(email):
        clear_user_state(context)
        return ConversationHandler.END

    if not is_valid_gmail_address(email):
        await update.message.reply_text("❌ সঠিক Gmail এড্রেস দিন (যেমন: example@gmail.com)!", reply_markup=_old_done_kb(user_id, "single"))
        return STATE_TASK_OLD_SINGLE_INPUT

    email = email.lower()
    context.user_data["old_gmail_single"] = email
    await update.message.reply_text(
        f"✅ Old Gmail গ্রহণ করা হয়েছে: <code>{html.escape(email)}</code>\n\nজমা দিতে নিচের <b>Submit Old Gmail</b> বাটনে চাপ দিন।",
        reply_markup=_old_done_kb(user_id, "single"),
        parse_mode=ParseMode.HTML
    )
    return STATE_TASK_OLD_SINGLE_INPUT


async def old_single_done_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    user_id = query.from_user.id
    email = (context.user_data.get("old_gmail_single") or "").strip().lower()

    if not email or not is_valid_gmail_address(email):
        await query.answer("❌ আগে সঠিক Old Gmail এড্রেসটি লিখুন!", show_alert=True)
        return STATE_TASK_OLD_SINGLE_INPUT

    with get_db() as conn:
        cursor = conn.cursor()
        dup = cursor.execute(
            "SELECT 1 FROM task_submissions WHERE LOWER(TRIM(submitted_username)) = ? AND status IN ('Pending','Hold','Approved') LIMIT 1",
            (email,)
        ).fetchone()
        if dup:
            await query.answer("❌ এই Gmail আগে জমা দেওয়া হয়েছে!", show_alert=True)
            return STATE_TASK_OLD_SINGLE_INPUT

        await query.answer()
        if "cred_msg_id" in context.user_data:
            await safe_delete_message(context, query.message.chat_id, context.user_data["cred_msg_id"])

        pwd = context.user_data.get("assigned_pass", "oldpass123")
        reward = context.user_data.get("reward_amount", 30.0)
        cat_label = context.user_data.get("old_cat_label", "Old Gmail")
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cursor.execute(
            """
            INSERT INTO task_submissions
            (user_id, task_type, submitted_username, submitted_password,
             secret_key_2fa, status, reward_amount, is_archived, category_label, created_at, updated_at)
            VALUES (?, 'old_gmail', ?, ?, 'N/A', 'Pending', ?, 0, ?, ?, ?)
            """,
            (user_id, email, pwd, reward, cat_label, now, now)
        )
        sub_id = cursor.lastrowid
        conn.commit()

    await query.edit_message_text(tr(user_id, "gmail_submitted_success", sub_id=sub_id), parse_mode=ParseMode.HTML)
    await context.bot.send_message(chat_id=user_id, text="📋 আপনার Old Gmail টাস্কটি রিভিউয়ের জন্য পাঠানো হয়েছে।", reply_markup=get_main_keyboard(user_id))
    await notify_admins_new_submission(context, sub_id, user_id, f"old_gmail ({cat_label})", email, pwd, "N/A", "", reward)
    clear_user_state(context)
    return ConversationHandler.END


async def process_old_bulk_submission(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    raw_text = (update.message.text or "").strip()

    if is_menu_or_admin_button(raw_text):
        clear_user_state(context)
        return ConversationHandler.END

    lines = [ln.strip().lower() for ln in re.split(r"[\r\n,\s]+", raw_text) if ln.strip()]
    unique_lines = list(dict.fromkeys(lines))
    valid_gmails, invalid_gmails, dup_gmails = [], [], []

    with get_db() as conn:
        cursor = conn.cursor()
        for em in unique_lines:
            if not is_valid_gmail_address(em):
                invalid_gmails.append(em)
                continue
            if cursor.execute("SELECT 1 FROM task_submissions WHERE LOWER(TRIM(submitted_username)) = ? AND status IN ('Pending','Hold','Approved') LIMIT 1", (em,)).fetchone():
                dup_gmails.append(em)
            else:
                valid_gmails.append(em)

    if not valid_gmails:
        await update.message.reply_text(
            f"❌ কোনো নতুন বৈধ Old Gmail পাওয়া যায়নি! (ভুল: {len(invalid_gmails)} | ডুপ্লিকেট: {len(dup_gmails)})",
            reply_markup=_old_done_kb(user_id, "bulk", 0)
        )
        return STATE_TASK_OLD_BULK_INPUT

    context.user_data["old_gmail_bulk"] = valid_gmails[:50]
    chosen = context.user_data["old_gmail_bulk"]
    preview = "\n".join(f"• <code>{html.escape(g)}</code>" for g in chosen[:15])
    if len(chosen) > 15:
        preview += f"\n... এবং আরও {len(chosen) - 15} টি"

    await update.message.reply_text(
        f"✅ <b>মোট {len(chosen)} টি বৈধ Old Gmail পাওয়া গেছে:</b>\n{preview}\n\nজমা দিতে <b>Submit All ({len(chosen)})</b> চাপুন।",
        reply_markup=_old_done_kb(user_id, "bulk", len(chosen)),
        parse_mode=ParseMode.HTML
    )
    return STATE_TASK_OLD_BULK_INPUT


async def old_bulk_done_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    user_id = query.from_user.id
    gmails = context.user_data.get("old_gmail_bulk") or []
    if not gmails:
        await query.answer("❌ আগে প্রতি লাইনে একটি করে Gmail লিখে পাঠান!", show_alert=True)
        return STATE_TASK_OLD_BULK_INPUT

    pwd = context.user_data.get("assigned_pass", "oldpass123")
    reward = context.user_data.get("reward_amount", 30.0)
    cat_label = context.user_data.get("old_cat_label", "Old Gmail")
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    submitted_items = []

    with get_db() as conn:
        cursor = conn.cursor()
        for em in gmails:
            if cursor.execute("SELECT 1 FROM task_submissions WHERE LOWER(TRIM(submitted_username)) = ? AND status IN ('Pending','Hold','Approved') LIMIT 1", (em,)).fetchone():
                continue
            cursor.execute(
                """
                INSERT INTO task_submissions
                (user_id, task_type, submitted_username, submitted_password,
                 secret_key_2fa, status, reward_amount, is_archived, category_label, created_at, updated_at)
                VALUES (?, 'old_gmail', ?, ?, 'N/A', 'Pending', ?, 0, ?, ?, ?)
                """,
                (user_id, em, pwd, reward, cat_label, now, now)
            )
            submitted_items.append((cursor.lastrowid, em))
        conn.commit()

    if not submitted_items:
        await query.answer("❌ এই Gmail-গুলো ইতিমধ্যেই জমা হয়ে গেছে!", show_alert=True)
        return STATE_TASK_OLD_BULK_INPUT

    await query.answer()
    if "cred_msg_id" in context.user_data:
        await safe_delete_message(context, query.message.chat_id, context.user_data["cred_msg_id"])

    await query.edit_message_text(tr(user_id, "bulk_gmail_submitted_success", count=len(submitted_items)), parse_mode=ParseMode.HTML)
    await context.bot.send_message(chat_id=user_id, text="📋 আপনার Bulk Old Gmail টাস্কগুলো রিভিউয়ের জন্য পাঠানো হয়েছে।", reply_markup=get_main_keyboard(user_id))

    admin_id_setting = ADMIN_ID or get_setting_val("admin_id", "")
    admins = list({int(x.strip()) for x in admin_id_setting.split(",") if x.strip().isdigit()})
    lines_str = "\n".join(f"• #{sid} — <code>{html.escape(em)}</code>" for sid, em in submitted_items[:25])
    if len(submitted_items) > 25:
        lines_str += f"\n... (+{len(submitted_items) - 25} more)"
    admin_msg = (
        f"♻️ <b>নতুন Bulk Old Gmail ({html.escape(cat_label)}) — {len(submitted_items)} টি!</b>\n\n"
        f"👤 <b>User ID:</b> <code>{user_id}</code>\n"
        f"🔑 <b>Password:</b> <code>{html.escape(pwd)}</code>\n"
        f"💵 <b>Rate:</b> ৳{reward:.2f} / each\n\n{lines_str}"
    )
    for admin in admins:
        try:
            await context.bot.send_message(chat_id=admin, text=admin_msg, parse_mode=ParseMode.HTML)
        except Exception:
            pass

    clear_user_state(context)
    return ConversationHandler.END


async def old_gmail_cancel_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    clear_user_state(context)
    await query.edit_message_text(tr(query.from_user.id, "operation_cancelled"))
    await context.bot.send_message(chat_id=query.message.chat_id, text=tr(query.from_user.id, "returning_main_menu"), reply_markup=get_main_keyboard(query.from_user.id))
    return ConversationHandler.END

# --- ADMIN NOTIFICATION HELPER ---
async def notify_admins_new_submission(context, sub_id, user_id, task_type, username, password, secret_key, cookies, reward, file_id=None):
    admin_id_setting = ADMIN_ID or get_setting_val("admin_id", "")
    admins = [int(x.strip()) for x in admin_id_setting.split(",") if x.strip().isdigit()]

    try:
        with get_db() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT user_id FROM sub_admins")
            for row in cursor.fetchall():
                admins.append(row["user_id"])
    except Exception:
        pass
    admins = list(set(admins))

    task_icon = "📸" if "ig" in task_type else "📘" if "fb" in task_type else "✉️"
    task_name = {
        "ig_2fa": "Instagram 2FA",
        "ig_seed": "Instagram Seed",
        "fb_2fa": "Facebook 2FA",
        "fb_cookies": "Facebook Cookies",
        "gmail": "Gmail",
            "old_gmail": "Old Gmail",
                "old_gmail": "Old Gmail"
    }.get(task_type, task_type.upper())

    admin_msg = (
        f"📥 **নতুন টাস্ক সাবমিট হয়েছে! / New Task Submitted!**\n\n"
        f"🆔 **টাস্ক আইডি / Task ID:** #{sub_id}\n"
        f"👤 **ইউজার আইডি / User ID:** <code>{user_id}</code>\n"
        f"📱 **টাইপ / Type:** {task_icon} {task_name}\n"
        f"👤 **সাবমিশন / Submission:** <code>{username}</code>\n"
        f"🔑 **পাসওয়ার্ড / Password:** <code>{password}</code>\n"
    )
    if secret_key and secret_key != "N/A":
        admin_msg += f"🔐 **2FA Secret:** <code>{secret_key}</code>\n"
    if cookies:
        admin_msg += f"🍪 **Cookies:** <code>{cookies[:50]}...</code>\n"
    admin_msg += f"💵 **মূল্য / Price:** ৳{reward:.2f}"

    buttons = [
        [
            ikb_btn("✅ Approve", style="success", callback_data=f"adm_app_sub_{sub_id}"),
            ikb_btn("⏳ Hold", callback_data=f"adm_hold_sub_{sub_id}")
        ],
        [ikb_btn("❌ Reject", style="danger", callback_data=f"adm_rej_prompt_{sub_id}")]
    ]

    for admin in admins:
        try:
            if file_id:
                await context.bot.send_document(
                    chat_id=admin,
                    document=file_id,
                    caption=admin_msg,
                    reply_markup=InlineKeyboardMarkup(buttons),
                    parse_mode=ParseMode.HTML
                )
            else:
                await context.bot.send_message(
                    chat_id=admin,
                    text=admin_msg,
                    reply_markup=InlineKeyboardMarkup(buttons),
                    parse_mode=ParseMode.HTML
                )
        except Exception as e:
            logger.error(f"Failed to notify admin {admin}: {e}")

# -----------------------------------------------------------------------------
# WALLET & WITHDRAWAL SYSTEM
# -----------------------------------------------------------------------------

async def handle_wallet_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    clear_user_state(context)
    user_id = update.effective_user.id
    if is_banned(user_id):
        return ConversationHandler.END

    is_joined, _ = await check_force_join(user_id, context)
    if not is_joined:
        await start_command(update, context)
        return ConversationHandler.END

    user_data = get_user_data(user_id)
    if not user_data:
        register_user(update.effective_user)
        user_data = get_user_data(user_id)

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT SUM(amount) FROM withdrawals WHERE user_id = ? AND status = 'Pending'", (user_id,))
        p_row = cursor.fetchone()
        pending_withdraw = p_row[0] if p_row[0] else 0.0

        cursor.execute("SELECT SUM(reward_amount) FROM task_submissions WHERE user_id = ? AND status = 'Approved'", (user_id,))
        a_row = cursor.fetchone()
        approved_earnings = a_row[0] if a_row[0] else 0.0

        cursor.execute("SELECT COUNT(*) FROM task_submissions WHERE user_id = ? AND status = 'Approved'", (user_id,))
        completed_tasks = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM task_submissions WHERE user_id = ? AND status = 'Pending'", (user_id,))
        pending_tasks = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM task_submissions WHERE user_id = ? AND status = 'Hold'", (user_id,))
        held_tasks = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM task_submissions WHERE user_id = ? AND status = 'Rejected'", (user_id,))
        rejected_tasks = cursor.fetchone()[0]

    lifetime_earnings = approved_earnings + user_data["ref_earnings"]
    usdt_rate = max(float(get_setting_val("usdt_rate_bdt", "120.0")), 0.000001)
    usdt_balance = float(user_data["balance"]) / usdt_rate

    wallet_text = tr(
        user_id, "wallet_text",
        balance=user_data['balance'],
        usdt_balance=usdt_balance,
        usdt_rate=usdt_rate,
        pending_wd=pending_withdraw,
        lifetime=lifetime_earnings,
        ref_earnings=user_data['ref_earnings'],
        completed=completed_tasks,
        pending=pending_tasks,
        held=held_tasks,
        rejected=rejected_tasks
    )

    await update.message.reply_text(
        wallet_text,
        reply_markup=get_wallet_keyboard(user_id),
        parse_mode=ParseMode.MARKDOWN
    )
    return ConversationHandler.END

async def handle_withdraw_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    clear_user_state(context)
    user_id = update.effective_user.id
    if is_banned(user_id):
        return ConversationHandler.END

    is_joined, _ = await check_force_join(user_id, context)
    if not is_joined:
        await start_command(update, context)
        return ConversationHandler.END

    bkash_active = get_setting_val("wd_bkash_active", "ON") == "ON"
    nagad_active = get_setting_val("wd_nagad_active", "ON") == "ON"
    usdt_active = get_setting_val("wd_usdt_active", "ON") == "ON"

    if not bkash_active and not nagad_active and not usdt_active:
        await update.message.reply_text(tr(user_id, "wd_all_disabled"), reply_markup=get_main_keyboard(user_id))
        return ConversationHandler.END

    await update.message.reply_text(
        tr(user_id, "wd_gateway_prompt"),
        reply_markup=get_withdraw_keyboard(user_id)
    )
    return STATE_WD_METHOD

async def handle_withdraw_method(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    user_id = update.effective_user.id

    if text in ["🏠 মূল মেনু", "🏠 Main Menu", "🔙 মূল মেনু", "মূল মেনু", "🔙 Main Menu", "Main Menu"]:
        await update.message.reply_text(tr(user_id, "returning_main_menu"), reply_markup=get_main_keyboard(user_id))
        return ConversationHandler.END

    method = None
    if text.startswith("Bkash") or text.startswith("বিকাশ"):
        if get_setting_val("wd_bkash_active", "ON") == "ON":
            method = "bkash"
    elif text.startswith("Nagad") or text.startswith("নগদ"):
        if get_setting_val("wd_nagad_active", "ON") == "ON":
            method = "nagad"
    elif text in ["USDT (BEP-20)"]:
        if get_setting_val("wd_usdt_active", "ON") == "ON":
            method = "usdt"
    elif text in ["Binance UID"]:
        if get_setting_val("wd_usdt_active", "ON") == "ON":
            method = "binance_uid"

    if not method:
        await update.message.reply_text(tr(user_id, "wd_invalid_method"), reply_markup=get_withdraw_keyboard(user_id))
        return STATE_WD_METHOD

    context.user_data["wd_method"] = method

    if method == "usdt":
        msg = tr(user_id, "wd_enter_account_usdt")
    elif method == "binance_uid":
        msg = tr(user_id, "wd_enter_account_binance_uid")
    else:
        msg = tr(user_id, "wd_enter_account_mfs", method=method.upper())

    await update.message.reply_text(msg, reply_markup=get_cancel_keyboard(user_id))
    return STATE_WD_NUMBER

async def handle_withdraw_number(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    number = update.message.text.strip()

    if number in ["❌ বাতিল", "❌ Cancel", "🏠 মূল মেনু", "🏠 Main Menu", "❌ বাতিল", "❌ Cancel", "🔙 মূল মেনু", "মূল মেনু", "🔙 Main Menu", "Main Menu"]:
        await update.message.reply_text(tr(user_id, "operation_cancelled"), reply_markup=get_main_keyboard(user_id))
        return ConversationHandler.END

    method = context.user_data.get("wd_method")

    if method == "binance_uid":
        if not re.fullmatch(r"[0-9]{5,20}", number):
            await update.message.reply_text("❌ সঠিক Binance UID দিন (শুধু ৫-২০ ডিজিট)।", reply_markup=get_cancel_keyboard(user_id))
            return STATE_WD_NUMBER
    elif method == "usdt":
        if not re.fullmatch(r"0x[a-fA-F0-9]{40}", number):
            await update.message.reply_text("❌ সঠিক USDT BEP-20 wallet address দিন (0x দিয়ে শুরু, মোট 42 অক্ষর)।", reply_markup=get_cancel_keyboard(user_id))
            return STATE_WD_NUMBER
    elif not re.match(r"^[0-9]{11,14}$", number):
        await update.message.reply_text(tr(user_id, "wd_invalid_account"), reply_markup=get_cancel_keyboard(user_id))
        return STATE_WD_NUMBER

    context.user_data["wd_number"] = number
    user_data = get_user_data(user_id)

    limit_method = "usdt" if method in ("usdt", "binance_uid") else method
    min_limit = float(get_setting_val(f"min_withdraw_{limit_method}", "50.0"))
    max_limit = float(get_setting_val(f"max_withdraw_{limit_method}", "5000.0"))

    if method in ("usdt", "binance_uid"):
        usdt_rate = max(float(get_setting_val("usdt_rate_bdt", "120.0")), 0.000001)
        usdt_balance = float(user_data["balance"]) / usdt_rate
        charge_status = get_setting_val("usdt_charge_enabled", "ON") == "ON"
        charge_usdt = float(get_setting_val("usdt_charge_usdt", "0.05")) if charge_status else 0.0
        msg = (
            "💵 **USDT withdrawal amount লিখুন:**\n"
            "━━━━━━━━━━━━━━━━━━━━━━\n"
            f"💰 Available: **${usdt_balance:.4f} USDT**\n"
            f"💱 Rate: **৳{usdt_rate:.2f} = 1 USDT**\n"
            f"🔻 Minimum: **{min_limit} USDT**\n"
            f"🔺 Maximum: **{max_limit} USDT**\n"
            f"💳 Charge: **{charge_usdt:.4f} USDT**" if charge_status else
            "💵 **USDT withdrawal amount লিখুন:**\n"
            "━━━━━━━━━━━━━━━━━━━━━━\n"
            f"💰 Available: **${usdt_balance:.4f} USDT**\n"
            f"💱 Rate: **৳{usdt_rate:.2f} = 1 USDT**\n"
            f"🔻 Minimum: **{min_limit} USDT**\n"
            f"🔺 Maximum: **{max_limit} USDT**\n"
            "💳 Charge: **OFF**"
        )
    else:
        chg_bdt = float(get_setting_val(f"{method}_charge_bdt", "0.0"))
        msg = tr(user_id, "wd_enter_amount", balance=user_data['balance'], min_limit=min_limit, max_limit=max_limit)
        if chg_bdt > 0:
            msg += f"\n💳 **Charge ({method.upper()}):** ৳{chg_bdt:g}"

    await update.message.reply_text(msg, reply_markup=get_cancel_keyboard(user_id), parse_mode=ParseMode.MARKDOWN)
    return STATE_WD_AMOUNT

async def handle_withdraw_amount(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    amount_str = update.message.text.strip()

    if amount_str in ["❌ বাতিল", "❌ Cancel", "🏠 মূল মেনু", "🏠 Main Menu", "❌ বাতিল", "❌ Cancel", "🔙 মূল মেনু", "মূল মেনু", "🔙 Main Menu", "Main Menu"]:
        await update.message.reply_text(tr(user_id, "operation_cancelled"), reply_markup=get_main_keyboard(user_id))
        return ConversationHandler.END

    try:
        amount = float(amount_str)
    except ValueError:
        await update.message.reply_text(tr(user_id, "wd_invalid_amount_number"), reply_markup=get_cancel_keyboard(user_id))
        return STATE_WD_AMOUNT

    user_data = get_user_data(user_id)
    method = context.user_data.get("wd_method")

    limit_method = "usdt" if method in ("usdt", "binance_uid") else method
    min_limit = float(get_setting_val(f"min_withdraw_{limit_method}", "50.0"))
    max_limit = float(get_setting_val(f"max_withdraw_{limit_method}", "5000.0"))

    if amount < min_limit:
        await update.message.reply_text(tr(user_id, "wd_min_limit_err", min_limit=min_limit), reply_markup=get_cancel_keyboard(user_id))
        return STATE_WD_AMOUNT

    if amount > max_limit:
        await update.message.reply_text(tr(user_id, "wd_max_limit_err", max_limit=max_limit), reply_markup=get_cancel_keyboard(user_id))
        return STATE_WD_AMOUNT

    number = context.user_data.get("wd_number")
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    fee_usdt = 0.0
    usdt_rate = 0.0
    usdt_amount = None
    debit_amount_bdt = amount

    if method in ("usdt", "binance_uid"):
        usdt_rate = max(float(get_setting_val("usdt_rate_bdt", "120.0")), 0.000001)
        usdt_amount = amount
        charge_enabled = get_setting_val("usdt_charge_enabled", "ON") == "ON"
        fee_usdt = float(get_setting_val("usdt_charge_usdt", "0.05")) if charge_enabled else 0.0
        debit_amount_bdt = (usdt_amount + fee_usdt) * usdt_rate
        if debit_amount_bdt > float(user_data["balance"]):
            await update.message.reply_text(
                f"❌ ব্যালেন্স যথেষ্ট নয়। প্রয়োজন: ৳{debit_amount_bdt:.2f} (প্রায় ${debit_amount_bdt/usdt_rate:.4f} USDT)",
                reply_markup=get_cancel_keyboard(user_id)
            )
            return STATE_WD_AMOUNT
    elif amount > user_data["balance"]:
        await update.message.reply_text(tr(user_id, "wd_insufficient_bal"), reply_markup=get_cancel_keyboard(user_id))
        return STATE_WD_AMOUNT

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("UPDATE users SET balance = balance - ? WHERE user_id = ?", (debit_amount_bdt, user_id))
        cursor.execute(
            """INSERT INTO withdrawals
               (user_id, method, number, amount, status, created_at, usdt_amount, usdt_fee, usdt_rate, currency)
               VALUES (?, ?, ?, ?, 'Pending', ?, ?, ?, ?, ?)""",
            (
                user_id, method, number, debit_amount_bdt, now,
                usdt_amount, fee_usdt, usdt_rate,
                "USDT" if method in ("usdt", "binance_uid") else "BDT"
            )
        )
        wd_id = cursor.lastrowid
        conn.commit()

    if method in ("usdt", "binance_uid"):
        fee_text = f" (Charge: ${fee_usdt:.4f} USDT)" if fee_usdt > 0 else " (Charge: OFF)"
        shown_amount = usdt_amount
        shown_currency = "$"
    else:
        chg_bdt = float(get_setting_val(f"{method}_charge_bdt", "0.0"))
        if chg_bdt > 0:
            if amount <= chg_bdt:
                await update.message.reply_text(tr(user_id, "wd_usdt_fee_err"), reply_markup=get_cancel_keyboard(user_id))
                return STATE_WD_AMOUNT
            net_bdt = amount - chg_bdt
            fee_text = f" (Charge: ৳{chg_bdt:g} | Net: ৳{net_bdt:g})"
        else:
            fee_text = ""
        shown_amount = amount
        shown_currency = "৳"

    await update.message.reply_text(
        tr(user_id, "wd_success_msg", wd_id=wd_id, amount=shown_amount, fee_text=fee_text, method=("USDT (BEP20)" if method == "usdt" else "Binance UID" if method == "binance_uid" else method.upper()), number=number).replace("৳", shown_currency),
        reply_markup=get_main_keyboard(user_id),
        parse_mode=ParseMode.MARKDOWN
    )

    admin_id_setting = ADMIN_ID or get_setting_val("admin_id", "")
    admins = [int(x.strip()) for x in admin_id_setting.split(",") if x.strip().isdigit()]

    admin_msg = (
        f"💸 **নতুন উইথড্র রিকোয়েস্ট! / New Withdraw Request!**\n\n"
        f"🆔 **উইথড্র আইডি:** #{wd_id}\n"
        f"👤 **ইউজার আইডি:** <code>{user_id}</code>\n"
        f"💵 **পরিমাণ:** " + (f"${usdt_amount:.4f} USDT (Debit: ৳{debit_amount_bdt:.2f})" if method in ("usdt", "binance_uid") else f"৳{amount:.2f}{fee_text}") + "\n"
        f"📱 **গেটওয়ে:** {('USDT (BEP20)' if method == 'usdt' else 'Binance UID' if method == 'binance_uid' else method.upper())} - <code>{number}</code>"
    )
    buttons = [
        [
            ikb_btn("✅ Approve", style="success", callback_data=f"adm_app_wd_{wd_id}"),
            ikb_btn("❌ Reject", style="danger", callback_data=f"adm_rej_wd_{wd_id}")
        ]
    ]

    for admin in admins:
        try:
            await context.bot.send_message(
                chat_id=admin,
                text=admin_msg,
                reply_markup=InlineKeyboardMarkup(buttons),
                parse_mode=ParseMode.HTML
            )
        except Exception as e:
            logger.error(f"Failed to alert admin of withdraw: {e}")

    return ConversationHandler.END

# -----------------------------------------------------------------------------
# REFERRAL, LEADERBOARD, SUPPORT, LANGUAGE
# -----------------------------------------------------------------------------

async def handle_referrals(update: Update, context: ContextTypes.DEFAULT_TYPE):
    clear_user_state(context)
    user_id = update.effective_user.id
    is_joined, _ = await check_force_join(user_id, context)
    if not is_joined:
        await start_command(update, context)
        return ConversationHandler.END

    user_data = get_user_data(user_id)
    if not user_data:
        register_user(update.effective_user)
        user_data = get_user_data(user_id)

    bot_username = (await context.bot.get_me()).username
    ref_link = f"https://t.me/{bot_username}?start={user_id}"
    ref_percent = get_setting_val("ref_bonus_percent", "10.0")

    custom_fixed = user_data["custom_ref_fixed"] if "custom_ref_fixed" in user_data.keys() else None
    if custom_fixed is not None:
        ref_text = (
            f"👥 <b>আপনার ব্যক্তিগত রেফারেল প্রোগ্রাম</b> 💎\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"আপনার রেফারকৃত কর্মী প্রতিটি কাজ সম্পন্ন ও এপ্রুভ করালে আপনি পাবেন সরাসরি <b>৳{float(custom_fixed):g}</b> করে বোনাস (রিজেক্ট হলে কোনো বোনাস পাবেন না)।\n\n"
            f"🔗 <b>আপনার রেফারেল লিঙ্ক:</b>\n<code>{ref_link}</code>\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"📊 <b>রেফারেল ট্র্যাকার:</b>\n"
            f"• মোট আমন্ত্রিত কর্মী: <code>{user_data['ref_count']}</code> জন\n"
            f"• রেফারেল বোনাস থেকে আয়: <code>৳{user_data['ref_earnings']:.2f}</code>"
        )
    else:
        ref_text = tr(
            user_id, "ref_text",
            ref_link=ref_link,
            ref_count=user_data['ref_count'],
            ref_earnings=user_data['ref_earnings'],
            ref_percent=ref_percent
        )
    from urllib.parse import quote as _url_quote
    share_text = "💎 Join & earn money by submitting Gmail accounts! Start now:"
    share_url = f"https://t.me/share/url?url={_url_quote(ref_link, safe='')}&text={_url_quote(share_text, safe='')}"
    share_kb = InlineKeyboardMarkup([
        [ikb_btn("Share Referral Link", style="success", url=share_url)]
    ])
    await update.message.reply_text(ref_text, reply_markup=share_kb, parse_mode=ParseMode.HTML)

async def handle_leaderboard(update: Update, context: ContextTypes.DEFAULT_TYPE):
    clear_user_state(context)
    user_id = update.effective_user.id
    is_joined, _ = await check_force_join(user_id, context)
    if not is_joined:
        await start_command(update, context)
        return ConversationHandler.END

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT full_name, balance FROM users WHERE is_banned = 0 ORDER BY balance DESC LIMIT 10")
        top_earners = cursor.fetchall()

    leaderboard = tr(user_id, "leaderboard_title")
    if not top_earners:
        leaderboard += tr(user_id, "no_leaderboard_data")
    else:
        for idx, row in enumerate(top_earners, 1):
            name = html.escape(row["full_name"])
            badge = "🥇" if idx == 1 else "🥈" if idx == 2 else "🥉" if idx == 3 else f"{idx}."
            leaderboard += f"{badge} **{name}** — ৳{row['balance']:.2f}\n"

    await update.message.reply_text(leaderboard, reply_markup=get_main_keyboard(user_id), parse_mode=ParseMode.MARKDOWN)

async def handle_support(update: Update, context: ContextTypes.DEFAULT_TYPE):
    clear_user_state(context)
    user_id = update.effective_user.id
    is_joined, _ = await check_force_join(user_id, context)
    if not is_joined:
        await start_command(update, context)
        return ConversationHandler.END

    support_handle = get_setting_val("support_handle", "https://t.me/TrustVaultMails_Owners")
    official_channel = get_setting_val("official_channel_link", "https://t.me/TrustVaultMailsOfficial")

    support_text = tr(user_id, "support_text")

    buttons = [
        [ikb_btn(tr(user_id, "btn_admin_support"), style="success", url=support_handle)],
        [ikb_btn(tr(user_id, "btn_official_channel"), url=official_channel)]
    ]

    await update.message.reply_text(
        support_text,
        reply_markup=InlineKeyboardMarkup(buttons),
        parse_mode=ParseMode.MARKDOWN
    )

async def handle_language(update: Update, context: ContextTypes.DEFAULT_TYPE):
    clear_user_state(context)
    user_id = update.effective_user.id
    is_joined, _ = await check_force_join(user_id, context)
    if not is_joined:
        await start_command(update, context)
        return ConversationHandler.END

    buttons = [
        [ikb_btn("🇧🇩 বাংলা", style="success", callback_data="lang_bn"),
         ikb_btn("🇺🇸 English", style="primary", callback_data="lang_en")]
    ]
    await update.message.reply_text(
        tr(user_id, "lang_select_prompt"),
        reply_markup=InlineKeyboardMarkup(buttons),
        parse_mode=ParseMode.MARKDOWN
    )

async def language_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    user_id = query.from_user.id

    lang = "bn" if data == "lang_bn" else "en"
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("UPDATE users SET language = ? WHERE user_id = ?", (lang, user_id))
        conn.commit()

    msg = tr(lang, "lang_set_bn") if lang == "bn" else tr(lang, "lang_set_en")
    await query.message.reply_text(msg, reply_markup=get_main_keyboard(user_id, lang=lang), parse_mode=ParseMode.MARKDOWN)

# -----------------------------------------------------------------------------
# ADMIN PANEL HANDLERS
# -----------------------------------------------------------------------------

async def handle_admin_panel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    clear_user_state(context)
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM users")
        total_users = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM task_submissions WHERE status = 'Pending'")
        pending_tasks = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM task_submissions WHERE status = 'Hold'")
        held_tasks = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM withdrawals WHERE status = 'Pending'")
        pending_wd = cursor.fetchone()[0]

    dash_msg = tr(
        user_id, "admin_dash_title",
        bot_name=get_setting_val("bot_display_name", "Gmail Work BD"),
        total_users=total_users,
        pending_tasks=pending_tasks,
        held_tasks=held_tasks,
        pending_wd=pending_wd
    )

    await update.message.reply_text(
        dash_msg,
        reply_markup=get_admin_keyboard(user_id),
        parse_mode=ParseMode.MARKDOWN
    )

async def admin_pending_tasks(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM task_submissions WHERE status = 'Pending' ORDER BY created_at DESC LIMIT 10")
        tasks = cursor.fetchall()
    if not tasks:
        await update.message.reply_text(tr(user_id, "admin_no_pending_tasks"), reply_markup=get_admin_keyboard(user_id))
        return
    msg = tr(user_id, "admin_pending_tasks_title")
    keyboard = []
    for task in tasks:
        msg += (
            f"🆔 #{task['submission_id']} | {task['task_type'].upper()}\n"
            f"👤 {html.escape(task['submitted_username'])} | ৳{task['reward_amount']:.2f}\n"
            f"📅 {task['created_at']}\n\n"
        )
        keyboard.append([
            ikb_btn(f"✅ App #{task['submission_id']}", callback_data=f"adm_app_sub_{task['submission_id']}"),
            ikb_btn(f"⏳ Hold", callback_data=f"adm_hold_sub_{task['submission_id']}"),
            ikb_btn(f"❌ Rej", callback_data=f"adm_rej_prompt_{task['submission_id']}")
        ])
    keyboard.append([ikb_btn(tr(user_id, "btn_admin_menu"), callback_data="adm_back_to_admin")])
    await update.message.reply_text(msg, reply_markup=InlineKeyboardMarkup(keyboard), parse_mode=ParseMode.HTML)

async def admin_held_tasks(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM task_submissions WHERE status = 'Hold' ORDER BY updated_at DESC LIMIT 15")
        tasks = cursor.fetchall()
    if not tasks:
        await update.message.reply_text(tr(user_id, "admin_no_held_tasks"), reply_markup=get_admin_keyboard(user_id))
        return
    msg = tr(user_id, "admin_held_tasks_title")
    keyboard = []
    for task in tasks:
        msg += (
            f"🆔 <b>Task #{task['submission_id']}</b> | 📱 {task['task_type'].upper()}\n"
            f"👤 <b>User ID:</b> <code>{task['user_id']}</code>\n"
            f"👤 <b>Submission:</b> <code>{html.escape(task['submitted_username'])}</code>\n"
        )
        if task['submitted_password'] and task['submitted_password'] != 'N/A':
            msg += f"🔑 <b>Pass:</b> <code>{html.escape(task['submitted_password'])}</code>\n"
        if task['secret_key_2fa'] and task['secret_key_2fa'] != 'N/A':
            msg += f"🔐 <b>2FA Key:</b> <code>{html.escape(task['secret_key_2fa'])}</code>\n"
        if task['cookies_data']:
            msg += f"🍪 <b>Data:</b> <code>{html.escape(task['cookies_data'][:40])}...</code>\n"
        msg += (
            f"💵 <b>Price:</b> ৳{task['reward_amount']:.2f}\n"
            f"📅 <b>Held At:</b> {task['updated_at']}\n━━━━━━━━━━━━━━━━━━\n"
        )
        keyboard.append([
            ikb_btn(f"✅ App #{task['submission_id']}", callback_data=f"adm_app_sub_{task['submission_id']}"),
            ikb_btn(f"❌ Rej #{task['submission_id']}", callback_data=f"adm_rej_prompt_{task['submission_id']}")
        ])
    keyboard.append([ikb_btn(tr(user_id, "btn_admin_menu"), callback_data="adm_back_to_admin")])
    await update.message.reply_text(msg, reply_markup=InlineKeyboardMarkup(keyboard), parse_mode=ParseMode.HTML)

# --- BULK APPROVE & BULK REJECT ---
async def admin_bulk_approve_prompt(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return
    await update.message.reply_text(
        tr(user_id, "admin_bulk_app_prompt"),
        reply_markup=get_cancel_keyboard(user_id),
        parse_mode=ParseMode.MARKDOWN
    )
    return STATE_ADMIN_BULK_APPROVE

async def admin_bulk_reject_prompt(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return
    await update.message.reply_text(
        tr(user_id, "admin_bulk_rej_prompt"),
        reply_markup=get_cancel_keyboard(user_id),
        parse_mode=ParseMode.MARKDOWN
    )
    return STATE_ADMIN_BULK_REJECT

async def process_bulk_approve(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return ConversationHandler.END

    raw_text = update.message.text.strip()
    if is_menu_or_admin_button(raw_text):
        clear_user_state(context)
        return ConversationHandler.END

    items = [x.strip() for x in re.split(r'[\n,\s]+', raw_text) if x.strip()]
    cleaned_items = list(dict.fromkeys(items))

    if not cleaned_items:
        await update.message.reply_text("❌ কোনো ভ্যালিড ডাটা পাওয়া যায়নি।", reply_markup=get_admin_keyboard(user_id))
        return ConversationHandler.END

    approved_count = 0
    matched_count = 0
    unmatched_list = []
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for item in cleaned_items:
        with get_db() as conn:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT * FROM task_submissions WHERE LOWER(submitted_username) = LOWER(?) AND status IN ('Pending', 'Hold')",
                (item,)
            )
            subs = cursor.fetchall()
            
            if not subs:
                unmatched_list.append(item)
                continue

            matched_count += len(subs)

            for sub in subs:
                sub_id = sub["submission_id"]
                sub_user_id = sub["user_id"]
                reward = sub["reward_amount"]

                cursor.execute("UPDATE users SET balance = balance + ? WHERE user_id = ?", (reward, sub_user_id))
                cursor.execute("UPDATE task_submissions SET status = 'Approved', updated_at = ? WHERE submission_id = ?", (now, sub_id))

                # Lifetime Referral Bonus logic
                cursor.execute("SELECT referred_by FROM users WHERE user_id = ?", (sub_user_id,))
                user_ref = cursor.fetchone()
                if user_ref and user_ref["referred_by"]:
                    referrer_id = user_ref["referred_by"]
                    commission_bonus = get_referrer_commission(referrer_id, float(reward))
                    if commission_bonus > 0:
                        cursor.execute(
                            "UPDATE users SET balance = balance + ?, ref_earnings = ref_earnings + ? WHERE user_id = ?",
                            (commission_bonus, commission_bonus, referrer_id)
                        )
                        try:
                            await context.bot.send_message(
                                chat_id=referrer_id,
                                text=tr(referrer_id, "user_ref_commission", bonus=commission_bonus),
                                parse_mode=ParseMode.HTML
                            )
                        except Exception as e:
                            logger.error(f"Failed to notify referrer {referrer_id}: {e}")

                conn.commit()
                approved_count += 1

                sub_username = html.escape(str(sub['submitted_username'] or '').strip())
                try:
                    await context.bot.send_message(
                        chat_id=sub_user_id,
                        text=tr(sub_user_id, "user_task_approved", task_display=sub["task_type"].upper(), username=sub_username, reward=reward),
                        parse_mode=ParseMode.HTML
                    )
                except Exception as e:
                    logger.error(f"Failed to notify user {sub_user_id}: {e}")

    report_msg = (
        f"✅ **বাল্ক এপ্রুভ সম্পন্ন রিপোর্ট:**\n━━━━━━━━━━━━━━━━━━━━━━\n"
        f"• ইনপুট দেওয়া মোট আইটেম: `{len(cleaned_items)}` টি\n"
        f"• ডাটাবেজে ম্যাচ পাওয়া টাস্ক: `{matched_count}` টি\n"
        f"• সফলভাবে এপ্রুভ হয়েছে: `{approved_count}` টি\n"
    )
    if unmatched_list:
        report_msg += f"• **পেন্ডিং/হোল্ডে খুঁজে পাওয়া যায়নি ({len(unmatched_list)} টি):**\n`" + ", ".join(unmatched_list[:20]) + ("..." if len(unmatched_list)>20 else "") + "`"

    await update.message.reply_text(
        report_msg,
        reply_markup=get_admin_keyboard(user_id),
        parse_mode=ParseMode.MARKDOWN
    )
    clear_user_state(context)
    return ConversationHandler.END

async def process_bulk_reject(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return ConversationHandler.END

    raw_text = update.message.text.strip()
    if is_menu_or_admin_button(raw_text):
        clear_user_state(context)
        return ConversationHandler.END

    items = [x.strip() for x in re.split(r'[\n,\s]+', raw_text) if x.strip()]
    cleaned_items = list(dict.fromkeys(items))

    if not cleaned_items:
        await update.message.reply_text("❌ কোনো ভ্যালিড ডাটা পাওয়া যায়নি।", reply_markup=get_admin_keyboard(user_id))
        return ConversationHandler.END

    rejected_count = 0
    matched_count = 0
    unmatched_list = []
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    reason = "আপনার Gmail অ্যাকাউন্ট লগইন করা সম্ভব হয়নি। তাই অ্যাকাউন্টটি রিজেক্ট করা হয়েছে।"

    for item in cleaned_items:
        with get_db() as conn:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT * FROM task_submissions WHERE LOWER(submitted_username) = LOWER(?) AND status IN ('Pending', 'Hold')",
                (item,)
            )
            subs = cursor.fetchall()

            if not subs:
                unmatched_list.append(item)
                continue

            matched_count += len(subs)

            for sub in subs:
                sub_id = sub["submission_id"]
                sub_user_id = sub["user_id"]

                cursor.execute(
                    "UPDATE task_submissions SET status = 'Rejected', reject_reason = ?, updated_at = ? WHERE submission_id = ?",
                    (reason, now, sub_id)
                )
                conn.commit()
                rejected_count += 1

                sub_username = html.escape(str(sub['submitted_username'] or '').strip())
                esc_reason = html.escape(reason)
                try:
                    await context.bot.send_message(
                        chat_id=sub_user_id,
                        text=tr(sub_user_id, "user_task_rejected", task_display=sub["task_type"].upper(), username=sub_username, reason=esc_reason),
                        parse_mode=ParseMode.HTML
                    )
                except Exception as e:
                    logger.error(f"Failed to notify user {sub_user_id}: {e}")

    report_msg = (
        f"❌ **বাল্ক রিজেক্ট সম্পন্ন রিপোর্ট:**\n━━━━━━━━━━━━━━━━━━━━━━\n"
        f"• ইনপুট দেওয়া মোট আইটেম: `{len(cleaned_items)}` টি\n"
        f"• ডাটাবেজে ম্যাচ পাওয়া টাস্ক: `{matched_count}` টি\n"
        f"• সফলভাবে রিজেক্ট হয়েছে: `{rejected_count}` টি\n"
    )
    if unmatched_list:
        report_msg += f"• **পেন্ডিং/হোল্ডে খুঁজে পাওয়া যায়নি ({len(unmatched_list)} টি):**\n`" + ", ".join(unmatched_list[:20]) + ("..." if len(unmatched_list)>20 else "") + "`"

    await update.message.reply_text(
        report_msg,
        reply_markup=get_admin_keyboard(user_id),
        parse_mode=ParseMode.MARKDOWN
    )
    clear_user_state(context)
    return ConversationHandler.END

# --- SPREADSHEET UI EXPORT ---
async def admin_excel_export_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return

    buttons = [
        [ikb_btn("Gmail (Download)", style="success", callback_data="adm_xlsx_gmail"),
         ikb_btn("Clear Gmail", style="danger", callback_data="adm_cls_gmail")],
        [ikb_btn("Old Gmail (Download)", style="success", callback_data="adm_xlsx_old_gmail"),
         ikb_btn("Clear Old Gmail", style="danger", callback_data="adm_cls_old_gmail")],
        [ikb_btn("All Tasks (Download)", style="primary", callback_data="adm_xlsx_all"),
         ikb_btn("Clear ALL Data", style="danger", callback_data="adm_cls_all")],
        [ikb_btn(tr(user_id, "btn_admin_menu"), callback_data="adm_back_to_admin")]
    ]

    await update.message.reply_text(
        "📊 **এক্সেল শীট (.xlsx) ডাউনলোড এবং ক্লিয়ার কন্ট্রোল প্যানেল:**\n━━━━━━━━━━━━━━━━━━━━━━\n"
        "• **Download:** রিজেক্ট হওয়া কাজ বাদ দিয়ে শুধুমাত্র ভ্যালিড কাজের ফাইল ডাউনলোড করবে।\n"
        "• **Clear:** ডাটা আর্কাইভ করবে (রিসেট করবে) যাতে নতুন দিনের কাজ ফ্রেশভাবে ১ থেকে শুরু হয়।",
        reply_markup=InlineKeyboardMarkup(buttons),
        parse_mode=ParseMode.MARKDOWN
    )

async def admin_excel_export_download_cb(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    user_id = query.from_user.id
    await query.answer()

    if not is_admin(user_id):
        return

    data = query.data.replace("adm_xlsx_", "")

    with get_db() as conn:
        cursor = conn.cursor()
        if data == "all":
            cursor.execute("SELECT * FROM task_submissions WHERE status != 'Rejected' AND is_archived = 0 ORDER BY submission_id DESC")
            task_label = "সকল টাস্ক (All Tasks)"
        else:
            cursor.execute("SELECT * FROM task_submissions WHERE task_type = ? AND status != 'Rejected' AND is_archived = 0 ORDER BY submission_id DESC", (data,))
            task_label = {
                "ig_2fa": "Instagram 2FA",
                "ig_seed": "Instagram Seed",
                "fb_2fa": "Facebook 2FA",
                "fb_cookies": "Facebook Cookies",
                "gmail": "Gmail",
            "old_gmail": "Old Gmail",
                "old_gmail": "Old Gmail"
            }.get(data, data.upper())

        tasks = cursor.fetchall()

        if data == "all":
            cursor.execute("SELECT COUNT(*), status FROM task_submissions WHERE status != 'Rejected' AND is_archived = 0 GROUP BY status")
        else:
            cursor.execute("SELECT COUNT(*), status FROM task_submissions WHERE task_type = ? AND status != 'Rejected' AND is_archived = 0 GROUP BY status", (data,))
        status_counts = dict(cursor.fetchall())

        if data == "all":
            cursor.execute("SELECT SUM(reward_amount) FROM task_submissions WHERE status != 'Rejected' AND is_archived = 0")
        else:
            cursor.execute("SELECT SUM(reward_amount) FROM task_submissions WHERE task_type = ? AND status != 'Rejected' AND is_archived = 0", (data,))
        total_amount = cursor.fetchone()[0] or 0.0

    pending = status_counts.get("Pending", 0)
    approved = status_counts.get("Approved", 0)
    hold = status_counts.get("Hold", 0)

    clear_btn_kb = InlineKeyboardMarkup([
        [ikb_btn(f"🗑️ {task_label} ডাটা ক্লিয়ার করুন (0)", callback_data=f"adm_cls_{data}")],
        [ikb_btn(tr(user_id, "btn_admin_menu"), callback_data="adm_back_to_admin")]
    ])

    if HAS_OPENPYXL:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = task_label[:30]
        ws.views.sheetView[0].showGridLines = True

        ws.merge_cells('A1:J1')
        summary_cell = ws['A1']
        summary_cell.value = f"📊 {task_label.upper()} DASHBOARD | Pending: {pending} | Approved: {approved} | Hold: {hold} | Total Amount: ৳{total_amount:.2f}"
        summary_cell.font = Font(bold=True, size=11, color="FFFFFF")
        summary_cell.fill = PatternFill("solid", fgColor="1F4E78")
        summary_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 25

        headers = ["Task ID", "User ID", "Username", "Task Type", "Password", "2FA Secret Key", "Cookies / Data", "Price (BDT)", "Status", "Submitted Time"]
        ws.append([])
        ws.append(headers)
        ws.row_dimensions[3].height = 22

        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill("solid", fgColor="2F5597")
        header_align = Alignment(horizontal="center", vertical="center")

        thin_border = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )

        for col_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col_idx)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align

        ws.freeze_panes = 'A4'

        fill_approved = PatternFill("solid", fgColor="E2EFDA")
        fill_pending = PatternFill("solid", fgColor="FFF2CC")
        fill_hold = PatternFill("solid", fgColor="FCE4D6")

        font_approved = Font(color="375623", bold=True)
        font_pending = Font(color="7F6000", bold=True)
        font_hold = Font(color="C65911", bold=True)

        for task in tasks:
            status = task["status"]

            row_data = [
                task["submission_id"],
                task["user_id"],
                task["submitted_username"],
                task["task_type"].upper(),
                task["submitted_password"] or "N/A",
                task["secret_key_2fa"] or "N/A",
                task["cookies_data"] or "N/A",
                task["reward_amount"],
                status,
                task["created_at"]
            ]
            ws.append(row_data)
            current_row = ws.max_row
            ws.row_dimensions[current_row].height = 19

            for col_idx in range(1, 11):
                cell = ws.cell(row=current_row, column=col_idx)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center")

            status_cell = ws.cell(row=current_row, column=9)
            if status == "Approved":
                status_cell.fill = fill_approved
                status_cell.font = font_approved
            elif status == "Pending":
                status_cell.fill = fill_pending
                status_cell.font = font_pending
            elif status == "Hold":
                status_cell.fill = fill_hold
                status_cell.font = font_hold

        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 40)

        bio = io.BytesIO()
        wb.save(bio)
        bio.seek(0)
        bio.name = f"{data}_valid_tasks_{datetime.date.today()}.xlsx"

        await context.bot.send_document(
            chat_id=query.message.chat_id,
            document=bio,
            caption=f"📊 **{task_label} Excel Report (Valid Tasks Only)**\n━━━━━━━━━━━━━━━━━━━━━━\n• **Total Valid Tasks:** {len(tasks)}\n• **Pending:** {pending}\n• **Approved:** {approved}\n• **Hold:** {hold}\n• **Total Amount:** ৳{total_amount:.2f}\n\n*(রিজেক্ট হওয়া কাজগুলো রিপোর্ট থেকে বাদ দেওয়া হয়েছে)*",
            reply_markup=clear_btn_kb,
            parse_mode=ParseMode.MARKDOWN
        )
    else:
        buffer = io.StringIO()
        buffer.write("Task ID,User ID,Username,Task Type,Password,2FA Secret,Cookies,Price,Status,Submitted Time\n")
        for t in tasks:
            buffer.write(f"{t['submission_id']},{t['user_id']},{t['submitted_username']},{t['task_type']},{t['submitted_password']},{t['secret_key_2fa']},{t['cookies_data']},{t['reward_amount']},{t['status']},{t['created_at']}\n")
        
        file_bytes = buffer.getvalue().encode('utf-8')
        bio = io.BytesIO(file_bytes)
        bio.name = f"{data}_tasks.csv"
        await context.bot.send_document(
            chat_id=query.message.chat_id,
            document=bio,
            caption=f"📊 **{task_label} CSV Report**",
            reply_markup=clear_btn_kb
        )

# --- SOFT DELETE / ARCHIVING CLEAR CALLBACK ---
async def admin_excel_clear_cb(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    user_id = query.from_user.id
    await query.answer()

    if not is_admin(user_id):
        return

    data = query.data

    if data.startswith("adm_cls_conf_"):
        cat = data.replace("adm_cls_conf_", "")
        with get_db() as conn:
            cursor = conn.cursor()
            if cat == "all":
                cursor.execute("UPDATE task_submissions SET is_archived = 1 WHERE is_archived = 0")
                label = "সকল কাজের"
            else:
                cursor.execute("UPDATE task_submissions SET is_archived = 1 WHERE task_type = ? AND is_archived = 0", (cat,))
                label = {
                    "ig_2fa": "Instagram 2FA",
                    "ig_seed": "Instagram Seed",
                    "fb_2fa": "Facebook 2FA",
                    "fb_cookies": "Facebook Cookies",
                    "gmail": "Gmail",
            "old_gmail": "Old Gmail",
                "old_gmail": "Old Gmail"
                }.get(cat, cat.upper())
            conn.commit()

        await query.edit_message_text(
            f"✅ **{label} সমস্ত ডাটা সফলভাবে ক্লিয়ার (আর্কাইভ) করা হয়েছে!**\n\n"
            f"পূর্বে ডাউনলোডকৃত ডাটা নিরাপদ রয়েছে এবং নতুন জমার কাজ ১ থেকে গণনাকৃত হবে।",
            reply_markup=InlineKeyboardMarkup([[ikb_btn(tr(user_id, "btn_admin_menu"), callback_data="adm_back_to_admin")]]),
            parse_mode=ParseMode.MARKDOWN
        )
        return

    if data.startswith("adm_cls_"):
        cat = data.replace("adm_cls_", "")
        label = {
            "ig_2fa": "Instagram 2FA",
            "ig_seed": "Instagram Seed",
            "fb_2fa": "Facebook 2FA",
            "fb_cookies": "Facebook Cookies",
            "gmail": "Gmail",
            "old_gmail": "Old Gmail",
                "old_gmail": "Old Gmail",
            "all": "সকল কাজের (ALL)"
        }.get(cat, cat.upper())

        buttons = [
            [ikb_btn("✅ হ্যাঁ, ক্লিয়ার করুন (Clear 0)", callback_data=f"adm_cls_conf_{cat}")],
            [ikb_btn("❌ বাতিল", callback_data="adm_back_to_admin")]
        ]

        await query.edit_message_text(
            f"⚠️ **আপনি কি নিশ্চিত যে {label} এর জমার ডাটা ক্লিয়ার (আর্কাইভ) করতে চান?**\n\n"
            f"ডাটা ক্লিয়ার করলে ডাটাবেজে রেকর্ড সেইফ রেখে এক্সেল ফাইল রিসেট হয়ে যাবে।",
            reply_markup=InlineKeyboardMarkup(buttons),
            parse_mode=ParseMode.MARKDOWN
        )

# --- USER MANAGEMENT & ADMIN CONTROL PANELS ---
async def admin_user_management(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return

    buttons = [
        [ikb_btn("Ban User", callback_data="adm_usr_ban"), ikb_btn("Unban User", callback_data="adm_usr_unban")],
        [ikb_btn("Add Balance", callback_data="adm_usr_add_bal"), ikb_btn("Deduct Balance", callback_data="adm_usr_sub_bal")],
        [ikb_btn("Reset All Balances", callback_data="adm_usr_reset_balances")],
        [ikb_btn("Custom Ref Bonus (৳)", style="primary", callback_data="adm_usr_set_ref"), ikb_btn("Banned List", callback_data="adm_usr_banned_list")],
        [ikb_btn("Search User Profile", style="success", callback_data="adm_usr_search")],
        [ikb_btn("Admin Menu", callback_data="adm_back_to_admin")]
    ]

    await update.message.reply_text(
        "👥 **ইউজার ম্যানেজমেন্ট প্যানেল (User Management):**\n━━━━━━━━━━━━━━━━━━━━━━\nনিচের অপশন থেকে নির্বাচন করুন:",
        reply_markup=InlineKeyboardMarkup(buttons),
        parse_mode=ParseMode.MARKDOWN
    )

async def admin_control_panel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return

    force_join = get_setting_val("force_join_enabled", "OFF")

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT user_id FROM sub_admins")
        sub_admins = cursor.fetchall()

    sub_list_str = "\n".join([f"• <code>{row['user_id']}</code>" for row in sub_admins]) if sub_admins else "কোনো সাব-এডমিন নেই।"

    buttons = [
        [ikb_btn(f"🏷️ Bot Name: {bot_name}\n🎧 Support: {support_link}\n📢 Channel: {channel_link}\n🔒 Force Join: {force_join}", callback_data="adm_toggle_force_join_enabled"),
         ikb_btn("📢 Edit Force Channels List", callback_data="adm_edit_force_channels_list")],
        [ikb_btn("➕ Add Sub-Admin", callback_data="adm_add_subadmin"), ikb_btn("➖ Remove Sub-Admin", callback_data="adm_rem_subadmin")],
        [ikb_btn("Admin Menu", callback_data="adm_back_to_admin")]
    ]

    await update.message.reply_text(
        f"👑 **এডমিন কন্ট্রোল প্যানেল (Admin Control):**\n━━━━━━━━━━━━━━━━━━━━━━\n"
        f"👑 **Main Owner ID:** <code>{get_setting_val('admin_id')}</code>\n\n"
        f"👥 **Sub-Admins:**\n{sub_list_str}",
        reply_markup=InlineKeyboardMarkup(buttons),
        parse_mode=ParseMode.HTML
    )


async def admin_old_categories_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return
    cats = get_old_categories()
    buttons = []
    for c in cats:
        buttons.append([ikb_btn(f"❌ {c['label']} (৳{float(c['rate']):g})", style="danger", callback_data=f"adm_oldcat_del_{c['cat_key']}")])
    buttons.append([ikb_btn("➕ Add Old Category", style="success", callback_data="adm_oldcat_add")])
    buttons.append([ikb_btn("Admin Menu", style="primary", callback_data="adm_back_to_admin")])

    lines = "\n".join(f"• <b>{html.escape(c['label'])}</b> — ৳{float(c['rate']):g}" for c in cats) if cats else "এখনো কোনো ক্যাটাগরি যোগ করা হয়নি।"
    await update.message.reply_text(
        f"♻️ <b>Old Gmail Categories ({len(cats)} টি)</b>\n━━━━━━━━━━━━━━━━━━━━━━\n{lines}\n\n"
        f"মুছতে চাইলে নিচের ক্যাটাগরিতে চাপ দিন অথবা নতুন যোগ করুন।",
        reply_markup=InlineKeyboardMarkup(buttons),
        parse_mode=ParseMode.HTML
    )


async def process_admin_old_cat_add(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return ConversationHandler.END
    raw = (update.message.text or "").strip()
    if is_menu_or_admin_button(raw):
        clear_user_state(context)
        return ConversationHandler.END
    if "|" not in raw:
        await update.message.reply_text("❌ সঠিক ফরম্যাটে দিন: <code>নাম | রেট</code> (যেমন: <code>2015-2018 Old | 80</code>)", parse_mode=ParseMode.HTML, reply_markup=get_cancel_keyboard(user_id))
        return STATE_ADMIN_OLD_CAT_ADD
    label, rate_str = [x.strip() for x in raw.split("|", 1)]
    try:
        rate = float(rate_str)
        if not label or rate <= 0:
            raise ValueError
    except ValueError:
        await update.message.reply_text("❌ সঠিক নাম এবং পজিটিভ রেট দিন।", reply_markup=get_cancel_keyboard(user_id))
        return STATE_ADMIN_OLD_CAT_ADD

    add_old_category(label, rate)
    await update.message.reply_text(
        f"✅ নতুন Old Gmail ক্যাটাগরি যোগ হয়েছে:\n<b>{html.escape(label)}</b> — ৳{rate:g}",
        reply_markup=get_admin_keyboard(user_id),
        parse_mode=ParseMode.HTML
    )
    clear_user_state(context)
    return ConversationHandler.END

async def admin_search_prompt(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return
    context.user_data["admin_search"] = True
    await update.message.reply_text(tr(user_id, "admin_search_prompt_msg"), reply_markup=get_cancel_keyboard(user_id))
    return STATE_ADMIN_SEARCH_QUERY

async def admin_export_prompt(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return
    context.user_data["admin_export"] = True
    await update.message.reply_text(tr(user_id, "admin_export_prompt_msg"), reply_markup=get_cancel_keyboard(user_id))
    return STATE_ADMIN_EXPORT_USER

async def admin_password_manager(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return
    gmail_pass = get_setting_val("gmail_default_password", "aass1122")
    old_pass = get_setting_val("old_gmail_default_password", "oldpass123")
    old_rec = get_setting_val("old_gmail_recovery_email", "") or "(Not Set)"
    msg = tr(user_id, "admin_pwd_mgr_title", gmail_pass=gmail_pass, old_pass=old_pass, old_rec=old_rec)
    keyboard = [
        [ikb_btn("New/Bulk Gmail Pass", style="primary", callback_data="adm_edit_gmail_default_password"),
         ikb_btn("Old Gmail Pass", style="success", callback_data="adm_edit_old_gmail_default_password")],
        [ikb_btn("Old Gmail Recovery", style="danger", callback_data="adm_edit_old_gmail_recovery_email")],
        [ikb_btn(tr(user_id, "btn_admin_menu"), style="primary", callback_data="adm_back_to_admin")]
    ]
    await update.message.reply_text(msg, reply_markup=InlineKeyboardMarkup(keyboard), parse_mode=ParseMode.HTML)

async def admin_price_manager(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return
    gmail = get_setting_val("gmail_task_price", "22.00")
    ref_bonus = get_setting_val("ref_bonus_percent", "10.0")

    msg = tr(
        user_id, "admin_price_mgr_title",
        gmail=gmail, ref_bonus=ref_bonus
    )
    keyboard = [
        [ikb_btn(f"✉️ Gmail (৳{gmail})", callback_data="adm_edit_gmail_task_price"),
         ikb_btn(f"👥 Ref Commission ({ref_bonus}%)", callback_data="adm_edit_ref_bonus_percent")],
        [ikb_btn(tr(user_id, "btn_admin_menu"), callback_data="adm_back_to_admin")]
    ]
    await update.message.reply_text(msg, reply_markup=InlineKeyboardMarkup(keyboard), parse_mode=ParseMode.HTML)

async def admin_withdraw_requests(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM withdrawals WHERE status = 'Pending' ORDER BY created_at DESC LIMIT 10")
        wds = cursor.fetchall()
    if not wds:
        await update.message.reply_text(tr(user_id, "admin_no_pending_wds"), reply_markup=get_admin_keyboard(user_id))
        return
    msg = tr(user_id, "admin_pending_wds_title")
    keyboard = []
    for wd in wds:
        msg += f"🆔 #{wd['id']} | {wd['method'].upper()} | ৳{wd['amount']:.2f} | <code>{wd['number']}</code>\n"
        keyboard.append([
            ikb_btn(f"✅ App #{wd['id']}", callback_data=f"adm_app_wd_{wd['id']}"),
            ikb_btn(f"❌ Rej #{wd['id']}", callback_data=f"adm_rej_wd_{wd['id']}")
        ])
    keyboard.append([ikb_btn(tr(user_id, "btn_admin_menu"), callback_data="adm_back_to_admin")])
    await update.message.reply_text(msg, reply_markup=InlineKeyboardMarkup(keyboard), parse_mode=ParseMode.HTML)

async def admin_system_settings(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return
    force_join = get_setting_val("force_join_enabled", "OFF")
    maint = get_setting_val("maintenance_mode", "OFF")
    gmail = get_setting_val("gmail_task_active", "ON")
    bulk_gmail = get_setting_val("bulk_gmail_task_active", "ON")
    old_gmail = get_setting_val("old_gmail_task_active", "ON")
    old_2fa = get_setting_val("old_gmail_2fa_active", "ON")
    
    bkash_wd = get_setting_val("wd_bkash_active", "ON")
    nagad_wd = get_setting_val("wd_nagad_active", "ON")
    usdt_wd = get_setting_val("wd_usdt_active", "ON")

    min_bkash = get_setting_val("min_withdraw_bkash", "50.0")
    max_bkash = get_setting_val("max_withdraw_bkash", "5000.0")
    min_nagad = get_setting_val("min_withdraw_nagad", "50.0")
    max_nagad = get_setting_val("max_withdraw_nagad", "5000.0")
    min_usdt = get_setting_val("min_withdraw_usdt", "0.25")
    max_usdt = get_setting_val("max_withdraw_usdt", "100.0")
    usdt_rate = get_setting_val("usdt_rate_bdt", "120.0")
    charge_status = get_setting_val("usdt_charge_enabled", "ON")
    charge_usdt = get_setting_val("usdt_charge_usdt", "0.05")
    bkash_charge = get_setting_val("bkash_charge_bdt", "0.0")
    nagad_charge = get_setting_val("nagad_charge_bdt", "0.0")

    msg = tr(
        user_id, "admin_sys_settings_title",
        bot_name=get_setting_val("bot_display_name", "Gmail Work BD"),
        support_link=get_setting_val("support_handle", ""),
        channel_link=get_setting_val("official_channel_link", ""),
        force_join=force_join, maint=maint, gmail=gmail, bulk_gmail=bulk_gmail, old_gmail=old_gmail, old_2fa=old_2fa,
        bkash_wd=bkash_wd, nagad_wd=nagad_wd, usdt_wd=usdt_wd,
        min_bkash=min_bkash, max_bkash=max_bkash,
        min_nagad=min_nagad, max_nagad=max_nagad,
        min_usdt=min_usdt, max_usdt=max_usdt, usdt_rate=usdt_rate,
        charge_status=charge_status, charge_usdt=charge_usdt, bkash_charge=bkash_charge, nagad_charge=nagad_charge
    )
    keyboard = [
        [ikb_btn(f"Force Join: {force_join}", callback_data="adm_toggle_force_join_enabled"),
         ikb_btn("📢 Edit Force Channels", callback_data="adm_edit_force_channels_list")],
        [ikb_btn(f"Maintenance: {maint}", callback_data="adm_toggle_maintenance_mode")],
        [ikb_btn(f"Bkash: {bkash_wd}", style="danger", callback_data="adm_toggle_wd_bkash_active"),
         ikb_btn("Bkash Min/Max", style="primary", callback_data="adm_edit_bkash_limits"),
         ikb_btn(f"Charge: ৳{bkash_charge}", style="success", callback_data="adm_edit_bkash_charge_bdt")],
        [ikb_btn(f"Nagad: {nagad_wd}", style="primary", callback_data="adm_toggle_wd_nagad_active"),
         ikb_btn("Nagad Min/Max", style="success", callback_data="adm_edit_nagad_limits"),
         ikb_btn(f"Charge: ৳{nagad_charge}", style="danger", callback_data="adm_edit_nagad_charge_bdt")],
        [ikb_btn(f"USDT: {usdt_wd}", callback_data="adm_toggle_wd_usdt_active"),
         ikb_btn("✏️ USDT Min/Max", callback_data="adm_edit_usdt_limits")],
        [ikb_btn("💱 Edit USDT Rate", callback_data="adm_edit_usdt_rate_bdt"),
         ikb_btn(f"💳 Charge: {charge_status}", callback_data="adm_toggle_usdt_charge_enabled")],
        [ikb_btn("✏️ Edit Charge (USDT)", callback_data="adm_edit_usdt_charge_usdt")],
        [ikb_btn(f"New Gmail: {gmail}", style="success", callback_data="adm_toggle_gmail_task_active"),
         ikb_btn(f"Bulk Gmail: {bulk_gmail}", style="primary", callback_data="adm_toggle_bulk_gmail_task_active")],
        [ikb_btn(f"Old Gmail: {old_gmail}", style="danger", callback_data="adm_toggle_old_gmail_task_active"),
         ikb_btn(f"Old 2FA: {old_2fa}", style="success", callback_data="adm_toggle_old_gmail_2fa_active")],
        [ikb_btn("🏷️ Bot Name", style="primary", callback_data="adm_edit_bot_display_name"),
         ikb_btn("🎧 Support ID/Link", style="success", callback_data="adm_edit_support_handle")],
        [ikb_btn("📢 Official Channel", style="danger", callback_data="adm_edit_official_channel_link"),
         ikb_btn("🖼️ Welcome Banner", style="primary", callback_data="adm_edit_welcome_banner")],
        [ikb_btn(tr(user_id, "btn_admin_menu"), callback_data="adm_back_to_admin")]
    ]
    await update.message.reply_text(msg, reply_markup=InlineKeyboardMarkup(keyboard), parse_mode=ParseMode.HTML)

async def admin_live_stats(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM users")
        total_users = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM task_submissions WHERE status = 'Pending'")
        pending_tasks = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM task_submissions WHERE status = 'Hold'")
        held_tasks = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM task_submissions WHERE status = 'Approved'")
        approved_tasks = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM task_submissions WHERE status = 'Rejected'")
        rejected_tasks = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM withdrawals WHERE status = 'Pending'")
        pending_wd = cursor.fetchone()[0]
        cursor.execute("SELECT SUM(balance) FROM users")
        total_balance = cursor.fetchone()[0] or 0.0

    msg = tr(
        user_id, "admin_live_stats_title",
        total_users=total_users,
        pending_tasks=pending_tasks,
        held_tasks=held_tasks,
        approved_tasks=approved_tasks,
        rejected_tasks=rejected_tasks,
        pending_wd=pending_wd,
        total_balance=total_balance
    )
    await update.message.reply_text(msg, reply_markup=get_admin_keyboard(user_id), parse_mode=ParseMode.MARKDOWN)


async def admin_backup_db(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return

    if not os.path.exists(DB_FILE):
        await update.message.reply_text("❌ ডাটাবেস ফাইল পাওয়া যায়নি।", reply_markup=get_admin_keyboard(user_id))
        return

    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    tmp_path = f"backup_{ts}.db"
    try:
        src_conn = sqlite3.connect(DB_FILE, timeout=30.0)
        dst_conn = sqlite3.connect(tmp_path)
        with dst_conn:
            src_conn.backup(dst_conn)
        dst_conn.close()
        src_conn.close()

        with get_db() as conn:
            c = conn.cursor()
            c.execute("SELECT COUNT(*) FROM users")
            u_cnt = c.fetchone()[0]
            c.execute("SELECT COUNT(*) FROM task_submissions")
            t_cnt = c.fetchone()[0]

        with open(tmp_path, "rb") as f:
            await update.message.reply_document(
                document=f,
                filename=tmp_path,
                caption=(
                    f"💾 <b>Database Backup Complete</b>\n"
                    f"━━━━━━━━━━━━━━━━━━━━━━\n"
                    f"👥 <b>Total Users:</b> {u_cnt}\n"
                    f"📧 <b>Total Tasks:</b> {t_cnt}\n"
                    f"🕒 <b>Time:</b> {ts}"
                ),
                parse_mode=ParseMode.HTML,
                reply_markup=get_admin_keyboard(user_id)
            )
    except Exception as e:
        logger.error(f"Database backup failed: {e}")
        await update.message.reply_text(f"❌ ব্যাকআপ নিতে সমস্যা হয়েছে: {e}", reply_markup=get_admin_keyboard(user_id))
    finally:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass



async def process_admin_welcome_banner(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return ConversationHandler.END

    if update.message.photo:
        file_id = update.message.photo[-1].file_id
        set_setting_val("welcome_banner_file_id", file_id)
        await update.message.reply_text(
            "✅ <b>Welcome Banner সফলভাবে সেট হয়েছে!</b>\nএখন /start দিলে ওয়েলকাম মেসেজের সাথে এই ব্যানার দেখাবে।",
            reply_markup=get_admin_keyboard(user_id),
            parse_mode=ParseMode.HTML
        )
        clear_user_state(context)
        return ConversationHandler.END

    txt = (update.message.text or "").strip().lower()
    if txt in ("off", "remove", "clear", "none", "delete"):
        set_setting_val("welcome_banner_file_id", "")
        await update.message.reply_text(
            "🗑️ <b>Welcome Banner মুছে ফেলা হয়েছে।</b>",
            reply_markup=get_admin_keyboard(user_id),
            parse_mode=ParseMode.HTML
        )
        clear_user_state(context)
        return ConversationHandler.END

    await update.message.reply_text(
        "❌ অনুগ্রহ করে একটি <b>Photo (ছবি)</b> পাঠান অথবা মুছতে <code>off</code> লিখুন।",
        reply_markup=get_cancel_keyboard(user_id),
        parse_mode=ParseMode.HTML
    )
    return STATE_ADMIN_WELCOME_BANNER

async def admin_restore_prompt(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return ConversationHandler.END
    await update.message.reply_text(
        "📥 <b>Restore Database</b>\n━━━━━━━━━━━━━━━━━━━━━━\n"
        "আপনার ব্যাকআপ <code>.db</code> ফাইলটি এখানে Document হিসেবে পাঠান।\n"
        "⚠️ এটি বর্তমান ডাটাবেসকে ব্যাকআপের ডাটা দিয়ে প্রতিস্থাপন করবে।",
        reply_markup=get_cancel_keyboard(user_id),
        parse_mode=ParseMode.HTML
    )
    return STATE_ADMIN_RESTORE_DB


async def process_admin_restore_db(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return ConversationHandler.END

    if not update.message.document:
        await update.message.reply_text(
            "❌ অনুগ্রহ করে একটি <code>.db</code> ব্যাকআপ ফাইল পাঠান (অথবা Cancel চাপুন)।",
            reply_markup=get_cancel_keyboard(user_id),
            parse_mode=ParseMode.HTML
        )
        return STATE_ADMIN_RESTORE_DB

    doc = update.message.document
    if not (doc.file_name or "").lower().endswith(".db"):
        await update.message.reply_text(
            "❌ শুধু <code>.db</code> ফরম্যাটের ব্যাকআপ ফাইল গ্রহণযোগ্য।",
            reply_markup=get_cancel_keyboard(user_id),
            parse_mode=ParseMode.HTML
        )
        return STATE_ADMIN_RESTORE_DB

    tmp_restore = f"restore_tmp_{int(time.time())}.db"
    try:
        tg_file = await doc.get_file()
        await tg_file.download_to_drive(tmp_restore)

        # Validate SQLite integrity and required tables
        test_conn = sqlite3.connect(tmp_restore)
        cur = test_conn.cursor()
        cur.execute("PRAGMA integrity_check")
        chk = cur.fetchone()
        if not chk or chk[0].lower() != "ok":
            test_conn.close()
            raise ValueError("Invalid or corrupted SQLite database file.")
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name IN ('users', 'task_submissions')")
        tables = {r[0] for r in cur.fetchall()}
        if not {"users", "task_submissions"}.issubset(tables):
            test_conn.close()
            raise ValueError("Missing required bot tables in backup file.")

        cur.execute("SELECT COUNT(*) FROM users")
        u_cnt = cur.fetchone()[0]
        cur.execute("SELECT COUNT(*) FROM task_submissions")
        t_cnt = cur.fetchone()[0]

        # Restore into active DB_FILE safely
        dst_conn = sqlite3.connect(DB_FILE, timeout=30.0)
        with dst_conn:
            test_conn.backup(dst_conn)
        dst_conn.close()
        test_conn.close()

        _db_pool.reload_pool()
        init_db()

        await update.message.reply_text(
            f"✅ <b>Database Restored Successfully!</b>\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"👥 <b>Restored Users:</b> {u_cnt}\n"
            f"📧 <b>Restored Tasks:</b> {t_cnt}",
            reply_markup=get_admin_keyboard(user_id),
            parse_mode=ParseMode.HTML
        )
    except Exception as e:
        logger.error(f"Database restore failed: {e}")
        await update.message.reply_text(
            f"❌ রিস্টোর করা যায়নি: {html.escape(str(e))}",
            reply_markup=get_admin_keyboard(user_id),
            parse_mode=ParseMode.HTML
        )
    finally:
        if os.path.exists(tmp_restore):
            try:
                os.remove(tmp_restore)
            except Exception:
                pass

    clear_user_state(context)
    return ConversationHandler.END


async def auto_backup_loop(app: Application):
    """Automatically send a safe DB backup to the primary owner every 6 hours."""
    await asyncio.sleep(60)
    while True:
        try:
            admin_id_setting = ADMIN_ID or get_setting_val("admin_id", "")
            owner_id = next((int(x.strip()) for x in admin_id_setting.split(",") if x.strip().isdigit()), None)
            if owner_id and os.path.exists(DB_FILE):
                with get_db() as conn:
                    c = conn.cursor()
                    c.execute("SELECT COUNT(*) FROM users")
                    u_cnt = c.fetchone()[0]
                if u_cnt > 0:
                    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    tmp_path = f"auto_backup_{ts}.db"
                    try:
                        src_conn = sqlite3.connect(DB_FILE, timeout=30.0)
                        dst_conn = sqlite3.connect(tmp_path)
                        with dst_conn:
                            src_conn.backup(dst_conn)
                        dst_conn.close()
                        src_conn.close()
                        with open(tmp_path, "rb") as f:
                            await app.bot.send_document(
                                chat_id=owner_id,
                                document=f,
                                filename=tmp_path,
                                caption=f"🔄 <b>Auto Backup</b> ({u_cnt} users) — {ts}",
                                parse_mode=ParseMode.HTML
                            )
                    finally:
                        if os.path.exists(tmp_path):
                            os.remove(tmp_path)
        except Exception as e:
            logger.error(f"Auto backup loop error: {e}")
        await asyncio.sleep(6 * 3600)


async def _on_startup(app: Application):
    asyncio.create_task(auto_backup_loop(app))

async def admin_broadcast_prompt(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return
    context.user_data["admin_broadcast"] = True
    await update.message.reply_text(
        tr(user_id, "admin_broadcast_prompt_msg"),
        reply_markup=get_cancel_keyboard(user_id),
        parse_mode=ParseMode.MARKDOWN
    )
    return STATE_ADMIN_BROADCAST

# -----------------------------------------------------------------------------
# ADMIN CALLBACK QUERY HANDLER
# -----------------------------------------------------------------------------
async def admin_callback_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    user_id = query.from_user.id
    await query.answer()

    if not is_admin(user_id):
        await query.edit_message_text("❌ You are not authorized.")
        return

    data = query.data

    if data == "adm_back_to_admin":
        await safe_delete_message(context, query.message.chat_id, query.message.message_id)
        await context.bot.send_message(
            chat_id=query.message.chat_id,
            text="🔑 <b>Admin Panel</b>",
            reply_markup=get_admin_keyboard(user_id),
            parse_mode=ParseMode.HTML
        )
        return

    if data.startswith("adm_toggle_"):
        key = data.replace("adm_toggle_", "")
        current = get_setting_val(key, "OFF")
        new_val = "ON" if current == "OFF" else "OFF"
        set_setting_val(key, new_val)
        await query.edit_message_text(f"✅ <code>{html.escape(key)}</code> → <b>{html.escape(new_val)}</b>", parse_mode=ParseMode.HTML)
        return

    if data == "adm_edit_force_channels_list":
        raw_list = get_setting_val(
            "force_channels_list",
            "@TrustVaultMailsOfficial|https://t.me/TrustVaultMailsOfficial"
        )
        channels = parse_force_join_channels(raw_list)
        fj_status = get_setting_val("force_join_enabled", "OFF")
        buttons = []
        for idx, (ch_id, _link) in enumerate(channels):
            buttons.append([ikb_btn(f"❌ Remove {ch_id}", style="danger", callback_data=f"adm_fj_del_{idx}")])
        buttons.append([
            ikb_btn("➕ Add Channel", style="success", callback_data="adm_fj_add"),
            ikb_btn(f"Force Join: {fj_status}", style="primary", callback_data="adm_toggle_force_join_enabled")
        ])
        if channels:
            buttons.append([ikb_btn("🗑️ Clear All Channels", style="danger", callback_data="adm_fj_clear")])
        buttons.append([ikb_btn("Admin Menu", style="primary", callback_data="adm_back_to_admin")])

        ch_lines = "\n".join(f"{i+1}. <code>{html.escape(c[0])}</code>" for i, c in enumerate(channels)) if channels else "কোনো চ্যানেল যোগ করা নেই।"
        await query.edit_message_text(
            f"📢 <b>Force Join Channels Manager ({len(channels)} টি)</b>\n━━━━━━━━━━━━━━━━━━━━━━\n"
            f"{ch_lines}\n\n"
            f"👉 নতুন চ্যানেল যোগ করতে <b>➕ Add Channel</b> চাপুন (একাধিক চ্যানেল একসাথে বা একে একে যোগ করা যাবে)।",
            reply_markup=InlineKeyboardMarkup(buttons),
            parse_mode=ParseMode.HTML
        )
        return

    if data == "adm_fj_add":
        await query.edit_message_text(
            "➕ <b>নতুন Force Join চ্যানেল যোগ করুন</b>\n━━━━━━━━━━━━━━━━━━━━━━\n"
            "একটি বা একাধিক চ্যানেল প্রতি লাইনে দিন (আগের চ্যানেলগুলো মুছবে না):\n\n"
            "<b>Public:</b> <code>@channelusername</code>\n"
            "<b>Private:</b> <code>-1001234567890|https://t.me/+invite_link</code>\n\n"
            "⚠️ বটকে ওই চ্যানেলের Admin রাখতে হবে।",
            parse_mode=ParseMode.HTML
        )
        return STATE_ADMIN_FORCE_CHANNELS_EDIT

    if data.startswith("adm_fj_del_"):
        idx = int(data.replace("adm_fj_del_", ""))
        raw_list = get_setting_val("force_channels_list", "")
        channels = parse_force_join_channels(raw_list)
        if 0 <= idx < len(channels):
            removed = channels.pop(idx)
            set_setting_val("force_channels_list", "\n".join(f"{c[0]}|{c[1]}" for c in channels))
            await query.edit_message_text(
                f"✅ চ্যানেল <code>{html.escape(removed[0])}</code> রিমুভ করা হয়েছে! (বাকি: {len(channels)} টি)",
                reply_markup=InlineKeyboardMarkup([[ikb_btn("📢 Back to Force Channels", style="primary", callback_data="adm_edit_force_channels_list")]]),
                parse_mode=ParseMode.HTML
            )
        return

    if data == "adm_fj_clear":
        set_setting_val("force_channels_list", "")
        await query.edit_message_text(
            "🗑️ সব Force Join চ্যানেল ক্লিয়ার করা হয়েছে!",
            reply_markup=InlineKeyboardMarkup([[ikb_btn("📢 Back to Force Channels", style="primary", callback_data="adm_edit_force_channels_list")]])
        )
        return

    if data in ["adm_edit_bkash_limits", "adm_edit_nagad_limits", "adm_edit_usdt_limits"]:
        m_name = data.replace("adm_edit_", "").replace("_limits", "")
        context.user_data["edit_limits_method"] = m_name
        current_min = get_setting_val(f"min_withdraw_{m_name}", "50.0")
        current_max = get_setting_val(f"max_withdraw_{m_name}", "5000.0")
        await query.edit_message_text(
            f"✏️ **{m_name.upper()} উইথড্র লিমিট সেট করুন:**\n\n"
            f"বর্তমান সর্বনিম্ন: `{current_min}` | সর্বোচ্চ: `{current_max}`\n\n"
            f"পাঠানোর ফরম্যাট: `সর্বনিম্ন সর্বোচ্চ` (যেমন: `50 5000`)",
            parse_mode=ParseMode.MARKDOWN
        )
        return STATE_ADMIN_SETTING_VAL

    if data == "adm_oldcat_add":
        await query.edit_message_text(
            "➕ <b>নতুন Old Gmail ক্যাটাগরি যোগ করুন</b>\n━━━━━━━━━━━━━━━━━━━━━━\n"
            "ফরম্যাট: <code>নাম | রেট</code>\n\n"
            "উদাহরণ:\n<code>2015 - 2018 Old | 80</code>\n<code>2019 - 2025 Old | 35</code>",
            parse_mode=ParseMode.HTML
        )
        return STATE_ADMIN_OLD_CAT_ADD

    if data.startswith("adm_oldcat_del_"):
        ckey = data.replace("adm_oldcat_del_", "", 1)
        del_old_category(ckey)
        await query.edit_message_text(
            "🗑️ ক্যাটাগরি মুছে ফেলা হয়েছে।",
            reply_markup=InlineKeyboardMarkup([[ikb_btn("Admin Menu", style="primary", callback_data="adm_back_to_admin")]])
        )
        return

    if data == "adm_edit_welcome_banner":
        has_banner = "✅ Set" if get_setting_val("welcome_banner_file_id", "").strip() else "❌ None"
        await query.edit_message_text(
            f"🖼️ <b>Welcome Banner ({has_banner})</b>\n━━━━━━━━━━━━━━━━━━━━━━\n"
            f"ওয়েলকাম মেসেজের সাথে দেখানোর জন্য যেকোনো একটি <b>Photo (ছবি)</b> এখানে পাঠান।\n\n"
            f"ব্যানার মুছে ফেলতে চাইলে <code>off</code> লিখে পাঠান।",
            parse_mode=ParseMode.HTML
        )
        return STATE_ADMIN_WELCOME_BANNER

    if data.startswith("adm_edit_"):
        setting_key = data.replace("adm_edit_", "")
        context.user_data["edit_setting_key"] = setting_key
        await query.edit_message_text(f"✏️ Send new value for <code>{html.escape(setting_key)}</code>:", parse_mode=ParseMode.HTML)
        return STATE_ADMIN_SETTING_VAL

    # Sub Admin Controls
    if data == "adm_add_subadmin":
        if not is_owner(user_id):
            await query.answer("❌ শুধুমাত্র প্রধান মালিক সাব-এডমিন যোগ করতে পারবেন!", show_alert=True)
            return
        await query.edit_message_text("➕ **নতুন সাব-এডমিনের টেলিগ্রাম ইউজার আইডি পাঠাবে:**")
        return STATE_ADMIN_ADD_SUBADMIN

    if data == "adm_rem_subadmin":
        if not is_owner(user_id):
            await query.answer("❌ শুধুমাত্র প্রধান মালিক সাব-এডমিন রিমুভ করতে পারবেন!", show_alert=True)
            return
        await query.edit_message_text("➖ **রিমুভ করার জন্য সাব-এডমিনের টেলিগ্রাম ইউজার আইডি পাঠাবে:**")
        return STATE_ADMIN_REM_SUBADMIN

    # User Management Callbacks
    if data == "adm_usr_reset_balances":
        buttons = [
            [ikb_btn("✅ হ্যাঁ, সবার ব্যালেন্স ০ করুন", callback_data="adm_usr_reset_balances_confirm")],
            [ikb_btn("❌ বাতিল", callback_data="adm_back_to_admin")]
        ]
        await query.edit_message_text(
            "⚠️ <b>সতর্কতা!</b>\n\nআপনি কি সত্যিই সব ইউজারের বর্তমান ব্যালেন্স ০ করতে চান?\n\nএই কাজটি পূর্বাবস্থায় ফিরিয়ে আনা যাবে না। আগে ব্যাকআপ নিন।",
            reply_markup=InlineKeyboardMarkup(buttons),
            parse_mode=ParseMode.HTML
        )
        return

    if data == "adm_usr_reset_balances_confirm":
        with get_db() as conn:
            cursor = conn.cursor()
            cursor.execute("UPDATE users SET balance = 0.0")
            conn.commit()
        await query.edit_message_text(
            "✅ <b>সফলভাবে সব ইউজারের ব্যালেন্স ০ করা হয়েছে।</b>",
            reply_markup=InlineKeyboardMarkup([[ikb_btn("Admin Menu", callback_data="adm_back_to_admin")]]),
            parse_mode=ParseMode.HTML
        )
        logger.warning("ADMIN %s reset all user balances to zero", user_id)
        return

    if data == "adm_usr_ban":
        await query.edit_message_text("🚫 **ব্যান করার জন্য ইউজারের টেলিগ্রাম আইডি পাঠান:**")
        return STATE_ADMIN_BAN_USER

    if data == "adm_usr_unban":
        await query.edit_message_text("✅ **আনব্যান করার জন্য ইউজারের টেলিগ্রাম আইডি পাঠান:**")
        return STATE_ADMIN_UNBAN_USER

    if data == "adm_usr_add_bal":
        await query.edit_message_text(
            "💰 **ইউজারের ওয়ালেটে ব্যালেন্স যোগ করুন:**\n\nফরম্যাট: `ইউজার_আইডি পরিমাণ` (যেমন: `123456789 50`)",
            parse_mode=ParseMode.MARKDOWN
        )
        return STATE_ADMIN_ADD_BAL

    if data == "adm_usr_sub_bal":
        await query.edit_message_text(
            "💸 **ইউজারের ওয়ালেট থেকে ব্যালেন্স কাটুন:**\n\nফরম্যাট: `ইউজার_আইডি পরিমাণ` (যেমন: `123456789 20`)",
            parse_mode=ParseMode.MARKDOWN
        )
        return STATE_ADMIN_SUB_BAL

    if data == "adm_usr_banned_list":
        with get_db() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT user_id, username, full_name FROM users WHERE is_banned = 1")
            banned = cursor.fetchall()
        if not banned:
            msg = "📜 **কোনো ব্যান ইউজার পাওয়া যায়নি।**"
        else:
            msg = "📜 **নিষিদ্ধ (Banned) ইউজার লিস্ট:**\n━━━━━━━━━━━━━━━━━━━━━━\n"
            for u in banned:
                uname = f"@{u['username']}" if u['username'] else "No Username"
                msg += f"• 🆔 ID: <code>{u['user_id']}</code> | Username: {uname} | Name: {html.escape(u['full_name'] or 'N/A')}\n"
        
        buttons = [[ikb_btn("Admin Menu", callback_data="adm_back_to_admin")]]
        await query.edit_message_text(msg, reply_markup=InlineKeyboardMarkup(buttons), parse_mode=ParseMode.HTML)
        return

    if data == "adm_usr_set_ref":
        await query.edit_message_text(
            "🎁 <b>নির্দিষ্ট ইউজারের জন্য কাস্টম রেফার বোনাস সেট করুন:</b>\n\n"
            "• প্রতি এপ্রুভড কাজে ৪ বা ৫ টাকা দিতে লিখুন: <code>ইউজার_আইডি টাকা</code> (যেমন: <code>123456789 5</code>)\n"
            "• আগের সাধারণ নিয়মে (%) ফেরত নিতে লিখুন: <code>ইউজার_আইডি default</code>",
            parse_mode=ParseMode.HTML
        )
        return STATE_ADMIN_SET_REF_BONUS

    if data == "adm_usr_search":
        await query.edit_message_text("🔍 **প্রোফাইল সার্চ করতে ইউজারের Telegram ID বা Username লিখুন:**")
        return STATE_ADMIN_USER_SEARCH

    # Approve Task (With Lifetime Referral Commission)
    if data.startswith("adm_app_sub_"):
        sub_id = int(data.replace("adm_app_sub_", ""))
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with get_db() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM task_submissions WHERE submission_id = ?", (sub_id,))
            sub = cursor.fetchone()
            if not sub or sub["status"] not in ("Pending", "Hold"):
                # The task may have been processed through bulk approve/reject.
                # Do not show an unnecessary "already processed" popup.
                try:
                    await query.answer()
                except Exception:
                    pass
                try:
                    await query.edit_message_reply_markup(reply_markup=None)
                except Exception:
                    pass
                return

            cursor.execute("UPDATE users SET balance = balance + ? WHERE user_id = ?", (sub["reward_amount"], sub["user_id"]))
            cursor.execute("UPDATE task_submissions SET status = 'Approved', updated_at = ? WHERE submission_id = ?", (now, sub_id))

            cursor.execute("SELECT referred_by FROM users WHERE user_id = ?", (sub["user_id"],))
            user_ref = cursor.fetchone()
            if user_ref and user_ref["referred_by"]:
                referrer_id = user_ref["referred_by"]
                commission_bonus = get_referrer_commission(referrer_id, float(sub["reward_amount"]))
                if commission_bonus > 0:
                    cursor.execute(
                        "UPDATE users SET balance = balance + ?, ref_earnings = ref_earnings + ? WHERE user_id = ?",
                        (commission_bonus, commission_bonus, referrer_id)
                    )
                    try:
                        await context.bot.send_message(
                            chat_id=referrer_id,
                            text=tr(referrer_id, "user_ref_commission", bonus=commission_bonus),
                            parse_mode=ParseMode.HTML
                        )
                    except Exception as e:
                        logger.error(f"Error notifying referrer {referrer_id}: {e}")
            conn.commit()

        task_display = sub["task_type"].upper()
        sub_username = html.escape(str(sub['submitted_username'] or '').strip())
        try:
            await context.bot.send_message(
                chat_id=sub["user_id"],
                text=tr(sub["user_id"], "user_task_approved", task_display=task_display, username=sub_username, reward=sub['reward_amount']),
                parse_mode=ParseMode.HTML
            )
        except Exception as e:
            logger.error(f"Error sending approval DM to user {sub['user_id']}: {e}")

        await query.edit_message_text(f"✅ Task #{sub_id} successfully approved!")
        return

    # Hold Task
    if data.startswith("adm_hold_sub_"):
        sub_id = int(data.replace("adm_hold_sub_", ""))
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with get_db() as conn:
            cursor = conn.cursor()
            cursor.execute("UPDATE task_submissions SET status = 'Hold', updated_at = ? WHERE submission_id = ?", (now, sub_id))
            cursor.execute("SELECT user_id FROM task_submissions WHERE submission_id = ?", (sub_id,))
            row = cursor.fetchone()
            conn.commit()

        if row:
            target_user_id = row["user_id"]
            try:
                await context.bot.send_message(
                    chat_id=target_user_id,
                    text=tr(target_user_id, "user_task_held", sub_id=sub_id),
                    parse_mode=ParseMode.HTML
                )
            except Exception as e:
                logger.error(f"Failed to notify user {target_user_id} of hold: {e}")

        await query.edit_message_text(f"⌛ Task #{sub_id} placed on Hold.")
        return

    # Reject Prompt
    if data.startswith("adm_rej_prompt_"):
        sub_id = int(data.replace("adm_rej_prompt_", ""))
        buttons = [
            [ikb_btn("📋 Default Reason", callback_data=f"adm_rej_act_{sub_id}_default")],
            [ikb_btn("🔑 Wrong Password", callback_data=f"adm_rej_act_{sub_id}_wrongpass")],
            [ikb_btn("📝 Custom Reason", callback_data=f"adm_rej_custom_{sub_id}")]
        ]
        await query.edit_message_text("❌ Select rejection reason:", reply_markup=InlineKeyboardMarkup(buttons))
        return

    # Reject Action
    if data.startswith("adm_rej_act_"):
        parts = data.replace("adm_rej_act_", "").split("_")
        sub_id = int(parts[0])
        reason_key = parts[1]
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        if reason_key == "default":
            reason = "আপনার Gmail অ্যাকাউন্ট লগইন করা সম্ভব হয়নি। তাই অ্যাকাউন্টটি রিজেক্ট করা হয়েছে।"
        elif reason_key == "wrongpass":
            reason = "Wrong password or username mismatch"
        else:
            reason = "Submission criteria not met"

        with get_db() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM task_submissions WHERE submission_id = ?", (sub_id,))
            sub = cursor.fetchone()

            if not sub or sub["status"] not in ("Pending", "Hold"):
                try:
                    await query.answer()
                except Exception:
                    pass
                try:
                    await query.edit_message_reply_markup(reply_markup=None)
                except Exception:
                    pass
                return

            cursor.execute(
                "UPDATE task_submissions SET status = 'Rejected', reject_reason = ?, updated_at = ? WHERE submission_id = ?",
                (reason, now, sub_id)
            )
            conn.commit()

        if sub:
            task_display = sub["task_type"].upper()
            sub_username = html.escape(str(sub['submitted_username'] or '').strip())
            esc_reason = html.escape(reason)
            try:
                await context.bot.send_message(
                    chat_id=sub["user_id"],
                    text=tr(sub["user_id"], "user_task_rejected", task_display=task_display, username=sub_username, reason=esc_reason),
                    parse_mode=ParseMode.HTML
                )
            except Exception as e:
                logger.error(f"Error sending reject DM to user {sub['user_id']}: {e}")

        await query.edit_message_text(f"❌ Task #{sub_id} rejected.")
        return

    # Approve Withdraw
    if data.startswith("adm_app_wd_"):
        wd_id = int(data.replace("adm_app_wd_", ""))
        with get_db() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM withdrawals WHERE id = ?", (wd_id,))
            wd = cursor.fetchone()
            if not wd or wd["status"] != "Pending":
                await query.edit_message_text("❌ Already approved or processed!")
                return
            cursor.execute("UPDATE withdrawals SET status = 'Approved' WHERE id = ?", (wd_id,))
            conn.commit()
        try:
            approved_amount = wd["usdt_amount"] if wd["method"] in ("usdt", "binance_uid") and wd["usdt_amount"] is not None else wd["amount"]
            approved_text = f"${approved_amount:.4f} USDT" if wd["method"] in ("usdt", "binance_uid") else f"৳{approved_amount:.2f}"
            approval_msg = tr(wd["user_id"], "user_wd_approved", wd_id=wd_id, amount=approved_amount, number=wd["number"]).replace(f"৳{approved_amount:.2f}", approved_text)
            await context.bot.send_message(
                chat_id=wd["user_id"],
                text=approval_msg,
                parse_mode=ParseMode.HTML
            )
        except Exception as e:
            logger.error(f"Failed to notify user of withdrawal approval: {e}")

        await query.edit_message_text(f"✅ Withdraw #{wd_id} (৳{wd['amount']}) approved!")
        return

    # Reject Withdraw
    if data.startswith("adm_rej_wd_"):
        wd_id = int(data.replace("adm_rej_wd_", ""))
        with get_db() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM withdrawals WHERE id = ?", (wd_id,))
            wd = cursor.fetchone()
            if not wd or wd["status"] != "Pending":
                await query.edit_message_text("❌ Already rejected or processed!")
                return
            cursor.execute("UPDATE users SET balance = balance + ? WHERE user_id = ?", (wd["amount"], wd["user_id"]))
            cursor.execute("UPDATE withdrawals SET status = 'Rejected' WHERE id = ?", (wd_id,))
            conn.commit()
        try:
            rejected_amount = wd["usdt_amount"] if wd["method"] in ("usdt", "binance_uid") and wd["usdt_amount"] is not None else wd["amount"]
            rejected_text = f"${rejected_amount:.4f} USDT" if wd["method"] in ("usdt", "binance_uid") else f"৳{rejected_amount:.2f}"
            rejection_msg = tr(wd["user_id"], "user_wd_rejected", wd_id=wd_id, amount=rejected_amount).replace(f"৳{rejected_amount:.2f}", rejected_text)
            await context.bot.send_message(
                chat_id=wd["user_id"],
                text=rejection_msg,
                parse_mode=ParseMode.HTML
            )
        except Exception as e:
            logger.error(f"Failed to notify user of withdrawal rejection: {e}")

        await query.edit_message_text(f"❌ Withdraw #{wd_id} (৳{wd['amount']}) rejected & refunded to user!")
        return

    if data.startswith("adm_rej_custom_"):
        sub_id = int(data.replace("adm_rej_custom_", ""))
        context.user_data["rej_sub_id"] = sub_id
        await query.edit_message_text(tr(user_id, "admin_rej_custom_prompt"))
        return STATE_ADMIN_REJECT_REASON

# -----------------------------------------------------------------------------
# HIGH-SPEED ASYNC BATCH BROADCAST ENGINE
# -----------------------------------------------------------------------------
async def process_admin_broadcast_media(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return ConversationHandler.END

    msg = update.message

    if msg.text:
        text_str = msg.text.strip()
        if text_str in ["❌ বাতিল", "❌ Cancel", "🏠 মূল মেনু", "🏠 Main Menu", "❌ বাতিল", "❌ Cancel", "🔙 মূল মেনু", "মূল মেনু", "🔙 Main Menu", "Main Menu"]:
            clear_user_state(context)
            await update.message.reply_text(
                tr(user_id, "operation_cancelled"),
                reply_markup=get_admin_keyboard(user_id)
            )
            return ConversationHandler.END

    if msg.text:
        msg_type_str = "📝 Text"
    elif msg.photo:
        msg_type_str = "🖼️ Photo"
    elif msg.video:
        msg_type_str = "🎥 Video"
    elif msg.document:
        msg_type_str = "📄 Document"
    elif msg.audio:
        msg_type_str = "🎵 Audio"
    elif msg.voice:
        msg_type_str = "🎤 Voice"
    elif msg.sticker:
        msg_type_str = "🎨 Sticker"
    elif msg.video_note:
        msg_type_str = "📹 Video Note"
    else:
        msg_type_str = "📦 Media"

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT user_id FROM users WHERE is_banned = 0")
        users = cursor.fetchall()

    total_users = len(users)
    start_time = time.time()

    status_msg = await update.message.reply_text(
        f"🚀 **High-Speed Broadcast Started...**\n\n"
        f"👥 **Total Target Users:** {total_users}\n"
        f"📨 **Type:** {msg_type_str}\n"
        f"⏳ **Status:** Preparing batch send...",
        parse_mode=ParseMode.MARKDOWN
    )

    success_count = 0
    blocked_count = 0
    deleted_count = 0
    other_failed_count = 0

    semaphore = asyncio.Semaphore(25)

    async def send_to_user(target_id):
        nonlocal success_count, blocked_count, deleted_count, other_failed_count
        async with semaphore:
            for attempt in range(2):
                try:
                    await context.bot.copy_message(
                        chat_id=target_id,
                        from_chat_id=update.effective_chat.id,
                        message_id=msg.message_id
                    )
                    success_count += 1
                    return
                except RetryAfter as e:
                    await asyncio.sleep(e.retry_after + 0.2)
                except Forbidden:
                    blocked_count += 1
                    return
                except (BadRequest, TelegramError) as e:
                    err_str = str(e).lower()
                    if "chat not found" in err_str or "user is deactivated" in err_str or "bot was blocked" in err_str:
                        deleted_count += 1
                    else:
                        other_failed_count += 1
                    return
                except Exception:
                    other_failed_count += 1
                    return

    chunk_size = 50
    user_ids = [u["user_id"] for u in users]
    
    for i in range(0, total_users, chunk_size):
        chunk = user_ids[i:i + chunk_size]
        await asyncio.gather(*[send_to_user(uid) for uid in chunk])
        
        processed = min(i + chunk_size, total_users)
        if processed % 200 == 0 or processed == total_users:
            try:
                elapsed = round(time.time() - start_time, 1)
                await status_msg.edit_text(
                    f"⚡ **High-Speed Broadcast Processing...**\n\n"
                    f"👥 **Total:** {total_users}\n"
                    f"⏳ **Processed:** {processed}/{total_users}\n"
                    f"✅ **Sent:** {success_count}\n"
                    f"❌ **Failed:** {blocked_count + deleted_count + other_failed_count}\n"
                    f"⏱️ **Time:** {elapsed}s",
                    parse_mode=ParseMode.MARKDOWN
                )
            except Exception:
                pass
        
        await asyncio.sleep(0.05)

    total_failed = blocked_count + deleted_count + other_failed_count
    success_rate = round((success_count / total_users * 100), 1) if total_users > 0 else 0.0
    elapsed_total = round(time.time() - start_time, 1)

    final_report = tr(
        user_id, "admin_broadcast_done",
        total=total_users,
        msg_type=msg_type_str,
        count=success_count,
        failed=total_failed,
        success_rate=success_rate,
        time_taken=elapsed_total,
        blocked=blocked_count,
        deleted=deleted_count,
        other=other_failed_count
    )

    try:
        await status_msg.edit_text(final_report, parse_mode=ParseMode.MARKDOWN)
    except Exception:
        await update.message.reply_text(final_report, parse_mode=ParseMode.MARKDOWN)

    await update.message.reply_text(
        "🔙 **অ্যাডমিন প্যানেল মেনু:**",
        reply_markup=get_admin_keyboard(user_id)
    )

    clear_user_state(context)
    return ConversationHandler.END

# -----------------------------------------------------------------------------
# PROCESS CUSTOM ADMIN REJECTION REASON
# -----------------------------------------------------------------------------
async def process_admin_reject_reason(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Consume the next admin text message strictly as the custom rejection reason."""
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return ConversationHandler.END

    reason = (update.message.text or "").strip()
    sub_id = context.user_data.get("rej_sub_id")

    if not sub_id:
        await update.message.reply_text(
            "⚠️ কোনো রিজেক্ট টাস্ক নির্বাচন করা নেই।",
            reply_markup=get_admin_keyboard(user_id)
        )
        return ConversationHandler.END

    if not reason:
        await update.message.reply_text(
            "⚠️ রিজেক্ট করার কারণ খালি রাখা যাবে না। আবার কারণটি লিখুন:",
            reply_markup=get_cancel_keyboard(user_id)
        )
        return STATE_ADMIN_REJECT_REASON

    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute(
            "UPDATE task_submissions SET status = 'Rejected', reject_reason = ?, updated_at = ? WHERE submission_id = ?",
            (reason, now, sub_id)
        )
        cursor.execute("SELECT * FROM task_submissions WHERE submission_id = ?", (sub_id,))
        sub = cursor.fetchone()
        conn.commit()

    if sub:
        task_display = sub["task_type"].upper()
        sub_username = html.escape(str(sub["submitted_username"] or "").strip())
        esc_reason = html.escape(reason)
        try:
            await context.bot.send_message(
                chat_id=sub["user_id"],
                text=tr(sub["user_id"], "user_task_rejected", task_display=task_display, username=sub_username, reason=esc_reason),
                parse_mode=ParseMode.HTML
            )
        except Exception as e:
            logger.error(f"Error sending custom reject DM to user {sub['user_id']}: {e}")

    await update.message.reply_text(
        f"❌ Task #{sub_id} rejected.",
        reply_markup=get_admin_keyboard(user_id)
    )
    context.user_data.pop("rej_sub_id", None)
    return ConversationHandler.END


# -----------------------------------------------------------------------------
# PROCESS ADMIN TEXT INPUTS
# -----------------------------------------------------------------------------
async def process_admin_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return ConversationHandler.END

    text = update.message.text.strip()

    if context.user_data.get("admin_search"):
        with get_db() as conn:
            cursor = conn.cursor()
            if text.isdigit():
                cursor.execute("SELECT * FROM task_submissions WHERE submission_id = ? OR user_id = ?", (int(text), int(text)))
            else:
                cursor.execute("SELECT * FROM task_submissions WHERE submitted_username LIKE ?", (f"%{text}%",))
            results = cursor.fetchall()

        if not results:
            await update.message.reply_text(tr(user_id, "admin_search_no_results"), reply_markup=get_admin_keyboard(user_id))
        else:
            msg = tr(user_id, "admin_search_results_title")
            for r in results[:10]:
                msg += f"🆔 #{r['submission_id']} | {r['task_type'].upper()} | {html.escape(r['submitted_username'])} | Status: {r['status']}\n"
            await update.message.reply_text(msg, reply_markup=get_admin_keyboard(user_id), parse_mode=ParseMode.HTML)

        context.user_data["admin_search"] = False
        return ConversationHandler.END

    if context.user_data.get("admin_export"):
        if not text.isdigit():
            await update.message.reply_text(tr(user_id, "admin_export_invalid_id"), reply_markup=get_cancel_keyboard(user_id))
            return STATE_ADMIN_EXPORT_USER

        target_id = int(text)
        with get_db() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM task_submissions WHERE user_id = ?", (target_id,))
            subs = cursor.fetchall()

        if not subs:
            await update.message.reply_text(tr(user_id, "admin_export_no_tasks"), reply_markup=get_admin_keyboard(user_id))
        else:
            buffer = io.StringIO()
            buffer.write(f"GMAIL WORK BD - USER {target_id} REPORT\n")
            buffer.write("="*50+"\n")
            for s in subs:
                buffer.write(f"ID:{s['submission_id']} Type:{s['task_type']} User:{s['submitted_username']} Pass:{s['submitted_password']} Status:{s['status']}\n")

            file_bytes = buffer.getvalue().encode('utf-8')
            bio = io.BytesIO(file_bytes)
            bio.name = f"user_{target_id}.txt"
            await context.bot.send_document(chat_id=update.effective_chat.id, document=bio, caption=f"📂 Report User {target_id}")

        context.user_data["admin_export"] = False
        return ConversationHandler.END

    if context.user_data.get("edit_limits_method"):
        m_name = context.user_data["edit_limits_method"]
        parts = text.split()
        if len(parts) != 2:
            await update.message.reply_text("❌ সঠিক ফরম্যাটে পাঠাবে (যেমন: `50 5000`)")
            return STATE_ADMIN_SETTING_VAL

        try:
            mn = float(parts[0])
            mx = float(parts[1])
            set_setting_val(f"min_withdraw_{m_name}", str(mn))
            set_setting_val(f"max_withdraw_{m_name}", str(mx))
            await update.message.reply_text(f"✅ **{m_name.upper()}** উইথড্র লিমিট সেট হয়েছে:\n• Min: ৳{mn}\n• Max: ৳{mx}", reply_markup=get_admin_keyboard(user_id), parse_mode=ParseMode.MARKDOWN)
        except ValueError:
            await update.message.reply_text("❌ সঠিক সংখ্যা লিখুন।")
            return STATE_ADMIN_SETTING_VAL

        context.user_data["edit_limits_method"] = None
        return ConversationHandler.END

    if context.user_data.get("edit_setting_key"):
        key = context.user_data["edit_setting_key"]
        val = text
        if key in ("support_handle", "official_channel_link"):
            if not val.startswith("http://") and not val.startswith("https://"):
                val = "https://t.me/" + val.lstrip("@").strip()
        set_setting_val(key, val)
        await update.message.reply_text(
            f"✅ <code>{html.escape(key)}</code> updated to <code>{html.escape(val)}</code>.",
            reply_markup=get_admin_keyboard(user_id),
            parse_mode=ParseMode.HTML
        )
        context.user_data["edit_setting_key"] = None
        return ConversationHandler.END

    await update.message.reply_text("🔙 **অ্যাডমিন প্যানেল মেনু:**", reply_markup=get_admin_keyboard(user_id))
    return ConversationHandler.END

async def process_admin_force_channels_edit(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return ConversationHandler.END

    text = update.message.text.strip()
    if is_menu_or_admin_button(text):
        clear_user_state(context)
        return ConversationHandler.END

    channels = parse_force_join_channels(text)
    if not channels:
        await update.message.reply_text(
            "❌ কোনো বৈধ channel পাওয়া যায়নি।\n\n"
            "উদাহরণ: <code>@mychannel|https://t.me/mychannel</code>",
            parse_mode=ParseMode.HTML,
            reply_markup=get_cancel_keyboard(user_id)
        )
        return STATE_ADMIN_FORCE_CHANNELS_EDIT

    existing = parse_force_join_channels(get_setting_val("force_channels_list", ""))
    merged = list(existing)
    seen = {c[0].lower() for c in merged}
    added_count = 0
    for channel_id, join_link in channels:
        if channel_id.lower() not in seen:
            seen.add(channel_id.lower())
            merged.append((channel_id, join_link))
            added_count += 1
    set_setting_val("force_channels_list", "\n".join(f"{c[0]}|{c[1]}" for c in merged))
    await update.message.reply_text(
        f"✅ <b>নতুন {added_count}টি চ্যানেল যোগ হয়েছে!</b>\n\n"
        f"📢 মোট Force Join চ্যানেল: <b>{len(merged)} টি</b>",
        parse_mode=ParseMode.HTML,
        reply_markup=get_admin_keyboard(user_id)
    )
    clear_user_state(context)
    return ConversationHandler.END

async def process_admin_user_search(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return ConversationHandler.END

    text = update.message.text.strip().replace("@", "")
    if is_menu_or_admin_button(text):
        clear_user_state(context)
        return ConversationHandler.END

    with get_db() as conn:
        cursor = conn.cursor()
        if text.isdigit():
            cursor.execute("SELECT * FROM users WHERE user_id = ? OR LOWER(username) = LOWER(?)", (int(text), text))
        else:
            cursor.execute("SELECT * FROM users WHERE LOWER(username) = LOWER(?)", (text,))
        u = cursor.fetchone()

    buttons = [[ikb_btn("Admin Menu", callback_data="adm_back_to_admin")]]

    if not u:
        await update.message.reply_text(
            "❌ **ইউজার পাওয়া যায়নি!**\nসঠিক Telegram User ID অথবা Username লিখুন।",
            reply_markup=InlineKeyboardMarkup(buttons),
            parse_mode=ParseMode.MARKDOWN
        )
    else:
        ban_str = "🚫 Banned" if u["is_banned"] else "✅ Active"
        uname_str = f"@{u['username']}" if u["username"] else "N/A"
        msg = (
            f"👤 **User Profile Details:**\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"🆔 **User ID:** <code>{u['user_id']}</code>\n"
            f"👤 **Username:** {uname_str}\n"
            f"📝 **Full Name:** {html.escape(u['full_name'] or 'N/A')}\n"
            f"💰 **Balance:** ৳{u['balance']:.2f}\n"
            f"🚫 **Ban Status:** {ban_str}\n"
            f"📅 **Joined Date:** {u['joined_at'] or 'N/A'}\n"
            f"👥 **Referral Count:** {u['ref_count']} জন\n"
            f"💵 **Referral Earnings:** ৳{u['ref_earnings']:.2f}\n"
            + (f"🎁 **Custom Ref Bonus:** ৳{float(u['custom_ref_fixed']):g} / approved task\n" if u["custom_ref_fixed"] is not None else "🎁 **Ref Bonus:** Default (%)\n")
            + f"━━━━━━━━━━━━━━━━━━━━━━"
        )
        await update.message.reply_text(
            msg,
            reply_markup=InlineKeyboardMarkup(buttons),
            parse_mode=ParseMode.HTML
        )

    clear_user_state(context)
    return ConversationHandler.END

async def process_admin_add_subadmin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_owner(user_id):
        return ConversationHandler.END

    text = update.message.text.strip()
    if not text.isdigit():
        await update.message.reply_text("❌ সঠিক টেলিগ্রাম ইউজার আইডি পাঠাবে।")
        return STATE_ADMIN_ADD_SUBADMIN

    sub_id = int(text)
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("INSERT OR REPLACE INTO sub_admins (user_id, added_by, created_at) VALUES (?, ?, ?)", (sub_id, user_id, now))
        conn.commit()

    await update.message.reply_text(f"✅ ইউজার আইডি ````{sub_id}```` সফলভাবে সাব-এডমিন হিসেবে যুক্ত করা হয়েছে।", reply_markup=get_admin_keyboard(user_id), parse_mode=ParseMode.MARKDOWN)
    clear_user_state(context)
    return ConversationHandler.END

async def process_admin_rem_subadmin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_owner(user_id):
        return ConversationHandler.END

    text = update.message.text.strip()
    if not text.isdigit():
        await update.message.reply_text("❌ সঠিক টেলিগ্রাম ইউজার আইডি পাঠাবে।")
        return STATE_ADMIN_REM_SUBADMIN

    sub_id = int(text)

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM sub_admins WHERE user_id = ?", (sub_id,))
        conn.commit()

    await update.message.reply_text(f"✅ সাব-এডমিন ````{sub_id}```` সফলভাবে রিমুভ করা হয়েছে।", reply_markup=get_admin_keyboard(user_id), parse_mode=ParseMode.MARKDOWN)
    clear_user_state(context)
    return ConversationHandler.END

async def process_admin_ban_user(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return ConversationHandler.END

    text = update.message.text.strip()
    if not text.isdigit():
        await update.message.reply_text("❌ সঠিক টেলিগ্রাম ইউজার আইডি পাঠাবে।")
        return STATE_ADMIN_BAN_USER

    target_id = int(text)
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("UPDATE users SET is_banned = 1 WHERE user_id = ?", (target_id,))
        conn.commit()

    await update.message.reply_text(f"🚫 ইউজার ````{target_id}```` কে সিস্টেমে ব্যান করা হয়েছে।", reply_markup=get_admin_keyboard(user_id), parse_mode=ParseMode.MARKDOWN)
    clear_user_state(context)
    return ConversationHandler.END

async def process_admin_unban_user(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return ConversationHandler.END

    text = update.message.text.strip()
    if not text.isdigit():
        await update.message.reply_text("❌ সঠিক টেলিগ্রাম ইউজার আইডি পাঠাবে।")
        return STATE_ADMIN_UNBAN_USER

    target_id = int(text)
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("UPDATE users SET is_banned = 0 WHERE user_id = ?", (target_id,))
        conn.commit()

    await update.message.reply_text(f"✅ ইউজার ````{target_id}```` কে আনব্যান করা হয়েছে।", reply_markup=get_admin_keyboard(user_id), parse_mode=ParseMode.MARKDOWN)
    clear_user_state(context)
    return ConversationHandler.END

async def process_admin_add_bal(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return ConversationHandler.END

    text = update.message.text.strip()
    parts = text.split()
    if len(parts) != 2 or not parts[0].isdigit():
        await update.message.reply_text("❌ সঠিক ফরম্যাটে পাঠাবে (যেমন: `123456789 50`)")
        return STATE_ADMIN_ADD_BAL

    target_id = int(parts[0])
    try:
        amount = float(parts[1])
    except ValueError:
        await update.message.reply_text("❌ সঠিক টাকার পরিমাণ লিখুন।")
        return STATE_ADMIN_ADD_BAL

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("UPDATE users SET balance = balance + ? WHERE user_id = ?", (amount, target_id))
        conn.commit()

    try:
        await context.bot.send_message(
            chat_id=target_id,
            text=f"🎉 <b>আপনার ওয়ালেটে ৳{amount:.2f} এডমিন কর্তৃক যুক্ত করা হয়েছে।</b>",
            parse_mode=ParseMode.HTML
        )
    except Exception:
        pass

    await update.message.reply_text(f"✅ ইউজার ````{target_id}```` এর ওয়ালেটে ৳{amount:.2f} সফলভাবে যোগ করা হয়েছে।", reply_markup=get_admin_keyboard(user_id), parse_mode=ParseMode.MARKDOWN)
    clear_user_state(context)
    return ConversationHandler.END

async def process_admin_sub_bal(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return ConversationHandler.END

    text = update.message.text.strip()
    parts = text.split()
    if len(parts) != 2 or not parts[0].isdigit():
        await update.message.reply_text("❌ সঠিক ফরম্যাটে পাঠাবে (যেমন: `123456789 20`)")
        return STATE_ADMIN_SUB_BAL

    target_id = int(parts[0])
    try:
        amount = float(parts[1])
    except ValueError:
        await update.message.reply_text("❌ সঠিক টাকার পরিমাণ লিখুন।")
        return STATE_ADMIN_SUB_BAL

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("UPDATE users SET balance = MAX(0.0, balance - ?) WHERE user_id = ?", (amount, target_id))
        conn.commit()

    try:
        await context.bot.send_message(
            chat_id=target_id,
            text=f"⚠️ <b>আপনার ওয়ালেট থেকে ৳{amount:.2f} কেটে নেওয়া হয়েছে।</b>",
            parse_mode=ParseMode.HTML
        )
    except Exception:
        pass

    await update.message.reply_text(f"✅ ইউজার ````{target_id}```` এর ওয়ালেট থেকে ৳{amount:.2f} কেটে নেওয়া হয়েছে।", reply_markup=get_admin_keyboard(user_id), parse_mode=ParseMode.MARKDOWN)
    clear_user_state(context)
    return ConversationHandler.END


async def process_admin_set_ref_bonus(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return ConversationHandler.END

    text = update.message.text.strip()
    parts = text.split()
    if len(parts) != 2 or not parts[0].isdigit():
        await update.message.reply_text(
            "❌ সঠিক ফরম্যাটে পাঠান (যেমন: <code>123456789 5</code> অথবা <code>123456789 default</code>)",
            parse_mode=ParseMode.HTML,
            reply_markup=get_cancel_keyboard(user_id)
        )
        return STATE_ADMIN_SET_REF_BONUS

    target_id = int(parts[0])
    val_str = parts[1].lower()
    if val_str in ("default", "reset", "off", "0"):
        new_val = None
    else:
        try:
            new_val = float(val_str)
            if new_val < 0:
                raise ValueError
        except ValueError:
            await update.message.reply_text("❌ সঠিক টাকার পরিমাণ দিন (যেমন: 4 বা 5)।")
            return STATE_ADMIN_SET_REF_BONUS

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT user_id FROM users WHERE user_id = ?", (target_id,))
        if not cursor.fetchone():
            await update.message.reply_text("❌ এই আইডির কোনো ইউজার পাওয়া যায়নি।", reply_markup=get_admin_keyboard(user_id))
            clear_user_state(context)
            return ConversationHandler.END
        cursor.execute("UPDATE users SET custom_ref_fixed = ? WHERE user_id = ?", (new_val, target_id))
        conn.commit()

    if new_val is None:
        await update.message.reply_text(
            f"✅ ইউজার <code>{target_id}</code>-এর রেফার বোনাস ডিফল্ট (%) নিয়মে রিসেট করা হয়েছে।",
            reply_markup=get_admin_keyboard(user_id),
            parse_mode=ParseMode.HTML
        )
    else:
        await update.message.reply_text(
            f"✅ ইউজার <code>{target_id}</code>-এর জন্য প্রতি এপ্রুভড রেফার কাজে <b>৳{new_val:g}</b> সেট করা হয়েছে (রিজেক্ট হলে কিছু পাবে না)।",
            reply_markup=get_admin_keyboard(user_id),
            parse_mode=ParseMode.HTML
        )
    clear_user_state(context)
    return ConversationHandler.END

async def _end_silently(update: Update, context: ContextTypes.DEFAULT_TYPE):
    clear_user_state(context)
    return ConversationHandler.END

def _auto_cancel_to(handler_fn):
    async def _wrapper(update: Update, context: ContextTypes.DEFAULT_TYPE):
        clear_user_state(context)
        await handler_fn(update, context)
        return ConversationHandler.END
    return _wrapper

async def _conversation_timeout_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    clear_user_state(context)
    return ConversationHandler.END

async def cancel_operation(update: Update, context: ContextTypes.DEFAULT_TYPE):
    clear_user_state(context)
    user_id = update.effective_user.id
    await update.message.reply_text(tr(user_id, "operation_cancelled"), reply_markup=get_main_keyboard(user_id))
    return ConversationHandler.END


async def expired_callback_fallback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    if query:
        try:
            await query.answer("⚠️ সেশন শেষ হয়ে গেছে! অনুগ্রহ করে আবার মেনু থেকে বাটন চাপুন।", show_alert=True)
        except Exception:
            pass

async def global_error_handler(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Global error handler to prevent internal error notices to users"""
    logger.error("Exception while handling an update:", exc_info=context.error)

# -----------------------------------------------------------------------------
# MAIN APPLICATION SETUP
# -----------------------------------------------------------------------------
def main():
    logger.info("Initializing bot database and environment...")
    init_db()

    token = BOT_TOKEN
    if not token:
        raise RuntimeError("BOT_TOKEN is not set. Add your NEW BotFather token as the BOT_TOKEN environment variable.")
    app = ApplicationBuilder().token(token).post_init(_on_startup).build()

    app.add_error_handler(global_error_handler)

    # Regex filters for clean English buttons
    btn_task_re = r"^Gmail Work$"
    btn_wallet_re = r"^Wallet$"
    btn_withdraw_re = r"^Withdraw$"
    btn_ref_re = r"^My Referrals$"
    btn_leaderboard_re = r"^Leaderboard$"
    btn_support_re = r"^Support$"
    btn_lang_re = r"^Language$"
    btn_admin_re = r"^Admin Panel$"
    btn_main_menu_re = r"^Main Menu$"
    btn_cancel_re = r"^Cancel$"
    btn_done_re = r"^Submit Gmail$"

    btn_gmail_re = r"^New Gmail Work$"
    btn_bulk_gmail_re = r"^Bulk Gmail Work$"
    btn_new_single_re = r"^Single New Gmail$"
    btn_new_bulk_re = r"^Bulk New Gmail$"
    btn_old_gmail_re = r"^Old Gmail Work$"
    btn_old_single_re = r"^Single Old Gmail$"
    btn_old_bulk_re = r"^Bulk Old Gmail$"
    btn_back_gmail_work_re = r"^Back to Gmail Work$"

    btn_bkash_re = r"^Bkash(?: -> .*)?$"
    btn_nagad_re = r"^Nagad(?: -> .*)?$"
    btn_usdt_re = r"^USDT \(BEP-20\)$"
    btn_binance_uid_re = r"^Binance UID$"

    btn_adm_pending_re = r"^Pending Tasks$"
    btn_adm_held_re = r"^Held Tasks$"
    btn_adm_search_re = r"^Task Search$"
    btn_adm_export_re = r"^User Export$"
    btn_adm_pwd_re = r"^Password Manager$"
    btn_adm_price_re = r"^Price Manager$"
    btn_adm_wd_re = r"^Withdraw Requests$"
    btn_adm_sys_re = r"^System Settings$"
    btn_adm_stats_re = r"^Live Stats$"
    btn_adm_bcast_re = r"^Broadcast$"
    btn_adm_backup_re = r"^Backup DB$"
    btn_adm_restore_re = r"^Restore DB$"
    btn_adm_old_cats_re = r"^Old Gmail Categories$"

    btn_adm_bulk_app_re = r"^Approve All$"
    btn_adm_bulk_rej_re = r"^Reject All$"
    btn_adm_excel_re = r"^Spreadsheet UI$"
    btn_adm_user_mgmt_re = r"^User Management$"
    btn_adm_admin_mgmt_re = r"^Admin Control$"

    all_menu_buttons_re = (
        r"^(Gmail Work|New Gmail Work|Bulk Gmail Work|Single New Gmail|Bulk New Gmail|Old Gmail Work|"
        r"Single Old Gmail|Bulk Old Gmail|Back to Gmail Work|"
        r"Wallet|Withdraw|My Referrals|Leaderboard|Support|Language|Admin Panel|Main Menu|Cancel|"
        r"Pending Tasks|Held Tasks|Approve All|Reject All|Password Manager|Price Manager|"
        r"Task Search|User Export|Spreadsheet UI|Withdraw Requests|User Management|"
        r"Admin Control|System Settings|Live Stats|Broadcast|Backup DB|Old Gmail Categories|Restore DB)$"
    )
    not_menu_text = filters.TEXT & ~filters.COMMAND & ~filters.Regex(all_menu_buttons_re)

    common_nav_fallbacks = [
        CommandHandler("start", start_command),
        MessageHandler(filters.Regex(btn_main_menu_re), start_command),
        MessageHandler(filters.Regex(btn_cancel_re), cancel_operation),
        MessageHandler(filters.Regex(btn_task_re), handle_task_menu),
        MessageHandler(filters.Regex(btn_gmail_re), handle_new_gmail_submenu),
        MessageHandler(filters.Regex(btn_new_single_re), handle_gmail_menu),
        MessageHandler(filters.Regex(btn_new_bulk_re), handle_bulk_gmail_menu),
        MessageHandler(filters.Regex(btn_bulk_gmail_re), handle_bulk_gmail_menu),
        MessageHandler(filters.Regex(btn_old_gmail_re), handle_old_gmail_menu),
        MessageHandler(filters.Regex(btn_old_single_re), handle_old_single_or_bulk_start),
        MessageHandler(filters.Regex(btn_old_bulk_re), handle_old_single_or_bulk_start),
        MessageHandler(filters.Regex(btn_back_gmail_work_re), handle_task_menu),
        MessageHandler(filters.Regex(btn_wallet_re), _auto_cancel_to(handle_wallet_menu)),
        MessageHandler(filters.Regex(btn_ref_re), _auto_cancel_to(handle_referrals)),
        MessageHandler(filters.Regex(btn_leaderboard_re), _auto_cancel_to(handle_leaderboard)),
        MessageHandler(filters.Regex(btn_support_re), _auto_cancel_to(handle_support)),
        MessageHandler(filters.Regex(btn_lang_re), _auto_cancel_to(handle_language)),
        MessageHandler(filters.Regex(btn_admin_re), _auto_cancel_to(handle_admin_panel)),
        MessageHandler(filters.Regex(all_menu_buttons_re), _end_silently),
    ]

    # --- Task Conversation Handler (Group 0) ---
    task_conv = ConversationHandler(
        entry_points=[
            MessageHandler(filters.Regex(btn_task_re), handle_task_menu),
            MessageHandler(filters.Regex(btn_gmail_re), handle_new_gmail_submenu),
            MessageHandler(filters.Regex(btn_new_single_re), handle_gmail_menu),
            MessageHandler(filters.Regex(btn_new_bulk_re), handle_bulk_gmail_menu),
            MessageHandler(filters.Regex(btn_bulk_gmail_re), handle_bulk_gmail_menu),
            MessageHandler(filters.Regex(btn_old_gmail_re), handle_old_gmail_menu),
            MessageHandler(filters.Regex(btn_old_single_re), handle_old_single_or_bulk_start),
            MessageHandler(filters.Regex(btn_old_bulk_re), handle_old_single_or_bulk_start),
            MessageHandler(filters.Regex(btn_back_gmail_work_re), handle_task_menu),
        ],
        states={
            STATE_TASK_CATEGORY: [
                MessageHandler(filters.Regex(btn_gmail_re), handle_new_gmail_submenu),
                MessageHandler(filters.Regex(btn_new_single_re), handle_gmail_menu),
                MessageHandler(filters.Regex(btn_new_bulk_re), handle_bulk_gmail_menu),
                MessageHandler(filters.Regex(btn_bulk_gmail_re), handle_bulk_gmail_menu),
                MessageHandler(filters.Regex(btn_old_gmail_re), handle_old_gmail_menu),
            ],
            STATE_TASK_NEW_MENU: [
                MessageHandler(filters.Regex(btn_new_single_re), handle_gmail_menu),
                MessageHandler(filters.Regex(btn_new_bulk_re), handle_bulk_gmail_menu),
                MessageHandler(filters.Regex(btn_bulk_gmail_re), handle_bulk_gmail_menu),
                MessageHandler(filters.Regex(btn_back_gmail_work_re), handle_task_menu),
                MessageHandler(filters.Regex(btn_old_gmail_re), handle_old_gmail_menu),
            ],
            STATE_TASK_OLD_MENU: [
                MessageHandler(filters.Regex(btn_old_single_re), handle_old_single_or_bulk_start),
                MessageHandler(filters.Regex(btn_old_bulk_re), handle_old_single_or_bulk_start),
                MessageHandler(filters.Regex(btn_back_gmail_work_re), handle_task_menu),
                MessageHandler(filters.Regex(btn_gmail_re), handle_gmail_menu),
                MessageHandler(filters.Regex(btn_bulk_gmail_re), handle_bulk_gmail_menu),
            ],
            STATE_TASK_OLD_CAT: [
                CallbackQueryHandler(old_category_selected_callback, pattern=r"^oldcat_"),
                MessageHandler(filters.Regex(btn_old_single_re), handle_old_single_or_bulk_start),
                MessageHandler(filters.Regex(btn_old_bulk_re), handle_old_single_or_bulk_start),
                MessageHandler(filters.Regex(btn_back_gmail_work_re), handle_task_menu),
                MessageHandler(filters.Regex(btn_gmail_re), handle_gmail_menu),
                MessageHandler(filters.Regex(btn_bulk_gmail_re), handle_bulk_gmail_menu),
                MessageHandler(filters.Regex(btn_old_gmail_re), handle_old_gmail_menu),
            ],
            STATE_TASK_OLD_SINGLE_INPUT: [
                CallbackQueryHandler(old_single_done_callback, pattern=r"^old_single_done$"),
                CallbackQueryHandler(old_gmail_cancel_callback, pattern=r"^old_gmail_cancel$"),
                MessageHandler(filters.Regex(btn_cancel_re), cancel_operation),
                MessageHandler(filters.Regex(btn_main_menu_re), start_command),
                MessageHandler(not_menu_text, process_old_single_submission),
            ],
            STATE_TASK_OLD_BULK_INPUT: [
                CallbackQueryHandler(old_bulk_done_callback, pattern=r"^old_bulk_done$"),
                CallbackQueryHandler(old_gmail_cancel_callback, pattern=r"^old_gmail_cancel$"),
                MessageHandler(filters.Regex(btn_cancel_re), cancel_operation),
                MessageHandler(filters.Regex(btn_main_menu_re), start_command),
                MessageHandler(not_menu_text, process_old_bulk_submission),
            ],
            STATE_TASK_GMAIL_INPUT: [
                CallbackQueryHandler(gmail_done_callback, pattern=r"^gmail_done$"),
                CallbackQueryHandler(gmail_cancel_callback, pattern=r"^gmail_cancel$"),
                MessageHandler(filters.Regex(btn_cancel_re), cancel_operation),
                MessageHandler(filters.Regex(btn_main_menu_re), start_command),
                MessageHandler(not_menu_text, process_gmail_submission),
            ],
            STATE_TASK_BULK_GMAIL_INPUT: [
                CallbackQueryHandler(bulk_gmail_done_callback, pattern=r"^bulk_gmail_done$"),
                CallbackQueryHandler(bulk_gmail_cancel_callback, pattern=r"^bulk_gmail_cancel$"),
                MessageHandler(filters.Regex(btn_cancel_re), cancel_operation),
                MessageHandler(filters.Regex(btn_main_menu_re), start_command),
                MessageHandler(not_menu_text, process_bulk_gmail_submission),
            ],
            ConversationHandler.TIMEOUT: [
                MessageHandler(filters.ALL, _conversation_timeout_handler),
                CallbackQueryHandler(_conversation_timeout_handler),
            ],
        },
        fallbacks=common_nav_fallbacks,
        conversation_timeout=300,
        per_message=False,
        allow_reentry=True
    )
    app.add_handler(task_conv, group=0)

    # --- Main Navigation Handlers (Group 0, after task_conv) ---
    app.add_handler(CommandHandler("start", start_command), group=0)
    app.add_handler(MessageHandler(filters.Regex(btn_wallet_re), handle_wallet_menu), group=0)
    app.add_handler(MessageHandler(filters.Regex(btn_ref_re), handle_referrals), group=0)
    app.add_handler(MessageHandler(filters.Regex(btn_leaderboard_re), handle_leaderboard), group=0)
    app.add_handler(MessageHandler(filters.Regex(btn_support_re), handle_support), group=0)
    app.add_handler(MessageHandler(filters.Regex(btn_lang_re), handle_language), group=0)
    app.add_handler(MessageHandler(filters.Regex(btn_admin_re), handle_admin_panel), group=0)
    app.add_handler(MessageHandler(filters.Regex(btn_main_menu_re), start_command), group=0)

    # --- Admin ReplyKeyboard Buttons (Group 0) ---
    app.add_handler(MessageHandler(filters.Regex(btn_adm_pending_re), admin_pending_tasks), group=0)
    app.add_handler(MessageHandler(filters.Regex(btn_adm_held_re), admin_held_tasks), group=0)
    app.add_handler(MessageHandler(filters.Regex(btn_adm_pwd_re), admin_password_manager), group=0)
    app.add_handler(MessageHandler(filters.Regex(btn_adm_price_re), admin_price_manager), group=0)
    app.add_handler(MessageHandler(filters.Regex(btn_adm_wd_re), admin_withdraw_requests), group=0)
    app.add_handler(MessageHandler(filters.Regex(btn_adm_sys_re), admin_system_settings), group=0)
    app.add_handler(MessageHandler(filters.Regex(btn_adm_stats_re), admin_live_stats), group=0)
    app.add_handler(MessageHandler(filters.Regex(btn_adm_excel_re), admin_excel_export_menu), group=0)
    app.add_handler(MessageHandler(filters.Regex(btn_adm_user_mgmt_re), admin_user_management), group=0)
    app.add_handler(MessageHandler(filters.Regex(btn_adm_admin_mgmt_re), admin_control_panel), group=0)
    app.add_handler(MessageHandler(filters.Regex(btn_adm_backup_re), admin_backup_db), group=0)
    app.add_handler(CommandHandler("backup", admin_backup_db), group=0)
    app.add_handler(MessageHandler(filters.Regex(btn_adm_old_cats_re), admin_old_categories_menu), group=0)

    # --- Withdraw Conversation Handler (Group 1) ---
    withdraw_conv = ConversationHandler(
        entry_points=[
            MessageHandler(filters.Regex(btn_withdraw_re), handle_withdraw_start),
        ],
        states={
            STATE_WD_METHOD: [
                MessageHandler(filters.Regex(btn_bkash_re), handle_withdraw_method),
                MessageHandler(filters.Regex(btn_nagad_re), handle_withdraw_method),
                MessageHandler(filters.Regex(btn_usdt_re), handle_withdraw_method),
                MessageHandler(filters.Regex(btn_binance_uid_re), handle_withdraw_method),
            ],
            STATE_WD_NUMBER: [
                MessageHandler(filters.Regex(btn_cancel_re), cancel_operation),
                MessageHandler(filters.Regex(btn_main_menu_re), start_command),
                MessageHandler(not_menu_text, handle_withdraw_number),
            ],
            STATE_WD_AMOUNT: [
                MessageHandler(filters.Regex(btn_cancel_re), cancel_operation),
                MessageHandler(filters.Regex(btn_main_menu_re), start_command),
                MessageHandler(not_menu_text, handle_withdraw_amount),
            ],
            ConversationHandler.TIMEOUT: [
                MessageHandler(filters.ALL, _conversation_timeout_handler),
            ],
        },
        fallbacks=[
            CommandHandler("start", _end_silently),
            MessageHandler(filters.Regex(btn_cancel_re), cancel_operation),
            MessageHandler(filters.Regex(all_menu_buttons_re), _end_silently),
        ],
        conversation_timeout=300,
        per_message=False,
        allow_reentry=True
    )
    app.add_handler(withdraw_conv, group=1)

    # --- Admin Input Conversation Handler (Group 2) ---
    admin_text_conv = ConversationHandler(
        entry_points=[
            MessageHandler(filters.Regex(btn_adm_search_re), admin_search_prompt),
            MessageHandler(filters.Regex(btn_adm_export_re), admin_export_prompt),
            MessageHandler(filters.Regex(btn_adm_bcast_re), admin_broadcast_prompt),
            MessageHandler(filters.Regex(btn_adm_bulk_app_re), admin_bulk_approve_prompt),
            MessageHandler(filters.Regex(btn_adm_bulk_rej_re), admin_bulk_reject_prompt),
            MessageHandler(filters.Regex(btn_adm_restore_re), admin_restore_prompt),
            CommandHandler("restore", admin_restore_prompt),
            CallbackQueryHandler(
                admin_callback_handler,
                pattern="^(adm_usr_ban|adm_usr_unban|adm_usr_add_bal|adm_usr_sub_bal|adm_usr_search|adm_usr_set_ref|adm_add_subadmin|adm_rem_subadmin|adm_edit_(?!force_channels_list)|adm_fj_add|adm_oldcat_add|adm_rej_custom_)"
            ),
        ],
        states={
            STATE_ADMIN_SEARCH_QUERY: [
                MessageHandler(not_menu_text, process_admin_text),
            ],
            STATE_ADMIN_EXPORT_USER: [
                MessageHandler(not_menu_text, process_admin_text),
            ],
            STATE_ADMIN_BROADCAST: [
                MessageHandler((filters.ALL & ~filters.COMMAND & ~filters.Regex(all_menu_buttons_re)), process_admin_broadcast_media),
            ],
            STATE_ADMIN_REJECT_REASON: [
                MessageHandler(not_menu_text, process_admin_reject_reason),
            ],
            STATE_ADMIN_SETTING_VAL: [
                MessageHandler(not_menu_text, process_admin_text),
            ],
            STATE_ADMIN_BULK_APPROVE: [
                MessageHandler(not_menu_text, process_bulk_approve),
            ],
            STATE_ADMIN_BULK_REJECT: [
                MessageHandler(not_menu_text, process_bulk_reject),
            ],
            STATE_ADMIN_ADD_SUBADMIN: [
                MessageHandler(not_menu_text, process_admin_add_subadmin),
            ],
            STATE_ADMIN_REM_SUBADMIN: [
                MessageHandler(not_menu_text, process_admin_rem_subadmin),
            ],
            STATE_ADMIN_BAN_USER: [
                MessageHandler(not_menu_text, process_admin_ban_user),
            ],
            STATE_ADMIN_UNBAN_USER: [
                MessageHandler(not_menu_text, process_admin_unban_user),
            ],
            STATE_ADMIN_ADD_BAL: [
                MessageHandler(not_menu_text, process_admin_add_bal),
            ],
            STATE_ADMIN_SUB_BAL: [
                MessageHandler(not_menu_text, process_admin_sub_bal),
            ],
            STATE_ADMIN_USER_SEARCH: [
                MessageHandler(not_menu_text, process_admin_user_search),
            ],
            STATE_ADMIN_FORCE_CHANNELS_EDIT: [
                MessageHandler(not_menu_text, process_admin_force_channels_edit),
            ],
            STATE_ADMIN_SET_REF_BONUS: [
                MessageHandler(not_menu_text, process_admin_set_ref_bonus),
            ],
            STATE_ADMIN_RESTORE_DB: [
                MessageHandler(filters.Document.ALL, process_admin_restore_db),
                MessageHandler(not_menu_text, process_admin_restore_db),
            ],
            STATE_ADMIN_WELCOME_BANNER: [
                MessageHandler(filters.PHOTO, process_admin_welcome_banner),
                MessageHandler(not_menu_text, process_admin_welcome_banner),
            ],
            STATE_ADMIN_OLD_CAT_ADD: [
                MessageHandler(not_menu_text, process_admin_old_cat_add),
            ],
            ConversationHandler.TIMEOUT: [
                MessageHandler(filters.ALL, _conversation_timeout_handler),
            ],
        },
        fallbacks=[
            CommandHandler("start", _end_silently),
            MessageHandler(filters.Regex(btn_cancel_re), cancel_operation),
            MessageHandler(filters.Regex(all_menu_buttons_re), _end_silently),
        ],
        conversation_timeout=300,
        per_message=False,
        allow_reentry=True
    )
    app.add_handler(admin_text_conv, group=2)

    # --- Global Callback Query Handlers ---
    app.add_handler(CallbackQueryHandler(language_callback, pattern="^(lang_bn|lang_en)$"))
    app.add_handler(CallbackQueryHandler(check_force_join_cb, pattern="^check_force_join_cb$"))
    app.add_handler(CallbackQueryHandler(admin_excel_export_download_cb, pattern="^adm_xlsx_"))
    app.add_handler(CallbackQueryHandler(admin_excel_clear_cb, pattern="^adm_cls_"))
    app.add_handler(CallbackQueryHandler(
        admin_callback_handler,
        pattern="^(adm_app_sub_|adm_hold_sub_|adm_rej_prompt_|adm_rej_act_|adm_app_wd_|adm_rej_wd_|adm_back_to_admin|adm_toggle_|adm_edit_force_channels_list|adm_fj_del_|adm_fj_clear|adm_oldcat_del_|adm_usr_banned_list|adm_usr_reset_balances)"
    ))

    app.add_handler(CallbackQueryHandler(expired_callback_fallback), group=0)

    logger.info("Gmail Work BD Bot successfully started!")
    app.run_polling()

if __name__ == "__main__":
    main()