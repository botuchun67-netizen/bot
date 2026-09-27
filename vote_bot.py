import asyncio
import logging
import os
import re
import sqlite3
import sys
from datetime import datetime, timedelta, timezone
from functools import wraps
from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from telegram import InlineKeyboardButton as Btn
from telegram import InlineKeyboardMarkup as Kb
from telegram import Update
from telegram.error import BadRequest, Conflict, NetworkError, RetryAfter
from telegram.ext import (Application, ApplicationBuilder, ApplicationHandlerStop,
                          CallbackQueryHandler, CommandHandler, ContextTypes,
                          MessageHandler, TypeHandler, filters)

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("vote_bot")

# ===================== SOZLAMALAR =====================
TOKEN = os.getenv("BOT_TOKEN", "BOT_TOKEN_SHU_YERGA")
SUPER_ADMIN = 8355611778  # asosiy admin: hech kim uni o'chira olmaydi, faqat u boshqa adminlar qo'sha oladi
TZ = timezone(timedelta(hours=5), "Toshkent")  # O'zbekiston: UTC+5, yozgi vaqt yo'q

DATA_DIR = os.getenv("RAILWAY_VOLUME_MOUNT_PATH") or os.getenv("DATA_DIR", ".")
os.makedirs(DATA_DIR, exist_ok=True)
DB = os.path.join(DATA_DIR, "council.db")  # yangi versiya: yangi baza fayli

PUBLIC_SHOWS_VOTERS = False  # oddiy foydalanuvchilarga to'liq (ovoz berganlar bilan) Excel yuborilsinmi

DEFAULT_CATEGORIES = ["💻 IT", "📚 Ta'lim", "⚽ Sport", "🎬 Media",
                      "🎭 Madaniyat", "🌿 Ekologiya", "⚖️ Ombudsman"]
DEFAULT_CANDIDATES = [
    ["Elmurod Saliybayev", "Ruzimbetov Sanjar", "Saliybayeva Zarina"],
    ["Chorshanbayev Sarvar", "Gaipnazarov Anvar", "Yusupova Charos"],
    ["Ayitbayev O'ktam", "Kamarov Bunyod", "Mirzayev Shohrux", "Qamaraddinov Jahongir"],
    ["Jumaboyev Farxod", "Ramatullayev Oʻlmasbek", "Saliybayeva Shamsiya", "Xamrayev Umidjon"],
    ["Ulugʻbekova Sohiba", "Kenjayev Xurshid", "Toʻrabayeva Nodira"],
    ["Abdullayeva Shahnoza", "Jangirova Elnora"],
    ["Mirzayev Mirzohid", "Sapayev Behruz"],
]

WELCOME_TEXT = """👋 Assalomu alaykum!

Siz 70-maktab O'quvchilar Kengashi rasmiy botidasiz.

Bu yerda siz:
🗳 Yo'nalish sardorlariga ovoz bera olasiz
📞 Yo'nalish sardorlari va maktab direktori bilan bog'lana olasiz
📝 Ariza yoki murojaat yubora olasiz
📢 Maktab e'lonlari va yangiliklaridan xabardor bo'lib turasiz

Quyidagi menyudan kerakli bo'limni tanlang 👇"""
# ======================================================

BACK_MENU = [Btn("⬅️ Bosh menyu", callback_data="menu")]
PANEL_BACK = [Btn("⬅️ Admin panel", callback_data="a:panel")]
CANCEL = [Btn("❌ Bekor qilish", callback_data="a:panel")]
CAP_PUB = "📊 Natijalar (Excel)"


# ---------- Baza ----------

_con = None
_end = None


def open_db():
    global _con
    _con = sqlite3.connect(DB, check_same_thread=False, isolation_level=None)
    try:
        _con.execute("PRAGMA journal_mode=WAL")
        _con.execute("PRAGMA synchronous=NORMAL")
    except sqlite3.DatabaseError:
        log.warning("WAL yoqilmadi, oddiy rejimda davom etadi")


def ex(sql, args=()):
    return _con.execute(sql, args)


def one(sql, args=()):
    return _con.execute(sql, args).fetchone()


def allr(sql, args=()):
    return _con.execute(sql, args).fetchall()


def init_db():
    _con.executescript("""
    CREATE TABLE IF NOT EXISTS categories(
        id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, pos INTEGER NOT NULL);
    CREATE TABLE IF NOT EXISTS candidates(
        id INTEGER PRIMARY KEY AUTOINCREMENT, cat_id INTEGER NOT NULL, name TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS votes(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        voter_key TEXT UNIQUE NOT NULL, voter_name TEXT NOT NULL, cand_id INTEGER NOT NULL,
        tg_id INTEGER, tg_username TEXT, ts TEXT);
    CREATE TABLE IF NOT EXISTS users(user_id INTEGER PRIMARY KEY);
    CREATE TABLE IF NOT EXISTS settings(k TEXT PRIMARY KEY, v TEXT);
    CREATE TABLE IF NOT EXISTS admins(user_id INTEGER PRIMARY KEY, ts TEXT);
    CREATE TABLE IF NOT EXISTS blocked(user_id INTEGER PRIMARY KEY, ts TEXT);
    CREATE TABLE IF NOT EXISTS arizas(
        id INTEGER PRIMARY KEY AUTOINCREMENT, tg_id INTEGER, tg_username TEXT, tg_name TEXT,
        cat_id INTEGER, text TEXT, file_id TEXT, file_type TEXT, status TEXT DEFAULT 'new', ts TEXT);
    CREATE TABLE IF NOT EXISTS broadcasts(
        id INTEGER PRIMARY KEY AUTOINCREMENT, text TEXT, file_id TEXT, file_type TEXT,
        scheduled_ts TEXT, sent INTEGER DEFAULT 0, ts TEXT);
    CREATE TABLE IF NOT EXISTS audit_log(
        id INTEGER PRIMARY KEY AUTOINCREMENT, ts TEXT, admin_id INTEGER, action TEXT);
    CREATE INDEX IF NOT EXISTS idx_votes_cand ON votes(cand_id);
    CREATE INDEX IF NOT EXISTS idx_cand_cat ON candidates(cat_id);
    """)
    if one("SELECT COUNT(*) FROM categories")[0] == 0:
        for pos, name in enumerate(DEFAULT_CATEGORIES, 1):
            ex("INSERT INTO categories(name, pos) VALUES(?,?)", (name, pos))
        cats = {name: cid for cid, name in allr("SELECT id, name FROM categories")}
        for cname, names in zip(DEFAULT_CATEGORIES, DEFAULT_CANDIDATES):
            cid = cats[cname]
            for n in names:
                ex("INSERT INTO candidates(cat_id, name) VALUES(?,?)", (cid, n))
        log.info("Standart yo'nalish va nomzodlar qo'shildi")


def get_categories():
    return allr("SELECT id, name FROM categories ORDER BY pos, id")


def cat_name(cid):
    r = one("SELECT name FROM categories WHERE id=?", (cid,))
    return r[0] if r else "?"


def load_end():
    global _end
    r = one("SELECT v FROM settings WHERE k='end'")
    _end = datetime.fromisoformat(r[0]) if r else None


def get_end():
    return _end


def save_end(dt):
    global _end
    ex("INSERT OR REPLACE INTO settings VALUES('end', ?)", (dt.isoformat(),))
    _end = dt


def is_ended():
    return _end is not None and datetime.now(TZ) >= _end


# ---------- Aloqa (sardor / direktor) ----------

def get_contact(key):
    r = one("SELECT v FROM settings WHERE k=?", (key,))
    if not r:
        return None
    parts = (r[0].split("|") + ["", "", "", ""])[:4]
    name, username, phone, tgid = parts
    return {"name": name or None, "username": username or None, "phone": phone or None,
            "tgid": int(tgid) if tgid.strip().lstrip("-").isdigit() else None}


def set_contact(key, name, username, phone, tgid):
    v = "|".join([name or "", username or "", phone or "", str(tgid) if tgid else ""])
    ex("INSERT OR REPLACE INTO settings VALUES(?,?)", (key, v))


def contact_buttons(c):
    row = []
    if c and c.get("username"):
        u = c["username"].lstrip("@")
        row.append(Btn(f"💬 @{u}", url=f"https://t.me/{u}"))
    if c and c.get("phone"):
        row.append(Btn(f"📞 {c['phone']}", url=f"tel:{c['phone']}"))
    return [row] if row else []


def contact_text(label, c):
    if not c or not c.get("name"):
        return f"{label}: hali belgilanmagan"
    lines = [f"{label}: {c['name']}"]
    if c.get("username"):
        lines.append(f"💬 @{c['username'].lstrip('@')}")
    if c.get("phone"):
        lines.append(f"📞 {c['phone']}")
    return "\n".join(lines)


# ---------- Admin / blok ----------

def is_admin(uid):
    return uid == SUPER_ADMIN or one("SELECT 1 FROM admins WHERE user_id=?", (uid,)) is not None


def is_blocked(uid):
    return one("SELECT 1 FROM blocked WHERE user_id=?", (uid,)) is not None


def all_admin_ids():
    return [SUPER_ADMIN] + [r[0] for r in allr("SELECT user_id FROM admins")]


def log_action(admin_id, text):
    ex("INSERT INTO audit_log(ts, admin_id, action) VALUES(?,?,?)",
       (datetime.now(TZ).isoformat(), admin_id, text))


def admin_only(fn):
    @wraps(fn)
    async def wrapper(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        if not is_admin(update.effective_user.id):
            if update.callback_query:
                await ans(update.callback_query, "Ruxsat yo'q", True)
            return
        return await fn(update, ctx)
    return wrapper


def super_admin_only(fn):
    @wraps(fn)
    async def wrapper(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        if update.effective_user.id != SUPER_ADMIN:
            if update.callback_query:
                await ans(update.callback_query, "Faqat asosiy admin uchun", True)
            return
        return await fn(update, ctx)
    return wrapper


# ---------- Ism-familiya (ovoz beruvchi uchun) ----------

WORD = re.compile(r"[^\W\d_]+(?:['\-][^\W\d_]+)*")


def clean_name(raw):
    s = raw.strip()
    for ch in "ʻʼ’‘`´":
        s = s.replace(ch, "'")
    words = s.split()
    if not 2 <= len(words) <= 5 or len(s) > 60:
        return None
    if not all(len(w) >= 2 and WORD.fullmatch(w) for w in words):
        return None
    display = " ".join(w[:1].upper() + w[1:] for w in words)
    key = " ".join(sorted(w.casefold() for w in words))
    return display, key


# ---------- Natijalar / Excel ----------

def results_text():
    cats = get_categories()
    rows = allr("""
        SELECT c.cat_id, c.name, COUNT(v.id) n FROM candidates c
        LEFT JOIN votes v ON v.cand_id = c.id
        GROUP BY c.id ORDER BY c.cat_id, n DESC, c.name""")
    out = ["🏆 Ovoz berish natijalari"]
    for cid, cname in cats:
        out.append(f"\n📌 {cname}")
        items = [r for r in rows if r[0] == cid]
        if not items:
            out.append("  — nomzod yo'q")
        for _, n, cnt in items:
            out.append(f"  {n} — {cnt} ovoz")
    return "\n".join(out)


def _style(ws):
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="2F5597")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for col in ws.columns:
        width = max(len(str(c.value)) if c.value is not None else 0 for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(width + 3, 60)
    ws.freeze_panes = "A2"


def make_excel(full):
    con = sqlite3.connect(DB)
    try:
        cats = {cid: name for cid, name in con.execute("SELECT id, name FROM categories ORDER BY pos, id")}
        summary = con.execute("""
            SELECT c.id, c.cat_id, c.name, COUNT(v.id) FROM candidates c
            LEFT JOIN votes v ON v.cand_id = c.id GROUP BY c.id
            ORDER BY c.cat_id, COUNT(v.id) DESC, c.name""").fetchall()
        voters = con.execute("""
            SELECT v.cand_id, v.voter_name, v.ts, v.tg_id, v.tg_username, c.cat_id, c.name
            FROM votes v JOIN candidates c ON c.id = v.cand_id ORDER BY v.id""").fetchall()
    finally:
        con.close()

    by_cand = {}
    for cid, vname, *_ in voters:
        by_cand.setdefault(cid, []).append(vname)

    wb = Workbook()
    ws = wb.active
    ws.title = "Natijalar"
    ws.append(["Yo'nalish", "O'rin", "Nomzod", "Ovozlar soni"] + (["Ovoz berganlar"] if full else []))
    prev_cat, prev_n, rank, pos = None, None, 0, 0
    for cid, cat_id, name, n in summary:
        if cat_id != prev_cat:
            prev_cat, prev_n, pos = cat_id, None, 0
        pos += 1
        if n != prev_n:
            rank, prev_n = pos, n
        row = [cats.get(cat_id, "?"), rank, name, n]
        if full:
            row.append(", ".join(by_cand.get(cid, [])))
        ws.append(row)
    ws.append([])
    ws.append(["Jami ovozlar", "", "", len(voters)])
    _style(ws)
    if full:
        for r in ws.iter_rows(min_row=2, min_col=5, max_col=5):
            r[0].alignment = Alignment(wrap_text=True, vertical="top")
        w2 = wb.create_sheet("Ovoz berganlar")
        w2.append(["№", "Ism-familiya", "Yo'nalish", "Nomzod", "Vaqt", "Telegram ID", "Username"])
        for i, (_, vname, ts, tg_id, un, cat_id, cname) in enumerate(voters, 1):
            t = datetime.fromisoformat(ts).strftime("%d.%m.%Y %H:%M:%S") if ts else ""
            w2.append([i, vname, cats.get(cat_id, "?"), cname, t, tg_id, f"@{un}" if un else ""])
        _style(w2)

    bio = BytesIO()
    wb.save(bio)
    return bio.getvalue()


async def send_results(bot, chat_id, full):
    data = await asyncio.to_thread(make_excel, full)
    name = "natijalar_toliq.xlsx" if full else "natijalar.xlsx"
    cap = "📊 To'liq natijalar (ovoz berganlar bilan)" if full else CAP_PUB
    await bot.send_document(chat_id, data, filename=name, caption=cap)


# ---------- Umumiy yordamchilar ----------

def cat_rows(prefix):
    btns = [Btn(name, callback_data=f"{prefix}:{cid}") for cid, name in get_categories()]
    return [btns[j:j + 2] for j in range(0, len(btns), 2)]


def main_kb(uid):
    rows = cat_rows("cat")
    rows.append([Btn("📊 Natijalar", callback_data="res")])
    rows.append([Btn("👥 Sardorlar", callback_data="sd"), Btn("📞 Direktorga murojaat", callback_data="dir")])
    rows.append([Btn("📝 Ariza yuborish", callback_data="ariza:start")])
    if is_admin(uid):
        rows.append([Btn("⚙️ Admin panel", callback_data="a:panel")])
    return Kb(rows)


async def ans(q, text=None, alert=False):
    try:
        await q.answer(text, show_alert=alert)
    except Exception:
        pass


async def show(q, text, kb=None):
    try:
        await q.message.edit_text(text, reply_markup=kb)
    except BadRequest as e:
        if "not modified" in str(e).lower():
            return
        await q.message.reply_text(text, reply_markup=kb)


MAIN_TEXT = "Yo'nalishni tanlang.\n⚠️ Har bir ism-familiya faqat 1 marta ovoz bera oladi."


# ---------- Foydalanuvchi: asosiy menyu ----------

async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data.clear()
    uid = update.effective_user.id
    ex("INSERT OR IGNORE INTO users VALUES(?)", (uid,))
    await update.message.reply_text(WELCOME_TEXT, reply_markup=main_kb(uid))


async def menu(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await ans(q)
    ctx.user_data.clear()
    await show(q, MAIN_TEXT, main_kb(q.from_user.id))


# ---------- Ovoz berish ----------

async def show_cat(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    if is_ended():
        await ans(q, "Ovoz berish yakunlangan.", True)
        return
    await ans(q)
    ctx.user_data.clear()
    cid = int(q.data.split(":")[1])
    title = cat_name(cid)
    cands = allr("SELECT id, name FROM candidates WHERE cat_id=? ORDER BY id", (cid,))
    contact = get_contact(f"sardor:{cid}")
    text = f"📌 {title}\n\n{contact_text('👤 Sardor', contact)}\n\n"
    text += "Ovoz bermoqchi bo'lgan nomzodni tanlang:" if cands else "Bu yo'nalishda hozircha nomzod yo'q."
    rows = contact_buttons(contact) + [[Btn(name, callback_data=f"v:{i}")] for i, name in cands]
    rows.append(BACK_MENU)
    await show(q, text, Kb(rows))


async def vote(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    if is_ended():
        await ans(q, "Ovoz berish yakunlangan.", True)
        return
    cid = int(q.data.split(":")[1])
    cand = one("SELECT name, cat_id FROM candidates WHERE id=?", (cid,))
    if not cand:
        await ans(q, "Nomzod topilmadi.", True)
        return
    await ans(q)
    ctx.user_data.clear()
    ctx.user_data.update(vstate="name", vcand=cid)
    await show(q, f"🗳 Nomzod: {cand[0]}\n\n✍️ Ism va familiyangizni yozing (masalan: Jasur Karimov).\n"
                  "⚠️ Har bir ism-familiya faqat 1 marta ovoz bera oladi, keyin o'zgartirib bo'lmaydi.",
               Kb([[Btn("❌ Bekor qilish", callback_data=f"cat:{cand[1]}")]]))


async def vote_name(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    m, ud = update.message, ctx.user_data
    cancel = Kb([[Btn("❌ Bekor qilish", callback_data="menu")]])
    if is_ended():
        ud.clear()
        await m.reply_text("Ovoz berish yakunlangan.")
        return
    parsed = clean_name(m.text)
    if not parsed:
        await m.reply_text("❗ Ism va familiyangizni to'liq, harflar bilan yozing.\nMasalan: Jasur Karimov",
                           reply_markup=cancel)
        return
    name, key = parsed
    taken = one("SELECT 1 FROM votes WHERE voter_key=?", (key,))
    cand = one("SELECT name, cat_id FROM candidates WHERE id=?", (ud.get("vcand"),))
    if taken:
        await m.reply_text("⛔ Bu ism-familiya bilan allaqachon ovoz berilgan.\n"
                           "Boshqa ism-familiya yozing yoki bekor qiling.", reply_markup=cancel)
        return
    if not cand:
        ud.clear()
        await m.reply_text("Nomzod topilmadi.", reply_markup=Kb([BACK_MENU]))
        return
    ud.update(vstate="confirm", vname=name, vkey=key)
    kb = Kb([[Btn("✅ Tasdiqlash", callback_data="vc:yes"), Btn("❌ Bekor qilish", callback_data="menu")]])
    await m.reply_text(f"👤 Ovoz beruvchi: {name}\n📌 {cat_name(cand[1])}\n🗳 Nomzod: {cand[0]}\n\n"
                       "⚠️ Ovoz qabul qilingach, o'zgartirib bo'lmaydi.", reply_markup=kb)


async def vote_confirm(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q, ud = update.callback_query, ctx.user_data
    if is_ended():
        await ans(q, "Ovoz berish yakunlangan.", True)
        return
    if ud.get("vstate") != "confirm":
        await ans(q, "Bu ovoz allaqachon hisoblangan yoki sessiya tugagan. /start bosing.", True)
        return
    cid, name, key = ud["vcand"], ud["vname"], ud["vkey"]
    ud.clear()
    if not one("SELECT 1 FROM candidates WHERE id=?", (cid,)):
        await ans(q, "Nomzod topilmadi.", True)
        return
    try:
        ex("""INSERT INTO votes(voter_key, voter_name, cand_id, tg_id, tg_username, ts)
              VALUES(?,?,?,?,?,?)""",
           (key, name, cid, q.from_user.id, q.from_user.username, datetime.now(TZ).isoformat()))
    except sqlite3.IntegrityError:
        await ans(q, "Bu ism-familiya bilan allaqachon ovoz berilgan.", True)
        return
    await ans(q, "✅ Ovozingiz qabul qilindi!", True)
    await show(q, f"✅ {name}, ovozingiz qabul qilindi.\nNatijalar ovoz berish tugagach e'lon qilinadi.",
               Kb([BACK_MENU]))


async def results(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    is_adm = is_admin(q.from_user.id)
    if not (is_ended() or is_adm):
        await ans(q, "Natijalar ovoz berish tugagach e'lon qilinadi.", True)
        return
    await ans(q)
    await show(q, results_text(), Kb([BACK_MENU]))
    await send_results(ctx.bot, q.message.chat_id, is_adm or PUBLIC_SHOWS_VOTERS)


# ---------- Sardorlar / Direktor (oddiy foydalanuvchi) ----------

async def sardorlar_menu(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await ans(q)
    rows = [[Btn(name, callback_data=f"sdv:{cid}")] for cid, name in get_categories()]
    rows.append(BACK_MENU)
    await show(q, "👥 Sardorlar\n\nYo'nalishni tanlang:", Kb(rows))


async def sardor_view(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await ans(q)
    cid = int(q.data.split(":")[1])
    contact = get_contact(f"sardor:{cid}")
    text = f"📌 {cat_name(cid)}\n\n{contact_text('👤 Sardor', contact)}"
    rows = contact_buttons(contact) + [[Btn("⬅️ Sardorlar", callback_data="sd")]]
    await show(q, text, Kb(rows))


async def director_view(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await ans(q)
    contact = get_contact("director")
    text = f"🎓 Maktab direktori\n\n{contact_text('Direktor', contact)}"
    if contact and (contact.get("username") or contact.get("phone")):
        text += "\n\n📩 Murojaat qilish uchun quyidagi tugmani bosing — chat to'g'ridan-to'g'ri ochiladi."
    rows = contact_buttons(contact) + [BACK_MENU]
    await show(q, text, Kb(rows))


# ---------- Ariza yuborish ----------

async def ariza_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await ans(q)
    ctx.user_data.clear()
    rows = cat_rows("arizacat")
    rows.append(BACK_MENU)
    await show(q, "📝 Ariza qaysi yo'nalish bo'yicha? Tanlang:", Kb(rows))


async def ariza_pick_cat(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await ans(q)
    cid = int(q.data.split(":")[1])
    ctx.user_data.clear()
    ctx.user_data.update(flow="ariza", step="text", cat=cid)
    await show(q, f"📌 {cat_name(cid)}\n\n✍️ Arizangiz matnini yozing:",
               Kb([[Btn("❌ Bekor qilish", callback_data="menu")]]))


async def ariza_attach_choice(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q, ud = update.callback_query, ctx.user_data
    if ud.get("flow") != "ariza":
        await ans(q, "Sessiya tugagan. Qaytadan boshlang.", True)
        return
    await ans(q)
    if q.data.endswith("yes"):
        ud["step"] = "attach_wait"
        await show(q, "📎 Endi rasm yoki faylni yuboring:", Kb([[Btn("❌ Bekor qilish", callback_data="menu")]]))
    else:
        await submit_ariza(ctx, q.from_user, ud, q.message.chat_id)


async def submit_ariza(ctx, user, ud, chat_id):
    cid, text = ud["cat"], ud.get("text", "")
    file_id, file_type = ud.get("file_id"), ud.get("file_type")
    cur = ex("""INSERT INTO arizas(tg_id, tg_username, tg_name, cat_id, text, file_id, file_type, status, ts)
                VALUES(?,?,?,?,?,?,?, 'new', ?)""",
             (user.id, user.username, user.full_name, cid, text, file_id, file_type,
              datetime.now(TZ).isoformat()))
    aid = cur.lastrowid
    ud.clear()
    caption = (f"📥 Yangi ariza #{aid}\n📌 {cat_name(cid)}\n"
              f"👤 {user.full_name} (@{user.username or '—'}, ID: {user.id})\n\n{text}")
    targets = set(all_admin_ids())
    sardor = get_contact(f"sardor:{cid}")
    if sardor and sardor.get("tgid"):
        targets.add(sardor["tgid"])
    for uid in targets:
        try:
            if file_type == "photo":
                await ctx.bot.send_photo(uid, file_id, caption=caption)
            elif file_type == "document":
                await ctx.bot.send_document(uid, file_id, caption=caption)
            else:
                await ctx.bot.send_message(uid, caption)
        except Exception:
            pass
    await ctx.bot.send_message(chat_id, "✅ Arizangiz qabul qilindi va tegishli shaxslarga yuborildi.",
                               reply_markup=Kb([BACK_MENU]))


# ---------- Matn va media qabul qiluvchi umumiy handlerlar ----------

ADMIN_TEXT_STATES = {"add_name", "set_end", "cat_add", "cat_rename",
                     "contact_name", "contact_username", "contact_phone", "contact_tgid",
                     "announce_time", "block_id", "unblock_id", "addadm_id", "deladm_id"}


async def on_text(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    ud = ctx.user_data
    if is_admin(uid) and ud.get("state") in ADMIN_TEXT_STATES:
        await admin_text(update, ctx)
    elif ud.get("flow") in ("ariza", "announce") and ud.get("step") == "text":
        await flow_text(update, ctx)
    elif ud.get("vstate") == "name":
        await vote_name(update, ctx)
    else:
        await update.message.reply_text("Boshlash uchun /start bosing.")


async def flow_text(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    m, ud = update.message, ctx.user_data
    if ud["flow"] == "ariza":
        ud["text"] = m.text.strip()[:1000]
        ud["step"] = "attach_choice"
        await m.reply_text("📎 Rasm yoki fayl biriktirasizmi?",
                           reply_markup=Kb([[Btn("📎 Ha", callback_data="arizaattach:yes"),
                                            Btn("➡️ Yo'q", callback_data="arizaattach:no")]]))
    else:  # announce
        ud["text"] = m.text.strip()[:2000]
        ud["step"] = "attach_choice"
        await m.reply_text("📎 Rasm yoki fayl biriktirasizmi?",
                           reply_markup=Kb([[Btn("📎 Ha", callback_data="anattach:yes"),
                                            Btn("➡️ Yo'q", callback_data="anattach:no")]]))


async def on_media(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    m, ud = update.message, ctx.user_data
    if ud.get("flow") not in ("ariza", "announce") or ud.get("step") != "attach_wait":
        return
    if m.photo:
        ud["file_id"], ud["file_type"] = m.photo[-1].file_id, "photo"
    elif m.document:
        ud["file_id"], ud["file_type"] = m.document.file_id, "document"
    else:
        await m.reply_text("Iltimos, rasm yoki fayl yuboring.")
        return
    if ud["flow"] == "ariza":
        await submit_ariza(ctx, update.effective_user, ud, m.chat_id)
    else:
        ud["step"] = "send_choice"
        await m.reply_text("Qachon yuborilsin?",
                           reply_markup=Kb([[Btn("🚀 Hozir", callback_data="ansend:now"),
                                            Btn("⏰ Rejalashtirish", callback_data="ansend:sched")]]))


# ---------- Admin: bosh panel ----------

def panel_text():
    users = one("SELECT COUNT(*) FROM users")[0]
    votes = one("SELECT COUNT(*) FROM votes")[0]
    cands = one("SELECT COUNT(*) FROM candidates")[0]
    cats = one("SELECT COUNT(*) FROM categories")[0]
    new_ar = one("SELECT COUNT(*) FROM arizas WHERE status='new'")[0]
    admins_n = one("SELECT COUNT(*) FROM admins")[0]
    blocked_n = one("SELECT COUNT(*) FROM blocked")[0]
    end = get_end()
    end_s = f"{end:%d.%m.%Y %H:%M}" if end else "belgilanmagan"
    return (f"⚙️ Admin panel\n\n👥 Foydalanuvchilar: {users}\n🗳 Ovozlar: {votes}\n"
            f"👤 Nomzodlar: {cands}\n🏷 Yo'nalishlar: {cats}\n📥 Yangi arizalar: {new_ar}\n"
            f"👮 Qo'shimcha adminlar: {admins_n}\n🚫 Bloklanganlar: {blocked_n}\n⏰ Tugash: {end_s}")


def panel_kb():
    return Kb([
        [Btn("🗳 Ovoz berish", callback_data="a:vote")],
        [Btn("🏷 Yo'nalishlar", callback_data="a:cats"), Btn("👤 Sardor/Direktor", callback_data="a:contacts")],
        [Btn("📥 Arizalar", callback_data="a:ariza"), Btn("📢 E'lonlar", callback_data="a:ann")],
        [Btn("🔐 Foydalanuvchilar", callback_data="a:users"), Btn("📜 Jurnal", callback_data="a:log")],
        [Btn("⬅️ Asosiy menyu", callback_data="menu")],
    ])


@admin_only
async def admin_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data.clear()
    await update.message.reply_text(panel_text(), reply_markup=panel_kb())


# ---------- Admin: ovoz berish bo'limi (nomzod/tugash/natija/reset) ----------

VOTE_BACK = [Btn("⬅️ Ovoz berish bo'limi", callback_data="a:vote")]


async def admin_vote_menu(q, ctx):
    kb = Kb([
        [Btn("➕ Nomzod qo'shish", callback_data="a:add"), Btn("📋 Nomzodlar", callback_data="a:list")],
        [Btn("⏰ Tugash vaqti", callback_data="a:end"), Btn("📊 Natijalar", callback_data="a:res")],
        [Btn("🧹 Ovozlarni nolga tushirish", callback_data="a:reset")],
        PANEL_BACK,
    ])
    await show(q, "🗳 Ovoz berish bo'limi", kb)


async def admin_vote_cb(sec, p, q, ctx):
    admin_id = q.from_user.id

    if sec == "add":
        await show(q, "Qaysi yo'nalishga nomzod qo'shamiz?", Kb(cat_rows("a:addcat") + [VOTE_BACK]))

    elif sec == "addcat":
        n = int(p[2])
        ctx.user_data.clear()
        ctx.user_data.update(state="add_name", cat=n)
        await show(q, f"{cat_name(n)}\n\n✍️ Nomzodning ism-familiyasini yozing:",
                   Kb([[Btn("❌ Bekor qilish", callback_data="a:vote")]]))

    elif sec == "list":
        await show(q, "Qaysi yo'nalish nomzodlari?", Kb(cat_rows("a:listcat") + [VOTE_BACK]))

    elif sec == "listcat":
        n = int(p[2])
        rows = allr("""SELECT c.id, c.name, COUNT(v.id) FROM candidates c
                       LEFT JOIN votes v ON v.cand_id = c.id WHERE c.cat_id=? GROUP BY c.id ORDER BY c.id""", (n,))
        if not rows:
            await show(q, f"{cat_name(n)}: nomzod yo'q.", Kb([VOTE_BACK]))
            return
        kb = [[Btn(f"🗑 {name} ({cnt})", callback_data=f"a:delask:{cid}")] for cid, name, cnt in rows]
        await show(q, f"{cat_name(n)}\n\nO'chirish uchun nomzodni bosing (qavsda ovozlar soni):",
                   Kb(kb + [VOTE_BACK]))

    elif sec == "delask":
        cid = int(p[2])
        r = one("SELECT name, cat_id FROM candidates WHERE id=?", (cid,))
        if not r:
            await show(q, "Nomzod topilmadi.", Kb([VOTE_BACK]))
            return
        kb = Kb([[Btn("✅ Ha, o'chirish", callback_data=f"a:del:{cid}"),
                  Btn("❌ Yo'q", callback_data=f"a:listcat:{r[1]}")]])
        await show(q, f"🗑 {r[0]} o'chirilsinmi?\nUning ovozlari ham o'chadi.", kb)

    elif sec == "del":
        cid = int(p[2])
        r = one("SELECT name, cat_id FROM candidates WHERE id=?", (cid,))
        if r:
            ex("DELETE FROM votes WHERE cand_id=?", (cid,))
            ex("DELETE FROM candidates WHERE id=?", (cid,))
            log_action(admin_id, f"Nomzod o'chirildi: {r[0]}")
        back = [Btn("⬅️ Ro'yxat", callback_data=f"a:listcat:{r[1]}")] if r else []
        await show(q, "🗑 O'chirildi", Kb([back, VOTE_BACK] if back else [VOTE_BACK]))

    elif sec == "end":
        kb = Kb([
            [Btn("+1 kun", callback_data="a:endin:1"), Btn("+3 kun", callback_data="a:endin:3"),
             Btn("+7 kun", callback_data="a:endin:7")],
            [Btn("✍️ Sana va vaqtni yozish", callback_data="a:endman")],
            [Btn("🏁 Hozir yakunlash", callback_data="a:endnow")],
            VOTE_BACK,
        ])
        await show(q, "⏰ Ovoz berish qachon tugasin? (hozirdan boshlab)", kb)

    elif sec == "endin":
        dt = datetime.now(TZ) + timedelta(days=int(p[2]))
        save_end(dt)
        schedule_finish(ctx.application)
        log_action(admin_id, f"Tugash vaqti: {dt:%d.%m.%Y %H:%M}")
        await show(q, f"✅ Tugash vaqti: {dt:%d.%m.%Y %H:%M}", Kb([VOTE_BACK]))

    elif sec == "endman":
        ctx.user_data.clear()
        ctx.user_data["state"] = "set_end"
        await show(q, "Sana va vaqtni yozing (Toshkent vaqti):\n2026-10-05 18:00",
                   Kb([[Btn("❌ Bekor qilish", callback_data="a:vote")]]))

    elif sec == "endnow":
        kb = Kb([[Btn("✅ Ha, yakunlash", callback_data="a:endyes"),
                  Btn("❌ Yo'q", callback_data="a:vote")]])
        await show(q, "⚠️ Ovoz berish hoziroq tugaydi va natijalar hammaga yuboriladi. Tasdiqlaysizmi?", kb)

    elif sec == "endyes":
        save_end(datetime.now(TZ))
        schedule_finish(ctx.application)
        log_action(admin_id, "Ovoz berish darhol yakunlandi")
        await show(q, "🏁 Ovoz berish yakunlandi. Natijalar yuborilmoqda.", Kb([VOTE_BACK]))

    elif sec == "res":
        await show(q, results_text(), Kb([VOTE_BACK]))
        await send_results(ctx.bot, q.message.chat_id, True)

    elif sec == "reset":
        n = one("SELECT COUNT(*) FROM votes")[0]
        kb = Kb([[Btn("✅ Ha, nolga tushirish", callback_data="a:resetyes"),
                  Btn("❌ Yo'q", callback_data="a:vote")]])
        await show(q, f"⚠️ Barcha {n} ta ovoz o'chiriladi va hisob 0 dan boshlanadi.\n"
                      "Nomzodlar va tugash vaqti saqlanadi. Tasdiqlaysizmi?", kb)

    elif sec == "resetyes":
        ex("DELETE FROM votes")
        log_action(admin_id, "Barcha ovozlar nolga tushirildi")
        await show(q, "✅ Barcha ovozlar o'chirildi. Hisob 0 dan boshlandi.", Kb([VOTE_BACK]))


# ---------- Admin: yo'nalishlar ----------

async def admin_cats_menu(q, ctx):
    rows = [[Btn(name, callback_data=f"a:catv:{cid}")] for cid, name in get_categories()]
    rows.append([Btn("➕ Yangi yo'nalish", callback_data="a:catadd")])
    rows.append(PANEL_BACK)
    await show(q, "🏷 Yo'nalishlar", Kb(rows))


async def admin_cats_cb(sec, p, q, ctx):
    admin_id = q.from_user.id
    CATS_BACK = [Btn("⬅️ Yo'nalishlar", callback_data="a:cats")]

    if sec == "catv":
        cid = int(p[2])
        n_cand = one("SELECT COUNT(*) FROM candidates WHERE cat_id=?", (cid,))[0]
        n_vote = one("""SELECT COUNT(*) FROM votes v JOIN candidates c ON c.id=v.cand_id
                        WHERE c.cat_id=?""", (cid,))[0]
        kb = Kb([[Btn("✏️ Nomini o'zgartirish", callback_data=f"a:catren:{cid}")],
                 [Btn("🗑 O'chirish", callback_data=f"a:catdelask:{cid}")], CATS_BACK])
        await show(q, f"📌 {cat_name(cid)}\n👤 Nomzodlar: {n_cand}\n🗳 Ovozlar: {n_vote}", kb)

    elif sec == "catadd":
        ctx.user_data.clear()
        ctx.user_data["state"] = "cat_add"
        await show(q, "✍️ Yangi yo'nalish nomini yozing:", Kb([CATS_BACK]))

    elif sec == "catren":
        cid = int(p[2])
        ctx.user_data.clear()
        ctx.user_data.update(state="cat_rename", cat=cid)
        await show(q, f"'{cat_name(cid)}' uchun yangi nom yozing:", Kb([CATS_BACK]))

    elif sec == "catdelask":
        cid = int(p[2])
        n_cand = one("SELECT COUNT(*) FROM candidates WHERE cat_id=?", (cid,))[0]
        kb = Kb([[Btn("✅ Ha, o'chirish", callback_data=f"a:catdel:{cid}"),
                  Btn("❌ Yo'q", callback_data=f"a:catv:{cid}")]])
        await show(q, f"🗑 '{cat_name(cid)}' o'chirilsinmi?\n{n_cand} ta nomzod va ularning ovozlari ham o'chadi.", kb)

    elif sec == "catdel":
        cid = int(p[2])
        name = cat_name(cid)
        cand_ids = [r[0] for r in allr("SELECT id FROM candidates WHERE cat_id=?", (cid,))]
        for c in cand_ids:
            ex("DELETE FROM votes WHERE cand_id=?", (c,))
        ex("DELETE FROM candidates WHERE cat_id=?", (cid,))
        ex("DELETE FROM categories WHERE id=?", (cid,))
        ex("DELETE FROM settings WHERE k=?", (f"sardor:{cid}",))
        log_action(admin_id, f"Yo'nalish o'chirildi: {name}")
        await show(q, f"🗑 '{name}' yo'nalishi o'chirildi.", Kb([CATS_BACK]))


# ---------- Admin: sardor / direktor kontaktlari ----------

async def admin_contacts_menu(q, ctx):
    rows = [[Btn("🎓 Direktor", callback_data="a:cv:d")]]
    rows += [[Btn(name, callback_data=f"a:cv:s:{cid}")] for cid, name in get_categories()]
    rows.append(PANEL_BACK)
    await show(q, "👤 Sardor va direktor kontaktlari", Kb(rows))


async def admin_contacts_cb(sec, p, q, ctx):
    CONTACTS_BACK = [Btn("⬅️ Kontaktlar", callback_data="a:contacts")]

    if sec == "cv":
        if p[2] == "d":
            key, label = "director", "🎓 Direktor"
            edit_cb = "a:ce:d"
        else:
            cid = int(p[3])
            key, label = f"sardor:{cid}", f"📌 {cat_name(cid)} sardori"
            edit_cb = f"a:ce:s:{cid}"
        contact = get_contact(key)
        text = contact_text(label, contact)
        if contact and contact.get("tgid"):
            text += f"\n🆔 Telegram ID: {contact['tgid']}"
        kb = Kb([[Btn("✏️ Tahrirlash", callback_data=edit_cb)], CONTACTS_BACK])
        await show(q, text, kb)

    elif sec == "ce":
        if p[2] == "d":
            key, label = "director", "Direktor"
        else:
            cid = int(p[3])
            key, label = f"sardor:{cid}", f"{cat_name(cid)} sardori"
        ctx.user_data.clear()
        ctx.user_data.update(state="contact_name", ckey=key, clabel=label)
        await show(q, f"✏️ {label}\n\n✍️ Ism-familiyasini yozing:", Kb([CONTACTS_BACK]))


# ---------- Admin: arizalar ----------

STATUS_ICON = {"new": "🆕", "reviewed": "✅"}


async def admin_ariza_menu(q, ctx):
    kb = Kb([
        [Btn("🆕 Yangi", callback_data="a:arlist:new"), Btn("✅ Ko'rilgan", callback_data="a:arlist:reviewed")],
        [Btn("🔁 Barchasi", callback_data="a:arlist:all")],
        PANEL_BACK,
    ])
    await show(q, "📥 Arizalar", kb)


async def admin_ariza_cb(sec, p, q, ctx):
    ARIZA_BACK = [Btn("⬅️ Arizalar", callback_data="a:ariza")]

    if sec == "arlist":
        f = p[2]
        if f == "all":
            rows = allr("SELECT id, cat_id, status, ts FROM arizas ORDER BY id DESC LIMIT 15")
        else:
            rows = allr("SELECT id, cat_id, status, ts FROM arizas WHERE status=? ORDER BY id DESC LIMIT 15", (f,))
        if not rows:
            await show(q, "Ariza topilmadi.", Kb([ARIZA_BACK]))
            return
        kb = [[Btn(f"{STATUS_ICON.get(st,'')} #{aid} {cat_name(cid)} — "
                   f"{datetime.fromisoformat(ts):%d.%m %H:%M}", callback_data=f"a:arv:{aid}")]
              for aid, cid, st, ts in rows]
        await show(q, "📥 Arizalar ro'yxati:", Kb(kb + [ARIZA_BACK]))

    elif sec == "arv":
        aid = int(p[2])
        r = one("""SELECT tg_name, tg_username, tg_id, cat_id, text, file_type, status, ts
                   FROM arizas WHERE id=?""", (aid,))
        if not r:
            await show(q, "Ariza topilmadi.", Kb([ARIZA_BACK]))
            return
        name, uname, tgid, cid, text, ftype, status, ts = r
        info = (f"📥 Ariza #{aid} {STATUS_ICON.get(status,'')}\n📌 {cat_name(cid)}\n"
               f"👤 {name} (@{uname or '—'}, ID: {tgid})\n"
               f"🕒 {datetime.fromisoformat(ts):%d.%m.%Y %H:%M}\n")
        if ftype:
            info += "📎 Fayl biriktirilgan\n"
        info += f"\n{text}"
        kb = []
        if status == "new":
            kb.append([Btn("✅ Ko'rib chiqilgan deb belgilash", callback_data=f"a:arok:{aid}")])
        kb.append(ARIZA_BACK)
        await show(q, info, Kb(kb))

    elif sec == "arok":
        aid = int(p[2])
        ex("UPDATE arizas SET status='reviewed' WHERE id=?", (aid,))
        log_action(q.from_user.id, f"Ariza ko'rib chiqildi: #{aid}")
        await show(q, f"✅ Ariza #{aid} ko'rib chiqilgan deb belgilandi.", Kb([ARIZA_BACK]))


# ---------- Admin: e'lonlar ----------

async def admin_ann_menu(q, ctx):
    kb = Kb([
        [Btn("📝 Yangi e'lon", callback_data="a:annew")],
        [Btn("🗓 Rejalashtirilgan", callback_data="a:ansched"), Btn("📜 Tarix", callback_data="a:anhist")],
        PANEL_BACK,
    ])
    await show(q, "📢 E'lonlar", kb)


async def admin_ann_cb(sec, p, q, ctx):
    ANN_BACK = [Btn("⬅️ E'lonlar", callback_data="a:ann")]

    if sec == "annew":
        ctx.user_data.clear()
        ctx.user_data.update(flow="announce", step="text")
        await show(q, "✍️ E'lon matnini yozing:", Kb([ANN_BACK]))

    elif sec == "ansched":
        rows = allr("SELECT id, text, scheduled_ts FROM broadcasts WHERE sent=0 ORDER BY scheduled_ts")
        if not rows:
            await show(q, "Rejalashtirilgan e'lon yo'q.", Kb([ANN_BACK]))
            return
        kb = [[Btn(f"⏰ {datetime.fromisoformat(ts):%d.%m %H:%M} — {t[:20]}",
                   callback_data=f"a:anview:{aid}")] for aid, t, ts in rows]
        await show(q, "🗓 Rejalashtirilgan e'lonlar:", Kb(kb + [ANN_BACK]))

    elif sec == "anview":
        aid = int(p[2])
        r = one("SELECT text, scheduled_ts FROM broadcasts WHERE id=?", (aid,))
        if not r:
            await show(q, "Topilmadi.", Kb([ANN_BACK]))
            return
        text, ts = r
        info = f"⏰ {datetime.fromisoformat(ts):%d.%m.%Y %H:%M}\n\n{text}"
        kb = Kb([[Btn("🗑 Bekor qilish", callback_data=f"a:ancancel:{aid}")], ANN_BACK])
        await show(q, info, kb)

    elif sec == "ancancel":
        aid = int(p[2])
        ex("DELETE FROM broadcasts WHERE id=?", (aid,))
        for job in ctx.application.job_queue.get_jobs_by_name(f"ann{aid}"):
            job.schedule_removal()
        log_action(q.from_user.id, f"Rejalashtirilgan e'lon bekor qilindi: #{aid}")
        await show(q, "🗑 Bekor qilindi.", Kb([ANN_BACK]))

    elif sec == "anhist":
        rows = allr("SELECT text, ts FROM broadcasts WHERE sent=1 ORDER BY id DESC LIMIT 10")
        if not rows:
            await show(q, "Hali e'lon yuborilmagan.", Kb([ANN_BACK]))
            return
        out = ["📜 Yuborilgan e'lonlar:"]
        for text, ts in rows:
            out.append(f"\n🕒 {datetime.fromisoformat(ts):%d.%m %H:%M}\n{text[:100]}")
        await show(q, "\n".join(out), Kb([ANN_BACK]))


async def ann_attach_choice(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q, ud = update.callback_query, ctx.user_data
    if not is_admin(q.from_user.id):
        await ans(q, "Ruxsat yo'q", True)
        return
    if ud.get("flow") != "announce":
        await ans(q, "Sessiya tugagan.", True)
        return
    await ans(q)
    if q.data.endswith("yes"):
        ud["step"] = "attach_wait"
        await show(q, "📎 Endi rasm yoki faylni yuboring:", Kb([[Btn("❌ Bekor qilish", callback_data="a:ann")]]))
    else:
        ud["step"] = "send_choice"
        await show(q, "Qachon yuborilsin?", Kb([[Btn("🚀 Hozir", callback_data="ansend:now"),
                                                 Btn("⏰ Rejalashtirish", callback_data="ansend:sched")]]))


async def ann_send_choice(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q, ud = update.callback_query, ctx.user_data
    if not is_admin(q.from_user.id):
        await ans(q, "Ruxsat yo'q", True)
        return
    if ud.get("flow") != "announce":
        await ans(q, "Sessiya tugagan.", True)
        return
    await ans(q)
    text, file_id, file_type = ud.get("text", ""), ud.get("file_id"), ud.get("file_type")
    if q.data.endswith("now"):
        ex("INSERT INTO broadcasts(text, file_id, file_type, scheduled_ts, sent, ts) VALUES(?,?,?,?,1,?)",
           (text, file_id, file_type, None, datetime.now(TZ).isoformat()))
        log_action(q.from_user.id, "E'lon darhol yuborildi")
        ud.clear()
        await show(q, "🚀 E'lon yuborilmoqda...", Kb([PANEL_BACK]))
        await broadcast_now(ctx.bot, text, file_id, file_type)
    else:
        ud["state"] = "announce_time"
        await show(q, "Sana va vaqtni yozing (Toshkent vaqti):\n2026-10-05 08:00",
                   Kb([[Btn("❌ Bekor qilish", callback_data="a:ann")]]))


async def _bsend_one(bot, uid, text, file_id, file_type):
    for _ in range(2):
        try:
            if file_type == "photo":
                await bot.send_photo(uid, file_id, caption=text)
            elif file_type == "document":
                await bot.send_document(uid, file_id, caption=text)
            else:
                await bot.send_message(uid, text)
            return
        except RetryAfter as e:
            wait = e.retry_after.total_seconds() if hasattr(e.retry_after, "total_seconds") else float(e.retry_after)
            await asyncio.sleep(wait + 1)
        except Exception:
            return


async def broadcast_now(bot, text, file_id, file_type):
    users = [r[0] for r in allr("SELECT user_id FROM users")]
    for i in range(0, len(users), 10):
        await asyncio.gather(*(_bsend_one(bot, u, text, file_id, file_type) for u in users[i:i + 10]))
        await asyncio.sleep(1)


async def send_scheduled(ctx: ContextTypes.DEFAULT_TYPE):
    aid = ctx.job.data
    row = one("SELECT text, file_id, file_type, sent FROM broadcasts WHERE id=?", (aid,))
    if not row or row[3]:
        return
    await broadcast_now(ctx.bot, row[0], row[1], row[2])
    ex("UPDATE broadcasts SET sent=1 WHERE id=?", (aid,))


def schedule_announcement(app: Application, aid, dt):
    delay = max((dt - datetime.now(TZ)).total_seconds(), 1)
    app.job_queue.run_once(send_scheduled, delay, name=f"ann{aid}", data=aid)


def schedule_pending_announcements(app: Application):
    now = datetime.now(TZ)
    for aid, sched in allr("SELECT id, scheduled_ts FROM broadcasts WHERE sent=0 AND scheduled_ts IS NOT NULL"):
        dt = datetime.fromisoformat(sched)
        delay = max((dt - now).total_seconds(), 1)
        app.job_queue.run_once(send_scheduled, delay, name=f"ann{aid}", data=aid)


# ---------- Admin: foydalanuvchilar (blok / adminlar) ----------

async def admin_users_menu(q, ctx):
    rows = [
        [Btn("🚫 Bloklash", callback_data="a:ub:block"), Btn("✅ Blokdan chiqarish", callback_data="a:ub:unblock")],
        [Btn("📋 Bloklanganlar", callback_data="a:ub:listb")],
    ]
    if q.from_user.id == SUPER_ADMIN:
        rows.append([Btn("👮 Admin qo'shish", callback_data="a:ub:addadm"),
                     Btn("🗑 Admin o'chirish", callback_data="a:ub:deladm")])
        rows.append([Btn("📋 Adminlar ro'yxati", callback_data="a:ub:listadm")])
    rows.append(PANEL_BACK)
    await show(q, "🔐 Foydalanuvchilar", Kb(rows))


async def admin_users_cb(p, q, ctx):
    act = p[2]
    USERS_BACK = [Btn("⬅️ Foydalanuvchilar", callback_data="a:users")]
    ctx.user_data.clear()

    if act == "block":
        ctx.user_data["state"] = "block_id"
        await show(q, "🚫 Bloklamoqchi bo'lgan foydalanuvchining Telegram ID sini yozing:", Kb([USERS_BACK]))
    elif act == "unblock":
        ctx.user_data["state"] = "unblock_id"
        await show(q, "✅ Blokdan chiqarmoqchi bo'lgan foydalanuvchining Telegram ID sini yozing:", Kb([USERS_BACK]))
    elif act == "listb":
        rows = allr("SELECT user_id FROM blocked")
        text = "🚫 Bloklanganlar:\n" + "\n".join(str(r[0]) for r in rows) if rows else "Bloklangan yo'q."
        await show(q, text, Kb([USERS_BACK]))
    elif act == "addadm" and q.from_user.id == SUPER_ADMIN:
        ctx.user_data["state"] = "addadm_id"
        await show(q, "👮 Admin qilmoqchi bo'lgan foydalanuvchining Telegram ID sini yozing:", Kb([USERS_BACK]))
    elif act == "deladm" and q.from_user.id == SUPER_ADMIN:
        ctx.user_data["state"] = "deladm_id"
        await show(q, "🗑 Adminlikdan olib tashlamoqchi bo'lgan Telegram ID ni yozing:", Kb([USERS_BACK]))
    elif act == "listadm" and q.from_user.id == SUPER_ADMIN:
        rows = allr("SELECT user_id FROM admins")
        text = f"👮 Asosiy admin: {SUPER_ADMIN}\n\nQo'shimcha adminlar:\n" + \
               ("\n".join(str(r[0]) for r in rows) if rows else "yo'q")
        await show(q, text, Kb([USERS_BACK]))
    else:
        await ans(q, "Ruxsat yo'q", True)


async def admin_log(q, ctx):
    rows = allr("SELECT ts, admin_id, action FROM audit_log ORDER BY id DESC LIMIT 20")
    if not rows:
        text = "Jurnal bo'sh."
    else:
        text = "📜 So'nggi amallar:\n" + "\n".join(
            f"🕒 {datetime.fromisoformat(ts):%d.%m %H:%M} (ID {aid}): {act}" for ts, aid, act in rows)
    await show(q, text, Kb([PANEL_BACK]))


# ---------- Admin: matn qabul qilish (barcha wizardlar) ----------

async def admin_text(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    m, ud = update.message, ctx.user_data
    state = ud.get("state")
    admin_id = update.effective_user.id

    if state == "add_name":
        name = " ".join(m.text.split())
        if not name or len(name) > 60:
            await m.reply_text("Ism 1–60 belgi bo'lsin. Qayta yozing:")
            return
        n = ud["cat"]
        ex("INSERT INTO candidates(cat_id, name) VALUES(?,?)", (n, name))
        log_action(admin_id, f"Nomzod qo'shildi: {name} → {cat_name(n)}")
        ud.clear()
        kb = Kb([[Btn("➕ Yana qo'shish", callback_data=f"a:addcat:{n}")], VOTE_BACK])
        await m.reply_text(f"✅ Qo'shildi: {name} → {cat_name(n)}", reply_markup=kb)

    elif state == "set_end":
        try:
            dt = datetime.strptime(m.text.strip(), "%Y-%m-%d %H:%M").replace(tzinfo=TZ)
        except ValueError:
            await m.reply_text("Format noto'g'ri. Masalan: 2026-10-05 18:00")
            return
        save_end(dt)
        schedule_finish(ctx.application)
        log_action(admin_id, f"Tugash vaqti: {dt:%d.%m.%Y %H:%M}")
        ud.clear()
        await m.reply_text(f"✅ Tugash vaqti: {dt:%d.%m.%Y %H:%M}", reply_markup=Kb([VOTE_BACK]))

    elif state == "cat_add":
        name = m.text.strip()
        if not name or len(name) > 40:
            await m.reply_text("Nom 1–40 belgi bo'lsin. Qayta yozing:")
            return
        pos = (one("SELECT MAX(pos) FROM categories")[0] or 0) + 1
        ex("INSERT INTO categories(name, pos) VALUES(?,?)", (name, pos))
        log_action(admin_id, f"Yangi yo'nalish: {name}")
        ud.clear()
        await m.reply_text(f"✅ '{name}' yo'nalishi qo'shildi.",
                           reply_markup=Kb([[Btn("⬅️ Yo'nalishlar", callback_data="a:cats")]]))

    elif state == "cat_rename":
        name = m.text.strip()
        if not name or len(name) > 40:
            await m.reply_text("Nom 1–40 belgi bo'lsin. Qayta yozing:")
            return
        old = cat_name(ud["cat"])
        ex("UPDATE categories SET name=? WHERE id=?", (name, ud["cat"]))
        log_action(admin_id, f"Yo'nalish nomi o'zgardi: {old} → {name}")
        ud.clear()
        await m.reply_text(f"✅ Yangi nom: {name}", reply_markup=Kb([[Btn("⬅️ Yo'nalishlar", callback_data="a:cats")]]))

    elif state == "contact_name":
        ud["cname"] = m.text.strip()[:60]
        ud["state"] = "contact_username"
        await m.reply_text("Username kiriting (masalan @ali_1990).\nBo'lmasa, - deb yozing:")

    elif state == "contact_username":
        t = m.text.strip()
        ud["cusername"] = "" if t == "-" else t.lstrip("@")
        ud["state"] = "contact_phone"
        await m.reply_text("Telefon raqamini kiriting (masalan +998901234567).\nBo'lmasa, - deb yozing:")

    elif state == "contact_phone":
        t = m.text.strip()
        ud["cphone"] = "" if t == "-" else t
        ud["state"] = "contact_tgid"
        await m.reply_text("Telegram ID sini kiriting (raqam — botga ariza kelishi uchun kerak).\n"
                           "Bilmasangiz, - deb yozing:")

    elif state == "contact_tgid":
        t = m.text.strip()
        tgid = int(t) if t.lstrip("-").isdigit() else None
        set_contact(ud["ckey"], ud["cname"], ud.get("cusername"), ud.get("cphone"), tgid)
        log_action(admin_id, f"Kontakt saqlandi: {ud['clabel']}")
        label = ud["clabel"]
        ud.clear()
        await m.reply_text(f"✅ {label} kontakti saqlandi.",
                           reply_markup=Kb([[Btn("⬅️ Kontaktlar", callback_data="a:contacts")]]))

    elif state == "announce_time":
        try:
            dt = datetime.strptime(m.text.strip(), "%Y-%m-%d %H:%M").replace(tzinfo=TZ)
        except ValueError:
            await m.reply_text("Format noto'g'ri. Masalan: 2026-10-05 08:00")
            return
        text, file_id, file_type = ud.get("text", ""), ud.get("file_id"), ud.get("file_type")
        cur = ex("INSERT INTO broadcasts(text, file_id, file_type, scheduled_ts, sent, ts) VALUES(?,?,?,?,0,?)",
                 (text, file_id, file_type, dt.isoformat(), datetime.now(TZ).isoformat()))
        aid = cur.lastrowid
        schedule_announcement(ctx.application, aid, dt)
        log_action(admin_id, f"E'lon rejalashtirildi: {dt:%d.%m.%Y %H:%M}")
        ud.clear()
        await m.reply_text(f"✅ E'lon {dt:%d.%m.%Y %H:%M} ga rejalashtirildi.",
                           reply_markup=Kb([[Btn("⬅️ E'lonlar", callback_data="a:ann")]]))

    elif state == "block_id":
        t = m.text.strip()
        if not t.lstrip("-").isdigit():
            await m.reply_text("Faqat raqam kiriting.")
            return
        uid = int(t)
        if uid in (SUPER_ADMIN,) or is_admin(uid):
            await m.reply_text("Adminni bloklab bo'lmaydi.")
            ud.clear()
            return
        ex("INSERT OR IGNORE INTO blocked VALUES(?,?)", (uid, datetime.now(TZ).isoformat()))
        log_action(admin_id, f"Foydalanuvchi bloklandi: {uid}")
        ud.clear()
        await m.reply_text(f"🚫 {uid} bloklandi.", reply_markup=Kb([[Btn("⬅️ Foydalanuvchilar", callback_data="a:users")]]))

    elif state == "unblock_id":
        t = m.text.strip()
        if not t.lstrip("-").isdigit():
            await m.reply_text("Faqat raqam kiriting.")
            return
        uid = int(t)
        ex("DELETE FROM blocked WHERE user_id=?", (uid,))
        log_action(admin_id, f"Blokdan chiqarildi: {uid}")
        ud.clear()
        await m.reply_text(f"✅ {uid} blokdan chiqarildi.", reply_markup=Kb([[Btn("⬅️ Foydalanuvchilar", callback_data="a:users")]]))

    elif state == "addadm_id":
        t = m.text.strip()
        if not t.lstrip("-").isdigit():
            await m.reply_text("Faqat raqam kiriting.")
            return
        uid = int(t)
        ex("INSERT OR IGNORE INTO admins VALUES(?,?)", (uid, datetime.now(TZ).isoformat()))
        log_action(admin_id, f"Yangi admin qo'shildi: {uid}")
        ud.clear()
        await m.reply_text(f"👮 {uid} endi admin.", reply_markup=Kb([[Btn("⬅️ Foydalanuvchilar", callback_data="a:users")]]))

    elif state == "deladm_id":
        t = m.text.strip()
        if not t.lstrip("-").isdigit():
            await m.reply_text("Faqat raqam kiriting.")
            return
        uid = int(t)
        if uid == SUPER_ADMIN:
            await m.reply_text("Asosiy adminni o'chirib bo'lmaydi.")
            ud.clear()
            return
        ex("DELETE FROM admins WHERE user_id=?", (uid,))
        log_action(admin_id, f"Admin o'chirildi: {uid}")
        ud.clear()
        await m.reply_text(f"🗑 {uid} adminlikdan olindi.", reply_markup=Kb([[Btn("⬅️ Foydalanuvchilar", callback_data="a:users")]]))


# ---------- Admin dispatcher ----------

@admin_only
async def admin_cb(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await ans(q)
    p = q.data.split(":")
    sec = p[1]

    if sec == "panel":
        ctx.user_data.clear()
        await show(q, panel_text(), panel_kb())
    elif sec == "vote":
        await admin_vote_menu(q, ctx)
    elif sec in ("add", "addcat", "list", "listcat", "delask", "del",
                "end", "endin", "endman", "endnow", "endyes", "res", "reset", "resetyes"):
        await admin_vote_cb(sec, p, q, ctx)
    elif sec == "cats":
        await admin_cats_menu(q, ctx)
    elif sec in ("catv", "catadd", "catren", "catdelask", "catdel"):
        await admin_cats_cb(sec, p, q, ctx)
    elif sec == "contacts":
        await admin_contacts_menu(q, ctx)
    elif sec in ("cv", "ce"):
        await admin_contacts_cb(sec, p, q, ctx)
    elif sec == "ariza":
        await admin_ariza_menu(q, ctx)
    elif sec in ("arlist", "arv", "arok"):
        await admin_ariza_cb(sec, p, q, ctx)
    elif sec == "ann":
        await admin_ann_menu(q, ctx)
    elif sec in ("annew", "ansched", "anhist", "anview", "ancancel"):
        await admin_ann_cb(sec, p, q, ctx)
    elif sec == "users":
        await admin_users_menu(q, ctx)
    elif sec == "ub":
        await admin_users_cb(p, q, ctx)
    elif sec == "log":
        await admin_log(q, ctx)


# ---------- Bloklangan foydalanuvchilarni to'xtatish ----------

async def block_guard(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    if not user or user.id == SUPER_ADMIN or not is_blocked(user.id):
        return
    if update.callback_query:
        await ans(update.callback_query, "🚫 Siz botdan foydalanish huquqidan mahrum qilingansiz.", True)
    elif update.message:
        await update.message.reply_text("🚫 Siz botdan foydalanish huquqidan mahrum qilingansiz.")
    raise ApplicationHandlerStop


# ---------- Ovoz berish yakuni ----------

async def finish(ctx: ContextTypes.DEFAULT_TYPE):
    text = results_text()
    public = await asyncio.to_thread(make_excel, PUBLIC_SHOWS_VOTERS)
    users = [r[0] for r in allr("SELECT user_id FROM users") if r[0] != SUPER_ADMIN]
    box = {"fid": None}

    async def notify(uid):
        for _ in range(2):
            try:
                await ctx.bot.send_message(uid, text)
                if box["fid"]:
                    await ctx.bot.send_document(uid, box["fid"], caption=CAP_PUB)
                else:
                    msg = await ctx.bot.send_document(uid, public, filename="natijalar.xlsx", caption=CAP_PUB)
                    box["fid"] = msg.document.file_id
                return
            except RetryAfter as e:
                w = e.retry_after.total_seconds() if hasattr(e.retry_after, "total_seconds") else float(e.retry_after)
                await asyncio.sleep(w + 1)
            except Exception:
                return

    if users:
        await notify(users[0])
        for i in range(1, len(users), 10):
            await asyncio.gather(*(notify(u) for u in users[i:i + 10]))
            await asyncio.sleep(1)
    try:
        await ctx.bot.send_message(SUPER_ADMIN, text)
        await send_results(ctx.bot, SUPER_ADMIN, True)
    except Exception:
        log.exception("Adminga natijani yuborib bo'lmadi")


def schedule_finish(app: Application):
    for job in app.job_queue.get_jobs_by_name("finish"):
        job.schedule_removal()
    end = get_end()
    if end:
        delay = max((end - datetime.now(TZ)).total_seconds(), 1)
        app.job_queue.run_once(finish, delay, name="finish")


# ---------- Xatolar, ishga tushirish ----------

async def on_error(update: object, ctx: ContextTypes.DEFAULT_TYPE):
    err = ctx.error
    if isinstance(err, Conflict):
        log.error("⚠️ Bot BOSHQA JOYDA ham ishlayapti (Conflict). Eski nusxani to'xtating "
                  "yoki BotFather'da /revoke qilib tokenni almashtiring.")
    elif isinstance(err, NetworkError):
        log.warning("Tarmoq xatosi: %s", err)
    else:
        log.error("Kutilmagan xato", exc_info=err)


async def post_init(app: Application):
    schedule_finish(app)
    schedule_pending_announcements(app)


def main():
    if TOKEN.startswith("BOT_TOKEN_SHU"):
        sys.exit("BOT_TOKEN o'rnatilmagan: Railway → Variables ga BOT_TOKEN qo'shing.")
    if (os.getenv("RAILWAY_ENVIRONMENT_NAME") or os.getenv("RAILWAY_ENVIRONMENT")) \
            and not os.getenv("RAILWAY_VOLUME_MOUNT_PATH"):
        log.warning("Volume ulanmagan: qayta deploydan keyin ma'lumotlar yo'qoladi!")
    open_db()
    init_db()
    load_end()

    app = (ApplicationBuilder().token(TOKEN)
           .concurrent_updates(True)
           .connection_pool_size(64)
           .post_init(post_init)
           .build())

    app.add_handler(TypeHandler(Update, block_guard), group=-1)

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("admin", admin_cmd))
    app.add_handler(CallbackQueryHandler(menu, pattern=r"^menu$"))
    app.add_handler(CallbackQueryHandler(show_cat, pattern=r"^cat:\d+$"))
    app.add_handler(CallbackQueryHandler(vote, pattern=r"^v:\d+$"))
    app.add_handler(CallbackQueryHandler(vote_confirm, pattern=r"^vc:yes$"))
    app.add_handler(CallbackQueryHandler(results, pattern=r"^res$"))
    app.add_handler(CallbackQueryHandler(sardorlar_menu, pattern=r"^sd$"))
    app.add_handler(CallbackQueryHandler(sardor_view, pattern=r"^sdv:\d+$"))
    app.add_handler(CallbackQueryHandler(director_view, pattern=r"^dir$"))
    app.add_handler(CallbackQueryHandler(ariza_start, pattern=r"^ariza:start$"))
    app.add_handler(CallbackQueryHandler(ariza_pick_cat, pattern=r"^arizacat:\d+$"))
    app.add_handler(CallbackQueryHandler(ariza_attach_choice, pattern=r"^arizaattach:(yes|no)$"))
    app.add_handler(CallbackQueryHandler(ann_attach_choice, pattern=r"^anattach:(yes|no)$"))
    app.add_handler(CallbackQueryHandler(ann_send_choice, pattern=r"^ansend:(now|sched)$"))
    app.add_handler(CallbackQueryHandler(admin_cb, pattern=r"^a:"))
    app.add_handler(MessageHandler(filters.PHOTO | filters.Document.ALL, on_media))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, on_text))
    app.add_error_handler(on_error)

    log.info("Bot ishga tushdi")
    app.run_polling(allowed_updates=["message", "callback_query"], drop_pending_updates=True)


if __name__ == "__main__":
    main()
