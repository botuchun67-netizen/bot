import asyncio
import logging
import os
import re
import sqlite3
import sys
from datetime import datetime, timedelta
from functools import wraps
from io import BytesIO
from zoneinfo import ZoneInfo

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from telegram import InlineKeyboardButton as Btn
from telegram import InlineKeyboardMarkup as Kb
from telegram import Update
from telegram.ext import (Application, ApplicationBuilder, CallbackQueryHandler,
                          CommandHandler, ContextTypes, MessageHandler,
                          PersistenceInput, PicklePersistence, filters)

logging.basicConfig(level=logging.INFO)
log = logging.getLogger("vote_bot")

TOKEN = os.getenv("BOT_TOKEN", "8708873894:AAFgu5CVlxpEjO8FmmOTqnUXbxeNk0aTDM4")
ADMIN_ID = 8355611778
TZ = ZoneInfo("Asia/Tashkent")

# Railway Volume ulangan bo'lsa ma'lumotlar shu papkada saqlanadi (qayta deploydan keyin ham yo'qolmaydi)
DATA_DIR = os.getenv("RAILWAY_VOLUME_MOUNT_PATH") or os.getenv("DATA_DIR", ".")
os.makedirs(DATA_DIR, exist_ok=True)
DB = os.path.join(DATA_DIR, "votes.db")
STATE_FILE = os.path.join(DATA_DIR, "state.pickle")

# Har bir ovozdan keyin bot o'zini qayta ishga tushiradi (o'chirish uchun False qiling)
RESTART_AFTER_VOTE = True
# False: oddiy foydalanuvchilarga Excelda faqat nomzodlar natijasi, kim ovoz bergani faqat adminga
PUBLIC_SHOWS_VOTERS = False

CATEGORIES = ["💻 IT", "📚 Ta'lim", "⚽ Sport", "🎬 Media",
              "🎭 Madaniyat", "🌿 Ekologiya", "⚖️ Ombudsman"]

admin = filters.User(ADMIN_ID)
BACK_MENU = [Btn("⬅️ Yo'nalishlar", callback_data="menu")]
PANEL_BACK = [Btn("⬅️ Admin panel", callback_data="a:panel")]
CANCEL = [Btn("❌ Bekor qilish", callback_data="a:panel")]


# ---------- Baza ----------

def db():
    return sqlite3.connect(DB)


def init_db():
    with db() as c:
        cols = [r[1] for r in c.execute("PRAGMA table_info(votes)")]
        if cols and "voter_key" not in cols:  # eski versiya jadvali
            c.execute("ALTER TABLE votes RENAME TO votes_old")
        c.executescript("""
        CREATE TABLE IF NOT EXISTS candidates(
            id INTEGER PRIMARY KEY AUTOINCREMENT, cat INTEGER, name TEXT, photo TEXT);
        CREATE TABLE IF NOT EXISTS votes(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            voter_key TEXT UNIQUE NOT NULL,
            voter_name TEXT NOT NULL,
            cand_id INTEGER NOT NULL,
            tg_id INTEGER, tg_username TEXT, ts TEXT);
        CREATE TABLE IF NOT EXISTS users(user_id INTEGER PRIMARY KEY);
        CREATE TABLE IF NOT EXISTS settings(k TEXT PRIMARY KEY, v TEXT);
        """)


def get_end():
    with db() as c:
        row = c.execute("SELECT v FROM settings WHERE k='end'").fetchone()
    return datetime.fromisoformat(row[0]) if row else None


def save_end(dt):
    with db() as c:
        c.execute("INSERT OR REPLACE INTO settings VALUES('end', ?)", (dt.isoformat(),))


def is_ended():
    end = get_end()
    return end is not None and datetime.now(TZ) >= end


# ---------- Ism-familiya ----------

WORD = re.compile(r"[^\W\d_]+(?:['\-][^\W\d_]+)*")


def clean_name(raw):
    """(ko'rinadigan ism, noyob kalit) yoki None. Ism-familiya tartibi farq qilmaydi."""
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


# ---------- Natijalar ----------

def results_text():
    with db() as c:
        rows = c.execute("""
            SELECT c.cat, c.name, COUNT(v.id) n FROM candidates c
            LEFT JOIN votes v ON v.cand_id = c.id
            GROUP BY c.id ORDER BY c.cat, n DESC""").fetchall()
    out = ["🏆 Ovoz berish natijalari"]
    for i, name in enumerate(CATEGORIES, 1):
        out.append(f"\n📌 {name}")
        items = [r for r in rows if r[0] == i]
        if not items:
            out.append("  — nomzod yo'q")
        for _, cname, n in items:
            out.append(f"  {cname} — {n} ovoz")
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
    """full=True: ovoz berganlar ismlari va ID lari bilan to'liq hisobot."""
    with db() as c:
        summary = c.execute("""
            SELECT c.id, c.cat, c.name, COUNT(v.id) FROM candidates c
            LEFT JOIN votes v ON v.cand_id = c.id GROUP BY c.id
            ORDER BY c.cat, COUNT(v.id) DESC, c.name""").fetchall()
        voters = c.execute("""
            SELECT v.cand_id, v.voter_name, v.ts, v.tg_id, v.tg_username, c.cat, c.name
            FROM votes v JOIN candidates c ON c.id = v.cand_id ORDER BY v.id""").fetchall()

    by_cand = {}
    for cid, vname, *_ in voters:
        by_cand.setdefault(cid, []).append(vname)

    wb = Workbook()
    ws = wb.active
    ws.title = "Natijalar"
    ws.append(["Yo'nalish", "O'rin", "Nomzod", "Ovozlar soni"] + (["Ovoz berganlar"] if full else []))
    prev_cat, prev_n, rank, pos = None, None, 0, 0
    for cid, cat, name, n in summary:
        if cat != prev_cat:
            prev_cat, prev_n, pos = cat, None, 0
        pos += 1
        if n != prev_n:
            rank, prev_n = pos, n
        row = [CATEGORIES[cat - 1], rank, name, n]
        if full:
            row.append(", ".join(by_cand.get(cid, [])))
        ws.append(row)
    ws.append([])
    ws.append(["Jami ovozlar", "", "", len(voters)])
    _style(ws)
    if full:
        for r in ws.iter_rows(min_row=2, min_col=5, max_col=5):
            r[0].alignment = Alignment(wrap_text=True, vertical="top")

    if full:
        w2 = wb.create_sheet("Ovoz berganlar")
        w2.append(["№", "Ism-familiya", "Yo'nalish", "Nomzod", "Vaqt", "Telegram ID", "Username"])
        for i, (_, vname, ts, tg_id, un, cat, cname) in enumerate(voters, 1):
            t = datetime.fromisoformat(ts).strftime("%d.%m.%Y %H:%M:%S") if ts else ""
            w2.append([i, vname, CATEGORIES[cat - 1], cname, t, tg_id, f"@{un}" if un else ""])
        _style(w2)

    bio = BytesIO()
    wb.save(bio)
    return bio.getvalue()


async def send_results(bot, chat_id, full):
    name = "natijalar_toliq.xlsx" if full else "natijalar.xlsx"
    cap = "📊 To'liq natijalar (ovoz berganlar bilan)" if full else "📊 Natijalar (Excel)"
    await bot.send_document(chat_id, make_excel(full), filename=name, caption=cap)


# ---------- Tugma yordamchilari ----------

def cat_rows(prefix):
    btns = [Btn(n, callback_data=f"{prefix}:{i}") for i, n in enumerate(CATEGORIES, 1)]
    return [btns[j:j + 2] for j in range(0, len(btns), 2)]


def main_kb(uid):
    rows = cat_rows("cat")
    rows.append([Btn("📊 Natijalar", callback_data="res")])
    if uid == ADMIN_ID:
        rows.append([Btn("⚙️ Admin panel", callback_data="a:panel")])
    return Kb(rows)


async def show(q, text, kb=None):
    """Xabarni tahrirlaydi; rasmli xabar bo'lsa yangisini yuboradi."""
    try:
        await q.message.edit_text(text, reply_markup=kb)
    except Exception:
        await q.message.reply_text(text, reply_markup=kb)


MAIN_TEXT = "Yo'nalishni tanlang.\n⚠️ Har bir ism-familiya faqat 1 marta ovoz bera oladi."


# ---------- Foydalanuvchi ----------

async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data.clear()
    uid = update.effective_user.id
    with db() as c:
        c.execute("INSERT OR IGNORE INTO users VALUES(?)", (uid,))
    text = "Ovoz berish yakunlangan." if is_ended() else "Salom! " + MAIN_TEXT
    await update.message.reply_text(text, reply_markup=main_kb(uid))


async def menu(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    ctx.user_data.clear()
    await show(q, MAIN_TEXT, main_kb(q.from_user.id))


async def show_cat(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    if is_ended():
        await q.answer("Ovoz berish yakunlangan.", show_alert=True)
        return
    await q.answer()
    cat = int(q.data.split(":")[1])
    with db() as c:
        cands = c.execute("SELECT id, name, photo FROM candidates WHERE cat=?", (cat,)).fetchall()
    if not cands:
        await show(q, f"{CATEGORIES[cat - 1]}\n\nBu yo'nalishda hozircha nomzod yo'q.", Kb([BACK_MENU]))
        return
    await q.message.reply_text(f"📌 {CATEGORIES[cat - 1]}")
    for cid, name, photo in cands:
        kb = Kb([[Btn("🗳 Ovoz berish", callback_data=f"vote:{cid}")]])
        await ctx.bot.send_photo(q.message.chat_id, photo, caption=name, reply_markup=kb)
    await q.message.reply_text("Boshqa yo'nalishlarni ko'rish:", reply_markup=Kb([BACK_MENU]))


async def vote(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """1-qadam: nomzod tanlandi, ism-familiya so'raladi."""
    q = update.callback_query
    if is_ended():
        await q.answer("Ovoz berish yakunlangan.", show_alert=True)
        return
    cid = int(q.data.split(":")[1])
    with db() as c:
        cand = c.execute("SELECT name FROM candidates WHERE id=?", (cid,)).fetchone()
    if not cand:
        await q.answer("Nomzod topilmadi.", show_alert=True)
        return
    await q.answer()
    ctx.user_data.update(vstate="name", vcand=cid)
    await q.message.reply_text(
        f"🗳 Nomzod: {cand[0]}\n\n✍️ Ism va familiyangizni yozing (masalan: Jasur Karimov).\n"
        "⚠️ Har bir ism-familiya faqat 1 marta ovoz bera oladi, keyin o'zgartirib bo'lmaydi.",
        reply_markup=Kb([[Btn("❌ Bekor qilish", callback_data="menu")]]))


async def vote_name(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """2-qadam: ism-familiya yozildi, tasdiqlash so'raladi."""
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
    with db() as c:
        taken = c.execute("SELECT 1 FROM votes WHERE voter_key=?", (key,)).fetchone()
        cand = c.execute("SELECT name, cat FROM candidates WHERE id=?", (ud["vcand"],)).fetchone()
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
    await m.reply_text(f"👤 Ovoz beruvchi: {name}\n📌 {CATEGORIES[cand[1] - 1]}\n🗳 Nomzod: {cand[0]}\n\n"
                       "⚠️ Ovoz qabul qilingach, o'zgartirib bo'lmaydi.", reply_markup=kb)


async def vote_confirm(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """3-qadam: ovozni bazaga yozish."""
    q, ud = update.callback_query, ctx.user_data
    if is_ended():
        await q.answer("Ovoz berish yakunlangan.", show_alert=True)
        return
    if ud.get("vstate") != "confirm":
        await q.answer("Sessiya tugagan. Qaytadan boshlang.", show_alert=True)
        return
    with db() as c:
        exists = c.execute("SELECT 1 FROM candidates WHERE id=?", (ud["vcand"],)).fetchone()
    if not exists:
        ud.clear()
        await q.answer("Nomzod topilmadi.", show_alert=True)
        return
    try:
        with db() as c:
            c.execute("""INSERT INTO votes(voter_key, voter_name, cand_id, tg_id, tg_username, ts)
                         VALUES(?,?,?,?,?,?)""",
                      (ud["vkey"], ud["vname"], ud["vcand"], q.from_user.id,
                       q.from_user.username, datetime.now(TZ).isoformat()))
    except sqlite3.IntegrityError:
        ud.clear()
        await q.answer("Bu ism-familiya bilan allaqachon ovoz berilgan.", show_alert=True)
        return
    name = ud["vname"]
    ud.clear()
    await q.answer("✅ Ovozingiz qabul qilindi!", show_alert=True)
    await show(q, f"✅ {name}, ovozingiz qabul qilindi.\nNatijalar ovoz berish tugagach e'lon qilinadi.",
               Kb([BACK_MENU]))
    schedule_restart(ctx.application)


async def results(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    is_admin = q.from_user.id == ADMIN_ID
    if not (is_ended() or is_admin):
        await q.answer("Natijalar ovoz berish tugagach e'lon qilinadi.", show_alert=True)
        return
    await q.answer()
    await show(q, results_text(), Kb([BACK_MENU]))
    await send_results(ctx.bot, q.message.chat_id, is_admin or PUBLIC_SHOWS_VOTERS)


# ---------- Admin ----------

def admin_only(fn):
    @wraps(fn)
    async def wrapper(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        if update.effective_user.id != ADMIN_ID:
            if update.callback_query:
                await update.callback_query.answer("Ruxsat yo'q", show_alert=True)
            return
        return await fn(update, ctx)
    return wrapper


def panel_text():
    with db() as c:
        users = c.execute("SELECT COUNT(*) FROM users").fetchone()[0]
        votes = c.execute("SELECT COUNT(*) FROM votes").fetchone()[0]
        cands = c.execute("SELECT COUNT(*) FROM candidates").fetchone()[0]
    end = get_end()
    end_s = f"{end:%d.%m.%Y %H:%M}" if end else "belgilanmagan"
    return (f"⚙️ Admin panel\n\n👥 Foydalanuvchilar: {users}\n🗳 Ovozlar: {votes}\n"
            f"👤 Nomzodlar: {cands}\n⏰ Tugash: {end_s}")


def panel_kb():
    return Kb([
        [Btn("➕ Nomzod qo'shish", callback_data="a:add"), Btn("📋 Nomzodlar", callback_data="a:list")],
        [Btn("⏰ Tugash vaqti", callback_data="a:end"), Btn("📊 Natijalar", callback_data="a:res")],
        [Btn("⬅️ Asosiy menyu", callback_data="menu")],
    ])


@admin_only
async def admin_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data.clear()
    await update.message.reply_text(panel_text(), reply_markup=panel_kb())


@admin_only
async def admin_cb(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    p = q.data.split(":")
    act = p[1]

    if act == "panel":
        ctx.user_data.clear()
        await show(q, panel_text(), panel_kb())

    elif act == "add":
        await show(q, "Qaysi yo'nalishga nomzod qo'shamiz?", Kb(cat_rows("a:addcat") + [PANEL_BACK]))

    elif act == "addcat":
        n = int(p[2])
        ctx.user_data.update(state="add_name", cat=n)
        await show(q, f"{CATEGORIES[n - 1]}\n\n✍️ Nomzodning ism-familiyasini yozing:", Kb([CANCEL]))

    elif act == "list":
        await show(q, "Qaysi yo'nalish nomzodlari?", Kb(cat_rows("a:listcat") + [PANEL_BACK]))

    elif act == "listcat":
        n = int(p[2])
        with db() as c:
            rows = c.execute("""SELECT c.id, c.name, c.photo, COUNT(v.id) FROM candidates c
                LEFT JOIN votes v ON v.cand_id = c.id WHERE c.cat=? GROUP BY c.id""", (n,)).fetchall()
        if not rows:
            await show(q, f"{CATEGORIES[n - 1]}: nomzod yo'q.", Kb([PANEL_BACK]))
            return
        await q.message.reply_text(f"📌 {CATEGORIES[n - 1]}")
        for cid, name, photo, cnt in rows:
            kb = Kb([[Btn("🗑 O'chirish", callback_data=f"a:del:{cid}")]])
            await ctx.bot.send_photo(q.message.chat_id, photo,
                                     caption=f"{name}\n🗳 {cnt} ovoz", reply_markup=kb)
        await q.message.reply_text("Admin panelga qaytish:", reply_markup=Kb([PANEL_BACK]))

    elif act == "del":
        cid = int(p[2])
        with db() as c:
            c.execute("DELETE FROM votes WHERE cand_id=?", (cid,))
            c.execute("DELETE FROM candidates WHERE id=?", (cid,))
        await q.message.edit_caption("🗑 O'chirildi", reply_markup=None)

    elif act == "end":
        kb = Kb([
            [Btn("+1 kun", callback_data="a:endin:1"), Btn("+3 kun", callback_data="a:endin:3"),
             Btn("+7 kun", callback_data="a:endin:7")],
            [Btn("✍️ Sana va vaqtni yozish", callback_data="a:endman")],
            [Btn("🏁 Hozir yakunlash", callback_data="a:endnow")],
            PANEL_BACK,
        ])
        await show(q, "⏰ Ovoz berish qachon tugasin? (hozirdan boshlab)", kb)

    elif act == "endin":
        dt = datetime.now(TZ) + timedelta(days=int(p[2]))
        save_end(dt)
        schedule_finish(ctx.application)
        await show(q, f"✅ Tugash vaqti: {dt:%d.%m.%Y %H:%M}", Kb([PANEL_BACK]))

    elif act == "endman":
        ctx.user_data["state"] = "set_end"
        await show(q, "Sana va vaqtni yozing (Toshkent vaqti):\n2026-10-05 18:00", Kb([CANCEL]))

    elif act == "endnow":
        kb = Kb([[Btn("✅ Ha, yakunlash", callback_data="a:endyes"),
                  Btn("❌ Yo'q", callback_data="a:panel")]])
        await show(q, "⚠️ Ovoz berish hoziroq tugaydi va natijalar hammaga yuboriladi. Tasdiqlaysizmi?", kb)

    elif act == "endyes":
        save_end(datetime.now(TZ))
        schedule_finish(ctx.application)
        await show(q, "🏁 Ovoz berish yakunlandi. Natijalar yuborilmoqda.", Kb([PANEL_BACK]))

    elif act == "res":
        await show(q, results_text(), Kb([PANEL_BACK]))
        await send_results(ctx.bot, q.message.chat_id, True)


async def admin_text(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    m = update.message
    state = ctx.user_data.get("state")
    if state == "add_name":
        ctx.user_data.update(name=m.text.strip(), state="add_photo")
        await m.reply_text("📷 Endi nomzodning rasmini yuboring:", reply_markup=Kb([CANCEL]))
    elif state == "set_end":
        try:
            dt = datetime.strptime(m.text.strip(), "%Y-%m-%d %H:%M").replace(tzinfo=TZ)
        except ValueError:
            await m.reply_text("Format noto'g'ri. Masalan: 2026-10-05 18:00")
            return
        save_end(dt)
        schedule_finish(ctx.application)
        ctx.user_data.clear()
        await m.reply_text(f"✅ Tugash vaqti: {dt:%d.%m.%Y %H:%M}", reply_markup=Kb([PANEL_BACK]))


@admin_only
async def admin_photo(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if ctx.user_data.get("state") != "add_photo":
        return
    n, name = ctx.user_data["cat"], ctx.user_data["name"]
    with db() as c:
        c.execute("INSERT INTO candidates(cat, name, photo) VALUES(?,?,?)",
                  (n, name, update.message.photo[-1].file_id))
    ctx.user_data.clear()
    kb = Kb([[Btn("➕ Yana qo'shish", callback_data=f"a:addcat:{n}")], PANEL_BACK])
    await update.message.reply_text(f"✅ Qo'shildi: {name} → {CATEGORIES[n - 1]}", reply_markup=kb)


async def on_text(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ud = ctx.user_data
    if update.effective_user.id == ADMIN_ID and ud.get("state") in ("add_name", "set_end"):
        await admin_text(update, ctx)
    elif ud.get("vstate") == "name":
        await vote_name(update, ctx)


# ---------- Avtomatik yakunlash va restart ----------

async def finish(ctx: ContextTypes.DEFAULT_TYPE):
    text = results_text()
    public = make_excel(PUBLIC_SHOWS_VOTERS)
    with db() as c:
        users = {r[0] for r in c.execute("SELECT user_id FROM users")} - {ADMIN_ID}
    file_id = None
    for uid in users:
        try:
            await ctx.bot.send_message(uid, text)
            if file_id:
                await ctx.bot.send_document(uid, file_id, caption="📊 Natijalar (Excel)")
            else:
                msg = await ctx.bot.send_document(uid, public, filename="natijalar.xlsx",
                                                  caption="📊 Natijalar (Excel)")
                file_id = msg.document.file_id
        except Exception:
            pass
        await asyncio.sleep(0.05)
    try:
        await ctx.bot.send_message(ADMIN_ID, text)
        await send_results(ctx.bot, ADMIN_ID, True)
    except Exception:
        log.exception("Adminga natijani yuborib bo'lmadi")


def schedule_finish(app: Application):
    for job in app.job_queue.get_jobs_by_name("finish"):
        job.schedule_removal()
    end = get_end()
    if end:
        delay = max((end - datetime.now(TZ)).total_seconds(), 1)
        app.job_queue.run_once(finish, delay, name="finish")


def schedule_restart(app: Application):
    """Ovoz qabul qilingach 3 soniyadan keyin restart (ketma-ket ovozlar bitta restartga birlashadi)."""
    if RESTART_AFTER_VOTE and not app.job_queue.get_jobs_by_name("restart"):
        app.job_queue.run_once(restart, 3, name="restart")


async def restart(ctx: ContextTypes.DEFAULT_TYPE):
    script = sys.argv[0] if sys.argv else ""
    if not (script and os.path.isfile(script)):
        log.warning("Avto-restart o'tkazib yuborildi: skript yo'li aniqlanmadi.")
        return
    try:
        await ctx.application.update_persistence()  # jarayondagi ovoz berish holatlari saqlanadi
        log.info("Bot qayta ishga tushmoqda...")
        os.execv(sys.executable, [sys.executable] + sys.argv)
    except Exception:
        log.exception("Avto-restart ishlamadi, bot davom etadi.")


async def post_init(app: Application):
    schedule_finish(app)


def main():
    if TOKEN.startswith("BOT_TOKEN_SHU"):
        sys.exit("BOT_TOKEN o'rnatilmagan: Railway → Variables ga BOT_TOKEN qo'shing.")
    if (os.getenv("RAILWAY_ENVIRONMENT_NAME") or os.getenv("RAILWAY_ENVIRONMENT")) \
            and not os.getenv("RAILWAY_VOLUME_MOUNT_PATH"):
        log.warning("Volume ulanmagan: qayta deploydan keyin ovozlar va nomzodlar yo'qoladi!")
    init_db()
    persistence = PicklePersistence(
        STATE_FILE, update_interval=5,
        store_data=PersistenceInput(bot_data=False, chat_data=False, callback_data=False))
    app = (ApplicationBuilder().token(TOKEN).persistence(persistence)
           .post_init(post_init).build())
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("admin", admin_cmd, filters=admin))
    app.add_handler(CallbackQueryHandler(menu, pattern=r"^menu$"))
    app.add_handler(CallbackQueryHandler(show_cat, pattern=r"^cat:\d+$"))
    app.add_handler(CallbackQueryHandler(vote, pattern=r"^vote:\d+$"))
    app.add_handler(CallbackQueryHandler(vote_confirm, pattern=r"^vc:yes$"))
    app.add_handler(CallbackQueryHandler(results, pattern=r"^res$"))
    app.add_handler(CallbackQueryHandler(admin_cb, pattern=r"^a:"))
    app.add_handler(MessageHandler(filters.PHOTO & admin, admin_photo))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, on_text))
    app.run_polling()


if __name__ == "__main__":
    main()
