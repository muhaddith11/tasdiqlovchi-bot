from flask import Flask, request, jsonify
import asyncio
import os
import psycopg2
import html as html_mod
import requests as req_lib
import openpyxl
import io
import datetime
import time
import pytz

from telegram import Bot, InlineKeyboardButton, InlineKeyboardMarkup, Update

app = Flask(__name__)

BOT_TOKEN = os.environ.get('BOT_TOKEN', '')
ADMIN_ID = int(os.environ.get('ADMIN_ID', '807823872'))
GROUP_ID = int(os.environ.get('GROUP_ID', '-1002449545348'))
DATABASE_URL = os.environ.get('DATABASE_URL', '')
ONEDRIVE_URL = os.environ.get('ONEDRIVE_URL', 'https://1drv.ms/x/c/0434e9c0edef097b/IQASHiM8IYUQSZNJNl0nojFBAcv7R4dXvdm4vdX1NQN-AJw?e=38Ffl7&download=1')
TASHKENT_TZ = pytz.timezone('Asia/Tashkent')
ALLOWED_COMMANDS = ['/kirim', '/chiqim']


# ─── Database ─────────────────────────────────────────────────────────────────

def get_db():
    return psycopg2.connect(DATABASE_URL)


def ensure_table():
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute('''
            CREATE TABLE IF NOT EXISTS pending (
                key TEXT PRIMARY KEY,
                sender_name TEXT,
                username TEXT,
                chat_id BIGINT,
                msg_type TEXT,
                msg_text TEXT,
                file_id TEXT,
                caption TEXT,
                created_at TIMESTAMP DEFAULT NOW()
            )
        ''')
        conn.commit()
        cur.close()
        conn.close()
    except Exception as e:
        print(f"DB init error: {e}")


# ─── Excel ────────────────────────────────────────────────────────────────────

def excel_yukla():
    url = ONEDRIVE_URL + f"&nocache={int(time.time())}"
    r = req_lib.get(url, timeout=30, headers={"Cache-Control": "no-cache"})
    return openpyxl.load_workbook(io.BytesIO(r.content), data_only=True, read_only=True)


def get_kunlik_tushum():
    try:
        wb = excel_yukla()
        today = datetime.datetime.now(TASHKENT_TZ).date()
        sheets = ['Салом сити-1', 'Салом сити-2', 'МЖК-1', 'МЖК-2']
        results = {}
        for sheet_name in sheets:
            ws = wb[sheet_name]
            total = 0
            for row in ws.iter_rows(values_only=True):
                date_cell = row[9] if len(row) > 9 else None
                amount_cell = row[10] if len(row) > 10 else None
                if isinstance(date_cell, datetime.datetime) and date_cell.date() == today:
                    total += (amount_cell or 0)
            results[sheet_name] = total
        jami = sum(results.values())
        today_str = today.strftime("%d.%m.%Y")
        text = f"📊 Кунлик тушум\n📅 Сана: {today_str}\n{'─' * 30}\n"
        for name, val in results.items():
            text += f"▪️ {name:<18} {val:>10,.0f}\n"
        text += f"{'─' * 30}\n💰 Жами:               {jami:>10,.0f}"
        return text
    except Exception as e:
        return f"❌ Xatolik: {e}"


def get_bugungi_tulumlar():
    try:
        wb = excel_yukla()
        today = datetime.datetime.now(TASHKENT_TZ).date()
        sheets = ['Салом сити-1', 'Салом сити-2', 'МЖК-1', 'МЖК-2']
        messages = []
        for sheet_name in sheets:
            ws = wb[sheet_name]
            current_apt = None
            sheet_payments = []
            payment_count = 0
            last_payment_date = None
            for row in ws.iter_rows(min_row=7, values_only=True):
                if len(row) < 11:
                    continue
                fio = row[5]
                if fio and isinstance(fio, str) and fio.strip() and fio.strip() != 'фио':
                    foiz_val = row[12] if len(row) > 12 else None
                    current_apt = {
                        'fio': fio.strip(), 'kv': row[2], 'dom': row[3],
                        'etaj': row[4], 'tulangan': row[10], 'qarz': row[11], 'foiz': foiz_val
                    }
                    payment_count = 0
                    last_payment_date = None
                if current_apt:
                    date_val = row[9]
                    amount = row[10]
                    if isinstance(date_val, datetime.datetime) and amount:
                        payment_count += 1
                        if date_val.date() == today:
                            sheet_payments.append({**current_apt, 'berdi': amount,
                                                   'toliq_son': payment_count, 'oldingi_tulov': last_payment_date})
                        else:
                            last_payment_date = date_val.date()
            if sheet_payments:
                text = f"🏢 {sheet_name} — {today.strftime('%d.%m.%Y')}\n{'─' * 30}\n\n"
                for i, p in enumerate(sheet_payments, 1):
                    foiz = p['foiz']
                    foiz_str = f"{foiz * 100:.0f}%" if isinstance(foiz, float) else "—"
                    oldingi = p['oldingi_tulov']
                    oldingi_str = oldingi.strftime("%d.%m.%Y") if oldingi else "birinchi to'lov"
                    fio_e = html_mod.escape(str(p['fio']))
                    dom_e = html_mod.escape(str(p['dom']))
                    etaj_e = html_mod.escape(str(p['etaj']))
                    kv_e = html_mod.escape(str(p['kv']))
                    text += (
                        f"{i}. 👤 {fio_e}\n"
                        f"   🏠 {dom_e}-дом, {etaj_e}-этаж, {kv_e}-кв\n"
                        f"   🔢 {p['toliq_son']}-chi to'lov\n"
                        f"   📅 Oldingi to'lov: {oldingi_str}\n"
                        f"   💵 <b>Bugun berdi:   ${p['berdi']:,.0f}</b>\n"
                        f"   ✅ Jami to'lagan: ${p['tulangan']:,.0f}\n"
                        f"   ❌ Qolgan qarz:   ${p['qarz']:,.0f}\n"
                        f"   📊 Foizda: {foiz_str}\n\n"
                    )
                text += f"{'─' * 30}\nJami: {len(sheet_payments)} ta to'lov"
                messages.append(text)
        return messages if messages else ["📭 Bugun hech qanday to'lov kiritilmagan"]
    except Exception as e:
        return [f"❌ Xatolik: {e}"]


# ─── Telegram handlers ────────────────────────────────────────────────────────

async def handle_group_message(bot, message):
    if not message or message.chat.type not in ['group', 'supergroup']:
        return
    text = message.text or message.caption or ""
    if not any(text.startswith(cmd) for cmd in ALLOWED_COMMANDS):
        return

    sender = message.from_user
    sender_name = sender.full_name
    username = f"@{sender.username}" if sender.username else "username yo'q"
    chat_id = message.chat_id
    key = f"{chat_id}_{message.message_id}"

    if message.text:
        msg_type, msg_text, file_id, caption = 'text', message.text, None, None
    elif message.photo:
        msg_type, msg_text, file_id, caption = 'photo', None, message.photo[-1].file_id, message.caption or ""
    elif message.video:
        msg_type, msg_text, file_id, caption = 'video', None, message.video.file_id, message.caption or ""
    elif message.document:
        msg_type, msg_text, file_id, caption = 'document', None, message.document.file_id, message.caption or ""
    else:
        return

    conn = get_db()
    cur = conn.cursor()
    cur.execute('''INSERT INTO pending (key, sender_name, username, chat_id, msg_type, msg_text, file_id, caption)
        VALUES (%s,%s,%s,%s,%s,%s,%s,%s) ON CONFLICT (key) DO NOTHING''',
        (key, sender_name, username, chat_id, msg_type, msg_text, file_id, caption))
    conn.commit()
    cur.close()
    conn.close()

    keyboard = [[
        InlineKeyboardButton("✅ Tasdiqlash", callback_data=f"approve_{key}"),
        InlineKeyboardButton("❌ Rad etish", callback_data=f"reject_{key}")
    ]]
    footer = f"\n\n👤 Yuboruvchi: {sender_name}\n📱 Username: {username}"

    if msg_type == 'text':
        await bot.send_message(chat_id=ADMIN_ID, text=msg_text + footer, reply_markup=InlineKeyboardMarkup(keyboard))
    elif msg_type == 'photo':
        await bot.send_photo(chat_id=ADMIN_ID, photo=file_id, caption=caption + footer, reply_markup=InlineKeyboardMarkup(keyboard))
    elif msg_type == 'video':
        await bot.send_video(chat_id=ADMIN_ID, video=file_id, caption=caption + footer, reply_markup=InlineKeyboardMarkup(keyboard))
    elif msg_type == 'document':
        await bot.send_document(chat_id=ADMIN_ID, document=file_id, caption=caption + footer, reply_markup=InlineKeyboardMarkup(keyboard))


async def handle_callback(bot, query):
    await query.answer()
    parts = query.data.split('_', 1)
    if len(parts) != 2:
        return
    action, key = parts

    conn = get_db()
    cur = conn.cursor()
    cur.execute('SELECT sender_name, username, chat_id, msg_type, msg_text, file_id, caption FROM pending WHERE key=%s', (key,))
    row = cur.fetchone()
    if not row:
        await query.edit_message_text("Bu xabar allaqachon ko'rib chiqilgan.")
        cur.close()
        conn.close()
        return

    sender_name, username, chat_id, msg_type, msg_text, file_id, caption = row
    cur.execute('DELETE FROM pending WHERE key=%s', (key,))
    conn.commit()
    cur.close()
    conn.close()

    status = "✅ Boshliq tasdiqladi" if action == "approve" else "❌ Boshliq rad etdi"
    footer = f"\n\n👤 Yuboruvchi: {sender_name}\n📱 Username: {username}\n\n{status}"

    if msg_type == 'text':
        await bot.send_message(chat_id=chat_id, text=msg_text + footer)
    elif msg_type == 'photo':
        await bot.send_photo(chat_id=chat_id, photo=file_id, caption=(caption or "") + footer)
    elif msg_type == 'video':
        await bot.send_video(chat_id=chat_id, video=file_id, caption=(caption or "") + footer)
    elif msg_type == 'document':
        await bot.send_document(chat_id=chat_id, document=file_id, caption=(caption or "") + footer)

    original_text = query.message.text or query.message.caption or ""
    new_text = original_text + f"\n\n{status}"
    try:
        if query.message.text:
            await query.edit_message_text(text=new_text)
        elif query.message.caption is not None:
            await query.edit_message_caption(caption=new_text)
    except Exception:
        pass


async def process_update(data):
    bot = Bot(token=BOT_TOKEN)
    update = Update.de_json(data, bot)
    if update.message:
        msg = update.message
        text = msg.text or msg.caption or ""
        if text.startswith('/hisobot'):
            for m in get_bugungi_tulumlar():
                await bot.send_message(chat_id=msg.chat_id, text=m, parse_mode='HTML')
            await bot.send_message(chat_id=msg.chat_id, text=get_kunlik_tushum())
        elif text.startswith('/chatid'):
            await bot.send_message(chat_id=msg.chat_id, text=f"Chat ID: `{msg.chat_id}`", parse_mode='Markdown')
        elif any(text.startswith(cmd) for cmd in ALLOWED_COMMANDS):
            await handle_group_message(bot, msg)
    elif update.callback_query:
        await handle_callback(bot, update.callback_query)


async def send_daily_report():
    bot = Bot(token=BOT_TOKEN)
    tulumlar = get_bugungi_tulumlar()
    tushum = get_kunlik_tushum()
    for chat_id in [ADMIN_ID, GROUP_ID]:
        for msg in tulumlar:
            await bot.send_message(chat_id=chat_id, text=msg, parse_mode='HTML')
        await bot.send_message(chat_id=chat_id, text=tushum)


# ─── Routes ───────────────────────────────────────────────────────────────────

@app.route('/api/webhook', methods=['POST'])
def webhook():
    ensure_table()
    data = request.get_json(force=True, silent=True) or {}
    try:
        asyncio.run(process_update(data))
    except Exception as e:
        print(f"Webhook error: {e}")
    return jsonify({"ok": True})


@app.route('/api/cron', methods=['GET'])
def cron():
    try:
        asyncio.run(send_daily_report())
    except Exception as e:
        print(f"Cron error: {e}")
    return jsonify({"ok": True})


@app.route('/', methods=['GET'])
def index():
    return jsonify({"status": "running"})
