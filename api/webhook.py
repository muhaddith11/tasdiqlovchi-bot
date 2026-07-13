from http.server import BaseHTTPRequestHandler
import json
import asyncio
import os
import psycopg2
import html as html_mod

from telegram import Bot, InlineKeyboardButton, InlineKeyboardMarkup, Update

BOT_TOKEN = os.environ.get('BOT_TOKEN', '')
ADMIN_ID = int(os.environ.get('ADMIN_ID', '807823872'))
GROUP_ID = int(os.environ.get('GROUP_ID', '-1002449545348'))
DATABASE_URL = os.environ.get('DATABASE_URL', '')
ALLOWED_COMMANDS = ['/kirim', '/chiqim']


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


async def handle_message(bot, message):
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
    cur.execute('''
        INSERT INTO pending (key, sender_name, username, chat_id, msg_type, msg_text, file_id, caption)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s) ON CONFLICT (key) DO NOTHING
    ''', (key, sender_name, username, chat_id, msg_type, msg_text, file_id, caption))
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
    cur.execute('SELECT sender_name, username, chat_id, msg_type, msg_text, file_id, caption FROM pending WHERE key = %s', (key,))
    row = cur.fetchone()

    if not row:
        await query.edit_message_text("Bu xabar allaqachon ko'rib chiqilgan.")
        cur.close()
        conn.close()
        return

    sender_name, username, chat_id, msg_type, msg_text, file_id, caption = row
    cur.execute('DELETE FROM pending WHERE key = %s', (key,))
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


async def handle_hisobot(bot, chat_id):
    from lib.excel import get_bugungi_tulumlar, get_kunlik_tushum
    messages = get_bugungi_tulumlar()
    for msg in messages:
        await bot.send_message(chat_id=chat_id, text=msg, parse_mode='HTML')
    await bot.send_message(chat_id=chat_id, text=get_kunlik_tushum())


async def process_update(data):
    bot = Bot(token=BOT_TOKEN)
    update = Update.de_json(data, bot)

    if update.message:
        msg = update.message
        text = msg.text or msg.caption or ""

        if text.startswith('/hisobot'):
            await handle_hisobot(bot, msg.chat_id)
        elif text.startswith('/chatid'):
            await bot.send_message(chat_id=msg.chat_id, text=f"Chat ID: `{msg.chat_id}`", parse_mode='Markdown')
        elif any(text.startswith(cmd) for cmd in ALLOWED_COMMANDS):
            await handle_message(bot, msg)

    elif update.callback_query:
        await handle_callback(bot, update.callback_query)


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        ensure_table()
        length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(length)
        try:
            asyncio.run(process_update(json.loads(body)))
        except Exception as e:
            print(f"Webhook error: {e}")
        self.send_response(200)
        self.send_header('Content-type', 'application/json')
        self.end_headers()
        self.wfile.write(b'{"ok":true}')

    def log_message(self, format, *args):
        pass
