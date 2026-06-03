import logging
import requests
import openpyxl
import io
import json
import os
import datetime
import time
import pytz
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, BotCommand
from telegram.ext import Application, MessageHandler, CallbackQueryHandler, CommandHandler, filters, ContextTypes

BOT_TOKEN = "8847266024:AAGA00Bqrw3ekbo5TCSmusK3Yd0FU2exTsM"
ADMIN_ID = 8042807902
ONEDRIVE_URL = "https://1drv.ms/x/c/0434e9c0edef097b/IQASHiM8IYUQSZNJNl0nojFBAcv7R4dXvdm4vdX1NQN-AJw?e=38Ffl7&download=1"
TASHKENT_TZ = pytz.timezone('Asia/Tashkent')
DATA_FILE = "data.json"

logging.basicConfig(level=logging.INFO)
pending = {}
ALLOWED_COMMANDS = ['/kirim', '/chiqim']


# ─── Ma'lumotlarni saqlash ────────────────────────────────────────────────────

def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE) as f:
            return json.load(f)
    return {}

def save_data(data):
    with open(DATA_FILE, 'w') as f:
        json.dump(data, f)

def get_group_chat_id():
    return load_data().get('group_chat_id')

def set_group_chat_id(chat_id):
    data = load_data()
    data['group_chat_id'] = chat_id
    save_data(data)


# ─── Excel dan kunlik tushum ──────────────────────────────────────────────────

def get_kunlik_tushum():
    try:
        url = ONEDRIVE_URL + f"&nocache={int(time.time())}"
        r = requests.get(url, timeout=30)
        wb = openpyxl.load_workbook(io.BytesIO(r.content), data_only=True)

        # Bugungi sana (Toshkent vaqti)
        today = datetime.datetime.now(TASHKENT_TZ).date()

        # Har bir varaqdan bugungi to'lovlarni yig'amiz
        sheets = ['Салом сити-1', 'Салом сити-2', 'МЖК-1', 'МЖК-2']
        results = {}

        for sheet_name in sheets:
            ws = wb[sheet_name]
            total = 0
            for row in ws.iter_rows(values_only=True):
                date_cell = row[8] if len(row) > 8 else None   # I ustun - sana
                amount_cell = row[9] if len(row) > 9 else None  # J ustun - summa
                if isinstance(date_cell, datetime.datetime) and date_cell.date() == today:
                    total += (amount_cell or 0)
            results[sheet_name] = total

        jami = sum(results.values())
        today_str = today.strftime("%d.%m.%Y")

        text = (
            f"📊 Кунлик тушум\n"
            f"📅 Сана: {today_str}\n"
            f"{'─' * 30}\n"
        )
        for name, val in results.items():
            text += f"▪️ {name:<18} {val:>10,.0f}\n"
        text += (
            f"{'─' * 30}\n"
            f"💰 Жами:               {jami:>10,.0f}"
        )
        return text

    except Exception as e:
        logging.error(f"Excel xatosi: {e}")
        return "❌ Ma'lumot olishda xatolik yuz berdi."


# ─── Xabar moderatsiyasi ──────────────────────────────────────────────────────

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    message = update.message
    if not message or message.chat.type not in ['group', 'supergroup']:
        return

    set_group_chat_id(message.chat_id)

    text = message.text or message.caption or ""
    if not any(text.startswith(cmd) for cmd in ALLOWED_COMMANDS):
        return

    sender = message.from_user
    sender_name = sender.full_name
    username = f"@{sender.username}" if sender.username else "username yo'q"
    chat_id = message.chat_id
    key = f"{chat_id}_{message.message_id}"

    pending[key] = {
        'sender_name': sender_name,
        'username': username,
        'chat_id': chat_id,
        'message': message
    }

    keyboard = [[
        InlineKeyboardButton("✅ Tasdiqlash", callback_data=f"approve_{key}"),
        InlineKeyboardButton("❌ Rad etish", callback_data=f"reject_{key}")
    ]]
    reply_markup = InlineKeyboardMarkup(keyboard)
    footer = f"\n\n👤 Yuboruvchi: {sender_name}\n📱 Username: {username}"

    if message.text:
        await context.bot.send_message(chat_id=ADMIN_ID, text=message.text + footer, reply_markup=reply_markup)
    elif message.photo:
        await context.bot.send_photo(chat_id=ADMIN_ID, photo=message.photo[-1].file_id, caption=(message.caption or "") + footer, reply_markup=reply_markup)
    elif message.video:
        await context.bot.send_video(chat_id=ADMIN_ID, video=message.video.file_id, caption=(message.caption or "") + footer, reply_markup=reply_markup)
    elif message.document:
        await context.bot.send_document(chat_id=ADMIN_ID, document=message.document.file_id, caption=(message.caption or "") + footer, reply_markup=reply_markup)


async def handle_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    action, key = query.data.split('_', 1)

    if key not in pending:
        await query.edit_message_text("Bu xabar allaqachon ko'rib chiqilgan.")
        return

    info = pending.pop(key)
    message = info['message']
    sender_name = info['sender_name']
    username = info['username']
    chat_id = info['chat_id']

    status = "✅ Boshliq tasdiqladi" if action == "approve" else "❌ Boshliq rad etdi"
    footer = f"\n\n👤 Yuboruvchi: {sender_name}\n📱 Username: {username}\n\n{status}"

    if message.text:
        await context.bot.send_message(chat_id=chat_id, text=message.text + footer)
    elif message.photo:
        await context.bot.send_photo(chat_id=chat_id, photo=message.photo[-1].file_id, caption=(message.caption or "") + footer)
    elif message.video:
        await context.bot.send_video(chat_id=chat_id, video=message.video.file_id, caption=(message.caption or "") + footer)
    elif message.document:
        await context.bot.send_document(chat_id=chat_id, document=message.document.file_id, caption=(message.caption or "") + footer)

    original_text = query.message.text or query.message.caption or ""
    new_text = original_text + f"\n\n{status}"
    if query.message.text:
        await query.edit_message_text(text=new_text)
    elif query.message.caption is not None:
        await query.edit_message_caption(caption=new_text)


# ─── Hisobot ─────────────────────────────────────────────────────────────────

async def hisobot_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = get_kunlik_tushum()
    await update.message.reply_text(text)


async def send_daily_report(context: ContextTypes.DEFAULT_TYPE):
    group_chat_id = get_group_chat_id()
    if not group_chat_id:
        logging.warning("Guruh ID topilmadi!")
        return
    text = get_kunlik_tushum()
    await context.bot.send_message(chat_id=group_chat_id, text=text)


# ─── Asosiy ──────────────────────────────────────────────────────────────────

async def post_init(app):
    await app.bot.set_my_commands([
        BotCommand("hisobot", "Kunlik hisobot"),
    ])
    # Har kuni soat 21:00 Toshkent vaqtida (JobQueue - ichki, xatosiz)
    app.job_queue.run_daily(
        send_daily_report,
        time=datetime.time(hour=21, minute=0, tzinfo=TASHKENT_TZ)
    )


def main():
    app = Application.builder().token(BOT_TOKEN).post_init(post_init).build()

    app.add_handler(MessageHandler(filters.ChatType.GROUPS, handle_message))
    app.add_handler(CallbackQueryHandler(handle_callback))
    app.add_handler(CommandHandler("hisobot", hisobot_command))

    app.run_polling()


if __name__ == '__main__':
    main()
