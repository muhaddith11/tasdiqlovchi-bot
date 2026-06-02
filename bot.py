import logging
import requests
import openpyxl
import io
import json
import os
from datetime import datetime
import pytz
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, BotCommand
from telegram.ext import Application, MessageHandler, CallbackQueryHandler, CommandHandler, filters, ContextTypes
from apscheduler.schedulers.asyncio import AsyncIOScheduler

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
        r = requests.get(ONEDRIVE_URL, timeout=30)
        wb = openpyxl.load_workbook(io.BytesIO(r.content), data_only=True)
        ws = wb.worksheets[0]  # Реестр varag'i

        # Jadval: 6-12 qatorlar, E-F ustunlar (5-6)
        rows = list(ws.iter_rows(min_row=6, max_row=12, min_col=5, max_col=6, values_only=True))

        # Sana (7-qator, F ustun)
        sana = rows[1][1]
        if isinstance(sana, datetime):
            sana_str = sana.strftime("%d.%m.%Y")
        else:
            sana_str = str(sana) if sana else "—"

        # 4 ta qator: салом сити-1, салом сити-2, мжк-1, мжк-2
        items = []
        for row in rows[2:6]:
            name = str(row[0] or "").capitalize()
            val = row[1] or 0
            items.append((name, val))

        jami = rows[6][1] or 0

        # Chiroyli matn
        text = (
            f"📊 Кунлик тушум\n"
            f"📅 Сана: {sana_str}\n"
            f"{'─' * 30}\n"
        )
        for name, val in items:
            text += f"▪️ {name:<18} {val:>10,.0f}\n"
        text += (
            f"{'─' * 30}\n"
            f"💰 Жами:               {jami:>10,.0f}"
        )
        return text

    except Exception as e:
        logging.error(f"Excel xatosi: {e}")
        return f"❌ Ma'lumot olishda xatolik yuz berdi."


# ─── Xabar moderatsiyasi ──────────────────────────────────────────────────────

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    message = update.message
    if not message or message.chat.type not in ['group', 'supergroup']:
        return

    # Guruh ID ni saqlash (hisobot uchun kerak)
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
        await context.bot.send_message(
            chat_id=ADMIN_ID,
            text=message.text + footer,
            reply_markup=reply_markup
        )
    elif message.photo:
        await context.bot.send_photo(
            chat_id=ADMIN_ID,
            photo=message.photo[-1].file_id,
            caption=(message.caption or "") + footer,
            reply_markup=reply_markup
        )
    elif message.video:
        await context.bot.send_video(
            chat_id=ADMIN_ID,
            video=message.video.file_id,
            caption=(message.caption or "") + footer,
            reply_markup=reply_markup
        )
    elif message.document:
        await context.bot.send_document(
            chat_id=ADMIN_ID,
            document=message.document.file_id,
            caption=(message.caption or "") + footer,
            reply_markup=reply_markup
        )


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
        await context.bot.send_photo(
            chat_id=chat_id,
            photo=message.photo[-1].file_id,
            caption=(message.caption or "") + footer
        )
    elif message.video:
        await context.bot.send_video(
            chat_id=chat_id,
            video=message.video.file_id,
            caption=(message.caption or "") + footer
        )
    elif message.document:
        await context.bot.send_document(
            chat_id=chat_id,
            document=message.document.file_id,
            caption=(message.caption or "") + footer
        )

    # Boshliq botida statusni yangilash
    original_text = query.message.text or query.message.caption or ""
    new_text = original_text + f"\n\n{status}"
    if query.message.text:
        await query.edit_message_text(text=new_text)
    elif query.message.caption is not None:
        await query.edit_message_caption(caption=new_text)


# ─── Hisobot buyrug'i ─────────────────────────────────────────────────────────

async def hisobot_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = get_kunlik_tushum()
    await update.message.reply_text(text)


# ─── Har kuni 21:00 da avtomatik yuborish ─────────────────────────────────────

async def send_daily_report(app):
    group_chat_id = get_group_chat_id()
    if not group_chat_id:
        logging.warning("Guruh ID topilmadi!")
        return
    text = get_kunlik_tushum()
    await app.bot.send_message(chat_id=group_chat_id, text=text)


# ─── Asosiy ──────────────────────────────────────────────────────────────────

async def post_init(app):
    await app.bot.set_my_commands([
        BotCommand("hisobot", "📊 Kunlik hisobot"),
    ])
    # Scheduler event loop ichida ishga tushiriladi
    scheduler = AsyncIOScheduler(timezone=TASHKENT_TZ)
    scheduler.add_job(send_daily_report, trigger='cron', hour=21, minute=0, args=[app])
    scheduler.start()


def main():
    app = Application.builder().token(BOT_TOKEN).post_init(post_init).build()

    app.add_handler(MessageHandler(filters.ChatType.GROUPS, handle_message))
    app.add_handler(CallbackQueryHandler(handle_callback))
    app.add_handler(CommandHandler("hisobot", hisobot_command))

    app.run_polling()


if __name__ == '__main__':
    main()
