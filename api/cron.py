from http.server import BaseHTTPRequestHandler
import asyncio
import os

from telegram import Bot

BOT_TOKEN = os.environ.get('BOT_TOKEN', '')
ADMIN_ID = int(os.environ.get('ADMIN_ID', '807823872'))
GROUP_ID = int(os.environ.get('GROUP_ID', '-1002449545348'))


async def send_report():
    from lib.excel import get_bugungi_tulumlar, get_kunlik_tushum
    bot = Bot(token=BOT_TOKEN)
    tulumlar = get_bugungi_tulumlar()
    tushum = get_kunlik_tushum()
    for chat_id in [ADMIN_ID, GROUP_ID]:
        for msg in tulumlar:
            await bot.send_message(chat_id=chat_id, text=msg, parse_mode='HTML')
        await bot.send_message(chat_id=chat_id, text=tushum)


class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        try:
            asyncio.run(send_report())
        except Exception as e:
            print(f"Cron error: {e}")
        self.send_response(200)
        self.send_header('Content-type', 'application/json')
        self.end_headers()
        self.wfile.write(b'{"ok":true}')

    def log_message(self, format, *args):
        pass
