"""
Telegram-бот для анализа отчёта по часу продаж.
"""
import os
import logging
import tempfile
import threading
from http.server import HTTPServer, BaseHTTPRequestHandler
from telegram import Update
from telegram.ext import Application, MessageHandler, CommandHandler, filters, ContextTypes

from report_analyzer import parse_report, generate_pdf

BOT_TOKEN = '8769728583:AAGlxs1A1mIN9bhkTueMu_z71ZO1VLTfUM8'
ALLOWED_USER_ID = 418375683

logging.basicConfig(format='%(asctime)s - %(message)s', level=logging.INFO)
logger = logging.getLogger(__name__)

class HealthHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"OK")
    def do_HEAD(self):
        self.send_response(200)
        self.end_headers()
    def log_message(self, format, *args):
        pass

def start_health_server():
    port = int(os.environ.get("PORT", 10000))
    server = HTTPServer(('0.0.0.0', port), HealthHandler)
    server.serve_forever()

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if ALLOWED_USER_ID and update.effective_user.id != ALLOWED_USER_ID:
        return
    await update.message.reply_text("Привет! Отправь Excel-файл (.xlsx) с отчётом, и я верну PDF с анализом.")

async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Отправь Excel-файл (.xlsx) - получишь PDF с анализом подразделений и регионов.")

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if ALLOWED_USER_ID and update.effective_user.id != ALLOWED_USER_ID:
        return
    doc = update.message.document
    if not doc.file_name.lower().endswith(('.xlsx', '.xlsm')):
        await update.message.reply_text("Пришли файл .xlsx")
        return
    status_msg = await update.message.reply_text("Обрабатываю...")
    try:
        with tempfile.TemporaryDirectory() as tmp:
            xlsx_path = os.path.join(tmp, 'input.xlsx')
            pdf_path = os.path.join(tmp, 'report.pdf')
            file = await doc.get_file()
            await file.download_to_drive(xlsx_path)
            data = parse_report(xlsx_path)
            generate_pdf(data, pdf_path)
            date_str = data['meta'].get('date', '').replace('.', '-')
            time_str = data['meta'].get('time', '').replace(':', '-')
            with open(pdf_path, 'rb') as f:
                await update.message.reply_document(
                    document=f,
                    filename=f"Анализ_{date_str}_{time_str}.pdf",
                    caption=f"Готово! {len(data['stores'])} магазинов, {len(data['podr'])} подразделений, {len(data['totals'])} регионов"
                )
            await status_msg.delete()
    except Exception as e:
        logger.exception("Ошибка")
        await status_msg.edit_text(f"Ошибка: {e}")

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if ALLOWED_USER_ID and update.effective_user.id != ALLOWED_USER_ID:
        return
    await update.message.reply_text("Пришли Excel-файл (.xlsx). /help - справка.")

def main():
    threading.Thread(target=start_health_server, daemon=True).start()
    print("Бот запущен!")
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler('start', start))
    app.add_handler(CommandHandler('help', help_cmd))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    app.run_polling()

if __name__ == '__main__':
    main()
