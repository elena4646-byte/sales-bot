"""
Telegram-бот для анализа отчёта по часу продаж.
Принимает Excel-файл, возвращает PDF с анализом.
"""
import os
import logging
import tempfile
import asyncio
import nest_asyncio
from telegram import Update
from telegram.ext import Application, MessageHandler, CommandHandler, filters, ContextTypes

from report_analyzer import parse_report, generate_pdf

# ===== НАСТРОЙКИ =====
BOT_TOKEN = '8769728583:AAGlxs1A1mIN9bhkTueMu_z71ZO1VLTfUM8'

# Ваш Telegram ID
ALLOWED_USER_ID = 418375683

logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if ALLOWED_USER_ID and update.effective_user.id != ALLOWED_USER_ID:
        await update.message.reply_text("Нет доступа")
        return
    await update.message.reply_text(
        "Привет!\n\n"
        "Отправь мне Excel-файл с отчётом по часу продаж, "
        "и я верну PDF с анализом.\n\n"
        "Просто прикрепи .xlsx файл."
    )


async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Как пользоваться:\n\n"
        "1. Отправь Excel-файл (.xlsx)\n"
        "2. Получишь PDF с анализом\n\n"
        "В PDF:\n"
        "- Сравнение подразделений КЗ\n"
        "- Сравнение регионов сети\n"
        "- Карточки по подразделениям"
    )


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if ALLOWED_USER_ID and update.effective_user.id != ALLOWED_USER_ID:
        await update.message.reply_text("Нет доступа")
        return

    doc = update.message.document
    if not doc.file_name.lower().endswith(('.xlsx', '.xlsm')):
        await update.message.reply_text("Пришли файл формата .xlsx")
        return

    status_msg = await update.message.reply_text("Обрабатываю отчёт...")

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
            out_name = f"Анализ_{date_str}_{time_str}.pdf"

            with open(pdf_path, 'rb') as f:
                await update.message.reply_document(
                    document=f,
                    filename=out_name,
                    caption=(
                        f"Готово!\n"
                        f"{len(data['stores'])} магазинов, "
                        f"{len(data['podr'])} подразделений, "
                        f"{len(data['totals'])} регионов"
                    )
                )

            await status_msg.delete()

    except Exception as e:
        logger.exception("Ошибка обработки")
        await status_msg.edit_text(f"Ошибка: {e}")


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if ALLOWED_USER_ID and update.effective_user.id != ALLOWED_USER_ID:
        return
    await update.message.reply_text(
        "Пришли Excel-файл (.xlsx) с отчётом.\n"
        "/help - справка."
    )


def main():
    nest_asyncio.apply()

    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler('start', start))
    app.add_handler(CommandHandler('help', help_cmd))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))

    print("Бот запущен!")
    app.run_polling()


if __name__ == '__main__':
    main()
