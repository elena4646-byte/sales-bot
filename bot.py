"""
Telegram-бот для анализа отчёта по часу продаж.
Принимает Excel-файл → возвращает PDF с анализом сравнения подразделений и регионов.
"""
import os
import logging
import tempfile
from telegram import Update
from telegram.ext import Application, MessageHandler, CommandHandler, filters, ContextTypes

from report_analyzer import parse_report, generate_pdf

# ===== НАСТРОЙКИ =====
BOT_TOKEN = os.environ.get('BOT_TOKEN', 'ВСТАВЬТЕ_СЮДА_ТОКЕН_ОТ_BOTFATHER')
ALLOWED_USER_ID = int(os.environ.get('ALLOWED_USER_ID', '0'))  # ваш Telegram ID, 0 = без ограничения

logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Команда /start"""
    if ALLOWED_USER_ID and update.effective_user.id != ALLOWED_USER_ID:
        await update.message.reply_text("⛔ Нет доступа")
        return
    
    await update.message.reply_text(
        "👋 Привет!\n\n"
        "Отправь мне Excel-файл с отчётом по часу продаж, "
        "и я верну PDF с анализом сравнения подразделений и регионов.\n\n"
        "📎 Просто прикрепи .xlsx файл к сообщению."
    )


async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Команда /help"""
    await update.message.reply_text(
        "📖 Как пользоваться:\n\n"
        "1. Отправь мне Excel-файл отчёта (.xlsx)\n"
        "2. Через несколько секунд получишь PDF с анализом\n\n"
        "В PDF будет:\n"
        "• Сравнение подразделений (Казахстан)\n"
        "• Сравнение регионов сети\n"
        "• Анализ отклонений от эталона (Kari)\n"
        "• Выводы и рекомендации"
    )


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработка полученного файла"""
    # Проверка доступа
    if ALLOWED_USER_ID and update.effective_user.id != ALLOWED_USER_ID:
        await update.message.reply_text("⛔ Нет доступа")
        return
    
    doc = update.message.document
    if not doc.file_name.lower().endswith(('.xlsx', '.xlsm')):
        await update.message.reply_text("⚠️ Пришли файл формата .xlsx")
        return
    
    status_msg = await update.message.reply_text("⏳ Обрабатываю отчёт...")
    
    try:
        # Временные файлы
        with tempfile.TemporaryDirectory() as tmp:
            xlsx_path = os.path.join(tmp, 'input.xlsx')
            pdf_path = os.path.join(tmp, 'report.pdf')
            
            # Скачиваем Excel
            file = await doc.get_file()
            await file.download_to_drive(xlsx_path)
            
            # Парсим и генерируем PDF
            data = parse_report(xlsx_path)
            generate_pdf(data, pdf_path)
            
            # Отправляем PDF
            date_str = data['meta'].get('date', '').replace('.', '-')
            time_str = data['meta'].get('time', '').replace(':', '-')
            out_name = f"Анализ_{date_str}_{time_str}.pdf"
            
            with open(pdf_path, 'rb') as f:
                await update.message.reply_document(
                    document=f,
                    filename=out_name,
                    caption=(
                        f"✅ Готово!\n"
                        f"📊 Обработано: {len(data['stores'])} магазинов, "
                        f"{len(data['podr'])} подразделений, "
                        f"{len(data['totals'])} регионов"
                    )
                )
            
            await status_msg.delete()
    
    except Exception as e:
        logger.exception("Ошибка обработки")
        await status_msg.edit_text(f"❌ Ошибка: {e}")


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Любой текст"""
    if ALLOWED_USER_ID and update.effective_user.id != ALLOWED_USER_ID:
        return
    await update.message.reply_text(
        "📎 Пришли Excel-файл (.xlsx) с отчётом.\n"
        "Команда /help — справка."
    )


def main():
    """Запуск бота"""
    if BOT_TOKEN == 'ВСТАВЬТЕ_СЮДА_ТОКЕН_ОТ_BOTFATHER':
        print("❌ Укажите BOT_TOKEN в переменной окружения или в коде!")
        return
    
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler('start', start))
    app.add_handler(CommandHandler('help', help_cmd))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    
    print("🤖 Бот запущен. Нажми Ctrl+C для остановки.")
    app.run_polling()


if __name__ == '__main__':
    main()
