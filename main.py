import telebot
from config import TOKEN
from handlers import register_handlers
from excel_manager import initialize_excel

# Инициализация бота
bot = telebot.TeleBot(TOKEN)

# Инициализация Excel файла
initialize_excel()

# Регистрация обработчиков
register_handlers(bot)

# Запуск бота
bot.polling(none_stop=True)

