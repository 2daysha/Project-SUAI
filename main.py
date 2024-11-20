from bot_instance import bot  # Импортируем bot
from handlers import register_handlers  # Импортируем функцию регистрации хендлеров
from excel_manager import initialize_excel
initialize_excel()
from handlers import register_handlers
register_handlers(bot)

# Регистрируем хендлеры
register_handlers(bot)

# Запуск бота
if __name__ == "__main__":
    bot.polling(none_stop=True)
