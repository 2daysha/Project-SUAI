from telebot import types
from excel_manager import add_user_to_excel, add_project_to_excel

def register_handlers(bot):
    @bot.message_handler(commands=['start'])
    def start(message):
        # Узнаем Telegram ID пользователя
        user_id = message.chat.id
        username = message.from_user.username

        # Создаем кнопки для регистрации
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        btn1 = types.KeyboardButton("Зарегистрироваться как студент")
        btn2 = types.KeyboardButton("Зарегистрироваться как преподаватель")
        markup.add(btn1, btn2)

        # Отправляем приветствие
        bot.send_message(
            message.chat.id,
            f"Привет, {username}! Ваш Telegram ID: {user_id}. Выберите вашу роль:",
            reply_markup=markup
        )

    @bot.message_handler(func=lambda message: message.text in ["Зарегистрироваться как студент", "Зарегистрироваться как преподаватель"])
    def register_user(message):
        # Определяем роль на основе выбора
        role = "student" if message.text == "Зарегистрироваться как студент" else "teacher"
        username = message.from_user.username
        telegram_id = message.chat.id

        # Регистрируем пользователя в Excel
        add_user_to_excel(telegram_id, username, role)

        # Подтверждение регистрации
        bot.send_message(message.chat.id, f"Вы успешно зарегистрировались как {role}!")

    @bot.message_handler(commands=['add_project'])
    def add_project_command(message):
        # Инициируем добавление проекта
        bot.send_message(message.chat.id, "Введите название проекта:")
        bot.register_next_step_handler(message, get_project_title)

    def get_project_title(message):
        # Получаем название проекта
        title = message.text
        bot.send_message(message.chat.id, "Введите описание проекта:")
        bot.register_next_step_handler(message, lambda msg: save_project(msg, title))

    def save_project(message, title):
        # Сохраняем проект
        description = message.text
        add_project_to_excel(
            student_id=message.chat.id,
            username=message.from_user.username,
            teacher_id=None,  # Учитель пока не назначен
            title=title,
            description=description,
            status="pending"
        )
        bot.send_message(message.chat.id, "Проект успешно добавлен в базу!")
