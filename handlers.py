from telebot import types
from excel_manager import (
    add_project_to_excel,
    add_user_to_excel,
    get_projects_by_student,
    get_all_projects,
    get_user_role
)

# ID администратора
ADMIN_ID = 1008919333
# ID преподавателя
TEACHER_ID = 1008919333


def register_handlers(bot):
    @bot.message_handler(commands=['start'])
    def start(message):
        user_id = message.chat.id
        username = message.from_user.username

        # Проверка, является ли пользователь администратором
        if user_id == ADMIN_ID:
            role = "admin"
        else:
            role = get_user_role(user_id)

        # Создаем кнопки для регистрации
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        btn1 = types.KeyboardButton("Зарегистрироваться как студент")
        btn2 = types.KeyboardButton("Зарегистрироваться как преподаватель")
        markup.add(btn1, btn2)

        bot.send_message(
            message.chat.id,
            f"Привет, {username}! Ваш Telegram ID: {user_id}. Выберите вашу роль:",
            reply_markup=markup
        )

    @bot.message_handler(
        func=lambda message: message.text in ["Зарегистрироваться как студент", "Зарегистрироваться как преподаватель"])
    def register_user(message):
        user_id = message.chat.id
        role = "student" if message.text == "Зарегистрироваться как студент" else "teacher"
        username = message.from_user.username

        # Если это администратор, можно регистрировать как студент или преподаватель
        if user_id == ADMIN_ID:
            role = "admin"
            add_user_to_excel(user_id, username, role)
            bot.send_message(message.chat.id, f"Вы успешно зарегистрированы как {role}!")
            show_menu(bot, message, role)
            return

        if role == "teacher" and user_id != TEACHER_ID:
            bot.send_message(message.chat.id,
                             "Вы не можете зарегистрироваться как преподаватель. Только администратор может зарегистрировать преподавателя.")
            return

        add_user_to_excel(user_id, username, role)
        bot.send_message(message.chat.id, f"Вы успешно зарегистрировались как {role}!")
        show_menu(bot, message, role)

    def show_menu(bot, message, role):
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)

        if role == "student":
            btn1 = types.KeyboardButton("Добавить проект")
            btn2 = types.KeyboardButton("Мои проекты")
            markup.add(btn1, btn2)
            bot.send_message(message.chat.id, "Выберите действие:", reply_markup=markup)
        elif role == "teacher" or role == "admin":
            btn1 = types.KeyboardButton("Просмотреть проекты")
            btn2 = types.KeyboardButton("Добавить проект")
            markup.add(btn1, btn2)
            bot.send_message(message.chat.id, "Выберите действие:", reply_markup=markup)

    @bot.message_handler(func=lambda message: True)
    def handle_user_actions(message):
        user_id = message.chat.id
        role = get_user_role(user_id)

        if not role:
            bot.send_message(message.chat.id,
                             "Вы не зарегистрированы. Пожалуйста, выполните регистрацию с помощью команды /start.")
            return

        if message.text == "Добавить проект" and (role == "teacher" or role == "admin"):
            bot.send_message(message.chat.id, "Введите название проекта:")
            bot.register_next_step_handler(message, get_project_title)
        elif message.text == "Мои проекты" and role == "student":
            my_projects_command(message)
        elif message.text == "Просмотреть проекты" and (role == "teacher" or role == "admin"):
            view_projects_command(message)
        else:
            bot.send_message(message.chat.id, "Действие недоступно или команда не распознана.")

    def get_project_title(message):
        title = message.text
        bot.send_message(message.chat.id, "Введите описание проекта:")
        bot.register_next_step_handler(message, lambda msg: save_project(msg, title))

    def save_project(message, title):
        description = message.text
        teacher_id = TEACHER_ID if message.chat.id != ADMIN_ID else None  # Преподаватель будет только для админа, для обычных преподавателей
        add_project_to_excel(
            student_id=message.chat.id,
            username=message.from_user.username,
            teacher_id=teacher_id,  # Устанавливаем преподавателя после одобрения
            title=title,
            description=description,
            status="pending"
        )
        bot.send_message(message.chat.id, "Проект успешно добавлен в базу!")

    def my_projects_command(message):
        student_id = message.chat.id
        projects = get_projects_by_student(student_id)

        if not projects:
            bot.send_message(message.chat.id, "У вас пока нет проектов.")
        else:
            for project in projects:
                bot.send_message(
                    message.chat.id,
                    f"Название: {project['title']}\n"
                    f"Описание: {project['description']}\n"
                    f"Статус: {project['status']}"
                )

    def view_projects_command(message):
        projects = get_all_projects()

        if not projects:
            bot.send_message(message.chat.id, "Проекты отсутствуют.")
        else:
            for project in projects:
                bot.send_message(
                    message.chat.id,
                    f"Студент: {project['student_username']} (ID: {project['student_id']})\n"
                    f"Название: {project['title']}\n"
                    f"Описание: {project['description']}\n"
                    f"Статус: {project['status']}\n"
                    f"Назначенный преподаватель: {project['teacher_id'] or 'Не назначен'}"
                )
