import pandas as pd
from telebot import types
from excel_manager import get_projects_by_student, add_message_to_excel, FILE_NAME

# Пример Teacher ID (замените на реальный ID преподавателя)
TEACHER_ID = 1008919333

def student_menu(bot, message):
    """Главное меню студента."""
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    btn1 = types.KeyboardButton("Работа с проектом")
    btn2 = types.KeyboardButton("Связаться с преподавателем")
    markup.add(btn1, btn2)
    bot.send_message(message.chat.id, "Выберите действие:", reply_markup=markup)

def project_work_menu(bot, message):
    """Меню работы с проектами для студента."""
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    btn1 = types.KeyboardButton("Предложить тему проекта")
    btn2 = types.KeyboardButton("Мои проекты")
    btn3 = types.KeyboardButton("Связаться с преподавателем")
    markup.add(btn1, btn2, btn3)
    bot.send_message(message.chat.id, "Выберите действие:", reply_markup=markup)

def my_projects_command(bot, message):
    """Команда для отображения проектов студента."""
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

def save_proposed_topic(message, bot):
    """Сохранение предложенной темы проекта и отправка преподавателю."""
    proposed_topic = message.text
    user_id = message.chat.id
    username = message.from_user.username

    try:
        # Отправка темы преподавателю
        bot.send_message(
            TEACHER_ID,
            f"Студент {username} (ID: {user_id}) предложил тему проекта:\n{proposed_topic}"
        )
        # Уведомление студента
        bot.send_message(user_id, "Ваша тема проекта была отправлена преподавателю.")
    except Exception as e:
        bot.send_message(user_id, f"Ошибка при отправке темы: {str(e)}")

def propose_project_topic(bot, message):
    """Запрос темы проекта у студента."""
    bot.send_message(message.chat.id, "Введите вашу тему проекта:")
    bot.register_next_step_handler(message, save_proposed_topic, bot)

def handle_reply_to_message(message, bot, original_message):
    """Обработка ответа на сообщение."""
    reply_text = message.text
    sender_id = message.chat.id
    receiver_id = original_message['sender_id']

    try:
        # Сохранение ответа в базе данных
        add_message_to_excel(
            sender_id=sender_id,
            receiver_id=receiver_id,
            message_text=reply_text,
            project_id=original_message['project_id']
        )
        # Отправка ответа
        bot.send_message(receiver_id, f"Ответ от {message.from_user.username}:\n{reply_text}")
        bot.send_message(sender_id, "Ваш ответ успешно отправлен.")
    except Exception as e:
        bot.send_message(sender_id, f"Ошибка при отправке ответа: {str(e)}")

def get_user_role(user_id):
    """Возвращает роль пользователя (студент или преподаватель)."""
    try:
        # Читаем файл с пользователями
        df = pd.read_excel(FILE_NAME, sheet_name="Users")
        user = df[df['user_id'] == user_id]

        if user.empty:
            return None  # Пользователь не найден
        return user.iloc[0]['role']  # Возвращаем роль
    except Exception as e:
        print(f"Ошибка при получении роли пользователя: {e}")
        return None