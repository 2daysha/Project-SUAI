from excel_manager import get_projects_by_teacher, add_message_to_excel, get_new_messages_for_user, mark_message_as_read
from telebot import types

# Функция для отправки сообщения студенту
def send_message_to_student(bot, teacher_id, student_id, message_text):
    try:
        bot.send_message(student_id, f"Сообщение от преподавателя:\n{message_text}")
        bot.send_message(teacher_id, "Сообщение успешно отправлено студенту.")
    except Exception as e:
        bot.send_message(teacher_id, f"Ошибка при отправке сообщения: {str(e)}")

# Функция для проверки наличия новых сообщений
def check_for_new_messages(bot, user_id):
    new_messages = get_new_messages_for_user(user_id)
    if new_messages:
        for msg in new_messages:
            bot.send_message(user_id, f"Новое сообщение от {msg['sender_username']}:\n{msg['message_text']}")
            mark_message_as_read(msg['message_id'])
    else:
        bot.send_message(user_id, "Нет новых сообщений.")

# Функция для отправки ответа на сообщение
def send_reply(bot, original_message, reply_text):
    sender_id = original_message['sender_id']
    receiver_id = original_message['receiver_id']

    # Сохраняем ответ в базе
    add_message_to_excel(sender_id=sender_id, receiver_id=receiver_id, message_text=reply_text, project_id=original_message['project_id'])

    # Отправляем ответ
    bot.send_message(receiver_id, f"Ответ от {original_message['sender_username']}:\n{reply_text}")
    bot.send_message(sender_id, "Ваш ответ успешно отправлен.")

def send_message_to_student_step(message, bot):
    """Отправка сообщения студенту по ID. Запрашивает проект и сообщение у преподавателя."""
    try:
        # Получаем проекты преподавателя
        projects = get_projects_by_teacher(message.chat.id)

        if not projects:
            bot.send_message(message.chat.id, "У вас нет проектов.")
            return

        # Создаем клавиатуру для выбора проекта
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
        for project in projects:
            markup.add(types.KeyboardButton(project['title']))

        bot.send_message(
            message.chat.id,
            "Выберите проект, чтобы отправить сообщение студенту:",
            reply_markup=markup,
        )

        # Переход к обработчику выбора проекта
        bot.register_next_step_handler(message, handle_message_to_student, bot, projects)

    except Exception as e:
        bot.send_message(message.chat.id, f"Ошибка: {str(e)}")

def handle_message_to_student(message, bot, projects):
    """
    Обрабатывает выбор проекта преподавателем.
    Запрашивает текст сообщения.
    """
    try:
        # Ищем проект по названию
        selected_project = next((proj for proj in projects if proj['title'] == message.text), None)

        if selected_project:
            bot.send_message(message.chat.id, "Введите сообщение для студента:")
            bot.register_next_step_handler(
                message, send_message_to_student_final, bot, selected_project
            )
        else:
            bot.send_message(message.chat.id, "Неверный выбор. Попробуйте еще раз.")
            send_message_to_student_step(message, bot)  # Перезапуск выбора проекта
    except Exception as e:
        bot.send_message(message.chat.id, f"Ошибка: {str(e)}")

def send_message_to_student_final(message, bot, project):
    """
    Отправляет сообщение студенту, связанное с выбранным проектом.
    Сохраняет сообщение в базе.
    """
    try:
        student_id = project['student_id']

        if not student_id:
            bot.send_message(message.chat.id, "Ошибка: студент не найден для этого проекта.")
            return

        # Сохраняем сообщение в базе
        add_message_to_excel(
            sender_id=message.chat.id,
            receiver_id=student_id,
            message_text=message.text,
            project_id=project['id'],
        )

        # Отправляем сообщение студенту
        bot.send_message(student_id, f"Сообщение от преподавателя:\n{message.text}")
        bot.send_message(message.chat.id, "Сообщение успешно отправлено студенту.")

    except Exception as e:
        bot.send_message(message.chat.id, f"Ошибка при отправке сообщения: {str(e)}")

def handle_reply_to_message(message, bot, original_message):
    """
    Обрабатывает ответ студента на сообщение преподавателя.
    """
    try:
        reply_text = message.text
        sender_id = message.chat.id
        receiver_id = original_message['sender_id']

        # Сохраняем ответ в базе
        add_message_to_excel(
            sender_id=sender_id,
            receiver_id=receiver_id,
            message_text=reply_text,
            project_id=original_message['project_id'],
        )

        # Отправляем ответ
        bot.send_message(
            receiver_id, f"Ответ от {message.from_user.username}:\n{reply_text}"
        )
        bot.send_message(sender_id, "Ваш ответ успешно отправлен.")

    except Exception as e:
        bot.send_message(message.chat.id, f"Ошибка: {str(e)}")
