import telebot
from telebot import types
import openpyxl
import os

TOKEN = '7701145484:AAGc9fPzzhuhwNW8MQebVXc8SUZJYQtg7es'
bot = telebot.TeleBot(TOKEN)

USER_STATE = {}

STUDENT_FILE = 'students.xlsx'
TEACHER_FILE = 'teachers.xlsx'
PROJECTS_FILE = 'projects.xlsx'
PROPOSED_PROJECTS_FILE = 'proposed_projects.xlsx'

def create_file_if_missing(file, headers):
    try:
        openpyxl.load_workbook(file)
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(headers)
        wb.save(file)

create_file_if_missing(STUDENT_FILE, ["ID", "Фамилия", "Имя", "Группа"])
create_file_if_missing(TEACHER_FILE, ["ID", "Фамилия", "Имя", "Предмет"])
create_file_if_missing(PROJECTS_FILE, ["ID Преподавателя", "Название проекта", "Фамилия студента",
                                      "Имя студента", "Группа студента", "Статус", "Оценка"])
create_file_if_missing(PROPOSED_PROJECTS_FILE, ["ID Студента", "Название проекта", "Фамилия студента",
                                                "Имя студента", "Группа студента", "ID Преподавателя"])

def is_user_registered(user_id):
    return is_user_in_file(user_id, STUDENT_FILE) or is_user_in_file(user_id, TEACHER_FILE)

def is_teacher(user_id):
    try:
        wb = openpyxl.load_workbook(TEACHER_FILE)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] == user_id:
                return True
        return False
    except Exception as e:
        print(f"Ошибка проверки роли преподавателя: {e}")
        return False

def is_student(user_id):
    try:
        wb = openpyxl.load_workbook(STUDENT_FILE)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] == user_id:
                return True
        return False
    except Exception as e:
        print(f"Ошибка проверки роли студента: {e}")
        return False

def is_user_in_file(user_id, file):
    try:
        wb = openpyxl.load_workbook(file)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            if str(row[0]) == str(user_id):
                return True
        return False
    except FileNotFoundError:
        return False

@bot.message_handler(commands=['start'])
def start(message):
    user_id = message.from_user.id
    if is_user_registered(user_id):
        send_main_menu(user_id)
    else:
        markup = types.InlineKeyboardMarkup()
        markup.add(types.InlineKeyboardButton("Зарегистрироваться как студент", callback_data="register_student"))
        markup.add(types.InlineKeyboardButton("Зарегистрироваться как преподаватель", callback_data="register_teacher"))
        bot.send_message(user_id, "Привет! Я помогу тебе организовать работу с проектами. Выберите роль:", reply_markup=markup)

def send_main_menu(user_id):
    role = get_user_role(user_id)
    if role == 'student':
        send_student_menu(user_id)
    elif role == 'teacher':
        send_teacher_menu(user_id)

def get_user_role(user_id):
    if is_user_in_file(user_id, STUDENT_FILE):
        return 'student'
    elif is_user_in_file(user_id, TEACHER_FILE):
        return 'teacher'
    return None

def send_student_menu(user_id):
    markup = types.InlineKeyboardMarkup()
    markup.add(types.InlineKeyboardButton("Связаться с преподавателем", callback_data="contact_teacher"))
    markup.add(types.InlineKeyboardButton("Предложить тему проекта", callback_data="suggest_project"))
    markup.add(types.InlineKeyboardButton("Мои проекты", callback_data="my_projects"))
    bot.send_message(user_id, "Добро пожаловать, студент! Выберите действие:", reply_markup=markup)

def send_teacher_menu(user_id):
    markup = types.InlineKeyboardMarkup()
    markup.add(types.InlineKeyboardButton("Добавить проект", callback_data="add_project"))
    markup.add(types.InlineKeyboardButton("Поиск проекта", callback_data="search_project"))
    markup.add(types.InlineKeyboardButton("Скачать отчет", callback_data="download_report"))
    markup.add(types.InlineKeyboardButton("Изменить статус проекта", callback_data="change_status"))
    markup.add(types.InlineKeyboardButton("Оценить проект", callback_data="evaluate_project"))
    bot.send_message(user_id, "Добро пожаловать, преподаватель! Выберите действие:", reply_markup=markup)


@bot.callback_query_handler(func=lambda call: call.data in ["register_student", "register_teacher"])
def register_user(call):
    user_id = call.from_user.id
    USER_STATE[user_id] = {'role': call.data.split("_")[1]}
    bot.send_message(user_id, "Введите вашу фамилию:")
    bot.register_next_step_handler(call.message, get_last_name)

def get_last_name(message):
    user_id = message.from_user.id
    USER_STATE[user_id]['last_name'] = message.text.strip()
    bot.send_message(user_id, "Введите ваше имя:")
    bot.register_next_step_handler(message, get_first_name)

def get_first_name(message):
    user_id = message.from_user.id
    USER_STATE[user_id]['first_name'] = message.text.strip()

    role = USER_STATE[user_id]['role']
    if role == "student":
        bot.send_message(user_id, "Введите вашу группу:")
        bot.register_next_step_handler(message, finalize_student_registration)
    elif role == "teacher":
        bot.send_message(user_id, "Введите ваш предмет:")
        bot.register_next_step_handler(message, finalize_teacher_registration)

def finalize_student_registration(message):
    user_id = message.from_user.id
    group = message.text.strip()

    try:
        wb = openpyxl.load_workbook(STUDENT_FILE)
        ws = wb.active
        ws.append([user_id, USER_STATE[user_id]['last_name'], USER_STATE[user_id]['first_name'], group])
        wb.save(STUDENT_FILE)
        bot.send_message(user_id, "Регистрация завершена! Вы зарегистрированы как студент.")
        send_student_menu(user_id)
    except Exception as e:
        bot.send_message(user_id, f"Ошибка при сохранении данных: {e}")
    finally:
        USER_STATE.pop(user_id, None)

def finalize_teacher_registration(message):
    user_id = message.from_user.id
    subject = message.text.strip()

    try:
        wb = openpyxl.load_workbook(TEACHER_FILE)
        ws = wb.active
        ws.append([user_id, USER_STATE[user_id]['last_name'], USER_STATE[user_id]['first_name'], subject])
        wb.save(TEACHER_FILE)
        bot.send_message(user_id, "Регистрация завершена! Вы зарегистрированы как преподаватель.")
        send_teacher_menu(user_id)
    except Exception as e:
        bot.send_message(user_id, f"Ошибка при сохранении данных: {e}")
    finally:
        USER_STATE.pop(user_id, None)

def initialize_projects_file():
    if not os.path.exists(PROJECTS_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Projects"
        ws.append(["Project ID", "Title", "Description", "Teacher ID", "Student ID", "Status"])
        wb.save(PROJECTS_FILE)

@bot.callback_query_handler(func=lambda call: call.data == "add_project")
def add_project_handler(call):
    user_id = call.from_user.id
    if not is_teacher(user_id):
        bot.send_message(user_id, "Ошибка! Вы не преподаватель.")
        return
    bot.send_message(user_id, "Введите название проекта:")
    bot.register_next_step_handler(call.message, get_project_title)

def get_project_title(message):
    user_id = message.from_user.id
    project_title = message.text.strip()
    bot.send_message(user_id, "Введите описание проекта:")
    bot.register_next_step_handler(message, get_project_description, project_title)

def get_project_description(message, project_title):
    user_id = message.from_user.id
    project_description = message.text.strip()

    try:
        wb = openpyxl.load_workbook(PROJECTS_FILE)
        ws = wb.active
        ws.append([user_id, project_title, project_description, 'active'])
        wb.save(PROJECTS_FILE)
        bot.send_message(user_id, f"Проект '{project_title}' успешно добавлен!")
    except Exception as e:
        bot.send_message(user_id, f"Ошибка при добавлении проекта: {e}")

@bot.callback_query_handler(func=lambda call: call.data == "search_project")
def search_project_handler(call):
    user_id = call.from_user.id
    if not is_teacher(user_id):
        bot.send_message(user_id, "Ошибка! Вы не преподаватель.")
        return
    bot.send_message(user_id, "Введите ID, название или статус проекта для поиска:")
    bot.register_next_step_handler(call.message, search_project)

def search_project(message):
    user_id = message.from_user.id
    search_query = message.text.strip()

    try:
        wb = openpyxl.load_workbook(PROJECTS_FILE)
        ws = wb.active
        found_projects = []

        for row in ws.iter_rows(min_row=2, values_only=True):
            if (str(search_query) in str(row[0])) or (search_query.lower() in str(row[1]).lower()) or (search_query.lower() in str(row[5]).lower()):
                found_projects.append(row)

        if found_projects:
            result_message = "Найденные проекты:\n"
            for project in found_projects:
                result_message += f"ID проекта: {project[0]}, Название: {project[1]}, Статус: {project[5]}\n"
            bot.send_message(user_id, result_message)
        else:
            bot.send_message(user_id, "Проекты по вашему запросу не найдены.")
    except Exception as e:
        bot.send_message(user_id, f"Ошибка при поиске проектов: {e}")


def send_student_menu(user_id):
    markup = types.InlineKeyboardMarkup()
    markup.add(types.InlineKeyboardButton("Связаться с преподавателем", callback_data="contact_teacher"))
    markup.add(types.InlineKeyboardButton("Предложить тему проекта", callback_data="suggest_project"))
    markup.add(types.InlineKeyboardButton("Мои проекты", callback_data="my_projects"))
    bot.send_message(user_id, "Добро пожаловать, студент! Выберите действие:", reply_markup=markup)

def get_student_group(user_id):
    try:
        wb = openpyxl.load_workbook("students.xlsx")
        ws = wb.active

        # Ищем строку, в которой находится student_id
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] == user_id:  # Предполагаем, что student_id находится в первой колонке
                return row[2]  # Предполагаем, что группа находится в третьей колонке

        return "Группа не указана"  # Если не нашли, возвращаем дефолтное значение
    except Exception as e:
        print(f"Ошибка при загрузке группы: {e}")
        return "Группа не указана"

@bot.callback_query_handler(func=lambda call: call.data == "my_projects")
def my_projects(call):
    user_id = call.from_user.id
    if not is_student(user_id):  # Проверяем, является ли пользователь студентом
        bot.send_message(user_id, "Ошибка! Вы не студент.")
        return

    student_group = get_student_group(user_id)  # Получаем группу студента

    # Получаем проекты студента из файла
    try:
        wb = openpyxl.load_workbook(PROJECTS_FILE)
        ws = wb.active
        student_projects = []

        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[2] == user_id:  # Если студент связан с проектом
                project_info = f"Название: {row[1]}, Преподаватель: {row[0]}, Статус: {row[5]}, Оценка: {row[6]}, Группа: {student_group}"
                student_projects.append(project_info)

        if student_projects:
            # Отправляем список проектов
            bot.send_message(user_id, "\n\n".join(student_projects))
        else:
            bot.send_message(user_id, "У вас нет проектов.")
    except Exception as e:
        bot.send_message(user_id, f"Ошибка при загрузке проектов: {e}")

@bot.callback_query_handler(func=lambda call: call.data == "change_status")
def change_status(call):
    user_id = call.from_user.id
    if not is_teacher(user_id):
        bot.send_message(user_id, "Ошибка! Вы не преподаватель.")
        return
    bot.send_message(user_id, "Введите ID проекта, для которого хотите изменить статус:")
    bot.register_next_step_handler(call.message, get_project_id_for_status)


def get_project_id_for_status(message):
    user_id = message.from_user.id
    try:
        project_id = int(message.text.strip())
        # Проверим, существует ли проект с таким ID
        wb = openpyxl.load_workbook(PROJECTS_FILE)
        ws = wb.active
        project_found = False
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] == project_id:
                project_found = True
                break
        if project_found:
            bot.send_message(user_id, "Введите новый статус проекта (например, 'Завершен', 'Отменен'):")
            bot.register_next_step_handler(message, set_project_status, project_id)
        else:
            bot.send_message(user_id, "Проект с таким ID не найден.")
    except ValueError:
        bot.send_message(user_id, "Некорректный ID проекта. Попробуйте еще раз.")


def set_project_status(message, project_id):
    user_id = message.from_user.id
    new_status = message.text.strip()

    try:
        wb = openpyxl.load_workbook(PROJECTS_FILE)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=False):
            if row[0].value == project_id:
                row[5].value = new_status  # Обновляем статус проекта
                wb.save(PROJECTS_FILE)
                bot.send_message(user_id, f"Статус проекта с ID {project_id} успешно обновлен на '{new_status}'.")
                break
    except Exception as e:
        bot.send_message(user_id, f"Ошибка при обновлении статуса проекта: {e}")

@bot.callback_query_handler(func=lambda call: call.data == "evaluate_project")
def evaluate_project(call):
    user_id = call.from_user.id
    if not is_teacher(user_id):
        bot.send_message(user_id, "Ошибка! Вы не преподаватель.")
        return
    bot.send_message(user_id, "Введите ID проекта, который хотите оценить:")
    bot.register_next_step_handler(call.message, get_project_id_for_evaluation)

def get_project_id_for_evaluation(message):
    user_id = message.from_user.id
    try:
        project_id = int(message.text.strip())
        # Проверим, существует ли проект с таким ID
        wb = openpyxl.load_workbook(PROJECTS_FILE)
        ws = wb.active
        project_found = False
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] == project_id:
                project_found = True
                break
        if project_found:
            bot.send_message(user_id, "Введите оценку для проекта:")
            bot.register_next_step_handler(message, set_project_evaluation, project_id)
        else:
            bot.send_message(user_id, "Проект с таким ID не найден.")
    except ValueError:
        bot.send_message(user_id, "Некорректный ID проекта. Попробуйте еще раз.")

def set_project_evaluation(message, project_id):
    user_id = message.from_user.id
    evaluation = message.text.strip()

    try:
        wb = openpyxl.load_workbook(PROJECTS_FILE)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=False):
            if row[0].value == project_id:
                row[6].value = evaluation  # Обновляем оценку проекта
                wb.save(PROJECTS_FILE)
                bot.send_message(user_id, f"Оценка проекта с ID {project_id} успешно обновлена на '{evaluation}'.")
                break
    except Exception as e:
        bot.send_message(user_id, f"Ошибка при обновлении оценки проекта: {e}")

@bot.callback_query_handler(func=lambda call: call.data == "download_report")
def download_report(call):
    user_id = call.from_user.id
    if not is_teacher(user_id):
        bot.send_message(user_id, "Ошибка! Вы не преподаватель.")
        return

    try:
        report_file = 'project_report.xlsx'
        wb = openpyxl.load_workbook(PROJECTS_FILE)
        ws = wb.active
        wb.save(report_file)

        with open(report_file, 'rb') as f:
            bot.send_document(user_id, f)
        os.remove(report_file)
    except Exception as e:
        bot.send_message(user_id, f"Ошибка при скачивании отчета: {e}")


@bot.callback_query_handler(func=lambda call: call.data == "suggest_project")
def suggest_project(call):
    user_id = call.from_user.id
    bot.send_message(user_id, "Выберите преподавателя, которому хотите предложить тему проекта:")

    markup = types.InlineKeyboardMarkup()
    try:
        wb = openpyxl.load_workbook('teachers.xlsx')
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            teacher_id = row[0]
            teacher_name = f"{row[1]} {row[2]}"
            markup.add(types.InlineKeyboardButton(teacher_name, callback_data=f"teacher_{teacher_id}"))
        bot.send_message(user_id, "Выберите преподавателя из списка:", reply_markup=markup)
    except Exception as e:
        bot.send_message(user_id, f"Ошибка при загрузке преподавателей: {e}")


@bot.callback_query_handler(func=lambda call: call.data.startswith("teacher_"))
def teacher_selected(call):
    teacher_id = call.data.split("_")[1]
    user_id = call.from_user.id

    bot.send_message(user_id, "Введите тему проекта, которую вы хотите предложить:")
    bot.register_next_step_handler(call.message, lambda msg: save_project_suggestion(msg, teacher_id))

def save_project_suggestion(message, teacher_id):
    user_id = message.from_user.id
    suggestion = message.text.strip()

    student_name = message.from_user.first_name
    student_lastname = message.from_user.last_name
    student_group = "Группа не указана"  # Здесь можно добавить код для получения группы из базы, если нужно

    try:
        # Добавляем предложенную тему проекта в файл предложенных проектов
        wb = openpyxl.load_workbook(PROPOSED_PROJECTS_FILE)
        ws = wb.active
        ws.append([user_id, suggestion, student_lastname, student_name, student_group, teacher_id])
        wb.save(PROPOSED_PROJECTS_FILE)

        bot.send_message(user_id, f"Тема проекта '{suggestion}' успешно предложена преподавателю!")
        bot.send_message(teacher_id, f"Студент {student_name} {student_lastname}, группа {student_group} предложил тему: {suggestion}\n"
                                     "Нажмите одну из кнопок ниже для одобрения или отклонения этой темы.")

        # Добавляем кнопки для преподавателя
        markup = types.InlineKeyboardMarkup()
        markup.add(types.InlineKeyboardButton("Одобрить", callback_data=f"approve_{user_id}_{suggestion}"))
        markup.add(types.InlineKeyboardButton("Не одобрить", callback_data=f"reject_{user_id}_{suggestion}"))
        bot.send_message(teacher_id, "Одобрить или отклонить тему проекта:", reply_markup=markup)
    except Exception as e:
        bot.send_message(message.from_user.id, f"Ошибка при предложении темы: {e}")


@bot.callback_query_handler(func=lambda call: call.data.startswith("approve_") or call.data.startswith("reject_"))
def approve_or_reject(call):
    # Сразу отправляем ответ на callback
    bot.answer_callback_query(call.id)

    data = call.data.split("_")

    # Проверяем, что в данных достаточно элементов (должно быть 3)
    if len(data) < 3:
        bot.send_message(call.from_user.id, "Ошибка: некорректные данные.")
        return

    action = data[0]
    user_id = data[1]
    suggestion = "_".join(data[2:])  # В случае, если название проекта состоит из нескольких слов

    if action == "approve":
        # Перемещаем проект в файл проектов
        try:
            wb = openpyxl.load_workbook(PROPOSED_PROJECTS_FILE)
            ws = wb.active
            for row in ws.iter_rows(min_row=2, values_only=False):  # Изменено на values=False для работы с ячейками
                if row[0].value == int(user_id) and row[1].value == suggestion:
                    # Одобряем проект
                    wb2 = openpyxl.load_workbook(PROJECTS_FILE)
                    ws2 = wb2.active
                    ws2.append([row[5].value, suggestion, row[2].value, row[3].value, row[4].value, "Одобрено", "Не оценено"])
                    wb2.save(PROJECTS_FILE)

                    # Удаляем из предложенных
                    ws.delete_rows(row[0].row)  # Удаляем строку
                    wb.save(PROPOSED_PROJECTS_FILE)

                    bot.send_message(call.from_user.id, f"Тема '{suggestion}' одобрена и добавлена в проекты!")
                    break
        except Exception as e:
            bot.send_message(call.from_user.id, f"Ошибка при обработке одобрения: {e}")
    elif action == "reject":
        # Отправка уведомления студенту
        bot.send_message(user_id, f"Тема проекта '{suggestion}' отклонена преподавателем.")


# Запуск бота
if __name__ == "__main__":
    bot.polling(none_stop=True)