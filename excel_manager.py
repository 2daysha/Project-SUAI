from openpyxl import Workbook, load_workbook
import os

FILE_NAME = "project_manager.xlsx"


# Инициализация файла Excel
def initialize_excel():
    if not os.path.exists(FILE_NAME):
        workbook = Workbook()

        # Лист для пользователей
        sheet_users = workbook.active
        sheet_users.title = "Users"
        sheet_users.append(["Telegram ID", "Username", "Role"])  # Заголовки

        # Лист для проектов
        sheet_projects = workbook.create_sheet("Projects")
        sheet_projects.append(
            ["Student ID", "Student Username", "Teacher ID", "Title", "Description", "Status"])  # Заголовки

        workbook.save(FILE_NAME)


# Добавление пользователя
def add_user_to_excel(telegram_id, username, role):
    initialize_excel()
    workbook = load_workbook(FILE_NAME)
    sheet = workbook["Users"]

    # Проверяем, зарегистрирован ли пользователь
    for row in sheet.iter_rows(values_only=True):
        if row[0] == telegram_id:
            return  # Пользователь уже существует, ничего не делаем

    # Добавляем пользователя
    sheet.append([telegram_id, username, role])
    workbook.save(FILE_NAME)


# Добавление проекта
def add_project_to_excel(student_id, username, teacher_id, title, description, status):
    initialize_excel()
    workbook = load_workbook(FILE_NAME)
    sheet = workbook["Projects"]

    # Добавляем проект
    sheet.append([student_id, username, teacher_id, title, description, status])
    workbook.save(FILE_NAME)

def get_all_projects():
    initialize_excel()
    workbook = load_workbook(FILE_NAME)
    sheet = workbook["Projects"]

    projects = []
    for row in sheet.iter_rows(values_only=True):
        if row[0] == "Student ID":  # Пропускаем заголовок
            continue
        projects.append({
            "student_id": row[0],
            "student_username": row[1],
            "teacher_id": row[2],
            "title": row[3],
            "description": row[4],
            "status": row[5],
        })
    return projects

def get_user_role(user_id):
    """
    Получает роль пользователя по Telegram ID.
    Возвращает 'student', 'teacher' или None, если пользователь не найден.
    """
    initialize_excel()
    workbook = load_workbook(FILE_NAME)
    sheet = workbook["Users"]

    for row in sheet.iter_rows(values_only=True):
        if row[0] == "Telegram ID":  # Пропускаем заголовок
            continue
        if row[0] == user_id:  # Telegram ID совпал
            return row[2]  # Возвращаем роль (student/teacher)

    return None  # Если пользователь не найден

# Получение всех пользователей
def get_all_users():
    initialize_excel()
    workbook = load_workbook(FILE_NAME)
    sheet = workbook["Users"]
    return list(sheet.iter_rows(values_only=True))[1:]  # Пропускаем заголовок

def get_projects_by_student(student_id):
    initialize_excel()
    workbook = load_workbook(FILE_NAME)
    sheet = workbook["Projects"]

    projects = []
    for row in sheet.iter_rows(values_only=True):
        if row[0] == student_id:  # Сравниваем Student ID
            projects.append({
                "title": row[3],
                "description": row[4],
                "status": row[5],
            })
    return projects
