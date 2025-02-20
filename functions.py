
from psycopg2 import sql
from database import get_connection

conn = get_connection()
cur = conn.cursor()
def add_person(last_name, first_name, father_name, group_id, type):
    try:
        query = sql.SQL("INSERT INTO people (last_name, first_name, father_name, group_id, type) VALUES (%s, %s, %s, %s, %s);")
        cur.execute(query, (last_name, first_name, father_name, group_id, type))
        conn.commit()
        print(f"Новое имя добавлено в базу данных.")
        return id
    except Exception as e:
        print(f"Ошибка при добавлении пользователя: {e}")
        return None
    finally:
        cur.close()
        conn.close()

def delete_person(id):
    try:
        query = sql.SQL("DELETE FROM people WHERE id = %s;")
        cur.execute(query, (id,))
        conn.commit()
        print(f"Пользователь удален из базы данных")
        return 1
    except Exception as e:
        print(f"Ошибка удалении пользователя: {e}")
        return None
    finally:
        cur.close()
        conn.close()

def add_group(group_name):
    try:
        query = sql.SQL("INSERT INTO groups (name) VALUES (%s);")
        cur.execute(query, (group_name,))
        conn.commit()
        print(f"Группа добавлена в базу данных")
        return 1
    except Exception as e:
        print(f"Ошибка добавления группы: {e}")
        return None
    finally:
        cur.close()
        conn.close()

def delete_group(id):
    try:
        query = sql.SQL("DELETE FROM groups WHERE id = %s;")
        cur.execute(query, (id,))
        conn.commit()
        print(f"Группа удалена из базы данных")
        return 1
    except Exception as e:
        print(f"Ошибка удаления группы: {e}")
        return None
    finally:
        cur.close()
        conn.close()

def add_grade(student_id, subject_id, teacher_id, value):
    try:
        query = sql.SQL("INSERT INTO marks (student_id, subject_id, teacher_id, value) VALUES (%s, %s, %s, %s);")
        cur.execute(query, (student_id, subject_id, teacher_id, value))
        conn.commit()
        print(f"Оценка добавлена")
        return 1
    except Exception as e:
        print(f"Ошибка добавления оценки: {e}")
        return None
    finally:
        cur.close()
        conn.close()

def delete_grade(id):
    try:
        query = sql.SQL("DELETE FROM marks WHERE id = %s;")
        cur.execute(query, (id,))
        conn.commit()
        print(f"Оценка удалена из базы данных")
        return 1
    except Exception as e:
        print(f"Ошибка удаления оценки: {e}")
        return None
    finally:
        cur.close()
        conn.close()

def add_subject(name):
    try:
        query = sql.SQL("INSERT INTO subjects (name) VALUES (%s)")
        cur.execute(query, (name,))
        conn.commit()
        print(f"Предмет добавлен в базу данных")
        return 1
    except Exception as e:
        print(f"Ошибка добавления предмета: {e}")
        return None
    finally:
        cur.close()
        conn.close()
def delete_subject(id):
    try:
        query = sql.SQL("DELETE FROM marks WHERE id = %s;")
        cur.execute(query, (id,))
        conn.commit()
        print(f"Предмет удален из базы данных")
        return 1
    except Exception as e:
        print(f"Ошибка удаления предмета: {e}")
        return None
    finally:
        cur.close()
        conn.close()

def load_teacher_grade():
    try:
        query = sql.SQL("SELECT * FROM TeacherSubjectAvgMarks;")
        conn = get_connection()  # Функция для подключения к базе данных
        cur = conn.cursor()
        cur.execute(query)
        rows = cur.fetchall()  # Извлечение всех строк результата
        columns = [desc[0] for desc in cur.description]  # Заголовки столбцов
        print(f"Представление успешно загружено.")
        return rows, columns
    except Exception as e:
        print(f"Ошибка при выводе представления: {e}")
        return None, None
    finally:
        cur.close()
        conn.close()

def load_year_grade():
    try:
        query = sql.SQL("SELECT * FROM YearlyAvgMarks;")
        conn = get_connection()  # Функция для подключения к базе данных
        cur = conn.cursor()
        cur.execute(query)
        rows = cur.fetchall()  # Извлечение всех строк результата
        columns = [desc[0] for desc in cur.description]  # Заголовки столбцов
        print(f"Представление успешно загружено.")
        return rows, columns
    except Exception as e:
        print(f"Ошибка при выводе представления: {e}")
        return None, None
    finally:
        cur.close()
        conn.close()
def load_year_avg_grade(year):
    try:
        query = sql.SQL("SELECT * FROM SubjectsWithAvgMarksForYear(%s);")
        conn = get_connection()  # Функция для подключения к базе данных
        cur = conn.cursor()
        cur.execute(query, (year,))
        rows = cur.fetchall()  # Извлечение всех строк результата
        columns = [desc[0] for desc in cur.description]  # Заголовки столбцов
        print(f"Таблица успешно загружена.")
        return rows, columns
    except Exception as e:
        print(f"Ошибка при выводе таблицы: {e}")
        return None, None
    finally:
        cur.close()
        conn.close()

def load_years_avg_grade(year1, year2):
    try:
        query = sql.SQL("SELECT * FROM get_avg_marks_by_time_interval(%s, %s);")
        conn = get_connection()  # Функция для подключения к базе данных
        cur = conn.cursor()
        cur.execute(query, (year1, year2, ))
        rows = cur.fetchall()  # Извлечение всех строк результата
        columns = [desc[0] for desc in cur.description]  # Заголовки столбцов
        print(f"Таблица успешно загружена.")
        return rows, columns
    except Exception as e:
        print(f"Ошибка при выводе таблицы: {e}")
        return None, None
    finally:
        cur.close()
        conn.close()

def load_teachers():
    try:
        query = sql.SQL("SELECT * FROM teachers;")
        conn = get_connection()
        cur = conn.cursor()
        cur.execute(query)
        rows = cur.fetchall()
        columns = [desc[0] for desc in cur.description]
        print(f"Представление успешно загружено.")
        return rows, columns
    except Exception as e:
        print(f"Ошибка при выводе представления: {e}")
        return None, None
    finally:
        cur.close()
        conn.close()

def load_students_and_groups():
    try:
        query = sql.SQL("SELECT * FROM students_and_groups;")
        conn = get_connection()
        cur = conn.cursor()
        cur.execute(query)
        rows = cur.fetchall()
        columns = [desc[0] for desc in cur.description]
        print(f"Представление успешно загружено.")
        return rows, columns
    except Exception as e:
        print(f"Ошибка при выводе представления: {e}")
        return None, None
    finally:
        cur.close()
        conn.close()

def load_groups():
    try:
        query = sql.SQL("SELECT * FROM st_groups;")
        conn = get_connection()
        cur = conn.cursor()
        cur.execute(query)
        rows = cur.fetchall()
        columns = [desc[0] for desc in cur.description]
        print(f"Представление успешно загружено.")
        return rows, columns
    except Exception as e:
        print(f"Ошибка при выводе представления: {e}")
        return None, None
    finally:
        cur.close()
        conn.close()

def load_subjects():
    try:
        query = sql.SQL("SELECT * FROM all_subjects;")
        conn = get_connection()
        cur = conn.cursor()
        cur.execute(query)
        rows = cur.fetchall()
        columns = [desc[0] for desc in cur.description]
        print(f"Представление успешно загружено.")
        return rows, columns
    except Exception as e:
        print(f"Ошибка при выводе представления: {e}")
        return None, None
    finally:
        cur.close()
        conn.close()

def load_students_and_marks():
    try:
        query = sql.SQL("SELECT * FROM students_with_marks;")
        conn = get_connection()
        cur = conn.cursor()
        cur.execute(query)
        rows = cur.fetchall()
        columns = [desc[0] for desc in cur.description]
        print(f"Представление успешно загружено.")
        return rows, columns
    except Exception as e:
        print(f"Ошибка при выводе представления: {e}")
        return None, None
    finally:
        cur.close()
        conn.close()

def load_subjects_and_marks():
    try:
        query = sql.SQL("SELECT * FROM subjects_with_average_marks;")
        conn = get_connection()
        cur = conn.cursor()
        cur.execute(query)
        rows = cur.fetchall()
        columns = [desc[0] for desc in cur.description]
        print(f"Представление успешно загружено.")
        return rows, columns
    except Exception as e:
        print(f"Ошибка при выводе представления: {e}")
        return None, None
    finally:
        cur.close()
        conn.close()

def load_student_grades(id):
    try:
        query = sql.SQL("SELECT * FROM get_student_performance_by_id(%s);")
        conn = get_connection()
        cur = conn.cursor()
        cur.execute(query, (id,))
        rows = cur.fetchall()
        columns = [desc[0] for desc in cur.description]
        print(f"Таблица успешно загружена.")
        return rows, columns
    except Exception as e:
        print(f"Ошибка при выводе таблицы: {e}")
        return None, None
    finally:
        cur.close()
        conn.close()

def load_years_subject(id):
    try:
        query = sql.SQL("SELECT * FROM get_avg_mark_by_subject(%s);")
        conn = get_connection()
        cur = conn.cursor()
        cur.execute(query, (id,))
        rows = cur.fetchall()
        columns = [desc[0] for desc in cur.description]
        print(f"Таблица успешно загружена.")
        return rows, columns
    except Exception as e:
        print(f"Ошибка при выводе таблицы: {e}")
        return None, None
    finally:
        cur.close()
        conn.close()