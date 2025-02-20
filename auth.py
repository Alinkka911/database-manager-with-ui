
import bcrypt
from database import get_connection
from psycopg2 import sql
import database


def add_user(username, password, role):
    conn = get_connection()
    cur = conn.cursor()
    try:
        password_hash = bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')

        query = sql.SQL("""
            INSERT INTO users (username, password_hash, role)
            VALUES (%s, %s, %s);
        """)
        cur.execute(query, (username, password_hash, role))
        conn.commit()
        print(f"Пользователь '{username}' добавлен.")
    except ValueError as message:
        print(f"Ошибка: Пользователь с именем '{username}' уже существует.")
    except Exception as e:
        print(f"Ошибка при добавлении пользователя: {e}")
    finally:
        cur.close()
        conn.close()
def authenticate_user(username, password):

    query = "SELECT password_hash, role FROM users WHERE username = %s;"
    result = database.execute_query(query, (username,))

    if result:
        stored_hash, role = result[0]
        if bcrypt.checkpw(password.encode('utf-8'), stored_hash.encode('utf-8')):
            return role
    return None