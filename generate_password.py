import random
import string
import bcrypt
def generate_password(length=12):
    characters = string.ascii_letters + string.digits
    return ''.join(random.choice(characters) for _ in range(length))

print(generate_password())


def add_user(username, role):
    # Создаём случайный пароль
    plain_password = "password123"  # Можно использовать генератор паролей

    # Генерируем хэш пароля
    password_hash = bcrypt.hashpw(plain_password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')

    # Добавляем пользователя в базу
    conn = psycopg2.connect(
        host="localhost", database="your_database", user="your_user", password="your_password"
    )
    cur = conn.cursor()
    try:
        cur.execute(
            "INSERT INTO users (username, password_hash, role) VALUES (%s, %s, %s)",
            (username, password_hash, role),
        )
        conn.commit()
        print(f"Пользователь '{username}' успешно добавлен!")
        print(f"Логин: {username}")
        print(f"Пароль: {plain_password}")
    except Exception as e:
        print(f"Ошибка: {e}")
    finally:
        cur.close()
        conn.close()