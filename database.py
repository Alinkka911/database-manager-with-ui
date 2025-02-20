import psycopg2
import configparser

def get_connection():
    config = configparser.ConfigParser()
    config.read('config.ini')
    db_config = config['database']

    try:
        conn = psycopg2.connect(
            host=db_config['host'],
            database=db_config['database'],
            user=db_config['user'],
            password=db_config['password'],
            port=db_config['port']
        )
        return conn
    except Exception as e:
        print(f"Ошибка подключения: {e}")
        return None

def execute_query(query, params=None):
    conn = get_connection()
    if conn:
        cur = conn.cursor()
        try:
            cur.execute(query, params)
            result = cur.fetchall()
            conn.commit()
            return result
        except Exception as e:
            print(f"Ошибка выполнения запроса: {e}")
        finally:
            cur.close()
            conn.close()
    return None
