import sqlite3
import pandas as pd

DB_NAME = "food_requests.db"
conn = sqlite3.connect(DB_NAME)

try:
    # Проверка подключения к БД
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tables = cursor.fetchall()
    print("Таблицы в базе данных:", tables)
    
    # Проверка данных
    cursor.execute("SELECT COUNT(*) FROM requests")
    count = cursor.fetchone()[0]
    print(f"Всего заявок: {count}")
    
    # Тестовый экспорт
    df = pd.read_sql("SELECT * FROM requests", conn)
    print(f"Данные:\n{df.head()}")
    
    df.to_excel("test_export.xlsx", index=False)
    print("Файл test_export.xlsx создан")
    
except Exception as e:
    print(f"Ошибка: {e}")
finally:
    conn.close()