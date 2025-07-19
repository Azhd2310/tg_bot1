import logging
import os
import sqlite3
import zipfile
from datetime import datetime, date
from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, ReplyKeyboardRemove
from aiogram.utils.keyboard import ReplyKeyboardBuilder
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from apscheduler.triggers.cron import CronTrigger

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Токен бота
TOKEN = "8041866728:AAGaQGnSE0JR7Oyt8fJg8aV6A_9AWxxXv_E"

# ID администратора (ваш Telegram ID)
ADMIN_ID = 189380617

# Инициализация бота и диспетчера
bot = Bot(token=TOKEN)
dp = Dispatcher()

# Настройка базы данных
DB_NAME = "food_requests.db"
EXCEL_FOLDER = os.path.join(os.getcwd(), "excel_reports")
os.makedirs(EXCEL_FOLDER, exist_ok=True)
logger.info(f"Папка для отчетов: {EXCEL_FOLDER}")

# Создание таблиц в базе данных
def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    
    # Таблица пользователей
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        telegram_id INTEGER UNIQUE,
        full_name TEXT,
        last_update DATE
    )
    ''')
    
    # Таблица заявок
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS requests (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER,
        date DATE,
        canteen TEXT,
        FOREIGN KEY(user_id) REFERENCES users(id)
    )
    ''')
    
    conn.commit()
    conn.close()
    logger.info("База данных инициализирована")

# Инициализация базы данных
init_db()

# Состояния для FSM
class Form(StatesGroup):
    waiting_for_name = State()
    waiting_for_canteen = State()

# ================= Обработчики команд =================

# Стартовая команда
@dp.message(Command("start"))
async def cmd_start(message: types.Message, state: FSMContext):
    user_id = message.from_user.id
    
    # Проверяем, есть ли пользователь в базе
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT full_name FROM users WHERE telegram_id = ?", (user_id,))
    user = cursor.fetchone()
    conn.close()
    
    if user:
        await message.answer(
            f"С возвращением, {user[0]}!\n"
            "Чтобы подать заявку на питание, используйте команду /order.\n"
            "Чтобы изменить ФИО, используйте /change_name."
        )
    else:
        await message.answer(
            "Привет! Я бот для подачи заявок на питание.\n"
            "Пожалуйста, введите ваше ФИО (полностью):"
        )
        await state.set_state(Form.waiting_for_name)

# Обработчик ввода ФИО
@dp.message(Form.waiting_for_name)
async def process_name(message: types.Message, state: FSMContext):
    full_name = message.text.strip()
    user_id = message.from_user.id
    today = date.today()
    
    # Сохраняем в базу данных
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    
    try:
        # Проверяем существование пользователя
        cursor.execute("SELECT id FROM users WHERE telegram_id = ?", (user_id,))
        user = cursor.fetchone()
        
        if user:
            # Обновляем существующего пользователя
            cursor.execute(
                "UPDATE users SET full_name = ?, last_update = ? WHERE telegram_id = ?",
                (full_name, today, user_id))
        else:
            # Добавляем нового пользователя
            cursor.execute(
                "INSERT INTO users (telegram_id, full_name, last_update) VALUES (?, ?, ?)",
                (user_id, full_name, today))
        
        conn.commit()
        await message.answer(
            f"Спасибо, {full_name}! Ваше ФИО сохранено.\n"
            "Теперь вы можете подать заявку на питание с помощью команды /order."
        )
        logger.info(f"Пользователь {user_id} сохранен: {full_name}")
    except Exception as e:
        logger.error(f"Ошибка при сохранении ФИО: {e}", exc_info=True)
        await message.answer("❌ Произошла ошибка при сохранении данных. Попробуйте еще раз.")
    finally:
        conn.close()
        await state.clear()

# Команда для изменения ФИО
@dp.message(Command("change_name"))
async def cmd_change_name(message: types.Message, state: FSMContext):
    await message.answer("Введите ваше новое ФИО:")
    await state.set_state(Form.waiting_for_name)
    logger.info(f"Пользователь {message.from_user.id} запросил изменение ФИО")

# Команда для подачи заявки
@dp.message(Command("order"))
async def cmd_order(message: types.Message, state: FSMContext):
    user_id = message.from_user.id
    
    # Проверяем, зарегистрирован ли пользователь
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT id, full_name FROM users WHERE telegram_id = ?", (user_id,))
    user = cursor.fetchone()
    
    if not user:
        conn.close()
        await message.answer("Пожалуйста, сначала зарегистрируйтесь с помощью команды /start.")
        return
    
    today = date.today()
    
    # Проверяем существующую заявку
    cursor.execute(
        "SELECT id FROM requests WHERE user_id = ? AND date = ?",
        (user[0], today))
    existing_request = cursor.fetchone()
    
    if existing_request:
        # Удаляем существующую заявку
        cursor.execute(
            "DELETE FROM requests WHERE id = ?",
            (existing_request[0],))
        conn.commit()
        await message.answer("✅ Ваша предыдущая заявка на сегодня удалена. Пожалуйста, создайте новую.")
    
    conn.close()
    
    # Создаем клавиатуру с выбором столовой
    builder = ReplyKeyboardBuilder()
    builder.add(KeyboardButton(text="Центр"))
    builder.add(KeyboardButton(text="Ястреб"))
    builder.adjust(2)
    
    await message.answer(
        f"{user[1]}, выберите столовую на сегодня:",
        reply_markup=builder.as_markup(resize_keyboard=True, one_time_keyboard=True))
    await state.set_state(Form.waiting_for_canteen)
    logger.info(f"Пользователь {user_id} начал подачу заявки")

# Обработчик выбора столовой
@dp.message(Form.waiting_for_canteen, F.text.in_(["Центр", "Ястреб"]))
async def process_canteen(message: types.Message, state: FSMContext):
    canteen = message.text
    user_id = message.from_user.id
    today = date.today()
    
    # Получаем ID пользователя
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT id FROM users WHERE telegram_id = ?", (user_id,))
    user = cursor.fetchone()
    
    if not user:
        await message.answer("❌ Ошибка: пользователь не найден. Начните с /start.")
        await state.clear()
        conn.close()
        return
    
    # Сохраняем заявку
    try:
        cursor.execute(
            "INSERT INTO requests (user_id, date, canteen) VALUES (?, ?, ?)",
            (user[0], today, canteen))
        conn.commit()
        await message.answer(
            f"✅ Заявка на питание в столовой '{canteen}' сохранена на сегодня!",
            reply_markup=ReplyKeyboardRemove())
        logger.info(f"Заявка сохранена: {user_id} -> {canteen}")
    except Exception as e:
        logger.error(f"Ошибка при сохранении заявки: {e}", exc_info=True)
        await message.answer("❌ Произошла ошибка при сохранении заявки. Попробуйте еще раз.")
    finally:
        conn.close()
        await state.clear()

# Команда для экспорта в Excel (без pandas)
@dp.message(Command("export"))
async def cmd_export(message: types.Message):
    # Проверяем, является ли пользователь администратором
    if message.from_user.id != ADMIN_ID:
        await message.answer("❌ Эта команда доступна только администратору.")
        logger.warning(f"Пользователь {message.from_user.id} попытался использовать /export без прав")
        return
    
    try:
        logger.info("Начало экспорта в Excel...")
        
        # Создаем подключение к базе данных
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        
        # Получаем данные
        query = """
        SELECT 
            u.full_name AS ФИО,
            r.date AS Дата,
            r.canteen AS Столовая
        FROM requests r
        JOIN users u ON u.id = r.user_id
        ORDER BY r.date DESC, u.full_name
        """
        
        logger.info("Выполняем SQL-запрос...")
        cursor.execute(query)
        rows = cursor.fetchall()
        logger.info(f"Получено {len(rows)} записей из БД")
        
        # Если данных нет
        if not rows:
            await message.answer("Нет данных для экспорта.")
            conn.close()
            logger.info("Нет данных для экспорта")
            return
        
        # Создаем Excel-файл
        wb = Workbook()
        ws = wb.active
        ws.title = "Заявки"
        
        # Заголовки
        headers = ["ФИО", "Дата", "Столовая"]
        ws.append(headers)
        
        # Стили для заголовков
        bold_font = Font(bold=True)
        center_alignment = Alignment(horizontal='center')
        
        for col in range(1, 4):
            cell = ws.cell(row=1, column=col)
            cell.font = bold_font
            cell.alignment = center_alignment
        
        # Данные
        for row in rows:
            # Преобразуем дату из строки (если в БД она хранится как строка) в формат дд.мм.гггг
            # Если в БД дата хранится в формате YYYY-MM-DD, то:
            date_str = row[1]
            try:
                # Пытаемся преобразовать строку в дату
                date_obj = datetime.strptime(date_str, "%Y-%m-%d")
                formatted_date = date_obj.strftime("%d.%m.%Y")
            except:
                formatted_date = date_str  # если не удалось, оставляем как есть
            
            ws.append([row[0], formatted_date, row[2]])
        
        # Устанавливаем ширину столбцов
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 15
        
        # Сохраняем файл
        today_str = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        excel_filename = f"meal_requests_{today_str}.xlsx"
        excel_path = os.path.join(EXCEL_FOLDER, excel_filename)
        wb.save(excel_path)
        logger.info(f"Excel-файл сохранен: {excel_path}")
        
        # Проверяем размер файла
        file_size = os.path.getsize(excel_path)
        logger.info(f"Размер файла: {file_size} байт")
        
        # Отправляем файл пользователю
        await message.answer_document(
            types.FSInputFile(excel_path, filename=excel_filename),
            caption=f"Экспорт заявок на питание ({len(rows)} записей)"
        )
        logger.info("Файл успешно отправлен")
        
        # Отправляем подтверждение
        await message.answer("✅ Файл успешно экспортирован и отправлен!")
        
    except Exception as e:
        error_msg = f"❌ Ошибка при создании отчета: {str(e)}"
        logger.error(f"Ошибка при экспорте в Excel: {e}", exc_info=True)
        await message.answer(error_msg)
    finally:
        conn.close()

# Команда для получения статистики (только для администратора)
@dp.message(Command("stats"))
async def cmd_stats(message: types.Message):
    # Проверяем, является ли пользователь администратором
    if message.from_user.id != ADMIN_ID:
        await message.answer("❌ Эта команда доступна только администратору.")
        return
    
    try:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        
        # Количество пользователей
        cursor.execute("SELECT COUNT(*) FROM users")
        users_count = cursor.fetchone()[0]
        
        # Количество заявок
        cursor.execute("SELECT COUNT(*) FROM requests")
        requests_count = cursor.fetchone()[0]
        
        # Последняя заявка
        cursor.execute("SELECT MAX(date) FROM requests")
        last_request_date = cursor.fetchone()[0] or "нет данных"
        
        # Статистика по столовым
        cursor.execute("SELECT canteen, COUNT(*) FROM requests GROUP BY canteen")
        canteen_stats = cursor.fetchall()
        
        stats_message = (
            "📊 Статистика бота:\n"
            f"👤 Пользователей: {users_count}\n"
            f"📝 Заявок: {requests_count}\n"
            f"📅 Последняя заявка: {last_request_date}\n\n"
            "🍽️ Статистика по столовым:\n"
        )
        
        for canteen, count in canteen_stats:
            stats_message += f"- {canteen}: {count} заявок\n"
        
        await message.answer(stats_message)
        
    except Exception as e:
        logger.error(f"Ошибка при получении статистики: {e}")
        await message.answer(f"❌ Ошибка при получении статистики: {str(e)}")
    finally:
        conn.close()

# Функция для отправки напоминаний
async def send_reminders():
    try:
        logger.info("Начало отправки напоминаний...")
        today = date.today()
        
        # Проверяем, будний ли день (пн-пт = 0-4)
        if today.weekday() < 5:  # 0-пн, 4-пт
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            
            # Получаем всех пользователей
            cursor.execute("SELECT telegram_id, full_name FROM users")
            users = cursor.fetchall()
            
            for user_id, full_name in users:
                # Проверяем, есть ли заявка на сегодня
                cursor.execute(
                    "SELECT COUNT(*) FROM requests r JOIN users u ON u.id = r.user_id WHERE u.telegram_id = ? AND r.date = ?",
                    (user_id, today)
                )
                has_request = cursor.fetchone()[0] == 0
                
                if has_request:
                    try:
                        await bot.send_message(
                            user_id,
                            f"⏰ {full_name}, не забудьте подать заявку на питание на сегодня!\n"
                            "Используйте команду /order"
                        )
                        logger.info(f"Напоминание отправлено {user_id}")
                    except Exception as e:
                        logger.error(f"Ошибка отправки напоминания {user_id}: {e}")
            
            conn.close()
        logger.info("Завершена отправка напоминаний")
    except Exception as e:
        logger.error(f"Ошибка в функции напоминаний: {e}", exc_info=True)

# Обработчик неизвестных команд
@dp.message()
async def unknown_command(message: types.Message):
    await message.answer(
        "Я вас не понял. Используйте следующие команды:\n"
        "/start - начать работу\n"
        "/order - подать заявку на питание\n"
        "/change_name - изменить ФИО\n"
        "/export - экспорт данных в Excel (только для администратора)\n"
        "/stats - статистика бота (только для администратора)"
    )

# Запуск бота
async def main():
    logger.info("Бот запущен")
    # Отправляем сообщение администратору при запуске
    try:
        await bot.send_message(
            ADMIN_ID, 
            "✅ Бот успешно запущен!\n"
            f"Папка с отчетами: {EXCEL_FOLDER}\n"
            "Используйте /export для получения данных"
        )
    except Exception as e:
        logger.error(f"Не удалось отправить сообщение администратору: {e}")
    
    # Создаем планировщик для напоминаний
    scheduler = AsyncIOScheduler()
    # Напоминание каждый будний день в 10:00 по Москве
    scheduler.add_job(
        send_reminders,
        trigger=CronTrigger(day_of_week="mon-fri", hour=10, minute=0, timezone="Europe/Moscow"),
    )
    scheduler.start()
    logger.info("Планировщик напоминаний запущен")
    
    await dp.start_polling(bot)

if __name__ == "__main__":
    import asyncio
    asyncio.run(main())