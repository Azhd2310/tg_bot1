import logging
import os
import sqlite3
import re
from datetime import datetime, date, timedelta
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
import calendar

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
    
    # Таблица заявок (добавлены поля submission_date и submission_time)
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS requests (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER,
        meal_date DATE,          -- Дата питания
        submission_date DATE,    -- Дата подачи заявки
        submission_time TIME,    -- Время подачи заявки
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
    waiting_for_meal_date = State()
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
            "Используйте кнопки меню для работы с ботом.",
            reply_markup=create_main_menu(user_id == ADMIN_ID))
        await state.clear()
    else:
        await message.answer(
            "Привет! Я бот для подачи заявок на питание.\n"
            "Пожалуйста, введите ваше ФИО в формате <b>Фамилия И.О.</b>\n"
            "Например: <b>Иванов И.И.</b>",
            parse_mode="HTML"
        )
        await state.set_state(Form.waiting_for_name)

# Создание главного меню
def create_main_menu(is_admin=False):
    builder = ReplyKeyboardBuilder()
    builder.add(KeyboardButton(text="🍽 Подать заявку"))
    builder.add(KeyboardButton(text="✏️ Изменить ФИО"))
    builder.add(KeyboardButton(text="❌ Удалить мои данные"))
    
    if is_admin:
        builder.add(KeyboardButton(text="📊 Статистика"))
        builder.add(KeyboardButton(text="📥 Экспорт в Excel"))
        builder.add(KeyboardButton(text="🧹 Очистить базу"))
    
    builder.adjust(2, 2, 2)
    return builder.as_markup(resize_keyboard=True)

# Обработчик ввода ФИО (с проверкой формата)
@dp.message(Form.waiting_for_name)
async def process_name(message: types.Message, state: FSMContext):
    full_name = message.text.strip()
    user_id = message.from_user.id
    today = date.today()
    
    # Проверка формата ФИО
    if not re.match(r'^[А-ЯЁ][а-яё]+\s+[А-ЯЁ]\.[А-ЯЁ]\.$', full_name):
        await message.answer(
            "❌ Неверный формат ФИО. Пожалуйста, введите в формате: <b>Фамилия И.О.</b>\n"
            "Например: <b>Иванов И.И.</b>\n"
            "Фамилия с заглавной буквы, затем инициалы с точками (И.О.).",
            parse_mode="HTML"
        )
        return
    
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
            f"✅ Спасибо, {full_name}! Ваше ФИО сохранено.\n"
            "Теперь вы можете подать заявку на питание с помощью кнопки меню.",
            reply_markup=create_main_menu(user_id == ADMIN_ID)
        )
        logger.info(f"Пользователь {user_id} сохранен: {full_name}")
    except Exception as e:
        logger.error(f"Ошибка при сохранении ФИО: {e}", exc_info=True)
        await message.answer("❌ Произошла ошибка при сохранении данных. Попробуйте еще раз.")
    finally:
        conn.close()
        await state.clear()

# Обработчик кнопки "Изменить ФИО"
@dp.message(F.text == "✏️ Изменить ФИО")
async def change_name_handler(message: types.Message, state: FSMContext):
    await message.answer("Введите ваше новое ФИО в формате <b>Фамилия И.О.</b>:", 
                         reply_markup=ReplyKeyboardRemove(),
                         parse_mode="HTML")
    await state.set_state(Form.waiting_for_name)
    logger.info(f"Пользователь {message.from_user.id} запросил изменение ФИО")

# Обработчик кнопки "Подать заявку"
@dp.message(F.text == "🍽 Подать заявку")
async def order_handler(message: types.Message, state: FSMContext):
    user_id = message.from_user.id
    
    # Проверяем, зарегистрирован ли пользователь
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT id, full_name FROM users WHERE telegram_id = ?", (user_id,))
    user = cursor.fetchone()
    conn.close()
    
    if not user:
        await message.answer("Пожалуйста, сначала зарегистрируйтесь с помощью команды /start.")
        return
    
    # Создаем клавиатуру с датами (завтра, послезавтра, через 2 дня)
    today = date.today()
    dates = [
        today + timedelta(days=1),
        today + timedelta(days=2),
        today + timedelta(days=3)
    ]
    
    builder = ReplyKeyboardBuilder()
    for d in dates:
        builder.add(KeyboardButton(text=d.strftime("%d.%m.%Y")))
    builder.add(KeyboardButton(text="↩️ Назад"))
    builder.adjust(3, 1)
    
    await message.answer(
        f"{user[1]}, выберите дату питания:",
        reply_markup=builder.as_markup(resize_keyboard=True))
    
    # Сохраняем user_id в контексте
    await state.update_data(user_id=user[0], full_name=user[1])
    await state.set_state(Form.waiting_for_meal_date)
    logger.info(f"Пользователь {user_id} начал подачу заявки")

# Обработчик выбора даты питания
@dp.message(Form.waiting_for_meal_date)
async def process_meal_date(message: types.Message, state: FSMContext):
    # Обработка кнопки "Назад"
    if message.text == "↩️ Назад":
        await message.answer("Главное меню:", reply_markup=create_main_menu(message.from_user.id == ADMIN_ID))
        await state.clear()
        return
    
    date_str = message.text.strip()
    
    # Парсим дату из разных форматов
    try:
        # Пробуем разные форматы даты
        for fmt in ("%d.%m.%Y", "%d.%m.%y", "%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d"):
            try:
                meal_date = datetime.strptime(date_str, fmt).date()
                break
            except ValueError:
                continue
        else:
            raise ValueError("Неверный формат даты")
        
        # Проверяем, что дата не в прошлом
        today = date.today()
        if meal_date < today:
            await message.answer("❌ Нельзя выбрать прошедшую дату. Пожалуйста, выберите другую дату.")
            return
            
    except Exception as e:
        logger.error(f"Ошибка парсинга даты: {e}")
        await message.answer("❌ Неверный формат даты. Пожалуйста, введите дату в формате ДД.ММ.ГГГГ")
        return
    
    # Сохраняем дату в контексте
    await state.update_data(meal_date=meal_date)
    
    # Создаем клавиатуру с выбором столовой
    builder = ReplyKeyboardBuilder()
    builder.add(KeyboardButton(text="Центр"))
    builder.add(KeyboardButton(text="Ястреб"))
    builder.add(KeyboardButton(text="↩️ Назад"))
    builder.adjust(2, 1)
    
    await message.answer(
        f"✅ Вы выбрали дату: {meal_date.strftime('%d.%m.%Y')}\nТеперь выберите столовую:",
        reply_markup=builder.as_markup(resize_keyboard=True))
    
    await state.set_state(Form.waiting_for_canteen)

# Обработчик выбора столовой (с сохранением даты и времени подачи)
@dp.message(Form.waiting_for_canteen)
async def process_canteen(message: types.Message, state: FSMContext):
    # Обработка кнопки "Назад"
    if message.text == "↩️ Назад":
        # Возвращаемся к выбору даты
        user_data = await state.get_data()
        full_name = user_data.get('full_name', 'пользователь')
        
        # Создаем клавиатуру с датами
        today = date.today()
        dates = [
            today + timedelta(days=1),
            today + timedelta(days=2),
            today + timedelta(days=3)
        ]
        
        builder = ReplyKeyboardBuilder()
        for d in dates:
            builder.add(KeyboardButton(text=d.strftime("%d.%m.%Y")))
        builder.add(KeyboardButton(text="↩️ Назад"))
        builder.adjust(3, 1)
        
        await message.answer(
            f"{full_name}, выберите дату питания:",
            reply_markup=builder.as_markup(resize_keyboard=True))
        await state.set_state(Form.waiting_for_meal_date)
        return
    
    # Проверяем, что выбрана столовая
    if message.text not in ["Центр", "Ястреб"]:
        await message.answer("❌ Пожалуйста, выберите столовую из предложенных вариантов.")
        return
    
    canteen = message.text
    user_id = message.from_user.id
    
    # Получаем данные из состояния
    data = await state.get_data()
    meal_date = data.get('meal_date')
    user_id_db = data.get('user_id')
    full_name = data.get('full_name')
    
    if not meal_date or not user_id_db:
        await message.answer("❌ Ошибка: данные сессии утеряны. Начните заново.")
        await state.clear()
        return
    
    # Фиксируем дату и время подачи заявки
    submission_datetime = datetime.now()
    submission_date = submission_datetime.date()
    submission_time = submission_datetime.time()
    
    # Сохраняем заявку
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    
    try:
        # Проверяем, есть ли уже заявка на эту дату
        cursor.execute(
            "SELECT id FROM requests WHERE user_id = ? AND meal_date = ?",
            (user_id_db, meal_date.strftime("%Y-%m-%d"))
        )
        existing_request = cursor.fetchone()
        
        if existing_request:
            # Обновляем существующую заявку
            cursor.execute(
                "UPDATE requests SET canteen = ?, submission_date = ?, submission_time = ? WHERE id = ?",
                (canteen, submission_date.strftime("%Y-%m-%d"), submission_time.strftime("%H:%M:%S"), existing_request[0]))
            action_msg = "обновлена"
        else:
            # Добавляем новую заявку
            cursor.execute(
                "INSERT INTO requests (user_id, meal_date, submission_date, submission_time, canteen) VALUES (?, ?, ?, ?, ?)",
                (user_id_db, meal_date.strftime("%Y-%m-%d"), submission_date.strftime("%Y-%m-%d"), submission_time.strftime("%H:%M:%S"), canteen))
            action_msg = "подана"
        
        conn.commit()
        await message.answer(
            f"✅ Заявка на питание в столовой '{canteen}' {action_msg} на {meal_date.strftime('%d.%m.%Y')}!\n"
            f"📅 Дата подачи: {submission_date.strftime('%d.%m.%Y')}\n"
            f"⏰ Время подачи: {submission_time.strftime('%H:%M:%S')}",
            reply_markup=create_main_menu(user_id == ADMIN_ID))
        logger.info(f"Заявка сохранена: {user_id} -> {canteen} на {meal_date}")
    except Exception as e:
        logger.error(f"Ошибка при сохранении заявки: {e}", exc_info=True)
        await message.answer("❌ Произошла ошибка при сохранении заявки. Попробуйте еще раз.")
    finally:
        conn.close()
        await state.clear()

# Обработчик кнопки "Экспорт в Excel" (с объединением даты и времени подачи)
@dp.message(F.text == "📥 Экспорт в Excel")
async def export_handler(message: types.Message):
    # Проверяем, является ли пользователь администратором
    if message.from_user.id != ADMIN_ID:
        await message.answer("❌ Эта команда доступна только администратору.")
        logger.warning(f"Пользователь {message.from_user.id} попытался использовать экспорт без прав")
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
            r.meal_date AS Дата_питания,
            r.submission_date AS Дата_подачи,
            r.submission_time AS Время_подачи,
            r.canteen AS Столовая
        FROM requests r
        JOIN users u ON u.id = r.user_id
        ORDER BY r.meal_date DESC, u.full_name
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
        headers = ["ФИО", "Дата питания", "Дата и время подачи", "Столовая"]
        ws.append(headers)
        
        # Стили для заголовков
        bold_font = Font(bold=True)
        center_alignment = Alignment(horizontal='center')
        
        for col in range(1, 5):
            cell = ws.cell(row=1, column=col)
            cell.font = bold_font
            cell.alignment = center_alignment
        
        # Данные
        for row in rows:
            # Форматируем даты
            meal_date = datetime.strptime(row[1], "%Y-%m-%d").strftime("%d.%m.%Y")
            submission_date = datetime.strptime(row[2], "%Y-%m-%d").strftime("%d.%m.%Y")
            submission_time = row[3]  # Время уже в формате строки
            
            # Объединяем дату и время подачи
            full_submission = f"{submission_date} {submission_time}"
            
            ws.append([row[0], meal_date, full_submission, row[4]])
        
        # Устанавливаем ширину столбцов
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 20  # Увеличим для полной даты-времени
        ws.column_dimensions['D'].width = 15
        
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

# Обработчик кнопки "Очистить базу"
@dp.message(F.text == "🧹 Очистить базу")
async def clear_db_handler(message: types.Message):
    # Проверяем, является ли пользователь администратором
    if message.from_user.id != ADMIN_ID:
        await message.answer("❌ Эта команда доступна только администратору.")
        return
    
    try:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        
        # Удаляем все данные
        cursor.execute("DELETE FROM requests")
        cursor.execute("DELETE FROM users")
        conn.commit()
        
        await message.answer("✅ База данных полностью очищена!")
        logger.info("База данных очищена администратором")
        
    except Exception as e:
        logger.error(f"Ошибка при очистке БД: {e}")
        await message.answer(f"❌ Ошибка при очистке базы данных: {str(e)}")
    finally:
        conn.close()

# Обработчик кнопки "Удалить мои данные"
@dp.message(F.text == "❌ Удалить мои данные")
async def delete_me_handler(message: types.Message):
    user_id = message.from_user.id
    
    try:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        
        # Удаляем заявки пользователя
        cursor.execute("DELETE FROM requests WHERE user_id IN (SELECT id FROM users WHERE telegram_id = ?)", (user_id,))
        
        # Удаляем самого пользователя
        cursor.execute("DELETE FROM users WHERE telegram_id = ?", (user_id,))
        conn.commit()
        
        if cursor.rowcount > 0:
            await message.answer("✅ Ваши данные полностью удалены из системы!", reply_markup=ReplyKeyboardRemove())
            logger.info(f"Пользователь {user_id} удалил свои данные")
        else:
            await message.answer("ℹ️ Ваши данные не найдены в системе.")
        
    except Exception as e:
        logger.error(f"Ошибка при удалении пользователя {user_id}: {e}")
        await message.answer(f"❌ Произошла ошибка при удалении данных: {str(e)}")
    finally:
        conn.close()

# Обработчик кнопки "Статистика"
@dp.message(F.text == "📊 Статистика")
async def stats_handler(message: types.Message):
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
        cursor.execute("SELECT MAX(meal_date) FROM requests")
        last_meal_date = cursor.fetchone()[0] or "нет данных"
        
        # Статистика по столовым
        cursor.execute("SELECT canteen, COUNT(*) FROM requests GROUP BY canteen")
        canteen_stats = cursor.fetchall()
        
        # Статистика по датам
        cursor.execute("SELECT meal_date, COUNT(*) FROM requests GROUP BY meal_date ORDER BY meal_date DESC LIMIT 7")
        date_stats = cursor.fetchall()
        
        stats_message = (
            "📊 Статистика бота:\n"
            f"👤 Пользователей: {users_count}\n"
            f"📝 Заявок: {requests_count}\n"
            f"📅 Последняя дата питания: {last_meal_date}\n\n"
            "🍽️ Статистика по столовым:\n"
        )
        
        for canteen, count in canteen_stats:
            stats_message += f"- {canteen}: {count} заявок\n"
        
        stats_message += "\n📅 Заявки за последние 7 дней:\n"
        for meal_date, count in date_stats:
            stats_message += f"- {meal_date}: {count} заявок\n"
        
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
                # Проверяем, есть ли заявка на завтра
                tomorrow = today + timedelta(days=1)
                cursor.execute(
                    "SELECT COUNT(*) FROM requests r JOIN users u ON u.id = r.user_id WHERE u.telegram_id = ? AND r.meal_date = ?",
                    (user_id, tomorrow.strftime("%Y-%m-%d"))
                )
                has_request = cursor.fetchone()[0] == 0
                
                if has_request:
                    try:
                        await bot.send_message(
                            user_id,
                            f"⏰ {full_name}, не забудьте подать заявку на питание на завтра ({tomorrow.strftime('%d.%m.%Y')})!\n"
                            "Используйте кнопку '🍽 Подать заявку'"
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
        "Я вас не понял. Используйте кнопки меню для работы с ботом.",
        reply_markup=create_main_menu(message.from_user.id == ADMIN_ID)
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
            "Используйте меню для управления ботом.\n\n"
            "Основные улучшения:\n"
            "- Проверка формата ФИО (Фамилия И.О.)\n"
            "- Раздельное сохранение даты и времени подачи заявки\n"
            "- Улучшенный экспорт в Excel\n"
            "- Подробные уведомления о подаче заявок"
        )
    except Exception as e:
        logger.error(f"Не удалось отправить сообщение администратору: {e}")
    
    # Создаем планировщик для напоминаний
    scheduler = AsyncIOScheduler()
    # Напоминание каждый будний день в 16:00 по Москве
    scheduler.add_job(
        send_reminders,
        trigger=CronTrigger(day_of_week="mon-fri", hour=16, minute=0, timezone="Europe/Moscow"),
    )
    scheduler.start()
    logger.info("Планировщик напоминаний запущен (будни в 16:00)")
    
    await dp.start_polling(bot)

if __name__ == "__main__":
    import asyncio
    asyncio.run(main())