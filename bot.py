import asyncio
import logging
import sqlite3
import io
import calendar
import re
import os
from datetime import datetime, time, timedelta
from math import radians, sin, cos, sqrt, asin

import pandas as pd
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    ReplyKeyboardMarkup, KeyboardButton,
    InlineKeyboardMarkup, InlineKeyboardButton,
    BufferedInputFile
)
from aiohttp import web

# ================= КОНФИГУРАЦИЯ =================
BOT_TOKEN = "8836765870:AAHA5NiXfxxnADr2sHGI-w6E6HB5gob4nGQ"
ADMIN_ID = 769121021
MAX_ADMINS = 3

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(storage=MemoryStorage())
logging.basicConfig(level=logging.INFO)

DAYS_MAP = {1: "Пн", 2: "Вт", 3: "Ср", 4: "Чт", 5: "Пт", 6: "Сб", 7: "Вс"}

# Порог переработки: меньше этого значения (в минутах) не считается переработкой
OVERTIME_THRESHOLD_MIN = 60
# Допуск опоздания: 0 = любое опоздание считается
LATE_TOLERANCE_MIN = 0

# ================= НАСТРОЙКИ ЗАЩИТЫ ОТ ФЕЙКОВОГО GPS =================
# Максимально допустимая точность (меньше = подозрительно, часто у фейков 0)
MIN_GPS_ACCURACY = 5  # метров
# Максимально допустимая точность (больше = подозрительно)
MAX_GPS_ACCURACY = 500  # метров
# Максимальная скорость перемещения (м/с). 200 км/ч ≈ 55 м/с
MAX_GPS_SPEED = 55  # м/с
# Максимальное расстояние от офиса для начала смены (100 км)
MAX_DISTANCE_FROM_OFFICE = 100000  # метров (100 км)
# Максимальное расстояние между двумя последовательными отметками (за 1 час)
MAX_HOURLY_MOVEMENT = 500000  # метров (500 км)


# ================= БАЗА ДАННЫХ =================

def init_db():
    conn = sqlite3.connect('clockster.db')
    cursor = conn.cursor()

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            user_id INTEGER PRIMARY KEY,
            username TEXT,
            full_name TEXT,
            role TEXT DEFAULT 'employee',
            department TEXT DEFAULT 'Общий',
            job_title TEXT DEFAULT 'Сотрудник',
            phone_number TEXT,
            target_lat REAL,
            target_lon REAL,
            radius INTEGER DEFAULT 100,
            is_working INTEGER DEFAULT 0,
            shift_start_time TEXT,
            schedule_start TEXT DEFAULT '09:00',
            schedule_end TEXT DEFAULT '18:00',
            monthly_salary REAL DEFAULT 0,
            work_days_week TEXT DEFAULT '1,2,3,4,5',
            last_start_reminder TEXT,
            last_end_reminder TEXT,
            last_location_lat REAL,
            last_location_lon REAL,
            last_location_time TEXT
        )
    ''')

    for col in [
        "work_days_week TEXT DEFAULT '1,2,3,4,5'",
        "last_start_reminder TEXT",
        "last_end_reminder TEXT",
        "last_location_lat REAL",
        "last_location_lon REAL",
        "last_location_time TEXT"
    ]:
        try:
            cursor.execute(f"ALTER TABLE users ADD COLUMN {col}")
        except sqlite3.OperationalError:
            pass

    cursor.execute(
        "INSERT OR IGNORE INTO users (user_id, username, full_name, role, department, job_title, "
        "schedule_start, schedule_end, work_days_week) "
        "VALUES (?, ?, ?, 'admin', 'Администрация', 'Системный администратор', '09:00', '18:00', '1,2,3,4,5')",
        (ADMIN_ID, f"admin_{ADMIN_ID}", "Admin",)
    )
    cursor.execute(
        "UPDATE users SET department='Администрация', job_title='Системный администратор', "
        "schedule_start='09:00', schedule_end='18:00', work_days_week='1,2,3,4,5' WHERE user_id = ?",
        (ADMIN_ID,)
    )

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS shifts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            start_time TEXT,
            end_time TEXT,
            duration_min INTEGER,
            overtime_min INTEGER DEFAULT 0,
            is_late INTEGER DEFAULT 0,
            late_minutes INTEGER DEFAULT 0
        )
    ''')

    try:
        cursor.execute("ALTER TABLE shifts ADD COLUMN late_minutes INTEGER DEFAULT 0")
    except sqlite3.OperationalError:
        pass

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS gps_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            timestamp TEXT,
            latitude REAL,
            longitude REAL,
            accuracy REAL,
            speed REAL,
            distance_from_office REAL,
            reason TEXT,
            action_taken TEXT
        )
    ''')

    conn.commit()
    conn.close()


def log_suspicious_gps(user_id, latitude, longitude, accuracy, speed, distance, reason, action_taken="blocked"):
    try:
        conn = sqlite3.connect('clockster.db')
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO gps_log (user_id, timestamp, latitude, longitude, accuracy, speed, distance_from_office, reason, action_taken) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (user_id, datetime.now().isoformat(), latitude, longitude, accuracy, speed, distance, reason, action_taken)
        )
        conn.commit()
        conn.close()
        logging.warning(f"🚨 Подозрительный GPS: user={user_id}, причина={reason}, действие={action_taken}")
    except Exception as e:
        logging.error(f"Ошибка логирования GPS: {e}")


def migrate_late_minutes():
    try:
        conn = sqlite3.connect('clockster.db')
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        cursor.execute("SELECT * FROM shifts")
        shifts = [dict(r) for r in cursor.fetchall()]

        updated = 0
        for s in shifts:
            try:
                start_dt = datetime.fromisoformat(s['start_time'])
                cursor.execute('SELECT schedule_start FROM users WHERE user_id = ?', (s['user_id'],))
                u = cursor.fetchone()
                if not u:
                    continue
                schedule_start_str = u['schedule_start'] or '09:00'
                s_start_time = datetime.strptime(schedule_start_str, '%H:%M').time()
                scheduled_start = datetime.combine(start_dt.date(), s_start_time)
                diff_min = int((start_dt - scheduled_start).total_seconds() / 60)

                is_late = 1 if diff_min > LATE_TOLERANCE_MIN else 0
                late_minutes = max(0, diff_min)

                cursor.execute(
                    'UPDATE shifts SET is_late = ?, late_minutes = ? WHERE id = ?',
                    (is_late, late_minutes, s['id'])
                )
                updated += 1
            except Exception:
                continue

        conn.commit()
        conn.close()

        if updated > 0:
            logging.info(f"🔄 Миграция: обновлено {updated} записей")
        return updated
    except Exception as e:
        logging.error(f"Ошибка миграции: {e}")
        return 0


def is_working_day(date_obj, work_days_str):
    try:
        work_days = [int(x) for x in work_days_str.split(',')]
    except ValueError:
        work_days = [1, 2, 3, 4, 5]
    iso_weekday = date_obj.weekday() + 1
    return iso_weekday in work_days


def get_working_days_in_month(year, month, work_days_str):
    try:
        work_days = [int(x) for x in work_days_str.split(',')]
    except ValueError:
        work_days = [1, 2, 3, 4, 5]
    count = 0
    _, num_days = calendar.monthrange(year, month)
    for day in range(1, num_days + 1):
        iso_weekday = datetime(year, month, day).weekday() + 1
        if iso_weekday in work_days:
            count += 1
    return count


def format_work_days(work_days_str):
    try:
        days = [int(x) for x in work_days_str.split(',')]
        return ", ".join([DAYS_MAP[d] for d in sorted(days) if d in DAYS_MAP])
    except ValueError:
        return "Пн-Пт"


def format_duration(minutes):
    if minutes is None or minutes <= 0:
        return "0 мин"
    try:
        minutes = int(minutes)
    except (ValueError, TypeError):
        return "0 мин"
    hours = minutes // 60
    mins = minutes % 60
    if hours == 0:
        return f"{mins} мин"
    elif mins == 0:
        return f"{hours}ч"
    else:
        return f"{hours}ч {mins}мин"


def get_user_by_username(username):
    if not username: return None
    username = username.lstrip('@').strip()
    conn = sqlite3.connect('clockster.db')
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM users WHERE LOWER(username) = LOWER(?)", (username,))
    row = cursor.fetchone()
    conn.close()
    return dict(row) if row else None


def get_user_by_phone(phone):
    if not phone: return None
    clean_phone = ''.join(filter(str.isdigit, phone))
    conn = sqlite3.connect('clockster.db')
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM users WHERE phone_number = ? OR username = ?", (clean_phone, clean_phone))
    row = cursor.fetchone()
    conn.close()
    return dict(row) if row else None


def get_users_by_filters(department=None, job_title=None, role=None, user_id=None):
    conn = sqlite3.connect('clockster.db')
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    query = "SELECT * FROM users WHERE 1=1"
    params = []
    if role:
        query += " AND role = ?";
        params.append(role)
    if department and department != "Все":
        query += " AND department = ?";
        params.append(department)
    if job_title and job_title != "Все":
        query += " AND job_title = ?";
        params.append(job_title)
    if user_id:
        query += " AND user_id = ?";
        params.append(user_id)
    cursor.execute(query, params)
    users = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return users


def get_unique_values(column):
    conn = sqlite3.connect('clockster.db')
    cursor = conn.cursor()
    cursor.execute(f"SELECT DISTINCT {column} FROM users WHERE role IN ('employee', 'admin')")
    values = [v[0] for v in cursor.fetchall() if v[0]]
    conn.close()
    return values


def get_user_by_id(user_id):
    conn = sqlite3.connect('clockster.db')
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM users WHERE user_id = ?", (user_id,))
    row = cursor.fetchone()
    conn.close()
    return dict(row) if row else None


def get_all_employees():
    conn = sqlite3.connect('clockster.db')
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM users WHERE role IN ('employee', 'admin')")
    users = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return users


def get_all_admins():
    conn = sqlite3.connect('clockster.db')
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM users WHERE role = 'admin' AND user_id != ?", (ADMIN_ID,))
    users = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return users


def get_admin_count():
    conn = sqlite3.connect('clockster.db')
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM users WHERE role = 'admin'")
    return cursor.fetchone()[0]


def is_admin(user_id):
    if user_id == ADMIN_ID: return True
    user = get_user_by_id(user_id)
    return bool(user and user.get('role') == 'admin')


def delete_user_by_id(user_id):
    if user_id == ADMIN_ID: return False
    conn = sqlite3.connect('clockster.db')
    cursor = conn.cursor()
    cursor.execute("DELETE FROM users WHERE user_id = ?", (user_id,))
    deleted = cursor.rowcount > 0
    if deleted:
        cursor.execute("DELETE FROM shifts WHERE user_id = ?", (user_id,))
    conn.commit()
    conn.close()
    return deleted


def update_user_data(record_id, **kwargs):
    kwargs = {k: v for k, v in kwargs.items() if v is not None}
    if not kwargs: return
    conn = sqlite3.connect('clockster.db')
    cursor = conn.cursor()
    set_clause = ", ".join([f"{k} = ?" for k in kwargs.keys()])
    values = list(kwargs.values()) + [record_id]
    cursor.execute(f"UPDATE users SET {set_clause} WHERE user_id = ?", values)
    conn.commit()
    conn.close()


def calculate_late(start_dt, schedule_start_str):
    try:
        s_start_time = datetime.strptime(schedule_start_str, "%H:%M").time()
        scheduled_start = datetime.combine(start_dt.date(), s_start_time)
        diff_min = int((start_dt - scheduled_start).total_seconds() / 60)
        is_late = 1 if diff_min > LATE_TOLERANCE_MIN else 0
        return is_late, max(0, diff_min)
    except ValueError:
        return 0, 0


def calculate_overtime(end_dt, schedule_end_str):
    try:
        s_end_time = datetime.strptime(schedule_end_str, "%H:%M").time()
        s_end_dt = datetime.combine(end_dt.date(), s_end_time)
        if end_dt > s_end_dt:
            overtime_min = int((end_dt - s_end_dt).total_seconds() / 60)
            if overtime_min <= OVERTIME_THRESHOLD_MIN:
                return 0
            return overtime_min
        return 0
    except ValueError:
        return 0


def create_shift_record(user_id, start_str, end_str, duration, overtime, is_late, late_minutes=0):
    conn = sqlite3.connect('clockster.db')
    cursor = conn.cursor()

    cursor.execute(
        "SELECT id FROM shifts WHERE user_id = ? AND start_time = ? AND end_time = ?",
        (user_id, start_str, end_str)
    )
    if cursor.fetchone():
        conn.close()
        logging.info(f"️ Дубликат смены не создан: {user_id} {start_str} - {end_str}")
        return

    cursor.execute(
        "INSERT INTO shifts (user_id, start_time, end_time, duration_min, overtime_min, is_late, late_minutes) "
        "VALUES (?, ?, ?, ?, ?, ?, ?)",
        (user_id, start_str, end_str, duration, overtime, is_late, late_minutes)
    )
    conn.commit()
    conn.close()


def update_shift_time(shift_id, new_start, new_end):
    conn = sqlite3.connect('clockster.db')
    cursor = conn.cursor()
    cursor.execute("SELECT user_id FROM shifts WHERE id = ?", (shift_id,))
    res = cursor.fetchone()
    if not res:
        conn.close()
        return False
    user_id = res[0]
    user = get_user_by_id(user_id)
    start_dt = datetime.fromisoformat(new_start)
    end_dt = datetime.fromisoformat(new_end)
    duration_min = int((end_dt - start_dt).total_seconds() / 60)

    schedule_start_str = user.get('schedule_start') or "09:00"
    schedule_end_str = user.get('schedule_end') or "18:00"

    is_late, late_minutes = calculate_late(start_dt, schedule_start_str)
    overtime_min = calculate_overtime(end_dt, schedule_end_str)

    cursor.execute(
        "UPDATE shifts SET start_time=?, end_time=?, duration_min=?, overtime_min=?, is_late=?, late_minutes=? WHERE id=?",
        (new_start, new_end, duration_min, overtime_min, is_late, late_minutes, shift_id)
    )
    conn.commit()
    conn.close()
    return True


def get_user_shifts(user_id, month_filter=None):
    conn = sqlite3.connect('clockster.db')
    cursor = conn.cursor()
    if month_filter:
        cursor.execute(
            "SELECT * FROM shifts WHERE user_id = ? AND strftime('%Y-%m', start_time) = ? ORDER BY start_time DESC",
            (user_id, month_filter))
    else:
        cursor.execute("SELECT * FROM shifts WHERE user_id = ? ORDER BY start_time DESC", (user_id,))
    shifts = cursor.fetchall()
    conn.close()
    return shifts


def get_unique_shifts(shifts):
    seen = set()
    unique = []
    for s in shifts:
        start_time = s[2]
        if start_time not in seen:
            seen.add(start_time)
            unique.append(s)
    return unique


def get_available_months():
    conn = sqlite3.connect('clockster.db')
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT strftime('%Y-%m', start_time) as month FROM shifts ORDER BY month DESC LIMIT 12")
    months = [row[0] for row in cursor.fetchall()]
    conn.close()
    return months


def count_shifts_by_day_type(shifts, work_days_str):
    work_days_count = 0
    weekend_count = 0
    for s in shifts:
        try:
            start_dt = datetime.fromisoformat(s[2])
            if is_working_day(start_dt, work_days_str):
                work_days_count += 1
            else:
                weekend_count += 1
        except (ValueError, IndexError):
            continue
    return work_days_count, weekend_count


def remove_duplicate_shifts():
    conn = sqlite3.connect('clockster.db')
    cursor = conn.cursor()
    cursor.execute('''
        DELETE FROM shifts 
        WHERE id NOT IN (
            SELECT MIN(id) 
            FROM shifts 
            GROUP BY user_id, start_time, end_time
        )
    ''')
    deleted_count = cursor.rowcount
    conn.commit()
    conn.close()
    if deleted_count > 0:
        logging.info(f"️ Удалено {deleted_count} дубликатов смен")
    return deleted_count


def safe_shift_value(s, idx, default=0):
    try:
        if len(s) > idx and s[idx] is not None:
            return s[idx]
        return default
    except (IndexError, TypeError):
        return default


# ================= ЗАЩИТА ОТ ФЕЙКОВОГО GPS =================

def validate_gps_location(location, user):
    lat = location.latitude
    lon = location.longitude
    accuracy = getattr(location, 'horizontal_accuracy', None)
    speed = getattr(location, 'speed', None)

    if lat == 0 and lon == 0:
        return False, "Координаты (0, 0) — невозможно", 0

    if not (-90 <= lat <= 90) or not (-180 <= lon <= 180):
        return False, "Невозможные координаты", 0

    office_lat = user.get('target_lat')
    office_lon = user.get('target_lon')

    if office_lat is None or office_lon is None:
        return False, "Геозона офиса не настроена администратором", 0

    distance = haversine(lon, lat, office_lon, office_lat)

    if distance > MAX_DISTANCE_FROM_OFFICE:
        return False, f"Слишком далеко от офиса ({int(distance / 1000)} км). Возможно фейковый GPS", distance

    if accuracy is not None:
        if accuracy < MIN_GPS_ACCURACY:
            return False, f"Подозрительная точность GPS: {accuracy}м (возможно фейк)", distance
        if accuracy > MAX_GPS_ACCURACY:
            return False, f"Слишком низкая точность GPS: {accuracy}м", distance

    if speed is not None and speed > MAX_GPS_SPEED:
        return False, f"Подозрительная скорость: {speed} м/с ({speed * 3.6:.0f} км/ч)", distance

    last_loc_time = user.get('last_location_time')
    last_loc_lat = user.get('last_location_lat')
    last_loc_lon = user.get('last_location_lon')

    if last_loc_time and last_loc_lat is not None and last_loc_lon is not None:
        try:
            last_time = datetime.fromisoformat(last_loc_time)
            time_diff = (datetime.now() - last_time).total_seconds()

            if time_diff < 3600 and time_diff > 0:
                movement_distance = haversine(lon, lat, last_loc_lon, last_loc_lat)

                if movement_distance > MAX_HOURLY_MOVEMENT:
                    return False, f"Невозможное перемещение: {int(movement_distance / 1000)} км за {int(time_diff / 60)} мин", distance
        except Exception:
            pass

    return True, "OK", distance


# ================= РАСЧЕТ ЗАРПЛАТЫ =================

def calculate_salary(user, shifts, period_str):
    monthly_salary = user.get('monthly_salary') or 0
    work_days_str = user.get('work_days_week') or '1,2,3,4,5'
    try:
        year, month = map(int, period_str.split('-'))
        work_days_norm = get_working_days_in_month(year, month, work_days_str)
    except Exception:
        work_days_norm = 22
    daily_rate = monthly_salary / work_days_norm if work_days_norm > 0 else 0
    actual_days = len(shifts) if shifts else 0
    base_pay = round(daily_rate * actual_days, 2)
    final_pay = round(base_pay, 2)
    return final_pay, base_pay, daily_rate, work_days_norm


# ================= ВСПОМОГАТЕЛЬНЫЕ =================

def haversine(lon1, lat1, lon2, lat2):
    lon1, lat1, lon2, lat2 = map(radians, [lon1, lat1, lon2, lat2])
    dlon = lon2 - lon1
    dlat = lat2 - lat1
    a = sin(dlat / 2) ** 2 + cos(lat1) * cos(lat2) * sin(dlon / 2) ** 2
    c = 2 * asin(sqrt(a))
    return c * 6371000


def auto_finish_old_shift(user):
    if not user.get('is_working'):
        return False
    shift_start_iso = user.get('shift_start_time')
    if not shift_start_iso:
        return False
    try:
        start_dt = datetime.fromisoformat(shift_start_iso)
        now = datetime.now()
        if start_dt.date() >= now.date():
            return False
        schedule_start_str = user.get('schedule_start') or "09:00"
        schedule_end_str = user.get('schedule_end') or "18:00"
        try:
            s_end_time = datetime.strptime(schedule_end_str, "%H:%M").time()
            end_dt = datetime.combine(start_dt.date(), s_end_time)
        except ValueError:
            end_dt = datetime.combine(start_dt.date() + timedelta(days=1), time(0, 0))
        duration_min = int((end_dt - start_dt).total_seconds() / 60)

        is_late, late_minutes = calculate_late(start_dt, schedule_start_str)
        overtime_min = 0

        create_shift_record(user['user_id'], shift_start_iso, end_dt.isoformat(), duration_min, overtime_min, is_late,
                            late_minutes)
        update_user_data(user['user_id'], is_working=0, shift_start_time=None)
        logging.info(
            f"🔄 Автозавершение {user['user_id']}: смена от {start_dt.strftime('%d.%m %H:%M')} → {end_dt.strftime('%H:%M')} (по графику)")
        return True
    except Exception as e:
        logging.error(f"Ошибка автозавершения {user['user_id']}: {e}")
        return False


def finish_shift_for_user(user_id):
    user = get_user_by_id(user_id)
    if not user or not user.get('is_working'):
        return None, "Пользователь не на смене."
    start_iso = user.get('shift_start_time')
    if not start_iso:
        return None, "Нет данных о начале смены."
    try:
        end_dt = datetime.now()
        start_dt = datetime.fromisoformat(start_iso)
    except Exception as e:
        logging.error(f"Ошибка парсинга: {e}")
        return None, "Ошибка данных смены."
    duration_min = int((end_dt - start_dt).total_seconds() / 60)

    schedule_start_str = user.get('schedule_start') or "09:00"
    schedule_end_str = user.get('schedule_end') or "18:00"

    is_late, late_minutes = calculate_late(start_dt, schedule_start_str)
    overtime_min = calculate_overtime(end_dt, schedule_end_str)

    create_shift_record(user_id, start_iso, end_dt.isoformat(), duration_min, overtime_min, is_late, late_minutes)
    update_user_data(user_id, is_working=0, shift_start_time=None)

    work_days_str = user.get('work_days_week') or '1,2,3,4,5'
    was_weekend = not is_working_day(start_dt, work_days_str)

    result_msg = f"🔴 Смена завершена. Длительность: {duration_min // 60}ч {duration_min % 60}м."
    if was_weekend:
        result_msg += f"\n🌟 Смена в выходной ({DAYS_MAP[start_dt.weekday() + 1]})"
    if overtime_min > 0:
        result_msg += f" Переработка: {format_duration(overtime_min)}."
    if is_late:
        result_msg += f"\n⚠️ Опоздание: {format_duration(late_minutes)}."
    return user, result_msg


def finish_shift_at_time(user_id, end_dt):
    user = get_user_by_id(user_id)
    if not user or not user.get('is_working'):
        return None, "Пользователь не на смене."
    start_iso = user.get('shift_start_time')
    if not start_iso:
        return None, "Нет данных о начале смены."
    try:
        start_dt = datetime.fromisoformat(start_iso)
    except Exception as e:
        logging.error(f"Ошибка парсинга: {e}")
        return None, "Ошибка данных смены."
    if end_dt <= start_dt:
        return None, "❌ Время конца должно быть позже времени начала."
    duration_min = int((end_dt - start_dt).total_seconds() / 60)

    schedule_start_str = user.get('schedule_start') or "09:00"
    schedule_end_str = user.get('schedule_end') or "18:00"

    is_late, late_minutes = calculate_late(start_dt, schedule_start_str)
    overtime_min = calculate_overtime(end_dt, schedule_end_str)

    create_shift_record(user_id, start_iso, end_dt.isoformat(), duration_min, overtime_min, is_late, late_minutes)
    update_user_data(user_id, is_working=0, shift_start_time=None)

    work_days_str = user.get('work_days_week') or '1,2,3,4,5'
    was_weekend = not is_working_day(start_dt, work_days_str)

    result_msg = f"🔴 Смена завершена. Длительность: {duration_min // 60}ч {duration_min % 60}м."
    if was_weekend:
        result_msg += f"\n Смена в выходной ({DAYS_MAP[start_dt.weekday() + 1]})"
    if overtime_min > 0:
        result_msg += f" Переработка: {format_duration(overtime_min)}."
    if is_late:
        result_msg += f"\n⚠️ Опоздание: {format_duration(late_minutes)}."
    return user, result_msg


def parse_callback_value(data: str, prefix: str) -> str:
    if not data.startswith(prefix + "_") and not data.startswith(prefix + "::"):
        return ""
    separator = "::" if "::" in data else "_"
    return data[len(prefix) + len(separator):]


def format_money(amount):
    return f"{amount:,.2f} ₸".replace(",", " ")


def format_month_display(month_str):
    try:
        dt = datetime.strptime(month_str, '%Y-%m')
        months_ru = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
                     'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь']
        return f"{months_ru[dt.month - 1]} {dt.year}"
    except:
        return month_str


# ================= ФОНОВЫЕ НАПОМИНАНИЯ =================

def is_time_in_window(current_time, target_time, window_minutes=2):
    current_minutes = current_time.hour * 60 + current_time.minute
    target_minutes = target_time.hour * 60 + target_time.minute
    return abs(current_minutes - target_minutes) <= window_minutes


async def reminder_loop():
    await asyncio.sleep(10)
    logging.info("🔔 Система напоминаний запущена")
    while True:
        try:
            now = datetime.now()
            current_time = now.time()
            today_str = now.strftime('%Y-%m-%d')
            conn = sqlite3.connect('clockster.db')
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM users WHERE role IN ('employee', 'admin') AND user_id > 0")
            users = [dict(r) for r in cursor.fetchall()]
            conn.close()
            for user in users:
                user_id = user['user_id']
                schedule_start_str = user.get('schedule_start') or "09:00"
                schedule_end_str = user.get('schedule_end') or "18:00"
                work_days_str = user.get('work_days_week') or '1,2,3,4,5'
                try:
                    start_time = datetime.strptime(schedule_start_str, "%H:%M").time()
                    end_time = datetime.strptime(schedule_end_str, "%H:%M").time()
                except ValueError:
                    continue
                was_auto_finished = auto_finish_old_shift(user)
                is_currently_working = user.get('is_working') and not was_auto_finished
                reminder_start = (datetime.combine(now.date(), start_time) - timedelta(minutes=10)).time()
                if is_time_in_window(current_time, reminder_start, 1):
                    last_rem = user.get('last_start_reminder')
                    if last_rem != today_str:
                        if is_working_day(now, work_days_str) and not is_currently_working:
                            try:
                                await bot.send_message(user_id,
                                                       f"⏰ <b>Напоминание</b>\nЧерез 10 минут начало смены ({schedule_start_str}).\nНе забудьте отметиться!",
                                                       parse_mode="HTML")
                                update_user_data(user_id, last_start_reminder=today_str)
                            except Exception as e:
                                logging.warning(f"Не удалось отправить {user_id}: {e}")
                reminder_end = (datetime.combine(now.date(), end_time) - timedelta(minutes=10)).time()
                if is_time_in_window(current_time, reminder_end, 1):
                    last_rem = user.get('last_end_reminder')
                    if last_rem != today_str:
                        if is_currently_working:
                            try:
                                await bot.send_message(user_id,
                                                       f"⏰ <b>Напоминание</b>\nЧерез 10 минут конец смены ({schedule_end_str}).\nНе забудьте завершить смену!",
                                                       parse_mode="HTML")
                                update_user_data(user_id, last_end_reminder=today_str)
                            except Exception as e:
                                logging.warning(f"Не удалось отправить {user_id}: {e}")
        except Exception as e:
            logging.error(f"Ошибка в reminder_loop: {e}")
        await asyncio.sleep(60)


# ================= FSM =================

class SettingsState(StatesGroup):
    choosing_action = State()
    choosing_target = State()
    choosing_dept = State()
    choosing_user = State()
    choosing_target_loc = State()
    input_start_time = State()
    input_end_time = State()
    input_radius = State()
    input_salary = State()
    selecting_work_days = State()


class AddEmployeeState(StatesGroup):
    input_identifier = State()
    input_dept = State()
    input_title = State()
    selecting_work_days = State()


class DeleteEmployeeState(StatesGroup):
    selecting_user = State()


class DeleteAdminState(StatesGroup):
    selecting_user = State()


class AddAdminState(StatesGroup):
    input_identifier = State()
    confirm = State()


class ReportState(StatesGroup):
    choosing_type = State()
    filter_dept = State()
    choosing_period = State()
    showing_result = State()


class EditShiftState(StatesGroup):
    choosing_dept = State()
    selecting_user = State()
    selecting_shift = State()
    input_start = State()
    input_end = State()
    current_shift_action = State()
    editing_current_start = State()
    editing_current_end = State()


class EditNameState(StatesGroup):
    choosing_dept = State()
    selecting_user = State()
    input_name = State()


class DashboardState(StatesGroup):
    choosing_filter = State()


class BroadcastState(StatesGroup):
    waiting_text = State()


# ================= КЛАВИАТУРЫ =================

def get_employee_keyboard():
    return ReplyKeyboardMarkup(resize_keyboard=True, keyboard=[
        [KeyboardButton(text="📍 Начать смену", request_location=True)],
        [KeyboardButton(text=" Завершить смену")],
        [KeyboardButton(text="📊 Моя статистика")]
    ])


def get_admin_main_keyboard():
    return ReplyKeyboardMarkup(resize_keyboard=True, keyboard=[
        [KeyboardButton(text="📍 Начать смену", request_location=True), KeyboardButton(text=" Завершить смену")],
        [KeyboardButton(text="📊 Моя статистика")],
        [KeyboardButton(text="⚙️ Настройки"), KeyboardButton(text="📋 Список сотрудников")],
        [KeyboardButton(text=" Отчеты и Excel"), KeyboardButton(text="📢 Рассылка")],
        [KeyboardButton(text="➕ Добавить сотрудника"), KeyboardButton(text="🗑 Удалить сотрудника")],
        [KeyboardButton(text="✏️ Исправить смену"), KeyboardButton(text="📝 Изменить имя")],
        [KeyboardButton(text="➕ Назначить Админа"), KeyboardButton(text="🗑 Удалить Админа")],
        [KeyboardButton(text="🗑️ Удалить дубликаты"), KeyboardButton(text="🔄 Сброс состояния")],
        [KeyboardButton(text="🛡️ Журнал GPS")]
    ])


def get_cancel_keyboard():
    return InlineKeyboardMarkup(
        inline_keyboard=[[InlineKeyboardButton(text="🔙 Отмена", callback_data="cancel_action")]])


def get_settings_menu():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📅 График дня (начало/конец)", callback_data="action_schedule")],
        [InlineKeyboardButton(text="🗓 Рабочие дни недели", callback_data="action_work_days")],
        [InlineKeyboardButton(text="📍 Геозона отметки", callback_data="action_location")],
        [InlineKeyboardButton(text="💰 Оклад (Зарплата)", callback_data="action_salary")]
    ])


def get_target_menu():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="👥 Всем", callback_data="target_all")],
        [InlineKeyboardButton(text="🏢 Отделу (выбрать из списка)", callback_data="target_dept")],
        [InlineKeyboardButton(text="👤 Сотруднику (выбрать из списка)", callback_data="target_user")],
        [InlineKeyboardButton(text="🔙 Отмена", callback_data="cancel_action")]
    ])


def get_departments_keyboard(prefix="dept"):
    departments = get_unique_values('department')
    buttons = [[InlineKeyboardButton(text=f"🏢 {d}", callback_data=f"{prefix}::{d}")] for d in departments]
    buttons.append([InlineKeyboardButton(text="🔙 Отмена", callback_data="cancel_action")])
    return InlineKeyboardMarkup(inline_keyboard=buttons)


def get_employees_keyboard(users, prefix="emp"):
    buttons = []
    for u in users:
        name = u.get('full_name') or u.get('username') or u.get('phone_number') or "Без имени"
        job_title = u.get('job_title', '')
        uid = u['user_id']
        display = f"{name} ({job_title})" if job_title else name
        buttons.append([InlineKeyboardButton(text=display, callback_data=f"{prefix}::{uid}")])
    buttons.append([InlineKeyboardButton(text="🔙 Отмена", callback_data="cancel_action")])
    return InlineKeyboardMarkup(inline_keyboard=buttons)


def get_report_menu():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="💰 Расчет Зарплаты", callback_data="report_salary")],
        [InlineKeyboardButton(text="🔥 Переработки", callback_data="report_overtime")],
        [InlineKeyboardButton(text="⚠️ Опоздания", callback_data="report_late")],
        [InlineKeyboardButton(text="🔙 Отмена", callback_data="cancel_action")]
    ])


def get_filter_buttons(items, callback_prefix):
    buttons = [[InlineKeyboardButton(text="Все", callback_data=f"{callback_prefix}_Все")]]
    for item in items:
        buttons.append([InlineKeyboardButton(text=str(item), callback_data=f"{callback_prefix}::{item}")])
    buttons.append([InlineKeyboardButton(text="🔙 Отмена", callback_data="cancel_action")])
    return InlineKeyboardMarkup(inline_keyboard=buttons)


def get_period_keyboard():
    current_month = datetime.now().strftime('%Y-%m')
    available_months = get_available_months()
    buttons = [[InlineKeyboardButton(text=f"📅 {format_month_display(current_month)} (текущий)",
                                     callback_data=f"period_{current_month}")]]
    for month in available_months:
        if month != current_month:
            buttons.append([InlineKeyboardButton(text=format_month_display(month), callback_data=f"period_{month}")])
    buttons.append([InlineKeyboardButton(text="🔙 Отмена", callback_data="cancel_action")])
    return InlineKeyboardMarkup(inline_keyboard=buttons)


def get_users_list_keyboard(users, prefix):
    buttons = []
    for u in users:
        name = u.get('full_name') or u.get('username') or u.get('phone_number') or "Без имени"
        job_title = u.get('job_title', '')
        uid = u['user_id']
        display = f"{name} ({job_title})" if job_title else name
        buttons.append([InlineKeyboardButton(text=display, callback_data=f"{prefix}::{uid}")])
    buttons.append([InlineKeyboardButton(text=" Отмена", callback_data="cancel_action")])
    return InlineKeyboardMarkup(inline_keyboard=buttons)


def get_shifts_list_keyboard(shifts, prefix, work_days_str='1,2,3,4,5', has_current_shift=False):
    buttons = []
    if has_current_shift:
        buttons.append([InlineKeyboardButton(text="⚡ ТЕКУЩАЯ СМЕНА (исправить)", callback_data=f"{prefix}_current")])
    for s in shifts:
        try:
            start_dt = datetime.fromisoformat(s[2])
            end_dt = datetime.fromisoformat(s[3])
            label = f"{start_dt.strftime('%d.%m %H:%M')} - {end_dt.strftime('%H:%M')} ({s[4] // 60}ч)"
            if not is_working_day(start_dt, work_days_str): label += " "
            if safe_shift_value(s, 6, 0) == 1:
                late_min = safe_shift_value(s, 7, 0)
                label += f" ⚠️{format_duration(late_min)}"
            buttons.append([InlineKeyboardButton(text=label, callback_data=f"{prefix}::{s[0]}")])
        except Exception:
            continue
    buttons.append([InlineKeyboardButton(text="🔙 Отмена", callback_data="cancel_action")])
    return InlineKeyboardMarkup(inline_keyboard=buttons)


def get_current_shift_actions_keyboard():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="️ Изменить время начала", callback_data="csa_change_start")],
        [InlineKeyboardButton(text="🛑 Завершить смену (указать время)", callback_data="csa_finish_shift")],
        [InlineKeyboardButton(text="🔙 Отмена", callback_data="cancel_action")]
    ])


def get_work_days_keyboard(selected_days):
    buttons = []
    row = []
    for i in range(1, 8):
        mark = "✅" if i in selected_days else "⬜"
        row.append(InlineKeyboardButton(text=f"{mark} {DAYS_MAP[i]}", callback_data=f"wd_toggle_{i}"))
        if i == 4 or i == 7:
            buttons.append(row)
            row = []
    buttons.append([InlineKeyboardButton(text=" Сохранить", callback_data="wd_save")])
    buttons.append([InlineKeyboardButton(text=" Отмена", callback_data="cancel_action")])
    return InlineKeyboardMarkup(inline_keyboard=buttons)


def get_share_contact_keyboard():
    return ReplyKeyboardMarkup(resize_keyboard=True,
                               keyboard=[[KeyboardButton(text="📞 Подтвердить номер", request_contact=True)]])


# ================= ХЕНДЛЕРЫ =================

@dp.callback_query(F.data == "cancel_action")
async def cancel_action_handler(call: types.CallbackQuery, state: FSMContext):
    await state.clear()
    try:
        await call.message.edit_text(" Действие отменено.")
    except Exception:
        pass
    await call.answer()
    if is_admin(call.from_user.id):
        await call.message.answer("Главное меню:", reply_markup=get_admin_main_keyboard())
    else:
        await call.message.answer("Главное меню:", reply_markup=get_employee_keyboard())


@dp.message(Command("start"))
async def cmd_start(message: types.Message, state: FSMContext):
    try:
        await state.clear()
        user_id = message.from_user.id
        username = message.from_user.username
        full_name = message.from_user.full_name
        db_user = get_user_by_id(user_id)
        if db_user:
            update_user_data(user_id, username=username or db_user.get('username'), full_name=full_name)
            if user_id == ADMIN_ID or db_user.get('role') == 'admin':
                await message.answer("Панель Администратора.", reply_markup=get_admin_main_keyboard())
            else:
                work_days_text = format_work_days(db_user.get('work_days_week') or '1,2,3,4,5')
                await message.answer(
                    f"Привет! График: {db_user.get('schedule_start')}-{db_user.get('schedule_end')}\n"
                    f"Рабочие дни: {work_days_text}\n"
                    f"💡 Можно начать смену в любой день\n"
                    f" Напоминания за 10 мин до начала/конца\n"
                    f"️ Защита от фейкового GPS активна",
                    reply_markup=get_employee_keyboard()
                )
            return
        if message.contact and message.contact.user_id == user_id:
            phone_user = get_user_by_phone(message.contact.phone_number)
            if phone_user:
                await _activate_by_contact(message, phone_user)
                return
        await message.answer("Вы не найдены в базе.\nЕсли вас добавили по номеру, нажмите кнопку ниже.",
                             reply_markup=get_share_contact_keyboard())
    except Exception as e:
        logging.error(f"Ошибка в cmd_start: {e}", exc_info=True)
        await message.answer("Произошла ошибка. Попробуйте еще раз.")


async def _activate_by_contact(message: types.Message, phone_user: dict):
    try:
        real_id = message.from_user.id
        ghost_id = phone_user['user_id']
        username = message.from_user.username
        full_name = message.from_user.full_name
        if ghost_id < 0:
            real_user = get_user_by_id(real_id)
            if real_user:
                update_user_data(real_id, department=phone_user.get('department'),
                                 job_title=phone_user.get('job_title'),
                                 phone_number=phone_user.get('phone_number'), full_name=full_name, username=username,
                                 work_days_week=phone_user.get('work_days_week') or '1,2,3,4,5')
                conn = sqlite3.connect('clockster.db')
                conn.cursor().execute("DELETE FROM users WHERE user_id = ?", (ghost_id,))
                conn.commit();
                conn.close()
                await message.answer("✅ Профиль активирован!", reply_markup=get_employee_keyboard())
            else:
                update_user_data(ghost_id, user_id=real_id, username=username, full_name=full_name)
                await message.answer("✅ Вы успешно привязаны!", reply_markup=get_employee_keyboard())
        else:
            update_user_data(ghost_id, username=username, full_name=full_name)
            await message.answer("Вы в системе.", reply_markup=get_employee_keyboard())
    except Exception as e:
        logging.error(f"Ошибка в _activate_by_contact: {e}", exc_info=True)
        await message.answer("Произошла ошибка при активации.")


@dp.message(F.content_type == "contact")
async def handle_user_contact(message: types.Message):
    try:
        contact = message.contact
        if not contact: return
        if contact.user_id and contact.user_id != message.from_user.id:
            await message.answer("❌ Поделитесь своим номером.");
            return
        phone_user = get_user_by_phone(contact.phone_number)
        if phone_user:
            await _activate_by_contact(message, phone_user)
        else:
            await message.answer("❌ Номер не найден.")
    except Exception as e:
        logging.error(f"Ошибка в handle_user_contact: {e}", exc_info=True)


# --- ЖУРНАЛ GPS (АДМИН) ---
@dp.message(F.text == "🛡️ Журнал GPS")
async def show_gps_log(message: types.Message):
    try:
        if not is_admin(message.from_user.id): return

        conn = sqlite3.connect('clockster.db')
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM gps_log ORDER BY id DESC LIMIT 20")
        logs = [dict(r) for r in cursor.fetchall()]
        conn.close()

        if not logs:
            await message.answer("🛡️ Журнал GPS пуст. Подозрительных попыток не было.",
                                 reply_markup=get_admin_main_keyboard())
            return

        text = "🛡️ <b>Последние 20 попыток GPS:</b>\n\n"
        for log in logs:
            try:
                user = get_user_by_id(log['user_id'])
                name = user.get('full_name') or user.get('username') if user else f"ID {log['user_id']}"
                ts = datetime.fromisoformat(log['timestamp']).strftime('%d.%m %H:%M')
                status = "🚫 ЗАБЛОКИРОВАНО" if log['action_taken'] == 'blocked' else "⚠️ ПРЕДУПРЕЖДЕНИЕ"
                text += f"• <b>{name}</b> ({ts})\n"
                text += f"  {status}: {log['reason']}\n"
                if log.get('distance_from_office'):
                    text += f"  📍 Расстояние от офиса: {int(log['distance_from_office'])}м\n"
                text += "\n"
            except Exception:
                continue

        await message.answer(text, parse_mode="HTML", reply_markup=get_admin_main_keyboard())
    except Exception as e:
        logging.error(f"Ошибка в show_gps_log: {e}", exc_info=True)
        await message.answer("Произошла ошибка при загрузке журнала.")


# --- УДАЛЕНИЕ ДУБЛИКАТОВ ---
@dp.message(F.text == "️ Удалить дубликаты")
async def remove_duplicates_cmd(message: types.Message):
    try:
        if not is_admin(message.from_user.id): return
        deleted_count = remove_duplicate_shifts()
        if deleted_count > 0:
            await message.answer(f"✅ Удалено {deleted_count} дубликатов смен.", reply_markup=get_admin_main_keyboard())
        else:
            await message.answer("✅ Дубликатов не найдено.", reply_markup=get_admin_main_keyboard())
    except Exception as e:
        logging.error(f"Ошибка в remove_duplicates_cmd: {e}", exc_info=True)


# --- СБРОС СОСТОЯНИЯ ---
@dp.message(F.text == "🔄 Сброс состояния")
async def reset_state_cmd(message: types.Message, state: FSMContext):
    try:
        if not is_admin(message.from_user.id): return
        await state.clear()
        await message.answer("✅ Состояние сброшено.", reply_markup=get_admin_main_keyboard())
    except Exception as e:
        logging.error(f"Ошибка в reset_state_cmd: {e}", exc_info=True)


# --- МОЯ СТАТИСТИКА (ИСПРАВЛЕНО) ---
@dp.message(F.text == "📊 Моя статистика")
async def cmd_my_stats(message: types.Message):
    try:
        user = get_user_by_id(message.from_user.id)
        if not user:
            await message.answer("❌ Вы не найдены в базе.")
            return

        cm = datetime.now().strftime('%Y-%m')
        mn = format_month_display(cm)
        shifts = get_unique_shifts(get_user_shifts(user['user_id'], month_filter=cm))
        wds = user.get('work_days_week') or '1,2,3,4,5'
        y, m = map(int, cm.split('-'))
        wn = get_working_days_in_month(y, m, wds)
        wc, wkc = count_shifts_by_day_type(shifts, wds)

        t = f"👤 <b>{user.get('full_name') or user.get('username')}</b>\n"
        t += f" {mn}\n"
        t += f"📅 Рабочие: {format_work_days(wds)}\n"
        t += f"📊 Норма: {wn}\n\n"

        if shifts:
            tm = 0;
            lt = 0;
            lm = 0
            for s in shifts:
                tm += safe_shift_value(s, 4, 0)
                if safe_shift_value(s, 6, 0) == 1:
                    lt += 1
                    lm += safe_shift_value(s, 7, 0)
            t += f"✅ Всего: {len(shifts)}\n"
            t += f"   • Рабочие: {wc}\n"
            t += f"   • 🌟 Выходные: {wkc}\n"
            t += f"⏱ Часов: {tm // 60}ч {tm % 60}м\n"
            t += f"⚠️ Опозданий: {lt}"
            if lm > 0:
                t += f"\n⏱ Общее время опозданий: {format_duration(lm)}"
            t += "\n\n<b>История:</b>\n"
            for s in shifts:
                try:
                    sd = datetime.fromisoformat(s[2])
                    ed = datetime.fromisoformat(s[3])
                    dr = f"{s[4] // 60}ч {s[4] % 60}м"
                    mk = ""
                    if not is_working_day(sd, wds): mk += " 🌟"
                    if safe_shift_value(s, 6, 0) == 1:
                        mk += f" ⚠️{format_duration(safe_shift_value(s, 7, 0))}"
                    t += f"📅 {sd.strftime('%H:%M %d.%m')} ({DAYS_MAP[sd.weekday() + 1]}) - {ed.strftime('%H:%M')} ({dr}){mk}\n"
                except:
                    continue
        else:
            t += "Смен нет."

        kb = get_admin_main_keyboard() if is_admin(message.from_user.id) else get_employee_keyboard()
        await message.answer(t, parse_mode="HTML", reply_markup=kb)
    except Exception as e:
        logging.error(f"Ошибка в cmd_my_stats: {e}", exc_info=True)
        await message.answer("❌ Произошла ошибка при загрузке статистики.")


# --- СПИСОК СОТРУДНИКОВ (ИСПРАВЛЕНО) ---
@dp.message(F.text == "📋 Список сотрудников")
async def show_dashboard_menu(message: types.Message, state: FSMContext):
    try:
        if not is_admin(message.from_user.id): return
        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text=" Все сотрудники", callback_data="dash_all")],
            [InlineKeyboardButton(text=" Выбрать отдел", callback_data="dash_dept")]
        ])
        await message.answer("Показать:", reply_markup=kb)
        await state.set_state(DashboardState.choosing_filter)
    except Exception as e:
        logging.error(f"Ошибка в show_dashboard_menu: {e}", exc_info=True)
        await message.answer("❌ Произошла ошибка.")


@dp.callback_query(DashboardState.choosing_filter, F.data == "dash_all")
async def show_all_employees(call: types.CallbackQuery, state: FSMContext):
    try:
        if not is_admin(call.from_user.id):
            await call.answer()
            return
        text = "📋 <b>Все сотрудники:</b>\n\n" + _format_employees_list(get_all_employees())
        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text=" Фильтр", callback_data="dash_dept")],
            [InlineKeyboardButton(text=" Назад", callback_data="cancel_action")]
        ])
        await call.message.edit_text(text, reply_markup=kb, parse_mode="HTML")
        await call.answer()
    except Exception as e:
        logging.error(f"Ошибка в show_all_employees: {e}", exc_info=True)
        await call.answer("Произошла ошибка")


@dp.callback_query(DashboardState.choosing_filter, F.data == "dash_dept")
async def choose_dept_dash(call: types.CallbackQuery, state: FSMContext):
    try:
        if not is_admin(call.from_user.id):
            await call.answer()
            return
        await call.message.edit_text("Отдел:", reply_markup=get_departments_keyboard("dash_dept"))
        await call.answer()
    except Exception as e:
        logging.error(f"Ошибка в choose_dept_dash: {e}", exc_info=True)
        await call.answer("Произошла ошибка")


@dp.callback_query(DashboardState.choosing_filter, F.data.startswith("dash_dept::"))
async def show_dept_emps(call: types.CallbackQuery, state: FSMContext):
    try:
        if not is_admin(call.from_user.id):
            await call.answer()
            return
        dept = parse_callback_value(call.data, "dash_dept")
        text = f" <b>«{dept}»:</b>\n\n" + _format_employees_list(get_users_by_filters(department=dept))
        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="👥 Все", callback_data="dash_all")],
            [InlineKeyboardButton(text="🏢 Другой", callback_data="dash_dept")],
            [InlineKeyboardButton(text="🔙 Назад", callback_data="cancel_action")]
        ])
        await call.message.edit_text(text, reply_markup=kb, parse_mode="HTML")
        await call.answer()
    except Exception as e:
        logging.error(f"Ошибка в show_dept_emps: {e}", exc_info=True)
        await call.answer("Произошла ошибка")


def _format_employees_list(emps):
    text = ""
    count = 0
    for e in emps:
        try:
            if e['user_id'] < 0: continue
            count += 1
            iw = bool(e.get('is_working'))
            st = "🟢" if iw else "🔴"
            nm = e.get('full_name') or e.get('username') or e.get('phone_number') or "?"
            rl = " 👑" if e.get('role') == 'admin' else ""
            jt = e.get('job_title', '')
            wdt = format_work_days(e.get('work_days_week') or '1,2,3,4,5')
            sal = e.get('monthly_salary') or 0
            ls = get_unique_shifts(get_user_shifts(e['user_id']))
            wds = e.get('work_days_week') or '1,2,3,4,5'

            if iw and e.get('shift_start_time'):
                try:
                    so = datetime.fromisoformat(e['shift_start_time'])
                    wm = " 🌟" if not is_working_day(so, wds) else ""
                    si = f"⏰ Начал {so.strftime('%H:%M %d.%m')} · на смене{wm}"
                except ValueError:
                    si = "⏰ На смене"
            elif ls:
                try:
                    so = datetime.fromisoformat(ls[0][2])
                    eo = datetime.fromisoformat(ls[0][3])
                    wm = " 🌟" if not is_working_day(so, wds) else ""
                    late_mark = ""
                    if safe_shift_value(ls[0], 6, 0) == 1:
                        late_mark = f" ⚠️{format_duration(safe_shift_value(ls[0], 7, 0))}"
                    si = f"⏰ {so.strftime('%H:%M %d.%m')} → {eo.strftime('%H:%M %d.%m')}{wm}{late_mark}"
                except ValueError:
                    si = "⏰ Данные"
            else:
                si = "Смен не было"

            li = "📍 Задана" if (e.get('target_lat') and e.get('target_lon')) else "⚠️ Нет зоны"
            sai = f"💰 {format_money(sal)}" if sal > 0 else "💰 Не указан"
            nd = f"{nm}{rl} ({jt})" if jt else f"{nm}{rl}"

            text += f"{st} <b>{nd}</b>\n"
            text += f"🏢 {e.get('department')} | 🗓 {wdt} | ⏰ {e.get('schedule_start')}-{e.get('schedule_end')}\n"
            text += f"{sai} | {li}\n"
            text += f"{si}\n\n"
        except Exception as e:
            logging.error(f"Ошибка форматирования сотрудника {e.get('user_id', '?')}: {e}")
            continue

    if count == 0:
        text = "Пусто."
    return text


# --- ДОБАВЛЕНИЕ СОТРУДНИКА ---
@dp.message(F.text == "➕ Добавить сотрудника")
async def start_add_employee(message: types.Message, state: FSMContext):
    try:
        if not is_admin(message.from_user.id): return
        await message.answer("Введите @username ИЛИ номер телефона:", reply_markup=get_cancel_keyboard())
        await state.set_state(AddEmployeeState.input_identifier)
    except Exception as e:
        logging.error(f"Ошибка в start_add_employee: {e}", exc_info=True)


@dp.message(AddEmployeeState.input_identifier)
async def process_add_identifier(message: types.Message, state: FSMContext):
    try:
        text = message.text.strip()
        digits_only = ''.join(filter(str.isdigit, text))
        is_phone = len(digits_only) >= 9 or text.startswith('+')
        if is_phone:
            id_type, clean_id = 'phone', digits_only
        else:
            id_type, clean_id = 'username', text.lstrip('@')
        exists = (get_user_by_phone(clean_id) if is_phone else get_user_by_username(clean_id)) is not None
        if exists:
            await message.answer("❌ Уже есть.", reply_markup=get_cancel_keyboard())
            return
        await state.update_data(identifier=clean_id, id_type=id_type)
        await message.answer("Введите Отдел:", reply_markup=get_cancel_keyboard())
        await state.set_state(AddEmployeeState.input_dept)
    except Exception as e:
        logging.error(f"Ошибка в process_add_identifier: {e}", exc_info=True)


@dp.message(AddEmployeeState.input_dept)
async def process_add_dept(message: types.Message, state: FSMContext):
    try:
        await state.update_data(department=message.text)
        await message.answer("Введите Должность:", reply_markup=get_cancel_keyboard())
        await state.set_state(AddEmployeeState.input_title)
    except Exception as e:
        logging.error(f"Ошибка в process_add_dept: {e}", exc_info=True)


@dp.message(AddEmployeeState.input_title)
async def process_add_title(message: types.Message, state: FSMContext):
    try:
        await state.update_data(title=message.text)
        await message.answer("Выберите рабочие дни:", reply_markup=get_work_days_keyboard([1, 2, 3, 4, 5]))
        await state.update_data(temp_work_days=[1, 2, 3, 4, 5])
        await state.set_state(AddEmployeeState.selecting_work_days)
    except Exception as e:
        logging.error(f"Ошибка в process_add_title: {e}", exc_info=True)


@dp.callback_query(AddEmployeeState.selecting_work_days, F.data.startswith("wd_toggle_"))
async def toggle_work_day_add(call: types.CallbackQuery, state: FSMContext):
    try:
        day = int(parse_callback_value(call.data, "wd_toggle"))
        data = await state.get_data()
        selected = data.get('temp_work_days', [1, 2, 3, 4, 5])
        if day in selected:
            selected.remove(day)
        else:
            selected.append(day); selected.sort()
        await state.update_data(temp_work_days=selected)
        await call.message.edit_reply_markup(reply_markup=get_work_days_keyboard(selected))
        await call.answer()
    except Exception as e:
        logging.error(f"Ошибка в toggle_work_day_add: {e}", exc_info=True)


@dp.callback_query(AddEmployeeState.selecting_work_days, F.data == "wd_save")
async def save_add_employee(call: types.CallbackQuery, state: FSMContext):
    try:
        data = await state.get_data()
        identifier, id_type = data['identifier'], data['id_type']
        dept, title = data['department'], data['title']
        work_days_str = ",".join(map(str, data.get('temp_work_days', [1, 2, 3, 4, 5])))
        conn = sqlite3.connect('clockster.db')
        cursor = conn.cursor()
        try:
            if id_type == 'username':
                cursor.execute(
                    "INSERT INTO users (username, role, department, job_title, work_days_week) VALUES (?, 'employee', ?, ?, ?)",
                    (identifier, dept, title, work_days_str))
                msg = f"@{identifier}"
            else:
                temp_id = -1 * int(datetime.now().timestamp() * 1000)
                cursor.execute(
                    "INSERT INTO users (user_id, username, phone_number, role, department, job_title, work_days_week) VALUES (?, ?, ?, 'employee', ?, ?, ?)",
                    (temp_id, identifier, identifier, dept, title, work_days_str))
                msg = f"Номер: {identifier}"
            conn.commit()
            await call.message.edit_text(
                f"✅ Добавлен ({msg}, {dept} - {title})\nРабочие дни: {format_work_days(work_days_str)}")
            await call.message.answer("Меню:", reply_markup=get_admin_main_keyboard())
        except sqlite3.IntegrityError:
            await call.message.edit_text("❌ Уже существует.")
            await call.message.answer("Меню:", reply_markup=get_admin_main_keyboard())
        finally:
            conn.close()
        await state.clear()
        await call.answer()
    except Exception as e:
        logging.error(f"Ошибка в save_add_employee: {e}", exc_info=True)


# --- УДАЛЕНИЕ СОТРУДНИКА ---
@dp.message(F.text == "🗑 Удалить сотрудника")
async def start_delete_employee(message: types.Message, state: FSMContext):
    try:
        if not is_admin(message.from_user.id): return
        emps = get_all_employees()
        if not emps:
            await message.answer("Пусто.", reply_markup=get_admin_main_keyboard())
            return
        await message.answer("Выберите:", reply_markup=get_users_list_keyboard(emps, "del_emp"))
        await state.set_state(DeleteEmployeeState.selecting_user)
    except Exception as e:
        logging.error(f"Ошибка в start_delete_employee: {e}", exc_info=True)


@dp.callback_query(DeleteEmployeeState.selecting_user, F.data.startswith("del_emp::"))
async def select_employee_to_delete(call: types.CallbackQuery, state: FSMContext):
    try:
        uid = int(parse_callback_value(call.data, "del_emp"))
        user = get_user_by_id(uid)
        if not user:
            await call.message.edit_text("Ошибка.")
            await state.clear()
            await call.answer()
            return
        await state.update_data(delete_user_id=uid)
        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="✅ Да", callback_data="confirm_del_user")],
            [InlineKeyboardButton(text="❌ Отмена", callback_data="cancel_action")]
        ])
        await call.message.edit_text(f"Удалить <b>{user.get('full_name') or user.get('username')}</b>?",
                                     reply_markup=kb, parse_mode="HTML")
        await call.answer()
    except ValueError:
        await call.answer("Ошибка", show_alert=True)
    except Exception as e:
        logging.error(f"Ошибка в select_employee_to_delete: {e}", exc_info=True)


@dp.callback_query(F.data == "confirm_del_user")
async def confirm_delete_user(call: types.CallbackQuery, state: FSMContext):
    try:
        if not is_admin(call.from_user.id):
            await call.answer()
            return
        data = await state.get_data()
        if data.get('delete_user_id') and delete_user_by_id(data.get('delete_user_id')):
            await call.message.edit_text("✅ Удалено.")
        else:
            await call.message.edit_text("❌ Ошибка.")
        await state.clear()
        await call.message.answer("Меню:", reply_markup=get_admin_main_keyboard())
        await call.answer()
    except Exception as e:
        logging.error(f"Ошибка в confirm_delete_user: {e}", exc_info=True)


# --- УДАЛЕНИЕ/НАЗНАЧЕНИЕ АДМИНА ---
@dp.message(F.text == "🗑 Удалить Админа")
async def start_delete_admin(message: types.Message, state: FSMContext):
    try:
        if not is_admin(message.from_user.id): return
        admins = get_all_admins()
        if not admins:
            await message.answer("Нет админов.", reply_markup=get_admin_main_keyboard())
            return
        await message.answer("Выберите:", reply_markup=get_users_list_keyboard(admins, "del_adm"))
        await state.set_state(DeleteAdminState.selecting_user)
    except Exception as e:
        logging.error(f"Ошибка в start_delete_admin: {e}", exc_info=True)


@dp.callback_query(DeleteAdminState.selecting_user, F.data.startswith("del_adm::"))
async def select_admin_to_delete(call: types.CallbackQuery, state: FSMContext):
    try:
        uid = int(parse_callback_value(call.data, "del_adm"))
        user = get_user_by_id(uid)
        await state.update_data(delete_user_id=uid)
        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="✅ Да", callback_data="confirm_del_admin")],
            [InlineKeyboardButton(text="❌ Отмена", callback_data="cancel_action")]
        ])
        await call.message.edit_text(f"Удалить админа <b>{user.get('full_name')}</b>?", reply_markup=kb,
                                     parse_mode="HTML")
        await call.answer()
    except ValueError:
        await call.answer("Ошибка", show_alert=True)
    except Exception as e:
        logging.error(f"Ошибка в select_admin_to_delete: {e}", exc_info=True)


@dp.callback_query(F.data == "confirm_del_admin")
async def confirm_delete_admin(call: types.CallbackQuery, state: FSMContext):
    try:
        if not is_admin(call.from_user.id):
            await call.answer()
            return
        data = await state.get_data()
        if data.get('delete_user_id') and delete_user_by_id(data.get('delete_user_id')):
            await call.message.edit_text("✅ Админ удалён.")
        await state.clear()
        await call.message.answer("Меню:", reply_markup=get_admin_main_keyboard())
        await call.answer()
    except Exception as e:
        logging.error(f"Ошибка в confirm_delete_admin: {e}", exc_info=True)


@dp.message(F.text == "➕ Назначить Админа")
async def start_add_admin(message: types.Message, state: FSMContext):
    try:
        if not is_admin(message.from_user.id): return
        if get_admin_count() >= MAX_ADMINS:
            await message.answer(f"❌ Максимум ({MAX_ADMINS}).", reply_markup=get_admin_main_keyboard())
            return
        emps = [e for e in get_all_employees() if e.get('role') != 'admin']
        if not emps:
            await message.answer("Нет сотрудников.", reply_markup=get_admin_main_keyboard())
            return
        await message.answer("Выберите:", reply_markup=get_users_list_keyboard(emps, "addadm"))
        await state.set_state(AddAdminState.input_identifier)
    except Exception as e:
        logging.error(f"Ошибка в start_add_admin: {e}", exc_info=True)


@dp.callback_query(AddAdminState.input_identifier, F.data.startswith("addadm::"))
async def process_add_admin_identifier(call: types.CallbackQuery, state: FSMContext):
    try:
        uid = int(parse_callback_value(call.data, "addadm"))
        user = get_user_by_id(uid)
        if not user:
            await call.message.edit_text("Не найден.")
            await state.clear()
            await call.answer()
            return
        await state.update_data(target_user_id=uid)
        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="✅ Да", callback_data="confirm_add_admin")],
            [InlineKeyboardButton(text=" Отмена", callback_data="cancel_action")]
        ])
        await call.message.edit_text(f"Назначить <b>{user.get('full_name') or user.get('username')}</b> админом?",
                                     reply_markup=kb, parse_mode="HTML")
        await state.set_state(AddAdminState.confirm)
        await call.answer()
    except ValueError:
        await call.answer("Ошибка", show_alert=True)
    except Exception as e:
        logging.error(f"Ошибка в process_add_admin_identifier: {e}", exc_info=True)


@dp.callback_query(AddAdminState.confirm, F.data == "confirm_add_admin")
async def confirm_add_admin(call: types.CallbackQuery, state: FSMContext):
    try:
        if not is_admin(call.from_user.id) or get_admin_count() >= MAX_ADMINS:
            await call.answer()
            await state.clear()
            return
        data = await state.get_data()
        if data.get('target_user_id'):
            update_user_data(data['target_user_id'], role='admin')
            await call.message.edit_text("✅ Назначен.")
        await state.clear()
        await call.message.answer("Меню:", reply_markup=get_admin_main_keyboard())
        await call.answer()
    except Exception as e:
        logging.error(f"Ошибка в confirm_add_admin: {e}", exc_info=True)


# --- ОТЧЕТЫ ---
@dp.message(F.text == "📊 Отчеты и Excel")
async def open_reports(message: types.Message, state: FSMContext):
    try:
        if not is_admin(message.from_user.id): return
        await message.answer("Тип отчета:", reply_markup=get_report_menu())
        await state.set_state(ReportState.choosing_type)
    except Exception as e:
        logging.error(f"Ошибка в open_reports: {e}", exc_info=True)


@dp.callback_query(ReportState.choosing_type, F.data.startswith("report_"))
async def choose_report_type(call: types.CallbackQuery, state: FSMContext):
    try:
        await state.update_data(report_type=parse_callback_value(call.data, "report"))
        await call.message.edit_text("Выберите отдел:",
                                     reply_markup=get_filter_buttons(get_unique_values('department'), "dept"))
        await state.set_state(ReportState.filter_dept)
    except Exception as e:
        logging.error(f"Ошибка в choose_report_type: {e}", exc_info=True)
    finally:
        try:
            await call.answer()
        except:
            pass


@dp.callback_query(ReportState.filter_dept, F.data.startswith("dept"))
async def filter_dept(call: types.CallbackQuery, state: FSMContext):
    try:
        dept = parse_callback_value(call.data, "dept")
        await state.update_data(department="Все" if dept == "Все" else dept)
        logging.info(f"🏢 Выбран отдел: {dept}")
        await call.message.edit_text("Выберите период:", reply_markup=get_period_keyboard())
        await state.set_state(ReportState.choosing_period)
    except Exception as e:
        logging.error(f"❌ Ошибка в filter_dept: {e}", exc_info=True)
        try:
            await call.message.edit_text(f"❌ Ошибка: {e}", reply_markup=get_cancel_keyboard())
        except:
            pass
    finally:
        try:
            await call.answer()
        except:
            pass


@dp.callback_query(ReportState.choosing_period, F.data.startswith("period_"))
async def choose_period(call: types.CallbackQuery, state: FSMContext):
    try:
        period = parse_callback_value(call.data, "period")
        await state.update_data(period=period)
        logging.info(f"📅 Выбран период: {period}")
        await generate_report(call, state)
    except Exception as e:
        logging.error(f"❌ Ошибка в choose_period: {e}", exc_info=True)
        try:
            await call.message.edit_text(f"❌ Ошибка при формировании отчёта:\n{e}", reply_markup=get_cancel_keyboard())
        except:
            pass
    finally:
        try:
            await call.answer()
        except:
            pass


async def generate_report(call: types.CallbackQuery, state: FSMContext):
    try:
        data = await state.get_data()
        period = data.get('period')
        department = data.get('department', 'Все')
        report_type = data.get('report_type', 'salary')

        users = get_users_by_filters(department=department if department != 'Все' else None)
        report_data, total_pay, total_overtime, total_lates, total_weekend, total_late_minutes = [], 0, 0, 0, 0, 0

        for u in users:
            shifts = get_unique_shifts(get_user_shifts(u['user_id'], month_filter=period))
            wds = u.get('work_days_week') or '1,2,3,4,5'
            ms = u.get('monthly_salary') or 0
            ad = len(shifts)

            om = 0;
            lt = 0;
            lm = 0
            if shifts:
                for s in shifts:
                    try:
                        om += safe_shift_value(s, 5, 0)
                        if safe_shift_value(s, 6, 0) == 1:
                            lt += 1
                            lm += safe_shift_value(s, 7, 0)
                    except Exception:
                        pass

            total_late_minutes += lm
            wc, wkc = count_shifts_by_day_type(shifts, wds)
            total_weekend += wkc
            nm = u.get('full_name') or u.get('username') or u.get('phone_number') or "Без имени"

            row = {
                "ФИО Сотрудника": nm,
                "Отдел": u.get('department'),
                "Должность": u.get('job_title'),
                "Рабочие дни": format_work_days(wds),
                "Всего смен": ad,
                "Смен в раб. дни": wc,
                "Смен в выходные": wkc,
                "Переработка": format_duration(om),
                "Кол-во опозданий": lt,
                "Время опозданий": format_duration(lm)
            }

            if report_type == 'salary':
                fp, bp, dr, norm = calculate_salary(u, shifts, period)
                row.update({
                    "Оклад (₸)": ms,
                    "Норма дней в мес.": norm,
                    "Дневная ставка (₸)": round(dr, 2),
                    "Итого к выплате (₸)": round(fp, 2)
                })
                total_pay += fp

            report_data.append(row)
            total_overtime += om
            total_lates += lt

        ft = f"Отдел: {department}" if department != 'Все' else "Отдел: Все"
        tt = {'salary': ' Расчет Зарплаты', 'overtime': ' Переработки', 'late': '️ Опоздания'}
        title = tt.get(report_type, 'Отчет')

        st = f"📊 <b>Отчет: {title}</b>\n🗓 Период: {format_month_display(period)}\n🔍 {ft}\n Сотрудников: {len(users)}\n"

        if report_type == 'salary':
            st += f"💵 <b>Итого: {format_money(total_pay)}</b>\n Выходных смен: {total_weekend}"
        elif report_type == 'overtime':
            st += f"🔥 <b>Переработки: {format_duration(total_overtime)}</b>"
        elif report_type == 'late':
            st += f"⚠️ <b>Опозданий: {total_lates}</b>\n⏱ <b>Общее время опозданий: {format_duration(total_late_minutes)}</b>"

        await state.update_data(excel_data=report_data, filename=f"report_{report_type}_{period}.xlsx")

        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="📥 Скачать Excel", callback_data="download_excel")],
            [InlineKeyboardButton(text="🔙 В меню", callback_data="cancel_action")]
        ])

        try:
            await call.message.edit_text(st, reply_markup=kb, parse_mode="HTML")
        except Exception as e:
            logging.error(f"Ошибка edit_text в generate_report: {e}")
            st_plain = re.sub(r'<[^>]+>', '', st)
            await call.message.edit_text(st_plain, reply_markup=kb)

        await state.set_state(ReportState.showing_result)
    except Exception as e:
        logging.error(f"Ошибка в generate_report: {e}", exc_info=True)


@dp.callback_query(ReportState.showing_result, F.data == "download_excel")
async def download_excel(call: types.CallbackQuery, state: FSMContext):
    try:
        data = await state.get_data()
        excel_data = data.get('excel_data') or []
        if not excel_data:
            await call.answer("Нет данных", show_alert=True)
            return
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df = pd.DataFrame(excel_data)
            df.to_excel(writer, index=False, sheet_name='Сводка')
            ws = writer.sheets['Сводка']
            for col in range(1, len(df.columns) + 1):
                ws[get_column_letter(col) + '1'].font = Font(bold=True)
            for col in ws.columns:
                ml = max((len(str(c.value)) for c in col if c.value), default=0)
                ws.column_dimensions[col[0].column_letter].width = min(ml + 2, 35)
        output.seek(0)
        await call.message.answer_document(
            BufferedInputFile(output.read(), filename=data.get('filename') or "report.xlsx"))
        await call.answer()
    except Exception as e:
        logging.error(f"Ошибка в download_excel: {e}", exc_info=True)


# --- ИСПРАВИТЬ СМЕНУ ---
@dp.message(F.text == "✏️ Исправить смену")
async def start_edit_shift(message: types.Message, state: FSMContext):
    try:
        if not is_admin(message.from_user.id): return
        await message.answer("Выберите отдел:", reply_markup=get_departments_keyboard("eshift_dept"))
        await state.set_state(EditShiftState.choosing_dept)
    except Exception as e:
        logging.error(f"Ошибка в start_edit_shift: {e}", exc_info=True)


@dp.callback_query(EditShiftState.choosing_dept, F.data.startswith("eshift_dept::"))
async def choose_dept_for_shift_edit(call: types.CallbackQuery, state: FSMContext):
    try:
        dept = parse_callback_value(call.data, "eshift_dept")
        await state.update_data(edit_dept=dept)
        users = get_users_by_filters(department=dept)
        if not users:
            await call.message.edit_text("Нет сотрудников.")
            await state.clear()
            await call.message.answer("Меню:", reply_markup=get_admin_main_keyboard())
            await call.answer()
            return
        await call.message.edit_text(f"Отдел <b>{dept}</b>. Выберите:",
                                     reply_markup=get_employees_keyboard(users, "eshift_user"), parse_mode="HTML")
        await state.set_state(EditShiftState.selecting_user)
        await call.answer()
    except Exception as e:
        logging.error(f"Ошибка в choose_dept_for_shift_edit: {e}", exc_info=True)


@dp.callback_query(EditShiftState.selecting_user, F.data.startswith("eshift_user::"))
async def select_user_for_shift_edit(call: types.CallbackQuery, state: FSMContext):
    try:
        uid = int(parse_callback_value(call.data, "eshift_user"))
        user = get_user_by_id(uid)
        await state.update_data(edit_user_id=uid)
        has_current_shift = bool(user.get('is_working') and user.get('shift_start_time'))
        shifts = get_unique_shifts(get_user_shifts(uid, month_filter=datetime.now().strftime('%Y-%m')))
        wds = user.get('work_days_week') or '1,2,3,4,5'
        if not shifts and not has_current_shift:
            await call.message.edit_text("Нет смен.")
            await state.clear()
            await call.message.answer("Меню:", reply_markup=get_admin_main_keyboard())
            await call.answer()
            return
        text = " - выходной | ⚠️ - опоздание\n"
        if has_current_shift:
            try:
                start_dt = datetime.fromisoformat(user.get('shift_start_time'))
                text += f"\n⚡ <b>Текущая смена:</b> начало {start_dt.strftime('%d.%m %H:%M')}\n"
            except Exception:
                text += f"\n⚡ <b>Текущая смена:</b> активна\n"
        await call.message.edit_text(
            text,
            reply_markup=get_shifts_list_keyboard(shifts, "eshift", wds, has_current_shift=has_current_shift),
            parse_mode="HTML"
        )
        await state.set_state(EditShiftState.selecting_shift)
        await call.answer()
    except ValueError:
        await call.answer("Ошибка", show_alert=True)
    except Exception as e:
        logging.error(f"Ошибка в select_user_for_shift_edit: {e}", exc_info=True)


@dp.callback_query(EditShiftState.selecting_shift, F.data == "eshift_current")
async def select_current_shift(call: types.CallbackQuery, state: FSMContext):
    try:
        data = await state.get_data()
        uid = data.get('edit_user_id')
        user = get_user_by_id(uid)
        if not user or not user.get('is_working') or not user.get('shift_start_time'):
            await call.message.edit_text("❌ Смена уже завершена.")
            await state.clear()
            await call.message.answer("Меню:", reply_markup=get_admin_main_keyboard())
            await call.answer()
            return
        try:
            start_dt = datetime.fromisoformat(user.get('shift_start_time'))
            start_str = start_dt.strftime('%d.%m.%Y %H:%M')
        except Exception:
            start_str = "неизвестно"
        user_name = user.get('full_name') or user.get('username') or "?"
        await call.message.edit_text(
            f" <b>Текущая смена</b>\n"
            f" {user_name}\n"
            f"⏰ Начало: {start_str}\n\n"
            f"Выберите действие:",
            reply_markup=get_current_shift_actions_keyboard(),
            parse_mode="HTML"
        )
        await state.set_state(EditShiftState.current_shift_action)
        await call.answer()
    except Exception as e:
        logging.error(f"Ошибка в select_current_shift: {e}", exc_info=True)


@dp.callback_query(EditShiftState.current_shift_action, F.data == "csa_change_start")
async def change_current_start(call: types.CallbackQuery, state: FSMContext):
    try:
        await call.message.edit_text(
            "Введите новое время НАЧАЛА текущей смены (ДД.ММ ЧЧ:ММ):\n"
            "Например: 12.09 07:30",
            reply_markup=get_cancel_keyboard()
        )
        await state.set_state(EditShiftState.editing_current_start)
        await call.answer()
    except Exception as e:
        logging.error(f"Ошибка в change_current_start: {e}", exc_info=True)


@dp.message(EditShiftState.editing_current_start)
async def input_current_new_start(message: types.Message, state: FSMContext):
    try:
        y = datetime.now().year
        t = message.text.strip()
        if len(t.split()[0].split('.')) == 2:
            new_start = datetime.strptime(f"{t} {y}", "%d.%m %H:%M %Y")
        else:
            new_start = datetime.strptime(t, "%d.%m.%y %H:%M")
        data = await state.get_data()
        uid = data.get('edit_user_id')
        user = get_user_by_id(uid)
        if not user or not user.get('is_working'):
            await message.answer("❌ Смена уже завершена.", reply_markup=get_admin_main_keyboard())
            await state.clear()
            return
        update_user_data(uid, shift_start_time=new_start.isoformat())
        await message.answer(
            f"✅ Время начала текущей смены изменено на:\n"
            f"<b>{new_start.strftime('%d.%m.%Y %H:%M')}</b>",
            reply_markup=get_admin_main_keyboard(),
            parse_mode="HTML"
        )
        await state.clear()
    except ValueError:
        await message.answer("❌ Неверный формат. Используйте ДД.ММ ЧЧ:ММ", reply_markup=get_cancel_keyboard())
    except Exception as e:
        logging.error(f"Ошибка в input_current_new_start: {e}", exc_info=True)


@dp.callback_query(EditShiftState.current_shift_action, F.data == "csa_finish_shift")
async def finish_current_shift_prompt(call: types.CallbackQuery, state: FSMContext):
    try:
        await call.message.edit_text(
            "Введите время КОНЦА текущей смены (ДД.ММ ЧЧ:ММ):\n"
            "Например: 12.09 17:30",
            reply_markup=get_cancel_keyboard()
        )
        await state.set_state(EditShiftState.editing_current_end)
        await call.answer()
    except Exception as e:
        logging.error(f"Ошибка в finish_current_shift_prompt: {e}", exc_info=True)


@dp.message(EditShiftState.editing_current_end)
async def input_current_end_time(message: types.Message, state: FSMContext):
    try:
        y = datetime.now().year
        t = message.text.strip()
        if len(t.split()[0].split('.')) == 2:
            end_dt = datetime.strptime(f"{t} {y}", "%d.%m %H:%M %Y")
        else:
            end_dt = datetime.strptime(t, "%d.%m.%y %H:%M")
        data = await state.get_data()
        uid = data.get('edit_user_id')
        user, result_msg = finish_shift_at_time(uid, end_dt)
        if user is None:
            await message.answer(f"❌ {result_msg}", reply_markup=get_admin_main_keyboard())
        else:
            await message.answer(
                f"{result_msg}\n\n✅ Смена завершена с указанным временем.",
                reply_markup=get_admin_main_keyboard()
            )
        await state.clear()
    except ValueError:
        await message.answer("❌ Неверный формат. Используйте ДД.ММ ЧЧ:ММ", reply_markup=get_cancel_keyboard())
    except Exception as e:
        logging.error(f"Ошибка в input_current_end_time: {e}", exc_info=True)


@dp.callback_query(EditShiftState.selecting_shift, F.data.startswith("eshift::"))
async def select_shift_to_edit(call: types.CallbackQuery, state: FSMContext):
    try:
        sid = int(parse_callback_value(call.data, "eshift"))
        await state.update_data(shift_id=sid)
        await call.message.edit_text("Новое время НАЧАЛА (ДД.ММ ЧЧ:ММ):", reply_markup=get_cancel_keyboard())
        await state.set_state(EditShiftState.input_start)
        await call.answer()
    except ValueError:
        await call.answer("Ошибка", show_alert=True)
    except Exception as e:
        logging.error(f"Ошибка в select_shift_to_edit: {e}", exc_info=True)


@dp.message(EditShiftState.input_start)
async def input_new_start(message: types.Message, state: FSMContext):
    try:
        y = datetime.now().year
        t = message.text.strip()
        dt = datetime.strptime(f"{t} {y}", "%d.%m %H:%M %Y") if len(
            t.split()[0].split('.')) == 2 else datetime.strptime(t, "%d.%m.%y %H:%M")
        await state.update_data(new_start=dt.isoformat())
        await message.answer("Новое время КОНЦА (ДД.ММ ЧЧ:ММ):", reply_markup=get_cancel_keyboard())
        await state.set_state(EditShiftState.input_end)
    except ValueError:
        await message.answer("❌ Формат: ДД.ММ ЧЧ:ММ", reply_markup=get_cancel_keyboard())
    except Exception as e:
        logging.error(f"Ошибка в input_new_start: {e}", exc_info=True)


@dp.message(EditShiftState.input_end)
async def input_new_end(message: types.Message, state: FSMContext):
    try:
        y = datetime.now().year
        t = message.text.strip()
        dt = datetime.strptime(f"{t} {y}", "%d.%m %H:%M %Y") if len(
            t.split()[0].split('.')) == 2 else datetime.strptime(t, "%d.%m.%y %H:%M")
        data = await state.get_data()
        if update_shift_time(data.get('shift_id'), data.get('new_start'), dt.isoformat()):
            await message.answer("✅ Исправлено!", reply_markup=get_admin_main_keyboard())
        else:
            await message.answer("❌ Ошибка.", reply_markup=get_admin_main_keyboard())
        await state.clear()
    except ValueError:
        await message.answer("❌ Формат: ДД.ММ ЧЧ:ММ", reply_markup=get_cancel_keyboard())
    except Exception as e:
        logging.error(f"Ошибка в input_new_end: {e}", exc_info=True)


# --- ИЗМЕНИТЬ ИМЯ ---
@dp.message(F.text == "📝 Изменить имя")
async def start_edit_name(message: types.Message, state: FSMContext):
    try:
        if not is_admin(message.from_user.id): return
        await message.answer("Выберите отдел:", reply_markup=get_departments_keyboard("ename_dept"))
        await state.set_state(EditNameState.choosing_dept)
    except Exception as e:
        logging.error(f"Ошибка в start_edit_name: {e}", exc_info=True)


@dp.callback_query(EditNameState.choosing_dept, F.data.startswith("ename_dept::"))
async def choose_dept_for_name_edit(call: types.CallbackQuery, state: FSMContext):
    try:
        dept = parse_callback_value(call.data, "ename_dept")
        await state.update_data(edit_dept=dept)
        users = get_users_by_filters(department=dept)
        if not users:
            await call.message.edit_text("Нет сотрудников.")
            await state.clear()
            await call.message.answer("Меню:", reply_markup=get_admin_main_keyboard())
            await call.answer()
            return
        await call.message.edit_text(f"Отдел <b>{dept}</b>. Выберите:",
                                     reply_markup=get_employees_keyboard(users, "ename_user"), parse_mode="HTML")
        await state.set_state(EditNameState.selecting_user)
        await call.answer()
    except Exception as e:
        logging.error(f"Ошибка в choose_dept_for_name_edit: {e}", exc_info=True)


@dp.callback_query(EditNameState.selecting_user, F.data.startswith("ename_user::"))
async def select_user_for_name_edit(call: types.CallbackQuery, state: FSMContext):
    try:
        uid = int(parse_callback_value(call.data, "ename_user"))
        user = get_user_by_id(uid)
        await state.update_data(edit_user_id=uid)
        await call.message.edit_text(f"Имя: <b>{user.get('full_name') or 'Нет'}</b>\nНовое ФИО:",
                                     reply_markup=get_cancel_keyboard(), parse_mode="HTML")
        await state.set_state(EditNameState.input_name)
        await call.answer()
    except ValueError:
        await call.answer("Ошибка", show_alert=True)
    except Exception as e:
        logging.error(f"Ошибка в select_user_for_name_edit: {e}", exc_info=True)


@dp.message(EditNameState.input_name)
async def input_new_name(message: types.Message, state: FSMContext):
    try:
        data = await state.get_data()
        nn = message.text.strip()
        if not nn:
            await message.answer("Пустое.", reply_markup=get_cancel_keyboard())
            return
        update_user_data(data.get('edit_user_id'), full_name=nn)
        await message.answer(f"✅ Имя: <b>{nn}</b>", reply_markup=get_admin_main_keyboard(), parse_mode="HTML")
        await state.clear()
    except Exception as e:
        logging.error(f"Ошибка в input_new_name: {e}", exc_info=True)


# --- НАСТРОЙКИ ---
@dp.message(F.text == "⚙️ Настройки")
async def open_settings(message: types.Message, state: FSMContext):
    try:
        if not is_admin(message.from_user.id): return
        await message.answer("Настроить:", reply_markup=get_settings_menu())
        await state.set_state(SettingsState.choosing_action)
    except Exception as e:
        logging.error(f"Ошибка в open_settings: {e}", exc_info=True)


@dp.callback_query(SettingsState.choosing_action, F.data.startswith("action_"))
async def process_action(call: types.CallbackQuery, state: FSMContext):
    try:
        action = parse_callback_value(call.data, "action")
        await state.update_data(action=action)
        await state.set_state(SettingsState.choosing_target)
        await call.message.edit_text("Кому?", reply_markup=get_target_menu())
        await call.answer()
    except Exception as e:
        logging.error(f"Ошибка в process_action: {e}", exc_info=True)


@dp.callback_query(SettingsState.choosing_target, F.data.startswith("target_"))
async def process_target(call: types.CallbackQuery, state: FSMContext):
    try:
        target = parse_callback_value(call.data, "target")
        data = await state.get_data()
        action = data.get('action')
        await state.update_data(target=target)
        kc = get_cancel_keyboard()
        if target == 'all':
            if action == "schedule":
                await call.message.edit_text("Начало (09:00):", reply_markup=kc)
                await state.set_state(SettingsState.input_start_time)
            elif action == "location":
                await call.message.edit_text("Геолокация:", reply_markup=kc)
                await state.set_state(SettingsState.choosing_target_loc)
            elif action == "salary":
                await call.message.edit_text("⚠️ Массовая установка не поддерживается.", reply_markup=kc)
            elif action == "work_days":
                await call.message.edit_text("Рабочие дни для ВСЕХ:",
                                             reply_markup=get_work_days_keyboard([1, 2, 3, 4, 5]))
                await state.update_data(temp_work_days=[1, 2, 3, 4, 5])
                await state.set_state(SettingsState.selecting_work_days)
        elif target == 'dept':
            await call.message.edit_text("Отдел:", reply_markup=get_departments_keyboard("set_dept"))
            await state.set_state(SettingsState.choosing_dept)
        elif target == 'user':
            await call.message.edit_text("Сотрудник:",
                                         reply_markup=get_employees_keyboard(get_all_employees(), "set_user"))
            await state.set_state(SettingsState.choosing_user)
        await call.answer()
    except Exception as e:
        logging.error(f"Ошибка в process_target: {e}", exc_info=True)


@dp.callback_query(SettingsState.choosing_dept, F.data.startswith("set_dept::"))
async def process_dept_selection(call: types.CallbackQuery, state: FSMContext):
    try:
        dept = parse_callback_value(call.data, "set_dept")
        data = await state.get_data()
        action = data.get('action')
        await state.update_data(identifier=dept)
        kc = get_cancel_keyboard()
        if action == "schedule":
            await call.message.edit_text(f"<b>{dept}</b> Начало (09:00):", reply_markup=kc, parse_mode="HTML")
            await state.set_state(SettingsState.input_start_time)
        elif action == "location":
            await call.message.edit_text(f"<b>{dept}</b> Геолокация:", reply_markup=kc, parse_mode="HTML")
            await state.set_state(SettingsState.choosing_target_loc)
        elif action == "salary":
            await call.message.edit_text(f"<b>{dept}</b> Оклад (₸):", reply_markup=kc, parse_mode="HTML")
            await state.set_state(SettingsState.input_salary)
        elif action == "work_days":
            await call.message.edit_text(f"<b>{dept}</b> Рабочие дни:",
                                         reply_markup=get_work_days_keyboard([1, 2, 3, 4, 5]), parse_mode="HTML")
            await state.update_data(temp_work_days=[1, 2, 3, 4, 5])
            await state.set_state(SettingsState.selecting_work_days)
        await call.answer()
    except Exception as e:
        logging.error(f"Ошибка в process_dept_selection: {e}", exc_info=True)


@dp.callback_query(SettingsState.choosing_user, F.data.startswith("set_user::"))
async def process_user_selection(call: types.CallbackQuery, state: FSMContext):
    try:
        uid = int(parse_callback_value(call.data, "set_user"))
        user = get_user_by_id(uid)
        if not user:
            await call.message.edit_text("Не найден.")
            await state.clear()
            await call.message.answer("Меню:", reply_markup=get_admin_main_keyboard())
            await call.answer()
            return
        data = await state.get_data()
        action = data.get('action')
        un = user.get('full_name') or user.get('username') or "?"
        await state.update_data(identifier=user.get('username') or user.get('phone_number') or str(uid),
                                selected_user_id=uid)
        kc = get_cancel_keyboard()
        if action == "schedule":
            await call.message.edit_text(f"<b>{un}</b> Начало (09:00):", reply_markup=kc, parse_mode="HTML")
            await state.set_state(SettingsState.input_start_time)
        elif action == "location":
            await call.message.edit_text(f"<b>{un}</b> Геолокация:", reply_markup=kc, parse_mode="HTML")
            await state.set_state(SettingsState.choosing_target_loc)
        elif action == "salary":
            await call.message.edit_text(f"<b>{un}</b> Оклад (₸):", reply_markup=kc, parse_mode="HTML")
            await state.set_state(SettingsState.input_salary)
        elif action == "work_days":
            cw = user.get('work_days_week') or '1,2,3,4,5'
            try:
                sel = [int(x) for x in cw.split(',')]
            except:
                sel = [1, 2, 3, 4, 5]
            await call.message.edit_text(f"<b>{un}</b> Рабочие дни:", reply_markup=get_work_days_keyboard(sel),
                                         parse_mode="HTML")
            await state.update_data(temp_work_days=sel)
            await state.set_state(SettingsState.selecting_work_days)
        await call.answer()
    except ValueError:
        await call.answer("Ошибка", show_alert=True)
    except Exception as e:
        logging.error(f"Ошибка в process_user_selection: {e}", exc_info=True)


@dp.callback_query(SettingsState.selecting_work_days, F.data.startswith("wd_toggle_"))
async def toggle_wd(call: types.CallbackQuery, state: FSMContext):
    try:
        day = int(parse_callback_value(call.data, "wd_toggle"))
        data = await state.get_data()
        sel = data.get('temp_work_days', [1, 2, 3, 4, 5])
        if day in sel:
            sel.remove(day)
        else:
            sel.append(day)
            sel.sort()
        await state.update_data(temp_work_days=sel)
        await call.message.edit_reply_markup(reply_markup=get_work_days_keyboard(sel))
        await call.answer()
    except Exception as e:
        logging.error(f"Ошибка в toggle_wd: {e}", exc_info=True)


@dp.callback_query(SettingsState.selecting_work_days, F.data == "wd_save")
async def save_wd(call: types.CallbackQuery, state: FSMContext):
    try:
        data = await state.get_data()
        wds = ",".join(map(str, data.get('temp_work_days', [1, 2, 3, 4, 5])))
        target = data.get('target', 'all')
        ident = data.get('identifier', '')
        suid = data.get('selected_user_id')
        if target == 'all':
            for u in get_all_employees():
                update_user_data(u['user_id'], work_days_week=wds)
            await call.message.edit_text(f"✅ Установлено ВСЕМ: {format_work_days(wds)}")
        elif target == 'dept':
            du = get_users_by_filters(department=ident)
            if du:
                for u in du:
                    update_user_data(u['user_id'], work_days_week=wds)
                await call.message.edit_text(f"✅ Отделу <b>{ident}</b>: {format_work_days(wds)}", parse_mode="HTML")
            else:
                await call.message.edit_text(f"❌ Отдел {ident} не найден.")
        elif target == 'user' and suid:
            u = get_user_by_id(suid)
            if u:
                update_user_data(u['user_id'], work_days_week=wds)
                await call.message.edit_text(f"✅ <b>{u.get('full_name')}</b>: {format_work_days(wds)}",
                                             parse_mode="HTML")
            else:
                await call.message.edit_text(" Не найден.")
        await state.clear()
        await call.message.answer("Меню:", reply_markup=get_admin_main_keyboard())
        await call.answer()
    except Exception as e:
        logging.error(f"Ошибка в save_wd: {e}", exc_info=True)


@dp.message(SettingsState.input_start_time)
async def input_start(message: types.Message, state: FSMContext):
    try:
        datetime.strptime(message.text.strip(), "%H:%M")
    except ValueError:
        await message.answer("Формат 09:00", reply_markup=get_cancel_keyboard())
        return
    await state.update_data(start_time=message.text.strip())
    await message.answer("Конец (18:00):", reply_markup=get_cancel_keyboard())
    await state.set_state(SettingsState.input_end_time)


@dp.message(SettingsState.input_end_time)
async def input_end(message: types.Message, state: FSMContext):
    try:
        datetime.strptime(message.text.strip(), "%H:%M")
    except ValueError:
        await message.answer("Формат 18:00", reply_markup=get_cancel_keyboard())
        return
    data = await state.get_data()
    users, msg = await _resolve_target_users(data)
    if users is None:
        await message.answer("Не найден.", reply_markup=get_admin_main_keyboard())
        await state.clear()
        return
    for u in users:
        update_user_data(u['user_id'], schedule_start=data['start_time'], schedule_end=message.text.strip())
    await message.answer(f"✅ {data['start_time']}-{message.text.strip()} для {msg}.",
                         reply_markup=get_admin_main_keyboard())
    await state.clear()


async def _resolve_target_users(data: dict):
    if data.get('target') == 'all':
        return get_all_employees(), "ВСЕМ"
    suid = data.get('selected_user_id')
    if suid:
        u = get_user_by_id(suid)
        if u:
            return [u], u.get('full_name') or u.get('username')
    ident = data.get('identifier', '')
    digits = ''.join(filter(str.isdigit, ident))
    user = get_user_by_phone(digits) if len(digits) >= 9 or ident.startswith('+') else get_user_by_username(
        ident.lstrip('@'))
    if user:
        return [user], user.get('full_name') or user.get('username')
    du = get_users_by_filters(department=ident)
    if du:
        return du, f"отделу {ident}"
    return None, None


@dp.message(SettingsState.choosing_target_loc, F.content_type == "location")
async def input_loc(message: types.Message, state: FSMContext):
    try:
        await state.update_data(lat=message.location.latitude, lon=message.location.longitude)
        await message.answer("Радиус (м):", reply_markup=get_cancel_keyboard())
        await state.set_state(SettingsState.input_radius)
    except Exception as e:
        logging.error(f"Ошибка в input_loc: {e}", exc_info=True)


@dp.message(SettingsState.choosing_target_loc)
async def input_loc_fb(message: types.Message, state: FSMContext):
    try:
        await message.answer("Отправьте геолокацию (📎 → Геопозиция).", reply_markup=get_cancel_keyboard())
    except Exception as e:
        logging.error(f"Ошибка в input_loc_fb: {e}", exc_info=True)


@dp.message(SettingsState.input_radius)
async def input_rad(message: types.Message, state: FSMContext):
    try:
        r = int(message.text.strip())
        if r <= 0:
            raise ValueError
    except ValueError:
        await message.answer("Положительное число.", reply_markup=get_cancel_keyboard())
        return
    data = await state.get_data()
    if 'lat' not in data or 'lon' not in data:
        await message.answer("❌ Нет геолокации.", reply_markup=get_admin_main_keyboard())
        await state.clear()
        return
    users, msg = await _resolve_target_users(data)
    if users is None:
        await message.answer("Не найден.", reply_markup=get_admin_main_keyboard())
        await state.clear()
        return
    for u in users:
        update_user_data(u['user_id'], target_lat=data['lat'], target_lon=data['lon'], radius=r)
    await message.answer(f"✅ Зона для {msg}.", reply_markup=get_admin_main_keyboard())
    await state.clear()


@dp.message(SettingsState.input_salary)
async def input_sal(message: types.Message, state: FSMContext):
    try:
        s = float(message.text.strip().replace(',', '.'))
        if s < 0:
            raise ValueError
    except ValueError:
        await message.answer("Введите число.", reply_markup=get_cancel_keyboard())
        return
    data = await state.get_data()
    suid = data.get('selected_user_id')
    if suid:
        u = get_user_by_id(suid)
        if u:
            update_user_data(u['user_id'], monthly_salary=s)
            n = get_working_days_in_month(datetime.now().year, datetime.now().month,
                                          u.get('work_days_week') or '1,2,3,4,5')
            await message.answer(
                f"✅ {format_money(s)} для <b>{u.get('full_name') or u.get('username')}</b>.\nСтавка: {format_money(s / n)}",
                reply_markup=get_admin_main_keyboard(), parse_mode="HTML")
            await state.clear()
            return
    ident = data.get('identifier', '')
    digits = ''.join(filter(str.isdigit, ident))
    user = get_user_by_phone(digits) if len(digits) >= 9 or ident.startswith('+') else get_user_by_username(
        ident.lstrip('@'))
    if user:
        update_user_data(user['user_id'], monthly_salary=s)
        n = get_working_days_in_month(datetime.now().year, datetime.now().month,
                                      user.get('work_days_week') or '1,2,3,4,5')
        await message.answer(
            f"✅ {format_money(s)} для <b>{user.get('full_name') or user.get('username')}</b>.\nСтавка: {format_money(s / n)}",
            reply_markup=get_admin_main_keyboard(), parse_mode="HTML")
    else:
        du = get_users_by_filters(department=ident)
        if du:
            for u in du:
                update_user_data(u['user_id'], monthly_salary=s)
            await message.answer(f"✅ {format_money(s)} для отдела <b>{ident}</b>.",
                                 reply_markup=get_admin_main_keyboard(), parse_mode="HTML")
        else:
            await message.answer("Не найден.", reply_markup=get_admin_main_keyboard())
    await state.clear()


# --- РАССЫЛКА ---
@dp.message(F.text == " Рассылка")
async def ask_news(message: types.Message, state: FSMContext):
    try:
        if not is_admin(message.from_user.id): return
        await message.answer("Текст:", reply_markup=get_cancel_keyboard())
        await state.set_state(BroadcastState.waiting_text)
    except Exception as e:
        logging.error(f"Ошибка в ask_news: {e}", exc_info=True)


@dp.message(BroadcastState.waiting_text)
async def process_broadcast(message: types.Message, state: FSMContext):
    try:
        if not is_admin(message.from_user.id):
            await state.clear()
            return
        text = message.text or ""
        if not text.strip():
            await message.answer("Пусто.", reply_markup=get_cancel_keyboard())
            return
        conn = sqlite3.connect('clockster.db')
        ids = [u[0] for u in conn.cursor().execute(
            "SELECT user_id FROM users WHERE role IN ('employee', 'admin') AND user_id > 0").fetchall()]
        conn.close()
        c = 0
        for uid in ids:
            try:
                await bot.send_message(uid, f"📢 {text}")
                c += 1
            except Exception as e:
                logging.warning(f"Не отправлено {uid}: {e}")
        await message.answer(f"✅ {c}/{len(ids)}.", reply_markup=get_admin_main_keyboard())
        await state.clear()
    except Exception as e:
        logging.error(f"Ошибка в process_broadcast: {e}", exc_info=True)


# --- ЛОКАЦИЯ С ЗАЩИТОЙ ОТ ФЕЙКОВОГО GPS ---
@dp.message(F.content_type == "location")
async def handle_employee_location(message: types.Message):
    try:
        user = get_user_by_id(message.from_user.id)
        if not user:
            return

        # === ПРОВЕРКА ФЕЙКОВОГО GPS ===
        is_valid, reason, distance = validate_gps_location(message.location, user)

        if not is_valid:
            log_suspicious_gps(
                user_id=user['user_id'],
                latitude=message.location.latitude,
                longitude=message.location.longitude,
                accuracy=getattr(message.location, 'horizontal_accuracy', None),
                speed=getattr(message.location, 'speed', None),
                distance=distance,
                reason=reason,
                action_taken="blocked"
            )

            admin_msg = (
                f"🚨 <b>ПОДОЗРИТЕЛЬНАЯ ПОПЫТКА GPS</b>\n"
                f" {user.get('full_name') or user.get('username')}\n"
                f" Координаты: {message.location.latitude}, {message.location.longitude}\n"
                f"❌ Причина: {reason}\n"
                f"📏 Расстояние от офиса: {int(distance)}м"
            )
            try:
                await bot.send_message(ADMIN_ID, admin_msg, parse_mode="HTML")
            except Exception:
                pass

            await message.answer(f"❌ Ошибка геолокации: {reason}\nСмена не начата.")
            return

        # === ОБЫЧНАЯ ЛОГИКА НАЧАЛА СМЕНЫ ===
        if user.get('is_working'):
            si = user.get('shift_start_time')
            if si:
                try:
                    sd = datetime.fromisoformat(si)
                    now = datetime.now()
                    if sd.date() == now.date():
                        await message.answer("Вы уже на смене.")
                        return
                    ese = user.get('schedule_end') or "18:00"
                    try:
                        s_end_time = datetime.strptime(ese, "%H:%M").time()
                        ed = datetime.combine(sd.date(), s_end_time)
                    except ValueError:
                        ed = datetime.combine(sd.date() + timedelta(days=1), time(0, 0))
                    dm = int((ed - sd).total_seconds() / 60)
                    ess = user.get('schedule_start') or "09:00"
                    is_late, late_minutes = calculate_late(sd, ess)
                    om = 0
                    create_shift_record(user['user_id'], si, ed.isoformat(), dm, om, is_late, late_minutes)
                    update_user_data(user['user_id'], is_working=0, shift_start_time=None)
                    wds = user.get('work_days_week') or '1,2,3,4,5'
                    ww = not is_working_day(sd, wds)
                    wi = f" (🌟 {DAYS_MAP[sd.weekday() + 1]})" if ww else ""
                    late_info = f" Опоздание: {format_duration(late_minutes)}." if is_late else ""
                    await message.answer(
                        f"⚠️ Незавершённая смена от {sd.strftime('%d.%m %H:%M')} автозавершена в {ed.strftime('%H:%M')} (по графику).{late_info}{wi}")
                except Exception as e:
                    logging.error(f"Ошибка проверки активной смены: {e}")
                    await message.answer("Вы уже на смене.")
                    return

        update_user_data(
            user['user_id'],
            last_location_lat=message.location.latitude,
            last_location_lon=message.location.longitude,
            last_location_time=datetime.now().isoformat()
        )

        wds = user.get('work_days_week') or '1,2,3,4,5'
        td = is_working_day(datetime.now(), wds)
        update_user_data(user['user_id'], is_working=1, shift_start_time=datetime.now().isoformat())

        distance_info = f" (расстояние от офиса: {int(distance)}м)"
        if not td:
            await message.answer(f"🌟 Смена в выходной ({DAYS_MAP[datetime.now().weekday() + 1]})!{distance_info}")
        else:
            await message.answer(f"✅ Смена началась!{distance_info}")
        kb = get_admin_main_keyboard() if is_admin(message.from_user.id) else get_employee_keyboard()
        await message.answer("Меню:", reply_markup=kb)
    except Exception as e:
        logging.error(f"Ошибка в handle_employee_location: {e}", exc_info=True)
        await message.answer("❌ Произошла ошибка при обработке геолокации.")


@dp.message(F.text == "📍 Начать смену")
async def prompt_start(message: types.Message):
    try:
        await message.answer("📍 Отправьте геолокацию (📎 → Геопозиция).")
    except Exception as e:
        logging.error(f"Ошибка в prompt_start: {e}", exc_info=True)


@dp.message(F.text == "🛑 Завершить смену")
async def cmd_stop(message: types.Message):
    try:
        u, m = finish_shift_for_user(message.from_user.id)
        kb = get_admin_main_keyboard() if is_admin(message.from_user.id) else get_employee_keyboard()
        await message.answer(m, reply_markup=kb)
    except Exception as e:
        logging.error(f"Ошибка в cmd_stop: {e}", exc_info=True)
        await message.answer("❌ Произошла ошибка при завершении смены.")


# Добавляем импорт для веб-сервера в начало файла (если его там нет):
# from aiohttp import web

async def start_web_server():
    """Фальшивый веб-сервер, чтобы Render не выключал бесплатный тариф"""
    app = web.Application()

    async def handle(request):
        return web.Response(text="Bot is running")

    app.router.add_get('/', handle)

    runner = web.AppRunner(app)
    await runner.setup()
    # Render требует слушать порт из переменной окружения PORT
    port = int(os.environ.get('PORT', 8080))
    site = web.TCPSite(runner, '0.0.0.0', port)
    await site.start()
    print(f"✅ Веб-сервер запущен на порту {port} для обхода ограничений Render")


async def main():
    init_db()
    migrate_late_minutes()
    deleted = remove_duplicate_shifts()
    if deleted > 0:
        print(f"🗑️ При запуске удалено {deleted} дубликатов смен")

    print(f"Бот запущен (Версия 23.0)...")

    # Запускаем бота и веб-сервер ОДНОВРЕМЕННО
    await asyncio.gather(
        dp.start_polling(bot),
        start_web_server()
    )


if __name__ == "__main__":
    asyncio.run(main())