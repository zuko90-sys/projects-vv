import imaplib
import smtplib
import ssl
import sys
import os
import csv
import re
import mimetypes
import email
from email.header import decode_header
from email.utils import parsedate_to_datetime
from email.message import EmailMessage
from datetime import datetime, timedelta, date
from zoneinfo import ZoneInfo
from typing import Tuple, List, Dict

# --- НОВОЕ: Библиотека для Google Таблиц ---
import gspread

# Настройка для работы графиков в фоне
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

# --- НАСТРОЙКИ ---
IMAP_HOST = "imap.yandex.com"
IMAP_PORT = 993
SMTP_HOST = "smtp.yandex.com"
SMTP_PORT = 465

TZ = ZoneInfo("Europe/Moscow")

FOLDER = "INBOX"
CATEGORY_LABEL = "10 тип"
FROM_EMAIL = "no-reply@vkusvill.ru"
SUBJECT_KEYS = ["10_1", "10_2"]

# ВТОРОЙ ПОЛУЧАТЕЛЬ
REPORT_TO = "hline493@vkusvill.ru"

IMAP_MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", 
               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_DIR = os.path.join(BASE_DIR, "logs")
STATE_DIR = os.path.join(BASE_DIR, "state")
HISTORY_DIR = os.path.join(BASE_DIR, "history")
HISTORY_FILE = os.path.join(HISTORY_DIR, "mail_history.csv")
CHART_FILE = os.path.join(BASE_DIR, "chart_weekly.png")

# --- НАСТРОЙКИ GOOGLE ТАБЛИЦ ---
GOOGLE_KEY_FILE = os.path.join(BASE_DIR, "google_key.json")
SHEET_NAME = "Статистика по 10 типу"


def get_credentials() -> Tuple[str, str]:
    u = os.getenv("EMAIL_USER")
    p = os.getenv("EMAIL_PASSWORD")
    if u and p: return u, p
    try:
        from dotenv import load_dotenv
        load_dotenv(os.path.join(BASE_DIR, ".env"))
        u = os.getenv("EMAIL_USER")
        p = os.getenv("EMAIL_PASSWORD")
        if u and p: return u, p
    except ImportError:
        pass
    cred_path = os.path.join(BASE_DIR, "credentials.txt")
    if os.path.exists(cred_path):
        with open(cred_path, "r", encoding="utf-8") as f:
            lines = [l.strip() for l in f if l.strip()]
            if len(lines) >= 2: return lines[0], lines[1]
    raise RuntimeError("Не найдены EMAIL_USER/PASSWORD")


def ensure_dirs():
    for d in [LOG_DIR, STATE_DIR, HISTORY_DIR]:
        os.makedirs(d, exist_ok=True)


def write_log(msg: str):
    ensure_dirs()
    now = datetime.now(TZ)
    log_file = os.path.join(LOG_DIR, f"log_{now.strftime('%Y-%m-%d')}.txt")
    try:
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(f"[{now.strftime('%H:%M:%S')}] {msg}\n")
    except:
        pass


def imap_date_str(d: date) -> str:
    return f"{d.day}-{IMAP_MONTHS[d.month - 1]}-{d.year}"


def decode_subject(header_value) -> str:
    if not header_value: return ""
    try:
        decoded_list = decode_header(header_value)
        parts = []
        for content, encoding in decoded_list:
            if isinstance(content, bytes):
                parts.append(content.decode(encoding or 'utf-8', errors='ignore'))
            else:
                parts.append(str(content))
        return "".join(parts)
    except Exception:
        return str(header_value)


def count_emails_robust(email_user: str, password: str, target_date: date, silent: bool = False) -> int:
    imap_start = target_date - timedelta(days=1)
    imap_end = target_date + timedelta(days=2)
    
    search_criteria = ["SINCE", imap_date_str(imap_start), "BEFORE", imap_date_str(imap_end), "FROM", FROM_EMAIL]
    
    count = 0
    if not silent:
        print(f"[IMAP] Проверяю дату: {target_date.strftime('%d.%m.%Y')}")

    try:
        mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
        mail.login(email_user, password)
        mail.select(FOLDER)
        
        status, data = mail.search(None, *search_criteria)
        if status != "OK": 
            mail.logout()
            return -1
            
        ids = data[0].split()
        if not ids:
            mail.logout()
            return 0
            
        id_str = b",".join(ids).decode('utf-8')
        typ, msg_data = mail.fetch(id_str, '(BODY.PEEK[HEADER.FIELDS (SUBJECT DATE)])')
        
        for response_part in msg_data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                decoded_subj = decode_subject(msg.get("Subject"))
                if not any(key in decoded_subj for key in SUBJECT_KEYS):
                    continue

                raw_date = msg.get("Date")
                try:
                    dt_obj = parsedate_to_datetime(raw_date)
                    dt_msk = dt_obj.astimezone(TZ)
                    if dt_msk.date() == target_date:
                        count += 1
                except:
                    continue
        mail.logout()
        return count
    except Exception as e:
        if not silent: write_log(f"Ошибка IMAP: {e}")
        return -1


def get_week_stats(user: str, pwd: str, end_date: date) -> Dict[date, int]:
    stats = {}
    print("\n[CHART] Сбор данных за неделю...")
    for i in range(6, -1, -1):
        day = end_date - timedelta(days=i)
        count = count_emails_robust(user, pwd, day, silent=True)
        stats[day] = count
    return stats


def create_chart(stats: Dict[date, int], filename: str):
    dates = [d.strftime("%d.%m") for d in stats.keys()]
    counts = list(stats.values())

    plt.figure(figsize=(10, 6))
    bars = plt.bar(dates, counts, color='#96C11F')
    plt.title(f"Количество писем '{CATEGORY_LABEL}' за неделю", fontsize=14)
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    
    for bar in bars:
        height = bar.get_height()
        plt.text(bar.get_x() + bar.get_width()/2., height, f'{height}', ha='center', va='bottom')

    plt.tight_layout()
    plt.savefig(filename)
    plt.close()


def send_report(user: str, pwd: str, subject: str, body: str, attachment_path: str = None):
    msg = EmailMessage()
    msg["From"] = user
    
    recipients = [user]
    if REPORT_TO:
        recipients.append(REPORT_TO)
    
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = subject
    msg.set_content(body)

    if attachment_path and os.path.exists(attachment_path):
        ctype, encoding = mimetypes.guess_type(attachment_path)
        if ctype is None or encoding is not None:
            ctype = 'application/octet-stream'
        maintype, subtype = ctype.split('/', 1)

        with open(attachment_path, 'rb') as f:
            file_data = f.read()
            msg.add_attachment(file_data, maintype=maintype, subtype=subtype, filename=os.path.basename(attachment_path))

    ctx = ssl.create_default_context()
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=ctx) as s:
        s.login(user, pwd)
        s.send_message(msg)


# --- НОВАЯ ФУНКЦИЯ ДЛЯ GOOGLE ТАБЛИЦ ---
def append_to_gsheet(report_date: date, mode: str, count: int):
    if not os.path.exists(GOOGLE_KEY_FILE):
        print("[GSHEET SKIP] Файл google_key.json не найден.")
        return

    try:
        # Авторизуемся по ключу
        gc = gspread.service_account(filename=GOOGLE_KEY_FILE)
        # Открываем таблицу по имени
        sh = gc.open(SHEET_NAME)
        # Выбираем первый лист
        worksheet = sh.sheet1
        
        # Подготавливаем данные для записи
        date_str = report_date.strftime('%d.%m.%Y')
        
        # Записываем строку
        worksheet.append_row([date_str, mode, count])
        print("[GSHEET] Данные успешно записаны в Google Таблицу.")
        write_log(f"GSHEET SUCCESS: {date_str}, {mode}, {count}")
    except Exception as e:
        print(f"[GSHEET ERROR] Ошибка записи в таблицу: {e}")
        write_log(f"GSHEET ERROR: {e}")
# ----------------------------------------


def run(mode: str):
    ensure_dirs()
    now = datetime.now(TZ)
    user, pwd = get_credentials()
    today = now.date()
    
    attachment = None 
    error_msg = None
    sunday_count = 0 

    if mode == "today":
        report_date = today
        date_str = report_date.strftime('%d.%m.%y')
        period_body = f"{date_str} 00:00-23:59 (Сегодня)"
        should_record = False 
        
        count = count_emails_robust(user, pwd, report_date)
        if count == -1:
            error_msg = "Ошибка доступа к данным (сервер Яндекса не ответил)."
            count = 0

    elif mode == "daily":
        report_date = today - timedelta(days=1)
        date_str = report_date.strftime('%d.%m.%y')
        period_body = f"{date_str} 00:00-23:59 (Вчера)"
        should_record = True
        
        count = count_emails_robust(user, pwd, report_date)
        if count == -1:
            error_msg = "Ошибка доступа к данным (сервер Яндекса не ответил)."
            count = 0

    elif mode == "weekly":
        report_date = today - timedelta(days=1) 
        start_week = report_date - timedelta(days=6) 
        date_str = f"{start_week.strftime('%d.%m')} - {report_date.strftime('%d.%m.%y')}"
        period_body = f"Неделя: {date_str}"
        should_record = True
        
        stats = get_week_stats(user, pwd, report_date)
        
        if -1 in stats.values():
            error_msg = "Ошибка доступа к данным при сборе статистики. График может быть неточным."
            stats = {k: (0 if v == -1 else v) for k, v in stats.items()}
            
        create_chart(stats, CHART_FILE)
        attachment = CHART_FILE
        
        count = sum(stats.values()) 
        sunday_count = stats.get(report_date, 0) 

    else:
        return

    state_file = os.path.join(STATE_DIR, f"{mode}_{report_date}.sent")
    if should_record and os.path.exists(state_file):
        print(f"[SKIP] Отчет '{mode}' за {report_date} уже был отправлен.")
        return

    subj = f"Отчет по письмам по 10 типу [{mode}]: {date_str}"
    
    body_lines = [
        f"Категория: {CATEGORY_LABEL}",
        f"Период: {period_body}",
        f"Время формирования: {now.strftime('%H:%M %d.%m.%y')}",
        "-----------------"
    ]
    
    if error_msg:
        body_lines.append(f"ВНИМАНИЕ: {error_msg}")
        
    if mode == "weekly":
        body_lines.append(f"За вчера (Воскресенье {report_date.strftime('%d.%m')}): {sunday_count} шт.")
        body_lines.append(f"Всего за неделю: {count} шт.")
    else:
        if error_msg:
            body_lines.append(f"Удалось подсчитать: {count} шт.")
        else:
            body_lines.append(f"Количество писем: {count}")
            
    body = "\n".join(body_lines)

    try:
        # 1. Отправляем письмо
        send_report(user, pwd, subj, body, attachment_path=attachment)
        write_log(f"SUCCESS: {mode}, date={report_date}, count={count}")
        print(f"[SUCCESS] Отправлено: {user}, {REPORT_TO}")
        
        # 2. Если отправка прошла успешно, сохраняем данные
        if should_record:
            # Создаем файл-метку
            with open(state_file, "w") as f: f.write(now.isoformat())
            # Пишем в локальный CSV
            with open(HISTORY_FILE, "a", newline="", encoding="utf-8") as f:
                csv.writer(f).writerow([report_date, mode, count])
            
            # Пишем в облачную Google Таблицу
            append_to_gsheet(report_date, mode, count)

    except Exception as e:
        write_log(f"ERROR sending mail: {e}")
        print(f"Ошибка отправки: {e}")
        sys.exit(1)

if __name__ == "__main__":
    if len(sys.argv) > 1:
        run(sys.argv[1])
    else:
        now_msk = datetime.now(TZ)
        if now_msk.weekday() == 0:
            print("[AUTO] Сегодня Понедельник -> Запуск WEEKLY")
            run("weekly")
        else:
            print("[AUTO] Обычный день -> Запуск DAILY")
            run("daily")