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
        print(f"📡 Проверяю дату: {target_date.strftime('%d.%m.%Y')}")

    try:
        mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
        mail.login(email_user, password)
        mail.select(FOLDER)
        
        status, data = mail.search(None, *search_criteria)
        if status != "OK": 
            mail.logout()
            return 0
            
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
        return 0


def get_week_stats(user: str, pwd: str, end_date: date) -> Dict[date, int]:
    stats = {}
    print("\n📊 Сбор данных за неделю...")
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
        height =