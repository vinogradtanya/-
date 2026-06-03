"""Работа с MySQL для читательского дневника."""
from __future__ import annotations

import os
import shutil
import uuid
import math
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

import sys

import mysql.connector
from mysql.connector import MySQLConnection
from mysql.connector.abstracts import MySQLConnectionAbstract
from mysql.connector.pooling import PooledMySQLConnection

from config import COVERS_DIR
from dotenv import load_dotenv

# При запуске через PyInstaller exe находится в sys.executable,
# при обычном запуске — рядом с db.py
_BASE_DIR = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
load_dotenv(_BASE_DIR / ".env")


READING_PLANNED = "planned"
READING_READING = "reading"
READING_FINISHED = "finished"
READING_ABANDONED = "abandoned"
READING_STATUS_VALUES = (READING_PLANNED, READING_READING, READING_FINISHED, READING_ABANDONED)


def connect() -> PooledMySQLConnection | MySQLConnectionAbstract:
    """Создает и возвращает подключение к БД."""
    return mysql.connector.connect(
        host=os.getenv("DB_HOST"),
        user=os.getenv("DB_USER"),
        password=os.getenv("DB_PASSWORD"),
        database=os.getenv("DB_NAME"),
        port=int(os.getenv("DB_PORT"))
    )


def ensure_schema() -> None:
    """Добавляет недостающие столбцы и таблицы в существующей БД."""
    try:
        with connect() as conn:
            cur = conn.cursor()
            cur.execute(
                """
                SELECT COUNT(*) FROM information_schema.COLUMNS
                WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'books'
                  AND COLUMN_NAME = 'reading_status'
                """
            )
            if cur.fetchone()[0] == 0:
                cur.execute(
                    """
                    ALTER TABLE books
                    ADD COLUMN reading_status VARCHAR(20) NOT NULL DEFAULT 'planned'
                    """
                )
            # Добавляем колонки page_start и page_end
            cur.execute(
                """
                SELECT COUNT(*) FROM information_schema.COLUMNS
                WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'books'
                  AND COLUMN_NAME = 'page_start'
                """
            )
            if cur.fetchone()[0] == 0:
                cur.execute(
                    """
                    ALTER TABLE books
                    ADD COLUMN page_start INT DEFAULT 0
                    """
                )
            cur.execute(
                """
                SELECT COUNT(*) FROM information_schema.COLUMNS
                WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'books'
                  AND COLUMN_NAME = 'page_end'
                """
            )
            if cur.fetchone()[0] == 0:
                cur.execute(
                    """
                    ALTER TABLE books
                    ADD COLUMN page_end INT DEFAULT 0
                    """
                )
            cur.execute(
                """
                SELECT COUNT(*) FROM information_schema.TABLES
                WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'book_tags'
                """
            )
            if cur.fetchone()[0] == 0:
                cur.execute(
                    """
                    CREATE TABLE book_tags (
                        book_id INT NOT NULL,
                        tag VARCHAR(80) NOT NULL,
                        PRIMARY KEY (book_id, tag),
                        INDEX idx_book_tags_tag (tag)
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
                    """
                )
            cur.execute(
                """
                SELECT COUNT(*) FROM information_schema.TABLES
                WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'reading_goals'
                """
            )
            if cur.fetchone()[0] == 0:
                cur.execute(
                    """
                    CREATE TABLE reading_goals (
                        scope_year INT NOT NULL,
                        scope_month TINYINT NOT NULL DEFAULT 0,
                        target_pages INT NOT NULL DEFAULT 0,
                        target_minutes INT NOT NULL DEFAULT 0,
                        PRIMARY KEY (scope_year, scope_month)
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
                    """
                )
            conn.commit()
    except Exception:
        pass
    ensure_book_goals_table()
    _migrate_abandoned_status()
    _migrate_quotes_page_number()


def _migrate_quotes_page_number() -> None:
    """Миграция: добавляет колонку page_number в таблицу quotes."""
    try:
        with connect() as conn:
            cur = conn.cursor()
            cur.execute(
                """
                SELECT COUNT(*) FROM information_schema.COLUMNS
                WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'quotes'
                  AND COLUMN_NAME = 'page_number'
                """
            )
            if cur.fetchone()[0] == 0:
                cur.execute("ALTER TABLE quotes ADD COLUMN page_number INT NULL")
                conn.commit()
    except Exception:
        pass


def _migrate_abandoned_status() -> None:
    """Миграция: расширяет VARCHAR reading_status до 20 символов (уже достаточно для 'abandoned')."""
    try:
        with connect() as conn:
            cur = conn.cursor()
            # Проверяем что поле достаточной длины (abandoned = 9 символов, planned = 7)
            cur.execute(
                """
                SELECT CHARACTER_MAXIMUM_LENGTH
                FROM information_schema.COLUMNS
                WHERE TABLE_SCHEMA = DATABASE()
                  AND TABLE_NAME = 'books'
                  AND COLUMN_NAME = 'reading_status'
                """
            )
            row = cur.fetchone()
            if row and int(row[0]) < 20:
                cur.execute(
                    "ALTER TABLE books MODIFY COLUMN reading_status VARCHAR(20) NOT NULL DEFAULT 'planned'"
                )
                conn.commit()
    except Exception:
        pass


def _row_book(r: dict) -> dict:
    if not r:
        return r
    out = dict(r)
    for k in ("date_started", "date_finished", "created_at", "updated_at"):
        v = out.get(k)
        if hasattr(v, "isoformat"):
            out[k] = v.isoformat() if v else None
    return out


# Books

def list_books_for_month(year: int, month: int) -> list[dict]:
    with connect() as conn:
        cur = conn.cursor(dictionary=True)
        cur.execute(
            "SELECT * FROM books WHERE plan_year = %s AND plan_month = %s ORDER BY title",
            (year, month),
        )
        return [_row_book(x) for x in cur.fetchall()]


def list_books_for_year(year: int) -> list[dict]:
    with connect() as conn:
        cur = conn.cursor(dictionary=True)
        cur.execute(
            "SELECT * FROM books WHERE plan_year = %s ORDER BY plan_month, title",
            (year,),
        )
        return [_row_book(x) for x in cur.fetchall()]


def list_books_by_status(status: str) -> list[dict]:
    if status not in READING_STATUS_VALUES:
        status = READING_PLANNED
    with connect() as conn:
        cur = conn.cursor(dictionary=True)
        cur.execute(
            "SELECT * FROM books WHERE reading_status = %s ORDER BY updated_at DESC, id DESC",
            (status,),
        )
        return [_row_book(x) for x in cur.fetchall()]


def list_all_books() -> list[dict]:
    with connect() as conn:
        cur = conn.cursor(dictionary=True)
        cur.execute("SELECT * FROM books ORDER BY updated_at DESC, id DESC")
        return [_row_book(x) for x in cur.fetchall()]


def get_book(book_id: int) -> dict | None:
    with connect() as conn:
        cur = conn.cursor(dictionary=True)
        cur.execute("SELECT * FROM books WHERE id = %s", (book_id,))
        return _row_book(cur.fetchone())


def create_book(
    title: str,
    author: str = "",
    genre: str = "",
    page_count: int = 0,
    page_start: int = 0,
    page_end: int = 0,
    format_type: str = "paper",
    plan_year: int | None = None,
    plan_month: int | None = None,
    cover_path: str | None = None,
    date_started: date | None = None,
    date_finished: date | None = None,
    reading_status: str = READING_PLANNED,
) -> int:
    if reading_status not in READING_STATUS_VALUES:
        reading_status = READING_PLANNED
    # Если указаны page_start и page_end, рассчитываем page_count
    if page_start > 0 and page_end > 0 and page_end >= page_start:
        page_count = page_end - page_start + 1
    with connect() as conn:
        cur = conn.cursor()
        cur.execute(
            """INSERT INTO books (title, author, genre, page_count, page_start, page_end, format_type,
               plan_year, plan_month, cover_path, date_started, date_finished, reading_status)
               VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)""",
            (title, author, genre, page_count, page_start, page_end, format_type, plan_year,
             plan_month, cover_path, date_started, date_finished, reading_status),
        )
        conn.commit()
        return cur.lastrowid


def update_book(book_id: int, **fields: Any) -> None:
    allowed = {
        "title", "author", "genre", "page_count", "page_start", "page_end", "format_type",
        "plan_year", "plan_month", "cover_path", "date_started", "date_finished",
        "reading_status",
    }
    sets, vals = [], []
    for k, v in fields.items():
        if k in allowed:
            sets.append(f"{k} = %s")
            vals.append(v)
    if not sets:
        return

    vals.append(book_id)
    sql = f"UPDATE books SET {', '.join(sets)} WHERE id = %s"

    with connect() as conn:
        cur = conn.cursor()
        cur.execute(sql, vals)
        conn.commit()


def delete_book(book_id: int) -> None:
    with connect() as conn:
        cur = conn.cursor()
        cur.execute("DELETE FROM daily_reading WHERE book_id = %s", (book_id,))
        cur.execute("DELETE FROM quotes WHERE book_id = %s", (book_id,))
        cur.execute("DELETE FROM reviews WHERE book_id = %s", (book_id,))
        cur.execute("DELETE FROM book_tags WHERE book_id = %s", (book_id,))
        cur.execute("DELETE FROM books WHERE id = %s", (book_id,))
        conn.commit()


def save_cover_from_path(book_id: int, src_path: str) -> str | None:
    if not src_path:
        return None
    src = Path(src_path)
    if not src.is_file():
        return None

    ext = src.suffix.lower() if src.suffix else ".jpg"
    if ext not in (".jpg", ".jpeg", ".png", ".gif", ".webp"):
        ext = ".jpg"

    dest = COVERS_DIR / f"{book_id}_{uuid.uuid4().hex}{ext}"
    shutil.copy2(src, dest)
    return str(dest.resolve())


# Quotes

def list_quotes(book_id: int) -> list[dict]:
    with connect() as conn:
        cur = conn.cursor(dictionary=True)
        cur.execute("SELECT * FROM quotes WHERE book_id = %s ORDER BY id", (book_id,))
        return cur.fetchall()


def add_quote(book_id: int, text: str, page_number: int | None = None) -> int:
    with connect() as conn:
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO quotes (book_id, quote_text, page_number) VALUES (%s, %s, %s)",
            (book_id, text, page_number)
        )
        conn.commit()
        return cur.lastrowid


def update_quote(quote_id: int, text: str, page_number: int | None = None) -> None:
    with connect() as conn:
        cur = conn.cursor()
        cur.execute(
            "UPDATE quotes SET quote_text = %s, page_number = %s WHERE id = %s",
            (text, page_number, quote_id)
        )
        conn.commit()


def delete_quote(quote_id: int) -> None:
    with connect() as conn:
        cur = conn.cursor()
        cur.execute("DELETE FROM quotes WHERE id = %s", (quote_id,))
        conn.commit()


# Reviews

def get_review(book_id: int) -> dict | None:
    with connect() as conn:
        cur = conn.cursor(dictionary=True)
        cur.execute("SELECT * FROM reviews WHERE book_id = %s", (book_id,))
        return cur.fetchone()


def upsert_review(
    book_id: int,
    rating_idea: int,
    rating_plot: int,
    rating_characters: int,
    rating_author_skill: int,
    review_text: str | None,
) -> None:
    with connect() as conn:
        cur = conn.cursor()
        cur.execute(
            """INSERT INTO reviews (book_id, rating_idea, rating_plot,
               rating_characters, rating_author_skill, review_text)
               VALUES (%s, %s, %s, %s, %s, %s)
               ON DUPLICATE KEY UPDATE
               rating_idea = VALUES(rating_idea),
               rating_plot = VALUES(rating_plot),
               rating_characters = VALUES(rating_characters),
               rating_author_skill = VALUES(rating_author_skill),
               review_text = VALUES(review_text)""",
            (book_id, rating_idea, rating_plot, rating_characters, rating_author_skill, review_text),
        )
        conn.commit()


# Daily reading

def list_daily_for_book(book_id: int) -> list[dict]:
    with connect() as conn:
        cur = conn.cursor(dictionary=True)
        cur.execute(
            "SELECT * FROM daily_reading WHERE book_id = %s ORDER BY read_date DESC",
            (book_id,),
        )
        rows = cur.fetchall()

    for r in rows:
        d = r.get("read_date")
        if hasattr(d, "isoformat"):
            r["read_date"] = d.isoformat()
    return rows


def list_daily_for_date(d: date) -> list[dict]:
    with connect() as conn:
        cur = conn.cursor(dictionary=True)
        cur.execute(
            """SELECT dr.*, b.title AS book_title FROM daily_reading dr
               JOIN books b ON b.id = dr.book_id WHERE dr.read_date = %s
               ORDER BY b.title""",
            (d,),
        )
        rows = cur.fetchall()

    for r in rows:
        x = r.get("read_date")
        if hasattr(x, "isoformat"):
            r["read_date"] = x.isoformat()
    return rows


def upsert_daily(
    book_id: int,
    read_date: date,
    pages_read: int,
    minutes_read: int = 0,
    note: str | None = None,
    page_start: int = 0,
    page_end: int = 0,
) -> None:
    with connect() as conn:
        cur = conn.cursor()
        cur.execute(
            """INSERT INTO daily_reading (book_id, read_date, page_start, page_end, pages_read, minutes_read, note)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
            ON DUPLICATE KEY UPDATE
            page_start = VALUES(page_start),
            page_end = VALUES(page_end),
            pages_read = VALUES(pages_read),
            minutes_read = VALUES(minutes_read),
            note = VALUES(note)""",
            (book_id, read_date, page_start, page_end, pages_read, minutes_read, note),
        )
        conn.commit()


def delete_daily_entry(book_id: int, read_date: date) -> None:
    with connect() as conn:
        cur = conn.cursor()
        cur.execute(
            "DELETE FROM daily_reading WHERE book_id = %s AND read_date = %s",
            (book_id, read_date),
        )
        conn.commit()


# Tags

def _norm_tag(s: str) -> str:
    t = (s or "").strip().lower()
    return t[:80] if t else ""


def replace_book_tags(book_id: int, tags: list[str]) -> None:
    seen: set[str] = set()
    clean: list[str] = []
    for raw in tags:
        t = _norm_tag(raw)
        if t and t not in seen:
            seen.add(t)
            clean.append(t)
    with connect() as conn:
        cur = conn.cursor()
        cur.execute("DELETE FROM book_tags WHERE book_id = %s", (book_id,))
        for t in clean:
            cur.execute(
                "INSERT INTO book_tags (book_id, tag) VALUES (%s, %s)",
                (book_id, t),
            )
        conn.commit()


def get_book_tags(book_id: int) -> list[str]:
    with connect() as conn:
        cur = conn.cursor()
        cur.execute(
            "SELECT tag FROM book_tags WHERE book_id = %s ORDER BY tag",
            (book_id,),
        )
        return [r[0] for r in cur.fetchall()]


def list_distinct_tags() -> list[str]:
    with connect() as conn:
        cur = conn.cursor()
        cur.execute("SELECT DISTINCT tag FROM book_tags ORDER BY tag")
        return [r[0] for r in cur.fetchall()]


def list_books_by_tag(tag: str) -> list[dict]:
    t = _norm_tag(tag)
    if not t:
        return []
    with connect() as conn:
        cur = conn.cursor(dictionary=True)
        cur.execute(
            """SELECT b.* FROM books b
               INNER JOIN book_tags bt ON bt.book_id = b.id AND bt.tag = %s
               ORDER BY b.updated_at DESC, b.id DESC""",
            (t,),
        )
        return [_row_book(x) for x in cur.fetchall()]


# Reading goals (scope_month: 0 = весь год, 1–12 = месяц)

def upsert_reading_goal(
    scope_year: int,
    scope_month: int,
    target_books: int,
) -> None:
    scope_month = max(0, min(12, scope_month))
    target_books = max(0, target_books)
    with connect() as conn:
        cur = conn.cursor()
        cur.execute(
            """INSERT INTO reading_goals (scope_year, scope_month, target_pages, target_minutes)
               VALUES (%s, %s, %s, 0)
               ON DUPLICATE KEY UPDATE
               target_pages = VALUES(target_pages)""",
            (scope_year, scope_month, target_books),
        )
        conn.commit()


def get_reading_goal(scope_year: int, scope_month: int) -> dict | None:
    with connect() as conn:
        cur = conn.cursor(dictionary=True)
        cur.execute(
            "SELECT * FROM reading_goals WHERE scope_year = %s AND scope_month = %s",
            (scope_year, scope_month),
        )
        return cur.fetchone()


def count_finished_books_between(d0: date, d1: date) -> int:
    """Количество книг со статусом finished и датой окончания в диапазоне."""
    with connect() as conn:
        cur = conn.cursor()
        cur.execute(
            """SELECT COUNT(*) FROM books
               WHERE reading_status = 'finished'
               AND date_finished BETWEEN %s AND %s""",
            (d0, d1),
        )
        return int(cur.fetchone()[0] or 0)


# Book reading goal (цель по дням для конкретной книги)

def ensure_book_goals_table() -> None:
    with connect() as conn:
        cur = conn.cursor()
        cur.execute(
            """
            SELECT COUNT(*) FROM information_schema.TABLES
            WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'book_goals'
            """
        )
        if cur.fetchone()[0] == 0:
            cur.execute(
                """
                CREATE TABLE book_goals (
                    book_id INT NOT NULL PRIMARY KEY,
                    target_days INT NOT NULL DEFAULT 0,
                    deadline_date DATE NULL
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
                """
            )
            conn.commit()


def upsert_book_goal(book_id: int, target_days: int, deadline_date: date | None) -> None:
    with connect() as conn:
        cur = conn.cursor()
        cur.execute(
            """INSERT INTO book_goals (book_id, target_days, deadline_date)
               VALUES (%s, %s, %s)
               ON DUPLICATE KEY UPDATE
               target_days = VALUES(target_days),
               deadline_date = VALUES(deadline_date)""",
            (book_id, max(0, target_days), deadline_date),
        )
        conn.commit()


def get_book_goal(book_id: int) -> dict | None:
    with connect() as conn:
        cur = conn.cursor(dictionary=True)
        cur.execute("SELECT * FROM book_goals WHERE book_id = %s", (book_id,))
        row = cur.fetchone()
    if row and row.get("deadline_date") and hasattr(row["deadline_date"], "isoformat"):
        row["deadline_date"] = row["deadline_date"].isoformat()
    return row


def count_reading_days_for_book(book_id: int) -> int:
    """Количество дней, в которые была запись в трекере для книги."""
    with connect() as conn:
        cur = conn.cursor()
        cur.execute(
            "SELECT COUNT(DISTINCT read_date) FROM daily_reading WHERE book_id = %s",
            (book_id,),
        )
        return int(cur.fetchone()[0] or 0)


# Stats

def _iso_date_row(r: dict) -> None:
    d = r.get("read_date")
    if hasattr(d, "isoformat"):
        r["read_date"] = d.isoformat()


def sum_daily_between(d0: date, d1: date) -> tuple[int, int]:
    with connect() as conn:
        cur = conn.cursor()
        cur.execute(
            """SELECT COALESCE(SUM(pages_read), 0), COALESCE(SUM(minutes_read), 0)
               FROM daily_reading WHERE read_date BETWEEN %s AND %s""",
            (d0, d1),
        )
        row = cur.fetchone()
        return int(row[0] or 0), int(row[1] or 0)


def daily_breakdown_between(d0: date, d1: date) -> list[dict]:
    with connect() as conn:
        cur = conn.cursor(dictionary=True)
        cur.execute(
            """SELECT read_date AS read_date, SUM(pages_read) AS pages_read,
                      SUM(minutes_read) AS minutes_read
               FROM daily_reading WHERE read_date BETWEEN %s AND %s
               GROUP BY read_date ORDER BY read_date DESC""",
            (d0, d1),
        )
        rows = cur.fetchall()
    for r in rows:
        _iso_date_row(r)
        r["pages_read"] = int(r.get("pages_read") or 0)
        r["minutes_read"] = int(r.get("minutes_read") or 0)
    return rows


def total_pages_logged_for_book(book_id: int) -> int:
    with connect() as conn:
        cur = conn.cursor()
        cur.execute(
            "SELECT COALESCE(SUM(pages_read), 0) FROM daily_reading WHERE book_id = %s",
            (book_id,),
        )
        return int(cur.fetchone()[0] or 0)


def book_pace_estimate(book_id: int, window_days: int = 14) -> dict | None:
    b = get_book(book_id)
    if not b:
        return None
    total = int(b.get("page_count") or 0)
    logged = total_pages_logged_for_book(book_id)
    remaining = max(0, total - logged) if total > 0 else None
    d_end = date.today()
    d_start = d_end - timedelta(days=window_days)
    with connect() as conn:
        cur = conn.cursor(dictionary=True)
        cur.execute(
            """SELECT read_date, pages_read FROM daily_reading
               WHERE book_id = %s AND read_date BETWEEN %s AND %s AND pages_read > 0
               ORDER BY read_date""",
            (book_id, d_start, d_end),
        )
        rows = cur.fetchall()
    pages_in_window = sum(int(r.get("pages_read") or 0) for r in rows)
    days_with = len({r["read_date"].isoformat() if hasattr(r["read_date"], "isoformat") else str(r["read_date"])[:10] for r in rows})
    avg = pages_in_window / max(1, days_with) if rows else 0.0
    est_days: int | None = None
    if remaining is not None and remaining > 0 and avg > 0:
        est_days = int(math.ceil(remaining / avg))
    return {
        "title": b.get("title") or "",
        "page_count": total,
        "pages_logged": logged,
        "remaining_pages": remaining,
        "avg_pages_per_active_day": round(avg, 2),
        "days_with_reading_in_window": days_with,
        "estimated_days_to_finish": est_days,
        "window_days": window_days,
    }


def _like_contains(q: str) -> str:
    esc = q.replace("|", "||").replace("%", "|%").replace("_", "|_")
    return f"%{esc}%"


def search_books(q: str) -> list[dict]:
    q = (q or "").strip()
    if len(q) < 1:
        return []
    like = _like_contains(q)
    found: dict[int, dict] = {}
    with connect() as conn:
        cur = conn.cursor(dictionary=True)
        cur.execute(
            """SELECT id, title FROM books
               WHERE title LIKE %s ESCAPE '|' OR author LIKE %s ESCAPE '|'
                     OR genre LIKE %s ESCAPE '|'""",
            (like, like, like),
        )
        for row in cur.fetchall():
            bid = int(row["id"])
            found[bid] = {"id": bid, "title": row.get("title") or "", "match": "книга"}

        cur.execute(
            """SELECT DISTINCT b.id, b.title FROM books b
               INNER JOIN quotes q ON q.book_id = b.id
               WHERE q.quote_text LIKE %s ESCAPE '|'""",
            (like,),
        )
        for row in cur.fetchall():
            bid = int(row["id"])
            if bid not in found:
                found[bid] = {"id": bid, "title": row.get("title") or "", "match": "цитата"}
            else:
                found[bid]["match"] = found[bid]["match"] + ", цитата"

        cur.execute(
            """SELECT DISTINCT b.id, b.title FROM books b
               INNER JOIN daily_reading d ON d.book_id = b.id
               WHERE d.note LIKE %s ESCAPE '|'""",
            (like,),
        )
        for row in cur.fetchall():
            bid = int(row["id"])
            if bid not in found:
                found[bid] = {"id": bid, "title": row.get("title") or "", "match": "заметка трекера"}
            else:
                found[bid]["match"] = found[bid]["match"] + ", заметка"

        cur.execute(
            """SELECT DISTINCT b.id, b.title FROM books b
               INNER JOIN book_tags bt ON bt.book_id = b.id
               WHERE bt.tag LIKE %s ESCAPE '|'""",
            (like,),
        )
        for row in cur.fetchall():
            bid = int(row["id"])
            if bid not in found:
                found[bid] = {"id": bid, "title": row.get("title") or "", "match": "тег"}
            else:
                found[bid]["match"] = found[bid]["match"] + ", тег"

    return sorted(found.values(), key=lambda x: (x["title"].lower(), x["id"]))



def list_all_quotes_export() -> list[dict]:
    with connect() as conn:
        cur = conn.cursor(dictionary=True)
        cur.execute("SELECT * FROM quotes ORDER BY id")
        return cur.fetchall()


def list_all_reviews_export() -> list[dict]:
    with connect() as conn:
        cur = conn.cursor(dictionary=True)
        cur.execute("SELECT * FROM reviews ORDER BY id")
        return cur.fetchall()


def list_all_daily_export() -> list[dict]:
    with connect() as conn:
        cur = conn.cursor(dictionary=True)
        cur.execute("SELECT * FROM daily_reading ORDER BY book_id, read_date")
        rows = cur.fetchall()
    for r in rows:
        _iso_date_row(r)
    return rows


def list_all_book_tags_export() -> list[dict]:
    with connect() as conn:
        cur = conn.cursor(dictionary=True)
        cur.execute("SELECT book_id, tag FROM book_tags ORDER BY book_id, tag")
        return cur.fetchall()


def list_all_goals_export() -> list[dict]:
    with connect() as conn:
        cur = conn.cursor(dictionary=True)
        cur.execute(
            "SELECT scope_year, scope_month, target_pages, target_minutes FROM reading_goals ORDER BY scope_year, scope_month"
        )
        return cur.fetchall()


def parse_import_date(v: Any) -> date | None:
    if v is None:
        return None
    if isinstance(v, date) and not isinstance(v, datetime):
        return v
    if isinstance(v, datetime):
        return v.date()
    s = str(v)[:10]
    try:
        return date.fromisoformat(s)
    except ValueError:
        return None


def export_full_snapshot() -> dict[str, Any]:
    return {
        "version": 2,
        "books": list_all_books(),
        "quotes": list_all_quotes_export(),
        "reviews": list_all_reviews_export(),
        "daily_reading": list_all_daily_export(),
        "book_tags": list_all_book_tags_export(),
        "reading_goals": list_all_goals_export(),
    }


def import_replace_snapshot(data: dict[str, Any]) -> None:
    if not isinstance(data, dict) or "books" not in data:
        raise ValueError("Некорректный файл резервной копии")

    books = data.get("books") or []
    quotes = data.get("quotes") or []
    reviews = data.get("reviews") or []
    daily = data.get("daily_reading") or []
    tags = data.get("book_tags") or []
    goals = data.get("reading_goals") or []

    with connect() as conn:
        cur = conn.cursor()
        cur.execute("SET FOREIGN_KEY_CHECKS = 0")
        cur.execute("DELETE FROM daily_reading")
        cur.execute("DELETE FROM quotes")
        cur.execute("DELETE FROM reviews")
        cur.execute("DELETE FROM book_tags")
        cur.execute("DELETE FROM reading_goals")
        cur.execute("DELETE FROM books")

        for b in books:
            bid = int(b["id"])
            cur.execute(
                """INSERT INTO books (id, title, author, genre, page_count, format_type,
                   plan_year, plan_month, cover_path, date_started, date_finished, reading_status)
                   VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)""",
                (
                    bid,
                    b.get("title") or "",
                    b.get("author") or "",
                    b.get("genre") or "",
                    int(b.get("page_count") or 0),
                    b.get("format_type") or "paper",
                    b.get("plan_year"),
                    b.get("plan_month"),
                    b.get("cover_path"),
                    parse_import_date(b.get("date_started")),
                    parse_import_date(b.get("date_finished")),
                    b.get("reading_status") or READING_PLANNED,
                ),
            )

        for q in quotes:
            cur.execute(
                "INSERT INTO quotes (id, book_id, quote_text) VALUES (%s, %s, %s)",
                (int(q["id"]), int(q["book_id"]), q.get("quote_text") or ""),
            )

        for rev in reviews:
            cur.execute(
                """INSERT INTO reviews (id, book_id, rating_idea, rating_plot,
                   rating_characters, rating_author_skill, review_text)
                   VALUES (%s, %s, %s, %s, %s, %s, %s)""",
                (
                    int(rev["id"]),
                    int(rev["book_id"]),
                    int(rev.get("rating_idea") or 3),
                    int(rev.get("rating_plot") or 3),
                    int(rev.get("rating_characters") or 3),
                    int(rev.get("rating_author_skill") or 3),
                    rev.get("review_text"),
                ),
            )

        for dr in daily:
            rd = parse_import_date(dr.get("read_date"))
            if rd is None:
                continue
            cur.execute(
                """INSERT INTO daily_reading (book_id, read_date, pages_read, minutes_read, note)
                   VALUES (%s, %s, %s, %s, %s)""",
                (
                    int(dr["book_id"]),
                    rd,
                    int(dr.get("pages_read") or 0),
                    int(dr.get("minutes_read") or 0),
                    dr.get("note"),
                ),
            )

        for t in tags:
            tag = _norm_tag(t.get("tag") or "")
            if not tag:
                continue
            cur.execute(
                "INSERT INTO book_tags (book_id, tag) VALUES (%s, %s)",
                (int(t["book_id"]), tag),
            )

        for g in goals:
            cur.execute(
                """INSERT INTO reading_goals (scope_year, scope_month, target_pages, target_minutes)
                   VALUES (%s, %s, %s, %s)""",
                (
                    int(g["scope_year"]),
                    int(g.get("scope_month") or 0),
                    int(g.get("target_pages") or 0),
                    int(g.get("target_minutes") or 0),
                ),
            )

        for tbl in ("books", "quotes", "reviews"):
            cur.execute(f"SELECT COALESCE(MAX(id), 0) + 1 FROM {tbl}")
            nxt = cur.fetchone()[0]
            cur.execute(f"ALTER TABLE {tbl} AUTO_INCREMENT = {nxt}")

        cur.execute("SET FOREIGN_KEY_CHECKS = 1")
        conn.commit()

# В конец файла db.py

def get_reading_stats_per_month() -> list[tuple]:
    """Возвращает список: [(Год-Месяц, Всего страниц), ...] за последние 12 месяцев."""
    with connect() as conn:
        cur = conn.cursor()
        cur.execute(
            """
            SELECT 
                DATE_FORMAT(read_date, '%Y-%m') as month_str,
                SUM(pages_read) as total_pages
            FROM daily_reading
            GROUP BY month_str
            ORDER BY month_str DESC
            LIMIT 12
            """
        )
        # Возвращаем список кортежей и переворачиваем его (чтобы график шел слева направо: от старого к новому)
        return cur.fetchall()[::-1]

def get_genres_stats() -> list[tuple]:
    """Возвращает список жанров и количества книг в них."""
    with connect() as conn:
        cur = conn.cursor()
        cur.execute(
            """
            SELECT genre, COUNT(*) as cnt 
            FROM books 
            WHERE genre IS NOT NULL AND genre != '' 
            GROUP BY genre 
            ORDER BY cnt DESC
            """
        )
        return cur.fetchall()