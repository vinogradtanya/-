"""Юниттесты для всех модулей читательского дневника.

Запуск:
    py -3.12 -m pytest test_units.py -v
    # или без pytest:
    py -3.12 test_units.py
"""
from __future__ import annotations

import json
import sys
import types
import unittest
from datetime import date, datetime
from pathlib import Path
from unittest.mock import MagicMock, patch

# Заглушки для GUI-зависимостей (до импорта модулей)
for _mod in ("customtkinter", "tkinter", "matplotlib", "matplotlib.pyplot",
             "matplotlib.backends", "matplotlib.backends.backend_tkagg",
             "matplotlib.ticker", "openpyxl", "openpyxl.styles",
             "openpyxl.drawing", "openpyxl.drawing.image",
             "plyer", "plyer.notification", "PIL", "PIL.Image",
             "PIL.ImageTk"):
    if _mod not in sys.modules:
        sys.modules[_mod] = MagicMock()

# Заглушка tkinter.messagebox и filedialog
_tk = sys.modules.get("tkinter", MagicMock())
_tk.messagebox = MagicMock()
_tk.filedialog = MagicMock()
sys.modules["tkinter"] = _tk
sys.modules["tkinter.messagebox"] = _tk.messagebox
sys.modules["tkinter.filedialog"] = _tk.filedialog
sys.modules["tkinter.simpledialog"] = MagicMock()

# Импорт тестируемых модулей
import utils
import constants
import config


# utils.py

class TestParseDate(unittest.TestCase):

    def test_valid_date(self):
        self.assertEqual(utils.parse_date("2024-06-15"), date(2024, 6, 15))

    def test_empty_string(self):
        self.assertIsNone(utils.parse_date(""))

    def test_none_input(self):
        self.assertIsNone(utils.parse_date(None))

    def test_whitespace(self):
        self.assertIsNone(utils.parse_date("   "))

    def test_invalid_format(self):
        self.assertIsNone(utils.parse_date("15.06.2024"))

    def test_invalid_date(self):
        self.assertIsNone(utils.parse_date("2024-13-01"))

    def test_strips_whitespace(self):
        self.assertEqual(utils.parse_date("  2024-01-01  "), date(2024, 1, 1))


class TestFmtDate(unittest.TestCase):

    def test_formats_date(self):
        self.assertEqual(utils.fmt_date(date(2024, 6, 15)), "2024-06-15")

    def test_none_returns_empty(self):
        self.assertEqual(utils.fmt_date(None), "")


class TestIntOr(unittest.TestCase):

    def test_valid_int(self):
        self.assertEqual(utils.int_or("42"), 42)

    def test_empty_string(self):
        self.assertEqual(utils.int_or(""), 0)

    def test_none(self):
        self.assertEqual(utils.int_or(None), 0)

    def test_non_numeric(self):
        self.assertEqual(utils.int_or("abc"), 0)

    def test_custom_default(self):
        self.assertEqual(utils.int_or("xyz", 99), 99)

    def test_negative(self):
        self.assertEqual(utils.int_or("-5"), -5)

    def test_whitespace_around_number(self):
        self.assertEqual(utils.int_or("  7  "), 7)


class TestToday(unittest.TestCase):

    def test_returns_date(self):
        self.assertIsInstance(utils.today(), date)

    def test_equals_date_today(self):
        self.assertEqual(utils.today(), date.today())

# constants.py

class TestConstants(unittest.TestCase):

    def test_colors_keys(self):
        for key in ("primary", "success", "danger", "warning", "secondary", "muted"):
            self.assertIn(key, constants.COLORS)

    def test_colors_are_hex(self):
        for v in constants.COLORS.values():
            self.assertTrue(v.startswith("#"), f"{v} не является hex-цветом")

    def test_format_labels_keys(self):
        self.assertIn("paper", constants.FORMAT_LABELS)
        self.assertIn("ebook", constants.FORMAT_LABELS)
        self.assertIn("audiobook", constants.FORMAT_LABELS)

    def test_format_values(self):
        self.assertEqual(set(constants.FORMAT_VALUES), {"paper", "ebook", "audiobook"})

    def test_status_label_to_key(self):
        self.assertEqual(constants.STATUS_LABEL_TO_KEY["В планах"], "planned")
        self.assertEqual(constants.STATUS_LABEL_TO_KEY["Читаю"], "reading")
        self.assertEqual(constants.STATUS_LABEL_TO_KEY["Прочитано"], "finished")
        self.assertEqual(constants.STATUS_LABEL_TO_KEY["Брошено"], "abandoned")

    def test_status_key_to_label_roundtrip(self):
        for label, key in constants.STATUS_LABEL_TO_KEY.items():
            self.assertEqual(constants.STATUS_KEY_TO_LABEL[key], label)

    def test_window_dimensions_positive(self):
        self.assertGreater(constants.WINDOW_MIN_WIDTH, 0)
        self.assertGreater(constants.WINDOW_MIN_HEIGHT, 0)
        self.assertGreaterEqual(constants.WINDOW_DEFAULT_WIDTH, constants.WINDOW_MIN_WIDTH)
        self.assertGreaterEqual(constants.WINDOW_DEFAULT_HEIGHT, constants.WINDOW_MIN_HEIGHT)

    def test_cover_dimensions_positive(self):
        self.assertGreater(constants.COVER_DISPLAY_WIDTH, 0)
        self.assertGreater(constants.COVER_DISPLAY_HEIGHT, 0)

# config.py

class TestConfig(unittest.TestCase):

    def test_data_dir_is_path(self):
        self.assertIsInstance(config.DATA_DIR, Path)

    def test_covers_dir_is_path(self):
        self.assertIsInstance(config.COVERS_DIR, Path)

    def test_covers_dir_inside_data_dir(self):
        self.assertTrue(str(config.COVERS_DIR).startswith(str(config.DATA_DIR)))

    def test_mysql_config_keys(self):
        cfg = config.mysql_config()
        for key in ("host", "port", "user", "password", "database"):
            self.assertIn(key, cfg)

    def test_mysql_config_port_is_int(self):
        self.assertIsInstance(config.mysql_config()["port"], int)

# db.py — через моки (без реальной БД)

def _make_cursor(rows=None, lastrowid=1):
    """Создаёт мок-курсор MySQL."""
    cur = MagicMock()
    cur.fetchone.return_value = rows[0] if rows else None
    cur.fetchall.return_value = rows or []
    cur.lastrowid = lastrowid
    return cur


def _make_conn(cursor):
    conn = MagicMock()
    conn.cursor.return_value = cursor
    conn.__enter__ = lambda s: s
    conn.__exit__ = MagicMock(return_value=False)
    return conn


class TestDbRowBook(unittest.TestCase):
    """Тесты вспомогательной функции _row_book."""

    def setUp(self):
        import db as _db
        self.db = _db

    def test_none_returns_none(self):
        self.assertIsNone(self.db._row_book(None))

    def test_date_converted_to_iso(self):
        row = {"date_started": date(2024, 1, 15), "title": "Test"}
        result = self.db._row_book(row)
        self.assertEqual(result["date_started"], "2024-01-15")

    def test_datetime_converted_to_iso(self):
        row = {"date_finished": datetime(2024, 3, 20, 10, 0), "title": "Test"}
        result = self.db._row_book(row)
        self.assertEqual(result["date_finished"], "2024-03-20T10:00:00")

    def test_none_date_stays_none(self):
        row = {"date_started": None, "title": "Test"}
        result = self.db._row_book(row)
        self.assertIsNone(result["date_started"])

    def test_non_date_fields_unchanged(self):
        row = {"title": "Война и мир", "author": "Толстой"}
        result = self.db._row_book(row)
        self.assertEqual(result["title"], "Война и мир")


class TestDbParseImportDate(unittest.TestCase):

    def setUp(self):
        import db as _db
        self.db = _db

    def test_none(self):
        self.assertIsNone(self.db.parse_import_date(None))

    def test_date_object(self):
        d = date(2024, 5, 10)
        self.assertEqual(self.db.parse_import_date(d), d)

    def test_datetime_object(self):
        dt = datetime(2024, 5, 10, 12, 0)
        self.assertEqual(self.db.parse_import_date(dt), date(2024, 5, 10))

    def test_iso_string(self):
        self.assertEqual(self.db.parse_import_date("2024-05-10"), date(2024, 5, 10))

    def test_invalid_string(self):
        self.assertIsNone(self.db.parse_import_date("not-a-date"))


class TestDbNormTag(unittest.TestCase):

    def setUp(self):
        import db as _db
        self.db = _db

    def test_lowercases(self):
        self.assertEqual(self.db._norm_tag("Фантастика"), "фантастика")

    def test_strips_whitespace(self):
        self.assertEqual(self.db._norm_tag("  тег  "), "тег")

    def test_empty_string(self):
        self.assertEqual(self.db._norm_tag(""), "")

    def test_none(self):
        self.assertEqual(self.db._norm_tag(None), "")

    def test_truncates_to_80(self):
        long_tag = "а" * 100
        self.assertEqual(len(self.db._norm_tag(long_tag)), 80)


class TestDbStatusConstants(unittest.TestCase):

    def setUp(self):
        import db as _db
        self.db = _db

    def test_status_values_complete(self):
        self.assertIn("planned", self.db.READING_STATUS_VALUES)
        self.assertIn("reading", self.db.READING_STATUS_VALUES)
        self.assertIn("finished", self.db.READING_STATUS_VALUES)
        self.assertIn("abandoned", self.db.READING_STATUS_VALUES)

    def test_constants_match(self):
        self.assertEqual(self.db.READING_PLANNED, "planned")
        self.assertEqual(self.db.READING_READING, "reading")
        self.assertEqual(self.db.READING_FINISHED, "finished")
        self.assertEqual(self.db.READING_ABANDONED, "abandoned")


class TestDbCreateBook(unittest.TestCase):

    def setUp(self):
        import db as _db
        self.db = _db

    def test_create_book_returns_id(self):
        cur = _make_cursor(lastrowid=42)
        conn = _make_conn(cur)
        with patch.object(self.db, "connect", return_value=conn):
            bid = self.db.create_book("Тест", author="Автор", page_count=300)
        self.assertEqual(bid, 42)

    def test_invalid_status_defaults_to_planned(self):
        cur = _make_cursor(lastrowid=1)
        conn = _make_conn(cur)
        with patch.object(self.db, "connect", return_value=conn):
            self.db.create_book("Книга", reading_status="invalid_status")
        args = cur.execute.call_args[0][1]
        # reading_status — последний аргумент в INSERT
        self.assertEqual(args[-1], "planned")

    def test_page_count_calculated_from_range(self):
        cur = _make_cursor(lastrowid=1)
        conn = _make_conn(cur)
        with patch.object(self.db, "connect", return_value=conn):
            self.db.create_book("Книга", page_start=10, page_end=109)
        args = cur.execute.call_args[0][1]
        # page_count = page_end - page_start + 1 = 100
        self.assertEqual(args[3], 100)


class TestDbGetBook(unittest.TestCase):

    def setUp(self):
        import db as _db
        self.db = _db

    def test_returns_none_when_not_found(self):
        cur = _make_cursor(rows=[None])
        conn = _make_conn(cur)
        cur.fetchone.return_value = None
        with patch.object(self.db, "connect", return_value=conn):
            result = self.db.get_book(999)
        self.assertIsNone(result)

    def test_returns_dict(self):
        cur = _make_cursor(rows=[{"id": 1, "title": "Тест", "date_started": None}])
        conn = _make_conn(cur)
        with patch.object(self.db, "connect", return_value=conn):
            result = self.db.get_book(1)
        self.assertEqual(result["title"], "Тест")


class TestDbUpdateBook(unittest.TestCase):

    def setUp(self):
        import db as _db
        self.db = _db

    def test_no_fields_does_nothing(self):
        with patch.object(self.db, "connect") as mock_connect:
            self.db.update_book(1)
        mock_connect.assert_not_called()

    def test_unknown_fields_ignored(self):
        with patch.object(self.db, "connect") as mock_connect:
            self.db.update_book(1, unknown_field="value")
        mock_connect.assert_not_called()

    def test_valid_field_executes_query(self):
        cur = _make_cursor()
        conn = _make_conn(cur)
        with patch.object(self.db, "connect", return_value=conn):
            self.db.update_book(1, title="Новое название")
        cur.execute.assert_called_once()
        sql = cur.execute.call_args[0][0]
        self.assertIn("UPDATE books", sql)


class TestDbDeleteBook(unittest.TestCase):

    def setUp(self):
        import db as _db
        self.db = _db

    def test_deletes_related_tables(self):
        cur = _make_cursor()
        conn = _make_conn(cur)
        with patch.object(self.db, "connect", return_value=conn):
            self.db.delete_book(5)
        calls_sql = [c[0][0] for c in cur.execute.call_args_list]
        self.assertTrue(any("daily_reading" in s for s in calls_sql))
        self.assertTrue(any("quotes" in s for s in calls_sql))
        self.assertTrue(any("reviews" in s for s in calls_sql))
        self.assertTrue(any("book_tags" in s for s in calls_sql))
        self.assertTrue(any("DELETE FROM books" in s for s in calls_sql))


class TestDbQuotes(unittest.TestCase):

    def setUp(self):
        import db as _db
        self.db = _db

    def test_add_quote_returns_id(self):
        cur = _make_cursor(lastrowid=7)
        conn = _make_conn(cur)
        with patch.object(self.db, "connect", return_value=conn):
            qid = self.db.add_quote(1, "Текст цитаты", page_number=42)
        self.assertEqual(qid, 7)

    def test_delete_quote_executes(self):
        cur = _make_cursor()
        conn = _make_conn(cur)
        with patch.object(self.db, "connect", return_value=conn):
            self.db.delete_quote(3)
        sql = cur.execute.call_args[0][0]
        self.assertIn("DELETE FROM quotes", sql)

    def test_update_quote_executes(self):
        cur = _make_cursor()
        conn = _make_conn(cur)
        with patch.object(self.db, "connect", return_value=conn):
            self.db.update_quote(1, "Новый текст", 10)
        sql = cur.execute.call_args[0][0]
        self.assertIn("UPDATE quotes", sql)

    def test_list_quotes_returns_list(self):
        rows = [{"id": 1, "book_id": 1, "quote_text": "Цитата"}]
        cur = _make_cursor(rows=rows)
        conn = _make_conn(cur)
        with patch.object(self.db, "connect", return_value=conn):
            result = self.db.list_quotes(1)
        self.assertEqual(result, rows)


class TestDbTags(unittest.TestCase):

    def setUp(self):
        import db as _db
        self.db = _db

    def test_replace_book_tags_deduplicates(self):
        cur = _make_cursor()
        conn = _make_conn(cur)
        with patch.object(self.db, "connect", return_value=conn):
            self.db.replace_book_tags(1, ["Тег", "тег", "ТЕГ", "другой"])
        insert_calls = [c for c in cur.execute.call_args_list
                        if "INSERT INTO book_tags" in str(c)]
        self.assertEqual(len(insert_calls), 2)  # "тег" и "другой"

    def test_replace_book_tags_skips_empty(self):
        cur = _make_cursor()
        conn = _make_conn(cur)
        with patch.object(self.db, "connect", return_value=conn):
            self.db.replace_book_tags(1, ["", "  ", "нормальный"])
        insert_calls = [c for c in cur.execute.call_args_list
                        if "INSERT INTO book_tags" in str(c)]
        self.assertEqual(len(insert_calls), 1)

    def test_get_book_tags_returns_list(self):
        cur = _make_cursor(rows=[("фантастика",), ("детектив",)])
        conn = _make_conn(cur)
        with patch.object(self.db, "connect", return_value=conn):
            tags = self.db.get_book_tags(1)
        self.assertEqual(tags, ["фантастика", "детектив"])


class TestDbUpsertDaily(unittest.TestCase):

    def setUp(self):
        import db as _db
        self.db = _db

    def test_upsert_daily_executes_insert(self):
        cur = _make_cursor()
        conn = _make_conn(cur)
        with patch.object(self.db, "connect", return_value=conn):
            self.db.upsert_daily(1, date(2024, 6, 1), pages_read=30)
        sql = cur.execute.call_args[0][0]
        self.assertIn("INSERT INTO daily_reading", sql)

    def test_delete_daily_entry_executes(self):
        cur = _make_cursor()
        conn = _make_conn(cur)
        with patch.object(self.db, "connect", return_value=conn):
            self.db.delete_daily_entry(1, date(2024, 6, 1))
        sql = cur.execute.call_args[0][0]
        self.assertIn("DELETE FROM daily_reading", sql)


class TestDbReview(unittest.TestCase):

    def setUp(self):
        import db as _db
        self.db = _db

    def test_upsert_review_executes(self):
        cur = _make_cursor()
        conn = _make_conn(cur)
        with patch.object(self.db, "connect", return_value=conn):
            self.db.upsert_review(1, 4, 3, 5, 4, "Отличная книга")
        sql = cur.execute.call_args[0][0]
        self.assertIn("INSERT INTO reviews", sql)

    def test_get_review_returns_none_when_missing(self):
        cur = _make_cursor()
        cur.fetchone.return_value = None
        conn = _make_conn(cur)
        with patch.object(self.db, "connect", return_value=conn):
            result = self.db.get_review(999)
        self.assertIsNone(result)


class TestDbReadingGoals(unittest.TestCase):

    def setUp(self):
        import db as _db
        self.db = _db

    def test_upsert_reading_goal_clamps_month(self):
        cur = _make_cursor()
        conn = _make_conn(cur)
        with patch.object(self.db, "connect", return_value=conn):
            self.db.upsert_reading_goal(2024, 15, 10)  # месяц > 12
        args = cur.execute.call_args[0][1]
        self.assertEqual(args[1], 12)  # зажат до 12

    def test_upsert_reading_goal_negative_month(self):
        cur = _make_cursor()
        conn = _make_conn(cur)
        with patch.object(self.db, "connect", return_value=conn):
            self.db.upsert_reading_goal(2024, -1, 5)
        args = cur.execute.call_args[0][1]
        self.assertEqual(args[1], 0)  # зажат до 0

    def test_upsert_reading_goal_negative_books(self):
        cur = _make_cursor()
        conn = _make_conn(cur)
        with patch.object(self.db, "connect", return_value=conn):
            self.db.upsert_reading_goal(2024, 6, -5)
        args = cur.execute.call_args[0][1]
        self.assertEqual(args[2], 0)  # зажат до 0


class TestDbSearch(unittest.TestCase):

    def setUp(self):
        import db as _db
        self.db = _db

    def test_empty_query_returns_empty(self):
        result = self.db.search_books("")
        self.assertEqual(result, [])

    def test_like_contains_escapes_percent(self):
        result = self.db._like_contains("100%")
        self.assertIn("|%", result)

    def test_like_contains_escapes_underscore(self):
        result = self.db._like_contains("test_book")
        self.assertIn("|_", result)

    def test_like_contains_wraps_with_percent(self):
        result = self.db._like_contains("слово")
        self.assertTrue(result.startswith("%"))
        self.assertTrue(result.endswith("%"))


class TestDbExportSnapshot(unittest.TestCase):

    def setUp(self):
        import db as _db
        self.db = _db

    def test_snapshot_has_required_keys(self):
        with patch.object(self.db, "list_all_books", return_value=[]), \
             patch.object(self.db, "list_all_quotes_export", return_value=[]), \
             patch.object(self.db, "list_all_reviews_export", return_value=[]), \
             patch.object(self.db, "list_all_daily_export", return_value=[]), \
             patch.object(self.db, "list_all_book_tags_export", return_value=[]), \
             patch.object(self.db, "list_all_goals_export", return_value=[]):
            snap = self.db.export_full_snapshot()
        for key in ("version", "books", "quotes", "reviews", "daily_reading", "book_tags", "reading_goals"):
            self.assertIn(key, snap)

    def test_snapshot_version(self):
        with patch.object(self.db, "list_all_books", return_value=[]), \
             patch.object(self.db, "list_all_quotes_export", return_value=[]), \
             patch.object(self.db, "list_all_reviews_export", return_value=[]), \
             patch.object(self.db, "list_all_daily_export", return_value=[]), \
             patch.object(self.db, "list_all_book_tags_export", return_value=[]), \
             patch.object(self.db, "list_all_goals_export", return_value=[]):
            snap = self.db.export_full_snapshot()
        self.assertEqual(snap["version"], 2)


class TestDbImportSnapshot(unittest.TestCase):

    def setUp(self):
        import db as _db
        self.db = _db

    def test_invalid_data_raises(self):
        with self.assertRaises(ValueError):
            self.db.import_replace_snapshot({"no_books_key": []})

    def test_non_dict_raises(self):
        with self.assertRaises((ValueError, AttributeError)):
            self.db.import_replace_snapshot("not a dict")

# backup.py

class TestBackup(unittest.TestCase):

    def setUp(self):
        import backup as _backup
        self.backup = _backup

    def test_export_writes_json(self):
        snapshot = {"version": 2, "books": [], "quotes": [], "reviews": [],
                    "daily_reading": [], "book_tags": [], "reading_goals": []}
        import db as _db
        mock_path = MagicMock()
        with patch.object(_db, "export_full_snapshot", return_value=snapshot), \
             patch("backup.Path", return_value=mock_path):
            self.backup.export_to_json("/tmp/test.json")
        mock_path.write_text.assert_called_once()

    def test_import_calls_db(self):
        data = {"version": 2, "books": [], "quotes": [], "reviews": [],
                "daily_reading": [], "book_tags": [], "reading_goals": []}
        import db as _db
        mock_path = MagicMock()
        mock_path.read_text.return_value = json.dumps(data)
        with patch("backup.Path", return_value=mock_path), \
             patch.object(_db, "import_replace_snapshot") as mock_import:
            self.backup.import_from_json("/tmp/test.json")
        mock_import.assert_called_once_with(data)

    def test_export_adds_exported_at(self):
        snapshot = {"version": 2, "books": [], "quotes": [], "reviews": [],
                    "daily_reading": [], "book_tags": [], "reading_goals": []}
        written_data = {}

        def fake_write_text(text, encoding=None):
            written_data["content"] = json.loads(text)

        import db as _db
        mock_path = MagicMock()
        mock_path.write_text.side_effect = fake_write_text

        with patch.object(_db, "export_full_snapshot", return_value=snapshot), \
             patch("backup.Path", return_value=mock_path):
            self.backup.export_to_json("/tmp/test.json")

        self.assertIn("exported_at", written_data.get("content", {}))

# export_excel.py

class TestExportExcel(unittest.TestCase):

    def setUp(self):
        import export_excel as _ex
        self.ex = _ex

    def test_get_export_status_returns_tuple(self):
        ok, msg = self.ex.get_export_status()
        self.assertIsInstance(ok, bool)
        self.assertIsInstance(msg, str)

    def test_get_export_status_message_not_empty(self):
        _, msg = self.ex.get_export_status()
        self.assertTrue(len(msg) > 0)

    def test_export_raises_without_openpyxl(self):
        original = self.ex.HAS_OPENPYXL
        try:
            self.ex.HAS_OPENPYXL = False
            with self.assertRaises(ImportError):
                self.ex.export_statistics_to_excel("/tmp/test.xlsx")
        finally:
            self.ex.HAS_OPENPYXL = original

# Точка входа

if __name__ == "__main__":
    unittest.main(verbosity=2)
