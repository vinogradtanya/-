"""Вспомогательные функции и утилиты."""
from datetime import date, datetime


def today() -> date:
    """Возвращает сегодняшнюю дату."""
    return date.today()


def parse_date(s: str) -> date | None:
    """Парсирует строку в дату (формат YYYY-MM-DD)."""
    s = (s or "").strip()
    if not s:
        return None
    try:
        return datetime.strptime(s, "%Y-%m-%d").date()
    except ValueError:
        return None


def fmt_date(d: date | None) -> str:
    """Форматирует дату в строку (ISO 8601)."""
    return d.isoformat() if d else ""


def int_or(s: str, default: int = 0) -> int:
    """Преобразует строку в целое число или возвращает default."""
    try:
        return int((s or "").strip())
    except ValueError:
        return default
