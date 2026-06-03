"""Константы и конфигурация приложения."""
from pathlib import Path

from config import DATA_DIR

# Современные цвета интерфейса
COLORS = {
    "primary": "#3B82F6",      # Синий
    "success": "#10B981",      # Зеленый
    "danger": "#EF4444",       # Красный
    "warning": "#F59E0B",      # Оранжевый
    "secondary": "#8B5CF6",    # Фиолетовый
    "muted": "#6B7280",        # Серый
}

# Форматы книг
FORMAT_LABELS = {
    "paper": "Бумажная",
    "ebook": "Электронная",
    "audiobook": "Аудиокнига",
}
FORMAT_VALUES = list(FORMAT_LABELS.keys())

# Статусы чтения
STATUS_LABEL_TO_KEY = {"В планах": "planned", "Читаю": "reading", "Прочитано": "finished", "Брошено": "abandoned"}
STATUS_KEY_TO_LABEL = {v: k for k, v in STATUS_LABEL_TO_KEY.items()}

# Пути
SETTINGS_PATH = DATA_DIR / "settings.json"

# Окно приложения
WINDOW_MIN_WIDTH = 1100
WINDOW_MIN_HEIGHT = 720
WINDOW_DEFAULT_WIDTH = 1240
WINDOW_DEFAULT_HEIGHT = 800

# Размеры обложек
COVER_DISPLAY_WIDTH = 200
COVER_DISPLAY_HEIGHT = 280
