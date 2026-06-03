"""Загрузка настроек подключения к MySQL (.env или переменные окружения)."""
import os
from pathlib import Path

try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).resolve().parent / ".env")
except ImportError:
    pass

DATA_DIR = Path(__file__).resolve().parent / "data"
DATA_DIR.mkdir(exist_ok=True)
COVERS_DIR = DATA_DIR / "covers"
COVERS_DIR.mkdir(exist_ok=True)


def mysql_config():
    return {
        "host": os.getenv("MYSQL_HOST"),
        "port": int(os.getenv("MYSQL_PORT", "3306")),
        "user": os.getenv("MYSQL_USER"),
        "password": os.getenv("MYSQL_PASSWORD"),
        "database": os.getenv("MYSQL_DATABASE"),
        "charset": "utf8mb4",
        "use_unicode": True,
        "autocommit": True,
    }
