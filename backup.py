"""Экспорт и импорт данных в JSON (без внешних сервисов)."""
from __future__ import annotations
import json
from datetime import datetime
from pathlib import Path

import db

def export_to_json(path: str | Path) -> None:
    data = db.export_full_snapshot()
    data["exported_at"] = datetime.now().isoformat(timespec="seconds")
    Path(path).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def import_from_json(path: str | Path) -> None:
    raw = Path(path).read_text(encoding="utf-8")
    data = json.loads(raw)
    db.import_replace_snapshot(data)
