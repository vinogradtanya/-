"""Диалоговые окна для файловых операций."""
from tkinter import messagebox, filedialog
from pathlib import Path
from datetime import datetime

import backup
import export_excel
import db


def export_review_dialog(book_id: int) -> None:
    """Экспорт отзыва на книгу в TXT или PDF."""
    book = db.get_book(book_id)
    if not book:
        return
    rev = db.get_review(book_id)
    quotes = db.list_quotes(book_id)
    tags = db.get_book_tags(book_id)

    title = book.get("title") or "Без названия"
    author = book.get("author") or ""
    genre = book.get("genre") or ""
    pages = book.get("page_count") or 0
    started = (book.get("date_started") or "")[:10]
    finished = (book.get("date_finished") or "")[:10]

    status_map = {"planned": "В планах", "reading": "Читаю", "finished": "Прочитано"}
    status = status_map.get(book.get("reading_status") or "", "")

    lines = [
        f"ЧИТАТЕЛЬСКИЙ ДНЕВНИК",
        f"Дата экспорта: {datetime.now().strftime('%d.%m.%Y %H:%M')}",
        "=" * 60,
        f"",
        f"Название:  {title}",
        f"Автор:     {author}",
        f"Жанр:      {genre}",
        f"Страниц:   {pages}",
        f"Статус:    {status}",
    ]
    if started:
        lines.append(f"Начало:    {started}")
    if finished:
        lines.append(f"Конец:     {finished}")
    if tags:
        lines.append(f"Теги:      {', '.join(tags)}")

    if rev:
        lines += [
            "",
            "ОЦЕНКИ",
            "-" * 40,
            f"Идея:              {rev.get('rating_idea', '-')} / 5",
            f"Сюжет:             {rev.get('rating_plot', '-')} / 5",
            f"Персонажи:         {rev.get('rating_characters', '-')} / 5",
            f"Мастерство автора: {rev.get('rating_author_skill', '-')} / 5",
        ]
        avg = sum([
            int(rev.get('rating_idea') or 0),
            int(rev.get('rating_plot') or 0),
            int(rev.get('rating_characters') or 0),
            int(rev.get('rating_author_skill') or 0),
        ]) / 4
        lines.append(f"Средняя оценка:    {avg:.1f} / 5")
        rt = (rev.get("review_text") or "").strip()
        if rt:
            lines += ["", "ОТЗЫВ", "-" * 40, rt]

    if quotes:
        lines += ["", "ЦИТАТЫ", "-" * 40]
        for i, q in enumerate(quotes, 1):
            lines.append(f"{i}. {q.get('quote_text', '')}")

    text = "\n".join(lines)

    path = filedialog.asksaveasfilename(
        defaultextension=".txt",
        filetypes=[("Текстовый файл", "*.txt"), ("Все", "*.*")],
        initialfile=f"Отзыв — {title}",
        title="Экспорт отзыва",
    )
    if not path:
        return
    try:
        Path(path).write_text(text, encoding="utf-8")
        messagebox.showinfo("Экспорт", f"Отзыв сохранён:\n{path}")
    except Exception as e:
        messagebox.showerror("Ошибка", str(e))


def export_backup_dialog() -> None:
    """Диалог для экспорта данных в JSON."""
    path = filedialog.asksaveasfilename(
        defaultextension=".json",
        filetypes=[("JSON", "*.json"), ("Все", "*.*")],
        title="Экспорт данных",
    )
    if not path:
        return
    try:
        backup.export_to_json(path)
        messagebox.showinfo("Экспорт", f"Данные сохранены в файл:\n{path}")
    except Exception as e:
        messagebox.showerror("Ошибка", str(e))


def import_backup_dialog() -> tuple[bool, str | None]:
    """
    Диалог для импорта данных из JSON.
    
    Возвращает (успешно ли импортировано, путь к файлу).
    """
    path = filedialog.askopenfilename(
        filetypes=[("JSON", "*.json"), ("Все", "*.*")],
        title="Импорт данных",
    )
    if not path:
        return False, None
    
    if not messagebox.askyesno(
        "Импорт",
        "Текущие данные в базе будут полностью заменены содержимым файла. Продолжить?",
    ):
        return False, None
    
    try:
        backup.import_from_json(path)
        messagebox.showinfo("Импорт", "Данные загружены.")
        return True, path
    except Exception as e:
        messagebox.showerror("Ошибка", str(e))
        return False, None


def export_excel_dialog() -> None:
    """Диалог для экспорта статистики в Excel."""
    # Проверяем доступность openpyxl
    available, msg = export_excel.get_export_status()
    if not available:
        messagebox.showwarning(
            "Excel экспорт",
            f"Недоступно: {msg}\n\nУстановите пакеты командой:\npip install openpyxl pillow"
        )
        return
    
    path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel", "*.xlsx"), ("Все", "*.*")],
        title="Экспорт статистики в Excel",
    )
    if not path:
        return
    
    try:
        export_excel.export_statistics_to_excel(path)
        messagebox.showinfo("Экспорт", f"Статистика сохранена в файл:\n{path}")
    except Exception as e:
        messagebox.showerror("Ошибка при экспорте", str(e))
