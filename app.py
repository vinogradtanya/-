"""
Читательский дневник — десктоп-приложение (CustomTkinter + MySQL).
Архитектура: основной класс приложения (Controller) с делегированием UI модулям.
"""
from __future__ import annotations

import calendar
import json
from datetime import date, datetime, timedelta
from pathlib import Path
from tkinter import messagebox, filedialog
import tkinter.simpledialog as sd
import tkinter.messagebox as messagebox

import customtkinter as ctk
from PIL import Image

import db
import dialogs
from pages import home
from pages import settings
from pages import stats
from pages import plans
from constants import (
    COLORS, FORMAT_LABELS, FORMAT_VALUES, STATUS_LABEL_TO_KEY, STATUS_KEY_TO_LABEL,
    SETTINGS_PATH, WINDOW_MIN_WIDTH, WINDOW_MIN_HEIGHT, WINDOW_DEFAULT_WIDTH, WINDOW_DEFAULT_HEIGHT
)
from utils import today, parse_date, fmt_date, int_or

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class ReadingDiaryApp(ctk.CTk):
    """Главное окно приложения 'Читательский дневник'."""

    def __init__(self) -> None:
        super().__init__()
        self.title("Читательский дневник")
        self.minsize(WINDOW_MIN_WIDTH, WINDOW_MIN_HEIGHT)
        self.geometry(f"{WINDOW_DEFAULT_WIDTH}x{WINDOW_DEFAULT_HEIGHT}")

        db.ensure_schema()

        # Тема
        self.var_theme = ctk.StringVar(value="Тёмная")
        self._theme_map = {"Тёмная": "dark", "Светлая": "light", "Системная": "System"}
        self._theme_inv = {v: k for k, v in self._theme_map.items()}

        # Настройки приложения
        self._app_settings = {
            "appearance": "dark",
            "ui_scale": 1.0,
            "compact_list": False,
            "reminder_enabled": False,
            "reminder_time": "20:00",
            "last_reminder_date": "",
        }
        self._load_settings()

        # Состояние приложения
        self._current_book_id: int | None = None
        self._books_cache: list[dict] = []
        self._cover_photo: ctk.CTkImage | None = None
        self._cover_pil_image: Image.Image | None = None
        self._cover_pil_image_original: Image.Image | None = None
        self._page_frames: dict[str, ctk.CTkFrame] = {}
        self._current_page_key = "home"
        self._status_scrolls: dict[str, ctk.CTkScrollableFrame] = {}

        # Общие переменные, используемые в нескольких модулях
        cy, cm = today().year, today().month
        self.var_plan_year = ctk.StringVar(value=str(cy))
        self.var_plan_month = ctk.StringVar(value=str(cm))

        # Инициализация UI
        self._build()
        self._show_page("home")
        self._refresh_book_list(select_id=None)
        self.after(2000, self._reminder_tick)

    # Основная сборка окна
    def _build(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # Заголовок
        head = ctk.CTkFrame(self, fg_color="transparent")
        head.grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 12))
        head.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(head, text="📚 Читательский дневник", font=ctk.CTkFont(size=28, weight="bold")).grid(
            row=0, column=0, sticky="w"
        )
        ctk.CTkLabel(head, text="Тема интерфейса: ", font=ctk.CTkFont(size=11)).grid(
            row=0, column=1, padx=(16, 6), sticky="e"
        )
        ctk.CTkOptionMenu(
            head, values=["Тёмная", "Светлая", "Системная"], variable=self.var_theme,
            command=self._on_theme_change, width=130, fg_color=COLORS["primary"], button_color=COLORS["primary"]
        ).grid(row=0, column=2, sticky="e")

        # Навигация
        nav = ctk.CTkFrame(self, fg_color="transparent")
        nav.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 8))

        self._nav_buttons = {}
        row1 = [("🏠 Главная", "home"), ("📅 План/месяц", "plan_month"), ("📆 План/год", "plan_year"),
                ("📋 В планах", "status_planned"), ("📖 Читаю", "status_reading"),
                ("✅ Прочитано", "status_finished"), ("🚫 Брошено", "status_abandoned")]
        row2 = [("📊 Аналитика", "stats"), ("📈 Графики", "charts"), ("🔍 Поиск", "search"), ("⚙️ Настройки", "settings")]

        for col, (text, key) in enumerate(row1):
            btn = ctk.CTkButton(nav, text=text, width=110, command=lambda k=key: self._show_page(k),
                                fg_color=COLORS["primary"], hover_color=COLORS["secondary"], font=ctk.CTkFont(size=11))
            btn.grid(row=0, column=col, padx=3, pady=(4, 2))
            self._nav_buttons[key] = btn

        for col, (text, key) in enumerate(row2):
            btn = ctk.CTkButton(nav, text=text, width=130, command=lambda k=key: self._show_page(k),
                                fg_color=COLORS["secondary"], hover_color=COLORS["primary"], font=ctk.CTkFont(size=11))
            btn.grid(row=1, column=col, padx=3, pady=(2, 4))
            self._nav_buttons[key] = btn

        # Контейнер для страниц
        self._content_host = ctk.CTkFrame(self, fg_color="transparent")
        self._content_host.grid(row=2, column=0, sticky="nsew", padx=16, pady=(0, 16))
        self._content_host.grid_columnconfigure(0, weight=1)
        self._content_host.grid_rowconfigure(0, weight=1)

        self._setup_hotkeys()

    # Навигация и ленивая загрузка страниц
    def _show_page(self, key: str) -> None:
        if key not in self._page_frames:
            # Ленивая загрузка из модулей
            if key == "home":
                self._page_frames["home"] = home.build(self)
            elif key == "plan_month":
                self._page_frames["plan_month"] = plans.build_month(self)
            elif key == "plan_year":
                self._page_frames["plan_year"] = plans.build_year(self)
            elif key == "status_planned":
                self._page_frames[key] = plans.build_status(self, "planned", key, "Книги: в планах")
            elif key == "status_reading":
                self._page_frames[key] = plans.build_status(self, "reading", key, "Книги: читаю")
            elif key == "status_finished":
                self._page_frames[key] = plans.build_status(self, "finished", key, "Книги: прочитано")
            elif key == "status_abandoned":
                self._page_frames[key] = plans.build_status(self, "abandoned", key, "Книги: брошено")
            elif key == "stats":
                self._page_frames["stats"] = stats.build_stats(self)
            elif key == "charts":
                self._page_frames["charts"] = stats.build_charts(self)
            elif key == "search":
                self._page_frames["search"] = stats.build_search(self)
            elif key == "settings":
                self._page_frames["settings"] = settings.build(self)
            else:
                return

        # Переключение видимости
        for frame in self._page_frames.values():
            frame.grid_remove()
        self._page_frames[key].grid(row=0, column=0, sticky="nsew")
        self._current_page_key = key

        # Вызов обновления данных для текущей страницы
        if key == "plan_month":
            plans.refresh_month(self)
        elif key == "plan_year":
            plans.refresh_year(self)
        elif key.startswith("status_"):
            st_key = key.replace("status_", "")
            plans.refresh_status(self, st_key, self._status_scrolls.get(key))
        elif key == "stats":
            stats.refresh_stats(self)
        elif key == "charts":
            stats.refresh_charts(self)
        elif key == "settings":
            settings.sync_widgets(self)

    # Управление книгами и списком
    def _refresh_book_list(self, select_id: int | None) -> None:
        if "home" not in self._page_frames or not hasattr(self, "list_books"):
            return
        for w in self.list_books.winfo_children():
            w.destroy()

        self._books_cache = self._filtered_books()
        try:
            all_for_tracker = db.list_all_books()
        except Exception:
            all_for_tracker = self._books_cache

        self._tracker_book_ids = [b['id'] for b in all_for_tracker]
        titles_for_combo = [(b.get('title') or '')[:42] for b in all_for_tracker]
        self.combo_track_book.configure(values=titles_for_combo or ["— нет книг —"])
        if titles_for_combo:
            self.combo_track_book.set(titles_for_combo[0])

        compact = bool(self._app_settings.get("compact_list"))
        tlen = 22 if compact else 42
        count = len(self._books_cache)
        self.list_books.configure(label_text=f"Книг в списке: {count}")
        for b in self._books_cache:
            title = (b.get("title") or "")[:tlen]
            btn = ctk.CTkButton(
                self.list_books, text=title, anchor="w", fg_color="transparent",
                hover_color=("gray70", "gray50"), text_color=("gray10", "gray90"),
                command=lambda bid=b["id"]: self._load_book(bid)
            )
            btn.pack(fill="x", pady=2)

        if select_id and any(b["id"] == select_id for b in self._books_cache):
            self._load_book(select_id)
        elif self._books_cache:
            self._load_book(self._books_cache[0]["id"])
        else:
            self._new_book(clear_only=True)

    def _filtered_books(self) -> list[dict]:
        try:
            all_b = db.list_all_books()
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
            return []

        mode = self.var_filter_mode.get()
        if mode == "План на месяц":
            y = int_or(self.var_plan_year.get(), today().year)
            m = max(1, min(12, int_or(self.var_plan_month.get(), today().month)))
            return [b for b in all_b if b.get("plan_year") == y and b.get("plan_month") == m]
        if mode == "По тегу":
            sel = (self.var_filter_tag.get() or "").strip()
            if not sel or sel.startswith("—"):
                return all_b
            try:
                return db.list_books_by_tag(sel)
            except Exception as e:
                messagebox.showerror("Ошибка", str(e))
                return []
        return all_b

    def _sync_filter_widgets(self) -> None:
        if self.var_filter_mode.get() == "По тегу":
            self._row_filter_tag.grid()
            self._refresh_tag_filter_values()
        else:
            self._row_filter_tag.grid_remove()

    def _refresh_tag_filter_values(self) -> None:
        try:
            tags = db.list_distinct_tags()
        except Exception:
            tags = []
        vals = ["— выберите тег —"] + tags
        self.combo_filter_tag.configure(values=vals)
        cur = (self.var_filter_tag.get() or "").strip()
        if cur in vals:
            self.combo_filter_tag.set(cur)
        else:
            self.combo_filter_tag.set(vals[0])
            self.var_filter_tag.set(vals[0])

    def _on_filter_mode_change(self, _value: str | None = None) -> None:
        self._sync_filter_widgets()
        self._refresh_book_list(select_id=self._current_book_id)

    def _on_list_refresh(self) -> None:
        self._refresh_book_list(select_id=self._current_book_id)

    # Работа с карточкой книги
    def _new_book(self, clear_only: bool = False) -> None:
        self._current_book_id = None
        self.var_title.set("")
        self.var_author.set("")
        self.var_genre.set("")
        self.var_pages.set("0")
        self.var_started.set("")
        self.var_finished.set("")
        self.var_format.set("paper")
        self.var_plan_y.set("")
        self.var_plan_m.set("")
        self.var_reading_status.set("В планах")
        self.var_tags.set("")
        self._sync_plan_fields_visibility()
        self.txt_new_quote.delete("1.0", "end")
        self._render_quotes([])
        for sl in (self.slider_idea, self.slider_plot, self.slider_chars, self.slider_skill):
            sl.set(3)
        self._sync_rating_labels()
        self.txt_review.delete("1.0", "end")
        self._set_cover_preview(None)
        self._render_daily([])
        if hasattr(self, 'progress_bar'):
            self.progress_bar.set(0)
            self.lbl_progress.configure(text="Прочитано: 0 стр. (0%)")
        if not clear_only:
            messagebox.showinfo("Новая книга", "Заполните поля и нажмите «Сохранить».")

    def _sync_plan_fields_visibility(self, _value: str | None = None) -> None:
        if self.var_reading_status.get() == "В планах":
            self._frm_book_plan.grid()
        else:
            self._frm_book_plan.grid_remove()

    def _load_book(self, book_id: int) -> None:
        try:
            b = db.get_book(book_id)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
            return
        if not b:
            return

        self._current_book_id = book_id
        self.var_title.set(b.get("title") or "")
        self.var_author.set(b.get("author") or "")
        self.var_genre.set(b.get("genre") or "")
        self.var_pages.set(str(b.get("page_count") or 0))
        self.var_started.set((b.get("date_started") or "")[:10] if b.get("date_started") else "")
        self.var_finished.set((b.get("date_finished") or "")[:10] if b.get("date_finished") else "")

        fmt = b.get("format_type") or "paper"
        self.var_format.set(fmt if fmt in FORMAT_VALUES else "paper")

        self.var_plan_y.set(str(b.get("plan_year", "")) or "")
        self.var_plan_m.set(str(b.get("plan_month", "")) or "")

        st = b.get("reading_status") or "planned"
        self.var_reading_status.set(STATUS_KEY_TO_LABEL.get(st, "В планах"))
        self._sync_plan_fields_visibility()
        self._set_cover_preview(b.get("cover_path"))
        self.txt_new_quote.delete("1.0", "end")

        try:
            quotes = db.list_quotes(book_id)
            rev = db.get_review(book_id)
            daily = db.list_daily_for_book(book_id)
            tag_list = db.get_book_tags(book_id)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
            return

        self.var_tags.set(", ".join(tag_list))
        self._render_quotes(quotes)

        if rev:
            self.slider_idea.set(max(1, min(5, int(rev.get("rating_idea") or 3))))
            self.slider_plot.set(max(1, min(5, int(rev.get("rating_plot") or 3))))
            self.slider_chars.set(max(1, min(5, int(rev.get("rating_characters") or 3))))
            self.slider_skill.set(max(1, min(5, int(rev.get("rating_author_skill") or 3))))
            self.txt_review.delete("1.0", "end")
            rt = rev.get("review_text") or ""
            if rt:
                self.txt_review.insert("1.0", rt)
        else:
            for sl in (self.slider_idea, self.slider_plot, self.slider_chars, self.slider_skill):
                sl.set(3)
            self.txt_review.delete("1.0", "end")

        self._sync_rating_labels()
        self._render_daily(daily)
        self._set_tracker_combo_for_book(book_id)
        self._update_progress_bar(book_id)

    def _save_book(self) -> None:
        title = self.var_title.get().strip()
        if not title:
            messagebox.showwarning("Название", "Укажите название книги.")
            return

        pages = max(0, int_or(self.var_pages.get(), 0))
        ds = parse_date(self.var_started.get())
        df = parse_date(self.var_finished.get())
        fmt = self.var_format.get()
        if fmt not in FORMAT_VALUES:
            fmt = "paper"

        rstatus = STATUS_LABEL_TO_KEY.get(self.var_reading_status.get(), "planned")
        plan_year, plan_month = None, None
        if rstatus == "planned":
            py_raw = self.var_plan_y.get().strip()
            pm_raw = self.var_plan_m.get().strip()
            plan_year = int_or(py_raw, 0) if py_raw else None
            plan_month = max(1, min(12, int_or(pm_raw, 0))) if pm_raw else None

        try:
            if self._current_book_id is None:
                bid = db.create_book(
                    title=title, author=self.var_author.get().strip(), genre=self.var_genre.get().strip(),
                    page_count=pages, format_type=fmt, plan_year=plan_year, plan_month=plan_month,
                    date_started=ds, date_finished=df, reading_status=rstatus
                )
                self._current_book_id = bid
            else:
                db.update_book(
                    self._current_book_id, title=title, author=self.var_author.get().strip(),
                    genre=self.var_genre.get().strip(), page_count=pages, format_type=fmt,
                    plan_year=plan_year, plan_month=plan_month, date_started=ds, date_finished=df, reading_status=rstatus
                )

            idea = int(round(self.slider_idea.get()))
            plot = int(round(self.slider_plot.get()))
            chars = int(round(self.slider_chars.get()))
            skill = int(round(self.slider_skill.get()))
            rtext = self.txt_review.get("1.0", "end-1c").strip() or None
            db.upsert_review(self._current_book_id, idea, plot, chars, skill, rtext)

            raw_tags = [x.strip() for x in self.var_tags.get().split(",") if x.strip()]
            db.replace_book_tags(self._current_book_id, raw_tags)
        except Exception as e:
            messagebox.showerror("Ошибка сохранения", str(e))
            return

        self._refresh_book_list(select_id=self._current_book_id)
        messagebox.showinfo("Сохранено", "Книга и отзыв сохранены.")

    def _delete_book(self) -> None:
        if not self._current_book_id:
            messagebox.showinfo("Удаление", "Не выбрана книга.")
            return
        if not messagebox.askyesno("Удалить", "Удалить эту книгу и все связанные записи?"):
            return
        try:
            db.delete_book(self._current_book_id)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
            return
        self._current_book_id = None
        self._refresh_book_list(select_id=None)

    def _goto_book(self, book_id: int) -> None:
        self.var_filter_mode.set("Все книги")
        self._show_page("home")
        self._refresh_book_list(select_id=book_id)

    # Трекер чтения
    def _on_tracker_save(self) -> None:
        d = parse_date(self.var_track_date.get())
        if not d:
            messagebox.showwarning("Дата", "Введите дату в формате ГГГГ-ММ-ДД.")
            return
        sel = self.combo_track_book.get()
        ids = getattr(self, '_tracker_book_ids', [])
        titles = list(self.combo_track_book.cget('values'))
        idx = titles.index(sel) if sel in titles else -1
        if idx < 0 or idx >= len(ids):
            messagebox.showwarning("Книга", "Выберите книгу из списка.")
            return
        book_id = ids[idx]

        page_start = max(0, int_or(self.var_track_page_start.get(), 0))
        page_end = max(0, int_or(self.var_track_page_end.get(), 0))
        pages = (page_end - page_start + 1) if (page_start > 0 and page_end > 0 and page_end >= page_start) else 0
        note = self.var_track_note.get().strip() or None

        if pages == 0:
            messagebox.showwarning("Страницы", "Укажите корректный диапазон страниц.")
            return

        try:
            db.upsert_daily(book_id, d, pages, 0, note, page_start, page_end)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
            return

        self.var_track_page_start.set("0")
        self.var_track_page_end.set("0")
        self.var_track_note.set("")
        if self._current_book_id == book_id:
            self._render_daily(db.list_daily_for_book(book_id))
        messagebox.showinfo("Трекер", "Запись за день сохранена.")

    def _set_tracker_combo_for_book(self, book_id: int) -> None:
        ids = getattr(self, '_tracker_book_ids', [])
        titles = list(self.combo_track_book.cget('values'))
        if book_id in ids:
            self.combo_track_book.set(titles[ids.index(book_id)])

    def _edit_daily_row(self, row: dict) -> None:
        rd = row.get("read_date")
        ds = rd.isoformat()[:10] if hasattr(rd, "isoformat") else (str(rd)[:10] if rd else "")
        if not ds:
            return
        self.var_track_date.set(ds)

        page_start = row.get("page_start", 0)
        page_end = row.get("page_end", 0)
        if page_start == 0 and page_end == 0:
            self.var_track_page_start.set("0")
            self.var_track_page_end.set(str(row.get("pages_read", 0)))
        else:
            self.var_track_page_start.set(str(page_start))
            self.var_track_page_end.set(str(page_end))

        self.var_track_note.set(row.get("note") or "")
        bid = row.get("book_id")
        if bid is not None:
            self._set_tracker_combo_for_book(int(bid))

    def _delete_daily_row(self, row: dict) -> None:
        bid = row.get("book_id")
        d = parse_date(str(row.get("read_date")))
        if bid is None or not d:
            return
        if not messagebox.askyesno("Удалить", "Удалить запись трекера за этот день?"):
            return
        try:
            db.delete_daily_entry(int(bid), d)
            if self._current_book_id == int(bid):
                self._render_daily(db.list_daily_for_book(int(bid)))
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    # Интерфейс карточки (цитаты, отзывы, обложки, рейтинг)
    def _render_quotes(self, quotes: list) -> None:
        for w in self.frame_quotes.winfo_children():
            w.destroy()
        for q in quotes:
            fr = ctk.CTkFrame(self.frame_quotes)
            fr.pack(fill="x", pady=4)
            fr.grid_columnconfigure(0, weight=1)

            page = q.get("page_number")
            page_txt = f"  [стр. {page}]" if page else ""
            ctk.CTkLabel(
                fr, text=q.get("quote_text", "")[:500] + page_txt,
                wraplength=580, justify="left"
            ).grid(row=0, column=0, sticky="ew", padx=4, pady=4)

            btn_frame = ctk.CTkFrame(fr, fg_color="transparent")
            btn_frame.grid(row=0, column=1, sticky="ne", padx=4)
            ctk.CTkButton(
                btn_frame, text="✏️", width=32,
                fg_color=COLORS["primary"], hover_color=COLORS["secondary"],
                command=lambda q_=q: self._edit_quote(q_)
            ).pack(side="left", padx=(0, 2))
            ctk.CTkButton(
                btn_frame, text="❌", width=32,
                fg_color=COLORS["danger"], hover_color=("gray70", "gray50"),
                command=lambda i=q["id"]: self._del_quote(i)
            ).pack(side="left")

    def _del_quote(self, qid: int) -> None:
        if not self._current_book_id:
            return
        try:
            db.delete_quote(qid)
            self._load_book(self._current_book_id)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def _add_quote(self) -> None:
        if not self._current_book_id:
            messagebox.showinfo("Книга", "Сохраните книгу перед добавлением цитат.")
            return
        text = self.txt_new_quote.get("1.0", "end").strip()
        if not text:
            return
        try:
            db.add_quote(self._current_book_id, text)
            self.txt_new_quote.delete("1.0", "end")
            self._load_book(self._current_book_id)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def _edit_quote(self, q: dict) -> None:
        quote_id = q.get("id")
        if quote_id is None:
            messagebox.showerror("Ошибка", "Отсутствует ID цитаты.")
            return

        new_text = sd.askstring(
            "Редактирование цитаты",
            "Текст цитаты:",
            initialvalue=q.get("quote_text", "")
        )
        if new_text is None:
            return

        # Разрешаем редактирование только если текст не пустой
        if not new_text.strip():
            messagebox.showwarning("Внимание", "Текст цитаты не может быть пустым.")
            return

        # Безопасная обработка None для страницы (0 считается валидным номером)
        pg = q.get("page_number")
        initial_page = str(pg) if pg is not None else ""

        page_raw = sd.askstring(
            "Страница цитаты",
            "Номер страницы (оставьте пустым, если не знаете):",
            initialvalue=initial_page
        )
        if page_raw is None:
            return

        page = int(page_raw.strip()) if page_raw.strip().isdigit() else None

        try:
            db.update_quote(quote_id, new_text.strip(), page)
            self._load_book(self._current_book_id)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить изменения:\n{e}")


    def _render_daily(self, rows: list) -> None:
        for w in self.frame_daily.winfo_children():
            w.destroy()

        for row in rows[:40]:
            rd = row.get("read_date")
            date_str = str(rd) if rd is not None else "—"

            note = (row.get("note") or "").strip()
            note_txt = f" — {note[:48]}…" if len(note) > 48 else (f" — {note}" if note else "")

            ps = row.get("page_start", 0)
            pe = row.get("page_end", 0)
            # Защита от None в числовых полях
            ps = ps if isinstance(ps, (int, float)) else 0
            pe = pe if isinstance(pe, (int, float)) else 0
            
            pages_read = row.get("pages_read", 0) or 0
            pages_info = f"стр. {ps}–{pe}" if (ps > 0 and pe > 0) else f"{pages_read} стр."

            fr = ctk.CTkFrame(self.frame_daily)
            fr.pack(fill="x", pady=2, padx=5)

            ctk.CTkLabel(
                fr, 
                text=f"{date_str} — {pages_info}{note_txt}", 
                anchor="w"
            ).pack(side="left", padx=4, fill="x", expand=True)

            ctk.CTkButton(
                fr, text="✏️ Изменить", width=80,
                fg_color=COLORS.get("primary", "#1f6aa5"),
                hover_color=COLORS.get("secondary", "#144870"),
                command=lambda r=row: self._edit_daily_row(r)
            ).pack(side="right", padx=2)

            ctk.CTkButton(
                fr, text="❌ Удалить", width=72,
                fg_color=COLORS.get("danger", "#d32f2f"),
                hover_color=COLORS.get("danger_hover", "#b71c1c"),
                command=lambda r=row: self._delete_daily_row(r)
            ).pack(side="right", padx=2)
    def _wire_rating_slider(self, sl: ctk.CTkSlider, lbl: ctk.CTkLabel, prefix: str) -> None:
        def upd(v=None):
            lbl.configure(text=f"{prefix}: {int(round(float(sl.get())))}")
        sl.configure(command=lambda _: upd())

    def _sync_rating_labels(self) -> None:
        self._wire_rating_slider(self.slider_idea, self.lbl_idea, "Идея")
        self._wire_rating_slider(self.slider_plot, self.lbl_plot, "Сюжет")
        self._wire_rating_slider(self.slider_chars, self.lbl_chars, "Персонажи")
        self._wire_rating_slider(self.slider_skill, self.lbl_skill, "Мастерство автора")
        # Принудительно обновляем текст лейблов по текущим значениям слайдеров
        self.lbl_idea.configure(text=f"Идея: {int(round(self.slider_idea.get()))}")
        self.lbl_plot.configure(text=f"Сюжет: {int(round(self.slider_plot.get()))}")
        self.lbl_chars.configure(text=f"Персонажи: {int(round(self.slider_chars.get()))}")
        self.lbl_skill.configure(text=f"Мастерство автора: {int(round(self.slider_skill.get()))}")

    def _update_progress_bar(self, book_id: int) -> None:
        if not hasattr(self, 'progress_bar'):
            return
        try:
            b = db.get_book(book_id)
            total = int(b.get('page_count') or 0) if b else 0
            logged = db.total_pages_logged_for_book(book_id)
        except Exception:
            return
        if total > 0:
            pct = min(1.0, logged / total)
            self.progress_bar.set(pct)
            self.lbl_progress.configure(text=f"Прочитано: {logged} стр. из {total} ({int(pct * 100)}%)")
        else:
            self.progress_bar.set(0)
            self.lbl_progress.configure(text=f"Прочитано: {logged} стр. (объём книги не указан)")

    def _export_review(self) -> None:
        if not self._current_book_id:
            messagebox.showinfo("Экспорт", "Сначала выберите книгу.")
            return
        import dialogs
        dialogs.export_review_dialog(self._current_book_id)

    def _set_cover_preview(self, path: str | None) -> None:
        self._cover_photo = None
        self._cover_pil_image = None
        self._cover_pil_image_original = None

        if not path:
            try:
                self.lbl_cover.configure(image="", text="Нет обложки")
                self.lbl_cover.image = None
            except Exception:
                pass
            return
        try:
            path_obj = Path(path)
            if not path_obj.is_file():
                self.lbl_cover.configure(image="", text="Файл не найден")
                self.lbl_cover.image = None
                return

            original = Image.open(path_obj)
            self._cover_pil_image_original = original
            converted = original.convert("RGBA")
            converted.thumbnail((200, 280), Image.Resampling.LANCZOS)
            self._cover_pil_image = converted

            self._cover_photo = ctk.CTkImage(light_image=converted, dark_image=converted, size=converted.size)
            self.lbl_cover.configure(image=self._cover_photo, text="")
            self.lbl_cover.image = self._cover_photo  # защита от GC
        except Exception as e:
            print(f"Error loading cover: {e}")
            self.lbl_cover.configure(image="", text="Ошибка загрузки")
            self.lbl_cover.image = None

    def _pick_cover(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Изображения", "*.png *.jpg *.jpeg *.gif *.webp"), ("Все", "*.*")])
        if not path:
            return
        if not self._current_book_id:
            messagebox.showinfo("Сначала сохраните книгу", "Нажмите «Сохранить», затем добавьте обложку.")
            return
        try:
            saved = db.save_cover_from_path(self._current_book_id, path)
            if saved:
                db.update_book(self._current_book_id, cover_path=saved)
                self._set_cover_preview(saved)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    # Аналитика, цели, графики, поиск
    def _stats_period_range(self) -> tuple[date, date]:
        t = today()
        mode = self.var_stats_period.get()
        if mode == "Неделя":
            start = t - timedelta(days=t.weekday())
            return start, start + timedelta(days=6)
        y = int_or(self.var_plan_year.get(), t.year)
        m = max(1, min(12, int_or(self.var_plan_month.get(), t.month)))
        if mode == "Месяц":
            return date(y, m, 1), date(y, m, calendar.monthrange(y, m)[1])
        return date(y, 1, 1), date(y, 12, 31)

    def _refresh_stats_page(self) -> None:
        stats.refresh_stats(self)

    def _refresh_pace_only(self) -> None:
        stats.refresh_pace(self)

    def _on_save_reading_goal(self) -> None:
        period = self.var_stats_period.get()
        y = int_or(self.var_plan_year.get(), today().year)
        m = max(1, min(12, int_or(self.var_plan_month.get(), today().month)))
        if period == "Месяц":
            sm = m
        elif period == "Год":
            sm = 0
        else:
            messagebox.showinfo("Цели", "Для недели цели не задаются. Выберите «Месяц» или «Год».")
            return
        tb = max(0, int_or(self.var_goal_books_s.get(), 0))
        try:
            db.upsert_reading_goal(y, sm, tb)
            self._refresh_stats_page()
            messagebox.showinfo("Сохранено", "Цель сохранена.")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def _refresh_charts_page(self) -> None:
        stats.refresh_charts(self)

    def _run_search(self) -> None:
        stats.run_search(self)

    def _refresh_plan_month_list(self) -> None:
        plans.refresh_month(self)

    def _refresh_plan_year_list(self) -> None:
        plans.refresh_year(self)

    def _refresh_status_list(self, status_key: str, scroll: ctk.CTkScrollableFrame) -> None:
        plans.refresh_status(self, status_key, scroll)

    # Настройки и горячие клавиши
    def _load_settings(self) -> None:
        try:
            if SETTINGS_PATH.is_file():
                disk = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
                if isinstance(disk, dict):
                    self._app_settings.update(disk)
        except Exception:
            pass
        mode = self._app_settings.get("appearance", "dark")
        if mode in self._theme_inv:
            self.var_theme.set(self._theme_inv[mode])
        ctk.set_appearance_mode(mode)
        try:
            ctk.set_widget_scaling(float(self._app_settings.get("ui_scale", 1.0)))
        except Exception:
            ctk.set_widget_scaling(1.0)

    def _save_settings(self) -> None:
        try:
            self._app_settings["appearance"] = self._theme_map.get(self.var_theme.get(), "dark")
            SETTINGS_PATH.parent.mkdir(parents=True, exist_ok=True)
            SETTINGS_PATH.write_text(json.dumps(self._app_settings, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    def _on_theme_change(self, _value: str | None = None) -> None:
        mode = self._theme_map.get(self.var_theme.get(), "dark")
        ctk.set_appearance_mode(mode)
        self._save_settings()

    def _apply_app_preferences(self) -> None:
        try:
            self._app_settings["ui_scale"] = float(self.var_ui_scale_menu.get())
        except ValueError:
            self._app_settings["ui_scale"] = 1.0
        self._app_settings["compact_list"] = bool(self.var_compact_chk.get())
        self._app_settings["reminder_enabled"] = bool(self.var_reminder_on.get())
        self._app_settings["reminder_time"] = (self.var_reminder_time.get() or "20:00").strip()[:5]
        try:
            ctk.set_widget_scaling(float(self._app_settings["ui_scale"]))
        except Exception:
            pass
        self._save_settings()
        self._refresh_book_list(select_id=self._current_book_id)
        messagebox.showinfo("Настройки", "Параметры сохранены.")

    def _export_backup_dialog(self) -> None:
        dialogs.export_backup_dialog()

    def _import_backup_dialog(self) -> None:
        success, _ = dialogs.import_backup_dialog()
        if success:
            self._current_book_id = None
            self._show_page("home")
            self._refresh_book_list(select_id=None)

    def _export_excel_dialog(self) -> None:
        dialogs.export_excel_dialog()

    def _setup_hotkeys(self) -> None:
        def bind(seq, fn):
            def h(ev):
                fn()
                return "break"
            self.bind_all(seq, h)

        bind("<Control-Shift-s>", self._hotkey_save)
        bind("<Control-Shift-n>", self._hotkey_new)
        bind("<Control-Shift-f>", self._hotkey_search)
        bind("<Control-Shift-t>", self._hotkey_tracker_focus)

    def _hotkey_save(self) -> None:
        if self._current_page_key == "home":
            self._save_book()

    def _hotkey_new(self) -> None:
        if self._current_page_key == "home":
            self._new_book()

    def _hotkey_search(self) -> None:
        self._show_page("search")

    def _hotkey_tracker_focus(self) -> None:
        self._show_page("home")
        try:
            self.lift()
            self.focus_force()
            if hasattr(self, "var_track_date"):
                self.focus()
        except Exception:
            pass

    def _reminder_tick(self) -> None:
        try:
            if self._app_settings.get("reminder_enabled"):
                parts = (self._app_settings.get("reminder_time") or "20:00").replace(" ", ":").split(":")
                th = int(parts[0]) if parts else 20
                tm = int(parts[1]) if len(parts) > 1 else 0
                now = datetime.now()
                today_s = now.date().isoformat()
                if self._app_settings.get("last_reminder_date") != today_s:
                    if now.hour > th or (now.hour == th and now.minute >= tm):
                        self._app_settings["last_reminder_date"] = today_s
                        self._save_settings()
                        try:
                            from plyer import notification
                            notification.notify(
                                title="📚 Читательский дневник",
                                message="Не забудьте отметить чтение в трекере.",
                                app_name="Читательский дневник",
                                timeout=8,
                            )
                        except Exception:
                            messagebox.showinfo("Напоминание", "Не забудьте отметить чтение в трекере.")
        except Exception:
            pass
        self.after(60_000, self._reminder_tick)


def _start_mysql_service() -> bool:
    """Пытается запустить службу MySQL. Возвращает True если удалось подключиться."""
    import subprocess
    import time
    # Перебираем возможные имена службы
    for svc in ("MySQL_ChitatelDnevnik", "MySQL80", "MySQL", "MySQL57"):
        try:
            subprocess.run(["net", "start", svc], capture_output=True, timeout=30)
        except Exception:
            pass
    # Ждём до 15 секунд пока MySQL поднимется
    for _ in range(15):
        try:
            db.connect().close()
            return True
        except Exception:
            time.sleep(1)
    return False


def main() -> None:
    try:
        db.connect().close()
    except Exception:
        import tkinter as tk
        from tkinter import messagebox as mb
        root = tk.Tk()
        root.withdraw()
        mb.showinfo(
            "Запуск базы данных",
            "MySQL не запущен. Выполняется автоматический запуск службы...\n"
            "Пожалуйста, подождите."
        )
        root.destroy()
        if not _start_mysql_service():
            import tkinter as tk
            from tkinter import messagebox as mb
            root = tk.Tk()
            root.withdraw()
            mb.showerror(
                "Ошибка подключения",
                "Не удалось подключиться к MySQL.\n\n"
                "Убедитесь что:\n"
                "• Служба MySQL запущена\n"
                "• Параметры в файле .env верны\n"
                "• Файл .env находится рядом с приложением"
            )
            root.destroy()
            return
    try:
        app = ReadingDiaryApp()
        app.mainloop()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()

