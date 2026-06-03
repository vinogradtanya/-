"""Главная страница: трекер чтения, список книг, карточка книги."""
from __future__ import annotations

from pathlib import Path
from tkinter import messagebox, filedialog

import customtkinter as ctk
from PIL import Image

import db
from constants import COLORS, FORMAT_VALUES, FORMAT_LABELS, STATUS_LABEL_TO_KEY, STATUS_KEY_TO_LABEL
from utils import today, parse_date, int_or


def build(app) -> ctk.CTkFrame:
    home = ctk.CTkFrame(app._content_host, fg_color="transparent")
    home.grid_columnconfigure(0, weight=1)
    home.grid_rowconfigure(1, weight=1)
    _build_tracker(app, home)
    _build_main_area(app, home)
    return home


def _build_tracker(app, parent: ctk.CTkFrame) -> None:
    tracker = ctk.CTkFrame(parent)
    tracker.grid(row=0, column=0, sticky="ew", pady=(0, 8))
    tracker.grid_columnconfigure(10, weight=1)

    ctk.CTkLabel(tracker, text="Трекер чтения").grid(
        row=0, column=0, columnspan=10, padx=12, pady=(8, 4), sticky="w"
    )
    ctk.CTkLabel(tracker, text="Дата").grid(row=1, column=0, padx=(12, 4), pady=8)
    app.var_track_date = ctk.StringVar(value=today().isoformat())
    ctk.CTkEntry(tracker, textvariable=app.var_track_date, width=110).grid(row=1, column=1, padx=4, pady=8)

    ctk.CTkLabel(tracker, text="Книга").grid(row=1, column=2, padx=(12, 4), pady=8)
    app.combo_track_book = ctk.CTkComboBox(tracker, width=260, values=[])
    app.combo_track_book.grid(row=1, column=3, padx=4, pady=8)

    ctk.CTkLabel(tracker, text="От страниц").grid(row=1, column=4, padx=(8, 4), pady=8)
    app.var_track_page_start = ctk.StringVar(value="0")
    ctk.CTkEntry(tracker, textvariable=app.var_track_page_start, width=56).grid(row=1, column=5, padx=4, pady=8)

    ctk.CTkLabel(tracker, text="До страниц").grid(row=1, column=6, padx=(8, 4), pady=8)
    app.var_track_page_end = ctk.StringVar(value="0")
    ctk.CTkEntry(tracker, textvariable=app.var_track_page_end, width=56).grid(row=1, column=7, padx=4, pady=8)

    ctk.CTkButton(
        tracker, text="📝 Записать", width=100, command=app._on_tracker_save,
        fg_color=COLORS["success"], hover_color="#059669"
    ).grid(row=1, column=8, padx=8, pady=8)

    app.var_track_note = ctk.StringVar()
    ctk.CTkEntry(tracker, textvariable=app.var_track_note, placeholder_text="Заметка…", width=200).grid(
        row=1, column=9, padx=4, pady=8, sticky="w"
    )


def _build_main_area(app, parent: ctk.CTkFrame) -> None:
    main = ctk.CTkFrame(parent, fg_color="transparent")
    main.grid(row=1, column=0, sticky="nsew")
    main.grid_columnconfigure(1, weight=1)
    main.grid_rowconfigure(0, weight=1)
    _build_left_panel(app, main)
    _build_right_panel(app, main)


def _build_left_panel(app, parent: ctk.CTkFrame) -> None:
    left = ctk.CTkFrame(parent, width=300)
    left.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
    left.grid_propagate(False)
    left.grid_columnconfigure(0, weight=1)
    left.grid_rowconfigure(3, weight=1)

    ctk.CTkLabel(left, text="Книги").grid(row=0, column=0, padx=8, pady=(8, 4), sticky="w")

    fil = ctk.CTkFrame(left, fg_color="transparent")
    fil.grid(row=1, column=0, sticky="ew", padx=8, pady=4)
    fil.grid_columnconfigure(1, weight=1)

    app.var_filter_mode = ctk.StringVar(value="Все книги")
    ctk.CTkOptionMenu(
        fil, values=["Все книги", "План на месяц", "По тегу"],
        variable=app.var_filter_mode, command=app._on_filter_mode_change,
    ).grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 4))

    app._row_filter_tag = ctk.CTkFrame(fil, fg_color="transparent")
    app._row_filter_tag.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 4))
    app._row_filter_tag.grid_columnconfigure(1, weight=1)
    ctk.CTkLabel(app._row_filter_tag, text="Тег").grid(row=0, column=0, sticky="w")
    app.var_filter_tag = ctk.StringVar(value="")
    app.combo_filter_tag = ctk.CTkComboBox(
        app._row_filter_tag, values=["— теги —"], variable=app.var_filter_tag, width=200,
        command=lambda _: app._refresh_book_list(select_id=app._current_book_id),
    )
    app.combo_filter_tag.grid(row=0, column=1, sticky="ew", padx=(6, 0))

    cy, cm = today().year, today().month
    app.var_plan_year = ctk.StringVar(value=str(cy))
    app.var_plan_month = ctk.StringVar(value=str(cm))
    ctk.CTkLabel(fil, text="Год").grid(row=2, column=0, sticky="w")
    ctk.CTkEntry(fil, textvariable=app.var_plan_year, width=72).grid(row=2, column=1, sticky="w", padx=(4, 0))
    ctk.CTkLabel(fil, text="Месяц").grid(row=3, column=0, sticky="w", pady=2)
    ctk.CTkEntry(fil, textvariable=app.var_plan_month, width=48).grid(row=3, column=1, sticky="w", padx=(4, 0), pady=2)

    app._row_filter_tag.grid_remove()
    app._sync_filter_widgets()

    btns = ctk.CTkFrame(left, fg_color="transparent")
    btns.grid(row=2, column=0, sticky="ew", padx=8, pady=8)
    ctk.CTkButton(btns, text="🔄 Обновить", width=80, command=app._on_list_refresh,
                  fg_color=COLORS["muted"], hover_color=COLORS["primary"]).pack(side="left", padx=(0, 6))
    ctk.CTkButton(btns, text="➕ Новая", width=80, command=app._new_book,
                  fg_color=COLORS["success"], hover_color=COLORS["primary"]).pack(side="left")

    app.list_books = ctk.CTkScrollableFrame(left, label_text="")
    app.list_books.grid(row=3, column=0, sticky="nsew", padx=8, pady=8)

    delbar = ctk.CTkFrame(left, fg_color="transparent")
    delbar.grid(row=4, column=0, sticky="ew", padx=8, pady=(0, 8))
    ctk.CTkButton(delbar, text="🗑️ Удалить книгу", fg_color=COLORS["danger"],
                  hover_color="#B91C1C", command=app._delete_book, text_color="white").pack(fill="x")


def _build_right_panel(app, parent: ctk.CTkFrame) -> None:
    right = ctk.CTkScrollableFrame(parent)
    right.grid(row=0, column=1, sticky="nsew")
    app._panel = right
    build_detail_panel(app, right)


def _insert_template(app, text: str) -> None:
    """Вставляет шаблон в поле отзыва."""
    app.txt_review.delete("1.0", "end")
    app.txt_review.insert("1.0", text)


def build_detail_panel(app, parent: ctk.CTkScrollableFrame) -> None:
    r = 0
    ctk.CTkLabel(parent, text="Карточка книги", font=ctk.CTkFont(size=18, weight="bold")).grid(
        row=r, column=0, columnspan=4, sticky="w", pady=(0, 12)
    )
    r += 1

    app.lbl_cover = ctk.CTkLabel(parent, text="Нет обложки", width=200, height=280)
    app.lbl_cover.grid(row=r, column=0, rowspan=6, sticky="nw", padx=(0, 16), pady=4)
    ctk.CTkButton(parent, text="🖼️ Выбрать фото", command=app._pick_cover,
                  fg_color=COLORS["primary"], hover_color=COLORS["secondary"]).grid(
        row=r + 6, column=0, sticky="w", pady=(8, 0)
    )

    def add_row(label: str, widget, col_offset: int = 1):
        nonlocal r
        ctk.CTkLabel(parent, text=label).grid(row=r, column=col_offset, sticky="ne", padx=(0, 8), pady=4)
        widget.grid(row=r, column=col_offset + 1, columnspan=2, sticky="ew", pady=4)
        parent.grid_columnconfigure(col_offset + 1, weight=1)
        r += 1

    app.var_title = ctk.StringVar()
    app.var_author = ctk.StringVar()
    app.var_genre = ctk.StringVar()
    app.var_pages = ctk.StringVar(value="0")
    app.var_started = ctk.StringVar()
    app.var_finished = ctk.StringVar()
    app.var_plan_y = ctk.StringVar()
    app.var_plan_m = ctk.StringVar()

    add_row("Название", ctk.CTkEntry(parent, textvariable=app.var_title))
    add_row("Автор", ctk.CTkEntry(parent, textvariable=app.var_author))
    add_row("Жанр", ctk.CTkEntry(parent, textvariable=app.var_genre))

    app.var_tags = ctk.StringVar()
    add_row("Теги (через запятую)", ctk.CTkEntry(parent, textvariable=app.var_tags))

    app.var_reading_status = ctk.StringVar(value="В планах")
    app.combo_reading_status = ctk.CTkOptionMenu(
        parent, values=["В планах", "Читаю", "Прочитано", "Брошено"],
        variable=app.var_reading_status, command=app._sync_plan_fields_visibility, width=200,
    )
    add_row("Статус чтения", app.combo_reading_status)
    add_row("Страниц", ctk.CTkEntry(parent, textvariable=app.var_pages, width=80))

    app.var_format = ctk.StringVar(value="paper")
    fmt_frame = ctk.CTkFrame(parent, fg_color="transparent")
    for key in FORMAT_VALUES:
        ctk.CTkRadioButton(fmt_frame, text=FORMAT_LABELS[key], variable=app.var_format, value=key).pack(
            side="left", padx=(0, 12)
        )
    add_row("Формат", fmt_frame)
    add_row("Начало (ГГГГ-ММ-ДД)", ctk.CTkEntry(parent, textvariable=app.var_started))
    add_row("Конец (ГГГГ-ММ-ДД)", ctk.CTkEntry(parent, textvariable=app.var_finished))

    app._frm_book_plan = ctk.CTkFrame(parent, fg_color="transparent")
    app._frm_book_plan.grid_columnconfigure(1, weight=1)
    ctk.CTkLabel(
        app._frm_book_plan,
        text="Год и месяц, в которые книга попадает в разделы «План на месяц» и «План на год». "
             "Заполняйте только для книг со статусом «В планах».",
        wraplength=520, justify="left", font=ctk.CTkFont(size=12),
    ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))
    ctk.CTkLabel(app._frm_book_plan, text="В плане: год").grid(row=1, column=0, sticky="ne", padx=(0, 8), pady=4)
    ctk.CTkEntry(app._frm_book_plan, textvariable=app.var_plan_y, width=80).grid(row=1, column=1, sticky="w", pady=4)
    ctk.CTkLabel(app._frm_book_plan, text="В плане: месяц (1–12)").grid(row=2, column=0, sticky="ne", padx=(0, 8), pady=4)
    ctk.CTkEntry(app._frm_book_plan, textvariable=app.var_plan_m, width=80).grid(row=2, column=1, sticky="w", pady=4)
    app._frm_book_plan.grid(row=r, column=0, columnspan=4, sticky="ew", pady=4)
    r += 1
    app._sync_plan_fields_visibility()

    # Прогресс-бар прочитанного
    app.lbl_progress = ctk.CTkLabel(parent, text="Прочитано: 0 стр. (0%)", font=ctk.CTkFont(size=12))
    app.lbl_progress.grid(row=r, column=0, columnspan=4, sticky="w", pady=(8, 2))
    r += 1
    app.progress_bar = ctk.CTkProgressBar(parent, height=14)
    app.progress_bar.set(0)
    app.progress_bar.grid(row=r, column=0, columnspan=4, sticky="ew", pady=(0, 8))
    r += 1

    ctk.CTkLabel(parent, text="Цитаты", font=ctk.CTkFont(weight="bold")).grid(
        row=r, column=0, columnspan=4, sticky="w", pady=(8, 8)
    )
    r += 1
    qf = ctk.CTkFrame(parent, fg_color="transparent")
    qf.grid(row=r, column=0, columnspan=4, sticky="ew")
    qf.grid_columnconfigure(0, weight=1)
    app.txt_new_quote = ctk.CTkTextbox(qf, height=60)
    app.txt_new_quote.grid(row=0, column=0, sticky="ew", padx=(0, 8))
    page_frame = ctk.CTkFrame(qf, fg_color="transparent")
    page_frame.grid(row=1, column=0, sticky="w", pady=(4, 0))
    ctk.CTkLabel(page_frame, text="Страница:").pack(side="left", padx=(0, 6))
    app.var_quote_page = ctk.StringVar()
    ctk.CTkEntry(page_frame, textvariable=app.var_quote_page, width=72,
                 placeholder_text="необяз.").pack(side="left")
    ctk.CTkButton(qf, text="➕ Добавить", command=app._add_quote,
                  fg_color=COLORS["primary"], hover_color=COLORS["secondary"], text_color="white").grid(
        row=0, column=1, rowspan=2, sticky="ne"
    )
    r += 1
    app.frame_quotes = ctk.CTkScrollableFrame(parent, height=140)
    app.frame_quotes.grid(row=r, column=0, columnspan=4, sticky="ew", pady=8)
    r += 1

    ctk.CTkLabel(parent, text="Отзыв", font=ctk.CTkFont(weight="bold")).grid(
        row=r, column=0, columnspan=4, sticky="w", pady=(8, 4)
    )
    r += 1

    # Шаблоны отзывов
    tpl_frame = ctk.CTkFrame(parent, fg_color="transparent")
    tpl_frame.grid(row=r, column=0, columnspan=4, sticky="ew", pady=(0, 6))
    ctk.CTkLabel(tpl_frame, text="Шаблон:", font=ctk.CTkFont(size=12)).pack(side="left", padx=(0, 6))
    _TEMPLATES = [
        ("Краткий", "Что понравилось?\nЧто не понравилось?\nРекомендую ли?"),
        ("Глубокий", "О чём эта книга?\nЧто автор хотел сказать?\nКакие мысли или чувства вызвала?\nЧто запомнилось больше всего?\nРекомендую ли и кому?"),
        ("Аналитика", "Сильные стороны:\nСлабые стороны:\nКлючевые идеи:\nЛюбимый персонаж:\nОценка стиля автора:"),
    ]
    for name, text in _TEMPLATES:
        ctk.CTkButton(
            tpl_frame, text=name, width=80, height=24,
            fg_color=COLORS["muted"], hover_color=COLORS["primary"],
            font=ctk.CTkFont(size=11),
            command=lambda t=text: _insert_template(app, t),
        ).pack(side="left", padx=3)
    r += 1

    app.slider_idea = ctk.CTkSlider(parent, from_=1, to=5, number_of_steps=4)
    app.lbl_idea = ctk.CTkLabel(parent, text="Идея: 3")
    app._wire_rating_slider(app.slider_idea, app.lbl_idea, "Идея")
    app.slider_plot = ctk.CTkSlider(parent, from_=1, to=5, number_of_steps=4)
    app.lbl_plot = ctk.CTkLabel(parent, text="Сюжет: 3")
    app._wire_rating_slider(app.slider_plot, app.lbl_plot, "Сюжет")
    app.slider_chars = ctk.CTkSlider(parent, from_=1, to=5, number_of_steps=4)
    app.lbl_chars = ctk.CTkLabel(parent, text="Персонажи: 3")
    app._wire_rating_slider(app.slider_chars, app.lbl_chars, "Персонажи")
    app.slider_skill = ctk.CTkSlider(parent, from_=1, to=5, number_of_steps=4)
    app.lbl_skill = ctk.CTkLabel(parent, text="Мастерство автора: 3")
    app._wire_rating_slider(app.slider_skill, app.lbl_skill, "Мастерство автора")

    for i, (sl, lb) in enumerate([
        (app.slider_idea, app.lbl_idea), (app.slider_plot, app.lbl_plot),
        (app.slider_chars, app.lbl_chars), (app.slider_skill, app.lbl_skill),
    ]):
        sl.set(3)
        sl.grid(row=r + i, column=1, sticky="ew", pady=4)
        lb.grid(row=r + i, column=0, sticky="e", padx=(0, 8))
    r += 4

    ctk.CTkLabel(parent, text="Текст отзыва").grid(row=r, column=0, sticky="ne", padx=(0, 8), pady=8)
    app.txt_review = ctk.CTkTextbox(parent, height=100)
    app.txt_review.grid(row=r, column=1, columnspan=3, sticky="ew", pady=8)
    r += 1

    sf = ctk.CTkFrame(parent, fg_color="transparent")
    sf.grid(row=r, column=0, columnspan=4, pady=16, sticky="ew")
    ctk.CTkButton(sf, text="💾 Сохранить книгу", command=app._save_book, height=36,
                  fg_color=COLORS["success"], hover_color="#059669").pack(side="left", padx=(0, 8))
    ctk.CTkButton(sf, text="🔄 Сбросить", command=app._new_book,
                  fg_color=COLORS["muted"], hover_color=COLORS["primary"]).pack(side="left", padx=(0, 8))
    ctk.CTkButton(sf, text="📄 Экспорт отзыва", command=app._export_review,
                  fg_color=COLORS["secondary"], hover_color=COLORS["primary"]).pack(side="left")
    r += 1

    ctk.CTkLabel(parent, text="История чтения (выбранная книга)", font=ctk.CTkFont(weight="bold")).grid(
        row=r, column=0, columnspan=4, sticky="w", pady=(16, 8)
    )
    r += 1
    app.frame_daily = ctk.CTkScrollableFrame(parent, height=120)
    app.frame_daily.grid(row=r, column=0, columnspan=4, sticky="ew")
