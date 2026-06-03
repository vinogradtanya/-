"""Страницы: аналитика и цели, графики, поиск."""
from __future__ import annotations

import calendar
from datetime import date, timedelta
from tkinter import messagebox

import customtkinter as ctk

import db
from charts import create_charts_frame
from utils import today, int_or


# Аналитика

def build_stats(app) -> ctk.CTkFrame:
    f = ctk.CTkFrame(app._content_host, fg_color="transparent")
    f.grid_columnconfigure(0, weight=1)
    f.grid_rowconfigure(7, weight=1)
    ctk.CTkLabel(f, text="Аналитика и цели", font=ctk.CTkFont(size=18, weight="bold")).grid(
        row=0, column=0, sticky="w", pady=(0, 8)
    )
    bar = ctk.CTkFrame(f, fg_color="transparent")
    bar.grid(row=1, column=0, sticky="ew", pady=(0, 6))
    app.var_stats_period = ctk.StringVar(value="Месяц")
    ctk.CTkLabel(bar, text="Период").pack(side="left", padx=(0, 8))
    ctk.CTkOptionMenu(
        bar, values=["Неделя", "Месяц", "Год"], variable=app.var_stats_period,
        command=lambda _: app._refresh_stats_page(), width=110,
    ).pack(side="left", padx=4)
    ctk.CTkLabel(bar, text="(год и месяц — как в фильтре на главной)").pack(side="left", padx=12)
    ctk.CTkButton(bar, text="Обновить", width=100, command=app._refresh_stats_page).pack(side="left", padx=12)

    app.lbl_stat_sum = ctk.CTkLabel(f, text="")
    app.lbl_stat_sum.grid(row=2, column=0, sticky="w", pady=4)

    app._frm_stat_goals = ctk.CTkFrame(f, fg_color="transparent")
    app._frm_stat_goals.grid(row=3, column=0, sticky="ew", pady=6)
    app._frm_stat_goals.grid_columnconfigure(3, weight=1)
    ctk.CTkLabel(app._frm_stat_goals, text="Цель по книгам на период", font=ctk.CTkFont(weight="bold")).grid(
        row=0, column=0, columnspan=4, sticky="w", pady=(0, 6)
    )
    app.var_goal_books_s = ctk.StringVar(value="0")
    ctk.CTkLabel(app._frm_stat_goals, text="Цель, книг:").grid(row=1, column=0, padx=(0, 4))
    ctk.CTkEntry(app._frm_stat_goals, textvariable=app.var_goal_books_s, width=72).grid(row=1, column=1, padx=4)
    ctk.CTkButton(app._frm_stat_goals, text="Сохранить", command=app._on_save_reading_goal).grid(row=1, column=2, padx=12)
    app.prog_goal_books = ctk.CTkProgressBar(app._frm_stat_goals, width=420)
    app.prog_goal_books.set(0)
    app.prog_goal_books.grid(row=2, column=0, columnspan=4, sticky="ew", pady=(8, 4))
    app.lbl_goal_books_txt = ctk.CTkLabel(app._frm_stat_goals, text="")
    app.lbl_goal_books_txt.grid(row=3, column=0, columnspan=4, sticky="w")

    pace_fr = ctk.CTkFrame(f, fg_color="transparent")
    pace_fr.grid(row=4, column=0, sticky="ew", pady=(12, 4))
    ctk.CTkLabel(pace_fr, text="Темп по книге").pack(side="left", padx=(0, 8))
    app.combo_pace_book = ctk.CTkComboBox(pace_fr, values=[], width=340, command=lambda _: app._refresh_pace_only())
    app.combo_pace_book.pack(side="left", padx=4)
    ctk.CTkButton(pace_fr, text="Пересчитать", width=100, command=app._refresh_pace_only).pack(side="left", padx=8)

    app.lbl_pace = ctk.CTkLabel(f, text="", wraplength=760, justify="left", anchor="w")
    app.lbl_pace.grid(row=5, column=0, sticky="ew", pady=4)

    ctk.CTkLabel(f, text="Сводка по дням", font=ctk.CTkFont(weight="bold")).grid(
        row=6, column=0, sticky="w", pady=(12, 4)
    )
    app.stats_breakdown = ctk.CTkScrollableFrame(f, height=260)
    app.stats_breakdown.grid(row=7, column=0, sticky="nsew")
    return f


def stats_period_range(app) -> tuple[date, date]:
    t = today()
    mode = app.var_stats_period.get()
    if mode == "Неделя":
        start = t - timedelta(days=t.weekday())
        return start, start + timedelta(days=6)
    y = int_or(app.var_plan_year.get(), t.year)
    m = max(1, min(12, int_or(app.var_plan_month.get(), t.month)))
    if mode == "Месяц":
        return date(y, m, 1), date(y, m, calendar.monthrange(y, m)[1])
    return date(y, 1, 1), date(y, 12, 31)


def refresh_stats(app) -> None:
    d0, d1 = stats_period_range(app)
    try:
        pgs, _ = db.sum_daily_between(d0, d1)
        finished = db.count_finished_books_between(d0, d1)
        rows = db.daily_breakdown_between(d0, d1)
        all_b = db.list_all_books()
    except Exception as e:
        messagebox.showerror("Ошибка", str(e))
        return

    app.lbl_stat_sum.configure(
        text=f"За период: прочитано {finished} кн., {pgs} стр.   ({d0.isoformat()} — {d1.isoformat()})"
    )

    period = app.var_stats_period.get()
    tg_b = 0
    y = int_or(app.var_plan_year.get(), today().year)
    m = max(1, min(12, int_or(app.var_plan_month.get(), today().month)))
    if period == "Месяц":
        g = db.get_reading_goal(y, m)
        tg_b = int(g.get("target_pages") or 0) if g else 0
        app.var_goal_books_s.set(str(tg_b))
        app._frm_stat_goals.grid()
    elif period == "Год":
        g = db.get_reading_goal(y, 0)
        tg_b = int(g.get("target_pages") or 0) if g else 0
        app.var_goal_books_s.set(str(tg_b))
        app._frm_stat_goals.grid()
    else:
        app._frm_stat_goals.grid_remove()

    if period != "Неделя":
        app.prog_goal_books.set(0 if tg_b <= 0 else min(1.0, finished / tg_b))
        app.lbl_goal_books_txt.configure(
            text=f"Книги: {finished} / {tg_b}" if tg_b else "Цель по книгам не задана."
        )

    titles_pb = [(b.get('title') or '')[:40] for b in all_b]
    app._pace_book_ids = [b['id'] for b in all_b]
    app.combo_pace_book.configure(values=titles_pb or ["— нет книг —"])
    if titles_pb and app.combo_pace_book.get() not in titles_pb:
        app.combo_pace_book.set(titles_pb[0])
    app._refresh_pace_only()

    for w in app.stats_breakdown.winfo_children():
        w.destroy()
    if not rows:
        ctk.CTkLabel(app.stats_breakdown, text="Нет записей трекера за этот период.").pack(anchor="w", pady=4)
    else:
        for r in rows[:90]:
            ctk.CTkLabel(
                app.stats_breakdown,
                text=f"{r.get('read_date')} — {r.get('pages_read', 0)} стр.",
                anchor="w",
            ).pack(fill="x", pady=2)


def refresh_pace(app) -> None:
    sel = app.combo_pace_book.get()
    if not sel or sel.startswith("—"):
        app.lbl_pace.configure(text="Выберите книгу в списке.")
        return
    try:
        ids = getattr(app, '_pace_book_ids', [])
        titles = app.combo_pace_book.cget('values')
        idx = list(titles).index(sel) if sel in titles else -1
        bid = ids[idx] if 0 <= idx < len(ids) else None
        if bid is None:
            return
    except (ValueError, IndexError):
        return
    try:
        p = db.book_pace_estimate(bid, 14)
    except Exception as e:
        app.lbl_pace.configure(text=str(e))
        return
    if not p:
        return
    lines = [f"«{p['title']}»: в книге указано {p['page_count']} стр., по трекеру набрано {p['pages_logged']} стр."]
    if p["remaining_pages"] is not None:
        lines.append(f"Осталось страниц: {p['remaining_pages']}.")
    lines.append(
        f"За последние {p['window_days']} дн.: в среднем {p['avg_pages_per_active_day']} стр. в день "
        f"({p['days_with_reading_in_window']} таких дней)."
    )
    if p["estimated_days_to_finish"] is not None:
        lines.append(f"При таком темпе до конца ориентировочно ~{p['estimated_days_to_finish']} дн.")
    else:
        lines.append("Срок до конца оценить нельзя (нет объёма книги или нет устойчивого темпа).")
    app.lbl_pace.configure(text="\n".join(lines))


# Графики

def build_charts(app) -> ctk.CTkFrame:
    f = ctk.CTkFrame(app._content_host, fg_color="transparent")
    f.grid_columnconfigure(0, weight=1)
    f.grid_rowconfigure(1, weight=1)
    ctk.CTkLabel(f, text="Графики", font=ctk.CTkFont(size=18, weight="bold")).grid(
        row=0, column=0, sticky="w", pady=(0, 8)
    )
    app._charts_host = ctk.CTkFrame(f, fg_color="transparent")
    app._charts_host.grid(row=1, column=0, sticky="nsew")
    app._charts_host.grid_columnconfigure(0, weight=1)
    app._charts_host.grid_rowconfigure(0, weight=1)
    return f


def refresh_charts(app) -> None:
    for w in app._charts_host.winfo_children():
        w.destroy()
    mode = app._theme_map.get(app.var_theme.get(), "dark")
    try:
        frame = create_charts_frame(app._charts_host, theme_mode=mode)
        frame.grid(row=0, column=0, sticky="nsew")
    except Exception as e:
        ctk.CTkLabel(app._charts_host, text=f"Ошибка загрузки графиков: {e}").grid(row=0, column=0)


# Поиск

def build_search(app) -> ctk.CTkFrame:
    f = ctk.CTkFrame(app._content_host, fg_color="transparent")
    f.grid_columnconfigure(0, weight=1)
    f.grid_rowconfigure(3, weight=1)
    ctk.CTkLabel(f, text="Поиск по книгам", font=ctk.CTkFont(size=18, weight="bold")).grid(
        row=0, column=0, sticky="w", pady=(0, 8)
    )
    bar = ctk.CTkFrame(f, fg_color="transparent")
    bar.grid(row=1, column=0, sticky="ew")
    app.var_search_q = ctk.StringVar()
    ent = ctk.CTkEntry(bar, textvariable=app.var_search_q, width=420,
                       placeholder_text="Название, автор, жанр, цитата, тег, заметка в трекере…")
    ent.pack(side="left", padx=(0, 8))
    ent.bind("<Return>", lambda e: app._run_search())
    ctk.CTkButton(bar, text="Найти", width=100, command=app._run_search).pack(side="left", padx=4)
    ctk.CTkLabel(
        f,
        text="Горячие клавиши: Ctrl+Shift+F — поиск, Ctrl+Shift+S — сохранить книгу, "
             "Ctrl+Shift+N — новая книга, Ctrl+Shift+T — фокус на страницах трекера.",
        font=ctk.CTkFont(size=12), wraplength=760, justify="left", anchor="w",
    ).grid(row=2, column=0, sticky="w", pady=6)
    app.search_results = ctk.CTkScrollableFrame(f, height=480)
    app.search_results.grid(row=3, column=0, sticky="nsew", pady=8)
    return f


def run_search(app) -> None:
    q = app.var_search_q.get().strip()
    for w in app.search_results.winfo_children():
        w.destroy()
    if not q:
        ctk.CTkLabel(app.search_results, text="Введите запрос.").pack(anchor="w", pady=8)
        return
    try:
        hits = db.search_books(q)
    except Exception as e:
        messagebox.showerror("Ошибка", str(e))
        return
    if not hits:
        ctk.CTkLabel(app.search_results, text="Ничего не найдено.").pack(anchor="w", pady=8)
        return
    for h in hits:
        label = f"{(h.get('title') or '')[:56]}  —  {h.get('match', '')}"
        ctk.CTkButton(
            app.search_results, text=label, anchor="w",
            fg_color="transparent", hover_color=("gray70", "gray40"),
            command=lambda bid=int(h["id"]): app._goto_book(bid),
        ).pack(fill="x", pady=2)
