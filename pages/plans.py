"""Страницы: план на месяц, план на год, статусы чтения."""
from __future__ import annotations

from tkinter import messagebox

import customtkinter as ctk

import db
from utils import today, int_or


def build_month(app) -> ctk.CTkFrame:
    f = ctk.CTkFrame(app._content_host, fg_color="transparent")
    f.grid_columnconfigure(0, weight=1)
    f.grid_rowconfigure(2, weight=1)
    ctk.CTkLabel(f, text="План чтения на месяц", font=ctk.CTkFont(size=18, weight="bold")).grid(
        row=0, column=0, sticky="w", pady=(0, 8)
    )
    bar = ctk.CTkFrame(f, fg_color="transparent")
    bar.grid(row=1, column=0, sticky="ew", pady=(0, 8))
    ctk.CTkLabel(bar, text="Год").pack(side="left", padx=(0, 6))
    ctk.CTkEntry(bar, textvariable=app.var_plan_year, width=80).pack(side="left", padx=4)
    ctk.CTkLabel(bar, text="Месяц").pack(side="left", padx=(16, 6))
    ctk.CTkEntry(bar, textvariable=app.var_plan_month, width=56).pack(side="left", padx=4)
    ctk.CTkButton(bar, text="Обновить список", command=app._refresh_plan_month_list).pack(side="left", padx=20)
    app._scroll_plan_month = ctk.CTkScrollableFrame(f, height=520)
    app._scroll_plan_month.grid(row=2, column=0, sticky="nsew")
    return f


def build_year(app) -> ctk.CTkFrame:
    f = ctk.CTkFrame(app._content_host, fg_color="transparent")
    f.grid_columnconfigure(0, weight=1)
    f.grid_rowconfigure(2, weight=1)
    ctk.CTkLabel(f, text="План чтения на год", font=ctk.CTkFont(size=18, weight="bold")).grid(
        row=0, column=0, sticky="w", pady=(0, 8)
    )
    bar = ctk.CTkFrame(f, fg_color="transparent")
    bar.grid(row=1, column=0, sticky="ew", pady=(0, 8))
    ctk.CTkLabel(bar, text="Год плана").pack(side="left", padx=(0, 6))
    ctk.CTkEntry(bar, textvariable=app.var_plan_year, width=80).pack(side="left", padx=4)
    ctk.CTkButton(bar, text="Обновить список", command=app._refresh_plan_year_list).pack(side="left", padx=20)
    app._scroll_plan_year = ctk.CTkScrollableFrame(f, height=520)
    app._scroll_plan_year.grid(row=2, column=0, sticky="nsew")
    return f


def build_status(app, status_key: str, page_key: str, heading: str) -> ctk.CTkFrame:
    f = ctk.CTkFrame(app._content_host, fg_color="transparent")
    f.grid_columnconfigure(0, weight=1)
    f.grid_rowconfigure(2, weight=1)
    ctk.CTkLabel(f, text=heading, font=ctk.CTkFont(size=18, weight="bold")).grid(
        row=0, column=0, sticky="w", pady=(0, 8)
    )
    scroll = ctk.CTkScrollableFrame(f, height=520)
    scroll.grid(row=2, column=0, sticky="nsew")
    ctk.CTkButton(
        f, text="Обновить список",
        command=lambda sk=status_key, sc=scroll: refresh_status(app, sk, sc),
    ).grid(row=1, column=0, sticky="w", pady=(0, 8))
    if not hasattr(app, "_status_scrolls"):
        app._status_scrolls = {}
    app._status_scrolls[page_key] = scroll
    refresh_status(app, status_key, scroll)
    return f


def refresh_month(app) -> None:
    for w in app._scroll_plan_month.winfo_children():
        w.destroy()
    y = int_or(app.var_plan_year.get(), today().year)
    m = max(1, min(12, int_or(app.var_plan_month.get(), today().month)))
    try:
        books = db.list_books_for_month(y, m)
    except Exception as e:
        messagebox.showerror("Ошибка", str(e))
        return
    if not books:
        ctk.CTkLabel(app._scroll_plan_month, text="Нет книг в плане на выбранный месяц.").pack(anchor="w", pady=8)
        return
    ctk.CTkLabel(app._scroll_plan_month, text=f"Книг в списке: {len(books)}",
                 text_color=("gray40", "gray60")).pack(anchor="w", pady=(0, 4))
    for b in books:
        ctk.CTkButton(
            app._scroll_plan_month, text=(b.get('title') or '')[:60],
            anchor="w", fg_color="transparent", hover_color=("gray70", "gray50"),
            command=lambda bid=b["id"]: app._goto_book(bid),
        ).pack(fill="x", pady=2)


def refresh_year(app) -> None:
    for w in app._scroll_plan_year.winfo_children():
        w.destroy()
    y = int_or(app.var_plan_year.get(), today().year)
    try:
        books = db.list_books_for_year(y)
    except Exception as e:
        messagebox.showerror("Ошибка", str(e))
        return
    if not books:
        ctk.CTkLabel(app._scroll_plan_year, text="Нет книг с планом на выбранный год.").pack(anchor="w", pady=8)
        return
    ctk.CTkLabel(app._scroll_plan_year, text=f"Книг в списке: {len(books)}",
                 text_color=("gray40", "gray60")).pack(anchor="w", pady=(0, 4))
    for b in books:
        pm = b.get("plan_month")
        mo = f", мес. {pm}" if pm else ""
        ctk.CTkButton(
            app._scroll_plan_year, text=f"{(b.get('title') or '')[:60]}{mo}",
            anchor="w", fg_color="transparent", hover_color=("gray70", "gray30"),
            command=lambda bid=b["id"]: app._goto_book(bid),
        ).pack(fill="x", pady=2)


def refresh_status(app, status_key: str, scroll: ctk.CTkScrollableFrame) -> None:
    for w in scroll.winfo_children():
        w.destroy()
    try:
        books = db.list_books_by_status(status_key)
    except Exception as e:
        messagebox.showerror("Ошибка", str(e))
        return
    if not books:
        ctk.CTkLabel(scroll, text="Список пуст.").pack(anchor="w", pady=8)
        return
    ctk.CTkLabel(scroll, text=f"Книг в списке: {len(books)}",
                 text_color=("gray40", "gray60")).pack(anchor="w", pady=(0, 4))
    for b in books:
        ctk.CTkButton(
            scroll, text=(b.get('title') or '')[:60],
            anchor="w", fg_color="transparent", hover_color=("gray70", "gray30"),
            command=lambda bid=b["id"]: app._goto_book(bid),
        ).pack(fill="x", pady=2)
