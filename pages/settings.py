"""Страница настроек приложения."""
from __future__ import annotations

from tkinter import messagebox

import customtkinter as ctk

import dialogs


def build(app) -> ctk.CTkFrame:
    f = ctk.CTkFrame(app._content_host, fg_color="transparent")
    f.grid_columnconfigure(0, weight=1)
    ctk.CTkLabel(f, text="Настройки", font=ctk.CTkFont(size=18, weight="bold")).grid(
        row=0, column=0, sticky="w", pady=(0, 12)
    )

    app.var_ui_scale_menu = ctk.StringVar(value="1.0")
    g1 = ctk.CTkFrame(f, fg_color="transparent")
    g1.grid(row=1, column=0, sticky="w", pady=6)
    ctk.CTkLabel(g1, text="Масштаб интерфейса").pack(side="left", padx=(0, 12))
    ctk.CTkOptionMenu(g1, values=["0.9", "1.0", "1.1", "1.25"], variable=app.var_ui_scale_menu, width=100).pack(side="left")

    app.var_compact_chk = ctk.BooleanVar(value=False)
    ctk.CTkCheckBox(f, text="Компактный список книг (короткие названия)", variable=app.var_compact_chk).grid(
        row=2, column=0, sticky="w", pady=6
    )

    app.var_reminder_on = ctk.BooleanVar(value=False)
    ctk.CTkCheckBox(
        f, text="Напоминание раз в день записать чтение (локально, без интернета)",
        variable=app.var_reminder_on,
    ).grid(row=3, column=0, sticky="w", pady=6)
    rt = ctk.CTkFrame(f, fg_color="transparent")
    rt.grid(row=4, column=0, sticky="w", pady=4)
    ctk.CTkLabel(rt, text="Время напоминания (ЧЧ:ММ)").pack(side="left", padx=(0, 8))
    app.var_reminder_time = ctk.StringVar(value="20:00")
    ctk.CTkEntry(rt, textvariable=app.var_reminder_time, width=72).pack(side="left")

    ctk.CTkButton(f, text="Сохранить настройки", command=app._apply_app_preferences, height=34).grid(
        row=5, column=0, sticky="w", pady=16
    )

    ctk.CTkLabel(f, text="Резервная копия (JSON)", font=ctk.CTkFont(weight="bold")).grid(
        row=6, column=0, sticky="w", pady=(16, 8)
    )
    bf = ctk.CTkFrame(f, fg_color="transparent")
    bf.grid(row=7, column=0, sticky="w")
    ctk.CTkButton(bf, text="Экспорт в JSON…", command=lambda: dialogs.export_backup_dialog()).pack(side="left", padx=(0, 8))
    ctk.CTkButton(bf, text="Импорт из JSON…", fg_color="#8B4513", command=lambda: _import(app)).pack(side="left")

    ctk.CTkLabel(f, text="Экспорт статистики в Excel", font=ctk.CTkFont(weight="bold")).grid(
        row=8, column=0, sticky="w", pady=(16, 8)
    )
    ef = ctk.CTkFrame(f, fg_color="transparent")
    ef.grid(row=9, column=0, sticky="w")
    ctk.CTkButton(ef, text="Экспорт в Excel…", fg_color="#207245", command=lambda: dialogs.export_excel_dialog()).pack(side="left")
    ctk.CTkLabel(
        f,
        text="Экспортирует список всех книг с обложками, датами, статусом, количеством страниц и общую статистику.",
        font=ctk.CTkFont(size=11), wraplength=600, justify="left", anchor="w", text_color="gray",
    ).grid(row=10, column=0, sticky="w", pady=(4, 0))

    return f


def _import(app) -> None:
    success, _ = dialogs.import_backup_dialog()
    if success:
        app._current_book_id = None
        app._show_page("home")
        app._refresh_book_list(select_id=None)


def sync_widgets(app) -> None:
    sc = str(app._app_settings.get("ui_scale", 1.0))
    app.var_ui_scale_menu.set(sc if sc in ("0.9", "1.0", "1.1", "1.25") else "1.0")
    app.var_compact_chk.set(bool(app._app_settings.get("compact_list")))
    app.var_reminder_on.set(bool(app._app_settings.get("reminder_enabled")))
    app.var_reminder_time.set(str(app._app_settings.get("reminder_time", "20:00")))
