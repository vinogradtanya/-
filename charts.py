"""
Модуль для визуализации статистики (Matplotlib + Tkinter).
"""
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.ticker as ticker
import customtkinter as ctk

# Настройка шрифтов Matplotlib (чтобы русский текст не был "квадратиками")
plt.rcParams['font.sans-serif'] = ['DejaVu Sans']


def create_charts_frame(parent, theme_mode: str = "dark") -> ctk.CTkFrame:
    """
    Создает фрейм с двумя графиками и встраивает его через grid.
    parent: родительский виджет (CTkFrame)
    theme_mode: 'dark' или 'light'
    """
    import db

    bg = "#1f1f1f" if theme_mode == "dark" else "#f3f4f6"
    fg_text = "white" if theme_mode == "dark" else "black"
    bar_color = "#3b82f6" if theme_mode == "dark" else "#2563eb"

    main_frame = ctk.CTkFrame(parent, fg_color="transparent")
    main_frame.grid_columnconfigure(0, weight=1)
    main_frame.grid_columnconfigure(1, weight=1)
    main_frame.grid_rowconfigure(0, weight=1)

    # График 1: Страницы по месяцам
    fig1, ax1 = plt.subplots(figsize=(7, 4))
    fig1.patch.set_facecolor(bg)
    ax1.set_facecolor(bg)

    raw_data = db.get_reading_stats_per_month()
    if raw_data:
        months = [x[0] for x in raw_data]
        pages = [int(x[1]) for x in raw_data]
        ax1.bar(months, pages, color=bar_color, width=0.6)
        ax1.set_ylabel("Страниц", color=fg_text)
        ax1.set_title("Активность чтения (по месяцам)", color=fg_text)
        ax1.tick_params(colors=fg_text, axis="both")
        ax1.yaxis.set_major_formatter(ticker.EngFormatter())
        plt.setp(ax1.get_xticklabels(), rotation=30, ha="right", color=fg_text)
        for spine in ax1.spines.values():
            spine.set_edgecolor(fg_text)
    else:
        ax1.text(0.5, 0.5, "Нет данных о чтении", ha="center", va="center", color="gray", transform=ax1.transAxes)
        ax1.set_xticks([])
        ax1.set_yticks([])

    fig1.tight_layout()
    canvas1 = FigureCanvasTkAgg(fig1, master=main_frame)
    canvas1.draw()
    canvas1.get_tk_widget().grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=10)

    # График 2: Жанры
    fig2, ax2 = plt.subplots(figsize=(7, 4))
    fig2.patch.set_facecolor(bg)
    ax2.set_facecolor(bg)

    raw_genres = db.get_genres_stats()
    if raw_genres:
        genres = [x[0] for x in raw_genres]
        counts = [int(x[1]) for x in raw_genres]
        colors = list(plt.get_cmap("tab10").colors)
        wedges, texts, autotexts = ax2.pie(
            counts, labels=genres, autopct="%1.1f%%",
            startangle=90, colors=colors,
        )
        for t in texts + autotexts:
            t.set_color(fg_text)
        ax2.set_title("Распределение по жанрам", color=fg_text)
    else:
        ax2.text(0.5, 0.5, "Нет данных о жанрах", ha="center", va="center", color="gray", transform=ax2.transAxes)
        ax2.set_xticks([])
        ax2.set_yticks([])

    fig2.tight_layout()
    canvas2 = FigureCanvasTkAgg(fig2, master=main_frame)
    canvas2.draw()
    canvas2.get_tk_widget().grid(row=0, column=1, sticky="nsew", padx=(8, 0), pady=10)

    return main_frame