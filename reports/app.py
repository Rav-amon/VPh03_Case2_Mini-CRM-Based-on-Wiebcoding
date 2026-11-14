"""Окно генератора отчётов.

Человек задаёт период, подразделение, вид отчёта и ответственного.
Остальные цифры моделируются, лист оформляется как документ и пишется
в Google Таблицу через google_sheets.py.
"""

from __future__ import annotations

import threading
import tkinter as tk
import webbrowser
from datetime import date, datetime
from random import Random
from tkinter import ttk

from google_sheets import GoogleSheetsClient, GoogleSheetsError, PublishedSheet, service_account_email
from report_model import (
    DEPARTMENTS,
    REPORT_KINDS,
    BuiltReport,
    ReportParams,
    format_money,
    format_percent,
    parse_report_params,
    build_report,
)

_WINDOW_BG = "#F4F7F8"
_INK = "#1F2933"
_MUTED = "#5C6B73"
_NAVY = "#1B3A4B"
_LINE = "#D5DEE7"
_PAPER = "#FFFFFF"


class ReportApp(tk.Tk):
    """Форма из нескольких полей и кнопка записи листа."""

    def __init__(self) -> None:
        super().__init__()
        self.title("Генератор отчётов")
        self.geometry("680x680")
        self.minsize(620, 600)
        self.configure(background=_WINDOW_BG)
        self._busy = False
        self._url = tk.StringVar(value="")
        today = date.today()
        self._start = tk.StringVar(value=today.replace(day=1).strftime("%d.%m.%Y"))
        self._end = tk.StringVar(value=today.strftime("%d.%m.%Y"))
        self._department = tk.StringVar(value=DEPARTMENTS[0])
        self._kind = tk.StringVar(value=REPORT_KINDS[0])
        self._author = tk.StringVar(value="")
        self._apply_theme()
        self._build()
        self.bind("<Return>", lambda _event: self._on_submit())

    def _apply_theme(self) -> None:
        """Спокойная палитра, близкая к шапке самого отчёта."""
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            return
        style.configure(".", background=_WINDOW_BG, foreground=_INK, font=("Segoe UI", 10))
        style.configure("TFrame", background=_WINDOW_BG)
        style.configure("Card.TFrame", background=_PAPER)
        style.configure("TLabel", background=_WINDOW_BG, foreground=_INK)
        style.configure("Card.TLabel", background=_PAPER, foreground=_INK)
        style.configure(
            "Title.TLabel",
            background=_WINDOW_BG,
            foreground=_NAVY,
            font=("Georgia", 20, "bold"),
        )
        style.configure(
            "Hint.TLabel",
            background=_WINDOW_BG,
            foreground=_MUTED,
            font=("Segoe UI", 9),
        )
        style.configure(
            "CardHint.TLabel",
            background=_PAPER,
            foreground=_MUTED,
            font=("Segoe UI", 9),
        )
        style.configure(
            "Accent.TButton",
            background=_NAVY,
            foreground=_PAPER,
            font=("Segoe UI", 10, "bold"),
            padding=(14, 8),
        )
        style.map(
            "Accent.TButton",
            background=[("disabled", "#9AA8B2"), ("active", "#245068")],
            foreground=[("disabled", _PAPER)],
        )
        style.configure("TEntry", padding=4)
        style.configure("TCombobox", padding=4)

    def _build(self) -> None:
        """Собрать форму. Сетка простая: подпись слева, поле справа."""
        root = ttk.Frame(self, padding=(24, 18, 24, 16))
        root.pack(fill="both", expand=True)
        ttk.Label(root, text="Генератор отчётов", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            root,
            text="Укажите период и несколько полей. Цифры по дням программа смоделирует сама.",
            style="Hint.TLabel",
        ).pack(anchor="w", pady=(4, 14))

        card = ttk.Frame(root, style="Card.TFrame", padding=16)
        card.pack(fill="x")
        card.columnconfigure(1, weight=1)

        self._add_period(card)
        self._add_choice(card, 1, "Подразделение", self._department, DEPARTMENTS)
        self._add_choice(card, 2, "Вид отчёта", self._kind, REPORT_KINDS)
        self._add_author(card)

        actions = ttk.Frame(root)
        actions.pack(fill="x", pady=(14, 8))
        self._button = ttk.Button(
            actions,
            text="Сформировать отчёт",
            style="Accent.TButton",
            command=self._on_submit,
        )
        self._button.pack(side="left")
        self._open_button = ttk.Button(
            actions,
            text="Открыть таблицу",
            command=self._open_sheet,
        )
        self._open_button.pack(side="left", padx=(8, 0))
        self._open_button.state(["disabled"])

        self._log = tk.Text(
            root,
            height=8,
            wrap="word",
            relief="flat",
            background=_PAPER,
            foreground=_INK,
            font=("Segoe UI", 10),
            padx=12,
            pady=10,
            highlightthickness=1,
            highlightbackground=_LINE,
            highlightcolor=_NAVY,
        )
        self._log.pack(fill="both", expand=True, pady=(8, 8))
        self._log.bind("<Key>", self._keep_log_readonly)
        self._log.bind("<<Paste>>", lambda _event: "break")
        self._set_log("Таблица берётся из .env. После записи здесь появится ссылка на лист.")

        url_row = ttk.Frame(root)
        url_row.pack(fill="x")
        ttk.Label(url_row, text="Ссылка", style="Hint.TLabel").pack(anchor="w")
        self._url_entry = ttk.Entry(url_row, textvariable=self._url, state="readonly")
        self._url_entry.pack(fill="x", pady=(2, 0))

    def _add_period(self, card: ttk.Frame) -> None:
        """Два поля даты в одной строке."""
        ttk.Label(card, text="Период", style="Card.TLabel").grid(
            row=0, column=0, sticky="w", padx=(0, 12), pady=6
        )
        period = ttk.Frame(card, style="Card.TFrame")
        period.grid(row=0, column=1, sticky="ew", pady=6)
        ttk.Entry(period, textvariable=self._start, width=14).pack(side="left")
        ttk.Label(period, text="—", style="Card.TLabel").pack(side="left", padx=8)
        ttk.Entry(period, textvariable=self._end, width=14).pack(side="left")
        ttk.Label(period, text="ДД.ММ.ГГГГ", style="CardHint.TLabel").pack(side="left", padx=(10, 0))

    def _add_choice(
        self,
        card: ttk.Frame,
        row: int,
        label: str,
        variable: tk.StringVar,
        values: tuple[str, ...],
    ) -> None:
        """Выпадающий список только из заранее заданных значений."""
        ttk.Label(card, text=label, style="Card.TLabel").grid(
            row=row, column=0, sticky="w", padx=(0, 12), pady=6
        )
        box = ttk.Combobox(
            card,
            textvariable=variable,
            values=values,
            state="readonly",
            width=36,
        )
        box.grid(row=row, column=1, sticky="ew", pady=6)

    def _add_author(self, card: ttk.Frame) -> None:
        """Свободное поле. Длинное имя в лист всё равно не влезет."""
        ttk.Label(card, text="Ответственный", style="Card.TLabel").grid(
            row=3, column=0, sticky="w", padx=(0, 12), pady=6
        )
        ttk.Entry(card, textvariable=self._author).grid(row=3, column=1, sticky="ew", pady=6)

    def _on_submit(self) -> None:
        """Проверить поля и увести запись в фон, чтобы окно не замирало."""
        if self._busy:
            return
        try:
            params = parse_report_params(
                self._start.get(),
                self._end.get(),
                self._department.get(),
                self._kind.get(),
                self._author.get(),
            )
        except ValueError as error:
            self._set_log(str(error))
            return
        self._busy = True
        self._button.state(["disabled"])
        self._set_log("Собираю дни и записываю лист в таблицу…")
        worker = threading.Thread(target=self._publish, args=(params,), daemon=True)
        worker.start()

    def _publish(self, params: ReportParams) -> None:
        """Смоделировать отчёт и записать его. Выполняется не в потоке окна."""
        try:
            report = build_report(params, Random(), datetime.now())
            client = GoogleSheetsClient.from_env()
            published = client.publish_document(
                report.sheet_name,
                report.rows,
                report.presentation,
            )
        except GoogleSheetsError as error:
            self.after(0, lambda error=error: self._show_failure(_format_sheets_error(error)))
            return
        except OSError as error:
            self.after(0, lambda error=error: self._show_failure(f"Не удалось открыть ключ или сеть.\n{error}"))
            return
        except Exception as error:
            self.after(0, lambda error=error: self._show_failure(f"Не удалось записать отчёт.\n{error}"))
            return
        self.after(0, lambda report=report, published=published: self._show_success(report, published))

    def _show_success(self, report: BuiltReport, published: PublishedSheet) -> None:
        """Показать ссылку и сразу открыть лист."""
        self._busy = False
        self._button.state(["!disabled"])
        self._open_button.state(["!disabled"])
        self._url.set(published.url)
        self._set_log(
            f"Лист «{published.title}» записан.\n"
            f"Дней: {report.day_count}. Заказов: {report.orders}. "
            f"Выручка: {format_money(report.revenue)}. "
            f"Маржа: {format_percent(report.margin)}.\n\n"
            "Сводные показатели и строка «Итого» в листе считаются формулами."
        )
        self._open_sheet()

    def _show_failure(self, message: str) -> None:
        """Вернуть кнопку и показать текст ошибки."""
        self._busy = False
        self._button.state(["!disabled"])
        self._set_log(message)

    def _open_sheet(self) -> None:
        """Открыть последнюю ссылку в браузере."""
        url = self._url.get().strip()
        if not url:
            return
        try:
            webbrowser.open(url)
        except OSError as error:
            self._set_log(f"Ссылка готова, но браузер не открылся.\n{error}\n{url}")

    def _set_log(self, message: str) -> None:
        """Заменить текст журнала. Правка с клавиатуры всё равно заблокирована."""
        self._log.delete("1.0", "end")
        self._log.insert("1.0", message)

    def _keep_log_readonly(self, event: tk.Event[tk.Misc]) -> str | None:
        """Оставить копирование и стрелки, остальной ввод в журнал не пускать."""
        if event.state & 0x4 and event.keysym.lower() in {"c", "a"}:
            return None
        allowed = {"Left", "Right", "Up", "Down", "Home", "End", "Prior", "Next", "Shift_L", "Shift_R"}
        if event.keysym in allowed:
            return None
        return "break"


def _format_sheets_error(error: GoogleSheetsError) -> str:
    """Собрать понятный текст. При отказе в доступе подсказать почту сервисного аккаунта."""
    lines = [error.message]
    if error.code in {"FORBIDDEN", "NOT_FOUND", "UNAUTHORIZED"}:
        try:
            email = service_account_email()
        except GoogleSheetsError:
            email = ""
        if email:
            lines.append(f"Выдайте доступ «Редактор» для {email}.")
    if isinstance(error.details, dict):
        reason = error.details.get("reason")
        if isinstance(reason, str) and reason and reason not in lines:
            lines.append(reason)
    return "\n".join(lines)


def main() -> None:
    """Запустить окно."""
    app = ReportApp()
    app.mainloop()


if __name__ == "__main__":
    main()
