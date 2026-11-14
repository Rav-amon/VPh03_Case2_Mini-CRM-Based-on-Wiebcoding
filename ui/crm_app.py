"""Окно мини-CRM.

Это клиент локального сервера: таблицы и формы только показывают данные
и отправляют HTTP-запросы на http://127.0.0.1:8000. Базу окно само не открывает.

Запуск из корня проекта:

    .\\.venv\\Scripts\\python.exe ui\\crm_app.py

Сервер, если он ещё не запущен:

    .\\.venv\\Scripts\\python.exe backend\\api.py
"""

from __future__ import annotations

import json
import queue
import subprocess
import sys
import threading
import tkinter as tk
import webbrowser
from tkinter import messagebox, ttk
import urllib.error
import urllib.parse
import urllib.request
from collections.abc import Callable
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any

from google_export import (
    DriveFailure,
    ExportError,
    GoogleSettings,
    SheetsFailure,
    TableView,
    ask_google_settings,
    load_google_settings,
    publish_report,
)

API_HOST = "127.0.0.1"
API_PORT = 8000
API_BASE = f"http://{API_HOST}:{API_PORT}"
API_PREFIX = "/api/v1"
API_TIMEOUT_SECONDS = 8
PAGE_LIMIT = 100
MINOR_SCALE_DEFAULT = 100
CREATE_NEW_CONSOLE = 0x00000010
SERVER_POLL_ATTEMPTS = 15
SERVER_POLL_MS = 400

_WINDOW_BG = "#F4F7F8"
_INK = "#1F2933"
_MUTED = "#5C6B73"
_NAVY = "#1B3A4B"
_PAPER = "#FFFFFF"
_LINE = "#D5DEE7"

_CLIENT_STATUSES = (
    ("lead", "Потенциальный"),
    ("active", "Действующий"),
    ("inactive", "Неактивный"),
    ("archived", "Архив"),
)
_DEAL_STATUSES = (
    ("new", "Новая"),
    ("qualification", "Уточнение"),
    ("proposal", "Предложение"),
    ("negotiation", "Согласование"),
    ("won", "Выиграна"),
    ("lost", "Проиграна"),
    ("archived", "Архив"),
)
_ORDER_STATUSES = (
    ("draft", "Черновик"),
    ("confirmed", "Подтверждён"),
    ("in_progress", "В работе"),
    ("fulfilled", "Выполнен"),
    ("cancelled", "Отменён"),
)
_TASK_KINDS = (
    ("task", "Задача"),
    ("reminder", "Напоминание"),
)
_DONE_FILTERS = (
    ("", "Все"),
    ("false", "Открытые"),
    ("true", "Выполненные"),
)
_NONE_OPTION = ("", "— нет —")


class ApiError(Exception):
    """Ошибка запроса в том же виде, что отдаёт сервер: code, message, details."""

    def __init__(self, code: str, message: str, details: Any = None) -> None:
        self.code = code
        self.message = message
        self.details = details
        super().__init__(message)


@dataclass(frozen=True)
class FieldSpec:
    """Поле формы: ключ, подпись и вид виджета."""

    key: str
    label: str
    kind: str
    options: tuple[tuple[str, str], ...] = ()


def parse_amount(text: str, scale: int) -> int:
    """Перевести «1 500,50» в мелкие единицы. Пустая строка — ноль."""
    cleaned = text.strip().replace(" ", "").replace(",", ".")
    if cleaned == "":
        return 0
    try:
        value = Decimal(cleaned)
    except InvalidOperation as error:
        raise ApiError(
            "VALIDATION_ERROR",
            "Сумма должна быть числом, например 1500.50.",
            {"amount": text},
        ) from error
    if value < 0:
        raise ApiError("VALIDATION_ERROR", "Сумма не может быть отрицательной.", None)
    minor = (value * Decimal(scale)).quantize(Decimal("1"), rounding=ROUND_HALF_UP)
    return int(minor)


def optional_id(value: object, label: str) -> int | None:
    """Пустая связь остаётся пустой, иначе это положительный идентификатор."""
    if value is None or value == "":
        return None
    if not isinstance(value, str) or not value.isdigit():
        raise ApiError(
            "VALIDATION_ERROR",
            f"Некорректная связь: {label}.",
            {"value": value},
        )
    number = int(value)
    if number < 1:
        raise ApiError(
            "VALIDATION_ERROR",
            f"Некорректная связь: {label}.",
            {"value": value},
        )
    return number


def require_text(value: object, label: str) -> str:
    """Обрезать строку и не пропускать пустое обязательное поле."""
    if not isinstance(value, str) or value.strip() == "":
        raise ApiError("VALIDATION_ERROR", f"Заполните поле «{label}».", {"field": label})
    return value.strip()


def plain_text(value: object) -> str:
    """Строка из формы. Чужой тип становится пустой строкой."""
    if not isinstance(value, str):
        return ""
    return value.strip()


def build_client_body(raw: dict[str, Any]) -> dict[str, Any]:
    """Тело POST/PUT клиента."""
    return {
        "full_name": require_text(raw.get("full_name"), "имя"),
        "company": plain_text(raw.get("company")),
        "email": plain_text(raw.get("email")),
        "phone": plain_text(raw.get("phone")),
        "status": plain_text(raw.get("status")) or "active",
        "source": plain_text(raw.get("source")),
        "notes": plain_text(raw.get("notes")),
    }


def build_deal_body(raw: dict[str, Any], scale: int) -> dict[str, Any]:
    """Тело POST/PUT сделки. Сумма уходит в копейках."""
    return {
        "title": require_text(raw.get("title"), "название"),
        "client_id": optional_id(raw.get("client_id"), "клиент"),
        "status": plain_text(raw.get("status")) or "new",
        "amount_minor": parse_amount(plain_text(raw.get("amount")), scale),
        "currency": (plain_text(raw.get("currency")) or "RUB").upper(),
        "expected_close_on": plain_text(raw.get("expected_close_on")),
        "notes": plain_text(raw.get("notes")),
    }


def build_order_body(raw: dict[str, Any], scale: int) -> dict[str, Any]:
    """Тело POST/PUT заказа."""
    return {
        "title": require_text(raw.get("title"), "название"),
        "client_id": optional_id(raw.get("client_id"), "клиент"),
        "deal_id": optional_id(raw.get("deal_id"), "сделка"),
        "status": plain_text(raw.get("status")) or "draft",
        "amount_minor": parse_amount(plain_text(raw.get("amount")), scale),
        "currency": (plain_text(raw.get("currency")) or "RUB").upper(),
        "due_on": plain_text(raw.get("due_on")),
        "notes": plain_text(raw.get("notes")),
    }


def build_task_body(raw: dict[str, Any]) -> dict[str, Any]:
    """Тело POST/PUT задачи. Напоминание без срока сервер тоже не примет."""
    kind = plain_text(raw.get("kind")) or "task"
    due_at = plain_text(raw.get("due_at"))
    if kind == "reminder" and due_at == "":
        raise ApiError("VALIDATION_ERROR", "У напоминания укажите срок.", {"field": "due_at"})
    done = raw.get("is_done")
    return {
        "title": require_text(raw.get("title"), "название"),
        "details": plain_text(raw.get("details")),
        "kind": kind,
        "due_at": due_at,
        "is_done": bool(done),
        "client_id": optional_id(raw.get("client_id"), "клиент"),
        "deal_id": optional_id(raw.get("deal_id"), "сделка"),
        "order_id": optional_id(raw.get("order_id"), "заказ"),
    }


def _parse_payload(raw: bytes) -> dict[str, Any]:
    try:
        payload = json.loads(raw.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as error:
        raise ApiError("BAD_RESPONSE", "Сервер вернул непонятный ответ.", None) from error
    if not isinstance(payload, dict):
        raise ApiError("BAD_RESPONSE", "Сервер вернул непонятный ответ.", None)
    return payload


def _error_from_payload(payload: dict[str, Any], status: int) -> ApiError:
    error = payload.get("error")
    if isinstance(error, dict):
        code = error.get("code")
        message = error.get("message")
        code_text = code if isinstance(code, str) and code else "HTTP_ERROR"
        message_text = message if isinstance(message, str) and message else f"Ошибка {status}."
        return ApiError(code_text, message_text, error.get("details"))
    message = payload.get("message")
    if isinstance(message, str) and message:
        return ApiError("HTTP_ERROR", message, None)
    return ApiError("HTTP_ERROR", f"Сервер ответил кодом {status}.", None)


class CrmApi:
    """Тонкий клиент локального API. Один метод — один запрос."""

    def __init__(self, base: str = API_BASE) -> None:
        self.base = base.rstrip("/")

    def request(
        self,
        method: str,
        path: str,
        body: dict[str, Any] | None = None,
        query: dict[str, Any] | None = None,
    ) -> Any:
        """Выполнить запрос и вернуть поле data. Ошибку поднять как ApiError."""
        params: list[tuple[str, str]] = []
        if query is not None:
            for key, value in query.items():
                if value is None:
                    continue
                params.append((key, str(value)))
        url = self.base + path
        if params:
            url = url + "?" + urllib.parse.urlencode(params)
        data = None
        headers = {"Accept": "application/json"}
        if body is not None:
            data = json.dumps(body, ensure_ascii=False).encode("utf-8")
            headers["Content-Type"] = "application/json; charset=utf-8"
        request = urllib.request.Request(url, data=data, headers=headers, method=method)
        try:
            with urllib.request.urlopen(request, timeout=API_TIMEOUT_SECONDS) as response:
                payload = _parse_payload(response.read())
                status = response.status
        except urllib.error.HTTPError as error:
            payload = _parse_payload(error.read())
            raise _error_from_payload(payload, error.code) from error
        except urllib.error.URLError as error:
            raise ApiError(
                "CONNECTION",
                "Нет связи с локальным сервером.",
                {"reason": str(error.reason)},
            ) from error
        if payload.get("error"):
            raise _error_from_payload(payload, status)
        return payload.get("data")

    def health(self) -> Any:
        """Проверить, что сервер отвечает."""
        return self.request("GET", f"{API_PREFIX}/health")

    def meta(self) -> Any:
        """Справочник статусов."""
        return self.request("GET", f"{API_PREFIX}/meta")

    def summary(self) -> Any:
        """Счётчики для шапки."""
        return self.request("GET", f"{API_PREFIX}/summary")

    def list_clients(self, status: str | None, offset: int = 0) -> Any:
        """Список клиентов."""
        return self.request(
            "GET",
            f"{API_PREFIX}/clients",
            query={"status": status, "limit": PAGE_LIMIT, "offset": offset},
        )

    def search_clients(self, query_text: str, offset: int = 0) -> Any:
        """Поиск клиентов."""
        return self.request(
            "GET",
            f"{API_PREFIX}/clients/search",
            query={"q": query_text, "limit": PAGE_LIMIT, "offset": offset},
        )

    def create_client(self, body: dict[str, Any]) -> Any:
        """Создать клиента."""
        return self.request("POST", f"{API_PREFIX}/clients", body)

    def update_client(self, client_id: int, body: dict[str, Any]) -> Any:
        """Заменить карточку клиента."""
        return self.request("PUT", f"{API_PREFIX}/clients/{client_id}", body)

    def set_client_status(self, client_id: int, status: str) -> Any:
        """Сменить статус клиента."""
        return self.request(
            "PATCH",
            f"{API_PREFIX}/clients/{client_id}/status",
            {"status": status},
        )

    def delete_client(self, client_id: int) -> Any:
        """Удалить клиента."""
        return self.request("DELETE", f"{API_PREFIX}/clients/{client_id}")

    def list_deals(self, status: str | None, offset: int = 0) -> Any:
        """Список сделок."""
        return self.request(
            "GET",
            f"{API_PREFIX}/deals",
            query={"status": status, "limit": PAGE_LIMIT, "offset": offset},
        )

    def search_deals(self, query_text: str, offset: int = 0) -> Any:
        """Поиск сделок."""
        return self.request(
            "GET",
            f"{API_PREFIX}/deals/search",
            query={"q": query_text, "limit": PAGE_LIMIT, "offset": offset},
        )

    def create_deal(self, body: dict[str, Any]) -> Any:
        """Создать сделку."""
        return self.request("POST", f"{API_PREFIX}/deals", body)

    def update_deal(self, deal_id: int, body: dict[str, Any]) -> Any:
        """Заменить сделку."""
        return self.request("PUT", f"{API_PREFIX}/deals/{deal_id}", body)

    def set_deal_status(self, deal_id: int, status: str) -> Any:
        """Сменить статус сделки."""
        return self.request(
            "PATCH",
            f"{API_PREFIX}/deals/{deal_id}/status",
            {"status": status},
        )

    def delete_deal(self, deal_id: int) -> Any:
        """Удалить сделку."""
        return self.request("DELETE", f"{API_PREFIX}/deals/{deal_id}")

    def list_orders(self, status: str | None, offset: int = 0) -> Any:
        """Список заказов."""
        return self.request(
            "GET",
            f"{API_PREFIX}/orders",
            query={"status": status, "limit": PAGE_LIMIT, "offset": offset},
        )

    def search_orders(self, query_text: str, offset: int = 0) -> Any:
        """Поиск заказов."""
        return self.request(
            "GET",
            f"{API_PREFIX}/orders/search",
            query={"q": query_text, "limit": PAGE_LIMIT, "offset": offset},
        )

    def create_order(self, body: dict[str, Any]) -> Any:
        """Создать заказ."""
        return self.request("POST", f"{API_PREFIX}/orders", body)

    def update_order(self, order_id: int, body: dict[str, Any]) -> Any:
        """Заменить заказ."""
        return self.request("PUT", f"{API_PREFIX}/orders/{order_id}", body)

    def set_order_status(self, order_id: int, status: str) -> Any:
        """Сменить статус заказа."""
        return self.request(
            "PATCH",
            f"{API_PREFIX}/orders/{order_id}/status",
            {"status": status},
        )

    def delete_order(self, order_id: int) -> Any:
        """Удалить заказ."""
        return self.request("DELETE", f"{API_PREFIX}/orders/{order_id}")

    def list_tasks(self, kind: str | None, is_done: str | None, offset: int = 0) -> Any:
        """Список задач. is_done — пусто, true или false."""
        return self.request(
            "GET",
            f"{API_PREFIX}/tasks",
            query={"kind": kind, "is_done": is_done, "limit": PAGE_LIMIT, "offset": offset},
        )

    def create_task(self, body: dict[str, Any]) -> Any:
        """Создать задачу."""
        return self.request("POST", f"{API_PREFIX}/tasks", body)

    def update_task(self, task_id: int, body: dict[str, Any]) -> Any:
        """Заменить задачу."""
        return self.request("PUT", f"{API_PREFIX}/tasks/{task_id}", body)

    def set_task_done(self, task_id: int, is_done: bool) -> Any:
        """Отметить задачу выполненной или открыть снова."""
        return self.request(
            "PATCH",
            f"{API_PREFIX}/tasks/{task_id}/done",
            {"is_done": is_done},
        )

    def delete_task(self, task_id: int) -> Any:
        """Удалить задачу."""
        return self.request("DELETE", f"{API_PREFIX}/tasks/{task_id}")


def _pairs_from_meta(
    items: object,
    fallback: tuple[tuple[str, str], ...],
) -> tuple[tuple[str, str], ...]:
    """Взять коды из справочника сервера. Если форма другая — оставить запасные."""
    if not isinstance(items, list):
        return fallback
    pairs: list[tuple[str, str]] = []
    for item in items:
        if not isinstance(item, dict):
            return fallback
        code = item.get("code")
        label = item.get("label")
        if not isinstance(code, str) or not isinstance(label, str):
            return fallback
        pairs.append((code, label))
    if not pairs:
        return fallback
    return tuple(pairs)


def _label_map(options: tuple[tuple[str, str], ...]) -> dict[str, str]:
    return {code: label for code, label in options}


def _with_current(
    options: tuple[tuple[str, str], ...],
    current_id: object,
    current_name: object,
) -> tuple[tuple[str, str], ...]:
    """Добавить уже выбранную связь, если её нет в свежем списке."""
    if not isinstance(current_id, int):
        return options
    key = str(current_id)
    for code, _label in options:
        if code == key:
            return options
    name = current_name if isinstance(current_name, str) and current_name else key
    return options + ((key, name),)


def raise_window(window: tk.Misc) -> None:
    """Поднять окно над остальными, иначе модальный диалог остаётся сзади и клики пропадают."""
    try:
        window.lift()
        window.focus_force()
        window.attributes("-topmost", True)
        window.after(250, lambda: window.attributes("-topmost", False))
    except tk.TclError:
        return


class Editor(tk.Toplevel):
    """Модальное окно создания и правки одной карточки."""

    def __init__(
        self,
        master: tk.Misc,
        title: str,
        specs: tuple[FieldSpec, ...],
        initial: dict[str, Any],
        on_submit: Callable[[dict[str, Any], Callable[[ApiError | None], None]], None],
    ) -> None:
        super().__init__(master)
        self.title(title)
        self.configure(background=_WINDOW_BG)
        self.minsize(480, 220)
        self.transient(master)
        self._specs = specs
        self._on_submit = on_submit
        self._vars: dict[str, tk.Variable] = {}
        self._texts: dict[str, tk.Text] = {}
        self._code_by_label: dict[str, dict[str, str]] = {}
        self._build(initial)
        self.bind("<Escape>", lambda _event: self.destroy())
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.grab_set()
        raise_window(self)

    def _build(self, initial: dict[str, Any]) -> None:
        """Собрать подписи, поля и кнопки."""
        frame = ttk_frame(self)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(1, weight=1)
        for row, spec in enumerate(self._specs):
            ttk_label(frame, spec.label).grid(
                row=row,
                column=0,
                sticky="nw",
                padx=(0, 12),
                pady=5,
            )
            self._place_field(frame, spec, initial, row)
        buttons = ttk_frame(frame)
        buttons.grid(row=len(self._specs), column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk_button(buttons, "Отмена", self.destroy).pack(side="right")
        self._save_button = ttk_button(buttons, "Сохранить", self._submit, accent=True)
        self._save_button.pack(side="right", padx=(0, 8))

    def _place_field(
        self,
        frame: tk.Widget,
        spec: FieldSpec,
        initial: dict[str, Any],
        row: int,
    ) -> None:
        """Поставить виджет по виду поля."""
        if spec.kind == "text":
            widget = tk.Text(
                frame,
                height=4,
                width=42,
                relief="flat",
                highlightthickness=1,
                highlightbackground=_LINE,
                font=("Segoe UI", 10),
            )
            value = initial.get(spec.key, "")
            if isinstance(value, str) and value:
                widget.insert("1.0", value)
            widget.grid(row=row, column=1, sticky="ew", pady=5)
            self._texts[spec.key] = widget
            return
        if spec.kind == "check":
            variable = tk.BooleanVar(value=bool(initial.get(spec.key, False)))
            tk.Checkbutton(
                frame,
                variable=variable,
                background=_WINDOW_BG,
                activebackground=_WINDOW_BG,
                font=("Segoe UI", 10),
            ).grid(row=row, column=1, sticky="w", pady=5)
            self._vars[spec.key] = variable
            return
        variable = tk.StringVar()
        if spec.kind == "combo":
            labels = [label for _code, label in spec.options]
            mapping = {label: code for code, label in spec.options}
            self._code_by_label[spec.key] = mapping
            code = initial.get(spec.key, "")
            code_text = "" if code is None else str(code)
            chosen = labels[0] if labels else ""
            for option_code, option_label in spec.options:
                if option_code == code_text:
                    chosen = option_label
            variable.set(chosen)
            widget = ttk.Combobox(
                frame,
                textvariable=variable,
                values=labels,
                state="readonly",
                width=40,
            )
        else:
            value = initial.get(spec.key, "")
            variable.set(value if isinstance(value, str) else "")
            widget = ttk.Entry(frame, textvariable=variable, width=42)
        widget.grid(row=row, column=1, sticky="ew", pady=5)
        self._vars[spec.key] = variable

    def _collect(self) -> dict[str, Any]:
        """Считать форму. Списки возвращают код, а не подпись."""
        body: dict[str, Any] = {}
        for spec in self._specs:
            if spec.kind == "text":
                body[spec.key] = self._texts[spec.key].get("1.0", "end").strip()
                continue
            if spec.kind == "check":
                body[spec.key] = bool(self._vars[spec.key].get())
                continue
            raw = self._vars[spec.key].get()
            text = raw.strip() if isinstance(raw, str) else ""
            if spec.kind == "combo":
                body[spec.key] = self._code_by_label[spec.key].get(text, "")
                continue
            body[spec.key] = text
        return body

    def _submit(self) -> None:
        """Отдать значения владельцу. Окно закроется, когда сервер подтвердит запись."""
        try:
            body = self._collect()
        except ApiError as error:
            messagebox.showerror("Мини-CRM", error.message, parent=self)
            return
        self._save_button.state(["disabled"])

        def finish(error: ApiError | None) -> None:
            if not self.winfo_exists():
                return
            self._save_button.state(["!disabled"])
            if error is not None:
                messagebox.showerror("Мини-CRM", error.message, parent=self)
                return
            self.destroy()

        self._on_submit(body, finish)


def ttk_frame(master: tk.Misc) -> ttk.Frame:
    """Рамка с полями окна."""
    return ttk.Frame(master, padding=16)


def ttk_label(master: tk.Misc, text: str) -> ttk.Label:
    """Подпись поля."""
    return ttk.Label(master, text=text)


def ttk_button(
    master: tk.Misc,
    text: str,
    command: Callable[[], None],
    accent: bool = False,
) -> ttk.Button:
    """Кнопка. accent — главное действие на форме."""
    style = "Accent.TButton" if accent else "TButton"
    return ttk.Button(master, text=text, command=command, style=style)


class ReportReadyDialog(tk.Toplevel):
    """Ссылка на новую таблицу и кнопка открытия в браузере."""

    def __init__(self, master: tk.Misc, file_name: str, url: str) -> None:
        super().__init__(master)
        self._url = url.strip()
        self.title("Таблица создана")
        self.resizable(False, False)
        self.transient(master)
        self.configure(background=_WINDOW_BG)
        frame = ttk.Frame(self, padding=16)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text=f"Таблица «{file_name}» создана.").pack(anchor="w")
        link = tk.Label(
            frame,
            text=self._url or "Ссылка не пришла.",
            fg="#1B4F72",
            cursor="hand2",
            background=_WINDOW_BG,
            font=("Segoe UI", 10, "underline"),
            wraplength=460,
            justify="left",
        )
        link.pack(anchor="w", pady=(10, 0))
        link.bind("<Button-1>", lambda _event: self._open())
        buttons = ttk.Frame(frame)
        buttons.pack(anchor="e", pady=(16, 0))
        ttk_button(buttons, "Закрыть", self.destroy).pack(side="right")
        open_button = ttk_button(buttons, "Открыть", self._open, accent=True)
        open_button.pack(side="right", padx=(0, 8))
        if self._url == "":
            open_button.state(["disabled"])
        self.bind("<Escape>", lambda _event: self.destroy())
        self.grab_set()
        raise_window(self)

    def _open(self) -> None:
        """Открыть ссылку на таблицу в браузере."""
        if self._url == "":
            return
        webbrowser.open(self._url, new=2)


class CrmApp(tk.Tk):
    """Главное окно: четыре таблицы и диалоги правки."""

    def __init__(self) -> None:
        super().__init__()
        self.title("Мини-CRM")
        self.geometry("1120x720")
        self.minsize(920, 560)
        self.configure(background=_WINDOW_BG)
        self.api = CrmApi()
        self._scale = MINOR_SCALE_DEFAULT
        self._server_process: subprocess.Popen[Any] | None = None
        self._poll_left = 0
        self._client_statuses = _CLIENT_STATUSES
        self._deal_statuses = _DEAL_STATUSES
        self._order_statuses = _ORDER_STATUSES
        self._task_kinds = _TASK_KINDS
        self._labels = {
            "clients": _label_map(_CLIENT_STATUSES),
            "deals": _label_map(_DEAL_STATUSES),
            "orders": _label_map(_ORDER_STATUSES),
            "tasks": _label_map(_TASK_KINDS),
        }
        self._offsets = {"clients": 0, "deals": 0, "orders": 0, "tasks": 0}
        self._totals = {"clients": 0, "deals": 0, "orders": 0, "tasks": 0}
        self._notebook: ttk.Notebook | None = None
        self._tab_ids: dict[str, str] = {}
        self._page_status: dict[str, str] = {}
        self._records: dict[str, dict[str, dict[str, Any]]] = {
            "clients": {},
            "deals": {},
            "orders": {},
            "tasks": {},
        }
        self._people: list[dict[str, Any]] = []
        self._deal_rows: list[dict[str, Any]] = []
        self._order_rows: list[dict[str, Any]] = []
        self._search_vars: dict[str, tk.StringVar] = {}
        self._filter_vars: dict[str, tk.StringVar] = {}
        self._filter_codes: dict[str, dict[str, str]] = {}
        self._filter_widgets: dict[str, ttk.Combobox] = {}
        self._trees: dict[str, ttk.Treeview] = {}
        self._exporting = False
        self._status = tk.StringVar(value="Нет связи с сервером.")
        self._summary = tk.StringVar(value="Сводка появится после подключения.")
        self._ui_queue: queue.Queue[Callable[[], None]] = queue.Queue()
        self._apply_theme()
        self._build()
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.after(50, self._drain_ui)
        self.after(100, lambda: raise_window(self))
        self.after(200, lambda: self._connect(silent=True))

    def _apply_theme(self) -> None:
        """Та же спокойная палитра, что у окна отчётов."""
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            return
        style.configure(".", background=_WINDOW_BG, foreground=_INK, font=("Segoe UI", 10))
        style.configure("TFrame", background=_WINDOW_BG)
        style.configure("TLabel", background=_WINDOW_BG, foreground=_INK)
        style.configure(
            "Hint.TLabel",
            background=_WINDOW_BG,
            foreground=_MUTED,
            font=("Segoe UI", 9),
        )
        style.configure(
            "Title.TLabel",
            background=_WINDOW_BG,
            foreground=_NAVY,
            font=("Georgia", 20, "bold"),
        )
        style.configure(
            "Accent.TButton",
            background=_NAVY,
            foreground=_PAPER,
            font=("Segoe UI", 10, "bold"),
            padding=(12, 6),
        )
        style.map(
            "Accent.TButton",
            background=[("disabled", "#9AA8B2"), ("active", "#245068")],
            foreground=[("disabled", _PAPER)],
        )
        style.configure("TButton", padding=(10, 6))
        style.configure(
            "Treeview",
            rowheight=28,
            font=("Segoe UI", 10),
            background=_PAPER,
            fieldbackground=_PAPER,
        )
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))
        style.configure("TNotebook", background=_WINDOW_BG)
        style.configure("TNotebook.Tab", padding=(14, 6))

    def _build(self) -> None:
        """Шапка, вкладки таблиц и строка состояния."""
        root = ttk.Frame(self, padding=(18, 14, 18, 12))
        root.pack(fill="both", expand=True)
        header = ttk.Frame(root)
        header.pack(fill="x")
        ttk.Label(header, text="Мини-CRM", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text=f"Окно само спрашивает локальный сервер {API_BASE}",
            style="Hint.TLabel",
        ).pack(anchor="w", pady=(2, 8))
        actions = ttk.Frame(header)
        actions.pack(fill="x", pady=(0, 6))
        ttk_button(actions, "Подключиться", lambda: self._connect(silent=False)).pack(side="left")
        ttk_button(actions, "Запустить сервер", self._start_server).pack(side="left", padx=(8, 0))
        ttk_button(actions, "Настройки Google", self._open_google_settings).pack(
            side="left",
            padx=(8, 0),
        )
        ttk.Label(actions, textvariable=self._summary, style="Hint.TLabel").pack(
            side="left",
            padx=(16, 0),
        )

        notebook = ttk.Notebook(root)
        notebook.pack(fill="both", expand=True, pady=(8, 8))
        notebook.bind("<<NotebookTabChanged>>", lambda _event: self._show_tab_status())
        self._notebook = notebook
        self._build_clients(notebook)
        self._build_deals(notebook)
        self._build_orders(notebook)
        self._build_tasks(notebook)
        ttk.Label(root, textvariable=self._status, style="Hint.TLabel").pack(anchor="w")

    def _build_clients(self, notebook: ttk.Notebook) -> None:
        """Вкладка клиентов."""
        frame = self._tab(notebook, "Клиенты", "clients")
        self._search_row(frame, "clients")
        filters = ttk.Frame(frame)
        filters.pack(fill="x", pady=(0, 8))
        self._filter_box(
            filters,
            "clients-status",
            "Показать",
            (("", "Все"),) + self._client_statuses,
        )
        self._action_row(
            frame,
            "clients",
            (
                ("Обновить", lambda: self._fetch_page("clients")),
                ("Создать", self._create_client),
                ("Изменить", self._edit_client),
                ("Статус", self._status_client),
                ("Удалить", self._delete_client),
                ("Выгрузить отчёт", lambda: self._export_table("clients")),
            ),
        )
        columns = (
            ("id", "ID", 50),
            ("full_name", "Имя", 180),
            ("company", "Компания", 150),
            ("phone", "Телефон", 140),
            ("email", "Почта", 180),
            ("status", "Статус", 130),
            ("source", "Источник", 120),
        )
        self._trees["clients"] = self._tree(frame, "clients", columns, self._edit_client)

    def _build_deals(self, notebook: ttk.Notebook) -> None:
        """Вкладка сделок."""
        frame = self._tab(notebook, "Сделки", "deals")
        self._search_row(frame, "deals")
        filters = ttk.Frame(frame)
        filters.pack(fill="x", pady=(0, 8))
        self._filter_box(filters, "deals-status", "Показать", (("", "Все"),) + self._deal_statuses)
        self._action_row(
            frame,
            "deals",
            (
                ("Обновить", lambda: self._fetch_page("deals")),
                ("Создать", self._create_deal),
                ("Изменить", self._edit_deal),
                ("Статус", self._status_deal),
                ("Удалить", self._delete_deal),
                ("Выгрузить отчёт", lambda: self._export_table("deals")),
            ),
        )
        columns = (
            ("id", "ID", 50),
            ("title", "Название", 220),
            ("client_name", "Клиент", 180),
            ("status", "Статус", 140),
            ("amount", "Сумма", 120),
            ("expected_close_on", "Дата", 110),
        )
        self._trees["deals"] = self._tree(frame, "deals", columns, self._edit_deal)

    def _build_orders(self, notebook: ttk.Notebook) -> None:
        """Вкладка заказов."""
        frame = self._tab(notebook, "Заказы", "orders")
        self._search_row(frame, "orders")
        filters = ttk.Frame(frame)
        filters.pack(fill="x", pady=(0, 8))
        self._filter_box(
            filters,
            "orders-status",
            "Показать",
            (("", "Все"),) + self._order_statuses,
        )
        self._action_row(
            frame,
            "orders",
            (
                ("Обновить", lambda: self._fetch_page("orders")),
                ("Создать", self._create_order),
                ("Изменить", self._edit_order),
                ("Статус", self._status_order),
                ("Удалить", self._delete_order),
                ("Выгрузить отчёт", lambda: self._export_table("orders")),
            ),
        )
        columns = (
            ("id", "ID", 50),
            ("title", "Название", 200),
            ("client_name", "Клиент", 160),
            ("deal_title", "Сделка", 160),
            ("status", "Статус", 130),
            ("amount", "Сумма", 120),
            ("due_on", "Срок", 110),
        )
        self._trees["orders"] = self._tree(frame, "orders", columns, self._edit_order)

    def _build_tasks(self, notebook: ttk.Notebook) -> None:
        """Вкладка задач и напоминаний."""
        frame = self._tab(notebook, "Задачи", "tasks")
        self._search_row(frame, "tasks")
        filters = ttk.Frame(frame)
        filters.pack(fill="x", pady=(0, 8))
        self._filter_box(filters, "tasks-kind", "Вид", (("", "Все"),) + self._task_kinds)
        self._filter_box(filters, "tasks-done", "Состояние", _DONE_FILTERS)
        self._action_row(
            frame,
            "tasks",
            (
                ("Обновить", lambda: self._fetch_page("tasks")),
                ("Создать", self._create_task),
                ("Изменить", self._edit_task),
                ("Выполнена", self._toggle_task),
                ("Удалить", self._delete_task),
                ("Выгрузить отчёт", lambda: self._export_table("tasks")),
            ),
        )
        columns = (
            ("id", "ID", 50),
            ("title", "Название", 220),
            ("kind", "Вид", 120),
            ("due_at", "Срок", 170),
            ("is_done", "Готово", 80),
            ("client_name", "Клиент", 140),
            ("deal_title", "Сделка", 140),
            ("order_title", "Заказ", 140),
        )
        self._trees["tasks"] = self._tree(frame, "tasks", columns, self._edit_task)

    def _tab(self, notebook: ttk.Notebook, title: str, key: str) -> ttk.Frame:
        """Пустая вкладка и переменная поиска."""
        frame = ttk.Frame(notebook, padding=12)
        notebook.add(frame, text=title)
        self._tab_ids[key] = notebook.tabs()[-1]
        self._search_vars[key] = tk.StringVar(value="")
        return frame

    def _search_row(self, frame: ttk.Frame, kind: str) -> None:
        """Строка поиска. Пустой запрос снова показывает список."""
        row = ttk.Frame(frame)
        row.pack(fill="x", pady=(0, 8))
        ttk.Label(row, text="Поиск").pack(side="left")
        entry = ttk.Entry(row, textvariable=self._search_vars[kind])
        entry.pack(side="left", fill="x", expand=True, padx=8)
        entry.bind("<Return>", lambda _event, name=kind: self._fetch_page(name, True))
        ttk_button(row, "Найти", lambda name=kind: self._fetch_page(name, True)).pack(side="left")

    def _filter_box(
        self,
        frame: ttk.Frame,
        key: str,
        caption: str,
        options: tuple[tuple[str, str], ...],
    ) -> None:
        """Выпадающий фильтр. В переменной хранится подпись, код лежит рядом."""
        ttk.Label(frame, text=caption).pack(side="left")
        variable = tk.StringVar(value=options[0][1] if options else "")
        self._filter_vars[key] = variable
        self._filter_codes[key] = {label: code for code, label in options}
        combo = ttk.Combobox(
            frame,
            textvariable=variable,
            values=[label for _code, label in options],
            state="readonly",
            width=18,
        )
        combo.pack(side="left", padx=(6, 16))
        combo.bind("<<ComboboxSelected>>", lambda _event, name=key: self._on_filter(name))
        self._filter_widgets[key] = combo

    def _action_row(
        self,
        frame: ttk.Frame,
        kind: str,
        buttons: tuple[tuple[str, Callable[[], None]], ...],
    ) -> None:
        """Кнопки над таблицей. Первая обычная, «Создать» выделена, если есть."""
        row = ttk.Frame(frame)
        row.pack(fill="x", pady=(0, 8))
        for text, command in buttons:
            accent = text == "Создать"
            button = ttk_button(row, text, command, accent=accent)
            if text == "Выгрузить отчёт":
                button.pack(side="right")
                continue
            button.pack(side="left", padx=(0, 6))
        ttk_button(row, "Дальше", lambda name=kind: self._shift_page(name, PAGE_LIMIT)).pack(
            side="right",
            padx=(6, 8),
        )
        ttk_button(row, "Назад", lambda name=kind: self._shift_page(name, -PAGE_LIMIT)).pack(
            side="right",
        )
        del kind

    def _open_google_settings(self) -> None:
        """Окно путей к JSON и папке. Результат уже записан в текстовый файл."""
        ask_google_settings(self, load_google_settings())

    def _export_table(self, kind: str) -> None:
        """Снять видимые строки и отправить их в новую таблицу Google."""
        if self._exporting:
            self._set_status("Выгрузка уже идёт.")
            return
        settings = self._settings_for_export()
        if settings is None:
            return
        view = self._visible_table(kind)
        self._exporting = True
        self._set_status("Собираю отчёт и создаю таблицу Google…")

        def work() -> None:
            try:
                published = publish_report(settings, view)
            except (ExportError, DriveFailure, SheetsFailure) as error:
                self._ui(lambda error=error: self._export_failed(error.message))
                return
            except Exception as error:
                self._ui(lambda error=error: self._export_failed(str(error)))
                return
            self._ui(
                lambda published=published: self._export_done(
                    published.url,
                    published.file_name,
                )
            )

        self._run(work)

    def _settings_for_export(self) -> GoogleSettings | None:
        """Взять сохранённые пути. Если их нет, открыть окно настроек."""
        settings = load_google_settings()
        if not settings.missing():
            return settings
        chosen = ask_google_settings(self, settings)
        if chosen is None:
            return None
        return chosen

    def _visible_table(self, kind: str) -> TableView:
        """Заголовки и строки, которые сейчас нарисованы. Другое окно не открывается."""
        tree = self._trees[kind]
        headers = tuple(str(tree.heading(column)["text"]) for column in tree["columns"])
        rows: list[tuple[str, ...]] = []
        amounts: list[tuple[str, str]] = []
        for iid in tree.get_children():
            values = tree.item(iid, "values")
            rows.append(tuple(str(value) for value in values))
            record = self._records[kind].get(iid)
            if kind in {"deals", "orders"} and isinstance(record, dict):
                amounts.append((_text(record.get("amount")), _text(record.get("currency"))))
        return TableView(
            section=kind,
            headers=headers,
            rows=tuple(rows),
            amounts=tuple(amounts),
        )

    def _export_done(self, url: str, file_name: str) -> None:
        """Показать ссылку и кнопку, которая открывает таблицу в браузере."""
        self._exporting = False
        self._set_status(f"Отчёт «{file_name}» записан.")
        ReportReadyDialog(self, file_name, url)

    def _export_failed(self, message: str) -> None:
        """Вернуть кнопку в работу и показать текст ошибки Google."""
        self._exporting = False
        self._set_status(message)
        messagebox.showerror("Мини-CRM", message, parent=self)

    def _tree(
        self,
        frame: ttk.Frame,
        kind: str,
        columns: tuple[tuple[str, str, int], ...],
        on_open: Callable[[], None],
    ) -> ttk.Treeview:
        """Таблица с вертикальной прокруткой."""
        holder = ttk.Frame(frame)
        holder.pack(fill="both", expand=True)
        holder.rowconfigure(0, weight=1)
        holder.columnconfigure(0, weight=1)
        tree = ttk.Treeview(
            holder,
            columns=[key for key, _title, _width in columns],
            show="headings",
            selectmode="browse",
        )
        for key, title, width in columns:
            tree.heading(key, text=title)
            tree.column(key, width=width, anchor="w", stretch=True)
        tree.tag_configure("muted", foreground=_MUTED)
        scroll = ttk.Scrollbar(holder, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        tree.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")

        def open_row(event: tk.Event[Any]) -> None:
            if tree.identify_region(event.x, event.y) != "cell":
                return
            on_open()

        tree.bind("<Double-1>", open_row)
        tree.bind("<Delete>", lambda _event, name=kind: self._delete_kind(name))
        return tree

    def _ui(self, callback: Callable[[], None]) -> None:
        """Отложить действие до главного потока. Из сети Tk вызывать нельзя."""
        self._ui_queue.put(callback)

    def _drain_ui(self) -> None:
        """Выполнить ответы сервера, которые накопились в очереди."""
        try:
            alive = bool(self.winfo_exists())
        except tk.TclError:
            return
        if not alive:
            return
        try:
            while True:
                callback = self._ui_queue.get_nowait()
                callback()
        except queue.Empty:
            pass
        self.after(50, self._drain_ui)

    def _run(self, work: Callable[[], None]) -> None:
        """Сеть не блокирует окно. Ошибки вне ApiError попадают в строку состояния."""
        def wrapped() -> None:
            try:
                work()
            except ApiError:
                raise
            except Exception as error:
                self._ui(lambda error=error: self._set_status(f"Сбой: {error}"))

        threading.Thread(target=wrapped, daemon=True).start()

    def _set_status(self, text: str) -> None:
        """Подпись внизу окна."""
        self._status.set(text)

    def _current_kind(self) -> str:
        """Вкладка, которую сейчас видно."""
        if self._notebook is None:
            return "clients"
        current = self._notebook.select()
        for kind, tab_id in self._tab_ids.items():
            if tab_id == current:
                return kind
        return "clients"

    def _remember_status(self, kind: str, text: str) -> None:
        """Запомнить подпись вкладки и показать её, только если вкладка открыта."""
        self._page_status[kind] = text
        if self._current_kind() == kind:
            self._set_status(text)

    def _show_tab_status(self) -> None:
        """При смене вкладки вернуть её последнюю подпись."""
        text = self._page_status.get(self._current_kind())
        if text:
            self._set_status(text)

    def _connect(self, silent: bool) -> None:
        """Спросить здоровье, справочник и сводку, затем обновить таблицы."""
        self._set_status("Подключение…")

        def work() -> None:
            try:
                health = self.api.health()
                meta = self.api.meta()
                summary = self.api.summary()
            except ApiError as error:
                self._ui(lambda error=error: self._offline(error, silent))
                return
            self._ui(lambda: self._online(health, meta, summary))

        self._run(work)

    def _offline(self, error: ApiError, silent: bool) -> None:
        """Сервер не ответил."""
        self._set_status(error.message)
        if not silent:
            messagebox.showerror("Мини-CRM", error.message, parent=self)

    def _online(self, health: Any, meta: Any, summary: Any) -> None:
        """Принять справочник и загрузить все вкладки."""
        del health
        self._apply_meta(meta)
        self._apply_summary(summary)
        self._set_status("Сервер на связи.")
        for kind in ("clients", "deals", "orders", "tasks"):
            self._fetch_page(kind)

    def _apply_meta(self, meta: Any) -> None:
        """Подписи статусов берём с сервера, чтобы окно не расходилось с API."""
        if not isinstance(meta, dict):
            return
        clients = meta.get("clients")
        deals = meta.get("deals")
        orders = meta.get("orders")
        tasks = meta.get("tasks")
        money = meta.get("money")
        self._client_statuses = _pairs_from_meta(
            clients.get("statuses") if isinstance(clients, dict) else None,
            _CLIENT_STATUSES,
        )
        self._deal_statuses = _pairs_from_meta(
            deals.get("statuses") if isinstance(deals, dict) else None,
            _DEAL_STATUSES,
        )
        self._order_statuses = _pairs_from_meta(
            orders.get("statuses") if isinstance(orders, dict) else None,
            _ORDER_STATUSES,
        )
        self._task_kinds = _pairs_from_meta(
            tasks.get("kinds") if isinstance(tasks, dict) else None,
            _TASK_KINDS,
        )
        self._labels["clients"] = _label_map(self._client_statuses)
        self._labels["deals"] = _label_map(self._deal_statuses)
        self._labels["orders"] = _label_map(self._order_statuses)
        self._labels["tasks"] = _label_map(self._task_kinds)
        if isinstance(money, dict) and isinstance(money.get("minor_scale"), int):
            scale = money["minor_scale"]
            if scale > 0:
                self._scale = scale
        self._replace_filter("clients-status", (("", "Все"),) + self._client_statuses)
        self._replace_filter("deals-status", (("", "Все"),) + self._deal_statuses)
        self._replace_filter("orders-status", (("", "Все"),) + self._order_statuses)
        self._replace_filter("tasks-kind", (("", "Все"),) + self._task_kinds)

    def _replace_filter(self, key: str, options: tuple[tuple[str, str], ...]) -> None:
        """Обновить пункты фильтра, не сбрасывая «Все», если оно выбрано."""
        variable = self._filter_vars.get(key)
        if variable is None:
            return
        self._filter_codes[key] = {label: code for code, label in options}
        labels = [label for _code, label in options]
        combo = self._filter_widgets.get(key)
        if combo is not None:
            combo["values"] = labels
        if variable.get() not in labels and labels:
            variable.set(labels[0])

    def _apply_summary(self, summary: Any) -> None:
        """Короткая строка счётчиков."""
        if not isinstance(summary, dict):
            return
        clients = _count_sum(summary.get("clients"))
        deals = _count_sum(summary.get("deals"))
        orders = _count_sum(summary.get("orders"))
        open_tasks = summary.get("tasks_open")
        open_text = open_tasks if isinstance(open_tasks, int) else 0
        self._summary.set(
            f"Клиенты: {clients}   Сделки: {deals}   "
            f"Заказы: {orders}   Открытые задачи: {open_text}"
        )

    def _reload_summary(self) -> None:
        """Обновить шапку после записи."""
        def work() -> None:
            try:
                summary = self.api.summary()
            except ApiError as error:
                self._ui(lambda error=error: self._set_status(error.message))
                return
            self._ui(lambda: self._apply_summary(summary))

        self._run(work)

    def _start_server(self) -> None:
        """Открыть консоль с backend/api.py и подождать, пока порт ответит."""
        if self._server_process is not None and self._server_process.poll() is None:
            self._set_status("Сервер уже запускается.")
            return
        script = Path(__file__).resolve().parent.parent / "backend" / "api.py"
        if not script.is_file():
            messagebox.showerror("Мини-CRM", "Не найден backend\\api.py.", parent=self)
            return
        self._server_process = subprocess.Popen(
            [sys.executable, str(script)],
            cwd=str(script.parent.parent),
            creationflags=CREATE_NEW_CONSOLE,
        )
        self._poll_left = SERVER_POLL_ATTEMPTS
        self._set_status("Жду ответ сервера…")
        self.after(SERVER_POLL_MS, self._poll_server)

    def _poll_server(self) -> None:
        """Несколько попыток health после запуска процесса."""
        self._poll_left -= 1

        def work() -> None:
            try:
                self.api.health()
            except ApiError:
                self._ui(self._poll_again)
                return
            self._ui(lambda: self._connect(silent=True))

        self._run(work)

    def _poll_again(self) -> None:
        """Повторить ожидание или сообщить, что сервер не поднялся."""
        if self._poll_left > 0:
            self.after(SERVER_POLL_MS, self._poll_server)
            return
        self._set_status("Сервер не ответил. Посмотрите окно консоли.")

    def _on_filter(self, key: str) -> None:
        """Смена фильтра сразу перечитывает свою таблицу."""
        kind = key.split("-", 1)[0]
        self._fetch_page(kind, reset_page=True)

    def _filter_code(self, key: str) -> str | None:
        """Код выбранного фильтра. «Все» — None."""
        variable = self._filter_vars.get(key)
        if variable is None:
            return None
        code = self._filter_codes.get(key, {}).get(variable.get(), "")
        if code == "":
            return None
        return code

    def _shift_page(self, kind: str, delta: int) -> None:
        """Следующая или предыдущая сотня строк с сервера."""
        current = self._offsets.get(kind, 0)
        total = self._totals.get(kind, 0)
        nxt = current + delta
        if nxt < 0:
            return
        if total > 0 and nxt >= total:
            return
        self._offsets[kind] = nxt
        self._fetch_page(kind)

    def _fetch_page(self, kind: str, reset_page: bool = False) -> None:
        """Прочитать фильтры в этом потоке и запросить страницу."""
        if reset_page:
            self._offsets[kind] = 0
        query = self._search_vars[kind].get().strip()
        status = self._filter_code(f"{kind}-status")
        task_kind = self._filter_code("tasks-kind") if kind == "tasks" else None
        done = self._filter_code("tasks-done") if kind == "tasks" else None
        offset = self._offsets.get(kind, 0)
        if self._current_kind() == kind:
            self._set_status("Загрузка…")

        def work() -> None:
            try:
                page = self._load_page(kind, query, status, task_kind, done, offset)
            except ApiError as error:
                self._ui(lambda error=error: self._set_status(error.message))
                return
            self._ui(
                lambda page=page, offset=offset: self._show_page(
                    kind,
                    page,
                    query,
                    status,
                    offset,
                )
            )

        self._run(work)

    def _load_page(
        self,
        kind: str,
        query: str,
        status: str | None,
        task_kind: str | None,
        done: str | None,
        offset: int,
    ) -> Any:
        """Выбрать список или поиск. Значения уже сняты с виджетов."""
        if kind == "clients":
            if query:
                return self.api.search_clients(query, offset)
            return self.api.list_clients(status, offset)
        if kind == "deals":
            if query:
                return self.api.search_deals(query, offset)
            return self.api.list_deals(status, offset)
        if kind == "orders":
            if query:
                return self.api.search_orders(query, offset)
            return self.api.list_orders(status, offset)
        if query:
            return self._search_tasks_locally(query, task_kind, done, offset)
        return self.api.list_tasks(task_kind, done, offset)

    def _search_tasks_locally(
        self,
        query: str,
        task_kind: str | None,
        done: str | None,
        offset: int,
    ) -> dict[str, Any]:
        """У задач нет отдельного поиска: читаем страницы и оставляем совпадения."""
        needle = query.casefold().replace("ё", "е")
        matched: list[dict[str, Any]] = []
        api_offset = 0
        total = 0
        while True:
            page = self.api.list_tasks(task_kind, done, api_offset)
            if not isinstance(page, dict):
                break
            raw_total = page.get("total")
            total = raw_total if isinstance(raw_total, int) else total
            items = page.get("items")
            if not isinstance(items, list) or not items:
                break
            for item in items:
                if isinstance(item, dict) and needle in _task_haystack(item):
                    matched.append(item)
            api_offset += len(items)
            if api_offset >= total or api_offset >= PAGE_LIMIT * 50:
                break
        return {"items": matched[offset:offset + PAGE_LIMIT], "total": len(matched)}

    def _show_page(
        self,
        kind: str,
        page: Any,
        query: str,
        status: str | None,
        offset: int,
    ) -> None:
        """Заполнить таблицу. Поиск и фильтр вместе дофильтровываются здесь."""
        if offset != self._offsets.get(kind, 0):
            return
        if not isinstance(page, dict):
            self._remember_status(kind, "Сервер вернул пустую страницу.")
            return
        items = page.get("items")
        if not isinstance(items, list):
            self._remember_status(kind, "Сервер вернул пустую страницу.")
            return
        rows = [item for item in items if isinstance(item, dict)]
        if query and status is not None and kind != "tasks":
            rows = [row for row in rows if row.get("status") == status]
        kept = [row for row in rows if isinstance(row.get("id"), int)]
        self._records[kind] = {str(row["id"]): row for row in kept}
        if kind == "clients":
            self._people = list(self._records[kind].values())
        elif kind == "deals":
            self._deal_rows = list(self._records[kind].values())
        elif kind == "orders":
            self._order_rows = list(self._records[kind].values())
        tree = self._trees[kind]
        selected = tree.selection()
        tree.delete(*tree.get_children())
        for row in rows:
            if not isinstance(row.get("id"), int):
                continue
            iid = str(row["id"])
            tree.insert("", "end", iid=iid, values=self._cells(kind, row), tags=self._tags(row))
        if selected and tree.exists(selected[0]):
            tree.selection_set(selected[0])
        total = page.get("total")
        total_value = total if isinstance(total, int) else shown
        self._totals[kind] = total_value
        shown = len(tree.get_children())
        if shown == 0 and offset > 0:
            self._offsets[kind] = max(0, offset - PAGE_LIMIT)
            self._fetch_page(kind)
            return
        if shown == 0:
            self._remember_status(kind, f"Записей: 0 из {total_value}.")
            return
        start = offset + 1
        end = offset + shown
        self._remember_status(kind, f"Показаны {start}–{end} из {total_value}.")

    def _cells(self, kind: str, row: dict[str, Any]) -> tuple[str, ...]:
        """Значения строки таблицы."""
        if kind == "clients":
            return (
                str(row.get("id", "")),
                _text(row.get("full_name")),
                _text(row.get("company")),
                _text(row.get("phone")),
                _text(row.get("email")),
                self._labels["clients"].get(_text(row.get("status")), _text(row.get("status"))),
                _text(row.get("source")),
            )
        if kind == "deals":
            return (
                str(row.get("id", "")),
                _text(row.get("title")),
                _text(row.get("client_name")),
                self._labels["deals"].get(_text(row.get("status")), _text(row.get("status"))),
                _money_text(row),
                _text(row.get("expected_close_on")),
            )
        if kind == "orders":
            return (
                str(row.get("id", "")),
                _text(row.get("title")),
                _text(row.get("client_name")),
                _text(row.get("deal_title")),
                self._labels["orders"].get(_text(row.get("status")), _text(row.get("status"))),
                _money_text(row),
                _text(row.get("due_on")),
            )
        done = "Да" if row.get("is_done") is True else "Нет"
        return (
            str(row.get("id", "")),
            _text(row.get("title")),
            self._labels["tasks"].get(_text(row.get("kind")), _text(row.get("kind"))),
            _text(row.get("due_at")),
            done,
            _text(row.get("client_name")),
            _text(row.get("deal_title")),
            _text(row.get("order_title")),
        )

    def _tags(self, row: dict[str, Any]) -> tuple[str, ...]:
        """Архив и выполненные задачи рисуются спокойнее."""
        if row.get("status") == "archived" or row.get("is_done") is True:
            return ("muted",)
        return ()

    def _selected(self, kind: str) -> dict[str, Any] | None:
        """Выбранная строка или подсказка, если кликнули мимо."""
        tree = self._trees[kind]
        chosen = tree.selection()
        if not chosen:
            messagebox.showinfo("Мини-CRM", "Сначала выберите строку.", parent=self)
            return None
        row = self._records[kind].get(chosen[0])
        if row is None:
            messagebox.showinfo(
                "Мини-CRM",
                "Строка уже не в списке. Обновите таблицу.",
                parent=self,
            )
            return None
        return row

    def _people_options(self, current: dict[str, Any] | None) -> tuple[tuple[str, str], ...]:
        """Клиенты для выпадающего списка связи."""
        options = [_NONE_OPTION]
        for row in self._people:
            if isinstance(row.get("id"), int):
                options.append((str(row["id"]), _text(row.get("full_name"))))
        base = tuple(options)
        if current is None:
            return base
        return _with_current(base, current.get("client_id"), current.get("client_name"))

    def _deal_options(self, current: dict[str, Any] | None) -> tuple[tuple[str, str], ...]:
        """Сделки для связи заказа и задачи."""
        options = [_NONE_OPTION]
        for row in self._deal_rows:
            if isinstance(row.get("id"), int):
                options.append((str(row["id"]), _text(row.get("title"))))
        base = tuple(options)
        if current is None:
            return base
        return _with_current(base, current.get("deal_id"), current.get("deal_title"))

    def _order_options(self, current: dict[str, Any] | None) -> tuple[tuple[str, str], ...]:
        """Заказы для связи задачи."""
        options = [_NONE_OPTION]
        for row in self._order_rows:
            if isinstance(row.get("id"), int):
                options.append((str(row["id"]), _text(row.get("title"))))
        base = tuple(options)
        if current is None:
            return base
        return _with_current(base, current.get("order_id"), current.get("order_title"))

    def _open_editor(
        self,
        title: str,
        specs: tuple[FieldSpec, ...],
        initial: dict[str, Any],
        on_submit: Callable[[dict[str, Any], Callable[[ApiError | None], None]], None],
    ) -> None:
        """Показать форму поверх главного окна."""
        Editor(self, title, specs, initial, on_submit)

    def _saved(self, finish: Callable[[ApiError | None], None], kind: str) -> None:
        """Закрыть форму, обновить таблицу и шапку."""
        finish(None)
        self._fetch_page(kind)
        self._reload_summary()

    def _create_client(self) -> None:
        """Окно нового клиента."""
        self._open_editor(
            "Новый клиент",
            self._client_fields(),
            {"status": "active"},
            lambda raw, finish: self._save_client(None, raw, finish),
        )

    def _edit_client(self) -> None:
        """Окно правки выбранного клиента."""
        row = self._selected("clients")
        if row is None:
            return
        self._open_editor(
            "Клиент",
            self._client_fields(),
            row,
            lambda raw, finish, client_id=row["id"]: self._save_client(client_id, raw, finish),
        )

    def _client_fields(self) -> tuple[FieldSpec, ...]:
        return (
            FieldSpec("full_name", "Имя", "entry"),
            FieldSpec("company", "Компания", "entry"),
            FieldSpec("phone", "Телефон", "entry"),
            FieldSpec("email", "Почта", "entry"),
            FieldSpec("status", "Статус", "combo", self._client_statuses),
            FieldSpec("source", "Источник", "entry"),
            FieldSpec("notes", "Заметки", "text"),
        )

    def _save_client(
        self,
        client_id: int | None,
        raw: dict[str, Any],
        finish: Callable[[ApiError | None], None],
    ) -> None:
        """Проверить форму и отправить карточку клиента."""
        try:
            body = build_client_body(raw)
        except ApiError as error:
            finish(error)
            return

        def work() -> None:
            try:
                if client_id is None:
                    self.api.create_client(body)
                else:
                    self.api.update_client(client_id, body)
            except ApiError as error:
                self._ui(lambda error=error: finish(error))
                return
            self._ui(lambda: self._saved(finish, "clients"))

        self._run(work)

    def _status_client(self) -> None:
        """Маленькое окно смены статуса, в том числе архива."""
        row = self._selected("clients")
        if row is None:
            return
        self._open_status(
            "Статус клиента",
            self._client_statuses,
            _text(row.get("status")),
            lambda code, finish, client_id=row["id"]: self._patch_status(
                "clients",
                lambda status: self.api.set_client_status(client_id, status),
                code,
                finish,
            ),
        )

    def _delete_client(self) -> None:
        """Удалить клиента после подтверждения."""
        row = self._selected("clients")
        if row is None:
            return
        self._confirm_delete(
            _text(row.get("full_name")),
            lambda finish, client_id=row["id"]: self._remove(
                "clients",
                lambda: self.api.delete_client(client_id),
                finish,
            ),
        )

    def _create_deal(self) -> None:
        """Окно новой сделки. Клиента можно не выбирать."""
        self._open_editor(
            "Новая сделка",
            self._deal_fields(None),
            {"status": "new", "currency": "RUB", "client_id": ""},
            lambda raw, finish: self._save_deal(None, raw, finish),
        )

    def _edit_deal(self) -> None:
        """Окно правки сделки. Здесь же меняется прикреплённый клиент."""
        row = self._selected("deals")
        if row is None:
            return
        initial = dict(row)
        initial["amount"] = _text(row.get("amount"))
        initial["client_id"] = "" if row.get("client_id") is None else str(row.get("client_id"))
        self._open_editor(
            "Сделка",
            self._deal_fields(row),
            initial,
            lambda raw, finish, deal_id=row["id"]: self._save_deal(deal_id, raw, finish),
        )

    def _deal_fields(self, current: dict[str, Any] | None) -> tuple[FieldSpec, ...]:
        return (
            FieldSpec("title", "Название", "entry"),
            FieldSpec("client_id", "Клиент", "combo", self._people_options(current)),
            FieldSpec("status", "Статус", "combo", self._deal_statuses),
            FieldSpec("amount", "Сумма", "entry"),
            FieldSpec("currency", "Валюта", "entry"),
            FieldSpec("expected_close_on", "Дата (ГГГГ-ММ-ДД)", "entry"),
            FieldSpec("notes", "Заметки", "text"),
        )

    def _save_deal(
        self,
        deal_id: int | None,
        raw: dict[str, Any],
        finish: Callable[[ApiError | None], None],
    ) -> None:
        """Проверить форму и отправить сделку."""
        try:
            body = build_deal_body(raw, self._scale)
        except ApiError as error:
            finish(error)
            return

        def work() -> None:
            try:
                if deal_id is None:
                    self.api.create_deal(body)
                else:
                    self.api.update_deal(deal_id, body)
            except ApiError as error:
                self._ui(lambda error=error: finish(error))
                return
            self._ui(lambda: self._saved(finish, "deals"))

        self._run(work)

    def _status_deal(self) -> None:
        """Смена статуса сделки."""
        row = self._selected("deals")
        if row is None:
            return
        self._open_status(
            "Статус сделки",
            self._deal_statuses,
            _text(row.get("status")),
            lambda code, finish, deal_id=row["id"]: self._patch_status(
                "deals",
                lambda status: self.api.set_deal_status(deal_id, status),
                code,
                finish,
            ),
        )

    def _delete_deal(self) -> None:
        """Удалить сделку после подтверждения."""
        row = self._selected("deals")
        if row is None:
            return
        self._confirm_delete(
            _text(row.get("title")),
            lambda finish, deal_id=row["id"]: self._remove(
                "deals",
                lambda: self.api.delete_deal(deal_id),
                finish,
            ),
        )

    def _create_order(self) -> None:
        """Окно нового заказа."""
        self._open_editor(
            "Новый заказ",
            self._order_fields(None),
            {"status": "draft", "currency": "RUB", "client_id": "", "deal_id": ""},
            lambda raw, finish: self._save_order(None, raw, finish),
        )

    def _edit_order(self) -> None:
        """Окно правки заказа."""
        row = self._selected("orders")
        if row is None:
            return
        initial = dict(row)
        initial["amount"] = _text(row.get("amount"))
        initial["client_id"] = "" if row.get("client_id") is None else str(row.get("client_id"))
        initial["deal_id"] = "" if row.get("deal_id") is None else str(row.get("deal_id"))
        self._open_editor(
            "Заказ",
            self._order_fields(row),
            initial,
            lambda raw, finish, order_id=row["id"]: self._save_order(order_id, raw, finish),
        )

    def _order_fields(self, current: dict[str, Any] | None) -> tuple[FieldSpec, ...]:
        return (
            FieldSpec("title", "Название", "entry"),
            FieldSpec("client_id", "Клиент", "combo", self._people_options(current)),
            FieldSpec("deal_id", "Сделка", "combo", self._deal_options(current)),
            FieldSpec("status", "Статус", "combo", self._order_statuses),
            FieldSpec("amount", "Сумма", "entry"),
            FieldSpec("currency", "Валюта", "entry"),
            FieldSpec("due_on", "Срок (ГГГГ-ММ-ДД)", "entry"),
            FieldSpec("notes", "Заметки", "text"),
        )

    def _save_order(
        self,
        order_id: int | None,
        raw: dict[str, Any],
        finish: Callable[[ApiError | None], None],
    ) -> None:
        """Проверить форму и отправить заказ."""
        try:
            body = build_order_body(raw, self._scale)
        except ApiError as error:
            finish(error)
            return

        def work() -> None:
            try:
                if order_id is None:
                    self.api.create_order(body)
                else:
                    self.api.update_order(order_id, body)
            except ApiError as error:
                self._ui(lambda error=error: finish(error))
                return
            self._ui(lambda: self._saved(finish, "orders"))

        self._run(work)

    def _status_order(self) -> None:
        """Смена статуса заказа."""
        row = self._selected("orders")
        if row is None:
            return
        self._open_status(
            "Статус заказа",
            self._order_statuses,
            _text(row.get("status")),
            lambda code, finish, order_id=row["id"]: self._patch_status(
                "orders",
                lambda status: self.api.set_order_status(order_id, status),
                code,
                finish,
            ),
        )

    def _delete_order(self) -> None:
        """Удалить заказ после подтверждения."""
        row = self._selected("orders")
        if row is None:
            return
        self._confirm_delete(
            _text(row.get("title")),
            lambda finish, order_id=row["id"]: self._remove(
                "orders",
                lambda: self.api.delete_order(order_id),
                finish,
            ),
        )

    def _create_task(self) -> None:
        """Окно новой задачи или напоминания."""
        self._open_editor(
            "Новая задача",
            self._task_fields(None),
            {"kind": "task", "is_done": False, "client_id": "", "deal_id": "", "order_id": ""},
            lambda raw, finish: self._save_task(None, raw, finish),
        )

    def _edit_task(self) -> None:
        """Окно правки задачи."""
        row = self._selected("tasks")
        if row is None:
            return
        initial = dict(row)
        initial["client_id"] = "" if row.get("client_id") is None else str(row.get("client_id"))
        initial["deal_id"] = "" if row.get("deal_id") is None else str(row.get("deal_id"))
        initial["order_id"] = "" if row.get("order_id") is None else str(row.get("order_id"))
        self._open_editor(
            "Задача",
            self._task_fields(row),
            initial,
            lambda raw, finish, task_id=row["id"]: self._save_task(task_id, raw, finish),
        )

    def _task_fields(self, current: dict[str, Any] | None) -> tuple[FieldSpec, ...]:
        return (
            FieldSpec("title", "Название", "entry"),
            FieldSpec("kind", "Вид", "combo", self._task_kinds),
            FieldSpec("due_at", "Срок (ISO или ГГГГ-ММ-ДД)", "entry"),
            FieldSpec("is_done", "Выполнена", "check"),
            FieldSpec("client_id", "Клиент", "combo", self._people_options(current)),
            FieldSpec("deal_id", "Сделка", "combo", self._deal_options(current)),
            FieldSpec("order_id", "Заказ", "combo", self._order_options(current)),
            FieldSpec("details", "Описание", "text"),
        )

    def _save_task(
        self,
        task_id: int | None,
        raw: dict[str, Any],
        finish: Callable[[ApiError | None], None],
    ) -> None:
        """Проверить форму и отправить задачу."""
        try:
            body = build_task_body(raw)
        except ApiError as error:
            finish(error)
            return

        def work() -> None:
            try:
                if task_id is None:
                    self.api.create_task(body)
                else:
                    self.api.update_task(task_id, body)
            except ApiError as error:
                self._ui(lambda error=error: finish(error))
                return
            self._ui(lambda: self._saved(finish, "tasks"))

        self._run(work)

    def _toggle_task(self) -> None:
        """Переключить отметку выполнения у выбранной задачи."""
        row = self._selected("tasks")
        if row is None:
            return
        target = row.get("is_done") is not True
        task_id = row["id"]

        def work() -> None:
            try:
                self.api.set_task_done(task_id, target)
            except ApiError as error:
                self._ui(lambda error=error: self._fail(error))
                return
            self._ui(lambda: self._saved(lambda _error: None, "tasks"))

        self._run(work)

    def _delete_task(self) -> None:
        """Удалить задачу после подтверждения."""
        row = self._selected("tasks")
        if row is None:
            return
        self._confirm_delete(
            _text(row.get("title")),
            lambda finish, task_id=row["id"]: self._remove(
                "tasks",
                lambda: self.api.delete_task(task_id),
                finish,
            ),
        )

    def _delete_kind(self, kind: str) -> None:
        """Клавиша Delete вызывает удаление на текущей вкладке."""
        if kind == "clients":
            self._delete_client()
            return
        if kind == "deals":
            self._delete_deal()
            return
        if kind == "orders":
            self._delete_order()
            return
        self._delete_task()

    def _open_status(
        self,
        title: str,
        options: tuple[tuple[str, str], ...],
        current: str,
        on_submit: Callable[[str, Callable[[ApiError | None], None]], None],
    ) -> None:
        """Окно из одного списка статусов."""
        def submit(raw: dict[str, Any], finish: Callable[[ApiError | None], None]) -> None:
            code = plain_text(raw.get("status"))
            if code == "":
                finish(ApiError("VALIDATION_ERROR", "Выберите статус.", None))
                return
            on_submit(code, finish)

        self._open_editor(
            title,
            (FieldSpec("status", "Статус", "combo", options),),
            {"status": current},
            submit,
        )

    def _patch_status(
        self,
        kind: str,
        call: Callable[[str], Any],
        code: str,
        finish: Callable[[ApiError | None], None],
    ) -> None:
        """Отправить новый статус и обновить таблицу."""
        def work() -> None:
            try:
                call(code)
            except ApiError as error:
                self._ui(lambda error=error: finish(error))
                return
            self._ui(lambda: self._saved(finish, kind))

        self._run(work)

    def _confirm_delete(
        self,
        name: str,
        action: Callable[[Callable[[ApiError | None], None]], None],
    ) -> None:
        """Спросить подтверждение и только потом удалять."""
        if not messagebox.askyesno("Удаление", f"Удалить «{name}»?", parent=self):
            return
        action(self._fail)

    def _remove(
        self,
        kind: str,
        call: Callable[[], Any],
        finish: Callable[[ApiError | None], None],
    ) -> None:
        """Удалить запись на сервере."""
        def work() -> None:
            try:
                call()
            except ApiError as error:
                self._ui(lambda error=error: finish(error))
                return
            self._ui(lambda: self._saved(finish, kind))

        self._run(work)

    def _fail(self, error: ApiError | None) -> None:
        """Показать ошибку действия, у которого нет своей формы."""
        if error is None:
            return
        self._set_status(error.message)
        messagebox.showerror("Мини-CRM", error.message, parent=self)


def _count_sum(value: object) -> int:
    """Сложить счётчики статусов из сводки."""
    if not isinstance(value, dict):
        return 0
    total = 0
    for item in value.values():
        if isinstance(item, int):
            total += item
    return total


def _text(value: object) -> str:
    """Текст ячейки. Числа и пустые значения не роняют таблицу."""
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    return str(value)


def _money_text(row: dict[str, Any]) -> str:
    """Сумма так, как её уже посчитал сервер, плюс валюта."""
    amount = _text(row.get("amount"))
    currency = _text(row.get("currency"))
    if amount == "":
        return currency
    if currency == "":
        return amount
    return f"{amount} {currency}"


def _task_haystack(row: dict[str, Any]) -> str:
    """Текст задачи для локального поиска без учёта регистра и буквы ё."""
    parts = (
        _text(row.get("title")),
        _text(row.get("details")),
        _text(row.get("client_name")),
        _text(row.get("deal_title")),
        _text(row.get("order_title")),
    )
    return " ".join(parts).casefold().replace("ё", "е")


def main() -> None:
    """Открыть окно."""
    app = CrmApp()
    app.mainloop()


if __name__ == "__main__":
    main()
