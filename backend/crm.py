"""Класс SQLite для мини-CRM: схема и CRUD.

Файл базы задаётся при создании класса. Таблицы из models.py
создаются, если их ещё нет. Каждая операция открывает своё соединение,
поэтому путь :memory: не подходит.

Поиска ILIKE в SQLite нет. При открытии соединения регистрируется
функция ilike: шаблон LIKE без учёта регистра, включая кириллицу и ё/е.
"""

from __future__ import annotations

import logging
import os
import re
import sqlite3
from collections.abc import Callable, Iterator, Mapping
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, TypeVar

from models import (
    CLIENT_STATUSES,
    COMPANY_MAX_LENGTH,
    DEAL_STATUSES,
    DEFAULT_PAGE_LIMIT,
    DETAILS_MAX_LENGTH,
    EMAIL_MAX_LENGTH,
    MAX_AMOUNT_MINOR,
    MAX_PAGE_LIMIT,
    NAME_MAX_LENGTH,
    NOTES_MAX_LENGTH,
    ORDER_STATUSES,
    PHONE_MAX_LENGTH,
    QUERY_MAX_LENGTH,
    SOURCE_MAX_LENGTH,
    TABLES,
    TASK_KINDS,
    TITLE_MAX_LENGTH,
    Client,
    ClientInput,
    Deal,
    DealInput,
    Order,
    OrderInput,
    Page,
    Task,
    TaskInput,
    TaskKind,
)

logger = logging.getLogger("crm")
if logger.level == logging.NOTSET:
    logger.setLevel(logging.INFO)

T = TypeVar("T")

_SQLITE_TIMEOUT_SECONDS = 5.0
_KNOWN_TABLES = frozenset({"clients", "deals", "orders", "tasks"})
_EMAIL_PATTERN = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
_PHONE_PATTERN = re.compile(r"^[0-9+\-() ]+$")
_CURRENCY_PATTERN = re.compile(r"^[A-Z]{3}$")
_ORDER_BY_SQL = re.compile(
    r"^[a-z_]+(\.[a-z_]+)?( ASC| DESC)?"
    r"(, [a-z_]+(\.[a-z_]+)?( ASC| DESC)?)*$"
)
_ILIKE_PART = re.compile(r"^ilike\([a-z_]+\.[a-z_]+, \?\)$")

_CLIENT_SEARCH = (
    "ilike(clients.full_name, ?)",
    "ilike(clients.company, ?)",
    "ilike(clients.email, ?)",
    "ilike(clients.phone, ?)",
    "ilike(clients.source, ?)",
    "ilike(clients.notes, ?)",
)
_DEAL_SEARCH = (
    "ilike(deals.title, ?)",
    "ilike(deals.notes, ?)",
    "ilike(deals.status, ?)",
    "ilike(clients.full_name, ?)",
    "ilike(clients.company, ?)",
    "ilike(clients.phone, ?)",
    "ilike(clients.email, ?)",
)
_ORDER_SEARCH = (
    "ilike(orders.title, ?)",
    "ilike(orders.notes, ?)",
    "ilike(orders.status, ?)",
    "ilike(clients.full_name, ?)",
    "ilike(clients.company, ?)",
    "ilike(deals.title, ?)",
)

_DEAL_FROM = "FROM deals LEFT JOIN clients ON clients.id = deals.client_id"
_DEAL_SELECT = f"""
SELECT
    deals.id,
    deals.title,
    deals.client_id,
    deals.status,
    deals.amount_minor,
    deals.currency,
    deals.expected_close_on,
    deals.notes,
    deals.created_at,
    deals.updated_at,
    COALESCE(clients.full_name, '') AS client_name
{_DEAL_FROM}
""".strip()

_ORDER_FROM = """
FROM orders
LEFT JOIN clients ON clients.id = orders.client_id
LEFT JOIN deals ON deals.id = orders.deal_id
""".strip()
_ORDER_SELECT = f"""
SELECT
    orders.id,
    orders.title,
    orders.client_id,
    orders.deal_id,
    orders.status,
    orders.amount_minor,
    orders.currency,
    orders.due_on,
    orders.notes,
    orders.created_at,
    orders.updated_at,
    COALESCE(clients.full_name, '') AS client_name,
    COALESCE(deals.title, '') AS deal_title
{_ORDER_FROM}
""".strip()

_TASK_FROM = """
FROM tasks
LEFT JOIN clients ON clients.id = tasks.client_id
LEFT JOIN deals ON deals.id = tasks.deal_id
LEFT JOIN orders ON orders.id = tasks.order_id
""".strip()
_TASK_SELECT = f"""
SELECT
    tasks.id,
    tasks.title,
    tasks.details,
    tasks.kind,
    tasks.due_at,
    tasks.is_done,
    tasks.client_id,
    tasks.deal_id,
    tasks.order_id,
    tasks.created_at,
    tasks.updated_at,
    tasks.completed_at,
    COALESCE(clients.full_name, '') AS client_name,
    COALESCE(deals.title, '') AS deal_title,
    COALESCE(orders.title, '') AS order_title
{_TASK_FROM}
""".strip()


class CrmError(Exception):
    """Ошибка CRM в формате code / message / details."""

    def __init__(self, code: str, message: str, details: Any = None) -> None:
        self.code = code
        self.message = message
        self.details = details
        super().__init__(message)

    def to_dict(self) -> dict[str, Any]:
        """Словарь для ответа API и журнала."""
        return {
            "code": self.code,
            "message": self.message,
            "details": self.details,
        }


def default_db_path() -> Path:
    """Файл базы: CRM_DB_PATH или data/crm.sqlite рядом с этим модулем."""
    raw = os.environ.get("CRM_DB_PATH", "").strip()
    if raw:
        return Path(raw)
    return Path(__file__).resolve().parent / "data" / "crm.sqlite"


def _log_event(message: str, context: dict[str, Any]) -> None:
    """Журнал без телефона, почты и заметок: только действие и идентификаторы."""
    logger.info("%s", {"level": "info", "message": message, "context": context})


def _utc_now() -> str:
    """Текущее время UTC без микросекунд, чтобы строки дат сортировались."""
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def _fold_text(value: str) -> str:
    """Сложить регистр и считать ё и е одной буквой."""
    return value.casefold().replace("ё", "е")


def _like_to_regex(pattern: str) -> re.Pattern[str]:
    """Перевести шаблон LIKE (% и _, экранирование \\) в регулярное выражение."""
    parts: list[str] = ["^"]
    index = 0
    length = len(pattern)
    while index < length:
        char = pattern[index]
        if char == "\\" and index + 1 < length:
            parts.append(re.escape(pattern[index + 1]))
            index += 2
            continue
        if char == "%":
            parts.append(".*")
        elif char == "_":
            parts.append(".")
        else:
            parts.append(re.escape(char))
        index += 1
    parts.append("$")
    return re.compile("".join(parts), re.DOTALL)


def _sqlite_ilike(value: object, pattern: object) -> int:
    """Функция SQL ilike(значение, шаблон). 1 — совпало, 0 — нет."""
    if not isinstance(value, str) or not isinstance(pattern, str):
        return 0
    expression = _like_to_regex(_fold_text(pattern))
    if expression.fullmatch(_fold_text(value)) is None:
        return 0
    return 1


def _contains_pattern(query: str) -> str:
    """Шаблон «содержит» с экранированием % , _ и обратной косой."""
    escaped = query.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")
    return f"%{escaped}%"


def _text(value: object, field: str, max_length: int, *, required: bool) -> str:
    if not isinstance(value, str):
        raise CrmError(
            "VALIDATION_ERROR",
            f"Поле «{field}» должно быть строкой.",
            {"field": field},
        )
    if "\x00" in value:
        raise CrmError(
            "VALIDATION_ERROR",
            f"Поле «{field}» содержит недопустимый символ.",
            {"field": field},
        )
    cleaned = value.strip()
    if required and cleaned == "":
        raise CrmError(
            "VALIDATION_ERROR",
            f"Заполните поле «{field}».",
            {"field": field},
        )
    if len(cleaned) > max_length:
        raise CrmError(
            "VALIDATION_ERROR",
            f"Поле «{field}» длиннее {max_length} символов.",
            {"field": field, "max_length": max_length},
        )
    return cleaned


def _choice(value: object, allowed: tuple[str, ...], field: str) -> str:
    cleaned = _text(value, field, max_length=32, required=True)
    if cleaned not in allowed:
        raise CrmError(
            "VALIDATION_ERROR",
            f"Недопустимое значение поля «{field}».",
            {"field": field, "allowed": list(allowed)},
        )
    return cleaned


def _positive_id(value: object, label: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < 1:
        raise CrmError(
            "VALIDATION_ERROR",
            f"Некорректный идентификатор {label}.",
            {"id": value},
        )
    return value


def _optional_id(value: object, label: str) -> int | None:
    if value is None:
        return None
    return _positive_id(value, label)


def _amount(value: object) -> int:
    if isinstance(value, bool) or not isinstance(value, int):
        raise CrmError(
            "VALIDATION_ERROR",
            "Сумма должна быть целым числом в мелких единицах.",
            {"amount_minor": value},
        )
    if value < 0 or value > MAX_AMOUNT_MINOR:
        raise CrmError(
            "VALIDATION_ERROR",
            "Сумма выходит за допустимые пределы.",
            {"amount_minor": value, "max": MAX_AMOUNT_MINOR},
        )
    return value


def _currency(value: object) -> str:
    if not isinstance(value, str):
        raise CrmError(
            "VALIDATION_ERROR",
            "Валюта должна быть строкой.",
            {"field": "currency"},
        )
    cleaned = value.strip().upper()
    if _CURRENCY_PATTERN.fullmatch(cleaned) is None:
        raise CrmError(
            "VALIDATION_ERROR",
            "Валюта — это трёхбуквенный код, например RUB.",
            {"currency": value},
        )
    return cleaned


def _email(value: object) -> str:
    cleaned = _text(value, "почта", EMAIL_MAX_LENGTH, required=False)
    if cleaned and _EMAIL_PATTERN.fullmatch(cleaned) is None:
        raise CrmError(
            "VALIDATION_ERROR",
            "Почта выглядит некорректно.",
            {"email": cleaned},
        )
    return cleaned


def _phone(value: object) -> str:
    cleaned = _text(value, "телефон", PHONE_MAX_LENGTH, required=False)
    if cleaned and _PHONE_PATTERN.fullmatch(cleaned) is None:
        raise CrmError(
            "VALIDATION_ERROR",
            "В телефоне допустимы цифры, пробел, +, - и скобки.",
            {"phone": cleaned},
        )
    return cleaned


def _iso_date(value: object, field: str) -> str:
    if not isinstance(value, str):
        raise CrmError(
            "VALIDATION_ERROR",
            f"Поле «{field}» должно быть строкой.",
            {"field": field},
        )
    cleaned = value.strip()
    if cleaned == "":
        return ""
    try:
        parsed = datetime.strptime(cleaned, "%Y-%m-%d")
    except ValueError as error:
        raise CrmError(
            "VALIDATION_ERROR",
            f"Поле «{field}» ждёт дату в формате ГГГГ-ММ-ДД.",
            {"field": field},
        ) from error
    return parsed.date().isoformat()


def _iso_datetime(value: object, field: str) -> str:
    if not isinstance(value, str):
        raise CrmError(
            "VALIDATION_ERROR",
            f"Поле «{field}» должно быть строкой.",
            {"field": field},
        )
    cleaned = value.strip()
    if cleaned == "":
        return ""
    try:
        parsed = datetime.fromisoformat(cleaned.replace("Z", "+00:00"))
    except ValueError as error:
        raise CrmError(
            "VALIDATION_ERROR",
            f"Поле «{field}» ждёт дату и время в формате ISO.",
            {"field": field},
        ) from error
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)
    return parsed.astimezone(timezone.utc).replace(microsecond=0).isoformat()


def _page_bounds(limit: int, offset: int) -> tuple[int, int]:
    if isinstance(limit, bool) or not isinstance(limit, int):
        raise CrmError(
            "VALIDATION_ERROR",
            "Параметр limit должен быть целым числом.",
            {"limit": limit},
        )
    if limit < 1 or limit > MAX_PAGE_LIMIT:
        raise CrmError(
            "VALIDATION_ERROR",
            f"Параметр limit должен быть от 1 до {MAX_PAGE_LIMIT}.",
            {"limit": limit},
        )
    if isinstance(offset, bool) or not isinstance(offset, int) or offset < 0:
        raise CrmError(
            "VALIDATION_ERROR",
            "Параметр offset должен быть целым числом от 0.",
            {"offset": offset},
        )
    return limit, offset


def _flag_value(value: object, field: str) -> bool:
    if not isinstance(value, bool):
        raise CrmError(
            "VALIDATION_ERROR",
            f"Поле «{field}» должно быть логическим.",
            {"field": field},
        )
    return value


def _completed_at(existing: Task | None, is_done: bool, now: str) -> str:
    """Сохранить момент первого выполнения и стереть его, если задачу открыли."""
    if not is_done:
        return ""
    if existing is not None and existing.is_done and existing.completed_at:
        return existing.completed_at
    return now


def _normalize_client(payload: ClientInput) -> ClientInput:
    return ClientInput(
        full_name=_text(payload.full_name, "имя", NAME_MAX_LENGTH, required=True),
        company=_text(payload.company, "компания", COMPANY_MAX_LENGTH, required=False),
        email=_email(payload.email),
        phone=_phone(payload.phone),
        status=_choice(payload.status, CLIENT_STATUSES, "статус клиента"),
        source=_text(payload.source, "источник", SOURCE_MAX_LENGTH, required=False),
        notes=_text(payload.notes, "заметки", NOTES_MAX_LENGTH, required=False),
    )


def _normalize_deal(payload: DealInput) -> DealInput:
    return DealInput(
        title=_text(payload.title, "название сделки", TITLE_MAX_LENGTH, required=True),
        client_id=_optional_id(payload.client_id, "клиента"),
        status=_choice(payload.status, DEAL_STATUSES, "статус сделки"),
        amount_minor=_amount(payload.amount_minor),
        currency=_currency(payload.currency),
        expected_close_on=_iso_date(payload.expected_close_on, "ожидаемая дата"),
        notes=_text(payload.notes, "заметки сделки", NOTES_MAX_LENGTH, required=False),
    )


def _normalize_order(payload: OrderInput) -> OrderInput:
    return OrderInput(
        title=_text(payload.title, "название заказа", TITLE_MAX_LENGTH, required=True),
        client_id=_optional_id(payload.client_id, "клиента"),
        deal_id=_optional_id(payload.deal_id, "сделки"),
        status=_choice(payload.status, ORDER_STATUSES, "статус заказа"),
        amount_minor=_amount(payload.amount_minor),
        currency=_currency(payload.currency),
        due_on=_iso_date(payload.due_on, "срок заказа"),
        notes=_text(payload.notes, "заметки заказа", NOTES_MAX_LENGTH, required=False),
    )


def _normalize_task(payload: TaskInput) -> TaskInput:
    kind = _choice(payload.kind, TASK_KINDS, "вид задачи")
    due_at = _iso_datetime(payload.due_at, "срок")
    if kind == TaskKind.REMINDER.value and due_at == "":
        raise CrmError(
            "VALIDATION_ERROR",
            "У напоминания укажите срок.",
            {"field": "due_at"},
        )
    return TaskInput(
        title=_text(payload.title, "название задачи", TITLE_MAX_LENGTH, required=True),
        details=_text(payload.details, "описание", DETAILS_MAX_LENGTH, required=False),
        kind=kind,
        due_at=due_at,
        is_done=_flag_value(payload.is_done, "выполнена"),
        client_id=_optional_id(payload.client_id, "клиента"),
        deal_id=_optional_id(payload.deal_id, "сделки"),
        order_id=_optional_id(payload.order_id, "заказа"),
    )


def _ilike_clause(parts: tuple[str, ...]) -> str:
    """Собрать OR из заранее заданных вызовов ilike, не из текста пользователя."""
    checked: list[str] = []
    for part in parts:
        if _ILIKE_PART.fullmatch(part) is None:
            raise CrmError("DATABASE_ERROR", "Некорректное условие поиска.", None)
        checked.append(part)
    return "(" + " OR ".join(checked) + ")"


def _translate_integrity(error: sqlite3.IntegrityError) -> CrmError:
    reason = str(error)
    lowered = reason.lower()
    if "foreign key" in lowered:
        return CrmError(
            "VALIDATION_ERROR",
            "Связанная запись не найдена.",
            {"reason": reason},
        )
    if "check constraint" in lowered:
        return CrmError(
            "VALIDATION_ERROR",
            "Значение не проходит ограничение таблицы.",
            {"reason": reason},
        )
    return CrmError(
        "CONFLICT",
        "Запись конфликтует с данными в базе.",
        {"reason": reason},
    )


def _execute(
    connection: sqlite3.Connection,
    sql: str,
    params: tuple[Any, ...] = (),
) -> sqlite3.Cursor:
    try:
        return connection.execute(sql, params)
    except sqlite3.IntegrityError as error:
        raise _translate_integrity(error) from error
    except sqlite3.OperationalError as error:
        raise CrmError(
            "DATABASE_ERROR",
            "Не удалось выполнить запрос к базе.",
            {"reason": str(error)},
        ) from error


def _require_table(table: str) -> str:
    if table not in _KNOWN_TABLES:
        raise CrmError("DATABASE_ERROR", "Неизвестная таблица.", {"table": table})
    return table


def _require(
    connection: sqlite3.Connection,
    table: str,
    entity_id: int,
    message: str,
) -> None:
    safe_table = _require_table(table)
    row = _execute(
        connection,
        f"SELECT id FROM {safe_table} WHERE id = ?",
        (entity_id,),
    ).fetchone()
    if row is None:
        raise CrmError("NOT_FOUND", message, {"id": entity_id})


def _find_row(
    connection: sqlite3.Connection,
    table: str,
    entity_id: int | None,
    message: str,
) -> sqlite3.Row | None:
    if entity_id is None:
        return None
    safe_table = _require_table(table)
    row = _execute(
        connection,
        f"SELECT * FROM {safe_table} WHERE id = ?",
        (entity_id,),
    ).fetchone()
    if row is None:
        raise CrmError("NOT_FOUND", message, {"id": entity_id})
    return row


def _same_link(
    left: object,
    right: int,
    message: str,
    details: dict[str, Any],
) -> None:
    """Если у связанной записи уже есть id, он должен совпасть."""
    if left is None:
        return
    if int(left) != right:
        raise CrmError("VALIDATION_ERROR", message, details)


def _ensure_order_links(
    connection: sqlite3.Connection,
    client_id: int | None,
    deal_id: int | None,
) -> None:
    _find_row(connection, "clients", client_id, "Клиент не найден.")
    deal = _find_row(connection, "deals", deal_id, "Сделка не найдена.")
    if deal is None or client_id is None:
        return
    _same_link(
        deal["client_id"],
        client_id,
        "Клиент заказа не совпадает с клиентом сделки.",
        {"client_id": client_id, "deal_id": deal_id},
    )


def _ensure_task_links(
    connection: sqlite3.Connection,
    client_id: int | None,
    deal_id: int | None,
    order_id: int | None,
) -> None:
    _find_row(connection, "clients", client_id, "Клиент не найден.")
    deal = _find_row(connection, "deals", deal_id, "Сделка не найдена.")
    order = _find_row(connection, "orders", order_id, "Заказ не найден.")
    if deal is not None and client_id is not None:
        _same_link(
            deal["client_id"],
            client_id,
            "Клиент задачи не совпадает с клиентом сделки.",
            {"client_id": client_id, "deal_id": deal_id},
        )
    if order is None:
        return
    if client_id is not None:
        _same_link(
            order["client_id"],
            client_id,
            "Клиент задачи не совпадает с клиентом заказа.",
            {"client_id": client_id, "order_id": order_id},
        )
    if deal_id is not None:
        _same_link(
            order["deal_id"],
            deal_id,
            "Сделка задачи не совпадает со сделкой заказа.",
            {"deal_id": deal_id, "order_id": order_id},
        )


def _insert_id(
    connection: sqlite3.Connection,
    sql: str,
    params: tuple[Any, ...],
) -> int:
    cursor = _execute(connection, sql, params)
    new_id = cursor.lastrowid
    if new_id is None:
        raise CrmError("DATABASE_ERROR", "База не вернула идентификатор записи.", None)
    return int(new_id)


def _load_one(
    connection: sqlite3.Connection,
    sql: str,
    entity_id: int,
    message: str,
    factory: Callable[[Mapping[str, Any]], T],
) -> T:
    row = _execute(connection, sql, (entity_id,)).fetchone()
    if row is None:
        raise CrmError("NOT_FOUND", message, {"id": entity_id})
    return factory(row)


def _resolve_db_path(db_path: str | Path) -> Path:
    raw = str(db_path).strip()
    if raw in {"", ":memory:"}:
        raise CrmError(
            "VALIDATION_ERROR",
            "Укажите файл базы SQLite. Общая память не подходит: "
            "каждое обращение открывает своё соединение.",
            {"db_path": raw},
        )
    path = Path(raw)
    if not path.is_absolute():
        path = Path(__file__).resolve().parent / path
    return path


class CrmDatabase:
    """CRUD мини-CRM в одном файле SQLite.

    При создании проверяет файл и создаёт недостающие таблицы.
    Пример:

        db = CrmDatabase("data/crm.sqlite")
        client = db.create_client(ClientInput(full_name="Иванова Анна"))
    """

    def __init__(self, db_path: str | Path) -> None:
        self.db_path = _resolve_db_path(db_path)
        if self.db_path.exists() and self.db_path.is_dir():
            raise CrmError(
                "VALIDATION_ERROR",
                "Путь базы указывает на папку, а не на файл.",
                {"db_path": str(self.db_path)},
            )
        try:
            self.db_path.parent.mkdir(parents=True, exist_ok=True)
        except OSError as error:
            raise CrmError(
                "DATABASE_ERROR",
                "Не удалось создать папку для файла базы.",
                {"path": str(self.db_path.parent), "reason": str(error)},
            ) from error
        self._create_schema()

    @contextmanager
    def _connection(self) -> Iterator[sqlite3.Connection]:
        """Короткое соединение: чужие ключи включены, ilike зарегистрирован."""
        try:
            connection = sqlite3.connect(
                self.db_path,
                timeout=_SQLITE_TIMEOUT_SECONDS,
            )
        except sqlite3.Error as error:
            raise CrmError(
                "DATABASE_ERROR",
                "Не удалось открыть файл базы.",
                {"reason": str(error)},
            ) from error
        try:
            connection.row_factory = sqlite3.Row
            connection.execute("PRAGMA foreign_keys = ON")
            connection.execute("PRAGMA journal_mode = WAL")
            flag = connection.execute("PRAGMA foreign_keys").fetchone()
            if flag is None or int(flag[0]) != 1:
                raise CrmError("DATABASE_ERROR", "SQLite не включил внешние ключи.", None)
            connection.create_function("ilike", 2, _sqlite_ilike, deterministic=True)
            yield connection
            connection.commit()
        except Exception:
            connection.rollback()
            raise
        finally:
            connection.close()

    def _create_schema(self) -> None:
        """Создать каждую таблицу и её индексы, если их ещё нет."""
        try:
            with self._connection() as connection:
                for table in TABLES:
                    connection.execute(table.create_sql())
                    for index in table.indexes:
                        connection.execute(index.create_sql(table.name))
        except sqlite3.DatabaseError as error:
            raise CrmError(
                "DATABASE_ERROR",
                "Не удалось подготовить таблицы CRM.",
                {"reason": str(error)},
            ) from error
        _log_event("схема проверена", {"tables": [table.name for table in TABLES]})

    def _page(
        self,
        *,
        select_sql: str,
        count_sql: str,
        params: tuple[Any, ...],
        order_by: str,
        limit: int,
        offset: int,
        factory: Callable[[Mapping[str, Any]], T],
    ) -> Page[T]:
        # order_by приходит только из литералов этого файла и ещё проверяется.
        if _ORDER_BY_SQL.fullmatch(order_by) is None:
            raise CrmError(
                "DATABASE_ERROR",
                "Некорректная сортировка запроса.",
                {"order_by": order_by},
            )
        with self._connection() as connection:
            total_row = _execute(connection, count_sql, params).fetchone()
            total = int(total_row[0]) if total_row is not None else 0
            fetched = _execute(
                connection,
                f"{select_sql} ORDER BY {order_by} LIMIT ? OFFSET ?",
                (*params, limit, offset),
            ).fetchall()
            # Снимок до закрытия соединения: Row после close читать нельзя.
            rows = [dict(zip(row.keys(), tuple(row))) for row in fetched]
        items = tuple(factory(row) for row in rows)
        return Page(items=items, total=total, limit=limit, offset=offset)

    def create_client(self, payload: ClientInput) -> Client:
        """Создать клиента."""
        fields = _normalize_client(payload)
        now = _utc_now()
        with self._connection() as connection:
            new_id = _insert_id(
                connection,
                """
                INSERT INTO clients (
                    full_name, company, email, phone, status, source, notes,
                    created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    fields.full_name,
                    fields.company,
                    fields.email,
                    fields.phone,
                    fields.status,
                    fields.source,
                    fields.notes,
                    now,
                    now,
                ),
            )
            created = _load_one(
                connection,
                "SELECT * FROM clients WHERE id = ?",
                new_id,
                "Клиент не найден.",
                Client.from_mapping,
            )
        _log_event("клиент создан", {"id": created.id})
        return created

    def get_client(self, client_id: int) -> Client:
        """Вернуть клиента по идентификатору."""
        client_id = _positive_id(client_id, "клиента")
        with self._connection() as connection:
            return _load_one(
                connection,
                "SELECT * FROM clients WHERE id = ?",
                client_id,
                "Клиент не найден.",
                Client.from_mapping,
            )

    def list_clients(
        self,
        *,
        status: str | None = None,
        limit: int = DEFAULT_PAGE_LIMIT,
        offset: int = 0,
    ) -> Page[Client]:
        """Список клиентов. status ограничивает выборку кодом статуса."""
        limit, offset = _page_bounds(limit, offset)
        clauses: list[str] = []
        params: list[Any] = []
        if status is not None:
            clauses.append("clients.status = ?")
            params.append(_choice(status, CLIENT_STATUSES, "статус клиента"))
        where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
        return self._page(
            select_sql=f"SELECT clients.* FROM clients {where}",
            count_sql=f"SELECT COUNT(*) FROM clients {where}",
            params=tuple(params),
            order_by="clients.updated_at DESC, clients.id DESC",
            limit=limit,
            offset=offset,
            factory=Client.from_mapping,
        )

    def search_clients(
        self,
        query: str,
        *,
        limit: int = DEFAULT_PAGE_LIMIT,
        offset: int = 0,
    ) -> Page[Client]:
        """Найти клиентов по имени, компании, почте, телефону, источнику и заметкам."""
        limit, offset = _page_bounds(limit, offset)
        pattern = _contains_pattern(
            _text(query, "поисковый запрос", QUERY_MAX_LENGTH, required=True)
        )
        where = f"WHERE {_ilike_clause(_CLIENT_SEARCH)}"
        return self._page(
            select_sql=f"SELECT clients.* FROM clients {where}",
            count_sql=f"SELECT COUNT(*) FROM clients {where}",
            params=(pattern,) * len(_CLIENT_SEARCH),
            order_by="clients.updated_at DESC, clients.id DESC",
            limit=limit,
            offset=offset,
            factory=Client.from_mapping,
        )

    def update_client(self, client_id: int, payload: ClientInput) -> Client:
        """Заменить карточку клиента."""
        client_id = _positive_id(client_id, "клиента")
        fields = _normalize_client(payload)
        with self._connection() as connection:
            _require(connection, "clients", client_id, "Клиент не найден.")
            _execute(
                connection,
                """
                UPDATE clients
                SET full_name = ?, company = ?, email = ?, phone = ?, status = ?,
                    source = ?, notes = ?, updated_at = ?
                WHERE id = ?
                """,
                (
                    fields.full_name,
                    fields.company,
                    fields.email,
                    fields.phone,
                    fields.status,
                    fields.source,
                    fields.notes,
                    _utc_now(),
                    client_id,
                ),
            )
            updated = _load_one(
                connection,
                "SELECT * FROM clients WHERE id = ?",
                client_id,
                "Клиент не найден.",
                Client.from_mapping,
            )
        _log_event("клиент изменён", {"id": client_id})
        return updated

    def set_client_status(self, client_id: int, status: str) -> Client:
        """Сменить статус. Архив — это status=archived, карточка остаётся."""
        client_id = _positive_id(client_id, "клиента")
        normalized = _choice(status, CLIENT_STATUSES, "статус клиента")
        with self._connection() as connection:
            _require(connection, "clients", client_id, "Клиент не найден.")
            _execute(
                connection,
                "UPDATE clients SET status = ?, updated_at = ? WHERE id = ?",
                (normalized, _utc_now(), client_id),
            )
            updated = _load_one(
                connection,
                "SELECT * FROM clients WHERE id = ?",
                client_id,
                "Клиент не найден.",
                Client.from_mapping,
            )
        _log_event("статус клиента изменён", {"id": client_id, "status": normalized})
        return updated

    def delete_client(self, client_id: int) -> int:
        """Удалить клиента. Сделки, заказы и задачи остаются, связь обнуляется."""
        client_id = _positive_id(client_id, "клиента")
        with self._connection() as connection:
            _require(connection, "clients", client_id, "Клиент не найден.")
            _execute(connection, "DELETE FROM clients WHERE id = ?", (client_id,))
        _log_event("клиент удалён", {"id": client_id})
        return client_id

    def create_deal(self, payload: DealInput) -> Deal:
        """Создать сделку. Клиента можно не указывать."""
        fields = _normalize_deal(payload)
        now = _utc_now()
        with self._connection() as connection:
            _find_row(connection, "clients", fields.client_id, "Клиент не найден.")
            new_id = _insert_id(
                connection,
                """
                INSERT INTO deals (
                    title, client_id, status, amount_minor, currency,
                    expected_close_on, notes, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    fields.title,
                    fields.client_id,
                    fields.status,
                    fields.amount_minor,
                    fields.currency,
                    fields.expected_close_on,
                    fields.notes,
                    now,
                    now,
                ),
            )
            created = _load_one(
                connection,
                f"{_DEAL_SELECT} WHERE deals.id = ?",
                new_id,
                "Сделка не найдена.",
                Deal.from_mapping,
            )
        _log_event("сделка создана", {"id": created.id})
        return created

    def get_deal(self, deal_id: int) -> Deal:
        """Вернуть сделку по идентификатору."""
        deal_id = _positive_id(deal_id, "сделки")
        with self._connection() as connection:
            return _load_one(
                connection,
                f"{_DEAL_SELECT} WHERE deals.id = ?",
                deal_id,
                "Сделка не найдена.",
                Deal.from_mapping,
            )

    def list_deals(
        self,
        *,
        status: str | None = None,
        client_id: int | None = None,
        limit: int = DEFAULT_PAGE_LIMIT,
        offset: int = 0,
    ) -> Page[Deal]:
        """Список сделок с фильтром по статусу и клиенту."""
        limit, offset = _page_bounds(limit, offset)
        clauses: list[str] = []
        params: list[Any] = []
        if status is not None:
            clauses.append("deals.status = ?")
            params.append(_choice(status, DEAL_STATUSES, "статус сделки"))
        if client_id is not None:
            clauses.append("deals.client_id = ?")
            params.append(_positive_id(client_id, "клиента"))
        where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
        return self._page(
            select_sql=f"{_DEAL_SELECT} {where}",
            count_sql=f"SELECT COUNT(*) {_DEAL_FROM} {where}",
            params=tuple(params),
            order_by="deals.updated_at DESC, deals.id DESC",
            limit=limit,
            offset=offset,
            factory=Deal.from_mapping,
        )

    def search_deals(
        self,
        query: str,
        *,
        limit: int = DEFAULT_PAGE_LIMIT,
        offset: int = 0,
    ) -> Page[Deal]:
        """Найти сделки по названию, заметкам, статусу и карточке клиента."""
        limit, offset = _page_bounds(limit, offset)
        pattern = _contains_pattern(
            _text(query, "поисковый запрос", QUERY_MAX_LENGTH, required=True)
        )
        where = f"WHERE {_ilike_clause(_DEAL_SEARCH)}"
        return self._page(
            select_sql=f"{_DEAL_SELECT} {where}",
            count_sql=f"SELECT COUNT(*) {_DEAL_FROM} {where}",
            params=(pattern,) * len(_DEAL_SEARCH),
            order_by="deals.updated_at DESC, deals.id DESC",
            limit=limit,
            offset=offset,
            factory=Deal.from_mapping,
        )

    def update_deal(self, deal_id: int, payload: DealInput) -> Deal:
        """Заменить карточку сделки, в том числе прикреплённого клиента."""
        deal_id = _positive_id(deal_id, "сделки")
        fields = _normalize_deal(payload)
        with self._connection() as connection:
            _require(connection, "deals", deal_id, "Сделка не найдена.")
            _find_row(connection, "clients", fields.client_id, "Клиент не найден.")
            _execute(
                connection,
                """
                UPDATE deals
                SET title = ?, client_id = ?, status = ?, amount_minor = ?,
                    currency = ?, expected_close_on = ?, notes = ?, updated_at = ?
                WHERE id = ?
                """,
                (
                    fields.title,
                    fields.client_id,
                    fields.status,
                    fields.amount_minor,
                    fields.currency,
                    fields.expected_close_on,
                    fields.notes,
                    _utc_now(),
                    deal_id,
                ),
            )
            updated = _load_one(
                connection,
                f"{_DEAL_SELECT} WHERE deals.id = ?",
                deal_id,
                "Сделка не найдена.",
                Deal.from_mapping,
            )
        _log_event("сделка изменена", {"id": deal_id})
        return updated

    def set_deal_status(self, deal_id: int, status: str) -> Deal:
        """Сменить статус сделки."""
        deal_id = _positive_id(deal_id, "сделки")
        normalized = _choice(status, DEAL_STATUSES, "статус сделки")
        with self._connection() as connection:
            _require(connection, "deals", deal_id, "Сделка не найдена.")
            _execute(
                connection,
                "UPDATE deals SET status = ?, updated_at = ? WHERE id = ?",
                (normalized, _utc_now(), deal_id),
            )
            updated = _load_one(
                connection,
                f"{_DEAL_SELECT} WHERE deals.id = ?",
                deal_id,
                "Сделка не найдена.",
                Deal.from_mapping,
            )
        _log_event("статус сделки изменён", {"id": deal_id, "status": normalized})
        return updated

    def set_deal_client(self, deal_id: int, client_id: int | None) -> Deal:
        """Прикрепить клиента к сделке или открепить, если client_id пустой."""
        deal_id = _positive_id(deal_id, "сделки")
        client_id = _optional_id(client_id, "клиента")
        with self._connection() as connection:
            _require(connection, "deals", deal_id, "Сделка не найдена.")
            _find_row(connection, "clients", client_id, "Клиент не найден.")
            _execute(
                connection,
                "UPDATE deals SET client_id = ?, updated_at = ? WHERE id = ?",
                (client_id, _utc_now(), deal_id),
            )
            updated = _load_one(
                connection,
                f"{_DEAL_SELECT} WHERE deals.id = ?",
                deal_id,
                "Сделка не найдена.",
                Deal.from_mapping,
            )
        _log_event("клиент сделки изменён", {"id": deal_id, "client_id": client_id})
        return updated

    def delete_deal(self, deal_id: int) -> int:
        """Удалить сделку. У заказов и задач ссылка на неё обнуляется."""
        deal_id = _positive_id(deal_id, "сделки")
        with self._connection() as connection:
            _require(connection, "deals", deal_id, "Сделка не найдена.")
            _execute(connection, "DELETE FROM deals WHERE id = ?", (deal_id,))
        _log_event("сделка удалена", {"id": deal_id})
        return deal_id

    def create_order(self, payload: OrderInput) -> Order:
        """Создать заказ."""
        fields = _normalize_order(payload)
        now = _utc_now()
        with self._connection() as connection:
            _ensure_order_links(connection, fields.client_id, fields.deal_id)
            new_id = _insert_id(
                connection,
                """
                INSERT INTO orders (
                    title, client_id, deal_id, status, amount_minor, currency,
                    due_on, notes, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    fields.title,
                    fields.client_id,
                    fields.deal_id,
                    fields.status,
                    fields.amount_minor,
                    fields.currency,
                    fields.due_on,
                    fields.notes,
                    now,
                    now,
                ),
            )
            created = _load_one(
                connection,
                f"{_ORDER_SELECT} WHERE orders.id = ?",
                new_id,
                "Заказ не найден.",
                Order.from_mapping,
            )
        _log_event("заказ создан", {"id": created.id})
        return created

    def get_order(self, order_id: int) -> Order:
        """Вернуть заказ по идентификатору."""
        order_id = _positive_id(order_id, "заказа")
        with self._connection() as connection:
            return _load_one(
                connection,
                f"{_ORDER_SELECT} WHERE orders.id = ?",
                order_id,
                "Заказ не найден.",
                Order.from_mapping,
            )

    def list_orders(
        self,
        *,
        status: str | None = None,
        client_id: int | None = None,
        deal_id: int | None = None,
        limit: int = DEFAULT_PAGE_LIMIT,
        offset: int = 0,
    ) -> Page[Order]:
        """Список заказов с фильтром по статусу, клиенту и сделке."""
        limit, offset = _page_bounds(limit, offset)
        clauses: list[str] = []
        params: list[Any] = []
        if status is not None:
            clauses.append("orders.status = ?")
            params.append(_choice(status, ORDER_STATUSES, "статус заказа"))
        if client_id is not None:
            clauses.append("orders.client_id = ?")
            params.append(_positive_id(client_id, "клиента"))
        if deal_id is not None:
            clauses.append("orders.deal_id = ?")
            params.append(_positive_id(deal_id, "сделки"))
        where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
        return self._page(
            select_sql=f"{_ORDER_SELECT} {where}",
            count_sql=f"SELECT COUNT(*) {_ORDER_FROM} {where}",
            params=tuple(params),
            order_by="orders.updated_at DESC, orders.id DESC",
            limit=limit,
            offset=offset,
            factory=Order.from_mapping,
        )

    def search_orders(
        self,
        query: str,
        *,
        limit: int = DEFAULT_PAGE_LIMIT,
        offset: int = 0,
    ) -> Page[Order]:
        """Найти заказы по названию, заметкам, статусу, клиенту и сделке."""
        limit, offset = _page_bounds(limit, offset)
        pattern = _contains_pattern(
            _text(query, "поисковый запрос", QUERY_MAX_LENGTH, required=True)
        )
        where = f"WHERE {_ilike_clause(_ORDER_SEARCH)}"
        return self._page(
            select_sql=f"{_ORDER_SELECT} {where}",
            count_sql=f"SELECT COUNT(*) {_ORDER_FROM} {where}",
            params=(pattern,) * len(_ORDER_SEARCH),
            order_by="orders.updated_at DESC, orders.id DESC",
            limit=limit,
            offset=offset,
            factory=Order.from_mapping,
        )

    def update_order(self, order_id: int, payload: OrderInput) -> Order:
        """Заменить карточку заказа."""
        order_id = _positive_id(order_id, "заказа")
        fields = _normalize_order(payload)
        with self._connection() as connection:
            _require(connection, "orders", order_id, "Заказ не найден.")
            _ensure_order_links(connection, fields.client_id, fields.deal_id)
            _execute(
                connection,
                """
                UPDATE orders
                SET title = ?, client_id = ?, deal_id = ?, status = ?,
                    amount_minor = ?, currency = ?, due_on = ?, notes = ?,
                    updated_at = ?
                WHERE id = ?
                """,
                (
                    fields.title,
                    fields.client_id,
                    fields.deal_id,
                    fields.status,
                    fields.amount_minor,
                    fields.currency,
                    fields.due_on,
                    fields.notes,
                    _utc_now(),
                    order_id,
                ),
            )
            updated = _load_one(
                connection,
                f"{_ORDER_SELECT} WHERE orders.id = ?",
                order_id,
                "Заказ не найден.",
                Order.from_mapping,
            )
        _log_event("заказ изменён", {"id": order_id})
        return updated

    def set_order_status(self, order_id: int, status: str) -> Order:
        """Сменить статус заказа."""
        order_id = _positive_id(order_id, "заказа")
        normalized = _choice(status, ORDER_STATUSES, "статус заказа")
        with self._connection() as connection:
            _require(connection, "orders", order_id, "Заказ не найден.")
            _execute(
                connection,
                "UPDATE orders SET status = ?, updated_at = ? WHERE id = ?",
                (normalized, _utc_now(), order_id),
            )
            updated = _load_one(
                connection,
                f"{_ORDER_SELECT} WHERE orders.id = ?",
                order_id,
                "Заказ не найден.",
                Order.from_mapping,
            )
        _log_event("статус заказа изменён", {"id": order_id, "status": normalized})
        return updated

    def delete_order(self, order_id: int) -> int:
        """Удалить заказ. У задач ссылка на него обнуляется."""
        order_id = _positive_id(order_id, "заказа")
        with self._connection() as connection:
            _require(connection, "orders", order_id, "Заказ не найден.")
            _execute(connection, "DELETE FROM orders WHERE id = ?", (order_id,))
        _log_event("заказ удалён", {"id": order_id})
        return order_id

    def create_task(self, payload: TaskInput) -> Task:
        """Создать задачу или напоминание."""
        fields = _normalize_task(payload)
        now = _utc_now()
        completed_at = _completed_at(None, fields.is_done, now)
        with self._connection() as connection:
            _ensure_task_links(
                connection,
                fields.client_id,
                fields.deal_id,
                fields.order_id,
            )
            new_id = _insert_id(
                connection,
                """
                INSERT INTO tasks (
                    title, details, kind, due_at, is_done, client_id, deal_id,
                    order_id, created_at, updated_at, completed_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    fields.title,
                    fields.details,
                    fields.kind,
                    fields.due_at,
                    int(fields.is_done),
                    fields.client_id,
                    fields.deal_id,
                    fields.order_id,
                    now,
                    now,
                    completed_at,
                ),
            )
            created = _load_one(
                connection,
                f"{_TASK_SELECT} WHERE tasks.id = ?",
                new_id,
                "Задача не найдена.",
                Task.from_mapping,
            )
        _log_event("задача создана", {"id": created.id, "kind": created.kind})
        return created

    def get_task(self, task_id: int) -> Task:
        """Вернуть задачу по идентификатору."""
        task_id = _positive_id(task_id, "задачи")
        with self._connection() as connection:
            return _load_one(
                connection,
                f"{_TASK_SELECT} WHERE tasks.id = ?",
                task_id,
                "Задача не найдена.",
                Task.from_mapping,
            )

    def list_tasks(
        self,
        *,
        client_id: int | None = None,
        deal_id: int | None = None,
        order_id: int | None = None,
        is_done: bool | None = None,
        kind: str | None = None,
        due_before: str | None = None,
        limit: int = DEFAULT_PAGE_LIMIT,
        offset: int = 0,
    ) -> Page[Task]:
        """Список задач. due_before оставляет напоминания со сроком не позже даты."""
        limit, offset = _page_bounds(limit, offset)
        clauses: list[str] = []
        params: list[Any] = []
        if client_id is not None:
            clauses.append("tasks.client_id = ?")
            params.append(_positive_id(client_id, "клиента"))
        if deal_id is not None:
            clauses.append("tasks.deal_id = ?")
            params.append(_positive_id(deal_id, "сделки"))
        if order_id is not None:
            clauses.append("tasks.order_id = ?")
            params.append(_positive_id(order_id, "заказа"))
        if is_done is not None:
            clauses.append("tasks.is_done = ?")
            params.append(int(_flag_value(is_done, "выполнена")))
        if kind is not None:
            clauses.append("tasks.kind = ?")
            params.append(_choice(kind, TASK_KINDS, "вид задачи"))
        if due_before is not None:
            boundary = _iso_datetime(due_before, "срок до")
            if boundary == "":
                raise CrmError(
                    "VALIDATION_ERROR",
                    "Укажите срок в поле «срок до».",
                    {"field": "due_before"},
                )
            clauses.append("tasks.due_at != '' AND tasks.due_at <= ?")
            params.append(boundary)
        where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
        return self._page(
            select_sql=f"{_TASK_SELECT} {where}",
            count_sql=f"SELECT COUNT(*) {_TASK_FROM} {where}",
            params=tuple(params),
            order_by="tasks.is_done ASC, tasks.due_at ASC, tasks.id DESC",
            limit=limit,
            offset=offset,
            factory=Task.from_mapping,
        )

    def update_task(self, task_id: int, payload: TaskInput) -> Task:
        """Заменить задачу, её связи и отметку выполнения."""
        task_id = _positive_id(task_id, "задачи")
        fields = _normalize_task(payload)
        with self._connection() as connection:
            existing = _load_one(
                connection,
                f"{_TASK_SELECT} WHERE tasks.id = ?",
                task_id,
                "Задача не найдена.",
                Task.from_mapping,
            )
            _ensure_task_links(
                connection,
                fields.client_id,
                fields.deal_id,
                fields.order_id,
            )
            now = _utc_now()
            _execute(
                connection,
                """
                UPDATE tasks
                SET title = ?, details = ?, kind = ?, due_at = ?, is_done = ?,
                    client_id = ?, deal_id = ?, order_id = ?, updated_at = ?,
                    completed_at = ?
                WHERE id = ?
                """,
                (
                    fields.title,
                    fields.details,
                    fields.kind,
                    fields.due_at,
                    int(fields.is_done),
                    fields.client_id,
                    fields.deal_id,
                    fields.order_id,
                    now,
                    _completed_at(existing, fields.is_done, now),
                    task_id,
                ),
            )
            updated = _load_one(
                connection,
                f"{_TASK_SELECT} WHERE tasks.id = ?",
                task_id,
                "Задача не найдена.",
                Task.from_mapping,
            )
        _log_event("задача изменена", {"id": task_id})
        return updated

    def set_task_done(self, task_id: int, is_done: bool) -> Task:
        """Отметить задачу выполненной или снова открыть её."""
        task_id = _positive_id(task_id, "задачи")
        done = _flag_value(is_done, "выполнена")
        with self._connection() as connection:
            existing = _load_one(
                connection,
                f"{_TASK_SELECT} WHERE tasks.id = ?",
                task_id,
                "Задача не найдена.",
                Task.from_mapping,
            )
            now = _utc_now()
            _execute(
                connection,
                """
                UPDATE tasks
                SET is_done = ?, completed_at = ?, updated_at = ?
                WHERE id = ?
                """,
                (int(done), _completed_at(existing, done, now), now, task_id),
            )
            updated = _load_one(
                connection,
                f"{_TASK_SELECT} WHERE tasks.id = ?",
                task_id,
                "Задача не найдена.",
                Task.from_mapping,
            )
        _log_event("отметка задачи изменена", {"id": task_id, "is_done": done})
        return updated

    def delete_task(self, task_id: int) -> int:
        """Удалить задачу."""
        task_id = _positive_id(task_id, "задачи")
        with self._connection() as connection:
            _require(connection, "tasks", task_id, "Задача не найдена.")
            _execute(connection, "DELETE FROM tasks WHERE id = ?", (task_id,))
        _log_event("задача удалена", {"id": task_id})
        return task_id

    def summary(self) -> dict[str, Any]:
        """Короткие счётчики для главного экрана."""
        with self._connection() as connection:
            return {
                "clients": _count_by(connection, "clients", "status"),
                "deals": _count_by(connection, "deals", "status"),
                "orders": _count_by(connection, "orders", "status"),
                "tasks_open": _scalar(
                    connection,
                    "SELECT COUNT(*) FROM tasks WHERE is_done = 0",
                ),
                "reminders_open": _scalar(
                    connection,
                    "SELECT COUNT(*) FROM tasks WHERE is_done = 0 AND kind = ?",
                    (TaskKind.REMINDER.value,),
                ),
            }


def _scalar(
    connection: sqlite3.Connection,
    sql: str,
    params: tuple[Any, ...] = (),
) -> int:
    row = _execute(connection, sql, params).fetchone()
    if row is None:
        return 0
    return int(row[0])


def _count_by(
    connection: sqlite3.Connection,
    table: str,
    column: str,
) -> dict[str, int]:
    safe_table = _require_table(table)
    if column not in {"status", "kind"}:
        raise CrmError("DATABASE_ERROR", "Неизвестная колонка сводки.", {"column": column})
    rows = _execute(
        connection,
        f"SELECT {column} AS key, COUNT(*) AS total FROM {safe_table} GROUP BY {column}",
    ).fetchall()
    return {str(row["key"]): int(row["total"]) for row in rows}
