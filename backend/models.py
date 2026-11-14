"""Модели мини-CRM и схема SQLite.

Деньги лежат в INTEGER (мелкие единицы, для рубля это копейки).
Даты и время — TEXT в ISO-8601. Логический флаг — INTEGER 0 или 1.
Тип REAL сознательно не используем: дробные деньги в SQLite плавают.
"""

from __future__ import annotations

import re
from collections.abc import Mapping
from dataclasses import dataclass
from enum import Enum
from typing import Any, Generic, Protocol, TypeVar


# Типы хранения SQLite. Класс базы собирает из них CREATE TABLE.
SQLITE_INTEGER = "INTEGER"
SQLITE_TEXT = "TEXT"
SQLITE_REAL = "REAL"

NAME_MAX_LENGTH = 200
COMPANY_MAX_LENGTH = 200
EMAIL_MAX_LENGTH = 200
PHONE_MAX_LENGTH = 40
SOURCE_MAX_LENGTH = 120
TITLE_MAX_LENGTH = 200
NOTES_MAX_LENGTH = 4000
DETAILS_MAX_LENGTH = 4000
QUERY_MAX_LENGTH = 200

DEFAULT_PAGE_LIMIT = 50
MAX_PAGE_LIMIT = 100
DEFAULT_PAGE_OFFSET = 0

# 100 мелких единиц = 1.00. Потолок — защита от случайного мусора, не бизнес-лимит.
MINOR_SCALE = 100
MAX_AMOUNT_MINOR = 10**14
DEFAULT_CURRENCY = "RUB"

_IDENTIFIER = re.compile(r"^[a-z_][a-z0-9_]*$")
_STATUS_CODE = re.compile(r"^[a-z_]{1,32}$")


class ClientStatus(str, Enum):
    """Код статуса клиента. В таблице clients колонка status имеет тип TEXT."""

    LEAD = "lead"
    ACTIVE = "active"
    INACTIVE = "inactive"
    ARCHIVED = "archived"


class DealStatus(str, Enum):
    """Воронка сделки. Код хранится в deals.status (TEXT)."""

    NEW = "new"
    QUALIFICATION = "qualification"
    PROPOSAL = "proposal"
    NEGOTIATION = "negotiation"
    WON = "won"
    LOST = "lost"
    ARCHIVED = "archived"


class OrderStatus(str, Enum):
    """Статус заказа. Код хранится в orders.status (TEXT)."""

    DRAFT = "draft"
    CONFIRMED = "confirmed"
    IN_PROGRESS = "in_progress"
    FULFILLED = "fulfilled"
    CANCELLED = "cancelled"


class TaskKind(str, Enum):
    """Задача или напоминание. Код хранится в tasks.kind (TEXT)."""

    TASK = "task"
    REMINDER = "reminder"


CLIENT_STATUSES: tuple[str, ...] = tuple(item.value for item in ClientStatus)
DEAL_STATUSES: tuple[str, ...] = tuple(item.value for item in DealStatus)
ORDER_STATUSES: tuple[str, ...] = tuple(item.value for item in OrderStatus)
TASK_KINDS: tuple[str, ...] = tuple(item.value for item in TaskKind)

CLIENT_STATUS_LABELS: dict[str, str] = {
    ClientStatus.LEAD.value: "Потенциальный",
    ClientStatus.ACTIVE.value: "Действующий",
    ClientStatus.INACTIVE.value: "Неактивный",
    ClientStatus.ARCHIVED.value: "Архив",
}
DEAL_STATUS_LABELS: dict[str, str] = {
    DealStatus.NEW.value: "Новая",
    DealStatus.QUALIFICATION.value: "Уточнение",
    DealStatus.PROPOSAL.value: "Предложение",
    DealStatus.NEGOTIATION.value: "Согласование",
    DealStatus.WON.value: "Выиграна",
    DealStatus.LOST.value: "Проиграна",
    DealStatus.ARCHIVED.value: "Архив",
}
ORDER_STATUS_LABELS: dict[str, str] = {
    OrderStatus.DRAFT.value: "Черновик",
    OrderStatus.CONFIRMED.value: "Подтверждён",
    OrderStatus.IN_PROGRESS.value: "В работе",
    OrderStatus.FULFILLED.value: "Выполнен",
    OrderStatus.CANCELLED.value: "Отменён",
}
TASK_KIND_LABELS: dict[str, str] = {
    TaskKind.TASK.value: "Задача",
    TaskKind.REMINDER.value: "Напоминание",
}


def _assert_labels() -> None:
    """Уронить импорт, если у кода статуса нет русской подписи."""
    pairs = (
        (CLIENT_STATUSES, CLIENT_STATUS_LABELS),
        (DEAL_STATUSES, DEAL_STATUS_LABELS),
        (ORDER_STATUSES, ORDER_STATUS_LABELS),
        (TASK_KINDS, TASK_KIND_LABELS),
    )
    for codes, labels in pairs:
        if set(codes) != set(labels):
            raise RuntimeError("Подписи справочника не совпадают с кодами.")


_assert_labels()


def _assert_identifier(value: str) -> str:
    """Пропустить только простое имя таблицы или колонки."""
    if _IDENTIFIER.fullmatch(value) is None:
        raise ValueError(f"Недопустимое имя SQL: {value}")
    return value


def _in_check(column: str, values: tuple[str, ...]) -> str:
    """CHECK (column IN (...)) из наших кодов, не из ввода пользователя."""
    _assert_identifier(column)
    quoted: list[str] = []
    for value in values:
        if _STATUS_CODE.fullmatch(value) is None:
            raise ValueError(f"Недопустимый код справочника: {value}")
        quoted.append(f"'{value}'")
    return f"CHECK ({column} IN ({', '.join(quoted)}))"


def _status_tail(default: str, column: str, values: tuple[str, ...]) -> str:
    """NOT NULL, значение по умолчанию и проверка кода."""
    if _STATUS_CODE.fullmatch(default) is None:
        raise ValueError("Недопустимое значение по умолчанию.")
    return f"NOT NULL DEFAULT '{default}' {_in_check(column, values)}"


@dataclass(frozen=True)
class SqlColumn:
    """Колонка: имя, тип SQLite и хвост ограничений."""

    name: str
    sqlite_type: str
    definition_tail: str = ""

    def __post_init__(self) -> None:
        _assert_identifier(self.name)
        if self.sqlite_type not in {SQLITE_INTEGER, SQLITE_TEXT, SQLITE_REAL}:
            raise ValueError(f"Неизвестный тип SQLite: {self.sqlite_type}")

    def ddl(self) -> str:
        """Фрагмент определения колонки."""
        if self.definition_tail:
            return f"{self.name} {self.sqlite_type} {self.definition_tail}"
        return f"{self.name} {self.sqlite_type}"


@dataclass(frozen=True)
class SqlIndex:
    """Индекс по колонкам одной таблицы."""

    name: str
    columns: tuple[str, ...]

    def __post_init__(self) -> None:
        _assert_identifier(self.name)
        if not self.columns:
            raise ValueError(f"У индекса {self.name} нет колонок.")
        for column in self.columns:
            _assert_identifier(column)

    def create_sql(self, table: str) -> str:
        """CREATE INDEX IF NOT EXISTS."""
        _assert_identifier(table)
        columns = ", ".join(self.columns)
        return f"CREATE INDEX IF NOT EXISTS {self.name} ON {table} ({columns})"


@dataclass(frozen=True)
class SqlTable:
    """Таблица, которую класс базы создаёт, если её ещё нет."""

    name: str
    columns: tuple[SqlColumn, ...]
    indexes: tuple[SqlIndex, ...] = ()

    def __post_init__(self) -> None:
        _assert_identifier(self.name)
        if not self.columns:
            raise ValueError(f"У таблицы {self.name} нет колонок.")
        known = {column.name for column in self.columns}
        for index in self.indexes:
            for column in index.columns:
                if column not in known:
                    raise ValueError(
                        f"Индекс {index.name} ссылается на неизвестную колонку {column}."
                    )

    def create_sql(self) -> str:
        """CREATE TABLE IF NOT EXISTS из колонок и их типов SQLite."""
        body = ",\n    ".join(column.ddl() for column in self.columns)
        return f"CREATE TABLE IF NOT EXISTS {self.name} (\n    {body}\n)"


# Клиент может быть действующим, потенциальным или в архиве. Удаление отдельно.
CLIENTS_TABLE = SqlTable(
    name="clients",
    columns=(
        SqlColumn("id", SQLITE_INTEGER, "PRIMARY KEY"),
        SqlColumn("full_name", SQLITE_TEXT, "NOT NULL"),
        SqlColumn("company", SQLITE_TEXT, "NOT NULL DEFAULT ''"),
        SqlColumn("email", SQLITE_TEXT, "NOT NULL DEFAULT ''"),
        SqlColumn("phone", SQLITE_TEXT, "NOT NULL DEFAULT ''"),
        SqlColumn(
            "status",
            SQLITE_TEXT,
            _status_tail(ClientStatus.ACTIVE.value, "status", CLIENT_STATUSES),
        ),
        SqlColumn("source", SQLITE_TEXT, "NOT NULL DEFAULT ''"),
        SqlColumn("notes", SQLITE_TEXT, "NOT NULL DEFAULT ''"),
        SqlColumn("created_at", SQLITE_TEXT, "NOT NULL"),
        SqlColumn("updated_at", SQLITE_TEXT, "NOT NULL"),
    ),
    indexes=(
        SqlIndex("idx_clients_status", ("status",)),
        SqlIndex("idx_clients_updated_at", ("updated_at",)),
    ),
)

# Сделка может жить без клиента: client_id пустой, пока человека не прикрепят.
DEALS_TABLE = SqlTable(
    name="deals",
    columns=(
        SqlColumn("id", SQLITE_INTEGER, "PRIMARY KEY"),
        SqlColumn("title", SQLITE_TEXT, "NOT NULL"),
        SqlColumn(
            "client_id",
            SQLITE_INTEGER,
            "REFERENCES clients(id) ON DELETE SET NULL",
        ),
        SqlColumn(
            "status",
            SQLITE_TEXT,
            _status_tail(DealStatus.NEW.value, "status", DEAL_STATUSES),
        ),
        SqlColumn(
            "amount_minor",
            SQLITE_INTEGER,
            "NOT NULL DEFAULT 0 CHECK (amount_minor >= 0)",
        ),
        SqlColumn(
            "currency",
            SQLITE_TEXT,
            "NOT NULL DEFAULT 'RUB' CHECK (length(currency) = 3)",
        ),
        SqlColumn("expected_close_on", SQLITE_TEXT, "NOT NULL DEFAULT ''"),
        SqlColumn("notes", SQLITE_TEXT, "NOT NULL DEFAULT ''"),
        SqlColumn("created_at", SQLITE_TEXT, "NOT NULL"),
        SqlColumn("updated_at", SQLITE_TEXT, "NOT NULL"),
    ),
    indexes=(
        SqlIndex("idx_deals_client_id", ("client_id",)),
        SqlIndex("idx_deals_status", ("status",)),
        SqlIndex("idx_deals_updated_at", ("updated_at",)),
    ),
)

# Заказ — конкретная работа. Сделка и клиент необязательны и ставятся по отдельности.
ORDERS_TABLE = SqlTable(
    name="orders",
    columns=(
        SqlColumn("id", SQLITE_INTEGER, "PRIMARY KEY"),
        SqlColumn("title", SQLITE_TEXT, "NOT NULL"),
        SqlColumn(
            "client_id",
            SQLITE_INTEGER,
            "REFERENCES clients(id) ON DELETE SET NULL",
        ),
        SqlColumn(
            "deal_id",
            SQLITE_INTEGER,
            "REFERENCES deals(id) ON DELETE SET NULL",
        ),
        SqlColumn(
            "status",
            SQLITE_TEXT,
            _status_tail(OrderStatus.DRAFT.value, "status", ORDER_STATUSES),
        ),
        SqlColumn(
            "amount_minor",
            SQLITE_INTEGER,
            "NOT NULL DEFAULT 0 CHECK (amount_minor >= 0)",
        ),
        SqlColumn(
            "currency",
            SQLITE_TEXT,
            "NOT NULL DEFAULT 'RUB' CHECK (length(currency) = 3)",
        ),
        SqlColumn("due_on", SQLITE_TEXT, "NOT NULL DEFAULT ''"),
        SqlColumn("notes", SQLITE_TEXT, "NOT NULL DEFAULT ''"),
        SqlColumn("created_at", SQLITE_TEXT, "NOT NULL"),
        SqlColumn("updated_at", SQLITE_TEXT, "NOT NULL"),
    ),
    indexes=(
        SqlIndex("idx_orders_client_id", ("client_id",)),
        SqlIndex("idx_orders_deal_id", ("deal_id",)),
        SqlIndex("idx_orders_status", ("status",)),
        SqlIndex("idx_orders_updated_at", ("updated_at",)),
    ),
)

# Задача или напоминание. Связь с клиентом, сделкой и заказом необязательна.
TASKS_TABLE = SqlTable(
    name="tasks",
    columns=(
        SqlColumn("id", SQLITE_INTEGER, "PRIMARY KEY"),
        SqlColumn("title", SQLITE_TEXT, "NOT NULL"),
        SqlColumn("details", SQLITE_TEXT, "NOT NULL DEFAULT ''"),
        SqlColumn(
            "kind",
            SQLITE_TEXT,
            _status_tail(TaskKind.TASK.value, "kind", TASK_KINDS),
        ),
        SqlColumn("due_at", SQLITE_TEXT, "NOT NULL DEFAULT ''"),
        SqlColumn(
            "is_done",
            SQLITE_INTEGER,
            "NOT NULL DEFAULT 0 CHECK (is_done IN (0, 1))",
        ),
        SqlColumn(
            "client_id",
            SQLITE_INTEGER,
            "REFERENCES clients(id) ON DELETE SET NULL",
        ),
        SqlColumn(
            "deal_id",
            SQLITE_INTEGER,
            "REFERENCES deals(id) ON DELETE SET NULL",
        ),
        SqlColumn(
            "order_id",
            SQLITE_INTEGER,
            "REFERENCES orders(id) ON DELETE SET NULL",
        ),
        SqlColumn("created_at", SQLITE_TEXT, "NOT NULL"),
        SqlColumn("updated_at", SQLITE_TEXT, "NOT NULL"),
        SqlColumn("completed_at", SQLITE_TEXT, "NOT NULL DEFAULT ''"),
    ),
    indexes=(
        SqlIndex("idx_tasks_client_id", ("client_id",)),
        SqlIndex("idx_tasks_deal_id", ("deal_id",)),
        SqlIndex("idx_tasks_order_id", ("order_id",)),
        SqlIndex("idx_tasks_is_done", ("is_done",)),
        SqlIndex("idx_tasks_due_at", ("due_at",)),
    ),
)

TABLES: tuple[SqlTable, ...] = (
    CLIENTS_TABLE,
    DEALS_TABLE,
    ORDERS_TABLE,
    TASKS_TABLE,
)


@dataclass(frozen=True)
class ClientInput:
    """Поля клиента без идентификатора и служебных дат."""

    full_name: str
    company: str = ""
    email: str = ""
    phone: str = ""
    status: str = ClientStatus.ACTIVE.value
    source: str = ""
    notes: str = ""


@dataclass(frozen=True)
class DealInput:
    """Поля сделки. client_id пустой, пока клиента не прикрепили."""

    title: str
    client_id: int | None = None
    status: str = DealStatus.NEW.value
    amount_minor: int = 0
    currency: str = DEFAULT_CURRENCY
    expected_close_on: str = ""
    notes: str = ""


@dataclass(frozen=True)
class OrderInput:
    """Поля заказа. Клиент и сделка независимы и могут отсутствовать."""

    title: str
    client_id: int | None = None
    deal_id: int | None = None
    status: str = OrderStatus.DRAFT.value
    amount_minor: int = 0
    currency: str = DEFAULT_CURRENCY
    due_on: str = ""
    notes: str = ""


@dataclass(frozen=True)
class TaskInput:
    """Поля задачи или напоминания. Связи можно не заполнять."""

    title: str
    details: str = ""
    kind: str = TaskKind.TASK.value
    due_at: str = ""
    is_done: bool = False
    client_id: int | None = None
    deal_id: int | None = None
    order_id: int | None = None


def format_amount(amount_minor: int) -> str:
    """Показать сумму с двумя знаками. Для рубля это рубли и копейки."""
    if amount_minor < 0:
        sign = "-"
        minor = -amount_minor
    else:
        sign = ""
        minor = amount_minor
    whole, fraction = divmod(minor, MINOR_SCALE)
    return f"{sign}{whole}.{fraction:02d}"


def _required_int(row: Mapping[str, Any], key: str) -> int:
    value = row[key]
    if isinstance(value, bool) or not isinstance(value, int):
        raise TypeError(f"Колонка {key} должна быть INTEGER.")
    return value


def _optional_int(row: Mapping[str, Any], key: str) -> int | None:
    value = row[key]
    if value is None:
        return None
    if isinstance(value, bool) or not isinstance(value, int):
        raise TypeError(f"Колонка {key} должна быть INTEGER.")
    return value


def _required_str(row: Mapping[str, Any], key: str) -> str:
    value = row[key]
    if not isinstance(value, str):
        raise TypeError(f"Колонка {key} должна быть TEXT.")
    return value


def _joined_str(row: Mapping[str, Any], key: str) -> str:
    """Имя из JOIN. Если колонки нет, связь просто не подгружали."""
    try:
        value = row[key]
    except (KeyError, IndexError):
        return ""
    if value is None:
        return ""
    if not isinstance(value, str):
        raise TypeError(f"Колонка {key} должна быть TEXT.")
    return value


def _flag(row: Mapping[str, Any], key: str) -> bool:
    value = row[key]
    if value in (0, 1):
        return bool(value)
    raise TypeError(f"Колонка {key} должна быть 0 или 1.")


@dataclass(frozen=True)
class Client:
    """Строка таблицы clients."""

    id: int
    full_name: str
    company: str
    email: str
    phone: str
    status: str
    source: str
    notes: str
    created_at: str
    updated_at: str

    @classmethod
    def from_mapping(cls, row: Mapping[str, Any]) -> Client:
        """Собрать клиента из строки SQLite."""
        return cls(
            id=_required_int(row, "id"),
            full_name=_required_str(row, "full_name"),
            company=_required_str(row, "company"),
            email=_required_str(row, "email"),
            phone=_required_str(row, "phone"),
            status=_required_str(row, "status"),
            source=_required_str(row, "source"),
            notes=_required_str(row, "notes"),
            created_at=_required_str(row, "created_at"),
            updated_at=_required_str(row, "updated_at"),
        )

    def to_dict(self) -> dict[str, Any]:
        """Словарь для ответа API."""
        return {
            "id": self.id,
            "full_name": self.full_name,
            "company": self.company,
            "email": self.email,
            "phone": self.phone,
            "status": self.status,
            "source": self.source,
            "notes": self.notes,
            "created_at": self.created_at,
            "updated_at": self.updated_at,
        }


@dataclass(frozen=True)
class Deal:
    """Строка таблицы deals плюс имя клиента, если он прикреплён."""

    id: int
    title: str
    client_id: int | None
    status: str
    amount_minor: int
    currency: str
    expected_close_on: str
    notes: str
    created_at: str
    updated_at: str
    client_name: str = ""

    @classmethod
    def from_mapping(cls, row: Mapping[str, Any]) -> Deal:
        """Собрать сделку из строки SQLite."""
        return cls(
            id=_required_int(row, "id"),
            title=_required_str(row, "title"),
            client_id=_optional_int(row, "client_id"),
            status=_required_str(row, "status"),
            amount_minor=_required_int(row, "amount_minor"),
            currency=_required_str(row, "currency"),
            expected_close_on=_required_str(row, "expected_close_on"),
            notes=_required_str(row, "notes"),
            created_at=_required_str(row, "created_at"),
            updated_at=_required_str(row, "updated_at"),
            client_name=_joined_str(row, "client_name"),
        )

    def to_dict(self) -> dict[str, Any]:
        """Словарь для ответа API."""
        return {
            "id": self.id,
            "title": self.title,
            "client_id": self.client_id,
            "client_name": self.client_name,
            "status": self.status,
            "amount_minor": self.amount_minor,
            "amount": format_amount(self.amount_minor),
            "currency": self.currency,
            "expected_close_on": self.expected_close_on,
            "notes": self.notes,
            "created_at": self.created_at,
            "updated_at": self.updated_at,
        }


@dataclass(frozen=True)
class Order:
    """Строка таблицы orders плюс подписи клиента и сделки."""

    id: int
    title: str
    client_id: int | None
    deal_id: int | None
    status: str
    amount_minor: int
    currency: str
    due_on: str
    notes: str
    created_at: str
    updated_at: str
    client_name: str = ""
    deal_title: str = ""

    @classmethod
    def from_mapping(cls, row: Mapping[str, Any]) -> Order:
        """Собрать заказ из строки SQLite."""
        return cls(
            id=_required_int(row, "id"),
            title=_required_str(row, "title"),
            client_id=_optional_int(row, "client_id"),
            deal_id=_optional_int(row, "deal_id"),
            status=_required_str(row, "status"),
            amount_minor=_required_int(row, "amount_minor"),
            currency=_required_str(row, "currency"),
            due_on=_required_str(row, "due_on"),
            notes=_required_str(row, "notes"),
            created_at=_required_str(row, "created_at"),
            updated_at=_required_str(row, "updated_at"),
            client_name=_joined_str(row, "client_name"),
            deal_title=_joined_str(row, "deal_title"),
        )

    def to_dict(self) -> dict[str, Any]:
        """Словарь для ответа API."""
        return {
            "id": self.id,
            "title": self.title,
            "client_id": self.client_id,
            "client_name": self.client_name,
            "deal_id": self.deal_id,
            "deal_title": self.deal_title,
            "status": self.status,
            "amount_minor": self.amount_minor,
            "amount": format_amount(self.amount_minor),
            "currency": self.currency,
            "due_on": self.due_on,
            "notes": self.notes,
            "created_at": self.created_at,
            "updated_at": self.updated_at,
        }


@dataclass(frozen=True)
class Task:
    """Строка таблицы tasks. is_done в базе INTEGER, здесь bool."""

    id: int
    title: str
    details: str
    kind: str
    due_at: str
    is_done: bool
    client_id: int | None
    deal_id: int | None
    order_id: int | None
    created_at: str
    updated_at: str
    completed_at: str
    client_name: str = ""
    deal_title: str = ""
    order_title: str = ""

    @classmethod
    def from_mapping(cls, row: Mapping[str, Any]) -> Task:
        """Собрать задачу из строки SQLite."""
        return cls(
            id=_required_int(row, "id"),
            title=_required_str(row, "title"),
            details=_required_str(row, "details"),
            kind=_required_str(row, "kind"),
            due_at=_required_str(row, "due_at"),
            is_done=_flag(row, "is_done"),
            client_id=_optional_int(row, "client_id"),
            deal_id=_optional_int(row, "deal_id"),
            order_id=_optional_int(row, "order_id"),
            created_at=_required_str(row, "created_at"),
            updated_at=_required_str(row, "updated_at"),
            completed_at=_required_str(row, "completed_at"),
            client_name=_joined_str(row, "client_name"),
            deal_title=_joined_str(row, "deal_title"),
            order_title=_joined_str(row, "order_title"),
        )

    def to_dict(self) -> dict[str, Any]:
        """Словарь для ответа API."""
        return {
            "id": self.id,
            "title": self.title,
            "details": self.details,
            "kind": self.kind,
            "due_at": self.due_at,
            "is_done": self.is_done,
            "client_id": self.client_id,
            "client_name": self.client_name,
            "deal_id": self.deal_id,
            "deal_title": self.deal_title,
            "order_id": self.order_id,
            "order_title": self.order_title,
            "created_at": self.created_at,
            "updated_at": self.updated_at,
            "completed_at": self.completed_at,
        }


class Record(Protocol):
    """Объект, который умеет отдать себя словарём."""

    def to_dict(self) -> dict[str, Any]:
        """Словарь для JSON."""


TRecord = TypeVar("TRecord", bound=Record)


@dataclass(frozen=True)
class Page(Generic[TRecord]):
    """Страница списка: записи и общее число без учёта limit."""

    items: tuple[TRecord, ...]
    total: int
    limit: int
    offset: int

    def to_dict(self) -> dict[str, Any]:
        """Словарь для ответа API."""
        return {
            "items": [item.to_dict() for item in self.items],
            "total": self.total,
            "limit": self.limit,
            "offset": self.offset,
        }


def _labeled(codes: tuple[str, ...], labels: dict[str, str]) -> list[dict[str, str]]:
    return [{"code": code, "label": labels[code]} for code in codes]


def catalog() -> dict[str, Any]:
    """Справочник кодов, подписей и правил денег для интерфейса."""
    return {
        "clients": {"statuses": _labeled(CLIENT_STATUSES, CLIENT_STATUS_LABELS)},
        "deals": {"statuses": _labeled(DEAL_STATUSES, DEAL_STATUS_LABELS)},
        "orders": {"statuses": _labeled(ORDER_STATUSES, ORDER_STATUS_LABELS)},
        "tasks": {"kinds": _labeled(TASK_KINDS, TASK_KIND_LABELS)},
        "money": {
            "minor_scale": MINOR_SCALE,
            "default_currency": DEFAULT_CURRENCY,
            "note": (
                "amount_minor — INTEGER. 100 единиц равны 1.00. "
                "Для рубля это копейки."
            ),
        },
        "search": {
            "operator": "ilike",
            "note": (
                "Поиск без учёта регистра по образцу ILIKE. "
                "В SQLite такого оператора нет, класс базы регистрирует функцию ilike."
            ),
        },
    }
