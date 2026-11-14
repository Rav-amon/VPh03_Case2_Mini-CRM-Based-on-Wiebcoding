"""Локальный API мини-CRM.

Сервер слушает только 127.0.0.1 и не проверяет вход: это личная база
на компьютере, в сеть её выставлять не нужно.

Запуск:

    .\\.venv\\Scripts\\python.exe api.py

Документация: http://127.0.0.1:8000/docs
Файл базы: CRM_DB_PATH или data/crm.sqlite рядом с модулем.
"""

from __future__ import annotations

import json
import logging
from typing import Any

from fastapi import APIRouter, Depends, FastAPI, Query, Request
from fastapi.exceptions import RequestValidationError
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from pydantic import BaseModel, ConfigDict, Field
from starlette.exceptions import HTTPException as StarletteHTTPException

from crm import CrmDatabase, CrmError, default_db_path
from models import (
    COMPANY_MAX_LENGTH,
    DEFAULT_PAGE_LIMIT,
    DETAILS_MAX_LENGTH,
    EMAIL_MAX_LENGTH,
    MAX_AMOUNT_MINOR,
    MAX_PAGE_LIMIT,
    NAME_MAX_LENGTH,
    NOTES_MAX_LENGTH,
    PHONE_MAX_LENGTH,
    QUERY_MAX_LENGTH,
    SOURCE_MAX_LENGTH,
    TITLE_MAX_LENGTH,
    ClientInput,
    ClientStatus,
    DealInput,
    DealStatus,
    OrderInput,
    OrderStatus,
    TaskInput,
    TaskKind,
    catalog,
)

logger = logging.getLogger("crm.api")

API_HOST = "127.0.0.1"
API_PORT = 8000

_STATUS_BY_CODE = {
    "NOT_FOUND": 404,
    "VALIDATION_ERROR": 422,
    "CONFLICT": 409,
    "DATABASE_ERROR": 500,
}

_store: CrmDatabase | None = None


class UnicodeJSONResponse(JSONResponse):
    """JSON с русскими буквами как есть, без экранирования \\uXXXX."""

    def render(self, content: Any) -> bytes:
        return json.dumps(content, ensure_ascii=False).encode("utf-8")


def init_store(db_path: str | None = None) -> CrmDatabase:
    """Открыть файл базы и запомнить его для зависимостей FastAPI."""
    global _store
    if db_path is None:
        _store = CrmDatabase(default_db_path())
    else:
        _store = CrmDatabase(db_path)
    return _store


def get_db() -> CrmDatabase:
    """Отдать уже открытую базу или создать файл по умолчанию."""
    if _store is None:
        return init_store()
    return _store


def _ok(data: Any, message: str) -> dict[str, Any]:
    return {"data": data, "error": None, "message": message}


def _fail(status_code: int, code: str, message: str, details: Any) -> JSONResponse:
    return UnicodeJSONResponse(
        status_code=status_code,
        content={
            "data": None,
            "error": {"code": code, "message": message, "details": details},
            "message": message,
        },
    )


class ClientBody(BaseModel):
    """Карточка клиента для создания."""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    full_name: str = Field(
        min_length=1,
        max_length=NAME_MAX_LENGTH,
        examples=["Иванова Анна"],
    )
    company: str = Field(default="", max_length=COMPANY_MAX_LENGTH)
    email: str = Field(default="", max_length=EMAIL_MAX_LENGTH)
    phone: str = Field(default="", max_length=PHONE_MAX_LENGTH)
    status: ClientStatus = ClientStatus.ACTIVE
    source: str = Field(default="", max_length=SOURCE_MAX_LENGTH)
    notes: str = Field(default="", max_length=NOTES_MAX_LENGTH)


class ClientUpdate(BaseModel):
    """Полная замена карточки. Все поля нужно прислать, даже пустые строки."""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    full_name: str = Field(min_length=1, max_length=NAME_MAX_LENGTH)
    company: str = Field(max_length=COMPANY_MAX_LENGTH)
    email: str = Field(max_length=EMAIL_MAX_LENGTH)
    phone: str = Field(max_length=PHONE_MAX_LENGTH)
    status: ClientStatus
    source: str = Field(max_length=SOURCE_MAX_LENGTH)
    notes: str = Field(max_length=NOTES_MAX_LENGTH)


class ClientStatusBody(BaseModel):
    """Смена статуса. Архив — код archived."""

    model_config = ConfigDict(extra="forbid")

    status: ClientStatus


class DealBody(BaseModel):
    """Новая сделка. client_id можно не передавать."""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    title: str = Field(min_length=1, max_length=TITLE_MAX_LENGTH, examples=["Сайт-визитка"])
    client_id: int | None = Field(default=None, ge=1)
    status: DealStatus = DealStatus.NEW
    amount_minor: int = Field(default=0, ge=0, le=MAX_AMOUNT_MINOR)
    currency: str = Field(default="RUB", min_length=3, max_length=3)
    expected_close_on: str = Field(default="", max_length=40)
    notes: str = Field(default="", max_length=NOTES_MAX_LENGTH)


class DealUpdate(BaseModel):
    """Полная замена сделки. client_id: null открепляет клиента."""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    title: str = Field(min_length=1, max_length=TITLE_MAX_LENGTH)
    client_id: int | None = Field(ge=1)
    status: DealStatus
    amount_minor: int = Field(ge=0, le=MAX_AMOUNT_MINOR)
    currency: str = Field(min_length=3, max_length=3)
    expected_close_on: str = Field(max_length=40)
    notes: str = Field(max_length=NOTES_MAX_LENGTH)


class DealStatusBody(BaseModel):
    """Смена статуса сделки."""

    model_config = ConfigDict(extra="forbid")

    status: DealStatus


class DealClientBody(BaseModel):
    """Прикрепить клиента или открепить его значением null."""

    model_config = ConfigDict(extra="forbid")

    client_id: int | None = Field(ge=1)


class OrderBody(BaseModel):
    """Новый заказ. Клиент и сделка необязательны."""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    title: str = Field(min_length=1, max_length=TITLE_MAX_LENGTH)
    client_id: int | None = Field(default=None, ge=1)
    deal_id: int | None = Field(default=None, ge=1)
    status: OrderStatus = OrderStatus.DRAFT
    amount_minor: int = Field(default=0, ge=0, le=MAX_AMOUNT_MINOR)
    currency: str = Field(default="RUB", min_length=3, max_length=3)
    due_on: str = Field(default="", max_length=40)
    notes: str = Field(default="", max_length=NOTES_MAX_LENGTH)


class OrderUpdate(BaseModel):
    """Полная замена заказа."""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    title: str = Field(min_length=1, max_length=TITLE_MAX_LENGTH)
    client_id: int | None = Field(ge=1)
    deal_id: int | None = Field(ge=1)
    status: OrderStatus
    amount_minor: int = Field(ge=0, le=MAX_AMOUNT_MINOR)
    currency: str = Field(min_length=3, max_length=3)
    due_on: str = Field(max_length=40)
    notes: str = Field(max_length=NOTES_MAX_LENGTH)


class OrderStatusBody(BaseModel):
    """Смена статуса заказа."""

    model_config = ConfigDict(extra="forbid")

    status: OrderStatus


class TaskBody(BaseModel):
    """Новая задача. Для kind=reminder нужен due_at."""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    title: str = Field(min_length=1, max_length=TITLE_MAX_LENGTH)
    details: str = Field(default="", max_length=DETAILS_MAX_LENGTH)
    kind: TaskKind = TaskKind.TASK
    due_at: str = Field(default="", max_length=40)
    is_done: bool = False
    client_id: int | None = Field(default=None, ge=1)
    deal_id: int | None = Field(default=None, ge=1)
    order_id: int | None = Field(default=None, ge=1)


class TaskUpdate(BaseModel):
    """Полная замена задачи. Пустой идентификатор связи передаётся как null."""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    title: str = Field(min_length=1, max_length=TITLE_MAX_LENGTH)
    details: str = Field(max_length=DETAILS_MAX_LENGTH)
    kind: TaskKind
    due_at: str = Field(max_length=40)
    is_done: bool
    client_id: int | None = Field(ge=1)
    deal_id: int | None = Field(ge=1)
    order_id: int | None = Field(ge=1)


class TaskDoneBody(BaseModel):
    """Отметка выполнения."""

    model_config = ConfigDict(extra="forbid")

    is_done: bool


def _client_input(body: ClientBody | ClientUpdate) -> ClientInput:
    return ClientInput(
        full_name=body.full_name,
        company=body.company,
        email=body.email,
        phone=body.phone,
        status=body.status.value,
        source=body.source,
        notes=body.notes,
    )


def _deal_input(body: DealBody | DealUpdate) -> DealInput:
    return DealInput(
        title=body.title,
        client_id=body.client_id,
        status=body.status.value,
        amount_minor=body.amount_minor,
        currency=body.currency,
        expected_close_on=body.expected_close_on,
        notes=body.notes,
    )


def _order_input(body: OrderBody | OrderUpdate) -> OrderInput:
    return OrderInput(
        title=body.title,
        client_id=body.client_id,
        deal_id=body.deal_id,
        status=body.status.value,
        amount_minor=body.amount_minor,
        currency=body.currency,
        due_on=body.due_on,
        notes=body.notes,
    )


def _task_input(body: TaskBody | TaskUpdate) -> TaskInput:
    return TaskInput(
        title=body.title,
        details=body.details,
        kind=body.kind.value,
        due_at=body.due_at,
        is_done=body.is_done,
        client_id=body.client_id,
        deal_id=body.deal_id,
        order_id=body.order_id,
    )


def _pagination(
    limit: int = Query(DEFAULT_PAGE_LIMIT, ge=1, le=MAX_PAGE_LIMIT),
    offset: int = Query(0, ge=0),
) -> tuple[int, int]:
    """Общий размер страницы для списков и поиска."""
    return limit, offset


system = APIRouter(prefix="/api/v1", tags=["Сводка"])
clients = APIRouter(prefix="/api/v1/clients", tags=["Клиенты"])
deals = APIRouter(prefix="/api/v1/deals", tags=["Сделки"])
orders = APIRouter(prefix="/api/v1/orders", tags=["Заказы"])
tasks = APIRouter(prefix="/api/v1/tasks", tags=["Задачи"])


@system.get("/health")
def health(db: CrmDatabase = Depends(get_db)) -> dict[str, Any]:
    """Проверить, что процесс жив и файл базы открыт."""
    names = db.summary()
    return _ok(
        {"database_file": db.db_path.name, "sections": sorted(names)},
        "База готова.",
    )


@system.get("/meta")
def meta() -> dict[str, Any]:
    """Коды статусов, виды задач и правило суммы."""
    return _ok(catalog(), "Справочник.")


@system.get("/summary")
def summary(db: CrmDatabase = Depends(get_db)) -> dict[str, Any]:
    """Счётчики клиентов, сделок, заказов и открытых задач."""
    return _ok(db.summary(), "Сводка.")


@clients.get("")
def list_clients(
    status: ClientStatus | None = Query(default=None),
    page: tuple[int, int] = Depends(_pagination),
    db: CrmDatabase = Depends(get_db),
) -> dict[str, Any]:
    """Список клиентов. Без status возвращаются и архивные."""
    limit, offset = page
    status_value = None if status is None else status.value
    found = db.list_clients(status=status_value, limit=limit, offset=offset)
    return _ok(found.to_dict(), "Список клиентов.")


@clients.get("/search")
def search_clients(
    q: str = Query(min_length=1, max_length=QUERY_MAX_LENGTH),
    page: tuple[int, int] = Depends(_pagination),
    db: CrmDatabase = Depends(get_db),
) -> dict[str, Any]:
    """Поиск клиентов без учёта регистра, аналог ILIKE."""
    limit, offset = page
    found = db.search_clients(q, limit=limit, offset=offset)
    return _ok(found.to_dict(), "Найденные клиенты.")


@clients.post("")
def create_client(body: ClientBody, db: CrmDatabase = Depends(get_db)) -> dict[str, Any]:
    """Создать клиента."""
    created = db.create_client(_client_input(body))
    return _ok(created.to_dict(), "Клиент создан.")


@clients.get("/{client_id}")
def get_client(client_id: int, db: CrmDatabase = Depends(get_db)) -> dict[str, Any]:
    """Карточка клиента."""
    return _ok(db.get_client(client_id).to_dict(), "Клиент.")


@clients.put("/{client_id}")
def update_client(
    client_id: int,
    body: ClientUpdate,
    db: CrmDatabase = Depends(get_db),
) -> dict[str, Any]:
    """Полностью заменить карточку клиента."""
    updated = db.update_client(client_id, _client_input(body))
    return _ok(updated.to_dict(), "Клиент изменён.")


@clients.patch("/{client_id}/status")
def set_client_status(
    client_id: int,
    body: ClientStatusBody,
    db: CrmDatabase = Depends(get_db),
) -> dict[str, Any]:
    """Сменить статус. Для архива передайте archived."""
    updated = db.set_client_status(client_id, body.status.value)
    return _ok(updated.to_dict(), "Статус клиента изменён.")


@clients.delete("/{client_id}")
def delete_client(client_id: int, db: CrmDatabase = Depends(get_db)) -> dict[str, Any]:
    """Удалить клиента. Связанные сделки и задачи останутся без него."""
    deleted_id = db.delete_client(client_id)
    return _ok({"id": deleted_id}, "Клиент удалён.")


@deals.get("")
def list_deals(
    status: DealStatus | None = Query(default=None),
    client_id: int | None = Query(default=None, ge=1),
    page: tuple[int, int] = Depends(_pagination),
    db: CrmDatabase = Depends(get_db),
) -> dict[str, Any]:
    """Список сделок, можно сузить по статусу или клиенту."""
    limit, offset = page
    status_value = None if status is None else status.value
    found = db.list_deals(
        status=status_value,
        client_id=client_id,
        limit=limit,
        offset=offset,
    )
    return _ok(found.to_dict(), "Список сделок.")


@deals.get("/search")
def search_deals(
    q: str = Query(min_length=1, max_length=QUERY_MAX_LENGTH),
    page: tuple[int, int] = Depends(_pagination),
    db: CrmDatabase = Depends(get_db),
) -> dict[str, Any]:
    """Поиск сделок и прикреплённых клиентов без учёта регистра."""
    limit, offset = page
    found = db.search_deals(q, limit=limit, offset=offset)
    return _ok(found.to_dict(), "Найденные сделки.")


@deals.post("")
def create_deal(body: DealBody, db: CrmDatabase = Depends(get_db)) -> dict[str, Any]:
    """Создать сделку. Клиент необязателен."""
    created = db.create_deal(_deal_input(body))
    return _ok(created.to_dict(), "Сделка создана.")


@deals.get("/{deal_id}")
def get_deal(deal_id: int, db: CrmDatabase = Depends(get_db)) -> dict[str, Any]:
    """Карточка сделки."""
    return _ok(db.get_deal(deal_id).to_dict(), "Сделка.")


@deals.put("/{deal_id}")
def update_deal(
    deal_id: int,
    body: DealUpdate,
    db: CrmDatabase = Depends(get_db),
) -> dict[str, Any]:
    """Полностью заменить сделку."""
    updated = db.update_deal(deal_id, _deal_input(body))
    return _ok(updated.to_dict(), "Сделка изменена.")


@deals.patch("/{deal_id}/status")
def set_deal_status(
    deal_id: int,
    body: DealStatusBody,
    db: CrmDatabase = Depends(get_db),
) -> dict[str, Any]:
    """Сменить статус сделки."""
    updated = db.set_deal_status(deal_id, body.status.value)
    return _ok(updated.to_dict(), "Статус сделки изменён.")


@deals.patch("/{deal_id}/client")
def set_deal_client(
    deal_id: int,
    body: DealClientBody,
    db: CrmDatabase = Depends(get_db),
) -> dict[str, Any]:
    """Прикрепить клиента к сделке или открепить."""
    updated = db.set_deal_client(deal_id, body.client_id)
    return _ok(updated.to_dict(), "Клиент сделки обновлён.")


@deals.delete("/{deal_id}")
def delete_deal(deal_id: int, db: CrmDatabase = Depends(get_db)) -> dict[str, Any]:
    """Удалить сделку."""
    deleted_id = db.delete_deal(deal_id)
    return _ok({"id": deleted_id}, "Сделка удалена.")


@orders.get("")
def list_orders(
    status: OrderStatus | None = Query(default=None),
    client_id: int | None = Query(default=None, ge=1),
    deal_id: int | None = Query(default=None, ge=1),
    page: tuple[int, int] = Depends(_pagination),
    db: CrmDatabase = Depends(get_db),
) -> dict[str, Any]:
    """Список заказов."""
    limit, offset = page
    status_value = None if status is None else status.value
    found = db.list_orders(
        status=status_value,
        client_id=client_id,
        deal_id=deal_id,
        limit=limit,
        offset=offset,
    )
    return _ok(found.to_dict(), "Список заказов.")


@orders.get("/search")
def search_orders(
    q: str = Query(min_length=1, max_length=QUERY_MAX_LENGTH),
    page: tuple[int, int] = Depends(_pagination),
    db: CrmDatabase = Depends(get_db),
) -> dict[str, Any]:
    """Поиск заказов без учёта регистра."""
    limit, offset = page
    found = db.search_orders(q, limit=limit, offset=offset)
    return _ok(found.to_dict(), "Найденные заказы.")


@orders.post("")
def create_order(body: OrderBody, db: CrmDatabase = Depends(get_db)) -> dict[str, Any]:
    """Создать заказ."""
    created = db.create_order(_order_input(body))
    return _ok(created.to_dict(), "Заказ создан.")


@orders.get("/{order_id}")
def get_order(order_id: int, db: CrmDatabase = Depends(get_db)) -> dict[str, Any]:
    """Карточка заказа."""
    return _ok(db.get_order(order_id).to_dict(), "Заказ.")


@orders.put("/{order_id}")
def update_order(
    order_id: int,
    body: OrderUpdate,
    db: CrmDatabase = Depends(get_db),
) -> dict[str, Any]:
    """Полностью заменить заказ."""
    updated = db.update_order(order_id, _order_input(body))
    return _ok(updated.to_dict(), "Заказ изменён.")


@orders.patch("/{order_id}/status")
def set_order_status(
    order_id: int,
    body: OrderStatusBody,
    db: CrmDatabase = Depends(get_db),
) -> dict[str, Any]:
    """Сменить статус заказа."""
    updated = db.set_order_status(order_id, body.status.value)
    return _ok(updated.to_dict(), "Статус заказа изменён.")


@orders.delete("/{order_id}")
def delete_order(order_id: int, db: CrmDatabase = Depends(get_db)) -> dict[str, Any]:
    """Удалить заказ."""
    deleted_id = db.delete_order(order_id)
    return _ok({"id": deleted_id}, "Заказ удалён.")


@tasks.get("")
def list_tasks(
    client_id: int | None = Query(default=None, ge=1),
    deal_id: int | None = Query(default=None, ge=1),
    order_id: int | None = Query(default=None, ge=1),
    is_done: bool | None = Query(default=None),
    kind: TaskKind | None = Query(default=None),
    due_before: str | None = Query(default=None, max_length=40),
    page: tuple[int, int] = Depends(_pagination),
    db: CrmDatabase = Depends(get_db),
) -> dict[str, Any]:
    """Список задач и напоминаний по связям, виду и сроку."""
    limit, offset = page
    kind_value = None if kind is None else kind.value
    boundary = None if due_before is None or due_before.strip() == "" else due_before
    found = db.list_tasks(
        client_id=client_id,
        deal_id=deal_id,
        order_id=order_id,
        is_done=is_done,
        kind=kind_value,
        due_before=boundary,
        limit=limit,
        offset=offset,
    )
    return _ok(found.to_dict(), "Список задач.")


@tasks.post("")
def create_task(body: TaskBody, db: CrmDatabase = Depends(get_db)) -> dict[str, Any]:
    """Создать задачу или напоминание."""
    created = db.create_task(_task_input(body))
    return _ok(created.to_dict(), "Задача создана.")


@tasks.get("/{task_id}")
def get_task(task_id: int, db: CrmDatabase = Depends(get_db)) -> dict[str, Any]:
    """Карточка задачи."""
    return _ok(db.get_task(task_id).to_dict(), "Задача.")


@tasks.put("/{task_id}")
def update_task(
    task_id: int,
    body: TaskUpdate,
    db: CrmDatabase = Depends(get_db),
) -> dict[str, Any]:
    """Полностью заменить задачу."""
    updated = db.update_task(task_id, _task_input(body))
    return _ok(updated.to_dict(), "Задача изменена.")


@tasks.patch("/{task_id}/done")
def set_task_done(
    task_id: int,
    body: TaskDoneBody,
    db: CrmDatabase = Depends(get_db),
) -> dict[str, Any]:
    """Отметить задачу выполненной или открыть снова."""
    updated = db.set_task_done(task_id, body.is_done)
    return _ok(updated.to_dict(), "Отметка задачи изменена.")


@tasks.delete("/{task_id}")
def delete_task(task_id: int, db: CrmDatabase = Depends(get_db)) -> dict[str, Any]:
    """Удалить задачу."""
    deleted_id = db.delete_task(task_id)
    return _ok({"id": deleted_id}, "Задача удалена.")


app = FastAPI(
    title="Мини-CRM",
    version="1.0.0",
    description=(
        "Локальный MVP: клиенты, сделки, заказы и задачи. "
        "Слушает только этот компьютер."
    ),
    default_response_class=UnicodeJSONResponse,
)
app.add_middleware(
    CORSMiddleware,
    allow_origin_regex=r"http://(localhost|127\.0\.0\.1)(:\d+)?",
    allow_methods=["*"],
    allow_headers=["*"],
)
app.include_router(system)
app.include_router(clients)
app.include_router(deals)
app.include_router(orders)
app.include_router(tasks)


@app.exception_handler(CrmError)
async def handle_crm_error(request: Request, exc: CrmError) -> JSONResponse:
    """Ответ {data, error, message} для ошибок базы и проверок."""
    del request
    status_code = _STATUS_BY_CODE.get(exc.code, 400)
    logger.error(
        "%s",
        {"level": "error", "message": exc.message, "context": exc.to_dict()},
    )
    return _fail(status_code, exc.code, exc.message, exc.details)


@app.exception_handler(RequestValidationError)
async def handle_validation(request: Request, exc: RequestValidationError) -> JSONResponse:
    """Поля запроса, которые не прошли проверку схемы."""
    del request
    details = [
        {
            "field": ".".join(str(part) for part in item.get("loc", ())),
            "message": str(item.get("msg", "")),
            "type": str(item.get("type", "")),
        }
        for item in exc.errors()
    ]
    message = "Проверьте поля запроса."
    return _fail(422, "VALIDATION_ERROR", message, details)


# Роутер поднимает starlette.HTTPException, не подкласс FastAPI.
# Обработчик вешается на базовый класс, иначе 404 остаётся текстом «Not Found».
_HTTP_MESSAGES = {
    404: "Маршрут не найден.",
    405: "Метод не поддерживается.",
}


@app.exception_handler(StarletteHTTPException)
async def handle_http(request: Request, exc: StarletteHTTPException) -> JSONResponse:
    """Единый конверт и для ошибок маршрута, например 404."""
    del request
    message = _HTTP_MESSAGES.get(exc.status_code)
    if message is None and isinstance(exc.detail, str):
        message = exc.detail
    if message is None:
        message = "Запрос отклонён."
    return _fail(exc.status_code, "HTTP_ERROR", message, exc.detail)


@app.exception_handler(Exception)
async def handle_unexpected(request: Request, exc: Exception) -> JSONResponse:
    """Не отдавать наружу текст внутренней ошибки."""
    del request
    logger.error(
        "%s",
        {
            "level": "error",
            "message": "необработанная ошибка",
            "context": {"error": type(exc).__name__},
        },
        exc_info=exc,
    )
    return _fail(500, "INTERNAL_ERROR", "Внутренняя ошибка сервера.", None)


def main() -> None:
    """Поднять API на локальном порту."""
    import uvicorn

    init_store()
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    logger.info(
        "%s",
        {
            "level": "info",
            "message": "мини-CRM слушает только локальный адрес",
            "context": {"docs": f"http://{API_HOST}:{API_PORT}/docs"},
        },
    )
    uvicorn.run(app, host=API_HOST, port=API_PORT)


if __name__ == "__main__":
    main()
