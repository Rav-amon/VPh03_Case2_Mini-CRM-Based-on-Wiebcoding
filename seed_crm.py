"""Заполнить локальный API мини-CRM похожими на живые карточками.

Тысяча клиентов, тысяча сделок и тысяча задач. Скрипт только шлёт POST
на http://127.0.0.1:8000 и сам базу не открывает.

Запуск из корня проекта, когда сервер уже слушает порт:

    .\\.venv\\Scripts\\python.exe seed_crm.py
"""

from __future__ import annotations

import random
import sys
import time
from datetime import date, datetime, timedelta, timezone
from typing import Any

import requests

API_BASE = "http://127.0.0.1:8000"
API_PREFIX = "/api/v1"
RECORD_COUNT = 1000
TIMEOUT_SECONDS = 20
MAX_ATTEMPTS = 3
RETRY_PAUSE_SECONDS = 0.4
PROGRESS_EVERY = 100
DEAL_WITHOUT_CLIENT_PERCENT = 15
TASK_CLIENT_ONLY_PERCENT = 40
TASK_WITH_DEAL_PERCENT = 35
REMINDER_PERCENT = 40
DONE_PERCENT = 30
TRANSIENT_STATUSES = frozenset({429, 500, 502, 503})

CLIENT_STATUSES = (
    ("active", 50),
    ("lead", 25),
    ("inactive", 15),
    ("archived", 10),
)
DEAL_STATUSES = (
    ("new", 15),
    ("qualification", 15),
    ("proposal", 20),
    ("negotiation", 15),
    ("won", 15),
    ("lost", 12),
    ("archived", 8),
)
SOURCES = (
    "Сайт",
    "Рекомендация",
    "Холодный звонок",
    "Выставка",
    "Реклама",
    "Повторная покупка",
)
COMPANY_KINDS = ("ООО", "АО", "ИП")
COMPANY_NAMES = (
    "Северная линия",
    "Тихий двор",
    "Балтийский контур",
    "Полевая мастерская",
    "Ясный берег",
    "Медная полка",
    "Городской склад",
    "Лёгкий шаг",
    "Домашний очаг",
    "Речной причал",
    "Сибирский мёд",
    "Каменный мост",
    "Утренний хлеб",
    "Восточный двор",
    "Белая бумага",
    "Тёплый цех",
    "Новая верфь",
    "Степной ветер",
    "Книжная лавка",
    "Зелёный двор",
)
FEMALE_SURNAMES = (
    "Иванова",
    "Петрова",
    "Смирнова",
    "Кузнецова",
    "Соколова",
    "Попова",
    "Лебедева",
    "Козлова",
    "Новикова",
    "Морозова",
    "Волкова",
    "Соловьёва",
    "Васильева",
    "Зайцева",
    "Павлова",
    "Семёнова",
    "Голубева",
    "Виноградова",
    "Богданова",
    "Воробьёва",
)
MALE_SURNAMES = (
    "Иванов",
    "Петров",
    "Смирнов",
    "Кузнецов",
    "Соколов",
    "Попов",
    "Лебедев",
    "Козлов",
    "Новиков",
    "Морозов",
    "Волков",
    "Соловьёв",
    "Васильев",
    "Зайцев",
    "Павлов",
    "Семёнов",
    "Голубев",
    "Виноградов",
    "Богданов",
    "Воробьёв",
)
WOMEN = (
    ("Анна", "Сергеевна"),
    ("Мария", "Игоревна"),
    ("Елена", "Андреевна"),
    ("Ольга", "Петровна"),
    ("Наталья", "Викторовна"),
    ("Татьяна", "Алексеевна"),
    ("Ирина", "Дмитриевна"),
    ("Светлана", "Николаевна"),
    ("Екатерина", "Олеговна"),
    ("Дарья", "Ильинична"),
    ("Алина", "Романовна"),
    ("Юлия", "Павловна"),
)
MEN = (
    ("Иван", "Петрович"),
    ("Алексей", "Сергеевич"),
    ("Дмитрий", "Андреевич"),
    ("Сергей", "Игоревич"),
    ("Андрей", "Николаевич"),
    ("Николай", "Олегович"),
    ("Павел", "Викторович"),
    ("Роман", "Алексеевич"),
    ("Илья", "Дмитриевич"),
    ("Максим", "Иванович"),
    ("Артём", "Юрьевич"),
    ("Кирилл", "Борисович"),
)
DEAL_ACTIONS = (
    "Поставка оборудования",
    "Сопровождение сайта",
    "Внедрение учёта",
    "Рекламная кампания",
    "Ремонт склада",
    "Аудит договоров",
    "Обучение сотрудников",
    "Разработка витрины",
)
TASK_TITLES = (
    "Позвонить и уточнить бюджет",
    "Отправить коммерческое предложение",
    "Согласовать дату встречи",
    "Напомнить об оплате счёта",
    "Проверить реквизиты договора",
    "Уточнить состав поставки",
    "Подготовить акт выполненных работ",
    "Перенести созвон на удобное время",
)
EMAIL_DOMAINS = ("mail.ru", "yandex.ru", "gmail.com", "inbox.ru")
_TRANSLIT = {
    "а": "a",
    "б": "b",
    "в": "v",
    "г": "g",
    "д": "d",
    "е": "e",
    "ё": "e",
    "ж": "zh",
    "з": "z",
    "и": "i",
    "й": "y",
    "к": "k",
    "л": "l",
    "м": "m",
    "н": "n",
    "о": "o",
    "п": "p",
    "р": "r",
    "с": "s",
    "т": "t",
    "у": "u",
    "ф": "f",
    "х": "h",
    "ц": "ts",
    "ч": "ch",
    "ш": "sh",
    "щ": "sch",
    "ъ": "",
    "ы": "y",
    "ь": "",
    "э": "e",
    "ю": "yu",
    "я": "ya",
}


class SeedError(Exception):
    """Сервер отклонил карточку или не ответил."""

    def __init__(self, message: str) -> None:
        self.message = message
        super().__init__(message)


def transliterate(value: str) -> str:
    """Имя для почты: кириллица в латиницу, без пробелов."""
    letters: list[str] = []
    for char in value.casefold():
        letters.append(_TRANSLIT.get(char, char))
    cleaned = "".join(letters)
    return "".join(char for char in cleaned if char.isascii() and char.isalnum())


def weighted_code(rng: random.Random, options: tuple[tuple[str, int], ...]) -> str:
    """Один код статуса с учётом веса."""
    codes = [code for code, _weight in options]
    weights = [weight for _code, weight in options]
    chosen = rng.choices(codes, weights=weights, k=1)
    return chosen[0]


def make_phone(rng: random.Random) -> str:
    """Российский мобильный номер в допустимом для API виде."""
    number = rng.randint(9000000000, 9999999999)
    text = str(number)
    return f"+7 {text[0:3]} {text[3:6]}-{text[6:8]}-{text[8:10]}"


def make_email(rng: random.Random, surname: str, first_name: str, index: int) -> str:
    """Почта из фамилии и имени. Номер в конце убирает совпадения."""
    local = f"{transliterate(surname)}.{transliterate(first_name)}{index}"
    return f"{local}@{rng.choice(EMAIL_DOMAINS)}"


def make_person(rng: random.Random, index: int) -> tuple[str, str, str]:
    """ФИО. Чётный номер — женщина, нечётный — мужчина, чтобы фамилия совпала."""
    if index % 2 == 0:
        surname = rng.choice(FEMALE_SURNAMES)
        first_name, patronymic = rng.choice(WOMEN)
    else:
        surname = rng.choice(MALE_SURNAMES)
        first_name, patronymic = rng.choice(MEN)
    return surname, first_name, patronymic


def make_company(rng: random.Random) -> str:
    """Короткое название фирмы."""
    return f"{rng.choice(COMPANY_KINDS)} «{rng.choice(COMPANY_NAMES)}»"


def build_client(rng: random.Random, index: int) -> dict[str, Any]:
    """Карточка клиента для POST /clients."""
    surname, first_name, patronymic = make_person(rng, index)
    return {
        "full_name": f"{surname} {first_name} {patronymic}",
        "company": make_company(rng),
        "email": make_email(rng, surname, first_name, index + 1),
        "phone": make_phone(rng),
        "status": weighted_code(rng, CLIENT_STATUSES),
        "source": rng.choice(SOURCES),
        "notes": "Первый контакт с сайта или по рекомендации менеджера.",
    }


def build_deal(
    rng: random.Random,
    index: int,
    client_ids: list[int],
    today: date,
) -> dict[str, Any]:
    """Сделка. Часть карточек остаётся без клиента, как позволяет API."""
    client_id: int | None = None
    if rng.randint(1, 100) > DEAL_WITHOUT_CLIENT_PERCENT and client_ids:
        client_id = rng.choice(client_ids)
    amount_rubles = rng.randrange(15_000, 2_400_000, 500)
    close_on = today + timedelta(days=rng.randint(-40, 160))
    return {
        "title": f"{rng.choice(DEAL_ACTIONS)} №{index + 1}",
        "client_id": client_id,
        "status": weighted_code(rng, DEAL_STATUSES),
        "amount_minor": amount_rubles * 100,
        "currency": "RUB",
        "expected_close_on": close_on.isoformat(),
        "notes": "Сумма согласована предварительно, срок можно сдвинуть.",
    }


def build_task(
    rng: random.Random,
    index: int,
    client_ids: list[int],
    deals: list[tuple[int, int | None]],
    moment: datetime,
) -> dict[str, Any]:
    """Задача или напоминание. Напоминание всегда получает срок."""
    kind = "task"
    if rng.randint(1, 100) <= REMINDER_PERCENT:
        kind = "reminder"
    due = moment + timedelta(days=rng.randint(-20, 40), hours=rng.randint(9, 18))
    due_at = ""
    if kind == "reminder" or rng.randint(1, 100) <= 70:
        due_at = due.replace(minute=0, second=0, microsecond=0).isoformat()
    client_id, deal_id = _task_links(rng, client_ids, deals)
    return {
        "title": f"{rng.choice(TASK_TITLES)} ({index + 1})",
        "details": "Связаться с клиентом и записать итог в карточку.",
        "kind": kind,
        "due_at": due_at,
        "is_done": rng.randint(1, 100) <= DONE_PERCENT,
        "client_id": client_id,
        "deal_id": deal_id,
        "order_id": None,
    }


def _task_links(
    rng: random.Random,
    client_ids: list[int],
    deals: list[tuple[int, int | None]],
) -> tuple[int | None, int | None]:
    """Связь задачи: только клиент, сделка с её клиентом или без связи."""
    roll = rng.randint(1, 100)
    if roll <= TASK_CLIENT_ONLY_PERCENT and client_ids:
        return rng.choice(client_ids), None
    deal_ceiling = TASK_CLIENT_ONLY_PERCENT + TASK_WITH_DEAL_PERCENT
    if roll <= deal_ceiling and deals:
        deal_id, deal_client_id = rng.choice(deals)
        return deal_client_id, deal_id
    return None, None


def post_row(session: requests.Session, path: str, body: dict[str, Any]) -> dict[str, Any]:
    """POST одной карточки. Повтор только для обрыва сети и временных кодов."""
    url = f"{API_BASE}{API_PREFIX}{path}"
    last_message = "Сервер не ответил."
    for _attempt in range(MAX_ATTEMPTS):
        try:
            response = session.post(url, json=body, timeout=TIMEOUT_SECONDS)
        except requests.RequestException as error:
            last_message = str(error)
            time.sleep(RETRY_PAUSE_SECONDS)
            continue
        if response.status_code in TRANSIENT_STATUSES:
            last_message = response.text
            time.sleep(RETRY_PAUSE_SECONDS)
            continue
        return _created_row(response)
    raise SeedError(last_message)


def _created_row(response: requests.Response) -> dict[str, Any]:
    """Достать data из конверта или остановиться с текстом ошибки API."""
    try:
        payload = response.json()
    except ValueError as error:
        raise SeedError("Сервер вернул не JSON.") from error
    if not isinstance(payload, dict):
        raise SeedError("Сервер вернул пустой ответ.")
    error = payload.get("error")
    if response.status_code >= 400 or error is not None:
        message = payload.get("message")
        if not isinstance(message, str) or message == "":
            message = "Карточка не записана."
        raise SeedError(message)
    data = payload.get("data")
    if not isinstance(data, dict) or not isinstance(data.get("id"), int):
        raise SeedError("В ответе нет идентификатора новой записи.")
    return data


def check_health(session: requests.Session) -> None:
    """Убедиться, что API уже запущен, до первой записи."""
    try:
        response = session.get(f"{API_BASE}{API_PREFIX}/health", timeout=TIMEOUT_SECONDS)
    except requests.RequestException as error:
        raise SeedError(
            "API не отвечает на http://127.0.0.1:8000. Сначала запустите сервер."
        ) from error
    if response.status_code != 200:
        raise SeedError("Проверка /api/v1/health не прошла.")


def load_clients(session: requests.Session, rng: random.Random) -> list[int]:
    """Создать клиентов и вернуть их id."""
    created: list[int] = []
    for index in range(RECORD_COUNT):
        row = post_row(session, "/clients", build_client(rng, index))
        created.append(int(row["id"]))
        _progress("Клиенты", index + 1)
    return created


def load_deals(
    session: requests.Session,
    rng: random.Random,
    client_ids: list[int],
) -> list[tuple[int, int | None]]:
    """Создать сделки. Пара — id сделки и id её клиента, если он есть."""
    today = datetime.now(timezone.utc).date()
    created: list[tuple[int, int | None]] = []
    for index in range(RECORD_COUNT):
        body = build_deal(rng, index, client_ids, today)
        row = post_row(session, "/deals", body)
        created.append((int(row["id"]), body["client_id"]))
        _progress("Сделки", index + 1)
    return created


def load_tasks(
    session: requests.Session,
    rng: random.Random,
    client_ids: list[int],
    deals: list[tuple[int, int | None]],
) -> None:
    """Создать задачи и напоминания."""
    moment = datetime.now(timezone.utc).replace(microsecond=0)
    for index in range(RECORD_COUNT):
        body = build_task(rng, index, client_ids, deals, moment)
        post_row(session, "/tasks", body)
        _progress("Задачи", index + 1)


def _progress(title: str, done: int) -> None:
    """Короткая отметка, чтобы длинная заливка не молчала."""
    if done % PROGRESS_EVERY == 0 or done == RECORD_COUNT:
        print(f"{title}: {done} из {RECORD_COUNT}", flush=True)


def main() -> None:
    """Залить три раздела и напечатать итог."""
    rng = random.Random()
    session = requests.Session()
    try:
        check_health(session)
        client_ids = load_clients(session, rng)
        deals = load_deals(session, rng, client_ids)
        load_tasks(session, rng, client_ids, deals)
    except SeedError as error:
        print(error.message, file=sys.stderr)
        raise SystemExit(1) from error
    print(
        f"Готово: клиентов {len(client_ids)}, сделок {len(deals)}, задач {RECORD_COUNT}."
    )


if __name__ == "__main__":
    main()
