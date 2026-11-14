"""Клиент CRUD для файлов Google Диска через официальный Drive API.

Сервисный аккаунт, как в google_sheets.py:

    from google_drive import GoogleDriveClient

    drive = GoogleDriveClient.from_env()
    files = drive.read_all()

Личный аккаунт через OAuth2 создаёт документы и таблицы от имени пользователя:

    from google_drive import GoogleDriveUserClient

    user_drive = GoogleDriveUserClient.from_env(folder_id="id_папки")
    document = user_drive.create_document("Договор")
    sheet = user_drive.create_spreadsheet("Смета")
"""

from __future__ import annotations

import json
import os
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

from google.auth.exceptions import RefreshError
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials as UserCredentials
from google.oauth2.service_account import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Полный доступ нужен для создания и удаления. Чтение им тоже покрывается.
_DRIVE_SCOPE = "https://www.googleapis.com/auth/drive"

_FOLDER_ID_ENV = "GOOGLE_DRIVE_FOLDER_ID"
_OAUTH_CLIENT_ENV = "GOOGLE_OAUTH_CLIENT_SECRETS"
_OAUTH_TOKEN_ENV = "GOOGLE_OAUTH_TOKEN_PATH"
_DEFAULT_TOKEN_NAME = "google_oauth_token.json"
# Полный доступ к Диску нужен, чтобы положить файл в уже существующую папку пользователя.
_USER_DRIVE_SCOPES = ("https://www.googleapis.com/auth/drive",)
_OAUTH_LOCAL_PORT = 0
_CREDENTIALS_PATH_ENVS = (
    "GOOGLE_CREDENTIALS_PATH",
    "GOOGLE_APPLICATION_CREDENTIALS",
)
_FOLDER_URL_MARKER = "/folders/"

_PAGE_SIZE = 100
_MAX_LIST_PAGES = 50

_MAX_RETRY_COUNT = 3
_RETRY_BACKOFF_SECONDS = 1.0
_TRANSIENT_HTTP_STATUSES = frozenset({429, 500, 502, 503})

_HTTP_ERROR_CODES = {
    400: "BAD_REQUEST",
    401: "UNAUTHORIZED",
    403: "FORBIDDEN",
    404: "NOT_FOUND",
    429: "RATE_LIMIT",
    500: "INTERNAL_ERROR",
    502: "BAD_GATEWAY",
    503: "SERVICE_UNAVAILABLE",
}

_FILE_FIELDS = "id, name, mimeType, webViewLink, modifiedTime, parents, owners(emailAddress)"
_LIST_FIELDS = f"nextPageToken, files({_FILE_FIELDS})"

MIME_DOCUMENT = "application/vnd.google-apps.document"
MIME_SPREADSHEET = "application/vnd.google-apps.spreadsheet"
MIME_PRESENTATION = "application/vnd.google-apps.presentation"
MIME_FOLDER = "application/vnd.google-apps.folder"

_MIME_LABELS = {
    MIME_FOLDER: "Папка",
    MIME_SPREADSHEET: "Google Таблица",
    MIME_DOCUMENT: "Google Документ",
    MIME_PRESENTATION: "Google Презентация",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "Excel",
}

_STORAGE_QUOTA_MARKER = "storage quota"

ApiPayload = dict[str, Any]
RequestFactory = Callable[[], Any]

_MODULE_DIR = Path(__file__).resolve().parent


class GoogleDriveError(Exception):
    """Ошибка Диска в едином формате code / message / details."""

    def __init__(self, code: str, message: str, details: Any = None) -> None:
        self.code = code
        self.message = message
        self.details = details
        super().__init__(message)

    def to_dict(self) -> dict[str, Any]:
        """Вернуть ошибку словарём для логов и API-ответов."""
        return {
            "code": self.code,
            "message": self.message,
            "details": self.details,
        }


@dataclass(frozen=True)
class DriveFile:
    """Файл или папка, которые вернул Drive API."""

    id: str
    name: str
    mime_type: str
    web_view_link: str
    modified_time: str
    parents: tuple[str, ...] = ()
    owner_email: str = ""

    @property
    def kind_label(self) -> str:
        """Короткое русское имя типа для консоли."""
        return _MIME_LABELS.get(self.mime_type, self.mime_type or "Файл")


def normalize_folder_id(value: str) -> str:
    """Достать идентификатор из ссылки на папку или вернуть уже готовый id."""
    cleaned = value.strip()
    if _FOLDER_URL_MARKER not in cleaned:
        return cleaned
    tail = cleaned.split(_FOLDER_URL_MARKER, 1)[1]
    folder_id = tail.split("/", 1)[0].split("?", 1)[0].strip()
    if not folder_id:
        raise GoogleDriveError(
            "VALIDATION_ERROR",
            "В ссылке на папку нет идентификатора.",
            {"value": value},
        )
    return folder_id


def _load_env_file(path: Path) -> None:
    """Подставить переменные из .env, не затирая уже заданное окружение."""
    if not path.is_file():
        return
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def _read_json(path: Path) -> ApiPayload:
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError as error:
        raise GoogleDriveError(
            "CREDENTIALS_NOT_FOUND",
            "Файл ключа сервисного аккаунта не найден.",
            {"path": str(path)},
        ) from error
    except json.JSONDecodeError as error:
        raise GoogleDriveError(
            "CREDENTIALS_INVALID",
            "JSON-ключ сервисного аккаунта повреждён.",
            {"path": str(path)},
        ) from error
    if not isinstance(payload, dict):
        raise GoogleDriveError(
            "CREDENTIALS_INVALID",
            "JSON-ключ должен быть объектом.",
            {"path": str(path)},
        )
    return payload


def resolve_credentials_path(credentials_path: str | Path | None = None) -> Path:
    """Найти JSON-ключ: явный путь, переменная окружения или единственный ключ в проекте."""
    if credentials_path is not None and str(credentials_path).strip():
        path = Path(credentials_path)
        if not path.is_absolute():
            path = _MODULE_DIR / path
        return path

    for env_name in _CREDENTIALS_PATH_ENVS:
        env_value = os.environ.get(env_name, "").strip()
        if not env_value:
            continue
        path = Path(env_value)
        if not path.is_absolute():
            path = _MODULE_DIR / path
        return path

    matches: list[Path] = []
    for path in sorted(_MODULE_DIR.glob("*.json")):
        try:
            payload = _read_json(path)
        except GoogleDriveError:
            continue
        if payload.get("type") == "service_account":
            matches.append(path)

    if len(matches) == 1:
        return matches[0]
    if not matches:
        raise GoogleDriveError(
            "CREDENTIALS_NOT_FOUND",
            "Не найден JSON-ключ сервисного аккаунта. Укажите credentials_path.",
            None,
        )
    raise GoogleDriveError(
        "AMBIGUOUS_CREDENTIALS",
        "В папке несколько ключей сервисного аккаунта. Укажите credentials_path явно.",
        {"files": [path.name for path in matches]},
    )


def service_account_email(credentials_path: str | Path | None = None) -> str:
    """Вернуть email сервисного аккаунта, которому нужно открыть папку Диска."""
    path = resolve_credentials_path(credentials_path)
    email = _read_json(path).get("client_email")
    if not isinstance(email, str) or not email:
        raise GoogleDriveError(
            "CREDENTIALS_INVALID",
            "В JSON-ключе нет client_email.",
            {"path": str(path)},
        )
    return email


def _credentials(path: Path) -> Credentials:
    try:
        return Credentials.from_service_account_file(str(path), scopes=[_DRIVE_SCOPE])
    except (ValueError, OSError) as error:
        raise GoogleDriveError(
            "CREDENTIALS_INVALID",
            "Не удалось прочитать ключ сервисного аккаунта.",
            {"path": str(path)},
        ) from error


def _is_oauth_client(payload: ApiPayload) -> bool:
    """Проверить, что JSON — клиент OAuth (installed или web), а не ключ сервисного аккаунта."""
    return isinstance(payload.get("installed"), dict) or isinstance(payload.get("web"), dict)


def resolve_oauth_client_path(client_secrets_path: str | Path | None = None) -> Path:
    """Найти JSON OAuth-клиента: явный путь, переменная окружения или единственный client secret."""
    if client_secrets_path is not None and str(client_secrets_path).strip():
        path = Path(client_secrets_path)
        if not path.is_absolute():
            path = _MODULE_DIR / path
        return path

    env_value = os.environ.get(_OAUTH_CLIENT_ENV, "").strip()
    if env_value:
        path = Path(env_value)
        if not path.is_absolute():
            path = _MODULE_DIR / path
        return path

    matches: list[Path] = []
    for path in sorted(_MODULE_DIR.glob("*.json")):
        try:
            payload = _read_json(path)
        except GoogleDriveError:
            continue
        if _is_oauth_client(payload):
            matches.append(path)

    if len(matches) == 1:
        return matches[0]
    if not matches:
        raise GoogleDriveError(
            "OAUTH_CLIENT_NOT_FOUND",
            "Не найден JSON OAuth-клиента. Укажите client_secrets_path.",
            None,
        )
    raise GoogleDriveError(
        "AMBIGUOUS_OAUTH_CLIENT",
        "В папке несколько OAuth-клиентов. Укажите client_secrets_path явно.",
        {"files": [path.name for path in matches]},
    )


def _resolve_token_path(token_path: str | Path | None) -> Path:
    """Путь к сохранённому токену пользователя."""
    if token_path is not None and str(token_path).strip():
        path = Path(token_path)
    else:
        env_value = os.environ.get(_OAUTH_TOKEN_ENV, "").strip()
        path = Path(env_value) if env_value else Path(_DEFAULT_TOKEN_NAME)
    if not path.is_absolute():
        path = _MODULE_DIR / path
    return path


def _token_is_usable(credentials: UserCredentials) -> bool:
    """Токен ещё живой и покрывает доступ к Диску."""
    if not credentials.valid:
        return False
    try:
        return bool(credentials.has_scopes(list(_USER_DRIVE_SCOPES)))
    except (ValueError, OSError):
        return False


def _read_saved_user_token(token_path: Path) -> UserCredentials | None:
    """Прочитать ранее выданный токен. Битый файл не мешает новому входу."""
    if not token_path.is_file():
        return None
    try:
        # Scope не подменяем: в файле лежат реально выданные права.
        return UserCredentials.from_authorized_user_file(str(token_path))
    except (ValueError, OSError):
        return None


def _save_user_token(token_path: Path, credentials: UserCredentials) -> None:
    """Сохранить токен, чтобы следующий запуск не открывал браузер."""
    token_path.write_text(credentials.to_json(), encoding="utf-8")


def _authorize_in_browser(client_secrets_path: Path) -> UserCredentials:
    """Открыть локальный вход Google и вернуть учётные данные пользователя."""
    try:
        flow = InstalledAppFlow.from_client_secrets_file(
            str(client_secrets_path),
            list(_USER_DRIVE_SCOPES),
        )
        # 127.0.0.1, а не localhost: на Windows браузер открывает localhost
        # по IPv6, а сервер входа слушает только IPv4, и страница не открывается.
        credentials = flow.run_local_server(
            host="127.0.0.1",
            bind_addr="127.0.0.1",
            port=_OAUTH_LOCAL_PORT,
            open_browser=True,
            prompt="consent",
            authorization_prompt_message=(
                "Войдите в Google в браузере, чтобы создавать файлы от своего имени:\n{url}"
            ),
            success_message="Вход выполнен. Это окно можно закрыть и вернуться в программу.",
        )
    except Exception as error:
        raise GoogleDriveError(
            "OAUTH_FAILED",
            "Не удалось войти через OAuth2.",
            {"reason": str(error)},
        ) from error
    if not isinstance(credentials, UserCredentials):
        raise GoogleDriveError("OAUTH_FAILED", "Google не вернул токен пользователя.", None)
    return credentials


def load_user_credentials(
    client_secrets_path: Path,
    token_path: Path,
) -> UserCredentials:
    """Вернуть OAuth-учётные данные пользователя.

    Сначала берётся сохранённый токен. Если он истёк, обновляется без браузера.
    Иначе открывается страница согласия Google.
    """
    credentials = _read_saved_user_token(token_path)
    if credentials is not None and credentials.refresh_token and not _token_is_usable(credentials):
        try:
            credentials.refresh(Request())
        except RefreshError:
            credentials = None
        else:
            if credentials is not None and _token_is_usable(credentials):
                _save_user_token(token_path, credentials)
                return credentials
    if credentials is not None and _token_is_usable(credentials):
        return credentials
    credentials = _authorize_in_browser(client_secrets_path)
    _save_user_token(token_path, credentials)
    return credentials


def _current_user(drive: Any) -> tuple[str, str]:
    """Почта и имя пользователя, от чьего имени выдан токен."""
    payload = _call_with_retry(
        lambda: drive.about().get(fields="user(emailAddress,displayName)")
    )
    user = payload.get("user")
    if not isinstance(user, dict):
        return "", ""
    email = user.get("emailAddress")
    name = user.get("displayName")
    return (
        email if isinstance(email, str) else "",
        name if isinstance(name, str) else "",
    )


def _http_error(error: HttpError) -> GoogleDriveError:
    status = int(getattr(error.resp, "status", 0) or 0)
    details = error.reason if getattr(error, "reason", None) else str(error)
    reason = str(details)
    if _STORAGE_QUOTA_MARKER in reason.lower():
        return GoogleDriveError(
            "STORAGE_QUOTA",
            (
                "У сервисного аккаунта нет своего места на Диске. "
                "Создавайте файл в папке, открытой этому аккаунту."
            ),
            {"status": status, "reason": reason},
        )
    code = _HTTP_ERROR_CODES.get(status, "GOOGLE_API_ERROR")
    if status == 404:
        message = "Файл или папка не найдены, либо к ним нет доступа."
    elif status == 403:
        message = "Нет доступа к файлу или папке на Диске."
    else:
        message = f"Google Drive API вернул ошибку {status}."
    return GoogleDriveError(code, message, {"status": status, "reason": reason})


def _call_with_retry(factory: RequestFactory) -> ApiPayload:
    """Выполнить запрос и повторить его при временных сбоях Google API."""
    delay_seconds = _RETRY_BACKOFF_SECONDS
    attempt = 0
    while True:
        try:
            payload = factory().execute()
        except HttpError as error:
            status = int(getattr(error.resp, "status", 0) or 0)
            can_retry = status in _TRANSIENT_HTTP_STATUSES and attempt < _MAX_RETRY_COUNT
            if not can_retry:
                raise _http_error(error) from error
            time.sleep(delay_seconds)
            delay_seconds *= 2
            attempt += 1
            continue
        if isinstance(payload, dict):
            return payload
        return {}


def _escape_query(value: str) -> str:
    """Экранировать значение для языка запросов Drive API."""
    return value.replace("\\", "\\\\").replace("'", "\\'")


def _parse_drive_file(item: ApiPayload) -> DriveFile | None:
    file_id = item.get("id")
    name = item.get("name")
    if not isinstance(file_id, str) or not isinstance(name, str):
        return None
    mime_type = item.get("mimeType")
    link = item.get("webViewLink")
    modified = item.get("modifiedTime")
    raw_parents = item.get("parents")
    parents: tuple[str, ...] = ()
    if isinstance(raw_parents, list):
        parents = tuple(parent for parent in raw_parents if isinstance(parent, str))
    return DriveFile(
        id=file_id,
        name=name,
        mime_type=mime_type if isinstance(mime_type, str) else "",
        web_view_link=link if isinstance(link, str) else "",
        modified_time=modified if isinstance(modified, str) else "",
        parents=parents,
        owner_email=_owner_email(item),
    )


def _owner_email(item: ApiPayload) -> str:
    """Достать почту первого владельца из ответа Drive API."""
    owners = item.get("owners")
    if not isinstance(owners, list) or not owners:
        return ""
    first = owners[0]
    if not isinstance(first, dict):
        return ""
    email = first.get("emailAddress")
    return email if isinstance(email, str) else ""


def _create_empty_file(
    drive: Any,
    name: str,
    mime_type: str,
    parent: str | None,
) -> DriveFile:
    """Создать пустой файл Google и вернуть его карточку."""
    cleaned_name = name.strip()
    if not cleaned_name:
        raise GoogleDriveError("VALIDATION_ERROR", "Не задано имя файла.", None)
    cleaned_mime = mime_type.strip()
    if not cleaned_mime:
        raise GoogleDriveError("VALIDATION_ERROR", "Не задан тип файла.", None)
    body: ApiPayload = {"name": cleaned_name, "mimeType": cleaned_mime}
    if parent:
        body["parents"] = [parent]
    payload = _call_with_retry(
        lambda: drive.files().create(
            body=body,
            fields=_FILE_FIELDS,
            supportsAllDrives=True,
        )
    )
    parsed = _parse_drive_file(payload)
    if parsed is None:
        raise GoogleDriveError(
            "GOOGLE_API_ERROR",
            "Google не вернул созданный файл.",
            payload,
        )
    return parsed


class GoogleDriveClient:
    """Создание, чтение и удаление файлов на Google Диске."""

    def __init__(
        self,
        credentials_path: str | Path | None = None,
        folder_id: str | None = None,
    ) -> None:
        self.credentials_path = resolve_credentials_path(credentials_path)
        self.account_email = service_account_email(self.credentials_path)
        cleaned_folder = folder_id.strip() if folder_id and folder_id.strip() else None
        self.folder_id = normalize_folder_id(cleaned_folder) if cleaned_folder else None
        credentials = _credentials(self.credentials_path)
        # Один клиент на экземпляр: и список, и запись идут через него.
        self._drive = build("drive", "v3", credentials=credentials)

    @classmethod
    def from_env(
        cls,
        env_path: str | Path | None = None,
        credentials_path: str | Path | None = None,
        folder_id: str | None = None,
    ) -> GoogleDriveClient:
        """Собрать клиент из .env и переменных окружения.

        Явные аргументы важнее файла. Так модуль одинаково работает
        локально и внутри другого приложения.
        """
        path = Path(env_path) if env_path is not None else _MODULE_DIR / ".env"
        _load_env_file(path)
        resolved_folder = folder_id if folder_id is not None else os.environ.get(_FOLDER_ID_ENV)
        return cls(credentials_path=credentials_path, folder_id=resolved_folder)

    def read_all(self, folder_id: str | None = None) -> list[DriveFile]:
        """Прочитать все файлы рабочей папки. Без папки — всё, что видит аккаунт."""
        parent = self._resolve_folder_id(folder_id)
        if parent:
            query = f"'{_escape_query(parent)}' in parents and trashed = false"
        else:
            query = "trashed = false"
        files = self._list_files(query)
        return sorted(files, key=lambda item: (item.mime_type != MIME_FOLDER, item.name.lower()))

    def read_file(self, file_id: str) -> DriveFile:
        """Прочитать карточку одного файла: имя, тип, ссылку и родителей."""
        cleaned_id = file_id.strip()
        if not cleaned_id:
            raise GoogleDriveError("VALIDATION_ERROR", "Не задан идентификатор файла.", None)
        payload = _call_with_retry(
            lambda: self._drive.files().get(
                fileId=cleaned_id,
                fields=_FILE_FIELDS,
                supportsAllDrives=True,
            )
        )
        parsed = _parse_drive_file(payload)
        if parsed is None:
            raise GoogleDriveError(
                "GOOGLE_API_ERROR",
                "Google не вернул карточку файла.",
                payload,
            )
        return parsed

    def create_file(
        self,
        name: str,
        mime_type: str = MIME_DOCUMENT,
        folder_id: str | None = None,
    ) -> DriveFile:
        """Создать пустой файл Google в рабочей папке и вернуть его карточку."""
        parent = self._resolve_folder_id(folder_id)
        return _create_empty_file(self._drive, name, mime_type, parent)

    def delete_file(self, file_id: str) -> None:
        """Удалить файл или папку по идентификатору без помещения в корзину."""
        cleaned_id = file_id.strip()
        if not cleaned_id:
            raise GoogleDriveError("VALIDATION_ERROR", "Не задан идентификатор файла.", None)
        _call_with_retry(
            lambda: self._drive.files().delete(
                fileId=cleaned_id,
                supportsAllDrives=True,
            )
        )

    def _resolve_folder_id(self, folder_id: str | None) -> str | None:
        if folder_id is None:
            return self.folder_id
        cleaned = folder_id.strip()
        if not cleaned:
            return None
        return normalize_folder_id(cleaned)

    def _list_files(self, query: str) -> list[DriveFile]:
        files: list[DriveFile] = []
        page_token: str | None = None
        for _ in range(_MAX_LIST_PAGES):
            token = page_token
            payload = _call_with_retry(
                lambda: self._drive.files().list(
                    q=query,
                    fields=_LIST_FIELDS,
                    pageSize=_PAGE_SIZE,
                    pageToken=token,
                    supportsAllDrives=True,
                    includeItemsFromAllDrives=True,
                )
            )
            raw_files = payload.get("files", [])
            if isinstance(raw_files, list):
                for item in raw_files:
                    if not isinstance(item, dict):
                        continue
                    parsed = _parse_drive_file(item)
                    if parsed is not None:
                        files.append(parsed)
            next_token = payload.get("nextPageToken")
            if not isinstance(next_token, str) or not next_token:
                return files
            page_token = next_token
        raise GoogleDriveError(
            "LIST_TOO_LARGE",
            "Файлов слишком много, список обрезан по лимиту страниц.",
            {"limit": _MAX_LIST_PAGES * _PAGE_SIZE},
        )


class GoogleDriveUserClient:
    """Создание файлов Google от имени пользователя через OAuth2.

    Первый вызов открывает браузер. Токен сохраняется и дальше обновляется сам.
    Файл принадлежит вошедшему пользователю и лежит в указанной папке.
    """

    def __init__(
        self,
        client_secrets_path: str | Path | None = None,
        token_path: str | Path | None = None,
        folder_id: str | None = None,
    ) -> None:
        self.client_secrets_path = resolve_oauth_client_path(client_secrets_path)
        self.token_path = _resolve_token_path(token_path)
        cleaned_folder = folder_id.strip() if folder_id and folder_id.strip() else None
        self.folder_id = normalize_folder_id(cleaned_folder) if cleaned_folder else None
        user_credentials = load_user_credentials(self.client_secrets_path, self.token_path)
        self._drive = build("drive", "v3", credentials=user_credentials)
        self.user_email, self.user_name = _current_user(self._drive)

    @classmethod
    def from_env(
        cls,
        env_path: str | Path | None = None,
        client_secrets_path: str | Path | None = None,
        token_path: str | Path | None = None,
        folder_id: str | None = None,
    ) -> GoogleDriveUserClient:
        """Собрать клиент пользователя из .env и переменных окружения."""
        path = Path(env_path) if env_path is not None else _MODULE_DIR / ".env"
        _load_env_file(path)
        resolved_folder = folder_id if folder_id is not None else os.environ.get(_FOLDER_ID_ENV)
        return cls(
            client_secrets_path=client_secrets_path,
            token_path=token_path,
            folder_id=resolved_folder,
        )

    def create_document(self, name: str, folder_id: str | None = None) -> DriveFile:
        """Создать пустой Google Документ в указанной папке от имени пользователя."""
        return self._create_in_folder(name, MIME_DOCUMENT, folder_id)

    def create_spreadsheet(self, name: str, folder_id: str | None = None) -> DriveFile:
        """Создать пустую Google Таблицу в указанной папке от имени пользователя."""
        return self._create_in_folder(name, MIME_SPREADSHEET, folder_id)

    def grant_editor(self, file_id: str, email: str) -> None:
        """Открыть файл почте как редактору, без письма-уведомления.

        Таблицу создаёт личный аккаунт. Запись через GoogleSheetsClient идёт
        от сервисного аккаунта, поэтому без этого доступа он получает 404.
        """
        cleaned_id = file_id.strip()
        cleaned_email = email.strip()
        if not cleaned_id:
            raise GoogleDriveError("VALIDATION_ERROR", "Не задан идентификатор файла.", None)
        if not cleaned_email:
            raise GoogleDriveError("VALIDATION_ERROR", "Не задана почта редактора.", None)
        _call_with_retry(
            lambda: self._drive.permissions().create(
                fileId=cleaned_id,
                body={
                    "type": "user",
                    "role": "writer",
                    "emailAddress": cleaned_email,
                },
                sendNotificationEmail=False,
                supportsAllDrives=True,
            )
        )

    def _create_in_folder(
        self,
        name: str,
        mime_type: str,
        folder_id: str | None,
    ) -> DriveFile:
        parent = self._require_folder_id(folder_id)
        return _create_empty_file(self._drive, name, mime_type, parent)

    def _require_folder_id(self, folder_id: str | None) -> str:
        parent = self._resolve_folder_id(folder_id)
        if parent:
            return parent
        raise GoogleDriveError(
            "FOLDER_NOT_CONFIGURED",
            "Укажите папку в аргументе folder_id или в GOOGLE_DRIVE_FOLDER_ID.",
            None,
        )

    def _resolve_folder_id(self, folder_id: str | None) -> str | None:
        if folder_id is None:
            return self.folder_id
        cleaned = folder_id.strip()
        if not cleaned:
            return None
        return normalize_folder_id(cleaned)


def _print_files(files: list[DriveFile]) -> None:
    if not files:
        print("(пусто)")
        return
    for item in files:
        parents = ", ".join(item.parents) if item.parents else "—"
        print(f"{item.kind_label}\t{item.name}\t{item.id}\tродители: {parents}")
        if item.web_view_link:
            print(f"  {item.web_view_link}")


def main() -> None:
    """Прочитать все файлы рабочей папки и вывести их в консоль."""
    try:
        client = GoogleDriveClient.from_env()
        files = client.read_all()
    except GoogleDriveError as error:
        print(f"[{error.code}] {error.message}")
        if error.details:
            print(error.details)
        raise SystemExit(1) from error

    print(f"Сервисный аккаунт: {client.account_email}")
    if client.folder_id:
        print(f"Рабочая папка: {client.folder_id}")
    else:
        print("GOOGLE_DRIVE_FOLDER_ID не задан. Показываю все файлы, которые видит аккаунт.")
    print(f"Найдено: {len(files)}")
    _print_files(files)


if __name__ == "__main__":
    main()
