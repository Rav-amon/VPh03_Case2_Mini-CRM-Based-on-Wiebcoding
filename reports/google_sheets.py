"""Клиент CRUD для Google Таблицы через официальный Google Sheets API.

Модуль подключается в любое приложение:

    from google_sheets import GoogleSheetsClient

    table = GoogleSheetsClient(
        spreadsheet_id="идентификатор_или_ссылка",
        credentials_path="путь/к/ключу.json",
    )
    rows = table.read_all()
"""

from __future__ import annotations

import io
import json
import os
import time
import zipfile
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Callable

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from openpyxl import load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.worksheet import Worksheet

# Таблица — для нативной Google Таблицы. Диск — чтобы читать и менять загруженный Excel.
_SHEETS_SCOPE = "https://www.googleapis.com/auth/spreadsheets"
_DRIVE_SCOPE = "https://www.googleapis.com/auth/drive"
_DRIVE_READONLY_SCOPE = "https://www.googleapis.com/auth/drive.readonly"
_GOOGLE_SHEET_MIME = "application/vnd.google-apps.spreadsheet"
_XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
_OFFICE_FILE_MARKER = "must not be an Office file"

_SPREADSHEET_ID_ENV = "GOOGLE_SPREADSHEET_ID"
_SHEET_NAME_ENV = "GOOGLE_SHEET_NAME"
_CREDENTIALS_PATH_ENVS = (
    "GOOGLE_CREDENTIALS_PATH",
    "GOOGLE_APPLICATION_CREDENTIALS",
)
_URL_ID_MARKER = "/spreadsheets/d/"

_VALUE_INPUT_OPTION = "USER_ENTERED"
_INSERT_DATA_OPTION = "INSERT_ROWS"
_ROW_DIMENSION = "ROWS"

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

CellValue = str | int | float | bool
Row = list[CellValue]
ApiPayload = dict[str, Any]
RequestFactory = Callable[[], Any]

# Символы, которые нельзя класть в имя листа ни в Excel, ни в Google Таблицах.
_FORBIDDEN_SHEET_TITLE_CHARS = frozenset("\\/?*[]:")
_DEFAULT_BORDER_COLOR = "D0D7DE"
_EXCEL_TITLE_LIMIT = 31
_GOOGLE_TITLE_LIMIT = 100

_SHEETS_HORIZONTAL = {"left": "LEFT", "center": "CENTER", "right": "RIGHT"}
_SHEETS_VERTICAL = {"top": "TOP", "center": "MIDDLE", "bottom": "BOTTOM"}
_OPENPYXL_VERTICAL = {"top": "top", "center": "center", "bottom": "bottom"}

_MODULE_DIR = Path(__file__).resolve().parent


class GoogleSheetsError(Exception):
    """Ошибка работы с таблицей в едином формате code / message / details."""

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


def normalize_spreadsheet_id(value: str) -> str:
    """Достать идентификатор из ссылки на таблицу или вернуть уже готовый id."""
    cleaned = value.strip()
    if _URL_ID_MARKER not in cleaned:
        return cleaned
    tail = cleaned.split(_URL_ID_MARKER, 1)[1]
    spreadsheet_id = tail.split("/", 1)[0].split("?", 1)[0].strip()
    if not spreadsheet_id:
        raise GoogleSheetsError(
            "VALIDATION_ERROR",
            "В ссылке на таблицу нет идентификатора.",
            {"value": value},
        )
    return spreadsheet_id


def quote_sheet_title(sheet_name: str) -> str:
    """Экранировать имя листа для A1-нотации."""
    escaped = sheet_name.replace("'", "''")
    return f"'{escaped}'"


def build_a1_range(sheet_name: str, cell_range: str | None = None) -> str:
    """Собрать диапазон. Без cell_range возвращается весь лист."""
    quoted = quote_sheet_title(sheet_name)
    if cell_range is None or cell_range.strip() == "":
        return quoted
    if "!" in cell_range:
        return cell_range
    return f"{quoted}!{cell_range}"


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
        raise GoogleSheetsError(
            "CREDENTIALS_NOT_FOUND",
            "Файл ключа сервисного аккаунта не найден.",
            {"path": str(path)},
        ) from error
    except json.JSONDecodeError as error:
        raise GoogleSheetsError(
            "CREDENTIALS_INVALID",
            "JSON-ключ сервисного аккаунта повреждён.",
            {"path": str(path)},
        ) from error
    if not isinstance(payload, dict):
        raise GoogleSheetsError(
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
        except GoogleSheetsError:
            continue
        if payload.get("type") == "service_account":
            matches.append(path)

    if len(matches) == 1:
        return matches[0]
    if not matches:
        raise GoogleSheetsError(
            "CREDENTIALS_NOT_FOUND",
            "Не найден JSON-ключ сервисного аккаунта. Укажите credentials_path.",
            None,
        )
    raise GoogleSheetsError(
        "AMBIGUOUS_CREDENTIALS",
        "В папке несколько ключей сервисного аккаунта. Укажите credentials_path явно.",
        {"files": [path.name for path in matches]},
    )


def service_account_email(credentials_path: str | Path | None = None) -> str:
    """Вернуть email сервисного аккаунта, которому нужно выдать доступ к таблице."""
    path = resolve_credentials_path(credentials_path)
    email = _read_json(path).get("client_email")
    if not isinstance(email, str) or not email:
        raise GoogleSheetsError(
            "CREDENTIALS_INVALID",
            "В JSON-ключе нет client_email.",
            {"path": str(path)},
        )
    return email


def _credentials(path: Path, scopes: list[str]) -> Credentials:
    try:
        return Credentials.from_service_account_file(str(path), scopes=scopes)
    except (ValueError, OSError) as error:
        raise GoogleSheetsError(
            "CREDENTIALS_INVALID",
            "Не удалось прочитать ключ сервисного аккаунта.",
            {"path": str(path)},
        ) from error


def _http_error(error: HttpError) -> GoogleSheetsError:
    status = int(getattr(error.resp, "status", 0) or 0)
    details = error.reason if getattr(error, "reason", None) else str(error)
    if _OFFICE_FILE_MARKER in str(details):
        return GoogleSheetsError(
            "OFFICE_FILE",
            "Это файл Excel на Диске. Sheets API открывает только настоящую Google Таблицу.",
            {"status": status, "reason": details},
        )
    code = _HTTP_ERROR_CODES.get(status, "GOOGLE_API_ERROR")
    return GoogleSheetsError(
        code,
        f"Google API вернул ошибку {status}.",
        {"status": status, "reason": details},
    )


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


def _validate_rows(rows: list[Row]) -> None:
    if not rows:
        raise GoogleSheetsError("VALIDATION_ERROR", "Нужно передать хотя бы одну строку.", None)
    for index, row in enumerate(rows, start=1):
        if not isinstance(row, list):
            raise GoogleSheetsError(
                "VALIDATION_ERROR",
                "Каждая строка должна быть списком ячеек.",
                {"row_number": index},
            )


def list_accessible_spreadsheets(
    credentials_path: str | Path | None = None,
) -> list[dict[str, str]]:
    """Список таблиц, к которым у сервисного аккаунта уже есть доступ."""
    path = resolve_credentials_path(credentials_path)
    credentials = _credentials(path, [_DRIVE_READONLY_SCOPE])
    drive = build("drive", "v3", credentials=credentials)
    query = (
        "(mimeType='application/vnd.google-apps.spreadsheet' or "
        "mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') "
        "and trashed=false"
    )
    payload = _call_with_retry(
        lambda: drive.files().list(
            q=query,
            fields="files(id, name)",
            pageSize=100,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        )
    )
    files = payload.get("files", [])
    spreadsheets: list[dict[str, str]] = []
    if not isinstance(files, list):
        return spreadsheets
    for item in files:
        if not isinstance(item, dict):
            continue
        file_id = item.get("id")
        file_name = item.get("name")
        if isinstance(file_id, str) and isinstance(file_name, str):
            spreadsheets.append({"id": file_id, "name": file_name})
    return spreadsheets


def _unquote_sheet(value: str) -> str:
    """Снять одинарные кавычки A1-нотации с имени листа."""
    text = value.strip()
    if len(text) >= 2 and text.startswith("'") and text.endswith("'"):
        return text[1:-1].replace("''", "'")
    return text


def _is_cell_range(value: str) -> bool:
    """Проверить, что строка похожа на диапазон ячеек, а не на имя листа."""
    try:
        range_boundaries(value)
    except ValueError:
        return False
    return True


def _split_range(cell_range: str) -> tuple[str | None, str | None]:
    """Разделить A1-адрес на имя листа и диапазон ячеек."""
    text = cell_range.strip()
    if "!" in text:
        sheet_part, cells = text.split("!", 1)
        cleaned_cells = cells.strip()
        return _unquote_sheet(sheet_part), cleaned_cells or None
    if _is_cell_range(text):
        return None, text
    return _unquote_sheet(text), None


def _cell_to_value(value: Any) -> CellValue:
    """Привести значение ячейки Excel к простому типу."""
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float, str)):
        return value
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(value, date):
        return value.isoformat()
    if value is None:
        return ""
    return str(value)


def _trim_row_tail(row: Row) -> Row:
    end = len(row)
    while end > 0 and row[end - 1] == "":
        end -= 1
    return row[:end]


def _trim_empty_edges(rows: list[Row]) -> list[Row]:
    trimmed = [_trim_row_tail(row) for row in rows]
    while trimmed and not trimmed[-1]:
        trimmed.pop()
    return trimmed


def _sheet_rows(
    sheet: Worksheet,
    bounds: tuple[int, int, int, int] | None = None,
) -> list[Row]:
    """Считать прямоугольник ячеек. Без bounds обрезает пустые края листа."""
    if bounds is None:
        min_col, min_row = 1, 1
        max_col = sheet.max_column or 1
        max_row = sheet.max_row or 1
        trim = True
    else:
        min_col, min_row, max_col, max_row = bounds
        trim = False
    if max_row < min_row or max_col < min_col:
        return []
    rows: list[Row] = []
    for excel_row in sheet.iter_rows(
        min_row=min_row,
        max_row=max_row,
        min_col=min_col,
        max_col=max_col,
        values_only=True,
    ):
        rows.append([_cell_to_value(cell) for cell in excel_row])
    if trim:
        return _trim_empty_edges(rows)
    return rows


def _cell_bounds(sheet: Worksheet, cells: str) -> tuple[int, int, int, int]:
    try:
        min_col, min_row, max_col, max_row = range_boundaries(cells)
    except ValueError as error:
        raise GoogleSheetsError(
            "VALIDATION_ERROR",
            "Некорректный диапазон ячеек.",
            {"range": cells},
        ) from error
    return (
        min_col or 1,
        min_row or 1,
        max_col or sheet.max_column or 1,
        max_row or sheet.max_row or 1,
    )


@dataclass(frozen=True)
class RangeStyle:
    """Оформление прямоугольника ячеек. None значит «это свойство не трогать»."""

    start_row: int
    start_col: int
    end_row: int
    end_col: int
    font_name: str | None = None
    font_size: int | None = None
    font_color: str | None = None
    bold: bool | None = None
    italic: bool | None = None
    fill_color: str | None = None
    horizontal: str | None = None
    vertical: str | None = None
    wrap: bool | None = None
    number_format: str | None = None
    border: str | None = None
    border_color: str | None = None


@dataclass(frozen=True)
class SheetPresentation:
    """Макет листа, который должен читаться как страница документа."""

    column_widths: tuple[float, ...]
    styles: tuple[RangeStyle, ...]
    merges: tuple[tuple[int, int, int, int], ...] = ()
    row_heights: dict[int, float] | None = None
    default_row_height: float = 22
    tab_color: str = "1B3A4B"
    hide_gridlines: bool = True


@dataclass(frozen=True)
class PublishedSheet:
    """Созданный лист и ссылка, по которой его открыть."""

    title: str
    url: str


def sanitize_sheet_title(title: str, max_length: int) -> str:
    """Убрать из имени листа символы, из-за которых Excel и Google отклоняют лист."""
    cleaned = "".join(
        char
        for char in title
        if char not in _FORBIDDEN_SHEET_TITLE_CHARS and ord(char) >= 32
    )
    cleaned = cleaned.strip().strip("'")
    if not cleaned:
        cleaned = "Отчёт"
    return cleaned[:max_length]


def unique_sheet_title(title: str, existing: list[str], max_length: int) -> str:
    """Подобрать имя, которого ещё нет среди листов. Сравнение без учёта регистра."""
    base = sanitize_sheet_title(title, max_length)
    taken = {name.casefold() for name in existing}
    if base.casefold() not in taken:
        return base
    for index in range(2, 100):
        suffix = f" ({index})"
        candidate = f"{base[: max_length - len(suffix)]}{suffix}"
        if candidate.casefold() not in taken:
            return candidate
    raise GoogleSheetsError(
        "VALIDATION_ERROR",
        "Не удалось подобрать уникальное имя листа.",
        {"title": title},
    )


def _sheets_color(hex_color: str) -> dict[str, float]:
    """Перевести RRGGBB в цвет Google Sheets API."""
    cleaned = hex_color.strip().lstrip("#")
    if len(cleaned) != 6:
        raise GoogleSheetsError(
            "VALIDATION_ERROR",
            "Цвет должен быть в виде шестизначного HEX.",
            {"color": hex_color},
        )
    try:
        red = int(cleaned[0:2], 16)
        green = int(cleaned[2:4], 16)
        blue = int(cleaned[4:6], 16)
    except ValueError as error:
        raise GoogleSheetsError(
            "VALIDATION_ERROR",
            "Цвет должен быть в виде шестизначного HEX.",
            {"color": hex_color},
        ) from error
    return {"red": red / 255, "green": green / 255, "blue": blue / 255}


def _column_width_to_pixels(width: float) -> int:
    """Ширина в символах Excel примерно переводится в пиксели Google Таблицы."""
    return max(int(width * 7.5 + 5), 24)


def _points_to_pixels(points: float) -> int:
    """Высота строки в пунктах для Excel и в пикселях для Google Таблицы."""
    return max(int(round(points * 4 / 3)), 8)


def _validate_span(
    start_row: int,
    start_col: int,
    end_row: int,
    end_col: int,
    row_count: int,
    col_count: int,
) -> None:
    if (
        start_row < 1
        or start_col < 1
        or end_row < start_row
        or end_col < start_col
        or end_row > row_count
        or end_col > col_count
    ):
        raise GoogleSheetsError(
            "VALIDATION_ERROR",
            "Диапазон оформления выходит за пределы листа.",
            {
                "start_row": start_row,
                "start_col": start_col,
                "end_row": end_row,
                "end_col": end_col,
            },
        )


def _validate_presentation(presentation: SheetPresentation, row_count: int) -> None:
    """Проверить, что макет попадает в записанный прямоугольник ячеек."""
    col_count = len(presentation.column_widths)
    if col_count < 1:
        raise GoogleSheetsError("VALIDATION_ERROR", "Не заданы ширины колонок.", None)
    if presentation.default_row_height <= 0:
        raise GoogleSheetsError(
            "VALIDATION_ERROR",
            "Высота строки должна быть больше нуля.",
            None,
        )
    for style in presentation.styles:
        _validate_span(
            style.start_row,
            style.start_col,
            style.end_row,
            style.end_col,
            row_count,
            col_count,
        )
        if style.horizontal is not None and style.horizontal not in _SHEETS_HORIZONTAL:
            raise GoogleSheetsError(
                "VALIDATION_ERROR",
                "Неизвестное выравнивание по горизонтали.",
                {"horizontal": style.horizontal},
            )
        if style.vertical is not None and style.vertical not in _SHEETS_VERTICAL:
            raise GoogleSheetsError(
                "VALIDATION_ERROR",
                "Неизвестное выравнивание по вертикали.",
                {"vertical": style.vertical},
            )
        if style.border is not None and style.border not in {"all", "bottom", "top"}:
            raise GoogleSheetsError(
                "VALIDATION_ERROR",
                "Неизвестный вид границы.",
                {"border": style.border},
            )
    for start_row, start_col, end_row, end_col in presentation.merges:
        _validate_span(start_row, start_col, end_row, end_col, row_count, col_count)
    heights = presentation.row_heights or {}
    for row_index, height in heights.items():
        if row_index < 1 or row_index > row_count or height <= 0:
            raise GoogleSheetsError(
                "VALIDATION_ERROR",
                "Некорректная высота строки.",
                {"row": row_index, "height": height},
            )


def _pad_rows(rows: list[Row], width: int) -> list[Row]:
    """Дополнить строки пустыми ячейками, чтобы прямоугольник был ровным."""
    padded: list[Row] = []
    for row in rows:
        if len(row) >= width:
            padded.append(list(row[:width]))
            continue
        padded.append([*row, *[""] * (width - len(row))])
    return padded


def _repeat_cell_request(sheet_id: int, style: RangeStyle) -> dict[str, Any] | None:
    """Один repeatCell только с теми полями, которые заданы в стиле."""
    user_format: dict[str, Any] = {}
    fields: list[str] = []
    text_format: dict[str, Any] = {}
    if style.bold is not None:
        text_format["bold"] = style.bold
        fields.append("userEnteredFormat.textFormat.bold")
    if style.italic is not None:
        text_format["italic"] = style.italic
        fields.append("userEnteredFormat.textFormat.italic")
    if style.font_size is not None:
        text_format["fontSize"] = style.font_size
        fields.append("userEnteredFormat.textFormat.fontSize")
    if style.font_name is not None:
        text_format["fontFamily"] = style.font_name
        fields.append("userEnteredFormat.textFormat.fontFamily")
    if style.font_color is not None:
        text_format["foregroundColor"] = _sheets_color(style.font_color)
        fields.append("userEnteredFormat.textFormat.foregroundColor")
    if text_format:
        user_format["textFormat"] = text_format
    if style.fill_color is not None:
        user_format["backgroundColor"] = _sheets_color(style.fill_color)
        fields.append("userEnteredFormat.backgroundColor")
    if style.horizontal is not None:
        user_format["horizontalAlignment"] = _SHEETS_HORIZONTAL[style.horizontal]
        fields.append("userEnteredFormat.horizontalAlignment")
    if style.vertical is not None:
        user_format["verticalAlignment"] = _SHEETS_VERTICAL[style.vertical]
        fields.append("userEnteredFormat.verticalAlignment")
    if style.wrap is not None:
        user_format["wrapStrategy"] = "WRAP" if style.wrap else "CLIP"
        fields.append("userEnteredFormat.wrapStrategy")
    if style.number_format is not None:
        kind = "PERCENT" if "%" in style.number_format else "NUMBER"
        user_format["numberFormat"] = {"type": kind, "pattern": style.number_format}
        fields.append("userEnteredFormat.numberFormat")
    if style.border is not None:
        color = style.border_color or _DEFAULT_BORDER_COLOR
        borders, border_fields = _sheets_borders(style.border, color)
        user_format["borders"] = borders
        fields.extend(border_fields)
    if not fields:
        return None
    return {
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": style.start_row - 1,
                "endRowIndex": style.end_row,
                "startColumnIndex": style.start_col - 1,
                "endColumnIndex": style.end_col,
            },
            "cell": {"userEnteredFormat": user_format},
            "fields": ",".join(fields),
        }
    }


def _sheets_borders(mode: str, color: str) -> tuple[dict[str, Any], list[str]]:
    """Границы и маска полей. Частичная маска не стирает остальные рёбра."""
    def side(kind: str) -> dict[str, Any]:
        sheets_style = "SOLID_MEDIUM" if kind == "medium" else "SOLID"
        return {"style": sheets_style, "color": _sheets_color(color)}

    borders: dict[str, Any] = {}
    fields: list[str] = []
    if mode == "all":
        drawn = side("thin")
        for edge in ("top", "bottom", "left", "right"):
            borders[edge] = drawn
            fields.append(f"userEnteredFormat.borders.{edge}")
        return borders, fields
    edge = "bottom" if mode == "bottom" else "top"
    borders[edge] = side("medium")
    fields.append(f"userEnteredFormat.borders.{edge}")
    return borders, fields


def _google_presentation_requests(
    sheet_id: int,
    row_count: int,
    presentation: SheetPresentation,
) -> list[dict[str, Any]]:
    """Собрать пакет: стили, объединения, ширины и печатную страницу A4."""
    requests: list[dict[str, Any]] = []
    for style in presentation.styles:
        request = _repeat_cell_request(sheet_id, style)
        if request is not None:
            requests.append(request)
    for start_row, start_col, end_row, end_col in presentation.merges:
        requests.append(
            {
                "mergeCells": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": start_row - 1,
                        "endRowIndex": end_row,
                        "startColumnIndex": start_col - 1,
                        "endColumnIndex": end_col,
                    },
                    "mergeType": "MERGE_ALL",
                }
            }
        )
    requests.append(
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": 0,
                    "endIndex": row_count,
                },
                "properties": {"pixelSize": _points_to_pixels(presentation.default_row_height)},
                "fields": "pixelSize",
            }
        }
    )
    for row_index, height in (presentation.row_heights or {}).items():
        requests.append(
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "ROWS",
                        "startIndex": row_index - 1,
                        "endIndex": row_index,
                    },
                    "properties": {"pixelSize": _points_to_pixels(height)},
                    "fields": "pixelSize",
                }
            }
        )
    for index, width in enumerate(presentation.column_widths):
        requests.append(
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": index,
                        "endIndex": index + 1,
                    },
                    "properties": {"pixelSize": _column_width_to_pixels(width)},
                    "fields": "pixelSize",
                }
            }
        )
    # Печатные поля A4 остаются в ветке Excel. У Sheets API этого поля в свойствах листа нет.
    return requests


def _added_sheet_id(payload: ApiPayload) -> int:
    """Достать id листа из ответа addSheet."""
    replies = payload.get("replies")
    if not isinstance(replies, list) or not replies or not isinstance(replies[0], dict):
        raise GoogleSheetsError("GOOGLE_API_ERROR", "Google не вернул новый лист.", payload)
    added = replies[0].get("addSheet")
    properties = added.get("properties") if isinstance(added, dict) else None
    sheet_id = properties.get("sheetId") if isinstance(properties, dict) else None
    if not isinstance(sheet_id, int):
        raise GoogleSheetsError("GOOGLE_API_ERROR", "Google не вернул номер нового листа.", payload)
    return sheet_id


def _font_rgb(font: Font) -> str:
    """Сохранить уже стоящий цвет шрифта, если стиль его не меняет."""
    try:
        rgb = font.color.rgb if font.color is not None else None
    except (AttributeError, TypeError):
        return "1F2933"
    if not isinstance(rgb, str) or len(rgb) < 6:
        return "1F2933"
    return rgb[-6:]


def _apply_openpyxl_cell(cell: Cell, style: RangeStyle) -> None:
    """Наложить только заданные поля стиля, не затирая остальные."""
    if any(
        value is not None
        for value in (
            style.font_name,
            style.font_size,
            style.font_color,
            style.bold,
            style.italic,
        )
    ):
        current = cell.font
        cell.font = Font(
            name=style.font_name or current.name or "Calibri",
            size=style.font_size or current.size or 11,
            bold=current.bold if style.bold is None else style.bold,
            italic=current.italic if style.italic is None else style.italic,
            color=style.font_color or _font_rgb(current),
        )
    if style.fill_color is not None:
        cell.fill = PatternFill(fill_type="solid", fgColor=style.fill_color)
    if style.horizontal is not None or style.vertical is not None or style.wrap is not None:
        current_alignment = cell.alignment
        vertical = style.vertical or current_alignment.vertical or "center"
        if vertical not in _OPENPYXL_VERTICAL:
            vertical = "center"
        wrap = current_alignment.wrap_text if style.wrap is None else style.wrap
        cell.alignment = Alignment(
            horizontal=style.horizontal or current_alignment.horizontal or "left",
            vertical=_OPENPYXL_VERTICAL[vertical],
            wrap_text=bool(wrap),
        )
    if style.number_format is not None:
        cell.number_format = style.number_format
    if style.border is not None:
        color = style.border_color or _DEFAULT_BORDER_COLOR
        cell.border = _openpyxl_border(cell.border, style.border, color)


def _openpyxl_border(current: Border, mode: str, color: str) -> Border:
    """Поставить рамку. Верх или низ дописываются к уже существующей рамке."""
    thin = Side(style="thin", color=color)
    medium = Side(style="medium", color=color)
    if mode == "all":
        return Border(left=thin, right=thin, top=thin, bottom=thin)
    if mode == "bottom":
        return Border(left=current.left, right=current.right, top=current.top, bottom=medium)
    return Border(left=current.left, right=current.right, top=medium, bottom=current.bottom)


def _write_openpyxl_sheet(
    sheet: Worksheet,
    rows: list[Row],
    presentation: SheetPresentation,
) -> None:
    """Записать значения и сверстать лист Excel как печатную страницу."""
    for row_index, row in enumerate(rows, start=1):
        for col_index, value in enumerate(row, start=1):
            sheet.cell(row=row_index, column=col_index, value=None if value == "" else value)
    for style in presentation.styles:
        for excel_row in sheet.iter_rows(
            min_row=style.start_row,
            max_row=style.end_row,
            min_col=style.start_col,
            max_col=style.end_col,
        ):
            for cell in excel_row:
                if isinstance(cell, Cell):
                    _apply_openpyxl_cell(cell, style)
    for start_row, start_col, end_row, end_col in presentation.merges:
        sheet.merge_cells(
            start_row=start_row,
            start_column=start_col,
            end_row=end_row,
            end_column=end_col,
        )
    sheet.sheet_format.defaultRowHeight = presentation.default_row_height
    for row_index, height in (presentation.row_heights or {}).items():
        sheet.row_dimensions[row_index].height = height
    for index, width in enumerate(presentation.column_widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    last_column = get_column_letter(len(presentation.column_widths))
    sheet.print_area = f"A1:{last_column}{len(rows)}"
    sheet.page_setup.orientation = "portrait"
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.page_setup.horizontalCentered = True
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_margins = PageMargins(
        left=0.6,
        right=0.6,
        top=0.7,
        bottom=0.6,
        header=0.3,
        footer=0.3,
    )
    sheet.print_options.horizontalCentered = True
    sheet.sheet_view.showGridLines = not presentation.hide_gridlines
    sheet.sheet_view.zoomScale = 110
    sheet.sheet_view.view = "pageLayout"
    sheet.sheet_properties.tabColor = presentation.tab_color
    sheet.oddFooter.left.text = "Смоделированный отчёт"
    sheet.oddFooter.left.font = "Calibri"
    sheet.oddFooter.left.size = 9
    sheet.oddFooter.center.text = "Страница &P из &N"
    sheet.oddFooter.center.font = "Calibri"
    sheet.oddFooter.center.size = 9


class _ExcelDriveBook:
    """CRUD для .xlsx на Google Диске: Sheets API такие файлы не принимает."""

    def __init__(self, drive: Any, file_id: str, mime_type: str) -> None:
        self._drive = drive
        self._file_id = file_id
        self._mime_type = mime_type

    def list_sheet_titles(self) -> list[str]:
        return list(self._download().sheetnames)

    def read_range(self, cell_range: str, fallback_sheet: str | None) -> list[Row]:
        sheet_name, cells = _split_range(cell_range)
        workbook = self._download()
        sheet = self._worksheet(workbook, self._sheet_name(workbook, sheet_name, fallback_sheet))
        if cells is None:
            return _sheet_rows(sheet)
        return _sheet_rows(sheet, _cell_bounds(sheet, cells))

    def append_rows(self, rows: list[Row], sheet_name: str) -> int:
        workbook = self._download()
        sheet = self._worksheet(workbook, sheet_name)
        for row in rows:
            sheet.append(list(row))
        self._upload(workbook)
        return len(rows)

    def update_range(self, cell_range: str, rows: list[Row], fallback_sheet: str | None) -> int:
        sheet_name, cells = _split_range(cell_range)
        if cells is None:
            raise GoogleSheetsError(
                "VALIDATION_ERROR",
                "Для обновления нужен диапазон ячеек, а не только имя листа.",
                {"range": cell_range},
            )
        workbook = self._download()
        sheet = self._worksheet(workbook, self._sheet_name(workbook, sheet_name, fallback_sheet))
        min_col, min_row, _, _ = _cell_bounds(sheet, cells)
        updated = 0
        for row_offset, row in enumerate(rows):
            for col_offset, value in enumerate(row):
                sheet.cell(
                    row=min_row + row_offset,
                    column=min_col + col_offset,
                    value=value,
                )
                updated += 1
        self._upload(workbook)
        return updated

    def clear_range(self, cell_range: str, fallback_sheet: str | None) -> str:
        sheet_name, cells = _split_range(cell_range)
        if cells is None:
            raise GoogleSheetsError(
                "VALIDATION_ERROR",
                "Для очистки нужен диапазон ячеек, а не только имя листа.",
                {"range": cell_range},
            )
        workbook = self._download()
        sheet = self._worksheet(workbook, self._sheet_name(workbook, sheet_name, fallback_sheet))
        min_col, min_row, max_col, max_row = _cell_bounds(sheet, cells)
        for row_index in range(min_row, max_row + 1):
            for column_index in range(min_col, max_col + 1):
                sheet.cell(row=row_index, column=column_index, value=None)
        self._upload(workbook)
        return cell_range

    def delete_rows(self, start_row: int, end_row: int, sheet_name: str) -> None:
        workbook = self._download()
        sheet = self._worksheet(workbook, sheet_name)
        deleted_count = end_row - start_row + 1
        sheet.delete_rows(start_row, deleted_count)
        self._upload(workbook)

    def publish_document(
        self,
        sheet_name: str,
        rows: list[Row],
        presentation: SheetPresentation,
    ) -> str:
        """Добавить оформленный лист в книгу Excel и залить файл обратно на Диск."""
        workbook = self._download()
        title = unique_sheet_title(sheet_name, workbook.sheetnames, _EXCEL_TITLE_LIMIT)
        sheet = workbook.create_sheet(title)
        _write_openpyxl_sheet(sheet, rows, presentation)
        position = workbook.sheetnames.index(sheet.title)
        if position:
            workbook.move_sheet(sheet, offset=-position)
        workbook.active = sheet
        workbook.calculation.calcMode = "auto"
        workbook.calculation.fullCalcOnLoad = True
        self._upload(workbook)
        return sheet.title

    def _sheet_name(
        self,
        workbook: Workbook,
        sheet_name: str | None,
        fallback_sheet: str | None,
    ) -> str:
        if sheet_name:
            return sheet_name
        if fallback_sheet:
            return fallback_sheet
        if not workbook.sheetnames:
            raise GoogleSheetsError("SHEET_NOT_FOUND", "В книге Excel нет листов.", None)
        return workbook.sheetnames[0]

    def _worksheet(self, workbook: Workbook, sheet_name: str) -> Worksheet:
        if sheet_name not in workbook.sheetnames:
            raise GoogleSheetsError(
                "SHEET_NOT_FOUND",
                f"Лист «{sheet_name}» не найден.",
                {"sheet_name": sheet_name},
            )
        sheet = workbook[sheet_name]
        if not isinstance(sheet, Worksheet):
            raise GoogleSheetsError(
                "SHEET_NOT_SUPPORTED",
                f"Лист «{sheet_name}» не является обычной таблицей.",
                {"sheet_name": sheet_name},
            )
        return sheet

    def _download(self) -> Workbook:
        buffer = io.BytesIO()
        request = self._drive.files().get_media(
            fileId=self._file_id,
            supportsAllDrives=True,
        )
        downloader = MediaIoBaseDownload(buffer, request)
        finished = False
        try:
            while not finished:
                _, finished = downloader.next_chunk()
        except HttpError as error:
            raise _http_error(error) from error
        buffer.seek(0)
        try:
            return load_workbook(buffer)
        except (InvalidFileException, zipfile.BadZipFile, OSError, ValueError) as error:
            raise GoogleSheetsError(
                "WORKBOOK_INVALID",
                "Не удалось прочитать файл Excel.",
                {"reason": str(error)},
            ) from error

    def _upload(self, workbook: Workbook) -> None:
        def factory() -> Any:
            buffer = io.BytesIO()
            workbook.save(buffer)
            buffer.seek(0)
            media = MediaIoBaseUpload(buffer, mimetype=self._mime_type, resumable=False)
            return self._drive.files().update(
                fileId=self._file_id,
                media_body=media,
                supportsAllDrives=True,
            )

        _call_with_retry(factory)


class GoogleSheetsClient:
    """CRUD для Google Таблицы или файла Excel .xlsx на Диске."""

    def __init__(
        self,
        spreadsheet_id: str,
        credentials_path: str | Path | None = None,
        sheet_name: str | None = None,
    ) -> None:
        cleaned_id = spreadsheet_id.strip()
        if not cleaned_id:
            raise GoogleSheetsError(
                "VALIDATION_ERROR",
                "Не задан идентификатор Google Таблицы.",
                None,
            )
        self.spreadsheet_id = normalize_spreadsheet_id(cleaned_id)
        self.credentials_path = resolve_credentials_path(credentials_path)
        self.default_sheet_name = sheet_name.strip() if sheet_name and sheet_name.strip() else None
        self._sheet_ids: dict[str, int] = {}
        credentials = _credentials(
            self.credentials_path,
            [_SHEETS_SCOPE, _DRIVE_SCOPE],
        )
        drive = build("drive", "v3", credentials=credentials)
        # Тип файла решает путь: нативная таблица идёт в Sheets API, .xlsx — в Диск.
        metadata = _call_with_retry(
            lambda: drive.files().get(
                fileId=self.spreadsheet_id,
                fields="id,name,mimeType",
                supportsAllDrives=True,
            )
        )
        mime_type = metadata.get("mimeType")
        self.file_name = metadata.get("name") if isinstance(metadata.get("name"), str) else ""
        self._sheets: Any = None
        self._book: _ExcelDriveBook | None = None
        if mime_type == _GOOGLE_SHEET_MIME:
            self._sheets = build("sheets", "v4", credentials=credentials).spreadsheets()
        elif mime_type == _XLSX_MIME:
            self._book = _ExcelDriveBook(drive, self.spreadsheet_id, mime_type)
        else:
            raise GoogleSheetsError(
                "UNSUPPORTED_DOCUMENT",
                "Файл не является Google Таблицей или Excel .xlsx.",
                {"name": self.file_name, "mimeType": mime_type},
            )

    @classmethod
    def from_env(
        cls,
        env_path: str | Path | None = None,
        spreadsheet_id: str | None = None,
        credentials_path: str | Path | None = None,
        sheet_name: str | None = None,
    ) -> GoogleSheetsClient:
        """Собрать клиент из .env и переменных окружения.

        Явные аргументы важнее файла. Так модуль одинаково работает
        локально и внутри другого приложения.
        """
        path = Path(env_path) if env_path is not None else _MODULE_DIR / ".env"
        _load_env_file(path)
        resolved_id = spreadsheet_id or os.environ.get(_SPREADSHEET_ID_ENV, "")
        resolved_sheet = sheet_name if sheet_name is not None else os.environ.get(_SHEET_NAME_ENV)
        return cls(
            spreadsheet_id=resolved_id,
            credentials_path=credentials_path,
            sheet_name=resolved_sheet,
        )

    def list_sheet_titles(self) -> list[str]:
        """Имена всех листов книги."""
        if self._book is not None:
            return self._book.list_sheet_titles()
        payload = self._call(lambda: self._sheets.get(spreadsheetId=self.spreadsheet_id))
        titles: list[str] = []
        sheets = payload.get("sheets", [])
        if not isinstance(sheets, list):
            return titles
        for sheet in sheets:
            if not isinstance(sheet, dict):
                continue
            properties = sheet.get("properties", {})
            if not isinstance(properties, dict):
                continue
            title = properties.get("title")
            sheet_id = properties.get("sheetId")
            if isinstance(title, str):
                titles.append(title)
                if isinstance(sheet_id, int):
                    self._sheet_ids[title] = sheet_id
        return titles

    def read_all(self, sheet_name: str | None = None) -> list[Row]:
        """Прочитать все заполненные ячейки листа."""
        target = self._resolve_sheet_name(sheet_name)
        return self.read_range(build_a1_range(target))

    def read_workbook(self) -> dict[str, list[Row]]:
        """Прочитать все листы. Ключ словаря — имя листа."""
        if self.default_sheet_name:
            return {self.default_sheet_name: self.read_all(self.default_sheet_name)}
        return {title: self.read_all(title) for title in self.list_sheet_titles()}

    def read_range(self, cell_range: str) -> list[Row]:
        """Прочитать произвольный диапазон в A1-нотации."""
        if not cell_range.strip():
            raise GoogleSheetsError("VALIDATION_ERROR", "Диапазон ячеек не задан.", None)
        if self._book is not None:
            return self._book.read_range(cell_range, self.default_sheet_name)
        payload = self._call(
            lambda: self._sheets.values().get(
                spreadsheetId=self.spreadsheet_id,
                range=cell_range,
            )
        )
        values = payload.get("values", [])
        if not isinstance(values, list):
            return []
        return [row for row in values if isinstance(row, list)]

    def append_rows(self, rows: list[Row], sheet_name: str | None = None) -> int:
        """Добавить строки в конец листа. Возвращает число добавленных строк."""
        _validate_rows(rows)
        target_sheet = self._resolve_sheet_name(sheet_name)
        if self._book is not None:
            return self._book.append_rows(rows, target_sheet)
        target = build_a1_range(target_sheet)
        payload = self._call(
            lambda: self._sheets.values().append(
                spreadsheetId=self.spreadsheet_id,
                range=target,
                valueInputOption=_VALUE_INPUT_OPTION,
                insertDataOption=_INSERT_DATA_OPTION,
                body={"values": rows},
            )
        )
        updates = payload.get("updates", {})
        if not isinstance(updates, dict):
            return 0
        updated_rows = updates.get("updatedRows", 0)
        return int(updated_rows) if isinstance(updated_rows, int) else 0

    def update_range(self, cell_range: str, rows: list[Row]) -> int:
        """Перезаписать диапазон. Возвращает число изменённых ячеек."""
        if not cell_range.strip():
            raise GoogleSheetsError("VALIDATION_ERROR", "Диапазон для обновления не задан.", None)
        _validate_rows(rows)
        if self._book is not None:
            return self._book.update_range(cell_range, rows, self.default_sheet_name)
        payload = self._call(
            lambda: self._sheets.values().update(
                spreadsheetId=self.spreadsheet_id,
                range=cell_range,
                valueInputOption=_VALUE_INPUT_OPTION,
                body={"values": rows},
            )
        )
        updated_cells = payload.get("updatedCells", 0)
        return int(updated_cells) if isinstance(updated_cells, int) else 0

    def clear_range(self, cell_range: str) -> str:
        """Очистить значения в диапазоне, не удаляя сами строки."""
        if not cell_range.strip():
            raise GoogleSheetsError("VALIDATION_ERROR", "Диапазон для очистки не задан.", None)
        if self._book is not None:
            return self._book.clear_range(cell_range, self.default_sheet_name)
        payload = self._call(
            lambda: self._sheets.values().clear(
                spreadsheetId=self.spreadsheet_id,
                range=cell_range,
                body={},
            )
        )
        cleared = payload.get("clearedRange", cell_range)
        return str(cleared)

    def delete_rows(self, start_row: int, end_row: int, sheet_name: str | None = None) -> None:
        """Удалить строки включительно. Нумерация с 1, как в интерфейсе таблицы."""
        if start_row < 1 or end_row < start_row:
            raise GoogleSheetsError(
                "VALIDATION_ERROR",
                "Некорректный диапазон строк. Нумерация начинается с 1.",
                {"start_row": start_row, "end_row": end_row},
            )
        target = self._resolve_sheet_name(sheet_name)
        if self._book is not None:
            self._book.delete_rows(start_row, end_row, target)
            return
        sheet_id = self._sheet_id(target)
        self._call(
            lambda: self._sheets.batchUpdate(
                spreadsheetId=self.spreadsheet_id,
                body={
                    "requests": [
                        {
                            "deleteDimension": {
                                "range": {
                                    "sheetId": sheet_id,
                                    "dimension": _ROW_DIMENSION,
                                    "startIndex": start_row - 1,
                                    "endIndex": end_row,
                                }
                            }
                        }
                    ]
                },
            )
        )

    def publish_document(
        self,
        sheet_name: str,
        rows: list[Row],
        presentation: SheetPresentation,
    ) -> PublishedSheet:
        """Создать новый лист и записать на него оформленный документ.

        Для Google Таблицы оформление уходит через Sheets API.
        Для .xlsx на Диске тот же макет собирается в openpyxl и файл заливается целиком.
        """
        if not sheet_name.strip():
            raise GoogleSheetsError("VALIDATION_ERROR", "Не задано имя листа.", None)
        _validate_rows(rows)
        width = len(presentation.column_widths)
        padded = _pad_rows(rows, width)
        _validate_presentation(presentation, len(padded))
        if self._book is not None:
            title = self._book.publish_document(sheet_name, padded, presentation)
            url = f"https://docs.google.com/spreadsheets/d/{self.spreadsheet_id}/edit"
            return PublishedSheet(title=title, url=url)
        blank = self._only_blank_sheet()
        if blank is not None:
            return self._fill_only_sheet(blank, sheet_name, padded, presentation)
        return self._publish_google_document(sheet_name, padded, presentation)

    def _publish_google_document(
        self,
        sheet_name: str,
        rows: list[Row],
        presentation: SheetPresentation,
    ) -> PublishedSheet:
        """Создать лист Google Таблицы, записать значения и применить макет."""
        title = unique_sheet_title(
            sheet_name,
            self.list_sheet_titles(),
            _GOOGLE_TITLE_LIMIT,
        )
        width = len(presentation.column_widths)
        created = self._call(
            lambda: self._sheets.batchUpdate(
                spreadsheetId=self.spreadsheet_id,
                body={
                    "requests": [
                        {
                            "addSheet": {
                                "properties": {
                                    "title": title,
                                    "index": 0,
                                    "tabColor": _sheets_color(presentation.tab_color),
                                    "gridProperties": {
                                        "rowCount": len(rows) + 5,
                                        "columnCount": width,
                                        "hideGridlines": presentation.hide_gridlines,
                                    },
                                }
                            }
                        }
                    ]
                },
            )
        )
        sheet_id = _added_sheet_id(created)
        self._sheet_ids[title] = sheet_id
        self.update_range(build_a1_range(title, "A1"), rows)
        self._apply_presentation(sheet_id, len(rows), presentation)
        url = f"https://docs.google.com/spreadsheets/d/{self.spreadsheet_id}/edit#gid={sheet_id}"
        return PublishedSheet(title=title, url=url)

    def _only_blank_sheet(self) -> str | None:
        """Единственный лист новой таблицы. В него и нужно писать, а не создавать второй."""
        titles = self.list_sheet_titles()
        if len(titles) != 1:
            return None
        if self.read_all(titles[0]):
            return None
        return titles[0]

    def _fill_only_sheet(
        self,
        current_title: str,
        sheet_name: str,
        rows: list[Row],
        presentation: SheetPresentation,
    ) -> PublishedSheet:
        """Переименовать пустой лист и записать отчёт прямо в него."""
        sheet_id = self._sheet_id(current_title)
        title = unique_sheet_title(sheet_name, [], _GOOGLE_TITLE_LIMIT)
        width = len(presentation.column_widths)
        properties = {
            "sheetId": sheet_id,
            "title": title,
            "tabColor": _sheets_color(presentation.tab_color),
            "gridProperties": {
                "rowCount": len(rows) + 5,
                "columnCount": width,
                "hideGridlines": presentation.hide_gridlines,
            },
        }
        fields = (
            "title,tabColor,gridProperties.rowCount,"
            "gridProperties.columnCount,gridProperties.hideGridlines"
        )
        self._call(
            lambda: self._sheets.batchUpdate(
                spreadsheetId=self.spreadsheet_id,
                body={
                    "requests": [
                        {
                            "updateSheetProperties": {
                                "properties": properties,
                                "fields": fields,
                            }
                        }
                    ]
                },
            )
        )
        self._sheet_ids.pop(current_title, None)
        self._sheet_ids[title] = sheet_id
        self.update_range(build_a1_range(title, "A1"), rows)
        self._apply_presentation(sheet_id, len(rows), presentation)
        url = f"https://docs.google.com/spreadsheets/d/{self.spreadsheet_id}/edit#gid={sheet_id}"
        return PublishedSheet(title=title, url=url)

    def _apply_presentation(
        self,
        sheet_id: int,
        row_count: int,
        presentation: SheetPresentation,
    ) -> None:
        """Наложить макет. Значения уже лежат в ячейках и остаются, даже если стиль не принят."""
        requests = _google_presentation_requests(sheet_id, row_count, presentation)
        if not requests:
            return
        self._call(
            lambda: self._sheets.batchUpdate(
                spreadsheetId=self.spreadsheet_id,
                body={"requests": requests},
            )
        )

    def _resolve_sheet_name(self, sheet_name: str | None) -> str:
        if sheet_name and sheet_name.strip():
            return sheet_name.strip()
        if self.default_sheet_name:
            return self.default_sheet_name
        titles = self.list_sheet_titles()
        if not titles:
            raise GoogleSheetsError(
                "SHEET_NOT_FOUND",
                "В таблице нет ни одного листа.",
                {"spreadsheet_id": self.spreadsheet_id},
            )
        return titles[0]

    def _sheet_id(self, sheet_name: str) -> int:
        cached = self._sheet_ids.get(sheet_name)
        if cached is not None:
            return cached
        self.list_sheet_titles()
        cached = self._sheet_ids.get(sheet_name)
        if cached is None:
            raise GoogleSheetsError(
                "SHEET_NOT_FOUND",
                f"Лист «{sheet_name}» не найден.",
                {"sheet_name": sheet_name},
            )
        return cached

    def _call(self, factory: RequestFactory) -> ApiPayload:
        return _call_with_retry(factory)


def _client_for_check() -> GoogleSheetsClient:
    """Клиент для локальной проверки: id из .env или единственная доступная таблица."""
    _load_env_file(_MODULE_DIR / ".env")
    spreadsheet_id = os.environ.get(_SPREADSHEET_ID_ENV, "").strip()
    if spreadsheet_id:
        return GoogleSheetsClient.from_env()

    try:
        spreadsheets = list_accessible_spreadsheets()
    except GoogleSheetsError as error:
        email = service_account_email()
        raise GoogleSheetsError(
            error.code,
            (
                "Не задан GOOGLE_SPREADSHEET_ID, и список таблиц получить не удалось. "
                f"Откройте таблицу для {email} и впишите её id в .env."
            ),
            error.details,
        ) from error

    if len(spreadsheets) == 1:
        only = spreadsheets[0]
        print(f"GOOGLE_SPREADSHEET_ID не задан. Читаю таблицу «{only['name']}».")
        return GoogleSheetsClient.from_env(spreadsheet_id=only["id"])

    email = service_account_email()
    if not spreadsheets:
        raise GoogleSheetsError(
            "SPREADSHEET_NOT_CONFIGURED",
            (
                "Сервисный аккаунт не видит ни одной таблицы. "
                f"Выдайте доступ «Редактор» для {email} и укажите GOOGLE_SPREADSHEET_ID в .env."
            ),
            None,
        )

    raise GoogleSheetsError(
        "AMBIGUOUS_SPREADSHEET",
        "Доступно несколько таблиц. Укажите нужную в GOOGLE_SPREADSHEET_ID.",
        {"spreadsheets": spreadsheets},
    )


def _print_rows(rows: list[Row]) -> None:
    if not rows:
        print("(пусто)")
        return
    for row in rows:
        printable = ["" if cell is None else str(cell) for cell in row]
        print("\t".join(printable))


def main() -> None:
    """Прочитать все ячейки и вывести их в консоль."""
    try:
        client = _client_for_check()
        workbook = client.read_workbook()
    except GoogleSheetsError as error:
        print(f"[{error.code}] {error.message}")
        if error.details:
            print(error.details)
        raise SystemExit(1) from error

    if not workbook:
        print("В таблице нет листов.")
        return

    for title, rows in workbook.items():
        print(f"=== {title} ===")
        _print_rows(rows)


if __name__ == "__main__":
    main()
