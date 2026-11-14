"""Выгрузка видимой таблицы CRM в Google Таблицу.

Окно само не ходит в API Google. Создание файла делает GoogleDriveUserClient,
запись с оформлением — GoogleSheetsClient.publish_document. Оба модуля лежат
в папке reports и подключаются здесь, в GUI, чтобы вход OAuth шёл с компьютера
пользователя.

Пути к JSON и идентификатор папки пишутся в google_settings.txt рядом с окном.
"""

from __future__ import annotations

import sys
import time
import tkinter as tk
from collections import Counter
from collections.abc import Callable
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

_UI_DIR = Path(__file__).resolve().parent
_REPORTS_DIR = _UI_DIR.parent / "reports"
if str(_REPORTS_DIR) not in sys.path:
    sys.path.insert(0, str(_REPORTS_DIR))

from google_drive import (
    GoogleDriveError,
    GoogleDriveUserClient,
    normalize_folder_id,
    service_account_email,
)
from google_sheets import (
    GoogleSheetsClient,
    GoogleSheetsError,
    RangeStyle,
    Row,
    SheetPresentation,
)

SETTINGS_PATH = _UI_DIR / "google_settings.txt"
TOKEN_PATH = _UI_DIR / "google_oauth_token.json"

_SECTION_TITLES = {
    "clients": "Клиенты",
    "deals": "Сделки",
    "orders": "Заказы",
    "tasks": "Задачи",
}
_GROUP_HEADERS = {
    "clients": "Статус",
    "deals": "Статус",
    "orders": "Статус",
    "tasks": "Вид",
}
_MONEY_SECTIONS = frozenset({"deals", "orders"})
_NAVY = "1B3A4B"
_PAPER = "FFFFFF"
_INK = "1F2933"
_MUTED = "5C6B73"
_ZEBRA = "F4F7F8"
_LINE = "D5DEE7"
_EMPTY_LABEL = "Не указано"
_SETTINGS_KEYS = ("oauth_client", "service_account", "folder_id")


class ExportError(Exception):
    """Ошибка сборки отчёта до обращения к Google."""

    def __init__(self, message: str) -> None:
        self.message = message
        super().__init__(message)


@dataclass(frozen=True)
class GoogleSettings:
    """Пути, которые пользователь выбирает один раз."""

    oauth_client: str
    service_account: str
    folder_id: str

    def missing(self) -> tuple[str, ...]:
        """Подписи пустых или пропавших полей."""
        gaps: list[str] = []
        if not Path(self.oauth_client.strip()).is_file():
            gaps.append("JSON OAuth-клиента")
        if not Path(self.service_account.strip()).is_file():
            gaps.append("JSON сервисного аккаунта")
        if self.folder_id.strip() == "":
            gaps.append("папка Google Диска")
        return tuple(gaps)


@dataclass(frozen=True)
class TableView:
    """Видимые ячейки таблицы и суммы сделок или заказов."""

    section: str
    headers: tuple[str, ...]
    rows: tuple[tuple[str, ...], ...]
    amounts: tuple[tuple[str, str], ...] = ()


@dataclass(frozen=True)
class PublishedReport:
    """Созданный файл и лист, куда легли данные."""

    file_name: str
    file_id: str
    sheet_title: str
    url: str


def load_google_settings(path: Path = SETTINGS_PATH) -> GoogleSettings:
    """Прочитать текстовый файл настроек. Нет файла — пустые поля."""
    found = {key: "" for key in _SETTINGS_KEYS}
    if not path.is_file():
        return GoogleSettings("", "", "")
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        cleaned = key.strip()
        if cleaned in found:
            found[cleaned] = value.strip()
    return GoogleSettings(
        oauth_client=found["oauth_client"],
        service_account=found["service_account"],
        folder_id=found["folder_id"],
    )


def save_google_settings(settings: GoogleSettings, path: Path = SETTINGS_PATH) -> None:
    """Записать пути в текстовый файл, чтобы следующий запуск их подставил."""
    lines = (
        "# Настройки Google для окна мини-CRM.",
        f"oauth_client={settings.oauth_client.strip()}",
        f"service_account={settings.service_account.strip()}",
        f"folder_id={settings.folder_id.strip()}",
        "",
    )
    path.write_text("\n".join(lines), encoding="utf-8")


def build_document(
    view: TableView,
    generated_at: datetime,
) -> tuple[str, str, list[Row], SheetPresentation]:
    """Собрать имя файла, строки и макет. Сеть здесь не вызывается."""
    if not view.headers:
        raise ExportError("У таблицы нет колонок, отчёт собрать не из чего.")
    title = _SECTION_TITLES.get(view.section, view.section)
    file_name = f"CRM {title} {generated_at.strftime('%Y-%m-%d %H-%M')}"
    width = len(view.headers)
    analysis = _analysis_lines(view)
    rows: list[Row] = []
    merges: list[tuple[int, int, int, int]] = []
    analysis_rows: list[int] = []

    def add(values: list[str]) -> int:
        rows.append(values)
        return len(rows)

    title_row = add([f"Отчёт CRM. {title}"])
    subtitle_row = add([f"Сформировано {generated_at.strftime('%Y-%m-%d %H:%M')}"])
    add([""])
    for line in analysis:
        analysis_rows.append(add([line]))
    add([""])
    header_row = add(list(view.headers))
    for data_row in view.rows:
        add(list(data_row))

    if width > 1:
        merges.append((title_row, 1, title_row, width))
        merges.append((subtitle_row, 1, subtitle_row, width))
        for index in analysis_rows:
            merges.append((index, 1, index, width))

    styles = _styles(header_row, len(rows), width, view.headers, analysis_rows)
    heights = {title_row: 32.0, subtitle_row: 20.0}
    for index in analysis_rows:
        heights[index] = 24.0
    presentation = SheetPresentation(
        column_widths=_column_widths(view.headers),
        styles=tuple(styles),
        merges=tuple(merges),
        row_heights=heights,
        default_row_height=22,
        tab_color=_NAVY,
        hide_gridlines=True,
    )
    return file_name, title, rows, presentation


def folder_id_for_drive(value: str) -> str:
    """Достать id папки из ссылки. Ссылку на таблицу папкой не считать."""
    cleaned = value.strip()
    if cleaned == "":
        raise ExportError("Укажите папку Google Диска.")
    if "/spreadsheets/" in cleaned.casefold():
        raise ExportError(
            "В поле папки сохранена ссылка на таблицу. "
            "Нужна ссылка на папку: в адресе должно быть /folders/."
        )
    try:
        folder_id = normalize_folder_id(cleaned)
    except GoogleDriveError as error:
        raise ExportError(error.message) from error
    if "/" in folder_id or ":" in folder_id or " " in folder_id:
        raise ExportError(
            "Не удалось прочитать папку. Вставьте ссылку "
            "https://drive.google.com/drive/folders/… или сам идентификатор."
        )
    return folder_id


def _retry_not_found(action: Callable[[], None], attempts: int = 6) -> None:
    """Повторить шаг, пока Google не увидит только что созданный файл."""
    last_error: GoogleDriveError | None = None
    for attempt in range(attempts):
        try:
            action()
            return
        except GoogleDriveError as error:
            last_error = error
            if error.code != "NOT_FOUND" or attempt == attempts - 1:
                raise
            time.sleep(2)
    if last_error is not None:
        raise last_error


def _open_sheets(
    spreadsheet_id: str,
    credentials_path: str,
    sheet_name: str,
) -> GoogleSheetsClient:
    """Открыть таблицу сервисным аккаунтом. Доступ после выдачи прав приходит не сразу."""
    last_error: GoogleSheetsError | None = None
    for attempt in range(6):
        try:
            return GoogleSheetsClient(
                spreadsheet_id=spreadsheet_id,
                credentials_path=credentials_path,
                sheet_name=sheet_name,
            )
        except GoogleSheetsError as error:
            last_error = error
            if error.code != "NOT_FOUND" or attempt == 5:
                raise
            time.sleep(2)
    if last_error is None:
        raise ExportError("Не удалось открыть созданную таблицу.")
    raise last_error


def publish_report(settings: GoogleSettings, view: TableView) -> PublishedReport:
    """Создать таблицу на Диске и записать в неё собранный отчёт."""
    gaps = settings.missing()
    if gaps:
        raise ExportError("Сначала заполните настройки Google: " + ", ".join(gaps) + ".")
    folder_id = folder_id_for_drive(settings.folder_id)
    file_name, sheet_name, rows, presentation = build_document(view, datetime.now())
    drive = GoogleDriveUserClient(
        client_secrets_path=settings.oauth_client,
        token_path=TOKEN_PATH,
        folder_id=folder_id,
    )
    created = drive.create_spreadsheet(file_name)
    editor = service_account_email(settings.service_account)
    _retry_not_found(lambda: drive.grant_editor(created.id, editor))
    sheets = _open_sheets(created.id, settings.service_account, sheet_name)
    published = sheets.publish_document(sheet_name, rows, presentation)
    return PublishedReport(
        file_name=created.name or file_name,
        file_id=created.id,
        sheet_title=published.title,
        url=published.url,
    )


def _analysis_lines(view: TableView) -> tuple[str, ...]:
    """Короткий разбор видимых строк: сколько, какой статус чаще, какая сумма."""
    count = len(view.rows)
    if count == 0:
        return ("Записей в таблице нет.",)
    lines = [f"Записей в таблице: {count}."]
    group_header = _GROUP_HEADERS.get(view.section)
    if group_header is not None:
        counted = _count_column(view.headers, view.rows, group_header)
        if counted:
            parts = [f"{label} — {number}" for label, number in counted]
            lines.append(f"По колонке «{group_header}»: " + ", ".join(parts) + ".")
            leader, leader_count = counted[0]
            lines.append(f"Больше всего «{leader}»: {leader_count} из {count}.")
    if view.section == "tasks":
        done = _count_column(view.headers, view.rows, "Готово")
        if done:
            parts = [f"{label} — {number}" for label, number in done]
            lines.append("Выполнение: " + ", ".join(parts) + ".")
    if view.section in _MONEY_SECTIONS:
        money = _sum_amounts(view.amounts)
        if money:
            lines.append("Сумма видимых строк: " + money + ".")
    return tuple(lines)


def _count_column(
    headers: tuple[str, ...],
    rows: tuple[tuple[str, ...], ...],
    header: str,
) -> tuple[tuple[str, int], ...]:
    """Сколько раз встречается каждое значение колонки, сначала самые частые."""
    if header not in headers:
        return ()
    index = headers.index(header)
    counter: Counter[str] = Counter()
    for row in rows:
        label = row[index].strip() if index < len(row) else ""
        if label == "":
            label = _EMPTY_LABEL
        counter[label] += 1
    ordered = sorted(counter.items(), key=lambda item: (-item[1], item[0]))
    return tuple(ordered)


def _sum_amounts(amounts: tuple[tuple[str, str], ...]) -> str:
    """Сложить суммы по валютам. Пустые и нечисловые ячейки пропускаются."""
    totals: dict[str, Decimal] = {}
    for raw_amount, raw_currency in amounts:
        cleaned = raw_amount.strip().replace(" ", "").replace(",", ".")
        if cleaned == "":
            continue
        try:
            value = Decimal(cleaned)
        except InvalidOperation:
            continue
        currency = raw_currency.strip() or "RUB"
        totals[currency] = totals.get(currency, Decimal("0")) + value
    if not totals:
        return ""
    parts = [f"{value:.2f} {code}" for code, value in sorted(totals.items())]
    return ", ".join(parts)


def _column_widths(headers: tuple[str, ...]) -> tuple[float, ...]:
    """Ширина колонок под заголовки окна, а не под длину каждой ячейки."""
    widths: list[float] = []
    for header in headers:
        if header == "ID":
            widths.append(8)
        elif header in {"Готово", "Дата", "Срок"}:
            widths.append(16)
        elif header in {"Название", "Почта", "Клиент"}:
            widths.append(28)
        else:
            widths.append(18)
    return tuple(widths)


def _styles(
    header_row: int,
    row_count: int,
    width: int,
    headers: tuple[str, ...],
    analysis_rows: list[int],
) -> list[RangeStyle]:
    """Шапка документа, разбор и таблица. Поздний стиль перекрывает только свои поля."""
    styles = [
        RangeStyle(
            start_row=1,
            start_col=1,
            end_row=1,
            end_col=width,
            font_name="Calibri",
            font_size=18,
            font_color=_NAVY,
            bold=True,
            vertical="center",
        ),
        RangeStyle(
            start_row=2,
            start_col=1,
            end_row=2,
            end_col=width,
            font_name="Calibri",
            font_size=10,
            font_color=_MUTED,
            vertical="center",
        ),
    ]
    for index in analysis_rows:
        styles.append(
            RangeStyle(
                start_row=index,
                start_col=1,
                end_row=index,
                end_col=width,
                font_name="Calibri",
                font_size=11,
                font_color=_INK,
                wrap=True,
                vertical="center",
            )
        )
    styles.append(
        RangeStyle(
            start_row=header_row,
            start_col=1,
            end_row=row_count,
            end_col=width,
            font_name="Calibri",
            font_size=10,
            border="all",
            border_color=_LINE,
            wrap=True,
            vertical="center",
        )
    )
    styles.append(
        RangeStyle(
            start_row=header_row,
            start_col=1,
            end_row=header_row,
            end_col=width,
            bold=True,
            font_color=_PAPER,
            fill_color=_NAVY,
            horizontal="center",
        )
    )
    data_row = header_row + 1
    while data_row <= row_count:
        if (data_row - header_row) % 2 == 0:
            styles.append(
                RangeStyle(
                    start_row=data_row,
                    start_col=1,
                    end_row=data_row,
                    end_col=width,
                    fill_color=_ZEBRA,
                )
            )
        data_row += 1
    if "Сумма" in headers and row_count >= header_row:
        money_column = headers.index("Сумма") + 1
        styles.append(
            RangeStyle(
                start_row=header_row,
                start_col=money_column,
                end_row=row_count,
                end_col=money_column,
                horizontal="right",
            )
        )
    return styles


class GoogleSettingsDialog(tk.Toplevel):
    """Окно путей к JSON и папки. Сохранение пишет текстовый файл."""

    def __init__(self, master: tk.Misc, current: GoogleSettings) -> None:
        super().__init__(master)
        self.result: GoogleSettings | None = None
        self.title("Google")
        self.resizable(False, False)
        self.transient(master)
        self.configure(background="#F4F7F8")
        self._oauth = tk.StringVar(value=current.oauth_client)
        self._service = tk.StringVar(value=current.service_account)
        self._folder = tk.StringVar(value=current.folder_id)
        self._build()
        self.grab_set()
        self.lift()
        self.focus_force()
        self.attributes("-topmost", True)
        self.after(250, lambda: self.attributes("-topmost", False))
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.bind("<Escape>", lambda _event: self._cancel())

    def _build(self) -> None:
        """Три поля: два JSON и папка с кнопкой вставки."""
        frame = ttk.Frame(self, padding=16)
        frame.pack(fill="both", expand=True)
        ttk.Label(
            frame,
            text="Файлы запоминаются в google_settings.txt и подставятся в следующий раз.",
            wraplength=520,
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 12))
        self._file_row(
            frame,
            1,
            "OAuth-клиент",
            self._oauth,
            "Нужен, чтобы создать таблицу от вашего имени.",
        )
        self._file_row(
            frame,
            4,
            "Сервисный аккаунт",
            self._service,
            "Нужен, чтобы записать строки и оформление в созданную таблицу.",
        )
        ttk.Label(frame, text="Папка на Диске").grid(row=7, column=0, sticky="w", pady=(8, 0))
        folder = ttk.Entry(frame, textvariable=self._folder, width=48)
        folder.grid(row=8, column=0, columnspan=2, sticky="we", pady=(4, 0))
        ttk.Button(frame, text="Вставить", command=self._paste_folder).grid(
            row=8,
            column=2,
            sticky="w",
            padx=(8, 0),
            pady=(4, 0),
        )
        folder.bind("<Control-v>", self._paste_event)
        folder.bind("<Control-V>", self._paste_event)
        folder.bind("<Shift-Insert>", self._paste_event)
        ttk.Label(
            frame,
            text="Ссылка на папку Диска, в адресе должно быть /folders/. Кнопка берёт текст из буфера.",
            wraplength=520,
        ).grid(row=9, column=0, columnspan=3, sticky="w", pady=(4, 0))
        buttons = ttk.Frame(frame)
        buttons.grid(row=10, column=0, columnspan=3, sticky="e", pady=(16, 0))
        ttk.Button(buttons, text="Отмена", command=self._cancel).pack(side="right")
        ttk.Button(buttons, text="Сохранить", command=self._save).pack(side="right", padx=(0, 8))
        frame.columnconfigure(0, weight=1)

    def _file_row(
        self,
        frame: ttk.Frame,
        row: int,
        caption: str,
        variable: tk.StringVar,
        hint: str,
    ) -> None:
        """Подпись, путь и кнопка выбора JSON."""
        ttk.Label(frame, text=caption).grid(row=row, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(frame, textvariable=variable, width=48, state="readonly").grid(
            row=row + 1,
            column=0,
            columnspan=2,
            sticky="we",
            pady=(4, 0),
        )
        ttk.Button(
            frame,
            text="Выбрать",
            command=lambda variable=variable: self._choose_json(variable),
        ).grid(row=row + 1, column=2, sticky="w", padx=(8, 0), pady=(4, 0))
        ttk.Label(frame, text=hint, wraplength=520).grid(
            row=row + 2,
            column=0,
            columnspan=3,
            sticky="w",
            pady=(4, 0),
        )

    def _choose_json(self, variable: tk.StringVar) -> None:
        """Диалог JSON. Выбранный путь сразу попадает в поле."""
        initial = str(_REPORTS_DIR if _REPORTS_DIR.is_dir() else _UI_DIR)
        selected = filedialog.askopenfilename(
            parent=self,
            title="JSON для Google",
            initialdir=initial,
            filetypes=(("JSON", "*.json"), ("Все файлы", "*.*")),
        )
        if selected:
            variable.set(selected)

    def _paste_event(self, _event: object) -> str:
        """Ctrl-V и Shift-Insert вставляют идентификатор папки целиком."""
        self._paste_folder()
        return "break"

    def _paste_folder(self) -> None:
        """Кнопка «Вставить»: в поле папки кладётся текст из буфера обмена."""
        try:
            text = self.clipboard_get()
        except tk.TclError:
            messagebox.showinfo("Google", "В буфере нет текста.", parent=self)
            return
        cleaned = text.strip()
        if cleaned == "":
            messagebox.showinfo("Google", "В буфере нет текста.", parent=self)
            return
        try:
            folder_id = folder_id_for_drive(cleaned)
        except ExportError as error:
            messagebox.showerror("Google", error.message, parent=self)
            return
        self._folder.set(folder_id)

    def _save(self) -> None:
        """Проверить поля и записать текстовый файл."""
        try:
            folder_id = folder_id_for_drive(self._folder.get())
        except ExportError as error:
            messagebox.showerror("Google", error.message, parent=self)
            return
        self._folder.set(folder_id)
        settings = GoogleSettings(
            oauth_client=self._oauth.get(),
            service_account=self._service.get(),
            folder_id=folder_id,
        )
        gaps = settings.missing()
        if gaps:
            messagebox.showerror(
                "Google",
                "Заполните: " + ", ".join(gaps) + ".",
                parent=self,
            )
            return
        try:
            save_google_settings(settings)
        except OSError as error:
            messagebox.showerror("Google", f"Не удалось записать настройки: {error}", parent=self)
            return
        self.result = settings
        self.destroy()

    def _cancel(self) -> None:
        """Закрыть окно, не меняя файл."""
        self.result = None
        self.destroy()


def ask_google_settings(master: tk.Misc, current: GoogleSettings) -> GoogleSettings | None:
    """Открыть окно настроек и дождаться сохранения или отмены."""
    dialog = GoogleSettingsDialog(master, current)
    master.wait_window(dialog)
    return dialog.result


# Ошибки готовых клиентов. Окно ловит их и не импортирует модули Google само.
DriveFailure = GoogleDriveError
SheetsFailure = GoogleSheetsError
