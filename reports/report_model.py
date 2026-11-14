"""Сборка смоделированного отчёта: цифры случайные, итоги в листе — формулы.

Функции не ходят в сеть. Одинаковые входные данные и один и тот же генератор
случайных чисел дают один и тот же лист.
"""

from __future__ import annotations

import math
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from random import Random

from google_sheets import CellValue, RangeStyle, Row, SheetPresentation

COLUMN_COUNT = 6
MAX_REPORT_DAYS = 62
MIN_AUTHOR_LENGTH = 2
MAX_AUTHOR_LENGTH = 40

DEPARTMENTS: tuple[str, ...] = (
    "Отдел продаж",
    "Логистика",
    "Производство",
    "Клиентский сервис",
)
REPORT_KINDS: tuple[str, ...] = (
    "Операционный",
    "Финансовый",
    "Контроль качества",
)

_DATE_FORMAT = "%d.%m.%Y"
_WEEKDAYS = ("пн", "вт", "ср", "чт", "пт", "сб", "вс")
_COLUMN_WIDTHS = (24.0, 14.0, 18.0, 16.0, 18.0, 38.0)

_NAVY = "1B3A4B"
_INK = "1F2933"
_MUTED = "5C6B73"
_LINE = "D5DEE7"
_GOLD = "C4A35A"
_ZEBRA = "F4F7F8"
_HEADER = "E7EEF2"
_PAPER = "FFFFFF"
_BAND = "1B3A4B"
_BAND_SOFT = "E8EEF2"
_NOTE = "FBF8F3"
_NOTE_LINE = "E6D9C8"
_TOTAL = "E7EEF2"

_MONEY_FORMAT = '#,##0.00" ₽"'
_INT_FORMAT = "#,##0"
_PERCENT_FORMAT = "0.0%"

_REPORT_TITLES = {
    "Операционный": "ОПЕРАЦИОННЫЙ ОТЧЁТ",
    "Финансовый": "ФИНАНСОВЫЙ ОТЧЁТ",
    "Контроль качества": "ОТЧЁТ ПО КАЧЕСТВУ",
}
_ORDER_BASE = {
    "Отдел продаж": 46,
    "Логистика": 28,
    "Производство": 18,
    "Клиентский сервис": 34,
}
_CHECK_BOUNDS = {
    "Отдел продаж": (2500.0, 8700.0),
    "Логистика": (1400.0, 4600.0),
    "Производство": (9000.0, 26000.0),
    "Клиентский сервис": (800.0, 3400.0),
}
_COST_RATIO = {
    "Операционный": (0.62, 0.78),
    "Финансовый": (0.55, 0.84),
    "Контроль качества": (0.60, 0.75),
}
_RETURN_RATE = {
    "Операционный": (0.01, 0.06),
    "Финансовый": (0.01, 0.05),
    "Контроль качества": (0.02, 0.09),
}
_COMMENT_POOL = {
    "Операционный": (
        "Штатный ход смены",
        "Ровный поток заказов",
        "Без срывов по срокам",
        "Нагрузка ближе к вечеру",
        "Сверка с планом прошла",
    ),
    "Финансовый": (
        "Оплата прошла в срок",
        "Скидки в пределах лимита",
        "Закупка без отклонений",
        "Касса сверена",
        "Отклонений по цене нет",
    ),
    "Контроль качества": (
        "Повторная проверка партии",
        "Единичные замечания",
        "Рекламаций не было",
        "Контроль выборки пройден",
        "Замечания закрыты в смене",
    ),
}


@dataclass(frozen=True)
class ReportParams:
    """То, что человек вводит руками. Остальное моделируется."""

    start: date
    end: date
    department: str
    kind: str
    author: str


@dataclass(frozen=True)
class DayRecord:
    """Один смоделированный день."""

    day: date
    orders: int
    revenue: float
    returns_count: int
    cost: float
    comment: str


@dataclass(frozen=True)
class BuiltReport:
    """Готовый лист и цифры для строки состояния в окне."""

    sheet_name: str
    rows: list[Row]
    presentation: SheetPresentation
    day_count: int
    orders: int
    revenue: float
    margin: float
    return_rate: float


class _Grid:
    """Прямоугольник ячеек, стили и объединения. Строки нумеруются с 1."""

    def __init__(self) -> None:
        self.rows: list[Row] = []
        self.styles: list[RangeStyle] = []
        self.merges: list[tuple[int, int, int, int]] = []
        self.heights: dict[int, float] = {}

    def add(self, values: list[CellValue], height: float | None = None) -> int:
        """Добавить строку и вернуть её номер."""
        padded = list(values[:COLUMN_COUNT])
        if len(padded) < COLUMN_COUNT:
            padded.extend([""] * (COLUMN_COUNT - len(padded)))
        self.rows.append(padded)
        index = len(self.rows)
        if height is not None:
            self.heights[index] = height
        return index

    def put(self, row: int, col: int, value: CellValue) -> None:
        """Записать ячейку по номеру строки и колонки, оба с 1."""
        self.rows[row - 1][col - 1] = value

    def merge(self, row: int, start_col: int, end_col: int, end_row: int | None = None) -> None:
        """Объединить ячейки. end_row по умолчанию — та же строка."""
        self.merges.append((row, start_col, end_row if end_row is not None else row, end_col))

    def paint(self, style: RangeStyle) -> None:
        """Добавить слой оформления. Следующий слой перекрывает только свои поля."""
        self.styles.append(style)


def format_money(value: float) -> str:
    """Деньги в виде «1 234 567,89 ₽»."""
    formatted = f"{value:,.2f}".replace(",", " ").replace(".", ",")
    return f"{formatted} ₽"


def format_int(value: float) -> str:
    """Целое с пробелом между тысячами."""
    return f"{int(round(value)):,}".replace(",", " ")


def format_percent(ratio: float) -> str:
    """Доля 0.128 как «12,8%»."""
    formatted = f"{ratio * 100:,.1f}".replace(",", " ").replace(".", ",")
    return f"{formatted}%"


def parse_user_date(value: str) -> date:
    """Разобрать дату из поля формы. Ожидается ДД.ММ.ГГГГ."""
    try:
        return datetime.strptime(value.strip(), _DATE_FORMAT).date()
    except ValueError as error:
        raise ValueError("Дата должна быть в формате ДД.ММ.ГГГГ.") from error


def parse_report_params(
    start_text: str,
    end_text: str,
    department: str,
    kind: str,
    author: str,
) -> ReportParams:
    """Проверить поля формы и собрать параметры отчёта."""
    start = parse_user_date(start_text)
    end = parse_user_date(end_text)
    if end < start:
        raise ValueError("Дата окончания раньше даты начала.")
    day_count = (end - start).days + 1
    if day_count > MAX_REPORT_DAYS:
        raise ValueError(f"Период длиннее {MAX_REPORT_DAYS} дней.")
    if department not in DEPARTMENTS:
        raise ValueError("Выберите подразделение из списка.")
    if kind not in REPORT_KINDS:
        raise ValueError("Выберите вид отчёта из списка.")
    cleaned_author = " ".join(author.split())
    if len(cleaned_author) < MIN_AUTHOR_LENGTH:
        raise ValueError("Укажите ответственного.")
    if len(cleaned_author) > MAX_AUTHOR_LENGTH:
        raise ValueError("Имя ответственного слишком длинное.")
    return ReportParams(
        start=start,
        end=end,
        department=department,
        kind=kind,
        author=cleaned_author,
    )


def suggest_sheet_name(params: ReportParams, generated_at: datetime) -> str:
    """Короткое имя листа: у файла Excel предел 31 символ."""
    if params.start.year == params.end.year:
        period = f"{params.start:%d.%m}-{params.end:%d.%m}"
    else:
        period = f"{params.start:%d.%m.%y}-{params.end:%d.%m.%y}"
    return f"Отчёт {period} {generated_at:%H-%M}"


def build_report(params: ReportParams, rng: Random, generated_at: datetime) -> BuiltReport:
    """Собрать лист-документ. Сводка и итог ссылаются формулами на таблицу дней."""
    records = _simulate_days(params, rng)
    grid = _Grid()
    days_row = _append_header(grid, params, generated_at, rng)
    kpi_rows = _append_summary(grid)
    # До таблицы дней идут заголовок раздела, пустая строка и шапка колонок.
    rows_before_details = len(grid.rows)
    _append_details(grid, records)
    data_start = rows_before_details + 4
    data_end = data_start + len(records) - 1
    if grid.rows[data_start - 2][0] != "Дата":
        raise RuntimeError("Сместилась шапка таблицы дней.")
    total_row = _append_total(grid, data_start, data_end)
    _fill_summary_formulas(grid, kpi_rows, days_row, data_start, data_end)
    _append_conclusion(grid, params, records)
    _append_signature(grid, params, generated_at)
    _append_footer(grid)
    _paint_document(grid, kpi_rows, data_start, data_end, total_row)
    orders = sum(record.orders for record in records)
    revenue = sum(record.revenue for record in records)
    cost = sum(record.cost for record in records)
    returns_count = sum(record.returns_count for record in records)
    margin = 0.0 if revenue == 0 else (revenue - cost) / revenue
    return_rate = 0.0 if orders == 0 else returns_count / orders
    presentation = SheetPresentation(
        column_widths=_COLUMN_WIDTHS,
        styles=tuple(grid.styles),
        merges=tuple(grid.merges),
        row_heights=dict(grid.heights),
        default_row_height=22,
        tab_color=_NAVY,
        hide_gridlines=True,
    )
    return BuiltReport(
        sheet_name=suggest_sheet_name(params, generated_at),
        rows=grid.rows,
        presentation=presentation,
        day_count=len(records),
        orders=orders,
        revenue=revenue,
        margin=margin,
        return_rate=return_rate,
    )


def _append_header(
    grid: _Grid,
    params: ReportParams,
    generated_at: datetime,
    rng: Random,
) -> int:
    """Шапка как у титула документа: полоса, период и реквизиты.

    Возвращает строку, где потом появится формула числа дней.
    """
    title_row = grid.add([_REPORT_TITLES[params.kind]], 46)
    grid.merge(title_row, 1, COLUMN_COUNT)
    subtitle = f"за период с {params.start:%d.%m.%Y} по {params.end:%d.%m.%Y}"
    subtitle_row = grid.add([subtitle], 28)
    grid.merge(subtitle_row, 1, COLUMN_COUNT)
    grid.add([""], 12)
    grid.add(
        [
            "Подразделение",
            params.department,
            "Вид отчёта",
            params.kind,
            "Ответственный",
            params.author,
        ],
        24,
    )
    report_no = f"ВН-{generated_at:%y%m%d}-{rng.randint(100, 999)}"
    days_row = grid.add(
        [
            "Сформирован",
            generated_at.strftime("%d.%m.%Y %H:%M"),
            "Дней в выборке",
            "",
            "Номер",
            report_no,
        ],
        24,
    )
    grid.add([""], 14)
    return days_row


def _append_summary(grid: _Grid) -> dict[str, int]:
    """Блок показателей. Значения пока пустые: их заполнят формулы."""
    section = grid.add(["1. Сводные показатели"], 30)
    grid.merge(section, 1, COLUMN_COUNT)
    grid.add([""], 8)
    header = grid.add(["Показатель", "", "Значение", "", "Пояснение", ""], 22)
    grid.merge(header, 1, 2)
    grid.merge(header, 3, 4)
    grid.merge(header, 5, 6)
    rows = {
        "revenue": _summary_row(grid, "Выручка", "Сумма по дням периода"),
        "orders": _summary_row(grid, "Количество заказов", "Сумма заказов за период"),
        "returns": _summary_row(grid, "Возвраты, шт.", "Сумма возвратов за период"),
        "return_rate": _summary_row(grid, "Доля возвратов", "Возвраты к числу заказов"),
        "cost": _summary_row(grid, "Себестоимость", "Сумма затрат за период"),
        "margin": _summary_row(grid, "Маржинальность", "Выручка минус себестоимость"),
        "check": _summary_row(grid, "Средний чек", "Выручка на один заказ"),
    }
    grid.add([""], 14)
    return rows


def _summary_row(grid: _Grid, label: str, note: str) -> int:
    """Строка сводки из трёх объединённых пар колонок."""
    row = grid.add([label, "", "", "", note, ""], 24)
    grid.merge(row, 1, 2)
    grid.merge(row, 3, 4)
    grid.merge(row, 5, 6)
    return row


def _append_details(grid: _Grid, records: list[DayRecord]) -> None:
    """Таблица дней. Это единственное место, куда числа записываются как значения."""
    section = grid.add(["2. Детализация по дням"], 30)
    grid.merge(section, 1, COLUMN_COUNT)
    grid.add([""], 8)
    grid.add(
        ["Дата", "Заказы", "Выручка", "Возвраты", "Себестоимость", "Комментарий"],
        24,
    )
    for record in records:
        grid.add(
            [
                f"{record.day:%d.%m.%Y}, {_WEEKDAYS[record.day.weekday()]}",
                record.orders,
                record.revenue,
                record.returns_count,
                record.cost,
                record.comment,
            ]
        )


def _append_total(grid: _Grid, data_start: int, data_end: int) -> int:
    """Строка итога. Суммы считает сама таблица."""
    return grid.add(
        [
            "Итого",
            _sum_formula("B", data_start, data_end),
            _sum_formula("C", data_start, data_end),
            _sum_formula("D", data_start, data_end),
            _sum_formula("E", data_start, data_end),
            "За весь период",
        ],
        24,
    )


def _fill_summary_formulas(
    grid: _Grid,
    kpi_rows: dict[str, int],
    days_row: int,
    data_start: int,
    data_end: int,
) -> None:
    """Подставить формулы сводки, когда известны номера строк детализации."""
    revenue = f"SUM(C{data_start}:C{data_end})"
    orders = f"SUM(B{data_start}:B{data_end})"
    returns = f"SUM(D{data_start}:D{data_end})"
    cost = f"SUM(E{data_start}:E{data_end})"
    grid.put(kpi_rows["revenue"], 3, _sum_formula("C", data_start, data_end))
    grid.put(kpi_rows["orders"], 3, _sum_formula("B", data_start, data_end))
    grid.put(kpi_rows["returns"], 3, _sum_formula("D", data_start, data_end))
    grid.put(kpi_rows["return_rate"], 3, _ratio_formula(returns, orders))
    grid.put(kpi_rows["cost"], 3, _sum_formula("E", data_start, data_end))
    grid.put(kpi_rows["margin"], 3, _ratio_formula(f"{revenue}-{cost}", revenue))
    grid.put(kpi_rows["check"], 3, _ratio_formula(revenue, orders))
    grid.put(days_row, 4, f"=COUNTA(A{data_start}:A{data_end})")


def _append_conclusion(grid: _Grid, params: ReportParams, records: list[DayRecord]) -> None:
    """Текст вывода фиксируется при сборке. Числовые блоки выше остаются формулами."""
    section = grid.add(["3. Выводы"], 30)
    grid.merge(section, 1, COLUMN_COUNT)
    grid.add([""], 8)
    text = _conclusion_text(params, records)
    first = grid.add([text], 22)
    last = first
    for _ in range(3):
        last = grid.add([""], 22)
    grid.merge(first, 1, COLUMN_COUNT, last)
    grid.paint(_note_style(first, last))
    grid.add([""], 12)


def _append_signature(grid: _Grid, params: ReportParams, generated_at: datetime) -> None:
    """Блок подписи: фамилия, дата и пустая линия."""
    signed = grid.add(
        [
            "Ответственный",
            params.author,
            "",
            "Дата",
            generated_at.strftime("%d.%m.%Y"),
            "",
        ],
        24,
    )
    grid.merge(signed, 2, 3)
    grid.merge(signed, 5, 6)
    line = grid.add(["Подпись", "", "", "", "", ""], 26)
    grid.merge(line, 2, 3)
    grid.paint(
        RangeStyle(
            start_row=line,
            start_col=2,
            end_row=line,
            end_col=3,
            border="bottom",
            border_color=_NAVY,
        )
    )
    grid.add([""], 10)


def _append_footer(grid: _Grid) -> None:
    """Нижняя оговорка: откуда взялись цифры и что пересчитывается само."""
    row = grid.add(
        [
            "Документ собран автоматически. Показатели дней смоделированы и не являются "
            "данными учёта. Сводка и строка «Итого» пересчитываются формулами, "
            "текст вывода зафиксирован на момент формирования."
        ],
        42,
    )
    grid.merge(row, 1, COLUMN_COUNT)
    grid.paint(
        RangeStyle(
            start_row=row,
            start_col=1,
            end_row=row,
            end_col=COLUMN_COUNT,
            font_size=9,
            font_color=_MUTED,
            italic=True,
            vertical="top",
            wrap=True,
        )
    )


def _paint_document(
    grid: _Grid,
    kpi_rows: dict[str, int],
    data_start: int,
    data_end: int,
    total_row: int,
) -> None:
    """Наложить оформление страницы. Базовый слой идёт первым."""
    row_count = len(grid.rows)
    base = RangeStyle(
        start_row=1,
        start_col=1,
        end_row=row_count,
        end_col=COLUMN_COUNT,
        font_name="Calibri",
        font_size=11,
        font_color=_INK,
        bold=False,
        italic=False,
        fill_color=_PAPER,
        horizontal="left",
        vertical="center",
        wrap=False,
    )
    grid.styles.insert(0, base)
    grid.paint(
        RangeStyle(
            start_row=1,
            start_col=1,
            end_row=1,
            end_col=COLUMN_COUNT,
            font_name="Georgia",
            font_size=22,
            font_color="FFFFFF",
            bold=True,
            fill_color=_BAND,
            horizontal="center",
            vertical="center",
        )
    )
    grid.paint(
        RangeStyle(
            start_row=2,
            start_col=1,
            end_row=2,
            end_col=COLUMN_COUNT,
            font_name="Georgia",
            font_size=12,
            font_color=_BAND_SOFT,
            fill_color=_BAND,
            horizontal="center",
            border="bottom",
            border_color=_GOLD,
        )
    )
    meta_rows = (_find_label(grid, "Подразделение"), _find_label(grid, "Сформирован"))
    for row in meta_rows:
        for column in (1, 3, 5):
            grid.paint(
                RangeStyle(
                    start_row=row,
                    start_col=column,
                    end_row=row,
                    end_col=column,
                    font_size=9,
                    font_color=_MUTED,
                )
            )
        for column in (2, 4, 6):
            grid.paint(
                RangeStyle(
                    start_row=row,
                    start_col=column,
                    end_row=row,
                    end_col=column,
                    font_size=11,
                    font_color=_NAVY,
                    bold=True,
                )
            )
    days_row = _find_label(grid, "Сформирован")
    grid.paint(
        RangeStyle(
            start_row=days_row,
            start_col=4,
            end_row=days_row,
            end_col=4,
            number_format=_INT_FORMAT,
            horizontal="left",
            bold=True,
            font_color=_NAVY,
        )
    )
    for row in _section_rows(grid):
        grid.paint(_section_style(row))
    _paint_summary_table(grid, kpi_rows)
    header_row = data_start - 1
    grid.paint(
        RangeStyle(
            start_row=header_row,
            start_col=1,
            end_row=total_row,
            end_col=COLUMN_COUNT,
            border="all",
            border_color=_LINE,
        )
    )
    grid.paint(
        RangeStyle(
            start_row=header_row,
            start_col=1,
            end_row=header_row,
            end_col=COLUMN_COUNT,
            font_color="FFFFFF",
            bold=True,
            fill_color=_NAVY,
            horizontal="center",
        )
    )
    for offset in range(data_end - data_start + 1):
        if offset % 2 == 1:
            row = data_start + offset
            grid.paint(
                RangeStyle(
                    start_row=row,
                    start_col=1,
                    end_row=row,
                    end_col=COLUMN_COUNT,
                    fill_color=_ZEBRA,
                )
            )
    _paint_numeric_column(grid, data_start, total_row, 2, _INT_FORMAT)
    _paint_numeric_column(grid, data_start, total_row, 3, _MONEY_FORMAT)
    _paint_numeric_column(grid, data_start, total_row, 4, _INT_FORMAT)
    _paint_numeric_column(grid, data_start, total_row, 5, _MONEY_FORMAT)
    grid.paint(
        RangeStyle(
            start_row=total_row,
            start_col=1,
            end_row=total_row,
            end_col=COLUMN_COUNT,
            bold=True,
            fill_color=_TOTAL,
            font_color=_NAVY,
            border="top",
            border_color=_NAVY,
        )
    )


def _paint_summary_table(grid: _Grid, kpi_rows: dict[str, int]) -> None:
    """Рамка и форматы блока сводных показателей."""
    first = min(kpi_rows.values()) - 1
    last = max(kpi_rows.values())
    grid.paint(
        RangeStyle(
            start_row=first,
            start_col=1,
            end_row=last,
            end_col=COLUMN_COUNT,
            border="all",
            border_color=_LINE,
        )
    )
    grid.paint(
        RangeStyle(
            start_row=first,
            start_col=1,
            end_row=first,
            end_col=COLUMN_COUNT,
            bold=True,
            fill_color=_HEADER,
            font_color=_NAVY,
            horizontal="center",
        )
    )
    for key, row in kpi_rows.items():
        grid.paint(
            RangeStyle(
                start_row=row,
                start_col=1,
                end_row=row,
                end_col=2,
                bold=True,
                font_color=_NAVY,
            )
        )
        grid.paint(
            RangeStyle(
                start_row=row,
                start_col=5,
                end_row=row,
                end_col=6,
                font_color=_MUTED,
                font_size=10,
            )
        )
    money_keys = ("revenue", "cost", "check")
    int_keys = ("orders", "returns")
    percent_keys = ("return_rate", "margin")
    for key in money_keys:
        _paint_kpi_value(grid, kpi_rows[key], _MONEY_FORMAT)
    for key in int_keys:
        _paint_kpi_value(grid, kpi_rows[key], _INT_FORMAT)
    for key in percent_keys:
        _paint_kpi_value(grid, kpi_rows[key], _PERCENT_FORMAT)


def _paint_kpi_value(grid: _Grid, row: int, number_format: str) -> None:
    """Число сводки стоит в колонке C и растянуто на D."""
    grid.paint(
        RangeStyle(
            start_row=row,
            start_col=3,
            end_row=row,
            end_col=4,
            bold=True,
            font_size=14,
            font_color=_NAVY,
            horizontal="right",
            number_format=number_format,
        )
    )


def _paint_numeric_column(
    grid: _Grid,
    start_row: int,
    end_row: int,
    column: int,
    number_format: str,
) -> None:
    """Выровнять числовой столбец детализации и итога вправо."""
    grid.paint(
        RangeStyle(
            start_row=start_row,
            start_col=column,
            end_row=end_row,
            end_col=column,
            horizontal="right",
            number_format=number_format,
        )
    )


def _find_label(grid: _Grid, label: str) -> int:
    """Найти строку по подписи в первой колонке."""
    for index, row in enumerate(grid.rows, start=1):
        if row and row[0] == label:
            return index
    raise RuntimeError(f"В макете нет строки «{label}».")


def _section_rows(grid: _Grid) -> list[int]:
    """Найти строки заголовков разделов по тексту в первой колонке."""
    markers = ("1. Сводные показатели", "2. Детализация по дням", "3. Выводы")
    return [
        index
        for index, row in enumerate(grid.rows, start=1)
        if row and row[0] in markers
    ]


def _section_style(row: int) -> RangeStyle:
    """Заголовок раздела с золотой чертой, как в печатном документе."""
    return RangeStyle(
        start_row=row,
        start_col=1,
        end_row=row,
        end_col=COLUMN_COUNT,
        font_name="Georgia",
        font_size=14,
        font_color=_NAVY,
        bold=True,
        vertical="center",
        border="bottom",
        border_color=_GOLD,
    )


def _note_style(start_row: int, end_row: int) -> RangeStyle:
    """Абзац вывода в рамке."""
    return RangeStyle(
        start_row=start_row,
        start_col=1,
        end_row=end_row,
        end_col=COLUMN_COUNT,
        fill_color=_NOTE,
        font_size=11,
        font_color=_INK,
        vertical="top",
        wrap=True,
        border="all",
        border_color=_NOTE_LINE,
    )


def _sum_formula(column: str, start_row: int, end_row: int) -> str:
    """Сумма колонки. Без разделителя аргументов, чтобы формула жила в любой локали."""
    return f"=SUM({column}{start_row}:{column}{end_row})"


def _ratio_formula(numerator: str, denominator: str) -> str:
    """Частное. Если знаменатель нулевой, в делитель добавляется 1."""
    return f"=({numerator})/(({denominator})+(({denominator})=0))"


def _simulate_days(params: ReportParams, rng: Random) -> list[DayRecord]:
    """Смоделировать каждый день периода. Выходные заметно тише будней."""
    records: list[DayRecord] = []
    day_count = (params.end - params.start).days + 1
    base = _ORDER_BASE[params.department]
    low_check, high_check = _CHECK_BOUNDS[params.department]
    low_cost, high_cost = _COST_RATIO[params.kind]
    low_return, high_return = _RETURN_RATE[params.kind]
    for index in range(day_count):
        current = params.start + timedelta(days=index)
        position = index / max(day_count - 1, 1)
        wave = 1 + 0.12 * math.sin(position * math.pi)
        volume = base * _weekend_factor(current, params.department) * wave
        orders = int(round(volume * rng.uniform(0.82, 1.18)))
        if orders < 0:
            orders = 0
        if orders == 0:
            records.append(
                DayRecord(
                    day=current,
                    orders=0,
                    revenue=0.0,
                    returns_count=0,
                    cost=0.0,
                    comment="Операций не было",
                )
            )
            continue
        revenue = round(orders * rng.uniform(low_check, high_check), 2)
        returns_count = int(round(orders * rng.uniform(low_return, high_return)))
        if returns_count > orders:
            returns_count = orders
        cost = round(revenue * rng.uniform(low_cost, high_cost), 2)
        records.append(
            DayRecord(
                day=current,
                orders=orders,
                revenue=revenue,
                returns_count=returns_count,
                cost=cost,
                comment=_day_comment(params, current, base, orders, returns_count, rng),
            )
        )
    return records


def _weekend_factor(day: date, department: str) -> float:
    """Выходные тише. Производство и логистика снижают поток меньше, чем офис."""
    if day.weekday() < 5:
        return 1.0
    if department in {"Производство", "Логистика"}:
        return 0.72
    return 0.42


def _day_comment(
    params: ReportParams,
    day: date,
    base: int,
    orders: int,
    returns_count: int,
    rng: Random,
) -> str:
    """Короткая пометка дня. Зависит от нагрузки, а не только от жребия."""
    if day.weekday() >= 5 and orders < int(base * 0.7):
        return "Выходной, поток ниже обычного"
    if orders > 0 and returns_count / orders >= 0.07:
        return "Повышенная доля возвратов"
    if orders >= int(base * 1.25):
        return "Пиковая нагрузка"
    pool = _COMMENT_POOL[params.kind]
    return pool[rng.randrange(len(pool))]


def _conclusion_text(params: ReportParams, records: list[DayRecord]) -> str:
    """Короткий вывод по уже посчитанным итогам. В лист попадает готовым текстом."""
    orders = sum(record.orders for record in records)
    revenue = sum(record.revenue for record in records)
    cost = sum(record.cost for record in records)
    returns_count = sum(record.returns_count for record in records)
    margin = 0.0 if revenue == 0 else (revenue - cost) / revenue
    return_rate = 0.0 if orders == 0 else returns_count / orders
    average_check = 0.0 if orders == 0 else revenue / orders
    sentences = [
        (
            f"За период подразделение «{params.department}» провело {format_int(orders)} "
            f"заказов на {format_money(revenue)}."
        ),
        (
            f"Средний чек — {format_money(average_check)}, маржинальность — "
            f"{format_percent(margin)}, доля возвратов — {format_percent(return_rate)}."
        ),
    ]
    if margin < 0.22:
        sentences.append(
            "Маржа ниже ориентира 22%: стоит проверить скидки, закупки и списания."
        )
    elif margin > 0.35:
        sentences.append(
            "Маржа выше обычного коридора, период можно взять за ориентир."
        )
    else:
        sentences.append("Маржа в рабочем коридоре, резких отклонений по прибыли нет.")
    if return_rate > 0.05:
        sentences.append(
            "Возвратов больше обычного: разбор лучше начать с дней, "
            "где в комментарии отмечена повышенная доля."
        )
    else:
        sentences.append("Возвраты в пределах нормы, отдельный разбор не требуется.")
    return " ".join(sentences)
