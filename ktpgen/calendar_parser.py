"""Чтение и разбор календаря-графика часов (XLSX).

Ожидаемый формат листа (как в «календарь часов 2026-2027.xlsx»):
  * строка 1 — дни недели (понедельник … воскресенье), каждый занимает 2 колонки;
  * далее сетка: в ячейках даты (datetime), рядом — счётчики недель/часов;
  * праздничные (нерабочие) дни помечены КРАСНЫМ цветом заливки.

Из такого листа извлекается:
  * множество рабочих дат и множество праздничных дат;
  * для каждого дня недели упорядоченный список рабочих дат.
"""

from __future__ import annotations

import datetime
from collections import defaultdict
from dataclasses import dataclass, field

import openpyxl

WEEKDAY_NAMES = [
    "понедельник", "вторник", "среда", "четверг",
    "пятница", "суббота", "воскресенье",
]


def _normalize_date(d: datetime.datetime | datetime.date) -> datetime.date:
    """Календарь может содержать «эталонный» год с неверным веком (2018 вместо 2026).

    Сохраняем месяц/день и подбираем год так, чтобы даты шли непрерывно
    от самой ранней даты листа (обычно сентябрь учебного года).
    """
    return d.date() if isinstance(d, datetime.datetime) else d


@dataclass
class Calendar:
    """Набор рабочих/праздничных дат по дням недели."""

    working_dates: list[datetime.date] = field(default_factory=list)
    holidays: set[datetime.date] = field(default_factory=set)
    by_weekday: dict[int, list[datetime.date]] = field(default_factory=dict)
    sheet_name: str = ""

    def dates_for_weekday(self, weekday: int) -> list[datetime.date]:
        """Рабочие даты указанного дня недели (0 = понедельник)."""
        return self.by_weekday.get(weekday, [])


def _is_red_fill(cell) -> bool:
    """True, если ячейка залита красным (праздник)."""
    fill = cell.fill
    if fill is None or fill.patternType is None:
        return False
    fg = fill.fgColor
    try:
        if fg.type == "theme":
            # theme 5 в стандартной теме Office — Accent2 (красный/оранжево-красный)
            return fg.theme in (4, 5, 6, 7, 8, 9, 10) and fg.theme == 5
        if fg.type == "rgb":
            rgb = str(fg.rgb)
            if len(rgb) >= 8:
                r, g, b = int(rgb[2:4], 16), int(rgb[4:6], 16), int(rgb[6:8], 16)
                return r > 150 and g < 130 and b < 130
        if fg.type == "indexed":
            return fg.indexed in (3, 10, 11, 12, 22, 25, 46)
    except (TypeError, ValueError):
        pass
    return False


def _is_weekend(d: datetime.date) -> bool:
    """Учитывает праздничные переносы РФ: 31.12 (суббота/воскресенье) — рабочий день."""
    if d.weekday() >= 5:
        if (d.month, d.day) == (12, 31):
            return False
        return True
    return False


def _detect_year_shift(raw_dates: list[datetime.date], target_hint: int | None) -> int:
    """Подбирает сдвиг года, чтобы набор дат соответствовал учебному году.

    Если в листе лежат даты 2018/2019/2024 (шаблон), а нужен 2026/2027 —
    сдвигаем все даты на один год вперёд относительно самого частого «якоря».
    Простейшая эвристика: берём модальный год и, если он не совпадает с
    подсказкой target_hint (или лежит вне диапазона 2025..2100 при её отсутствии),
    сдвигаем так, чтобы месяц/день сохранялись, а год стал target_hint для
    дат с месяцем >= 8 (сентябрь–декабрь) и target_hint+1 для месяцев <= 7.
    """
    if not raw_dates:
        return 0
    years = defaultdict(int)
    for d in raw_dates:
        years[d.year] += 1
    modal_year = max(years, key=lambda y: years[y])
    hint = target_hint if target_hint else None
    if hint is None:
        # если модальный год правдоподобен как текущий учебный — оставляем как есть
        if 2020 <= modal_year <= 2100:
            return 0
        hint = 2026
    shift = hint - modal_year
    return shift


def parse_calendar(path_or_stream, target_year: int | None = None,
                   remap_years: bool = True) -> Calendar:
    """Читает XLSX-календарь и возвращает структуру рабочих/праздничных дат.

    target_year — первый год учебного года (например, 2026 для 2026/2027).
    remap_years — переносить «шаблонные» годы (2018/2019/2024) в нужный год,
                  сохраняя месяц и день.
    """
    wb = openpyxl.load_workbook(path_or_stream, data_only=True)
    ws = wb.active

    # 1. Определяем сетку: строка заголовков с названиями дней недели и колонки дней.
    day_col: dict[int, int] = {}
    header_row_idx = None
    for row in ws.iter_rows(min_row=1, max_row=10):
        found = {}
        for cell in row:
            if isinstance(cell.value, str):
                name = cell.value.strip().lower()
                for wd, wname in enumerate(WEEKDAY_NAMES):
                    if name.startswith(wname[:4]):   # «понедельник» -> «поне»
                        found.setdefault(wd, cell.column)
        if len(found) >= 5:
            day_col = found
            header_row_idx = row[0].row
            break

    # 2. Собираем даты; день недели берём из колонки (в шаблоне stored-год
    #    не совпадает с фактическим днём недели сетки — это вёрстка календаря).
    raw: dict[datetime.date, bool] = {}          # дата -> красный?
    col_of: dict[datetime.date, int] = {}        # дата -> день недели по сетке
    explicit_non_red_cols: dict[tuple[int, int], int] = {}  # (строка, колонка) -> wd
    inv_day_col = {c: wd for wd, c in day_col.items()}

    def _wd_of(cell) -> int | None:
        if day_col:
            return inv_day_col.get(cell.column)
        v = cell.value
        return _normalize_date(v).weekday() if isinstance(v, (datetime.datetime, datetime.date)) else None

    for row in ws.iter_rows(min_row=(header_row_idx + 1) if header_row_idx else 1):
        for cell in row:
            v = cell.value
            if isinstance(v, (datetime.datetime, datetime.date)):
                d = _normalize_date(v)
                wd = _wd_of(cell)
                red = _is_red_fill(cell)
                if not red and wd is not None and wd >= 5:
                    explicit_non_red_cols[(cell.row, cell.column)] = wd
                prev = raw.get(d)
                raw[d] = bool(prev) or red if prev is not None else red
                if wd is not None:
                    col_of[d] = wd

    raw_dates = sorted(raw)
    if not raw_dates:
        raise ValueError("В календаре не найдено ни одной даты")

    # 3. Перенос дат шаблона (например 2018/2019/2024) в нужный учебный год.
    #    Месяц/день сохраняются; сентябрь–декабрь -> target_year, январь–август -> +1.
    modal_year = max(set(d.year for d in raw_dates),
                     key=lambda y: sum(1 for d in raw_dates if d.year == y))
    if remap_years and target_year is None and not (2020 <= modal_year <= 2100):
        target_year = 2026
    if target_year is not None and modal_year != target_year:
        def remap(d: datetime.date) -> datetime.date:
            ny = target_year if d.month >= 9 else target_year + 1
            try:
                return d.replace(year=ny)
            except ValueError:  # 29 февраля
                return d.replace(year=ny, day=28)
        remapped: dict[datetime.date, bool] = {}
        remap_wd: dict[datetime.date, int] = {}
        for d, red in raw.items():
            nd = remap(d)
            remapped[nd] = red or remapped.get(nd, False)
            if d in col_of:
                remap_wd[nd] = col_of[d]
        raw = remapped
        col_of = remap_wd

    holidays = {d for d, red in raw.items() if red}

    # 4. Выходные: сб/вс нерабочие. 31 декабря при субботе/воскресенье —
    #    рабочий день (праздничные переносы РФ).
    working = []
    for d in raw:
        if d in holidays:
            continue
        wd = col_of.get(d, d.weekday())
        if wd >= 5 and not (d.month == 12 and d.day == 31):
            continue
        working.append(d)
    working.sort()

    by_weekday: dict[int, list[datetime.date]] = defaultdict(list)
    for d in working:
        by_weekday[col_of.get(d, d.weekday())].append(d)

    return Calendar(
        working_dates=working,
        holidays=holidays,
        by_weekday=dict(by_weekday),
        sheet_name=ws.title,
    )
