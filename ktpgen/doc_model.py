"""Разбор текста КТП-документа: извлечение списка тем занятий и их дат."""

from __future__ import annotations

import re
from dataclasses import dataclass, field

# Управление текстом в .doc: \x07 — конец строки/ячейки таблицы, \r — конец абзаца,
# \x0b — «мягкая» расстановка (внутри ячейки), \x13..\x15 — поля.
CELL_SEP = "\x07"
PARA_SEP = "\r"
LINE_SEP = "\x0b"
FIELD_BEGIN, FIELD_SEP, FIELD_END = "\x13", "\x15", "\x14"

DATE_RE = re.compile(r"\d{1,2}\.\d{2}")
THEME_RE = re.compile(r"^\s*(?:№?\s*)?(Тема|Кейс)\s*[\d.]", re.IGNORECASE)


@dataclass
class LessonRow:
    """Строка КТП: одна тема занятия."""

    index: int                      # сквозной номер занятия (1..N)
    title: str                      # наименование темы
    hours_theory: str = ""          # теория (часов)
    hours_practice: str = ""         # практика (часов)
    hours_total: str = ""            # всего (часов)
    date: str | None = None          # дата в документе (дд.мм)
    date_pos: tuple[int, int] | None = None  # позиция даты в тексте документа
    is_case_header: bool = False     # строка — кейс/раздел, а не занятие


@dataclass
class KtpDocument:
    lessons: list[LessonRow] = field(default_factory=list)

    @property
    def session_rows(self) -> list[LessonRow]:
        return [l for l in self.lessons if not l.is_case_header]


def _strip_fields(text: str) -> str:
    """Удаляет коды управления полями, оставляя их отображаемый результат."""
    out = []
    depth = 0
    skip_result = False
    for ch in text:
        if ch == FIELD_BEGIN:
            depth += 1
            skip_result = False
        elif ch == FIELD_SEP:
            skip_result = True
        elif ch == FIELD_END:
            depth -= 1
            skip_result = False
        else:
            if depth == 0 or not skip_result:
                out.append(ch)
    return "".join(out)


def parse_lessons_from_text(text: str) -> KtpDocument:
    """Разбирает текст документа на строки-темы с датами.

    Анализируются «строки таблиц»: фрагменты между PARAF/CELL маркерами.
    Тема считается занятием, если её название начинается с «Тема …».
    Строки «Кейс …» — разделы (перенумеровываются, но без дат).
    """
    doc = KtpDocument()
    # Разбиваем по концам строк таблиц (\x07 идёт после каждой ячейки/строки)
    rows_raw = text.split(CELL_SEP)

    idx = 0
    current: list[str] = []
    merged_rows: list[list[str]] = []
    # Каждая строка таблицы заканчивается одиночным \x07 (TAP), ячейки — тоже \x07.
    # Проще: идём по ячейкам и группируем по признаку «начинается с номера/Тема/Кейс/Итого».
    cells = [c.replace("\n", " ").replace(LINE_SEP, " ").strip() for c in rows_raw]

    lesson_re = re.compile(r"^Тема\s+[\d.]+", re.IGNORECASE)
    case_re = re.compile(r"^Кейс\s+\d", re.IGNORECASE)
    total_re = re.compile(r"^Итого$", re.IGNORECASE)

    i = 0
    while i < len(cells):
        c = cells[i]
        if lesson_re.match(c) or case_re.match(c):
            # следующая ячейка может содержать продолжение заголовка + часы;
            # формат: [заголовок], [''], [теория], [практика], [всего], [дата], ['']
            title = c
            j = i + 1
            nums: list[str] = []
            date_val = None
            date_cell_index = None
            while j < len(cells) and j < i + 8:
                nxt = cells[j]
                if lesson_re.match(nxt) or case_re.match(nxt) or total_re.match(nxt):
                    break
                m = DATE_RE.fullmatch(nxt)
                if m and date_val is None and nums:
                    date_val = nxt
                    date_cell_index = j
                    j += 1
                    continue
                if re.fullmatch(r"\d{1,3}", nxt):
                    nums.append(nxt)
                elif nxt == "":
                    pass
                else:
                    # текст продолжения заголовка
                    if not nums:
                        title += " " + nxt
                j += 1
            is_case = bool(case_re.match(c))
            row = LessonRow(
                index=(idx := idx + 1),
                title=title,
                hours_theory=nums[0] if len(nums) > 0 else "",
                hours_practice=nums[1] if len(nums) > 1 else "",
                hours_total=nums[2] if len(nums) > 2 else "",
                date=date_val,
                is_case_header=is_case,
            )
            doc.lessons.append(row)
            i = j if j > i else i + 1
            continue
        i += 1

    # Позиции дат в исходном тексте сопоставляем строкам-занятиям эвристикой:
    # дата занятия находится сразу после заголовка темы (в пределах ~250 символов).
    positions = [(m.start(), m.group()) for m in DATE_RE.finditer(text)]
    used = set()
    pi = 0
    for row in doc.lessons:
        tpos = text.find(row.title[:30])
        if tpos < 0:
            continue
        best = None
        for k in range(pi, len(positions)):
            p, v = positions[k]
            if p < tpos:
                continue
            if p - tpos > 400:
                break
            if row.is_case_header and p - tpos > 60:
                break
            best = (k, p, v)
            break
        if best:
            k, p, v = best
            row.date = v
            row.date_pos = (p, p + len(v))
            used.add(k)
            pi = k + 1
    return doc
