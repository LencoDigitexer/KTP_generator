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

    Текст .doc разбивается по \\x07 (конец ячейки/строки таблицы). Формат строк
    КТП: [заголовок темы] [''] [теория] [практика] [всего] [дата дд.мм] [''].
    Строки «Кейс …» — разделы (без дат). Позиции дат в исходном тексте
    вычисляются точно (offset ячейки), что позволяет безопасно их перезаписать.
    """
    doc = KtpDocument()
    # offset каждой ячейки в исходном тексте
    offsets: list[int] = []
    off = 0
    raw_cells = text.split(CELL_SEP)
    for c in raw_cells:
        offsets.append(off)
        off += len(c) + 1
    cells = [c.replace("\n", " ").replace(LINE_SEP, " ").strip() for c in raw_cells]

    lesson_re = re.compile(r"^Тема\s+[\d.]+", re.IGNORECASE)
    case_re = re.compile(r"^Кейс\s+\d", re.IGNORECASE)
    total_re = re.compile(r"^Итого$", re.IGNORECASE)
    num_re = re.compile(r"^\d{1,3}$")
    date_re = re.compile(r"^\d{1,2}\.\d{2}$")

    idx = 0
    i = 0
    n = len(cells)
    while i < n:
        c = cells[i]
        if not (lesson_re.match(c) or case_re.match(c)):
            i += 1
            continue
        title = c
        nums: list[str] = []
        date_val: str | None = None
        date_span: tuple[int, int] | None = None
        j = i + 1
        while j < n and j - i <= 14:
            nxt = cells[j]
            if lesson_re.match(nxt) or case_re.match(nxt) or total_re.match(nxt):
                break
            if date_re.match(nxt):
                if date_val is None:
                    date_val = nxt
                    s = offsets[j] + raw_cells[j].find(nxt)
                    date_span = (s, s + len(nxt))
                j += 1
                continue
            if num_re.match(nxt):
                nums.append(nxt)
            elif nxt == "":
                pass
            elif not nums:
                title += " " + nxt
            else:
                break
            j += 1
        is_case = bool(case_re.match(c))
        row = LessonRow(
            index=(idx := idx + 1),
            title=title,
            hours_theory=nums[0] if len(nums) > 0 else "",
            hours_practice=nums[1] if len(nums) > 1 else "",
            hours_total=nums[2] if len(nums) > 2 else "",
            date=date_val,
            date_pos=date_span,
            is_case_header=is_case,
        )
        doc.lessons.append(row)
        i = j if j > i else i + 1
    return doc
