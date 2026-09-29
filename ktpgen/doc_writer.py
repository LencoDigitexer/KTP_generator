"""Чтение и запись Word 97-2003 (.doc, OLE Compound File) без MS Word.

Реализовано:
  * extract_doc_text() — извлечение основного текста через piece table (CLX/FIB);
  * rewrite_doc_dates() — замена дат «дд.мм» в тексте .doc-файла.

Метод замены безопасен, потому что все новые даты имеют ту же длину (5 символов),
что и старые:_offsets форматирования (CHPX/PAPX индексируются по cp) не смещаются,
и достаточно побайтовой перезаписи кусков (pieces) в потоке WordDocument.
Для UTF-16-кусков дата кодируется напрямую; для 8-битных (compressed) кусков
даты — ASCII, поэтому тоже пишутся как есть.
"""

from __future__ import annotations

import re
import struct

import olefile

DATE_RE = re.compile(r"\d{1,2}\.\d{2}")

# Смещения в FIB
FIB_BASE_FLAGS = 0x0A
FC_MIN = 0x18
FC_MAC = 0x1C
CCP_TEXT = 0x4C          # FibRgLw97.cbMac at 64, ccpText at 68
_CCPT = 68
_FIB_BLOB = 154          # начало FibRgFcLcb97
_FC_CLX = _FIB_BLOB + 264
_LCB_CLX = _FIB_BLOB + 268


def _table_stream_name(word: bytes) -> str:
    return "1Table" if (struct.unpack_from("<H", word, FIB_BASE_FLAGS)[0] & 0x0200) else "0Table"


def _read_pieces(word: bytes, table: bytes):
    """Возвращает список кусков: dict(cp_start, cp_end, fc, compressed)."""
    fcClx = struct.unpack_from("<I", word, _FC_CLX)[0]
    lcbClx = struct.unpack_from("<I", word, _LCB_CLX)[0]
    clx = table[fcClx:fcClx + lcbClx]

    pos = 0
    while pos < len(clx) and clx[pos] == 1:          # Prc — пропускаем
        cb = struct.unpack_from("<H", clx, pos + 1)[0]
        pos += 3 + cb
    if pos >= len(clx) or clx[pos] != 2:             # должен быть Pcdt
        raise ValueError("Не найден Pcdt (piece table) в CLX")
    lcb = struct.unpack_from("<I", clx, pos + 1)[0]
    plcpcd = clx[pos + 5:pos + 5 + lcb]
    n = (lcb - 4) // 12
    cps = [struct.unpack_from("<I", plcpcd, i * 4)[0] for i in range(n + 1)]
    pcds_off = (n + 1) * 4
    pieces = []
    for i in range(n):
        fc_raw = struct.unpack_from("<I", plcpcd, pcds_off + i * 8 + 2)[0]
        compressed = bool(fc_raw & 0x40000000)
        fc = fc_raw & 0x3FFFFFFF
        pieces.append(dict(cp_start=cps[i], cp_end=cps[i + 1],
                           fc=fc, compressed=compressed))
    return pieces


def _piece_bytes(piece, word: bytes) -> bytes:
    cch = piece["cp_end"] - piece["cp_start"]
    if piece["compressed"]:
        return word[piece["fc"] // 2: piece["fc"] // 2 + cch]
    return word[piece["fc"]: piece["fc"] + 2 * cch]


def _piece_text(piece, word: bytes) -> str:
    raw = _piece_bytes(piece, word)
    if piece["compressed"]:
        return raw.decode("cp1251", errors="replace")
    return raw.decode("utf-16-le", errors="replace")


def extract_doc_text(data: bytes | str) -> str:
    """Извлекает основной текст документа .doc (путь или байты)."""
    with olefile.OleFileIO(data) as ole:
        word = ole.openstream("WordDocument").read()
        table = ole.openstream(_table_stream_name(word)).read()
    return "".join(_piece_text(p, word) for p in _read_pieces(word, table))


def rewrite_doc_dates(doc_path_or_bytes, replacements, out_path=None) -> bytes:
    """Заменяет участки основного текста .doc-файла.

    replacements: iterable of (cp_start, cp_end, new_text); длины old/new совпадают.
    Возвращает содержимое нового файла; при заданном out_path сохраняет его.
    """
    src = doc_path_or_bytes
    with olefile.OleFileIO(src) as ole:
        word = bytearray(ole.openstream("WordDocument").read())
        tbl_name = _table_stream_name(bytes(word))
        table = ole.openstream(tbl_name).read()
        pieces = _read_pieces(bytes(word), table)

    # группируем замены по кускам
    by_piece: dict[int, list[tuple[int, int, str]]] = {}
    for start, end, repl in replacements:
        old_len = end - start
        assert len(repl) == old_len, \
            f"Несовпадение длин: [{start}:{end}] {old_len} != {len(repl)} ({repl!r})"
        placed = False
        for pi, p in enumerate(pieces):
            if p["cp_start"] <= start and end <= p["cp_end"]:
                by_piece.setdefault(pi, []).append((start, end, repl))
                placed = True
                break
        if not placed:
            raise ValueError(f"Участок [{start}:{end}] пересекает границы кусков — "
                             "требуется пересборка piece table")

    for pi, reps in by_piece.items():
        p = pieces[pi]
        if p["compressed"]:
            base = p["fc"] // 2
            buf = bytearray(word[base: base + (p["cp_end"] - p["cp_start"])])
            for start, end, repl in reps:
                off = start - p["cp_start"]
                seg = repl.encode("cp1251", errors="replace")
                assert len(seg) == end - start
                buf[off: off + len(seg)] = seg
            word[base: base + len(buf)] = buf
        else:
            base = p["fc"]
            for start, end, repl in reps:
                off = (start - p["cp_start"]) * 2
                seg = repl.encode("utf-16-le")
                word[base + off: base + off + len(seg)] = seg

    result = bytes(word)
    # сохраняем остальные потоки без изменений, переписывая WordDocument
    if out_path:
        _save_doc(src, {"WordDocument": result}, out_path)
    return result


def _save_doc(src, updated_streams: dict[str, bytes], out_path: str):
    """Сохраняет .doc с заменёнными потоками через olefile write-поддержку нет —
    поэтому делаем побайтовую замену внутри OLE: читаем все потоки и пересобираем."""
    import io
    streams: dict[str, bytes] = {}
    with olefile.OleFileIO(src) as ole:
        for en in ole.listdir():
            name = "/".join(en)
            data = ole.openstream(en).read()
            streams[name] = updated_streams.get(name, data)
    data = _build_ole(streams)
    with open(out_path, "wb") as f:
        f.write(data)


# ------------------------------------------------------------- OLE builder
SECT = 512
MAXREG = 0xFFFFFFFA
DIFSECT = 0xFFFFFFFC
FATSECT = 0xFFFFFFFD
ENDOFCHAIN = 0xFFFFFFFE
FREESECT = 0xFFFFFFFF


def _chain_fat(fat: list[int], start: int, nsects: int):
    for j in range(nsects):
        s = start + j
        fat[s] = s + 1 if j + 1 < nsects else ENDOFCHAIN


def _build_ole(streams: dict[str, bytes]) -> bytes:
    """Собирает минимальный OLE CFB (без MiniStream — порог mini = 0)."""
    names = list(streams.keys())
    n_sects = {k: (len(v) + SECT - 1) // SECT for k, v in streams.items()}
    total_data = sum(n_sects.values())
    dir_sects = max(1, (len(names) + 1 + 3) // 4)

    fat_sects = 1
    while True:
        total = fat_sects + dir_sects + total_data
        need = max(1, (total * 4 + SECT - 1) // SECT)
        if need <= fat_sects:
            break
        fat_sects = need

    fat_start = 0
    dir_start = fat_sects
    data_start = fat_sects + dir_sects
    total_sects = fat_sects + dir_sects + total_data

    fat = [FREESECT] * (fat_sects * SECT // 4)
    _chain_fat(fat, fat_start, fat_sects)
    _chain_fat(fat, dir_start, dir_sects)

    cur = data_start
    placement = {}
    for k in names:
        placement[k] = (cur, n_sects[k])
        _chain_fat(fat, cur, n_sects[k])
        cur += n_sects[k]

    fat_bytes = b"".join(struct.pack("<I", x) for x in fat[:fat_sects * 128])

    def dir_entry(name, etype, start, size, child=-1, left=-1, right=-1):
        e = bytearray(128)
        nb = name.encode("utf-16-le") + b"\x00\x00"
        e[0:len(nb)] = nb
        struct.pack_into("<H", e, 64, len(nb))
        e[66] = etype
        e[67] = 1  # black
        for off, val in ((68, left), (72, right), (76, child)):
            struct.pack_into("<I", e, off, (val & 0xFFFFFFFF) if val >= 0 else ENDOFCHAIN)
        struct.pack_into("<I", e, 120, ENDOFCHAIN if start < 0 else start)
        struct.pack_into("<Q", e, 124, size)   # low+high 32-bit
        return e

    entries = [dir_entry("Root Entry", 5, -1, 0, child=1 if names else -1)]
    for idx, nm in enumerate(names):
        st, ns = placement[nm]
        left_i = idx - 1 if idx > 0 else -1
        right_i = idx + 1 if idx + 1 < len(names) else -1
        entries.append(dir_entry(nm, 2, st, len(streams[nm]),
                                 left=left_i, right=right_i))
    dir_bytes = b"".join(bytes(e) for e in entries)
    dir_bytes = dir_bytes.ljust(dir_sects * SECT, b"\x00")

    h = bytearray(SECT)
    h[0:8] = bytes.fromhex("d0cf11e0a1b11ae1")
    struct.pack_into("<H", h, 0x18, 0x3E)
    struct.pack_into("<H", h, 0x1A, 0x0003)
    struct.pack_into("<H", h, 0x1C, 0xFFFE)
    struct.pack_into("<H", h, 0x1E, 9)
    struct.pack_into("<H", h, 0x20, 6)
    struct.pack_into("<I", h, 0x28, fat_sects)
    struct.pack_into("<I", h, 0x30, dir_start)
    struct.pack_into("<I", h, 0x38, 0xFFFFFFFE)     # mini stream start (нет)
    struct.pack_into("<I", h, 0x3C, 0)              # num mini fat sectors
    struct.pack_into("<I", h, 0x40, ENDOFCHAIN)     # mini fat start
    struct.pack_into("<I", h, 0x44, 0)              # num difat
    struct.pack_into("<I", h, 0x48, ENDOFCHAIN)     # difat start
    struct.pack_into("<I", h, 0x4C, 0)
    for i in range(109):
        struct.pack_into("<I", h, 0x4C + 4 + 4 * i,
                         fat_start + i if i < fat_sects else FREESECT)
    struct.pack_into("<I", h, 0x34, 0)              # num ministream sectors
    struct.pack_into("<I", h, 0x18 + 0x18, 0x3E)    # no-op keep signature minor

    out = io.BytesIO()
    out.write(bytes(h))
    out.write(fat_bytes)
    out.write(dir_bytes)
    for k in names:
        d = streams[k]
        out.write(d)
        out.write(b"\x00" * (n_sects[k] * SECT - len(d)))
    return out.getvalue()
