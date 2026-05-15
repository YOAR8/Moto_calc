#!/usr/bin/env python3
import argparse
import datetime as dt
import os
import platform
import subprocess
import re
import shutil
import sys
import traceback
from pathlib import Path
from typing import Dict, Tuple

import xlrd
from xlutils.copy import copy as xl_copy

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except Exception:  # pragma: no cover
    tk = None
    filedialog = None
    messagebox = None
    ttk = None

IS_WINDOWS = platform.system().lower().startswith("win")


def runtime_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def default_output_dir(base_dir: Path) -> Path:
    if IS_WINDOWS:
        return Path.home() / "Documents" / "MotoCalc" / "out"
    return base_dir / "out"


def a1_to_rc(addr: str) -> Tuple[int, int]:
    m = re.fullmatch(r"([A-Z]+)(\d+)", addr.upper())
    if not m:
        raise ValueError(f"Invalid address: {addr}")
    col_s, row_s = m.groups()
    col = 0
    for ch in col_s:
        col = col * 26 + (ord(ch) - 64)
    return int(row_s) - 1, col - 1


def rc_to_a1(row: int, col: int) -> str:
    c = col + 1
    out = ""
    while c:
        c, rem = divmod(c - 1, 26)
        out = chr(65 + rem) + out
    return f"{out}{row + 1}"


def read_xls_cell(path: Path, sheet_name: str, addr: str):
    wb = xlrd.open_workbook(str(path), formatting_info=False)
    sh = wb.sheet_by_name(sheet_name)
    r, c = a1_to_rc(addr)
    return sh.cell_value(r, c)


def write_xls_cells(path: Path, sheet_name: str, updates: Dict[str, object], backup: bool = True) -> None:
    if backup:
        ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = path.with_name(f"{path.stem}_backup_{ts}{path.suffix}")
        shutil.copy2(path, backup_path)

    rb = xlrd.open_workbook(str(path), formatting_info=True)
    wb = xl_copy(rb)

    sheet_idx = None
    for i, name in enumerate(rb.sheet_names()):
        if name == sheet_name:
            sheet_idx = i
            break
    if sheet_idx is None:
        raise ValueError(f"Sheet not found: {sheet_name}")

    ws = wb.get_sheet(sheet_idx)
    for addr, value in updates.items():
        r, c = a1_to_rc(addr)
        ws.write(r, c, value)

    wb.save(str(path))


def split_number_date(text: str) -> Tuple[str, str]:
    # Example: "№ 8424/26/000308 від 15 травня 2026 року"
    parts = str(text).split()
    if len(parts) >= 7 and parts[0] == "№" and "від" in parts:
        number = parts[1]
        try:
            vidx = parts.index("від")
            date_txt = " ".join(parts[vidx + 1 : vidx + 5])
        except ValueError:
            date_txt = ""
        return number, date_txt

    # Fallback parser
    m = re.search(r"№\s*([^\s]+).*?від\s+(.+)$", str(text))
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return "", ""


def short_name(full_name: str) -> str:
    tokens = [t for t in str(full_name).strip().split() if t]
    if len(tokens) < 3:
        return str(full_name).strip()
    return f"{tokens[0]} {tokens[1][0]}. {tokens[2][0]}."


def load_source_values(source_6055: Path) -> Dict[str, object]:
    wb = xlrd.open_workbook(str(source_6055), formatting_info=False)
    sh = wb.sheet_by_name("Worksheet")

    def g(addr: str):
        r, c = a1_to_rc(addr)
        return sh.cell_value(r, c)

    a3 = str(g("A3"))
    number, date_txt = split_number_date(a3)
    fio = str(g("C15"))

    return {
        "number": number,
        "date": date_txt,
        "fio": fio,
        "fio_short": short_name(fio),
        "pasport": str(g("C17")),
        "tax": str(g("C18")),
        "birthday": str(g("C12")),
        "adres": str(g("C16")),
        "decl": str(g("C47")),
        "model": str(g("C20")),
        "year": str(g("C29")),
        "color": str(g("C28")),
        "numberdv": str(g("C41")),
        "cuzov": str(g("C39")),
        "cub": str(g("C36")),
        "znak": str(g("C50")),
        "price": str(g("C46")),
    }


def transfer_to_act_non_windows(source_6055: Path, moto_template: Path, out_path: Path) -> None:
    src = load_source_values(source_6055)

    shutil.copy2(moto_template, out_path)
    updates = {
        "B12": src["number"],
        "L12": src["date"],
        "H18": src["fio"],
        "C24": src["model"],
        "E24": src["year"],
        "H24": src["cuzov"],
        "K24": src["color"],
        "N33": src["fio_short"],
    }
    write_xls_cells(out_path, "Worksheet", updates, backup=False)


def transfer_to_word_preview(source_6055: Path, out_txt: Path) -> None:
    src = load_source_values(source_6055)
    lines = [
        "This is a Linux preview of Word bookmark payload.",
        "Use Windows mode to write into DOGOVIR_6055_template.doc bookmarks.",
        "",
    ]
    mapping = [
        ("Number", src["number"]),
        ("Data", src["date"]),
        ("FIO", src["fio"]),
        ("pasport", src["pasport"]),
        ("TaxNumber", src["tax"]),
        ("BirthDay", src["birthday"]),
        ("adres", src["adres"]),
        ("decl", src["decl"]),
        ("model", src["model"]),
        ("year", src["year"]),
        ("color", src["color"]),
        ("numberdv", src["numberdv"]),
        ("cuzov", src["cuzov"]),
        ("cub", src["cub"]),
        ("znak", src["znak"]),
        ("price", src["price"]),
        ("sumtext", "(computed by VBA in original workbook)"),
        ("FIO2", src["fio"]),
    ]
    for key, value in mapping:
        lines.append(f"{key}: {value}")

    out_txt.write_text("\n".join(lines), encoding="utf-8")


def windows_transfer_all(source_6055: Path, moto_template: Path, dogovir_template: Path, out_dir: Path) -> None:
    import win32com.client as win32  # type: ignore

    out_dir.mkdir(parents=True, exist_ok=True)
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    word = win32.Dispatch("Word.Application")
    word.Visible = False

    try:
        wb_src = excel.Workbooks.Open(str(source_6055.resolve()))
        wb_dst = excel.Workbooks.Open(str(moto_template.resolve()))

        ws_src = wb_src.Worksheets("Worksheet")
        ws_dst = wb_dst.Worksheets("Worksheet")

        a3 = str(ws_src.Range("A3").Value)
        number, date_txt = split_number_date(a3)
        fio = str(ws_src.Range("C15").Value)

        ws_dst.Range("B12").Value = number
        ws_dst.Range("L12").Value = date_txt
        ws_dst.Range("H18").Value = fio
        ws_dst.Range("C24").Value = ws_src.Range("C20").Value
        ws_dst.Range("E24").Value = ws_src.Range("C29").Value
        ws_dst.Range("H24").Value = ws_src.Range("C39").Value
        ws_dst.Range("K24").Value = ws_src.Range("C28").Value
        ws_dst.Range("N33").Value = short_name(fio)

        ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        out_act = out_dir / f"6055_MOTO_filled_{ts}.xls"
        wb_dst.SaveAs(str(out_act.resolve()))

        doc = word.Documents.Open(str(dogovir_template.resolve()))
        bm = doc.Bookmarks

        payload = load_source_values(source_6055)
        payload["Number"] = payload.pop("number")
        payload["Data"] = payload.pop("date")
        payload["TaxNumber"] = payload.pop("tax")
        payload["BirthDay"] = payload.pop("birthday")
        payload["FIO"] = payload.pop("fio")
        payload["FIO2"] = payload["FIO"]

        bookmark_map = {
            "Number": "Number",
            "Data": "Data",
            "FIO": "FIO",
            "pasport": "pasport",
            "TaxNumber": "TaxNumber",
            "BirthDay": "BirthDay",
            "adres": "adres",
            "decl": "decl",
            "model": "model",
            "year": "year",
            "color": "color",
            "numberdv": "numberdv",
            "cuzov": "cuzov",
            "cub": "cub",
            "znak": "znak",
            "price": "price",
            "FIO2": "FIO2",
        }

        for key, bookmark_name in bookmark_map.items():
            if bm.Exists(bookmark_name):
                rng = bm(bookmark_name).Range
                rng.Text = str(payload[key])

        out_doc = out_dir / f"DOGOVIR_6055_filled_{ts}.doc"
        doc.SaveAs(str(out_doc.resolve()))
        doc.Close(SaveChanges=False)

        wb_dst.Close(SaveChanges=False)
        wb_src.Close(SaveChanges=True)
    finally:
        try:
            excel.Quit()
        except Exception:
            pass
        try:
            word.Quit()
        except Exception:
            pass


FIELD_SECTIONS = [
    (
        "Документ",
        [
            ("A3", "Номер і дата", True),
            ("C51", "Дата набуття права", False),
            ("C47", "Митна декларація", False),
        ],
    ),
    (
        "Покупець",
        [
            ("C15", "ПІБ покупця", True),
            ("E15", "ПІБ скорочено", False),
            ("C12", "Дата народження", False),
            ("C16", "Адреса", True),
            ("C17", "Паспорт", True),
            ("C18", "ІПН / код", True),
        ],
    ),
    (
        "Транспорт 1",
        [
            ("C20", "Марка, модель", True),
            ("C21", "Тип ТЗ", False),
            ("C22", "Призначення", False),
            ("C25", "Категорія", False),
            ("C26", "Тип кузова", False),
            ("C27", "Паливо", False),
            ("C28", "Колір", True),
            ("C29", "Рік випуску", True),
            ("C30", "Повна маса", False),
            ("C31", "Маса без навантаження", False),
            ("C33", "Кількість місць", False),
        ],
    ),
    (
        "Транспорт 2",
        [
            ("C35", "Кількість циліндрів", False),
            ("C36", "Об'єм двигуна", True),
            ("C37", "Потужність кВт", False),
            ("C39", "VIN / номер рами", True),
            ("C41", "Номер двигуна", True),
            ("C42", "Номер шасі", True),
            ("C43", "Номер рами", True),
            ("C44", "Ціна без ПДВ", True),
            ("C45", "ПДВ", True),
            ("C46", "Ціна з ПДВ", True),
            ("C50", "Номерні знаки", False),
        ],
    ),
    (
        "Додатково",
        [
            ("C23", "Сертифікат відповідності", False),
            ("C24", "Номер сертифіката типу", False),
            ("C32", "Кількість дверей", False),
            ("C34", "Кількість стоячих місць", False),
            ("C38", "Потужність к.с.", False),
            ("C40", "Номер кузова", False),
            ("C48", "Акт приймання-передачі", False),
            ("C49", "Свідоцтво про реєстрацію", False),
            ("C51", "Дата набуття права", False),
            ("C53", "Висновок", False),
            ("A56", "Продавець", False),
            ("D56", "Покупець (підпис)", False),
        ],
    ),
]


def iter_field_specs():
    for group_name, fields in FIELD_SECTIONS:
        for cell, label, required in fields:
            yield group_name, cell, label, required


FORM_CELLS = [cell for _, cell, _, _ in iter_field_specs()]


def safe_read_cell(sheet, addr: str):
    try:
        r, c = a1_to_rc(addr)
        if r < 0 or c < 0 or r >= sheet.nrows or c >= sheet.ncols:
            return ""
        value = sheet.cell_value(r, c)
        if isinstance(value, float) and value.is_integer():
            return int(value)
        return value
    except Exception:
        return ""


def load_form_state(source_path: Path) -> Dict[str, str]:
    wb = xlrd.open_workbook(str(source_path), formatting_info=False)
    sh = wb.sheet_by_name("Worksheet")
    state: Dict[str, str] = {}
    for cell in FORM_CELLS:
        state[cell] = str(safe_read_cell(sh, cell)).strip()
    if not state.get("E15"):
        state["E15"] = short_name(state.get("C15", ""))
    return state


def parse_state(state: Dict[str, str]) -> Dict[str, str]:
    payload = dict(state)
    number, date_txt = split_number_date(payload.get("A3", ""))
    payload["Number"] = number
    payload["Data"] = date_txt
    payload["FIO"] = payload.get("C15", "")
    payload["FIO2"] = payload.get("C15", "")
    payload["TaxNumber"] = payload.get("C18", "")
    payload["BirthDay"] = payload.get("C12", "")
    payload["adres"] = payload.get("C16", "")
    payload["decl"] = payload.get("C47", "")
    payload["model"] = payload.get("C20", "")
    payload["year"] = payload.get("C29", "")
    payload["color"] = payload.get("C28", "")
    payload["numberdv"] = payload.get("C41", "")
    payload["cuzov"] = payload.get("C39", "") or payload.get("C42", "") or payload.get("C43", "")
    payload["cub"] = payload.get("C36", "")
    payload["znak"] = payload.get("C50", "")
    payload["price"] = payload.get("C46", "")
    payload["sumtext"] = amount_to_words_uah(payload.get("C46", ""))
    payload["fio_short"] = short_name(payload.get("C15", ""))
    return payload


def amount_to_words_uah(value) -> str:
    try:
        amount = int(float(str(value).replace(",", ".")))
    except Exception:
        return ""

    if amount <= 0:
        return "нуль грн. 00 коп."

    ones = ["", "один", "два", "три", "чотири", "п'ять", "шість", "сім", "вісім", "дев'ять"]
    ones_f = ["", "одна", "дві", "три", "чотири", "п'ять", "шість", "сім", "вісім", "дев'ять"]
    teens = ["десять", "одинадцять", "дванадцять", "тринадцять", "чотирнадцять", "п'ятнадцять", "шістнадцять", "сімнадцять", "вісімнадцять", "дев'ятнадцять"]
    tens = ["", "", "двадцять", "тридцять", "сорок", "п'ятдесят", "шістдесят", "сімдесят", "вісімдесят", "дев'яносто"]
    hundreds = ["", "сто", "двісті", "триста", "чотириста", "п'ятсот", "шістсот", "сімсот", "вісімсот", "дев'ятсот"]

    def triad_words(num: int, feminine: bool = False) -> str:
        if num == 0:
            return ""
        parts = []
        h = num // 100
        t = (num // 10) % 10
        u = num % 10
        if h:
            parts.append(hundreds[h])
        if t == 1:
            parts.append(teens[u])
        else:
            if t:
                parts.append(tens[t])
            if u:
                parts.append((ones_f if feminine else ones)[u])
        return " ".join([p for p in parts if p])

    groups = [
        (amount % 1000, False, ""),
        ((amount // 1000) % 1000, True, "тисяч"),
        ((amount // 1000000) % 1000, False, "мільйонів"),
    ]

    chunks = []
    for idx, (value_part, feminine, suffix) in enumerate(groups):
        if not value_part:
            continue
        text = triad_words(value_part, feminine=feminine)
        if idx == 1:
            last_two = value_part % 100
            last = value_part % 10
            if 10 < last_two < 20:
                suffix = "тисяч"
            elif last == 1:
                suffix = "тисяча"
            elif last in (2, 3, 4):
                suffix = "тисячі"
            else:
                suffix = "тисяч"
        elif idx == 2:
            last_two = value_part % 100
            last = value_part % 10
            if 10 < last_two < 20:
                suffix = "мільйонів"
            elif last == 1:
                suffix = "мільйон"
            elif last in (2, 3, 4):
                suffix = "мільйона"
            else:
                suffix = "мільйонів"
        if suffix:
            chunks.append(f"{text} {suffix}".strip())
        else:
            chunks.append(text)

    return " ".join(reversed([chunk for chunk in chunks if chunk])).strip() + " грн. 00 коп."


def parse_decimal(value: str) -> float | None:
    try:
        cleaned = str(value).strip().replace(" ", "").replace(",", ".")
        if not cleaned:
            return None
        return float(cleaned)
    except Exception:
        return None


def normalize_token(value: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(value).upper())


def valid_date_text(value: str) -> bool:
    value = str(value).strip()
    if not value:
        return False
    if re.fullmatch(r"\d{2}\.\d{2}\.\d{4}", value):
        return True
    return bool(re.search(r"\b\d{1,2}\b.+\b\d{4}\b", value))


def preview_blocks_for_act(payload: Dict[str, str]) -> list[tuple[str, list[tuple[str, str]]]]:
    return [
        ("Реквізити акта", [("Номер", payload.get("Number", "")), ("Дата", payload.get("Data", ""))]),
        (
            "Покупець",
            [
                ("ПІБ", payload.get("C15", "")),
                ("Адреса", payload.get("C16", "")),
                ("Паспорт", payload.get("C17", "")),
                ("ІПН", payload.get("C18", "")),
            ],
        ),
        (
            "Транспортний засіб",
            [
                ("Марка / модель", payload.get("C20", "")),
                ("Тип", payload.get("C21", "")),
                ("Рік", payload.get("C29", "")),
                ("Колір", payload.get("C28", "")),
                ("VIN / рама", payload.get("cuzov", "")),
                ("Номер двигуна", payload.get("C41", "")),
            ],
        ),
        ("Підписні дані", [("Скорочений ПІБ", payload.get("fio_short", "")), ("Дата набуття права", payload.get("C51", ""))]),
    ]


def preview_blocks_for_contract(payload: Dict[str, str]) -> list[tuple[str, list[tuple[str, str]]]]:
    return [
        ("Договір", [("Номер", payload.get("Number", "")), ("Дата", payload.get("Data", "")), ("Ціна", payload.get("price", "")), ("Сума прописом", payload.get("sumtext", ""))]),
        (
            "Покупець",
            [
                ("ПІБ", payload.get("FIO", "")),
                ("Дата народження", payload.get("BirthDay", "")),
                ("Паспорт", payload.get("pasport", "")),
                ("ІПН", payload.get("TaxNumber", "")),
                ("Адреса", payload.get("adres", "")),
            ],
        ),
        (
            "Мотоцикл",
            [
                ("Модель", payload.get("model", "")),
                ("Рік", payload.get("year", "")),
                ("Колір", payload.get("color", "")),
                ("Рама / кузов", payload.get("cuzov", "")),
                ("Двигун", payload.get("numberdv", "")),
                ("Об'єм", payload.get("cub", "")),
                ("Номерні знаки", payload.get("znak", "")),
            ],
        ),
        ("Митні дані", [("Декларація", payload.get("decl", ""))]),
    ]


def preview_blocks_for_vidatkova(payload: Dict[str, str]) -> list[tuple[str, list[tuple[str, str]]]]:
    return [
        ("Видаткова накладна", [("Дата / номер", payload.get("A3", "")), ("Одержувач", payload.get("C15", ""))]),
        (
            "Позиція",
            [
                ("Тип ТЗ", payload.get("C21", "")),
                ("Марка, модель", payload.get("C20", "")),
                ("Номер рами", payload.get("cuzov", "")),
                ("Кількість", "1"),
            ],
        ),
        (
            "Суми",
            [
                ("Ціна без ПДВ", payload.get("C44", "")),
                ("ПДВ", payload.get("C45", "")),
                ("Всього", payload.get("C46", "")),
                ("Сума прописом", payload.get("sumtext", "")),
            ],
        ),
    ]


def preview_blocks(kind: str, payload: Dict[str, str]) -> list[tuple[str, list[tuple[str, str]]]]:
    if kind == "act":
        return preview_blocks_for_act(payload)
    if kind == "contract":
        return preview_blocks_for_contract(payload)
    return preview_blocks_for_vidatkova(payload)


def validate_state(state: Dict[str, str]) -> Tuple[Dict[str, str], list[str], list[str]]:
    payload = parse_state(state)
    errors: list[str] = []
    warnings: list[str] = []

    required = [cell for _, cell, _, req in iter_field_specs() if req]
    for cell in required:
        if not str(state.get(cell, "")).strip():
            errors.append(f"{cell} порожнє")

    if not payload["Number"] or not payload["Data"]:
        errors.append("A3 не парситься як номер + дата")
    elif not re.fullmatch(r"\d{4}/\d{2}/\d{6}", payload["Number"]):
        warnings.append("Номер документа в A3 має незвичний формат")

    if state.get("C51") and not valid_date_text(state.get("C51", "")):
        warnings.append("C51 має підозрілий формат дати")

    if state.get("C12") and not valid_date_text(state.get("C12", "")):
        warnings.append("C12 має підозрілий формат дати народження")

    frame_values = [str(state.get(c, "")).strip() for c in ("C39", "C42", "C43") if str(state.get(c, "")).strip()]
    if len(set(frame_values)) > 1:
        warnings.append("C39 / C42 / C43 не збігаються")

    if payload.get("cuzov"):
        vin_token = normalize_token(payload["cuzov"])
        if len(vin_token) < 8:
            errors.append("VIN / рама занадто короткий")
        elif len(vin_token) not in (8, 9, 10, 11, 12, 13, 14, 15, 16, 17):
            warnings.append("VIN / рама має незвичну довжину")

    if payload.get("C41"):
        engine_token = normalize_token(payload["C41"])
        if len(engine_token) < 5:
            warnings.append("Номер двигуна схожий на короткий або неповний")
        if payload.get("cuzov") and engine_token == normalize_token(payload["cuzov"]):
            warnings.append("Номер двигуна збігається з VIN / рамою, перевірте дані")

    if payload.get("C28") and payload["C28"].strip().isdigit():
        warnings.append("Колір виглядає як число")

    if payload.get("C18"):
        tax = re.sub(r"\D", "", payload["C18"])
        if len(tax) not in (8, 10):
            warnings.append("ІПН / код має незвичну довжину")

    if payload.get("C29"):
        try:
            year = int(str(payload["C29"]).strip())
            if year < 1900 or year > dt.datetime.now().year + 1:
                warnings.append("Рік випуску поза нормальним діапазоном")
        except Exception:
            warnings.append("Рік випуску не є числом")

    if payload.get("C46"):
        try:
            float(str(payload["C46"]).replace(",", "."))
        except Exception:
            errors.append("C46 не є числом")

    price_no_vat = parse_decimal(payload.get("C44", ""))
    vat = parse_decimal(payload.get("C45", ""))
    price_total = parse_decimal(payload.get("C46", ""))
    if price_no_vat is not None and vat is not None and price_total is not None:
        if abs((price_no_vat + vat) - price_total) > 0.01:
            errors.append("C44 + C45 не дорівнює C46")

    if price_no_vat is not None and price_total is not None and vat is not None and price_total > 0:
        vat_rate = round(vat / price_no_vat * 100, 2) if price_no_vat else None
        if price_no_vat and vat_rate is not None and vat_rate not in (20.0, 7.0, 0.0):
            warnings.append("Ставка ПДВ має незвичне значення")

    return payload, errors, warnings


def preview_text_for_act(payload: Dict[str, str]) -> str:
    return "\n".join([
        "ЧОРНОВИК АКТА 6055",
        "",
        f"Номер: {payload.get('Number', '')}",
        f"Дата: {payload.get('Data', '')}",
        f"Покупець: {payload.get('C15', '')}",
        f"Адреса: {payload.get('C16', '')}",
        f"Паспорт: {payload.get('C17', '')}",
        f"ІПН / код: {payload.get('C18', '')}",
        f"Модель: {payload.get('C20', '')}",
        f"Рік: {payload.get('C29', '')}",
        f"Рама / VIN: {payload.get('cuzov', '')}",
        f"Колір: {payload.get('C28', '')}",
        f"Скорочений ПІБ: {payload.get('fio_short', '')}",
    ])


def preview_text_for_contract(payload: Dict[str, str]) -> str:
    lines = ["ЧОРНОВИК ДОГОВОРУ", ""]
    for key in [
        ("Number", "Number"),
        ("Data", "Data"),
        ("FIO", "FIO"),
        ("pasport", "pasport"),
        ("TaxNumber", "TaxNumber"),
        ("BirthDay", "BirthDay"),
        ("adres", "adres"),
        ("decl", "decl"),
        ("model", "model"),
        ("year", "year"),
        ("color", "color"),
        ("numberdv", "numberdv"),
        ("cuzov", "cuzov"),
        ("cub", "cub"),
        ("znak", "znak"),
        ("price", "price"),
        ("sumtext", "sumtext"),
        ("FIO2", "FIO2"),
    ]:
        lines.append(f"{key[0]}: {payload.get(key[1], '')}")
    return "\n".join(lines)


def preview_text_for_vidatkova(payload: Dict[str, str]) -> str:
    return "\n".join([
        "ЧОРНОВИК ВИДАТКОВОЇ",
        "",
        f"Одержувач: {payload.get('C15', '')}",
        f"Дата / номер: {payload.get('A3', '')}",
        f"Марка, модель: {payload.get('C20', '')}",
        f"Номер рами: {payload.get('cuzov', '')}",
        f"Ціна без ПДВ: {payload.get('C44', '')}",
        f"ПДВ: {payload.get('C45', '')}",
        f"Всього: {payload.get('C46', '')}",
        f"Сума прописом: {payload.get('sumtext', '')}",
    ])


def generate_act_xls_from_state(state: Dict[str, str], moto_template: Path, out_path: Path) -> Path:
    payload = parse_state(state)
    shutil.copy2(moto_template, out_path)
    updates = {
        "B12": payload["Number"],
        "L12": payload["Data"],
        "H18": payload["FIO"],
        "C24": payload["model"],
        "E24": payload["year"],
        "H24": payload["cuzov"],
        "K24": payload["color"],
        "N33": payload["fio_short"],
    }
    write_xls_cells(out_path, "Worksheet", updates, backup=False)
    return out_path


def generate_vidatkova_xls_from_state(state: Dict[str, str], template_path: Path, out_path: Path) -> Path:
    payload = parse_state(state)
    shutil.copy2(template_path, out_path)
    updates = {
        "C6": payload["FIO"],
        "C7": "той самий",
        "G9": payload["A3"],
        "D10": payload["A3"],
        "B13": payload.get("C21", "МОПЕД"),
        "C13": payload["model"],
        "D13": payload["cuzov"],
        "E13": "шт",
        "F13": 1,
        "G13": payload.get("C44", ""),
        "H13": payload.get("C44", ""),
        "H15": payload.get("C44", ""),
        "H16": payload.get("C45", ""),
        "H17": payload.get("C46", ""),
        "A19": payload.get("sumtext", ""),
        "B20": payload.get("C45", ""),
        "F23": payload["FIO"],
    }
    write_xls_cells(out_path, "Лист1", updates, backup=False)
    return out_path


def generate_contract_doc_windows_from_state(state: Dict[str, str], dogovir_template: Path, out_path: Path) -> Path:
    payload = parse_state(state)
    import win32com.client as win32  # type: ignore

    word = win32.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(str(dogovir_template.resolve()))
        bookmark_map = {
            "Number": "Number",
            "Data": "Data",
            "FIO": "FIO",
            "pasport": "pasport",
            "TaxNumber": "TaxNumber",
            "BirthDay": "BirthDay",
            "adres": "adres",
            "decl": "decl",
            "model": "model",
            "year": "year",
            "color": "color",
            "numberdv": "numberdv",
            "cuzov": "cuzov",
            "cub": "cub",
            "znak": "znak",
            "price": "price",
            "sumtext": "sumtext",
            "FIO2": "FIO2",
        }
        for key, bookmark_name in bookmark_map.items():
            if doc.Bookmarks.Exists(bookmark_name):
                doc.Bookmarks(bookmark_name).Range.Text = str(payload.get(key, ""))
        doc.SaveAs(str(out_path.resolve()))
        doc.Close(SaveChanges=False)
    finally:
        try:
            word.Quit()
        except Exception:
            pass
    return out_path


def save_text_preview(path: Path, title: str, body: str) -> Path:
    path.write_text(f"{title}\n\n{body}\n", encoding="utf-8")
    return path


def open_file_with_preference(path: Path, editor_path: str = "") -> None:
    if editor_path:
        subprocess.Popen([editor_path, str(path)])
        return
    if IS_WINDOWS:
        os.startfile(str(path))  # type: ignore[attr-defined]
    else:
        subprocess.Popen(["xdg-open", str(path)])


def slugify_case_name(text: str) -> str:
    raw = re.sub(r"\s+", "_", str(text).strip())
    raw = re.sub(r"[^0-9A-Za-zА-Яа-яІіЇїЄєҐґ_\-]", "", raw)
    return raw.strip("_") or "case"


def build_case_folder_name(state: Dict[str, str]) -> str:
    short_fio = short_name(state.get("C15", "")).replace(".", "")
    return slugify_case_name(short_fio)


def ensure_case_dir(base_out_dir: Path, state: Dict[str, str], unique: bool = False) -> Path:
    folder_name = build_case_folder_name(state)
    case_dir = base_out_dir / folder_name
    if unique and case_dir.exists():
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        case_dir = base_out_dir / f"{folder_name}_{stamp}"
    case_dir.mkdir(parents=True, exist_ok=True)
    return case_dir


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.base_dir = runtime_base_dir()
        self.out_dir = default_output_dir(self.base_dir)

        self.root.title("MotoCalc 6055")
        self.root.geometry("1280x900")
        self.root.minsize(1120, 760)

        self.source_path = tk.StringVar(value=str(self.base_dir / "6055.xls"))
        self.moto_path = tk.StringVar(value=str(self.base_dir / "6055_MOTO_template.xls"))
        self.dogovir_path = tk.StringVar(value=str(self.base_dir / "DOGOVIR_6055_template.doc"))
        self.vidatkova_path = tk.StringVar(value=str(self.base_dir / "vidatkova.xls"))
        self.editor_path = tk.StringVar(value="")
        self.open_after_save = tk.BooleanVar(value=True)

        self.state_vars: Dict[str, tk.StringVar] = {}
        self.widgets: Dict[str, tk.Entry] = {}
        self.summary_var = tk.StringVar(value="")
        self.error_var = tk.StringVar(value="")
        self.warning_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value="Готово")
        self.log_lines: list[str] = []
        self._syncing_form = False

        self._load_state_from_source()
        self._build_ui()
        self.refresh_validation()

    def _load_state_from_source(self) -> None:
        source = Path(self.source_path.get())
        state = load_form_state(source)
        for _, cell, _, _ in iter_field_specs():
            self.state_vars[cell] = tk.StringVar(value=state.get(cell, ""))
        self.state_vars["E15"].set(short_name(self.state_vars["C15"].get()))

    def _build_ui(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)
        self.root.configure(bg="#f4efe7")

        header = tk.Frame(self.root, bg="#1f2937", padx=18, pady=14)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)

        tk.Label(header, text="MotoCalc 6055", fg="white", bg="#1f2937", font=("Segoe UI", 20, "bold")).grid(row=0, column=0, sticky="w")
        tk.Label(header, text="Одна форма для акта, договору та видаткової", fg="#d1d5db", bg="#1f2937", font=("Segoe UI", 10)).grid(row=1, column=0, sticky="w", pady=(4, 0))

        main = ttk.Frame(self.root, padding=12)
        main.grid(row=1, column=0, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.rowconfigure(1, weight=1)

        actions = ttk.Frame(main)
        actions.grid(row=0, column=0, sticky="ew")
        for i in range(8):
            actions.columnconfigure(i, weight=1)

        ttk.Button(actions, text="Перевірити", command=self.refresh_validation).grid(row=0, column=0, sticky="ew", padx=4)
        ttk.Button(actions, text="Заповнити з 6055", command=self.reload_source).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(actions, text="Очистити форму", command=self.clear_form).grid(row=0, column=2, sticky="ew", padx=4)
        ttk.Button(actions, text="Акт", command=lambda: self.open_draft("act")).grid(row=0, column=3, sticky="ew", padx=4)
        ttk.Button(actions, text="Договір", command=lambda: self.open_draft("contract")).grid(row=0, column=4, sticky="ew", padx=4)
        ttk.Button(actions, text="Видаткова", command=lambda: self.open_draft("vidatkova")).grid(row=0, column=5, sticky="ew", padx=4)
        ttk.Button(actions, text="Зберегти 6055", command=self.save_source_changes).grid(row=0, column=6, sticky="ew", padx=4)
        ttk.Button(actions, text="Генерувати", command=self.generate_all).grid(row=0, column=7, sticky="ew", padx=4)

        summary = ttk.Frame(main)
        summary.grid(row=1, column=0, sticky="ew", pady=(10, 8))
        summary.columnconfigure(0, weight=1)
        summary.columnconfigure(1, weight=1)

        ttk.Label(summary, textvariable=self.summary_var, anchor="w", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, columnspan=2, sticky="ew")
        ttk.Label(summary, textvariable=self.error_var, foreground="#9f1239", anchor="w", wraplength=1100).grid(row=1, column=0, columnspan=2, sticky="ew", pady=(4, 0))
        ttk.Label(summary, textvariable=self.warning_var, foreground="#a16207", anchor="w", wraplength=1100).grid(row=2, column=0, columnspan=2, sticky="ew")

        self.notebook = ttk.Notebook(main)
        self.notebook.grid(row=2, column=0, sticky="nsew")

        data_tab = ttk.Frame(self.notebook)
        settings_tab = ttk.Frame(self.notebook)
        log_tab = ttk.Frame(self.notebook)
        self.notebook.add(data_tab, text="Дані")
        self.notebook.add(settings_tab, text="Налаштування")
        self.notebook.add(log_tab, text="Лог")

        self._build_data_tab(data_tab)
        self._build_settings_tab(settings_tab)
        self._build_log_tab(log_tab)

        self.write_log(f"Platform: {platform.system()}")
        self.write_log(f"Output folder: {self.out_dir}")
        if not IS_WINDOWS:
            self.write_log("Linux mode: Word COM write is unavailable; preview text will be generated.")
        ttk.Label(main, textvariable=self.status_var, anchor="w").grid(row=3, column=0, sticky="ew", pady=(8, 0))

    def _build_data_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)

        outer = tk.Frame(parent, bg="#f4efe7")
        outer.grid(row=0, column=0, sticky="nsew")
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(0, weight=1)

        canvas = tk.Canvas(outer, bg="#f4efe7", highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        content = tk.Frame(canvas, bg="#f4efe7")
        content.columnconfigure(0, weight=1)
        content.columnconfigure(1, weight=1)
        canvas.create_window((0, 0), window=content, anchor="nw")
        content.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure("all", width=e.width))

        title = tk.Frame(content, bg="#f4efe7", padx=10, pady=10)
        title.grid(row=0, column=0, columnspan=2, sticky="ew")
        tk.Label(title, text="АКТ / ГОЛОВНА ФОРМА 6055", font=("Segoe UI", 18, "bold"), bg="#f4efe7", fg="#111827").pack(anchor="w")
        tk.Label(title, text="Заповніть поля один раз, далі формуйте чорновик акту, договору або видаткової.", font=("Segoe UI", 10), bg="#f4efe7", fg="#4b5563").pack(anchor="w", pady=(4, 0))

        for index, (group_name, fields) in enumerate(FIELD_SECTIONS):
            row = index // 2 + 1
            col = index % 2
            card = tk.Frame(content, bg="#fffdf8", highlightbackground="#d6c7ad", highlightthickness=1, padx=14, pady=14)
            card.grid(row=row, column=col, sticky="nsew", padx=8, pady=8)
            content.grid_columnconfigure(col, weight=1)
            self._build_field_section(card, group_name, fields)

    def _build_field_section(self, parent, group_name: str, fields):
        tk.Label(parent, text=group_name, bg="#fffdf8", fg="#7c2d12", font=("Segoe UI", 12, "bold")).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        parent.columnconfigure(1, weight=1)
        for row, (cell, label, required) in enumerate(fields):
            row += 1
            display = f"{label} {'*' if required else ''} ({cell})"
            tk.Label(parent, text=display, bg="#fffdf8", fg="#374151", anchor="w").grid(row=row, column=0, sticky="w", padx=(0, 10), pady=4)

            var = self.state_vars[cell]
            entry = tk.Entry(parent, textvariable=var, relief="flat", highlightthickness=1, bd=0, bg="white", insertbackground="#111827")
            if cell == "E15":
                entry.configure(state="readonly", readonlybackground="#eef2ff", fg="#374151")
            entry.grid(row=row, column=1, sticky="ew", pady=4)
            self.widgets[cell] = entry

            if cell != "E15":
                var.trace_add("write", self._on_state_change)

        if {f[0] for f in fields} >= {"C39", "C42", "C43"}:
            ttk.Button(parent, text="Синхронізувати C39/C42/C43", command=self.sync_frame_fields).grid(row=len(fields) + 1, column=0, columnspan=2, sticky="w", pady=(10, 0))

    def _build_settings_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(1, weight=1)

        rows = [
            ("6055.xls", self.source_path, True),
            ("6055_MOTO_template.xls", self.moto_path, True),
            ("DOGOVIR_6055_template.doc", self.dogovir_path, True),
            ("vidatkova.xls", self.vidatkova_path, False),
            ("Output folder", tk.StringVar(value=str(self.out_dir)), False),
            ("Editor path", self.editor_path, False),
        ]

        self.output_dir_var = rows[4][1]

        for row, (label, var, editable) in enumerate(rows):
            ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=(0, 10), pady=5)
            entry = ttk.Entry(parent, textvariable=var)
            entry.grid(row=row, column=1, sticky="ew", pady=5)
            if label in {"6055.xls", "6055_MOTO_template.xls", "DOGOVIR_6055_template.doc", "vidatkova.xls", "Output folder", "Editor path"}:
                ttk.Button(parent, text="Огляд", command=lambda v=var: self.browse_path(v)).grid(row=row, column=2, padx=(8, 0))

        ttk.Checkbutton(parent, text="Відкривати файл після генерації", variable=self.open_after_save).grid(row=6, column=0, columnspan=2, sticky="w", pady=(12, 0))
        ttk.Button(parent, text="Перезавантажити 6055", command=self.reload_source).grid(row=7, column=0, sticky="w", pady=(12, 0))
        ttk.Button(parent, text="Зберегти 6055 зараз", command=self.save_source_changes).grid(row=7, column=1, sticky="w", pady=(12, 0))

    def _build_log_tab(self, parent: ttk.Frame) -> None:
        parent.rowconfigure(0, weight=1)
        parent.columnconfigure(0, weight=1)
        self.log = tk.Text(parent, height=12, wrap="word")
        self.log.grid(row=0, column=0, sticky="nsew")

    def browse_path(self, var: tk.StringVar) -> None:
        if var is self.output_dir_var:
            value = filedialog.askdirectory(initialdir=var.get() or str(self.out_dir)) if filedialog else ""
        else:
            value = filedialog.askopenfilename(initialdir=str(self.base_dir)) if filedialog else ""
        if value:
            var.set(value)

    def _on_state_change(self, *_):
        if self._syncing_form:
            return
        self._syncing_form = True
        self.state_vars["E15"].set(short_name(self.state_vars["C15"].get()))
        self.refresh_validation(silent=True)
        self._syncing_form = False

    def collect_state(self) -> Dict[str, str]:
        state = {cell: var.get().strip() for cell, var in self.state_vars.items()}
        state["E15"] = short_name(state.get("C15", ""))
        return state

    def refresh_validation(self, silent: bool = False):
        state = self.collect_state()
        payload, errors, warnings = validate_state(state)

        invalid_cells: set[str] = set()
        warning_cells: set[str] = set()

        for err in errors:
            if err.startswith("A3"):
                invalid_cells.add("A3")
            elif err.startswith("C46"):
                invalid_cells.add("C46")
            else:
                match = re.match(r"([A-Z]+\d+)\s", err)
                if match:
                    invalid_cells.add(match.group(1))

        for warn in warnings:
            if "C39 / C42 / C43" in warn:
                warning_cells.update({"C39", "C42", "C43"})
            if "Номер двигуна" in warn:
                warning_cells.add("C41")
            if "Колір" in warn:
                warning_cells.add("C28")
            if "Рік випуску" in warn:
                warning_cells.add("C29")

        for cell, entry in self.widgets.items():
            if cell == "E15":
                continue
            if cell in invalid_cells:
                entry.configure(bg="#fee2e2")
            elif cell in warning_cells:
                entry.configure(bg="#fef3c7")
            else:
                entry.configure(bg="#ffffff")

        self.summary_var.set(f"Заповнено: {sum(1 for v in state.values() if v)} полів. Номер: {payload.get('Number', '')} | Дата: {payload.get('Data', '')}")
        self.error_var.set("Помилки: " + ("; ".join(errors) if errors else "немає"))
        self.warning_var.set("Попередження: " + ("; ".join(warnings) if warnings else "немає"))

        if not silent:
            self.write_log(self.error_var.get())
            self.write_log(self.warning_var.get())

        return payload, errors, warnings

    def sync_frame_fields(self) -> None:
        state = self.collect_state()
        frame_value = state.get("C39") or state.get("C42") or state.get("C43")
        if frame_value:
            self._syncing_form = True
            for cell in ("C39", "C42", "C43"):
                self.state_vars[cell].set(frame_value)
            self._syncing_form = False
            self.write_log(f"Синхронізовано рамні поля: {frame_value}")
            self.refresh_validation(silent=True)

    def clear_form(self) -> None:
        self._syncing_form = True
        for cell in FORM_CELLS:
            if cell == "E15":
                self.state_vars[cell].set("")
            else:
                self.state_vars[cell].set("")
        self._syncing_form = False
        self.status_var.set("Форму очищено")
        self.write_log("Форму очищено")
        self.refresh_validation(silent=True)

    def reload_source(self) -> None:
        source = Path(self.source_path.get())
        state = load_form_state(source)
        self._syncing_form = True
        for cell in FORM_CELLS:
            self.state_vars[cell].set(state.get(cell, ""))
        self.state_vars["E15"].set(short_name(self.state_vars["C15"].get()))
        self._syncing_form = False
        self.write_log(f"Перезавантажено дані з {source}")
        self.status_var.set(f"Дані завантажено з {source.name}")
        self.refresh_validation(silent=True)

    def write_log(self, text: str) -> None:
        self.log_lines.append(text)
        if hasattr(self, "log"):
            self.log.insert("end", text + "\n")
            self.log.see("end")

    def save_source_changes(self) -> None:
        source = Path(self.source_path.get())
        state = self.collect_state()
        updates = {cell: state.get(cell, "") for cell in FORM_CELLS if cell != "E15"}
        updates["E15"] = short_name(state.get("C15", ""))
        updates["C39"] = state.get("C39", "")
        updates["C42"] = state.get("C42", "")
        updates["C43"] = state.get("C43", "")
        write_xls_cells(source, "Worksheet", updates, backup=True)
        self.write_log(f"Збережено 6055 у {source}")
        self.status_var.set(f"Збережено у {source.name}")

    def current_payload(self) -> Dict[str, str]:
        payload, errors, warnings = validate_state(self.collect_state())
        return payload

    def open_draft(self, kind: str) -> None:
        DraftWindow(self, kind)

    def _ensure_output_dir(self) -> Path:
        out_dir = Path(self.output_dir_var.get().strip() or str(self.out_dir))
        out_dir.mkdir(parents=True, exist_ok=True)
        return out_dir

    def save_draft(self, kind: str, open_after: bool = True, output_dir: Path | None = None) -> Path:
        state = self.collect_state()
        payload, errors, warnings = validate_state(state)
        if errors:
            raise ValueError("Потрібно виправити помилки перед збереженням: " + "; ".join(errors))

        out_dir = output_dir or self._ensure_output_dir()
        ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")

        if kind == "act":
            out_path = out_dir / f"6055_MOTO_filled_{ts}.xls"
            generate_act_xls_from_state(state, Path(self.moto_path.get()), out_path)
        elif kind == "contract":
            if IS_WINDOWS:
                out_path = out_dir / f"DOGOVIR_6055_filled_{ts}.doc"
                generate_contract_doc_windows_from_state(state, Path(self.dogovir_path.get()), out_path)
            else:
                out_path = out_dir / f"DOGOVIR_preview_{ts}.txt"
                save_text_preview(out_path, "ЧОРНОВИК ДОГОВОРУ", preview_text_for_contract(payload))
        elif kind == "vidatkova":
            out_path = out_dir / f"vidatkova_filled_{ts}.xls"
            generate_vidatkova_xls_from_state(state, Path(self.vidatkova_path.get()), out_path)
        else:
            raise ValueError(f"Невідомий тип чорновика: {kind}")

        self.write_log(f"Згенеровано: {out_path}")
        self.status_var.set(f"Створено {out_path.name}")
        if open_after and self.open_after_save.get():
            try:
                open_file_with_preference(out_path, self.editor_path.get().strip())
            except Exception as exc:
                self.write_log(f"Не вдалося відкрити файл: {exc}")
        return out_path

    def generate_all(self) -> None:
        try:
            self.save_source_changes()
            state = self.collect_state()
            base_out_dir = self._ensure_output_dir()
            case_dir = ensure_case_dir(base_out_dir, state, unique=True)
            self.save_draft("act", output_dir=case_dir)
            self.save_draft("contract", output_dir=case_dir)
            self.save_draft("vidatkova", output_dir=case_dir)
            self.status_var.set(f"Усі документи збережено у {case_dir}")
            self.write_log(f"Усі документи збережено у {case_dir}")
            if self.open_after_save.get():
                try:
                    open_file_with_preference(case_dir, self.editor_path.get().strip())
                except Exception as exc:
                    self.write_log(f"Не вдалося відкрити папку кейсу: {exc}")
        except Exception as exc:
            self.write_log(f"ERROR: {exc}")
            self.write_log(traceback.format_exc())
            if messagebox:
                messagebox.showerror("Error", str(exc))


if tk is not None:
    class DraftWindow(tk.Toplevel):
        def __init__(self, app: App, kind: str):
            super().__init__(app.root)
            self.app = app
            self.kind = kind
            self.title(self._title())
            self.geometry("860x720")
            self.minsize(720, 560)

            self.columnconfigure(0, weight=1)
            self.rowconfigure(1, weight=1)

            header = tk.Frame(self, bg="#0f766e", padx=14, pady=12)
            header.grid(row=0, column=0, sticky="ew")
            header.columnconfigure(0, weight=1)
            tk.Label(header, text=self._title(), fg="white", bg="#0f766e", font=("Segoe UI", 16, "bold")).grid(row=0, column=0, sticky="w")
            tk.Label(header, text=self._subtitle(), fg="#ccfbf1", bg="#0f766e", font=("Segoe UI", 10)).grid(row=1, column=0, sticky="w", pady=(2, 0))

            self.notice = tk.Label(self, text="", fg="#991b1b", anchor="w", justify="left", wraplength=820)
            self.notice.grid(row=1, column=0, sticky="ew", padx=14, pady=(10, 0))

            preview_host = tk.Frame(self, bg="#dbe4ea")
            preview_host.grid(row=2, column=0, sticky="nsew", padx=14, pady=10)
            preview_host.columnconfigure(0, weight=1)
            preview_host.rowconfigure(0, weight=1)

            self.preview_canvas = tk.Canvas(preview_host, bg="#dbe4ea", highlightthickness=0)
            preview_scroll = ttk.Scrollbar(preview_host, orient="vertical", command=self.preview_canvas.yview)
            self.preview_canvas.configure(yscrollcommand=preview_scroll.set)
            self.preview_canvas.grid(row=0, column=0, sticky="nsew")
            preview_scroll.grid(row=0, column=1, sticky="ns")

            self.preview_frame = tk.Frame(self.preview_canvas, bg="#dbe4ea")
            self.preview_window = self.preview_canvas.create_window((0, 0), window=self.preview_frame, anchor="nw")
            self.preview_frame.bind("<Configure>", lambda e: self.preview_canvas.configure(scrollregion=self.preview_canvas.bbox("all")))
            self.preview_canvas.bind("<Configure>", self._resize_preview_window)

            actions = ttk.Frame(self, padding=(14, 0, 14, 14))
            actions.grid(row=3, column=0, sticky="ew")
            actions.columnconfigure(0, weight=1)
            actions.columnconfigure(1, weight=1)
            actions.columnconfigure(2, weight=1)

            ttk.Button(actions, text="Зберегти", command=self.save).grid(row=0, column=0, sticky="ew", padx=4)
            ttk.Button(actions, text="Відкрити у редакторі", command=self.open_external).grid(row=0, column=1, sticky="ew", padx=4)
            ttk.Button(actions, text="Закрити", command=self.destroy).grid(row=0, column=2, sticky="ew", padx=4)

            self._refresh_preview()

        def _title(self) -> str:
            return {
                "act": "Чорновик акту 6055",
                "contract": "Чорновик договору",
                "vidatkova": "Чорновик видаткової",
            }[self.kind]

        def _subtitle(self) -> str:
            return {
                "act": "Це попередній перегляд перед генерацією фінального файлу.",
                "contract": "На Windows зберігає Word-документ, на інших системах створює текстовий чорновик.",
                "vidatkova": "Окрема видаткова накладна, заповнена з тієї ж форми даних.",
            }[self.kind]

        def _resize_preview_window(self, event) -> None:
            self.preview_canvas.itemconfigure(self.preview_window, width=event.width)

        def _render_preview_document(self, payload: Dict[str, str]) -> None:
            for child in self.preview_frame.winfo_children():
                child.destroy()

            paper = tk.Frame(self.preview_frame, bg="#fffdf8", highlightbackground="#cbd5e1", highlightthickness=1, padx=24, pady=22)
            paper.grid(row=0, column=0, sticky="ew", padx=20, pady=12)
            paper.columnconfigure(0, weight=1)
            paper.columnconfigure(1, weight=1)

            tk.Label(paper, text=self._title().upper(), font=("Segoe UI", 18, "bold"), bg="#fffdf8", fg="#111827").grid(row=0, column=0, columnspan=2, sticky="w")
            tk.Label(paper, text=self._subtitle(), font=("Segoe UI", 10), bg="#fffdf8", fg="#6b7280").grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 14))

            sections = preview_blocks(self.kind, payload)
            for index, (section_title, rows) in enumerate(sections):
                row = index // 2 + 2
                col = index % 2
                card = tk.Frame(paper, bg="#ffffff", highlightbackground="#e5e7eb", highlightthickness=1, padx=14, pady=12)
                card.grid(row=row, column=col, sticky="nsew", padx=8, pady=8)
                paper.grid_columnconfigure(col, weight=1)

                tk.Label(card, text=section_title, font=("Segoe UI", 11, "bold"), bg="#ffffff", fg="#7c2d12").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
                card.columnconfigure(1, weight=1)
                for item_index, (label, value) in enumerate(rows, start=1):
                    tk.Label(card, text=label, font=("Segoe UI", 9, "bold"), bg="#ffffff", fg="#374151").grid(row=item_index, column=0, sticky="nw", padx=(0, 10), pady=3)
                    value_box = tk.Label(card, text=value or "-", font=("Segoe UI", 10), bg="#f8fafc", fg="#111827", anchor="w", justify="left", wraplength=280, padx=8, pady=6)
                    value_box.grid(row=item_index, column=1, sticky="ew", pady=3)

            footer = tk.Frame(paper, bg="#fffdf8")
            footer.grid(row=10, column=0, columnspan=2, sticky="ew", pady=(16, 0))
            footer.columnconfigure(0, weight=1)
            footer.columnconfigure(1, weight=1)
            tk.Label(footer, text="Перевірте реквізити перед збереженням.", font=("Segoe UI", 9, "italic"), bg="#fffdf8", fg="#6b7280").grid(row=0, column=0, sticky="w")
            tk.Label(footer, text=f"Покупець: {payload.get('C15', payload.get('FIO', ''))}", font=("Segoe UI", 9), bg="#fffdf8", fg="#6b7280").grid(row=0, column=1, sticky="e")

        def _refresh_preview(self) -> None:
            payload, errors, warnings = validate_state(self.app.collect_state())
            self.notice.configure(text=("\n".join(errors + warnings) if errors or warnings else "Помилок і попереджень немає."))
            self._render_preview_document(payload)

        def save(self) -> None:
            try:
                path = self.app.save_draft(self.kind)
                self.app.write_log(f"Чорновик збережено: {path}")
                if messagebox:
                    messagebox.showinfo("Збережено", f"Файл збережено:\n{path}")
            except Exception as exc:
                if messagebox:
                    messagebox.showerror("Помилка", str(exc))
                self.app.write_log(f"ERROR: {exc}")

        def open_external(self) -> None:
            try:
                out_dir = self.app._ensure_output_dir()
                if self.kind == "act":
                    matches = sorted(out_dir.glob("6055_MOTO_filled_*.xls"))
                    file_name = matches[-1] if matches else None
                elif self.kind == "contract":
                    patterns = ["DOGOVIR_6055_filled_*.doc", "DOGOVIR_preview_*.txt"]
                    matches = []
                    for pattern in patterns:
                        matches.extend(sorted(out_dir.glob(pattern)))
                    file_name = matches[-1] if matches else None
                else:
                    matches = sorted(out_dir.glob("vidatkova_filled_*.xls"))
                    file_name = matches[-1] if matches else None
                if file_name:
                    open_file_with_preference(file_name, self.app.editor_path.get().strip())
            except Exception as exc:
                if messagebox:
                    messagebox.showerror("Помилка", str(exc))
else:
    class DraftWindow:
        def __init__(self, *args, **kwargs):
            raise RuntimeError("Tkinter is unavailable in this environment")


def run_demo(base_dir: Path) -> None:
    source = base_dir / "6055.xls"
    moto = base_dir / "6055_MOTO_template.xls"
    dog = base_dir / "DOGOVIR_6055_template.doc"

    updates = {
        "A3": "№ 8424/26/009999 від 15 травня 2026 року",
        "C15": "ТЕСТОВИЙ ІВАН ПЕТРОВИЧ",
        "E15": "ТЕСТОВИЙ І. П.",
        "C39": "TESTFRAME123456789",
        "C42": "TESTFRAME123456789",
        "C43": "TESTFRAME123456789",
        "D39": "",
        "D42": "",
        "D43": "",
    }
    write_xls_cells(source, "Worksheet", updates, backup=True)

    out_dir = base_dir / "out"
    out_dir.mkdir(parents=True, exist_ok=True)
    ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")

    out_act = out_dir / f"6055_MOTO_filled_{ts}.xls"
    transfer_to_act_non_windows(source, moto, out_act)

    if IS_WINDOWS:
        windows_transfer_all(source, moto, dog, out_dir)
    else:
        out_prev = out_dir / f"DOGOVIR_preview_{ts}.txt"
        transfer_to_word_preview(source, out_prev)

    print(f"Demo completed. Output folder: {out_dir}")


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="6055 transfer helper")
    p.add_argument("--demo", action="store_true", help="Run non-interactive demo update + transfer")
    return p.parse_args()


def main() -> int:
    args = parse_args()
    base_dir = runtime_base_dir()

    if args.demo:
        run_demo(base_dir)
        return 0

    if tk is None:
        print("Tkinter is unavailable in this environment.")
        return 1

    root = tk.Tk()
    App(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
