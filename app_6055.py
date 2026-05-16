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


def runtime_app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def runtime_resource_dir() -> Path:
    if getattr(sys, "frozen", False):
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            return Path(meipass).resolve()
    return Path(__file__).resolve().parent


def default_output_dir(base_dir: Path) -> Path:
    if IS_WINDOWS:
        return Path.home() / "Documents" / "JapanMoto" / "out"
    return base_dir / "out"


def safe_contract_output_path(out_path: Path) -> Path:
    return out_path.with_suffix(".txt")


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
CELL_LABELS: Dict[str, str] = {cell: label for _, cell, label, _ in iter_field_specs()}


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
    payload["pasport"] = payload.get("C17", "")
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
    invalid_cells: set[str] = set()
    warning_cells: set[str] = set()

    required = [cell for _, cell, _, req in iter_field_specs() if req]
    for cell in required:
        if not str(state.get(cell, "")).strip():
            label = CELL_LABELS.get(cell, cell)
            errors.append(f"«{label}» порожнє")
            invalid_cells.add(cell)

    if not payload["Number"] or not payload["Data"]:
        errors.append("Номер і дата документа (A3) не розпізнані")
        invalid_cells.add("A3")
    elif not re.fullmatch(r"\d{4}/\d{2}/\d{6}", payload["Number"]):
        warnings.append("Номер документа в A3 має незвичний формат")
        warning_cells.add("A3")

    if state.get("C51") and not valid_date_text(state.get("C51", "")):
        warnings.append("Дата набуття права — підозрілий формат")
        warning_cells.add("C51")

    if state.get("C12") and not valid_date_text(state.get("C12", "")):
        warnings.append("Дата народження — підозрілий формат")
        warning_cells.add("C12")

    frame_values = [str(state.get(c, "")).strip() for c in ("C39", "C42", "C43") if str(state.get(c, "")).strip()]
    if len(set(frame_values)) > 1:
        warnings.append("VIN / рама / шасі не збігаються")
        warning_cells.update({"C39", "C42", "C43"})

    if payload.get("cuzov"):
        vin_token = normalize_token(payload["cuzov"])
        if len(vin_token) < 8:
            errors.append("VIN / рама занадто короткий")
            invalid_cells.add("C39")
        elif len(vin_token) not in (8, 9, 10, 11, 12, 13, 14, 15, 16, 17):
            warnings.append("VIN / рама має незвичну довжину")
            warning_cells.add("C39")

    if payload.get("C41"):
        engine_token = normalize_token(payload["C41"])
        if len(engine_token) < 5:
            warnings.append("Номер двигуна схожий на короткий або неповний")
            warning_cells.add("C41")
        if payload.get("cuzov") and engine_token == normalize_token(payload["cuzov"]):
            warnings.append("Номер двигуна збігається з VIN / рамою, перевірте дані")
            warning_cells.add("C41")

    if payload.get("C28") and payload["C28"].strip().isdigit():
        warnings.append("Колір виглядає як число")
        warning_cells.add("C28")

    if payload.get("C18"):
        tax = re.sub(r"\D", "", payload["C18"])
        if len(tax) not in (8, 10):
            warnings.append("ІПН / код має незвичну довжину")
            warning_cells.add("C18")

    if payload.get("C29"):
        try:
            year = int(str(payload["C29"]).strip())
            if year < 1900 or year > dt.datetime.now().year + 1:
                warnings.append("Рік випуску поза нормальним діапазоном")
                warning_cells.add("C29")
        except Exception:
            warnings.append("Рік випуску не є числом")
            warning_cells.add("C29")

    if payload.get("C46"):
        try:
            float(str(payload["C46"]).replace(",", "."))
        except Exception:
            errors.append("Ціна з ПДВ — не є числом")
            invalid_cells.add("C46")

    price_no_vat = parse_decimal(payload.get("C44", ""))
    vat = parse_decimal(payload.get("C45", ""))
    price_total = parse_decimal(payload.get("C46", ""))
    if price_no_vat is not None and vat is not None and price_total is not None:
        if abs((price_no_vat + vat) - price_total) > 0.01:
            errors.append("Ціна без ПДВ + ПДВ ≠ Ціна з ПДВ")
            invalid_cells.update({"C44", "C45", "C46"})

    if price_no_vat is not None and price_total is not None and vat is not None and price_total > 0:
        vat_rate = round(vat / price_no_vat * 100, 2) if price_no_vat else None
        if price_no_vat and vat_rate is not None and vat_rate not in (20.0, 7.0, 0.0):
            warnings.append("Ставка ПДВ має незвичне значення")
            warning_cells.add("C45")

    payload["_invalid_cells"] = invalid_cells
    payload["_warning_cells"] = warning_cells
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


def generate_act_xls_from_state(state: Dict[str, str], source_6055: Path, out_path: Path) -> Path:
    shutil.copy2(source_6055, out_path)
    updates = {cell: state.get(cell, "") for cell in FORM_CELLS}
    full_fio = str(state.get("C15", "")).strip()
    short_fio = str(state.get("E15", "")).strip()
    if full_fio and short_fio:
        updates["C15"] = f"{full_fio} / {short_fio}"
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


def generate_contract_docx_fallback(state: Dict[str, str], out_path: Path) -> "Path | None":
    """Generate contract .docx without Word COM (python-docx fallback). Returns None if python-docx unavailable."""
    try:
        from docx import Document  # type: ignore
        from docx.shared import Cm  # type: ignore
        from docx.enum.text import WD_ALIGN_PARAGRAPH  # type: ignore
    except ImportError:
        return None

    payload = parse_state(state)
    doc = Document()
    for section in doc.sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)

    t = doc.add_heading("ДОГОВІР КУПІВЛІ-ПРОДАЖУ ТРАНСПОРТНОГО ЗАСОБУ", level=1)
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"\u2116 {payload.get('Number', '')}  \u0432\u0456\u0434  {payload.get('Data', '')}")
    run.bold = True

    def _add_field(label: str, value: str) -> None:
        para = doc.add_paragraph()
        rl = para.add_run(f"{label}: ")
        rl.bold = True
        para.add_run(value if value else "___")

    doc.add_heading("\u0421\u0442\u043e\u0440\u043e\u043d\u0438 \u0434\u043e\u0433\u043e\u0432\u043e\u0440\u0443", level=2)
    _add_field("\u041f\u043e\u043a\u0443\u043f\u0435\u0446\u044c (\u041f\u0406\u0411)", payload.get("FIO", ""))
    _add_field("\u0414\u0430\u0442\u0430 \u043d\u0430\u0440\u043e\u0434\u0436\u0435\u043d\u043d\u044f", payload.get("BirthDay", ""))
    _add_field("\u041f\u0430\u0441\u043f\u043e\u0440\u0442", payload.get("pasport", ""))
    _add_field("\u0406\u041f\u041d / \u0420\u041d\u041e\u041a\u041f\u041f", payload.get("TaxNumber", ""))
    _add_field("\u0410\u0434\u0440\u0435\u0441\u0430", payload.get("adres", ""))
    _add_field("\u041c\u0438\u0442\u043d\u0430 \u0434\u0435\u043a\u043b\u0430\u0440\u0430\u0446\u0456\u044f", payload.get("decl", ""))

    doc.add_heading("\u0422\u0440\u0430\u043d\u0441\u043f\u043e\u0440\u0442\u043d\u0438\u0439 \u0437\u0430\u0441\u0456\u0431", level=2)
    _add_field("\u041c\u0430\u0440\u043a\u0430, \u043c\u043e\u0434\u0435\u043b\u044c", payload.get("model", ""))
    _add_field("\u0420\u0456\u043a \u0432\u0438\u043f\u0443\u0441\u043a\u0443", payload.get("year", ""))
    _add_field("\u041a\u043e\u043b\u0456\u0440", payload.get("color", ""))
    _add_field("\u041d\u043e\u043c\u0435\u0440 \u0434\u0432\u0438\u0433\u0443\u043d\u0430", payload.get("numberdv", ""))
    _add_field("VIN / \u041d\u043e\u043c\u0435\u0440 \u0440\u0430\u043c\u0438", payload.get("cuzov", ""))
    _add_field("\u041e\u0431'\u0454\u043c \u0434\u0432\u0438\u0433\u0443\u043d\u0430", payload.get("cub", ""))
    _add_field("\u041d\u043e\u043c\u0435\u0440\u043d\u0456 \u0437\u043d\u0430\u043a\u0438", payload.get("znak", ""))

    doc.add_heading("\u0412\u0430\u0440\u0442\u0456\u0441\u0442\u044c", level=2)
    _add_field("\u0426\u0456\u043d\u0430 (\u0437 \u041f\u0414\u0412)", payload.get("price", ""))
    _add_field("\u0421\u0443\u043c\u0430 \u043f\u0440\u043e\u043f\u0438\u0441\u043e\u043c", payload.get("sumtext", ""))

    doc.add_paragraph()
    tbl = doc.add_table(rows=3, cols=2)
    tbl.cell(0, 0).text = "\u041f\u0440\u043e\u0434\u0430\u0432\u0435\u0446\u044c"
    tbl.cell(0, 1).text = "\u041f\u043e\u043a\u0443\u043f\u0435\u0446\u044c"
    tbl.cell(1, 0).text = "________________________"
    tbl.cell(1, 1).text = "________________________"
    tbl.cell(2, 0).text = ""
    tbl.cell(2, 1).text = payload.get("FIO2", "")

    out_docx = out_path.with_suffix(".docx")
    doc.save(str(out_docx))
    return out_docx


def generate_contract_doc_windows_from_state(state: Dict[str, str], dogovir_template: Path, out_path: Path) -> Path:
    payload = parse_state(state)
    try:
        import win32com.client as win32  # type: ignore

        word = win32.Dispatch("Word.Application")
        word.Visible = False
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
        return out_path
    except Exception:
        fb_docx = generate_contract_docx_fallback(state, out_path.with_suffix(".docx"))
        if fb_docx:
            return fb_docx
        fb_txt = out_path.with_suffix(".txt")
        save_text_preview(fb_txt, "ЧОРНОВИК ДОГОВОРУ", preview_text_for_contract(payload))
        return fb_txt
    finally:
        try:
            word.Quit()  # type: ignore[name-defined]
        except Exception:
            pass


def save_text_preview(path: Path, title: str, body: str) -> Path:
    path.write_text(f"{title}\n\n{body}\n", encoding="utf-8-sig")
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
        self.app_dir = runtime_app_dir()
        self.resource_dir = runtime_resource_dir()
        self.out_dir = default_output_dir(self.app_dir)

        self.root.title("Japan moto")
        self.root.geometry("1320x920")
        self.root.minsize(1160, 780)

        self.source_path = tk.StringVar(value=str(self.resource_dir / "6055.xls"))
        self.moto_path = tk.StringVar(value=str(self.resource_dir / "6055_MOTO_template.xls"))
        self.dogovir_path = tk.StringVar(value=str(self.resource_dir / "DOGOVIR_6055_template.doc"))
        self.vidatkova_path = tk.StringVar(value=str(self.resource_dir / "vidatkova.xls"))
        self.output_dir_var = tk.StringVar(value=str(self.out_dir))
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
        self._generation_suggested = False

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
        self.root.rowconfigure(2, weight=0)
        self.root.configure(bg="#f4efe7")

        header = tk.Frame(self.root, bg="#111827", padx=10, pady=8)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)
        header.columnconfigure(1, weight=0)
        header.columnconfigure(2, weight=0)

        title_frame = tk.Frame(header, bg="#111827")
        title_frame.grid(row=0, column=0, sticky="w")
        tk.Label(title_frame, text="Japan moto", fg="white", bg="#111827", font=("Segoe UI", 18, "bold")).pack(anchor="w")
        tk.Label(title_frame, text="Один екран для акта, договору та видаткової", fg="#cbd5e1", bg="#111827", font=("Segoe UI", 9)).pack(anchor="w")

        toolbar = tk.Frame(header, bg="#111827")
        toolbar.grid(row=0, column=1, sticky="e")

        self._make_toolbar_button(toolbar, "Акт", lambda: self.open_draft("act"), 0)
        self._make_toolbar_button(toolbar, "Договір", lambda: self.open_draft("contract"), 1)
        self._make_toolbar_button(toolbar, "Видаткова", lambda: self.open_draft("vidatkova"), 2)
        self._make_toolbar_button(toolbar, "↺ Відновити з файлу", self.reload_source, 3)
        self._make_toolbar_button(toolbar, "✖ Очистити", self.clear_form, 4)
        self._make_toolbar_button(toolbar, "⟲ Генерувати", self.generate_all, 5)

        gear_btn = tk.Button(
            header, text="⚙", command=self.open_settings_dialog,
            bg="#111827", fg="#9ca3af", activebackground="#1f2937",
            activeforeground="white", bd=0, relief="flat",
            font=("Segoe UI", 14), padx=8, pady=4, cursor="hand2",
        )
        gear_btn.grid(row=0, column=2, sticky="ne", padx=(0, 4), pady=2)

        main = ttk.Frame(self.root, padding=12)
        main.grid(row=1, column=0, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.rowconfigure(0, weight=1)
        self._build_data_tab(main)

        bottom = tk.Frame(self.root, bg="#111827", padx=12, pady=6)
        bottom.grid(row=2, column=0, sticky="ew")
        bottom.columnconfigure(0, weight=1)
        bottom.columnconfigure(1, weight=1)
        bottom.columnconfigure(2, weight=1)

        tk.Label(bottom, textvariable=self.status_var, bg="#111827", fg="#d1d5db", anchor="w").grid(row=0, column=0, sticky="ew")
        tk.Label(bottom, textvariable=self.error_var, bg="#111827", fg="#fecaca", anchor="w", wraplength=640).grid(row=0, column=1, sticky="ew")
        tk.Label(bottom, textvariable=self.warning_var, bg="#111827", fg="#fde68a", anchor="w", wraplength=640).grid(row=0, column=2, sticky="ew")

        self.write_log(f"Platform: {platform.system()}")
        self.write_log(f"Output folder: {self.out_dir}")
        if not IS_WINDOWS:
            self.write_log("Linux mode: Word COM write is unavailable; preview text will be generated.")

    def _make_toolbar_button(self, parent, text: str, command, column: int) -> None:
        button = tk.Button(parent, text=text, command=command, bg="#111827", fg="white", activebackground="#1f2937", activeforeground="white", bd=0, relief="flat", padx=10, pady=4, cursor="hand2")
        button.grid(row=0, column=column, padx=(0, 6), sticky="e")

    def _add_copy_paste_menu(self, entry: tk.Entry) -> None:
        menu = tk.Menu(entry, tearoff=0)
        menu.add_command(label="Вирізати", command=lambda: entry.event_generate("<<Cut>>"))
        menu.add_command(label="Копіювати", command=lambda: entry.event_generate("<<Copy>>"))
        menu.add_command(label="Вставити", command=lambda: entry.event_generate("<<Paste>>"))
        menu.add_separator()
        menu.add_command(label="Виділити все", command=lambda: (entry.select_range(0, "end"), entry.focus_set()))

        def _show_menu(event: "tk.Event") -> None:
            menu.tk_popup(event.x_root, event.y_root)

        def _select_all(event: "tk.Event") -> str:
            event.widget.select_range(0, "end")
            event.widget.icursor("end")
            return "break"

        entry.bind("<Button-3>", _show_menu)
        entry.bind("<Control-a>", _select_all)

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
        content_win = canvas.create_window((0, 0), window=content, anchor="nw")
        content.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(content_win, width=e.width))

        def _scroll(e):
            canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")

        def _scroll_up(e):
            canvas.yview_scroll(-1, "units")

        def _scroll_dn(e):
            canvas.yview_scroll(1, "units")

        canvas.bind_all("<MouseWheel>", _scroll)
        canvas.bind_all("<Button-4>", _scroll_up)
        canvas.bind_all("<Button-5>", _scroll_dn)

        title = tk.Frame(content, bg="#f4efe7", padx=10, pady=10)
        title.grid(row=0, column=0, columnspan=2, sticky="ew")
        tk.Label(title, text="АКТ / ГОЛОВНА ФОРМА 6055", font=("Segoe UI", 18, "bold"), bg="#f4efe7", fg="#111827").pack(anchor="w")
        tk.Label(title, text="Заповніть поля один раз, далі формуйте чорновик акту, договору або видаткової.", font=("Segoe UI", 10), bg="#f4efe7", fg="#4b5563").pack(anchor="w", pady=(4, 0))

        last_idx = len(FIELD_SECTIONS) - 1
        for index, (group_name, fields) in enumerate(FIELD_SECTIONS):
            row = index // 2 + 1
            col = index % 2
            card = tk.Frame(content, bg="#fffdf8", highlightbackground="#d6c7ad", highlightthickness=1, padx=14, pady=14)
            if index == last_idx and last_idx % 2 == 0:
                card.grid(row=row, column=0, columnspan=2, sticky="nsew", padx=8, pady=8)
            else:
                card.grid(row=row, column=col, sticky="nsew", padx=8, pady=8)
            content.grid_columnconfigure(col, weight=1)
            self._build_field_section(card, group_name, fields)

    def _build_field_section(self, parent, group_name: str, fields):
        tk.Label(parent, text=group_name, bg="#fffdf8", fg="#7c2d12", font=("Segoe UI", 12, "bold")).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        parent.columnconfigure(1, weight=1)
        for row, (cell, label, required) in enumerate(fields):
            row += 1
            display = f"{label}{' *' if required else ''}"
            tk.Label(parent, text=display, bg="#fffdf8", fg="#374151", anchor="w").grid(row=row, column=0, sticky="w", padx=(0, 10), pady=4)

            var = self.state_vars[cell]
            entry = tk.Entry(parent, textvariable=var, relief="flat", highlightthickness=1, bd=0, bg="white", insertbackground="#111827")
            if cell == "E15":
                entry.configure(state="readonly", readonlybackground="#eef2ff", fg="#374151")
            entry.grid(row=row, column=1, sticky="ew", pady=4)
            self.widgets[cell] = entry

            self._add_copy_paste_menu(entry)
            if cell != "E15":
                var.trace_add("write", self._on_state_change)

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
            value = filedialog.askopenfilename(initialdir=str(self.app_dir)) if filedialog else ""
        if value:
            var.set(value)

    def open_settings_dialog(self) -> None:
        dialog = tk.Toplevel(self.root)
        dialog.title("Шаблони і вихід")
        dialog.geometry("760x320")
        dialog.transient(self.root)
        dialog.grab_set()

        dialog.columnconfigure(1, weight=1)
        fields = [
            ("Витягнути інфо з файла", self.source_path),
            ("Шаблон акту", self.moto_path),
            ("Шаблон договору", self.dogovir_path),
            ("Шаблон видаткової", self.vidatkova_path),
            ("Папка збереження", self.output_dir_var),
            ("Редактор", self.editor_path),
        ]

        tk.Label(dialog, text="Параметри шаблонів", font=("Segoe UI", 14, "bold")).grid(row=0, column=0, columnspan=3, sticky="w", padx=16, pady=(14, 8))
        for row, (label, var) in enumerate(fields, start=1):
            tk.Label(dialog, text=label).grid(row=row, column=0, sticky="w", padx=16, pady=6)
            tk.Entry(dialog, textvariable=var).grid(row=row, column=1, sticky="ew", pady=6)
            tk.Button(dialog, text="Огляд", command=lambda v=var: self.browse_path(v)).grid(row=row, column=2, padx=12)

        tk.Checkbutton(dialog, text="Відкривати файл після генерації", variable=self.open_after_save).grid(row=7, column=0, columnspan=3, sticky="w", padx=16, pady=(10, 6))
        tk.Button(dialog, text="↺ Перезавантажити 6055", command=lambda: [self.reload_source(), dialog.destroy()]).grid(row=8, column=0, sticky="w", padx=16, pady=8)
        tk.Button(dialog, text="Зберегти 6055", command=lambda: [self.save_source_changes(), dialog.destroy()]).grid(row=8, column=1, sticky="w", pady=8)
        tk.Button(dialog, text="Закрити", command=dialog.destroy).grid(row=8, column=2, sticky="e", padx=12, pady=12)

    def _on_state_change(self, *_):
        if self._syncing_form:
            return
        self._syncing_form = True
        c15_val = self.state_vars["C15"].get()
        self.state_vars["E15"].set(short_name(c15_val))
        self.state_vars["D56"].set(c15_val)
        _payload, errors, _warnings = self.refresh_validation(silent=True)
        self._syncing_form = False
        if not errors and not self._generation_suggested:
            self._generation_suggested = True
            if messagebox and messagebox.askyesno(
                "Готово до генерації",
                "Всі обов'язкові поля заповнено!\n\nСгенерувати всі документи зараз?",
            ):
                self.generate_all()

    def collect_state(self) -> Dict[str, str]:
        state = {cell: var.get().strip() for cell, var in self.state_vars.items()}
        state["E15"] = short_name(state.get("C15", ""))
        return state

    def refresh_validation(self, silent: bool = False):
        state = self.collect_state()
        payload, errors, warnings = validate_state(state)

        invalid_cells: set[str] = payload.pop("_invalid_cells", set())
        warning_cells: set[str] = payload.pop("_warning_cells", set())

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
            self.state_vars[cell].set("")
        self._syncing_form = False
        self._generation_suggested = False
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

    def save_draft(self, kind: str, open_after: bool = True, output_dir: Path | None = None, use_timestamp: bool = True) -> Path:
        state = self.collect_state()
        payload, errors, warnings = validate_state(state)
        if errors:
            raise ValueError("Потрібно виправити помилки перед збереженням: " + "; ".join(errors))

        out_dir = output_dir or self._ensure_output_dir()
        ts = f"_{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}" if use_timestamp else ""
        doc_num = re.sub(r'[\\/:*?"<>|]', "-", payload.get("Number", "")).strip("-")
        num_part = f" \u2116{doc_num}" if doc_num else ""

        if kind == "act":
            out_path = out_dir / f"Акт{num_part}{ts}.xls"
            generate_act_xls_from_state(state, Path(self.source_path.get()), out_path)
        elif kind == "contract":
            if IS_WINDOWS:
                out_path = out_dir / f"Договір{num_part}{ts}.doc"
                out_path = generate_contract_doc_windows_from_state(state, Path(self.dogovir_path.get()), out_path)
            else:
                out_path = out_dir / f"Договір{num_part}{ts}.docx"
                result = generate_contract_docx_fallback(state, out_path)
                if result:
                    out_path = result
                else:
                    out_path = out_path.with_suffix(".txt")
                    save_text_preview(out_path, "ЧОРНОВИК ДОГОВОРУ", preview_text_for_contract(payload))
        elif kind == "vidatkova":
            out_path = out_dir / f"видаткова{num_part}{ts}.xls"
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
            state = self.collect_state()
            payload, errors, _ = validate_state(state)
            if errors:
                messagebox.showerror("Помилки валідації", "\n".join(errors))
                return
            base_out_dir = self._ensure_output_dir()
            case_dir = base_out_dir / build_case_folder_name(state)
            if case_dir.exists():
                if not messagebox.askyesno(
                    "Перезаписати?",
                    f"Папка вже існує:\n{case_dir}\n\nПерезаписати всі документи?",
                ):
                    return
            case_dir.mkdir(parents=True, exist_ok=True)
            self.save_draft("act", open_after=False, output_dir=case_dir, use_timestamp=False)
            self.save_draft("contract", open_after=False, output_dir=case_dir, use_timestamp=False)
            self.save_draft("vidatkova", open_after=False, output_dir=case_dir, use_timestamp=False)
            self.status_var.set(f"Усі документи збережено у {case_dir.name}")
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
                messagebox.showerror("Помилка", str(exc))


if tk is not None:
    class DraftWindow(tk.Toplevel):
        _DEST_STEMS = {"act": "6055_akt", "contract": "dogovir", "vidatkova": "vidatkova"}

        def __init__(self, app: App, kind: str):
            super().__init__(app.root)
            self.app = app
            self.kind = kind
            self._temp_path: "Path | None" = None
            self.title(self._title())
            self.transient(app.root)

            self.columnconfigure(0, weight=1)

            header = tk.Frame(self, bg="#0f766e", padx=14, pady=12)
            header.grid(row=0, column=0, sticky="ew")
            header.columnconfigure(0, weight=1)
            tk.Label(header, text=self._title(), fg="white", bg="#0f766e", font=("Segoe UI", 15, "bold")).grid(row=0, column=0, sticky="w")
            tk.Label(header, text="Файл відкрито у системній програмі. Перевірте і збережіть.", fg="#ccfbf1", bg="#0f766e", font=("Segoe UI", 9)).grid(row=1, column=0, sticky="w")

            body = tk.Frame(self, bg="#f4efe7", padx=16, pady=14)
            body.grid(row=1, column=0, sticky="ew")
            body.columnconfigure(0, weight=1)
            self.status_lbl = tk.Label(body, text="Генерація файлу...", bg="#f4efe7", fg="#374151", anchor="w", wraplength=600)
            self.status_lbl.grid(row=0, column=0, sticky="ew")
            self.path_lbl = tk.Label(body, text="", bg="#f4efe7", fg="#6b7280", anchor="w", wraplength=600, font=("Segoe UI", 9))
            self.path_lbl.grid(row=1, column=0, sticky="ew", pady=(4, 0))

            actions = tk.Frame(self, bg="#f4efe7", padx=16, pady=12)
            actions.grid(row=2, column=0, sticky="ew")
            ttk.Button(actions, text="Зберегти в папку клієнта", command=self.save_to_case).pack(side="left", padx=(0, 8))
            ttk.Button(actions, text="Відкрити знову", command=self._open_in_app).pack(side="left", padx=(0, 8))
            ttk.Button(actions, text="Закрити", command=self.destroy).pack(side="left")

            self.update_idletasks()
            _w = 640
            _h = self.winfo_reqheight()
            _px = max(0, app.root.winfo_rootx() + (app.root.winfo_width() - _w) // 2)
            _py = max(0, app.root.winfo_rooty() + (app.root.winfo_height() - _h) // 2)
            self.geometry(f"{_w}x{_h}+{_px}+{_py}")
            self.resizable(False, False)
            self.after(100, self._generate_and_open)

        def _title(self) -> str:
            return {"act": "Акт 6055", "contract": "Договір", "vidatkova": "Видаткова"}[self.kind]

        def _generate_and_open(self) -> None:
            import tempfile
            try:
                tmp_dir = Path(tempfile.mkdtemp(prefix="japan_moto_"))
                self._temp_path = self.app.save_draft(self.kind, open_after=False, output_dir=tmp_dir)
                self.status_lbl.configure(text="Файл відкрито. Перевірте документ у системній програмі, потім збережіть або закрийте.")
                self.path_lbl.configure(text=f"Тимчасовий файл: {self._temp_path}")
                self._open_in_app()
            except Exception as exc:
                self.status_lbl.configure(text=f"Помилка генерації: {exc}", fg="#991b1b")
                self.app.write_log(f"DraftWindow error: {exc}")

        def _open_in_app(self) -> None:
            if self._temp_path and self._temp_path.exists():
                try:
                    open_file_with_preference(self._temp_path, self.app.editor_path.get().strip())
                except Exception as exc:
                    messagebox.showerror("Помилка", str(exc), parent=self)
            else:
                messagebox.showwarning("Немає файлу", "Файл не знайдено. Зачекайте генерацію.", parent=self)

        def save_to_case(self) -> None:
            if not self._temp_path or not self._temp_path.exists():
                messagebox.showwarning("Немає файлу", "Спочатку дочекайтесь генерації.", parent=self)
                return
            state = self.app.collect_state()
            base_out_dir = self.app._ensure_output_dir()
            case_dir = base_out_dir / build_case_folder_name(state)
            payload_s = parse_state(state)
            doc_num_s = re.sub(r'[\\/:*?"<>|]', "-", payload_s.get("Number", "")).strip("-")
            num_part_s = f" №{doc_num_s}" if doc_num_s else ""
            _stem_map = {"act": "Акт", "contract": "Договір", "vidatkova": "видаткова"}
            dest_name = _stem_map.get(self.kind, self._temp_path.stem) + num_part_s + self._temp_path.suffix
            dest = case_dir / dest_name
            if dest.exists():
                if not messagebox.askyesno("Перезаписати?", f"Файл вже існує:\n{dest}\n\nПерезаписати?", parent=self):
                    return
            case_dir.mkdir(parents=True, exist_ok=True)
            shutil.copy2(self._temp_path, dest)
            messagebox.showinfo("Збережено", f"Файл збережено:\n{dest}", parent=self)
            self.app.write_log(f"Збережено: {dest}")
            self.app.status_var.set(f"Збережено {dest.name}")
else:
    class DraftWindow:
        def __init__(self, *args, **kwargs):
            raise RuntimeError("Tkinter is unavailable in this environment")


def run_demo(app_dir: Path, resource_dir: Path) -> None:
    source = app_dir / "6055.xls"
    moto = resource_dir / "6055_MOTO_template.xls"
    dog = resource_dir / "DOGOVIR_6055_template.doc"

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

    out_dir = app_dir / "out"
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
    app_dir = runtime_app_dir()
    resource_dir = runtime_resource_dir()

    if args.demo:
        run_demo(app_dir, resource_dir)
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
