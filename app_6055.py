#!/usr/bin/env python3
import argparse
import datetime as dt
import os
import platform
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


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("6055 Transfer Tool")
        self.base_dir = runtime_base_dir()
        self.out_dir = default_output_dir(self.base_dir)

        self.source_path = tk.StringVar(value=str(self.base_dir / "6055.xls"))
        self.moto_path = tk.StringVar(value=str(self.base_dir / "6055_MOTO_template.xls"))
        self.dogovir_path = tk.StringVar(value=str(self.base_dir / "DOGOVIR_6055_template.doc"))

        self.a3_text = tk.StringVar(value=str(read_xls_cell(Path(self.source_path.get()), "Worksheet", "A3")))
        self.fio_text = tk.StringVar(value=str(read_xls_cell(Path(self.source_path.get()), "Worksheet", "C15")))
        self.frame_text = tk.StringVar(value=str(read_xls_cell(Path(self.source_path.get()), "Worksheet", "C43")))

        self._build_ui()

    def _build_ui(self) -> None:
        frm = ttk.Frame(self.root, padding=12)
        frm.grid(sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        for i in range(2):
            frm.columnconfigure(i, weight=1)

        ttk.Label(frm, text="Source 6055.xls").grid(row=0, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.source_path).grid(row=0, column=1, sticky="ew", padx=6)

        ttk.Label(frm, text="MOTO template").grid(row=1, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.moto_path).grid(row=1, column=1, sticky="ew", padx=6)

        ttk.Label(frm, text="DOGOVIR template").grid(row=2, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.dogovir_path).grid(row=2, column=1, sticky="ew", padx=6)

        ttk.Separator(frm, orient="horizontal").grid(row=3, column=0, columnspan=2, sticky="ew", pady=10)

        ttk.Label(frm, text="A3 (номер + дата)").grid(row=4, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.a3_text).grid(row=4, column=1, sticky="ew", padx=6)

        ttk.Label(frm, text="C15 (ПІБ)").grid(row=5, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.fio_text).grid(row=5, column=1, sticky="ew", padx=6)

        ttk.Label(frm, text="C43 (номер рами)").grid(row=6, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.frame_text).grid(row=6, column=1, sticky="ew", padx=6)

        btns = ttk.Frame(frm)
        btns.grid(row=7, column=0, columnspan=2, sticky="ew", pady=10)
        for i in range(4):
            btns.columnconfigure(i, weight=1)

        ttk.Button(btns, text="1) Зберегти зміни в 6055", command=self.save_source_changes).grid(row=0, column=0, sticky="ew", padx=4)
        ttk.Button(btns, text="2) Перенести в акт", command=self.transfer_act).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(btns, text="3) Перенести в договір", command=self.transfer_dogovir).grid(row=0, column=2, sticky="ew", padx=4)
        ttk.Button(btns, text="4) Зробити все", command=self.transfer_all).grid(row=0, column=3, sticky="ew", padx=4)

        self.log = tk.Text(frm, height=12)
        self.log.grid(row=8, column=0, columnspan=2, sticky="nsew")
        frm.rowconfigure(8, weight=1)

        self.write_log(f"Platform: {platform.system()}")
        self.write_log(f"Output folder: {self.out_dir}")
        if not IS_WINDOWS:
            self.write_log("Linux mode: Word COM write is unavailable, preview file will be created.")

    def write_log(self, text: str) -> None:
        self.log.insert("end", text + "\n")
        self.log.see("end")

    def save_source_changes(self) -> None:
        source = Path(self.source_path.get())
        updates = {
            "A3": self.a3_text.get(),
            "C15": self.fio_text.get(),
            "E15": short_name(self.fio_text.get()),
            # Keep VIN/chassis/frame synchronized for this workflow.
            "C39": self.frame_text.get(),
            "C42": self.frame_text.get(),
            "C43": self.frame_text.get(),
            "D39": "",
            "D42": "",
            "D43": "",
        }
        write_xls_cells(source, "Worksheet", updates, backup=True)
        self.write_log(f"Saved source updates into: {source}")

    def transfer_act(self) -> None:
        source = Path(self.source_path.get())
        moto = Path(self.moto_path.get())
        out_dir = self.out_dir
        out_dir.mkdir(parents=True, exist_ok=True)
        ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        out_act = out_dir / f"6055_MOTO_filled_{ts}.xls"

        if IS_WINDOWS:
            self.write_log("On Windows, use 'Зробити все' to run COM transfer with Word.")
            transfer_to_act_non_windows(source, moto, out_act)
        else:
            transfer_to_act_non_windows(source, moto, out_act)

        self.write_log(f"ACT created: {out_act}")

    def transfer_dogovir(self) -> None:
        source = Path(self.source_path.get())
        out_dir = self.out_dir
        out_dir.mkdir(parents=True, exist_ok=True)
        ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")

        if IS_WINDOWS:
            self.write_log("Use 'Зробити все' for full COM pipeline on Windows.")
        else:
            out_preview = out_dir / f"DOGOVIR_preview_{ts}.txt"
            transfer_to_word_preview(source, out_preview)
            self.write_log(f"DOGOVIR preview created: {out_preview}")

    def transfer_all(self) -> None:
        try:
            self.save_source_changes()
            source = Path(self.source_path.get())
            moto = Path(self.moto_path.get())
            dog = Path(self.dogovir_path.get())
            out_dir = self.out_dir
            out_dir.mkdir(parents=True, exist_ok=True)

            if IS_WINDOWS:
                windows_transfer_all(source, moto, dog, out_dir)
                self.write_log("Windows COM transfer complete (ACT + DOGOVIR).")
            else:
                ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
                out_act = out_dir / f"6055_MOTO_filled_{ts}.xls"
                out_prev = out_dir / f"DOGOVIR_preview_{ts}.txt"
                transfer_to_act_non_windows(source, moto, out_act)
                transfer_to_word_preview(source, out_prev)
                self.write_log(f"ACT created: {out_act}")
                self.write_log(f"DOGOVIR preview created: {out_prev}")
        except Exception as exc:
            self.write_log(f"ERROR: {exc}")
            self.write_log(traceback.format_exc())
            if messagebox:
                messagebox.showerror("Error", str(exc))


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
