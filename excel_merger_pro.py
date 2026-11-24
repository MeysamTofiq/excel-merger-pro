import os
import re
import sys
import copy

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string

from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QColor, QFont
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFrame, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit, QFileDialog, QMessageBox,
    QProgressBar, QCheckBox, QButtonGroup, QListWidget, QGraphicsDropShadowEffect,
)
from PyQt6.QtGui import QIcon

def resource_path(relative):
    import sys
    base = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base, relative)

MODE_SHEETS = "sheets"
MODE_STACK = "stack"
MODE_ONE_SHEET = "one_sheet"

MODES = [
    (MODE_SHEETS, "🗂  هر فایل ← یک شیت جدا"),
    (MODE_STACK, "⬇️  ادغام فایل‌ها زیر هم"),
    (MODE_ONE_SHEET, "🧩  شیت‌های یک فایل ← یک شیت"),
]

EXCEL_EXT = (".xlsx", ".xls")
OUTPUT_NAMES = {"merged_excels.xlsx", "merged_stacked.xlsx"}


# ======================= توابع کمکی اکسل =======================

def copy_cell_style(src, dst):
    if src.has_style:
        dst.font = copy.copy(src.font)
        dst.border = copy.copy(src.border)
        dst.fill = copy.copy(src.fill)
        dst.number_format = src.number_format
        dst.protection = copy.copy(src.protection)
        dst.alignment = copy.copy(src.alignment)


def copy_sheet_with_styles(src_ws, dst_ws):
    """کپی کامل سلول‌ها همراه با استایل، سلول‌های ادغام‌شده و ابعاد"""
    for row in src_ws.iter_rows():
        for cell in row:
            dst = dst_ws[cell.coordinate]
            dst.value = cell.value
            copy_cell_style(cell, dst)
    for rng in src_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(rng))
    for letter, dim in src_ws.column_dimensions.items():
        if dim.width:
            dst_ws.column_dimensions[letter].width = dim.width
    for idx, dim in src_ws.row_dimensions.items():
        if dim.height:
            dst_ws.row_dimensions[idx].height = dim.height


def unique_sheet_name(wb, name):
    """نام شیت یکتا، بدون کاراکترهای ممنوع و حداکثر ۳۱ کاراکتر"""
    name = re.sub(r'[\[\]:*?/\\]', "_", name or "Sheet")[:31] or "Sheet"
    candidate, i = name, 1
    while candidate in wb.sheetnames:
        suffix = f"_{i}"
        candidate = name[:31 - len(suffix)] + suffix
        i += 1
    return candidate


def write_dataframe(ws, df):
    ws.append([str(c) for c in df.columns])
    for row in df.itertuples(index=False):
        ws.append([None if pd.isna(v) else v for v in row])


def load_sources(path):
    """لیست (نام شیت, worksheet) و ورک‌بوک (برای بستن). برای .xls داده‌ها به ورک‌بوک موقت منتقل می‌شوند."""
    if path.lower().endswith(".xls"):
        frames = pd.read_excel(path, engine="xlrd", sheet_name=None)
        tmp = Workbook()
        tmp.remove(tmp.active)
        sources = []
        for name, df in frames.items():
            ws = tmp.create_sheet(unique_sheet_name(tmp, name))
            write_dataframe(ws, df)
            sources.append((name, ws))
        return sources, None
    wb = load_workbook(path)
    return [(ws.title, ws) for ws in wb.worksheets], wb


def read_sheet_names(path):
    if path.lower().endswith(".xls"):
        return pd.ExcelFile(path, engine="xlrd").sheet_names
    wb = load_workbook(path, read_only=True)
    names = wb.sheetnames
    wb.close()
    return names


def list_excel_files(folder):
    files = []
    for f in sorted(os.listdir(folder)):
        if f.startswith("~$") or f in OUTPUT_NAMES or f.endswith("_one_sheet.xlsx"):
            continue
        if f.lower().endswith(EXCEL_EXT):
            files.append(os.path.join(folder, f))
    return files


def append_rows(src_ws, out_ws, state, label, add_source):
    """ردیف‌های یک شیت را زیر ردیف‌های قبلی اضافه می‌کند و سرستون تکراری را حذف می‌کند"""
    rows = list(src_ws.iter_rows())
    if not rows:
        return
    offset = 1 if add_source else 0

    head = [c.value for c in rows[0]]
    start = 0
    if state["header"] is None:
        state["header"] = head
    elif head == state["header"]:
        start = 1

    if not state["widths_done"]:
        for letter, dim in src_ws.column_dimensions.items():
            if dim.width:
                idx = column_index_from_string(letter) + offset
                out_ws.column_dimensions[get_column_letter(idx)].width = dim.width
        if add_source:
            out_ws.column_dimensions["A"].width = 24
        state["widths_done"] = True

    for row in rows[start:]:
        if all(c.value is None for c in row):
            continue
        r = state["next_row"]
        if add_source:
            out_ws.cell(row=r, column=1, value="منبع" if r == 1 else label)
        for c in row:
            dst = out_ws.cell(row=r, column=c.column + offset, value=c.value)
            copy_cell_style(c, dst)
        state["next_row"] += 1


def new_stack_state():
    return {"header": None, "next_row": 1, "widths_done": False}


# ======================= Worker (Thread جداگانه) =======================

class MergeWorker(QThread):
    log = pyqtSignal(str)
    progress = pyqtSignal(int)
    done = pyqtSignal(bool, str, str)

    def __init__(self, mode, files, output_path, find_text, replace_text, add_source):
        super().__init__()
        self.mode = mode
        self.files = files
        self.output_path = output_path
        self.find_text = find_text
        self.replace_text = replace_text
        self.add_source = add_source

    def rename(self, text):
        if self.find_text:
            text = text.replace(self.find_text, self.replace_text)
        return text

    def run(self):
        try:
            if self.mode == MODE_SHEETS:
                wb, ok = self.merge_as_sheets()
            elif self.mode == MODE_STACK:
                wb, ok = self.stack_files()
            else:
                wb, ok = self.one_sheet()

            if ok == 0:
                self.done.emit(False, "هیچ فایلی با موفقیت ادغام نشد.", "")
                return

            wb.save(self.output_path)
            self.progress.emit(100)
            self.log.emit(f"\n🎯 {ok} مورد با موفقیت ادغام شد → {self.output_path}")
            self.done.emit(True, f"{ok} مورد ترکیب شد.", self.output_path)
        except Exception as e:
            self.log.emit(f"❌ خطای کلی: {e}")
            self.done.emit(False, str(e), "")

    def append_file(self, path, out_ws, state, by_sheet_name):
        sources, wb = load_sources(path)
        stem = os.path.splitext(os.path.basename(path))[0]
        try:
            for sheet_name, ws in sources:
                if by_sheet_name:
                    label = sheet_name
                elif len(sources) == 1:
                    label = stem
                else:
                    label = f"{stem} | {sheet_name}"
                append_rows(ws, out_ws, state, label, self.add_source)
        finally:
            if wb is not None:
                wb.close()

    # حالت ۱: هر فایل یک شیت جدا
    def merge_as_sheets(self):
        out_wb = Workbook()
        out_wb.remove(out_wb.active)
        ok = 0
        total = len(self.files)
        for i, path in enumerate(self.files, 1):
            name = os.path.basename(path)
            try:
                sources, wb = load_sources(path)
                stem = os.path.splitext(name)[0]
                if len(sources) > 1:
                    for sheet_name, ws in sources:
                        target = out_wb.create_sheet(unique_sheet_name(out_wb, self.rename(sheet_name)))
                        copy_sheet_with_styles(ws, target)
                else:
                    target = out_wb.create_sheet(unique_sheet_name(out_wb, self.rename(stem)))
                    copy_sheet_with_styles(sources[0][1], target)
                if wb is not None:
                    wb.close()
                ok += 1
                self.log.emit(f"✅ {name} اضافه شد")
            except Exception as e:
                self.log.emit(f"⚠️ خطا در '{name}': {e}")
            self.progress.emit(int(i * 100 / total))
        return out_wb, ok

    # حالت ۲: ادغام فایل‌ها زیر هم
    def stack_files(self):
        out_wb = Workbook()
        out_ws = out_wb.active
        out_ws.title = "Merged"
        state = new_stack_state()
        ok = 0
        total = len(self.files)
        for i, path in enumerate(self.files, 1):
            name = os.path.basename(path)
            try:
                self.append_file(path, out_ws, state, by_sheet_name=False)
                ok += 1
                self.log.emit(f"✅ {name} زیر هم اضافه شد")
            except Exception as e:
                self.log.emit(f"⚠️ خطا در '{name}': {e}")
            self.progress.emit(int(i * 100 / total))
        return out_wb, ok

    # حالت ۳: شیت‌های یک فایل در یک شیت
    def one_sheet(self):
        out_wb = Workbook()
        out_ws = out_wb.active
        stem = os.path.splitext(os.path.basename(self.files[0]))[0]
        out_ws.title = unique_sheet_name(out_wb, stem)
        state = new_stack_state()
        self.append_file(self.files[0], out_ws, state, by_sheet_name=True)
        self.log.emit(f"✅ همه شیت‌های '{os.path.basename(self.files[0])}' در یک شیت ادغام شد")
        return out_wb, 1


# ======================= کارت و استایل =======================

class Card(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("card")
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(34)
        shadow.setOffset(0, 10)
        shadow.setColor(QColor(0, 0, 0, 170))
        self.setGraphicsEffect(shadow)


def section_label(text):
    lbl = QLabel(text)
    lbl.setObjectName("section")
    return lbl


QSS = """
QWidget#root { background: #0b0d14; }
QFrame#card { background: #151927; border: 1px solid #1f2540; border-radius: 20px; }
QLabel { color: #e4e7f3; background: transparent; font-size: 13px; }
QLabel#title { color: #ffffff; font-size: 22px; font-weight: bold; }
QLabel#muted { color: #7b8199; font-size: 12px; }
QLabel#section { color: #c8ccff; font-size: 14px; font-weight: bold; }
QLabel#badge { color: #b9a6ff; background: #241f45; border-radius: 11px; padding: 4px 12px; font-size: 12px; }
QLineEdit { background: #0e1220; color: #e4e7f3; border: 1px solid #262d45; border-radius: 12px; padding: 9px 12px; }
QLineEdit:focus { border: 1px solid #7c5cff; }
QPushButton#mode { background: #0e1220; color: #9aa1bd; border: 1px solid #232942; border-radius: 14px; padding: 13px 14px; font-size: 13px; }
QPushButton#mode:hover { border: 1px solid #3b3f66; color: #e4e7f3; }
QPushButton#mode:checked { color: #ffffff; border: 1px solid #7c5cff; background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #2a2166, stop:1 #3a2a7a); }
QPushButton#ghost { background: #1d2340; color: #c8ccff; border: none; border-radius: 12px; padding: 9px 14px; }
QPushButton#ghost:hover { background: #262e52; }
QPushButton#primary { background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #7c5cff, stop:1 #b86bff); color: #ffffff; border: none; border-radius: 14px; padding: 13px; font-size: 14px; font-weight: bold; }
QPushButton#primary:hover { background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #8f73ff, stop:1 #c983ff); }
QPushButton#primary:disabled { background: #262b40; color: #6b7189; }
QCheckBox { color: #c8ccff; spacing: 8px; font-size: 13px; }
QCheckBox::indicator { width: 18px; height: 18px; border-radius: 6px; border: 1px solid #3a4266; background: #0e1220; }
QCheckBox::indicator:checked { background: #7c5cff; border: 1px solid #7c5cff; }
QListWidget#files { background: #0e1220; color: #dfe3f5; border: none; border-radius: 14px; padding: 8px; font-size: 13px; }
QListWidget#files::item { padding: 7px 10px; border-radius: 8px; }
QListWidget#files::item:selected { background: #2a2166; }
QTextEdit#log { background: #0e1220; color: #b7bdd8; border: none; border-radius: 14px; padding: 10px; font-family: Consolas; font-size: 12px; }
QProgressBar { background: #0e1220; border: none; border-radius: 5px; }
QProgressBar::chunk { border-radius: 5px; background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #7c5cff, stop:1 #c983ff); }
QScrollBar:vertical { background: transparent; width: 8px; margin: 2px; }
QScrollBar::handle:vertical { background: #2a3150; border-radius: 4px; min-height: 30px; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QMessageBox { background: #151927; }
QMessageBox QLabel { color: #e4e7f3; }
QMessageBox QPushButton { background: #7c5cff; color: #ffffff; border: none; border-radius: 10px; padding: 8px 18px; min-width: 80px; }
"""


# ======================= پنجره اصلی =======================

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Merger Pro")
        self.resize(1040, 700)

        self.mode = MODE_SHEETS
        self.source_dir = None
        self.files = []
        self.worker = None

        root = QWidget()
        root.setObjectName("root")
        self.setCentralWidget(root)
        main = QVBoxLayout(root)
        main.setContentsMargins(20, 20, 20, 20)
        main.setSpacing(16)

        # ---- هدر ----
        header = Card()
        hl = QHBoxLayout(header)
        hl.setContentsMargins(24, 18, 24, 18)
        titles = QVBoxLayout()
        title = QLabel("Excel Merger Pro")
        title.setObjectName("title")
        sub = QLabel("ادغام هوشمند فایل‌ها و شیت‌های اکسل")
        sub.setObjectName("muted")
        titles.addWidget(title)
        titles.addWidget(sub)
        hl.addLayout(titles)
        hl.addStretch()
        badge = QLabel("PyQt6 • Dark")
        badge.setObjectName("badge")
        hl.addWidget(badge)
        main.addWidget(header)

        body = QHBoxLayout()
        body.setSpacing(16)
        main.addLayout(body, 1)

        # ---- ستون تنظیمات ----
        left = Card()
        left.setFixedWidth(330)
        ll = QVBoxLayout(left)
        ll.setContentsMargins(20, 20, 20, 20)
        ll.setSpacing(10)
        ll.addWidget(section_label("حالت ادغام"))

        self.mode_group = QButtonGroup(self)
        self.mode_group.setExclusive(True)
        self.mode_buttons = {}
        for key, text in MODES:
            btn = QPushButton(text)
            btn.setObjectName("mode")
            btn.setCheckable(True)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            self.mode_group.addButton(btn)
            ll.addWidget(btn)
            self.mode_buttons[key] = btn
            btn.clicked.connect(lambda _checked, k=key: self.set_mode(k))
        self.mode_buttons[MODE_SHEETS].setChecked(True)

        ll.addSpacing(8)
        ll.addWidget(section_label("منبع"))
        self.source_btn = QPushButton()
        self.source_btn.setObjectName("ghost")
        self.source_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.source_btn.clicked.connect(self.pick_source)
        ll.addWidget(self.source_btn)

        self.source_label = QLabel()
        self.source_label.setObjectName("muted")
        self.source_label.setWordWrap(True)
        ll.addWidget(self.source_label)

        self.replace_box = QWidget()
        rb = QVBoxLayout(self.replace_box)
        rb.setContentsMargins(0, 8, 0, 0)
        rb.setSpacing(8)
        rb.addWidget(section_label("جایگزینی در نام"))
        self.find_entry = QLineEdit()
        self.find_entry.setPlaceholderText("کلمه‌ی قدیمی")
        self.replace_entry = QLineEdit()
        self.replace_entry.setPlaceholderText("کلمه‌ی جدید")
        rb.addWidget(self.find_entry)
        rb.addWidget(self.replace_entry)
        ll.addWidget(self.replace_box)

        self.source_chk = QCheckBox("افزودن ستون نام منبع")
        ll.addWidget(self.source_chk)

        ll.addStretch()

        self.start_btn = QPushButton("▶  شروع ادغام")
        self.start_btn.setObjectName("primary")
        self.start_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.start_btn.clicked.connect(self.start_merge)
        ll.addWidget(self.start_btn)

        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setFixedHeight(10)
        self.progress_bar.setValue(0)
        ll.addWidget(self.progress_bar)

        body.addWidget(left)

        # ---- ستون لیست و گزارش ----
        right = QVBoxLayout()
        right.setSpacing(16)

        files_card = Card()
        fl = QVBoxLayout(files_card)
        fl.setContentsMargins(20, 16, 20, 16)
        fl.setSpacing(10)
        fh = QHBoxLayout()
        self.files_title = section_label("📁 فایل‌های یافت‌شده")
        self.files_count = QLabel("0")
        self.files_count.setObjectName("badge")
        fh.addWidget(self.files_title)
        fh.addStretch()
        fh.addWidget(self.files_count)
        self.files_list = QListWidget()
        self.files_list.setObjectName("files")
        fl.addLayout(fh)
        fl.addWidget(self.files_list)
        right.addWidget(files_card, 1)

        log_card = Card()
        l2 = QVBoxLayout(log_card)
        l2.setContentsMargins(20, 16, 20, 16)
        l2.setSpacing(10)
        lh = QHBoxLayout()
        lh.addWidget(section_label("📜 گزارش عملیات"))
        lh.addStretch()
        clear_btn = QPushButton("پاک کردن")
        clear_btn.setObjectName("ghost")
        clear_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        clear_btn.clicked.connect(lambda: self.log_box.clear())
        lh.addWidget(clear_btn)
        self.log_box = QTextEdit()
        self.log_box.setObjectName("log")
        self.log_box.setReadOnly(True)
        l2.addLayout(lh)
        l2.addWidget(self.log_box)
        right.addWidget(log_card, 1)

        body.addLayout(right, 1)

        self.set_mode(MODE_SHEETS)

    # ---------- منطق UI ----------

    def log(self, text):
        self.log_box.append(text)
        sb = self.log_box.verticalScrollBar()
        sb.setValue(sb.maximum())

    def set_mode(self, key):
        # اگر همان حالت است و منبعی انتخاب شده، چیزی را ریست نکن
        if key == self.mode and self.files:
            return
        self.mode = key
        self.source_dir = None
        self.files = []
        self.files_list.clear()
        self.files_count.setText("0")
        self.source_label.setText("هنوز منبعی انتخاب نشده است.")

        is_one = key == MODE_ONE_SHEET
        self.source_btn.setText("📄  انتخاب فایل اکسل" if is_one else "📂  انتخاب پوشه فایل‌ها")
        self.files_title.setText("📑 شیت‌های این فایل" if is_one else "📁 فایل‌های یافت‌شده")
        self.replace_box.setVisible(key == MODE_SHEETS)
        self.source_chk.setVisible(key != MODE_SHEETS)

    def pick_source(self):
        if self.mode == MODE_ONE_SHEET:
            path, _ = QFileDialog.getOpenFileName(
                self, "انتخاب فایل اکسل", "", "Excel Files (*.xlsx *.xls)"
            )
            if not path:
                return
            try:
                names = read_sheet_names(path)
            except Exception as e:
                QMessageBox.warning(self, "خطا", f"خواندن فایل ممکن نبود:\n{e}")
                return
            self.source_dir = os.path.dirname(path)
            self.files = [path]
            self.source_label.setText(path)
            self.files_list.clear()
            self.files_list.addItems(names)
            self.files_count.setText(str(len(names)))
            self.log(f"📄 {os.path.basename(path)} — {len(names)} شیت یافت شد")
        else:
            folder = QFileDialog.getExistingDirectory(self, "پوشه فایل‌های اکسل را انتخاب کنید")
            if not folder:
                return
            self.source_dir = folder
            self.files = list_excel_files(folder)
            self.source_label.setText(folder)
            self.files_list.clear()
            names = [os.path.basename(f) for f in self.files]
            if names:
                self.files_list.addItems(names)
            self.files_count.setText(str(len(names)))
            if names:
                self.log(f"📁 {len(names)} فایل یافت شد در: {folder}")
            else:
                self.log("⚠️ هیچ فایل اکسلی در این پوشه یافت نشد.")

    def start_merge(self):
        if self.worker and self.worker.isRunning():
            return
        if not self.files:
            QMessageBox.warning(self, "هشدار", "ابتدا منبع (پوشه یا فایل) را انتخاب کنید.")
            return

        if self.mode == MODE_SHEETS:
            out = os.path.join(self.source_dir, "merged_excels.xlsx")
        elif self.mode == MODE_STACK:
            out = os.path.join(self.source_dir, "merged_stacked.xlsx")
        else:
            stem = os.path.splitext(os.path.basename(self.files[0]))[0]
            out = os.path.join(self.source_dir, f"{stem}_one_sheet.xlsx")

        self.worker = MergeWorker(
            self.mode,
            list(self.files),
            out,
            self.find_entry.text().strip(),
            self.replace_entry.text().strip(),
            self.source_chk.isChecked(),
        )
        self.worker.log.connect(self.log)
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.done.connect(self.on_done)

        self.progress_bar.setValue(0)
        self.start_btn.setEnabled(False)
        self.log("🚀 شروع فرآیند ادغام...")
        self.worker.start()

    def on_done(self, ok, message, path):
        self.start_btn.setEnabled(True)
        if ok:
            self.progress_bar.setValue(100)
            QMessageBox.information(self, "انجام شد", f"{message}\n\nفایل خروجی:\n{path}")
        else:
            QMessageBox.warning(self, "خطا", message)


def main():
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon(resource_path("Exl.ico")))
    app.setStyle("Fusion")
    app.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
    app.setFont(QFont("Tahoma", 10))
    app.setStyleSheet(QSS)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
