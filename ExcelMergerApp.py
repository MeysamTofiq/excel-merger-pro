import os
import copy
import pandas as pd
import openpyxl
import customtkinter as ctk
from tkinter import filedialog, messagebox

# تنظیمات اولیه تم
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


def copy_sheet_with_styles(src_ws, dst_ws):
    """کپی کامل سلول‌ها همراه با استایل، سلول‌های ادغام‌شده و ابعاد"""
    for row in src_ws.iter_rows():
        for cell in row:
            dst = dst_ws[cell.coordinate]
            dst.value = cell.value
            if cell.has_style:
                dst.font = copy.copy(cell.font)
                dst.border = copy.copy(cell.border)
                dst.fill = copy.copy(cell.fill)
                dst.number_format = cell.number_format
                dst.protection = copy.copy(cell.protection)
                dst.alignment = copy.copy(cell.alignment)
    for rng in src_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(rng))
    for letter, dim in src_ws.column_dimensions.items():
        dst_ws.column_dimensions[letter].width = dim.width
    for idx, dim in src_ws.row_dimensions.items():
        dst_ws.row_dimensions[idx].height = dim.height


def unique_sheet_name(wb, name):
    """ساخت نام شیت یکتا (بیشتر از ۳۱ کاراکتر و تکراری نباشد)"""
    name = (name or "Sheet")[:31]
    candidate = name
    i = 1
    while candidate in wb.sheetnames:
        suffix = f"_{i}"
        candidate = name[:31 - len(suffix)] + suffix
        i += 1
    return candidate

class ExcelMergerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("📘 Excel Merger - CustomTkinter Edition")
        self.geometry("800x600")
        self.resizable(False, False)

        self.folder_path = None
        self.excel_files = []

        # فریم انتخاب پوشه
        frame_top = ctk.CTkFrame(self)
        frame_top.pack(fill="x", padx=10, pady=(10,5))

        self.select_btn = ctk.CTkButton(frame_top, text="📂 انتخاب پوشه فایل‌ها", command=self.select_folder)
        self.select_btn.pack(side="left", padx=10, pady=10)

        self.replace_label = ctk.CTkLabel(frame_top, text="جایگزینی در نام فایل:")
        self.replace_label.pack(side="left", padx=(30,5))

        self.find_entry = ctk.CTkEntry(frame_top, width=120, placeholder_text="کلمه‌ی قدیمی")
        self.find_entry.pack(side="left", padx=5)

        self.replace_entry = ctk.CTkEntry(frame_top, width=120, placeholder_text="کلمه‌ی جدید")
        self.replace_entry.pack(side="left", padx=5)

        self.start_btn = ctk.CTkButton(frame_top, text="▶️ شروع ادغام", command=self.merge_excels)
        self.start_btn.pack(side="right", padx=10, pady=10)

        # فریم لیست فایل‌ها
        frame_list = ctk.CTkFrame(self)
        frame_list.pack(fill="both", expand=True, padx=10, pady=5)

        self.files_listbox = ctk.CTkTextbox(frame_list, wrap="none", height=200)
        self.files_listbox.pack(fill="both", expand=True, padx=10, pady=10)

        # فریم گزارش عملیات
        frame_log = ctk.CTkFrame(self)
        frame_log.pack(fill="both", expand=True, padx=10, pady=(5,10))

        log_label = ctk.CTkLabel(frame_log, text="📜 گزارش عملیات:", anchor="w")
        log_label.pack(fill="x", padx=10, pady=(5,0))

        self.log_box = ctk.CTkTextbox(frame_log, wrap="word")
        self.log_box.pack(fill="both", expand=True, padx=10, pady=10)

    def log(self, text):
        """ثبت متن در باکس لاگ"""
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")
        self.update_idletasks()

    def select_folder(self):
        """انتخاب پوشه و لیست فایل‌های اکسل"""
        folder = filedialog.askdirectory(title="پوشه فایل‌های اکسل را انتخاب کنید")
        if not folder:
            return
        self.folder_path = folder
        self.excel_files = [f for f in os.listdir(folder) if f.endswith(('.xlsx', '.xls'))]

        self.files_listbox.delete("1.0", "end")
        if not self.excel_files:
            self.files_listbox.insert("end", "هیچ فایل اکسل در این پوشه یافت نشد.\n")
        else:
            self.files_listbox.insert("end", "\n".join(self.excel_files))

        self.log(f"📁 {len(self.excel_files)} فایل یافت شد در: {folder}")

    def merge_excels(self):
        """ادغام فایل‌ها در یک فایل نهایی"""
        if not self.folder_path or not self.excel_files:
            messagebox.showwarning("هشدار", "ابتدا پوشه فایل‌های اکسل را انتخاب کنید.")
            return

        find_text = self.find_entry.get().strip()
        replace_text = self.replace_entry.get().strip()

        output_path = os.path.join(self.folder_path, "merged_excels.xlsx")
        successful = 0

        self.log("🚀 شروع فرآیند ادغام...\n")

        out_wb = openpyxl.Workbook()
        out_wb.remove(out_wb.active)

        for file in self.excel_files:
            file_path = os.path.join(self.folder_path, file)
            try:
                sheet_name = os.path.splitext(file)[0]
                if find_text:
                    sheet_name = sheet_name.replace(find_text, replace_text)

                if file.lower().endswith('.xls'):
                    df = pd.read_excel(file_path, engine='xlrd')
                    ws = out_wb.create_sheet(unique_sheet_name(out_wb, sheet_name))
                    for col_idx, col_name in enumerate(df.columns, start=1):
                        ws.cell(row=1, column=col_idx, value=col_name)
                    for row_idx, values in enumerate(df.values, start=2):
                        for col_idx, value in enumerate(values, start=1):
                            ws.cell(row=row_idx, column=col_idx, value=value)
                    self.log(f"⚠️ '{file}' فرمت قدیمی (xls) است؛ فقط مقادیر کپی شد.")
                else:
                    src_wb = openpyxl.load_workbook(file_path)
                    if len(src_wb.worksheets) > 1:
                        for src_ws in src_wb.worksheets:
                            target = out_wb.create_sheet(unique_sheet_name(out_wb, src_ws.title))
                            copy_sheet_with_styles(src_ws, target)
                    else:
                        target = out_wb.create_sheet(unique_sheet_name(out_wb, sheet_name))
                        copy_sheet_with_styles(src_wb.worksheets[0], target)
                    src_wb.close()

                self.log(f"✅ {file} → شیت '{sheet_name}' اضافه شد")
                successful += 1

            except Exception as e:
                self.log(f"⚠️ خطا در '{file}': {e}")

        if successful == 0:
            self.log("❌ هیچ فایلی با موفقیت اضافه نشد.")
            messagebox.showerror("خطا", "هیچ فایلی با موفقیت ادغام نشد.")
        else:
            out_wb.save(output_path)
            self.log(f"\n🎯 {successful} فایل با موفقیت ادغام شد → {output_path}")
            messagebox.showinfo("انجام شد", f"{successful} فایل ترکیب شد.\n\nفایل خروجی:\n{output_path}")

if __name__ == "__main__":
    app = ExcelMergerApp()
    app.mainloop()
