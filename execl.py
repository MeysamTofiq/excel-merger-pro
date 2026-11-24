import os
import copy
import pandas as pd
import openpyxl
from tkinter import Tk, filedialog


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


Tk().withdraw()

folder_path = filedialog.askdirectory(title="پوشه‌ای که فایل‌های اکسل داخلش هستند را انتخاب کنید")
if not folder_path:
    print("پوشه‌ای انتخاب نشد.")
    exit()

excel_files = [f for f in os.listdir(folder_path) if f.endswith(('.xlsx', '.xls'))]

if not excel_files:
    print("هیچ فایل اکسل در این پوشه یافت نشد.")
    exit()

output_path = os.path.join(folder_path, "merged_excels.xlsx")

successful = 0

out_wb = openpyxl.Workbook()
out_wb.remove(out_wb.active)

for file in excel_files:
    file_path = os.path.join(folder_path, file)
    try:
        if file.lower().endswith('.xls'):
            df = pd.read_excel(file_path, engine='xlrd')
            ws = out_wb.create_sheet(unique_sheet_name(out_wb, os.path.splitext(file)[0]))
            for col_idx, col_name in enumerate(df.columns, start=1):
                ws.cell(row=1, column=col_idx, value=col_name)
            for row_idx, values in enumerate(df.values, start=2):
                for col_idx, value in enumerate(values, start=1):
                    ws.cell(row=row_idx, column=col_idx, value=value)
            print(f"⚠️ فایل '{file}' فرمت قدیمی (xls) است؛ فقط مقادیر کپی شد.")
        else:
            src_wb = openpyxl.load_workbook(file_path)
            if len(src_wb.worksheets) > 1:
                for src_ws in src_wb.worksheets:
                    target = out_wb.create_sheet(unique_sheet_name(out_wb, src_ws.title))
                    copy_sheet_with_styles(src_ws, target)
            else:
                target = out_wb.create_sheet(unique_sheet_name(out_wb, os.path.splitext(file)[0]))
                copy_sheet_with_styles(src_wb.worksheets[0], target)
            src_wb.close()
        successful += 1
        print(f"✅ فایل '{file}' اضافه شد")

    except Exception as e:
        print(f"⚠️ خطا در خواندن '{file}': {e}")

if successful == 0:
    print("\n⚠️ هیچ فایلی با موفقیت خوانده نشد. فایل خروجی ساخته نشد.")
    if os.path.exists(output_path):
        os.remove(output_path)
else:
    out_wb.save(output_path)
    print(f"\n🎯 {successful} فایل با موفقیت ترکیب شد → {output_path}")
