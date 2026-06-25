import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font, NamedStyle
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from datetime import datetime
import random

def create_excel_file():
    wb = Workbook()
    
    if 'Sheet' in wb.sheetnames:
        default_sheet = wb['Sheet']
        wb.remove(default_sheet)
    
    # ایجاد 4 شیت ساده
    create_simple_sheet1(wb)
    create_simple_sheet2(wb)
    create_simple_sheet3(wb)
    create_simple_sheet4(wb)
    
    filename = f"مدیریت_مشتریان_جعفریان_{datetime.now().strftime('%Y%m%d')}.xlsx"
    wb.save(filename)
    
    print(f"✅ فایل اکسل ایجاد شد: {filename}")
    print("📁 فایل در پوشه فایل‌های Codespace ذخیره شد")
    print("📥 برای دانلود: روی پنل فایل‌ها سمت چپ کلیک کنید → روی فایل .xlsx راست کلیک → Download")
    
    return filename

def create_simple_sheet1(wb):
    ws = wb.create_sheet(title="مشتریان و خرید ماهانه")
    
    # سرستون‌ها
    headers = [
        "کد مشتری", "نام مشتری", "مجموع خرید اردیبهشت (ریال)", 
        "مجموع خرید خرداد (ریال)", "درصد تغییر", "گروه مشتری",
        "رتبه فروش", "سبد اصلی", "تاریخ آخرین خرید", 
        "ویزیتور مسئول", "وضعیت", "هدف ماه آینده (ریال)"
    ]
    
    for col, header in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=header)
    
    # داده‌های نمونه
    for i in range(1, 44):
        ws.cell(row=i+1, column=1, value=i)
        ws.cell(row=i+1, column=2, value=f"مشتری {i}")
        ws.cell(row=i+1, column=3, value=random.randint(100000000, 900000000))
        ws.cell(row=i+1, column=4, value=random.randint(80000000, 1000000000))
        ws.cell(row=i+1, column=5, value=f"=(D{i+1}-C{i+1})/C{i+1}")
        ws.cell(row=i+1, column=6, value=random.choice(["A", "B", "C", "D"]))
        ws.cell(row=i+1, column=7, value=f"=RANK(D{i+1},D$2:D$44)")
        ws.cell(row=i+1, column=8, value=random.choice(["پنیر، شیر", "ماست، دوغ", "شیر، خامه"]))
        ws.cell(row=i+1, column=9, value=f"1403/0{random.randint(1,3)}/{random.randint(10,30)}")
        ws.cell(row=i+1, column=10, value=random.choice(["ویزیتور تهران", "ویزیتور اصفهان"]))
        ws.cell(row=i+1, column=11, value=random.choice(["رشد بالا", "رشد", "کاهش", "کاهش شدید"]))
        ws.cell(row=i+1, column=12, value=f"=D{i+1}*1.15")
    
    # تنظیم عرض ستون‌ها
    for col in range(1, 13):
        ws.column_dimensions[get_column_letter(col)].width = 20

def create_simple_sheet2(wb):
    ws = wb.create_sheet(title="تحلیل محصولات")
    
    headers = [
        "کد مشتری", "نام مشتری", "محصول", 
        "مقدار اردیبهشت", "مقدار خرداد", "تغییر مقدار",
        "درصد تغییر", "رتبه محصول", "گروه محصول",
        "سهم در کل فروش"
    ]
    
    for col, header in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=header)
    
    row_idx = 2
    products = ["شیر استریل", "پنیر ورقه‌ای", "ماست یونانی", "دوغ بطری", "خامه استریل"]
    for i in range(100):
        ws.cell(row=row_idx, column=1, value=random.randint(1, 43))
        ws.cell(row=row_idx, column=2, value=f"مشتری {random.randint(1, 43)}")
        ws.cell(row=row_idx, column=3, value=random.choice(products))
        ws.cell(row=row_idx, column=4, value=random.randint(100, 10000))
        ws.cell(row=row_idx, column=5, value=random.randint(50, 12000))
        ws.cell(row=row_idx, column=6, value=f"=E{row_idx}-D{row_idx}")
        ws.cell(row=row_idx, column=7, value=f"=F{row_idx}/D{row_idx}")
        ws.cell(row=row_idx, column=8, value=f"=RANK(E{row_idx},E$2:E$101)")
        ws.cell(row=row_idx, column=9, value=random.choice(["شیر", "پنیر", "ماست", "دوغ"]))
        ws.cell(row=row_idx, column=10, value=f"=E{row_idx}/SUM(E$2:E$101)")
        row_idx += 1
    
    for col in range(1, 11):
        ws.column_dimensions[get_column_letter(col)].width = 18

def create_simple_sheet3(wb):
    ws = wb.create_sheet(title="طرح فروش پیشنهادی")
    
    ws.cell(row=1, column=1, value="اهداف ماهانه مشتریان")
    ws.merge_cells('A1:I1')
    
    headers = [
        "کد مشتری", "نام مشتری", "خرید خرداد (ریال)",
        "هدف ماه آینده (ریال)", "افزایش مورد انتظار (ریال)",
        "محصولات پیشنهادی", "بودجه تخصیص‌یافته (ریال)",
        "اولویت ویزیت", "استراتژی فروش"
    ]
    
    for col, header in enumerate(headers, 1):
        ws.cell(row=2, column=col, value=header)
    
    for i in range(1, 44):
        row = i + 2
        ws.cell(row=row, column=1, value=i)
        ws.cell(row=row, column=2, value=f"مشتری {i}")
        ws.cell(row=row, column=3, value=random.randint(100000000, 1000000000))
        ws.cell(row=row, column=4, value=f"=C{row}*1.15")
        ws.cell(row=row, column=5, value=f"=D{row}-C{row}")
        ws.cell(row=row, column=6, value=random.choice(["شیر استریل, پنیر ورقه‌ای", "ماست یونانی, دوغ بطری"]))
        ws.cell(row=row, column=7, value=f"=E{row}*0.1")
        ws.cell(row=row, column=8, value=random.randint(1, 5))
        ws.cell(row=row, column=9, value=random.choice(["Cross-sell", "Up-sell", "New product"]))

def create_simple_sheet4(wb):
    ws = wb.create_sheet(title="داشبورد مدیریتی")
    
    ws.cell(row=1, column=1, value="داشبورد مدیریت فروش - شرکت جعفریان")
    ws.merge_cells('A1:H1')
    
    # KPI
    ws.cell(row=3, column=1, value="شاخص‌های عملکرد (KPI)")
    ws.merge_cells('A3:D3')
    
    kpi_headers = ["شاخص", "مقدار", "هدف", "وضعیت"]
    for col, header in enumerate(kpi_headers, 1):
        ws.cell(row=4, column=col, value=header)
    
    kpis = [
        ["فروش کل خرداد", "='مشتریان و خرید ماهانه'!D46", 15000000000],
        ["میانگین خرید مشتری", "='مشتریان و خرید ماهانه'!C50", 350000000],
        ["تعداد مشتریان رشد", "='مشتریان و خرید ماهانه'!C51", 30],
        ["نرخ رشد ماهانه", "='مشتریان و خرید ماهانه'!E46", 0.15]
    ]
    
    for idx, kpi in enumerate(kpis, 5):
        ws.cell(row=idx, column=1, value=kpi[0])
        ws.cell(row=idx, column=2, value=kpi[1])
        ws.cell(row=idx, column=3, value=kpi[2])

# اجرای کد
if __name__ == "__main__":
    create_excel_file()
