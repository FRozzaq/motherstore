from openpyxl import *
from datetime import datetime
import os
import calendar

FILE_DIR = "D:/My work/Project/MS/"

today = datetime.today()
last_month = today.replace(month=today.month - 1)
month_range = calendar.monthrange(today.year, today.month)
last_month_range = calendar.monthrange(last_month.year, last_month.month)


wb = load_workbook(f"{FILE_DIR}LAPORAN PENJUALAN VENDOR 1-{last_month_range[1]} {last_month.strftime('%B %Y')}.xlsx")
wb_data = load_workbook(f"{FILE_DIR}LAPORAN PENJUALAN VENDOR 1-{last_month_range[1]} "
                        f"{last_month.strftime('%B %Y')}.xlsx", data_only=True)


def get_stok_akhir(vendor_name):
    vendor_sheet = wb_data[vendor_name]
    return [stok_akhir.value for stok_akhir in vendor_sheet["L"][12:]]


vendor_names = wb.sheetnames
stok_akhir_dict = {vendor: get_stok_akhir(vendor) for vendor in vendor_names}

# set stok awal = stok akhir
for name in vendor_names:
    vendor = wb[name]
    if vendor == wb["SUMMARY"]:
        vendor["I8"] = today.strftime("%B")
        vendor["I9"] = today.strftime(f"1-{today.strftime('%d %B %Y')}")
    else:
        # set periode
        vendor["N8"] = today.strftime("=SUMMARY!I8")
        vendor["N9"] = today.strftime(f"=SUMMARY!I9")

        for index in range(len(stok_akhir_dict[name])):
            vendor[f"F{index+13}"] = stok_akhir_dict[name][index]
            vendor[f"H{index+13}"] = 0
            vendor[f"N{index+13}"] = ""

FILE_NAME = f"LAPORAN PENJUALAN VENDOR 1-{month_range[1]} {today.strftime('%B %Y')}.xlsx"
wb.save(f"{FILE_DIR}{FILE_NAME}")

os.system(f'start excel.exe "{FILE_DIR}{FILE_NAME}"')
