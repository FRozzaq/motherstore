from datetime import datetime
from openpyxl import *
import calendar


class FillReport:
    def __init__(self):
        self.today = datetime.today()

    def get_file_name(self):
        file_dir = "D:/My work/Project/MS/"
        month_range = calendar.monthrange(self.today.year, self.today.month)
        return f"{file_dir}LAPORAN PENJUALAN VENDOR 1-{month_range[1]} {self.today.strftime('%B %Y')}.xlsx"

    def load_wb(self):
        # file_dir = "D:/My work/Project/MS/"
        # month_range = calendar.monthrange(self.today.year, self.today.month)
        file_name = self.get_file_name()
        wb = load_workbook(file_name)
        return wb

    def get_vendor_name(self):
        wb = self.load_wb()
        vendor = wb.sheetnames
        return vendor[1:]

    def get_vendor_data(self, vendor_name):
        wb = self.load_wb()
        return wb[vendor_name]


# data = FillReport()
# vendor_data = data.get_vendor_data("BLUES STRONG")
# print(vendor_data["B13"].value)
# products = [product.value for product in vendor_data["B"][12:] if product != "None"]
# print(products[:-2])
# print(len(products))
