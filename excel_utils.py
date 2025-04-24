from openpyxl import Workbook

def create_excel_file():
    workbook = Workbook()
    workbook.save("mon_fichier.xlsx")