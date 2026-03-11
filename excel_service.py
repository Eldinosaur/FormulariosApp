import os
from openpyxl import Workbook, load_workbook

FILE_PATH = 'data/data.xlsx'

def create_excel_file(data):
    if not os.path.exists(FILE_PATH):
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.append(['ID', 'Name', 'Email'])
        workbook.save(FILE_PATH)
        
    workbook = load_workbook(FILE_PATH)
    worksheet = workbook.active
    worksheet.append(data)
    workbook.save(FILE_PATH)