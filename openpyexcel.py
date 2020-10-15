import openpyxl

import pandas as pd

def read_excel(file_path, keep_header):
    wb=openpyxl.load_workbook(file_path)
    ws = wb.active
    if not keep_header:
        start = 1
    else:
        start = 0

    return [
        tuple(cell.value for cell in row)
        for row in ws
    ][start:]


def combine_excel_files(folder_path):
    data = []
    keep_header = True
    for file_path in folder_path.glob('*.xls*'):
        data += read_excel(file_path, keep_header)
        keep_header = False
    return data



def save_dataframe(df, file_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Combined Data'
    for row in df:
        ws.append(row)
    wb.save(file_path)