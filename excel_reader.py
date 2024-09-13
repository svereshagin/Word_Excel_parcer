import openpyxl

def excel_function():
    file_path = 'список ФИО для отчетов.xlsx'
    wb = openpyxl.load_workbook(file_path)

    sheet = wb.active

    data = []

    for row in sheet.iter_rows(min_row=2, values_only=True):
        fio = row[0]
        email = row[1]
        data.append((fio, email))
    return data


data = tuple(excel_function())