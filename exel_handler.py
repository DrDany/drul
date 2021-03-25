from openpyxl import load_workbook
import datetime

def input_cell(page,start_cell, end_cell, word):
    cell_index = 0
    cells = page[start_cell:end_cell][0]
    for char in word:
        cells[cell_index].value = char
        cell_index = cell_index + 4


def add_new_exel(surname, name, patranomic):
    wb = load_workbook('exel.xlsx')
    sheet = wb.get_sheet_by_name('стр.1')
    sheet2 = wb.get_sheet_by_name('стр.2')
    sheet3 = wb.get_sheet_by_name('стр.3')
    sheet4 = wb.get_sheet_by_name('стр.4')
    one_year_from_now = datetime.datetime.now()
    date_formated = one_year_from_now.strftime("%d-%m-%Y")
    file_name = surname + ' ' + date_formated + '.xlsx'

    if len(surname) > 35:
        raise Exception('Surname is long')

    surname_upper = surname.upper()

    input_cell(sheet, 'N11', 'DN11', surname_upper)
    input_cell(sheet3, 'N31', 'DN31', surname_upper)

    input_cell(sheet, 'N13', 'DN13', name)
    input_cell(sheet3, 'N33', 'DN33', name)

    input_cell(sheet, 'Z15', 'DN15', patranomic)
    input_cell(sheet3, 'AH35', 'DN35', patranomic)





    wb.save(file_name)
