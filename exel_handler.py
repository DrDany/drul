from openpyxl import load_workbook
import datetime

def input_cell(page,start_cell, end_cell, word):
    cell_index = 0
    cells = page[start_cell:end_cell][0]
    for char in word:
        cells[cell_index].value = char
        cell_index = cell_index + 4


def add_new_exel(surname, name, patranomic, citizen, birthdate, gender):
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

    input_cell(sheet, 'V17', 'DN17', citizen)
    input_cell(sheet, 'Z22', 'DN22', citizen)
    input_cell(sheet3, 'R37', 'DN37', citizen)
    input_cell(sheet3, 'Z41', 'DN41', citizen)

    # birthdate
    # insert birth date
    sheet["AD20"].value = birthdate[0]
    sheet["AH20"].value = birthdate[1]

    sheet3["AA39"].value = birthdate[0]
    sheet3["AE39"].value = birthdate[1]
    # month
    sheet["AT20"].value = birthdate[3]
    sheet["AX20"].value = birthdate[4]

    sheet3["AQ39"].value = birthdate[3]
    sheet3["AU39"].value = birthdate[4]

    sheet["BF20"].value = birthdate[6]
    sheet["BJ20"].value = birthdate[7]
    sheet["BN20"].value = birthdate[8]
    sheet["BR20"].value = birthdate[9]

    sheet3["BC39"].value = birthdate[6]
    sheet3["BG39"].value = birthdate[7]
    sheet3["BK39"].value = birthdate[8]
    sheet3["BO39"].value = birthdate[9]


    if gender == "female":
        sheet['DB20'] = 'X'
        sheet3['DB39'] = 'X'

    else:
        sheet['CL20'] = 'X'
        sheet3['CL39'] = 'X'




    wb.save(file_name)
