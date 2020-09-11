from openpyxl import load_workbook


def add_new_exel(surname, name, birthdate):
    wb = load_workbook('./exel.xlsx')
    sheet = wb.get_sheet_by_name('стр.1')
    file_name = surname + '.xlsx'

    # insert birth date
    sheet["AE24"].value = birthdate[0]
    sheet["AI24"].value = birthdate[1]

    sheet["AU24"].value = birthdate[3]
    sheet["AY24"].value = birthdate[4]

    sheet["BG24"].value = birthdate[6]
    sheet["BK24"].value = birthdate[7]
    sheet["BO24"].value = birthdate[8]
    sheet["BS24"].value = birthdate[9]

    # insert pasport





    if len(surname) > 35:
        raise Exception('Surname is long')

    surname_upper = surname.upper()
    name_upper = name.upper()

    cells_surname = sheet['W16':'FC16'][0]
    cell_index_surname = 0
    for char in surname_upper:
        cells_surname[cell_index_surname].value = char
        cell_index_surname = cell_index_surname + 4

    cells_name = sheet['AI18':'FC18'][0]
    cell_index_name = 0
    for char_name in name_upper:
        cells_name[cell_index_name].value = char_name
        cell_index_name = cell_index_name + 4

    wb.save(file_name)










