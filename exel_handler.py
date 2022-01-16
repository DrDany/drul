# coding=utf-8
from openpyxl import load_workbook
import datetime


def input_cell(page, start_cell, end_cell, word):
    cell_index = 0
    cells = page[start_cell:end_cell][0]
    for char in word:
        cells[cell_index].value = char
        cell_index = cell_index + 4


def add_new_exel(surname='', name='', patranomic='', citizen='', birthdate='', gender='', doc_seria='', doc_number='',
                 doc_date='', doc_end='',
                 profession='', date_income='',
                 date_stay_to='', mig_card_ser='',
                 mig_card_number='', surname_host='', name_host='', patr_host='', host_doc_seria='', host_doc_number='',
                 date_host_pass=''):
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
    if not birthdate:
        print("string empty")

    if birthdate:
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

    # document

    input_cell(sheet, 'BF28', 'BR28', doc_seria)
    input_cell(sheet3, 'BF47', 'BR47', doc_seria)

    input_cell(sheet, 'BZ28', 'DN28', doc_number)
    input_cell(sheet3, 'BZ47', 'DN47', doc_number)

    if not doc_date:
        print("string empty")

    if doc_date:
    # doc_date
        sheet["I30"].value = doc_date[0]
        sheet["M30"].value = doc_date[1]

        sheet3["I49"].value = doc_date[0]
        sheet3["M49"].value = doc_date[1]
        # month
        sheet["Z30"].value = doc_date[3]
        sheet["AD30"].value = doc_date[4]

        sheet3["Z49"].value = doc_date[3]
        sheet3["AD49"].value = doc_date[4]

        sheet["AL30"].value = doc_date[6]
        sheet["AP30"].value = doc_date[7]
        sheet["AT30"].value = doc_date[8]
        sheet["AX30"].value = doc_date[9]

        sheet3["AL49"].value = doc_date[6]
        sheet3["AP49"].value = doc_date[7]
        sheet3["AT49"].value = doc_date[8]
        sheet3["AX49"].value = doc_date[9]

    # doc_end

    if not doc_end:
        print("string empty")

    if doc_end:
        sheet["BN30"].value = doc_end[0]
        sheet["BR30"].value = doc_end[1]

        sheet3["BN49"].value = doc_end[0]
        sheet3["BR49"].value = doc_end[1]
        # month
        sheet["CD30"].value = doc_end[3]
        sheet["CH30"].value = doc_end[4]

        sheet3["CD49"].value = doc_end[3]
        sheet3["CH49"].value = doc_end[4]
        #
        sheet["CP30"].value = doc_end[6]
        sheet["CT30"].value = doc_end[7]
        sheet["CX30"].value = doc_end[8]
        sheet["DB30"].value = doc_end[9]

        sheet3["CP49"].value = doc_end[6]
        sheet3["CT49"].value = doc_end[7]
        sheet3["CX49"].value = doc_end[8]
        sheet3["DB49"].value = doc_end[9]

    input_cell(sheet, 'R44', 'DN44', profession)

    if not date_income:
        print("string empty")

    if date_income:
        sheet["I46"].value = date_income[0]
        sheet["M46"].value = date_income[1]

        sheet["Z46"].value = date_income[3]
        sheet["AD46"].value = date_income[4]

        sheet["AL46"].value = date_income[6]
        sheet["AP46"].value = date_income[7]
        sheet["AT46"].value = date_income[8]
        sheet["AX46"].value = date_income[9]

    if not date_stay_to:
        print("string empty")
    if date_stay_to:
        sheet["BN46"].value = date_stay_to[0]
        sheet["BR46"].value = date_stay_to[1]

        sheet["CD46"].value = date_stay_to[3]
        sheet["CH46"].value = date_stay_to[4]

        sheet["CP46"].value = date_stay_to[6]
        sheet["CT46"].value = date_stay_to[7]
        sheet["CX46"].value = date_stay_to[8]
        sheet["DB46"].value = date_stay_to[9]

        sheet3["I68"].value = date_stay_to[0]
        sheet3["M68"].value = date_stay_to[1]

        sheet3["AA68"].value = date_stay_to[3]
        sheet3["AE68"].value = date_stay_to[4]

        sheet3["AM68"].value = date_stay_to[6]
        sheet3["AQ68"].value = date_stay_to[7]
        sheet3["AU68"].value = date_stay_to[8]
        sheet3["AY68"].value = date_stay_to[9]

    input_cell(sheet, 'AP48', 'BB48', mig_card_ser)
    input_cell(sheet, 'BJ48', 'CX48', mig_card_number)

    input_cell(sheet3, 'N5', 'DN5', surname_host)
    input_cell(sheet4, 'N27', 'DN27', surname_host)
    input_cell(sheet3, 'N7', 'DN7', name_host)
    input_cell(sheet4, 'N29', 'DN29', name_host)
    input_cell(sheet3, 'AH9', 'DN9', patr_host)
    input_cell(sheet4, 'Z31', 'DN31', patr_host)

    input_cell(sheet3, 'BF11', 'BR11', host_doc_seria)
    input_cell(sheet3, 'BZ11', 'DN11', host_doc_number)

    if not date_host_pass:
        print("string empty")
    if date_host_pass:
        sheet3["I13"].value = date_host_pass[0]
        sheet3["M13"].value = date_host_pass[1]

        sheet3["Z13"].value = date_host_pass[3]
        sheet3["AD13"].value = date_host_pass[4]

        sheet3["AL13"].value = date_host_pass[6]
        sheet3["AP13"].value = date_host_pass[7]
        sheet3["AT13"].value = date_host_pass[8]
        sheet3["AX13"].value = date_host_pass[9]

    wb.save(file_name)
