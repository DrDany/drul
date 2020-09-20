from openpyxl import load_workbook


def add_new_exel(surname, name, birthdate, citizen, birth_place, birth_city, doc_type, doc_seria, doc_number, doc_date, doc_end):
    wb = load_workbook('./exel.xlsx')
    sheet = wb.get_sheet_by_name('стр.1')
    file_name = surname + '.xlsx'

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

    # insert citizen
    citizen_upper = citizen.upper()

    cells_citizen = sheet['AA21':'FC21'][0]
    cell_index_citizen = 0
    for char in citizen_upper:
        cells_citizen[cell_index_citizen].value = char
        cell_index_citizen = cell_index_citizen + 4

    # insert birth place
    birth_place_upper = birth_place.upper()

    cells_birth_place = sheet['AE27':'FC27'][0]
    cell_index_birth_place = 0
    for char in birth_place_upper:
        cells_birth_place[cell_index_birth_place].value = char
        cell_index_birth_place = cell_index_birth_place + 4

    # insert birth city
    birth_city_upper = birth_city.upper()

    cells_birth_city = sheet['AE30':'FC30'][0]
    cell_index_birth_city = 0
    for char in birth_city_upper:
        cells_birth_city[cell_index_birth_city].value = char
        cell_index_birth_city = cell_index_birth_city + 4

    # insert doc type
    doc_type_upper = doc_type.upper()

    cells_doc_type = sheet['BC33':'CQ33'][0]
    cell_index_doc_type = 0
    for char in doc_type_upper:
        cells_doc_type[cell_index_doc_type].value = char
        cell_index_doc_type = cell_index_doc_type + 4

    # insert seria and number
    cells_doc_seria = sheet['DC33':'DO33'][0]
    cell_index_doc_seria = 0
    for char in doc_seria:
        cells_doc_seria[cell_index_doc_seria].value = char
        cell_index_doc_seria = cell_index_doc_seria + 4

    cells_doc_number = sheet['DW33':'FC33'][0]
    cell_index_doc_number = 0
    for char in doc_number:
        cells_doc_number[cell_index_doc_number].value = char
        cell_index_doc_number = cell_index_doc_number + 4

    # if len(surname) > 35:
    #     raise Exception('Surname is long')
    #
    # surname_upper = surname.upper()
    # name_upper = name.upper()
    #
    # cells_surname = sheet['W16':'FC16'][0]
    # cell_index_surname = 0
    # for char in surname_upper:
    #     cells_surname[cell_index_surname].value = char
    #     cell_index_surname = cell_index_surname + 4
    #
    # cells_name = sheet['AI18':'FC18'][0]
    # cell_index_name = 0
    # for char_name in name_upper:
    #     cells_name[cell_index_name].value = char_name
    #     cell_index_name = cell_index_name + 4

    # insert document start date
    sheet["AA35"].value = doc_end[0]
    sheet["AE35"].value = doc_end[1]

    sheet["AQ35"].value = doc_end[3]
    sheet["AU35"].value = doc_end[4]

    sheet["BC35"].value = doc_end[6]
    sheet["BG35"].value = doc_end[7]
    sheet["BK35"].value = doc_end[8]
    sheet["BO35"].value = doc_end[9]

    # insert document end date
    sheet["CM35"].value = doc_end[0]
    sheet["CQ35"].value = doc_end[1]

    sheet["DC35"].value = doc_end[3]
    sheet["DG35"].value = doc_end[4]

    sheet["DO35"].value = doc_end[6]
    sheet["DS35"].value = doc_end[7]
    sheet["DW35"].value = doc_end[8]
    sheet["EA35"].value = doc_end[9]

    wb.save(file_name)










