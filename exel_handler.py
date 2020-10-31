from openpyxl import load_workbook
import datetime


def add_new_exel(surname, name, birthdate, citizen, birth_place, birth_city, doc_type, doc_seria, doc_number, doc_date,
                 doc_end, profession, date_income, region, district, city, street, street_number, flat_number, gender,
                 mig_card_ser, mig_card_number, mig_card_region, mig_card_city):
    wb = load_workbook('./exel.xlsx')
    sheet = wb.get_sheet_by_name('стр.1')
    sheet2 = wb.get_sheet_by_name('стр.2')
    one_year_from_now = datetime.datetime.now()
    date_formated = one_year_from_now.strftime("%d-%m-%Y")
    file_name = surname + ' ' + date_formated + '.xlsx'

    if len(surname) > 35:
        raise Exception('Surname is long')

    surname_upper = surname.upper()
    name_upper = name.upper()

    if gender == "female":
        sheet['DS24'] = 'X'

    else:
        sheet['CY24'] = 'X'

    cells_surname = sheet['W16':'FC16'][0]
    cell_index_surname = 0
    for char in surname_upper:
        cells_surname[cell_index_surname].value = char
        cell_index_surname = cell_index_surname + 4

    cells_surname1 = sheet['W71':'FC71'][0]
    cell_index_surname1 = 0
    for char in surname_upper:
        cells_surname1[cell_index_surname1].value = char
        cell_index_surname1 = cell_index_surname1 + 4

    cells_name = sheet['AI18':'FC18'][0]
    cell_index_name = 0
    for char_name in name_upper:
        cells_name[cell_index_name].value = char_name
        cell_index_name = cell_index_name + 4

    cells_name1 = sheet['AI73':'FC73'][0]
    cell_index_name1 = 0
    for char_name1 in name_upper:
        cells_name1[cell_index_name1].value = char_name1
        cell_index_name1 = cell_index_name1 + 4

    # insert birth date
    sheet["AE24"].value = birthdate[0]
    sheet["AI24"].value = birthdate[1]

    sheet["AE79"].value = birthdate[0]
    sheet["AI79"].value = birthdate[1]

    sheet["AU24"].value = birthdate[3]
    sheet["AY24"].value = birthdate[4]

    sheet["AU79"].value = birthdate[3]
    sheet["AY79"].value = birthdate[4]

    sheet["BG24"].value = birthdate[6]
    sheet["BK24"].value = birthdate[7]
    sheet["BO24"].value = birthdate[8]
    sheet["BS24"].value = birthdate[9]

    sheet["BG79"].value = birthdate[6]
    sheet["BK79"].value = birthdate[7]
    sheet["BO79"].value = birthdate[8]
    sheet["BS79"].value = birthdate[9]

    # insert pasport

    # insert citizen
    citizen_upper = citizen.upper()

    cells_citizen = sheet['AA21':'FC21'][0]
    cell_index_citizen = 0
    for char in citizen_upper:
        cells_citizen[cell_index_citizen].value = char
        cell_index_citizen = cell_index_citizen + 4

    cells_citizen1 = sheet['AA76':'FC76'][0]
    cell_index_citizen1 = 0
    for char in citizen_upper:
        cells_citizen1[cell_index_citizen1].value = char
        cell_index_citizen1 = cell_index_citizen1 + 4

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

    cells_doc_type1 = sheet['BC82':'CQ82'][0]
    cell_index_doc_type1 = 0
    for char in doc_type_upper:
        cells_doc_type1[cell_index_doc_type1].value = char
        cell_index_doc_type1 = cell_index_doc_type1 + 4

    # insert seria and number
    cells_doc_seria = sheet['DC33':'DO33'][0]
    cell_index_doc_seria = 0
    for char in doc_seria:
        cells_doc_seria[cell_index_doc_seria].value = char
        cell_index_doc_seria = cell_index_doc_seria + 4

    # insert seria and number
    cells_doc_seria1 = sheet['DC82':'DO82'][0]
    cell_index_doc_seria1 = 0
    for char in doc_seria:
        cells_doc_seria1[cell_index_doc_seria1].value = char
        cell_index_doc_seria1 = cell_index_doc_seria1 + 4

    cells_doc_number = sheet['DW33':'FC33'][0]
    cell_index_doc_number = 0
    for char in doc_number:
        cells_doc_number[cell_index_doc_number].value = char
        cell_index_doc_number = cell_index_doc_number + 4

    cells_doc_number1 = sheet['DW82':'FC82'][0]
    cell_index_doc_number1 = 0
    for char in doc_number:
        cells_doc_number1[cell_index_doc_number1].value = char
        cell_index_doc_number1 = cell_index_doc_number1 + 4

    cells_mig_card_ser = sheet['AQ52':'BC52'][0]
    cell_index_mig_card_ser = 0
    for char in mig_card_ser:
        cells_mig_card_ser[cell_index_mig_card_ser].value = char
        cell_index_mig_card_ser = cell_index_mig_card_ser + 4

    cells_mig_card_number = sheet['BK52':'CY52'][0]
    cell_index_mig_card_number = 0
    for char in mig_card_number:
        cells_mig_card_number[cell_index_mig_card_number].value = char
        cell_index_mig_card_number = cell_index_mig_card_number + 4

    cells_mig_card_region = sheet['AA60':'CU60'][0]
    cell_index_mig_card_region = 0
    for char in mig_card_region:
        cells_mig_card_region[cell_index_mig_card_region].value = char
        cell_index_mig_card_region = cell_index_mig_card_region + 4

    cells_mig_card_city = sheet['AA62':'CU62'][0]
    cell_index_mig_card_city = 0
    for char in mig_card_city:
        cells_mig_card_city[cell_index_mig_card_city].value = char
        cell_index_mig_card_city = cell_index_mig_card_city + 4

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
    sheet["AA35"].value = doc_date[0]
    sheet["AE35"].value = doc_date[1]

    sheet["AQ35"].value = doc_date[3]
    sheet["AU35"].value = doc_date[4]

    sheet["BC35"].value = doc_date[6]
    sheet["BG35"].value = doc_date[7]
    sheet["BK35"].value = doc_date[8]
    sheet["BO35"].value = doc_date[9]

    # insert document end date
    sheet["CM35"].value = doc_end[0]
    sheet["CQ35"].value = doc_end[1]

    sheet["DC35"].value = doc_end[3]
    sheet["DG35"].value = doc_end[4]

    sheet["DO35"].value = doc_end[6]
    sheet["DS35"].value = doc_end[7]
    sheet["DW35"].value = doc_end[8]
    sheet["EA35"].value = doc_end[9]

    # insert profession
    profession_upper = profession.upper()

    cells_profession_upper = sheet['AA48':'DK48'][0]
    cell_index_profession = 0
    for char in profession_upper:
        cells_profession_upper[cell_index_profession].value = char
        cell_index_profession = cell_index_profession + 4

    # insert income date
    sheet["AI50"].value = date_income[0]
    sheet["AM50"].value = date_income[1]

    sheet["AY50"].value = date_income[3]
    sheet["BC50"].value = date_income[4]

    sheet["BK50"].value = date_income[6]
    sheet["BO50"].value = date_income[7]
    sheet["BS50"].value = date_income[8]
    sheet["BW50"].value = date_income[9]

    region_upper = region.upper()

    cells_region_upper = sheet['AQ85':'FC85'][0]
    cell_index_region = 0
    for char in region_upper:
        cells_region_upper[cell_index_region].value = char
        cell_index_region = cell_index_region + 4
    #
    district_upper = district.upper()

    cells_district_upper = sheet['W88':'FC88'][0]
    cell_index_district = 0
    for char in district_upper:
        cells_district_upper[cell_index_district].value = char
        cell_index_district = cell_index_district + 4

    city_upper = city.upper()

    cells_city_upper = sheet['AE90':'FC90'][0]
    cell_index_city = 0
    for char in city_upper:
        cells_city_upper[cell_index_city].value = char
        cell_index_city = cell_index_city + 4

    street_upper = street.upper()

    cells_street_upper = sheet['W93':'FC93'][0]
    cell_index_street = 0
    for char in street_upper:
        cells_street_upper[cell_index_street].value = char
        cell_index_street = cell_index_street + 4

    cells_street_number = sheet['AI95':'FC95'][0]
    cell_index_street_number = 0
    for char in street_number:
        cells_street_number[cell_index_street_number].value = char
        cell_index_street_number = cell_index_street_number + 4

    cells_float_number = sheet['EM95':'FC95'][0]
    cell_index_flat_number = 0
    for char in flat_number:
        cells_float_number[cell_index_flat_number].value = char
        cell_index_flat_number = cell_index_flat_number + 4




    # сведения о месте пребывания

    cells_region_upper = sheet2['AQ14':'FC14'][0]
    cell_index_region = 0
    for char in region_upper:
        cells_region_upper[cell_index_region].value = char
        cell_index_region = cell_index_region + 4

    cells_district_upper = sheet2['W17':'FC17'][0]
    cell_index_district = 0
    for char in district_upper:
        cells_district_upper[cell_index_district].value = char
        cell_index_district = cell_index_district + 4

    cells_city_upper = sheet2['AE19':'FC19'][0]
    cell_index_city = 0
    for char in city_upper:
        cells_city_upper[cell_index_city].value = char
        cell_index_city = cell_index_city + 4

    cells_street_upper = sheet2['W22':'FC22'][0]
    cell_index_street = 0
    for char in street_upper:
        cells_street_upper[cell_index_street].value = char
        cell_index_street = cell_index_street + 4

    cells_street_number = sheet2['AQ24':'BS24'][0]
    cell_index_street_number = 0
    for char in street_number:
        cells_street_number[cell_index_street_number].value = char
        cell_index_street_number = cell_index_street_number + 4

    cells_float_number = sheet2['EQ24':'FC24'][0]
    cell_index_flat_number = 0
    for char in flat_number:
        cells_float_number[cell_index_flat_number].value = char
        cell_index_flat_number = cell_index_flat_number + 4


    # Сведения о принимающей стороне
    #
    cells_region_upper = sheet2['AQ41':'FC41'][0]
    cell_index_region = 0
    for char in region_upper:
        cells_region_upper[cell_index_region].value = char
        cell_index_region = cell_index_region + 4

    cells_district_upper = sheet2['W44':'FC44'][0]
    cell_index_district = 0
    for char in district_upper:
        cells_district_upper[cell_index_district].value = char
        cell_index_district = cell_index_district + 4

    cells_city_upper = sheet2['AE46':'FC46'][0]
    cell_index_city = 0
    for char in city_upper:
        cells_city_upper[cell_index_city].value = char
        cell_index_city = cell_index_city + 4

    cells_street_upper = sheet2['W49':'FC49'][0]
    cell_index_street = 0
    for char in street_upper:
        cells_street_upper[cell_index_street].value = char
        cell_index_street = cell_index_street + 4
    #
    cells_street_number = sheet2['S51':'AE51'][0]
    cell_index_street_number = 0
    for char in street_number:
        cells_street_number[cell_index_street_number].value = char
        cell_index_street_number = cell_index_street_number + 4
    #
    cells_float_number = sheet2['CM51':'CY51'][0]
    cell_index_flat_number = 0
    for char in flat_number:
        cells_float_number[cell_index_flat_number].value = char
        cell_index_flat_number = cell_index_flat_number + 4



    wb.save(file_name)
