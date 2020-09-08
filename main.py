from openpyxl import load_workbook
import argparse

parser = argparse.ArgumentParser()
parser.add_argument("--surname", type=str, required=True, help="second name", action="store", dest="surname")
args = parser.parse_args()
surname = args.surname
surname_upper = surname.upper()
file_name = surname + '.xlsx'
result = list(surname_upper)




wb = load_workbook('./exel.xlsx')

sheet = wb.get_sheet_by_name('стр.1')


#
sheet['W16'] = surname_upper[0]
sheet['AA16'] = surname_upper[1]
sheet['AE16'] = surname_upper[2]
sheet['AI16'] = surname_upper[3]
sheet['AM16'] = surname_upper[4]
sheet['AQ16'] = surname_upper[5]
# sheet['AU16'] = surname_upper[6]
# sheet['AY16'] = 'К'
# sheet['BC16'] = 'О'
#
wb.save(file_name)
