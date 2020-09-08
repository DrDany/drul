from openpyxl import load_workbook
import argparse

parser = argparse.ArgumentParser()
parser.add_argument("--surname", type=str, required=True, help="second name", action="store", dest="surname")
args = parser.parse_args()
surname = args.surname
surname_upper = surname.upper()
file_name = surname + '.xlsx'
result = list(surname_upper)

list = ('W16', 'AA16', 'AE16', 'AI16', 'AM16', 'AQ16', 'AU16', 'AY16', 'BC16', 'BG16', 'BK16', 'BO16', 'BS16', 'BW16', 'CA16', 'CE16', 'CI16', 'CM16','CQ16')


wb = load_workbook('./exel.xlsx')

sheet = wb.get_sheet_by_name('стр.1')

for letter in range(len(surname)):
    sheet[list[letter]] = surname_upper[letter]

#
wb.save(file_name)
