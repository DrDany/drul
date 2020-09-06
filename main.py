from openpyxl import load_workbook
wb = load_workbook('./exel.xlsx')
print(wb.get_sheet_names())

sheet = wb.get_sheet_by_name('стр.1')

K16 = sheet['K16'].value
print(K16)

sheet['W16'] = 'Ф'
sheet['AA16'] = 'E'
sheet['AE16'] = 'Д'
sheet['AI16'] = 'О'
sheet['AM16'] = 'Р'
sheet['AQ16'] = 'Е'
sheet['AU16'] = 'Н'
sheet['AY16'] = 'К'
sheet['BC16'] = 'О'

wb.save('fedorenko.xlsx')
