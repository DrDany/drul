import datetime

one_year_from_now = datetime.datetime.now()
date_formated = one_year_from_now.strftime("%d-%m-%Y")
print (date_formated)