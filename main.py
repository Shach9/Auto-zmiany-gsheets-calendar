import gspread

sa = gspread.service_account()
sh = sa.open("Cinema")

wks = sh.worksheet("Godziny")

data = wks.acell('C4').value

data_godzina = data + 'T' + wks.acell('D4').value + ':00'
print(data_godzina)
