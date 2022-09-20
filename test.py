import gspread

#Accesessing sheet
sa = gspread.service_account()
sh = sa.open("Cinema")
wks = sh.worksheet("Godziny")


#Getting shift type and shift time values then moving to next row
row = 4
while row < 24:
    row = str(row)
    dzien = wks.acell('C' + row).value
    if dzien == None:
        break
    data_start = wks.acell('D'+ row).value
    data_end = wks.acell('E' + row).value
    zmiana = wks.acell('B' + row).value

    row = int(row)
    row += 1



