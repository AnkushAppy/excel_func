# Read but writes in new copied file
# Change input file_name

import xlrd
from xlrd import open_workbook
from xlutils.copy import copy
#file_path = 'C:/Users/Dolcera/Desktop/script/f2.xlsx'
file_path = 'demo_1.xlsx'
book = xlrd.open_workbook(file_path)
sh = book.sheet_by_index(0)
wb = copy(book)
ws = wb.get_sheet(0)
count = 0

abandon_list = ['EXPIRED',"LAPS","LAPSE","ABANDON","FAILURE","CEASE","CEASED","EXPIRY","WITHDRAWN","DEAD"]

for rx in range(sh.nrows):
    legal_status = sh.cell_value(rowx=rx,colx=15)
    first_split = legal_status.split('|')[0]
    #print first_split
    for x in abandon_list:
        if x in first_split:
            ws.write(rx,26,"Abandon")
            count = count +1
            break
        else:
            ws.write(rx,26,"Fine")
wb.save('demo_example.xls')
#print count
