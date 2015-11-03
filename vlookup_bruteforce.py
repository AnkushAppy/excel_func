import xlrd
from xlrd import open_workbook
from xlutils.copy import copy
import time

start_time = time.time()

##lookup_table = 'book_for_vlookup.xlsx'
##data_table = 'book_for_data_copy.xlsx'
##lookup_table = 'Adobe_2015_RAW_clean.xlsx'
##data_table = 'adobe_app_1_cleaned.xlsx'
##lookup_table = 'ibm_app_removed_small_normalised.xlsx'
##data_table = 'ibm_grant_all_normalized_small.xlsx'
lookup_table = 'test_1.xlsx'
data_table = 'test_1_copy.xlsx'
book_vl = xlrd.open_workbook(lookup_table)
book_dt = xlrd.open_workbook(data_table)
sheet_vl = book_vl.sheet_by_index(0)
sheet_dt = book_dt.sheet_by_index(0)

temp_bk = copy(book_vl)
temp_sh = temp_bk.get_sheet(0)

vlk_inp_1 = int(raw_input("1st vlookup input "))
vlk_inp_2 = int(raw_input("2nd vlookup input "))
vlk_inp_3 = int(raw_input("3rd vlookup input "))
vlk_inp_4 = int(raw_input("4th vlookup input column tobe modified "))

for x in range(sheet_vl.nrows):
    val = sheet_vl.cell_value(rowx = x,colx = vlk_inp_1)
    #print '-->',val
    for y in range (sheet_dt.nrows):
        val_y = sheet_dt.cell_value( rowx = y, colx = vlk_inp_2)
        #sheet_dt.cell_value( rowx = y, colx = 1)
        #print val_y
        if(val == val_y):
            temp_sh.write(x,vlk_inp_4,sheet_dt.cell_value( rowx = y, colx = vlk_inp_3))
            break


#temp_bk.save('example.xls')
temp_bk.save('example_v1.xls')
end_time = time.time()
print end_time - start_time
