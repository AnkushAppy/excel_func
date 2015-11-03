import xlrd
from xlrd import open_workbook
from xlutils.copy import copy
import time
import xlsxwriter

start_time = time.time()


##lookup_table = 'book_for_vlookup.xlsx'
##data_table = 'book_for_data_copy.xlsx'
#lookup_table = 'Adobe_2015_RAW_clean.xlsx'
#data_table = 'adobe_app_1_cleaned.xlsx'
# lookup_table = 'test_1.xlsx'
# data_table = 'test_1_copy.xlsx'
lookup_table = 'sheet_maked.xlsx'
data_table = 'sheet_maked_2.xlsx'
book_vl = xlrd.open_workbook(lookup_table)
book_dt = xlrd.open_workbook(data_table)
sheet_vl = book_vl.sheet_by_index(0)
sheet_dt = book_dt.sheet_by_index(0)


#temp_bk = copy(book_vl)
#temp_sh = temp_bk.get_sheet(0)

temp_book = xlsxwriter.Workbook('example_2.xlsx')
temp_worksheet = temp_book.add_worksheet()

vlk_inp_1 = int(raw_input("1st vlookup input "))
vlk_inp_2 = int(raw_input("2nd vlookup input "))
vlk_inp_3 = int(raw_input("3rd vlookup input "))
vlk_inp_4 = int(raw_input("4th vlookup input column to be modified "))

dictionary_for_data_table = {}   ## Dictionary for second sheet
duplicate_count_in_sheet_2 = 0
for y in range (sheet_dt.nrows):
    val = sheet_dt.cell_value( rowx = y, colx = vlk_inp_2)
    checker = dictionary_for_data_table.has_key(val)        # gives true -> (if duplicate found) and flase
    if checker:
        duplicate_count_in_sheet_2 = duplicate_count_in_sheet_2 + 1
    else:
        print val,sheet_dt.cell_value( rowx = y, colx = vlk_inp_3)
        dictionary_for_data_table[val] = sheet_dt.cell_value( rowx = y, colx = vlk_inp_3)

print "%d Duplicate(s) Found! Press X to Exit OR Keep Going"% duplicate_count_in_sheet_2
print "Enter 'X' to proceed"
string_raw = raw_input()
if string_raw == 'X':       
    for x in range(sheet_vl.nrows):
        val = sheet_vl.cell_value(rowx = x,colx = vlk_inp_1)
        for col in range(sheet_vl.ncols):
            temp_worksheet.write(x,col,sheet_vl.cell_value(rowx = x,colx = col))

        checker = dictionary_for_data_table.has_key(val)
        if checker:
            if dictionary_for_data_table.get(val) == '':
                temp_worksheet.write(x,vlk_inp_4,'#Blank')
            else:
                temp_worksheet.write(x,vlk_inp_4,dictionary_for_data_table.get(val))
        else:
            temp_worksheet.write(x,vlk_inp_4,'#N/A')       

    temp_book.close()

    #temp_bk.save('example.xls')
    #temp_bk.save('example_1.xls')
    end_time = time.time()
    print end_time - start_time
else:
    print "Exited!!"
