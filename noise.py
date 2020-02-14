import random
import xlwt 
from xlwt import Workbook 

noise_list = [-1.0,0,1]
  
# Workbook is created 
wb = Workbook() 
  
# add_sheet is used to create sheet. 
sheet1 = wb.add_sheet('Sheet 1', cell_overwrite_ok=True) 
# worksheet = workbook.add_sheet("Sheet 1", cell_overwrite_ok=True)

for i in range(380):
    for j in range(2):
        sheet1.write(i, 2,random.choice(noise_list)) 
        sheet1.write(i, 3,random.choice(noise_list)) 

wb.save('noise.xls') 


''' Code Insert for adding noise'''

'''
loc2 = ("noise.xls")
    wb2 = xlrd.open_workbook(loc2)
    sheet2 = wb2.sheet_by_index(0) 


scorei = int(sheet.cell_value(row,2))
scorej = int(sheet.cell_value(row,3))
    if scorei == 0 and int(sheet2.cell_value(row,2)) == -1:
        scorei = 0
    elif scorej == 0 and int(sheet2.cell_value(row,3)) == -1:
        scorej = 0
    else:
        scorei += int(sheet2.cell_value(row,2))
        scorej += int(sheet2.cell_value(row,3))
'''