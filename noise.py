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