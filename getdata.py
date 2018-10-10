import xlrd
from oauth2client import file, client, tools
group_name = 'ФИИТ-15'
book = xlrd.open_workbook('imi2018.xls')
curr_sheet = book.sheets()[8]
for i in range(20): 
    if group_name in curr_sheet.cell(2,i).value: col = i
for j in range(38):
    time=curr_sheet.cell(j,1).value
    subj=curr_sheet.cell(j,col).value
    print(time + subj)