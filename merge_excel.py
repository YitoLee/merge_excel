
import xlwings as xw
from datetime import datetime
import os
import time


combo_data={}
allfiles=os.listdir()
for filesname in allfiles:
#  通过循环目录下所有的excel文件,每个文件的每个表格,形成一个字典包含
#   combo_data['文件名字']['sheet1']=range()
#   combo_data['文件名字']['sheet2']=range()
#   combo_data['文件名字']['sheet3']=range()
#.......
    try:
        if filesname.split(".")[1]=='xlsx' or filesname.split(".")[1]=='xls':
            
            wb=xw.Book(filesname)
            combo_data[filesname]=[]
            wbsheets=wb.sheets
            for sheet in wbsheets:
                #  print(sheet.name)
                used_range_rows = (sheet.api.UsedRange.Row,
                        sheet.api.UsedRange.Row + sheet.api.UsedRange.Rows.Count)
                used_range_cols = (sheet.api.UsedRange.Column,
                        sheet.api.UsedRange.Column + sheet.api.UsedRange.Columns.Count)
                #  print(used_range_rows,used_range_cols)
                used_range = sheet.range(*zip(used_range_rows, used_range_cols))
                #  print(sheet.name)
                print(used_range)
                combo_data[filesname].append(used_range.value)
                #  print(combo_data[filesname])
            wb.close
            for app in xw.apps:
                app.quit()
            
    except Exception as e:
        #  print(e)
        pass
time.sleep(3)
combo_wb = xw.Book()
for filesname in combo_data:
    for sheets_num in range(len(combo_data[filesname])):
        try:
            used_rows = combo_wb.sheets[sheets_num].api.UsedRange.Rows.Count + 1
            combo_wb.sheets[sheets_num].range("A"+str(used_rows)).value = combo_data[filesname][sheets_num]
        except IndexError:
            combo_wb.sheets.add(after=combo_wb.sheets[sheets_num-1])
            used_rows = combo_wb.sheets[sheets_num].api.UsedRange.Rows.Count + 1
            combo_wb.sheets[sheets_num].range("A"+str(used_rows)).value = combo_data[filesname][sheets_num]


