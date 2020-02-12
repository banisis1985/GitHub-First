import openpyxl as xl
import glob, os
from datetime import datetime
from openpyxl.utils import column_index_from_string
from openpyxl import load_workbook
from pathlib import Path
all_files = glob.glob('d:/vbaniukaitis/Desktop/codingCsv/dataIN/*.xlsx')
latest_xlsx = max(all_files, key=os.path.getctime)
wb = xl.load_workbook(latest_xlsx)
ws = wb.active
mylist= []
datestring = datetime.strftime(datetime.now(), ' %Y_%m_%d %H_%M_%S')
for column in range(1, ws.max_column + 1):
    if ws.cell(1, column).value == 'ExpirationDate':
        stulpelis = column
for row in range(1, ws.max_row +1):
    mylist.append(ws.cell(row,stulpelis).value)
    for index in mylist:
        index = str(index)
        if index != datetime:
            continue
    if index.isalpha():
        continue
    date = datetime.strptime(index, '%Y-%m-%d %H:%M:%S')
    newdate = date.replace(hour=23, minute=59)
    cell = ws.cell(row, stulpelis)
    cell.value = (newdate)
    print(newdate)
wb.save('d:/vbaniukaitis/Desktop/codingCsv/naujas' + datestring + '.xlsx')
print('buvo nuskaitytas failas ' +latest_xlsx)
