from openpyxl import Workbook
from openpyxl import load_workbook
import datetime

from datetime import date
import calendar
mydays = (calendar.day_name)
pageNumber = 6

people = [
    'Hehe1',
    'Hehe2',
    'Hehe3'
]

wb = Workbook()
ws = wb.active

for i in range(1, len(people) + 1):
    ws.merge_cells(None, i,i,i,i+1)
    ws.cell(row=1, column=i).value = people[i-1]

for i in range(pageNumber):
    wb.create_sheet(mydays[i])
wb.save('template.xlsx')
