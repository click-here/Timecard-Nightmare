from openpyxl import Workbook
import random

wb = Workbook()

weekdays = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']

projects = ['Redesign Mobile UI',
            'Write unit tests',
            'Raspberry PI Project',
            'Power BI Dashboard',
            'Fuzzy Search Implementation',
            'Dynamics Plugin Development',
            'SCOTUS Opinions Data Science',
            'Build CSS Piano Animation',
            'Honeypot Project']

dest_filename = 'timecard.xlsx'

ws = wb.active
ws.title = "Timecard"

ws['F3'] = 'Employee'
ws['F4'] = 'Week Ending'
ws['F5'] = 'Manager'

weekday_row = 7
weekday_offset = 3
for col in enumerate(weekdays):
    ws.cell(column=col[0] + weekday_offset, row=weekday_row, value=col[1])

timecard_projects = random.sample(projects, random.randint(1, 4))

for proj in enumerate(timecard_projects):
    row = weekday_row + proj[0] + 1
    ws.cell(column=weekday_offset - 1, row=row, value=proj[1])

wb.save(filename=dest_filename)
