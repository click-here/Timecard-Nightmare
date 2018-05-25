from openpyxl import Workbook
import random
import numpy as np

def generate_weekly_hours(median,days_per_week):
    return np.round(np.random.normal(median,2,days_per_week)*4)/4

wb = Workbook()

weekdays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']

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

hours = generate_weekly_hours(8,5)

weekday_row = 7
weekday_offset = 3
timecard_projects = random.sample(projects, random.randint(1, 4))

for col in enumerate(weekdays):
    ws.cell(column=col[0] + weekday_offset, row=weekday_row, value=col[1])
    for proj in enumerate(timecard_projects):



for proj in enumerate(timecard_projects):
    row = weekday_row + proj[0] + 1
    ws.cell(column=weekday_offset - 1, row=row, value=proj[1])

wb.save(filename=dest_filename)
