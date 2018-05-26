from openpyxl import Workbook
import random
import numpy as np
from openpyxl.utils import get_column_letter

def generate_weekly_hours(median,days_per_week):
    return np.round(np.random.normal(median,2,days_per_week)*4)/4

def col2num(col): # from https://stackoverflow.com/a/12640614
    num = 0
    for c in col:
        if c in string.ascii_letters:
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num

wb = Workbook()

weekdays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday','Total']

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

def get_rng(col,row):
    return get_column_letter(col) + str(row)

def generate_proj_hours(projects):
    shift_length = 8
    proj_cnt = len(projects)
    project_split = shift_length / proj_cnt
    return np.round(np.random.normal(project_split, 1, proj_cnt) * 4) / 4

# set the top left of the hours table
top_left_cell = 'C7'


for col in enumerate(weekdays):
    print(col[1])
    ws.cell(column=col[0] + weekday_offset, row=weekday_row, value=col[1])

    project_count = random.randint(1, 4)
    timecard_projects = random.sample(projects, project_count)
    daily_hours = generate_proj_hours(timecard_projects)

    for proj in enumerate(timecard_projects):
        if col[1] != 'Total':
            task_hours = daily_hours[proj[0]]
            ws.cell(column=col[0] + weekday_offset, row=weekday_row+proj[0]+1, value = task_hours)
        else:
            sum_formula = '=sum()'
            ws.cell(column=col[0] + weekday_offset, row=weekday_row + proj[0] + 1, value=task_hours)


wb.save(filename=dest_filename)
