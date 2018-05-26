from openpyxl import Workbook
import random
import numpy as np
import re
import string
from openpyxl.utils import get_column_letter

def generate_weekly_hours(median,days_per_week):
    return np.round(np.random.normal(median,2,days_per_week)*4)/4

def rng2tuple(rng): # adapted from https://stackoverflow.com/a/12640614
    col,row = re.split('(\D*)',rng)[1:]
    num = 0
    for c in col:
        if c in string.ascii_letters:
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num, int(row)

def get_rng(col,row):
    return get_column_letter(col) + str(row)

def generate_proj_hours(projects):
    shift_length = 8
    proj_cnt = len(projects)
    project_split = shift_length / proj_cnt
    return np.round(np.random.normal(project_split, 1, proj_cnt) * 4) / 4

def build_timecard_table():
    wb = Workbook()

    col_headers = ['Projects', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Total']

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

    # set the top left of the hours table
    top_left_cell = 'C7'

    project_count = random.randint(1, 4)
    timecard_projects = random.sample(projects, project_count)


    for col in range(len(col_headers)):
        header = col_headers[col]
        c,r = rng2tuple(top_left_cell)
        active_col = c + col
        ws.cell(column=active_col, row=r, value= header)
        task_hours = generate_proj_hours(timecard_projects)
        for proj in range(len(timecard_projects)):
            active_row = r + proj + 1
            if header == 'Projects':
                project = timecard_projects[proj]
                ws.cell(column=active_col, row=active_row, value=project)
            elif header == 'Total':
                rbound = get_rng(active_col-1,active_row)
                lbound = get_rng(active_col-5, active_row)
                ws.cell(column=active_col, row=active_row, value='=sum(%s:%s)'%(lbound,rbound))
            else:
                ws.cell(column=active_col, row=active_row, value=task_hours[proj])
    c, r = rng2tuple(top_left_cell)

    # write total for week formula

    tbound = get_rng(active_col, active_row)
    bbound = get_rng(active_col, active_row - proj)
    ws.cell(column=active_col, row=active_row + 1, value='=sum(%s:%s)'%(tbound,bbound))

    wb.save(filename=dest_filename)

build_timecard_table()


