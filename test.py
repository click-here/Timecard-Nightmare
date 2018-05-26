import random
import numpy as np
import math




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

def generate_proj_hours(projects):
    shift_length = 8
    proj_cnt = len(projects)
    project_split = shift_length / proj_cnt
    return np.round(np.random.normal(project_split, 1, proj_cnt) * 4) / 4



weekday_row = 7
weekday_offset = 3


for col in enumerate(weekdays):
    print(col[1])
    project_count = random.randint(1, 4)
    timecard_projects = random.sample(projects, project_count)
    daily_hours = generate_proj_hours(timecard_projects)
    for proj in enumerate(timecard_projects):
        print(proj[1])
        print(daily_hours[proj[0]])
