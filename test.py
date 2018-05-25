import random
import numpy as np
import math


def decomposition(i):  # from https://stackoverflow.com/a/10305400
    while i > 0:
        n = random.randint(1, i)
        yield n
        i -= n

def distribute_hours(hrs):
    split = math.modf(hrs)
    whole_hours = list(decomposition(split[1]))
    indx = random.randint(1,len(whole_hours))
    whole_hours[indx-1] += split[0]
    return whole_hours


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


def generate_weekly_hours(median, days_per_week):
    return np.round(np.random.normal(median, 2, days_per_week) * 4) / 4


hours = generate_weekly_hours(8, 5)

weekday_row = 7
weekday_offset = 3
timecard_projects = random.sample(projects, random.randint(1, 4))

for col in enumerate(weekdays):
    print(col[1])
    daily_hours = distribute_hours(hours[col[0]])
    for proj in enumerate(timecard_projects):
        print(daily_hours[proj[0]])