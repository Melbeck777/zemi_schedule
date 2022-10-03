import os
import datetime
import sys
import pandas as pd
import jpholiday

args = sys.argv
week_day = args[1]
decided_time  = args[2]
week_days = ['月', '火', '水', '木', '金']
week_index = week_days.index(week_day)
today = datetime.date.today()

data = pd.read_csv('{}/late/duration.csv'.format(today.year))
start_date = str(data[data.columns[1]]).split(' ')[5]
start_date = datetime.datetime.strptime(start_date, '%Y-%m-%d').date()
print('default date is {}'.format(start_date))
print('week day is {}'.format(week_days[week_index]))

date_first_offset = datetime.timedelta(days=week_index)
start_date += date_first_offset
offset = datetime.timedelta(weeks=1)
today = datetime.date.today()

term = 'early'
if today.month > 8:
    term = 'late'


presenter_list_file = '{}/{}/presenter.txt'.format(today.year,term)
event_date = []

with open(presenter_list_file, 'r', encoding='utf-8') as f:
    presenters = f.read().split('\n')
    
output_file = '{}/{}/order.txt'.format(today.year,term)
with open(output_file, 'w',encoding='utf-8') as f:
    for i in range(len(presenters)):
        while jpholiday.is_holiday_name(start_date):
            start_date += offset
        f.write('{:<2}/{:<2} ({}) {} {}\n'.format(start_date.month, start_date.day, week_day, decided_time, presenters[i]))
        start_date += offset
print('Output schedule as "{}".'.format(output_file))