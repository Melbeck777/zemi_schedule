import os
from tkinter import E
import pandas as pd
import datetime
import csv
import jpholiday
import time


today = datetime.datetime.today()
input_marks = ['◯', '△','×']

member_file = os.path.join(str(today.year),'member','member.xlsx')
data = pd.read_excel(member_file)
labels = ['学籍番号','学年','研究室','研究班','氏名']
member_data = {}
for index,col in enumerate(data[labels[0]]):
    current_list = []
    for other_label in labels[1:]:
        current_list.append(data.iloc[index][other_label])
    member_data[str(col)] = [current_list,[],[]]

terms = {0:'early',2:'late'}
url_file = os.path.join(str(today.year),'url.csv')
url_data = pd.read_csv(url_file)
start = 0
if today.month > 8:
    start = 2


duration_data = pd.read_csv('{}/{}/duration.csv'.format(today.year,terms[start]))

first_columns = []
strangers = []
for url_index in range(2):
    columns = []
    names = []
    with open('{}/{}/{}Q_data.csv'.format(str(today.year),terms[start], str(start+1+url_index))) as f:
        reader = csv.reader(f)
        for row_index, row in enumerate(reader):
            if row_index < 1:
                continue
            elif row_index == 2:
                for col_index, col in enumerate(row):
                    if col == '日程':
                        continue
                    names.append(col)
            elif row_index > 2:
                flag = False
                for col_index, col in enumerate(row):
                    if flag:
                        flag = True
                        member_data[names[col_index-1]][url_index+1].append(col)
                    elif col not in input_marks and col != ' ':
                        if col == 'コメント':
                            flag = True
                        columns.append(col)
                    else:
                        member_data[names[col_index-1]][url_index+1].append(col)
    # Arrange columns for write.
    input_columns = labels.copy()
    for now in columns:
        input_columns.append(now)
    # Arrange each person data of schedule to match the columns.
    save_data = [input_columns]
    for name in member_data:
        current_data = [name]
        for label_index in range(len(labels)-1):
            current_data.append(member_data[name][0][label_index])
        date_index = 1+url_index
        if len(member_data[name][date_index]) == 0 and  date_index != 1:
            date_index -= 1
        for date in member_data[name][date_index]:
            current_data.append(date)
        save_data.append(current_data)
    if url_index == 0:
        first_columns = columns
    save_name = os.path.join(str(today.year),'{}Q.csv'.format(url_index+start+1))
    df = pd.DataFrame(save_data)
    df.to_csv(save_name,encoding='cp932',header=False, index=False)
    time.sleep(3)

# Find don't input date people
none_data_people = [[],[]]
for person in member_data:
    if len(member_data[person][1]) == 0:
        none_data_people[0].append(person)
    elif len(member_data[person][2]) == 0:
        none_data_people[1].append(person)
non_data_file = os.path.join(str(today.year), terms[start], 'none_data.txt')
# none_data_people[0] = sorted(none_data_people[0])
# none_data_people[1] = sorted(none_data_people[1])
# Output don't input person number.
with open(non_data_file, 'w', encoding='utf-8') as f:
    degree = ''
    for index, element in enumerate(none_data_people):
        f.write('{}Q 未記入者\n'.format(start+1+index))
        for it in element:
            if degree != member_data[it][0][0]:
                degree = member_data[it][0][0]
                if degree == '教員':
                    f.write('{} {}\n'.format(degree,it))
                else:
                    f.write('{:^4} {}\n'.format(degree,it))
            else:
                f.write('     {}\n'.format(it))
        f.write('\n')
    if len(strangers) > 0:
        f.write('学籍番号入力ミス')
    for stranger in strangers:
        f.write('{}\n'.format(stranger))
print('Output "{}".'.format(non_data_file))

# Count people who can join.
number_data = [[[0,x] for x in first_columns[:-1]],[[0,x] for x in first_columns[:-1]]]
print(number_data)
for quarter_index in range(2):
    for columns_index in range(5*5):
        for name in member_data:
            if len(member_data[name][quarter_index+1]) == 0:
                if quarter_index == 0:
                    continue
                elif len(member_data[name][quarter_index]) > 0:
                    if member_data[name][quarter_index][columns_index] == input_marks[0]:
                        number_data[quarter_index][columns_index][0] += 1
            else:
                if member_data[name][quarter_index+1][columns_index] == input_marks[0]:
                    number_data[quarter_index][columns_index][0] += 1
print(number_data[0])
print(number_data[1])
decided_date = [max(number_data[0])[1].split(' '),max(number_data[1])[1].split(' ')]
print(decided_date)
week_days = ['月', '火', '水', '木', '金']
week_index = week_days.index(decided_date[0][0])
start_date = str(duration_data[duration_data.columns[1]]).split('     ')[1].split('\n')[0]
start_date = datetime.datetime.strptime(start_date, '%Y-%m-%d').date()
end_date = str(duration_data[duration_data.columns[2]]).split('     ')[1].split('\n')[0]
end_date = datetime.datetime.strptime(end_date, '%Y-%m-%d').date()

date_first_offset = datetime.timedelta(days=week_index)
start_date += date_first_offset
offset = datetime.timedelta(weeks=1)

presenter_list_file = '{}/{}/presenter.txt'.format(today.year,terms[start])
event_date = []
with open(presenter_list_file, 'r', encoding='utf-8') as f:
    presenters = f.read().split('\n')
    
schedule_output_file = '{}/{}/order.txt'.format(today.year,terms[start])
counts = [0,0]
with open(schedule_output_file, 'w',encoding='utf-8') as f:
    for i in range(len(presenters)):
        while jpholiday.is_holiday_name(start_date):
            start_date += offset
        if start_date < end_date:
            counts[0] += 1
        else:
            counts[1] += 1
        f.write('{:<2}/{:<2} ({}) {} {}\n'.format(start_date.month, start_date.day, decided_date[0][0], decided_date[0][1], presenters[i]))
        start_date += offset
print('Output schedule as "{}".'.format(schedule_output_file))