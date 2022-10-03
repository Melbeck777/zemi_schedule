import pandas as pd
import os
import datetime
import requests
from bs4 import BeautifulSoup
import hashlib

today = datetime.date.today()

member_file = os.path.join(str(today.year),'member','member.xlsx')
data = pd.read_excel(member_file)
labels = ['学籍番号','学年','研究室','研究班','氏名']
member_data = {}
for index,col in enumerate(data[labels[0]]):
    current_list = []
    for other_label in labels[1:]:
        current_list.append(data.iloc[index][other_label])
    member_data[str(col)] = [current_list,[],[]]


url_file = os.path.join(str(today.year),'url.csv')
url_data = pd.read_csv(url_file)
start = 0
if today.month > 8:
    start = 2

input_type = ['×','○','△' ]

for url_index in range(2):
    columns = []
    names = []
    name_index = 0
    current_url = url_data.iloc[start+url_index]['url']
    hash_name = hashlib.md5(current_url.encode()).hexdigest()+'.html'
    try:
        f = open(hash_name, encoding='utf-8')
        html = f.read()
    except:
        r = requests.get(current_url)
        html = r.text
        with open(hash_name, 'w', encoding='utf-8') as f:
            f.write(r.text)
    soup  = BeautifulSoup(html, 'html.parser')
    current_name = ''
    comment_flag = False
    for index,table in enumerate(soup.find_all('td')):
        name = table.find('a')
        if name is not None:
            names.append(name.text)
        elif table is not None:
            if table.text != '\n\n×\t\t\t\t\t\t\t\n◯\t\t\t\t\t\t\t\n△\t\t\t\t\t\t\t\n\n' and table.text != '日程':
                if comment_flag:
                    member_data[names[name_index]][url_index+1].append(table.text)
                    name_index += 1
                    name_index %= len(names)
                elif table.text in input_type:
                    if names[name_index] not in member_data:
                        name_index += 1
                        name_index %= len(names)
                        continue
                    # print(names[name_index])
                    member_data[names[name_index]][url_index+1].append(table.text)
                    name_index += 1
                    name_index %= len(names)
                elif table.text != '\n':
                    if table.text == 'コメント':
                        comment_flag = True
                    columns.append(table.text)
    input_columns = labels.copy()
    for now in columns:
        input_columns.append(now)
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
    save_name = os.path.join(str(today.year),'{}Q.csv'.format(url_index+start+1))
    df = pd.DataFrame(save_data)
    df.to_csv(save_name,encoding='cp932',header=False, index=False)



none_data_people = [[],[]]
for person in member_data:
    if len(member_data[person][1]) == 0:
        none_data_people[0].append(person)
    elif len(member_data[person][2]):
        none_data_people[1].append(person)


with open('none_data.txt', 'w', encoding='utf-8') as f:
    for index, element in enumerate(none_data_people):
        f.write('{}Q 未記入者\n'.format(start+1+index))
        for it in element:
            f.write('{}\n'.format(it))
        f.write('\n') 