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
    for index,table in enumerate(soup.find_all('td')):
        name = table.find('a')
        if name is not None:
            names.append(name.text)
        elif table is not None:
            if table.text != '\n\n×\t\t\t\t\t\t\t\n◯\t\t\t\t\t\t\t\n△\t\t\t\t\t\t\t\n\n' and table.text != '日程':
                if table.text in input_type:
                    if names[name_index] not in member_data:
                        name_index += 1
                        name_index %= len(names)
                        continue
                    # print(names[name_index])
                    member_data[names[name_index]][url_index+1].append(table.text)
                    name_index += 1
                    name_index %= len(names)

                elif len(columns) > 0 and columns[0] == table.text:
                    name_index += 1
                elif table.text != '\n':
                    columns.append(table.text)
    input_columns = labels.copy()
    for now in columns:
        input_columns.append(now)
    save_data = [input_columns]
    if url_index == 0:
        for name in names:
            current_data = [name]
            if name not in member_data:
                continue
            for label_index in range(len(labels)-1):
                current_data.append(member_data[name][0][label_index])
            for date in member_data[name][url_index+1]:
                current_data.append(date)
            save_data.append(current_data)
    else:
        for name in member_data:
            current_data = [name]
            for label_index in range(len(labels)-1):
                current_data.append(member_data[name][0][label_index])
            date_index = 1
            if len(member_data[name][0][label_index]) == 0:
                date_index = 0
            for date in member_data[name][date_index]:
                current_data.append(date)
            save_data.append(current_data)
    save_name = os.path.join(str(today.year),'{}Q.csv'.format(url_index+start+1))
    df = pd.DataFrame(save_data)
    df.to_csv(save_name,encoding='shift_jis',header=False, index=False)