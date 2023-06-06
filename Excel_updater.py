import pandas as pd
import os
import time

in_folder = 'pg_lists/'
out_folder = 'pg_tables/'
table_name = 'pg_' + str(time.gmtime()[2]) + '_' + str(time.gmtime()[1]) + '.xlsx'
output_sheet_name = 'Main'

dataset = {'Rank':[],
           'Nickname':[],
           'Points':[],
           'Credit':[],
           'Status':[],
           'Tries':[] }

print('Введите минималку:')
minimal = int(input())
sheets = list(os.listdir('pg_tables'))
print('Какая таблица самая новая?')
for i in range(len(sheets)):
    print(str(i) + '. ' + sheets[i])
void = int(input())
illegal = int(not(bool(void)))
data = list()
for i in range(len(sheets)):
    data.append(pd.read_excel('pg_tables/'+sheets[i]).to_dict())

dataset['Rank'] = list(data[void]['Rank'].values())
dataset['Nickname'] = list(data[void]['Nickname'].values())
dataset['Points'] = list(data[void]['Points'].values())
dataset['Credit'] = list(data[void]['Credit'].values())
dataset['Status'] = list(data[void]['Status'].values())
dataset['Tries'] = list(data[void]['Tries'].values())

names = list(data[illegal]['Nickname'].values())
credit = list(data[illegal]['Credit'].values())

have_credit_nicknames = list()
have_credit = list()
for i in range(len(dataset['Nickname'])):
    if str(credit[i]) != "0": 
        have_credit_nicknames.append(names[i])
        have_credit.append(int(credit[i]))

for i in range(len(have_credit_nicknames)):
    try:
        id = dataset['Nickname'].index(have_credit_nicknames[i])
        num = int(dataset['Credit'][id]) + have_credit[i]
        if num > 200:
            print(dataset['Nickname'][id], num)
            dataset['Status'][id] = "Kicked"
            dataset['Credit'][id] = "0"
        else:
            if dataset['Points'][id] > minimal:    
                ret = num - (dataset['Points'][id] - minimal)
                if ret < 0:
                    dataset['Status'][id] = "Free"
                    dataset['Credit'][id] = "0"
                else:
                    dataset['Status'][id] = "Free"
                    dataset['Credit'][id] = str(ret)
            else: 
                ret = num + (minimal - dataset['Points'][id])
                if ret > 200:
                    dataset['Status'][id] = "Kicked"
                    dataset['Credit'][id] = "0"
                else:
                    dataset['Status'][id] = "Free"
                    dataset['Credit'][id] = str(num)
    except ValueError:
        None

d = pd.DataFrame(dataset)
writer = pd.ExcelWriter(out_folder+"new"+table_name)
d.to_excel(writer, index_label='Rank', sheet_name=output_sheet_name)
writer.save()