import os
import pandas as pd

# Получение списка таблиц
sheets = list(os.listdir('tables'))
columns = pd.read_excel(os.path.join('tables', sheets[0])).columns
# Объединение таблиц в один список
data_frames = []
for sheet in sheets:
    if sheet.endswith('.xlsx'):
        data_frames.append(pd.read_excel(os.path.join('tables', sheet),
            parse_dates=['First Hunt Time', 'Last Hunt Time'], date_format='%Y-%m-%d %H:%M:%S'))

# Создание словаря для хранения данных
players = {}
# Обработка строк таблиц
for df in data_frames:
    for index, row in df.iterrows():
        user_id = row['User ID']
        dirow = row.to_dict()
        # Создание новой записи для игрока
        if user_id not in players:
            players[user_id] = dirow
        else:
            for key, value in dirow.items():
                if ((type(value) == type(1)) or type(value) == type(1.1)) and key != "User ID":
                    players[user_id][key] += value
                if dirow['First Hunt Time'] < players[user_id]['First Hunt Time']:
                    players[user_id]['First Hunt Time'] = dirow['First Hunt Time']
                if dirow['Last Hunt Time'] > players[user_id]['Last Hunt Time']:
                    players[user_id]['Last Hunt Time'] = dirow['Last Hunt Time']

data_frame = {}
index_list = {columns[i]:[] for i in range(len(columns))}
for key, values in players.items():
    for k, value in values.items():
        index_list[k].append(value)

for i in range(len(columns)):
    data_frame.update({columns[i]:index_list[columns[i]]})
data_frame = pd.DataFrame(data_frame)
writer = pd.ExcelWriter('sheet_week.xlsx')
data_frame.to_excel(writer)
writer._save()
