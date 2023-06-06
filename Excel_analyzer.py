import pandas as pd
import os


sheets = list(os.listdir('tables'))
tab = list()
for i in range(len(sheets)):
    print(sheets[i])
    data = pd.read_excel('tables/'+sheets[i]).to_dict()
    new_data = dict()
    for k in range(len(data['Nickname.1'])):
        new_data.update({
            data['Nickname.1'][k]: [data['Common'][k], data['Uncommon'][k], data['Rare'][k], data['Epic'][k], data['Legendary'][k], data['Total'][k]]
        })
    tab.append(new_data)

data_frame = {
    'Nicknames': [],
    'Common': [],
    'Uncommon': [],
    'Rare': [],
    'Epic': [],
    'Legendary': [],
    'Total': [],
    'Points': []
}
for i in range(len(tab)):
    for k in range(len(tab[i].keys())):
        name = list(tab[i].keys())[k]
        if data_frame['Nicknames'].count(name) != 0:
            name_list = tab[i][name]
            index = data_frame['Nicknames'].index(name)
            data_frame['Common'][index] += int(tab[i][name][0])
            data_frame['Uncommon'][index] += int(tab[i][name][1])
            data_frame['Rare'][index] += int(tab[i][name][2])
            data_frame['Epic'][index] += int(tab[i][name][3])
            data_frame['Legendary'][index] += int(tab[i][name][4])
            data_frame['Total'][index] += int(tab[i][name][5])
            data_frame['Points'][index] += int(tab[i][name][0]) + int(tab[i][name][1])*4 + int(tab[i][name][2])*9 + int(tab[i][name][3])*20 + int(tab[i][name][4]*50)
        else:
            data_frame['Nicknames'].append(name)
            data_frame['Common'].append(int(tab[i][name][0]))
            data_frame['Uncommon'].append(int(tab[i][name][1]))
            data_frame['Rare'].append(int(tab[i][name][2]))
            data_frame['Epic'].append(int(tab[i][name][3]))
            data_frame['Legendary'].append(int(tab[i][name][4]))
            data_frame['Total'].append(int(tab[i][name][5]))
            data_frame['Points'].append(int(tab[i][name][0]) + int(tab[i][name][1])*4 + int(tab[i][name][2])*9 + int(tab[i][name][3])*20 + int(tab[i][name][4]*50))

z = zip(data_frame['Nicknames'], data_frame['Common'], data_frame['Uncommon'], data_frame['Rare'], data_frame['Epic'], data_frame['Legendary'], data_frame['Total'], data_frame['Points'])
zs = sorted(z, key=lambda tup: tup[7])
data_frame['Nicknames'] = [z[0] for z in zs]
data_frame['Common'] = [z[1] for z in zs]
data_frame['Uncommon'] = [z[2] for z in zs]
data_frame['Rare'] = [z[3] for z in zs]
data_frame['Epic'] = [z[4] for z in zs]
data_frame['Legendary'] = [z[5] for z in zs]
data_frame['Total'] = [z[6] for z in zs]
data_frame['Points'] = [z[7] for z in zs]

df = pd.DataFrame(data_frame)
writer = pd.ExcelWriter('tables/sheet_week.xlsx')
df.to_excel(writer)
writer.save()