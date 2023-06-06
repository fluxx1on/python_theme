import pyautogui as pg
import time
import discord as ds
import os
import keyboard
import pytesseract as pt
import pandas as pd
import cv2

# Presets

settings = {
    'token':'OTUwMzI5OTIzODkyMDQzNzc2.YiXVtg.DlpKBF0WgIlaRg-W6kcf5rkuegE',
    'bot': 'Counter',
    'id': 962924805894602792,
    'prefix': '%'
}

print('Введите минималку:')
minimal = int(input())
confidence = 0.9
tesseract_config = '-c tessedit_char_whitelist=" 0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz" --oem 3 --psm 7'
pt.pytesseract.tesseract_cmd = r'C:\Windows.old\Users\USER\AppData\Local\Tesseract-OCR\tesseract.exe'

in_folder = 'pg_lists/'
out_folder = 'pg_tables/'
table_name = 'pg_' + str(time.gmtime()[2]) + '_' + str(time.gmtime()[1]) + '.xlsx'
output_sheet_name = 'Main'

# Main
dataset = {'Rank':[],
           'Nickname':[],
           'Points':[],
           'Credit':[],
           'Status':[],
           'Tries':[] }

def creditation(points):
    points = int(points)
    if (minimal < points + 200) and (minimal > points):
        boolean = ((minimal-points)//200) > 0
        return int(boolean)*200 + int(not(boolean))*(minimal-points), 'Free' 
    elif minimal - points > 200:
        return 0, "Kicked"
    return 0, 'Free'

sheets = list(os.listdir(in_folder[:-1:]))
for i in range(len(sheets)):
    img = cv2.imread(in_folder + sheets[i])
    for k in range(round(len(img)/67)):
        img = cv2.imread(in_folder + sheets[i])
        img = img[67*k:67*(k+1), 0:]
        sender_fragment = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        sender_fragment = cv2.threshold(sender_fragment, 127, 255, cv2.THRESH_BINARY_INV)[1]
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
        sender_fragment = cv2.morphologyEx(sender_fragment, cv2.MORPH_CLOSE, kernel)
        sender_fragment = cv2.blur(sender_fragment, (3, 3))
        sender = pt.image_to_string(sender_fragment, lang='eng', config=tesseract_config)
        sender = sender.split()
        if sender != []:
            print(sender)
            rank = sender[0]
            nickname = ''
            for t in sender[1:-2:]:
                nickname += t
            points = sender[-2]
            credit, status = creditation(points)
            if sender[-1][0:2:] == '10':
                tries = '10/10'
            elif sender[-1][-2::] == '10':
                tries = sender[-1][0] + '/' + sender[-1][-2] + sender[-1][-1]
            else:
                tries = sender[-1][0] + '/' + sender[-1][-1]
            dataset['Rank'].append(int(rank))
            dataset['Nickname'].append(nickname)
            dataset['Points'].append(points)
            dataset['Credit'].append(credit)
            dataset['Status'].append(status)
            dataset['Tries'].append(tries)



z = zip(dataset['Rank'], dataset['Nickname'], dataset['Points'], dataset['Credit'], dataset['Status'], dataset['Tries'])
zs = sorted(z, key=lambda tup: tup[0])
dataset['Rank'] = [z[0] for z in zs]
dataset['Nickname'] = [z[1] for z in zs]
dataset['Points'] = [z[2] for z in zs]
dataset['Credit'] = [z[3] for z in zs]
dataset['Status'] = [z[4] for z in zs]
dataset['Tries'] = [z[5] for z in zs]

for i in range(len(dataset['Rank'])):
    dataset['Rank'][i] = str(dataset['Rank'][i])
for i in range(len(dataset['Credit'])):
    dataset['Credit'][i] = str(dataset['Credit'][i])

data = pd.DataFrame(dataset)
writer = pd.ExcelWriter(out_folder+table_name)
data.to_excel(writer, index_label='Rank', sheet_name=output_sheet_name)
writer.save()

# Discord
"""
client = ds.Client()

@client.event
async def on_ready():
    channel = client.get_channel(id=950459413867671582)
    await channel.send(file=ds.File(out_folder+table_name))

client.run(settings['token'])
"""