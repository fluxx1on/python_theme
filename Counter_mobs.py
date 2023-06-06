from logging import exception
import pyautogui as pg
from time import sleep
import discord as ds
import os
import keyboard
import pytesseract as pt
import pandas as pd
import cv2

"""
im1 = pg.screenshot()
im1.save('players_dataset/my_screenshot.png')
im2 = pg.screenshot('players_dataset/my_screenshot.png')

bot = commands.Bot(command_prefix = settings['prefix'])

@bot.command() # Не передаём аргумент pass_context, так как он был нужен в старых версиях.
async def h(ctx): # Создаём функцию и передаём аргумент ctx.
    author = ctx.message.author # Объявляем переменную author и записываем туда информацию об авторе.
    if str(author) == "flunkli#1381":
        print(ctx.message)
        await ctx.send(f'Hello, {author.mention}!') # Выводим сообщение с упоминанием автора, обращаясь к переменной author.

bot.run(settings['token'])

"""

settings = {
    'token':'OTUwMzI5OTIzODkyMDQzNzc2.YiXVtg.DlpKBF0WgIlaRg-W6kcf5rkuegE',
    'bot': 'Counter',
    'id': 950329923892043776,
    'prefix': '%'
}

data = 'icons_lords_mobile/rarity/'
rarities = {
    'Common': data+'01 Common.jpg',
    'Uncommon': data+'02 Uncommon.jpg',
    'Rare': data+'03 Rare.jpg',
    'Epic': data+'04 Epic.jpg',
    'Legendary': data+'05 Legendary.jpg',
}

directory = 'player_dataset'
gift_icon = 'icons_lords_mobile/gift_icon.png'
giftfrom = 'icons_lords_mobile/giftfrom.png'
guild_icon = 'icons_lords_mobile/guild_icon.png'
open_icon = 'icons_lords_mobile/open_icon.png'
text_icon = 'icons_lords_mobile/text_icon.png'

output_sheet_name = 'Main'
confidence = 0.85
tesseract_config = '-c tessedit_char_whitelist=" 0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz" --oem 3 --psm 7'
pt.pytesseract.tesseract_cmd = r'C:\Windows.old\Users\USER\AppData\Local\Tesseract-OCR\tesseract.exe'

def rarity_check(region):
    keys = [i for i in rarities.keys()]
    for i in range(5):
        test = pg.locateOnScreen(rarities[keys[i]], grayscale=True, region=region, confidence=confidence-0.3)
        if test is not None:
            return i

def ExceptBeginScreen(conf, counter):
    print('Exception: Вернитесь на основной экран крепости или перейдите на карту королевства')
    if counter >= 3:
        try:
            p = int(input('Введите точность целым числом от 50 до 100, где 100 обозначает идентичность, а ' + str(conf*100) + 'является предыдущим показателем: '))
            if p >= 50 and p <= 100:
                return int(p)/100
            else:
                print('клоун')
                ExceptBeginScreen(conf, counter)
        except:
            print('ты себя то не обманывай')
            ExceptBeginScreen(conf, counter)
    else:
        return conf

def begin(conf, counter):
    guild_coords = pg.locateCenterOnScreen(guild_icon, confidence=conf, grayscale=True)
    if guild_coords is not None:
        pg.click(guild_coords)
        sleep(1)
        pg.click(pg.locateCenterOnScreen(gift_icon, confidence=conf, grayscale=True))
        sleep(1)
        if pg.locateCenterOnScreen(text_icon, confidence=conf, grayscale=True) is not None:
            pg.click(pg.locateCenterOnScreen(text_icon, confidence=conf, grayscale=True))
        sleep(1)
        giftfrom_appear = pg.locateOnScreen(giftfrom, confidence=conf, grayscale=True)
        open_icon_appear = pg.locateCenterOnScreen(open_icon, confidence=conf, grayscale=True)
        return (giftfrom_appear, open_icon_appear)
    else:
        conf1 = ExceptBeginScreen(conf, counter)
        sleep(3)
        counter += 1
        if conf1 == conf:
            begin(conf1, counter)
        else:
            begin(conf1, 0)


sleep(3)
ga, oa = begin(confidence, 0)
name_region = (ga.left + ga.width-1, ga.top-5, ga.width*2, ga.height+10)
Nicknames = []
hunting = []
rarity_region = (round(ga.left - ga.width*0.8), ga.top-3*ga.height, ga.width*2, ga.height*3)

while 1:
    img = pg.screenshot(region=name_region)
    img.save('my_screenshot.png')
    sender_fragment = cv2.imread('my_screenshot.png')
    sender_fragment = cv2.cvtColor(sender_fragment, cv2.COLOR_BGR2GRAY)
    sender_fragment = cv2.threshold(sender_fragment, 127, 255, cv2.THRESH_BINARY_INV)[1]
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
    sender_fragment = cv2.morphologyEx(sender_fragment, cv2.MORPH_CLOSE, kernel)
    sender_fragment = cv2.blur(sender_fragment, (3, 3))
    sender = pt.image_to_string(sender_fragment, lang='eng', config=tesseract_config)[:-1]
    rarity = rarity_check(rarity_region)
    if rarity is None:
        break
    if Nicknames.count(sender) != 1:
        Nicknames.append(sender)
        hunting.append([0 for i in range(5)])
        hunting[-1][rarity] = 1
    else:
        hunting[Nicknames.index(sender)][rarity] += 1 
    if pg.locateCenterOnScreen(open_icon, confidence=confidence, grayscale=True) == oa:
        pg.click(oa)
        sleep(1)
    else: None
    pg.click(oa)
    sleep(1)
    if keyboard.is_pressed('shift') == True:
        break

#excel

z = zip(Nicknames, hunting)
zs = sorted(z, key=lambda tup: tup[0])
Nicknames = [z[0] for z in zs]
hunting = [z[1] for z in zs]

folder = 'tables/'
sheets = list(os.listdir('tables'))
max_num = 0
for i in range(len(sheets)):
    num = int(((str(sheets[i]).replace('sheet', ' ')).replace('.xlsx', ' ')).strip(' '))
    if num > max_num: max_num = num
table_name = 'sheet' + str(int(max_num)+1) + '.xlsx'
Nicknames.append('Total')
tab = {'Nickname': Nicknames}
for i in range(5):
    arr = []
    keys = [c for c in rarities.keys()]
    for k in hunting:
        arr.append(k[i])
    arr.append(sum(arr))
    tab.update({keys[i]: arr})
summa = list()
for i in hunting:
    summa.append(sum(i))
summa.append(sum(summa))
tab.update({'Total': summa})
df = pd.DataFrame(tab)
"""df['Total'] = df.sum(axis='columns', numeric_only=True)
df = df.sort_index(key=lambda x: x.str.lower())
df = df.append(df.sum().rename('-- Total --'))"""

writer = pd.ExcelWriter(folder+table_name)
df.to_excel(writer, index_label='Nickname', sheet_name=output_sheet_name)
writer.save()
"""workbook = writer.book
worksheet = writer.sheets[output_sheet_name]
worksheet.set_column('A:G', 15)"""
    

#discord
client = ds.Client()

@client.event
async def on_ready():
    channel = client.get_channel(id=950459413867671582)
    await channel.send(file=ds.File(folder+table_name))

client.run(settings['token'])