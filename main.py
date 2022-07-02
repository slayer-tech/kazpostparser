import requests
import json
import pandas as pd
import openpyxl
import time

i = 0

fn = 'posts.xlsx'
df = pd.DataFrame({ 
                "Трек код": [],
                "Направление": [],
                "Дата отправления": [],
                "Статус": [],
                "Имя отправителя": [],
                "Страна отправителя": [],
                "Адресс отправителя":[],
                "Имя получателя": [],
                "Страна получателя": [],
                "Адресс получателя": [],
                "Метод доставки": [],
                "Вес": []
            }, index=[])

df.to_excel('./' + fn)

wb = openpyxl.load_workbook(fn)
ws = wb['Sheet1']

for x in map("{:0>8}".format, range(100000000)):
    x1 = int(x[0])
    x2 = int(x[1])
    x3 = int(x[2])
    x4 = int(x[3])
    x5 = int(x[4])
    x6 = int(x[5])
    x7 = int(x[6])
    x8 = int(x[7])

    x9 = 11 - (x1 * 8 + x2 * 6 + x3 * 4 + x4 * 2 + x5 * 3 + x6 * 5 + x7 * 9 + x8 * 7) % 11

    if (x9 > 9 or x9 < 0):
        x9 = 0

    for departure_type in ['RR', 'RW', 'CP', 'EE', 'CV']:
        URL = "https://track.kazpost.kz./api/v2/" + departure_type + str(x1) + str(x2) + str(x3) + str(x4) + str(x5) + str(x6) + str(x7) + str(x8) + str(x9) + "KZ"

        r = requests.get(URL)

        data = json.loads(r.text)

        print(data['trackid'])

        if 'error' not in data: 
            ws.append([
                    i,
                    data['trackid'], 
                    data['direction'],
                    data['origin']['date'],
                    data['status'],
                    data['sender']['name'],
                    data['sender']['country'],
                    data['sender']['address'],
                    data['receiver']['name'],
                    data['sender']['country'],
                    data['sender']['address'],
                    data['delivery_method'],
                    data['weight']
                ])

            i += 1

            wb.save(fn)
        
print("Finish")
wb.close()