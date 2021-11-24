import xlrd
import json
import pandas 
import random
from pathlib import Path

def randomSpeech(fileName):
    if (Path(fileName).suffix) == '.xlsx':
        with open('table.json', 'w', encoding='utf-8') as file:
            pandas.read_excel(fileName).to_json(file, force_ascii=False, orient='columns')

        with open('table.json', encoding = 'UTF-8') as file:
            data = json.load(file)
        print(random.choice(list(data["Столбец 1"].values())), random.choice(list(data["Столбец 2"].values())), random.choice(list(data["Столбец 3"].values())), random.choice(list(data["Столбец 4"].values())), random.choice(list(data["Столбец 5"].values())))
    
    elif (Path(fileName).suffix) == '.json':
        with open(fileName, encoding = 'UTF-8') as file:
            data = json.load(file)
        print(random.choice(list(data["Столбец 1"].values())), random.choice(list(data["Столбец 2"].values())), random.choice(list(data["Столбец 3"].values())), random.choice(list(data["Столбец 4"].values())), random.choice(list(data["Столбец 5"].values())))

randomSpeech('table.xlsx')
# randomSpeech('table.json')
