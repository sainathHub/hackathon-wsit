from bs4 import BeautifulSoup
from urllib.request import urlopen
from urllib.request import urlopen
import requests
import pandas as pd
from pandas import ExcelWriter

import math
import csv

def clean(data):
    new_data = []
    for row in data:
        if type(row[0]) != str and math.isnan(row[0]):
            break        
        row = [val for val in row if not val!=val]
        if  len(row)>3:
            row[2]=row[2] + " " + row[3]
            row.pop(3)

        new_data.append(row)

    return new_data

def gh(row):
    ret = []
    if 'plus' in row[0][0]:
         ret.append(row[0][1])
         for i in range(1,len(row)):
            ret.append(row[i][0]+row[i][1])
    else:
        ret.append(row[0][0])
        for i  in range(1,len(row)):
            ret.append(row[i][0]+row[i][1])
        
    return tuple(ret)


def compare_report(oldfdata, exceldata):
    hash_map = {}



    header = ['Type',	'ReasonForGAndE',	'name'	,'Organization'	,'Role','location in web sheet']
    
    for i,row in enumerate(oldfdata):
        hash_map[gh(row)] = (True, i)

    keys = hash_map.keys()
    for i, row in enumerate(exceldata):
        if  gh(row)  in keys:
            row[0].append(hash_map[gh(row)][1]+1)
    
        else :
            row[o].append('Missing')

   # exceldata.insert(0, header)
    

    expanded = []
    for row in exceldata:
        expanded.extend(row)

    df1 = pd.DataFrame(expanded,
                   columns=header)

    with ExcelWriter(r'C:\Users\siddh\Documents\github23.12.20\out.xlsx') as writer:
        df1.to_excel(writer)


def get_pluscount(rows, i):
    pluscount = 0
    for row in rows[i+1:]:
        if len(row) < 4:
            pluscount = pluscount + 1
        else :
            break
    return pluscount

URL = r'C:\Users\siddh\Documents\github23.12.20\static.html'
soup = BeautifulSoup(open(URL), "html.parser")
data = []
table = soup.find('table')
table_body = table.find('tbody')

rows = table_body.find_all('tr')
for i, row in enumerate(rows):
    cols = row.find_all('td')
    cols = [ele.text.strip() for ele in cols]
    data.append([ele for ele in cols if ele]) # Get rid of empty values


data.remove(data[0])
for i,row in enumerate(data):
    if row[0] == "Forename":
        data.remove(row)

oldf_data = []

i=0
while i < len(data):
    entry = []
    if "plus" in data[i][0] :
        entry.append(data[i][1:-3])
    pluscount = get_pluscount(data, i)
  
    i = i+1
    if pluscount > 0:
        for j in range(pluscount):
            entry.append(["Plus 1", data[i][0], data[i][1]])
            i=i+1
    oldf_data.append(entry)

print(oldf_data)

df = pd.read_excel(r'C:\Users\siddh\Documents\github23.12.20\Book1.xlsx',sheet_name="Sheet1")
excel_data = df.values.tolist()
excel_data = clean(excel_data)

new_excel = []

data = excel_data
i=0
while i < len(data):
    entry = []
    if "Plus" not in data[i][0] :
        entry.append(data[i] )
    pluscount = get_pluscount(data, i)
  
    i = i+1
    if pluscount > 0:
        for j in range(pluscount):
            entry.append(data[i])
            i=i+1
    new_excel.append(entry)



print(new_excel)

compare_report(oldf_data, new_excel)