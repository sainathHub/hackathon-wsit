from bs4 import BeautifulSoup
from urllib.request import urlopen
from urllib.request import urlopen
import requests
import pandas as pd
import math
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



URL = 'C:/Users/siddh/Desktop/1.html'

soup = BeautifulSoup(open(URL), "html.parser")

data = []
table = soup.find('table')

table_body = table.find('tbody')

rows = table_body.find_all('tr')
for row in rows:
    cols = row.find_all('td')
    cols = [ele.text.strip() for ele in cols]
    
    data.append([ele for ele in cols if ele]) # Get rid of empty values
   


oldf_data = []

i = 1
while(i < len(data)):
    row = data[i]
    newrow = [row[1], row[2],row[3],row[4],row[5]]
    oldf_data.append(newrow)
    i = i+1

    if  "Plus 1(1)" == row[0] :
        i= i+1 #skipping the hreader and current
        row = data[i]
        oldf_data.append(["Plus 1",row[0],row[1]])
        i=i+2 #skipping the space and current

print(oldf_data)
print()
df = pd.read_excel(r'book1.xlsx',sheet_name="Sheet1")
excel_data = df.values.tolist()
excel_data = clean(excel_data)
print(excel_data)