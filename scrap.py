from bs4 import BeautifulSoup
import requests
import pandas as pd
from pandas import ExcelWriter
import random
import math
import csv
from collections import Counter
import matplotlib.pyplot as plt
from fpdf import FPDF 
from matplotlib.backends.backend_pdf import PdfPages


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
    header = ['Type',	'ReasonForGAndE','name','Organization','Role','mappingindex']
    
    for i, row in enumerate(exceldata):
        hash_map[gh(row)] = [False, i]


    outlier = []
    keys = hash_map.keys()
    for i, row in enumerate(oldfdata):
        hashval =  gh(row)

        if  hashval in keys:
            row[0].append(hash_map[hashval][1]+1)
            hash_map[hashval][0] = True
        
        else:
            row[0].append("outlier(Not found).")
            outlier.append(row)


    NotFound = []
    for row in keys:

        if hash_map[row][0] == False:
            ind = hash_map[row][1]
            NotFound.append(exceldata[ind])
 


    expanded = []
    for row in oldfdata:
        expanded.extend(row)
    print(expanded)

    with open('report.csv', mode='w') as reportFile:
        csvWriter = csv.writer(reportFile)
        csvWriter.writerow(header)

        csvWriter.writerows(expanded)

        csvWriter.writerows([[],["Outliers"]])
        csvWriter.writerows(outlier)

        csvWriter.writerows([[],["MISSED"]])
        csvWriter.writerows(NotFound)

    pluscount = {}
    rolecount = {}
    typeCount = {}
    missedCount = len(NotFound)
    missedDic = {"Parsed in web" : len(exceldata)-missedCount, "missed in web": missedCount}
    outlierCount = len(outlier)
    outlierDic = {"valid in web" : len(oldfdata)-outlierCount, "Outliers in web": outlierCount}

    pluscount = Counter(["plus 1(" +str(len(row)-1)+")" for row in exceldata])
    typeCount = Counter([row[0][0] for row in exceldata])
    rolecount = Counter([row[0][4] for row in exceldata])



#next line
    
    #plt.show()

    
    data = [header[:-1],]
    
    for row in NotFound:
        data.extend(row)

    spacing=1
    pdf = FPDF()
    pdf.set_font("Arial", size=12)
    pdf.add_page()
    pdf.cell(200, 10, txt="Welcome to, Validation and correctness Report of GAndE recipients data.", ln=1, align="C")

    pdf.cell(100, 10, txt=" ", ln=1, align="Left")
    pdf.cell(100, 10, txt="mapping index to excel", ln=1, align="Left")
    col_width = pdf.w / 4.5
    row_height = pdf.font_size
    for row in data:
        for item in row:
            pdf.cell(col_width, row_height*spacing, txt=item, border=1)
        pdf.ln(row_height*spacing)
   #add_table()
    pdf.cell(100, 10, txt="Visualiztion of data", ln=1, align="C")


    for row in data:
        for item in row:
            pdf.cell(col_width, row_height*spacing,
                     txt=item, border=1)
        pdf.ln(row_height*spacing)
   #add_table()
    pdf.output('Report.pdf') 
    plt.grid(True)

    fig =  plt.figure(figsize=(12,24))
    fig.suptitle("Analysis of recipient data.", fontsize=16)
    
    axes = plt.subplot(231)
    axes.set_title("plus 1's of representative")
    axes.bar(list(pluscount.keys()),list(pluscount.values()), width = 0.4)   

    
    axes = plt.subplot(232)
    axes.set_title("types of memberss")

    axes.bar(list(typeCount.keys()),list(typeCount.values()))

    axes = plt.subplot(233)
    axes.set_title("various roles")
    axes.bar(list(rolecount.keys()),list(rolecount.values()),width = 0.4)
    
    a = plt.subplot(234)
    a.set_title("Pie graphs for above")
    a.pie(list(pluscount.values()), labels = list(pluscount.keys()), autopct="%0.f%%")
    
    a = plt.subplot(235)
    a.set_title("     ")
    a.pie(list(typeCount.values()), labels = list(typeCount.keys()), autopct="%0.f%%")

    plt.subplot(236).pie(list(rolecount.values()), labels = list(rolecount.keys()), autopct="%0.f%%")
    #plt.show()

    export_pdf = PdfPages("Report.pdf")
    export_pdf.savefig()
    plt.close()

    #nd
    fig = plt.figure(figsize = (10,20))
    fig.suptitle("missed and outlier graphs")
    a = plt.subplot(221)
    a.bar(list(missedDic.keys()),list(missedDic.values()), width=0.4)
    a.set_title("No: of missed recipients")
    a=plt.subplot(222)
    a.bar(list(outlierDic.keys()),list(outlierDic.values()), width=0.4)
    a.set_title("number of wrong recipients")

    a = plt.subplot(223)
    a.pie(list(missedDic.values()),labels = list(missedDic.keys()),autopct="%0.f%%")
    a.set_title("Pie charts for above")
    a=plt.subplot(224)
    a.pie(list(outlierDic.values()),labels=list(outlierDic.keys()), autopct="%0.f%%")
    a.set_title("number of wrong recipients")

    export_pdf.savefig()
    plt.close()
    

      # simple_demo.py



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



rInt = random.randint(0,len(new_excel)-1)
new_excel.pop(rInt)
print(new_excel)

rInt = random.randint(0,len(oldf_data)-1)
oldf_data.pop(rInt)
print(oldf_data)




compare_report(oldf_data, new_excel)