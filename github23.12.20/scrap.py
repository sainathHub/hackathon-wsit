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
import numpy as np
import tkinter as tk
    
from selenium import webdriver

import time
import contextlib
import selenium.webdriver as webdriver
import selenium.webdriver.support.ui as ui
from selenium.common.exceptions import NoSuchElementException


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


def create_table(data, header, export_pdf, title):
    # rcolors = plt.cm.BuPu(np.full(len(row_headers), 0.1))
    # ccolors = plt.cm.BuPu(np.full(len(column_headers), 0.1)
    # ccolors = ["#56b5fd" for _ in range(len(header))]
    fig, ax =plt.subplots(figsize=(12,4))
    plt.title(title)
    ax.axis('tight')
    ax.axis('off')
    print("____________________________________________________")
    print(header, data)
    if len(data)==0:
        data = [["*" for _ in range(len(header))],]

        the_table = ax.table(cellText=data,colLabels=header,loc='center')
        export_pdf.savefig(fig, bbox_inches='tight')
        return

    for row in data:
        if len(row) < len(header):
            row.extend([" " for _ in range(len(header)-len(row))])
        elif "Coorperate" in row[1]:
            row[1] = "Coorperate Banking"
    
    the_table = ax.table(cellText=data,colLabels=header,loc='center')
    export_pdf.savefig(fig, bbox_inches='tight')
    

def compare_report(oldfdata, exceldata):

   
    hash_map = {}
    header = ['Type','ReasonForGAndE','name','Organization','Role','mappingindex']
    
    for i, row in enumerate(exceldata):
        hash_map[gh(row)] = [False, i]


    outlier = []
    keys = hash_map.keys()
    for i, row in enumerate(oldfdata):
        hashval =  gh(row)

        if  hashval in keys:
            row[0][-1] = (hash_map[hashval][1]+1)
            hash_map[hashval][0] = True
        
        else:
            row[0].append("outlier(Not found).")
            outlier.append(row)


    NotFound = []
    for row in keys:

        if hash_map[row][0] == False:
            ind = hash_map[row][1]
            NotFound.append(exceldata[ind])

 
    expanded_oldf = []
    for row in oldfdata:
        expanded_oldf.extend(row)

    outlierexp = []
    for row in outlier:
        outlierexp.extend(row)

    NotFoundexp = []
    for row in NotFound:
        NotFoundexp.extend(row)



 


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





    with PdfPages(r"C:\Users\siddh\Downloads\Report.pdf") as export_pdf,open('report.csv', mode='w') as reportFile:

  

        title = "The index mapping of recepients from web to sheet"
        create_table(oldfdata, header, export_pdf,title)
        csvWriter = csv.writer(reportFile)
        csvWriter.writerow(header)
        csvWriter.writerows(expanded_oldf)
    


        fig =  plt.figure(figsize=(12,6))
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
        a.set_title("  ")
        a.pie(list(pluscount.values()), labels = list(pluscount.keys()), autopct="%0.f%%")
    
        a = plt.subplot(235)
        a.set_title("     ")
        a.pie(list(typeCount.values()), labels = list(typeCount.keys()), autopct="%0.f%%")

        plt.subplot(236).pie(list(rolecount.values()), labels = list(rolecount.keys()), autopct="%0.f%%")
        plt.grid(True)
        export_pdf.savefig()
        
    


        #nd
        fig = plt.figure(figsize = (12,8))
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

        plt.grid(True)
        export_pdf.savefig()
        
        print(NotFound, outlier)
        title = "Missed Recepients"
        create_table(NotFoundexp, header, export_pdf, title)
        csvWriter.writerows([[],["MISSED"]])
        csvWriter.writerows(NotFound)



        title = "Incorrent data of Recepients"
        create_table(outlierexp, header, export_pdf, title)
        csvWriter.writerows([[],["Outliers"]])
        csvWriter.writerows(outlierexp)

      # simple_demo.py



def get_pluscount(rows, i):
    pluscount = 0
    for row in rows[i+1:]:
        if len(row) < 4:
            pluscount = pluscount + 1
        else :
            break
    return pluscount

def main_script():
    URL = r'C:\Users\siddh\Documents\github23.12.20\static\xlsx\saved.html'
    soup = BeautifulSoup(open(URL), "html.parser")
    data = []
    div = soup.find(id="dvparsed")
    table = div.find("table")
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


    df = pd.read_excel(r'C:\Users\siddh\Documents\github23.12.20\static\xlsx\Book2.xlsx',sheet_name="Sheet1")
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

    for ind, row in enumerate(new_excel):
        row[0].append(ind)
    for ind, row in enumerate(oldf_data):
        row[0].append(ind)

    rInt = random.randint(0,len(new_excel)-1)
    new_excel.pop(rInt)
    

    rInt = random.randint(0,len(oldf_data)-1)
    oldf_data.pop(rInt)
    rInt = random.randint(0,len(oldf_data)-1)
    oldf_data.pop(rInt)

  

    print(new_excel)
    print(oldf_data)

    compare_report(oldf_data, new_excel)
    import os
    filename = "Report.pdf"
    os.startfile(filename)




def open_browser():
    
    executable_path = (r"C:\Program Files\Google\Chrome Beta\Application\chrome");
    browser = webdriver.Chrome()
    URL = 'file:///C:/Users/siddh/Documents/github23.12.20/final.html'
    search = browser.get(URL) 
    
    while True:
        try:
            elem = browser.find_element_by_css_selector("#dvparsed.tablelizer-table")
        except NoSuchElementException:  #spelling error making this code not work as expected
            print("not yet")
        else:
            print("hurray found element",elem)
            break
        time.sleep(1)

    time.sleep(100)


# main_script()

# from selenium import webdriver

# driver = webdriver.Chrome(executable_path=r"C:\Program Files (x86)\Selenium\chromedriver.exe")

# driver.get("http://www.example.com")
# with open('page.html', 'w') as f:
#     f.write(driver.page_source)


root = tk.Tk() 
canvas1 = tk.Canvas(root, width = 300, height = 300)
canvas1.pack() 

exportButton = tk.Button(root, text='export PDF',command=main_script, bg='brown', fg='white')
#openButton = tk.Button (root, text='Export PDF',command=open_browser, bg='brown', fg='white')
canvas1.create_window(150, 150, window=exportButton)   
  

root.mainloop()








































 

