import tabula as tb
import pandas as pd
import re
import os
import xlsxwriter

# selects current working directory to source pdf file
dir_path = os.path.dirname(os.path.realpath(__file__))
os.chdir(dir_path)
os.system('cls')

print("Job running......") 

workbook = xlsxwriter.Workbook('export.xlsx', {'strings_to_numbers': True})
worksheet = workbook.add_worksheet("Tab1")

pagecount = 2    #PAGE COUNT start must be Manually defined here!
date = str()
price = str()
xrow = 0
errorcount = 0
errorpage = list()

while pagecount <= 71:   #PAGE COUNT MAX must be Manually defined here!
    file = 'dellprice.pdf'
    #file read area should be (70,0,1000,800)
    #area parameter for all pages should be (180,300,380,500)
    df = tb.read_pdf(file, area = (70, 0, 1000, 800), columns = [180,300,380,500], pages = pagecount, pandas_options={'header': [0,1]}, stream=True)[0]
    print(df)
    # print(df[0][0])
    dfrow = 0

    try:
        for row in df.index:
            date = df[0][dfrow]
            price = df[1][dfrow]
            worksheet.write(xrow,0,str(date))
            worksheet.write(xrow,1,str(price))
            xrow += 1
            dfrow += 1
    except:
        errorcount += 1
        errorpage.append(pagecount)

    dfrow = 0

    try:
        for row in df.index:
            date = df[2][dfrow]
            price = df[3][dfrow]
            worksheet.write(xrow,0,str(date))
            worksheet.write(xrow,1,str(price))
            xrow += 1
            dfrow += 1
    except:
        errorcount += 1
        errorpage.append(pagecount)
    pagecount += 1

print("Export Complete")    
print(errorcount , "data block failed to process on page(s)",errorpage)    

workbook.close()