# I created this program as an experiment to read data out of PDF files. In particular, this file uses a 71 page PDF of Dell stock prices (prior to going private in 2013), reading from pages 2-71 and ignoring the first page which is context on historical splits. The program creates an excel workbook, then executes a loop that works through the PDF page by page. For each page, two dataframes are asynchronously created to read from the two blocks of data. The data is simultaneously written into the Excel file row by row.

import tabula as tb
import pandas as pd
import os
import xlsxwriter

# selects current working directory to source pdf file
dir_path = os.path.dirname(os.path.realpath(__file__))
os.chdir(dir_path)
os.system('cls')

print("Job running......") 

# assigns xlsxwriter excel export object and creates a new tab for export
workbook = xlsxwriter.Workbook('export.xlsx', {'strings_to_numbers': True})
worksheet = workbook.add_worksheet("Tab1")

pagecount = 2         #PAGE COUNT start must be Manually defined here!
date = str()          #declares new variable for the date field from PDF
price = str()         #declares new variable for the price field from PDF
xrow = 0              #initializes row counter for assignment into excel export
errorcount = 0        #initializes error counter variable for when a data frame can't be created
errorpage = list()    #declares list to signify which page(s) an error was found

while pagecount <= 71:   #PAGE COUNT MAX must be manually defined here!
    file = 'dellprice.pdf' #initializes file to be read
    #Tabula's file scanning area should be (70,0,1000,800) for this specific PDF
    #Tabula's area parameter for all pages should be (180,300,380,500) for this specific PDF
    df = tb.read_pdf(file, area = (70, 0, 1000, 800), columns = [180,300,380,500], pages = pagecount, pandas_options={'header': [0,1]}, stream=True)[0]
    # print(df)
    # print(df[0][0])

    try:
        for row in df.index:
            date = df[0][row]    #reads first column of block and assigns to date
            price = df[1][row]    #reads second column of block and assigns to price
            worksheet.write(xrow,0,str(date))    #writes date to excel
            worksheet.write(xrow,1,str(price))    #writes price to excel
            xrow += 1    #increments row in excel for each time loop runs
    except:    #exception adds context to the two error variables if a dataframe cannot be created
        errorcount += 1
        errorpage.append(pagecount)

    # second try statement reads from second block of same data on every page
    try:
        for row in df.index:
            date = df[2][row]
            price = df[3][row]
            worksheet.write(xrow,0,str(date))
            worksheet.write(xrow,1,str(price))
            xrow += 1
    except:
        errorcount += 1
        errorpage.append(pagecount)
        
    pagecount += 1

print("Export Complete")    
print(errorcount , "data blocks failed to process on page(s)",errorpage)    

workbook.close()
