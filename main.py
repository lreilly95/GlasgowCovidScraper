import requests
import openpyxl
from datetime import date
from bs4 import BeautifulSoup as bs

# Get website HTML and load excel file
URL = 'https://www.gov.scot/publications/coronavirus-covid-19-daily-data-for-scotland/'
page = requests.get(URL)
workbook = openpyxl.load_workbook('GCC2.xlsx')
worksheet = workbook.get_sheet_by_name('Sheet1')
today=date.today()

# Get table from website
soup = bs(page.content)
table=soup.find('tbody')

# Find all span elements on page
spanList = soup.find_all('span')
spanTextList = []
indices=[]

# Get contents of span elements
for span in spanList:
    spanTextList.append(span.text)

# Get index of span element containing Greater Glasgow and Clyde
for i, elem in enumerate(spanTextList):
    if 'Greater Glasgow and Clyde' in elem:
        indices.append(i)

# Column indices
dateCol = 0
casesCol = 1
activeCol = 2

# Cases and active cases relative to health board name
cases = spanTextList[indices[1]+1]
active = spanTextList[indices[1]+4]

# Set current row to first empty row
currentRow=worksheet.get_highest_row()

# Write data to worksheet
worksheet.cell(row=currentRow,column=datesCol).value = today.strftime("%d-%b-%Y")
worksheet.cell(row=currentRow,column=casesCol).value = spanTextList[indices[1]+1]
worksheet.cell(row=currentRow,column=datesCol).value = spanTextList[indices[1]+4]
