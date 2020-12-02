import xlrd
from selenium import webdriver
financials_url = 'https://www.wsj.com/market-data/quotes/HK/XHKG/1810/financials'
browser = webdriver.Chrome()

workbook = xlrd.open_workbook('./')

worksheet = workbook.sheet_by_index(0)
first_row = [] # The row where we stock the name of the column
for col in range(worksheet.ncols):
    first_row.append( worksheet.cell_value(0,col) )
# tronsform the workbook to a list of dictionnary
data =[]
for row in range(1, worksheet.nrows):
    elm = {}
    for col in range(worksheet.ncols):
        elm[first_row[col]]=worksheet.cell_value(row,col)
    data.append(elm)