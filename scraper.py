from openpyxl import Workbook, load_workbook
from selenium import webdriver

# financials_url = 'https://www.wsj.com/market-data/quotes/HK/XHKG/1810/financials'
PATH = "C:\chromedriver.exe"

# wb = Workbook()
wb = load_workbook(filename="copy_hk_stock.xlsx")
# print(wb.sheetnames)
sheet = wb.active
sheet["B1"] = "Hello"

browser = webdriver.Chrome(PATH)

# sheet = workbook.active

#Take categories and iterate over them. 
wb.save(filename="copy_hk_stock.xlsx")