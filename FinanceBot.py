from openpyxl import Workbook, load_workbook
from selenium import webdriver
from time import sleep
import re 


class FinanceBot():
  
  def __init__(self):
    self.driver = webdriver.Chrome( "C:\chromedriver")
    self.info = {
"Enterprise Value to Sales": 0,
"EBIT": 0,
"Sales/Revenue":0,
"Cash":0,
"Total Current Assets":0,
"short term debt":0,
"current portion of long term debt":0,
"total current liabilities":0,
"Net Property, Plant & Equipment":0
}
  def format_worksheet(self):
    wb = load_workbook(filename="copy_hk_stock.xlsx")
    # print(wb.sheetnames)
    sheet = wb.active
    sheet["B1"] = "Hello"

    # Clear instruction columns
    sheet.delete_cols(3,4)

    #Populate column names
    col_i = 3
    for name in column_names.keys():
      sheet.cell(row=1, column=col_i).value = name
      col_i += 1

    #Take categories and iterate over them. 
    wb.save(filename="copy_hk_stock.xlsx")


  financials_url = 'https://www.wsj.com/market-data/quotes/HK/XHKG/1810/financials'
  # PATH = "/Users/work/Documents/projects/chromedriver"

  def extract_income_statement_page(self, company_code):
    '''
    Extracts pertinent data from the income statement page
    '''
    self.driver.get(f"https://www.wsj.com/market-data/quotes/HK/XHKG/{company_code}/financials/annual/income-statement")
  
    sleep(1)
    #GET sales 
    sales_revenue = self.driver.find_element_by_xpath('//*[@id="cr_cashflow"]/div[2]/div/table/tbody/tr[1]/td[2]').text
    sales_revenue = float(sales_revenue)
    values["Sales/Revenue"] = sales_revenue

    #Get EBIT
    
    ebit = self.driver.find_element_by_xpath('//*[@id="cr_cashflow"]/div[2]/div/table/tbody/tr[56]/td[2]').text
    values['EBIT'] = float(ebit)
   
def extract_balance_sheet_data(self, company_code):
  
  self.driver.get(f'https://www.wsj.com/market-data/quotes/HK/XHKG/{company_code}/financials/annual/balance-sheet')
  cash = self.driver.find_element_by_xpath('//*[@id="cr_cashflow"]/div[2]/div[2]/table/tbody/tr[3]/td[2]').text
  
  total_current_assets = self.driver.get('//*[@id="cr_cashflow"]/div[2]/div[2]/table/tbody/tr[22]/td[1]').text
  total_current_assets = float(total_current_assets)
  values["Total Current Assets"] = total_current_assets 

  

  pass

  def extract_finance_data(self, company_code):
    self.driver.get(f"https://www.wsj.com/market-data/quotes/HK/XHKG/{company_code}/financials")
    
    company_name = self.driver.find_element_by_class_name('companyName').text
    company_name = company_name[:len(company_name)-2] #Remove Full stop and white space from title
    
    digit_regex = '[+-]?([0-9]*[.])?[0-9]+'
    #Extract data from financials page
    entrprse_value_to_sales = self.driver.find_element_by_xpath('/html/body/div[2]/section[2]/div[2]/div[1]/div[3]/div/div[2]/div[1]/div[1]/table/tbody/tr[7]/td/span[2]/span')
    entrprse_value_to_sales = float(enterprise_value_to_sales.text)
    self.values["Enterprise Value To Sales"] = entrprse_value_to_sales

    #Visit  income statement page
    self.extract_income_statement_page()


    
    #Narrow down list using .contains
    #Split list into billions and decimals 
    #apply regex
    #Check for B
    #standerdize values 


    #TODO Create a separate way to extract Total Current Assets
    #REGEx for getting digits 


      
    


