from openpyxl import Workbook, load_workbook
from selenium import webdriver
from time import sleep
import re 
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class FinanceBot():
  
  def __init__(self):
    self.driver = webdriver.Chrome( "C:\chromedriver")
    self.info = {
"Enterprise Value to Sales": 0,
"EBIT": 0,
"Sales/Revenue":0,
"Cash":0,
"Total Current Assets":0,
"Short Term Debt":0,
"Current Portion of Long Term Debt":0,
"Total Current Liabilities":0,
"Net Property, Plant & Equipment":0
}
    self.wait = WebDriverWait(self.driver, 10)
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
     #get sales/revenue and EBIT
    sales_revenue = 0 
    ebit = 0 
    try:
        sales_rev_element = self.wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="cr_cashflow"]/div[2]/div/table/tbody/tr[1]/td[2]')))
        sales_revenue = float(sales_revenue_element.text)
        #
        ebit_element = self.wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="cr_cashflow"]/div[2]/div/table/tbody/tr[1]/td[2]')))
        ebit = float(ebit_element.text)
    except:
        self.driver.quit() 

    values["Sales/Revenue"] = sales_revenue
    values['EBIT'] = float(ebit)

def get_latest_cell(self,row_name)->float:
  elmt = self.driver.find_element_by_xpath(f"//td[text()='{row_name}']/following-sibling::td[1]")
  return float(elmt.text)

   
def extract_balance_sheet_data(self, company_code):
  
  self.driver.get(f'https://www.wsj.com/market-data/quotes/HK/XHKG/{company_code}/financials/annual/balance-sheet')
 
  try:
      cash_element = self.wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="cr_cashflow"]/div[2]/div[2]/table/tbody/tr[2]/td[2]')))
      cash = float(cash_element.text)

      total_current_assets_element = self.wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="cr_cashflow"]/div[2]/div[2]/table/tbody/tr[22]/td[1]')))
      total_current_assets = float(total_current_assets_element.text)

      #Expand Liabilities & Shareholders' Equity Section
      liabilities_btn = self.driver.find_element_by_xpath('//*[@id="cr_cashflow"]/div[3]/div[1]')
      liabilities_btn.click()

      short_term_debt_element = self.driver.find_element_by_xpath("//td[text()='Short Term Debt']/following-sibling::td[1]")
      short_term_debt = float(short_term_debt_element.text)


      lt_debt_elmt = self.driver.find_element_by_xpath("//td[text()='Current Portion of Long Term Debt']/following-sibling::td[1]")
      lt_debt = float(lt_debt_elmt.text)

      total_curr_liabilites_elmt = self.driver.find_element_by_xpath("//td[text()='Total Current Liabilities']/following-sibling::td[1]")
      toatl_curr_liabilites = float(toatl_curr_liabilites.text)
      
      net_property_and_plan_equipment = self.get_latest_cell("Net Property, Plant & Equipment")

  except:
      self.driver.quit() 

  '''
  Net Property, Plant & Equipment
  '''
  

  values["Total Current Assets"] = total_current_assets 
  values["Cash"] = cash
  values["Short Term Debt"] = short_term_debt
  values["Current Portion of Long Term Debt"] = lt_debt
  values["Total Current Liabilities"] = toatl_curr_liabilites
  values["Net Property, Plant & Equipment"] = net_property_and_plan_equipment

  

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



