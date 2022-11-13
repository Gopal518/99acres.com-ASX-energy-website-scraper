# -*- coding: utf-8 -*-
"""
Created on Wed Aug  3 11:00:31 2022

@author: 44814
"""

import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
import pandas as pd
chrome_options = Options()
chrome_options.add_argument("--headless")


driver = webdriver.Chrome('C:/Users/44814/Downloads/chromedriver.exe', options=chrome_options)
date = '20220708'
url = 'https://www.asxenergy.com.au/wrap_au/' + date
driver.get(url)
time.sleep(5)

count = 0
table_list_name = ['Calendar Year Price','Base Quarters','Cap Quarters','Options Volatility','Futures and Caps Open Interest'
  ,'Options Open Interest','Historical Trade Volumes']
table_list =  ['Calendar Year Price','Base Quarters','Cap Quarters','Options Volatility','Futures and Caps Open Interest'
  ,'Options Open Interest','Historical Trade Volumes']
tables = WebDriverWait(driver,5).until(EC.presence_of_all_elements_located((By.CLASS_NAME, "dataset")))
for table in tables:
    # print(table_list[count])
    table_list[count] = pd.read_html(table.get_attribute('outerHTML'))
    count +=1


with pd.ExcelWriter(date +'.xlsx') as writer:
    for i,table in enumerate(table_list):
        print(table)
        table[0].to_excel(writer, sheet_name= str(table_list_name[i]))
        
        
        
    
