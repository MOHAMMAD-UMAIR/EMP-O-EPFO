from selenium import webdriver
from selenium.webdriver.common.by import By
from openpyxl import Workbook
from bs4 import BeautifulSoup
from openpyxl.styles import PatternFill
from selenium.common.exceptions import NoSuchElementException
import time
import re
import math

def logo():
    print("""   / ____// __ \ / ____// __ \
                / __/  / /_/ // /_   / / / /
                / /___ / ____// __/  / /_/ / 
                /_____//_/    /_/     \____/  
    """)


logo()
# def row_count(driver):
#     # count the number of rows in the table
#     table = driver.find_element(By.ID,"example")
#     rows = table.find_elements(By.TAG_NAME, "tr")
#     print("Number of rows in the table: ",len(rows))
#     return len(rows)

# def table_click(driver,row):
#     # 
#     links = driver.find_elements(By.XPATH, '//*[@id="example"]/tbody/tr[{0}]/td[5]/a',str(row))
#     if links:
#         links[0].click()
