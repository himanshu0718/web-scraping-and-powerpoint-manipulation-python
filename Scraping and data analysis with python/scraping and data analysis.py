from selenium import webdriver
from selenium.webdriver.common.by import By
import time
from selenium.webdriver.support.ui import Select
import os
import pandas as pd
import win32com.client

driver = webdriver.Chrome()

driver.get("https://www.mcxindia.com/market-data/spot-market-price")
driver.maximize_window()
time.sleep(1)
driver.find_element(By.XPATH,"//div[@class='today']").click()     #clicking recent
time.sleep(1)
driver.find_element(By.ID,"ctl00_cph_InnerContainerRight_C004_ddlSymbols_Input").click()  #clicking commodity
time.sleep(1)

driver.find_element(By.ID,"ctl00_cph_InnerContainerRight_C004_ddlSymbols_Input").send_keys("G")   #selecting GOLD as input  # entering G
time.sleep(1)
driver.find_element(By.XPATH,"(//li[@class='rcbItem'])[18]").click()  #selecting GOLD from index 18
time.sleep(1)

Select(driver.find_element(By.XPATH,"//select[@id='cph_InnerContainerRight_C004_ddlSession']")).select_by_index(2) #selecting session 2 from dropdown

# from calender
driver.find_element(By.ID,"txtFromDate").click()
year = Select(driver.find_element(By.XPATH,"//select[@title='Change the year']"))
time.sleep(1)
year.select_by_visible_text("2023")
month = Select(driver.find_element(By.XPATH,"//select[@title='Change the month']"))
time.sleep(1)
month.select_by_visible_text("November")
driver.find_element(By.XPATH,"//a[@title='Select Wednesday, Nov 1, 2023']").click()

# to calender
driver.find_element(By.ID,"txtToDate").click()
year = Select(driver.find_element(By.XPATH,"//select[@title='Change the year']"))
time.sleep(1)
year.select_by_visible_text("2024")
month = Select(driver.find_element(By.XPATH,"//select[@title='Change the month']"))
time.sleep(1)
month.select_by_visible_text("January")
driver.find_element(By.XPATH,"//a[@title='Select Wednesday, Jan 24, 2024']").click()
driver.find_element(By.XPATH,"//a[@id='btnShowArchive']").click()
time.sleep(1)
driver.find_element(By.XPATH,"//a[@id='cph_InnerContainerRight_C004_lnkExportToExcel']").click()
time.sleep(5)

#fetching latest file in download directory
download_directory = "C:\\Users\\himan\\Downloads" 
latest_file = max([download_directory + "\\" + f for f in os.listdir(download_directory)], key=os.path.getctime) 

#Step 6: Open the Excel file using win32com
excel = win32com.client.Dispatch("Excel.Application")
workbook = excel.Workbooks.Open(latest_file)
worksheet = workbook.ActiveSheet

# Step 6a: Count the total number of rows in the table
total_rows = worksheet.UsedRange.Rows.Count
print("a.)- Total Rows: ",total_rows)

# Step 6b: Finding date with the highest "Spot Price (Rs.)"
max_spot_price = worksheet.Cells(2, 4).Value  # As 2nd row, 4th column contains first "Spot Price (Rs.)" , setting it as max spot price
max_spot_price_row = 2                        # setting row number of max spot price as 2
for row in range(3, total_rows + 1):
    spot_price = worksheet.Cells(row, 4).Value  # As spot price is in 4th column
    if spot_price > max_spot_price:
        max_spot_price = spot_price
        max_spot_price_row = row
max_spot_price_date = worksheet.Cells(max_spot_price_row, 6).Value  # As date is in Sixth column
print("b.)- Date with highest Spot Price: ",max_spot_price_date)

# Step 6c: Saving the extracted data to an Excel workbook with sheet name as "Raw Data"
workbook.SaveAs('D:\\Himanshu 2.0\\IMARC - Python Assignment\\Scraping and data analysis with python\\Raw_data.xls')
print("Data extracted successfully with name Raw_data.xls")

# Closing the workbook and quit Excel
workbook.Close()
excel.Quit()

