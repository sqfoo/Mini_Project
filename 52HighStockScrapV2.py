from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import xlsxwriter

TargetVolume = 1000000

# The browser used Here is Safari
driver = webdriver.Safari()
MainURL = "https://www.investing.com"
FiftyTwoHighUrl = '/equities/52-week-high'
driver.get(MainURL+FiftyTwoHighUrl)

Link = driver.find_elements(By.CLASS_NAME, "alertBellGrayPlus")
FiftyWeekHigh = []
i = 0
for link in Link:
    FiftyWeekHigh.append(link.get_attribute("data-name"))

Stocks = []
Sectors = []
Volumes = []
# Exchanges = []
for stockName in FiftyWeekHigh:
    driver.get(MainURL + FiftyTwoHighUrl)
    TempUrl = driver.find_element(By.XPATH, "//a[@title='"+stockName+"']").get_attribute('href')
    driver.get(TempUrl)
    TempVolume = driver.find_element(By.XPATH, "//div[@data-test='volume-value']").text
    TempVolume = int(TempVolume.replace(',', ''))
    if TempVolume >= TargetVolume:
        i += 1
        driver.get(TempUrl+"-company-profile")
        Stocks.append(stockName)
        Volumes.append(TempVolume)
        Sectors.append(driver.find_element(By.XPATH, '//*[@id="leftColumn"]/div[8]/div[2]/a').text)


time.sleep(5)
driver.quit()

# Create and Open an Excel File
FileName = "StockList.xlsx"
workbook = xlsxwriter.Workbook(FileName)
worksheet = workbook.add_worksheet("Sheet 1")
worksheet.write(0, 0, "Stock")
worksheet.write(0, 1, "Vol.")
worksheet.write(0, 2, "Sector")

for index in range(i):
    # print(Stocks[index], Volumes[index],  Sectors[index])
    worksheet.write(index+1, 0, Stocks[index])
    worksheet.write(index + 1, 1, Volumes[index])
    worksheet.write(index + 1, 2, Sectors[index])
#
workbook.close()

