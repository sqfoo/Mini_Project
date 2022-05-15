from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import time
import xlsxwriter

# Create and Open an Excel File
FileName = "StockList.xlsx"
workbook = xlsxwriter.Workbook(FileName)
worksheet = workbook.add_worksheet("Sheet 1")
worksheet.write(0, 0, "Stock")
worksheet.write(0, 1, "Vol.")
worksheet.write(0,2, "Sector")

# The browser used Here is Safari
driver = webdriver.Safari()
driver.get("https://www.investing.com")

actions = ActionChains(driver)
MarketButton = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.CLASS_NAME, 'nav'))  # This is a dummy element
)
actions.move_to_element(MarketButton)
StockButton = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="navMenu"]/ul/li[1]/ul/li[2]'))  # This is a dummy element
)
actions.move_to_element(StockButton)
actions.perform()


FiftyTwoHighButton = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="navMenu"]/ul/li[1]/ul/li[2]/div/ul[1]/li[8]/a'))  # This is a dummy element
)
driver.get(FiftyTwoHighButton.get_attribute("href"))
ChangeVolumeButton = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="stockPageInnerContent"]/table/thead/tr/th[8]/span'))  # This is a dummy element
)
ChangeVolumeButton.click()

Link = driver.find_elements(By.CLASS_NAME, "alertBellGrayPlus")

i = 0
Stocks = []
Sectors = []
Volumes = []
Exchanges = []
for link in Link:
    Stocks.append(link.get_attribute("data-name"))
    i += 1

for stock in Stocks:
    driver.get("https://www.investing.com")
    SearchButton = driver.find_element(By.CLASS_NAME, "searchText")
    SearchButton.send_keys(stock)
    ResultButton = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/div[5]/header/div[2]/div/div[3]/div[2]/div[1]/div[1]/div[2]/div/a/span[3]'))  # This is a dummy element
    )
    ResultButton.click()
    Volume = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CLASS_NAME, 'trading-hours_value__2MrOn'))
    )
    Volumes.append(Volume.text)
    Profile = driver.find_element(By.LINK_TEXT, 'Profile')
    url = Profile.get_attribute('href')
    driver.get(url)
    Sector = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="leftColumn"]/div[8]/div[2]/a'))
    )
    Sectors.append(Sector.text)

time.sleep(5)
driver.quit()


for index in range(i):
    # print(Stocks[index], Volumes[index],  Sectors[index])
    worksheet.write(index+1, 0, Stocks[index])
    worksheet.write(index + 1, 1, Volumes[index])
    worksheet.write(index + 1, 2, Sectors[index])

workbook.close()

