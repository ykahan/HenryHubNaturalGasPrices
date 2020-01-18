import time
from selenium import webdriver
import shutil
import pandas as pd

source = "C:\\Users\\USER\\Java_Workspace_Yehoshua\\Selenium Dependencies"
source += "\\chromedriver.exe"
driver = webdriver.Chrome(source)  # Optional argument, if not specified will search path.
target = "http://www.eia.gov/dnav/ng/hist/rngwhhdm.htm"
driver.implicitly_wait(10)
driver.get(target)

dailyPricesXpath = "//a[contains(@href, 'whhdD') and contains(@class, 'Nav')]"
weeklyPricesXpath = "//a[contains(@href, 'whhdW') and contains(@class, 'Nav')]"
monthlyPricesXpath = "//a[contains(@href, 'whhdM') and contains(@class, 'Nav')]"
yearlyPricesXpath = "//a[contains(@href, 'whhdA') and contains(@class, 'Nav')]"

dailyPrices = driver.find_element_by_xpath(dailyPricesXpath)
weeklyPrices = driver.find_element_by_xpath(weeklyPricesXpath)
monthlyPrices = driver.find_element_by_xpath(monthlyPricesXpath)
yearlyPrices = driver.find_element_by_xpath(yearlyPricesXpath)

datasets = [dailyPricesXpath, weeklyPricesXpath, monthlyPricesXpath, yearlyPricesXpath]
for dataset in datasets:
    target = driver.find_element_by_xpath(dataset)
    target.click()

    downloadXpath = "//a[contains(text(), 'Download')]"
    downloadLink = driver.find_element_by_xpath(downloadXpath)
    downloadLink.click()

    time.sleep(5)

url = driver.current_url
driver.quit()

downloadsPath = "C:\\Users\\USER\\Downloads\\"
annualDataName = "RNGWHHDa"
monthlyDataName = "RNGWHHDm (1)"
weeklyDataName = "RNGWHHDw"
dailyDataName = "RNGWHHDd"

names = [annualDataName, monthlyDataName, weeklyDataName, dailyDataName]

for name in names:
    excel = name + ".xls"
    if name.endswith("a"):
        csv = "annual"
    elif name.endswith("w"):
        csv = "weekly"
    elif name.endswith("d"):
        csv = "daily"
    else:
        csv = "monthly"
    csv += ".csv"

    source = downloadsPath + excel
    destExcel = "C:\\Users\\USER\\Documents\\Yehoshua\\Programming\\Python\\HenryHubNaturalGasPricesOld\\XlsFiles\\"
    shutil.move(source, destExcel)

    destCsv = "C:\\Users\\USER\\Documents\\Yehoshua\\Programming\\Python\\HenryHubNaturalGasPricesOld\\CsvFiles\\"

    excelPath = destExcel + "\\" + excel
    csvPath = destCsv + csv

    read_file = pd.read_excel(excelPath, "Data 1")
    read_file.to_csv(csvPath, index=None, header=True)
