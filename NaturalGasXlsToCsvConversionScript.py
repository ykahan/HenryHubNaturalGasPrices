import shutil
import time
from os import path, remove, rename

import pandas as pd
from selenium import webdriver

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
monthlyDataName = "RNGWHHDm"
weeklyDataName = "RNGWHHDw"
dailyDataName = "RNGWHHDd"

names = [annualDataName, monthlyDataName, weeklyDataName, dailyDataName]

for name in names:
    excel = name + ".xls"
    newExcel = ""
    if name.endswith("a"):
        newExcel = "annual.xls"
        csv = "annual"
    elif name.endswith("w"):
        newExcel = "weekly.xls"
        csv = "weekly"
    elif name.endswith("d"):
        newExcel = "daily.xls"
        csv = "daily"
    else:
        newExcel = "monthly.xls"
        csv = "monthly"
    csv += ".csv"

    source = downloadsPath + excel
    destinationExcel = "C:\\Users\\USER\\Documents\\Yehoshua\\Programming\\Python\\HenryHubNaturalGasPricesOld" \
                       "\\XlsFiles\\"
    destinationCsv = "C:\\Users\\USER\\Documents\\Yehoshua\\Programming\\Python\\HenryHubNaturalGasPricesOld" \
                     "\\CsvFiles\\"

    excelPath = destinationExcel + "\\" + excel
    newExcelPath = destinationExcel + "\\" + newExcel
    csvPath = destinationCsv + csv

    excelPathExists = path.exists(excelPath)
    newExcelPathExists = path.exists(newExcelPath)
    csvPathExists = path.exists(csvPath)

    if excelPathExists:
        remove(excelPath)

    shutil.move(source, destinationExcel)

    if newExcelPathExists:
        remove(newExcelPath)

    rename(excelPath, newExcelPath)
    read_file = pd.read_excel(newExcelPath, "Data 1")

    if csvPathExists:
        remove(csvPath)

    read_file.to_csv(csvPath, index=None, header=True)
