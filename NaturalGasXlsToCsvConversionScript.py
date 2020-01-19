import shutil
import time
from os import path, remove, rename, listdir

import pandas as pd
from selenium import webdriver


def remove_files(provided_path, provided_name):
    for file in listdir(provided_path):
        if file.startswith(provided_name):
            remove(downloadsPath + file)


def prepare_driver(path_to_chromedriver, seconds):
    drive = webdriver.Chrome(path_to_chromedriver)
    drive.implicitly_wait(seconds)
    return drive


source = "C:\\Users\\USER\\Java_Workspace_Yehoshua\\Selenium Dependencies"
source += "\\chromedriver.exe"
wait = 10

driver = prepare_driver(source, wait)

target = "http://www.eia.gov/dnav/ng/hist/rngwhhdm.htm"
driver.get(target)

dailyPricesXpath = "//a[contains(@href, 'whhdD') and contains(@class, 'Nav')]"
weeklyPricesXpath = "//a[contains(@href, 'whhdW') and contains(@class, 'Nav')]"
monthlyPricesXpath = "//a[contains(@href, 'whhdM') and contains(@class, 'Nav')]"
yearlyPricesXpath = "//a[contains(@href, 'whhdA') and contains(@class, 'Nav')]"

downloadsPath = "C:\\Users\\USER\\Downloads\\"
commonName = "RNGWHH"

remove_files(downloadsPath, commonName)

datasets = [dailyPricesXpath, weeklyPricesXpath, monthlyPricesXpath, yearlyPricesXpath]


def download_file(xpath):
    download_link = driver.find_element_by_xpath(xpath)
    download_link.click()


for dataset in datasets:
    target = driver.find_element_by_xpath(dataset)
    target.click()

    downloadXpath = "//a[contains(text(), 'Download')]"
    download_file(downloadXpath)

    time.sleep(5)

driver.quit()

annualDataName = "RNGWHHDa"
monthlyDataName = "RNGWHHDm"
weeklyDataName = "RNGWHHDw"
dailyDataName = "RNGWHHDd"

old_names = [annualDataName, monthlyDataName, weeklyDataName, dailyDataName]


def get_file_names(old_name):
    # global excel, newExcel, csv
    exc = old_name + ".xls"
    if old_name.endswith("a"):
        new_excel = "annual.xls"
        comma_sep_values = "annual"
    elif old_name.endswith("w"):
        new_excel = "weekly.xls"
        comma_sep_values = "weekly"
    elif old_name.endswith("d"):
        new_excel = "daily.xls"
        comma_sep_values = "daily"
    else:
        new_excel = "monthly.xls"
        comma_sep_values = "monthly"
    comma_sep_values += ".csv"
    return [exc, new_excel, comma_sep_values]


def move_file(path_exists, src, destination):
    if path_exists:
        remove(destination)
    shutil.move(src, destination)


def rename_file(path_exists, old_name, new_name):
    if path_exists:
        remove(new_name)
    rename(old_name, new_name)


def convert_from_xls_to_csv(local_excel, local_sheet, csv_path):
    read_file = pd.read_excel(local_excel, local_sheet)
    read_file.to_csv(csv_path, index=None, header=True)


for name in old_names:
    new_names = get_file_names(name)
    excel = new_names[0]
    newExcel = new_names[1]
    csv = new_names[2]

    source = downloadsPath + excel
    destinationExcel = "C:\\Users\\USER\\Documents\\Yehoshua\\Programming\\Python\\HenryHubNaturalGasPricesOld" \
                       "\\XlsFiles\\"

    excelPath = destinationExcel + "\\" + excel
    newExcelPath = destinationExcel + "\\" + newExcel

    destinationCsv = "C:\\Users\\USER\\Documents\\Yehoshua\\Programming\\Python\\HenryHubNaturalGasPricesOld" \
                     "\\CsvFiles\\"
    csvPath = destinationCsv + csv

    excelPathExists = path.exists(excelPath)
    newExcelPathExists = path.exists(newExcelPath)
    csvPathExists = path.exists(csvPath)

    move_file(excelPathExists, source, excelPath)

    rename_file(newExcelPathExists, excelPath, newExcelPath)

    if csvPathExists:
        remove(csvPath)

    sheet = "Data 1"
    convert_from_xls_to_csv(newExcelPath, sheet, csvPath)
