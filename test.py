from zipfile import ZipFile
import pandas as pd
import io
import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By

DOWNLOAD_PATH = os.getcwd()
URL = 'https://cde.ucr.cjis.gov/LATEST/webapp/#/pages/home'
URL_SECTION = 'Documents & Downloads'
DOWNLOAD_SECTION = 'National Incident-Based Reporting System (NIBRS) Tables'
DOWNLOAD_SECTION_ID = 'dwnnibrs-download-select'
DOWNLOAD_TABLE_OPTION = 'nb-option-71'
DOWNLOAD_YEAR = 'dwnnibrscol-year-select'
DOWNLOAD_YEAR_OPTION = 'nb-option-85'
DOWNLOAD_LOCATION = 'dwnnibrsloc-select'
DOWNLOAD_LOCATION_OPTION = 'nb-option-243'
DOWNLOAD_BUTTON = 'nibrs-download-button'
EXCEL_CATEGORY = 'Crimes Against Property'

def read_zip(zip_path, zip_extract):
    zf = ZipFile(zip_path)
    return zf.read(zip_extract)

def read_xlxs(path, sheet_name=0):
    df = pd.read_excel(path, sheet_name=sheet_name)
    return df

def to_csv(df: pd.DataFrame, index=False, output_path='test.csv'):
    return df.to_csv(output_path, index=index, header= False)  

def sub_df(df, export_rows, exclude_columns):
    return df.iloc[export_rows].drop(axis=1,columns=df.columns[exclude_columns])

def find_first(df, str_to_find) -> int:
    for index, row in df.iterrows():
        if(row.isin([str_to_find]).any()):
            return index
    return 0

def find_footer(df, start_index=0):
    for index, row in df.iloc[start_index:].iterrows():
        if(row.isnull().any()):
            return index
    return -1        

def scrap_excel(zip_path, xlsx):
    table = read_xlxs(io.BytesIO(read_zip(zip_path, xlsx)))
    first_index = find_first(table, EXCEL_CATEGORY)+1
    footer_index = find_footer(table, start_index=first_index)
    
    export_rows = list(range(first_index, footer_index))
    exclude_columns = [find_first(table, 'Total Victims1')+1]
    to_csv(sub_df(table, export_rows=export_rows, exclude_columns=exclude_columns))
    return

def create_firefox_options():
    profile= webdriver.FirefoxProfile()
    profile.set_preference("browser.download.folderList",2)
    profile.set_preference("browser.download.manager.showWhenStarting",False)
    profile.set_preference("browser.download.dir", DOWNLOAD_PATH)
    profile.set_preference("browser.helperApps.neverAsk.saveToDisk","application/octet-stream")

    option_firefox = webdriver.FirefoxOptions()
    option_firefox.profile = profile
    return option_firefox

def scrap_web():
    browser = webdriver.Firefox(options=create_firefox_options())
    browser.get(URL)
    browser.find_element(By.LINK_TEXT, URL_SECTION).click() 
    
    browser.find_element(By.ID, DOWNLOAD_SECTION_ID).click()
    browser.find_element(By.ID, DOWNLOAD_TABLE_OPTION).click()
    
    browser.find_element(By.ID, DOWNLOAD_YEAR).click()
    browser.find_element(By.ID, DOWNLOAD_YEAR_OPTION).click()
    
    browser.find_element(By.ID, DOWNLOAD_LOCATION).click()
    browser.find_element(By.ID, DOWNLOAD_LOCATION_OPTION).click()

    browser.find_element(By.ID, DOWNLOAD_BUTTON).click()
    time.sleep(5)
    browser.quit()

if __name__ == "__main__":
    scrap_web()
    print('Generando el archivo csv..')
    scrap_excel('victims.zip', 'Victims_Age_by_Offense_Category_2022.xlsx')
    print('Resultado generado como test.csv')