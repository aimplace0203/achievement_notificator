import os
import re
import csv
import sys
import json
import shutil
import datetime
import requests
import gspread
from time import sleep
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from fake_useragent import UserAgent
from webdriver_manager.chrome import ChromeDriverManager
from oauth2client.service_account import ServiceAccountCredentials

# Logger setting
from logging import getLogger, FileHandler, DEBUG
logger = getLogger(__name__)
today = datetime.datetime.now()
os.makedirs('./log', exist_ok=True)
handler = FileHandler(f'log/{today.strftime("%Y-%m-%d")}_result.log', mode='a')
handler.setLevel(DEBUG)
logger.setLevel(DEBUG)
logger.addHandler(handler)
logger.propagate = False

### functions ###
def importCsvFromChapup(downloadsDirPath, day):
    login = os.environ['BRESMILE_ID']
    password = os.environ['BRESMILE_PASS']
    url = f"https://{login}:{password}@bresmile.jp/media_adv/"

    ua = UserAgent()
    logger.debug(f'importCsvFromBresmile: UserAgent: {ua.chrome}')

    options = Options()
    options.add_argument('--headless')
    options.add_argument(f'user-agent={ua.chrome}')

    prefs = {
        "profile.default_content_settings.popups": 1,
        "download.default_directory": 
                os.path.abspath(downloadsDirPath),
        "directory_upgrade": True
    }
    options.add_experimental_option("prefs", prefs)
    
    try:
        driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
        
        driver.get(url)
        driver.maximize_window()
        sleep(3)
        driver.implicitly_wait(30)

        dropdown = driver.find_element(By.NAME, 'search_startyear')
        select = Select(dropdown)
        select.select_by_value(str(day.year))

        dropdown = driver.find_element(By.NAME, 'search_startmonth')
        select = Select(dropdown)
        select.select_by_value(str(day.month))

        dropdown = driver.find_element(By.NAME, 'search_startday')
        select = Select(dropdown)
        select.select_by_value(str(day.day))

        dropdown = driver.find_element(By.NAME, 'search_endyear')
        select = Select(dropdown)
        select.select_by_value(str(day.year))

        dropdown = driver.find_element(By.NAME, 'search_endmonth')
        select = Select(dropdown)
        select.select_by_value(str(day.month))

        dropdown = driver.find_element(By.NAME, 'search_endday')
        select = Select(dropdown)
        select.select_by_value(str(day.day))

        driver.find_element(By.NAME, 'kikan').click()
        sleep(3)
        driver.find_element(By.NAME, 'csv').click()
        logger.info('importCsvFromBresmile: Complete download')
        sleep(3)

        soup = BeautifulSoup(driver.page_source, "html.parser")
        els = iter(soup.find_all("td", align="center"))
        global code
        while True:
            try:
                id = next(els).text
                name = re.sub('^LI_', '', next(els).text)
                code[id] = name
            except StopIteration:
                break

        driver.close()
        driver.quit()
    except Exception as err:
        logger.debug(f'Error: importCsvFromBresmile: {err}')
        exit(1)

def getLatestDownloadedFileName(downloadsDirPath):
    if len(os.listdir(downloadsDirPath)) == 0:
        return None
    return max (
        [downloadsDirPath + '/' + f for f in os.listdir(downloadsDirPath)],
        key=os.path.getctime
    )

### Achievement! ###
def readCsvData(csvPath, code):
    with open(csvPath, newline='', encoding=code) as csvfile:
        buf = csv.reader(csvfile, delimiter=',', lineterminator='\r\n', skipinitialspace=True)
        for row in buf:
            yield row

def getAchievementData(data, day):
    header = data.pop(0)
    for i, d in enumerate(header):
        if re.search('ADコード／媒体名', d):
            ad = int(i)
        elif re.search('注文回数\(合計\)', d):
            cnt = int(i)
    
    global code
    res = {
        'yss': 0,
        'gsn': 0,
        'yda': 0,
        'gdn': 0,
        'line': 0,
        'tiktok': 0,
        'banner': 0
    }
    for item in data:
        id = item[ad]
        if re.search('yss', code[id], re.IGNORECASE):
            res['yss'] += int(item[cnt])
        elif re.search('gsn', code[id], re.IGNORECASE):
            res['gsn'] += int(item[cnt])
        elif re.search('yda', code[id], re.IGNORECASE):
            res['yda'] += int(item[cnt])
        elif re.search('gdn', code[id], re.IGNORECASE):
            res['gdn'] += int(item[cnt])
        elif re.search('line', code[id], re.IGNORECASE):
            res['line'] += int(item[cnt])
        elif re.search('tiktok', code[id], re.IGNORECASE):
            res['tiktok'] += int(item[cnt])
        elif re.search('記事離脱用', code[id], re.IGNORECASE):
            res['banner'] += int(item[cnt])
    logger.info(f'date: {day}, data: {res}')
    
    writeOrderData(res, day)

def getCsvPath(dirPath, day):
    os.makedirs(dirPath, exist_ok=True)
    importCsvFromChapup(dirPath, day)

    csvPath = getLatestDownloadedFileName(dirPath)
    logger.info(f"achievement_notificator: download completed: {csvPath}")

    return csvPath

def writeOrderData(data, day):
    SPREADSHEET_ID = os.environ['BRESMILE_SSID']
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('spreadsheet.json', scope)
    gc = gspread.authorize(credentials)
    sheet = gc.open_by_key(SPREADSHEET_ID).worksheet(day.strftime("%Y%m"))

    sheet.update_cell(4 + int(day.strftime("%d")), 6, data['banner'])
    sheet.update_cell(4 + int(day.strftime("%d")), 14, data['gsn'])
    sheet.update_cell(4 + int(day.strftime("%d")), 17, data['gdn'])
    sheet.update_cell(4 + int(day.strftime("%d")), 20, data['yss'])
    sheet.update_cell(4 + int(day.strftime("%d")), 23, data['yda'])
    sheet.update_cell(4 + int(day.strftime("%d")), 26, data['line'])
    sheet.update_cell(4 + int(day.strftime("%d")), 29, data['tiktok'])


### main_script ###
if __name__ == '__main__':

    r = 1
    if len(sys.argv) > 1:
        r = 15

    code = dict()
    try:
        os.makedirs('./csv/bresmile/', exist_ok=True)

        for i in range(0, r):
            day = today - datetime.timedelta(days=i)
            csvPath = getCsvPath('./csv/bresmile/', day)
            data = list(readCsvData(csvPath, 'cp932'))
            getAchievementData(data, day)

        if r > 2:
            shutil.rmtree('./csv/bresmile/')
        logger.info("achievement_notificator: Finish")
        exit(0)
    except Exception as err:
        logger.debug(f'achievement_notificator: {err}')
        exit(1)
