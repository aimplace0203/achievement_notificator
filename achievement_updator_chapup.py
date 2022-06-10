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
    login = os.environ['CHAPUP_ID']
    password = os.environ['CHAPUP_PASS']
    url = f"https://{login}:{password}@chapup.jp/media_adv/"

    ua = UserAgent()
    logger.debug(f'importCsvFromChapup: UserAgent: {ua.chrome}')

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

        dropdown = driver.find_element_by_name('search_startyear')
        select = Select(dropdown)
        select.select_by_value(str(day.year))

        dropdown = driver.find_element_by_name('search_startmonth')
        select = Select(dropdown)
        select.select_by_value(str(day.month))

        dropdown = driver.find_element_by_name('search_startday')
        select = Select(dropdown)
        select.select_by_value(str(day.day))

        dropdown = driver.find_element_by_name('search_endyear')
        select = Select(dropdown)
        select.select_by_value(str(day.year))

        dropdown = driver.find_element_by_name('search_endmonth')
        select = Select(dropdown)
        select.select_by_value(str(day.month))

        dropdown = driver.find_element_by_name('search_endday')
        select = Select(dropdown)
        select.select_by_value(str(day.day))

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

        driver.find_element_by_name('kikan').click()
        sleep(3)
        driver.find_element_by_name('csv').click()
        logger.info('importCsvFromChapup: Complete download')
        sleep(3)

        driver.close()
        driver.quit()
    except Exception as err:
        logger.debug(f'Error: importCsvFromChapup: {err}')
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
        'tiktok': 0
    }
    for item in data:
        id = item[ad]
        if re.search('YSS', code[id]):
            res['yss'] += int(item[cnt])
        elif re.search('GSN', code[id]):
            res['gsn'] += int(item[cnt])
        elif re.search('YDA', code[id]):
            res['yda'] += int(item[cnt])
        elif re.search('GDN', code[id]):
            res['gdn'] += int(item[cnt])
        elif re.search('LINE', code[id]):
            res['line'] += int(item[cnt])
        elif re.search('tiktok', code[id]):
            res['tiktok'] += int(item[cnt])
    logger.info(f'date: {day}, data: {res}')
    
    writeOrderData(res, day)

def getCsvPath(dirPath, day):
    os.makedirs(dirPath, exist_ok=True)
    importCsvFromChapup(dirPath, day)

    csvPath = getLatestDownloadedFileName(dirPath)
    logger.info(f"achievement_notificator: download completed: {csvPath}")

    return csvPath

def writeOrderData(data, day):
    SPREADSHEET_ID = os.environ['HAIR_PROMOTION_SSID']
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('spreadsheet.json', scope)
    gc = gspread.authorize(credentials)
    sheet = gc.open_by_key(SPREADSHEET_ID).worksheet(day.strftime("%Y%m"))

    sheet.update_cell(4 + int(day.strftime("%d")), 15, data['gsn'])
    sheet.update_cell(4 + int(day.strftime("%d")), 18, data['gdn'])
    sheet.update_cell(4 + int(day.strftime("%d")), 21, data['yss'])
    sheet.update_cell(4 + int(day.strftime("%d")), 24, data['yda'])
    sheet.update_cell(4 + int(day.strftime("%d")), 27, data['line'])
    sheet.update_cell(4 + int(day.strftime("%d")), 30, data['tiktok'])


### main_script ###
if __name__ == '__main__':

    r = 1
    if len(sys.argv) > 1:
        r = 6

    code = dict()
    try:
        os.makedirs('./csv/chapup2/', exist_ok=True)

        for i in range(0, r):
            day = today - datetime.timedelta(days=i)
            csvPath = getCsvPath('./csv/chapup2/', day)
            data = list(readCsvData(csvPath, 'cp932'))
            getAchievementData(data, day)

        if r > 2:
            shutil.rmtree('./csv/chapup2/')
        logger.info("achievement_notificator: Finish")
        exit(0)
    except Exception as err:
        logger.debug(f'achievement_notificator: {err}')
        exit(1)
