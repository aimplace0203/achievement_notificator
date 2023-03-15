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
def importCsvFromAfb(downloadsDirPath, d):
    url = "https://www.afi-b.com/"
    login = os.environ['AFB_ID']
    password = os.environ['AFB_PASS']

    logger.debug(d)
    if d == 0:
        da = "td"
    elif d == 1:
        da = "ytd"
    elif d == 2:
        da = "bytd"
    
    ua = UserAgent()
    logger.debug(f'importCsvFromAfb: UserAgent: {ua.chrome}')

    options = Options()
    #options.add_argument('--headless')
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
        driver.implicitly_wait(30)

        driver.find_element(By.XPATH, '//input[@name="login_name"]').send_keys(login)
        driver.find_element(By.XPATH, '//input[@name="password"]').send_keys(password)
        driver.find_element(By.XPATH, '//button[@type="submit"]').click()

        logger.debug('importCsvFromAfb: afb login')
        driver.implicitly_wait(60)
        
        driver.find_element(By.XPATH, '//a[@href="/pa/result/"]').click()
        driver.implicitly_wait(30)
        driver.find_element(By.CLASS_NAME, 'chzn-single').click()
        driver.implicitly_wait(30)
        sites = driver.find_element(By.CLASS_NAME, 'chzn-results')
        options = sites.find_elements(By.CLASS_NAME, 'active-result')
        for option in options:
            if re.search('806580', option.text):
                option.click()

        logger.info('importCsvFromAfb: select site')
        driver.implicitly_wait(30)

        driver.find_element(By.XPATH, f'//input[@value="{da}"]').click()
        logger.info('importCsvFromAfb: select date range')
        driver.implicitly_wait(30)

        driver.find_element(By.XPATH, '//input[@src="/assets/img/report/btn_original_csv.gif"]').click()
        sleep(10)

        driver.close()
        driver.quit()
    except Exception as err:
        logger.debug(f'Error: importCsvFromAfb: {err}')
        exit(1)

def getDataFromLinkA(d):
    url = "https://link-ag.net/partner/sign_in"
    login = os.environ['LINKA_ID']
    password = os.environ['LINKA_PASS']

    ua = UserAgent()
    logger.debug(f'importCsvFromLinkA: UserAgent: {ua.chrome}')

    options = Options()
    #options.add_argument('--headless')
    options.add_argument(f'user-agent={ua.chrome}')

    try:
        driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)

        driver.get(url)
        driver.maximize_window()
        driver.implicitly_wait(30)

        driver.find_element(By.ID, 'login_id').send_keys(login)
        driver.find_element(By.ID, 'password').send_keys(password)
        driver.find_element(By.XPATH, '//input[@type="submit"]').click()

        logger.debug('importCsvFromLinkA: linka login')
        driver.get('https://link-ag.net/partner/achievements')
        driver.implicitly_wait(30)

        driver.find_elements(By.ID, 'occurrence_time_occurrence_time')[d].click()
        driver.implicitly_wait(30)

        logger.info('importCsvFromLinkA: select date range')
        driver.implicitly_wait(30)

        driver.find_element(By.XPATH, '//input[@value="検索"]').click()
        driver.implicitly_wait(30)

        dropdown = driver.find_element(By.ID, "separator")
        select = Select(dropdown)
        select.select_by_value('comma')
        driver.implicitly_wait(30)

        text = driver.find_element(By.CLASS_NAME, 'partnerMain-scroll').text
        driver.close()
        driver.quit()

        r = re.search('合計[\d]*件', text)
        sum = 0
        if r:
            r2 = re.search('[\d]+', r.group())
            if r2:
                sum = r2.group()
        return sum
    except Exception as err:
        logger.debug(f'Error: importCsvFromLinkA: {err}')
        exit(1)

def getLatestDownloadedFileName(downloadsDirPath):
    if len(os.listdir(downloadsDirPath)) == 0:
        return None
    return max (
        [downloadsDirPath + '/' + f for f in os.listdir(downloadsDirPath)],
        key=os.path.getctime
    )

def getCsvData(csvPath):
    with open(csvPath, newline='', encoding='cp932') as csvfile:
        buf = csv.reader(csvfile, delimiter=',', lineterminator='\r\n', skipinitialspace=True)
        head = next(buf)
        for i, d in enumerate(head):
            if re.search('プロモーション名', d):
                name = int(i)
            elif re.search('報酬', d):
                price = int(i)
        sum = 0
        for row in buf:
            if re.search('ポリピュア', row[name]):
                sum += 23000
            else:
                sum += int(row[price])
        return sum

def writeUploadData(linka, afb, day):
    SPREADSHEET_ID = os.environ['HAIR_PROMOTION_SSID']
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('spreadsheet.json', scope)
    gc = gspread.authorize(credentials)
    sheet = gc.open_by_key(SPREADSHEET_ID).worksheet(day.strftime("%Y%m"))

    sheet.update_cell(4 + int(day.strftime("%d")), 8, linka)
    sheet.update_cell(4 + int(day.strftime("%d")), 10, afb)

### main_script ###
if __name__ == '__main__':

    d = 0
    if len(sys.argv) > 1:
        d = int(sys.argv[1])

    try:
        linka = getDataFromLinkA(d)
        logger.info(f'LinkA: {linka}')

        dirPath = './csv/afb2'
        os.makedirs(dirPath, exist_ok=True)
        importCsvFromAfb(dirPath, d)
        afbCsvPath = getLatestDownloadedFileName(dirPath)
        afb = getCsvData(afbCsvPath)
        logger.info(f'Afb: {afb}')

        day = today - datetime.timedelta(days=d)
        writeUploadData(linka, afb, day)

        if d > 0:
            shutil.rmtree('./csv/afb2/')
        logger.info("achievement_updator_chapup2: Finish")
        exit(0)
    except Exception as err:
        logger.info(f"achievement_updator_chapup2: {err}")
        exit(1)
