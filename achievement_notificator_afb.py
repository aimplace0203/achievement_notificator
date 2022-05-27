import os
import re
import csv
import sys
import json
import shutil
import datetime
import requests
from time import sleep
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from fake_useragent import UserAgent
from webdriver_manager.chrome import ChromeDriverManager

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
def importCsvFromAfb(downloadsDirPath):
    url = "https://www.afi-b.com/"
    login = os.environ['AFB_ID']
    password = os.environ['AFB_PASS']

    ua = UserAgent()
    logger.debug(f'importCsvFromAfb: UserAgent: {ua.chrome}')

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

        driver.find_element_by_xpath('//input[@name="login_name"]').send_keys(login)
        driver.find_element_by_xpath('//input[@name="password"]').send_keys(password)
        driver.find_element_by_xpath('//button[@type="submit"]').click()

        logger.debug('importCsvFromAfb: afb login')
        sleep(3)
        driver.implicitly_wait(60)
        
        driver.find_element_by_xpath('//a[@href="/pa/result/"]').click()
        sleep(3)
        driver.implicitly_wait(30)

        driver.find_element_by_xpath(f'//input[@value="td"]').click()
        logger.info('importCsvFromAfb: select date range')
        sleep(3)
        driver.implicitly_wait(30)

        driver.find_element_by_xpath('//input[@src="/assets/img/report/btn_original_csv.gif"]').click()
        sleep(5)

        driver.close()
        driver.quit()
    except Exception as err:
        logger.debug(f'Error: importCsvFromAfb: {err}')
        exit(1)

def getLatestDownloadedFileName(downloadsDirPath):
    if len(os.listdir(downloadsDirPath)) == 0:
        return None
    return max (
        [downloadsDirPath + '/' + f for f in os.listdir(downloadsDirPath)],
        key=os.path.getctime
    )

def sendChatworkNotification(message):
    try:
        url = f'https://api.chatwork.com/v2/rooms/{os.environ["CHATWORK_ROOM_ID_ACHIEVEMENT"]}/messages'
        headers = { 'X-ChatWorkToken': os.environ["CHATWORK_API_TOKEN"] }
        params = { 'body': message }
        requests.post(url, headers=headers, params=params)
    except Exception as err:
        logger.error(f'Error: sendChatworkNotification: {err}')
        exit(1)

### Achievement! ###
def readCsvData(csvPath, code):
    with open(csvPath, newline='', encoding=code) as csvfile:
        buf = csv.reader(csvfile, delimiter=',', lineterminator='\r\n', skipinitialspace=True)
        for row in buf:
            yield row

def getAchievementData(data, previous):
    header = data.pop(0)
    for i, d in enumerate(header):
        if re.search('発生日', d):
            date = int(i)
        elif re.search('報酬', d):
            reward = int(i)
        elif re.search('サイト名', d):
            site = int(i)
        elif re.search('PID', d):
            pid = int(i)
        elif re.search('プロモーション名', d):
            name = int(i)
    
    diff = len(data) - len(previous)

    while diff > 0:
        item = data.pop(0)
        yield [item[date], item[reward], item[site], item[pid], item[name]]
        diff -= 1

def getCsvPath(dirPath):
    os.makedirs(dirPath, exist_ok=True)
    importCsvFromAfb(dirPath)

    csvPath = getLatestDownloadedFileName(dirPath)
    logger.info(f"achievement_notificator: download completed: {csvPath}")

    return csvPath

def createCsvFile(data, outputFilePath):
    header = ["発生日","報酬","サイト名","PID","プロモーションID"]
    with open(outputFilePath, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f, delimiter=',', lineterminator='\r\n',  quoting=csv.QUOTE_ALL)
        writer.writerow(header)
        writer.writerows(data)

### main_script ###
if __name__ == '__main__':

    if len(sys.argv) > 1:
        print("Complete remove")
        shutil.rmtree('./csv/afb/')
        shutil.rmtree('./data/afb/')
        exit(0)

    try:
        os.makedirs('./csv/afb/', exist_ok=True)
        csvPath = getCsvPath('./csv/afb/')

        previous = []
        if os.path.exists('./data/afb/data.csv'):
            data = list(readCsvData('./data/afb/data.csv', 'utf-8'))
            if len(data) > 0:
                data.pop(0)
                previous = data
        all_list = previous

        data = list(readCsvData(csvPath, 'cp932'))
        new = list(getAchievementData(data, previous))
        new.reverse()

        if len(new) == 0:
            logger.info("No new achievements")
        else:
            all_list.extend(new)
            os.makedirs('./data/', exist_ok=True)
            createCsvFile(all_list, './data/afb/data.csv')

            total = 0
            for item in all_list:
                total += int(item[1])
            total = '{:,}'.format(total)

            message = "[info][title]【祝】新規成果発生のお知らせ！[/title]"
            message += f"新規で【{len(new)}件】成果が発生しました。\n"
            message += f"本日の累計成果報酬は【¥{total}】です。\n"
            for item in new:
                message += '\n＋＋＋\n\n'
                message += f'発生日：{item[0]}\n'
                reward = '{:,}'.format(int(item[1]))
                message += f'報酬：¥{reward}\n'
                message += f'サイト名：{item[2]}\n'
                message += f'プロモーションID：{item[3]}\n'
                message += f'プロモーション名：{item[4]}\n'
            message += '[/info]'

            sendChatworkNotification(message)

        logger.info("achievement_notificator: Finish")
        exit(0)
    except Exception as err:
        logger.debug(f'achievement_notificator: {err}')
        exit(1)
