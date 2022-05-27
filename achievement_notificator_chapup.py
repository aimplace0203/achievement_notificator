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
def importCsvFromChapup(downloadsDirPath):
    login = os.environ['CHAPUP_ID']
    password = os.environ['CHAPUP_PASS']
    url = f"https://{login}:{password}@chapup.jp/media_adv/"

    ua = UserAgent()
    logger.debug(f'importCsvFromChapup: UserAgent: {ua.chrome}')

    options = Options()
#    options.add_argument('--headless')
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
        select.select_by_value(str(today.year))

        dropdown = driver.find_element_by_name('search_startmonth')
        select = Select(dropdown)
        select.select_by_value(str(today.month))

        dropdown = driver.find_element_by_name('search_startday')
        select = Select(dropdown)
        select.select_by_value(str(today.day))

        dropdown = driver.find_element_by_name('search_endyear')
        select = Select(dropdown)
        select.select_by_value(str(today.year))

        dropdown = driver.find_element_by_name('search_endmonth')
        select = Select(dropdown)
        select.select_by_value(str(today.month))

        dropdown = driver.find_element_by_name('search_endday')
        select = Select(dropdown)
        select.select_by_value(str(today.day))

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

def sendChatworkNotification(message):
    try:
        url = f'https://api.chatwork.com/v2/rooms/{os.environ["CHATWORK_ROOM_ID_CHAPUP"]}/messages'
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

def getAchievementData(data, prev):
    header = data.pop(0)
    for i, d in enumerate(header):
        if re.search('ADコード／媒体名', d):
            ad = int(i)
        elif re.search('定期申込数\(集計期間内のみ\)', d):
            cnt = int(i)
    
    global code
    global output
    for item in data:
        id = item[ad]
        count = item[cnt]
        output[id] = count
    
    if len(prev) == 0:
        for k, v in output.items():
            if int(v) > 0:
                yield [k, v]
    else:
        for k, v in output.items():
            diff = int(v) - int(prev[k])
            if diff > 0:
                yield [k, diff]

def getCsvPath(dirPath):
    os.makedirs(dirPath, exist_ok=True)
    importCsvFromChapup(dirPath)

    csvPath = getLatestDownloadedFileName(dirPath)
    logger.info(f"achievement_notificator: download completed: {csvPath}")

    return csvPath

### main_script ###
if __name__ == '__main__':

    option = 0
    if len(sys.argv) > 1:
        option = int(sys.argv[1])

    code = dict()
    output = dict()
    try:
        os.makedirs('./csv/chapup/', exist_ok=True)
        csvPath = getCsvPath('./csv/chapup/')

        prev = {}
        if os.path.exists('./data/chapup/data.json'):
            with open('./data/chapup/data.json', 'r') as f:
                prev = json.load(f)

        data = list(readCsvData(csvPath, 'cp932'))
        new = list(getAchievementData(data, prev))

        total = 0
        for item in new:
            total += int(item[1])
        if total == 0:
            logger.info("No new achievements")
        else:
            os.makedirs('./data/chapup/', exist_ok=True)
            with open('./data/chapup/data.json', 'w') as f:
                json.dump(output, f, ensure_ascii=False, indent=4)

            message = "[info][title]【祝】新規成果発生のお知らせ！[/title]"
            message += f"新規で【{total}件】申込が発生しました。\n"
            message += f"本日の累計成果報酬は【¥{total * 22000}】です。\n"
            for item in new:
                message += '\n＋＋＋\n\n'
                message += f'新規発生件数：{item[1]}\n'
                reward = '{:,}'.format(int(item[1]) * 22000)
                message += f'報酬：¥{reward}\n'
                message += f'媒体名：{code[item[0]]}\n'
            message += '[/info]'

            print(message)
#            sendChatworkNotification(message)

        if not option == 0:
            print("remove")
            shutil.rmtree('./csv/chapup/')
            shutil.rmtree('./data/chapup/')

        logger.info("achievement_notificator: Finish")
        exit(0)
    except Exception as err:
        logger.debug(f'achievement_notificator: {err}')
        exit(1)
