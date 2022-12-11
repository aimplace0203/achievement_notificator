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
# Common #
def getLatestDownloadedFileName(downloadsDirPath):
    if len(os.listdir(downloadsDirPath)) == 0:
        return None
    return max (
        [downloadsDirPath + f for f in os.listdir(downloadsDirPath)],
        key=os.path.getctime
    )

def sendChatworkNotification(message, uname):
    try:
        url = f'https://api.chatwork.com/v2/rooms/{os.environ[f"CHATWORK_ROOM_ID_{uname}"]}/messages'
        headers = { 'X-ChatWorkToken': os.environ["CHATWORK_API_TOKEN"] }
        params = { 'body': message }
        requests.post(url, headers=headers, params=params)
    except Exception as err:
        logger.error(f'Error: sendChatworkNotification: {err}')
        exit(1)

def readCsvData(csvPath, code):
    with open(csvPath, newline='', encoding=code) as csvfile:
        buf = csv.reader(csvfile, delimiter=',', lineterminator='\r\n', skipinitialspace=True)
        for row in buf:
            yield row


# LinkA #
def getDataFromLinkA(downloadsDirPath, d):
    url = "https://link-ag.net/partner/sign_in"
    login = os.environ['LINKA_ID']
    password = os.environ['LINKA_PASS']

    ua = UserAgent()
    logger.debug(f'importCsvFromLinkA: UserAgent: {ua.chrome}')

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
        driver.implicitly_wait(30)

        driver.find_element_by_id('login_id').send_keys(login)
        driver.find_element_by_id('password').send_keys(password)
        driver.find_element_by_xpath('//input[@type="submit"]').click()

        logger.debug('importCsvFromLinkA: linka login')
        driver.get('https://link-ag.net/partner/achievements')
        driver.implicitly_wait(30)

        driver.find_elements_by_id('occurrence_time_occurrence_time')[d].click()
        driver.implicitly_wait(30)

        logger.info('importCsvFromLinkA: select date range')
        driver.implicitly_wait(30)

        driver.find_element_by_xpath('//input[@value="検索"]').click()
        driver.implicitly_wait(30)

        dropdown = driver.find_element_by_id("separator")
        select = Select(dropdown)
        select.select_by_value('comma')
        driver.implicitly_wait(30)

        driver.find_element_by_class_name('partnerMain-btn.partnerMain-btn-md.partnerMain-btn').click()
        logger.info('getDataFromLinkA: Complete download')
        sleep(3)

        driver.close()
        driver.quit()
    except Exception as err:
        logger.debug(f'Error: importCsvFromLinkA: {err}')
        exit(1)

def getCsvPathLinkA(dirPath, d):
    os.makedirs(dirPath, exist_ok=True)
    getDataFromLinkA(dirPath, d)

    csvPath = getLatestDownloadedFileName(dirPath)
    logger.info(f"achievement_notificator: download completed: {csvPath}")
    return csvPath

def getAchievementDataLinkA(data, prev):
    header = data.pop(0)
    for i, d in enumerate(header):
        if re.search('発生日時', d):
            date = int(i)
        elif re.search('報酬金額（税込）', d):
            reward = int(i)
        elif re.search('広告名', d):
            ad = int(i)
    
    diff = len(data) - len(prev)

    while diff > 0:
        item = data.pop(0)
        yield [item[date], item[reward], item[ad]]
        diff -= 1

def createCsvFileLinkA(data, outputFilePath):
    header = ["発生日時","報酬金額（税込）","広告名"]
    with open(outputFilePath, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f, delimiter=',', lineterminator='\r\n',  quoting=csv.QUOTE_ALL)
        writer.writerow(header)
        writer.writerows(data)

def getInfoLinkA():
    previous = []
    if os.path.exists(f'./data/chapup/linka.csv'):
        data = list(readCsvData(f'./data/chapup/linka.csv', 'utf-8'))
        if len(data) > 0:
            data.pop(0)
            previous = data
    all_list = previous

    csvPath = getCsvPathLinkA('./csv/linka/', 0)

    data = list(readCsvData(csvPath, 'utf-16'))
    new = list(getAchievementDataLinkA(data, previous))
    new.reverse()

    total_price = 0
    for item in all_list:
        total_price += int(int(item[1]) / 11 * 10)

    message = ''
    if len(new) == 0:
        logger.info("LinkA: No new achievements")
    else:
        all_list.extend(new)
        os.makedirs(f'./data/chapup/', exist_ok=True)
        createCsvFileLinkA(all_list, f'./data/chapup/linka.csv')

        for item in new:
            total_price += int(int(item[1]) / 11 * 10)
            message += '\n＋＋＋\n\n'
            message += f'発生日：{item[0]}\n'
            reward = '{:,}'.format(int(int(item[1]) / 11 * 10))
            message += f'報酬：¥{reward}\n'
            message += f'広告名：{item[2]}\n'
    return {
        'new': len(new),
        'total': len(all_list),
        'total_price': total_price,
        'message': message
    }


# Afb #
def getAchievementDataAfb(data, prev):
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
    
    chapup_data = []
    for item in data:
        if re.search('育毛剤',item[site]) or re.search('チャップアップ',item[site]):
            chapup_data.append(item)
    
    diff = len(chapup_data) - len(prev)

    while diff > 0:
        item = chapup_data.pop(0)
        yield [item[date], item[reward], item[site], item[pid], item[name]]
        diff -= 1

def createCsvFileAfb(data, outputFilePath):
    header = ["発生日","報酬","サイト名","PID","プロモーションID"]
    with open(outputFilePath, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f, delimiter=',', lineterminator='\r\n',  quoting=csv.QUOTE_ALL)
        writer.writerow(header)
        writer.writerows(data)

def getInfoAfb():
    previous = []
    if os.path.exists(f'./data/chapup/afb.csv'):
        data = list(readCsvData(f'./data/chapup/afb.csv', 'utf-8'))
        if len(data) > 0:
            data.pop(0)
            previous = data
    all_list = previous

    afbCsvPath = getLatestDownloadedFileName('./csv/afb/')

    data = list(readCsvData(afbCsvPath, 'cp932'))
    new = list(getAchievementDataAfb(data, previous))
    new.reverse()

    total_price = 0
    for item in all_list:
        total_price += int(item[1])

    message = ""
    if len(new) == 0:
        logger.info("Afb: No new achievements")
    else:
        all_list.extend(new)
        os.makedirs(f'./data/chapup/', exist_ok=True)
        createCsvFileAfb(all_list, f'./data/chapup/afb.csv')

        for item in new:
            total_price += int(item[1])
            message += '\n＋＋＋\n\n'
            message += f'発生日：{item[0]}\n'
            reward = '{:,}'.format(int(item[1]))
            message += f'報酬：¥{reward}\n'
            message += f'サイト名：{item[2]}\n'
            message += f'プロモーションID：{item[3]}\n'
            message += f'プロモーション名：{item[4]}\n'
    return {
        'new': len(new),
        'total': len(all_list),
        'total_price': total_price,
        'message': message
    }


# Media #
def importCsvFromMedia(downloadsDirPath, uname, domain):
    login = os.environ[f'{uname}_ID']
    password = os.environ[f'{uname}_PASS']
    url = f"https://{login}:{password}@{domain}/media_adv/"

    ua = UserAgent()
    logger.debug(f'importCsvFromMedia: UserAgent: {ua.chrome}')

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

        driver.find_element_by_name('kikan').click()
        sleep(5)

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

        driver.find_element_by_name('csv').click()
        logger.info('importCsvFromMedia: Complete download')
        sleep(3)

        driver.close()
        driver.quit()
    except Exception as err:
        logger.debug(f'Error: importCsvFromMedia: {err}')
        exit(1)

def getCsvPath(dirPath, uname, domain):
    os.makedirs(dirPath, exist_ok=True)
    importCsvFromMedia(dirPath, uname, domain)

    csvPath = getLatestDownloadedFileName(dirPath)
    logger.info(f"achievement_notificator: download completed: {csvPath}")

    return csvPath

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
            try:
                diff = int(v) - int(prev[k])
            except Exception as err:
                logger.debug(f'achievement_notificator: {err}')
                diff = int(v)
            if diff > 0:
                yield [k, diff]

def getInfoMedia(name, uname, domain, price):
    os.makedirs(f'./csv/{name}/', exist_ok=True)
    csvPath = getCsvPath(f'./csv/{name}/', uname, domain)

    prev = {}
    if os.path.exists(f'./data/{name}/data.json'):
        with open(f'./data/{name}/data.json', 'r') as f:
            prev = json.load(f)

    data = list(readCsvData(csvPath, 'cp932'))
    new = list(getAchievementData(data, prev))

    global output
    n = 0
    total = 0
    total_price = 0
    message = ''
    for v in output.values():
        total += int(v)
        total_price += int(v) * price
    for item in new:
        n += int(item[1])
    if n == 0:
        logger.info(f"{name}: No new achievements")
    else:
        os.makedirs(f'./data/{name}/', exist_ok=True)
        with open(f'./data/{name}/data.json', 'w') as f:
            json.dump(output, f, ensure_ascii=False, indent=4)

        for item in new:
            message += '\n＋＋＋\n\n'
            message += f'新規発生件数：{item[1]}\n'
            reward = '{:,}'.format(int(item[1]) * price)
            message += f'報酬：¥{reward}\n'
            message += f'媒体名：{code[item[0]]}\n'
    return {
        'new': n,
        'total': total,
        'total_price': total_price,
        'message': message
    }


### main_script ###
if __name__ == '__main__':

    if len(sys.argv) > 1:
        if sys.argv[1] == 'cleanup':
            if os.path.exists('./csv/chapup/'):
                shutil.rmtree('./csv/chapup/')
            if os.path.exists('./data/chapup/'):
                shutil.rmtree('./data/chapup/')
            if os.path.exists('./csv/bresmile/'):
                shutil.rmtree('./csv/bresmile/')
            if os.path.exists('./data/bresmile/'):
                shutil.rmtree('./data/bresmile/')
            if os.path.exists('./csv/linka/'):
                shutil.rmtree('./csv/linka/')
            exit(0)

    new = 0
    total = 0
    total_price = 0
    body_message = ''

    code = dict()
    output = dict()
    try:
        res = getInfoLinkA()
        new += res['new']
        total += res['total']
        total_price += res['total_price']
        body_message += res['message']

        res = getInfoAfb()
        new += res['new']
        total += res['total']
        total_price += res['total_price']
        body_message += res['message']

        res = getInfoMedia('chapup', 'CHAPUP', 'chapup.jp', 22000)
        new += res['new']
        total += res['total']
        total_price += res['total_price']
        body_message += res['message']

        if new == 0:
            logger.info(f"ALL: No new achievements")
        else:
            message = "[info][title]【祝】新規成果発生のお知らせ！[/title]"
            message += f"新規で【{new}件】申込が発生しました。\n"
            total_price = '{:,}'.format(total_price)
            message += f"本日の累計発生件数 : 成果報酬は【{total}件 : ¥{total_price}】です。\n"
            message += body_message
            message += '[/info]'

            #print(message)
            sendChatworkNotification(message, 'CHAPUP')

        logger.info("achievement_notificator: Finish")
        handler.close()
        os.remove(f'log/{today.strftime("%Y-%m-%d")}_result.log')
        exit(0)
    except Exception as err:
        logger.debug(f'achievement_notificator: {err}')
        exit(1)
