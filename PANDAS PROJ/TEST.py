import bs4 as bs
import requests
import codecs
from selenium import webdriver
import time
from datetime import date
from DCFWEBSCRAPE import wait
from DCFWEBSCRAPE import asktckr
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

livewebsite = True

def scrapecomp1(TCKR):

    MNURL = 'https://www.marketwatch.com/investing/stock/' + TCKR

    if livewebsite == True:

        # TURN BACK ON WHEN READY FOR PRODUCTION CODE
        driver = webdriver.Chrome('C:/Users/Stephen/Downloads/SPRING SEMESTER 2021-2022/PYTHON/GitHub/steveproject/PANDAS PROJ/chromedriver_win32/chromedriver.exe')
        # return the main page
        driver.get(MNURL)
        # clicking the first comp
        count = 2
        count = str(count)
        XPATH = '//*[@id="maincontent"]/div[7]/div[3]/div/div/table/tbody/tr[' + count + ']/td[1]/a'
        button = driver.find_element(By.XPATH, XPATH)
        driver.execute_script("arguments[0].click();", button)
        # wait
        wait()
        # saving the HTML script into the python application just like beautiful soup
        MWHOME = driver.execute_script('return document.body.innerHTML;')
        # closing the driver
        driver.close()

    else:
        # USE FOR TESTING CODE WITHOUT PINGING WEBSITE
        MWHOME = codecs.open("MWHOME.htm", 'r', 'utf-8')
        MWHOME = MWHOME.read()

    return MWHOME

def getpe(MWHOME):

    soup = bs.BeautifulSoup(MWHOME, 'lxml')
    Pull = soup.find_all("span", {'class':'primary'})
    PE = Pull[14].text

    return PE

def repeat_collect(MWHOME, dict2, TCKR):

    try:
        for _ in range(50):
            TICKER, count = scrapecomp1()
            dict2.update({count : TICKER})

    except IndexError:
        print("table end")

    return dict2

def executescript():
    TCKR = asktckr()
    repeat_collect(MWHOME, dict2, TCKR)
    PE = getpe(MWHOME)
    print(PE)
    return

executescript()