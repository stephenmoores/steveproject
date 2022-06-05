from DCFWEBSCRAPE import *
import pandas as pd
import bs4 as bs
import xlwings as xw
import requests
import codecs
from selenium import webdriver
import time
from datetime import date
import DCFWEBSCRAPE
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def createdict():

    dict2 = {}

    return dict2

def scrapecomp(TCKR):

    MNURL = 'https://www.marketwatch.com/investing/stock/' + TCKR

    if livewebsite == True:
        # TURN BACK ON WHEN READY FOR PRODUCTION CODE
        driver = webdriver.Chrome('C:/Users/Stephen/Downloads/SPRING SEMESTER 2021-2022/PYTHON/GitHub/steveproject/PANDAS PROJ/chromedriver_win32/chromedriver.exe')
        driver.maximize_window()
        # return the main page
        driver.get(MNURL)
        # clicking the first comp
        scrapecomp.counter += 1
        count = scrapecomp.counter
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
        print('livewebsite is set to false')

    return MWHOME, count

def repeat_collect(TCKR, dict2):

    try:
        for _ in range(50):
            MWHOME, count = scrapecomp(TCKR)
            PE = getpe(MWHOME)
            dict2.update({count : PE})

    except NoSuchElementException:
        print("end comps")

    return dict2

def getpe(MWHOME):

    soup = bs.BeautifulSoup(MWHOME, 'lxml')
    Pull = soup.find_all("span", {'class':'primary'})
    PE = Pull[14].text

    return PE

def changetodf(dict2):

    df2 = pd.DataFrame([dict2])
    df2 = df2.T

    return df2

def executescript2():
    dict2 = createdict()
    scrapecomp.counter = 0
    TCKR = asktckr()
    dict2 = repeat_collect(TCKR, dict2)
    df2 = changetodf(dict2)
    wb = xw.Book('RECDCF.xlsx')
    sheet = wb.sheets['DATA_COMPS']
    sheet.range('A1').value = df2
    print('Comps data successfully exported into Excel File')
    return

executescript2()
