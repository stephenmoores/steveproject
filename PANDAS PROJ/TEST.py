import openpyxl
import pandas as pd
import bs4 as bs
import requests
import codecs
import lxml
import xlwings as xw
from selenium import webdriver
from selenium import webdriver
from bs4 import BeautifulSoup
import pandas as pd
import datetime
import time
import re
import webbrowser

livewebsite = True

def asktckr():

    if livewebsite == True:
        print("Welcome to Steve's DCF generator")
        TCKR = input("What is the ticker of the stock you would like to run a valuation on?  ")

    else:
        TCKR = 'NVDA'
        print("LIVEWEBSITE = FALSE , RETURNING NVIDIA'S INFORMATION AS TEST")

    return TCKR

def scrapeCCE(TCKR):
    baseURL = 'https://www.marketwatch.com/investing/stock/'
    mainURL = baseURL + TCKR + '/financials/balance-sheet'
    print(mainURL)

    if livewebsite == True:

        # TURN BACK ON WHEN READY FOR PRODUCTION CODE
        driver = webdriver.Chrome('C:/Users/Stephen/Downloads/SPRING SEMESTER 2021-2022/PYTHON/GitHub/steveproject/PANDAS PROJ/chromedriver_win32/chromedriver.exe')
        # return the main page
        driver.get(mainURL)
        # saving the HTML script into the python application just like beautiful soup
        BShtml = driver.execute_script('return document.body.innerHTML;')
        # closing the driver
        driver.close()

    else:
        # USE FOR TESTING CODE WITHOUT PINGING WEBSITE
        BShtml = codecs.open("MWBS.htm", 'r', 'utf-8')
        BShtml = BShtml.read()

    soup = bs.BeautifulSoup(BShtml, 'lxml')
    CCEPull = soup.find_all('div', {'class': 'cell__content'})
    CCE = CCEPull[14].text

    DENOM = CCE[-1]
    CCE = CCE[:-1]
    CCE = float(CCE)

    # getting the cash and cash equivalents to a dollar figure that we can later divide by the units to get useable data

    if DENOM == 'B':
        CCE = CCE*1000000000

    elif DENOM == 'M':
        CCE = CCE*1000000

    elif DENOM == 'K':
        CCE = CCE*1000

    else:
        print('unit error')

    return BShtml, CCE, DENOM



def test():

    return

def executescript():
    TCKR = asktckr()
    MWBShtml, CCE, DENOM = scrapeCCE(TCKR)

    print(CCE)
    print(DENOM)
    return

executescript()
