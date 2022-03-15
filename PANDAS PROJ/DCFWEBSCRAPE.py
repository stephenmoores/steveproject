import openpyxl
import pandas as pd
import bs4 as bs
import requests
import codecs
import lxml
import xlwings as xw
from selenium import webdriver

# False means switch to HTML code and True means switch to live website opener
livewebsite = False


def asktckr():
    print("Welcome to Steve's DCF generator")
    TCKR = input("What is the ticker of the stock you would like to run a valuation on?  ")
    return TCKR

def scrapedates(TCKR):
    headers = {'User-Agent':'custom'}
    baseURL = 'https://finance.yahoo.com/quote/'
    mainURL = baseURL + TCKR + '/financials'
    print(mainURL)

    if livewebsite == True:

        # TURN BACK ON WHEN READY FOR PRODUCTION CODE
        driver = webdriver.Chrome('C:/Users/Stephen/Downloads/SPRING SEMESTER 2021-2022/PYTHON/GitHub/steveproject/PANDAS PROJ/chromedriver_win32/chromedriver.exe')
        # return the main page
        driver.get(mainURL)
        # saving the HTML script into the python application just like beautiful soup
        IShtml = driver.execute_script('return document.body.innerHTML;')
        # closing the driver
        driver.close()

    else:
        # USE FOR TESTING CODE WITHOUT PINGING WEBSITE
        IShtml = codecs.open("IS.htm", 'r', 'utf-8')
        IShtml = IShtml.read()

    soup = bs.BeautifulSoup(IShtml, 'lxml')
    datePull = soup.find_all("div", {'class':'Ta(c) Py(6px) Bxz(bb) BdB Bdc($seperatorColor) Miw(120px) Miw(100px)--pnclg D(ib) Fw(b)'})
    Y = datePull[0].text
    Y2 = datePull[1].text

    soup = bs.BeautifulSoup(IShtml, 'lxml')
    datePull = soup.find_all("div", {'class':'Ta(c) Py(6px) Bxz(bb) BdB Bdc($seperatorColor) Miw(120px) Miw(100px)--pnclg D(ib) Fw(b) Bgc($lv1BgColor)'})
    Y1 = datePull[0].text


    return Y, Y1, Y2, IShtml

def scraperev(TCKR, IShtml):
    headers = {'User-Agent':'custom'}
    baseURL = 'https://finance.yahoo.com/quote/'
    mainURL = baseURL + TCKR + '/financials'
    print(mainURL)


    soup = bs.BeautifulSoup(IShtml, 'lxml')
    revPull = soup.find_all("div", {'class':'Ta(c) Py(6px) Bxz(bb) BdB Bdc($seperatorColor) Miw(120px) Miw(100px)--pnclg D(ib) Fw(b)'})
    Rev = revPull[0].text
    Rev2 = revPull[1].text

    soup = bs.BeautifulSoup(IShtml, 'lxml')
    revPull = soup.find_all("div", {'class':'Ta(c) Py(6px) Bxz(bb) BdB Bdc($seperatorColor) Miw(120px) Miw(100px)--pnclg D(ib) Fw(b) Bgc($lv1BgColor)'})
    Rev1 = revPull[0].text


    return Rev, Rev1, Rev2, IShtml

def executescript():
    TCKR = asktckr()
    Y, Y1, Y2, IShtml = scrapedates(TCKR)
    Rev, Rev1, Rev2, IShtml = scraperev(TCKR, IShtml)
    print(Y, Y1, Y2)
    print(Rev, Rev1, Rev2)
    return

if __name__ == "__main__":
    executescript()