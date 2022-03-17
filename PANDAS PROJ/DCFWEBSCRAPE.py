import openpyxl
import pandas as pd
import bs4 as bs
import requests
import codecs
import lxml
import xlwings as xw
from selenium import webdriver
import datetime
import time

from datetime import date

# False means switch to HTML code and True means switch to live website opener
livewebsite = True


def asktckr():
    print("Welcome to Steve's DCF generator")
    TCKR = input("What is the ticker of the stock you would like to run a valuation on?  ")
    # TCKR = 'NVDA'
    return TCKR

def scrapedates(TCKR):

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

    currentdate = date.today()

    return Y, Y1, Y2, currentdate, IShtml

def scraperev(IShtml):

    soup = bs.BeautifulSoup(IShtml, 'lxml')
    revPull = soup.find_all("div", {'class':'Ta(c) Py(6px) Bxz(bb) BdB Bdc($seperatorColor) Miw(120px) Miw(100px)--pnclg D(tbc)'})
    Rev = revPull[0].text
    Rev2 = revPull[1].text

    soup = bs.BeautifulSoup(IShtml, 'lxml')
    revPull = soup.find_all("div", {'class':'Ta(c) Py(6px) Bxz(bb) BdB Bdc($seperatorColor) Miw(120px) Miw(100px)--pnclg Bgc($lv1BgColor) fi-row:h_Bgc($hoverBgColor) D(tbc)'})
    Rev1 = revPull[1].text


    return Rev, Rev1, Rev2, IShtml

def scrapeEBIT(IShtml):

    soup = bs.BeautifulSoup(IShtml, 'lxml')
    EBITPull = soup.find_all("div", {'class':'Ta(c) Py(6px) Bxz(bb) BdB Bdc($seperatorColor) Miw(120px) Miw(100px)--pnclg D(tbc)'})
    EBIT = EBITPull[44].text
    EBIT2 = EBITPull[45].text

    soup = bs.BeautifulSoup(IShtml, 'lxml')
    EBITPull = soup.find_all("div", {'class':'Ta(c) Py(6px) Bxz(bb) BdB Bdc($seperatorColor) Miw(120px) Miw(100px)--pnclg Bgc($lv1BgColor) fi-row:h_Bgc($hoverBgColor) D(tbc)'})
    EBIT1 = EBITPull[67].text


    return EBIT, EBIT1, EBIT2, IShtml

def scrapedep(IShtml):

    soup = bs.BeautifulSoup(IShtml, 'lxml')
    depPull = soup.find_all("div", {'class':'Ta(c) Py(6px) Bxz(bb) BdB Bdc($seperatorColor) Miw(120px) Miw(100px)--pnclg D(tbc)'})
    dep = depPull[50].text
    dep2 = depPull[51].text

    soup = bs.BeautifulSoup(IShtml, 'lxml')
    depPull = soup.find_all("div", {'class':'Ta(c) Py(6px) Bxz(bb) BdB Bdc($seperatorColor) Miw(120px) Miw(100px)--pnclg Bgc($lv1BgColor) fi-row:h_Bgc($hoverBgColor) D(tbc)'})
    dep1 = depPull[76].text

    return dep, dep1, dep2, IShtml

def EBITDACALC(dep, dep1, dep2, EBIT, EBIT1, EBIT2):
    dep = dep.replace(",", "")
    dep1 = dep1.replace(",", "")
    dep2 = dep2.replace(",", "")

    EBIT = EBIT.replace(",", "")
    EBIT1 = EBIT1.replace(",", "")
    EBIT2 = EBIT2.replace(",", "")

    dep = float(dep)
    dep1 = float(dep1)
    dep2 = float(dep2)

    EBIT = float(EBIT)
    EBIT1 = float(EBIT1)
    EBIT2 = float(EBIT2)

    EBITDA = EBIT + dep
    EBITDA1 = EBIT1 + dep1
    EBITDA2 = EBIT2 + dep2

    return EBITDA, EBITDA1, EBITDA2

def scrapepretax(IShtml):

    soup = bs.BeautifulSoup(IShtml, 'lxml')
    pretaxPull = soup.find_all("div", {'class':'Ta(c) Py(6px) Bxz(bb) BdB Bdc($seperatorColor) Miw(120px) Miw(100px)--pnclg D(tbc)'})
    pretax = pretaxPull[14].text
    pretax2 = pretaxPull[15].text

    soup = bs.BeautifulSoup(IShtml, 'lxml')
    pretaxPull = soup.find_all("div", {'class':'Ta(c) Py(6px) Bxz(bb) BdB Bdc($seperatorColor) Miw(120px) Miw(100px)--pnclg Bgc($lv1BgColor) fi-row:h_Bgc($hoverBgColor) D(tbc)'})
    pretax1 = pretaxPull[22].text

    return pretax, pretax1, pretax2, IShtml

def scrapetax(IShtml):

    soup = bs.BeautifulSoup(IShtml, 'lxml')
    taxPull = soup.find_all("div", {'class': 'Ta(c) Py(6px) Bxz(bb) BdB Bdc($seperatorColor) Miw(120px) Miw(100px)--pnclg D(tbc)'})
    tax = taxPull[16].text
    tax2 = taxPull[17].text

    soup = bs.BeautifulSoup(IShtml, 'lxml')
    taxPull = soup.find_all("div", {'class': 'Ta(c) Py(6px) Bxz(bb) BdB Bdc($seperatorColor) Miw(120px) Miw(100px)--pnclg Bgc($lv1BgColor) fi-row:h_Bgc($hoverBgColor) D(tbc)'})
    tax1 = taxPull[25].text

    return tax, tax1, tax2, IShtml

def scrapeCAPEX(TCKR):

    baseURL = 'https://finance.yahoo.com/quote/'
    mainURL = baseURL + TCKR + '/cash-flow'
    print(mainURL)

    if livewebsite == True:

        # TURN BACK ON WHEN READY FOR PRODUCTION CODE
        driver = webdriver.Chrome('C:/Users/Stephen/Downloads/SPRING SEMESTER 2021-2022/PYTHON/GitHub/steveproject/PANDAS PROJ/chromedriver_win32/chromedriver.exe')
        # return the main page
        driver.get(mainURL)
        # saving the HTML script into the python application just like beautiful soup
        CFShtml = driver.execute_script('return document.body.innerHTML;')
        # closing the driver
        driver.close()

    else:
        # USE FOR TESTING CODE WITHOUT PINGING WEBSITE
        CFShtml = codecs.open("CFS.htm", 'r', 'utf-8')
        CFShtml = CFShtml.read()

    soup = bs.BeautifulSoup(CFShtml, 'lxml')

    capexPull = soup.find_all("div", {'class':'Ta(c) Py(6px) Bxz(bb) BdB Bdc($seperatorColor) Miw(120px) Miw(100px)--pnclg D(tbc)'})
    capex = capexPull[12].text
    capex2 = capexPull[13].text

    soup = bs.BeautifulSoup(CFShtml, 'lxml')
    capexPull = soup.find_all("div", {'class':'Ta(c) Py(6px) Bxz(bb) BdB Bdc($seperatorColor) Miw(120px) Miw(100px)--pnclg Bgc($lv1BgColor) fi-row:h_Bgc($hoverBgColor) D(tbc)'})
    capex1 = capexPull[19].text

    return capex, capex1, capex2, CFShtml

def executescript():
    TCKR = asktckr()
    Y, Y1, Y2, currentdate, IShtml = scrapedates(TCKR)
    time.sleep(3)
    capex, capex1, capex2, CFShtml = scrapeCAPEX(TCKR)
    # time.sleep(10)
    Rev, Rev1, Rev2, IShtml = scraperev(IShtml)
    EBIT, EBIT1, EBIT2, IShtml = scrapeEBIT(IShtml)
    dep, dep1, dep2, IShtml = scrapedep(IShtml)
    EBITDA, EBITDA1, EBITDA2 = EBITDACALC(dep, dep1, dep2, EBIT, EBIT1, EBIT2)
    pretax, pretax1, pretax2, IShtml = scrapepretax(IShtml)
    tax, tax1, tax2, IShtml = scrapetax(IShtml)


    print(Y, Y1, Y2, currentdate)
    print(Rev, Rev1, Rev2)
    print(EBIT, EBIT1, EBIT2)
    print(dep, dep1, dep2)
    print(EBITDA, EBITDA1, EBITDA2)
    print(pretax, pretax1, pretax2)
    print(tax, tax1, tax2)
    print(capex, capex1, capex2)

    return

if __name__ == "__main__":
    executescript()