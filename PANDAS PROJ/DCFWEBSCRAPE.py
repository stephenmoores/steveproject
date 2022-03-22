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
import re

from datetime import date

# False means switch to HTML code and True means switch to live website opener
livewebsite = True


def asktckr():

    if livewebsite == True:
        print("Welcome to Steve's DCF generator")
        TCKR = input("What is the ticker of the stock you would like to run a valuation on?  ")

    else:
        TCKR = 'NVDA'
        print("LIVEWEBSITE = FALSE , RETURNING NVIDIA'S INFORMATION AS TEST")

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

def scrapeNWC(TCKR):

    baseURL = 'https://finance.yahoo.com/quote/'
    mainURL = baseURL + TCKR + '/balance-sheet'
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
        BShtml = codecs.open("BS.htm", 'r', 'utf-8')
        BShtml = BShtml.read()

    soup = bs.BeautifulSoup(BShtml, 'lxml')

    NWCPull = soup.find_all("div", {'class':'Ta(c) Py(6px) Bxz(bb) BdB Bdc($seperatorColor) Miw(120px) Miw(100px)--pnclg D(tbc)'})
    NWC = NWCPull[14].text
    NWC2 = NWCPull[15].text

    soup = bs.BeautifulSoup(BShtml, 'lxml')
    NWCPull = soup.find_all("div", {'class':'Ta(c) Py(6px) Bxz(bb) BdB Bdc($seperatorColor) Miw(120px) Miw(100px)--pnclg Bgc($lv1BgColor) fi-row:h_Bgc($hoverBgColor) D(tbc)'})
    NWC1 = NWCPull[14].text


    return NWC, NWC1, NWC2, BShtml

def imnotacrook():

    if livewebsite == True:
        time.sleep(3)

    else:
        print('live website is not in use: no delay')

    return

def scrapeunits(IShtml):

    soup = bs.BeautifulSoup(IShtml, 'lxml')
    UNITSPull = soup.find_all("span", {'class':'Fz(xs) C($tertiaryColor) Mstart(25px) smartphone_Mstart(0px) smartphone_D(b) smartphone_Mt(5px)'})
    UNITS = UNITSPull[0].text

    return UNITS

def scrapeRF():
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.131 Safari/537.36'}
    ticker = '%5ETNX'
    baseURL = 'https://finance.yahoo.com/quote/'
    mainURL = baseURL + ticker

    if livewebsite == True:
        # TURN BACK ON WHEN READY FOR PRODUCTION CODE
        res = requests.get(mainURL, headers)
        res = res.text

    else:
        # USE FOR TESTING CODE WITHOUT PINGING WEBSITE
        res = codecs.open("RFHOME.htm", 'r', 'utf-8')
        res = res.read()

    soup = bs.BeautifulSoup(res, 'lxml')
    RF = soup.find_all("fin-streamer", {'data-test':'qsp-price'})
    RF = RF[0].text
    RF = float(RF)/100

    return RF

def scrapeHOME(TCKR):

    baseURL = 'https://finance.yahoo.com/quote/'
    mainURL = baseURL + TCKR
    print(mainURL)

    if livewebsite == True:

        # TURN BACK ON WHEN READY FOR PRODUCTION CODE
        driver = webdriver.Chrome('C:/Users/Stephen/Downloads/SPRING SEMESTER 2021-2022/PYTHON/GitHub/steveproject/PANDAS PROJ/chromedriver_win32/chromedriver.exe')
        # return the main page
        driver.get(mainURL)
        # saving the HTML script into the python application just like beautiful soup
        HOMEhtml = driver.execute_script('return document.body.innerHTML;')
        # closing the driver
        driver.close()

    else:
        # USE FOR TESTING CODE WITHOUT PINGING WEBSITE
        HOMEhtml = codecs.open("HOME.htm", 'r', 'utf-8')
        HOMEhtml = HOMEhtml.read()

    soup = bs.BeautifulSoup(HOMEhtml, 'lxml')
    BETAPull = soup.find_all("td", {'data-test':'BETA_5Y-value'})
    BETA = BETAPull[0].text

    soup = bs.BeautifulSoup(HOMEhtml, 'lxml')
    NAMEPull = soup.find_all("h1", {'class': 'D(ib) Fz(18px)'})
    NAME = NAMEPull[0].text

    soup = bs.BeautifulSoup(HOMEhtml, 'lxml')
    PRICEPull = soup.find_all("fin-streamer", {'class':'Fw(b) Fz(36px) Mb(-4px) D(ib)'})
    PRICE = PRICEPull[0].text

    return NAME, PRICE, BETA, HOMEhtml

def scrapeCOD(IShtml, BShtml):

    soup = bs.BeautifulSoup(IShtml, 'lxml')
    intPull = soup.find_all("div", {'class':'Ta(c) Py(6px) Bxz(bb) BdB Bdc($seperatorColor) Miw(120px) Miw(100px)--pnclg D(tbc)'})
    int = intPull[40].text

    soup = bs.BeautifulSoup(BShtml, 'lxml')
    debtPull = soup.find_all("div", {'class':'Ta(c) Py(6px) Bxz(bb) BdB Bdc($seperatorColor) Miw(120px) Miw(100px)--pnclg D(tbc)'})
    TD = debtPull[20].text

    int = int.replace(",", "")
    TD = TD.replace(",", "")

    int = float(int)
    TD = float(TD)

    COD = int/TD

    return int, TD, COD, IShtml, BShtml

def executescript():
    TCKR = asktckr()
    Y, Y1, Y2, currentdate, IShtml = scrapedates(TCKR)
    imnotacrook()
    capex, capex1, capex2, CFShtml = scrapeCAPEX(TCKR)
    imnotacrook()
    NWC, NWC1, NWC2, BShtml = scrapeNWC(TCKR)
    imnotacrook()
    NAME, PRICE, BETA, HOMEhtml = scrapeHOME(TCKR)
    imnotacrook()
    RF = scrapeRF()
    Rev, Rev1, Rev2, IShtml = scraperev(IShtml)
    EBIT, EBIT1, EBIT2, IShtml = scrapeEBIT(IShtml)
    dep, dep1, dep2, IShtml = scrapedep(IShtml)
    EBITDA, EBITDA1, EBITDA2 = EBITDACALC(dep, dep1, dep2, EBIT, EBIT1, EBIT2)
    pretax, pretax1, pretax2, IShtml = scrapepretax(IShtml)
    tax, tax1, tax2, IShtml = scrapetax(IShtml)
    UNITS = scrapeunits(IShtml)
    int, TD, COD, IShtml, BShtml = scrapeCOD(IShtml, BShtml)

    # print(Y, Y1, Y2, currentdate)
    # print(Rev, Rev1, Rev2)
    # print(EBIT, EBIT1, EBIT2)
    # print(dep, dep1, dep2)
    # print(EBITDA, EBITDA1, EBITDA2)
    # print(pretax, pretax1, pretax2)
    # print(tax, tax1, tax2)
    # print(capex, capex1, capex2)
    # print(NWC, NWC1, NWC2)
    # print(UNITS)
    # print(RF)
    # print(NAME, PRICE, BETA)
    # print(int, TD, COD)
    return

if __name__ == "__main__":
    executescript()