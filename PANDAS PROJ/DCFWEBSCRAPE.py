import bs4 as bs
import requests
import codecs
from selenium import webdriver
import time
from datetime import date

# False means switch to HTML code and True means switch to live website opener
livewebsite = True

def asktckr():

    if livewebsite == True:
        print("Welcome to Steve's DCF generator, this code currently works with companies that are not: banks, insurance providers, companies who do not report in USD")
        print('Input ticker in all caps for example: NVDA')
        TCKR = input("What is the ticker of the stock you would like to run a valuation on?  ")

    else:
        TCKR = 'NVDA'
        print("LIVEWEBSITE = FALSE , RETURNING NVIDIA'S INFORMATION (as of February, 2022) AS TEST")

    return TCKR

def askpga():

    if livewebsite == True:
        PGA = input("What is your assumed growth rate at perpetuity? (input as a decimal EX: 0.03 percent for 3%)  ")

    else:
        PGA = 0.03
        print("LIVEWEBSITE = FALSE , USING 3.0% AS THE PERPETUITY GROWTH RATE")

    return PGA

def askEM():

    if livewebsite == True:
        EM = input("What is your assumed exit multiple? (more patches coming to improve this feature)  ")

    else:
        EM = 25
        print("LIVEWEBSITE = FALSE , USING 25 AS THE EXIT MULTIPLE")

    return EM

def cleanscrape(var, DENOM, UNITS):

    if DENOM == 'M' or DENOM == 'B' or DENOM == 'K' or DENOM == 'T':
        var = var[:-1]
        var = float(var)

        if DENOM == 'T':
            var = var * 1000000000000

        if DENOM == 'B':
            var = var * 1000000000

        elif DENOM == 'M':
            var = var * 1000000

        elif DENOM == 'K':
            var = var * 1000

    if DENOM == ')':
        # getting what units the number is displayed as in marketwatch
        DENOM = var[-2]
        # getting all characters between the letter and the parenthesis
        var = var[1:-2]
        # converting var to a float so we can work with it
        var = float(var)
        # this if statement only should be triggered in the case of a negative number therefore we convert the float to a negative float
        var = var*-1

        if DENOM == 'T':
            var = var * 1000000000000

        if DENOM == 'B':
            var = var * 1000000000

        if DENOM == 'M':
            var = var * 1000000

        if DENOM == 'K':
            var = var * 1000

    if DENOM == '-':
        var = 0

    if UNITS == 'All numbers in thousands':
        var = var / 1000

    elif UNITS == 'All numbers in millions':
        var = var / 1000000

    return var, DENOM

def scrapedatesunits(TCKR):

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
        IShtml = codecs.open("DATE.htm", 'r', 'utf-8')
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

def scraperev(TCKR, UNITS):

    baseURL = 'https://www.marketwatch.com/investing/stock/'
    mainURL = baseURL + TCKR + '/financials'
    print(mainURL)

    if livewebsite == True:

        # TURN BACK ON WHEN READY FOR PRODUCTION CODE
        driver = webdriver.Chrome('C:/Users/Stephen/Downloads/SPRING SEMESTER 2021-2022/PYTHON/GitHub/steveproject/PANDAS PROJ/chromedriver_win32/chromedriver.exe')
        # return the main page
        driver.get(mainURL)
        # saving the HTML script into the python application just like beautiful soup
        MWIShtml = driver.execute_script('return document.body.innerHTML;')
        # closing the driver
        driver.close()

    else:
        # USE FOR TESTING CODE WITHOUT PINGING WEBSITE
        MWIShtml = codecs.open("MWIS.htm", 'r', 'utf-8')
        MWIShtml = MWIShtml.read()

    soup = bs.BeautifulSoup(MWIShtml, 'lxml')
    RevPull = soup.find_all('div', {'class': 'cell__content'})
    Rev = RevPull[14].text
    DENOM = Rev[-1]
    # CLEANING THE DATA
    var = Rev
    var, DENOM = cleanscrape(var, DENOM, UNITS)
    Rev = var

    soup = bs.BeautifulSoup(MWIShtml, 'lxml')
    Rev1Pull = soup.find_all('div', {'class': 'cell__content'})
    Rev1 = Rev1Pull[13].text
    DENOM = Rev1[-1]
    # CLEANING THE DATA
    var = Rev1
    var, DENOM = cleanscrape(var, DENOM, UNITS)
    Rev1 = var


    soup = bs.BeautifulSoup(MWIShtml, 'lxml')
    Rev2Pull = soup.find_all('div', {'class': 'cell__content'})
    Rev2 = Rev2Pull[12].text
    DENOM = Rev2[-1]
    # CLEANING THE DATA
    var = Rev2
    var, DENOM = cleanscrape(var, DENOM, UNITS)
    Rev2 = var

    return Rev, Rev1, Rev2, MWIShtml, UNITS

def scrapeEBITDA(MWIShtml, UNITS):

    soup = bs.BeautifulSoup(MWIShtml, 'lxml')
    EBITDAPull = soup.find_all('div', {'class': 'cell__content'})
    EBITDA = EBITDAPull[446].text
    DENOM = EBITDA[-1]

    var = EBITDA
    var, DENOM = cleanscrape(var, DENOM, UNITS)
    EBITDA = var

    soup = bs.BeautifulSoup(MWIShtml, 'lxml')
    EBITDA1Pull = soup.find_all('div', {'class': 'cell__content'})
    EBITDA1 = EBITDA1Pull[445].text
    DENOM = EBITDA1[-1]

    var = EBITDA1
    var, DENOM = cleanscrape(var, DENOM, UNITS)
    EBITDA1 = var

    soup = bs.BeautifulSoup(MWIShtml, 'lxml')
    EBITDA2Pull = soup.find_all('div', {'class': 'cell__content'})
    EBITDA2 = EBITDA2Pull[444].text
    DENOM = EBITDA2[-1]

    var = EBITDA2
    var, DENOM = cleanscrape(var, DENOM, UNITS)
    EBITDA2 = var


    return EBITDA, EBITDA1, EBITDA2, MWIShtml, UNITS

def scrapedepna(MWIShtml, UNITS):

    soup = bs.BeautifulSoup(MWIShtml, 'lxml')
    depnaPull = soup.find_all('div', {'class': 'cell__content'})
    depna = depnaPull[54].text
    DENOM = depna[-1]

    var = depna
    var, DENOM = cleanscrape(var, DENOM, UNITS)
    depna = var

    soup = bs.BeautifulSoup(MWIShtml, 'lxml')
    depna1Pull = soup.find_all('div', {'class': 'cell__content'})
    depna1 = depna1Pull[53].text
    DENOM = depna1[-1]

    var = depna1
    var, DENOM = cleanscrape(var, DENOM, UNITS)
    depna1 = var

    soup = bs.BeautifulSoup(MWIShtml, 'lxml')
    depna2Pull = soup.find_all('div', {'class': 'cell__content'})
    depna2 = depna2Pull[52].text
    DENOM = depna2[-1]

    var = depna2
    var, DENOM = cleanscrape(var, DENOM, UNITS)
    depna2 = var

    return depna, depna1, depna2, MWIShtml, UNITS

def calcEBIT(EBITDA, EBITDA1, EBITDA2, depna, depna1, depna2):

    EBIT = EBITDA - depna
    EBIT1 = EBITDA1 - depna1
    EBIT2 = EBITDA2 - depna2

    return EBIT, EBIT1, EBIT2

def scrapepretax(MWIShtml, UNITS):

    soup = bs.BeautifulSoup(MWIShtml, 'lxml')
    pretaxPull = soup.find_all('div', {'class': 'cell__content'})
    pretax = pretaxPull[214].text
    DENOM = pretax[-1]

    var = pretax
    var, DENOM = cleanscrape(var, DENOM, UNITS)
    pretax = var

    soup = bs.BeautifulSoup(MWIShtml, 'lxml')
    pretax1Pull = soup.find_all('div', {'class': 'cell__content'})
    pretax1 = pretax1Pull[213].text
    DENOM = pretax1[-1]

    var = pretax1
    var, DENOM = cleanscrape(var, DENOM, UNITS)
    pretax1 = var

    soup = bs.BeautifulSoup(MWIShtml, 'lxml')
    pretax2Pull = soup.find_all('div', {'class': 'cell__content'})
    pretax2 = pretax2Pull[212].text
    DENOM = pretax2[-1]

    var = pretax2
    var, DENOM = cleanscrape(var, DENOM, UNITS)
    pretax2 = var

    return pretax, pretax1, pretax2, MWIShtml, UNITS

def scrapetax(MWIShtml, UNITS):

    soup = bs.BeautifulSoup(MWIShtml, 'lxml')
    taxPull = soup.find_all('div', {'class': 'cell__content'})
    tax = taxPull[238].text
    DENOM = tax[-1]

    var = tax
    var, DENOM = cleanscrape(var, DENOM, UNITS)
    tax = var

    soup = bs.BeautifulSoup(MWIShtml, 'lxml')
    tax1Pull = soup.find_all('div', {'class': 'cell__content'})
    tax1 = tax1Pull[237].text
    DENOM = tax1[-1]

    var = tax1
    var, DENOM = cleanscrape(var, DENOM, UNITS)
    tax1 = var

    soup = bs.BeautifulSoup(MWIShtml, 'lxml')
    tax2Pull = soup.find_all('div', {'class': 'cell__content'})
    tax2 = tax2Pull[236].text
    DENOM = tax2[-1]

    var = tax2
    var, DENOM = cleanscrape(var, DENOM, UNITS)
    tax2 = var

    return tax, tax1, tax2, MWIShtml, UNITS

def scrapeCAPEX(TCKR, UNITS):

    baseURL = 'https://www.marketwatch.com/investing/stock/'
    mainURL = baseURL + TCKR + '/financials/cash-flow'
    print(mainURL)

    if livewebsite == True:

        # TURN BACK ON WHEN READY FOR PRODUCTION CODE
        driver = webdriver.Chrome('C:/Users/Stephen/Downloads/SPRING SEMESTER 2021-2022/PYTHON/GitHub/steveproject/PANDAS PROJ/chromedriver_win32/chromedriver.exe')
        # return the main page
        driver.get(mainURL)
        # saving the HTML script into the python application just like beautiful soup
        MWCFShtml = driver.execute_script('return document.body.innerHTML;')
        # closing the driver
        driver.close()

    else:
        # USE FOR TESTING CODE WITHOUT PINGING WEBSITE
        MWCFShtml = codecs.open("MWCFS.htm", 'r', 'utf-8')
        MWCFShtml = MWCFShtml.read()

    soup = bs.BeautifulSoup(MWCFShtml, 'lxml')
    capexPull = soup.find_all('div', {'class': 'cell__content'})
    capex = capexPull[166].text
    DENOM = capex[-1]

    var = capex
    var, DENOM = cleanscrape(var, DENOM, UNITS)
    capex = var

    soup = bs.BeautifulSoup(MWCFShtml, 'lxml')
    capex1Pull = soup.find_all('div', {'class': 'cell__content'})
    capex1 = capex1Pull[165].text
    DENOM = capex1[-1]

    var = capex1
    var, DENOM = cleanscrape(var, DENOM, UNITS)
    capex1 = var

    soup = bs.BeautifulSoup(MWCFShtml, 'lxml')
    capex2Pull = soup.find_all('div', {'class': 'cell__content'})
    capex2 = capex2Pull[164].text
    DENOM = capex2[-1]

    var = capex2
    var, DENOM = cleanscrape(var, DENOM, UNITS)
    capex2 = var

    return capex, capex1, capex2, MWCFShtml

def scrapeNWC(TCKR, UNITS):

    baseURL = 'https://www.marketwatch.com/investing/stock/'
    mainURL = baseURL + TCKR + '/financials/balance-sheet'
    print(mainURL)

    if livewebsite == True:

        # TURN BACK ON WHEN READY FOR PRODUCTION CODE
        driver = webdriver.Chrome('C:/Users/Stephen/Downloads/SPRING SEMESTER 2021-2022/PYTHON/GitHub/steveproject/PANDAS PROJ/chromedriver_win32/chromedriver.exe')
        # return the main page
        driver.get(mainURL)
        # saving the HTML script into the python application just like beautiful soup
        MWBShtml = driver.execute_script('return document.body.innerHTML;')
        # closing the driver
        driver.close()

    else:
        # USE FOR TESTING CODE WITHOUT PINGING WEBSITE
        MWBShtml = codecs.open("MWBS.htm", 'r', 'utf-8')
        MWBShtml = MWBShtml.read()

    soup = bs.BeautifulSoup(MWBShtml, 'lxml')
    CAPull = soup.find_all('div', {'class': 'cell__content'})
    CA = CAPull[166].text
    DENOM = CA[-1]

    var = CA
    var, DENOM = cleanscrape(var, DENOM, UNITS)
    CA = var

    soup = bs.BeautifulSoup(MWBShtml, 'lxml')
    CA1Pull = soup.find_all('div', {'class': 'cell__content'})
    CA1 = CA1Pull[165].text
    DENOM = CA1[-1]

    var = CA1
    var, DENOM = cleanscrape(var, DENOM, UNITS)
    CA1 = var

    soup = bs.BeautifulSoup(MWBShtml, 'lxml')
    CA2Pull = soup.find_all('div', {'class': 'cell__content'})
    CA2 = CA2Pull[164].text
    DENOM = CA2[-1]

    var = CA2
    var, DENOM = cleanscrape(var, DENOM, UNITS)
    CA2 = var

    soup = bs.BeautifulSoup(MWBShtml, 'lxml')
    CLPull = soup.find_all('div', {'class': 'cell__content'})
    CL = CLPull[390].text
    DENOM = CL[-1]

    var = CL
    var, DENOM = cleanscrape(var, DENOM, UNITS)
    CL = var

    soup = bs.BeautifulSoup(MWBShtml, 'lxml')
    CL1Pull = soup.find_all('div', {'class': 'cell__content'})
    CL1 = CL1Pull[389].text
    DENOM = CL1[-1]

    var = CL1
    var, DENOM = cleanscrape(var, DENOM, UNITS)
    CL1 = var

    soup = bs.BeautifulSoup(MWBShtml, 'lxml')
    CL2Pull = soup.find_all('div', {'class': 'cell__content'})
    CL2 = CL2Pull[388].text
    DENOM = CL2[-1]

    var = CL2
    var, DENOM = cleanscrape(var, DENOM, UNITS)
    CL2 = var

    NWC = CA - CL
    NWC1 = CA1 - CL1
    NWC2 = CA2 - CL2

    return NWC, NWC1, NWC2, MWBShtml

def wait():

    if livewebsite == True:
        time.sleep(6)
        print('three second delay to maintain low frequency of pings')

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
    print(mainURL)

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

def scrapeHOME(TCKR, UNITS):

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
    PRICE = float(PRICE.replace(',', ''))

    soup = bs.BeautifulSoup(HOMEhtml, 'lxml')
    CAPPull = soup.find_all("td", {'data-test':'MARKET_CAP-value'})
    CAP = CAPPull[0].text
    DENOM = CAP[-1]

    var = CAP

    if DENOM == 'M' or DENOM == 'B' or DENOM == 'K' or DENOM == 'T':
        var = var[:-1]
        var = float(var)

        if DENOM == 'T':
            var = var * 1000000000000

        elif DENOM == 'B':
            var = var * 1000000000

        elif DENOM == 'M':
            var = var * 1000000

        elif DENOM == 'K':
            var = var * 1000

    if UNITS == 'All numbers in thousands':
        var = var / 1000

    elif UNITS == 'All numbers in millions':
        var = var / 1000000

    CAP = var

    SHARES = CAP/PRICE

    return NAME, PRICE, BETA, CAP, SHARES, HOMEhtml

def scrapeCOD(MWIShtml, MWBShtml, UNITS):

    soup = bs.BeautifulSoup(MWIShtml, 'lxml')
    interestPull = soup.find_all('div', {'class': 'cell__content'})
    interest = interestPull[182].text
    DENOM = interest[-1]

    var = interest
    var, DENOM = cleanscrape(var, DENOM, UNITS)
    interest = var

    soup = bs.BeautifulSoup(MWBShtml, 'lxml')
    STDPull = soup.find_all('div', {'class': 'cell__content'})
    STD = STDPull[318].text
    DENOM = STD[-1]

    var = STD
    var, DENOM = cleanscrape(var, DENOM, UNITS)
    STD = var

    soup = bs.BeautifulSoup(MWBShtml, 'lxml')
    LTDPull = soup.find_all('div', {'class': 'cell__content'})
    LTD = LTDPull[398].text
    DENOM = LTD[-1]

    var = LTD
    var, DENOM = cleanscrape(var, DENOM, UNITS)
    LTD = var

    # ST AND LTD ARE SHORT TERM AND LONG TERM DEBT THEIR TOTAL MAKES TOTAL DEBT

    TD = STD + LTD
    COD = interest/TD

    return interest, TD, COD, MWIShtml, MWBShtml

def scrapeCCEandSTLandLTL(MWBShtml, UNITS):


    soup = bs.BeautifulSoup(MWBShtml, 'lxml')
    CCEPull = soup.find_all('div', {'class': 'cell__content'})
    CCE = CCEPull[14].text
    DENOM = CCE[-1]

    var = CCE
    var, DENOM = cleanscrape(var, DENOM, UNITS)
    CCE = var

    soup = bs.BeautifulSoup(MWBShtml, 'lxml')
    STLPull = soup.find_all('div', {'class': 'cell__content'})
    STL = STLPull[390].text
    DENOM = STL[-1]

    var = STL
    var, DENOM = cleanscrape(var, DENOM, UNITS)
    STL = var

    soup = bs.BeautifulSoup(MWBShtml, 'lxml')
    TLPull = soup.find_all('div', {'class': 'cell__content'})
    TL = TLPull[494].text
    DENOM = TL[-1]

    var = TL
    var, DENOM = cleanscrape(var, DENOM, UNITS)
    TL = var

    LTL = TL - STL

    return CCE, STL, LTL, UNITS

def executescript():
    TCKR = asktckr()
    PGA = askpga()
    EM = askEM()
    Y, Y1, Y2, currentdate, IShtml = scrapedatesunits(TCKR)
    UNITS = scrapeunits(IShtml)
    wait()
    Rev, Rev1, Rev2, MWIShtml, UNITS = scraperev(TCKR, UNITS)
    wait()
    capex, capex1, capex2, MWCFShtml = scrapeCAPEX(TCKR, UNITS)
    wait()
    NWC, NWC1, NWC2, MWBShtml = scrapeNWC(TCKR, UNITS)
    wait()
    NAME, PRICE, BETA, CAP, SHARES, HOMEhtml = scrapeHOME(TCKR, UNITS)
    wait()
    RF = scrapeRF()
    CCE, STL, LTL, UNITS = scrapeCCEandSTLandLTL(MWBShtml, UNITS)
    EBITDA, EBITDA1, EBITDA2, MWIShtml, UNITS = scrapeEBITDA(MWIShtml, UNITS)
    depna, depna1, depna2, MWIShtml, UNITS = scrapedepna(MWIShtml, UNITS)
    EBIT, EBIT1, EBIT2 = calcEBIT(EBITDA, EBITDA1, EBITDA2, depna, depna1, depna2)
    pretax, pretax1, pretax2, MWIShtml, UNITS = scrapepretax(MWIShtml, UNITS)
    tax, tax1, tax2, MWIShtml, UNITS = scrapetax(MWIShtml, UNITS)
    interest, TD, COD, MWIShtml, MWBShtml = scrapeCOD(MWIShtml, MWBShtml, UNITS)

    return

if __name__ == "__main__":
    executescript()

