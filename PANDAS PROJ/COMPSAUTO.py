import pandas as pd
import bs4 as bs
from selenium import webdriver
import time
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import xlwings as xw
from selenium.webdriver.chrome.options import Options



livewebsite = True

chrome_options = Options()
chrome_options.add_argument('--disable-dev-shm-usage')
chrome_options.add_argument('--no-sandbox')

def wait():

    if livewebsite == True:
        time.sleep(3)
        print('three second delay to maintain low frequency of pings')

    else:
        print('live website is not in use: no delay')

    return

def asktckr2():

    if livewebsite == True:
        print("Welcome to Steve's comps analysis")
        print('Input ticker in all caps for example: NVDA')
        TCKR = input("What is the ticker of the stock you would like to run a valuation on?  ")

    else:
        TCKR = 'NVDA'
        print("LIVEWEBSITE = FALSE , RETURNING NVIDIA'S INFORMATION (as of February, 2022) AS TEST")

    return TCKR

def createdicts():

    tckrdict = {}
    PElist = []
    PBlist = []
    PSlist = []
    QUICKlist = []
    CURRENTlist = []
    DTOElist = []
    ROAlist = []
    ROElist = []
    ROIlist = []
    GMlist = []

    return tckrdict, PElist, PBlist, PSlist, QUICKlist, CURRENTlist, DTOElist, ROAlist, ROElist, ROIlist, GMlist

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

def getcountry(MWHOME):

    soup = bs.BeautifulSoup(MWHOME, 'lxml')
    Pull = soup.find_all("span", {'itemprop':'name'})
    country = Pull[4].text
    print(country)

    return country

def getticker(MWHOME):

    soup = bs.BeautifulSoup(MWHOME, 'lxml')
    Pull = soup.find_all("span", {'itemprop':'name'})
    compTCKR = Pull[5].text
    print(compTCKR)

    return  compTCKR

def getratiohome(compTCKR):

    MNURL = 'https://finviz.com/quote.ashx?t=' + compTCKR

    if livewebsite == True:
        # TURN BACK ON WHEN READY FOR PRODUCTION CODE
        driver = webdriver.Chrome('C:/Users/Stephen/Downloads/SPRING SEMESTER 2021-2022/PYTHON/GitHub/steveproject/PANDAS PROJ/chromedriver_win32/chromedriver.exe')
        driver.maximize_window()
        # return the main page
        driver.get(MNURL)
        # wait
        wait()
        # saving the HTML script into the python application just like beautiful soup
        compHOME = driver.execute_script('return document.body.innerHTML;')
        # closing the driver
        driver.close()

    return compHOME

def getratios(compHOME):

    soup = bs.BeautifulSoup(compHOME, 'lxml')
    Pull = soup.find_all("td", {'class':'snapshot-td2'})
    PE = Pull[1].text
    PB = Pull[25].text
    PS = Pull[19].text
    QUICK = Pull[43].text
    CURRENT = Pull[49].text
    DTOE = Pull[55].text
    ROA = Pull[27].text
    ROE = Pull[33].text
    ROI = Pull[39].text
    GM = Pull[45].text
    print(PE, PB, PS, QUICK, CURRENT, DTOE, ROA, ROE, ROI, GM)

    return PE, PB, PS, QUICK, CURRENT, DTOE, ROA, ROE, ROI, GM

def collectMAIN(TCKR):

    MNURL = 'https://finviz.com/quote.ashx?t=' + TCKR

    if livewebsite == True:
        # TURN BACK ON WHEN READY FOR PRODUCTION CODE
        driver = webdriver.Chrome('C:/Users/Stephen/Downloads/SPRING SEMESTER 2021-2022/PYTHON/GitHub/steveproject/PANDAS PROJ/chromedriver_win32/chromedriver.exe')
        driver.maximize_window()
        # return the main page
        driver.get(MNURL)
        # wait
        wait()
        # saving the HTML script into the python application just like beautiful soup
        MAINHOME = driver.execute_script('return document.body.innerHTML;')
        # closing the driver
        driver.close()

    soup = bs.BeautifulSoup(MAINHOME, 'lxml')
    Pull = soup.find_all("td", {'class':'snapshot-td2'})
    MAINPE = Pull[1].text
    MAINPB = Pull[25].text
    MAINPS = Pull[19].text
    MAINQUICK = Pull[43].text
    MAINCURRENT = Pull[49].text
    MAINDTOE = Pull[55].text
    MAINROA = Pull[27].text
    MAINROE = Pull[33].text
    MAINROI = Pull[39].text
    MAINGM = Pull[45].text

    listMAIN = [TCKR, MAINPE, MAINPB, MAINPS, MAINQUICK, MAINCURRENT, MAINDTOE, MAINROA, MAINROE, MAINROI, MAINGM]

    return MAINHOME, listMAIN

def repeat_collect(TCKR, tckrdict, PElist, PBlist, PSlist, QUICKlist, CURRENTlist, DTOElist, ROAlist, ROElist, ROIlist, GMlist):

    try:
        for _ in range(2):
            MWHOME, count = scrapecomp(TCKR)
            country = getcountry(MWHOME)

            if country == 'United States':
                compTCKR = getticker(MWHOME)
                compHOME = getratiohome(compTCKR)
                PE, PB, PS, QUICK, CURRENT, DTOE, ROA, ROE, ROI, GM = getratios(compHOME)


            else:
                compTCKR = None
                PE = None
                PB = None
                PS = None
                QUICK = None
                CURRENT = None
                DTOE = None
                ROA = None
                ROE = None
                ROI = None
                GM = None

            tckrdict.update({count : compTCKR})
            PElist.append([PE])
            PBlist.append([PB])
            PSlist.append([PS])
            QUICKlist.append([QUICK])
            CURRENTlist.append([CURRENT])
            DTOElist.append([DTOE])
            ROAlist.append([ROA])
            ROElist.append([ROE])
            ROIlist.append([ROI])
            GMlist.append([GM])

    except NoSuchElementException:
        print("end comps")

    return tckrdict, PElist, PBlist, PSlist, QUICKlist, CURRENTlist, DTOElist, ROAlist, ROElist, ROIlist, GMlist

def executescript2():
    tckrdict, PElist, PBlist, PSlist, QUICKlist, CURRENTlist, DTOElist, ROAlist, ROElist, ROIlist, GMlist = createdicts()
    scrapecomp.counter = 0
    TCKR = asktckr2()
    tckrdict, PElist, PBlist, PSlist, QUICKlist, CURRENTlist, DTOElist, ROAlist, ROElist, ROIlist, GMlist = repeat_collect(TCKR, tckrdict, PElist, PBlist, PSlist, QUICKlist, CURRENTlist, DTOElist, ROAlist, ROElist, ROIlist, GMlist)
    MAINHOME, listMAIN = collectMAIN(TCKR)

    # export to file
    wb = xw.Book('RECDCF.xlsx')
    sheet = wb.sheets['DATA_COMPS']
    sheet.range('A2:L12').clear_contents()
    sheet.range('B2').value = listMAIN
    sheet.range('A3').value = tckrdict
    sheet.range('C3').value = PElist
    sheet.range('D3').value = PBlist
    sheet.range('E3').value = PSlist
    sheet.range('F3').value = QUICKlist
    sheet.range('G3').value = CURRENTlist
    sheet.range('H3').value = DTOElist
    sheet.range('I3').value = ROAlist
    sheet.range('J3').value = ROElist
    sheet.range('K3').value = ROIlist
    sheet.range('L3').value = GMlist

    print('Comps data successfully exported into Excel File')

    return

if __name__ == "__main__":
    executescript2()


