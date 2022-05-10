import bs4 as bs
import codecs
from selenium import webdriver

   # you can alter the URL it uses to scrape in the "askURL" function
   # you can alter the path it takes in BS4 in the "collect" function

livewebsite = True

def askurl():

    mainURL = "https://www.marketwatch.com/investing/stock/nvda/financials"

    return mainURL

def returnhtml(mainURL):

    if livewebsite == True:

        # TURN BACK ON WHEN READY FOR PRODUCTION CODE
        driver = webdriver.Chrome('C:/Users/Stephen/Downloads/SPRING SEMESTER 2021-2022/PYTHON/GitHub/steveproject/PANDAS PROJ/chromedriver_win32/chromedriver.exe')
        # return the main page
        driver.get(mainURL)
        # saving the HTML script into the python application just like beautiful soup
        html = driver.execute_script('return document.body.innerHTML;')
        # closing the driver
        driver.close()

    else:
        # USE FOR TESTING CODE WITHOUT PINGING WEBSITE
        html = codecs.open("MWIS.htm", 'r', 'utf-8')
        html = html.read()

    return html

def collect(html):

    collect.counter += 1
    count = collect.counter

    soup = bs.BeautifulSoup(html, 'lxml')
    Pull = soup.find_all("div", {'class':'cell__content'})
    var = Pull[count].text

    print(f"{var} - {collect.counter}")

    return

def repeat_collect(html):

    try:
        for _ in range(10000):
            collect(html)

    except IndexError:
        print("table end")

    return

def executescript():

    mainURL = askurl()
    html = returnhtml(mainURL)
    collect.counter = -1
    repeat_collect(html)

    return

executescript()