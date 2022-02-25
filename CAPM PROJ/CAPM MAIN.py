import bs4 as bs
import requests
import codecs
import lxml

print("Welcome to Steve's CAPM model")
print("### NOTE ### - Please input percentages as decimals (e.g., 2% = 0.02)")

def scrapeBeta(ticker):
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.131 Safari/537.36'}

    baseURL = 'https://finance.yahoo.com/quote/'
    mainURL = baseURL + ticker

    # TURN BACK ON WHEN READY FOR PRODUCTION CODE
    res = requests.get(mainURL, headers)
    res = res.text


    # USE FOR TESTING CODE WITHOUT PINGING WEBSITE
    # res = codecs.open("testSite.html", 'r', 'utf-8')
    # res = res.read()

    soup = bs.BeautifulSoup(res, 'lxml')
    beta = soup.find_all("td", {'data-test':'BETA_5Y-value'})
    beta = beta[0].text
    beta = float(beta)

    return beta

def scrapeRF():
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.131 Safari/537.36'}
    ticker = '%5ETNX'
    baseURL = 'https://finance.yahoo.com/quote/'
    mainURL = baseURL + ticker

    # TURN BACK ON WHEN READY FOR PRODUCTION CODE
    res = requests.get(mainURL, headers)
    res = res.text


    # USE FOR TESTING CODE WITHOUT PINGING WEBSITE
    # res = codecs.open("pogfile.html", 'r', 'utf-8')
    # res = res.read()

    soup = bs.BeautifulSoup(res, 'lxml')
    RF = soup.find_all("fin-streamer", {'data-test':'qsp-price'})
    RF = RF[0].text
    RF = float(RF)/100

    return RF

def calcCAPM():
    TCKR = input("What is the ticker of the stock you wish to evaluate?")
    RF = scrapeRF()
    RM = 0.11
    BETA = scrapeBeta(TCKR)
    print(f'the current required return of the market in use is {RM*100}%')
    print(f'The risk-free rate in use is {round(RF*100, 2)}%')
    print(f'The beta of {TCKR} is = {BETA}')

    CAPM = ((RM - RF) * BETA) + RF

    print(f'The expected return of {TCKR} is {round(CAPM * 100, 4)} percent')

    return



def executeScript():
    calcCAPM()
    return

executeScript()