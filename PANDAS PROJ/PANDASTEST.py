import openpyxl
import pandas as pd
import bs4 as bs
import requests
import codecs
import lxml
import xlsxwriter
import xlwings as xw
import datetime


from DCFWEBSCRAPE import asktckr
from DCFWEBSCRAPE import scrapedates
from DCFWEBSCRAPE import scraperev
from DCFWEBSCRAPE import scrapedep
from DCFWEBSCRAPE import scrapeEBIT
from DCFWEBSCRAPE import EBITDACALC
from DCFWEBSCRAPE import scrapepretax
from DCFWEBSCRAPE import scrapetax
from DCFWEBSCRAPE import scrapeCAPEX
from DCFWEBSCRAPE import scrapeNWC
from DCFWEBSCRAPE import scrapeunits
from DCFWEBSCRAPE import imnotacrook
from DCFWEBSCRAPE import scrapeRF
from DCFWEBSCRAPE import scrapeHOME


from datetime import date

from DCFWEBSCRAPE import executescript

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



# DataFrame Creation
# the format of the dataframes is as follows: in quotations the title of the column followed by in
# brackets the {row number: corresponding data, row number: corresponding data, row number: corresponding data,

df = pd.DataFrame({ Y :{'Revenue': Rev, 'EBITDA': EBITDA, 'EBIT': EBIT,
                                             'Pretax Income': pretax, 'Tax Exp': tax, 'CAPEX': capex,
                                             'NWC': NWC},

                             Y1: {'Revenue': Rev1, 'EBITDA': EBITDA1,
                                               'EBIT': EBIT1, 'Pretax Income': pretax1,
                                               'Tax Exp': tax1, 'CAPEX': capex1,
                                               'NWC': NWC1},

                             Y2: {'Revenue': Rev2, 'EBITDA': EBITDA2, 'EBIT': EBIT2, 'Pretax Income': pretax2,
                                                'Tax Exp': tax2, 'CAPEX': capex2, 'NWC': NWC2},

                             '': {'Revenue': 'Ticker:', 'EBITDA': 'Todays date', 'EBIT': "Current Price", 'Pretax Income': "Beta",
                                                'Tax Exp': 'RF Rate', 'CAPEX': 'Cost of Debt', 'NWC': 'LTL', 'STL': 'STL', 'Cash + Cash eq' : 'Cash + Cash eq', 'Shares outst' : 'Shares outst', 'TERM G RATE' : 'TERM G RATE', 'Exit Mult' : 'Exit Mult', 'UNITS': 'UNITS', 'NAME': 'NAME'},

                             'DESC': {'Revenue': TCKR, 'EBITDA': currentdate, 'EBIT': PRICE, 'Pretax Income': BETA,
                                                'Tax Exp': RF, 'CAPEX': 0, 'NWC': 0, 'STL': 1, 'Cash + Cash eq' : 2, 'Shares outst' : 4, 'TERM G RATE' : 5, 'Exit Mult' : 6, 'UNITS' : UNITS, 'NAME': NAME}})

wb = xw.Book('RECDCF.xlsx')
sheet = wb.sheets['DATA']
sheet.range('A1').value = df

print('Data successfully exported into Excel File')
