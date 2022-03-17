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

from datetime import date

from DCFWEBSCRAPE import executescript

TCKR = asktckr()
Y, Y1, Y2, currentdate, IShtml = scrapedates(TCKR)
capex, capex1, capex2, CFShtml = scrapeCAPEX(TCKR)
Rev, Rev1, Rev2, IShtml = scraperev(IShtml)
EBIT, EBIT1, EBIT2, IShtml = scrapeEBIT(IShtml)
dep, dep1, dep2, IShtml = scrapedep(IShtml)
EBITDA, EBITDA1, EBITDA2 = EBITDACALC(dep, dep1, dep2, EBIT, EBIT1, EBIT2)
pretax, pretax1, pretax2, IShtml = scrapepretax(IShtml)
tax, tax1, tax2, IShtml = scrapetax(IShtml)


# DataFrame Creation
# the format of the dataframes is as follows: in quotations the title of the column followed by in
# brackets the {row number: corresponding data, row number: corresponding data, row number: corresponding data,

df = pd.DataFrame({ Y :{'Revenue': Rev, 'EBITDA': EBITDA, 'EBIT': EBIT,
                                             'Pretax Income': pretax, 'Tax Exp': tax, 'CAPEX': capex,
                                             'NWC': 107},

                             Y1: {'Revenue': Rev1, 'EBITDA': EBITDA1,
                                               'EBIT': EBIT1, 'Pretax Income': pretax1,
                                               'Tax Exp': tax1, 'CAPEX': capex1,
                                               'NWC': 'Android Box'},

                             Y2: {'Revenue': Rev2, 'EBITDA': EBITDA2, 'EBIT': EBIT2, 'Pretax Income': pretax2,
                                                'Tax Exp': tax2, 'CAPEX': capex2, 'NWC': 1800},

                             '': {'Revenue': 'Ticker:', 'EBITDA': 'Todays date', 'EBIT': "Current Price", 'Pretax Income': "Beta",
                                                'Tax Exp': 'RF Rate', 'CAPEX': 'Cost of Debt', 'NWC': 'LTL', 'STL': 'STL', 'Cash' : 'Cash', 'Cash Eq' : 'Cash Eq', 'Shares outst' : 'Shares outst', 'TERM G RATE' : 'TERM G RATE', 'Exit Mult' : 'Exit Mult'},

                             'DESC': {'Revenue': TCKR, 'EBITDA': currentdate, 'EBIT': 200, 'Pretax Income': 0,
                                                'Tax Exp': 0, 'CAPEX': 0, 'NWC': 0, 'STL': 1, 'Cash' : 2, 'Cash Eq' : 3, 'Shares outst' : 4, 'TERM G RATE' : 5, 'Exit Mult' : 6}})

wb = xw.Book('RECDCF.xlsx')
sheet = wb.sheets['DATA']
sheet.range('A1').value = df

print('Data successfully exported into Excel File')
