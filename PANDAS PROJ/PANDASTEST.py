import pandas as pd
import xlwings as xw

from DCFWEBSCRAPE import asktckr
from DCFWEBSCRAPE import scrapedatesunits
from DCFWEBSCRAPE import scraperev
from DCFWEBSCRAPE import scrapedepna
from DCFWEBSCRAPE import scrapepretax
from DCFWEBSCRAPE import scrapetax
from DCFWEBSCRAPE import scrapeCAPEX
from DCFWEBSCRAPE import scrapeNWC
from DCFWEBSCRAPE import scrapeunits
from DCFWEBSCRAPE import wait
from DCFWEBSCRAPE import scrapeRF
from DCFWEBSCRAPE import scrapeHOME
from DCFWEBSCRAPE import scrapeCOD
from DCFWEBSCRAPE import askEM
from DCFWEBSCRAPE import askpga
from DCFWEBSCRAPE import scrapeCCEandSTLandLTL
from DCFWEBSCRAPE import calcEBIT
from DCFWEBSCRAPE import scrapeEBITDA

# DCFWEBSCRAPE
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
                                                'Tax Exp': RF, 'CAPEX': COD, 'NWC': LTL, 'STL': STL, 'Cash + Cash eq' : CCE, 'Shares outst' : SHARES, 'TERM G RATE' : PGA, 'Exit Mult' : EM, 'UNITS' : UNITS, 'NAME': NAME}})

wb = xw.Book('RECDCF.xlsx')
sheet = wb.sheets['DATA']
sheet.range('A1').value = df
print('DCF data successfully exported into Excel File')

