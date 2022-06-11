import pandas as pd
import xlwings as xw

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
import COMPSAUTO

livewebsite = True

# asktckr
def asktckr3():
    if livewebsite == True:
        print("Welcome to Steve's DCF generator, this code currently works with companies that are not: banks, insurance providers, companies who do not report in USD")
        print('Input ticker in all caps for example: NVDA')
        global TCKR
        TCKR = input("What is the ticker of the stock you would like to run a valuation on?  ")


    else:
        TCKR = 'NVDA'
        print("LIVEWEBSITE = FALSE , RETURNING NVIDIA'S INFORMATION (as of February, 2022) AS TEST")

    return TCKR

# -------------------DCFWEBSCRAPE-------------------
TCKR = asktckr3()
PGA = askpga()
EM = askEM()
Y, Y1, Y2, currentdate, IShtml, = scrapedatesunits(TCKR)
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
# export to file
wb = xw.Book('RECDCF.xlsx')
sheet = wb.sheets['DATA']
sheet.range('A1').value = df
print('DCF data successfully exported into Excel File')


# -------------------COMPSAUTO-------------------
tckrdict, PElist, PBlist, PSlist, QUICKlist, CURRENTlist, DTOElist, ROAlist, ROElist, ROIlist, GMlist = COMPSAUTO.createdicts()
COMPSAUTO.scrapecomp.counter = 0
tckrdict, PElist, PBlist, PSlist, QUICKlist, CURRENTlist, DTOElist, ROAlist, ROElist, ROIlist, GMlist = COMPSAUTO.repeat_collect(
    TCKR, tckrdict, PElist, PBlist, PSlist, QUICKlist, CURRENTlist, DTOElist, ROAlist, ROElist, ROIlist, GMlist)
MAINHOME, listMAIN = COMPSAUTO.collectMAIN(TCKR)

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