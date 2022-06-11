import xlwings as xw

listMAIN = ['TCKR', 'MAINPE', 'MAINPB', 'MAINPS', 'MAINQUICK', 'MAINCURRENT', 'MAINDTOE', 'MAINROA', 'MAINROE', 'MAINROI', 'MAINGM']
print(listMAIN)
wb = xw.Book('RECDCF.xlsx')
sheet = wb.sheets['DATA_COMPS']
sheet.range('A2:L12').clear_contents()
sheet.range('B2').value = listMAIN
