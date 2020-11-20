import openpyxl

book = openpyxl.load_workbook('prueba_escritura.xlsx')

sheet = book.active

sheet['E1'] = 'suma total'
sheet['E2'] = '=SUM(B2:B14)'

book.save('prueba_escritura.xlsx')
