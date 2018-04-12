import openpyxl

book = openpyxl.load_workbook('lotto.xlsx')
sheet = book.active
rows = sheet.rows

data_ragne = sheet['B4':'T804']
for row in data_ragne:
    values = []
    for cell in row:
        values.append(cell.value)
    print(values)




