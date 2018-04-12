import openpyxl

book = openpyxl.load_workbook('lotto.xlsx')
sheet = book.active
rows = sheet.rows

for row in rows:
    values = []
    for cell in row:
        values.append(cell.value)
    print(values)
