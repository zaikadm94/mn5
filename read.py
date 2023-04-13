import openpyxl

book = openpyxl.open("5MN.xlsx")

sheet = book.worksheets[3]
for row in range(11, sheet.max_row + 1):
    for col in range(10, sheet.max_column + 1):
        cell = sheet.cell(row=row, column=col)
        print(cell.value, end='')
print()