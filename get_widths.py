import openpyxl

FILE_MODELO = r"d:\APRENDIZADO APP\MEDIÇÕES\MEDIÇÕES.xlsx"
wb = openpyxl.load_workbook(FILE_MODELO)
ws = wb['Medições']

widths = {}
for i in range(1, ws.max_column + 1):
    col_letter = openpyxl.utils.get_column_letter(i)
    width = ws.column_dimensions[col_letter].width
    header_val = ws.cell(row=2, column=i).value # Cabeçalho está na linha 2 conforme testes anteriores
    if header_val:
        widths[header_val.replace('\n', ' ')] = width

print(widths)
