import openpyxl

FILE_MODELO = r"d:\APRENDIZADO APP\MEDIÇÕES\MEDIÇÕES.xlsx"
wb = openpyxl.load_workbook(FILE_MODELO)
ws = wb['Medições']

for row_idx in range(1, 4):
    for col_idx in range(1, 4):
        cell = ws.cell(row=row_idx, column=col_idx)
        print(f"R{row_idx}C{col_idx}: Val={cell.value}, Fill={cell.fill.start_color.rgb}, FontBold={cell.font.bold}")
