import openpyxl

FILE_MODELO = r"d:\APRENDIZADO APP\MEDIÇÕES\MEDIÇÕES.xlsx"
wb = openpyxl.load_workbook(FILE_MODELO)
ws = wb['Medições']

# Check first few cells of first row to see patterns
for i in range(1, 10):
    cell = ws.cell(row=1, column=i)
    print(f"Col {i} ({cell.value}): Fill={cell.fill.start_color.rgb}, Font={cell.font.color.rgb if cell.font.color else 'None'}, Bold={cell.font.bold}")
