import openpyxl

FILE_MODELO = r"d:\APRENDIZADO APP\MEDIÇÕES\MEDIÇÕES.xlsx"
wb = openpyxl.load_workbook(FILE_MODELO)
ws = wb['Medições']

cell = ws['A1']
print(f"CELL_A1_VALUE: {cell.value}")
print(f"FILL_RGB: {cell.fill.start_color.rgb}")
print(f"FONT_RGB: {cell.font.color.rgb if cell.font.color and hasattr(cell.font.color, 'rgb') else 'N/A'}")
print(f"BOLD: {cell.font.bold}")
