import openpyxl

FILE_MODELO = r"d:\APRENDIZADO APP\MEDIÇÕES\MEDIÇÕES.xlsx"
wb = openpyxl.load_workbook(FILE_MODELO)
ws = wb['Medições']

# Procurar ESPECIAIS
for row in ws.iter_rows(min_row=2, max_row=200, min_col=1, max_col=5):
    for cell in row:
        val = str(cell.value).strip().upper() if cell.value else ""
        if "ESPECIAIS" in val:
            print(f"ESPECIAIS Found: Fill={cell.fill.start_color.rgb}")
            break
