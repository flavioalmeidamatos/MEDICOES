import openpyxl

FILE_MODELO = r"d:\APRENDIZADO APP\MEDIÇÕES\MEDIÇÕES.xlsx"
wb = openpyxl.load_workbook(FILE_MODELO)
ws = wb['Medições']

# Procurar os textos CIVIS, CONTINGENCIA, ESPECIAIS e pegar as cores
styles = {}
for row in ws.iter_rows(min_row=2, max_row=100, min_col=1, max_col=15):
    for cell in row:
        val = str(cell.value).strip().upper() if cell.value else ""
        if val in ["CIVIS", "CONTINGENCIA", "CONTINGÊNCIA", "ESPECIAIS"] and val not in styles:
            styles[val] = {
                "fill": cell.fill.start_color.rgb if cell.fill else "N/A",
                "font_bold": cell.font.bold if cell.font else False
            }

print(styles)
