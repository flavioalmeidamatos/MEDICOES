import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter
from datetime import datetime
import os
import re

# Caminhos dos arquivos
CWD = r"d:\APRENDIZADO APP\MEDIÇÕES"
FILE_BASE = os.path.join(CWD, "BASE.xlsx")
FILE_ANALITICA = os.path.join(CWD, "ANALITICA.xlsx")
FILE_AUXILIAR = os.path.join(CWD, "AUXILIAR.xlsx")
FILE_COMISSOES = os.path.join(CWD, "COMISSÕES POR REGIAO.xlsx")
FILE_OUTPUT = os.path.join(CWD, "MEDIÇÕES_CONSOLIDADO.xlsx")

def clean_sei(val):
    if pd.isna(val): return ""
    return str(val).strip()

def to_numeric(val):
    if pd.isna(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    # Remove R$, espaços, pontos de milhar, troca vírgula por ponto
    s = str(val).replace("R$", "").replace("\xa0", "").replace(" ", "")
    # Se houver pontos e vírgulas, assume que ponto é milhar e vírgula é decimal
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return round(float(s), 2)
    except:
        return 0.0

def get_region_mapping():
    # Lê AUXILIAR.xlsx para mapear Município -> Região
    df_aux = pd.read_excel(FILE_AUXILIAR, sheet_name="AUXILIAR")
    mapping = {}
    siglas = {
        "BAIXADA": "BX",
        "METROPOLITANA": "MT",
        "SUL FLUMINENSE": "SL",
        "NORTE": "NT"
    }
    # O arquivo AUXILIAR tem colunas com nomes das Regiões e os municípios abaixo
    for col in df_aux.columns:
        if "Unnamed" in col: continue
        reg_name = col.strip().upper()
        if reg_name in siglas:
            sigla = siglas[reg_name]
            for muni in df_aux[col].dropna():
                mapping[str(muni).strip().upper()] = sigla
    return mapping

def normalize_name(name):
    if not name or pd.isna(name): return ""
    # Remove pontos, traços, barras e espaços múltiplos para comparação
    n = str(name).upper().strip()
    n = re.sub(r'[\.\-\/]', ' ', n)
    n = re.sub(r'\s+', ' ', n).strip()
    return n

def get_contractor_mapping():
    # Lê AUXILIAR.xlsx para mapear CONTRATADA -> RESUMIDO
    df_aux = pd.read_excel(FILE_AUXILIAR, sheet_name="AUXILIAR")
    mapping = {}
    if 'CONTRATADA' in df_aux.columns and 'RESUMIDO' in df_aux.columns:
        for _, row in df_aux[['CONTRATADA', 'RESUMIDO']].dropna(subset=['CONTRATADA', 'RESUMIDO']).iterrows():
            orig = normalize_name(row['CONTRATADA'])
            res = str(row['RESUMIDO']).strip()
            if orig:
                mapping[orig] = res
    return mapping

def get_comissoes_data():
    xl = pd.ExcelFile(FILE_COMISSOES)
    data = {}
    
    # 1. Lê a aba AUXILIAR para mapear SEI -> STATUS, GESTOR e LOCAL
    aux_sheet = next((s for s in xl.sheet_names if s.upper() == "AUXILIAR"), None)
    if aux_sheet:
        df_aux = pd.read_excel(FILE_COMISSOES, sheet_name=aux_sheet)
        # Normaliza colunas
        df_aux.columns = [str(c).upper().strip() for c in df_aux.columns]
        if 'SEI' in df_aux.columns:
            for _, row in df_aux.iterrows():
                sei = clean_sei(row['SEI'])
                if not sei: continue
                data[sei] = {
                    'gestor': str(row['GESTOR']).strip() if 'GESTOR' in df_aux.columns and pd.notna(row['GESTOR']) else "",
                    'status_aux': str(row['STATUS']).replace('#', '').strip().upper() if 'STATUS' in df_aux.columns and pd.notna(row['STATUS']) else "",
                    'local': str(row['LOCAL']).strip().upper() if 'LOCAL' in df_aux.columns and pd.notna(row['LOCAL']) else "CIVIS"
                }

    # 2. Percorre as abas específicas para garantir o LOCAL correto baseado na aba
    for sheet in xl.sheet_names:
        if sheet.upper() == "AUXILIAR": continue
        
        local_val = "CIVIS"
        if sheet.upper() == "CONTIGENCIA": local_val = "CONTINGENCIA"
        elif sheet.upper() == "ESPECIAIS": local_val = "ESPECIAIS"
            
        df = pd.read_excel(FILE_COMISSOES, sheet_name=sheet, header=None)
        
        sei_idx = None
        gestor_idx = None
        start_row = 0
        
        for i in range(min(10, len(df))):
            row_vals = df.iloc[i].fillna("").astype(str).str.upper().tolist()
            s, g = None, None
            for j, val in enumerate(row_vals):
                v = val.replace("\n", " ").strip()
                if "SEI" == v or "PROCESSO SEI" == v: s = j
                if any(x in v for x in ["GESTOR", "GESTOR(A)", "GESTOR ATUANTE"]): g = j
            if s is not None and g is not None:
                sei_idx, gestor_idx, start_row = s, g, i + 1
                break
        
        if sei_idx is not None:
            for i in range(start_row, len(df)):
                row = df.iloc[i]
                sei = clean_sei(row[sei_idx])
                if not sei or sei.upper() == "NAN": continue
                
                gestor = str(row[gestor_idx]).strip() if gestor_idx is not None and pd.notna(row[gestor_idx]) else ""
                
                if sei not in data:
                    data[sei] = {'gestor': gestor, 'local': local_val, 'status_aux': ''}
                else:
                    # Se já existe (pela AUXILIAR), atualiza o LOCAL se o da aba for mais específico
                    data[sei]['local'] = local_val
                    if gestor and not data[sei]['gestor']:
                        data[sei]['gestor'] = gestor
    return data

def main():
    print("Iniciando..." )
    
    # 1. Carregar mapeamentos
    region_map = get_region_mapping()
    comissoes_map = get_comissoes_data()
    contractor_map = get_contractor_mapping()
    
    # 2. Carregar ANALITICA (Metadados) e remover duplicatas de SEI
    df_ana = pd.read_excel(FILE_ANALITICA)
    df_ana['SEI_CLEAN'] = df_ana['Processo SEI'].apply(clean_sei)
    
    # Se houver duplicatas de SEI na métrica analítica, mantemos apenas o primeiro
    df_ana = df_ana.drop_duplicates(subset=['SEI_CLEAN']).copy()
    
    # 3. Carregar BASE (Lançamentos) e pivotar
    df_base = pd.read_excel(FILE_BASE)
    df_base['SEI_CLEAN'] = df_base['Processo SEI'].apply(clean_sei)
    df_base['Valor'] = df_base['Valor das medições'].apply(to_numeric)
    
    # Mapeamento de meses para string JAN/21
    meses_pt = {1: "JAN", 2: "FEV", 3: "MAR", 4: "ABR", 5: "MAI", 6: "JUN", 7: "JUL", 8: "AGO", 9: "SET", 10: "OUT", 11: "NOV", 12: "DEZ"}
    df_base['MesAno'] = df_base.apply(lambda r: f"{meses_pt[r['Mês']]}/{str(r['Ano'])[2:]}", axis=1)
    
    # Pivotar: Index SEI, Columns MesAno, Values Valor
    df_pivot = df_base.pivot_table(index='SEI_CLEAN', columns='MesAno', values='Valor', aggfunc='sum').fillna(0)
    
    # 4. Consolidar dados
    final_rows = []
    gestores_faltantes = []
    
    # Ordenar colunas de meses (JAN/21 a DEZ/25)
    all_months = []
    for ano in range(21, 26):
        for mes in range(1, 13):
            all_months.append(f"{meses_pt[mes]}/{str(ano)}")
    
    # Nomes exatos com quebras de linha conforme o modelo
    h_prazo = "PRAZO\nEXECUÇÃO"
    h_inicio = "ORDEM\nDE INÍCIO"
    h_vlr_contr = "VLR.CONTRATO\nC/ADITIVO"
    h_med_acum = "MEDIÇÕES\nACUMULADAS"
    h_saldo = "SALDO DO\nCONTRATO"

    for _, row in df_ana.iterrows():
        sei = row['SEI_CLEAN']
        
        info = comissoes_map.get(sei, {'gestor': '', 'local': 'CIVIS', 'status_aux': ''})
        gestor = info['gestor']
        local = info['local']
        status_final = info.get('status_aux') if info.get('status_aux') else str(row['Fase']).strip().upper()
        
        if not gestor:
            gestores_faltantes.append({'SEI': row['Processo SEI'], 'CONTRATADA': row['Contratada']})
            
        muni = str(row['Municipio']).strip().upper()
        regiao = region_map.get(muni, "")
        
        ordem_inicio = pd.to_datetime(row['Ordem de Início'], errors='coerce')
        data_final = pd.to_datetime(row['Prazo Final'], errors='coerce')
        prazo = (data_final - ordem_inicio).days if pd.notnull(data_final) and pd.notnull(ordem_inicio) else ""
        
        vlr_contrato = to_numeric(row['Valor contrato (Atual)'])
        
        med_months = {}
        for m in all_months:
            val = df_pivot.loc[sei, m] if sei in df_pivot.index and m in df_pivot.columns else 0.0
            med_months[m] = round(val, 2)
            
        med_2025 = sum(med_months.get(f"{meses_pt[m]}/25", 0.0) for m in range(1, 13))
        med_acumulada = sum(med_months.values())
        
        perc_exec = med_acumulada / vlr_contrato if vlr_contrato > 0 else 0.0
        saldo = vlr_contrato - med_acumulada
        # Contratada (Substituir pelo Resumido se houver match)
        fullname = str(row['Contratada']).strip()
        norm_name = normalize_name(fullname)
        contratada_final = contractor_map.get(norm_name, fullname)
        
        # Montar a linha final conforme layout
        linha = {
            "SEI": row['Processo SEI'],
            "LOCAL": local,
            h_prazo: prazo,
            h_inicio: ordem_inicio,
            "DATA FINAL": data_final,
            h_vlr_contr: vlr_contrato,
            "STATUS": status_final,
            "GESTOR": gestor,
            "REGIÃO": regiao,
            "MUNICIPIO": row['Municipio'],
            "CONTRATADA": contratada_final,
            "MEDIÇÕES 2025": round(med_2025, 2),
            h_med_acum: round(med_acumulada, 2),
            "% EXEC.": perc_exec,
            h_saldo: round(saldo, 2)
        }
        linha.update(med_months)
        final_rows.append(linha)
        
    df_final = pd.DataFrame(final_rows)
    
    # 1. Filtrar somente registros em "EXECUÇÃO"
    df_final = df_final[df_final['STATUS'] == 'EXECUÇÃO'].copy()
    
    # 2. Remover duplicatas residuais caso existam por outras joins
    df_final = df_final.drop_duplicates(subset=['SEI']).copy()
    
    # 3. Definir ordem customizada para LOCAL: CIVIS -> CONTINGENCIA -> ESPECIAIS
    local_rank = {"CIVIS": 0, "CONTINGENCIA": 1, "ESPECIAIS": 2}
    df_final['_rank_local'] = df_final['LOCAL'].map(lambda x: local_rank.get(str(x).upper().strip(), 99))
    
    # 4. Ordenar primeiro pelo Rank do LOCAL (Agrupar Civis, depois Conting., depois Especiais), e depois por CONTRATADA
    df_final = df_final.sort_values(by=["_rank_local", "CONTRATADA"], ascending=[True, True]).reset_index(drop=True)
    df_final = df_final.drop(columns=['_rank_local'])
    
    # 5. Numerar sequencialmente após todas as filtragens e ordenações
    if "Nº" in df_final.columns: df_final = df_final.drop(columns=["Nº"])
    df_final.insert(0, "Nº", range(1, len(df_final) + 1))
    
    with pd.ExcelWriter(FILE_OUTPUT, engine='openpyxl') as writer:
        df_final.to_excel(writer, sheet_name='Medições', index=False)
        if gestores_faltantes:
            pd.DataFrame(gestores_faltantes).to_excel(writer, sheet_name='GESTOR_FALTANTES', index=False)
            
    wb = openpyxl.load_workbook(FILE_OUTPUT)
    ws = wb['Medições']
    
    from openpyxl.styles import Alignment, Border, Side
    
    # Border style
    thin_border = Border(
        left=Side(style='thin'), 
        right=Side(style='thin'), 
        top=Side(style='thin'), 
        bottom=Side(style='thin')
    )
    
    # Fill colors
    fill_header = PatternFill(start_color="E6E6E6", end_color="E6E6E6", fill_type="solid")
    
    # Colors for LOCAL (vibrantes conforme imagem)
    fills_local = {
        "CIVIS": PatternFill(start_color="F4B084", end_color="F4B084", fill_type="solid"),        # Laranja
        "CONTINGENCIA": PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid"),  # Amarelo
        "ESPECIAIS": PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")     # Verde Água
    }

    fills_regiao = {
        "SL": PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
        "NT": PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid"),
        "BX": PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"),
        "MT": PatternFill(start_color="00B0F0", end_color="00B0F0", fill_type="solid")
    }
    
    money_fmt = '_-R$ * #.##0,00_-;_-R$ * -#.##0,00_-;_-R$ * "-"??_-;_-@_-'
    date_fmt = 'DD/MM/YYYY'
    perc_fmt = '0.00%'
    
    header = [cell.value for cell in ws[1]]
    col_map = {name: i+1 for i, name in enumerate(header)}
    
    # --- LEITURA DE ESTILOS E LARGURAS DO MODELO ---
    model_widths = {}
    model_header_style = {}
    try:
        wb_mod = openpyxl.load_workbook(FILE_BASE.replace("BASE.xlsx", "MEDIÇÕES.xlsx"), data_only=False)
        ws_mod = wb_mod['Medições']
        # No modelo as colunas começam na linha 2 (visto em testes anteriores)
        for i in range(1, ws_mod.max_column + 1):
            col_let = get_column_letter(i)
            w = ws_mod.column_dimensions[col_let].width
            cell_mod = ws_mod.cell(row=2, column=i)
            val_mod = cell_mod.value
            if val_mod:
                name_clean = str(val_mod).replace('\n', ' ').strip()
                model_widths[name_clean] = w
                model_header_style[name_clean] = {
                    'fill': cell_mod.fill.start_color.rgb if cell_mod.fill else None,
                    'font_bold': cell_mod.font.bold if cell_mod.font else False,
                    'font_color': cell_mod.font.color.rgb if cell_mod.font and cell_mod.font.color else None
                }
    except Exception as e:
        print(f"Aviso: Não foi possível ler larguras do modelo: {e}")

    # Header format
    for cell in ws[1]:
        name_in_cell = str(cell.value).replace('\n', ' ').strip()
        
        # Estilo base
        cell.alignment = Alignment(wrapText=True, horizontal='center', vertical='center')
        cell.border = thin_border
        
        # Tenta aplicar do modelo
        if name_in_cell in model_header_style:
            style = model_header_style[name_in_cell]
            if style['fill'] and style['fill'] != '00000000':
                cell.fill = PatternFill(start_color=style['fill'], end_color=style['fill'], fill_type="solid")
            else:
                cell.fill = fill_header # Fallback cinza
            cell.font = Font(bold=style['font_bold'], color=style['font_color'])
        else:
            cell.fill = fill_header
            cell.font = Font(bold=True)

    # Data content
    for row in range(2, ws.max_row + 1):
        for col in range(1, len(header) + 1):
            ws.cell(row=row, column=col).border = thin_border

        local_cell = ws.cell(row=row, column=col_map["LOCAL"])
        local_val = str(local_cell.value).strip().upper()
        if local_val in fills_local:
            local_cell.fill = fills_local[local_val]
            local_cell.font = Font(bold=True)

        reg_val = ws.cell(row=row, column=col_map["REGIÃO"]).value
        if reg_val in fills_regiao:
            ws.cell(row=row, column=col_map["REGIÃO"]).fill = fills_regiao[reg_val]
            
        money_fmt = '_-R$ * #,##0.00_-;_-R$ * -#,##0.00_-;_-R$ * "-"??_-;_-@_-'
        money_cols = [h_vlr_contr, "MEDIÇÕES 2025", h_med_acum, h_saldo] + all_months
        for mc in money_cols:
            cell = ws.cell(row=row, column=col_map[mc])
            cell.number_format = money_fmt
            try:
                if cell.value is not None:
                    val_clean = str(cell.value).replace('R$', '').replace(' ', '')
                    if ',' in val_clean and '.' not in val_clean: val_clean = val_clean.replace(',', '.')
                    cell.value = round(float(val_clean), 2)
                else: cell.value = 0.0
            except: cell.value = 0.0
            
        for dc in [h_inicio, "DATA FINAL"]:
            cell = ws.cell(row=row, column=col_map[dc])
            if cell.value: cell.number_format = 'DD/MM/YYYY'
        ws.cell(row=row, column=col_map["% EXEC."]).number_format = '0.00%'

    # --- Aplicação de Larguras ---
    for col in ws.columns:
        column_letter = col[0].column_letter
        val_header = str(col[0].value).replace('\n', ' ').strip()
        
        if val_header in model_widths and model_widths[val_header] is not None:
            ws.column_dimensions[column_letter].width = model_widths[val_header]
        else:
            # Fallback autofit simples se não achar no modelo
            max_len = max(len(str(cell.value)) for cell in col)
            ws.column_dimensions[column_letter].width = min(max_len + 5, 40)

    wb.save(FILE_OUTPUT)
    print(f"Finalizado: {FILE_OUTPUT}")

if __name__ == "__main__":
    main()
