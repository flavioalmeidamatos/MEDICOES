import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import os
import re

# Caminhos dos arquivos
CWD = r"d:\APRENDIZADO APP\MEDICOES"
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

def apply_sheet_formatting(ws, col_map, header, all_months, model_widths, model_header_style,
                           h_vlr_contr, h_med_acum, h_saldo, h_inicio):
    """Aplica formatação idêntica (cabeçalhos, cores, bordas, larguras) a uma worksheet."""
    
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
                cell.fill = fill_header  # Fallback cinza
            cell.font = Font(bold=style['font_bold'], color=style['font_color'])
        else:
            cell.fill = fill_header
            cell.font = Font(bold=True)

    # Data content
    money_fmt = '_-R$ * #,##0.00_-;_-R$ * -#,##0.00_-;_-R$ * "-"??_-;_-@_-'
    # Utilizar as colunas que foram passadas como "all_months" ou colunas financeiras
    # Mas aqui precisamos saber quais sao financeiras.
    # Vamos inferir: tudo que é mês, valor, saldo.
    
    for row in range(2, ws.max_row + 1):
        for col in range(1, len(header) + 1):
            ws.cell(row=row, column=col).border = thin_border

        # Formatação LOCAL
        if "LOCAL" in col_map:
            local_cell = ws.cell(row=row, column=col_map["LOCAL"])
            local_val = str(local_cell.value).strip().upper()
            if local_val in fills_local:
                local_cell.fill = fills_local[local_val]
                local_cell.font = Font(bold=True)

        # Formatação REGIÃO
        if "REGIÃO" in col_map:
            reg_val = ws.cell(row=row, column=col_map["REGIÃO"]).value
            if reg_val in fills_regiao:
                ws.cell(row=row, column=col_map["REGIÃO"]).fill = fills_regiao[reg_val]

        # Formatação Financeira (heurística + lista passada)
        # Se o nome da coluna estiver em all_months ou contiver VLR, SALDO, MEDIÇÕES (exceto ano solto)
        for col_name, col_idx in col_map.items():
            is_money = False
            # Verifica meses
            c_clean = col_name.replace(" ", "")
            if re.match(r'^[A-Z]{3}/\d{2}$', c_clean):
                is_money = True
            elif any(k in col_name.upper() for k in ["VLR", "VALOR", "SALDO", "MEDIÇÕES", "MEDICOES"]):
                # Evita "MEDIÇÕES 2025" se for só contagem, mas aqui é valor, então ok
                is_money = True
            
            if is_money:
                cell = ws.cell(row=row, column=col_idx)
                cell.number_format = money_fmt
                try:
                    if cell.value is not None:
                        val_clean = str(cell.value).replace('R$', '').replace(' ', '')
                        if ',' in val_clean and '.' not in val_clean:
                            val_clean = val_clean.replace(',', '.')
                        cell.value = round(float(val_clean), 2)
                    else:
                        cell.value = 0.0
                except:
                    cell.value = 0.0

        # Formatação de Datas
        for dc in [h_inicio, "DATA FINAL", "Prazo Final", "Ordem de Início"]:
            if dc in col_map:
                cell = ws.cell(row=row, column=col_map[dc])
                if cell.value:
                    cell.number_format = 'DD/MM/YYYY'
        
        if "% EXEC." in col_map:
            ws.cell(row=row, column=col_map["% EXEC."]).number_format = '0.00%'

    # --- Aplicação de Larguras ---
    for col in ws.columns:
        column_letter = col[0].column_letter
        val_header = str(col[0].value).replace('\n', ' ').strip()

        if val_header in model_widths and model_widths[val_header] is not None:
            ws.column_dimensions[column_letter].width = model_widths[val_header]
        else:
            # Fallback autofit simples
            max_len = max(len(str(cell.value)) for cell in col)
            ws.column_dimensions[column_letter].width = min(max_len + 5, 40)


def prepare_dataframe(df, status_filter='EXECUÇÃO', keep_execution=True):
    """Filtra, ordena e numera o DataFrame conforme o status desejado."""
    if keep_execution:
        df_filtered = df[df['STATUS'] == status_filter].copy()
    else:
        df_filtered = df[df['STATUS'] != status_filter].copy()

    # Remover duplicatas residuais
    df_filtered = df_filtered.drop_duplicates(subset=['SEI']).copy()

    # Definir ordem customizada para LOCAL
    local_rank = {"CIVIS": 0, "CONTINGENCIA": 1, "ESPECIAIS": 2}
    df_filtered['_rank_local'] = df_filtered['LOCAL'].map(lambda x: local_rank.get(str(x).upper().strip(), 99))

    # Ordenar
    df_filtered = df_filtered.sort_values(by=["_rank_local", "CONTRATADA"], ascending=[True, True]).reset_index(drop=True)
    df_filtered = df_filtered.drop(columns=['_rank_local'])

    # Numerar sequencialmente
    if "Nº" in df_filtered.columns:
        df_filtered = df_filtered.drop(columns=["Nº"])
    df_filtered.insert(0, "Nº", range(1, len(df_filtered) + 1))

    return df_filtered

def get_model_structure():
    """Lê o arquivo modelo MEDIÇÕES.xlsx para obter a ordem exata das colunas e estilos."""
    model_path = FILE_BASE.replace("BASE.xlsx", "MEDIÇÕES.xlsx")
    model_widths = {}
    model_header_style = {}
    ordered_columns = []
    
    try:
        wb_mod = openpyxl.load_workbook(model_path, data_only=False)
        ws_mod = wb_mod['Medições']
        
        # Ler cabeçalhos da linha 2
        for i in range(1, ws_mod.max_column + 1):
            cell_mod = ws_mod.cell(row=2, column=i)
            val_mod = cell_mod.value
            if val_mod:
                name_clean = str(val_mod).replace('\n', ' ').strip()
                ordered_columns.append(name_clean)
                
                col_let = get_column_letter(i)
                w = ws_mod.column_dimensions[col_let].width
                model_widths[name_clean] = w
                
                model_header_style[name_clean] = {
                    'fill': cell_mod.fill.start_color.rgb if cell_mod.fill else None,
                    'font_bold': cell_mod.font.bold if cell_mod.font else False,
                    'font_color': cell_mod.font.color.rgb if cell_mod.font and cell_mod.font.color else None
                }
    except Exception as e:
        print(f"Erro ao ler modelo: {e}")
        return [], {}, {}
        
    return ordered_columns, model_widths, model_header_style

def main():
    print("Iniciando...")
    
    # 1. Obter estrutura do modelo
    ordered_columns, model_widths, model_header_style = get_model_structure()
    
    if not ordered_columns:
        print("ALERTA: Não foi possível ler colunas do modelo. Usando fallback.")
        return 

    # 2. Carregar mapeamentos
    region_map = get_region_mapping()
    comissoes_map = get_comissoes_data()
    contractor_map = get_contractor_mapping()

    # 3. Carregar DADOS
    df_ana = pd.read_excel(FILE_ANALITICA)
    df_ana['SEI_CLEAN'] = df_ana['Processo SEI'].apply(clean_sei)
    df_ana = df_ana.drop_duplicates(subset=['SEI_CLEAN']).copy()

    df_base = pd.read_excel(FILE_BASE)
    df_base['SEI_CLEAN'] = df_base['Processo SEI'].apply(clean_sei)
    val_col = 'Valor' if 'Valor' in df_base.columns else 'Valor das medições'
    if val_col not in df_base.columns:
         # Fallback cleaning if needed or error
         print(f"Erro: Coluna de valor não encontrada. Colunas: {df_base.columns}")
         return

    df_base['Valor'] = df_base[val_col].apply(to_numeric)

    meses_pt = {1: "JAN", 2: "FEV", 3: "MAR", 4: "ABR", 5: "MAI", 6: "JUN", 7: "JUL", 8: "AGO", 9: "SET", 10: "OUT", 11: "NOV", 12: "DEZ"}
    mes_map_str = {
        "JANEIRO": "JAN", "FEVEREIRO": "FEV", "MARÇO": "MAR", "ABRIL": "ABR", "MAIO": "MAI", "JUNHO": "JUN",
        "JULHO": "JUL", "AGOSTO": "AGO", "SETEMBRO": "SET", "OUTUBRO": "OUT", "NOVEMBRO": "NOV", "DEZEMBRO": "DEZ"
    }

    def get_mes_abrev(mes_val):
        s = str(mes_val).strip().upper()
        if s.isdigit():
            return meses_pt.get(int(s), "???")
        return mes_map_str.get(s, "???")

    df_base['MesAno'] = df_base.apply(lambda r: f"{get_mes_abrev(r['Mês'])}/{str(r['Ano'])[2:]}", axis=1)

    df_pivot = df_base.pivot_table(index='SEI_CLEAN', columns='MesAno', values='Valor', aggfunc='sum').fillna(0)

    # 4. Consolidar dados
    final_rows = []
    gestores_faltantes = []
    
    for _, row in df_ana.iterrows():
        sei = row['SEI_CLEAN']
        info = comissoes_map.get(sei, {'gestor': '', 'local': 'CIVIS', 'status_aux': ''})
        
        # Dados básicos
        dados = {
            "SEI": row['Processo SEI'],
            "LOCAL": info['local'],
            "STATUS": info.get('status_aux') if info.get('status_aux') else str(row['Fase']).strip().upper(),
            "GESTOR": info['gestor'],
            "MUNICIPIO": row['Municipio'],
            "REGIÃO": region_map.get(str(row['Municipio']).strip().upper(), ""),
            "CONTRATADA": contractor_map.get(normalize_name(str(row['Contratada'])), str(row['Contratada']).strip())
        }

        if not dados['GESTOR']:
            gestores_faltantes.append({'SEI': dados['SEI'], 'CONTRATADA': row['Contratada']})

        # Datas e Prazos
        dt_ini = pd.to_datetime(row['Ordem de Início'], errors='coerce')
        dt_fim = pd.to_datetime(row['Prazo Final'], errors='coerce')
        dados["ORDEM DE INÍCIO"] = dt_ini
        dados["DATA FINAL"] = dt_fim
        dados["PRAZO EXECUÇÃO"] = (dt_fim - dt_ini).days if pd.notnull(dt_fim) and pd.notnull(dt_ini) else ""

        # Financeiro
        vlr_contr = to_numeric(row['Valor contrato (Atual)'])
        dados["VLR.CONTRATO C/ADITIVO"] = vlr_contr

        med_acumulada = 0.0
        med_2025 = 0.0
        
        # Preencher meses baseado nas colunas do modelo
        for col_name in ordered_columns:
            # Tenta casar com padrão de mês
            col_clean = col_name.replace(" ", "")
            
            # Se a coluna existe no pivot, puxa o valor
            if col_clean in df_pivot.columns:
                val = df_pivot.loc[sei, col_clean] if sei in df_pivot.index else 0.0
                dados[col_name] = round(val, 2)
                med_acumulada += val
                if "/25" in col_clean:
                    # Verifica se é coluna de mês para somar no MEDIÇÕES 2025
                    if re.match(r'^[A-Z]{3}/\d{2}$', col_clean):
                        med_2025 += val
            elif col_name not in dados:
                 # Coluna do modelo que não é dado básico e não está no pivot (ex: mês futuro ou coluna calculada)
                 # Se for mês futuro (ex: JAN/26) e não tem no pivot, é 0.
                 pass

        dados["MEDIÇÕES ACUMULADAS"] = round(med_acumulada, 2)
        dados["MEDIÇÕES 2025"] = round(med_2025, 2)
        dados["SALDO DO CONTRATO"] = round(vlr_contr - med_acumulada, 2)
        dados["% EXEC."] = (med_acumulada / vlr_contr) if vlr_contr > 0 else 0.0

        # Montar linha final ordenada
        linha_ordenada = {}
        for col in ordered_columns:
            val = dados.get(col)
            
            # Fallback de Nomes
            if val is None:
                if "PRAZO" in col and "EXECUÇÃO" in col: val = dados.get("PRAZO EXECUÇÃO")
                elif "ORDEM" in col and "INÍCIO" in col: val = dados.get("ORDEM DE INÍCIO")
                elif "VLR" in col and "CONTRATO" in col: val = dados.get("VLR.CONTRATO C/ADITIVO")
                elif "MEDIÇÕES" in col and "ACUMULADAS" in col: val = dados.get("MEDIÇÕES ACUMULADAS")
                elif "SALDO" in col and "CONTRATO" in col: val = dados.get("SALDO DO CONTRATO")
                elif "%" in col and "EXEC" in col: val = dados.get("% EXEC.")
            
            linha_ordenada[col] = val if val is not None else ""
            
        final_rows.append(linha_ordenada)

    df_all = pd.DataFrame(final_rows)
    # Garante a ordem das colunas
    df_all = df_all[ordered_columns]

    # Separar em EXECUÇÃO e PROBLEMAS
    df_execucao = prepare_dataframe(df_all, status_filter='EXECUÇÃO', keep_execution=True)
    df_problemas = prepare_dataframe(df_all, status_filter='EXECUÇÃO', keep_execution=False)

    # Escrever
    with pd.ExcelWriter(FILE_OUTPUT, engine='openpyxl') as writer:
        df_execucao.to_excel(writer, sheet_name='Medições', index=False)
        if not df_problemas.empty:
            df_problemas.to_excel(writer, sheet_name='PROBLEMAS', index=False)
        if gestores_faltantes:
            pd.DataFrame(gestores_faltantes).to_excel(writer, sheet_name='GESTOR_FALTANTES', index=False)

    # Formatar
    wb = openpyxl.load_workbook(FILE_OUTPUT)
    
    for sheet_name in ['Medições', 'PROBLEMAS']:
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            apply_sheet_formatting(ws, 
                                   col_map={c: i+1 for i, c in enumerate(ordered_columns)}, 
                                   header=ordered_columns, 
                                   all_months=[], # Não usado mais na nova lógica de formatação interna
                                   model_widths=model_widths, 
                                   model_header_style=model_header_style,
                                   h_vlr_contr="VLR.CONTRATO C/ADITIVO", 
                                   h_med_acum="MEDIÇÕES ACUMULADAS", 
                                   h_saldo="SALDO DO CONTRATO", 
                                   h_inicio="ORDEM DE INÍCIO")

    wb.save(FILE_OUTPUT)
    print(f"Finalizado: {FILE_OUTPUT}")
    print(f"  - Aba 'Medições': {len(df_execucao)} obras em EXECUÇÃO")
    print(f"  - Aba 'PROBLEMAS': {len(df_problemas)} obras com status != EXECUÇÃO")
    if gestores_faltantes:
        print(f"  - Aba 'GESTOR_FALTANTES': {len(gestores_faltantes)} registros sem gestor")

if __name__ == "__main__":
    main()
