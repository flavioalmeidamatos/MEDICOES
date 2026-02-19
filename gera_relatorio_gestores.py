
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# Define paths
INPUT_FILE = r"d:\APRENDIZADO APP\MEDICOES\CONTROLES POR COMISSÃO E GESTORES.xlsx"
OUTPUT_FILE = r"d:\APRENDIZADO APP\MEDICOES\RELATORIO DE OBRAS POR GESTORES E FISCAIS.xlsx"

def load_data(file_path):
    print(f"Reading {file_path}...")
    # Load the entire sheet to parse manually due to multi-header structure
    try:
        df_raw = pd.read_excel(file_path, header=None)
    except Exception as e:
        print(f"Error reading Excel: {e}")
        return None

    data_rows = []
    current_region = None
    
    # Iterate rows to find blocks
    # Structure:
    # Row N: Region Name (BAIXADA, etc)
    # Row N+1: Headers (SEI, GESTOR, etc)
    # Row N+M: Data...
    
    # We look for the "SEI" header to identify start of data
    header_row_idx = -1
    cols_map = {}
    
    # Standard headers we expect
    expected_headers = ['SEI', 'GESTOR(A) ATUANTE', 'FISCAL NOMEADO', 'MUNICIPIO', 'EMPRESA', '%EXEC', 'STATUS']
    
    i = 0
    while i < len(df_raw):
        row = df_raw.iloc[i].astype(str).tolist()
        # Clean row values: handle newlines and strip
        row_clean = [str(v).replace('\n', ' ').strip().upper() for v in row]
        
        # Check if this is a header row
        # Relaxed check: Look for "SEI" and "GESTOR" part
        has_sei = "SEI" in row_clean
        has_gestor = any("GESTOR" in c for c in row_clean)
        
        if has_sei and has_gestor:
            # Found a header row
            print(f"DEBUG: Found header at row {i}")
            # The region should be in the row above (i-1) in column 0 or 1
            if i > 0:
                possible_region = str(df_raw.iloc[i-1, 0]).strip()
                if possible_region.lower() == 'nan' or not possible_region:
                     possible_region = str(df_raw.iloc[i-1, 1]).strip()
                
                # If valid region text, use it. Otherwise keep previous or default.
                if possible_region.upper() not in ['NAN', '', 'TOTAL']:
                     current_region = possible_region.upper()
            
            # Map columns
            cols_map = {}
            for idx, val in enumerate(row_clean):
                 if val:
                    cols_map[val] = idx
            
            print(f"DEBUG: Columns found: {list(cols_map.keys())}")

            # Process data lines until empty line or "Total"
            j = i + 1
            while j < len(df_raw):
                d_row = df_raw.iloc[j]
                d_row_str = [str(v).strip() for v in d_row]
                
                # Check for end of block
                first_val = d_row_str[0].upper() if len(d_row_str)>0 else ""
                second_val = d_row_str[1].upper() if len(d_row_str)>1 else ""
                
                if not first_val and 'TOTAL' in second_val:
                    break # End of block
                if first_val in ['BAIXADA', 'SUL', 'NORTE', 'METROPOLITANA', 'CENTRO'] and j > i+1: 
                    # If we encounter a region name in col 0, it's likely a header for next block
                    break 
                
                # Check if it has data (SEI usually)
                if 'SEI' in cols_map:
                    sei_idx = cols_map['SEI']
                    sei_val = str(d_row[sei_idx]).strip()
                    
                    if sei_val and sei_val.lower() != 'nan' and sei_val.upper() != 'SEI' and "TOTAL" not in sei_val.upper():
                        # Extract values
                        record = {'REGIÃO': current_region}
                        for head, col_idx in cols_map.items():
                            val = d_row[col_idx]
                            record[head] = val
                        data_rows.append(record)
                
                j += 1
            i = j # Move outer loop
        else:
            i += 1
            
    return pd.DataFrame(data_rows)

def generate_report(df):
    if df is None or df.empty:
        print("No data found to generate report.")
        return

    # Filter/Clean Data
    # Ensure we have relevant fields
    # Standardize names
    df['GESTOR(A) ATUANTE'] = df['GESTOR(A) ATUANTE'].fillna('NÃO DEFINIDO').astype(str).str.strip().str.upper()
    df['FISCAL NOMEADO'] = df['FISCAL NOMEADO'].fillna('NÃO DEFINIDO').astype(str).str.strip().str.upper()
    df['REGIÃO'] = df['REGIÃO'].fillna('SEM REGIÃO').astype(str).str.strip().str.upper()

    # Expand rows where there are multiple fiscals separated by "/"
    # distinct rows for each fiscal
    df_expanded = df.assign(FISCAL_INDIVIDUAL=df['FISCAL NOMEADO'].str.split('/')).explode('FISCAL_INDIVIDUAL')
    df_expanded['FISCAL_INDIVIDUAL'] = df_expanded['FISCAL_INDIVIDUAL'].str.strip()

    # Sort: Region -> Fiscal -> Municipio
    df_sorted = df_expanded.sort_values(by=['REGIÃO', 'FISCAL_INDIVIDUAL', 'MUNICIPIO'])
    
    # Columns to show in report
    # Based on PDF "Report of Works by Managers and Fiscals", likely we want:
    # Region | Fiscal | SEI | Municipio | Empresa | Status | % Exec | Gestor
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Relatório Obras Gestores Fiscais"
    
    # Styles
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid") # Dark Blue
    
    region_font = Font(bold=True, size=12, color="000000")
    region_fill = PatternFill(start_color="BFBFBF", end_color="BFBFBF", fill_type="solid") # Grey
    
    fiscal_font = Font(bold=True, italic=True)
    fiscal_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid") # Light Blue
    
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    thick_border = Border(left=Side(style='medium'), right=Side(style='medium'), top=Side(style='medium'), bottom=Side(style='medium'))
    
    # Columns Layout
    cols = ['SEI', 'MUNICIPIO', 'EMPRESA', 'STATUS', '%EXEC', 'GESTOR(A) ATUANTE']
    start_row = 1
    
    # Main Title
    # Merge cells for title 
    title_range = ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(cols))
    title_cell = ws.cell(row=1, column=1, value="RELATÓRIO DE OBRAS POR GESTORES E FISCAIS")
    title_cell.font = Font(bold=True, size=14)
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Apply border to merged title range
    # Ideally apply to all cells in range, but usually top-left is enough for content, borders need loop
    for r in range(1, 2):
        for c in range(1, len(cols) + 1):
             ws.cell(row=r, column=c).border = thick_border
    
    current_row = 3
    
    # Loop through groups: Region -> Fiscal
    for region, group_r in df_sorted.groupby('REGIÃO'):
        # Region Header
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=len(cols))
        cell_r = ws.cell(row=current_row, column=1, value=f"REGIÃO: {region}")
        cell_r.font = region_font
        cell_r.fill = region_fill
        cell_r.alignment = Alignment(horizontal='left')
        for c in range(1, len(cols) + 1):
             ws.cell(row=current_row, column=c).border = thin_border
        current_row += 1
        
        for fiscal, group_f in group_r.groupby('FISCAL_INDIVIDUAL'):
            if not fiscal: continue
            
            # Fiscal Sub-header
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=len(cols))
            cell_f = ws.cell(row=current_row, column=1, value=f"  FISCAL: {fiscal}")
            cell_f.font = fiscal_font
            cell_f.fill = fiscal_fill
             # Apply border to fiscal header row
            for c in range(1, len(cols) + 1):
                ws.cell(row=current_row, column=c).border = thin_border
            current_row += 1
            
            # Column Headers
            for c_idx, col_name in enumerate(cols, 1):
                cell = ws.cell(row=current_row, column=c_idx, value=col_name)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal='center')
                cell.border = thin_border
            current_row += 1
            
            # Data Rows
            for _, row in group_f.iterrows():
                for c_idx, col_name in enumerate(cols, 1):
                    val = row.get(col_name, "")
                    cell = ws.cell(row=current_row, column=c_idx, value=val)
                    cell.border = thin_border
                    cell.alignment = Alignment(horizontal='left')
                    
                    # Format numbers/percentages
                    if col_name == '%EXEC':
                        # Try to format as percentage
                         try:
                             # Expected format 25.58% or 0.2558
                             s_val = str(val).replace('%', '').replace(',', '.')
                             f_val = float(s_val)
                             if f_val > 1.05: # Likely 25.58 meaning 25%
                                 f_val = f_val / 100.0
                             cell.value = f_val
                             cell.number_format = '0.00%'
                         except:
                             pass
                current_row += 1
            
            current_row += 1 # Spacer between fiscals
        
        current_row += 1 # Spacer between Regions

    # Autofit columns
    for column_cells in ws.columns:
        # Get column letter safely
        try:
            col_letter = get_column_letter(column_cells[0].column)
        except:
             continue
             
        length = 0
        for cell in column_cells:
            try:
                if cell.value:
                    length = max(length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[col_letter].width = min(length + 2, 50)
        
    wb.save(OUTPUT_FILE)
    print(f"Report generated: {OUTPUT_FILE}")

if __name__ == "__main__":
    df = load_data(INPUT_FILE)
    if df is not None:
        generate_report(df)
