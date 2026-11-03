import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
import re
import os

def clean_number(val):
    """Limpa e converte valores financeiros."""
    if pd.isna(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    v_str = str(val).strip()
    if ',' in v_str and '.' in v_str:
        if v_str.rfind(',') > v_str.rfind('.'): v_str = v_str.replace('.', '').replace(',', '.')
        else: v_str = v_str.replace(',', '')
    elif ',' in v_str: v_str = v_str.replace(',', '.')
    v_str = re.sub(r'[^\d\.\-]', '', v_str)
    try: return float(v_str)
    except: return 0.0

def process_file(file_input, path_depara, path_template, dt_lanc_contabil, dt_lanc, tipo_provisao):
    # Verificação de existência dos arquivos locais
    if not os.path.exists(path_depara):
        raise FileNotFoundError(f"Arquivo de mapping não encontrado: {path_depara}")
    if not os.path.exists(path_template):
        raise FileNotFoundError(f"Template SAP não encontrado: {path_template}")

    # 1. Carregar De/Para local
    df_mapa = pd.read_excel(path_depara)
    df_mapa.columns = df_mapa.columns.astype(str).str.strip().str.upper()
    df_mapa['COSTCENTER'] = df_mapa['COSTCENTER'].astype(str).str.strip().str.zfill(10)
    de_para_dict = df_mapa.set_index('COSTCENTER').to_dict('index')

    # 2. Carregar Input (Upload do usuário)
    input_df = pd.read_excel(file_input, header=None)
    
    # Criar Texto de Referência Padronizado (XBLNR) com base na data escolhida
    mes_ano = dt_lanc_contabil.strftime("%m/%Y")
    if tipo_provisao == "Férias":
        doc_ref = f"PROV FF {mes_ano}"
    else:
        doc_ref = f"PROV 13 {mes_ano}"

    # Identificar limites (Total Geral)
    start_idx = 4
    end_idx = len(input_df)
    for idx in range(start_idx, len(input_df)):
        val = str(input_df.iloc[idx, 3]).strip()
        if "Total Geral" in val:
            end_idx = idx
            break

    # Configuração de Colunas
    if tipo_provisao == "Férias":
        # Férias: AJ(35), AL(37), AP(41)
        col_base, col_inss, col_fgts = 35, 37, 41
        map_contas = {
            "base": ("FÉRIAS", "21501400", "61022000"),
            "inss": ("INSS FÉRIAS", "21539100", "61023050"),
            "fgts": ("FGTS FÉRIAS", "21539107", "61023030")
        }
    else:
        # 13º: AI(34), AK(36), AO(40)
        col_base, col_inss, col_fgts = 34, 36, 40
        map_contas = {
            "base": ("13º", "21501200", "61020000"),
            "inss": ("INSS EMPRESA 13", "21539106", "61023050"),
            "fgts": ("FGTS 13", "21539103", "61023030")
        }

    items = []
    first_company_code = ""

    for idx in range(start_idx, end_idx):
        row = input_df.iloc[idx]
        cc = str(row[3]).strip()[-10:]
        cc_info = de_para_dict.get(cc, {})
        comp = cc_info.get("COMPANYCODE", "")
        bp = cc_info.get("BUSINESSPLACE", "")

        if not comp: continue
        if not first_company_code: first_company_code = comp

        vals = [clean_number(row[col_base]), clean_number(row[col_inss]), clean_number(row[col_fgts])]
        keys = ["base", "inss", "fgts"]

        for v, k in zip(vals, keys):
            if round(v, 2) == 0: continue
            txt, c_passivo, c_despesa = map_contas[k]
            v_abs = abs(v)
            if v > 0:
                items.append((comp, c_passivo, txt, 0, v_abs, cc, bp))
                items.append((comp, c_despesa, txt, v_abs, 0, cc, bp))
            else:
                items.append((comp, c_passivo, txt, v_abs, 0, cc, bp))
                items.append((comp, c_despesa, txt, 0, v_abs, cc, bp))

    # 3. Preencher Template
    wb = load_workbook(path_template)
    ws = wb.active
    
    # Cabeçalho (Linha 9)
    h_map = {cell.value: cell.column for cell in ws[9] if cell.value}
    ws.cell(row=11, column=h_map["BUKRS"], value=first_company_code)
    ws.cell(row=11, column=h_map["BLART"], value="SA")
    
    # ==== AJUSTE DE DATA AQUI ====
    # Passamos o dt_lanc (que já é tipo Date) e aplicamos a formatação nativa do Excel
    ws.cell(row=11, column=h_map["BLDAT"], value=dt_lanc_contabil).number_format = 'DD/MM/YYYY'
    ws.cell(row=11, column=h_map["BUDAT"], value=dt_lanc).number_format = 'DD/MM/YYYY'
    # =============================
    
    ws.cell(row=11, column=h_map["WAERS"], value="BRL")
    ws.cell(row=11, column=h_map["XBLNR"], value=doc_ref)

    # Itens (Linha 15 e 16 para busca, 17 para escrita)
    i_map = {cell.value: cell.column for cell in ws[15] if cell.value}

    for i, item in enumerate(items):
        r = 17 + i
        c_code, conta, txt, deb, cred, costc, bplace = item
        
        if "BUKRS" in i_map: ws.cell(row=r, column=i_map["BUKRS"], value=c_code)
        if "HKONT" in i_map: ws.cell(row=r, column=i_map["HKONT"], value=conta)
        if "SGTXT" in i_map: ws.cell(row=r, column=i_map["SGTXT"], value=txt)
        if "WRSOL" in i_map and deb > 0: ws.cell(row=r, column=i_map["WRSOL"], value=deb)
        if "WRHAB" in i_map and cred > 0: ws.cell(row=r, column=i_map["WRHAB"], value=cred)
        if "KOSTL" in i_map: ws.cell(row=r, column=i_map["KOSTL"], value=costc)
        if "VBUND" in i_map: ws.cell(row=r, column=i_map["VBUND"], value=bplace)
        if "BUPLA" in i_map: ws.cell(row=r, column=i_map["BUPLA"], value=bplace)

    buffer = BytesIO()
    wb.save(buffer)
    return buffer.getvalue()