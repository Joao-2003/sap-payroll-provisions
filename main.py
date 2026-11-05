import datetime
import io
import os
import sys

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

from converter import process_file

AQUI = os.path.dirname(os.path.abspath(__file__))

# Arquivos fixos na pasta. Sem eles (ou no navegador), o app usa os exemplos de demo/.
FILE_DEPARA = os.path.join(AQUI, "de_para_CC.xlsx")
FILE_TEMPLATE = os.path.join(AQUI, "JournalEntry_Template (1).XLSX")
MODO_DEMO = sys.platform == "emscripten" or not (os.path.exists(FILE_DEPARA) and os.path.exists(FILE_TEMPLATE))
if MODO_DEMO:
    FILE_DEPARA = os.path.join(AQUI, "de_para_CC_exemplo.xlsx")
    FILE_TEMPLATE = os.path.join(AQUI, "demo", "template_lancamento_exemplo.xlsx")
RELATORIO_EXEMPLO = os.path.join(AQUI, "demo", "relatorio_provisoes_exemplo.xlsx")

st.set_page_config(page_title="Provisões SAP", page_icon="📄")

st.title("Automação de Provisões (férias e 13º)")
st.caption("Converte relatórios de domínio para o padrão SAP S/4HANA.")
if MODO_DEMO:
    st.info("Demonstração: de-para de centros de custo, template e relatório são fictícios. "
            "A conversão é a mesma usada pela Contabilidade.")
st.divider()

col1, col2 = st.columns([1, 1.5])

with col1:
    tipo = st.radio("Tipo de Provisão:", ["Férias", "13º"])
    dt_cont = st.date_input("Data Lanç. Contábil:", datetime.date.today(), format="DD/MM/YYYY")
    dt_lanc = st.date_input("Data Lançamento:", datetime.date.today(), format="DD/MM/YYYY")

with col2:
    uploaded_file = st.file_uploader("Upload do arquivo de Provisão (XLSX):", type=["xlsx"])
    if MODO_DEMO:
        usar_exemplo = st.checkbox("Usar o relatório de exemplo", value=uploaded_file is None)
        with open(RELATORIO_EXEMPLO, "rb") as fh:
            st.download_button("⬇️ Baixar o relatório de exemplo", fh.read(), file_name="relatorio_provisoes_exemplo.xlsx")
        if usar_exemplo and uploaded_file is None:
            uploaded_file = RELATORIO_EXEMPLO

    depara_exists = os.path.exists(FILE_DEPARA)
    temp_exists = os.path.exists(FILE_TEMPLATE)
    if not depara_exists:
        st.error(f"⚠️ Arquivo '{os.path.basename(FILE_DEPARA)}' não encontrado na pasta!")
    if not temp_exists:
        st.error(f"⚠️ Arquivo '{os.path.basename(FILE_TEMPLATE)}' não encontrado na pasta!")


def previa(out_bytes):
    """Lê de volta os itens gravados no template (a partir da linha 17) para conferência na tela."""
    ws = load_workbook(io.BytesIO(out_bytes)).active
    cab = {cell.value: cell.column for cell in ws[15] if cell.value}
    linhas = []
    for r in range(17, ws.max_row + 1):
        if not ws.cell(row=r, column=cab["HKONT"]).value:
            continue
        linhas.append({"Empresa": ws.cell(row=r, column=cab["BUKRS"]).value, "Conta": ws.cell(row=r, column=cab["HKONT"]).value,
                       "Texto": ws.cell(row=r, column=cab["SGTXT"]).value,
                       "Débito": ws.cell(row=r, column=cab["WRSOL"]).value or 0.0,
                       "Crédito": ws.cell(row=r, column=cab["WRHAB"]).value or 0.0,
                       "Centro de custo": ws.cell(row=r, column=cab["KOSTL"]).value})
    return pd.DataFrame(linhas)


if st.button("Gerar Arquivo SAP", type="primary", disabled=not uploaded_file):
    if depara_exists and temp_exists:
        try:
            with st.spinner("Processando..."):
                out_bytes = process_file(uploaded_file, FILE_DEPARA, FILE_TEMPLATE, dt_cont, dt_lanc, tipo)

            st.success("Pronto!")
            st.download_button(
                label="📥 Baixar Arquivo SAP",
                data=out_bytes,
                file_name=f"SAP_{tipo}_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            df = previa(out_bytes)
            if not df.empty:
                c1, c2, c3 = st.columns(3)
                c1.metric("Partidas", len(df))
                c2.metric("Débitos", f"R$ {df['Débito'].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                c3.metric("Créditos", f"R$ {df['Crédito'].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                st.caption("Débito = despesa e crédito = passivo de provisão; valores negativos no relatório invertem a natureza (estorno).")
                st.dataframe(df, hide_index=True, width="stretch",
                             column_config={"Débito": st.column_config.NumberColumn(format="R$ %.2f"),
                                            "Crédito": st.column_config.NumberColumn(format="R$ %.2f")})
        except Exception as e:
            st.error(f"Erro: {e}")
    else:
        st.error("Verifique se os arquivos de configuração estão na mesma pasta do código.")
