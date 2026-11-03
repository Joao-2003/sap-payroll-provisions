import streamlit as st
import datetime
import os
from converter import process_file

# Configuração dos arquivos fixos na pasta
FILE_DEPARA = "de_para_CC.xlsx"
FILE_TEMPLATE = "JournalEntry_Template (1).XLSX"

st.set_page_config(page_title="Provisões SAP", page_icon="📄")

st.title("Automação de Provisões (férias e 13º)")
st.caption("Converte relatórios de domínio para o padrão SAP S/4HANA.")
st.divider()

# Interface simplificada
col1, col2 = st.columns([1, 1.5])

with col1:
    tipo = st.radio("Tipo de Provisão:", ["Férias", "13º"])
    dt_cont = st.date_input("Data Lanç. Contábil:", datetime.date.today())
    dt_lanc = st.date_input("Data Lançamento:", datetime.date.today())

with col2:
    uploaded_file = st.file_uploader("Upload do arquivo de Provisão (XLSX):", type=["xlsx"])
    
    # Verificação visual se os arquivos base estão na pasta
    depara_exists = os.path.exists(FILE_DEPARA)
    temp_exists = os.path.exists(FILE_TEMPLATE)
    
    if not depara_exists: st.error(f"⚠️ Arquivo '{FILE_DEPARA}' não encontrado na pasta!")
    if not temp_exists: st.error(f"⚠️ Arquivo '{FILE_TEMPLATE}' não encontrado na pasta!")

if st.button("Gerar Arquivo SAP", type="primary", disabled=not uploaded_file):
    if depara_exists and temp_exists:
        try:
            with st.spinner('Processando...'):
                out_bytes = process_file(uploaded_file, FILE_DEPARA, FILE_TEMPLATE, dt_cont, dt_lanc, tipo)
                
            st.success("Pronto!")
            st.download_button(
                label="📥 Baixar Arquivo SAP",
                data=out_bytes,
                file_name=f"SAP_{tipo}_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Erro: {e}")
    else:
        st.error("Verifique se os arquivos de configuração estão na mesma pasta do código.")