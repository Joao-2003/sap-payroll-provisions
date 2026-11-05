"""Gera os arquivos fictícios da demonstração (pasta demo/).

- template_lancamento_exemplo.xlsx: estrutura mínima do template "Upload General Journal Entries" do SAP
  (cabeçalho na linha 9, itens na linha 15, dados a partir da 11 e da 17), só com as colunas usadas;
- relatorio_ferias_exemplo.xlsx e relatorio_13_exemplo.xlsx: relatório de provisões da folha no layout
  que o conversor lê (centro de custo na coluna D, valores nas colunas AI a AP, "Total Geral" no fim).

    python tools/gerar_exemplos.py
"""
import os
import random

from openpyxl import Workbook

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEMO = os.path.join(RAIZ, "demo")
rnd = random.Random(13)
CENTROS = [("Produção - Usinagem", "0101010101"), ("Produção - Montagem", "0101020101"),
           ("Administrativo - Filial", "0201010101")]


def template():
    wb = Workbook()
    ws = wb.active
    ws.title = "Lançamentos"
    ws["A1"] = "Template de exemplo (estrutura simplificada do upload de lançamentos contábeis do SAP S/4HANA)"
    for col, campo in enumerate(["BUKRS", "BLART", "BLDAT", "BUDAT", "WAERS", "XBLNR"], start=1):
        ws.cell(row=8, column=col, value={"BUKRS": "Empresa", "BLART": "Tipo doc.", "BLDAT": "Data doc.",
                                           "BUDAT": "Data lanç.", "WAERS": "Moeda", "XBLNR": "Referência"}[campo])
        ws.cell(row=9, column=col, value=campo)
    for col, campo in enumerate(["BUKRS", "HKONT", "SGTXT", "WRSOL", "WRHAB", "KOSTL", "VBUND", "BUPLA"], start=1):
        ws.cell(row=14, column=col, value={"BUKRS": "Empresa", "HKONT": "Conta", "SGTXT": "Texto", "WRSOL": "Débito",
                                            "WRHAB": "Crédito", "KOSTL": "Centro de custo", "VBUND": "Sociedade parceira",
                                            "BUPLA": "Local de negócios"}[campo])
        ws.cell(row=15, column=col, value=campo)
    wb.save(os.path.join(DEMO, "template_lancamento_exemplo.xlsx"))


def relatorio(nome, tipo):
    wb = Workbook()
    ws = wb.active
    ws["A1"] = f"Relatório de provisão de {tipo} (fictício)"
    ws["A2"] = "Competência: mês corrente"
    ws.cell(row=4, column=4, value="Centro de custo")
    for c, rot in zip(range(35, 43), ["13º", "Férias", "INSS 13º", "INSS Férias", "", "", "FGTS 13º", "FGTS Férias"]):
        ws.cell(row=4, column=c, value=rot)
    linha = 5
    for desc, cc in CENTROS:
        for _ in range(rnd.randint(4, 7)):
            ws.cell(row=linha, column=4, value=f"{desc} - {cc}")
            base = round(rnd.uniform(800, 9000), 2) * (1 if rnd.random() > 0.08 else -0.3)  # algumas reversões
            ws.cell(row=linha, column=35, value=round(base, 2))          # AI: 13º
            ws.cell(row=linha, column=36, value=round(base * 1.1, 2))    # AJ: férias
            ws.cell(row=linha, column=37, value=round(base * 0.2, 2))    # AK: INSS 13º
            ws.cell(row=linha, column=38, value=round(base * 0.22, 2))   # AL: INSS férias
            ws.cell(row=linha, column=41, value=round(base * 0.08, 2))   # AO: FGTS 13º
            ws.cell(row=linha, column=42, value=round(base * 0.088, 2))  # AP: FGTS férias
            linha += 1
    ws.cell(row=linha, column=4, value="Total Geral")
    wb.save(os.path.join(DEMO, nome))


if __name__ == "__main__":
    os.makedirs(DEMO, exist_ok=True)
    template()
    relatorio("relatorio_provisoes_exemplo.xlsx", "férias e 13º")
    print("demo/ pronto")
