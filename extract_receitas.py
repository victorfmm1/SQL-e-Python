 # =========================================================
# ETL - RECEITAS FINANÇAS PESSOAIS
# Autor: Victor
# Descrição: Extrai receitas tributadas e não tributadas
#            de múltiplos arquivos Excel (multi-ano),
#            consolida e gera base tratada.
# =========================================================

import pandas as pd
import os
import re

# CONFIGURAÇÕES
PASTA_LOAD = r"D:\ARQUIVOS\Python\Controle Financeiro Pessoal\ETL\EXTRACT"
ARQUIVO_SAIDA = r"D:\ARQUIVOS\Python\Controle Financeiro Pessoal\ETL\TRANSFORM\receitas_raw.xlsx"

# LISTA DE ARQUIVOS EXCEL
arquivos = [
    os.path.join(PASTA_LOAD, f)
    for f in os.listdir(PASTA_LOAD)
    if f.endswith(".xlsx") and not f.startswith("~$")
]

lista_receitas = []

for arquivo in arquivos:

    nome_arquivo = os.path.basename(arquivo)
    ano = re.search(r"\d{4}", nome_arquivo).group()
    print(f"Processando arquivo: {nome_arquivo} | Ano: {ano}")
    base = pd.ExcelFile(arquivo)

    # LOOP POR ABA (MÊS)

    for mes in base.sheet_names:
        # RECEITAS TRIBUTADAS (A2:C9)
        df_trib = pd.read_excel(
            arquivo,
            sheet_name=mes,
            usecols="A:C",
            skiprows=1,
            nrows=30
        )

        df_trib.columns = ["descricao", "teto", "realizado"]
        df_trib["mes"] = mes
        df_trib["ano"] = ano
        df_trib["tipo_receita"] = ""

        df_trib = df_trib.dropna(subset=["descricao"])
        lista_receitas.append(df_trib)


# CONSOLIDAÇÃO FINAL
df_receitas = pd.concat(lista_receitas, ignore_index=True)
df_receitas.to_excel(ARQUIVO_SAIDA, index=False)

print("ETL FINALIZADO COM SUCESSO!")
print(f"Arquivo gerado em: {ARQUIVO_SAIDA}")