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
ARQUIVO_SAIDA = r"D:\ARQUIVOS\Python\Controle Financeiro Pessoal\ETL\TRANSFORM\despesas_raw.xlsx"

# LISTA DE ARQUIVOS EXCEL
arquivos = [
    os.path.join(PASTA_LOAD, f)
    for f in os.listdir(PASTA_LOAD)
    if f.endswith(".xlsx") and not f.startswith("~$")
]

lista_despesas = []

for arquivo in arquivos:

    nome_arquivo = os.path.basename(arquivo)
    ano = re.search(r"\d{4}", nome_arquivo).group()
    print(f"Processando arquivo: {nome_arquivo} | Ano: {ano}")
    base = pd.ExcelFile(arquivo)

    # LOOP POR ABA (MÊS)

    for mes in base.sheet_names: 
        df_despesas = pd.read_excel(
            arquivo,
            sheet_name=mes,
            usecols="E:G",
            skiprows=1,
            nrows=25
        )

        df_despesas.columns = ["descricao", "teto", "realizado"]
        df_despesas["mes"] = mes
        df_despesas["ano"] = ano
        df_despesas["categoria"] = " "

        df_despesas = df_despesas.dropna(subset=["descricao"])
        lista_despesas.append(df_despesas)


# CONSOLIDAÇÃO FINAL
df_despesas = pd.concat(lista_despesas, ignore_index=True)
df_despesas.to_excel(ARQUIVO_SAIDA, index=False)

print("ETL FINALIZADO COM SUCESSO!")
print(f"Arquivo gerado em: {ARQUIVO_SAIDA}")