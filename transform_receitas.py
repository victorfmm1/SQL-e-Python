# =========================================================
# ETL - RECEITAS FINANÇAS PESSOAIS
# Autor: Victor
# Descrição: TRANSFORM para ajustar descrição, remover caracteres especiais e afins
# =========================================================

import pandas as pd
import re
import unicodedata

# CONFIGURAÇÕES
ARQUIVO_ENTRADA = r"D:\ARQUIVOS\Python\Controle Financeiro Pessoal\ETL\TRANSFORM\receitas_raw.xlsx"
ARQUIVO_SAIDA = r"D:\ARQUIVOS\Python\Controle Financeiro Pessoal\ETL\LOAD\receitas_tratadas.xlsx"
    
def normalizar_texto(texto):
    if not isinstance(texto, str):
        return texto
    # lower case
    texto = texto.lower()
    # remover acentos
    texto = unicodedata.normalize("NFKD", texto)
    texto = texto.encode("ASCII", "ignore").decode("utf-8")
    # remover caracteres especiais (mantém letras, números e espaço)
    texto = re.sub(r"[^a-z0-9\s]", "", texto)
    # remover espaços duplicados
    texto = re.sub(r"\s+", " ", texto).strip()

    return texto

def padronizar_13(texto):
    if not isinstance(texto, str):
        return texto
    t = texto
    if t.startswith("13"):
        if "victor" in t:
            return "13 Victor"
        if "suelynn" in t or "suely" in t or "suellyn" in t:
            return "13 Suellyn"
    return texto

def padronizar_bonificacao(texto):
    if not isinstance(texto, str):
        return texto
    t = texto

    if t.startswith("bonificacao"):
        if "victor" in t:
            return "bonificacao Victor"
        if "suelynn" in t or "suely" in t or "suellyn" in t:
            return "bonificacao Suellyn"
        else:
            return "bonificacao Suellyn"
    return texto

def padronizar_vales(texto):
    if not isinstance(texto, str):
        return texto
    t = texto

    if t.startswith("ticket"):
        return "vale alimentacao"
    
    if t.startswith("vale"):
        if "refeicao" in t:
            return "vale alimentacao"
    return texto

def padronizar_comissao(texto):
    if not isinstance(texto, str):
        return texto
    t = texto

    if t.startswith("comissao"):
        if "magalu" in t:
            return "parceiro magalu"
        return "comissao victor"
    return texto

def padronizar_fgts(texto):
    if not isinstance(texto,str):
        return texto
    t = texto
    if t.startswith("fgts"):
        if "victor" in t:
            return "fgts victor"
        if "suelynn" in t or "suely" in t or "suellyn" in t or "varoa" in t:
            return "fgts suellyn"
    return texto

def padronizar_seguro(texto):
    if not isinstance(texto, str):
        return texto
    t = texto

    if t.startswith("seguro"):
        return "seguro desemprego"
    return texto

def tipo_receita(texto):
    if not isinstance(texto, str):
        return texto
    t = texto

    if t.startswith("vale"):
        return "Nao Tributada"
    return "Tributada" 

# leitura da base raw
df = pd.read_excel(ARQUIVO_ENTRADA)

# cria coluna normalizada (NUNCA sobrescreve a original)
df["descricao_normalizada"] = (
    df["descricao"]
    .apply(normalizar_texto)
    .apply(padronizar_bonificacao)
    .apply(padronizar_13)
    .apply(padronizar_vales)
    .apply(padronizar_comissao)
    .apply(padronizar_fgts)
    .apply(padronizar_seguro)
)

df["tipo_receita"] = df["descricao_normalizada"].apply(tipo_receita)

#remove linhas RECEITAS Não tributadas e TOTAL
tupla_linhas = ("RECEITAS NÃO TRIBUTÁVEIS", "TOTAL", "DESCRICAO")
df = df.drop(df[df["descricao"].isin(tupla_linhas)].index)

# salva base tratada
df.to_excel(ARQUIVO_SAIDA, index=False)

print("TRANSFORM FINALIZADO COM SUCESSO!")
print(f"Arquivo gerado em: {ARQUIVO_SAIDA}")