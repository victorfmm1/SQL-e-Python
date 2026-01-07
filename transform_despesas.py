# =========================================================
# etl - receitas finanças pessoais
# autor: victor
# descrição: motor de regras para normalização de descrições
# =========================================================

import pandas as pd
import re
import unicodedata

# configurações
arquivo_entrada = r"d:\arquivos\python\controle financeiro pessoal\etl\transform\despesas_raw.xlsx"
arquivo_saida = r"d:\arquivos\python\controle financeiro pessoal\etl\load\despesas_tratadas.xlsx"

def motor_de_regras(texto):
    if not isinstance(texto, str):
        return texto

    # normalização básica
    t = texto.lower()
    t = unicodedata.normalize("NFKD", t).encode("ascii", "ignore").decode("utf-8")
    t = re.sub(r"[^a-z0-9\s]", "", t)
    t = re.sub(r"\s+", " ", t).strip()

    # regras específicas
    if "aporte" in t or re.search(r"\b\d+\s*investimento\b", t):
        if "caixa" in t:
            return "aporte de caixa"
        if "carro" in t:
            return "aporte para carro"
        return "aporte de investimento"

    if t.startswith("nu"):
        if "victor" in t:
            return "nubank victor"
        if "suellyn" in t:
            return "nubank suellyn"
        return t

    if t.startswith("coisas") or t.startswith("coisa"):
        if "conrado" in t or t.startswith("coisa"):
            return "coisas conrado"
        if "casa" in t:
            return "coisas casa"
        if "carro" in t:
            return "coisas carro"
        return t

    if t.startswith("parcela"):
        return "parcela carro"

    if t.startswith("quitar"):
        return "quitar menor divida"

    return t

def definir_categoria(texto_normalizado):
    # dicionário de mapeamento: "termo_chave": "categoria"
    mapeamento_categorias = {
        # --- Deus ---
        "dizimo": "Deus",
        "oferta" : "Deus",

        # --- investimentos ---
        "aporte de caixa": "caixa",
        "aporte para carro": "investimentos",
        "aporte de investimento": "investimentos",
        "conrado": "investimentos",
        "bernardo": "investimentos",
        
        # --- cartões e bancos ---
        "nubank victor": "cartao de credito",
        "nubank suellyn": "cartao de credito",
        "xp casal": "cartao de credito",
        
        # --- moradia e família ---
        "coisas casa": "moradia",
        "coisas conrado": "dependentes",
        "aluguel": "moradia",
        "energia": "moradia",
        "internet": "moradia",
        "felicidade suellyn": "lazer",
        "felicidade victor": "lazer",
        
        # --- transporte ---
        "coisas carro": "transporte",
        "parcela carro": "transporte",
        "combustivel": "transporte",
        "gasolina": "transporte",
        "uber": "transporte",
        
        # --- dívidas ---
        "quitar menor divida": "dividas",
        "financiamento casa": "dividas",
        
        # --- alimentaçao e consumo (exemplos comuns) ---
        "ifood": "alimentacao",
        "supermercado": "alimentacao",
        "restaurante": "lazer",
        "farmacia": "saude",
    }
    
    return mapeamento_categorias.get(texto_normalizado, "outros")
# leitura da base raw
df = pd.read_excel(arquivo_entrada)

# cria coluna normalizada
df["descricao_normalizada"] = df["descricao"].apply(motor_de_regras)
df["categoria"] = df["descricao_normalizada"].apply(definir_categoria)

tupla_linhas = ("TOTAL", "DESCRICAO")
df = df.drop(df[df["descricao"].isin(tupla_linhas)].index)

# salva base tratada
df.to_excel(arquivo_saida, index=False)

print("transform finalizado com sucesso!")
print(f"arquivo gerado em: {arquivo_saida}")