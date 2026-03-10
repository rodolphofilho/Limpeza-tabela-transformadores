import pandas as pd 
import os 

# juntas os arquivos excel 

PASTA_EXCEL = "excel_cromatografia"

dfs = [] # lista para armazenar 

# conectar os aquivos excel 

for arq in os.listdir(PASTA_EXCEL):
    if arq.endswith('.xlsx'):
        caminho = os.path.join(PASTA_EXCEL, arq)
        df = pd.read_excel(caminho)
        dfs.append(df)

df_conectar = pd.concat(dfs, ignore_index=True) 

#print(df_conectar)

# Colunas Selecionadas 

colunas_selecionadas = [
    'TAG',
    'Amostra', 
    'Série',
    'Subestação', 
    'Data da Amostragem', 
    'Hidrogênio (µL/L)', 
    'Oxigênio (µL/L)', 
    'Nitrogênio (µL/L)', 
    'Monóxido de Carbono (µL/L)',
    'Metano (µL/L)', 
    'Dióxido de Carbono (µL/L)', 
    'Etileno (µL/L)', 
    'Etano (µL/L)', 
    'Acetileno (µL/L)',
    'Total de Gases Combustíveis (µL/L)', 
    'Total de gases Dissolvidos (µL/L)', 
    'Relação CO2/CO (-)'
]

df_selecionadas = df_conectar[colunas_selecionadas]

#print(df_selecionadas)

# Alterar os nomes das colunas

novos_nomes = {
    'Amostra': 'ID_AMOSTRA',  
    'Série': 'NUMERO_SERIE',
    'Subestação': 'SUBESTACAO', 
    'Data da Amostragem': 'DATA', 
    'Hidrogênio (µL/L)': 'HIDROGENIO', 
    'Oxigênio (µL/L)': 'OXIGENIO', 
    'Nitrogênio (µL/L)': 'NITROGENIO', 
    'Monóxido de Carbono (µL/L)': 'MONOXIDO DE CARBONO',
    'Metano (µL/L)': 'METANO', 
    'Dióxido de Carbono (µL/L)': 'DIOXIDO DE CARBONO', 
    'Etileno (µL/L)': 'ETILENO', 
    'Etano (µL/L)': 'ETANO', 
    'Acetileno (µL/L)': 'ACETILENO',
    'Total de Gases Combustíveis (µL/L)': 'TOTAL DE GASES COMBUSTIVEIS', 
    'Total de gases Dissolvidos (µL/L)': 'TOTAL DE GASES DISSOLVIDOS', 
    'Relação CO2/CO (-)': 'RELACAO CO2/CO'
}

df_selecionadas = df_selecionadas.rename(columns=novos_nomes)

#print(df_selecionadas)

# Organizar as colunas 

df_selecionadas["TAG"] = df_selecionadas["TAG"].astype(str).str.replace(r"[^A-Za-z0-9]", "", regex=True) # Limpar os dados das colunas "TAG", "AMOSTRA" e "NUMERO_SERIE" removendo caracteres especiais e espaços em branco
df_selecionadas["ID_AMOSTRA"] = df_selecionadas["ID_AMOSTRA"].astype(str).str.replace(r"[^A-Za-z0-9]", "", regex=True)
df_selecionadas["NUMERO_SERIE"] = df_selecionadas["NUMERO_SERIE"].astype(str).str.replace(r"[^A-Za-z0-9]", "", regex=True)

coluna = [
    "OXIGENIO",
    "HIDROGENIO",
    "MONOXIDO DE CARBONO",
    "METANO",
    "DIOXIDO DE CARBONO",
    "ETILENO",
    "ETANO",
    "ACETILENO",
    "TOTAL DE GASES COMBUSTIVEIS",
    "TOTAL DE GASES DISSOLVIDOS"
]

for col in coluna:
    df_selecionadas[col] = (
        df_selecionadas[col]
        .astype(str)                        # força virar texto
        .str.replace(",", ".", regex=False) # troca vírgula por ponto
        .str.strip()                        # remove espaços
    )

    df_selecionadas[col] = pd.to_numeric(
        df_selecionadas[col],
        errors="coerce"                     # o que não virar número vira NaN
    )

    df_selecionadas[col] = (
        df_selecionadas[col]
        .fillna(0)                                  # NaN vira 0
        .replace([float("inf"), float("-inf")], 0)  # remove infinito
        .astype(int)                                # agora sim vira inteiro
    )



# SALVAR O DATAFRAME EM UM ARQUIVO EXCEL

PASTA_SAIDA = 'saida'

if not os.path.exists(PASTA_SAIDA):
    os.makedirs(PASTA_SAIDA)

caminho_saida = os.path.join(PASTA_SAIDA, 'cromatografia.xlsx')

df_selecionadas.to_excel(caminho_saida, index=False) # Salvar o DataFrame em um arquivo Excel

print(f"Arquivo salvo em: {caminho_saida}")