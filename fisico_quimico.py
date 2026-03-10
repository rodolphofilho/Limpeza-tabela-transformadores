import pandas as pd
import os

#juntar os arquivos excel gerados em um unico arquivo excel

PASTA_EXCEL = 'excel_fisicoquimico'

dfs = [] # Lista para armazenar os excel

for arq in os.listdir(PASTA_EXCEL):
    if arq.endswith('.xlsx'):
        caminho = os.path.join(PASTA_EXCEL, arq)
        df = pd.read_excel(caminho)
        dfs.append(df) # Adiciona o excel à lista

df_conectado = pd.concat(dfs, ignore_index=True) # Concatena os execel em um único execel

#print(df_conectado) # Exibe o DataFrame resultante da concatenação
#print(df_conectado.info()) # Exibe o número de linhas e colunas do DataFrame resultante da concatenação

# Colunas Selecionadas

colunas_selecionadas = [
    "TAG",
    "Amostra",
    "Série",
    "Subestação",
    "Data da Amostragem",
    "Densidade 20/4 °C (g/cm3)",
    "Cor (Não se aplica)",
    "Fator de Perdas (Tang. δ) a 90 °C (%)",
    "Teor de Água (mg/Kg)",
    "Índice de Neutralização (mgKOH/g)",
    "Tensão Interfacial (mN/m)",
    "Rigidez Dielétrica (kV)"
]

df_selecionado = df_conectado[colunas_selecionadas] # selecionar as colunas desejadas 

# print(df_selecionado) # Exibe o DataFrame resultante da seleção de colunas

# ALTERAR OS NOMES DAS COLUNAS 

novos_nomes = {
    "Amostra": "ID_AMOSTRA",
    "Série": "NUMERO_SERIE",
    "Subestação": "SUBESTACAO",
    "Data da Amostragem": "DATA",
    "Densidade 20/4 °C (g/cm3)": "DENSIDADE",
    "Cor (Não se aplica)": "COR",
    "Fator de Perdas (Tang. δ) a 90 °C (%)": "FATOR_PERDAS",
    "Teor de Água (mg/Kg)": "TEOR_AGUA",
    "Índice de Neutralização (mgKOH/g)": "INDICE_NEUTRALIZACAO",
    "Tensão Interfacial (mN/m)": "TENSAO_INTERFACIAL",
    "Rigidez Dielétrica (kV)": "RIGIDEZ_DIELETRICA"
}

df_selecionado = df_selecionado.rename(columns=novos_nomes) # Renomear as colunas do EXECEL 

#print(df_selecionado) # Exibe o DataFrame resultante do renomeamento das colunas

# ORGANIZAR AS COLUNAS 

df_selecionado["TAG"] = df_selecionado["TAG"].astype(str).str.replace(r"[^A-Za-z0-9]", "", regex=True) # Limpar os dados das colunas "TAG", "AMOSTRA" e "NUMERO_SERIE" removendo caracteres especiais e espaços em branco
df_selecionado["ID_AMOSTRA"] = df_selecionado["ID_AMOSTRA"].astype(str).str.replace(r"[^A-Za-z0-9]", "", regex=True)
df_selecionado["NUMERO_SERIE"] = df_selecionado["NUMERO_SERIE"].astype(str).str.replace(r"[^A-Za-z0-9]", "", regex=True)


# SALVAR O DATAFRAME EM UM ARQUIVO EXCEL

PASTA_SAIDA = 'saida'

if not os.path.exists(PASTA_SAIDA):
    os.makedirs(PASTA_SAIDA)

caminho_saida = os.path.join(PASTA_SAIDA, 'fisicoquimica.xlsx')

df_selecionado.to_excel(caminho_saida, index=False) # Salvar o DataFrame em um arquivo Excel

print(f"Arquivo salvo em: {caminho_saida}")  