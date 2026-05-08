'''
Ideia Geral do código para identificação dos estudantes em RAPP em 2026:
1. Utilizar dados de 2025 de estudantes em RAPP;
2. Utilizar informações do Relatório Geral Matrículas (2026 e atual) para atualizar as informações da escola e DIREC dos estudantes;

4. Deixar somente componentes da BNCC e estudantes dos Anos Finais Ensino Fundamental e Ensino Médio (6º ano EF até 3ª Série EM);
5. Padronizar CPFs para não ter diferenças de formatação entre as bases;
6. Excluir duplicatas da combinação de componente + CPF;
7. Conseguir informações de necessidades especiais por CPF em algum relatório;
8. Conseguir informações de turno por CPF em algum relatório;
9. Incluir informações de necessidades especiais e turno na base principal;
10. Fazer as segmentações e contagem de interesse: estudantes por DIREC; por componente; por turno; por Série; necessidades especiais etc.
'''

# Importação das bibliotecas
import pandas as pd
import glob
import os
from tqdm import tqdm  # Para barra de progresso
import numpy as np
import warnings
warnings.filterwarnings('ignore')
import openpyxl

# Nessa primeira vez, não farei os passos 1 e 2, pois vou utilizar os dados prontos do Uanderson.
# Nas próximas eu faço os passos 1 e 2.

# Usei a base do Uanderson da aba '2025' acrescentando as informações da aba 'NAO_ENCONTRADOS_POR_COMPONENTE'.

# Extrair os dados de 2025 e 2026 em Excel e passar para dataframe
df = pd.read_excel(r"D:\Scripts_Python\FGV\RAPP_2026\20260430_RELATORIO DE ONDE ESTA O ESTUDANTE EM RAPP.xlsx", sheet_name="TOTAL RAPP")


# Trocar nomenclatura das séries para padronizar:
mapeamento = {
    '6º Ano': '6º ANO',
    '7º Ano': '7º ANO',
    '8º Ano': '8º ANO',
    '9º Ano': '9º ANO'
}

df['SÉRIE'] = df['SÉRIE'].replace(mapeamento)


# Criar coluna de "ETAPA_RESUMIDA" para indicar Anos Finais ou Ensino Médio, de acordo com a série
mapeamento_etapa = {
    '1ª SÉRIE': 'Ensino Médio',
    '2ª SÉRIE': 'Ensino Médio',
    '3ª SÉRIE': 'Ensino Médio',
    '6º ANO': 'Ens. Fund. - Anos Finais',
    '7º ANO': 'Ens. Fund. - Anos Finais',
    '8º ANO': 'Ens. Fund. - Anos Finais',
    '9º ANO': 'Ens. Fund. - Anos Finais'
}

df['ETAPA_RESUMIDA'] = df['SÉRIE'].map(mapeamento_etapa)

# Padronizar CPF: manter apenas dígitos, completar com zeros à esquerda e formatar como XXX.XXX.XXX-XX
df['CPF_Padronizado'] = (
    df['CPF ALUNO']
        .astype(str)
        .str.replace(r'\D', '', regex=True)   # remove tudo que não é dígito
        .str.zfill(11)                        # completa com zeros à esquerda
        .str.replace(
            r'(\d{3})(\d{3})(\d{3})(\d{2})',
            r'\1.\2.\3-\4',
            regex=True
        )
)


# 3. Concatenar com as 16 planilhas (1 para cada DIREC) dos 'Relatórios de Acompanhamento das Turmas de Progressão Parcial' (SigEduc)
###########################
# RELATÓRIO DE ACOMPANHAMENTO DE TURMAS EM PROGRESSÃO PARCIALR (SIGEDUC) - 16 PLANILHAS (1 para cada DIREC) com informações por componente para cada estudante
# caminho da pasta onde estão os arquivos
pasta = r"C:\Users\hugob\Downloads\Alunos_RAPP_2026"

# lista todos os arquivos .xlsx da pasta
arquivos = glob.glob(os.path.join(pasta, "*.xlsx"))

# lista para armazenar os dataframes
dfs = []

for arquivo in tqdm(arquivos, desc="Processando arquivos"):
    # lê cada arquivo, pulando as 2 primeiras linhas
    df_unico = pd.read_excel(arquivo, skiprows=2)
    dfs.append(df_unico)

# concatena todos em um único dataframe
df_relatorio_RAPP = pd.concat(dfs, ignore_index=True)


# Manter apenas as colunas de interesse:
df_relatorio_RAPP = df_relatorio_RAPP[
    ['DIREC', 
     'ESCOLA',
     'INEP ESCOLA',
     'ETAPA ENSINO',
     'SÉRIE',
     'COMPONENTE CURRICULAR',
     'CPF',
     'ESTUDANTE',
     'MATRÍCULA']]


# Padronizar CPF: manter apenas dígitos, completar com zeros à esquerda e formatar como XXX.XXX.XXX-XX
df_relatorio_RAPP['CPF_Padronizado'] = (
    df_relatorio_RAPP['CPF']
        .astype(str)
        .str.replace(r'\D', '', regex=True)   # remove tudo que não é dígito
        .str.zfill(11)                        # completa com zeros à esquerda
        .str.replace(
            r'(\d{3})(\d{3})(\d{3})(\d{2})',
            r'\1.\2.\3-\4',
            regex=True
        )
)


# Trocar nomenclatura das séries para padronizar:
mapeamento = {
    '1º Período (1ª Série)': '1ª SÉRIE',
    '2º Período (2ª Série)': '2ª SÉRIE',
    '3º Período (3ª Série)': '3ª SÉRIE',
    '6º Ano': '6º ANO',
    '7º Ano': '7º ANO',
    '8º Ano': '8º ANO',
    '9º Ano': '9º ANO',
    '5º PERÍODO (8° E 9° ANO - ANOS FINAIS)': '9º ANO',
    '4º PERÍODO (6° E 7° ANO - ANOS FINAIS)': '7º ANO',
    'TURMA I (6° E 7° ANOS)': '7º ANO',
    'TURMA II (8° E 9° ANOS)': '9º ANO'
}

df_relatorio_RAPP['SÉRIE'] = df_relatorio_RAPP['SÉRIE'].replace(mapeamento)


# Criar coluna de "ETAPA_RESUMIDA" para indicar Anos Finais ou Ensino Médio, de acordo com a série
mapeamento_etapa = {
    '1ª SÉRIE': 'Ensino Médio',
    '2ª SÉRIE': 'Ensino Médio',
    '3ª SÉRIE': 'Ensino Médio',
    '6º ANO': 'Ens. Fund. - Anos Finais',
    '7º ANO': 'Ens. Fund. - Anos Finais',
    '8º ANO': 'Ens. Fund. - Anos Finais',
    '9º ANO': 'Ens. Fund. - Anos Finais'
}

df_relatorio_RAPP['ETAPA_RESUMIDA'] = df_relatorio_RAPP['SÉRIE'].map(mapeamento_etapa)


# Padronizar os nomes das colunas para concatenar os 2 dataframes
df = df.rename(columns={
    'ALUNO': 'ESTUDANTE',
    'CPF ALUNO': 'CPF',
    'MATRÍCULA ESTUDANTE': 'MATRÍCULA'
})


# Concatenar os 2 dataframes
df_concat = pd.concat([df, df_relatorio_RAPP], ignore_index=True)

# Excluir coluna 'CPF' para só usar o 'CPF_Padronizado' e evitar confusão
df_concat = df_concat.drop(columns=['CPF'])

# 4. Deixar somente componentes da BNCC e estudantes dos Anos Finais Ensino Fundamental e Ensino Médio (6º ano EF até 3ª Série EM);
# Nesse caso já estão somente esses valores

# 5. Padronizar CPFs para não ter diferenças de formatação entre as bases;
# Já está padronizado nos 2 dataframnes e na concatenação na coluna 'CPF_Padronizado'

# 6. Excluir duplicatas da combinação de componente + CPF
# A base tem alguns erros de mesmo estudante em séries e escolas diferentes
total_duplicadas = df_concat.duplicated(subset=['CPF_Padronizado', 'COMPONENTE CURRICULAR']).sum()
total_duplicadas


df_final = df_concat.drop_duplicates(subset=['CPF_Padronizado', 'COMPONENTE CURRICULAR'], keep='first', ignore_index=True)



#################################### ANÁLISES ####################################
# Estudantes por DIREC
# Contagem de CPF_PAdronizado distinto por DIREC
direc = df_final.groupby('DIREC')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})
direc

total_direc = direc['Quantidade de Estudantes Distintos'].sum()
total_direc

# Estudantes por Componente Curricular
# Contagem de CPF_Padronizado distinto por Componente Curricular
componente = df_final.groupby('COMPONENTE CURRICULAR')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})
componente

total_componente = componente['Quantidade de Estudantes Distintos'].sum()
total_componente


# Estudantes por Série
# Contagem de CPF_Padronizado distinto por Série
serie = df_final.groupby('SÉRIE')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})
serie

total_serie = serie['Quantidade de Estudantes Distintos'].sum()
total_serie

# Juntar os valores por DIREC, Componente e Série e salvar em um arquivo Excel, com cada tabela em uma aba diferente
with pd.ExcelWriter(r"D:\Scripts_Python\FGV\RAPP_2026\20260508_analises_RAPP.xlsx") as writer:
    df_final.to_excel(writer, sheet_name='Base RAPP', index=False)
    direc.to_excel(writer, sheet_name='Estudantes por DIREC', index=False)
    componente.to_excel(writer, sheet_name='Estudantes por Componente', index=False)
    serie.to_excel(writer, sheet_name='Estudantes por Série', index=False)















'''


7. Conseguir informações de necessidades especiais por CPF em algum relatório;
8. Conseguir informações de turno por CPF em algum relatório;
9. Incluir informações de necessidades especiais e turno na base principal;
10. Fazer as segmentações e contagem de interesse: estudantes por DIREC; por componente; por turno; por Série; necessidades especiais etc.
'''

