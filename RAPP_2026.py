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
import re

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

# Excluir coluna 'CPF ALUNO' para só usar o 'CPF_Padronizado' e evitar confusão
df = df.drop(columns=['CPF ALUNO'])

# 4. Deixar somente componentes da BNCC e estudantes dos Anos Finais Ensino Fundamental e Ensino Médio (6º ano EF até 3ª Série EM);
# Nesse caso já estão somente esses valores

# 5. Padronizar CPFs para não ter diferenças de formatação entre as bases;
# Já está padronizado e só usa 1 base

# 6. Excluir duplicatas da combinação de componente + CPF
# A base tem alguns erros de mesmo estudante em séries e escolas diferentes
total_duplicadas = df.duplicated(subset=['CPF_Padronizado', 'COMPONENTE CURRICULAR']).sum()
total_duplicadas


df_sem_duplicata = df.drop_duplicates(subset=['CPF_Padronizado', 'COMPONENTE CURRICULAR'], keep='first', ignore_index=True)

# 7. Conseguir informações de necessidades especiais por CPF em algum relatório;
# Dataframe com os dados do relatório geral e padronização do CPF para cruzamento com o df_final
df_rel_geral = pd.read_excel(r"D:\Scripts_Python\FGV\RAPP_2026\Relatório Geral de Estudantes - Matrículas.xlsx", skiprows=2)

# Padronizar CPF no relatório geral
df_rel_geral['CPF_Padronizado'] = (
    df_rel_geral['CPF']
        .astype(str)
        .str.replace(r'\D', '', regex=True)
        .str.zfill(11)
        .str.replace(
            r'(\d{3})(\d{3})(\d{3})(\d{2})',
            r'\1.\2.\3-\4',
            regex=True
        )
)

# 9. Incluir informações de necessidades especiais e turma na base principal;
# Realiza merge dos dataframes
# Relatório Geral de Matrículas tem CPFs duplicados
total_duplicados = df_rel_geral['CPF_Padronizado'].duplicated().sum()
print(f"Existem {total_duplicados} CPFs repetidos.")

# Caso tenham vários valores para 'TUMA' e 'TIPO NECESSIDADE ESPECÍFICA INFORMADAS' para um mesmo CPF, vamos manter a moda.
# Função simples para retornar a moda (ou o primeiro valor caso haja empate)
def pegar_moda(x):
    m = x.mode()
    return m.iloc[0] if not m.empty else None

# Agrupamos pelo CPF e aplicamos a moda nas colunas desejadas
df_rel_limpo = df_rel_geral.groupby('CPF_Padronizado').agg({
    'TURMA': pegar_moda,
    'TIPO NECESSIDADE ESPECÍFICA INFORMADAS': pegar_moda
}).reset_index()


# Merge para incluir as informações de turma e necessidades especiais na base principal
df_final = df_sem_duplicata.merge(
    df_rel_limpo, 
    on='CPF_Padronizado', 
    how='left'
)

# Trocar valores NaN por "-" em 'TURMA' e 'TIPO NECESSIDADE ESPECÍFICA INFORMADAS'
df_final['TURMA'] = df_final['TURMA'].fillna('-')
df_final['TIPO NECESSIDADE ESPECÍFICA INFORMADAS'] = df_final['TIPO NECESSIDADE ESPECÍFICA INFORMADAS'].fillna('-')


# 8. Conseguir informações de turno por CPF em algum relatório;
# Turno será extraído de acordo com o código da turma.
#  Criar a coluna de turno a partir do código da turma, considerando:

# Garante que o df_final é um objeto independente (evita avisos do Pandas)
df_final = df_final.copy()

# Valor padrão
df_final['TURNO'] = "Não Identificado"

# Dicionário de mapeamento
mapeamento = {
    r'INT\d': 'Integral',
    r'M\d': 'Matutino',
    r'V\d': 'Vespertino',
    r'N\d': 'Noturno'
}

# Aplicando a lógica
for padrao, nome_turno in mapeamento.items():
    # Usamos o .loc para garantir que a alteração ocorra no DataFrame original
    mascara = df_final['TURMA'].astype(str).str.contains(padrao, case=True, na=False, regex=True)
    df_final.loc[mascara, 'TURNO'] = nome_turno


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


# Estudante por Série em cada DIREC
serie_direc = df_final.groupby(['DIREC', 'SÉRIE'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})
serie_direc


# Estudante por componente em cada DIREC
componente_direc = df_final.groupby(['DIREC', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})
componente_direc


# Estudante por Componente por Série
componente_serie = df_final.groupby(['SÉRIE', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})
componente_serie


# Estudante por Componente, por Série e por DIREC
componente_serie_direc = df_final.groupby(['DIREC', 'SÉRIE', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})
componente_serie_direc




# Juntar os valores por DIREC, Componente e Série e salvar em um arquivo Excel, com cada tabela em uma aba diferente
with pd.ExcelWriter(r"D:\Scripts_Python\FGV\RAPP_2026\20260508_analises_RAPP.xlsx") as writer:
    df_final.to_excel(writer, sheet_name='Base RAPP', index=False)
    direc.to_excel(writer, sheet_name='DIREC', index=False)
    componente.to_excel(writer, sheet_name='Componente', index=False)
    serie.to_excel(writer, sheet_name='Serie', index=False)
    serie_direc.to_excel(writer, sheet_name='Serie e DIREC', index=False)
    componente_direc.to_excel(writer, sheet_name='Componente e DIREC', index=False)
    componente_serie.to_excel(writer, sheet_name='Componente e Serie', index=False)
    componente_serie_direc.to_excel(writer, sheet_name='Componente, Serie e DIREC', index=False)





