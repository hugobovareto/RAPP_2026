'''
Ideia Geral do código para identificação dos estudantes em RAPP em 2026:
1. Utilizar Relatório Geral de Matrículas de 2025 para identificar estudantes em RAPP (SITUAÇÃO = 'PROGRESSÃO PARCIAL' e 'APENAS PROG. PARCIAL');
2. Padronizar CPFs na base de 2025 para cruzamento com a base de 2026;
3. Excluir valores indesejados de etapa;
4. Padronizar CPFs na base de 2026 para cruzamento com a base de 2025;
5. Padronizar CPFs nos Relatórios de Notas 2025 para cruzamento com a base de relatório geral de matrículas;
6. Utilizar Relatório de Notas 2025 para preencher os componentes curriculares reprovados dos estudantes identificados em RAPP.
7. Utilizar informações do Relatório Geral Matrículas (2026 e atual) para retirar estudantes não encontrados nos dados de 2026 ou com SITUAÇÃO = 'CANCELADO', 'DEIXOU DE FREQUENTAR', 'TRANSFERIDO';
8. Utilizar informações do Relatório Geral Matrículas (2026 e atual) para atualizar as informações de DIREC, Escola e turma (variáveis: DIREC; CÓDIGO INEP ESCOLA; ESCOLA; TURMA);
9. Excluir duplicatas da combinação de componente + CPF;
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

# 1. Utilizar Relatório Geral de Matrículas de 2025 para identificar estudantes em RAPP (SITUAÇÃO = 'PROGRESSÃO PARCIAL' e 'APENAS PROG. PARCIAL');
# Dataframe com os dados do Relatório Geral de Matrículas de 2025
df_geral_2025 = pd.read_excel(r"D:\Scripts_Python\FGV\RAPP_2026\2025_Relatório Geral de Estudantes - Matrículas.xlsx", skiprows=2)

# Excluir colunas que não são do interesse
df_geral_2025 = df_geral_2025.drop(columns=[
    'ANO',
    'PERÍODO',
    'MATRÍCULA',
    'DATA DE EFETIVAÇÃO DA MATRÍCULA',
    'NOME SOCIAL',
    'ID_PESSOA',
    'SEXO',
    'NOME DA FILIAÇÃO 1',
    'NOME DA FILIAÇÃO 2',
    'NOME DO RESPONSÁVEL MATRÍCULA',
    'DATA DE NASCIMENTO',
    'NÚMERO DE IDENTIDADE',
    'CÓDIGO NIS',
    'CÓDIGO SUS',
    'CÓDIGO INEP',
    'ENDEREÇO',
    'TELEFONES',
    'E-MAIL INSTITUCIONAL',
    'ID_UNIDADE_ESCOLA',
    'ID_SEGMENTO',
    'ID_SÉRIE',
    'ID_TURMA',
    'RESPÓNSÁVEL',
    'MATRICULA SERVIDOR',
    'LOGIN',
    'ID_USUÁRIO',
    'PARTICIPA PÉ DE MEIA',
    'RAÇA / COR',
    'UTILIZA TRANSPORTE ESCOLAR',
    'TIPO DE TRANSPORTE ESCOLAR',
    'TIPO'])


# Manter somente Estudantes em RAPP (SITUAÇÃO = 'PROGRESSÃO PARCIAL' e 'APENAS PROG. PARCIAL')
df_geral_2025 = df_geral_2025[df_geral_2025['SITUAÇÃO'].isin(['PROGRESSÃO PARCIAL', 'APENAS PROG. PARCIAL'])]


# 2. Padronizar CPFs na base de 2025 para cruzamento com a base de 2026;
# Padronizar CPF: manter apenas dígitos, completar com zeros à esquerda e formatar como XXX.XXX.XXX-XX
df_geral_2025['CPF_Padronizado'] = (
    df_geral_2025['CPF']
        .astype(str)
        .str.replace(r'\D', '', regex=True)   # remove tudo que não é dígito
        .str.zfill(11)                        # completa com zeros à esquerda
        .str.replace(
            r'(\d{3})(\d{3})(\d{3})(\d{2})',
            r'\1.\2.\3-\4',
            regex=True
        )
)

# Excluir coluna 'CPF' para só usar o 'CPF_Padronizado' e evitar confusão
df_geral_2025 = df_geral_2025.drop(columns=['CPF'])


# Verificar se tem CPFs duplicados (se sim, será mantido o com 'Data da Operação' mais recente)
total_duplicados = df_geral_2025['CPF_Padronizado'].duplicated().sum()
print(f"Total de registros duplicados: {total_duplicados}")

# Exclusão de CPFs duplicados, mantendo o valor mais recente baseado na 'DATA DA OPERAÇÃO'
# Converter a coluna de data para datetime
df_geral_2025['DATA DA OPERAÇÃO'] = pd.to_datetime(df_geral_2025['DATA DA OPERAÇÃO'], errors='coerce')

# Ordenar do mais recente para o mais antigo
df_geral_2025 = df_geral_2025.sort_values('DATA DA OPERAÇÃO', ascending=False)

# Remover CPFs duplicados, mantendo o registro mais recente
df_geral_2025 = df_geral_2025.drop_duplicates(subset='CPF_Padronizado', keep='first')


# 3. Excluir valores indesejados de etapa
# Como também contabiliza EPT, não será feito exclusão a partir do valor da SÉRIE
# Excluir etapas de ensino indesejadas (EJA e EJATEC)
# Lista de valores a removar para ETAPA DE ENSINO
etapas_remover = [
    'EJA - ENSINO MÉDIO',
    'ENSINO FUND. 2º SEGMENTO (ANOS FINAIS) - EDUCAÇÃO DE JOVENS E ADULTOS',
    'EJA - ENS FUNDAMENTAL - 2º SEGMENTO ANUAL',
    'CURSO TÉCNICO DE NÍVEL MÉDIO EM ADMINISTRAÇÃO NA FORMA ARTICULADA INTEGRADA A EDUCAÇÃO DE JOVENS E ADULTOS - EJATEC',
    'FUNDAMENTAL ANOS FINAIS - EJA',
    'CURSO TÉCNICO DE NÍVEL MÉDIO EM LOGÍSTICA NA FORMA ARTICULADA INTEGRADA A EDUCAÇÃO DE JOVENS E ADULTOS - EJATEC'
]

# Excluir os registros que possuem as etapas de ensino indesejadas
df_geral_2025 = df_geral_2025[~df_geral_2025['ETAPA DE ENSINO'].isin(etapas_remover)]


# Dataframe com os dados do Relatório Geral de Matrículas de 2026
df_geral_2026 = pd.read_excel(r"D:\Scripts_Python\FGV\RAPP_2026\2026_Relatório Geral de Estudantes - Matrículas.xlsx", skiprows=2)

# 4. Padronizar CPFs na base de 2026 para cruzamento com a base de 2025;
# Padronizar CPF: manter apenas dígitos, completar com zeros à esquerda e formatar como XXX.XXX.XXX-XX
df_geral_2026['CPF_Padronizado'] = (
    df_geral_2026['CPF']
        .astype(str)
        .str.replace(r'\D', '', regex=True)   # remove tudo que não é dígito
        .str.zfill(11)                        # completa com zeros à esquerda
        .str.replace(
            r'(\d{3})(\d{3})(\d{3})(\d{2})',
            r'\1.\2.\3-\4',
            regex=True
        )
)

# Excluir coluna 'CPF' para só usar o 'CPF_Padronizado' e evitar confusão
df_geral_2026 = df_geral_2026.drop(columns=['CPF'])

# Excluir colunas que não são do interesse
df_geral_2026 = df_geral_2026.drop(columns=[
    'ANO',
    'PERÍODO',
    'MATRÍCULA',
    'DATA DE EFETIVAÇÃO DA MATRÍCULA',
    'NOME SOCIAL',
    'ID_PESSOA',
    'SEXO',
    'NOME DA FILIAÇÃO 1',
    'NOME DA FILIAÇÃO 2',
    'NOME DO RESPONSÁVEL MATRÍCULA',
    'DATA DE NASCIMENTO',
    'NÚMERO DE IDENTIDADE',
    'CÓDIGO NIS',
    'CÓDIGO SUS',
    'CÓDIGO INEP',
    'ENDEREÇO',
    'TELEFONES',
    'E-MAIL INSTITUCIONAL',
    'ID_UNIDADE_ESCOLA',
    'ID_SEGMENTO',
    'ID_SÉRIE',
    'ID_TURMA',
    'RESPÓNSÁVEL',
    'MATRICULA SERVIDOR',
    'LOGIN',
    'ID_USUÁRIO',
    'PARTICIPA PÉ DE MEIA',
    'RAÇA / COR',
    'UTILIZA TRANSPORTE ESCOLAR',
    'TIPO DE TRANSPORTE ESCOLAR',
    'TIPO'])

# Verificar se tem CPFs duplicados (se sim, será mantido o com 'Data da Operação' mais recente)
total_duplicados = df_geral_2026['CPF_Padronizado'].duplicated().sum()
print(f"Total de registros duplicados: {total_duplicados}")

# Exclusão de CPFs duplicados, mantendo o valor mais recente baseado na 'DATA DA OPERAÇÃO'
# Converter a coluna de data para datetime
df_geral_2026['DATA DA OPERAÇÃO'] = pd.to_datetime(df_geral_2026['DATA DA OPERAÇÃO'], errors='coerce')  # Coerce para lidar com datas inválidas, se houver

# Ordenar do mais recente para o mais antigo
df_geral_2026 = df_geral_2026.sort_values('DATA DA OPERAÇÃO', ascending=False)

# Remover CPFs duplicados, mantendo o registro mais recente
df_geral_2026 = df_geral_2026.drop_duplicates(subset='CPF_Padronizado', keep='first')


# Relatórios de Notas 2025
# caminho da pasta onde estão os arquivos
pasta = r"C:\Users\hugob\Downloads\Notas"

# lista todos os arquivos .xlsx da pasta
arquivos = glob.glob(os.path.join(pasta, "*.xlsx"))

# lista para armazenar os dataframes
dfs = []

for arquivo in tqdm(arquivos, desc="Processando arquivos"):
    # lê cada arquivo, pulando as 2 primeiras linhas
    df_unico = pd.read_excel(arquivo, skiprows=2)
    dfs.append(df_unico)

# concatena todos em um único dataframe
df_notas = pd.concat(dfs, ignore_index=True)

# Manter somente os componentes curriculares reprovados (RESULTADO FINAL = 'REPROVADO')
df_notas = df_notas[df_notas['RESULTADO FINAL'] == 'REPROVADO']

# 5. Padronizar CPFs nos Relatórios de Notas 2025 para cruzamento com a base de relatório geral de matrículas;
# Padronizar CPF: manter apenas dígitos, completar com zeros à esquerda e formatar como XXX.XXX.XXX-XX
df_notas['CPF_Padronizado'] = (
    df_notas['CPF PESSOA']
        .astype(str)
        .str.replace(r'\D', '', regex=True)   # remove tudo que não é dígito
        .str.zfill(11)                        # completa com zeros à esquerda
        .str.replace(
            r'(\d{3})(\d{3})(\d{3})(\d{2})',
            r'\1.\2.\3-\4',
            regex=True
        )
)

# Excluir coluna 'CPF PESSOA' para só usar o 'CPF_Padronizado' e evitar confusão
df_notas = df_notas.drop(columns=['CPF PESSOA'])


# Excluir colunas que não são do interesse
df_notas = df_notas.drop(columns=[
    'ID DIREC',
    'ID MUNICÍPIO',
    'ID ETAPA ENSINO',
    'PERIODICIDADE ETAPA ENSINO',
    'ID SÉRIE',
    'ID TURMA',
    'ID PESSOA (PROFESSOR)',
    'MATRICULA (PROFESSOR)',
    'VÍNCULO',
    'NOME DO PROFESSOR',
    'DATA INÍCIO ALOCAÇÃO',
    'DATA FIM ALOCAÇÃO',
    'ID COMPONENTE CURRICULAR',
    'PERIODICIDADE COMPONENTE CURRICULAR',
    'ID PESSOA',
    'MATRÍCULA ESTUDANTE',
    'APROVEITAMENTO DE ESTUDO',
    'NOTA 1º BIMESTRE',
    'NOTA 2º BIMESTRE',
    'NOTA 3º BIMESTRE',
    'NOTA 4º BIMESTRE',
    'EXAME FINAL',
    'AVALIAÇÃO ESPECIAL'
])


# 6. Utilizar Relatório de Notas 2025 para preencher os componentes curriculares reprovados dos estudantes identificados em RAPP 2025.
# Selecionar somente acoluna de CPF e componente curricular para fazer o merge
colunas_notas = ['CPF_Padronizado', 'COMPONENTE CURRICULAR']

# Merge para incluir os componentes curriculares reprovados na base de estudantes em RAPP 2025
df_final = pd.merge(
    df_geral_2025, 
    df_notas[colunas_notas], 
    on='CPF_Padronizado', 
    how='left'
)

# Ordenar para componentes de cada estudante ficarem juntos
df_final = df_final.sort_values(by=['CPF_Padronizado', 'COMPONENTE CURRICULAR'])

# Verificar quantas linhas tem com Componente Curricular nulo (estudantes sem componente reprovado)
total_nulos = df_final['COMPONENTE CURRICULAR'].isna().sum()
print(f"Quantidade de linhas com Componente Curricular nulo: {total_nulos}")
# Tem estudantes que estão como Progressão Parcial no relatório geral de matrículas 2025, mas não apresentam componentes curriculares reprovados no relatório de notas

# Para os estudantes sem identificação do componente curricular (NaN), colocar "Não Identificado"
df_final['COMPONENTE CURRICULAR'] = df_final['COMPONENTE CURRICULAR'].fillna('Não Identificado')


# 7. Utilizar informações do Relatório Geral Matrículas (2026 e atual) para retirar estudantes não encontrados nos dados de 2026 ou com SITUAÇÃO = 'CANCELADO', 'DEIXOU DE FREQUENTAR', 'TRANSFERIDO';
# Criar uma lista de todos os CPFs que aparecem no relatório de 2026
cpfs_2026 = df_geral_2026['CPF_Padronizado'].unique()

# Estudantes que estão no df_final mas NÃO aparecem em 2026
estudantes_nao_encontrados = df_final[~df_final['CPF_Padronizado'].isin(cpfs_2026)]

# Estudantes que estão no df_final, aparecem em 2026, mas com situação de saída
situacoes_saida = ['CANCELADO', 'TRANSFERIDO', 'DEIXOU DE FREQUENTAR']

# CPFs em 2026 com status de saída
cpfs_saida_2026 = df_geral_2026[df_geral_2026['SITUAÇÃO'].isin(situacoes_saida)]['CPF_Padronizado']

# Estudantes do df_final em situação de saída
estudantes_saida = df_final[df_final['CPF_Padronizado'].isin(cpfs_saida_2026)]

# Concatenar os dois perfis
df_revisar = pd.concat([estudantes_nao_encontrados, estudantes_saida])

# Salvar em Excel
df_revisar.to_excel("estudantes_para_remover_ou_revisar.xlsx", index=False)


# Remover esses estudantes do df_final
df_final = df_final[~df_final['CPF_Padronizado'].isin(df_revisar['CPF_Padronizado'])]


# 8. Utilizar informações do Relatório Geral Matrículas (2026 e atual) para atualizar as informações de DIREC, Escola e turma (variáveis: DIREC; CÓDIGO INEP ESCOLA; ESCOLA; TURMA);
# Colunas para atualizar
colunas_atualizar = ['DIREC', 'CÓDIGO INEP ESCOLA', 'ESCOLA', 'TURMA']

# Dataframe de referência para atualização, com CPF como índice
df_ref_2026 = df_geral_2026.set_index('CPF_Padronizado')[colunas_atualizar]

# Atualização do df_final com os valores de 2026 para DIREC; CÓDIGO INEP ESCOLA; ESCOLA; TURMA
for coluna in colunas_atualizar:
    df_final[coluna] = df_final['CPF_Padronizado'].map(df_ref_2026[coluna]).fillna(df_final[coluna])


# Trocar nomenclatura das séries para padronizar:
mapeamento = {
    '6º Ano': '6º ANO',
    '7º Ano': '7º ANO',
    '8º Ano': '8º ANO',
    '9º Ano': '9º ANO'
}

df_final['SÉRIE'] = df_final['SÉRIE'].replace(mapeamento)


# 9. Excluir duplicatas da combinação de componente + CPF;
total_duplicados = df_final.duplicated(subset=['CPF_Padronizado', 'COMPONENTE CURRICULAR']).sum()
total_duplicados

# Remover asociação de CPF + Componente Curricular duplicada, mantendo o primeiro registro encontrado
df_final = df_final.drop_duplicates(subset=['CPF_Padronizado', 'COMPONENTE CURRICULAR'], keep='first')

# Criar a coluna 'ETAPA_RESUMIDA' a partir da SÉRIE
mapeamento_etapa = {
    '1ª SÉRIE': 'Ensino Médio',
    '2ª SÉRIE': 'Ensino Médio',
    '3ª SÉRIE': 'Ensino Médio',
    '1º PERÍODO': 'Ensino Médio',
    '2º PERÍODO': 'Ensino Médio',
    '3º PERÍODO': 'Ensino Médio',
    '1º SEMESTRE': 'Ensino Médio',
    '2º SEMESTRE': 'Ensino Médio',
    'TURMA II (8° E 9° ANOS)': 'Ens. Fund. - Anos Finais',
    '6º ANO': 'Ens. Fund. - Anos Finais',
    '7º ANO': 'Ens. Fund. - Anos Finais',
    '8º ANO': 'Ens. Fund. - Anos Finais',
    '9º ANO': 'Ens. Fund. - Anos Finais'
}

df_final['ETAPA_RESUMIDA'] = df_final['SÉRIE'].map(mapeamento_etapa)


# Criar coluna 'CATEGORIA_COMPONENTE' para diferenciar componentes da BNCC e EPT
# Componentes da BNCC
lista_bncc = [
    'Matemática',
    'Física',
    'Química',
    'Língua Portuguesa',
    'História', 
    'Biologia',
    'Geografia',
    'Língua Inglesa',
    'Sociologia', 
    'Língua Espanhola',
    'Arte',
    'Filosofia',
    'Educação Física',
    'Ciências', 
    'Educação Física',
    'Ensino Religioso'
]

# Criar a nova variável 'CATEGORIA_COMPONENTE' com base na lista de componentes da BNCC (caso não seja BNCC, classificado como EPT)
df_final['CATEGORIA_COMPONENTE'] = np.where(
    df_final['COMPONENTE CURRICULAR'].isin(lista_bncc), 
    'BNCC', 
    'EPT'
)


# 10. Fazer as segmentações e contagem de interesse: estudantes por DIREC; por componente; por turno; por Série; necessidades especiais etc.

#################################### ANÁLISES ####################################
# Estudantes por DIREC
# Contagem de CPF_Padronizado distinto por DIREC
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

# Estudantes por tipo de Necessidade Especial
necessidade_especial = df_final.groupby('TIPO NECESSIDADE ESPECÍFICA INFORMADAS')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})
necessidade_especial

# Estudante por Série em cada DIREC
serie_direc = df_final.groupby(['DIREC', 'SÉRIE'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})
serie_direc


# Estudante por componente em cada DIREC
componente_direc = df_final.groupby(['DIREC', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})
componente_direc


# Estudante por Componente por Série
componente_serie = df_final.groupby(['SÉRIE', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})
componente_serie


# Necessidade Especial por DIREC
necessidade_direc = df_final.groupby(['DIREC', 'TIPO NECESSIDADE ESPECÍFICA INFORMADAS'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})
necessidade_direc

# Estudante por Componente, por Série e por DIREC
componente_serie_direc = df_final.groupby(['DIREC', 'SÉRIE', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})
componente_serie_direc


# Juntar os valores por DIREC, Componente e Série e salvar em um arquivo Excel, com cada tabela em uma aba diferente
with pd.ExcelWriter(r"D:\Scripts_Python\FGV\RAPP_2026\20260513_analises_RAPP.xlsx") as writer:
    df_final.to_excel(writer, sheet_name='Base RAPP', index=False)
    direc.to_excel(writer, sheet_name='DIREC', index=False)
    componente.to_excel(writer, sheet_name='Componente', index=False)
    serie.to_excel(writer, sheet_name='Serie', index=False)
    necessidade_especial.to_excel(writer, sheet_name='Neces. Especial', index=False)
    serie_direc.to_excel(writer, sheet_name='Serie e DIREC', index=False)
    componente_direc.to_excel(writer, sheet_name='Componente e DIREC', index=False)
    componente_serie.to_excel(writer, sheet_name='Componente e Serie', index=False)
    necessidade_direc.to_excel(writer, sheet_name='Neces. Especial e DIREC', index=False)
    componente_serie_direc.to_excel(writer, sheet_name='Componente, Serie e DIREC', index=False)


####################################################################
# CATEGORIA_COMPONENTE = 'BNCC'
df_bncc = df_final[df_final['CATEGORIA_COMPONENTE'] == 'BNCC']


# Estudantes por DIREC
# Contagem de CPF_Padronizado distinto por DIREC
direc_bncc = df_bncc.groupby('DIREC')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})
direc_bncc


# Estudantes por Componente Curricular
# Contagem de CPF_Padronizado distinto por Componente Curricular
componente_bncc = df_bncc.groupby('COMPONENTE CURRICULAR')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})
componente_bncc


# Estudantes por Série
# Contagem de CPF_Padronizado distinto por Série
serie_bncc = df_bncc.groupby('SÉRIE')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})
serie_bncc


# Estudantes por tipo de Necessidade Especial
necessidade_especial_bncc = df_bncc.groupby('TIPO NECESSIDADE ESPECÍFICA INFORMADAS')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})
necessidade_especial_bncc

# Estudante por Série em cada DIREC
serie_direc_bncc = df_bncc.groupby(['DIREC', 'SÉRIE'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})
serie_direc_bncc


# Estudante por componente em cada DIREC
componente_direc_bncc = df_bncc.groupby(['DIREC', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})
componente_direc_bncc


# Estudante por Componente por Série
componente_serie_bncc = df_bncc.groupby(['SÉRIE', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})
componente_serie_bncc


# Necessidade Especial por DIREC
necessidade_direc_bncc = df_bncc.groupby(['DIREC', 'TIPO NECESSIDADE ESPECÍFICA INFORMADAS'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})
necessidade_direc_bncc

# Estudante por Componente, por Série e por DIREC
componente_serie_direc_bncc = df_bncc.groupby(['DIREC', 'SÉRIE', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})
componente_serie_direc_bncc


# Juntar os valores por DIREC, Componente e Série e salvar em um arquivo Excel, com cada tabela em uma aba diferente
with pd.ExcelWriter(r"D:\Scripts_Python\FGV\RAPP_2026\20260513_BNCC_analises_RAPP.xlsx") as writer:
    df_bncc.to_excel(writer, sheet_name='BNCC_Base RAPP', index=False)
    direc_bncc.to_excel(writer, sheet_name='BNCC_DIREC', index=False)
    componente_bncc.to_excel(writer, sheet_name='BNCC_Componente', index=False)
    serie_bncc.to_excel(writer, sheet_name='BNCC_Serie', index=False)
    necessidade_especial_bncc.to_excel(writer, sheet_name='BNCC_Neces. Especial', index=False)
    serie_direc_bncc.to_excel(writer, sheet_name='BNCC_Serie e DIREC', index=False)
    componente_direc_bncc.to_excel(writer, sheet_name='BNCC_Componente e DIREC', index=False)
    componente_serie_bncc.to_excel(writer, sheet_name='BNCC_Componente e Serie', index=False)
    necessidade_direc_bncc.to_excel(writer, sheet_name='BNCC_Neces. Especial e DIREC', index=False)
    componente_serie_direc_bncc.to_excel(writer, sheet_name='BNCC_Componente, Serie e DIREC', index=False)



######################################################################
# CATEGORIA_COMPONENTE = 'EPT'
df_ept = df_final[df_final['CATEGORIA_COMPONENTE'] == 'EPT']


# Estudantes por DIREC
# Contagem de CPF_Padronizado distinto por DIREC
direc_ept = df_ept.groupby('DIREC')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})
direc_ept


# Estudantes por Componente Curricular
# Contagem de CPF_Padronizado distinto por Componente Curricular
componente_ept = df_ept.groupby('COMPONENTE CURRICULAR')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})
componente_ept


# Estudantes por Série
# Contagem de CPF_Padronizado distinto por Série
serie_ept = df_ept.groupby('SÉRIE')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})
serie_ept


# Estudantes por tipo de Necessidade Especial
necessidade_especial_ept = df_ept.groupby('TIPO NECESSIDADE ESPECÍFICA INFORMADAS')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})
necessidade_especial_ept

# Estudante por Série em cada DIREC
serie_direc_ept = df_ept.groupby(['DIREC', 'SÉRIE'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})
serie_direc_ept


# Estudante por componente em cada DIREC
componente_direc_ept = df_ept.groupby(['DIREC', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})
componente_direc_ept


# Estudante por Componente por Série
componente_serie_ept = df_ept.groupby(['SÉRIE', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})
componente_serie_ept


# Necessidade Especial por DIREC
necessidade_direc_ept = df_ept.groupby(['DIREC', 'TIPO NECESSIDADE ESPECÍFICA INFORMADAS'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})
necessidade_direc_ept

# Estudante por Componente, por Série e por DIREC
componente_serie_direc_ept = df_ept.groupby(['DIREC', 'SÉRIE', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})
componente_serie_direc_ept


# Juntar os valores por DIREC, Componente e Série e salvar em um arquivo Excel, com cada tabela em uma aba diferente
with pd.ExcelWriter(r"D:\Scripts_Python\FGV\RAPP_2026\20260513_EPT_analises_RAPP.xlsx") as writer:
    df_ept.to_excel(writer, sheet_name='EPT_Base RAPP', index=False)
    direc_ept.to_excel(writer, sheet_name='EPT_DIREC', index=False)
    componente_ept.to_excel(writer, sheet_name='EPT_Componente', index=False)
    serie_ept.to_excel(writer, sheet_name='EPT_Serie', index=False)
    necessidade_especial_ept.to_excel(writer, sheet_name='EPT_Neces. Especial', index=False)
    serie_direc_ept.to_excel(writer, sheet_name='EPT_Serie e DIREC', index=False)
    componente_direc_ept.to_excel(writer, sheet_name='EPT_Componente e DIREC', index=False)
    componente_serie_ept.to_excel(writer, sheet_name='EPT_Componente e Serie', index=False)
    necessidade_direc_ept.to_excel(writer, sheet_name='EPT_Neces. Especial e DIREC', index=False)
    componente_serie_direc_ept.to_excel(writer, sheet_name='EPT_Componente, Serie e DIREC', index=False)


