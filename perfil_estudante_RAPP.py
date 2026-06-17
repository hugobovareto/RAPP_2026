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
from datetime import datetime
import matplotlib.pyplot as plt
import seaborn as sns


# Dados do Relatório Geral de Matrículas de 2026 para cruzamentos
df_geral_2026 = pd.read_excel(r"D:\Scripts_Python\FGV\RAPP_2026\20260610_2026_Relatório Geral de Estudantes - Matrículas.xlsx", skiprows=2)

# Dados do Relatório Geral de Matrículas de 2025 para cruzamentos
df_geral_2025 = pd.read_excel(r"D:\Scripts_Python\FGV\RAPP_2026\20260610_2025_Relatório Geral de Estudantes - Matrículas.xlsx", skiprows=2)


# Dados finais e tratados dos esudantes em RAPP para identificação do CPF
df_RAPP = pd.read_excel(r"D:\Scripts_Python\FGV\RAPP_2026\Resultados\20260608_RAPP\20260608_GERAL_analises_RAPP.xlsx", sheet_name='Base RAPP')

# Do df_RAPP pegar todos os CPF_Padronizados únicos
cpf_unicos_RAPP = df_RAPP['CPF_Padronizado'].unique()

# Deixar o df_RAPP com apenas 1 entrada por CPF_Padronizado para evitar que a quantidade de componentes dêem mais pesos para alguns estudantes.
df_rapp_unicos = df_RAPP.drop_duplicates(subset=['CPF_Padronizado'])


# CPF padronizados para df_geral_2026
# Padronizar CPF: manter apenas dígitos, completar com zeros à esquerda e formatar como XXX.XXX.XXX-XX
df_geral_2026['CPF_Padronizado'] = (
    df_geral_2026['CPF']
        .astype(str)
        .str.replace(r'\.0$', '', regex=True)   # REMOVE O ".0" DO FINAL (caso seja float)
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


# CPF padronizados para df_geral_2025
# Padronizar CPF: manter apenas dígitos, completar com zeros à esquerda e formatar como XXX.XXX.XXX-XX
df_geral_2025['CPF_Padronizado'] = (
    df_geral_2025['CPF']
        .astype(str)
        .str.replace(r'\.0$', '', regex=True)   # REMOVE O ".0" DO FINAL (caso seja float)
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


# Nos dados do Relatório Geral de Matrículas de 2025 e 2026, manter a informação somente com a 'DATA DA OPERAÇÃO' mais recente
# Exclusão de CPFs duplicados, mantendo o valor mais recente baseado na 'DATA DA OPERAÇÃO'
# 2025
# Converter a coluna de data para datetime
df_geral_2025['DATA DA OPERAÇÃO'] = pd.to_datetime(df_geral_2025['DATA DA OPERAÇÃO'], errors='coerce')

# Ordenar do mais recente para o mais antigo
df_geral_2025 = df_geral_2025.sort_values('DATA DA OPERAÇÃO', ascending=False)

# Remover CPFs duplicados, mantendo o registro mais recente
df_geral_2025 = df_geral_2025.drop_duplicates(subset='CPF_Padronizado', keep='first')


# 2026
# Converter a coluna de data para datetime
df_geral_2026['DATA DA OPERAÇÃO'] = pd.to_datetime(df_geral_2026['DATA DA OPERAÇÃO'], errors='coerce')

# Ordenar do mais recente para o mais antigo
df_geral_2026 = df_geral_2026.sort_values('DATA DA OPERAÇÃO', ascending=False)

# Remover CPFs duplicados, mantendo o registro mais recente
df_geral_2026 = df_geral_2026.drop_duplicates(subset='CPF_Padronizado', keep='first')


# Do df_geral_2025 preciso pegar as informações de gênero, raça/cor e idade dos estudantes de RAPP (df_RAPP)
# GÊNERO / SEXO
# Adicionar a coluna de SEXO vindo do relatório Geral de Matrículas (df_geral_2025) para o df_rapp_unicos, usando o CPF_Padronizado como chave de junção
df_rapp_unicos = df_rapp_unicos.merge(
    df_geral_2025[['CPF_Padronizado', 'SEXO']],  # Traz apenas a chave e a coluna desejada
    on='CPF_Padronizado',                       # Chave de ligação em ambos os DataFrames
    how='left'                                  # Mantém todos os registros de df_rapp_unicos
)

df_rapp_unicos['SEXO'].isna().sum()


# RAÇA / COR
# Adicionar a coluna de RAÇA / COR vindo do relatório Geral de Matrículas (df_geral_2025) para o df_rapp_unicos, usando o CPF_Padronizado como chave de junção
df_rapp_unicos = df_rapp_unicos.merge(
    df_geral_2025[['CPF_Padronizado', 'RAÇA / COR']],  # Traz apenas a chave e a coluna desejada
    on='CPF_Padronizado',                           # Chave de ligação em ambos os DataFrames
    how='left'                                      # Mantém todos os registros de df_rapp_unicos
)

df_rapp_unicos['RAÇA / COR'].isna().sum()


# IDADE
# Adicionar a coluna de DATA DE NASCIMENTO vindo do relatório Geral de Matrículas (df_geral_2025) para o df_rapp_unicos, usando o CPF_Padronizado como chave de junção
df_rapp_unicos = df_rapp_unicos.merge(
    df_geral_2025[['CPF_Padronizado', 'DATA DE NASCIMENTO']],  # Traz apenas a chave e a coluna desejada
    on='CPF_Padronizado',                           # Chave de ligação em ambos os DataFrames
    how='left'                                      # Mantém todos os registros de df_rapp_unicos
)

df_rapp_unicos['DATA DE NASCIMENTO'].isna().sum()


# Definir a data de referência para calcular a idade (10/06/2026)
data_referencia = datetime(2026, 6, 10)

# Converter a coluna de data para formato datetime
df_rapp_unicos['DATA DE NASCIMENTO'] = pd.to_datetime(df_rapp_unicos['DATA DE NASCIMENTO'], format='%d/%m/%Y %H:%M')

# Calcular as idades usando a data de referência
df_rapp_unicos['IDADE'] = (data_referencia - df_rapp_unicos['DATA DE NASCIMENTO']).dt.days // 365


#########################################################################
########## Sexo ##########
# Distribuição por sexo
df_dist_sexo = (
    df_rapp_unicos
        .groupby('SEXO')
        .agg(quantidade=('CPF_Padronizado', 'nunique'))
        .assign(
            percentual=lambda x: (x['quantidade'] / x['quantidade'].sum() * 100).round(2)
        )
        .reset_index()
)

df_dist_sexo


########## Raça / Cor ##########
# Distribuição por raça / cor
df_dist_raca = (
    df_rapp_unicos
        .groupby('RAÇA / COR')
        .agg(quantidade=('CPF_Padronizado', 'nunique'))
        .assign(
            percentual=lambda x: (x['quantidade'] / x['quantidade'].sum() * 100).round(2)
        )
        .reset_index()
)

df_dist_raca

########## Idade ##########
# Distribuição por idade
df_dist_idade = (
    df_rapp_unicos
        .groupby('IDADE')
        .agg(quantidade=('CPF_Padronizado', 'nunique'))
        .assign(
            percentual=lambda x: (x['quantidade'] / x['quantidade'].sum() * 100).round(2)
        )
        .reset_index()
)

df_dist_idade


########## Idade (agrupado de 20 ou +)##########
# Cria uma coluna temporária com os grupos de idade
df_rapp_unicos['IDADE_AGRUPADA'] = np.where(
    df_rapp_unicos['IDADE'] >= 20, 
    '20 ou +', 
    df_rapp_unicos['IDADE'].astype(str)
)

# Faz a distribuição baseada na nova coluna
df_dist_idade_agrupada = (
    df_rapp_unicos
        .groupby('IDADE_AGRUPADA')
        .agg(quantidade=('CPF_Padronizado', 'nunique'))
        .assign(
            percentual=lambda x: (x['quantidade'] / x['quantidade'].sum() * 100).round(2)
        )
        .reset_index()
)

# Remove a coluna temporária do DataFrame principal
df_rapp_unicos.drop(columns=['IDADE_AGRUPADA'], inplace=True)

df_dist_idade_agrupada


# Salvar o dataframe em Excel e as distribuições em diferentes abas
# Criando o arquivo Excel com múltiplas abas
with pd.ExcelWriter(r"D:\Scripts_Python\FGV\RAPP_2026\Perfil_Estudante\20260610_Perfil_Estudante_RAPP.xlsx") as writer:
    df_rapp_unicos.to_excel(writer, sheet_name='RAPP_Estudantes unicos', index=False)
    df_dist_sexo.to_excel(writer, sheet_name='Distr_Sexo', index=False)
    df_dist_raca.to_excel(writer, sheet_name='Distr_Raca', index=False)
    df_dist_idade.to_excel(writer, sheet_name='Distr_Idade', index=False)
    df_dist_idade_agrupada.to_excel(writer, sheet_name='Distr_Idade_Agrupada', index=False)




#####################################################################################################
# Total de estudantes da rede
df_geral_2026['CPF_Padronizado'].nunique()


# Filtra excluindo as situações desejadas e conta os CPFs únicos
cpfs_unicos_filtrados = df_geral_2026.loc[
    ~df_geral_2026['SITUAÇÃO'].isin(['TRANSFERIDO', 'CANCELADO', 'DEIXOU DE FREQUENTAR']),
    'CPF_Padronizado'
].nunique()

print(f"Quantidade de CPFs únicos ativos: {cpfs_unicos_filtrados}")




#####################################################################################################
# Gráfico de barras de distribuição de idade (df_dist_idade_agrupada)
# 1. Dados originais
idades_originais = df_dist_idade_agrupada['IDADE_AGRUPADA'].astype(str).tolist()
valores = df_dist_idade_agrupada['quantidade'].tolist()
percentuais = df_dist_idade_agrupada['percentual'].tolist()

# Adaptação do eixo X: adiciona " anos" se não for o "20 ou +"
idades_formatadas = [
    f"{id_} anos" if id_ != "20 ou +" else id_ 
    for id_ in idades_originais
]

# 2. Configurar a figura (Largura 20, Altura 5.5)
fig, ax = plt.subplots(figsize=(20, 5.5))

# 3. Desenhar as barras verticais com as idades formatadas
cor_verde = '#58a75b'
barras = ax.bar(idades_formatadas, valores, color=cor_verde, width=0.65, edgecolor='none')

# 4. Adicionar os rótulos (porcentagem inteira) no topo de cada barra
for barra, pct in zip(barras, percentuais):
    altura = barra.get_height()
    texto_pct = f'{int(round(pct))}%'
    
    ax.annotate(
        texto_pct,
        xy=(barra.get_x() + barra.get_width() / 2, altura),
        xytext=(0, 6),
        textcoords="offset points",
        ha='center', 
        va='bottom', 
        fontsize=14,
        weight='bold',
        color='#333333'
    )

# 5. Limpeza do layout
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.spines['left'].set_visible(False)
ax.spines['bottom'].set_color('#cccccc')
ax.spines['bottom'].set_linewidth(1.5)

# Ocultar o eixo Y
ax.get_yaxis().set_visible(False)

# Ajustar tamanho das fontes do eixo X (com os novos rótulos por extenso)
ax.tick_params(axis='x', colors='#333333', labelsize=13)

# 6. Salvamento e exibição
plt.savefig(
    r"D:\Scripts_Python\FGV\RAPP_2026\Perfil_Estudante\20260610_Distr_Idade_RAPP_Formatado.pdf", 
    bbox_inches='tight'
)


###############
# Gráfico de pizza de distribuição de raça/ cor (df_dist_raca)
# Gráfico de pizza de distribuição de raça/ cor (df_dist_raca)
# 1. Dados e ordenação (ajuda a organizar as fatias da maior para a menor)
df_ordenado = df_dist_raca.sort_values(by='quantidade', ascending=False)
labels = df_ordenado['RAÇA / COR'].tolist()
valores = df_ordenado['quantidade'].tolist()

# 2. Paleta de cores na ordem enviada (ajustada para as 6 categorias)
cores_hex = ['#aed13f', '#58a75b', '#fff9d9', '#f9d736', '#ffe6d9', '#cc8a42']

# 3. Configurar a figura (ajustada para formato mais largo/horizontal)
fig, ax = plt.subplots(figsize=(10, 5.5))

# Função para mostrar a porcentagem como número inteiro, ocultando fatias menores que 1%
def porcentagem_inteira(val):
    return f'{int(round(val))}%' if val >= 1 else ''

# 4. Desenhar o gráfico de pizza tradicional (sem o círculo central)
# Removi os labels textuais da borda da pizza (labels=None) para restarem apenas na legenda
wedges, texts, autotexts = ax.pie(
    valores, 
    labels=None,                  # Garante que não haverá texto duplicado nas fatias
    autopct=porcentagem_inteira,  # Rótulos internos com números inteiros
    startangle=90, 
    colors=cores_hex,
    pctdistance=0.75,             # Aproxima os números do centro para não vazarem
    textprops=dict(color="black", fontsize=10, weight="bold")
)

# 5. Adicionar a Legenda do lado esquerdo (Corrigido)
# Usamos 'center left' e ajustamos o bbox_to_anchor para ancorar fora do círculo, à esquerda
ax.legend(
    wedges, 
    labels,
    title="Raça / Cor",
    title_fontproperties={'weight': 'bold', 'size': 11},
    loc="center left",             # Posição de ancoragem correta
    bbox_to_anchor=(-0.25, 0.5),   # Desloca bem para a esquerda (-0.25) e centraliza verticalmente (0.5)
    fontsize=10,
    frameon=True                  # Adiciona uma borda leve na legenda para organização
)

# 6. Título e ajustes de layout
plt.title('Distribuição por Raça / Cor', fontsize=14, weight='bold', pad=10)

# Garante que a pizza continue redonda mesmo em uma figura retangular
ax.axis('equal')  

# Salvar gráfico em pdf
# O 'bbox_inches=tight' é fundamental aqui, pois ele força o PDF a expandir e incluir a legenda que colocamos para fora
plt.savefig(
    r"D:\Scripts_Python\FGV\RAPP_2026\Perfil_Estudante\20260610_Perfil_Estudante_RAPP.pdf", 
    bbox_inches='tight'
)






