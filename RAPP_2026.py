'''
Ideia Geral do código para identificação dos estudantes em RAPP em 2026:
1. Utilizar Relatório Geral de Matrículas de 2025 para identificar estudantes em RAPP (SITUAÇÃO = 'PROGRESSÃO PARCIAL' e 'APENAS PROG. PARCIAL');
2. Padronizar CPFs na base de 2025 para cruzamento com a base de 2026;
3. Excluir valores indesejados de etapa;
4. Padronizar CPFs na base de 2026 para cruzamento com a base de 2025;
5. Padronizar CPFs nos Relatórios de Notas 2025 para cruzamento com a base de relatório geral de matrículas;
6. Utilizar Relatório de Notas 2025 para preencher os componentes curriculares reprovados dos estudantes identificados em RAPP.
7. Utilizar informações do Relatório Geral Matrículas (2026 e atual) para atualizar as informações de DIREC, Escola e turma (variáveis: DIREC; CÓDIGO INEP ESCOLA; ESCOLA; TURMA);
8. Excluir duplicatas da combinação de componente + CPF;
9. Fazer as segmentações e contagem de interesse: estudantes por DIREC; por componente; por turno; por Série; necessidades especiais etc.
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


# 2. Padronizar CPFs na base de 2025 para cruzamento com a base de 2026;
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

# Lista de estudantes sem CPF
estudantes_sem_cpf_2025 = df_geral_2025[df_geral_2025['CPF_Padronizado'] == '000.000.000-00']

# Salvar em Excel
estudantes_sem_cpf_2025.to_excel(r"D:\Scripts_Python\FGV\RAPP_2026\2025_Matriculas_estudantes_sem_cpf.xlsx", index=False)

# Manter somente Estudantes em RAPP (SITUAÇÃO = 'PROGRESSÃO PARCIAL' e 'APENAS PROG. PARCIAL')
df_geral_2025 = df_geral_2025[df_geral_2025['SITUAÇÃO'].isin(['PROGRESSÃO PARCIAL', 'APENAS PROG. PARCIAL'])]


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

# Excluir CPF_Padronizado (000.000.000-00)
df_geral_2025 = df_geral_2025[df_geral_2025['CPF_Padronizado'] != '000.000.000-00']


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
    'CURSO TÉCNICO DE NÍVEL MÉDIO EM LOGÍSTICA NA FORMA ARTICULADA INTEGRADA A EDUCAÇÃO DE JOVENS E ADULTOS - EJATEC',
    'TÉCNICO DE NÍVEL MÉDIO EM ADMINISTRAÇÃO NA FORMA ARTICULADA INTEGRADA À EDUCAÇÃO DE JOVENS E ADULTOS - EJATEC'
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
        .str.replace(r'\.0$', '', regex=True)   # REMOVE O ".0" DO FINAL (caso seja float)
        .str.replace(r'\D', '', regex=True)   # remove tudo que não é dígito
        .str.zfill(11)                        # completa com zeros à esquerda
        .str.replace(
            r'(\d{3})(\d{3})(\d{3})(\d{2})',
            r'\1.\2.\3-\4',
            regex=True
        )
)

# Lista de estudantes sem CPF
estudantes_sem_cpf_2026 = df_geral_2026[df_geral_2026['CPF_Padronizado'] == '000.000.000-00']

# Salvar em Excel
estudantes_sem_cpf_2026.to_excel(r"D:\Scripts_Python\FGV\RAPP_2026\2026_Matriculas_estudantes_sem_cpf.xlsx", index=False)

# Excluir CPF_Padronizado (000.000.000-00)
df_geral_2026 = df_geral_2026[df_geral_2026['CPF_Padronizado'] != '000.000.000-00']

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

# 5. Padronizar CPFs nos Relatórios de Notas 2025 para cruzamento com a base de relatório geral de matrículas;
# Padronizar CPF: manter apenas dígitos, completar com zeros à esquerda e formatar como XXX.XXX.XXX-XX
df_notas['CPF_Padronizado'] = (
    df_notas['CPF PESSOA']
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

# Excluir componentes da base de dados que não devem reprovar
componentes_excluir = [
    'Projeto de Vida',
    'Recomposição das Aprendizagens em Língua Portuguesa e Matemática',
    'Recomposição de Aprendizagens - Matemática',
    'Recomposição de Aprendizagens - Língua Portuguesa',
    'Recomposição de Aprendizagens em Matemática',
    'Recomposição de Aprendizagens em Língua Portuguesa',
    'Orientação Acadêmica'
]

df_notas = df_notas[~df_notas['COMPONENTE CURRICULAR'].isin(componentes_excluir)]


# Lista de estudantes sem CPF
estudantes_sem_cpf_notas = df_notas[df_notas['CPF_Padronizado'] == '000.000.000-00']

# Salvar em Excel
estudantes_sem_cpf_notas.to_excel(r"D:\Scripts_Python\FGV\RAPP_2026\Notas_estudantes_sem_cpf.xlsx", index=False)

# Salvar em um único Excel os estudantes sem CPF das bases de Matrículas 2025, Matrículas 2026 e Notas 2025
with pd.ExcelWriter(r"D:\Scripts_Python\FGV\RAPP_2026\20260601_sem_CPF.xlsx") as writer:
    estudantes_sem_cpf_2025.to_excel(writer, sheet_name='Matriculas_2025', index=False)
    estudantes_sem_cpf_2026.to_excel(writer, sheet_name='Matriculas_2026', index=False)
    estudantes_sem_cpf_notas.to_excel(writer, sheet_name='Notas_2025', index=False)


# Criar base de dados do relatório de notas onde RESULTADO FINAL = ‘Matriculado'
df_notas_matriculados = df_notas[df_notas['RESULTADO FINAL'] == 'MATRICULADO']

# Salvar em Excel
df_notas_matriculados.to_excel("notas_matriculados.xlsx", index=False)

# Excluir CPF_Padronizado (000.000.000-00)
df_notas = df_notas[df_notas['CPF_Padronizado'] != '000.000.000-00']


# Manter somente os componentes curriculares diferentes de 'APROVADO' (RESULTADO FINAL = 'APROVADO') (o que inclui: Reprovado; Matriculado; Aproveitamento de Estudos)
df_notas = df_notas[df_notas['RESULTADO FINAL'] != 'APROVADO']


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


# 6. Utilizar Relatório de Notas 2025 para preencher os componentes curriculares não aprovados dos estudantes identificados em RAPP 2025.
# Selecionar somente a coluna de CPF e componente curricular para fazer o merge
colunas_notas = ['CPF_Padronizado', 'COMPONENTE CURRICULAR', 'RESULTADO FINAL']

# Merge para incluir os componentes curriculares não aprovados na base de estudantes em RAPP 2025
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
    '1º SEMESTRE': 'Ensino Médio',
    '2º SEMESTRE': 'Ensino Médio',
    'TURMA II (8° E 9° ANOS)': 'Ens. Fund. - Anos Finais',
    '6º ANO': 'Ens. Fund. - Anos Finais',
    '7º ANO': 'Ens. Fund. - Anos Finais',
    '8º ANO': 'Ens. Fund. - Anos Finais',
    '9º ANO': 'Ens. Fund. - Anos Finais'
}

df_final['ETAPA_RESUMIDA'] = df_final['SÉRIE'].map(mapeamento_etapa)


# Utilizar informações do Relatório Geral Matrículas (2026 e atual) para acrescentar a coluna 'SITUAÇÃO_2026';
df_final = pd.merge(
    df_final,
    df_geral_2026[['CPF_Padronizado', 'SITUAÇÃO']].rename(columns={'SITUAÇÃO': 'SITUAÇÃO_2026'}),
    on='CPF_Padronizado',
    how='left'
)


# Criar coluna 'CATEGORIA_COMPONENTE' para diferenciar componentes da BNCC, EPT e Específicos
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

# Componentes EPT
lista_ept = [
    'Informática Básica',
    'Eletricidade Básica',
    'Desenho Técnico',
    'Fundamentos de Lógica e Algoritmos',
    'Arquitetura e Organização de Computadores',
    'Teoria e Fundamentos da Administração',
    'Estatística',
    'Prevenção e Combate a Sinistros',
    'Gestão de Pessoas',
    'Gestão Pública e Terceiro Setor',
    'Controle Ambiental',
    'Lógica de Programação (Algoritmos)',
    'Metodologia do Trabalho Cientifico',
    'Direito Empresarial, Trabalhista e Tributário',
    'Instalações Elétricas de Baixa Tensão',
    'Eletrônica Aplicada',
    'Introdução à Segurança do Trabalho',
    'Fundamentos de Redes de Computadores',
    'Empreendedorismo',
    'Manutenção e Configuração de Computadores',
    'Matemática Financeira',
    'Noções de Eletrônica e Eletricidade',
    'Programação Estruturada',
    'Estudo dos Solos e Materiais de Construção',
    'Sociologia do Trabalho',
    'Contabilidade Geral',
    'Energia Eólica',
    'Estatística Aplicada à Segurança do Trabalho',
    'Programação WEB I e II',
    'Eletrônica Analógica',
    'Tipos de Energia Renovável',
    'Princípios da Agroecologia',
    'Programação Estruturada e Orientada a Objetos',
    'Gestão Organizacional',
    'Metodologia do Trabalho Científico',
    'Prevenção e Controle de Perdas',
    'Psicologia do Trabalho',
    'Energia Solar, Térmica e Fotovoltaica',
    'Fundamentos do Trabalho do Técnico em Redes de Computadores',
    'Gestão de Saúde e Segurança Ocupacional',
    'Primeiros Socorros',
    'Cabeamento Estruturado e Redes de Acesso',
    'Agricultura Familiar',
    'Banco de Dados',
    'Educação Ambiental e Eco Turismo',
    'Educação Digital',
    'Matemática Básica',
    'Planejamento Estratégico',
    'Segurança da Informação',
    'Sistemas Operacionais',
    'Fundamentos da Computação',
    'Fundamentos do Técnico em Redes de Computadores',
    'Gestão Financeira',
    'Instalações Prediais Hidrossanitárias e Elétricas',
    'Lógica de Programação - Algoritmos',
    'Matematica Básica',
    'Mineralogia',
    'Petrografia',
    'Redes de Computadores',
    'Tecnologia em Mídias Digitais',
    'Avaliação e Educação Nutricional',
    'Desenvolvimento Sustentável',
    'Gestão Ambiental',
    'Microbiologia Ambiental',
    'Sequenciamento da Produção',
    'Trabalho de Conclusão de Curso - TCC',
    'Cabeamento Estruturado e Redes de Acesso e Eletricidade Básica',
    'Gestão Organizacional e Segurança do Trabalho',
    'Hospitalidade e Meios de Hospedagem',
    'Impactos Ambientais',
    'Legislação Turistica',
    'Legislação e Segurança do Trabalho',
    'Química e Bioquímica dos Alimentos',
    'Sociedade, Cultura e Meio Ambiente',
    'Tecnologia de Implementação de Redes',
    'Anatomia e Fisiologia Humana - Noções Básicas',
    'Fitopatologia e Dietoterapia da Nutrição',
    'Geologia Geral',
    'História do RN Aplicada ao Turismo',
    'Marketing e Serviços',
    'Matemática Financeira e Estatística',
    'Microbiologia dos Alimentos',
    'Máquinas e Acionamentos Elétricos',
    'Orçamento e Estabilidade',
    'Química Orgânica',
    'Saúde Pública',
    'Tecnologia da Costura, do Enfesto e Corte',
    'Tecnologia da Modelagem',
    'Tecnologia de Cereais'
]

# Criar a nova variável 'CATEGORIA_COMPONENTE' com base na lista de componentes da BNCC e EPT (caso não seja nenhum dos dois, classificar como Específico)
# Definir regras/ condições
condicoes = [
    df_final['COMPONENTE CURRICULAR'].isin(lista_bncc),
    df_final['COMPONENTE CURRICULAR'].isin(lista_ept)
]

# Rótulos par as condições
rotulos = ['BNCC', 'EPT']

# Aplicar o select com o valor padrão para o que sobrar
df_final['CATEGORIA_COMPONENTE'] = np.select(condicoes, rotulos, default='Específico')


# Criar a nova variável 'CATEGORIA_NECESSIDADES ESPECIAIS' para agrupar os tipos de necessidades especiais
# Lista de tipos de necessidades especiais para cada categoria
categorias_necessidades_especiais = {
    'Altas habilidades/superdotação': 'Transtorno do Neurodesenvolvimento',
    'Transtorno de Déficit de Atenção e Hiperatividade (TDAH)': 'Transtorno do Neurodesenvolvimento',
    'Transtorno do Espectro Autista (TEA)': 'Transtorno do Neurodesenvolvimento',
    'Baixa visão': 'Deficiência Visual',
    'Baixa audição': 'Deficiência Auditiva',
    'Surdez': 'Deficiência Auditiva',
    'Discalculia': 'Transtorno do Neurodesenvolvimento',
    'Disgrafia': 'Transtorno do Neurodesenvolvimento',
    'Deficiência Auditiva': 'Deficiência Auditiva',
    'Deficiência física': 'Deficiência Física',
    'Deficiência intelectual': 'Deficiência Intelectual',
    'Deficiência múltipla': 'Deficiência Múltipla',
    'Deficiência intelectual': 'Deficiência Intelectual',
    'Visão monocular': 'Deficiência Visual',
    'Cegueira': 'Deficiência Visual',
    'Perda de visão periférica': 'Deficiência Visual',
    'Paraplegia': 'Deficiência Física',
    'Tetraplegia': 'Deficiência Física',
    'Hemiegia': 'Deficiência Física',
    'Paralisia cerebral': 'Deficiência Física',
    'Amputações e deformidades congênitas': 'Deficiência Física',
    'Síndrone de down': 'Deficiência Intelectual',
    'Síndromes genéticas': 'Deficiência Intelectual',
    'Dislexia': 'Transtorno do Neurodesenvolvimento',
    'Dislalia': 'Transtorno do Neurodesenvolvimento',
    'Disortografia': 'Transtorno do Neurodesenvolvimento'
}

def classificar_necessidade_especial(texto):
    # Trata valores nulos ou vazios
    if pd.isna(texto) or str(texto).strip() in ["", "-", "NÃO INFORMADO"]:
        return "Sem necessidade especial informada"
    
    texto = str(texto).strip()

    # Caso 1: A linha é exatamente igual a uma das chaves do mapeamento
    if texto in categorias_necessidades_especiais:
        return categorias_necessidades_especiais[texto]
    
    # Caso 2: Se não é uma opção única, verificar se existem múltiplas opções conhecidas dentro do texto
    # Contar quantas chaves do dicionário estão presentes na string
    opcoes_encontradas = [opcao for opcao in categorias_necessidades_especiais.keys() if opcao in texto]
    
    if len(opcoes_encontradas) > 1:
        return "Deficiência Múltipla"
    
    # Caso 3: Se encontrou apenas uma (mas talvez com algum caractere extra ou espaço diferente)
    if len(opcoes_encontradas) == 1:
        return categorias_necessidades_especiais[opcoes_encontradas[0]]

    return "Outra Categoria / Não Mapeado"

# 2. Aplicar a nova função
df_final['CATEGORIA_NECESSIDADES ESPECIAIS'] = df_final['TIPO NECESSIDADE ESPECÍFICA INFORMADAS'].apply(classificar_necessidade_especial)


# 10. Fazer as segmentações e contagem de interesse: estudantes por DIREC; por componente; por turno; por Série; necessidades especiais etc.
#################################### ANÁLISES ####################################
# Estudantes por DIREC
# Contagem de CPF_Padronizado distinto por DIREC
direc = df_final.groupby('DIREC')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudantes por Componente Curricular
# Contagem de CPF_Padronizado distinto por Componente Curricular
componente = df_final.groupby('COMPONENTE CURRICULAR')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudantes por Série
# Contagem de CPF_Padronizado distinto por Série
serie = df_final.groupby('SÉRIE')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudantes por Etapa de Ensino
# Contagem de CPF_Padronizado distinto por Etapa de Ensino
etapa = df_final.groupby('ETAPA_RESUMIDA')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudantes por tipo de Necessidade Especial
necessidade_especial = df_final.groupby('CATEGORIA_NECESSIDADES ESPECIAIS')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por Série em cada DIREC
serie_direc = df_final.groupby(['DIREC', 'SÉRIE'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por Etapa em cada DIREC
etapa_direc = df_final.groupby(['DIREC', 'ETAPA_RESUMIDA'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por componente em cada DIREC
componente_direc = df_final.groupby(['DIREC', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por Componente por Série
componente_serie = df_final.groupby(['SÉRIE', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por Componente por Etapa de Ensino
componente_etapa = df_final.groupby(['ETAPA_RESUMIDA', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Necessidade Especial por DIREC
necessidade_direc = df_final.groupby(['DIREC', 'CATEGORIA_NECESSIDADES ESPECIAIS'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Necessidade Especial por ETAPA
necessidade_etapa = df_final.groupby(['ETAPA_RESUMIDA', 'CATEGORIA_NECESSIDADES ESPECIAIS'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por Componente, por Série e por DIREC
componente_serie_direc = df_final.groupby(['DIREC', 'SÉRIE', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})


# Juntar os valores por DIREC, Componente e Série e salvar em um arquivo Excel, com cada tabela em uma aba diferente
with pd.ExcelWriter(r"D:\Scripts_Python\FGV\RAPP_2026\20260601_analises_RAPP.xlsx") as writer:
    df_final.to_excel(writer, sheet_name='Base RAPP', index=False)
    direc.to_excel(writer, sheet_name='DIREC', index=False)
    componente.to_excel(writer, sheet_name='Componente', index=False)
    serie.to_excel(writer, sheet_name='Serie', index=False)
    etapa.to_excel(writer, sheet_name='Etapa', index=False)
    necessidade_especial.to_excel(writer, sheet_name='Neces. Especial', index=False)
    serie_direc.to_excel(writer, sheet_name='Serie e DIREC', index=False)
    etapa_direc.to_excel(writer, sheet_name='Etapa e DIREC', index=False)
    componente_direc.to_excel(writer, sheet_name='Componente e DIREC', index=False)
    componente_serie.to_excel(writer, sheet_name='Componente e Serie', index=False)
    componente_etapa.to_excel(writer, sheet_name='Componente e Etapa', index=False)
    necessidade_direc.to_excel(writer, sheet_name='Neces. Especial e DIREC', index=False)
    necessidade_etapa.to_excel(writer, sheet_name='Neces. Especial e Etapa', index=False)
    componente_serie_direc.to_excel(writer, sheet_name='Componente, Serie e DIREC', index=False)


####################################################################
# CATEGORIA_COMPONENTE = 'BNCC'
df_bncc = df_final[df_final['CATEGORIA_COMPONENTE'] == 'BNCC']


# Estudantes por DIREC
# Contagem de CPF_Padronizado distinto por DIREC
direc_bncc = df_bncc.groupby('DIREC')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudantes por Componente Curricular
# Contagem de CPF_Padronizado distinto por Componente Curricular
componente_bncc = df_bncc.groupby('COMPONENTE CURRICULAR')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudantes por Série
# Contagem de CPF_Padronizado distinto por Série
serie_bncc = df_bncc.groupby('SÉRIE')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudantes por Etapa de Ensino
# Contagem de CPF_Padronizado distinto por Etapa de Ensino
etapa_bncc = df_bncc.groupby('ETAPA_RESUMIDA')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudantes por tipo de Necessidade Especial
necessidade_especial_bncc = df_bncc.groupby('CATEGORIA_NECESSIDADES ESPECIAIS')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por Série em cada DIREC
serie_direc_bncc = df_bncc.groupby(['DIREC', 'SÉRIE'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por Etapa em cada DIREC
etapa_direc_bncc = df_bncc.groupby(['DIREC', 'ETAPA_RESUMIDA'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por componente em cada DIREC
componente_direc_bncc = df_bncc.groupby(['DIREC', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por Componente por Série
componente_serie_bncc = df_bncc.groupby(['SÉRIE', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por Componente por Etapa de Ensino
componente_etapa_bncc = df_bncc.groupby(['ETAPA_RESUMIDA', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Necessidade Especial por DIREC
necessidade_direc_bncc = df_bncc.groupby(['DIREC', 'CATEGORIA_NECESSIDADES ESPECIAIS'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Necessidade Especial por ETAPA
necessidade_etapa_bncc = df_bncc.groupby(['ETAPA_RESUMIDA', 'CATEGORIA_NECESSIDADES ESPECIAIS'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por Componente, por Série e por DIREC
componente_serie_direc_bncc = df_bncc.groupby(['DIREC', 'SÉRIE', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})


# Juntar os valores por DIREC, Componente e Série e salvar em um arquivo Excel, com cada tabela em uma aba diferente
with pd.ExcelWriter(r"D:\Scripts_Python\FGV\RAPP_2026\20260601_BNCC_analises_RAPP.xlsx") as writer:
    df_bncc.to_excel(writer, sheet_name='BNCC_Base RAPP', index=False)
    direc_bncc.to_excel(writer, sheet_name='BNCC_DIREC', index=False)
    componente_bncc.to_excel(writer, sheet_name='BNCC_Componente', index=False)
    serie_bncc.to_excel(writer, sheet_name='BNCC_Serie', index=False)
    etapa_bncc.to_excel(writer, sheet_name='BNCC_Etapa', index=False)
    necessidade_especial_bncc.to_excel(writer, sheet_name='BNCC_Neces. Especial', index=False)
    serie_direc_bncc.to_excel(writer, sheet_name='BNCC_Serie e DIREC', index=False)
    etapa_direc_bncc.to_excel(writer, sheet_name='BNCC_Etapa e DIREC', index=False)
    componente_direc_bncc.to_excel(writer, sheet_name='BNCC_Componente e DIREC', index=False)
    componente_serie_bncc.to_excel(writer, sheet_name='BNCC_Componente e Serie', index=False)
    componente_etapa_bncc.to_excel(writer, sheet_name='BNCC_Componente e Etapa', index=False)
    necessidade_direc_bncc.to_excel(writer, sheet_name='BNCC_Neces. Especial e DIREC', index=False)
    necessidade_etapa_bncc.to_excel(writer, sheet_name='BNCC_Neces. Especial e Etapa', index=False)
    componente_serie_direc_bncc.to_excel(writer, sheet_name='BNCC_Componente, Serie e DIREC', index=False)


######################################################################
# CATEGORIA_COMPONENTE = 'EPT'
df_ept = df_final[df_final['CATEGORIA_COMPONENTE'] == 'EPT']


# Estudantes por DIREC
# Contagem de CPF_Padronizado distinto por DIREC
direc_ept = df_ept.groupby('DIREC')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudantes por Componente Curricular
# Contagem de CPF_Padronizado distinto por Componente Curricular
componente_ept = df_ept.groupby('COMPONENTE CURRICULAR')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudantes por Série
# Contagem de CPF_Padronizado distinto por Série
serie_ept = df_ept.groupby('SÉRIE')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudantes por Etapa de Ensino
# Contagem de CPF_Padronizado distinto por Série
etapa_ept = df_ept.groupby('ETAPA_RESUMIDA')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudantes por tipo de Necessidade Especial
necessidade_especial_ept = df_ept.groupby('CATEGORIA_NECESSIDADES ESPECIAIS')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por Série em cada DIREC
serie_direc_ept = df_ept.groupby(['DIREC', 'SÉRIE'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por Etapa em cada DIREC
etapa_direc_ept = df_ept.groupby(['DIREC', 'ETAPA_RESUMIDA'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por componente em cada DIREC
componente_direc_ept = df_ept.groupby(['DIREC', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por Componente por Série
componente_serie_ept = df_ept.groupby(['SÉRIE', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por Componente por Etapa de Ensino
componente_etapa_ept = df_ept.groupby(['ETAPA_RESUMIDA', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Necessidade Especial por DIREC
necessidade_direc_ept = df_ept.groupby(['DIREC', 'CATEGORIA_NECESSIDADES ESPECIAIS'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Necessidade Especial por ETAPA
necessidade_etapa_ept = df_ept.groupby(['ETAPA_RESUMIDA', 'CATEGORIA_NECESSIDADES ESPECIAIS'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por Componente, por Série e por DIREC
componente_serie_direc_ept = df_ept.groupby(['DIREC', 'SÉRIE', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})


# Juntar os valores por DIREC, Componente e Série e salvar em um arquivo Excel, com cada tabela em uma aba diferente
with pd.ExcelWriter(r"D:\Scripts_Python\FGV\RAPP_2026\20260601_EPT_analises_RAPP.xlsx") as writer:
    df_ept.to_excel(writer, sheet_name='EPT_Base RAPP', index=False)
    direc_ept.to_excel(writer, sheet_name='EPT_DIREC', index=False)
    componente_ept.to_excel(writer, sheet_name='EPT_Componente', index=False)
    serie_ept.to_excel(writer, sheet_name='EPT_Serie', index=False)
    etapa_ept.to_excel(writer, sheet_name='EPT_Etapa', index=False)
    necessidade_especial_ept.to_excel(writer, sheet_name='EPT_Neces. Especial', index=False)
    serie_direc_ept.to_excel(writer, sheet_name='EPT_Serie e DIREC', index=False)
    etapa_direc_ept.to_excel(writer, sheet_name='EPT_Etapa e DIREC', index=False)
    componente_direc_ept.to_excel(writer, sheet_name='EPT_Componente e DIREC', index=False)
    componente_serie_ept.to_excel(writer, sheet_name='EPT_Componente e Serie', index=False)
    componente_etapa_ept.to_excel(writer, sheet_name='EPT_Componente e Etapa', index=False)
    necessidade_direc_ept.to_excel(writer, sheet_name='EPT_Neces. Especial e DIREC', index=False)
    necessidade_etapa_ept.to_excel(writer, sheet_name='EPT_Neces. Especial e Etapa', index=False)
    componente_serie_direc_ept.to_excel(writer, sheet_name='EPT_Componente, Serie e DIREC', index=False)


######################################################################
# CATEGORIA_COMPONENTE = 'Específico'
df_esp = df_final[df_final['CATEGORIA_COMPONENTE'] == 'Específico']


# Estudantes por DIREC
# Contagem de CPF_Padronizado distinto por DIREC
direc_esp = df_esp.groupby('DIREC')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudantes por Componente Curricular
# Contagem de CPF_Padronizado distinto por Componente Curricular
componente_esp = df_esp.groupby('COMPONENTE CURRICULAR')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudantes por Série
# Contagem de CPF_Padronizado distinto por Série
serie_esp = df_esp.groupby('SÉRIE')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudantes por Etapa de Ensino
# Contagem de CPF_Padronizado distinto por Série
etapa_esp = df_esp.groupby('ETAPA_RESUMIDA')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudantes por tipo de Necessidade Especial
necessidade_especial_esp = df_esp.groupby('CATEGORIA_NECESSIDADES ESPECIAIS')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por Série em cada DIREC
serie_direc_esp = df_esp.groupby(['DIREC', 'SÉRIE'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por Etapa em cada DIREC
etapa_direc_esp = df_esp.groupby(['DIREC', 'ETAPA_RESUMIDA'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por componente em cada DIREC
componente_direc_esp = df_esp.groupby(['DIREC', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por Componente por Série
componente_serie_esp = df_esp.groupby(['SÉRIE', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por Componente por Etapa de Ensino
componente_etapa_esp = df_esp.groupby(['ETAPA_RESUMIDA', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Necessidade Especial por DIREC
necessidade_direc_esp = df_esp.groupby(['DIREC', 'CATEGORIA_NECESSIDADES ESPECIAIS'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Necessidade Especial por ETAPA
necessidade_etapa_esp = df_esp.groupby(['ETAPA_RESUMIDA', 'CATEGORIA_NECESSIDADES ESPECIAIS'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por Componente, por Série e por DIREC
componente_serie_direc_esp = df_esp.groupby(['DIREC', 'SÉRIE', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})


# Juntar os valores por DIREC, Componente e Série e salvar em um arquivo Excel, com cada tabela em uma aba diferente
with pd.ExcelWriter(r"D:\Scripts_Python\FGV\RAPP_2026\20260601_ESPECIFICO_analises_RAPP.xlsx") as writer:
    df_esp.to_excel(writer, sheet_name='ESP_Base RAPP', index=False)
    direc_esp.to_excel(writer, sheet_name='ESP_DIREC', index=False)
    componente_esp.to_excel(writer, sheet_name='ESP_Componente', index=False)
    serie_esp.to_excel(writer, sheet_name='ESP_Serie', index=False)
    etapa_esp.to_excel(writer, sheet_name='ESP_Etapa', index=False)
    necessidade_especial_esp.to_excel(writer, sheet_name='ESP_Neces. Especial', index=False)
    serie_direc_esp.to_excel(writer, sheet_name='ESP_Serie e DIREC', index=False)
    etapa_direc_esp.to_excel(writer, sheet_name='ESP_Etapa e DIREC', index=False)
    componente_direc_esp.to_excel(writer, sheet_name='ESP_Componente e DIREC', index=False)
    componente_serie_esp.to_excel(writer, sheet_name='ESP_Componente e Serie', index=False)
    componente_etapa_esp.to_excel(writer, sheet_name='ESP_Componente e Etapa', index=False)
    necessidade_direc_esp.to_excel(writer, sheet_name='ESP_Neces. Especial e DIREC', index=False)
    necessidade_etapa_esp.to_excel(writer, sheet_name='ESP_Neces. Especial e Etapa', index=False)
    componente_serie_direc_esp.to_excel(writer, sheet_name='ESP_Componente, Serie e DIREC', index=False)


######################################################################
# CATEGORIA_COMPONENTE = 'BNCC' + 'Específico'
df_bncc_esp = df_final[
    df_final['CATEGORIA_COMPONENTE'].isin(['BNCC', 'Específico'])]


# Estudantes por DIREC
# Contagem de CPF_Padronizado distinto por DIREC
direc_bncc_esp = df_bncc_esp.groupby('DIREC')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudantes por Componente Curricular
# Contagem de CPF_Padronizado distinto por Componente Curricular
componente_bncc_esp = df_bncc_esp.groupby('COMPONENTE CURRICULAR')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudantes por Série
# Contagem de CPF_Padronizado distinto por Série
serie_bncc_esp = df_bncc_esp.groupby('SÉRIE')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudantes por Etapa de Ensino
# Contagem de CPF_Padronizado distinto por Série
etapa_bncc_esp = df_bncc_esp.groupby('ETAPA_RESUMIDA')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudantes por tipo de Necessidade Especial
necessidade_especial_bncc_esp = df_bncc_esp.groupby('CATEGORIA_NECESSIDADES ESPECIAIS')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por Série em cada DIREC
serie_direc_bncc_esp = df_bncc_esp.groupby(['DIREC', 'SÉRIE'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por Etapa em cada DIREC
etapa_direc_bncc_esp = df_bncc_esp.groupby(['DIREC', 'ETAPA_RESUMIDA'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por componente em cada DIREC
componente_direc_bncc_esp = df_bncc_esp.groupby(['DIREC', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por Componente por Série
componente_serie_bncc_esp = df_bncc_esp.groupby(['SÉRIE', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por Componente por Etapa de Ensino
componente_etapa_bncc_esp = df_bncc_esp.groupby(['ETAPA_RESUMIDA', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Necessidade Especial por DIREC
necessidade_direc_bncc_esp = df_bncc_esp.groupby(['DIREC', 'CATEGORIA_NECESSIDADES ESPECIAIS'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Necessidade Especial por ETAPA
necessidade_etapa_bncc_esp = df_bncc_esp.groupby(['ETAPA_RESUMIDA', 'CATEGORIA_NECESSIDADES ESPECIAIS'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por Componente, por Série e por DIREC
componente_serie_direc_bncc_esp = df_bncc_esp.groupby(['DIREC', 'SÉRIE', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})


# Juntar os valores por DIREC, Componente e Série e salvar em um arquivo Excel, com cada tabela em uma aba diferente
with pd.ExcelWriter(r"D:\Scripts_Python\FGV\RAPP_2026\20260601_BNCC_ESPECIFICO_analises_RAPP.xlsx") as writer:
    df_bncc_esp.to_excel(writer, sheet_name='BNCC_ESP_Base RAPP', index=False)
    direc_bncc_esp.to_excel(writer, sheet_name='BNCC_ESP_DIREC', index=False)
    componente_bncc_esp.to_excel(writer, sheet_name='BNCC_ESP_Componente', index=False)
    serie_bncc_esp.to_excel(writer, sheet_name='BNCC_ESP_Serie', index=False)
    etapa_bncc_esp.to_excel(writer, sheet_name='BNCC_ESP_Etapa', index=False)
    necessidade_especial_bncc_esp.to_excel(writer, sheet_name='BNCC_ESP_Neces. Especial', index=False)
    serie_direc_bncc_esp.to_excel(writer, sheet_name='BNCC_ESP_Serie e DIREC', index=False)
    etapa_direc_bncc_esp.to_excel(writer, sheet_name='BNCC_ESP_Etapa e DIREC', index=False)
    componente_direc_bncc_esp.to_excel(writer, sheet_name='BNCC_ESP_Componente e DIREC', index=False)
    componente_serie_bncc_esp.to_excel(writer, sheet_name='BNCC_ESP_Componente e Serie', index=False)
    componente_etapa_bncc_esp.to_excel(writer, sheet_name='BNCC_ESP_Componente e Etapa', index=False)
    necessidade_direc_bncc_esp.to_excel(writer, sheet_name='BNCC_ESP_Neces Espec e DIREC', index=False)
    necessidade_etapa_bncc_esp.to_excel(writer, sheet_name='BNCC_ESP_Neces Espec e Etapa', index=False)
    componente_serie_direc_bncc_esp.to_excel(writer, sheet_name='BNCC_ESP_Comp, Serie e DIREC', index=False)


###########################################################################################################################################

# Endereços das escolas de estudantes em RAPP

# Importação da base tratada de estudantes em RAPP
df_rapp = pd.read_excel(r"D:\Scripts_Python\FGV\RAPP_2026\20260601_analises_RAPP.xlsx", sheet_name='Base RAPP')

# Importação de base de endereços
df_enderecos = pd.read_csv(r"D:\Scripts_Python\FGV\RAPP_2026\Enderecos_Escolas_RN.csv", sep=',')


# Adicionar a coluna de endereço na base de estudantes em RAPP, utilizando o código INEP da escola
# As colunas de código Inep no mesmo formato
df_rapp["CÓDIGO INEP ESCOLA"] = df_rapp["CÓDIGO INEP ESCOLA"].astype(str)
df_enderecos["Código INEP"] = df_enderecos["Código INEP"].astype(str)

# Só há um valor de código INEP para cada escola no df_enderecos
print(df_enderecos["Código INEP"].is_unique)

# Criar um mapeamento (dicionário/série) de Código INEP para Endereço
mapeamento_enderecos = df_enderecos.drop_duplicates(subset=["Código INEP"]).set_index("Código INEP")["Endereço"]

# 2. Cria a nova coluna direto no df_rapp
df_rapp["Endereço_Escolas"] = df_rapp["CÓDIGO INEP ESCOLA"].map(mapeamento_enderecos)

# Apenas estudantes do 6º ano do Ensino Fundamental
df_rapp_6ano = df_rapp[df_rapp['SÉRIE'] == '6º ANO']


# Salvar as bases com Endereço e somente 6º ano do Ensino Fundamental em arquivos Excel
with pd.ExcelWriter(r"D:\Scripts_Python\FGV\RAPP_2026\20260601_Enderecos_RAPP.xlsx") as writer: 
    df_rapp.to_excel(writer, sheet_name='Base RAPP', index=False)
    df_rapp_6ano.to_excel(writer, sheet_name='6ano', index=False)


# Salvar somente lista de endereços para as escolas
# Manter um dataframe apenas com os endereços de cada escola
colunas_selecionadas = [
    'DIREC',
    'CÓDIGO INEP ESCOLA',
    'ESCOLA',
    'MUNICÍPIO',
    'Endereço_Escolas'
]

# Filtrar o DataFrame e remover as duplicatas mantendo apenas a primeira ocorrência
df_rapp_escolas = df_rapp[colunas_selecionadas].drop_duplicates(subset=['CÓDIGO INEP ESCOLA'])

# Filtrar o dataframe de 6º ano e remover as duplicatas mantendo apenas a primeira ocorrência
df_rapp_6ano_escolas = df_rapp_6ano[colunas_selecionadas].drop_duplicates(subset=['CÓDIGO INEP ESCOLA'])


# Salvar em Excel os endereços das escolas do RAPP e dos estudantes do 6º ano
with pd.ExcelWriter(r"D:\Scripts_Python\FGV\RAPP_2026\20260601_Enderecos_RAPP_Escolas.xlsx") as writer:
    df_rapp_escolas.to_excel(writer, sheet_name='Endereços Escolas RAPP', index=False)
    df_rapp_6ano_escolas.to_excel(writer, sheet_name='Endereços Escolas 6º ano', index=False)







##########################################################################################################################################
'''
Base de estudantes em RAPP a partir de dados passados diretamente pela equipe do GPD da SEEC. 
A extração e tratamento foi feita pelo GPD da SEEC.

Procedimento:
1. Fonte inicial: dados postos > aba; ‘RAPP REGULAR E TEC’
2. Cruzar com a aba ‘REULTADOS’ para conseguir o ‘tipo de necessidade específica informada’
3. 'Cruzar com a aba 'COMPONENTES' para conseguir todos os componentes reprovados por estudante'
3. Criar a ‘CATEGORIA_NECESSIDADES ESPECIAIS’, ‘ETAPA_RESUMIDA’ e ‘CATEGORIA_COMPONENTE’ de acordo com o que tinha feito anteriormente.
4. fazer todas as segmentações feitas anteriormente;
5. Gerar 1 planilha para geral (todos os estudantes) e 1 para cada DIREC (17 planilhas final).

'''
# Utilizar as abas 'RAPP REGULAR E TEC' e 'RESULTADOS' como fontes de informação para análise
df_rapp = pd.read_excel(r"D:\Scripts_Python\FGV\RAPP_2026\20260603_Firmino_RAPP 2026.xlsx", sheet_name='RAPP REGULAR E TEC')
df_necessidades = pd.read_excel(r"D:\Scripts_Python\FGV\RAPP_2026\20260603_Firmino_RAPP 2026.xlsx", sheet_name='RESULTADOS')
df_componentes = pd.read_excel(r"D:\Scripts_Python\FGV\RAPP_2026\20260603_Firmino_RAPP 2026.xlsx", sheet_name='COMPONENTES')


# Padronizar o CPF dos 2 dataframes para cruzamento
# Padronizar CPF: manter apenas dígitos, completar com zeros à esquerda e formatar como XXX.XXX.XXX-XX
df_rapp['CPF_Padronizado'] = (
    df_rapp['CPF PESSOA']
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

df_necessidades['CPF_Padronizado'] = (
    df_necessidades['CPF']
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

df_componentes['CPF_Padronizado'] = (
    df_componentes['CPF PESSOA']
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


# Acrescentar no df_rapp a coluna 'TIPO NECESSIDADE ESPECÍFICA INFORMADAS', cruzando com o df_necessidades a partir do CPF_Padronizado
# Filtra e remove CPFs duplicados no df_necessidades, mantendo apenas a primeira aparição
df_necessidades_filtrado = df_necessidades[
    ["CPF_Padronizado", "TIPO NECESSIDADE ESPECÍFICA INFORMADAS"]
].drop_duplicates(subset=["CPF_Padronizado"])

# Merge para acrescentar a coluna de tipo de necessidade específica informada no df_rapp
df_rapp = df_rapp.merge(
    df_necessidades_filtrado,
    on='CPF_Padronizado',
    how='left'
)


# Excluir a coluna 'COMPONENTE CURRICULAR' do df_rapp já que essa não apresenta informação relevante e substituir por uma nova coluna 'COMPONENTE CURRICULAR' vindo da aba 'COMPONENTES'
df_rapp = df_rapp.drop(columns=['COMPONENTE CURRICULAR'])


# Acrescentar no df_rapp a coluna 'COMPONENTE CURRICULAR', cruzando com o df_componentes a partir do CPF_Padronizado
df_rapp = df_rapp.merge(
    df_componentes[['CPF_Padronizado', 'COMPONENTE CURRICULAR']], 
    on='CPF_Padronizado', 
    how='left'
)

# Trocar nomenclatura das séries para padronizar:
mapeamento = {
    '6º Ano': '6º ANO',
    '7º Ano': '7º ANO',
    '8º Ano': '8º ANO',
    '9º Ano': '9º ANO'
}

df_rapp['SÉRIE'] = df_rapp['SÉRIE'].replace(mapeamento)


# Criar a coluna 'ETAPA_RESUMIDA' a partir da SÉRIE
mapeamento_etapa = {
    '1ª SÉRIE': 'Ensino Médio',
    '2ª SÉRIE': 'Ensino Médio',
    '3ª SÉRIE': 'Ensino Médio',
    '1º SEMESTRE': 'Ensino Médio',
    '2º SEMESTRE': 'Ensino Médio',
    '3° PERÍODO': 'Ensino Médio',
    'TURMA II (8° E 9° ANOS)': 'Ens. Fund. - Anos Finais',
    '6º ANO': 'Ens. Fund. - Anos Finais',
    '7º ANO': 'Ens. Fund. - Anos Finais',
    '8º ANO': 'Ens. Fund. - Anos Finais',
    '9º ANO': 'Ens. Fund. - Anos Finais'
}

df_rapp['ETAPA_RESUMIDA'] = df_rapp['SÉRIE'].map(mapeamento_etapa)


# Criar coluna 'CATEGORIA_COMPONENTE' para diferenciar componentes da BNCC, EPT e Específicos
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

# Componentes EPT
lista_ept = [
    'Informática Básica',
    'Eletricidade Básica',
    'Desenho Técnico',
    'Fundamentos de Lógica e Algoritmos',
    'Arquitetura e Organização de Computadores',
    'Teoria e Fundamentos da Administração',
    'Estatística',
    'Prevenção e Combate a Sinistros',
    'Gestão de Pessoas',
    'Gestão Pública e Terceiro Setor',
    'Controle Ambiental',
    'Lógica de Programação (Algoritmos)',
    'Metodologia do Trabalho Cientifico',
    'Direito Empresarial, Trabalhista e Tributário',
    'Instalações Elétricas de Baixa Tensão',
    'Eletrônica Aplicada',
    'Introdução à Segurança do Trabalho',
    'Fundamentos de Redes de Computadores',
    'Empreendedorismo',
    'Manutenção e Configuração de Computadores',
    'Matemática Financeira',
    'Noções de Eletrônica e Eletricidade',
    'Programação Estruturada',
    'Estudo dos Solos e Materiais de Construção',
    'Sociologia do Trabalho',
    'Contabilidade Geral',
    'Energia Eólica',
    'Estatística Aplicada à Segurança do Trabalho',
    'Programação WEB I e II',
    'Eletrônica Analógica',
    'Tipos de Energia Renovável',
    'Princípios da Agroecologia',
    'Programação Estruturada e Orientada a Objetos',
    'Gestão Organizacional',
    'Metodologia do Trabalho Científico',
    'Prevenção e Controle de Perdas',
    'Psicologia do Trabalho',
    'Energia Solar, Térmica e Fotovoltaica',
    'Fundamentos do Trabalho do Técnico em Redes de Computadores',
    'Gestão de Saúde e Segurança Ocupacional',
    'Primeiros Socorros',
    'Cabeamento Estruturado e Redes de Acesso',
    'Agricultura Familiar',
    'Banco de Dados',
    'Educação Ambiental e Eco Turismo',
    'Educação Digital',
    'Matemática Básica',
    'Planejamento Estratégico',
    'Segurança da Informação',
    'Sistemas Operacionais',
    'Fundamentos da Computação',
    'Fundamentos do Técnico em Redes de Computadores',
    'Gestão Financeira',
    'Instalações Prediais Hidrossanitárias e Elétricas',
    'Lógica de Programação - Algoritmos',
    'Matematica Básica',
    'Mineralogia',
    'Petrografia',
    'Redes de Computadores',
    'Tecnologia em Mídias Digitais',
    'Avaliação e Educação Nutricional',
    'Desenvolvimento Sustentável',
    'Gestão Ambiental',
    'Microbiologia Ambiental',
    'Sequenciamento da Produção',
    'Trabalho de Conclusão de Curso - TCC',
    'Cabeamento Estruturado e Redes de Acesso e Eletricidade Básica',
    'Gestão Organizacional e Segurança do Trabalho',
    'Hospitalidade e Meios de Hospedagem',
    'Impactos Ambientais',
    'Legislação Turistica',
    'Legislação e Segurança do Trabalho',
    'Química e Bioquímica dos Alimentos',
    'Sociedade, Cultura e Meio Ambiente',
    'Tecnologia de Implementação de Redes',
    'Anatomia e Fisiologia Humana - Noções Básicas',
    'Fitopatologia e Dietoterapia da Nutrição',
    'Geologia Geral',
    'História do RN Aplicada ao Turismo',
    'Marketing e Serviços',
    'Matemática Financeira e Estatística',
    'Microbiologia dos Alimentos',
    'Máquinas e Acionamentos Elétricos',
    'Orçamento e Estabilidade',
    'Química Orgânica',
    'Saúde Pública',
    'Tecnologia da Costura, do Enfesto e Corte',
    'Tecnologia da Modelagem',
    'Tecnologia de Cereais'
]

# Criar a nova variável 'CATEGORIA_COMPONENTE' com base na lista de componentes da BNCC e EPT (caso não seja nenhum dos dois, classificar como Específico)
# Definir regras/ condições
condicoes = [
    df_rapp['COMPONENTE CURRICULAR'].isin(lista_bncc),
    df_rapp['COMPONENTE CURRICULAR'].isin(lista_ept)
]

# Rótulos par as condições
rotulos = ['BNCC', 'EPT']

# Aplicar o select com o valor padrão para o que sobrar
df_rapp['CATEGORIA_COMPONENTE'] = np.select(condicoes, rotulos, default='Específico')


# Criar a nova variável 'CATEGORIA_NECESSIDADES ESPECIAIS' para agrupar os tipos de necessidades especiais
# Lista de tipos de necessidades especiais para cada categoria
categorias_necessidades_especiais = {
    'Altas habilidades/superdotação': 'Transtorno do Neurodesenvolvimento',
    'Transtorno de Déficit de Atenção e Hiperatividade (TDAH)': 'Transtorno do Neurodesenvolvimento',
    'Transtorno do Espectro Autista (TEA)': 'Transtorno do Neurodesenvolvimento',
    'Baixa visão': 'Deficiência Visual',
    'Baixa audição': 'Deficiência Auditiva',
    'Surdez': 'Deficiência Auditiva',
    'Discalculia': 'Transtorno do Neurodesenvolvimento',
    'Disgrafia': 'Transtorno do Neurodesenvolvimento',
    'Deficiência Auditiva': 'Deficiência Auditiva',
    'Deficiência física': 'Deficiência Física',
    'Deficiência intelectual': 'Deficiência Intelectual',
    'Deficiência múltipla': 'Deficiência Múltipla',
    'Deficiência intelectual': 'Deficiência Intelectual',
    'Visão monocular': 'Deficiência Visual',
    'Cegueira': 'Deficiência Visual',
    'Perda de visão periférica': 'Deficiência Visual',
    'Paraplegia': 'Deficiência Física',
    'Tetraplegia': 'Deficiência Física',
    'Hemiegia': 'Deficiência Física',
    'Paralisia cerebral': 'Deficiência Física',
    'Amputações e deformidades congênitas': 'Deficiência Física',
    'Síndrone de down': 'Deficiência Intelectual',
    'Síndromes genéticas': 'Deficiência Intelectual',
    'Dislexia': 'Transtorno do Neurodesenvolvimento',
    'Dislalia': 'Transtorno do Neurodesenvolvimento',
    'Disortografia': 'Transtorno do Neurodesenvolvimento'
}

def classificar_necessidade_especial(texto):
    # Trata valores nulos ou vazios
    if pd.isna(texto) or str(texto).strip() in ["", "-", "NÃO INFORMADO"]:
        return "Sem necessidade especial informada"
    
    texto = str(texto).strip()

    # Caso 1: A linha é exatamente igual a uma das chaves do mapeamento
    if texto in categorias_necessidades_especiais:
        return categorias_necessidades_especiais[texto]
    
    # Caso 2: Se não é uma opção única, verificar se existem múltiplas opções conhecidas dentro do texto
    # Contar quantas chaves do dicionário estão presentes na string
    opcoes_encontradas = [opcao for opcao in categorias_necessidades_especiais.keys() if opcao in texto]
    
    if len(opcoes_encontradas) > 1:
        return "Deficiência Múltipla"
    
    # Caso 3: Se encontrou apenas uma (mas talvez com algum caractere extra ou espaço diferente)
    if len(opcoes_encontradas) == 1:
        return categorias_necessidades_especiais[opcoes_encontradas[0]]

    return "Outra Categoria / Não Mapeado"

# 2. Aplicar a nova função
df_rapp['CATEGORIA_NECESSIDADES ESPECIAIS'] = df_rapp['TIPO NECESSIDADE ESPECÍFICA INFORMADAS'].apply(classificar_necessidade_especial)


# Fazer as segmentações e contagem de interesse: estudantes por DIREC; por componente; por turno; por Série; necessidades especiais etc.
#################################### ANÁLISES ####################################
# Estudantes por DIREC
# Contagem de CPF_Padronizado distinto por DIREC
direc = df_rapp.groupby('DIREC')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudantes por Componente Curricular
# Contagem de CPF_Padronizado distinto por Componente Curricular
componente = df_rapp.groupby('COMPONENTE CURRICULAR')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudantes por Série
# Contagem de CPF_Padronizado distinto por Série
serie = df_rapp.groupby('SÉRIE')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudantes por Etapa de Ensino
# Contagem de CPF_Padronizado distinto por Etapa de Ensino
etapa = df_rapp.groupby('ETAPA_RESUMIDA')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudantes por tipo de Necessidade Especial
necessidade_especial = df_rapp.groupby('CATEGORIA_NECESSIDADES ESPECIAIS')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por Série em cada DIREC
serie_direc = df_rapp.groupby(['DIREC', 'SÉRIE'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por Etapa em cada DIREC
etapa_direc = df_rapp.groupby(['DIREC', 'ETAPA_RESUMIDA'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por componente em cada DIREC
componente_direc = df_rapp.groupby(['DIREC', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por Componente por Série
componente_serie = df_rapp.groupby(['SÉRIE', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por Componente por Etapa de Ensino
componente_etapa = df_rapp.groupby(['ETAPA_RESUMIDA', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Necessidade Especial por DIREC
necessidade_direc = df_rapp.groupby(['DIREC', 'CATEGORIA_NECESSIDADES ESPECIAIS'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Necessidade Especial por ETAPA
necessidade_etapa = df_rapp.groupby(['ETAPA_RESUMIDA', 'CATEGORIA_NECESSIDADES ESPECIAIS'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por Componente, por Série e por DIREC
componente_serie_direc = df_rapp.groupby(['DIREC', 'SÉRIE', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})


# Caminho da pasta onde os arquivos serão salvos
pasta_destino = r"D:\Scripts_Python\FGV\RAPP_2026\20260608_RAPP"


# Salva o arquivo GERAL
caminho_geral = os.path.join(pasta_destino, "20260608_GERAL_analises_RAPP.xlsx")

with pd.ExcelWriter(caminho_geral) as writer:
    df_rapp.to_excel(writer, sheet_name='Base RAPP', index=False)
    direc.to_excel(writer, sheet_name='DIREC', index=False)
    componente.to_excel(writer, sheet_name='Componente', index=False)
    serie.to_excel(writer, sheet_name='Serie', index=False)
    etapa.to_excel(writer, sheet_name='Etapa', index=False)
    necessidade_especial.to_excel(writer, sheet_name='Neces. Especial', index=False)
    serie_direc.to_excel(writer, sheet_name='Serie e DIREC', index=False)
    etapa_direc.to_excel(writer, sheet_name='Etapa e DIREC', index=False)
    componente_direc.to_excel(writer, sheet_name='Componente e DIREC', index=False)
    componente_serie.to_excel(writer, sheet_name='Componente e Serie', index=False)
    componente_etapa.to_excel(writer, sheet_name='Componente e Etapa', index=False)
    necessidade_direc.to_excel(writer, sheet_name='Neces. Especial e DIREC', index=False)
    necessidade_etapa.to_excel(writer, sheet_name='Neces. Especial e Etapa', index=False)
    componente_serie_direc.to_excel(writer, sheet_name='Componente, Serie e DIREC', index=False)


#################################################################
# GERAR 1 PLANILHA PARA CADA DIREC
print("\nIniciando a criação das planilhas por DIREC...")

# Obtém a lista de DIRECs únicas diretamente da coluna para evitar erros de digitação
lista_direcs = df_rapp["DIREC"].dropna().unique()

for d in lista_direcs:
    print(f" -> Processando: {d}")

    # Formata o nome do arquivo (ex: '01ª DIREC - NATAL' vira '01_DIREC_NATAL')
    nome_limpo = (
        str(d).replace("ª", "").replace(" - ", "_").replace(" ", "_")
    )
    caminho_direc = os.path.join(
        pasta_destino, f"20260608_{nome_limpo}_analises_RAPP.xlsx"
    )

    # Passo chave: Filtra a base principal APENAS para a DIREC do loop atual
    df_rapp_filtrado = df_rapp[df_rapp["DIREC"] == d]

    # Recalcula todas as tabelas usando o DataFrame filtrado
    direc_f = (
        df_rapp_filtrado.groupby("DIREC")["CPF_Padronizado"]
        .nunique()
        .reset_index()
        .rename(
            columns={"CPF_Padronizado": "Quantidade de Estudantes Distintos"}
        )
    )
    componente_f = (
        df_rapp_filtrado.groupby("COMPONENTE CURRICULAR")["CPF_Padronizado"]
        .nunique()
        .reset_index()
        .rename(
            columns={"CPF_Padronizado": "Quantidade de Estudantes Distintos"}
        )
    )
    serie_f = (
        df_rapp_filtrado.groupby("SÉRIE")["CPF_Padronizado"]
        .nunique()
        .reset_index()
        .rename(
            columns={"CPF_Padronizado": "Quantidade de Estudantes Distintos"}
        )
    )
    etapa_f = (
        df_rapp_filtrado.groupby("ETAPA_RESUMIDA")["CPF_Padronizado"]
        .nunique()
        .reset_index()
        .rename(
            columns={"CPF_Padronizado": "Quantidade de Estudantes Distintos"}
        )
    )
    necessidade_especial_f = (
        df_rapp_filtrado.groupby("CATEGORIA_NECESSIDADES ESPECIAIS")[
            "CPF_Padronizado"
        ]
        .nunique()
        .reset_index()
        .rename(
            columns={"CPF_Padronizado": "Quantidade de Estudantes Distintos"}
        )
    )
    serie_direc_f = (
        df_rapp_filtrado.groupby(["DIREC", "SÉRIE"])["CPF_Padronizado"]
        .nunique()
        .reset_index()
        .rename(
            columns={"CPF_Padronizado": "Quantidade de Estudantes Distintos"}
        )
    )
    etapa_direc_f = (
        df_rapp_filtrado.groupby(["DIREC", "ETAPA_RESUMIDA"])[
            "CPF_Padronizado"
        ]
        .nunique()
        .reset_index()
        .rename(
            columns={"CPF_Padronizado": "Quantidade de Estudantes Distintos"}
        )
    )
    componente_direc_f = (
        df_rapp_filtrado.groupby(["DIREC", "COMPONENTE CURRICULAR"])[
            "CPF_Padronizado"
        ]
        .nunique()
        .reset_index()
        .rename(
            columns={"CPF_Padronizado": "Quantidade de Estudantes Distintos"}
        )
    )
    componente_serie_f = (
        df_rapp_filtrado.groupby(["SÉRIE", "COMPONENTE CURRICULAR"])[
            "CPF_Padronizado"
        ]
        .nunique()
        .reset_index()
        .rename(
            columns={"CPF_Padronizado": "Quantidade de Estudantes Distintos"}
        )
    )
    componente_etapa_f = (
        df_rapp_filtrado.groupby(["ETAPA_RESUMIDA", "COMPONENTE CURRICULAR"])[
            "CPF_Padronizado"
        ]
        .nunique()
        .reset_index()
        .rename(
            columns={"CPF_Padronizado": "Quantidade de Estudantes Distintos"}
        )
    )
    necessidade_direc_f = (
        df_rapp_filtrado.groupby(["DIREC", "CATEGORIA_NECESSIDADES ESPECIAIS"])[
            "CPF_Padronizado"
        ]
        .nunique()
        .reset_index()
        .rename(
            columns={"CPF_Padronizado": "Quantidade de Estudantes Distintos"}
        )
    )
    necessidade_etapa_f = (
        df_rapp_filtrado.groupby(
            ["ETAPA_RESUMIDA", "CATEGORIA_NECESSIDADES ESPECIAIS"]
        )["CPF_Padronizado"]
        .nunique()
        .reset_index()
        .rename(
            columns={"CPF_Padronizado": "Quantidade de Estudantes Distintos"}
        )
    )
    componente_serie_direc_f = (
        df_rapp_filtrado.groupby(["DIREC", "SÉRIE", "COMPONENTE CURRICULAR"])[
            "CPF_Padronizado"
        ]
        .nunique()
        .reset_index()
        .rename(
            columns={"CPF_Padronizado": "Quantidade de Estudantes Distintos"}
        )
    )

    # Grava o arquivo Excel da DIREC atual com todas as abas
    with pd.ExcelWriter(caminho_direc) as writer:
        df_rapp_filtrado.to_excel(writer, sheet_name="Base RAPP", index=False)
        direc_f.to_excel(writer, sheet_name="DIREC", index=False)
        componente_f.to_excel(writer, sheet_name="Componente", index=False)
        serie_f.to_excel(writer, sheet_name="Serie", index=False)
        etapa_f.to_excel(writer, sheet_name="Etapa", index=False)
        necessidade_especial_f.to_excel(
            writer, sheet_name="Neces. Especial", index=False
        )
        serie_direc_f.to_excel(writer, sheet_name="Serie e DIREC", index=False)
        etapa_direc_f.to_excel(writer, sheet_name="Etapa e DIREC", index=False)
        componente_direc_f.to_excel(
            writer, sheet_name="Componente e DIREC", index=False
        )
        componente_serie_f.to_excel(
            writer, sheet_name="Componente e Serie", index=False
        )
        componente_etapa_f.to_excel(
            writer, sheet_name="Componente e Etapa", index=False
        )
        necessidade_direc_f.to_excel(
            writer, sheet_name="Neces. Especial e DIREC", index=False
        )
        necessidade_etapa_f.to_excel(
            writer, sheet_name="Neces. Especial e Etapa", index=False
        )
        componente_serie_direc_f.to_excel(
            writer, sheet_name="Componente, Serie e DIREC", index=False
        )


print("\nTodo o processo foi concluído com sucesso!")










##########################################################################################################################################
'''
Base de estudantes em RAPP a partir de dados passados diretamente pela equipe do GPD da SEEC. Nesse caso, os dados já estavam cruzados entre estudantes e componentes.
A extração e tratamento foi feita pelo GPD da SEEC.

Procedimento:
1. Fonte inicial: dados postos > aba; ‘RAPP_x_COMPONENTES’
2. Criar a ‘CATEGORIA_NECESSIDADES ESPECIAIS’, ‘ETAPA_RESUMIDA’ e ‘CATEGORIA_COMPONENTE’ de acordo com o que tinha feito anteriormente.
3. fazer todas as segmentações feitas anteriormente;
4. Gerar 1 planilha para geral (todos os estudantes) e 1 para cada DIREC (17 planilhas final).

'''
# Utilizar a aba 'RAPP_x_COMPONENTES' como fonte de informação para análise
df_rapp = pd.read_excel(r"D:\Scripts_Python\FGV\RAPP_2026\20260617_Firmino_RAPP 2026 ATUALIZADO EM 16.06.2026.xlsb", sheet_name='RAPP_x_COMPONENTES', engine="pyxlsb")


# Padronizar o CPF do dataframe
# Padronizar CPF: manter apenas dígitos, completar com zeros à esquerda e formatar como XXX.XXX.XXX-XX
df_rapp['CPF_Padronizado'] = (
    df_rapp['CPF']
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

# Excluir componentes que tenham algum valor em 'NOTA_APROVAÇÃO'
df_rapp = df_rapp[df_rapp['NOTA_APROVAÇÃO'].isna()]


# Trocar nomenclatura das séries para padronizar:
mapeamento = {
    '6º Ano': '6º ANO',
    '7º Ano': '7º ANO',
    '8º Ano': '8º ANO',
    '9º Ano': '9º ANO'
}

df_rapp['SÉRIE'] = df_rapp['SÉRIE'].replace(mapeamento)
df_rapp['SÉRIE.1'] = df_rapp['SÉRIE.1'].replace(mapeamento)

# Conferir se as séries são diferentes entre as colunas 'SÉRIE e 'SÉRIE.1'
df_diferentes = df_rapp[df_rapp['SÉRIE'] != df_rapp['SÉRIE.1']]
df_diferentes
# Não há diferenças entre as séries.


# Criar a coluna 'ETAPA_RESUMIDA' a partir da SÉRIE
mapeamento_etapa = {
    '1ª SÉRIE': 'Ensino Médio',
    '2ª SÉRIE': 'Ensino Médio',
    '3ª SÉRIE': 'Ensino Médio',
    '1º SEMESTRE': 'Ensino Médio',
    '2º SEMESTRE': 'Ensino Médio',
    '3° PERÍODO': 'Ensino Médio',
    'TURMA II (8° E 9° ANOS)': 'Ens. Fund. - Anos Finais',
    '6º ANO': 'Ens. Fund. - Anos Finais',
    '7º ANO': 'Ens. Fund. - Anos Finais',
    '8º ANO': 'Ens. Fund. - Anos Finais',
    '9º ANO': 'Ens. Fund. - Anos Finais'
}

df_rapp['ETAPA_RESUMIDA'] = df_rapp['SÉRIE'].map(mapeamento_etapa)


# Criar coluna 'CATEGORIA_COMPONENTE' para diferenciar componentes da BNCC, EPT e Específicos
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

# Componentes EPT
lista_ept = [
    'Informática Básica',
    'Eletricidade Básica',
    'Desenho Técnico',
    'Fundamentos de Lógica e Algoritmos',
    'Arquitetura e Organização de Computadores',
    'Teoria e Fundamentos da Administração',
    'Estatística',
    'Prevenção e Combate a Sinistros',
    'Gestão de Pessoas',
    'Gestão Pública e Terceiro Setor',
    'Controle Ambiental',
    'Lógica de Programação (Algoritmos)',
    'Metodologia do Trabalho Cientifico',
    'Direito Empresarial, Trabalhista e Tributário',
    'Instalações Elétricas de Baixa Tensão',
    'Eletrônica Aplicada',
    'Introdução à Segurança do Trabalho',
    'Fundamentos de Redes de Computadores',
    'Empreendedorismo',
    'Manutenção e Configuração de Computadores',
    'Matemática Financeira',
    'Noções de Eletrônica e Eletricidade',
    'Programação Estruturada',
    'Estudo dos Solos e Materiais de Construção',
    'Sociologia do Trabalho',
    'Contabilidade Geral',
    'Energia Eólica',
    'Estatística Aplicada à Segurança do Trabalho',
    'Programação WEB I e II',
    'Eletrônica Analógica',
    'Tipos de Energia Renovável',
    'Princípios da Agroecologia',
    'Programação Estruturada e Orientada a Objetos',
    'Gestão Organizacional',
    'Metodologia do Trabalho Científico',
    'Prevenção e Controle de Perdas',
    'Psicologia do Trabalho',
    'Energia Solar, Térmica e Fotovoltaica',
    'Fundamentos do Trabalho do Técnico em Redes de Computadores',
    'Gestão de Saúde e Segurança Ocupacional',
    'Primeiros Socorros',
    'Cabeamento Estruturado e Redes de Acesso',
    'Agricultura Familiar',
    'Banco de Dados',
    'Educação Ambiental e Eco Turismo',
    'Educação Digital',
    'Matemática Básica',
    'Planejamento Estratégico',
    'Segurança da Informação',
    'Sistemas Operacionais',
    'Fundamentos da Computação',
    'Fundamentos do Técnico em Redes de Computadores',
    'Gestão Financeira',
    'Instalações Prediais Hidrossanitárias e Elétricas',
    'Lógica de Programação - Algoritmos',
    'Matematica Básica',
    'Mineralogia',
    'Petrografia',
    'Redes de Computadores',
    'Tecnologia em Mídias Digitais',
    'Avaliação e Educação Nutricional',
    'Desenvolvimento Sustentável',
    'Gestão Ambiental',
    'Microbiologia Ambiental',
    'Sequenciamento da Produção',
    'Trabalho de Conclusão de Curso - TCC',
    'Cabeamento Estruturado e Redes de Acesso e Eletricidade Básica',
    'Gestão Organizacional e Segurança do Trabalho',
    'Hospitalidade e Meios de Hospedagem',
    'Impactos Ambientais',
    'Legislação Turistica',
    'Legislação e Segurança do Trabalho',
    'Química e Bioquímica dos Alimentos',
    'Sociedade, Cultura e Meio Ambiente',
    'Tecnologia de Implementação de Redes',
    'Anatomia e Fisiologia Humana - Noções Básicas',
    'Fitopatologia e Dietoterapia da Nutrição',
    'Geologia Geral',
    'História do RN Aplicada ao Turismo',
    'Marketing e Serviços',
    'Matemática Financeira e Estatística',
    'Microbiologia dos Alimentos',
    'Máquinas e Acionamentos Elétricos',
    'Orçamento e Estabilidade',
    'Química Orgânica',
    'Saúde Pública',
    'Tecnologia da Costura, do Enfesto e Corte',
    'Tecnologia da Modelagem',
    'Tecnologia de Cereais'
]

# Criar a nova variável 'CATEGORIA_COMPONENTE' com base na lista de componentes da BNCC e EPT (caso não seja nenhum dos dois, classificar como Específico)
# Definir regras/ condições
condicoes = [
    df_rapp['COMPONENTE CURRICULAR'].isin(lista_bncc),
    df_rapp['COMPONENTE CURRICULAR'].isin(lista_ept)
]

# Rótulos par as condições
rotulos = ['BNCC', 'EPT']

# Aplicar o select com o valor padrão para o que sobrar
df_rapp['CATEGORIA_COMPONENTE'] = np.select(condicoes, rotulos, default='Específico')


# Criar a nova variável 'CATEGORIA_NECESSIDADES ESPECIAIS' para agrupar os tipos de necessidades especiais
# Lista de tipos de necessidades especiais para cada categoria
categorias_necessidades_especiais = {
    'Altas habilidades/superdotação': 'Transtorno do Neurodesenvolvimento',
    'Transtorno de Déficit de Atenção e Hiperatividade (TDAH)': 'Transtorno do Neurodesenvolvimento',
    'Transtorno do Espectro Autista (TEA)': 'Transtorno do Neurodesenvolvimento',
    'Baixa visão': 'Deficiência Visual',
    'Baixa audição': 'Deficiência Auditiva',
    'Surdez': 'Deficiência Auditiva',
    'Discalculia': 'Transtorno do Neurodesenvolvimento',
    'Disgrafia': 'Transtorno do Neurodesenvolvimento',
    'Deficiência Auditiva': 'Deficiência Auditiva',
    'Deficiência física': 'Deficiência Física',
    'Deficiência intelectual': 'Deficiência Intelectual',
    'Deficiência múltipla': 'Deficiência Múltipla',
    'Deficiência intelectual': 'Deficiência Intelectual',
    'Visão monocular': 'Deficiência Visual',
    'Cegueira': 'Deficiência Visual',
    'Perda de visão periférica': 'Deficiência Visual',
    'Paraplegia': 'Deficiência Física',
    'Tetraplegia': 'Deficiência Física',
    'Hemiegia': 'Deficiência Física',
    'Paralisia cerebral': 'Deficiência Física',
    'Amputações e deformidades congênitas': 'Deficiência Física',
    'Síndrone de down': 'Deficiência Intelectual',
    'Síndromes genéticas': 'Deficiência Intelectual',
    'Dislexia': 'Transtorno do Neurodesenvolvimento',
    'Dislalia': 'Transtorno do Neurodesenvolvimento',
    'Disortografia': 'Transtorno do Neurodesenvolvimento'
}

def classificar_necessidade_especial(texto):
    # Trata valores nulos ou vazios
    if pd.isna(texto) or str(texto).strip() in ["", "-", "NÃO INFORMADO"]:
        return "Sem necessidade especial informada"
    
    texto = str(texto).strip()

    # Caso 1: A linha é exatamente igual a uma das chaves do mapeamento
    if texto in categorias_necessidades_especiais:
        return categorias_necessidades_especiais[texto]
    
    # Caso 2: Se não é uma opção única, verificar se existem múltiplas opções conhecidas dentro do texto
    # Contar quantas chaves do dicionário estão presentes na string
    opcoes_encontradas = [opcao for opcao in categorias_necessidades_especiais.keys() if opcao in texto]
    
    if len(opcoes_encontradas) > 1:
        return "Deficiência Múltipla"
    
    # Caso 3: Se encontrou apenas uma (mas talvez com algum caractere extra ou espaço diferente)
    if len(opcoes_encontradas) == 1:
        return categorias_necessidades_especiais[opcoes_encontradas[0]]

    return "Outra Categoria / Não Mapeado"

# 2. Aplicar a nova função
df_rapp['CATEGORIA_NECESSIDADES ESPECIAIS'] = df_rapp['TIPO NECESSIDADE ESPECÍFICA INFORMADAS'].apply(classificar_necessidade_especial)


############ ENDEREÇOS DAS ESCOLAS DE ESTUDANTES EM RAPP
# Adicionar a coluna 'Endereco_escola' para cada estudante
# Importação de base de endereços
df_enderecos = pd.read_csv(r"D:\Scripts_Python\FGV\RAPP_2026\Enderecos_Escolas_RN.csv", sep=',')


# Adicionar a coluna de endereço na base de estudantes em RAPP, utilizando o código INEP da escola
# As colunas de código Inep no mesmo formato
df_rapp["CÓDIGO INEP ESCOLA"] = df_rapp["CÓDIGO INEP ESCOLA"].astype(str)
df_enderecos["Código INEP"] = df_enderecos["Código INEP"].astype(str)

# Só há um valor de código INEP para cada escola no df_enderecos
print(df_enderecos["Código INEP"].is_unique)

# Criar um mapeamento (dicionário/série) de Código INEP para Endereço
mapeamento_enderecos = df_enderecos.drop_duplicates(subset=["Código INEP"]).set_index("Código INEP")["Endereço"]

# 2. Cria a nova coluna direto no df_rapp
df_rapp["Endereço_Escolas"] = df_rapp["CÓDIGO INEP ESCOLA"].map(mapeamento_enderecos)

# Apenas estudantes do 6º ano do Ensino Fundamental
df_rapp_6ano = df_rapp[df_rapp['SÉRIE'] == '6º ANO']


# Salvar em Excel os endereços das escolas dos estudantes do RAPP e dos estudantes do 6º ano
with pd.ExcelWriter(r"D:\Scripts_Python\FGV\RAPP_2026\Resultados\20260617_RAPP\20260617_Enderecos_RAPP_Estudantes.xlsx") as writer:
    df_rapp.to_excel(writer, sheet_name='Endereços Escolas RAPP', index=False)
    df_rapp_6ano.to_excel(writer, sheet_name='Endereços Escolas 6º ano', index=False)


# Salvar somente lista de endereços para as escolas
# Manter um dataframe apenas com os endereços de cada escola
colunas_selecionadas = [
    'DIREC',
    'CÓDIGO INEP ESCOLA',
    'ESCOLA',
    'MUNICÍPIO',
    'Endereço_Escolas'
]

# Filtrar o DataFrame e remover as duplicatas mantendo apenas a primeira ocorrência
df_rapp_escolas = df_rapp[colunas_selecionadas].drop_duplicates(subset=['CÓDIGO INEP ESCOLA'])

# Filtrar o dataframe de 6º ano e remover as duplicatas mantendo apenas a primeira ocorrência
df_rapp_6ano_escolas = df_rapp_6ano[colunas_selecionadas].drop_duplicates(subset=['CÓDIGO INEP ESCOLA'])


# Salvar em Excel os endereços das escolas do RAPP e dos estudantes do 6º ano
with pd.ExcelWriter(r"D:\Scripts_Python\FGV\RAPP_2026\Resultados\20260617_RAPP\20260617_Enderecos_RAPP_Escolas.xlsx") as writer:
    df_rapp_escolas.to_excel(writer, sheet_name='Endereços Escolas RAPP', index=False)
    df_rapp_6ano_escolas.to_excel(writer, sheet_name='Endereços Escolas 6º ano', index=False)


######### (NÃO VAI SER EXCLUÍDO. É CONTADO DUAS VEZES)
######### Excluir valores duplicados de componente-CPF_Padronizado
##### df_rapp = df_rapp.drop_duplicates(subset=['COMPONENTE CURRICULAR', 'CPF_Padronizado'], keep='first')




# Fazer as segmentações e contagem de interesse: estudantes por DIREC; por componente; por turno; por Série; necessidades especiais etc.
#################################### ANÁLISES ####################################
# Estudantes por DIREC
# Contagem de CPF_Padronizado distinto por DIREC
direc = df_rapp.groupby('DIREC')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudantes por Componente Curricular
# Contagem de CPF_Padronizado distinto por Componente Curricular
componente = df_rapp.groupby('COMPONENTE CURRICULAR')['CPF_Padronizado'].count().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes'})

# Estudantes por Série
# Contagem de CPF_Padronizado distinto por Série
serie = df_rapp.groupby('SÉRIE')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudantes por Etapa de Ensino
# Contagem de CPF_Padronizado distinto por Etapa de Ensino
etapa = df_rapp.groupby('ETAPA_RESUMIDA')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudantes por tipo de Necessidade Especial
necessidade_especial = df_rapp.groupby('CATEGORIA_NECESSIDADES ESPECIAIS')['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por Série em cada DIREC
serie_direc = df_rapp.groupby(['DIREC', 'SÉRIE'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por Etapa em cada DIREC
etapa_direc = df_rapp.groupby(['DIREC', 'ETAPA_RESUMIDA'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por componente em cada DIREC
componente_direc = df_rapp.groupby(['DIREC', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].count().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes'})

# Estudante por Componente por Série
componente_serie = df_rapp.groupby(['SÉRIE', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].count().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes'})

# Estudante por Componente por Etapa de Ensino
componente_etapa = df_rapp.groupby(['ETAPA_RESUMIDA', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].count().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes'})

# Necessidade Especial por DIREC
necessidade_direc = df_rapp.groupby(['DIREC', 'CATEGORIA_NECESSIDADES ESPECIAIS'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes'})

# Necessidade Especial por ETAPA
necessidade_etapa = df_rapp.groupby(['ETAPA_RESUMIDA', 'CATEGORIA_NECESSIDADES ESPECIAIS'])['CPF_Padronizado'].nunique().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes Distintos'})

# Estudante por Componente, por Série e por DIREC
componente_serie_direc = df_rapp.groupby(['DIREC', 'SÉRIE', 'COMPONENTE CURRICULAR'])['CPF_Padronizado'].count().reset_index().rename(columns={'CPF_Padronizado': 'Quantidade de Estudantes'})


# Caminho da pasta onde os arquivos serão salvos
pasta_destino = r"D:\Scripts_Python\FGV\RAPP_2026\Resultados\20260625_RAPP"


# Salva o arquivo GERAL
caminho_geral = os.path.join(pasta_destino, "20260625_GERAL_analises_RAPP.xlsx")

with pd.ExcelWriter(caminho_geral) as writer:
    df_rapp.to_excel(writer, sheet_name='Base RAPP', index=False)
    direc.to_excel(writer, sheet_name='DIREC', index=False)
    componente.to_excel(writer, sheet_name='Componente', index=False)
    serie.to_excel(writer, sheet_name='Serie', index=False)
    etapa.to_excel(writer, sheet_name='Etapa', index=False)
    necessidade_especial.to_excel(writer, sheet_name='Neces. Especial', index=False)
    serie_direc.to_excel(writer, sheet_name='Serie e DIREC', index=False)
    etapa_direc.to_excel(writer, sheet_name='Etapa e DIREC', index=False)
    componente_direc.to_excel(writer, sheet_name='Componente e DIREC', index=False)
    componente_serie.to_excel(writer, sheet_name='Componente e Serie', index=False)
    componente_etapa.to_excel(writer, sheet_name='Componente e Etapa', index=False)
    necessidade_direc.to_excel(writer, sheet_name='Neces. Especial e DIREC', index=False)
    necessidade_etapa.to_excel(writer, sheet_name='Neces. Especial e Etapa', index=False)
    componente_serie_direc.to_excel(writer, sheet_name='Componente, Serie e DIREC', index=False)


 
#################################################################
# GERAR 1 PLANILHA PARA CADA DIREC
print("\nIniciando a criação das planilhas por DIREC...")

# Obtém a lista de DIRECs únicas diretamente da coluna para evitar erros de digitação
lista_direcs = df_rapp["DIREC"].dropna().unique()

for d in lista_direcs:
    print(f" -> Processando: {d}")

    # Formata o nome do arquivo (ex: '01ª DIREC - NATAL' vira '01_DIREC_NATAL')
    nome_limpo = (
        str(d).replace("ª", "").replace(" - ", "_").replace(" ", "_")
    )
    caminho_direc = os.path.join(
        pasta_destino, f"20260617_{nome_limpo}_analises_RAPP.xlsx"
    )

    # Passo chave: Filtra a base principal APENAS para a DIREC do loop atual
    df_rapp_filtrado = df_rapp[df_rapp["DIREC"] == d]

    # Recalcula todas as tabelas usando o DataFrame filtrado
    direc_f = (
        df_rapp_filtrado.groupby("DIREC")["CPF_Padronizado"]
        .nunique()
        .reset_index()
        .rename(
            columns={"CPF_Padronizado": "Quantidade de Estudantes Distintos"}
        )
    )
    componente_f = (
        df_rapp_filtrado.groupby("COMPONENTE CURRICULAR")["CPF_Padronizado"]
        .count()
        .reset_index()
        .rename(
            columns={"CPF_Padronizado": "Quantidade de Estudantes"}
        )
    )
    serie_f = (
        df_rapp_filtrado.groupby("SÉRIE")["CPF_Padronizado"]
        .nunique()
        .reset_index()
        .rename(
            columns={"CPF_Padronizado": "Quantidade de Estudantes Distintos"}
        )
    )
    etapa_f = (
        df_rapp_filtrado.groupby("ETAPA_RESUMIDA")["CPF_Padronizado"]
        .nunique()
        .reset_index()
        .rename(
            columns={"CPF_Padronizado": "Quantidade de Estudantes Distintos"}
        )
    )
    necessidade_especial_f = (
        df_rapp_filtrado.groupby("CATEGORIA_NECESSIDADES ESPECIAIS")[
            "CPF_Padronizado"
        ]
        .nunique()
        .reset_index()
        .rename(
            columns={"CPF_Padronizado": "Quantidade de Estudantes Distintos"}
        )
    )
    serie_direc_f = (
        df_rapp_filtrado.groupby(["DIREC", "SÉRIE"])["CPF_Padronizado"]
        .nunique()
        .reset_index()
        .rename(
            columns={"CPF_Padronizado": "Quantidade de Estudantes Distintos"}
        )
    )
    etapa_direc_f = (
        df_rapp_filtrado.groupby(["DIREC", "ETAPA_RESUMIDA"])[
            "CPF_Padronizado"
        ]
        .nunique()
        .reset_index()
        .rename(
            columns={"CPF_Padronizado": "Quantidade de Estudantes Distintos"}
        )
    )
    componente_direc_f = (
        df_rapp_filtrado.groupby(["DIREC", "COMPONENTE CURRICULAR"])[
            "CPF_Padronizado"
        ]
        .count()
        .reset_index()
        .rename(
            columns={"CPF_Padronizado": "Quantidade de Estudantes"}
        )
    )
    componente_serie_f = (
        df_rapp_filtrado.groupby(["SÉRIE", "COMPONENTE CURRICULAR"])[
            "CPF_Padronizado"
        ]
        .count()
        .reset_index()
        .rename(
            columns={"CPF_Padronizado": "Quantidade de Estudantes"}
        )
    )
    componente_etapa_f = (
        df_rapp_filtrado.groupby(["ETAPA_RESUMIDA", "COMPONENTE CURRICULAR"])[
            "CPF_Padronizado"
        ]
        .count()
        .reset_index()
        .rename(
            columns={"CPF_Padronizado": "Quantidade de Estudantes"}
        )
    )
    necessidade_direc_f = (
        df_rapp_filtrado.groupby(["DIREC", "CATEGORIA_NECESSIDADES ESPECIAIS"])[
            "CPF_Padronizado"
        ]
        .nunique()
        .reset_index()
        .rename(
            columns={"CPF_Padronizado": "Quantidade de Estudantes Distintos"}
        )
    )
    necessidade_etapa_f = (
        df_rapp_filtrado.groupby(
            ["ETAPA_RESUMIDA", "CATEGORIA_NECESSIDADES ESPECIAIS"]
        )["CPF_Padronizado"]
        .nunique()
        .reset_index()
        .rename(
            columns={"CPF_Padronizado": "Quantidade de Estudantes Distintos"}
        )
    )
    componente_serie_direc_f = (
        df_rapp_filtrado.groupby(["DIREC", "SÉRIE", "COMPONENTE CURRICULAR"])[
            "CPF_Padronizado"
        ]
        .count()
        .reset_index()
        .rename(
            columns={"CPF_Padronizado": "Quantidade de Estudantes"}
        )
    )

    # Grava o arquivo Excel da DIREC atual com todas as abas
    with pd.ExcelWriter(caminho_direc) as writer:
        df_rapp_filtrado.to_excel(writer, sheet_name="Base RAPP", index=False)
        direc_f.to_excel(writer, sheet_name="DIREC", index=False)
        componente_f.to_excel(writer, sheet_name="Componente", index=False)
        serie_f.to_excel(writer, sheet_name="Serie", index=False)
        etapa_f.to_excel(writer, sheet_name="Etapa", index=False)
        necessidade_especial_f.to_excel(
            writer, sheet_name="Neces. Especial", index=False
        )
        serie_direc_f.to_excel(writer, sheet_name="Serie e DIREC", index=False)
        etapa_direc_f.to_excel(writer, sheet_name="Etapa e DIREC", index=False)
        componente_direc_f.to_excel(
            writer, sheet_name="Componente e DIREC", index=False
        )
        componente_serie_f.to_excel(
            writer, sheet_name="Componente e Serie", index=False
        )
        componente_etapa_f.to_excel(
            writer, sheet_name="Componente e Etapa", index=False
        )
        necessidade_direc_f.to_excel(
            writer, sheet_name="Neces. Especial e DIREC", index=False
        )
        necessidade_etapa_f.to_excel(
            writer, sheet_name="Neces. Especial e Etapa", index=False
        )
        componente_serie_direc_f.to_excel(
            writer, sheet_name="Componente, Serie e DIREC", index=False
        )


print("\nTodo o processo foi concluído com sucesso!")






