# %% [markdown]
# #### Importar libs

# %%
from datetime import datetime
from zoneinfo import ZoneInfo
import requests
import pandas as pd
from io import BytesIO

import os
import sys

sys.path.insert(0, r'C:\Scripts\modules\database')
sys.path.insert(0, r'C:\Scripts\modules\parameters')

from connection import Connect
from bases import Bases

# %% [markdown]
# #### Atualizar o ROA e os volumes detalhados

# %%
conexao = Connect.connect_techdb()
hist = Connect.import_table(conexao, "previa_receita_nova")
conexao.close()

# %%
brasilia_tz = ZoneInfo('America/Sao_Paulo')

# 2. Remove o mês atual do histórico para evitar duplicação
hist = hist[hist['Data'] != datetime.now(brasilia_tz).strftime("%Y-%m")]

# %%
def load_previas(advisor_name, link: str):
    response = requests.get("{}".format(link), params={"downloadformat": "excel"})
    df = pd.read_excel(BytesIO(response.content), sheet_name='Meta')
    
    # Mapeamento direto das colunas pela posição (Índice):
    # Coluna 0: Nome da Categoria (RF, RV, Fundos...)
    # Coluna 1: Meta Volume
    # Coluna 2: Realizado Volume (O cabeçalho é o nome do assessor)
    
    df.rename(columns={
        df.columns[0]: "Categoria - Acompanhamento Next",
        df.columns[1]: "META - ROA",
        df.columns[2]: "REALIZADO - ROA",
        df.columns[4]: "REALIZADO - VOLUME"
    }, inplace=True)

    # Cria as colunas de ROA zeradas para satisfazer a estrutura da tabela do banco
    df['META - VOLUME'] = 0.0
    
    # Remove linha de TOTAL se existir, para não duplicar valores no BI
    df = df[df['Categoria - Acompanhamento Next'] != "TOTAL"]
    
    # Adiciona o nome do assessor
    df['Assessor'] = advisor_name
    
    # Reordena para garantir que bate com a ordem do banco/histórico
    colunas_ordem = [
        'Categoria - Acompanhamento Next', 
        'META - VOLUME', 
        'REALIZADO - VOLUME', 
        'META - ROA', 
        'REALIZADO - ROA', 
        'Assessor'
    ]
    
    return df[colunas_ordem]

# %%
rodrigo = load_previas(link="https://netorg18892072-my.sharepoint.com/:x:/g/personal/joao_aquino_atriacm_com_br/IQBVuGicHybdRrC4d1MtFO8vAbY4Kw4m4_8gNo8EKu3BN4I?download=1", advisor_name="RODRIGO DE MELLO D’ELIA")
caic = load_previas(link="https://netorg18892072-my.sharepoint.com/:x:/g/personal/joao_aquino_atriacm_com_br/IQByqpoVZXN3TYdcUlFmYqE9Af8AsDkk6umaNM26wLfQlo4?download=1", advisor_name="CAIC ZEM GOMES")
fernando = load_previas(link="https://netorg18892072-my.sharepoint.com/:x:/g/personal/joao_aquino_atriacm_com_br/IQAFP5cFxU9QRL_OWClG-BJdAT3Wvk18_VoypIF3CxIpQYY?download=1", advisor_name="FERNANDO DOMINGUES DA SILVA")
saadallah = load_previas(link="https://netorg18892072-my.sharepoint.com/:x:/g/personal/joao_aquino_atriacm_com_br/IQDjXAOoHHxgQYo_aeQQdXz4AXzuH0TotDouMPG_NNH4H-4?download=1", advisor_name="SAADALLAH JOSE ASSAD")
paulo = load_previas(link="https://netorg18892072-my.sharepoint.com/:x:/g/personal/joao_aquino_atriacm_com_br/IQA6mQn9Z8lwTL7HfLdw1UzSAbCoa8qOg03wNZvbJyDoIaw?download=1", advisor_name="PAULO ROBERTO FARIA SILVA")
marcos = load_previas(link="https://netorg18892072-my.sharepoint.com/:x:/g/personal/joao_aquino_atriacm_com_br/IQAtEitVDNlnT4zWgSIonhfzAQlTKUlOX3xzRRYfgcoHrZE?download=1", advisor_name="MARCOS SOARES PEREIRA FILHO")
renan_bento = load_previas(link="https://netorg18892072-my.sharepoint.com/:x:/g/personal/joao_aquino_atriacm_com_br/IQCr2UivmbvdQ4EIlpS8fcTfAVmuuqLXeLhqroDonRcYXDQ?download=1", advisor_name = "RENAN BENTO DA SILVA")
rosana = load_previas(link="https://netorg18892072-my.sharepoint.com/:x:/g/personal/joao_aquino_atriacm_com_br/IQDmCPZulmfjTYtC94MGxHh1ARzYBjyg9gWNL9v2U--M0CQ?download=1", advisor_name = "ROSANA APARECIDA PAVANI DA SILVA")
rafael_pasold = load_previas(link="https://netorg18892072-my.sharepoint.com/:x:/g/personal/joao_aquino_atriacm_com_br/IQBVpoNJYZ8NSY9rnRotpPSPAXlir-PVsE7SItom19ZhijI?download=1", advisor_name = "RAFAEL PASOLD MEDEIROS")
felipe = load_previas(link="https://netorg18892072-my.sharepoint.com/:x:/g/personal/joao_aquino_atriacm_com_br/IQCy_TXmqwgLSL2M6iwrcsp0AWNG5plDMI-xPRY363bX2dA?download=1", advisor_name = "FELIPE AUGUSTO CARDOSO")
guilherme = load_previas(link="https://netorg18892072-my.sharepoint.com/:x:/g/personal/joao_aquino_atriacm_com_br/IQCZlPDCr4lsTpfTGlsnjvQiAT0lnyonDCFpCKlASYnfES4?download=1", advisor_name = "GUILHERME DE LUCCA BERTELONI")
izadora = load_previas(link="https://netorg18892072-my.sharepoint.com/:x:/g/personal/joao_aquino_atriacm_com_br/IQD4lZoreFnUQYKgxi0ibLONAV1CwfMLplcscrKNd6ZJPis?download=1", advisor_name = "IZADORA VILLELA FREITAS")
vitor = load_previas(link="https://netorg18892072-my.sharepoint.com/:x:/g/personal/joao_aquino_atriacm_com_br/IQAmZYnDo_EeToIhGV7QRgYJAb6YYoMOMaWaAEadD6WOS4I?download=1", advisor_name = "VITOR OLIVEIRA DOS REIS")

# %%
# 1. Lista para concatenar
lista_dfs = [rodrigo, caic, fernando, saadallah, paulo, marcos, renan_bento, rosana, rafael_pasold, felipe, guilherme, izadora, vitor] 

previa_receita = pd.concat(lista_dfs, ignore_index=True)

# 2. Adiciona Data e Hora atuais
# (Certifique-se que você importou o 'brasilia_tz' antes ou use datetime.now() puro se der erro)
previa_receita['Data'] = pd.to_datetime(datetime.now().strftime("%Y-%m"))
previa_receita['Hora Atualizado'] = pd.to_datetime(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

# Lista quem carregou com sucesso
print("Assessores identificados na base consolidada:")
assessores_unicos = previa_receita['Assessor'].unique()
print(list(assessores_unicos))

print("-" * 30)
print("Quantidade de linhas por Assessor:")
print(previa_receita['Assessor'].value_counts())

print("-" * 30)
print("RESUMO FINANCEIRO CONSOLIDADO:")

# Verifica se as colunas existem antes de somar para evitar erro
coluna_roa = 'REALIZADO - ROA' 
coluna_meta = 'META - ROA'       

if coluna_roa in previa_receita.columns and coluna_meta in previa_receita.columns:
    # Soma (fillna(0) garante que células vazias não atrapalhem a conta)
    soma_roa = previa_receita[coluna_roa].fillna(0).sum()
    soma_meta = previa_receita[coluna_meta].fillna(0).sum()

    # Formatação B
    roa_fmt = f"R$ {soma_roa:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    meta_fmt = f"R$ {soma_meta:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    print(f"Soma {coluna_roa}:   {roa_fmt}")
    print(f"Soma {coluna_meta}:          {meta_fmt}")


print("-" * 30)
print(f"Colunas atuais: {list(previa_receita.columns)}")
print("="*50)

# %%
# 4. Junta o histórico (categorias antigas) com o atual (categorias novas)
previa_receita = pd.concat([hist, previa_receita], axis=0)

# %%
# 5. Tratamento final de nulos
previa_receita.loc[previa_receita['META - VOLUME'] == '-', 'META - VOLUME'] = 0
previa_receita.fillna(0, inplace=True)

# %%
# 6. Salva no Banco mantendo a estrutura original
conexao = Connect.connect_techdb()
previa_receita.to_sql('previa_receita_nova', 
                    conexao, 
                    if_exists='replace', 
                    index=False,
                    schema="principal")
conexao.close()

# %% [markdown]
# #### Atualizar receita previa por assessor

# %%
# 1. Agrupamento da prévia atual
previa_receita_assessor = previa_receita.loc[:, ["Assessor", "Data", "META - ROA", "REALIZADO - ROA"]].groupby(['Assessor', 'Data']).sum().reset_index()

# %%
# 2. Carrega histórico
conexao = Connect.connect_techdb()
historico_previa_receita = Connect.import_table(conexao, "previa_receita_assessor_historico")
conexao.close()

# %%
# 3. Normaliza as datas (Remove horas/dias quebrados e deixa dia 01)
historico_previa_receita['Data'] = (
    pd.to_datetime(historico_previa_receita['Data'])
    .dt.to_period('M')
    .dt.to_timestamp()
)

previa_receita_assessor['Data'] = (
    pd.to_datetime(previa_receita_assessor['Data'])
    .dt.to_period('M')
    .dt.to_timestamp()
)

# 4. Define mês atual (baseado em Brasília)
mes_atual = pd.Timestamp.now(tz=brasilia_tz).to_period('M').to_timestamp()

# 5. Remove APENAS o mês atual do histórico (limpa para atualizar)
historico_previa_receita = historico_previa_receita[
    historico_previa_receita['Data'] != mes_atual
]

# 6. Seleciona apenas o mês atual da nova tabela
previa_receita_mes_atual = previa_receita_assessor[
    previa_receita_assessor['Data'] == mes_atual
]

# 7. Reconstrói histórico (Histórico Antigo + Mês Novo)
# (O processamento continua ocorrendo para garantir que a variável final exista
previa_receita_assessor_historico = pd.concat(
    [historico_previa_receita, previa_receita_mes_atual],
    axis=0,
    ignore_index=True
)

print("\n" + "="*50)
print("Dataframe - Mês Atual")
print("="*50)

# Validação de Data e Volume
print(f"Mês de Referência: {mes_atual}")
print(f"Total de linhas a inserir: {len(previa_receita_mes_atual)}")
print("-" * 30)

# Validação dos valores acumulados (Soma rápida para checagem)
soma_meta = previa_receita_mes_atual['META - ROA'].sum()
soma_realizado = previa_receita_mes_atual['REALIZADO - ROA'].sum()
print(f"Total META (Mês): {soma_meta:,.2f}")
print(f"Total REALIZADO (Mês): {soma_realizado:,.2f}")

print("-" * 30)
print("Visualização da Tabela (Apenas mês atual):")


# %%
conexao = Connect.connect_techdb()
previa_receita_assessor_historico.to_sql(
    'previa_receita_assessor_historico',
    conexao,
    if_exists='replace',
    index=False,
    schema="principal"
)
conexao.close()