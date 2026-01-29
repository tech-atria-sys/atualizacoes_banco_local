# %%
import pandas as pd 
from pandas.tseries.offsets import BDay
from datetime import datetime, timedelta
from dateutil import relativedelta
import calendar 
import win32com.client as win32
from email.message import EmailMessage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase
from email import encoders
import smtplib, ssl
import os
import xlsxwriter

import sys
sys.path.insert(0, r'C:\Scripts\modules\database')
sys.path.insert(0, r'C:\Scripts\modules\parameters')

from connection import Connect
from bases import Bases

# %%
# Se True: Manda para todos os emails que estiverem na tabela 'times'.
# Se False: Manda APENAS para os nomes que estiverem na lista 'ALVOS_ESPECIFICOS'.
ENVIAR_PARA_TODOS = True

# Coloque aqui os nomes dos assessores
DESTINATARIO = [#'FERNANDO DOMINGUES DA SILVA',
#'PAULO ROBERTO FARIA SILVA',
#'SAADALLAH JOSE ASSAD',
#'RODRIGO DE MELLO D’ELIA',
#'ROSANA APARECIDA PAVANI DA SILVA',
#'GABRIEL GUERRERO TORRES FONSECA',
'CAIC ZEM GOMES',
#'RENAN BENTO DA SILVA',
#'RAFAEL PASOLD MEDEIROS',
#'FELIPE AUGUSTO CARDOSO',
#'MARCOS SOARES PEREIRA FILHO',
#'IZADORA VILLELA FREITAS',
#'GUILHERME DE LUCCA BERTELONI',
#'VITOR OLIVEIRA DOS REIS',
]

# Se True, ele apenas printa que enviaria, mas não envia
MODO_TESTE = True

# %%
## Adicionar coluna de time e de pf ou pj

# %%
conexao = Connect.connect_techdb()
basebtg = Connect.import_table(conexao, "base_btg")
times = Connect.import_table(conexao, "times_nova_empresa")
saldo = Connect.import_table(conexao, "saldo_conta_corrente")
posicoes = Connect.import_table(conexao, "posicao")
conexao.close()

# %%
basebtg = basebtg.merge(times, on='Assessor', how='left')

# %%
## Clientes com Conta, Assessor e Nome completo
basebtg = basebtg[['Conta', 'Assessor', 'Tipo', 'TIME']]
nomes_clientes = pd.read_excel(r"C:\Scripts\nomes_clientes\Nomes_clientes2.xlsx")
nomes_clientes.rename(columns={"Código":"Conta"}, inplace=True)
basebtg = basebtg.merge(nomes_clientes, on='Conta', how='left')

# %%
## Criar novo dataframe com os clientes + saldo em cc
saldo.rename(columns={"CONTA":"Conta"}, inplace=True)
table = pd.merge(basebtg, saldo[['Conta', 'SALDO']], on='Conta', how='left')

# %%
for coluna in posicoes.columns:
    posicoes.rename(columns={coluna:coluna.upper()}, inplace=True)

# %%
## Adiconar LFTs
lft = posicoes[posicoes['PRODUTO'] == 'LFT']
lft = lft.iloc[:,[0, 16]].groupby("CONTA").sum().reset_index()
lft.rename(columns={"VALOR LÍQUIDO":'LFT'}, inplace=True)

table.rename(columns={"Conta":"CONTA"}, inplace=True)
table = table.merge(lft[['CONTA', 'LFT']], on='CONTA', how='left')

# %%
## Adicionar CDBs de líquidez diária
cdb = posicoes[posicoes['PRODUTO'] == "CDB"]

conexao = Connect.connect_techdb()
rf_coe = Connect.import_table(conexao, "renda_fixa")
conexao.close()

for coluna in rf_coe.columns:
    rf_coe.rename(columns={coluna:coluna.upper()}, inplace=True)

rf_coe = rf_coe[['ATIVO', 'LIQUIDEZ']]

cdb = cdb.merge(rf_coe, on='ATIVO', how='left')
cdb_liq_diaria = cdb[cdb['LIQUIDEZ'] == "Liquidez Diaria"]
cdb_liq_diaria = cdb_liq_diaria[['CONTA', 'VALOR LÍQUIDO']].groupby("CONTA").sum().reset_index()
cdb_liq_diaria.rename(columns={"VALOR LÍQUIDO":"CDB Líquidez Diária"}, inplace=True)
table = table.merge(cdb_liq_diaria, on='CONTA', how='left')

# %%
hoje = datetime.today()
res = calendar.monthrange(hoje.year, hoje.month)
day = res[1]
ultimo_dia_do_mes = datetime(hoje.year, hoje.month, day).strftime("%Y-%m-%d")

# %%
## Adicionar os vencimentos até o fim do mes
vencimento = posicoes[posicoes['VENCIMENTO'] <= ultimo_dia_do_mes]
vencimento = vencimento[vencimento['MERCADO'] != "Valor em Trânsito"]
vencimento = vencimento[['CONTA', 'VALOR LÍQUIDO']].groupby("CONTA").sum().reset_index()
vencimento.rename(columns={"VALOR LÍQUIDO":"Vencimentos Mes Atual"}, inplace=True)
table = table.merge(vencimento, on='CONTA', how='left')

# %%
## Adicionar fundos d+0 e d+1
#fundos = fundos_posicao.copy()
#fundos.rename(columns={"'DL_D_ContaAssessor'[NR_CONTA]":"CONTA"}, inplace=True)
#for coluna in fundos.columns:
#    fundos.rename(columns={coluna:coluna.upper()}, inplace=True)
#fundos.rename(columns={"RESGATE (D+)":"REGATE (D+)"}, inplace=True)
#fundos = fundos[(fundos['REGATE (D+)'] == 0) | (fundos['REGATE (D+)'] == 1)]
#fundos_liq = fundos[['CONTA', 'VALOR LÍQUIDO']].groupby("CONTA").sum().reset_index()
#fundos_liq.rename(columns={"VALOR LÍQUIDO":"Fundos d+0/d+1"}, inplace=True)
#fundos_liq['CONTA'] = fundos_liq['CONTA'].astype(str)
#table = table.merge(fundos_liq, on='CONTA', how='left')

# %%
table.fillna(0, inplace=True)

# %%
conexao = Connect.connect_techdb()
basebtg = Connect.import_table(conexao, "base_btg")
conexao.close()

# %%
basebtg = basebtg[['Conta', 'PL Total', 'PL Declarado']]
basebtg.rename(columns={"Conta":"CONTA"}, inplace=True)
basebtg['% share of wallet'] = basebtg['PL Total']/basebtg['PL Declarado']
table = table.merge(basebtg[['CONTA', "% share of wallet"]], on='CONTA', how='left')

# %%
conexao = Connect.connect_techdb()
proventos = Connect.import_table(conexao, 'proventos_futuros_rf')
conexao.close()

# %%
proventos.head(2)

# %%
proventos.rename(columns={"Mes":"Data", "Conta":"CONTA"}, inplace=True)
proventos_por_mes = proventos.iloc[:,[0, 7, 9]].groupby(["CONTA", "Data"]).sum().reset_index()

# %%
proventos_no_mes = \
    proventos_por_mes[
        (proventos_por_mes['Data'] == datetime.today().strftime("%Y-%m"))
    ]

# %%
proventos_no_mes.rename(columns={"Total proventos":"Proventos no Mês"}, inplace=True)

# %%
table = table.merge(proventos_no_mes[['CONTA', 'Proventos no Mês']], on='CONTA', how='left')

# %%
table['Proventos no Mês'].fillna(0, inplace=True)
table['Soma Liquidez'] = table['SALDO'] + table['LFT'] + table['CDB Líquidez Diária'] + table['Vencimentos Mes Atual'] + table['Proventos no Mês']

# %%
pl_clientes = basebtg.loc[:,["CONTA", "PL Total"]]

# %%
table = table.merge(pl_clientes)
table["Líquidez da Carteira"] = table["Soma Liquidez"]/table["PL Total"]

# %%
liquidez_assessor = table.loc[:,["Assessor", "Soma Liquidez", "PL Total"]].groupby("Assessor").sum().reset_index()
liquidez_assessor["Liquidez da carteira"] = liquidez_assessor["Soma Liquidez"]/liquidez_assessor["PL Total"]
liquidez_assessor = liquidez_assessor.sort_values("Liquidez da carteira", ascending=False)
liquidez_assessor[(liquidez_assessor['Assessor'].isin(times['Assessor'])) & (liquidez_assessor['Assessor'] != "DIGITAL")].loc[:,["Assessor", "Liquidez da carteira"]]

# %%
liquidez_assessor[(liquidez_assessor['Assessor'].isin(times['Assessor'])) & (liquidez_assessor['Assessor'] != "DIGITAL")]['Liquidez da carteira'].mean()

# %%
conexao = Connect.connect_techdb()
table.to_sql(con=conexao, name="liquidez_clientes", 
             if_exists='replace', 
             index=False,
             schema='principal')
conexao.close()
table.to_excel(r"C:\Scripts\liquidez\pacote_liquidez.xlsx", header=True, index=False)

# %%
infos = table[[ 
    'CONTA', 'Assessor', 'Nome', 'SALDO','PL Total', 'CDB Líquidez Diária', 'LFT',
    'Vencimentos Mes Atual', 'Proventos no Mês', 'Soma Liquidez',
    'Líquidez da Carteira'
]]

infos.to_excel(
    r"C:\Scripts\liquidez\infos_todos.xlsx",
    index=False
)

# %%
infos

# %%
caminho_base = r"C:\Scripts\liquidez"

outlook = win32.Dispatch('outlook.application')

# Garante que a pasta existe
if not os.path.exists(caminho_base):
    os.makedirs(caminho_base)

# Lógica de seleção de lista
if ENVIAR_PARA_TODOS:
    # Pega todos os assessores únicos que estão na base de dados 'times'
    lista_assessores = times['Assessor'].unique()
else:
    # Usa a lista manual definida
    lista_assessores = DESTINATARIO

print(f"Iniciando processamento para {len(lista_assessores)} assessores...\n")

for nome_assessor in lista_assessores:
    dados_assessor = times[times['Assessor'] == nome_assessor]
    if dados_assessor.empty:
        continue 
    email_destino = dados_assessor['Email'].iloc[0]
    
    # Usamos .copy() para evitar avisos do Pandas
    infos_filtrada = infos[infos['Assessor'] == nome_assessor].copy()
    
    if infos_filtrada.empty:
        print(f"[AVISO] {nome_assessor} sem dados. Pulando.")
        continue

    # 1. Remove a coluna Assessor
    infos_filtrada = infos_filtrada.drop(columns=['Assessor'])

    # 2. Ordena pelo SALDO (Fazemos isso antes de cortar as colunas, para garantir a ordem)
    if 'SALDO' in infos_filtrada.columns:
        infos_filtrada = infos_filtrada.sort_values(by='SALDO', ascending=False)

    # --- TRATAMENTO ESPECÍFICO PARA O CAIC ---
    if nome_assessor == 'CAIC ZEM GOMES':
        # Mantém apenas as 4 primeiras colunas (A, B, C, D)
        infos_filtrada = infos_filtrada.iloc[:, :4]
    # -----------------------------------------

    # adiciona o nome do arquivo na pasta
    nome_arquivo = f"infos_{nome_assessor}.xlsx"
    caminho_completo = os.path.join(caminho_base, nome_arquivo)
    
    with pd.ExcelWriter(caminho_completo, engine='xlsxwriter') as writer:

        # Joga os dados para a planilha
        infos_filtrada.to_excel(writer, index=False, sheet_name='Relatorio')
        
        # Acessa os objetos de formatação
        workbook = writer.book
        worksheet = writer.sheets['Relatorio']
        
        # --- ADICIONA O FILTRO ---
        # Pega o número de linhas e colunas do dataframe filtrado
        max_row = len(infos_filtrada)
        max_col = infos_filtrada.shape[1] - 1  # -1 pois o índice começa em 0
        
        # Aplica o filtro da célula A1 (0,0) até a última célula ocupada
        worksheet.autofilter(0, 0, max_row, max_col)

        # --- CRIAÇÃO DOS ESTILOS ---
        fmt_moeda = workbook.add_format({'num_format': 'R$ #,##0.00'})   
        fmt_cabecalho = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1})
        fmt_porcentagem = workbook.add_format({'num_format': '0.00%'})

        # --- APLICANDO NOS DADOS ---
        
        # Ajustar largura das colunas (A até Z ficam com largura 23)
        worksheet.set_column('A:Z', 23) 
        worksheet.set_column('B:B', 40) 
        worksheet.set_column('A:A', 14) 

        # Obs: Se o arquivo for do CAIC e não tiver colunas E:L, 
        # esses comandos abaixo apenas formatarão células vazias (não gera erro).
        
        # Aplicar formato de monetário
        worksheet.set_column('C:I', 23, fmt_moeda) 
        
        # Aplica porcentagem
        worksheet.set_column('J:J', 23, fmt_porcentagem)


    # 3. ENVIO DE E-MAIL
    if MODO_TESTE:
        print(f"--- [TESTE] Arquivo gerado: {nome_arquivo} | Cols: {infos_filtrada.shape[1]} | Para: {email_destino}")
    else:
        try:
            email = outlook.CreateItem(0)
            email.To = email_destino
            email.Subject = "Relatório de Saldo e Liquidez - Carteira"
            email.HTMLBody = f"""
            <p>Olá, {nome_assessor.split()[0].title()}</p>
            <p>Segue em anexo a planilha formatada com os saldos atualizados.</p>
            <p>Abs,</p>
            """
            email.Attachments.Add(caminho_completo)
            email.Send()
            print(f"[SUCESSO] Enviado para: {nome_assessor}")
        except Exception as e:
            print(f"[FALHA] Erro ao enviar para {nome_assessor}: {e}")

print("Fim do processamento.")


