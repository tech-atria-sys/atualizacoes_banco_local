from parametros import Parametros
import pandas as pd
import sys
sys.path.insert(0, r"C:\Scripts\modules\database")
from connection import Connect

class Bases:
    def __init__(self) -> None:
        pass

    def basebtg():
        conexao = Connect.connect_techdb()
        base_btg = Connect.import_table(conexao, "base_btg")
        conexao.close()
        return base_btg
    
    def dados_de_abertura():
        dados_abertura = pd.read_excel(Parametros.diretorio_base()+"%s\\Dados de Abertura da Conta.xlsx" % (Parametros.dia_de_hoje()), header=2)
        for columns in dados_abertura.columns:
            dados_abertura.rename(columns={columns:columns.replace(" (R$)","")}, inplace=True)
        return dados_abertura
    
    def times():
        conexao = Connect.connect_techdb()
        times_potenza = Connect.import_table(conexao, "times")
        conexao.close()
        return times_potenza        

    def informacoes():
        informacoes_clientes = pd.read_excel(Parametros.diretorio_base()+"%s\\%s\\informações.xlsx" % (Parametros.dia_de_hoje(), Parametros.horario()), header=2)
        return informacoes_clientes

    def pf_crm():
        pessoa_fisica = pd.read_excel(Parametros.diretorio_base()+"ligreufor\\Pessoa_Física_Zoho_CRM_.xlsx")
        return pessoa_fisica