import pandas as pd
from datetime import datetime
import os

class Parametros:
    """
    Centraliza caminhos e datas para automação de relatórios.
    Estrutura adaptada para: C:\\Scripts\\modules_azure\\parameters\\parametros_azure.py
    """

    @staticmethod
    def diretorio_base():
        """
        Retorna a raiz onde ficam as pastas de datas.
        """
        # Ajuste aqui se sua pasta raiz for diferente
        return r"C:\Scripts\relatórios\\"

    @staticmethod
    def dia_de_hoje():
        """
        Retorna a data de hoje no formato dia-mes-ano (ex: 05-02-2026).
        """
        return datetime.today().strftime("%d-%m-%Y")

    @staticmethod
    def horario():
        """
        Retorna o nome da subpasta de horário (ex: '09h15') dentro da pasta de hoje.
        Pega sempre a primeira pasta encontrada.
        """
        caminho_do_dia = os.path.join(Parametros.diretorio_base(), Parametros.dia_de_hoje())
        
        try:
            arquivos = os.listdir(caminho_do_dia)
            # Filtra apenas o que é pasta (evita pegar arquivos soltos se houver)
            # Se só tiver pasta de horário, retorna a primeira [0]
            if arquivos:
                return arquivos[0]
            else:
                raise FileNotFoundError("Pasta do dia existe, mas está vazia.")
        except FileNotFoundError:
            print(f"⚠️ [Parametros] Atenção: Pasta {caminho_do_dia} não encontrada.")
            return None

    @staticmethod
    def caminho_completo_atual():
        """
        FACILITADOR: Retorna o caminho completo já pronto para uso.
        Ex: C:\\Scripts\\relatórios\\05-02-2026\\10h00\\
        """
        h = Parametros.horario()
        if h:
            path = os.path.join(Parametros.diretorio_base(), Parametros.dia_de_hoje(), h)
            return path + "\\" # Garante a barra no final
        return None

    @staticmethod
    def nomes_dos_clientes():
        """
        Lê e trata o Excel de nomes de clientes (De/Para).
        """
        caminho_arquivo = r"C:\Scripts\nomes_clientes\Nomes_Clientes2.xlsx"
        
        try:
            clientes_nomes = pd.read_excel(caminho_arquivo)
            # Renomeia para padronizar com o banco
            clientes_nomes.rename(columns={"Código": "Conta", "Nome": "Nome Completo"}, inplace=True)
            return clientes_nomes
        except Exception as e:
            print(f"❌ [Parametros] Erro ao ler Nomes_Clientes2.xlsx: {e}")
            return pd.DataFrame() # Retorna vazio para não quebrar tudo