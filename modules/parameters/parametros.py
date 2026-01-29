import pandas as pd
from datetime import datetime
import os

class Parametros:
    def __init__(self) -> None:
        pass

    def diretorio_base():
        """
        Retorna o diretório base onde os relatórios
        estão localizados.
        """
        return "C:\\Scripts\\scripts_python\\relatórios\\"

    def nomes_dos_clientes():
        """
        Retorna o nome completo dos clientes.
        """
        clientes_nomes = pd.read_excel(r"C:\\Scripts\\scripts_python\\nomes_clientes\\Nomes_Clientes2.xlsx")
        clientes_nomes.rename(columns={"Código":"Conta", "Nome":"Nome Completo"}, inplace=True)
        return clientes_nomes

    def dia_de_hoje():
        """
        Retorna o a data de hoje no formato
        dia-mes-ano.
        """
        data_hoje = data = datetime.today().strftime("%d-%m-%Y")
        return data_hoje

    def horario():
        """
        Retorna o horário em que a pasta dos
        relatórios do dia foi criada.
        """
        today_hour = os.listdir(Parametros.diretorio_base()+Parametros.dia_de_hoje())[0]
        return today_hour