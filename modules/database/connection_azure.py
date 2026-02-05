from sqlalchemy import create_engine, text
import pandas as pd
import urllib.parse  
import os
from dotenv import load_dotenv
from pathlib import Path

caminho_atual = Path(__file__).resolve()
caminho_env = caminho_atual.parents[2] / '.env'
# Isso procura o .env na pasta raiz do projeto automaticamente
load_dotenv(dotenv_path=caminho_env, override=True) 

# ... resto do código

class Connect:
    """
    Conectar ao banco de dados Azure SQL.
    """
    def __init__(self) -> None:
        self.server = os.getenv('SERVER_NAME')
        self.database = os.getenv('DATABASE_NAME')
        self.username = os.getenv('USERNAME')
        self.password = os.getenv('PASSWORD')
        self.driver = 'ODBC Driver 17 for SQL Server' 

    def connect_techdb(self):
        # Codifica a senha para evitar erros com caracteres especiais (@, /, etc)
        password_encoded = urllib.parse.quote_plus(self.password)
        
        # String de conexão para Azure SQL (MSSQL)
        conn_string = (
            f"mssql+pyodbc://{self.username}:{password_encoded}@"
            f"{self.server}/{self.database}?"
            f"driver={self.driver.replace(' ', '+')}"
        )
        
        engine = create_engine(conn_string)
        conn = engine.connect()
        return conn
    
    def import_table(self, active_connection, table_name, schema='dbo'):
        """
        Lê uma tabela do banco e retorna um DataFrame.
        Padrão do Azure é schema='dbo', mude se criou outro.
        """
        # Nota: No SQL Server usa-se colchetes ou aspas duplas, mas aspas funcionam se QUOTED_IDENTIFIER estiver ON.
        # Ajustei para f-string mais segura
        query_str = f'SELECT * FROM {schema}."{table_name}";'
        stmt = text(query_str)
        
        # Execução
        result = active_connection.execute(stmt)

        # Lógica de conversão mantida, mas simplificada pois pandas lê direto do SQLAlchemy result
        try:
            df = pd.DataFrame(result.fetchall(), columns=result.keys())
        except Exception as e:
            print(f"Erro ao converter para DataFrame: {e}")
            df = pd.DataFrame() # Retorna vazio em caso de erro

        return df


    # ... (seu código da classe acima continua igual)

if __name__ == "__main__": # Boa prática para garantir que só rode se for o arquivo principal
    try:
        print("Tentando conectar...")
        db = Connect()
        conn = db.connect_techdb()
        print("Sucesso! Conexão estabelecida com o Azure.")
        
        # Teste importar uma tabela (substitua 'NomeDaSuaTabela' por uma real)
        # df = db.import_table(conn, 'NomeDaSuaTabela')
        # print(df.head())
        
        conn.close()
        print("Conexão fechada.")
        
    except Exception as e:
        print(f"Erro ao conectar: {e}")
