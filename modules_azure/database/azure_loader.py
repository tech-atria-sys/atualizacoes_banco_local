import pandas as pd
from sqlalchemy import text
import math
from connection_azure import Connect 

class AzureLoader:
    """
    Classe responsável por todas as operações de escrita e leitura
    no Banco de Dados Azure SQL.
    Otimizada para Standard S0 e limites do SQL Server.
    """
    
    @staticmethod
    def enviar_df(df, nome_tabela, col_pk=None, if_exists='replace', schema='dbo'):
        if df.empty:
            print(f"[AzureLoader] Aviso: DataFrame '{nome_tabela}' está vazio.")
            return

        print(f"[AzureLoader] Subindo tabela: {schema}.{nome_tabela} ({len(df)} linhas)...")
        
        # --- CÁLCULO DE SEGURANÇA DUPLO ---
        # Regra 1: Max 2100 parâmetros (para tabelas com muitas colunas)
        # Regra 2: Max 1000 linhas por INSERT (para tabelas com poucas colunas)
        
        num_colunas = len(df.columns)
        if num_colunas > 0:
            limit_params = math.floor(2090 / num_colunas)
            limit_rows = 1000 # Limite hard-coded do SQL Server
            
            # O chunksize deve ser o MENOR dos dois limites
            safe_chunksize = min(limit_params, limit_rows)
        else:
            safe_chunksize = 1000
        
        safe_chunksize = max(1, safe_chunksize) # Garante mínimo de 1
        
        print(f"[AzureLoader] Chunksize calculado: {safe_chunksize} linhas/lote (Colunas: {num_colunas}).")

        db = Connect()
        conn = db.connect_techdb()
        
        try:
            df.to_sql(
                name=nome_tabela,
                con=conn,
                schema=schema,
                if_exists=if_exists,
                index=False,
                chunksize=safe_chunksize, 
                method='multi' 
            )
            
            # Criação de PK (apenas se for replace)
            if col_pk and if_exists == 'replace':
                print(f"[AzureLoader] Configurando Primary Key: {col_pk}")
                conn.execute(text(f'ALTER TABLE {schema}."{nome_tabela}" ALTER COLUMN "{col_pk}" VARCHAR(450) NOT NULL'))
                conn.execute(text(f'ALTER TABLE {schema}."{nome_tabela}" ADD PRIMARY KEY ("{col_pk}")'))
                conn.commit() 
                
            print(f"[AzureLoader] Concluido: {nome_tabela} atualizada.")
            
        except Exception as e:
            print(f"[AzureLoader] ERRO em {nome_tabela}: {e}")
            raise e
        finally:
            conn.close()

    @staticmethod
    def ler_tabela(nome_tabela, schema='dbo'):
        db = Connect()
        conn = db.connect_techdb()
        try:
            print(f"[AzureLoader] Lendo: {schema}.{nome_tabela}...")
            df = pd.read_sql_table(nome_tabela, conn, schema=schema)
            return df
        except Exception as e:
            print(f"[AzureLoader] Erro ao ler ou tabela inexistente: {e}")
            return pd.DataFrame()
        finally:
            conn.close()