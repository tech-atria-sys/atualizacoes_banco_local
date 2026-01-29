from sqlalchemy import create_engine, text
import pandas as pd


class Connect:
    """
    Conectar ao banco de dados.
    """
    def __init__(self) -> None:
        pass

    def connect_techdb():
        conn_string = 'postgresql://postgres:Atria202501@localhost:5432/postgres'
        engine = create_engine(conn_string)
        conn = engine.connect()
        return conn
    
    def import_table(active_connection, table_name):
        # Build query string and ensure it's executed as a SQLAlchemy TextClause
        query_str = f'SELECT * FROM principal."{table_name}";'
        stmt = text(query_str)
        result = active_connection.execute(stmt)

        # Try to build DataFrame from result; support both fetchall()/keys() and mappings().all()
        try:
            rows = result.fetchall()
            cols = result.keys()
            df = pd.DataFrame(rows, columns=cols)
        except Exception:
            try:
                rows = result.mappings().all()
                df = pd.DataFrame(rows)
            except Exception:
                # Fallback: attempt to coerce the result directly
                df = pd.DataFrame(result)

        return df
    