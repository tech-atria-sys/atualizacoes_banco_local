import sys
sys.path.insert(0, '/Users/Potenza Capital/Zoho WorkDrive (Potenza Capital)/POTENZA TECH/Scripts de rotina/arquivos_banco/modules/parameters')
from apis import dadosFinanceiros
sys.path.insert(0, '/Users/Potenza Capital/Zoho WorkDrive (Potenza Capital)/POTENZA TECH/Scripts de rotina/arquivos_banco/modules/database')
from connection import Connect

conn = Connect.connect_techdb()

cdi = dadosFinanceiros.consulta_bc(12)
cdi_mensal = cdi.resample("M").sum()
cdi_mensal['valor'] = cdi_mensal['valor']/100
cdi_mensal['CDI ACUMULADO'] = cdi_mensal.groupby(cdi_mensal.index.year)['valor'].cumsum()
cdi_mensal = cdi_mensal["2021":]
cdi_mensal.reset_index(inplace=True)
cdi_mensal.rename(columns={"valor":"CDI", "data":"DATA"}, inplace=True)
#cdi_mensal.to_sql('CDI', index=False, con=conn, if_exists='replace')