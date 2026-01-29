import shutil
import os
from datetime import datetime

now=datetime.today()
ano=str(now.year)

if len(str(now.day)) == 1:
    dia='0'+str(now.day)
else:
    dia=str(now.day)
    
if len(str(now.month)) == 1:
    mes='0'+str(now.month)
else:
    mes=str(now.month)
    
if len(str(now.hour)) == 1:
    hora='0'+str(now.hour)
else:
    hora=str(now.hour)
    
if len(str(now.minute)) == 1:
    minuto='0'+str(now.minute)
else:
    minuto=str(now.minute)

tempo=hora+minuto
data=dia+'-'+mes+'-'+ano

renomes={'Evolução do PL':'evolução_pl',
         'NNM':'evolução_nnm',
         'Saldo em CC (D 0)':'saldo_cc',
         'data':'movimentação',
         'Net New Money Detalhado':'nnm',
         'data (1)':'posição',
         'Valor em Trânsito':'valor_em_transito',
         'Histórico de Operações':'operações',
         'Operações do Dia':'operações_do_dia',
         'Subscrição':'subscrição',
         'data (2)':'ativação_de_mercado',
         'data (3)':'base_btg',
         'Informações':'informações',
         'data (4)':'rentabilidade',
         'POSIÇÕES':'fundos_posição',
         'MOVIMENTAÇÕES':'fundos',
         'data (5)':'ipo',
         'PREVIDÊNCIA':'previdência',
         'OPÇÕES FLEXÍVEIS':'opções_flexíveis',
         'POSIÇÕES (1)':'renda_fixa_coe',
         'POSIÇÕES (2)':'renda_variável',
         'MOVIMENTAÇÕES (1)':'custódia',
          }

for i in renomes:
    os.rename(i+'.xlsx', renomes[i]+'.xlsx')

try:
    os.mkdir('C:/Users/Potenza Capital/Zoho WorkDrive (Potenza Capital)/POTENZA TECH/relatórios/'+data)
except Exception:
    pass

try:
    os.mkdir('C:/Users/Potenza Capital/Zoho WorkDrive (Potenza Capital)/POTENZA TECH/relatórios/'+data+'/'+tempo)
except Exception:
    pass


for j in renomes.values():
    file=j+'.xlsx'
    os.replace(file, 'C:/Users/Potenza Capital/Zoho WorkDrive (Potenza Capital)/POTENZA TECH/relatórios/'+data+'/'+tempo+'/'+file)