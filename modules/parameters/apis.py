import pandas as pd

class dadosFinanceiros:
    def consulta_bc(codigo_bcb: int):
        url = 'https://api.bcb.gov.br/dados/serie/bcdata.sgs.{}/dados?formato=json'.format(codigo_bcb)
        df = pd.read_json(url)
        df['data'] = pd.to_datetime(df['data'], dayfirst=True)
        df.set_index('data', inplace=True)
        return df
    
    def consultaExpectativasMensais(indice: str):
        url = "https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/odata/ExpectativaMercadoMensais?$format=json&$filter=Indicador%20eq%20'{}'".format(indice)
        df = pd.read_json(url)
        df.drop('@odata.context', axis=1, inplace=True)
        return df
    
    def consultaExpectativasAnuais(indice: str):
        url = "https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/odata/ExpectativasMercadoAnuais?$format=json&$filter=Indicador%20eq%20'{}'".format(indice)
        df = pd.read_json(url)
        df.drop('@odata.context', axis=1, inplace=True)
        return df