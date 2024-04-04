import pandas as pd

class Leitor_excel:
    def __init__(self, arquivo_excel):
        self.arquivo_excel = arquivo_excel

    def format_dados_nfts(self, colunas_desejadas, novos_titulos):
        df = pd.read_excel(self.arquivo_excel, sheet_name='Dados_NF', usecols=colunas_desejadas, names=novos_titulos)
        return df