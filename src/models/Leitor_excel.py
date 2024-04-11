import pandas as pd
import tkinter as tk
from tkinter import filedialog

class Leitor_excel:
    # def __init__(self, arquivo_excel):
    #     self.arquivo_excel = arquivo_excel

    # def format_dados_nfts(self, colunas_desejadas, novos_titulos):
    #     df = pd.read_excel(self.arquivo_excel, sheet_name='Dados_NF', usecols=colunas_desejadas, names=novos_titulos)
    #     return df
    
    # consumir os dados dos relatorios de pagamento e criar um novo df somente com as informações que serão usadas no padrão da prefeitura
    def format_dataframe_nfts(file_path, colunas, novos_nomes):
        df = pd.read_excel(file_path)
        df_dados_nfts = df.loc[:, colunas]
        df_dados_nfts = df_dados_nfts.rename(columns=novos_nomes)
        return df_dados_nfts

    # exemplo de uso - a chamada abaixo ficara na views/app
    file_path = 'caminho/do/arquivo.xlsx'
    colunas_desejadas = ['Coluna1', 'Coluna2', 'Coluna3']
    novos_nomes = {'Coluna1': 'NovoNome1', 'Coluna2': 'NovoNome2', 'Coluna3': 'NovoNome3'}
    df_info_lote_nfts = format_dataframe_nfts(file_path, colunas_desejadas, novos_nomes)

    print(df_info_lote_nfts.head())

    # permite que o usuario selecione um arquivo presente em sua maquina local ou rede
    def selecionar_arquivo():
        try:
            root = tk.Tk()
            root.withdraw()  # Esconde a janela principal

            # Abre o explorador de arquivos
            arquivo = filedialog.askopenfilename()

            # Verifica se um arquivo foi selecionado
            if arquivo:
                print("Arquivo selecionado:", arquivo)
                return arquivo
            else:
                print("Nenhum arquivo selecionado.")
                return None
        except tk.TclError as e:
            print("Erro ao abrir o diálogo de seleção de arquivo:", e)
        except FileNotFoundError as e:
            print("Arquivo não encontrado:", e)
        except PermissionError as e:
            print("Permissão negada:", e)

    # devera ser chamado na views/app
    selecionar_arquivo()