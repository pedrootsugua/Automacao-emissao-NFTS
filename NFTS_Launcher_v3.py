#!/usr/bin/env python
# coding: utf-8
##################### Importing all the necessary libraries ##############################
import pandas as pd
from datetime import datetime
import os

def formatar(valor):
    return "{:.2f}".format(valor)

##################### Coletting all the paths necessary ##############################
Path = fr'./data/NFTS_db.xlsx'
Wb_Tomadora = pd.read_excel(Path, sheet_name='Dados_Tomadora', dtype=str)
Wb_Info = pd.read_excel(Path, sheet_name='Dados_NF', dtype=str)

##################### Extracting data from Cabeçalho ##############################
TR_CB = '1' # 1)
Versao = '001' # 2)
Inscricao_Mun = str(Wb_Tomadora.iloc[1, 1]) # 3)
Data_Inicio = pd.to_datetime(Wb_Tomadora.iloc[1, 2], format='%d/%m/%Y').strftime('%Y%m%d') # 4)
Data_Fim = pd.to_datetime(Wb_Tomadora.iloc[1, 3], format='%d/%m/%Y').strftime('%Y%m%d')

##################### Getting Cabecalho in one variable ##############################

Cabeca = (
    TR_CB +
    Versao +
    Inscricao_Mun +
    Data_Inicio +
    Data_Fim
)

##################### Extracting NF Info's from lines #############################
ID_pagamento = Wb_Info['ID_pagamento'].str.replace('.', '')

TR_NF = '4' # 1)
Tipo_Documento = '02' # 2)
Serie_NF = ' ' * 5 # 3)
Numero_NF = Wb_Info['Numero_NF'].astype(str).str.zfill(12) # 4)


Data_NF = pd.to_datetime(Wb_Info['Data_Prestação']).dt.strftime('%Y-%d-%m').str.replace('-','') # 5)

Situacao_NF = 'N' # 6)
# N = Normal
# C = Cancelado

Tributacao = Wb_Info['Tributacao'] # 7)
# T - Operação normal
# I - Imune
# J – ISS Suspenso por Decisão Judicial

Valor_Servico = Wb_Info['Valor_Servico'].astype(float).apply(formatar).str.replace('.','').str.zfill(15) # 8)

Valor_Deducoes = Wb_Info['Valor_Deducoes'].astype(float).apply(formatar).str.replace('.','').str.zfill(15) # 9)

Cod_Serv = Wb_Info['Cod_Serv'].astype(str).str.zfill(5) # 10

Cod_Subitem = Wb_Info['Cod_Subitem'].astype(str).str.replace('.','').str.zfill(4)  # 11)

Aliquota = Wb_Info['Aliquota'].astype(float).apply(formatar).str.replace('.','').str.zfill(4) # 12)

ISS = Wb_Info['ISS_Retido'] # 13)
# 1 – ISS Retido pelo tomador.
# 2 – NFTS sem ISS Retido.
# 3 – ISS Retido pelo intermediário.
# 4 – ISS Retido pelo tomador (descumprimento do Art. 8º A, §1º, da Lei Complementar 116, de 31 de julho de 2003)
# 5 – ISS Retido pelo intermediário (descumprimento do Art. 8º A, §1º, da Lei Complementar 116, de 31 de julho de 2003)


Indica_CNPJ= '2' # 14)
# 1 para CPF.
# 2 para CNPJ.
# 3 para Prestador estabelecido no exterior.

CNPJ = Wb_Info['CNPJ'].astype(str).str.replace('.','').str.replace('/','').str.replace('-','').str.zfill(14) # 15)

CCM_Prestador = ' ' * 8 # 16)

Razao_social = Wb_Info['Razao_Social'].str.ljust(75) # 17)

Endereco = ' ' * 123 # 18, 19, 20, 21, 22)

Cidade_Prestador = Wb_Info['Cidade_Prestador'].str.ljust(50) # 23)

UF_Prestador = Wb_Info['UF_Prestador'].str.ljust(2) # 24

CEP_Prestador = Wb_Info['CEP_Prestador'].astype(str).str.replace('.','').str.replace('-','').str.zfill(8) # 25)

Email = ' ' * 75 # 26

Tipo_NFTS = '1' # 27
# 1 - Nota Fiscal do Tomador;
# 2 - Nota Fiscal do Intermediário.

Regime = '0' # 28
# 0 – Normal ou Simples Nacional (DAMSP);
# 4 – Simples Nacional (DAS);
# 5 – Microempreendedor Individual - MEI.

Dt_Pgto = ' ' * 8 # 29

Descriminacao = Wb_Info['Descriminacao'] # 30

##################### Getting NF Info's in one variable ##############################

Info_nf = (
    TR_NF +
    Tipo_Documento +
    Serie_NF +
    Numero_NF +
    Data_NF +
    Situacao_NF +
    Tributacao +
    Valor_Servico +
    Valor_Deducoes +
    Cod_Serv +
    Cod_Subitem +
    Aliquota +
    ISS +
    Indica_CNPJ +
    CNPJ +
    CCM_Prestador +
    Razao_social +
    Endereco +
    Cidade_Prestador +
    UF_Prestador +
    CEP_Prestador +
    Email +
    Tipo_NFTS +
    Regime +
    Dt_Pgto +
    Descriminacao
)

##################### Extracting rodapé #############################
TR_RP = '9' # 1)
N_Linhas = str(Wb_Info.shape[0]).zfill(7) # 2) Contador de linhas

Soma_VS = Wb_Info['Valor_Servico'].astype(float).apply(formatar).str.replace('.','') # 3)
Soma_VS = Soma_VS.astype(int).sum()
Soma_VS = Soma_VS.astype(str).zfill(15)

Soma_VD = Wb_Info['Valor_Deducoes'].astype(float).apply(formatar).str.replace('.','') # 4)
Soma_VD = Soma_VD.astype(int).sum()
Soma_VD = Soma_VD.astype(str).zfill(15)

##################### Getting Rodapé Infos in one variable ##############################
Rodape = (
    TR_RP +
    N_Linhas +
    Soma_VS +
    Soma_VD
)

####################### Making .Txt file ##############################
data_atual = datetime.now()
cont = 1
nome_base = 'Lote_NFTS_' + data_atual.strftime("%Y-%m-%d")
nome_arquivo = f"{nome_base}_{cont}.txt"


while os.path.exists(nome_arquivo):
    cont += 1
    nome_arquivo = f"{nome_base}_{cont}.txt"

try:
    with open(nome_arquivo, 'w') as arquivo:
        # Escrever cabeçalho
        arquivo.write(Cabeca + '\n')

        # Escrever informações das notas fiscais
        for linha_info_nf in Info_nf:
            arquivo.write(linha_info_nf + '\n')

        # Escrever rodapé
        arquivo.write(Rodape + '\n')

        print(f"Arquivo criado {nome_arquivo} criado com sucesso!")

except Exception as e:    
    print("Ocorreu um erro ao criar o arquivo:", e)

