##################### Importing all the necessary libraries ##############################
import pandas as pd
from datetime import datetime
import os

def format_number(x, zeros=15):
    rounded_number = round(float(x), 2)
    formatted_number = '{:.2f}'.format(rounded_number)
    padded_number = formatted_number.replace('.', '').zfill(zeros)
    return padded_number

##################### Coletting all the paths necessary ##############################

Wb_Tomadora_Path = fr'./Dados/Target_Workbook.xlsx'
Wb_Tomadora = pd.read_excel(Wb_Tomadora_Path, sheet_name='Dados_Tomadora')
Wb_Info = pd.read_excel(Wb_Tomadora_Path, sheet_name='Dados_NF')

##################### Extracting data from Cabeçalho ##############################

TR_CB = '1' # 1)
Versao = '001' # 2)
Inscricao_Mun = Wb_Tomadora.iloc[1, 1] # 3)
Data_Inicio = pd.to_datetime(Wb_Tomadora.iloc[1, 2], format='%d/%m/%Y').strftime('%Y%m%d') # 4)
Data_Fim = pd.to_datetime(Wb_Tomadora.iloc[1, 3], format='%d/%m/%Y').strftime('%Y%m%d')

##################### Getting Cabecalho in one variable ##############################

Cabeca = (
    TR_CB +
    Versao +
    str(Inscricao_Mun) +
    Data_Inicio +
    Data_Fim
)

##################### Extracting NF Info's from lines #############################

ID_pagamento = Wb_Info['ID_pagamento'].astype(str).str.replace('.', '')

TR_NF = '4' # 1)
Tipo_Documento = '02' # 2)
Serie_NF = '0' * 5 # 3)
Numero_NF = Wb_Info['Numero_NF'].astype(str).str.zfill(12) # 4)
Data_NF = str(pd.to_datetime(Wb_Info['Data_Prestação'], format='%d/%m/%Y').dt.strftime('%Y%m%d')) # 5)

Situacao_NF = 'N' # 6)
# N = Normal
# C = Cancelado

Tributacao = Wb_Info['Tributacao'] # 7)
# T - Operação normal
# I - Imune
# J – ISS Suspenso por Decisão Judicial

# Treating values
Valor_Servico = Wb_Info['Valor_Servico'].apply(format_number, zeros=15) # 8) - Using the def we created to format the values in R$
Valor_Deducoes = Wb_Info['Valor_Deducoes'].apply(format_number, zeros=15) # 9)

Cod_Serv = Wb_Info['Cod_Serv'].astype(str).str.zfill(5) # 10
Cod_Subitem = Wb_Info['Cod_Subitem'].astype(str).str.replace('.', '').str.zfill(4) # 11)
Aliquota = Wb_Info['Aliquota'].astype(str).str.replace(',', '').str.replace('%', '').str.zfill(4) # 12)

ISS_Retido = str(Wb_Info['ISS_Retido']) # 13)
# 1 – ISS Retido pelo tomador.
# 2 – NFTS sem ISS Retido.
# 3 – ISS Retido pelo intermediário.
# 4 – ISS Retido pelo tomador (descumprimento do Art. 8º A, §1º, da Lei Complementar 116, de 31 de julho de 2003)
# 5 – ISS Retido pelo intermediário (descumprimento do Art. 8º A, §1º, da Lei Complementar 116, de 31 de julho de 2003)

Indica_CNPJ= '2' # 14)
# 1 para CPF.
# 2 para CNPJ.
# 3 para Prestador estabelecido no exterior.

CNPJ = Wb_Info['CNPJ'].astype(str).str.replace('-', '').str.replace('.', '').str.replace('/', '') # 15)

CCM_Prestador = '0' * 8 # 16)
Razao_social = '0' * 75 # 17)

# Sequency to complete all the empty spaces in the code
Endereco = ' ' * 123 # 18,19,20,21,22)

Cidade_Prestador = Wb_Info['Cidade_Prestador'].str.rjust(50) # 23)
UF_Prestador = Wb_Info['UF_Prestador'] # 24
CEP_Prestador = Wb_Info['CEP_Prestador'].astype(str).str.replace('-', '') # 25
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
    ISS_Retido +
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

TR_RP = '9'

Wb_Info_Value = str(len(Wb_Info['ID_pagamento'])).zfill(7) # Contador de linhas

Soma_Valor_Deducoes = float(Wb_Info['Valor_Deducoes'].sum())
Soma_Valor_Deducoes_format = '{:016.2f}'.format(round(Soma_Valor_Deducoes, 2)).replace('.', '').replace(',', '') # Valor da coluna Deduções somados

Soma_Valor_Servico = float(Wb_Info['Valor_Servico'].sum())
Soma_Valor_Servico_format = '{:016.2f}'.format(round(Soma_Valor_Servico, 2)).replace('.', '').replace(',', '') # Valor da coluna Serviços somados

##################### Getting Rodapé Infos in one variable ##############################

Rodape = (
    TR_RP +
    Wb_Info_Value +
    Soma_Valor_Deducoes_format +
    Soma_Valor_Servico_format
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
