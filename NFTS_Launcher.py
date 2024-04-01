import pandas as pd
from datetime import datetime
import os

Wb_Tomadora_Path = fr'./Dados/Target_Workbook.xlsx'
Wb_Tomadora = pd.read_excel(Wb_Tomadora_Path, sheet_name='Dados_Tomadora')
Wb_Info = pd.read_excel(Wb_Tomadora_Path, sheet_name='Dados_NF')


# ##### Extracting data from Cabeçalho


Tipo_Registro = '1'
Versao = '001'
Inscricao_Mun = Wb_Tomadora.iloc[1, 1]
Data_Inicio = pd.to_datetime(Wb_Tomadora.iloc[1, 2], format='%d/%m/%Y').strftime('%Y%m%d')
Data_Fim = pd.to_datetime(Wb_Tomadora.iloc[1, 3], format='%d/%m/%Y').strftime('%Y%m%d')

### passing everything to string type
Tipo_Registro = str(Tipo_Registro)
Versao = str(Versao)
Inscricao_Mun = str(Inscricao_Mun)
Data_Inicio = str(Data_Inicio)
Data_Fim = str(Data_Fim)

# Getting everything in one variable
Cabeca = (
    Tipo_Registro +
    Versao +
    Inscricao_Mun +
    Data_Inicio +
    Data_Fim
)


### Extracting data from NF's Details
ID_pagamento = Wb_Info['ID_pagamento'].astype(str).str.replace('.', '')

TR_NF = '4'

Tipo_Documento = '02' # Ponto importante de se fazer um if/else caso vá realizar transmissão em lote de câmbios.

Serie_NF = '0' * 5
Data_NF = pd.to_datetime(Wb_Info['Data_Prestação'], format='%d/%m/%Y').dt.strftime('%Y%m%d')
Numero_NF = Wb_Info['Numero_NF'].astype(str).str.zfill(12)
Situacao_NF = 'N'
Tributacao = Wb_Info['Tributacao']
Valor_Servico = Wb_Info['Valor_Servico'].astype(str).str.replace(',', '').str.replace('.', '').str.zfill(15)
Valor_Deducoes = Wb_Info['Valor_Deducoes'].astype(str).str.replace(',', '').str.replace('.', '').str.zfill(15)
Cod_Serv = Wb_Info['Cod_Serv'].astype(str).str.zfill(5)
Cod_Subitem = Wb_Info['Cod_Subitem'].astype(str).str.replace('.', '').str.zfill(4)
Aliquota = Wb_Info['Aliquota'].astype(str).str.replace(',', '').str.replace('%', '').str.zfill(4)
ISS_Retido = Wb_Info['ISS_Retido'].astype(str) # 1 = retido pelo tomador / 2 = Sem retenção / 4 = Retido pelo Tomador
Indica_CNPJ= '2' # Emissor é cnpj, sempre tem que ser 2
CNPJ = Wb_Info['CNPJ'].astype(str).str.replace('-', '').str.replace('.', '').str.replace('/', '')
CCM_Prestador = '0' * 8
Razao_social = '0' * 75
# sequencia pra completar
Endereco = '0'*123
Cidade_Prestador = Wb_Info['Cidade_Prestador'].str.rjust(50)
UF_Prestador = Wb_Info['UF_Prestador']
CEP_Prestador = Wb_Info['CEP_Prestador'].astype(str).str.replace('-', '')
Email = ' ' * 75
Tipo_NFTS = '1'
Regime = '0'
Dt_Pgto = '0' * 8
Descriminacao = Wb_Info['Descriminacao']

# Getting everything in one variable

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

# Extracting rodapé
Tipo_Registro_RP = '9'

#CRIAR UM COUNT PARA A QTD DE NOTAS EXISTENTES NO TXT
Wb_Info_Value = str(len(Wb_Info['ID_pagamento'])).zfill(7) #(VERIFICAR) DEVE SER INFORMADO A QTD DE LINHAS(DADOS NFE) QUE EXISTEM NO TIPO 4(LINHAS DE DETALHE)

# formatação valor deduções
Soma_Valor_Deducoes = float(Wb_Info['Valor_Deducoes'].sum())
Soma_Valor_Deducoes_format = format(Soma_Valor_Deducoes, '.2f')
Soma_Valor_Deducoes_format = Soma_Valor_Deducoes_format.replace(',','').replace('.','').zfill(15)
print(Soma_Valor_Deducoes)#VERIFICAR CALCULO DA SOMA

# formatação valor serviço
Soma_Valor_Servico = float(Wb_Info['Valor_Servico'].sum())
Soma_Valor_Servico_format = format(Soma_Valor_Servico, '.2f')
Soma_Valor_Servico_format = Soma_Valor_Servico_format.replace(',','').replace('.','').zfill(15)
print(Soma_Valor_Servico)#VERIFICAR CALCULO DA SOMA

# Passing everything to string type
Tipo_Registro_RP = str(Tipo_Registro_RP)
Wb_Info_Value = str(Wb_Info_Value)
Soma_Valor_Servico = str(Soma_Valor_Servico)
Soma_Valor_Deducoes = str(Soma_Valor_Deducoes)


# Getting everything in one variable
Rodape = (
    Tipo_Registro_RP +
    Wb_Info_Value +
    Soma_Valor_Servico +
    Soma_Valor_Deducoes
)

# Making our .Txt file
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
