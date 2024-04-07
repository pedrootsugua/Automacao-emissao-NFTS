#!/usr/bin/env python
# coding: utf-8
##################### Importing all the necessary libraries ##############################
import pandas as pd
from datetime import datetime
import os

def format_number_sum(x, zeros=15):
    # Remover pontos e substituir vírgulas por pontos para garantir que o número seja interpretado corretamente
    padded_number = x.replace('.', '').replace(',', '.')
    # Converter para float, arredondar para duas casas decimais e formatar como string
    rounded_number = '{:.2f}'.format(round(float(padded_number), 2))
    # Preencher com zeros à esquerda até atingir a quantidade de zeros especificada
    stringed_number = rounded_number.replace('.', '').zfill(zeros)
    return stringed_number

def formatar_coluna_texto(df, coluna, novo_nome, comprimento, tipo=None):
    """
    Formata uma coluna existente no DataFrame.

    Parâmetros:
        df (DataFrame): O DataFrame onde a coluna está presente.
        coluna (str): O nome da coluna a ser formatada.
        novo_nome (str): O nome para a nova coluna formatada.
        comprimento (int): O comprimento desejado da string, para preencher com espaços.
        
    Retorna:
        DataFrame: O DataFrame atualizado com a nova coluna formatada.
    """
    nova_coluna = []
    for valor in df[coluna]:
        if pd.notna(valor):
            # Insira aqui as operações de formatação desejadas
            if tipo == None:
                valor_str = str(valor)
            else:
                valor_str = str(tipo(valor))  # Converter para tipo e depois para string
            valor_str = valor_str.replace('.' , '').replace(',' , '').replace('/' , '').replace('-' , '')
            valor_formatado = valor_str.ljust(comprimento)
            nova_coluna.append(valor_formatado)
        else:
            nova_coluna.append(' ' * comprimento)  # Preencher com espaços
    df[novo_nome] = nova_coluna  # Adiciona a nova coluna ao DataFrame com o nome especificado
    return df

def formatar_coluna_valor(df, coluna, novo_nome, comprimento, tipo=None):
    """
    Formata uma coluna existente no DataFrame.

    Parâmetros:
        df (DataFrame): O DataFrame onde a coluna está presente.
        coluna (str): O nome da coluna a ser formatada.
        novo_nome (str): O nome para a nova coluna formatada.
        comprimento (int): O comprimento desejado da string, para preencher com espaços.
        
    Retorna:
        DataFrame: O DataFrame atualizado com a nova coluna formatada.
    """
    nova_coluna = []
    for valor in df[coluna]:
        if pd.notna(valor):
            # Insira aqui as operações de formatação desejadas
            if tipo == None:
                valor_str = str(valor)
            else:
                valor_str = str(tipo(valor))  # Converter para tipo e depois para string
            valor_str = valor_str.replace('.' , '').replace(',' , '').replace('/' , '').replace('-' , '')
            valor_formatado = valor_str.zfill(comprimento)
            nova_coluna.append(valor_formatado)
        else:
            nova_coluna.append('0' * comprimento)  # Preencher com zeros
    df[novo_nome] = nova_coluna  # Adiciona a nova coluna ao DataFrame com o nome especificado
    return df

def formatar_coluna_ccm(df, coluna, novo_nome, comprimento, tipo=None):
    """
    Formata uma coluna existente no DataFrame.

    Parâmetros:
        df (DataFrame): O DataFrame onde a coluna está presente.
        coluna (str): O nome da coluna a ser formatada.
        novo_nome (str): O nome para a nova coluna formatada.
        comprimento (int): O comprimento desejado da string, para preencher com espaços.
        
    Retorna:
        DataFrame: O DataFrame atualizado com a nova coluna formatada.
    """
    nova_coluna = []
    for valor in df[coluna]:
        if pd.notna(valor):
            # Insira aqui as operações de formatação desejadas
            if tipo == None:
                valor_str = str(valor)
            else:
                valor_str = str(tipo(valor))  # Converter para tipo e depois para string
            valor_str = valor_str.replace('.' , '').replace(',' , '').replace('/' , '').replace('-' , '')
            valor_formatado = valor_str.zfill(comprimento)
            nova_coluna.append(valor_formatado)
        else:
            nova_coluna.append(' ' * comprimento)  # Preencher com zeros
    df[novo_nome] = nova_coluna  # Adiciona a nova coluna ao DataFrame com o nome especificado
    return df

##################### Coletting all the paths necessary ##############################
Path = fr'./Dados/NFTS_db.xlsx'
Wb_Tomadora = pd.read_excel(Path, sheet_name='Dados_Tomadora')
Wb_Info = pd.read_excel(Path, sheet_name='Dados_NF')

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

ID_pagamento = Wb_Info['ID_pagamento'].astype(str).str.replace('.', '')

TR_NF = '4' # 1)
Tipo_Documento = '02' # 2)
Serie_NF = ' ' * 5 # 3)
Numero_NF = Wb_Info['Numero_NF'].astype(str).str.zfill(12) # 4)

Data_NF = pd.to_datetime(Wb_Info['Data_Prestação'], format='%d/%m/%Y').dt.strftime('%Y%m%d') # 5)

Situacao_NF = 'N' # 6)
# N = Normal
# C = Cancelado

Tributacao = Wb_Info['Tributacao'] # 7)
# T - Operação normal
# I - Imune
# J – ISS Suspenso por Decisão Judicial

# Listas para armazenar os valores formatados de serviços e deduções
valores_servicos_formatados = []
valores_deducoes_formatados = []

# Itera sobre os valores de serviços e formata cada um individualmente
for valor_servico in Wb_Info['Valor_Servico']:
    valor_formatado = valor_servico.replace('.', '').replace(',', '.')  # Remove pontos de milhar e substitui vírgulas por pontos
    valor_formatado = '{:.2f}'.format(float(valor_formatado))  # Formata para duas casas decimais
    valor_formatado = valor_formatado.replace('.', '').zfill(15)  # Remove o ponto decimal e preenche com zeros à esquerda
    valores_servicos_formatados.append(valor_formatado)  # Adiciona o valor formatado à lista de serviços

# Itera sobre os valores de deduções e formata cada um individualmente
for valor_deducao in Wb_Info['Valor_Deducoes']:
    valor_formatado = valor_deducao.replace('.', '').replace(',', '.')  # Remove pontos de milhar e substitui vírgulas por pontos
    valor_formatado = '{:.2f}'.format(float(valor_formatado))  # Formata para duas casas decimais
    valor_formatado = valor_formatado.replace('.', '').zfill(15)  # Remove o ponto decimal e preenche com zeros à esquerda
    valores_deducoes_formatados.append(valor_formatado)  # Adiciona o valor formatado à lista de deduções


formatar_coluna_valor(Wb_Info, 'Cod_Serv', 'Cod_Serv_str', 5)
Cod_Serv = Wb_Info['Cod_Serv_str'] # 10

formatar_coluna_valor(Wb_Info, 'Cod_Subitem', 'Cod_Subitem_str', 4)
Cod_Subitem = Wb_Info['Cod_Subitem_str']  # 11)

formatar_coluna_valor(Wb_Info, 'Aliquota', 'Aliquota_str', 4)
Aliquota = Wb_Info['Aliquota_str'] # 12)

formatar_coluna_valor(Wb_Info, 'ISS_Retido', 'ISS_str', 1, tipo=None)
ISS = Wb_Info['ISS_str']  # 13)
# 1 – ISS Retido pelo tomador.
# 2 – NFTS sem ISS Retido.
# 3 – ISS Retido pelo intermediário.
# 4 – ISS Retido pelo tomador (descumprimento do Art. 8º A, §1º, da Lei Complementar 116, de 31 de julho de 2003)
# 5 – ISS Retido pelo intermediário (descumprimento do Art. 8º A, §1º, da Lei Complementar 116, de 31 de julho de 2003)

Indica_CNPJ= '2' # 14)
# 1 para CPF.
# 2 para CNPJ.
# 3 para Prestador estabelecido no exterior.

formatar_coluna_valor(Wb_Info, 'CNPJ', 'CNPJ_str', 14, tipo=None)
CNPJ = Wb_Info['CNPJ_str'] # 15)


formatar_coluna_ccm(Wb_Info, 'CCM', 'CCM_str', 8, tipo=None)
CCM_Prestador = Wb_Info['CCM_str']# 16)

formatar_coluna_texto(Wb_Info, 'Razao_Social', 'Razao_Social_str', 75)
Razao_social = Wb_Info['Razao_Social_str'] # 17)

# Endereco
for x in CCM_Prestador:
    if x == '        ':
        formatar_coluna_texto(Wb_Info, 'Tipo_Endereco', 'Tipo_Endereco_str', 3)
        TP_End = Wb_Info['Tipo_Endereco_str'] # 18
        
        formatar_coluna_texto(Wb_Info, 'Endereco', 'Endereco_str', 50)
        Endereco = Wb_Info['Endereco_str'] # 19
        
        formatar_coluna_texto(Wb_Info, 'Numero', 'Numero_str', 10)
        N_End = Wb_Info['Numero_str'] # 20
        
        formatar_coluna_texto(Wb_Info, 'Complemento', 'Complemento_str', 30)
        Complemento = Wb_Info['Complemento_str'] # 21
        
        formatar_coluna_texto(Wb_Info, 'Bairro', 'Bairro_str', 30)
        Bairro = Wb_Info['Bairro_str'] # 22
        
    else:
        TP_End = ' ' * 3
        Endereco = ' ' * 50
        N_End = ' ' * 10
        Complemento = ' ' * 30
        Bairro = ' ' * 30

formatar_coluna_texto(Wb_Info, 'Cidade_Prestador', 'Cidade_Prestador_str', 50)
Cidade_Prestador = Wb_Info['Cidade_Prestador_str'] # 23)

formatar_coluna_texto(Wb_Info, 'UF_Prestador', 'UF_Prestador_str', 2)
UF_Prestador = Wb_Info['UF_Prestador_str'] # 24

formatar_coluna_valor(Wb_Info, 'CEP_Prestador', 'CEP_Prestador_str', 8)
CEP_Prestador = Wb_Info['CEP_Prestador_str'] # 25

Email = ' ' * 75 # 26

Tipo_NFTS = '1' # 27
# 1 - Nota Fiscal do Tomador;
# 2 - Nota Fiscal do Intermediário.

Regime = '0' # 28
# 0 – Normal ou Simples Nacional (DAMSP);
# 4 – Simples Nacional (DAS);
# 5 – Microempreendedor Individual - MEI.

Dt_Pgto = ' ' * 8 # 29

Descriminacao = Wb_Info['Descriminacao'].astype(str) # 30

##################### Getting NF Info's in one variable ##############################

Info_nf = (
    TR_NF +
    Tipo_Documento +
    Serie_NF +
    Numero_NF +
    Data_NF +
    Situacao_NF +
    Tributacao +
    valores_servicos_formatados +
    valores_deducoes_formatados +
    Cod_Serv +
    Cod_Subitem +
    Aliquota +
    ISS +
    Indica_CNPJ +
    CNPJ +
    CCM_Prestador +
    Razao_social +
    TP_End +
    Endereco +
    N_End +
    Complemento +
    Bairro +
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

soma_deducoes = 0
soma_servicos = 0

TR_RP = '9'

N_Linhas = str(Wb_Info.shape[0]).zfill(7) # Contador de linhas

for valor_servicos in Wb_Info['Valor_Servico']:
    # Remove a vírgula e converte para ponto flutuante
    valor_servicos = valor_servicos.replace('.', '').replace(',', '.')
    valor_servicos = float(valor_servicos)
    # Adiciona o valor convertido à soma total
    soma_servicos += valor_servicos
# Formata a soma das deduções como valor monetário brasileiro
soma_servicos_formatada = '{:.2f}'.format(soma_servicos)
soma_servicos_formatada = soma_servicos_formatada.replace('.', '')
soma_servicos_formatada = soma_servicos_formatada.zfill(15)

# Itera sobre os valores de deduções
for valor_deducao in Wb_Info['Valor_Deducoes']:
    # Remove a vírgula e converte para ponto flutuante
    valor_deducao = valor_deducao.replace('.', '').replace(',', '.')
    valor_deducao = float(valor_deducao)
    # Adiciona o valor convertido à soma total
    soma_deducoes += valor_deducao
# Formata a soma das deduções como valor monetário brasileiro
soma_deducoes_formatada = '{:.2f}'.format(soma_deducoes)
soma_deducoes_formatada = soma_deducoes_formatada.replace('.', '')
soma_deducoes_formatada = soma_deducoes_formatada.zfill(15)

##################### Getting Rodapé Infos in one variable ##############################
Rodape = (
    TR_RP +
    N_Linhas +
    soma_servicos_formatada +
    soma_deducoes_formatada
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
