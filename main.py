from Classes.Leitor_excel import Leitor_excel

leitor_excel = Leitor_excel('./Dados/Target_Workbook.xlsx')

# Definindo as colunas desejadas e os novos titulos das colunas
colunas_desejadas = ['Numero_NF', 'Data_Prestação', 'Valor_Servico', 'Valor_Deducoes', 'Cod_Serv', 'Cidade_Prestador']
novos_titulos = ['Número NF', 'Data de Prestação', 'Valor do Serviço', 'Valor das Deduções', 'Código do Serviço', 'Cidade do Prestador']

df = leitor_excel.format_dados_nfts(colunas_desejadas, novos_titulos)
        
# Exibir o DataFrame com as colunas renomeadas
print(df)