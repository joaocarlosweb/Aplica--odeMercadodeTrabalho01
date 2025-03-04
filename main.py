import pandas as pd
import openpyxl
import pathlib
import funçoes as fun


email = pd.read_excel(r'Bases de Dados\Emails.xlsx')
lojas = pd.read_csv(r'Bases de Dados\Lojas.csv', encoding="latin1", sep=";")
vendas = pd.read_excel(r'Bases de Dados\Vendas.xlsx')

# print(email)
# print(vendas)
# print(loja)
# mesclar o arquivo de loja com vendas para juntar o nome da loja com o ID
vendas = vendas.merge(lojas, on='ID Loja')
# print(vendas['ID Loja'])

# Criando uma planilha para cada loja
dicionario_lojas ={}
for loja in lojas['Loja']:
    dicionario_lojas[loja] = vendas.loc[vendas['Loja']== loja, :]

# print(dicionario_lojas['Iguatemi Campinas'])

dia_indicador = vendas["Data"].max()
print(f'Onepage do dia {dia_indicador.day}/{dia_indicador.month}/{dia_indicador.year}')

caminho_backup = pathlib.Path(r'Backup Arquivos Lojas')
arquivos_pasta_backup = caminho_backup.iterdir()
lista_names_backup = [arquivo.name for arquivo in arquivos_pasta_backup]
# print(lista_names_backup)

for loja in dicionario_lojas:
    if loja not in lista_names_backup:
        nova_pasta = caminho_backup / loja
        nova_pasta.mkdir(parents=True, exist_ok=True)
    nome_arquivo = f'{dia_indicador.day}_{dia_indicador.month}_{loja}.xlsx'
    local_arquivo = caminho_backup / loja / nome_arquivo
    dicionario_lojas[loja].to_excel(local_arquivo)
    # print(dicionario_lojas[loja])

    vendas_loja = dicionario_lojas[loja]
    faturamento_ano = vendas_loja['Valor Final'].sum()
    # print(f' Faturamento anual da Loja {loja} R$:{faturamento_ano:,.2f}')
    # print('-'*40)
    vendas_loja_dia = vendas_loja.loc[vendas["Data"]== dia_indicador, :]
    faturamento_dia = vendas_loja_dia['Valor Final'].sum()
    # print(f'Faturamento do dia da loja {loja} R$:{faturamento_dia:,.2f}')
    # print('-' * 40)
    quant_produtos_ano = len(vendas_loja['Produto'].unique())
    # print(f'Quantidade de produtos vendido no ano {quant_produtos_ano}')
    # print('-' * 40)
    quant_produtos_dia = len(vendas_loja_dia['Produto'].unique())
    # print(f'Quantidade de produtos vendido por dia {quant_produtos_dia}')
    # print('-' * 40)
    vendas_loja.loc[:, 'Valor Final'] = pd.to_numeric(vendas_loja['Valor Final'], errors='coerce')
    valor_venda = vendas_loja.groupby('Código Venda')['Valor Final'].sum()
    tikte_medio_ano = valor_venda.mean()
    vendas_loja_dia.loc[:, 'Valor Final'] = pd.to_numeric(vendas_loja_dia['Valor Final'], errors='coerce')
    valor_venda_dia = vendas_loja_dia.groupby('Código Venda')['Valor Final'].sum()
    tikte_medio_dia = valor_venda_dia.mean()
    # print(f'Tikte médio por dia R$:{tikte_medio_dia:,.2f}')
    # print('-' * 40)
    # print(f'Tikte médio por ano R$:{tikte_medio_ano:,.2f}')
    print('-' * 40)
    fun.enviar_mensagem()

    # meta_faturamento_dia = 1000
    # meta_faturamento_ano = 1650000
    # meta_quant_prod_dia = 4
    # meta_quant_prod_ano = 120
    # meta_tiktemedia_dia = 500
    # meta_tiktemedia_ano = 500


