import pandas as pd

base_dados = pd.read_excel('Base_de_dados/Produtos.xlsx')#Le a base de dados
orders = pd.read_excel('orders_formatada.xlsx')#Le a planilha - ISSO PRECISA SER SUBSTITUIDO PELA FUNCAO DE UPLOAD NO SITE(APENAS PARA TESTE NO MOMENTO)

Base_dados_ajustada = base_dados.rename(columns={'REFERENCIA': 'ITEM CODE'})# Renomear a coluna 'REFERENCIA' na base de dados para 'ITEM CODE'

Colunas_Nomes_filtrados = ['Order number:', 'Delivery Date' , 'LINE', 'ITEM CODE', 'ITEM', 'ITEM PRICE', 'U. M.', 'ORDERED QUANTITY']#Os valores que eu busco

Ajuste = orders[Colunas_Nomes_filtrados].rename(columns={
    
    'ORDERED QUANTITY': 'QTY',
    'LINE': 'Seq'
    
    })#Altera os nomes originais da coluna para qual eu quiser.

base_dados_unica = Base_dados_ajustada.drop_duplicates(subset=['ITEM CODE'], keep='first')# O heroi, remove o problema de duplicacao que eu tinha

#Aqui faço o merge, para juntar o bd com o item code
dados_combinados = pd.merge(
    Ajuste,
    base_dados_unica[['ITEM CODE', 'ID_PRODUTO', 'VOLUME', 'FANTASIA', 'PESO_CAIXA']],
    on='ITEM CODE',
    how='left'
)

dados_combinados = dados_combinados.rename(columns={
    'ID_PRODUTO': 'ID',
    'FANTASIA': 'Nome',
    'PESO_CAIXA': 'Unid'
    
})#Renomeia a coluna para ID(Aqui tudo da base de dados pode ser renomeado)

dados_combinados['Cxs'] = dados_combinados['VOLUME'] / dados_combinados['Unid']#divide o volume total (VOLUME) pelo peso de uma unidade (Unid) para determinar quantas caixas são necessárias para o pedido.

#Defini o tipo de pallet pelo numero de caixas
def tipo_pallet(maior):
    if maior['Cxs'] >= 70: #aaaaaa era so corrigir para ' > ' FIQUEI 3 HORAS REVISANDO ISSO AQUI DEUS.
        return 'E'
    else:
        return 'P'

dados_combinados['Tipo'] = dados_combinados.apply(tipo_pallet, axis=1)

#Verifica a capacidade o tipo de pallet, e defini quantas caixas cabem em cada pallet.
def capacidade_pallet(Tipo):
    if Tipo == 'P':
        return 50
    else:
        return 70

dados_combinados['Caixas pallet'] = dados_combinados['Tipo'].apply(capacidade_pallet)

dados_combinados['V Pallet'] = dados_combinados['Cxs'] / dados_combinados['Caixas pallet']#determina quantos pallets são necessários para armazenar todas as caixas. Ele divide o total de caixas (Cxs) pela capacidade de caixas do pallet (Caixas pallet).

# Adiciona a coluna de "Sobra" e "Pallets" com os valores corretos----------------------------
dados_combinados['Sobra'] = dados_combinados['Cxs'] % dados_combinados['Caixas pallet']#encontra as caixas que sobram após preencher todos os pallets inteiro.
dados_combinados['Pallets'] = dados_combinados['Cxs'] // dados_combinados['Caixas pallet']#encontra o número inteiro de pallets necessários usando divisão inteira
#---------------------------------------------------------------------------------------------

#dados_combinados[['Order number:', 'Delivery Date']] = dados_combinados[['Order number:', 'Delivery Date']].drop_duplicates()#Aqui eu apenas pego esses dados e tiro a duplicata deles, estava enfrentando problemas 
dados_combinados['Seq'] = range(1, len(dados_combinados) + 1)#Faz que a parte de linha fique de forma numerica.

colunas_final = ['Order number:', 'Delivery Date', 'ID', 'Nome', 'Seq', 'ITEM CODE', 'ITEM', 'ITEM PRICE', 'U. M.', 'QTY', 'VOLUME', 'Unid', 'Cxs', 'Tipo', 'Caixas pallet', 'V Pallet', 'Pallets', 'Sobra']#Apenas as colunas que eu quero que apareca

#essa linhas estao gerando varias datas || dados_combinados['Delivery Date'] = dados_combinados ['Delivery Date'].dt.date.iloc[0]#Tira o horario das datas
#essa linhas estao gerando varias datas || dados_combinados['Delivery Date'] = dados_combinados ['Delivery Date'].apply(lambda x: x.strftime('%d/%m/%Y'))#Define o formato pt-br

#dados_combinados['Order number:'] = dados_combinados['Order number:'].mask(dados_combinados['Order number:'].duplicated(), '')
#dados_combinados['Delivery Date'] = dados_combinados['Delivery Date'].mask(dados_combinados['Delivery Date'].duplicated(), '')

dados_combinados = dados_combinados[colunas_final]#Junto tudo

print(dados_combinados)

#dados_combinados.to_excel('Ajuste.xlsx', index=False)#Imprime
