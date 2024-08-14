import pandas as pd

base_dados = pd.read_excel('Base_de_dados/Produtos.xlsx')#Le a base de dados
orders = pd.read_excel('orders_formatada.xlsx')#Le a planilha - ISSO PRECISA SER SUBSTITUIDO PELA FUNCAO DE UPLOAD NO SITE(APENAS PARA TESTE NO MOMENTO)

Base_dados_ajustada = base_dados.rename(columns={'REFERENCIA': 'ITEM CODE'})# Renomear a coluna 'REFERENCIA' na base de dados para 'ITEM CODE'

Colunas_Nomes_filtrados = ['Order number:', 'Delivery Date' , 'LINE', 'ITEM CODE', 'ITEM', 'ITEM PRICE', 'U. M.', 'ORDERED QUANTITY']#Os valores que eu busco

Ajuste = orders[Colunas_Nomes_filtrados].rename(columns={'ITEM': 'Nome'})#Altera os nomes originais da coluna para qual eu quiser.

base_dados_unica = Base_dados_ajustada.drop_duplicates(subset=['ITEM CODE'], keep='first')# O heroi, remove o problema de duplicacao que eu tinha
#Aqui faço o merge, para juntar o bd com o item code
dados_combinados = pd.merge(
    Ajuste,
    base_dados_unica[['ITEM CODE', 'ID_PRODUTO', 'PESO_CAIXA']],
    on='ITEM CODE',
    how='left'
)

dados_combinados = dados_combinados.rename(columns={'ID_PRODUTO': 'ID'})#Renomeia a coluna para ID

colunas_final = ['Order number:', 'Delivery Date' , 'LINE', 'ID', 'ITEM CODE', 'Nome', 'ITEM PRICE', 'U. M.', 'ORDERED QUANTITY', 'PESO_CAIXA']#Apenas as colunas que eu quero que apareca

dados_combinados[['Order number:', 'Delivery Date']] = dados_combinados[['Order number:', 'Delivery Date']].drop_duplicates()#Aqui eu apenas pego esses dados e tiro a duplicata deles, estava enfrentando problemas 
dados_combinados['LINE'] = range(1, len(dados_combinados) + 1)#Faz que a parte de linha fique de forma numerica.

#essa linhas estao gerando varias datas || dados_combinados['Delivery Date'] = dados_combinados ['Delivery Date'].dt.date.iloc[0]#Tira o horario das datas
#essa linhas estao gerando varias datas || dados_combinados['Delivery Date'] = dados_combinados ['Delivery Date'].apply(lambda x: x.strftime('%d/%m/%Y'))#Define o formato pt-br

dados_combinados = dados_combinados[colunas_final]#Junto tudo


print(dados_combinados)

dados_combinados.to_excel('Ajuste.xlsx', index=False)#Imprime

#faca todo o codigo e depois a gente passa ele para o flaske e adciona a logica para apenas receber a planilha.



''''o merge esta duplicando as coisas resultando em erro'''