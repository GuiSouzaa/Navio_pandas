from flask import Flask, render_template, request, redirect, url_for, send_file
import pandas as pd
import io

app = Flask(__name__)
app.secret_key = 'sua_chave_secreta'  # Defina uma chave secreta para sessões, se necessário
tipo_de_planilha = {'xlsx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in tipo_de_planilha

@app.route('/', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        senha = request.form['senha']
        if senha == '1':
            return redirect(url_for('upload'))
        else:
            return render_template('login.html', mensagem='Senha incorreta')
    
    return render_template('login.html')

@app.route('/upload', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':
        if 'file' not in request.files:
            return 'Nenhum arquivo selecionado'
        file = request.files['file']
        if file.filename == '':
            return 'Nenhum arquivo selecionado'
        if file and allowed_file(file.filename):
            # Carregar e processar o arquivo em memória
            orders = pd.read_excel(file.stream)  # Le a planilha enviada
            
            base_dados = pd.read_excel('Base_de_dados/Produtos.xlsx')  # Le a base de dados
            Base_dados_ajustada = base_dados.rename(columns={'REFERENCIA': 'ITEM CODE'})  # Renomear a coluna 'REFERENCIA' na base de dados para 'ITEM CODE'
            
            Colunas_Nomes_filtrados = ['Order number:', 'Delivery Date', 'LINE', 'ITEM CODE', 'ITEM', 'ITEM PRICE', 'U. M.', 'ORDERED QUANTITY']  # Os valores que eu busco
            
            Ajuste = orders[Colunas_Nomes_filtrados].rename(columns={
                'ORDERED QUANTITY': 'QTY',
                'LINE': 'Seq'
            })  # Altera os nomes originais da coluna para qual eu quiser.
            
            base_dados_unica = Base_dados_ajustada.drop_duplicates(subset=['ITEM CODE'], keep='first')  # Remove duplicações
            
            # Aqui faço o merge, para juntar o bd com o item code
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
            })  # Renomeia as colunas
            
            dados_combinados['Cxs'] = dados_combinados['VOLUME'] / dados_combinados['Unid']  # Calcula o número de caixas
            
            # Define o tipo de pallet pelo número de caixas
            def tipo_pallet(maior):
                if maior['Cxs'] >= 70:
                    return 'E'
                else:
                    return 'P'
            
            dados_combinados['Tipo'] = dados_combinados.apply(tipo_pallet, axis=1)
            
            # Verifica a capacidade do tipo de pallet
            def capacidade_pallet(Tipo):
                if Tipo == 'P':
                    return 50
                else:
                    return 70
            
            dados_combinados['Caixas pallet'] = dados_combinados['Tipo'].apply(capacidade_pallet)
            
            dados_combinados['V Pallet'] = dados_combinados['Cxs'] / dados_combinados['Caixas pallet']  # Calcula o número de pallets
            
            # Adiciona as colunas de "Sobra" e "Pallets"
            dados_combinados['Sobra'] = dados_combinados['Cxs'] % dados_combinados['Caixas pallet']
            dados_combinados['Pallets'] = dados_combinados['Cxs'] // dados_combinados['Caixas pallet']
            
            dados_combinados['Vlr estimado'] = dados_combinados['ITEM PRICE'] * dados_combinados['VOLUME']  # Valor estimado
            
            dados_combinados['Seq'] = range(1, len(dados_combinados) + 1)  # Adiciona a coluna de sequência
            
            colunas_final = ['Order number:', 'Delivery Date', 'ID', 'Nome', 'Seq', 'ITEM CODE', 'ITEM', 'ITEM PRICE', 'U. M.', 'QTY', 'VOLUME', 'Unid', 'Cxs', 'Tipo', 'Caixas pallet', 'V Pallet', 'Pallets', 'Sobra', 'Vlr estimado']  # Colunas finais
            
            dados_combinados = dados_combinados[colunas_final]  # Seleciona as colunas finais
            
            # Salva o resultado em um arquivo Excel em memória
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                dados_combinados.to_excel(writer, index=False)
            output.seek(0)
            
            # Envia o arquivo para o cliente
            return send_file(output, as_attachment=True, download_name='Ajuste.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    
    return render_template('upload.html')

if __name__ == '__main__':
    app.run(debug=True)
