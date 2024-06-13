from flask import Flask, request, jsonify
import pandas as pd
import os

app = Flask(__name__)

@app.route('/')
def index():
    return open('index.html').read()

@app.route('/processar', methods=['POST'])
def processar_arquivo():
    try:
        arquivo_excel = request.files['excel_file']

        if arquivo_excel.filename == '':
            return 'Nenhum arquivo selecionado.'

        # Salvar o arquivo temporariamente
        arquivo_temp = 'temp.xlsx'
        arquivo_excel.save(arquivo_temp)

        # Ler o arquivo Excel
        xls = pd.ExcelFile(arquivo_temp)

        # Iterar sobre cada aba no arquivo Excel
        for aba in xls.sheet_names:
            df = pd.read_excel(arquivo_temp, sheet_name=aba)

            # Remover itens duplicados
            df_sem_duplicatas = df.drop_duplicates()

            # Salvar a aba sem duplicatas no novo arquivo Excel
            arquivo_saida = f"{aba}-semduplicatas.xlsx"
            df_sem_duplicatas.to_excel(arquivo_saida, index=False)

        # Remover arquivo temporário
        os.remove(arquivo_temp)

        return 'Arquivo processado com sucesso.'
    except Exception as e:
        return str(e), 500

if __name__ == "__main__":
    app.run(debug=True)
