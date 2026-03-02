import os
import sys
import tempfile
import webbrowser
from threading import Timer
from flask import Flask, render_template, request, jsonify, send_file
from core.processador import gerar_tabela, salvar_excel_estilizado

# Determina o caminho base (se está rodando como script ou executável)
if getattr(sys, 'frozen', False):
    template_folder = os.path.join(sys._MEIPASS, 'templates')
    static_folder = os.path.join(sys._MEIPASS, 'static')
    app = Flask(__name__, template_folder=template_folder, static_folder=static_folder)
else:
    app = Flask(__name__)
TEMP_DIR = tempfile.gettempdir()


@app.route('/')
def index():
    """Carrega a interface principal."""
    return render_template('index.html')


@app.route('/processar', methods=['POST'])
def processar():
    """
    Processa os arquivos e gera o preview colorido para a interface
    e o arquivo Excel final para download.
    """
    try:
        xml_file = request.files.get('xml')
        csv_file = request.files.get('csv')
        params = request.form.to_dict()

        if not xml_file or not csv_file:
            return jsonify({'sucesso': False, 'erro': 'Por favor, anexe os arquivos XML e CSV.'})

        # Caminhos temporários
        xml_path = os.path.join(TEMP_DIR, 'nfe_temp.xml')
        csv_path = os.path.join(TEMP_DIR, 'sys_temp.csv')
        xml_file.save(xml_path)
        csv_file.save(csv_path)

        fornecedor = params.get('fornecedor', 'FORNECEDOR').strip()
        nota = params.get('nota', '000').strip()

        # Executa o motor de cálculo
        sucesso, df = gerar_tabela(xml_path, csv_path, fornecedor, nota, params)

        if sucesso:
            # 1. Gera o arquivo Excel físico com todas as cores e bordas
            nome_excel = f"Precificacao_{fornecedor}_NF_{nota}.xlsx"
            caminho_excel = os.path.join(TEMP_DIR, nome_excel)
            salvar_excel_estilizado(df, caminho_excel)

            # 2. Constrói o Preview HTML para o navegador com as cores solicitadas
            df_view = df.head(21)  # Pega o header de taxas + 20 itens

            cols_azul = ['PRODUTO +', 'PREÇO VAREJO', 'MARGEM VAREJO']
            cols_amarelo = ['PRODUTO +.1', 'DESCONTO ATC (15%)', 'MARGEM ATC']

            html = '<table class="table table-sm table-bordered"><thead><tr>'
            for col in df_view.columns:
                html += f'<th>{col}</th>'
            html += '</tr></thead><tbody>'

            for i, row in df_view.iterrows():
                # Detecta se a linha é ST (Substituição Tributária)
                st_val = 0
                try:
                    st_val = float(row['ST'])
                except:
                    pass

                # Se for ST, a linha inteira recebe a classe table-st (cor de pele)
                tr_class = 'class="table-st"' if st_val > 0 else ''
                html += f'<tr {tr_class}>'

                for col_name in df_view.columns:
                    val = row[col_name]
                    td_class = ''

                    # Formatação visual de porcentagem apenas para Margens
                    if col_name in ['MARGEM VAREJO', 'MARGEM ATC'] and i > 0:
                        try:
                            val = f"{float(val):.2f}%"
                        except:
                            pass

                    # Aplica cores Azul e Amarelo apenas se a linha NÃO for ST
                    if st_val <= 0:
                        if col_name in cols_azul:
                            td_class = 'class="bg-azul"'
                        elif col_name in cols_amarelo:
                            td_class = 'class="bg-amarelo"'

                    html += f'<td {td_class}>{val}</td>'
                html += '</tr>'

            html += '</tbody></table>'

            return jsonify({
                'sucesso': True,
                'tabela': html,
                'download_url': f'/download/{nome_excel}'
            })

        return jsonify({'sucesso': False, 'erro': df})

    except Exception as e:
        return jsonify({'sucesso': False, 'erro': str(e)})


@app.route('/download/<filename>')
def download(filename):
    """Rota para baixar o arquivo gerado."""
    caminho = os.path.join(TEMP_DIR, filename)
    return send_file(caminho, as_attachment=True)


def open_browser():
    """Abre o navegador automaticamente."""
    webbrowser.open_new("http://127.0.0.1:5000")


if __name__ == '__main__':
    Timer(1, open_browser).start()
    app.run(port=5000, debug=False)