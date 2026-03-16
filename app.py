import os
import sys
import tempfile
import webbrowser
from threading import Timer
from flask import Flask, render_template, request, jsonify, send_file

from core.processador import gerar_tabela, salvar_excel_estilizado, gerar_dashboard_html


def _arredondar_99(c_real):
    if c_real <= 0:
        return 0.0
    x = c_real * 2
    return int(x) - (1 if (x - int(x)) <= 0.5 else 0) + 0.99


def _calcular(row, P):
    nf_u  = row['nf_u']
    st_u  = row['st_u']
    ant_u = row['ant_u']
    ipi_u = row['ipi_u']
    cst   = str(row['cst'])
    qtd   = row['qtd']

    c_real = round(nf_u * P['mult'], 2)
    frete  = round(c_real * P['frete'], 2)
    desp   = round(c_real * P['desp'], 2)

    if st_u > 0.005 or cst in {'40', '60', '102', '500'}:
        cred = 0.0
    else:
        cred = round(nf_u * P['cred'], 2)

    c_ent = round(c_real + st_u + ant_u + ipi_u + frete + desp - cred, 2)

    p_atual = row['p_atual']
    p_var   = p_atual if p_atual > 0 else round(_arredondar_99(c_real), 2)

    fed    = round(p_var * P['fed'],    2)
    cartao = round(p_var * P['cartao'], 2)
    icms_s = 0.0 if st_u > 0.005 else round(p_var * P['icm'], 2)
    c_saida = round(c_ent + fed + cartao + icms_s, 2)

    meta   = 0.15
    p_min  = round(c_saida / (1 - meta), 2) if meta > 0 and c_saida > 0 else 0.0
    margem = round((p_var - c_saida) / p_var, 4) if p_var > 0 else 0.0
    lucro  = round((p_var - c_saida) * qtd, 2)

    # Atacado
    mult_atc = P.get('mult_atc', 1.3)
    desc_atc = P.get('desc_atc', 0.15)
    nf_atc      = round(nf_u * mult_atc, 2)
    p_atc       = round(p_var * (1 - desc_atc), 2)
    fed_atc     = round(nf_atc * P['fed'], 2)
    cart_atc    = round(p_atc  * P['cartao'], 2)
    icm_atc     = 0.0 if st_u > 0.005 else round(nf_atc * P['icm'], 2)
    c_saida_atc = round(c_ent + fed_atc + cart_atc + icm_atc, 2)
    margem_atc  = round((p_atc - c_saida_atc) / p_atc, 4) if p_atc > 0 else 0.0

    return dict(c_real=c_real, frete=frete, desp=desp, cred=cred,
                c_ent=c_ent, fed=fed, cartao=cartao, icms_s=icms_s,
                c_saida=c_saida, p_min=p_min, p_var=p_var, margem=margem,
                lucro=lucro,
                nf_atc=nf_atc, p_atc=p_atc, fed_atc=fed_atc,
                cart_atc=cart_atc, icm_atc=icm_atc,
                c_saida_atc=c_saida_atc, margem_atc=margem_atc)


if getattr(sys, 'frozen', False):
    template_folder = os.path.join(sys._MEIPASS, 'templates')
    static_folder   = os.path.join(sys._MEIPASS, 'static')
    app = Flask(__name__, template_folder=template_folder, static_folder=static_folder)
else:
    app = Flask(__name__)

TEMP_DIR = tempfile.gettempdir()


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/processar', methods=['POST'])
def processar():
    try:
        xml_file = request.files.get('xml')
        csv_file = request.files.get('csv')
        params   = request.form.to_dict()

        if not xml_file or not csv_file:
            return jsonify({'sucesso': False, 'erro': 'Por favor, anexe os arquivos XML e CSV.'})

        xml_path = os.path.join(TEMP_DIR, 'nfe_temp.xml')
        csv_path = os.path.join(TEMP_DIR, 'sys_temp.csv')
        xml_file.save(xml_path)
        csv_file.save(csv_path)

        fornecedor = params.get('fornecedor', 'FORNECEDOR').strip()
        nota       = params.get('nota', '000').strip()

        sucesso, resultado = gerar_tabela(xml_path, csv_path, fornecedor, nota, params)

        if not sucesso:
            return jsonify({'sucesso': False, 'erro': resultado})

        rows, P, num_nf = resultado

        # ── Gerar Excel ──────────────────────────────────────────────────────
        nome_excel   = f"Precificacao_{fornecedor}_NF_{nota}.xlsx"
        caminho_excel = os.path.join(TEMP_DIR, nome_excel)
        salvar_excel_estilizado(resultado, caminho_excel)

        # ── Calcular métricas para prévia e dashboard ────────────────────────
        metricas    = [_calcular(row, P) for row in rows]
        lucro_total = sum(m['lucro'] for m in metricas)

        # ── Montar tabela HTML de prévia (primeiros 20 produtos) ─────────────
        # ── Grupos de colunas com rótulo separador no cabeçalho ─────────────
        # Índice 0-based de cada coluna na lista de cells abaixo:
        #   0-4   : NF, DESC, REF, SKU, QTD
        #   5-8   : NF UNIT, ST, ANT, IPI
        #   9-13  : FRETE, DESPESA, CRED ICMS, C. REAL, C. ENTRADA
        #   14    : CST
        #  ── VAREJO ──
        #   15-18 : FEDERAL, CARTÃO, ICMS S., C. SAÍDA
        #   19-20 : META %, PREÇO MÍN
        #   21-23 : PREÇO ATUAL, PREÇO VAREJO, MARGEM
        #  ── ATACADO ──
        #   24-28 : NF ATC, FEDERAL ATC, CARTÃO ATC, ICMS ATC, C. SAÍDA ATC
        #   29-30 : PREÇO ATC, MARGEM ATC

        col_headers = [
            'NF', 'DESCRIÇÃO', 'REF', 'SKU', 'QTD',
            'NF UNIT', 'ST UNIT', 'ANT UNIT', 'IPI UNIT',
            'FRETE', 'DESPESA', 'CRED ICMS',
            'C. REAL', 'C. ENTRADA', 'CST',
            # varejo
            'FEDERAL', 'CARTÃO', 'ICMS S.', 'C. SAÍDA',
            'META %', 'PREÇO MÍN',
            'PREÇO ATUAL', 'PREÇO VAREJO', 'MARGEM',
            # atacado
            'NF ATC', 'FEDERAL ATC', 'CARTÃO ATC', 'ICMS ATC', 'C. SAÍDA ATC',
            'PREÇO ATC', 'MARGEM ATC',
        ]

        # Cabeçalho com separador visual entre varejo e atacado
        html  = '<table class="table table-sm table-bordered table-hover"><thead>'
        html += '<tr>'
        html += '<th colspan="15" style="background:#f8f9fa;text-align:center"></th>'
        html += '<th colspan="9" style="background:#dbeafe;text-align:center;font-size:0.7rem;letter-spacing:1px;color:#1d4ed8">VAREJO</th>'
        html += '<th colspan="7" style="background:#ede9fe;text-align:center;font-size:0.7rem;letter-spacing:1px;color:#5b21b6">ATACADO</th>'
        html += '</tr><tr>'
        html += ''.join(f'<th>{h}</th>' for h in col_headers)
        html += '</tr></thead><tbody>'

        # Índices (0-based) com cor especial em linhas normais:
        # azul  → PREÇO VAREJO (22), MARGEM (23)
        # amar  → PREÇO MÍN (20), PREÇO ATC (29), MARGEM ATC (30)
        _AMAR = {20, 29, 30}
        _AZUL = {22, 23}

        for row, m in zip(rows[:20], metricas[:20]):
            has_st  = row['st_u']  > 0.005
            has_ant = row['ant_u'] > 0.005

            if has_st:
                row_bg = '#FCE4D6'   # peach  — produto com ST
            elif has_ant:
                row_bg = '#D1FAE5'   # verde menta — produto com ANT (antecipação, sem ST)
            else:
                row_bg = None

            html += '<tr>'

            cells = [
                row['nf'],                                                                    # 0
                row['desc'][:55],                                                             # 1
                row['ref'],                                                                   # 2
                row['sku'],                                                                   # 3
                int(row['qtd']),                                                              # 4
                f"R$ {row['nf_u']:.2f}",                                                     # 5
                f"R$ {row['st_u']:.2f}"  if row['st_u']  > 0.005 else '-',                  # 6
                f"R$ {row['ant_u']:.2f}" if row['ant_u'] > 0.005 else '-',                  # 7
                f"R$ {row['ipi_u']:.2f}" if row['ipi_u'] > 0.005 else '-',                  # 8
                f"R$ {m['frete']:.2f}",                                                      # 9
                f"R$ {m['desp']:.2f}",                                                       # 10
                f"R$ {m['cred']:.2f}",                                                       # 11
                f"R$ {m['c_real']:.2f}",                                                     # 12
                f"R$ {m['c_ent']:.2f}",                                                      # 13
                row['cst'],                                                                   # 14
                # varejo
                f"R$ {m['fed']:.2f}",                                                        # 15
                f"R$ {m['cartao']:.2f}",                                                     # 16
                f"R$ {m['icms_s']:.2f}",                                                     # 17
                f"R$ {m['c_saida']:.2f}",                                                    # 18
                '15%',                                                                        # 19
                f"R$ {m['p_min']:.2f}",                                                      # 20
                f"R$ {row['p_atual']:.2f}" if row['p_atual'] > 0 else '<span style="color:#e74c3c">SEM PREÇO</span>', # 21
                f"R$ {m['p_var']:.2f}",                                                      # 22
                f"{m['margem']*100:.1f}%",                                                   # 23
                # atacado
                f"R$ {m['nf_atc']:.2f}",                                                     # 24
                f"R$ {m['fed_atc']:.2f}",                                                    # 25
                f"R$ {m['cart_atc']:.2f}",                                                   # 26
                f"R$ {m['icm_atc']:.2f}",                                                    # 27
                f"R$ {m['c_saida_atc']:.2f}",                                                # 28
                f"R$ {m['p_atc']:.2f}",                                                      # 29
                f"{m['margem_atc']*100:.1f}%",                                               # 30
            ]

            # Faixa de colunas atacado (24-30)
            _ATC_RANGE = set(range(24, 31))

            for idx, val in enumerate(cells):
                if has_st:
                    # ST: peach em tudo, sem exceção
                    style = ' style="background:#FCE4D6"'
                elif idx in _AZUL:
                    # PREÇO VAREJO / MARGEM: azul sempre (prevalece sobre ANT)
                    style = ' style="background:#BDD7EE;font-weight:600"'
                elif idx in _AMAR:
                    # PREÇO MÍN / PREÇO ATC / MARGEM ATC: amarelo sempre (prevalece sobre ANT)
                    style = ' style="background:#FFF2CC;font-weight:600"'
                elif has_ant:
                    # ANT: verde menta nas demais células
                    style = ' style="background:#D1FAE5"'
                elif idx in _ATC_RANGE:
                    # Normal: roxo suave na faixa atacado
                    style = ' style="background:#F3E8FF"'
                else:
                    style = ''
                html += f'<td{style}>{val}</td>'

            html += '</tr>'

        html += '</tbody></table>'
        if len(rows) > 20:
            html += f'<p class="text-muted small">Mostrando 20 de {len(rows)} produtos. Baixe o Excel para ver todos.</p>'

        # ── Dashboard HTML ───────────────────────────────────────────────────
        dashboard_html = gerar_dashboard_html(rows, lucro_total, metricas)

        return jsonify({
            'sucesso':      True,
            'tabela':       html,
            'dashboard':    dashboard_html,
            'download_url': f'/download/{nome_excel}',
            'total_itens':  len(rows),
        })

    except Exception as e:
        import traceback
        return jsonify({'sucesso': False, 'erro': traceback.format_exc()})


@app.route('/download/<filename>')
def download(filename):
    caminho = os.path.join(TEMP_DIR, filename)
    return send_file(caminho, as_attachment=True)


def open_browser():
    webbrowser.open_new("http://127.0.0.1:8080")


if __name__ == '__main__':
    Timer(1, open_browser).start()
    app.run(port=8080, debug=False)