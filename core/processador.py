import pandas as pd
import xml.etree.ElementTree as ET
import math
import requests
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font


def limpar_str(val):
    if pd.isna(val) or str(val).strip() == "": return ""
    s = str(val).replace('="', '').replace('"', '').strip()
    if s.endswith('.0'): s = s[:-2]
    return s


def limpar_preco(val):
    if pd.isna(val) or str(val).strip() == "": return 0.0
    try:
        return float(str(val).replace('R$', '').replace('.', '').replace(',', '.').strip())
    except:
        return 0.0


def arredondar_99(valor):
    if pd.isna(valor) or valor <= 0: return 0.0
    inteiro = math.floor(valor)
    decimal = valor - inteiro
    return float(inteiro - 1 if decimal <= 0.50 else inteiro) + 0.99


def gerar_tabela(xml_path, csv_path, fornecedor, nota_ref, params):
    try:
        P_MULT = float(params.get('mult_var', 2.0))
        P_ATC_MULT = float(params.get('mult_atc', 1.3))
        P_DESP = float(params.get('desp', 10)) / 100
        P_FRETE = float(params.get('frete', 10)) / 100
        P_CRED_ICMS = float(params.get('cred_icms', 4)) / 100
        P_FED = float(params.get('fed', 9.13)) / 100
        P_ICM = float(params.get('icm', 20)) / 100
        P_CARTAO = float(params.get('cartao', 4)) / 100
        P_DESC_ATC = 1.0 - (float(params.get('desc_atc', 15)) / 100)

        tree = ET.parse(xml_path)
        ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
        infNFe = tree.getroot().find('.//nfe:infNFe', ns)
        tag_nnf = infNFe.find('.//nfe:nNF', ns)
        num_nf_xml = tag_nnf.text if tag_nnf is not None else ''
        chave_nfe = infNFe.attrib.get('Id', '')[3:] if infNFe is not None else ""

        # --- REQUISIÇÃO API SEFAZ ALAGOAS (ST/ANTECIPADO) ---
        dados_api_st = []
        try:
            api_url = f"https://contribuinte.sefaz.al.gov.br/cobrancadfe/sfz-cobranca-dfe-api/api/detalhe-calculo-nfes?chaveNota.equals={chave_nfe}"
            headers = {
                'Accept': 'application/json, text/plain, */*',
                'Referer': 'https://contribuinte.sefaz.al.gov.br/cobrancadfe/',
                'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15',
                'Accept-Language': 'pt-BR,pt;q=0.9'
            }
            cookies = {'cookiesAccepted': 'true'}
            response = requests.get(api_url, headers=headers, cookies=cookies, timeout=10)
            if response.status_code == 200:
                dados_api_st = response.json()
        except:
            pass

        itens_xml = []
        for det in tree.getroot().findall('.//nfe:det', ns):
            p = det.find('nfe:prod', ns)
            i = det.find('nfe:imposto', ns)

            ean_xml = limpar_str(p.find('nfe:cEAN', ns).text if p.find('nfe:cEAN', ns) is not None else '')
            ref_xml = limpar_str(p.find('nfe:cProd', ns).text if p.find('nfe:cProd', ns) is not None else '')
            desc_xml = p.find('nfe:xProd', ns).text if p.find('nfe:xProd', ns) is not None else ''

            # ST do XML Original
            v_st_xml = 0.0
            if i is not None:
                st_node = i.find('.//nfe:vICMSST', ns)
                if st_node is not None: v_st_xml = float(st_node.text)

            # Cruzamento com API SEFAZ pela Descrição do Produto
            v_st_api = 0.0
            for item_api in dados_api_st:
                desc_api = str(item_api.get('descricaoProduto', '')).strip().upper()
                if desc_api in desc_xml.strip().upper() or desc_xml.strip().upper() in desc_api:
                    v_st_api = float(item_api.get('valorIcmsCalculado', 0) or 0) + float(
                        item_api.get('valorFecoepCalculado', 0) or 0)
                    break

            v_st_total = v_st_xml + v_st_api

            # Busca CST
            cst_csosn = ""
            if i is not None:
                icms_node = i.find('.//nfe:ICMS', ns)
                if icms_node is not None:
                    for child in icms_node:
                        cst_tag = child.find('nfe:CST', ns) or child.find('nfe:CSOSN', ns)
                        if cst_tag is not None: cst_csosn = cst_tag.text; break

            itens_xml.append({
                'ean_xml': ean_xml,
                'ref_xml': ref_xml,
                'desc_xml': desc_xml,
                'qCom': float(p.find('nfe:qCom', ns).text),
                'vProd': float(p.find('nfe:vProd', ns).text),
                'vIPI': float(i.find('.//nfe:vIPI', ns).text) if i is not None and i.find('.//nfe:vIPI',
                                                                                          ns) is not None else 0.0,
                'vST': v_st_total,
                'nf_xml_base': num_nf_xml,
                'cst': cst_csosn
            })

        df_xml = pd.DataFrame(itens_xml)

        # 2. CSV DO SISTEMA - LÓGICA ORIGINAL RESTAURADA
        try:
            df_sys = pd.read_csv(csv_path, sep=';', encoding='utf-8', on_bad_lines='skip', dtype=str)
        except:
            df_sys = pd.read_csv(csv_path, sep=';', encoding='latin1', on_bad_lines='skip', dtype=str)

        df_sys.columns = df_sys.columns.str.replace('"', '').str.strip().str.upper()
        df_sys['ean_sys'] = df_sys['BARRA'].apply(limpar_str) if 'BARRA' in df_sys.columns else ''
        df_sys['ref_sys'] = df_sys['REFERÊNCIA'].apply(limpar_str) if 'REFERÊNCIA' in df_sys.columns else ''
        df_sys['nota_sys'] = df_sys['NOTA'].apply(limpar_str) if 'NOTA' in df_sys.columns else ''
        df_sys['preco_sys'] = df_sys['PREÇO'].apply(limpar_preco) if 'PREÇO' in df_sys.columns else 0.0

        # MERGE PADRÃO DO BASELINE
        df_base = pd.merge(df_xml, df_sys[['ean_sys', 'ref_sys', 'preco_sys', 'nota_sys']],
                           left_on='ean_xml', right_on='ean_sys', how='left')

        # FALLBACK PARA REFERÊNCIA (QUANDO O CÓDIGO DE BARRAS NÃO BATE)
        sem_match = (df_base['preco_sys'].isna()) | (df_base['preco_sys'] == 0)
        if sem_match.any():
            fallback = pd.merge(df_base.loc[sem_match, ['ref_xml']],
                                df_sys[['ean_sys', 'ref_sys', 'preco_sys', 'nota_sys']],
                                left_on='ref_xml', right_on='ref_sys', how='left')
            df_base.loc[sem_match, 'preco_sys'] = fallback['preco_sys'].values
            df_base.loc[sem_match, 'ean_sys'] = fallback['ean_sys'].values
            df_base.loc[sem_match, 'nota_sys'] = fallback['nota_sys'].values

        df_base['preco_sys'] = df_base['preco_sys'].fillna(0.0)
        df_base['sku_final'] = df_base.apply(
            lambda r: r['ean_sys'] if pd.notna(r['ean_sys']) and r['ean_sys'] != "" else r['ean_xml'], axis=1)

        # 3. CÁLCULO DAS LINHAS
        def calcular_linha(row):
            qtd = row['qCom']
            if qtd <= 0: return pd.Series(dtype=float)

            nf_un = round(row['vProd'] / qtd, 2)
            st_un = round(row.get('vST', 0) / qtd, 2)

            # REGRA MANTIDA: Custo Real Base não soma o ST.
            c_real = round(nf_un * P_MULT, 2)
            nf_atc = round(nf_un * P_ATC_MULT + st_un, 2)
            despesa = round(c_real * P_DESP, 2)
            frete = round(c_real * P_FRETE, 2)

            sem_credito = ['40', '41', '50', '60', '102', '103', '300', '400', '500']
            cred_icms = 0.0 if (str(row.get('cst', '')) in sem_credito or st_un > 0) else round(nf_un * P_CRED_ICMS, 2)

            ipi_un = round(row['vIPI'] / qtd, 2)
            p_varejo = arredondar_99(row['preco_sys'] if row['preco_sys'] > 0 else (c_real * 2.5))
            p_atc = round(p_varejo * P_DESC_ATC, 2)

            fed_var = round(p_varejo * P_FED, 2)
            icm_var = round(p_varejo * P_ICM, 2)
            prem_var = round(p_varejo * P_CARTAO, 2)

            # PRODUTO + SOMA O ST_UN
            p_mais_var = round(c_real + st_un + despesa + frete + ipi_un + fed_var + icm_var + prem_var - cred_icms, 2)

            return pd.Series([
                row['nf_xml_base'], row['desc_xml'], row['ref_xml'], row['sku_final'], qtd,
                nf_un, st_un, c_real, nf_atc, despesa, frete, -cred_icms, 0.0, ipi_un,
                fed_var, round(nf_atc * P_FED, 2), icm_var, round(nf_atc * P_ICM, 2), prem_var,
                round(p_atc * P_CARTAO, 2),
                p_mais_var, 0.0, p_atc, p_varejo, row['preco_sys'],
                round(((p_varejo - p_mais_var) / p_varejo) * 100, 2) if p_varejo > 0 else 0.0, 0.0
            ])

        header_vals = [
            ['', '', '', '', '', '', 'ST_XML', P_MULT, P_ATC_MULT, P_DESP, P_FRETE, -P_CRED_ICMS, 'ALIQ', 0.0, P_FED,
             P_FED, P_ICM, P_ICM, P_CARTAO, P_CARTAO, 'IMP VAR', 'IMP ATC', P_DESC_ATC, 'PREÇO VAR', 'PREÇO ATUAL',
             'M. VAR', 'M. ATC']]

        df_final = pd.concat([pd.DataFrame(header_vals), df_base.apply(calcular_linha, axis=1)], ignore_index=True)
        df_final.columns = ['NF', 'DESCRIÇÃO', 'REF', 'SKU', 'QTD', 'NF.1', 'ST', 'CUSTO REAL', 'NF ATC', 'DESPESA',
                            'FRETE', 'CRED. ICMS', 'IPI (ALIQUOTA)', 'IPI.1', 'FEDERAL VAREJO', 'FEDERAL ATC',
                            'ICM VENDA VAREJO', 'ICM VENDA ATC', 'PREM + CARTÃO VAREJO', 'PREM + CARTÃO ATC',
                            'PRODUTO +', 'PRODUTO +.1', 'DESCONTO ATC (15%)', 'PREÇO VAREJO', 'PREÇO ATUAL',
                            'MARGEM VAREJO', 'MARGEM ATC']
        return True, df_final
    except Exception as e:
        return False, str(e)


def salvar_excel_estilizado(df, path):
    fill_azul = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
    fill_amarelo = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_pele = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    borda = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Precificação')
        ws = writer.sheets['Precificação']
        cols_percent = ['MARGEM VAREJO', 'MARGEM ATC']
        col_map = {col: i + 1 for i, col in enumerate(df.columns)}

        for row_idx in range(1, len(df) + 2):
            is_st = False
            if row_idx > 2:
                try:
                    if float(df.iloc[row_idx - 2]['ST']) > 0: is_st = True
                except:
                    pass

            for col_name, col_idx in col_map.items():
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.border = borda
                cell.alignment = Alignment(horizontal='center')

                if col_name in cols_percent and row_idx > 1:
                    if isinstance(cell.value, (int, float)):
                        if abs(cell.value) > 1: cell.value = cell.value / 100
                    cell.number_format = '0.00%'
                elif row_idx > 1 and isinstance(cell.value, (int, float)):
                    cell.number_format = '0.00'

                if row_idx == 1:
                    cell.font = Font(bold=True)
                elif row_idx == 2:
                    if col_idx in [8, 9, 10, 11, 12, 15, 16, 17, 18, 19, 20, 23]:
                        cell.number_format = '0.00%'
                    cell.font = Font(bold=True, color="333333")
                elif is_st:
                    cell.fill = fill_pele
                else:
                    if col_name in ['PRODUTO +', 'PREÇO VAREJO', 'MARGEM VAREJO']:
                        cell.fill = fill_azul
                    elif col_name in ['PRODUTO +.1', 'DESCONTO ATC (15%)', 'MARGEM ATC']:
                        cell.fill = fill_amarelo

            if row_idx >= 3:
                r = row_idx
                ws.cell(row=r, column=8).value = f"=ROUND(F{r}*H$2, 2)"  # CUSTO REAL sem ST
                ws.cell(row=r, column=9).value = f"=ROUND((F{r}*I$2)+G{r}, 2)"
                ws.cell(row=r, column=10).value = f"=ROUND(H{r}*J$2, 2)"
                ws.cell(row=r, column=11).value = f"=ROUND(H{r}*K$2, 2)"

                try:
                    val_cred = float(df.iloc[row_idx - 2]['CRED. ICMS'])
                    if val_cred == 0.0:
                        ws.cell(row=r, column=12).value = 0
                    else:
                        ws.cell(row=r, column=12).value = f"=ROUND(F{r}*L$2, 2)"
                except:
                    ws.cell(row=r, column=12).value = f"=ROUND(F{r}*L$2, 2)"

                ws.cell(row=r, column=15).value = f"=ROUND(X{r}*O$2, 2)"  # FEDERAL sobre VENDA (X)
                ws.cell(row=r, column=16).value = f"=ROUND(I{r}*P$2, 2)"
                ws.cell(row=r, column=17).value = f"=ROUND(X{r}*Q$2, 2)"
                ws.cell(row=r, column=18).value = f"=ROUND(I{r}*R$2, 2)"
                ws.cell(row=r, column=19).value = f"=ROUND(X{r}*S$2, 2)"
                ws.cell(row=r, column=20).value = f"=ROUND(W{r}*T$2, 2)"

                # PRODUTO + soma G(ST) e H(CustoReal) separados
                ws.cell(row=r, column=21).value = f"=ROUND(H{r}+G{r}+J{r}+K{r}+N{r}+O{r}+Q{r}+S{r}+L{r}, 2)"
                ws.cell(row=r, column=22).value = f"=ROUND(H{r}+G{r}+J{r}+K{r}+N{r}+P{r}+R{r}+T{r}+L{r}, 2)"
                ws.cell(row=r, column=23).value = f"=ROUND(X{r}*W$2, 2)"
                ws.cell(row=r, column=26).value = f"=IF(X{r}>0, (X{r}-U{r})/X{r}, 0)"
                ws.cell(row=r, column=27).value = f"=IF(W{r}>0, (W{r}-V{r})/W{r}, 0)"

        for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = 20