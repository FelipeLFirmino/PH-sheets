import pandas as pd
import xml.etree.ElementTree as ET
import math
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font

def limpar_str(val):
    if pd.isna(val) or str(val).strip() == "": return ""
    return str(val).replace('="', '').replace('"', '').strip()

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
        num_nf_xml = infNFe.find('.//nfe:nNF', ns).text if infNFe is not None else "000"

        itens_xml = []
        for det in tree.getroot().findall('.//nfe:det', ns):
            p = det.find('nfe:prod', ns)
            i = det.find('nfe:imposto', ns)
            v_st = float(i.find('.//nfe:vICMSST', ns).text) if i.find('.//nfe:vICMSST', ns) is not None else 0.0
            itens_xml.append({
                'ean_xml': limpar_str(p.find('nfe:cEAN', ns).text),
                'ref_xml': limpar_str(p.find('nfe:cProd', ns).text),
                'desc_xml': p.find('nfe:xProd', ns).text,
                'qCom': float(p.find('nfe:qCom', ns).text),
                'vProd': float(p.find('nfe:vProd', ns).text),
                'vIPI': float(i.find('.//nfe:vIPI', ns).text) if i.find('.//nfe:vIPI', ns) is not None else 0.0,
                'vST': v_st
            })
        df_xml = pd.DataFrame(itens_xml)

        df_sys = pd.read_csv(csv_path, sep=';', encoding='utf-8', dtype=str)
        df_sys.columns = df_sys.columns.str.replace('"', '').str.strip().str.upper()

        df_sys.loc[:, 'BARRA'] = df_sys['BARRA'].apply(limpar_str)
        df_sys.loc[:, 'REFERÊNCIA'] = df_sys['REFERÊNCIA'].apply(limpar_str)
        df_sys.loc[:, 'preco_sys'] = df_sys['PREÇO'].apply(limpar_preco)

        df_base = pd.merge(df_xml, df_sys[['BARRA', 'REFERÊNCIA', 'preco_sys']],
                           left_on='ean_xml', right_on='BARRA', how='left')

        mask_vazio = (df_base['preco_sys'].isna()) | (df_base['preco_sys'] == 0)
        if mask_vazio.any():
            map_ref_preco = df_sys.set_index('REFERÊNCIA')['preco_sys'].to_dict()
            map_ref_barra = df_sys.set_index('REFERÊNCIA')['BARRA'].to_dict()

            df_base.loc[mask_vazio, 'preco_sys'] = df_base.loc[mask_vazio, 'ref_xml'].map(map_ref_preco)
            df_base.loc[mask_vazio, 'BARRA'] = df_base.loc[mask_vazio, 'ref_xml'].map(map_ref_barra)

        df_base.loc[:, 'preco_sys'] = df_base['preco_sys'].fillna(0.0)

        def calcular_linha(row):
            qtd = row['qCom']
            if qtd <= 0: return pd.Series(dtype=float)

            sku_final = row['BARRA'] if pd.notna(row['BARRA']) and row['BARRA'] != "" else row['ean_xml']

            nf_un = round(row['vProd'] / qtd, 2)
            st_un = round(row['vST'] / qtd, 2)
            c_real = round((nf_un * P_MULT) + st_un, 2)
            nf_atc = round((nf_un * P_ATC_MULT) + st_un, 2)
            despesa = round(c_real * P_DESP, 2)
            frete = round(c_real * P_FRETE, 2)
            cred_icms = round(nf_un * P_CRED_ICMS, 2)
            ipi_un = round(row['vIPI'] / qtd, 2)
            fed_var = round(nf_atc * P_FED, 2)
            fed_atc = round(nf_atc * P_FED, 2)
            p_varejo = arredondar_99(row['preco_sys'] if row['preco_sys'] > 0 else (c_real * 2.5))
            p_atc = round(p_varejo * P_DESC_ATC, 2)
            icm_var = round(p_varejo * P_ICM, 2)
            icm_atc = round(nf_atc * P_ICM, 2)
            prem_var = round(p_varejo * P_CARTAO, 2)
            prem_atc = round(p_atc * P_CARTAO, 2)
            p_mais_var = round(c_real + despesa + frete + ipi_un + fed_var + icm_var + prem_var - cred_icms, 2)
            p_mais_atc = round(c_real + despesa + frete + ipi_un + fed_atc + icm_atc + prem_atc - cred_icms, 2)

            return pd.Series([
                num_nf_xml, row['desc_xml'], row['ref_xml'], sku_final, qtd,
                nf_un, st_un, c_real, nf_atc, despesa, frete, -cred_icms, ipi_un, ipi_un,
                fed_var, fed_atc, icm_var, icm_atc, prem_var, prem_atc,
                p_mais_var, p_mais_atc, p_atc, p_varejo, row['preco_sys'],
                round(((p_varejo - p_mais_var) / p_varejo) * 100, 2) if p_varejo > 0 else 0.0,
                round(((p_atc - p_mais_atc) / p_atc) * 100, 2) if p_atc > 0 else 0.0
            ])

        df_calc = df_base.apply(calcular_linha, axis=1)
        header_vals = [
            ['', '', '', '', '', '', 'ST_XML', P_MULT, P_ATC_MULT, P_DESP, P_FRETE, -P_CRED_ICMS, 'ALIQUOTA', 0.0975,
             P_FED, P_FED, P_ICM, P_ICM, P_CARTAO, P_CARTAO, 'IMPOSTOS VAR', 'IMPOSTOS ATC', P_DESC_ATC, 'PREÇO VAR',
             'PREÇO ATUAL', 'M. VAR', 'M. ATC']]
        df_final = pd.concat([pd.DataFrame(header_vals), df_calc], ignore_index=True)
        df_final.columns = [
            'NF', 'DESCRIÇÃO', 'REF', 'SKU', 'QTD', 'NF.1', 'ST', 'CUSTO REAL', 'NF ATC',
            'DESPESA', 'FRETE', 'CRED. ICMS', 'IPI (ALIQUOTA)', 'IPI.1', 'FEDERAL VAREJO', 'FEDERAL ATC',
            'ICM VENDA VAREJO', 'ICM VENDA ATC', 'PREM + CARTÃO VAREJO', 'PREM + CARTÃO ATC',
            'PRODUTO +', 'PRODUTO +.1', 'DESCONTO ATC (15%)', 'PREÇO VAREJO', 'PREÇO ATUAL', 'MARGEM VAREJO',
            'MARGEM ATC'
        ]
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
                    st_val = float(df.iloc[row_idx - 2]['ST'])
                    if st_val > 0: is_st = True
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
                    if col_idx in [10, 11, 12, 14, 15, 16, 17, 18, 19, 20]:
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
                ws.cell(row=r, column=8).value = f"=ROUND((F{r}*H$2)+G{r}, 2)"
                ws.cell(row=r, column=9).value = f"=ROUND((F{r}*I$2)+G{r}, 2)"
                ws.cell(row=r, column=10).value = f"=ROUND(H{r}*J$2, 2)"
                ws.cell(row=r, column=11).value = f"=ROUND(H{r}*K$2, 2)"
                ws.cell(row=r, column=12).value = f"=ROUND(F{r}*L$2, 2)"
                ws.cell(row=r, column=15).value = f"=ROUND(I{r}*O$2, 2)"
                ws.cell(row=r, column=16).value = f"=ROUND(I{r}*P$2, 2)"
                ws.cell(row=r, column=17).value = f"=ROUND(X{r}*Q$2, 2)"
                ws.cell(row=r, column=18).value = f"=ROUND(I{r}*R$2, 2)"
                ws.cell(row=r, column=19).value = f"=ROUND(X{r}*S$2, 2)"
                ws.cell(row=r, column=20).value = f"=ROUND(W{r}*T$2, 2)"
                ws.cell(row=r, column=21).value = f"=ROUND(H{r}+J{r}+K{r}+M{r}+O{r}+Q{r}+S{r}+L{r}, 2)"
                ws.cell(row=r, column=22).value = f"=ROUND(H{r}+J{r}+K{r}+N{r}+P{r}+R{r}+T{r}+L{r}, 2)"
                ws.cell(row=r, column=23).value = f"=ROUND(X{r}*W$2, 2)"
                ws.cell(row=r, column=26).value = f"=IF(X{r}>0, (X{r}-U{r})/X{r}, 0)"
                ws.cell(row=r, column=27).value = f"=IF(W{r}>0, (W{r}-V{r})/W{r}, 0)"

        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 20