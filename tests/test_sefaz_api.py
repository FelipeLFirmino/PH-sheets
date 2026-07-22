"""
Testes de parear_impostos_api() — lógica de combinação XML × SEFAZ AL.

Bug documentado (NF 13215 / 2026-03-09):
  Para produtos com ST (tipoImposto='ST'), a API retorna valorIcmsCalculado
  com o MESMO valor que já está em vICMSST no XML. O código antigo somava
  ambos → double-counting de até 73% do valor do produto (ex: ICMS-ST R$1,60
  num produto de R$2,25).

  O valorFecoepCalculado (FECOEP — Fundo de Combate à Pobreza) NÃO está no
  XML e DEVE ser somado ao vST mesmo para produtos ST.

Bug documentado (NF 59263 / 2026-04-07):
  Algumas NFs não destacam ST no XML (vICMSST=0), mas a API SEFAZ AL retorna
  os valores completos (ICMS + FECOEP). O código somava apenas FECOEP,
  ignorando o ICMS da API → st_u ficava apenas com o FECOEP (ex: R$0,47
  ao invés de R$8,67 por unidade).

Bug documentado (NF 16750 / 2026-07-20):
  Quando a NF tem múltiplos itens <det> com a MESMA descricaoProduto
  (produtos diferentes — REF/EAN distintos — ex: variações de cor/tamanho
  do mesmo item), o casamento por substring somava TODOS os registros da
  API daquela descrição em CADA linha do XML que batia, em vez de dividir
  um registro por linha. Ex: "FLOR ARTIFICIAL CX120" com 5 linhas no XML e
  5 registros ANT de R$146,80 cada (total real R$734,00) — o código antigo
  aplicava R$734,00 (soma dos 5) em CADA uma das 5 linhas, inflando o grupo
  para R$3.670,00 (5x o valor correto).

  Confirmado empiricamente que a API SEFAZ AL preserva a ordem exata dos
  itens <det> do XML e que descricaoProduto é o texto completo e literal
  do XML (não truncado) — inclusive replicando erros de digitação do
  fornecedor. Isso torna seguro casar por IGUALDADE EXATA de descrição e,
  dentro de cada grupo de descrição idêntica, por POSIÇÃO (o item N-ésimo
  do XML casa com o registro N-ésimo da API, na ordem).

  Quando a contagem de itens do XML não bate com a contagem de registros
  da API para a mesma descrição exata, cai em rateio proporcional a vProd
  e marca os itens como 'rateado' para conferência manual.

Regras de combinação (por item já pareado):
  - tipoImposto='ST' e vICMSST_xml > 0 → só adiciona valorFecoepCalculado (ICMS já no XML)
  - tipoImposto='ST' e vICMSST_xml = 0 → adiciona valorIcmsCalculado + valorFecoepCalculado (NF sem destaque)
  - tipoImposto='ANT'                  → adiciona valorIcmsCalculado + valorFecoepCalculado (XML tem zero)
"""
import pytest
from core.processador import parear_impostos_api


def _merge_single(v_st_xml, desc_xml, dados_api, vProd=1.0):
    """
    Helper de teste: pareia um único item do XML (sem grupo de duplicidade)
    contra a API, replicando a assinatura conveniente usada antes de
    parear_impostos_api() passar a operar sobre a nota inteira de uma vez.
    """
    itens = [{'desc_xml': desc_xml, 'v_st_xml': v_st_xml, 'vProd': vProd}]
    resultado = parear_impostos_api(itens, dados_api)
    vst, vant, _rateado = resultado[0]
    return vst, vant


# ─── Fixtures de payloads API reais (NF 13215) ──────────────────────────────

KIT_PRENDEDOR_CX180_API = {
    'descricaoProduto':     'KIT PRENDEDOR DE CABELO E ESPONJA DE MAQUIAGEM CX180',
    'tipoImposto':          'ST',
    'valorIcmsCalculado':   287.93,
    'valorFecoepCalculado': 16.86,
}
SACOLA_NAMORADO_CX240_API = {
    'descricaoProduto':     'SACOLA PARA PRESENTES NAMORADO CX240',
    'tipoImposto':          'ANT',
    'valorIcmsCalculado':   222.24,
    'valorFecoepCalculado': 14.82,
}
PRENDEDOR_CX144_API = {
    'descricaoProduto':     'PRENDEDOR DE CABELO CX144',
    'tipoImposto':          'ST',
    'valorIcmsCalculado':   89.58,
    'valorFecoepCalculado': 5.25,
}


# ─── Produto ST — não deve double-count ICMS ─────────────────────────────────

class TestProdutoST:
    """
    Para produtos ST, o XML já contém o ICMS em vICMSST.
    A API retorna o mesmo ICMS + o FECOEP.
    Só o FECOEP deve ser adicionado ao custo.
    """

    def test_st_nao_double_conta_icms(self):
        """
        REGRESSÃO: o bug somava vICMSST_xml + valorIcmsCalculado_api = dobro.
        Resultado correto: vST = vICMSST_xml + valorFecoepCalculado_api apenas.
        """
        v_st_xml = 287.93
        desc_xml = 'KIT PRENDEDOR DE CABELO E ESPONJA DE MAQUIAGEM CX180'
        dados_api = [KIT_PRENDEDOR_CX180_API]

        vST, vANT = _merge_single(v_st_xml, desc_xml, dados_api)

        # Correto: XML ICMS + FECOEP
        assert vST  == pytest.approx(287.93 + 16.86)  # 304.79
        assert vANT == pytest.approx(0.0)

    def test_st_nao_adiciona_icms_api(self):
        """Valor que seria dobrado (vICMSST_xml + valorIcmsCalculado_api) não deve aparecer."""
        v_st_xml = 89.58
        desc_xml = 'PRENDEDOR DE CABELO CX144'
        dados_api = [PRENDEDOR_CX144_API]

        vST, _ = _merge_single(v_st_xml, desc_xml, dados_api)

        dobro_errado = 89.58 + 89.58 + 5.25  # comportamento antigo
        assert vST != pytest.approx(dobro_errado), "ICMS está sendo dobrado"
        assert vST == pytest.approx(89.58 + 5.25)  # 94.83

    def test_st_sem_match_api_usa_so_xml(self):
        """Se a API não retornar esse produto, vST = só o XML (sem FECOEP)."""
        vST, vANT = _merge_single(287.93, 'KIT PRENDEDOR CX180', dados_api=[])
        assert vST  == pytest.approx(287.93)
        assert vANT == pytest.approx(0.0)

    def test_st_fecoep_zero_nao_altera_vst(self):
        """FECOEP = 0 → vST permanece igual ao XML."""
        api = {**KIT_PRENDEDOR_CX180_API, 'valorFecoepCalculado': 0}
        vST, _ = _merge_single(287.93, 'KIT PRENDEDOR DE CABELO E ESPONJA DE MAQUIAGEM CX180', [api])
        assert vST == pytest.approx(287.93)


# ─── Produto ANT — deve somar ICMS + FECOEP inteiros ─────────────────────────

class TestProdutoANT:
    """
    Para produtos ANT, o XML tem vICMSST=0.
    A API é a única fonte: soma ICMS + FECOEP completos → vai para vANT.
    """

    def test_ant_soma_icms_e_fecoep(self):
        """vANT = valorIcmsCalculado + valorFecoepCalculado."""
        v_st_xml = 0.0
        desc_xml = 'SACOLA PARA PRESENTES NAMORADO CX240'
        dados_api = [SACOLA_NAMORADO_CX240_API]

        vST, vANT = _merge_single(v_st_xml, desc_xml, dados_api)

        assert vST  == pytest.approx(0.0)
        assert vANT == pytest.approx(222.24 + 14.82)  # 237.06

    def test_ant_nao_vai_para_vst(self):
        """Valor ANT não deve contaminar vST."""
        vST, _ = _merge_single(0.0, 'SACOLA PARA PRESENTES NAMORADO CX240',
                                [SACOLA_NAMORADO_CX240_API])
        assert vST == pytest.approx(0.0)

    def test_ant_sem_match_api_retorna_zeros(self):
        """Sem match na API, produto ANT fica com custo zero (SEFAZ fora do ar)."""
        vST, vANT = _merge_single(0.0, 'PRODUTO SEM MATCH', [SACOLA_NAMORADO_CX240_API])
        assert vST  == pytest.approx(0.0)
        assert vANT == pytest.approx(0.0)


# ─── Matching por igualdade exata ─────────────────────────────────────────────

class TestMatching:
    """
    O matching é por IGUALDADE EXATA de descrição (não mais substring) —
    confirmado que descricaoProduto da API é o texto completo e literal do
    XML. Testa os casos esperados e os limites desse comportamento.
    """

    def test_match_exato(self):
        desc = 'KIT PRENDEDOR DE CABELO E ESPONJA DE MAQUIAGEM CX180'
        api  = [{**KIT_PRENDEDOR_CX180_API, 'descricaoProduto': desc}]
        vST, _ = _merge_single(100.0, desc, api)
        assert vST == pytest.approx(100.0 + 16.86)

    def test_descricao_parcial_nao_bate_mais(self):
        """
        Antes (substring): api_desc sendo prefixo de desc_xml gerava match.
        Agora (igualdade exata): só bate se for o texto inteiro idêntico.
        """
        api = [{**KIT_PRENDEDOR_CX180_API,
                'descricaoProduto': 'KIT PRENDEDOR DE CABELO'}]  # só um prefixo
        vST, _ = _merge_single(
            100.0, 'KIT PRENDEDOR DE CABELO E ESPONJA DE MAQUIAGEM CX180', api
        )
        assert vST == pytest.approx(100.0), "prefixo não deveria mais casar"

    def test_sem_match_nenhum_valor_adicionado(self):
        """Produto com descrição totalmente diferente não recebe nenhum valor da API."""
        vST, vANT = _merge_single(0.0, 'PRODUTO XPTO 99', [KIT_PRENDEDOR_CX180_API])
        assert vST  == pytest.approx(0.0)
        assert vANT == pytest.approx(0.0)

    def test_api_dados_vazios(self):
        """API fora do ar → dados_api=[] → vST = só XML, vANT = 0."""
        vST, vANT = _merge_single(50.0, 'QUALQUER PRODUTO', [])
        assert vST  == pytest.approx(50.0)
        assert vANT == pytest.approx(0.0)

    def test_multiplos_itens_mesma_nota_nao_contaminam(self):
        """
        API retorna vários itens da NF. Produto X não deve acumular valores
        de produto Y mesmo que estejam na mesma lista.
        """
        dados_api = [
            KIT_PRENDEDOR_CX180_API,    # ST — não deve bater com SACOLA
            SACOLA_NAMORADO_CX240_API,  # ANT — não deve bater com KIT
        ]
        # Processando apenas a SACOLA
        vST, vANT = _merge_single(0.0, 'SACOLA PARA PRESENTES NAMORADO CX240', dados_api)
        assert vST  == pytest.approx(0.0)
        assert vANT == pytest.approx(222.24 + 14.82)  # só a SACOLA


# ─── Duplicidade de descrição — bug NF 16750 ─────────────────────────────────

class TestDuplicidadeDescricao:
    """
    Grupo de N itens do XML com a MESMA descrição casado com M registros da
    API da mesma descrição. Pareamento é posicional dentro do grupo — cada
    item recebe UM registro da API, nunca a soma de todos.
    """

    def _api_ant(self, desc, icms, fecoep):
        return {
            'descricaoProduto':     desc,
            'tipoImposto':          'ANT',
            'valorIcmsCalculado':   icms,
            'valorFecoepCalculado': fecoep,
        }

    def test_grupo_n_igual_m_pareia_por_posicao_nao_soma(self):
        """
        REGRESSÃO NF 16750: 5 itens "FLOR ARTIFICIAL CX120" no XML, 5
        registros ANT de 138.41/8.39 (=146.80) cada na API. Cada item deve
        receber 146.80 — não 5x146.80=734.00 (o bug antigo aplicava a soma
        dos 5 registros em cada uma das 5 linhas).
        """
        desc = 'FLOR ARTIFICIAL CX120'
        itens = [{'desc_xml': desc, 'v_st_xml': 0.0, 'vProd': 600.0} for _ in range(5)]
        dados_api = [self._api_ant(desc, 138.41, 8.39) for _ in range(5)]

        resultado = parear_impostos_api(itens, dados_api)

        assert len(resultado) == 5
        for idx in range(5):
            vst, vant, rateado = resultado[idx]
            assert vant == pytest.approx(146.80)
            assert vst  == pytest.approx(0.0)
            assert rateado is False

    def test_grupo_com_valores_distintos_mantem_ordem(self):
        """
        REGRESSÃO NF 16750: "ABAJUR DECORATIVO CX12" com 2 itens de vProd
        diferentes (270 e 300) casa com 2 registros ST de valores diferentes
        (141.46/7.43 e 157.17/8.25), na mesma ordem de aparição.
        """
        desc = 'ABAJUR DECORATIVO CX12'
        itens = [
            {'desc_xml': desc, 'v_st_xml': 0.0, 'vProd': 270.0},
            {'desc_xml': desc, 'v_st_xml': 0.0, 'vProd': 300.0},
        ]
        dados_api = [
            {'descricaoProduto': desc, 'tipoImposto': 'ST',
             'valorIcmsCalculado': 141.46, 'valorFecoepCalculado': 7.43, 'id': 23},
            {'descricaoProduto': desc, 'tipoImposto': 'ST',
             'valorIcmsCalculado': 157.17, 'valorFecoepCalculado': 8.25, 'id': 24},
        ]

        resultado = parear_impostos_api(itens, dados_api)

        vst0, _, _ = resultado[0]
        vst1, _, _ = resultado[1]
        assert vst0 == pytest.approx(141.46 + 7.43)
        assert vst1 == pytest.approx(157.17 + 8.25)

    def test_grupos_diferentes_intercalados_nao_se_misturam(self):
        """
        Réplica do padrão real da NF 16750: descrições diferentes
        intercaladas no XML não devem se misturar entre grupos.
        """
        itens = [
            {'desc_xml': 'PLANTA X', 'v_st_xml': 0.0, 'vProd': 100.0},
            {'desc_xml': 'FLOR X',   'v_st_xml': 0.0, 'vProd': 100.0},
            {'desc_xml': 'PLANTA X', 'v_st_xml': 0.0, 'vProd': 100.0},
        ]
        dados_api = [
            self._api_ant('PLANTA X', 10.0, 1.0),
            self._api_ant('FLOR X',   20.0, 2.0),
            self._api_ant('PLANTA X', 30.0, 3.0),
        ]

        resultado = parear_impostos_api(itens, dados_api)

        assert resultado[0][1] == pytest.approx(11.0)  # 1º PLANTA X
        assert resultado[1][1] == pytest.approx(22.0)  # FLOR X
        assert resultado[2][1] == pytest.approx(33.0)  # 2º PLANTA X

    def test_contagem_diferente_cai_em_rateio_e_sinaliza(self):
        """
        Quando N (itens XML) != M (registros API) para a mesma descrição,
        não arrisca pareamento posicional não verificável — rateia
        proporcional a vProd e marca 'rateado' para conferência manual.
        """
        desc = 'PRODUTO AMBIGUO CX10'
        itens = [
            {'desc_xml': desc, 'v_st_xml': 0.0, 'vProd': 100.0},
            {'desc_xml': desc, 'v_st_xml': 0.0, 'vProd': 300.0},
        ]
        # só 1 registro na API para 2 itens no XML → N=2, M=1
        dados_api = [self._api_ant(desc, 80.0, 20.0)]

        resultado = parear_impostos_api(itens, dados_api)

        vant0, vant1 = resultado[0][1], resultado[1][1]
        assert resultado[0][2] is True, "deveria estar marcado como rateado"
        assert resultado[1][2] is True
        # total do grupo preservado, dividido proporcionalmente a vProd (1:3)
        assert vant0 + vant1 == pytest.approx(100.0)
        assert vant0 == pytest.approx(25.0)   # 100 * (100/400)
        assert vant1 == pytest.approx(75.0)   # 100 * (300/400)


# ─── Valores por unidade (integração com divisão por quantidade) ──────────────

class TestValoresPorUnidade:
    """
    Valida os valores unitários que chegam à planilha após a divisão por qtd.
    Baseia-se nos dados reais da NF 13215 para confirmar que o fix eliminou
    o ratio >70% reportado pelo usuário.
    """

    @pytest.mark.parametrize("desc,v_st_xml,api_item,qcom,fecoep,ratio_max", [
        (
            'KIT PRENDEDOR DE CABELO E ESPONJA DE MAQUIAGEM CX180',
            287.93, KIT_PRENDEDOR_CX180_API, 360, 16.86, 0.50,
        ),
        (
            'PRENDEDOR DE CABELO CX144',
            89.58,  PRENDEDOR_CX144_API,      144, 5.25,  0.50,
        ),
    ])
    def test_ratio_st_por_unidade_abaixo_de_50pct(
        self, desc, v_st_xml, api_item, qcom, fecoep, ratio_max
    ):
        """
        Após o fix, st_u / nf_u deve ser menor que 50%.
        Antes do fix, era ~73% para esses produtos (ICMS dobrado).
        """
        vProd = {'KIT PRENDEDOR DE CABELO E ESPONJA DE MAQUIAGEM CX180': 810.0,
                 'PRENDEDOR DE CABELO CX144': 252.0}[desc]

        vST, _ = _merge_single(v_st_xml, desc, [api_item], vProd=vProd)

        nf_u  = round(vProd / qcom, 2)
        st_u  = round(vST   / qcom, 2)
        ratio = st_u / nf_u

        assert ratio < ratio_max, (
            f"st_u/nf_u={ratio:.0%} — ICMS provavelmente dobrado. "
            f"vST={vST:.2f}, nf_u={nf_u:.2f}"
        )


# ─── Fixtures de payloads API reais (NF 59263) ──────────────────────────────

VENTILADOR_PRETO_API = {
    'descricaoProduto':     'VENTILADOR PEDESTAL 220V - 50W - FUTURO - PRETO',
    'tipoImposto':          'ST',
    'valorIcmsCalculado':   2458.57,
    'valorFecoepCalculado': 142.83,
}
VENTILADOR_BRANCO_PRETO_API = {
    'descricaoProduto':     'VENTILADOR PEDESTAL 220V - 50W - FUTURO - BRANCO COM PRETO',
    'tipoImposto':          'ST',
    'valorIcmsCalculado':   1639.04,
    'valorFecoepCalculado': 95.22,
}


# ─── NF sem destaque de ST (vICMSST=0 no XML) ────────────────────────────────

class TestProdutoST_SemDestaqueNF:
    """
    Bug NF 59263 (2026-04-07): NF não destaca ST no XML (vICMSST=0), mas a API
    retorna os valores completos. O código anterior somava só FECOEP → st_u
    ficava com ~R$0,47 ao invés de ~R$8,67 por unidade.

    Quando v_st_xml=0 e tipoImposto='ST', deve usar ICMS + FECOEP da API.
    """

    def test_st_sem_destaque_xml_usa_icms_e_fecoep_api(self):
        """v_st_xml=0 → vST deve ser icms + fecoep da API (não só fecoep)."""
        vST, vANT = _merge_single(
            0.0,
            'VENTILADOR PEDESTAL 220V - 50W - FUTURO - PRETO',
            [VENTILADOR_PRETO_API],
        )
        assert vST  == pytest.approx(2458.57 + 142.83)  # 2601.40
        assert vANT == pytest.approx(0.0)

    def test_st_sem_destaque_xml_nao_usa_so_fecoep(self):
        """REGRESSÃO: comportamento antigo retornava apenas fecoep (142.83)."""
        vST, _ = _merge_single(
            0.0,
            'VENTILADOR PEDESTAL 220V - 50W - FUTURO - PRETO',
            [VENTILADOR_PRETO_API],
        )
        assert vST != pytest.approx(142.83), "st_u está usando apenas FECOEP — ICMS ausente"

    def test_st_sem_destaque_valor_unitario_correto(self):
        """
        Dados reais NF 59263: qty=300, vProd=6379.41, vICMSST=0.
        API: icms=2458.57, fecoep=142.83.
        Esperado: st_u = (2458.57 + 142.83) / 300 ≈ 8.67.
        """
        vST, _ = _merge_single(
            0.0,
            'VENTILADOR PEDESTAL 220V - 50W - FUTURO - PRETO',
            [VENTILADOR_PRETO_API],
        )
        st_u = round(vST / 300, 2)
        assert st_u == pytest.approx(8.67)

    def test_st_com_destaque_xml_nao_usa_icms_api(self):
        """
        Garante que o comportamento antigo (v_st_xml > 0) não regrediu:
        quando XML já tem vICMSST, ainda deve somar só FECOEP.
        """
        v_st_xml = 2458.57  # XML tem o valor
        vST, _ = _merge_single(
            v_st_xml,
            'VENTILADOR PEDESTAL 220V - 50W - FUTURO - PRETO',
            [VENTILADOR_PRETO_API],
        )
        # Correto: XML + fecoep apenas
        assert vST == pytest.approx(2458.57 + 142.83)
        # Errado seria: 2458.57 + 2458.57 + 142.83 (double-count)
        assert vST != pytest.approx(2458.57 + 2458.57 + 142.83)

    def test_dois_produtos_st_sem_destaque_nao_contaminam(self):
        """
        Dois produtos ST sem destaque na mesma NF, descrições distintas.
        Cada um deve receber apenas os valores do seu próprio grupo.
        """
        dados_api = [VENTILADOR_PRETO_API, VENTILADOR_BRANCO_PRETO_API]

        vST_preto, _ = _merge_single(
            0.0, 'VENTILADOR PEDESTAL 220V - 50W - FUTURO - PRETO', dados_api
        )
        vST_branco, _ = _merge_single(
            0.0, 'VENTILADOR PEDESTAL 220V - 50W - FUTURO - BRANCO COM PRETO', dados_api
        )

        assert vST_preto  == pytest.approx(2458.57 + 142.83)   # 2601.40
        assert vST_branco == pytest.approx(1639.04 + 95.22)    # 1734.26
