"""
Testes das funções utilitárias de baixo nível.

Cobre:
  - limpar_preco()   — normalização de valores monetários do CSV
  - limpar_str()     — normalização de strings (EAN, REF)
  - _arredondar_x9() — arredondamento para X.X9 usado no preço varejo
"""
import math
import pytest
from core.processador import limpar_preco, limpar_str
from app import _arredondar_x9


# ─── limpar_preco ─────────────────────────────────────────────────────────────

class TestLimparPreco:
    def test_formato_brasileiro_virgula(self):
        # Formato padrão do CSV: vírgula como decimal
        assert limpar_preco("12,50") == 12.50

    def test_ponto_e_tratado_como_milhar(self):
        # A função remove pontos antes de processar — comportamento intencional
        # para suportar "1.234,56" (formato BR). "12.50" vira 1250, não 12.5.
        assert limpar_preco("12.50") == 1250.0

    def test_ponto_milhar_virgula_decimal(self):
        # Formato brasileiro completo: "1.234,56"
        assert limpar_preco("1.234,56") == 1234.56

    def test_prefixo_rs(self):
        assert limpar_preco("R$ 25,99") == 25.99

    def test_nan(self):
        import math
        assert limpar_preco(float('nan')) == 0.0

    def test_string_vazia(self):
        assert limpar_preco("") == 0.0

    def test_string_invalida(self):
        assert limpar_preco("abc") == 0.0

    def test_zero_string(self):
        assert limpar_preco("0") == 0.0

    def test_valor_inteiro(self):
        assert limpar_preco("100") == 100.0


# ─── limpar_str ──────────────────────────────────────────────────────────────

class TestLimparStr:
    def test_string_normal(self):
        assert limpar_str("ABC123") == "ABC123"

    def test_remove_aspas_e_igual(self):
        # CSVs exportados às vezes encapsulam: ="7891234"
        assert limpar_str('="7891234"') == "7891234"

    def test_remove_ponto_zero_float(self):
        # Pandas lê EANs inteiros como float: 7891234567890.0
        assert limpar_str("7891234567890.0") == "7891234567890"

    def test_strip_espacos(self):
        assert limpar_str("  REF123  ") == "REF123"

    def test_nan_retorna_vazio(self):
        assert limpar_str(float('nan')) == ""

    def test_string_vazia_retorna_vazio(self):
        assert limpar_str("") == ""


# ─── _arredondar_x9 ──────────────────────────────────────────────────────────

class TestArredondarX9:
    """
    Lógica: int(val * 10) / 10 + 0.09
    Garante preços terminando em X.X9 (ex: 12.99, 34.09, 7.79).
    """

    def test_valor_basico(self):
        # 12.2 → int(122)/10 + 0.09 = 12.2 + 0.09 = 12.29
        assert _arredondar_x9(12.2) == pytest.approx(12.29)

    def test_valor_inteiro(self):
        # 10.0 → int(100)/10 + 0.09 = 10.0 + 0.09 = 10.09
        assert _arredondar_x9(10.0) == pytest.approx(10.09)

    def test_valor_alto(self):
        # 46.39 → int(463.9)/10 + 0.09 = 46.3 + 0.09 = 46.39
        assert _arredondar_x9(46.39) == pytest.approx(46.39)

    def test_valor_alto_arredondado(self):
        # 46.4 → int(464)/10 + 0.09 = 46.4 + 0.09 = 46.49
        assert _arredondar_x9(46.4) == pytest.approx(46.49)

    def test_zero_retorna_zero(self):
        assert _arredondar_x9(0) == 0.0

    def test_negativo_retorna_zero(self):
        assert _arredondar_x9(-5.0) == 0.0

    def test_valor_pequeno(self):
        # 0.5 → int(5)/10 + 0.09 = 0.5 + 0.09 = 0.59
        assert _arredondar_x9(0.5) == pytest.approx(0.59)
