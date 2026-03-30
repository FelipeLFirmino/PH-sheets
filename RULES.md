# RULES.md — Regras de Negócio da Precificadora PH

Este arquivo é a fonte de verdade para toda lógica de classificação tributária e cálculo de preços.
Qualquer alteração no código deve ser validada contra as regras aqui documentadas.

---

## 1. Taxonomia de Regimes de ICMS (Legislação Brasileira)

Uma NF-e pode conter produtos em qualquer um dos seguintes regimes. O sistema deve tratá-los de forma independente — **não existe apenas ST e ANT**.

| Regime | CST Típico | Descrição |
|--------|-----------|-----------|
| **Normal** | 00 | Débito/crédito padrão. Cada elo da cadeia paga seu próprio ICMS na venda. |
| **ST** (Substituição Tributária) | 10 / 60 | Um contribuinte (geralmente indústria/importador) recolhe o ICMS de toda a cadeia antecipadamente, com base em uma MVA presumida. |
| **ANT** (Antecipação Tributária) | — | Ocorre quando a mercadoria chega a um estado sem acordo inter-estadual de ST vigente, ou quando a ST não foi retida na origem. O destinatário recolhe ao entrar no estado. |
| **Isenção** | 40 / 41 | Produto isento de ICMS. Nenhum imposto é devido. |
| **Diferimento** | 51 | Pagamento adiado para etapa posterior da cadeia (ex.: frigorífico paga pelo pecuarista). |
| **Suspensão** | 50 | Tributação suspensa por prazo determinado. |
| **Redução de base** | 20 | Isenção parcial via redução da base de cálculo. |

> **Base legal:** CF Art. 150 §7 (ST), Convênio ICMS 52/2017 (ANT), RIPI/RICMS estaduais.

---

## 2. Regras de Classificação por Produto

### 2.1 Como identificar o regime de cada produto

```
Produto tem ST?
    → st_u > 0  (valor vindo do XML campo vICMSST + API SEFAZ)
    → Regime: ST

Produto tem ANT?
    → ant_u > 0  (valor vindo exclusivamente da API SEFAZ, tipoImposto = 'ANT')
    → Regime: ANT

Produto tem CST isento?
    → CST IN [40, 41, 50, 51, 60]
    → Regime: Isento/Diferido/Suspenso

Nenhum dos anteriores?
    → Regime: Normal (CST 00 — débito/crédito padrão)
```

### 2.2 Threshold de comparação

Valores `st_u` e `ant_u` são arredondados a 2 casas decimais em `processador.py`.
Para evitar falsos positivos por arredondamento de float, usar **`> 0.005`** como threshold em todas as comparações, nunca `> 0`.

```python
# CORRETO
has_st  = row['st_u']  > 0.005
has_ant = row['ant_u'] > 0.005

# ERRADO — pode capturar near-zeros de floating point
has_st  = row['st_u']  > 0
```

### 2.3 Um produto pode ter ST e ANT ao mesmo tempo?

Não na prática — ou o ICMS foi retido na origem (ST) ou será retido na entrada do estado (ANT). Se ambos aparecerem, prevalece o tratamento de ST.

---

## 3. Regra do Crédito de ICMS (CRED ICMS)

Esta é a regra mais crítica e com maior histórico de erro.

### 3.1 Definição

O crédito de ICMS é o valor que o comprador pode abater do ICMS que vai pagar na revenda. Ele existe porque o ICMS é um imposto não-cumulativo — você desconta o que pagou na compra do que deve na venda.

### 3.2 Quando o crédito é ZERO

| Condição | Motivo | CST |
|----------|--------|-----|
| `st_u > 0.005` (produto com ST) | O ICMS de toda a cadeia já foi pago pelo substituto. Não há crédito a recuperar pois o produto sai sem débito de ICMS. | 60 |
| `CST = 40` | Isento — não há débito na saída, logo não há crédito na entrada. | 40 |
| `CST = 41` | Não tributado — mesmo raciocínio. | 41 |
| `CST = 50` | Suspensão — tributação suspensa, crédito suspenso junto. | 50 |
| `CST = 60` | ST anteriormente cobrada — mesmo que ST acima. | 60 |
| `CST = 102` / `500` | Simples Nacional — regime simplificado, sem crédito de ICMS. | CSOSN |

### 3.3 Quando o crédito NÃO é zerado

| Condição | Motivo |
|----------|--------|
| `ant_u > 0.005` (ANT padrão, sem encerramento) | **ANT padrão NÃO elimina o crédito.** Conforme SEFAZ AL FAQ (Q.6) e Lei 6.474/2004: *"A antecipação prevista não encerra a fase de tributação."* O comprador continua no regime débito/crédito normal e pode creditar o ICMS destacado na NF do fornecedor. |
| `CST = 00` (Normal) | Regime padrão — crédito integral. |
| `CST = 10` | Tributado + ST nas operações subsequentes — o vendedor ainda débita ICMS, gerando crédito. |
| `CST = 20` | Redução de base — crédito proporcional à base reduzida. |

### 3.4 Exceção: ANT com encerramento de tributação

Alguns produtos específicos têm **ANT com encerramento** — funciona igual ao ST, fecha a cadeia tributária e **elimina o crédito**. Em Alagoas, o exemplo atual é **calçados** (Art. 591-H do RICMS/AL + Anexo XXXVI).

> **Atenção:** o sistema atual **não distingue** ANT padrão de ANT com encerramento — trata todo ANT como padrão (com crédito). Caso o negócio trabalhe com calçados ou outros produtos com ANT com encerramento, essa lógica precisará de revisão.

### 3.5 O que o comprador NÃO pode fazer com crédito em produtos ANT

O Art. 433, §2° do RICMS/AL proíbe usar créditos acumulados para compensar ou deduzir o **valor da antecipação em si** — a ANT é uma dívida autônoma paga via DAR (código 1542-3). Mas o crédito do ICMS da NF de compra continua existindo normalmente para abater o ICMS das vendas futuras.

### 3.6 Implementação correta

```python
# app.py — _calcular()  (prévia/dashboard)
# CORRETO:
if st_u > 0.005 or cst in {'40', '41', '50', '60', '102', '500'}:
    cred = 0.0
else:
    cred_pct = row.get('cred_pct', P['cred'])
    cred = round(nf_u * cred_pct, 2)

# processador.py — coluna % CRED ICMS (col 13)
# A coluna armazena o PERCENTUAL por produto (0.04, 0.07, 0.12, 0.19 ou 0.0).
# NÃO armazena o valor monetário. Cada célula recebe cor conforme paleta §16.
cred_rate = 0.0 if (st_u > 0.005 or cst in isentos) else row['cred_pct']
cell.value          = cred_rate      # ex: 0.04
cell.number_format  = '0.00%'        # exibe "4,00%"
cell.fill           = cred_fill(cred_rate)  # cor da paleta

# processador.py — fórmula C_ENT (col 15)
# O crédito é descontado inline na fórmula de custo de entrada:
=ROUND(C_REAL + ST + ANT + IPI + FRETE + DESP - ROUND(NF_U × CRED%, 2), 2)
# Nota: H (ANT) NÃO zera o crédito — ANT não está na condição de isenção.
```

### 3.7 Crédito por origem do produto

| Origem | % Crédito | 1º dígito do CST |
|--------|-----------|-----------------|
| Nacional | 19% (AL) / 12% (PE) / 7% (SP) | 0 |
| Importado / conteúdo importado ≥ 40% | 4% | 1, 2, 3, 6, 7, 8 |

> O percentual padrão configurável é 4% (importado), ajustável no formulário por nota.

---

## 4. Fluxo de Cálculo por Produto (Ordem Obrigatória)

Os passos DEVEM ser executados nesta ordem — dependências entre eles não permitem reordenação.

```
ENTRADA: nf_u, st_u, ant_u, ipi_u, cst, qtd, p_atual, params (mult, frete%, desp%, cred%, fed%, icms%, cart%)

Passo 1 — CUSTO REAL
    c_real = ROUND(nf_u × mult, 2)
    ← Multiplicador aplicado PRIMEIRO, antes de qualquer percentual

Passo 2 — FRETE
    frete = ROUND(c_real × frete%, 2)

Passo 3 — DESPESA
    desp = ROUND(c_real × desp%, 2)

Passo 4 — CRÉDITO ICMS
    SE st_u > 0.005 OU cst IN isentos:
        cred = 0.0
    SENÃO:          ← inclui ANT — ANT não zera o crédito
        cred = ROUND(nf_u × cred%, 2)

Passo 5 — CUSTO ENTRADA
    c_ent = ROUND(c_real + st_u + ant_u + ipi_u + frete + desp - cred, 2)
    ← cred é subtraído (benefício fiscal)
    ← ST, ANT e IPI são somados (custos reais de entrada)

Passo 6 — PREÇO VAREJO
    SE p_atual > 0:
        p_var = p_atual
    SENÃO:
        p_var = arredondar_99(c_real × 2)
        arredondar_99(x) = INT(x) - IF(x - INT(x) <= 0.5, 1, 0) + 0.99

Passo 7 — FEDERAL
    fed = ROUND(p_var × fed%, 2)

Passo 8 — CARTÃO
    cart = ROUND(p_var × cart%, 2)

Passo 9 — ICMS SAÍDA
    icms_s = MAX(0, ROUND(p_var × icms%, 2) − ant_u)
    ← ANT já pago na entrada é abatido do ICMS da saída (funciona como crédito)

Passo 10 — CUSTO SAÍDA
    c_saida = ROUND(c_ent + fed + cart + icms_s, 2)

Passo 11 — META %
    meta = 0.15 (padrão editável por produto na planilha)

Passo 12 — PREÇO MÍN VIÁVEL
    SE meta > 0:
        p_min = ROUND(c_saida / (1 - meta), 2)
    SENÃO:
        p_min = 0

Passo 13 — MARGEM REAL
    SE p_var > 0:
        margem = ROUND((p_var - c_saida) / p_var, 4)
    SENÃO:
        margem = 0

Passo 14 — LUCRO ESTIMADO (dashboard/preview)
    lucro = ROUND((p_var - c_saida) × qtd, 2)

--- ATACADO (passos adicionais, mesma entrada) ---

Passo 15 — NF ATC (custo atacado)
    nf_atc = ROUND(nf_u × mult_atc, 2)
    ← Multiplicador atacado separado (padrão 1.3 — menor que varejo 2.0)

Passo 16 — PREÇO ATC
    p_atc = ROUND(p_var × (1 - desc_atc), 2)
    ← Preço atacado = preço varejo com desconto (padrão 15%)

Passo 17 — FEDERAL ATC
    fed_atc = ROUND(nf_atc × fed%, 2)
    ← Base: NF_ATC (custo atacado), não preço de venda

Passo 18 — CARTÃO ATC
    cart_atc = ROUND(p_atc × cart%, 2)
    ← Base: P_ATC (preço de venda atacado)

Passo 19 — ICMS ATC
    icm_atc = MAX(0, ROUND(nf_atc × icms%, 2) − ant_u)
    ← Base: NF_ATC; ANT também abate o ICMS atacado

Passo 20 — CUSTO SAÍDA ATC
    c_saida_atc = ROUND(c_ent + fed_atc + cart_atc + icm_atc, 2)
    ← Usa o mesmo C_ENT do varejo — o custo de entrada é o mesmo

Passo 21 — MARGEM ATC
    SE p_atc > 0:
        margem_atc = ROUND((p_atc - c_saida_atc) / p_atc, 4)
    SENÃO:
        margem_atc = 0
```

---

## 5. Regras da API SEFAZ AL

### 5.1 Dados retornados

A API retorna uma lista de itens por NF. Cada item tem:
- `descricaoProduto` — descrição do produto no sistema SEFAZ
- `tipoImposto` — `'ST'` ou `'ANT'`
- `valorIcmsCalculado` — valor principal do imposto
- `valorFecoepCalculado` — fundo de combate à pobreza (soma com ICMS)

O valor total do imposto por item = `valorIcmsCalculado + valorFecoepCalculado`.

### 5.2 Matching produto XML × item API

O matching atual usa substring: `api_desc IN desc_xml`. Isso é frágil — uma descrição curta da API pode bater em produto errado.

**Regra:** Se um produto receber `ant_u > 0` mas o usuário confirmar que não é ANT, o matching da API pode estar errado. Verificar com o log de terminal se a `api_desc` realmente corresponde ao produto.

### 5.3 Falha da API

A API tem timeout de 10s. Se falhar:
- `v_st_api = 0`, `v_ant_api = 0`
- O processamento continua usando apenas `vICMSST` do XML
- **Nunca** abortar por falha da API

---

## 6. Regras de Preço e Fallback

### 6.1 Busca de preço no CSV do sistema

```
Tentativa 1: JOIN por EAN (campo BARRA no CSV == cEAN no XML)
    → Se preço encontrado e > 0.01: usar

Tentativa 2 (fallback): JOIN por REF (campo REFERÊNCIA no CSV == cProd no XML)
    → Se preço encontrado e > 0.01: usar

Tentativa 3 (fallback final): sem preço no sistema
    → p_atual = 0
    → p_var = arredondar_99(c_real × 2)  ← preço calculado automático
```

### 6.2 Detecção de embalagem

```
Regex: (?:CAIXA COM|PACOTE COM|KIT COM|PCT\s*C/|CX\s*C/|C/)\s*(\d+)

SE match e qtd_emb > 1:
    Proteção anti-absurdo:
        SE p_sys × qtd_emb > nf_u × mult × 3:
            qtd_emb = 1   ← sistema já tem preço da caixa completa
        SENÃO:
            p_atual = p_sys × qtd_emb
```

### 6.3 Arredondar 99

Função aplicada quando não há preço no sistema:
```
arredondar_99(x):
    resultado = INT(x) - (1 se x - INT(x) <= 0.5, senão 0) + 0.99

Exemplos:
    x = 10.00 → INT=10, decimal=0.00 ≤ 0.5 → 10 - 1 + 0.99 = 9.99
    x = 10.60 → INT=10, decimal=0.60 > 0.5 → 10 - 0 + 0.99 = 10.99
    x = 11.50 → INT=11, decimal=0.50 ≤ 0.5 → 11 - 1 + 0.99 = 10.99
```

---

## 7. Regras do Dashboard

### 7.1 Classificação de produtos (cards e barra de distribuição)

```python
com_st    = count(st_u  > 0.005)
com_ant   = count(ant_u > 0.005 AND st_u <= 0.005)   # ANT puro, sem ST
sem_ambos = total - com_st - com_ant                  # Normal, Isento, etc.
```

> **Regra:** ST e ANT são contadores independentes. Um produto não pode ser contado em ambos.
> Se tiver `st_u > 0` E `ant_u > 0`, é classificado como ST (tem prioridade).

### 7.2 Barra de distribuição

Três segmentos, em ordem:
```
[== Normal (azul) ==][= ST (pêssego) =][= ANT (verde menta) =]
```
Segmentos com 0% são omitidos para não quebrar o layout.

### 7.3 Margem estimada no dashboard

Calculada comparando preço atual do sistema com custo NF (simplificado, sem todos os impostos de saída):
```
margem_estimada = mean( (p_atual - nf_u) / p_atual )  para produtos com p_atual > 0
```
Esta é uma estimativa rápida — a margem real está na coluna X da planilha.

---

## 8. Regras de Cores (Preview e Excel)

### 8.1 Preview HTML (por linha)

| Condição | Cor de fundo das células |
|----------|--------------------------|
| `st_u > 0.005` | Pêssego `#FCE4D6` — linha inteira |
| `ant_u > 0.005` (sem ST) | Verde menta `#D1FAE5` — linha inteira |
| Normal | Sem cor de fundo |

### 8.2 Preview HTML (por coluna, apenas linhas normais)

A preview tem **cabeçalho duplo**: a primeira linha agrupa as colunas com rótulos "VAREJO" (fundo azul) e "ATACADO" (fundo amarelo). A segunda linha traz os nomes individuais.

| Coluna (índice 0-based) | Cor | Grupo |
|------------------------|-----|-------|
| 20 — PREÇO MÍN | Amarelo `#FFF2CC` + bold | Varejo |
| 22 — PREÇO VAREJO | Azul `#BDD7EE` + bold | Varejo |
| 23 — MARGEM | Azul `#BDD7EE` + bold | Varejo |
| 24 — NF ATC | Amarelo suave `#F3E8FF` | Atacado |
| 25 — FEDERAL ATC | Amarelo suave `#F3E8FF` | Atacado |
| 26 — CARTÃO ATC | Amarelo suave `#F3E8FF` | Atacado |
| 27 — ICMS ATC | Amarelo suave `#F3E8FF` | Atacado |
| 28 — C. SAÍDA ATC | Amarelo suave `#F3E8FF` | Atacado |
| 29 — PREÇO ATC | Amarelo `#FFF2CC` + bold | Atacado |
| 30 — MARGEM ATC | Amarelo `#FFF2CC` + bold | Atacado |

> ST e ANT têm prioridade — linha pêssego/verde menta sobrepõe todas as cores de coluna.
> A faixa inteira de colunas atacado (24-30) recebe `#F3E8FF` em linhas normais para separação visual, exceto PREÇO ATC e MARGEM ATC que recebem amarelo forte.

### 8.3 Excel (openpyxl)

| Cor | Hex | Aplicação |
|-----|-----|-----------|
| Azul | `#BDD7EE` | PREÇO VAREJO + MARGEM REAL (linhas normais) |
| Amarelo | `#FFF2CC` | PREÇO MÍN VIÁVEL + PREÇO ATC + MARGEM ATC (linhas normais) |
| Verde menta | `#D1FAE5` | Linha inteira de produto ANT (ant_u > 0.005 e st_u ≤ 0.005) |
| Verde | `#E2EFDA` | META % (editável por produto — todas as linhas) |
| Pêssego | `#FCE4D6` | Linha inteira de produto ST (st_u > 0.005) |
| Cinza | `#D9D9D9` | Colunas de auditoria P.UNIT SISTEMA e QTD EMB |
| Param | `#F2F2F2` | Linha 1 (cabeçalhos) e linha 2 (parâmetros) |

---

## 9. Invariantes — Nunca Quebrar

Estas são restrições que o código deve sempre respeitar:

1. **`st_u` nunca zera o crédito de ANT** — ANT não é ST, não elimina o direito ao crédito de ICMS.
2. **Threshold `> 0.005` em todas as comparações de imposto** — nunca usar `> 0` diretamente em floats vindos de divisão ou API.
3. **A API SEFAZ nunca pode travar o processamento** — sempre em try/except com fallback para zero.
4. **`p_var` nunca pode ser 0** — se não há preço no sistema, usar `arredondar_99(c_real × 2)`.
5. **Multiplicador aplicado antes de qualquer percentual** — C_REAL = NF_U × MULT é sempre o Passo 1.
6. **Crédito de ICMS é subtraído** no C_ENT via `ROUND(NF_U × CRED%, 2)` — a coluna CRED armazena o **percentual** (não o valor). A fórmula do C_ENT desconta o crédito inline.
7. **Federal (varejo), Cartão (varejo) e ICMS Saída (varejo) incidem sobre P_VAR** (preço de venda varejo).
8. **Federal ATC e ICMS ATC incidem sobre NF_ATC** (custo atacado) — base de cálculo diferente do varejo.
9. **Cartão ATC incide sobre P_ATC** (preço de venda atacado) — mesma lógica do cartão varejo mas com base atacado.
10. **C_ENT (custo de entrada) é COMPARTILHADO** entre varejo e atacado — a diferença está nos tributos de saída.
11. **ANT colorido como verde menta, ST como pêssego** — nunca misturar as cores dos dois regimes.
12. **ICMS Saída é ZERO quando há ST** — produto com ST já teve ICMS recolhido na cadeia; na revenda sai sem débito (CST 60 na saída). Aplica-se varejo e atacado.
13. **Custos NF divididos pela embalagem** — NF_U, ST_U, ANT_U e IPI_U são divididos por qtd_emb. Tudo por UNIDADE vendida.
14. **Preço do sistema é SEMPRE por unidade** — nunca multiplicar p_sys pela embalagem.
15. **Crédito de ICMS vem do XML por produto** — pICMS de cada item do XML. Fallback pro parâmetro do formulário se XML não tiver.
16. **Sem preço real (≤ R$ 0,02), não dividir por embalagem** — p_sys placeholder não permite validar anti-absurdo. EMB=1 é mais seguro.
17. **Frete CIF = 0%** — se modFrete=0 no XML, o fornecedor paga o frete. Parâmetro do formulário é ignorado.

---

## 10. Detecção de Embalagem

### 10.1 Padrões reconhecidos (3 padrões, nesta ordem)

| Prioridade | Padrão | Exemplos | Fornecedores |
|-----------|--------|----------|-------------|
| 1 (explícito) | `CAIXA COM N`, `PCT C/N`, `KIT COM N`, `CX C/N`, `C/N` | "CAIXA COM 12", "PCT C/24" | Diversos |
| 2 (início) | `N PRODUTO...` (número no início, sem sufixo de medida) | "12 CANECAS...", "6 TIGELAS..." | Oxford, Biona |
| 3 (final) | `PT N UN`, `DP N UN`, `POTE N UN`, `POLYBAG N UN`, `DISPLAY N UN` | "PT24UN", "DP 18 UN", "POTE 48 UN" | Summit, papelarias |
| 4 (DS Display) | `DS NOME - N` ou `DS NOME - N PCS` | "DS FESTIVAL LAVANDERIA LAVANDA - 84", "DS ORG GIRE E TRAVE M - 207 PCS" | Fornecedores que usam "DS" (Display) |

> **Padrão 2 — sufixos excluídos:** se o número for seguido de `ML`, `MG`, `KG`, `GR`, `LT`, `L`, `CM`, `MM`, `M`, `UN`, `PC`, `PCS`, `PCT`, `UNID`, o padrão NÃO é aplicado (são medidas, não contagens). Ex: "2 LITROS", "500 ML", "250 G" → `qtd_emb = 1`.

> **Padrão 4 — DS Display:** `DS` no início indica "Display" (embalagem de exposição). O número após o traço final é a quantidade de unidades no display. Sufixo `PCS` é opcional e ignorado. Tratado como explícito (não requer `p_sys` para validação — apenas os anti-absurdos numéricos se aplicam).

### 10.2 Anti-absurdo (3 travas)

```
1. SE qtd_emb <= 1 → EMB = 1
2. SE padrão_ambíguo (Padrão 2) E p_sys <= 0.02 → EMB = 1
   ← Padrões 1, 3 e 4 são explícitos o suficiente para dispensar essa trava
3. SE p_sys > 0 E p_sys × qtd_emb > vUnCom × mult × 3 → EMB = 1
```

> **Padrões 1, 3 e 4 (CAIXA COM, POTE N UN, DS NOME - N etc.) não requerem preço no sistema** para aplicar a embalagem — a descrição já é inequívoca. O Padrão 2 (número no início) é ambíguo e continua exigindo `p_sys > 0.02`.

### 10.3 Divisão

Quando `qtd_emb > 1`: nf_u, st_u, ant_u, ipi_u divididos por qtd_emb.
Preço do sistema (`p_sys`) NÃO é multiplicado — já é por unidade.

---

## 11. Crédito de ICMS Automático

O percentual de crédito é extraído do campo `pICMS` de cada item do XML. Cada produto pode ter % diferente na mesma nota (ex: 7% nacionais, 4% importados).

```
SE pICMS > 0 no XML → cred_pct = pICMS / 100 (por produto)
SENÃO → cred_pct = parâmetro do formulário (fallback)
```

O crédito é gravado como **valor monetário negativo** na coluna `CRED ICMS` (col 13), ex: `-3,12`. A **cor da célula** indica qual percentual foi usado (paleta §16). O valor é calculado como `ROUND(NF_U × cred_pct, 2)` e somado ao C_ENT (como negativo, reduz o custo). A legenda de cores é exibida automaticamente abaixo da tabela no Excel sempre que houver dados na NF.

---

## 12. Frete CIF/FOB Automático

O campo `modFrete` do XML indica quem paga o frete:

| modFrete | Tipo | Comportamento |
|----------|------|--------------|
| 0 | CIF | Emitente paga → frete = 0% (ignora formulário) |
| 1 | FOB | Destinatário paga → usa % do formulário |
| 2 | Terceiros | Usa % do formulário |
| 9 | Sem frete | Usa % do formulário |

---

## 13. Arredondamento X.X9

Quando não há preço no sistema, o preço varejo é calculado como C_REAL × 2 arredondado pro próximo centavo terminado em 9:

```
arredondar_x9(x):
    SE x é inteiro (ex: 10.00) → permanece igual
    SENÃO → INT(x × 10) / 10 + 0.09

Exemplos:
    6.72 → 6.79
    3.10 → 3.19
    9.00 → 9.00 (inteiro, permanece)
    10.48 → 10.49
    25.96 → 25.99
```

---

## 14. Colunas de Pacote (Atacado)

Duas colunas após MARGEM ATC mostram os valores **por caixa/pacote completo** (usando `QTD_EMB`):

| Coluna | Fórmula | Cor | Significado |
|--------|---------|-----|-------------|
| **PREÇO PCT ATC** (col 32) | `P_ATC × QTD_EMB` | Amarelo `#FFF2CC` | Quanto o cliente **paga** pela caixa no atacado |
| **P. COMPRA PCT** (col 33) | `NF_U × QTD_EMB` | Roxo suave `#F3E8FF` | Quanto você **pagou** pela caixa ao fornecedor |

Se `QTD_EMB = 1` (produto unitário), os valores coincidem com P_ATC e NF_U respectivamente.
Útil para: cotação de atacado, comparação custo × venda por pacote, margem visual.

---

## 15. Dashboard — Cards Exibidos

| Card | Cálculo |
|------|---------|
| Total Itens | `len(rows)` |
| Com ST | `count(st_u > 0.005)` |
| Com ANT | `count(ant_u > 0.005 and st_u <= 0.005)` |
| Sem Preço no Sistema | `count(p_atual <= 0)` |
| Embalagens | `count(qtd_emb > 1)` de N |
| Crédito ICMS | Badges coloridos por faixa (ex: `30× 4%` em laranja, `10× 7%` em orquídea) — paleta §16 |
| Lucro Estimado | `sum((p_var - c_saida) × qtd)` |

Cards removidos: "Normal" (não relevante), "Valor Total NF" (não batia com XML), "Margem Estimada" (impreciso).

---

## 16. Paleta de Cores — Crédito ICMS

A coluna `% CRED ICMS` (col 13) usa cores distintas por faixa de percentual. As cores aparecem:
- Na **célula individual** da coluna CRED de cada produto (Excel e preview HTML)
- Nos **badges do card** "Crédito ICMS" no dashboard
- Na **legenda automática** abaixo da tabela no Excel (gerada apenas quando há mais de uma faixa na NF)

| % Crédito | Cor | Hex | Origem típica |
|-----------|-----|-----|---------------|
| **4%** | Laranja vivo | `#FFB347` | Importado / conteúdo importado ≥ 40% |
| **7%** | Orquídea | `#DA70D6` | Nacional — SP |
| **12%** | Turquesa | `#40E0D0` | Nacional — PE |
| **19%** | Rosa pink | `#FF69B4` | Local — AL |
| **0%** | Cinza médio | `#B0B0B0` | ST / Isento (sem crédito) |

> Essas cores não colidem com nenhuma outra cor do sistema (ST-pêssego, ANT-verde menta, VAREJO-azul, PREÇO-amarelo, META-verde claro, ATACADO-roxo suave).

### 16.1 Legenda automática no Excel

Gerada automaticamente logo abaixo da tabela de dados quando a NF contém **mais de uma faixa** de crédito ICMS. Cada linha da legenda mostra:
- Swatch colorido com o percentual (ex: `4%` em laranja)
- Descrição (ex: `4% — Importado`)

Quando a NF tem apenas uma faixa, a legenda é omitida (desnecessária).

### 16.2 Invariante de cores CRED

A cor da célula CRED **nunca é sobrescrita** pelas cores de linha ST (pêssego) ou ANT (verde menta). A coluna CRED está no conjunto `_skip` do loop de colorização de linhas.