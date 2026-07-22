# TODO

## Bug: ST/ANT inflado em produtos com descrição duplicada na NF (merge_impostos_api)

**Onde:** `core/processador.py`, função `merge_impostos_api()` (linha ~190), chamada em `gerar_tabela()` (linha ~357).

**O que acontece:** quando uma NFe tem múltiplos itens `<det>` com a mesma `descricaoProduto` (produtos diferentes — REF/EAN distintos — mas mesmo texto de descrição, ex: variações de cor do mesmo item), o casamento com a API da SEFAZ AL é feito só por substring de texto (`api_desc in desc_upper`). Isso faz o código somar **todos** os registros da API daquela descrição em **cada** linha do XML que bate, em vez de dividir um registro por linha.

**Confirmado com a NF 16750** (`/Users/felipelopesfirmino/Downloads/33260649682710000138550010000167501001682880.xml`): "FLOR ARTIFICIAL CX120" tem 5 linhas no XML (REFs DV80496, DV80498, DV80500, DV80501, DV80503 — EANs diferentes), e a API retorna 5 registros ANT de R$146,80 cada (total real R$734,00). O código aplica R$734,00 (soma dos 5) em **cada uma** das 5 linhas — inflando o grupo para R$3.670,00, 5× o valor correto. Mesmo padrão em "ADESIVO DECORATIVO DE PVC PARA COZINHA CX48 30CM X 60CM" (5×), "BANDEJA DECORATIVA CX72" (5×), "PLANTA/FLOR ARTIFICIAL... CX288" e "ABAJUR DECORATIVO CX12" (2×).

**Por que a API não ajuda a resolver isso diretamente:** investigamos os 24 campos que a API devolve (`id, idNFe, chaveNota, numeroNota, indiceProduto, responsavel, aliquotaIcms, aliquotaFecoep, mvaValor, redutor, redutorCredito, segmento, pauta, codigoNcm, codigoCest, dataEmissao, ..., descricaoProduto, valorIcmsCalculado, valorFecoepCalculado, tipoImposto, numDocResponsavel`). Nenhum identifica o produto por REF/EAN. Os candidatos naturais (`indiceProduto`, `codigoCest`, `segmento`, `pauta`, `mvaValor`) vêm `null` nessa nota. O único campo que varia entre registros de mesma descrição é `id` (inteiro sequencial da tabela da SEFAZ — não é um identificador de produto, mas parece preservar a ordem de processamento dos itens da nota).

### Opção 1 — Pareamento posicional por fila (recomendada)

Para cada grupo de descrição igual, tratar os itens do XML (na ordem em que aparecem) e os registros da API (ordenados por `id` crescente) como duas filas e consumir um registro da API por item do XML (round-robin/pop), em vez de somar tudo em cada linha.

- Prós: resolve o caso 100% sem exigir nenhuma mudança externa; barato de implementar; não depende de a SEFAZ mudar nada.
- Contras: depende da suposição (não garantida contratualmente pela SEFAZ) de que a ordem de `id`/retorno da API acompanha a ordem dos itens no XML. Se a SEFAZ um dia mudar a ordem de processamento/retorno, o pareamento pode ficar errado silenciosamente. Precisa de teste de regressão específico para grupos de descrição duplicada.
- Esforço: baixo (algumas horas) — mexe só em `merge_impostos_api()` e no loop que a chama em `gerar_tabela()`.

### Opção 2 — Fallback com aviso/rateio proporcional quando há ambiguidade

Quando um grupo de descrição tiver N itens no XML e M registros na API com N == M, aplicar o pareamento posicional (Opção 1). Quando N != M ou a correspondência for ambígua, **não tentar adivinhar**: dividir o total do grupo proporcionalmente ao valor de compra (`vProd`) de cada linha, e sinalizar visualmente na planilha/dashboard (ex: coluna de auditoria ou alerta) que aquele produto teve valor de ST/ANT estimado por rateio, não vindo direto da API — para o usuário conferir manualmente.

- Prós: mais seguro quando a suposição de ordem falha ou quando N ≠ M (ex: SEFAZ agrupou/desmembrou itens de forma diferente do XML); dá transparência ao usuário em vez de errar silenciosamente.
- Contras: mais trabalho (precisa de UI/sinalização, lógica de rateio, mensagens de alerta); ainda existe uma zona cinzenta (rateio por valor não é garantidamente o critério real usado pela SEFAZ para dividir o imposto entre os itens).
- Esforço: médio — inclui a lógica da Opção 1 como caso feliz, mais tratamento do caso de ambiguidade e algum indicador visual novo.

**Decisão pendente:** escolher entre implementar só a Opção 1 (mais rápida, cobre o caso observado) ou já entrar com a Opção 2 (mais robusta a longo prazo, mas mais esforço).
