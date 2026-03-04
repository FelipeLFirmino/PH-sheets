
 Precificadora Inteligente

Este projeto é uma ferramenta local de organização financeira e precificação automatizada. Ele processa arquivos **XML (NFe)** e **CSV (Sistema)** para gerar planilhas de precificação detalhadas em formato Excel, aplicando regras de negócio específicas para varejo e atacado.

## 🚀 Tecnologias Utilizadas

* **Backend:** Python 3.x com Flask.
* **Processamento de Dados:** Pandas para manipulação de tabelas e cruzamento de dados.
* **Manipulação de XML:** ElementTree para extração de dados da NFe.
* **Geração de Excel:** Openpyxl para aplicação de estilos, bordas e cores.
* **Frontend:** HTML5, CSS3 (Bootstrap) e JavaScript nativo para suporte a Drag & Drop.

## 🛠️ Instalação e Configuração

Siga os passos abaixo para configurar o ambiente em sua máquina local (macOS com chip M1/M2/M3):

1. **Clonar o repositório:**
```bash
git clone https://github.com/seu-usuario/ph-SheetGenerator.git
cd ph-SheetGenerator

```


2. **Criar e ativar o ambiente virtual:**
```bash
python3 -m venv venv
source venv/bin/activate

```


3. **Instalar as dependências:**
```bash
pip install Flask pandas openpyxl

```



## 📖 Como Usar

1. Execute o servidor Flask:
```bash
python3 app.py

```


2. O navegador abrirá automaticamente no endereço `http://127.0.0.1:5000`.
3. Insira o **Nome do Fornecedor** e o **Número da Nota**.
4. Arraste o arquivo **XML da NFe** e o **CSV do seu sistema** para as áreas indicadas.
5. Ajuste os parâmetros financeiros (Margens, Impostos, Frete) conforme necessário.
6. Clique em **Gerar Prévia da Planilha** para visualizar os dados na tela.
7. Clique em **Confirmar e Baixar Excel** para salvar o arquivo estilizado.

## 🧠 Lógica de Processamento e Regras de Negócio

O motor de cálculo utiliza uma lógica de busca resiliente para garantir que o preço atual do produto seja encontrado no sistema:

* **Mecanismo de Fallback:** O sistema tenta primeiro cruzar os dados pelo **EAN (Código de Barras)**. Caso não encontre um preço válido, ele realiza uma segunda tentativa utilizando a **Referência/Código do Produto**.
* **Cálculo de Impostos:** ICMS, Impostos Federais e Taxas de Cartão são calculados com base no preço de venda final.
* **Arredondamento de Marketing:** Preços de varejo são automaticamente arredondados para a terminação `.99`.

## 🎨 Formatação da Planilha (Excel)

A planilha gerada segue um padrão visual rígido para facilitar a análise:

* **Bordas:** Todas as células possuem bordas finas pretas para melhor legibilidade.
* **Formatação Numérica:** Colunas de Margem Varejo e Margem Atacado são formatadas nativamente como porcentagem (`%`).
* **Destaque por Cores:**
* **Azul:** Colunas de Produto, Preço e Margem de Varejo.
* **Amarelo:** Colunas de Produto, Desconto e Margem de Atacado.
* **Cor de Pele (Pêssego):** Linhas completas de produtos identificados com Substituição Tributária (ST > 0).



## 📂 Estrutura do Projeto

```text
ph-SheetGenerator/
├── app.py              # Ponto de entrada e servidor Flask
├── core/
│   ├── __init__.py
│   └── processador.py  # Motor de cálculo e estilização Excel
├── static/
│   └── logo.png        # Logo da aplicação
└── templates/
    └── index.html      # Interface web (Frontend)

```

---
Para transformar sua aplicação Flask em um executável (.exe para Windows ou .app para Mac) de forma que o usuário precise apenas dar um duplo clique, utilizaremos o **PyInstaller**. Como você é desenvolvedor Fullstack, o processo será simples, mas exige um ajuste técnico no seu `app.py` para que ele encontre as pastas de HTML e CSS dentro do arquivo compactado.

### 1. Ajuste Necessário no `app.py`

Quando o PyInstaller cria um executável, ele descompacta os arquivos em uma pasta temporária. Você precisa avisar o Flask onde encontrar essa pasta. Altere o início do seu `app.py`:

```python
import sys
import os

# Determina o caminho base (se está rodando como script ou executável)
if getattr(sys, 'frozen', False):
    template_folder = os.path.join(sys._MEIPASS, 'templates')
    static_folder = os.path.join(sys._MEIPASS, 'static')
    app = Flask(__name__, template_folder=template_folder, static_folder=static_folder)
else:
    app = Flask(__name__)

```

---

### 2. Fluxo de Instalação na Máquina do Usuário

Como você pretende baixar do GitHub e gerar o executável na máquina dele, siga este roteiro:

#### Passo A: Preparação (Windows)

1. Instale o **Python 3.12** (ou a versão estável que você preferir) no computador do usuário. Marque a opção "Add Python to PATH".
2. Abra o Terminal (PowerShell ou CMD) na pasta onde deseja colocar o projeto.

#### Passo B: Clonagem e Dependências

```powershell
# Baixe o código
git clone https://github.com/seu-usuario/ph-SheetGenerator.git
cd ph-SheetGenerator

# Crie o ambiente e instale tudo
python -m venv venv
.\venv\Scripts\activate
pip install Flask pandas openpyxl pyinstaller

```

#### Passo C: Gerar o Executável

Execute o comando abaixo. Note que no Windows usamos `;` para separar os caminhos, e no Mac usamos `:`.

**Comando para Windows:**

```powershell
pyinstaller --noconfirm --onefile --windowed --add-data "templates;templates" --add-data "static;static" --icon "static/logo.png" app.py

```

* `--onefile`: Cria um único arquivo .exe (mais fácil para o usuário).
* `--windowed`: Impede que uma tela preta de terminal abra junto com o programa.
* `--add-data`: Inclui suas pastas de HTML/Imagens dentro do executável.

---


### 2. Comando para Gerar o App (`.app`)

Execute o comando abaixo na raiz do projeto. Este comando cria um bundle de aplicativo do macOS que pode ser movido para a pasta /Applications.

```bash
pyinstaller --noconfirm --onefile --windowed \
    --add-data "templates:templates" \
    --add-data "static:static" \
    --icon "static/logo.png" \
    --name "PrecificadoraDaOnda" \
    app.py

```

### Detalhes das Flags:

* **`--windowed`**: Cria um arquivo `.app`. Sem isso, o Mac trataria apenas como um script de terminal.
* **`--add-data "templates:templates"`**: Copia a pasta de HTML para dentro do pacote. Note o uso do `:` (padrão Unix/Mac).
* **`--name`**: Define o nome que aparecerá no Finder.


---

### 3. Entrega Final ao Usuário

Após o processo terminar, uma pasta chamada `dist` será criada. Dentro dela estará o seu arquivo `app.exe`.

1. Pegue esse arquivo `app.exe`.
2. Mova-o para a Área de Trabalho do usuário.
3. Você pode renomeá-lo para "Precificadora Da Onda".
4. O usuário não precisa de mais nada (nem de Python instalado após o build) para rodar o programa.

> **Nota Importante:** Ao rodar pela primeira vez no Windows, o antivírus ou o SmartScreen pode exibir um alerta. Como o executável não possui uma assinatura digital paga, basta o usuário clicar em "Mais informações" e "Executar assim mesmo".

### como Atualizar o App para os Usuários

Como você transformou a aplicação em um executável local (o arquivo .app ou .exe), ela não se atualiza sozinha pela internet. Ela agora é como um programa normal do computador (tipo o Word ou o Photoshop antigo). O GitHub serve apenas para você versionar o código-fonte, não para entregar o executável aos leigos.

1. Você altera o código: Você melhora algo no código e testa no seu Mac.

2. Gera um novo Build: Você roda aquele comando do PyInstaller novamente (pyinstaller --noconfirm --windowed ... app.py). O PyInstaller vai apagar a versão velha na pasta dist e criar o novo PrecificadoraDaOnda.app atualizado.

3. Distribuição: Você compacta esse novo .app (botão direito -> Comprimir) para virar um arquivo .zip e envia para os seus funcionários (por WhatsApp, Slack, Google Drive, etc.).

4. Instalação no usuário: O funcionário apaga o ícone antigo que ele tinha na Área de Trabalho e substitui pelo ícone novo que você enviou. Pronto.