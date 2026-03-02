
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

### 3. Automatizando com um Script de Build

Para facilitar sua vida ao configurar várias máquinas, você pode deixar um arquivo chamado `build.bat` na raiz do seu projeto no GitHub com este conteúdo:

```batch
@echo off
echo Iniciando criacao do executavel...
python -m venv venv
call venv\Scripts\activate
pip install -r requirements.txt
pip install pyinstaller
pyinstaller --noconfirm --onefile --windowed --add-data "templates;templates" --add-data "static;static" app.py
echo Concluido! O arquivo esta na pasta dist.
pause

```

---

### 4. Entrega Final ao Usuário

Após o processo terminar, uma pasta chamada `dist` será criada. Dentro dela estará o seu arquivo `app.exe`.

1. Pegue esse arquivo `app.exe`.
2. Mova-o para a Área de Trabalho do usuário.
3. Você pode renomeá-lo para "Precificadora Da Onda".
4. O usuário não precisa de mais nada (nem de Python instalado após o build) para rodar o programa.

> **Nota Importante:** Ao rodar pela primeira vez no Windows, o antivírus ou o SmartScreen pode exibir um alerta. Como o executável não possui uma assinatura digital paga, basta o usuário clicar em "Mais informações" e "Executar assim mesmo".

Gostaria que eu montasse o script de build específico para o seu Mac, caso queira gerar a versão `.app` para outros usuários de Apple?