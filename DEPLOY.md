# Deploy — PH-Sheets

## Visão geral

```
Push para main
     │
     ▼
GitHub Actions (.github/workflows/ci.yml)
     ├─ FAIL → bloqueia, deploy não acontece
     └─ PASS
          │
          ▼
     Railway detecta o push
          └─ nixpacks.toml → Python 3.12 + requirements-server.txt
               └─ gunicorn app:app
```

O Railway só recebe o push depois que o código chegou em `main`. Com branch protection ativada, `main` só aceita código que passou nos testes — a sequência CI → deploy é garantida por estrutura, não por configuração extra.

---

## Arquivos de deploy

| Arquivo | Função |
|---------|--------|
| `.github/workflows/ci.yml` | Roda `pytest` a cada push/PR em `main` |
| `nixpacks.toml` | Configura o build no Railway: Python 3.12, instala `requirements-server.txt`, inicia com gunicorn |
| `requirements-server.txt` | Dependências do servidor — sem PyInstaller, Pillow, macholib (só desktop) |
| `Procfile` | Fallback de start caso o Railway ignore o nixpacks.toml |

---

## Por que Python 3.12 no servidor

O projeto usa Python 3.14 localmente (macOS). O Railway usa Python 3.13 por padrão, mas `pandas==2.2.1` e `numpy==1.26.4` não têm wheels pré-compiladas para 3.13 — o pip tenta compilar do zero e falha. Python 3.12 tem wheels prontas para essas versões, o build é instantâneo.

O `nixpacks.toml` fixa isso:

```toml
[phases.setup]
nixPkgs = ["python312"]

[phases.install]
cmds = ["pip install -r requirements-server.txt"]

[start]
cmd = "gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120"
```

---

## GitHub Actions — CI

**Arquivo:** `.github/workflows/ci.yml`

Dispara em: push ou PR para `main`.

Passos:
1. Checkout do código
2. Setup Python 3.12 (com cache de pip via `requirements-server.txt`)
3. `pip install -r requirements-server.txt pytest pytest-cov`
4. `python -m pytest tests/ -v --tb=short`

Se qualquer teste falhar, o job falha e o merge/push é bloqueado (se branch protection estiver ativa).

---

## Branch protection (configurar uma vez no GitHub)

Acesse: **GitHub → Settings → Branches → Add rule → `main`**

Marcar:
- [x] Require status checks to pass before merging
  - Status check: `test`
- [x] Require branches to be up to date before merging

Com isso, ninguém consegue fazer merge para `main` sem passar nos testes — nem via PR nem via push direto.

---

## Variáveis de ambiente

Nenhuma variável obrigatória para o servidor funcionar. A aplicação não usa secrets, banco de dados nem API keys fixas.

A única API externa (SEFAZ AL) é chamada com a chave da NFe enviada pelo usuário em cada requisição — sem credenciais armazenadas no servidor.

Se precisar adicionar variáveis no futuro: **Railway → projeto → Variables**.

---

## Limitações do ambiente de produção

**Arquivos temporários:** os Excel gerados ficam em `/tmp` do container Railway. Funcionam para download imediato após o processamento. Se a instância reiniciar (deploy novo, crash, idle), os arquivos somem — isso é esperado dado o fluxo de uso (gera → baixa na hora).

**Upload:** limite de 10 MB por requisição (`MAX_CONTENT_LENGTH` em `app.py`). XMLs de NFe raramente passam de 1 MB, então é seguro.

**Workers:** 2 workers gunicorn + timeout de 120s. O processamento paralelo de até 3 lotes usa `ThreadPoolExecutor` dentro do worker — adequado para o volume esperado.

---

## Fluxo de trabalho diário

```bash
# Desenvolvimento normal
git checkout -b feature/minha-mudanca
# ... edita código ...
python -m pytest tests/ -v          # roda local antes de abrir PR
git push origin feature/minha-mudanca
# Abre PR → CI roda automaticamente → merge → Railway deploya
```

```bash
# Hotfix urgente direto na main (requer desabilitar branch protection temporariamente)
git commit -m "fix: ..."
git push origin main
# CI roda → Railway deploya em ~2 min
```

---

## Repositório

`git@github.com:FelipeLFirmino/PH-sheets.git`
