# Apuração de PIS/Cofins a partir do Balancete Contábil

Ferramenta em Python que lê o **balancete contábil** e o **cadastro de contas
Cofins**, classifica automaticamente cada conta, apura a base de cálculo, calcula
**PIS** e **Cofins** (regime **cumulativo** ou **não-cumulativo**) e gera um
**relatório Excel** com **Dashboard**, **Log de Inconsistências** e **Auditoria
Tributária**.

## Arquivos do projeto

| Arquivo | Função |
|---|---|
| `apuracao_pis_cofins.py` | Motor principal + linha de comando |
| `regras.py` | Alíquotas, CSTs e regras de classificação automática (parametrizável) |
| `relatorio_excel.py` | Geração do relatório Excel (dashboard, abas e gráficos) |
| `gerar_exemplos.py` | Cria arquivos de exemplo `balancete.xlsx` e `cadastro_cofins.xlsx` |
| `README.md` | Este arquivo |

## Requisitos

Python 3.9+ com `pandas` e `openpyxl`:

```bash
pip install pandas openpyxl
```

## Como usar

1. Gere os arquivos de exemplo (opcional, para testar):

```bash
python gerar_exemplos.py
```

2. Rode a apuração:

```bash
python apuracao_pis_cofins.py \
    --balancete balancete.xlsx \
    --cadastro cadastro_cofins.xlsx \
    --regime nao_cumulativo \
    --competencia 07/2026 \
    --empresa "Minha Empresa LTDA" \
    --cnpj 00.000.000/0001-00 \
    --saida apuracao_pis_cofins.xlsx
```

Parâmetros:

- `--regime` : `cumulativo` (PIS 0,65% / Cofins 3,00%, **sem crédito**) ou
  `nao_cumulativo` (PIS 1,65% / Cofins 7,60%, **com crédito** sobre despesas
  dedutíveis). Padrão: `nao_cumulativo`.
- `--balancete`, `--cadastro`, `--saida` : caminhos dos arquivos.
- `--competencia`, `--empresa`, `--cnpj` : informativos, aparecem no relatório.

Você também pode usar o motor dentro de outro script:

```python
from apuracao_pis_cofins import executar
from regras import Config
import relatorio_excel

cfg = Config(regime="cumulativo", competencia="07/2026", empresa="ACME")
apur = executar("balancete.xlsx", "cadastro_cofins.xlsx", cfg)
relatorio_excel.gerar(apur, "saida.xlsx")
print(apur.pis_a_pagar, apur.cofins_a_pagar)
```

## Layout esperado dos arquivos de entrada

### `balancete.xlsx` — aba `Balancete`

| Coluna | Obrigatória | Descrição |
|---|:--:|---|
| `Conta` | ✔ | Código da conta contábil (texto). Ex.: `3.1.1.001` |
| `Descricao` | ✔ | Descrição da conta |
| `Saldo_Atual` | ✔ | Saldo final do período (número) — **é o valor usado na apuração** |
| `Saldo_Anterior`, `Debito`, `Credito` | — | Opcionais, apenas informativos |

Convenção de sinal: receitas e despesas com **saldo positivo**. Saldos negativos
são aceitos (usa-se o valor absoluto), mas geram um alerta no log para conferência.

### `cadastro_cofins.xlsx` — aba `Cadastro`

| Coluna | Obrigatória | Descrição |
|---|:--:|---|
| `Conta` | ✔ | Código da conta (mesmo padrão do balancete) |
| `Descricao` | — | Descrição |
| `Natureza` | — | `RECEITA`, `DESPESA` ou `PATRIMONIAL` |
| `Classificacao` | ✔ | Uma das chaves abaixo |
| `CST_PIS`, `CST_COFINS` | — | Código CST; se vazio, usa o padrão da classificação |
| `Gera_Credito` | — | `S`/`N` — só relevante no regime não-cumulativo |
| `Observacao` | — | Texto livre |

Valores válidos para `Classificacao`:

- `RECEITA_TRIBUTADA` — receita sujeita à incidência (entra na base)
- `RECEITA_EXCLUIDA` — exclusões da base (exportação, cancelamentos, descontos
  incondicionais, IPI, ICMS-ST, reversões, isentas, etc.)
- `DESPESA_DEDUTIVEL` — despesa que **gera crédito** no não-cumulativo
- `DESPESA_NAO_DEDUTIVEL` — despesa que **não** gera crédito
- `NAO_CLASSIFICADA` — conta patrimonial / fora da apuração

## Como funciona a classificação automática

Para cada conta do balancete, na ordem:

1. **Cadastro** — se a conta está no cadastro Cofins, usa a classificação de lá
   (confiança **alta**).
2. **Palavra-chave** — se não está no cadastro, tenta reconhecer pela descrição
   (ex.: "exportação" → excluída; "energia elétrica" → dedutível; "salário" →
   não dedutível). Confiança **média**.
3. **Prefixo da conta** — fallback pelo início do código (`3…` = receita,
   `4…` = despesa não dedutível por segurança). Confiança **baixa**.
4. **Indefinido** — se nada casar, fica fora da apuração e vai para o log.

Toda conta classificada fora do cadastro é registrada no **Log de Inconsistências**
para revisão. As regras ficam em `regras.py` e podem ser ajustadas livremente.

## Cálculo

```
Receita Bruta Total   = Receita Tributada + Receita Excluída
Base de Cálculo       = Receita Tributada            (= Receita Bruta − Exclusões)

PIS (débito)          = Base × alíquota PIS
Cofins (débito)       = Base × alíquota Cofins

Crédito PIS           = Despesa Dedutível (que gera crédito) × alíquota PIS      *
Crédito Cofins        = Despesa Dedutível (que gera crédito) × alíquota Cofins   *

PIS a pagar           = máx(PIS débito − Crédito PIS, 0)
Cofins a pagar        = máx(Cofins débito − Crédito Cofins, 0)
```

\* Créditos só se aplicam no regime **não-cumulativo**.

## O relatório Excel

Cinco abas:

1. **Dashboard** — KPIs (receita bruta, base, PIS/Cofins a pagar, total a
   recolher, nº de inconsistências) e gráficos de composição e de tributos.
2. **Apuração** — memória de cálculo com **fórmulas vivas** (`SUMIFS`): se você
   alterar um saldo ou uma classificação na planilha, os totais recalculam.
3. **Balancete Classificado** — cada conta com sua classificação, CST, origem da
   decisão e nível de confiança.
4. **Log de Inconsistências** — pontos que exigem revisão, ordenados por
   severidade, com ação sugerida.
5. **Auditoria Tributária** — trilha por conta: como foi classificada, com que
   confiança, valor, CSTs e contribuição para cada tributo.

## Aviso

As alíquotas e regras de classificação são um **ponto de partida** e refletem o
tratamento geral. Casos específicos (regimes monofásicos, alíquota zero,
suspensões, créditos presumidos, atividades imobiliárias, etc.) devem ser
revisados por profissional habilitado e ajustados em `regras.py`. Esta ferramenta
não substitui a análise contábil/fiscal.
