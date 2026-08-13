# Calculadora de PIS/COFINS — Instituição Financeira (COSIF)

Apuração de PIS e COFINS para **instituições financeiras** (bancos, financeiras,
cooperativas de crédito, corretoras/DTVM, arrendamento mercantil), a partir do
balancete no **plano COSIF**. Diferente da versão comercial, esta calculadora é
construída para a estrutura de resultado do setor (grupo 7 = rendas, grupo 8 =
despesas) e para a apuração cumulativa própria das IF.

## Regras aplicadas (base legal)

- **Regime cumulativo obrigatório, sem crédito** — art. 8º, I, da Lei 10.637/2002
  e art. 10, I, da Lei 10.833/2003 (as PJ do § 6º do art. 3º da Lei 9.718/98
  permanecem no cumulativo).
- **PIS 0,65% e COFINS 4,00%** — art. 18 da Lei 10.684/2003.
- **Base = Rendas Tributadas − Deduções de intermediação financeira** — art. 3º,
  § 6º, I, alíneas "a" a "e", da Lei 9.718/98:
  - a) despesas de intermediação financeira (captação);
  - b) obrigações por empréstimos e repasses;
  - c) deságio na colocação de títulos;
  - d) perdas com títulos de renda fixa e variável, exceto ações;
  - e) perdas com ativos financeiros e mercadorias em operações de hedge.

## Arquivos

| Arquivo | Função |
|---|---|
| `apuracao_if.py` | Motor: lê, classifica, calcula e comanda o relatório |
| `regras_if.py` | Regras: alíquotas, classificação COSIF e palavras-chave |
| `relatorio_if.py` | Gera o relatório Excel (5 abas + gráficos) |
| `exemplos_if.py` | Cria balancete e cadastro de exemplo (banco COSIF) |
| `APURAR_IF.bat` | Atalho de dois cliques no Windows |

## Como rodar

Dois cliques em `APURAR_IF.bat` (Windows), ou pela linha de comando:

```bash
python apuracao_if.py --balancete balancete.xlsx --cadastro cadastro.xlsx \
    --competencia 07/2026 --empresa "Banco Exemplo S.A." --saida apuracao.xlsx
```

## Como o cálculo funciona

```
Receita Bruta Operacional = Rendas Tributadas + Rendas Excluídas   (grupo 7)
Base de Cálculo           = Rendas Tributadas − Deduções (§6º I)
PIS    = Base × 0,65%
COFINS = Base × 4,00%
Total  = PIS + COFINS      (sem crédito)
```

As **despesas operacionais** (pessoal, administrativas, tributárias, PCLD, etc.)
aparecem apenas como memória — **não** reduzem a base. Em especial, a **PCLD não
é dedutível** da base de PIS/COFINS (não consta no rol taxativo do § 6º, I).

## O relatório (5 abas)

Dashboard (KPIs e gráficos), Apuração (memória de cálculo com fórmulas vivas e a
Base Legal no rodapé), Balancete Classificado, Log de Inconsistências e
Auditoria Tributária.

## Layout de entrada

**balancete.xlsx** (aba `Balancete`): `Conta`, `Descricao`, `Saldo_Atual`
(obrigatórias). **cadastro.xlsx** (aba `Cadastro`): `Conta`, `Descricao`,
`Classificacao` (obrigatórias) e, opcionalmente, `Natureza`, `CST_PIS`,
`CST_COFINS`, `Observacao`. Classificações válidas: `RENDA_TRIBUTADA`,
`RENDA_EXCLUIDA`, `DEDUCAO_INTERMEDIACAO`, `DESPESA_OPERACIONAL`, `FORA_APURACAO`.

## Aviso

As deduções "a"–"e" aplicam-se a bancos e instituições de crédito. Seguradoras
(§6º II), previdência (§6º III) e capitalização (§6º IV) têm deduções próprias.
A lista de contas dedutíveis deve ser validada pela área fiscal e mapeada no
cadastro. Esta ferramenta não substitui a análise profissional.
