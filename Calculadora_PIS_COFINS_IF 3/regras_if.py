# -*- coding: utf-8 -*-
"""
regras_if.py
============
Regras da apuração de PIS/COFINS para INSTITUIÇÕES FINANCEIRAS (bancos,
financeiras, cooperativas de crédito, corretoras/DTVM, arrendamento mercantil),
com base no plano de contas COSIF.

Características do setor (base legal):
  - Regime CUMULATIVO obrigatório, SEM crédito
    (art. 8º, I, da Lei 10.637/2002 e art. 10, I, da Lei 10.833/2003 — as PJ do
     § 6º do art. 3º da Lei 9.718/98 permanecem no regime cumulativo).
  - PIS 0,65% e COFINS 4,00% (art. 18 da Lei 10.684/2003).
  - Base = Receita Bruta Operacional − Exclusões − Deduções de intermediação
    financeira (art. 3º, § 6º, I, "a" a "e", da Lei 9.718/98).

Plano COSIF (grupos de resultado):
  - Grupo 7 = Contas de Resultado Credoras (RENDAS / receitas)
  - Grupo 8 = Contas de Resultado Devedoras (DESPESAS)
  - Grupos 1,2,3,4,5,6,9 = patrimoniais / compensação (fora da apuração)

Tudo aqui é PARAMETRIZÁVEL: ajuste as listas de palavras-chave e as alíquotas
conforme a realidade da instituição.
"""

from __future__ import annotations
from dataclasses import dataclass
from typing import List
import unicodedata


# --------------------------------------------------------------------------- #
# Classificações (chaves internas) próprias de instituição financeira
# --------------------------------------------------------------------------- #
RENDA_TRIBUTADA = "RENDA_TRIBUTADA"                 # rendas que compõem a base
RENDA_EXCLUIDA = "RENDA_EXCLUIDA"                   # rendas fora da base / não incidência
DEDUCAO_INTERMEDIACAO = "DEDUCAO_INTERMEDIACAO"     # deduções da base (§6º I "a"-"e")
DESPESA_OPERACIONAL = "DESPESA_OPERACIONAL"         # demais despesas (não afetam a base)
FORA_APURACAO = "FORA_APURACAO"                     # patrimonial / compensação

CLASSIFICACOES_VALIDAS = {
    RENDA_TRIBUTADA, RENDA_EXCLUIDA, DEDUCAO_INTERMEDIACAO,
    DESPESA_OPERACIONAL, FORA_APURACAO,
}

ROTULOS = {
    RENDA_TRIBUTADA: "Renda Tributada",
    RENDA_EXCLUIDA: "Renda Excluída / Não Tributada",
    DEDUCAO_INTERMEDIACAO: "Dedução - Intermediação Financeira (§6º I)",
    DESPESA_OPERACIONAL: "Despesa Operacional (não deduz base)",
    FORA_APURACAO: "Fora da Apuração (patrimonial)",
}

# CST padrão por classificação (rendas). Créditos não existem (regime cumulativo).
CST_PADRAO = {
    RENDA_TRIBUTADA: {"pis": "01", "cofins": "01"},
    RENDA_EXCLUIDA: {"pis": "08", "cofins": "08"},
    DEDUCAO_INTERMEDIACAO: {"pis": "", "cofins": ""},
    DESPESA_OPERACIONAL: {"pis": "", "cofins": ""},
    FORA_APURACAO: {"pis": "", "cofins": ""},
}


# --------------------------------------------------------------------------- #
# Alíquotas (fixas para instituição financeira)
# --------------------------------------------------------------------------- #
ALIQUOTA_PIS = 0.0065      # 0,65%
ALIQUOTA_COFINS = 0.04     # 4,00% (majorada - Lei 10.684/2003)


@dataclass
class Config:
    """Parâmetros gerais da apuração de uma instituição financeira."""
    competencia: str = ""
    empresa: str = ""
    cnpj: str = ""
    tipo_instituicao: str = "Banco / crédito / corretora"

    # Prefixos COSIF: grupo 7 = rendas, grupo 8 = despesas.
    prefixos_renda: tuple = ("7",)
    prefixos_despesa: tuple = ("8",)


# --------------------------------------------------------------------------- #
# PALAVRAS-CHAVE (descrições típicas do COSIF)
# Ordem de teste: exclusões > deduções (§6 I) > rendas tributadas >
#                 despesas operacionais. A primeira que casar vence.
# --------------------------------------------------------------------------- #

# Rendas EXCLUÍDAS da base (não incidência / exclusões do art. 3º da Lei 9.718/98):
#  exportação de serviços; reversões de provisões; recuperações de créditos
#  baixados como perda; resultado de participações/equivalência; dividendos/JCP;
#  ganhos de avaliação a valor justo; ajuste a valor presente; alienação de bens
#  do ativo; rendas isentas/alíquota zero.
KEYWORDS_RENDA_EXCLUIDA: List[str] = [
    "exportacao", "prestacao de servicos ao exterior", "servicos ao exterior",
    "reversao de provisao", "reversao de provisoes", "reversao",
    "recuperacao de creditos baixados", "recuperacao de encargos",
    "recuperacao de creditos",
    "resultado de participacoes", "equivalencia patrimonial",
    "lucros e dividendos", "dividendo", "juros sobre capital",
    "ganho de avaliacao a valor justo", "avaliacao a valor justo", "valor justo",
    "ajuste a valor presente",
    "lucro na alienacao", "alienacao de bens", "alienacao de investimentos",
    "venda de imobilizado", "ganho de capital",
    "rendas isentas", "isenta", "aliquota zero",
]

# Deduções da base - art. 3º, § 6º, I, alíneas "a" a "e", da Lei 9.718/98:
#  a) despesas de intermediação financeira (captação);
#  b) obrigações por empréstimos e repasses;
#  c) deságio na colocação de títulos;
#  d) perdas com títulos de renda fixa e variável (exceto ações);
#  e) perdas com ativos financeiros e mercadorias em operações de hedge.
KEYWORDS_DEDUCAO_INTERMEDIACAO: List[str] = [
    # a) captação / intermediação financeira
    "despesas de captacao", "despesa de captacao", "captacao no mercado",
    "operacoes de captacao", "intermediacao financeira", "despesas de intermediacao",
    # b) obrigações por empréstimos e repasses
    "obrigacoes por emprestimos", "obrigacoes por repasses",
    "operacoes de emprestimos e repasses", "emprestimos e repasses",
    # c) deságio na colocação de títulos
    "desagio na colocacao de titulos", "desagio",
    # d) perdas com títulos de renda fixa e variável (exceto ações)
    "perdas com titulos de renda", "perdas com titulos", "prejuizo com titulos",
    # e) perdas com ativos financeiros e mercadorias em operações de hedge
    "operacoes de hedge", "perdas em hedge", "perdas com instrumentos financeiros derivativos",
    "perdas com ativos financeiros",
]

# Rendas TRIBUTADAS (compõem a base): receita operacional típica de banco.
KEYWORDS_RENDA_TRIBUTADA: List[str] = [
    "rendas de operacoes de credito", "operacoes de credito",
    "rendas de arrendamento", "arrendamento mercantil",
    "rendas de cambio", "resultado de operacoes de cambio", "operacoes de cambio",
    "rendas de aplicacoes interfinanceiras", "aplicacoes interfinanceiras",
    "rendas de titulos", "rendas com titulos e valores mobiliarios",
    "titulos e valores mobiliarios", "tvm",
    "rendas de prestacao de servicos", "prestacao de servicos", "tarifas",
    "rendas de garantias prestadas", "rendas de tarifas",
    "rendas de operacoes com", "outras receitas operacionais", "rendas",
]

# Despesas operacionais que NÃO reduzem a base (apenas informativas): pessoal,
# administrativas, tributárias, PCLD, depreciação, etc.
KEYWORDS_DESPESA_OPERACIONAL: List[str] = [
    "despesas de pessoal", "despesas administrativas", "despesas tributarias",
    "provisao para creditos de liquidacao duvidosa", "pcld",
    "provisoes operacionais", "outras despesas operacionais",
    "depreciacao", "amortizacao", "despesas de comercializacao",
    "despesa", "despesas",
]

# Contas fora da apuração (patrimoniais / compensação) — pistas auxiliares.
KEYWORDS_FORA: List[str] = [
    "disponibilidades", "caixa", "reservas", "patrimonio liquido",
    "capital social", "conta de compensacao", "compensacao",
]


def normalizar(texto: str) -> str:
    """Minúsculas, sem acentos e sem espaços duplicados (para casar palavras-chave)."""
    if texto is None:
        return ""
    txt = str(texto).strip().lower()
    txt = unicodedata.normalize("NFKD", txt)
    txt = "".join(c for c in txt if not unicodedata.combining(c))
    return " ".join(txt.split())


@dataclass
class ResultadoClassificacao:
    classificacao: str
    origem: str            # "cadastro", "palavra-chave", "grupo" ou "indefinido"
    regra: str
    confianca: str         # "alta", "media", "baixa"


def classificar_por_heuristica(conta: str, descricao: str, cfg: Config) -> ResultadoClassificacao:
    """
    Classifica uma conta COSIF que não está no cadastro:
      1) palavras-chave (exclusão > dedução §6 I > renda tributada > despesa);
      2) grupo COSIF pelo prefixo (7 = renda tributada, 8 = despesa operacional);
      3) indefinido.
    """
    desc = normalizar(descricao)
    conta_str = str(conta).strip()

    # 1) Palavras-chave
    for kw in KEYWORDS_FORA:
        if kw in desc:
            return ResultadoClassificacao(
                FORA_APURACAO, "palavra-chave",
                f"Descrição contém '{kw}' (conta patrimonial/compensação)", "media")

    for kw in KEYWORDS_RENDA_EXCLUIDA:
        if kw in desc:
            return ResultadoClassificacao(
                RENDA_EXCLUIDA, "palavra-chave",
                f"Descrição contém '{kw}' (exclusão da base)", "media")

    for kw in KEYWORDS_DEDUCAO_INTERMEDIACAO:
        if kw in desc:
            return ResultadoClassificacao(
                DEDUCAO_INTERMEDIACAO, "palavra-chave",
                f"Descrição contém '{kw}' (dedução §6º I - Lei 9.718/98)", "media")

    for kw in KEYWORDS_RENDA_TRIBUTADA:
        if kw in desc:
            return ResultadoClassificacao(
                RENDA_TRIBUTADA, "palavra-chave",
                f"Descrição contém '{kw}'", "media")

    for kw in KEYWORDS_DESPESA_OPERACIONAL:
        if kw in desc:
            return ResultadoClassificacao(
                DESPESA_OPERACIONAL, "palavra-chave",
                f"Descrição contém '{kw}' (não reduz a base)", "media")

    # 2) Grupo COSIF pelo prefixo do código
    if conta_str.startswith(cfg.prefixos_renda):
        return ResultadoClassificacao(
            RENDA_TRIBUTADA, "grupo",
            "Conta do grupo 7 (rendas) — assumida tributada", "baixa")
    if conta_str.startswith(cfg.prefixos_despesa):
        return ResultadoClassificacao(
            DESPESA_OPERACIONAL, "grupo",
            "Conta do grupo 8 (despesas) — assumida NÃO dedutível por segurança", "baixa")

    # 3) Indefinido -> fora da apuração
    return ResultadoClassificacao(
        FORA_APURACAO, "indefinido",
        "Conta patrimonial/compensação ou não reconhecida", "baixa")
