# -*- coding: utf-8 -*-
"""
regras.py
=========
Parametrização da apuração de PIS/Cofins:

- Alíquotas por regime (cumulativo e não-cumulativo)
- Códigos CST padrão por classificação tributária
- Regras de classificação automática (por prefixo de conta e por palavras-chave)

Tudo aqui é PARAMETRIZÁVEL. Ajuste os dicionários/listas conforme a realidade
contábil e fiscal da sua empresa. As regras de palavra-chave são um ponto de
partida conservador; contas não reconhecidas com segurança são sempre enviadas
para o Log de Inconsistências para revisão manual.
"""

from __future__ import annotations
from dataclasses import dataclass, field
from typing import Dict, List
import unicodedata


# --------------------------------------------------------------------------- #
# Classificações tributárias possíveis (chaves internas)
# --------------------------------------------------------------------------- #
RECEITA_TRIBUTADA = "RECEITA_TRIBUTADA"
RECEITA_EXCLUIDA = "RECEITA_EXCLUIDA"
DESPESA_DEDUTIVEL = "DESPESA_DEDUTIVEL"       # gera crédito (só no não-cumulativo)
DESPESA_NAO_DEDUTIVEL = "DESPESA_NAO_DEDUTIVEL"  # não gera crédito
NAO_CLASSIFICADA = "NAO_CLASSIFICADA"         # não entra na apuração (patrimonial, etc.)

CLASSIFICACOES_VALIDAS = {
    RECEITA_TRIBUTADA,
    RECEITA_EXCLUIDA,
    DESPESA_DEDUTIVEL,
    DESPESA_NAO_DEDUTIVEL,
    NAO_CLASSIFICADA,
}

# Rótulos amigáveis para exibição em relatórios
ROTULOS = {
    RECEITA_TRIBUTADA: "Receita Tributada",
    RECEITA_EXCLUIDA: "Receita Excluída",
    DESPESA_DEDUTIVEL: "Despesa Dedutível (gera crédito)",
    DESPESA_NAO_DEDUTIVEL: "Despesa Não Dedutível",
    NAO_CLASSIFICADA: "Não Classificada / Fora da Apuração",
}


# --------------------------------------------------------------------------- #
# CST padrão por classificação (usado quando o cadastro não informa o CST)
#   Receitas (saídas):   01 tributada, 06 alíquota zero, 07 isenta, 08 sem incidência
#   Créditos (entradas): 50 crédito vinculado a receita tributada, 70 sem crédito
# --------------------------------------------------------------------------- #
CST_PADRAO = {
    RECEITA_TRIBUTADA: {"pis": "01", "cofins": "01"},
    RECEITA_EXCLUIDA: {"pis": "06", "cofins": "06"},
    DESPESA_DEDUTIVEL: {"pis": "50", "cofins": "50"},
    DESPESA_NAO_DEDUTIVEL: {"pis": "70", "cofins": "70"},
    NAO_CLASSIFICADA: {"pis": "", "cofins": ""},
}


# --------------------------------------------------------------------------- #
# Alíquotas por regime
# --------------------------------------------------------------------------- #
@dataclass(frozen=True)
class Aliquotas:
    pis: float
    cofins: float
    permite_credito: bool


REGIMES: Dict[str, Aliquotas] = {
    # Lucro Presumido (em geral): sem direito a crédito
    "cumulativo": Aliquotas(pis=0.0065, cofins=0.03, permite_credito=False),
    # Lucro Real (em geral): com direito a crédito sobre insumos/despesas elegíveis
    "nao_cumulativo": Aliquotas(pis=0.0165, cofins=0.076, permite_credito=True),
}


# --------------------------------------------------------------------------- #
# Regras de classificação automática
# --------------------------------------------------------------------------- #
@dataclass
class Config:
    """Parâmetros gerais da apuração."""
    regime: str = "nao_cumulativo"          # "cumulativo" ou "nao_cumulativo"
    competencia: str = ""                    # ex.: "07/2026" (informativo)
    empresa: str = ""                        # razão social (informativo)
    cnpj: str = ""                           # informativo

    # Prefixos de conta (plano de contas). Ajuste ao seu plano.
    # Padrão comum: 3 = Receitas, 4 = Custos/Despesas.
    prefixos_receita: tuple = ("3",)
    prefixos_despesa: tuple = ("4",)

    def aliquotas(self) -> Aliquotas:
        if self.regime not in REGIMES:
            raise ValueError(
                f"Regime inválido: {self.regime!r}. "
                f"Use um de: {list(REGIMES.keys())}"
            )
        return REGIMES[self.regime]


# Palavras-chave -> classificação. Ordem importa: a primeira regra que casar vence.
# As chaves são fragmentos SEM acento e em minúsculas (o texto é normalizado antes).
#
# Receita Excluída: valores que reduzem a base (não tributados) ou fora do campo
# de incidência das contribuições.
KEYWORDS_RECEITA_EXCLUIDA: List[str] = [
    "venda cancelad", "vendas cancelad", "cancelamento",
    "devolucao", "devolvid",
    "desconto incondicional", "abatimento",
    "ipi", "icms st", "icms substituicao", "icms-st", "substituicao tributaria",
    "exportacao", "receita de exportacao", "mercado externo",
    "isenta", "isencao", "aliquota zero", "aliquota-zero",
    "monofasic",
    "reversao de provisao", "reversao de provisoes", "reversao",
    "equivalencia patrimonial",
    "dividendo", "juros sobre capital",
    "subvencao", "subvencoes",
    "venda de imobilizado", "venda de ativo", "alienacao de bens",
    "recuperacao de despesa", "recuperacao de tributo",
]

# Receita Tributada: receita operacional sujeita à incidência.
KEYWORDS_RECEITA_TRIBUTADA: List[str] = [
    "receita de venda", "receita bruta", "venda de mercadoria",
    "venda de produto", "receita de servico", "prestacao de servico",
    "receita operacional", "faturamento", "receita financeira",
    "receita de locacao", "receita de aluguel", "receita",
]

# Despesa Dedutível (gera crédito no não-cumulativo): insumos e despesas elegíveis.
KEYWORDS_DESPESA_DEDUTIVEL: List[str] = [
    "insumo", "materia-prima", "materia prima", "materia-prima",
    "mercadoria para revenda", "compra de mercadoria", "aquisicao de mercadoria",
    "energia eletrica",
    "aluguel", "arrendamento", "leasing", "arrendamento mercantil",
    "frete", "transporte de carga", "armazenagem",
    "depreciacao", "amortizacao",
    "combustivel", "manutencao de maquina", "manutencao de equipamento",
    "servico aplicado na producao", "embalagem",
]

# Despesa Não Dedutível: mão de obra, tributos, financeiras, etc.
KEYWORDS_DESPESA_NAO_DEDUTIVEL: List[str] = [
    "salario", "ordenado", "folha de pagamento", "pro-labore", "pro labore",
    "13o salario", "ferias", "encargo social", "fgts", "inss",
    "vale transporte", "vale refeicao", "vale alimentacao",
    "honorario", "comissao",
    "multa", "juros de mora", "despesa financeira", "iof",
    "brinde", "doacao", "confraternizacao", "presente",
    "imposto", "taxa", "contribuicao sindical",
    "propaganda", "publicidade", "marketing",
    "viagem", "hospedagem", "refeicao",
    "material de escritorio", "material de expediente",
    "telefone", "internet", "software",
]

# Contas patrimoniais / de resultado que NÃO entram na apuração de PIS/Cofins
# (ativos, passivos, patrimônio líquido, apuração de IRPJ/CSLL, etc.)
KEYWORDS_NAO_CLASSIFICADA: List[str] = [
    "irpj", "csll", "provisao para imposto de renda",
    "resultado do exercicio", "lucros acumulados", "reserva de",
    "capital social", "conta corrente", "caixa e equivalente",
]


def normalizar(texto: str) -> str:
    """Minúsculas, sem acentos e sem espaços duplicados — para casar palavras-chave."""
    if texto is None:
        return ""
    txt = str(texto).strip().lower()
    txt = unicodedata.normalize("NFKD", txt)
    txt = "".join(c for c in txt if not unicodedata.combining(c))
    return " ".join(txt.split())


@dataclass
class ResultadoClassificacao:
    classificacao: str
    origem: str            # "cadastro", "palavra-chave", "prefixo" ou "indefinido"
    regra: str             # descrição da regra aplicada (trilha de auditoria)
    confianca: str         # "alta", "media" ou "baixa"


def classificar_por_heuristica(conta: str, descricao: str, cfg: Config) -> ResultadoClassificacao:
    """
    Classifica uma conta que NÃO está no cadastro Cofins, usando:
      1) palavras-chave na descrição (mais específico);
      2) prefixo do código da conta (fallback);
      3) indefinido (vai para o log com confiança baixa).
    """
    desc = normalizar(descricao)
    conta_str = str(conta).strip()

    # 1) Palavras-chave (prioridade: excluída > dedutível > não dedutível > tributada)
    for kw in KEYWORDS_NAO_CLASSIFICADA:
        if kw in desc:
            return ResultadoClassificacao(
                NAO_CLASSIFICADA, "palavra-chave",
                f"Descrição contém '{kw}' (conta patrimonial/fora da apuração)", "media")

    for kw in KEYWORDS_RECEITA_EXCLUIDA:
        if kw in desc:
            return ResultadoClassificacao(
                RECEITA_EXCLUIDA, "palavra-chave",
                f"Descrição contém '{kw}'", "media")

    for kw in KEYWORDS_DESPESA_DEDUTIVEL:
        if kw in desc:
            return ResultadoClassificacao(
                DESPESA_DEDUTIVEL, "palavra-chave",
                f"Descrição contém '{kw}'", "media")

    for kw in KEYWORDS_DESPESA_NAO_DEDUTIVEL:
        if kw in desc:
            return ResultadoClassificacao(
                DESPESA_NAO_DEDUTIVEL, "palavra-chave",
                f"Descrição contém '{kw}'", "media")

    for kw in KEYWORDS_RECEITA_TRIBUTADA:
        if kw in desc:
            return ResultadoClassificacao(
                RECEITA_TRIBUTADA, "palavra-chave",
                f"Descrição contém '{kw}'", "media")

    # 2) Prefixo da conta (fallback grosseiro)
    if conta_str.startswith(cfg.prefixos_receita):
        return ResultadoClassificacao(
            RECEITA_TRIBUTADA, "prefixo",
            f"Conta inicia com prefixo de receita {cfg.prefixos_receita}", "baixa")
    if conta_str.startswith(cfg.prefixos_despesa):
        return ResultadoClassificacao(
            DESPESA_NAO_DEDUTIVEL, "prefixo",
            f"Conta inicia com prefixo de despesa {cfg.prefixos_despesa} "
            f"(assumida NÃO dedutível por segurança)", "baixa")

    # 3) Indefinido
    return ResultadoClassificacao(
        NAO_CLASSIFICADA, "indefinido",
        "Não foi possível classificar por cadastro, palavra-chave ou prefixo", "baixa")
