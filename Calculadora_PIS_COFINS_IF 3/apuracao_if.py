# -*- coding: utf-8 -*-
"""
apuracao_if.py
==============
Motor de apuração de PIS/COFINS para INSTITUIÇÕES FINANCEIRAS (plano COSIF).

Fluxo:
  1. Ler balancete COSIF (.xlsx)               — grupos 7 (rendas) e 8 (despesas)
  2. Ler cadastro de contas (.xlsx)            — classificação PIS/COFINS de cada conta
  3. Classificar cada conta                    — cadastro -> heurística -> log
  4. Somar: Rendas Tributadas, Rendas Excluídas, Deduções §6 I, Despesas
  5. Apurar a base:  Base = Rendas Tributadas − Deduções de intermediação
  6. Calcular PIS (0,65%) e COFINS (4%)        — regime cumulativo, sem crédito
  7. Gerar relatório Excel

Uso:
    python apuracao_if.py --balancete balancete.xlsx --cadastro cadastro.xlsx
        --competencia 07/2026 --empresa "Banco Exemplo S.A." --saida apuracao.xlsx
"""

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass, field
from typing import List, Optional

import openpyxl

import regras_if
from regras_if import (
    Config, classificar_por_heuristica,
    RENDA_TRIBUTADA, RENDA_EXCLUIDA, DEDUCAO_INTERMEDIACAO,
    DESPESA_OPERACIONAL, FORA_APURACAO, CLASSIFICACOES_VALIDAS,
    CST_PADRAO, ROTULOS, ALIQUOTA_PIS, ALIQUOTA_COFINS,
)


# --------------------------------------------------------------------------- #
# Estruturas de dados
# --------------------------------------------------------------------------- #
@dataclass
class ContaApurada:
    conta: str
    descricao: str
    saldo: float
    classificacao: str
    natureza: str
    cst_pis: str
    cst_cofins: str
    origem: str
    regra: str
    confianca: str


@dataclass
class Inconsistencia:
    conta: str
    descricao: str
    saldo: float
    tipo: str
    detalhe: str
    severidade: str
    acao_sugerida: str


@dataclass
class Apuracao:
    cfg: Config
    contas: List[ContaApurada] = field(default_factory=list)
    inconsistencias: List[Inconsistencia] = field(default_factory=list)

    # totais por categoria
    renda_tributada: float = 0.0
    renda_excluida: float = 0.0
    deducao_intermediacao: float = 0.0
    despesa_operacional: float = 0.0
    receita_bruta_operacional: float = 0.0

    # base e tributos
    base_calculo: float = 0.0
    aliquota_pis: float = ALIQUOTA_PIS
    aliquota_cofins: float = ALIQUOTA_COFINS
    pis_a_pagar: float = 0.0
    cofins_a_pagar: float = 0.0


# --------------------------------------------------------------------------- #
# Leitura dos arquivos
# --------------------------------------------------------------------------- #
COLS_BALANCETE_OBRIG = ["Conta", "Descricao", "Saldo_Atual"]
COLS_CADASTRO_OBRIG = ["Conta", "Descricao", "Classificacao"]


def _num(valor) -> Optional[float]:
    """Converte célula em float (aceita 1.234,56 ou 1,234.56). None se não for número."""
    if valor is None or isinstance(valor, bool):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    s = str(valor).strip().replace("R$", "").replace(" ", "")
    if s == "":
        return None
    negativo = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        n = float(s)
        return -n if negativo else n
    except ValueError:
        return None


def _ler_planilha(caminho: str, nome: str):
    try:
        wb = openpyxl.load_workbook(caminho, data_only=True, read_only=True)
    except FileNotFoundError:
        raise SystemExit(f"[ERRO] Arquivo de {nome} não encontrado: {caminho}")
    except Exception as e:
        raise SystemExit(f"[ERRO] Falha ao ler {nome} ({caminho}): {e}")
    ws = wb.active
    linhas = list(ws.iter_rows(values_only=True))
    wb.close()
    if not linhas:
        raise SystemExit(f"[ERRO] {nome} está vazio: {caminho}")
    cabecalho = [str(c).strip() if c is not None else "" for c in linhas[0]]
    registros = []
    for r in linhas[1:]:
        if r is None or all(c is None for c in r):
            continue
        d = {}
        for i, col in enumerate(cabecalho):
            if col:
                d[col] = r[i] if i < len(r) else None
        registros.append(d)
    return cabecalho, registros


def ler_balancete(caminho: str) -> List[dict]:
    cabecatro, registros = _ler_planilha(caminho, "balancete COSIF")

    mapa_colunas = {
        "Conta COSIF": "Conta",
        "Descrição da Conta": "Descricao",
        "Saldo Final": "Saldo_Atual",
    }

    cabecalho = [mapa_colunas.get(col, col) for col in cabecalho]

    for registro in registros:
        for origem, destino in mapa_colunas.items():
            if origem in registro:
                registro[destino] = registro[origem]

    faltando = [c for c in COLS_BALANCETE_OBRIG if c not in cabecalho]

    if faltando:
        raise SystemExit(
            f"[ERRO] Balancete sem as colunas obrigatórias: {faltando}. "
            f"Colunas encontradas: {cabecalho}"
        )

    linhas = []

    for d in registros:
        conta = str(d.get("Conta") or "").strip()

        if conta == "":
            continue

        linhas.append({
            "Conta": conta,
            "Descricao": str(d.get("Descricao") or "").strip(),
            "Saldo_Atual": _num(d.get("Saldo_Atual")),
        })

def ler_cadastro(caminho: str) -> List[dict]:
    cabecalho, registros = _ler_planilha(caminho, "cadastro")
    faltando = [c for c in COLS_CADASTRO_OBRIG if c not in cabecalho]
    if faltando:
        raise SystemExit(
            f"[ERRO] Cadastro sem as colunas obrigatórias: {faltando}. "
            f"Colunas encontradas: {cabecalho}")
    linhas = []
    for d in registros:
        conta = str(d.get("Conta") or "").strip()
        if conta == "":
            continue
        linhas.append({
            "Conta": conta,
            "Descricao": str(d.get("Descricao") or "").strip(),
            "Natureza": str(d.get("Natureza") or "").strip().upper(),
            "Classificacao": str(d.get("Classificacao") or "").strip().upper(),
            "CST_PIS": str(d.get("CST_PIS") or "").strip(),
            "CST_COFINS": str(d.get("CST_COFINS") or "").strip(),
            "Observacao": str(d.get("Observacao") or "").strip(),
        })
    return linhas


# --------------------------------------------------------------------------- #
# Classificação
# --------------------------------------------------------------------------- #
def classificar(balancete: List[dict], cadastro: List[dict],
                cfg: Config, apur: Apuracao) -> None:
    cad_idx = {row["Conta"]: row for row in cadastro}
    contas_vistas = set()

    for lin in balancete:
        conta, desc, saldo = lin["Conta"], lin["Descricao"], lin["Saldo_Atual"]

        if saldo is None:
            apur.inconsistencias.append(Inconsistencia(
                conta, desc, 0.0, "Saldo inválido",
                "Coluna Saldo_Atual vazia ou não numérica.",
                "Alta", "Corrigir o valor do saldo no balancete."))
            continue

        if conta in contas_vistas:
            apur.inconsistencias.append(Inconsistencia(
                conta, desc, float(saldo), "Conta duplicada",
                "Conta aparece mais de uma vez; os saldos serão somados.",
                "Média", "Verificar duplicidade no balancete."))
        contas_vistas.add(conta)

        if conta in cad_idx:
            reg = cad_idx[conta]
            classif = reg["Classificacao"]
            if classif not in CLASSIFICACOES_VALIDAS:
                res = classificar_por_heuristica(conta, desc, cfg)
                apur.inconsistencias.append(Inconsistencia(
                    conta, desc, float(saldo), "Classificação inválida no cadastro",
                    f"Valor '{classif}' inválido. Aplicada heurística: {ROTULOS[res.classificacao]}.",
                    "Alta", "Corrigir a coluna Classificacao no cadastro."))
                classif, origem, regra, conf = res.classificacao, "palavra-chave", res.regra, "baixa"
                cst_pis = CST_PADRAO[classif]["pis"]
                cst_cofins = CST_PADRAO[classif]["cofins"]
            else:
                origem = "cadastro"
                regra = "Classificação obtida diretamente do cadastro."
                conf = "alta"
                cst_pis = str(reg.get("CST_PIS", "") or CST_PADRAO[classif]["pis"]).strip()
                cst_cofins = str(reg.get("CST_COFINS", "") or CST_PADRAO[classif]["cofins"]).strip()
            natureza = str(reg.get("Natureza", "") or "").strip().upper()
        else:
            res = classificar_por_heuristica(conta, desc, cfg)
            classif = res.classificacao
            origem, regra, conf = res.origem, res.regra, res.confianca
            cst_pis = CST_PADRAO[classif]["pis"]
            cst_cofins = CST_PADRAO[classif]["cofins"]
            natureza = _natureza(classif)
            sev = "Alta" if conf == "baixa" else "Média"
            apur.inconsistencias.append(Inconsistencia(
                conta, desc, float(saldo), "Conta ausente no cadastro",
                f"Classificação automática: {ROTULOS[classif]} "
                f"(origem: {origem}, confiança: {conf}). {regra}",
                sev, "Incluir a conta no cadastro com a classificação correta."))

        # Sinal atípico em rendas (grupo 7 costuma ser credor/positivo)
        if saldo < 0 and classif in (RENDA_TRIBUTADA, RENDA_EXCLUIDA):
            apur.inconsistencias.append(Inconsistencia(
                conta, desc, float(saldo), "Sinal atípico",
                "Renda com saldo negativo — verificar convenção de sinal.",
                "Média", "Conferir o sinal do saldo no balancete."))

        apur.contas.append(ContaApurada(
            conta=conta, descricao=desc, saldo=float(saldo),
            classificacao=classif, natureza=natureza,
            cst_pis=cst_pis, cst_cofins=cst_cofins,
            origem=origem, regra=regra, confianca=conf))

    contas_balancete = set(l["Conta"] for l in balancete)
    for conta, reg in cad_idx.items():
        if conta not in contas_balancete:
            apur.inconsistencias.append(Inconsistencia(
                conta, str(reg.get("Descricao", "")), 0.0,
                "Conta do cadastro sem saldo",
                "Conta existe no cadastro mas não aparece no balancete.",
                "Baixa", "Apenas informativo."))


def _natureza(classif: str) -> str:
    if classif in (RENDA_TRIBUTADA, RENDA_EXCLUIDA):
        return "RENDA"
    if classif in (DEDUCAO_INTERMEDIACAO, DESPESA_OPERACIONAL):
        return "DESPESA"
    return "PATRIMONIAL"


# --------------------------------------------------------------------------- #
# Cálculo
# --------------------------------------------------------------------------- #
def calcular(apur: Apuracao) -> None:
    for c in apur.contas:
        valor = abs(c.saldo)
        if c.classificacao == RENDA_TRIBUTADA:
            apur.renda_tributada += valor
        elif c.classificacao == RENDA_EXCLUIDA:
            apur.renda_excluida += valor
        elif c.classificacao == DEDUCAO_INTERMEDIACAO:
            apur.deducao_intermediacao += valor
        elif c.classificacao == DESPESA_OPERACIONAL:
            apur.despesa_operacional += valor
        # FORA_APURACAO não entra

    # Receita Bruta Operacional = todas as rendas do grupo 7 (tributadas + excluídas)
    apur.receita_bruta_operacional = apur.renda_tributada + apur.renda_excluida

    # Base = Rendas Tributadas − Deduções de intermediação (§6º I). Nunca negativa.
    apur.base_calculo = max(apur.renda_tributada - apur.deducao_intermediacao, 0.0)

    # Regime cumulativo: sem crédito.
    apur.pis_a_pagar = round(apur.base_calculo * ALIQUOTA_PIS, 2)
    apur.cofins_a_pagar = round(apur.base_calculo * ALIQUOTA_COFINS, 2)

    if apur.renda_tributada == 0:
        apur.inconsistencias.append(Inconsistencia(
            "-", "Apuração", 0.0, "Base zerada",
            "Nenhuma renda tributada identificada — base igual a zero.",
            "Alta", "Revisar a classificação das rendas (grupo 7)."))


# --------------------------------------------------------------------------- #
# Orquestração e linha de comando
# --------------------------------------------------------------------------- #
def executar(balancete_path: str, cadastro_path: str, cfg: Config) -> Apuracao:
    balancete = ler_balancete(balancete_path)
    cadastro = ler_cadastro(cadastro_path)
    apur = Apuracao(cfg=cfg)
    classificar(balancete, cadastro, cfg, apur)
    calcular(apur)
    return apur


def _resumo(apur: Apuracao) -> None:
    a = apur
    print("=" * 62)
    print("  APURAÇÃO PIS/COFINS — INSTITUIÇÃO FINANCEIRA (cumulativo)")
    if a.cfg.competencia:
        print(f"  Competência: {a.cfg.competencia}   Empresa: {a.cfg.empresa}")
    print("=" * 62)
    print(f"  Rendas Tributadas ............ R$ {a.renda_tributada:,.2f}")
    print(f"  (+) Rendas Excluídas ......... R$ {a.renda_excluida:,.2f}")
    print(f"  (=) Receita Bruta Operacional  R$ {a.receita_bruta_operacional:,.2f}")
    print(f"  (-) Deduções Intermediação ... R$ {a.deducao_intermediacao:,.2f}")
    print(f"  (=) Base de Cálculo .......... R$ {a.base_calculo:,.2f}")
    print(f"  Despesas Operacionais (info) . R$ {a.despesa_operacional:,.2f}")
    print("-" * 62)
    print(f"  PIS    (0,65%) ............... R$ {a.pis_a_pagar:,.2f}")
    print(f"  COFINS (4,00%) .............. R$ {a.cofins_a_pagar:,.2f}")
    print("-" * 62)
    print(f"  TOTAL A RECOLHER ............. R$ {a.pis_a_pagar + a.cofins_a_pagar:,.2f}")
    print(f"  Inconsistências .............. {len(a.inconsistencias)}")
    print("=" * 62)


def main(argv: Optional[List[str]] = None) -> int:
    p = argparse.ArgumentParser(
        description="Apuração de PIS/COFINS para instituições financeiras (COSIF).")
    p.add_argument("--balancete", default="balancete.xlsx")
    p.add_argument("--cadastro", default="cadastro.xlsx")
    p.add_argument("--competencia", default="")
    p.add_argument("--empresa", default="")
    p.add_argument("--cnpj", default="")
    p.add_argument("--saida", default="apuracao_if.xlsx")
    args = p.parse_args(argv)

    cfg = Config(competencia=args.competencia, empresa=args.empresa, cnpj=args.cnpj)
    apur = executar(args.balancete, args.cadastro, cfg)
    _resumo(apur)

    import relatorio_if
    relatorio_if.gerar(apur, args.saida)
    print(f"\nRelatório gerado: {args.saida}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
