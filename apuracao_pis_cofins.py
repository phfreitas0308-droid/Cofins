# -*- coding: utf-8 -*-
"""
apuracao_pis_cofins.py
======================
Motor de apuração de PIS e Cofins a partir do balancete contábil e do cadastro
de contas Cofins.

Fluxo:
  1. Ler balancete contábil (.xlsx)
  2. Ler cadastro das contas Cofins (.xlsx)
  3. Classificar automaticamente cada conta (cadastro -> heurística -> log)
  4. Calcular Receita Tributada, Receita Excluída, Despesa Dedutível e Não Dedutível
  5. Apurar a base de cálculo
  6. Calcular PIS e Cofins (regime cumulativo ou não-cumulativo)
  7. Gerar relatório Excel (Dashboard, Balancete Classificado, Apuração,
     Log de Inconsistências, Auditoria Tributária)

Uso (linha de comando):
    python apuracao_pis_cofins.py \
        --balancete balancete.xlsx \
        --cadastro cadastro_cofins.xlsx \
        --regime nao_cumulativo \
        --competencia 07/2026 \
        --empresa "Minha Empresa LTDA" \
        --cnpj 00.000.000/0001-00 \
        --saida apuracao_pis_cofins.xlsx
"""

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass, field
from typing import Dict, List, Optional

import openpyxl

import regras
from regras import (
    Config, ResultadoClassificacao, classificar_por_heuristica, normalizar,
    RECEITA_TRIBUTADA, RECEITA_EXCLUIDA, DESPESA_DEDUTIVEL,
    DESPESA_NAO_DEDUTIVEL, NAO_CLASSIFICADA, CLASSIFICACOES_VALIDAS,
    CST_PADRAO, ROTULOS,
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
    gera_credito: bool
    origem: str            # "cadastro" | "palavra-chave" | "prefixo" | "indefinido"
    regra: str             # trilha de auditoria
    confianca: str         # "alta" | "media" | "baixa"


@dataclass
class Inconsistencia:
    conta: str
    descricao: str
    saldo: float
    tipo: str              # categoria do problema
    detalhe: str
    severidade: str        # "Alta" | "Média" | "Baixa"
    acao_sugerida: str


@dataclass
class Apuracao:
    cfg: Config
    contas: List[ContaApurada] = field(default_factory=list)
    inconsistencias: List[Inconsistencia] = field(default_factory=list)

    # totais por categoria
    receita_tributada: float = 0.0
    receita_excluida: float = 0.0
    despesa_dedutivel: float = 0.0
    despesa_nao_dedutivel: float = 0.0
    receita_bruta_total: float = 0.0

    # base e tributos
    base_calculo: float = 0.0
    aliquota_pis: float = 0.0
    aliquota_cofins: float = 0.0
    pis_debito: float = 0.0
    cofins_debito: float = 0.0
    credito_pis: float = 0.0
    credito_cofins: float = 0.0
    pis_a_pagar: float = 0.0
    cofins_a_pagar: float = 0.0


# --------------------------------------------------------------------------- #
# Leitura dos arquivos
# --------------------------------------------------------------------------- #
COLS_BALANCETE_OBRIG = ["Conta", "Descricao", "Saldo_Atual"]
COLS_CADASTRO_OBRIG = ["Conta", "Descricao", "Classificacao"]


def _num(valor) -> Optional[float]:
    """Converte o valor de uma célula em float. Retorna None se não for número.
    Aceita número puro ou texto no formato brasileiro (1.234,56) ou americano."""
    if valor is None:
        return None
    if isinstance(valor, bool):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    s = str(valor).strip().replace("R$", "").replace(" ", "")
    if s == "":
        return None
    negativo = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    if "," in s and "." in s:          # 1.234,56 -> 1234.56
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:                       # 1234,56 -> 1234.56
        s = s.replace(",", ".")
    try:
        n = float(s)
        return -n if negativo else n
    except ValueError:
        return None


def _ler_planilha(caminho: str, nome: str):
    """Lê a primeira aba do .xlsx com openpyxl e devolve (cabeçalho, lista de dicts)."""
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
    cabecalho, registros = _ler_planilha(caminho, "balancete")
    faltando = [c for c in COLS_BALANCETE_OBRIG if c not in cabecalho]
    if faltando:
        raise SystemExit(
            f"[ERRO] Balancete sem as colunas obrigatórias: {faltando}. "
            f"Colunas encontradas: {cabecalho}")
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
    return linhas


def ler_cadastro(caminho: str) -> List[dict]:
    cabecalho, registros = _ler_planilha(caminho, "cadastro Cofins")
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
            "Gera_Credito": str(d.get("Gera_Credito") or "N").strip(),
            "Observacao": str(d.get("Observacao") or "").strip(),
        })
    return linhas


# --------------------------------------------------------------------------- #
# Classificação
# --------------------------------------------------------------------------- #
def _to_bool_credito(valor) -> bool:
    return str(valor).strip().upper() in ("S", "SIM", "TRUE", "1", "Y", "YES")


def classificar(balancete: List[dict], cadastro: List[dict],
                cfg: Config, apur: Apuracao) -> None:
    """Percorre o balancete, aplica cadastro ou heurística e popula apur.contas."""
    cad_idx = {row["Conta"]: row for row in cadastro}
    contas_vistas = set()

    for lin in balancete:
        conta = lin["Conta"]
        desc = lin["Descricao"]
        saldo = lin["Saldo_Atual"]

        # Saldo inválido -> inconsistência e pula
        if saldo is None:
            apur.inconsistencias.append(Inconsistencia(
                conta, desc, 0.0, "Saldo inválido",
                "Coluna Saldo_Atual vazia ou não numérica.",
                "Alta", "Corrigir o valor do saldo no balancete."))
            continue

        if conta in contas_vistas:
            apur.inconsistencias.append(Inconsistencia(
                conta, desc, float(saldo), "Conta duplicada",
                "Conta aparece mais de uma vez no balancete; os saldos serão somados.",
                "Média", "Verificar duplicidade de lançamento no balancete."))
        contas_vistas.add(conta)

        if conta in cad_idx:
            reg = cad_idx[conta]
            classif = reg["Classificacao"]
            if classif not in CLASSIFICACOES_VALIDAS:
                # cadastro com classificação inválida -> heurística + log
                res = classificar_por_heuristica(conta, desc, cfg)
                apur.inconsistencias.append(Inconsistencia(
                    conta, desc, float(saldo), "Classificação inválida no cadastro",
                    f"Valor '{classif}' não é uma classificação válida. "
                    f"Aplicada heurística: {ROTULOS[res.classificacao]}.",
                    "Alta", "Corrigir a coluna Classificacao no cadastro."))
                classif = res.classificacao
                origem, regra, conf = "palavra-chave", res.regra, "baixa"
                cst_pis = CST_PADRAO[classif]["pis"]
                cst_cofins = CST_PADRAO[classif]["cofins"]
                gera_cred = classif == DESPESA_DEDUTIVEL
            else:
                origem = "cadastro"
                regra = "Classificação obtida diretamente do cadastro Cofins."
                conf = "alta"
                cst_pis = str(reg.get("CST_PIS", "") or CST_PADRAO[classif]["pis"]).strip()
                cst_cofins = str(reg.get("CST_COFINS", "") or CST_PADRAO[classif]["cofins"]).strip()
                gera_cred = _to_bool_credito(reg.get("Gera_Credito", "N"))
            natureza = str(reg.get("Natureza", "") or "").strip().upper()
        else:
            # não está no cadastro -> heurística + log
            res = classificar_por_heuristica(conta, desc, cfg)
            classif = res.classificacao
            origem, regra, conf = res.origem, res.regra, res.confianca
            cst_pis = CST_PADRAO[classif]["pis"]
            cst_cofins = CST_PADRAO[classif]["cofins"]
            gera_cred = classif == DESPESA_DEDUTIVEL
            natureza = _natureza_por_classificacao(classif)

            sev = "Alta" if conf == "baixa" else "Média"
            apur.inconsistencias.append(Inconsistencia(
                conta, desc, float(saldo), "Conta ausente no cadastro",
                f"Não cadastrada. Classificação automática: {ROTULOS[classif]} "
                f"(origem: {origem}, confiança: {conf}). Regra: {regra}",
                sev, "Incluir a conta no cadastro Cofins com a classificação correta."))

        # Coerência de sinal: despesa com saldo negativo ou receita negativa
        if saldo < 0 and classif in (RECEITA_TRIBUTADA, RECEITA_EXCLUIDA):
            apur.inconsistencias.append(Inconsistencia(
                conta, desc, float(saldo), "Sinal atípico",
                "Conta de receita com saldo negativo — verificar convenção de sinal.",
                "Média", "Conferir o sinal do saldo no balancete."))
        if saldo < 0 and classif in (DESPESA_DEDUTIVEL, DESPESA_NAO_DEDUTIVEL):
            apur.inconsistencias.append(Inconsistencia(
                conta, desc, float(saldo), "Sinal atípico",
                "Conta de despesa com saldo negativo — verificar convenção de sinal.",
                "Média", "Conferir o sinal do saldo no balancete."))

        apur.contas.append(ContaApurada(
            conta=conta, descricao=desc, saldo=float(saldo),
            classificacao=classif, natureza=natureza,
            cst_pis=cst_pis, cst_cofins=cst_cofins, gera_credito=bool(gera_cred),
            origem=origem, regra=regra, confianca=conf))

    # Contas no cadastro mas ausentes no balancete (informativo)
    contas_balancete = set(l["Conta"] for l in balancete)
    for conta, reg in cad_idx.items():
        if conta not in contas_balancete:
            apur.inconsistencias.append(Inconsistencia(
                conta, str(reg.get("Descricao", "")), 0.0,
                "Conta do cadastro sem saldo",
                "Conta existe no cadastro Cofins mas não aparece no balancete.",
                "Baixa", "Apenas informativo — nenhuma ação obrigatória."))


def _natureza_por_classificacao(classif: str) -> str:
    if classif in (RECEITA_TRIBUTADA, RECEITA_EXCLUIDA):
        return "RECEITA"
    if classif in (DESPESA_DEDUTIVEL, DESPESA_NAO_DEDUTIVEL):
        return "DESPESA"
    return "PATRIMONIAL"


# --------------------------------------------------------------------------- #
# Cálculo de categorias, base e tributos
# --------------------------------------------------------------------------- #
def calcular(apur: Apuracao) -> None:
    cfg = apur.cfg
    aliq = cfg.aliquotas()

    for c in apur.contas:
        valor = abs(c.saldo)  # usa valor absoluto por classificação
        if c.classificacao == RECEITA_TRIBUTADA:
            apur.receita_tributada += valor
        elif c.classificacao == RECEITA_EXCLUIDA:
            apur.receita_excluida += valor
        elif c.classificacao == DESPESA_DEDUTIVEL:
            apur.despesa_dedutivel += valor
        elif c.classificacao == DESPESA_NAO_DEDUTIVEL:
            apur.despesa_nao_dedutivel += valor
        # NAO_CLASSIFICADA não entra

    apur.receita_bruta_total = apur.receita_tributada + apur.receita_excluida

    # Base de cálculo = receita tributada (receita bruta - exclusões)
    apur.base_calculo = apur.receita_tributada

    apur.aliquota_pis = aliq.pis
    apur.aliquota_cofins = aliq.cofins

    # Débitos (contribuição sobre a receita)
    apur.pis_debito = round(apur.base_calculo * aliq.pis, 2)
    apur.cofins_debito = round(apur.base_calculo * aliq.cofins, 2)

    # Créditos (somente no regime não-cumulativo, sobre despesas dedutíveis)
    if aliq.permite_credito:
        base_credito = sum(abs(c.saldo) for c in apur.contas
                           if c.classificacao == DESPESA_DEDUTIVEL and c.gera_credito)
        apur.credito_pis = round(base_credito * aliq.pis, 2)
        apur.credito_cofins = round(base_credito * aliq.cofins, 2)
    else:
        apur.credito_pis = 0.0
        apur.credito_cofins = 0.0

    apur.pis_a_pagar = round(max(apur.pis_debito - apur.credito_pis, 0.0), 2)
    apur.cofins_a_pagar = round(max(apur.cofins_debito - apur.credito_cofins, 0.0), 2)

    # Sanidade: nenhuma receita tributada => alerta
    if apur.receita_tributada == 0:
        apur.inconsistencias.append(Inconsistencia(
            "-", "Apuração", 0.0, "Base zerada",
            "Nenhuma receita tributada identificada — base de cálculo igual a zero.",
            "Alta", "Revisar a classificação das contas de receita."))


# --------------------------------------------------------------------------- #
# Orquestração
# --------------------------------------------------------------------------- #
def executar(balancete_path: str, cadastro_path: str, cfg: Config) -> Apuracao:
    balancete = ler_balancete(balancete_path)
    cadastro = ler_cadastro(cadastro_path)
    apur = Apuracao(cfg=cfg)
    classificar(balancete, cadastro, cfg, apur)
    calcular(apur)
    return apur


def _resumo_console(apur: Apuracao) -> None:
    a = apur
    print("=" * 60)
    print(f"  APURAÇÃO PIS/COFINS — regime: {a.cfg.regime}")
    if a.cfg.competencia:
        print(f"  Competência: {a.cfg.competencia}   Empresa: {a.cfg.empresa}")
    print("=" * 60)
    print(f"  Receita Bruta Total .......... R$ {a.receita_bruta_total:,.2f}")
    print(f"  (-) Receita Excluída ......... R$ {a.receita_excluida:,.2f}")
    print(f"  (=) Receita Tributada / Base . R$ {a.base_calculo:,.2f}")
    print(f"  Despesa Dedutível ............ R$ {a.despesa_dedutivel:,.2f}")
    print(f"  Despesa Não Dedutível ........ R$ {a.despesa_nao_dedutivel:,.2f}")
    print("-" * 60)
    print(f"  PIS   ({a.aliquota_pis*100:.2f}%) débito .... R$ {a.pis_debito:,.2f}"
          f"   crédito R$ {a.credito_pis:,.2f}   a pagar R$ {a.pis_a_pagar:,.2f}")
    print(f"  Cofins({a.aliquota_cofins*100:.2f}%) débito .. R$ {a.cofins_debito:,.2f}"
          f"   crédito R$ {a.credito_cofins:,.2f}   a pagar R$ {a.cofins_a_pagar:,.2f}")
    print("-" * 60)
    print(f"  TOTAL A RECOLHER ............. R$ {a.pis_a_pagar + a.cofins_a_pagar:,.2f}")
    print(f"  Inconsistências registradas .. {len(a.inconsistencias)}")
    print("=" * 60)


def main(argv: Optional[List[str]] = None) -> int:
    p = argparse.ArgumentParser(
        description="Apuração de PIS/Cofins a partir do balancete contábil.")
    p.add_argument("--balancete", default="balancete.xlsx", help="Arquivo .xlsx do balancete")
    p.add_argument("--cadastro", default="cadastro_cofins.xlsx", help="Arquivo .xlsx do cadastro Cofins")
    p.add_argument("--regime", default="nao_cumulativo",
                   choices=["cumulativo", "nao_cumulativo"], help="Regime de apuração")
    p.add_argument("--competencia", default="", help="Competência, ex.: 07/2026")
    p.add_argument("--empresa", default="", help="Razão social (informativo)")
    p.add_argument("--cnpj", default="", help="CNPJ (informativo)")
    p.add_argument("--saida", default="apuracao_pis_cofins.xlsx", help="Arquivo .xlsx de saída")
    args = p.parse_args(argv)

    cfg = Config(regime=args.regime, competencia=args.competencia,
                 empresa=args.empresa, cnpj=args.cnpj)
    apur = executar(args.balancete, args.cadastro, cfg)
    _resumo_console(apur)

    # geração do relatório (import tardio para permitir uso do motor sem openpyxl de report)
    import relatorio_excel
    relatorio_excel.gerar(apur, args.saida)
    print(f"\nRelatório gerado: {args.saida}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
