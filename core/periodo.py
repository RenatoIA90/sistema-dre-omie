"""Relatórios de um período (vários meses): o Excel do período tem uma coluna por mês e a coluna TOTAL com a
soma do período (year to date), no layout do REPORT padrão, com SALDOS do período, quadros, indicadores e a
BASE CLASSIFICADA de todos os meses (fórmulas SUMIFS por mês, como no report mensal)."""
from copy import deepcopy
from datetime import date

import openpyxl

from . import dre as D
from . import excel, hagili, ms_festas, psp

MESES_ABREV = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]
MAX_MESES = 24


def meses_entre(a0, m0, a1, m1):
    """[(ano, mes)] de a0/m0 até a1/m1, inclusive."""
    i, fim = a0 * 12 + m0 - 1, a1 * 12 + m1 - 1
    return [(k // 12, k % 12 + 1) for k in range(i, fim + 1)]


def rotulo(ano, mes):
    return f"{MESES_ABREV[mes - 1]}/{ano}"


def saldos_periodo(saldos_por_mes):
    """Saldo inicial do primeiro mês e saldo final do último mês de cada conta."""
    contas = {}
    for saldos in saldos_por_mes:
        for s in saldos or []:
            c = contas.setdefault(s["conta"], {"conta": s["conta"], "saldo_inicial": s["saldo_inicial"]})
            c["saldo_final"] = s["saldo_final"]
    return list(contas.values())


def montar(bases, cfg, modelo=""):
    """bases: [(ano, mes, base do mês)] já classificadas. modelo: 'ms', 'psp', 'hagili' ou '' (padrão).
    Devolve (base de todos os meses com coluna = mês, DRE com uma coluna por mês, cfg usado)."""
    labels, todas = [], []
    for a, m, base in bases:
        lab = rotulo(a, m)
        labels.append(lab)
        todas += [dict(b, coluna=lab) for b in base]
    c = deepcopy(cfg)
    c.update(ordem_colunas=labels, grupos_colunas={}, modo_colunas="periodo")
    if modelo == "ms":
        ms_festas.codificar(todas)
        d = ms_festas.montar_dre(todas, c, colunas=labels)
    elif modelo == "psp":
        psp.codificar(todas)
        d = psp.montar_dre(todas, c, colunas=labels)
    elif modelo == "hagili":
        hagili.codificar(todas)  # mesmos códigos para as linhas novas em todos os meses
        d = hagili.montar_dre(todas, c, colunas=labels)
    else:
        D._codificar_contas(todas, None)
        d = D.montar_dre(todas, c)
        d["colunas"] = labels  # meses sem movimento também aparecem
        for l in d["linhas"]:
            l["valores"] = {lab: l["valores"].get(lab, 0.0) for lab in labels}
    return todas, d, c


def gerar_excel(caminho, empresa, bases, saldos_por_mes, cfg, modelo=""):
    """Excel do período: REPORT com uma coluna por mês + TOTAL DO PERÍODO (soma = year to date)."""
    (a0, m0), (a1, m1) = bases[0][:2], bases[-1][:2]
    todas, d, c = montar(bases, cfg, modelo)
    excel.gerar_excel(caminho, empresa, a1, m1, todas, d, saldos_periodo(saldos_por_mes), c)

    # ajustes de cabeçalho: cada coluna com o seu mês, TOTAL como soma do período
    wb = openpyxl.load_workbook(caminho)
    ws = wb.worksheets[0]
    ws.title = f"REPORT {m0:02d}-{a0} A {m1:02d}-{a1}"
    periodo = f"{m0:02d}/{a0} a {m1:02d}/{a1}"
    col = 3
    for a, m, _ in bases:
        ws.cell(3, col).value = date(a, m, 1)
        col += 3
    ws.cell(2, col).value = "TOTAL DO PERÍODO (YTD)"
    ws.cell(3, col).value = periodo
    ws.cell(3, col).number_format = "General"
    if "SALDOS" in wb.sheetnames:
        pa, pm = (a0 - 1, 12) if m0 == 1 else (a0, m0 - 1)
        import calendar
        wb["SALDOS"].cell(2, 3).value = f"SALDO INICIAL {calendar.monthrange(pa, pm)[1]:02d}/{pm:02d}/{pa}"
    if "INDICADORES" in wb.sheetnames:
        wb["INDICADORES"].cell(2, 3).value = f"Período {periodo}"
        wb["INDICADORES"].cell(2, 3).number_format = "General"
    wb.save(caminho)
    return caminho
