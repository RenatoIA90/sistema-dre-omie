"""Fluxo de caixa do mês (previsto x realizado) a partir da mesma base classificada da DRE.

Estrutura: SALDO INICIAL → seções da DRE (grupos e contas) → transferências entre contas →
VARIAÇÃO DO CAIXA → SALDO FINAL. Cada linha tem o valor previsto e o realizado.

Previsto de cada conta: o valor digitado pelo usuário para o mês ou, se não houver, a sugestão
realizado + lançamentos ainda em aberto no OMIE no mês (previstos, atrasados, a vencer).

O Excel segue a formatação do modelo de DRE da Savian (FECHAMENTO HAGILI): fonte Aptos Narrow,
cabeçalhos azul-marinho, grupos em cinza, seções com borda média e valores puxados por SUMIFS da base.
"""
import calendar
from collections import defaultdict
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

from . import dre as D
from .plano import IGNORAR, SECOES, norm

TRANSF = ("TRANSFERENCIAS", "(+/-) TRANSFERÊNCIAS ENTRE CONTAS")
REALIZADO, EM_ABERTO = "Realizado", "Em aberto"


def chave(grupo, conta):
    return f"{grupo}|{conta}"


def base_pendente(lancs, cfg, mapa):
    """Lançamentos do mês ainda não realizados (previstos/atrasados/a vencer), classificados como na DRE."""
    pend = [l for l in lancs if not D.realizado(l)]
    return D.montar_base(pend, dict(cfg, apenas_realizado=False, rateios_mes=[]), mapa) if pend else []


def codigos(base, pendentes):
    """{(grupo, conta): 'código-conta'}: os mesmos códigos da DRE para as contas realizadas; contas só com
    lançamentos em aberto recebem os próximos números do grupo."""
    cod = {(b["grupo"], b["conta"]): b["conta_fechamento"] for b in base}
    usados = defaultdict(int)
    for (g, _c), cf in cod.items():
        n = cf[4:6]
        if g != IGNORAR and n.isdigit():
            usados[g] = max(usados[g], int(n))
    for g, c in sorted({(b["grupo"], b["conta"]) for b in pendentes} - set(cod), key=lambda x: (x[0], norm(x[1]))):
        usados[g] += 1
        cod[(g, c)] = c if g == IGNORAR else f"{g[:4]}{usados[g]:02d}-{c}"
    return cod


def montar_fluxo(base, pendentes, saldos, cfg, previstos=None):
    """Linhas do fluxo: {nivel, tipo, descricao, chave, codigo, realizado, previsto, sugestao, digitado}.

    previstos: {chave(grupo, conta): valor} digitados pelo usuário para o mês.
    """
    previstos = previstos or {}
    real, pend = defaultdict(float), defaultdict(float)
    for b in base:
        real[(b["grupo"], b["conta"])] += b["valor"]
    for b in pendentes:
        pend[(b["grupo"], b["conta"])] += b["valor"]
    contas = set(real) | set(pend)
    cod = codigos(base, pendentes)

    ini = sum(s.get("saldo_inicial") or 0 for s in saldos or [])
    linhas = [{"nivel": 1, "tipo": "saldo", "descricao": "SALDO INICIAL", "chave": "", "realizado": ini,
               "previsto": ini, "sugestao": ini}]
    grupos = list(D.grupos_da_empresa(cfg))
    secoes = list(SECOES) + [TRANSF]
    tot_real = tot_prev = 0.0
    for sec, rotulo in secoes:
        lin_sec = {"nivel": 1, "tipo": "secao", "descricao": rotulo, "chave": "", "realizado": 0.0, "previsto": 0.0}
        filhos = []
        for cod_g, nome, s, _rel in grupos:
            if (sec == "TRANSFERENCIAS") != (cod_g == IGNORAR) or (sec != "TRANSFERENCIAS" and s != sec):
                continue
            cs = sorted((c for g, c in contas if g == cod_g), key=lambda c: norm(cod[(cod_g, c)]))
            if not cs:
                continue
            lin_g = {"nivel": 2, "tipo": "grupo", "descricao": nome if cod_g == IGNORAR else f"{cod_g}-{nome}",
                     "chave": "", "realizado": 0.0, "previsto": 0.0}
            filhos.append(lin_g)
            for c in cs:
                k = chave(cod_g, c)
                r = round(real.get((cod_g, c), 0.0), 2)
                sug = round(r + pend.get((cod_g, c), 0.0), 2)
                p = round(float(previstos[k]), 2) if k in previstos else sug
                filhos.append({"nivel": 3, "tipo": "conta", "descricao": c, "codigo": cod[(cod_g, c)], "chave": k,
                               "realizado": r, "previsto": p, "sugestao": sug, "digitado": k in previstos})
                lin_g["realizado"] += r
                lin_g["previsto"] += p
            lin_sec["realizado"] += lin_g["realizado"]
            lin_sec["previsto"] += lin_g["previsto"]
        if filhos:
            linhas.append(lin_sec)
            linhas += filhos
            tot_real += lin_sec["realizado"]
            tot_prev += lin_sec["previsto"]
    linhas.append({"nivel": 1, "tipo": "resultado", "descricao": "(=) VARIAÇÃO DO CAIXA", "chave": "",
                   "realizado": tot_real, "previsto": tot_prev})
    linhas.append({"nivel": 1, "tipo": "saldo", "descricao": "SALDO FINAL", "chave": "",
                   "realizado": ini + tot_real, "previsto": ini + tot_prev})
    return linhas


def resumo(linhas):
    """Totais para os indicadores da tela: saldo inicial, entradas, saídas, saldo final (realizado e previsto)."""
    contas = [l for l in linhas if l["tipo"] == "conta" and not l["chave"].startswith(IGNORAR + "|")]
    return {"saldo_inicial": linhas[0]["realizado"],
            "entradas": sum(l["realizado"] for l in contas if l["realizado"] > 0),
            "saidas": sum(l["realizado"] for l in contas if l["realizado"] < 0),
            "entradas_previstas": sum(l["previsto"] for l in contas if l["previsto"] > 0),
            "saidas_previstas": sum(l["previsto"] for l in contas if l["previsto"] < 0),
            "saldo_final": linhas[-1]["realizado"],
            "saldo_final_previsto": linhas[-1]["previsto"]}


def diario(base, saldos, pendentes=()):
    """Movimento por dia: realizado (entradas, saídas, transferências, saldo) e em aberto no OMIE
    (entradas e saídas previstas, saldo projetado = realizado + em aberto acumulados)."""
    campos = ("entradas", "saidas", "transferencias", "entradas_prev", "saidas_prev")
    dias = defaultdict(lambda: dict.fromkeys(campos, 0.0))
    for b in base:
        d = dias[b["data"]]
        if b["grupo"] == IGNORAR:
            d["transferencias"] += b["valor"]
        elif b["valor"] >= 0:
            d["entradas"] += b["valor"]
        else:
            d["saidas"] += b["valor"]
    for b in pendentes:
        if b["grupo"] != IGNORAR and b["data"]:
            dias[b["data"]]["entradas_prev" if b["valor"] >= 0 else "saidas_prev"] += b["valor"]
    saldo = proj = sum(s.get("saldo_inicial") or 0 for s in saldos or [])
    out = []
    for data in sorted(dias):
        d = dias[data]
        saldo += d["entradas"] + d["saidas"] + d["transferencias"]
        proj += d["entradas"] + d["saidas"] + d["transferencias"] + d["entradas_prev"] + d["saidas_prev"]
        out.append({"data": data, **{k: round(v, 2) for k, v in d.items()}, "saldo": round(saldo, 2),
                    "saldo_projetado": round(proj, 2)})
    return out


# ---------------------------------------------------------------- Excel (formatação do modelo de DRE)
FONTE = "Aptos Narrow"
MARINHO = "1F3864"
CINZA_GRUPO = "E8E8E8"
AZUL_DIGITADO = "0000FF"
FMT = '#,##0;\\(#,##0\\);\\-'
FMT_CENT = '#,##0.00;\\(#,##0.00\\);\\-'
FMT_AV = '0%;\\-0%;\\-'
FMT_MES = '[$-416]mmm\\-yy;@'
_fino = Side(style="thin")
_medio = Side(style="medium")
GRADE = Border(top=_fino, bottom=_fino, left=_fino, right=_fino)
SECAO = Border(top=_medio, bottom=_medio)


def _fonte(bold=False, size=11, color=None, italic=False):
    return Font(name=FONTE, size=size, bold=bold, italic=italic, color=color)


def _cab(ws, row, col, texto, fill=MARINHO, cor="FFFFFF", borda=GRADE):
    c = ws.cell(row, col, texto)
    c.font = _fonte(bold=True, color=cor)
    if fill:
        c.fill = PatternFill("solid", fgColor=fill)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if borda:
        c.border = borda
    return c


def gerar_excel_fluxo(caminho, empresa, ano, mes, linhas, dias, base, pendentes, saldos, com_previsto=True):
    wb = Workbook()
    ws = wb.active
    ws.title = f"FLUXO DE CAIXA {mes:02d} {ano}"
    cod = codigos(base, pendentes)
    n_base = _aba_base(wb, base, pendentes, cod)
    tot_saldos = _aba_saldos(wb, saldos, ano, mes) if saldos else None
    _aba_fluxo(ws, empresa, ano, mes, linhas, n_base, tot_saldos, com_previsto)
    _aba_diario(wb, dias, tot_saldos, linhas[0]["realizado"])
    wb.move_sheet("BASE FLUXO", offset=len(wb.sheetnames))
    wb.save(caminho)
    return caminho


def _aba_fluxo(ws, empresa, ano, mes, linhas, n_base, tot_saldos, com_previsto):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 54
    ws.row_dimensions[2].height = 45
    ws.freeze_panes = "C4"
    ws.cell(1, 2, empresa["nome"]).font = _fonte(bold=True, size=12, color=MARINHO)
    _cab(ws, 2, 2, "Fluxo de Caixa - Previsto x Realizado" if com_previsto else "Fluxo de Caixa - Realizado",
         fill=None, cor=None, borda=None)
    _cab(ws, 3, 2, "Descrição da Conta", fill=None, cor=None, borda=None)

    # blocos de colunas: (título, rótulo da 2ª coluna)
    blocos = [("PREVISTO", "AV"), ("REALIZADO", "AV"), ("DIFERENÇA", "% REAL.")]
    if not com_previsto:
        blocos = [blocos[1]]
    col = {}
    for i, (titulo, sub) in enumerate(blocos):
        cp = 3 + 3 * i
        col[titulo.split()[0]] = cp
        ws.merge_cells(start_row=2, start_column=cp, end_row=2, end_column=cp + 1)
        _cab(ws, 2, cp, titulo)
        _cab(ws, 2, cp + 1, None)
        m = _cab(ws, 3, cp, date(ano, mes, 1), fill=None, cor=None, borda=None)
        m.number_format = FMT_MES
        _cab(ws, 3, cp + 1, sub, fill=None, cor=None, borda=None)
        ws.column_dimensions[L(cp)].width = 13 if i < 2 else 14
        ws.column_dimensions[L(cp + 1)].width = 6.57 if i < 2 else 8
        ws.column_dimensions[L(cp + 2)].width = 1.14
    ult = 3 + 3 * len(blocos) - 2  # última coluna com valores
    col_conf = ult + 2
    _cab(ws, 2, col_conf, "CONFERÊNCIA (REALIZADO − BASE)", fill=None, cor=None, borda=None).font = \
        _fonte(bold=True, size=9)
    ws.column_dimensions[L(col_conf)].width = 14

    rng_v = f"'BASE FLUXO'!$J$2:$J${n_base}"
    rng_c = f"'BASE FLUXO'!$G$2:$G${n_base}"
    rng_s = f"'BASE FLUXO'!$B$2:$B${n_base}"
    R, P = L(col["REALIZADO"]), (L(col["PREVISTO"]) if com_previsto else None)

    r0 = 4
    rows = {i: r0 + i for i in range(len(linhas))}
    filhos, secoes, g_atual, s_atual = defaultdict(list), [], None, None
    for i, l in enumerate(linhas):
        if l["tipo"] == "secao":
            s_atual = rows[i]
            secoes.append(s_atual)
        elif l["tipo"] == "grupo":
            g_atual = rows[i]
            filhos[s_atual].append(g_atual)
        elif l["tipo"] == "conta":
            filhos[g_atual].append(rows[i])
    ent = next((rows[i] for i, l in enumerate(linhas) if l["descricao"] == SECOES[0][1]), None)
    var_row, ini_row, fim_row = rows[len(linhas) - 2], rows[0], rows[len(linhas) - 1]

    for i, l in enumerate(linhas):
        r, t = rows[i], l["tipo"]
        ws.cell(r, 1, l["nivel"]).font = _fonte(size=8, color="FFFFFF")
        bold = t != "conta"
        cor = "FFFFFF" if t == "saldo" else None
        fill = PatternFill("solid", fgColor=MARINHO if t == "saldo" else CINZA_GRUPO) if t in ("saldo", "grupo") \
            else None
        c = ws.cell(r, 2, (l.get("codigo") or l["descricao"]) if t == "conta" else l["descricao"])
        c.font = _fonte(bold=bold, color=cor)
        if t == "conta":
            c.alignment = Alignment(indent=1)

        def formula(X, previsto=False):
            if t == "conta":
                if previsto:
                    if l.get("digitado"):
                        return round(l["previsto"], 2)
                    return f"={R}{r}+SUMIFS({rng_v},{rng_c},$B{r},{rng_s},\"{EM_ABERTO}\")"
                return f"=SUMIFS({rng_v},{rng_c},$B{r},{rng_s},\"{REALIZADO}\")"
            if t in ("grupo", "secao"):
                fs = filhos.get(r, [])
                if t == "grupo" and fs:
                    return f"=SUM({X}{fs[0]}:{X}{fs[-1]})"
                return "=" + "+".join(f"{X}{x}" for x in fs) if fs else 0
            if t == "resultado":
                return "=" + "+".join(f"{X}{x}" for x in secoes) if secoes else 0
            if r == ini_row:
                return f"=SALDOS!C{tot_saldos}" if tot_saldos else round(l["realizado"], 2)
            return f"={X}{ini_row}+{X}{var_row}"

        for nome, cp in col.items():
            X = L(cp)
            if nome == "DIFERENÇA":
                v, av = f"={R}{r}-{P}{r}", (f"=IFERROR({R}{r}/{P}{r},0)" if t != "saldo" else None)
            else:
                v = formula(X, previsto=nome == "PREVISTO")
                av = f"=IFERROR({X}{r}/{X}${ent},0)" if ent and t != "saldo" else None
            cv = ws.cell(r, cp, v)
            cv.number_format = FMT
            azul = nome == "PREVISTO" and t == "conta" and l.get("digitado")
            cv.font = _fonte(bold=bold, color=AZUL_DIGITADO if azul else cor)
            ca = ws.cell(r, cp + 1, av)
            ca.number_format = FMT_AV
            ca.font = _fonte(bold=bold, size=10, color=cor)
        alvo = range(2, ult + 2)
        for cc in alvo:
            cel = ws.cell(r, cc)
            if fill:
                cel.fill = fill
            if t in ("secao", "resultado") and cc not in [cp + 2 for cp in col.values()]:
                cel.border = SECAO
        if t == "conta":
            cc = ws.cell(r, col_conf, f"=ROUND({R}{r}-SUMIFS({rng_v},{rng_c},$B{r},{rng_s},\"{REALIZADO}\"),2)")
            cc.number_format = FMT_CENT
            cc.font = _fonte(size=8, color="7F7F7F")

    fim = r0 + len(linhas)
    ws.cell(fim + 1, 2, "Conferência: variação realizada − soma da BASE FLUXO realizada (deve ser 0)").font = \
        _fonte(italic=True, size=10)
    ws.cell(fim + 1, col["REALIZADO"], f"=ROUND({R}{var_row}-SUMIFS({rng_v},{rng_s},\"{REALIZADO}\"),2)")
    if tot_saldos:
        ws.cell(fim + 2, 2, "Conferência: saldo final realizado − saldo final da aba SALDOS (deve ser 0)").font = \
            _fonte(italic=True, size=10)
        ws.cell(fim + 2, col["REALIZADO"], f"=ROUND({R}{fim_row}-SALDOS!D{tot_saldos},2)")
    for rr in (fim + 1, fim + 2):
        x = ws.cell(rr, col["REALIZADO"])
        x.number_format = FMT_CENT
        x.font = _fonte(size=10)
    ws.cell(fim + 4, 2, "Notas:").font = _fonte(bold=True, size=10)
    notas = [
        "• Realizado: SUMIFS da aba BASE FLUXO (lançamentos do extrato OMIE no mês, mesma classificação da DRE).",
        "• Previsto: realizado + lançamentos em aberto no OMIE no mês (previstos, atrasados e a vencer).",
        "• Valores em azul na coluna PREVISTO foram digitados no sistema e substituem a sugestão do OMIE.",
        "• Saídas aparecem entre parênteses. Transferências entre contas da empresa não alteram o saldo total.",
    ]
    if com_previsto:
        notas.append("• DIFERENÇA = realizado − previsto; % REAL. = quanto do previsto já foi realizado.")
    for i, n in enumerate(notas):
        ws.cell(fim + 5 + i, 2, n).font = _fonte(size=10)


def _aba_base(wb, base, pendentes, cod):
    ws = wb.create_sheet("BASE FLUXO")
    hdr = ["Data", "Status", "Cliente ou Fornecedor", "Conta Corrente", "Categoria (original)", "Grupo",
           "Conta do Fluxo", "Projeto", "Situação", "Valor (R$)", "Tipo", "Observações"]
    larg = [11, 11, 38, 18, 34, 9, 40, 22, 13, 14, 11, 60]
    for i, (h, w) in enumerate(zip(hdr, larg), start=1):
        _cab(ws, 1, i, h)
        ws.column_dimensions[L(i)].width = w
    linhas = [(b, REALIZADO) for b in base] + [(b, EM_ABERTO) for b in pendentes]
    for r, (b, status) in enumerate(linhas, start=2):
        vals = [b["data"], status, b["favorecido"], b["conta_corrente"], b["categoria"], b["grupo"],
                cod[(b["grupo"], b["conta"])], b["projeto"], b["situacao"], b["valor"], b["tipo"],
                (b["observacoes"] or "")[:500]]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(r, c, v)
            cell.font = _fonte(color=AZUL_DIGITADO if status == EM_ABERTO and c == 2 else None)
            if c == 1:
                cell.number_format = "DD/MM/YYYY"
            if c == 10:
                cell.number_format = FMT_CENT
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{L(len(hdr))}{max(len(linhas), 1) + 1}"
    return max(len(linhas), 1) + 1


def _aba_saldos(wb, saldos, ano, mes):
    """Saldos por conta no formato do modelo; devolve a linha do TOTAL."""
    ws = wb.create_sheet("SALDOS")
    ws.sheet_view.showGridLines = False
    ant_a, ant_m = (ano - 1, 12) if mes == 1 else (ano, mes - 1)
    ini = f"{calendar.monthrange(ant_a, ant_m)[1]:02d}/{ant_m:02d}/{ant_a}"
    fim = f"{calendar.monthrange(ano, mes)[1]:02d}/{mes:02d}/{ano}"
    for i, h in enumerate(["CONTA", f"SALDO INICIAL {ini}", f"SALDO FINAL {fim}", "VARIAÇÃO"], start=2):
        _cab(ws, 2, i, h, borda=None)
    for r, s in enumerate(saldos, start=3):
        ws.cell(r, 2, s["conta"]).font = _fonte()
        for c, v in ((3, s.get("saldo_inicial") or 0), (4, s.get("saldo_final") or 0), (5, f"=D{r}-C{r}")):
            x = ws.cell(r, c, v)
            x.number_format = FMT_CENT
            x.font = _fonte(color=AZUL_DIGITADO if c < 5 else None)
    tot = 3 + len(saldos)
    ws.cell(tot, 2, "TOTAL").font = _fonte(bold=True)
    for c in (2, 3, 4, 5):
        x = ws.cell(tot, c)
        if c > 2:
            x.value = f"=SUM({L(c)}3:{L(c)}{tot - 1})"
            x.number_format = FMT_CENT
            x.font = _fonte(bold=True)
        x.fill = PatternFill("solid", fgColor=CINZA_GRUPO)
        x.border = SECAO
    for c, w in zip("BCDE", (34, 24, 24, 18)):
        ws.column_dimensions[c].width = w
    return tot


def _aba_diario(wb, dias, tot_saldos, saldo_ini):
    ws = wb.create_sheet("FLUXO DIÁRIO")
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[2].height = 45
    hdr = [("Data", 12), ("Entradas", 15), ("Saídas", 15), ("Transferências", 15), ("Saldo realizado", 16),
           ("Entradas previstas (em aberto)", 16), ("Saídas previstas (em aberto)", 16), ("Saldo projetado", 16)]
    for i, (h, w) in enumerate(hdr, start=2):
        _cab(ws, 2, i, h)
        ws.column_dimensions[L(i)].width = w
    ws.column_dimensions["A"].width = 2
    r = 3
    ws.cell(r, 2, "SALDO INICIAL")
    ws.cell(r, 6, f"=SALDOS!C{tot_saldos}" if tot_saldos else round(saldo_ini, 2))
    ws.cell(r, 9, f"=F{r}")
    for c in range(2, 10):
        x = ws.cell(r, c)
        x.font = _fonte(bold=True)
        x.fill = PatternFill("solid", fgColor=CINZA_GRUPO)
        x.border = GRADE
        x.number_format = FMT_CENT
    for d in dias:
        r += 1
        vals = [d["data"], d["entradas"], d["saidas"], d["transferencias"], f"=F{r - 1}+SUM(C{r}:E{r})",
                d["entradas_prev"], d["saidas_prev"], f"=I{r - 1}+SUM(C{r}:E{r})+SUM(G{r}:H{r})"]
        for c, v in enumerate(vals, start=2):
            x = ws.cell(r, c, v)
            x.font = _fonte()
            x.border = GRADE
            x.number_format = "DD/MM/YYYY" if c == 2 else FMT_CENT
    r += 1
    ws.cell(r, 2, "TOTAL / SALDO FINAL")
    for c in range(2, 10):
        x = ws.cell(r, c)
        if c in (3, 4, 5, 7, 8):
            x.value = f"=SUM({L(c)}4:{L(c)}{r - 1})"
        elif c in (6, 9):
            x.value = f"={L(c)}{r - 1}"
        x.font = _fonte(bold=True, color="FFFFFF")
        x.fill = PatternFill("solid", fgColor=MARINHO)
        x.border = GRADE
        x.number_format = FMT_CENT
    ws.freeze_panes = "C3"
