"""Gera o Excel do fechamento no formato do modelo Savian: aba de REPORT com
fórmulas SUMIFS sobre a BASE CLASSIFICADA, mais SALDOS, RECEBIMENTOS etc."""
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

from . import dre as D
from .plano import IGNORAR

AZUL = "0B2A7A"
AZUL_CLARO = "DDEBF7"
CINZA = "D9D9D9"
CINZA_CLARO = "F2F2F2"
AMARELO = "FFFF00"
FMT = '#,##0.00;(#,##0.00);"-"'
FMT_PCT = '0%;-0%;"0%"'
fino = Side(style="thin", color="808080")
BORDA = Border(top=fino, bottom=fino, left=fino, right=fino)


def _cab(ws, row, col, texto, fill=AZUL, cor="FFFFFF"):
    c = ws.cell(row, col, texto)
    c.font = Font(bold=True, color=cor)
    c.fill = PatternFill("solid", fgColor=fill)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = BORDA
    return c


def gerar_excel(caminho, empresa, ano, mes, base, dre, saldos, cfg):
    wb = Workbook()
    ws = wb.active
    ws.title = f"REPORT {mes:02d} {ano}"
    n_base = len(base) + 1
    _aba_base(wb, base)
    _aba_report(ws, empresa, ano, mes, dre, n_base)
    _aba_saldos(wb, saldos, ano, mes)
    _aba_quadro(wb, "RECEBIMENTOS", D.recebimentos(base, cfg), "RECEBIDO", com_av=True)
    _aba_quadro(wb, "FORNECEDORES", D.pagamentos(base, cfg, "fornecedores"), "VALOR PAGO")
    _aba_quadro(wb, "FUNCIONARIOS", D.pagamentos(base, cfg, "funcionarios"), "VALOR PAGO")
    _aba_transferencias(wb, D.transferencias(base))
    _aba_indicadores(wb, D.indicadores(dre), ano, mes)
    wb.move_sheet("BASE CLASSIFICADA", offset=len(wb.sheetnames))
    wb.save(caminho)
    return caminho


def _aba_report(ws, empresa, ano, mes, dre, n_base):
    cols, agrup = dre["colunas"], dre["agrupamentos"]
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 52
    ws.cell(1, 2, f"{empresa['nome']} — DRE Financeiro {mes:02d}/{ano}").font = Font(bold=True, size=14, color=AZUL)
    _cab(ws, 2, 2, "CashFlowReport")
    _cab(ws, 3, 2, "Descrição da Conta")
    ws.freeze_panes = "C4"

    # layout das colunas: [coluna, coluna, ..., subtotal agrupamento], ..., TOTAL
    blocos, usados = [], set()
    for g, membros in agrup.items():
        for m in membros:
            blocos.append(("col", m))
            usados.add(m)
        blocos.append(("sub", g))
    for c in cols:
        if c not in usados:
            blocos.append(("col", c))
    blocos.append(("tot", f"TOTAL {empresa['nome']}".upper()))

    pos, cpos = {}, 3
    for tipo, nome in blocos:
        pos[(tipo, nome)] = cpos
        fill = AZUL if tipo == "col" else ("1F4E79" if tipo == "sub" else "000000")
        ws.merge_cells(start_row=2, start_column=cpos, end_row=2, end_column=cpos + 1)
        _cab(ws, 2, cpos, nome if tipo != "sub" else f"TOTAL {nome}".upper(), fill=fill)
        _cab(ws, 2, cpos + 1, None, fill=fill)
        _cab(ws, 3, cpos, f"{mes:02d}/{ano}", fill=fill)
        _cab(ws, 3, cpos + 1, "AV", fill=fill)
        ws.column_dimensions[L(cpos)].width = 15
        ws.column_dimensions[L(cpos + 1)].width = 7
        ws.column_dimensions[L(cpos + 2)].width = 2
        cpos += 3
    col_total = pos[("tot", blocos[-1][1])]
    col_conf = cpos
    _cab(ws, 2, col_conf, "CONFERÊNCIA (TOTAL − BASE)", fill="7F7F7F")
    ws.column_dimensions[L(col_conf)].width = 16

    rng_v = f"'BASE CLASSIFICADA'!$K$2:$K${n_base}"
    rng_c = f"'BASE CLASSIFICADA'!$G$2:$G${n_base}"
    rng_col = f"'BASE CLASSIFICADA'!$I$2:$I${n_base}"

    linhas = dre["linhas"]
    r0 = 4
    row_of = {i: r0 + i for i in range(len(linhas))}
    secao_row, grupos_por_secao, contas_por_grupo = {}, {}, {}
    for i, l in enumerate(linhas):
        if l["tipo"] == "secao":
            secao_row[l["secao"]] = row_of[i]
            grupos_por_secao[l["secao"]] = []
        elif l["tipo"] == "grupo":
            grupos_por_secao[l["secao"]].append(row_of[i])
            contas_por_grupo[row_of[i]] = []
            g_atual = row_of[i]
        elif l["tipo"] == "conta":
            contas_por_grupo[g_atual].append(row_of[i])

    from .plano import RESULTADOS, SECOES
    for i, l in enumerate(linhas):
        r = row_of[i]
        ws.cell(r, 1, l["nivel"]).font = Font(color="A6A6A6", size=8)
        c = ws.cell(r, 2, l["descricao"])
        bold = l["tipo"] != "conta"
        fill = {"secao": CINZA, "grupo": CINZA_CLARO, "resultado": AZUL_CLARO}.get(l["tipo"])
        c.font = Font(bold=bold, color=AZUL if l["tipo"] == "resultado" else "000000")
        if l["tipo"] == "conta":
            c.alignment = Alignment(indent=2)
        elif l["tipo"] == "grupo":
            c.alignment = Alignment(indent=1)
        for (tipo, nome), cp in pos.items():
            X = L(cp)
            if tipo == "col":
                if l["tipo"] == "conta":
                    f = f"=SUMIFS({rng_v},{rng_c},$B{r},{rng_col},\"{nome}\")"
                elif l["tipo"] == "grupo":
                    rs = contas_por_grupo[r]
                    f = f"=SUM({X}{rs[0]}:{X}{rs[-1]})" if rs else 0
                elif l["tipo"] == "secao":
                    rs = grupos_por_secao.get(l["secao"], [])
                    f = "=" + "+".join(f"{X}{x}" for x in rs) if rs else 0
                else:
                    if l["secao"] in RESULTADOS:
                        secs = RESULTADOS[l["secao"]][1]
                    else:
                        secs = [s for s, _ in SECOES]
                    f = "=" + "+".join(f"{X}{secao_row[s]}" for s in secs if s in secao_row)
            elif tipo == "sub":
                membros = agrup[nome]
                f = "=" + "+".join(f"{L(pos[('col', m)])}{r}" for m in membros)
            else:
                f = "=" + "+".join(f"{L(pos[('col', m)])}{r}" for m in cols) if cols else 0
            cv = ws.cell(r, cp, f)
            cv.number_format = FMT
            cv.font = Font(bold=bold or tipo != "col")
            ent = secao_row.get("ENTRADAS", r0)
            ca = ws.cell(r, cp + 1, f"=IFERROR({X}{r}/{X}${ent},0)")
            ca.number_format = FMT_PCT
            ca.font = Font(size=8, italic=True, color="595959")
            for cc in (cv, ca):
                if fill:
                    cc.fill = PatternFill("solid", fgColor=fill)
                cc.border = Border(bottom=Side(style="hair", color="BFBFBF"))
        if fill:
            c.fill = PatternFill("solid", fgColor=fill)
        c.border = Border(bottom=Side(style="hair", color="BFBFBF"))
        if l["tipo"] == "conta":
            T = L(col_total)
            cc = ws.cell(r, col_conf, f"=ROUND({T}{r}-SUMIFS({rng_v},{rng_c},$B{r}),2)")
            cc.number_format = FMT
            cc.font = Font(size=8, color="7F7F7F")

    fim = r0 + len(linhas)
    T = L(col_total)
    liq_row = r0 + next(i for i, l in enumerate(linhas) if l["descricao"] == D.RESULTADO_LIQUIDO)
    ws.cell(fim + 1, 2, "Conferência: TOTAL − soma da BASE CLASSIFICADA (deve ser 0)").font = Font(italic=True, size=9)
    rng_g = f"'BASE CLASSIFICADA'!$F$2:$F${n_base}"
    ws.cell(fim + 1, col_total, f"=ROUND({T}{liq_row}-SUMIFS({rng_v},{rng_g},\"<>{IGNORAR}\"),2)").number_format = FMT
    ws.cell(fim + 2, 2, "Conferência: resultado líquido + transferências − variação de caixa (deve ser 0)").font = Font(italic=True, size=9)
    ws.cell(fim + 2, col_total,
            f"=ROUND({T}{liq_row}+SUMIFS({rng_v},{rng_g},\"{IGNORAR}\")-SALDOS!E{{TOT}},2)").number_format = FMT
    ws._savian_conf_cell = ws.cell(fim + 2, col_total)
    ws.cell(fim + 4, 2, "Notas:").font = Font(bold=True)
    notas = [
        "• Valores puxados por SUMIFS da aba BASE CLASSIFICADA (cada lançamento do extrato OMIE + linhas de rateio).",
        "• Categorias com % no nome (ex.: \"Combustível (98%); Tarifas (2%)\") foram divididas entre as contas pelos percentuais.",
        "• DRE financeira (regime de caixa): considera apenas lançamentos realizados; transferências entre contas ficam fora.",
    ]
    for i, n in enumerate(notas):
        ws.cell(fim + 5 + i, 2, n).font = Font(size=9, color="595959")


def _aba_base(wb, base):
    ws = wb.create_sheet("BASE CLASSIFICADA")
    hdr = ["Data", "Cliente ou Fornecedor", "Conta Corrente", "Categoria (original)", "Componente",
           "Grupo", "Conta do Fechamento", "Projeto (original)", "Coluna (fechamento)", "Situação", "Valor (R$)",
           "Tipo", "Observações"]
    larg = [11, 38, 18, 34, 26, 9, 40, 22, 26, 13, 14, 11, 60]
    for i, (h, w) in enumerate(zip(hdr, larg), start=1):
        _cab(ws, 1, i, h)
        ws.column_dimensions[L(i)].width = w
    for r, b in enumerate(base, start=2):
        vals = [b["data"], b["favorecido"], b["conta_corrente"], b["categoria"], b["componente"], b["grupo"],
                b["conta_fechamento"], b["projeto"], b["coluna"], b["situacao"], b["valor"], b["tipo"],
                (b["observacoes"] or "")[:500]]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(r, c, v)
            if c == 1:
                cell.number_format = "DD/MM/YYYY"
            if c == 11:
                cell.number_format = FMT
            if b["tipo"] != "Lançamento":
                cell.fill = PatternFill("solid", fgColor="FFF2CC")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{L(len(hdr))}{len(base) + 1}"


def _aba_saldos(wb, saldos, ano, mes):
    import calendar
    ws = wb.create_sheet("SALDOS")
    ws.sheet_view.showGridLines = False
    ult = calendar.monthrange(ano, mes)[1]
    ini = f"{calendar.monthrange(ano - (mes == 1), 12 if mes == 1 else mes - 1)[1]:02d}/{12 if mes == 1 else mes - 1:02d}/{ano - (mes == 1)}"
    for i, h in enumerate(["CONTA", f"SALDO INICIAL {ini}", f"SALDO FINAL {ult:02d}/{mes:02d}/{ano}", "VARIAÇÃO"], start=2):
        _cab(ws, 2, i, h, fill=AMARELO, cor="000000")
    for r, s in enumerate(saldos, start=3):
        ws.cell(r, 2, s["conta"]).border = BORDA
        for c, v in ((3, s["saldo_inicial"]), (4, s["saldo_final"]), (5, f"=D{r}-C{r}")):
            x = ws.cell(r, c, v)
            x.number_format = FMT
            x.border = BORDA
    tot = 3 + len(saldos)
    ws.cell(tot, 2, "TOTAL").font = Font(bold=True)
    for c in (3, 4, 5):
        x = ws.cell(tot, c, f"=SUM({L(c)}3:{L(c)}{tot - 1})" if saldos else 0)
        x.number_format = FMT
        x.font = Font(bold=True)
        x.border = BORDA
    for col, w in zip("BCDE", (34, 24, 24, 18)):
        ws.column_dimensions[col].width = w
    rep = wb.worksheets[0]
    cel = getattr(rep, "_savian_conf_cell", None)
    if cel is not None:
        cel.value = cel.value.replace("{TOT}", str(tot))


def _aba_quadro(wb, nome, quadro, rotulo_valor, com_av=False):
    ws = wb.create_sheet(nome)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["B"].width = 50
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 9
    r = 2
    total_rows = []
    for bloco, itens in quadro.items():
        _cab(ws, r, 2, bloco, fill=AMARELO, cor="000000")
        _cab(ws, r, 3, rotulo_valor, fill=AMARELO, cor="000000")
        if com_av:
            _cab(ws, r, 4, "AV%", fill=AMARELO, cor="000000")
        r += 1
        ini = r
        for item, v in itens:
            ws.cell(r, 2, item).border = BORDA
            x = ws.cell(r, 3, v)
            x.number_format = FMT
            x.border = BORDA
            r += 1
        ws.cell(r, 2, f"TOTAL {bloco}").font = Font(bold=True, color="FFFFFF")
        x = ws.cell(r, 3, f"=SUM(C{ini}:C{r - 1})")
        x.number_format = FMT
        x.font = Font(bold=True, color="FFFFFF")
        for c in (2, 3):
            ws.cell(r, c).fill = PatternFill("solid", fgColor="000000")
        if com_av:
            for rr in range(ini, r):
                a = ws.cell(rr, 4, f"=IFERROR(C{rr}/C${r},0)")
                a.number_format = "0.0%"
                a.border = BORDA
        total_rows.append(r)
        r += 2
    if len(total_rows) > 1:
        ws.cell(r, 2, f"TOTAL GERAL").font = Font(bold=True)
        x = ws.cell(r, 3, "=" + "+".join(f"C{t}" for t in total_rows))
        x.number_format = FMT
        x.font = Font(bold=True)
    if not quadro:
        ws.cell(2, 2, "Sem lançamentos no mês para este quadro.")


def _aba_transferencias(wb, transf):
    ws = wb.create_sheet("TRANSFERENCIAS")
    ws.sheet_view.showGridLines = False
    ws.merge_cells("B2:D2")
    _cab(ws, 2, 2, "TRANSFERÊNCIAS ENTRE CONTAS DA EMPRESA", fill=AMARELO, cor="000000")
    for i, h in enumerate(["CONTA", "TRANSFERÊNCIA ENVIADA", "TRANSFERÊNCIA RECEBIDA"], start=2):
        _cab(ws, 3, i, h, fill=AMARELO, cor="000000")
    r = 4
    for conta, v in transf.items():
        ws.cell(r, 2, conta).border = BORDA
        for c, val in ((3, v["enviada"]), (4, v["recebida"])):
            x = ws.cell(r, c, round(val, 2))
            x.number_format = FMT
            x.border = BORDA
        r += 1
    ws.cell(r, 2, "TOTAL TRANSFERIDO").font = Font(bold=True)
    for c in (3, 4):
        x = ws.cell(r, c, f"=SUM({L(c)}4:{L(c)}{r - 1})" if transf else 0)
        x.number_format = FMT
        x.font = Font(bold=True)
        x.fill = PatternFill("solid", fgColor=AZUL_CLARO)
    for col, w in zip("BCD", (34, 24, 24)):
        ws.column_dimensions[col].width = w


def _aba_indicadores(wb, ind, ano, mes):
    ws = wb.create_sheet("INDICADORES")
    ws.sheet_view.showGridLines = False
    for i, h in enumerate(["Indicador", f"{mes:02d}/{ano}", "% dos recebimentos"], start=2):
        _cab(ws, 2, i, h)
    for r, (k, v) in enumerate(ind.items(), start=3):
        ws.cell(r, 2, k).border = BORDA
        a = ws.cell(r, 3, v["valor"])
        a.number_format = FMT
        a.border = BORDA
        b = ws.cell(r, 4, v["pct"])
        b.number_format = "0.0%"
        b.border = BORDA
        if k.startswith("Resultado") or k in ("Entradas", "Receita Líquida"):
            for c in (2, 3, 4):
                ws.cell(r, c).font = Font(bold=True)
    for col, w in zip("BCD", (30, 20, 20)):
        ws.column_dimensions[col].width = w
