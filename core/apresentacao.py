"""Gera o PowerPoint do Report Mensal no padrão visual da Savian Contabilidade."""
import os

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt
from lxml import etree

from . import dre as D
from .analise import brl, pct

ASSETS = os.path.join(os.path.dirname(os.path.dirname(__file__)), "assets")
AZUL = RGBColor(0x0B, 0x2A, 0x7A)
AZUL2 = RGBColor(0x00, 0x77, 0xC8)
PRETO = RGBColor(0, 0, 0)
BRANCO = RGBColor(0xFF, 0xFF, 0xFF)
AMARELO = RGBColor(0xFF, 0xFF, 0x00)
CINZA = RGBColor(0xE7, 0xE6, 0xE6)
AZUL_CLARO = RGBColor(0xDD, 0xEB, 0xF7)
FONTE = "Calibri"
LARG, ALT = 960, 540  # pontos (16:9)


def _p(x):
    return Pt(x)


def _fundo(slide, arquivo):
    slide.shapes.add_picture(os.path.join(ASSETS, arquivo), 0, 0, _p(LARG), _p(ALT))


def _texto(slide, x, y, w, h, texto, tam=14, cor=PRETO, negrito=False, alinh=PP_ALIGN.LEFT, italico=False,
           ancora=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(_p(x), _p(y), _p(w), _p(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = ancora
    for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, m, Pt(2))
    linhas = texto if isinstance(texto, list) else [texto]
    for i, ln in enumerate(linhas):
        par = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        par.alignment = alinh
        if isinstance(ln, tuple):  # (texto, dict de estilo)
            ln, est = ln
        else:
            est = {}
        r = par.add_run()
        r.text = ln
        f = r.font
        f.name = FONTE
        f.size = Pt(est.get("tam", tam))
        f.bold = est.get("negrito", negrito)
        f.italic = est.get("italico", italico)
        f.color.rgb = est.get("cor", cor)
        if est.get("espaco"):
            par.space_before = Pt(est["espaco"])
    return tb


def _pilula(slide, titulo):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, _p(188), _p(33), _p(551), _p(52))
    s.adjustments[0] = 0.5
    s.line.fill.background()
    s.shadow.inherit = False
    fill = s.fill
    fill.gradient()
    fill.gradient_angle = 0
    st = fill.gradient_stops
    st[0].color.rgb = AZUL2
    st[0].position = 0
    st[1].color.rgb = RGBColor(0x0B, 0x3D, 0x91)
    st[1].position = 1.0
    tf = s.text_frame
    tf.margin_left = tf.margin_right = Pt(10)
    par = tf.paragraphs[0]
    par.alignment = PP_ALIGN.CENTER
    r = par.add_run()
    r.text = titulo
    r.font.name = FONTE
    r.font.size = Pt(20 if len(titulo) < 48 else 16)
    r.font.bold = True
    r.font.color.rgb = BRANCO


def _slide_conteudo(prs, titulo):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _fundo(s, "fundo_conteudo.png")
    _pilula(s, titulo)
    return s


def _borda(cell, cor="000000", larg=6350):
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        for e in tcPr.findall(qn(tag)):
            tcPr.remove(e)
        ln = etree.SubElement(tcPr, qn(tag), w=str(larg), cap="flat", cmpd="sng", algn="ctr")
        sf = etree.SubElement(ln, qn("a:solidFill"))
        etree.SubElement(sf, qn("a:srgbClr"), val=cor)
        etree.SubElement(ln, qn("a:prstDash"), val="solid")


def _cel(cell, texto, tam=9, negrito=False, cor=PRETO, fundo=None, alinh=PP_ALIGN.LEFT, borda=True):
    cell.text = ""
    tf = cell.text_frame
    tf.word_wrap = True
    cell.margin_left = cell.margin_right = Pt(3)
    cell.margin_top = cell.margin_bottom = Pt(0.5)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    par = tf.paragraphs[0]
    par.alignment = alinh
    r = par.add_run()
    r.text = str(texto)
    r.font.name = FONTE
    r.font.size = Pt(tam)
    r.font.bold = negrito
    r.font.color.rgb = cor
    if fundo is None:
        cell.fill.background()
    else:
        cell.fill.solid()
        cell.fill.fore_color.rgb = fundo
    if borda:
        _borda(cell)


def _tabela(slide, x, y, w, larguras, n_linhas, alt_linha):
    shp = slide.shapes.add_table(n_linhas, len(larguras), _p(x), _p(y), _p(w), _p(alt_linha * n_linhas))
    tbl = shp.table
    # remove estilo padrão (faixas azuis)
    tblPr = tbl._tbl.tblPr
    tblPr.set("firstRow", "0")
    tblPr.set("bandRow", "0")
    tot = sum(larguras)
    for i, lw in enumerate(larguras):
        tbl.columns[i].width = Emu(int(_p(w) * lw / tot))
    for r in tbl.rows:
        r.height = _p(alt_linha)
    return tbl


def _num(v):
    """formato da DRE: milhar sem centavos, negativo entre parênteses."""
    s = f"{abs(v):,.0f}".replace(",", ".")
    return f"({s})" if round(v) < 0 else s


# ---------------- slides ----------------

def _capa(prs, empresa, ano, mes):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _fundo(s, "fundo_capa.jpg")
    s.shapes.add_picture(os.path.join(ASSETS, "logo_savian.png"), _p(53), _p(33), _p(247))
    _texto(s, 128, 236, 330, 40, "Report Mensal", tam=28, cor=BRANCO, negrito=True)
    _texto(s, 128, 286, 330, 40, f"{mes:02d}/{ano}", tam=28, cor=BRANCO, negrito=True)
    ln = s.shapes.add_connector(1, _p(132), _p(333), _p(260), _p(333))
    ln.line.color.rgb = RGBColor(0x4F, 0xC3, 0xF7)
    ln.line.width = Pt(2)
    linhas = [("EMPRESA", {"negrito": True})] + [(empresa["nome"].upper(), {"negrito": True})]
    for c in [x.strip() for x in str(empresa.get("cnpj") or "").replace(";", "\n").split("\n") if x.strip()]:
        linhas.append((c, {"negrito": True}))
    _texto(s, 125, 352, 300, 150, linhas, tam=16, cor=BRANCO)


def _servicos(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _fundo(s, "slide_servicos.jpg")


def _blocos_dre(dre, max_cols=5):
    """Divide as colunas da DRE em páginas. Cada item: ('col'|'sub'|'tot', nome, valores)."""
    cols, agrup = dre["colunas"], dre["agrupamentos"]
    seq, usados = [], set()
    for g, membros in agrup.items():
        for m in membros:
            seq.append(("col", m))
            usados.add(m)
        seq.append(("sub", g))
    seq += [("col", c) for c in cols if c not in usados]
    if len(cols) == 1 and not agrup:
        seq = []  # só uma coluna: mostra direto o total
    seq.append(("tot", None))
    paginas, atual = [], []
    for item in seq:
        atual.append(item)
        if len(atual) >= max_cols and item[0] != "tot":
            paginas.append(atual)
            atual = []
    if atual:
        paginas.append(atual)
    return paginas


def _valor_bloco(l, tipo, nome, dre):
    if tipo == "col":
        return l["valores"].get(nome, 0.0)
    if tipo == "sub":
        return sum(l["valores"].get(m, 0.0) for m in dre["agrupamentos"][nome])
    return l["total"]


def _slide_dre(prs, empresa, ano, mes, dre, detalhado=False):
    linhas = [l for l in dre["linhas"] if detalhado or l["nivel"] <= 2]
    ent_l = D.linha(dre, "(+) ENTRADAS")
    paginas = [[("tot", None)]] if detalhado else _blocos_dre(dre, max_cols=5)
    por_pag = 34 if not detalhado else 30
    mes_txt = f"{['jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez'][mes - 1]}-{str(ano)[2:]}"
    for pi, blocos in enumerate(paginas):
        for ri in range(0, len(linhas), por_pag):
            pedaco = linhas[ri:ri + por_pag]
            tit = f"Gerencial Report – DRE {'DETALHADA' if detalhado else 'FINANCEIRO'} {mes:02d}-{ano}"
            s = _slide_conteudo(prs, tit)
            n = len(pedaco) + 2
            alt = min(13.0, 380.0 / n)
            tam = 8 if alt >= 11 else 7
            larg = [3.4] + [1.05, 0.5] * len(blocos)
            w = min(870, 260 + 150 * len(blocos)) if not detalhado else 700
            x = (LARG - w) / 2
            tbl = _tabela(s, x, 100, w, larg, n, alt)
            _cel(tbl.cell(0, 0), "CashFlowReport - Por " + {"projeto": "Projeto", "conta": "Conta"}.get(
                dre.get("modo", ""), "Conta"), tam=tam, negrito=True, alinh=PP_ALIGN.CENTER, borda=False)
            _cel(tbl.cell(1, 0), "Descrição da Conta", tam=tam, negrito=True, alinh=PP_ALIGN.CENTER, borda=False)
            for bi, (tipo, nome) in enumerate(blocos):
                c = 1 + bi * 2
                rot = nome if tipo == "col" else (f"TOTAL {nome}" if tipo == "sub" else f"TOTAL {empresa['nome']}")
                fundo = None if tipo == "col" else (AZUL if tipo == "sub" else PRETO)
                cor = PRETO if tipo == "col" else BRANCO
                tbl.cell(0, c).merge(tbl.cell(0, c + 1))
                _cel(tbl.cell(0, c), rot.upper(), tam=tam - 0.5, negrito=True, cor=cor, fundo=fundo, alinh=PP_ALIGN.CENTER)
                _cel(tbl.cell(1, c), mes_txt, tam=tam, negrito=True, alinh=PP_ALIGN.CENTER)
                _cel(tbl.cell(1, c + 1), "AV", tam=tam, negrito=True, alinh=PP_ALIGN.CENTER)
            for r, l in enumerate(pedaco, start=2):
                destaque = l["nivel"] == 1
                fundo = CINZA if l["nivel"] == 2 else (AZUL_CLARO if l["tipo"] == "resultado" else None)
                ind = "    " if l["nivel"] == 3 else ""
                _cel(tbl.cell(r, 0), ind + l["descricao"], tam=tam, negrito=destaque, fundo=fundo, borda=False)
                for bi, (tipo, nome) in enumerate(blocos):
                    c = 1 + bi * 2
                    v = _valor_bloco(l, tipo, nome, dre)
                    base_ent = _valor_bloco(ent_l, tipo, nome, dre) if ent_l.get("valores") else 0
                    av = v / base_ent if base_ent else 0
                    _cel(tbl.cell(r, c), _num(v), tam=tam, negrito=destaque or tipo != "col", fundo=fundo,
                         alinh=PP_ALIGN.RIGHT, borda=False)
                    _cel(tbl.cell(r, c + 1), pct(av), tam=tam, negrito=destaque, fundo=fundo, alinh=PP_ALIGN.RIGHT,
                         borda=False)


def _slide_saldos(prs, saldos, ano, mes):
    import calendar
    if not saldos:
        return
    s = _slide_conteudo(prs, "Gerencial Report – SALDOS NAS CONTAS")
    pm, pa = (12, ano - 1) if mes == 1 else (mes - 1, ano)
    ini = f"{calendar.monthrange(pa, pm)[1]:02d}/{pm:02d}/{pa}"
    fim = f"{calendar.monthrange(ano, mes)[1]:02d}/{mes:02d}/{ano}"
    n = len(saldos) + 2
    alt = min(28, 340 / n)
    tam = 16 if alt >= 24 else 12
    w = 630
    tbl = _tabela(s, (LARG - w) / 2, 150, w, [2.1, 2.1, 2.1], n, alt)
    for c, h in enumerate(["CONTA", f"SALDO INICIAL {ini}", f"SALDO FINAL {fim}"]):
        _cel(tbl.cell(0, c), h, tam=tam, negrito=True, fundo=AMARELO)
    for r, sd in enumerate(saldos, start=1):
        _cel(tbl.cell(r, 0), sd["conta"].upper(), tam=tam)
        _cel(tbl.cell(r, 1), brl(sd["saldo_inicial"]), tam=tam, alinh=PP_ALIGN.RIGHT)
        _cel(tbl.cell(r, 2), brl(sd["saldo_final"]), tam=tam, alinh=PP_ALIGN.RIGHT)
    r = n - 1
    _cel(tbl.cell(r, 0), "TOTAL", tam=tam, negrito=True)
    _cel(tbl.cell(r, 1), brl(sum(x["saldo_inicial"] for x in saldos)), tam=tam, negrito=True, alinh=PP_ALIGN.RIGHT)
    _cel(tbl.cell(r, 2), brl(sum(x["saldo_final"] for x in saldos)), tam=tam, negrito=True, alinh=PP_ALIGN.RIGHT)


def _slide_blocos(prs, titulo, quadro, total_geral_rot=None, max_itens=40):
    """Quadros 'título amarelo + itens + total' distribuídos em até 3 colunas por slide."""
    if not quadro:
        return
    cartoes = []
    for bloco, itens in quadro.items():
        if len(itens) > max_itens:
            resto = itens[max_itens - 1:]
            itens = itens[:max_itens - 1] + [(f"DEMAIS ({len(resto)})", sum(v for _, v in resto))]
        total = sum(v for _, v in itens)
        cap = 28
        for i in range(0, len(itens), cap):
            parte = itens[i:i + cap]
            ultimo = i + cap >= len(itens)
            cartoes.append((bloco + (" (cont.)" if i else ""), parte, total if ultimo else None))
    geral = sum(v for its in quadro.values() for _, v in its)
    alt = 14.5
    lim = 30  # linhas por coluna
    paginas, pag, col, usado = [], [[]], 0, 0
    for c in cartoes:
        h = len(c[1]) + 1 + (1 if c[2] is not None else 0)
        if usado and usado + h + 1 > lim:
            col += 1
            usado = 0
            if col == 3:
                paginas.append(pag)
                pag, col = [[]], 0
            else:
                pag.append([])
        pag[-1].append(c)
        usado += h + 1
    paginas.append(pag)
    for pi, pag in enumerate(paginas):
        s = _slide_conteudo(prs, titulo)
        ncol = len(pag)
        wcol = {1: 520, 2: 380, 3: 290}[ncol]
        gap = 12
        x0 = (LARG - (wcol * ncol + gap * (ncol - 1))) / 2
        tam = 10 if ncol == 1 else 8
        for ci, coluna in enumerate(pag):
            y = 105
            for bloco, itens, total in coluna:
                n = len(itens) + 1 + (1 if total is not None else 0)
                tbl = _tabela(s, x0 + ci * (wcol + gap), y, wcol, [3.2, 1.1], n, alt)
                _cel(tbl.cell(0, 0), bloco, tam=tam, negrito=True, fundo=AMARELO)
                _cel(tbl.cell(0, 1), "VALOR", tam=tam, negrito=True, fundo=AMARELO)
                for r, (item, v) in enumerate(itens, start=1):
                    _cel(tbl.cell(r, 0), str(item).upper()[:60], tam=tam)
                    _cel(tbl.cell(r, 1), brl(v), tam=tam, alinh=PP_ALIGN.RIGHT)
                if total is not None:
                    _cel(tbl.cell(n - 1, 0), f"TOTAL", tam=tam, negrito=True, cor=BRANCO, fundo=PRETO)
                    _cel(tbl.cell(n - 1, 1), brl(total), tam=tam, negrito=True, cor=BRANCO, fundo=PRETO,
                         alinh=PP_ALIGN.RIGHT)
                y += n * alt + alt
        if total_geral_rot and pi == len(paginas) - 1 and len(quadro) > 1:
            w = 520
            tbl = _tabela(s, (LARG - w) / 2, 500, w, [3.2, 1.1], 1, 18)
            _cel(tbl.cell(0, 0), total_geral_rot, tam=11, negrito=True, cor=BRANCO, fundo=PRETO)
            _cel(tbl.cell(0, 1), brl(geral), tam=11, negrito=True, cor=BRANCO, fundo=PRETO, alinh=PP_ALIGN.RIGHT)


def _slide_transf(prs, transf):
    if not transf:
        return
    s = _slide_conteudo(prs, "Gerencial Report – TRANSFERÊNCIAS ENTRE CONTAS")
    n = len(transf) + 3
    alt = 19
    w = 620
    tbl = _tabela(s, (LARG - w) / 2, 200, w, [2.2, 1.9, 1.9], n, alt)
    tbl.cell(0, 0).merge(tbl.cell(0, 2))
    _cel(tbl.cell(0, 0), "TRANSFERÊNCIAS ENTRE CONTAS DA EMPRESA", tam=12, negrito=True, fundo=AMARELO,
         alinh=PP_ALIGN.CENTER)
    for c, h in enumerate(["CONTA", "TRANSFERÊNCIA ENVIADA", "TRANSFERÊNCIA RECEBIDA"]):
        _cel(tbl.cell(1, c), h, tam=12, negrito=True, fundo=AMARELO)
    for r, (conta, v) in enumerate(transf.items(), start=2):
        _cel(tbl.cell(r, 0), conta.upper(), tam=12)
        _cel(tbl.cell(r, 1), brl(v["enviada"]) if v["enviada"] else "R$ -", tam=12, alinh=PP_ALIGN.RIGHT)
        _cel(tbl.cell(r, 2), brl(v["recebida"]) if v["recebida"] else "R$ -", tam=12, alinh=PP_ALIGN.RIGHT)
    _cel(tbl.cell(n - 1, 0), "TOTAL TRANSFERIDO", tam=12, fundo=AZUL_CLARO)
    _cel(tbl.cell(n - 1, 1), brl(sum(v["enviada"] for v in transf.values())), tam=12, fundo=AZUL_CLARO, alinh=PP_ALIGN.RIGHT)
    _cel(tbl.cell(n - 1, 2), brl(sum(v["recebida"] for v in transf.values())), tam=12, fundo=AZUL_CLARO, alinh=PP_ALIGN.RIGHT)


def _slide_analise(prs, textos, dre, ano, mes, empresa):
    s = _slide_conteudo(prs, f"Gerencial Report – Análise {mes:02d}/{ano}")
    fech = [x for x in textos["fechamento"].split("\n") if x.strip()]
    tam = 10 if len(fech) <= 10 else 9
    linhas = [(textos["titulo"], {"negrito": True, "tam": 15, "cor": AZUL}),
              ("Valores consolidados (coluna TOTAL do relatório)", {"italico": True, "tam": 9, "cor": RGBColor(0x59, 0x59, 0x59)}),
              ("Fechamento do mês", {"negrito": True, "tam": 12, "cor": AZUL, "espaco": 8})]
    for f in fech:
        if ":" in f:
            linhas.append(("• " + f, {"tam": tam, "espaco": 2}))
        else:
            linhas.append(("• " + f, {"tam": tam, "espaco": 2}))
    _texto(s, 95, 100, 800, 230, linhas)
    ind = D.indicadores(dre)
    n = len(ind) + 1
    alt = 15
    tbl = _tabela(s, 180, 545 - 40 - n * alt, 600, [2.6, 1.6, 1.3], n, alt)
    for c, h in enumerate(["Indicador", f"{['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'][mes-1]}/{ano}", "% dos recebimentos"]):
        _cel(tbl.cell(0, c), h, tam=9, negrito=True, cor=BRANCO, fundo=AZUL, alinh=PP_ALIGN.CENTER)
    for r, (k, v) in enumerate(ind.items(), start=1):
        forte = k in ("Entradas", "Receita Líquida", "Resultado Bruto", "Resultado Operacional", "Resultado Líquido")
        fundo = RGBColor(0xE2, 0xEF, 0xDA) if k in ("Resultado Operacional", "Resultado Líquido") else None
        _cel(tbl.cell(r, 0), k, tam=9, negrito=forte, fundo=fundo)
        _cel(tbl.cell(r, 1), brl(v["valor"]), tam=9, negrito=forte, fundo=fundo, alinh=PP_ALIGN.RIGHT)
        _cel(tbl.cell(r, 2), pct(v["pct"], 1), tam=9, negrito=forte, fundo=fundo, alinh=PP_ALIGN.RIGHT)


def _slide_pontos(prs, textos, ano, mes):
    s = _slide_conteudo(prs, f"Gerencial Report – Análise {mes:02d}/{ano}")
    pos = [x for x in textos["positivos"].split("\n") if x.strip()]
    ate = [x for x in textos["atencao"].split("\n") if x.strip()]
    linhas = [("☆ Principais pontos positivos:", {"negrito": True, "tam": 14, "cor": AZUL})]
    linhas += [("•  " + x, {"tam": 11, "espaco": 2}) for x in pos]
    linhas += [("⚲ Principais pontos de atenção:", {"negrito": True, "tam": 14, "cor": AZUL, "espaco": 14})]
    linhas += [("•  " + x, {"tam": 11, "espaco": 2}) for x in ate]
    if textos.get("conclusao"):
        linhas += [(textos["conclusao"], {"italico": True, "tam": 11, "espaco": 16})]
    _texto(s, 80, 115, 800, 390, linhas)


def gerar_pptx(caminho, empresa, ano, mes, base, dre, saldos, cfg, textos, incluir_detalhada=True):
    prs = Presentation()
    prs.slide_width = _p(LARG)
    prs.slide_height = _p(ALT)
    dre = dict(dre, modo=cfg.get("modo_colunas"))
    _capa(prs, empresa, ano, mes)
    _servicos(prs)
    _slide_dre(prs, empresa, ano, mes, dre)
    if incluir_detalhada:
        _slide_dre(prs, empresa, ano, mes, dre, detalhado=True)
    _slide_saldos(prs, saldos, ano, mes)
    _slide_blocos(prs, f"Gerencial Report – RECEBIMENTOS {mes:02d}/{ano}", D.recebimentos(base, cfg),
                  f"TOTAL RECEBIMENTO {mes:02d}/{ano}")
    _slide_blocos(prs, "Gerencial Report – PAGAMENTO DE FORNECEDORES", D.pagamentos(base, cfg, "fornecedores"),
                  "TOTAL FORNECEDORES")
    _slide_blocos(prs, "Gerencial Report – PAGAMENTOS POR FUNCIONÁRIO", D.pagamentos(base, cfg, "funcionarios"))
    _slide_transf(prs, D.transferencias(base))
    _slide_analise(prs, textos, dre, ano, mes, empresa)
    _slide_pontos(prs, textos, ano, mes)
    prs.save(caminho)
    return caminho
