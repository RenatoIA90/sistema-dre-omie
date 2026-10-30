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
CINZA = RGBColor(0xE8, 0xE8, 0xE8)
# cores e fonte do REPORT do Excel (modelo HAGILI), usadas nas tabelas
MARINHO = RGBColor(0x1F, 0x38, 0x64)
AZUL_MEDIO = RGBColor(0x2F, 0x54, 0x96)
AZUL_TOTAL = RGBColor(0xD9, 0xE2, 0xF3)
ZEBRA = RGBColor(0xF2, 0xF5, 0xFA)
GRADE_CLARA = "BFC9D9"
FINA, MEDIA = 6350, 19050  # espessura das bordas (EMU)
FONTE = "Calibri"
FONTE_TAB = "Aptos Narrow"
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


def _borda(cell, cor="000000", larg=FINA, lados="LRTB"):
    tcPr = cell._tc.get_or_add_tcPr()
    for lado in lados:
        tag = "a:ln" + lado
        for e in tcPr.findall(qn(tag)):
            tcPr.remove(e)
        ln = etree.Element(qn(tag), w=str(larg), cap="flat", cmpd="sng", algn="ctr")
        sf = etree.SubElement(ln, qn("a:solidFill"))
        etree.SubElement(sf, qn("a:srgbClr"), val=cor)
        etree.SubElement(ln, qn("a:prstDash"), val="solid")
        # o schema exige a ordem lnL, lnR, lnT, lnB e todas antes do preenchimento da célula
        ordem = ["a:lnL", "a:lnR", "a:lnT", "a:lnB"]
        depois = [qn(t) for t in ordem[ordem.index(tag) + 1:]] + [qn("a:noFill"), qn("a:solidFill")]
        prox = next((e for e in tcPr if e.tag in depois), None)
        if prox is None:
            tcPr.append(ln)
        else:
            prox.addprevious(ln)


def _cel(cell, texto, tam=9, negrito=False, cor=PRETO, fundo=None, alinh=PP_ALIGN.LEFT, borda=False,
         italico=False):
    cell.text = ""
    tf = cell.text_frame
    tf.word_wrap = True
    cell.margin_left = cell.margin_right = Pt(min(4, tam * 0.45))
    cell.margin_top = cell.margin_bottom = Pt(0.5)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    par = tf.paragraphs[0]
    par.alignment = alinh
    r = par.add_run()
    r.text = str(texto)
    r.font.name = FONTE_TAB
    r.font.size = Pt(tam)
    r.font.bold = negrito
    r.font.italic = italico
    r.font.color.rgb = cor
    # sem isso a célula vazia usa 18pt e a linha fica mais alta que o definido
    par._p.get_or_add_endParaRPr().set("sz", str(int(tam * 100)))
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
    # estilo "sem estilo, sem grade": só aparecem as bordas e cores definidas aqui
    tblPr = tbl._tbl.tblPr
    tblPr.set("firstRow", "0")
    tblPr.set("bandRow", "0")
    sid = tblPr.find(qn("a:tableStyleId"))
    if sid is not None:
        sid.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"
    tot = sum(larguras)
    for i, lw in enumerate(larguras):
        tbl.columns[i].width = Emu(int(_p(w) * lw / tot))
    for r in tbl.rows:
        r.height = _p(alt_linha)
    return tbl


def _num(v):
    """formato da DRE: milhar sem centavos, negativo entre parênteses, zero como traço."""
    if round(v) == 0:
        return "-"
    s = f"{abs(v):,.0f}".replace(",", ".")
    return f"({s})" if v < 0 else s


def _av(v):
    x = round(v * 100)
    return "-" if x == 0 else f"{x:.0f}%"


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


def _blocos_dre(dre):
    """Colunas da DRE na ordem do Excel. Cada item: ('col'|'sub'|'tot', nome)."""
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
    return seq + [("tot", None)]


def _cabe(texto, largura, tam):
    """Corta o texto para caber em uma linha da largura dada (pontos) na fonte das tabelas."""
    n = max(6, int(largura / (tam * 0.56)))
    texto = str(texto)
    return texto if len(texto) <= n else texto[:n - 1].rstrip() + "…"


def _valor_bloco(l, tipo, nome, dre):
    if tipo == "col":
        return l["valores"].get(nome, 0.0)
    if tipo == "sub":
        return sum(l["valores"].get(m, 0.0) for m in dre["agrupamentos"][nome])
    return l["total"]


def _slide_dre(prs, empresa, ano, mes, dre):
    """DRE financeira (seções, resultados e grupos) no visual do REPORT do Excel (modelo HAGILI):
    cabeçalhos com borda fina e só os totais em azul-marinho, seções e resultados com borda média
    em cima e embaixo, grupos em cinza, AV% em itálico e valores sem centavos.
    Sempre em um slide só: com muitas colunas ou linhas a fonte e a altura das linhas diminuem."""
    linhas = [l for l in dre["linhas"] if l["nivel"] <= 2]
    ent_l = D.linha(dre, "(+) ENTRADAS")
    mes_txt = f"{['jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez'][mes - 1]}-{str(ano)[2:]}"
    titulo = "CashFlowReport - Por " + {"projeto": "Projeto", "conta": "Conta"}.get(dre.get("modo", ""), "Conta")
    blocos = _blocos_dre(dre)
    s = _slide_conteudo(prs, f"Gerencial Report – DRE FINANCEIRO {mes:02d}-{ano}")
    n = len(linhas) + 2
    alt = min(13.5, 405.0 / (n + 0.6))
    nb = len(blocos)
    tam = min(9 if alt >= 12 else 8 if alt >= 10.5 else 7 if alt >= 9 else 6,
              9 if nb <= 4 else 8 if nb <= 6 else 7 if nb <= 8 else 6 if nb <= 11 else 5)
    # descrição | valor, AV e um espaçador estreito por bloco (o último sem espaçador)
    desc = 3.4 if nb <= 5 else 2.6
    larg = [desc]
    for bi in range(nb):
        larg += [1.05, 0.5] + ([0.1] if bi < nb - 1 else [])
    w = min(840, 270 + 150 * nb)  # sem invadir o enfeite do canto direito
    w_desc = w * desc / sum(larg)
    linhas = [dict(l, descricao=_cabe(l["descricao"], w_desc - 8, tam)) for l in linhas]
    y = 100 + max(0, (410 - (n + 0.6) * alt) / 2)
    tbl = _tabela(s, (LARG - w) / 2, y, w, larg, n, alt)
    tbl.rows[0].height = _p(alt * 1.6)
    _cel(tbl.cell(0, 0), titulo, tam=tam + 1, negrito=True, alinh=PP_ALIGN.CENTER)
    _cel(tbl.cell(1, 0), "Descrição da Conta", tam=tam, negrito=True, alinh=PP_ALIGN.CENTER)
    for c in range(1, len(larg)):
        _cel(tbl.cell(0, c), "", tam=tam)
        _cel(tbl.cell(1, c), "", tam=tam)
    for bi, (tipo, nome) in enumerate(blocos):
        c = 1 + bi * 3
        rot = nome if tipo == "col" else (f"TOTAL {nome}" if tipo == "sub" else f"TOTAL {empresa['nome']}")
        destaque = tipo != "col"
        tbl.cell(0, c).merge(tbl.cell(0, c + 1))
        _cel(tbl.cell(0, c), rot.upper(), tam=tam, negrito=True, cor=BRANCO if destaque else PRETO,
             fundo=MARINHO if destaque else None, alinh=PP_ALIGN.CENTER, borda=True)
        _borda(tbl.cell(0, c + 1))
        _cel(tbl.cell(1, c), mes_txt, tam=tam, negrito=True, alinh=PP_ALIGN.CENTER)
        _cel(tbl.cell(1, c + 1), "AV", tam=tam, negrito=True, alinh=PP_ALIGN.CENTER)
    for r, l in enumerate(linhas, start=2):
        forte = l["tipo"] != "conta"
        fundo = CINZA if l["tipo"] == "grupo" else None
        for c in range(len(larg)):
            _cel(tbl.cell(r, c), "", tam=tam, fundo=fundo)
        _cel(tbl.cell(r, 0), l["descricao"], tam=tam, negrito=forte, fundo=fundo)
        for bi, (tipo, nome) in enumerate(blocos):
            c = 1 + bi * 3
            v = _valor_bloco(l, tipo, nome, dre)
            base_ent = _valor_bloco(ent_l, tipo, nome, dre) if ent_l.get("valores") else 0
            av = v / base_ent if base_ent else 0
            _cel(tbl.cell(r, c), _num(v), tam=tam, negrito=forte, fundo=fundo, alinh=PP_ALIGN.RIGHT)
            _cel(tbl.cell(r, c + 1), _av(av), tam=tam - 0.5, negrito=forte, italico=True, fundo=fundo,
                 alinh=PP_ALIGN.RIGHT)
        if l["tipo"] in ("secao", "resultado"):
            for c in range(len(larg)):
                _borda(tbl.cell(r, c), larg=MEDIA, lados="TB")
    # espaçadores sem margem, senão o PowerPoint alarga a coluna e a tabela
    for c in range(3, len(larg), 3):
        for r in range(n):
            tbl.cell(r, c).margin_left = tbl.cell(r, c).margin_right = 0


# ---- quadros (saldos, recebimentos, fornecedores, funcionários, transferências) ----

def _cab_quadro(cell, texto, tam, fundo=MARINHO):
    _cel(cell, texto, tam=tam, negrito=True, cor=BRANCO, fundo=fundo, alinh=PP_ALIGN.CENTER)
    _borda(cell, cor=GRADE_CLARA)


def _item_quadro(cell, texto, tam, zebra):
    _cel(cell, texto, tam=tam, fundo=ZEBRA if zebra else BRANCO, alinh=PP_ALIGN.CENTER)
    _borda(cell, cor=GRADE_CLARA)


def _total_quadro(cell, texto, tam):
    _cel(cell, texto, tam=tam, negrito=True, cor=MARINHO, fundo=AZUL_TOTAL, alinh=PP_ALIGN.CENTER)
    _borda(cell, cor=GRADE_CLARA)
    _borda(cell, cor="1F3864", larg=MEDIA, lados="T")


def _y_centro(altura, topo=100, base=515):
    """topo para centralizar verticalmente um conteúdo na área abaixo do título."""
    return topo + max(0, (base - topo - altura) / 2)


def _slide_saldos(prs, saldos, ano, mes):
    import calendar
    saldos = [x for x in saldos or [] if round(x["saldo_inicial"], 2) or round(x["saldo_final"], 2)]
    if not saldos:
        return
    s = _slide_conteudo(prs, "Gerencial Report – SALDOS NAS CONTAS")
    pm, pa = (12, ano - 1) if mes == 1 else (mes - 1, ano)
    ini = f"{calendar.monthrange(pa, pm)[1]:02d}/{pm:02d}/{pa}"
    fim = f"{calendar.monthrange(ano, mes)[1]:02d}/{mes:02d}/{ano}"
    n = len(saldos) + 2
    alt = min(30, 400 / n)
    tam = min(14, round(alt * 0.55 * 2) / 2)
    w = 680
    tbl = _tabela(s, (LARG - w) / 2, _y_centro(n * alt), w, [2.4, 2.0, 2.0], n, alt)
    for c, h in enumerate(["CONTA", f"SALDO INICIAL {ini}", f"SALDO FINAL {fim}"]):
        _cab_quadro(tbl.cell(0, c), h, tam)
    for r, sd in enumerate(saldos, start=1):
        z = r % 2 == 0
        _item_quadro(tbl.cell(r, 0), sd["conta"].upper(), tam, z)
        _item_quadro(tbl.cell(r, 1), brl(sd["saldo_inicial"]), tam, z)
        _item_quadro(tbl.cell(r, 2), brl(sd["saldo_final"]), tam, z)
    r = n - 1
    _total_quadro(tbl.cell(r, 0), "TOTAL", tam)
    _total_quadro(tbl.cell(r, 1), brl(sum(x["saldo_inicial"] for x in saldos)), tam)
    _total_quadro(tbl.cell(r, 2), brl(sum(x["saldo_final"] for x in saldos)), tam)


def _resumir(itens, cap):
    """Mantém os maiores itens e junta o resto em 'DEMAIS (n)', totalizando cap linhas."""
    if len(itens) <= cap:
        return itens
    resto = itens[cap - 1:]
    return itens[:cap - 1] + [(f"DEMAIS ({len(resto)})", sum(v for _, v in resto))]


def _empacotar(blocos, lim):
    """Distribui os quadros em colunas de até `lim` linhas (cabeçalho, itens, total e uma linha de espaço
    entre quadros). Um quadro que não cabe no fim de uma coluna continua na próxima como '(cont.)'."""
    colunas, usado = [[]], 0
    for nome, itens, total in blocos:
        resto, cont = list(itens), False
        while True:
            gap = 1 if usado else 0
            livre = lim - usado - gap - 1  # tira o cabeçalho
            if livre >= len(resto) + 1:
                colunas[-1].append((nome + (" (cont.)" if cont else ""), resto, total))
                usado += gap + len(resto) + 2
                break
            if livre >= 2:
                colunas[-1].append((nome + (" (cont.)" if cont else ""), resto[:livre], None))
                resto, cont = resto[livre:], True
            colunas.append([])
            usado = 0
    return colunas


def _slide_blocos(prs, titulo, quadro, total_geral_rot=None):
    """Quadros 'cabeçalho azul + itens + total' em um slide só, em até 3 colunas. Se não couber, os
    menores itens de cada quadro são somados em 'DEMAIS (n)' e as linhas e a fonte diminuem."""
    if not quadro:
        return
    geral = sum(v for its in quadro.values() for _, v in its)
    com_geral = bool(total_geral_rot) and len(quadro) > 1
    base = 488 if com_geral else 515
    alt_min = 10.5
    lim_max = int((base - 100) / alt_min)
    maior = max(len(i) for i in quadro.values())
    escolha = None
    for cap in range(maior, 1, -1):
        blocos = [(b, _resumir(i, cap), sum(v for _, v in i)) for b, i in quadro.items()]
        total_linhas = sum(len(i) + 3 for _, i, _ in blocos) - 1
        for ncol in (1, 2, 3):
            for lim in range(max(4, -(-total_linhas // ncol)), lim_max + 1):
                cols = _empacotar(blocos, lim)
                if len(cols) <= ncol:
                    escolha = cols
                    break
            if escolha:
                break
        if escolha:
            break
    if not escolha:  # quadros demais: aceita linhas mais baixas
        blocos = [(b, _resumir(i, 2), sum(v for _, v in i)) for b, i in quadro.items()]
        escolha = _empacotar(blocos, -(-(sum(len(i) + 3 for _, i, _ in blocos)) // 3) + 3)
    s = _slide_conteudo(prs, titulo)
    ncol = len(escolha)
    wcol = {1: 540, 2: 390, 3: 280}[ncol]
    gap = 14
    x0 = (LARG - (wcol * ncol + gap * (ncol - 1))) / 2
    # com poucos itens as linhas e a fonte crescem para ocupar melhor o slide
    linhas = max(sum(len(i) + 1 + (t is not None) + 1 for _, i, t in coluna) - 1 for coluna in escolha)
    alt = min(26.0, (base - 100 - 4) / linhas)
    tam = min(13 if ncol == 1 else 10, round(alt * 0.6 * 2) / 2)
    w_nome = wcol * 3.2 / 4.4 - 2 * min(4, tam * 0.45)
    y0 = _y_centro(linhas * alt, base=base)
    for ci, coluna in enumerate(escolha):
        y = y0
        for bloco, itens, total in coluna:
            n = len(itens) + 1 + (1 if total is not None else 0)
            tbl = _tabela(s, x0 + ci * (wcol + gap), y, wcol, [3.2, 1.2], n, alt)
            _cab_quadro(tbl.cell(0, 0), _cabe(bloco, w_nome, tam), tam)
            _cab_quadro(tbl.cell(0, 1), "VALOR", tam)
            for r, (item, v) in enumerate(itens, start=1):
                _item_quadro(tbl.cell(r, 0), _cabe(str(item).upper(), w_nome, tam), tam, r % 2 == 0)
                _item_quadro(tbl.cell(r, 1), brl(v), tam, r % 2 == 0)
            if total is not None:
                _total_quadro(tbl.cell(n - 1, 0), "TOTAL", tam)
                _total_quadro(tbl.cell(n - 1, 1), brl(total), tam)
            y += n * alt + alt
    if com_geral:
        w = 540
        tbl = _tabela(s, (LARG - w) / 2, 495, w, [3.2, 1.2], 1, 20)
        _cab_quadro(tbl.cell(0, 0), total_geral_rot, 11)
        _cab_quadro(tbl.cell(0, 1), brl(geral), 11)


def _slide_transf(prs, transf):
    transf = {k: v for k, v in (transf or {}).items() if round(v["enviada"], 2) or round(v["recebida"], 2)}
    if not transf:
        return
    s = _slide_conteudo(prs, "Gerencial Report – TRANSFERÊNCIAS ENTRE CONTAS")
    n = len(transf) + 3
    alt = min(21, 400 / n)
    tam = min(12, round(alt * 0.58 * 2) / 2)
    w = 660
    tbl = _tabela(s, (LARG - w) / 2, _y_centro(n * alt), w, [2.2, 1.9, 1.9], n, alt)
    tbl.cell(0, 0).merge(tbl.cell(0, 2))
    _cab_quadro(tbl.cell(0, 0), "TRANSFERÊNCIAS ENTRE CONTAS DA EMPRESA", tam)
    for c, h in enumerate(["CONTA", "TRANSFERÊNCIA ENVIADA", "TRANSFERÊNCIA RECEBIDA"]):
        _cab_quadro(tbl.cell(1, c), h, tam - 1, fundo=AZUL_MEDIO)
    for r, (conta, v) in enumerate(transf.items(), start=2):
        z = r % 2 == 1
        _item_quadro(tbl.cell(r, 0), conta.upper(), tam, z)
        _item_quadro(tbl.cell(r, 1), brl(v["enviada"]) if v["enviada"] else "R$ -", tam, z)
        _item_quadro(tbl.cell(r, 2), brl(v["recebida"]) if v["recebida"] else "R$ -", tam, z)
    _total_quadro(tbl.cell(n - 1, 0), "TOTAL TRANSFERIDO", tam)
    _total_quadro(tbl.cell(n - 1, 1), brl(sum(v["enviada"] for v in transf.values())), tam)
    _total_quadro(tbl.cell(n - 1, 2), brl(sum(v["recebida"] for v in transf.values())), tam)


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
        _cel(tbl.cell(0, c), h, tam=9, negrito=True, cor=BRANCO, fundo=AZUL, alinh=PP_ALIGN.CENTER, borda=True)
    for r, (k, v) in enumerate(ind.items(), start=1):
        forte = k in ("Entradas", "Receita Líquida", "Resultado Bruto", "Resultado Operacional", "Resultado Líquido")
        fundo = RGBColor(0xE2, 0xEF, 0xDA) if k in ("Resultado Operacional", "Resultado Líquido") else None
        _cel(tbl.cell(r, 0), k, tam=9, negrito=forte, fundo=fundo, borda=True)
        _cel(tbl.cell(r, 1), brl(v["valor"]), tam=9, negrito=forte, fundo=fundo, alinh=PP_ALIGN.RIGHT, borda=True)
        _cel(tbl.cell(r, 2), pct(v["pct"], 1), tam=9, negrito=forte, fundo=fundo, alinh=PP_ALIGN.RIGHT, borda=True)


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


def gerar_pptx(caminho, empresa, ano, mes, base, dre, saldos, cfg, textos):
    prs = Presentation()
    prs.slide_width = _p(LARG)
    prs.slide_height = _p(ALT)
    dre = dict(dre, modo=cfg.get("modo_colunas"))
    _capa(prs, empresa, ano, mes)
    _servicos(prs)
    _slide_dre(prs, empresa, ano, mes, dre)  # a DRE detalhada vai só no Excel
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
