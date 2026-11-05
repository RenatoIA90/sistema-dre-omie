"""Regras exclusivas da empresa MS FESTAS E EVENTOS.

- Plano de contas e layout do Excel iguais ao modelo "FECHAMENTO MS FESTAS 09_2026.xlsx"
  (assets/modelos/ms_festas_fechamento.xlsx): REPORT por conta (ITAÚ/STONE DJ CLUB = FILIAL,
  ITAÚ/STONE MADAME = MATRIZ), SALDOS, FORNECEDORES, FUNCIONÁRIOS, RECEBIMENTOS, TRANSFERÊNCIAS e INSIGHTS.
- Planilha à parte "Comparativos MS FESTAS" no layout do modelo "MS FESTAS REPORT COMPARATIVO 08 E 09_2026.xlsx"
  (assets/modelos/ms_festas_comparativo.xlsx): uma aba por mês e a aba de comparação com a diferença.

A formatação vem das próprias planilhas-modelo: cada célula escrita recebe o estilo da célula equivalente do
modelo (fonte, cores do tema, bordas, formatos de número), então o arquivo sai igual ao modelo.
Os arquivos de modelo guardados no sistema não têm valores, só a formatação.
"""
import os
from collections import OrderedDict, defaultdict
from copy import copy, deepcopy
from datetime import datetime

import openpyxl
from openpyxl.utils import get_column_letter as L

from . import dre as D
from .modelos import completar_grupos
from .plano import FCL, IGNORAR, RESULTADO_LIQUIDO, RESULTADOS, SECOES, norm

ASSETS = os.path.join(os.path.dirname(os.path.dirname(__file__)), "assets", "modelos")
MODELO_FECHAMENTO = os.path.join(ASSETS, "ms_festas_fechamento.xlsx")
MODELO_COMPARATIVO = os.path.join(ASSETS, "ms_festas_comparativo.xlsx")

NOME_MODELO = "MS Festas e Eventos (modelo FECHAMENTO MS FESTAS) (pronto)"
CNPJ_RAIZ = "35094562"
TITULO_TOTAL = "TOTAL MS FESTAS"
MESES_ABREV = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]

# Plano do modelo, na ordem em que aparece no REPORT: código, nome, seção, relatório, contas (código, nome).
PLANO = [
    ("010100", "RECEIMENTOS", "ENTRADAS", "recebimentos", [
        ("010101", "Cartão Crédito - Rede Amex"), ("010102", "Cartão Crédito - Rede Elo"),
        ("010103", "Cartão Crédito - Rede Master"), ("010104", "Cartão Crédito - Rede Visa"),
        ("010105", "Cartão Debito - Rede Amex"), ("010106", "Cartão Debito - Rede Elo"),
        ("010107", "Cartão Debito - Rede Master"), ("010108", "Cartão Debito - Rede Visa"),
        ("010109", "PIX")]),
    ("020100", "IMPOSTOS", "DEDUCOES", "", [
        ("020101", "(-)DAS SIMPLES NACIONAL"), ("020102", "(-) ISS"), ("020103", "(-) PIS"),
        ("020104", "(-) COFINS"), ("020105", "(-) OUTROS IMPOSTOS")]),
    ("040100", "PRESTADORES DE SERVIÇOS", "CUSTOS", "funcionarios", [
        ("040101", "Funcionários"), ("040102", "Dj Musicos"), ("040103", "Cachê Show (Performances, Bandas)"),
        ("040104", "Seguranças"), ("040105", "Coletor de Lixo"), ("040106", "Serviço de Manutenção "),
        ("040107", "Produção"), ("040108", "Tecnico Som"), ("040109", "Transporte")]),
    ("040400", "LOJINHA", "CUSTOS", "", [("040112", "Ativo Circulante - Lojinha")]),
    ("040300", "FORNECEDORES", "CUSTOS", "fornecedores", [
        ("040301", "Fornecedores de Bebidas"), ("040302", "Gelo"), ("040303", "Fornecedores Alimenticios"),
        ("040304", "Gráfica"), ("040305", "Decoração"), ("040305", "Móveis e Utensílios (Cadeiras)")]),
    ("040200", "LOCAÇÃO/AQUISIÇÃO E MANUTENÇÃO", "CUSTOS", "", [
        ("040201", "Locação/Aquisição de Equipamentos"), ("040202", "Manutenção de Equipamentos"),
        ("040203", "Máquinas e Equipamentos (Acustica Gerador)"), ("040204", "Locação de Espaço Evento")]),
    ("060100", "CONTABILIDADE", "DESPESAS", "", [("060102", "Savian Contabilidade")]),
    ("060200", "SISTEMA", "DESPESAS", "", [("060201", "Ponto de Venda")]),
    ("060300", "TECNOLOGIA DA INFORMAÇÃO", "DESPESAS", "", [
        ("060301", "SOFTWARE"), ("060302", "HARDWARE"), ("060303", "ASSESSORIA EM T.I")]),
    ("060400", "ADMINISTRATIVO", "DESPESAS", "", [
        ("060401", "Água e Esgoto"), ("060402", "Energia Eletrica"), ("060403", "Marketing / Comissões"),
        ("060404", "Limpeza"), ("060405", "Telefonia/Internet"), ("060406", "ECAD Direitos de Imagem"),
        ("060407", "Multas de Transito")]),
    ("060500", "ALUGUEL", "DESPESAS", "", [
        ("060501", "Locação das Casas"), ("060502", "Condominios"), ("060502", "IPTU")]),
    ("060600", "JURÍDICO", "DESPESAS", "", [("060601", "Advogados")]),
    ("060700", "SÓCIOS", "DESPESAS", "", [
        ("060701", "Pro-Labore"), ("060701", "Socios (Reembolsos, Assitencia Médica)")]),
    ("080100", "RECEITAS FINANCEIRAS", "FINANCEIRO", "", [("080101", "Rendimento Bancário")]),
    ("080200", "DESPESAS FINANCEIRAS", "FINANCEIRO", "", [
        ("080201", "Operadora de Cartão"), ("080202", "Tarifa Bancária")]),
    ("090200", "DESPESAS NÃO OPERACIONAIS", "NAO_OPERACIONAL", "", [("090201", "DESPESAS ITAU")]),
]
# rótulos dos quadros de FUNCIONÁRIOS no modelo
ROTULOS_FUNC = {"funcionarios": "FUNCIONÁRIO", "coletor de lixo": "COLETOR DE LIXO", "dj musicos": "DJ",
                "cache show (performances, bandas)": "CACHE PARA BANDA", "servico de manutencao": "MANUTENÇÃO"}
# contas correntes do OMIE -> coluna do relatório e agrupamento (como no modelo)
COLUNAS = [  # (palavras que precisam aparecer no nome da conta, nome da coluna, agrupamento)
    (("itau", "dj"), "CONTA ITAU DJ CLUB", "FILIAL"),
    (("stone", "dj"), "CONTA STONE DJCLUB", "FILIAL"),
    (("itau", "madame"), "CONTA ITAU MADAME", "MATRIZ"),
    (("stone", "madame"), "CONTA STONE MADAME", "MATRIZ"),
]


def _rotulo(cod, nome):
    return f"{cod}-{nome}"


_CONTA_MODELO = {(g, norm(n)): _rotulo(c, n) for g, _, _, _, contas in PLANO for c, n in contas}
_GRUPOS_MODELO = [g for g, *_ in PLANO]


def eh_ms_festas(empresa):
    if not empresa:
        return False
    cnpjs = "".join(ch for ch in str(empresa.get("cnpj") or "") if ch.isdigit() or ch == "\n")
    if any(c.startswith(CNPJ_RAIZ) for c in cnpjs.split("\n")):
        return True
    return "ms festas" in norm(empresa.get("nome"))


def modelo():
    """Modelo de DRE (grupos + contas) no formato de core.modelos."""
    grupos = [[g, n, s, r] for g, n, s, r, _ in PLANO]
    contas = [[n.strip(), g] for g, _, _, _, cs in PLANO for _, n in cs]
    return {"grupos": grupos, "contas": contas}


def precisa_aplicar_modelo(cfg):
    """Aplica o modelo MS FESTAS uma única vez; se depois o usuário trocar o modelo, a escolha dele vale."""
    return not cfg.get("ms_festas_modelo_aplicado")


def aplicar_modelo(cfg, mapa):
    m = modelo()
    cfg["grupos"] = completar_grupos(m["grupos"])
    cfg["contas_modelo"] = m["contas"]
    cfg["modelo_dre"] = NOME_MODELO
    cfg["grupos_criados"] = []
    cfg["ms_festas_modelo_aplicado"] = True
    D.reclassificar(mapa, cfg)


def preparar_colunas(cfg, nomes_contas):
    """Colunas do REPORT como no modelo (apelidos, agrupamentos FILIAL/MATRIZ e ordem), sem sobrescrever o que já
    foi configurado para a empresa. Altera cfg em memória."""
    apel = dict(cfg.get("apelidos_conta") or {})
    grp = dict(cfg.get("grupos_colunas") or {})
    for nome in nomes_contas:
        t = norm(nome)
        for palavras, coluna, agrupamento in COLUNAS:
            if all(p in t for p in palavras):
                apel.setdefault(nome, coluna)
                grp.setdefault(apel[nome], agrupamento)
                break
    cfg["apelidos_conta"] = apel
    cfg["grupos_colunas"] = grp
    if not cfg.get("ordem_colunas"):
        cfg["ordem_colunas"] = [c for _, c, _ in COLUNAS]
    return cfg


def nome_curto(conta, cfg):
    """Nome da conta corrente nas abas SALDOS/TRANSFERÊNCIAS (ex.: 'ITAU DJ CLUB')."""
    apel = (cfg.get("apelidos_conta") or {}).get(conta, conta)
    return apel[6:] if apel.upper().startswith("CONTA ") else apel


# ------------------------------------------------------------------ base e DRE
def codificar(base, grupos=None):
    """conta_fechamento com os códigos do modelo (ex.: 040101-Funcionários). Contas fora do modelo ganham o próximo
    código livre do grupo, em ordem alfabética, depois das contas do modelo."""
    extras = defaultdict(set)
    for b in base:
        if b["grupo"] != IGNORAR and (b["grupo"], norm(b["conta"])) not in _CONTA_MODELO:
            extras[b["grupo"]].add(b["conta"])
    codigos = {}
    for g, contas in extras.items():
        usados = [int(lbl[4:6]) for (gg, _), lbl in _CONTA_MODELO.items() if gg == g]
        n = max(usados, default=0)
        for conta in sorted(contas, key=norm):
            n += 1
            codigos[(g, conta)] = f"{g[:4]}{n:02d}-{conta}"
    for b in base:
        if b["grupo"] == IGNORAR:
            b["conta_fechamento"] = b["conta"]
        else:
            b["conta_fechamento"] = _CONTA_MODELO.get((b["grupo"], norm(b["conta"]))) or codigos[(b["grupo"], b["conta"])]


def montar_base(lancs, cfg, mapa):
    return D.montar_base(lancs, cfg, mapa, codificar=codificar)


def montar_dre(base, cfg, colunas=None, contas_extra=()):
    """DRE no layout fixo do modelo: todos os grupos e contas do modelo aparecem (zerados se não houver
    movimento) e contas/grupos novos entram no fim do grupo/seção. Mesmo formato de core.dre.montar_dre."""
    grupos_cfg = D.grupos_da_empresa(cfg)
    cols = list(colunas) if colunas is not None else D.colunas(base, cfg)
    soma = defaultdict(lambda: defaultdict(float))
    conta_grupo = {}
    for b in base:
        if b["grupo"] == IGNORAR:
            continue
        soma[b["conta_fechamento"]][b["coluna"]] += b["valor"]
        conta_grupo[b["conta_fechamento"]] = b["grupo"]
    for lbl, g in contas_extra:
        conta_grupo.setdefault(lbl, g)

    nomes_cfg = {g[0]: g for g in grupos_cfg}
    linhas, tot_secao = [], {}
    for sec, rotulo in SECOES:
        lin_sec = {"nivel": 1, "descricao": rotulo, "secao": sec, "grupo": "", "valores": {c: 0.0 for c in cols},
                   "tipo": "secao"}
        linhas.append(lin_sec)
        do_modelo = [(g, n, cs) for g, n, s, _, cs in PLANO if s == sec]
        com_contas = set(conta_grupo.values())
        outros = sorted(g for g, info in nomes_cfg.items()
                        if info[2] == sec and g not in _GRUPOS_MODELO and g in com_contas)
        for g, nome, contas_modelo in do_modelo + [(g, nomes_cfg[g][1], []) for g in outros]:
            lin_g = {"nivel": 2, "descricao": _rotulo(g, nome), "secao": sec, "grupo": g,
                     "valores": {c: 0.0 for c in cols}, "tipo": "grupo"}
            linhas.append(lin_g)
            rots = [_rotulo(c, n) for c, n in contas_modelo]
            rots += sorted((cf for cf, gg in conta_grupo.items() if gg == g and cf not in rots), key=norm)
            for cf in rots:
                vals = {c: round(soma[cf].get(c, 0.0), 2) for c in cols}
                linhas.append({"nivel": 3, "descricao": cf, "secao": sec, "grupo": g, "valores": vals,
                               "tipo": "conta"})
                for c in cols:
                    lin_g["valores"][c] += vals[c]
                    lin_sec["valores"][c] += vals[c]
        tot_secao[sec] = lin_sec["valores"]
        if sec in RESULTADOS:
            rot, secs = RESULTADOS[sec]
            linhas.append({"nivel": 1, "descricao": rot, "secao": sec, "grupo": "", "tipo": "resultado",
                           "valores": {c: sum(tot_secao.get(s, {}).get(c, 0.0) for s in secs) for c in cols}})
    liq = {c: sum(tot_secao[s][c] for s, _ in SECOES) for c in cols}
    linhas.append({"nivel": 1, "descricao": RESULTADO_LIQUIDO, "secao": "", "grupo": "", "tipo": "resultado",
                   "valores": dict(liq)})
    linhas.append({"nivel": 1, "descricao": FCL, "secao": "", "grupo": "", "tipo": "resultado", "valores": dict(liq)})
    for l in linhas:
        l["valores"] = {c: round(v, 2) for c, v in l["valores"].items()}
        l["total"] = round(sum(l["valores"].values()), 2)
    agrup = OrderedDict()
    for c in cols:
        g = (cfg.get("grupos_colunas") or {}).get(c)
        if g:
            agrup.setdefault(g, []).append(c)
    return {"colunas": cols, "agrupamentos": agrup, "linhas": linhas}


def montar_meses(meses, cfg, mapa):
    """meses: [(ano, mes, lancs)] em ordem. Monta base e DRE de cada mês com as MESMAS linhas e colunas
    (união dos meses), para o comparativo. Devolve [(ano, mes, base, dre)] e as categorias novas do mapa."""
    bases, novas = [], []
    for ano, mes, lancs in meses:
        c = deepcopy(cfg)
        c["rateios_mes"] = (cfg.get("rateios") or {}).get(f"{ano}-{mes:02d}", [])
        novas += D.completar_mapeamento(mapa, lancs, c)
        if c.get("grupos") != cfg.get("grupos"):  # grupo novo criado por categoria fora do modelo
            cfg["grupos"], cfg["grupos_criados"] = c["grupos"], c.get("grupos_criados", [])
        bases.append((ano, mes, montar_base(lancs, c, mapa)))
    todas = [b for _, _, base in bases for b in base]
    codificar(todas)
    uniao = montar_dre(todas, cfg)
    extra = [(l["descricao"], l["grupo"]) for l in uniao["linhas"] if l["tipo"] == "conta"]
    return [(a, m, base, montar_dre(base, cfg, colunas=uniao["colunas"], contas_extra=extra))
            for a, m, base in bases], novas


# ------------------------------------------------------------------ estilos copiados do modelo
def _st(ws, coord):
    return copy(ws[coord]._style)


def _put(ws, r, c, valor, estilo):
    cell = ws.cell(r, c)
    if valor is not None:
        cell.value = valor
    cell._style = copy(estilo)
    return cell


def _limpar(ws):
    for rng in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(rng))
    ws._cells.clear()
    for k in list(ws.row_dimensions.keys()):
        del ws.row_dimensions[k]


def _estilos_report(ws):
    p = {k: _st(ws, c) for k, c in {
        "titulo": "B2", "desc": "B3", "A2": "A2",
        ("h", "col", 0): "C2", ("h", "col", 1): "D2", ("h", "col", 2): "E2",
        ("h", "sub", 0): "I2", ("h", "sub", 1): "J2", ("h", "sub", 2): "K2",
        ("h", "tot", 0): "U2", ("h", "tot", 1): "V2", ("h", "tot", 2): "W2",
        ("d", "col", 0): "C3", ("d", "col", 1): "D3", ("d", "col", 2): "E3",
        ("d", "sub", 0): "I3", ("d", "sub", 1): "J3", ("d", "sub", 2): "K3",
        ("d", "tot", 0): "U3", ("d", "tot", 1): "V3", ("d", "tot", 2): "W3",
    }.items()}
    for tipo, r in (("secao", 4), ("grupo", 5), ("conta", 6), ("resultado", 22)):
        p[(tipo, "A")] = _st(ws, f"A{r}")
        p[(tipo, "B")] = _st(ws, f"B{r}")
        for bloco, letras in (("col", "CDE"), ("sub", "IJK"), ("tot", "UVW")):
            for i, x in enumerate(letras):
                p[(tipo, bloco, i)] = _st(ws, f"{x}{r}")
    p["larguras"] = {k: v.width for k, v in ws.column_dimensions.items()}
    return p


def _blocos(dre):
    cols, agrup = dre["colunas"], dre["agrupamentos"]
    seq, usados = [], set()
    for g, membros in agrup.items():
        seq += [("col", m) for m in membros] + [("sub", g)]
        usados.update(membros)
    seq += [("col", c) for c in cols if c not in usados]
    return seq + [("tot", TITULO_TOTAL)]


_LAYOUT_MODELO = ["col", "col", "sub", "col", "col", "sub", "tot"]


def _escrever_report(ws, p, dre, ano, mes, valor_conta):
    """REPORT no layout do modelo. valor_conta(linha, nome_coluna, row) devolve o valor ou fórmula da conta."""
    _limpar(ws)
    blocos = _blocos(dre)
    mesmo_layout = [t for t, _ in blocos] == _LAYOUT_MODELO
    if not mesmo_layout:
        for k in list(ws.column_dimensions.keys()):
            if k not in ("A", "B"):
                del ws.column_dimensions[k]
    _put(ws, 2, 1, None, p["A2"])
    _put(ws, 2, 2, "CashFlowReport - Por Conta", p["titulo"])
    _put(ws, 3, 2, "Descrição da Conta", p["desc"])
    pos, c = {}, 3
    for i, (tipo, nome) in enumerate(blocos):
        pos[(tipo, nome)] = c
        rot = f"TOTAL {nome}".upper() if tipo == "sub" else nome
        _put(ws, 2, c, rot, p[("h", tipo, 0)])
        _put(ws, 2, c + 1, None, p[("h", tipo, 1)])
        ws.merge_cells(start_row=2, start_column=c, end_row=2, end_column=c + 1)
        _put(ws, 3, c, datetime(ano, mes, 1), p[("d", tipo, 0)])
        _put(ws, 3, c + 1, "AV", p[("d", tipo, 1)])
        if tipo != "tot":
            _put(ws, 2, c + 2, None, p[("h", tipo, 2)])
            _put(ws, 3, c + 2, None, p[("d", tipo, 2)])
        if not mesmo_layout:
            ws.column_dimensions[L(c)].width = 11.71
            ws.column_dimensions[L(c + 2)].width = 1.14
        c += 3

    linhas = dre["linhas"]
    r0 = 4
    row = {i: r0 + i for i in range(len(linhas))}
    sec_row, grupos_sec, contas_g, res_row = {}, defaultdict(list), defaultdict(list), {}
    g_atual = None
    for i, l in enumerate(linhas):
        if l["tipo"] == "secao":
            sec_row[l["secao"]] = row[i]
        elif l["tipo"] == "grupo":
            grupos_sec[l["secao"]].append(row[i])
            g_atual = row[i]
        elif l["tipo"] == "conta":
            contas_g[g_atual].append(row[i])
        else:
            res_row[l["descricao"]] = row[i]
    ent = sec_row.get("ENTRADAS", r0)
    rot_rliq, rot_bruto, rot_fco = (RESULTADOS[s][0] for s in ("DEDUCOES", "CUSTOS", "DESPESAS"))
    cadeia = {  # resultados como no modelo: cada um soma o anterior com a seção de cima
        rot_rliq: lambda X: [f"{X}{sec_row[s]}" for s in ("ENTRADAS", "DEDUCOES")],
        rot_bruto: lambda X: [f"{X}{res_row[rot_rliq]}", f"{X}{sec_row['CUSTOS']}"],
        rot_fco: lambda X: [f"{X}{res_row[rot_bruto]}", f"{X}{sec_row['DESPESAS']}"],
        RESULTADO_LIQUIDO: lambda X: [f"{X}{res_row[rot_fco]}", f"{X}{sec_row['FINANCEIRO']}",
                                      f"{X}{sec_row['NAO_OPERACIONAL']}"],
    }

    def soma(refs):
        if not refs:
            return 0
        return f"={refs[0]}" if len(refs) == 1 else f"=SUM({','.join(refs)})"

    agrup = dre["agrupamentos"]
    for i, l in enumerate(linhas):
        r, t = row[i], l["tipo"]
        _put(ws, r, 1, l["nivel"], p[(t, "A")])
        _put(ws, r, 2, l["descricao"], p[(t, "B")])
        if t in ("secao", "resultado"):
            ws.row_dimensions[r].height = 15.75
        for (tipo, nome), cp in pos.items():
            X = L(cp)
            if tipo == "col":
                if t == "conta":
                    f = valor_conta(l, nome, r)
                elif t == "grupo":
                    rs = contas_g[r]
                    f = f"=SUM({X}{rs[0]}:{X}{rs[-1]})" if len(rs) > 1 else (f"={X}{rs[0]}" if rs else 0)
                elif t == "secao":
                    f = soma([f"{X}{x}" for x in grupos_sec.get(l["secao"], [])])
                elif l["descricao"] == FCL:
                    f = f"={X}{res_row[RESULTADO_LIQUIDO]}"
                else:
                    f = soma(cadeia[l["descricao"]](X))
            elif tipo == "sub":
                f = soma([f"{L(pos[('col', m)])}{r}" for m in agrup[nome]])
            else:
                partes = [(tp, nm) for tp, nm in pos if tp == "sub"]
                partes += [("col", c) for c in dre["colunas"] if not any(c in m for m in agrup.values())]
                f = soma([f"{L(pos[k])}{r}" for k in partes])
            _put(ws, r, cp, f, p[(t, tipo, 0)])
            _put(ws, r, cp + 1, f"=IFERROR({X}{r}/${X}${ent},0)", p[(t, tipo, 1)])
            if tipo != "tot":
                _put(ws, r, cp + 2, None, p[(t, tipo, 2)])
    return {"pos": pos, "linhas": {l["descricao"]: row[i] for i, l in enumerate(linhas)},
            "col_total": pos[("tot", TITULO_TOTAL)], "fim": r0 + len(linhas) - 1}


# ------------------------------------------------------------------ Excel do fechamento
def gerar_excel(caminho, empresa, ano, mes, base, dre, saldos, cfg, textos=None):
    wb = openpyxl.load_workbook(MODELO_FECHAMENTO)
    ws_rep, ws_sal, ws_forn, ws_func, ws_rec, ws_tr, ws_ins = wb.worksheets
    ws_rep.title = f"REPORT {mes:02d} {ano} - POR CONTA"
    ws_ins.title = f"INSIGHTS {mes:02d} {ano}"
    rep = ws_rep.title

    from .excel import _aba_base
    n_base = len(base) + 1
    rng_v = f"'BASE CLASSIFICADA'!$K$2:$K${n_base}"
    rng_c = f"'BASE CLASSIFICADA'!$G$2:$G${n_base}"
    rng_col = f"'BASE CLASSIFICADA'!$I$2:$I${n_base}"

    def sumifs(l, coluna, r):
        return f"=SUMIFS({rng_v},{rng_c},$B{r},{rng_col},\"{coluna}\")"

    info = _escrever_report(ws_rep, _estilos_report(ws_rep), dre, ano, mes, sumifs)
    T = L(info["col_total"])
    lin = info["linhas"]

    _aba_saldos(ws_sal, saldos, cfg, ano, mes)
    _aba_fornecedores(ws_forn, base, cfg)
    _aba_funcionarios(ws_func, base, cfg)
    _aba_recebimentos(ws_rec, dre, rep, T, lin, ano, mes)
    _aba_transferencias(ws_tr, D.transferencias(base), cfg)
    _aba_insights(ws_ins, empresa, ano, mes, dre, rep, T, lin, saldos, textos)
    _aba_base(wb, base)
    wb.active = 0
    wb.save(caminho)
    return caminho


def _aba_saldos(ws, saldos, cfg, ano, mes):
    import calendar
    p = {k: _st(ws, k) for k in ("B2", "C2", "D2", "B4", "C4", "D4", "D7")}
    _limpar(ws)
    pm, pa = (12, ano - 1) if mes == 1 else (mes - 1, ano)
    for c, (txt, k) in enumerate(((f"EMPRESA", "B2"),
                                  (f"SALDO INICIAL {calendar.monthrange(pa, pm)[1]:02d}/{pm:02d}/{pa}", "C2"),
                                  (f"SALDO FINAL {calendar.monthrange(ano, mes)[1]:02d}/{mes:02d}/{ano}", "D2")), 2):
        _put(ws, 2, c, txt, p[k])
    r = 3
    for s in saldos or []:
        _put(ws, r, 2, nome_curto(s["conta"], cfg).upper(), p["B4"])
        _put(ws, r, 3, round(s["saldo_inicial"], 2), p["C4"])
        _put(ws, r, 4, round(s["saldo_final"], 2), p["D4"])
        r += 1
    _put(ws, r, 4, f"=SUM(D3:D{r - 1})" if r > 3 else 0, p["D7"])
    ws._ms_total = r


def _pagamentos_por_conta(base, grupos_rel):
    """{conta: [(favorecido, valor positivo)]} dos pagamentos do mês nos grupos dados."""
    acc = defaultdict(lambda: defaultdict(float))
    ordem = {}
    for b in base:
        if b["grupo"] in grupos_rel and b["tipo"] == "Lançamento" and b["valor"] < 0:
            acc[b["conta_fechamento"]][b["favorecido"] or "(sem nome)"] -= b["valor"]
            ordem.setdefault(b["conta_fechamento"], b["conta_fechamento"])
    out = OrderedDict()
    for conta in sorted(acc):
        itens = sorted(((f, round(v, 2)) for f, v in acc[conta].items() if round(v, 2)), key=lambda x: -x[1])
        if itens:
            out[conta] = itens
    return out


def _grupos_rel(cfg, rel):
    return {g[0] for g in D.grupos_da_empresa(cfg) if g[3] == rel}


def _sem_codigo(lbl):
    return lbl.split("-", 1)[1].strip() if "-" in lbl[:7] else lbl.strip()


def _aba_fornecedores(ws, base, cfg):
    p = {k: _st(ws, k) for k in ("B2", "C2", "B3", "C3", "B16", "C16")}
    _limpar(ws)
    quadros = _pagamentos_por_conta(base, _grupos_rel(cfg, "fornecedores"))
    r, faixas = 2, []
    for conta, itens in quadros.items():
        _put(ws, r, 2, _sem_codigo(conta).upper(), p["B2"])
        _put(ws, r, 3, "VALOR", p["C2"])
        r += 1
        ini = r
        for fav, v in itens:
            _put(ws, r, 2, fav, p["B3"])
            _put(ws, r, 3, v, p["C3"])
            r += 1
        faixas.append(f"C{ini}:C{r - 1}")
        r += 1
    _put(ws, r, 2, "TOTAL FORNECEDORES", p["B16"])
    _put(ws, r, 3, f"=SUM({','.join(faixas)})" if faixas else 0, p["C16"])


def _aba_funcionarios(ws, base, cfg):
    p = {k: _st(ws, k) for k in ("B2", "C2", "B3", "C3")}
    _limpar(ws)
    quadros = _pagamentos_por_conta(base, _grupos_rel(cfg, "funcionarios"))
    if not quadros:
        return
    blocos = [(ROTULOS_FUNC.get(norm(_sem_codigo(c)), _sem_codigo(c).upper()), itens) for c, itens in quadros.items()]
    # o maior quadro (funcionários) na primeira coluna; os outros vão para a coluna mais curta, como no modelo
    blocos.sort(key=lambda b: -len(b[1]))
    colunas = [[blocos[0]], [], []]
    altura = [len(blocos[0][1]) + 1, 0, 0]
    for b in blocos[1:]:
        i = 1 if altura[1] <= altura[2] else 2
        colunas[i].append(b)
        altura[i] += len(b[1]) + 1 + (1 if altura[i] else 0)
    for i, coluna in enumerate(colunas):
        cn = 2 + 3 * i  # B, E, H
        r = 2
        for rot, itens in coluna:
            _put(ws, r, cn, rot, p["B2"])
            _put(ws, r, cn + 1, "VALOR", p["C2"])
            r += 1
            for fav, v in itens:
                _put(ws, r, cn, fav, p["B3"])
                _put(ws, r, cn + 1, v, p["C3"])
                r += 1
            r += 1


def _aba_recebimentos(ws, dre, rep, T, lin, ano, mes):
    p = {k: _st(ws, k) for k in ("B2", "C2", "B3", "C3", "B7", "C7", "B19", "C19")}
    _limpar(ws)
    grupos = {}
    for l in dre["linhas"]:
        if l["tipo"] == "conta" and l["secao"] == "ENTRADAS":
            t = norm(l["descricao"])
            chave = ("CRÉDITO", "Total Crédito") if "credito" in t else ("DEBITO", "Total Débito") if "debito" in t \
                else ("PIX", "Total PIX") if "pix" in t else ("OUTROS", "Total Outros")
            if chave[0] == "OUTROS" and not l["total"]:
                continue
            grupos.setdefault(chave, []).append(l)
    r, totais = 2, []
    for (bloco, rot_total), itens in grupos.items():
        _put(ws, r, 2, f"RECEBIMENTOS {bloco}", p["B2"])
        _put(ws, r, 3, None, p["C2"])
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        r += 1
        ini = r
        for l in itens:
            _put(ws, r, 2, _sem_codigo(l["descricao"]), p["B3"])
            _put(ws, r, 3, f"='{rep}'!{T}{lin[l['descricao']]}", p["C3"])
            r += 1
        if bloco != "PIX" or len(itens) > 1:
            _put(ws, r, 2, rot_total, p["B7"])
            _put(ws, r, 3, f"=SUM(C{ini}:C{r - 1})", p["C7"])
            totais.append(f"C{r}")
            r += 2
        else:
            totais.append(f"C{ini}")
            r += 1
    _put(ws, r, 2, f"TOTAL RECEBIMENTO {mes:02d}/{ano}", p["B19"])
    _put(ws, r, 3, f"=SUM({','.join(totais)})" if totais else 0, p["C19"])


def _aba_transferencias(ws, transf, cfg):
    p = {k: _st(ws, k) for k in ("B2", "C2", "D2", "B3", "C3", "D3", "B4", "C4", "D4", "B7", "C7", "D7",
                                 "B8", "C8", "D8")}
    _limpar(ws)
    _put(ws, 2, 2, "TRANSFERENCIAS ENTRE CONTAS DA EMPRESA", p["B2"])
    _put(ws, 2, 3, None, p["C2"])
    _put(ws, 2, 4, None, p["D2"])
    ws.merge_cells("B2:D2")
    for c, (txt, k) in enumerate((("EMPRESA", "B3"), ("TRANSFERENCIA ENVIADA", "C3"),
                                  ("TRANSFERENCIA RECEBIDA", "D3")), 2):
        _put(ws, 3, c, txt, p[k])
    itens = list(transf.items())
    r = 4
    for i, (conta, v) in enumerate(itens):
        ult = "7" if i == len(itens) - 1 else "4"
        _put(ws, r, 2, nome_curto(conta, cfg).upper(), p["B" + ult])
        _put(ws, r, 3, round(v["enviada"], 2), p["C" + ult])
        _put(ws, r, 4, round(v["recebida"], 2), p["D" + ult])
        r += 1
    _put(ws, r, 2, "TOTAL TRANSFERIDO ", p["B8"])
    _put(ws, r, 3, f"=SUM(C4:C{r - 1})" if itens else 0, p["C8"])
    _put(ws, r, 4, f"=SUM(D4:D{r - 1})" if itens else 0, p["D8"])
    ws._ms_total = r


_EMOJI = [("recebimento", "💰"), ("imposto", "🧾"), ("resultado bruto", "📈"), ("custo", "🏗️"),
          ("despesas operacionais", "🏢"), ("resultado operacional", "🟢"), ("resultado financeiro", "💳"),
          ("nao operacional", "⚠️"), ("resultado liquido", "🟢"), ("caixa", "🏦")]


def _com_emoji(txt):
    t = norm(txt)
    for chave, em in _EMOJI:
        if t.startswith(chave):
            return f"{em} {txt}"
    return f"• {txt}"


def _aba_insights(ws, empresa, ano, mes, dre, rep, T, lin, saldos, textos):
    from .analise import MESES, gerar_textos
    p = {k: _st(ws, k) for k in ("A1", "B2", "B3", "B5", "B6", "B16", "B17", "C17", "D17", "B18", "C18", "D18",
                                 "B29", "B30", "C30", "B34", "C34", "B40", "B55", "F2")}
    _limpar(ws)
    textos = textos or gerar_textos(empresa, ano, mes, [], dre, saldos or [], {})
    ref = lambda rot: f"'{rep}'!{T}{lin[rot]}"
    competencia = f"{MESES[mes - 1]}/{ano}"
    linhas_fech = [x.strip() for x in str(textos.get("fechamento", "")).split("\n") if x.strip()]
    pos = [x.strip().lstrip("•").strip() for x in str(textos.get("positivos", "")).split("\n") if x.strip()]
    ate = [x.strip().lstrip("•").strip() for x in str(textos.get("atencao", "")).split("\n") if x.strip()]
    n_saldos = len(saldos or [])
    ultima = 2 + 3 + len(linhas_fech) + 2 + 12 + 2 + 2 + n_saldos + 1 + 2 + 2 + len(pos) + 2 + len(ate) + 4
    for r in range(1, max(60, ultima) + 1):  # fundo branco do modelo
        for c in range(1, 27):
            _put(ws, r, c, None, p["A1"] if c == 1 else p["F2"])
    alturas = {}

    def titulo(r, txt, k="B5"):
        _put(ws, r, 2, txt, p[k])
        alturas[r] = 17.25

    _put(ws, 2, 2, "📊 " + str(textos.get("titulo") or f"Fechamento Financeiro — {competencia} | MS Festas"), p["B2"])
    alturas[2] = 21.0
    _put(ws, 3, 2, "Valores consolidados das quatro contas (coluna TOTAL MS FESTAS do relatório)", p["B3"])
    titulo(5, f"📊 Fechamento — {competencia}")
    r = 6
    for x in linhas_fech:
        _put(ws, r, 2, _com_emoji(x), p["B6"])
        r += 1
    r += 1
    titulo(r, "📋 Indicadores do mês")
    r += 1
    for c, (txt, k) in enumerate((("Indicador", "B16"), (competencia, "B16"), ("% dos recebimentos", "B16")), 2):
        _put(ws, r, c, txt, p[k])
    r += 1
    ind = [("💰 Entradas", "(+) ENTRADAS", True), ("🧾 Deduções (impostos)", "(-) DEDUÇÕES S/ RECEITA BRUTA", False),
           ("💵 Receita Líquida", RESULTADOS["DEDUCOES"][0], True), ("🏗️ Custos Operacionais", "(-) CUSTOS OPERACIONAIS", False),
           ("📈 Resultado Bruto", RESULTADOS["CUSTOS"][0], True), ("🏢 Despesas Operacionais", "(-) DESPESAS OPERACIONAIS", False),
           ("🟢 Resultado Operacional", RESULTADOS["DESPESAS"][0], True),
           ("💳 Resultado Financeiro", "(+/-) RESULTADO FINANCEIRO", False),
           ("⚠️ Não Operacional", "(+/-) RESULTADO NÃO OPERACIONAL", False), ("🟢 Resultado Líquido", RESULTADO_LIQUIDO, True)]
    r_ent = r
    for nome, rot, forte in ind:
        s = "17" if forte else "18"
        _put(ws, r, 2, nome, p["B" + s])
        _put(ws, r, 3, f"={ref(rot)}", p["C" + s])
        _put(ws, r, 4, f"=IFERROR(C{r}/$C${r_ent},0)", p["D" + s])
        r += 1
    r += 1
    titulo(r, "🏦 Saldos bancários")
    r += 1
    import calendar
    pm, pa = (12, ano - 1) if mes == 1 else (mes - 1, ano)
    for c, txt in enumerate(("Conta", f"Saldo inicial {calendar.monthrange(pa, pm)[1]:02d}/{pm:02d}",
                             f"Saldo final {calendar.monthrange(ano, mes)[1]:02d}/{mes:02d}", "Variação"), 2):
        _put(ws, r, c, txt, p["B29"])
    r += 1
    ini = r
    for i in range(n_saldos):
        _put(ws, r, 2, f"=SALDOS!B{3 + i}", p["B30"])
        _put(ws, r, 3, f"=SALDOS!C{3 + i}", p["C30"])
        _put(ws, r, 4, f"=SALDOS!D{3 + i}", p["C30"])
        _put(ws, r, 5, f"=D{r}-C{r}", p["C30"])
        r += 1
    _put(ws, r, 2, "Total", p["B34"])
    for c in (3, 4):
        _put(ws, r, c, f"=SUM({L(c)}{ini}:{L(c)}{r - 1})" if n_saldos else 0, p["C34"])
    _put(ws, r, 5, f"=D{r}-C{r}", p["C34"])
    r += 2
    titulo(r, "✅ Conclusão")
    r += 2
    titulo(r, "⭐ Principais pontos positivos:")
    r += 1
    for x in pos:
        _put(ws, r, 2, f"•  {x}", p["B40"])
        r += 1
    r += 1
    titulo(r, "🔎 Principais pontos de atenção:")
    r += 1
    for x in ate:
        _put(ws, r, 2, f"•  {x}", p["B40"])
        r += 1
    r += 1
    if textos.get("conclusao"):
        _put(ws, r, 2, textos["conclusao"], p["B55"])
    for k, h in alturas.items():
        ws.row_dimensions[k].height = h


# ------------------------------------------------------------------ planilha Comparativos MS FESTAS
def _estilos_comparativo(ws):
    p = {"titulo": _st(ws, "B2"), "desc": _st(ws, "B3"), "A2": _st(ws, "A2"),
         ("h", 0): _st(ws, "C2"), ("h", 1): _st(ws, "D2"), ("h", 2): _st(ws, "E2"), ("h", "esp"): _st(ws, "F2"),
         ("d", 0): _st(ws, "C3"), ("d", 1): _st(ws, "D3"), ("d", 2): _st(ws, "E3"), ("d", "esp"): _st(ws, "F3"),
         ("chave", "h"): _st(ws, "AE3"), ("sinal", "h"): _st(ws, "AF3"), ("esp2", "h"): _st(ws, "AD3")}
    for tipo, r in (("secao", 4), ("grupo", 5), ("conta", 6), ("resultado", 22)):
        p[(tipo, "A")] = _st(ws, f"A{r}")
        p[(tipo, "B")] = _st(ws, f"B{r}")
        for i, x in enumerate("CDE"):
            p[(tipo, i)] = _st(ws, f"{x}{r}")
        p[(tipo, "esp")] = _st(ws, f"F{r}")
        p[(tipo, "chave")] = _st(ws, f"AE{r}")
        p[(tipo, "sinal")] = _st(ws, f"AF{r}")
        p[(tipo, "esp2")] = _st(ws, f"AD{r}")
    return p


def _sinal(l):
    """Sinal da coluna Diferença, como no modelo: +1 para entradas, receitas e resultados; -1 para gastos
    (custos e despesas são negativos, então diferença positiva = gastou mais que no mês anterior)."""
    if l["tipo"] == "resultado" or l["secao"] == "ENTRADAS":
        return 1
    if l["tipo"] == "secao":
        return 1 if l["secao"] in ("FINANCEIRO", "NAO_OPERACIONAL") else -1
    return 1 if l["grupo"].startswith(("0801", "0901")) else -1


def nome_aba_mes(ano, mes):
    return f"{mes:02d}{ano}"


def gerar_comparativo(caminho, empresa, meses):
    """meses: [(ano, mes, base, dre)] em ordem cronológica, o último é o mês do report (saída de montar_meses).
    Gera uma aba por mês (REPORT no layout do modelo, com valores) e a aba de comparação com a diferença entre
    o mês do report e o mês anterior."""
    wb = openpyxl.load_workbook(MODELO_COMPARATIVO)
    ws_m1, ws_m2, ws_cmp = wb.worksheets
    p_rep = _estilos_report(ws_m2)
    grade = ws_m2.sheet_view.showGridLines
    ws_m1.title, ws_m2.title = "_modelo_1", "_modelo_2"  # libera os nomes 082026 / 092026
    abas = []
    for i, (ano, mes, _base, dre) in enumerate(meses):
        ws = wb.copy_worksheet(ws_m2)
        ws.sheet_view.showGridLines = grade
        ws.title = nome_aba_mes(ano, mes)

        def valor(l, coluna, r):
            return round(l["valores"].get(coluna, 0.0), 2)

        info = _escrever_report(ws, p_rep, dre, ano, mes, valor)
        abas.append((ws.title, info))
    p = _estilos_comparativo(ws_cmp)
    wb.remove(ws_m1)
    wb.remove(ws_m2)
    a0, m0 = meses[0][:2]
    a1, m1 = meses[-1][:2]
    ws_cmp.title = (f"Comparativo {m0:02d}x{m1:02d}" if len(meses) == 2 else
                    f"Comparativo {m0:02d}-{a0 % 100:02d} a {m1:02d}-{a1 % 100:02d}")
    wb.move_sheet(ws_cmp, offset=len(wb.sheetnames))
    _escrever_comparativo(ws_cmp, p, meses, abas)
    wb.active = len(wb.sheetnames) - 1
    for ws in wb.worksheets:
        ws.sheet_view.tabSelected = ws is ws_cmp
    wb.save(caminho)
    return caminho


def _escrever_comparativo(ws, p, meses, abas):
    larg = {k: v.width for k, v in ws.column_dimensions.items()}
    zoom = ws.sheet_view.zoomScale
    _limpar(ws)
    for k in list(ws.column_dimensions.keys()):
        if k not in ("A", "B"):
            del ws.column_dimensions[k]
    dre = meses[-1][3]
    blocos = _blocos(dre)
    n = len(meses)
    largura = n + 1  # meses + diferença
    _put(ws, 2, 1, None, p["A2"])
    _put(ws, 2, 2, "CashFlowReport - Por Conta", p["titulo"])
    _put(ws, 3, 2, "Descrição da Conta", p["desc"])
    aba_atual = abas[-1][0]
    pos = {}
    c = 3
    for tipo, nome in blocos:
        pos[(tipo, nome)] = c
        col_mes = L(abas[-1][1]["pos"][(tipo, nome)])
        _put(ws, 2, c, f"='{aba_atual}'!{col_mes}2", p[("h", 0)])
        for k in range(1, largura):
            _put(ws, 2, c + k, None, p[("h", 1 if k < largura - 1 else 2)])
        ws.merge_cells(start_row=2, start_column=c, end_row=2, end_column=c + largura - 1)
        for k, (aba, _info) in enumerate(abas):
            _put(ws, 3, c + k, f"='{aba}'!$C$3", p[("d", 0 if k == 0 else 1)])
        _put(ws, 3, c + n, "Diferença", p[("d", 2)])
        _put(ws, 2, c + largura, None, p[("h", "esp")])
        _put(ws, 3, c + largura, None, p[("d", "esp")])
        ws.column_dimensions[L(c)].width = larg.get("C", 14.43)
        ws.column_dimensions[L(c + largura)].width = larg.get("F", 1.14)
        c += largura + 1
    col_chave, col_sinal = c + 1, c + 2
    _put(ws, 3, c, None, p[("esp2", "h")])
    _put(ws, 3, col_chave, "Conta (chave)", p[("chave", "h")])
    _put(ws, 3, col_sinal, "Sinal dif. (-1 = gasto)", p[("sinal", "h")])
    ws.column_dimensions[L(col_chave)].width = larg.get("AE", 48.14)
    ws.column_dimensions[L(col_chave)].hidden = True
    ws.column_dimensions[L(col_sinal)].width = larg.get("AF", 22.29)
    ws.column_dimensions[L(col_sinal)].hidden = True
    K, S = L(col_chave), L(col_sinal)
    for i, l in enumerate(dre["linhas"]):
        r, t = 4 + i, l["tipo"]
        _put(ws, r, 1, l["nivel"], p[(t, "A")])
        _put(ws, r, 2, l["descricao"], p[(t, "B")])
        if t in ("secao", "resultado"):
            ws.row_dimensions[r].height = 15.75
        for (tipo, nome), cp in pos.items():
            for k, (aba, info) in enumerate(abas):
                X = L(info["pos"][(tipo, nome)])
                f = f"=IFERROR(INDEX('{aba}'!{X}:{X},MATCH(${K}{r},'{aba}'!$B:$B,0)),0)"
                _put(ws, r, cp + k, f, p[(t, 0 if k == 0 else 1)])
            ant, atu = L(cp + n - 2), L(cp + n - 1)
            _put(ws, r, cp + n, f"=({atu}{r}-{ant}{r})*${S}{r}", p[(t, 2)])
            _put(ws, r, cp + largura, None, p[(t, "esp")])
        _put(ws, r, c, None, p[(t, "esp2")])
        _put(ws, r, col_chave, f"=B{r}", p[(t, "chave")])
        _put(ws, r, col_sinal, _sinal(l), p[(t, "sinal")])
    ws.freeze_panes = "C4"
    ws.sheet_view.zoomScale = zoom


# ------------------------------------------------------------------ dados dos gráficos do PowerPoint
def dados_graficos(dre):
    """(receita x custos x despesas, custos e despesas por conta) para os gráficos de pizza."""
    ind = D.indicadores(dre)
    resumo = [("Receita", abs(ind["Entradas"]["valor"])), ("Custos", abs(ind["Custos Operacionais"]["valor"])),
              ("Despesas", abs(ind["Despesas Operacionais"]["valor"]))]
    gastos = [(_sem_codigo(l["descricao"]), -l["total"]) for l in dre["linhas"]
              if l["tipo"] == "conta" and l["secao"] in ("CUSTOS", "DESPESAS") and l["total"] < 0]
    gastos.sort(key=lambda x: -x[1])
    return resumo, gastos
