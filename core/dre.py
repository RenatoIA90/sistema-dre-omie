"""Monta a base classificada e a DRE financeira (regime de caixa) de um mês."""
from collections import OrderedDict, defaultdict

from .plano import (FCL, GRUPOS_PADRAO, IGNORAR, RESULTADO_LIQUIDO, RESULTADOS, SECOES,
                    dividir_categoria, norm, sugerir_grupo)
from .modelos import A_CLASSIFICAR, sugerir_pelo_modelo

SEM_CC = "DESPESAS SEM CENTRO DE CUSTO"
NAO_REALIZADO = {"atrasado", "a vencer", "vence hoje", "previsto", "aberto", "em aberto", "a pagar", "a receber"}


def grupos_da_empresa(cfg):
    """Lista de grupos (codigo, nome, secao, relatorio) da empresa, ou o padrão."""
    gs = cfg.get("grupos") or [list(g) for g in GRUPOS_PADRAO]
    if not any(g[0] == A_CLASSIFICAR[0] for g in gs):
        gs = list(gs) + [list(A_CLASSIFICAR)]
    if not any(g[0] == IGNORAR for g in gs):
        gs = list(gs) + [list(GRUPOS_PADRAO[-1])]
    return [tuple(g) for g in gs]


def sugerir(categoria, valor, transferencia, cfg):
    """(grupo, nome da conta) para uma categoria: primeiro o modelo da empresa, depois as regras do plano
    padrão, adaptadas aos grupos do modelo (mesmo código, senão o primeiro grupo da mesma seção)."""
    if transferencia:
        return IGNORAR, categoria.strip()
    pelo_modelo = sugerir_pelo_modelo(categoria, cfg)
    if pelo_modelo:
        return pelo_modelo
    g = sugerir_grupo(categoria, valor)
    grupos = grupos_da_empresa(cfg)
    if g in {x[0] for x in grupos}:
        return g, categoria.strip()
    secao = next((x[2] for x in GRUPOS_PADRAO if x[0] == g), "DESPESAS")
    mesma = [x[0] for x in grupos if x[2] == secao and x[0] != A_CLASSIFICAR[0]]
    return (mesma[0] if mesma else A_CLASSIFICAR[0]), categoria.strip()


def realizado(l):
    return norm(l.get("situacao")) not in NAO_REALIZADO


def componentes(lancs):
    """Explode categorias rateadas ('A (30%); B (70%)') em componentes, com centavos fechando."""
    out = []
    for l in lancs:
        partes = dividir_categoria(l.get("categoria"))
        resto = round(l["valor"], 2)
        for i, (nome, pct) in enumerate(partes):
            v = resto if i == len(partes) - 1 else round(l["valor"] * pct, 2)
            resto = round(resto - v, 2)
            out.append(dict(l, componente=nome, valor=v))
    return out


def categorias_do_mes(lancs):
    """{chave_normalizada: {'categoria', 'valor', 'qtd', 'transferencia'}} dos componentes do mês."""
    res = OrderedDict()
    for c in componentes(lancs):
        k = norm(c["componente"]) or "(sem categoria)"
        r = res.setdefault(k, {"categoria": c["componente"] or "(sem categoria)", "valor": 0.0, "qtd": 0,
                               "transferencia": False})
        r["valor"] += c["valor"]
        r["qtd"] += 1
        r["transferencia"] |= bool(c.get("transferencia"))
    return res


def completar_mapeamento(mapa, lancs, cfg=None):
    """Garante uma entrada no mapeamento para cada categoria do mês. Devolve chaves novas."""
    novas = []
    for k, info in categorias_do_mes(lancs).items():
        if k not in mapa:
            grupo, conta = sugerir(info["categoria"], info["valor"], info["transferencia"], cfg or {})
            mapa[k] = {"categoria": info["categoria"], "grupo": grupo, "conta": conta, "automatico": True}
            novas.append(k)
    return novas


def reclassificar(mapa, cfg):
    """Depois de trocar o modelo: refaz as sugestões automáticas e as que apontam para grupos que não existem
    mais. Classificações revisadas pelo usuário em grupos que continuam existindo são mantidas."""
    codigos = {g[0] for g in grupos_da_empresa(cfg)}
    for m in mapa.values():
        if m.get("automatico") or m.get("grupo") not in codigos:
            transf = m.get("grupo") == IGNORAR
            sinal = 1 if m.get("grupo", "").startswith(("01", "0901", "0801")) else -1
            m["grupo"], m["conta"] = sugerir(m.get("categoria", ""), sinal, transf, cfg)
            m["automatico"] = True


def _coluna(l, cfg):
    modo = cfg.get("modo_colunas", "conta")
    if modo == "projeto":
        p = (l.get("projeto") or "").strip()
        p = cfg.get("apelidos_projeto", {}).get(p, p)
        return p or SEM_CC
    if modo == "conta":
        c = (l.get("conta_corrente") or "").strip()
        return cfg.get("apelidos_conta", {}).get(c, c) or "CONTA"
    return "TOTAL"


def montar_base(lancs, cfg, mapa):
    """Base classificada: uma linha por componente de lançamento (+ linhas de rateio manual)."""
    grupos = {g[0]: g for g in grupos_da_empresa(cfg)}
    so_realizado = cfg.get("apenas_realizado", True)
    base = []
    for c in componentes(lancs):
        if so_realizado and not realizado(c):
            continue
        k = norm(c["componente"]) or "(sem categoria)"
        m = mapa.get(k) or {"grupo": sugerir_grupo(c["componente"], c["valor"]), "conta": c["componente"]}
        g = m["grupo"] if m["grupo"] in grupos else "060900"
        base.append({
            "data": c["data"], "favorecido": c.get("favorecido") or c.get("razao_social") or "",
            "conta_corrente": c.get("conta_corrente", ""), "categoria": c.get("categoria", ""),
            "componente": c["componente"], "grupo": g, "conta": m.get("conta") or c["componente"],
            "projeto": c.get("projeto", ""), "coluna": _coluna(c, cfg), "valor": c["valor"],
            "tipo": "Lançamento", "situacao": c.get("situacao", ""), "observacoes": c.get("observacoes", ""),
        })
    _codificar_contas(base, grupos)
    for a in cfg.get("rateios_mes", []) or []:
        _aplicar_rateio(base, a)
    return base


def _codificar_contas(base, grupos):
    """Atribui códigos de nível 3 (ex.: 040101-Salários) por grupo, em ordem alfabética."""
    por_grupo = defaultdict(set)
    for b in base:
        por_grupo[b["grupo"]].add(b["conta"])
    codigos = {}
    for g, contas in por_grupo.items():
        for i, conta in enumerate(sorted(contas, key=norm), start=1):
            codigos[(g, conta)] = conta if g == IGNORAR else f"{g[:4]}{i:02d}-{conta}"
    for b in base:
        b["conta_fechamento"] = codigos[(b["grupo"], b["conta"])]


def _aplicar_rateio(base, a):
    """Rateio manual: move `valor` da conta de uma coluna (ex.: sem centro de custo) para outra."""
    try:
        valor = abs(float(a.get("valor") or 0))
    except (TypeError, ValueError):
        return
    alvo = [b for b in base if b["conta_fechamento"] == a.get("conta") or b["conta"] == a.get("conta")]
    if not valor or not alvo:
        return
    ref = alvo[0]
    sinal = 1 if sum(b["valor"] for b in alvo) >= 0 else -1
    for coluna, v in ((a.get("de"), -sinal * valor), (a.get("para"), sinal * valor)):
        base.append(dict(ref, data=ref["data"], favorecido=a.get("descricao") or "Rateio",
                         coluna=coluna, valor=round(v, 2), tipo="Rateio", situacao="",
                         observacoes=a.get("descricao") or "Rateio manual"))


def colunas(base, cfg):
    cols = []
    for b in base:
        if b["grupo"] != IGNORAR and b["coluna"] not in cols:
            cols.append(b["coluna"])
    ordem = cfg.get("ordem_colunas") or []
    cols.sort(key=lambda c: (c == SEM_CC, ordem.index(c) if c in ordem else len(ordem), norm(c)))
    return cols


def montar_dre(base, cfg):
    """Retorna dict com colunas, subtotais de colunas e linhas da DRE.

    Cada linha: {nivel, descricao, secao, grupo, valores{col: v}, total, tipo}
    tipo: 'secao' | 'grupo' | 'conta' | 'resultado'
    """
    grupos = grupos_da_empresa(cfg)
    cols = colunas(base, cfg)
    soma = defaultdict(lambda: defaultdict(float))  # conta_fechamento -> col -> v
    conta_info = {}
    for b in base:
        if b["grupo"] == IGNORAR:
            continue
        soma[b["conta_fechamento"]][b["coluna"]] += b["valor"]
        conta_info[b["conta_fechamento"]] = b["grupo"]

    def vazio():
        return {c: 0.0 for c in cols}

    linhas, tot_secao = [], {}
    for sec, rotulo in SECOES:
        lin_sec = {"nivel": 1, "descricao": rotulo, "secao": sec, "grupo": "", "valores": vazio(), "tipo": "secao"}
        linhas.append(lin_sec)
        for cod, nome, s, _rel in grupos:
            if s != sec:
                continue
            contas = sorted([cf for cf, g in conta_info.items() if g == cod])
            if not contas and not cfg.get("mostrar_grupos_vazios", False):
                continue
            lin_g = {"nivel": 2, "descricao": f"{cod}-{nome}", "secao": sec, "grupo": cod, "valores": vazio(),
                     "tipo": "grupo"}
            linhas.append(lin_g)
            for cf in contas:
                vals = {c: round(soma[cf].get(c, 0.0), 2) for c in cols}
                linhas.append({"nivel": 3, "descricao": cf, "secao": sec, "grupo": cod, "valores": vals,
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

    # subtotais por agrupamento de colunas (ex.: TOTAL MATRIZ / TOTAL FILIAL)
    agrup = OrderedDict()
    for c in cols:
        g = (cfg.get("grupos_colunas") or {}).get(c)
        if g:
            agrup.setdefault(g, []).append(c)
    return {"colunas": cols, "agrupamentos": agrup, "linhas": linhas}


def linha(dre, descricao_inicio):
    for l in dre["linhas"]:
        if l["descricao"].startswith(descricao_inicio):
            return l
    return {"valores": {}, "total": 0.0}


def indicadores(dre):
    """Totais da coluna TOTAL para a análise do mês."""
    def t(prefixo):
        return linha(dre, prefixo)["total"]
    ent = t("(+) ENTRADAS")
    ind = OrderedDict([
        ("Entradas", t("(+) ENTRADAS")),
        ("Deduções (impostos)", t("(-) DEDUÇÕES")),
        ("Receita Líquida", t("(=) RECEITA OPERACIONAL")),
        ("Custos Operacionais", t("(-) CUSTOS")),
        ("Resultado Bruto", t("(=) RESULTADO BRUTO")),
        ("Despesas Operacionais", t("(-) DESPESAS")),
        ("Resultado Operacional", t("(=) FCO")),
        ("Resultado Financeiro", t("(+/-) RESULTADO FINANCEIRO")),
        ("Não Operacional", t("(+/-) RESULTADO NÃO")),
        ("Resultado Líquido", t("(=) RESULTADO LÍQUIDO")),
    ])
    return {k: {"valor": v, "pct": (v / ent if ent else 0.0)} for k, v in ind.items()}


# ---------- quadros auxiliares (Excel e PowerPoint) ----------

def quadro_por(base, filtro, chave_bloco, chave_item, sinal=-1):
    """{bloco: [(item, valor_absoluto)]} ordenado por valor decrescente."""
    acc = defaultdict(lambda: defaultdict(float))
    for b in base:
        if filtro(b):
            acc[chave_bloco(b)][chave_item(b) or "(sem nome)"] += b["valor"] * sinal
    out = OrderedDict()
    for bloco in sorted(acc, key=lambda k: -sum(acc[k].values())):
        itens = [(i, round(v, 2)) for i, v in acc[bloco].items() if round(v, 2) != 0]
        itens.sort(key=lambda x: -x[1])
        if itens:
            out[bloco] = itens
    return out


def _rel_grupos(cfg, rel):
    return {g[0] for g in grupos_da_empresa(cfg) if g[3] == rel}


def recebimentos(base, cfg):
    gs = _rel_grupos(cfg, "recebimentos")
    f = lambda b: b["grupo"] in gs and b["tipo"] == "Lançamento"
    if cfg.get("modo_colunas") == "projeto":
        return quadro_por(base, f, lambda b: "RECEBIMENTOS POR PROJETO", lambda b: b["coluna"], sinal=1)
    por_cat = quadro_por(base, f, lambda b: "RECEBIMENTOS POR CATEGORIA", lambda b: b["conta"], sinal=1)
    poucas = sum(len(v) for v in por_cat.values()) <= 2
    if cfg.get("recebimentos_por_cliente") or poucas:
        return quadro_por(base, f, lambda b: "RECEBIMENTOS - " + b["conta"].upper(), lambda b: b["favorecido"], sinal=1)
    return por_cat


def pagamentos(base, cfg, rel):
    gs = _rel_grupos(cfg, rel)
    f = lambda b: b["grupo"] in gs and b["tipo"] == "Lançamento" and b["valor"] < 0
    return quadro_por(base, f, lambda b: b["conta"].upper(), lambda b: b["favorecido"])


def transferencias(base):
    acc = OrderedDict()
    for b in base:
        if b["grupo"] != IGNORAR:
            continue
        r = acc.setdefault(b["conta_corrente"] or "CONTA", {"enviada": 0.0, "recebida": 0.0})
        r["enviada" if b["valor"] < 0 else "recebida"] += b["valor"]
    return acc


def conferencia(base, dre, saldos):
    total_base = round(sum(b["valor"] for b in base if b["grupo"] != IGNORAR and b["tipo"] == "Lançamento"), 2)
    total_dre = linha(dre, RESULTADO_LIQUIDO)["total"]
    transf = round(sum(b["valor"] for b in base if b["grupo"] == IGNORAR), 2)
    var_caixa = round(sum(s["saldo_final"] - s["saldo_inicial"] for s in saldos), 2) if saldos else None
    return {
        "dre_menos_base": round(total_dre - total_base, 2),
        "resultado_liquido": total_dre,
        "transferencias_liquidas": transf,
        "variacao_caixa": var_caixa,
        "dif_caixa": None if var_caixa is None else round(total_dre + transf - var_caixa, 2),
    }
