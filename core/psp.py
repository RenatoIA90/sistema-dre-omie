"""Regras exclusivas da empresa PSP HUB - ESTUDOS EM INFRAESTRUTURA E URBANISMO.

Fechamento no layout do modelo "REPORT - FECHAMENTO PSP 06_2026 DRE.xlsx" (assets/modelos/psp_fechamento.xlsx):
- as linhas (nível 3) são os FORNECEDORES e CLIENTES, não as categorias do OMIE. O fornecedor que já está no
  modelo define a linha e o grupo (ex.: MASTER PRO fica em 040100-CACHE PARA PROFESSORES mesmo lançado em
  "Cursos e treinamentos"); fornecedor novo entra no grupo da categoria do OMIE, no fim do grupo, com o
  próximo código livre. Impostos e resultado financeiro/não operacional continuam por categoria
  (ISS, INSS/IRRF/CSRF, rendimentos, tarifas, empréstimos...).
- colunas: mesmo mês do ano anterior | mês do report | diferença (AH) | acumulado do ano (YTD), com AV.

A formatação vem da própria planilha-modelo: cada célula escrita recebe o estilo da célula equivalente do modelo.
"""
import os
import re
from collections import defaultdict
from copy import copy, deepcopy
from datetime import datetime

import openpyxl

from . import dre as D
from .modelos import completar_grupos, relatorio_sugerido
from .plano import FCL, IGNORAR, RESULTADO_LIQUIDO, RESULTADOS, SECOES, norm

MODELO_FECHAMENTO = os.path.join(os.path.dirname(os.path.dirname(__file__)), "assets", "modelos",
                                 "psp_fechamento.xlsx")
NOME_MODELO = "PSP HUB (modelo REPORT - FECHAMENTO PSP) (pronto)"
CNPJ_RAIZ = "48816886"
NOME_COLUNA = "PSP"

# Plano do modelo, na ordem em que aparece no REPORT: código, nome, seção, linhas (código, fornecedor/cliente).
PLANO = [
    ("010200", "SERVIÇOS PRESTADOS", "ENTRADAS", [
        ("010201", "MUNICIPIO DE PORTO VELHO"),
        ("010102", "HIRIA A NUERNBERGMESSE BRASIL BUSINESS"),
        ("010103", "ENLEV PARTICIPACOES EM SAUDE S.A"),
        ("010104", "IPERON"),
        ("010105", "PEZCO CONSULTORIA"),
        ("010106", "TOCANTINS PARCERIAS"),
        ("010107", "MARQUISE SERVIÇOS AMBIENTAIS"),
        ("010107", "OP EXTERIOR 3026813145"),
    ]),
    ("020100", "IMPOSTOS", "DEDUCOES", [
        ("020102", "(-) ISS"),
        ("020103", "(-) INSS / IRRF / CSRF"),
    ]),
    ("040100", "CACHE PARA PROFESSORES", "CUSTOS", [
        ("040101", "ALCANCCI ASSESSORIA E CONSULTORIA"),
        ("040102", "PRISCILLA MACEDO MAGALHAES BELLAZZI"),
        ("040103", "PROJECT CONSULTING & CORPORATE TRAINING"),
        ("040104", "LIDIA GRACE MONTEIRO LIMA 04273485281"),
        ("040105", "M & M GRAFICA E EDITORA LTDA"),
        ("040106", "NOESIS SP"),
        ("040107", "JULIANA CRISTINA LUVIZOTTO"),
        ("040108", "MAXX TRAINING"),
        ("040109", "REVIVI TREINAMENTOS"),
        ("040110", "HARPA 645"),
        ("040111", "MASTER PRO"),
        ("040112", "PROJECT CONSULTING E CORPORATE TRAINNING"),
        ("040113", "RICARDO TARDELLI"),
    ]),
    ("040200", "SALARIO", "CUSTOS", [
        ("040201", "LUCAS FRANCO DE GODOY"),
        ("040201", "LUCAS PAULO DE ALMEIDA COSTA"),
    ]),
    ("040300", "MÃO DE OBRA TEMPORARIA", "CUSTOS", [
        ("040301", "TALIANE CRISTINA RODRIGUES DE SOUZA"),
    ]),
    ("060100", "CONTABILIDADE", "DESPESAS", [
        ("060102", "SAVIAN CONTABILIDADE"),
    ]),
    ("060200", "VIAGENS / HOSPEDAGEM", "DESPESAS", [
        ("060201", "FLAMBOYANT COMERCIO HOTELEIRAS LTDA"),
        ("060202", "SLAVIERO EMPREENDI HOTELEIROS LTDA"),
        ("060203", "ELAINE OLIVEIRA MENDES"),
    ]),
    ("060300", "MARKETING", "DESPESAS", [
        ("060301", "MB PRODUCTION"),
        ("060302", "ENGRENAGEM EDITORIAL LTDA."),
        ("060303", "FACEBOOK SERVICOS ONLINE DO BRASIL LTDA."),
        ("060304", "VERSATIL COMUNICACAO VISUAL (GRAFICA)"),
        ("060305", "COPIADORA RORIZ LTDA (GRAFICA)"),
        ("060306", "EDNEY EITY HATYIA 17681044841 (FOTOGRAFO)"),
        ("060307", "ARRISCA COMUNICACAO VISUAL LTDA"),
        ("060308", "LOJA DO PORTA BANNER"),
        ("060309", "ENGAGED"),
    ]),
    ("060400", "ADMINISTRATIVO", "DESPESAS", [
        ("060401", "R I SERVICOS DE IMPRESSOS PERSONALIZADOS LTDA"),
        ("060402", "ENDEREÇO FISCAL SP LTDA (END FISCAL)"),
        ("060403", "CHEERS BAR E DRINKS"),
        ("060404", "VERDE LEASING"),
        ("060405", "33.379.949 NALIAN RODRIGUES MATIOLLI"),
        ("060406", "DELICIAS DA CIDA"),
        ("060407", "BPLUS (CERTIFICADO DIGITAL)"),
        ("060408", "NALIAN RODRIGUES MATIOLLI"),
    ]),
    ("060500", "CONSULTORIA", "DESPESAS", [
        ("060501", "TERRANOVA CONSULTORIA"),
        ("060502", "FABIO DE SOUSA SANTOS"),
        ("060503", "ELAINE OLIVEIRA MENDES"),
        ("060504", "LOS BUENOS SERVICOS ADMINISTRATIVOS"),
        ("060505", "ALMEIDA COSTA CONSULTORIA"),
        ("060506", "PEZCO CONSULTORIA DESENVOLVIMENTO LTDA"),
        ("060507", "KRAFT ENGENHARIA E CONSTRUCOES"),
        ("060508", "CHRYSTIAN JOSE CUNHA TRINDADE"),
        ("060509", "FINANZA PLANEJAMENTO FINANCEIRO"),
        ("060510", "GIOVANI SANTANA DE OLIVEIRA"),
        ("060511", "IGS - INSTITUTO DE GESTAO SETORIAL"),
    ]),
    ("060800", "CURSOS E TREINAMENTOS", "DESPESAS", [
        ("060801", "CCC CONSULTORIA LTDA"),
        ("060802", "COLOR SYSTEM"),
        ("060803", "CENTRO INTEGRACAO EMPRESA ESCOLA CIEE"),
    ]),
    ("060600", "JURÍDICO", "DESPESAS", [
        ("060601", "ADVOCACIA LUIS HENRIQUE DA SILVA"),
        ("060602", "APPARECIDO E SOCIEDADE DE ADVOGADOS"),
        ("060603", "JOSE APARECIDO SOCIEDADE INDIVIDUAL DE ADVOCACIA"),
        ("060604", "MATEUS KLEIN - SOCIEDADE INDIVIDUAL DE ADVOCACIA"),
        ("060604", "CORDEIRO, LIMA"),
    ]),
    ("060700", "REEMBOLSOS", "DESPESAS", [
        ("060701", "FREDERICO ARAUJO TUROLLA"),
        ("060702", "ANA CAROLINE OLIVEIRA DE SOUZA BUENO"),
        ("060703", "CARLOS AUGUSTO ANACLETO"),
        ("060704", "ANDRE CASTRO CARVALHO"),
        ("060705", "Isabella Gomide"),
        ("060706", "MARINA FARIA FRANCO DE CAMPOS"),
    ]),
    ("080100", "RECEITAS FINANCEIRAS", "FINANCEIRO", [
        ("080101", "RENDIMENTOS BANCÁRIOS"),
        ("080102", "EMPRESTIMO BANCÁRIOS"),
    ]),
    ("080200", "DESPESAS FINANCEIRAS", "FINANCEIRO", [
        ("080201", "EMPRESTIMOS"),
        ("080202", "TARIFAS BANCÁRIAS"),
        ("080203", "IOF"),
        ("080204", "JUROS CHEQUE ESPECIAL"),
    ]),
    ("090100", "RECEITAS NÃO OPERACIONAIS", "NAO_OPERACIONAL", [
        ("090103", "RECEITAS NÃO IDENTIFICADAS"),
    ]),
    ("090200", "DESPESAS NÃO OPERACIONAIS", "NAO_OPERACIONAL", [
        ("090201", "DISCOTECA DJ'S & ESTRUTURAS - LOCAÇÃO SOM"),
    ]),
]

# Categorias do OMIE da PSP -> grupo do modelo (classificação inicial; pode ser revista na aba Classificação).
CATEGORIAS = [
    ("Clientes - Serviços Prestados", "010200"), ("Prestação de serviços", "010200"), ("Patrocínio", "010200"),
    ("Serviços Prestados", "010200"),
    ("ISS", "020100"), ("CSRF", "020100"), ("INSS", "020100"), ("IRRF", "020100"), ("PIS", "020100"),
    ("COFINS", "020100"), ("CSLL", "020100"), ("Simples Nacional", "020100"),
    ("Cachê para professor", "040100"), ("Salários", "040200"), ("Mão de obra temporária", "040300"),
    ("Contabilidade", "060100"), ("Despesas de Viagens", "060200"), ("Hospedagem", "060200"),
    ("Marketing", "060300"), ("Fotográfo", "060300"), ("Endereço Fiscal", "060400"),
    ("Material Para Evento", "060400"), ("Consultoria", "060500"), ("Advogados", "060600"),
    ("Reembolso de despesas", "060700"), ("Retirada de Socios", "060700"), ("Cursos e treinamentos", "060800"),
    ("Rendimentos de Aplicações", "080100"), ("Empréstimos Bancários", "080100"),
    ("Juros sobre Empréstimos", "080200"), ("Pagamento de Empréstimos", "080200"), ("Tarifas Bancárias", "080200"), ("IOF", "080200"),
    ("Receitas não identificadas", "090100"),
    ("Adiantamento de Clientes", "010200"), ("DCTFWEB", "020100"), ("Vale Transporte", "040200"),
    ("Grafica", "060300"), ("Certificação", "060400"), ("Demais Despesas", "060400"),
    ("Despesas com eventos", "060400"),
]

# Grupos em que a linha é a CATEGORIA (não o fornecedor): (regex na categoria, entrada?, grupo, linha).
POR_CATEGORIA = {"020100", "080100", "080200", "090100"}
REGRAS_CATEGORIA = [
    (r"\biss\b", None, "020100", "(-) ISS"),
    (r"\b(inss|irrf|csrf|csll|pis|cofins|irpj|dctf\w*)\b", None, "020100", "(-) INSS / IRRF / CSRF"),
    (r"rendimento", True, "080100", "RENDIMENTOS BANCARIOS"),
    (r"emprestimo|financiamento", True, "080100", "EMPRESTIMO BANCARIOS"),
    (r"cheque especial", False, "080200", "JUROS CHEQUE ESPECIAL"),
    (r"emprestimo|financiamento", False, "080200", "EMPRESTIMOS"),
    (r"\biof\b", False, "080200", "IOF"),
    (r"tarifa", False, "080200", "TARIFAS BANCARIAS"),
    (r"nao identificad", True, "090100", "RECEITAS NAO IDENTIFICADAS"),
]


def _rotulo(cod, nome):
    return f"{cod}-{nome}"


_GRUPOS_MODELO = [g for g, *_ in PLANO]
_SECAO = {g: s for g, _, s, _ in PLANO}
# (grupo, nome normalizado) -> rótulo do modelo
_LINHA_MODELO = {(g, norm(n)): _rotulo(c, n) for g, _, _, linhas in PLANO for c, n in linhas}


def _entrada(grupo):
    return _SECAO.get(grupo) == "ENTRADAS" or grupo.startswith(("01", "0801", "0901"))


def eh_psp(empresa):
    if not empresa:
        return False
    cnpjs = "".join(ch for ch in str(empresa.get("cnpj") or "") if ch.isdigit() or ch == "\n")
    if any(c.startswith(CNPJ_RAIZ) for c in cnpjs.split("\n")):
        return True
    return "psp hub" in norm(empresa.get("nome"))


def modelo():
    """Modelo de DRE (grupos + contas) no formato de core.modelos; as contas são as categorias do OMIE."""
    grupos = [[g, n, s, relatorio_sugerido(n, s)] for g, n, s, _ in PLANO]
    return {"grupos": grupos, "contas": [list(c) for c in CATEGORIAS]}


# Versão do modelo acima. Ao subir a versão, atualizar_modelo() leva as mudanças às empresas já configuradas.
# v2: "Pagamento de Empréstimos" em 080200-DESPESAS FINANCEIRAS (linha EMPRESTIMOS), e não num grupo não operacional.
VERSAO_MODELO = 2


def precisa_aplicar_modelo(cfg):
    """Aplica o modelo PSP uma única vez; se depois o usuário trocar o modelo, a escolha dele vale."""
    return not cfg.get("psp_modelo_aplicado")


def aplicar_modelo(cfg, mapa):
    m = modelo()
    cfg["grupos"] = completar_grupos(m["grupos"])
    cfg["contas_modelo"] = m["contas"]
    cfg["modelo_dre"] = NOME_MODELO
    cfg["grupos_criados"] = []
    cfg["psp_modelo_aplicado"] = True
    cfg["psp_modelo_versao"] = VERSAO_MODELO
    D.reclassificar(mapa, cfg)


def atualizar_modelo(cfg, mapa):
    """Leva as mudanças do modelo PSP (VERSAO_MODELO) a uma empresa que já usa o modelo: refaz as classificações
    automáticas (as revistas na aba Classificação ficam como estão) e tira os grupos criados que ficaram sem uso.
    Roda uma vez por versão. Devolve True se mudou algo."""
    if not cfg.get("psp_modelo_aplicado") or cfg.get("psp_modelo_versao", 1) >= VERSAO_MODELO:
        return False
    cfg["psp_modelo_versao"] = VERSAO_MODELO
    if cfg.get("modelo_dre") == NOME_MODELO:
        cfg["contas_modelo"] = modelo()["contas"]
        D.reclassificar(mapa, cfg)
        em_uso = {m.get("grupo") for m in mapa.values()}
        criados = {g[:6] for g in cfg.get("grupos_criados") or []}
        cfg["grupos"] = [g for g in cfg["grupos"] if g[0] not in criados or g[0] in em_uso]
        cfg["grupos_criados"] = [g for g in cfg.get("grupos_criados") or [] if g[:6] in em_uso]
    return True


# ------------------------------------------------------------------ fornecedor -> linha do modelo
_PALAVRAS_VAZIAS = {"ltda", "lt", "sa", "s", "a", "me", "eireli", "epp", "de", "da", "do", "das", "dos", "e", "em"}


def _palavras(txt):
    return [w for w in re.sub(r"[^a-z0-9 ]", " ", norm(txt)).split() if w not in _PALAVRAS_VAZIAS]


def _iguais(a, b):
    """Mesma palavra, ou uma é abreviação da outra (ADVO ~ ADVOGADOS, EXT ~ EXTERIOR, PRODUCTION ~ PRODUCTIONS)."""
    return a == b or (min(len(a), len(b)) >= 3 and (a.startswith(b) or b.startswith(a)))


def semelhanca(a, b):
    """0 a ~1,5: quanto o nome a (favorecido do OMIE) parece o nome b (linha do modelo). Exige a mesma primeira
    palavra (fora números) e que pelo menos 2/3 das palavras de um dos dois nomes apareçam no outro."""
    pa, pb = _palavras(a), _palavras(b)
    if not pa or not pb:
        return 0.0
    if pa == pb:
        return 1.5
    pri = lambda p: next((w for w in p if not w.isdigit()), p[0])
    if not _iguais(pri(pa), pri(pb)):
        return 0.0
    sa, sb = list(dict.fromkeys(pa)), list(dict.fromkeys(pb))
    cobre = lambda x, y: sum(1 for w in x if any(_iguais(w, z) for z in y))
    acertos = max(cobre(sa, sb), cobre(sb, sa))
    frac = max(cobre(sa, sb) / len(sa), cobre(sb, sa) / len(sb))
    return frac + acertos * 0.01 if frac >= 2 / 3 else 0.0


def linha_do_fornecedor(favorecido, valor, grupo_categoria=None):
    """(grupo, rótulo) da linha do modelo para o fornecedor/cliente, ou None. Entradas só casam com linhas de
    entrada e saídas com linhas de gasto; empate fica com o grupo da categoria do OMIE."""
    if not str(favorecido or "").strip():
        return None
    entrada = valor >= 0
    melhor = None
    for i, (g, _, _, linhas) in enumerate(PLANO):
        if g in POR_CATEGORIA or _entrada(g) != entrada:
            continue
        for c, n in linhas:
            s = semelhanca(favorecido, n)
            if s:
                chave = (s, g == grupo_categoria, -i)
                if melhor is None or chave > melhor[0]:
                    melhor = (chave, g, _rotulo(c, n))
    return (melhor[1], melhor[2]) if melhor else None


def linha_da_categoria(categoria, valor, grupo):
    """(grupo, nome da linha) para os grupos por categoria (impostos, financeiro, não operacional)."""
    t = norm(categoria)
    for rx, entrada, g, nome in REGRAS_CATEGORIA:
        if re.search(rx, t) and (entrada is None or entrada == (valor >= 0)):
            return g, nome
    return grupo, str(categoria or "").strip()


# ------------------------------------------------------------------ base e DRE
def codificar(base, grupos=None):
    """Define grupo e conta_fechamento de cada lançamento pelo fornecedor/cliente (ou pela categoria nos grupos
    por categoria). Linhas fora do modelo recebem o próximo código livre do grupo, em ordem alfabética.
    Pode ser chamada de novo sobre a mesma base (usa o grupo original da categoria)."""
    novos = defaultdict(dict)  # grupo -> {nome: nome agrupado}
    destino = []
    for b in base:
        b.setdefault("grupo_categoria", b["grupo"])
        b.setdefault("conta_categoria", b["conta"])
        g0 = b["grupo_categoria"]
        if b.get("tipo", "Lançamento") != "Lançamento":  # rateio manual: fica na linha do lançamento de origem
            destino.append(("MANTER", None, None))
            continue
        if g0 == IGNORAR:
            destino.append((IGNORAR, b["conta_categoria"], None))
            continue
        if g0 in POR_CATEGORIA:
            g, nome = linha_da_categoria(b["componente"] or b["conta_categoria"], b["valor"], g0)
            rot = _LINHA_MODELO.get((g, norm(nome)))
            destino.append((g, nome, rot))
            if not rot:
                novos[g].setdefault(nome, nome)
            continue
        achado = linha_do_fornecedor(b["favorecido"], b["valor"], g0)
        if achado:
            destino.append((achado[0], None, achado[1]))
            continue
        nome = (b["favorecido"] or b["conta_categoria"] or "(sem nome)").strip().upper()
        # variações do mesmo fornecedor novo (ex.: "X LTDA" e "X") viram uma linha só
        igual = next((n for n in novos[g0] if semelhanca(nome, n) >= 1), None)
        nome = novos[g0].setdefault(igual or nome, igual or nome)
        destino.append((g0, nome, None))
    codigos = {}
    for g, nomes in novos.items():
        usados = [int(lbl[4:6]) for (gg, _), lbl in _LINHA_MODELO.items() if gg == g]
        n = max(usados, default=0)
        for nome in sorted(nomes, key=norm):
            n += 1
            codigos[(g, nome)] = f"{g[:4]}{n:02d}-{nome}"
    for b, (g, nome, rot) in zip(base, destino):
        if g == "MANTER":
            continue
        b["grupo"] = g
        if g == IGNORAR:
            b["conta"] = b["conta_fechamento"] = nome
            continue
        b["conta_fechamento"] = rot or codigos[(g, nome)]
        b["conta"] = b["conta_fechamento"].split("-", 1)[1]


def montar_base(lancs, cfg, mapa):
    base = D.montar_base(lancs, cfg, mapa, codificar=codificar)
    return base


def montar_dre(base, cfg, colunas=None):
    """DRE no layout fixo do modelo: todos os grupos e linhas do modelo aparecem (zerados se não houver
    movimento) e linhas/grupos novos entram no fim do grupo/seção. Mesmo formato de core.dre.montar_dre."""
    grupos_cfg = {g[0]: g for g in D.grupos_da_empresa(cfg)}
    cols = list(colunas) if colunas is not None else D.colunas(base, cfg)
    soma = defaultdict(lambda: defaultdict(float))
    conta_grupo = {}
    for b in base:
        if b["grupo"] == IGNORAR:
            continue
        soma[b["conta_fechamento"]][b["coluna"]] += b["valor"]
        conta_grupo[b["conta_fechamento"]] = b["grupo"]
    linhas, tot_secao = [], {}
    com_contas = set(conta_grupo.values())
    for sec, rotulo in SECOES:
        lin_sec = {"nivel": 1, "descricao": rotulo, "secao": sec, "grupo": "", "valores": {c: 0.0 for c in cols},
                   "tipo": "secao"}
        linhas.append(lin_sec)
        do_modelo = [(g, n, ls) for g, n, s, ls in PLANO if s == sec]
        outros = sorted(g for g in com_contas if g not in _GRUPOS_MODELO
                        and (grupos_cfg.get(g, (g, g, "DESPESAS"))[2] == sec))
        for g, nome, linhas_modelo in do_modelo + [(g, grupos_cfg.get(g, (g, g))[1], []) for g in outros]:
            lin_g = {"nivel": 2, "descricao": _rotulo(g, nome), "secao": sec, "grupo": g,
                     "valores": {c: 0.0 for c in cols}, "tipo": "grupo"}
            linhas.append(lin_g)
            rots = [_rotulo(c, n) for c, n in linhas_modelo]
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
    return {"colunas": cols, "agrupamentos": {}, "linhas": linhas}


ANT, MES, YTD = "ANO ANTERIOR", "MES", "ACUMULADO"


def montar_periodos(periodos, cfg, mapa):
    """periodos: {ANT: [(ano, mes, lancs)], MES: [...], YTD: [...]}. Monta a DRE do REPORT com as três colunas
    (mesmas linhas para todas). Devolve (dre, base do mês, categorias novas do mapa)."""
    todas, novas, base_mes = [], [], []
    for chave in (ANT, MES, YTD):
        for ano, mes, lancs in periodos.get(chave) or []:
            c = deepcopy(cfg)
            c["rateios_mes"] = (cfg.get("rateios") or {}).get(f"{ano}-{mes:02d}", [])
            novas += D.completar_mapeamento(mapa, lancs, c)
            if c.get("grupos") != cfg.get("grupos"):  # grupo novo criado por categoria fora do modelo
                cfg["grupos"], cfg["grupos_criados"] = c["grupos"], c.get("grupos_criados", [])
            base = montar_base(lancs, c, mapa)
            for b in base:
                b["coluna"] = chave
            todas += base
            if chave == MES:
                base_mes += base
    codificar(todas)
    return montar_dre(todas, cfg, colunas=[ANT, MES, YTD]), base_mes, novas


# ------------------------------------------------------------------ Excel no layout do modelo
def _st(ws, coord):
    return copy(ws[coord]._style)


def _put(ws, r, c, valor, estilo):
    cell = ws.cell(r, c)
    if valor is not None:
        cell.value = valor
    cell._style = copy(estilo)
    return cell


# colunas do modelo: C/D = ano anterior, F/G = mês, I/J = diferença e AH, L/M = acumulado; E, H, K, N separadores
_LETRAS = "ABCDEFGHIJKLMN"


def _estilos(ws):
    p = {}
    for tipo, r in (("h", 2), ("d", 3), ("secao", 4), ("grupo", 5), ("conta", 6), ("resultado", 18)):
        for x in _LETRAS:
            p[(tipo, x)] = _st(ws, f"{x}{r}")
    return p


def gerar_excel(caminho, empresa, ano, mes, dre3, base_mes):
    """REPORT no layout do modelo PSP (dre3 = saída de montar_periodos) + aba BASE CLASSIFICADA do mês."""
    wb = openpyxl.load_workbook(MODELO_FECHAMENTO)
    ws = wb.worksheets[0]
    p = _estilos(ws)
    for rng in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(rng))
    ws._cells.clear()
    for k in list(ws.row_dimensions.keys()):
        if k != 1:
            del ws.row_dimensions[k]
    ws.title = f"REPORT {mes:02d} {ano} - POR CONTA"

    cab2 = {"B": "CashFlowReport - Por Conta", "C": NOME_COLUNA, "F": NOME_COLUNA, "I": "Total",
            "L": f"Acumulado {ano}"}
    cab3 = {"B": "Descrição da Conta", "C": datetime(ano - 1, mes, 1), "D": "AV", "F": datetime(ano, mes, 1),
            "G": "AV", "I": f"{(ano - 1) % 100:02d}-{ano % 100:02d}", "J": "AH", "L": f"YTD - {ano}", "M": "AV"}
    for i, x in enumerate(_LETRAS, 1):
        _put(ws, 2, i, cab2.get(x), p[("h", x)])
        _put(ws, 3, i, cab3.get(x), p[("d", x)])
    for a, b in (("C", "D"), ("F", "G"), ("I", "J"), ("L", "M")):
        ws.merge_cells(f"{a}2:{b}2")

    linhas = dre3["linhas"]
    r0 = 4
    sec_row, grupos_sec, contas_g, res_row = {}, defaultdict(list), defaultdict(list), {}
    g_atual = None
    for i, l in enumerate(linhas):
        r = r0 + i
        if l["tipo"] == "secao":
            sec_row[l["secao"]] = r
        elif l["tipo"] == "grupo":
            grupos_sec[l["secao"]].append(r)
            g_atual = r
        elif l["tipo"] == "conta":
            contas_g[g_atual].append(r)
        else:
            res_row[l["descricao"]] = r
    ent = sec_row["ENTRADAS"]
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

    for i, l in enumerate(linhas):
        r, t = r0 + i, l["tipo"]
        if t in ("secao", "resultado"):
            ws.row_dimensions[r].height = 15.75
        cel = {"A": l["nivel"], "B": l["descricao"]}
        for X, chave in (("C", ANT), ("F", MES), ("L", YTD)):
            if t == "conta":
                f = round(l["valores"].get(chave, 0.0), 2)
            elif t == "grupo":
                rs = contas_g[r]
                f = f"=SUM({X}{rs[0]}:{X}{rs[-1]})" if len(rs) > 1 else (f"={X}{rs[0]}" if rs else 0)
            elif t == "secao":
                f = soma([f"{X}{x}" for x in grupos_sec.get(l["secao"], [])])
            elif l["descricao"] == FCL:
                f = f"={X}{res_row[RESULTADO_LIQUIDO]}"
            else:
                f = soma(cadeia[l["descricao"]](X))
            cel[X] = f
        cel["D"] = f"=IFERROR(C{r}/$C${ent},0)"
        cel["G"] = f"=IFERROR(F{r}/$F${ent},0)"
        cel["I"] = f"=F{r}-C{r}"
        cel["J"] = f"=IFERROR(I{r}/C{r},0)"
        cel["M"] = f"=IFERROR(L{r}/$L${ent},0)"
        for c, x in enumerate(_LETRAS, 1):
            _put(ws, r, c, cel.get(x), p[(t, x)])

    from .excel import _aba_base
    _aba_base(wb, base_mes)
    wb.active = 0
    wb.save(caminho)
    return caminho
