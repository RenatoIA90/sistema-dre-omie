"""Regras exclusivas da empresa HAGILI SERVIÇOS TERCEIRIZADOS LTDA.

Fechamento no layout do modelo "FECHAMENTO_HAGILI_09_2026 - POR PROJETO (ajustado)":
- colunas = PROJETOS (centros de custo do OMIE), em ordem alfabética, com DESPESAS SEM CENTRO DE CUSTO no fim e a
  coluna TOTAL HAGILI; é a mesma visão da tabela dinâmica Projeto x Categoria feita sobre a base do mês;
- linhas = plano fixo do modelo (todas aparecem, zeradas se não houver movimento). Categoria do OMIE que não está
  no modelo entra no fim do grupo dela, com o próximo código livre;
- projetos unificados como no modelo: OGGI SORVETES + SO GELO = OGGI / SÓ GELO; CONDOMINIO GR HORTOLANDIA =
  BTG HORTOLÂNDIA; TB HORTOLANDIA (e demais projetos "TB ...") + TRANSPORTADORA BARBARENSE = TRANSPORTADORA
  BARBARENSE LTDA; projetos "A DEFINIR CC", "(vazio)" e sem projeto = DESPESAS SEM CENTRO DE CUSTO;
- combustível: no projeto "COMBUSTIVEL MAQUINÁRIO - A DEFINIR CC" vai para Combustível Maquinário, no resto para
  Combustível Veículos;
- "Saída de Transferência" para o cartão de crédito = 090201-Pagamento de Cartão de Crédito; as demais
  transferências (aplicação CDB, entre contas) ficam fora da DRE.

Excel: REPORT por projeto (SUMIFS sobre a BASE CLASSIFICADA), SALDOS, RECEBIMENTOS (recebido x faturado), RATEIOS e
BASE CLASSIFICADA, na formatação do modelo.
"""
import re
from collections import defaultdict
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, PatternFill, Side
from openpyxl.utils import get_column_letter as L

from . import dre as D
from .excel import AMARELO_RATEIO, TOTAL_MEDIO, _aba_saldos
from .fluxo import AZUL_DIGITADO, CINZA_GRUPO, FMT, FMT_AV, FMT_CENT, FMT_MES, GRADE, MARINHO, SECAO, _cab, _fonte
from .modelos import completar_grupos
from .plano import FCL, IGNORAR, RESULTADO_LIQUIDO, RESULTADOS, SECOES, norm

NOME_MODELO = "HAGILI (modelo FECHAMENTO POR PROJETO) (pronto)"
CNPJ_RAIZ = "48277409"
TOTAL = "TOTAL HAGILI"
SEM_CC = D.SEM_CC
OGGI = "OGGI / SÓ GELO"
BARBARENSE = "TRANSPORTADORA BARBARENSE LTDA"
FORA_DRE = "FORA DA DRE - Transferência entre contas / aplicação"

# Plano do modelo, na ordem do REPORT: grupo, nome, seção, relatório (PowerPoint), linhas [(código, nome)].
PLANO = [
    ("010100", "RECEBIMENTOS", "ENTRADAS", "recebimentos", [("010101", "Serviço Prestado")]),
    ("020100", "IMPOSTOS", "DEDUCOES", "", [("020101", "Impostos s/ Faturamento")]),
    ("040100", "MÃO DE OBRA OPERACIONAL", "CUSTOS", "funcionarios", [
        ("040101", "Salários"), ("040102", "FT Freelancer"), ("040103", "Rescisões"),
        ("040104", "Adiantamento Salarial")]),
    ("040200", "ENCARGOS SOCIAIS", "CUSTOS", "", [
        ("040201", "FGTS Mensal"), ("040202", "FGTS Rescisório"), ("040203", "INSS")]),
    ("040300", "BENEFÍCIOS", "CUSTOS", "", [
        ("040301", "Vale Refeição"), ("040302", "Vale Alimentação"), ("040303", "Vale Transporte"),
        ("040304", "Seguro de Vida"), ("040305", "Assistência Médica")]),
    ("040400", "MATERIAIS E UNIFORMES", "CUSTOS", "fornecedores", [("040401", "Limpeza"), ("040402", "Uniformes")]),
    ("040500", "MÁQUINAS, EQUIPAMENTOS E VEÍCULOS", "CUSTOS", "fornecedores", [
        ("040501", "Máquinas e Equipamentos"), ("040502", "Manutenção de Equipamentos"), ("040503", "Fio de Nylon"),
        ("040504", "Combustível Maquinário"), ("040505", "Combustível Veículos"), ("040506", "Pedágio"),
        ("040507", "Locação de Veículos")]),
    ("060100", "CONTABILIDADE", "DESPESAS", "", [("060101", "Contabilidade")]),
    ("060200", "SISTEMA", "DESPESAS", "", [
        ("060201", "Sistemas (SOLIDES)"), ("060202", "Ponto Eletrônico"), ("060203", "Certificado Digital")]),
    ("060300", "PRESTADORES DE SERVIÇO / CONSULTORIA", "DESPESAS", "fornecedores", [
        ("060301", "Prestação de Serviço - Emerson"), ("060302", "Prestação de Serviço - Huiala"),
        ("060303", "Prestação de Serviço - Marcos"), ("060304", "Prestação de Serviço - Fábio"),
        ("060305", "Conselho Estratégico Empresarial")]),
    ("060400", "ADMINISTRATIVO", "DESPESAS", "", [
        ("060401", "Telefonia"), ("060402", "Marketing"), ("060403", "Material de Escritório"),
        ("060404", "Recrutamento/Pesquisa de Funcionário"), ("060405", "Reembolsos")]),
    ("060500", "ALUGUEL", "DESPESAS", "", [("060501", "Aluguel de Sala")]),
    ("060600", "JURÍDICO", "DESPESAS", "", [("060601", "Advogados")]),
    ("060700", "SÓCIOS", "DESPESAS", "", [("060701", "Saída de Sócio")]),
    ("080100", "RECEITAS FINANCEIRAS", "FINANCEIRO", "", [("080101", "Rendimentos de Aplicações")]),
    ("080200", "DESPESAS FINANCEIRAS", "FINANCEIRO", "", [
        ("080201", "Tarifa Bancária"), ("080202", "Tarifa Supera (Pagamentos PIX)"),
        ("080203", "Tarifas VR - VR e VA")]),
    ("090200", "DESPESAS NÃO OPERACIONAIS", "NAO_OPERACIONAL", "", [("090201", "Pagamento de Cartão de Crédito")]),
]

# Categoria do OMIE (como vem no extrato) -> linha do modelo, quando o nome não é o mesmo.
APELIDOS_CATEGORIA = {
    "Clientes - Serviços Prestados": "Serviço Prestado", "Serviços Prestados": "Serviço Prestado",
    "Simples Nacional": "Impostos s/ Faturamento", "DAS": "Impostos s/ Faturamento",
    "ISS": "Impostos s/ Faturamento", "Impostos": "Impostos s/ Faturamento",
    "Salário": "Salários", "Rescisão": "Rescisões", "Adiantamento": "Adiantamento Salarial",
    "FGTS": "FGTS Mensal", "Uniforme": "Uniformes", "Manutenção Equipamentos": "Manutenção de Equipamentos",
    "Combustivel": "Combustível Veículos", "Combustível": "Combustível Veículos",
    "Sistemas": "Sistemas (SOLIDES)", "Sistema": "Sistemas (SOLIDES)", "Ponto Eletronico": "Ponto Eletrônico",
    "Prestação de Serviços - Fábio": "Prestação de Serviço - Fábio", "Telefonia / Internet": "Telefonia",
    "Reembolso": "Reembolsos", "Reembolso de despesas": "Reembolsos", "Aluguel": "Aluguel de Sala",
    "Saida de Sócio": "Saída de Sócio", "Retirada de Sócios": "Saída de Sócio",
    "Tarifas Bancárias": "Tarifa Bancária", "Tarifa Supera": "Tarifa Supera (Pagamentos PIX)",
    "Saída de Transferência": "Pagamento de Cartão de Crédito",
}

# Categorias que não são linha do modelo mas têm grupo certo (viram linha nova no fim do grupo).
CATEGORIAS_EXTRAS = [
    ("Brinde Clientes", "060400"), ("Kit Boas Vindas", "060400"), ("Coletes", "040400"), ("EPI", "040400"),
    ("Estacionamento", "040500"), ("Tarifas Ticket", "080200"), ("IOF", "080200"), ("Juros", "080200"),
    ("Exames", "040300"), ("Exame Admissional", "040300"), ("Férias", "040100"), ("13º Salário", "040100"),
]

_GRUPOS_MODELO = [g for g, *_ in PLANO]
_SECAO = {g: s for g, _, s, _, _ in PLANO}
_LINHA = {}  # nome normalizado (linha do modelo ou apelido) -> (grupo, rótulo "040101-Salários")
for _g, _, _, _, _ls in PLANO:
    for _c, _n in _ls:
        _LINHA[norm(_n)] = (_g, f"{_c}-{_n}")
for _cat, _n in APELIDOS_CATEGORIA.items():
    _LINHA.setdefault(norm(_cat), _LINHA[norm(_n)])
_ROTULOS = {rot for _, rot in _LINHA.values()}


def eh_hagili(empresa):
    if not empresa:
        return False
    if any(c.startswith(CNPJ_RAIZ) for c in re.findall(r"\d+", re.sub(r"[./-]", "", str(empresa.get("cnpj") or "")))):
        return True
    return "hagili" in norm(empresa.get("nome"))


def modelo():
    """Modelo de DRE (grupos + contas) no formato de core.modelos; as contas são as categorias do OMIE."""
    grupos = [[g, n, s, r] for g, n, s, r, _ in PLANO]
    contas = [[n, g] for g, _, _, _, ls in PLANO for _, n in ls]
    contas += [[cat, _LINHA[norm(n)][0]] for cat, n in APELIDOS_CATEGORIA.items()]
    contas += [list(c) for c in CATEGORIAS_EXTRAS]
    return {"grupos": grupos, "contas": contas}


# Versão do modelo acima. Ao subir a versão, atualizar_modelo() leva as mudanças à empresa.
VERSAO_MODELO = 1


def precisa_aplicar_modelo(cfg):
    """Aplica o modelo HAGILI uma única vez; se depois o usuário trocar o modelo, a escolha dele vale."""
    return not cfg.get("hagili_modelo_aplicado")


def aplicar_modelo(cfg, mapa):
    m = modelo()
    cfg["grupos"] = completar_grupos(m["grupos"])
    cfg["contas_modelo"] = m["contas"]
    cfg["modelo_dre"] = NOME_MODELO
    cfg["grupos_criados"] = []
    cfg["modo_colunas"] = "projeto"
    cfg["hagili_modelo_aplicado"] = True
    cfg["hagili_modelo_versao"] = VERSAO_MODELO
    D.reclassificar(mapa, cfg)


def atualizar_modelo(cfg, mapa):
    """Leva as mudanças do modelo (VERSAO_MODELO) à empresa: refaz as classificações automáticas."""
    if not cfg.get("hagili_modelo_aplicado") or cfg.get("hagili_modelo_versao", 1) >= VERSAO_MODELO:
        return False
    cfg["hagili_modelo_versao"] = VERSAO_MODELO
    if cfg.get("modelo_dre") == NOME_MODELO:
        cfg["contas_modelo"] = modelo()["contas"]
        D.reclassificar(mapa, cfg)
    return True


def em_uso(cfg):
    """True se a empresa está com o modelo HAGILI (não trocou por outro na tela do plano de contas)."""
    return cfg.get("modelo_dre") == NOME_MODELO


# ------------------------------------------------------------------ projetos (colunas)
def coluna_projeto(nome):
    """Nome da coluna do REPORT para o projeto do OMIE (já com o apelido da tela de cadastro aplicado)."""
    n = norm(nome)
    if not n or n in ("(vazio)", "vazio") or "a definir" in n or "sem centro de custo" in n:
        return SEM_CC
    if n.startswith("oggi") or "so gelo" in n:
        return OGGI
    if n == "condominio gr hortolandia":
        return "BTG HORTOLÂNDIA"
    if "barbarense" in n or n == "tb" or n.startswith(("tb ", "tb-", "tb/")):  # TB HORTOLANDIA, TB LOG...
        return BARBARENSE
    return str(nome).strip()


def _eh_cartao(b):
    t = norm(f"{b.get('favorecido')} {b.get('observacoes')}")
    return "cartao" in t or "fatura" in t


# ------------------------------------------------------------------ base e DRE
def codificar(base, mapa=None):
    """Define coluna (projeto), grupo e conta_fechamento de cada lançamento no plano do modelo. Linhas fora do
    modelo recebem o próximo código livre do grupo, em ordem alfabética. Pode ser chamada de novo sobre a mesma
    base (usa o grupo e a conta originais da categoria)."""
    mapa = mapa or {}
    novos = defaultdict(dict)  # grupo -> {nome normalizado: nome}
    destino = []
    for b in base:
        if b.get("tipo", "Lançamento") != "Lançamento":  # rateio manual: fica na linha do lançamento de origem
            destino.append(("MANTER", None, None))
            continue
        b.setdefault("grupo_categoria", b["grupo"])
        b.setdefault("conta_categoria", b["conta"])
        if "projeto_relatorio" not in b:  # só na 1ª vez: no Excel do período a coluna já é o mês
            b["projeto_relatorio"] = b["coluna"]
            b["coluna"] = coluna_projeto(b["coluna"])
        g0, comp = b["grupo_categoria"], norm(b["componente"])
        # classificação revista pelo usuário na aba Classificação (guardada para recodificar no período)
        revisado = b.setdefault("revisado", mapa.get(comp, {}).get("automatico") is False)
        transf = comp in ("saida de transferencia", "entrada de transferencia")
        if g0 == IGNORAR or (transf and not revisado):
            if b["valor"] < 0 and _eh_cartao(b):
                g, rot = _LINHA[norm("Pagamento de Cartão de Crédito")]
                destino.append((g, None, rot))
            else:
                destino.append((IGNORAR, None, FORA_DRE))
            continue
        achado = None
        if comp.startswith("combust"):
            proj = norm(b.get("projeto"))
            nome = "Combustível Maquinário" if "maquin" in proj else (
                "Combustível Veículos" if "veicul" in proj else b["conta_categoria"])
            achado = _LINHA.get(norm(nome)) or _LINHA.get(comp)
        else:
            achado = _LINHA.get(norm(b["conta_categoria"])) or _LINHA.get(comp)
        if achado and revisado and achado[0] != g0:  # o usuário mudou o grupo na aba Classificação
            achado = None
        if achado:
            destino.append((achado[0], None, achado[1]))
            continue
        nome = str(b["conta_categoria"] or b["componente"] or "(sem categoria)").strip()
        nome = novos[g0].setdefault(norm(nome), nome)
        destino.append((g0, nome, None))
    codigos = {}
    for g, nomes in novos.items():
        usados = [int(r[4:6]) for gg, r in _LINHA.values() if gg == g]
        n = max(usados, default=0)
        for nome in sorted(nomes.values(), key=norm):
            n += 1
            codigos[(g, nome)] = f"{g[:4]}{n:02d}-{nome}"
    for b, (g, nome, rot) in zip(base, destino):
        if g == "MANTER":
            continue
        b["grupo"] = g
        b["conta_fechamento"] = rot or codigos[(g, nome)]
        b["conta"] = b["conta_fechamento"].split("-", 1)[1] if g != IGNORAR else b["conta_categoria"]


def montar_base(lancs, cfg, mapa):
    return D.montar_base(lancs, cfg, mapa, codificar=lambda base, _grupos: codificar(base, mapa))


def montar_dre(base, cfg, colunas=None):
    """DRE no layout fixo do modelo: todos os grupos e linhas do modelo aparecem (zerados se não houver
    movimento) e linhas/grupos novos entram no fim do grupo/seção. Colunas = projetos em ordem alfabética, SEM
    CENTRO DE CUSTO no fim. Mesmo formato de core.dre.montar_dre."""
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
        do_modelo = [(g, n, ls) for g, n, s, _, ls in PLANO if s == sec]
        outros = sorted(g for g in com_contas if g not in _GRUPOS_MODELO
                        and grupos_cfg.get(g, (g, g, "DESPESAS"))[2] == sec)
        for g, nome, linhas_modelo in do_modelo + [(g, grupos_cfg.get(g, (g, g))[1], []) for g in outros]:
            lin_g = {"nivel": 2, "descricao": f"{g}-{nome}", "secao": sec, "grupo": g,
                     "valores": {c: 0.0 for c in cols}, "tipo": "grupo"}
            linhas.append(lin_g)
            rots = [f"{c}-{n}" for c, n in linhas_modelo]
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
    return {"colunas": cols, "agrupamentos": {}, "linhas": linhas, "modo": "projeto"}


# ------------------------------------------------------------------ Excel no layout do modelo
_fino = Side(style="thin")
NOTAS = [
    "• Valores puxados por SUMIFS da aba BASE CLASSIFICADA (cada lançamento do extrato + linhas de rateio em "
    "amarelo claro).",
    "• Colunas = projetos do OMIE (mesma visão da tabela dinâmica Projeto x Categoria da base do mês).",
    "• Categorias com % no nome (ex.: \"Combustivel (98%); Tarifas Ticket (2%)\") foram divididas entre as contas "
    "pelos percentuais.",
    "• OGGI SORVETES + SO GELO unificados em OGGI / SÓ GELO; CONDOMINIO GR HORTOLANDIA = BTG HORTOLÂNDIA; "
    "TB HORTOLANDIA + TRANSPORTADORA BARBARENSE unificados em TRANSPORTADORA BARBARENSE LTDA.",
    "• O que não tinha projeto marcado (inclusive \"A DEFINIR CC\" não rateados) ficou em DESPESAS SEM CENTRO DE "
    "CUSTO. Rateios por posto: aba 2. Rateios do sistema (listados na aba RATEIOS).",
    "• Combustível do projeto \"COMBUSTIVEL MAQUINÁRIO - A DEFINIR CC\" = Combustível Maquinário; demais "
    "combustíveis = Combustível Veículos.",
    "• \"Saída de Transferência\" para cartão de crédito = 090201-Pagamento de Cartão de Crédito; aplicações e "
    "transferências entre contas ficam fora da DRE (entram só na conferência com o caixa).",
    "• Coluna TOTAL = soma das colunas de projeto + SEM CENTRO DE CUSTO; as conferências acima comparam com a base "
    "e com o caixa.",
]


def gerar_excel(caminho, empresa, ano, mes, base, dre, saldos, cfg):
    wb = Workbook()
    ws = wb.active
    ws.title = f"REPORT {mes:02d} {ano} - POR PROJETO"
    n_base = len(base) + 1
    _aba_report(ws, ano, mes, dre, n_base)
    _aba_saldos(wb, saldos, ano, mes)
    _aba_recebimentos(wb, ws, dre, ano, mes, n_base)
    _aba_rateios(wb, cfg.get("rateios_mes") or [])
    _aba_base(wb, base)
    wb.save(caminho)
    return caminho


def _rng(col, n_base):
    return f"'BASE CLASSIFICADA'!${col}$2:${col}${n_base}"


def _aba_report(ws, ano, mes, dre, n_base):
    cols = dre["colunas"]
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 54
    ws.row_dimensions[2].height = 45
    _cab(ws, 2, 2, "CashFlowReport - Por Projeto", fill=None, cor=None, borda=None)
    c = ws.cell(3, 2, "Descrição da Conta")
    c.font = _fonte(bold=True)
    c.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "C4"

    competencia = date(ano, mes, 1)
    pos = {}
    for i, nome in enumerate(cols + [TOTAL]):
        cp = 3 + 3 * i
        pos[nome] = cp
        ws.merge_cells(start_row=2, start_column=cp, end_row=2, end_column=cp + 1)
        tot = nome == TOTAL
        _cab(ws, 2, cp, nome, fill=MARINHO if tot else None, cor="FFFFFF" if tot else None)
        ws.cell(2, cp + 1).border = Border(top=_fino, bottom=_fino, right=_fino)
        for col, v, fmt in ((cp, competencia, FMT_MES), (cp + 1, "AV", "General")):
            h = ws.cell(3, col, v)
            h.font = _fonte(bold=True)
            h.alignment = Alignment(horizontal="center")
            h.number_format = fmt
        ws.column_dimensions[L(cp)].width = 13
        ws.column_dimensions[L(cp + 1)].width = 6.57
        ws.column_dimensions[L(cp + 2)].width = 1.14
    col_tot = pos[TOTAL]
    col_conf = col_tot + 3
    _cab(ws, 2, col_conf, "CONFERÊNCIA (TOTAL − BASE)", fill=None, cor=None, borda=None).font = _fonte(bold=True,
                                                                                                         size=9)
    ws.column_dimensions[L(col_conf)].width = 14
    faixa = range(2, col_tot + 2)

    rng_v, rng_c, rng_p = _rng("H", n_base), _rng("E", n_base), _rng("G", n_base)
    linhas = dre["linhas"]
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
    cadeia = {  # como no modelo: cada resultado soma o anterior com a seção de cima
        rot_rliq: lambda X: [f"{X}{sec_row['ENTRADAS']}", f"{X}{sec_row['DEDUCOES']}"],
        rot_bruto: lambda X: [f"{X}{res_row[rot_rliq]}", f"{X}{sec_row['CUSTOS']}"],
        rot_fco: lambda X: [f"{X}{res_row[rot_bruto]}", f"{X}{sec_row['DESPESAS']}"],
        RESULTADO_LIQUIDO: lambda X: [f"{X}{res_row[rot_fco]}", f"{X}{sec_row['FINANCEIRO']}",
                                      f"{X}{sec_row['NAO_OPERACIONAL']}"],
        FCL: lambda X: [f"{X}{res_row[RESULTADO_LIQUIDO]}"],
    }

    def soma(refs):
        return f"=SUM({','.join(refs)})" if refs else 0

    for i, l in enumerate(linhas):
        r, t = r0 + i, l["tipo"]
        bold = t != "conta"
        ws.cell(r, 1, l["nivel"]).font = _fonte(size=8, color="FFFFFF")
        c = ws.cell(r, 2, l["descricao"])
        c.font = _fonte(bold=bold)
        if t == "conta":
            c.alignment = Alignment(indent=1)
        for nome, cp in pos.items():
            X = L(cp)
            if nome == TOTAL:
                f = soma([f"{L(pos[x])}{r}" for x in cols])
            elif t == "conta":
                f = f"=SUMIFS({rng_v},{rng_c},$B{r},{rng_p},{X}$2)"
            elif t == "grupo":
                rs = contas_g[r]
                f = f"=SUM({X}{rs[0]}:{X}{rs[-1]})" if len(rs) > 1 else (f"={X}{rs[0]}" if rs else 0)
            elif t == "secao":
                rs = grupos_sec.get(l["secao"], [])
                f = f"={X}{rs[0]}" if len(rs) == 1 else soma([f"{X}{x}" for x in rs])
            else:
                f = soma(cadeia[l["descricao"]](X))
            cv = ws.cell(r, cp, f)
            cv.number_format = FMT
            cv.font = _fonte(bold=bold)
            ca = ws.cell(r, cp + 1, f"=IFERROR({X}{r}/{X}${ent},0)")
            ca.number_format = FMT_AV
            ca.font = _fonte(bold=bold, size=10, italic=True)
        if t in ("secao", "resultado"):
            ws.row_dimensions[r].height = 15.75
            for cc in faixa:
                ws.cell(r, cc).border = SECAO
        elif t == "grupo":
            for cc in faixa:
                ws.cell(r, cc).fill = PatternFill("solid", fgColor=CINZA_GRUPO)
        if t == "conta":
            T = L(col_tot)
            cc = ws.cell(r, col_conf, f"=ROUND({T}{r}-SUMIFS({rng_v},{rng_c},$B{r}),2)")
            cc.number_format = FMT_CENT
            cc.font = _fonte(size=10, italic=True, color="808080")

    fim = r0 + len(linhas)
    T = L(col_tot)
    fcl = res_row[FCL]
    conf = [
        ("Conferência: TOTAL − soma da BASE CLASSIFICADA (deve ser 0)",
         f"=ROUND({T}{fcl}-SUM({rng_v})+SUMIFS({rng_v},{rng_c},\"{FORA_DRE}\"),2)"),
        ("Conferência: resultado líquido + transferências fora da DRE − variação de caixa (deve ser 0)",
         f"=ROUND({T}{fcl}+SUMIFS({rng_v},{rng_c},\"{FORA_DRE}\")-SALDOS!E{{TOT}},2)"),
    ]
    for k, (txt, f) in enumerate(conf):
        ws.cell(fim + 1 + k, 2, txt).font = _fonte(size=10, italic=True)
        x = ws.cell(fim + 1 + k, col_tot, f)
        x.number_format = FMT_CENT
        x.font = _fonte(size=10, italic=True)
    ws._savian_conf_cell = ws.cell(fim + 2, col_tot)  # _aba_saldos troca {TOT} pela linha do total
    ws.cell(fim + 4, 2, "Notas:").font = _fonte(bold=True, size=10)
    for i, n in enumerate(NOTAS):
        ws.cell(fim + 5 + i, 2, n).font = _fonte(size=10)
    ws._hagili = {"pos": pos, "contas_entrada": [r0 + i for i, l in enumerate(linhas)
                                                 if l["tipo"] == "conta" and l["grupo"] == "010100"]}


def _aba_recebimentos(wb, rep, dre, ano, mes, n_base):
    """Recebido por projeto (SUMIFS da base nas linhas de 010100) x FATURADO (digitar) e a diferença."""
    ws = wb.create_sheet("RECEBIMENTOS")
    for col, w in zip("ABCDEFG", (2, 38, 12.43, 5.71, 12.71, 13.43, 31.57)):
        ws.column_dimensions[col].width = w
    ws.merge_cells("B2:G2")
    _cab(ws, 2, 2, "RECEBIMENTOS POR PROJETO - SERVIÇO PRESTADO", fill=None, cor=None)
    for c in range(3, 8):
        ws.cell(2, c).border = Border(top=_fino, bottom=_fino, right=_fino if c == 7 else None)
    for i, h in enumerate(["PROJETO", "RECEBIDO", "AV%", "FATURADO", "DIFERENÇA", "OBSERVAÇÃO"], start=2):
        _cab(ws, 3, i, h)
    lin = next((l for l in dre["linhas"] if l["tipo"] == "grupo" and l["grupo"] == "010100"), None)
    projetos = [c for c in dre["colunas"] if lin and round(lin["valores"].get(c, 0.0), 2) != 0]
    rotulos = [l["descricao"] for l in dre["linhas"] if l["tipo"] == "conta" and l["grupo"] == "010100"]
    tot = 4 + len(projetos)
    for r, p in enumerate(projetos, start=4):
        partes = [f"SUMIFS({_rng('H', n_base)},{_rng('E', n_base)},\"{x}\",{_rng('G', n_base)},B{r})"
                  for x in rotulos]
        vals = [p, "=" + "+".join(partes), f"=IFERROR(C{r}/$C${tot},0)", None, f"=IF(E{r}=\"\",\"\",C{r}-E{r})",
                None]
        for c, v in enumerate(vals, start=2):
            x = ws.cell(r, c, v)
            x.font = _fonte(color=AZUL_DIGITADO if c in (5, 7) else None)
            x.border = GRADE
            x.number_format = {3: FMT_CENT, 4: "0.0%", 5: FMT_CENT, 6: FMT_CENT}.get(c, "General")
    vals = [f"TOTAL RECEBIMENTO {mes:02d}/{ano}", f"=SUM(C4:C{tot - 1})" if projetos else 0,
            f"=SUM(D4:D{tot - 1})" if projetos else 0, f"=SUM(E4:E{tot - 1})" if projetos else 0,
            f"=IF(E{tot}=0,\"\",C{tot}-E{tot})", None]
    for c, v in enumerate(vals, start=2):
        x = ws.cell(tot, c, v)
        x.font = _fonte(bold=True)
        x.border = TOTAL_MEDIO
        x.number_format = {3: FMT_CENT, 4: "0.0%", 5: FMT_CENT, 6: FMT_CENT}.get(c, "General")
    ws.cell(tot + 2, 2, "FATURADO e OBSERVAÇÃO (em azul) são digitados conforme as notas emitidas no mês.").font = \
        _fonte(size=10, italic=True)


def _aba_rateios(wb, rateios):
    ws = wb.create_sheet("RATEIOS")
    for col, w in zip("ABCDEF", (2, 48, 40, 30, 30, 16)):
        ws.column_dimensions[col].width = w
    for i, h in enumerate(["RATEIO", "CONTA", "DE (PROJETO)", "PARA (PROJETO)", "VALOR RATEADO"], start=2):
        _cab(ws, 2, i, h)
    r = 3
    for a in rateios:
        vals = [a.get("descricao") or "Rateio", a.get("conta"), coluna_projeto(a.get("de")) if a.get("de") else "",
                coluna_projeto(a.get("para")), abs(float(a.get("valor") or 0))]
        for c, v in enumerate(vals, start=2):
            x = ws.cell(r, c, v)
            x.font = _fonte()
            x.border = GRADE
            if c == 6:
                x.number_format = FMT_CENT
        r += 1
    if rateios:
        for c, v in ((2, "TOTAL RATEADO"), (6, f"=SUM(F3:F{r - 1})")):
            x = ws.cell(r, c, v)
            x.font = _fonte(bold=True)
            x.border = TOTAL_MEDIO
            x.number_format = FMT_CENT
    else:
        ws.cell(3, 2, "Sem rateios manuais neste mês (cadastre na aba 2. Rateios do sistema).").font = _fonte()


def _aba_base(wb, base):
    ws = wb.create_sheet("BASE CLASSIFICADA")
    hdr = ["Data", "Cliente ou Fornecedor", "Categoria (original)", "Componente", "Conta do Fechamento",
           "Projeto (original)", "Projeto (fechamento)", "Valor (R$)", "Tipo", "Observações"]
    larg = [11, 38, 34, 26, 40, 32, 30, 13, 11, 60]
    for i, (h, w) in enumerate(zip(hdr, larg), start=1):
        _cab(ws, 1, i, h, borda=None)
        ws.column_dimensions[L(i)].width = w
    amarelo = PatternFill("solid", fgColor=AMARELO_RATEIO)
    for r, b in enumerate(base, start=2):
        proj = b.get("projeto_relatorio", b.get("projeto")) or ""
        vals = [b["data"], b["favorecido"], b["categoria"], b["componente"], b["conta_fechamento"],
                proj if b["tipo"] == "Lançamento" else b.get("projeto") or "", b["coluna"], b["valor"], b["tipo"],
                (b.get("observacoes") or "")[:500]]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(r, c, v)
            cell.font = _fonte()
            if c == 1:
                cell.number_format = "dd/mm/yyyy"
            if c == 8:
                cell.number_format = FMT_CENT
            if b["tipo"] != "Lançamento":
                cell.fill = amarelo
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{L(len(hdr))}{len(base) + 1}"
