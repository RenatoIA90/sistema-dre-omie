"""Textos automáticos da análise do mês (editáveis na tela antes de gerar o PowerPoint)."""
from . import dre as D

MESES = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro",
         "Outubro", "Novembro", "Dezembro"]


def brl(v, casas=2):
    s = f"{abs(v):,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return ("-R$ " if v < 0 else "R$ ") + s


def pct(v, casas=0):
    x = round(v * 100, casas)
    if x == 0:
        x = 0.0
    return f"{x:.{casas}f}%".replace(".", ",")


def _grupos_secao(dre, secao, n=3):
    gs = [l for l in dre["linhas"] if l["tipo"] == "grupo" and l["secao"] == secao and l["total"]]
    gs.sort(key=lambda l: abs(l["total"]))
    gs.reverse()
    return [(l["descricao"].split("-", 1)[-1].strip().lower(), l["total"]) for l in gs[:n]]


def gerar_textos(empresa, ano, mes, base, dre, saldos, cfg):
    ind = D.indicadores(dre)
    v = {k: x["valor"] for k, x in ind.items()}
    ent = v["Entradas"] or 0.0
    p = lambda x: abs(x / ent) if ent else 0.0
    rec = D.recebimentos(base, cfg)
    itens_rec = sorted([i for its in rec.values() for i in its], key=lambda x: -x[1])
    custos = _grupos_secao(dre, "CUSTOS")
    desp = _grupos_secao(dre, "DESPESAS")
    conf = D.conferencia(base, dre, saldos)
    s_ini = sum(s["saldo_inicial"] for s in saldos)
    s_fim = sum(s["saldo_final"] for s in saldos)
    transf = D.transferencias(base)

    def lista(pares):
        return ", ".join(f"{n} ({brl(abs(x))})" for n, x in pares)

    fech = []
    principais = ", ".join(f"{n} ({pct(x / ent)})" for n, x in itens_rec[:4]) if ent else ""
    fech.append(f"Recebimentos: a empresa teve {brl(ent)} em entradas no mês."
                + (f" Os principais foram {principais}." if principais else ""))
    if v["Deduções (impostos)"]:
        fech.append(f"Impostos: {brl(abs(v['Deduções (impostos)']))} ({pct(p(v['Deduções (impostos)']), 1)} dos recebimentos).")
    fech.append(f"Custos: os custos operacionais somaram {brl(abs(v['Custos Operacionais']))}, representando "
                f"{pct(p(v['Custos Operacionais']))} dos recebimentos" + (f", principalmente {lista(custos)}." if custos else "."))
    fech.append(f"Resultado Bruto: após os custos, ficou um resultado de {brl(v['Resultado Bruto'])}, equivalente a "
                f"{pct(v['Resultado Bruto'] / ent if ent else 0)} dos recebimentos.")
    fech.append(f"Despesas Operacionais: totalizaram {brl(abs(v['Despesas Operacionais']))} ({pct(p(v['Despesas Operacionais']))})"
                + (f", com destaque para {lista(desp)}." if desp else "."))
    fech.append(f"Resultado Operacional: foi de {brl(v['Resultado Operacional'])}, representando "
                f"{pct(v['Resultado Operacional'] / ent if ent else 0)} do total recebido.")
    if round(v["Resultado Financeiro"], 2):
        fech.append(f"Resultado Financeiro: {brl(v['Resultado Financeiro'])} (rendimentos menos tarifas e juros).")
    if round(v["Não Operacional"], 2):
        nao_op = _grupos_secao(dre, "NAO_OPERACIONAL")
        fech.append(f"Não Operacional: {brl(v['Não Operacional'])} ({pct(p(v['Não Operacional']))}), referente a {lista(nao_op)}.")
    fech.append(f"Resultado Líquido: o mês fechou com {brl(v['Resultado Líquido'])}, ou seja, "
                f"{pct(v['Resultado Líquido'] / ent if ent else 0)} de tudo que entrou.")
    if saldos:
        fech.append(f"Caixa: o saldo {'das contas' if len(saldos) > 1 else 'da conta'} passou de {brl(s_ini)} para "
                    f"{brl(s_fim)} (variação de {brl(s_fim - s_ini)}).")

    pos, aten = [], []
    if ent:
        pos.append(f"Volume de recebimentos: {brl(ent)};")
    if v["Resultado Bruto"] > 0 and ent and v["Resultado Bruto"] / ent >= 0.3:
        pos.append(f"Resultado bruto de {brl(v['Resultado Bruto'])} ({pct(v['Resultado Bruto'] / ent)});")
    if v["Resultado Operacional"] > 0:
        pos.append(f"Resultado operacional positivo de {brl(v['Resultado Operacional'])} ({pct(v['Resultado Operacional'] / ent if ent else 0)});")
    if v["Resultado Líquido"] > 0:
        pos.append(f"Resultado líquido de {brl(v['Resultado Líquido'])}, com cerca de {pct(v['Resultado Líquido'] / ent if ent else 0)} dos recebimentos permanecendo como resultado;")
    if saldos and s_fim > s_ini:
        pos.append(f"Caixa cresceu {brl(s_fim - s_ini)} no mês;")
    if v["Deduções (impostos)"] and p(v["Deduções (impostos)"]) < 0.05:
        pos.append(f"Impostos baixos: {brl(abs(v['Deduções (impostos)']))} ({pct(p(v['Deduções (impostos)']), 1)} dos recebimentos);")
    if transf:
        env = sum(t["enviada"] for t in transf.values())
        recb = sum(t["recebida"] for t in transf.values())
        if round(env + recb, 2) == 0:
            pos.append(f"Transferências entre contas fecharam certinho: {brl(recb)} enviados e recebidos;")
        else:
            aten.append(f"Transferências entre contas não fecham: enviados {brl(abs(env))} x recebidos {brl(recb)};")

    if v["Resultado Líquido"] < 0:
        aten.append(f"Resultado líquido negativo de {brl(v['Resultado Líquido'])} — as saídas superaram as entradas no mês;")
    if v["Resultado Operacional"] < 0:
        aten.append(f"Operação no negativo: resultado operacional de {brl(v['Resultado Operacional'])};")
    if ent and p(v["Custos Operacionais"]) >= 0.45:
        aten.append(f"Custos operacionais representam {pct(p(v['Custos Operacionais']))} dos recebimentos"
                    + (f", principalmente {lista(custos[:2])};" if custos else ";"))
    for l in dre["linhas"]:
        if l["tipo"] != "grupo" or not ent:
            continue
        nome = l["descricao"].split("-", 1)[-1].strip()
        if l["grupo"].startswith("0607") and p(l["total"]) >= 0.03:
            aten.append(f"Gastos com sócios de {brl(abs(l['total']))} ({pct(p(l['total']), 1)} dos recebimentos);")
        elif l["grupo"].startswith("0609") and l["total"]:
            aten.append(f"Existem {brl(abs(l['total']))} em despesas sem classificação definida — revisar categorias;")
        elif l["secao"] == "NAO_OPERACIONAL" and l["total"] < 0 and p(l["total"]) >= 0.02:
            aten.append(f"{nome.capitalize()}: {brl(abs(l['total']))} — importante identificar a origem;")
    if itens_rec and ent and itens_rec[0][1] / ent >= 0.5 and len(itens_rec) > 1:
        aten.append(f"Recebimentos concentrados em {itens_rec[0][0]} ({pct(itens_rec[0][1] / ent)} das entradas);")
    if saldos and s_fim < s_ini:
        aten.append(f"Caixa reduziu {brl(s_ini - s_fim)} no mês;")
    if saldos and s_fim < 0:
        aten.append(f"Saldo final negativo: {brl(s_fim)};")
    if conf.get("dif_caixa"):
        aten.append(f"Conferência com o caixa apresenta diferença de {brl(conf['dif_caixa'])} — verificar lançamentos não conciliados;")

    if v["Resultado Líquido"] >= 0:
        concl = ("O principal ponto para os próximos meses é manter o volume de receitas e acompanhar de perto os custos "
                 "operacionais e as despesas, buscando preservar a margem de resultado.")
    else:
        concl = ("O principal ponto para os próximos meses é recompor o resultado: revisar os maiores custos e despesas, "
                 "acompanhar o fluxo de caixa semanalmente e avaliar o aumento das receitas.")
    return {
        "titulo": f"Fechamento Financeiro — {MESES[mes - 1]}/{ano} | {empresa['nome']}",
        "fechamento": "\n".join(fech),
        "positivos": "\n".join(pos) or "—",
        "atencao": "\n".join(aten) or "Nenhum ponto crítico identificado no mês.",
        "conclusao": concl,
    }
