"""Savian Contabilidade — Sistema de Fechamento Mensal (DRE financeira a partir do OMIE).

Rodar:  streamlit run app.py
"""
import copy
import io
import os
from datetime import date

import pandas as pd
import streamlit as st

from core import analise, apresentacao, db, dre, excel
from core.importador import ler_exportacao_omie
from core.omie import OmieCliente, OmieErro, baixar_mes, salvar_log
from core.plano import GRUPOS_PADRAO, IGNORAR, SECOES

st.set_page_config(page_title="Savian | Fechamento DRE", page_icon="📊", layout="wide")
ASSETS = os.path.join(os.path.dirname(__file__), "assets")
MESES = analise.MESES
SECOES_OPC = [s for s, _ in SECOES] + ["FORA"]
REL_OPC = ["", "recebimentos", "fornecedores", "funcionarios", "transferencias"]

st.markdown("""
<style>
 .block-container {padding-top: 1.5rem;}
 h1, h2, h3 {color: #0B2A7A;}
 div[data-testid="stMetricValue"] {font-size: 1.3rem;}
</style>""", unsafe_allow_html=True)

with st.sidebar:
    st.image(os.path.join(ASSETS, "logo_savian.png"), use_container_width=True)
    pagina = st.radio("Menu", ["Fechamento do mês", "Empresas", "Plano da DRE"], label_visibility="collapsed")
    st.caption("Report BPO Financeiro • dados locais em `dados/`")


def opcoes_grupos(cfg):
    return {g[0]: f"{g[0]} - {g[1]}" for g in dre.grupos_da_empresa(cfg)}


def brl(v):
    return analise.brl(v)


# =====================================================================
# EMPRESAS
# =====================================================================
def tela_empresas():
    st.title("Empresas cadastradas")
    empresas = db.listar()
    nomes = ["➕ Nova empresa"] + [f"{e['nome']} (#{e['id']})" for e in empresas]
    escolha = st.selectbox("Selecione para editar", nomes)
    emp = None if escolha.startswith("➕") else empresas[nomes.index(escolha) - 1]

    with st.form("form_empresa"):
        st.subheader("Dados e chaves da API OMIE")
        c1, c2 = st.columns(2)
        nome = c1.text_input("Nome da empresa (aparece no relatório)", value=emp["nome"] if emp else "")
        cnpj = c2.text_area("CNPJ(s) — um por linha", value=emp["cnpj"] if emp else "", height=70)
        c3, c4 = st.columns(2)
        app_key = c3.text_input("App Key", value=emp["app_key"] if emp else "")
        app_secret = c4.text_input("App Secret", value=emp["app_secret"] if emp else "", type="password")
        st.caption("No OMIE: Configurações › Integrações › Minhas aplicações (ou developer.omie.com.br › Meus aplicativos).")
        salvar = st.form_submit_button("💾 Salvar empresa", type="primary")
    if salvar:
        if not nome.strip():
            st.error("Informe o nome da empresa.")
        else:
            eid = db.salvar(nome.strip(), cnpj.strip(), app_key.strip(), app_secret.strip(), emp["id"] if emp else None)
            st.success(f"Empresa salva (#{eid}).")
            st.rerun()

    if not emp:
        return
    cfg = emp["config"]

    if st.button("🔌 Testar conexão com o OMIE"):
        try:
            with st.spinner("Consultando contas correntes..."):
                contas = OmieCliente(emp["app_key"], emp["app_secret"]).testar()
            st.success(f"Conexão OK — {len(contas)} conta(s) corrente(s) encontrada(s).")
            st.session_state[f"contas_{emp['id']}"] = contas
        except OmieErro as e:
            st.error(str(e))

    st.subheader("Configuração do relatório")
    c1, c2 = st.columns(2)
    modos = {"conta": "Por conta corrente (ex.: MS Festas)", "projeto": "Por projeto / centro de custo (ex.: Hagili)",
             "total": "Somente total da empresa"}
    modo = c1.radio("Colunas da DRE", list(modos), format_func=modos.get,
                    index=list(modos).index(cfg.get("modo_colunas", "conta")))
    so_real = c2.checkbox("Considerar só lançamentos realizados (ignora atrasados / a vencer)",
                          value=cfg.get("apenas_realizado", True))
    rec_cli = c2.checkbox("Slide de recebimentos por cliente (em vez de por categoria)",
                          value=cfg.get("recebimentos_por_cliente", False))

    contas = st.session_state.get(f"contas_{emp['id']}")
    if contas:
        st.markdown("**Contas correntes que entram no fechamento** (vazio = todas)")
        opc = {str(c["nCodCC"]): f"{c.get('descricao')} ({c['nCodCC']})" for c in contas}
        sel = st.multiselect("Contas", list(opc), default=[x for x in cfg.get("contas_ativas", []) if x in opc],
                             format_func=opc.get, label_visibility="collapsed")
    else:
        sel = cfg.get("contas_ativas", [])
        st.caption("Clique em “Testar conexão” para escolher quais contas correntes entram no fechamento.")

    st.markdown("**Nomes das colunas no relatório e agrupamentos** (ex.: juntar contas em TOTAL MATRIZ / TOTAL FILIAL)")
    chave_ap = "apelidos_projeto" if modo == "projeto" else "apelidos_conta"
    conhecidos = st.session_state.get(f"colunas_vistas_{emp['id']}", [])
    linhas = []
    for orig in sorted(set(conhecidos) | set(cfg.get(chave_ap, {}))):
        novo = cfg.get(chave_ap, {}).get(orig, orig)
        linhas.append({"Nome no OMIE": orig, "Nome no relatório": novo,
                       "Agrupamento (opcional)": cfg.get("grupos_colunas", {}).get(novo, ""),
                       "Ordem": (cfg.get("ordem_colunas") or []).index(novo) + 1 if novo in (cfg.get("ordem_colunas") or []) else None})
    df = pd.DataFrame(linhas or [{"Nome no OMIE": "", "Nome no relatório": "", "Agrupamento (opcional)": "", "Ordem": None}])
    df = st.data_editor(df, num_rows="dynamic", use_container_width=True, key=f"ap_{emp['id']}_{modo}",
                        column_config={"Ordem": st.column_config.NumberColumn(min_value=1, step=1)})
    st.caption("Dica: os nomes do OMIE aparecem aqui depois do primeiro fechamento. Dois nomes do OMIE com o mesmo "
               "“Nome no relatório” viram uma coluna só (ex.: OGGI SORVETES + SO GELO → OGGI / SÓ GELO).")

    if st.button("💾 Salvar configuração", type="primary"):
        apel, grp, ordem = {}, {}, []
        for _, r in df.dropna(subset=["Nome no OMIE"]).iterrows():
            o, n = str(r["Nome no OMIE"]).strip(), str(r["Nome no relatório"] or "").strip()
            if not o:
                continue
            n = n or o
            if n != o:
                apel[o] = n
            if str(r.get("Agrupamento (opcional)") or "").strip() not in ("", "nan", "None"):
                grp[n] = str(r["Agrupamento (opcional)"]).strip()
            if pd.notna(r.get("Ordem")):
                ordem.append((int(r["Ordem"]), n))
        cfg.update({"modo_colunas": modo, "apenas_realizado": so_real, "recebimentos_por_cliente": rec_cli,
                    "contas_ativas": sel, chave_ap: apel, "grupos_colunas": grp,
                    "ordem_colunas": [n for _, n in sorted(ordem)]})
        db.salvar_config(emp["id"], cfg)
        st.success("Configuração salva.")

    with st.expander("🗑️ Excluir empresa"):
        if st.button(f"Excluir {emp['nome']}", type="secondary"):
            db.excluir(emp["id"])
            st.rerun()


# =====================================================================
# PLANO DA DRE
# =====================================================================
def tela_plano():
    st.title("Plano da DRE (grupos)")
    empresas = db.listar()
    if not empresas:
        st.info("Cadastre uma empresa primeiro.")
        return
    emp = st.selectbox("Empresa", empresas, format_func=lambda e: e["nome"])
    cfg = emp["config"]
    st.caption("Os grupos são o nível 2 da DRE (ex.: 040100-MÃO DE OBRA OPERACIONAL). As contas de nível 3 são as "
               "categorias do OMIE, classificadas na tela de fechamento. A coluna “Relatório” define em qual slide os "
               "pagamentos do grupo aparecem.")
    df = pd.DataFrame(dre.grupos_da_empresa(cfg), columns=["Código", "Nome", "Seção", "Relatório"])
    df = st.data_editor(df, num_rows="dynamic", use_container_width=True, key=f"plano_{emp['id']}", column_config={
        "Seção": st.column_config.SelectboxColumn(options=SECOES_OPC, required=True),
        "Relatório": st.column_config.SelectboxColumn(options=REL_OPC),
    })
    c1, c2 = st.columns(2)
    if c1.button("💾 Salvar plano", type="primary"):
        gs = [[str(r["Código"]).strip(), str(r["Nome"]).strip(), r["Seção"], r["Relatório"] or ""]
              for _, r in df.iterrows() if str(r["Código"]).strip() not in ("", "nan", "None")]
        gs.sort(key=lambda g: (g[0] == IGNORAR, g[0]))
        cfg["grupos"] = gs
        db.salvar_config(emp["id"], cfg)
        st.success("Plano salvo.")
    if c2.button("↩️ Voltar ao plano padrão"):
        cfg["grupos"] = None
        db.salvar_config(emp["id"], cfg)
        st.rerun()

    st.subheader("Copiar classificação de outra empresa")
    outras = [e for e in empresas if e["id"] != emp["id"] and e["mapa"]]
    if outras:
        src = st.selectbox("Origem", outras, format_func=lambda e: f"{e['nome']} ({len(e['mapa'])} categorias)")
        if st.button("Copiar categorias que ainda não existem nesta empresa"):
            mapa = emp["mapa"]
            n = 0
            for k, v in src["mapa"].items():
                if k not in mapa:
                    mapa[k] = dict(v, automatico=False)
                    n += 1
            db.salvar_mapa(emp["id"], mapa)
            st.success(f"{n} categoria(s) copiada(s).")
    else:
        st.caption("Nenhuma outra empresa com classificação salva.")


# =====================================================================
# FECHAMENTO
# =====================================================================
def tela_fechamento():
    st.title("Fechamento do mês")
    empresas = db.listar()
    if not empresas:
        st.info("Nenhuma empresa cadastrada. Vá em **Empresas** e cadastre a primeira (nome + App Key/App Secret do OMIE).")
        return
    hoje = date.today()
    pm = hoje.month - 1 or 12
    pa = hoje.year if hoje.month > 1 else hoje.year - 1
    c1, c2, c3, c4 = st.columns([3, 2, 1.2, 2])
    emp = c1.selectbox("Empresa", empresas, format_func=lambda e: e["nome"])
    mes = c2.selectbox("Mês", list(range(1, 13)), index=pm - 1, format_func=lambda m: MESES[m - 1])
    ano = c3.number_input("Ano", 2015, 2100, pa)
    fonte = c4.radio("Fonte", ["API OMIE", "Planilha exportada do OMIE"], horizontal=False)

    arquivo = None
    if fonte.startswith("Planilha"):
        arquivo = st.file_uploader("Finanças › Movimentação da Conta Corrente › Exportar para Excel", type=["xlsx"])

    chave = f"dados_{emp['id']}_{ano}_{mes}"
    if st.button("⬇️ Puxar dados e montar DRE", type="primary"):
        try:
            if fonte == "API OMIE":
                if not emp["app_key"] or not emp["app_secret"]:
                    st.error("Cadastre a App Key e o App Secret desta empresa.")
                    return
                log = []
                barra = st.progress(0.0, "Conectando ao OMIE...")
                cli = OmieCliente(emp["app_key"], emp["app_secret"], log=log)
                r = baixar_mes(cli, int(ano), int(mes), emp["config"].get("contas_ativas"),
                               progresso=lambda p, t: barra.progress(p, t))
                pasta = os.path.join(db.PASTA_DADOS, "logs")
                os.makedirs(pasta, exist_ok=True)
                salvar_log(log, os.path.join(pasta, f"omie_{emp['id']}_{ano}_{mes:02d}.json"))
                st.session_state[f"contas_{emp['id']}"] = r["contas"]
                lancs, saldos = r["lancamentos"], r["saldos"]
            else:
                if not arquivo:
                    st.error("Envie a planilha exportada do OMIE.")
                    return
                lancs, saldos = ler_exportacao_omie(arquivo)
            st.session_state[chave] = {"lancs": lancs, "saldos": saldos}
        except (OmieErro, ValueError) as e:
            st.error(str(e))
            return

    dados = st.session_state.get(chave)
    if not dados:
        st.info("Escolha a empresa e o mês e clique em **Puxar dados**.")
        return

    cfg = copy.deepcopy(emp["config"])
    mapa = emp["mapa"]
    novas = dre.completar_mapeamento(mapa, dados["lancs"])
    if novas:
        db.salvar_mapa(emp["id"], mapa)
    cfg["rateios_mes"] = cfg.get("rateios", {}).get(f"{ano}-{mes:02d}", [])
    saldos = dados["saldos"]

    base = dre.montar_base(dados["lancs"], cfg, mapa)
    d = dre.montar_dre(base, cfg)
    vistas = sorted({(l.get("projeto") if cfg["modo_colunas"] == "projeto" else l.get("conta_corrente")) or ""
                     for l in dados["lancs"]} - {""})
    st.session_state[f"colunas_vistas_{emp['id']}"] = vistas
    conf = dre.conferencia(base, d, saldos)
    ind = dre.indicadores(d)

    nao_rev = [k for k in categorias_mes(dados["lancs"]) if mapa.get(k, {}).get("automatico")]
    a_class = [b for b in base if b["grupo"].startswith("0609")]
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Lançamentos", len(base))
    m2.metric("Entradas", brl(ind["Entradas"]["valor"]))
    m3.metric("Resultado líquido", brl(ind["Resultado Líquido"]["valor"]), f"{ind['Resultado Líquido']['pct'] * 100:.1f}%")
    m4.metric("Variação de caixa", brl(conf["variacao_caixa"]) if conf["variacao_caixa"] is not None else "—")
    if conf["dif_caixa"] in (None,):
        m5.metric("Conferência caixa", "sem saldos")
    else:
        m5.metric("Conferência caixa", "OK ✅" if abs(conf["dif_caixa"]) < 0.01 else brl(conf["dif_caixa"]))
    if nao_rev:
        st.warning(f"{len(nao_rev)} categoria(s) foram classificadas automaticamente e ainda não foram revisadas. "
                   "Confira na aba **1. Classificação**.")
    if a_class:
        st.warning(f"{len(a_class)} lançamento(s) caíram em OUTRAS DESPESAS (A CLASSIFICAR).")

    abas = st.tabs(["1. Classificação", "2. Rateios", "3. Saldos", "4. DRE", "5. Base", "6. Análise e arquivos"])
    with abas[0]:
        aba_classificacao(emp, cfg, mapa, dados["lancs"])
    with abas[1]:
        aba_rateios(emp, cfg, d, ano, mes)
    with abas[2]:
        aba_saldos(chave, saldos)
    with abas[3]:
        st.dataframe(dre_df(d), use_container_width=True, height=700, hide_index=True)
    with abas[4]:
        dfb = pd.DataFrame(base)
        st.dataframe(dfb, use_container_width=True, height=600, hide_index=True)
    with abas[5]:
        aba_arquivos(emp, cfg, base, d, saldos, int(ano), int(mes))


def categorias_mes(lancs):
    return dre.categorias_do_mes(lancs)


def aba_classificacao(emp, cfg, mapa, lancs):
    st.caption("Cada categoria do OMIE vira uma conta (nível 3) dentro de um grupo da DRE. A classificação fica salva "
               "para a empresa e é reaproveitada nos próximos meses. Categorias com o mesmo “Nome no relatório” e "
               "grupo são somadas na mesma linha.")
    cats = categorias_mes(lancs)
    opc = opcoes_grupos(cfg)
    linhas = []
    for k, info in cats.items():
        m = mapa.get(k, {})
        linhas.append({"chave": k, "Categoria no OMIE": info["categoria"], "Valor no mês": round(info["valor"], 2),
                       "Grupo da DRE": opc.get(m.get("grupo"), opc.get("060900")),
                       "Nome no relatório": m.get("conta", info["categoria"]),
                       "Revisar": "⚠️" if m.get("automatico") else ""})
    df = pd.DataFrame(linhas).sort_values(["Grupo da DRE", "Categoria no OMIE"])
    ed = st.data_editor(df, use_container_width=True, hide_index=True, height=min(40 + 36 * len(df), 700),
                        key=f"cls_{emp['id']}", disabled=["chave", "Categoria no OMIE", "Valor no mês", "Revisar"],
                        column_config={
                            "chave": None,
                            "Valor no mês": st.column_config.NumberColumn(format="%.2f"),
                            "Grupo da DRE": st.column_config.SelectboxColumn(options=list(opc.values()), required=True,
                                                                             width="large"),
                        })
    if st.button("💾 Salvar classificação e recalcular", type="primary"):
        inv = {v: k for k, v in opc.items()}
        for _, r in ed.iterrows():
            mapa[r["chave"]] = {"categoria": r["Categoria no OMIE"], "grupo": inv.get(r["Grupo da DRE"], "060900"),
                                "conta": str(r["Nome no relatório"] or r["Categoria no OMIE"]).strip(),
                                "automatico": False}
        db.salvar_mapa(emp["id"], mapa)
        st.rerun()


def aba_rateios(emp, cfg, d, ano, mes):
    st.caption("Rateio manual: move um valor de uma conta entre colunas (ex.: combustível que veio “sem centro de custo” "
               "para cada posto). Informe o valor positivo; o sinal segue a conta. Fica salvo para este mês.")
    contas = [l["descricao"] for l in d["linhas"] if l["tipo"] == "conta"]
    cols = d["colunas"] + ([dre.SEM_CC] if dre.SEM_CC not in d["colunas"] else [])
    chave = f"{ano}-{mes:02d}"
    atuais = emp["config"].get("rateios", {}).get(chave, [])
    df = pd.DataFrame(atuais or [], columns=["descricao", "conta", "de", "para", "valor"])
    ed = st.data_editor(df, num_rows="dynamic", use_container_width=True, key=f"rat_{emp['id']}_{chave}",
                        column_config={
                            "descricao": st.column_config.TextColumn("Descrição"),
                            "conta": st.column_config.SelectboxColumn("Conta", options=contas, width="large"),
                            "de": st.column_config.SelectboxColumn("De (coluna)", options=cols),
                            "para": st.column_config.SelectboxColumn("Para (coluna)", options=cols),
                            "valor": st.column_config.NumberColumn("Valor", format="%.2f", min_value=0.0),
                        })
    if st.button("💾 Salvar rateios do mês"):
        conf = emp["config"]
        conf.setdefault("rateios", {})[chave] = [
            {k: (None if pd.isna(v) else v) for k, v in r.items()} for r in ed.to_dict("records")
            if r.get("conta") and r.get("para") and r.get("valor")]
        db.salvar_config(emp["id"], conf)
        st.rerun()


def aba_saldos(chave, saldos):
    st.caption("Saldos inicial e final de cada conta (vêm do extrato do OMIE). Ajuste se necessário — usados no slide "
               "de saldos e na conferência com o caixa.")
    df = pd.DataFrame(saldos or [], columns=["conta", "saldo_inicial", "saldo_final"])
    ed = st.data_editor(df[["conta", "saldo_inicial", "saldo_final"]], num_rows="dynamic", use_container_width=True,
                        key=f"sal_{chave}", column_config={
                            "conta": "Conta", "saldo_inicial": st.column_config.NumberColumn("Saldo inicial", format="%.2f"),
                            "saldo_final": st.column_config.NumberColumn("Saldo final", format="%.2f")})
    if st.button("💾 Aplicar saldos"):
        st.session_state[chave]["saldos"] = [
            {"conta": r["conta"], "saldo_inicial": float(r["saldo_inicial"] or 0), "saldo_final": float(r["saldo_final"] or 0)}
            for r in ed.to_dict("records") if r.get("conta")]
        st.rerun()


def dre_df(d):
    rows = []
    for l in d["linhas"]:
        ind = {1: "", 2: "   ", 3: "        "}[l["nivel"]]
        r = {"Descrição da Conta": ind + l["descricao"]}
        for c in d["colunas"]:
            r[c] = l["valores"].get(c, 0.0)
        for g, membros in d["agrupamentos"].items():
            r[f"TOTAL {g}"] = sum(l["valores"].get(m, 0.0) for m in membros)
        r["TOTAL"] = l["total"]
        rows.append(r)
    df = pd.DataFrame(rows)
    fmt = {c: (lambda v: f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")) for c in df.columns[1:]}
    return df.style.format(fmt)


def aba_arquivos(emp, cfg, base, d, saldos, ano, mes):
    textos = analise.gerar_textos(emp, ano, mes, base, d, saldos, cfg)
    st.caption("Textos gerados automaticamente a partir dos números. Edite à vontade antes de gerar a apresentação.")
    k = f"txt_{emp['id']}_{ano}_{mes}"
    titulo = st.text_input("Título da análise", textos["titulo"], key=k + "t")
    fech = st.text_area("Fechamento do mês (um item por linha)", textos["fechamento"], height=260, key=k + "f")
    c1, c2 = st.columns(2)
    pos = c1.text_area("Pontos positivos", textos["positivos"], height=200, key=k + "p")
    ate = c2.text_area("Pontos de atenção", textos["atencao"], height=200, key=k + "a")
    concl = st.text_area("Conclusão", textos["conclusao"], height=80, key=k + "c")
    detal = st.checkbox("Incluir slide com a DRE detalhada (contas de nível 3)", value=True)

    if st.button("📦 Gerar Excel e PowerPoint", type="primary"):
        pasta = os.path.join(db.PASTA_DADOS, "saidas", emp["nome"].replace("/", "-"), f"{ano}-{mes:02d}")
        os.makedirs(pasta, exist_ok=True)
        base_nome = f"FECHAMENTO {emp['nome']} {mes:02d}_{ano}".replace("/", "-")
        xlsx = os.path.join(pasta, base_nome + ".xlsx")
        pptx = os.path.join(pasta, f"{emp['nome']} REPORT {mes:02d}_{ano}.pptx".replace("/", "-"))
        with st.spinner("Gerando arquivos..."):
            excel.gerar_excel(xlsx, emp, ano, mes, base, d, saldos, cfg)
            apresentacao.gerar_pptx(pptx, emp, ano, mes, base, d, saldos, cfg,
                                    {"titulo": titulo, "fechamento": fech, "positivos": pos, "atencao": ate,
                                     "conclusao": concl}, incluir_detalhada=detal)
        st.session_state[k + "arqs"] = (xlsx, pptx)
        st.success(f"Arquivos salvos em `{pasta}`")
    arqs = st.session_state.get(k + "arqs")
    if arqs:
        c1, c2 = st.columns(2)
        with open(arqs[0], "rb") as f:
            c1.download_button("⬇️ Baixar Excel (DRE)", f.read(), file_name=os.path.basename(arqs[0]),
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with open(arqs[1], "rb") as f:
            c2.download_button("⬇️ Baixar PowerPoint", f.read(), file_name=os.path.basename(arqs[1]),
                               mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")


if pagina == "Empresas":
    tela_empresas()
elif pagina == "Plano da DRE":
    tela_plano()
else:
    tela_fechamento()
