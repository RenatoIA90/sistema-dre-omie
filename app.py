"""Savian Contabilidade — Sistema de Fechamento Mensal (DRE financeira a partir do OMIE).

Rodar:  streamlit run app.py
"""
import copy
import getpass
import html
import io
import os
from datetime import date

import pandas as pd
import streamlit as st

from core import analise, apresentacao, db, dre, excel, fluxo, modelos, recebimento
from core.importador import ler_exportacao_omie
from core.omie import OmieCliente, OmieErro, baixar_mes, salvar_log
from core.plano import GRUPOS_PADRAO, IGNORAR, SECOES

st.set_page_config(page_title="Savian | Reports BPO", page_icon="📊", layout="wide")
ASSETS = os.path.join(os.path.dirname(__file__), "assets")
MESES = analise.MESES
SECOES_OPC = [s for s, _ in SECOES] + ["FORA"]
REL_OPC = ["", "recebimentos", "fornecedores", "funcionarios", "transferencias"]
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
MIME_PPTX = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
REL_RECEB = "Relatório Recebimento"

AZUL = "#213A6B"
st.markdown(f"""
<style>
 .stApp {{background: #F5F6F8;}}
 header[data-testid="stHeader"] {{background: transparent;}}
 .block-container {{padding-top: 1rem; max-width: 1400px;}}
 h1, h2, h3 {{color: #1A1A1A;}}
 div[data-testid="stMetricValue"] {{font-size: 1.3rem;}}

 /* barra lateral azul, no visual do portal Savian */
 section[data-testid="stSidebar"] {{background: {AZUL}; min-width: 240px !important; max-width: 240px !important;}}
 section[data-testid="stSidebar"] * {{color: #FFFFFF;}}
 section[data-testid="stSidebar"] hr {{border-color: rgba(255,255,255,.25); margin: .6rem 0 1rem;}}
 .menu-secao {{font-size: .68rem; letter-spacing: .06em; opacity: .75; margin: .2rem 0 .3rem .1rem;}}
 section[data-testid="stSidebar"] .stButton button {{
   width: 100%; justify-content: flex-start; border: 0; border-radius: 4px; padding: .45rem .7rem;
   background: transparent; font-weight: 500; box-shadow: none;}}
 section[data-testid="stSidebar"] .stButton button > div {{justify-content: flex-start; width: 100%;}}
 section[data-testid="stSidebar"] .stButton button:hover {{background: rgba(255,255,255,.08);}}
 section[data-testid="stSidebar"] .stButton button[kind="primary"] {{background: rgba(255,255,255,.16);}}
 .rodape-menu {{font-size: .72rem; opacity: .6; margin-top: 2rem;}}

 /* cabeçalho da página e cartões brancos */
 .topo {{display: flex; align-items: center; gap: .8rem; background: #FFFFFF; border-bottom: 1px solid #E3E6EA;
   margin: -1rem -1rem 1.4rem; padding: .85rem 1.2rem; font-weight: 600; color: {AZUL}; font-size: 1.05rem;}}
 .topo .usuario {{margin-left: auto; font-weight: 400; color: #6B7280; font-size: .9rem;}}
 div[class*="st-key-cartao"] {{background: #FFFFFF; border: 0 !important; border-radius: 6px;
   box-shadow: 0 1px 3px rgba(16,24,40,.08); padding: 1rem 1.1rem;}}
 div[data-testid="stForm"] {{background: #FFFFFF; border: 0; border-radius: 6px;
   box-shadow: 0 1px 3px rgba(16,24,40,.08);}}
 .tabela-cab {{font-weight: 600; font-size: .9rem; border-bottom: 1px solid #E3E6EA; padding-bottom: .4rem;}}
 .cinza {{color: #6B7280;}}
 div[class*="st-key-linkbtn"] button {{background: transparent; border: 0; padding: 0; min-height: 0;
   color: {AZUL}; text-decoration: underline; font-size: .85rem; box-shadow: none;}}
 div[class*="st-key-linkbtn"] button:hover {{color: #0B2A7A; background: transparent;}}
 div[class*="st-key-linkbtn"] button p {{font-size: .85rem;}}

 /* mensagens de validação do cadastro */
 .alerta-erro {{display: flex; gap: .8rem; align-items: flex-start; background: #FEF3F2; border: 1px solid #FECDCA;
   border-left: 5px solid #D92D20; border-radius: 6px; padding: .85rem 1rem; margin: .4rem 0 1rem; color: #7A271A;}}
 .alerta-erro .icone {{background: #D92D20; color: #FFFFFF; border-radius: 50%; min-width: 1.6rem; height: 1.6rem;
   display: flex; align-items: center; justify-content: center; font-weight: 700;}}
 .alerta-erro b {{color: #B42318;}}
 .alerta-erro ul {{margin: .3rem 0 0 1rem; padding: 0;}}
 .alerta-erro li {{margin: .1rem 0;}}
 /* botões de excluir (vermelhos) */
 div[class*="st-key-perigo"] button {{background: #D92D20; border-color: #D92D20; color: #FFFFFF;}}
 div[class*="st-key-perigo"] button:hover {{background: #B42318; border-color: #B42318; color: #FFFFFF;}}
 div[class*="st-key-xemp"] button {{border: 1px solid #FECDCA; color: #D92D20; background: #FFFFFF; font-weight: 700;}}
 div[class*="st-key-xemp"] button:hover {{background: #FEF3F2; border-color: #D92D20; color: #B42318;}}
</style>""", unsafe_allow_html=True)

PAGINAS = {"bpo": ("Cadastrar Empresas", ":material/apartment:"), "report": ("Report", ":material/bar_chart:")}
st.session_state.setdefault("pagina", "report")

with st.sidebar:
    st.image(os.path.join(ASSETS, "logo_savian_branco.png"), width=170)
    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown('<div class="menu-secao">BPO</div>', unsafe_allow_html=True)
    for chave_pg, (rotulo, icone) in PAGINAS.items():
        if st.button(rotulo, icon=icone, key=f"menu_{chave_pg}", width="stretch",
                     type="primary" if st.session_state["pagina"] == chave_pg else "secondary"):
            st.session_state["pagina"] = chave_pg
            st.rerun()
    st.markdown('<div class="rodape-menu">Savian Contabilidade</div>', unsafe_allow_html=True)


def contas_conhecidas(emp):
    """Contas correntes do OMIE guardadas no cadastro da empresa (atualizadas a cada consulta)."""
    return emp["config"].get("contas_omie") or []


def guardar_contas(emp, contas):
    lista = [{"nCodCC": str(c["nCodCC"]), "descricao": html.unescape(str(c.get("descricao") or c["nCodCC"]))}
             for c in contas if str(c.get("inativo", "N")).upper() != "S"]
    if lista != contas_conhecidas(emp):
        emp["config"]["contas_omie"] = lista
        db.salvar_config(emp["id"], emp["config"])


def ir_para(pagina):
    st.session_state["pagina"] = pagina


def topo(titulo):
    st.markdown(f'<div class="topo">☰&nbsp; {titulo}<span class="usuario">{getpass.getuser()}</span></div>',
                unsafe_allow_html=True)


def opcoes_grupos(cfg):
    return {g[0]: f"{g[0]} - {g[1]}" for g in dre.grupos_da_empresa(cfg)}


def brl(v):
    return analise.brl(v)


def alerta_erro(titulo, itens):
    """Caixa vermelha de erro com a lista do que precisa ser corrigido."""
    lis = "".join(f"<li>{html.escape(i)}</li>" for i in itens)
    st.markdown(f'<div class="alerta-erro"><div class="icone">!</div><div><b>{html.escape(titulo)}</b>'
                f'<ul>{lis}</ul></div></div>', unsafe_allow_html=True)


def cnpj_valido(c):
    """Confere os dígitos verificadores de um CNPJ (só números)."""
    if len(c) != 14 or len(set(c)) == 1:
        return False
    for n in (12, 13):
        pesos = list(range(n - 7, 1, -1)) + list(range(9, 1, -1))
        dv = sum(int(d) * p for d, p in zip(c[:n], pesos)) % 11
        if int(c[n]) != (0 if dv < 2 else 11 - dv):
            return False
    return True


def formatar_cnpj(c):
    return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}"


def avisos_pendentes():
    """Mensagem guardada antes de um st.rerun (ex.: empresa excluída)."""
    msg = st.session_state.pop("aviso_ok", None)
    if msg:
        st.toast(msg, icon="✅")


@st.dialog("Excluir empresa")
def dialogo_excluir(emp):
    st.markdown(f"Tem certeza que deseja excluir **{html.escape(emp['nome'])}**?")
    st.caption("A empresa sai do cadastro e das listas do Report. Os arquivos já gerados em "
               "Documentos › Savian DRE continuam salvos.")
    c1, c2 = st.columns(2)
    if c1.button("Sim, excluir", key="perigo_confirma", width="stretch"):
        db.excluir(emp["id"])
        st.session_state["cad_emp_proxima"] = "➕ Nova empresa"
        st.session_state["rep_emp_proxima"] = None
        st.session_state["aviso_ok"] = f"Empresa {emp['nome']} excluída."
        st.rerun()
    if c2.button("Cancelar", width="stretch"):
        st.rerun()


# =====================================================================
# EMPRESAS
# =====================================================================
def tela_bpo():
    topo("Cadastrar Empresas")
    st.subheader("Cadastro de empresas")
    abas = st.tabs(["Cadastro de empresas", "Plano da DRE"])
    with abas[0]:
        tela_empresas()
    with abas[1]:
        tela_plano()


def validar_empresa(nome, cnpj, app_key, app_secret, eid=None):
    """(campos não preenchidos, outros erros) do formulário de empresa."""
    faltando = [rot for rot, v in (("Nome da empresa", nome), ("CNPJ", cnpj), ("App Key", app_key),
                                   ("App Secret", app_secret)) if not v.strip()]
    erros = []
    lista = db.cnpjs(cnpj)
    invalidos = [c for c in lista if not cnpj_valido(c)]
    if invalidos:
        erros.append("CNPJ inválido: " + ", ".join(c if len(c) != 14 else formatar_cnpj(c) for c in invalidos)
                     + " (confira os 14 números).")
    repetidos = {c for c in lista if lista.count(c) > 1}
    if repetidos:
        erros.append("CNPJ repetido no próprio cadastro: " + ", ".join(formatar_cnpj(c) for c in repetidos) + ".")
    for c, outra in db.cnpj_em_uso(lista, eid).items():
        erros.append(f"O CNPJ {formatar_cnpj(c)} já está cadastrado na empresa {outra}. "
                     "Não é permitido cadastrar o mesmo CNPJ duas vezes.")
    return faltando, erros


def tela_empresas():
    avisos_pendentes()
    empresas = db.listar()
    nomes = ["➕ Nova empresa"] + [f"{e['nome']} (#{e['id']})" for e in empresas]
    if "cad_emp_proxima" in st.session_state:
        st.session_state["cad_emp"] = st.session_state.pop("cad_emp_proxima")
    if st.session_state.get("cad_emp") not in nomes:
        st.session_state["cad_emp"] = nomes[0]
    escolha = st.selectbox("Selecione para editar", nomes, key="cad_emp")
    emp = None if escolha.startswith("➕") else empresas[nomes.index(escolha) - 1]

    with st.form("form_empresa"):
        st.markdown("**Dados e chaves da API OMIE** — todos os campos são obrigatórios")
        c1, c2 = st.columns(2)
        nome = c1.text_input("Nome da empresa (aparece no relatório) *", value=emp["nome"] if emp else "")
        cnpj = c2.text_area("CNPJ(s) — um por linha *", value=emp["cnpj"] if emp else "", height=70)
        c3, c4 = st.columns(2)
        app_key = c3.text_input("App Key *", value=emp["app_key"] if emp else "")
        app_secret = c4.text_input("App Secret *", value=emp["app_secret"] if emp else "", type="password")
        ramo = None
        if not emp:
            opcoes = todos_modelos()
            ramo = st.selectbox("Ramo da empresa (modelo de DRE) *", list(opcoes),
                                help="Define os grupos da DRE. Facilities usa o modelo da HAGILI. Dá para trocar "
                                     "depois em “Modelo de DRE”.")
        st.caption("No OMIE: Configurações › Integrações › Minhas aplicações (ou developer.omie.com.br › Meus aplicativos).")
        salvar = st.form_submit_button("💾 Salvar empresa", type="primary")
    if salvar:
        faltando, erros = validar_empresa(nome, cnpj, app_key, app_secret, emp["id"] if emp else None)
        if faltando:
            alerta_erro("Existe campo obrigatório não preenchido" if len(faltando) == 1 else
                        "Existem campos obrigatórios não preenchidos",
                        [f"Preencha o campo “{f}”." for f in faltando])
        if erros:
            alerta_erro("Não foi possível salvar a empresa", erros)
        if not faltando and not erros:
            cnpj_fmt = "\n".join(formatar_cnpj(c) for c in db.cnpjs(cnpj))
            eid = db.salvar(nome.strip(), cnpj_fmt, app_key.strip(), app_secret.strip(), emp["id"] if emp else None)
            if ramo:
                aplicar_modelo(db.obter(eid), ramo, todos_modelos()[ramo])
            st.session_state["aviso_ok"] = f"Empresa {nome.strip()} salva."
            st.session_state["cad_emp_proxima"] = f"{nome.strip()} (#{eid})"
            st.rerun()

    if not emp:
        return
    cfg = emp["config"]

    b1, _, b2 = st.columns([2, 4, 1.6])
    if b2.button("Excluir empresa", icon=":material/delete:", key="perigo_excluir_cad", width="stretch"):
        dialogo_excluir(emp)
    if b1.button("🔌 Testar conexão com o OMIE"):
        try:
            with st.spinner("Consultando contas correntes..."):
                cli = OmieCliente(emp["app_key"], emp["app_secret"])
                contas = cli.testar()
                dados_omie = cli.empresa()
            st.success(f"Conexão OK — {dados_omie.get('razao_social', '')} ({dados_omie.get('cnpj', '')}): "
                       f"{len(contas)} conta(s) corrente(s) encontrada(s).")
            cnpj_omie = db.so_digitos(dados_omie.get("cnpj"))
            if not emp["cnpj"] and cnpj_omie:
                outra = db.cnpj_em_uso([cnpj_omie], emp["id"]).get(cnpj_omie)
                if outra:
                    alerta_erro("CNPJ já cadastrado", [f"O CNPJ do OMIE ({formatar_cnpj(cnpj_omie)}) já está na "
                                                       f"empresa {outra}. Confira se as chaves são desta empresa."])
                else:
                    db.salvar(emp["nome"], formatar_cnpj(cnpj_omie), emp["app_key"], emp["app_secret"], emp["id"])
                    st.info("CNPJ preenchido com o cadastro do OMIE.")
            guardar_contas(emp, contas)
        except OmieErro as e:
            st.error(str(e))

    secao_modelo_dre(emp)

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

    contas = contas_conhecidas(emp)
    if contas:
        st.markdown("**Contas correntes que entram no fechamento** (vazio = todas)")
        opc = {c["nCodCC"]: f"{c['descricao']} ({c['nCodCC']})" for c in contas}
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
    df = st.data_editor(df, num_rows="dynamic", width="stretch", key=f"ap_{emp['id']}_{modo}",
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


def todos_modelos():
    return {**{f"{n} (pronto)": m for n, m in modelos.MODELOS_PRONTOS.items()},
            **{f"{n} (importado)": m for n, m in db.listar_modelos().items()}}


def aplicar_modelo(emp, nome, modelo):
    cfg = emp["config"]
    cfg["grupos"] = modelos.completar_grupos(modelo["grupos"])
    cfg["contas_modelo"] = [list(c) for c in modelo["contas"]]
    cfg["modelo_dre"] = nome
    cfg["grupos_criados"] = []
    mapa = emp["mapa"]
    dre.reclassificar(mapa, cfg)
    db.salvar_config(emp["id"], cfg)
    db.salvar_mapa(emp["id"], mapa)
    st.session_state.pop(f"modelo_sel_{emp['id']}", None)


def secao_modelo_dre(emp):
    st.subheader("Modelo de DRE")
    cfg = emp["config"]
    atual = modelos.NOMES_ANTIGOS.get(cfg.get("modelo_dre"), cfg.get("modelo_dre")) or f"{modelos.MODELO_PADRAO} (pronto)"
    st.caption(f"Modelo em uso: **{atual}**. O modelo define os grupos da DRE conforme o ramo da empresa "
               "(Facilities = modelo HAGILI, casa noturna, artistas...) e ajuda a classificar as categorias do OMIE. "
               "Para trocar, escolha outro modelo e clique em Aplicar. Classificações que você já revisou continuam "
               "valendo se o grupo existir no novo modelo.")
    opcoes = todos_modelos()
    c1, c2, c3 = st.columns([3, 1.3, 1.3], vertical_alignment="bottom")
    nomes = list(opcoes)
    escolha = c1.selectbox("Modelos disponíveis", nomes, index=nomes.index(atual) if atual in nomes else 0,
                           key=f"modelo_sel_{emp['id']}")
    if c2.button("Aplicar à empresa", key=f"modelo_apl_{emp['id']}", width="stretch"):
        aplicar_modelo(emp, escolha, opcoes[escolha])
        st.rerun()
    buf = io.BytesIO()
    m = opcoes[escolha]
    modelos.exportar_excel_modelo(buf, escolha, modelos.completar_grupos(m["grupos"]), m["contas"])
    c3.download_button("Baixar em Excel", buf.getvalue(), file_name=f"MODELO DRE - {escolha}.xlsx",
                       mime=MIME_XLSX, key=f"modelo_xls_{emp['id']}", width="stretch")

    criar = st.toggle("Criar grupos novos no DRE para categorias que não existem no modelo",
                      value=cfg.get("criar_grupos", True), key=f"criar_grupos_{emp['id']}",
                      help="Ex.: uma balada com a categoria “Pirotecnia” ganha o grupo 060800-PIROTECNIA, no mesmo "
                           "padrão de códigos do modelo. Desligado, a categoria vai para OUTRAS DESPESAS (A CLASSIFICAR).")
    if criar != cfg.get("criar_grupos", True):
        cfg["criar_grupos"] = criar
        db.salvar_config(emp["id"], cfg)
    if cfg.get("grupos_criados"):
        st.caption("Grupos criados automaticamente para esta empresa (podem ser renomeados ou removidos na aba "
                   "**Plano da DRE**): " + ", ".join(f"`{g}`" for g in cfg["grupos_criados"]))

    with st.expander("📥 Importar um DRE de exemplo em Excel e criar um modelo novo"):
        st.caption("Envie um report de fechamento (como o da Hagili ou o que o sistema gera) ou o modelo baixado "
                   "acima depois de editado. O sistema lê as linhas no formato “040100-NOME DO GRUPO” (grupos) e "
                   "“040101-Nome da conta” (contas) e os títulos das seções, como “(-) CUSTOS OPERACIONAIS”.")
        arq = st.file_uploader("Excel do DRE de exemplo", type=["xlsx", "xlsm"], key=f"modelo_up_{emp['id']}")
        if arq:
            try:
                grupos, contas = modelos.ler_excel_modelo(arq)
            except Exception as e:  # arquivo fora do formato ou corrompido
                st.error(f"Não consegui ler esse Excel: {e}")
                return
            st.success(f"Encontrei {len(grupos)} grupos e {len(contas)} contas.")
            prev = pd.DataFrame(grupos, columns=["Código", "Grupo", "Seção", "Relatório"])
            prev["Contas"] = [", ".join(n for n, g in contas if g == cod) for cod in prev["Código"]]
            st.dataframe(prev, hide_index=True, width="stretch", height=min(40 + 35 * len(prev), 500))
            nome = st.text_input("Nome do modelo", value=os.path.splitext(arq.name)[0][:60],
                                 key=f"modelo_nome_{emp['id']}")
            if st.button("💾 Salvar modelo e aplicar a esta empresa", type="primary", key=f"modelo_salvar_{emp['id']}"):
                if not nome.strip():
                    st.error("Dê um nome ao modelo.")
                    return
                db.salvar_modelo(nome.strip(), grupos, contas)
                aplicar_modelo(emp, f"{nome.strip()} (importado)", {"grupos": grupos, "contas": contas})
                st.success("Modelo salvo. Ele também fica disponível para as outras empresas.")
                st.rerun()


# =====================================================================
# PLANO DA DRE
# =====================================================================
def tela_plano():
    empresas = db.listar()
    if not empresas:
        st.info("Cadastre uma empresa primeiro.")
        return
    emp = st.selectbox("Empresa", empresas, format_func=lambda e: e["nome"], key="plano_empresa")
    cfg = emp["config"]
    st.caption("Os grupos são o nível 2 da DRE (ex.: 040100-MÃO DE OBRA OPERACIONAL). As contas de nível 3 são as "
               "categorias do OMIE, classificadas na tela de fechamento. A coluna “Relatório” define em qual slide os "
               "pagamentos do grupo aparecem.")
    df = pd.DataFrame(dre.grupos_da_empresa(cfg), columns=["Código", "Nome", "Seção", "Relatório"])
    df = st.data_editor(df, num_rows="dynamic", width="stretch", key=f"plano_{emp['id']}", column_config={
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
        cfg.update({"grupos": None, "contas_modelo": [], "modelo_dre": None, "grupos_criados": []})
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
def tela_report():
    topo("Reports BPO")
    avisos_pendentes()
    c1, c2 = st.columns([5, 1])
    c1.subheader("Reports BPO")
    c2.button("Cadastrar Empresas", key="linkbtn_empresas", on_click=ir_para, args=("bpo",))
    empresas = db.listar()
    if not empresas:
        st.info("Nenhuma empresa cadastrada. Vá em **Cadastrar Empresas** e cadastre a primeira "
                "(nome, CNPJ, App Key e App Secret do OMIE).")
        return
    hoje = date.today()
    st.session_state.setdefault("rep_mes", hoje.month - 1 or 12)
    st.session_state.setdefault("rep_ano", hoje.year if hoje.month > 1 else hoje.year - 1)
    ids = [e["id"] for e in empresas]
    if "rep_emp_proxima" in st.session_state:
        st.session_state["rep_emp"] = st.session_state.pop("rep_emp_proxima")
    if st.session_state.get("rep_emp") not in ids:
        st.session_state["rep_emp"] = ids[0]
    por_id = {e["id"]: e for e in empresas}

    with st.container(key="cartao_filtros"):
        tipo = st.segmented_control("Relatório", ["DRE", "Fluxo de Caixa", REL_RECEB], default="DRE",
                                    key="rep_tipo") or "DRE"
        if tipo == REL_RECEB:
            filtros = filtros_recebimento(por_id, ids)
        else:
            c1, cx, c2, c3, c4, c5 = st.columns([4, .45, 1.1, 1.1, 2, 1.8], vertical_alignment="bottom")
            emp = por_id[c1.selectbox("Empresa", ids, format_func=lambda i: por_id[i]["nome"], key="rep_emp")]
            if cx.button("✕", key="xemp_report", help=f"Excluir a empresa {emp['nome']}", width="stretch"):
                dialogo_excluir(emp)
            mes = c2.selectbox("Mês", list(range(1, 13)), format_func=lambda m: f"{m:02d}", key="rep_mes")
            ano = c3.number_input("Ano", 2015, 2100, key="rep_ano")
            fonte = c4.selectbox("Fonte", ["API OMIE", "Planilha exportada do OMIE"])
            puxar = c5.button("Montar DRE" if tipo == "DRE" else "Montar Fluxo", type="primary", width="stretch")
            arquivo = None
            sel = []
            if fonte.startswith("Planilha"):
                arquivo = st.file_uploader("Finanças › Movimentação da Conta Corrente › Exportar para Excel",
                                           type=["xlsx"])
            else:
                sel = seletor_contas(emp)
    if tipo == REL_RECEB:
        tela_recebimento(*filtros)
        return
    puxar = puxar or st.session_state.pop("rep_auto", False)

    chave = f"dados_{emp['id']}_{ano}_{mes}_{'-'.join(sorted(sel))}"
    if puxar:
        try:
            if fonte == "API OMIE":
                if not emp["app_key"] or not emp["app_secret"]:
                    st.error("Cadastre a App Key e o App Secret desta empresa.")
                    return
                if contas_conhecidas(emp) and not sel:
                    st.error("Selecione ao menos uma conta corrente para importar.")
                    return
                if contas_conhecidas(emp):
                    todas = len(sel) == len(contas_conhecidas(emp))
                    emp["config"]["contas_ativas"] = [] if todas else sel
                    db.salvar_config(emp["id"], emp["config"])
                log = []
                barra = st.progress(0.0, "Conectando ao OMIE...")
                cli = OmieCliente(emp["app_key"], emp["app_secret"], log=log)
                r = baixar_mes(cli, int(ano), int(mes), sel or None,
                               progresso=lambda p, t: barra.progress(p, t))
                pasta = os.path.join(db.PASTA_DADOS, "logs")
                os.makedirs(pasta, exist_ok=True)
                salvar_log(log, os.path.join(pasta, f"omie_{emp['id']}_{ano}_{mes:02d}.json"))
                guardar_contas(emp, r["contas"])
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
        ultimas_geracoes()
        return

    cfg = copy.deepcopy(emp["config"])
    mapa = emp["mapa"]
    antes = list(cfg.get("grupos_criados") or [])
    novas = dre.completar_mapeamento(mapa, dados["lancs"], cfg)
    if novas:
        db.salvar_mapa(emp["id"], mapa)
    criados = [g for g in cfg.get("grupos_criados") or [] if g not in antes]
    if criados:
        emp["config"]["grupos"] = cfg["grupos"]
        emp["config"]["grupos_criados"] = cfg["grupos_criados"]
        db.salvar_config(emp["id"], emp["config"])
        st.info(f"{len(criados)} categoria(s) deste mês não existiam no modelo de DRE da empresa e viraram grupos "
                "novos, no mesmo padrão: " + ", ".join(f"**{g}**" for g in criados)
                + ". Confira na aba **1. Classificação**.", icon="🆕")
    cfg["rateios_mes"] = cfg.get("rateios", {}).get(f"{ano}-{mes:02d}", [])
    saldos = dados["saldos"]

    base = dre.montar_base(dados["lancs"], cfg, mapa)
    d = dre.montar_dre(base, cfg)
    vistas = sorted({(l.get("projeto") if cfg["modo_colunas"] == "projeto" else l.get("conta_corrente")) or ""
                     for l in dados["lancs"]} - {""})
    st.session_state[f"colunas_vistas_{emp['id']}"] = vistas
    if tipo == "Fluxo de Caixa":
        tela_fluxo(emp, cfg, mapa, dados["lancs"], base, saldos, int(ano), int(mes))
        return
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
        st.dataframe(dre_df(d), width="stretch", height=700, hide_index=True)
    with abas[4]:
        dfb = pd.DataFrame(base)
        st.dataframe(dfb, width="stretch", height=600, hide_index=True)
    with abas[5]:
        aba_arquivos(emp, cfg, base, d, saldos, int(ano), int(mes))
    st.divider()
    ultimas_geracoes()


def tela_fluxo(emp, cfg, mapa, lancs, base, saldos, ano, mes):
    competencia = f"{ano}-{mes:02d}"
    previstos = cfg.get("previstos", {}).get(competencia, {})
    pendentes = fluxo.base_pendente(lancs, cfg, mapa)
    linhas = fluxo.montar_fluxo(base, pendentes, saldos, cfg, previstos)
    res = fluxo.resumo(linhas)
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Saldo inicial", brl(res["saldo_inicial"]))
    m2.metric("Entradas realizadas", brl(res["entradas"]), f"previsto {brl(res['entradas_previstas'])}",
              delta_color="off")
    m3.metric("Saídas realizadas", brl(res["saidas"]), f"previsto {brl(res['saidas_previstas'])}", delta_color="off")
    m4.metric("Saldo final realizado", brl(res["saldo_final"]))
    m5.metric("Saldo final previsto", brl(res["saldo_final_previsto"]))

    abas = st.tabs(["1. Previsto x Realizado", "2. Fluxo diário", "3. Excel"])
    with abas[0]:
        st.caption("**Previsto**: digite o valor previsto de cada conta (saídas com sinal negativo). Enquanto não for "
                   "digitado, o sistema sugere o realizado mais o que ainda está em aberto no OMIE no mês (previstos, "
                   "atrasados, a vencer). **Realizado** = movimento do extrato no mês. A classificação das categorias "
                   "é a mesma da DRE.")
        df = pd.DataFrame([{"chave": l["chave"], "tipo": l["tipo"],
                            "Descrição": {"grupo": "   ", "conta": "      "}.get(l["tipo"], "")
                            + (l.get("codigo") or l["descricao"]),
                            "Previsto": round(l["previsto"], 2), "Realizado": round(l["realizado"], 2),
                            "Diferença": round(l["realizado"] - l["previsto"], 2),
                            "% Realizado": (l["realizado"] / l["previsto"] * 100) if l["previsto"] else None}
                           for l in linhas])
        ed = st.data_editor(df, hide_index=True, width="stretch", height=min(40 + 35 * len(df), 720),
                            key=f"fluxo_{emp['id']}_{competencia}",
                            disabled=["Descrição", "Realizado", "Diferença", "% Realizado"],
                            column_config={"chave": None, "tipo": None,
                                           "Previsto": st.column_config.NumberColumn(format="%.2f"),
                                           "Realizado": st.column_config.NumberColumn(format="%.2f"),
                                           "Diferença": st.column_config.NumberColumn(format="%.2f"),
                                           "% Realizado": st.column_config.NumberColumn(format="%.0f%%")})
        c1, c2 = st.columns([1, 1])
        if c1.button("💾 Salvar previsto", type="primary"):
            novos = {r["chave"]: float(r["Previsto"] or 0) for r in ed.to_dict("records") if r["tipo"] == "conta"}
            emp["config"].setdefault("previstos", {})[competencia] = novos
            db.salvar_config(emp["id"], emp["config"])
            st.rerun()
        if previstos and c2.button("↩️ Voltar para a sugestão do OMIE"):
            emp["config"].get("previstos", {}).pop(competencia, None)
            db.salvar_config(emp["id"], emp["config"])
            st.rerun()
        st.caption("Só as linhas de conta (as mais recuadas) são gravadas; grupos, seções e saldos são recalculados.")
    dias = fluxo.diario(base, saldos, pendentes)
    with abas[1]:
        dfd = pd.DataFrame(dias)
        if dfd.empty:
            st.info("Sem movimento no mês.")
        else:
            dfd["data"] = pd.to_datetime(dfd["data"]).dt.strftime("%d/%m/%Y")
            st.dataframe(dfd.rename(columns={"data": "Data", "entradas": "Entradas", "saidas": "Saídas",
                                             "transferencias": "Transferências", "saldo": "Saldo realizado",
                                             "entradas_prev": "Entradas previstas", "saidas_prev": "Saídas previstas",
                                             "saldo_projetado": "Saldo projetado"}),
                         hide_index=True, width="stretch", height=min(40 + 35 * len(dfd), 720))
            st.caption("Previstas = lançamentos ainda em aberto no OMIE, na data de cada um. "
                       "Saldo projetado = saldo realizado + o que está em aberto.")
    with abas[2]:
        st.caption("Excel no mesmo padrão visual do modelo de DRE (Aptos Narrow, cabeçalhos azul-marinho, grupos em "
                   "cinza), com Previsto, Realizado e Diferença puxados por fórmulas da aba BASE FLUXO.")
        com_prev = st.checkbox("Incluir as colunas Previsto e Diferença ao lado do Realizado", value=True)
        k = f"fluxo_arq_{emp['id']}_{competencia}"
        if st.button("📦 Gerar Excel do Fluxo de Caixa", type="primary"):
            pasta = os.path.join(db.PASTA_DADOS, "saidas", emp["nome"].replace("/", "-"), competencia)
            os.makedirs(pasta, exist_ok=True)
            arq = os.path.join(pasta, f"FLUXO DE CAIXA {emp['nome']} {mes:02d}_{ano}.xlsx".replace("/", "-"))
            fluxo.gerar_excel_fluxo(arq, emp, ano, mes, linhas, dias, base, pendentes, saldos, com_prev)
            st.session_state[k] = arq
            st.success(f"Arquivo salvo em `{pasta}`")
        if st.session_state.get(k):
            with open(st.session_state[k], "rb") as f:
                st.download_button("⬇️ Baixar Excel do Fluxo de Caixa", f.read(),
                                   file_name=os.path.basename(st.session_state[k]), mime=MIME_XLSX)
    st.divider()
    ultimas_geracoes()


def filtros_recebimento(por_id, ids):
    """Linha de filtros do Relatório Recebimento: empresa e fonte (a base vem inteira, sem período)."""
    c1, cx, c4, c5 = st.columns([4, .45, 2, 1.8], vertical_alignment="bottom")
    emp = por_id[c1.selectbox("Empresa", ids, format_func=lambda i: por_id[i]["nome"], key="rep_emp")]
    if cx.button("✕", key="xemp_report", help=f"Excluir a empresa {emp['nome']}", width="stretch"):
        dialogo_excluir(emp)
    fonte = c4.selectbox("Fonte", ["API OMIE", "Planilha exportada do OMIE"], key="rec_fonte")
    puxar = c5.button("Montar Relatório", type="primary", width="stretch")
    arquivo = None
    if fonte.startswith("Planilha"):
        arquivo = st.file_uploader("Finanças › Contas a Receber › Exportar para Excel", type=["xlsx"],
                                   key="rec_arquivo")
    st.caption(f"Puxa todo o contas a receber da empresa, desde o início no OMIE. No relatório entra tudo o que "
               f"está **em aberto** (a vencer, vence hoje, atrasado, recebido parcialmente) e os **recebidos** "
               f"só dos dois últimos meses (desde {recebimento.inicio_recebidos():%d/%m/%Y}). "
               f"Cancelados ficam de fora.")
    return emp, fonte, arquivo, puxar


def tela_recebimento(emp, fonte, arquivo, puxar):
    chave = f"receb_{emp['id']}_{fonte}"
    if puxar:
        try:
            if fonte == "API OMIE":
                if not emp["app_key"] or not emp["app_secret"]:
                    st.error("Cadastre a App Key e o App Secret desta empresa.")
                    return
                log = []
                barra = st.progress(0.0, "Conectando ao OMIE...")
                cli = OmieCliente(emp["app_key"], emp["app_secret"], log=log)
                titulos = recebimento.baixar_titulos(cli, progresso=lambda p, t: barra.progress(p, t))
                pasta = os.path.join(db.PASTA_DADOS, "logs")
                os.makedirs(pasta, exist_ok=True)
                salvar_log(log, os.path.join(pasta, f"receber_{emp['id']}_{date.today():%Y_%m_%d}.json"))
            else:
                if not arquivo:
                    st.error("Envie a planilha exportada do OMIE.")
                    return
                titulos = recebimento.ler_exportacao(arquivo)
            st.session_state[chave] = titulos
        except (OmieErro, ValueError) as e:
            st.error(str(e))
            return

    titulos = st.session_state.get(chave)
    if titulos is None:
        st.info("Escolha a empresa e clique em **Montar Relatório** para puxar as contas a receber do OMIE.")
        return
    situacoes = st.multiselect("Situações", recebimento.SITUACOES, default=recebimento.SITUACOES,
                               key=f"rec_sit_{emp['id']}")
    sel = recebimento.filtrar(titulos, situacoes)
    res = recebimento.resumo(sel)
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Títulos", res["titulos"])
    m2.metric("Recebido (2 últimos meses)", brl(res["recebido"]))
    m3.metric("A receber", brl(res["a_receber"]))
    m4.metric("Atrasado", brl(res["atrasado"]))
    m5.metric("A vencer", brl(res["a_vencer"]))

    df = pd.DataFrame(sel, columns=[k for k, _, _ in recebimento.COLUNAS])
    for k in recebimento.DATAS:
        df[k] = pd.to_datetime(df[k])
    cfg_cols = {k: (st.column_config.DateColumn(t, format="DD/MM/YYYY") if k in recebimento.DATAS
                    else st.column_config.NumberColumn(t, format="%.2f") if k in recebimento.VALORES
                    else st.column_config.TextColumn(t))
                for k, t, _ in recebimento.COLUNAS}
    st.dataframe(df, hide_index=True, width="stretch", height=min(40 + 35 * len(df), 720), column_config=cfg_cols)

    k = f"rec_arq_{emp['id']}"
    if st.button("📦 Gerar Excel do Relatório de Recebimento", type="primary", disabled=not sel):
        hoje = date.today()
        pasta = os.path.join(db.PASTA_DADOS, "saidas", emp["nome"].replace("/", "-"), f"{hoje:%Y-%m}")
        os.makedirs(pasta, exist_ok=True)
        arq = os.path.join(pasta, f"RELATORIO DE RECEBIMENTO {emp['nome']} {hoje:%d_%m_%Y}.xlsx".replace("/", "-"))
        recebimento.gerar_excel(arq, emp, sel)
        st.session_state[k] = arq
        st.success(f"Arquivo salvo em `{pasta}`")
    if st.session_state.get(k) and os.path.exists(st.session_state[k]):
        with open(st.session_state[k], "rb") as f:
            st.download_button("⬇️ Baixar Excel do Relatório de Recebimento", f.read(),
                               file_name=os.path.basename(st.session_state[k]), mime=MIME_XLSX)


def seletor_contas(emp):
    """Escolha das contas correntes a importar; a escolha fica salva para a empresa."""
    contas = contas_conhecidas(emp)
    if not contas:
        c1, c2 = st.columns([4, 1.8], vertical_alignment="bottom")
        c1.caption("Todas as contas correntes da empresa serão importadas. Para escolher quais, carregue a lista do OMIE.")
        if c2.button("Carregar contas do OMIE", width="stretch"):
            try:
                with st.spinner("Consultando contas correntes..."):
                    guardar_contas(emp, OmieCliente(emp["app_key"], emp["app_secret"]).contas_correntes())
                st.rerun()
            except OmieErro as e:
                st.error(str(e))
        return []
    salvas = set(emp["config"].get("contas_ativas", [])) & {c["nCodCC"] for c in contas}
    st.markdown("**Contas correntes a importar**")
    cols = st.columns(6)
    sel = []
    for i, c in enumerate(contas):
        if cols[i % len(cols)].checkbox(c["descricao"], value=not salvas or c["nCodCC"] in salvas,
                                        key=f"rep_conta_{emp['id']}_{c['nCodCC']}"):
            sel.append(c["nCodCC"])
    return sel


def regerar(g):
    st.session_state.update({"rep_emp": g["empresa_id"], "rep_mes": g["mes"], "rep_ano": g["ano"], "rep_auto": True})


def ultimas_geracoes():
    st.markdown("**Últimas gerações**")
    gs = db.listar_geracoes()
    with st.container(key="cartao_geracoes"):
        if not gs:
            st.caption("Nenhum report gerado ainda. Escolha a empresa e o mês, clique em **Montar DRE** e gere os "
                       "arquivos na aba **6. Análise e arquivos**.")
            return
        larg = [3.4, 1, 1.4, 1.5, 0.7, 1.1, 1.3]
        cab = st.columns(larg)
        for col, t in zip(cab, ["Empresa", "Competência", "Gerado em", "Gerado por", "", "", ""]):
            col.markdown(f'<div class="tabela-cab">{t or "&nbsp;"}</div>', unsafe_allow_html=True)
        for g in gs:
            c = st.columns(larg, vertical_alignment="center")
            c[0].write(g["empresa"])
            c[1].write(f"{g['mes']:02d}/{g['ano']}")
            dt = g["gerado_em"]
            c[2].markdown(f'<span class="cinza">{dt[8:10]}/{dt[5:7]}/{dt[:4]} {dt[11:16]}</span>', unsafe_allow_html=True)
            c[3].markdown(f'<span class="cinza">{g["gerado_por"] or ""}</span>', unsafe_allow_html=True)
            c[4].button("Regerar", key=f"linkbtn_reg_{g['id']}", on_click=regerar, args=(g,))
            for col, arq, rot, mime in ((c[5], g["xlsx"], "Baixar Excel", MIME_XLSX), (c[6], g["pptx"], "Baixar PowerPoint", MIME_PPTX)):
                if arq and os.path.exists(arq):
                    with open(arq, "rb") as f:
                        col.download_button(f"⬇ {rot}", f.read(), file_name=os.path.basename(arq), mime=mime,
                                            key=f"linkbtn_{rot}_{g['id']}")


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
    ed = st.data_editor(df, width="stretch", hide_index=True, height=min(40 + 36 * len(df), 700),
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
    ed = st.data_editor(df, num_rows="dynamic", width="stretch", key=f"rat_{emp['id']}_{chave}",
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
    ed = st.data_editor(df[["conta", "saldo_inicial", "saldo_final"]], num_rows="dynamic", width="stretch",
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
                                     "conclusao": concl})
        st.session_state[k + "arqs"] = (xlsx, pptx)
        db.registrar_geracao(emp["id"], ano, mes, getpass.getuser(), xlsx, pptx)
        st.success(f"Arquivos salvos em `{pasta}`")
    arqs = st.session_state.get(k + "arqs")
    if arqs:
        c1, c2 = st.columns(2)
        with open(arqs[0], "rb") as f:
            c1.download_button("⬇️ Baixar Excel (DRE)", f.read(), file_name=os.path.basename(arqs[0]),
                               mime=MIME_XLSX)
        with open(arqs[1], "rb") as f:
            c2.download_button("⬇️ Baixar PowerPoint", f.read(), file_name=os.path.basename(arqs[1]),
                               mime=MIME_PPTX)


if st.session_state["pagina"] == "bpo":
    tela_bpo()
else:
    tela_report()
