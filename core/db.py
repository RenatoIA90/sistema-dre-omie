"""Cadastro de empresas (SQLite local em dados/savian.db)."""
import json
import os
import sqlite3
from datetime import datetime

PASTA_DADOS = os.environ.get("SAVIAN_DADOS", os.path.join(os.path.dirname(os.path.dirname(__file__)), "dados"))
os.makedirs(PASTA_DADOS, exist_ok=True)
ARQ = os.path.join(PASTA_DADOS, "savian.db")

CONFIG_PADRAO = {
    "modo_colunas": "conta",        # conta | projeto | total
    "apenas_realizado": True,       # ignora lançamentos atrasados / a vencer
    "apelidos_conta": {},           # nome no OMIE -> nome no relatório
    "apelidos_projeto": {},
    "grupos_colunas": {},           # coluna -> agrupamento (ex.: TOTAL MATRIZ)
    "ordem_colunas": [],
    "contas_ativas": [],            # nCodCC; vazio = todas
    "grupos": None,                 # None = plano padrão
    "recebimentos_por_cliente": False,
    "mostrar_grupos_vazios": False,
    "observacoes_mes": {},          # 'AAAA-MM' -> texto
    "rateios": {},                  # 'AAAA-MM' -> [ {descricao, conta, de, para, valor} ]
}


def _con():
    c = sqlite3.connect(ARQ)
    c.row_factory = sqlite3.Row
    c.execute("""CREATE TABLE IF NOT EXISTS empresas (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT NOT NULL, cnpj TEXT, app_key TEXT, app_secret TEXT,
        config TEXT DEFAULT '{}', mapa TEXT DEFAULT '{}', ativo INTEGER DEFAULT 1)""")
    c.execute("""CREATE TABLE IF NOT EXISTS geracoes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        empresa_id INTEGER NOT NULL, ano INTEGER NOT NULL, mes INTEGER NOT NULL,
        gerado_em TEXT NOT NULL, gerado_por TEXT, xlsx TEXT, pptx TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS modelos (
        id INTEGER PRIMARY KEY AUTOINCREMENT, nome TEXT NOT NULL UNIQUE, grupos TEXT NOT NULL, contas TEXT NOT NULL)""")
    return c


def _linha(r):
    d = dict(r)
    cfg = dict(CONFIG_PADRAO)
    cfg.update(json.loads(d.pop("config") or "{}"))
    d["config"] = cfg
    d["mapa"] = json.loads(d.get("mapa") or "{}")
    return d


def listar():
    with _con() as c:
        return [_linha(r) for r in c.execute("SELECT * FROM empresas WHERE ativo=1 ORDER BY nome")]


def obter(eid):
    with _con() as c:
        r = c.execute("SELECT * FROM empresas WHERE id=?", (eid,)).fetchone()
        return _linha(r) if r else None


def salvar(nome, cnpj, app_key, app_secret, eid=None):
    with _con() as c:
        if eid:
            c.execute("UPDATE empresas SET nome=?, cnpj=?, app_key=?, app_secret=? WHERE id=?",
                      (nome, cnpj, app_key, app_secret, eid))
            return eid
        cur = c.execute("INSERT INTO empresas (nome, cnpj, app_key, app_secret, config, mapa) VALUES (?,?,?,?,?,?)",
                        (nome, cnpj, app_key, app_secret, json.dumps(CONFIG_PADRAO), "{}"))
        return cur.lastrowid


def salvar_config(eid, config):
    with _con() as c:
        c.execute("UPDATE empresas SET config=? WHERE id=?", (json.dumps(config, ensure_ascii=False, default=str), eid))


def salvar_mapa(eid, mapa):
    with _con() as c:
        c.execute("UPDATE empresas SET mapa=? WHERE id=?", (json.dumps(mapa, ensure_ascii=False), eid))


def excluir(eid):
    with _con() as c:
        c.execute("UPDATE empresas SET ativo=0 WHERE id=?", (eid,))


def registrar_geracao(empresa_id, ano, mes, gerado_por, xlsx, pptx):
    """Guarda a última geração de cada empresa/competência (substitui a anterior)."""
    with _con() as c:
        c.execute("DELETE FROM geracoes WHERE empresa_id=? AND ano=? AND mes=?", (empresa_id, ano, mes))
        c.execute("INSERT INTO geracoes (empresa_id, ano, mes, gerado_em, gerado_por, xlsx, pptx) VALUES (?,?,?,?,?,?,?)",
                  (empresa_id, ano, mes, datetime.now().strftime("%Y-%m-%d %H:%M"), gerado_por, xlsx, pptx))


def listar_geracoes(limite=15):
    with _con() as c:
        return [dict(r) for r in c.execute(
            """SELECT g.*, e.nome AS empresa FROM geracoes g JOIN empresas e ON e.id = g.empresa_id
               WHERE e.ativo=1 ORDER BY g.gerado_em DESC, g.id DESC LIMIT ?""", (limite,))]


def salvar_modelo(nome, grupos, contas):
    """Modelo de DRE importado de um Excel; um nome já existente é substituído."""
    with _con() as c:
        c.execute("INSERT INTO modelos (nome, grupos, contas) VALUES (?,?,?) "
                  "ON CONFLICT(nome) DO UPDATE SET grupos=excluded.grupos, contas=excluded.contas",
                  (nome, json.dumps(grupos, ensure_ascii=False), json.dumps(contas, ensure_ascii=False)))


def listar_modelos():
    with _con() as c:
        return {r["nome"]: {"grupos": json.loads(r["grupos"]), "contas": json.loads(r["contas"])}
                for r in c.execute("SELECT * FROM modelos ORDER BY nome")}


def excluir_modelo(nome):
    with _con() as c:
        c.execute("DELETE FROM modelos WHERE nome=?", (nome,))
