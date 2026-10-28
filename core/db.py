"""Cadastro de empresas (SQLite local em dados/savian.db)."""
import json
import os
import sqlite3

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
