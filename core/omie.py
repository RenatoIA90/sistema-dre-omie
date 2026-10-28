"""Cliente da API do OMIE (https://developer.omie.com.br).

Usa o extrato de conta corrente (financas/extrato ListarExtrato), que é a mesma
base da tela "Finanças > Movimentação da Conta Corrente" exportada em Excel.
"""
import calendar
import html
import json
import re
import time
from datetime import date, datetime

import requests

BASE = "https://app.omie.com.br/api/v1/"


class OmieErro(Exception):
    pass


class OmieCliente:
    def __init__(self, app_key, app_secret, timeout=60, log=None):
        self.app_key = str(app_key).strip()
        self.app_secret = str(app_secret).strip()
        self.timeout = timeout
        self.log = log  # lista opcional que recebe (call, param, resposta) para diagnóstico

    def chamar(self, endpoint, call, param, tentativas=4):
        corpo = {"call": call, "app_key": self.app_key, "app_secret": self.app_secret, "param": [param]}
        espera = 2
        for tentativa in range(tentativas):
            try:
                r = requests.post(BASE + endpoint, json=corpo, timeout=self.timeout)
            except requests.RequestException as e:
                if tentativa == tentativas - 1:
                    raise OmieErro(f"Falha de conexão com o OMIE: {e}")
                time.sleep(espera)
                espera *= 2
                continue
            try:
                dados = r.json()
            except ValueError:
                dados = {"faultstring": r.text[:500]}
            if self.log is not None:
                self.log.append({"call": call, "param": param, "status": r.status_code, "resposta": dados})
            falha = dados.get("faultstring") if isinstance(dados, dict) else None
            if r.status_code == 200 and not falha:
                return dados
            recusa_omie = bool(falha)
            falha = str(falha or r.text[:300])
            # Sem registros não é erro para nós.
            if "não existem registros" in falha.lower() or "nao existem registros" in falha.lower():
                return {}
            # Limite de consumo / requisição redundante: espera e tenta de novo.
            # O OMIE devolve HTTP 500 para qualquer recusa (ex.: "Lançamento não cadastrado"); repetir esses
            # erros só gera bloqueio por "consumo redundante". Repete apenas limite de consumo e falha de servidor.
            if any(x in falha.lower() for x in ("redundante", "bloqueada", "limite", "aguarde", "too many")) \
                    or r.status_code in (425, 429, 502, 503) or (r.status_code == 500 and not recusa_omie):
                if tentativa < tentativas - 1:
                    # o OMIE costuma dizer quanto esperar ("Aguarde 46 segundos...")
                    pedido = re.search(r"aguarde (\d+) segundo", falha.lower())
                    time.sleep(min(int(pedido.group(1)) + 1, 120) if pedido else espera)
                    espera *= 2
                    continue
            raise OmieErro(f"OMIE recusou {call}: {falha}")
        raise OmieErro(f"OMIE não respondeu {call}")

    # ---- cadastros ----
    def _paginar(self, endpoint, call, chave_lista, extra=None, por_pagina=500):
        pagina, itens = 1, []
        while True:
            param = {"pagina": pagina, "registros_por_pagina": por_pagina}
            param.update(extra or {})
            d = self.chamar(endpoint, call, param)
            itens += d.get(chave_lista, []) or []
            total = d.get("total_de_paginas") or 1
            if pagina >= total:
                return itens
            pagina += 1

    def empresa(self):
        d = self._paginar("geral/empresas/", "ListarEmpresas", "empresas_cadastro", por_pagina=50)
        return d[0] if d else {}

    def contas_correntes(self):
        return self._paginar("geral/contacorrente/", "ListarContasCorrentes", "ListarContasCorrentes",
                             {"apenas_importado_api": "N"}, por_pagina=100)

    def categorias(self):
        return self._paginar("geral/categorias/", "ListarCategorias", "categoria_cadastro")

    def projetos(self):
        try:
            return self._paginar("geral/projetos/", "ListarProjetos", "cadastro", {"apenas_importado_api": "N"})
        except OmieErro:
            return []

    def extrato(self, n_cod_cc, inicio, fim):
        return self.chamar("financas/extrato/", "ListarExtrato", {
            "nCodCC": int(n_cod_cc), "cCodIntCC": "",
            "dPeriodoInicial": inicio.strftime("%d/%m/%Y"), "dPeriodoFinal": fim.strftime("%d/%m/%Y"),
            "cExibirApenasSaldo": "N",
        })

    def testar(self):
        """Valida as chaves; devolve lista de contas correntes."""
        return self.contas_correntes()


def periodo(ano, mes):
    return date(ano, mes, 1), date(ano, mes, calendar.monthrange(ano, mes)[1])


def _pega(d, *chaves, padrao=None):
    v = _pega_bruto(d, *chaves, padrao=padrao)
    # o OMIE devolve &, <, >, aspas como entidades HTML (ex.: "ALFA &amp; BETA")
    return html.unescape(v) if isinstance(v, str) else v


def _pega_bruto(d, *chaves, padrao=None):
    for c in chaves:
        if c in d and d[c] not in (None, ""):
            return d[c]
    # busca sem diferenciar maiúsculas
    baixo = {k.lower(): v for k, v in d.items()}
    for c in chaves:
        v = baixo.get(c.lower())
        if v not in (None, ""):
            return v
    return padrao


def _data(v):
    if isinstance(v, (date, datetime)):
        return v if isinstance(v, date) else v.date()
    v = str(v or "").strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%y"):
        try:
            return datetime.strptime(v[:10], fmt).date()
        except ValueError:
            pass
    return None


def _num(v):
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v or "0").strip().replace("R$", "").replace(" ", "")
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def baixar_mes(cli, ano, mes, contas_ativas=None, progresso=None):
    """Baixa o extrato do mês de todas as contas correntes.

    Retorna dict com: lancamentos (lista de dicts normalizados), saldos (lista),
    categorias (dict codigo->cadastro), contas (lista), projetos (dict codigo->nome).
    """
    ini, fim = periodo(ano, mes)
    contas = cli.contas_correntes()
    cats = {}
    for c in cli.categorias():
        cats[str(c.get("codigo"))] = c
    projetos = {str(p.get("codigo")): p.get("nome") for p in cli.projetos()}

    lancs, saldos = [], []
    contas_usadas = [c for c in contas if str(c.get("inativo", "N")).upper() != "S"]
    if contas_ativas:
        contas_usadas = [c for c in contas_usadas if str(c.get("nCodCC")) in {str(x) for x in contas_ativas}]
    for i, cc in enumerate(contas_usadas):
        nome_cc = html.unescape(str(cc.get("descricao") or cc.get("nCodCC")))
        if progresso:
            progresso(i / max(len(contas_usadas), 1), f"Baixando extrato: {nome_cc}")
        ext = cli.extrato(cc["nCodCC"], ini, fim)
        if not ext:
            continue
        saldos.append({
            "conta": nome_cc, "nCodCC": cc["nCodCC"],
            "saldo_inicial": _num(_pega(ext, "nSaldoAnterior", padrao=0)),
            "saldo_final": _num(_pega(ext, "nSaldoAtual", "nSaldoConciliado", padrao=0)),
            "saldo_disponivel": _num(_pega(ext, "nSaldoDisponivel", padrao=0)),
        })
        for m in ext.get("listaMovimentos", []) or []:
            l = normalizar_mov_api(m, nome_cc, cats, projetos)
            if l and l["data"] and ini <= l["data"] <= fim:
                lancs.append(l)
    if progresso:
        progresso(1.0, "Extratos baixados")
    return {"lancamentos": lancs, "saldos": saldos, "categorias": cats, "contas": contas, "projetos": projetos}


def normalizar_mov_api(m, nome_cc, cats, projetos):
    fav = _pega(m, "cDesCliente", "cRazCliente", padrao="")
    if str(fav).upper().startswith("SALDO"):
        return None
    cod_cat = str(_pega(m, "cCodCategoria", padrao="") or "")
    cat = _pega(m, "cDesCategoria", padrao="") or (cats.get(cod_cat) or {}).get("descricao", "")
    cad = cats.get(cod_cat) or {}
    proj = _pega(m, "cDesProjeto", "cProjeto", "cNomeProjeto", padrao="")
    if not proj:
        cod_proj = _pega(m, "nCodProjeto", "cCodProjeto")
        proj = projetos.get(str(cod_proj), "") if cod_proj else ""
    valor = _num(_pega(m, "nValorDocumento", padrao=0))
    natureza = str(_pega(m, "cNatureza", padrao="")).upper()[:1]
    if natureza in ("P", "D") and valor > 0:  # pagamento/débito vem positivo em algumas contas
        valor = -valor
    return {
        "data": _data(_pega(m, "dDataLancamento", "dDataPagamento")),
        "conta_corrente": nome_cc,
        "favorecido": fav,
        "razao_social": _pega(m, "cRazCliente", padrao=""),
        "documento_favorecido": _pega(m, "cDocCliente", padrao=""),
        "categoria": cat,
        "cod_categoria": cod_cat,
        "transferencia": str(cad.get("transferencia", "N")).upper() == "S" or "transfer" in str(_pega(m, "cOrigem", padrao="")).lower(),
        "valor": valor,
        "projeto": proj or "",
        "situacao": _pega(m, "cSituacao", padrao=""),
        "origem": _pega(m, "cOrigem", padrao=""),
        "tipo_documento": _pega(m, "cTipoDocumento", padrao=""),
        "documento": _pega(m, "cNumero", "cDocumentoFiscal", padrao=""),
        "observacoes": _pega(m, "cObservacoes", padrao=""),
        "id_omie": _pega(m, "nCodLancamento", padrao=""),
    }


def salvar_log(log, caminho):
    with open(caminho, "w", encoding="utf-8") as f:
        json.dump(log, f, ensure_ascii=False, indent=1, default=str)
