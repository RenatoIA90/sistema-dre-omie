"""Plano padrão da DRE financeira (regime de caixa) e classificação automática
das categorias do OMIE nos grupos da DRE.

Estrutura (igual aos relatórios-modelo da Savian):
  nível 1 = seção  ((+) ENTRADAS, (-) CUSTOS OPERACIONAIS ...)
  nível 2 = grupo  (040100-MÃO DE OBRA OPERACIONAL ...)
  nível 3 = conta  (040101-Salários ...)  -> uma por categoria do OMIE
"""
import re
import unicodedata

# Seções na ordem em que aparecem na DRE. sinal só documenta a natureza esperada.
SECOES = [
    ("ENTRADAS", "(+) ENTRADAS"),
    ("DEDUCOES", "(-) DEDUÇÕES S/ RECEITA BRUTA"),
    ("CUSTOS", "(-) CUSTOS OPERACIONAIS"),
    ("DESPESAS", "(-) DESPESAS OPERACIONAIS"),
    ("FINANCEIRO", "(+/-) RESULTADO FINANCEIRO"),
    ("NAO_OPERACIONAL", "(+/-) RESULTADO NÃO OPERACIONAL"),
]

# Linhas de resultado (nível 1) e onde entram: depois de qual seção, somando quais.
RESULTADOS = {
    "DEDUCOES": ("(=) RECEITA OPERACIONAL LÍQUIDA", ["ENTRADAS", "DEDUCOES"]),
    "CUSTOS": ("(=) RESULTADO BRUTO", ["ENTRADAS", "DEDUCOES", "CUSTOS"]),
    "DESPESAS": ("(=) FCO / RESULTADO OPERACIONAL", ["ENTRADAS", "DEDUCOES", "CUSTOS", "DESPESAS"]),
}
RESULTADO_LIQUIDO = "(=) RESULTADO LÍQUIDO"
FCL = "(=) FCL / FLUXO DE CAIXA LIVRE"

IGNORAR = "IGNORAR"  # transferências entre contas: ficam fora da DRE

# codigo, nome, secao, relatorio (bloco do PowerPoint em que os pagamentos aparecem)
GRUPOS_PADRAO = [
    ("010100", "RECEBIMENTOS", "ENTRADAS", "recebimentos"),
    ("010200", "OUTRAS ENTRADAS OPERACIONAIS", "ENTRADAS", "recebimentos"),
    ("020100", "IMPOSTOS", "DEDUCOES", ""),
    ("040100", "MÃO DE OBRA OPERACIONAL", "CUSTOS", "funcionarios"),
    ("040200", "ENCARGOS SOCIAIS", "CUSTOS", ""),
    ("040300", "BENEFÍCIOS", "CUSTOS", ""),
    ("040400", "MATERIAIS E UNIFORMES", "CUSTOS", "fornecedores"),
    ("040500", "MÁQUINAS, EQUIPAMENTOS E VEÍCULOS", "CUSTOS", "fornecedores"),
    ("040600", "FORNECEDORES / MERCADORIAS", "CUSTOS", "fornecedores"),
    ("040700", "PRESTADORES DE SERVIÇOS", "CUSTOS", "fornecedores"),
    ("060100", "CONTABILIDADE", "DESPESAS", ""),
    ("060200", "SISTEMAS E TECNOLOGIA", "DESPESAS", ""),
    ("060300", "CONSULTORIA / SERVIÇOS DE TERCEIROS", "DESPESAS", "fornecedores"),
    ("060400", "ADMINISTRATIVO", "DESPESAS", ""),
    ("060500", "OCUPAÇÃO (ALUGUEL, CONDOMÍNIO, UTILIDADES)", "DESPESAS", ""),
    ("060600", "JURÍDICO", "DESPESAS", ""),
    ("060700", "SÓCIOS", "DESPESAS", ""),
    ("060900", "OUTRAS DESPESAS (A CLASSIFICAR)", "DESPESAS", ""),
    ("080100", "RECEITAS FINANCEIRAS", "FINANCEIRO", ""),
    ("080200", "DESPESAS FINANCEIRAS", "FINANCEIRO", ""),
    ("090100", "ENTRADAS NÃO OPERACIONAIS", "NAO_OPERACIONAL", ""),
    ("090200", "SAÍDAS NÃO OPERACIONAIS", "NAO_OPERACIONAL", ""),
    (IGNORAR, "TRANSFERÊNCIA ENTRE CONTAS (fora da DRE)", "FORA", "transferencias"),
]


def norm(txt):
    """minúsculas, sem acento, espaços simples."""
    txt = unicodedata.normalize("NFKD", str(txt or "")).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", txt.lower()).strip()


# (lista de termos, grupo se for ENTRADA, grupo se for SAÍDA). Primeira regra que casa vence.
# Termos com \b são tratados como regex de palavra inteira.
REGRAS = [
    (["transferencia", "transf. entre", "resgate de aplicacao", "aplicacao financeira"], IGNORAR, IGNORAR),
    (["rendimento", "juros recebidos", "receita financeira"], "080100", "080200"),
    (["aporte", "emprestimo", "financiamento", "mutuo"], "090100", "090200"),
    (["cartao de credito", "fatura"], "010100", "090200"),
    (["pro-labore", "pro labore", "prolabore", "socio", "retirada", "distribuicao de lucro", "dividendo"], "090100", "060700"),
    (["fgts", r"\binss\b", "encargo", "gps", "esocial"], "010200", "040200"),
    ([r"\biss\b", r"\bpis\b", "cofins", "simples nacional", r"\bdas\b", "icms", "irpj", "csll", "imposto", "tributo", r"\bdifal\b", r"\bipi\b"], "010200", "020100"),
    (["salario", "freelanc", r"\bft\b", "rescis", "ferias", "13o", "decimo terceiro", "adiantamento salarial", "hora extra", "comiss", "folha", "diarista", "cache"], "010200", "040100"),
    (["vale", "beneficio", "seguro de vida", "assistencia medica", "plano de saude", "odonto", "exame", "cesta basica", "alimentacao funcionario"], "010200", "040300"),
    (["material de escritorio", "escritorio"], "010200", "060400"),
    (["limpeza", "uniforme", "material", "epi"], "010200", "040400"),
    (["combust", "pedagio", "veicul", "maquina", "equipamento", "manutencao", "estacionamento", "frete"], "010200", "040500"),
    (["prestador", "prestacao de servico", "compra de servico", "terceiriz", r"\bdj\b", "banda", "freela"], "010200", "040700"),
    (["contabil"], "010200", "060100"),
    (["sistema", "software", "tecnologia", "internet", "certificado digital", "ponto eletronico", "licenca", "hospedagem", r"\bti\b"], "010200", "060200"),
    (["consultoria", "assessoria", "conselho"], "010200", "060300"),
    (["aluguel", "condominio", "energia", "agua", "iptu", r"\bluz\b", "gas"], "010200", "060500"),
    (["advog", "juridic", "protesto", "cartorio", "custas", "honorario"], "010200", "060600"),
    (["tarifa", "juros", r"\biof\b", "multa", "taxa banc", "taxa de cartao", "antecipacao"], "080100", "080200"),
    (["fornecedor", "mercadoria", "compra", "bebida", "aliment", "insumo", "materia prima", "estoque"], "010200", "040600"),
    (["telefon", "celular", "marketing", "publicidade", "recrutamento", "reembolso", "correio", "viagem",
      "adiantamento", "despesa administrativa", "cancelamento", "brinde", "doacao"], "010200", "060400"),
    (["cliente", "servico prestado", "servicos prestados", "venda", "receita", "recebimento", "faturamento",
      "cartao", r"\bpix\b", "credito", "debito", "boleto"], "010100", "060900"),
]


def _casa(termo, texto):
    if termo.startswith(r"\b"):
        return re.search(termo, texto) is not None
    return termo in texto


def sugerir_grupo(categoria, valor_total, flag_transferencia=False):
    """Sugere o grupo da DRE para uma categoria do OMIE.

    valor_total: soma dos lançamentos da categoria (>0 entrada, <0 saída).
    """
    if flag_transferencia:
        return IGNORAR
    t = norm(categoria)
    entrada = valor_total >= 0
    if not t:
        return "010200" if entrada else "060900"
    for termos, g_ent, g_sai in REGRAS:
        if any(_casa(x, t) for x in termos):
            return g_ent if entrada else g_sai
    return "010200" if entrada else "060900"


def regra_encontrada(categoria):
    """True se alguma regra do plano padrão reconhece a categoria."""
    t = norm(categoria)
    return any(_casa(x, t) for termos, _e, _s in REGRAS for x in termos)


_RX_RATEIO = re.compile(r"\s*(.+?)\s*\((\d+(?:[.,]\d+)?)%\)\s*")


def dividir_categoria(categoria):
    """'Combustivel (8,55%); Limpeza (91,45%)' -> [('Combustivel', .0855), ('Limpeza', .9145)].
    Categoria simples -> [(categoria, 1.0)]."""
    cat = str(categoria or "").strip()
    if ";" not in cat and "%" not in cat:
        return [(cat, 1.0)]
    partes = []
    for pedaco in cat.split(";"):
        m = _RX_RATEIO.fullmatch(pedaco)
        if not m:
            return [(cat, 1.0)]
        partes.append((m.group(1).strip(), float(m.group(2).replace(",", ".")) / 100))
    soma = sum(p for _, p in partes)
    if soma <= 0:
        return [(cat, 1.0)]
    return [(n, p / soma) for n, p in partes]
