"""Gera exemplos a partir dos arquivos-modelo (sem API): python -m testes.gerar_exemplos SAIDA"""
import sys
from datetime import date

import openpyxl

from core import analise, apresentacao, dre, excel
from core.importador import ler_exportacao_omie

saida = sys.argv[1]
PF = "/mnt/project-files/"


def rodar(nome, lancs, saldos, cfg, empresa):
    mapa = {}
    dre.completar_mapeamento(mapa, lancs)
    base = dre.montar_base(lancs, cfg, mapa)
    d = dre.montar_dre(base, cfg)
    print(nome, dre.conferencia(base, d, saldos))
    textos = analise.gerar_textos(empresa, 2026, 9, base, d, saldos, cfg)
    excel.gerar_excel(f"{saida}/{nome}.xlsx", empresa, 2026, 9, base, d, saldos, cfg)
    apresentacao.gerar_pptx(f"{saida}/{nome}.pptx", empresa, 2026, 9, base, d, saldos, cfg, textos)


# 1) exportação OMIE (1 conta corrente)
l, s = ler_exportacao_omie(PF + "financas_363808126921847.xlsx")
rodar("VDM_09_2026", l, s, {"modo_colunas": "conta"}, {"nome": "VDM", "cnpj": "00.000.000/0001-00"})

# 2) base Hagili em modo projeto
wb = openpyxl.load_workbook(PF + "FECHAMENTO HAGILI 09_2026 - POR PROJETO (9).xlsx", data_only=True)
ws = wb["BASE CLASSIFICADA"]
lancs = []
for r in list(ws.iter_rows(values_only=True))[1:]:
    if not r[0] or r[8] != "Lançamento":
        continue
    lancs.append({"data": r[0].date(), "conta_corrente": "BTG", "favorecido": r[1] or "", "categoria": r[3],
                  "valor": r[7], "projeto": "" if r[6] == dre.SEM_CC else r[6], "situacao": "Conciliado",
                  "observacoes": r[9] or "", "transferencia": False})
sd = [{"conta": "HAGILI - Conta Corrente", "saldo_inicial": 361514.63, "saldo_final": 440983.65}]
rodar("HAGILI_09_2026", lancs, sd, {"modo_colunas": "projeto"}, {"nome": "HAGILI", "cnpj": ""})
