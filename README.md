# Savian Contabilidade — Sistema de Fechamento DRE (OMIE)

Você escolhe a **empresa** e o **mês**. O sistema puxa o extrato do OMIE pela API, classifica as categorias, monta a
**DRE financeira (regime de caixa)** no formato dos reports da Savian e gera:

- um **Excel** com a DRE (fórmulas SUMIFS sobre a aba BASE CLASSIFICADA, AV% e conferências), mais as abas de saldos,
  recebimentos, fornecedores, funcionários, transferências e indicadores;
- um **PowerPoint** do Report Mensal no visual Savian, com capa, nossos serviços, DRE, DRE detalhada, saldos,
  recebimentos, fornecedores, funcionários, transferências, análise e pontos positivos/de atenção.

## Como rodar

Precisa de Python 3.10 ou mais novo.

- **Windows:** dê dois cliques em `iniciar.bat` (na primeira vez ele instala tudo).
- **Mac/Linux:** `./iniciar.sh`

O sistema abre no navegador em `http://localhost:8501`.

## Passo a passo

1. **Empresas**: cadastre nome, CNPJ(s), **App Key** e **App Secret** do OMIE e clique em *Testar conexão*.
   Escolha como as colunas da DRE serão divididas: **por conta corrente** (ex.: MS Festas), **por projeto**
   (ex.: Hagili) ou só o total.
2. **Fechamento do mês**: escolha empresa, mês e ano e clique em *Puxar dados e montar DRE*.
   Sem API, use a opção *Planilha exportada do OMIE* (Finanças › Movimentação da Conta Corrente › Exportar).
3. Aba **1. Classificação**: confira em qual grupo da DRE cada categoria do OMIE caiu. Na primeira vez é uma sugestão
   automática (⚠️). Depois de salvar, a classificação fica gravada para a empresa e os próximos meses já saem prontos.
4. Aba **2. Rateios** (opcional): move valores entre colunas (ex.: combustível sem centro de custo → postos).
5. Aba **3. Saldos**: saldos inicial e final por conta. Vêm do OMIE e podem ser ajustados.
6. Aba **6. Análise e arquivos**: revise os textos automáticos e clique em *Gerar Excel e PowerPoint*.

Os arquivos ficam salvos em `dados/saidas/<empresa>/<AAAA-MM>/` e também podem ser baixados na própria tela.

## Conferências

- **TOTAL − BASE = 0**: tudo que está na base classificada entrou na DRE.
- **Resultado líquido + transferências − variação de caixa = 0**: a DRE fecha com o saldo bancário.

## Estrutura

| Arquivo | O que faz |
|---|---|
| `app.py` | Telas (Streamlit) |
| `core/omie.py` | Cliente da API OMIE (contas correntes, categorias, projetos, extrato) |
| `core/importador.py` | Leitura da planilha exportada do OMIE |
| `core/plano.py` | Plano padrão da DRE e classificação automática das categorias |
| `core/dre.py` | Base classificada, DRE, quadros e conferências |
| `core/excel.py` | Geração do Excel |
| `core/apresentacao.py` / `core/analise.py` | PowerPoint e textos da análise |
| `testes/gerar_exemplos.py` | Gera exemplos a partir dos arquivos-modelo |

## Segurança

As chaves do OMIE ficam no banco local `dados/savian.db`, que está fora do Git. Não publique essa pasta.
