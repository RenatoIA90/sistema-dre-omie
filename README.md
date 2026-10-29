# Savian Contabilidade — Sistema de Fechamento DRE (OMIE)

Você escolhe a **empresa** e o **mês**. O sistema puxa o extrato do OMIE pela API, classifica as categorias, monta a
**DRE financeira (regime de caixa)** no formato dos reports da Savian e gera:

- um **Excel** com a DRE (fórmulas SUMIFS sobre a aba BASE CLASSIFICADA, AV% e conferências), mais as abas de saldos,
  recebimentos, fornecedores, funcionários, transferências e indicadores;
- um **PowerPoint** do Report Mensal no visual Savian, com capa, nossos serviços, DRE, DRE detalhada, saldos,
  recebimentos, fornecedores, funcionários, transferências, análise e pontos positivos/de atenção.

## Como instalar e abrir (Windows)

1. No GitHub, clique em **Code › Download ZIP** e extraia a pasta onde preferir (ex.: `C:\Savian DRE`).
2. Dê dois cliques em **`iniciar.bat`**. Na primeira vez ele instala o Python (se faltar) e as bibliotecas,
   o que leva alguns minutos, e cria o atalho **Savian DRE** na área de trabalho.
3. O sistema abre no navegador (`http://localhost:8501`). Deixe a janela preta aberta enquanto usa.
   Das próximas vezes, é só usar o atalho da área de trabalho.

As empresas, chaves do OMIE, classificações e os arquivos gerados ficam em **Documentos › Savian DRE**, fora da
pasta do sistema. Para atualizar o sistema, baixe o ZIP de novo e substitua a pasta: nada do cadastro se perde.

Se o Python não instalar sozinho, instale pelo [python.org](https://www.python.org/downloads/) marcando
*Add python.exe to PATH* e abra o `iniciar.bat` de novo. Mac/Linux: `./iniciar.sh` (precisa de Python 3.10+).

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

## Fluxo de Caixa (previsto x realizado)

No Report, escolha **Fluxo de Caixa** em vez de DRE. O fluxo usa a mesma base puxada do OMIE e a mesma classificação
da DRE: saldo inicial, entradas e saídas por grupo/conta, transferências, variação e saldo final.

- **Previsto**: sugestão = realizado + lançamentos ainda em aberto no OMIE no mês; pode ser digitado por conta e fica salvo.
- **Realizado**: movimento do extrato no mês.
- **Fluxo diário**: saldo realizado e saldo projetado (com o que está em aberto) dia a dia.
- **Excel** no mesmo visual do modelo de DRE (Aptos Narrow, cabeçalhos azul-marinho, grupos em cinza, seções com
  borda), com Previsto, Realizado e Diferença puxados por SUMIFS da aba BASE FLUXO, mais SALDOS e FLUXO DIÁRIO.

Os arquivos ficam salvos em `Documentos/Savian DRE/saidas/<empresa>/<AAAA-MM>/` e também podem ser baixados na própria tela.

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
| `core/fluxo.py` | Fluxo de caixa previsto x realizado e Excel do fluxo |
| `core/excel.py` | Geração do Excel |
| `core/apresentacao.py` / `core/analise.py` | PowerPoint e textos da análise |
| `testes/gerar_exemplos.py` | Gera exemplos a partir dos arquivos-modelo |

## Segurança

As chaves do OMIE ficam no banco local `dados/savian.db`, que está fora do Git. Não publique essa pasta.
