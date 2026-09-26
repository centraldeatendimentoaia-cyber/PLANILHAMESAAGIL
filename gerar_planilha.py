# -*- coding: utf-8 -*-
"""Gera a planilha financeira do Mesa Ágil (Mesa_Agil_Financeiro.xlsx).

Estrutura inspirada na planilha Precifica Beleza Premium; cores do design system Saiaê
(esmeralda = marca/positivo, slate = texto, âmbar = células que você preenche, vermelho = negativo).
Rode:  python3 gerar_planilha.py
"""
import datetime as dt

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.cell.cell import MergedCell

from estilo_app import *  # noqa: F401,F403  (cores, fontes, put, entrada, botao, tile, card...)
from estilo_app import _s  # noqa: F401

SAIDA = "Mesa_Agil_Financeiro.xlsx"

wb = Workbook()
wb.remove(wb.active)

# ====================================================================== COMECE AQUI
guia = base(wb, "Comece Aqui", "MESA ÁGIL  ·  PLANEJAMENTO FINANCEIRO  ·  GUIA RÁPIDO", "C", MARCA,
            larguras={"B": 12, "C": 100, "D": 18})
botao(guia, "D2", "▦ DASHBOARD →", "#'Dashboard'!A1", AMBAR_CTA, PRETO)
passos = [
    ("PREMISSAS — preços dos planos (já vêm do escopo), cenários de venda, custos variáveis, marketing e custos fixos. Os campos AMARELO-CLAROS você edita.", "Premissas"),
    ("PREMISSAS › Cenário em uso — escolha Pessimista, Realista ou Otimista. Projeção, Fluxo de Caixa, Unit Economics e Dashboard passam a usar esse cenário.", "Premissas"),
    ("PROJEÇÃO 24 MESES — clientes (novos, cancelados, ativos por plano), MRR, ARR, ticket médio, custos, margem bruta e lucro líquido mês a mês.", "Projeção 24 meses"),
    ("FLUXO DE CAIXA — separa o dinheiro das mensalidades do dinheiro do plano anual (que entra de uma vez) e mostra a receita anual ainda a reconhecer.", "Fluxo de Caixa"),
    ("UNIT ECONOMICS — LTV, CAC, LTV/CAC e payback por plano e na média, mais o ponto de equilíbrio em número de clientes.", "Unit Economics"),
    ("CENÁRIOS — os três cenários lado a lado, sempre calculados ao mesmo tempo, com gráficos comparativos.", "Cenários"),
    ("DASHBOARD — resumo do cenário em uso: indicadores-chave e gráficos. Tudo automático.", "Dashboard"),
    ("MOTOR CENÁRIOS — cálculo automático que alimenta a aba Cenários. Não precisa mexer.", "Motor Cenários"),
]
for k, (txt, aba) in enumerate(passos):
    r = 4 + k
    put(guia, f"B{r}", f"{k + 1:02d}", 14, True, MOSTARDA_700, fam=F_DISPLAY)
    put(guia, f"C{r}", txt, 11, wrap=True, borda=borda_linha)
    botao(guia, f"D{r}", "ABRIR  →", f"#'{aba}'!A1", CARD, PRETO, borda=CLARO)
    guia.row_dimensions[r].height = 34
r = 4 + len(passos) + 1
put(guia, f"B{r}", "IMPORTANTE", 9, True, ATN_TXT)
put(guia, f"C{r}", "Só os PREÇOS dos planos vieram do escopo. Todos os outros campos amarelo-claros (clientes, churn, mix, custos, marketing) "
    "são EXEMPLOS estimados para a planilha já sair funcionando — troque pelos seus dados reais.", 9, cor=GRAFITE, wrap=True)
guia.row_dimensions[r].height = 30
put(guia, f"B{r + 2}", "CORES", 9, True, GRAFITE)
put(guia, f"C{r + 2}", "Campos AMARELO-CLAROS com borda mostarda = você preenche (clique neles para ver uma dica).  Células brancas = cálculo automático, não digite por cima.  "
    "Vermelho = resultado negativo.  Verde = resultado positivo.", 9, cor=GRAFITE, wrap=True)
guia.row_dimensions[r + 2].height = 30
put(guia, f"B{r + 4}", "COMO LER", 9, True, GRAFITE)
put(guia, f"C{r + 4}", "Receita (MRR, lucro) é por COMPETÊNCIA: o plano anual é distribuído em 12 meses.  Caixa é quando o dinheiro ENTRA: "
    "o plano anual entra inteiro no mês da venda/renovação.  Clientes aparecem com casas decimais porque são médias estatísticas.", 9, cor=GRAFITE, wrap=True)
guia.row_dimensions[r + 4].height = 42
put(guia, f"C{r + 6}", "Mesa Ágil (Powered by AIA) · comanda digital SaaS para barracas, food trucks, lanchonetes e restaurantes pequenos · mesaagil.pages.dev",
    9, cor=GRAFITE)

# ====================================================================== PREMISSAS
pr = base(wb, "Premissas", "PREMISSAS  ·  TODOS OS CAMPOS EDITÁVEIS", "G", MOSTARDA_600,
          "Campos AMARELO-CLAROS = você preenche. Preços dos planos vieram do escopo; o restante são exemplos para substituir pelos seus números.",
          {"B": 46, "C": 15, "D": 15, "E": 15, "F": 15, "G": 60})

secao(pr, "B5", "GERAL")
put(pr, "B6", "Nome do produto"); entrada(pr, "C6", "Mesa Ágil", al="left")
put(pr, "B7", "Mês de início da projeção"); entrada(pr, "C7", dt.date(2026, 10, 1), fmt=MES)
put(pr, "D7", "Primeiro mês de vendas (mês 1 da projeção)", 9, cor=GRAFITE)
put(pr, "B8", "Caixa inicial (R$)"); entrada(pr, "C8", 0, fmt=MOEDA)
put(pr, "D8", "Dinheiro já disponível no mês 1", 9, cor=GRAFITE)
put(pr, "B9", "Cenário em uso", b=True); entrada(pr, "C9", "Realista", al="center")
pr["C9"].font = font(11, True)
put(pr, "D9", "Pessimista, Realista ou Otimista — muda Projeção, Fluxo, Unit Economics e Dashboard", 9, cor=GRAFITE)
dv = DataValidation(type="list", formula1='"Pessimista,Realista,Otimista"', allow_blank=False,
                    showInputMessage=True, promptTitle="Cenário", prompt="Escolha o cenário que alimenta Projeção, Fluxo de Caixa, Unit Economics e Dashboard.")
pr.add_data_validation(dv); dv.add("C9")
dica(pr, "C7", "Início", "Data do primeiro mês de vendas. Ex.: 01/10/2026.")
dica(pr, "C8", "Caixa inicial", "Quanto dinheiro já existe para bancar o início (R$). Pode ser 0.")

secao(pr, "B11", "PREÇOS DOS PLANOS  (definidos no escopo)")
cabecalho(pr, 12, 2, ["PLANO", "MENSAL", "DESCONTO ANUAL", "ANUAL (POR MÊS)", "ANUAL (COBRADO DE UMA VEZ)", "O QUE INCLUI"])
planos = [
    ("Essencial", 57.90, 0.15, "1 barraca, lançar pedido, cozinha, chamada de senha, relatório do dia, histórico de 7 dias"),
    ("Pro", 87.90, 0.20, "Tudo do Essencial + várias barracas, histórico completo, exportar, custo/lucro/taxas, suporte prioritário no WhatsApp"),
]
for k, (nome, preco, desc, inclui) in enumerate(planos):
    r = 13 + k
    put(pr, f"B{r}", nome, b=True, borda=borda_linha)
    entrada(pr, f"C{r}", preco, fmt=MOEDA)
    entrada(pr, f"D{r}", desc, fmt=PCT)
    put(pr, f"E{r}", f"=C{r}*(1-D{r})", fmt=MOEDA, borda=borda_linha)
    put(pr, f"F{r}", f"=ROUND(E{r}*12,2)", fmt=MOEDA, borda=borda_linha)
    put(pr, f"G{r}", inclui, 9, cor=GRAFITE, wrap=True)
    pr.row_dimensions[r].height = 26
    dica(pr, f"C{r}", "Preço mensal", "Preço da assinatura mensal (do escopo).")
    dica(pr, f"D{r}", "Desconto anual", "Desconto do plano anual sobre o mensal. Ex.: 15% ou 0,15.")
put(pr, "B15", "No plano anual a RECEITA é reconhecida em 12 parcelas iguais (cobrado ÷ 12), mas o CAIXA entra inteiro no mês da venda.", 9, cor=GRAFITE)

secao(pr, "B17", "PREMISSAS DE VENDA POR CENÁRIO")
cabecalho(pr, 18, 2, ["PREMISSA", "PESSIMISTA", "REALISTA", "OTIMISTA", "EM USO", "COMO PREENCHER"])
cen = [
    ("Novos clientes pagantes no mês 1", (4, 8, 15), NUM1,
     "Quantas assinaturas novas no primeiro mês de vendas"),
    ("Crescimento mensal dos novos clientes", (0.03, 0.07, 0.12), PCT,
     "Quanto o nº de novos clientes cresce a cada mês. Ex.: 7% → 8; 8,6; 9,2..."),
    ("Churn mensal (plano mensal)", (0.08, 0.05, 0.03), PCT,
     "% dos clientes do plano MENSAL que cancelam por mês"),
    ("Renovação do plano anual", (0.50, 0.65, 0.80), PCT,
     "% dos clientes anuais que renovam ao fim dos 12 meses"),
    ("% dos novos clientes no plano Pro", (0.15, 0.25, 0.35), PCT,
     "Mix de planos: o restante vai para o Essencial"),
    ("% dos novos clientes no plano anual", (0.10, 0.20, 0.30), PCT,
     "Mix de cobrança: o restante paga mensal"),
    ("Conversão do teste grátis", (0.10, 0.20, 0.30), PCT,
     "% dos testes grátis que viram pagantes. Sem teste grátis? Coloque 100%"),
]
CEN_LINHA = {}
for k, (rot, vals, fmt, como) in enumerate(cen):
    r = 19 + k
    put(pr, f"B{r}", rot, borda=borda_linha)
    for j, v in enumerate(vals):
        entrada(pr, f"{L(3 + j)}{r}", v, fmt=fmt)
    put(pr, f"F{r}", f"=INDEX(C{r}:E{r},MATCH($C$9,$C$18:$E$18,0))", b=True, fmt=fmt, al="right", borda=borda_linha)
    put(pr, f"G{r}", como, 9, cor=GRAFITE)
    CEN_LINHA[k] = r
    dica(pr, f"C{r}:E{r}", rot[:32], como)
CHAVES_CEN = ["novos1", "cresc", "churn", "renov", "mixpro", "mixanual", "conv"]
LINHA_CEN = dict(zip(CHAVES_CEN, [19 + k for k in range(len(cen))]))

secao(pr, "B27", "CUSTOS VARIÁVEIS POR CLIENTE")
var = [
    ("Taxa do gateway de pagamento (%)", 0.035, PCT, "C28", "% cobrado pelo gateway (cartão/Pix/boleto) sobre cada cobrança. Média ponderada."),
    ("Tarifa fixa por cobrança (R$)", 0.49, MOEDA, "C29", "Valor fixo por cobrança gerada (se houver). Mensal = 1 cobrança/mês; anual = 1 por ano."),
    ("Imposto sobre faturamento (%)", 0.06, PCT, "C30", "Alíquota do Simples Nacional (SaaS costuma ser Anexo III, a partir de 6%). Confirme com o contador."),
    ("Infraestrutura por cliente ativo (R$/mês)", 0.50, MOEDA, "C31", "Custo extra de servidor/banco/tempo real para cada cliente ativo."),
]
for rot, v, fmt, ref, como in var:
    r = int(ref[1:])
    put(pr, f"B{r}", rot, borda=borda_linha)
    entrada(pr, ref, v, fmt=fmt)
    put(pr, f"D{r}", como, 9, cor=GRAFITE)
    dica(pr, ref, rot[:32], como)

secao(pr, "B33", "AQUISIÇÃO (MARKETING)")
aq = [
    ("Investimento em anúncios no mês 1 (R$)", 1500, MOEDA, "C34", "Meta/Google/Instagram pago no primeiro mês"),
    ("Crescimento mensal do investimento em anúncios", 0.03, PCT, "C35", "Quanto o orçamento de anúncios sobe por mês. 0% = fixo"),
    ("% dos novos clientes vindos de indicação/revenda", 0.20, PCT, "C36", "Parte dos novos clientes que gera comissão"),
    ("Comissão por cliente indicado (R$, paga uma vez)", 30, MOEDA, "C37", "Valor pago ao indicador/revendedor por cliente novo. 0 se não houver"),
]
for rot, v, fmt, ref, como in aq:
    r = int(ref[1:])
    put(pr, f"B{r}", rot, borda=borda_linha)
    entrada(pr, ref, v, fmt=fmt)
    put(pr, f"D{r}", como, 9, cor=GRAFITE)
    dica(pr, ref, rot[:32], como)
put(pr, "B38", "O CAC não é digitado: ele é calculado (anúncios + comissões ÷ novos clientes) na Projeção e no Unit Economics.", 9, cor=GRAFITE)

secao(pr, "B40", "CUSTOS FIXOS MENSAIS")
cabecalho(pr, 41, 2, ["DESCRIÇÃO", "VALOR MENSAL", "A PARTIR DO MÊS", "", "", "OBSERVAÇÃO"], altura=24)
fixos = [
    ("Hospedagem — Cloudflare Pages", 0, 1, "Plano gratuito atende no início"),
    ("Backend, banco de dados e tempo real", 150, 1, "Ex.: Supabase/Firebase pago (~US$ 25)"),
    ("Domínio", 10, 1, "Valor anual ÷ 12"),
    ("E-mail profissional", 30, 1, "Google Workspace, Zoho etc."),
    ("Ferramentas de IA", 110, 1, "Assinaturas de IA usadas no desenvolvimento/atendimento"),
    ("Design (Canva etc.)", 35, 1, ""),
    ("Site WordPress/Elementor + hospedagem", 60, 1, "Site institucional / página de vendas"),
    ("WhatsApp de atendimento (número/API)", 50, 1, "Suporte prioritário do plano Pro"),
    ("Pró-labore", 1500, 1, "Sua retirada mensal"),
    ("Suporte e atendimento (pessoa)", 1200, 7, "Exemplo: contratar a partir do mês 7"),
    ("Contador", 250, 1, ""),
    ("Impostos/taxas fixas (DAS MEI, anuidades)", 0, 1, "No Simples o imposto é % (linha 30); aqui só o que é fixo"),
]
FIX_INI, FIX_FIM = 42, 61
for k in range(FIX_FIM - FIX_INI + 1):
    r = FIX_INI + k
    d, v, m, obs = fixos[k] if k < len(fixos) else ("", None, None, "")
    entrada(pr, f"B{r}", d, al="left")
    entrada(pr, f"C{r}", v, fmt=MOEDA)
    entrada(pr, f"D{r}", m, fmt="0", al="center")
    put(pr, f"G{r}", obs, 9, cor=GRAFITE)
dica(pr, f"C{FIX_INI}:C{FIX_FIM}", "Valor mensal", "Quanto esse custo representa por mês (R$). Custos anuais: divida por 12.")
dica(pr, f"D{FIX_INI}:D{FIX_FIM}", "A partir do mês", "Mês da projeção (1 a 24) em que esse custo começa. Vazio = desde o mês 1.")
r = FIX_FIM + 1
put(pr, f"B{r}", "TOTAL (todos os custos já ativos)", b=True, borda=borda_total)
put(pr, f"C{r}", f"=SUM(C{FIX_INI}:C{FIX_FIM})", b=True, fmt=MOEDA, borda=borda_total)
pr.freeze_panes = "A4"

P = {
    "inicio": "Premissas!$C$7", "caixa0": "Premissas!$C$8",
    "pEM": "Premissas!$C$13", "pEC": "Premissas!$F$13", "pPM": "Premissas!$C$14", "pPC": "Premissas!$F$14",
    "gw": "Premissas!$C$28", "gwfix": "Premissas!$C$29", "imp": "Premissas!$C$30", "infra": "Premissas!$C$31",
    "ads1": "Premissas!$C$34", "adsg": "Premissas!$C$35", "indpct": "Premissas!$C$36", "indval": "Premissas!$C$37",
    "fixval": f"Premissas!$C${FIX_INI}:$C${FIX_FIM}", "fixmes": f"Premissas!$D${FIX_INI}:$D${FIX_FIM}",
}


def params(col_cen):
    """Premissas de cenário: col 'F' = em uso, 'C'/'D'/'E' = pessimista/realista/otimista."""
    d = dict(P)
    for k, r in LINHA_CEN.items():
        d[k] = f"Premissas!${col_cen}${r}"
    return d


# ====================================================================== MOTOR (mesma lógica em todas as abas)
FR, LR = 6, 29  # linhas dos meses 1..24

# (chave, cabeçalho, formato, grupo, fórmula)  — R(chave, desloc) referencia outra coluna; p = premissas
MOTOR = [
    ("t", "Nº", "0", "cli", lambda R, p, r: "1" if r == FR else f"{R('t', -1)}+1"),
    ("mes", "MÊS", MES, "cli", lambda R, p, r: f"EDATE({p['inicio']},{R('t')}-1)"),
    ("novos", "NOVOS CLIENTES", NUM1, "cli", lambda R, p, r: f"{p['novos1']}*(1+{p['cresc']})^({R('t')}-1)"),
    ("testes", "TESTES GRÁTIS NECESSÁRIOS", NUM, "cli", lambda R, p, r: f"IFERROR({R('novos')}/{p['conv']},0)"),
    ("cancel", "CANCELADOS", NUM1, "cli",
     lambda R, p, r: f"(N({R('aEM', -1)})+N({R('aPM', -1)}))*{p['churn']}+{R('naoren')}"),
    ("ativos", "CLIENTES ATIVOS", NUM1, "cli", lambda R, p, r: f"{R('aEM')}+{R('aEA')}+{R('aPM')}+{R('aPA')}"),
    ("aEM", "ESSENCIAL MENSAL", NUM1, "cli", lambda R, p, r: f"N({R('aEM', -1)})*(1-{p['churn']})+{R('nEM')}"),
    ("aEA", "ESSENCIAL ANUAL", NUM1, "cli",
     lambda R, p, r: f"SUM(INDEX({R.rng('entEA')},MAX(1,{R('t')}-11)):INDEX({R.rng('entEA')},{R('t')}))"),
    ("aPM", "PRO MENSAL", NUM1, "cli", lambda R, p, r: f"N({R('aPM', -1)})*(1-{p['churn']})+{R('nPM')}"),
    ("aPA", "PRO ANUAL", NUM1, "cli",
     lambda R, p, r: f"SUM(INDEX({R.rng('entPA')},MAX(1,{R('t')}-11)):INDEX({R.rng('entPA')},{R('t')}))"),
    ("mrr", "MRR (RECEITA RECORRENTE)", MOEDA, "dre",
     lambda R, p, r: f"{R('aEM')}*{p['pEM']}+{R('aEA')}*{p['pEC']}/12+{R('aPM')}*{p['pPM']}+{R('aPA')}*{p['pPC']}/12"),
    ("arr", "ARR (MRR × 12)", MOEDA, "dre", lambda R, p, r: f"{R('mrr')}*12"),
    ("arpu", "TICKET MÉDIO (ARPU)", MOEDA, "dre", lambda R, p, r: f"IFERROR({R('mrr')}/{R('ativos')},0)"),
    ("imp", "IMPOSTOS", MOEDA, "dre", lambda R, p, r: f"{R('mrr')}*{p['imp']}"),
    ("gwc", "GATEWAY", MOEDA, "dre", lambda R, p, r: f"{R('mrr')}*{p['gw']}+{R('cobr')}*{p['gwfix']}"),
    ("infra", "INFRA POR CLIENTE", MOEDA, "dre", lambda R, p, r: f"{R('ativos')}*{p['infra']}"),
    ("margem", "MARGEM BRUTA", MOEDA, "dre", lambda R, p, r: f"{R('mrr')}-{R('imp')}-{R('gwc')}-{R('infra')}"),
    ("margpct", "MARGEM BRUTA %", PCT, "dre", lambda R, p, r: f"IFERROR({R('margem')}/{R('mrr')},0)"),
    ("fixos", "CUSTOS FIXOS", MOEDA, "dre", lambda R, p, r: f"SUMPRODUCT({p['fixval']},--({p['fixmes']}<={R('t')}))"),
    ("ads", "ANÚNCIOS", MOEDA, "dre", lambda R, p, r: f"{p['ads1']}*(1+{p['adsg']})^({R('t')}-1)"),
    ("comis", "COMISSÕES DE INDICAÇÃO", MOEDA, "dre", lambda R, p, r: f"{R('novos')}*{p['indpct']}*{p['indval']}"),
    ("custot", "CUSTOS TOTAIS", MOEDA, "dre",
     lambda R, p, r: f"{R('imp')}+{R('gwc')}+{R('infra')}+{R('fixos')}+{R('ads')}+{R('comis')}"),
    ("lucro", "LUCRO LÍQUIDO", MOEDA, "dre", lambda R, p, r: f"{R('mrr')}-{R('custot')}"),
    ("lucroac", "LUCRO ACUMULADO", MOEDA, "dre", lambda R, p, r: f"N({R('lucroac', -1)})+{R('lucro')}"),
    ("cac", "CAC DO MÊS", MOEDA, "dre", lambda R, p, r: f"IFERROR(({R('ads')}+{R('comis')})/{R('novos')},0)"),
    ("equil", "CLIENTES P/ EQUILÍBRIO", NUM, "dre", lambda R, p, r: f"IFERROR({R('fixos')}/({R('margem')}/{R('ativos')}),0)"),
    # auxiliares
    ("nEM", "NOVOS ESS. MENSAL", NUM1, "aux", lambda R, p, r: f"{R('novos')}*(1-{p['mixpro']})*(1-{p['mixanual']})"),
    ("nEA", "NOVOS ESS. ANUAL", NUM1, "aux", lambda R, p, r: f"{R('novos')}*(1-{p['mixpro']})*{p['mixanual']}"),
    ("nPM", "NOVOS PRO MENSAL", NUM1, "aux", lambda R, p, r: f"{R('novos')}*{p['mixpro']}*(1-{p['mixanual']})"),
    ("nPA", "NOVOS PRO ANUAL", NUM1, "aux", lambda R, p, r: f"{R('novos')}*{p['mixpro']}*{p['mixanual']}"),
    ("entEA", "COBRANÇAS ANUAIS ESS. (NOVOS + RENOV.)", NUM1, "aux",
     lambda R, p, r: f"{R('nEA')}+IF({R('t')}>12,INDEX({R.rng('nEA')},{R('t')}-12)*{p['renov']},0)"),
    ("entPA", "COBRANÇAS ANUAIS PRO (NOVOS + RENOV.)", NUM1, "aux",
     lambda R, p, r: f"{R('nPA')}+IF({R('t')}>12,INDEX({R.rng('nPA')},{R('t')}-12)*{p['renov']},0)"),
    ("naoren", "ANUAIS QUE NÃO RENOVARAM", NUM1, "aux",
     lambda R, p, r: f"IF({R('t')}>12,(INDEX({R.rng('nEA')},{R('t')}-12)+INDEX({R.rng('nPA')},{R('t')}-12))*(1-{p['renov']}),0)"),
    ("cobr", "Nº DE COBRANÇAS NO MÊS", NUM1, "aux", lambda R, p, r: f"{R('aEM')}+{R('aPM')}+{R('entEA')}+{R('entPA')}"),
    ("lucropos", "MÊS COM LUCRO?", "0", "aux", lambda R, p, r: f'IF({R("lucro")}>0,{R("t")},"")'),
    # caixa
    ("cxm", "ENTRADA MENSALIDADES", MOEDA, "cx", lambda R, p, r: f"{R('aEM')}*{p['pEM']}+{R('aPM')}*{p['pPM']}"),
    ("cxa", "ENTRADA PLANO ANUAL (À VISTA)", MOEDA, "cx", lambda R, p, r: f"{R('entEA')}*{p['pEC']}+{R('entPA')}*{p['pPC']}"),
    ("cxin", "TOTAL DE ENTRADAS", MOEDA, "cx", lambda R, p, r: f"{R('cxm')}+{R('cxa')}"),
    ("cximp", "(−) IMPOSTOS", MOEDA, "cx", lambda R, p, r: f"{R('cxin')}*{p['imp']}"),
    ("cxgw", "(−) GATEWAY", MOEDA, "cx", lambda R, p, r: f"{R('cxin')}*{p['gw']}+{R('cobr')}*{p['gwfix']}"),
    ("cxinfra", "(−) INFRA", MOEDA, "cx", lambda R, p, r: f"{R('infra')}"),
    ("cxfix", "(−) CUSTOS FIXOS", MOEDA, "cx", lambda R, p, r: f"{R('fixos')}"),
    ("cxmkt", "(−) MARKETING", MOEDA, "cx", lambda R, p, r: f"{R('ads')}+{R('comis')}"),
    ("cxout", "TOTAL DE SAÍDAS", MOEDA, "cx",
     lambda R, p, r: f"{R('cximp')}+{R('cxgw')}+{R('cxinfra')}+{R('cxfix')}+{R('cxmkt')}"),
    ("cxsaldo", "SALDO DO MÊS", MOEDA, "cx", lambda R, p, r: f"{R('cxin')}-{R('cxout')}"),
    ("cxac", "SALDO ACUMULADO EM CAIXA", MOEDA, "cx",
     lambda R, p, r: f"IF({R('t')}=1,{p['caixa0']},N({R('cxac', -1)}))+{R('cxsaldo')}"),
    ("recan", "RECEITA ANUAL RECONHECIDA NO MÊS", MOEDA, "cx",
     lambda R, p, r: f"{R('aEA')}*{p['pEC']}/12+{R('aPA')}*{p['pPC']}/12"),
    ("difer", "RECEITA ANUAL A RECONHECER (JÁ RECEBIDA)", MOEDA, "cx",
     lambda R, p, r: f"N({R('difer', -1)})+{R('cxa')}-{R('recan')}"),
]
M = {k: (h, fmt, g, fn) for k, h, fmt, g, fn in MOTOR}
SOMA = {"novos", "testes", "cancel", "mrr", "custot", "imp", "gwc", "infra", "margem", "fixos", "ads", "comis", "lucro",
        "nEM", "nEA", "nPM", "nPA", "entEA", "entPA", "naoren", "cobr",
        "cxm", "cxa", "cxin", "cximp", "cxgw", "cxinfra", "cxfix", "cxmkt", "cxout", "cxsaldo", "recan"}


class Resolver:
    """Converte chave do motor em referência de célula (mesma aba ou outra)."""

    def __init__(self, mapa, aba_atual):
        self.mapa, self.aba, self.r = mapa, aba_atual, FR

    def _pref(self, k):
        aba, col = self.mapa[k]
        return ("" if aba == self.aba else f"'{aba}'!"), col

    def __call__(self, k, d=0):
        pref, col = self._pref(k)
        return f"{pref}{col}{self.r + d}"

    def rng(self, k):
        pref, col = self._pref(k)
        return f"{pref}${col}${FR}:${col}${LR}"


def escreve(ws, mapa, chaves, p, total=True, larg=14):
    R = Resolver(mapa, ws.title)
    for k in chaves:
        col = mapa[k][1]
        h, fmt, g, fn = M[k]
        ws.column_dimensions[col].width = 7 if k == "t" else (9 if k == "mes" else larg)
        for r in range(FR, LR + 1):
            R.r = r
            destaque = k in ("ativos", "mrr", "lucro", "cxac", "cxsaldo")
            put(ws, f"{col}{r}", "=" + fn(R, p, r), b=destaque, fmt=fmt,
                cor=GRAFITE if g == "aux" else PRETO, borda=borda_linha,
                al="center" if k in ("t", "mes") else None)
        if total:
            c = ws[f"{col}{LR + 1}"]
            c.border = borda_total
            if k in SOMA:
                put(ws, f"{col}{LR + 1}", f"=SUM({col}{FR}:{col}{LR})", b=True, fmt=fmt, borda=borda_total)


def mapa_bloco(aba, chaves, col_ini):
    return {k: (aba, L(col_ini + i)) for i, k in enumerate(chaves)}


# ---------------------------------------------------------------- PROJEÇÃO
PJ = "Projeção 24 meses"
ch_cli = [k for k, *_ in MOTOR if M[k][2] == "cli"]
ch_dre = [k for k, *_ in MOTOR if M[k][2] == "dre"]
ch_aux = [k for k, *_ in MOTOR if M[k][2] == "aux"]
ch_cx = [k for k, *_ in MOTOR if M[k][2] == "cx"]

vis = ch_cli + ch_dre
col_aux = 2 + len(vis) + 1
mapa = mapa_bloco(PJ, vis, 2)
mapa.update(mapa_bloco(PJ, ch_aux, col_aux))
FC = "Fluxo de Caixa"
mapa.update({"t_fc": (FC, "B")})
mapa.update(mapa_bloco(FC, ch_cx, 4))

ult = L(col_aux + len(ch_aux) - 1)
pj = base(wb, PJ, "", ult, GRAFITE)
pj["B2"].value = '="PROJEÇÃO 24 MESES  ·  CENÁRIO "&UPPER(Premissas!$C$9)'
put(pj, "B3", "Clientes, MRR, custos e lucro por COMPETÊNCIA (o plano anual é distribuído em 12 meses). Tudo automático — edite só a aba Premissas.",
    9, cor=GRAFITE)
# faixas de grupo
grupos = [("CLIENTES", ch_cli), ("RECEITA E RESULTADO", ch_dre), ("CÁLCULOS AUXILIARES", ch_aux)]
for nome, ks in grupos:
    c1, c2 = mapa[ks[0]][1], mapa[ks[-1]][1]
    pj.merge_cells(f"{c1}4:{c2}4")
    put(pj, f"{c1}4", nome, 9, True, "FFFFFF" if nome != "CÁLCULOS AUXILIARES" else PRETO,
        bg=MARCA_2 if nome != "CÁLCULOS AUXILIARES" else CLARO, al="center")
for k in vis + ch_aux:
    put(pj, f"{mapa[k][1]}5", M[k][0], 9, True, PRETO if M[k][2] != "aux" else GRAFITE,
        bg=SUB if M[k][2] != "aux" else CARD, al="center", wrap=True, borda=borda_linha)
pj.row_dimensions[5].height = 42
escreve(pj, mapa, vis + ch_aux, params("F"))
put(pj, f"B{LR + 1}", "TOTAL", b=True, borda=borda_total)
pj.merge_cells(f"B{LR + 1}:C{LR + 1}")
pj.column_dimensions[L(col_aux - 1)].width = 3
pj.freeze_panes = "D6"
for k in ("lucro", "lucroac", "margem"):
    c = mapa[k][1]
    cor_condicional(pj, f"{c}{FR}:{c}{LR + 1}", f"{c}{FR}<0")

# ---------------------------------------------------------------- FLUXO DE CAIXA
fc = base(wb, FC, "", L(3 + len(ch_cx)), GRAFITE)
fc["B2"].value = '="FLUXO DE CAIXA  ·  CENÁRIO "&UPPER(Premissas!$C$9)'
put(fc, "B3", "Quando o dinheiro ENTRA e SAI. O plano anual entra inteiro no mês da venda ou da renovação; a última coluna mostra quanto desse dinheiro "
    "ainda é receita dos meses seguintes (não gaste como se fosse lucro).", 9, cor=GRAFITE)
fc.column_dimensions["B"].width = 7
fc.column_dimensions["C"].width = 9
fc.merge_cells(f"D4:F4"); put(fc, "D4", "ENTRADAS", 9, True, "FFFFFF", bg=MARCA_2, al="center")
fc.merge_cells(f"G4:L4"); put(fc, "G4", "SAÍDAS", 9, True, "FFFFFF", bg=MARCA_2, al="center")
fc.merge_cells(f"M4:N4"); put(fc, "M4", "SALDO", 9, True, "FFFFFF", bg=MARCA_2, al="center")
fc.merge_cells(f"O4:P4"); put(fc, "O4", "PLANO ANUAL (COMPETÊNCIA)", 9, True, "FFFFFF", bg=MARCA_2, al="center")
cabecalho(fc, 5, 2, ["Nº", "MÊS"] + [M[k][0] for k in ch_cx], altura=42)
for r in range(FR, LR + 1):
    put(fc, f"B{r}", f"='{PJ}'!{mapa['t'][1]}{r}", fmt="0", al="center", borda=borda_linha)
    put(fc, f"C{r}", f"='{PJ}'!{mapa['mes'][1]}{r}", fmt=MES, al="center", borda=borda_linha)
escreve(fc, mapa, ch_cx, params("F"), larg=15)
put(fc, f"B{LR + 1}", "TOTAL", b=True, borda=borda_total)
fc.merge_cells(f"B{LR + 1}:C{LR + 1}")
for k in ("cxsaldo", "cxac"):
    c = mapa[k][1]
    cor_condicional(fc, f"{c}{FR}:{c}{LR + 1}", f"{c}{FR}<0")
fc.freeze_panes = "D6"
cxac_col = mapa["cxac"][1]
r0 = LR + 3
secao(fc, f"B{r0}", "RESUMO DO CAIXA")
resumo_cx = [
    ("Caixa no fim do mês 24", f"={cxac_col}{LR}", MOEDA),
    ("Menor saldo acumulado (necessidade máxima de capital)", f"=MIN({cxac_col}{FR}:{cxac_col}{LR})", MOEDA),
    ("Mês do menor saldo", f"=INDEX(C{FR}:C{LR},MATCH(MIN({cxac_col}{FR}:{cxac_col}{LR}),{cxac_col}{FR}:{cxac_col}{LR},0))", MES),
    ("Total recebido à vista em planos anuais (24 meses)", f"={mapa['cxa'][1]}{LR + 1}", MOEDA),
    ("Receita anual já recebida e ainda a reconhecer no mês 24", f"={mapa['difer'][1]}{LR}", MOEDA),
]
for k, (rot, f_, fmt) in enumerate(resumo_cx):
    r = r0 + 1 + k
    fc.merge_cells(f"B{r}:F{r}")
    put(fc, f"B{r}", rot, borda=borda_linha)
    put(fc, f"G{r}", f_, b=True, fmt=fmt, al="right", borda=borda_linha)
cor_condicional(fc, f"G{r0 + 1}:G{r0 + 2}", f"G{r0 + 1}<0")

# ---------------------------------------------------------------- UNIT ECONOMICS
ue = base(wb, "Unit Economics", "", "L", GRAFITE,
          larguras={"B": 58, **{L(c): 14 for c in range(3, 12)}})
ue["B2"].value = '="UNIT ECONOMICS  ·  CENÁRIO "&UPPER(Premissas!$C$9)'
put(ue, "B3", "Quanto cada cliente deixa, quanto custa para conquistá-lo e em quantos meses ele se paga. LTV = ticket × margem ÷ churn.", 9, cor=GRAFITE)

S = lambda k: f"'{PJ}'!${mapa[k][1]}${LR + 1}"          # total 24 meses
V = lambda k, r: f"'{PJ}'!${mapa[k][1]}${r}"            # valor de um mês

secao(ue, "B5", "CAC — CUSTO PARA CONQUISTAR UM CLIENTE")
cac_linhas = [
    ("Investimento em aquisição em 24 meses (anúncios + comissões)", f"={S('ads')}+{S('comis')}", MOEDA),
    ("Novos clientes em 24 meses", f"={S('novos')}", NUM),
    ("CAC médio dos 24 meses", "=IFERROR(C6/C7,0)", MOEDA),
    ("CAC no mês 1", f"={V('cac', FR)}", MOEDA),
    ("CAC no mês 24", f"={V('cac', LR)}", MOEDA),
]
for k, (rot, f_, fmt) in enumerate(cac_linhas):
    r = 6 + k
    put(ue, f"B{r}", rot, b=(r == 8), borda=borda_linha)
    put(ue, f"C{r}", f_, b=(r == 8), fmt=fmt, borda=borda_linha)
CAC = "$C$8"

secao(ue, "B12", "RETORNO POR CLIENTE — POR PLANO")
cabecalho(ue, 13, 2, ["PLANO / COBRANÇA", "PESO NO MIX", "RECEITA POR MÊS", "CUSTO VARIÁVEL POR MÊS", "MARGEM POR MÊS", "MARGEM %",
                      "CHURN MENSAL (EQUIV.)", "VIDA MÉDIA", "LTV", "LTV / CAC", "PAYBACK DO CAC"], altura=42)
pp = params("F")
seg = [
    ("Essencial mensal", f"(1-{pp['mixpro']})*(1-{pp['mixanual']})", pp["pEM"], "1", pp["churn"]),
    ("Essencial anual", f"(1-{pp['mixpro']})*{pp['mixanual']}", f"{pp['pEC']}/12", "1/12", f"(1-{pp['renov']})/12"),
    ("Pro mensal", f"{pp['mixpro']}*(1-{pp['mixanual']})", pp["pPM"], "1", pp["churn"]),
    ("Pro anual", f"{pp['mixpro']}*{pp['mixanual']}", f"{pp['pPC']}/12", "1/12", f"(1-{pp['renov']})/12"),
]
for k, (nome, peso, receita, cobr, churn) in enumerate(seg):
    r = 14 + k
    put(ue, f"B{r}", nome, borda=borda_linha)
    put(ue, f"C{r}", f"={peso}", fmt=PCT, borda=borda_linha)
    put(ue, f"D{r}", f"={receita}", fmt=MOEDA, borda=borda_linha)
    put(ue, f"E{r}", f"=D{r}*({pp['imp']}+{pp['gw']})+{cobr}*{pp['gwfix']}+{pp['infra']}", fmt=MOEDA, borda=borda_linha)
    put(ue, f"F{r}", f"=D{r}-E{r}", fmt=MOEDA, borda=borda_linha)
    put(ue, f"G{r}", f"=IFERROR(F{r}/D{r},0)", fmt=PCT, borda=borda_linha)
    put(ue, f"H{r}", f"={churn}", fmt=PCT, borda=borda_linha)
    put(ue, f"I{r}", f"=IFERROR(1/H{r},0)", fmt=MESES, borda=borda_linha)
    put(ue, f"J{r}", f"=IFERROR(D{r}*G{r}/H{r},0)", fmt=MOEDA, borda=borda_linha)
    put(ue, f"K{r}", f"=IFERROR(J{r}/{CAC},0)", fmt=VEZES, borda=borda_linha)
    put(ue, f"L{r}", f"=IFERROR({CAC}/F{r},0)", fmt=MESES, borda=borda_linha)
ue.column_dimensions["L"].width = 14
r = 18
put(ue, f"B{r}", "MÉDIA PONDERADA (mix em uso)", b=True, borda=borda_total)
put(ue, f"C{r}", "=SUM(C14:C17)", b=True, fmt=PCT, borda=borda_total)
for col in "DEFH":
    put(ue, f"{col}{r}", f"=SUMPRODUCT($C$14:$C$17,{col}14:{col}17)", b=True, fmt=MOEDA if col != "H" else PCT, borda=borda_total)
put(ue, f"G{r}", "=IFERROR(F18/D18,0)", b=True, fmt=PCT, borda=borda_total)
put(ue, f"I{r}", "=IFERROR(1/H18,0)", b=True, fmt=MESES, borda=borda_total)
put(ue, f"J{r}", "=IFERROR(D18*G18/H18,0)", b=True, fmt=MOEDA, borda=borda_total)
put(ue, f"K{r}", f"=IFERROR(J18/{CAC},0)", b=True, fmt=VEZES, borda=borda_total)
put(ue, f"L{r}", f"=IFERROR({CAC}/F18,0)", b=True, fmt=MESES, borda=borda_total)
put(ue, "B19", '=IF(K18>=3,"✔ Saudável: cada cliente devolve 3x ou mais o que custou para conquistar.",'
    'IF(K18>=1,"⚠ Atenção: LTV/CAC entre 1x e 3x — o cliente se paga, mas sobra pouco.",'
    '"✖ Cada cliente custa mais do que devolve. Reveja CAC, preço ou churn."))', 10, True)
ue.conditional_formatting.add("B19", FormulaRule(formula=["$K$18>=3"], font=Font(color=VERDE)))
ue.conditional_formatting.add("B19", FormulaRule(formula=["$K$18<1"], font=Font(color=VERMELHO)))
put(ue, "B20", "Plano anual: churn equivalente = (1 − renovação) ÷ 12.  Referência de mercado SaaS: LTV/CAC ≥ 3x e payback ≤ 12 meses.",
    9, cor=GRAFITE)

secao(ue, "B22", "VISÃO REAL DA PROJEÇÃO (24 MESES)")
real = [
    ("Ticket médio real (ARPU = MRR ÷ clientes ativos)", f"=IFERROR({S('mrr')}/SUM('{PJ}'!${mapa['ativos'][1]}${FR}:${mapa['ativos'][1]}${LR}),0)", MOEDA),
    ("Margem bruta real %", f"=IFERROR({S('margem')}/{S('mrr')},0)", PCT),
    ("Churn efetivo médio (cancelados ÷ ativos do mês anterior)",
     f"=IFERROR({S('cancel')}/(SUM('{PJ}'!${mapa['ativos'][1]}${FR}:${mapa['ativos'][1]}${LR})-'{PJ}'!${mapa['ativos'][1]}${LR}),0)", PCT),
    ("LTV (ticket × margem ÷ churn efetivo)", "=IFERROR(C23*C24/C25,0)", MOEDA),
    ("LTV / CAC", f"=IFERROR(C26/{CAC},0)", VEZES),
    ("Payback do CAC (meses)", f"=IFERROR({CAC}/(C23*C24),0)", MESES),
]
for k, (rot, f_, fmt) in enumerate(real):
    r = 23 + k
    put(ue, f"B{r}", rot, borda=borda_linha)
    put(ue, f"C{r}", f_, b=True, fmt=fmt, borda=borda_linha)
put(ue, "D25", "No início quase ninguém cancela (anuais só saem no mês 13), por isso o churn efetivo fica abaixo do churn digitado.",
    9, cor=GRAFITE)

secao(ue, "B30", "PONTO DE EQUILÍBRIO")
eq = [
    ("Margem média por cliente/mês (mix em uso)", "=F18", MOEDA),
    ("Custos fixos no mês 24", f"={V('fixos', LR)}", MOEDA),
    ("Clientes necessários para cobrir os custos fixos", "=IFERROR(ROUNDUP(C32/C31,0),0)", NUM),
    ("Clientes necessários para cobrir fixos + marketing do mês 24", f"=IFERROR(ROUNDUP((C32+{V('ads', LR)}+{V('comis', LR)})/C31,0),0)", NUM),
    ("Clientes ativos no mês 24", f"={V('ativos', LR)}", NUM),
    ("Primeiro mês com lucro líquido positivo",
     f"=IF(COUNT('{PJ}'!${mapa['lucropos'][1]}${FR}:${mapa['lucropos'][1]}${LR})=0,\"Não atinge em 24 meses\","
     f"INDEX('{PJ}'!${mapa['mes'][1]}${FR}:${mapa['mes'][1]}${LR},MIN('{PJ}'!${mapa['lucropos'][1]}${FR}:${mapa['lucropos'][1]}${LR})))", MES),
]
for k, (rot, f_, fmt) in enumerate(eq):
    r = 31 + k
    put(ue, f"B{r}", rot, borda=borda_linha, b=(r == 33))
    put(ue, f"C{r}", f_, b=True, fmt=fmt, borda=borda_linha, al="right")
ue.conditional_formatting.add("C35", FormulaRule(formula=["C35>=C33"], font=Font(color=VERDE)))
ue.conditional_formatting.add("C35", FormulaRule(formula=["C35<C33"], font=Font(color=VERMELHO)))
put(ue, "D33", "Mesmo número que a coluna CLIENTES P/ EQUILÍBRIO da Projeção (mês a mês).", 9, cor=GRAFITE)

# ---------------------------------------------------------------- MOTOR CENÁRIOS
MC = "Motor Cenários"
todas = ch_cli + ch_dre + ch_aux + ch_cx
mc = base(wb, MC, "MOTOR DOS CENÁRIOS  ·  CÁLCULO AUTOMÁTICO — NÃO EDITAR", L(2 + 3 * (len(todas) + 1)), CLARO,
          "Mesma lógica da Projeção e do Fluxo de Caixa, rodando os três cenários ao mesmo tempo para a aba Cenários.")
CENS = [("Pessimista", "C"), ("Realista", "D"), ("Otimista", "E")]
mapas_mc = {}
for i, (nome, col) in enumerate(CENS):
    ini = 2 + i * (len(todas) + 1)
    mp = mapa_bloco(MC, todas, ini)
    mapas_mc[nome] = mp
    c1, c2 = L(ini), L(ini + len(todas) - 1)
    mc.merge_cells(f"{c1}4:{c2}4")
    put(mc, f"{c1}4", nome.upper(), 9, True, "FFFFFF", bg=MARCA_2, al="left")
    for k in todas:
        put(mc, f"{mp[k][1]}5", M[k][0], 8, True, PRETO, bg=SUB, al="center", wrap=True, borda=borda_linha)
    escreve(mc, mp, todas, params(col), larg=12)
    mc.column_dimensions[L(ini + len(todas))].width = 3
mc.row_dimensions[5].height = 54
mc.freeze_panes = "A6"

# ---------------------------------------------------------------- CENÁRIOS
ce = base(wb, "Cenários", "CENÁRIOS  ·  PESSIMISTA × REALISTA × OTIMISTA", "M", MARCA,
          "Os três cenários calculados ao mesmo tempo. As premissas de cada um ficam na aba Premissas (tabela PREMISSAS DE VENDA POR CENÁRIO).",
          {"B": 46, "C": 16, "D": 16, "E": 16, **{L(c): 11 for c in range(6, 14)}})
cabecalho(ce, 5, 2, ["INDICADOR", "PESSIMISTA", "REALISTA", "OTIMISTA"], altura=24)


def mref(nome, k, r):
    return f"'{MC}'!${mapas_mc[nome][k][1]}${r}"


def mrng(nome, k):
    col = mapas_mc[nome][k][1]
    return f"'{MC}'!${col}${FR}:${col}${LR}"


ind = [
    ("Clientes ativos — mês 12", lambda n: f"={mref(n, 'ativos', FR + 11)}", NUM),
    ("Clientes ativos — mês 24", lambda n: f"={mref(n, 'ativos', LR)}", NUM),
    ("MRR — mês 12", lambda n: f"={mref(n, 'mrr', FR + 11)}", MOEDA0),
    ("MRR — mês 24", lambda n: f"={mref(n, 'mrr', LR)}", MOEDA0),
    ("ARR — mês 24", lambda n: f"={mref(n, 'arr', LR)}", MOEDA0),
    ("Ticket médio (ARPU) — mês 24", lambda n: f"={mref(n, 'arpu', LR)}", MOEDA),
    ("Lucro líquido — mês 24", lambda n: f"={mref(n, 'lucro', LR)}", MOEDA0),
    ("Lucro acumulado em 24 meses", lambda n: f"={mref(n, 'lucroac', LR)}", MOEDA0),
    ("Caixa no fim do mês 24", lambda n: f"={mref(n, 'cxac', LR)}", MOEDA0),
    ("Menor saldo de caixa (capital necessário)", lambda n: f"=MIN({mrng(n, 'cxac')})", MOEDA0),
    ("Primeiro mês com lucro",
     lambda n: f"=IF(COUNT({mrng(n, 'lucropos')})=0,\"Não atinge\",INDEX({mrng(n, 'mes')},MIN({mrng(n, 'lucropos')})))", MES),
    ("CAC médio (24 meses)", lambda n: f"=IFERROR(({mref(n, 'ads', LR + 1)}+{mref(n, 'comis', LR + 1)})/{mref(n, 'novos', LR + 1)},0)", MOEDA),
    ("LTV (ticket × margem ÷ churn equivalente)", None, MOEDA),
    ("LTV / CAC", None, VEZES),
]
for k, (rot, fn, fmt) in enumerate(ind):
    r = 6 + k
    put(ce, f"B{r}", rot, borda=borda_linha, b=rot.startswith(("MRR — mês 24", "Lucro líquido", "LTV / CAC")))
    for j, (nome, colp) in enumerate(CENS):
        col = L(3 + j)
        if rot.startswith("LTV ("):
            ch = (f"(Premissas!${colp}${LINHA_CEN['churn']}*(1-Premissas!${colp}${LINHA_CEN['mixanual']})"
                  f"+(1-Premissas!${colp}${LINHA_CEN['renov']})/12*Premissas!${colp}${LINHA_CEN['mixanual']})")
            f_ = f"=IFERROR({mref(nome, 'margem', LR)}/{mref(nome, 'ativos', LR)}/{ch},0)"
        elif rot == "LTV / CAC":
            f_ = f"=IFERROR({col}{r - 1}/{col}{r - 2},0)"
        else:
            f_ = fn(nome)
        put(ce, f"{col}{r}", f_, fmt=fmt, al="right", borda=borda_linha, b=(j == 1))
    if "Lucro" in rot or "Caixa" in rot or "saldo" in rot:
        cor_condicional(ce, f"C{r}:E{r}", f"C{r}<0")
fim = 6 + len(ind)
put(ce, f"B{fim}", "LTV aqui usa a margem por cliente do mês 24 e o churn equivalente do mix (mensal + anual).", 9, cor=GRAFITE)
put(ce, f"B{fim + 1}", f'=IF(ABS(INDEX(C9:E9,MATCH(Premissas!$C$9,$C$5:$E$5,0))-\'{PJ}\'!${mapa["mrr"][1]}${LR})<0.01,'
    '"✔ Conferência: o cenário em uso bate com a aba Projeção 24 meses.","✖ Conferência: diferença entre Motor e Projeção.")', 9, cor=GRAFITE)

# dados dos gráficos (MRR, clientes, caixa por cenário)
g0 = 54
secao(ce, f"B{g0}", "DADOS DOS GRÁFICOS")
cabecalho(ce, g0 + 1, 2, ["MÊS", "MRR PESSIMISTA", "MRR REALISTA", "MRR OTIMISTA", "CLIENTES PESS.", "CLIENTES REAL.", "CLIENTES OTIM.",
                          "CAIXA PESS.", "CAIXA REAL.", "CAIXA OTIM."], altura=30)
for i in range(24):
    r = g0 + 2 + i
    put(ce, f"B{r}", f"={mref('Realista', 'mes', FR + i)}", fmt=MES, al="left", cor=GRAFITE)
    for j, (nome, _) in enumerate(CENS):
        put(ce, f"{L(3 + j)}{r}", f"={mref(nome, 'mrr', FR + i)}", fmt=MOEDA0, cor=GRAFITE)
        put(ce, f"{L(6 + j)}{r}", f"={mref(nome, 'ativos', FR + i)}", fmt=NUM, cor=GRAFITE)
        put(ce, f"{L(9 + j)}{r}", f"={mref(nome, 'cxac', FR + i)}", fmt=MOEDA0, cor=GRAFITE)
gd1, gd2 = g0 + 2, g0 + 25


def grafico_linhas(ws, titulo, c1, c2, r_ini, r_fim, cat_col, ancora, fmt_y, cores, tracos=None):
    ch = LineChart()
    estilo_grafico(ch, titulo, fmt_y)
    data = Reference(ws, min_col=c1, max_col=c2, min_row=r_ini - 1, max_row=r_fim)
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=cat_col, min_row=r_ini, max_row=r_fim))
    for k, s in enumerate(ch.series):
        estilo_serie(s, cores[k], (tracos or [None] * 3)[k], 28000)
    ws.add_chart(ch, ancora)
    return ch


CORES3 = [ERR_COR, SERIE1, OK_COR]
TRACOS3 = ["dash", None, None]
grafico_linhas(ce, "MRR por cenário", 3, 5, gd1, gd2, 2, "G5", '"R$ "#,##0', CORES3, TRACOS3)
grafico_linhas(ce, "Clientes ativos por cenário", 6, 8, gd1, gd2, 2, "G21", "#,##0", CORES3, TRACOS3)
grafico_linhas(ce, "Saldo de caixa acumulado por cenário", 9, 11, gd1, gd2, 2, "G37", '"R$ "#,##0', CORES3, TRACOS3)

# ====================================================================== DASHBOARD (bento)
TILES = [2, 5, 8, 11, 14, 17]          # B, E, H, K, N, Q  (2 colunas por card + 1 de respiro)
larg_db = {}
for c in range(2, 19):
    larg_db[L(c)] = 3 if (c - 1) % 3 == 0 else 12
db = base(wb, "Dashboard", "", "R", MARCA, larguras=larg_db)
db["B2"].value = '="DASHBOARD  ·  "&UPPER(Premissas!$C$6)&"  ·  CENÁRIO "&UPPER(Premissas!$C$9)'
db["B2"].font = font(18, True, "FFFFFF", fam=F_DISPLAY)
put(db, "B3", "Tudo automático, sempre no cenário em uso. Para trocar o cenário ou os números, vá em Premissas.", 9, cor=GRAFITE)
botao(db, "Q3", "Premissas  →", "#'Premissas'!C9", AMBAR_CTA, PRETO, mescla="Q3:R3")

VPJ = lambda k, r: f"'{PJ}'!${mapa[k][1]}${r}"


secao(db, "B5", "PRINCIPAIS NÚMEROS  ·  MÊS 24")
tile(db, TILES[0], 6, "CLIENTES ATIVOS", f"={VPJ('ativos', LR)}", NUM, "no mês 24")
tile(db, TILES[1], 6, "MRR", f"={VPJ('mrr', LR)}", MOEDA0, f"={VPJ('mrr', FR + 11)}", '"mês 12: R$ "#,##0')
tile(db, TILES[2], 6, "ARR", f"={VPJ('arr', LR)}", MOEDA0, "MRR × 12")
tile(db, TILES[3], 6, "LUCRO LÍQUIDO", f"={VPJ('lucro', LR)}", MOEDA0, "no mês 24", neg=True)
ltv = tile(db, TILES[4], 6, "LTV / CAC", "='Unit Economics'!$K$18", VEZES, "ideal ≥ 3x")
tile(db, TILES[5], 6, "PAYBACK DO CAC", "='Unit Economics'!$L$18", '0.0 "meses"', "ideal ≤ 12 meses")
db.conditional_formatting.add(ltv, FormulaRule(formula=[f"AND(ISNUMBER({ltv}),{ltv}>=3)"], font=Font(color=VERDE)))
db.conditional_formatting.add(ltv, FormulaRule(formula=[f"AND(ISNUMBER({ltv}),{ltv}<1)"], font=Font(color=VERMELHO)))

tile(db, TILES[0], 10, "TICKET MÉDIO", f"={VPJ('arpu', LR)}", MOEDA, "ARPU")
tile(db, TILES[1], 10, "MARGEM BRUTA", f"={VPJ('margpct', LR)}", PCT, "após impostos, gateway, infra")
tile(db, TILES[2], 10, "EQUILÍBRIO", "='Unit Economics'!$C$33", '#,##0" clientes"', "cobrem os custos fixos")
tile(db, TILES[3], 10, "1º MÊS COM LUCRO", "='Unit Economics'!$C$36", MES, "lucro líquido > 0")
tile(db, TILES[4], 10, "LUCRO ACUMULADO", f"={VPJ('lucroac', LR)}", MOEDA0, "soma dos 24 meses", neg=True)
tile(db, TILES[5], 10, "MENOR SALDO DE CAIXA", f"='{FC}'!$G${r0 + 2}", MOEDA0, "capital necessário", neg=True)

# card de progresso rumo ao equilíbrio
secao(db, "B14", "RUMO AO EQUILÍBRIO")
db.merge_cells("B15:G15")
put(db, "B15", "Clientes ativos no mês 24  ×  clientes que cobrem os custos fixos", 11, al="left")
db.merge_cells("H15:N15")
put(db, "H15", '=REPT("█",ROUND(MIN(1,MAX(0,N(O15)))*30,0))&REPT("░",30-ROUND(MIN(1,MAX(0,N(O15)))*30,0))', 11, cor=CLARO, al="left")
put(db, "O15", "=IFERROR('Unit Economics'!$C$35/'Unit Economics'!$C$33,0)", 14, True, fmt="0%", al="center")
db.merge_cells("P15:R15")
put(db, "P15", "='Unit Economics'!$C$33", 9, cor=GRAFITE, fmt='"meta: "#,##0" clientes"', al="left")
db.row_dimensions[15].height = 30
card(db, "B", "R", 15, 15)
db.conditional_formatting.add("H15", FormulaRule(formula=["$O$15>=1"], font=Font(color=VERDE)))
db.conditional_formatting.add("H15", FormulaRule(formula=["$O$15<1"], font=Font(color=MOSTARDA_600)))
db.conditional_formatting.add("O15", FormulaRule(formula=["$O$15>=1"], font=Font(color=VERDE)))

secao(db, "B17", "EVOLUÇÃO EM 24 MESES")


def serie_ref(ch, ws, col, cor, tipo="line", traco=None):
    ch.add_data(Reference(ws, min_col=col, min_row=5, max_row=LR), titles_from_data=True)
    s = ch.series[-1]
    if tipo == "line":
        estilo_serie(s, cor, traco, 28000)
    else:
        s.graphicalProperties.solidFill = cor
        s.graphicalProperties.line.noFill = True
    return s


cats = Reference(pj, min_col=openpyxl_col(mapa["mes"][1]), min_row=FR, max_row=LR)
colpj = lambda k: openpyxl_col(mapa[k][1])

c1 = novo("line", "MRR × custos totais", '"R$ "#,##0')
serie_ref(c1, pj, colpj("mrr"), SERIE1)
serie_ref(c1, pj, colpj("custot"), PRETO, traco="dash")
c1.set_categories(cats)
db.add_chart(c1, "B18")

c2 = novo("bar", "Clientes ativos", "#,##0")
serie_ref(c2, pj, colpj("ativos"), SERIE_BAR, tipo="bar")
c2.set_categories(cats)
c2.legend = None
db.add_chart(c2, "K18")

c3 = novo("bar", "Lucro líquido mensal", '"R$ "#,##0')
s3 = serie_ref(c3, pj, colpj("lucro"), SERIE_BAR, tipo="bar")
s3.invertIfNegative = False
c3.set_categories(cats)
c3.legend = None
db.add_chart(c3, "B35")

c4 = novo("line", "Caixa: entradas × saldo acumulado", '"R$ "#,##0')
serie_ref(c4, fc, openpyxl_col(mapa["cxin"][1]), PRETO, traco="dash")
serie_ref(c4, fc, openpyxl_col(mapa["cxac"][1]), SERIE1)
c4.set_categories(Reference(fc, min_col=3, min_row=FR, max_row=LR))
db.add_chart(c4, "K35")

secao(db, "B52", "ALERTAS")
alertas = [
    ("='Unit Economics'!$K$18<3", "LTV/CAC abaixo de 3x — o cliente devolve pouco do que custa", "LTV/CAC saudável (3x ou mais)"),
    ("='Unit Economics'!$L$18>12", "Payback do CAC acima de 12 meses", "Payback do CAC em até 12 meses"),
    (f"='{FC}'!$G${r0 + 2}<0", "O caixa fica negativo em algum mês — é preciso capital inicial", "O caixa nunca fica negativo"),
    (f"={VPJ('lucro', LR)}<0", "Ainda dá prejuízo no mês 24", "Operação lucrativa no mês 24"),
]
AUXA = 22  # V
for k, (cond, ruim, bom) in enumerate(alertas):
    r = 53 + k
    put(db, f"{L(AUXA)}{r}", cond, 8, cor=GRAFITE)
    db.merge_cells(f"B{r}:R{r}")
    put(db, f"B{r}", f'=IF({L(AUXA)}{r},"✖   {ruim}","✔   {bom}")', 11, True, al="left")
    db.row_dimensions[r].height = 26
    # selo (pill): vermelho-claro ou esmeralda-claro, como os chips do app
    db.conditional_formatting.add(f"B{r}:R{r}", FormulaRule(formula=[f"${L(AUXA)}${r}"], font=Font(color=ERR_TXT, bold=True),
                                                            fill=fill(ERR_BG), border=Border(bottom=Side("thin", color=ERR_COR), top=Side("thin", color=ERR_COR))))
    db.conditional_formatting.add(f"B{r}:R{r}", FormulaRule(formula=[f"NOT(${L(AUXA)}${r})"], font=Font(color=OK_TXT, bold=True),
                                                            fill=fill(OK_BG), border=Border(bottom=Side("thin", color=OK_COR), top=Side("thin", color=OK_COR))))
db.column_dimensions[L(AUXA)].hidden = True


# ---------------------------------------------------------------- ordem das abas
ordem = ["Comece Aqui", "Dashboard", "Premissas", PJ, FC, "Unit Economics", "Cenários", MC]
finalizar(wb, ordem, "Comece Aqui")
wb.save(SAIDA)
print("ok ->", SAIDA)
