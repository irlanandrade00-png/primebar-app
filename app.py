from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import openpyxl
import json
import os
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import io
import re
import pdfplumber
import difflib
import unicodedata

app = Flask(__name__, static_folder='static')
CORS(app)

@app.route('/')
def index():
    return send_from_directory(app.static_folder, 'index.html')

@app.route('/<path:path>')
def static_files(path):
    return send_from_directory(app.static_folder, path)

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def get_sheets_service():
    creds_json = os.environ.get('GOOGLE_CREDENTIALS')
    if not creds_json:
        raise Exception("Credenciais do Google nao configuradas.")
    creds_info = json.loads(creds_json)
    creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
    return build('sheets', 'v4', credentials=creds)

# ===========================================================================
# CONSTANTES
# ===========================================================================

MAPA_SUBCAT = {
    'BEBIDAS NAO ALCOOLICAS':  'BEBIDAS NAO ALCOOLICAS',
    'BEBIDAS NÃO ALCOOLICAS':  'BEBIDAS NAO ALCOOLICAS',
    'BEBIDAS ALCOOLICAS':      'BEBIDAS ALCOOLICAS',
    'DESTILADOS':              'DESTILADOS',
    'DOSES':                   'DOSES & OUTROS',
    'DRINKS':                  'DRINK',
    'COMBOS':                  'COMBOS',
    'OUTROS':                  'DOSES & OUTROS',
}

# Produtos cujo YUZER envia categoria diferente da planilha
OVERRIDE_CAT = {
    'GELO SACOLINHA': 'BEBIDAS NAO ALCOOLICAS',
}

CAT_INICIO = {
    'BEBIDAS NAO ALCOOLICAS': 16,
    'BEBIDAS ALCOOLICAS':     32,
    'DESTILADOS':             39,
    'COMBOS':                 52,
    'DRINK':                  68,
    'DOSES & OUTROS':         79,
}

CAT_MAX = {
    'BEBIDAS NAO ALCOOLICAS': 15,
    'BEBIDAS ALCOOLICAS':     6,
    'DESTILADOS':             12,
    'COMBOS':                 15,
    'DRINK':                  10,
    'DOSES & OUTROS':         8,
}

OFFSET_ESTOQUE = -10

OFFSET_PRODUCAO = {
    'BEBIDAS NAO ALCOOLICAS': -11,
    'BEBIDAS ALCOOLICAS':     -10,
    'DESTILADOS':             -10,
    'COMBOS':                 -11,
    'DRINK':                  -11,
    'DOSES & OUTROS':         -10,
}

# ===========================================================================
# PARSERS
# ===========================================================================

def parse_produtos_xlsx(file_bytes):
    """Lê relatório de produtos YUZER. Suporta header em qualquer linha."""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    produtos = []
    header_found = False
    col_map = {}

    for row in ws.iter_rows(values_only=True):
        if not header_found:
            if row[0] == 'Produto':
                header_found = True
                col_map = {str(v).strip(): j for j, v in enumerate(row) if v}
            continue

        if row[0] is None:
            continue

        subcat_idx = col_map.get('Subcategoria', 3)
        if subcat_idx >= len(row) or not row[subcat_idx]:
            continue

        subcat = str(row[subcat_idx]).strip().upper()
        cat    = MAPA_SUBCAT.get(subcat, 'DOSES & OUTROS')
        nome   = str(row[0]).strip()

        if nome in OVERRIDE_CAT:
            cat = OVERRIDE_CAT[nome]

        qtd_idx   = col_map.get('Quantidade', 7)
        preco_idx = col_map.get('Preço', 8)

        try:
            qtd = int(float(str(row[qtd_idx] or 0).replace(',','.')))
        except Exception:
            qtd = 0

        try:
            v = row[preco_idx]
            preco = round(float(v), 2) if isinstance(v, (int, float)) else \
                    round(float(str(v or 0).replace('R$','').replace(',','.').strip()), 2)
        except Exception:
            preco = 0.0

        if qtd <= 0:
            continue

        produtos.append({
            'produto':      nome,
            'subcategoria': subcat,
            'cat':          cat,
            'qtd_vendida':  qtd,
            'preco':        preco,
        })

    return produtos


def _preco_str(s):
    if isinstance(s, (int, float)):
        return round(float(s), 2)
    s = str(s or '0').replace('R$','').replace('\xa0','').strip()
    if ',' in s:
        s = s.replace('.','').replace(',','.')
    try: return round(float(s), 2)
    except: return 0.0


def _normalizar_subcat(s):
    s = str(s).strip().upper().replace('\n',' ')
    if 'NÃO' in s or 'NAO' in s: return 'BEBIDAS NÃO ALCOOLICAS'
    if s in ('BEBIDAS','BEBIDAS ALCOOLICAS'): return 'BEBIDAS ALCOOLICAS'
    return s


def parse_bonus_pdf(file_bytes):
    """Parser PDF de bônus/cortesia do YUZER."""
    produtos = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            for table in (page.extract_tables() or []):
                for row in table:
                    if not row or not row[0]: continue
                    cell0 = str(row[0]).strip()
                    if cell0 == 'NOME' or not cell0 or re.match(r'^[\d\s]+$', cell0): continue

                    if row[1] is not None and row[5] is not None:
                        try:
                            qtd = int(str(row[5]).strip())
                            if qtd <= 0: continue
                            subcat = _normalizar_subcat(row[3] or '')
                            cat = MAPA_SUBCAT.get(subcat, 'DOSES & OUTROS')
                            produtos.append({'produto': cell0,'cat': cat,
                                             'qtd_vendida': qtd,'preco': _preco_str(row[8])})
                        except: pass

                    elif row[1] is None and '\n' in cell0:
                        lines = cell0.split('\n')
                        subcat = _normalizar_subcat(lines[0])
                        cat = MAPA_SUBCAT.get(subcat, 'DOSES & OUTROS')
                        for part in lines[1:]:
                            m = re.match(r'^(.+?)\s+FINAL\s+\S+\s+.+?\s+(\d+)\s+\d+\s+\d+\s+R\$\s*([\d.,]+)', part.strip())
                            if m:
                                qtd = int(m.group(2))
                                if qtd > 0:
                                    produtos.append({'produto': m.group(1).strip(),'cat': cat,
                                                     'qtd_vendida': qtd,
                                                     'preco': round(float(m.group(3).replace('.','').replace(',','.')), 2)})

                    elif row[1] is None and 'FINAL' in cell0:
                        m = re.match(r'^(.+?)\s+FINAL\s+\S+\s+(\S+)\s+\S+\s+(\d+)\s+\d+\s+\d+\s+R\$\s*([\d.,]+)', cell0)
                        if m:
                            qtd = int(m.group(3))
                            if qtd > 0:
                                subcat = _normalizar_subcat(m.group(2))
                                cat = MAPA_SUBCAT.get(subcat,'DOSES & OUTROS')
                                produtos.append({'produto': m.group(1).strip(),'cat': cat,
                                                 'qtd_vendida': qtd,
                                                 'preco': round(float(m.group(4).replace('.','').replace(',','.')), 2)})
    return produtos


def parse_caixas(file_bytes):
    """
    Lê exportação de caixas YUZER.
    Cols: [0]=Id [1]=Usuário [3]=Serial [5]=Operação [6]=Total
          [13]=Crédito [14]=Débito [15]=PIX [16]=Dinheiro [12]=Retornado
    """
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    caixas = []
    col_map = {}
    header_found = False

    def norm_key(s):
        s = str(s).strip().lower()
        for a, b in [('á','a'),('â','a'),('ã','a'),('é','e'),('ê','e'),
                     ('í','i'),('ó','o'),('ô','o'),('õ','o'),('ú','u'),('ç','c')]:
            s = s.replace(a, b)
        return s

    for row in ws.iter_rows(values_only=True):
        if not header_found:
            if row[0] == 'Id':
                header_found = True
                col_map = {norm_key(v): j for j, v in enumerate(row) if v}
            continue

        if not row[0] or len(str(row[0]).strip()) < 15:
            continue

        def gcol(names, fallback):
            for n in names:
                if n in col_map and col_map[n] < len(row):
                    try: return round(float(str(row[col_map[n]] or 0).replace(',','.').replace('R$','')), 2)
                    except: pass
            try: return round(float(row[fallback] or 0), 2)
            except: return 0.0

        def scol(names, fallback):
            for n in names:
                if n in col_map and col_map[n] < len(row):
                    return str(row[col_map[n]] or '').strip()
            return str(row[fallback] or '').strip() if fallback < len(row) else ''

        dinheiro_bruto = gcol(['dinheiro'], 16)
        devolvido      = gcol(['total produtos retornados','total retornado'], 12)
        dinheiro_liq   = round(max(0.0, dinheiro_bruto - devolvido), 2)

        caixas.append({
            'usuario':  scol(['usuario'], 1),
            'serial':   scol(['serial'], 3),
            'operacao': scol(['operacao'], 5),
            'total':    gcol(['total'], 6),
            'credito':  gcol(['credito'], 13),
            'debito':   gcol(['debito'], 14),
            'pix':      gcol(['pix'], 15),
            'dinheiro': dinheiro_liq,
        })

    return caixas


def parse_painel_vendas(file_bytes):
    """
    Lê painel de vendas YUZER.
    Lê apenas formas principais (PIX, DEBIT_CARD, CREDIT_CARD, CASH).
    Para na linha 'Total por bandeira' para não somar sub-bandeiras.
    """
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    painel = {}
    formas = {}
    lendo_formas = False
    passou_operacoes = False

    FORMAS_PRINCIPAIS = {'PIX', 'DEBIT_CARD', 'CREDIT_CARD', 'CASH', 'APP', 'CASHLESS'}

    for row in ws.iter_rows(values_only=True):
        if row[0] is None: continue
        key = str(row[0]).strip()
        val = row[1] if len(row) > 1 else None

        if key == 'Formas de Pagamento':
            lendo_formas = True
            continue

        if key.startswith('Total por bandeira') or key in ('Operacoes','Operações'):
            lendo_formas = False
            if 'Opera' in key: passou_operacoes = True
            continue

        if lendo_formas and val is not None and key in FORMAS_PRINCIPAIS:
            try: formas[key] = round(float(val or 0), 2)
            except: pass
            continue

        if not passou_operacoes and key in ('Total','Pedidos','Média','Media','Ticket'):
            if key not in painel:
                painel[key] = val

    painel['formas_pagamento'] = formas
    return painel


# ===========================================================================
# LEITURA DA PLANILHA GOOGLE
# ===========================================================================

def ler_cadastro(service, spreadsheet_id):
    """Lê CADASTRO col B (nome) e F (preço) por categoria."""
    catalogo = {cat: [] for cat in CAT_INICIO}
    for cat, inicio in CAT_INICIO.items():
        fim = inicio + CAT_MAX[cat] - 1
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"CADASTRO!B{inicio}:F{fim}"
        ).execute()
        for i, row in enumerate(result.get('values', [])):
            nome = str(row[0]).strip() if len(row) > 0 and row[0] else ''
            if not nome: continue
            try:
                raw = row[4] if len(row) > 4 else 0
                preco = round(float(raw), 2) if isinstance(raw, (int,float)) else \
                        round(float(str(raw).replace('R$','').replace(',','.').strip()), 2)
            except Exception:
                preco = 0.0
            catalogo[cat].append({
                'nome':           nome,
                'preco':          preco,
                'linha_cadastro': inicio + i,
            })
    return catalogo


def ler_mapa_linhas(service, spreadsheet_id):
    """
    Lê col A de ESTOQUE e PRODUÇÃO e retorna nome→linha.
    Usa fórmulas resolvidas para pegar o nome real do produto.
    """
    IGNORAR = {
        'PRODUTO','BEBIDAS NÃO ALCOOLICAS','BEBIDAS ALCOOLICAS','DESTILADOS',
        'COMBOS','DRINK','DOSES & OUTROS','FECHAMENTO GERAL BAR CONSUMO/VENDA',
        'OBSERVAÇÃO PREENCHER APENAS AS COLUNAS EM AMARELO',
        'CONSUMO PRODUÇÃO CAMARIM / BONUS','RESUMO ALIMENTAÇAO',
        'TOTAL / CARTÃO','TOTAL',
    }
    est_map = {}
    prod_map = {}

    r = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range="ESTOQUE!A1:A80"
    ).execute()
    for i, row in enumerate(r.get('values', []), 1):
        if row and row[0] and str(row[0]).strip() not in IGNORAR:
            est_map[str(row[0]).strip()] = i

    r = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range="PRODUÇÃO!A1:A80"
    ).execute()
    for i, row in enumerate(r.get('values', []), 1):
        if row and row[0] and str(row[0]).strip() not in IGNORAR:
            prod_map[str(row[0]).strip()] = i

    return est_map, prod_map


# ===========================================================================
# MOTOR DE CONCILIAÇÃO MULTI-ATRIBUTO
# ===========================================================================

STOP_WORDS = {
    'ml','l','lt','kg','g','un','und','cx',
    '100','130','150','160','180','200','220','250','260','269',
    '280','290','300','320','330','340','350','355','360','370',
    '380','410','430','440','473','500','550','600','750','1000',
    'com','de','do','da','e','em','para','o','a','os','as',
    'final','bebidas','drink','dose','doses','combo','outros',
    'garrafa','lata','long','neck','anos',
    '1','2','3','4','5','6','7','8','9','10',
}

ALIAS = {
    'buul':'bull','redbull':'red bull','redbuul':'red bull',
    'oldpar':'old parr','oldparr':'old parr',
    'tanquery':'tanqueray','moscow':'moscow','mulle':'mule',
    'limonade':'lemonade','tonica':'tonica','tonicas':'tonica',
}

def _norm_str(s):
    s = unicodedata.normalize('NFKD', str(s))
    s = ''.join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r'[^a-z0-9 ]', ' ', s.lower())
    for a, b in ALIAS.items():
        s = re.sub(r'\b' + a + r'\b', b, s)
    return s

def _tokens(nome):
    return set(t for t in _norm_str(nome).split() if t not in STOP_WORDS and len(t) > 1)

def _extrair_ml(nome):
    m = re.search(r'(\d+\.?\d*)\s*(ml|l|lt)\b', nome.lower().replace(' ',''))
    if m:
        val = float(m.group(1))
        return val * 1000 if m.group(2) in ('l','lt') else val
    return None

def _score_par(nome_p, preco_p, nome_y, preco_y):
    """Score combinado: 35% preço + 45% nome + 20% unidade."""
    # Preço
    if preco_p > 0 and preco_y > 0:
        s_preco = 1.0 if abs(preco_p-preco_y) < 0.01 else \
                  max(0.0, 1.0 - abs(preco_p-preco_y)/max(preco_p,preco_y)*3)
    else:
        s_preco = 0.0

    # Nome
    s_seq = difflib.SequenceMatcher(None, _norm_str(nome_p), _norm_str(nome_y)).ratio()
    tp, ty = _tokens(nome_p), _tokens(nome_y)
    s_tok = len(tp&ty)/len(tp|ty) if (tp and ty) else 0.0
    s_nome = s_seq*0.6 + s_tok*0.4

    # Unidade
    ml_p, ml_y = _extrair_ml(nome_p), _extrair_ml(nome_y)
    s_un = 1.0 if (ml_p and ml_y and abs(ml_p-ml_y)<1) else 0.5

    return s_preco*0.35 + s_nome*0.45 + s_un*0.20

def _similaridade(a, b):
    return _score_par(a, 0, b, 0)

def _normalizar(nome):
    return _norm_str(nome)

LIMIAR_SCORE = 0.40

def conciliar(catalogo, vendas, bonus, mapeamento=None):
    """
    Pré-match global greedy por score descendente.
    Aplica mapeamento fixo antes de calcular scores.
    Retorna detalhes de conciliação para preview e validação.
    """
    mapeamento = mapeamento or {}
    agrupado = {cat: [] for cat in CAT_INICIO}

    def aplicar_mapa(lista):
        for p in lista:
            if p['produto'] in mapeamento:
                p = dict(p)
                p['produto'] = mapeamento[p['produto']]
            yield p

    vendas_map = list(aplicar_mapa(vendas))
    bonus_map  = list(aplicar_mapa(bonus))

    def pre_match(planilha_items, yuzer_items):
        if not planilha_items or not yuzer_items:
            return {}
        pares = []
        for pi, item in enumerate(planilha_items):
            for yi, prod in enumerate(yuzer_items):
                s = _score_par(item['nome'], item['preco'],
                               prod['produto'], prod['preco'])
                if s >= LIMIAR_SCORE:
                    pares.append((s, pi, yi))
        pares.sort(reverse=True)
        plan_usado = set(); yuzer_usado = set(); resultado = {}
        for s, pi, yi in pares:
            if pi not in plan_usado and yi not in yuzer_usado:
                resultado[pi] = (yi, yuzer_items[yi]['produto'], s,
                                 yuzer_items[yi]['qtd_vendida'])
                plan_usado.add(pi); yuzer_usado.add(yi)
        return resultado

    for cat, itens in catalogo.items():
        v_cat = [p for p in vendas_map if p['cat'] == cat]
        b_cat = [p for p in bonus_map  if p['cat'] == cat]
        match_v = pre_match(itens, v_cat)
        match_b = pre_match(itens, b_cat)

        for pi, item in enumerate(itens):
            mv = match_v.get(pi)
            mb = match_b.get(pi)
            agrupado[cat].append({
                'nome':           item['nome'],
                'linha_cadastro': item['linha_cadastro'],
                'preco':          item['preco'],
                'qtd_venda':      mv[3] if mv else 0,
                'qtd_bonus':      mb[3] if mb else 0,
                'qtd_sistema':    (mv[3] if mv else 0) + (mb[3] if mb else 0),
                'match_venda':    mv[1] if mv else None,
                'score_venda':    mv[2] if mv else 0.0,
                'match_bonus':    mb[1] if mb else None,
                'score_bonus':    mb[2] if mb else 0.0,
                'conciliado':     mv is not None or mb is not None,
            })

    return agrupado


def gerar_sugestoes_mapeamento(vendas, bonus, catalogo):
    """
    Para cada produto YUZER sem match, sugere o melhor candidato da planilha.
    Retorna lista de sugestões para o usuário confirmar no preview.
    """
    todos_yuzer = vendas + bonus
    todos_planilha = [p for prods in catalogo.values() for p in prods]
    sugestoes = []

    for yu in todos_yuzer:
        # Calcular score contra todos os produtos da planilha
        scores = []
        for pl in todos_planilha:
            s = _score_par(pl['nome'], pl['preco'], yu['produto'], yu['preco'])
            if s >= 0.30:
                scores.append((s, pl['nome'], pl['preco']))
        scores.sort(reverse=True)

        sugestoes.append({
            'yuzer':      yu['produto'],
            'preco_yuzer': yu['preco'],
            'cat':        yu['cat'],
            'qtd':        yu['qtd_vendida'],
            'sugestao':   scores[0][1] if scores else None,
            'score':      round(scores[0][0], 2) if scores else 0,
            'alternativas': [{'nome': s[1], 'preco': s[2], 'score': round(s[0],2)} 
                             for s in scores[1:4]],
        })

    return sugestoes


# ===========================================================================
# BUILDERS GOOGLE SHEETS
# ===========================================================================

def build_estoque_updates(agrupado, est_map=None):
    """ESTOQUE col I = qtd_venda (só vendas, sem bônus)."""
    updates = []
    nao_conciliados = []
    for prods in agrupado.values():
        for p in prods:
            linha = p['linha_cadastro'] + OFFSET_ESTOQUE
            updates.append({'range': f"ESTOQUE!I{linha}", 'values': [[p['qtd_venda']]]})
            if not p.get('conciliado', True) and p['preco'] > 0:
                nao_conciliados.append(f"{p['nome']} (R${p['preco']:.2f})")
    return updates, nao_conciliados


def build_producao_updates(agrupado, prod_map=None):
    """PRODUÇÃO col C = qtd_bonus. Usa prod_map real ou offset como fallback."""
    updates = []
    avisos  = []
    for cat, prods in agrupado.items():
        for p in prods:
            if p['qtd_bonus'] <= 0: continue
            nome  = p['nome']
            linha = prod_map.get(nome) if prod_map else None
            if linha is None:
                linha = p['linha_cadastro'] + OFFSET_PRODUCAO.get(cat, -10)
            if linha and linha > 0:
                updates.append({'range': f"PRODUÇÃO!C{linha}", 'values': [[p['qtd_bonus']]]})
            if p.get('score_bonus', 1.0) < 0.40 and p.get('match_bonus'):
                avisos.append(f"Bônus incerto: '{nome}' → '{p['match_bonus']}' ({int(p['score_bonus']*100)}%)")
    return updates, avisos


# ===========================================================================
# MAPEAMENTO POR PLANILHA (item 2: salvar por spreadsheet_id)
# ===========================================================================
# Estrutura: { spreadsheet_id: { "yuzer_nome": "planilha_nome" } }
_mapeamento_store = {}

# Mapeamento global padrão (aprendido entre eventos)
_mapeamento_global = {
    'DRINK Tropical Gin':    'TROPICAL GIN ( GIN + RODELA DE LARANJA E RED BUUL TROPICAL )',
    'DRINK Melancita':       'MELANCITA ( GIN + RODELA DE LIMÃO E RED BUUL MELANCIA )',
    'DRINK Moscow Mule':     'MOSCOW MULLE ( VODKA + XAROPE DE GENGIBRE + SUMO DE LIMÃO E ESPUMA CITRICA )',
    'DRINK Pink Limonade':   'PINK LEMONADE ( GIN +  SUCO DE  LIMÃO + GROSELHA E RODELA DE LIMÃO SICILIANO)',
    'DRINK Gija':            'GIJA ( GIN + TONICA + XAROPE DE GENGIBRE + CANELA E RODELA DE LIMÃO SICILIANO )',
    'DRINK Gin Tônica':      'GIN TONICA ( GIN + TONICA E RODELA DE LIMÃO )',
    'DRINK Vodka + Red Bull':'VODKA E RED BUUL (VODKA + RED BUUL + ESCOLHA SEU SABOR )',
    'Old Parr+3 Red Bull':   'OLDPAR 12 ANOS 1L  + 3 REDBULL 250ML',
    'Old Parr+5 Águas de Coco': 'OLDPARR 12 ANOS 1L + 5 AGUA DE COCO',
}

def get_mapa(spreadsheet_id):
    """Retorna mapeamento: global + específico da planilha."""
    mapa = dict(_mapeamento_global)
    mapa.update(_mapeamento_store.get(spreadsheet_id, {}))
    return mapa


# ===========================================================================
# LIMPEZA DA PLANILHA
# ===========================================================================

def limpar_planilha(service, spreadsheet_id):
    """
    Zera apenas os campos que o sistema preenche.
    NÃO toca em RELATORIO DE VENDA (tem fórmulas).
    NÃO toca em col D do RESUMO (fórmula de FECHAMENTO CAIXAS).
    """
    ranges = [
        'RESUMO!B3:B9',
        'ESTOQUE!I6:I76',
        'PRODUÇÃO!C5:C77',
        'FECHAMENTO CAIXAS!B3:H32',
        'FECHAMENTO CAIXAS!B36:H50',
        'FECHAMENTO CAIXAS!B54:H83',
    ]
    service.spreadsheets().values().batchClear(
        spreadsheetId=spreadsheet_id,
        body={'ranges': ranges}
    ).execute()

    # PRODUÇÃO com acento — limpar separado
    try:
        service.spreadsheets().values().clear(
            spreadsheetId=spreadsheet_id,
            range='PRODUÇÃO!C5:C77'
        ).execute()
    except Exception:
        pass


# ===========================================================================
# VALIDAÇÃO DE TOTAIS
# ===========================================================================

def validar_totais(agrupado, painel):
    """Compara soma calculada vs total do painel. Retorna lista de avisos."""
    avisos = []
    total_painel = float(painel.get('Total', 0) or 0)
    if total_painel == 0: return avisos

    total_calc = sum(
        p['qtd_venda'] * p['preco']
        for prods in agrupado.values()
        for p in prods
    )
    if total_calc == 0: return avisos

    diferenca = abs(total_painel - total_calc)
    pct = diferenca / total_painel * 100
    if diferenca > 1.0:
        avisos.append(
            f"Divergência de totais: Painel=R${total_painel:,.2f} | "
            f"Calculado=R${total_calc:,.2f} | Diferença=R${diferenca:,.2f} ({pct:.1f}%)"
        )
    return avisos


# ===========================================================================
# RELATÓRIO DE CONCILIAÇÃO (item 6: download PDF)
# ===========================================================================

def gerar_relatorio_texto(agrupado, msgs, avisos, painel_data, caixas):
    """Gera texto do relatório de fechamento."""
    fp  = painel_data.get('formas_pagamento', {})
    tot = painel_data.get('Total', 0)
    linhas = [
        "PRIME BAR — RELATÓRIO DE FECHAMENTO",
        "=" * 50,
        "",
        "PAGAMENTOS",
        f"  Total:   R${float(tot or 0):>12,.2f}",
        f"  Dinheiro:R${fp.get('CASH',0):>12,.2f}",
        f"  Crédito: R${fp.get('CREDIT_CARD',0):>12,.2f}",
        f"  Débito:  R${fp.get('DEBIT_CARD',0):>12,.2f}",
        f"  PIX:     R${fp.get('PIX',0):>12,.2f}",
        "",
        "PRODUTOS PREENCHIDOS (ESTOQUE col I)",
        f"  {'Produto':<40} {'Qtd':>5} {'Score':>6}",
        "  " + "-"*54,
    ]
    for prods in agrupado.values():
        for p in prods:
            if p['qtd_venda'] > 0:
                s = p.get('score_venda', 0)
                flag = ' ⚠' if s < 0.5 else ''
                linhas.append(f"  {p['nome'][:39]:<40} {p['qtd_venda']:>5} {s:>5.2f}{flag}")

    linhas += ["", "BÔNUS/CORTESIA (PRODUÇÃO col C)"]
    for prods in agrupado.values():
        for p in prods:
            if p['qtd_bonus'] > 0:
                linhas.append(f"  {p['nome'][:39]:<40} {p['qtd_bonus']:>5}")

    nconc = [p for prods in agrupado.values() for p in prods
             if not p.get('conciliado') and p['preco'] > 0]
    if nconc:
        linhas += ["", "SEM VENDA NO YUZER (preencher manualmente se necessário)"]
        for p in nconc:
            linhas.append(f"  {p['nome']} R${p['preco']:.2f}")

    if avisos:
        linhas += ["", "AVISOS"]
        for a in avisos:
            linhas.append(f"  ⚠ {a}")

    if msgs:
        linhas += ["", "LOG DE ENVIO"]
        for m in msgs:
            linhas.append(f"  ✓ {m}")

    return "\n".join(linhas)


# ===========================================================================
# ROTAS
# ===========================================================================

@app.route('/api/preview', methods=['POST'])
def preview():
    try:
        result  = {}
        vendas  = []
        bonus   = []

        # Suporte a múltiplos arquivos de vendas (item 4)
        vendas_files = request.files.getlist('produtos_vendidos')
        for f in vendas_files:
            vendas += parse_produtos_xlsx(f.read())
        result['produtos'] = vendas

        if 'produtos_bonus' in request.files:
            b = request.files['produtos_bonus']
            bonus = parse_bonus_pdf(b.read()) if b.filename.lower().endswith('.pdf') \
                    else parse_produtos_xlsx(b.read())
            result['bonus'] = bonus

        if 'exportacao_caixas' in request.files:
            result['caixas'] = parse_caixas(request.files['exportacao_caixas'].read())

        if 'painel_de_vendas' in request.files:
            painel = parse_painel_vendas(request.files['painel_de_vendas'].read())
            fp = painel.get('formas_pagamento', {})
            result['painel']  = painel
            result['resumo'] = {
                'total_faturado': painel.get('Total', 0),
                'total_pedidos':  painel.get('Pedidos', 0),
                'ticket_medio':   painel.get('Média', painel.get('Media', 0)),
                'credito':  fp.get('CREDIT_CARD', 0),
                'debito':   fp.get('DEBIT_CARD', 0),
                'pix':      fp.get('PIX', 0),
                'dinheiro': fp.get('CASH', 0),
            }

        # Sugestões de mapeamento automático (item 1)
        spreadsheet_id = request.form.get('spreadsheet_id', '').strip()
        if spreadsheet_id and (vendas or bonus):
            try:
                sid = spreadsheet_id
                if 'docs.google.com' in sid:
                    m = re.search(r'/d/([a-zA-Z0-9-_]+)', sid)
                    if m: sid = m.group(1)

                service  = get_sheets_service()
                catalogo = ler_cadastro(service, sid)
                mapa_atual = get_mapa(sid)
                agrupado = conciliar(catalogo, vendas, bonus, mapa_atual)

                # Sugestões para produtos sem match
                sugestoes = gerar_sugestoes_mapeamento(vendas, bonus, catalogo)
                result['sugestoes_mapeamento'] = sugestoes

                # Resumo de conciliação para validação (item 3)
                total_prod  = sum(len(v) for v in agrupado.values())
                conciliados = sum(1 for ps in agrupado.values() for p in ps if p.get('conciliado'))
                sem_match   = [p for ps in agrupado.values() for p in ps
                               if not p.get('conciliado') and p['preco'] > 0]

                result['validacao'] = {
                    'total_produtos':  total_prod,
                    'conciliados':     conciliados,
                    'sem_match':       [{'nome': p['nome'], 'preco': p['preco']} for p in sem_match],
                    'score_baixo':     [{'planilha': p['nome'],
                                         'yuzer': p['match_venda'],
                                         'score': p['score_venda']}
                                        for ps in agrupado.values() for p in ps
                                        if p.get('match_venda') and 0 < p['score_venda'] < 0.50],
                    'pode_enviar':     True,
                }
            except Exception as e_prev:
                result['validacao'] = {'erro_preview': str(e_prev), 'pode_enviar': True}

        return jsonify({'success': True, 'data': result})
    except Exception as e:
        import traceback
        return jsonify({'success': False, 'error': str(e),
                        'trace': traceback.format_exc()}), 400


@app.route('/api/mapeamento', methods=['GET'])
def get_mapeamento():
    spreadsheet_id = request.args.get('spreadsheet_id', '')
    if spreadsheet_id and 'docs.google.com' in spreadsheet_id:
        m = re.search(r'/d/([a-zA-Z0-9-_]+)', spreadsheet_id)
        if m: spreadsheet_id = m.group(1)
    mapa = get_mapa(spreadsheet_id)
    return jsonify({'success': True, 'mapeamento': mapa, 'global': _mapeamento_global,
                    'especifico': _mapeamento_store.get(spreadsheet_id, {})})


@app.route('/api/mapeamento', methods=['POST'])
def save_mapeamento():
    data = request.get_json(force=True) or {}
    spreadsheet_id = data.get('spreadsheet_id', 'global').strip()
    if spreadsheet_id and 'docs.google.com' in spreadsheet_id:
        m = re.search(r'/d/([a-zA-Z0-9-_]+)', spreadsheet_id)
        if m: spreadsheet_id = m.group(1)
    mapa = data.get('mapeamento', {})

    if spreadsheet_id == 'global':
        _mapeamento_global.update(mapa)
    else:
        if spreadsheet_id not in _mapeamento_store:
            _mapeamento_store[spreadsheet_id] = {}
        _mapeamento_store[spreadsheet_id].update(mapa)

    # Aprendizado: persistir no global também (item 5)
    if data.get('aprender', False):
        _mapeamento_global.update(mapa)

    return jsonify({'success': True, 'total_global': len(_mapeamento_global),
                    'total_especifico': len(_mapeamento_store.get(spreadsheet_id, {}))})


@app.route('/api/enviar', methods=['POST'])
def enviar():
    try:
        spreadsheet_id = request.form.get('spreadsheet_id', '').strip()
        if not spreadsheet_id:
            return jsonify({'success': False, 'error': 'ID da planilha não informado.'}), 400
        if 'docs.google.com' in spreadsheet_id:
            m = re.search(r'/d/([a-zA-Z0-9-_]+)', spreadsheet_id)
            if m: spreadsheet_id = m.group(1)

        service = get_sheets_service()
        batch   = []
        msgs    = []
        avisos  = []
        agrupado    = {}
        painel_data = {}

        # Limpeza prévia
        try:
            limpar_planilha(service, spreadsheet_id)
            msgs.append('Planilha limpa (campos automáticos zerados)')
        except Exception as e_limpa:
            avisos.append(f'Limpeza falhou: {str(e_limpa)[:80]}')

        # Produtos vendidos + bônus
        # Suporte a múltiplos arquivos de vendas (item 4)
        vendas_files = request.files.getlist('produtos_vendidos')
        if vendas_files:
            vendas = []
            for f in vendas_files:
                vendas += parse_produtos_xlsx(f.read())

            bonus = []
            if 'produtos_bonus' in request.files:
                b = request.files['produtos_bonus']
                bonus = parse_bonus_pdf(b.read()) if b.filename.lower().endswith('.pdf') \
                        else parse_produtos_xlsx(b.read())

            # Mapeamento: global + específico desta planilha
            mapa_atual = get_mapa(spreadsheet_id)
            msgs.append(f'Mapeamento: {len(mapa_atual)} entradas ativas')

            # Catálogo e mapa de linhas
            catalogo = ler_cadastro(service, spreadsheet_id)
            total_cat = sum(len(v) for v in catalogo.values())
            msgs.append(f'CADASTRO: {total_cat} produtos lidos')

            _, prod_map_real = ler_mapa_linhas(service, spreadsheet_id)
            msgs.append(f'PRODUÇÃO: {len(prod_map_real)} linhas mapeadas')

            # Conciliação
            agrupado = conciliar(catalogo, vendas, bonus, mapa_atual)

            # Estatísticas de conciliação
            total_prod  = sum(len(v) for v in agrupado.values())
            conciliados = sum(1 for ps in agrupado.values() for p in ps if p.get('conciliado'))
            nao_conc    = sum(1 for ps in agrupado.values() for p in ps
                              if not p.get('conciliado') and p['preco'] > 0)
            msgs.append(f'Conciliação: {conciliados}/{total_prod} produtos encontrados')

            if nao_conc:
                nomes_nconc = [
                    f"{p['nome']} (R${p['preco']:.2f})"
                    for prods in agrupado.values() for p in prods
                    if not p.get('conciliado') and p['preco'] > 0
                ]
                avisos.append(
                    f'{nao_conc} produto(s) sem venda no YUZER (zerado — preencha manualmente): '
                    f'{"; ".join(nomes_nconc[:7])}'
                )

            # Matches com score muito baixo
            ruins = [
                f"'{p['nome']}' → '{p['match_venda']}' ({int(p['score_venda']*100)}%)"
                for prods in agrupado.values() for p in prods
                if p.get('match_venda') and 0 < p.get('score_venda', 1) < 0.40
            ]
            if ruins:
                avisos.append(f'Matches com score baixo — revisar: {"; ".join(ruins[:5])}')

            # ESTOQUE col I
            est_updates, est_nf = build_estoque_updates(agrupado)
            batch.extend(est_updates)
            msgs.append(f'ESTOQUE col I: {len(est_updates)} produtos preenchidos')
            for a in est_nf: avisos.append(a)

            # PRODUÇÃO col C
            prod_updates, prod_avisos = build_producao_updates(agrupado, prod_map_real)
            batch.extend(prod_updates)
            msgs.append(f'PRODUÇÃO col C: {len(prod_updates)} bônus preenchidos')
            for a in prod_avisos: avisos.append(a)

        # Caixas
        if 'exportacao_caixas' in request.files:
            caixas = parse_caixas(request.files['exportacao_caixas'].read())

            def op_norm(s):
                return str(s).upper().replace('Ç','C').replace('Ã','A').strip()

            # L3:L32 = GARÇOM PIX | L36:L50 = Caixa PIX | L54:L83 = Garçons crachá
            garcons_pix = [c for c in caixas if 'GARCOM' in op_norm(c['operacao'])]
            caixas_pix  = [c for c in caixas if 'CAIXA'  in op_norm(c['operacao'])]

            def to_rows(lista):
                return [[c['usuario'], c['serial'], c['total'],
                         c['dinheiro'], c['pix'], c['debito'], c['credito']]
                        for c in lista]

            if garcons_pix:
                rows = to_rows(garcons_pix)
                batch.append({'range': f"FECHAMENTO CAIXAS!B3:H{2+len(rows)}", 'values': rows})
                msgs.append(f'FECHAMENTO CAIXAS — Garçons: {len(rows)}')

            if caixas_pix:
                rows = to_rows(caixas_pix)
                batch.append({'range': f"FECHAMENTO CAIXAS!B36:H{35+len(rows)}", 'values': rows})
                msgs.append(f'FECHAMENTO CAIXAS — Caixas Fixos: {len(rows)}')

        # Painel → RESUMO col B
        if 'painel_de_vendas' in request.files:
            painel_data = parse_painel_vendas(request.files['painel_de_vendas'].read())
            fp = painel_data.get('formas_pagamento', {})
            total_fp = fp.get('CASH',0)+fp.get('CREDIT_CARD',0)+fp.get('DEBIT_CARD',0)+fp.get('PIX',0)
            batch.append({
                'range': 'RESUMO!B3:B9',
                'values': [
                    [0],
                    [fp.get('CASH',0)],
                    [fp.get('CREDIT_CARD',0)],
                    [fp.get('DEBIT_CARD',0)],
                    [fp.get('PIX',0)],
                    [0],
                    [total_fp],
                ],
            })
            msgs.append('RESUMO col B: formas de pagamento preenchidas')

            # Validação de totais
            if agrupado:
                for a in validar_totais(agrupado, painel_data):
                    avisos.append(a)

        # Enviar
        if batch:
            service.spreadsheets().values().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={'valueInputOption': 'USER_ENTERED', 'data': batch}
            ).execute()

        # Gerar relatório de fechamento (item 6)
        relatorio_txt = gerar_relatorio_texto(
            agrupado, msgs, avisos, painel_data,
            request.files.getlist('exportacao_caixas')
        )

        return jsonify({
            'success':   True,
            'message':   'Dados enviados com sucesso!',
            'detalhes':  msgs,
            'avisos':    avisos,
            'relatorio': relatorio_txt,
        })

    except Exception as e:
        import traceback
        return jsonify({'success': False, 'error': str(e),
                        'trace': traceback.format_exc()}), 400


@app.route('/api/health')
def health():
    return jsonify({'status': 'ok', 'app': 'Prime Bar YUZER v5',
                    'mapeamentos_globais': len(_mapeamento_global),
                    'planilhas_mapeadas': len(_mapeamento_store)})


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
