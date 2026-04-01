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

# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------

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

# Produtos do YUZER que vêm com categoria errada mas pertencem a outra cat na planilha
# Chave = nome do produto YUZER, valor = categoria correta da planilha
OVERRIDE_CAT = {
    'GELO SACOLINHA': 'BEBIDAS NAO ALCOOLICAS',
}

# Categorias e linha inicial no CADASTRO
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

# Cabeçalhos/categorias a ignorar ao ler nomes de produtos nas abas
IGNORAR_NOMES = {
    'PRODUTO', 'BEBIDAS NÃO ALCOOLICAS', 'BEBIDAS ALCOOLICAS',
    'DESTILADOS', 'COMBOS', 'DRINK', 'DOSES & OUTROS',
    'FECHAMENTO GERAL BAR CONSUMO/VENDA',
    'OBSERVAÇÃO PREENCHER APENAS AS COLUNAS EM AMARELO',
    'CONSUMO PRODUÇÃO CAMARIM / BONUS',
    'RESUMO ALIMENTAÇAO', 'TOTAL / CARTÃO', 'TOTAL',
}

# Offsets reais medidos na planilha SAMBA_NO_PARQUE_TESTE123
# CADASTRO linha X -> ESTOQUE col I linha (X + OFFSET_ESTOQUE)
OFFSET_ESTOQUE = -10  # igual para todas as categorias

# CADASTRO linha X -> PRODUÇÃO col C linha (X + OFFSET_PRODUCAO[cat])
OFFSET_PRODUCAO = {
    'BEBIDAS NAO ALCOOLICAS': -11,  # CAD L16 -> PROD L5
    'BEBIDAS ALCOOLICAS':     -10,  # CAD L32 -> PROD L22
    'DESTILADOS':             -10,  # CAD L39 -> PROD L29
    'COMBOS':                 -10,  # CAD L52 -> PROD L42
    'DRINK':                  -10,  # CAD L68 -> PROD L58
    'DOSES & OUTROS':         -10,  # CAD L79 -> PROD L69
}

# ---------------------------------------------------------------------------
# Parser: Produtos Vendidos XLSX (Arquivo 1)
# Cabeçalho dinâmico — busca linha com 'Produto' na col A
# Usa col 'Quantidade' (após devoluções) e col 'Preço'
# ---------------------------------------------------------------------------

def parse_produtos_xlsx(file_bytes):
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

        qtd_idx   = col_map.get('Quantidade', 7)
        preco_idx = col_map.get('Preço', 8)

        try:
            qtd = int(float(str(row[qtd_idx] or 0).replace(',','.')))
        except Exception:
            qtd = 0

        try:
            v = row[preco_idx]
            if isinstance(v, (int, float)):
                preco = round(float(v), 2)
            else:
                # String brasileira: "R$ 1.234,56" → remover R$, remover milhar, trocar vírgula
                s = str(v or 0).replace('R$','').replace(' ','').strip()
                # Detectar formato: se tem vírgula é BR, se só ponto é EN
                if ',' in s:
                    s = s.replace('.','').replace(',','.')
                preco = round(float(s), 2)
        except Exception:
            preco = 0.0

        if qtd <= 0:
            continue

        nome = str(row[0]).strip()
        # Corrigir categoria de produtos que o YUZER classifica diferente da planilha
        if nome in OVERRIDE_CAT:
            cat = OVERRIDE_CAT[nome]

        produtos.append({
            'produto':      nome,
            'subcategoria': subcat,
            'cat':          cat,
            'qtd_vendida':  qtd,
            'preco':        preco,
        })

    return produtos

# ---------------------------------------------------------------------------
# Parser: Bônus/Cortesia PDF (Arquivo 4)
# ---------------------------------------------------------------------------

def _preco_str(s):
    if isinstance(s, (int, float)):
        return round(float(s), 2)
    s = str(s or '0').replace('R$','').replace('\xa0','').strip()
    if ',' in s:
        s = s.replace('.','').replace(',','.')
    try: return round(float(s), 2)
    except: return 0.0

def _normalizar_subcat(s):
    s = str(s).strip().upper().replace('\n', ' ')
    if 'NÃO' in s or 'NAO' in s:
        return 'BEBIDAS NÃO ALCOOLICAS'
    if s in ('BEBIDAS', 'BEBIDAS ALCOOLICAS'):
        return 'BEBIDAS ALCOOLICAS'
    return s

def parse_bonus_pdf(file_bytes):
    produtos = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            for table in (page.extract_tables() or []):
                for row in table:
                    if not row or not row[0]:
                        continue
                    cell0 = str(row[0]).strip()
                    if cell0 == 'NOME' or not cell0:
                        continue
                    if re.match(r'^[\d\s]+$', cell0):
                        continue

                    # TIPO A — linha normal
                    if row[1] is not None and row[5] is not None:
                        try:
                            qtd = int(str(row[5]).strip())
                            if qtd <= 0:
                                continue
                            subcat = _normalizar_subcat(row[3] or '')
                            cat = MAPA_SUBCAT.get(subcat, 'DOSES & OUTROS')
                            produtos.append({
                                'produto':     cell0,
                                'cat':         cat,
                                'qtd_vendida': qtd,
                                'preco':       _preco_str(row[8]),
                            })
                        except Exception:
                            pass

                    # TIPO B — colada com \n
                    elif row[1] is None and '\n' in cell0:
                        lines = cell0.split('\n')
                        subcat = _normalizar_subcat(lines[0])
                        cat = MAPA_SUBCAT.get(subcat, 'DOSES & OUTROS')
                        for part in lines[1:]:
                            m = re.match(
                                r'^(.+?)\s+FINAL\s+\S+\s+.+?\s+(\d+)\s+\d+\s+\d+\s+R\$\s*([\d.,]+)',
                                part.strip()
                            )
                            if m:
                                qtd = int(m.group(2))
                                if qtd > 0:
                                    produtos.append({
                                        'produto':     m.group(1).strip(),
                                        'cat':         cat,
                                        'qtd_vendida': qtd,
                                        'preco':       round(float(m.group(3).replace('.','').replace(',','.')), 2),
                                    })

                    # TIPO C — totalmente colada sem \n
                    elif row[1] is None and 'FINAL' in cell0:
                        m = re.match(
                            r'^(.+?)\s+FINAL\s+\S+\s+(\S+)\s+\S+\s+(\d+)\s+\d+\s+\d+\s+R\$\s*([\d.,]+)',
                            cell0
                        )
                        if m:
                            qtd = int(m.group(3))
                            if qtd > 0:
                                subcat = _normalizar_subcat(m.group(2))
                                cat = MAPA_SUBCAT.get(subcat, 'DOSES & OUTROS')
                                produtos.append({
                                    'produto':     m.group(1).strip(),
                                    'cat':         cat,
                                    'qtd_vendida': qtd,
                                    'preco':       round(float(m.group(4).replace('.','').replace(',','.')), 2),
                                })
    return produtos

# ---------------------------------------------------------------------------
# Parser: Exportação Caixas XLSX (Arquivo 2)
# Cabeçalho dinâmico — busca linha com 'Id' na col A
# Colunas reais: [0]=Id [1]=Usuário [3]=Serial [5]=Operação [6]=Total
#                [13]=Crédito [14]=Débito [15]=Pix [16]=Dinheiro
# NÃO usar col [10] "Dinheiro em caixa" (contagem física ≠ pagamento)
# ---------------------------------------------------------------------------

def parse_caixas(file_bytes):
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

        dinheiro_bruto  = gcol(['dinheiro'], 16)         # col Q = Dinheiro recebido
        devolvido       = gcol(['total produtos retornados', 'total retornado'], 12)  # col M

        # Dinheiro líquido = Dinheiro recebido - valor devolvido em dinheiro
        # Usar max(0, ...) para evitar negativo
        dinheiro_liq = round(max(0.0, dinheiro_bruto - devolvido), 2)

        caixas.append({
            'usuario':  scol(['usuario'], 1),
            'serial':   scol(['serial'], 3),
            'operacao': scol(['operacao'], 5),
            'total':    gcol(['total'], 6),
            'credito':  gcol(['credito'], 13),
            'debito':   gcol(['debito'], 14),
            'pix':      gcol(['pix'], 15),
            'dinheiro': dinheiro_liq,   # líquido após devoluções
        })

    return caixas

# ---------------------------------------------------------------------------
# Parser: Painel de Vendas XLSX (Arquivo 3)
# Lê Total geral e as 4 formas principais: PIX, DEBIT_CARD, CREDIT_CARD, CASH
# NÃO soma sub-bandeiras (Maestro, Visa, etc.) — são sub-totais de DEBIT/CREDIT
# ---------------------------------------------------------------------------

def parse_painel_vendas(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    painel = {}
    formas = {}
    lendo_formas = False
    passou_operacoes = False

    FORMAS_PRINCIPAIS = {'PIX', 'DEBIT_CARD', 'CREDIT_CARD', 'CASH', 'APP', 'CASHLESS'}

    for row in ws.iter_rows(values_only=True):
        if row[0] is None:
            continue
        key = str(row[0]).strip()
        val = row[1] if len(row) > 1 else None

        if key == 'Formas de Pagamento':
            lendo_formas = True
            continue

        # Parar ao encontrar sub-bandeiras ou Operações
        if key.startswith('Total por bandeira') or key in ('Operacoes', 'Operações'):
            lendo_formas = False
            if 'Opera' in key:
                passou_operacoes = True
            continue

        # Ler só formas principais — ignorar Maestro, Visa, Elo, etc.
        if lendo_formas and val is not None and key in FORMAS_PRINCIPAIS:
            try: formas[key] = round(float(val or 0), 2)
            except: pass
            continue

        # Métricas gerais (Total, Pedidos, Ticket médio)
        if not passou_operacoes and key in ('Total', 'Pedidos', 'Média', 'Media', 'Ticket'):
            if key not in painel:
                painel[key] = val

    painel['formas_pagamento'] = formas
    return painel

# ---------------------------------------------------------------------------
# Leitura do CADASTRO da planilha Google (nome + preço por categoria)
# ---------------------------------------------------------------------------

def ler_cadastro(service, spreadsheet_id):
    """
    Retorna: {cat: [{nome, preco, linha_cadastro}, ...]}
    Lê CADASTRO col B (nome) e F (preço) para cada categoria.
    """
    catalogo = {cat: [] for cat in CAT_INICIO}
    for cat, inicio in CAT_INICIO.items():
        fim = inicio + CAT_MAX[cat] - 1
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"CADASTRO!B{inicio}:F{fim}"
        ).execute()
        for i, row in enumerate(result.get('values', [])):
            nome = str(row[0]).strip() if len(row) > 0 and row[0] else ''
            if not nome:
                continue
            try:
                preco = round(float(
                    str(row[4] if len(row) > 4 else 0)
                    .replace('R$','').replace('.','').replace(',','.').strip()
                ), 2)
            except Exception:
                preco = 0.0
            catalogo[cat].append({
                'nome':           nome,
                'preco':          preco,
                'linha_cadastro': inicio + i,
            })
    return catalogo

# ---------------------------------------------------------------------------
# Leitura do mapa de linhas: ESTOQUE e PRODUCAO (col A -> linha)
# ---------------------------------------------------------------------------

def ler_mapa_linhas(service, spreadsheet_id):
    """
    Lê col A das abas ESTOQUE e PRODUÇÃO e retorna
    dicionários nome_produto -> linha, para uso no batchUpdate.
    """
    est_map = {}
    prod_map = {}

    # ESTOQUE col A (linhas 1-80)
    r = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range="ESTOQUE!A1:A80"
    ).execute()
    for i, row in enumerate(r.get('values', []), 1):
        if row and row[0] and str(row[0]).strip() not in IGNORAR_NOMES:
            est_map[str(row[0]).strip()] = i

    # PRODUÇÃO col A (linhas 1-80)
    r = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range="PRODUÇÃO!A1:A80"
    ).execute()
    for i, row in enumerate(r.get('values', []), 1):
        if row and row[0] and str(row[0]).strip() not in IGNORAR_NOMES:
            prod_map[str(row[0]).strip()] = i

    return est_map, prod_map

# ---------------------------------------------------------------------------
# Motor de Conciliação Multi-Atributo
# Usa: categoria + preço + nome (sequência + tokens) + unidade de medida
# Algoritmo: pré-match global greedy por score descendente
# ---------------------------------------------------------------------------

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
    'buul': 'bull', 'redbull': 'red bull', 'redbuul': 'red bull',
    'oldpar': 'old parr', 'oldparr': 'old parr',
    'tanqueray': 'tanqueray', 'tanquery': 'tanqueray',
    'moscow': 'moscow', 'mulle': 'mule',
    'limonade': 'lemonade', 'limao': 'limao', 'gengibre': 'gengibre',
    'tonica': 'tonica', 'tonicas': 'tonica',
    'smirnoff': 'smirnoff', 'heineken': 'heineken', 'amstel': 'amstel',
    'budweiser': 'budweiser', 'ciroc': 'ciroc', 'ketel': 'ketel',
}

def _norm_str(s):
    """Remove acentos, lowercase, aplica aliases."""
    s = unicodedata.normalize('NFKD', str(s))
    s = ''.join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r'[^a-z0-9 ]', ' ', s.lower())
    for a, b in ALIAS.items():
        s = re.sub(r'\b' + a + r'\b', b, s)
    return s

def _tokens(nome):
    """Tokens significativos sem stop words."""
    return set(t for t in _norm_str(nome).split() if t not in STOP_WORDS and len(t) > 1)

def _extrair_ml(nome):
    """Extrai volume em ml (ex: 350ml → 350, 1l → 1000)."""
    m = re.search(r'(\d+\.?\d*)\s*(ml|l|lt)\b', nome.lower().replace(' ',''))
    if m:
        val = float(m.group(1))
        return val * 1000 if m.group(2) in ('l','lt') else val
    return None

def _score_par(nome_p, preco_p, nome_y, preco_y):
    """
    Score 0-1 combinando:
    - Preço: 35% — exact match = 1.0, cai linearmente com diferença %
    - Nome:  45% — 60% sequência + 40% interseção de tokens
    - Unidade: 20% — match exato de ML, neutro se ausente
    """
    # Preço
    if preco_p > 0 and preco_y > 0:
        if abs(preco_p - preco_y) < 0.01:
            s_preco = 1.0
        else:
            diff_pct = abs(preco_p - preco_y) / max(preco_p, preco_y)
            s_preco = max(0.0, 1.0 - diff_pct * 3)
    else:
        s_preco = 0.0

    # Nome: sequência + tokens
    np_norm, ny_norm = _norm_str(nome_p), _norm_str(nome_y)
    s_seq = difflib.SequenceMatcher(None, np_norm, ny_norm).ratio()
    tp, ty = _tokens(nome_p), _tokens(nome_y)
    s_tok = len(tp & ty) / len(tp | ty) if (tp and ty) else 0.0
    s_nome = s_seq * 0.6 + s_tok * 0.4

    # Unidade de medida
    ml_p, ml_y = _extrair_ml(nome_p), _extrair_ml(nome_y)
    if ml_p is not None and ml_y is not None:
        s_un = 1.0 if abs(ml_p - ml_y) < 1 else 0.0
    else:
        s_un = 0.5  # sem info = neutro

    return s_preco * 0.35 + s_nome * 0.45 + s_un * 0.20

def _similaridade(a, b):
    """Compatibilidade com código legado."""
    return _score_par(a, 0, b, 0)

def _normalizar(nome):
    """Compatibilidade com código legado."""
    return _norm_str(nome)

# ---------------------------------------------------------------------------
# Conciliação por Pré-Match Global (greedy por score descendente)
# Garante que cada produto YUZER é usado no máximo uma vez
# ---------------------------------------------------------------------------

LIMIAR_SCORE = 0.40

def conciliar(catalogo, vendas, bonus):
    """
    Para cada categoria, constrói matriz de scores (planilha × YUZER),
    ordena por score descendente e faz atribuição greedy.
    Retorna: {cat: [{nome, linha_cadastro, preco,
                     qtd_venda, qtd_bonus, qtd_sistema,
                     match_venda, score_venda,
                     match_bonus, score_bonus, conciliado}]}
    """
    agrupado = {cat: [] for cat in CAT_INICIO}

    for cat, itens in catalogo.items():
        v_cat = [p for p in vendas if p['cat'] == cat]
        b_cat = [p for p in bonus  if p['cat'] == cat]

        def pre_match(planilha_items, yuzer_items):
            """Pré-match global: retorna dict {idx_plan: (idx_yuzer, nome_yuzer, score)}"""
            if not planilha_items or not yuzer_items:
                return {}

            # Construir todos os pares acima do limiar
            pares = []
            for pi, item in enumerate(planilha_items):
                for yi, prod in enumerate(yuzer_items):
                    s = _score_par(item['nome'], item['preco'],
                                   prod['produto'], prod['preco'])
                    if s >= LIMIAR_SCORE:
                        pares.append((s, pi, yi))

            # Greedy: atribuir melhor score primeiro
            pares.sort(reverse=True)
            plan_usado = set()
            yuzer_usado = set()
            resultado = {}
            for s, pi, yi in pares:
                if pi not in plan_usado and yi not in yuzer_usado:
                    resultado[pi] = (yi, yuzer_items[yi]['produto'], s,
                                     yuzer_items[yi]['qtd_vendida'])
                    plan_usado.add(pi)
                    yuzer_usado.add(yi)
            return resultado

        match_v = pre_match(itens, v_cat)
        match_b = pre_match(itens, b_cat)

        for pi, item in enumerate(itens):
            mv = match_v.get(pi)
            mb = match_b.get(pi)

            qtd_venda = mv[3] if mv else 0
            qtd_bonus = mb[3] if mb else 0

            agrupado[cat].append({
                'nome':           item['nome'],
                'linha_cadastro': item['linha_cadastro'],
                'preco':          item['preco'],
                'qtd_venda':      qtd_venda,
                'qtd_bonus':      qtd_bonus,
                'qtd_sistema':    qtd_venda + qtd_bonus,
                'match_venda':    mv[1] if mv else None,
                'score_venda':    mv[2] if mv else 0.0,
                'match_bonus':    mb[1] if mb else None,
                'score_bonus':    mb[2] if mb else 0.0,
                'conciliado':     mv is not None or mb is not None,
            })

    return agrupado


# ---------------------------------------------------------------------------
# Builders Google Sheets — usando mapa de linhas real das abas
# ---------------------------------------------------------------------------

def build_estoque_updates(agrupado, est_map=None):
    """
    ESTOQUE col I = qtd_venda APENAS (sem bônus — bônus vai só em PRODUÇÃO col C).
    Retorna (updates, nao_conciliados).
    """
    updates = []
    nao_conciliados = []
    for prods in agrupado.values():
        for p in prods:
            linha = p['linha_cadastro'] + OFFSET_ESTOQUE
            # Só vendas — bônus fica exclusivamente na PRODUÇÃO col C
            updates.append({'range': f"ESTOQUE!I{linha}", 'values': [[p['qtd_venda']]]})
            if not p.get('conciliado', True) and p['preco'] > 0:
                nao_conciliados.append(
                    f"{p['nome']} (R${p['preco']:.2f}) — sem match no YUZER"
                )
    return updates, nao_conciliados

def build_producao_updates(agrupado, prod_map=None):
    """
    PRODUÇÃO col C = qtd_bonus (só onde bonus > 0).
    Usa prod_map (nome→linha) lido diretamente da planilha Google.
    O prod_map é construído em ler_mapa_linhas() e passado pelo enviar().
    Retorna (updates, avisos).
    """
    updates = []
    avisos  = []

    for cat, prods in agrupado.items():
        for p in prods:
            if p['qtd_bonus'] <= 0:
                continue

            nome  = p['nome']
            linha = None

            # 1. Usar prod_map se disponível (nome → linha real na planilha)
            if prod_map and nome in prod_map:
                linha = prod_map[nome]

            # 2. Fallback: calcular pelo offset (menos preciso)
            if linha is None:
                offset = OFFSET_PRODUCAO.get(cat, -10)
                linha  = p['linha_cadastro'] + offset

            if linha and linha > 0:
                updates.append({'range': f"PRODUÇÃO!C{linha}", 'values': [[p['qtd_bonus']]]})

            score = p.get('score_bonus', 1.0)
            if score < 0.5 and p.get('match_bonus'):
                avisos.append(
                    f"Bônus incerto: '{nome}' → '{p['match_bonus']}' ({int(score*100)}%) — verifique"
                )

    return updates, avisos

# ---------------------------------------------------------------------------
# Rotas
# ---------------------------------------------------------------------------

@app.route('/api/preview', methods=['POST'])
def preview():
    try:
        result = {}

        vendas = []
        bonus  = []

        if 'produtos_vendidos' in request.files:
            vendas = parse_produtos_xlsx(request.files['produtos_vendidos'].read())
            result['produtos'] = vendas

        if 'produtos_bonus' in request.files:
            b_bytes = request.files['produtos_bonus'].read()
            fname   = request.files['produtos_bonus'].filename or ''
            bonus   = (parse_bonus_pdf(b_bytes)
                       if fname.lower().endswith('.pdf')
                       else parse_produtos_xlsx(b_bytes))
            result['bonus'] = bonus

        if 'exportacao_caixas' in request.files:
            result['caixas'] = parse_caixas(request.files['exportacao_caixas'].read())

        if 'painel_de_vendas' in request.files:
            painel = parse_painel_vendas(request.files['painel_de_vendas'].read())
            fp = painel.get('formas_pagamento', {})
            result['painel'] = painel
            result['resumo'] = {
                'total_faturado': painel.get('Total', 0),
                'total_pedidos':  painel.get('Pedidos', 0),
                'ticket_medio':   painel.get('Média', painel.get('Media', 0)),
                'credito':        fp.get('CREDIT_CARD', 0),
                'debito':         fp.get('DEBIT_CARD', 0),
                'pix':            fp.get('PIX', 0),
                'dinheiro':       fp.get('CASH', 0),
            }

        # Conciliação: produtos do YUZER não encontrados na planilha
        # (só disponível se houver vendas e bônus — catálogo lido no /enviar)
        # No preview, indicamos quais preços existem nos arquivos YUZER
        precos_vendas = {}
        for p in vendas:
            precos_vendas.setdefault(p['cat'], set()).add(p['preco'])
        precos_bonus = {}
        for p in bonus:
            precos_bonus.setdefault(p['cat'], set()).add(p['preco'])

        result['conciliacao_info'] = {
            'total_vendas': len(vendas),
            'total_bonus':  len(bonus),
            'categorias_vendas': {k: len(v) for k, v in precos_vendas.items()},
        }

        return jsonify({'success': True, 'data': result})
    except Exception as e:
        import traceback
        return jsonify({'success': False, 'error': str(e),
                        'trace': traceback.format_exc()}), 400


# ---------------------------------------------------------------------------
# Mapeamento fixo YUZER → Planilha (salvo em memória do servidor)
# Estrutura: { "Nome YUZER": "Nome Planilha", ... }
# ---------------------------------------------------------------------------
_mapeamento_store = {
    # Drinks (todos R$35 — mapeamento fixo evita troca por posição)
    'DRINK Tropical Gin':    'TROPICAL GIN ( GIN + RODELA DE LARANJA E RED BUUL TROPICAL )',
    'DRINK Melancita':       'MELANCITA ( GIN + RODELA DE LIMÃO E RED BUUL MELANCIA )',
    'DRINK Moscow Mule':     'MOSCOW MULLE ( VODKA + XAROPE DE GENGIBRE + SUMO DE LIMÃO E ESPUMA CITRICA )',
    'DRINK Pink Limonade':   'PINK LEMONADE ( GIN +  SUCO DE  LIMÃO + GROSELHA E RODELA DE LIMÃO SICILIANO)',
    'DRINK Gija':            'GIJA ( GIN + TONICA + XAROPE DE GENGIBRE + CANELA E RODELA DE LIMÃO SICILIANO )',
    'DRINK Gin Tônica':      'GIN TONICA ( GIN + TONICA E RODELA DE LIMÃO )',
    'DRINK Vodka + Red Bull':'VODKA E RED BUUL (VODKA + RED BUUL + ESCOLHA SEU SABOR )',
    # Combos R$440
    'Old Parr+3 Red Bull':       'OLDPAR 12 ANOS 1L  + 3 REDBULL 250ML',
    'Old Parr+5 Águas de Coco':  'OLDPARR 12 ANOS 1L + 5 AGUA DE COCO',
}

@app.route('/api/mapeamento', methods=['GET'])
def get_mapeamento():
    return jsonify({'success': True, 'mapeamento': _mapeamento_store})

@app.route('/api/mapeamento', methods=['POST'])
def save_mapeamento():
    data = request.get_json(force=True) or {}
    mapa = data.get('mapeamento', {})
    _mapeamento_store.clear()
    _mapeamento_store.update(mapa)
    return jsonify({'success': True, 'total': len(_mapeamento_store)})

# ---------------------------------------------------------------------------
# Limpeza de células antes de enviar
# ---------------------------------------------------------------------------

def limpar_planilha(service, spreadsheet_id):
    """
    Zera ESTOQUE col I (linhas 6-76), PRODUCAO col C (linhas 5-70)
    e FECHAMENTO CAIXAS B3:H52 antes de escrever dados novos.
    Limpa cada aba individualmente para evitar erro com acentos na API.
    """
    ranges = [
        'RESUMO!B3:B9',               # Receita Bar
        'ESTOQUE!I6:I76',             # Consumo Sistema
        'PRODUÇÃO!C5:C77',            # Espelho Bonus
        'FECHAMENTO CAIXAS!B3:H32',   # Garçom PIX
        'FECHAMENTO CAIXAS!B36:H50',  # Caixas Fixos
        'FECHAMENTO CAIXAS!B54:H83',  # Garçons crachá
        # RELATORIO DE VENDA — NÃO limpar: tem fórmulas que puxam de ESTOQUE/PRODUÇÃO
    ]
    service.spreadsheets().values().batchClear(
        spreadsheetId=spreadsheet_id,
        body={'ranges': ranges}
    ).execute()

    # Limpar PRODUÇÃO separadamente (aba com acento)
    try:
        service.spreadsheets().values().clear(
            spreadsheetId=spreadsheet_id,
            range='PRODUÇÃO!C5:C70'
        ).execute()
    except Exception:
        # Tentar sem acento como fallback
        try:
            service.spreadsheets().values().clear(
                spreadsheetId=spreadsheet_id,
                range='PRODUCAO!C5:C70'
            ).execute()
        except Exception:
            pass  # Se falhar, segue sem limpar — não bloqueia o envio

# ---------------------------------------------------------------------------
# Validação de totais
# ---------------------------------------------------------------------------

def validar_totais(agrupado, painel):
    """
    Compara soma dos qtd_sistema × preço com o total do painel de vendas.
    Retorna lista de avisos.
    """
    avisos = []
    total_painel = float(painel.get('Total', 0) or 0)
    if total_painel == 0:
        return avisos

    total_calculado = sum(
        p['qtd_venda'] * p['preco']
        for prods in agrupado.values()
        for p in prods
    )

    if total_calculado == 0:
        return avisos

    diferenca = abs(total_painel - total_calculado)
    pct = (diferenca / total_painel) * 100 if total_painel else 0

    if diferenca > 1.0:  # tolerância de R$1,00 para arredondamentos
        avisos.append(
            f'Divergência de totais: Painel = R${total_painel:,.2f} | '
            f'Calculado pelo sistema = R${total_calculado:,.2f} | '
            f'Diferença = R${diferenca:,.2f} ({pct:.1f}%)'
        )
    return avisos


@app.route('/api/enviar', methods=['POST'])
def enviar():
    try:
        spreadsheet_id = request.form.get('spreadsheet_id', '').strip()
        if not spreadsheet_id:
            return jsonify({'success': False, 'error': 'ID da planilha nao informado.'}), 400
        if 'docs.google.com' in spreadsheet_id:
            m = re.search(r'/d/([a-zA-Z0-9-_]+)', spreadsheet_id)
            if m:
                spreadsheet_id = m.group(1)

        service = get_sheets_service()
        batch      = []
        msgs       = []
        avisos     = []
        painel_data = {}
        agrupado   = {}   # inicializar para evitar NameError na validação de totais

        # ---- LIMPEZA PRÉVIA ----
        try:
            limpar_planilha(service, spreadsheet_id)
            msgs.append('Planilha limpa (ESTOQUE col I, PRODUÇÃO col C, CAIXAS)')
        except Exception as e_limpa:
            avisos.append(f'Limpeza prévia falhou (dados anteriores podem permanecer): {str(e_limpa)[:80]}')

        # ---- Produtos vendidos + bônus ----
        if 'produtos_vendidos' in request.files:
            vendas = parse_produtos_xlsx(request.files['produtos_vendidos'].read())

            bonus = []
            if 'produtos_bonus' in request.files:
                b_bytes = request.files['produtos_bonus'].read()
                fname   = request.files['produtos_bonus'].filename or ''
                bonus   = (parse_bonus_pdf(b_bytes)
                           if fname.lower().endswith('.pdf')
                           else parse_produtos_xlsx(b_bytes))

            # Aplicar mapeamento fixo aos nomes do YUZER
            if _mapeamento_store:
                for p in vendas + bonus:
                    if p['produto'] in _mapeamento_store:
                        p['produto_original'] = p['produto']
                        p['produto'] = _mapeamento_store[p['produto']]

            # Ler catálogo do CADASTRO
            catalogo = ler_cadastro(service, spreadsheet_id)
            total_cat = sum(len(v) for v in catalogo.values())
            msgs.append(f'CADASTRO: {total_cat} produtos lidos')

            # Ler mapa de linhas reais de PRODUÇÃO (nome → linha)
            _, prod_map_real = ler_mapa_linhas(service, spreadsheet_id)
            msgs.append(f'PRODUÇÃO: {len(prod_map_real)} produtos mapeados')

            # Conciliar por preço + similaridade de nome
            agrupado = conciliar(catalogo, vendas, bonus)

            # Estatísticas de conciliação
            total_prod  = sum(len(v) for v in agrupado.values())
            conciliados = sum(1 for prods in agrupado.values() for p in prods if p.get('conciliado'))
            nao_conc    = sum(1 for prods in agrupado.values() for p in prods if not p.get('conciliado') and p['preco'] > 0)
            msgs.append(f'Conciliação: {conciliados}/{total_prod} produtos encontrados no YUZER')
            if nao_conc:
                avisos.append(f'{nao_conc} produto(s) da planilha sem correspondência no YUZER')

            # Avisos de score baixo
            matches_baixos = [
                f"'{p['nome']}' → '{p['match_venda']}' ({int(p['score_venda']*100)}%)"
                for prods in agrupado.values() for p in prods
                if p.get('match_venda') and 0 < p.get('score_venda', 1) < 0.6
            ]
            if matches_baixos:
                avisos.append(f'Matches incertos — confira: {"; ".join(matches_baixos[:5])}')

            # ESTOQUE col I
            est_updates, est_nf = build_estoque_updates(agrupado)
            batch.extend(est_updates)
            msgs.append(f'ESTOQUE col I: {len(est_updates)} produtos preenchidos (vendas + bônus)')
            for a in est_nf: avisos.append(a)

            # PRODUÇÃO col C
            prod_updates, prod_avisos = build_producao_updates(agrupado, prod_map_real)
            batch.extend(prod_updates)
            msgs.append(f'PRODUÇÃO col C: {len(prod_updates)} produtos com bônus/cortesia preenchidos')
            for a in prod_avisos: avisos.append(a)

        # ---- Caixas — estrutura FECHAMENTO CAIXAS (confirmada na planilha manual) ----
        # L3:L32  → "N° DA MAQUINA" = GARÇOM PIX (garçons com máquina)
        # L36:L50 → "CAIXAS FIXOS"  = Caixa PIX  (caixas fixos)
        # L54:L83 → "N° CRACHA"     = garçons sem máquina (vazio neste fluxo)
        # Colunas: B=Nome C=Serial/Crachá D=Total E=Dinheiro F=PIX G=Débito H=Crédito
        if 'exportacao_caixas' in request.files:
            caixas = parse_caixas(request.files['exportacao_caixas'].read())

            def op_norm(s):
                return str(s).upper().replace('Ç','C').replace('Ã','A').strip()

            caixas_pix  = [c for c in caixas if 'CAIXA' in op_norm(c['operacao'])]
            garcons_pix = [c for c in caixas if 'GARCOM' in op_norm(c['operacao']) or 'GARÇOM' in op_norm(c['operacao'])]

            def to_rows(lista):
                return [[c['usuario'], c['serial'], c['total'],
                         c['dinheiro'], c['pix'], c['debito'], c['credito']]
                        for c in lista]

            # Seção topo (L3:L32) → GARÇOM PIX
            if garcons_pix:
                rows = to_rows(garcons_pix)
                batch.append({'range': f"FECHAMENTO CAIXAS!B3:H{2+len(rows)}", 'values': rows})
                msgs.append(f'FECHAMENTO CAIXAS — Garçons: {len(rows)} operadores')

            # Seção meio (L36:L50) → Caixa PIX
            if caixas_pix:
                rows = to_rows(caixas_pix)
                batch.append({'range': f"FECHAMENTO CAIXAS!B36:H{35+len(rows)}", 'values': rows})
                msgs.append(f'FECHAMENTO CAIXAS — Caixas: {len(rows)} operadores')

            if not caixas:
                msgs.append('FECHAMENTO CAIXAS: nenhum operador encontrado')

        # ---- Painel → RESUMO col B (B3:B9) ----
        # Col D é preenchida via fórmula da aba FECHAMENTO CAIXAS — não preencher aqui
        if 'painel_de_vendas' in request.files:
            painel_data = parse_painel_vendas(request.files['painel_de_vendas'].read())
            fp = painel_data.get('formas_pagamento', {})
            batch.append({
                'range': 'RESUMO!B3:B9',
                'values': [
                    [0],                        # B3 APP
                    [fp.get('CASH', 0)],        # B4 Dinheiro
                    [fp.get('CREDIT_CARD', 0)], # B5 Crédito
                    [fp.get('DEBIT_CARD', 0)],  # B6 Débito
                    [fp.get('PIX', 0)],         # B7 PIX
                    [0],                        # B8 Cancelamento
                    [fp.get('CASH', 0) +        # B9 Receita Total
                     fp.get('CREDIT_CARD', 0) +
                     fp.get('DEBIT_CARD', 0) +
                     fp.get('PIX', 0)],
                ],
            })
            msgs.append('RESUMO col B (B3:B9): formas de pagamento preenchidas')

            # ---- VALIDAÇÃO DE TOTAIS ----
            if 'produtos_vendidos' in request.files and agrupado:
                for a in validar_totais(agrupado, painel_data):
                    avisos.append(a)

        # ---- Enviar tudo de uma vez ----
        if batch:
            service.spreadsheets().values().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={'valueInputOption': 'USER_ENTERED', 'data': batch}
            ).execute()

        return jsonify({
            'success':  True,
            'message':  'Dados enviados com sucesso!',
            'detalhes': msgs,
            'avisos':   avisos,
        })

    except Exception as e:
        import traceback
        return jsonify({'success': False, 'error': str(e),
                        'trace': traceback.format_exc()}), 400


@app.route('/api/health')
def health():
    return jsonify({'status': 'ok', 'app': 'Prime Bar YUZER v4'})


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
