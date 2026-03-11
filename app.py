from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import openpyxl
import json
import os
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import io
import re

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
# Parsers
# ---------------------------------------------------------------------------

def parse_produtos(file_bytes):
    """Parseia arquivo de produtos vendidos OU bonus - mesma estrutura YUZER."""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    produtos = []
    header_row = None
    for i, row in enumerate(ws.iter_rows(values_only=True), 1):
        if row[0] == 'Produto':
            header_row = i
            continue
        if header_row and row[0] is not None:
            produtos.append({
                'produto':       row[0],
                'subcategoria':  row[3],
                'qtd_vendida':   row[5] or 0,
                'preco':         row[8] or 0,
                'total_vendido': row[10] or 0,
            })
    return produtos

def parse_caixas(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    caixas = []
    header_row = None
    for i, row in enumerate(ws.iter_rows(values_only=True), 1):
        if row[0] == 'Id':
            header_row = i
            continue
        if header_row and row[0] is not None:
            caixas.append({
                'usuario':  row[1],
                'serial':   row[3],
                'total':    row[6] or 0,
                'credito':  row[12] or 0,
                'debito':   row[13] or 0,
                'pix':      row[14] or 0,
                'dinheiro': row[15] or 0,
            })
    return caixas

def parse_painel_vendas(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    painel = {}
    formas_pagamento = {}
    reading_pagamentos = False
    for row in ws.iter_rows(values_only=True):
        if row[0] is None:
            continue
        key = str(row[0]).strip()
        val = row[1] if len(row) > 1 else None
        if key == 'Formas de Pagamento':
            reading_pagamentos = True
            continue
        if key in ('Operacoes', 'Operações'):
            reading_pagamentos = False
            continue
        if reading_pagamentos and val is not None:
            formas_pagamento[key] = val
        elif key in ('Total', 'Pedidos', 'Media', 'Média'):
            painel[key] = val
    painel['formas_pagamento'] = formas_pagamento
    return painel

# ---------------------------------------------------------------------------
# Mapeamento YUZER -> Planilha
# ---------------------------------------------------------------------------

MAPA_SUBCAT = {
    'BEBIDAS NAO ALCOOLICAS': 'BEBIDAS NAO ALCOOLICAS',
    'BEBIDAS NÃO ALCOOLICAS': 'BEBIDAS NAO ALCOOLICAS',
    'BEBIDAS ALCOOLICAS':     'BEBIDAS ALCOOLICAS',
    'DESTILADOS':             'DESTILADOS',
    'DOSES':                  'DOSES & OUTROS',
    'DRINKS':                 'DRINK',
    'COMBOS':                 'COMBOS',
    'OUTROS':                 'DOSES & OUTROS',
}

# Linha inicial no CADASTRO por categoria
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

# CADASTRO linha X -> ESTOQUE / RELATORIO / PRODUCAO linha (X - 10)
OFFSET = -10


def organizar_por_categoria(produtos):
    agrupado = {cat: [] for cat in CAT_INICIO}
    for p in produtos:
        subcat = (p.get('subcategoria') or '').strip().upper()
        cat = MAPA_SUBCAT.get(subcat, 'DOSES & OUTROS')
        agrupado[cat].append(p)
    return agrupado


def merge_vendas_bonus(vendas, bonus):
    """
    Une produtos de vendas e bonus pelo nome.
      qtd_vendida_pura = so vendas  -> RELATORIO DE VENDA col B
      qtd_bonus        = so bonus   -> PRODUCAO col C (Cartao 1)
      qtd_sistema      = vendas + bonus -> ESTOQUE col I
    """
    bonus_map = {(p['produto'] or '').strip().upper(): p for p in bonus}
    merged = []
    nomes_vistos = set()

    for p in vendas:
        nome_norm = (p['produto'] or '').strip().upper()
        nomes_vistos.add(nome_norm)
        b = bonus_map.get(nome_norm)
        qtd_bon = b['qtd_vendida'] if b else 0
        merged.append({
            **p,
            'qtd_vendida_pura': p['qtd_vendida'],
            'qtd_bonus':        qtd_bon,
            'qtd_sistema':      p['qtd_vendida'] + qtd_bon,
        })

    # Produtos que existem apenas no bonus (nao estavam em vendas)
    for p in bonus:
        nome_norm = (p['produto'] or '').strip().upper()
        if nome_norm not in nomes_vistos:
            merged.append({
                **p,
                'qtd_vendida_pura': 0,
                'qtd_bonus':        p['qtd_vendida'],
                'qtd_sistema':      p['qtd_vendida'],
            })

    return merged

# ---------------------------------------------------------------------------
# Builders de update
# ---------------------------------------------------------------------------

def build_cadastro_updates(agrupado_vendas):
    """
    CADASTRO col B = nome do produto
    CADASTRO col F = preco de venda
    A planilha propaga automaticamente para ESTOQUE, RELATORIO, etc.
    """
    updates = []
    for cat, prods in agrupado_vendas.items():
        if not prods:
            continue
        inicio = CAT_INICIO[cat]
        prods = prods[:CAT_MAX[cat]]
        updates.append({
            'range': f"CADASTRO!B{inicio}:B{inicio + len(prods) - 1}",
            'values': [[p['produto']] for p in prods],
        })
        updates.append({
            'range': f"CADASTRO!F{inicio}:F{inicio + len(prods) - 1}",
            'values': [[p['preco']] for p in prods],
        })
    return updates


def build_estoque_updates(agrupado_merged):
    """
    ESTOQUE col I = qtd_sistema (vendas + bonus = Consumo Sistema total).
    Col K NAO e preenchida — a planilha ja puxa do CADASTRO via formula.
    """
    updates = []
    for cat, prods in agrupado_merged.items():
        if not prods:
            continue
        inicio_est = CAT_INICIO[cat] + OFFSET
        prods = prods[:CAT_MAX[cat]]
        updates.append({
            'range': f"ESTOQUE!I{inicio_est}:I{inicio_est + len(prods) - 1}",
            'values': [[p.get('qtd_sistema', p['qtd_vendida'])] for p in prods],
        })
    return updates


def build_relatorio_updates(agrupado_merged):
    """
    RELATORIO DE VENDA col B = apenas qtd_vendida_pura (SEM bonus).
    Somente Beb. Nao Alcoolicas precisa ser gravada diretamente;
    as demais categorias calculam automaticamente via formula
    =ESTOQUE!Ix - ESTOQUE!Gx ja existente na planilha.
    """
    updates = []
    cat = 'BEBIDAS NAO ALCOOLICAS'
    prods = agrupado_merged.get(cat, [])
    if not prods:
        return updates
    inicio_rel = CAT_INICIO[cat] + OFFSET
    prods = prods[:CAT_MAX[cat]]
    updates.append({
        'range': f"RELATORIO DE VENDA!B{inicio_rel}:B{inicio_rel + len(prods) - 1}",
        'values': [[p.get('qtd_vendida_pura', p['qtd_vendida'])] for p in prods],
    })
    return updates


def build_producao_cartao1_updates(agrupado_merged):
    """
    PRODUCAO col C (Cartao 1) = qtd_bonus por produto.
    Mesma logica de offset: CADASTRO linha X -> PRODUCAO linha (X - 10).
    Bonus e centralizado aqui para que ESTOQUE col G seja calculado
    automaticamente pela formula ='PRODUCAO'!O que ja existe na planilha.
    """
    updates = []
    for cat, prods in agrupado_merged.items():
        if not prods:
            continue
        inicio_prod = CAT_INICIO[cat] + OFFSET
        prods = prods[:CAT_MAX[cat]]
        # So grava se houver pelo menos um bonus > 0
        bonus_vals = [[p.get('qtd_bonus', 0)] for p in prods]
        if any(v[0] > 0 for v in bonus_vals):
            updates.append({
                'range': f"PRODUCAO!C{inicio_prod}:C{inicio_prod + len(prods) - 1}",
                'values': bonus_vals,
            })
    return updates

# ---------------------------------------------------------------------------
# Rotas
# ---------------------------------------------------------------------------

@app.route('/api/preview', methods=['POST'])
def preview():
    try:
        result = {}

        if 'produtos_vendidos' in request.files:
            result['produtos'] = parse_produtos(request.files['produtos_vendidos'].read())

        if 'produtos_bonus' in request.files:
            result['bonus'] = parse_produtos(request.files['produtos_bonus'].read())

        if 'exportacao_caixas' in request.files:
            result['caixas'] = parse_caixas(request.files['exportacao_caixas'].read())

        if 'painel_de_vendas' in request.files:
            painel = parse_painel_vendas(request.files['painel_de_vendas'].read())
            result['painel'] = painel
            fp = painel.get('formas_pagamento', {})
            result['resumo'] = {
                'total_faturado': painel.get('Total', 0),
                'total_pedidos':  painel.get('Pedidos', 0),
                'ticket_medio':   painel.get('Média', painel.get('Media', 0)),
                'credito':  fp.get('CREDIT_CARD', 0),
                'debito':   fp.get('DEBIT_CARD', 0),
                'pix':      fp.get('PIX', 0),
                'dinheiro': fp.get('CASH', 0),
            }

        return jsonify({'success': True, 'data': result})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 400


@app.route('/api/enviar', methods=['POST'])
def enviar():
    try:
        spreadsheet_id = request.form.get('spreadsheet_id', '').strip()
        if not spreadsheet_id:
            return jsonify({'success': False, 'error': 'ID da planilha nao informado.'}), 400

        if 'docs.google.com' in spreadsheet_id:
            match = re.search(r'/d/([a-zA-Z0-9-_]+)', spreadsheet_id)
            if match:
                spreadsheet_id = match.group(1)

        service = get_sheets_service()
        batch_data = []
        descriptions = []

        # --- Produtos vendidos + bonus ---
        if 'produtos_vendidos' in request.files:
            vendas = parse_produtos(request.files['produtos_vendidos'].read())
            bonus = []
            if 'produtos_bonus' in request.files:
                bonus = parse_produtos(request.files['produtos_bonus'].read())

            merged = merge_vendas_bonus(vendas, bonus)
            agrupado_vendas = organizar_por_categoria(vendas)
            agrupado_merged = organizar_por_categoria(merged)

            # 1. CADASTRO — nome (col B) e preco (col F)
            #    A planilha propaga automaticamente para ESTOQUE K, RELATORIO C, etc.
            for u in build_cadastro_updates(agrupado_vendas):
                batch_data.append({'range': u['range'], 'values': u['values']})
            total_prods = sum(len(v) for v in agrupado_vendas.values())
            descriptions.append(f'CADASTRO: {total_prods} produtos e precos preenchidos')

            # 2. ESTOQUE col I = vendas + bonus (Consumo Sistema total)
            for u in build_estoque_updates(agrupado_merged):
                batch_data.append({'range': u['range'], 'values': u['values']})
            descriptions.append('ESTOQUE: col I (Consumo Sistema = vendas + bonus)')

            # 3. RELATORIO DE VENDA col B = apenas vendas (sem bonus)
            for u in build_relatorio_updates(agrupado_merged):
                batch_data.append({'range': u['range'], 'values': u['values']})
            descriptions.append('RELATORIO DE VENDA: col B preenchida apenas com vendas')

            # 4. PRODUCAO col C (Cartao 1) = bonus por produto
            if bonus:
                for u in build_producao_cartao1_updates(agrupado_merged):
                    batch_data.append({'range': u['range'], 'values': u['values']})
                descriptions.append('PRODUCAO: col C (Cartao 1) preenchida com consumo bonus')

        # --- Caixas ---
        if 'exportacao_caixas' in request.files:
            caixas = parse_caixas(request.files['exportacao_caixas'].read())
            caixas_values = [[
                c['usuario'], c['serial'], c['total'],
                c['dinheiro'], c['pix'], c['debito'], c['credito'],
            ] for c in caixas]
            batch_data.append({
                'range': f"FECHAMENTO CAIXAS!B3:H{2 + len(caixas_values)}",
                'values': caixas_values,
            })
            descriptions.append(f'FECHAMENTO CAIXAS: {len(caixas_values)} operadores preenchidos')

        # --- Painel → RESUMO ---
        if 'painel_de_vendas' in request.files:
            painel = parse_painel_vendas(request.files['painel_de_vendas'].read())
            fp = painel.get('formas_pagamento', {})
            batch_data.append({
                'range': 'RESUMO!B3:B7',
                'values': [
                    [0],
                    [fp.get('CASH', 0)],
                    [fp.get('CREDIT_CARD', 0)],
                    [fp.get('DEBIT_CARD', 0)],
                    [fp.get('PIX', 0)],
                ],
            })
            descriptions.append('RESUMO: totais por forma de pagamento preenchidos')

        service.spreadsheets().values().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={'valueInputOption': 'USER_ENTERED', 'data': batch_data}
        ).execute()

        return jsonify({
            'success': True,
            'message': 'Dados enviados com sucesso para o Google Sheets!',
            'detalhes': descriptions
        })

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 400


@app.route('/api/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'app': 'Prime Bar - YUZER Integration'})


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
