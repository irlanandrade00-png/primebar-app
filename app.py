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
# Mapeamento YUZER -> Planilha
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

OFFSET = -10

# ---------------------------------------------------------------------------
# Parsers
# ---------------------------------------------------------------------------

def parse_produtos_xlsx(file_bytes):
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
                'produto':      str(row[0]).strip(),
                'subcategoria': str(row[3]).strip() if row[3] else '',
                'qtd_vendida':  int(row[5] or 0),
                'preco':        float(row[8] or 0),
                'total_vendido': float(row[10] or 0),
            })
    return produtos


def parse_bonus_pdf(file_bytes):
    produtos = []

    def normalizar_subcat(s):
        s = s.strip().upper()
        if 'NÃO' in s or 'NAO' in s:
            return 'BEBIDAS NÃO ALCOOLICAS'
        if s in ('BEBIDAS', 'BEBIDAS ALCOOLICAS'):
            return 'BEBIDAS ALCOOLICAS'
        return s

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if not row or not row[0]:
                        continue
                    nome_cell = str(row[0]).strip()
                    if nome_cell == 'NOME':
                        continue

                    # Linha normal
                    if row[1] is not None and row[5] is not None:
                        try:
                            qtd = int(str(row[5]).strip())
                            if qtd <= 0:
                                continue
                            subcat_raw = (row[3] or '').replace('\n', ' ')
                            preco_str = str(row[8] or '0').replace('R$','').replace('\xa0','').replace('.','').replace(',','.').strip()
                            preco = float(preco_str) if preco_str else 0
                            produtos.append({
                                'produto':      nome_cell,
                                'subcategoria': normalizar_subcat(subcat_raw),
                                'qtd_vendida':  qtd,
                                'preco':        preco,
                            })
                        except Exception:
                            pass

                    # Linha colada com \n
                    elif row[1] is None and '\n' in nome_cell:
                        lines = nome_cell.split('\n')
                        subcat_linha = normalizar_subcat(lines[0])
                        for part in lines[1:]:
                            part = part.strip()
                            m = re.match(
                                r'^(.+?)\s+FINAL\s+\S+\s+.+?\s+(\d+)\s+\d+\s+\d+\s+R\$\s*([\d.,]+)',
                                part
                            )
                            if m:
                                qtd = int(m.group(2))
                                if qtd <= 0:
                                    continue
                                preco_str = m.group(3).replace('.','').replace(',','.')
                                produtos.append({
                                    'produto':      m.group(1).strip(),
                                    'subcategoria': subcat_linha,
                                    'qtd_vendida':  qtd,
                                    'preco':        float(preco_str),
                                })

                    # Linha totalmente colada sem \n
                    elif row[1] is None and 'FINAL' in nome_cell:
                        m = re.match(
                            r'^(.+?)\s+FINAL\s+\S+\s+(\S+)\s+\S+\s+(\d+)\s+\d+\s+\d+\s+R\$\s*([\d.,]+)',
                            nome_cell
                        )
                        if m:
                            qtd = int(m.group(3))
                            if qtd <= 0:
                                continue
                            preco_str = m.group(4).replace('.','').replace(',','.')
                            produtos.append({
                                'produto':      m.group(1).strip(),
                                'subcategoria': normalizar_subcat(m.group(2)),
                                'qtd_vendida':  qtd,
                                'preco':        float(preco_str),
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
# Leitura do CADASTRO da planilha (fonte de verdade dos nomes)
# ---------------------------------------------------------------------------

def ler_cadastro_planilha(service, spreadsheet_id):
    catalogo = {}
    for cat, inicio in CAT_INICIO.items():
        fim = inicio + CAT_MAX[cat] - 1
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"CADASTRO!B{inicio}:B{fim}"
        ).execute()
        valores = result.get('values', [])
        for i, row in enumerate(valores):
            if row and row[0] and str(row[0]).strip():
                nome = str(row[0]).strip()
                catalogo[nome.upper()] = {
                    'nome_original': nome,
                    'cat': cat,
                    'linha_cadastro': inicio + i,
                }
    return catalogo

# ---------------------------------------------------------------------------
# Merge usando catálogo da planilha como referência
# ---------------------------------------------------------------------------

def merge_com_catalogo(vendas, bonus, catalogo):
    vendas_map = {(p['produto'] or '').strip().upper(): p for p in vendas}
    bonus_map  = {(p['produto'] or '').strip().upper(): p for p in bonus}

    agrupado = {cat: [] for cat in CAT_INICIO}
    por_cat = {cat: [] for cat in CAT_INICIO}

    for nome_norm, info in catalogo.items():
        por_cat[info['cat']].append((info['linha_cadastro'], nome_norm, info['nome_original']))

    for cat in CAT_INICIO:
        itens = sorted(por_cat[cat], key=lambda x: x[0])
        for linha, nome_norm, nome_orig in itens:
            v = vendas_map.get(nome_norm)
            b = bonus_map.get(nome_norm)
            qtd_venda = v['qtd_vendida'] if v else 0
            qtd_bon   = b['qtd_vendida'] if b else 0
            agrupado[cat].append({
                'produto':          nome_orig,
                'linha_cadastro':   linha,
                'qtd_vendida_pura': qtd_venda,
                'qtd_bonus':        qtd_bon,
                'qtd_sistema':      qtd_venda + qtd_bon,
                'preco':            v['preco'] if v else (b['preco'] if b else 0),
            })

    return agrupado

# ---------------------------------------------------------------------------
# Builders de update
# ---------------------------------------------------------------------------

def build_estoque_updates(agrupado):
    updates = []
    for cat, prods in agrupado.items():
        for p in prods:
            linha_est = p['linha_cadastro'] + OFFSET
            updates.append({
                'range':  f"ESTOQUE!I{linha_est}",
                'values': [[p['qtd_sistema']]],
            })
    return updates


def build_relatorio_updates(agrupado):
    updates = []
    for cat, prods in agrupado.items():
        for p in prods:
            linha_rel = p['linha_cadastro'] + OFFSET
            updates.append({
                'range':  f"RELATORIO DE VENDA!B{linha_rel}",
                'values': [[p['qtd_vendida_pura']]],
            })
    return updates


def build_producao_updates(agrupado):
    updates = []
    for cat, prods in agrupado.items():
        for p in prods:
            if p['qtd_bonus'] > 0:
                linha_prod = p['linha_cadastro'] + OFFSET
                updates.append({
                    'range':  f"PRODUCAO!C{linha_prod}",
                    'values': [[p['qtd_bonus']]],
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
            result['produtos'] = parse_produtos_xlsx(request.files['produtos_vendidos'].read())

        if 'produtos_bonus' in request.files:
            bonus_bytes = request.files['produtos_bonus'].read()
            fname = request.files['produtos_bonus'].filename or ''
            if fname.lower().endswith('.pdf'):
                result['bonus'] = parse_bonus_pdf(bonus_bytes)
            else:
                result['bonus'] = parse_produtos_xlsx(bonus_bytes)

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
        import traceback
        return jsonify({'success': False, 'error': str(e), 'trace': traceback.format_exc()}), 400


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
            vendas = parse_produtos_xlsx(request.files['produtos_vendidos'].read())

            bonus = []
            if 'produtos_bonus' in request.files:
                bonus_bytes = request.files['produtos_bonus'].read()
                fname = request.files['produtos_bonus'].filename or ''
                if fname.lower().endswith('.pdf'):
                    bonus = parse_bonus_pdf(bonus_bytes)
                else:
                    bonus = parse_produtos_xlsx(bonus_bytes)

            # Lê nomes da planilha como referência
            catalogo = ler_cadastro_planilha(service, spreadsheet_id)
            descriptions.append(f'CADASTRO: {len(catalogo)} produtos lidos da planilha como referencia')

            agrupado = merge_com_catalogo(vendas, bonus, catalogo)

            for u in build_estoque_updates(agrupado):
                batch_data.append({'range': u['range'], 'values': u['values']})
            descriptions.append('ESTOQUE: col I (Consumo Sistema = vendas + bonus) preenchida')

            for u in build_relatorio_updates(agrupado):
                batch_data.append({'range': u['range'], 'values': u['values']})
            descriptions.append('RELATORIO DE VENDA: col B preenchida com vendas por produto')

            for u in build_producao_updates(agrupado):
                batch_data.append({'range': u['range'], 'values': u['values']})
            descriptions.append('PRODUCAO: col C (Cartao 1) preenchida com cortesia/bonus')

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

        if batch_data:
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
        import traceback
        return jsonify({'success': False, 'error': str(e), 'trace': traceback.format_exc()}), 400


@app.route('/api/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'app': 'Prime Bar - YUZER Integration v2'})


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
