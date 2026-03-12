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
# Constantes de mapeamento
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

# ---------------------------------------------------------------------------
# Parsers YUZER
# ---------------------------------------------------------------------------

def parse_produtos_xlsx(file_bytes):
    """Parseia arquivo de produtos vendidos XLSX exportado do YUZER."""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    produtos = []
    header_row = None
    for row in ws.iter_rows(values_only=True):
        if row[0] == 'Produto':
            header_row = True
            continue
        if header_row and row[0] is not None:
            subcat = str(row[3]).strip().upper() if row[3] else ''
            cat = MAPA_SUBCAT.get(subcat, 'DOSES & OUTROS')
            produtos.append({
                'produto':      str(row[0]).strip(),
                'subcategoria': subcat,
                'cat':          cat,
                'qtd_vendida':  int(row[5] or 0),
                'preco':        round(float(row[8] or 0), 2),
                'total_vendido': round(float(row[10] or 0), 2),
            })
    return produtos


def parse_bonus_pdf(file_bytes):
    """Parseia PDF de bonus/cortesia do YUZER."""
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

                    # Linha normal: colunas separadas corretamente
                    if row[1] is not None and row[5] is not None:
                        try:
                            qtd = int(str(row[5]).strip())
                            if qtd <= 0:
                                continue
                            subcat_raw = (row[3] or '').replace('\n', ' ')
                            subcat = normalizar_subcat(subcat_raw)
                            preco_str = str(row[8] or '0').replace('R$','').replace('\xa0','').replace('.','').replace(',','.').strip()
                            preco = round(float(preco_str), 2) if preco_str else 0
                            cat = MAPA_SUBCAT.get(subcat, MAPA_SUBCAT.get(subcat.upper(), 'DOSES & OUTROS'))
                            produtos.append({
                                'produto':      nome_cell,
                                'subcategoria': subcat,
                                'cat':          cat,
                                'qtd_vendida':  qtd,
                                'preco':        preco,
                            })
                        except Exception:
                            pass

                    # Linha colada com \n: "SUBCATEGORIA\nNOME FINAL ..."
                    elif row[1] is None and '\n' in nome_cell:
                        lines = nome_cell.split('\n')
                        subcat = normalizar_subcat(lines[0])
                        cat = MAPA_SUBCAT.get(subcat, 'DOSES & OUTROS')
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
                                    'subcategoria': subcat,
                                    'cat':          cat,
                                    'qtd_vendida':  qtd,
                                    'preco':        round(float(preco_str), 2),
                                })

                    # Linha totalmente colada sem \n mas com FINAL
                    elif row[1] is None and 'FINAL' in nome_cell:
                        m = re.match(
                            r'^(.+?)\s+FINAL\s+\S+\s+(\S+)\s+\S+\s+(\d+)\s+\d+\s+\d+\s+R\$\s*([\d.,]+)',
                            nome_cell
                        )
                        if m:
                            qtd = int(m.group(3))
                            if qtd <= 0:
                                continue
                            subcat = normalizar_subcat(m.group(2))
                            preco_str = m.group(4).replace('.','').replace(',','.')
                            cat = MAPA_SUBCAT.get(subcat, 'DOSES & OUTROS')
                            produtos.append({
                                'produto':      m.group(1).strip(),
                                'subcategoria': subcat,
                                'cat':          cat,
                                'qtd_vendida':  qtd,
                                'preco':        round(float(preco_str), 2),
                            })

    return produtos


def parse_caixas(file_bytes):
    """Parseia exportacao_caixas XLSX do YUZER."""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    caixas = []
    header_row = None
    for row in ws.iter_rows(values_only=True):
        if row[0] == 'Id':
            header_row = True
            continue
        if header_row and row[0] is not None:
            caixas.append({
                'usuario':  row[1],
                'serial':   row[3],
                'total':    round(float(row[6] or 0), 2),
                'credito':  round(float(row[12] or 0), 2),
                'debito':   round(float(row[13] or 0), 2),
                'pix':      round(float(row[14] or 0), 2),
                'dinheiro': round(float(row[15] or 0), 2),
            })
    return caixas


def parse_painel_vendas(file_bytes):
    """Parseia exportacao_painel_de_vendas XLSX do YUZER."""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    painel = {}
    formas_pagamento = {}
    reading_pagamentos = False
    passou_operacoes = False

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
            passou_operacoes = True
            continue

        # Formas de pagamento (entre 'Formas de Pagamento' e 'Operações')
        if reading_pagamentos and val is not None:
            formas_pagamento[key] = val
            continue

        # Totais gerais — só antes de 'Operações' para evitar subtotais por operador
        if not passou_operacoes and key in ('Total', 'Pedidos', 'Media', 'Média', 'Ticket'):
            if key not in painel:
                painel[key] = val

    painel['formas_pagamento'] = formas_pagamento
    return painel

# ---------------------------------------------------------------------------
# Leitura do CADASTRO da planilha (nome + preço por categoria)
# ---------------------------------------------------------------------------

def ler_cadastro_planilha(service, spreadsheet_id):
    """
    Lê nome (col B) e preço (col F) de cada produto cadastrado na planilha.
    Retorna lista por categoria ordenada por posição (linha_cadastro).
    Estrutura: {cat: [{nome, preco, linha_cadastro}, ...]}
    """
    catalogo = {cat: [] for cat in CAT_INICIO}

    for cat, inicio in CAT_INICIO.items():
        fim = inicio + CAT_MAX[cat] - 1
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"CADASTRO!B{inicio}:F{fim}"
        ).execute()
        valores = result.get('values', [])
        for i, row in enumerate(valores):
            nome = str(row[0]).strip() if len(row) > 0 and row[0] else ''
            # Preço está na col F = índice 4 do range B:F
            preco_raw = row[4] if len(row) > 4 else ''
            if not nome:
                continue
            try:
                preco = round(float(str(preco_raw).replace('R$','').replace('.','').replace(',','.').strip()), 2)
            except Exception:
                preco = 0.0
            catalogo[cat].append({
                'nome':           nome,
                'preco':          preco,
                'linha_cadastro': inicio + i,
            })

    return catalogo

# ---------------------------------------------------------------------------
# Conciliação por PREÇO + POSIÇÃO
# ---------------------------------------------------------------------------

def conciliar_por_preco(catalogo, vendas, bonus):
    """
    Para cada categoria:
      1. Pega os produtos do CADASTRO (ordenados por posição = linha_cadastro)
      2. Pega os produtos do YUZER filtrados pela categoria
      3. Agrupa os do YUZER por preço
      4. Para cada item do CADASTRO, busca no YUZER pelo preço
         — se houver mais de um com o mesmo preço, usa a posição relativa
           (1º produto do CADASTRO com preço X = 1º produto do YUZER com preço X)
      5. Retorna agrupado com qtd_vendida_pura, qtd_bonus, qtd_sistema
    """
    agrupado = {cat: [] for cat in CAT_INICIO}

    for cat, itens_cadastro in catalogo.items():
        # Filtrar vendas e bonus pela categoria
        vendas_cat = [p for p in vendas if p.get('cat') == cat]
        bonus_cat  = [p for p in bonus  if p.get('cat') == cat]

        # Agrupar YUZER por preço, mantendo ordem de aparição (posição)
        def agrupar_por_preco(lista):
            grupos = {}  # preco -> [produto, ...]
            for p in lista:
                preco = p['preco']
                if preco not in grupos:
                    grupos[preco] = []
                grupos[preco].append(p)
            return grupos

        vendas_por_preco = agrupar_por_preco(vendas_cat)
        bonus_por_preco  = agrupar_por_preco(bonus_cat)

        # Contador de posição usada por preço (para desempate)
        pos_vendas = {}
        pos_bonus  = {}

        for item in itens_cadastro:
            preco = item['preco']
            linha = item['linha_cadastro']

            # Buscar venda correspondente (preço + posição)
            idx_v = pos_vendas.get(preco, 0)
            lista_v = vendas_por_preco.get(preco, [])
            v = lista_v[idx_v] if idx_v < len(lista_v) else None
            if v:
                pos_vendas[preco] = idx_v + 1

            # Buscar bonus correspondente (preço + posição)
            idx_b = pos_bonus.get(preco, 0)
            lista_b = bonus_por_preco.get(preco, [])
            b = lista_b[idx_b] if idx_b < len(lista_b) else None
            if b:
                pos_bonus[preco] = idx_b + 1

            qtd_venda = v['qtd_vendida'] if v else 0
            qtd_bon   = b['qtd_vendida'] if b else 0

            agrupado[cat].append({
                'nome':             item['nome'],
                'linha_cadastro':   linha,
                'preco':            preco,
                'qtd_vendida_pura': qtd_venda,
                'qtd_bonus':        qtd_bon,
                'qtd_sistema':      qtd_venda + qtd_bon,
            })

    return agrupado

# ---------------------------------------------------------------------------
# Builders de update para Google Sheets
# ---------------------------------------------------------------------------

def build_estoque_updates(agrupado):
    """ESTOQUE col I = qtd_sistema (vendas + bonus = Consumo Sistema)."""
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
    """RELATORIO DE VENDA col B = apenas vendas (sem bonus)."""
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
    """PRODUCAO col C (Cartao 1) = qtd_bonus por produto."""
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

            # Lê CADASTRO da planilha (nome + preço + posição)
            catalogo = ler_cadastro_planilha(service, spreadsheet_id)
            total_cadastro = sum(len(v) for v in catalogo.values())
            descriptions.append(f'CADASTRO: {total_cadastro} produtos lidos da planilha')

            # Concilia por preço + posição
            agrupado = conciliar_por_preco(catalogo, vendas, bonus)

            # ESTOQUE col I = vendas + bonus
            for u in build_estoque_updates(agrupado):
                batch_data.append({'range': u['range'], 'values': u['values']})
            descriptions.append('ESTOQUE: col I (Consumo Sistema = vendas + bonus) preenchida')

            # RELATORIO DE VENDA col B = apenas vendas
            for u in build_relatorio_updates(agrupado):
                batch_data.append({'range': u['range'], 'values': u['values']})
            descriptions.append('RELATORIO DE VENDA: col B preenchida com vendas')

            # PRODUCAO col C = bonus
            prod_updates = build_producao_updates(agrupado)
            for u in prod_updates:
                batch_data.append({'range': u['range'], 'values': u['values']})
            if prod_updates:
                descriptions.append(f'PRODUCAO: col C preenchida com {len(prod_updates)} produtos com bonus/cortesia')
            else:
                descriptions.append('PRODUCAO: nenhum bonus/cortesia encontrado')

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
                    [0],                          # B3 = APP (sempre 0)
                    [fp.get('CASH', 0)],          # B4 = Dinheiro
                    [fp.get('CREDIT_CARD', 0)],   # B5 = Crédito
                    [fp.get('DEBIT_CARD', 0)],    # B6 = Débito
                    [fp.get('PIX', 0)],           # B7 = PIX
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
    return jsonify({'status': 'ok', 'app': 'Prime Bar - YUZER Integration v3'})


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
