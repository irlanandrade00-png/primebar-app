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

# ---------------------------------------------------------------------------
# Parser: Produtos Vendidos XLSX (Arquivo 1)
# ---------------------------------------------------------------------------

def parse_produtos_xlsx(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    produtos = []
    header_found = False
    for row in ws.iter_rows(values_only=True):
        if row[0] == 'Produto':
            header_found = True
            continue
        if header_found and row[0] is not None:
            subcat = str(row[3]).strip().upper() if row[3] else ''
            cat = MAPA_SUBCAT.get(subcat, 'DOSES & OUTROS')
            produtos.append({
                'produto':     str(row[0]).strip(),
                'subcategoria': subcat,
                'cat':         cat,
                'qtd_vendida': int(row[5] or 0),
                'preco':       round(float(row[8] or 0), 2),
            })
    return produtos

# ---------------------------------------------------------------------------
# Parser: Bônus/Cortesia PDF (Arquivo 4)
# ---------------------------------------------------------------------------

def _preco_str(s):
    return round(float(str(s or '0').replace('R$','').replace('\xa0','')
                       .replace(' ','').replace('.','').replace(',','.')), 2)

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
# ---------------------------------------------------------------------------

def parse_caixas(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    caixas = []
    header_found = False
    for row in ws.iter_rows(values_only=True):
        if row[0] == 'Id':
            header_found = True
            continue
        if header_found and row[0] is not None:
            caixas.append({
                'usuario':  str(row[1] or ''),
                'serial':   str(row[3] or ''),
                'total':    round(float(row[6]  or 0), 2),
                'credito':  round(float(row[12] or 0), 2),
                'debito':   round(float(row[13] or 0), 2),
                'pix':      round(float(row[14] or 0), 2),
                'dinheiro': round(float(row[15] or 0), 2),
            })
    return caixas

# ---------------------------------------------------------------------------
# Parser: Painel de Vendas XLSX (Arquivo 3)
# ---------------------------------------------------------------------------

def parse_painel_vendas(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    painel = {}
    formas = {}
    lendo_formas = False
    passou_operacoes = False

    for row in ws.iter_rows(values_only=True):
        if row[0] is None:
            continue
        key = str(row[0]).strip()
        val = row[1] if len(row) > 1 else None

        if key == 'Formas de Pagamento':
            lendo_formas = True
            continue
        if key in ('Operacoes', 'Operações'):
            lendo_formas = False
            passou_operacoes = True
            continue
        if lendo_formas and val is not None:
            formas[key] = val
            continue
        if not passou_operacoes and key in ('Total', 'Pedidos', 'Media', 'Média', 'Ticket'):
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
# Conciliação por PREÇO + POSIÇÃO
# ---------------------------------------------------------------------------

def conciliar(catalogo, vendas, bonus):
    """
    Retorna: {cat: [{nome, linha_cadastro, preco,
                     qtd_venda, qtd_bonus, qtd_sistema}, ...]}
    """
    agrupado = {cat: [] for cat in CAT_INICIO}

    for cat, itens in catalogo.items():
        v_cat = [p for p in vendas if p['cat'] == cat]
        b_cat = [p for p in bonus  if p['cat'] == cat]

        def por_preco(lista):
            d = {}
            for p in lista:
                d.setdefault(p['preco'], []).append(p)
            return d

        v_map = por_preco(v_cat)
        b_map = por_preco(b_cat)
        v_pos = {}
        b_pos = {}

        for item in itens:
            preco = item['preco']

            vi = v_pos.get(preco, 0)
            vl = v_map.get(preco, [])
            v = vl[vi] if vi < len(vl) else None
            if v:
                v_pos[preco] = vi + 1

            bi = b_pos.get(preco, 0)
            bl = b_map.get(preco, [])
            b = bl[bi] if bi < len(bl) else None
            if b:
                b_pos[preco] = bi + 1

            qtd_venda = v['qtd_vendida'] if v else 0
            qtd_bonus = b['qtd_vendida'] if b else 0

            agrupado[cat].append({
                'nome':           item['nome'],
                'linha_cadastro': item['linha_cadastro'],
                'preco':          preco,
                'qtd_venda':      qtd_venda,
                'qtd_bonus':      qtd_bonus,
                'qtd_sistema':    qtd_venda + qtd_bonus,
            })

    return agrupado

# ---------------------------------------------------------------------------
# Builders Google Sheets — usando mapa de linhas real das abas
# ---------------------------------------------------------------------------

def build_estoque_updates(agrupado, est_map):
    """ESTOQUE col I = qtd_sistema (vendas + bonus)"""
    updates = []
    nao_encontrados = []
    for prods in agrupado.values():
        for p in prods:
            linha = est_map.get(p['nome'])
            if linha:
                updates.append({
                    'range':  f"ESTOQUE!I{linha}",
                    'values': [[p['qtd_sistema']]]
                })
            else:
                nao_encontrados.append(p['nome'])
    return updates, nao_encontrados

def build_producao_updates(agrupado, prod_map):
    """PRODUÇÃO col C = qtd_bonus (só onde bonus > 0)"""
    updates = []
    nao_encontrados = []
    for prods in agrupado.values():
        for p in prods:
            if p['qtd_bonus'] > 0:
                linha = prod_map.get(p['nome'])
                if linha:
                    updates.append({
                        'range':  f"PRODUÇÃO!C{linha}",
                        'values': [[p['qtd_bonus']]]
                    })
                else:
                    nao_encontrados.append(p['nome'])
    return updates, nao_encontrados

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
            b_bytes = request.files['produtos_bonus'].read()
            fname = request.files['produtos_bonus'].filename or ''
            result['bonus'] = (parse_bonus_pdf(b_bytes)
                               if fname.lower().endswith('.pdf')
                               else parse_produtos_xlsx(b_bytes))

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

        return jsonify({'success': True, 'data': result})
    except Exception as e:
        import traceback
        return jsonify({'success': False, 'error': str(e),
                        'trace': traceback.format_exc()}), 400


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
        batch = []
        msgs  = []
        avisos = []

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

            # Ler catálogo do CADASTRO
            catalogo = ler_cadastro(service, spreadsheet_id)
            total_cat = sum(len(v) for v in catalogo.values())
            msgs.append(f'CADASTRO: {total_cat} produtos lidos')

            # Ler mapas de linhas reais do ESTOQUE e PRODUÇÃO
            est_map, prod_map = ler_mapa_linhas(service, spreadsheet_id)
            msgs.append(f'ESTOQUE: {len(est_map)} produtos mapeados')
            msgs.append(f'PRODUÇÃO: {len(prod_map)} produtos mapeados')

            # Conciliar por preço + posição
            agrupado = conciliar(catalogo, vendas, bonus)

            # ESTOQUE col I
            est_updates, est_nf = build_estoque_updates(agrupado, est_map)
            batch.extend(est_updates)
            msgs.append(f'ESTOQUE col I: {len(est_updates)} produtos preenchidos (vendas + bônus)')
            if est_nf:
                avisos.append(f'ESTOQUE não encontrados: {est_nf}')

            # PRODUÇÃO col C
            prod_updates, prod_nf = build_producao_updates(agrupado, prod_map)
            batch.extend(prod_updates)
            msgs.append(f'PRODUÇÃO col C: {len(prod_updates)} produtos com bônus/cortesia preenchidos')
            if prod_nf:
                avisos.append(f'PRODUÇÃO não encontrados: {prod_nf}')

        # ---- Caixas ----
        if 'exportacao_caixas' in request.files:
            caixas = parse_caixas(request.files['exportacao_caixas'].read())
            rows = [[c['usuario'], c['serial'], c['total'],
                     c['dinheiro'], c['pix'], c['debito'], c['credito']]
                    for c in caixas]
            batch.append({
                'range':  f"FECHAMENTO CAIXAS!B3:H{2 + len(rows)}",
                'values': rows,
            })
            msgs.append(f'FECHAMENTO CAIXAS: {len(rows)} operadores preenchidos')

        # ---- Painel → RESUMO ----
        if 'painel_de_vendas' in request.files:
            painel = parse_painel_vendas(request.files['painel_de_vendas'].read())
            fp = painel.get('formas_pagamento', {})
            batch.append({
                'range': 'RESUMO!B3:B7',
                'values': [
                    [0],
                    [fp.get('CASH', 0)],
                    [fp.get('CREDIT_CARD', 0)],
                    [fp.get('DEBIT_CARD', 0)],
                    [fp.get('PIX', 0)],
                ],
            })
            msgs.append('RESUMO: formas de pagamento preenchidas')

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
