from flask import Flask, request, jsonify
from flask_cors import CORS
import openpyxl
import json
import os
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import io

app = Flask(__name__)
CORS(app)

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def get_sheets_service():
    creds_json = os.environ.get('GOOGLE_CREDENTIALS')
    if not creds_json:
        raise Exception("Credenciais do Google não configuradas. Defina a variável GOOGLE_CREDENTIALS.")
    creds_info = json.loads(creds_json)
    creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
    return build('sheets', 'v4', credentials=creds)

def parse_produtos_vendidos(file_bytes):
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
                'produto': row[0],
                'tipo': row[1],
                'categoria': row[2],
                'subcategoria': row[3],
                'qtd_vendida': row[5] or 0,
                'qtd_devolvida': row[6] or 0,
                'quantidade': row[7] or 0,
                'preco': row[8] or 0,
                'total_vendido': row[10] or 0,
                'total_devolvido': row[11] or 0,
                'total': row[12] or 0,
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
                'id': row[0],
                'usuario': row[1],
                'cpf': row[2],
                'serial': row[3],
                'painel': row[4],
                'operacao': row[5],
                'total': row[6] or 0,
                'reimpressoes': row[7] or 0,
                'canceladas': row[8] or 0,
                'qtd_vendas': row[9] or 0,
                'dinheiro_caixa': row[10] or 0,
                'produtos_retornados': row[11] or 0,
                'credito': row[12] or 0,
                'debito': row[13] or 0,
                'pix': row[14] or 0,
                'dinheiro': row[15] or 0,
                'insercao_cashless': row[16] or 0,
                'subtotal': row[17] or 0,
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
        if key == 'Operações':
            reading_pagamentos = False
            continue
        if reading_pagamentos and val is not None:
            formas_pagamento[key] = val
        elif key in ['Total', 'Pedidos', 'Média', 'Devolução de pedidos']:
            painel[key] = val
    painel['formas_pagamento'] = formas_pagamento
    return painel

@app.route('/api/preview', methods=['POST'])
def preview():
    try:
        files = request.files
        result = {}

        if 'produtos_vendidos' in files:
            result['produtos'] = parse_produtos_vendidos(files['produtos_vendidos'].read())

        if 'exportacao_caixas' in files:
            result['caixas'] = parse_caixas(files['exportacao_caixas'].read())

        if 'painel_de_vendas' in files:
            result['painel'] = parse_painel_vendas(files['painel_de_vendas'].read())

        # Summary
        if 'painel' in result:
            fp = result['painel'].get('formas_pagamento', {})
            result['resumo'] = {
                'total_faturado': result['painel'].get('Total', 0),
                'total_pedidos': result['painel'].get('Pedidos', 0),
                'ticket_medio': result['painel'].get('Média', 0),
                'credito': fp.get('CREDIT_CARD', 0),
                'debito': fp.get('DEBIT_CARD', 0),
                'pix': fp.get('PIX', 0),
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
            return jsonify({'success': False, 'error': 'ID da planilha não informado.'}), 400

        # Extract spreadsheet ID from URL if full URL provided
        if 'docs.google.com' in spreadsheet_id:
            import re
            match = re.search(r'/d/([a-zA-Z0-9-_]+)', spreadsheet_id)
            if match:
                spreadsheet_id = match.group(1)

        service = get_sheets_service()
        updates = []

        # Parse produtos
        if 'produtos_vendidos' in request.files:
            produtos = parse_produtos_vendidos(request.files['produtos_vendidos'].read())
            # Write to RELATORIO DE VENDA - column B (Sistema) starting row 5
            # Map by product name matching column A
            prod_map = {p['produto'].strip().upper(): p for p in produtos}
            updates.append({
                'range': 'RELATORIO DE VENDA!B5',
                'values': [[p['qtd_vendida']] for p in produtos[:15]],
                'description': f'{len(produtos)} produtos mapeados'
            })

        # Parse caixas
        if 'exportacao_caixas' in request.files:
            caixas = parse_caixas(request.files['exportacao_caixas'].read())
            # Write to FECHAMENTO CAIXAS starting row 3
            # Columns: B=nome, C=serial, D=total, E=dinheiro, F=pix, G=debito, H=credito
            caixas_values = []
            for c in caixas:
                caixas_values.append([
                    c['usuario'],
                    c['serial'],
                    c['total'],
                    c['dinheiro'],
                    c['pix'],
                    c['debito'],
                    c['credito'],
                ])
            updates.append({
                'range': f"FECHAMENTO CAIXAS!B3:H{2+len(caixas_values)}",
                'values': caixas_values,
                'description': f'{len(caixas_values)} caixas/garçons mapeados'
            })

        # Parse painel - write totals to RESUMO
        if 'painel_de_vendas' in request.files:
            painel = parse_painel_vendas(request.files['painel_de_vendas'].read())
            fp = painel.get('formas_pagamento', {})
            # RESUMO: B3=APP/0, B4=Dinheiro, B5=Credito, B6=Debito, B7=PIX
            resumo_values = [
                [0],                              # B3 APP
                [fp.get('CASH', 0)],              # B4 Dinheiro
                [fp.get('CREDIT_CARD', 0)],       # B5 Credito
                [fp.get('DEBIT_CARD', 0)],        # B6 Debito
                [fp.get('PIX', 0)],               # B7 PIX
            ]
            updates.append({
                'range': 'RESUMO!B3:B7',
                'values': resumo_values,
                'description': 'Totais por forma de pagamento no RESUMO'
            })

        # Execute all updates
        batch_data = []
        for upd in updates:
            batch_data.append({
                'range': upd['range'],
                'values': upd['values']
            })

        body = {
            'valueInputOption': 'USER_ENTERED',
            'data': batch_data
        }
        service.spreadsheets().values().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=body
        ).execute()

        descriptions = [u['description'] for u in updates]
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
