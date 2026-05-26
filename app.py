from flask import Flask, jsonify, request, send_file
from flask_cors import CORS
import gspread
from gspread.utils import rowcol_to_a1
from google.oauth2.service_account import Credentials
from datetime import datetime, date
from functools import lru_cache
import os
import json
from decimal import Decimal, InvalidOperation
import re
import math
from io import BytesIO
from openpyxl import Workbook

app = Flask(__name__)

cors_origins = os.environ.get('CORS_ORIGINS')
if cors_origins:
    CORS(app, origins=[origin.strip() for origin in cors_origins.split(',') if origin.strip()])
else:
    CORS(app)

SPREADSHEET_ID = '1zUMUAfEdVYrV8ogIpARV2LnOXuNFhObWBsKb0ONQmh4'
API_KEY = os.environ.get('API_KEY', '').strip()
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]


@app.route('/', methods=['GET'])
def home():
    return send_file('app_render.html')


def api_response(ok=True, status='ok', http_status=200, **payload):
    base = {'ok': ok, 'status': status}
    base.update(payload)
    return jsonify(base), http_status


def api_error(message='Erro interno no servidor', http_status=500, exc=None, status='erro'):
    if exc is not None:
        app.logger.exception(message)
        if app.debug:
            message = str(exc)
    return api_response(False, status, http_status, erro=message, mensagem=message)


@app.before_request
def require_api_key():
    if request.method == 'OPTIONS' or not API_KEY:
        return None
    if request.endpoint in {'home', 'status', 'health'}:
        return None
    provided = request.headers.get('X-API-Key') or request.args.get('api_key')
    if provided != API_KEY:
        return api_error('Não autorizado', 401)
    return None


@lru_cache(maxsize=1)
def get_sheet():
    client = get_client()
    return client.open_by_key(SPREADSHEET_ID)


@lru_cache(maxsize=1)
def get_client():
    creds_json = os.environ.get('GOOGLE_CREDENTIALS')
    if creds_json:
        creds_dict = json.loads(creds_json)
    else:
        cred_path = 'credentials.json'
        if not os.path.exists(cred_path):
            for candidate in os.listdir('.'):
                if not candidate.endswith('.json'):
                    continue
                try:
                    with open(candidate, 'r', encoding='utf-8') as f:
                        maybe_creds = json.load(f)
                    if maybe_creds.get('type') == 'service_account':
                        cred_path = candidate
                        break
                except (OSError, json.JSONDecodeError):
                    continue
        with open(cred_path, 'r', encoding='utf-8') as f:
            creds_dict = json.load(f)
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)


def unique_headers(headers):
    seen = {}
    normalized = []
    for idx, header in enumerate(headers):
        name = norm_text(header) or f'COL_{idx + 1}'
        count = seen.get(name, 0) + 1
        seen[name] = count
        normalized.append(name if count == 1 else f'{name}_{count}')
    return normalized


def records_from_values(values):
    if not values:
        return []
    header_idx = 0
    required = {'data', 'fonte'}
    for idx, row in enumerate(values):
        labels = {norm_text(value).lower() for value in row}
        if required.issubset(labels) and any(label in labels for label in ('valor', 'valor (r$)')):
            header_idx = idx
            break

    headers = unique_headers(values[header_idx])
    records = []
    for row in values[header_idx + 1:]:
        if not any(norm_text(value) for value in row):
            continue
        padded = row + [''] * (len(headers) - len(row))
        records.append(dict(zip(headers, padded[:len(headers)])))
    return records


def get_lancamentos_records(ws):
    return records_from_values(ws.get_all_values(value_render_option='UNFORMATTED_VALUE'))

def cell(r, *names, default=''):
    for name in names:
        value = r.get(name)
        if value not in (None, ''):
            return value
    return default


def norm_text(value):
    return str(value or '').strip()


def same_text(value, expected):
    return norm_text(value).lower() == norm_text(expected).lower()


def get_valor(r):
    return to_float(cell(r, 'Valor (R$)', 'Valor', default=0))


def format_data(value):
    parsed = parse_data(value)
    if not parsed:
        return norm_text(value)
    return parsed.strftime('%d/%m/%Y')


def norm(r):
    return {
        'data':      format_data(cell(r, 'Data')),
        'fonte':     norm_text(cell(r, 'Fonte')),
        'item':      norm_text(cell(r, 'Descrição', 'Descricao', 'Item')),
        'valor':     get_valor(r),
        'categoria': norm_text(cell(r, 'Categoria')),
        'geral':     norm_text(cell(r, 'Subcategoria', 'Geral')),
        'situacao':  norm_text(cell(r, 'Situação', 'Situacao')),
        'tipo':      norm_text(cell(r, 'Tipo')),
    }

def serial_to_label(serial):
    """Converte serial do Excel para label 'jan-26'."""
    MN = ['jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez']
    try:
        s = int(float(str(serial)))
        d = date(1899, 12, 30)
        from datetime import timedelta
        d = d + timedelta(days=s)
        return MN[d.month - 1] + '-' + str(d.year)[-2:]
    except:
        return str(serial)


def month_serial(value):
    text = norm_text(value).lower()
    try:
        serial = int(float(text))
        from datetime import timedelta
        dt = date(1899, 12, 30) + timedelta(days=serial)
        return (date(dt.year, dt.month, 1) - date(1899, 12, 30)).days
    except ValueError:
        pass

    aliases = {
        'jan': 1, 'janeiro': 1,
        'fev': 2, 'fevereiro': 2,
        'mar': 3, 'marco': 3, 'março': 3,
        'abr': 4, 'abril': 4,
        'mai': 5, 'maio': 5,
        'jun': 6, 'junho': 6,
        'jul': 7, 'julho': 7,
        'ago': 8, 'agosto': 8,
        'set': 9, 'setembro': 9,
        'out': 10, 'outubro': 10,
        'nov': 11, 'novembro': 11,
        'dez': 12, 'dezembro': 12,
    }
    match = re.search(r'([a-zç]+)\.?\s*[/\-]\s*(\d{2,4})', text)
    if not match:
        return value
    month = aliases.get(match.group(1).replace('.', ''))
    year = int(match.group(2))
    if year < 100:
        year += 2000
    if not month:
        return value
    return (date(year, month, 1) - date(1899, 12, 30)).days


def serial_to_month_date(serial):
    try:
        serial = int(float(serial))
        from datetime import timedelta
        dt = date(1899, 12, 30) + timedelta(days=serial)
        return date(dt.year, dt.month, 1)
    except (ValueError, TypeError):
        parsed = parse_mes_input(serial)
        return parsed


def parse_mes_input(value):
    text = norm_text(value).lower()
    if not text:
        return None

    aliases = {
        'jan': 1, 'janeiro': 1,
        'fev': 2, 'fevereiro': 2,
        'mar': 3, 'marco': 3, 'março': 3,
        'abr': 4, 'abril': 4,
        'mai': 5, 'maio': 5,
        'jun': 6, 'junho': 6,
        'jul': 7, 'julho': 7,
        'ago': 8, 'agosto': 8,
        'set': 9, 'setembro': 9,
        'out': 10, 'outubro': 10,
        'nov': 11, 'novembro': 11,
        'dez': 12, 'dezembro': 12,
    }

    for fmt in ('%Y-%m', '%m/%Y', '%m-%Y'):
        try:
            parsed = datetime.strptime(text, fmt).date()
            return date(parsed.year, parsed.month, 1)
        except ValueError:
            pass

    match = re.search(r'([a-zç]+)\.?\s*[/\- ]\s*(\d{2,4})', text)
    if match:
        month = aliases.get(match.group(1).replace('.', ''))
        year = int(match.group(2))
        if year < 100:
            year += 2000
        if month:
            return date(year, month, 1)

    try:
        return serial_to_month_date(int(float(text)))
    except (ValueError, TypeError):
        return None


def month_date_to_serial(month_date):
    return (date(month_date.year, month_date.month, 1) - date(1899, 12, 30)).days


def month_date_to_label(month_date):
    meses = ['jan.', 'fev.', 'mar.', 'abr.', 'mai.', 'jun.', 'jul.', 'ago.', 'set.', 'out.', 'nov.', 'dez.']
    return f'{meses[month_date.month - 1]}/{month_date.year}'


def orcamento_col_map(rows):
    grupos_row = rows[1] if len(rows) > 1 else []
    sub_row = rows[3] if len(rows) > 3 else []
    col_map = {}
    for ci, sub in enumerate(sub_row):
        if ci == 0:
            continue
        sub_clean = norm_text(sub).upper()
        if sub_clean == 'RECEITAS':
            col_map[ci] = {'grupo': 'RECEITAS', 'cat': 'RECEITAS', 'tipo': 'receita'}
        elif ' - PREVISTO' in sub_clean:
            cat = sub_clean.replace(' - PREVISTO', '').strip()
            grupo = norm_text(grupos_row[ci]) if ci < len(grupos_row) else ''
            col_map[ci] = {'grupo': grupo, 'cat': cat, 'tipo': 'previsto'}
        elif ' - REALIZADO' in sub_clean:
            cat = sub_clean.replace(' - REALIZADO', '').strip()
            grupo = norm_text(grupos_row[ci]) if ci < len(grupos_row) else ''
            col_map[ci] = {'grupo': grupo, 'cat': cat, 'tipo': 'realizado'}
    return col_map

def to_float(v):
    if isinstance(v, (int, float)):
        return float(v)
    try:
        raw = str(v).replace('R$', '').strip()
        raw = raw.replace(' ', '')
        if not raw:
            return 0.0
        if ',' in raw and '.' in raw:
            raw = raw.replace('.', '').replace(',', '.')
        elif ',' in raw:
            raw = raw.replace(',', '.')
        return float(Decimal(raw))
    except (InvalidOperation, ValueError):
        return 0.0


def parse_data(value):
    text = norm_text(value)
    for fmt in ('%d/%m/%Y', '%Y-%m-%d', '%d/%m/%y'):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    try:
        serial = int(float(text))
        from datetime import timedelta
        return date(1899, 12, 30) + timedelta(days=serial)
    except (ValueError, TypeError):
        return None


def mes_match(row, mes):
    if not mes:
        return True
    data = parse_data(cell(row, 'Data'))
    if not data:
        return False
    mes = norm_text(mes)
    if '-' in mes:
        try:
            y, m = mes.split('-', 1)
            return data.year == int(y) and data.month == int(m)
        except ValueError:
            return False
    if '/' in mes:
        parts = mes.split('/')
        try:
            if len(parts) == 2:
                return data.month == int(parts[0]) and data.year == int(parts[1])
            if len(parts) == 3:
                return data == datetime.strptime(mes, '%d/%m/%Y').date()
        except ValueError:
            return False
    return False


def filtrar_lancamentos(rows):
    mes = request.args.get('mes')
    categoria = request.args.get('categoria')
    tipo = request.args.get('tipo')
    situacao = request.args.get('situacao')

    if mes:
        rows = [r for r in rows if mes_match(r, mes)]
    if categoria:
        rows = [r for r in rows if same_text(cell(r, 'Categoria'), categoria)]
    if tipo:
        rows = [r for r in rows if same_text(cell(r, 'Tipo'), tipo)]
    if situacao:
        rows = [r for r in rows if same_text(cell(r, 'Situação', 'Situacao'), situacao)]
    return rows


def payload_json():
    dados = request.get_json(silent=True)
    if not isinstance(dados, dict):
        return None, api_error('JSON inválido ou ausente', 400)
    return dados, None


def validar_lancamento(dados):
    obrigatorios = ['data', 'fonte', 'item', 'valor', 'categoria', 'situacao', 'tipo']
    for campo in obrigatorios:
        if campo not in dados or dados[campo] in (None, ''):
            return f'Campo obrigatório: {campo}'
    if to_float(dados['valor']) <= 0:
        return 'Campo valor deve ser maior que zero'
    return None


def normalizar_data_input(data_str):
    data_str = norm_text(data_str)
    if '-' in data_str and len(data_str) == 10:
        y, m, d = data_str.split('-')
        return f'{d}/{m}/{y}'
    return data_str


def linha_lancamento(dados):
    return [
        normalizar_data_input(dados['data']),
        dados['fonte'],
        dados['item'],
        to_float(dados['valor']),
        norm_text(dados['categoria']).upper(),
        dados.get('subcategoria', dados.get('geral', '')),
        dados['situacao'],
        dados['tipo'],
    ]


def listar_lancamentos_data():
    sheet = get_sheet()
    ws = sheet.worksheet('Lançamentos')
    rows = filtrar_lancamentos(get_lancamentos_records(ws))
    normalized = [norm(r) for r in rows]
    return normalized, len(rows)


def criar_lancamento(dados):
    erro_validacao = validar_lancamento(dados)
    if erro_validacao:
        return api_error(erro_validacao, 400)
    sheet = get_sheet()
    ws = sheet.worksheet('Lançamentos')
    ws.append_row(linha_lancamento(dados), value_input_option='USER_ENTERED')
    return api_response(True, mensagem='Lançamento adicionado com sucesso')


# ── /status ───────────────────────────────────────────────────
@app.route('/status', methods=['GET'])
def status():
    return jsonify({'ok': True, 'timestamp': datetime.now().isoformat()})


# ── /health ───────────────────────────────────────────────────
@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'timestamp': datetime.now().isoformat()})


# ── /listar ───────────────────────────────────────────────────
@app.route('/listar', methods=['GET'])
def listar():
    try:
        rows, total = listar_lancamentos_data()
        return api_response(True, movimentacoes=rows, data=rows, total=total)

    except Exception as e:
        return api_error('Não foi possível listar os lançamentos', exc=e)


# ── /adicionar ────────────────────────────────────────────────
@app.route('/adicionar', methods=['POST'])
def adicionar():
    try:
        dados, erro = payload_json()
        if erro:
            return erro
        return criar_lancamento(dados)

    except Exception as e:
        return api_error('Não foi possível adicionar o lançamento', exc=e)


# ── /lancamentos GET ──────────────────────────────────────────
@app.route('/lancamentos', methods=['GET'])
def listar_lancamentos():
    try:
        rows, total = listar_lancamentos_data()
        return api_response(True, data=rows, movimentacoes=rows, total=total)
    except Exception as e:
        return api_error('Não foi possível listar os lançamentos', exc=e)


# ── /lancamentos POST ─────────────────────────────────────────
@app.route('/lancamentos', methods=['POST'])
def adicionar_lancamento():
    try:
        dados, erro = payload_json()
        if erro:
            return erro
        return criar_lancamento(dados)
    except Exception as e:
        return api_error('Não foi possível adicionar o lançamento', exc=e)


# ── /lancamentos/<linha> PUT ───────────────────────────────
@app.route('/lancamentos/<int:linha>', methods=['PUT'])
def atualizar_lancamento(linha):
    if linha < 2:
        return api_error('Linha inválida', 400)
    try:
        dados, erro = payload_json()
        if erro:
            return erro
        erro_validacao = validar_lancamento(dados)
        if erro_validacao:
            return api_error(erro_validacao, 400)

        sheet = get_sheet()
        ws = sheet.worksheet('Lançamentos')
        ws.update(f'A{linha}:H{linha}', [linha_lancamento(dados)], value_input_option='USER_ENTERED')
        return api_response(True, mensagem='Lançamento atualizado')
    except Exception as e:
        return api_error('Não foi possível atualizar o lançamento', exc=e)


# ── /lancamentos/<linha> DELETE ────────────────────────────
@app.route('/lancamentos/<int:linha>', methods=['DELETE'])
def remover_lancamento(linha):
    if linha < 2:
        return api_error('Linha inválida', 400)
    try:
        sheet = get_sheet()
        ws = sheet.worksheet('Lançamentos')
        ws.delete_rows(linha)
        return api_response(True, mensagem='Lançamento removido')
    except Exception as e:
        return api_error('Não foi possível remover o lançamento', exc=e)


# ── /orcamento ────────────────────────────────────────────────
@app.route('/orcamento', methods=['GET'])
def listar_orcamento():
    """
    Estrutura da planilha 'Desembolso Trimestral':
      Linha 1: título
      Linha 2: grupos (DESPESAS BÁSICAS, INVESTIMENTOS, LAZER, DIVERSO, REEMBOLSOS)
      Linha 3: categorias (MERCADO, FIXOS, ...) — col A = 'Mês'
      Linha 4: sub-cabeçalhos (CAT - PREVISTO, CAT - REALIZADO, ...)
      Linha 5+: dados — col A = serial de data do Excel
    """
    try:
        sheet = get_sheet()
        ws = sheet.worksheet('Desembolso Trimestral')
        rows = ws.get_all_values()

        # Monta mapa de colunas: índice → {grupo, cat, tipo: 'previsto'|'realizado'}
        col_map = orcamento_col_map(rows)

        # Coleta categorias únicas em ordem
        cats_seen = {}
        for ci, info in col_map.items():
            key = info['cat']
            if key not in cats_seen:
                cats_seen[key] = {'grupo': info['grupo'], 'cat': key, 'prev_col': None, 'real_col': None}
            if info['tipo'] == 'previsto':
                cats_seen[key]['prev_col'] = ci
            else:
                cats_seen[key]['real_col'] = ci

        # Processa meses (linhas 5+ = índice 4+)
        meses = []
        # estrutura: {serial: {cat: {previsto, realizado}}}
        mes_data = {}

        for row in rows[4:]:
            if not row or not row[0]:
                continue
            serial = row[0]
            serial_key = month_serial(serial)
            label = serial_to_label(serial)
            if serial_key not in mes_data:
                mes_data[serial_key] = {'label': label, 'serial': serial_key, 'receita': 0.0, 'cats': {}}
                meses.append({'serial': serial_key, 'label': label, 'receita': 0.0})

            for ci, info in col_map.items():
                if ci >= len(row):
                    continue
                if info['tipo'] == 'receita':
                    receita = to_float(row[ci])
                    mes_data[serial_key]['receita'] = receita
                    meses[-1]['receita'] = receita
                    continue
                cat = info['cat']
                if cat not in mes_data[serial_key]['cats']:
                    mes_data[serial_key]['cats'][cat] = {'previsto': 0.0, 'realizado': 0.0}
                v = to_float(row[ci])
                mes_data[serial_key]['cats'][cat][info['tipo']] = v

        # Monta resposta no formato esperado pelo app HTML
        # linhas: lista de {grupo, cod, desc, meses: [{previsto, realizado, saldo}]}
        linhas = []
        for cat_info in cats_seen.values():
            if cat_info['cat'] == 'RECEITAS':
                continue
            cat = cat_info['cat']
            cat_meses = []
            for m in meses:
                cd = mes_data[m['serial']]['cats'].get(cat, {'previsto': 0.0, 'realizado': 0.0})
                prev = cd['previsto']
                real = cd['realizado']
                cat_meses.append({
                    'previsto':  round(prev, 2),
                    'realizado': round(real, 2),
                    'saldo':     round(prev - real, 2)
                })
            linhas.append({
                'grupo':  cat_info['grupo'],
                'cod':    cat,
                'desc':   cat,
                'meses':  cat_meses
            })

        return api_response(True, meses=meses, linhas=linhas)

    except Exception as e:
        return api_error('Não foi possível carregar o orçamento', exc=e)


# ── /orcamento POST ─────────────────────────────────────────
@app.route('/orcamento', methods=['POST'])
def salvar_orcamento():
    try:
        dados, erro = payload_json()
        if erro:
            return erro

        mes = parse_mes_input(dados.get('mes'))
        if not mes:
            return api_error('Informe um mês válido, como 06/2026 ou jun/2026', 400)

        valores = dados.get('valores', {})
        if not isinstance(valores, dict):
            return api_error('Campo valores deve ser um objeto', 400)
        receita_payload = dados.get('receita', None)

        sheet = get_sheet()
        ws = sheet.worksheet('Desembolso Trimestral')
        rows = ws.get_all_values(value_render_option='FORMULA')
        col_map = orcamento_col_map(rows)
        if not col_map:
            return api_error('Cabeçalho do orçamento não encontrado')

        target_serial = month_date_to_serial(mes)
        target_row_idx = None
        for idx, row in enumerate(rows[4:], start=5):
            if row and month_serial(row[0]) == target_serial:
                target_row_idx = idx
                break

        width = max(col_map.keys()) + 1
        new_row_idx = target_row_idx or (len(rows) + 1)
        existing = []
        if target_row_idx:
            existing = rows[target_row_idx - 1]
        linha = existing[:width] + [''] * max(0, width - len(existing))
        linha[0] = month_date_to_label(mes)

        valores_norm = {norm_text(k).upper(): to_float(v) for k, v in valores.items() if to_float(v) > 0}
        receita_col = next((ci for ci, info in col_map.items() if info['tipo'] == 'receita'), None)
        receita = to_float(receita_payload) if receita_payload not in (None, '') else 0.0
        if receita_col is not None:
            if receita <= 0 and target_row_idx and receita_col < len(linha):
                receita = to_float(linha[receita_col])
            if receita <= 0:
                return api_error('Informe a receita do mês', 400)

        total_previsto = sum(valores_norm.values())
        if receita > 0 and total_previsto > receita:
            return api_error(
                f'A soma das previsões (R$ {total_previsto:.2f}) não pode passar da receita (R$ {receita:.2f})',
                400
            )

        if receita_col is not None:
            linha[receita_col] = receita

        for ci, info in col_map.items():
            if info['tipo'] == 'receita':
                continue
            if info['tipo'] == 'previsto':
                if info['cat'] in valores_norm:
                    linha[ci] = valores_norm[info['cat']]
                elif not target_row_idx:
                    linha[ci] = 0
            elif info['tipo'] == 'realizado' and not target_row_idx:
                linha[ci] = f'=SUMIFS(\'Lançamentos\'!$D:$D;\'Lançamentos\'!$E:$E;"{info["cat"]}";\'Lançamentos\'!$A:$A;$A{new_row_idx})'

        if target_row_idx:
            ws.update(f'A{target_row_idx}:{rowcol_to_a1(target_row_idx, width)}', [linha], value_input_option='USER_ENTERED')
            acao = 'atualizado'
        else:
            ws.append_row(linha, value_input_option='USER_ENTERED')
            acao = 'criado'

        return api_response(
            True,
            mensagem=f'Orçamento {acao} com sucesso',
            mes={'serial': target_serial, 'label': month_date_to_label(mes)}
        )
    except Exception as e:
        return api_error('Não foi possível salvar o orçamento', exc=e)



def pdf_escape(value):
    return str(value).replace('\\', '\\\\').replace('(', '\\(').replace(')', '\\)')


def brl(value):
    try:
        value = float(value)
    except (TypeError, ValueError):
        value = 0.0
    formatted = f'{abs(value):,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')
    return ('-R$ ' if value < 0 else 'R$ ') + formatted


def parse_br_date(value):
    parsed = parse_data(value)
    if parsed:
        return parsed
    return None


def make_simple_pdf(title, lines):
    width, height = 595, 842
    margin_x, y = 48, 790
    content = ['BT', '/F1 18 Tf', f'{margin_x} {y} Td', f'({pdf_escape(title)}) Tj']
    y -= 28
    content += ['/F1 9 Tf', f'{margin_x} {y} Td', f'(Gerado em {datetime.now().strftime("%d/%m/%Y %H:%M")}) Tj']
    y -= 28
    content += ['/F1 10 Tf']
    for line in lines:
        if y < 56:
            break
        if line == '':
            y -= 10
            continue
        if line.startswith('# '):
            y -= 8
            content += ['/F1 13 Tf', f'{margin_x} {y} Td', f'({pdf_escape(line[2:])}) Tj', '/F1 10 Tf']
            y -= 18
            continue
        safe = pdf_escape(line[:105])
        content += [f'{margin_x} {y} Td', f'({safe}) Tj']
        y -= 15
    content.append('ET')
    stream = '\n'.join(content).encode('latin-1', 'replace')
    objects = []
    objects.append(b'<< /Type /Catalog /Pages 2 0 R >>')
    objects.append(b'<< /Type /Pages /Kids [3 0 R] /Count 1 >>')
    objects.append(b'<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>')
    objects.append(b'<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>')
    objects.append(b'<< /Length ' + str(len(stream)).encode() + b' >>\nstream\n' + stream + b'\nendstream')
    pdf = bytearray(b'%PDF-1.4\n')
    offsets = []
    for idx, obj in enumerate(objects, start=1):
        offsets.append(len(pdf))
        pdf.extend(f'{idx} 0 obj\n'.encode())
        pdf.extend(obj)
        pdf.extend(b'\nendobj\n')
    xref = len(pdf)
    pdf.extend(f'xref\n0 {len(objects)+1}\n'.encode())
    pdf.extend(b'0000000000 65535 f \n')
    for off in offsets:
        pdf.extend(f'{off:010d} 00000 n \n'.encode())
    pdf.extend(f'trailer\n<< /Size {len(objects)+1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF'.encode())
    return bytes(pdf)


def resumo_executivo_data():
    sheet = get_sheet()
    lanc_rows = [norm(r) for r in get_lancamentos_records(sheet.worksheet('Lançamentos'))]
    today = date.today()
    month_rows = []
    paid_rows = []
    for row in lanc_rows:
        parsed = parse_br_date(row.get('data'))
        if parsed and parsed.year == today.year and parsed.month == today.month:
            month_rows.append(row)
            if same_text(row.get('situacao'), 'Pago'):
                paid_rows.append(row)

    despesas = [r for r in month_rows if same_text(r.get('tipo'), 'Despesa')]
    receitas = [r for r in month_rows if same_text(r.get('tipo'), 'Receita')]
    total_desp = sum(to_float(r.get('valor')) for r in despesas)
    total_rec = sum(to_float(r.get('valor')) for r in receitas)
    pendentes = sum(to_float(r.get('valor')) for r in month_rows if same_text(r.get('situacao'), 'Pendente'))
    por_cat = {}
    for r in despesas:
        cat = norm_text(r.get('categoria')) or 'OUTRA'
        por_cat[cat] = por_cat.get(cat, 0) + to_float(r.get('valor'))
    top_cats = sorted(por_cat.items(), key=lambda x: x[1], reverse=True)[:8]

    orc = None
    try:
        ws = sheet.worksheet('Desembolso Trimestral')
        rows = ws.get_all_values()
        col_map = orcamento_col_map(rows)
        target_serial = month_date_to_serial(date(today.year, today.month, 1))
        target_row = next((row for row in rows[4:] if row and month_serial(row[0]) == target_serial), None)
        if target_row:
            total_prev = total_real = receita_prev = 0.0
            estouros = []
            for ci, info in col_map.items():
                value = to_float(target_row[ci]) if ci < len(target_row) else 0.0
                if info['tipo'] == 'receita':
                    receita_prev = value
                elif info['tipo'] == 'previsto':
                    total_prev += value
                elif info['tipo'] == 'realizado':
                    total_real += value
            for ci, info in col_map.items():
                if info['tipo'] != 'previsto':
                    continue
                previsto = to_float(target_row[ci]) if ci < len(target_row) else 0.0
                real_col = next((rci for rci, rinfo in col_map.items() if rinfo['cat'] == info['cat'] and rinfo['tipo'] == 'realizado'), None)
                realizado = to_float(target_row[real_col]) if real_col is not None and real_col < len(target_row) else 0.0
                if previsto > 0 and realizado > previsto:
                    estouros.append((info['cat'], realizado, previsto, realizado - previsto))
            orc = {'receita': receita_prev, 'previsto': total_prev, 'realizado': total_real, 'estouros': sorted(estouros, key=lambda x: x[3], reverse=True)}
    except Exception:
        orc = None

    return {
        'mes': month_date_to_label(date(today.year, today.month, 1)),
        'despesas': total_desp,
        'receitas': total_rec,
        'saldo': total_rec - total_desp,
        'pendentes': pendentes,
        'lancamentos': len(month_rows),
        'top_cats': top_cats,
        'orc': orc,
    }

# ── /exportar ───────────────────────────────────────────────
@app.route('/exportar', methods=['POST'])
def exportar_planilha():
    try:
        sheet = get_sheet()
        workbook = Workbook()
        default_ws = workbook.active
        workbook.remove(default_ws)

        for worksheet in sheet.worksheets():
            title = worksheet.title[:31] or 'Planilha'
            ws_out = workbook.create_sheet(title=title)
            values = worksheet.get_all_values()
            if not values:
                ws_out.append([])
                continue
            for row in values:
                ws_out.append(row)

            for column_cells in ws_out.columns:
                max_len = max((len(str(cell.value)) for cell in column_cells if cell.value is not None), default=0)
                ws_out.column_dimensions[column_cells[0].column_letter].width = min(max(max_len + 2, 10), 42)

        output = BytesIO()
        workbook.save(output)
        output.seek(0)
        filename = f"Desembolso_Estruturado_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.xlsx"
        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        return api_error('Não foi possível exportar a planilha', exc=e)


@app.route('/exportar_pdf', methods=['POST'])
def exportar_pdf():
    try:
        data = resumo_executivo_data()
        lines = [
            f"Periodo: {data['mes']}",
            '',
            '# Resumo do mes',
            f"Receitas: {brl(data['receitas'])}",
            f"Despesas: {brl(data['despesas'])}",
            f"Saldo: {brl(data['saldo'])}",
            f"Pendentes: {brl(data['pendentes'])}",
            f"Lancamentos no mes: {data['lancamentos']}",
            '',
            '# Top categorias de despesa',
        ]
        if data['top_cats']:
            lines.extend([f"{idx+1}. {cat}: {brl(total)}" for idx, (cat, total) in enumerate(data['top_cats'])])
        else:
            lines.append('Sem despesas no periodo atual.')
        if data['orc']:
            orc = data['orc']
            lines.extend([
                '',
                '# Orcamento',
                f"Receita prevista: {brl(orc['receita'])}",
                f"Previsoes cadastradas: {brl(orc['previsto'])}",
                f"Realizado no orcamento: {brl(orc['realizado'])}",
                f"Saldo livre previsto: {brl(orc['receita'] - orc['previsto'])}",
                '',
                '# Alertas de orcamento',
            ])
            if orc['estouros']:
                lines.extend([f"{cat}: realizado {brl(real)} / previsto {brl(prev)} | excesso {brl(over)}" for cat, real, prev, over in orc['estouros'][:8]])
            else:
                lines.append('Nenhum estouro encontrado para o mes atual.')
        else:
            lines.extend(['', '# Orcamento', 'Orcamento do mes atual nao encontrado.'])
        pdf = make_simple_pdf('Controle Financeiro - Visao Executiva', lines)
        output = BytesIO(pdf)
        filename = f"Controle_Financeiro_Executivo_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.pdf"
        return send_file(output, as_attachment=True, download_name=filename, mimetype='application/pdf')
    except Exception as e:
        return api_error('Não foi possível gerar o PDF executivo', exc=e)


# ── /resumo ───────────────────────────────────────────────────
@app.route('/resumo', methods=['GET'])
def resumo():
    try:
        sheet = get_sheet()
        ws = sheet.worksheet('Lançamentos')
        rows = get_lancamentos_records(ws)
        total_d = sum(get_valor(r) for r in rows if same_text(cell(r, 'Tipo'), 'Despesa') and same_text(cell(r, 'Situação', 'Situacao'), 'Pago'))
        total_i = sum(get_valor(r) for r in rows if same_text(cell(r, 'Tipo'), 'Investimento') and same_text(cell(r, 'Situação', 'Situacao'), 'Pago'))
        total_p = sum(get_valor(r) for r in rows if same_text(cell(r, 'Situação', 'Situacao'), 'Pendente'))
        por_cat = {}
        for r in rows:
            if same_text(cell(r, 'Tipo'), 'Despesa') and same_text(cell(r, 'Situação', 'Situacao'), 'Pago'):
                c = norm_text(cell(r, 'Categoria', default='OUTROS')) or 'OUTROS'
                por_cat[c] = por_cat.get(c, 0) + get_valor(r)
        return api_response(
            True,
            total_despesas=round(total_d, 2),
            total_investido=round(total_i, 2),
            total_pendente=round(total_p, 2),
            por_categoria={k: round(v, 2) for k, v in sorted(por_cat.items(), key=lambda x: -x[1])}
        )
    except Exception as e:
        return api_error('Não foi possível carregar o resumo', exc=e)


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5050))
    app.run(host='0.0.0.0', port=port, debug=False)
