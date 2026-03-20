# -*- coding: utf-8 -*-
# app.py — Flask 메인 (라우팅 진입점)

from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os, sys, time, threading, json, requests
from datetime import datetime, timedelta

# ── 선택적 임포트 ──
try:
    from statsmodels.tsa.arima.model import ARIMA
    ARIMA_OK = True
except ImportError:
    ARIMA_OK = False
    print('⚠ statsmodels 미설치 (py -m pip install statsmodels)')

try:
    from prophet import Prophet
    PROPHET_OK = True
except ImportError:
    PROPHET_OK = False
    print('⚠ prophet 미설치 (py -m pip install prophet)')

try:
    import FinanceDataReader as fdr
    FDR_OK = True
except Exception as e:
    fdr = None
    FDR_OK = False
    print(f'⚠ FinanceDataReader 로드 실패: {e}')

import yfinance as yf
import pandas as pd

# ── .env 로드 ──
try:
    from dotenv import load_dotenv
    load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env'))
    print('✅ .env 로드 완료')
except ImportError:
    print('⚠ python-dotenv 미설치')

# ── Flask 앱 ──
app = Flask(__name__)
CORS(app)

def get_base_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE_DIR = get_base_path()

# ── CSV 종목 데이터 ──
krx_list  = []
target_csv_files = ['한국상장법인.csv', '미국snp.csv', '미국Nyse.csv', 'ETF.csv', 'ETN.csv']
file_mtimes = {f: 0 for f in target_csv_files}

def load_csvs_if_changed():
    global krx_list, file_mtimes
    changed = False
    for fname in target_csv_files:
        path = os.path.join(BASE_DIR, fname)
        if not os.path.exists(path): continue
        mtime = os.path.getmtime(path)
        if mtime != file_mtimes[fname]:
            file_mtimes[fname] = mtime
            changed = True
    if not changed and krx_list: return
    new_list = []
    import csv
    mkt_map = {'한국상장법인.csv': 'KOSPI', '미국snp.csv': 'S&P500',
               '미국Nyse.csv': 'NYSE', 'ETF.csv': 'ETF', 'ETN.csv': 'ETN'}
    for fname in target_csv_files:
        path = os.path.join(BASE_DIR, fname)
        if not os.path.exists(path): continue
        mkt = mkt_map.get(fname, '기타')
        for enc in ['utf-8-sig', 'cp949', 'euc-kr', 'utf-8']:
            try:
                with open(path, encoding=enc) as f:
                    for row in csv.DictReader(f):
                        name = row.get('Name') or row.get('회사명') or ''
                        sym  = row.get('Symbol') or row.get('티커') or row.get('단축코드') or ''
                        if name and sym:
                            new_list.append({'name': name.strip(), 'symbol': sym.strip(), 'market': mkt})
                break  # 성공하면 다음 파일로
            except UnicodeDecodeError:
                continue
            except Exception as e:
                print(f'CSV 로드 오류 {fname}: {e}')
                break
    krx_list = new_list
    print(f'✅ CSV 로드 완료: {len(krx_list)}개 종목')

# ── KIS 인증 ──
KIS_APP_KEY    = os.environ.get('KIS_APP_KEY', '')
KIS_APP_SECRET = os.environ.get('KIS_APP_SECRET', '')
KIS_ACCOUNT    = os.environ.get('KIS_ACCOUNT', '')
KIS_ACCOUNT_SUFFIX = os.environ.get('KIS_ACCOUNT_SUFFIX', '01')
KIS_BASE_URL   = 'https://openapi.koreainvestment.com:9443'

_kis_token = {'token': None, 'expires': 0}
_kis_lock  = threading.Lock()

def kis_get_token():
    with _kis_lock:
        now = time.time()
        if _kis_token['token'] and now < _kis_token['expires']:
            return _kis_token['token']
        if not KIS_APP_KEY: return None
        try:
            resp = requests.post(f'{KIS_BASE_URL}/oauth2/tokenP',
                headers={'Content-Type': 'application/json'},
                json={'grant_type': 'client_credentials',
                      'appkey': KIS_APP_KEY, 'appsecret': KIS_APP_SECRET}, timeout=10)
            t = resp.json().get('access_token')
            if t:
                _kis_token['token']   = t
                _kis_token['expires'] = now + 86000
                print('✅ KIS 토큰 발급')
                return t
        except Exception as e:
            print(f'❌ KIS 토큰 오류: {e}')
        return None

def kis_headers(tr_id, token=None):
    t = token or kis_get_token()
    return {'Content-Type': 'application/json', 'authorization': f'Bearer {t}',
            'appkey': KIS_APP_KEY, 'appsecret': KIS_APP_SECRET,
            'tr_id': tr_id, 'custtype': 'P'}

# ── 공유 컨텍스트 (Blueprint에 주입) ──
shared = {
    'fdr': fdr, 'FDR_OK': FDR_OK,
    'ARIMA_OK': ARIMA_OK, 'PROPHET_OK': PROPHET_OK,
    'krx_list': lambda: krx_list,
    'load_csvs_if_changed': load_csvs_if_changed,
    'kis_get_token': kis_get_token,
    'kis_headers': kis_headers,
    'KIS_BASE_URL': KIS_BASE_URL,
    'KIS_ACCOUNT': KIS_ACCOUNT,
    'KIS_ACCOUNT_SUFFIX': KIS_ACCOUNT_SUFFIX,
    'get_base_path': get_base_path,
}

# ── Blueprint 등록 ──
from routes.data     import data_bp,     init_data
from routes.kis      import kis_bp,      init_kis
from routes.forecast import forecast_bp, init_forecast
from routes.us       import us_bp

init_data(shared)
init_kis(shared)
init_forecast(shared)

app.register_blueprint(data_bp)
app.register_blueprint(kis_bp)
app.register_blueprint(forecast_bp)
app.register_blueprint(us_bp)

# ── 메인 라우트 ──
@app.route('/')
def home():
    path = os.path.join(BASE_DIR, 'stock_chart_app.html')
    if not os.path.exists(path):
        return f'stock_chart_app.html 파일이 없습니다 (경로: {path})'
    resp = send_file(path)
    resp.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    resp.headers['Pragma']  = 'no-cache'
    resp.headers['Expires'] = '0'
    return resp

# ── AI 채팅 ──
@app.route('/ai/chat', methods=['POST'])
def ai_chat():
    import anthropic
    try:
        body     = request.get_json()
        messages = body.get('messages', [])
        system   = body.get('system', '')
        normalized = []
        for msg in messages:
            content = msg.get('content', '')
            if isinstance(content, str):
                normalized.append({'role': msg['role'], 'content': [{'type':'text','text':content}]})
            else:
                normalized.append(msg)
        client = anthropic.Anthropic()
        resp   = client.messages.create(
            model='claude-sonnet-4-20250514', max_tokens=1500,
            system=system, messages=normalized)
        return jsonify({'content': resp.content[0].text})
    except Exception as e:
        print(f'AI 오류: {e}')
        return jsonify({'error': str(e)}), 500

# ── 정적 파일 서빙 ──
@app.route('/<path:filename>')
def serve_static(filename):
    path = os.path.join(get_base_path(), filename)
    if os.path.exists(path):
        return send_file(path)
    return jsonify({'error': '파일 없음'}), 404

# ── 시작 ──
if __name__ == '__main__':
    load_csvs_if_changed()
    app.run(port=5000, debug=True)
