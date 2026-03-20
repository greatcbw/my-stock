# -*- coding: utf-8 -*-
# routes/kis.py — 한국투자증권 KIS API

from flask import Blueprint, request, jsonify
import requests
from datetime import datetime
import os

from extensions import cache # 캐시 불러오기 추가

kis_bp = Blueprint('kis', __name__)

_shared = {}

def init_kis(shared):
    global _shared
    _shared = shared

def _get_token(): return _shared.get('kis_get_token', lambda: None)()
def _headers(tr_id, token=None): return _shared.get('kis_headers', lambda a,b: {})(tr_id, token)
def _base(): return _shared.get('KIS_BASE_URL', '')
def _account(): return _shared.get('KIS_ACCOUNT', '')
def _suffix(): return _shared.get('KIS_ACCOUNT_SUFFIX', '01')


@kis_bp.route('/kis/price')
@cache.cached(timeout=5, query_string=True) # 🌟 호가는 중요하므로 5초만 캐싱!
# kis.py 의 kis_price 라우트 함수 전체를 아래 코드로 교체해주세요.

@kis_bp.route('/kis/price')
@cache.cached(timeout=5, query_string=True) # Phase 2에서 추가한 캐시 유지
def kis_price():
    symbol = request.args.get('symbol', '')
    if not symbol:
        return jsonify({'error': '종목코드 필요'}), 400
        
    token = _get_token()
    
    # ── 1. KIS API 정상 호출 시도 ──
    if token:
        try:
            resp = requests.get(
                f'{_base()}/uapi/domestic-stock/v1/quotations/inquire-price',
                headers=_headers('FHKST01010100', token),
                params={'FID_COND_MRKT_DIV_CODE': 'J', 'FID_INPUT_ISCD': symbol},
                timeout=3 # 3초 이상 응답 없으면 실패로 간주
            )
            data = resp.json().get('output')
            if data:
                return jsonify({
                    'price': int(data.get('stck_prpr', 0)),
                    'change': int(data.get('prdy_vrss', 0)),
                    'change_pct': float(data.get('prdy_ctrt', 0)),
                    'volume': int(data.get('acml_vol', 0)),
                    'is_fallback': False # 정상 데이터임을 프론트엔드에 알림
                })
        except Exception as e:
            print(f"⚠ KIS API 지연/오류 발생. Fallback 실행: {e}")

    # ── 2. KIS 실패 시 Fallback: yfinance 로 우회 (주말/점검용) ──
    try:
        import yfinance as yf
        import pandas as pd
        
        # 코스피(.KS)인지 코스닥(.KQ)인지 판별 (단순화 로직)
        # 종목 코드만으로 yfinance 조회를 위해 기본적으로 .KS 시도 후 실패시 .KQ를 쓸 수도 있지만,
        # 야후 파이낸스는 한국 주식에 .KS / .KQ를 붙여야 합니다.
        yf_symbol = f"{symbol}.KS"
        tk = yf.Ticker(yf_symbol)
        
        # 만약 코스피가 아니라면(info가 비어있다면) 코스닥으로 재시도
        if not tk.fast_info.get('last_price'):
            yf_symbol = f"{symbol}.KQ"
            tk = yf.Ticker(yf_symbol)
            
        info = tk.fast_info
        hist = tk.history(period='2d')
        
        price = float(info.last_price or 0)
        prev_close = float(hist['Close'].iloc[-2]) if len(hist) >= 2 else price
        change = price - prev_close
        change_pct = (change / prev_close * 100) if prev_close else 0

        return jsonify({
            'price': int(price),
            'change': int(change),
            'change_pct': round(change_pct, 2),
            'volume': int(info.last_volume or 0),
            'is_fallback': True, # 야후 데이터임을 알림
            'msg': '⚠️ KIS 점검중 (야후 15분 지연 데이터)'
        })
    except Exception as fallback_e:
        print(f"Fallback 마저 실패: {fallback_e}")
        return jsonify({'error': '데이터를 불러올 수 없습니다.'}), 500


@kis_bp.route('/kis/orderbook')
def kis_orderbook():
    symbol = request.args.get('symbol', '')
    token  = _get_token()
    if not token:
        return jsonify({'error': 'KIS 토큰 없음'}), 500
    try:
        resp = requests.get(
            f'{_base()}/uapi/domestic-stock/v1/quotations/inquire-asking-price-exp-ccn',
            headers=_headers('FHKST01010200', token),
            params={'FID_COND_MRKT_DIV_CODE': 'J', 'FID_INPUT_ISCD': symbol},
            timeout=5
        )
        d = resp.json().get('output1', {})
        asks, bids = [], []
        for i in range(1, 11):
            asks.append({'price': int(d.get(f'askp{i}', 0)), 'qty': int(d.get(f'askp_rsqn{i}', 0))})
            bids.append({'price': int(d.get(f'bidp{i}', 0)), 'qty': int(d.get(f'bidp_rsqn{i}', 0))})
        return jsonify({'asks': asks, 'bids': bids})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@kis_bp.route('/kis/ws-key')
def kis_ws_key():
    token = _get_token()
    if not token:
        return jsonify({'error': 'KIS 토큰 없음'}), 500
    try:
        resp = requests.post(
            f'{_base()}/oauth2/Approval',
            json={'grant_type': 'client_credentials',
                  'appkey': os.environ.get('KIS_APP_KEY', ''),
                  'secretkey': os.environ.get('KIS_APP_SECRET', '')},
            timeout=5
        )
        return jsonify(resp.json())
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@kis_bp.route('/kis/investor')
def kis_investor():
    symbol = request.args.get('symbol', '')
    period = request.args.get('period', 'day')
    token  = _get_token()
    if not token:
        return jsonify({'error': 'KIS 토큰 없음'}), 500
    try:
        period_map = {'day': '0', 'week': '1', 'month': '2'}
        resp = requests.get(
            f'{_base()}/uapi/domestic-stock/v1/quotations/inquire-investor',
            headers=_headers('FHKST01010900', token),
            params={
                'FID_COND_MRKT_DIV_CODE': 'J',
                'FID_INPUT_ISCD': symbol,
                'FID_PERIOD_DIV_CODE': period_map.get(period, '0'),
            },
            timeout=5
        )
        return jsonify(resp.json().get('output', []))
    except Exception as e:
        return jsonify({'error': str(e)}), 500
