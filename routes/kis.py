# -*- coding: utf-8 -*-
# routes/kis.py — 한국투자증권 KIS API

from flask import Blueprint, request, jsonify
import requests
from datetime import datetime
import os

from extensions import cache

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
# @cache.cached(timeout=5, query_string=True)
def kis_price():
    symbol = request.args.get('symbol', '')
    if not symbol:
        return jsonify({'error': '종목코드 필요'}), 400
    token = _get_token()
    print(f"[DEBUG] symbol={symbol} token={'있음' if token else '없음'}")

    if token:
        try:
            for market_code in ['J', 'Q']:
                resp = requests.get(
                    f'{_base()}/uapi/domestic-stock/v1/quotations/inquire-price',
                    headers=_headers('FHKST01010100', token),
                    params={'FID_COND_MRKT_DIV_CODE': market_code, 'FID_INPUT_ISCD': symbol},
                    timeout=3
                )
                data = resp.json().get('output')
                if data:
                    print(f"[DEBUG] KIS 응답: {symbol} market={market_code} price={data.get('stck_prpr')} per={data.get('per')}")
                    if int(data.get('stck_prpr', 0)) > 0:
                        return jsonify({
                            'price': int(data.get('stck_prpr', 0)),
                            'change': int(data.get('prdy_vrss', 0)),
                            'change_pct': float(data.get('prdy_ctrt', 0)),
                            'volume': int(data.get('acml_vol', 0)),
                            'mktcap': int(data.get('hts_avls', 0)),
                            'open':  int(data.get('stck_oprc', 0)),
                            'high':  int(data.get('stck_hgpr', 0)),
                            'low':   int(data.get('stck_lwpr', 0)),
                            'per': float(data.get('per', 0)) or None,
                            'pbr': float(data.get('pbr', 0)) or None,
                            'eps': float(data.get('eps', 0)) or None,
                            'time': datetime.now().strftime('%H:%M:%S'),
                            'is_fallback': False
                        })
        except Exception as e:
            print(f"⚠ KIS API 오류: {e}")

    try:
        import yfinance as yf
        import pandas as pd
        print(f"[DEBUG] Fallback 시작: {symbol}")
        yf_symbol = f"{symbol}.KS"
        tk = yf.Ticker(yf_symbol)
        if not tk.fast_info.last_price:
            yf_symbol = f"{symbol}.KQ"
            tk = yf.Ticker(yf_symbol)
        info = tk.fast_info
        hist = tk.history(period='2d')
        price = float(info.last_price or 0)
        prev_close = float(hist['Close'].iloc[-2]) if len(hist) >= 2 else price
        change = price - prev_close
        change_pct = (change / prev_close * 100) if prev_close else 0
        tk_info = tk.info
        return jsonify({
            'price': int(price),
            'change': int(change),
            'change_pct': round(change_pct, 2),
            'volume': int(info.last_volume or 0),
            'per': round(tk_info.get('trailingPE', 0) or 0, 2) or None,
            'pbr': round(tk_info.get('priceToBook', 0) or 0, 2) or None,
            'eps': round(tk_info.get('epsTrailingTwelveMonths', 0) or 0, 2) or None,
            'is_fallback': True,
            'msg': '⚠️ KIS 점검중 (야후 15분 지연 데이터)'
        })
    except Exception as fallback_e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(fallback_e)}), 500

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
