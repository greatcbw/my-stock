# -*- coding: utf-8 -*-
# routes/forecast.py — ARIMA/Prophet 예측, 포트폴리오, 날짜별 메모

from flask import Blueprint, request, jsonify
from datetime import datetime, timedelta
import os, json, glob

forecast_bp = Blueprint('forecast', __name__)

_shared = {}

def init_forecast(shared):
    global _shared
    _shared = shared

def _fdr(): return _shared.get('fdr')
def _FDR_OK(): return _shared.get('FDR_OK', False)
def _ARIMA_OK(): return _shared.get('ARIMA_OK', False)
def _PROPHET_OK(): return _shared.get('PROPHET_OK', False)
def _base_path(): return _shared.get('get_base_path', lambda: '.')()


# ── ARIMA + Prophet 예측 ──
@forecast_bp.route('/api/forecast/<symbol>')
def get_forecast(symbol):
    if not _ARIMA_OK():
        return jsonify({'error': 'statsmodels 미설치'}), 500
    try:
        import pandas as pd
        import yfinance as yf
        from statsmodels.tsa.arima.model import ARIMA

        fdr    = _fdr()
        FDR_OK = _FDR_OK()
        VALID  = 10
        FUTURE = 10
        is_krx = bool(__import__('re').match(r'^[0-9]{6}$', symbol))

        end   = datetime.now()
        start = end - timedelta(days=250)

        df = None
        if is_krx and FDR_OK:
            try: df = fdr.DataReader(symbol, start, end)
            except Exception: df = None
        if df is None or (hasattr(df, 'empty') and df.empty):
            yf_sym = symbol + '.KS' if is_krx else symbol
            df = yf.download(yf_sym, start=start, end=end, progress=False)
        if df is None or df.empty:
            return jsonify({'error': '데이터 없음'}), 404

        if isinstance(df.columns, pd.MultiIndex):
            close = df['Close'].iloc[:, 0]
        else:
            close = df['Close']
        close.index = pd.to_datetime(close.index).tz_localize(None)
        series = close.resample('B').ffill().dropna()

        if len(series) < 30 + VALID:
            return jsonify({'error': '데이터 부족'}), 400

        # A: 과거 검증
        train = series.iloc[:-VALID]
        val   = series.iloc[-VALID:]
        v_fit = ARIMA(train, order=(5,1,0)).fit()
        v_pred = v_fit.forecast(steps=VALID)
        validation = [
            {'time': d.strftime('%Y-%m-%d'),
             'actual': round(float(a),2),
             'predicted': round(float(p),2),
             'error': round(float(a)-float(p),2)}
            for (d,a),(_, p) in zip(val.items(), v_pred.items())
        ]

        # B: 미래 예측 (ARIMA)
        f_fit    = ARIMA(series, order=(5,1,0)).fit()
        forecast = f_fit.forecast(steps=FUTURE)
        forecast_list = [
            {'time': d.strftime('%Y-%m-%d'), 'value': round(float(v),2)}
            for d, v in forecast.items()
        ]

        # C: Prophet + 앙상블
        prophet_list, ensemble_list = [], []
        if _PROPHET_OK():
            try:
                from prophet import Prophet
                df_p = pd.DataFrame({'ds': series.index, 'y': series.values}).reset_index(drop=True)
                m = Prophet(daily_seasonality=False, weekly_seasonality=True,
                            yearly_seasonality=True, changepoint_prior_scale=0.05)
                m.fit(df_p)
                future = m.make_future_dataframe(periods=FUTURE, freq='B')
                pf     = m.predict(future).tail(FUTURE)
                prophet_list = [
                    {'time': r['ds'].strftime('%Y-%m-%d'), 'value': round(float(r['yhat']),2)}
                    for _, r in pf.iterrows()
                ]
                ensemble_list = [
                    {'time': a['time'], 'value': round((a['value']+p['value'])/2, 2)}
                    for a, p in zip(forecast_list, prophet_list)
                ]
            except Exception as pe:
                print(f'[Prophet] {symbol}: {pe}')

        return jsonify({
            'forecast':   forecast_list,
            'validation': validation,
            'prophet':    prophet_list,
            'ensemble':   ensemble_list,
        })
    except Exception as e:
        print(f'[ARIMA] {symbol}: {e}')
        return jsonify({'error': str(e)}), 500


# ── 포트폴리오 ──
@forecast_bp.route('/portfolio', methods=['GET'])
def get_portfolio():
    path = os.path.join(_base_path(), 'portfolio.json')
    try:
        if os.path.exists(path):
            with open(path, 'r', encoding='utf-8') as f:
                return jsonify(json.load(f))
        return jsonify({'myList': [], 'watchList': [], 'holdings': {}})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@forecast_bp.route('/portfolio', methods=['POST'])
def save_portfolio():
    path = os.path.join(_base_path(), 'portfolio.json')
    try:
        data = request.get_json()
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return jsonify({'ok': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ── 날짜별 메모 ──
def _memo_dir():
    d = os.path.join(_base_path(), 'memos')
    os.makedirs(d, exist_ok=True)
    return d


@forecast_bp.route('/memo/list')
def list_memos():
    files = sorted(glob.glob(os.path.join(_memo_dir(), '*.txt')), reverse=True)
    return jsonify([os.path.basename(f).replace('.txt','') for f in files])


@forecast_bp.route('/memo/<date>', methods=['GET'])
def get_memo(date):
    path = os.path.join(_memo_dir(), f'{date}.txt')
    content = open(path, encoding='utf-8').read() if os.path.exists(path) else ''
    return jsonify({'date': date, 'content': content})


@forecast_bp.route('/memo/<date>', methods=['POST'])
def save_memo(date):
    path = os.path.join(_memo_dir(), f'{date}.txt')
    content = request.get_json().get('content', '')
    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
    return jsonify({'ok': True})
