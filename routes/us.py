# -*- coding: utf-8 -*-
# routes/us.py — 미국주식 실시간 (yfinance)

from flask import Blueprint, request, jsonify
import yfinance as yf

us_bp = Blueprint('us', __name__)


@us_bp.route('/us/quote')
def us_quote():
    symbol = request.args.get('symbol', '')
    try:
        tk   = yf.Ticker(symbol)
        info = tk.fast_info
        hist = tk.history(period='2d', interval='1d')
        price      = float(info.last_price or 0)
        prev_close = float(hist['Close'].iloc[-2]) if len(hist) >= 2 else price
        change     = round(price - prev_close, 4)
        change_pct = round((change / prev_close * 100) if prev_close else 0, 4)
        return jsonify({
            'price':      round(price, 4),
            'change':     change,
            'change_pct': change_pct,
            'volume':     int(info.three_month_average_volume or 0),
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@us_bp.route('/us/trades')
def us_trades():
    symbol = request.args.get('symbol', '')
    try:
        tk   = yf.Ticker(symbol)
        hist = tk.history(period='1d', interval='1m')
        if hist.empty:
            return jsonify([])
        result = []
        for dt, row in hist.tail(50).iterrows():
            result.append({
                'time':   int(dt.timestamp()),
                'price':  round(float(row['Close']), 4),
                'volume': int(row['Volume']),
            })
        return jsonify(result)
    except Exception as e:
        return jsonify({'error': str(e)}), 500
