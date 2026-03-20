# -*- coding: utf-8 -*-
# routes/data.py — 주가 데이터, 지수, 뉴스, 펀더멘털

from flask import Blueprint, request, jsonify
from datetime import datetime, timedelta
import requests
import yfinance as yf
import pandas as pd

from extensions import cache # 캐시 불러오기 추가

data_bp = Blueprint('data', __name__)

# 공유 변수는 app.py에서 주입
_shared = {}

def init_data(shared):
    """app.py에서 공유 변수 주입"""
    global _shared
    _shared = shared

def _fdr(): return _shared.get('fdr')
def _FDR_OK(): return _shared.get('FDR_OK', False)
def _krx_list(): return _shared.get('krx_list', [])
def _load_csvs(): return _shared.get('load_csvs_if_changed', lambda: None)()

# ── 종목 검색 ──
# data.py 하단에 추가
@data_bp.route('/sector/status', methods=['POST'])
def sector_status():
    try:
        symbols = request.get_json().get('symbols', [])
        if not symbols: return jsonify([])
        
        results = []
        for sym in symbols:
            # 기존 kis_price 로직 등을 재활용하여 현재가/등락률 수집
            # 여기서는 성능을 위해 yfinance fast_info를 활용한 예시입니다.
            tk = yf.Ticker(sym if not sym.isdigit() else f"{sym}.KS")
            hist = tk.history(period='2d')
            if len(hist) < 2: continue
            
            price = hist['Close'].iloc[-1]
            prev = hist['Close'].iloc[-2]
            chg_pct = ((price - prev) / prev) * 100
            results.append(chg_pct)
            
        avg_chg = sum(results) / len(results) if results else 0
        return jsonify({'avg_chg': round(avg_chg, 2)})
    except:
        return jsonify({'avg_chg': 0})


# ── OHLCV 데이터 ──
@data_bp.route('/data')
@cache.cached(timeout=60, query_string=True) # 🌟 60초 캐싱 적용!
def get_data():
    symbol   = request.args.get('symbol', '')
    interval = request.args.get('interval', 'day')
    market   = request.args.get('market', '')

    fdr     = _fdr()
    FDR_OK  = _FDR_OK()
    is_krx  = bool(__import__('re').match(r'^[0-9]{6}$', symbol))

    try:
        if interval == 'day':
            end   = datetime.now()
            start = end - timedelta(days=700)
            df = None
            if is_krx and FDR_OK:
                try: df = fdr.DataReader(symbol, start, end)
                except Exception: df = None
            if df is None or (hasattr(df, 'empty') and df.empty):
                yf_sym = symbol + '.KS' if is_krx else symbol
                df = yf.download(yf_sym, start=start, end=end, progress=False)
            if df is None or df.empty:
                return jsonify([])
            if isinstance(df.columns, pd.MultiIndex):
                df.columns = df.columns.get_level_values(0)
            df = df.dropna(subset=['Close'])
            result = []
            for dt, row in df.iterrows():
                result.append({
                    'time':   dt.strftime('%Y-%m-%d'),
                    'open':   round(float(row.get('Open', 0)), 2),
                    'high':   round(float(row.get('High', 0)), 2),
                    'low':    round(float(row.get('Low', 0)), 2),
                    'close':  round(float(row.get('Close', 0)), 2),
                    'volume': int(row.get('Volume', 0)),
                })
            return jsonify(result)
        else:
            # 분봉
            period_map = {'1m': '1d', '5m': '5d', '15m': '5d', '30m': '1mo'}
            period = period_map.get(interval, '1d')
            yf_sym = symbol + '.KS' if is_krx else symbol
            df = yf.download(yf_sym, period=period, interval=interval, progress=False)
            if df is None or df.empty:
                return jsonify([])
            if isinstance(df.columns, pd.MultiIndex):
                df.columns = df.columns.get_level_values(0)
            result = []
            for dt, row in df.iterrows():
                ts = int(dt.timestamp())
                result.append({
                    'time':   ts,
                    'open':   round(float(row.get('Open', 0)), 2),
                    'high':   round(float(row.get('High', 0)), 2),
                    'low':    round(float(row.get('Low', 0)), 2),
                    'close':  round(float(row.get('Close', 0)), 2),
                    'volume': int(row.get('Volume', 0)),
                })
            return jsonify(result)
    except Exception as e:
        print(f'[데이터 오류] {symbol}: {e}')
        return jsonify([])


# ── 주요 지수 ──
_indices_cache = {}

def _fetch_one_index(name, sym):
    import time
    cached = _indices_cache.get(name)
    if cached and time.time() - cached['ts'] < 25:
        return cached
    try:
        url  = f'https://query1.finance.yahoo.com/v8/finance/chart/{sym}?interval=1d&range=5d'
        resp = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=5)
        j    = resp.json()
        meta = j['chart']['result'][0]['meta']
        price     = float(meta.get('regularMarketPrice', 0))
        prev      = float(meta.get('previousClose', price))
        change    = round(price - prev, 4)
        change_pct= round((change / prev * 100) if prev else 0, 4)
        entry = {'name': name, 'price': price, 'change': change, 'change_pct': change_pct, 'ts': time.time()}
        _indices_cache[name] = entry
        return entry
    except Exception:
        return _indices_cache.get(name)


@data_bp.route('/indices')
def get_indices():
    from concurrent.futures import ThreadPoolExecutor, as_completed
    symbols = {
        'KOSPI':    '%5EKS11',
        'KOSDAQ':   '%5EKQ11',
        'S&P500':   '%5EGSPC',
        'NASDAQ':   '%5EIXIC',
        'DOW':      '%5EDJI',
        'USD/KRW':  'USDKRW%3DX',
        'WTI유가':  'CL%3DF',
        '금리(10Y)':'%5ETNX',
        '공포지수': '%5EVIX',
    }
    order = ['KOSPI','KOSDAQ','S&P500','NASDAQ','DOW','USD/KRW','WTI유가','금리(10Y)','공포지수']
    result = []
    with ThreadPoolExecutor(max_workers=9) as ex:
        futures = {ex.submit(_fetch_one_index, name, sym): name for name, sym in symbols.items()}
        for fut in futures:
            entry = fut.result()
            if entry: result.append(entry)
    result.sort(key=lambda x: order.index(x['name']) if x['name'] in order else 99)
    return jsonify(result)


# ── 뉴스 ──
@data_bp.route('/news')
def get_news():
    import xml.etree.ElementTree as ET
    feeds = [
        ('연합경제', 'https://www.yonhapnewstv.co.kr/category/news/economy/feed/'),
        ('연합증권', 'https://www.yna.co.kr/rss/economy.xml'),
        ('한국경제', 'https://www.hankyung.com/feed/economy'),
        ('매일경제', 'https://www.mk.co.kr/rss/30000001/'),
        ('전자신문', 'https://www.etnews.com/rss/allArticleList.xml'),
    ]
    items = []
    for src, url in feeds:
        try:
            r = requests.get(url, timeout=5, headers={'User-Agent': 'Mozilla/5.0'})
            root = ET.fromstring(r.content)
            for item in root.iter('item'):
                title = (item.findtext('title') or '').strip()
                link  = (item.findtext('link')  or '').strip()
                pub   = (item.findtext('pubDate') or '').strip()
                if title and link:
                    items.append({'source': src, 'title': title, 'link': link, 'pub': pub})
        except Exception:
            pass
    seen = set()
    unique = []
    for it in items:
        if it['title'] not in seen:
            seen.add(it['title'])
            unique.append(it)
    return jsonify(unique[:60])


# ── 펀더멘털 ──
@data_bp.route('/fundamentals')
def get_fundamentals():
    symbol = request.args.get('symbol', '')
    market = request.args.get('market', '')
    is_krx = bool(__import__('re').match(r'^[0-9]{6}$', symbol))
    yf_sym = symbol + '.KS' if is_krx else symbol
    try:
        tk = yf.Ticker(yf_sym)
        i  = tk.info
        def safe(k): return i.get(k)
        def pct(v):  return f'{round(v*100,2)}%' if v is not None else None
        def r2(v):   return round(v,2) if v is not None else None
        cap = safe('marketCap')
        def fmt_cap(v):
            if v is None: return None
            if v >= 1e12: return f'{v/1e12:.2f}조'
            if v >= 1e8:  return f'{v/1e8:.0f}억'
            return str(v)
        revenue = safe('totalRevenue')
        def fmt_rev(v):
            if v is None: return None
            if v >= 1e12: return f'{v/1e12:.2f}조'
            if v >= 1e8:  return f'{v/1e8:.0f}억'
            return str(v)
        return jsonify({
            'mktcap': fmt_cap(cap),
            'per': r2(safe('trailingPE')), 'per_fwd': r2(safe('forwardPE')),
            'pbr': r2(safe('priceToBook')), 'psr': r2(safe('priceToSalesTrailing12Months')),
            'div_yield': pct(safe('dividendYield')),
            'eps': r2(safe('epsTrailingTwelveMonths')), 'eps_fwd': r2(safe('epsForward')),
            'revenue': fmt_rev(revenue), 'rev_growth': pct(safe('revenueGrowth')),
            'earn_growth': pct(safe('earningsGrowth')),
            'oper_margin': pct(safe('operatingMargins')),
            'profit_margin': pct(safe('profitMargins')),
            'roe': pct(safe('returnOnEquity')), 'roa': pct(safe('returnOnAssets')),
            'fcf': fmt_rev(safe('freeCashflow')), 'ocf': fmt_rev(safe('operatingCashflow')),
            '52w_high': r2(safe('fiftyTwoWeekHigh')), '52w_low': r2(safe('fiftyTwoWeekLow')),
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500
