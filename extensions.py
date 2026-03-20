# extensions.py
from flask_caching import Cache

# 메모리 기반 캐시 생성 (기본 60초 동안 데이터 유지)
cache = Cache(config={
    'CACHE_TYPE': 'SimpleCache',
    'CACHE_DEFAULT_TIMEOUT': 60
})