// chart.js — 차트/팝업/지표/크로스알림 (자동 분리)



/* ══════════════════════════════════
   엑셀 출력
══════════════════════════════════ */
async function exportChart(cid, key) {
  const w = activeWidgets[cid];
  if (!w) { showToast('⚠ 차트 데이터 없음'); return; }

  // 데이터 소스: 모달이면 모달 데이터, 팝업이면 w.data
  const modal    = key ? chartModals[key] : null;
  const data     = (modal && modal._modalData) ? modal._modalData : w.data;
  const interval = (modal && modal._modalInterval) ? modal._modalInterval : (w.interval || 'day');
  
  

  if (!data || !data.length) { showToast('⚠ 차트 데이터 없음'); return; }
  showToast('📊 엑셀 파일 생성 중...');

  try {
    const wb   = XLSX.utils.book_new();
    const name = w.name || w.symbol;

    // ── OHLCV + 이평선 ──
    const ma20m  = new Map(calcMA(data,20).map(d  => [d.time, d.value]));
    const ma60m  = new Map(calcMA(data,60).map(d  => [d.time, d.value]));
    const ma120m = new Map(calcMA(data,120).map(d => [d.time, d.value]));
    const bb     = calcBB(data, 20);
    const bbUm   = new Map(bb.upper.map(d => [d.time, d.value]));
    const bbLm   = new Map(bb.lower.map(d => [d.time, d.value]));

    const rows1 = [['날짜/시간','시가','고가','저가','종가','거래량','20MA','60MA','120MA','BB상단','BB하단']];
    data.forEach(d => {
      const t = typeof d.time === 'number'
        ? new Date(d.time*1000).toISOString().slice(0,16).replace('T',' ')
        : String(d.time);
      rows1.push([t, d.open, d.high, d.low, d.close, d.volume,
        ma20m.get(d.time)  ? +ma20m.get(d.time).toFixed(2)  : '',
        ma60m.get(d.time)  ? +ma60m.get(d.time).toFixed(2)  : '',
        ma120m.get(d.time) ? +ma120m.get(d.time).toFixed(2) : '',
        bbUm.get(d.time)   ? +bbUm.get(d.time).toFixed(2)   : '',
        bbLm.get(d.time)   ? +bbLm.get(d.time).toFixed(2)   : '',
      ]);
    });

    // 일봉이면 ARIMA 예측 포함 (시트 순서: 미래예측 → 과거검증 → OHLCV)
    if (interval === 'day') {
      try {
        const res   = await fetch(`${API}/api/forecast/${w.symbol}`);
        const fdata = await res.json();
        if (!fdata.error) {
          const propMap = new Map((fdata.prophet||[]).map(d=>[d.time,d.value]));
          const ensMap  = new Map((fdata.ensemble||[]).map(d=>[d.time,d.value]));
          const todayPrice = data[data.length-1]?.close || 0;

          const rows2 = [['날짜','ARIMA','Prophet','앙상블','현재가(오늘)','앙상블-현재가','등락률(%)']];
          (fdata.forecast || []).forEach((d, i) => {
            const ens  = ensMap.get(d.time) || '';
            const diff = todayPrice && ens ? +((ens - todayPrice).toFixed(2)) : '';
            const pct  = todayPrice && ens ? +((((ens - todayPrice) / todayPrice) * 100).toFixed(2)) : '';
            rows2.push([d.time, d.value, propMap.get(d.time)||'', ens, i===0 ? todayPrice : '', diff, pct]);
          });
          XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows2), '미래예측(B)');

          const rows3 = [['날짜','실제값','검증예측','오차','오차율(%)']];
          (fdata.validation||[]).forEach(d => rows3.push([d.time, d.actual, d.predicted, d.error,
            d.actual ? +((d.error/d.actual)*100).toFixed(2) : '']));
          XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows3), '과거검증(A)');
        }
      } catch(e) {}
    }

    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows1), interval === 'day' ? 'OHLCV+이평선' : `OHLCV(${interval})`);

    const date = new Date().toISOString().slice(0,10);
    XLSX.writeFile(wb, `${name}_${interval}_${date}.xlsx`);
    showToast(`✅ ${name}_${interval}_${date}.xlsx 저장됨`);

  } catch(e) {
    console.error('엑셀 출력 오류:', e);
    showToast('❌ 엑셀 출력 실패');
  }
}

/* ══════════════════════════════════
   전역 상태
══════════════════════════════════ */
const API = 'http://127.0.0.1:5000';

// ── 전역 에러 캐처 (디버그용) ──
window.onerror = function(msg, src, line, col, err) {
  const div = document.createElement('div');
  div.style.cssText = 'position:fixed;top:50px;left:50%;transform:translateX(-50%);background:#ef4444;color:#fff;padding:12px 20px;border-radius:8px;z-index:99999;font-size:13px;max-width:700px;word-break:break-all;';
  div.textContent = '🚨 JS 에러: ' + msg + ' (L' + line + ':' + col + ')';
  document.body.appendChild(div);
  setTimeout(() => div.remove(), 15000);
  return false;
};
window.addEventListener('unhandledrejection', e => {
  const div = document.createElement('div');
  div.style.cssText = 'position:fixed;top:110px;left:50%;transform:translateX(-50%);background:#f59e0b;color:#fff;padding:12px 20px;border-radius:8px;z-index:99999;font-size:13px;max-width:700px;word-break:break-all;';
  div.textContent = '⚠️ Promise 에러: ' + (e.reason?.message || e.reason || '알 수 없음');
  document.body.appendChild(div);
  setTimeout(() => div.remove(), 15000);
});
let activeWidgets = {};   // cid → widget 객체
let curSym=null, curMkt=null, curName=null;
let isDark   = localStorage.getItem('isDark')    === 'true';
let isUS     = localStorage.getItem('isUSStyle') === 'true';
let soundOn  = false;
let zBase    = 600;

// 지수 갱신 관련 변수 — startIndexRefresh() 호출 전에 선언 필수
const REFRESH_SEC = 30;
let cdTimer = null, rfTimer = null, lastPrices = {}, cdRemain = REFRESH_SEC;
let _wcPriceTimer = null;
let currentFilter = localStorage.getItem('dashFilter') || 'all';
let currentSort = localStorage.getItem('dashSort') || 'manual';
const FILTER_LABELS = {
  all:'전체보기', my:'★ 내종목', watch:'⭐ 관심종목', kr:'🇰🇷 한국주식', us:'🇺🇸 미국주식'
};

if (isDark) document.body.classList.add('dark-theme');
document.getElementById('btn-theme').innerText = isDark ? '☀️' : '🌙';
document.getElementById('btn-style').innerText = isUS   ? '🇺🇸' : '🇰🇷';

// 검색창 외부 클릭 → 결과 닫기
document.addEventListener('click', e => {
  if (!document.getElementById('search').contains(e.target) &&
      !document.getElementById('results').contains(e.target))
    document.getElementById('results').innerHTML = '';
});

// 앱 시작: portfolio.json 있으면 복원 후 loadWatch
(async () => {
  const restored = await loadPortfolio();
  loadWatch();
  startIndexRefresh();
})();
setInterval(() => {
  Object.keys(activeWidgets).forEach(id => {
    const w = activeWidgets[id];
    if (w && w.interval !== 'day' && !w.minimized)
      fetchChart(id, w.symbol, w.interval, false);
  });
}, 60000);

/* ══════════════════════════════════
   지수 로드
══════════════════════════════════ */

function startIndexRefresh() {
  if (rfTimer)  clearInterval(rfTimer);
  if (cdTimer)  clearInterval(cdTimer);
  cdRemain = REFRESH_SEC;
  fetchNews();
  setInterval(fetchNews, 5 * 60 * 1000);
  loadIndices();
  rfTimer = setInterval(() => { cdRemain = REFRESH_SEC; loadIndices(); }, REFRESH_SEC * 1000);
  cdTimer = setInterval(() => {
    cdRemain = Math.max(0, cdRemain - 1);
    const el = document.getElementById('idx-countdown');
    if (el) el.textContent = `(${cdRemain}s)`;
  }, 1000);
}

/* ══════════════════════════════════
   뉴스 패널
══════════════════════════════════ */
/* 뉴스 패널 리사이즈 */
(function() {
  let startY = 0, startH = 0, dragging = false;
  const SAVED_KEY = 'newsPanelH';
  const MIN_H = 26, MAX_H = 600;

  function init() {
    const resizer  = document.getElementById('news-resizer');
    const newsList = document.getElementById('news-list');
    if (!resizer || !newsList) return;

    // 저장된 높이 복원
    const saved = localStorage.getItem(SAVED_KEY);
    if (saved) newsList.style.height = saved + 'px';

    resizer.addEventListener('mousedown', function(e) {
      dragging = true;
      startY = e.clientY;
      startH = newsList.offsetHeight;
      resizer.classList.add('dragging');
      document.body.style.userSelect = 'none';
      document.body.style.cursor = 'ns-resize';
      e.preventDefault();
    });
    document.addEventListener('mousemove', function(e) {
      if (!dragging) return;
      const newH = Math.min(MAX_H, Math.max(MIN_H, startH + (e.clientY - startY)));
      newsList.style.height = newH + 'px';
    });
    document.addEventListener('mouseup', function() {
      if (!dragging) return;
      dragging = false;
      resizer.classList.remove('dragging');
      document.body.style.userSelect = '';
      document.body.style.cursor = '';
      localStorage.setItem(SAVED_KEY, parseInt(newsList.style.height));
    });
  }
  // DOM 준비 후 실행
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();

async function fetchNews() {
  const list   = document.getElementById('news-list');
  const update = document.getElementById('news-update');
  if (!list) return;
  list.innerHTML = '<div id="news-loading">뉴스 불러오는 중...</div>';
  try {
    const items = await fetch(`${API}/news`).then(r => r.json());
    if (!items || !items.length) {
      list.innerHTML = '<div id="news-loading">뉴스를 가져올 수 없습니다.</div>';
      return;
    }
    list.innerHTML = items.map(n => {
      const cat = n.category || '뉴스';
      return `<a class="news-item" href="${n.link}" target="_blank" rel="noopener">
        <span class="news-cat ${cat}">${cat}</span>
        <span class="news-txt">${n.title}</span>
      </a>`;
    }).join('');
    if (update) update.textContent = '업데이트 ' + new Date().toLocaleTimeString('ko-KR');
  } catch(e) {
    list.innerHTML = '<div id="news-loading">⚠ 뉴스 로드 실패</div>';
  }
}

async function loadIndices() {
  const dot = document.getElementById('idx-live-dot');
  const txt = document.getElementById('idx-update-txt');
  if (dot) dot.style.background = '#f59e0b';
  try {
    const data = await fetch(`${API}/indices`).then(r => r.json());
    const nm   = {'KOSPI':0,'KOSDAQ':1,'S&P500':2,'NASDAQ':3,'DOW':4,'USD/KRW':5,'WTI유가':6,'금리(10Y)':7,'공포지수':8};
    const cards = document.querySelectorAll('#idx-grid .idx-card');
    data.forEach(d => {
      const i = nm[d.name]; if (i === undefined) return;
      const c  = cards[i];
      const kr = c.dataset.kr === '1';
      const up = d.change_pct >= 0;
      const cls = kr ? (up?'kr-up':'kr-dn') : (up?'us-up':'us-dn');
      const sg  = up ? '+' : '';
      const ar  = up ? '▲' : '▼';
      const pr  = d.name === 'USD/KRW'
        ? d.price.toLocaleString('ko-KR',{maximumFractionDigits:2}) + ' ₩'
        : d.price.toLocaleString('ko-KR',{maximumFractionDigits:2});
      const prev = lastPrices[d.name];
      const flashCls = prev == null ? '' : d.price > prev ? 'flash-up' : d.price < prev ? 'flash-dn' : '';
      lastPrices[d.name] = d.price;
      if (d.name === 'USD/KRW') _usdkrw = d.price;
      c.innerHTML = `
        <div class="idx-name">${d.name}</div>
        <div class="idx-price ${flashCls}">${pr}</div>
        <div class="idx-chg ${cls}">${ar} ${sg}${d.change.toLocaleString('ko-KR',{maximumFractionDigits:2})} (${sg}${d.change_pct.toFixed(2)}%)</div>`;
      if (flashCls) setTimeout(() => { const p=c.querySelector('.idx-price'); if(p) p.classList.remove(flashCls); }, 800);
    });
    const now = new Date();
    const ts  = now.toLocaleTimeString('ko-KR',{hour:'2-digit',minute:'2-digit',second:'2-digit'});
    if (dot) dot.style.background = '#22c55e';
    if (txt) txt.textContent = `업데이트 ${ts}`;
    cdRemain = REFRESH_SEC;
    const cd = document.getElementById('idx-countdown');
    if (cd) cd.textContent = `(${REFRESH_SEC}s)`;
  } catch(e) {
    console.warn('지수 로드 실패', e);
    if (dot) dot.style.background = '#ef4444';
    if (txt) txt.textContent = '연결 오류';
  }
}

/* ══════════════════════════════════
   내종목 + 관심종목
══════════════════════════════════ */
let dragFrom = null;
function getList()    { return JSON.parse(localStorage.getItem('watchPro')   || '[]'); }
function getMyList()  { return JSON.parse(localStorage.getItem('myStockPro') || '[]'); }
/* ══════════════════════════════════
   포트폴리오 자동 저장/로드 (portfolio.json)
══════════════════════════════════ */

// 전체 포트폴리오를 서버에 저장
async function savePortfolio() {
  try {
    const myList    = getMyList();
    const watchList = getList();
    // 보유정보: holding_* 키 전부 수집
    const holdings = {};
    myList.forEach(item => {
      const h = getHolding(item.symbol);
      if (h) holdings[item.symbol] = h;
    });
    await fetch(`${API}/portfolio`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ myList, watchList, holdings })
    });
  } catch(e) {
    console.warn('[포트폴리오 저장 오류]', e);
  }
}

function saveList(l)  { localStorage.setItem('watchPro',   JSON.stringify(l)); savePortfolio(); }
function saveMyList(l){ localStorage.setItem('myStockPro', JSON.stringify(l)); savePortfolio(); }

function addWatch() {
  if (!curSym || !curName) { showToast('⚠️ 검색 결과에서 종목을 먼저 클릭하세요.'); return; }
  const list = getList();
  if (list.find(x => x.symbol === curSym)) { showToast(`ℹ️ [${curName}] 이미 추가됐습니다.`); }
  else { list.push({name:curName, symbol:curSym, market:curMkt}); saveList(list); showToast(`✅ [${curName}] 관심종목 추가`); }
  loadWatch();
}
function addMyStock() {
  if (!curSym || !curName) { showToast('⚠️ 검색 결과에서 종목을 먼저 클릭하세요.'); return; }
  const list = getMyList();
  if (list.find(x => x.symbol === curSym)) { showToast(`ℹ️ [${curName}] 이미 내종목에 있습니다.`); }
  else { list.push({name:curName, symbol:curSym, market:curMkt}); saveMyList(list); showToast(`✅ [${curName}] 내종목 추가`); }
  loadWatch();
}
function removeWatch(sym, e) {
  e.stopPropagation();
  saveList(getList().filter(x => x.symbol !== sym));
  loadWatch();
}
function removeMyStock(sym, e) {
  e.stopPropagation();
  saveMyList(getMyList().filter(x => x.symbol !== sym));
  loadWatch();
}

function _renderSideCol(items, container, removeFn, listType) {
  container.innerHTML = '';
  items.forEach((item, i) => {
    const mc = mktCls(item.market);
    const el = document.createElement('div');
    el.className = 'watch-item'; el.dataset.index = i;
    el.innerHTML = `
      <span class="wi-drag">⠿</span>
      <div class="wi-info">
        <div class="wi-name">${item.name}</div>
        <div class="wi-sub"><span class="market-badge ${mc}">${item.market||'기타'}</span> ${item.symbol}</div>
      </div>
      <button class="wi-del">✕</button>`;
    el.querySelector('.wi-info').onclick = () => openPopup(item.name, item.symbol, item.market||'기타');
    el.querySelector('.wi-del').onclick  = e => removeFn(item.symbol, e);
    bindDrag(el, i, listType);
    container.appendChild(el);
  });
  bindColDrop(container, listType);
}

// 대시보드 카드 순서 저장
function getDashOrder() { return JSON.parse(localStorage.getItem('dashOrder') || '[]'); }
function saveDashOrder(l) { localStorage.setItem('dashOrder', JSON.stringify(l)); }

function syncDashOrder() {
  // 새로 추가된 종목은 뒤에 붙이고, 삭제된 건 제거
  const myList = getMyList(), wList = getList();
  const allSyms = [...myList, ...wList].map(x => x.symbol);
  const order = getDashOrder().filter(s => allSyms.includes(s));
  allSyms.forEach(s => { if (!order.includes(s)) order.push(s); });
  saveDashOrder(order);
  return order;
}


/* ── 정렬 상태 ── */
let _sortKey = localStorage.getItem('dashSortKey') || null; // 'invest'|'name'|'change'|null
let _sortDir = localStorage.getItem('dashSortDir') || 'desc'; // 'asc'|'desc'

function toggleSort(key) {
  if (_sortKey === key) {
    _sortDir = _sortDir === 'desc' ? 'asc' : 'desc';
  } else {
    _sortKey = key;
    _sortDir = 'desc';
  }
  localStorage.setItem('dashSortKey', _sortKey);
  localStorage.setItem('dashSortDir', _sortDir);
  updateSortBtnUI();
  sortCardsInDOM(); // ← DOM 재정렬 (빠름)
}
function getUsdKrw() {
  const el = document.querySelector('[id*="USDKRW"]');
  if (el?.dataset.price) return parseFloat(el.dataset.price);
  return 1500; // 환율 못 읽으면 기본값
}
function sortCardsInDOM() {
  const wc = document.getElementById('watch-cards');
  if (!wc) return;

  const cards = [...wc.querySelectorAll('.wc-card')];
  if (!cards.length) return;

  const usdKrw = getUsdKrw(); // 환율
  const dir = _sortDir === 'desc' ? -1 : 1;

  const items = cards.map(card => ({
    el:     card,
    symbol: card.dataset.symbol,
    name:   card.dataset.name || card.querySelector('.wc-name')?.textContent?.replace('★ ','') || ''
  }));

  items.sort((a, b) => {
    const isKrA = /^[0-9]{6}$/.test(a.symbol);
    const isKrB = /^[0-9]{6}$/.test(b.symbol);

    if (_sortKey === 'invest') {
      // 투자금액 = avg_price × qty (미국은 × 환율 → 원화 통일)
      const ha = JSON.parse(localStorage.getItem(`holding_${a.symbol}`) || 'null');
      const hb = JSON.parse(localStorage.getItem(`holding_${b.symbol}`) || 'null');
      const va = ha ? (ha.avg_price * ha.qty * (isKrA ? 1 : usdKrw)) : 0;
      const vb = hb ? (hb.avg_price * hb.qty * (isKrB ? 1 : usdKrw)) : 0;
      if (va === 0 && vb === 0) return 0;
      if (va === 0) return 1;
      if (vb === 0) return -1;
      return dir * (vb - va);

    } else if (_sortKey === 'eval') {
      // 평가금액 = 현재가 × qty (미국은 × 환율 → 원화 통일)
      const ha = JSON.parse(localStorage.getItem(`holding_${a.symbol}`) || 'null');
      const hb = JSON.parse(localStorage.getItem(`holding_${b.symbol}`) || 'null');
      const priceElA = document.getElementById(`wcp-${a.symbol}`);
      const priceElB = document.getElementById(`wcp-${b.symbol}`);
      const priceA = priceElA?.dataset.price ? parseFloat(priceElA.dataset.price) : 0;
      const priceB = priceElB?.dataset.price ? parseFloat(priceElB.dataset.price) : 0;
      const qtyA = ha ? ha.qty : 0;
      const qtyB = hb ? hb.qty : 0;
      const va = priceA * qtyA * (isKrA ? 1 : usdKrw);
      const vb = priceB * qtyB * (isKrB ? 1 : usdKrw);
      if (va === 0 && vb === 0) return 0;
      if (va === 0) return 1;
      if (vb === 0) return -1;
      return dir * (vb - va);

    } else if (_sortKey === 'name') {
      return dir * a.name.localeCompare(b.name, 'ko');

    } else if (_sortKey === 'change') {
      const ea = document.getElementById(`wcg-${a.symbol}`);
      const eb = document.getElementById(`wcg-${b.symbol}`);
      const va = ea?.dataset.changePct ? parseFloat(ea.dataset.changePct) : null;
      const vb = eb?.dataset.changePct ? parseFloat(eb.dataset.changePct) : null;
      if (va === null && vb === null) return 0;
      if (va === null) return 1;
      if (vb === null) return -1;
      return dir * (vb - va);
    }
    return 0;
  });

  // DOM 순서 재배치
  items.forEach(item => wc.appendChild(item.el));
}

function updateSortBtnUI() {
  const arrow = _sortDir === 'desc' ? ' ▼' : ' ▲';
  const labels = { invest: '💰 투자금액', eval: '📈 평가금액', name: '🔤 이름', change: '📊 등락률' };

  ['invest', 'name', 'change'].forEach(key => {
    const btn = document.getElementById(`sort-btn-${key}`);
    if (!btn) return;
    const isActive = _sortKey === key;
    btn.classList.toggle('active', isActive);
    btn.textContent = labels[key] + (isActive ? arrow : '');
  });
}

function getSortedItems(allItems) {
  if (!_sortKey) return allItems; // 정렬 없으면 수동순서 유지

	const dir = _sortDir === 'desc' ? -1 : 1;

	return [...allItems].sort((a, b) => {
	  if (_sortKey === 'invest') {
		const ha = JSON.parse(localStorage.getItem(`holding_${a.symbol}`) || 'null');
		const hb = JSON.parse(localStorage.getItem(`holding_${b.symbol}`) || 'null');
		const va = ha ? (ha.avg_price * ha.qty) : 0;
		const vb = hb ? (hb.avg_price * hb.qty) : 0;
	  // 보유정보 없는 종목은 항상 뒤로
	  if (va === 0 && vb === 0) return 0;
	  if (va === 0) return 1;
	  if (vb === 0) return -1;
	  return dir * (vb - va);
	} else if (_sortKey === 'name') {
	  return dir * a.name.localeCompare(b.name, 'ko');
	} else if (_sortKey === 'change') {
	  const ea = document.getElementById(`wcg-${a.symbol}`);
	  const eb = document.getElementById(`wcg-${b.symbol}`);
	  const va = ea?.dataset.changePct ? parseFloat(ea.dataset.changePct) : null; // ← changePct
	  const vb = eb?.dataset.changePct ? parseFloat(eb.dataset.changePct) : null; // ← changePct
	  // 가격 미로딩 종목은 항상 뒤로
	  if (va === null && vb === null) return 0;
	  if (va === null) return 1;
	  if (vb === null) return -1;
	  return dir * (vb - va);
	}
    return 0;
  });
}

let cardDragSym = null;

function bindCardDrag(card, symbol) {
  card.draggable = true;
  card.addEventListener('dragstart', e => {
    cardDragSym = symbol;
    card.classList.add('card-dragging');
    e.dataTransfer.effectAllowed = 'move';
  });
  card.addEventListener('dragend', () => {
    card.classList.remove('card-dragging');
    document.querySelectorAll('.wc-card').forEach(c => c.classList.remove('card-over'));
  });
  card.addEventListener('dragover', e => { e.preventDefault(); card.classList.add('card-over'); });
  card.addEventListener('dragleave', () => card.classList.remove('card-over'));
  card.addEventListener('drop', e => {
    e.preventDefault(); e.stopPropagation();
    card.classList.remove('card-over');
    if (!cardDragSym || cardDragSym === symbol) return;
    const order = getDashOrder();
    const fi = order.indexOf(cardDragSym);
    const ti = order.indexOf(symbol);
    if (fi < 0 || ti < 0) return;
    order.splice(fi, 1); order.splice(ti, 0, cardDragSym);
    saveDashOrder(order);
    cardDragSym = null;
    loadWatch();
  });
}

function loadWatch() {
  const wList  = getList();
  const myList = getMyList();

  /* 사이드바 2컬럼 */
  _renderSideCol(myList,  document.getElementById('my-list'),       removeMyStock, 'my');
  _renderSideCol(wList,   document.getElementById('watch-col-list'), removeWatch,  'watch');

  /* 대시보드 카드 — dashOrder 기준 정렬 */
  const wc = document.getElementById('watch-cards');
  const order = syncDashOrder();
  const allMap = {};
  [...myList.map(x=>({...x,_my:true})), ...wList.map(x=>({...x,_my:false}))]
    .forEach(x => allMap[x.symbol] = x);
  const allItems = order.map(s => allMap[s]).filter(Boolean);
  const sortedItems = getSortedItems(allItems); // ← 정렬 적용

  if (!sortedItems.length) {
    wc.innerHTML = '<div class="wc-empty"><p>👈 왼쪽에서 종목을 검색하고 추가하세요.<br>카드를 클릭하면 차트 팝업이 열립니다.</p></div>';
    return;
  }
  wc.innerHTML = '';
  if (_sectorViewOn) {
    renderSectorView(wc, sortedItems);   // ← sortedItems
  } else {
    sortedItems.forEach(item => wc.appendChild(buildCard(item)));  // ← sortedItems
  }
  const svBtn = document.getElementById('sector-view-btn');
  if (svBtn) svBtn.classList.toggle('on', _sectorViewOn);
  scheduleWatchPrices();
  applyFilter();
  loadHoldings();
}

/* 관심종목 카드 가격 갱신 */
async function loadWatchPrices() {
  const list = [...getMyList(), ...getList()];
  if (!list.length) return;

  await Promise.allSettled(list.map(async item => {
    const priceEl = document.getElementById(`wcp-${item.symbol}`);
    const chgEl   = document.getElementById(`wcg-${item.symbol}`);
    
    // 🌟 종목명이 들어있는 DOM 요소도 가져옵니다 (ID가 name-005930 형태라고 가정)
    // 만약 종목명 ID가 다르다면 그에 맞게 수정해 주세요!
    const nameEl  = document.getElementById(`name-${item.symbol}`); 
    
    if (!priceEl || !chgEl) return;

    try {
      const isKrx = /^[0-9]{6}$/.test(item.symbol);
      let price = 0, change = 0, changePct = 0, loaded = false;

      if (isKrx) {
        // 1차: KIS API
        try {
          const d = await fetch(`${API}/kis/price?symbol=${item.symbol}`).then(r => r.json());
          if (!d.error && d.price) {
            price = d.price; change = d.change; changePct = d.change_pct; loaded = true;
            
            // 🌟 [추가] 1차 KIS 성공 시: 혹시 있던 지연(⚠️) 마크를 지워줍니다.
            if (nameEl) {
              nameEl.innerHTML = nameEl.innerHTML.replace(/ <span.*⚠️지연<\/span>/g, '');
            }
          }
        } catch(e) {}

        // 2차 폴백: yfinance (.KS)
        if (!loaded) {
          try {
            const d = await fetch(`${API}/us/quote?symbol=${item.symbol}.KS`).then(r => r.json());
            if (!d.error && d.price) {
              price = d.price; change = d.change; changePct = d.change_pct; loaded = true;
              
              // 🌟 [추가] 2차 야후 폴백 성공 시: 종목명 옆에 지연(⚠️) 마크를 달아줍니다.
              if (nameEl && !nameEl.innerHTML.includes('⚠️')) {
                nameEl.innerHTML += ' <span style="font-size:10px; color:#f59e0b;" title="KIS 점검중 (야후 15분 지연)">⚠️지연</span>';
              }
            }
          } catch(e) {}
        }
      } else {
        // 미국 주식 등 기타 로직 (그대로 유지)
        try {
          const d = await fetch(`${API}/us/quote?symbol=${item.symbol}`).then(r => r.json());
          if (!d.error && d.price) {
            price = d.price; change = d.change; changePct = d.change_pct; loaded = true;
          }
        } catch(e) {}
      }
      
      // ... 이 아래에는 priceEl과 chgEl에 가격을 찍어주는 기존 코드가 있을 것입니다 ...

      if (!loaded) { priceEl.textContent = '-'; priceEl.dataset.price = ''; chgEl.textContent = '-'; chgEl.dataset.change = ''; return; }

      const isUp   = change > 0;
      const isDown = change < 0;
      const cls    = isUp ? 'up' : isDown ? 'down' : 'flat';
      const sign   = isUp ? '+' : '';
      const arrow  = isUp ? '▲' : isDown ? '▼' : '━';

      // 표시용 텍스트: 한국=쉼표 정수+'원', 미국=소수 2자리 고정
      const priceDisplay = isKrx
        ? fmt(price) + '원'
        : '$' + price.toFixed(2);
      priceEl.textContent  = priceDisplay;
      priceEl.dataset.price = price;   // 파싱 없이 숫자 원본 보존
      priceEl.className    = `wc-price ${cls}`;

      const changeDisplay = isKrx
        ? fmt(Math.abs(change))
        : Math.abs(change).toFixed(2);
      chgEl.textContent    = `${arrow} ${sign}${changeDisplay} (${sign}${changePct}%)`;
      chgEl.className      = `wc-chg ${cls}`;
      // 금일수익 계산용 — 부호 포함 실제 등락값 저장
      chgEl.dataset.change = change;
	  chgEl.dataset.changePct = changePct; // ← 이 줄 추가 (% 값)
      // 목표가/손절가 알림 체크
      checkPriceAlert(item.symbol, price);	  
    } catch(e) {
      if (priceEl) priceEl.textContent = '-';
    }
  }));
  // 가격 갱신 후 보유정보/스냅샷 재계산
  loadHoldings();
  if (_snapshotOn) updateSnapshot();
  savePortfolioSnapshot(); // ← 하루 1회 자동 스냅샷 저장
}



// 서버에서 포트폴리오 불러와서 복원
async function loadPortfolio() {
  try {
    const res  = await fetch(`${API}/portfolio`);
    const data = await res.json();
    if (!data) return false;

    let restored = false;

    // localStorage에 직접 저장 (saveMyList/saveList는 savePortfolio를 재호출하므로 제외)
    if (data.myList && data.myList.length) {
      localStorage.setItem('myStockPro', JSON.stringify(data.myList));
      restored = true;
    }
    if (data.watchList && data.watchList.length) {
      localStorage.setItem('watchPro', JSON.stringify(data.watchList));
      restored = true;
    }
    // 보유정보 복원
    if (data.holdings) {
      Object.entries(data.holdings).forEach(([symbol, h]) => {
        localStorage.setItem(`holding_${symbol}`, JSON.stringify({ avg_price: h.avg_price, qty: h.qty }));
      });
    }
    return restored;
  } catch(e) {
    console.warn('[포트폴리오 로드 오류]', e);
    return false;
  }
}

/* ══════════════════════════════════
   보유정보 직접 입력 (localStorage)
══════════════════════════════════ */
let _hmSymbol = '';  // 현재 편집 중인 종목코드

function getHolding(symbol) {
  return JSON.parse(localStorage.getItem(`holding_${symbol}`) || 'null');
}
function setHolding(symbol, avg, qty) {
  localStorage.setItem(`holding_${symbol}`, JSON.stringify({ avg_price: avg, qty }));
  savePortfolio();
}
function delHolding(symbol) {
  localStorage.removeItem(`holding_${symbol}`);
  savePortfolio();
}
/* ── 목표가/손절가 ── */
function getAlert(symbol) {
  return JSON.parse(localStorage.getItem(`alert_${symbol}`) || 'null');
}
function setAlert(symbol, target, stop) {
  if (!target && !stop) {
    localStorage.removeItem(`alert_${symbol}`);
    return;
  }
  localStorage.setItem(`alert_${symbol}`, JSON.stringify({ target, stop }));
}
function clearAlert(symbol) {
  localStorage.removeItem(`alert_${symbol}`);
}
/* ── 알림 체크 ── */
const _alertFired = new Set(); // 이미 울린 알림 (중복 방지)

function checkPriceAlert(symbol, price) {
  const a = getAlert(symbol);
  if (!a) return;

  // 목표가 도달
  if (a.target && price >= a.target) {
    const key = `${symbol}_target`;
    if (!_alertFired.has(key)) {
      _alertFired.add(key);
      const name = document.getElementById(`wcp-${symbol}`)
        ?.closest('.wc-card')?.dataset.name || symbol;
      showToast(`🎯 ${name} 목표가 도달! (${fmt(a.target)})`);
      playAlertSound('target');
    }
  } else {
    _alertFired.delete(`${symbol}_target`); // 가격 내려가면 초기화
  }

  // 손절가 도달
  if (a.stop && price <= a.stop) {
    const key = `${symbol}_stop`;
    if (!_alertFired.has(key)) {
      _alertFired.add(key);
      const name = document.getElementById(`wcp-${symbol}`)
        ?.closest('.wc-card')?.dataset.name || symbol;
      showToast(`🛑 ${name} 손절가 도달! (${fmt(a.stop)})`);
      playAlertSound('stop');
    }
  } else {
    _alertFired.delete(`${symbol}_stop`); // 가격 올라가면 초기화
  }
}

/* ── 알림 사운드 (Web Audio API) ── */
function playAlertSound(type) {
  try {
    const ctx  = new (window.AudioContext || window.webkitAudioContext)();
    const osc  = ctx.createOscillator();
    const gain = ctx.createGain();
    osc.connect(gain);
    gain.connect(ctx.destination);

    if (type === 'target') {
      // 목표가: 밝고 상승하는 소리 (도-미-솔)
      osc.type = 'sine';
      osc.frequency.setValueAtTime(523, ctx.currentTime);       // 도
      osc.frequency.setValueAtTime(659, ctx.currentTime + 0.15); // 미
      osc.frequency.setValueAtTime(784, ctx.currentTime + 0.30); // 솔
    } else {
      // 손절가: 낮고 하강하는 소리 (솔-미-도)
      osc.type = 'sine';
      osc.frequency.setValueAtTime(784, ctx.currentTime);       // 솔
      osc.frequency.setValueAtTime(659, ctx.currentTime + 0.15); // 미
      osc.frequency.setValueAtTime(523, ctx.currentTime + 0.30); // 도
    }

    gain.gain.setValueAtTime(0.4, ctx.currentTime);
    gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.6);
    osc.start(ctx.currentTime);
    osc.stop(ctx.currentTime + 0.6);
  } catch(e) {
    console.warn('사운드 재생 실패:', e);
  }
}
function openHoldingModal(symbol, name) {
  _hmSymbol = symbol;
  const isKrx = /^[0-9]{6}$/.test(symbol);
  const unit  = isKrx ? '원 (₩)' : '달러 ($)';
  document.getElementById('hm-title').innerHTML =
    `✏️ ${name} 보유정보 <button class="hm-close" onclick="closeHoldingModal()">✕</button>`;
  document.querySelector('#holding-modal .hm-label').textContent = `매수 평균가 (${unit})`;
  const h = getHolding(symbol);
  document.getElementById('hm-avgprice').value = h ? h.avg_price : '';
  document.getElementById('hm-qty').value      = h ? h.qty       : '';
  // 목표가/손절가 불러오기
  const t = getAlert(symbol);
  document.getElementById('hm-target').value = t ? t.target : '';
  document.getElementById('hm-stop').value   = t ? t.stop   : '';
  document.getElementById('holding-modal').classList.add('open');
  setTimeout(() => document.getElementById('hm-avgprice').focus(), 100);
}
function closeHoldingModal() {
  document.getElementById('holding-modal').classList.remove('open');
  _hmSymbol = '';
}
function saveHolding() {
  const avg = parseFloat(document.getElementById('hm-avgprice').value);
  const qty = parseFloat(document.getElementById('hm-qty').value);
  if (!avg || !qty || avg <= 0 || qty <= 0) {
    showToast('⚠ 평균가와 수량을 입력하세요'); return;
  }
  setHolding(_hmSymbol, avg, qty);
  // 목표가/손절가 저장
  const target = parseFloat(document.getElementById('hm-target').value) || 0;
  const stop   = parseFloat(document.getElementById('hm-stop').value)   || 0;
  setAlert(_hmSymbol, target, stop);
  closeHoldingModal();
  loadHoldings();
  showToast('✅ 보유정보 저장됨');
}
function clearHolding() {
  delHolding(_hmSymbol);
  closeHoldingModal();
  loadHoldings();
  showToast('🗑 보유정보 삭제됨');
}


// 투자 스냅샷 ON/OFF
let _snapshotOn = false;
let _usdkrw     = 1350;   // 환율 기본값, 지수 갱신 시 업데이트
let _usInKrw    = false;  // 미국주식 원화 표시 여부

function toggleUsUnit() {
  _usInKrw = !_usInKrw;
  const flag = document.getElementById('isn-us-flag');
  if (flag) flag.title = _usInKrw ? '원화(₩) 표시 중 — 클릭시 달러' : '달러($) 표시 중 — 클릭시 원화';
  updateSnapshot();
}
function toggleSnapshot() {
  _snapshotOn = !_snapshotOn;
  const btn = document.getElementById('isn-toggle');
  btn.className = _snapshotOn ? 'isn-toggle-btn on' : 'isn-toggle-btn';
  // 박스는 항상 표시, 값 셀만 숨김
  document.querySelectorAll('.isn-val').forEach(el => {
    el.style.color = _snapshotOn ? '' : 'transparent';
  });
  document.querySelectorAll('.isn-cat').forEach(el => {
    el.style.color = _snapshotOn ? '' : 'transparent';
  });
  if (_snapshotOn) updateSnapshot();
}

// 투자 스냅샷 업데이트 (현재가 DOM에서 직접 읽기)
function updateSnapshot() {
  if (!_snapshotOn) return;
  const myList = getMyList();
  let krBuy=0, krEval=0, usBuy=0, usEval=0;

  myList.forEach(item => {
    const h = getHolding(item.symbol);
    if (!h || !h.qty || !h.avg_price) return;

    // data-price 속성에서 원본 숫자 직접 읽기 (텍스트 파싱 불필요)
    const priceEl   = document.getElementById(`wcp-${item.symbol}`);
    const currPrice = priceEl?.dataset?.price ? parseFloat(priceEl.dataset.price) : 0;
    if (!currPrice) return;   // 가격 미로드 시 스냅샷 계산에서 제외
    const buyAmt  = h.avg_price * h.qty;
    const evalAmt = currPrice * h.qty;
    const isKrx   = /^[0-9]{6}$/.test(item.symbol);

    if (isKrx) {
      krBuy  += buyAmt;
      krEval += evalAmt;
    } else {
      // 미국: avg_price는 달러, currPrice도 달러 → 달러 저장
      usBuy  += buyAmt;
      usEval += evalAmt;
    }
  });

  const totBuy  = krBuy  + usBuy;
  const totEval = krEval + usEval;

  function setRow(ids, buy, eval_, unit='₩') {
    const profit = eval_ - buy;
    const rate   = buy ? (profit / buy * 100) : 0;
    const pCls   = profit > 0 ? 'pos' : profit < 0 ? 'neg' : 'muted';
    const sign   = profit > 0 ? '+' : profit < 0 ? '-' : '';
    const dec    = unit === '$' ? 2 : 0;
    const fmt    = v => unit + Math.abs(v).toLocaleString('ko-KR', {maximumFractionDigits: dec, minimumFractionDigits: dec});
    const set    = (id, txt, cls) => {
      const el = document.getElementById(id);
      if (!el) return;
      el.textContent = txt;
      el.style.color = '';  // 인라인 color 초기화
      el.classList.remove('pos','neg','muted');
      if (cls) el.classList.add(cls);
    };
    set(ids[0], fmt(buy),   '');
    set(ids[1], fmt(eval_), '');
    set(ids[2], `${sign}${fmt(Math.abs(profit))}`, pCls);
    set(ids[3], `${sign}${Math.abs(rate).toFixed(2)}%`, pCls);
  }

  // 미국: 달러 표시 or 원화 표시 토글
  const usKrwBuy  = usBuy  * _usdkrw;
  const usKrwEval = usEval * _usdkrw;
  const totBuyKrw  = krBuy  + usKrwBuy;
  const totEvalKrw = krEval + usKrwEval;

  setRow(['isn-kr-buy','isn-kr-eval','isn-kr-profit','isn-kr-rate'], krBuy, krEval, '₩');
  if (_usInKrw) {
    setRow(['isn-us-buy','isn-us-eval','isn-us-profit','isn-us-rate'], usKrwBuy, usKrwEval, '₩');
  } else {
    setRow(['isn-us-buy','isn-us-eval','isn-us-profit','isn-us-rate'], usBuy, usEval, '$');
  }
  setRow(['isn-tot-buy','isn-tot-eval','isn-tot-profit','isn-tot-rate'], totBuyKrw, totEvalKrw, '₩');
}
/* ── 포트폴리오 히스토리 스냅샷 ── */
function savePortfolioSnapshot() {
  const myList = getMyList();
  let krBuy=0, krEval=0, usBuy=0, usEval=0;

  myList.forEach(item => {
    const h = getHolding(item.symbol);
    if (!h || !h.qty || !h.avg_price) return;
    const priceEl  = document.getElementById(`wcp-${item.symbol}`);
    const currPrice = priceEl?.dataset?.price ? parseFloat(priceEl.dataset.price) : 0;
    if (!currPrice) return;
    const buyAmt  = h.avg_price * h.qty;
    const evalAmt = currPrice   * h.qty;
    const isKrx   = /^[0-9]{6}$/.test(item.symbol);
    if (isKrx) { krBuy += buyAmt; krEval += evalAmt; }
    else        { usBuy += buyAmt; usEval += evalAmt; }
  });

  // 가격 로드된 종목이 없으면 저장 안 함
  if (krEval === 0 && usEval === 0) return;

  const today    = new Date().toISOString().slice(0, 10); // 'YYYY-MM-DD'
  const history  = JSON.parse(localStorage.getItem('portfolioHistory') || '[]');

  // 오늘 데이터가 이미 있으면 업데이트, 없으면 추가
  const idx = history.findIndex(x => x.date === today);
  const snap = {
    date:    today,
    krBuy:   Math.round(krBuy),
    krEval:  Math.round(krEval),
    usBuy:   Math.round(usBuy  * _usdkrw),  // 원화 환산
    usEval:  Math.round(usEval * _usdkrw),  // 원화 환산
    total:   Math.round(krEval + usEval * _usdkrw)
  };

  if (idx >= 0) history[idx] = snap;
  else          history.push(snap);

  // 최대 180일치만 보관
  if (history.length > 180) history.splice(0, history.length - 180);
  localStorage.setItem('portfolioHistory', JSON.stringify(history));
}
/* ── 포트폴리오 히스토리 모달 ── */
let _historyRange = '1w';
let _historyChart = null;

function openHistoryModal() {
  document.getElementById('history-modal').style.display = 'flex';
  renderHistoryChart(_historyRange);
}

function closeHistoryModal() {
  document.getElementById('history-modal').style.display = 'none';
  if (_historyChart) {
    _historyChart.remove();
    _historyChart = null;
  }
}

function setHistoryRange(range, el) {
  _historyRange = range;
  // 버튼 활성화 상태 변경
  document.querySelectorAll('[id^="hist-btn-"]').forEach(b => b.classList.remove('active'));
  el.classList.add('active');
  renderHistoryChart(range);
}

function renderHistoryChart(range) {
  const allHistory = JSON.parse(localStorage.getItem('portfolioHistory') || '[]');

  // 기간 필터 적용
  const now   = new Date();
  const cutoff = new Date();
  if      (range === '1w')  cutoff.setDate(now.getDate() - 7);
  else if (range === '1m')  cutoff.setMonth(now.getMonth() - 1);
  else if (range === '3m')  cutoff.setMonth(now.getMonth() - 3);
  else                      cutoff.setFullYear(2000); // 전체

  const filtered = allHistory.filter(x => new Date(x.date) >= cutoff);

  const emptyEl = document.getElementById('history-empty');
  const chartEl = document.getElementById('history-chart');

  if (!filtered.length) {
    emptyEl.style.display = 'block';
    chartEl.style.display = 'none';
    document.getElementById('history-summary').innerHTML = '';
    return;
  }

  emptyEl.style.display = 'none';
  chartEl.style.display = 'block';

  // 기존 차트 제거
  if (_historyChart) { _historyChart.remove(); _historyChart = null; }
  chartEl.innerHTML = '';

  // LightweightCharts 생성
  _historyChart = LightweightCharts.createChart(chartEl, {
    width:  chartEl.clientWidth,
    height: 300,
    layout: {
      background: { color: 'transparent' },
      textColor:  '#94a3b8',
    },
    grid: {
      vertLines: { color: 'rgba(148,163,184,0.1)' },
      horzLines: { color: 'rgba(148,163,184,0.1)' },
    },
    rightPriceScale: { borderColor: 'rgba(148,163,184,0.2)' },
    timeScale:       { borderColor: 'rgba(148,163,184,0.2)', timeVisible: true },
  });

  // 평가금액 라인 (주황)
  const evalSeries = _historyChart.addLineSeries({
    color:     '#f59e0b',
    lineWidth: 2,
    title:     '평가금액',
  });

  // 투자금액 라인 (파랑 점선)
  const buySeries = _historyChart.addLineSeries({
    color:       '#3b82f6',
    lineWidth:   1,
    lineStyle:   2, // 점선
    title:       '투자금액',
  });

  const evalData = filtered.map(x => ({
    time:  x.date,
    value: x.total
  }));

  const buyData = filtered.map(x => ({
    time:  x.date,
    value: x.krBuy + x.usBuy
  }));

  evalSeries.setData(evalData);
  buySeries.setData(buyData);
  _historyChart.timeScale().fitContent();

  // 요약 정보
  const first = filtered[0];
  const last  = filtered[filtered.length - 1];
  const diff  = last.total - (first.krBuy + first.usBuy);
  const rate  = first.krBuy + first.usBuy
    ? (diff / (first.krBuy + first.usBuy) * 100).toFixed(2)
    : 0;
  const pCls  = diff >= 0 ? '#ef4444' : '#3b82f6';
  const sign  = diff >= 0 ? '+' : '';

  document.getElementById('history-summary').innerHTML = `
    <span style="color:#f59e0b;">● 평가금액 <b>${last.total.toLocaleString()}원</b></span>
    <span style="color:#3b82f6;">● 투자금액 <b>${(last.krBuy+last.usBuy).toLocaleString()}원</b></span>
    <span style="color:${pCls};">● 손익 <b>${sign}${diff.toLocaleString()}원 (${sign}${rate}%)</b></span>
  `;
}
// 카드 보유정보 표시 (현재가 기반 실시간 계산)
function loadHoldings() {
  getMyList().forEach(item => {
    const h    = getHolding(item.symbol);
    const rEl  = document.getElementById(`wcr-${item.symbol}`);
    const hEl  = document.getElementById(`wch-${item.symbol}`);
    if (!rEl || !hEl) return;
    if (!h || !h.qty || !h.avg_price) { rEl.classList.remove('visible'); return; }

    // data-price 속성에서 원본 숫자 직접 읽기 (텍스트 파싱 불필요)
    const priceEl  = document.getElementById(`wcp-${item.symbol}`);
    const currPrice = priceEl?.dataset?.price ? parseFloat(priceEl.dataset.price) : 0;

    const buyAmt   = h.avg_price * h.qty;
    // 가격 미로드 시 — 총평가/수익/수익률 모두 표시 보류 (가짜 0 방지)
    const priceReady = currPrice > 0;
    const evalAmt  = priceReady ? currPrice * h.qty : 0;
    const profit   = priceReady ? evalAmt - buyAmt  : 0;
    const profitRt = (priceReady && buyAmt) ? (profit / buyAmt * 100) : 0;

    // 금일 수익: dataset.change × 수량 (미로드 시 0 대신 --)
    const chgEl    = document.getElementById(`wcg-${item.symbol}`);
    const chgAmt   = (priceReady && chgEl?.dataset?.change) ? parseFloat(chgEl.dataset.change) : null;
    const todayPnl = chgAmt != null ? chgAmt * h.qty : null;

    const isKrx = /^[0-9]{6}$/.test(item.symbol);
    const fmtP  = v => isKrx
      ? Math.round(Math.abs(v)).toLocaleString('ko-KR') + '원'
      : '$' + Math.abs(v).toFixed(2);
    const fmtA  = v => isKrx
      ? Math.round(Math.abs(v)).toLocaleString('ko-KR')
      : '$' + Math.abs(v).toFixed(2);

    const pCls  = profit  >= 0 ? 'pos' : 'neg';
    const tCls  = (todayPnl != null && todayPnl >= 0) ? 'pos' : 'neg';
    const pSign = profit  >= 0 ? '+' : '-';
    const tSign = (todayPnl != null && todayPnl >= 0) ? '+' : '-';

    const evalStr   = priceReady ? fmtA(evalAmt) : '···';
    const profStr   = priceReady ? `${pSign}${fmtA(profit)}` : '···';
    const rateStr   = priceReady ? `${pSign}${Math.abs(profitRt).toFixed(2)}%` : '···';
    const todayStr  = todayPnl != null ? `${tSign}${fmtA(todayPnl)}` : '···';

    hEl.innerHTML = `
      <div class="wc-hold-row"><span class="wc-hold-label">수량</span><span class="wc-hold-val">${Math.round(h.qty).toLocaleString()}주</span></div>
      <div class="wc-hold-row"><span class="wc-hold-label">평균가</span><span class="wc-hold-val">${fmtP(h.avg_price)}</span></div>
      <div class="wc-hold-row"><span class="wc-hold-label">구매액</span><span class="wc-hold-val">${fmtA(buyAmt)}</span></div>
      <div class="wc-hold-row"><span class="wc-hold-label">총평가</span><span class="wc-hold-val">${evalStr}</span></div>
      <div class="wc-hold-row"><span class="wc-hold-label">수익</span><span class="wc-hold-val ${priceReady ? pCls : ''}">${profStr}</span></div>
      <div class="wc-hold-row"><span class="wc-hold-label">수익률</span><span class="wc-hold-val ${priceReady ? pCls : ''}">${rateStr}</span></div>
      <div class="wc-hold-row"><span class="wc-hold-label">금일</span><span class="wc-hold-val ${todayPnl != null ? tCls : ''}">${todayStr}</span></div>`;
    rEl.classList.add('visible');
  });
}

// 모달 바깥 클릭 닫기
window.addEventListener('DOMContentLoaded', () => {
  document.getElementById('holding-modal')?.addEventListener('click', function(e) {
    if (e.target === this) closeHoldingModal();
  });
  // Enter 키 저장
  ['hm-avgprice','hm-qty','hm-target','hm-stop'].forEach(id => {
    document.getElementById(id)?.addEventListener('keydown', e => {
      if (e.key === 'Enter') saveHolding();
    });
  });
});

// 관심종목 카드 렌더 후 가격 로딩 + 30초마다 갱신
function scheduleWatchPrices() {
  clearInterval(_wcPriceTimer);
  loadWatchPrices();
  _wcPriceTimer = setInterval(loadWatchPrices, 30000);
  // 크로스 알림 ON이면 체크
  if (localStorage.getItem('crossAlert') === 'on') loadAllCrossAlerts();
}

/* ══════════════════════════════════
   엑셀 출력
══════════════════════════════════ */





/* ── 마우스 플로팅 툴팁 ── */
function addFloatingTooltip(chart, container, mainSeries) {
  if (!container || container.querySelector('.floating-tooltip')) return;
  const toolTip = document.createElement('div');
  toolTip.className = 'floating-tooltip';
  container.appendChild(toolTip);

  chart.subscribeCrosshairMove(param => {
    if (
      param.point === undefined || !param.time ||
      param.point.x < 0 || param.point.x > container.clientWidth ||
      param.point.y < 0 || param.point.y > container.clientHeight
    ) { toolTip.style.display = 'none'; return; }

    let dateStr = '';
    if (typeof param.time === 'object') {
      dateStr = `${param.time.year}-${String(param.time.month).padStart(2,'0')}-${String(param.time.day).padStart(2,'0')}`;
    } else if (typeof param.time === 'string') {
      dateStr = param.time;
    } else {
      const d = new Date(param.time * 1000);
      dateStr = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')} ${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}`;
    }

    const data = param.seriesData.get(mainSeries);
    let priceText = '';
    if (data) {
      if (data.close !== undefined) priceText = `종가: ${data.close.toLocaleString()}`;
      else if (data.value !== undefined) priceText = `값: ${data.value.toLocaleString()}`;
    }

    toolTip.innerHTML = `<div style="color:var(--muted);margin-bottom:2px;">${dateStr}</div><div style="font-weight:bold;">${priceText}</div>`;
    toolTip.style.display = 'block';

    let left = param.point.x + 15;
    let top  = param.point.y + 15;
    if (left + toolTip.clientWidth  > container.clientWidth)  left = param.point.x - toolTip.clientWidth  - 15;
    if (top  + toolTip.clientHeight > container.clientHeight) top  = param.point.y - toolTip.clientHeight - 15;
    toolTip.style.left = left + 'px';
    toolTip.style.top  = top  + 'px';
  });
}

/* ══════════════════════════════════
   차트 패널 팝아웃 모달
══════════════════════════════════ */
const chartModals = {}; // key: cid+type

function popoutChart(cid, type) {
  const w = activeWidgets[cid];
  if (!w) return;

  const key = `${cid}-${type}`;
  // 이미 열려있으면 앞으로
  if (chartModals[key]) {
    chartModals[key].style.zIndex = ++zBase;
    return;
  }

  const titles = { candle:'캔들차트', macd:'MACD', rsi:'RSI(14)' };
  const backdrop = document.getElementById('chart-modal-backdrop');
  backdrop.classList.add('open');

  // 모달 생성
  const modal = document.createElement('div');
  modal.className = 'chart-modal';
  modal.style.cssText = `width:700px;height:450px;left:${100+Object.keys(chartModals).length*30}px;top:${80+Object.keys(chartModals).length*30}px;z-index:${++zBase}`;
modal.innerHTML = `
    <div class="chart-modal-bar" id="cmbar-${key}">
      <span class="chart-modal-title">📊 ${w.name} — ${titles[type]}</span>
      <div style="display:flex;align-items:center;gap:6px;margin-right:6px;">
        <select class="win-ctrl-select" id="cmtf-${key}" data-cid="${cid}" data-key="${key}" data-type="${type}" style="font-size:10px;padding:1px 4px;">
          <option value="day">일봉</option>
          <option value="1m">1분봉</option>
          <option value="5m">5분봉</option>
          <option value="15m">15분봉</option>
          <option value="30m">30분봉</option>
        </select>
        <button class="ret-name" onclick="zoomChartModal('${key}','1d',this)">오늘</button>
        <button class="ret-name" onclick="zoomChartModal('${key}','1w',this)">1주</button>
        <button class="ret-name" onclick="zoomChartModal('${key}','1mo',this)">1개월</button>
        <button class="ret-name" onclick="zoomChartModal('${key}','3mo',this)">3개월</button>
        <button class="ret-name" onclick="zoomChartModal('${key}','1y',this)">1년</button>
      </div>
      <button class="export-btn" style="margin-right:4px;" onclick="exportChart('${cid}','${key}')">📊 엑셀</button>
      <button class="chart-modal-close" onclick="closeChartModal('${key}')">✕</button>
    </div>
    <div class="chart-modal-body" id="cmbody-${key}">
      <div id="cmchart-${key}" style="width:100%;height:100%;position:relative;">
        <div class="rb-selection" id="cmrb-${key}"></div>
      </div>
    </div>
    <div class="cm-resize n"  data-key="${key}" data-dir="n"></div>
    <div class="cm-resize s"  data-key="${key}" data-dir="s"></div>
    <div class="cm-resize e"  data-key="${key}" data-dir="e"></div>
    <div class="cm-resize w"  data-key="${key}" data-dir="w"></div>
    <div class="cm-resize nw" data-key="${key}" data-dir="nw"></div>
    <div class="cm-resize ne" data-key="${key}" data-dir="ne"></div>
    <div class="cm-resize sw" data-key="${key}" data-dir="sw"></div>
    <div class="cm-resize se" data-key="${key}" data-dir="se"></div>`;

  backdrop.appendChild(modal);
  chartModals[key] = modal;

  modal.addEventListener('mousedown', () => { modal.style.zIndex = ++zBase; });
  // ESC 키로 이 모달 닫기
  modal._escHandler = (e) => {
    if (e.key === 'Escape') {
      // 가장 높은 z-index 모달 닫기
      const topKey = Object.entries(chartModals)
        .sort((a,b) => (+b[1].style.zIndex||0) - (+a[1].style.zIndex||0))[0]?.[0];
      if (topKey) closeChartModal(topKey);
    }
  };
  document.addEventListener('keydown', modal._escHandler);

  // 드래그 이동
  makeCMDraggable(modal, document.getElementById(`cmbar-${key}`));
  // 리사이즈
  makeCMResizable(modal, key);

  // 차트 생성 (약간 지연 - DOM 렌더 후)
  // 현재 팝업 봉 종류로 select 초기화
  setTimeout(() => {
    const sel = document.getElementById(`cmtf-${key}`);
    if (sel) {
      sel.value = activeWidgets[cid]?.interval || 'day';
      sel.addEventListener('change', function() {
        const c = this.dataset.cid;
        const k = this.dataset.key;
        const t = this.dataset.type;
        changeModalTF(c, k, t, this.value);
      });
    }
    initChartModal(cid, key, type);
  }, 60);
}

function closeChartModal(key) {
  const modal = chartModals[key];
  if (!modal) return;
  modal._disposed = true;  // forecast.js 가 체크하는 플래그
  if (modal._cmChart) { try { modal._cmChart.remove(); } catch(e){} }
  if (modal._cmRO)    { try { modal._cmRO.disconnect(); } catch(e){} }
  if (modal._escHandler) document.removeEventListener('keydown', modal._escHandler);
  modal.remove();
  delete chartModals[key];
  if (Object.keys(chartModals).length === 0)
    document.getElementById('chart-modal-backdrop').classList.remove('open');
}

function initChartModal(cid, key, type) {
  const w = activeWidgets[cid];
  if (!w) return;
  // modalData가 있으면 그걸 쓰고 없으면 w.data 사용
  const modal = chartModals[key];
  const initData = (modal && modal._modalData) ? modal._modalData : w.data;
  const initInterval = (modal && modal._modalInterval) ? modal._modalInterval : (w.interval || 'day');
  initChartModalWithData(cid, key, type, initData, initInterval);
}

function initChartModalWithData(cid, key, type, initData, initInterval) {
  const w = activeWidgets[cid];
  if (!w) return;
  const container = document.getElementById(`cmchart-${key}`);
  if (!container) return;

  const bg = cssVar('--cbg'), tc = cssVar('--ctxt'), gl = cssVar('--gcol');

  // ★ LightweightCharts 기본 스크롤/줌 완전히 끔 → 직접 구현
  const cmChart = LightweightCharts.createChart(container, {
    layout:{ background:{color:bg}, textColor:tc },
    grid:{ vertLines:{color:gl}, horzLines:{color:gl} },
    crosshair:{ mode:1 },
    timeScale:{ timeVisible:true, secondsVisible:false, borderColor:cssVar('--border') },
    handleScroll: false,
    handleScale:  false,
    width:  container.offsetWidth,
    height: container.offsetHeight,
  });
  chartModals[key]._cmChart = cmChart;

  // disposed 안전 래퍼
  function safeChart(fn) {
    const m = chartModals[key];
    if (!m || m._disposed || m._cmChart !== cmChart) return;
    try { fn(); } catch(e) { /* disposed 에러 무시 */ }
  }

  // ── Y축 줌 상태 (데이터 세팅 전에 선언해야 registerSeries 사용 가능) ──
  let _yRange = null;
  let _cmSeries = [];
  function registerSeries(s) { _cmSeries.push(s); }
  function applyYRange(min, max) {
    _yRange = { min, max };
    _cmSeries.forEach(s => s.applyOptions({
      autoscaleInfoProvider: () => ({
        priceRange: { minValue: _yRange.min, maxValue: _yRange.max },
        margins: { above: 0, below: 0 }
      })
    }));
    cmChart.priceScale('right').applyOptions({ autoScale: true });
  }
  function resetYRange() {
    _yRange = null;
    _cmSeries.forEach(s => s.applyOptions({ autoscaleInfoProvider: () => null }));
    cmChart.priceScale('right').applyOptions({ autoScale: true });
  }

  // 데이터 세팅
  const data = initData;
  if (!data || !data.length) { cmChart.timeScale().fitContent(); return; }

  if (type === 'candle') {
    const cs  = cmChart.addCandlestickSeries({borderVisible:false});
    const m20 = cmChart.addLineSeries({lineWidth:2,  color:'#f59e0b', title:'20MA'});
    const m60 = cmChart.addLineSeries({lineWidth:1.5,color:'#3b82f6', title:'60MA'});
    const m120= cmChart.addLineSeries({lineWidth:1.5,color:'#a855f7', title:'120MA'});
    const bbU = cmChart.addLineSeries({lineWidth:1,  color:'#94a3b8', lineStyle:2});
    const bbL = cmChart.addLineSeries({lineWidth:1,  color:'#94a3b8', lineStyle:2});
    cs.setData(data);
    const bb = calcBB(data,20);
    m20.setData(bb.sma); bbU.setData(bb.upper); bbL.setData(bb.lower);
    m60.setData(calcMA(data,60)); m120.setData(calcMA(data,120));
    [cs, m20, m60, m120, bbU, bbL].forEach(registerSeries);

    // 🌟 모달 툴팁
    addFloatingTooltip(cmChart, container, cs);

    // 🌟 모달 ARIMA 예측선 (일봉 전용)
    if (initInterval === 'day' && typeof drawForecastLine === 'function') {
      const ww = activeWidgets[cid];
      chartModals[key].chart          = cmChart;
      chartModals[key].symbol         = ww?.symbol;
      chartModals[key].forecastSeries = null;
      chartModals[key].validSeries    = null;
      chartModals[key].errorSeries    = null;
      chartModals[key].prophetSeries  = null;
      chartModals[key].ensembleSeries = null;
      drawForecastLine(chartModals[key]).then(() => {
        if (chartModals[key]?.forecastSeries) registerSeries(chartModals[key].forecastSeries);
      });
    }
  } else if (type === 'macd') {
    const md = calcMACD(data);
    const _mh = cmChart.addHistogramSeries(); _mh.setData(md.hist);
    const _ml = cmChart.addLineSeries({color:'#2962FF',lineWidth:1.5,title:'MACD'}); _ml.setData(md.macd);
    const _ms = cmChart.addLineSeries({color:'#FF6D00',lineWidth:1.5,title:'Signal'}); _ms.setData(md.signal);
    [_mh, _ml, _ms].forEach(registerSeries);
  } else if (type === 'rsi') {
    const rsiData = calcRSIFull(data);
    const _rr = cmChart.addLineSeries({color:'#a855f7',lineWidth:1.5,title:'RSI'}); _rr.setData(rsiData);
    const _ro = cmChart.addLineSeries({color:'#ef4444',lineWidth:1,lineStyle:2}); _ro.setData(rsiData.map(d=>({time:d.time,value:70})));
    const _rs = cmChart.addLineSeries({color:'#22c55e',lineWidth:1,lineStyle:2}); _rs.setData(rsiData.map(d=>({time:d.time,value:30})));
    [_rr, _ro, _rs].forEach(registerSeries);
  }
  // 오늘 기준 앞뒤 43일 표시 (마우스 10번 줌인 동일 효과)
  // 예측선 10일도 보이도록 to는 오늘+12일
  (function() {
    const now    = new Date();
    const from   = new Date(now); from.setDate(from.getDate() - 43);
    const to     = new Date(now); to.setDate(to.getDate() + 12);
    const fromTs = Math.floor(from.getTime() / 1000);
    const toTs   = Math.floor(to.getTime()   / 1000);
    try {
      cmChart.timeScale().setVisibleRange({ from: fromTs, to: toTs });
    } catch(e) {
      cmChart.timeScale().fitContent();
    }
  })();

  // ══ 커스텀 인터랙션 ══════════════════════════════

  // 헬퍼: X축 줌 (logicalRange 기준 중앙 확대/축소)
  function zoomX(factor) {
    const ts  = cmChart.timeScale();
    const lr  = ts.getVisibleLogicalRange();
    if (!lr) return;
    const mid  = (lr.from + lr.to) / 2;
    const half = (lr.to - lr.from) / 2 * factor;
    ts.setVisibleLogicalRange({ from: mid - half, to: mid + half });
  }

  // 현재 보이는 범위의 가격 min/max를 시리즈 데이터에서 직접 계산
  function getVisiblePriceRange() {
    const lr = cmChart.timeScale().getVisibleLogicalRange();
    if (!lr || !_cmSeries.length) return null;
    let minP = Infinity, maxP = -Infinity;
    _cmSeries.forEach(s => {
      try {
        const d = s.data ? s.data() : null;
        if (!d || !d.length) return;
        const from = Math.max(0, Math.floor(lr.from));
        const to   = Math.min(d.length - 1, Math.ceil(lr.to));
        for (let i = from; i <= to; i++) {
          const v = d[i];
          if (!v) continue;
          const hi = v.high ?? v.value ?? v.close ?? null;
          const lo = v.low  ?? v.value ?? v.close ?? null;
          if (hi != null && hi > maxP) maxP = hi;
          if (lo != null && lo < minP) minP = lo;
        }
      } catch(e) {}
    });
    return (minP < Infinity && maxP > -Infinity) ? { min: minP, max: maxP } : null;
  }

  function zoomY(factor) {
    if (!_yRange) {
      const pr = getVisiblePriceRange();
      if (!pr) return;
      // 약간 여유 추가
      const pad = (pr.max - pr.min) * 0.1;
      _yRange = { min: pr.min - pad, max: pr.max + pad };
    }
    const mid  = (_yRange.min + _yRange.max) / 2;
    const half = (_yRange.max - _yRange.min) / 2 * factor;
    applyYRange(mid - half, mid + half);
  }

  // 마우스 휠: 일반=XY / Ctrl=X만 / Shift=Y만
  // 줌인/아웃 대칭: BASE^1 (아웃) vs 1/BASE^1 (인) → 역수 관계로 정확한 대칭
  container.addEventListener('wheel', e => {
    e.preventDefault();
    e.stopPropagation();
    safeChart(() => {
      const BASE = 1.08;  // 한 스텝당 8% 변화 — 이 값만 조정하면 됨
      const factor = e.deltaY > 0 ? BASE : 1 / BASE;  // 줌아웃/줌인 대칭
      if (e.shiftKey)       zoomY(factor);
      else if (e.ctrlKey)   zoomX(factor);
      else { zoomX(factor); zoomY(factor); }
    });
  }, { passive: false });

  // 클릭 드래그: X(시간축 스크롤) + Y(가격축 이동) 동시
  let dragState = null;
  container.addEventListener('mousedown', e => {
    if (e.button !== 0) return;
    if (e.ctrlKey && e.altKey) return;  // Rubber Band 줌에 양보
    const ts = cmChart.timeScale();
    const lr = ts.getVisibleLogicalRange();
    const pr = _yRange || getVisiblePriceRange();
    dragState = {
      sx: e.clientX, sy: e.clientY,
      lr: lr ? {...lr} : null,
      lrSpan: lr ? (lr.to - lr.from) : 0,
      initYRange: pr ? { min: pr.min, max: pr.max } : null,
    };
    container.style.cursor = 'grabbing';
    e.preventDefault();
  });
  document.addEventListener('mousemove', e => {
    if (!dragState) return;
    safeChart(() => {
      const dx = e.clientX - dragState.sx;
      const dy = e.clientY - dragState.sy;
      const cw = container.offsetWidth  || 1;
      const ch = container.offsetHeight || 1;
      if (dragState.lr && dragState.lrSpan) {
        const dLogical = -(dx / cw) * dragState.lrSpan;
        cmChart.timeScale().setVisibleLogicalRange({
          from: dragState.lr.from + dLogical,
          to:   dragState.lr.to   + dLogical,
        });
      }
      if (dragState.initYRange) {
        const dPrice = (dy / ch) * (dragState.initYRange.max - dragState.initYRange.min);
        applyYRange(dragState.initYRange.min + dPrice, dragState.initYRange.max + dPrice);
      }
    });
  });
  const stopDrag = () => {
    if (dragState) { dragState = null; container.style.cursor = 'default'; }
  };
  document.addEventListener('mouseup', stopDrag);
  container.addEventListener('mouseleave', () => { container.style.cursor = 'default'; });
  chartModals[key]._stopDrag = stopDrag;

  // ── Ctrl+Alt+드래그: Rubber Band Zoom ──
  const rbEl = document.getElementById(`cmrb-${key}`);
  let rbState = null;

  container.addEventListener('mousedown', e => {
    if (!e.ctrlKey || !e.altKey || e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    const rect = container.getBoundingClientRect();
    rbState = {
      x0: e.clientX - rect.left,
      y0: e.clientY - rect.top,
      rect
    };
    if (rbEl) {
      rbEl.style.left   = rbState.x0 + 'px';
      rbEl.style.top    = rbState.y0 + 'px';
      rbEl.style.width  = '0px';
      rbEl.style.height = '0px';
      rbEl.style.display = 'block';
    }
    container.style.cursor = 'crosshair';
  });

  document.addEventListener('mousemove', e => {
    if (!rbState) return;
    const rect  = rbState.rect;
    const x1    = e.clientX - rect.left;
    const y1    = e.clientY - rect.top;
    const left  = Math.min(rbState.x0, x1);
    const top   = Math.min(rbState.y0, y1);
    const w_    = Math.abs(x1 - rbState.x0);
    const h_    = Math.abs(y1 - rbState.y0);
    if (rbEl) {
      rbEl.style.left   = left + 'px';
      rbEl.style.top    = top  + 'px';
      rbEl.style.width  = w_   + 'px';
      rbEl.style.height = h_   + 'px';
    }
  });

  document.addEventListener('mouseup', e => {
    if (!rbState) return;
    const rect = rbState.rect;
    const x0 = rbState.x0, y0 = rbState.y0;
    const x1 = e.clientX - rect.left;
    const y1 = e.clientY - rect.top;
    if (rbEl) rbEl.style.display = 'none';
    container.style.cursor = 'default';

    const minX = Math.min(x0, x1), maxX = Math.max(x0, x1);
    const minY = Math.min(y0, y1), maxY = Math.max(y0, y1);
    rbState = null;

    // 너무 작은 선택 무시
    if (maxX - minX < 5 || maxY - minY < 5) return;

    safeChart(() => {
      // X축: 픽셀 → logical range 변환
      const cw = container.offsetWidth || 1;
      const lr = cmChart.timeScale().getVisibleLogicalRange();
      if (!lr) return;
      const span = lr.to - lr.from;
      const newFrom = lr.from + (minX / cw) * span;
      const newTo   = lr.from + (maxX / cw) * span;
      cmChart.timeScale().setVisibleLogicalRange({ from: newFrom, to: newTo });

      // Y축: 픽셀 → 가격 변환
      if (!_yRange) {
        const pr = getVisiblePriceRange();
        if (!pr) return;
        const pad = (pr.max - pr.min) * 0.05;
        _yRange = { min: pr.min - pad, max: pr.max + pad };
      }
      const ch   = container.offsetHeight || 1;
      const span_y = _yRange.max - _yRange.min;
      // Y축은 위가 높은 가격 (픽셀 반전)
      const newMax = _yRange.max - (minY / ch) * span_y;
      const newMin = _yRange.max - (maxY / ch) * span_y;
      applyYRange(newMin, newMax);
    });
  });

  // ResizeObserver
  const ro = new ResizeObserver(() => {
    safeChart(() => {
      const b = container.getBoundingClientRect();
      if (b.width > 0 && b.height > 0)
        cmChart.applyOptions({width: b.width, height: b.height});
    });
  });
  ro.observe(container);
  chartModals[key]._cmRO = ro;
}

// 차트 모달 봉 종류 변경
async function changeModalTF(cid, key, type, interval) {
  const modal = chartModals[key];
  const w     = activeWidgets[cid];
  if (!modal || !w) return;

  // 기존 예측선 + 차트 제거
  modal._disposed = true;  // fetch 진행 중인 forecast 중단용
  ['forecastSeries','validSeries','errorSeries','prophetSeries','ensembleSeries'].forEach(key => {
    if (modal[key]) {
      try { modal._cmChart?.removeSeries(modal[key]); } catch(e) {}
      modal[key] = null;
    }
  });
  if (modal._cmChart) { try { modal._cmChart.remove(); } catch(e){} }
  if (modal._cmRO)    { try { modal._cmRO.disconnect(); } catch(e){} }
  modal._disposed = false; // 새 차트 시작 전 리셋
  modal._cmChart = null;

  // 컨테이너 초기화
  const container = document.getElementById(`cmchart-${key}`);
  if (!container) return;
  container.innerHTML = '';

  // 새 데이터 로드 (일봉은 w.data 재활용 가능하지만 통일을 위해 fetch)
  try {
    const url  = `${API}/data?symbol=${w.symbol}&interval=${interval}&market=${encodeURIComponent(w.market||'')}`;
    const data = await fetch(url).then(r => r.json());
    if (!data || !data.length) return;

    // 모달에 임시로 해당 데이터 저장
    modal._modalData = data;
    modal._modalInterval = interval;

    // 차트 재생성
    initChartModalWithData(cid, key, type, data, interval);
  } catch(e) {
    console.warn('모달 봉 변경 오류:', e);
  }
}

// 차트 모달 기간 줌
function zoomChartModal(key, period, btn) {
  const modal = chartModals[key];
  if (!modal || !modal._cmChart) return;
  // active 버튼 표시
  modal.querySelectorAll('.ret-name').forEach(b => b.classList.remove('active'));
  if (btn) btn.classList.add('active');

  const ts  = modal._cmChart.timeScale();
  const now = new Date();
  const from = new Date(now);
  if      (period === '1d')  from.setDate(from.getDate() - 1);
  else if (period === '1w')  from.setDate(from.getDate() - 7);
  else if (period === '1m' || period === '1mo') from.setMonth(from.getMonth() - 1);
  else if (period === '3m' || period === '3mo') from.setMonth(from.getMonth() - 3);
  else if (period === '1y')  from.setFullYear(from.getFullYear() - 1);

  ts.setVisibleRange({
    from: Math.floor(from.getTime() / 1000),
    to:   Math.floor(now.getTime()  / 1000)
  });
}

// 차트 모달 드래그
function makeCMDraggable(modal, bar) {
  let sx=0, sy=0, ox=0, oy=0, drag=false;
  bar.addEventListener('mousedown', e => {
    // select/button/input 은 드래그 제외 (클릭 동작 보존)
    if (['SELECT','BUTTON','INPUT'].includes(e.target.tagName)) return;
    drag=true; sx=e.clientX; sy=e.clientY;
    const r=modal.getBoundingClientRect(); ox=r.left; oy=r.top;
    document.body.style.userSelect='none';
    e.preventDefault();
  });
  document.addEventListener('mousemove', e => {
    if (!drag) return;
    modal.style.left = (ox + e.clientX - sx) + 'px';
    modal.style.top  = (oy + e.clientY - sy) + 'px';
  });
  document.addEventListener('mouseup', () => {
    drag=false; document.body.style.userSelect='';
  });
}

// 차트 모달 리사이즈 (8방향)
function makeCMResizable(modal, key) {
  modal.querySelectorAll('.cm-resize').forEach(handle => {
    const dir = handle.dataset.dir;
    let sx,sy,ow,oh,ox,oy,drag=false;
    handle.addEventListener('mousedown', e => {
      e.preventDefault(); e.stopPropagation();
      drag=true; sx=e.clientX; sy=e.clientY;
      const r=modal.getBoundingClientRect();
      ow=r.width; oh=r.height; ox=r.left; oy=r.top;
      document.body.style.userSelect='none';
      document.body.style.cursor = handle.style.cursor;
    });
    document.addEventListener('mousemove', e => {
      if (!drag) return;
      const dx=e.clientX-sx, dy=e.clientY-sy;
      const MIN_W=320, MIN_H=180;
      let nw=ow, nh=oh, nl=ox, nt=oy;
      if (dir.includes('e'))  nw = Math.max(MIN_W, ow+dx);
      if (dir.includes('s'))  nh = Math.max(MIN_H, oh+dy);
      if (dir.includes('w')) { nw=Math.max(MIN_W,ow-dx); nl=ox+ow-nw; }
      if (dir.includes('n')) { nh=Math.max(MIN_H,oh-dy); nt=oy+oh-nh; }
      modal.style.width  = nw+'px'; modal.style.height = nh+'px';
      modal.style.left   = nl+'px'; modal.style.top    = nt+'px';
    });
    document.addEventListener('mouseup', () => {
      if (!drag) return;
      drag=false;
      document.body.style.userSelect='';
      document.body.style.cursor='';
    });
  });
}

/* ══════════════════════════════════
   차트 기간 줌
══════════════════════════════════ */
function zoomChart(cid, period, btn) {
  const w = activeWidgets[cid];
  if (!w || !w.chart) return;

  // 버튼 active 표시
  const bar = document.getElementById(`ret-${cid}`);
  if (bar) bar.querySelectorAll('.ret-name').forEach(b => b.classList.remove('active'));
  if (btn) btn.classList.add('active');

  const ts  = w.chart.timeScale();
  const now = new Date();

  // 기간별 from 날짜 계산
  const from = new Date(now);
  if      (period === '1d') { from.setDate(from.getDate() - 1); }
  else if (period === '1w') { from.setDate(from.getDate() - 7); }
  else if (period === '1m') { from.setMonth(from.getMonth() - 1); }
  else if (period === '3m') { from.setMonth(from.getMonth() - 3); }
  else if (period === '1y') { from.setFullYear(from.getFullYear() - 1); }

  // 분봉의 경우 오늘 하루만
  const tf = document.getElementById(`tf-${cid}`)?.value || 'day';
  if (tf !== 'day' && period === '1d') {
    // 분봉: 오늘 장 시간만 (9:00~15:30)
    const todayOpen  = new Date(now); todayOpen.setHours(9,0,0,0);
    const todayClose = new Date(now); todayClose.setHours(15,30,0,0);
    ts.setVisibleRange({
      from: Math.floor(todayOpen.getTime()  / 1000),
      to:   Math.floor(todayClose.getTime() / 1000)
    });
    return;
  }

  // 일봉: Unix timestamp (초)
  ts.setVisibleRange({
    from: Math.floor(from.getTime() / 1000),
    to:   Math.floor(now.getTime()  / 1000)
  });
}

/* ══════════════════════════════════
   대시보드 필터
══════════════════════════════════ */
// currentFilter → 전역으로 이동

function setFilter(f, el) {
  currentFilter = f;
  localStorage.setItem('dashFilter', f);
  // 드롭다운 active 표시
  document.querySelectorAll('.dfm-item').forEach(i => i.classList.remove('active'));
  if (el) el.classList.add('active');
  else {
    document.querySelectorAll('.dfm-item').forEach(i => {
      if (i.onclick && i.onclick.toString().includes(`'${f}'`)) i.classList.add('active');
    });
  }
  const lbl = document.getElementById('filter-label');
  if (lbl) lbl.textContent = FILTER_LABELS[f] || '전체보기';
  closeFilterMenu();
  applyFilter();
}

function applyFilter() {
  const symbolsInSector = _activeSector
    ? (_sectorData[_activeSector]?.stocks || [])
    : null;  // ← 섹터 활성 여부 체크

  document.querySelectorAll('#watch-cards .wc-card').forEach(card => {
    const sym  = card.dataset.symbol || '';
    const isMy = card.dataset.my === '1' || card.classList.contains('my-stock');
    const isKr = /^[0-9]{6}$/.test(sym);
    let show = true;
    if      (currentFilter === 'my')    show = isMy;
    else if (currentFilter === 'watch') show = !isMy;
    else if (currentFilter === 'kr')    show = isKr;
    else if (currentFilter === 'us')    show = !isKr;

    // ← 섹터 필터 추가 적용 (dashFilter와 AND 조건)
    if (symbolsInSector && show) {
      show = symbolsInSector.includes(sym);
    }

    card.style.display = show ? '' : 'none';
  });
  // 섹터 뷰: 필터 후 카드가 없는 그룹 숨기기
  if (_sectorViewOn) {
    document.querySelectorAll('#watch-cards .sector-group').forEach(g => {
      const hasVisible = [...g.querySelectorAll('.wc-card')].some(c => c.style.display !== 'none');
      g.style.display = hasVisible ? '' : 'none';
    });
  }
}

function toggleFilterMenu(e) {
  e.stopPropagation();
  const menu = document.getElementById('dash-filter-menu');
  const btn  = e.currentTarget;
  const rect = btn.getBoundingClientRect();
  menu.style.top   = (rect.bottom + 4) + 'px';
  menu.style.right = (window.innerWidth - rect.right) + 'px';
  menu.classList.toggle('open');
}
function closeFilterMenu() {
  document.getElementById('dash-filter-menu')?.classList.remove('open');
}

// 우클릭 메뉴
function closeCtx() { document.getElementById('ctx-menu')?.classList.remove('open'); }

document.addEventListener('click', () => { closeFilterMenu(); closeCtx(); });
// 우클릭 메뉴 — 카드 감지
let _ctxSym = '', _ctxName = '', _ctxIsMy = false;
document.addEventListener('contextmenu', e => {
  const wc = document.getElementById('watch-cards');
  if (!wc) return;
  if (!wc.contains(e.target) && e.target !== wc) return;
  e.preventDefault();
  const card   = e.target.closest('.wc-card');
  _ctxSym    = card ? (card.dataset.symbol || '') : '';
  _ctxName   = card ? (card.dataset.name   || card.querySelector('.wc-name')?.textContent?.replace('★ ','') || '') : '';
  _ctxIsMy   = card ? (card.dataset.my === '1') : false;

  // 섹터 지정
  const sectorItem = document.getElementById('ctx-sector-item');
  if (sectorItem) sectorItem.style.display = _ctxSym ? '' : 'none';

  // 내종목/관심종목 이동
  const moveMyEl    = document.getElementById('ctx-move-my');
  const moveWatchEl = document.getElementById('ctx-move-watch');
  const moveDivEl   = document.getElementById('ctx-move-divider');
  if (_ctxSym) {
    if (_ctxIsMy) {
      // 내종목 → 관심종목으로 이동
      if (moveMyEl)    { moveMyEl.style.display    = 'none'; }
      if (moveWatchEl) { moveWatchEl.style.display = ''; }
    } else {
      // 관심종목 → 내종목으로 이동
      if (moveMyEl)    { moveMyEl.style.display    = ''; }
      if (moveWatchEl) { moveWatchEl.style.display = 'none'; }
    }
    if (moveDivEl) moveDivEl.style.display = '';
  } else {
    if (moveMyEl)    moveMyEl.style.display    = 'none';
    if (moveWatchEl) moveWatchEl.style.display = 'none';
    if (moveDivEl)   moveDivEl.style.display   = 'none';
  }

  const menu = document.getElementById('ctx-menu');
  if (!menu) return;
  menu.style.left = e.clientX + 'px';
  menu.style.top  = e.clientY + 'px';
  menu.classList.add('open');
});

/* ══════════════════════════════════
   최근 가격 조회
══════════════════════════════════ */
let _rpData = [];  // 최근 가격 데이터 캐시
let _rpChartOpen = false;

async function showRecentPrices(cid) {
  const w = activeWidgets[cid];
  if (!w) return;

  const modal = document.getElementById('recent-price-modal');
  const title = document.getElementById('rp-title');
  const tbody = document.getElementById('rp-tbody');
  if (!modal || !tbody) return;

  title.textContent = `📅 ${w.name} — 최근 1개월 가격`;
  window._rpName = w.name || w.symbol;
  tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;padding:20px;color:var(--muted);">불러오는 중...</td></tr>';
  modal.classList.add('open');
  _rpChartOpen = false;
  document.getElementById('rp-bar-wrap')?.classList.remove('open');
  document.getElementById('rp-chart-btn').textContent = '📊 차트 보기';

  try {
    // w.data 에서 최근 30 영업일 추출
    let rows = [];
    if (w.data && w.data.length) {
      // 최신순 정렬
      const sorted = [...w.data].sort((a,b) => {
        const ta = typeof a.time === 'number' ? a.time : new Date(a.time).getTime()/1000;
        const tb = typeof b.time === 'number' ? b.time : new Date(b.time).getTime()/1000;
        return tb - ta;
      });
      rows = sorted.slice(0, 22); // 약 1개월
    } else {
      // 데이터 없으면 API 호출
      const isKrx = /^[0-9]{6}$/.test(w.symbol);
      const mkt   = w.market || (isKrx ? 'KOSPI' : 'US');
      const data  = await fetch(`${API}/data?symbol=${w.symbol}&interval=day&market=${encodeURIComponent(mkt)}`).then(r=>r.json());
      const sorted = [...data].sort((a,b)=>{
        const ta = typeof a.time==='number'?a.time:new Date(a.time).getTime()/1000;
        const tb = typeof b.time==='number'?b.time:new Date(b.time).getTime()/1000;
        return tb-ta;
      });
      rows = sorted.slice(0, 22);
    }

    _rpData = rows;
    renderRpTable(rows, w);
  } catch(e) {
    tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;padding:20px;color:#ef4444;">데이터 로드 실패</td></tr>';
  }
}

function renderRpTable(rows, w) {
  const tbody = document.getElementById('rp-tbody');
  const isKrx = /^[0-9]{6}$/.test(w?.symbol || '');
  tbody.innerHTML = '';

  rows.forEach((d, i) => {
    const prev   = rows[i + 1];
    const change = prev ? d.close - prev.close : 0;
    const pct    = prev && prev.close ? (change / prev.close * 100) : 0;
    const cls    = change > 0 ? 'rp-up' : change < 0 ? 'rp-down' : 'rp-flat';
    const arrow  = change > 0 ? '▲' : change < 0 ? '▼' : '━';
    const sign   = change > 0 ? '+' : '';

    const dateStr = typeof d.time === 'number'
      ? new Date(d.time * 1000).toISOString().slice(0,10)
      : String(d.time);

    const priceStr = isKrx
      ? d.close.toLocaleString('ko-KR') + '원'
      : '$' + d.close.toFixed(2);
    const chgStr = isKrx
      ? `${arrow} ${sign}${Math.abs(change).toLocaleString('ko-KR')}`
      : `${arrow} ${sign}${Math.abs(change).toFixed(2)}`;
    const pctStr = `${sign}${pct.toFixed(2)}%`;
    const volStr = d.volume ? d.volume.toLocaleString('ko-KR') : '-';

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${dateStr}</td>
      <td>${priceStr}</td>
      <td class="${cls}">${chgStr}</td>
      <td class="${cls}">${pctStr}</td>
      <td style="color:var(--muted)">${volStr}</td>`;
    tbody.appendChild(tr);
  });
}

function toggleRpChart() {
  _rpChartOpen = !_rpChartOpen;
  const wrap = document.getElementById('rp-bar-wrap');
  const btn  = document.getElementById('rp-chart-btn');
  if (_rpChartOpen) {
    wrap.classList.add('open');
    btn.textContent = '📋 표 보기';
    drawRpBarChart();
  } else {
    wrap.classList.remove('open');
    btn.textContent = '📊 차트 보기';
  }
}

function drawRpBarChart() {
  const canvas = document.getElementById('rp-bar-canvas');
  if (!canvas || !_rpData.length) return;
  const ctx = canvas.getContext('2d');
  const dpr = window.devicePixelRatio || 1;
  const W   = canvas.offsetWidth;
  const H   = 120;
  canvas.width  = W * dpr;
  canvas.height = H * dpr;
  ctx.scale(dpr, dpr);
  ctx.clearRect(0, 0, W, H);

  // 최신→과거 역순으로 그리기 (왼쪽=최신)
  const rows   = [..._rpData].reverse(); // 과거→최신
  const prices = rows.map(d => d.close);
  const minP   = Math.min(...prices);
  const maxP   = Math.max(...prices);
  const range  = maxP - minP || 1;

  const pad   = { l: 8, r: 8, t: 10, b: 20 };
  const bw    = (W - pad.l - pad.r) / rows.length;
  const chartH = H - pad.t - pad.b;

  // 바 그리기
  rows.forEach((d, i) => {
    const prev  = rows[i - 1];
    const isUp  = prev ? d.close >= prev.close : true;
    const barH  = Math.max(2, ((d.close - minP) / range) * chartH * 0.85);
    const x     = pad.l + i * bw + bw * 0.1;
    const y     = pad.t + chartH - barH;
    ctx.fillStyle = isUp ? 'rgba(239,68,68,.8)' : 'rgba(59,130,246,.8)';
    ctx.fillRect(x, y, bw * 0.8, barH);
  });

  // 날짜 레이블 (4개만)
  ctx.fillStyle = getComputedStyle(document.documentElement).getPropertyValue('--muted') || '#94a3b8';
  ctx.font = `${10 * dpr / dpr}px Arial`;
  ctx.textAlign = 'center';
  const step = Math.floor(rows.length / 4);
  [0, step, step*2, rows.length-1].forEach(i => {
    if (!rows[i]) return;
    const dateStr = typeof rows[i].time === 'number'
      ? new Date(rows[i].time*1000).toISOString().slice(5,10)
      : String(rows[i].time).slice(5,10);
    const x = pad.l + i * bw + bw / 2;
    ctx.fillText(dateStr, x, H - 4);
  });
}

function exportRecentPrices() {
  if (!_rpData.length) { showToast('⚠ 데이터 없음'); return; }
  const title = document.getElementById('rp-title')?.textContent || '최근가격';
  const wb    = XLSX.utils.book_new();
  const isKrx = /원/.test(document.querySelector('#rp-tbody td')?.textContent || '');

  const rows = [['날짜','종가','전일대비','등락률(%)','거래량']];
  _rpData.forEach((d, i) => {
    const prev   = _rpData[i + 1];
    const change = prev ? d.close - prev.close : 0;
    const pct    = prev && prev.close ? +((change / prev.close) * 100).toFixed(2) : '';
    const dateStr = typeof d.time === 'number'
      ? new Date(d.time * 1000).toISOString().slice(0,10)
      : String(d.time);
    rows.push([dateStr, d.close, prev ? +change.toFixed(2) : '', pct, d.volume || '']);
  });

  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), '최근가격');
  const today = new Date();
  const ymd   = String(today.getFullYear()).slice(2)
    + String(today.getMonth()+1).padStart(2,'0')
    + String(today.getDate()).padStart(2,'0');
  const rpName = window._rpName || '종목';
  XLSX.writeFile(wb, `${rpName}_최근가격_${ymd}.xlsx`);
  showToast('✅ 엑셀 저장됨');
}

function closeRecentPrices() {
  document.getElementById('recent-price-modal')?.classList.remove('open');
  _rpData = [];
}

// 모달 바깥 클릭 시 닫기 (DOM 로드 후)
window.addEventListener('DOMContentLoaded', () => {
  document.getElementById('recent-price-modal')?.addEventListener('click', function(e) {
    if (e.target === this) closeRecentPrices();
  });
});

/* ══════════════════════════════════
   골든/데드 크로스 알림
══════════════════════════════════ */
function toggleCrossAlert() {
  const isOn = localStorage.getItem('crossAlert') === 'on';
  const next = isOn ? 'off' : 'on';
  localStorage.setItem('crossAlert', next);
  updateCrossToggleUI();
  if (next === 'on') {
    loadAllCrossAlerts();
  } else {
    // 뱃지 전부 제거
    document.querySelectorAll('.cross-badges').forEach(el => el.innerHTML = '');
  }
}

function updateCrossToggleUI() {
  const btn = document.getElementById('cross-toggle-btn');
  if (!btn) return;
  const isOn = localStorage.getItem('crossAlert') === 'on';
  btn.textContent = isOn ? '📡 크로스 알림 ON' : '📡 크로스 알림 OFF';
  btn.classList.toggle('on', isOn);
}

function calcMA(data, period) {
  // data: [{close},...] 최신순 정렬 (인덱스 0 = 가장 최근)
  const result = [];
  for (let i = 0; i <= data.length - period; i++) {
    let sum = 0;
    for (let j = i; j < i + period; j++) sum += data[j].close;
    result.push(sum / period);
  }
  return result; // result[0] = 최근 MA값
}

async function checkCross(symbol, market) {
  try {
    const isKrx = /^[0-9]{6}$/.test(symbol);
    const mkt   = market || (isKrx ? 'KOSPI' : 'US');
    const url   = `${API}/data?symbol=${symbol}&interval=day&market=${encodeURIComponent(mkt)}`;
    const raw   = await fetch(url).then(r => r.json());
    if (!raw || raw.length < 65) return [];

    // 최신순으로 정렬
    const data = [...raw].sort((a,b) => b.time - a.time);

    const badges = [];
    // 단기: 5MA vs 20MA
    const ma5  = calcMA(data, 5);
    const ma20 = calcMA(data, 20);
    // 중기: 20MA vs 60MA
    const ma60 = calcMA(data, 60);

    // 3일 이내 크로스 체크 (인덱스 0=오늘, 1=어제, 2=그제, 3=3일전)
    for (let i = 0; i <= 2; i++) {
      if (ma5[i] === undefined || ma5[i+1] === undefined) break;
      // 단기 골든: 오늘 5>20, 어제 5<=20
      if (ma5[i] > ma20[i] && ma5[i+1] <= ma20[i+1])
        badges.push({cls:'cb-golden', text:'단골'});
      // 단기 데드: 오늘 5<20, 어제 5>=20
      if (ma5[i] < ma20[i] && ma5[i+1] >= ma20[i+1])
        badges.push({cls:'cb-dead', text:'단데'});
    }
    for (let i = 0; i <= 2; i++) {
      if (ma20[i] === undefined || ma20[i+1] === undefined) break;
      if (ma60[i] === undefined || ma60[i+1] === undefined) break;
      if (ma20[i] > ma60[i] && ma20[i+1] <= ma60[i+1])
        badges.push({cls:'cb-golden', text:'중골'});
      if (ma20[i] < ma60[i] && ma20[i+1] >= ma60[i+1])
        badges.push({cls:'cb-dead', text:'중데'});
    }
    return badges;
  } catch(e) { return []; }
}

async function loadAllCrossAlerts() {
  const allItems = [...getMyList(), ...getList()];
  await Promise.allSettled(allItems.map(async item => {
    const el = document.getElementById(`cb-${item.symbol}`);
    if (!el) return;
    const badges = await checkCross(item.symbol, item.market);
    el.innerHTML = badges.map(b =>
      `<span class="cross-badge ${b.cls}">${b.text}</span>`
    ).join('');
  }));
}

// dragFrom: {idx, listType('my'|'watch')}
let dragState = null;

function bindDrag(el, idx, listType) {
  el.draggable = true;
  el.addEventListener('dragstart', function(e) {
    dragState = { idx, listType };
    this.classList.add('dragging');
    e.dataTransfer.effectAllowed = 'move';
  });
  el.addEventListener('dragover', function(e) {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    this.classList.add('over');
  });
  el.addEventListener('dragleave', function() { this.classList.remove('over'); });
  el.addEventListener('dragend',   function() {
    this.classList.remove('dragging');
    document.querySelectorAll('.watch-item').forEach(i => i.classList.remove('over'));
  });
  el.addEventListener('drop', function(e) {
    e.stopPropagation(); e.preventDefault();
    this.classList.remove('over');
    if (!dragState) return;
    const from = dragState;
    const toIdx = idx, toType = listType;
    if (from.idx === toIdx && from.listType === toType) return;

    const myL = getMyList(), wL = getList();
    const srcList  = from.listType === 'my' ? myL : wL;
    const dstList  = toType        === 'my' ? myL : wL;
    const [moved]  = srcList.splice(from.idx, 1);
    dstList.splice(toIdx, 0, moved);
    saveMyList(myL); saveList(wL);
    dragState = null;
    loadWatch();
  });
}

// 컬럼 빈 영역 드롭 지원 (맨 끝에 추가)
function bindColDrop(colBody, listType) {
  colBody.addEventListener('dragover', e => e.preventDefault());
  colBody.addEventListener('drop', function(e) {
    if (e.target !== this) return; // 아이템 위면 아이템 drop이 처리
    e.preventDefault();
    if (!dragState) return;
    const from = dragState;
    const myL = getMyList(), wL = getList();
    const srcList = from.listType === 'my' ? myL : wL;
    const dstList = listType      === 'my' ? myL : wL;
    if (from.listType === listType) return; // 같은 컬럼 끝으로는 아이템 drop이 처리
    const [moved] = srcList.splice(from.idx, 1);
    dstList.push(moved);
    saveMyList(myL); saveList(wL);
    dragState = null;
    currentFilter = localStorage.getItem('dashFilter') || 'all';
  loadWatch();
  updateCrossToggleUI();
  const lbl = document.getElementById('filter-label');
  if (lbl) lbl.textContent = FILTER_LABELS[currentFilter] || '전체보기';
  });
}

function mktCls(m) { return ((m||'기타').toLowerCase().replace(/[^a-z가-힣]/g,'') || '기타'); }

/* ══════════════════════════════════
   검색
══════════════════════════════════ */
let stimer = null;
function debounceSearch() { clearTimeout(stimer); stimer = setTimeout(doSearch, 280); }
async function doSearch() {
  const q   = document.getElementById('search').value.trim();
  const box = document.getElementById('results');
  box.innerHTML = '';
  if (!q) return;
  try {
    const data = await fetch(`${API}/search?q=${encodeURIComponent(q)}`).then(r => r.json());
    data.forEach(s => {
      const mc = mktCls(s.market);
      const d  = document.createElement('div');
      d.className = 'si';
      d.innerHTML = `<span class="market-badge ${mc}">${s.market}</span><span class="si-name">${s.name}</span><span class="si-code">${s.symbol}</span>`;
      d.onclick = () => {
        curSym = s.symbol; curMkt = s.market; curName = s.name;
        document.getElementById('search').value = s.name;
        box.innerHTML = '';
        // 추가버튼 영역 표시
        const addBtns = document.getElementById('sb-add-btns');
        const selName = document.getElementById('add-sel-name');
        if (addBtns) { addBtns.classList.add('visible'); }
        if (selName) selName.textContent = s.name;
      };
      box.appendChild(d);
    });
  } catch(e) { showToast('서버 연결 실패. app.py를 확인하세요.'); }
}

/* ══════════════════════════════════
   팝업 윈도우 시스템
══════════════════════════════════ */
function openPopup(name, symbol, market = '기타') {
  // 이미 열려 있는 경우 복원 + 앞으로
  const existing = Object.entries(activeWidgets).find(([,w]) => w.symbol === symbol);
  if (existing) {
    const [cid, w] = existing;
    restorePopup(cid);
    bringToFront(cid);
    return;
  }

  const cid = `chart-${symbol}-${Date.now()}`;
  const mc  = mktCls(market);
  const memo = localStorage.getItem(`memo_${symbol}`) || '';

  /* 팝업 DOM */
  const popup = document.createElement('div');
  popup.className = 'cpopup';
  popup.id = `popup-${cid}`;

  // 열린 창 수에 따라 약간씩 오프셋
  const n   = Object.keys(activeWidgets).length;
  const off = n * 28;
  const W   = Math.min(900, window.innerHeight * 0.90);
  const H   = W;
  const left = Math.max(0, (window.innerWidth  - W) / 2 + off);
  const top  = Math.max(50, (window.innerHeight - H) / 2 + off);
  popup.style.cssText = `left:${left}px;top:${top}px;width:${W}px;height:${H}px;`;

  popup.innerHTML = `
    <!-- 리사이즈 핸들 8방향 -->
    <div class="rs rs-n"  data-dir="n"></div>
    <div class="rs rs-s"  data-dir="s"></div>
    <div class="rs rs-e"  data-dir="e"></div>
    <div class="rs rs-w"  data-dir="w"></div>
    <div class="rs rs-ne" data-dir="ne"></div>
    <div class="rs rs-nw" data-dir="nw"></div>
    <div class="rs rs-se" data-dir="se"></div>
    <div class="rs rs-sw" data-dir="sw"></div>
    <!-- 타이틀바 -->
    <div class="win-bar" id="bar-${cid}">
      <div class="win-btns">
        <button class="wbtn wbtn-close" onclick="closePopup('${cid}')"   title="닫기">✕</button>
        <button class="wbtn wbtn-min"   onclick="minimizePopup('${cid}')" title="최소화">─</button>
        <button class="wbtn wbtn-max"   onclick="maximizePopup('${cid}')" title="최대화/복원">□</button>
      </div>
      <div class="win-title">
        <span class="market-badge ${mc}">${market}</span>
        <b>${name}</b>
        <span class="sym">${symbol}</span>
      </div>
      <div class="combo-signals" id="combo-${cid}"></div>
    </div>
    <!-- 컨트롤 -->
    <!-- ① 컨트롤 바: 봉종류 | RSI | 신호 | → | 호가창/체결/투자자 | 시총 -->
    <div class="win-ctrl">
      <select id="tf-${cid}" onchange="changeTF('${cid}',this.value)">
        <option value="day">일봉</option>
        <option value="1m">1분봉</option>
        <option value="5m">5분봉</option>
        <option value="15m">15분봉</option>
        <option value="30m">30분봉</option>
      </select>
      <span id="rsi-${cid}" class="badge b-hold">RSI: -</span>
      <span id="sig-${cid}" class="badge b-hold">분석중...</span>
      <div class="win-ctrl-right">
        <button class="kis-ob-btn" id="kis-ob-btn-${cid}" onclick="toggleOrderbook('${cid}')">호가창</button>
        <button class="kis-ob-btn" id="kis-tick-btn-${cid}" onclick="toggleTickStream('${cid}')">체결</button>
        <button class="kis-ob-btn" id="kis-inv-btn-${cid}" onclick="toggleInvestor('${cid}')">투자자</button>
        <span class="win-ctrl-mcap" id="kis-mktcap-${cid}" style="display:none"></span>
        <button class="kis-ob-btn" onclick="showRecentPrices('${cid}')">📅 최근가격</button>
        <button class="export-btn" onclick="exportChart('${cid}')">📊 엑셀</button>
		      <button class="btn-ai-analyze" onclick="analyzeChartWithAi('${cid}')">
        🤖 AI분석
      </button>
      </div>
    </div>
    <!-- ② 기간 수익률 -->
    <div class="returns-bar" id="ret-${cid}">
      <span class="ret-label">기간 수익률</span>
      <span class="ret-item"><button class="ret-name" onclick="zoomChart('${cid}','1d',this)">오늘</button><span class="ret-val" id="ret-d-${cid}">-</span></span>
      <span class="ret-item"><button class="ret-name" onclick="zoomChart('${cid}','1w',this)">1주</button><span class="ret-val" id="ret-w-${cid}">-</span></span>
      <span class="ret-item"><button class="ret-name" onclick="zoomChart('${cid}','1m',this)">1개월</button><span class="ret-val" id="ret-m-${cid}">-</span></span>
      <span class="ret-item"><button class="ret-name" onclick="zoomChart('${cid}','3m',this)">3개월</button><span class="ret-val" id="ret-3m-${cid}">-</span></span>
      <span class="ret-item"><button class="ret-name" onclick="zoomChart('${cid}','1y',this)">1년</button><span class="ret-val" id="ret-y-${cid}">-</span></span>
    </div>
    <!-- ③ KIS 현재가 + 펀더멘털 (한국주식) -->
    <div class="kis-panel" id="kis-panel-${cid}" style="display:none">
      <div class="kis-panel-inner">
        <!-- 왼쪽: 현재가행 + 시가고가저가행 -->
        <div style="display:flex;flex-direction:column;justify-content:center;flex-shrink:0;padding:3px 8px 3px 10px;min-width:0;">
          <div style="display:flex;align-items:center;gap:6px;margin-bottom:1px;">
            <span class="kis-dot" id="kis-dot-${cid}"></span>
            <span class="kis-price-main" id="kis-price-${cid}">-</span>
            <span class="kis-change" id="kis-change-${cid}">-</span>
          </div>
          <div class="kis-info-row" id="kis-info-${cid}"></div>
        </div>
        <!-- 오른쪽: 펀더멘털 탭 -->
        <div class="kis-fund-panel" id="kis-fund-${cid}" style="display:none;flex:1;border-left:1px solid var(--border);min-width:0;height:100%;flex-direction:column;">
        <div class="kis-fund-tabs">
          <button class="kis-fund-tab active" onclick="fundTab('${cid}',0,this)">밸류</button>
          <button class="kis-fund-tab" onclick="fundTab('${cid}',1,this)">실적</button>
          <button class="kis-fund-tab" onclick="fundTab('${cid}',2,this)">수익성</button>
          <button class="kis-fund-tab" onclick="fundTab('${cid}',3,this)">현금흐름</button>
		  <button class="kis-fund-tab" 
		  style="margin-left:auto; background:rgba(139,92,246,.15); color:#8b5cf6; border-color:#8b5cf6;"
          onclick="fundamentalAnalysis('${cid}')">🔍 기본분석</button>
        </div>
        <div class="kis-fund-page active" id="kf-p0-${cid}">
          <div class="kis-fund-item"><span class="kis-fund-label">시총</span><span class="kis-fund-val" id="kf-mcap-${cid}">-</span></div>
          <div class="kis-fund-item"><span class="kis-fund-label">PER</span><span class="kis-fund-val" id="kf-per-${cid}">-</span></div>
          <div class="kis-fund-item"><span class="kis-fund-label">PER(F)</span><span class="kis-fund-val" id="kf-perf-${cid}">-</span></div>
          <div class="kis-fund-item"><span class="kis-fund-label">PBR</span><span class="kis-fund-val" id="kf-pbr-${cid}">-</span></div>
          <div class="kis-fund-item"><span class="kis-fund-label">PSR</span><span class="kis-fund-val" id="kf-psr-${cid}">-</span></div>
          <div class="kis-fund-item"><span class="kis-fund-label">DY</span><span class="kis-fund-val" id="kf-div-${cid}">-</span></div>
        </div>
        <div class="kis-fund-page" id="kf-p1-${cid}">
          <div class="kis-fund-item"><span class="kis-fund-label">매출</span><span class="kis-fund-val" id="kf-rev-${cid}">-</span></div>
          <div class="kis-fund-item"><span class="kis-fund-label">RevG</span><span class="kis-fund-val" id="kf-revg-${cid}">-</span></div>
          <div class="kis-fund-item"><span class="kis-fund-label">EPS</span><span class="kis-fund-val" id="kf-eps-${cid}">-</span></div>
          <div class="kis-fund-item"><span class="kis-fund-label">EPS(F)</span><span class="kis-fund-val" id="kf-epsf-${cid}">-</span></div>
          <div class="kis-fund-item"><span class="kis-fund-label">EarnG</span><span class="kis-fund-val" id="kf-eg-${cid}">-</span></div>
        </div>
        <div class="kis-fund-page" id="kf-p2-${cid}">
          <div class="kis-fund-item"><span class="kis-fund-label">OPM</span><span class="kis-fund-val" id="kf-om-${cid}">-</span></div>
          <div class="kis-fund-item"><span class="kis-fund-label">NPM</span><span class="kis-fund-val" id="kf-pm-${cid}">-</span></div>
          <div class="kis-fund-item"><span class="kis-fund-label">ROE</span><span class="kis-fund-val" id="kf-roe-${cid}">-</span></div>
          <div class="kis-fund-item"><span class="kis-fund-label">ROA</span><span class="kis-fund-val" id="kf-roa-${cid}">-</span></div>
        </div>
        <div class="kis-fund-page" id="kf-p3-${cid}">
          <div class="kis-fund-item"><span class="kis-fund-label">FCF</span><span class="kis-fund-val" id="kf-fcf-${cid}">-</span></div>
          <div class="kis-fund-item"><span class="kis-fund-label">OCF</span><span class="kis-fund-val" id="kf-ocf-${cid}">-</span></div>
          <div class="kis-fund-item"><span class="kis-fund-label">52H</span><span class="kis-fund-val" id="kf-52h-${cid}">-</span></div>
          <div class="kis-fund-item"><span class="kis-fund-label">52L</span><span class="kis-fund-val" id="kf-52l-${cid}">-</span></div>
        </div>
      </div>
        </div><!-- fund-panel 끝 -->
      </div><!-- flex-row 끝 -->
    </div><!-- kis-panel 끝 -->
    <!-- 차트 분할 패널 -->
    <div class="win-body" id="wb-${cid}" style="flex:1;min-width:0;min-height:0;overflow:hidden;">
      <!-- 메인 차트 (캔들 + MA + BB + 거래량) -->
      <div class="chart-pane chart-pane-main" id="${cid}">
        <button class="pane-popout" onclick="popoutChart('${cid}','candle')">⤢</button>
        <span class="pane-label">캔들차트</span>
        <div class="chart-legend">
          <div class="legend-title">📊 지표</div>
          <div class="legend-item"><span style="color:#f59e0b;font-weight:700;margin-right:4px;">━</span>20MA &amp; BB</div>
          <label class="legend-item"><input type="checkbox" checked onchange="toggleMA('${cid}','ma60',this.checked)">
            <span style="color:#3b82f6;font-weight:700;margin:0 4px;">━</span>60MA</label>
          <label class="legend-item"><input type="checkbox" checked onchange="toggleMA('${cid}','ma120',this.checked)">
            <span style="color:#a855f7;font-weight:700;margin:0 4px;">━</span>120MA</label>
        </div>
      </div>
      <!-- 구분선 1 -->
      <div class="chart-divider" id="div1-${cid}" data-cid="${cid}" data-divider="1"></div>
      <!-- MACD 차트 -->
      <div class="chart-pane chart-pane-macd" id="macd-${cid}">
        <button class="pane-popout" onclick="popoutChart('${cid}','macd')">⤢</button>
        <span class="pane-label">MACD &nbsp;<span style="color:#2962FF;">━</span>&nbsp;<span style="color:#FF6D00;">━</span></span>
      </div>
      <!-- 구분선 2 -->
      <div class="chart-divider" id="div2-${cid}" data-cid="${cid}" data-divider="2"></div>
      <!-- RSI 차트 -->
      <div class="chart-pane chart-pane-rsi" id="rsi-pane-${cid}">
        <button class="pane-popout" onclick="popoutChart('${cid}','rsi')">⤢</button>
        <span class="pane-label">RSI(14)</span>
      </div>
    </div>
    </div><!-- win-body 끝 -->
    <!-- 투자자별 매매동향 패널 -->
    <div class="investor-panel" id="inv-panel-${cid}">
      <div class="inv-header">
        <span class="inv-title">👥 투자자별 매매동향</span>
        <div class="inv-tabs">
          <button class="inv-tab active" onclick="switchInvTab('${cid}','day',this)">당일</button>
          <button class="inv-tab"        onclick="switchInvTab('${cid}','week',this)">1주</button>
          <button class="inv-tab"        onclick="switchInvTab('${cid}','month',this)">1개월</button>
        </div>
        <button class="inv-close" onclick="toggleInvestor('${cid}')">✕</button>
      </div>
      <div id="inv-body-${cid}"><div class="inv-loading">불러오는 중...</div></div>
    </div>
    <!-- 메모 -->
    <div class="win-memo">
      <span style="font-size:12px;font-weight:700;color:var(--muted);">📝</span>
      <textarea id="memo-${cid}" onkeyup="saveMemo('${symbol}',this.value)"
        placeholder="${name} — 매매 시나리오, 진입/청산 목표가 등 (자동저장)"></textarea>
    </div>`;

  document.body.appendChild(popup);
  document.getElementById(`memo-${cid}`).value = memo;

  // 드래그 + 리사이즈
  makeDraggable(popup, document.getElementById(`bar-${cid}`), cid);
  makeResizable(popup, cid);

  // 클릭 시 앞으로
  popup.addEventListener('mousedown', () => bringToFront(cid));

  // 차트 초기화 (3개 독립 차트 + 시간축 동기화)
  const bg = cssVar('--cbg'), tc = cssVar('--ctxt'), gl = cssVar('--gcol');
  const chartOpts = {
    layout:{ background:{color:bg}, textColor:tc },
    grid:{ vertLines:{color:gl}, horzLines:{color:gl} },
    crosshair:{ mode:1 },
    timeScale:{ timeVisible:true, secondsVisible:false, borderColor:cssVar('--border') }
  };

  // Y축 너비를 세 차트 동일하게 고정 (정렬)
  const PRICE_SCALE_W = 80;

  // ① 메인 차트 (캔들 + MA + BB + 거래량)
  const chart = LightweightCharts.createChart(document.getElementById(cid), {
    ...chartOpts,
    rightPriceScale:{ scaleMargins:{top:.05, bottom:.20}, minimumWidth:PRICE_SCALE_W },
    timeScale:{ ...chartOpts.timeScale, visible:false }
  });
  const candleSeries = chart.addCandlestickSeries({borderVisible:false});
  const ma20Series   = chart.addLineSeries({lineWidth:2,  color:'#f59e0b', title:'20선'});
  const ma60Series   = chart.addLineSeries({lineWidth:1.5,color:'#3b82f6', title:'60선'});
  const ma120Series  = chart.addLineSeries({lineWidth:1.5,color:'#a855f7', title:'120선'});
  const bbUpSeries   = chart.addLineSeries({lineWidth:1,  color:'#94a3b8', lineStyle:2});
  const bbLowSeries  = chart.addLineSeries({lineWidth:1,  color:'#94a3b8', lineStyle:2});
  const volumeSeries = chart.addHistogramSeries({priceFormat:{type:'volume'}, priceScaleId:'vol'});
  chart.priceScale('vol').applyOptions({scaleMargins:{top:.75,bottom:.0},visible:false});

  // ② MACD 차트
  // ② MACD 차트
  const macdChart = LightweightCharts.createChart(document.getElementById(`macd-${cid}`), {
    ...chartOpts,
    rightPriceScale:{ scaleMargins:{top:.1, bottom:.1}, minimumWidth:PRICE_SCALE_W },
    timeScale:{ ...chartOpts.timeScale, visible:false }
  });
  const macdHistSeries   = macdChart.addHistogramSeries();
  const macdLineSeries   = macdChart.addLineSeries({color:'#2962FF',lineWidth:1.5,title:'MACD'});
  const signalLineSeries = macdChart.addLineSeries({color:'#FF6D00',lineWidth:1.5,title:'Signal'});

  // ③ RSI 차트
  const rsiChart = LightweightCharts.createChart(document.getElementById(`rsi-pane-${cid}`), {
    ...chartOpts,
    rightPriceScale:{ scaleMargins:{top:.1, bottom:.1}, minimumWidth:PRICE_SCALE_W },
    timeScale:{ ...chartOpts.timeScale, visible:true }
  });
  const rsiSeries    = rsiChart.addLineSeries({color:'#a855f7',lineWidth:1.5,title:'RSI'});
  const rsiOb        = rsiChart.addLineSeries({color:'#ef4444',lineWidth:1,lineStyle:2});
  const rsiOs        = rsiChart.addLineSeries({color:'#22c55e',lineWidth:1,lineStyle:2});

  // 시간축 동기화 (세 차트가 함께 스크롤/줌)
  function syncTimeRange(src, targets) {
    let syncing = false;
    src.timeScale().subscribeVisibleLogicalRangeChange(range => {
      if (syncing || !range) return;
      syncing = true;
      targets.forEach(t => t.timeScale().setVisibleLogicalRange(range));
      syncing = false;
    });
  }
  syncTimeRange(chart,     [macdChart, rsiChart]);
  syncTimeRange(macdChart, [chart,     rsiChart]);
  syncTimeRange(rsiChart,  [chart,     macdChart]);

  activeWidgets[cid] = {
    chart, macdChart, rsiChart,
    candleSeries, volumeSeries, ma20Series, ma60Series, ma120Series,
    bbUpSeries, bbLowSeries, macdHistSeries, macdLineSeries, signalLineSeries,
    rsiSeries, rsiOb, rsiOs,
    name, symbol, market, interval:'day', data:null, lastSignalType:null,
    popupEl:popup, minimized:false, maximized:false, savedRect:null,
    // KIS 실시간
    kisTimer:null, kisObTimer:null, kisWs:null,
    obVisible:false, tickVisible:false, lastPrice:0,
    invVisible:false, invPeriod:'day'
  };

  // 한국 주식이면 KIS 실시간 패널 활성화
  const isKrx = !symbol.includes('.') && /^[0-9]{6}$/.test(symbol);
  const isUs  = market === 'S&P500' || market === 'NYSE' || market === 'NASDAQ' ||
                market === 'ETF(미국)' || symbol.includes('.') === false && !isKrx &&
                /^[A-Z]{1,5}$/.test(symbol);

  if (isKrx) {
    document.getElementById(`kis-panel-${cid}`).style.display = 'flex';
    startKisRealtime(cid, symbol);
  } else if (isUs) {
    // 미국 주식: 현재가 바 표시 (호가/체결 버튼 포함), KIS 전용 버튼 숨김
    const panel = document.getElementById(`kis-panel-${cid}`);
    panel.style.display = 'flex';
    document.getElementById(`kis-dot-${cid}`).style.display    = 'none';
    document.getElementById(`kis-inv-btn-${cid}`).style.display = 'none';
    // 미국 현재가 폴링 시작 (15초마다)
    startUsRealtime(cid, symbol);
  }

  // 분할 드래그 초기화
  initChartDivider(cid);

  // ResizeObserver 등록 (3개 모두)
  ro.observe(document.getElementById(cid));
  ro.observe(document.getElementById(`macd-${cid}`));
  ro.observe(document.getElementById(`rsi-pane-${cid}`));

  applyCandleColors(activeWidgets[cid]);
  bringToFront(cid);
  addTaskbarChip(cid, name, symbol);
  fetchChart(cid, symbol, 'day', true);
}

/* ── 창 컨트롤 ── */
function bringToFront(cid) {
  // 다른 창들 포커스 해제
  document.querySelectorAll('.cpopup').forEach(p => p.classList.remove('focused'));
  const popup = document.getElementById(`popup-${cid}`);
  if (!popup) return;
  popup.style.zIndex = ++zBase;
  popup.classList.add('focused');
  // 태스크바 chip 활성
  document.querySelectorAll('.tb-chip').forEach(c => c.classList.remove('active'));
  const chip = document.getElementById(`chip-${cid}`);
  if (chip) chip.classList.add('active');
}

function minimizePopup(cid) {
  const w = activeWidgets[cid];
  if (!w) return;
  w.minimized = true;
  w.popupEl.classList.add('minimized');
  // 태스크바 chip 비활성
  document.querySelectorAll('.cpopup').forEach(p => p.classList.remove('focused'));
  const chip = document.getElementById(`chip-${cid}`);
  if (chip) { chip.classList.remove('active'); }
}

function restorePopup(cid) {
  const w = activeWidgets[cid];
  if (!w) return;
  w.minimized = false;
  w.popupEl.classList.remove('minimized');
  bringToFront(cid);
  setTimeout(() => resizeAllCharts(cid), 60);
}

function maximizePopup(cid) {
  const w = activeWidgets[cid]; if (!w) return;
  const popup = w.popupEl;
  if (!w.maximized) {
    w.savedRect = { left:popup.style.left, top:popup.style.top, width:popup.style.width, height:popup.style.height };
    popup.classList.add('maximized');
    popup.querySelectorAll('.rs').forEach(h => h.style.display = 'none');
    w.maximized = true;
  } else {
    popup.classList.remove('maximized');
    popup.querySelectorAll('.rs').forEach(h => h.style.display = '');
    if (w.savedRect) {
      popup.style.left   = w.savedRect.left;
      popup.style.top    = w.savedRect.top;
      popup.style.width  = w.savedRect.width;
      popup.style.height = w.savedRect.height;
    }
    w.maximized = false;
  }
  setTimeout(() => resizeAllCharts(cid), 60);
  bringToFront(cid);
}

function closePopup(cid) {
  const el = document.getElementById(cid); if (!el) return;
  ro.unobserve(el);
  const mel = document.getElementById(`macd-${cid}`);     if(mel) ro.unobserve(mel);
  const rel = document.getElementById(`rsi-pane-${cid}`); if(rel) ro.unobserve(rel);
  const w = activeWidgets[cid];
  if (w) {
    if (w.kisTimer)   clearInterval(w.kisTimer);
    if (w.kisObTimer) clearInterval(w.kisObTimer);
    if (w.kisWs)      w.kisWs.close();
    if (obModalCid === cid) closeObModal();
    if (w.chart)     w.chart.remove();
    if (w.macdChart) w.macdChart.remove();
    if (w.rsiChart)  w.rsiChart.remove();
    if (w.popupEl)   w.popupEl.remove();
  }
  delete activeWidgets[cid];
  removeTaskbarChip(cid);
}

function resizeAllCharts(cid) {
  const w = activeWidgets[cid]; if (!w) return;
  const panes = document.querySelectorAll(`#wb-${cid} .chart-pane`);
  panes.forEach(p => {
    const W = p.clientWidth, H = p.clientHeight;
    if (p.id === cid         && w.chart)     w.chart.applyOptions({width:W,height:H});
    if (p.id === `macd-${cid}`     && w.macdChart) w.macdChart.applyOptions({width:W,height:H});
    if (p.id === `rsi-pane-${cid}` && w.rsiChart)  w.rsiChart.applyOptions({width:W,height:H});
  });
}

/* ── 태스크바 칩 ── */
function addTaskbarChip(cid, name, symbol) {
  const tb   = document.getElementById('taskbar');
  const empty = document.getElementById('tb-empty');
  if (empty) empty.style.display = 'none';
  const chip = document.createElement('div');
  chip.className = 'tb-chip active';
  chip.id = `chip-${cid}`;
  chip.innerHTML = `<span class="tb-chip-name">📈 ${name}</span><span class="tb-chip-x" id="chipx-${cid}">✕</span>`;
  chip.onclick = e => {
    if (e.target.id === `chipx-${cid}`) { closePopup(cid); return; }
    const w = activeWidgets[cid]; if (!w) return;
    if (w.minimized) restorePopup(cid);
    else { minimizePopup(cid); }
  };
  tb.appendChild(chip);
}
function removeTaskbarChip(cid) {
  const chip = document.getElementById(`chip-${cid}`);
  if (chip) chip.remove();
  if (!document.querySelectorAll('.tb-chip').length) {
    const e = document.getElementById('tb-empty');
    if (e) e.style.display = '';
  }
}

/* ── 드래그 이동 ── */
function makeDraggable(popup, handle, cid) {
  let sx, sy, ox, oy, dragging = false;
  handle.addEventListener('mousedown', e => {
    // 버튼 클릭은 드래그 무시
    if (e.target.closest('.wbtn')) return;
    const w = activeWidgets[cid];
    if (w && w.maximized) return; // 최대화 중엔 드래그 없음
    e.preventDefault();
    dragging = true;
    const r = popup.getBoundingClientRect();
    ox = r.left; oy = r.top; sx = e.clientX; sy = e.clientY;
    popup.style.transform = 'none';
    popup.style.left = ox + 'px'; popup.style.top = oy + 'px';
  });
  document.addEventListener('mousemove', e => {
    if (!dragging) return;
    popup.style.left = (ox + e.clientX - sx) + 'px';
    popup.style.top  = (oy + e.clientY - sy) + 'px';
  });
  document.addEventListener('mouseup', () => { dragging = false; });
}

function makeResizable(popup, cid) {
  const MIN_W = 420, MIN_H = 300;
  let resizing = false, dir = '';
  let startX, startY, startW, startH, startL, startT;

  popup.querySelectorAll('.rs').forEach(handle => {
    handle.addEventListener('mousedown', e => {
      const w = activeWidgets[cid];
      if (w && w.maximized) return;
      e.preventDefault();
      e.stopPropagation();
      resizing = true;
      dir      = handle.dataset.dir;
      startX   = e.clientX;
      startY   = e.clientY;
      const r  = popup.getBoundingClientRect();
      startW   = r.width;
      startH   = r.height;
      startL   = r.left;
      startT   = r.top;
      popup.style.transform = 'none';
      popup.style.left  = startL + 'px';
      popup.style.top   = startT + 'px';
      popup.style.width = startW + 'px';
      popup.style.height= startH + 'px';
      bringToFront(cid);
    });
  });

  document.addEventListener('mousemove', e => {
    if (!resizing) return;
    const dx = e.clientX - startX;
    const dy = e.clientY - startY;
    let newW = startW, newH = startH, newL = startL, newT = startT;

    if (dir.includes('e'))  newW = Math.max(MIN_W, startW + dx);
    if (dir.includes('s'))  newH = Math.max(MIN_H, startH + dy);
    if (dir.includes('w')) { newW = Math.max(MIN_W, startW - dx); newL = startL + (startW - newW); }
    if (dir.includes('n')) { newH = Math.max(MIN_H, startH - dy); newT = startT + (startH - newH); }

    popup.style.width  = newW + 'px';
    popup.style.height = newH + 'px';
    popup.style.left   = newL + 'px';
    popup.style.top    = newT + 'px';

    // 차트 즉시 리사이즈
    setTimeout(() => resizeAllCharts(cid), 0);
  });

  document.addEventListener('mouseup', () => {
    if (resizing) {
      resizing = false;
      resizeAllCharts(cid);
    }
  });
}

/* ── ResizeObserver: 3개 차트 모두 ── */
const ro = new ResizeObserver(entries => {
  for (const e of entries) {
    const id  = e.target.id;
    const W   = e.contentRect.width;
    const H   = e.contentRect.height;
    // 메인 차트
    const wm = activeWidgets[id];
    if (wm && wm.chart) wm.chart.applyOptions({width:W, height:H});
    // MACD 차트 (macd-{cid})
    if (id.startsWith('macd-')) {
      const cid2 = id.slice(5);
      const ww = activeWidgets[cid2];
      if (ww && ww.macdChart) ww.macdChart.applyOptions({width:W, height:H});
    }
    // RSI 차트 (rsi-pane-{cid})
    if (id.startsWith('rsi-pane-')) {
      const cid2 = id.slice(9);
      const ww = activeWidgets[cid2];
      if (ww && ww.rsiChart) ww.rsiChart.applyOptions({width:W, height:H});
    }
  }
});

/* ── 차트 분할 드래그 ── */
function initChartDivider(cid) {
  [1, 2].forEach(n => {
    const div = document.getElementById(`div${n}-${cid}`);
    if (!div) return;
    let startY, startA, startB;

    div.addEventListener('mousedown', e => {
      e.preventDefault();
      div.classList.add('dragging');
      startY = e.clientY;
      const wb = document.getElementById(`wb-${cid}`);
      const panes = wb.querySelectorAll('.chart-pane');
      // div1 → pane[0]↔pane[1], div2 → pane[1]↔pane[2]
      const idx = n - 1;
      startA = panes[idx].getBoundingClientRect().height;
      startB = panes[idx+1].getBoundingClientRect().height;

      const onMove = me => {
        const delta = me.clientY - startY;
        const newA  = Math.max(50, startA + delta);
        const newB  = Math.max(50, startB - delta);
        panes[idx].style.flex   = 'none';
        panes[idx].style.height = newA + 'px';
        panes[idx+1].style.flex   = 'none';
        panes[idx+1].style.height = newB + 'px';
        // 차트 크기 즉시 반영
        const w = activeWidgets[cid];
        if (w) {
          const pw = panes[0].clientWidth;
          if (idx === 0 && w.chart)     w.chart.applyOptions({width:pw, height:newA});
          if (idx === 0 && w.macdChart) w.macdChart.applyOptions({width:pw, height:newB});
          if (idx === 1 && w.macdChart) w.macdChart.applyOptions({width:pw, height:newA});
          if (idx === 1 && w.rsiChart)  w.rsiChart.applyOptions({width:pw, height:newB});
        }
      };
      const onUp = () => {
        div.classList.remove('dragging');
        document.removeEventListener('mousemove', onMove);
        document.removeEventListener('mouseup',   onUp);
      };
      document.addEventListener('mousemove', onMove);
      document.addEventListener('mouseup',   onUp);
    });
  });
}

/* ══════════════════════════════════
   테마 / 색상 / 소리
══════════════════════════════════ */
function cssVar(v) { return getComputedStyle(document.body).getPropertyValue(v).trim(); }

function toggleTheme() {
  isDark = !isDark; localStorage.setItem('isDark', isDark);
  document.body.classList.toggle('dark-theme', isDark);
  document.getElementById('btn-theme').innerText = isDark ? '☀️' : '🌙';
  const bg=cssVar('--cbg'), tc=cssVar('--ctxt'), gl=cssVar('--gcol');
  const themeOpts = {layout:{background:{color:bg},textColor:tc},grid:{vertLines:{color:gl},horzLines:{color:gl}}};
  Object.values(activeWidgets).forEach(w => {
    w.chart.applyOptions(themeOpts);
    if (w.macdChart) w.macdChart.applyOptions(themeOpts);
    if (w.rsiChart)  w.rsiChart.applyOptions(themeOpts);
  });
}
function toggleCandleStyle() {
  isUS = !isUS; localStorage.setItem('isUSStyle', isUS);
  document.getElementById('btn-style').innerText = isUS ? '🇺🇸' : '🇰🇷';
  Object.values(activeWidgets).forEach(w => applyCandleColors(w));
}
function applyCandleColors(w) {
  const uc=isUS?'#22c55e':'#ef4444', dc=isUS?'#ef4444':'#3b82f6';
  const uv=isUS?'rgba(34,197,94,.4)':'rgba(239,68,68,.4)';
  const dv=isUS?'rgba(239,68,68,.4)':'rgba(59,130,246,.4)';
  w.candleSeries.applyOptions({upColor:uc,downColor:dc,wickUpColor:uc,wickDownColor:dc});
  if (w.data) w.volumeSeries.setData(w.data.map(d=>({time:d.time,value:d.volume,color:d.close>=d.open?uv:dv})));
}
function toggleMA(cid, type, vis) { const w=activeWidgets[cid]; if(w&&w[type+'Series']) w[type+'Series'].applyOptions({visible:vis}); }

function toggleSound() {
  soundOn = !soundOn;
  const b = document.getElementById('btn-sound');
  b.innerText = soundOn ? '🔔' : '🔕'; b.className = soundOn ? 'btn-icon active' : 'btn-icon';
  if (soundOn) showToast('소리 알림이 켜졌습니다.');
}
function playSound(type) {
  if (!soundOn) return;
  const ctx=new(window.AudioContext||window.webkitAudioContext)();
  const osc=ctx.createOscillator(), gain=ctx.createGain();
  osc.connect(gain); gain.connect(ctx.destination);
  if (type==='buy') { osc.type='sine'; osc.frequency.setValueAtTime(523.25,ctx.currentTime); osc.frequency.setValueAtTime(659.25,ctx.currentTime+.1); osc.frequency.setValueAtTime(783.99,ctx.currentTime+.2); }
  else { osc.type='triangle'; osc.frequency.setValueAtTime(440,ctx.currentTime); osc.frequency.setValueAtTime(349.23,ctx.currentTime+.2); }
  gain.gain.setValueAtTime(.1,ctx.currentTime); gain.gain.exponentialRampToValueAtTime(.00001,ctx.currentTime+.5);
  osc.start(); osc.stop(ctx.currentTime+.5);
}

/* ══════════════════════════════════
   차트 데이터 & 지표
══════════════════════════════════ */
function changeTF(cid, iv) {
  const w = activeWidgets[cid]; if (!w) return;
  w.interval = iv;
  const s = document.getElementById(`sig-${cid}`); if (s) s.innerText = '분석중...';
  fetchChart(cid, w.symbol, iv, true);
}
function saveMemo(sym, txt) { localStorage.setItem(`memo_${sym}`, txt); }

async function fetchChart(cid, symbol, interval, autoFit=false) {
  let w = activeWidgets[cid]; if (!w) return;
  try {
    const data = await fetch(`${API}/data?symbol=${symbol}&interval=${interval}&market=${encodeURIComponent(w.market||'')}`).then(r=>r.json());
    w = activeWidgets[cid]; if (!w) return;
    if (!Array.isArray(data) || !data.length) return;
    w.data = data;
    w.candleSeries.setData(data); applyCandleColors(w);

    // 🌟 메인 차트 툴팁
    const mainContainer = document.getElementById(cid);
    if (mainContainer) addFloatingTooltip(w.chart, mainContainer, w.candleSeries);

    const bb = calcBB(data, 20);
    w.ma20Series.setData(bb.sma); w.bbUpSeries.setData(bb.upper); w.bbLowSeries.setData(bb.lower);
    w.ma60Series.setData(calcMA(data,60)); w.ma120Series.setData(calcMA(data,120));

    const macd = calcMACD(data);
    w.macdLineSeries.setData(macd.macd); w.signalLineSeries.setData(macd.signal); w.macdHistSeries.setData(macd.hist);
    // RSI 차트 데이터
    const rsiVals = calcRSIFull(data);
    w.rsiSeries.setData(rsiVals);
    w.rsiOb.setData(data.map(d=>({time:d.time,value:70})));
    w.rsiOs.setData(data.map(d=>({time:d.time,value:30})));
    // RSI 배지 업데이트
    const rEl = document.getElementById(`rsi-${cid}`);
    if (rEl) {
      const lastRsi = rsiVals.length ? rsiVals[rsiVals.length-1].value : 50;
      rEl.innerText = `RSI: ${lastRsi.toFixed(1)}`;
      rEl.className = lastRsi>=70?'badge b-sell':lastRsi<=30?'badge b-buy':'badge b-hold';
    }
    analyzeSignal(w, data, bb.sma, macd, cid, interval);
    // 기간별 수익률 계산 (일봉일 때만)
    if (interval === 'day') calcReturns(cid, data);
    else document.getElementById(`ret-${cid}`) && (document.getElementById(`ret-${cid}`).style.opacity = '.4');

    // 🌟 ARIMA 예측선 (일봉 전용)
    ['forecastSeries','validSeries','errorSeries','prophetSeries','ensembleSeries'].forEach(k => {
      if (w[k]) {
        try { w.chart.removeSeries(w[k]); } catch(e) {}
        w[k] = null;
      }
    });
    if (interval === 'day' && typeof drawForecastLine === 'function') {
      drawForecastLine(w);
    }

    if (autoFit) {
      if (interval === 'day') {
        // 일봉: 기본 1주일 뷰 (1w 버튼 active 포함)
        const retBar = document.getElementById(`ret-${cid}`);
        const wBtn   = retBar?.querySelector('[onclick*="\'1w\'"]');
        zoomChart(cid, '1w', wBtn || null);
      } else {
        // 분봉: 오늘 날짜 범위만 표시
        fitToday(w, data);
      }
    }
  } catch(e) { console.error('차트 오류:', e); }
}

function fitToday(w, data) {
  if (!data || !data.length) return;

  // 오늘 날짜 (KST 기준)
  const now    = new Date();
  const todayY = now.getFullYear();
  const todayM = String(now.getMonth() + 1).padStart(2, '0');
  const todayD = String(now.getDate()).padStart(2, '0');
  const todayStr = `${todayY}-${todayM}-${todayD}`;

  // time이 Unix 타임스탬프인지 문자열인지 판별
  const isUnix = typeof data[0].time === 'number';

  // 오늘 시작/끝 Unix 타임스탬프 (KST = UTC+9)
  const todayStartSec = Math.floor(new Date(`${todayStr}T00:00:00+09:00`).getTime() / 1000);
  const todayEndSec   = todayStartSec + 86400;

  function toSec(t) {
    if (typeof t === 'number') return t;
    return Math.floor(new Date(t + 'T00:00:00+09:00').getTime() / 1000);
  }

  // 오늘 데이터의 인덱스 범위 찾기
  let firstIdx = -1, lastIdx = -1;
  for (let i = 0; i < data.length; i++) {
    const sec = toSec(data[i].time);
    if (sec >= todayStartSec && sec < todayEndSec) {
      if (firstIdx === -1) firstIdx = i;
      lastIdx = i;
    }
  }

  if (firstIdx === -1) {
    // 오늘 데이터 없으면 (주말/공휴일) 마지막 거래일 전체 표시
    w.chart.timeScale().fitContent();
    w.macdChart.timeScale().fitContent();
    w.rsiChart.timeScale().fitContent();
    return;
  }

  // 여유 공간 추가 (앞뒤 5봉)
  const from = Math.max(0, firstIdx - 5);
  const to   = Math.min(data.length - 1, lastIdx + 5);

  // Logical range로 설정 (세 차트 동기화)
  const range = { from, to };
  try {
    w.chart.timeScale().setVisibleLogicalRange(range);
    w.macdChart.timeScale().setVisibleLogicalRange(range);
    w.rsiChart.timeScale().setVisibleLogicalRange(range);
  } catch(e) {
    w.chart.timeScale().fitContent();
    w.macdChart.timeScale().fitContent();
    w.rsiChart.timeScale().fitContent();
  }
}

function calcReturns(cid, data) {
  if (!data || data.length < 2) return;
  const retBar = document.getElementById(`ret-${cid}`);
  if (!retBar) return;
  retBar.style.opacity = '1';

  const last    = data[data.length - 1];
  const current = last.close;

  // time이 Unix 타임스탬프(초)인지 문자열인지 판별 후 통일
  function toSec(t) {
    if (typeof t === 'number') return t;
    // 'YYYY-MM-DD' 문자열 → UTC 초
    return Math.floor(new Date(t + 'T00:00:00Z').getTime() / 1000);
  }

  const nowSec = toSec(last.time);

  // n일 이전 종가 찾기 (타임스탬프 기준)
  function getPast(days) {
    const targetSec = nowSec - days * 86400;
    // targetSec 보다 같거나 이전인 가장 최근 데이터
    for (let i = data.length - 2; i >= 0; i--) {
      if (toSec(data[i].time) <= targetSec) return data[i].close;
    }
    return null; // 데이터 부족
  }

  const periods = [
    { id: `ret-d-${cid}`,  days: 1,   label: '오늘' },
    { id: `ret-w-${cid}`,  days: 7,   label: '1주'  },
    { id: `ret-m-${cid}`,  days: 30,  label: '1개월'},
    { id: `ret-3m-${cid}`, days: 90,  label: '3개월'},
    { id: `ret-y-${cid}`,  days: 365, label: '1년'  },
  ];

  periods.forEach(p => {
    const el = document.getElementById(p.id);
    if (!el) return;
    const base = getPast(p.days);
    if (!base) { el.textContent = 'N/A'; el.className = 'ret-val flat'; return; }
    const pct  = ((current - base) / base * 100).toFixed(1);
    const sign = pct > 0 ? '+' : '';
    const cls  = pct > 0 ? 'up' : pct < 0 ? 'down' : 'flat';
    el.textContent = `${sign}${pct}%`;
    el.className   = `ret-val ${cls}`;
  });
}

function analyzeSignal(w, data, sma, macdData, cid, interval) {
  const last=data[data.length-1], lsma=sma[sma.length-1].value;
  const el=document.getElementById(`sig-${cid}`); if(!el) return;
  const sup20=last.low<=lsma&&last.close>=lsma;
  const m=macdData.macd, s=macdData.signal;
  let golden=false, dead=false;
  if (m.length>=2&&s.length>=2) {
    golden=m[m.length-2].value<=s[s.length-2].value&&m[m.length-1].value>s[s.length-1].value;
    dead  =m[m.length-2].value>=s[s.length-2].value&&m[m.length-1].value<s[s.length-1].value;
  }
  const unit=interval==='day'?'일선':'이평';
  let type='hold';
  if      (golden)  { el.innerText=`🔥 강력매수 (골든크로스)`; el.className='badge b-buy';  type='strong_buy'; }
  else if (dead)    { el.innerText=`❄️ 매도경고 (데드크로스)`; el.className='badge b-sell'; type='sell'; }
  else if (sup20)   { el.innerText=`📈 매수타점 (20${unit} 지지)`; el.className='badge b-buy'; type='buy'; }
  else              { el.innerText='➖ 관망'; el.className='badge b-hold'; }
  if (w.lastSignalType!==type&&type!=='hold') {
    if (type.includes('buy')) { playSound('buy');  showToast(`🔔 [매수] ${w.name} 신호 감지`); }
    if (type==='sell')        { playSound('sell'); showToast(`⚠️ [매도] ${w.name} 하락 신호`); }
  }
  w.lastSignalType = type;

  // ── 조합 신호 (win-bar 우측) ──
  const comboEl = document.getElementById(`combo-${cid}`);
  if (!comboEl || data.length < 20) return;

  const bb = calcBB(data, 20);
  const rsi = parseFloat(calcRSI(data));
  const combos = [];

  // 볼린저 밴드 위치
  const upperBB = bb.upper[bb.upper.length-1].value;
  const lowerBB = bb.lower[bb.lower.length-1].value;
  const midBB   = bb.sma[bb.sma.length-1].value;
  const price   = last.close;

  // 밴드 폭 수축 감지 (최근 20일 중 현재 폭이 가장 좁은 10% 이내)
  const bandWidths = bb.upper.map((u,i) => u.value - bb.lower[i].value);
  const recentW = bandWidths.slice(-20);
  const minW = Math.min(...recentW), maxW = Math.max(...recentW);
  const currW = upperBB - lowerBB;
  const isSqueeze = (currW - minW) / (maxW - minW + 0.0001) < 0.15;

  // ① BB 하단 + RSI 과매도 → 강력매수
  if (price <= lowerBB * 1.005 && rsi < 35)
    combos.push({cls:'cb-strong-buy', text:'🔥BB하단+RSI과매도'});
  // ② BB 상단 + RSI 과매수 → 강력매도 경고
  else if (price >= upperBB * 0.995 && rsi > 65)
    combos.push({cls:'cb-strong-sell', text:'⚠️BB상단+RSI과매수'});
  // ③ BB 수축 + MACD 크로스 → 돌파 임박
  if (isSqueeze && (golden || dead))
    combos.push({cls:'cb-breakout', text:'💥BB수축+MACD크로스'});
  // ④ BB 수축만 (방향 미확정)
  else if (isSqueeze)
    combos.push({cls:'cb-squeeze', text:'🔔BB수축(돌파임박)'});
  // ⑤ BB 상단 터치 + RSI 중간 → 주의
  if (price >= upperBB * 0.995 && rsi >= 50 && rsi <= 65)
    combos.push({cls:'cb-caution', text:'BB상단주의'});

  comboEl.innerHTML = combos.map(c =>
    `<span class="combo-badge ${c.cls}">${c.text}</span>`
  ).join('');
}

/* ══════════════════════════════════
   수학 지표
══════════════════════════════════ */
function calcMA(data,p){const r=[];for(let i=p-1;i<data.length;i++){let s=0;for(let j=0;j<p;j++)s+=data[i-j].close;r.push({time:data[i].time,value:s/p});}return r;}
function calcEMA(data,p){if(!data.length)return[];const k=2/(p+1);let e=data[0].close;const r=[{time:data[0].time,value:e}];for(let i=1;i<data.length;i++){e=(data[i].close-e)*k+e;r.push({time:data[i].time,value:e});}return r;}
function calcMACD(data,f=12,sl=26,sg=9){
  if(data.length<sl)return{macd:[],signal:[],hist:[]};
  const fe=calcEMA(data,f),se=calcEMA(data,sl);
  const macd=data.map((d,i)=>({time:d.time,value:fe[i].value-se[i].value}));
  const k=2/(sg+1);let cs=macd[0].value;
  const signal=[{time:macd[0].time,value:cs}];
  for(let i=1;i<macd.length;i++){cs=(macd[i].value-cs)*k+cs;signal.push({time:macd[i].time,value:cs});}
  const hist=macd.map((d,i)=>{const v=d.value-signal[i].value;return{time:d.time,value:v,color:v>=0?'rgba(38,166,154,.6)':'rgba(239,83,80,.6)'};});
  return{macd,signal,hist};
}
function calcBB(data,p){const sma=[],upper=[],lower=[];for(let i=p-1;i<data.length;i++){let s=0;for(let j=0;j<p;j++)s+=data[i-j].close;const avg=s/p;let sq=0;for(let j=0;j<p;j++)sq+=Math.pow(data[i-j].close-avg,2);const std=Math.sqrt(sq/p);sma.push({time:data[i].time,value:avg});upper.push({time:data[i].time,value:avg+std*2});lower.push({time:data[i].time,value:avg-std*2});}return{sma,upper,lower};}
function calcRSIFull(data,p=14){
  const r=[];
  if(data.length<=p) return r;
  let g=0,l=0;
  for(let i=1;i<=p;i++){const c=data[i].close-data[i-1].close;if(c>0)g+=c;else l-=c;}
  let ag=g/p,al=l/p;
  r.push({time:data[p].time,value:parseFloat((100-(100/(1+(ag/(al||.0001))))).toFixed(2))});
  for(let i=p+1;i<data.length;i++){
    const c=data[i].close-data[i-1].close;
    ag=(ag*13+(c>0?c:0))/14; al=(al*13+(c<0?-c:0))/14;
    r.push({time:data[i].time,value:parseFloat((100-(100/(1+(ag/(al||.0001))))).toFixed(2))});
  }
  return r;
}
function calcRSI(data,p=14){if(data.length<p)return 50;let g=0,l=0;for(let i=1;i<=p;i++){const c=data[i].close-data[i-1].close;if(c>0)g+=c;else l-=c;}let ag=g/p,al=l/p;for(let i=p+1;i<data.length;i++){const c=data[i].close-data[i-1].close;ag=(ag*13+(c>0?c:0))/14;al=(al*13+(c<0?-c:0))/14;}return(100-(100/(1+(ag/(al||1))))).toFixed(1);}

/* ── 토스트 ── */
/* ══════════════════════════════════
   AI 채팅 패널
══════════════════════════════════ */
let aiPanelOpen = false;
const aiHistory = []; // {role, content}
let aiPendingImage = null; // { base64, mediaType, previewUrl, name }

// toggleAIPanel 제거됨

// AI 패널 제거됨 — paste 리스너 불필요

function makeDraggableEl(el, handle) {
  // 이미 드래그 핸들이 등록됐으면 스킵
  if (handle._aiDrag) return;
  handle._aiDrag = true;
  let sx, sy, ox, oy, dragging = false;
  handle.addEventListener('mousedown', e => {
    if (e.target.id === 'ai-close') return;
    e.preventDefault();
    dragging = true;
    const r = el.getBoundingClientRect();
    ox = r.left; oy = r.top; sx = e.clientX; sy = e.clientY;
    el.style.right = 'auto';
    el.style.bottom = 'auto';
    el.style.left = ox + 'px';
    el.style.top  = oy + 'px';
  });
  document.addEventListener('mousemove', e => {
    if (!dragging) return;
    el.style.left = (ox + e.clientX - sx) + 'px';
    el.style.top  = (oy + e.clientY - sy) + 'px';
  });
  document.addEventListener('mouseup', () => { dragging = false; });
}

function aiQuick(text) {
  document.getElementById('ai-input').value = text;
  aiSend();
}

function aiInputKey(e) {
  if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); aiSend(); }
}



// 이미지 선택 처리
function aiImageSelected(input) {
  const file = input.files[0];
  if (!file) return;
  if (!file.type.startsWith('image/')) { showToast('이미지 파일만 첨부 가능합니다.'); return; }
  if (file.size > 5 * 1024 * 1024) { showToast('이미지는 5MB 이하만 가능합니다.'); return; }

  const reader = new FileReader();
  reader.onload = e => {
    const dataUrl   = e.target.result;
    const base64    = dataUrl.split(',')[1];
    const mediaType = file.type;
    aiPendingImage  = { base64, mediaType, previewUrl: dataUrl, name: file.name };

    // 미리보기 표시
    document.getElementById('ai-preview-img').src = dataUrl;
    document.getElementById('ai-img-name').textContent  = file.name;
    document.getElementById('ai-img-preview').classList.add('visible');
    document.getElementById('ai-img-btn').classList.add('has-img');
  };
  reader.readAsDataURL(file);
  input.value = ''; // 같은 파일 재선택 가능하게
}

function aiRemoveImg() {
  aiPendingImage = null;
  document.getElementById('ai-preview-img').src = '';
  document.getElementById('ai-img-preview').classList.remove('visible');
  document.getElementById('ai-img-btn').classList.remove('has-img');
}

function aiAddMsg(role, text, isLoading=false, imgUrl=null) {
  const box = document.getElementById('ai-messages');
  const div = document.createElement('div');
  div.className = `ai-msg ${role}${isLoading?' loading':''}`;
  let html = `<div class="ai-bubble">${text.replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\n/g,'<br>')}`;
  if (imgUrl) html += `<img class="chat-img" src="${imgUrl}" alt="첨부 이미지">`;
  html += `</div>`;
  div.innerHTML = html;
  box.appendChild(div);
  box.scrollTop = box.scrollHeight;
  return div;
}

// chart.js 의 aiSend 함수를 이 '무적 버전'으로 덮어씌우세요.
async function aiSend() {
  const inp = document.getElementById('ai-input') || document.getElementById('rp-ai-input');
  const text = inp?.value?.trim() || ""; 
  const img = window.aiPendingImage;

  if (!text && !img) return;

  // 1. UI 준비 (메시지창 가져오기)
  const msgContainer = document.getElementById('ai-messages') || document.getElementById('rp-ai-messages');
  if (!msgContainer) return;

  // 2. 사용자 메시지 표시
  if (text) {
    const userDiv = document.createElement('div');
    userDiv.className = 'ai-msg user';
    userDiv.innerHTML = `<div class="ai-bubble">${text}</div>`;
    msgContainer.appendChild(userDiv);
  }

  // 3. 봇의 답변을 위한 빈 버블 생성
  const botDiv = document.createElement('div');
  botDiv.className = 'ai-msg bot';
  botDiv.innerHTML = `<div class="ai-bubble">⏳ 차트 분석 중...</div>`;
  msgContainer.appendChild(botDiv);
  msgContainer.scrollTop = msgContainer.scrollHeight;

  const bubble = botDiv.querySelector('.ai-bubble');

  // 4. 전송 데이터 구성
  let userContent = img ? [
    { type: 'image', source: { type: 'base64', media_type: img.mediaType, data: img.base64 } },
    { type: 'text', text: text || '이 차트 이미지를 정밀하게 분석해줘.' }
  ] : text;

  aiHistory.push({ role: 'user', content: userContent });
  if (inp) inp.value = '';

  try {
    const res = await fetch(`${API}/ai/chat`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        system: "당신은 전문 주식 분석가입니다. 차트 이미지를 보고 기술적 지표와 추세를 분석하여 매수/매도 전략을 제시하세요.",
        messages: aiHistory.slice(-10)
      })
    });

    if (!res.ok) throw new Error("서버 응답 에러");

    // 🌟 스트리밍 데이터를 한 조각씩 읽기 위한 핵심 로직
    const reader = res.body.getReader();
    const decoder = new TextDecoder();
    let fullReply = "";
    bubble.innerText = ""; // 로딩 문구 제거

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;

      const chunk = decoder.decode(value, { stream: true });
      fullReply += chunk;
      
      // 화면에 실시간으로 타이핑 효과처럼 표시
      bubble.innerText = fullReply;
      msgContainer.scrollTop = msgContainer.scrollHeight;
    }

    aiHistory.push({ role: 'assistant', content: fullReply });
    window.aiPendingImage = null; // 캡처 데이터 초기화

  } catch (e) {
    console.error("AI 에러:", e);
    bubble.innerText = "❌ 분석 중 오류가 발생했습니다: " + e.message;
  }
}
// analyzeChartWithAi → 파일 하단 단일 정의 사용
/* ══════════════════════════════════
   펀더멘털 (PER/PBR/PSR/시총)
══════════════════════════════════ */
async function loadFundamentals(cid, symbol) {
  const panel = document.getElementById(`kis-fund-${cid}`);
  if (!panel) return;
  try {
    const d = await fetch(`${API}/fundamentals?symbol=${symbol}`).then(r => r.json());
    if (d.error) return;

    const set = (id, val, suffix='', isNum=true) => {
      const el = document.getElementById(`${id}-${cid}`);
      if (!el) return;
      if (val === null || val === undefined) { el.textContent = 'N/A'; return; }
      const txt = suffix ? val + suffix : val;
      el.textContent = txt;
      if (isNum && typeof val === 'number') {
        el.className = 'kis-fund-val' + (val > 0 ? ' pos' : val < 0 ? ' neg' : '');
      }
    };

    // 밸류
    set('kf-mcap',  d.mktcap, '', false);
    set('kf-per',   d.per,  '배');
    set('kf-perf',  d.per_fwd, '배');
    set('kf-pbr',   d.pbr,  '배');
    set('kf-psr',   d.psr,  '배');
    set('kf-div',   d.div_yield, '%');
    // 실적
    set('kf-rev',   d.revenue, '', false);
    set('kf-revg',  d.rev_growth, '%');
    set('kf-eps',   d.eps,  '');
    set('kf-epsf',  d.eps_fwd, '');
    set('kf-eg',    d.earn_growth, '%');
    // 수익성
    set('kf-om',    d.oper_margin, '%');
    set('kf-pm',    d.profit_margin, '%');
    set('kf-roe',   d.roe, '%');
    set('kf-roa',   d.roa, '%');
    // 현금흐름
    set('kf-fcf',   d.fcf, '', false);
    set('kf-ocf',   d.ocf, '', false);
    set('kf-52h',   d['52w_high'], '');
    set('kf-52l',   d['52w_low'],  '');

    panel.style.display = 'flex';
  } catch(e) {
    console.warn('펀더멘털 로드 실패:', e);
  }
}

function calcFundamentalScore(cid) {
  const get = id => {
    const el = document.getElementById(`${id}-${cid}`);
    return el ? parseFloat(el.textContent) : null;
  };
  const per  = get('kf-per');
  const pbr  = get('kf-pbr');
  const roe  = parseFloat(document.getElementById(`kf-roe-${cid}`)?.textContent) || null;
  const revg = parseFloat(document.getElementById(`kf-revg-${cid}`)?.textContent) || null;
  const opm  = parseFloat(document.getElementById(`kf-om-${cid}`)?.textContent) || null;

  let score = 0, total = 0, details = [];

  if (per !== null) {
    total += 25;
    if      (per < 10)  { score += 25; details.push(`PER ${per} → 매우 저평가 ✅`); }
    else if (per < 15)  { score += 20; details.push(`PER ${per} → 저평가 ✅`); }
    else if (per < 25)  { score += 12; details.push(`PER ${per} → 적정 수준`); }
    else if (per < 40)  { score += 5;  details.push(`PER ${per} → 고평가 주의 ⚠️`); }
    else                {              details.push(`PER ${per} → 매우 고평가 ❌`); }
  }
  if (pbr !== null) {
    total += 20;
    if      (pbr < 1)   { score += 20; details.push(`PBR ${pbr} → 자산 대비 저평가 ✅`); }
    else if (pbr < 2)   { score += 15; details.push(`PBR ${pbr} → 적정`); }
    else if (pbr < 4)   { score += 8;  details.push(`PBR ${pbr} → 다소 고평가`); }
    else                {              details.push(`PBR ${pbr} → 고평가 ❌`); }
  }
  if (roe !== null) {
    total += 25;
    if      (roe > 20)  { score += 25; details.push(`ROE ${roe}% → 우수 ✅`); }
    else if (roe > 15)  { score += 20; details.push(`ROE ${roe}% → 양호 ✅`); }
    else if (roe > 10)  { score += 12; details.push(`ROE ${roe}% → 보통`); }
    else if (roe > 0)   { score += 5;  details.push(`ROE ${roe}% → 저조 ⚠️`); }
    else                {              details.push(`ROE ${roe}% → 적자 ❌`); }
  }
  if (revg !== null) {
    total += 15;
    if      (revg > 20) { score += 15; details.push(`매출성장 ${revg}% → 고성장 ✅`); }
    else if (revg > 10) { score += 10; details.push(`매출성장 ${revg}% → 성장 ✅`); }
    else if (revg > 0)  { score += 5;  details.push(`매출성장 ${revg}% → 완만`); }
    else                {              details.push(`매출성장 ${revg}% → 역성장 ❌`); }
  }
  if (opm !== null) {
    total += 15;
    if      (opm > 20)  { score += 15; details.push(`영업이익률 ${opm}% → 우수 ✅`); }
    else if (opm > 10)  { score += 10; details.push(`영업이익률 ${opm}% → 양호`); }
    else if (opm > 5)   { score += 5;  details.push(`영업이익률 ${opm}% → 보통`); }
    else                {              details.push(`영업이익률 ${opm}% → 저조 ⚠️`); }
  }

  const pct = total > 0 ? Math.round(score / total * 100) : 0;
  const grade = pct >= 80 ? '🟢 매수 관심' : pct >= 60 ? '🟡 중립' : pct >= 40 ? '🟠 주의' : '🔴 회피';
  return { score: pct, grade, details };
}

async function fundamentalAnalysis(cid) {
  // 1. 로컬 점수 계산
  const result = calcFundamentalScore(cid);

  // 2. 펀더멘털 데이터 텍스트로 수집
  const get = id => document.getElementById(`${id}-${cid}`)?.textContent || 'N/A';
  const dataText = `
[기본 분석 요청]
종목: ${document.getElementById(`wtitle-${cid}`)?.textContent || cid}

■ 밸류에이션
- PER: ${get('kf-per')} / PER(선행): ${get('kf-perf')}
- PBR: ${get('kf-pbr')} / PSR: ${get('kf-psr')}
- 시가총액: ${get('kf-mcap')} / 배당수익률: ${get('kf-div')}

■ 실적
- 매출: ${get('kf-rev')} / 매출성장: ${get('kf-revg')}
- EPS: ${get('kf-eps')} / EPS(선행): ${get('kf-epsf')}
- 이익성장률: ${get('kf-eg')}

■ 수익성
- 영업이익률: ${get('kf-om')} / 순이익률: ${get('kf-pm')}
- ROE: ${get('kf-roe')} / ROA: ${get('kf-roa')}

■ 현금흐름
- FCF: ${get('kf-fcf')} / OCF: ${get('kf-ocf')}

■ 로컬 점수: ${result.score}점 / 종합의견: ${result.grade}
${result.details.join('\n')}

위 데이터를 바탕으로 기본 분석 리포트를 작성해줘.
투자 매력도, 리스크 요인, 매수/매도/관망 의견을 포함해줘.
  `.trim();

  // 3. 우측 패널 AI 입력창에 세팅 후 전송
  const inp = document.getElementById('rp-ai-input');
  if (inp) inp.value = dataText;

  if (typeof rpSendAi === 'function') {
    rpSendAi();
  } else {
    showToast('❌ AI 전송 함수를 찾을 수 없습니다.');
  }
}

function fundTab(cid, idx, btn) {
  document.querySelectorAll(`#kis-fund-${cid} .kis-fund-tab`).forEach((t,i) => {
    t.classList.toggle('active', i === idx);
  });
  document.querySelectorAll(`#kis-fund-${cid} .kis-fund-page`).forEach((p,i) => {
    p.classList.toggle('active', i === idx);
  });
}

/* ══════════════════════════════════
   KIS 실시간 API
══════════════════════════════════ */
const KIS_PRICE_INTERVAL  = 3000;   // 현재가 갱신 3초
const KIS_OB_INTERVAL     = 2000;   // 호가창 갱신 2초

function fmt(n) {
  return Number(n).toLocaleString('ko-KR');
}

async function fetchKisPrice(cid, symbol) {
  const w = activeWidgets[cid]; if (!w) return;
  try {
    const res = await fetch(`${API}/kis/price?symbol=${symbol}`);
    const d   = await res.json();

    // KIS 실패(주말/장마감) → yfinance .KS 폴백
    if (d.error || !d.price) {
      await fetchKisPriceFallback(cid, symbol);
      return;
    }

    const priceEl  = document.getElementById(`kis-price-${cid}`);
    const changeEl = document.getElementById(`kis-change-${cid}`);
    const dotEl    = document.getElementById(`kis-dot-${cid}`);
    if (!priceEl) return;

    const isUp   = d.change > 0;
    const isDown = d.change < 0;
    const sign   = isUp ? '▲' : isDown ? '▼' : '━';
    const cls    = isUp ? 'up' : isDown ? 'down' : '';

    if (w.lastPrice && w.lastPrice !== d.price) {
      priceEl.style.transition = 'none'; priceEl.style.opacity = '0.3';
      setTimeout(() => { priceEl.style.transition = 'opacity .3s'; priceEl.style.opacity = '1'; }, 50);
    }
    w.lastPrice = d.price;
    priceEl.className   = `kis-price-main ${cls}`;
    priceEl.textContent = fmt(d.price) + '원';
    changeEl.className  = `kis-change ${cls}`;
    changeEl.textContent = `${sign} ${fmt(Math.abs(d.change))} (${d.change_pct > 0 ? '+' : ''}${d.change_pct}%)`;

    const infoRow = document.getElementById(`kis-info-${cid}`);
    if (infoRow) {
      infoRow.innerHTML =
        `<span>시<b>${fmt(d.open)}</b></span>` +
        `<span>고<b style="color:#ef4444">${fmt(d.high)}</b></span>` +
        `<span>저<b style="color:#3b82f6">${fmt(d.low)}</b></span>` +
        `<span style="margin-left:4px">거래량 <b>${fmt(d.volume)}</b></span>` +
        `<span style="margin-left:auto">${d.time || ''}</span>`;
    }
    dotEl.classList.add('live');

    const mcapEl = document.getElementById(`kis-mktcap-${cid}`);
    if (mcapEl && d.mktcap) {
      const mcap = parseFloat(d.mktcap);
      mcapEl.textContent = mcap >= 10000 ? `시총 ${(mcap/10000).toFixed(1)}조` :
                           mcap >= 1000  ? `시총 ${(mcap/1000).toFixed(1)}천억` : `시총 ${Math.round(mcap)}억`;
      mcapEl.style.display = '';
    }
	// KIS PER/PBR → 펀더멘털 패널에 직접 주입
	if (d.per) {
	  const perEl = document.getElementById(`kf-per-${cid}`);
	  if (perEl && (perEl.textContent === '-' || perEl.textContent === 'N/A')) {
		perEl.textContent = d.per + '배';
	  }
	}
	if (d.pbr) {
	  const pbrEl = document.getElementById(`kf-pbr-${cid}`);
	  if (pbrEl && (pbrEl.textContent === '-' || pbrEl.textContent === 'N/A')) {
		pbrEl.textContent = d.pbr + '배';
	  }
	}
	if (d.eps) {
	  const epsEl = document.getElementById(`kf-eps-${cid}`);
	  if (epsEl && (epsEl.textContent === '-' || epsEl.textContent === 'N/A')) {
		epsEl.textContent = d.eps;
	  }
	  
	}	
	// KIS 시총 → 펀더멘털 탭에 주입
	if (d.mktcap) {
	  const mcapEl = document.getElementById(`kf-mcap-${cid}`);
	  if (mcapEl && (mcapEl.textContent === '-' || mcapEl.textContent === 'N/A')) {
		const mcap = parseInt(d.mktcap);
		mcapEl.textContent = mcap >= 10000 ? `${(mcap/10000).toFixed(1)}조` :
							 mcap >= 1000  ? `${(mcap/1000).toFixed(1)}천억` : `${mcap}억`;
	  }
	}
  } catch(e) {
    console.warn('KIS 현재가 오류, 폴백 시도:', e);
    await fetchKisPriceFallback(cid, symbol);
  }
}

// KIS 실패 시 yfinance .KS 폴백
async function fetchKisPriceFallback(cid, symbol) {
  const w = activeWidgets[cid]; if (!w) return;
  try {
    const d = await fetch(`${API}/us/quote?symbol=${symbol}.KS`).then(r => r.json());
    if (d.error || !d.price) return;

    const priceEl  = document.getElementById(`kis-price-${cid}`);
    const changeEl = document.getElementById(`kis-change-${cid}`);
    const infoRow  = document.getElementById(`kis-info-${cid}`);
    if (!priceEl) return;

    const isUp   = d.change > 0;
    const isDown = d.change < 0;
    const sign   = isUp ? '▲' : isDown ? '▼' : '━';
    const cls    = isUp ? 'up' : isDown ? 'down' : '';

    priceEl.className   = `kis-price-main ${cls}`;
    priceEl.textContent = fmt(d.price) + '원';
    changeEl.className  = `kis-change ${cls}`;
    changeEl.textContent = `${sign} ${fmt(Math.abs(d.change))} (${d.change_pct > 0 ? '+' : ''}${d.change_pct}%) ※지연`;

    if (infoRow) infoRow.innerHTML =
		`<span>매도 <b>${d.ask ? '$'+d.ask : '-'}</b></span>` +
		`<span>매수 <b>${d.bid ? '$'+d.bid : '-'}</b></span>` +
		`<span style="margin-left:auto;opacity:.7">※15분지연</span>`;
  } catch(e) { console.warn('KIS 폴백 오류:', e); }
}

async function fetchKisOrderbookModal(cid, symbol) {
  try {
    const res = await fetch(`${API}/kis/orderbook?symbol=${symbol}`);
    const d   = await res.json();
    if (d.error) return;
    const w        = activeWidgets[cid]; if (!w) return;
    const body     = document.getElementById('ob-body-modal');
    const askTotal = document.getElementById('ob-ask-total-modal');
    const bidTotal = document.getElementById('ob-bid-total-modal');
    if (!body) return;

    const maxQty = Math.max(...d.asks.map(a=>a.qty), ...d.bids.map(b=>b.qty), 1);
    let html = '';
    const asksRev = [...d.asks].reverse();
    asksRev.forEach(a => {
      const bw = Math.round(a.qty / maxQty * 100);
      html += `<div class="ob-row">
        <span class="ob-ask-qty" style="background:linear-gradient(to left,rgba(59,130,246,.15) ${bw}%,transparent ${bw}%);padding:1px 4px;border-radius:2px">${fmt(a.qty)}</span>
        <span class="ob-price ask">${fmt(a.price)}</span>
        <span></span></div>`;
    });
    html += `<div class="ob-row"><span></span><span class="ob-price current">${fmt(w.lastPrice||'-')}</span><span></span></div>`;
    d.bids.forEach(b => {
      const bw = Math.round(b.qty / maxQty * 100);
      html += `<div class="ob-row">
        <span></span>
        <span class="ob-price bid">${fmt(b.price)}</span>
        <span class="ob-bid-qty" style="background:linear-gradient(to right,rgba(239,68,68,.15) ${bw}%,transparent ${bw}%);padding:1px 4px;border-radius:2px">${fmt(b.qty)}</span></div>`;
    });
    body.innerHTML = html;
    if (askTotal) askTotal.textContent = `매도 ${fmt(d.total_ask)}`;
    if (bidTotal) bidTotal.textContent = `매수 ${fmt(d.total_bid)}`;
  } catch(e) { console.warn('호가 모달 오류:', e); }
}

async function fetchKisOrderbook(cid, symbol) {
  const w = activeWidgets[cid]; if (!w || !w.obVisible) return;
  try {
    const res = await fetch(`${API}/kis/orderbook?symbol=${symbol}`);
    const d   = await res.json();
    if (d.error) return;

    const body     = document.getElementById(`ob-body-${cid}`);
    const askTotal = document.getElementById(`ob-ask-total-${cid}`);
    const bidTotal = document.getElementById(`ob-bid-total-${cid}`);
    if (!body) return;

    // 최대 수량 (막대 너비 계산용)
    const maxQty = Math.max(...d.asks.map(a=>a.qty), ...d.bids.map(b=>b.qty), 1);
    let html = '';

    // 매도호가 10단계 (역순 — 가장 높은 매도가 위에)
    const asksRev = [...d.asks].reverse();
    asksRev.forEach(a => {
      const w = Math.round(a.qty / maxQty * 100);
      html += `<div class="ob-row">
        <span class="ob-ask-qty" style="background:linear-gradient(to left,rgba(59,130,246,.15) ${w}%,transparent ${w}%);padding:1px 4px;border-radius:2px">${fmt(a.qty)}</span>
        <span class="ob-price ask">${fmt(a.price)}</span>
        <span></span>
      </div>`;
    });

    // 현재가 구분선
    html += `<div class="ob-row"><span></span><span class="ob-price current">${fmt(w.lastPrice||'-')}</span><span></span></div>`;

    // 매수호가 10단계
    d.bids.forEach(b => {
      const bw = Math.round(b.qty / maxQty * 100);
      html += `<div class="ob-row">
        <span></span>
        <span class="ob-price bid">${fmt(b.price)}</span>
        <span class="ob-bid-qty" style="background:linear-gradient(to right,rgba(239,68,68,.15) ${bw}%,transparent ${bw}%);padding:1px 4px;border-radius:2px">${fmt(b.qty)}</span>
      </div>`;
    });

    body.innerHTML = html;
    if (askTotal) askTotal.textContent = `매도 ${fmt(d.total_ask)}`;
    if (bidTotal) bidTotal.textContent = `매수 ${fmt(d.total_bid)}`;
  } catch(e) { console.warn('호가 오류:', e); }
}

/* ══════════════════════════════════
   미국 주식 실시간 (yfinance, 15분 지연)
══════════════════════════════════ */
function startUsRealtime(cid, symbol) {
  const w = activeWidgets[cid]; if (!w) return;
  fetchUsPrice(cid, symbol);
  w.kisTimer = setInterval(() => fetchUsPrice(cid, symbol), 15000);
}

async function fetchUsPrice(cid, symbol) {
  const w = activeWidgets[cid]; if (!w) return;
  try {
    const res = await fetch(`${API}/us/quote?symbol=${symbol}`);
    const d   = await res.json();
    if (d.error) return;

    const priceEl  = document.getElementById(`kis-price-${cid}`);
    const changeEl = document.getElementById(`kis-change-${cid}`);
    if (!priceEl) return;

    const isUp  = d.change > 0;
    const isDown= d.change < 0;
    const sign  = isUp ? '▲' : isDown ? '▼' : '━';
    const cls   = isUp ? 'up' : isDown ? 'down' : '';

    priceEl.className   = `kis-price-main ${cls}`;
    priceEl.textContent = `$${d.price}`;
    changeEl.className  = `kis-change ${cls}`;
    changeEl.textContent= `${sign} ${Math.abs(d.change).toFixed(2)} (${d.change_pct > 0 ? '+' : ''}${d.change_pct}%)`;
    const infoRow = document.getElementById(`kis-info-${cid}`);
    if (infoRow) {
      infoRow.innerHTML =
        `<span>매도 <b>$${d.ask}</b>(${fmt(d.ask_size)})</span>` +
        `<span>매수 <b>$${d.bid}</b>(${fmt(d.bid_size)})</span>` +
        `<span style="margin-left:auto">${d.time} <span style="opacity:.6">※15분지연</span></span>`;
    }
    w.lastPrice = d.price;
    w.usQuote = d;
    // 펀더멘털 — 최초 1회
    if (!w.fundLoaded) {
      w.fundLoaded = true;
      loadFundamentals(cid, symbol);
    }
  } catch(e) { console.warn('US 현재가 오류:', e); }
}

async function fetchUsOrderbookModal(cid, symbol) {
  const w = activeWidgets[cid]; if (!w) return;
  try {
    const res = await fetch(`${API}/us/quote?symbol=${symbol}`);
    const d   = await res.json();
    if (d.error) return;
    w.usQuote = d;

    const body     = document.getElementById('ob-body-modal');
    const askTotal = document.getElementById('ob-ask-total-modal');
    const bidTotal = document.getElementById('ob-bid-total-modal');
    if (!body) return;

    const maxQty = Math.max(...d.asks.map(a=>a.qty), ...d.bids.map(b=>b.qty), 1);
    let html = '';
    // 매도호가 역순
    [...d.asks].reverse().forEach(a => {
      const bw = Math.round(a.qty / maxQty * 100);
      html += `<div class="ob-row">
        <span class="ob-ask-qty" style="background:linear-gradient(to left,rgba(59,130,246,.15) ${bw}%,transparent ${bw}%);padding:1px 4px;border-radius:2px">${fmt(a.qty)}</span>
        <span class="ob-price ask">$${a.price}</span>
        <span></span></div>`;
    });
    html += `<div class="ob-row"><span></span><span class="ob-price current">$${d.price}</span><span></span></div>`;
    d.bids.forEach(b => {
      const bw = Math.round(b.qty / maxQty * 100);
      html += `<div class="ob-row">
        <span></span>
        <span class="ob-price bid">$${b.price}</span>
        <span class="ob-bid-qty" style="background:linear-gradient(to right,rgba(239,68,68,.15) ${bw}%,transparent ${bw}%);padding:1px 4px;border-radius:2px">${fmt(b.qty)}</span></div>`;
    });
    body.innerHTML = html;
    if (askTotal) askTotal.textContent = `매도 ${fmt(d.ask_size)}`;
    if (bidTotal) bidTotal.textContent = `매수 ${fmt(d.bid_size)}`;
  } catch(e) { console.warn('US 호가 오류:', e); }
}

async function fetchUsTradesModal(cid, symbol) {
  const tickBody = document.getElementById('tick-body-modal');
  if (!tickBody) return;
  tickBody.innerHTML = '<div style="padding:8px;color:var(--muted);font-size:11px;text-align:center">불러오는 중...</div>';
  try {
    const res = await fetch(`${API}/us/trades?symbol=${symbol}`);
    const d   = await res.json();
    if (d.error || !d.trades) return;
    tickBody.innerHTML = '';
    d.trades.forEach(t => {
      const row = document.createElement('div');
      row.className = `tick-row ${t.dir}`;
      row.innerHTML = `
        <span class="tick-time">${t.time}</span>
        <span class="tick-price ${t.dir}">$${t.price}</span>
        <span class="tick-qty">${fmt(t.qty)}</span>`;
      tickBody.appendChild(row);
    });
  } catch(e) { console.warn('US 체결 오류:', e); }
}

function startKisRealtime(cid, symbol) {
  const w = activeWidgets[cid]; if (!w) return;

  // 펀더멘털은 KIS 성공 여부와 무관하게 즉시 로드 (주말/장마감도 동작)
  if (!w.fundLoaded) {
    w.fundLoaded = true;
    loadFundamentals(cid, symbol);
  }

  // 즉시 1회 조회
  fetchKisPrice(cid, symbol);

  // 현재가 타이머
  w.kisTimer = setInterval(() => fetchKisPrice(cid, symbol), KIS_PRICE_INTERVAL);

  // 호가 타이머
  w.kisObTimer = setInterval(() => fetchKisOrderbook(cid, symbol), KIS_OB_INTERVAL);

  // WebSocket 실시간 체결가
  startKisWebSocket(cid, symbol);
}

async function startKisWebSocket(cid, symbol) {
  const w = activeWidgets[cid]; if (!w) return;
  try {
    const res  = await fetch(`${API}/kis/ws-key`);
    const data = await res.json();
    if (data.error) { console.warn('WS 키 오류:', data.error); return; }

    const ws = new WebSocket('ws://ops.koreainvestment.com:21000');
    w.kisWs  = ws;

    ws.onopen = () => {
      // 실시간 체결가 구독
      ws.send(JSON.stringify({
        header: {
          approval_key: data.approval_key,
          custtype: 'P',
          tr_type: '1',
          content_type: 'utf-8'
        },
        body: {
          input: { tr_id: 'H0STCNT0', tr_key: symbol }
        }
      }));
      console.log(`✅ KIS WebSocket 연결: ${symbol}`);
    };

    ws.onmessage = e => {
      if (!e.data || e.data.startsWith('{')) return; // 컨트롤 메시지 무시
      const parts = e.data.split('^');
      if (parts.length < 13) return;
      const tickBody = document.getElementById(`tick-body-${cid}`);
      if (!tickBody) return;

      const price = parseInt(parts[2]);
      const qty   = parseInt(parts[12]);
      const time  = parts[1] ? parts[1].replace(/(\d{2})(\d{2})(\d{2})/, '$1:$2:$3') : '';
      const w2    = activeWidgets[cid];
      const dir   = w2 && price >= w2.lastPrice ? 'up' : 'down';

      const makeRow = () => {
        const row = document.createElement('div');
        row.className = `tick-row ${dir}`;
        row.innerHTML = `
          <span class="tick-time">${time}</span>
          <span class="tick-price ${dir}">${fmt(price)}</span>
          <span class="tick-qty">${fmt(qty)}</span>`;
        return row;
      };
      tickBody.insertBefore(makeRow(), tickBody.firstChild);
      while (tickBody.children.length > 100) tickBody.removeChild(tickBody.lastChild);

      // 모달 체결창 미러링
      if (obModalCid === cid) {
        const modalTickBody = document.getElementById('tick-body-modal');
        if (modalTickBody) {
          modalTickBody.insertBefore(makeRow(), modalTickBody.firstChild);
          while (modalTickBody.children.length > 100) modalTickBody.removeChild(modalTickBody.lastChild);
        }
      }
    };

    ws.onerror = e => console.warn('KIS WS 오류:', e);
    ws.onclose = () => console.log('KIS WS 연결 종료');
  } catch(e) { console.warn('KIS WebSocket 시작 오류:', e); }
}

// ── 글로벌 모달 상태 ──
let obModalCid    = null;   // 현재 모달에 연결된 팝업 cid
let obModalTab    = 'ob';   // 'ob' | 'tick' | 'both'
let obModalTimer  = null;

function openObModal(cid, tab) {
  const w = activeWidgets[cid]; if (!w) return;
  obModalCid = cid;
  obModalTab = tab;

  const modal   = document.getElementById('ob-global-modal');
  const title   = document.getElementById('ob-modal-title');
  const obCol   = document.getElementById('ob-modal-ob-col');
  const tickCol = document.getElementById('ob-modal-tick-col');

  // 탭에 따라 컬럼 표시
  const showOb   = (tab === 'ob'   || tab === 'both');
  const showTick = (tab === 'tick' || tab === 'both');
  obCol.style.display   = showOb   ? 'flex' : 'none';
  tickCol.style.display = showTick ? 'flex' : 'none';

  const labels = { ob:'📊 호가창', tick:'⚡ 실시간 체결', both:'📊 호가창 + 체결' };
  title.textContent = `${w.name}  ${labels[tab] || ''}`;

  // 팝업 근처에 위치
  const popup = w.popupEl;
  const pr    = popup.getBoundingClientRect();
  const mw    = tab === 'both' ? 410 : tab === 'ob' ? 210 : 200;
  let left    = pr.right + 8;
  let top     = pr.top + 80;
  if (left + mw > window.innerWidth - 10) left = pr.left - mw - 8;
  if (left < 4) left = 4;
  modal.style.left    = left + 'px';
  modal.style.top     = top  + 'px';
  // 높이: 헤더(36) + 컬럼헤더(28) + 행(11행 × 24) + 합계(28) = 약 356px
  modal.style.height  = 'auto';
  modal.style.maxHeight = Math.min(700, window.innerHeight * 0.88) + 'px';
  modal.classList.add('visible');

  // 드래그 등록
  makeObModalDraggable();

  // 타이머 시작
  clearInterval(obModalTimer);
  const isKrxModal = /^[0-9]{6}$/.test(w.symbol);
  if (isKrxModal) {
    if (showOb) fetchKisOrderbookModal(cid, w.symbol);
    if (showTick && w.kisWs && w.kisWs.readyState === 1) w.mirrorTick = true;
    if (showOb) obModalTimer = setInterval(() => {
      if (obModalCid && activeWidgets[obModalCid]) {
        fetchKisOrderbookModal(obModalCid, activeWidgets[obModalCid].symbol);
      }
    }, 2000);
  } else {
    // 미국 주식
    if (showOb)   fetchUsOrderbookModal(cid, w.symbol);
    if (showTick) fetchUsTradesModal(cid, w.symbol);
    if (showOb) obModalTimer = setInterval(() => {
      if (obModalCid && activeWidgets[obModalCid]) {
        fetchUsOrderbookModal(obModalCid, activeWidgets[obModalCid].symbol);
      }
    }, 15000);
    if (showTick) setInterval(() => {
      if (obModalCid && activeWidgets[obModalCid]) {
        fetchUsTradesModal(obModalCid, activeWidgets[obModalCid].symbol);
      }
    }, 60000);
  }
}

function closeObModal() {
  clearInterval(obModalTimer);
  obModalTimer = null;
  const modal = document.getElementById('ob-global-modal');
  modal.classList.remove('visible');
  // 버튼 active 해제
  if (obModalCid) {
    const w = activeWidgets[obModalCid];
    if (w) {
      w.obVisible   = false;
      w.tickVisible = false;
      w.mirrorTick  = false;
    }
    document.getElementById(`kis-ob-btn-${obModalCid}`)?.classList.remove('active');
    document.getElementById(`kis-tick-btn-${obModalCid}`)?.classList.remove('active');
  }
  obModalCid = null;
}

function makeObModalDraggable() {
  const modal = document.getElementById('ob-global-modal');
  const bar   = document.getElementById('ob-modal-bar');
  let dragging = false, sx, sy, ox, oy;
  bar.onmousedown = e => {
    dragging = true;
    const r = modal.getBoundingClientRect();
    ox = r.left; oy = r.top; sx = e.clientX; sy = e.clientY;
    e.preventDefault();
  };
  document.addEventListener('mousemove', e => {
    if (!dragging) return;
    modal.style.left = (ox + e.clientX - sx) + 'px';
    modal.style.top  = (oy + e.clientY - sy) + 'px';
  });
  document.addEventListener('mouseup', () => { dragging = false; });
}

function toggleOrderbook(cid) {
  const w = activeWidgets[cid]; if (!w) return;
  // 이미 이 cid의 ob 모달이 열려있으면 닫기
  if (obModalCid === cid && obModalTab === 'ob' &&
      document.getElementById('ob-global-modal').classList.contains('visible')) {
    closeObModal(); return;
  }
  w.obVisible = true;
  document.getElementById(`kis-ob-btn-${cid}`)?.classList.add('active');
  document.getElementById(`kis-tick-btn-${cid}`)?.classList.remove('active');
  openObModal(cid, 'ob');
}

function toggleTickStream(cid) {
  const w = activeWidgets[cid]; if (!w) return;
  if (obModalCid === cid && obModalTab === 'tick' &&
      document.getElementById('ob-global-modal').classList.contains('visible')) {
    closeObModal(); return;
  }
  w.tickVisible = true;
  document.getElementById(`kis-tick-btn-${cid}`)?.classList.add('active');
  document.getElementById(`kis-ob-btn-${cid}`)?.classList.remove('active');
  openObModal(cid, 'tick');
}

/* ══════════════════════════════════
   투자자별 매매동향
══════════════════════════════════ */
function toggleInvestor(cid) {
  const w = activeWidgets[cid]; if (!w) return;
  w.invVisible = !w.invVisible;
  const panel = document.getElementById(`inv-panel-${cid}`);
  const btn   = document.getElementById(`kis-inv-btn-${cid}`);
  if (panel) panel.classList.toggle('visible', w.invVisible);
  if (btn)   btn.classList.toggle('active',    w.invVisible);
  if (w.invVisible) {
    w.invPeriod = w.invPeriod || 'day';
    fetchInvestor(cid, w.symbol, w.invPeriod);
  }
}

function switchInvTab(cid, period, tabEl) {
  const w = activeWidgets[cid]; if (!w) return;
  w.invPeriod = period;
  // 탭 active 전환
  tabEl.closest('.inv-tabs').querySelectorAll('.inv-tab').forEach(t => t.classList.remove('active'));
  tabEl.classList.add('active');
  // 로딩 표시 후 데이터 조회
  const body = document.getElementById(`inv-body-${cid}`);
  if (body) body.innerHTML = '<div class="inv-loading">불러오는 중...</div>';
  fetchInvestor(cid, w.symbol, period);
}

async function fetchInvestor(cid, symbol, period) {
  const body = document.getElementById(`inv-body-${cid}`);
  if (!body) return;
  try {
    const res  = await fetch(`${API}/kis/investor?symbol=${symbol}&period=${period}`);
    const d    = await res.json();
    if (d.error) { body.innerHTML = `<div class="inv-loading">⚠ ${d.error}</div>`; return; }

    const rows = d.data || [];
    if (!rows.length) { body.innerHTML = '<div class="inv-loading">데이터 없음</div>'; return; }

    // 최대 순매수 절댓값 (바 차트용)
    const maxNet = Math.max(...rows.map(r => Math.abs(r.net || 0)), 1);
    const maxBuy = Math.max(...rows.map(r => Math.abs(r.buy || 0)), 1);

    let html = `<table class="inv-table">
      <tr>
        <th>구분</th>
        <th>매수</th>
        <th>매도</th>
        <th>순매수</th>
      </tr>`;

    rows.forEach(r => {
      const netCls = r.net > 0 ? 'inv-net-pos' : r.net < 0 ? 'inv-net-neg' : '';
      const netSign = r.net > 0 ? '+' : '';
      const buyStr  = r.buy  != null ? fmt(r.buy)  : '-';
      const sellStr = r.sell != null ? fmt(r.sell) : '-';
      html += `<tr>
        <td>${r.name}</td>
        <td class="inv-buy">${buyStr}</td>
        <td class="inv-sell">${sellStr}</td>
        <td class="${netCls}">${netSign}${fmt(r.net)}</td>
      </tr>`;
    });
    html += '</table>';

    // 순매수 비율 바 차트
    html += `<div class="inv-bar-wrap">
      <div class="inv-bar-label">순매수 비율</div>`;
    rows.forEach(r => {
      const pct = Math.round(Math.abs(r.net) / maxNet * 100);
      const fillCls = r.net >= 0 ? 'inv-bar-fill-buy' : 'inv-bar-fill-sell';
      html += `<div class="inv-bar-row">
        <span class="inv-bar-name">${r.name}</span>
        <div class="inv-bar-track">
          <div class="${fillCls}" style="width:${pct}%"></div>
        </div>
        <span style="font-size:11px;width:36px;text-align:right;color:${r.net>=0?'#ef4444':'#3b82f6'};font-weight:700">
          ${pct}%
        </span>
      </div>`;
    });
    html += '</div>';

    body.innerHTML = html;
  } catch(e) {
    if (body) body.innerHTML = `<div class="inv-loading">⚠ 오류: ${e.message}</div>`;
  }
}

/* ══════════════════════════════════
   삭제 확인 모달
══════════════════════════════════ */
let _confirmCb = null;

function showConfirm(msg, icon, okLabel, cb) {
  _confirmCb = cb;
  document.getElementById('confirm-icon').textContent  = icon  || '🗑';
  document.getElementById('confirm-msg').innerHTML     = msg;
  document.getElementById('confirm-ok-btn').textContent = okLabel || '삭제';
  document.getElementById('confirm-modal').classList.add('open');
}
function confirmOk() {
  document.getElementById('confirm-modal').classList.remove('open');
  if (_confirmCb) { _confirmCb(); _confirmCb = null; }
}
function confirmCancel() {
  document.getElementById('confirm-modal').classList.remove('open');
  _confirmCb = null;
}

/* ══════════════════════════════════
   deleteCard (카드 ✕ 버튼용)
══════════════════════════════════ */
function deleteCard(symbol, isMy, name) {
  const label = name || symbol;
  const listName = String(isMy) === 'true' ? '내종목' : '관심종목';
  showConfirm(`<b>${label}</b>을(를)<br>${listName}에서 삭제할까요?`, '🗑', '삭제', () => {
    if (String(isMy) === 'true') {
      saveMyList(getMyList().filter(x => x.symbol !== symbol));
    } else {
      saveList(getList().filter(x => x.symbol !== symbol));
    }
    loadWatch();
    showToast(`🗑 ${label} 삭제됨`);
  });
}

/* ══════════════════════════════════
   buildCard — 카드 DOM 생성 (공통)
══════════════════════════════════ */
function buildCard(item) {
  const mc      = mktCls(item.market);
  const isKr    = /^[0-9]{6}$/.test(item.symbol);
  const ctyCls  = isKr ? ' kr-card' : ' us-card';
  const card    = document.createElement('div');
  card.className = 'wc-card' + (item._my ? ' my-stock' : '') + ctyCls;
  card.onclick = () => openPopup(item.name, item.symbol, item.market||'기타');
  card.dataset.symbol = item.symbol;
  card.dataset.market = item.market || '';
  card.dataset.my     = item._my ? '1' : '0';
  card.dataset.name   = item.name;
  const sector = getSectorMap()[item.symbol];
  const sectorBadge = sector
    ? ` &nbsp;<span class="sector-badge" style="--sdot:${SECTOR_COLORS[sector]||'#94a3b8'}">${sector}</span>`
    : '';
  card.innerHTML = `
    <div class="wc-name">${item._my ? '★ ' : ''}${item.name}</div>
    <div class="wc-meta"><span class="market-badge ${mc}">${item.market||'기타'}</span> &nbsp;${item.symbol}${sectorBadge}</div>
    <div class="wc-body">
      <div class="wc-left">
        <div class="wc-price-row">
          <span class="wc-price" id="wcp-${item.symbol}">···</span>
          <span class="wc-chg"  id="wcg-${item.symbol}">-</span>
        </div>
        <div class="cross-badges" id="cb-${item.symbol}"></div>
      </div>
      <div class="wc-right" id="wcr-${item.symbol}">
        <div class="wc-holding" id="wch-${item.symbol}"></div>
      </div>
    </div>
    ${item._my ? `<button class="wc-edit-btn" onclick="event.stopPropagation();openHoldingModal('${item.symbol}','${item.name}')">✏️</button>` : ''}
    <button class="wc-del-btn" onclick="event.stopPropagation();deleteCard('${item.symbol}','${item._my}','${item.name.replace(/'/g,"\\'")}')">✕</button>
    <button class="wc-btn">📈 차트 열기</button>`;
  bindCardDrag(card, item.symbol);
  return card;
}

/* ══════════════════════════════════
   섹터 그룹화
══════════════════════════════════ */
const SECTOR_PRESETS = ['반도체','IT/기술','자동차','금융','바이오/헬스','에너지','소비재','미디어/엔터','부동산','기타'];
const SECTOR_COLORS  = {
  '반도체':'#3b82f6', 'IT/기술':'#8b5cf6', '자동차':'#f59e0b',
  '금융':'#10b981',   '바이오/헬스':'#ef4444', '에너지':'#f97316',
  '소비재':'#ec4899', '미디어/엔터':'#06b6d4', '부동산':'#84cc16',
  '기타':'#94a3b8',   '미분류':'#64748b',
};

function getSectorMap()         { return JSON.parse(localStorage.getItem('sectorMap')       || '{}'); }
function saveSectorMap(m)       { localStorage.setItem('sectorMap', JSON.stringify(m)); }
function getCustomSectors()     { return JSON.parse(localStorage.getItem('customSectors')   || '[]'); }
function saveCustomSectors(l)   { localStorage.setItem('customSectors', JSON.stringify(l)); }
function getSectorCollapsed()   { return JSON.parse(localStorage.getItem('sectorCollapsed') || '{}'); }
function saveSectorCollapsed(s) { localStorage.setItem('sectorCollapsed', JSON.stringify(s)); }

let _sectorViewOn = localStorage.getItem('sectorView') === 'on';

function toggleSectorView() {
  _sectorViewOn = !_sectorViewOn;
  localStorage.setItem('sectorView', _sectorViewOn ? 'on' : 'off');
  loadWatch();
}

function renderSectorView(wc, allItems) {
  // ← _sectorData에서 sectorMap 형식으로 변환 (symbol → 섹터이름)
  const sectorMap = {};
  Object.values(_sectorData).forEach(info => {
    (info.stocks || []).forEach(sym => {
      sectorMap[sym] = info.name;
    });
  });
  const collapsed = getSectorCollapsed();

  // 종목 → 섹터 그룹화
  const groups = {};
  allItems.forEach(item => {
    const sec = sectorMap[item.symbol] || '미분류';
    if (!groups[sec]) groups[sec] = [];
    groups[sec].push(item);
  });

  // 섹터 출력 순서: 프리셋 → 커스텀 → 미분류
  const customs = getCustomSectors().filter(s => groups[s] && !SECTOR_PRESETS.includes(s));
  const ordered = [
    ...SECTOR_PRESETS.filter(s => groups[s]),
    ...customs,
    ...Object.keys(groups).filter(s => !SECTOR_PRESETS.includes(s) && !customs.includes(s) && s !== '미분류'),
  ];
  if (groups['미분류']) ordered.push('미분류');

  ordered.forEach(sector => {
    const items = groups[sector];
    if (!items?.length) return;
    const isCollapsed = !!collapsed[sector];
    const dot  = SECTOR_COLORS[sector] || '#94a3b8';
    const safe = sector.replace(/[^a-z가-힣0-9]/gi, '_');
    const esc  = sector.replace(/'/g, "\\'");

    const groupEl = document.createElement('div');
    groupEl.className = 'sector-group' + (isCollapsed ? ' collapsed' : '');
    groupEl.dataset.sector = sector;
    groupEl.innerHTML = `
      <div class="sector-group-header" onclick="toggleSectorCollapse('${esc}')">
        <span class="sector-dot" style="background:${dot}"></span>
        <span class="sector-group-title">${sector}</span>
        <span class="sector-group-count">${items.length}</span>
        <span class="sector-group-arrow">${isCollapsed ? '▶' : '▼'}</span>
      </div>
      <div class="sector-group-cards" id="sgc-${safe}"></div>`;
    wc.appendChild(groupEl);

    const cardsEl = groupEl.querySelector('.sector-group-cards');
    items.forEach(item => cardsEl.appendChild(buildCard(item)));
  });
}

function toggleSectorCollapse(sector) {
  const collapsed = getSectorCollapsed();
  collapsed[sector] = !collapsed[sector];
  saveSectorCollapsed(collapsed);
  const groupEl = document.querySelector(`.sector-group[data-sector="${sector}"]`);
  if (!groupEl) return;
  groupEl.classList.toggle('collapsed', collapsed[sector]);
  const arrow = groupEl.querySelector('.sector-group-arrow');
  if (arrow) arrow.textContent = collapsed[sector] ? '▶' : '▼';
}

/* ── 섹터 지정 모달 ── */
let _smSym = '', _smName = '';

function openSectorModal(symbol, name) {
  _smSym  = symbol;
  _smName = name;
  const current = getSectorMap()[symbol] || '';
  document.getElementById('sector-modal-title').innerHTML =
    `📁 ${name} — 섹터 지정<button class="sm-close" onclick="closeSectorModal()">✕</button>`;
  document.getElementById('sector-modal-input').value = current;
  _renderSectorBtns(current);
  document.getElementById('sector-modal').classList.add('open');
  setTimeout(() => document.getElementById('sector-modal-input').focus(), 80);
}

function closeSectorModal() {
  document.getElementById('sector-modal').classList.remove('open');
  _smSym = _smName = '';
}

function _renderSectorBtns(current) {
  const wrap = document.getElementById('sector-modal-btns');
  if (!wrap) return;
  const all = [...SECTOR_PRESETS, ...getCustomSectors().filter(s => !SECTOR_PRESETS.includes(s))];
  wrap.innerHTML = all.map(s => {
    const active = s === current ? ' active' : '';
    const dot = SECTOR_COLORS[s] || '#94a3b8';
    return `<button class="sector-preset-btn${active}" onclick="selectSectorPreset('${s.replace(/'/g,"\\'")}')"><span class="sdot-sm" style="background:${dot}"></span>${s}</button>`;
  }).join('');
}

function selectSectorPreset(s) {
  document.getElementById('sector-modal-input').value = s;
  _renderSectorBtns(s);
}

function saveSectorModal() {
  const val = document.getElementById('sector-modal-input').value.trim();
  if (!val) { clearSectorAssign(); return; }
  const map = getSectorMap();
  map[_smSym] = val;
  saveSectorMap(map);
  if (!SECTOR_PRESETS.includes(val)) {
    const c = getCustomSectors();
    if (!c.includes(val)) { c.push(val); saveCustomSectors(c); }
  }
  showToast(`✅ ${_smName} → ${val}`);
  closeSectorModal();
  loadWatch();
}

function clearSectorAssign() {
  const map = getSectorMap();
  delete map[_smSym];
  saveSectorMap(map);
  showToast(`🗑 ${_smName} 섹터 해제`);
  closeSectorModal();
  loadWatch();
}

// 우클릭 메뉴 → 내종목/관심종목 이동
function ctxMoveCard(to) {
  closeCtx();
  if (!_ctxSym) return;
  const sym  = _ctxSym;
  const name = _ctxName;
  const item = (_ctxIsMy ? getMyList() : getList()).find(x => x.symbol === sym);
  if (!to || !item) return;

  if (to === 'my') {
    // 관심종목 → 내종목
    saveList(getList().filter(x => x.symbol !== sym));
    const myL = getMyList();
    if (!myL.find(x => x.symbol === sym)) { myL.push(item); saveMyList(myL); }
    showToast(`★ ${name} → 내종목 이동`);
  } else {
    // 내종목 → 관심종목
    saveMyList(getMyList().filter(x => x.symbol !== sym));
    const wL = getList();
    if (!wL.find(x => x.symbol === sym)) { wL.push(item); saveList(wL); }
    showToast(`⭐ ${name} → 관심종목 이동`);
  }
  loadWatch();
}

// 우클릭 메뉴 → 섹터 지정
function ctxSetSector() {
  closeCtx();
  if (_ctxSym) openSectorModal(_ctxSym, _ctxName);
}

// 섹터 모달 — Enter 저장 / 바깥 클릭 닫기
window.addEventListener('DOMContentLoaded', () => {
  document.getElementById('sector-modal')?.addEventListener('click', function(e) {
    if (e.target === this) closeSectorModal();
  });
  document.getElementById('sector-modal-input')?.addEventListener('keydown', e => {
    if (e.key === 'Enter') saveSectorModal();
  });
  document.getElementById('confirm-modal')?.addEventListener('click', function(e) {
    if (e.target === this) confirmCancel();
  });
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape' && document.getElementById('confirm-modal')?.classList.contains('open')) {
      confirmCancel();
    }
  });
});

// ── 🌟 섹터 관리 전역 변수 ──
let _activeSector = null; // 현재 선택된 섹터 ID
let _sectorData = {};

// ── 🌟 섹터 클릭 이벤트 초기화 함수 ──
function initSectorMenu() {
  const cards = document.querySelectorAll('.sector-card');

  cards.forEach(card => {
	let clickTimer = null; // 단일/더블 클릭을 구분하기 위한 타이머
    card.addEventListener('click', (e) => {
      const sectorId = card.dataset.sector || `custom-${card.dataset.index}`;
      
      // 더블 클릭 감지 로직
      if (clickTimer) {
        clearTimeout(clickTimer);
        clickTimer = null;
        // 🎯 더블 클릭 실행 (Phase D에서 모달창 연결할 곳)
		openSectorEditModal(sectorId, card);
      } else {
        // 단일 클릭 감지 로직 (250ms 대기 후 실행)
        clickTimer = setTimeout(() => {
          clickTimer = null;
          // 🎯 단일 클릭 실행 (필터링)
          toggleSectorFilter(card, sectorId);
        }, 250); 
      }
    });
  });
}

// ── 🌟 섹터 필터링 토글 함수 ──
function toggleSectorFilter(cardElement, sectorId) {
  const cards = document.querySelectorAll('.sector-card');
  
  if (_activeSector === sectorId) {
    // 1. 이미 선택된 섹터를 다시 누름 -> 필터 해제 (전체 보기)
    _activeSector = null;
    cardElement.classList.remove('active');
    filterDashboardBySector(null);
  } else {
    // ← 빈 섹터 체크 추가
    const stocks = _sectorData[sectorId]?.stocks || [];
    if (stocks.length === 0) {
      showToast('⚠ 섹터에 등록된 종목이 없습니다. 더블클릭으로 종목을 추가하세요.');
      return;
    }
    _activeSector = sectorId;
    cards.forEach(c => c.classList.remove('active'));
    cardElement.classList.add('active');
    filterDashboardBySector(sectorId);
  }
}

// ── 🌟 대시보드 화면 필터링 로직 ──
function filterDashboardBySector(sectorId) {
  // 대시보드의 모든 종목 카드를 가져옵니다. 
  // (클래스명이 다를 경우 실제 사용 중인 종목 카드 클래스명으로 변경해주세요)
  const stockCards = document.querySelectorAll('.wc-card'); 
  
  if (!sectorId) {
    // 필터가 해제되면 모든 종목 카드를 다시 보여줍니다.
    stockCards.forEach(card => card.style.display = '');
    return;
  }

  // 현재 선택된 섹터에 등록된 종목 코드 배열 가져오기
  const symbolsInSector = _sectorData[sectorId] || [];

  stockCards.forEach(card => {
    // 카드 안에서 종목 코드를 추출합니다. (wcp-005930 같은 ID에서 추출)
    // HTML 구조에 맞춰 종목 코드를 가져오도록 안전하게 정규식 사용
    const match = card.innerHTML.match(/wcp-([A-Za-z0-9]+)/);
    const symbol = match ? match[1] : null;

    if (symbol && symbolsInSector.includes(symbol)) {
      card.style.display = ''; // 섹터에 포함된 종목은 보여주기
    } else {
      card.style.display = 'none'; // 포함되지 않은 종목은 숨기기
    }
  });
}
// 페이지 로드 완료 시 섹터 메뉴 활성화
// ── 🌟 섹터 데이터 로드 및 초기화 ──
let _editingSectorId = null; // 현재 편집 중인 섹터 ID
let _editingStocks = [];     // 편집창에서 조작 중인 임시 종목 배열

// 기존 _sectorData 대신 localStorage에서 불러오도록 업그레이드
function loadSectorData() {
  const saved = localStorage.getItem('proPlusSectorData');
if (saved) {
    _sectorData = JSON.parse(saved);
  } else {
    // 저장된 데이터가 없으면 초기 기본값 생성
    _sectorData = {
      'semiconductor': { name: '半導體', stocks: ['005930', '000660'] },
      'it': { name: 'IT/소프트', stocks: ['035420', '035720'] },
      'auto': { name: '自動車', stocks: ['005380', '000270'] },
      'bio': { name: '바이오', stocks: ['207940', '068270'] },
      'untact': { name: '언택트', stocks: [] },
      'battery': { name: '二次電池', stocks: ['373220', '006400'] },
      'custom-1': { name: '섹터 지정1', stocks: [] },
      'custom-2': { name: '섹터 지정2', stocks: [] },
      'custom-3': { name: '섹터 지정3', stocks: [] },
      'custom-4': { name: '섹터 지정4', stocks: [] },
      'custom-5': { name: '섹터 지정5', stocks: [] },
      'custom-6': { name: '섹터 지정6', stocks: [] }
    };
  }
  renderSectorCards(); // ← 이 줄 추가
}

/* ── 섹터 카드 동적 렌더링 ── */
function renderSectorCards() {
  const row = document.getElementById('sector-card-row');
  if (!row) return;

  // ＋버튼 제외한 기존 카드 모두 제거
  row.querySelectorAll('.sector-card').forEach(c => c.remove());

  // _sectorData 기반으로 카드 생성
  Object.entries(_sectorData).forEach(([sectorId, info]) => {
    if (!info.name) return;
    const card = document.createElement('div');
    card.className = 'sector-card' + (sectorId.startsWith('custom-') ? ' custom' : ' preset');
    card.dataset.sector = sectorId;
    card.innerHTML = `
      <span class="sc-name">${info.name}</span>
      <span class="sc-del" onclick="sectorRemove('${sectorId}', event)">✕</span>`;
    row.insertBefore(card, row.querySelector('.sector-add-btn'));
  });

  // 이벤트 재연결
  initSectorMenu();
}

/* ── 섹터 추가 ── */
function sectorAddNew() {
  const name = prompt('새 섹터 이름을 입력하세요:');
  if (!name?.trim()) return;
    // ← 이 줄 추가
  if (Object.values(_sectorData).some(s => s.name === name.trim())) {
    showToast('⚠ 이미 같은 이름의 섹터가 있습니다.'); return;
  }
  const id = 'custom-' + Date.now();
  _sectorData[id] = { name: name.trim(), stocks: [] };
  saveSectorDataLocal();
  renderSectorCards();
}

/* ── 섹터 삭제 ── */
function sectorRemove(sectorId, e) {
  e.stopPropagation();
  if (!confirm(`"${_sectorData[sectorId]?.name}" 섹터를 삭제할까요?`)) return;
  delete _sectorData[sectorId];
  saveSectorDataLocal();
  renderSectorCards();
}

/* ── localStorage 저장 ── */
function saveSectorDataLocal() {
  localStorage.setItem('proPlusSectorData', JSON.stringify(_sectorData));
}

// Phase B에서 작성했던 필터링 로직 수정 (데이터 구조가 객체로 변경되었으므로)
function filterDashboardBySector(sectorId) {
  const stockCards = document.querySelectorAll('.wc-card'); 
  if (!sectorId) {
    stockCards.forEach(card => card.style.display = '');
    return;
  }
  const symbolsInSector = _sectorData[sectorId]?.stocks || [];
  stockCards.forEach(card => {
    const symbol = card.dataset.symbol || null;  // ← dataset.symbol 직접 참조
    card.style.display = (symbol && symbolsInSector.includes(symbol)) ? '' : 'none';
  });
}

// ── 🌟 섹터 편집 모달 열기/닫기 ──
// ── 🌟 [수정됨] 섹터 편집 모달 열기 ──
function openSectorEditModal(sectorId, cardElement) {
  _editingSectorId = sectorId;
  const sectorInfo = _sectorData[sectorId] || { name: cardElement.innerText, stocks: [] };
  
  // 편집용 임시 배열에 복사
  _editingStocks = [...sectorInfo.stocks]; 
  
  // 1. 이름 입력창 세팅 (Phase A에서 수정한 input 요소)
  const nameInput = document.getElementById('sector-edit-name');
  if (nameInput) {
    nameInput.value = sectorInfo.name;
    nameInput.disabled = !sectorId.startsWith('custom-'); // 프리셋은 이름 수정 불가
  }
  
  // 2. 검색창 초기화
  const searchInput = document.getElementById('source-stock-search');
  if (searchInput) {
    searchInput.value = '';
    searchInput.onkeyup = renderDualPanes; // 타자 칠 때마다 실시간 검색
  }

  // 3. 🌟 양쪽 화면 그리기 (에러가 나던 함수를 완전히 대체!)
  _selectedLeft = null;
  _selectedRight = null;
  renderDualPanes(); 
  
  // 4. 모달창 띄우기
  document.getElementById('sector-edit-backdrop').style.display = 'block';
  document.getElementById('sector-edit-modal').style.display = 'flex';
 // 🌟 [추가된 부분] 모달창이 화면에 나타난 직후에 드래그 기능을 확실하게 연결! 
  initSectorModalDrag();
}


// ── 🌟 [신규] 두 분할 패널(Target / Source) 렌더링 로직 ──
let _selectedLeft = null;  // 왼쪽 창에서 선택된 종목
let _selectedRight = null; // 오른쪽 창에서 선택된 종목

function renderDualPanes() {
  const leftList = document.getElementById('sector-target-list');
  const rightList = document.getElementById('sector-source-list');
  const searchKeyword = (document.getElementById('source-stock-search').value || '').toLowerCase();
  
  // 모달 안의 총 개수 업데이트
  const countSpan = document.getElementById('sector-edit-count');
  if(countSpan) countSpan.innerText = _editingStocks.length;

  // 1. 내 종목 & 관심 종목 합치기 (중복 제거 및 이름/코드 확보)
  const allMyStocks = [...getMyList(), ...getList()];
  const uniqueStocks = [];
  const seen = new Set();
  for (const s of allMyStocks) {
    if (!seen.has(s.symbol)) {
      seen.add(s.symbol);
      uniqueStocks.push(s);
    }
  }

  // 2. 왼쪽 패널 (현재 섹터에 들어온 종목들) 그리기
  leftList.innerHTML = '';
  _editingStocks.forEach(sym => {
    // 종목 코드(sym)를 가지고 내 리스트에서 이름 찾기
    const stockInfo = uniqueStocks.find(x => x.symbol === sym) || { name: '알 수 없음', market: '기타' };
    const isKr = /^[0-9]{6}$/.test(sym); // 6자리 숫자면 한국 주식
    const countryClass = isKr ? 'kr-card' : 'us-card'; // 한국 파랑, 미국 빨강 테두리
    
    const div = document.createElement('div');
    div.className = `sector-item ${countryClass} ${_selectedLeft === sym ? 'selected' : ''}`;
    div.innerHTML = `
      <div class="sector-item-info">
        <span class="sector-item-name">${stockInfo.name}</span>
        <span class="sector-item-symbol">${sym}</span>
      </div>
    `;
    // 클릭하면 선택됨 (회색 배경)
    div.onclick = () => { _selectedLeft = sym; _selectedRight = null; renderDualPanes(); };
    // 더블클릭하면 바로 오른쪽으로 빼버림
    div.ondblclick = () => { _selectedLeft = sym; moveStockFromSector(); }; 
    leftList.appendChild(div);
  });

  // 3. 오른쪽 패널 (내 종목/관심 종목 중 아직 추가 안 된 종목들) 그리기
  rightList.innerHTML = '';
  uniqueStocks.forEach(stock => {
    // 이미 왼쪽에 들어간 종목은 오른쪽에서 숨김
    if (_editingStocks.includes(stock.symbol)) return;
    
    // 검색어 필터링 (이름이나 코드에 검색어가 포함되어 있는지)
    if (searchKeyword && !stock.name.toLowerCase().includes(searchKeyword) && !stock.symbol.toLowerCase().includes(searchKeyword)) return;

    const isKr = /^[0-9]{6}$/.test(stock.symbol);
    const countryClass = isKr ? 'kr-card' : 'us-card';
    
    const div = document.createElement('div');
    div.className = `sector-item ${countryClass} ${_selectedRight === stock.symbol ? 'selected' : ''}`;
    div.innerHTML = `
      <div class="sector-item-info">
        <span class="sector-item-name">${stock.name}</span>
        <span class="sector-item-symbol">${stock.symbol}</span>
      </div>
    `;
    div.onclick = () => { _selectedRight = stock.symbol; _selectedLeft = null; renderDualPanes(); };
    div.ondblclick = () => { _selectedRight = stock.symbol; moveStockToSector(); };
    rightList.appendChild(div);
  });
}

// ── 🌟 [신규] 좌우 이동 버튼 로직 (<<, >>) ──
function moveStockToSector() {
  if (!_selectedRight) return; // 선택된 게 없으면 무시
  if (!_editingStocks.includes(_selectedRight)) {
    _editingStocks.push(_selectedRight);
  }
  _selectedRight = null; // 이동 후 선택 해제
  renderDualPanes(); // 화면 새로고침
}

function moveStockFromSector() {
  if (!_selectedLeft) return;
  _editingStocks = _editingStocks.filter(sym => sym !== _selectedLeft); // 배열에서 제거
  _selectedLeft = null;
  renderDualPanes();
}

function closeSectorEdit() {
  document.getElementById('sector-edit-backdrop').style.display = 'none';
  document.getElementById('sector-edit-modal').style.display = 'none';
  _editingSectorId = null;
}

// ── 🌟 모달창 내 종목 리스트 렌더링 ──
function renderSectorEditList() {
  const listDiv = document.getElementById('sector-edit-list');
  document.getElementById('sector-edit-count').innerText = _editingStocks.length;
  listDiv.innerHTML = '';
  
  if (_editingStocks.length === 0) {
    listDiv.innerHTML = '<div style="color:var(--muted); text-align:center; margin-top:50px; font-size:12px;">등록된 종목이 없습니다.</div>';
    return;
  }

  _editingStocks.forEach(sym => {
    const item = document.createElement('div');
    item.style.cssText = 'display:flex; justify-content:space-between; align-items:center; background:var(--panel); padding:6px 10px; border-radius:4px; border:1px solid var(--border); font-size:12px; font-weight:bold;';
    item.innerHTML = `
      <span>${sym}</span>
      <button style="background:none; border:none; color:#ef4444; cursor:pointer; font-weight:bold;" onclick="removeSectorStockUi('${sym}')">✕</button>
    `;
    listDiv.appendChild(item);
  });
}

// ── 🌟 종목 추가 및 삭제 로직 ──
function addSectorStockUi() {
  const input = document.getElementById('sector-edit-symbol');
  const sym = input.value.trim().toUpperCase(); // 알파벳(미국주식) 고려하여 대문자 변환
  if (!sym) return;
  
  if (_editingStocks.includes(sym)) {
    showToast('⚠ 이미 섹터에 포함된 종목입니다.');
    return;
  }
  
  _editingStocks.push(sym);
  input.value = '';
  renderSectorEditList();
}

function removeSectorStockUi(sym) {
  _editingStocks = _editingStocks.filter(s => s !== sym);
  renderSectorEditList();
}

// ── 🌟 저장 버튼 클릭 시 처리 ──
function saveSectorEdit() {
  if (!_editingSectorId) return;
  
  const newName = document.getElementById('sector-edit-name').value.trim() || '이름 없음';
  
  // 데이터 업데이트
  _sectorData[_editingSectorId] = {
    name: newName,
    stocks: [..._editingStocks]
  };
  
  // 로컬 스토리지에 저장 (영구 보존)
  localStorage.setItem('proPlusSectorData', JSON.stringify(_sectorData));
  
  // 대시보드 UI(섹터 박스 이름) 업데이트
  const card = document.querySelector(`.sector-card[data-sector="${_editingSectorId}"]`) || 
               document.querySelector(`.sector-card[data-index="${_editingSectorId.replace('custom-','')}"]`);
  
if (card) {
    const nameSpan = card.querySelector('.sc-name');  // ← sc-name span만 타겟
    if (nameSpan) nameSpan.textContent = newName;
    if (_editingSectorId.startsWith('custom-')) {
      card.classList.remove('empty');
    }
  }
  
  showToast(`✅ '${newName}' 섹터가 저장되었습니다.`);
  closeSectorEdit();
  
  // 만약 현재 필터링해서 보고 있던 섹터를 수정한 거라면 화면 새로고침
  if (_activeSector === _editingSectorId) {
    filterDashboardBySector(_activeSector);
  }
}
document.addEventListener('DOMContentLoaded', () => {
    loadSectorData();
    // 저장된 정렬 상태 복원
	updateSortBtnUI(); // ← 저장된 정렬 상태 UI 복원
});
// ── 🌟 모달창 드래그(이동) 및 크기 조절 활성화 함수 ──
// ── 🌟 [신규/수정] 타이밍 이슈를 해결한 무적의 드래그 로직 ──
let _isSectorDragInit = false; // 드래그 이벤트가 여러 번 등록되지 않도록 막는 플래그

function initSectorModalDrag() {
  if (_isSectorDragInit) return; // 이미 세팅되었다면 패스

  const modal = document.getElementById('sector-edit-modal');
  const handle = document.getElementById('sector-drag-handle');
  
  // 창이 아직 화면에 없으면 패스 (에러 방지)
  if (!modal || !handle) return;

  let isDragging = false;
  let offsetX, offsetY;

  // 1. 마우스를 꾹 눌렀을 때 (드래그 시작)
  handle.addEventListener('mousedown', (e) => {
    if (e.target.tagName === 'BUTTON') return; // X 닫기 버튼 누를 때는 무시
    
    isDragging = true;
    
    // 🌟 핵심: CSS의 transform(중앙 정렬)과 충돌하지 않도록 현재 위치를 픽셀(px)로 고정
    const rect = modal.getBoundingClientRect();
    modal.style.transform = 'none'; 
    modal.style.margin = '0';
    modal.style.top = rect.top + 'px';
    modal.style.left = rect.left + 'px';
    
    // 마우스 포인터와 모달창 모서리 사이의 간격 계산
    offsetX = e.clientX - rect.left;
    offsetY = e.clientY - rect.top;
  });

  // 2. 마우스를 움직일 때 (드래그 중)
  document.addEventListener('mousemove', (e) => {
    if (!isDragging) return;
    // 계산된 간격을 유지하며 모달창 이동
    modal.style.left = (e.clientX - offsetX) + 'px';
    modal.style.top = (e.clientY - offsetY) + 'px';
  });

  // 3. 마우스 클릭을 뗐을 때 (드래그 종료)
  document.addEventListener('mouseup', () => {
    isDragging = false;
  });

  _isSectorDragInit = true; // 세팅 완료!
}

// chart.js 에 추가
async function updateSectorPerformance() {
  for (const sId in _sectorData) {
    const stocks = _sectorData[sId].stocks;
    if (!stocks || stocks.length === 0) continue;

    // 서버 호출 없이 DOM 카드에서 직접 등락률 읽기
    let total = 0, count = 0;
    stocks.forEach(sym => {
      const el = document.getElementById(`wcp-${sym}`);
      if (!el) return;
      const txt = el.textContent.replace(/[^0-9.\-+%]/g, '');
      const match = el.closest('.wc-card')?.querySelector('.wc-chgr');
      if (!match) return;
      const val = parseFloat(match.textContent.replace(/[^0-9.\-]/g, ''));
      if (!isNaN(val)) { total += val; count++; }
    });

    if (count === 0) continue;
    const avg = (total / count).toFixed(2);

    const card = document.querySelector(`.sector-card[data-sector="${sId}"]`);
    if (!card) continue;

    let badge = card.querySelector('.sector-perf');
    if (!badge) {
      badge = document.createElement('span');
      badge.className = 'sector-perf';
      badge.style.cssText = 'font-size:10px; margin-left:4px; font-weight:bold;';
      card.querySelector('.sc-name')?.after(badge);
    }
    const num = parseFloat(avg);
    badge.textContent = (num > 0 ? '+' : '') + avg + '%';
    badge.style.color = num > 0 ? '#ef4444' : (num < 0 ? '#3b82f6' : 'var(--muted)');
  }
}
// 🌟 AI 분석 버튼을 눌렀을 때 실행되는 핵심 함수
// 🌟 [수정본] 어떤 상황에서도 차트 영역을 찾아내는 AI 분석 함수
/* ══════════════════════════════════
   차트 캡처 → AI 기술 분석 (단일 정의)
══════════════════════════════════ */
async function analyzeChartWithAi(cid) {
  const btn = event.currentTarget;

  // 버튼 기준으로 차트 바디 탐색
  const titleBar = btn.closest('.chart-modal-bar') || btn.closest('.win-ctrl-right')?.parentElement;
  let chartEl = titleBar ? titleBar.nextElementSibling : null;

  if (!chartEl || !chartEl.classList.contains('chart-modal-body')) {
    const parentModal = btn.closest('.chart-modal') || btn.closest('.modal-content');
    if (parentModal) chartEl = parentModal.querySelector('.chart-modal-body');
  }

  // 폴백: widgetId 기반 탐색
  if (!chartEl) chartEl = document.getElementById(`chart-cont-${cid}`);

  if (!chartEl) {
    showToast("❌ 차트 본문을 찾을 수 없습니다.");
    return;
  }

  showToast("🤖 AI 전략가가 차트를 정밀 분석 중입니다...");

  try {
	const w = activeWidgets[cid];
	if (!w || !w.chart) {
	  showToast("❌ 차트 인스턴스를 찾을 수 없습니다.");
	  return;
	}
	const canvas = w.chart.takeScreenshot();

    // 이미지 전역 변수 세팅 (rpSendAi가 읽어감)
    window.aiPendingImage = {
      base64: canvas.toDataURL('image/png').split(',')[1],
      mediaType: 'image/png',
      previewUrl: canvas.toDataURL('image/png')
    };

    // 우측 패널 입력창에 질문 세팅 후 rpSendAi 호출
    const inp = document.getElementById('rp-ai-input');
    if (inp) inp.value = "이 차트의 현재 추세와 주요 지표를 분석해서 기술적 리포트를 작성해줘.";

    if (typeof rpSendAi === 'function') {
      rpSendAi();
    } else {
      showToast("❌ AI 전송 함수(rpSendAi)를 찾을 수 없습니다.");
    }

  } catch (err) {
    console.error(err);
    showToast("❌ 캡처 중 오류 발생: " + err.message);
  }
}

// 5분마다 섹터 현황 갱신
setInterval(updateSectorPerformance, 5 * 60 * 1000);

function showToast(msg){const t=document.createElement('div');t.className='toast';t.innerText=msg;document.getElementById('toast-cont').appendChild(t);setTimeout(()=>t.remove(),4200);}
