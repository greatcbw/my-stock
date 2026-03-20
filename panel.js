// panel.js — 우측 패널: 검색, 메모/일지, AI 채팅

// API는 chart.js에서 선언됨

/* ══════════════════════════════════
   우측 패널 초기화
══════════════════════════════════ */
function initRightPanel() {
  const layout = document.getElementById('layout');
  if (!layout || document.getElementById('right-panel')) return;

  const rp = document.createElement('div');
  rp.id = 'right-panel';
  rp.style.cssText = 'width:calc(100vw/7);min-width:180px;max-width:300px;flex-shrink:0;height:100%;background:var(--panel);border-left:1px solid var(--border);display:flex;flex-direction:column;overflow:hidden;';
  rp.innerHTML = `
    <!-- 검색 -->
    <div class="rp-search-wrap">
      <input id="rp-search-input" class="rp-input" placeholder="🔍 종목명 / 코드 검색" onkeyup="rpDebounceSearch()">
      <div id="rp-search-results" class="rp-search-results" style="display:none;"></div>
      <div id="rp-add-btns" class="rp-add-btns" style="display:none;">
        <button class="btn-primary rp-add-btn" onclick="rpAddMyStock()">＋ 내종목</button>
        <button class="btn-secondary rp-add-btn" onclick="rpAddWatch()">＋ 관심</button>
      </div>
    </div>

    <!-- 메모/일지 (30% 고정) -->
    <div id="rp-memo" style="flex:0 0 30%;display:flex;flex-direction:column;border-bottom:1px solid var(--border);min-height:0;overflow:hidden;">
      <div class="rp-section-header">
        <span>📝 투자 일지</span>
        <button class="rp-save-btn" onclick="rpSaveMemo()">💾 저장</button>
      </div>
      <div class="rp-cal-bar">
        <button class="rp-cal-btn" onclick="rpMemoNav(-1)">◀</button>
        <select id="rp-memo-date" class="rp-cal-select" onchange="rpLoadMemo(this.value)"></select>
        <button class="rp-cal-btn" onclick="rpMemoNav(1)">▶</button>
        <button class="rp-cal-btn" onclick="rpLoadMemo('today')" title="오늘">📅</button>
      </div>
      <textarea id="rp-memo-text" class="rp-memo-text" placeholder="오늘의 투자 메모를 기록하세요..."></textarea>
    </div>

    <!-- AI 채팅 -->
    <div id="rp-ai" style="flex:1;display:flex;flex-direction:column;min-height:0;overflow:hidden;">
      <div class="rp-section-header" style="flex-shrink:0;">
        <span>🤖 AI 어시스턴트</span>
        <button class="rp-cal-btn" onclick="rpClearChat()" title="대화 초기화">🗑</button>
      </div>
      <!-- 대화 내역 -->
      <div id="rp-ai-msgs" class="rp-ai-msgs"></div>
      <!-- 질문 입력 (고정) -->
      <div class="rp-ai-input-wrap">
        <textarea id="rp-ai-input" class="rp-ai-input"
               placeholder="질문 입력... (Enter 전송 / Shift+Enter 줄바꿈)"
               onkeydown="if(event.key==='Enter'&&!event.shiftKey){event.preventDefault();rpSendAi();}"></textarea>
      </div>
    </div>`;

  layout.appendChild(rp);

  // 검색 외부 클릭 닫기
  document.addEventListener('click', e => {
    const res  = document.getElementById('rp-search-results');
    const wrap = document.querySelector('.rp-search-wrap');
    if (res && wrap && !wrap.contains(e.target)) res.style.display = 'none';
  });

  rpLoadMemo('today');
  rpRefreshMemoDates();
  rpAddChatMsg('bot', '안녕하세요! 주식 관련 궁금한 것을 물어보세요 😊');
}

/* ══════════════════════════════════
   검색
══════════════════════════════════ */
let _rpSearchTimer = null;
let _rpSelected    = null;

function rpDebounceSearch() {
  clearTimeout(_rpSearchTimer);
  _rpSearchTimer = setTimeout(rpDoSearch, 280);
}

async function rpDoSearch() {
  const q   = document.getElementById('rp-search-input')?.value.trim();
  const res = document.getElementById('rp-search-results');
  if (!res || !q) { if (res) res.style.display='none'; return; }
  try {
    const data = await fetch(`${API}/search?q=${encodeURIComponent(q)}`).then(r=>r.json());
    if (!data.length) { res.style.display='none'; return; }
    res.style.display = 'block';
    res.innerHTML = data.slice(0,10).map(s => {
      const mc = (window.mktCls || (() => ''))(s.market);
      const name = s.name.replace(/'/g, "\\'");
      return `<div class="si" onclick="rpSelectResult('${s.symbol}','${name}','${s.market||'기타'}')">
        <span class="market-badge ${mc}">${s.market}</span>
        <span class="si-name">${s.name}</span>
        <span class="si-code">${s.symbol}</span>
      </div>`;
    }).join('');
  } catch(e) { res.style.display='none'; }
}

function rpSelectResult(symbol, name, market) {
  _rpSelected = { symbol, name, market };
  const inp = document.getElementById('rp-search-input');
  const res = document.getElementById('rp-search-results');
  const btn = document.getElementById('rp-add-btns');
  if (inp) inp.value = `${name} (${symbol})`;
  if (res) res.style.display = 'none';
  if (btn) btn.style.display = 'flex';
  window.curSym = symbol; window.curName = name; window.curMkt = market;
}

function rpAddMyStock() {
  if (!_rpSelected) { window.showToast?.('⚠ 종목을 먼저 선택하세요'); return; }
  const list = (window.getMyList || (() => []))();
  if (list.find(x => x.symbol === _rpSelected.symbol)) { window.showToast?.('이미 내종목에 있습니다'); return; }
  list.push(_rpSelected);
  window.saveMyList?.(list);
  window.loadWatch?.();
  window.showToast?.(`✅ ${_rpSelected.name} 내종목 추가`);
  rpClearSearch();
}

function rpAddWatch() {
  if (!_rpSelected) { window.showToast?.('⚠ 종목을 먼저 선택하세요'); return; }
  const list = (window.getList || (() => []))();
  if (list.find(x => x.symbol === _rpSelected.symbol)) { window.showToast?.('이미 관심종목에 있습니다'); return; }
  list.push(_rpSelected);
  window.saveList?.(list);
  window.loadWatch?.();
  window.showToast?.(`✅ ${_rpSelected.name} 관심종목 추가`);
  rpClearSearch();
}

function rpClearSearch() {
  const inp = document.getElementById('rp-search-input');
  const btn = document.getElementById('rp-add-btns');
  if (inp) inp.value = '';
  if (btn) btn.style.display = 'none';
  _rpSelected = null;
}

/* ══════════════════════════════════
   날짜별 메모
══════════════════════════════════ */
let _rpMemoDate  = '';
let _rpMemoDates = [];

function _rpTodayStr() { return new Date().toISOString().slice(0,10); }

async function rpRefreshMemoDates() {
  try {
    const res = await fetch(`${API}/memo/list`);
    _rpMemoDates = await res.json();
    const today = _rpTodayStr();
    if (!_rpMemoDates.includes(today)) _rpMemoDates = [today, ..._rpMemoDates];
    const sel = document.getElementById('rp-memo-date');
    if (!sel) return;
    sel.innerHTML = _rpMemoDates.map(d =>
      `<option value="${d}">${d}${d === today ? ' (오늘)' : ''}</option>`
    ).join('');
    sel.value = _rpMemoDate || today;
  } catch(e) {}
}

async function rpLoadMemo(date) {
  if (date === 'today') date = _rpTodayStr();
  _rpMemoDate = date;
  try {
    const res  = await fetch(`${API}/memo/${date}`);
    const data = await res.json();
    const el   = document.getElementById('rp-memo-text');
    if (el) el.value = data.content || '';
    const sel = document.getElementById('rp-memo-date');
    if (sel) {
      if (!Array.from(sel.options).find(o => o.value === date)) {
        const opt = document.createElement('option');
        opt.value = date; opt.textContent = date;
        sel.insertBefore(opt, sel.firstChild);
      }
      sel.value = date;
    }
  } catch(e) {}
}

async function rpSaveMemo() {
  const val  = document.getElementById('rp-memo-text')?.value || '';
  const date = _rpMemoDate || _rpTodayStr();
  try {
    await fetch(`${API}/memo/${date}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ content: val })
    });
    window.showToast?.(`✅ ${date} 메모 저장됨`);
    await rpRefreshMemoDates();
  } catch(e) { window.showToast?.('❌ 메모 저장 실패'); }
}

function rpMemoNav(dir) {
  const sel  = document.getElementById('rp-memo-date');
  if (!sel) return;
  const opts = Array.from(sel.options).map(o => o.value);
  const idx  = opts.indexOf(_rpMemoDate);
  const next = idx + dir;
  if (next >= 0 && next < opts.length) rpLoadMemo(opts[next]);
}

/* ══════════════════════════════════
   AI 채팅
══════════════════════════════════ */
let _rpAiHistory = [];

function rpAddChatMsg(role, text) {
  const msgs = document.getElementById('rp-ai-msgs');
  if (!msgs) return;
  const div = document.createElement('div');
  div.className = `rp-ai-msg rp-ai-${role === 'user' ? 'user' : 'bot'}`;
  div.textContent = text;
  msgs.appendChild(div);
  msgs.scrollTop = msgs.scrollHeight;
  return div;
}

function rpClearChat() {
  const msgs = document.getElementById('rp-ai-msgs');
  if (msgs) msgs.innerHTML = '';
  _rpAiHistory = [];
  rpAddChatMsg('bot', '대화가 초기화되었습니다. 새로운 질문을 입력하세요 😊');
}

async function rpSendAi() {
  const input = document.getElementById('rp-ai-input');
  if (!input) return;
  const text = input.value.trim();
  const img  = window.aiPendingImage; // 차트 캡처 이미지 (있을 경우)

  if (!text && !img) return;

  // 1. 사용자 메시지 표시
  input.value = '';
  input.style.height = 'auto';
  rpAddChatMsg('user', text || '📊 차트 이미지를 첨부했습니다.');

  // 2. 이미지 미리보기 (차트 분석 요청 시)
  if (img) {
    const msgs = document.getElementById('rp-ai-msgs');
    if (msgs) {
      const prevDiv = document.createElement('div');
      prevDiv.className = 'rp-ai-msg rp-ai-user';
      prevDiv.innerHTML = `<img src="${img.previewUrl}" style="max-width:100%;border-radius:6px;margin-top:4px;">`;
      msgs.appendChild(prevDiv);
      msgs.scrollTop = msgs.scrollHeight;
    }
  }

  // 3. 히스토리에 추가 (이미지 포함 시 멀티모달 포맷)
  const userContent = img
    ? [
        { type: 'image', source: { type: 'base64', media_type: img.mediaType, data: img.base64 } },
        { type: 'text',  text: text || '이 차트의 현재 추세와 주요 지표를 분석해서 기술적 리포트를 작성해줘.' }
      ]
    : text;
  _rpAiHistory.push({ role: 'user', content: userContent });

  // 4. 봇 응답 버블 생성
  const loadDiv = rpAddChatMsg('bot', '⏳ 분석 중...');

  // 버튼 비활성화
  const btn = document.querySelector('.rp-ai-send-btn');
  if (btn) { btn.disabled = true; btn.textContent = '⏳'; }

  // 이미지 전역 변수 초기화
  window.aiPendingImage = null;

  try {
    const system = img
      ? '당신은 전문 주식 기술 분석가입니다. 차트 이미지를 보고 캔들 패턴, 이동평균선, RSI/MACD 등 보조지표를 종합 분석하여 매수/매도/관망 전략을 한국어 평문으로 간결하게 작성하세요.'
      : '당신은 주식 투자 전문 AI 어시스턴트입니다. 간결하고 명확하게 한국어로 답변하세요. 마크다운 사용을 지양하고 평문으로 작성하세요.';

    const res = await fetch(`${API}/ai/chat`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        messages: _rpAiHistory.slice(-10),
        system
      })
    });

    if (!res.ok) throw new Error('서버 응답 오류 ' + res.status);

    // 5. 스트리밍 수신
    const reader  = res.body.getReader();
    const decoder = new TextDecoder('utf-8');
    let fullReply = '';
    loadDiv.textContent = '';

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      fullReply += decoder.decode(value, { stream: true });
      loadDiv.innerHTML = fullReply.replace(/\n/g, '<br>');
      const msgs = document.getElementById('rp-ai-msgs');
      if (msgs) msgs.scrollTop = msgs.scrollHeight;
    }

    // 6. 히스토리 저장
    _rpAiHistory.push({ role: 'assistant', content: fullReply });
    if (_rpAiHistory.length > 20) _rpAiHistory = _rpAiHistory.slice(-20);

  } catch (err) {
    loadDiv.innerHTML = `<span style="color:#ef4444;">오류: ${err.message}</span>`;
  } finally {
    if (btn) { btn.disabled = false; btn.textContent = '➤'; }
  }
}

/* ══════════════════════════════════
   DOM 로드 후 초기화
══════════════════════════════════ */
document.addEventListener('DOMContentLoaded', initRightPanel);
