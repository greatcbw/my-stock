// forecast.js — ARIMA + Prophet + 앙상블 예측선

async function drawForecastLine(w) {
    ['forecastSeries', 'validSeries', 'errorSeries',
     'prophetSeries', 'ensembleSeries'].forEach(key => {
        if (w[key]) {
            try { w.chart.removeSeries(w[key]); } catch(e) {}
            w[key] = null;
        }
    });

    function isAlive() {
        return !w._disposed && !!w.chart;
    }

    try {
        const response = await fetch('/api/forecast/' + w.symbol);
        if (!isAlive()) return;
        if (!response.ok) {
            const err = await response.json().catch(() => ({}));
            throw new Error(err.error || '서버 응답 오류');
        }

        const data = await response.json();
        if (!isAlive()) return;
        if (data.error) throw new Error(data.error);

        const forecast   = data.forecast   || [];
        const validation = data.validation || [];
        const prophet    = data.prophet    || [];
        const ensemble   = data.ensemble   || [];

        if (!isAlive()) return;

        // ── A: 과거 검증 예측선 (하늘색 긴점선) ──
        if (validation.length) {
            w.validSeries = w.chart.addLineSeries({
                color: '#38bdf8', lineWidth: 2, lineStyle: 3,
                title: '검증예측(A)',
            });
            w.validSeries.setData(validation.map(d => ({
                time: d.time, value: d.predicted
            })));
            w.errorSeries = w.chart.addLineSeries({
                color: 'rgba(148,163,184,0.5)', lineWidth: 1, lineStyle: 4,
                title: '오차(A)',
            });
            w.errorSeries.setData(validation.map(d => ({
                time: d.time, value: d.actual + d.error * 0.5
            })));
        }

        // ── B: ARIMA 미래 예측선 (핑크 점선) ──
        if (forecast.length) {
            w.forecastSeries = w.chart.addLineSeries({
                color: '#ff2972', lineWidth: 2, lineStyle: 2,
                title: 'ARIMA(B)',
            });
            w.forecastSeries.setData(forecast);
        }

        // ── C: Prophet 미래 예측선 (보라 점선) ──
        if (prophet.length) {
            w.prophetSeries = w.chart.addLineSeries({
                color: '#a855f7', lineWidth: 2, lineStyle: 2,
                title: 'Prophet(C)',
            });
            w.prophetSeries.setData(prophet);
        }

        // ── D: 앙상블 (주황 실선) ──
        if (ensemble.length) {
            w.ensembleSeries = w.chart.addLineSeries({
                color: '#f59e0b', lineWidth: 2.5, lineStyle: 0,
                title: '앙상블(D)',
            });
            w.ensembleSeries.setData(ensemble);
        }

        // ★ fitContent/scrollToPosition 제거 — 모달 초기 줌 유지

    } catch (error) {
        console.warn('[예측선] 오류:', error.message);
    }
}

function calculateError(actualData, forecastData) {
    const errorData = [];
    const map = new Map();
    actualData.forEach(d => map.set(d.time, d.value ?? d.close));
    forecastData.forEach(f => {
        const actual = map.get(f.time);
        if (actual !== undefined) {
            errorData.push({ time: f.time, value: actual - f.value });
        }
    });
    return errorData;
}
