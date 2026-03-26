import { SharedData, runWorkerTask, showError, showSuccess } from '../utils.js';

let paretoChart = null;

export function initParetoEvents() {
    const sourceMode = document.getElementById('pareto-source-mode');
    if (sourceMode) sourceMode.addEventListener('change', window.handleParetoSourceChange);
    
    const fileInput = document.getElementById('pareto-file-input');
    if (fileInput) fileInput.addEventListener('change', handleParetoFileSelect);
}

window.handleParetoSourceChange = function() {
    const mode = document.getElementById('pareto-source-mode').value;
    document.getElementById('pareto-file-upload').style.display = mode === 'new' ? 'block' : 'none';
    if (mode === 'cpk') {
        if (SharedData.hasData) populateParetoSelectors(SharedData.headers);
        else showError("請先在 Cpk 頁面上傳數據或選擇上傳新檔案");
    }
}

async function handleParetoFileSelect(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    try {
        const buffer = await file.arrayBuffer();
        const res = await runWorkerTask('parseFirstSheet', { buffer });
        if (res.headers.length > 0) {
            SharedData.tempPareto = res;
            populateParetoSelectors(res.headers);
            showSuccess("柏拉圖數據已讀取");
        }
    } catch (err) {
        showError("讀取失敗: " + err.message);
    }
}

function populateParetoSelectors(headers) {
    const catSel = document.getElementById('pareto-cat-col');
    const valSel = document.getElementById('pareto-val-col');
    catSel.innerHTML = '';
    valSel.innerHTML = '<option value="">-- 無 (計算出現次數) --</option>';
    headers.forEach(h => {
        catSel.add(new Option(h, h));
        valSel.add(new Option(h, h));
    });
}

window.calculatePareto = function() {
    const mode = document.getElementById('pareto-source-mode').value;
    let data = [];
    if (mode === 'cpk') data = SharedData.rawJson;
    else if (SharedData.tempPareto) data = SharedData.tempPareto.data;
    
    if (!data || data.length === 0) return showError("無數據可分析");
    
    const catKey = document.getElementById('pareto-cat-col').value;
    const valKey = document.getElementById('pareto-val-col').value;
    if (!catKey) return showError("請選擇類別欄位");
    
    let counts = {};
    let total = 0;
    data.forEach(row => {
        const cat = row[catKey];
        if (cat === undefined || cat === null || String(cat).trim() === "") return;
        let val = 1;
        if (valKey) {
            val = parseFloat(String(row[valKey]).replace(/[^0-9.\-]/g, ''));
            if (isNaN(val)) val = 0;
        }
        if (!counts[cat]) counts[cat] = 0;
        counts[cat] += val;
        total += val;
    });
    
    let sorted = Object.entries(counts).sort((a, b) => b[1] - a[1]);
    let cum = 0;
    let chartLabels = [];
    let chartDataBar = [];
    let chartDataLine = [];
    
    sorted.forEach(([k, v]) => {
        cum += v;
        chartLabels.push(k);
        chartDataBar.push(v);
        chartDataLine.push((cum / total) * 100);
    });
    
    drawParetoChart(chartLabels, chartDataBar, chartDataLine);
}

function drawParetoChart(labels, bars, lines) {
    const ctx = document.getElementById('pareto-chart').getContext('2d');
    if (paretoChart) paretoChart.destroy();
    paretoChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [
                { label: '累積百分比 (%)', data: lines, type: 'line', borderColor: '#e74c3c', yAxisID: 'y1', tension: 0.1, pointRadius: 4 },
                { label: '數量/數值', data: bars, backgroundColor: '#3498db', yAxisID: 'y' }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: { beginAtZero: true, title: { display: true, text: '數量' } },
                y1: { beginAtZero: true, max: 100, position: 'right', title: { display: true, text: '累積百分比 (%)' }, grid: { drawOnChartArea: false } }
            }
        }
    });
}
