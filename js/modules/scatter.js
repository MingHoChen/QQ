import { SharedData, runWorkerTask, showError, showSuccess } from '../utils.js';

let scatterChart = null;

export function initScatterEvents() {
    const sourceMode = document.getElementById('scatter-source-mode');
    if (sourceMode) sourceMode.addEventListener('change', window.handleScatterSourceChange);
    
    const fileInput = document.getElementById('scatter-file-input');
    if (fileInput) fileInput.addEventListener('change', handleScatterFileSelect);
}

window.handleScatterSourceChange = function() {
    const mode = document.getElementById('scatter-source-mode').value;
    document.getElementById('scatter-file-upload').style.display = mode === 'new' ? 'block' : 'none';
    if (mode === 'cpk') {
        if (SharedData.hasData) populateScatterSelectors(SharedData.headers);
        else showError("請先在 Cpk 頁面上傳數據或選擇上傳新檔案");
    }
}

async function handleScatterFileSelect(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    try {
        const buffer = await file.arrayBuffer();
        const res = await runWorkerTask('parseFirstSheet', { buffer });
        if (res.headers.length > 0) {
            SharedData.tempScatter = res;
            populateScatterSelectors(res.headers);
            showSuccess("散佈圖數據已讀取");
        }
    } catch (err) {
        showError("讀取失敗: " + err.message);
    }
}

function populateScatterSelectors(headers) {
    const xSel = document.getElementById('scatter-x-col');
    const ySel = document.getElementById('scatter-y-col');
    const gSel = document.getElementById('scatter-group-col');
    [xSel, ySel, gSel].forEach(s => s.innerHTML = '');
    gSel.add(new Option("-- 不分組 --", ""));
    headers.forEach(h => {
        xSel.add(new Option(h, h));
        ySel.add(new Option(h, h));
        gSel.add(new Option(h, h));
    });
    if (ySel.options.length > 1) ySel.selectedIndex = 1;
}

window.calculateScatter = function() {
    const mode = document.getElementById('scatter-source-mode').value;
    let data = [];
    if (mode === 'cpk') data = SharedData.rawJson;
    else if (SharedData.tempScatter) data = SharedData.tempScatter.data;
    
    if (!data || data.length === 0) return showError("無數據可分析");
    
    const xKey = document.getElementById('scatter-x-col').value;
    const yKey = document.getElementById('scatter-y-col').value;
    const gKey = document.getElementById('scatter-group-col').value;
    if (!xKey || !yKey) return showError("請選擇 X 和 Y 軸欄位");
    
    let datasets = {};
    let xSum = 0, ySum = 0, n = 0, xSqSum = 0, ySqSum = 0, xySum = 0;
    
    data.forEach(row => {
        let x = parseFloat(String(row[xKey]).replace(/[^0-9.\-]/g, ''));
        let y = parseFloat(String(row[yKey]).replace(/[^0-9.\-]/g, ''));
        if (!isNaN(x) && !isNaN(y)) {
            let group = "Data";
            if (gKey && row[gKey] !== undefined) group = String(row[gKey]);
            if (!datasets[group]) datasets[group] = [];
            datasets[group].push({ x, y });
            xSum += x; ySum += y; xSqSum += x * x; ySqSum += y * y; xySum += x * y; n++;
        }
    });
    
    if (n < 2) return showError("有效數據點不足，無法計算相關性");
    
    let num = n * xySum - xSum * ySum;
    let den = Math.sqrt((n * xSqSum - xSum * xSum) * (n * ySqSum - ySum * ySum));
    let r = den === 0 ? 0 : num / den;
    document.getElementById('val-r').innerText = r.toFixed(4);
    
    drawScatterChart(datasets, xKey, yKey);
}

function drawScatterChart(datasetsObj, xLabel, yLabel) {
    const ctx = document.getElementById('scatter-chart').getContext('2d');
    if (scatterChart) scatterChart.destroy();
    
    const colors = ['#3498db', '#e74c3c', '#2ecc71', '#f1c40f', '#9b59b6', '#34495e'];
    let finalDatasets = [];
    let i = 0;
    for (const [grp, pts] of Object.entries(datasetsObj)) {
        finalDatasets.push({ label: grp, data: pts, backgroundColor: colors[i % colors.length], pointRadius: 5 });
        i++;
    }
    
    scatterChart = new Chart(ctx, {
        type: 'scatter',
        data: { datasets: finalDatasets },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: { title: { display: true, text: xLabel }, type: 'linear', position: 'bottom' },
                y: { title: { display: true, text: yLabel }, type: 'linear' }
            },
            plugins: {
                legend: { display: true },
                zoom: { 
                    zoom: { wheel: { enabled: true }, pinch: { enabled: true }, mode: 'xy' }, 
                    pan: { enabled: true, mode: 'xy' } 
                }
            }
        }
    });
}
