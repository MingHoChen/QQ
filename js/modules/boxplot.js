import { SharedData, runWorkerTask, showError, showSuccess } from '../utils.js';

let boxplotChart = null;

export function initBoxplotEvents() {
    const sourceMode = document.getElementById('boxplot-source-mode');
    if (sourceMode) sourceMode.addEventListener('change', window.handleBoxplotSourceChange);
    
    const fileInput = document.getElementById('boxplot-file-input');
    if (fileInput) fileInput.addEventListener('change', handleBoxplotFileSelect);
}

window.handleBoxplotSourceChange = function() {
    const mode = document.getElementById('boxplot-source-mode').value;
    document.getElementById('boxplot-file-upload').style.display = mode === 'new' ? 'block' : 'none';
    if (mode === 'cpk') {
        if (SharedData.hasData) populateBoxplotSelectors(SharedData.headers);
        else showError("請先在 Cpk 頁面上傳數據或選擇上傳新檔案");
    }
}

async function handleBoxplotFileSelect(e) {
    const file = e.target.files[0];
    if (!file) return;
    try {
        const buffer = await file.arrayBuffer();
        const res = await runWorkerTask('parseFirstSheet', { buffer });
        if (res.headers.length > 0) {
            SharedData.tempBoxplot = res;
            populateBoxplotSelectors(res.headers);
            showSuccess("盒鬚圖數據已讀取");
        }
    } catch (err) {
        showError("讀取失敗: " + err.message);
    }
}

function populateBoxplotSelectors(headers) {
    const groupSel = document.getElementById('boxplot-group-col');
    const valueSel = document.getElementById('boxplot-value-col');
    [groupSel, valueSel].forEach(s => s.innerHTML = '');
    headers.forEach(h => {
        groupSel.add(new Option(h, h));
        valueSel.add(new Option(h, h));
    });
    if (valueSel.options.length > 1) valueSel.selectedIndex = 1;
}

window.calculateBoxplot = function() {
    const mode = document.getElementById('boxplot-source-mode').value;
    let data = [];
    if (mode === 'cpk') data = SharedData.rawJson;
    else if (SharedData.tempBoxplot) data = SharedData.tempBoxplot.data;
    
    if (!data || data.length === 0) return showError("無數據可分析");
    
    const groupKey = document.getElementById('boxplot-group-col').value;
    const valueKey = document.getElementById('boxplot-value-col').value;
    if (!groupKey || !valueKey) return showError("請選擇分組和數值欄位");
    
    let grouped = {};
    data.forEach(row => {
        const group = row[groupKey];
        if (group === undefined || group === null) return;
        let val = parseFloat(String(row[valueKey]).replace(/[^0-9.\-]/g, ''));
        if (!isNaN(val)) {
            if (!grouped[group]) grouped[group] = [];
            grouped[group].push(val);
        }
    });
    
    if (Object.keys(grouped).length === 0) return showError("無有效數據");
    
    let boxplotData = [];
    for (const [group, values] of Object.entries(grouped)) {
        if (values.length < 2) continue;
        const sorted = values.slice().sort((a, b) => a - b);
        const stats = calculateBoxplotStats(sorted);
        boxplotData.push({ label: group, values: values, ...stats });
    }
    
    if (boxplotData.length === 0) return showError("數據不足以繪製盒鬚圖");
    
    const showMean = document.getElementById('boxplot-show-mean').checked;
    const showOutliers = document.getElementById('boxplot-show-outliers').checked;
    const runANOVA = document.getElementById('boxplot-run-anova').checked;
    
    drawBoxplotChart(boxplotData, showMean, showOutliers);
    
    if (runANOVA) {
        const anovaResult = calculateANOVA(grouped);
        displayANOVA(anovaResult);
    } else {
        document.getElementById('anova-results').style.display = 'none';
    }
}

function calculateBoxplotStats(sortedData) {
    const n = sortedData.length;
    const min = sortedData[0];
    const max = sortedData[n - 1];
    const q1Index = Math.floor(n * 0.25);
    const q2Index = Math.floor(n * 0.5);
    const q3Index = Math.floor(n * 0.75);
    const q1 = sortedData[q1Index];
    const median = n % 2 === 0 ? (sortedData[q2Index - 1] + sortedData[q2Index]) / 2 : sortedData[q2Index];
    const q3 = sortedData[q3Index];
    const mean = sortedData.reduce((a, b) => a + b, 0) / n;
    const iqr = q3 - q1;
    const lowerFence = q1 - 1.5 * iqr;
    const upperFence = q3 + 1.5 * iqr;
    const outliers = sortedData.filter(v => v < lowerFence || v > upperFence);
    return { min, q1, median, q3, max, mean, outliers, iqr };
}

function drawBoxplotChart(boxplotData, showMean, showOutliers) {
    const ctx = document.getElementById('boxplot-chart').getContext('2d');
    if (boxplotChart) boxplotChart.destroy();
    
    const labels = boxplotData.map(d => d.label);
    const datasets = [{
        label: '盒鬚圖',
        backgroundColor: 'rgba(54, 162, 235, 0.5)',
        borderColor: 'rgb(54, 162, 235)',
        borderWidth: 2,
        outlierBackgroundColor: 'rgba(231, 76, 60, 0.8)',
        outlierRadius: 4,
        itemRadius: 0,
        data: boxplotData.map(d => {
            const boxData = { min: d.min, q1: d.q1, median: d.median, q3: d.q3, max: d.max };
            if (showOutliers && d.outliers.length > 0) boxData.outliers = d.outliers;
            if (showMean) boxData.mean = d.mean;
            return boxData;
        })
    }];
    
    boxplotChart = new Chart(ctx, {
        type: 'boxplot',
        data: { labels: labels, datasets: datasets },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: true },
                tooltip: {
                    callbacks: {
                        label: function (context) {
                            const data = context.parsed;
                            return [
                                `最小值: ${data.min.toFixed(3)}`,
                                `Q1: ${data.q1.toFixed(3)}`,
                                `中位數: ${data.median.toFixed(3)}`,
                                `Q3: ${data.q3.toFixed(3)}`,
                                `最大值: ${data.max.toFixed(3)}`,
                                showMean ? `平均值: ${data.mean.toFixed(3)}` : ''
                            ].filter(s => s);
                        }
                    }
                }
            },
            scales: { y: { title: { display: true, text: '數值' } } }
        }
    });
}

function calculateANOVA(groupedData) {
    const groups = Object.values(groupedData); 
    const groupNames = Object.keys(groupedData); 
    const k = groups.length;
    let N = 0; 
    groups.forEach(g => N += g.length);
    
    let grandSum = 0; 
    groups.forEach(g => { g.forEach(v => grandSum += v); }); 
    const grandMean = grandSum / N;
    
    let ssb = 0; 
    groups.forEach(g => { 
        const groupMean = g.reduce((a, b) => a + b, 0) / g.length; 
        ssb += g.length * Math.pow(groupMean - grandMean, 2); 
    });
    
    let ssw = 0; 
    groups.forEach(g => { 
        const groupMean = g.reduce((a, b) => a + b, 0) / g.length; 
        g.forEach(v => { ssw += Math.pow(v - groupMean, 2); }); 
    });
    
    const sst = ssb + ssw; 
    const dfb = k - 1; 
    const dfw = N - k; 
    const dft = N - 1;
    const msb = ssb / dfb; 
    const msw = ssw / dfw; 
    const f = msb / msw;
    
    let pValue = "< 0.05"; 
    if (f < 3.0) pValue = "> 0.05"; 
    else if (f < 5.0) pValue = "< 0.05"; 
    else if (f < 10.0) pValue = "< 0.01"; 
    else pValue = "< 0.001";
    
    return { ssb, ssw, sst, dfb, dfw, dft, msb, msw, f, pValue, groupNames };
}

function displayANOVA(result) {
    const container = document.getElementById('anova-results'); 
    const tableDiv = document.getElementById('anova-table');
    let html = `<table style="width: 100%; border-collapse: collapse;">
    <thead><tr style="background: #3498db; color: white;">
    <th style="padding: 10px; border: 1px solid #ddd;">變異來源</th>
    <th style="padding: 10px; border: 1px solid #ddd;">平方和 (SS)</th>
    <th style="padding: 10px; border: 1px solid #ddd;">自由度 (df)</th>
    <th style="padding: 10px; border: 1px solid #ddd;">均方 (MS)</th>
    <th style="padding: 10px; border: 1px solid #ddd;">F 值</th>
    <th style="padding: 10px; border: 1px solid #ddd;">p 值</th>
    </tr></thead><tbody>
    <tr><td style="padding: 8px; border: 1px solid #ddd;">組間 (Between)</td><td style="padding: 8px; border: 1px solid #ddd;">${result.ssb.toFixed(4)}</td><td style="padding: 8px; border: 1px solid #ddd;">${result.dfb}</td><td style="padding: 8px; border: 1px solid #ddd;">${result.msb.toFixed(4)}</td><td style="padding: 8px; border: 1px solid #ddd;" rowspan="2">${result.f.toFixed(4)}</td><td style="padding: 8px; border: 1px solid #ddd;" rowspan="2">${result.pValue}</td></tr>
    <tr><td style="padding: 8px; border: 1px solid #ddd;">組內 (Within)</td><td style="padding: 8px; border: 1px solid #ddd;">${result.ssw.toFixed(4)}</td><td style="padding: 8px; border: 1px solid #ddd;">${result.dfw}</td><td style="padding: 8px; border: 1px solid #ddd;">${result.msw.toFixed(4)}</td></tr>
    <tr style="background: #ecf0f1;"><td style="padding: 8px; border: 1px solid #ddd;"><strong>總計 (Total)</strong></td><td style="padding: 8px; border: 1px solid #ddd;"><strong>${result.sst.toFixed(4)}</strong></td><td style="padding: 8px; border: 1px solid #ddd;"><strong>${result.dft}</strong></td><td style="padding: 8px; border: 1px solid #ddd;">-</td><td style="padding: 8px; border: 1px solid #ddd;">-</td><td style="padding: 8px; border: 1px solid #ddd;">-</td></tr>
    </tbody></table>
    <div style="margin-top: 15px; padding: 10px; background: #e8f4f8; border-left: 4px solid #3498db;"><strong>結論:</strong> ${result.pValue.includes('<') && !result.pValue.includes('> 0.05') ? '組間差異<strong>顯著</strong> (拒絕虛無假設)' : '組間差異<strong>不顯著</strong> (接受虛無假設)'}</div>`;
    tableDiv.innerHTML = html; 
    container.style.display = 'block';
}
