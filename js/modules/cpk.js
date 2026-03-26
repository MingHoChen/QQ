import { SharedData, SharedSpecs, runWorkerTask, showError, showSuccess } from '../utils.js';
// We need to attach drawCpkChart to window. Since utils.js doesn't export window, we use global window.
// Wait, we can just use `window.drawCpkChart`.

let loadedWorkbooksLocal = []; 
export let cpk_myChart = null;
export let cpkResults = [];

export function initCpkEvents() {
    const fileInput = document.getElementById('file-input');
    if (fileInput) fileInput.addEventListener('change', handleCpkFileSelect);
    
    const calcBtn = document.getElementById('calculate-btn');
    if (calcBtn) calcBtn.addEventListener('click', performCpkCalculation);
    
    const modeSelect = document.getElementById('data-mode-select');
    if (modeSelect) modeSelect.addEventListener('change', (e) => toggleSetupSections(e.target.value));
    
    const confirmBtn = document.getElementById('confirm-sheet-btn');
    if (confirmBtn) confirmBtn.addEventListener('click', processSelectedSheets);

    const exportBtn = document.getElementById('export-btn');
    if (exportBtn) exportBtn.addEventListener('click', handleCpkExport);

    const lsl = document.getElementById('lsl');
    if (lsl) lsl.addEventListener('input', (e) => SharedSpecs.lsl = parseFloat(e.target.value));
    
    const usl = document.getElementById('usl');
    if (usl) usl.addEventListener('input', (e) => SharedSpecs.usl = parseFloat(e.target.value));
}

async function handleCpkFileSelect(e) {
    const files = e.target.files;
    if (!files.length) return;
    
    const statusEl = document.getElementById('file-status');
    statusEl.innerHTML = '<span style="color:#f39c12">⏳ 解析檔案中，請稍候...</span>';
    loadedWorkbooksLocal = [];
    
    try {
        const filePayloads = await Promise.all(Array.from(files).map(async (f) => {
            const buffer = await f.arrayBuffer();
            return { name: f.name, buffer };
        }));
        
        const res = await runWorkerTask('readWorkbooks', { files: filePayloads });
        loadedWorkbooksLocal = res.results;
        
        statusEl.innerHTML = `<span style="color:#27ae60">✅ 已載入 ${loadedWorkbooksLocal.length} 個檔案。請於下方勾選工作表並讀取。</span>`;
        renderSheetSelector();
    } catch (err) {
        showError("檔案讀取失敗: " + err.message);
        statusEl.innerHTML = '<span style="color:#e74c3c">❌ 讀取失敗</span>';
    }
}

function renderSheetSelector() {
    const sheetSel = document.getElementById('sheet-selector');
    sheetSel.innerHTML = '';
    sheetSel.disabled = false;

    loadedWorkbooksLocal.forEach((wbData) => {
        wbData.sheetNames.forEach(sheetName => {
            const opt = new Option(`[${wbData.name}] ${sheetName}`, `${wbData.index}|${sheetName}`);
            sheetSel.add(opt);
        });
    });

    if (sheetSel.options.length === 1) {
        sheetSel.selectedIndex = 0;
    }
}

async function processSelectedSheets() {
    const sheetSel = document.getElementById('sheet-selector');
    const selectedOptions = Array.from(sheetSel.selectedOptions);
    
    if (selectedOptions.length === 0) return showError("請至少選擇一個工作表！");

    const statusEl = document.getElementById('file-status');
    statusEl.innerHTML = '<span style="color:#f39c12">⏳ 正在處理與抽取數據...</span>';
    
    const selections = selectedOptions.map(option => {
        const parts = option.value.split('|');
        return { fIdx: parseInt(parts[0]), sheetName: parts[1] };
    });
    
    try {
        const res = await runWorkerTask('processSheets', { selections });
        
        SharedData.rawJson = res.data;
        SharedData.headers = res.headers;
        
        if (res.successCount > 0 && SharedData.rawJson.length > 0) {
            SharedData.hasData = true;
            populateCpkSelectors();
            const mode = document.getElementById('data-mode-select').value;
            toggleSetupSections(mode);
            
            statusEl.innerHTML = `<span style="color:#27ae60">✅ 成功讀取 ${res.successCount} 個工作表，共 ${SharedData.rawJson.length} 筆數據。</span>`;
        } else {
            statusEl.innerHTML = '<span style="color:#e74c3c">❌ 選定的工作表無有效數據或標頭。</span>';
            showError("選定的工作表無有效數據。");
        }
    } catch (err) {
        showError("處理數據失敗: " + err.message);
        statusEl.innerHTML = '';
    }
}

function toggleSetupSections(mode) {
    document.getElementById('col-select-long').style.display = mode === 'long' ? 'block' : 'none';
    document.getElementById('col-select-wide').style.display = mode === 'wide' ? 'block' : 'none';
}

export function populateCpkSelectors() {
    const catSel = document.getElementById('category-col');
    const valSel = document.getElementById('value-col');
    const wideSel = document.getElementById('category-header-selector');

    [catSel, valSel, wideSel].forEach(s => s.innerHTML = '');
    
    const optAll = new Option("-- 不分組 (全部數據視為一組) --", "_ALL_");
    optAll.style.fontWeight = "bold";
    catSel.add(optAll);

    SharedData.headers.forEach(h => {
        catSel.add(new Option(h, h));
        valSel.add(new Option(h, h));
        wideSel.add(new Option(h, h));
    });

    catSel.selectedIndex = 0;
    if (valSel.options.length > 1) valSel.selectedIndex = 1;
}

function performCpkCalculation() {
    if (!SharedData.hasData) return showError("請先上傳並讀取數據");
    
    const lslStr = document.getElementById('lsl').value;
    const uslStr = document.getElementById('usl').value;
    const lsl = lslStr === "" ? NaN : parseFloat(lslStr);
    const usl = uslStr === "" ? NaN : parseFloat(uslStr);
    
    const mode = document.getElementById('data-mode-select').value;
    let grouped = {};
    let totalCount = 0;

    if (mode === 'long') {
        const catSelect = document.getElementById('category-col');
        const valKey = document.getElementById('value-col').value;
        const selectedCats = Array.from(catSelect.selectedOptions).map(opt => opt.value);
        
        if (selectedCats.includes('_ALL_') || selectedCats.length === 0) {
            grouped["Total_Data"] = [];
            SharedData.rawJson.forEach(row => {
                let val = parseFloat(String(row[valKey]).replace(/[^0-9.\-]/g, ''));
                if (!isNaN(val)) { grouped["Total_Data"].push(val); totalCount++; }
            });
        } else {
            SharedData.rawJson.forEach(row => {
                const compositeKey = selectedCats.map(k => row[k]).join('_');
                let val = parseFloat(String(row[valKey]).replace(/[^0-9.\-]/g, ''));
                if (compositeKey && !isNaN(val)) {
                    if (!grouped[compositeKey]) grouped[compositeKey] = [];
                    grouped[compositeKey].push(val);
                    totalCount++;
                }
            });
        }
    } else {
        const nameKey = document.getElementById('category-header-selector').value;
        SharedData.rawJson.forEach(row => {
            let rawName = row[nameKey];
            if (rawName === undefined || rawName === null || String(rawName).trim() === "") return;
            let catName = String(rawName).trim();
            let values = [];
            Object.keys(row).forEach(key => {
                if (key === nameKey) return;
                let val = parseFloat(String(row[key]).replace(/[^0-9.\-]/g, ''));
                if (!isNaN(val)) values.push(val);
            });
            if (values.length > 0) {
                if (!grouped[catName]) grouped[catName] = [];
                grouped[catName] = grouped[catName].concat(values);
                totalCount += values.length;
            }
        });
    }

    if (totalCount === 0) return showError("找不到有效數值，請檢查欄位選擇是否正確");

    cpkResults = [];
    const sortedKeys = Object.keys(grouped).sort((a, b) => {
        const numA = parseFloat(a);
        const numB = parseFloat(b);
        if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
        return a.localeCompare(b);
    });

    sortedKeys.forEach(cat => {
        const vals = grouped[cat];
        const stats = calculateStats(vals, lsl, usl);
        if (!stats.error) {
            cpkResults.push({ category: cat, values: vals, stats, lsl, usl });
        }
    });

    if (cpkResults.length === 0) return showError("無法計算！樣本數不足。");
    displayCpkTable();
}

export function calculateStats(data, lsl, usl) {
    if (!data || data.length < 2) return { error: true };
    const mean = data.reduce((a, b) => a + b, 0) / data.length;
    const variance = data.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / (data.length - 1);
    const stdDev = Math.sqrt(variance);
    if (stdDev === 0) return { mean, stdDev, lcl: mean, ucl: mean, cpk: NaN, error: false };
    const lcl = mean - 3 * stdDev; const ucl = mean + 3 * stdDev;
    let cpk = NaN;
    if (!isNaN(lsl) && !isNaN(usl)) { const cpu = (usl - mean) / (3 * stdDev); const cpl = (mean - lsl) / (3 * stdDev); cpk = Math.min(cpu, cpl); }
    else if (!isNaN(usl)) { cpk = (usl - mean) / (3 * stdDev); } else if (!isNaN(lsl)) { cpk = (mean - lsl) / (3 * stdDev); }
    return { mean, stdDev, lcl, ucl, cpk, error: false };
}

function displayCpkTable() {
    const div = document.getElementById('results-display');
    let html = `<table><thead><tr><th>項目</th><th>N</th><th>Mean</th><th>StdDev</th><th>LCL</th><th>UCL</th><th>Ppk (Overall)</th><th>狀態</th></tr></thead><tbody>`;
    cpkResults.forEach((item, i) => {
        const s = item.stats; let statusText = '-'; let statusClass = '';
        if (!isNaN(s.cpk)) { if (s.cpk >= 1.33) { statusText = '良好'; statusClass = 'status-good'; } else if (s.cpk >= 1.0) { statusText = '尚可'; statusClass = 'status-ok'; } else { statusText = '不足'; statusClass = 'status-bad'; } }
        html += `<tr onclick="window.drawCpkChart(${i})" style="cursor:pointer"><td>${item.category}</td><td>${item.values.length}</td><td>${s.mean.toFixed(3)}</td><td>${s.stdDev.toFixed(3)}</td><td>${s.lcl.toFixed(3)}</td><td>${s.ucl.toFixed(3)}</td><td>${isNaN(s.cpk) ? '-' : s.cpk.toFixed(3)}</td><td class="${statusClass}">${statusText}</td></tr>`;
    });
    html += `</tbody></table>`;
    div.innerHTML = html;
    document.getElementById('results-card').style.display = 'block';
    document.getElementById('export-btn').style.display = 'block';
    if (cpkResults.length > 0) window.drawCpkChart(0);
}

window.drawCpkChart = function(index) {
    document.getElementById('chart-card').style.display = 'block';
    const item = cpkResults[index];
    const ctx = document.getElementById('cpk-chart').getContext('2d');
    if (cpk_myChart) cpk_myChart.destroy();

    const vals = item.values;
    const min = Math.min(...vals);
    const max = Math.max(...vals);
    let plotMin = min, plotMax = max;

    if (!isNaN(item.lsl)) plotMin = Math.min(plotMin, item.lsl);
    if (!isNaN(item.usl)) plotMax = Math.max(plotMax, item.usl);
    plotMin = Math.min(plotMin, item.stats.lcl);
    plotMax = Math.max(plotMax, item.stats.ucl);
    if (plotMin === plotMax) { plotMin -= 0.5; plotMax += 0.5; }

    const bins = 20;
    const range = plotMax - plotMin;
    const step = range / bins;
    const labels = [], counts = new Array(bins).fill(0);

    for (let i = 0; i < bins; i++) labels.push((plotMin + i * step).toFixed(2));
    
    vals.forEach(v => {
        let idx = Math.floor((v - plotMin) / step);
        if (idx < 0) idx = 0;
        if (idx >= bins) idx = bins - 1;
        counts[idx]++;
    });

    const valToX = (v) => (v - plotMin) / step - 0.5;
    const annotations = {};

    if (!isNaN(item.lsl)) annotations.lsl = { type: 'line', xMin: valToX(item.lsl), xMax: valToX(item.lsl), borderColor: 'red', borderWidth: 3, label: { display: true, content: 'LSL', position: 'start', backgroundColor: 'red', color: 'white' } };
    if (!isNaN(item.usl)) annotations.usl = { type: 'line', xMin: valToX(item.usl), xMax: valToX(item.usl), borderColor: 'red', borderWidth: 3, label: { display: true, content: 'USL', position: 'start', backgroundColor: 'red', color: 'white' } };
    annotations.lcl = { type: 'line', xMin: valToX(item.stats.lcl), xMax: valToX(item.stats.lcl), borderColor: 'blue', borderDash: [6, 6], borderWidth: 2, label: { display: true, content: 'LCL', position: 'end', backgroundColor: 'blue', color: 'white' } };
    annotations.ucl = { type: 'line', xMin: valToX(item.stats.ucl), xMax: valToX(item.stats.ucl), borderColor: 'blue', borderDash: [6, 6], borderWidth: 2, label: { display: true, content: 'UCL', position: 'end', backgroundColor: 'blue', color: 'white' } };
    annotations.mean = { type: 'line', xMin: valToX(item.stats.mean), xMax: valToX(item.stats.mean), borderColor: 'green', borderWidth: 2, label: { display: true, content: 'Mean', position: 'center', backgroundColor: 'green', color: 'white' } };

    const curveData = [];
    const N = vals.length;
    const mean = item.stats.mean;
    const stdDev = item.stats.stdDev;
    for (let i = 0; i < bins; i++) {
        const xVal = plotMin + i * step + (step / 2);
        let expectedCount = 0;
        if (stdDev > 0) {
            const pdf = (1 / (stdDev * Math.sqrt(2 * Math.PI))) * Math.exp(-0.5 * Math.pow((xVal - mean) / stdDev, 2));
            expectedCount = pdf * N * step;
        }
        curveData.push(expectedCount);
    }

    cpk_myChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: [
                {
                    label: '趨勢線 (常態分佈)',
                    data: curveData,
                    type: 'line',
                    borderColor: 'rgba(255, 99, 132, 1)',
                    borderWidth: 2,
                    pointRadius: 0,
                    pointHoverRadius: 0,
                    tension: 0.4,
                    fill: false
                },
                { label: 'Count', data: counts, backgroundColor: 'rgba(54, 162, 235, 0.5)' }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { annotation: { annotations } }
        }
    });
    
    const rows = document.querySelectorAll('#results-display tr');
    rows.forEach(r => r.classList.remove('selected-row'));
    if (rows[index + 1]) rows[index + 1].classList.add('selected-row');
};

async function handleCpkExport() {
    if (cpkResults.length === 0) return showError("無數據可匯出");
    if (typeof ExcelJS === 'undefined') return showError("找不到 ExcelJS，請確認網路連線。");
    
    // Disable button to prevent multi-click
    const exportBtn = document.getElementById('export-btn');
    exportBtn.disabled = true;
    exportBtn.innerText = "正在匯出...";

    try {
        const wb = new ExcelJS.Workbook(); 
        const ws = wb.addWorksheet('Ppk Analysis');
        ws.columns = [
            { header: '項目', key: 'cat', width: 20 }, 
            { header: 'N', key: 'n', width: 10 }, 
            { header: 'Mean', key: 'mean', width: 12 }, 
            { header: 'StdDev', key: 'std', width: 12 }, 
            { header: 'LCL', key: 'lcl', width: 12 }, 
            { header: 'UCL', key: 'ucl', width: 12 }, 
            { header: 'Ppk', key: 'cpk', width: 14 }, 
            { header: 'Status', key: 'status', width: 10 }, 
            { header: 'LSL', key: 'lsl', width: 10 }, 
            { header: 'USL', key: 'usl', width: 10 }
        ];
        
        cpkResults.forEach(item => { 
            const s = item.stats; 
            ws.addRow({ 
                cat: item.category, n: item.values.length, 
                mean: s.mean, std: s.stdDev, lcl: s.lcl, 
                ucl: s.ucl, cpk: s.cpk, lsl: item.lsl, usl: item.usl,
                status: (isNaN(s.cpk) ? '-' : (s.cpk >= 1.33 ? '良好' : (s.cpk >= 1 ? '尚可' : '不足')))
            }); 
        });
        
        const canvas = document.getElementById('cpk-chart');
        if (canvas) {
            const chartImageId = wb.addImage({
                base64: canvas.toDataURL('image/png'),
                extension: 'png',
            });
            ws.addImage(chartImageId, {
                tl: { col: 11, row: 1 },
                ext: { width: 500, height: 300 }
            });
        }
        
        const buffer = await wb.xlsx.writeBuffer(); 
        saveAs(new Blob([buffer]), "Ppk_Report.xlsx");
        showSuccess("匯出成功！");
    } catch (err) {
        showError("匯出失敗: " + err.message);
    } finally {
        exportBtn.disabled = false;
        exportBtn.innerText = "匯出 Excel (含圖表)";
    }
}
