/**
 * QUALITY ENGINEERING INTEGRATED SYSTEM (v19.19 Add Normal Curve)
 * Features:
 * - Cpk/Ppk Calculation (Fixed Wide Mode: Row-based Analysis)
 * - SPC Control Charts
 * - Pareto Analysis
 * - Scatter Plot Analysis
 * - Box Plot Analysis
 * - Gage R&R Analysis
 */

if (typeof Chart !== 'undefined' && typeof chartjsPluginAnnotation !== 'undefined') {
    Chart.register(chartjsPluginAnnotation);
}

const SharedData = { rawJson: [], headers: [], hasData: false, tempPareto: null, tempScatter: null, tempBoxplot: null, tempGrr: null };
const SharedSpecs = { lsl: NaN, usl: NaN };
let CONFIG = { numSubgroups: 50, maxRows: 200, sampleSize: 5, decimalPlaces: 3 };
const CONSTANTS = { n: 5, A2: 0.577, D3: 0, D4: 2.114, d2: 2.326 };

let cpk_myChart = null;
let cpkResults = [];
let spc_charts = { xbar: null, range: null };
let grr_charts = {};
let spcGroupedData = {};
let loadedWorkbooks = [];

window.onload = function () {
    // Cpk Events
    document.getElementById('file-input').addEventListener('change', handleCpkFileSelect);
    document.getElementById('calculate-btn').addEventListener('click', performCpkCalculation);
    const modeSelect = document.getElementById('data-mode-select');
    if (modeSelect) {
        modeSelect.addEventListener('change', (e) => toggleSetupSections(e.target.value));
    }
    const exportBtn = document.getElementById('export-btn');
    if (exportBtn) exportBtn.addEventListener('click', handleCpkExport);

    document.getElementById('lsl').addEventListener('input', (e) => SharedSpecs.lsl = parseFloat(e.target.value));
    document.getElementById('usl').addEventListener('input', (e) => SharedSpecs.usl = parseFloat(e.target.value));

    // SPC Events
    initSpcTable();
    initSpcResizer();
    document.getElementById('spc-single-file').addEventListener('change', handleSpcSingleFile);
    document.getElementById('uslInput').addEventListener('change', calculateAndDraw);
    document.getElementById('lslInput').addEventListener('change', calculateAndDraw);
    handleSpcSourceChange();

    // Pareto Events
    document.getElementById('pareto-file-input').addEventListener('change', handleParetoFileSelect);

    // Scatter Events
    document.getElementById('scatter-file-input').addEventListener('change', handleScatterFileSelect);

    // Boxplot Events
    document.getElementById('boxplot-file-input').addEventListener('change', handleBoxplotFileSelect);

    // Gage R&R Events
    document.getElementById('grr-file-input').addEventListener('change', handleGrrFileSelect);
};

function switchTab(tabId) {
    document.querySelectorAll('.nav-tab').forEach(btn => btn.classList.remove('active'));
    event.currentTarget.classList.add('active');
    document.querySelectorAll('.tab-content').forEach(div => div.classList.remove('active'));
    document.getElementById('view-' + tabId).classList.add('active');

    if (tabId === 'spc') {
        if (!isNaN(SharedSpecs.lsl)) document.getElementById('lslInput').value = SharedSpecs.lsl;
        if (!isNaN(SharedSpecs.usl)) document.getElementById('uslInput').value = SharedSpecs.usl;
        handleSpcSourceChange();
        setTimeout(() => {
            if (spc_charts.xbar) spc_charts.xbar.resize();
            if (spc_charts.range) spc_charts.range.resize();
            calculateAndDraw();
        }, 100);
    } else if (tabId === 'pareto') {
        const mode = document.getElementById('pareto-source-mode').value;
        if (mode === 'cpk' && SharedData.hasData) populateParetoSelectors(SharedData.headers);
    } else if (tabId === 'scatter') {
        const mode = document.getElementById('scatter-source-mode').value;
        if (mode === 'cpk' && SharedData.hasData) populateScatterSelectors(SharedData.headers);
    } else if (tabId === 'boxplot') {
        const mode = document.getElementById('boxplot-source-mode').value;
        if (mode === 'cpk' && SharedData.hasData) populateBoxplotSelectors(SharedData.headers);
    } else if (tabId === 'grr') {
        const mode = document.getElementById('grr-source-mode').value;
        if (mode === 'cpk' && SharedData.hasData) populateGrrSelectors(SharedData.headers);
    }
}

// --- Common Utilities ---
function robustParseExcel(sheet) {
    const rawArray = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    if (!rawArray || rawArray.length === 0) return { headers: [], data: [] };
    const firstRow = rawArray[0];
    let isHeaderRow = true;
    let numericCount = 0;
    if (firstRow && firstRow.length > 0) {
        firstRow.forEach(cell => {
            if (typeof cell === 'number') numericCount++;
            else if (typeof cell === 'string' && !isNaN(parseFloat(cell))) numericCount++;
        });
        if (numericCount > firstRow.length / 2) isHeaderRow = false;
    }
    let safeHeaders = [];
    let dataStartIndex = 0;
    if (isHeaderRow) {
        let originalHeaders = rawArray[0] || [];
        dataStartIndex = 1;
        for (let i = 0; i < originalHeaders.length; i++) {
            let h = originalHeaders[i];
            let name = (h === undefined || h === null || String(h).trim() === "") ? `Col_${i + 1}` : String(h).trim();
            let baseName = name; let count = 1;
            while (safeHeaders.includes(name)) { count++; name = `${baseName}_${count}`; }
            safeHeaders.push(name);
        }
    } else {
        let maxCols = 0;
        rawArray.forEach(r => { if (r) maxCols = Math.max(maxCols, r.length); });
        for (let i = 0; i < maxCols; i++) safeHeaders.push(`Column_${i + 1}`);
        dataStartIndex = 0;
    }
    let data = [];
    for (let i = dataStartIndex; i < rawArray.length; i++) {
        let row = rawArray[i];
        if (!row || row.length === 0) continue;
        let obj = {}; let hasVal = false;
        safeHeaders.forEach((header, idx) => {
            let val = row[idx];
            obj[header] = val;
            if (val !== undefined && val !== null && val !== "") hasVal = true;
        });
        if (hasVal) data.push(obj);
    }
    return { headers: safeHeaders, data: data };
}

// --- Cpk Logic ---
function handleCpkFileSelect(e) {
    const files = e.target.files;
    if (!files.length) return;
    loadedWorkbooks = [];
    document.getElementById('file-status').innerText = "讀取中...";
    
    const promises = Array.from(files).map((file) => {
        return new Promise((resolve) => {
            const reader = new FileReader();
            reader.onload = (evt) => {
                try {
                    const wb = XLSX.read(new Uint8Array(evt.target.result), { type: 'array' });
                    resolve({ name: file.name, wb: wb });
                } catch (e) { console.error(e); resolve(null); }
            };
            reader.readAsArrayBuffer(file);
        });
    });

    Promise.all(promises).then(results => {
        loadedWorkbooks = results.filter(r => r !== null);
        document.getElementById('file-status').innerText = `已載入 ${loadedWorkbooks.length} 個檔案。`;
        renderSheetSelector();
    });
}

function renderSheetSelector() {
    const sheetSel = document.getElementById('sheet-selector');
    const btn = document.getElementById('confirm-sheet-btn');
    sheetSel.innerHTML = '';
    sheetSel.disabled = false;

    loadedWorkbooks.forEach((item, fIdx) => {
        item.wb.SheetNames.forEach(sheetName => {
            const opt = new Option(`[${item.name}] ${sheetName}`, `${fIdx}|${sheetName}`);
            sheetSel.add(opt);
        });
    });

    if (sheetSel.options.length === 1) {
        sheetSel.selectedIndex = 0;
    }
}

function processSelectedSheets() {
    const sheetSel = document.getElementById('sheet-selector');
    const selectedOptions = Array.from(sheetSel.selectedOptions);
    
    if (selectedOptions.length === 0) return alert("請至少選擇一個工作表！");

    SharedData.rawJson = [];
    SharedData.headers = [];
    let successCount = 0;
    
    selectedOptions.forEach(option => {
        const val = option.value;
        const parts = val.split('|');
        const fIdx = parseInt(parts[0]);
        const sheetName = parts[1];
        
        const wb = loadedWorkbooks[fIdx].wb;
        const sheet = wb.Sheets[sheetName];
        const result = robustParseExcel(sheet);

        if (result.headers.length > 0) {
            if (SharedData.headers.length === 0) {
                SharedData.headers = result.headers;
            }
            SharedData.rawJson = SharedData.rawJson.concat(result.data);
            successCount++;
        }
    });

    if (successCount > 0 && SharedData.rawJson.length > 0) {
        SharedData.hasData = true;
        populateCpkSelectors();
        const mode = document.getElementById('data-mode-select').value;
        toggleSetupSections(mode);
        
        document.getElementById('file-status').innerText = `成功讀取 ${successCount} 個工作表，共 ${SharedData.rawJson.length} 筆數據 ✅`;
        document.getElementById('file-status').style.color = "green";
    } else {
        document.getElementById('file-status').innerText = "❌ 讀取失敗：選定的工作表無有效數據或標頭。";
        document.getElementById('file-status').style.color = "red";
    }
}

function toggleSetupSections(mode) {
    document.getElementById('col-select-long').style.display = mode === 'long' ? 'block' : 'none';
    document.getElementById('col-select-wide').style.display = mode === 'wide' ? 'block' : 'none';
}

function populateCpkSelectors() {
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
    if (!SharedData.hasData) return alert("請先上傳並讀取數據");
    
    const lslStr = document.getElementById('lsl').value;
    const uslStr = document.getElementById('usl').value;
    const lsl = lslStr === "" ? NaN : parseFloat(lslStr);
    const usl = uslStr === "" ? NaN : parseFloat(uslStr);
    
    const mode = document.getElementById('data-mode-select').value;
    let grouped = {};
    let totalCount = 0;

    if (mode === 'long') {
        // 模式 A: 長格式 (每列是一筆記錄，欄位決定分組)
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
        // --- 模式 B: 寬格式 (Row-based) 更新 ---
        // 使用者選擇的欄位是「標頭名稱」(e.g. "LO(MHz)")
        // 我們要以該欄位的內容作為 "Category Name"
        // 並將該列的其他所有數值欄位作為 "Values"
        
        const nameKey = document.getElementById('category-header-selector').value;
        
        SharedData.rawJson.forEach(row => {
            // 1. 取得該列的名稱 (e.g. "6000", "6500")
            let rawName = row[nameKey];
            if (rawName === undefined || rawName === null || String(rawName).trim() === "") return;
            let catName = String(rawName).trim();

            // 2. 收集該列的所有數值 (排除 nameKey 欄位)
            let values = [];
            Object.keys(row).forEach(key => {
                if (key === nameKey) return; // 跳過標題欄
                
                // 嘗試解析數值
                let val = parseFloat(String(row[key]).replace(/[^0-9.\-]/g, ''));
                if (!isNaN(val)) {
                    values.push(val);
                }
            });

            if (values.length > 0) {
                // 如果已經有同名的 category (雖然通常不會)，合併數據
                if (!grouped[catName]) grouped[catName] = [];
                grouped[catName] = grouped[catName].concat(values);
                totalCount += values.length;
            }
        });
    }

    if (totalCount === 0) return alert("找不到有效數值，請檢查欄位選擇是否正確");

    // --- 計算邏輯 ---
    cpkResults = [];
    // 對 Group Name 進行排序 (如果是數字，按數字排序；否則按字母)
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

    if (cpkResults.length === 0) return alert(`無法計算！樣本數不足。`);
    displayCpkTable();
}

function calculateStats(data, lsl, usl) {
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
        html += `<tr onclick="drawCpkChart(${i})" style="cursor:pointer"><td>${item.category}</td><td>${item.values.length}</td><td>${s.mean.toFixed(3)}</td><td>${s.stdDev.toFixed(3)}</td><td>${s.lcl.toFixed(3)}</td><td>${s.ucl.toFixed(3)}</td><td>${isNaN(s.cpk) ? '-' : s.cpk.toFixed(3)}</td><td class="${statusClass}">${statusText}</td></tr>`;
    });
    html += `</tbody></table>`;
    div.innerHTML = html;
    document.getElementById('results-card').style.display = 'block';
    document.getElementById('export-btn').style.display = 'block';
    if (cpkResults.length > 0) drawCpkChart(0);
}
function drawCpkChart(index) {
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

    // 計算常態分佈趨勢線 (Normal Distribution Curve)
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
}

async function handleCpkExport() {
    if (cpkResults.length === 0) return alert("無數據可匯出");
    const wb = new ExcelJS.Workbook(); const ws = wb.addWorksheet('Ppk Analysis');
    ws.columns = [{ header: '項目', key: 'cat', width: 20 }, { header: 'N', key: 'n', width: 10 }, { header: 'Mean', key: 'mean', width: 12 }, { header: 'StdDev', key: 'std', width: 12 }, { header: 'LCL', key: 'lcl', width: 12 }, { header: 'UCL', key: 'ucl', width: 12 }, { header: 'Ppk', key: 'cpk', width: 14 }, { header: 'Status', key: 'status', width: 10 }, { header: 'LSL', key: 'lsl', width: 10 }, { header: 'USL', key: 'usl', width: 10 }];
    cpkResults.forEach(item => { const s = item.stats; ws.addRow({ cat: item.category, n: item.values.length, mean: s.mean, std: s.stdDev, lcl: s.lcl, ucl: s.ucl, cpk: s.cpk, lsl: item.lsl, usl: item.usl }); });
    const buffer = await wb.xlsx.writeBuffer(); saveAs(new Blob([buffer]), "Ppk_Report.xlsx");
}

// --- SPC Logic ---
function handleSpcSourceChange() {
    const mode = document.getElementById('spc-source-mode').value;
    document.getElementById('spc-tool-cpk').style.display = mode === 'cpk' ? 'flex' : 'none';
    document.getElementById('spc-tool-new').style.display = mode === 'new' ? 'flex' : 'none';
    const itemSel = document.getElementById('spc-inherit-item-select');
    itemSel.style.display = 'none'; itemSel.innerHTML = '<option value="">(請先讀取)</option>';
}
function parseSpcInheritedData() {
    if (!SharedData.hasData) return alert("請先在 Cpk 頁面上傳數據");
    const catKey = document.getElementById('category-col').value; const valKey = document.getElementById('value-col').value;
    processSpcGrouping(SharedData.rawJson, catKey, valKey);
}
function handleSpcSingleFile(e) {
    const file = e.target.files[0]; if (!file) return; clearSpcData(false);
    const reader = new FileReader(); reader.onload = (evt) => {
        try {
            const wb = XLSX.read(new Uint8Array(evt.target.result), { type: 'array' });
            const sheet = wb.Sheets[wb.SheetNames[0]];
            const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
            if (!rawData || rawData.length === 0) return alert("讀取失敗：無數據");
            let maxNums = -1; let targetColIdx = 0;
            if (rawData[0].length > 1) { /* simple heuristic to find numeric col */ targetColIdx = 0; }
            let allValues = [];
            rawData.forEach(row => { if (row && row[targetColIdx] !== undefined) { let val = parseFloat(row[targetColIdx]); if (!isNaN(val)) allValues.push(val); } });
            if (allValues.length === 0) return alert("未偵測到有效數值");
            fillSpcTableFromValues(allValues); calculateAndDraw();
        } catch (err) { console.error(err); alert("讀取錯誤"); }
    };
    reader.readAsArrayBuffer(file);
}
function processSpcGrouping(dataArray, catKey, valKey) {
    spcGroupedData = {};
    if (!catKey || catKey === '_ALL_') {
        spcGroupedData["Total_Data"] = []; dataArray.forEach(row => { let val = parseFloat(String(row[valKey]).replace(/[^0-9.\-]/g, '')); if (!isNaN(val)) spcGroupedData["Total_Data"].push(val); });
    } else {
        dataArray.forEach(row => { const cat = row[catKey]; let val = parseFloat(String(row[valKey]).replace(/[^0-9.\-]/g, '')); if (cat !== undefined && !isNaN(val)) { if (!spcGroupedData[cat]) spcGroupedData[cat] = []; spcGroupedData[cat].push(val); } });
    }
    const itemSel = document.getElementById('spc-inherit-item-select');
    itemSel.innerHTML = '<option value="">-- 選擇項目 --</option>';
    Object.keys(spcGroupedData).forEach(k => itemSel.add(new Option(k, k)));
    itemSel.style.display = 'inline-block';
    if (itemSel.options.length > 1) { itemSel.selectedIndex = 1; loadSpcItemToTable(); }
}
function loadSpcItemToTable() { const key = document.getElementById('spc-inherit-item-select').value; if (!key) return; fillSpcTableFromValues(spcGroupedData[key]); calculateAndDraw(); }
function fillSpcTableFromValues(values) {
    let chunked = []; for (let i = 0; i < values.length; i += 5) chunked.push(values.slice(i, i + 5));
    if (chunked.length > CONFIG.numSubgroups) { CONFIG.numSubgroups = Math.min(chunked.length, CONFIG.maxRows); initSpcTable(); } else { for (let i = 1; i <= CONFIG.numSubgroups; i++) for (let j = 1; j <= 5; j++) { const el = document.getElementById(`cell_${i}_${j}`); if (el) el.value = ""; } }
    chunked.forEach((grp, rIdx) => { if (rIdx < CONFIG.numSubgroups) grp.forEach((v, cIdx) => { const el = document.getElementById(`cell_${rIdx + 1}_${cIdx + 1}`); if (el) el.value = v; }); });
}
function calculateAndDraw() {
    try {
        let subgroups = [], allData = [], sumX = 0, sumR = 0, valid = 0;
        for (let i = 1; i <= CONFIG.numSubgroups; i++) {
            let vals = []; for (let j = 1; j <= 5; j++) { const el = document.getElementById(`cell_${i}_${j}`); if (el && el.value !== "") vals.push(parseFloat(el.value)); }
            let xEl = document.getElementById(`res_xbar_${i}`); let rEl = document.getElementById(`res_r_${i}`);
            if (!xEl) continue;
            if (vals.length === 5) {
                let x = vals.reduce((a, b) => a + b, 0) / 5; let r = Math.max(...vals) - Math.min(...vals);
                xEl.innerText = x.toFixed(3); rEl.innerText = r.toFixed(3);
                subgroups.push({ id: i, xbar: x, r: r }); allData.push(...vals); sumX += x; sumR += r; valid++;
            } else { xEl.innerText = "-"; rEl.innerText = "-"; subgroups.push({ id: i, xbar: null, r: null }); }
        }
        if (valid < 2) { spcClearUI(); return; }
        const xdb = sumX / valid; const rb = sumR / valid;
        const uclX = xdb + (CONSTANTS.A2 * rb); const lclX = xdb - (CONSTANTS.A2 * rb);
        const uclR = CONSTANTS.D4 * rb; const lclR = CONSTANTS.D3 * rb;
        updateSpcStatsUI(xdb, uclX, lclX, rb, uclR, lclR);
        let usl = parseFloat(document.getElementById('uslInput').value); let lsl = parseFloat(document.getElementById('lslInput').value);
        if (!isNaN(usl) || !isNaN(lsl)) {
            const si = rb / CONSTANTS.d2; const so = calculateStdDev(allData, xdb);
            updateSpcCapUI(calculateCapability(xdb, si, so, usl, lsl));
        } else { clearSpcCapUI(); }
        const xvals = subgroups.map(g => g.xbar); const violations = checkSPCRules(xvals, uclX, lclX, xdb);
        displayViolations(violations, subgroups);
        drawSpcChart('xbar', subgroups.map(g => g.id), xvals, uclX, lclX, xdb, 'X-Bar', violations, usl, lsl);
        drawSpcChart('range', subgroups.map(g => g.id), subgroups.map(g => g.r), uclR, lclR, rb, 'Range', [], null, null);
    } catch (e) { console.error(e); }
}
function spcSetText(id, v) { const e = document.getElementById(id); if (e) e.innerText = v; }
function updateSpcStatsUI(x, ux, lx, r, ur, lr) { const f = n => isNaN(n) ? '-' : n.toFixed(3); spcSetText('val-xdb', f(x)); spcSetText('val-uclx', f(ux)); spcSetText('val-lclx', f(lx)); spcSetText('val-rb', f(r)); spcSetText('val-uclr', f(ur)); spcSetText('val-lclr', lr === 0 ? "0.000" : f(lr)); }
function spcClearUI() { ['val-xdb', 'val-uclx', 'val-lclx', 'val-rb', 'val-uclr', 'val-lclr'].forEach(id => spcSetText(id, '-')); clearSpcCapUI(); }
function updateSpcCapUI(c) { const f = n => (n !== null && !isNaN(n)) ? n.toFixed(3) : '-'; spcSetText('val-cp', f(c.cp)); spcSetText('val-cpk', f(c.cpk)); spcSetText('val-pp', f(c.pp)); spcSetText('val-ppk', f(c.ppk)); }
function clearSpcCapUI() { ['val-cp', 'val-cpk', 'val-pp', 'val-ppk'].forEach(id => spcSetText(id, '-')); }
function drawSpcChart(type, lbls, data, ucl, lcl, cl, title, vio, usl, lsl) {
    const ctx = document.getElementById(type === 'xbar' ? 'xbarChart' : 'rChart').getContext('2d');
    ctx.save(); ctx.globalCompositeOperation = 'destination-over'; ctx.fillStyle = 'white'; ctx.fillRect(0, 0, ctx.canvas.width, ctx.canvas.height); ctx.restore();
    let cols = Array(data.length).fill('#3498db'); let rads = Array(data.length).fill(4);
    if (vio) vio.forEach(v => { cols[v.index] = '#e74c3c'; rads[v.index] = 7; });
    let anns = {};
    if (type === 'xbar') {
        if (!isNaN(usl)) anns.usl = { type: 'line', yMin: usl, yMax: usl, borderColor: '#f39c12', borderDash: [4, 4], borderWidth: 2, label: { display: true, content: 'USL', position: 'end' } };
        if (!isNaN(lsl)) anns.lsl = { type: 'line', yMin: lsl, yMax: lsl, borderColor: '#f39c12', borderDash: [4, 4], borderWidth: 2, label: { display: true, content: 'LSL', position: 'end' } };
    }
    if (spc_charts[type]) spc_charts[type].destroy();
    spc_charts[type] = new Chart(ctx, { type: 'line', data: { labels: lbls, datasets: [{ label: 'Data', data: data, borderColor: '#3498db', pointBackgroundColor: cols, pointRadius: rads, fill: false, tension: 0 }, { label: 'UCL', data: Array(lbls.length).fill(ucl), borderColor: '#e74c3c', borderDash: [5, 5], pointRadius: 0 }, { label: 'CL', data: Array(lbls.length).fill(cl), borderColor: '#2ecc71', pointRadius: 0 }, { label: 'LCL', data: Array(lbls.length).fill(lcl), borderColor: '#e74c3c', borderDash: [5, 5], pointRadius: 0 }] }, options: { responsive: true, maintainAspectRatio: false, animation: false, plugins: { annotation: { annotations: anns }, legend: { display: true, position: 'top' } }, scales: { y: { title: { display: true, text: title } } } } });
}
function clearSpcData(redraw = true) { for (let i = 1; i <= CONFIG.numSubgroups; i++) for (let j = 1; j <= 5; j++) { let el = document.getElementById(`cell_${i}_${j}`); if (el) el.value = ""; } if (redraw) calculateAndDraw(); }
function initSpcTable() { const table = document.getElementById('inputTable'); let html = `<thead><tr><th style="width:35px">n</th>`; for (let j = 1; j <= 5; j++) html += `<th>X${j}</th>`; html += `<th>X̄</th><th>R</th></tr></thead><tbody id="spcTableBody"></tbody>`; table.innerHTML = html; renderSpcRows(1, CONFIG.numSubgroups); }
function renderSpcRows(start, end) { const tbody = document.getElementById('spcTableBody'); let html = ''; for (let i = start; i <= end; i++) { html += `<tr><td>${i}</td>`; for (let j = 1; j <= 5; j++) html += `<td><input type="number" id="cell_${i}_${j}" onchange="calculateAndDraw()"></td>`; html += `<td class="calc-res" id="res_xbar_${i}">-</td><td class="calc-res" id="res_r_${i}">-</td></tr>`; } if (start === 1) tbody.innerHTML = html; else tbody.insertAdjacentHTML('beforeend', html); }
function addMoreRows() { if (CONFIG.numSubgroups >= CONFIG.maxRows) return alert("已達最大列數"); let s = CONFIG.numSubgroups + 1; let e = Math.min(CONFIG.numSubgroups + 25, CONFIG.maxRows); CONFIG.numSubgroups = e; renderSpcRows(s, e); }
function initSpcResizer() { const resizer = document.getElementById('dragMe'); const container = document.getElementById('spcMainContainer'); let isResizing = false; resizer.addEventListener('mousedown', () => { isResizing = true; document.body.style.cursor = 'col-resize'; }); document.addEventListener('mousemove', (e) => { if (!isResizing) return; let w = e.clientX; if (w < 300) w = 300; if (w > 800) w = 800; container.style.setProperty('--left-width', `${w}px`); }); document.addEventListener('mouseup', () => { isResizing = false; document.body.style.cursor = 'default'; }); }
async function exportSpcToExcel() { const wb = new ExcelJS.Workbook(); const ws = wb.addWorksheet('SPC'); ws.columns = [{ key: 'id' }, { key: 'x1' }, { key: 'x2' }, { key: 'x3' }, { key: 'x4' }, { key: 'x5' }, { key: 'xb' }, { key: 'r' }]; for (let i = 1; i <= CONFIG.numSubgroups; i++) { let v1 = document.getElementById(`cell_${i}_1`)?.value; if (!v1) continue; ws.addRow({ id: i, x1: parseFloat(v1), x2: parseFloat(document.getElementById(`cell_${i}_2`).value), x3: parseFloat(document.getElementById(`cell_${i}_3`).value), x4: parseFloat(document.getElementById(`cell_${i}_4`).value), x5: parseFloat(document.getElementById(`cell_${i}_5`).value), xb: parseFloat(document.getElementById(`res_xbar_${i}`).innerText), r: parseFloat(document.getElementById(`res_r_${i}`).innerText) }); } const b = await wb.xlsx.writeBuffer(); saveAs(new Blob([b]), `SPC_Report.xlsx`); }
function checkSPCRules(data, ucl, lcl, cl) { let vs = []; for (let i = 0; i < data.length; i++) { let v = data[i]; if (v === null) continue; let rs = []; if (v > ucl || v < lcl) rs.push("超出界限"); if (i >= 8) { let side = v > cl ? 1 : -1, run = true; for (let k = 1; k < 9; k++) if (data[i - k] === null || (data[i - k] > cl ? 1 : -1) !== side) { run = false; break; } if (run) rs.push("連9點同側"); } if (rs.length > 0) vs.push({ index: i, rules: rs }); } return vs; }
function displayViolations(violations, subgroups) { const box = document.getElementById('rule-violations'); box.innerHTML = ""; violations.forEach(v => { let div = document.createElement('div'); div.className = 'violation-item'; div.innerText = `第 ${subgroups[v.index].id} 組: ${v.rules.join(', ')}`; box.appendChild(div); }); }
function downloadChart(id) { const link = document.createElement('a'); link.download = id + ".png"; link.href = document.getElementById(id).toDataURL(); link.click(); }
function calculateStdDev(data, mean) { if (!data || data.length < 2) return 0; if (mean === undefined) { mean = data.reduce((a, b) => a + b, 0) / data.length; } const variance = data.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / (data.length - 1); return Math.sqrt(variance); }
function calculateCapability(mean, sigmaWithin, sigmaOverall, usl, lsl) { let cp = NaN, cpk = NaN, pp = NaN, ppk = NaN; if (!isNaN(usl) && !isNaN(lsl)) { if (sigmaWithin > 0) cp = (usl - lsl) / (6 * sigmaWithin); if (sigmaOverall > 0) pp = (usl - lsl) / (6 * sigmaOverall); } const calcK = (sigma) => { if (sigma === 0) return NaN; let k_u = NaN, k_l = NaN; if (!isNaN(usl)) k_u = (usl - mean) / (3 * sigma); if (!isNaN(lsl)) k_l = (mean - lsl) / (3 * sigma); if (!isNaN(k_u) && !isNaN(k_l)) return Math.min(k_u, k_l); if (!isNaN(k_u)) return k_u; if (!isNaN(k_l)) return k_l; return NaN; }; cpk = calcK(sigmaWithin); ppk = calcK(sigmaOverall); return { cp, cpk, pp, ppk }; }

// --- Pareto Logic ---
let paretoChart = null;
function handleParetoSourceChange() { const mode = document.getElementById('pareto-source-mode').value; document.getElementById('pareto-file-upload').style.display = mode === 'new' ? 'block' : 'none'; if (mode === 'cpk') { if (SharedData.hasData) populateParetoSelectors(SharedData.headers); else alert("請先在 Cpk 頁面上傳數據"); } }
function handleParetoFileSelect(e) { const file = e.target.files[0]; if (!file) return; const reader = new FileReader(); reader.onload = (evt) => { const wb = XLSX.read(new Uint8Array(evt.target.result), { type: 'array' }); const sheet = wb.Sheets[wb.SheetNames[0]]; const res = robustParseExcel(sheet); if (res.headers.length > 0) { SharedData.tempPareto = res; populateParetoSelectors(res.headers); alert("柏拉圖數據已讀取"); } }; reader.readAsArrayBuffer(file); }
function populateParetoSelectors(headers) { const catSel = document.getElementById('pareto-cat-col'); const valSel = document.getElementById('pareto-val-col'); catSel.innerHTML = ''; valSel.innerHTML = '<option value="">-- 無 (計算出現次數) --</option>'; headers.forEach(h => { catSel.add(new Option(h, h)); valSel.add(new Option(h, h)); }); }
function calculatePareto() {
    const mode = document.getElementById('pareto-source-mode').value; let data = [];
    if (mode === 'cpk') data = SharedData.rawJson; else if (SharedData.tempPareto) data = SharedData.tempPareto.data;
    if (!data || data.length === 0) return alert("無數據");
    const catKey = document.getElementById('pareto-cat-col').value; const valKey = document.getElementById('pareto-val-col').value;
    if (!catKey) return alert("請選擇類別欄位");
    let counts = {}; let total = 0;
    data.forEach(row => { const cat = row[catKey]; if (cat === undefined || cat === null || String(cat).trim() === "") return; let val = 1; if (valKey) { val = parseFloat(String(row[valKey]).replace(/[^0-9.\-]/g, '')); if (isNaN(val)) val = 0; } if (!counts[cat]) counts[cat] = 0; counts[cat] += val; total += val; });
    let sorted = Object.entries(counts).sort((a, b) => b[1] - a[1]);
    let cum = 0; let chartLabels = []; let chartDataBar = []; let chartDataLine = [];
    sorted.forEach(([k, v]) => { cum += v; chartLabels.push(k); chartDataBar.push(v); chartDataLine.push((cum / total) * 100); });
    drawParetoChart(chartLabels, chartDataBar, chartDataLine);
}
function drawParetoChart(labels, bars, lines) {
    const ctx = document.getElementById('pareto-chart').getContext('2d'); if (paretoChart) paretoChart.destroy();
    paretoChart = new Chart(ctx, { type: 'bar', data: { labels: labels, datasets: [{ label: '累積百分比 (%)', data: lines, type: 'line', borderColor: '#e74c3c', yAxisID: 'y1', tension: 0.1, pointRadius: 4 }, { label: '數量/數值', data: bars, backgroundColor: '#3498db', yAxisID: 'y' }] }, options: { responsive: true, maintainAspectRatio: false, scales: { y: { beginAtZero: true, title: { display: true, text: '數量' } }, y1: { beginAtZero: true, max: 100, position: 'right', title: { display: true, text: '累積百分比 (%)' }, grid: { drawOnChartArea: false } } } } });
}

// --- Scatter Logic ---
let scatterChart = null;
function handleScatterSourceChange() { const mode = document.getElementById('scatter-source-mode').value; document.getElementById('scatter-file-upload').style.display = mode === 'new' ? 'block' : 'none'; if (mode === 'cpk') { if (SharedData.hasData) populateScatterSelectors(SharedData.headers); else alert("請先在 Cpk 頁面上傳數據"); } }
function handleScatterFileSelect(e) { const file = e.target.files[0]; if (!file) return; const reader = new FileReader(); reader.onload = (evt) => { const wb = XLSX.read(new Uint8Array(evt.target.result), { type: 'array' }); const sheet = wb.Sheets[wb.SheetNames[0]]; const res = robustParseExcel(sheet); if (res.headers.length > 0) { SharedData.tempScatter = res; populateScatterSelectors(res.headers); alert("散佈圖數據已讀取"); } }; reader.readAsArrayBuffer(file); }
function populateScatterSelectors(headers) { const xSel = document.getElementById('scatter-x-col'); const ySel = document.getElementById('scatter-y-col'); const gSel = document.getElementById('scatter-group-col'); [xSel, ySel, gSel].forEach(s => s.innerHTML = ''); gSel.add(new Option("-- 不分組 --", "")); headers.forEach(h => { xSel.add(new Option(h, h)); ySel.add(new Option(h, h)); gSel.add(new Option(h, h)); }); if (ySel.options.length > 1) ySel.selectedIndex = 1; }
function calculateScatter() {
    const mode = document.getElementById('scatter-source-mode').value; let data = [];
    if (mode === 'cpk') data = SharedData.rawJson; else if (SharedData.tempScatter) data = SharedData.tempScatter.data;
    if (!data || data.length === 0) return alert("無數據");
    const xKey = document.getElementById('scatter-x-col').value; const yKey = document.getElementById('scatter-y-col').value; const gKey = document.getElementById('scatter-group-col').value;
    if (!xKey || !yKey) return alert("請選擇 X 和 Y 軸欄位");
    let datasets = {}; let xSum = 0, ySum = 0, n = 0, xSqSum = 0, ySqSum = 0, xySum = 0;
    data.forEach(row => {
        let x = parseFloat(String(row[xKey]).replace(/[^0-9.\-]/g, '')); let y = parseFloat(String(row[yKey]).replace(/[^0-9.\-]/g, ''));
        if (!isNaN(x) && !isNaN(y)) { let group = "Data"; if (gKey && row[gKey] !== undefined) group = row[gKey]; if (!datasets[group]) datasets[group] = []; datasets[group].push({ x, y }); xSum += x; ySum += y; xSqSum += x * x; ySqSum += y * y; xySum += x * y; n++; }
    });
    if (n < 2) return alert("有效數據點不足");
    let num = n * xySum - xSum * ySum; let den = Math.sqrt((n * xSqSum - xSum * xSum) * (n * ySqSum - ySum * ySum)); let r = den === 0 ? 0 : num / den;
    document.getElementById('val-r').innerText = r.toFixed(4);
    drawScatterChart(datasets, xKey, yKey);
}
function drawScatterChart(datasetsObj, xLabel, yLabel) {
    const ctx = document.getElementById('scatter-chart').getContext('2d'); if (scatterChart) scatterChart.destroy();
    const colors = ['#3498db', '#e74c3c', '#2ecc71', '#f1c40f', '#9b59b6', '#34495e']; let finalDatasets = []; let i = 0;
    for (const [grp, pts] of Object.entries(datasetsObj)) { finalDatasets.push({ label: grp, data: pts, backgroundColor: colors[i % colors.length], pointRadius: 5 }); i++; }
    scatterChart = new Chart(ctx, { type: 'scatter', data: { datasets: finalDatasets }, options: { responsive: true, maintainAspectRatio: false, scales: { x: { title: { display: true, text: xLabel }, type: 'linear', position: 'bottom' }, y: { title: { display: true, text: yLabel } } }, plugins: { legend: { display: true } } } });
}

// --- Box Plot Logic ---
let boxplotChart = null;
function handleBoxplotSourceChange() { const mode = document.getElementById('boxplot-source-mode').value; document.getElementById('boxplot-file-upload').style.display = mode === 'new' ? 'block' : 'none'; if (mode === 'cpk') { if (SharedData.hasData) populateBoxplotSelectors(SharedData.headers); else alert("請先在 Cpk 頁面上傳數據"); } }
function handleBoxplotFileSelect(e) { const file = e.target.files[0]; if (!file) return; const reader = new FileReader(); reader.onload = (evt) => { const wb = XLSX.read(new Uint8Array(evt.target.result), { type: 'array' }); const sheet = wb.Sheets[wb.SheetNames[0]]; const res = robustParseExcel(sheet); if (res.headers.length > 0) { SharedData.tempBoxplot = res; populateBoxplotSelectors(res.headers); alert("盒鬚圖數據已讀取"); } }; reader.readAsArrayBuffer(file); }
function populateBoxplotSelectors(headers) { const groupSel = document.getElementById('boxplot-group-col'); const valueSel = document.getElementById('boxplot-value-col'); [groupSel, valueSel].forEach(s => s.innerHTML = ''); headers.forEach(h => { groupSel.add(new Option(h, h)); valueSel.add(new Option(h, h)); }); if (valueSel.options.length > 1) valueSel.selectedIndex = 1; }
function calculateBoxplot() {
    const mode = document.getElementById('boxplot-source-mode').value; let data = [];
    if (mode === 'cpk') data = SharedData.rawJson; else if (SharedData.tempBoxplot) data = SharedData.tempBoxplot.data;
    if (!data || data.length === 0) return alert("無數據");
    const groupKey = document.getElementById('boxplot-group-col').value; const valueKey = document.getElementById('boxplot-value-col').value;
    if (!groupKey || !valueKey) return alert("請選擇分組和數值欄位");
    let grouped = {};
    data.forEach(row => { const group = row[groupKey]; if (group === undefined || group === null) return; let val = parseFloat(String(row[valueKey]).replace(/[^0-9.\-]/g, '')); if (!isNaN(val)) { if (!grouped[group]) grouped[group] = []; grouped[group].push(val); } });
    if (Object.keys(grouped).length === 0) return alert("無有效數據");
    let boxplotData = [];
    for (const [group, values] of Object.entries(grouped)) { if (values.length < 2) continue; const sorted = values.slice().sort((a, b) => a - b); const stats = calculateBoxplotStats(sorted); boxplotData.push({ label: group, values: values, ...stats }); }
    if (boxplotData.length === 0) return alert("數據不足以繪製盒鬚圖");
    const showMean = document.getElementById('boxplot-show-mean').checked; const showOutliers = document.getElementById('boxplot-show-outliers').checked; const runANOVA = document.getElementById('boxplot-run-anova').checked;
    drawBoxplotChart(boxplotData, showMean, showOutliers);
    if (runANOVA) { const anovaResult = calculateANOVA(grouped); displayANOVA(anovaResult); } else { document.getElementById('anova-results').style.display = 'none'; }
}
function calculateBoxplotStats(sortedData) {
    const n = sortedData.length; const min = sortedData[0]; const max = sortedData[n - 1];
    const q1Index = Math.floor(n * 0.25); const q2Index = Math.floor(n * 0.5); const q3Index = Math.floor(n * 0.75);
    const q1 = sortedData[q1Index]; const median = n % 2 === 0 ? (sortedData[q2Index - 1] + sortedData[q2Index]) / 2 : sortedData[q2Index]; const q3 = sortedData[q3Index];
    const mean = sortedData.reduce((a, b) => a + b, 0) / n;
    const iqr = q3 - q1; const lowerFence = q1 - 1.5 * iqr; const upperFence = q3 + 1.5 * iqr;
    const outliers = sortedData.filter(v => v < lowerFence || v > upperFence);
    return { min, q1, median, q3, max, mean, outliers, iqr };
}
function drawBoxplotChart(boxplotData, showMean, showOutliers) {
    const ctx = document.getElementById('boxplot-chart').getContext('2d'); if (boxplotChart) boxplotChart.destroy();
    const labels = boxplotData.map(d => d.label);
    const datasets = [{ label: '盒鬚圖', backgroundColor: 'rgba(54, 162, 235, 0.5)', borderColor: 'rgb(54, 162, 235)', borderWidth: 2, outlierBackgroundColor: 'rgba(231, 76, 60, 0.8)', outlierRadius: 4, itemRadius: 0, data: boxplotData.map(d => { const boxData = { min: d.min, q1: d.q1, median: d.median, q3: d.q3, max: d.max }; if (showOutliers && d.outliers.length > 0) { boxData.outliers = d.outliers; } if (showMean) { boxData.mean = d.mean; } return boxData; }) }];
    boxplotChart = new Chart(ctx, { type: 'boxplot', data: { labels: labels, datasets: datasets }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: true }, tooltip: { callbacks: { label: function (context) { const data = context.parsed; return [`最小值: ${data.min.toFixed(3)}`, `Q1: ${data.q1.toFixed(3)}`, `中位數: ${data.median.toFixed(3)}`, `Q3: ${data.q3.toFixed(3)}`, `最大值: ${data.max.toFixed(3)}`, showMean ? `平均值: ${data.mean.toFixed(3)}` : ''].filter(s => s); } } } }, scales: { y: { title: { display: true, text: '數值' } } } } });
}
function calculateANOVA(groupedData) {
    const groups = Object.values(groupedData); const groupNames = Object.keys(groupedData); const k = groups.length;
    let N = 0; groups.forEach(g => N += g.length);
    let grandSum = 0; groups.forEach(g => { g.forEach(v => grandSum += v); }); const grandMean = grandSum / N;
    let ssb = 0; groups.forEach(g => { const groupMean = g.reduce((a, b) => a + b, 0) / g.length; ssb += g.length * Math.pow(groupMean - grandMean, 2); });
    let ssw = 0; groups.forEach(g => { const groupMean = g.reduce((a, b) => a + b, 0) / g.length; g.forEach(v => { ssw += Math.pow(v - groupMean, 2); }); });
    const sst = ssb + ssw; const dfb = k - 1; const dfw = N - k; const dft = N - 1;
    const msb = ssb / dfb; const msw = ssw / dfw; const f = msb / msw;
    let pValue = "< 0.05"; if (f < 3.0) pValue = "> 0.05"; else if (f < 5.0) pValue = "< 0.05"; else if (f < 10.0) pValue = "< 0.01"; else pValue = "< 0.001";
    return { ssb, ssw, sst, dfb, dfw, dft, msb, msw, f, pValue, groupNames };
}
function displayANOVA(result) {
    const container = document.getElementById('anova-results'); const tableDiv = document.getElementById('anova-table');
    let html = `<table style="width: 100%; border-collapse: collapse;"><thead><tr style="background: #3498db; color: white;"><th style="padding: 10px; border: 1px solid #ddd;">變異來源</th><th style="padding: 10px; border: 1px solid #ddd;">平方和 (SS)</th><th style="padding: 10px; border: 1px solid #ddd;">自由度 (df)</th><th style="padding: 10px; border: 1px solid #ddd;">均方 (MS)</th><th style="padding: 10px; border: 1px solid #ddd;">F 值</th><th style="padding: 10px; border: 1px solid #ddd;">p 值</th></tr></thead><tbody><tr><td style="padding: 8px; border: 1px solid #ddd;">組間 (Between)</td><td style="padding: 8px; border: 1px solid #ddd;">${result.ssb.toFixed(4)}</td><td style="padding: 8px; border: 1px solid #ddd;">${result.dfb}</td><td style="padding: 8px; border: 1px solid #ddd;">${result.msb.toFixed(4)}</td><td style="padding: 8px; border: 1px solid #ddd;" rowspan="2">${result.f.toFixed(4)}</td><td style="padding: 8px; border: 1px solid #ddd;" rowspan="2">${result.pValue}</td></tr><tr><td style="padding: 8px; border: 1px solid #ddd;">組內 (Within)</td><td style="padding: 8px; border: 1px solid #ddd;">${result.ssw.toFixed(4)}</td><td style="padding: 8px; border: 1px solid #ddd;">${result.dfw}</td><td style="padding: 8px; border: 1px solid #ddd;">${result.msw.toFixed(4)}</td></tr><tr style="background: #ecf0f1;"><td style="padding: 8px; border: 1px solid #ddd;"><strong>總計 (Total)</strong></td><td style="padding: 8px; border: 1px solid #ddd;"><strong>${result.sst.toFixed(4)}</strong></td><td style="padding: 8px; border: 1px solid #ddd;"><strong>${result.dft}</strong></td><td style="padding: 8px; border: 1px solid #ddd;">-</td><td style="padding: 8px; border: 1px solid #ddd;">-</td><td style="padding: 8px; border: 1px solid #ddd;">-</td></tr></tbody></table><div style="margin-top: 15px; padding: 10px; background: #e8f4f8; border-left: 4px solid #3498db;"><strong>結論:</strong> ${result.pValue.includes('<') && !result.pValue.includes('> 0.05') ? '組間差異<strong>顯著</strong> (拒絕虛無假設)' : '組間差異<strong>不顯著</strong> (接受虛無假設)'}</div>`;
    tableDiv.innerHTML = html; container.style.display = 'block';
}

// --- Gage R&R Logic ---

function handleGrrSourceChange() {
    const mode = document.getElementById('grr-source-mode').value;
    document.getElementById('grr-file-upload').style.display = mode === 'new' ? 'block' : 'none';
    if (mode === 'cpk') {
        if (SharedData.hasData) populateGrrSelectors(SharedData.headers);
        else alert("請先在 Cpk 頁面上傳數據");
    }
}

function handleGrrFileSelect(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
        try {
            const wb = XLSX.read(new Uint8Array(evt.target.result), { type: 'array' });
            const sheet = wb.Sheets[wb.SheetNames[0]];
            const res = robustParseExcel(sheet);
            if (res.headers.length > 0) {
                SharedData.tempGrr = res;
                populateGrrSelectors(res.headers);
                alert("Gage R&R 數據已讀取");
            }
        } catch (err) { console.error(err); alert("讀取失敗"); }
    };
    reader.readAsArrayBuffer(file);
}

function populateGrrSelectors(headers) {
    const opSel = document.getElementById('grr-operator-col');
    const partSel = document.getElementById('grr-part-col');
    const measSel = document.getElementById('grr-measurement-col');
    [opSel, partSel, measSel].forEach(s => s.innerHTML = '');
    headers.forEach(h => {
        const opt = new Option(h, h);
        opSel.add(opt.cloneNode(true));
        partSel.add(opt.cloneNode(true));
        measSel.add(opt.cloneNode(true));
    });
    // Auto-select heuristic
    if (headers.length >= 3) {
        partSel.selectedIndex = 1;
        measSel.selectedIndex = 2;
    }
}

function calculateGrr() {
    const mode = document.getElementById('grr-source-mode').value;
    let data = [];
    if (mode === 'cpk') data = SharedData.rawJson;
    else if (SharedData.tempGrr) data = SharedData.tempGrr.data;

    if (!data || data.length === 0) return alert("無數據");

    const opKey = document.getElementById('grr-operator-col').value;
    const partKey = document.getElementById('grr-part-col').value;
    const measKey = document.getElementById('grr-measurement-col').value;

    if (!opKey || !partKey || !measKey) return alert("請選擇所有必要的欄位");

    // Prepare data structure for ANOVA & Charts
    let cleanData = [];
    data.forEach(row => {
        let op = row[opKey];
        let part = row[partKey];
        let val = parseFloat(String(row[measKey]).replace(/[^0-9.\-]/g, ''));
        if (op !== undefined && part !== undefined && !isNaN(val)) {
            cleanData.push({ op: String(op), part: String(part), val: val });
        }
    });

    if (cleanData.length < 10) return alert("數據太少，無法進行有效分析");

    try {
        const result = computeAnovaGrr(cleanData);
        renderGrrResults(result);
    } catch (e) {
        console.error(e);
        alert("計算發生錯誤: " + e.message + "\n請確認數據格式是否正確 (需有多個作業員、多個部件、重複量測)");
    }
}

function getControlChartConstants(n) {
    // n: Sample size (replications)
    if (n === 2) return { A2: 1.880, D3: 0, D4: 3.267 };
    if (n === 3) return { A2: 1.023, D3: 0, D4: 2.574 };
    if (n === 4) return { A2: 0.729, D3: 0, D4: 2.282 };
    if (n === 5) return { A2: 0.577, D3: 0, D4: 2.114 };
    if (n === 6) return { A2: 0.483, D3: 0, D4: 2.004 };
    if (n === 7) return { A2: 0.419, D3: 0.076, D4: 1.924 };
    if (n === 8) return { A2: 0.373, D3: 0.136, D4: 1.864 };
    if (n === 9) return { A2: 0.337, D3: 0.184, D4: 1.816 };
    if (n >= 10) return { A2: 0.308, D3: 0.223, D4: 1.777 };
    return { A2: 0.577, D3: 0, D4: 2.114 }; // Default fallback
}

function computeAnovaGrr(data) {
    // 1. Identify unique levels
    const operators = [...new Set(data.map(d => d.op))].sort();
    const parts = [...new Set(data.map(d => d.part))].sort();
    const a = operators.length; // Number of operators
    const b = parts.length;     // Number of parts
    
    // Group measurements to find n (replications)
    const totalCount = data.length;
    const n = Math.round(totalCount / (a * b)); 

    if (n < 2) throw new Error("每個部件每位作業員至少需要 2 次量測 (n >= 2)");

    // 2. Calculate Sums
    let grandSum = 0;
    let sumSqTotal = 0;
    let sumOp = {};
    let sumPart = {};
    let sumOpPart = {}; // Interaction sums
    
    // For X-bar R calculations later
    let subgroups = {}; // Key: "Op|Part", Value: [measurements]

    data.forEach(d => {
        grandSum += d.val;
        sumSqTotal += d.val * d.val;

        sumOp[d.op] = (sumOp[d.op] || 0) + d.val;
        sumPart[d.part] = (sumPart[d.part] || 0) + d.val;
        
        const intKey = d.op + "|" + d.part;
        sumOpPart[intKey] = (sumOpPart[intKey] || 0) + d.val;

        if (!subgroups[intKey]) subgroups[intKey] = [];
        subgroups[intKey].push(d.val);
    });

    const CF = (grandSum * grandSum) / totalCount;
    const SS_Total = sumSqTotal - CF;

    // SS Operator
    let sumSqOp = 0; for (let k in sumOp) sumSqOp += sumOp[k] * sumOp[k];
    const SS_Op = (sumSqOp / (b * n)) - CF;

    // SS Part
    let sumSqPart = 0; for (let k in sumPart) sumSqPart += sumPart[k] * sumPart[k];
    const SS_Part = (sumSqPart / (a * n)) - CF;

    // SS Interaction (Subtotals)
    let sumSqOpPart = 0; for (let k in sumOpPart) sumSqOpPart += sumOpPart[k] * sumOpPart[k];
    const SS_Subtotal = sumSqOpPart / n; 
    const SS_Interaction = SS_Subtotal - SS_Op - SS_Part - CF;

    // SS Error
    const SS_Error = SS_Total - SS_Op - SS_Part - SS_Interaction;

    // Degrees of Freedom
    const df_Op = a - 1;
    const df_Part = b - 1;
    const df_Int = (a - 1) * (b - 1);
    const df_Error = a * b * (n - 1);
    const df_Total = (a * b * n) - 1;

    // Mean Squares
    const MS_Op = SS_Op / df_Op;
    const MS_Part = SS_Part / df_Part;
    const MS_Int = SS_Interaction / df_Int;
    const MS_Error = SS_Error / df_Error;

    // F values
    const F_Op = MS_Op / MS_Int;
    const F_Part = MS_Part / MS_Int;
    const F_Int = MS_Int / MS_Error;

    // Variance Components
    let var_Equipment = MS_Error;
    let var_Interaction = (MS_Int - MS_Error) / n;
    if (var_Interaction < 0) var_Interaction = 0;

    let var_Operator = (MS_Op - MS_Int) / (b * n);
    if (var_Operator < 0) var_Operator = 0;

    let var_Part = (MS_Part - MS_Int) / (a * n);
    if (var_Part < 0) var_Part = 0;

    const var_GRR = var_Equipment + var_Operator + var_Interaction;
    const var_Total = var_GRR + var_Part;

    // Stats
    const std_Total = Math.sqrt(var_Total);
    const stats = {
        EV: { var: var_Equipment, std: Math.sqrt(var_Equipment) },
        AV: { var: var_Operator + var_Interaction, std: Math.sqrt(var_Operator + var_Interaction) },
        GRR: { var: var_GRR, std: Math.sqrt(var_GRR) },
        PV: { var: var_Part, std: Math.sqrt(var_Part) },
        TV: { var: var_Total, std: std_Total }
    };

    const report = [
        { name: "Repeatability (EV)", var: stats.EV.var, pctCont: (stats.EV.var/var_Total)*100, pctStudy: (stats.EV.std/std_Total)*100 },
        { name: "Reproducibility (AV)", var: stats.AV.var, pctCont: (stats.AV.var/var_Total)*100, pctStudy: (stats.AV.std/std_Total)*100 },
        { name: "Interaction", var: var_Interaction, pctCont: (var_Interaction/var_Total)*100, pctStudy: (Math.sqrt(var_Interaction)/std_Total)*100 },
        { name: "Gage R&R (GRR)", var: stats.GRR.var, pctCont: (stats.GRR.var/var_Total)*100, pctStudy: (stats.GRR.std/std_Total)*100 },
        { name: "Part Variation (PV)", var: stats.PV.var, pctCont: (stats.PV.var/var_Total)*100, pctStudy: (stats.PV.std/std_Total)*100 },
        { name: "Total Variation (TV)", var: stats.TV.var, pctCont: 100.0, pctStudy: 100.0 }
    ];

    return {
        anova: { SS: [SS_Op, SS_Part, SS_Interaction, SS_Error, SS_Total], df: [df_Op, df_Part, df_Int, df_Error, df_Total], MS: [MS_Op, MS_Part, MS_Int, MS_Error], F: [F_Op, F_Part, F_Int] },
        report: report,
        rawData: data,
        subgroups: subgroups,
        meta: { n: n, operators: operators, parts: parts }
    };
}

function renderGrrResults(result) {
    document.getElementById('grr-results-container').style.display = 'block';
    
    // 1. ANOVA Table HTML
    let anovaHtml = `
        <table class="w-full text-sm text-left">
            <thead class="bg-gray-100">
                <tr><th>來源</th><th>DF</th><th>SS</th><th>MS</th><th>F</th></tr>
            </thead>
            <tbody>
                <tr><td>作業員 (Op)</td><td>${result.anova.df[0]}</td><td>${result.anova.SS[0].toFixed(4)}</td><td>${result.anova.MS[0].toFixed(4)}</td><td>${result.anova.F[0].toFixed(2)}</td></tr>
                <tr><td>部件 (Part)</td><td>${result.anova.df[1]}</td><td>${result.anova.SS[1].toFixed(4)}</td><td>${result.anova.MS[1].toFixed(4)}</td><td>${result.anova.F[1].toFixed(2)}</td></tr>
                <tr><td>交互作用 (Int)</td><td>${result.anova.df[2]}</td><td>${result.anova.SS[2].toFixed(4)}</td><td>${result.anova.MS[2].toFixed(4)}</td><td>${result.anova.F[2].toFixed(2)}</td></tr>
                <tr><td>重複性 (Error)</td><td>${result.anova.df[3]}</td><td>${result.anova.SS[3].toFixed(4)}</td><td>${result.anova.MS[3].toFixed(4)}</td><td></td></tr>
                <tr style="font-weight:bold; background:#eee;"><td>總計 (Total)</td><td>${result.anova.df[4]}</td><td>${result.anova.SS[4].toFixed(4)}</td><td></td><td></td></tr>
            </tbody>
        </table>
    `;
    document.getElementById('grr-anova-table').innerHTML = anovaHtml;

    // 2. Report Table HTML
    const grrPct = result.report[3].pctStudy;
    let reportHtml = `
        <table class="w-full text-sm text-left">
            <thead class="bg-gray-100">
                <tr><th>變異來源</th><th>變異分量 (VarComp)</th><th>% 貢獻度 (%Contribution)</th><th>% 研究變異 (%Study Var)</th></tr>
            </thead>
            <tbody>
    `;
    result.report.forEach(row => {
        let style = row.name.includes("Total") ? "font-weight:bold; background:#eee;" : "";
        if (row.name.includes("Gage R&R")) style = "font-weight:bold; color:#d35400;";
        reportHtml += `<tr style="${style}">
            <td>${row.name}</td>
            <td>${row.var.toFixed(6)}</td>
            <td>${row.pctCont.toFixed(2)} %</td>
            <td>${row.pctStudy.toFixed(2)} %</td>
        </tr>`;
    });
    reportHtml += `</tbody></table>`;
    document.getElementById('grr-report-table').innerHTML = reportHtml;

    // 3. Verdict
    const summaryCard = document.getElementById('grr-summary-card');
    const verdictText = document.getElementById('grr-verdict-text');
    document.getElementById('grr-final-percent').innerText = grrPct.toFixed(2);
    summaryCard.style.display = 'block';

    if (grrPct < 10) {
        verdictText.innerText = "✅ 量測系統可接受 (Excellent)";
        verdictText.style.color = "#27ae60";
    } else if (grrPct < 30) {
        verdictText.innerText = "⚠️ 條件下可接受 (Acceptable)";
        verdictText.style.color = "#f39c12";
    } else {
        verdictText.innerText = "❌ 量測系統無法接受 (Unacceptable)";
        verdictText.style.color = "#c0392b";
    }

    // 4. Draw Six Pack Charts
    drawGrrSixPack(result);
}

function drawGrrSixPack(result) {
    // Clear old charts
    ['grr-chart-components', 'grr-chart-r', 'grr-chart-xbar', 'grr-chart-bypart', 'grr-chart-byop', 'grr-chart-interaction'].forEach(id => {
        if (grr_charts[id]) grr_charts[id].destroy();
    });

    const colors = ['#3498db', '#e74c3c', '#2ecc71', '#f1c40f', '#9b59b6', '#34495e'];

    // 1. Components of Variation (Stacked Bar)
    const grrVal = result.report[3].pctStudy;
    const pvVal = result.report[4].pctStudy;
    grr_charts['grr-chart-components'] = new Chart(document.getElementById('grr-chart-components'), {
        type: 'bar',
        data: {
            labels: ['變異來源'],
            datasets: [
                { label: '% Gage R&R', data: [grrVal], backgroundColor: '#e74c3c' },
                { label: '% Part Variation', data: [pvVal], backgroundColor: '#3498db' }
            ]
        },
        options: { indexAxis: 'y', responsive: true, maintainAspectRatio: false, scales: { x: { stacked: true, max: 100 }, y: { stacked: true } } }
    });

    // Calculate Control Limits Constants
    const consts = getControlChartConstants(result.meta.n);
    
    // Prepare Data for X-bar & R
    const subgroups = [];
    const xbars = [];
    const ranges = [];
    
    // Iterating subgroups
    // Key format: Op|Part
    const sortedKeys = Object.keys(result.subgroups).sort();
    sortedKeys.forEach(k => {
        const vals = result.subgroups[k];
        const mean = vals.reduce((a,b)=>a+b,0)/vals.length;
        const min = Math.min(...vals);
        const max = Math.max(...vals);
        const range = max - min;
        subgroups.push(k);
        xbars.push(mean);
        ranges.push(range);
    });

    const Rbar = ranges.reduce((a,b)=>a+b,0)/ranges.length;
    const Xbarbar = xbars.reduce((a,b)=>a+b,0)/xbars.length;

    const UCL_R = consts.D4 * Rbar;
    const LCL_R = consts.D3 * Rbar;
    const UCL_X = Xbarbar + (consts.A2 * Rbar);
    const LCL_X = Xbarbar - (consts.A2 * Rbar);

    // 2. R Chart
    grr_charts['grr-chart-r'] = new Chart(document.getElementById('grr-chart-r'), {
        type: 'line',
        data: {
            labels: subgroups,
            datasets: [{ label: 'Range', data: ranges, borderColor: '#3498db', tension: 0, fill: false }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: { 
                annotation: { 
                    annotations: {
                        ucl: { type: 'line', yMin: UCL_R, yMax: UCL_R, borderColor: 'red', borderWidth: 2, label: { display: true, content: 'UCL' } },
                        lcl: { type: 'line', yMin: LCL_R, yMax: LCL_R, borderColor: 'red', borderWidth: 2, label: { display: true, content: 'LCL' } },
                        cl: { type: 'line', yMin: Rbar, yMax: Rbar, borderColor: 'green', borderWidth: 2, label: { display: true, content: 'Rbar' } }
                    } 
                },
                legend: { display: false } 
            }
        }
    });

    // 3. X-Bar Chart
    grr_charts['grr-chart-xbar'] = new Chart(document.getElementById('grr-chart-xbar'), {
        type: 'line',
        data: {
            labels: subgroups,
            datasets: [{ label: 'Mean', data: xbars, borderColor: '#3498db', tension: 0, fill: false }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: { 
                annotation: { 
                    annotations: {
                        ucl: { type: 'line', yMin: UCL_X, yMax: UCL_X, borderColor: 'red', borderWidth: 2, label: { display: true, content: 'UCL' } },
                        lcl: { type: 'line', yMin: LCL_X, yMax: LCL_X, borderColor: 'red', borderWidth: 2, label: { display: true, content: 'LCL' } },
                        cl: { type: 'line', yMin: Xbarbar, yMax: Xbarbar, borderColor: 'green', borderWidth: 2, label: { display: true, content: 'Mean' } }
                    } 
                },
                legend: { display: false } 
            }
        }
    });

    // 4. By Part (Scatter with Mean Line)
    // Group by part
    const partGroups = {};
    result.meta.parts.forEach(p => partGroups[p] = []);
    result.rawData.forEach(d => partGroups[d.part].push(d.val));
    
    const byPartLabels = result.meta.parts;
    const byPartMeans = byPartLabels.map(p => partGroups[p].reduce((a,b)=>a+b,0)/partGroups[p].length);
    
    // Scatter points for By Part
    const byPartScatterData = [];
    result.rawData.forEach(d => {
        byPartScatterData.push({ x: d.part, y: d.val });
    });

    grr_charts['grr-chart-bypart'] = new Chart(document.getElementById('grr-chart-bypart'), {
        type: 'line',
        data: {
            labels: byPartLabels,
            datasets: [
                { type: 'line', label: 'Mean', data: byPartMeans, borderColor: '#3498db', tension: 0, fill: false },
                { type: 'scatter', label: 'Data', data: byPartScatterData, backgroundColor: '#95a5a6' }
            ]
        },
        options: { responsive: true, maintainAspectRatio: false }
    });

    // 5. By Operator
    const opGroups = {};
    result.meta.operators.forEach(o => opGroups[o] = []);
    result.rawData.forEach(d => opGroups[d.op].push(d.val));
    
    const byOpLabels = result.meta.operators;
    const byOpMeans = byOpLabels.map(o => opGroups[o].reduce((a,b)=>a+b,0)/opGroups[o].length);
    
    const byOpScatterData = [];
    result.rawData.forEach(d => {
        byOpScatterData.push({ x: d.op, y: d.val });
    });

    grr_charts['grr-chart-byop'] = new Chart(document.getElementById('grr-chart-byop'), {
        type: 'line',
        data: {
            labels: byOpLabels,
            datasets: [
                { type: 'line', label: 'Mean', data: byOpMeans, borderColor: '#e74c3c', tension: 0, fill: false },
                { type: 'scatter', label: 'Data', data: byOpScatterData, backgroundColor: '#95a5a6' }
            ]
        },
        options: { responsive: true, maintainAspectRatio: false }
    });

    // 6. Interaction (Operator * Part)
    // X-axis: Part, Lines: Operator
    const interactionDatasets = result.meta.operators.map((op, idx) => {
        const data = result.meta.parts.map(p => {
            const key = op + "|" + p;
            const vals = result.subgroups[key];
            return vals ? vals.reduce((a,b)=>a+b,0)/vals.length : null;
        });
        return {
            label: op,
            data: data,
            borderColor: colors[idx % colors.length],
            tension: 0,
            fill: false
        };
    });

    grr_charts['grr-chart-interaction'] = new Chart(document.getElementById('grr-chart-interaction'), {
        type: 'line',
        data: {
            labels: result.meta.parts,
            datasets: interactionDatasets
        },
        options: { responsive: true, maintainAspectRatio: false }
    });
}

// Version Control: v19.18-Fix-Wide-Mode-Row-Based-20251125