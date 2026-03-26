import { SharedData, SharedSpecs, CONFIG, CONSTANTS, runWorkerTask, showError, showSuccess } from '../utils.js';
import { populateCpkSelectors } from './cpk.js';

export let spc_charts = { xbar: null, range: null };
export let spcGroupedData = {};

export function initSpcEvents() {
    initSpcTable();
    initSpcResizer();
    
    const sourceMode = document.getElementById('spc-source-mode');
    if (sourceMode) sourceMode.addEventListener('change', handleSpcSourceChange);
    
    const singleFile = document.getElementById('spc-single-file');
    if (singleFile) singleFile.addEventListener('change', handleSpcSingleFile);
    
    document.getElementById('uslInput').addEventListener('change', window.calculateAndDraw);
    document.getElementById('lslInput').addEventListener('change', window.calculateAndDraw);
    
    handleSpcSourceChange();
}

export function handleSpcSourceChange() {
    const mode = document.getElementById('spc-source-mode').value;
    document.getElementById('spc-tool-cpk').style.display = mode === 'cpk' ? 'flex' : 'none';
    document.getElementById('spc-tool-new').style.display = mode === 'new' ? 'flex' : 'none';
    const itemSel = document.getElementById('spc-inherit-item-select');
    itemSel.style.display = 'none'; 
    itemSel.innerHTML = '<option value="">(請先讀取)</option>';
}

export function parseSpcInheritedData() {
    if (!SharedData.hasData) return showError("請先在 Cpk 頁面上傳數據");
    const catKey = document.getElementById('category-col').value; 
    const valKey = document.getElementById('value-col').value;
    processSpcGrouping(SharedData.rawJson, catKey, valKey);
}

async function handleSpcSingleFile(e) {
    const file = e.target.files[0]; 
    if (!file) return; 
    window.clearSpcData(false);
    
    const statusEl = document.getElementById('statusMsg');
    statusEl.innerHTML = '<span style="color:#f39c12">⏳ 解析數據中...</span>';

    try {
        const buffer = await file.arrayBuffer();
        const res = await runWorkerTask('parseFirstSheet', { buffer });
        
        let allValues = [];
        let targetColIdx = 0; 
        
        res.data.forEach(row => {
            const keys = Object.keys(row);
            if (keys.length > 0) {
                let val = parseFloat(String(row[keys[targetColIdx]]).replace(/[^0-9.\-]/g, ''));
                if (!isNaN(val)) allValues.push(val);
            }
        });
        
        if (allValues.length === 0) throw new Error("未偵測到有效數值");
        
        fillSpcTableFromValues(allValues); 
        window.calculateAndDraw();
        statusEl.innerHTML = '系統就緒';
    } catch (err) {
        showError("讀取失敗: " + err.message);
        statusEl.innerHTML = '<span style="color:#e74c3c">讀取失敗</span>';
    }
}

function processSpcGrouping(dataArray, catKey, valKey) {
    spcGroupedData = {};
    if (!catKey || catKey === '_ALL_') {
        spcGroupedData["Total_Data"] = []; 
        dataArray.forEach(row => { 
            let val = parseFloat(String(row[valKey]).replace(/[^0-9.\-]/g, '')); 
            if (!isNaN(val)) spcGroupedData["Total_Data"].push(val); 
        });
    } else {
        dataArray.forEach(row => { 
            const cat = row[catKey]; 
            let val = parseFloat(String(row[valKey]).replace(/[^0-9.\-]/g, '')); 
            if (cat !== undefined && !isNaN(val)) { 
                if (!spcGroupedData[cat]) spcGroupedData[cat] = []; 
                spcGroupedData[cat].push(val); 
            } 
        });
    }
    const itemSel = document.getElementById('spc-inherit-item-select');
    itemSel.innerHTML = '<option value="">-- 選擇項目 --</option>';
    Object.keys(spcGroupedData).forEach(k => itemSel.add(new Option(k, k)));
    itemSel.style.display = 'inline-block';
    if (itemSel.options.length > 1) { 
        itemSel.selectedIndex = 1; 
        window.loadSpcItemToTable(); 
    }
}

window.loadSpcItemToTable = function() { 
    const key = document.getElementById('spc-inherit-item-select').value; 
    if (!key) return; 
    fillSpcTableFromValues(spcGroupedData[key]); 
    window.calculateAndDraw(); 
}

function fillSpcTableFromValues(values) {
    let chunked = []; 
    for (let i = 0; i < values.length; i += 5) chunked.push(values.slice(i, i + 5));
    if (chunked.length > CONFIG.numSubgroups) { 
        CONFIG.numSubgroups = Math.min(chunked.length, CONFIG.maxRows); 
        initSpcTable(); 
    } else { 
        for (let i = 1; i <= CONFIG.numSubgroups; i++) 
            for (let j = 1; j <= 5; j++) { 
                const el = document.getElementById(`cell_${i}_${j}`); 
                if (el) el.value = ""; 
            } 
    }
    chunked.forEach((grp, rIdx) => { 
        if (rIdx < CONFIG.numSubgroups) grp.forEach((v, cIdx) => { 
            const el = document.getElementById(`cell_${rIdx + 1}_${cIdx + 1}`); 
            if (el) el.value = v; 
        }); 
    });
}

window.calculateAndDraw = function() {
    try {
        let subgroups = [], allData = [], sumX = 0, sumR = 0, valid = 0;
        for (let i = 1; i <= CONFIG.numSubgroups; i++) {
            let vals = []; 
            for (let j = 1; j <= 5; j++) { 
                const el = document.getElementById(`cell_${i}_${j}`); 
                if (el && el.value !== "") vals.push(parseFloat(el.value)); 
            }
            let xEl = document.getElementById(`res_xbar_${i}`); 
            let rEl = document.getElementById(`res_r_${i}`);
            if (!xEl) continue;
            if (vals.length === 5) {
                let x = vals.reduce((a, b) => a + b, 0) / 5; 
                let r = Math.max(...vals) - Math.min(...vals);
                xEl.innerText = x.toFixed(3); 
                rEl.innerText = r.toFixed(3);
                subgroups.push({ id: i, xbar: x, r: r }); 
                allData.push(...vals); 
                sumX += x; 
                sumR += r; 
                valid++;
            } else { 
                xEl.innerText = "-"; 
                rEl.innerText = "-"; 
                subgroups.push({ id: i, xbar: null, r: null }); 
            }
        }
        if (valid < 2) { 
            spcClearUI(); 
            return; 
        }
        const xdb = sumX / valid; 
        const rb = sumR / valid;
        const uclX = xdb + (CONSTANTS.A2 * rb); 
        const lclX = xdb - (CONSTANTS.A2 * rb);
        const uclR = CONSTANTS.D4 * rb; 
        const lclR = CONSTANTS.D3 * rb;
        const sigmaX = (uclX - xdb) / 3;

        updateSpcStatsUI(xdb, uclX, lclX, rb, uclR, lclR);
        
        let usl = parseFloat(document.getElementById('uslInput').value); 
        let lsl = parseFloat(document.getElementById('lslInput').value);
        if (!isNaN(usl) || !isNaN(lsl)) {
            const si = rb / CONSTANTS.d2; 
            const so = calculateStdDev(allData, xdb);
            updateSpcCapUI(calculateCapability(xdb, si, so, usl, lsl));
        } else { 
            clearSpcCapUI(); 
        }

        const xvals = subgroups.map(g => g.xbar); 
        const violations = checkNelsonRules(xvals, xdb, sigmaX, uclX, lclX);
        displayViolations(violations, subgroups);
        drawSpcChart('xbar', subgroups.map(g => g.id), xvals, uclX, lclX, xdb, 'X-Bar', violations, usl, lsl);
        drawSpcChart('range', subgroups.map(g => g.id), subgroups.map(g => g.r), uclR, lclR, rb, 'Range', [], null, null);
    } catch (e) { 
        console.error(e); 
    }
}

function checkNelsonRules(data, mean, sigma, ucl, lcl) {
    let violations = [];
    
    for (let i = 0; i < data.length; i++) {
        if (data[i] === null) continue;
        let rules = [];
        
        // Rule 1: 1 point > 3 sigma (超出管制界限)
        if (data[i] > ucl || data[i] < lcl) {
            rules.push("規則1 (超出 3σ)");
        }
        
        // Rule 2: 9 points in a row on same side of mean
        if (i >= 8) {
            let side = data[i] > mean ? 1 : -1;
            let rule2 = true;
            for (let j = 0; j < 9; j++) {
                if (data[i - j] === null || (data[i - j] > mean ? 1 : -1) !== side) { rule2 = false; break; }
            }
            if (rule2) rules.push("規則2 (連9點同側)");
        }
        
        // Rule 3: 6 points continually increasing/decreasing
        if (i >= 5) {
            let dec = true, inc = true;
            for (let j = 0; j < 5; j++) {
                if (data[i - j] === null || data[i - j - 1] === null) { inc = false; dec = false; break; }
                if (data[i - j] <= data[i - j - 1]) inc = false;
                if (data[i - j] >= data[i - j - 1]) dec = false;
            }
            if (inc || dec) rules.push("規則3 (連6點遞增/減)");
        }
        
        // Rule 4: 14 points alternating
        if (i >= 13) {
            let rule4 = true;
            for (let j = 0; j < 13; j++) {
                if (data[i - j] === null || data[i - j - 1] === null || data[i - j - 2] === null) { rule4 = false; break; }
                const diff1 = data[i - j] - data[i - j - 1];
                const diff2 = data[i - j - 1] - data[i - j - 2];
                if (diff1 * diff2 >= 0) { rule4 = false; break; }
            }
            if (rule4) rules.push("規則4 (連14點上下交替)");
        }
        
        // Rule 5: 2 out of 3 points > 2 sigma same side
        if (i >= 2) {
            let pts = [data[i], data[i-1], data[i-2]];
            if (!pts.includes(null)) {
                let countPos = pts.filter(p => p - mean > 2 * sigma).length;
                let countNeg = pts.filter(p => mean - p > 2 * sigma).length;
                if (countPos >= 2 || countNeg >= 2) rules.push("規則5 (3點內有2點超出同側 2σ)");
            }
        }
        
        // Rule 6: 4 out of 5 points > 1 sigma same side
        if (i >= 4) {
            let pts = [data[i], data[i-1], data[i-2], data[i-3], data[i-4]];
            if (!pts.includes(null)) {
                let countPos = pts.filter(p => p - mean > 1 * sigma).length;
                let countNeg = pts.filter(p => mean - p > 1 * sigma).length;
                if (countPos >= 4 || countNeg >= 4) rules.push("規則6 (5點內有4點超出同側 1σ)");
            }
        }
        
        // Rule 7: 15 points within 1 sigma
        if (i >= 14) {
            let rule7 = true;
            for (let j = 0; j < 15; j++) {
                if (data[i-j] === null || Math.abs(data[i-j] - mean) >= 1 * sigma) { rule7 = false; break; }
            }
            if (rule7) rules.push("規則7 (連15點分佈於中心線 1σ 內)");
        }
        
        // Rule 8: 8 points > 1 sigma on either side
        if (i >= 7) {
            let rule8 = true;
            for (let j = 0; j < 8; j++) {
                if (data[i-j] === null || Math.abs(data[i-j] - mean) <= 1 * sigma) { rule8 = false; break; }
            }
            if (rule8) rules.push("規則8 (連8點均落在中心線 1σ 外)");
        }

        if (rules.length > 0) violations.push({ index: i, rules: rules });
    }
    return violations;
}

function spcSetText(id, v) { 
    const e = document.getElementById(id); 
    if (e) e.innerText = v; 
}

function updateSpcStatsUI(x, ux, lx, r, ur, lr) { 
    const f = n => isNaN(n) ? '-' : n.toFixed(3); 
    spcSetText('val-xdb', f(x)); 
    spcSetText('val-uclx', f(ux)); 
    spcSetText('val-lclx', f(lx)); 
    spcSetText('val-rb', f(r)); 
    spcSetText('val-uclr', f(ur)); 
    spcSetText('val-lclr', lr <= 0 ? "0.000" : f(lr)); 
}

function spcClearUI() { 
    ['val-xdb', 'val-uclx', 'val-lclx', 'val-rb', 'val-uclr', 'val-lclr'].forEach(id => spcSetText(id, '-')); 
    clearSpcCapUI(); 
}

function updateSpcCapUI(c) { 
    const f = n => (n !== null && !isNaN(n)) ? n.toFixed(3) : '-'; 
    spcSetText('val-cp', f(c.cp)); 
    spcSetText('val-cpk', f(c.cpk)); 
    spcSetText('val-pp', f(c.pp)); 
    spcSetText('val-ppk', f(c.ppk)); 
}

function clearSpcCapUI() { 
    ['val-cp', 'val-cpk', 'val-pp', 'val-ppk'].forEach(id => spcSetText(id, '-')); 
}

function drawSpcChart(type, lbls, data, ucl, lcl, cl, title, vio, usl, lsl) {
    const ctx = document.getElementById(type === 'xbar' ? 'xbarChart' : 'rChart').getContext('2d');
    
    // Background reset
    ctx.save(); 
    ctx.globalCompositeOperation = 'destination-over'; 
    ctx.fillStyle = 'white'; 
    ctx.fillRect(0, 0, ctx.canvas.width, ctx.canvas.height); 
    ctx.restore();
    
    let cols = Array(data.length).fill('#3498db'); 
    let rads = Array(data.length).fill(4);
    if (vio) vio.forEach(v => { cols[v.index] = '#e74c3c'; rads[v.index] = 7; });
    
    let anns = {};
    if (type === 'xbar') {
        if (!isNaN(usl)) anns.usl = { type: 'line', yMin: usl, yMax: usl, borderColor: '#f39c12', borderDash: [4, 4], borderWidth: 2, label: { display: true, content: 'USL', position: 'end' } };
        if (!isNaN(lsl)) anns.lsl = { type: 'line', yMin: lsl, yMax: lsl, borderColor: '#f39c12', borderDash: [4, 4], borderWidth: 2, label: { display: true, content: 'LSL', position: 'end' } };
    }
    
    if (spc_charts[type]) spc_charts[type].destroy();
    
    spc_charts[type] = new Chart(ctx, { 
        type: 'line', 
        data: { 
            labels: lbls, 
            datasets: [
                { label: 'Data', data: data, borderColor: '#3498db', pointBackgroundColor: cols, pointRadius: rads, fill: false, tension: 0 }, 
                { label: 'UCL', data: Array(lbls.length).fill(ucl), borderColor: '#e74c3c', borderDash: [5, 5], pointRadius: 0 }, 
                { label: 'CL', data: Array(lbls.length).fill(cl), borderColor: '#2ecc71', pointRadius: 0 }, 
                { label: 'LCL', data: Array(lbls.length).fill(lcl), borderColor: '#e74c3c', borderDash: [5, 5], pointRadius: 0 }
            ] 
        }, 
        options: { 
            responsive: true, 
            maintainAspectRatio: false, 
            animation: false, 
            plugins: { 
                annotation: { annotations: anns }, 
                legend: { display: true, position: 'top' },
                zoom: { 
                    zoom: { wheel: { enabled: true }, mode: 'x', pinch: { enabled: true } }, 
                    pan: { enabled: true, mode: 'x' } 
                }
            }, 
            scales: { y: { title: { display: true, text: title } } } 
        } 
    });
}

window.clearSpcData = function(redraw = true) { 
    for (let i = 1; i <= CONFIG.numSubgroups; i++) 
        for (let j = 1; j <= 5; j++) { 
            let el = document.getElementById(`cell_${i}_${j}`); 
            if (el) el.value = ""; 
        } 
    if (redraw) window.calculateAndDraw(); 
}

function initSpcTable() { 
    const table = document.getElementById('inputTable'); 
    let html = `<thead><tr><th style="width:35px">n</th>`; 
    for (let j = 1; j <= 5; j++) html += `<th>X${j}</th>`; 
    html += `<th>X̄</th><th>R</th></tr></thead><tbody id="spcTableBody"></tbody>`; 
    table.innerHTML = html; 
    renderSpcRows(1, CONFIG.numSubgroups); 
}

function renderSpcRows(start, end) { 
    const tbody = document.getElementById('spcTableBody'); 
    let html = ''; 
    for (let i = start; i <= end; i++) { 
        html += `<tr><td>${i}</td>`; 
        for (let j = 1; j <= 5; j++) html += `<td><input type="number" id="cell_${i}_${j}" onchange="window.calculateAndDraw()"></td>`; 
        html += `<td class="calc-res" id="res_xbar_${i}">-</td><td class="calc-res" id="res_r_${i}">-</td></tr>`; 
    } 
    if (start === 1) tbody.innerHTML = html; 
    else tbody.insertAdjacentHTML('beforeend', html); 
}

window.addMoreRows = function() { 
    if (CONFIG.numSubgroups >= CONFIG.maxRows) return showError("已達最大列數"); 
    let s = CONFIG.numSubgroups + 1; 
    let e = Math.min(CONFIG.numSubgroups + 25, CONFIG.maxRows); 
    CONFIG.numSubgroups = e; 
    renderSpcRows(s, e); 
}

function initSpcResizer() { 
    const resizer = document.getElementById('dragMe'); 
    const container = document.getElementById('spcMainContainer'); 
    let isResizing = false; 
    resizer.addEventListener('mousedown', () => { isResizing = true; document.body.style.cursor = 'col-resize'; }); 
    document.addEventListener('mousemove', (e) => { 
        if (!isResizing) return; 
        let w = e.clientX; 
        if (w < 300) w = 300; 
        if (w > 800) w = 800; 
        container.style.setProperty('--left-width', `${w}px`); 
    }); 
    document.addEventListener('mouseup', () => { isResizing = false; document.body.style.cursor = 'default'; }); 
}

window.exportSpcToExcel = async function() { 
    if (typeof ExcelJS === 'undefined') return showError("尚未載入 ExcelJS");
    const wb = new ExcelJS.Workbook(); 
    const ws = wb.addWorksheet('SPC'); 
    ws.columns = [{ key: 'id' }, { key: 'x1' }, { key: 'x2' }, { key: 'x3' }, { key: 'x4' }, { key: 'x5' }, { key: 'xb' }, { key: 'r' }]; 
    for (let i = 1; i <= CONFIG.numSubgroups; i++) { 
        let v1 = document.getElementById(`cell_${i}_1`)?.value; 
        if (!v1) continue; 
        ws.addRow({ 
            id: i, 
            x1: parseFloat(v1), 
            x2: parseFloat(document.getElementById(`cell_${i}_2`).value), 
            x3: parseFloat(document.getElementById(`cell_${i}_3`).value), 
            x4: parseFloat(document.getElementById(`cell_${i}_4`).value), 
            x5: parseFloat(document.getElementById(`cell_${i}_5`).value), 
            xb: parseFloat(document.getElementById(`res_xbar_${i}`).innerText), 
            r: parseFloat(document.getElementById(`res_r_${i}`).innerText) 
        }); 
    } 
    const addCanvasToWs = (canvasId, colOff, rowOff) => {
        const canvas = document.getElementById(canvasId);
        if (canvas) {
            const imgId = wb.addImage({ base64: canvas.toDataURL('image/png'), extension: 'png' });
            ws.addImage(imgId, { tl: { col: colOff, row: rowOff }, ext: { width: 500, height: 250 } });
        }
    };
    addCanvasToWs('xbarChart', 10, 1);
    addCanvasToWs('rChart', 10, 16);

    const b = await wb.xlsx.writeBuffer(); 
    saveAs(new Blob([b]), `SPC_Report.xlsx`); 
    showSuccess("SPC 報表匯出成功");
}

function displayViolations(violations, subgroups) { 
    const box = document.getElementById('rule-violations'); 
    box.innerHTML = ""; 
    violations.forEach(v => { 
        let div = document.createElement('div'); 
        div.className = 'violation-item'; 
        div.innerText = `第 ${subgroups[v.index].id} 組: ${v.rules.join(', ')}`; 
        box.appendChild(div); 
    }); 
}

function calculateStdDev(data, mean) { 
    if (!data || data.length < 2) return 0; 
    if (mean === undefined) { mean = data.reduce((a, b) => a + b, 0) / data.length; } 
    const variance = data.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / (data.length - 1); 
    return Math.sqrt(variance); 
}

function calculateCapability(mean, sigmaWithin, sigmaOverall, usl, lsl) { 
    let cp = NaN, cpk = NaN, pp = NaN, ppk = NaN; 
    if (!isNaN(usl) && !isNaN(lsl)) { 
        if (sigmaWithin > 0) cp = (usl - lsl) / (6 * sigmaWithin); 
        if (sigmaOverall > 0) pp = (usl - lsl) / (6 * sigmaOverall); 
    } 
    const calcK = (sigma) => { 
        if (sigma === 0) return NaN; 
        let k_u = NaN, k_l = NaN; 
        if (!isNaN(usl)) k_u = (usl - mean) / (3 * sigma); 
        if (!isNaN(lsl)) k_l = (mean - lsl) / (3 * sigma); 
        if (!isNaN(k_u) && !isNaN(k_l)) return Math.min(k_u, k_l); 
        if (!isNaN(k_u)) return k_u; 
        if (!isNaN(k_l)) return k_l; 
        return NaN; 
    }; 
    cpk = calcK(sigmaWithin); 
    ppk = calcK(sigmaOverall); 
    return { cp, cpk, pp, ppk }; 
}

window.parseSpcInheritedData = parseSpcInheritedData;
