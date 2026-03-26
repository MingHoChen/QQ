import { SharedData, runWorkerTask, showError, showSuccess } from '../utils.js';

let grr_charts = {};

export function initGrrEvents() {
    const sourceMode = document.getElementById('grr-source-mode');
    if (sourceMode) sourceMode.addEventListener('change', window.handleGrrSourceChange);
    
    const fileInput = document.getElementById('grr-file-input');
    if (fileInput) fileInput.addEventListener('change', handleGrrFileSelect);
}

window.handleGrrSourceChange = function() {
    const mode = document.getElementById('grr-source-mode').value;
    document.getElementById('grr-file-upload').style.display = mode === 'new' ? 'block' : 'none';
    if (mode === 'cpk') {
        if (SharedData.hasData) populateGrrSelectors(SharedData.headers);
        else showError("請先在 Cpk 頁面上傳數據或選擇上傳新檔案");
    }
}

async function handleGrrFileSelect(e) {
    const file = e.target.files[0];
    if (!file) return;
    try {
        const buffer = await file.arrayBuffer();
        const res = await runWorkerTask('parseFirstSheet', { buffer });
        if (res.headers.length > 0) {
            SharedData.tempGrr = res;
            populateGrrSelectors(res.headers);
            showSuccess("Gage R&R 數據已讀取");
        }
    } catch (err) {
        showError("讀取失敗: " + err.message);
    }
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
    if (headers.length >= 3) {
        partSel.selectedIndex = 1;
        measSel.selectedIndex = 2;
    }
}

window.calculateGrr = function() {
    const mode = document.getElementById('grr-source-mode').value;
    let data = [];
    if (mode === 'cpk') data = SharedData.rawJson;
    else if (SharedData.tempGrr) data = SharedData.tempGrr.data;

    if (!data || data.length === 0) return showError("無數據");

    const opKey = document.getElementById('grr-operator-col').value;
    const partKey = document.getElementById('grr-part-col').value;
    const measKey = document.getElementById('grr-measurement-col').value;

    if (!opKey || !partKey || !measKey) return showError("請選擇所有必要的欄位");

    let cleanData = [];
    data.forEach(row => {
        let op = row[opKey];
        let part = row[partKey];
        let val = parseFloat(String(row[measKey]).replace(/[^0-9.\-]/g, ''));
        if (op !== undefined && part !== undefined && !isNaN(val)) {
            cleanData.push({ op: String(op), part: String(part), val: val });
        }
    });

    if (cleanData.length < 10) return showError("數據太少，無法進行有效分析");

    try {
        const result = computeAnovaGrr(cleanData);
        renderGrrResults(result);
    } catch (e) {
        console.error(e);
        showError("計算發生錯誤: " + e.message + "\n請確認數據格式是否正確 (需有多個作業員、多個部件、重複量測)");
    }
}

function getControlChartConstants(n) {
    if (n === 2) return { A2: 1.880, D3: 0, D4: 3.267 };
    if (n === 3) return { A2: 1.023, D3: 0, D4: 2.574 };
    if (n === 4) return { A2: 0.729, D3: 0, D4: 2.282 };
    if (n === 5) return { A2: 0.577, D3: 0, D4: 2.114 };
    if (n === 6) return { A2: 0.483, D3: 0, D4: 2.004 };
    if (n === 7) return { A2: 0.419, D3: 0.076, D4: 1.924 };
    if (n === 8) return { A2: 0.373, D3: 0.136, D4: 1.864 };
    if (n === 9) return { A2: 0.337, D3: 0.184, D4: 1.816 };
    if (n >= 10) return { A2: 0.308, D3: 0.223, D4: 1.777 };
    return { A2: 0.577, D3: 0, D4: 2.114 }; 
}

function computeAnovaGrr(data) {
    const operators = [...new Set(data.map(d => d.op))].sort();
    const parts = [...new Set(data.map(d => d.part))].sort();
    const a = operators.length; 
    const b = parts.length;     
    
    const totalCount = data.length;
    const n = Math.round(totalCount / (a * b)); 

    if (n < 2) throw new Error("每個部件每位作業員至少需要 2 次量測 (n >= 2)");

    let grandSum = 0;
    let sumSqTotal = 0;
    let sumOp = {};
    let sumPart = {};
    let sumOpPart = {}; 
    let subgroups = {}; 

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

    let sumSqOp = 0; for (let k in sumOp) sumSqOp += sumOp[k] * sumOp[k];
    const SS_Op = (sumSqOp / (b * n)) - CF;

    let sumSqPart = 0; for (let k in sumPart) sumSqPart += sumPart[k] * sumPart[k];
    const SS_Part = (sumSqPart / (a * n)) - CF;

    let sumSqOpPart = 0; for (let k in sumOpPart) sumSqOpPart += sumOpPart[k] * sumOpPart[k];
    const SS_Subtotal = sumSqOpPart / n; 
    const SS_Interaction = SS_Subtotal - SS_Op - SS_Part - CF;

    const SS_Error = SS_Total - SS_Op - SS_Part - SS_Interaction;

    const df_Op = a - 1;
    const df_Part = b - 1;
    const df_Int = (a - 1) * (b - 1);
    const df_Error = a * b * (n - 1);
    const df_Total = (a * b * n) - 1;

    const MS_Op = SS_Op / df_Op;
    const MS_Part = SS_Part / df_Part;
    const MS_Int = SS_Interaction / df_Int;
    const MS_Error = SS_Error / df_Error;

    const F_Op = MS_Op / MS_Int;
    const F_Part = MS_Part / MS_Int;
    const F_Int = MS_Int / MS_Error;

    let var_Equipment = MS_Error;
    let var_Interaction = (MS_Int - MS_Error) / n;
    if (var_Interaction < 0) var_Interaction = 0;

    let var_Operator = (MS_Op - MS_Int) / (b * n);
    if (var_Operator < 0) var_Operator = 0;

    let var_Part = (MS_Part - MS_Int) / (a * n);
    if (var_Part < 0) var_Part = 0;

    const var_GRR = var_Equipment + var_Operator + var_Interaction;
    const var_Total = var_GRR + var_Part;

    const std_Total = Math.sqrt(var_Total);
    const stats = {
        EV: { var: var_Equipment, std: Math.sqrt(var_Equipment) },
        AV: { var: var_Operator + var_Interaction, std: Math.sqrt(var_Operator + var_Interaction) },
        GRR: { var: var_GRR, std: Math.sqrt(var_GRR) },
        PV: { var: var_Part, std: Math.sqrt(var_Part) },
        TV: { var: var_Total, std: std_Total }
    };

    const report = [
        { name: "Repeatability (EV)", var: stats.EV.var, pctCont: Math.max(0, (stats.EV.var/var_Total)*100), pctStudy: Math.max(0, (stats.EV.std/std_Total)*100) },
        { name: "Reproducibility (AV)", var: stats.AV.var, pctCont: Math.max(0, (stats.AV.var/var_Total)*100), pctStudy: Math.max(0, (stats.AV.std/std_Total)*100) },
        { name: "Interaction", var: var_Interaction, pctCont: Math.max(0, (var_Interaction/var_Total)*100), pctStudy: Math.max(0, (Math.sqrt(var_Interaction)/std_Total)*100) },
        { name: "Gage R&R (GRR)", var: stats.GRR.var, pctCont: Math.max(0, (stats.GRR.var/var_Total)*100), pctStudy: Math.max(0, (stats.GRR.std/std_Total)*100) },
        { name: "Part Variation (PV)", var: stats.PV.var, pctCont: Math.max(0, (stats.PV.var/var_Total)*100), pctStudy: Math.max(0, (stats.PV.std/std_Total)*100) },
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

    drawGrrSixPack(result);
}

function drawGrrSixPack(result) {
    ['grr-chart-components', 'grr-chart-r', 'grr-chart-xbar', 'grr-chart-bypart', 'grr-chart-byop', 'grr-chart-interaction'].forEach(id => {
        if (grr_charts[id]) grr_charts[id].destroy();
    });

    const colors = ['#3498db', '#e74c3c', '#2ecc71', '#f1c40f', '#9b59b6', '#34495e'];

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

    const consts = getControlChartConstants(result.meta.n);
    const subgroups = [];
    const xbars = [];
    const ranges = [];
    
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

    const partGroups = {};
    result.meta.parts.forEach(p => partGroups[p] = []);
    result.rawData.forEach(d => partGroups[d.part].push(d.val));
    
    const byPartLabels = result.meta.parts;
    const byPartMeans = byPartLabels.map(p => partGroups[p].reduce((a,b)=>a+b,0)/partGroups[p].length);
    
    const byPartScatterData = [];
    result.rawData.forEach(d => { byPartScatterData.push({ x: d.part, y: d.val }); });

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

    const opGroups = {};
    result.meta.operators.forEach(o => opGroups[o] = []);
    result.rawData.forEach(d => opGroups[d.op].push(d.val));
    
    const byOpLabels = result.meta.operators;
    const byOpMeans = byOpLabels.map(o => opGroups[o].reduce((a,b)=>a+b,0)/opGroups[o].length);
    
    const byOpScatterData = [];
    result.rawData.forEach(d => { byOpScatterData.push({ x: d.op, y: d.val }); });

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

    const interactionDatasets = result.meta.operators.map((op, idx) => {
        const data = result.meta.parts.map(p => {
            const key = op + "|" + p;
            const vals = result.subgroups[key];
            return vals ? vals.reduce((a,b)=>a+b,0)/vals.length : null;
        });
        return { label: op, data: data, borderColor: colors[idx % colors.length], tension: 0, fill: false };
    });

    grr_charts['grr-chart-interaction'] = new Chart(document.getElementById('grr-chart-interaction'), {
        type: 'line',
        data: { labels: result.meta.parts, datasets: interactionDatasets },
        options: { responsive: true, maintainAspectRatio: false }
    });
}
