importScripts('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js');

let loadedWorkbooks = [];

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

self.onmessage = function(e) {
    const { action, payload, id } = e.data;
    
    try {
        if (action === 'readWorkbooks') {
            const files = payload.files; // [{ name, buffer }]
            loadedWorkbooks = [];
            const results = [];
            
            files.forEach((f, idx) => {
                const wb = XLSX.read(new Uint8Array(f.buffer), { type: 'array' });
                loadedWorkbooks.push({ name: f.name, wb: wb });
                results.push({ name: f.name, sheetNames: wb.SheetNames, index: idx });
            });
            self.postMessage({ id, action: 'readWorkbooksDone', results });
            
        } else if (action === 'processSheets') {
            const selections = payload.selections; // [{ fIdx, sheetName }]
            let combinedHeaders = [];
            let combinedData = [];
            let successCount = 0;
            
            selections.forEach(sel => {
                const wb = loadedWorkbooks[sel.fIdx].wb;
                const sheet = wb.Sheets[sel.sheetName];
                const res = robustParseExcel(sheet);
                if (res.headers.length > 0) {
                    if (combinedHeaders.length === 0) combinedHeaders = res.headers;
                    combinedData = combinedData.concat(res.data);
                    successCount++;
                }
            });
            self.postMessage({ id, action: 'processSheetsDone', headers: combinedHeaders, data: combinedData, successCount });
            
        } else if (action === 'parseFirstSheet') {
            const buffer = payload.buffer;
            const wb = XLSX.read(new Uint8Array(buffer), { type: 'array' });
            const sheet = wb.Sheets[wb.SheetNames[0]];
            const res = robustParseExcel(sheet);
            self.postMessage({ id, action: 'parseFirstSheetDone', headers: res.headers, data: res.data });
        }
    } catch (err) {
        self.postMessage({ id, action: 'error', message: err.message });
    }
};
