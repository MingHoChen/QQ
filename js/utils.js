export const SharedData = { rawJson: [], headers: [], hasData: false, tempPareto: null, tempScatter: null, tempBoxplot: null, tempGrr: null };
export const SharedSpecs = { lsl: NaN, usl: NaN };
export let CONFIG = { numSubgroups: 50, maxRows: 200, sampleSize: 5, decimalPlaces: 3 };
export const CONSTANTS = { n: 5, A2: 0.577, D3: 0, D4: 2.114, d2: 2.326 };

let globalWorker = null;
let workerCallbacks = {};

export function getWorker() {
    if (!globalWorker) {
        globalWorker = new Worker('js/worker.js');
        globalWorker.onmessage = (e) => {
            const data = e.data;
            if (data.id && workerCallbacks[data.id]) {
                if (data.action === 'error') {
                    workerCallbacks[data.id].reject(new Error(data.message));
                } else {
                    workerCallbacks[data.id].resolve(data);
                }
                delete workerCallbacks[data.id];
            }
        };
        globalWorker.onerror = (e) => {
            console.error("Worker error:", e);
        };
    }
    return globalWorker;
}

export function runWorkerTask(action, payload) {
    return new Promise((resolve, reject) => {
        const worker = getWorker();
        const id = Math.random().toString(36).substr(2, 9);
        workerCallbacks[id] = { resolve, reject };
        worker.postMessage({ id, action, payload });
    });
}

export function showError(msg) {
    if (typeof Swal !== 'undefined') {
        Swal.fire({ icon: 'error', title: '錯誤', text: msg, confirmButtonText: '確定' });
    } else {
        alert(msg);
    }
}

export function showSuccess(msg) {
    if (typeof Swal !== 'undefined') {
        Swal.fire({ icon: 'success', title: '成功', text: msg, timer: 1500, showConfirmButton: false });
    } else {
        alert(msg);
    }
}

export function downloadChart(id) { 
    const link = document.createElement('a'); 
    link.download = id + ".png"; 
    link.href = document.getElementById(id).toDataURL(); 
    link.click(); 
}
