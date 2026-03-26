import { initCpkEvents } from './modules/cpk.js';
import { initSpcEvents } from './modules/spc.js';
import { initParetoEvents } from './modules/pareto.js';
import { initScatterEvents } from './modules/scatter.js';
import { initBoxplotEvents } from './modules/boxplot.js';
import { initGrrEvents } from './modules/grr.js';
import { SharedSpecs, showSuccess, showError } from './utils.js';

window.onload = function () {
    initCpkEvents();
    initSpcEvents();
    initParetoEvents();
    initScatterEvents();
    initBoxplotEvents();
    initGrrEvents();
};

window.switchTab = function (tabId) {
    document.querySelectorAll('.nav-tab').forEach(btn => btn.classList.remove('active'));
    event.currentTarget.classList.add('active');
    document.querySelectorAll('.tab-content').forEach(div => div.classList.remove('active'));
    document.getElementById('view-' + tabId).classList.add('active');

    if (tabId === 'spc') {
        if (!isNaN(SharedSpecs.lsl)) document.getElementById('lslInput').value = SharedSpecs.lsl;
        if (!isNaN(SharedSpecs.usl)) document.getElementById('uslInput').value = SharedSpecs.usl;
        if (window.handleSpcSourceChange) window.handleSpcSourceChange();
        setTimeout(() => {
            if (window.calculateAndDraw) window.calculateAndDraw();
        }, 100);
    } else if (tabId === 'pareto') {
        if (window.handleParetoSourceChange) window.handleParetoSourceChange();
    } else if (tabId === 'scatter') {
        if (window.handleScatterSourceChange) window.handleScatterSourceChange();
    } else if (tabId === 'boxplot') {
        if (window.handleBoxplotSourceChange) window.handleBoxplotSourceChange();
    } else if (tabId === 'grr') {
        if (window.handleGrrSourceChange) window.handleGrrSourceChange();
    }
}


