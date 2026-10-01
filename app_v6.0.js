// ============================================================
//  V4.4 DEPOT IMPORT ENGINE (CSV, XLSX, PDF, JSON) & EMPFEHLUNGSLISTE
// ============================================================

const LIBRARY_KEY = 'depotBibliothek_V50';
const EMPFEHLUNG_KEY_PREFIX = 'empfehlungsliste';
function empfehlungsKey(blockId) { return `${blockId}::${EMPFEHLUNG_KEY_PREFIX}`; }

function anlageschwerpunktToBlock(schwerpunkt) {
    if (!schwerpunkt) return null;
    const s = schwerpunkt.toLowerCase().trim();
    if (s.includes('geldmarkt')) return 'block-kasse';
    if (s.includes('anleihen euro kurz') || s.includes('anleihen euro kurz laufzeit') ||
        s.includes('anleihen hochzins laufzeit')) return 'block-kasse';
    if (s.includes('anleihen') || s.includes('renten') || s.includes('bond')) return 'block-defensiv';
    if (s.includes('vermögensverwalter - defensiv') || (s.includes('verm') && s.includes('defensiv'))) return 'block-defensiv';
    if (s.includes('vermögensverwalter - ausgewogen') || (s.includes('verm') && s.includes('ausgewogen'))) return 'block-ausgewogen';
    if (s.includes('vermögensverwalter - dynamisch') || (s.includes('verm') && s.includes('dynamisch'))) return 'block-dynamisch';
    if (s.includes('alternative volatilitätsstrategien') || s.includes('alternative') || s.includes('spezial')) return 'block-spezial';
    if (s.includes('aktien weit') || s.includes('weites benchmarking')) return 'block-maerkte-weit';
    if (s.includes('aktien eng') || s.includes('enges benchmarking')) return 'block-maerkte-eng';
    if (s.includes('aktien')) return 'block-maerkte-weit';
    return null;
}

function parseGermanNumber(s) {
    if (!s) return 0;
    const cleaned = String(s).replace(/\./g, '').replace(',', '.').replace(/[^0-9.]/g, '');
    const v = parseFloat(cleaned);
    return isNaN(v) ? 0 : v;
}

function normalizeWKN(wkn) {
    return String(wkn).toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function buildWknLookup() {
    const map = {};
    if (typeof managementBlocks === 'undefined') return map;
    managementBlocks.forEach(block => {
        block.funds.forEach(fund => {
            if (fund._isEmpfehlungsliste) return;
            const matches = fund.info.matchAll(/WKN:\s*([A-Z0-9]{6})/gi);
            for (const m of matches) {
                const wkn = normalizeWKN(m[1]);
                if (!map[wkn]) map[wkn] = [];
                map[wkn].push({ block, fund });
            }
        });
    });
    const blockOrder = managementBlocks.map(b => b.id);
    Object.keys(map).forEach(wkn => {
        if (map[wkn].length <= 1) return;
        const nonTagesgeld = map[wkn].filter(e => e.block.id !== 'block-tagesgeld');
        const candidates = nonTagesgeld.length > 0 ? nonTagesgeld : map[wkn];
        candidates.sort((a, b) => blockOrder.indexOf(a.block.id) - blockOrder.indexOf(b.block.id));
        map[wkn] = [candidates[0]];
    });
    return map;
}

function parseCsvDepot(text, wknLookup) {
    const matched = [], unmatched = [];
    function cleanAmount(s) { return s.replace(/[^\d.,]/g, '').trim(); }
    const clean = text.split(/\r?\n/);

    for (let i = 0; i < clean.length; i++) {
        const cols = clean[i].split(';');
        const wknRaw = (cols[0] || '').trim();
        const wkn = normalizeWKN(wknRaw);
        if (wkn.length !== 6) continue;

        let schwerpunkt = (cols[2] || '').trim();
        let betragRaw = cleanAmount(cols[4] || cols[3] || '');
        let sparrateRaw = cleanAmount(cols[6] || cols[5] || '');
        let einmal = parseGermanNumber(betragRaw);
        let sparrate = parseGermanNumber(sparrateRaw);

        if (einmal <= 0 && sparrate <= 0 && i > 0) {
            const prevCols = clean[i - 1].split(';');
            const prevWknCheck = normalizeWKN((prevCols[0] || '').trim());
            if (prevWknCheck.length !== 6) {
                const prevBetragRaw = cleanAmount(prevCols[4] || prevCols[3] || '');
                const prevSparRaw = cleanAmount(prevCols[6] || prevCols[5] || '');
                einmal = parseGermanNumber(prevBetragRaw);
                sparrate = parseGermanNumber(prevSparRaw);
                if (!schwerpunkt) schwerpunkt = (prevCols[2] || '').trim();
            }
        }
        if (einmal <= 0 && sparrate <= 0) continue;
        const nameRaw = i + 1 < clean.length ? (clean[i + 1].split(';')[0] || '').trim() : '';

        if (wknLookup[wkn]) {
            wknLookup[wkn].forEach(entry => {
                if (!matched.some(m => m.block.id === entry.block.id && m.fund.name === entry.fund.name))
                    matched.push({ block: entry.block, fund: entry.fund, einmal, sparrate, wkn, schwerpunkt });
            });
        } else {
            if (!unmatched.some(u => u.wkn === wkn))
                unmatched.push({ name: nameRaw || `Unbekannter Fonds (${wkn})`, wkn, einmal, sparrate, schwerpunkt });
        }
    }
    return { matched, unmatched };
}

function parseXlsxDepot(workbook, wknLookup) {
    const matched = [], unmatched = [];
    if (typeof XLSX === 'undefined') return { matched, unmatched };
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    function getAmt(row) {
        for (const idx of [4, 3]) {
            const raw = String(row[idx] === null || row[idx] === undefined ? '' : row[idx]).replace(/\u00a0/g, '').replace(/€/g, '').trim();
            const v = parseGermanNumber(raw);
            if (v > 0) return v;
        }
        return 0;
    }

    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const rawCell = row[0];
        const wkn = normalizeWKN(String(rawCell === null || rawCell === undefined ? '' : rawCell).trim());
        if (wkn.length !== 6) continue;

        let schwerpunkt = String(row[2] || '').trim();
        let einmal = getAmt(row);
        let sparrate = 0;

        if (einmal <= 0 && i > 0) {
            const prev = rows[i - 1];
            const prevWkn = normalizeWKN(String(prev[0] === null || prev[0] === undefined ? '' : prev[0]).trim());
            if (prevWkn.length !== 6) {
                einmal = getAmt(prev);
                if (!schwerpunkt) schwerpunkt = String(prev[2] || '').trim();
            }
        }
        if (einmal <= 0 && sparrate <= 0) continue;
        const nameRaw = i + 1 < rows.length ? String(rows[i + 1][0] === null || rows[i + 1][0] === undefined ? '' : rows[i + 1][0]).trim() : '';

        if (wknLookup[wkn]) {
            wknLookup[wkn].forEach(entry => {
                if (!matched.some(m => m.block.id === entry.block.id && m.fund.name === entry.fund.name))
                    matched.push({ block: entry.block, fund: entry.fund, einmal, sparrate, wkn, schwerpunkt });
            });
        } else {
            if (!unmatched.some(u => u.wkn === wkn))
                unmatched.push({ name: nameRaw || `Unbekannter Fonds (${wkn})`, wkn, einmal, sparrate, schwerpunkt });
        }
    }
    return { matched, unmatched };
}

if (typeof pdfjsLib !== 'undefined')
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

async function extractPdfText(file) {
    if (typeof pdfjsLib === 'undefined') throw new Error('PDF.js nicht geladen.');
    const pdf = await pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
    let fullText = '';
    for (let i = 1; i <= pdf.numPages; i++) {
        const content = await (await pdf.getPage(i)).getTextContent();
        const items = content.items.sort((a,b) => {
            const ay = Math.round(a.transform[5]*10), by = Math.round(b.transform[5]*10);
            return ay !== by ? by - ay : a.transform[4] - b.transform[4];
        });
        fullText += items.map(it => it.str).join(' ') + '\n';
    }
    return fullText;
}

function parsePdfData(text, wknLookup) {
    const matched = [], unmatched = [];
    const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
    const wknPat = /\b([A-Z0-9]{6})\b/g;
    const amtPat = /(\d{1,3}(?:\.\d{3})*,\d{2})/g;
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        const wkns = []; let wm;
        wknPat.lastIndex = 0;
        while ((wm = wknPat.exec(line)) !== null) wkns.push(wm[1]);
        const ctx = [line, lines[i+1]||'', lines[i+2]||''].join(' ');
        const amts = []; let am;
        amtPat.lastIndex = 0;
        while ((am = amtPat.exec(ctx)) !== null) amts.push(parseGermanNumber(am[1]));
        for (const wknRaw of wkns) {
            const wkn = normalizeWKN(wknRaw);
            if (wkn.length !== 6) continue;
            if (wknLookup[wkn]) {
                if ((amts[0]||0) <= 0) continue;
                wknLookup[wkn].forEach(entry => {
                    if (!matched.some(m => m.block.id === entry.block.id && m.fund.name === entry.fund.name))
                        matched.push({ block: entry.block, fund: entry.fund, einmal: amts[0], sparrate: amts[1]||0, wkn, schwerpunkt: '' });
                });
            } else if ((amts[0]||0) > 0) {
                const prev = lines.slice(Math.max(0,i-3),i).filter(l => !/^\d/.test(l) && l.length > 5);
                if (!unmatched.some(u => u.wkn === wkn))
                    unmatched.push({ name: prev.at(-1) || `(${wkn})`, wkn, einmal: amts[0], sparrate: amts[1]||0, schwerpunkt: '' });
            }
        }
    }
    return { matched, unmatched };
}

function categorizeUnmatched(unmatched) {
    const mapped = [];
    const ambiguous = [];
    unmatched.forEach(u => {
        const blockId = anlageschwerpunktToBlock(u.schwerpunkt);
        if (blockId) mapped.push({ ...u, blockId });
        else ambiguous.push(u);
    });
    return { mapped, ambiguous };
}

function applyImportData(matched, mappedUnmatched) {
    portfolioGlobals.fundInvestments = {};
    portfolioGlobals.fundSparrates = {};
    portfolio = [];
    localStorage.removeItem('empfehlungslisteFonds_V44');
    localStorage.removeItem('delistedFunds_V44');

    matched.forEach(({ block, fund, einmal, sparrate }) => {
        const key = `${block.id}::${fund.name}`;
        if (einmal > 0) { portfolioGlobals.fundInvestments[key] = einmal; addFund(block, fund); }
        if (sparrate > 0) { portfolioGlobals.fundSparrates[key] = sparrate; addFund(block, fund); }
    });

    if (mappedUnmatched && mappedUnmatched.length > 0) {
        const empFondsMap = {};
        try { Object.assign(empFondsMap, JSON.parse(localStorage.getItem('empfehlungslisteFonds_V44') || '{}')); } catch {}
        mappedUnmatched.forEach(u => {
            const { blockId, einmal, sparrate } = u;
            const eKey = empfehlungsKey(blockId);
            portfolioGlobals.fundInvestments[eKey] = (portfolioGlobals.fundInvestments[eKey] || 0) + einmal;
            if (sparrate > 0) portfolioGlobals.fundSparrates[eKey] = (portfolioGlobals.fundSparrates[eKey] || 0) + sparrate;
            if (!empFondsMap[blockId]) empFondsMap[blockId] = [];
            if (!empFondsMap[blockId].some(f => f.wkn === u.wkn))
                empFondsMap[blockId].push({ name: u.name, wkn: u.wkn, schwerpunkt: u.schwerpunkt, einmal: u.einmal, sparrate: u.sparrate || 0 });
            
            const block = managementBlocks.find(b => b.id === blockId);
            if (block) {
                const layer = getOrCreateLayer(block.id, block.title);
                if (!layer.funds.some(f => f._isEmpfehlungsliste))
                    layer.funds.push({ name: EMPFEHLUNG_KEY_PREFIX, info: '', type: '', ertrag: '', _isEmpfehlungsliste: true });
            }
        });
        localStorage.setItem('empfehlungslisteFonds_V44', JSON.stringify(empFondsMap));
    }

    const totalEinmalGesamt = Math.round(Object.values(portfolioGlobals.fundInvestments).reduce((s, v) => s + v, 0) * 100) / 100;
    portfolioGlobals.totalInvestment = totalEinmalGesamt;
    portfolioGlobals.initialInvestments = { ...portfolioGlobals.fundInvestments };
    portfolioGlobals.initialSparrates = { ...portfolioGlobals.fundSparrates };
    portfolioGlobals.initialTotalInvestment = totalEinmalGesamt;
    const totInp = document.getElementById('total-investment-input');
    if (totInp) totInp.value = formatNumberInput(totalEinmalGesamt);
    saveGlobals(); savePortfolio();
}

// ============================================================
//  V4.4 DEPOT IMPORT ENGINE (CSV, XLSX, PDF, JSON) & EMPFEHLUNGSLISTE
// ============================================================


function anlageschwerpunktToBlock(schwerpunkt) {
    if (!schwerpunkt) return null;
    const s = schwerpunkt.toLowerCase().trim();
    if (s.includes('geldmarkt')) return 'block-kasse';
    if (s.includes('anleihen euro kurz') || s.includes('anleihen euro kurz laufzeit') ||
        s.includes('anleihen hochzins laufzeit')) return 'block-kasse';
    if (s.includes('anleihen') || s.includes('renten') || s.includes('bond')) return 'block-defensiv';
    if (s.includes('vermögensverwalter - defensiv') || (s.includes('verm') && s.includes('defensiv'))) return 'block-defensiv';
    if (s.includes('vermögensverwalter - ausgewogen') || (s.includes('verm') && s.includes('ausgewogen'))) return 'block-ausgewogen';
    if (s.includes('vermögensverwalter - dynamisch') || (s.includes('verm') && s.includes('dynamisch'))) return 'block-dynamisch';
    if (s.includes('alternative volatilitätsstrategien') || s.includes('alternative') || s.includes('spezial')) return 'block-spezial';
    if (s.includes('aktien weit') || s.includes('weites benchmarking')) return 'block-maerkte-weit';
    if (s.includes('aktien eng') || s.includes('enges benchmarking')) return 'block-maerkte-eng';
    if (s.includes('aktien')) return 'block-maerkte-weit';
    return null;
}

function parseGermanNumber(s) {
    if (!s) return 0;
    const cleaned = String(s).replace(/\./g, '').replace(',', '.').replace(/[^0-9.]/g, '');
    const v = parseFloat(cleaned);
    return isNaN(v) ? 0 : v;
}

function normalizeWKN(wkn) {
    return String(wkn).toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function buildWknLookup() {
    const map = {};
    if (typeof managementBlocks === 'undefined') return map;
    managementBlocks.forEach(block => {
        block.funds.forEach(fund => {
            if (fund._isEmpfehlungsliste) return;
            const matches = fund.info.matchAll(/WKN:\s*([A-Z0-9]{6})/gi);
            for (const m of matches) {
                const wkn = normalizeWKN(m[1]);
                if (!map[wkn]) map[wkn] = [];
                map[wkn].push({ block, fund });
            }
        });
    });
    const blockOrder = managementBlocks.map(b => b.id);
    Object.keys(map).forEach(wkn => {
        if (map[wkn].length <= 1) return;
        const nonTagesgeld = map[wkn].filter(e => e.block.id !== 'block-tagesgeld');
        const candidates = nonTagesgeld.length > 0 ? nonTagesgeld : map[wkn];
        candidates.sort((a, b) => blockOrder.indexOf(a.block.id) - blockOrder.indexOf(b.block.id));
        map[wkn] = [candidates[0]];
    });
    return map;
}

function parseCsvDepot(text, wknLookup) {
    const matched = [], unmatched = [];
    function cleanAmount(s) { return s.replace(/[^\d.,]/g, '').trim(); }
    const clean = text.split(/\r?\n/);

    for (let i = 0; i < clean.length; i++) {
        const cols = clean[i].split(';');
        const wknRaw = (cols[0] || '').trim();
        const wkn = normalizeWKN(wknRaw);
        if (wkn.length !== 6) continue;

        let schwerpunkt = (cols[2] || '').trim();
        let betragRaw = cleanAmount(cols[4] || cols[3] || '');
        let sparrateRaw = cleanAmount(cols[6] || cols[5] || '');
        let einmal = parseGermanNumber(betragRaw);
        let sparrate = parseGermanNumber(sparrateRaw);

        if (einmal <= 0 && sparrate <= 0 && i > 0) {
            const prevCols = clean[i - 1].split(';');
            const prevWknCheck = normalizeWKN((prevCols[0] || '').trim());
            if (prevWknCheck.length !== 6) {
                const prevBetragRaw = cleanAmount(prevCols[4] || prevCols[3] || '');
                const prevSparRaw = cleanAmount(prevCols[6] || prevCols[5] || '');
                einmal = parseGermanNumber(prevBetragRaw);
                sparrate = parseGermanNumber(prevSparRaw);
                if (!schwerpunkt) schwerpunkt = (prevCols[2] || '').trim();
            }
        }
        if (einmal <= 0 && sparrate <= 0) continue;
        const nameRaw = i + 1 < clean.length ? (clean[i + 1].split(';')[0] || '').trim() : '';

        if (wknLookup[wkn]) {
            wknLookup[wkn].forEach(entry => {
                if (!matched.some(m => m.block.id === entry.block.id && m.fund.name === entry.fund.name))
                    matched.push({ block: entry.block, fund: entry.fund, einmal, sparrate, wkn, schwerpunkt });
            });
        } else {
            if (!unmatched.some(u => u.wkn === wkn))
                unmatched.push({ name: nameRaw || `Unbekannter Fonds (${wkn})`, wkn, einmal, sparrate, schwerpunkt });
        }
    }
    return { matched, unmatched };
}

function parseXlsxDepot(workbook, wknLookup) {
    const matched = [], unmatched = [];
    if (typeof XLSX === 'undefined') return { matched, unmatched };
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    function getAmt(row) {
        for (const idx of [4, 3, 2, 5]) {
            const rawCell = row[idx];
            if (rawCell === null || rawCell === undefined) continue;
            const raw = String(rawCell).replace(/ /g, '').trim();
            if (!raw) continue;
            if (/^\d{6}$/.test(raw) || /^[A-Z0-9]{6}$/i.test(raw)) continue;
            const cleanStr = raw.replace(/€/g, '').trim();
            const v = parseGermanNumber(cleanStr);
            if (v > 0 && v < 5000000) return v;
        }
        return 0;
    }

    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const rawCell = row[0];
        const wkn = normalizeWKN(String(rawCell === null || rawCell === undefined ? '' : rawCell).trim());
        if (wkn.length !== 6) continue;

        let schwerpunkt = String(row[2] || '').trim();
        let einmal = getAmt(row);
        let sparrate = 0;

        if (einmal <= 0 && i > 0) {
            const prev = rows[i - 1];
            const prevWkn = normalizeWKN(String(prev[0] === null || prev[0] === undefined ? '' : prev[0]).trim());
            if (prevWkn.length !== 6) {
                einmal = getAmt(prev);
                if (!schwerpunkt) schwerpunkt = String(prev[2] || '').trim();
            }
        }
        if (einmal <= 0 && sparrate <= 0) continue;
        const nameRaw = i + 1 < rows.length ? String(rows[i + 1][0] === null || rows[i + 1][0] === undefined ? '' : rows[i + 1][0]).trim() : '';

        if (wknLookup[wkn]) {
            wknLookup[wkn].forEach(entry => {
                if (!matched.some(m => m.block.id === entry.block.id && m.fund.name === entry.fund.name))
                    matched.push({ block: entry.block, fund: entry.fund, einmal, sparrate, wkn, schwerpunkt });
            });
        } else {
            if (!unmatched.some(u => u.wkn === wkn))
                unmatched.push({ name: nameRaw || `Unbekannter Fonds (${wkn})`, wkn, einmal, sparrate, schwerpunkt });
        }
    }
    return { matched, unmatched };
}

if (typeof pdfjsLib !== 'undefined')
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

async function extractPdfText(file) {
    if (typeof pdfjsLib === 'undefined') throw new Error('PDF.js nicht geladen.');
    const pdf = await pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
    let fullText = '';
    for (let i = 1; i <= pdf.numPages; i++) {
        const content = await (await pdf.getPage(i)).getTextContent();
        const items = content.items.sort((a,b) => {
            const ay = Math.round(a.transform[5]*10), by = Math.round(b.transform[5]*10);
            return ay !== by ? by - ay : a.transform[4] - b.transform[4];
        });
        fullText += items.map(it => it.str).join(' ') + '\n';
    }
    return fullText;
}

// ============================================================
//  V5.0 – Robustes Portfolio
//  - Schichten 1-6 stehen direkt unter den Zeitphasen 1-6
//  - Tagesgeld & Unternehmerisches Risiko nebeneinander losgelöst
//  - Fondsauswahl mit "+ Auswählen" / "✓ Ausgewählt" Toggle
//  - Hover-Tooltip für Top 5 Ländergewichtungen je Fonds
//  - Vollständige Ländergewichtungen für Einmalbeitrag & Sparrate
// ============================================================

const layerColors = {
    "block-kasse": "#FFB300",
    "block-defensiv": "#6E9E2E",
    "block-ausgewogen": "#8CC63F",
    "block-dynamisch": "#A8D84E",
    "block-maerkte-weit": "#5D9CEC",
    "block-maerkte-eng": "#2B4C7E",
    "block-spezial": "#00B1EB",
    "block-tagesgeld": "#FFCA28"
};

function getManagementBlockDisplayTitle(blockId) {
    switch (blockId) {
        case 'block-tagesgeld':
            return 'Tagesgeld';
        case 'block-kasse':
            return 'Zeitphase 1: Geldmarkt';
        case 'block-defensiv':
            return 'Zeitphase 2: Zielrendite / WB Anleihen / EB Anleihen';
        case 'block-ausgewogen':
            return 'Zeitphase 3: Zielrendite / WB Anleihen / EB Anleihen';
        case 'block-dynamisch':
            return 'Zeitphase 4: WB Aktien/Anleihen';
        case 'block-maerkte-weit':
            return 'Zeitphase 5: WB Aktien';
        case 'block-maerkte-eng':
            return 'Zeitphase 6: EB Aktien';
        case 'block-spezial':
            return 'Echte / unechte Anlageklassen mit unternehmerischen Risiken';
        default: {
            if (typeof managementBlocks !== 'undefined') {
                const b = managementBlocks.find(mb => mb.id === blockId);
                if (b && b.title) return b.title.replace(/<br\s*\/?>/gi, ' / ');
            }
            return blockId;
        }
    }
}

function formatManagementBlockHeaderTitle(blockId) {
    const full = getManagementBlockDisplayTitle(blockId);
    if (full.includes(': ')) {
        const parts = full.split(': ');
        return `<span class="panel-layer-phase-num">${parts[0]}:</span><span class="panel-layer-phase-name">${parts.slice(1).join(': ')}</span>`;
    }
    return `<span class="panel-layer-phase-name">${full}</span>`;
}

let collapsedLayers = new Set();
try {
    const savedCollapsed = localStorage.getItem('portfolio_collapsed_layers_v6');
    if (savedCollapsed) {
        const parsed = JSON.parse(savedCollapsed);
        if (Array.isArray(parsed)) collapsedLayers = new Set(parsed);
    }
} catch (e) {}

let collapsedClustersEinmal = new Set();
try {
    const savedEinmal = localStorage.getItem('portfolio_collapsed_clusters_einmal_v6');
    if (savedEinmal) {
        const parsed = JSON.parse(savedEinmal);
        if (Array.isArray(parsed)) collapsedClustersEinmal = new Set(parsed);
    }
} catch (e) {}

let collapsedClustersSparrate = new Set();
try {
    const savedSpar = localStorage.getItem('portfolio_collapsed_clusters_sparrate_v6');
    if (savedSpar) {
        const parsed = JSON.parse(savedSpar);
        if (Array.isArray(parsed)) collapsedClustersSparrate = new Set(parsed);
    }
} catch (e) {}

let activeClustersEinmal = [];
let activeClustersSparrate = [];

const zeitphasen = [
    { id: "phase1", name: "Zeitphase 1", duration: "< 1 Jahr", assetClass: "Geldmarkt", mappedBlocks: ["block-kasse"], exactMatchBlock: "block-kasse" },
    { id: "phase2", name: "Zeitphase 2", duration: "> 2 Jahre", assetClass: "Zielrendite,<br>WB Anleihen,<br>EB Anleihen", mappedBlocks: ["block-defensiv"], exactMatchBlock: "block-defensiv" },
    { id: "phase3", name: "Zeitphase 3", duration: "> 4 Jahre", assetClass: "Zielrendite,<br>WB Anleihen,<br>EB Anleihen", mappedBlocks: ["block-ausgewogen"], exactMatchBlock: "block-ausgewogen" },
    { id: "phase4", name: "Zeitphase 4", duration: "> 6 Jahre", assetClass: "WB Aktien/Anleihen", mappedBlocks: ["block-dynamisch"], exactMatchBlock: "block-dynamisch" },
    { id: "phase5", name: "Zeitphase 5", duration: "> 8 Jahre", assetClass: "WB Aktien", mappedBlocks: ["block-maerkte-weit"], exactMatchBlock: "block-maerkte-weit" },
    { id: "phase6", name: "Zeitphase 6", duration: "> 10 Jahre", assetClass: "EB Aktien", mappedBlocks: ["block-maerkte-eng"], exactMatchBlock: "block-maerkte-eng" }
];

const CLUSTER_DEFS = {
    'Nordamerika': { color: '#3b82f6', countries: new Set(['USA', 'Vereinigte Staaten', 'Kanada', 'Mexiko']) },
    'Europa': { color: '#10b981', countries: new Set(['Deutschland', 'Frankreich', 'Großbritannien', 'Niederlande', 'Schweiz', 'Österreich', 'Spanien', 'Italien', 'Schweden', 'Dänemark', 'Finnland', 'Norwegen', 'Belgien', 'Luxemburg', 'Europa', 'Sonstige Länder']) },
    'Asien': { color: '#f59e0b', countries: new Set(['Japan', 'China', 'Indien', 'Taiwan', 'Südkorea', 'Hongkong', 'Singapur', 'Asien', 'Mauritius', 'Indonesien', 'Vietnam', 'Pakistan']) },
    'Schwellenländer': { color: '#8b5cf6', countries: new Set(['Brasilien', 'Schwellenländer', 'Chile', 'Peru', 'Kolumbien', 'Südafrika', 'Ägypten', 'Rumänien', 'Kuwait', 'Kasachstan', 'Namibia']) },
    'Sonstige': { color: '#64748b', countries: new Set(['global', 'Global', 'sonstige']) }
};
const CLUSTER_ORDER = ['Nordamerika', 'Europa', 'Asien', 'Schwellenländer', 'Sonstige'];

function classifyFund(type) {
    if (!type) return 'aktien';
    const t = type.toLowerCase();
    if (t.includes('anleihen') || t.includes('fixed income') || t.includes('bond') || t.includes('geldmarkt') || t.includes('kasse') || t.includes('renten')) {
        return 'anleihen';
    }
    return 'aktien';
}

const STORAGE_KEY  = 'portfolioV50_setup';
const GLOBALS_KEY  = 'portfolioGlobalsV50_setup';

const delistetFundName = (blockTitle) => `${blockTitle} – Fonds delistet`;

function ensureDelistetFunds() {
    if (typeof managementBlocks === 'undefined') return;
    managementBlocks.forEach(block => {
        if (block.id === 'block-tagesgeld' || (block.title && block.title.toLowerCase().includes('tagesgeld'))) {
            block.funds = block.funds.filter(f => !f._isDelistet && !f.name.toLowerCase().includes('delistet'));
            return;
        }
        const name = delistetFundName(block.title);
        if (!block.funds.some(f => f.name === name)) {
            block.funds.push({
                name,
                info: 'Fonds aus der VEM-Liste entfernt, aber noch im Depot',
                type: 'Delistet',
                ertrag: '',
                _isDelistet: true
            });
        }
    });
}
ensureDelistetFunds();

function resetPortfolio() {
    portfolio = [];
    portfolioGlobals.fundInvestments = {};
    portfolioGlobals.fundSparrates = {};
    portfolioGlobals.totalInvestment = 0;
    portfolioGlobals.totalSparrate = 0;
    try { localStorage.removeItem('empfehlungslisteFonds_V44'); } catch(e) {}
    savePortfolio();
    saveGlobals();
    const totalInput = document.getElementById('total-investment-input');
    const totalSpar = document.getElementById('total-investment-sparrate');
    if (totalInput) totalInput.value = '';
    if (totalSpar) totalSpar.value = '';
    if (typeof window.updatePortfolioUI === 'function') {
        window.updatePortfolioUI();
    } else if (typeof updatePortfolioUI === 'function') {
        updatePortfolioUI();
    }
}

let portfolio = loadPortfolio();

function loadPortfolio() {
    try {
        const r = localStorage.getItem(STORAGE_KEY);
        if (r) {
            let loaded = JSON.parse(r);
            if (Array.isArray(loaded)) {
                loaded.forEach(layer => {
                    if (layer.blockId === 'block-tagesgeld') {
                        layer.funds = layer.funds.filter(f => f.name === 'Tagesgeldkonto');
                    }
                });
                loaded = loaded.filter(layer => layer.funds && layer.funds.length > 0);
            }
            return loaded;
        }
    } catch (e) { }
    return [];
}
function savePortfolio() { localStorage.setItem(STORAGE_KEY, JSON.stringify(portfolio)); }

function sanitizeGlobals(g) {
    if (!g) return { totalInvestment: 0, totalSparrate: 0, fundInvestments: {}, fundSparrates: {}, initialInvestments: {}, initialSparrates: {}, initialTotalInvestment: 0 };
    const cleanInvestments = {};
    const cleanSparrates = {};
    const cleanInitial = {};
    if (g.fundInvestments) {
        Object.keys(g.fundInvestments).forEach(k => {
            const v = Math.round((Number(g.fundInvestments[k]) || 0) * 100) / 100;
            if (v > 0) {
                const normKey = k.includes('::') ? k : k.replace('_', '::');
                if (normKey.startsWith('block-tagesgeld::') && !normKey.endsWith('::Tagesgeldkonto')) return;
                if (!cleanInvestments[normKey]) {
                    cleanInvestments[normKey] = v;
                }
            }
        });
    }
    if (g.fundSparrates) {
        Object.keys(g.fundSparrates).forEach(k => {
            const v = Math.round((Number(g.fundSparrates[k]) || 0) * 100) / 100;
            if (v > 0) {
                const normKey = k.includes('::') ? k : k.replace('_', '::');
                if (normKey.startsWith('block-tagesgeld::') && !normKey.endsWith('::Tagesgeldkonto')) return;
                if (!cleanSparrates[normKey]) {
                    cleanSparrates[normKey] = v;
                }
            }
        });
    }
    const cleanInitialSparrates = {};
    if (g.initialInvestments) {
        Object.keys(g.initialInvestments).forEach(k => {
            const v = Math.round((Number(g.initialInvestments[k]) || 0) * 100) / 100;
            if (v > 0) {
                const normKey = k.includes('::') ? k : k.replace('_', '::');
                cleanInitial[normKey] = v;
            }
        });
    }
    if (g.initialSparrates) {
        Object.keys(g.initialSparrates).forEach(k => {
            const v = Math.round((Number(g.initialSparrates[k]) || 0) * 100) / 100;
            if (v > 0) {
                const normKey = k.includes('::') ? k : k.replace('_', '::');
                cleanInitialSparrates[normKey] = v;
            }
        });
    }
    return {
        totalInvestment: Math.round((Number(g.totalInvestment) || 0) * 100) / 100,
        totalSparrate: Math.round((Number(g.totalSparrate) || 0) * 100) / 100,
        fundInvestments: cleanInvestments,
        fundSparrates: cleanSparrates,
        initialInvestments: cleanInitial,
        initialSparrates: cleanInitialSparrates,
        initialTotalInvestment: Math.round((Number(g.initialTotalInvestment) || 0) * 100) / 100
    };
}

function loadGlobals() {
    try {
        const r = localStorage.getItem(GLOBALS_KEY);
        if (r) return sanitizeGlobals(JSON.parse(r));
    } catch (e) { }
    return {
        totalInvestment: 0,
        totalSparrate: 0,
        fundInvestments: {},
        fundSparrates: {},
        initialInvestments: {},
        initialSparrates: {},
        initialTotalInvestment: 0
    };
}
let portfolioGlobals = loadGlobals();
function saveGlobals() { localStorage.setItem(GLOBALS_KEY, JSON.stringify(portfolioGlobals)); }

window.__hardResetApp = function() {
    const keysToDelete = [];
    for (let i = 0; i < localStorage.length; i++) {
        const k = localStorage.key(i);
        if (k && (k.includes('V50') || k.includes('V44') || k.includes('V60') || k.includes('portfolioGlobals') ||
                  k.includes('delistedFunds') || k.includes('empfehlungslisteFonds') || k.includes('portfolio_'))) {
            if (!k.includes('depotBibliothek') && !k.includes('vem_library')) {
                keysToDelete.push(k);
            }
        }
    }
    keysToDelete.forEach(k => localStorage.removeItem(k));
    location.reload();
};

function getOrCreateLayer(blockId, blockTitle) {
    let layer = portfolio.find(l => l.blockId === blockId);
    if (!layer) {
        layer = { blockId, blockTitle, allocation: 0, funds: [] };
        portfolio.push(layer);
    }
    return layer;
}

function isFundSelected(blockId, fundName) {
    const l = portfolio.find(l => l.blockId === blockId);
    return l ? l.funds.some(f => f.name === fundName) : false;
}

function addFund(block, fund) {
    const layer = getOrCreateLayer(block.id, block.title);
    if (!isFundSelected(block.id, fund.name)) {
        const entry = { name: fund.name, info: fund.info, type: fund.type, ertrag: fund.ertrag };
        if (fund._isDelistet)         entry._isDelistet       = true;
        if (fund._isEmpfehlungsliste) entry._isEmpfehlungsliste = true;
        layer.funds.push(entry);
    }
    savePortfolio();
}

function removeFund(blockId, fundName) {
    const layer = portfolio.find(l => l.blockId === blockId);
    if (layer) {
        const idx = layer.funds.findIndex(f => f.name === fundName);
        if (idx > -1) {
            layer.funds.splice(idx, 1);
            if (layer.funds.length === 0) {
                portfolio = portfolio.filter(l => l.blockId !== blockId);
            }
        }
    }
    const kColon = `${blockId}::${fundName}`;
    const kUnder = `${blockId}_${fundName}`;
    delete portfolioGlobals.fundInvestments[kColon];
    delete portfolioGlobals.fundInvestments[kUnder];
    delete portfolioGlobals.fundSparrates[kColon];
    delete portfolioGlobals.fundSparrates[kUnder];
    savePortfolio();
    saveGlobals();
}

function formatCurrency(val) {
    return new Intl.NumberFormat('de-DE', { style: 'currency', currency: 'EUR' }).format(val || 0);
}
function parseCurrencyInput(str) {
    if (!str) return 0;
    const cleanStr = String(str).replace(/\./g, '').replace(',', '.').replace(/[^0-9.]/g, '');
    const num = parseFloat(cleanStr) || 0;
    return Math.round(num * 100) / 100;
}
function formatNumberInput(num) {
    if (!num && num !== 0) return '';
    const rounded = Math.round(Number(num) * 100) / 100;
    return new Intl.NumberFormat('de-DE', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(rounded);
}

document.addEventListener('DOMContentLoaded', () => {

    // ── V6.0 IMPORT-ERGEBNIS-MODAL AKTIONEN (BIBLIOTHEK, LADEN, EXPORT) ──
    const importSaveBtn = document.getElementById('import-save-btn');
    const importLoadBtn = document.getElementById('import-load-btn');
    const importJsonDlBtn = document.getElementById('import-json-dl-btn');
    const importSaveName = document.getElementById('import-save-name');
    const pdfImportModal = document.getElementById('pdf-import-modal');
    const pdfImportModalClose = document.getElementById('pdf-import-modal-close');
    const pdfImportModalOk = document.getElementById('pdf-import-modal-ok');

    const closeImportModal = () => {
        if (pdfImportModal) {
            pdfImportModal.classList.remove('open');
            setTimeout(() => { pdfImportModal.style.display = 'none'; }, 300);
        }
    };
    if (pdfImportModalClose) pdfImportModalClose.addEventListener('click', closeImportModal);
    if (pdfImportModalOk) pdfImportModalOk.addEventListener('click', closeImportModal);

    if (importSaveBtn) {
        importSaveBtn.addEventListener('click', () => {
            const name = (importSaveName?.value || (typeof _lastImportedName !== 'undefined' ? _lastImportedName : 'Depot')).trim() || 'Depot';
            const lib = getSavedLibrary();
            const existingIdx = lib.findIndex(e => e.name === name);
            const entry = {
                id: Date.now(),
                name: name,
                date: new Date().toISOString(),
                totalInvestment: portfolioGlobals.totalInvestment,
                totalSparrate: portfolioGlobals.totalSparrate,
                portfolio: JSON.parse(JSON.stringify(portfolio)),
                investments: { ...portfolioGlobals.fundInvestments },
                fundInvestments: { ...portfolioGlobals.fundInvestments },
                fundSparrates: { ...portfolioGlobals.fundSparrates },
                empfehlungslisteFonds: JSON.parse(localStorage.getItem('empfehlungslisteFonds_V44') || '{}')
            };
            if (existingIdx > -1) lib[existingIdx] = entry;
            else lib.unshift(entry);
            saveSavedLibrary(lib);
            importSaveBtn.textContent = '✅ Gespeichert!';
            setTimeout(() => { importSaveBtn.textContent = '💾 In Bibliothek speichern'; }, 2000);
        });
    }

    if (importLoadBtn) {
        importLoadBtn.addEventListener('click', () => {
            closeImportModal();
        });
    }

    if (importJsonDlBtn) {
        importJsonDlBtn.addEventListener('click', () => {
            const name = (importSaveName?.value || (typeof _lastImportedName !== 'undefined' ? _lastImportedName : 'Depot')).trim() || 'Depot';
            const setup = {
                version: "V6.0",
                name: name,
                date: new Date().toISOString(),
                totalInvestment: portfolioGlobals.totalInvestment,
                totalSparrate: portfolioGlobals.totalSparrate,
                portfolio: portfolio,
                fundInvestments: { ...portfolioGlobals.fundInvestments },
                fundSparrates: { ...portfolioGlobals.fundSparrates },
                empfehlungslisteFonds: JSON.parse(localStorage.getItem('empfehlungslisteFonds_V44') || '{}')
            };
            const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(setup, null, 2));
            const dl = document.createElement('a');
            dl.setAttribute("href", dataStr);
            dl.setAttribute("download", `${name.replace(/\s+/g, '_')}_v6.0.json`);
            document.body.appendChild(dl);
            dl.click();
            dl.remove();
        });
    }


    const hardResetBtn = document.getElementById('hard-reset-btn');
    if (hardResetBtn) {
        hardResetBtn.addEventListener('click', () => {
            const overlay = document.createElement('div');
            overlay.style.cssText = `
                position:fixed; inset:0; background:rgba(0,0,0,0.7);
                display:flex; align-items:center; justify-content:center; z-index:999999;`;
            overlay.innerHTML = `
                <div style="background:#1a1a2e; border:2px solid #dc2626; border-radius:14px;
                     padding:36px 40px; min-width:340px; max-width:420px; text-align:center;
                     box-shadow:0 8px 40px rgba(220,38,38,0.4);">
                    <div style="font-size:2.5em; margin-bottom:12px;">⚠️</div>
                    <div style="color:#fff; font-size:1.15em; font-weight:700; margin-bottom:8px;">Hard Reset</div>
                    <div style="color:#fca5a5; font-size:0.9em; margin-bottom:16px;">Alle Eingaben werden gelöscht.</div>
                    <div style="display:flex; gap:12px; justify-content:center;">
                        <button id="hr-cancel" style="padding:10px 24px; border-radius:8px; border:1px solid #4b5563; background:transparent; color:#d1d5db; cursor:pointer;">Abbrechen</button>
                        <button id="hr-confirm" style="padding:10px 28px; border-radius:8px; border:none; background:#dc2626; color:#fff; cursor:pointer; font-weight:700;">Ja, alles löschen</button>
                    </div>
                </div>`;
            document.body.appendChild(overlay);
            overlay.querySelector('#hr-cancel').addEventListener('click', () => overlay.remove());
            overlay.querySelector('#hr-confirm').addEventListener('click', () => {
                overlay.remove();
                window.__hardResetApp();
            });
            overlay.addEventListener('click', e => { if (e.target === overlay) overlay.remove(); });
        });
    }

    const gridContainer = document.getElementById('v5-grid-container');
    const tagesgeldContainer = document.getElementById('block-tagesgeld-container');
    const spezialContainer = document.getElementById('block-spezial-container');

    const modal = document.getElementById('funds-modal');
    const modalTitle = document.getElementById('modal-title');
    const modalList = document.getElementById('modal-funds-list');
    const modalFilter = document.getElementById('modal-filter');
    const closeBtn = document.querySelector('#funds-modal .close-modal');

    const totalInvestmentInput = document.getElementById('total-investment-input');
    const totalDistributedDisplay = document.getElementById('total-distributed-display');
    const totalRemainingDisplay = document.getElementById('total-remaining-display');
    const totalInvestmentSparrate = document.getElementById('total-investment-sparrate');
    const sparrateDistributedDisplay = document.getElementById('sparrate-distributed-display');
    const sparrateRemainingDisplay = document.getElementById('sparrate-remaining-display');

    const panelEmpty = document.getElementById('panel-empty');
    const panelLayers = document.getElementById('panel-layers');
    const panelHiddenState = document.getElementById('panel-hidden-state');
    const toggleVisibilityBtn = document.getElementById('toggle-managers-visibility-btn');
    const resetBtn = document.getElementById('reset-portfolio-btn');
    const savePortfolioBtn = document.getElementById('save-portfolio-btn');

    let managersHidden = false;

    function setManagersVisibility(hidden) {
        managersHidden = !!hidden;
        const btn = document.getElementById('toggle-managers-visibility-btn') || toggleVisibilityBtn;
        const pLayers = document.getElementById('panel-layers') || panelLayers;
        const pEmpty = document.getElementById('panel-empty') || panelEmpty;
        const pHidden = document.getElementById('panel-hidden-state') || panelHiddenState;

        if (btn) {
            if (managersHidden) {
                btn.classList.add('active');
                btn.title = "Auswahl operativer Manager einblenden";
                btn.setAttribute('aria-label', "Auswahl operativer Manager einblenden");
                btn.innerHTML = `<svg id="visibility-icon-svg" xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24"></path><line x1="1" y1="1" x2="23" y2="23"></line></svg>`;
            } else {
                btn.classList.remove('active');
                btn.title = "Auswahl operativer Manager ausblenden";
                btn.setAttribute('aria-label', "Auswahl operativer Manager ausblenden");
                btn.innerHTML = `<svg id="visibility-icon-svg" xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path><circle cx="12" cy="12" r="3"></circle></svg>`;
            }
        }

        if (managersHidden) {
            if (pLayers) pLayers.style.display = 'none';
            if (pEmpty) pEmpty.style.display = 'none';
            if (pHidden) pHidden.style.display = 'block';
        } else {
            if (pHidden) pHidden.style.display = 'none';
            if (pEmpty) pEmpty.style.display = 'none';
            if (pLayers) {
                pLayers.style.display = 'block';
                renderRightPanelLayers();
            }
        }
    }
    window.setManagersVisibility = setManagersVisibility;

    let currentBlockData = null;
    let _searchTargetFundName = null;
    let activePhaseId = null;

    // ── PROMINENT SEARCH FUNCTIONALITY ────────────────────────
    const searchInput = document.getElementById('fund-search-input');
    const searchResults = document.getElementById('fund-search-results');

    if (searchInput && searchResults) {
        searchInput.addEventListener('input', (e) => {
            const query = e.target.value.toLowerCase().trim();
            searchResults.innerHTML = '';
            if (query.length < 2) { searchResults.style.display = 'none'; return; }

            let matches = [];
            managementBlocks.forEach(block => {
                block.funds.forEach(fund => {
                    const searchStr = `${fund.name} ${fund.info} ${fund.type}`.toLowerCase();
                    if (searchStr.includes(query)) matches.push({ fund, block });
                });
            });

            if (matches.length === 0) {
                searchResults.innerHTML = '<li><span class="fund-search-meta">Keine Fonds gefunden.</span></li>';
                searchResults.style.display = 'block';
                return;
            }

            matches.forEach(match => {
                const li = document.createElement('li');
                const badgeSubtextHtml = match.fund.badgeSubtext ? `
                    <span class="badge-info-wrapper">
                        <i class="badge-info-icon">i</i>
                        <span class="badge-info-tooltip">${match.fund.badgeSubtext}</span>
                    </span>` : '';
                li.innerHTML = `
                    <span class="fund-search-name" style="display:flex; align-items:center; flex-wrap:wrap; gap:6px;">
                        ${match.fund.name}
                        <span class="fund-badge" style="font-size:0.72rem;padding:2px 6px;">${match.fund.type}</span>
                        ${badgeSubtextHtml}
                        ${match.fund.anlageschwerpunkt ? `<span class="fund-badge-schwerpunkt" style="font-size:0.72rem;padding:2px 6px;">${match.fund.anlageschwerpunkt}</span>` : ''}
                    </span>
                    <span class="fund-search-meta">${match.fund.info} | Topf: ${match.block.title.replace(/<br\s*\/?>/gi, ' ')}</span>`;
                li.addEventListener('click', () => {
                    searchInput.value = ''; searchResults.style.display = 'none';
                    _searchTargetFundName = match.fund.name;
                    openFundModalForBlock(match.block);
                });
                searchResults.appendChild(li);
            });
            searchResults.style.display = 'block';
        });

        document.addEventListener('click', e => {
            if (!searchInput.contains(e.target) && !searchResults.contains(e.target)) {
                searchResults.style.display = 'none';
            }
        });
    }

    // ── RENDER 6 PHASE COLUMNS ───────────────────────────────
    function renderGridColumns() {
        if (!gridContainer) return;
        gridContainer.innerHTML = '';

        zeitphasen.forEach(phase => {
            const col = document.createElement('div');
            col.className = 'v5-phase-col';

            const header = document.createElement('div');
            header.className = `v5-phase-header ${activePhaseId === phase.id ? 'active' : ''}`;
            header.dataset.phaseId = phase.id;
            header.innerHTML = `
                <div>${phase.name}</div>
                <span class="duration-badge">${phase.duration}</span>`;
            header.addEventListener('click', () => handlePhaseClick(phase));

            const blockId = phase.exactMatchBlock;
            const block = managementBlocks.find(b => b.id === blockId);
            const cylinder = createCylinderElement(block);

            col.appendChild(header);
            col.appendChild(cylinder);
            gridContainer.appendChild(col);
        });
    }

    // ── RENDER STANDALONE CYLINDERS (NEBENEINANDER OHNE ÜBERSCHRIFT) ──
    function renderStandaloneCylinders() {
        if (tagesgeldContainer) {
            tagesgeldContainer.innerHTML = '';
            const bTagesgeld = managementBlocks.find(b => b.id === 'block-tagesgeld');
            if (bTagesgeld) tagesgeldContainer.appendChild(createCylinderElement(bTagesgeld));
        }

        if (spezialContainer) {
            spezialContainer.innerHTML = '';
            const bSpezial = managementBlocks.find(b => b.id === 'block-spezial');
            if (bSpezial) spezialContainer.appendChild(createCylinderElement(bSpezial));
        }
    }

    function createCylinderElement(block) {
        const blockEl = document.createElement('div');
        blockEl.className = 'tower-layer highlighted';
        blockEl.id = block.id;
        blockEl.dataset.blockId = block.id;

        const bg = layerColors[block.id] || '#2563eb';
        blockEl.style.backgroundColor = bg;

        let sumEinmal = 0;
        let sumSpar = 0;
        let count = 0;

        block.funds.forEach(f => {
            const k = `${block.id}::${f.name}`;
            const kAlt = `${block.id}_${f.name}`;
            const eVal = (portfolioGlobals.fundInvestments[k] || portfolioGlobals.fundInvestments[kAlt] || 0);
            const sVal = (portfolioGlobals.fundSparrates[k] || portfolioGlobals.fundSparrates[kAlt] || 0);
            if (eVal > 0 || sVal > 0) {
                sumEinmal += eVal;
                sumSpar += sVal;
                count++;
            }
        });

        blockEl.innerHTML = `
            <div class="layer-content">
                <div class="block-title">${block.title}</div>
            </div>
            <div class="layer-selection-count ${count > 0 ? 'visible' : ''}">
                <span class="count-val">${count}</span> Fonds
            </div>
            <div class="layer-assigned-amount ${(sumEinmal > 0 || sumSpar > 0) ? 'visible' : ''}">
                ${sumEinmal > 0 ? `<span class="badge-einmal">${formatCurrency(sumEinmal)}</span>` : ''}
                ${sumSpar > 0 ? `<span class="badge-sparrate" style="font-size:0.68rem; color:#059669; font-weight:600;">${formatCurrency(sumSpar)} mtl.</span>` : ''}
            </div>`;

        blockEl.addEventListener('click', () => openFundModalForBlock(block));
        return blockEl;
    }

    function handlePhaseClick(phase) {
        if (activePhaseId === phase.id) {
            activePhaseId = null;
            resetHighlights();
            return;
        }
        activePhaseId = phase.id;

        document.querySelectorAll('.v5-phase-header').forEach(h => h.classList.remove('active'));
        const activeHeader = document.querySelector(`.v5-phase-header[data-phase-id="${phase.id}"]`);
        if (activeHeader) activeHeader.classList.add('active');

        document.querySelectorAll('.tower-layer').forEach(card => card.classList.remove('highlighted', 'exact-match'));
        
        phase.mappedBlocks.forEach(bId => {
            const card = document.getElementById(bId);
            if (card) {
                card.classList.add('highlighted', 'exact-match');
                const bgColor = window.getComputedStyle(card).backgroundColor;
                card.style.setProperty('--badge-color', bgColor);
            }
        });
    }

    function resetHighlights() {
        document.querySelectorAll('.v5-phase-header').forEach(h => h.classList.remove('active'));
        document.querySelectorAll('.tower-layer').forEach(card => {
            card.classList.add('highlighted');
            card.classList.remove('exact-match');
        });
    }

    // ── MODAL OPEN & FUND LIST WITH "+ AUSWÄHLEN" / "✓ AUSGEWÄHLT" & TOP 5 COUNTRY HOVER TOOLTIP ──
    function openFundModalForBlock(block) {
        currentBlockData = block;
        modalTitle.innerHTML = block.title.replace(/<br\s*\/?>/gi, ' ');
        modalFilter.value = 'all';
        populateFilterDropdown(block.funds, block.id);
        renderFundList(block.funds);
        modal.classList.add('open');

        if (_searchTargetFundName) {
            const targetName = _searchTargetFundName;
            _searchTargetFundName = null;
            setTimeout(() => {
                const allItems = modalList.querySelectorAll('.fund-list-item');
                allItems.forEach(item => {
                    if (item.dataset.fundName === targetName) {
                        item.scrollIntoView({ behavior: 'smooth', block: 'center' });
                        item.classList.add('search-highlight-flash');
                        setTimeout(() => item.classList.remove('search-highlight-flash'), 2500);
                    }
                });
            }, 200);
        }
    }

    function populateFilterDropdown(funds, blockId) {
        modalFilter.innerHTML = '<option value="all">Alle Arten anzeigen (Filter)</option>';
        const types = [...new Set(funds.map(f => f.type).filter(Boolean))];
        types.forEach(t => {
            const opt = document.createElement('option');
            opt.value = t; opt.textContent = t;
            modalFilter.appendChild(opt);
        });
    }

    const refreshModalButtons = () => {
        if (!currentBlockData) return;
        modalList.querySelectorAll('.fund-list-item').forEach(li => {
            const sel = isFundSelected(currentBlockData.id, li.dataset.fundName);
            const btn = li.querySelector('.btn-fund-select');
            if (!btn) return;
            li.classList.toggle('is-selected', sel);
            btn.classList.toggle('selected', sel);
            btn.textContent = sel ? '✓ Ausgewählt' : '+ Auswählen';
        });
    };

    function renderFundList(funds) {
        modalList.innerHTML = '';
        const filterVal = modalFilter ? modalFilter.value : 'all';
        const filtered = filterVal === 'all' ? funds : funds.filter(f => f.type === filterVal);

        if (!filtered || filtered.length === 0) {
            modalList.innerHTML = '<li class="fund-list-item"><span class="fund-name">Keine Fonds für diesen Filter.</span></li>';
            return;
        }

        // ── Sortierung: Aktive Fonds A→Z, dann delistete Fonds A→Z ──
        const activeFunds   = filtered.filter(f => !f._isDelistet).sort((a, b) => a.name.localeCompare(b.name, 'de'));
        const delistedFunds = filtered.filter(f =>  f._isDelistet).sort((a, b) => a.name.localeCompare(b.name, 'de'));

        // ── Abschnitts-Überschrift einfügen ──
        function insertSectionHeader(label, isFirst) {
            const header = document.createElement('li');
            header.className = 'fund-list-section-header';
            header.innerHTML = `
                ${!isFirst ? '<hr class="fund-list-section-divider">' : ''}
                <span class="fund-list-section-label">${label}</span>`;
            modalList.appendChild(header);
        }

        if (activeFunds.length > 0)   insertSectionHeader('Aktiv', true);
        const sortedFunds = [...activeFunds, ...delistedFunds];
        let delistedHeaderInserted = false;

        sortedFunds.forEach(fund => {
            // Überschrift "Delistet" vor dem ersten delisteten Fonds
            if (fund._isDelistet && !delistedHeaderInserted) {
                insertSectionHeader('Delistet', activeFunds.length === 0);
                delistedHeaderInserted = true;
            }

            const sel = isFundSelected(currentBlockData.id, fund.name);
            const keyColon = `${currentBlockData.id}::${fund.name}`;
            const keyUnder = `${currentBlockData.id}_${fund.name}`;
            const valEinmal = (portfolioGlobals.fundInvestments[keyColon] !== undefined ? portfolioGlobals.fundInvestments[keyColon] : portfolioGlobals.fundInvestments[keyUnder]) || 0;
            const valSpar = (portfolioGlobals.fundSparrates[keyColon] !== undefined ? portfolioGlobals.fundSparrates[keyColon] : portfolioGlobals.fundSparrates[keyUnder]) || 0;
            const key = keyColon;

            // ── TOP 5 COUNTRY HOVER TOOLTIP GENERATOR (EXACT V4.4) ──
            let countryTooltipHtml = '';
            if (fund.countryWeightings && fund.countryWeightings.length > 0) {
                const fClass = classifyFund(fund.type);
                const sortedCW = [...fund.countryWeightings].sort((a, b) => b.weight - a.weight).slice(0, 5);

                const listHtml = sortedCW.map(c => {
                    let subLine = '';
                    const formattedWeight = typeof c.weight === 'number' ? c.weight.toFixed(1).replace('.', ',') : c.weight;
                    if (fClass === 'aktien') {
                        subLine = `<div style="font-size:0.8em;color:#93c5fd;padding-left:4px;">&#x2514; Aktien ${formattedWeight}%</div>`;
                    } else if (fClass === 'anleihen') {
                        subLine = `<div style="font-size:0.8em;color:#86efac;padding-left:4px;">&#x2514; Anleihen ${formattedWeight}%</div>`;
                    }
                    return `<li style="margin-bottom:5px;"><div style="display:flex;justify-content:space-between;gap:10px;"><span>${c.country}</span><strong>${c.weight}%</strong></div>${subLine}</li>`;
                }).join('');

                const classLabel = fClass === 'aktien' ? '<span style="font-size:0.75em;background:#1e3a8a;color:#93c5fd;border-radius:4px;padding:1px 6px;margin-left:6px;">Aktien</span>'
                                 : fClass === 'anleihen' ? '<span style="font-size:0.75em;background:#14532d;color:#86efac;border-radius:4px;padding:1px 6px;margin-left:6px;">Anleihen</span>'
                                 : '<span style="font-size:0.75em;background:#334155;color:#cbd5e1;border-radius:4px;padding:1px 6px;margin-left:6px;">Gemischt</span>';

                countryTooltipHtml = `
                    <div class="fund-country-tooltip">
                        <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" class="country-icon">
                            <circle cx="12" cy="12" r="10"></circle>
                            <line x1="2" y1="12" x2="22" y2="12"></line>
                            <path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10z"></path>
                        </svg>
                        <div class="tooltip-content">
                            <strong style="font-size:1.05em; border-bottom:1px solid #555; display:flex; align-items:center; padding-bottom:5px; margin-bottom:5px;">Top 5 Länder ${classLabel}</strong>
                            <ul style="margin:0;padding:0;list-style:none;">${listHtml}</ul>
                        </div>
                    </div>`;
            }

            const ertragCls = fund.ertrag === 'thesaurierend' ? 'thesaurierend' : fund.ertrag === 'ausschüttend' ? 'ausschuettend' : '';
            const ertragHtml = fund.ertrag ? `<span class="fund-badge-ertrag ${ertragCls}">${fund.ertrag}</span>` : '';
            const isDelistet = !!fund._isDelistet;

            const item = document.createElement('li');
            item.className = `fund-list-item${sel ? ' is-selected' : ''}${isDelistet ? ' fund-delistet' : ''}`;
            item.dataset.fundName = fund.name;

            const delistetBadgeHtml = isDelistet
                ? `<span class="delisted-badge">Fonds delistet</span>`
                : '';

            // ── Eingabesperre für delistete Fonds ohne bestehende Investition ──
            const isDelistedLocked = isDelistet && valEinmal === 0 && valSpar === 0;
            const lockedAttr  = isDelistedLocked ? 'disabled title="Erstinvestition in delistete Fonds nicht möglich"' : '';
            const lockedClass = isDelistedLocked ? ' delisted-input-locked' : '';

            const hasInitialDepot = Boolean(portfolioGlobals.initialInvestments && Object.keys(portfolioGlobals.initialInvestments).length > 0);
            const valInitial = (portfolioGlobals.initialInvestments && (portfolioGlobals.initialInvestments[key] || portfolioGlobals.initialInvestments[`${currentBlockData.id}_${fund.name}`])) || 0;
            const valInitialSpar = (portfolioGlobals.initialSparrates && (portfolioGlobals.initialSparrates[key] || portfolioGlobals.initialSparrates[`${currentBlockData.id}_${fund.name}`])) || 0;

            let zuzahlungCurrent = 0;
            let isNeuanlage = false;
            if (hasInitialDepot) {
                if (valInitial === 0 && valEinmal > 0) {
                    zuzahlungCurrent = valEinmal;
                    isNeuanlage = true;
                } else if (valInitial > 0 && valEinmal > valInitial) {
                    zuzahlungCurrent = Math.round((valEinmal - valInitial) * 100) / 100;
                    isNeuanlage = false;
                }
            }

            let sparrateNeuCurrent = 0;
            if (hasInitialDepot) {
                if (valInitialSpar === 0 && valSpar > 0) {
                    sparrateNeuCurrent = valSpar;
                } else if (valInitialSpar > 0 && valSpar > valInitialSpar) {
                    sparrateNeuCurrent = Math.round((valSpar - valInitialSpar) * 100) / 100;
                }
            }

            const fundInput = hasInitialDepot ? `
                <div class="fund-input-wrapper">
                    <div class="fund-input-col">
                        <label class="fund-input-label">Bisheriger Bestand</label>
                        <div class="currency-input-wrapper">
                            <input type="text" class="fund-investment-input${lockedClass}" data-fund-key="${key}"
                                placeholder="0,00" inputmode="decimal"
                                value="${valInitial > 0 ? formatNumberInput(valInitial) : (valEinmal > 0 ? '0,00' : '')}" ${lockedAttr}>
                            <span class="currency-symbol">€</span>
                        </div>
                        ${valInitial > 0 ? `<div class="fund-bestand-tag">Bestand: <strong>${formatCurrency(valInitial)}</strong></div>` : (valEinmal > 0 ? `<div class="fund-bestand-tag" style="color:#64748b;">Bestand: <strong>0,00 €</strong></div>` : '')}
                    </div>
                    <div class="fund-input-col">
                        <label class="fund-input-label" style="color:#15803d; font-weight:700;">${valInitial === 0 ? '+ Neuanlage' : '+ Zuzahlung'}</label>
                        <div class="currency-input-wrapper">
                            <input type="text" class="fund-zuzahlung-input${lockedClass}" data-fund-key="${key}"
                                placeholder="+ 0,00" inputmode="decimal"
                                value="${zuzahlungCurrent > 0 ? formatNumberInput(zuzahlungCurrent) : ''}" ${lockedAttr}>
                            <span class="currency-symbol" style="color:#15803d; font-weight:700;">€</span>
                        </div>
                        <div class="fund-total-tag" style="font-size:10px; color:#15803d; font-weight:700; margin-top:2px;">
                            ${valEinmal > 0 ? `Gesamt: ${formatCurrency(valEinmal)}` : ''}
                        </div>
                    </div>
                    <div class="fund-input-col">
                        <label class="fund-input-label">Sparrate mtl.</label>
                        <div class="currency-input-wrapper">
                            <input type="text" class="fund-sparrate-input${lockedClass}" data-fund-key="${key}"
                                placeholder="0,00" inputmode="decimal"
                                value="${valSpar > 0 ? formatNumberInput(valSpar) : ''}" ${lockedAttr}>
                            <span class="currency-symbol">€</span>
                        </div>
                        ${sparrateNeuCurrent > 0 ? `<div style="font-size:10px; color:#15803d; font-weight:700; margin-top:2px;">+ ${formatCurrency(sparrateNeuCurrent)} mtl. neu</div>` : (valInitialSpar > 0 ? `<div class="fund-bestand-tag">Bestand: <strong>${formatCurrency(valInitialSpar)} mtl.</strong></div>` : '')}
                    </div>
                </div>` : `
                <div class="fund-input-wrapper">
                    <div class="fund-input-col">
                        <label class="fund-input-label">Gesamt Einmalbeitrag</label>
                        <div class="currency-input-wrapper">
                            <input type="text" class="fund-investment-input${lockedClass}" data-fund-key="${key}"
                                placeholder="0,00" inputmode="decimal"
                                value="${valEinmal > 0 ? formatNumberInput(valEinmal) : ''}" ${lockedAttr}>
                            <span class="currency-symbol">€</span>
                        </div>
                    </div>
                    <div class="fund-input-col">
                        <label class="fund-input-label">Sparrate mtl.</label>
                        <div class="currency-input-wrapper">
                            <input type="text" class="fund-sparrate-input${lockedClass}" data-fund-key="${key}"
                                placeholder="0,00" inputmode="decimal"
                                value="${valSpar > 0 ? formatNumberInput(valSpar) : ''}" ${lockedAttr}>
                            <span class="currency-symbol">€</span>
                        </div>
                    </div>
                </div>`;

            item.innerHTML = `
                <div class="fund-info-wrapper">
                    <div class="fund-header" style="display:flex; align-items:center; gap:8px;">
                        <span class="fund-name">${fund.name}</span>
                        ${delistetBadgeHtml}
                        ${countryTooltipHtml}
                    </div>
                    <div style="margin-top:4px;display:flex;gap:8px;flex-wrap:wrap;align-items:center;">
                        ${!isDelistet && fund.type ? `<span class="fund-badge">${fund.type}</span>` : ''}
                        ${!isDelistet && fund.badgeSubtext ? `
                            <span class="badge-info-wrapper">
                                <i class="badge-info-icon">i</i>
                                <span class="badge-info-tooltip">${fund.badgeSubtext}</span>
                            </span>` : ''}
                        ${ertragHtml}
                        ${!isDelistet && fund.anlageschwerpunkt ? `<span class="fund-badge-schwerpunkt">${fund.anlageschwerpunkt}</span>` : ''}
                    </div>
                    <span class="fund-info" style="margin-top:6px;display:block;">${fund.info}</span>
                </div>
                ${fundInput}
                <button class="btn-fund-select${sel ? ' selected' : ''}"
                    title="${sel ? 'Klicken zum Entfernen' : 'Zum Portfolio hinzufügen'}">
                    ${sel ? '✓ Ausgewählt' : '+ Auswählen'}
                </button>`;

            // Toggle select button
            const selBtn = item.querySelector('.btn-fund-select');
            const inpEinmal = item.querySelector('.fund-investment-input');
            const inpZuzahlung = item.querySelector('.fund-zuzahlung-input');
            const inpSpar = item.querySelector('.fund-sparrate-input');
            const kColon = `${currentBlockData.id}::${fund.name}`;

            selBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                if (isFundSelected(currentBlockData.id, fund.name)) {
                    removeFund(currentBlockData.id, fund.name);
                    if (inpEinmal) inpEinmal.value = '';
                    if (inpZuzahlung) inpZuzahlung.value = '';
                    if (inpSpar) inpSpar.value = '';
                    const totalTag = item.querySelector('.fund-total-tag');
                    if (totalTag) totalTag.innerHTML = '';
                } else {
                    addFund(currentBlockData, fund);
                }
                refreshModalButtons();
                updatePortfolioUI();
            });

            // Helper to commit Zuzahlung / Neuanlage
            const commitZuzahlung = (zVal) => {
                const newTotal = Math.round((valInitial + zVal) * 100) / 100;
                if (hasInitialDepot) {
                    if (inpEinmal) inpEinmal.value = valInitial > 0 ? formatNumberInput(valInitial) : (newTotal > 0 ? '0,00' : '');
                } else {
                    if (inpEinmal) inpEinmal.value = newTotal > 0 ? formatNumberInput(newTotal) : '';
                }
                const totalTag = item.querySelector('.fund-total-tag');
                if (totalTag) {
                    totalTag.innerHTML = newTotal > 0 ? `Gesamt: ${formatCurrency(newTotal)}` : '';
                }
                if (newTotal > 0) {
                    portfolioGlobals.fundInvestments[kColon] = newTotal;
                    delete portfolioGlobals.fundInvestments[`${currentBlockData.id}_${fund.name}`];
                    addFund(currentBlockData, fund);
                } else {
                    delete portfolioGlobals.fundInvestments[kColon];
                    delete portfolioGlobals.fundInvestments[`${currentBlockData.id}_${fund.name}`];
                    const sVal = portfolioGlobals.fundSparrates[kColon] || 0;
                    if (sVal === 0) removeFund(currentBlockData.id, fund.name);
                }
                saveGlobals();
                refreshModalButtons();
                updatePortfolioUI();
            };

            // EUR Input listeners for Zuzahlung
            if (inpZuzahlung) {
                inpZuzahlung.addEventListener('input', e => {
                    const z = parseCurrencyInput(e.target.value);
                    commitZuzahlung(z);
                });
                inpZuzahlung.addEventListener('blur', e => {
                    const z = parseCurrencyInput(e.target.value);
                    e.target.value = z > 0 ? formatNumberInput(z) : '';
                });
                inpZuzahlung.addEventListener('focus', e => {
                    const currentTotal = portfolioGlobals.fundInvestments[kColon] || portfolioGlobals.fundInvestments[`${currentBlockData.id}_${fund.name}`] || 0;
                    const z = Math.max(0, currentTotal - valInitial);
                    e.target.value = z > 0 ? (Math.round(z * 100) / 100).toFixed(2).replace('.', ',') : '';
                    e.target.select();
                });
                inpZuzahlung.addEventListener('keydown', e => { if (e.key === 'Enter') e.target.blur(); });
            }

            // EUR Input listeners for Einmalbeitrag / Bestand
            if (inpEinmal) {
                if (hasInitialDepot) {
                    inpEinmal.addEventListener('input', e => {
                        const v = parseCurrencyInput(e.target.value);
                        if (inpZuzahlung) {
                            const diff = Math.max(0, v - valInitial);
                            inpZuzahlung.value = diff > 0 ? formatNumberInput(diff) : '';
                        }
                        const totalTag = item.querySelector('.fund-total-tag');
                        if (totalTag) {
                            totalTag.innerHTML = v > 0 ? `Gesamt: ${formatCurrency(v)}` : '';
                        }
                    });
                    inpEinmal.addEventListener('blur', e => {
                        const v = parseCurrencyInput(e.target.value);
                        if (v >= valInitial) {
                            const diff = Math.round((v - valInitial) * 100) / 100;
                            commitZuzahlung(diff);
                            if (inpZuzahlung) inpZuzahlung.value = diff > 0 ? formatNumberInput(diff) : '';
                            e.target.value = valInitial > 0 ? formatNumberInput(valInitial) : (v > 0 ? '0,00' : '');
                        } else if (v > 0) {
                            portfolioGlobals.fundInvestments[kColon] = v;
                            delete portfolioGlobals.fundInvestments[`${currentBlockData.id}_${fund.name}`];
                            addFund(currentBlockData, fund);
                            if (inpZuzahlung) inpZuzahlung.value = '';
                            e.target.value = formatNumberInput(v);
                            saveGlobals();
                            refreshModalButtons();
                            updatePortfolioUI();
                        } else {
                            commitZuzahlung(0);
                            if (inpZuzahlung) inpZuzahlung.value = '';
                            e.target.value = valInitial > 0 ? formatNumberInput(valInitial) : '';
                        }
                    });
                    inpEinmal.addEventListener('focus', e => {
                        const currentTotal = portfolioGlobals.fundInvestments[kColon] || portfolioGlobals.fundInvestments[`${currentBlockData.id}_${fund.name}`] || 0;
                        e.target.value = currentTotal > 0 ? (Math.round(currentTotal * 100) / 100).toFixed(2).replace('.', ',') : '';
                        e.target.select();
                    });
                    inpEinmal.addEventListener('keydown', e => { if (e.key === 'Enter') e.target.blur(); });
                } else {
                    inpEinmal.addEventListener('input', e => {
                        const v = parseCurrencyInput(e.target.value);
                        if (v > 0) {
                            portfolioGlobals.fundInvestments[kColon] = v;
                            delete portfolioGlobals.fundInvestments[`${currentBlockData.id}_${fund.name}`];
                            addFund(currentBlockData, fund);
                        } else {
                            delete portfolioGlobals.fundInvestments[kColon];
                            delete portfolioGlobals.fundInvestments[`${currentBlockData.id}_${fund.name}`];
                            const sVal = portfolioGlobals.fundSparrates[kColon] || 0;
                            if (sVal === 0) removeFund(currentBlockData.id, fund.name);
                        }
                        saveGlobals();
                        refreshModalButtons();
                        updatePortfolioUI();
                    });
                    inpEinmal.addEventListener('blur', e => {
                        const v = parseCurrencyInput(e.target.value);
                        e.target.value = v > 0 ? formatNumberInput(v) : '';
                    });
                    inpEinmal.addEventListener('focus', e => {
                        const v = portfolioGlobals.fundInvestments[kColon] || portfolioGlobals.fundInvestments[`${currentBlockData.id}_${fund.name}`] || 0;
                        e.target.value = v > 0 ? (Math.round(v * 100) / 100).toFixed(2).replace('.', ',') : '';
                        e.target.select();
                    });
                    inpEinmal.addEventListener('keydown', e => { if (e.key === 'Enter') e.target.blur(); });
                }
            }

            // EUR Input listeners for Sparrate
            if (inpSpar) {
                inpSpar.addEventListener('input', e => {
                    const v = parseCurrencyInput(e.target.value);
                    if (v > 0) {
                        portfolioGlobals.fundSparrates[kColon] = v;
                        delete portfolioGlobals.fundSparrates[`${currentBlockData.id}_${fund.name}`];
                        addFund(currentBlockData, fund);
                    } else {
                        delete portfolioGlobals.fundSparrates[kColon];
                        delete portfolioGlobals.fundSparrates[`${currentBlockData.id}_${fund.name}`];
                        const eVal = portfolioGlobals.fundInvestments[kColon] || 0;
                        if (eVal === 0) {
                            removeFund(currentBlockData.id, fund.name);
                        }
                    }
                    saveGlobals();
                    refreshModalButtons();
                    updatePortfolioUI();
                });
                inpSpar.addEventListener('blur', e => {
                    const v = parseCurrencyInput(e.target.value);
                    e.target.value = v > 0 ? formatNumberInput(v) : '';
                });
                inpSpar.addEventListener('focus', e => {
                    const v = portfolioGlobals.fundSparrates[kColon] || portfolioGlobals.fundSparrates[`${currentBlockData.id}_${fund.name}`] || 0;
                    e.target.value = v > 0 ? (Math.round(v * 100) / 100).toFixed(2).replace('.', ',') : '';
                    e.target.select();
                });
                inpSpar.addEventListener('keydown', e => { if (e.key === 'Enter') e.target.blur(); });
            }

            // Hover Tooltip Positioner JS
            const tooltipWrapper = item.querySelector('.fund-country-tooltip');
            if (tooltipWrapper) {
                const tooltipContent = tooltipWrapper.querySelector('.tooltip-content');
                tooltipWrapper.addEventListener('mouseenter', () => {
                    const iconRect = tooltipWrapper.getBoundingClientRect();
                    const modalHeaderEl = document.querySelector('#funds-modal .modal-header');
                    const headerBottom = modalHeaderEl ? modalHeaderEl.getBoundingClientRect().bottom : 80;
                    const tooltipH = 175;
                    const spaceAbove = iconRect.top - headerBottom;

                    if (spaceAbove >= tooltipH + 10) {
                        tooltipContent.classList.remove('tip-below');
                        tooltipContent.classList.add('tip-above');
                        tooltipContent.style.top = (iconRect.top - tooltipH - 8) + 'px';
                    } else {
                        tooltipContent.classList.remove('tip-above');
                        tooltipContent.classList.add('tip-below');
                        tooltipContent.style.top = (iconRect.bottom + 8) + 'px';
                    }
                    const left = Math.max(4, iconRect.right - 210);
                    tooltipContent.style.left = left + 'px';
                });
            }

            modalList.appendChild(item);
        });
    }

    if (modalFilter) {
        modalFilter.addEventListener('change', () => {
            if (currentBlockData) renderFundList(currentBlockData.funds);
        });
    }

    const closeModal = () => {
        modal.classList.remove('open');
        resetHighlights();
    };

    if (closeBtn) closeBtn.addEventListener('click', closeModal);
    window.addEventListener('click', e => { if (e.target === modal) closeModal(); });

    // ── FINANCIAL INPUT LISTENERS (V6.0 TAUSENDERTRENNZEICHEN & 2 DEZIMALSTELLEN) ──
    if (totalInvestmentInput) {
        totalInvestmentInput.value = portfolioGlobals.totalInvestment > 0 ? formatNumberInput(portfolioGlobals.totalInvestment) : '';
        const onTotalInvCommit = (e) => {
            const val = parseCurrencyInput(e.target.value);
            portfolioGlobals.totalInvestment = val;
            totalInvestmentInput.value = val > 0 ? formatNumberInput(val) : '';
            saveGlobals();
            updatePortfolioUI();
        };
        totalInvestmentInput.addEventListener('change', onTotalInvCommit);
        totalInvestmentInput.addEventListener('blur', onTotalInvCommit);
        totalInvestmentInput.addEventListener('focus', (e) => {
            const val = portfolioGlobals.totalInvestment || 0;
            e.target.value = val > 0 ? (Math.round(val * 100) / 100).toFixed(2).replace('.', ',') : '';
            e.target.select();
        });
        totalInvestmentInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') e.target.blur();
        });
    }

    if (totalInvestmentSparrate) {
        totalInvestmentSparrate.value = portfolioGlobals.totalSparrate > 0 ? formatNumberInput(portfolioGlobals.totalSparrate) : '';
        const onTotalSparCommit = (e) => {
            const val = parseCurrencyInput(e.target.value);
            portfolioGlobals.totalSparrate = val;
            totalInvestmentSparrate.value = val > 0 ? formatNumberInput(val) : '';
            saveGlobals();
            updatePortfolioUI();
        };
        totalInvestmentSparrate.addEventListener('change', onTotalSparCommit);
        totalInvestmentSparrate.addEventListener('blur', onTotalSparCommit);
        totalInvestmentSparrate.addEventListener('focus', (e) => {
            const val = portfolioGlobals.totalSparrate || 0;
            e.target.value = val > 0 ? (Math.round(val * 100) / 100).toFixed(2).replace('.', ',') : '';
            e.target.select();
        });
        totalInvestmentSparrate.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') e.target.blur();
        });
    }

    // ── LÄNDERGEWICHTUNGEN & SCHICHT-BREAKDOWN ENGINE ──────
    function updateCountryBreakdown() {
        const infoBar = document.getElementById('portfolio-info-bar');
        const summaryEinmal = document.getElementById('portfolio-country-summary');
        const summarySpar = document.getElementById('portfolio-country-summary-sparrate');
        const clusterListEinmal = document.getElementById('portfolio-cluster-list');
        const clusterListSpar = document.getElementById('portfolio-cluster-list-sparrate');
        const layerBreakdownList = document.getElementById('layer-breakdown-list');

        if (!infoBar) return;

        let totalE = 0;
        let totalS = 0;
        portfolio.forEach(l => {
            l.funds.forEach(f => {
                const k = `${l.blockId}::${f.name}`;
                totalE += (portfolioGlobals.fundInvestments[k] || 0);
                totalS += (portfolioGlobals.fundSparrates[k] || 0);
            });
        });

        const totalInvestedGlobal = (portfolioGlobals.totalInvestment || 0);
        const totalSparGlobal = (portfolioGlobals.totalSparrate || 0);
        if (totalE === 0 && totalS === 0 && totalInvestedGlobal === 0 && totalSparGlobal === 0) {
            infoBar.style.display = 'none';
            return;
        }

        infoBar.style.display = 'grid';

        // 1. Verteilung nach Managementansatz (Kanonische Sortierung aller 8 Ansätze, unbesetzte mit 0,00 €)
        if (layerBreakdownList) {
            layerBreakdownList.innerHTML = '';
            managementBlocks.forEach(b => {
                const l = portfolio.find(item => item.blockId === b.id);
                let lE = 0; let lS = 0;
                if (l && l.funds) {
                    l.funds.forEach(f => {
                        const k = `${l.blockId}::${f.name}`;
                        const kAlt = `${l.blockId}_${f.name}`;
                        lE += (portfolioGlobals.fundInvestments[k] || portfolioGlobals.fundInvestments[kAlt] || 0);
                        lS += (portfolioGlobals.fundSparrates[k] || portfolioGlobals.fundSparrates[kAlt] || 0);
                    });
                }
                const eKey = `${b.id}::empfehlungsliste`;
                lE += (portfolioGlobals.fundInvestments[eKey] || 0);
                lS += (portfolioGlobals.fundSparrates[eKey] || 0);

                const color = layerColors[b.id] || '#2563eb';
                const displayTitle = getManagementBlockDisplayTitle(b.id);
                const isZero = (lE === 0 && lS === 0);

                const li = document.createElement('li');
                li.style.cssText = `display:flex; justify-content:space-between; align-items:center; margin-bottom:8px; padding:6px 10px; background:${isZero ? '#f8fafc' : '#ffffff'}; border:1px solid #e2e8f0; border-left:4px solid ${color}; border-radius:4px; font-size:12.5px; font-weight:600;`;
                li.innerHTML = `
                    <span style="color:${isZero ? '#64748b' : '#1e293b'};">${displayTitle}</span>
                    <span style="display:flex; flex-direction:column; align-items:flex-end; gap:2px; text-align:right;">
                        <span style="color:${isZero ? '#94a3b8' : '#0f172a'};">${formatCurrency(lE)}</span>
                        ${lS > 0 ? `<span style="font-size:11px; color:#15803d; font-weight:600;">+ ${formatCurrency(lS)} mtl.</span>` : ''}
                    </span>`;
                layerBreakdownList.appendChild(li);
            });
        }

        // 2. Compute Clusters for Einmal & Sparrate
        const computeData = (isSpar) => {
            let countryVals = {};
            let countryValsAktien = {};
            let countryValsAnleihen = {};
            let sum = 0;

            portfolio.forEach(layer => {
                layer.funds.forEach(fund => {
                    const k = `${layer.blockId}::${fund.name}`;
                    const val = isSpar ? (portfolioGlobals.fundSparrates[k] || 0) : (portfolioGlobals.fundInvestments[k] || 0);
                    if (val > 0) {
                        sum += val;
                        const blockData = managementBlocks.find(b => b.id === layer.blockId);
                        if (blockData) {
                            const fundData = blockData.funds.find(f => f.name === fund.name);
                            if (fundData && fundData.countryWeightings) {
                                const fClass = classifyFund(fundData.type);
                                fundData.countryWeightings.forEach(cw => {
                                    const v = val * (cw.weight / 100);
                                    const c = cw.country;
                                    countryVals[c] = (countryVals[c] || 0) + v;
                                    if (fClass === 'aktien') countryValsAktien[c] = (countryValsAktien[c] || 0) + v;
                                    else countryValsAnleihen[c] = (countryValsAnleihen[c] || 0) + v;
                                });
                            }
                        }
                    }
                });
            });
            return { countryVals, countryValsAktien, countryValsAnleihen, sum };
        };

        const renderClusterPills = (containerEl, listEl, data, label, type) => {
            const isSpar = (type === 'sparrate');
            const set = isSpar ? collapsedClustersSparrate : collapsedClustersEinmal;
            const storageKey = isSpar ? 'portfolio_collapsed_clusters_sparrate_v6' : 'portfolio_collapsed_clusters_einmal_v6';

            const { countryVals, sum } = data;
            if (sum <= 0 || Object.keys(countryVals).length === 0) {
                listEl.innerHTML = `<div style="color:#888; font-size:13px; text-align:center; padding:10px 0;">Keine ${label} verplant</div>`;
                if (isSpar) activeClustersSparrate = [];
                else activeClustersEinmal = [];
                updateToggleAllClustersBtn(type);
                return;
            }

            const clusterVals = {};
            const clusterDetails = {};
            CLUSTER_ORDER.forEach(k => { clusterVals[k] = 0; clusterDetails[k] = {}; });

            Object.keys(countryVals).forEach(country => {
                const val = countryVals[country];
                let assigned = false;
                for (const key of CLUSTER_ORDER) {
                    if (CLUSTER_DEFS[key].countries.has(country.trim())) {
                        clusterVals[key] += val;
                        clusterDetails[key][country] = (clusterDetails[key][country] || 0) + val;
                        assigned = true;
                        break;
                    }
                }
                if (!assigned) {
                    clusterVals['Sonstige'] += val;
                    clusterDetails['Sonstige'][country] = (clusterDetails['Sonstige'][country] || 0) + val;
                }
            });

            listEl.innerHTML = '';
            const fmt = n => n.toFixed(1).replace('.', ',') + '%';
            const activeKeys = [];

            CLUSTER_ORDER.forEach(cKey => {
                const pVal = clusterVals[cKey];
                if (pVal <= 0) return;
                activeKeys.push(cKey);
                const pct = (pVal / sum) * 100;
                const color = CLUSTER_DEFS[cKey].color;

                const topCountries = Object.entries(clusterDetails[cKey])
                    .map(([c, v]) => ({ c, p: (v / sum) * 100 }))
                    .sort((a, b) => b.p - a.p).slice(0, 5);

                const isCollapsed = set.has(cKey);

                const pill = document.createElement('div');
                pill.className = 'cluster-pill-item' + (isCollapsed ? ' is-collapsed' : '');
                pill.style.borderLeft = `4px solid ${color}`;
                pill.innerHTML = `
                    <div class="cluster-pill-header" data-cluster="${cKey}" style="cursor:pointer;" title="${isCollapsed ? 'Klicken zum Aufklappen' : 'Klicken zum Zuklappen'}" aria-expanded="${isCollapsed ? 'false' : 'true'}">
                        <div style="display:flex; align-items:center; gap:8px;">
                            <svg class="cluster-collapse-chevron" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="${color}" stroke-width="2.8" stroke-linecap="round" stroke-linejoin="round">
                                <polyline points="9 18 15 12 9 6"></polyline>
                            </svg>
                            <span class="cluster-pill-title" style="font-weight:700;">${cKey}</span>
                        </div>
                        <span class="cluster-pill-pct" style="font-weight:700;">${fmt(pct)}</span>
                    </div>
                    <ul class="cluster-top-countries">
                        ${topCountries.map(d => `<li style="display:flex; justify-content:space-between; margin-top:2px;"><span>${d.c}</span> <span>${fmt(d.p)}</span></li>`).join('')}
                    </ul>`;

                const headerEl = pill.querySelector('.cluster-pill-header');
                if (headerEl) {
                    headerEl.addEventListener('click', () => {
                        const nowCollapsed = !pill.classList.contains('is-collapsed');
                        if (nowCollapsed) {
                            pill.classList.add('is-collapsed');
                            set.add(cKey);
                            headerEl.setAttribute('title', 'Klicken zum Aufklappen');
                            headerEl.setAttribute('aria-expanded', 'false');
                        } else {
                            pill.classList.remove('is-collapsed');
                            set.delete(cKey);
                            headerEl.setAttribute('title', 'Klicken zum Zuklappen');
                            headerEl.setAttribute('aria-expanded', 'true');
                        }
                        try {
                            localStorage.setItem(storageKey, JSON.stringify([...set]));
                        } catch (err) {}
                        updateToggleAllClustersBtn(type);
                    });
                }

                listEl.appendChild(pill);
            });

            if (isSpar) activeClustersSparrate = activeKeys;
            else activeClustersEinmal = activeKeys;

            updateToggleAllClustersBtn(type);
        };

        if (summaryEinmal && clusterListEinmal) {
            renderClusterPills(summaryEinmal, clusterListEinmal, computeData(false), 'Einmalbeiträge', 'einmal');
        }
        if (summarySpar && clusterListSpar) {
            renderClusterPills(summarySpar, clusterListSpar, computeData(true), 'Sparraten', 'sparrate');
        }
    }

    function updatePortfolioUI() {
        renderGridColumns();
        renderStandaloneCylinders();
        updateCountryBreakdown();

        let sumEinmal = 0;
        let sumSpar = 0;

        const countedEinmal = new Set();
        Object.entries(portfolioGlobals.fundInvestments).forEach(([k, v]) => {
            const num = Number(v) || 0;
            if (num > 0) {
                const normKey = k.includes('::') ? k : k.replace('_', '::');
                if (!countedEinmal.has(normKey)) {
                    countedEinmal.add(normKey);
                    sumEinmal += num;
                }
            }
        });

        const countedSpar = new Set();
        Object.entries(portfolioGlobals.fundSparrates).forEach(([k, v]) => {
            const num = Number(v) || 0;
            if (num > 0) {
                const normKey = k.includes('::') ? k : k.replace('_', '::');
                if (!countedSpar.has(normKey)) {
                    countedSpar.add(normKey);
                    sumSpar += num;
                }
            }
        });

        const totalDist = document.getElementById('total-distributed-display') || totalDistributedDisplay;
        const totalRem = document.getElementById('total-remaining-display') || totalRemainingDisplay;
        const sparDist = document.getElementById('sparrate-distributed-display') || sparrateDistributedDisplay;
        const sparRem = document.getElementById('sparrate-remaining-display') || sparrateRemainingDisplay;

        if (totalDist) totalDist.textContent = formatCurrency(sumEinmal);
        if (totalRem) {
            const remE = (portfolioGlobals.totalInvestment || 0) - sumEinmal;
            totalRem.textContent = formatCurrency(remE);
            totalRem.style.color = remE < 0 ? '#ef4444' : '#10b981';
        }

        if (sparDist) sparDist.textContent = formatCurrency(sumSpar) + ' mtl.';
        if (sparRem) {
            const remS = (portfolioGlobals.totalSparrate || 0) - sumSpar;
            sparRem.textContent = formatCurrency(remS) + ' mtl.';
            sparRem.style.color = remS < 0 ? '#ef4444' : '#10b981';
        }

        // Update Allocation Summary Progress Bar (V4.4)
        const allocFill = document.getElementById('alloc-bar-fill');
        const allocTotal = document.getElementById('alloc-total');
        if (allocFill && allocTotal) {
            const totInv = portfolioGlobals.totalInvestment || 0;
            const pct = totInv > 0 ? (sumEinmal / totInv) * 100 : 0;
            allocTotal.textContent = pct.toFixed(1).replace('.', ',');
            allocFill.style.width = Math.min(pct, 100) + '%';
            if (Math.abs(pct - 100) < 0.1) {
                allocFill.style.backgroundColor = '#10b981';
            } else if (pct > 100) {
                allocFill.style.backgroundColor = '#ef4444';
            } else {
                allocFill.style.backgroundColor = '#3b82f6';
            }
        }

        const pLayers = document.getElementById('panel-layers') || panelLayers;
        const pEmpty = document.getElementById('panel-empty') || panelEmpty;
        const pHidden = document.getElementById('panel-hidden-state') || panelHiddenState;

        if (managersHidden) {
            if (pLayers) pLayers.style.display = 'none';
            if (pEmpty) pEmpty.style.display = 'none';
            if (pHidden) pHidden.style.display = 'block';
        } else {
            if (pHidden) pHidden.style.display = 'none';
            if (pEmpty) pEmpty.style.display = 'none';
            if (pLayers) {
                pLayers.style.display = 'block';
                renderRightPanelLayers();
            }
        }
    }
    window.updatePortfolioUI = updatePortfolioUI;
    window.renderGridColumns = renderGridColumns;
    window.renderStandaloneCylinders = renderStandaloneCylinders;
    window.refreshModalButtons = refreshModalButtons;

    function renderRightPanelLayers() {
        if (!panelLayers) return;
        panelLayers.innerHTML = '';

        let totalOverallZuzahlung = 0;
        let totalOverallSparNeu = 0;
        const hasBaseline = Boolean(portfolioGlobals.initialInvestments && Object.keys(portfolioGlobals.initialInvestments).length > 0);

        managementBlocks.forEach(block => {
            const layer = portfolio.find(l => l.blockId === block.id);
            const color = layerColors[block.id] || '#2563eb';
            const displayTitle = getManagementBlockDisplayTitle(block.id);
            const isCollapsed = collapsedLayers.has(block.id);

            const div = document.createElement('div');
            div.className = 'panel-layer-card' + (isCollapsed ? ' is-collapsed' : '');
            div.dataset.blockId = block.id;
            div.style.borderLeft = `4px solid ${color}`;

            let layerEinmal = 0;
            let layerSpar = 0;
            let layerZuzahlung = 0;

            const layerEl = document.createElement('div');

            // ── Sortierung: Aktive Fonds A→Z, dann delistete Fonds A→Z ──
            const regularFunds  = layer ? layer.funds.filter(f => !f._isEmpfehlungsliste && !f._isDelistet).sort((a, b) => a.name.localeCompare(b.name, 'de')) : [];
            const delistedFunds = layer ? layer.funds.filter(f => !f._isEmpfehlungsliste &&  f._isDelistet).sort((a, b) => a.name.localeCompare(b.name, 'de')) : [];
            const sortedLayerFunds = [...regularFunds, ...delistedFunds];

            let separatorAdded = false;
            sortedLayerFunds.forEach(f => {
                const k = `${block.id}::${f.name}`;
                const kAlt = `${block.id}_${f.name}`;
                const eVal = portfolioGlobals.fundInvestments[k] || portfolioGlobals.fundInvestments[kAlt] || 0;
                const sVal = portfolioGlobals.fundSparrates[k] || portfolioGlobals.fundSparrates[kAlt] || 0;
                const initialVal = portfolioGlobals.initialInvestments ? (portfolioGlobals.initialInvestments[k] || portfolioGlobals.initialInvestments[kAlt] || 0) : 0;
                const initialSpar = portfolioGlobals.initialSparrates ? (portfolioGlobals.initialSparrates[k] || portfolioGlobals.initialSparrates[kAlt] || 0) : 0;

                let zuzahlung = 0;
                let isNeuanlage = false;
                if (hasBaseline) {
                    if (initialVal === 0 && eVal > 0) {
                        zuzahlung = eVal;
                        isNeuanlage = true;
                    } else if (initialVal > 0 && eVal > initialVal) {
                        zuzahlung = Math.round((eVal - initialVal) * 100) / 100;
                        isNeuanlage = false;
                    }
                }

                let sparrateNeu = 0;
                let isSparrateNeuanlage = false;
                if (hasBaseline) {
                    if (initialSpar === 0 && sVal > 0) {
                        sparrateNeu = sVal;
                        isSparrateNeuanlage = true;
                    } else if (initialSpar > 0 && sVal > initialSpar) {
                        sparrateNeu = Math.round((sVal - initialSpar) * 100) / 100;
                        isSparrateNeuanlage = false;
                    }
                }

                layerEinmal += eVal;
                layerSpar += sVal;
                if (zuzahlung > 0) {
                    layerZuzahlung += zuzahlung;
                    totalOverallZuzahlung += zuzahlung;
                }
                if (sparrateNeu > 0) {
                    totalOverallSparNeu += sparrateNeu;
                }

                const ertragTxt = f.ertrag ? ` · ${f.ertrag}` : '';
                const fWknMatch = (f.info || '').match(/WKN:\s*([A-Z0-9]{6})/i);
                const fWkn = fWknMatch ? fWknMatch[1] : '';

                const badgeLabel = isNeuanlage ? `+ ${formatCurrency(zuzahlung)} Neuanlage` : `+ ${formatCurrency(zuzahlung)} Zuzahlung`;
                const sparrateBadgeLabel = isSparrateNeuanlage ? `+ ${formatCurrency(sparrateNeu)} mtl. neu` : `+ ${formatCurrency(sparrateNeu)} mtl. Erhöhung`;

                let metaContent = '';
                if (zuzahlung > 0) {
                    metaContent = `
                        <div style="display:flex; gap:6px; align-items:center; flex-wrap:wrap; margin-top:3px;">
                            <span class="zuzahlung-badge" style="background:#ecfdf5; color:#15803d; border:1px solid #86efac; border-radius:4px; padding:1px 6px; font-size:10px; font-weight:700;">${badgeLabel}</span>
                            <span style="color:#64748b; font-size:10.5px;">Bestand: <strong>${formatCurrency(initialVal)}</strong> · Gesamt: <strong style="color:var(--color-mlp-blau);">${formatCurrency(eVal)}</strong></span>
                            ${sparrateNeu > 0 ? `<span class="sparrate-badge" style="background:#ecfdf5; color:#15803d; border:1px solid #86efac; border-radius:4px; padding:1px 6px; font-size:10px; font-weight:700;">${sparrateBadgeLabel}</span>` : (sVal > 0 ? `<span style="color:#15803d; font-weight:600;">(${formatCurrency(sVal)} mtl.)</span>` : '')}
                        </div>`;
                } else if (hasBaseline && initialVal > 0) {
                    metaContent = `
                        <div style="display:flex; gap:6px; align-items:center; flex-wrap:wrap; margin-top:3px;">
                            <span style="color:#64748b; font-size:10.5px;">Bestand: <strong>${formatCurrency(initialVal)}</strong></span>
                            ${sparrateNeu > 0 ? `<span class="sparrate-badge" style="background:#ecfdf5; color:#15803d; border:1px solid #86efac; border-radius:4px; padding:1px 6px; font-size:10px; font-weight:700;">${sparrateBadgeLabel}</span>` : (sVal > 0 ? `<span style="color:#15803d; font-weight:600;">(${formatCurrency(sVal)} mtl.)</span>` : '')}
                        </div>`;
                } else {
                    metaContent = `${(eVal > 0 || sVal > 0) ? ` · <strong style="color:var(--color-mlp-blau);">${eVal > 0 ? formatCurrency(eVal) : ''} ${sVal > 0 ? `(${formatCurrency(sVal)} mtl.)` : ''}</strong>` : ''}`;
                    if (sparrateNeu > 0) {
                        metaContent += ` <span class="sparrate-badge" style="background:#ecfdf5; color:#15803d; border:1px solid #86efac; border-radius:4px; padding:1px 6px; font-size:10px; font-weight:700;">${sparrateBadgeLabel}</span>`;
                    }
                }

                const item = document.createElement('div');
                if (f._isDelistet) {
                    // ── Trennlinie vor erstem delisteten Fonds ──
                    if (!separatorAdded && regularFunds.length > 0) {
                        const sep = document.createElement('hr');
                        sep.className = 'panel-fund-delisted-separator';
                        layerEl.appendChild(sep);
                        separatorAdded = true;
                    }
                    // ── Dezentes, gedimmtes Styling für delistete Fonds ──
                    item.className = 'panel-fund-item panel-fund-delisted-card';
                    item.innerHTML = `
                        <div class="panel-fund-info" style="flex:1; padding-right:8px;">
                            <div class="panel-fund-name" style="font-weight:500; font-size:12px; color:#64748b; line-height:1.3;" title="${f.name}">
                                ${f.name}${fWkn ? ` · WKN ${fWkn}` : ''}
                            </div>
                            <div class="panel-fund-meta" style="font-size:11px; color:#94a3b8; margin-top:2px;">
                                ${f.anlageschwerpunkt || f.type || 'Fonds'}
                                ${metaContent}
                            </div>
                        </div>
                        <button class="panel-fund-remove" title="Entfernen" data-block="${block.id}" data-fund="${f.name}"
                            style="background:none; border:none; color:#94a3b8; font-size:14px; font-weight:bold; cursor:pointer; padding:2px 6px; border-radius:4px; transition:background 0.15s;">
                            ✕
                        </button>`;
                } else {
                    item.className = 'panel-fund-item';
                    item.style.cssText = 'display:flex; justify-content:space-between; align-items:center; padding:6px 0; border-top:1px dashed #e2e8f0;';
                    item.innerHTML = `
                        <div class="panel-fund-info" style="flex:1; padding-right:8px;">
                            <div class="panel-fund-name" style="font-weight:600; font-size:13px; color:#1e293b;" title="${f.name}">${f.name}</div>
                            <div class="panel-fund-meta" style="font-size:11px; color:#64748b; margin-top:2px;">
                                ${f.info || ''}${ertragTxt}
                                ${metaContent}
                            </div>
                        </div>
                        <button class="panel-fund-remove" title="Entfernen" data-block="${block.id}" data-fund="${f.name}"
                            style="background:none; border:none; color:#ef4444; font-size:14px; font-weight:bold; cursor:pointer; padding:2px 6px; border-radius:4px; transition:background 0.15s;">
                            ✕
                        </button>`;
                }
                layerEl.appendChild(item);
            });

            // ── Empfehlungsliste für diesen Block ──
            const eKey = `${block.id}::empfehlungsliste`;
            const eInv = portfolioGlobals.fundInvestments[eKey] || 0;
            const eSpar = portfolioGlobals.fundSparrates[eKey] || 0;
            const eInit = portfolioGlobals.initialInvestments ? (portfolioGlobals.initialInvestments[eKey] || 0) : 0;
            const eInitSpar = portfolioGlobals.initialSparrates ? (portfolioGlobals.initialSparrates[eKey] || 0) : 0;

            let empFondsForBlock = [];
            try {
                const allEmpFonds = JSON.parse(localStorage.getItem('empfehlungslisteFonds_V44') || '{}');
                empFondsForBlock = allEmpFonds[block.id] || [];
            } catch { empFondsForBlock = []; }

            if (eInv > 0 || eSpar > 0 || empFondsForBlock.length > 0) {
                let eZuz = 0;
                let eIsNeu = false;
                if (hasBaseline) {
                    if (eInit === 0 && eInv > 0) {
                        eZuz = eInv;
                        eIsNeu = true;
                    } else if (eInit > 0 && eInv > eInit) {
                        eZuz = Math.round((eInv - eInit) * 100) / 100;
                        eIsNeu = false;
                    }
                }

                let eSparNeu = 0;
                if (hasBaseline) {
                    if (eInitSpar === 0 && eSpar > 0) {
                        eSparNeu = eSpar;
                    } else if (eInitSpar > 0 && eSpar > eInitSpar) {
                        eSparNeu = Math.round((eSpar - eInitSpar) * 100) / 100;
                    }
                }

                layerEinmal += eInv;
                layerSpar += eSpar;
                if (eZuz > 0) {
                    layerZuzahlung += eZuz;
                    totalOverallZuzahlung += eZuz;
                }
                if (eSparNeu > 0) {
                    totalOverallSparNeu += eSparNeu;
                }

                const eRow = document.createElement('div');
                eRow.className = 'panel-fund-item panel-empfehlung-row';
                const eInvFmt = eInv > 0 ? formatNumberInput(eInv) : '';
                const eSparFmt = eSpar > 0 ? formatNumberInput(eSpar) : '';

                let popupHtml = '';
                if (empFondsForBlock.length > 0) {
                    const popupItems = empFondsForBlock.map(u =>
                        `<span><b>${u.name}</b> · WKN ${u.wkn} · ${u.schwerpunkt || '?'} · ${formatCurrency(u.einmal || 0)}${u.sparrate > 0 ? ' + ' + formatCurrency(u.sparrate) + ' mtl.' : ''}</span>`
                    ).join('');
                    popupHtml = `<div class="empfehlung-popup">${popupItems}</div>`;
                }

                const eBadge = eZuz > 0 ? `<span class="zuzahlung-badge" style="background:#ecfdf5; color:#15803d; border:1px solid #86efac; border-radius:4px; padding:1px 6px; font-size:10px; font-weight:700; margin-left:6px;">+ ${formatCurrency(eZuz)} ${eIsNeu ? 'Neuanlage' : 'Zuzahlung'}</span>` : '';
                const eSparBadge = eSparNeu > 0 ? `<span class="sparrate-badge" style="background:#ecfdf5; color:#15803d; border:1px solid #86efac; border-radius:4px; padding:1px 6px; font-size:10px; font-weight:700; margin-left:6px;">+ ${formatCurrency(eSparNeu)} mtl. neu</span>` : '';

                eRow.innerHTML = `
                    <div class="panel-fund-info" style="width:100%;">
                        <div class="panel-fund-name empfehlung-label empfehlung-label-hoverable" title="Fonds, die nicht mehr auf der VEM-Empfehlungsliste stehen" style="position:relative; cursor:${empFondsForBlock.length > 0 ? 'help' : 'default'}; font-weight:700;">
                            <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#d97706" stroke-width="2.5" style="vertical-align:-2px; margin-right:4px;"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>
                            Von Empfehlungsliste genommen
                            ${eBadge}
                            ${eSparBadge}
                            ${popupHtml}
                        </div>
                        <div class="panel-fund-meta" style="display:flex; gap:8px; align-items:center; flex-wrap:wrap; margin-top:6px;">
                            <input class="fund-investment-input empfehlung-input" type="text" inputmode="decimal"
                                placeholder="Einmal €" title="Einmalbetrag"
                                data-ekey="${eKey}" data-spar="0"
                                value="${eInvFmt}" style="width:100px;">
                            <input class="fund-investment-input empfehlung-input" type="text" inputmode="decimal"
                                placeholder="Sparrate €" title="Sparrate monatlich"
                                data-ekey="${eKey}" data-spar="1"
                                value="${eSparFmt}" style="width:100px;">
                        </div>
                    </div>`;

                eRow.querySelectorAll('.empfehlung-input').forEach(inp => {
                    const commitEmpfehlung = () => {
                        const v = parseCurrencyInput(inp.value);
                        const k = inp.dataset.ekey;
                        const isSpar = inp.dataset.spar === '1';
                        inp.value = v > 0 ? formatNumberInput(v) : '';
                        if (isSpar) {
                            if (v > 0) portfolioGlobals.fundSparrates[k] = v;
                            else delete portfolioGlobals.fundSparrates[k];
                        } else {
                            if (v > 0) portfolioGlobals.fundInvestments[k] = v;
                            else delete portfolioGlobals.fundInvestments[k];
                        }
                        saveGlobals();
                        updatePortfolioUI();
                    };
                    inp.addEventListener('change', commitEmpfehlung);
                    inp.addEventListener('blur', commitEmpfehlung);
                    inp.addEventListener('focus', () => {
                        const v = parseCurrencyInput(inp.value);
                        inp.value = v > 0 ? (Math.round(v * 100) / 100).toFixed(2).replace('.', ',') : '';
                        inp.select();
                    });
                    inp.addEventListener('keydown', (e) => { if (e.key === 'Enter') inp.blur(); });
                });

                layerEl.appendChild(eRow);
            }

            // ── Wenn keine Fonds in diesem Managementansatz: Zeige leeren Zustand mit Button ──
            const hasAnyFunds = (layerEl.children.length > 0);
            if (!hasAnyFunds) {
                const emptyItem = document.createElement('div');
                emptyItem.className = 'panel-fund-empty-item';
                emptyItem.style.cssText = 'padding:6px 0; border-top:1px dashed #e2e8f0; color:#94a3b8; font-size:12px; font-style:italic; display:flex; justify-content:space-between; align-items:center;';
                emptyItem.innerHTML = `
                    <span>Keine Manager ausgewählt (0,00 €)</span>
                    <button type="button" class="btn-select-block" data-block="${block.id}" style="background:#f1f5f9; border:1px solid #cbd5e1; border-radius:4px; font-size:11px; padding:2px 8px; cursor:pointer; color:#033d5d; font-weight:600;">+ Auswählen</button>
                `;
                layerEl.appendChild(emptyItem);
            } else {
                const addMoreWrapper = document.createElement('div');
                addMoreWrapper.className = 'panel-add-more-wrapper';
                addMoreWrapper.style.cssText = 'padding-top:6px; margin-top:4px; border-top:1px dashed #e2e8f0; display:flex; justify-content:flex-end;';
                addMoreWrapper.innerHTML = `
                    <button type="button" class="btn-select-more-funds" data-block="${block.id}" style="background:#f8fafc; border:1px solid #cbd5e1; border-radius:4px; font-size:11px; padding:3px 8px; cursor:pointer; color:#033d5d; font-weight:600; display:inline-flex; align-items:center; gap:4px; transition:background-color 0.15s;">
                        <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="5" x2="12" y2="19"></line><line x1="5" y1="12" x2="19" y2="12"></line></svg>
                        <span>+ Weitere Manager auswählen</span>
                    </button>
                `;
                layerEl.appendChild(addMoreWrapper);
            }

            div.innerHTML = `
                <div class="panel-layer-header" data-block="${block.id}" style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:6px; cursor:pointer;" title="${isCollapsed ? 'Klicken zum Aufklappen' : 'Klicken zum Zuklappen'}" aria-expanded="${isCollapsed ? 'false' : 'true'}">
                    <div style="display:flex; align-items:center; gap:8px;">
                        <svg class="collapse-chevron" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="${color}" stroke-width="2.8" stroke-linecap="round" stroke-linejoin="round">
                            <polyline points="9 18 15 12 9 6"></polyline>
                        </svg>
                        <div class="panel-layer-title-text" style="color:${color};">
                            ${formatManagementBlockHeaderTitle(block.id)}
                        </div>
                    </div>
                    <span style="display:flex; flex-direction:column; align-items:flex-end; gap:2px; text-align:right;">
                        <span style="font-weight:bold; font-size:12px; color:${hasAnyFunds && (layerEinmal > 0 || layerSpar > 0) ? '#334155' : '#94a3b8'};">
                            ${hasAnyFunds && (layerEinmal > 0 || layerSpar > 0) ? `${layerEinmal > 0 ? formatCurrency(layerEinmal) : ''} ${layerSpar > 0 ? `(${formatCurrency(layerSpar)} mtl.)` : ''}` : '0,00 €'}
                        </span>
                        ${layerZuzahlung > 0 ? `
                            <div class="panel-layer-zuzahlung-badge">
                                <span class="panel-layer-zuzahlung-amount">+&nbsp;${formatCurrency(layerZuzahlung)}</span>
                                <span class="panel-layer-zuzahlung-label">Zuzahlung/Neuanlage</span>
                            </div>` : ''}
                    </span>
                </div>`;

            const listWrapper = document.createElement('div');
            listWrapper.className = 'panel-funds-list';
            listWrapper.appendChild(layerEl);
            div.appendChild(listWrapper);

            // Accordion toggle on header
            const headerEl = div.querySelector('.panel-layer-header');
            if (headerEl) {
                headerEl.addEventListener('click', (e) => {
                    if (e.target.closest('button') || e.target.closest('input')) return;
                    const nowCollapsed = !div.classList.contains('is-collapsed');
                    if (nowCollapsed) {
                        div.classList.add('is-collapsed');
                        collapsedLayers.add(block.id);
                        headerEl.setAttribute('title', 'Klicken zum Aufklappen');
                        headerEl.setAttribute('aria-expanded', 'false');
                    } else {
                        div.classList.remove('is-collapsed');
                        collapsedLayers.delete(block.id);
                        headerEl.setAttribute('title', 'Klicken zum Zuklappen');
                        headerEl.setAttribute('aria-expanded', 'true');
                    }
                    try {
                        localStorage.setItem('portfolio_collapsed_layers_v6', JSON.stringify([...collapsedLayers]));
                    } catch (err) {}
                    if (typeof updateToggleAllCollapseBtn === 'function') {
                        updateToggleAllCollapseBtn();
                    }
                });
            }

            // Click listener for selecting / managing funds
            div.querySelectorAll('.btn-select-block, .btn-select-more-funds').forEach(clickable => {
                clickable.addEventListener('click', (e) => {
                    e.stopPropagation();
                    const bId = clickable.dataset.block || block.id;
                    const blk = managementBlocks.find(b => b.id === bId);
                    if (blk && typeof openFundModalForBlock === 'function') {
                        openFundModalForBlock(blk);
                    }
                });
            });

            div.querySelectorAll('.panel-fund-remove').forEach(btn => {
                btn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    const bId = btn.dataset.block;
                    const fName = btn.dataset.fund;
                    removeFund(bId, fName);
                    if (currentBlockData && modal && modal.classList.contains('open')) {
                        renderFundList(currentBlockData.funds);
                    }
                    updatePortfolioUI();
                });
            });

            panelLayers.appendChild(div);
        });

        // Top summary for right panel
        const zuzahlungSummaryEl = document.getElementById('zuzahlung-panel-summary');
        const zuzahlungValEl = document.getElementById('zuzahlung-panel-val');
        if (zuzahlungSummaryEl) {
            if (totalOverallZuzahlung > 0 || totalOverallSparNeu > 0) {
                zuzahlungSummaryEl.style.display = 'flex';
                let txt = `+ ${formatCurrency(totalOverallZuzahlung)}`;
                if (totalOverallSparNeu > 0) {
                    txt += ` (+ ${formatCurrency(totalOverallSparNeu)} mtl. neu)`;
                }
                if (zuzahlungValEl) zuzahlungValEl.textContent = txt;
            } else {
                zuzahlungSummaryEl.style.display = 'none';
            }
        }

        updateToggleAllCollapseBtn();
    }

    function updateToggleAllCollapseBtn() {
        const btn = document.getElementById('toggle-all-collapse-btn');
        if (!btn || typeof managementBlocks === 'undefined') return;
        const allBlockIds = managementBlocks.map(b => b.id);
        const allCollapsed = allBlockIds.length > 0 && allBlockIds.every(id => collapsedLayers.has(id));
        btn.title = allCollapsed ? 'Alle Zeitphasen aufklappen' : 'Alle Zeitphasen zuklappen';
        btn.setAttribute('aria-label', btn.title);
    }

    function updateToggleAllClustersBtn(type) {
        const isSpar = (type === 'sparrate');
        const btn = document.getElementById(isSpar ? 'toggle-all-clusters-sparrate-btn' : 'toggle-all-clusters-einmal-btn');
        const set = isSpar ? collapsedClustersSparrate : collapsedClustersEinmal;
        const activeKeys = isSpar ? activeClustersSparrate : activeClustersEinmal;

        if (!btn) return;
        if (!activeKeys || activeKeys.length === 0) {
            btn.style.display = 'none';
            return;
        }
        btn.style.display = 'flex';
        const allCollapsed = activeKeys.length > 0 && activeKeys.every(k => set.has(k));
        btn.title = allCollapsed ? 'Alle Länder aufklappen' : 'Alle Länder zuklappen';
        btn.setAttribute('aria-label', btn.title);
    }

    function exportProposalPdf() {
        if (!window.jspdf) {
            alert('PDF-Bibliothek ist noch nicht geladen.');
            return;
        }
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF('p', 'mm', 'a4');
        const now = new Date();
        const dateStr = now.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' });
        const fileName = `Anlagevorschlag_V6.0_${dateStr.replace(/\./g, '-')}.pdf`;

        // Check if there are any Zuzahlungen / Neuanlagen
        const hasBaseline = Boolean(portfolioGlobals.initialInvestments && Object.keys(portfolioGlobals.initialInvestments).length > 0);
        let totalZuzahlung = 0;
        let totalBestand = 0;
        let totalEinmal = 0;
        let totalSpar = 0;
        let totalSparNeu = 0;
        let totalInitialSpar = 0;

        managementBlocks.forEach(block => {
            block.funds.forEach(fund => {
                const key = `${block.id}::${fund.name}`;
                const keyAlt = `${block.id}_${fund.name}`;
                const eVal = portfolioGlobals.fundInvestments[key] || portfolioGlobals.fundInvestments[keyAlt] || 0;
                const sVal = portfolioGlobals.fundSparrates[key] || portfolioGlobals.fundSparrates[keyAlt] || 0;
                const iVal = portfolioGlobals.initialInvestments ? (portfolioGlobals.initialInvestments[key] || portfolioGlobals.initialInvestments[keyAlt] || 0) : 0;
                const iSpar = portfolioGlobals.initialSparrates ? (portfolioGlobals.initialSparrates[key] || portfolioGlobals.initialSparrates[keyAlt] || 0) : 0;

                totalEinmal += eVal;
                totalSpar += sVal;
                totalBestand += iVal;
                totalInitialSpar += iSpar;

                if (hasBaseline) {
                    if (iVal === 0 && eVal > 0) {
                        totalZuzahlung += eVal;
                    } else if (iVal > 0 && eVal > iVal) {
                        totalZuzahlung += (eVal - iVal);
                    }
                    if (iSpar === 0 && sVal > 0) {
                        totalSparNeu += sVal;
                    } else if (iSpar > 0 && sVal > iSpar) {
                        totalSparNeu += (sVal - iSpar);
                    }
                }
            });
        });

        managementBlocks.forEach(block => {
            const eKey = `${block.id}::empfehlungsliste`;
            const eInv = portfolioGlobals.fundInvestments[eKey] || 0;
            const eSpar = portfolioGlobals.fundSparrates[eKey] || 0;
            const eInit = portfolioGlobals.initialInvestments ? (portfolioGlobals.initialInvestments[eKey] || 0) : 0;
            const eInitSpar = portfolioGlobals.initialSparrates ? (portfolioGlobals.initialSparrates[eKey] || 0) : 0;
            totalEinmal += eInv;
            totalSpar += eSpar;
            totalBestand += eInit;
            totalInitialSpar += eInitSpar;
            if (hasBaseline) {
                if (eInit === 0 && eInv > 0) {
                    totalZuzahlung += eInv;
                } else if (eInit > 0 && eInv > eInit) {
                    totalZuzahlung += (eInv - eInit);
                }
                if (eInitSpar === 0 && eSpar > 0) {
                    totalSparNeu += eSpar;
                } else if (eInitSpar > 0 && eSpar > eInitSpar) {
                    totalSparNeu += (eSpar - eInitSpar);
                }
            }
        });

        const hasZuzahlung = Boolean(hasBaseline && (totalZuzahlung > 0 || totalSparNeu > 0 || totalBestand > 0));

        // ── KOPFBEREICH ──
        doc.setFontSize(16);
        doc.setFont(undefined, 'bold');
        doc.setTextColor(3, 61, 93); // MLP Blau
        doc.text('Anlagevorschlag - Zeitphasenmodell V6.0', 14, 18);

        // Dekorative Trennlinie
        doc.setDrawColor(3, 61, 93);
        doc.setLineWidth(0.6);
        doc.line(14, 22, 196, 22);

        doc.setFontSize(9.5);
        doc.setFont(undefined, 'normal');
        doc.setTextColor(71, 85, 105);
        doc.text(`Datum: ${dateStr}`, 14, 29);

        let startYTable = 45;
        if (hasZuzahlung) {
            const bestandSparStr = totalInitialSpar > 0 ? ` (+ ${formatCurrency(totalInitialSpar)} mtl.)` : '';
            doc.text(`Bisheriger Depotbestand: ${formatCurrency(totalBestand)}${bestandSparStr}`, 14, 35);
            doc.setTextColor(21, 128, 61);
            doc.setFont(undefined, 'bold');
            const zuzahlungSparStr = totalSparNeu > 0 ? ` (+ ${formatCurrency(totalSparNeu)} mtl. neu)` : '';
            doc.text(`Geplante Zuzahlung / Neuanlage: + ${formatCurrency(totalZuzahlung)}${zuzahlungSparStr}`, 14, 41);
            doc.setFont(undefined, 'normal');
            doc.setTextColor(15, 23, 42);
            const gesamtSparStr = totalSpar > 0 ? ` (+ ${formatCurrency(totalSpar)} mtl.)` : '';
            doc.text(`Neues Gesamtvermögen: ${formatCurrency(portfolioGlobals.totalInvestment || totalEinmal)}${gesamtSparStr}`, 14, 47);
            startYTable = 53;
        } else {
            doc.setTextColor(15, 23, 42);
            doc.text(`Anzulegendes Gesamtvermögen: ${formatCurrency(portfolioGlobals.totalInvestment || totalEinmal)}` + (totalSpar > 0 ? ` (+ ${formatCurrency(totalSpar)} mtl.)` : ''), 14, 35);
            startYTable = 42;
        }

        const formatWknIsin = (rawInfo) => {
            if (!rawInfo) return '';
            return String(rawInfo)
                .replace(/\s*\/\s*ISIN:\s*/i, '\nISIN: ')
                .replace(/\s*·\s*ISIN:\s*/i, '\nISIN: ')
                .trim();
        };

        const tableData = [];
        let totalInvestedRows = 0;

        managementBlocks.forEach(block => {
            const blockFunds = block.funds.filter(fund => {
                const key = `${block.id}::${fund.name}`;
                const keyAlt = `${block.id}_${fund.name}`;
                const amountEinmal = portfolioGlobals.fundInvestments[key] || portfolioGlobals.fundInvestments[keyAlt] || 0;
                const amountSpar = portfolioGlobals.fundSparrates[key] || portfolioGlobals.fundSparrates[keyAlt] || 0;
                const initialVal = portfolioGlobals.initialInvestments ? (portfolioGlobals.initialInvestments[key] || portfolioGlobals.initialInvestments[keyAlt] || 0) : 0;
                return (amountEinmal > 0 || amountSpar > 0 || initialVal > 0);
            });

            // Aktive Fonds A-Z, delistete Fonds A-Z
            const regularFunds = blockFunds.filter(f => !f._isDelistet).sort((a, b) => a.name.localeCompare(b.name, 'de'));
            const delistedFunds = blockFunds.filter(f => f._isDelistet).sort((a, b) => a.name.localeCompare(b.name, 'de'));
            const sortedBlockFunds = [...regularFunds, ...delistedFunds];

            // Empfehlungsliste für diesen Block
            const eKey = `${block.id}::empfehlungsliste`;
            const eInv = portfolioGlobals.fundInvestments[eKey] || 0;
            const eSpar = portfolioGlobals.fundSparrates[eKey] || 0;
            const eInit = portfolioGlobals.initialInvestments ? (portfolioGlobals.initialInvestments[eKey] || 0) : 0;
            const eInitSpar = portfolioGlobals.initialSparrates ? (portfolioGlobals.initialSparrates[eKey] || 0) : 0;
            const hasEmpfehlung = (eInv > 0 || eSpar > 0 || eInit > 0);

            const displayTitle = getManagementBlockDisplayTitle(block.id);

            if (sortedBlockFunds.length === 0 && !hasEmpfehlung) {
                // Wenn 0 €: Eintrag mit 0,00 € erzeugen
                if (hasZuzahlung) {
                    tableData.push([displayTitle, '-', '0,00 €', '-', '0,00 €']);
                } else {
                    tableData.push([displayTitle, '-', '-', '0,00 €']);
                }
                return;
            }

            totalInvestedRows += (sortedBlockFunds.length + (hasEmpfehlung ? 1 : 0));

            sortedBlockFunds.forEach(fund => {
                const key = `${block.id}::${fund.name}`;
                const keyAlt = `${block.id}_${fund.name}`;
                const amountEinmal = portfolioGlobals.fundInvestments[key] || portfolioGlobals.fundInvestments[keyAlt] || 0;
                const amountSpar = portfolioGlobals.fundSparrates[key] || portfolioGlobals.fundSparrates[keyAlt] || 0;
                const initialVal = portfolioGlobals.initialInvestments ? (portfolioGlobals.initialInvestments[key] || portfolioGlobals.initialInvestments[keyAlt] || 0) : 0;
                const initialSpar = portfolioGlobals.initialSparrates ? (portfolioGlobals.initialSparrates[key] || portfolioGlobals.initialSparrates[keyAlt] || 0) : 0;

                let zuzahlung = 0;
                let isNeuanlage = false;
                if (hasBaseline) {
                    if (initialVal === 0 && amountEinmal > 0) {
                        zuzahlung = amountEinmal;
                        isNeuanlage = true;
                    } else if (initialVal > 0 && amountEinmal > initialVal) {
                        zuzahlung = amountEinmal - initialVal;
                        isNeuanlage = false;
                    }
                }

                let sparrateNeu = 0;
                if (hasBaseline) {
                    if (initialSpar === 0 && amountSpar > 0) {
                        sparrateNeu = amountSpar;
                    } else if (initialSpar > 0 && amountSpar > initialSpar) {
                        sparrateNeu = amountSpar - initialSpar;
                    }
                }

                const ertragStr = fund.ertrag ? ` (${fund.ertrag})` : '';
                let amountStr = '';
                if (amountEinmal > 0) amountStr += formatCurrency(amountEinmal);
                if (amountSpar > 0) amountStr += (amountStr ? '\n+ ' : '') + formatCurrency(amountSpar) + ' mtl.';
                if (!amountStr) amountStr = '0,00 €';

                const wknIsinFormatted = formatWknIsin(fund.info);

                if (hasZuzahlung) {
                    let zuzahlungStr = '';
                    if (zuzahlung > 0) zuzahlungStr = `+ ${formatCurrency(zuzahlung)}`;
                    if (sparrateNeu > 0) zuzahlungStr += (zuzahlungStr ? '\n+ ' : '+ ') + formatCurrency(sparrateNeu) + ' mtl.';
                    if (!zuzahlungStr) zuzahlungStr = '-';

                    let bestandStr = '';
                    if (initialVal > 0) bestandStr = formatCurrency(initialVal);
                    else if (amountEinmal > 0 || isNeuanlage) bestandStr = '0,00 €';
                    else bestandStr = '-';
                    if (initialSpar > 0) bestandStr += '\n' + formatCurrency(initialSpar) + ' mtl.';

                    tableData.push([
                        fund.name,
                        wknIsinFormatted,
                        bestandStr,
                        zuzahlungStr,
                        amountStr
                    ]);
                } else {
                    tableData.push([
                        fund.name,
                        wknIsinFormatted,
                        `${fund.type || ''}${ertragStr}`,
                        amountStr
                    ]);
                }
            });

            if (hasEmpfehlung) {
                let eZuz = 0;
                let eIsNeu = false;
                if (hasBaseline) {
                    if (eInit === 0 && eInv > 0) {
                        eZuz = eInv;
                        eIsNeu = true;
                    } else if (eInit > 0 && eInv > eInit) {
                        eZuz = eInv - eInit;
                        eIsNeu = false;
                    }
                }
                let eSparNeu = 0;
                if (hasBaseline) {
                    if (eInitSpar === 0 && eSpar > 0) {
                        eSparNeu = eSpar;
                    } else if (eInitSpar > 0 && eSpar > eInitSpar) {
                        eSparNeu = eSpar - eInitSpar;
                    }
                }

                let amountStr = '';
                if (eInv > 0) amountStr += formatCurrency(eInv);
                if (eSpar > 0) amountStr += (amountStr ? '\n+ ' : '') + formatCurrency(eSpar) + ' mtl.';
                if (!amountStr) amountStr = '0,00 €';

                if (hasZuzahlung) {
                    let zuzahlungStr = '';
                    if (eZuz > 0) zuzahlungStr = `+ ${formatCurrency(eZuz)}`;
                    if (eSparNeu > 0) zuzahlungStr += (zuzahlungStr ? '\n+ ' : '+ ') + formatCurrency(eSparNeu) + ' mtl.';
                    if (!zuzahlungStr) zuzahlungStr = '-';

                    let bestandStr = '';
                    if (eInit > 0) bestandStr = formatCurrency(eInit);
                    else if (eInv > 0 || eIsNeu) bestandStr = '0,00 €';
                    else bestandStr = '-';
                    if (eInitSpar > 0) bestandStr += '\n' + formatCurrency(eInitSpar) + ' mtl.';

                    tableData.push([
                        'Von Empfehlungsliste genommen',
                        displayTitle,
                        bestandStr,
                        zuzahlungStr,
                        amountStr
                    ]);
                } else {
                    tableData.push(['Von Empfehlungsliste genommen', displayTitle, 'Sonderposition', amountStr]);
                }
            }
        });

        if (totalInvestedRows === 0 && (!portfolioGlobals.totalInvestment || portfolioGlobals.totalInvestment === 0)) {
            alert('Keine Fonds im Portfolio ausgewählt.');
            return;
        }

        const headColumns = hasZuzahlung
            ? [['Fondsname', 'WKN / ISIN', 'Bisheriger Bestand', 'Zuzahlung / Neuanlage', 'Gesamtanlage']]
            : [['Fondsname', 'WKN / ISIN', 'Typ / Ertrag', 'Anlagebetrag']];

        const totalGesamtStr = formatCurrency(portfolioGlobals.totalInvestment || totalEinmal) + (totalSpar > 0 ? `\n+ ${formatCurrency(totalSpar)} mtl.` : '');

        let totalZuzahlungFootStr = '';
        if (totalZuzahlung > 0) totalZuzahlungFootStr += `+ ${formatCurrency(totalZuzahlung)}`;
        if (totalSparNeu > 0) totalZuzahlungFootStr += (totalZuzahlungFootStr ? '\n+ ' : '+ ') + `${formatCurrency(totalSparNeu)} mtl.`;
        if (!totalZuzahlungFootStr) totalZuzahlungFootStr = '-';

        const totalBestandFootStr = formatCurrency(totalBestand) + (totalInitialSpar > 0 ? `\n+ ${formatCurrency(totalInitialSpar)} mtl.` : '');

        const footColumns = hasZuzahlung
            ? [['Gesamtsumme', '', totalBestandFootStr, totalZuzahlungFootStr, totalGesamtStr]]
            : [['Gesamtsumme', '', '', totalGesamtStr]];

        const columnStylesConfig = hasZuzahlung
            ? {
                0: { cellWidth: 56 },                              // Fondsname
                1: { cellWidth: 28, fontSize: 7.5 },              // WKN / ISIN (kompakt 2-zeilig)
                2: { cellWidth: 30, halign: 'right' },            // Bisheriger Bestand (rechtsbündig)
                3: { cellWidth: 34, halign: 'right', fontStyle: 'bold' }, // Zuzahlung / Neuanlage (34mm: auch 6-stellige Beträge einzeilig!)
                4: { cellWidth: 34, halign: 'right' }             // Gesamtanlage (rechtsbündig)
            }
            : {
                0: { cellWidth: 66 },
                1: { cellWidth: 32, fontSize: 7.5 },
                2: { cellWidth: 48 },
                3: { cellWidth: 36, halign: 'right' }
            };

        doc.autoTable({
            startY: startYTable,
            margin: { left: 14, right: 14, bottom: 16 },
            head: headColumns,
            body: tableData,
            foot: footColumns,
            showFoot: 'lastPage',
            theme: 'striped',
            styles: {
                fontSize: 8.5,
                cellPadding: { top: 2.8, right: 2.5, bottom: 2.8, left: 2.5 },
                valign: 'middle',
                overflow: 'linebreak',
                textColor: [30, 41, 59]
            },
            headStyles: {
                fillColor: [3, 61, 93],
                textColor: [255, 255, 255],
                fontStyle: 'bold',
                fontSize: 8.5,
                cellPadding: { top: 3.2, right: 2.5, bottom: 3.2, left: 2.5 }
            },
            footStyles: {
                fillColor: [241, 245, 249],
                textColor: [3, 61, 93],
                fontStyle: 'bold',
                fontSize: 8.5,
                cellPadding: { top: 3, right: 2.5, bottom: 3, left: 2.5 }
            },
            alternateRowStyles: {
                fillColor: [248, 250, 252]
            },
            columnStyles: columnStylesConfig,
            didParseCell: function(data) {
                const isZeroRow = data.row.raw && data.row.raw[1] === '-' && (data.row.raw[4] === '0,00 €' || data.row.raw[3] === '0,00 €');
                if (isZeroRow) {
                    data.cell.styles.textColor = [148, 163, 184];
                    data.cell.styles.fontStyle = 'italic';
                }
                if (hasZuzahlung) {
                    // Spalten 2, 3, 4 (Bestand, Zuzahlung/Neuanlage, Gesamtanlage) immer rechtsbündig
                    if ([2, 3, 4].includes(data.column.index)) {
                        data.cell.styles.halign = 'right';
                    }
                    // Zuzahlungen / Neuanlagen in Grün hervorheben (in Body und Foot)
                    if (data.column.index === 3 && !isZeroRow) {
                        const text = String(data.cell.raw || '');
                        if (text.includes('+')) {
                            data.cell.styles.textColor = [21, 128, 61];
                            data.cell.styles.fontStyle = 'bold';
                        }
                    }
                } else {
                    if (data.column.index === 3) {
                        data.cell.styles.halign = 'right';
                    }
                }
            }
        });

        // ── SEITENNUMMERIERUNG & FUSSZEILE ──
        const totalPages = doc.internal.getNumberOfPages();
        for (let i = 1; i <= totalPages; i++) {
            doc.setPage(i);
            doc.setFontSize(8);
            doc.setFont(undefined, 'normal');
            doc.setTextColor(148, 163, 184);
            doc.text('MLP Zeitphasenmodell V6.0 · Unverbindlicher Anlagevorschlag', 14, 290);
            doc.text(`Seite ${i} von ${totalPages}`, 196, 290, { align: 'right' });
        }

        doc.save(fileName);
    }

    if (savePortfolioBtn) {
        savePortfolioBtn.addEventListener('click', exportProposalPdf);
    }
    const savePortfolioPdfBtn = document.getElementById('save-portfolio-pdf-btn');
    if (savePortfolioPdfBtn) {
        savePortfolioPdfBtn.addEventListener('click', exportProposalPdf);
    }

    const printOverallBtn = document.getElementById('print-overall-btn');
    if (printOverallBtn) {
        printOverallBtn.addEventListener('click', () => {
            window.print();
        });
    }

    const setBaselineBtn = document.getElementById('set-baseline-btn');
    if (setBaselineBtn) {
        setBaselineBtn.addEventListener('click', () => {
            portfolioGlobals.initialInvestments = { ...portfolioGlobals.fundInvestments };
            portfolioGlobals.initialSparrates = { ...portfolioGlobals.fundSparrates };
            portfolioGlobals.initialTotalInvestment = portfolioGlobals.totalInvestment || 0;
            saveGlobals();
            updatePortfolioUI();
            if (currentBlockData && modal && modal.classList.contains('open')) {
                renderFundList(currentBlockData.funds);
            }
            alert('Aktueller Stand als Depot-Bestand gespeichert. Zukünftige Erhöhungen werden als Zuzahlung bzw. Neuanlage hervorgehoben.');
        });
    }

    if (toggleVisibilityBtn) {
        toggleVisibilityBtn.addEventListener('click', () => {
            setManagersVisibility(!managersHidden);
        });
    }

    const toggleAllCollapseBtn = document.getElementById('toggle-all-collapse-btn');
    if (toggleAllCollapseBtn) {
        toggleAllCollapseBtn.addEventListener('click', () => {
            if (typeof managementBlocks === 'undefined') return;
            const allBlockIds = managementBlocks.map(b => b.id);
            const allCollapsed = allBlockIds.length > 0 && allBlockIds.every(id => collapsedLayers.has(id));
            if (allCollapsed) {
                collapsedLayers.clear();
            } else {
                allBlockIds.forEach(id => collapsedLayers.add(id));
            }
            try {
                localStorage.setItem('portfolio_collapsed_layers_v6', JSON.stringify([...collapsedLayers]));
            } catch (e) {}
            renderRightPanelLayers();
        });
    }

    const toggleAllClustersEinmalBtn = document.getElementById('toggle-all-clusters-einmal-btn');
    if (toggleAllClustersEinmalBtn) {
        toggleAllClustersEinmalBtn.addEventListener('click', () => {
            const allCollapsed = activeClustersEinmal.length > 0 && activeClustersEinmal.every(k => collapsedClustersEinmal.has(k));
            if (allCollapsed) {
                collapsedClustersEinmal.clear();
            } else {
                activeClustersEinmal.forEach(k => collapsedClustersEinmal.add(k));
            }
            try {
                localStorage.setItem('portfolio_collapsed_clusters_einmal_v6', JSON.stringify([...collapsedClustersEinmal]));
            } catch (e) {}
            updateCountryBreakdown();
        });
    }

    const toggleAllClustersSparrateBtn = document.getElementById('toggle-all-clusters-sparrate-btn');
    if (toggleAllClustersSparrateBtn) {
        toggleAllClustersSparrateBtn.addEventListener('click', () => {
            const allCollapsed = activeClustersSparrate.length > 0 && activeClustersSparrate.every(k => collapsedClustersSparrate.has(k));
            if (allCollapsed) {
                collapsedClustersSparrate.clear();
            } else {
                activeClustersSparrate.forEach(k => collapsedClustersSparrate.add(k));
            }
            try {
                localStorage.setItem('portfolio_collapsed_clusters_sparrate_v6', JSON.stringify([...collapsedClustersSparrate]));
            } catch (e) {}
            updateCountryBreakdown();
        });
    }

    if (resetBtn) {
        resetBtn.addEventListener('click', () => {
            resetPortfolio();
            portfolioGlobals.fundInvestments = {};
            portfolioGlobals.fundSparrates = {};
            portfolioGlobals.initialInvestments = {};
            portfolioGlobals.initialSparrates = {};
            portfolioGlobals.initialTotalInvestment = 0;
            saveGlobals();
            setManagersVisibility(false);
            updatePortfolioUI();
        });
    }

    updatePortfolioUI();
});


// ============================================================
//  CRAWLER / LÄNDER AKTUALISIEREN (PIN 2203 / 220363)
// ============================================================
const CRAWLER_PORT = 8765;

function initCrawlerFeature() {
    const crawlerBtn       = document.getElementById('crawler-btn');
    const crawlerModal     = document.getElementById('crawler-modal');
    const crawlerModalClose= document.getElementById('crawler-modal-close');
    const crawlerNoServer  = document.getElementById('crawler-no-server');
    const crawlerReady     = document.getElementById('crawler-ready');
    const crawlerRunning   = document.getElementById('crawler-running');
    const crawlerDone      = document.getElementById('crawler-done');
    const crawlerStartBtn  = document.getElementById('crawler-start-btn');
    const crawlerLog       = document.getElementById('crawler-log');
    const crawlerBar       = document.getElementById('crawler-progress-bar');
    const crawlerLabel     = document.getElementById('crawler-progress-label');
    const crawlerDoneMsg   = document.getElementById('crawler-done-msg');
    const copyCmdBtn       = document.getElementById('copy-cmd-btn');
    let pollInterval       = null;

    const SERVER = `http://localhost:${CRAWLER_PORT}`;

    const showPane = (pane) => {
        [crawlerNoServer, crawlerReady, crawlerRunning, crawlerDone]
            .forEach(el => el && (el.style.display = 'none'));
        if (pane) pane.style.display = 'block';
    };

    const appendLog = (msg, color = '#cdd6f4') => {
        if (!crawlerLog) return;
        const line = document.createElement('div');
        line.style.color = color;
        line.textContent = msg;
        crawlerLog.appendChild(line);
        crawlerLog.scrollTop = crawlerLog.scrollHeight;
    };

    const stopPolling = () => { if (pollInterval) { clearInterval(pollInterval); pollInterval = null; } };

    const pollStatus = () => {
        fetch(`${SERVER}/status`)
            .then(r => r.json())
            .then(data => {
                const steps = data.progress || [];
                const total = steps.length > 0 ? steps[steps.length-1].total || 1 : 1;
                const done  = steps.filter(s => s.status !== 'running').length;
                if (crawlerBar)   crawlerBar.style.width = `${Math.round((done/total)*100)}%`;
                if (crawlerLabel) crawlerLabel.textContent = `[${done}/${total}] ${steps.length > 0 ? steps[steps.length-1].fund : ''}...`;
                if (steps.length > (crawlerLog ? crawlerLog.children.length : 0)) {
                    const last = steps[steps.length-1];
                    const ok = ['ok','ok_fallback'].includes(last.status);
                    appendLog(`[${done}/${total}] ${last.fund}: ${ok ? '✅' : '❌'} ${last.status}`,
                              ok ? '#a6e3a1' : '#f38ba8');
                }
                if (data.done) {
                    stopPolling();
                    const okCount = steps.filter(s => ['ok','ok_fallback'].includes(s.status)).length;
                    if (crawlerDoneMsg) crawlerDoneMsg.textContent = `Fertig! ✅ ${okCount} Fonds aktualisiert, ❌ ${steps.length - okCount} nicht gefunden.`;
                    showPane(crawlerDone);
                }
                if (data.error) {
                    stopPolling();
                    appendLog(`Fehler: ${data.error}`, '#f38ba8');
                }
            })
            .catch(() => stopPolling());
    };

    if (crawlerBtn) {
        crawlerBtn.addEventListener('click', () => {
            const pwOverlay = document.createElement('div');
            pwOverlay.style.cssText = `
                position:fixed; inset:0; background:rgba(0,0,0,0.65);
                display:flex; align-items:center; justify-content:center; z-index:9999;`;
            pwOverlay.innerHTML = `
                <div style="background:#1e2130; border:1px solid #3d4258; border-radius:14px;
                    padding:36px 40px; min-width:320px; text-align:center; box-shadow:0 8px 40px rgba(0,0,0,0.5);">
                    <div style="font-size:2em; margin-bottom:12px;">🔒</div>
                    <div style="color:#e0e4f0; font-size:1.1em; font-weight:600; margin-bottom:6px;">Zugang geschützt</div>
                    <div style="color:#9aa0bc; font-size:0.88em; margin-bottom:22px;">Bitte PIN eingeben, um die Aktualisierung zu starten.</div>
                    <input id="pw-input" type="password" maxlength="20"
                        placeholder="PIN eingeben"
                        style="width:100%; box-sizing:border-box; padding:10px 14px; font-size:1.1em;
                            border:1.5px solid #3d4258; border-radius:8px; background:#131520;
                            color:#e0e4f0; outline:none; text-align:center; letter-spacing:4px;"
                    />
                    <div id="pw-error" style="color:#f38ba8; font-size:0.85em; margin-top:10px; min-height:18px;"></div>
                    <div style="display:flex; gap:12px; margin-top:20px; justify-content:center;">
                        <button id="pw-cancel" style="padding:9px 24px; border-radius:8px; border:1px solid #3d4258;
                            background:transparent; color:#9aa0bc; cursor:pointer; font-size:0.95em;">Abbrechen</button>
                        <button id="pw-confirm" style="padding:9px 28px; border-radius:8px; border:none;
                            background:#034d6e; color:#fff; cursor:pointer; font-size:0.95em; font-weight:600;">Bestätigen</button>
                    </div>
                </div>`;
            document.body.appendChild(pwOverlay);
            const pwInput   = pwOverlay.querySelector('#pw-input');
            const pwError   = pwOverlay.querySelector('#pw-error');
            const pwConfirm = pwOverlay.querySelector('#pw-confirm');
            const pwCancel  = pwOverlay.querySelector('#pw-cancel');
            setTimeout(() => pwInput.focus(), 50);

            const checkPw = () => {
                const val = pwInput.value.trim();
                if (val === '220363') {
                    document.body.removeChild(pwOverlay);
                    if (crawlerModal) crawlerModal.style.display = 'flex';
                    showPane(null);
                    fetch(`${SERVER}/ping`, { signal: AbortSignal.timeout(2500) })
                        .then(r => r.json())
                        .then(() => showPane(crawlerReady))
                        .catch(() => showPane(crawlerNoServer));
                } else {
                    pwError.textContent = 'Falsche PIN. Bitte erneut versuchen.';
                    pwInput.value = '';
                    pwInput.focus();
                    pwInput.style.borderColor = '#f38ba8';
                    setTimeout(() => { pwInput.style.borderColor = '#3d4258'; pwError.textContent = ''; }, 2000);
                }
            };

            pwConfirm.addEventListener('click', checkPw);
            pwInput.addEventListener('keydown', e => { if (e.key === 'Enter') checkPw(); });
            pwCancel.addEventListener('click', () => document.body.removeChild(pwOverlay));
            pwOverlay.addEventListener('click', e => { if (e.target === pwOverlay) document.body.removeChild(pwOverlay); });
        });
    }

    if (crawlerModalClose) {
        crawlerModalClose.addEventListener('click', () => {
            if (crawlerModal) crawlerModal.style.display = 'none';
            stopPolling();
        });
    }
    window.addEventListener('click', e => {
        if (e.target === crawlerModal) { crawlerModal.style.display = 'none'; stopPolling(); }
    });

    if (copyCmdBtn) {
        copyCmdBtn.addEventListener('click', () => {
            const dir = window.location.href.replace('index.html','').replace('file://','');
            const cmd = `cd "${decodeURIComponent(dir)}" && python3 crawler_server.py`;
            navigator.clipboard.writeText(cmd).then(() => { copyCmdBtn.textContent = '✅ Kopiert!'; setTimeout(() => { copyCmdBtn.textContent = '📋 Kopieren'; }, 2000); });
        });
    }

    if (crawlerStartBtn) {
        crawlerStartBtn.addEventListener('click', () => {
            showPane(crawlerRunning);
            if (crawlerLog) crawlerLog.innerHTML = '';
            if (crawlerBar) crawlerBar.style.width = '0%';
            if (crawlerLabel) crawlerLabel.textContent = 'Starte Crawler...';
            appendLog('Verbinde mit Crawle-Server...');
            fetch(`${SERVER}/run`, { method: 'POST' })
                .then(r => r.json())
                .then(d => {
                    if (d.ok) {
                        appendLog('Crawler gestartet 🚀', '#89b4fa');
                        pollInterval = setInterval(pollStatus, 2000);
                    } else {
                        appendLog(`Fehler: ${d.error}`, '#f38ba8');
                    }
                })
                .catch(e => appendLog(`Verbindungsfehler: ${e}`, '#f38ba8'));
        });
    }
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initCrawlerFeature);
} else {
    initCrawlerFeature();
}


// ══════════════════════════════════════════════════════════════════════════════
// ══ RESTORED V6.0 TOOLBAR & MODAL ACTION HANDLERS (LIBRARY, UPLOAD, EXPORT) ══
// ══════════════════════════════════════════════════════════════════════════════

function getSavedLibrary() {
    const keys = [
        'vem_library_v6.0',
        'vem_library_v5.2.1',
        'vem_library_v5.2',
        'vem_library_v44',
        'depotBibliothek_V50',
        'depotBibliothek'
    ];
    for (const k of keys) {
        try {
            const raw = localStorage.getItem(k);
            if (raw) {
                const parsed = JSON.parse(raw);
                if (Array.isArray(parsed) && parsed.length > 0) {
                    return parsed;
                }
            }
        } catch(e) {}
    }
    try {
        const raw = localStorage.getItem('vem_library_v6.0');
        if (raw) {
            const parsed = JSON.parse(raw);
            if (Array.isArray(parsed)) return parsed;
        }
    } catch(e) {}
    return [];
}

function saveSavedLibrary(items) {
    try {
        const valid = Array.isArray(items) ? items : [];
        const json = JSON.stringify(valid);
        localStorage.setItem('vem_library_v6.0', json);
        localStorage.setItem('vem_library_v5.2.1', json);
        localStorage.setItem('depotBibliothek_V50', json);
    } catch(e) {}
}

function openLibraryModal() {
    const libModal = document.getElementById('library-modal');
    if (!libModal) return;
    renderLibraryModal();
    libModal.style.display = 'flex';
    libModal.classList.add('open');
}
window.openLibraryModal = openLibraryModal;

function closeLibraryModal() {
    const libModal = document.getElementById('library-modal');
    if (!libModal) return;
    libModal.classList.remove('open');
    libModal.style.display = 'none';
}
window.closeLibraryModal = closeLibraryModal;

function renderLibraryModal() {
    const libraryList = document.getElementById('library-list');
    const libraryEmpty = document.getElementById('library-empty');
    if (!libraryList || !libraryEmpty) return;

    const items = getSavedLibrary();
    libraryList.innerHTML = '';
    if (items.length === 0) {
        libraryEmpty.style.display = 'block';
        return;
    }
    libraryEmpty.style.display = 'none';

    items.forEach((item, index) => {
        const div = document.createElement('div');
        div.className = 'library-item';
        div.style.cssText = 'background:#f8fafc; border:1px solid #e2e8f0; border-radius:8px; padding:12px; margin-bottom:10px; display:flex; justify-content:space-between; align-items:center;';
        
        const d = item.date ? new Date(item.date) : null;
        const dateStr = (d && !isNaN(d.getTime())) ? d.toLocaleDateString('de-DE') : '';
        const inv = item.investments || item.fundInvestments || {};
        let totalAmount = (inv && typeof inv === 'object') ? Object.values(inv).reduce((a, b) => a + (parseFloat(b) || 0), 0) : 0;
        if (!totalAmount && item.totalInvestment) totalAmount = parseFloat(item.totalInvestment) || 0;

        div.innerHTML = `
            <div>
                <div style="font-weight:700; color:#0f172a; font-size:1.02rem;">${item.name || 'Unbenanntes Depot'}</div>
                <div style="font-size:0.82rem; color:#64748b; margin-top:2px;">Erstellt: ${dateStr} | Gesamt: ${formatCurrency(totalAmount)}</div>
            </div>
            <div style="display:flex; gap:8px;">
                <button class="btn-lib-load" data-index="${index}" style="background:#0284c7; color:#fff; border:none; padding:6px 12px; border-radius:6px; font-weight:600; cursor:pointer;">▶ Laden</button>
                <button class="btn-lib-export" data-index="${index}" style="background:#64748b; color:#fff; border:none; padding:6px 10px; border-radius:6px; font-weight:600; cursor:pointer;">⬇ JSON</button>
                <button class="btn-lib-del" data-index="${index}" style="background:#ef4444; color:#fff; border:none; padding:6px 10px; border-radius:6px; font-weight:600; cursor:pointer;">🗑</button>
            </div>
        `;
        libraryList.appendChild(div);
    });

    libraryList.querySelectorAll('.btn-lib-load').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const idx = parseInt(e.currentTarget.dataset.index, 10);
            const items = getSavedLibrary();
            if (items[idx]) {
                loadPortfolioSetup(items[idx]);
                closeLibraryModal();
            }
        });
    });

    libraryList.querySelectorAll('.btn-lib-export').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const idx = parseInt(e.currentTarget.dataset.index, 10);
            const items = getSavedLibrary();
            if (items[idx]) {
                const item = items[idx];
                const exportData = {
                    version: "V6.0",
                    name: item.name || 'depot',
                    date: item.date || new Date().toISOString(),
                    totalInvestment: item.totalInvestment || 0,
                    totalSparrate: item.totalSparrate || 0,
                    portfolio: item.portfolio || [],
                    investments: item.investments || item.fundInvestments || {},
                    fundInvestments: item.fundInvestments || item.investments || {},
                    fundSparrates: item.fundSparrates || item.sparrates || {}
                };
                const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(exportData, null, 2));
                const dlAnchor = document.createElement('a');
                dlAnchor.setAttribute("href", dataStr);
                dlAnchor.setAttribute("download", `${(item.name || 'depot').replace(/\s+/g, '_')}_v6.0.json`);
                document.body.appendChild(dlAnchor);
                dlAnchor.click();
                dlAnchor.remove();
            }
        });
    });

    libraryList.querySelectorAll('.btn-lib-del').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const idx = parseInt(e.currentTarget.dataset.index, 10);
            const items = getSavedLibrary();
            const itemName = items[idx]?.name || 'dieses Depot';
            if (!confirm(`Möchten Sie "${itemName}" wirklich aus der Bibliothek löschen?`)) return;
            items.splice(idx, 1);
            saveSavedLibrary(items);
            renderLibraryModal();
        });
    });
}

function parseGermanAmount(val) {
    if (val === null || val === undefined) return 0;
    if (typeof val === 'number') return val > 0 ? val : 0;
    let str = String(val).replace(/\u00a0/g, ' ').replace(/[Ââ€$]/g, '').replace(/\s+/g, '').trim();
    const numMatch = str.match(/(\d+[\d.,]*)/);
    if (!numMatch) return 0;
    let numStr = numMatch[1];
    if (numStr.includes(',')) numStr = numStr.replace(/\./g, '').replace(',', '.');
    const parsed = parseFloat(numStr);
    return (!isNaN(parsed) && parsed > 0) ? parsed : 0;
}

function matchFundByWknOrName(searchWkn, searchIsin, searchName) {
    let matchedBlock = null;
    let matchedFund = null;
    const sWkn = searchWkn ? searchWkn.toString().trim().toUpperCase().padStart(6, '0') : '';
    const sIsin = (searchIsin || '').toString().trim().toUpperCase();
    const sName = (searchName || '').toString().trim().toUpperCase();

    managementBlocks.forEach(block => {
        block.funds.forEach(fund => {
            if (matchedFund) return;
            const infoUpper = (fund.info || '').toUpperCase();
            const nameUpper = (fund.name || '').toUpperCase();
            const fWknMatch = infoUpper.match(/WKN:\s*([A-Z0-9]{6})/);
            const fIsinMatch = infoUpper.match(/ISIN:\s*([A-Z0-9]{12})/);
            const fWkn = fWknMatch ? fWknMatch[1] : '';
            const fIsin = fIsinMatch ? fIsinMatch[1] : '';

            if (sWkn && sWkn.length === 6 && (fWkn === sWkn || infoUpper.includes(sWkn))) {
                matchedBlock = block; matchedFund = fund;
            } else if (sIsin && sIsin.length === 12 && (fIsin === sIsin || infoUpper.includes(sIsin))) {
                matchedBlock = block; matchedFund = fund;
            } else if (sName && sName.length > 3 && (nameUpper.includes(sName) || sName.includes(nameUpper))) {
                matchedBlock = block; matchedFund = fund;
            }
        });
    });
    return { matchedBlock, matchedFund };
}

let _lastImportedName = "Depot";
function processImportedEntries(entries, sourceFileName) {
    _lastImportedName = sourceFileName ? sourceFileName.replace(/\.[^.]+$/, "") : "Depot";
    let matchedInvestments = {};
    let totalVol = 0;
    let matchedCount = 0;
    let unmatchedCount = 0;

    // Reset portfolio and globals on new depot import
    portfolio = [];
    portfolioGlobals.fundInvestments = {};
    portfolioGlobals.fundSparrates = {};
    try { localStorage.removeItem('empfehlungslisteFonds_V44'); } catch(e) {}

    const empFondsMap = {};

    entries.forEach(entry => {
        const { matchedBlock, matchedFund } = matchFundByWknOrName(entry.wkn, entry.isin, entry.name);
        const amount = entry.amount || 0;
        if (matchedFund && matchedBlock) {
            const keyColon = `${matchedBlock.id}::${matchedFund.name}`;
            matchedInvestments[keyColon] = (matchedInvestments[keyColon] || 0) + amount;
            totalVol += amount;
            matchedCount++;
            if (typeof addFund === 'function') {
                addFund(matchedBlock, matchedFund);
            }
        } else if (amount > 0 || entry.wkn || entry.name) {
            unmatchedCount++;
            const bId = (typeof anlageschwerpunktToBlock === 'function' && entry.schwerpunkt ? anlageschwerpunktToBlock(entry.schwerpunkt) : null) || 'block-defensiv';
            const eKey = `${bId}::empfehlungsliste`;
            matchedInvestments[eKey] = (matchedInvestments[eKey] || 0) + amount;
            totalVol += amount;
            if (!empFondsMap[bId]) empFondsMap[bId] = [];
            empFondsMap[bId].push({
                name: entry.name || `Fonds (${entry.wkn || 'delistet'})`,
                wkn: entry.wkn || '',
                schwerpunkt: entry.schwerpunkt || '',
                einmal: amount,
                sparrate: 0
            });
            const block = managementBlocks.find(b => b.id === bId);
            if (block) {
                const layer = getOrCreateLayer(block.id, block.title);
                if (!layer.funds.some(f => f._isEmpfehlungsliste)) {
                    layer.funds.push({ name: EMPFEHLUNG_KEY_PREFIX, info: '', type: '', ertrag: '', _isEmpfehlungsliste: true });
                }
            }
        }
    });

    if (Object.keys(empFondsMap).length > 0) {
        try { localStorage.setItem('empfehlungslisteFonds_V44', JSON.stringify(empFondsMap)); } catch(e) {}
    }

    portfolioGlobals.fundInvestments = matchedInvestments;
    portfolioGlobals.initialInvestments = { ...matchedInvestments };
    portfolioGlobals.initialSparrates = { ...portfolioGlobals.fundSparrates };
    if (totalVol > 0) {
        portfolioGlobals.totalInvestment = Math.round(totalVol * 100) / 100;
        portfolioGlobals.initialTotalInvestment = portfolioGlobals.totalInvestment;
        const totalInput = document.getElementById('total-investment-input');
        if (totalInput) totalInput.value = formatNumberInput(portfolioGlobals.totalInvestment);
    }
    saveGlobals();
    savePortfolio();

    if (typeof window.updatePortfolioUI === 'function') {
        window.updatePortfolioUI();
    } else if (typeof updatePortfolioUI === 'function') {
        updatePortfolioUI();
    }

    if (typeof window.renderGridColumns === 'function') {
        window.renderGridColumns();
    } else if (typeof renderGridColumns === 'function') {
        renderGridColumns();
    }

    const pdfImportModal = document.getElementById('pdf-import-modal');
    const summaryElem = document.getElementById('pdf-import-summary');
    const matchedCountElem = document.getElementById('pdf-matched-count');
    const unmatchedCountElem = document.getElementById('pdf-unmatched-count');

    if (summaryElem) summaryElem.textContent = `Datei "${sourceFileName}" verarbeitet: ${matchedCount} Fonds zugeordnet, Gesamtwert: ${formatCurrency(totalVol)}`;
    const saveNameInp = document.getElementById('import-save-name');
    if (saveNameInp) saveNameInp.value = _lastImportedName;
    if (matchedCountElem) matchedCountElem.textContent = matchedCount;
    if (unmatchedCountElem) unmatchedCountElem.textContent = unmatchedCount;
    if (pdfImportModal) pdfImportModal.classList.add('open');
}

function getRowMonetaryAmount(cells) {
    if (!cells || !Array.isArray(cells)) return 0;
    for (const c of cells) {
        if (c === null || c === undefined) continue;
        const raw = String(c);
        // Geschütztes Leerzeichen \xa0, Euro-Zeichen, Tausenderpunkte, Komma als Dezimaltrenner
        const str = raw.replace(/\u00a0/g, '').replace(/€/g, '').replace(/\s/g, '').replace(/\./g, '').replace(',', '.').trim();
        if (!str) continue;
        const rawStr = raw.replace(/\u00a0/g, '').replace(/€/g, '').trim();
        // WKN-Filter: Nur überspringen wenn der Wert WIRKLICH wie eine WKN aussieht
        // (keine Kommas, keine Leerzeichen, kein €, genau 6 alphanumerische Zeichen)
        // Beträge wie "9.773,02 €" dürfen NICHT gefiltert werden (enthalten Komma/Punkt/€)
        const hasComma = raw.includes(',');
        const hasCurrency = raw.includes('€') || raw.includes('\u00a0');
        if (!hasComma && !hasCurrency && /^[A-Z0-9]{6}$/i.test(rawStr.replace(/[^A-Z0-9]/gi, ''))) continue;
        const v = parseFloat(str);
        if (!isNaN(v) && v > 0 && v < 2000000) return v;
    }
    return 0;
}


function extractWknFromRow(cells) {
    if (!cells || !Array.isArray(cells)) return '';
    for (const c of cells) {
        if (c === null || c === undefined || c === '') continue;
        // Nachkomma entfernen (Excel speichert Integer manchmal als Float: 986838.0)
        const s = String(c).trim().toUpperCase().replace(/\.0+$/, '');
        // Alphanumerische WKNs direkt prüfen (6 Stellen, keine Führungsnullen ergänzen)
        if (/^[A-Z0-9]{6}$/.test(s) && s !== 'ANLEIH' && s !== 'AKTIEN') {
            return s;
        }
        // Rein numerische WKNs < 6 Stellen mit Führungsnullen auffüllen (z.B. 214466 → "214466")
        if (/^\d{1,6}$/.test(s)) {
            const padded = s.padStart(6, '0');
            if (!padded.startsWith('00')) return padded;
        }
    }
    return '';
}

function parseCsvDepot(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
        const text = e.target.result;
        const rawLines = text.split(/\r?\n/);
        let rowsCells = [];
        rawLines.forEach(line => {
            if (!line.trim()) return;
            const cells = line.split(/[;,\t]/).map(c => c.replace(/"/g, '').trim()).filter(Boolean);
            if (cells.length > 0) rowsCells.push(cells);
        });

        let entries = [];
        let seenWkns = new Set();
        for (let i = 0; i < rowsCells.length; i++) {
            const wkn = extractWknFromRow(rowsCells[i]);
            if (wkn && !seenWkns.has(wkn)) {
                let amt = getRowMonetaryAmount(rowsCells[i]);
                if (amt === 0 && i > 0) amt = getRowMonetaryAmount(rowsCells[i - 1]);
                if (amt === 0 && i + 1 < rowsCells.length) amt = getRowMonetaryAmount(rowsCells[i + 1]);
                if (amt > 0) {
                    seenWkns.add(wkn);
                    let name = '';
                    let schwerpunkt = '';
                    rowsCells[i].forEach(cell => {
                        const cStr = String(cell).trim();
                        if (cStr && !cStr.includes('€') && cStr !== wkn && !/^\d+$/.test(cStr)) {
                            if (typeof anlageschwerpunktToBlock === 'function' && anlageschwerpunktToBlock(cStr)) schwerpunkt = cStr;
                            else if (!name) name = cStr;
                        }
                    });
                    if (i + 1 < rowsCells.length) {
                        const nextRowFirst = String(rowsCells[i + 1][0] || '').trim();
                        if (nextRowFirst && !extractWknFromRow(rowsCells[i + 1])) {
                            if (!name) name = nextRowFirst;
                        }
                    }
                    entries.push({ wkn: wkn, isin: '', name: name, amount: amt, schwerpunkt: schwerpunkt });
                }
            }
        }
        processImportedEntries(entries, file.name);
    };
    reader.readAsText(file, 'ISO-8859-1');
}

function parseXlsxDepot(file) {
    if (typeof XLSX === 'undefined') {
        alert("XLSX-Bibliothek ist nicht geladen.");
        return;
    }
    const reader = new FileReader();
    reader.onerror = () => {
        alert("Fehler beim Lesen der Datei. Bitte prüfe, ob die Datei geöffnet oder gesperrt ist.");
    };
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);

            // Robuster Lesemodus: zuerst standard, dann Fallback ohne Passwortschutz-Features
            let workbook;
            try {
                workbook = XLSX.read(data, { type: 'array', cellDates: false, WTF: false });
            } catch(readErr) {
                try {
                    workbook = XLSX.read(data, { type: 'array', raw: true });
                } catch(readErr2) {
                    throw new Error('Excel-Datei konnte nicht gelesen werden: ' + readErr2.message);
                }
            }

            if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) {
                throw new Error('Die Excel-Datei enthält keine Tabellenblätter.');
            }

            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            if (!firstSheet) {
                throw new Error('Das erste Tabellenblatt ist leer oder nicht lesbar.');
            }

            const jsonRows = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });

            let rowsCells = [];
            jsonRows.forEach(row => {
                if (!Array.isArray(row)) return;
                const cells = row.map(c => (c === null || c === undefined) ? '' : String(c).trim()).filter(Boolean);
                if (cells.length > 0) rowsCells.push(cells);
            });

            if (rowsCells.length === 0) {
                alert('Die Excel-Datei scheint keine lesbaren Daten zu enthalten.');
                return;
            }

            let entries = [];
            let seenWkns = new Set();
            for (let i = 0; i < rowsCells.length; i++) {
                const wkn = extractWknFromRow(rowsCells[i]);
                if (wkn && !seenWkns.has(wkn)) {
                    let amt = getRowMonetaryAmount(rowsCells[i]);
                    if (amt === 0 && i > 0) amt = getRowMonetaryAmount(rowsCells[i - 1]);
                    if (amt === 0 && i + 1 < rowsCells.length) amt = getRowMonetaryAmount(rowsCells[i + 1]);
                    if (amt > 0) {
                        seenWkns.add(wkn);
                        let name = '';
                        let schwerpunkt = '';
                        rowsCells[i].forEach(cell => {
                            const cStr = String(cell).trim();
                            if (cStr && !cStr.includes('€') && cStr !== wkn && !/^\d+$/.test(cStr)) {
                                if (typeof anlageschwerpunktToBlock === 'function' && anlageschwerpunktToBlock(cStr)) schwerpunkt = cStr;
                                else if (!name) name = cStr;
                            }
                        });
                        if (i + 1 < rowsCells.length) {
                            const nextRowFirst = String(rowsCells[i + 1][0] || '').trim();
                            if (nextRowFirst && !extractWknFromRow(rowsCells[i + 1])) {
                                if (!name) name = nextRowFirst;
                            }
                        }
                        entries.push({ wkn: wkn, isin: '', name: name, amount: amt, schwerpunkt: schwerpunkt });
                    }
                }
            }

            if (entries.length === 0) {
                alert('Keine Fonds mit WKN und Betrag in der Datei gefunden.\n\nHinweis: Die Datei muss WKNs (6-stellig) und Beträge enthalten.');
                return;
            }

            processImportedEntries(entries, file.name);
        } catch(err) {
            console.error("Excel import error:", err);
            alert('Fehler beim Lesen der Excel-Datei:\n' + (err.message || err));
        }
    };
    reader.readAsArrayBuffer(file);
}


function parsePdfDepot(file) {
    if (typeof pdfjsLib === 'undefined') {
        alert("PDF.js-Bibliothek ist nicht geladen.");
        return;
    }
    const reader = new FileReader();
    reader.onload = function() {
        const typedarray = new Uint8Array(this.result);
        pdfjsLib.getDocument(typedarray).promise.then(pdf => {
            let maxPages = pdf.numPages;
            let countPromises = [];
            for (let i = 1; i <= maxPages; i++) {
                countPromises.push(pdf.getPage(i).then(page => page.getTextContent()));
            }
            Promise.all(countPromises).then(contents => {
                let fullText = '';
                contents.forEach(content => {
                    content.items.forEach(item => { fullText += item.str + ' '; });
                });
                let entries = [];
                const wknMatches = fullText.match(/\b([A-Z0-9]{6})\b/g) || [];
                wknMatches.forEach(wkn => {
                    entries.push({ wkn: wkn, isin: '', name: '', amount: 0 });
                });
                processImportedEntries(entries, file.name);
            });
        });
    };
    reader.readAsArrayBuffer(file);
}

function loadPortfolioSetup(imported) {
    if (!imported) return false;
    const src = imported.data ? imported.data : imported;
    const g = src.globals || imported.globals || {};
    const inv = src.fundInvestments || src.investments || g.fundInvestments || g.investments || {};
    const spar = src.fundSparrates || src.sparrates || g.fundSparrates || g.sparrates || {};
    const totalInv = Math.round((src.totalInvestment || g.totalInvestment || 0) * 100) / 100;
    const totalSpar = Math.round((src.totalSparrate || g.totalSparrate || 0) * 100) / 100;
    const initialInv = src.initialInvestments || imported.initialInvestments || { ...inv };
    const initialSpar = src.initialSparrates || imported.initialSparrates || { ...spar };
    const initialTot = Math.round((src.initialTotalInvestment || imported.initialTotalInvestment || totalInv) * 100) / 100;

    // Restore empfehlungslisteFonds if present
    if (src.empfehlungslisteFonds || imported.empfehlungslisteFonds) {
        try {
            localStorage.setItem('empfehlungslisteFonds_V44', JSON.stringify(src.empfehlungslisteFonds || imported.empfehlungslisteFonds));
        } catch(e) {}
    }

    portfolioGlobals = sanitizeGlobals({
        totalInvestment: totalInv,
        totalSparrate: totalSpar,
        fundInvestments: inv,
        fundSparrates: spar,
        initialInvestments: initialInv,
        initialSparrates: initialSpar,
        initialTotalInvestment: initialTot
    });
    portfolio = [];

    // Rebuild layers if explicit portfolio array is saved
    if (Array.isArray(imported.portfolio) && imported.portfolio.length > 0) {
        imported.portfolio.forEach(layer => {
            const block = managementBlocks.find(b => b.id === layer.blockId);
            if (block && Array.isArray(layer.funds)) {
                layer.funds.forEach(f => {
                    const fund = block.funds.find(bf => bf.name === f.name) || f;
                    addFund(block, fund);
                });
            }
        });
    }

    // Also ensure all funds with an investment or sparrate key are present
    const addFundFromKey = (k) => {
        const parts = k.split('::');
        if (parts.length >= 2) {
            const blockId = parts[0];
            const fundName = parts.slice(1).join('::');
            const block = managementBlocks.find(b => b.id === blockId);
            if (block) {
                let fund = block.funds.find(f => f.name === fundName);
                if (!fund) {
                    fund = { name: fundName, info: '', type: 'Fonds', ertrag: '' };
                }
                addFund(block, fund);
            }
        }
    };

    Object.keys(portfolioGlobals.fundInvestments).forEach(addFundFromKey);
    Object.keys(portfolioGlobals.fundSparrates).forEach(addFundFromKey);

    if (portfolioGlobals.totalInvestment > 0) {
        const totalInput = document.getElementById('total-investment-input');
        if (totalInput) totalInput.value = formatNumberInput(portfolioGlobals.totalInvestment);
    }
    if (portfolioGlobals.totalSparrate > 0) {
        const totalSpar = document.getElementById('total-investment-sparrate');
        if (totalSpar) totalSpar.value = formatNumberInput(portfolioGlobals.totalSparrate);
    }

    saveGlobals();
    savePortfolio();

    if (typeof window.updatePortfolioUI === 'function') {
        window.updatePortfolioUI();
    } else if (typeof updatePortfolioUI === 'function') {
        updatePortfolioUI();
    }
    if (typeof window.renderGridColumns === 'function') {
        window.renderGridColumns();
    } else if (typeof renderGridColumns === 'function') {
        renderGridColumns();
    }
    if (typeof window.renderStandaloneCylinders === 'function') {
        window.renderStandaloneCylinders();
    } else if (typeof renderStandaloneCylinders === 'function') {
        renderStandaloneCylinders();
    }
    return true;
}
window.loadPortfolioSetup = loadPortfolioSetup;

function parseJsonSetup(file) {
    const reader = new FileReader();
    reader.onload = (evt) => {
        try {
            const imported = JSON.parse(evt.target.result);
            if (imported && (imported.investments || imported.fundInvestments || imported.portfolio || imported.globals)) {
                loadPortfolioSetup(imported);
                const name = imported.name || (file && file.name ? file.name.replace(/\.[^.]+$/, '') : `Depot ${new Date().toLocaleDateString('de-DE')}`);
                const lib = getSavedLibrary();
                const existingIdx = lib.findIndex(e => e.name === name);
                const entry = {
                    id: Date.now(),
                    name: name,
                    date: imported.date || new Date().toISOString(),
                    totalInvestment: portfolioGlobals.totalInvestment,
                    totalSparrate: portfolioGlobals.totalSparrate,
                    initialTotalInvestment: portfolioGlobals.initialTotalInvestment || portfolioGlobals.totalInvestment,
                    portfolio: JSON.parse(JSON.stringify(portfolio)),
                    investments: { ...portfolioGlobals.fundInvestments },
                    fundInvestments: { ...portfolioGlobals.fundInvestments },
                    fundSparrates: { ...portfolioGlobals.fundSparrates },
                    initialInvestments: { ...(portfolioGlobals.initialInvestments || portfolioGlobals.fundInvestments) },
                    initialSparrates: { ...(portfolioGlobals.initialSparrates || portfolioGlobals.fundSparrates) },
                    empfehlungslisteFonds: JSON.parse(localStorage.getItem('empfehlungslisteFonds_V44') || '{}')
                };
                if (existingIdx > -1) lib[existingIdx] = entry;
                else lib.unshift(entry);
                saveSavedLibrary(lib);
                renderLibraryModal();
                closeLibraryModal();
            } else {
                alert("Ungültiges Dateiformat für das Portfolio-Setup.");
            }
        } catch(err) {
            console.error("JSON parse error:", err);
            alert("Fehler beim Lesen der JSON-Setup-Datei.");
        }
    };
    reader.readAsText(file);
}

// Global Toolbar Initialization Function
document.addEventListener('DOMContentLoaded', () => {

    // ── V6.0 IMPORT-ERGEBNIS-MODAL AKTIONEN (BIBLIOTHEK, LADEN, EXPORT) ──
    const importSaveBtn = document.getElementById('import-save-btn');
    const importLoadBtn = document.getElementById('import-load-btn');
    const importJsonDlBtn = document.getElementById('import-json-dl-btn');
    const importSaveName = document.getElementById('import-save-name');
    const pdfImportModal = document.getElementById('pdf-import-modal');
    const pdfImportModalClose = document.getElementById('pdf-import-modal-close');
    const pdfImportModalOk = document.getElementById('pdf-import-modal-ok');

    const closeImportModal = () => {
        if (pdfImportModal) {
            pdfImportModal.classList.remove('open');
            setTimeout(() => { pdfImportModal.style.display = 'none'; }, 300);
        }
    };
    if (pdfImportModalClose) pdfImportModalClose.addEventListener('click', closeImportModal);
    if (pdfImportModalOk) pdfImportModalOk.addEventListener('click', closeImportModal);

    if (importSaveBtn) {
        importSaveBtn.addEventListener('click', () => {
            const name = (importSaveName?.value || (typeof _lastImportedName !== 'undefined' ? _lastImportedName : 'Depot')).trim() || 'Depot';
            const lib = getSavedLibrary();
            const existingIdx = lib.findIndex(e => e.name === name);
            const entry = {
                id: Date.now(),
                name: name,
                date: new Date().toISOString(),
                totalInvestment: portfolioGlobals.totalInvestment,
                totalSparrate: portfolioGlobals.totalSparrate,
                initialTotalInvestment: portfolioGlobals.initialTotalInvestment || portfolioGlobals.totalInvestment,
                portfolio: JSON.parse(JSON.stringify(portfolio)),
                investments: { ...portfolioGlobals.fundInvestments },
                fundInvestments: { ...portfolioGlobals.fundInvestments },
                fundSparrates: { ...portfolioGlobals.fundSparrates },
                initialInvestments: { ...(portfolioGlobals.initialInvestments || portfolioGlobals.fundInvestments) },
                initialSparrates: { ...(portfolioGlobals.initialSparrates || portfolioGlobals.fundSparrates) },
                empfehlungslisteFonds: JSON.parse(localStorage.getItem('empfehlungslisteFonds_V44') || '{}')
            };
            if (existingIdx > -1) lib[existingIdx] = entry;
            else lib.unshift(entry);
            saveSavedLibrary(lib);
            importSaveBtn.textContent = '✅ Gespeichert!';
            setTimeout(() => { importSaveBtn.textContent = '💾 In Bibliothek speichern'; }, 2000);
        });
    }

    if (importLoadBtn) {
        importLoadBtn.addEventListener('click', () => {
            closeImportModal();
        });
    }

    if (importJsonDlBtn) {
        importJsonDlBtn.addEventListener('click', () => {
            const name = (importSaveName?.value || (typeof _lastImportedName !== 'undefined' ? _lastImportedName : 'Depot')).trim() || 'Depot';
            const setup = {
                version: "V6.0",
                name: name,
                date: new Date().toISOString(),
                totalInvestment: portfolioGlobals.totalInvestment,
                totalSparrate: portfolioGlobals.totalSparrate,
                initialTotalInvestment: portfolioGlobals.initialTotalInvestment || portfolioGlobals.totalInvestment,
                portfolio: portfolio,
                fundInvestments: { ...portfolioGlobals.fundInvestments },
                fundSparrates: { ...portfolioGlobals.fundSparrates },
                initialInvestments: { ...(portfolioGlobals.initialInvestments || portfolioGlobals.fundInvestments) },
                initialSparrates: { ...(portfolioGlobals.initialSparrates || portfolioGlobals.fundSparrates) },
                empfehlungslisteFonds: JSON.parse(localStorage.getItem('empfehlungslisteFonds_V44') || '{}')
            };
            const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(setup, null, 2));
            const dl = document.createElement('a');
            dl.setAttribute("href", dataStr);
            dl.setAttribute("download", `${name.replace(/\s+/g, '_')}_v6.0.json`);
            document.body.appendChild(dl);
            dl.click();
            dl.remove();
        });
    }

    const helpModal = document.getElementById('help-modal');
    const helpBtn = document.getElementById('help-btn');
    const helpCloseBtn = document.getElementById('help-modal-close');
    if (helpBtn && helpModal) helpBtn.addEventListener('click', () => helpModal.classList.add('open'));
    if (helpCloseBtn && helpModal) helpCloseBtn.addEventListener('click', () => helpModal.classList.remove('open'));

    const libModal = document.getElementById('library-modal');
    const libOpenBtn = document.getElementById('library-open-btn');
    const libCloseBtn = document.getElementById('library-modal-close');
    if (libOpenBtn) {
        libOpenBtn.addEventListener('click', openLibraryModal);
    }
    if (libCloseBtn) {
        libCloseBtn.addEventListener('click', closeLibraryModal);
    }
    if (libModal) {
        libModal.addEventListener('click', (e) => {
            if (e.target === libModal) closeLibraryModal();
        });
    }
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape' && libModal && (libModal.classList.contains('open') || libModal.style.display === 'flex')) {
            closeLibraryModal();
        }
    });

    const depotUploadInput = document.getElementById('depot-upload-input');
    const depotUploadLabel = document.getElementById('depot-upload-label');
    const importSetupFile = document.getElementById('import-setup-file');
    const libImportFile = document.getElementById('library-import-file');

    function handleFileSelection(file) {
        if (!file) return;
        const fn = file.name.toLowerCase();
        if (fn.endsWith('.csv')) parseCsvDepot(file);
        else if (fn.endsWith('.xlsx') || fn.endsWith('.xls')) parseXlsxDepot(file);
        else if (fn.endsWith('.pdf')) parsePdfDepot(file);
        else if (fn.endsWith('.json')) parseJsonSetup(file);
        else alert('Bitte eine .csv, .xlsx, .xls, .pdf oder .json Datei auswählen.');
    }
    window.handleFileSelection = handleFileSelection; // sofort global registrieren

    if (depotUploadInput) {
        depotUploadInput.addEventListener('change', (e) => {
            handleFileSelection(e.target.files[0]);
            e.target.value = '';
        });
    }
    if (depotUploadLabel && depotUploadInput) {
        depotUploadLabel.addEventListener('click', (e) => {
            if (e.target !== depotUploadInput) depotUploadInput.click();
        });
    }
    if (importSetupFile) {
        importSetupFile.addEventListener('change', (e) => {
            handleFileSelection(e.target.files[0]);
            e.target.value = '';
        });
    }
    if (libImportFile) {
        libImportFile.addEventListener('change', (e) => {
            handleFileSelection(e.target.files[0]);
            e.target.value = '';
        });
    }

    const saveBtn = document.getElementById('save-portfolio-btn');
    if (saveBtn) {
        saveBtn.addEventListener('click', () => {
            const name = prompt("Name für diese Depot-Konfiguration:", `Depot ${new Date().toLocaleDateString('de-DE')}`);
            if (!name) return;
            const lib = getSavedLibrary();
            lib.push({
                id: Date.now(),
                name: name,
                date: new Date().toISOString(),
                totalInvestment: portfolioGlobals.totalInvestment,
                totalSparrate: portfolioGlobals.totalSparrate,
                portfolio: portfolio,
                investments: { ...portfolioGlobals.fundInvestments },
                fundInvestments: { ...portfolioGlobals.fundInvestments },
                fundSparrates: { ...portfolioGlobals.fundSparrates }
            });
            saveSavedLibrary(lib);
            alert(`Depot "${name}" wurde in der Bibliothek gespeichert.`);
        });
    }

    const exportBtn = document.getElementById('export-setup-btn');
    if (exportBtn) {
        exportBtn.addEventListener('click', () => {
            const setup = {
                version: "V6.0",
                date: new Date().toISOString(),
                totalInvestment: portfolioGlobals.totalInvestment,
                totalSparrate: portfolioGlobals.totalSparrate,
                portfolio: portfolio,
                investments: { ...portfolioGlobals.fundInvestments },
                fundInvestments: { ...portfolioGlobals.fundInvestments },
                fundSparrates: { ...portfolioGlobals.fundSparrates }
            };
            const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(setup, null, 2));
            const dl = document.createElement('a');
            dl.setAttribute("href", dataStr);
            dl.setAttribute("download", `vem_setup_v6.0_${new Date().toISOString().slice(0,10)}.json`);
            document.body.appendChild(dl);
            dl.click();
            dl.remove();
        });
    }
});
