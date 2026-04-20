function showError(msg) { const el = document.getElementById('alertError'); el.querySelector('.msg').textContent = msg; el.classList.add('show'); }
function showWarn(msg) { const el = document.getElementById('alertWarn'); el.querySelector('.msg').textContent = msg; el.classList.add('show'); }

// Registrar plugin de etiquetas para todos os gráficos
Chart.register(ChartDataLabels);

/* ===== Util ===== */
const $ = s => document.querySelector(s);
const BRL = v => new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(+v || 0);
const NUM = v => new Intl.NumberFormat('pt-BR').format(+v || 0);
const PCT = v => (Number.isFinite(v) ? v.toFixed(2).replace('.', ',') : '—') + '%';

/* ===== Scroll estável ===== */
function withStableScroll(fn) {
    const x = window.pageXOffset || 0;
    const y = window.pageYOffset || 0;
    fn();
    requestAnimationFrame(() => {
        window.scrollTo(x, y);
        requestAnimationFrame(() => { window.scrollTo(x, y); });
    });
}

/* ===== Conversões ===== */
function toNumberBR(x) {
    if (x === null || x === undefined || x === '') return 0;
    if (typeof x === 'number') return x;
    let s = String(x).trim().replace(/[R$\s]/g, '');
    // Se tiver vírgula, assume formato BR (1.000,50 -> 1000.50)
    if (s.includes(',')) {
        s = s.replace(/\./g, '').replace(',', '.');
    }
    const n = parseFloat(s);
    return isNaN(n) ? 0 : n;
}

function parseDateBR(val) {
    if (!val) return null;
    if (typeof val === 'number') return parseDateExcel(val);
    const s = String(val).trim();
    // Tenta formato DD/MM/AAAA
    const m = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (m) return new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]));
    const d = new Date(s);
    return isNaN(+d) ? null : d;
}
function parseDateExcel(v) {
    if (typeof v === 'number') {
        const base = new Date(Date.UTC(1899, 11, 30));
        return new Date(base.getTime() + v * 86400000);
    }
    const d = new Date(v); return isNaN(+d) ? null : d;
}
function monthNameFromCell(cell) {
    const nomes = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];
    if (cell === null || cell === undefined || cell === '') return 'Não informado';

    if (typeof cell === 'string') {
        const m = cell.match(/\/(\d{2})\/(\d{4})/);
        if (m) {
            const mm = parseInt(m[1], 10);
            const aaaa = m[2];
            if (mm >= 1 && mm <= 12) {
                const nome = nomes[mm - 1];
                return (nome.charAt(0).toUpperCase() + nome.slice(1)) + ' ' + aaaa;
            }
        }
        const d = new Date(cell);
        if (!isNaN(+d)) {
            const nome = d.toLocaleDateString('pt-BR', { month: 'long' });
            const aaaa = d.getFullYear();
            return (nome.charAt(0).toUpperCase() + nome.slice(1)) + ' ' + aaaa;
        }
        return 'Não informado';
    }

    if (typeof cell === 'number') {
        const d = parseDateExcel(cell);
        if (d) {
            const nome = d.toLocaleDateString('pt-BR', { month: 'long' });
            const aaaa = d.getFullYear();
            return (nome.charAt(0).toUpperCase() + nome.slice(1)) + ' ' + aaaa;
        }
        return 'Não informado';
    }

    const d = new Date(cell);
    if (!isNaN(+d)) {
        const nome = d.toLocaleDateString('pt-BR', { month: 'long' });
        const aaaa = d.getFullYear();
        return (nome.charAt(0).toUpperCase() + nome.slice(1)) + ' ' + aaaa;
    }
    return 'Não informado';
}

function dayFromCell(cell) {
    if (cell === null || cell === undefined || cell === '') return null;
    if (typeof cell === 'number') {
        const d = parseDateExcel(cell);
        return d ? d.getDate() : null;
    }
    if (typeof cell === 'string') {
        const m = cell.match(/^(\d{2})\//);
        if (m) return parseInt(m[1], 10);
        const d = new Date(cell);
        return isNaN(+d) ? null : d.getDate();
    }
    const d = new Date(cell);
    return isNaN(+d) ? null : d.getDate();
}

function modelBase(txt) {
    if (!txt) return 'Não informado';
    const u = String(txt).toUpperCase();
    const cats = ['ARGO', 'CRONOS', 'MOBI', 'UNO', 'STRADA', 'TORO', 'FIORINO', 'DUCATO', 'HB20', 'HB20S', 'CRETA', 'GOL', 'POLO', 'VIRTUS', 'T-CROSS', 'GOLF', 'UP', 'PRISMA', 'ONIX', 'TRACKER', 'S10', 'ECOSPORT', 'KA', 'RANGER', 'CIVIC', 'CITY', 'FIT', 'COROLLA', 'HILUX', 'RAV4', 'PULSE', 'AGILE', 'AIRCROSS', 'ALL'];
    for (const c of cats) { if (u.includes(c)) return c; }
    return u.split(/\s+/).find(w => w.length > 2) || 'Outros';
}
function normalizeYear(val) {
    if (val === null || val === undefined || val === '') return null;
    const s = String(val).trim();
    const m4 = s.match(/\b(19|20)\d{2}\b/g);
    if (m4 && m4.length) return parseInt(m4.map(Number).sort((a, b) => b - a)[0], 10);
    const m2pair = s.match(/\b(\d{2})\s*[/\-]\s*(\d{2})\b/);
    if (m2pair) {
        const a = parseInt(m2pair[1], 10), b = parseInt(m2pair[2], 10);
        const two = Math.max(a, b);
        const now2 = new Date().getFullYear() % 100;
        const century = two <= now2 + 1 ? 2000 : 1900;
        return century + two;
    }
    const n = toNumberBR(s);
    if (Number.isFinite(n)) {
        if (n >= 1900 && n <= 2099) return Math.floor(n);
        if (n >= 0 && n < 100) {
            const now2 = new Date().getFullYear() % 100;
            const century = n <= now2 + 1 ? 2000 : 1900;
            return century + Math.floor(n);
        }
    }
    return null;
}
function normalizePlate(val) {
    if (val === null || val === undefined) return 'Não informado';
    let s = String(val).toUpperCase().trim();
    if (!s) return 'Não informado';
    const raw = s.replace(/[^A-Z0-9]/g, '');
    if (raw.length === 7) {
        if (/\d{4}$/.test(raw)) return raw.slice(0, 3) + '-' + raw.slice(3);
        return raw;
    }
    return s;
}

/* ====== NORMALIZADOR DE COR (usa coluna COR) ====== */
function normalizeColorName(raw) {
    if (raw === null || raw === undefined) return 'NÃO INFORMADA';
    let s = String(raw).toUpperCase();
    if (!s.trim()) return 'NÃO INFORMADA';

    // remove acentos simples
    s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, '');

    // mapeamento direto por palavras-chave (primeira cor que aparecer)
    const cores = [
        'PRETO', 'BRANCO', 'CINZA', 'PRATA', 'AZUL', 'VERMELHO', 'VERDE', 'AMARELO', 'MARROM', 'BEGE', 'DOURADO', 'LARANJA', 'ROXO', 'GRAFITE'
    ];
    for (const c of cores) {
        const re = new RegExp(`\\b${c}\\b`);
        if (re.test(s)) return c === 'GRAFITE' ? 'CINZA' : c;
    }

    // fallback: primeira palavra (muito útil para "AZUL AMALFI TETO PRE")
    const first = s.split(/\s+/)[0];
    return cores.includes(first) ? (first === 'GRAFITE' ? 'CINZA' : first) : first;
}

function findCol(headers, aliases) {
    const lc = headers.map(h => String(h || '').toLowerCase().trim());
    for (const a of aliases) {
        const idx = lc.indexOf(a.toLowerCase());
        if (idx >= 0) return idx;
    }
    return -1;
}

/* ===== Estado ===== */
const DOTACAO_CODES = [
    "155257016", "155256846", "155257015", "155256845", "155257959",
    "155257958", "155257960", "155257961", "155256245", "155255887",
    "155256390", "155256683", "155257695", "155256707", "9840660180",
    "9840659980", "9844133580", "155255862", "155255800", "9809532380",
    "7095074", "7095225", "7094994", "7095069", "7094991", "7094970",
    "7094971", "7095178", "7094996", "7094995", "7095416", "7094998"
];

let allData = [];
let filtered = [];
const state = { search: '', Grupo: [], Giro: [], CodItem: [], Descricao: [], Aging: [], Loja: [], Dotacao: [] };
const sortState = { key: 'Quantidade', type: 'num', dir: 'desc' };
let charts = { periodo: null, valorItem: null, loja: null, aging: null };

function getAgingLabel(r) {
    if (!r.UltimaCompraDate) return 'acima de 1000 dias';
    const today = new Date();
    const diffTime = today - r.UltimaCompraDate;
    const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
    if (diffDays <= 90) return 'de 0 a 90 dias';
    if (diffDays <= 180) return 'de 91 a 180 dias';
    if (diffDays <= 365) return 'de 181 a 365 dias';
    if (diffDays <= 1000) return 'de 366 a 1000 dias';
    return 'acima de 1000 dias';
}


/* ===== Pílulas ===== */
class MultiPill {
    constructor({ label, getter, setter }) {
        this.getter = getter; this.setter = setter;
        this.el = document.createElement('div');
        this.el.className = 'pill'; this.el.tabIndex = 0;
        this.el.innerHTML = `<strong>${label}:</strong> <span class="value">Todos</span> <i class="fa-solid fa-chevron-down chev"></i><div class="menu"></div>`;
        this.menu = this.el.querySelector('.menu'); this.valueEl = this.el.querySelector('.value');
        this.searchInput = null;

        this.el.addEventListener('click', e => {
            if (!e.target.closest('.menu')) {
                const x = window.pageXOffset, y = window.pageYOffset;
                this.el.classList.toggle('open');
                if (this.searchInput) this.searchInput.focus();
                window.scrollTo(x, y);
            }
        });
        document.addEventListener('click', e => {
            if (!this.el.contains(e.target)) {
                const x = window.pageXOffset, y = window.pageYOffset;
                this.el.classList.remove('open');
                window.scrollTo(x, y);
            }
        });
    }
    setOptions(opts, { keepOrder = false } = {}) {
        const unique = [...new Set(opts.filter(Boolean))];
        this.options = keepOrder ? unique : unique.sort((a, b) => String(a).localeCompare(String(b), 'pt-BR'));
        this.render();
    }
    render() {
        const sel = new Set(this.getter());
        this.menu.innerHTML = '';
        const s = document.createElement('input');
        s.placeholder = 'Pesquisar...'; s.className = 'menu-search'; this.menu.append(s); this.searchInput = s;
        const acts = document.createElement('div'); acts.className = 'menu-actions';
        const limpar = document.createElement('button'); limpar.textContent = 'Limpar';
        const aplicar = document.createElement('button'); aplicar.textContent = 'Aplicar'; aplicar.className = 'apply';
        acts.append(limpar, aplicar); this.menu.append(acts);
        const list = document.createElement('div'); list.className = 'list'; this.menu.append(list);

        for (const o of this.options) {
            const row = document.createElement('label'); row.className = 'opt'; row.dataset.text = String(o).toLowerCase();
            const chk = document.createElement('input'); chk.type = 'checkbox'; chk.value = o; chk.checked = sel.has(o);
            row.append(chk, document.createElement('span')); row.lastChild.textContent = o;
            list.append(row);
        }

        s.oninput = () => {
            const q = s.value.toLowerCase();
            list.querySelectorAll('.opt').forEach(el => {
                el.style.display = el.dataset.text.includes(q) ? '' : 'none';
            });
        };

        limpar.onclick = () => withStableScroll(() => {
            this.setter([]);
            this.sync();
            applyFilters();
        });
        aplicar.onclick = () => withStableScroll(() => {
            const vals = [...list.querySelectorAll('input:checked')].map(i => i.value);
            this.setter(vals);
            this.el.classList.remove('open');
            this.sync();
            applyFilters();
        });

        list.addEventListener('click', (ev) => {
            const input = ev.target.closest('input[type="checkbox"]');
            if (!input) return;
            const temp = new Set(this.getter());
            if (input.checked) temp.add(input.value); else temp.delete(input.value);
            const arr = [...temp];
            this.valueEl.textContent = !arr.length ? 'Todos' : (arr.length <= 2 ? arr.join(', ') : `${arr[0]}, ${arr[1]} (+${arr.length - 2})`);
        });

        this.sync();
    }
    sync() {
        const v = this.getter();
        this.valueEl.textContent = !v || v.length === 0 ? 'Todos' : (v.length <= 2 ? v.join(', ') : `${v[0]}, ${v[1]} (+${v.length - 2})`);
    }
}
const pillInstances = {
    Grupo: new MultiPill({ label: 'Grupo/Classe', getter: () => state.Grupo, setter: v => state.Grupo = v }),
    CodItem: new MultiPill({ label: 'Código do Item', getter: () => state.CodItem, setter: v => state.CodItem = v }),
    Descricao: new MultiPill({ label: 'Descrição', getter: () => state.Descricao, setter: v => state.Descricao = v }),
    Aging: new MultiPill({ label: 'Dias de Estoque', getter: () => state.Aging, setter: v => state.Aging = v }),
    Loja: new MultiPill({ label: 'Empresa', getter: () => state.Loja, setter: v => state.Loja = v }),
    Dotacao: new MultiPill({ label: 'Dotação', getter: () => state.Dotacao, setter: v => state.Dotacao = v }),
};
(function mountPills() {
    const holder = document.getElementById('filters');
    ['Loja', 'CodItem', 'Grupo', 'Descricao', 'Dotacao', 'Aging'].forEach(k => {
        if (pillInstances[k]) holder.appendChild(pillInstances[k].el);
    });
})();

/* ===== Upload ===== */
const fileInput = document.getElementById('fileInput');
const fileNameEl = document.getElementById('fileName');
fileInput.addEventListener('change', ev => { const f = ev.target.files?.[0]; if (f) { fileNameEl.textContent = f.name; readFile(f); } });

/* ===== Parser ===== */
async function readFile(file) {
    try {
        if (!window.XLSX) { showError('Biblioteca XLSX não carregada. Verifique sua internet.'); return; }
        const ext = (file.name.split('.').pop() || '').toLowerCase();
        if (!['xlsx', 'xls', 'csv'].includes(ext)) {
            showWarn('Formato não suportado. Use .xlsx, .xls ou .csv'); return;
        }
        if (ext === 'csv') {
            const text = await file.text();
            const wb = XLSX.read(text, { type: 'string' });
            processWorkbook(wb);
        } else {
            const buf = await file.arrayBuffer();
            const wb = XLSX.read(new Uint8Array(buf), { type: 'array' });
            processWorkbook(wb);
        }
    } catch (err) {
        console.error(err);
        showError('Falha ao processar arquivo: ' + (err.message || 'Erro desconhecido') + '. Se o arquivo estiver aberto no Excel, feche-o e tente novamente.');
    }
}

/* ===== Processamento ===== */
function processWorkbook(wb) {
    // Tenta encontrar a primeira aba com dados
    let ws = null;
    let sheetName = "";
    for (const name of wb.SheetNames) {
        const testWs = wb.Sheets[name];
        const range = XLSX.utils.decode_range(testWs['!ref'] || 'A1:A1');
        if ((range.e.r - range.s.r) >= 0) { // pelo menos 1 linha
            ws = testWs;
            sheetName = name;
            break;
        }
    }

    if (!ws) { showError('A planilha parece estar vazia em todas as abas.'); return; }

    const mat = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false });
    if (!mat || !mat.length) { showError('Não encontrei linhas com dados na aba "' + sheetName + '".'); return; }

    console.log('Dados lidos (matriz):', mat.slice(0, 5)); // Log para debug

    // Detecta se a primeira linha parece ser um cabeçalho ou dado direto
    let headerIdx = -1;
    for (let i = 0; i < Math.min(5, mat.length); i++) {
        const row = mat[i];
        // Se a linha tem textos como "Código", "Descrição", "Item", assumimos que é cabeçalho
        const rowStr = JSON.stringify(row).toLowerCase();
        if (rowStr.includes('código') || rowStr.includes('item') || rowStr.includes('descrição') || rowStr.includes('estoque')) {
            headerIdx = i;
            break;
        }
    }

    let headers = [];
    let dataRows = [];

    if (headerIdx > -1) {
        headers = mat[headerIdx].map(h => String(h || '').trim());
        dataRows = mat.slice(headerIdx + 1);
    } else {
        // Se não detectou cabeçalho, assume que os dados começam na linha 0
        headers = [];
        dataRows = mat;
    }

    parseRows(headers, dataRows);
}

/* ===== Leitura das linhas ===== */
function parseRows(headers, data) {
    // Mapeamento Flexível: Tenta achar pelo nome, se não achar usa a letra da coluna (A=0, B=1...)
    const getIdx = (aliases, defaultIdx) => {
        const found = findCol(headers, aliases);
        return found >= 0 ? found : defaultIdx;
    };

    const idx = {
        codItem: getIdx(['Código do Item', 'Código'], 0), // A
        descricao: getIdx(['Descrição'], 1),              // B
        quantidade: getIdx(['Quantidade da peça em estoque', 'Qtd', 'Quantidade'], 3), // D
        giro: getIdx(['Giro'], 2), 
        precoCusto: getIdx(['Preço de Custo', 'Custo'], 4), 
        precoVenda: getIdx(['Preço de Venda', 'Venda'], 7), // H
        ultimaVenda: getIdx(['Última Venda', 'Data Última Venda'], 18), // S
        loja: getIdx(['Nome da Empresa', 'Empresa', 'Loja'], 11), // L
        ultimaCompra: getIdx(['Data da compra da peça', 'Última Compra'], 17), // R
        grupo: getIdx(['Grupo'], 19) // T
    };

    allData = data.filter(r => r.length > 0).map(r => {
        const get = (i) => r[i];
        
        const fmtDate = (val) => {
            if (!val) return '—';
            if (typeof val === 'number') {
                // SheetJS às vezes lê datas como números de série do Excel
                try {
                    return parseDateExcel(val).toLocaleDateString('pt-BR');
                } catch(e) { return String(val); }
            }
            return String(val);
        };


        return {
            CodItem: get(idx.codItem) ?? '—',
            Descricao: get(idx.descricao) ?? '—',
            Quantidade: toNumberBR(get(idx.quantidade)),
            Giro: toNumberBR(get(idx.giro)),
            PrecoCusto: toNumberBR(get(idx.precoCusto)),
            PrecoVenda: toNumberBR(get(idx.precoVenda)),
            UltimaVenda: fmtDate(get(idx.ultimaVenda)),
            UltimaCompra: fmtDate(get(idx.ultimaCompra)),
            UltimaCompraDate: parseDateBR(get(idx.ultimaCompra)),
            Grupo: get(idx.grupo) ?? 'Outros',
            Loja: get(idx.loja) ?? 'Não informada'
        };
    });

    if (allData.length === 0) {
        showWarn('A planilha foi lida, mas nenhuma linha de dado válida foi encontrada.');
        return;
    }

    pillInstances.Grupo.setOptions(allData.map(x => x.Grupo));
    pillInstances.CodItem.setOptions(allData.map(x => x.CodItem));
    pillInstances.Descricao.setOptions(allData.map(x => x.Descricao));
    pillInstances.Loja.setOptions(allData.map(x => x.Loja));
    pillInstances.Aging.setOptions([
        'de 0 a 90 dias', 'de 91 a 180 dias', 'de 181 a 365 dias', 'de 366 a 1000 dias', 'acima de 1000 dias'
    ], { keepOrder: true });
    pillInstances.Dotacao.setOptions(['Sim', 'Não'], { keepOrder: true });

    applyFilters();
}

/* ===== Busca e Limpar ===== */
document.getElementById('q').addEventListener('input', e => { state.search = e.target.value.toLowerCase(); withStableScroll(() => applyFilters()); });
document.getElementById('clearAll').addEventListener('click', () => {
    withStableScroll(() => {
        state.search = ''; document.getElementById('q').value = '';
        state.Grupo = []; state.Giro = []; state.CodItem = []; state.Descricao = []; state.Loja = []; state.Aging = []; state.Dotacao = [];
        Object.values(pillInstances).forEach(p => { p.setter([]); p.render(); p.sync(); });
        applyFilters();
    });
});
document.getElementById('exportExcel').addEventListener('click', () => {
    if (!filtered || filtered.length === 0) {
        showWarn('Não há dados para exportar.');
        return;
    }
    
    const sortedData = sortRows([...filtered]);
    const dataToExport = sortedData.map(r => ({
        'Cód. Item': r.CodItem,
        'Descrição': r.Descricao,
        'Quantidade': r.Quantidade,
        'Giro': r.Giro,
        'Preço Custo': r.PrecoCusto,
        'Preço Venda': r.PrecoVenda,
        'Últ. Venda': r.UltimaVenda,
        'Últ. Compra': r.UltimaCompra,
        'Dias de Estoque': getAgingLabel(r),
        'Grupo/Classe': r.Grupo,
        'Empresa': r.Loja
    }));

    const ws = XLSX.utils.json_to_sheet(dataToExport);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Estoque");
    
    // Auto-size columns based on header length and content max length
    const cols = Object.keys(dataToExport[0] || {}).map(k => ({ wch: Math.max(k.length, 10) }));
    ws['!cols'] = cols;

    XLSX.writeFile(wb, "estoque_analise.xlsx");
});

function applyFilters() {
    filtered = allData.filter(r => {
        const grupoOk = state.Grupo.length === 0 || state.Grupo.includes(r.Grupo);
        const giroOk = state.Giro.length === 0 || state.Giro.includes(r.Giro);
        const codOk = state.CodItem.length === 0 || state.CodItem.includes(r.CodItem);
        const descOk = state.Descricao.length === 0 || state.Descricao.includes(r.Descricao);
        const lojaOk = state.Loja.length === 0 || state.Loja.includes(r.Loja);
        const agingOk = state.Aging.length === 0 || state.Aging.includes(getAgingLabel(r));

        let dotacaoOk = true;
        if (state.Dotacao.length === 1) {
            const isDotacao = DOTACAO_CODES.includes(String(r.CodItem));
            if (state.Dotacao.includes("Sim") && !isDotacao) dotacaoOk = false;
            if (state.Dotacao.includes("Não") && isDotacao) dotacaoOk = false;
        }

        const q = state.search;
        const qOk = !q || [r.Descricao, r.CodItem, r.Grupo, r.Loja].some(vv => (vv || '').toString().toLowerCase().includes(q));

        return grupoOk && giroOk && codOk && descOk && lojaOk && agingOk && dotacaoOk && qOk;
    });
    renderKpiGroups();
    renderTable(sortRows([...filtered]));
    renderCharts(filtered);
}

/* ===== KPIs (cálculo) ===== */
function calcKPIs(rows) {
    const totalItens = rows.reduce((s, r) => s + (+r.Quantidade || 0), 0);
    const totalVariedade = rows.length; 
    const valorCustoTotal = rows.reduce((s, r) => s + ((+r.Quantidade || 0) * (+r.PrecoCusto || 0)), 0);
    const valorVendaTotal = rows.reduce((s, r) => s + ((+r.Quantidade || 0) * (+r.PrecoVenda || 0)), 0);

    return {
        totalItens, totalVariedade, valorCustoTotal, valorVendaTotal
    };
}


/* ===== KPIs AGRUPADOS ===== */
function renderKpiGroups() {
    const k = calcKPIs(filtered);

    const gruposHTML = `
<div class="kpi-group">

<div class="kpi-grid">
    <div class="kpi"><div class="lab">Total de Peças</div><div class="val">${NUM(k.totalItens)}</div></div>
    <div class="kpi"><div class="lab">Quantidade de Itens</div><div class="val">${NUM(k.totalVariedade)}</div></div>
    <div class="kpi"><div class="lab">Valor Total (Custo)</div><div class="val">${BRL(k.valorCustoTotal)}</div></div>
    <div class="kpi"><div class="lab">Valor Total (Venda)</div><div class="val">${BRL(k.valorVendaTotal)}</div></div>
</div>
</div>
`;
    document.getElementById('kpiGroups').innerHTML = gruposHTML;
}

/* ===== Ordenação/Tabela ===== */
const collator = new Intl.Collator('pt-BR', { numeric: true, sensitivity: 'base' });
function sortRows(rows) {
    if (!sortState.key) return rows;
    const key = sortState.key, type = sortState.type, dir = sortState.dir === 'asc' ? 1 : -1;
    return rows.sort((a, b) => {
        let va = a[key], vb = b[key];
        if (type === 'num') {
            va = Number(va) || 0; vb = Number(vb) || 0;
            return dir * (va - vb);
        } else {
            return dir * collator.compare(String(va || ''), String(vb || ''));
        }
    });
}
function renderTable(rows) {
    const tb = document.querySelector('#grid tbody'); tb.innerHTML = '';
    for (const r of rows) {
        const tr = document.createElement('tr');
        tr.innerHTML = `
    <td>${r.CodItem || '—'}</td>
    <td>${r.Descricao || '—'}</td>
    <td>${NUM(r.Quantidade)}</td>
    <td>${NUM(r.Giro)}</td>
    <td>${BRL(r.PrecoCusto)}</td>
    <td>${BRL(r.PrecoVenda)}</td>
    <td>${r.UltimaVenda || '—'}</td>
    <td>${r.UltimaCompra || '—'}</td>
    <td>${r.Grupo || '—'}</td>
    <td>${r.Loja || '—'}</td>
`;
        tb.appendChild(tr);
    }
}
function initSorting() {
    const ths = document.querySelectorAll('#grid thead th.sortable');
    ths.forEach(th => {
        th.addEventListener('click', () => withStableScroll(() => {
            const key = th.dataset.key, type = th.dataset.type;
            if (sortState.key === key) {
                sortState.dir = sortState.dir === 'asc' ? 'desc' : 'asc';
            } else {
                sortState.key = key; sortState.type = type; sortState.dir = 'desc';
            }
            ths.forEach(t => { t.classList.remove('active'); t.querySelector('.sort').className = 'fa-solid fa-sort sort'; });
            th.classList.add('active');
            const icon = th.querySelector('.sort');
            icon.className = 'fa-solid ' + (sortState.dir === 'asc' ? 'fa-sort-up' : 'fa-sort-down') + ' sort';
            renderTable(sortRows([...filtered]));
        }));
    });
}
initSorting();


/* Tooltips e datalabels */
function pctTooltip() {
    return {
        callbacks: {
            label: function (ctx) {
                const val = Number(ctx.raw) || 0;
                const data = ctx.dataset.data || [];
                const total = data.reduce((a, b) => a + (Number(b) || 0), 0) || 1;
                const pct = ((val / total) * 100).toFixed(1).replace('.', ',');
                return `${val} (${pct}%)`;
            }
        }
    };
}

function datalabelsCenter() {
    return {
        color: '#ffffff',
        font: { weight: '700' },
        anchor: 'center',
        align: 'center',
        clip: true,
        formatter: (v, ctx) => {
            const data = ctx.dataset.data || [];
            const total = data.reduce((a, b) => a + (Number(b) || 0), 0) || 1;
            const pct = ((v / total) * 100).toFixed(1).replace('.', ',');
            return `${v} (${pct}%)`;
        }
    };
}

/* gráficos */
function renderCharts(rows) {
    const commonOpts = { responsive: true, maintainAspectRatio: false, animation: false, plugins: { datalabels: {} } };

    // 1. Estoque por Grupo e por Empresa
    const statsByGroup = {};
    const statsByLoja = {};
    rows.forEach(r => {
        const g = r.Grupo || 'Outros';
        if (!statsByGroup[g]) statsByGroup[g] = { qtd: 0, custo: 0 };
        statsByGroup[g].qtd += (+r.Quantidade || 0);
        statsByGroup[g].custo += ((+r.Quantidade || 0) * (+r.PrecoCusto || 0));

        const l = r.Loja || 'Não informada';
        if (!statsByLoja[l]) statsByLoja[l] = { qtd: 0, custo: 0 };
        statsByLoja[l].qtd += (+r.Quantidade || 0);
        statsByLoja[l].custo += ((+r.Quantidade || 0) * (+r.PrecoCusto || 0));
    });

    const groupsSorted = Object.entries(statsByGroup).sort((a, b) => b[1].qtd - a[1].qtd);

    if (charts.periodo) charts.periodo.destroy();
    charts.periodo = new Chart(document.getElementById('periodoChart').getContext('2d'), {
        type: 'bar',
        data: { 
            labels: groupsSorted.map(x => x[0]), 
            datasets: [{ 
                label: 'Quantidade Total', 
                data: groupsSorted.map(x => x[1].qtd), 
                backgroundColor: '#10b981',
                borderRadius: 8
            }] 
        },
        options: { 
            ...commonOpts, 
            plugins: { 
                ...commonOpts.plugins, 
                legend: { display: false },
                title: { display: true, text: 'Quantidade de Itens por Grupo' },
                tooltip: {
                    callbacks: {
                        label: (ctx) => {
                            const grupo = groupsSorted[ctx.dataIndex];
                            return [
                                `Quantidade: ${NUM(ctx.raw)}`,
                                `${BRL(grupo[1].custo)}`
                            ];
                        }
                    }
                },
                datalabels: {
                    display: true,
                    anchor: 'end',
                    align: 'top',
                    color: '#333',
                    textAlign: 'center',
                    font: { weight: 'bold', size: 10 },
                    formatter: (v, ctx) => {
                        const grupo = groupsSorted[ctx.dataIndex];
                        return `${NUM(v)}\n${BRL(grupo[1].custo)}`;
                    }
                }
            } 
        }
    });

    // 2. Investimento por Empresa (Custo Total)
    const lojasSorted = Object.entries(statsByLoja).sort((a, b) => b[1].custo - a[1].custo);
    const totalCustoLojas = Object.values(statsByLoja).reduce((acc, curr) => acc + curr.custo, 0) || 1;

    if (charts.loja) charts.loja.destroy();
    charts.loja = new Chart(document.getElementById('lojaChart').getContext('2d'), {
        type: 'bar',
        data: { 
            labels: lojasSorted.slice(0, 10).map(x => x[0]), 
            datasets: [{ 
                label: 'Investimento (Custo)',
                data: lojasSorted.slice(0, 10).map(x => x[1].custo),
                backgroundColor: '#3b82f6',
                borderRadius: 8
            }] 
        },
        options: { 
            ...commonOpts,
            indexAxis: 'y',
            plugins: { 
                ...commonOpts.plugins, 
                legend: { display: false },
                title: { display: true, text: 'Estoque por Empresa' },
                tooltip: {
                    callbacks: {
                        label: (ctx) => {
                            const val = Number(ctx.raw) || 0;
                            const pct = ((val / totalCustoLojas) * 100).toFixed(1).replace('.', ',');
                            return `${BRL(val)} (${pct}%)`;
                        }
                    }
                },
                datalabels: {
                    anchor: 'end',
                    align: 'right',
                    color: '#444',
                    font: { weight: 'bold', size: 10 },
                    formatter: (v) => BRL(v)
                }
            },
            layout: { padding: { right: 70 } }
        }
    });

    // 3. Top 10 Itens por Valor Total em Estoque
    const itemValues = rows.map(r => ({
        desc: r.Descricao,
        cod: r.CodItem,
        valor: (+r.Quantidade || 0) * (+r.PrecoCusto || 0),
        qtd: +r.Quantidade || 0
    })).sort((a, b) => b.valor - a.valor).slice(0, 10);

    if (charts.valorItem) charts.valorItem.destroy();
    charts.valorItem = new Chart(document.getElementById('valorItemChart').getContext('2d'), {
        type: 'bar',
        data: { 
            labels: itemValues.map(x => x.desc), 
            datasets: [{ 
                label: 'Valor Total (Custo)', 
                data: itemValues.map(x => x.valor), 
                backgroundColor: '#8b5cf6', 
                borderRadius: 8 
            }] 
        },
        options: { 
            ...commonOpts, 
            indexAxis: 'y', 
            plugins: { 
                ...commonOpts.plugins, 
                legend: { display: false }, 
                title: { display: true, text: 'Top 10 Itens por Valor Total' }, 
                datalabels: { 
                    anchor: 'end',
                    align: 'start',
                    color: '#fff',
                    font: { weight: 'bold', size: 10 },
                    formatter: (v, ctx) => {
                        const item = itemValues[ctx.dataIndex];
                        return `${BRL(v)} (Qtd: ${NUM(item.qtd)})`;
                    }
                },
                tooltip: {
                    callbacks: {
                        label: (ctx) => {
                            const item = itemValues[ctx.dataIndex];
                            return [
                                `Código: ${item.cod}`,
                                `Quantidade: ${NUM(item.qtd)}`,
                                `Total: ${BRL(ctx.raw)}`
                            ];
                        }
                    }
                } 
            } 
        }
    });

    // 4. Envelhecimento de Estoque (Aging) - Por Valor de Custo
    const today = new Date();
    const buckets = [
        { label: 'de 0 a 90 dias', min: 0, max: 90, custo: 0, qtd: 0 },
        { label: 'de 91 a 180 dias', min: 91, max: 180, custo: 0, qtd: 0 },
        { label: 'de 181 a 365 dias', min: 181, max: 365, custo: 0, qtd: 0 },
        { label: 'de 366 a 1000 dias', min: 366, max: 1000, custo: 0, qtd: 0 },
        { label: 'acima de 1000 dias', min: 1001, max: Infinity, custo: 0, qtd: 0 }
    ];

    rows.forEach(r => {
        const custoTotal = (+r.Quantidade || 0) * (+r.PrecoCusto || 0);
        const qtdItem = (+r.Quantidade || 0);
        if (!r.UltimaCompraDate) {
            buckets[4].custo += custoTotal; 
            buckets[4].qtd += qtdItem;
            return;
        }
        const diffTime = today - r.UltimaCompraDate;
        const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
        
        const b = buckets.find(b => diffDays >= b.min && diffDays <= b.max);
        if (b) {
            b.custo += custoTotal;
            b.qtd += qtdItem;
        } else if (diffDays > 1000) {
            buckets[4].custo += custoTotal;
            buckets[4].qtd += qtdItem;
        }
    });

    const totalCustoAging = buckets.reduce((acc, curr) => acc + curr.custo, 0) || 1;

    if (charts.aging) charts.aging.destroy();
    charts.aging = new Chart(document.getElementById('agingChart').getContext('2d'), {
        type: 'bar',
        data: {
            labels: buckets.map(b => b.label),
            datasets: [{
                label: '',
                data: buckets.map(b => b.custo),
                backgroundColor: ['#10b981', '#3b82f6', '#f59e0b', '#ef4444', '#71717a'],
                borderRadius: 8
            }]
        },
        options: {
            ...commonOpts,
            indexAxis: 'y',
            plugins: {
                ...commonOpts.plugins,
                legend: { display: false },
                title: { display: true, text: 'Dias de Estoque' },
                tooltip: {
                    callbacks: {
                        label: (ctx) => {
                            const b = buckets[ctx.dataIndex];
                            const pct = ((b.custo / totalCustoAging) * 100).toFixed(1).replace('.', ',');
                            return [
                                `${BRL(b.custo)}`,
                                `Quantidade: ${NUM(b.qtd)} peças`,
                                `Participação: ${pct}% do total`
                            ];
                        }
                    }
                },
                datalabels: {
                    anchor: 'end',
                    align: 'right',
                    color: '#475569',
                    offset: 8,
                    font: { weight: 'bold', size: 11 },
                    formatter: v => v > 0 ? BRL(v) : ''
                }
            },
            layout: {
                padding: { right: 80, left: 10 }
            }
        }
    });

    // Desativar outros gráficos que não são mais necessários ou não têm dados correspondentes
    ['modeloChart', 'faixasPrecoChart', 'vendedorChart', 'vendedorValorChart', 'tipoOsChart', 'vendasPorDiaChart'].forEach(id => {
        const canvas = document.getElementById(id);
        if (canvas) canvas.parentElement.style.display = 'none';
    });
}
