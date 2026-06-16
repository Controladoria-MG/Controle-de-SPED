// ================================================
// CONFIG
// ================================================
const CAMINHO_ARQUIVO = 'data/base_formatada/Base.xlsx';

let dadosProcessados = null;
let unidadeAtual = null;
let currentModalData = [];

// ================================================
// DATAS ÚTEIS DE JUNHO 2026
// ================================================
function getDiasUteisJunho2026() {
    const dias = [];
    for (let i = 1; i <= 30; i++) {
        const d = new Date(2026, 5, i); // mês 5 = junho
        if (d.getDay() !== 0 && d.getDay() !== 6) dias.push(d);
    }
    return dias;
}

function formatarData(d) {
    return d.toLocaleDateString('pt-BR', { day: '2-digit', month: '2-digit' });
}

function formatarDataCompleta(d) {
    if (!d) return '-';
    return d.toLocaleDateString('pt-BR');
}

function isMesmaData(a, b) {
    return a && b &&
        a.getFullYear() === b.getFullYear() &&
        a.getMonth() === b.getMonth() &&
        a.getDate() === b.getDate();
}

function extrairData(valor) {
    if (!valor) return null;
    try {
        if (typeof valor === 'number') {
            const d = XLSX.SSF.parse_date_code(valor);
            return new Date(d.y, d.m - 1, d.d);
        }
        if (valor instanceof Date) {
            return new Date(valor.getFullYear(), valor.getMonth(), valor.getDate());
        }
        if (typeof valor === 'string') {
            const partes = valor.split(' ')[0];
            const [dia, mes, ano] = partes.split('/');
            if (dia && mes && ano) return new Date(+ano, mes - 1, +dia);
            const d = new Date(valor);
            if (!isNaN(d)) return d;
        }
    } catch (e) {}
    return null;
}

function ehDataFutura(data) {
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);
    const d = new Date(data);
    d.setHours(0, 0, 0, 0);
    return d > hoje;
}

// ================================================
// CARREGAR ARQUIVO
// ================================================
async function carregarArquivo() {
    const response = await fetch(CAMINHO_ARQUIVO);
    if (!response.ok) throw new Error(`Arquivo não encontrado (HTTP ${response.status})`);
    const buffer = await response.arrayBuffer();
    const wb = XLSX.read(buffer, { type: 'array', cellDates: true });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(sheet, { raw: true });
}

// ================================================
// AUXILIARES DE DADOS
// ================================================
function getUnidades(dados) {
    const set = new Set(dados.map(r => r.Unidade).filter(Boolean));
    return [...set].sort();
}

function filtrarPorUnidade(dados, unidade) {
    return dados.filter(r => r.Unidade === unidade);
}

function agruparPorDepartamento(dados) {
    const grupos = {};
    dados.forEach(r => {
        const dept = String(r.Departamento || 'Sem Departamento').trim();
        if (!grupos[dept]) grupos[dept] = [];
        grupos[dept].push(r);
    });
    return grupos;
}

// ================================================
// CÁLCULO DE EVOLUÇÃO
// ================================================
function calcularEvolucao(base, dias) {
    const total = base.length;
    let pendente = total;
    const arr = [total]; // coluna Total

    dias.forEach(dia => {
        const baixados = base.filter(r => {
            const d = extrairData(r.DataAlteracaoEstagio);
            return d && isMesmaData(d, dia);
        }).length;
        pendente -= baixados;
        arr.push(pendente);
    });

    return arr; // arr[0]=Total, arr[1]=após dia 1, arr[2]=após dia 2...
}

// Retorna os registros ainda pendentes ao fim do dia informado
function getPendentesNaData(base, data) {
    if (!data) return [...base]; // coluna Total: todos pendentes no início
    return base.filter(r => {
        const d = extrairData(r.DataAlteracaoEstagio);
        return !d || d > data; // sem data de transmissão OU transmitido depois deste dia
    });
}

// ================================================
// LINHA DE PERCENTUAL COM BARRA
// ================================================
function criarLinhaPercentual(label, valores, tipo, dias) {
    const tr = document.createElement('tr');

    const tdLabel = document.createElement('td');
    tdLabel.textContent = label;
    tr.appendChild(tdLabel);

    valores.forEach((v, i) => {
        const td = document.createElement('td');
        const dataColuna = i === 0 ? null : dias[i - 1];
        const futura = dataColuna && ehDataFutura(dataColuna);

        if (futura) {
            td.className = 'celula-futura';
            tr.appendChild(td);
            return;
        }

        const wrapper = document.createElement('div');
        wrapper.className = 'percent-wrapper';

        const barBg = document.createElement('div');
        barBg.className = `percent-bar-bg ${tipo === 'danger' ? 'bar-danger' : 'bar-success'}`;
        barBg.style.width = `${Math.min(Math.round(v), 100)}%`;

        const valueSpan = document.createElement('span');
        valueSpan.className = 'percent-value';
        valueSpan.textContent = `${Math.round(v)}%`;

        wrapper.appendChild(barBg);
        wrapper.appendChild(valueSpan);
        td.appendChild(wrapper);
        tr.appendChild(td);
    });

    return tr;
}

// ================================================
// CRIAÇÃO DO BLOCO POR DEPARTAMENTO
// ================================================
function criarBloco(nomeDepto, dados, dias) {
    const total = dados.length;
    const evolucao = calcularEvolucao(dados, dias);

    const percPend = evolucao.map(v => (v / total) * 100);
    const percConc = evolucao.map(v => ((total - v) / total) * 100);

    const container = document.createElement('div');
    container.className = 'section-card';

    // Cabeçalho do card
    const header = document.createElement('div');
    header.className = 'section-header';
    header.innerHTML = `<h2>${nomeDepto} <span class="badge-total">${total} SPED${total !== 1 ? 's' : ''}</span></h2>`;
    container.appendChild(header);

    const scrollWrapper = document.createElement('div');
    scrollWrapper.className = 'scroll-wrapper';

    const tabela = document.createElement('table');
    tabela.className = 'dashboard-table';

    // Cabeçalho da tabela: Indicador | Total | datas...
    const thead = document.createElement('thead');
    const trHead = document.createElement('tr');

    const thInd = document.createElement('th');
    thInd.textContent = 'Indicador';
    trHead.appendChild(thInd);

    const thTotal = document.createElement('th');
    thTotal.textContent = 'Total';
    trHead.appendChild(thTotal);

    dias.forEach(d => {
        const th = document.createElement('th');
        th.textContent = formatarData(d);
        if (ehDataFutura(d)) th.classList.add('th-futura');
        trHead.appendChild(th);
    });

    thead.appendChild(trHead);
    tabela.appendChild(thead);

    const tbody = document.createElement('tbody');

    // Linha: Pendentes
    const trPend = document.createElement('tr');
    trPend.className = 'linha-pendentes';

    const tdPendLabel = document.createElement('td');
    tdPendLabel.textContent = 'Pendentes';
    trPend.appendChild(tdPendLabel);

    // evolucao[0] = Total, evolucao[1..] = por dia
    evolucao.forEach((v, i) => {
        const td = document.createElement('td');
        const dataColuna = i === 0 ? null : dias[i - 1];
        const futura = dataColuna && ehDataFutura(dataColuna);

        if (futura) {
            td.className = 'celula-futura';
        } else {
            td.textContent = v;
            td.className = 'clickable';
            td.title = `Clique para ver os ${v} SPED(s) pendente(s)`;
            td.onclick = () => {
                const lista = getPendentesNaData(dados, dataColuna);
                const titulo = dataColuna
                    ? `${nomeDepto} — Pendentes em ${formatarData(dataColuna)}`
                    : `${nomeDepto} — Total de SPEDs`;
                abrirModal(lista, titulo);
            };
        }

        trPend.appendChild(td);
    });

    tbody.appendChild(trPend);

    // Linha: % Pendente
    tbody.appendChild(criarLinhaPercentual('% Pendente', percPend, 'danger', dias));

    // Linha: % Entregue
    tbody.appendChild(criarLinhaPercentual('% Entregue', percConc, 'success', dias));

    tabela.appendChild(tbody);
    scrollWrapper.appendChild(tabela);
    container.appendChild(scrollWrapper);

    return container;
}

// ================================================
// ATUALIZAR DASHBOARD
// ================================================
function atualizarDashboard(dados, dias, unidade) {
    const container = document.getElementById('dashboards-container');
    container.innerHTML = '';

    if (!unidade) return;

    const filtrados = filtrarPorUnidade(dados, unidade);

    if (filtrados.length === 0) {
        container.innerHTML = '<p class="sem-dados">Nenhum dado encontrado para esta unidade.</p>';
        return;
    }

    const grupos = agruparPorDepartamento(filtrados);
    Object.keys(grupos)
        .sort()
        .forEach(nome => container.appendChild(criarBloco(nome, grupos[nome], dias)));
}

function criarBotoesUnidade(unidades, dados, dias) {
    const container = document.getElementById('unidades-buttons');
    container.innerHTML = '';

    unidades.forEach((unidade, idx) => {
        const btn = document.createElement('button');
        btn.className = 'btn-unidade' + (idx === 0 ? ' active' : '');
        btn.textContent = unidade;
        btn.setAttribute('data-unidade', unidade);
        btn.addEventListener('click', () => {
            document.querySelectorAll('.btn-unidade').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            unidadeAtual = unidade;
            atualizarDashboard(dados, dias, unidade);
        });
        container.appendChild(btn);
    });
}

// ================================================
// MODAL
// ================================================
function abrirModal(lista, titulo) {
    currentModalData = lista;

    const titleEl = document.getElementById('modal-title');
    if (titleEl) titleEl.textContent = titulo;

    renderModalTable(lista);

    const searchInput = document.getElementById('modalSearchInput');
    if (searchInput) searchInput.value = '';

    document.getElementById('modal').style.display = 'block';
}

function renderModalTable(lista) {
    const tbody = document.getElementById('modalTableBody');
    tbody.innerHTML = '';

    if (lista.length === 0) {
        const tr = document.createElement('tr');
        tr.innerHTML = '<td colspan="9" style="text-align:center;padding:30px;color:#64748b">Nenhum SPED pendente</td>';
        tbody.appendChild(tr);
        return;
    }

    lista.forEach(r => {
        const dataTransmissao = extrairData(r.DataAlteracaoEstagio);
        const tr = document.createElement('tr');
        tr.className = 'modal-table-row';
        tr.innerHTML = `
            <td>${r.IdCliente || '-'}</td>
            <td>${r.RazaoSocial || '-'}</td>
            <td>${r.Grupo || '-'}</td>
            <td>${r.GerenteDeContas || '-'}</td>
            <td>${r.Tributacao || '-'}</td>
            <td>${r.ResponsavelPadrao || '-'}</td>
            <td>${r.Status || '-'}</td>
            <td>${dataTransmissao ? formatarDataCompleta(dataTransmissao) : '-'}</td>
            <td>${r['Coordenador(a)'] || '-'}</td>
        `;
        tbody.appendChild(tr);
    });
}

function setupModalSearch() {
    const searchInput = document.getElementById('modalSearchInput');
    if (!searchInput) return;

    searchInput.addEventListener('input', e => {
        const term = e.target.value.toLowerCase().trim();
        document.querySelectorAll('#modalTableBody .modal-table-row').forEach(row => {
            row.classList.toggle('hidden', term !== '' && !row.textContent.toLowerCase().includes(term));
        });
    });
}

// ================================================
// EXPORTAR EXCEL
// ================================================
function exportToExcel() {
    if (!currentModalData.length) {
        alert('Nenhum dado para exportar.');
        return;
    }

    const exportData = currentModalData.map(r => {
        const d = extrairData(r.DataAlteracaoEstagio);
        return {
            'ID Cliente': r.IdCliente || '-',
            'Cliente': r.RazaoSocial || '-',
            'Grupo': r.Grupo || '-',
            'Gerente': r.GerenteDeContas || '-',
            'Tributação': r.Tributacao || '-',
            'Responsável': r.ResponsavelPadrao || '-',
            'Status': r.Status || '-',
            'Data Transmissão': d ? formatarDataCompleta(d) : '-',
            'Coordenador(a)': r['Coordenador(a)'] || '-',
        };
    });

    const ws = XLSX.utils.json_to_sheet(exportData);

    // Larguras das colunas
    ws['!cols'] = [
        { wch: 12 }, { wch: 50 }, { wch: 20 }, { wch: 20 },
        { wch: 18 }, { wch: 22 }, { wch: 18 }, { wch: 16 }, { wch: 20 }
    ];

    // Estilo do cabeçalho
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let C = range.s.c; C <= range.e.c; ++C) {
        const addr = XLSX.utils.encode_cell({ r: 0, c: C });
        if (ws[addr]) {
            ws[addr].s = {
                font: { bold: true, color: { rgb: 'FFFFFF' } },
                fill: { fgColor: { rgb: '215C98' } },
                alignment: { horizontal: 'left', vertical: 'center' }
            };
        }
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'SPEDs');
    XLSX.writeFile(wb, 'pendencias_sped.xlsx');
}

// ================================================
// INICIALIZAÇÃO
// ================================================
document.addEventListener('DOMContentLoaded', async () => {
    const dias = getDiasUteisJunho2026();
    const status = document.getElementById('statusMessage');

    try {
        const dados = await carregarArquivo();
        dadosProcessados = dados;

        const unidades = getUnidades(dados);
        unidadeAtual = unidades[0] || null;

        criarBotoesUnidade(unidades, dados, dias);
        atualizarDashboard(dados, dias, unidadeAtual);

        if (status) {
            status.innerHTML = '✅ Dados carregados com sucesso — ' + dados.length + ' registros';
            status.style.color = '#27ae60';
        }
    } catch (e) {
        if (status) {
            status.innerHTML = '❌ Erro ao carregar arquivo: ' + e.message;
            status.style.color = '#dc3545';
        }
        console.error(e);
    }

    setupModalSearch();

    document.getElementById('exportExcelBtn').addEventListener('click', exportToExcel);

    document.querySelector('.close-modal').onclick = () => {
        document.getElementById('modal').style.display = 'none';
    };

    window.onclick = e => {
        const modal = document.getElementById('modal');
        if (e.target === modal) modal.style.display = 'none';
    };
});
