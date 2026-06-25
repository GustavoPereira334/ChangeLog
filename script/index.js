// Inicializa o Google Charts
google.charts.load('current', { packages: ['corechart'] });

// Carrega o menu inicial
carregarMenuBanco();

// Helpers
function normalizar(valor) {
    return String(valor || '')
        .trim()
        .toUpperCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '');
}

function pegarCampo(obj, nomes, padrao = '-') {
    for (const nome of nomes) {
        if (obj[nome] !== undefined && obj[nome] !== null && String(obj[nome]).trim() !== '') {
            return obj[nome];
        }
    }
    return padrao;
}

// Seletor de sprint
document.getElementById('selectSprint').addEventListener('change', function () {
    const urlArquivo = this.value;
    if (!urlArquivo) return;

    fetch(urlArquivo)
        .then(res => {
            if (!res.ok) throw new Error('Arquivo não encontrado no servidor');
            return res.arrayBuffer();
        })
        .then(data => {
            const workbook = XLSX.read(data, { type: 'array' });
            processarDados(workbook);
        })
        .catch(err => {
            console.error('Erro ao carregar Excel:', err);
            alert('Verifique se o servidor está rodando.');
        });
});

function processarDados(workbook) {
    try {
        const abaDash = workbook.Sheets['Dashboard'];
        const abaResumo = workbook.Sheets['Resumo'];
        const abaTickets = workbook.Sheets['Tickets'];

        if (!abaDash || !abaResumo || !abaTickets) {
            alert('Formato do Excel inválido. Verifique se o arquivo possui as abas Dashboard, Resumo e Tickets.');
            return;
        }

        const vTotal = abaDash['A5'] ? abaDash['A5'].v : 0;
        const vAreas = abaDash['J5'] ? abaDash['J5'].v : 0;
        const vPeriodo = abaDash['B2'] ? (abaDash['B2'].w || abaDash['B2'].v) : 'N/A';

        const concluido = abaResumo['N10'] ? abaResumo['N10'].v : 0;
        const implantado = abaResumo['N11'] ? abaResumo['N11'].v : 0;
        const encerrado = abaResumo['N12'] ? abaResumo['N12'].v : 0;
        const vAtendidos = concluido + implantado + encerrado;
        const vPercentual = vTotal > 0 ? Math.round((vAtendidos / vTotal) * 100) : 0;

        document.getElementById('cardTotal').innerText = vTotal;
        document.getElementById('cardAtendidos').innerText = vAtendidos;
        document.getElementById('cardPercentual').innerText = vPercentual + '%';
        document.getElementById('cardAreas').innerText = vAreas;

        document.getElementById('resumoPeriodo').innerText = vPeriodo;
        document.getElementById('resumoTotal').innerText = vTotal;
        document.getElementById('resumoAreas').innerText = vAreas;
        document.getElementById('resumoAtendidos').innerText = vAtendidos;

        const jsonData = XLSX.utils.sheet_to_json(abaTickets);
        const jsonDataTabela = jsonData.filter(ticket => {
            const origem = normalizar(pegarCampo(ticket, ['Origem', 'origem'], 'Planejado'));
            return !origem.includes('FORA');
        });

        document.getElementById('output').innerHTML = XLSX.utils.sheet_to_html(XLSX.utils.json_to_sheet(jsonDataTabela));

        atualizaFiltrosTable();
        renderizarTabelaSLA(jsonData);

        google.charts.setOnLoadCallback(() => {
            renderizarTodosGraficos(abaResumo, jsonData);
            setTimeout(colorirNomesTabela, 120);
        });

    } catch (err) {
        console.error('Erro no processamento do arquivo:', err);
        document.getElementById('output').innerHTML = 'Erro ao carregar dados do Excel.';
    }
}

function renderizarTodosGraficos(abaResumo, jsonData) {
    const resumoJson = XLSX.utils.sheet_to_json(abaResumo, { header: 1 });

    let contagemArea = {};
    let contagemTipo = {};
    let contagemPrio = { 'Crítica': 0, 'Alta': 0, 'Média': 0, 'Baixa': 0 };
    let contagemPlanejamento = { 'Planejado': 0, 'Fora da Sprint': 0 };

    resumoJson.forEach((row, i) => {
        if (i < 9 || !row) return;

        if (row[0] !== undefined && row[1] !== undefined) {
            const valor = parseFloat(row[1]);
            const texto = String(row[0]).trim();
            if (!isNaN(valor) && texto !== '' &&
                !normalizar(texto).includes('TICKETS') &&
                !normalizar(texto).includes('DISTRIBUICAO') &&
                !normalizar(texto).includes('SEMANA')) {
                contagemArea[texto] = valor;
            }
        }

        if (row[3] !== undefined && row[4] !== undefined) {
            const valor = parseFloat(row[4]);
            const texto = String(row[3]).trim();
            if (!isNaN(valor) && texto !== '' &&
                !normalizar(texto).includes('DISTRIBUICAO') &&
                !normalizar(texto).includes('TIPO')) {
                contagemTipo[texto] = valor;
            }
        }

        if (row[6] !== undefined && row[7] !== undefined) {
            const valor = parseFloat(row[7]);
            const texto = String(row[6]).trim();
            const prioChave = ['MUITO ALTA', 'CRITICA'].includes(normalizar(texto)) ? 'Crítica' : texto;
            if (!isNaN(valor) && contagemPrio[prioChave] !== undefined && texto !== '') {
                contagemPrio[prioChave] = valor;
            }
        }
    });

    jsonData.forEach(ticket => {
        const origem = normalizar(pegarCampo(ticket, ['Origem', 'origem'], 'Planejado'));
        if (origem.includes('FORA')) {
            contagemPlanejamento['Fora da Sprint']++;
        } else {
            contagemPlanejamento['Planejado']++;
        }
    });

    let semanas = { 'Semana 1': 0, 'Semana 2': 0, 'Semana 3': 0, 'Semana 4': 0, 'Semana 5': 0 };

    jsonData.forEach(ticket => {
        const statusTicket = normalizar(ticket['Status']);
        let nomeSemana = String(ticket['Semana'] || ticket['semana'] || ticket['SEMANA'] || '').trim();

        if (nomeSemana === '1') nomeSemana = 'Semana 1';
        if (nomeSemana === '2') nomeSemana = 'Semana 2';
        if (nomeSemana === '3') nomeSemana = 'Semana 3';
        if (nomeSemana === '4') nomeSemana = 'Semana 4';
        if (nomeSemana === '5') nomeSemana = 'Semana 5';

        if (['CONCLUIDO', 'IMPLANTADO', 'ENCERRADO', 'FEITO', 'DONE', 'IMPLANTACAO'].includes(statusTicket)) {
            if (semanas[nomeSemana] !== undefined) semanas[nomeSemana]++;
        }
    });

    const dataSetor = new google.visualization.DataTable();
    dataSetor.addColumn('string', 'Área');
    dataSetor.addColumn('number', 'Quantidade');
    Object.entries(contagemArea).forEach(([k, v]) => dataSetor.addRow([k, v]));

    new google.visualization.ColumnChart(document.getElementById('chart_setores')).draw(dataSetor, {
        title: 'Tickets por Área',
        chartArea: { width: '80%', height: '70%' },
        colors: ['#034C8C'],
        legend: { position: 'none' },
        hAxis: { slantedText: true, slantedTextAngle: 45 },
        vAxis: { format: '0', viewWindow: { min: 0 } }
    });

    const dataTipo = new google.visualization.DataTable();
    dataTipo.addColumn('string', 'Tipo');
    dataTipo.addColumn('number', 'Qtd');
    Object.entries(contagemTipo).forEach(([k, v]) => dataTipo.addRow([k, v]));

    new google.visualization.PieChart(document.getElementById('chart_tipo')).draw(dataTipo, {
        title: 'Distribuição por Tipo',
        pieHole: 0.4,
        pieSliceText: 'value',
        chartArea: { width: '90%', height: '80%' },
        legend: { position: 'right' }
    });

    const dataPrio = new google.visualization.DataTable();
    dataPrio.addColumn('string', 'Prioridade');
    dataPrio.addColumn('number', 'Qtd');
    dataPrio.addRow(['Crítica', contagemPrio['Crítica'] || 0]);
    dataPrio.addRow(['Alta', contagemPrio['Alta'] || 0]);
    dataPrio.addRow(['Média', contagemPrio['Média'] || 0]);
    dataPrio.addRow(['Baixa', contagemPrio['Baixa'] || 0]);

    new google.visualization.ColumnChart(document.getElementById('chart_prioridades')).draw(dataPrio, {
        title: 'Prioridade',
        colors: ['#034C8C'],
        legend: { position: 'none' },
        vAxis: { minValue: 0, format: '0' }
    });

    const dataPlanejamento = google.visualization.arrayToDataTable([
        ['Origem', 'Quantidade'],
        ['Planejado', contagemPlanejamento['Planejado'] || 0],
        ['Fora da Sprint', contagemPlanejamento['Fora da Sprint'] || 0]
    ]);

    new google.visualization.PieChart(document.getElementById('chart_sla')).draw(dataPlanejamento, {
        title: 'Planejado x Fora da Sprint',
        colors: ['#034C8C', '#fd7e14'],
        chartArea: { width: '90%', height: '80%' },
        legend: { position: 'right' }
    });

    const dataBurn = new google.visualization.DataTable();
    dataBurn.addColumn('string', 'Semana');
    dataBurn.addColumn('number', 'Quantidade');

    ['Semana 1', 'Semana 2', 'Semana 3', 'Semana 4', 'Semana 5'].forEach(sem => {
        dataBurn.addRow([sem, semanas[sem]]);
    });

    new google.visualization.LineChart(document.getElementById('chart_burn_down')).draw(dataBurn, {
        title: 'Tickets Finalizados por Semana',
        legend: { position: 'none' },
        colors: ['#034C8C'],
        pointSize: 10,
        chartArea: { width: '85%', height: '75%' },
        vAxis: { minValue: 0, format: '0' }
    });
}

function colorirNomesTabela() {
    const container = document.getElementById('output');
    if (!container) return;

    const cores = {
        'ALTA': { color: '#fd7e14', bg: '#fff3e0' },
        'BAIXA': { color: 'green', bg: '#e8f5e9' },
        'MÉDIA': { color: '#007bff', bg: '#e3f2fd' },
        'CRÍTICA': { color: '#FF0A01', bg: '#ffebee' },
        'MUITO ALTA': { color: '#FF0A01', bg: '#ffebee' },
        'AUTOMAÇÃO': { color: '#fd7e14', bg: '#fff3e0' },
        'BUG': { color: 'blue', bg: '#e8eaf6' },
        'MELHORIA': { color: '#ffd000', bg: '#ffffeb' },
        'CORREÇÃO': { color: 'green', bg: '#e8f5e9' },
        'PROJETO': { color: '#00b7ff', bg: '#e6e4fc' },
        'NÃO RETORNADO': { color: '#6c757d', bg: '#f8f9fa' },
        'FEITO': { color: '#28a745', bg: '#e8f5e9' },
        'CONCLUÍDO': { color: '#28a745', bg: '#e8f5e9' },
        'IMPLANTADO': { color: '#034C8C', bg: '#e3f2fd' },
        'ENCERRADO': { color: '#6c757d', bg: '#f8f9fa' },
        'A FAZER': { color: '#fd7e14', bg: '#fff3e0' },
        'FAZENDO': { color: '#007bff', bg: '#e3f2fd' },
        'IMPEDIMENTO': { color: '#FF0A01', bg: '#ffebee' },
        'IMPLANTAÇÃO': { color: '#034C8C', bg: '#e3f2fd' },
        'HOMOLOGAÇÃO': { color: '#9c27b0', bg: '#f3e5f5' },
        'NOVO RECURSO': { color: '#964B00', bg: '#F5EBE0' },
        'NOVO SISTEMA': { color: '#964B00', bg: '#F5EBE0' },
        'ATUALIZAÇÃO DE VERSÃO': { color: '#6a1b9a', bg: '#f3e5f5' },
        'ANÁLISE': { color: 'white', bg: '#1C1C1C' },
        'PLANEJADO': { color: '#034C8C', bg: '#e3f2fd' },
        'FORA DA SPRINT': { color: '#fd7e14', bg: '#fff3e0' },
    };

    container.querySelectorAll('td').forEach(cell => {
        const texto = cell.innerText.trim().toUpperCase();
        if (cores[texto]) {
            cell.innerHTML = `<span class="badge-status" style="color:${cores[texto].color};background-color:${cores[texto].bg};">${texto}</span>`;
            cell.style.textAlign = 'center';
        }
    });
}

function renderizarTabelaSLA(jsonData) {
    const container = document.getElementById('outputSLA');

    const foraDaSprint = jsonData.filter(ticket => {
        const origem = normalizar(pegarCampo(ticket, ['Origem', 'origem'], ''));
        return origem.includes('FORA');
    });

    if (foraDaSprint.length === 0) {
        container.innerHTML = '<p style="color:#28a745;font-size:0.85rem;padding:10px 0;">Nenhuma demanda atendida fora da sprint.</p>';
        return;
    }

    const vistos = new Set();

    const linhas = foraDaSprint
        .filter(ticket => {
            const id = String(pegarCampo(ticket, ['Ticket', 'ticket'], '-')).trim();
            if (vistos.has(id)) return false;
            vistos.add(id);
            return true;
        })
        .map(ticket => {
            const id = pegarCampo(ticket, ['Ticket', 'ticket'], '-');
            const titulo = pegarCampo(ticket, ['Título', 'Titulo', 'T?tulo', 'T??tulo', 'titulo'], '-');
            const area = pegarCampo(ticket, ['Área', 'Area', '?rea', '??rea', 'area'], '-');

            return `
                <tr>
                    <td title="${id}">${id}</td>
                    <td title="${titulo}">${titulo}</td>
                    <td title="${area}">${area}</td>
                </tr>`;
        }).join('');

    container.innerHTML = `
        <table>
            <colgroup>
                <col class="col-id">
                <col>
                <col class="col-area">
            </colgroup>
            <thead>
                <tr>
                    <th>ID</th>
                    <th>Título</th>
                    <th>Área</th>
                </tr>
            </thead>
            <tbody>${linhas}</tbody>
        </table>`;
}

function filtrarTabela() {
    const busca = document.getElementById('busca').value.toUpperCase();
    const areaSelecionada = document.getElementById('areaBox').value.toUpperCase();
    const tipoSelecionado = document.getElementById('tipoBox').value.toUpperCase();

    const rows = document.getElementById('output').querySelectorAll('table tr');

    rows.forEach((row, index) => {
        if (index === 0 && row.querySelector('th')) return;

        const celulas = row.querySelectorAll('td');
        if (celulas.length === 0) return;

        const textoLinha = row.innerText.toUpperCase();
        const textoArea = celulas[1] ? celulas[1].innerText.toUpperCase() : '';
        const textoTipo = celulas[2] ? celulas[2].innerText.toUpperCase() : '';

        const bateBusca = textoLinha.includes(busca);
        const bateArea = areaSelecionada === 'TODASAREAS' || textoArea.includes(areaSelecionada);
        const bateTipo = tipoSelecionado === 'TODOSTIPOS' || textoTipo.includes(tipoSelecionado);

        row.style.display = bateBusca && bateArea && bateTipo ? '' : 'none';
    });
}

function atualizaFiltrosTable() {
    const areaBox = document.getElementById('areaBox');
    const rows = document.getElementById('output').querySelectorAll('table tr');
    const setoresUnicos = new Set();

    rows.forEach((row, index) => {
        if (index === 0) return;

        const celulas = row.querySelectorAll('td');
        if (celulas[1]) {
            const setor = celulas[1].innerText.trim();
            if (setor && setor !== 'Área') setoresUnicos.add(setor);
        }
    });

    areaBox.innerHTML = '<option value="todasAreas">Todas as Áreas</option>';

    setoresUnicos.forEach(setor => {
        const opt = document.createElement('option');
        opt.value = setor.toUpperCase();
        opt.textContent = setor;
        areaBox.appendChild(opt);
    });
}

document.getElementById('busca').addEventListener('input', filtrarTabela);
document.getElementById('areaBox').addEventListener('change', filtrarTabela);
document.getElementById('tipoBox').addEventListener('change', filtrarTabela);

function carregarMenuBanco() {
    const select = document.getElementById('selectSprint');
    if (!select) return;

    fetch(`${CONFIG.SERVER_BASE_URL}/api/sprints-da-pasta`)
        .then(res => {
            if (!res.ok) throw new Error('Servidor offline');
            return res.json();
        })
        .then(listaSprints => {
            if (!listaSprints || listaSprints.length === 0) throw new Error('Pasta vazia');
            popularSelectComMapeamento(listaSprints, true);
        })
        .catch(() => {
            console.warn('Servidor local offline. Carregando índice estático da pasta utils...');
            fetch('./utils/sprints.json')
                .then(resOffline => {
                    if (!resOffline.ok) throw new Error('Índice sprints.json não localizado');
                    return resOffline.json();
                })
                .then(listaEstatica => {
                    popularSelectComMapeamento(listaEstatica, false);
                })
                .catch(errOffline => {
                    console.error('Falha ao ler índice estático:', errOffline);
                    select.innerHTML = '<option value="">Nenhum arquivo Excel registrado na utils.</option>';
                });
        });
}

function popularSelectComMapeamento(lista, servidorOnline) {
    const select = document.getElementById('selectSprint');
    select.innerHTML = '<option value="">Selecione um arquivo Excel...</option>';

    lista.forEach(sprint => {
        const opt = document.createElement('option');
        opt.value = servidorOnline
            ? `${CONFIG.SERVER_BASE_URL}/${sprint.caminho_arquivo}`
            : `./${sprint.caminho_arquivo}`;

        opt.textContent = sprint.nome_exibicao;
        select.appendChild(opt);
    });

    if (lista.length > 0) {
        select.selectedIndex = 1;
        select.dispatchEvent(new Event('change'));
    }
}