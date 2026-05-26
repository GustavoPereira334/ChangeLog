// Inicializa o Google Charts
google.charts.load('current', { 'packages': ['corechart'] });

// Carrega o menu inicial ao inicializar o arquivo
carregarMenuBanco();

//LISTENERS DOS BOTÕES DE SINCRONIZAÇÃO
document.getElementById('btnSyncMovidesk').addEventListener('click', async function () {
    const btn = this;
    const original = btn.innerHTML;
    btn.disabled = true;
    btn.innerHTML = 'Sincronizando Movidesk...';
    btn.style.opacity = '0.6';

    try {
        const response = await fetch('http://localhost:3008/api/sincronizar-movidesk?sprint=Geral');
        const data = await response.json();
        if (data.success) {
            alert('Movidesk sincronizado com sucesso!');
            carregarMenuBanco();
        } else {
            alert('Erro: ' + (data.erro || 'Erro desconhecido'));
        }
    } catch (err) {
        alert('Servidor não está rodando.');
    } finally {
        btn.disabled = false;
        btn.innerHTML = original;
        btn.style.opacity = '1';
    }
});

document.getElementById('btnSyncMonday').addEventListener('click', async function () {
    const btn = this;
    const original = btn.innerHTML;
    btn.disabled = true;
    btn.innerHTML = 'Sincronizando Monday...';
    btn.style.opacity = '0.6';

    try {
        const response = await fetch('http://localhost:3008/api/sincronizar-monday-sprints');
        const data = await response.json();
        if (data.success) {
            alert('Monday sincronizado. Arquivos gerados: ' + (data.files || []).join(', '));
            carregarMenuBanco();
        } else {
            alert('Erro: ' + (data.erro || 'Erro desconhecido'));
        }
    } catch (err) {
        alert('Servidor não está rodando.');
    } finally {
        btn.disabled = false;
        btn.innerHTML = original;
        btn.style.opacity = '1';
    }
});

//  SELETOR DE SPRINT E CARREGAMENTO EXCEL
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
            alert('Erro ao abrir o arquivo Excel. Verifique se o servidor está rodando.');
        });
});

//  PROCESSAMENTO DOS DADOS DO EXCEL
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
        const vSLA = abaDash['D5'] ? abaDash['D5'].v : 0;
        const vTempo = abaDash['G5'] ? abaDash['G5'].v : 0;
        const vAreas = abaDash['J5'] ? abaDash['J5'].v : 0;
        const vPeriodo = abaDash['B2'] ? (abaDash['B2'].w || abaDash['B2'].v) : 'N/A';

        // Calcula atendidos Concluído + Implantado + Encerrado utilizando a aba Resumo
        const concluido = abaResumo['N10'] ? abaResumo['N10'].v : 0;
        const implantado = abaResumo['N11'] ? abaResumo['N11'].v : 0;
        const encerrado = abaResumo['N12'] ? abaResumo['N12'].v : 0;
        const vAtendidos = concluido + implantado + encerrado;
        const vPercentual = vTotal > 0 ? Math.round((vAtendidos / vTotal) * 100) : 0;

        // Renderiza valores nos blocos de cards superiores
        document.getElementById('cardTotal').innerText = vTotal;
        document.getElementById('cardAtendidos').innerText = vAtendidos;
        document.getElementById('cardPercentual').innerText = vPercentual + '%';
        document.getElementById('cardAreas').innerText = vAreas;

        // Renderiza valores nas estatísticas do header
        document.getElementById('resumoPeriodo').innerText = vPeriodo;
        document.getElementById('resumoTotal').innerText = vTotal;
        document.getElementById('resumoAreas').innerText = vAreas;
        document.getElementById('resumoAtendidos').innerText = vAtendidos;

        // Renderização da tabela principal de chamados
        const jsonData = XLSX.utils.sheet_to_json(abaTickets);
        document.getElementById('output').innerHTML = XLSX.utils.sheet_to_html(XLSX.utils.json_to_sheet(jsonData));

        // Inicializa o preenchimento automático das caixas de seleção de filtros
        atualizaFiltrosTable();

        // Renderização da Tabela de Chamados fora da Sprint
        const rangeOriginal = abaDash['!ref'];
        abaDash['!ref'] = 'J9:L20'; // Filtra apenas a região crítica do dashboard do Excel
        document.getElementById('outputSLA').innerHTML = XLSX.utils.sheet_to_html(abaDash);
        abaDash['!ref'] = rangeOriginal;

        //renderiza Google Charts
        google.charts.setOnLoadCallback(() => {
            renderizarTodosGraficos(abaResumo, jsonData);
            setTimeout(colorirNomesTabela, 120);
        });

    } catch (err) {
        console.error('Erro no processamento do arquivo:', err);
        document.getElementById('output').innerHTML = 'Erro ao carregar dados do Excel.';
    }
}

// MOTOR DE COMPOSIÇÃO DOS GRÁFICOS
function renderizarTodosGraficos(abaResumo, jsonData) {
    const resumoJson = XLSX.utils.sheet_to_json(abaResumo, { header: 1 });

    let contagemArea = {};
    let contagemTipo = {};
    let contagemPrio = { 'Crítica': 0, 'Alta': 0, 'Média': 0, 'Baixa': 0 };
    let contagemSLA = { 'Sim': 0, 'Não': 0 };

    // EXTRAÇÃO DE KPI'S MATRICIAIS DA ABA RESUMO
    resumoJson.forEach((row, i) => {
        if (i < 9) return;

        // Tickets por Área
        if (row && row[0] && row[1] !== undefined) {
            const valor = parseFloat(row[1]);
            const texto = String(row[0]).trim();
            if (!isNaN(valor) && !texto.toUpperCase().includes('TICKETS') && !texto.toUpperCase().includes('DISTRIBUIÇÃO') && !texto.toUpperCase().includes('SEMANA')) {
                contagemArea[texto] = valor;
            }
        }

        // Distribuição por Tipo 
        if (row && row[3] && row[4] !== undefined) {
            const valor = parseFloat(row[4]);
            const texto = String(row[3]).trim();
            if (!isNaN(valor) && !texto.toUpperCase().includes('DISTRIBUIÇÃO') && !texto.toUpperCase().includes('TIPO')) {
                contagemTipo[texto] = valor;
            }
        }

        // Tickets por Prioridade 
        if (row && row[6] && row[7] !== undefined) {
            const valor = parseFloat(row[7]);
            const texto = String(row[6]).trim();
            const prioChave = (texto === 'Muito Alta' || texto === 'CRÍTICA') ? 'Crítica' : texto;
            if (!isNaN(valor) && contagemPrio[prioChave] !== undefined) {
                contagemPrio[prioChave] = valor;
            }
        }

        //Cumprimento de SLA (Colunas J e K)
        if (row && row[9] && row[10] !== undefined) {
            const valor = parseFloat(row[10]);
            const texto = String(row[9]).trim();
            if (!isNaN(valor)) {
                if (texto === 'Sim') contagemSLA['Sim'] = valor;
                if (texto === 'Não') contagemSLA['Não'] = valor;
            }
        }
    });

    let semanas = { 'Semana 1': 0, 'Semana 2': 0, 'Semana 3': 0, 'Semana 4': 0, 'Semana 5': 0 };

    jsonData.forEach(ticket => {
        // Captura o status e a string da semana diretamente da aba Tickets do Excel
        const statusTicket = ticket['Status'] ? String(ticket['Status']).trim().toUpperCase() : '';

        // Mapeia os possíveis nomes que a coluna 'Semana' pode assumir na leitura do JSON
        const semanaCrua = ticket['Semana'] || ticket['semana'] || ticket['SEMANA'] || '';
        let nomeSemana = String(semanaCrua).trim();

        // Fallback Inteligente: Se o Excel trouxe apenas o número isolado (ex: 1, 2, 3) ao invés do texto completo
        if (nomeSemana === '1') nomeSemana = 'Semana 1';
        if (nomeSemana === '2') nomeSemana = 'Semana 2';
        if (nomeSemana === '3') nomeSemana = 'Semana 3';
        if (nomeSemana === '4') nomeSemana = 'Semana 4';
        if (nomeSemana === '5') nomeSemana = 'Semana 5';

        // Incrementa apenas se pertencer ao escopo de chamados finalizados
        if (['CONCLUÍDO', 'IMPLANTADO', 'ENCERRADO', 'FEITO', 'DONE'].includes(statusTicket)) {
            if (semanas[nomeSemana] !== undefined) {
                semanas[nomeSemana]++;
            }
        }
    });


    // Grafico Tickets por Área
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

    // GRÁFICO Distribuição por Tipo
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

    // GRÁFICO  Tickets por Prioridade
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

    // GRÁFICO Cumprimento de SLA
    const dataSLA = google.visualization.arrayToDataTable([
        ['SLA', 'Valor'],
        ['Sim', contagemSLA['Sim'] || 0],
        ['Não', contagemSLA['Não'] || 0]
    ]);

    new google.visualization.PieChart(document.getElementById('chart_sla')).draw(dataSLA, {
        title: 'Cumprimento de SLA',
        colors: ['#034C8C', '#fd7e14'],
        chartArea: { width: '90%', height: '80%' },
        legend: { position: 'right' }
    });

    // GRÁFICO Tickets por Semana
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

// Altera cores e fundo de certas "Escritas" nas tabelas
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
        'Correção': { color: 'green', bg: '#e8f5e9' },
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
        'CORREÇÃO': { color: '#800020', bg: '#F3E1EC' },
        'ANÁLISE': { color: 'White', bg: '#1C1C1C' },
    };

    container.querySelectorAll('td').forEach(cell => {
        const texto = cell.innerText.trim().toUpperCase();
        if (cores[texto]) {
            cell.innerHTML = `<span class="badge-status" style="color:${cores[texto].color};background-color:${cores[texto].bg};">${texto}</span>`;
            cell.style.textAlign = 'center';
        }
    });
}

// 6.  FILTROS DA TABELA
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

// Vincula ouvintes de digitação e alteração aos elementos de filtro
document.getElementById('busca').addEventListener('input', filtrarTabela);
document.getElementById('areaBox').addEventListener('change', filtrarTabela);
document.getElementById('tipoBox').addEventListener('change', filtrarTabela);

// 7. CARREGAMENTO DO SERVIDOR

function carregarMenuBanco() {
    const select = document.getElementById('selectSprint');
    if (!select) return;
    fetch('http://localhost:3008/api/sprints-da-pasta')
        .then(res => {
            if (!res.ok) throw new Error('Servidor offline');
            return res.json();
        })
        .then(listaSprints => {
            if (!listaSprints || listaSprints.length === 0) throw new Error('Pasta vazia');
            popularSelectComMapeamento(listaSprints, true);
        })
        .catch(err => {
            console.warn('Servidor local offline. Carregando índice estático da pasta utils...');
            fetch('./utils/sprints.json')
                .then(resOffline => {
                    if (!resOffline.ok) throw new Error('Índice sprints.json não localizado');
                    return resOffline.json();
                })
                .then(listaEstática => {
                    popularSelectComMapeamento(listaEstática, false);
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
        if (servidorOnline) {
            opt.value = sprint.caminho_arquivo.startsWith('http') ? sprint.caminho_arquivo : `http://localhost:3008/${sprint.caminho_arquivo}`;
        } else {
            opt.value = `./${sprint.caminho_arquivo}`;
        }

        opt.textContent = sprint.nome_exibicao;
        select.appendChild(opt);
    });

    // Dispara automaticamente a leitura do primeiro Excel da lista
    if (lista.length > 0) {
        select.selectedIndex = 1;
        select.dispatchEvent(new Event('change'));
    }
}