const axios = require('axios');
const ExcelJS = require('exceljs');
const path = require('path');
const https = require('https');
const fs = require('fs').promises;

// ==========================================
// 1. CONFIGURAÇÕES E MAPAS DE DE-PARA
// ==========================================
const mondayStatusMap = {
  "Concluído": "Concluído", "Done": "Concluído", "Feito": "Concluído",
  "Implantado": "Implantado", "Implantação": "Implantação",
  "Homologação": "Homologação", "Em andamento": "Em andamento",
  "Working on it": "Em andamento", "A fazer": "A fazer",
  "To Do": "A fazer", "Impedimento": "Impedimento",
  "Encerrado": "Encerrado", "Cancelado": "Cancelado",
  "": "Não definido"
};

const mondayPriorityMap = {
  "Alta": "Alta", "High": "Alta",
  "Média": "Média", "Medium": "Média",
  "Baixa": "Baixa", "Low": "Baixa",
  "Muito Alta": "Muito Alta", "Critical": "Muito Alta",
  "": "Não definida"
};

const urgencyMap = {
  53857: "Muito Alta", // Assumindo que 53860 também é "Muito Alta" e 53857 é "Alta"
  53859: "Baixa",
  53858: "Média",
  53860: "Muito Alta"
};

// CUIDADO: Desabilita a verificação de certificado TLS para requisições HTTPS.
// Mantido aqui conforme sua necessidade de ambiente, mas reafirmo que esta é
// uma vulnerabilidade de segurança grave e deve ser usada com extrema cautela,
// preferencialmente APENAS em ambientes de desenvolvimento isolados e NUNCA em produção.
const httpsAgent = new https.Agent({ rejectUnauthorized: false });

// ==========================================
// 2. FUNÇÕES AUXILIARES DE DATA
// ==========================================
function calcularSemana(dateObj) {
  if (!dateObj || isNaN(dateObj.getTime())) return 'Semana 1';
  const dia = dateObj.getDate();
  if (dia <= 7) return 'Semana 1';
  if (dia <= 14) return 'Semana 2';
  if (dia <= 21) return 'Semana 3';
  if (dia <= 28) return 'Semana 4';
  return 'Semana 5';
}

function nomeMes(dateObj) {
  if (!dateObj || isNaN(dateObj.getTime())) return 'Indefinido';
  const meses = [
    'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
    'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
  ];
  return meses[dateObj.getMonth()];
}

// =========================================================================
// 3.1. FUNÇÃO: EXTRAÇÃO DE VALORES DE CAMPOS PERSONALIZADOS DE PESSOAS (CLIENTES)
//      (Implementação COMPLETA para buscar o SETOR do cliente)
// =========================================================================
async function buscarValoresCamposPersonalizadosPessoas(token) {
  const mapaValoresCamposPessoas = {}; // Mapa para armazenar PersonId -> { customFieldId -> valor }
  let skip = 0;
  const take = 100;
  const maxRetries = 3;
  const retryDelayMs = 2000;

  console.log('[Movidesk] Iniciando varredura na API de Pessoas para mapear VALORES de campos personalizados...');

  while (true) {
    let resp;
    let retries = 0;
    let success = false;
    let falhaCritica = false;

    while (retries < maxRetries && !success) {
      try {
        resp = await axios.get('https://api.movidesk.com/public/v1/persons', { // <<< ENDPOINT CORRETO PARA PESSOAS
          params: {
            token,
            '$select': 'id', // Precisamos do ID da pessoa
            '$expand': 'customFieldValues($expand=items)', // Expande os VALORES dos campos personalizados da pessoa
            '$top': take,
            '$skip': skip,
          },
          httpsAgent,
          timeout: 90000,
        });
        success = true;
      } catch (err) {
        retries++;
        console.error(`[Movidesk] Erro ao buscar valores de pessoas skip=${skip} (tentativa ${retries}/${maxRetries}): ${err.message}`);
        if (retries < maxRetries) {
          await new Promise(resolve => setTimeout(resolve, retryDelayMs));
        } else {
          console.error(`[Movidesk] Interrompendo paginação de valores de pessoas em skip=${skip}.`);
          falhaCritica = true;
        }
      }
    }

    if (falhaCritica || !success) break;

    const pessoas = resp.data.value || (Array.isArray(resp.data) ? resp.data : []);
    if (pessoas.length === 0) break;

    for (const pessoa of pessoas) {
      const personId = String(pessoa.id);
      mapaValoresCamposPessoas[personId] = {};

      if (!pessoa.customFieldValues || pessoa.customFieldValues.length === 0) continue;

      for (const campo of pessoa.customFieldValues) {
        const idCampo = String(campo.customFieldId);
        if (campo.value !== undefined && campo.value !== null) {
            mapaValoresCamposPessoas[personId][idCampo] = campo.value;
        } else if (campo.items && campo.items.length > 0) {
          mapaValoresCamposPessoas[personId][idCampo] = campo.items[0].customFieldItem;
        }
      }
    }

    console.log(`[Movidesk] Mapeando valores de campos para ${pessoas.length} pessoas (skip=${skip})...`);
    skip += take;
  }

  console.log(`[Movidesk] Mapeamento de valores de campos de pessoas concluído. ${Object.keys(mapaValoresCamposPessoas).length} pessoas com campos personalizados identificadas.`);
  return mapaValoresCamposPessoas;
}

// =========================================================================
// 3.2. FUNÇÃO: EXTRAÇÃO DA ESTRUTURA DOS CAMPOS PERSONALIZADOS DE CLIENTES
//      (Corrigida para buscar a ESTRUTURA dos campos de pessoas, não tickets)
// =========================================================================
async function buscarEstruturaCamposPersonalizadosClientes(token) {
  const mapaCamposGlobais = {};
  let skip = 0;
  const take = 100;
  const maxRetries = 3;
  const retryDelayMs = 2000;

  console.log('[Movidesk] Iniciando varredura na API de CustomFields para mapear a ESTRUTURA de campos personalizados de clientes...');

  while (true) {
    let resp;
    let retries = 0;
    let success = false;
    let falhaCritica = false;

    while (retries < maxRetries && !success) {
      try {
        resp = await axios.get('https://api.movidesk.com/public/v1/customFields', { // <<< ENDPOINT CORRETO PARA ESTRUTURA DE CAMPOS
          params: {
            token,
            '$filter': `entityType eq 'Person'`, // Filtra por campos de Pessoa/Cliente
            '$top': take,
            '$skip': skip,
          },
          httpsAgent,
          timeout: 90000,
        });
        success = true;
      } catch (err) {
        retries++;
        console.error(`[Movidesk] Erro ao buscar estrutura de campos (tentativa ${retries}/${maxRetries}): ${err.message}`);
        if (retries < maxRetries) {
          await new Promise(resolve => setTimeout(resolve, retryDelayMs));
        } else {
          console.error(`[Movidesk] Interrompendo paginação de estrutura de campos em skip=${skip}.`);
          falhaCritica = true;
        }
      }
    }

    if (falhaCritica || !success) break;

    const campos = resp.data.value || (Array.isArray(resp.data) ? resp.data : []);
    if (campos.length === 0) break;

    for (const campo of campos) {
      const idCampo = String(campo.id); // O ID do campo personalizado é 'id' na resposta de customFields
      
      if (!mapaCamposGlobais[idCampo]) {
        mapaCamposGlobais[idCampo] = {
          customFieldId: campo.id,
          customFieldRuleId: campo.customFieldRuleId,
          type: campo.type || 'N/A',
          itensDisponiveis: {}
        };
      }

      if (campo.items && campo.items.length > 0) {
        for (const item of campo.items) {
          const idItem = String(item.customFieldItemId);
          mapaCamposGlobais[idCampo].itensDisponiveis[idItem] = item.customFieldItem;
        }
      }
    }

    console.log(`[Movidesk] Mapeando estrutura: ${campos.length} campos analisados (skip=${skip})...`);
    skip += take;
  }

  const relatorioFinal = Object.values(mapaCamposGlobais).map(campo => ({
    customFieldId: campo.customFieldId,
    customFieldRuleId: campo.customFieldRuleId,
    type: campo.type,
    items: Object.entries(campo.itensDisponiveis).map(([id, nome]) => ({
      customFieldItemId: Number(id),
      customFieldItem: nome
    }))
  }));

  console.log(`[Movidesk] Mapeamento concluído. ${relatorioFinal.length} estruturas de campos personalizados de clientes identificadas.`);
  return relatorioFinal;
}


// ==========================================
// 4. INTEGRAÇÃO DA API MOVIDESK TICKETS COM RETRY (CORRIGIDA)
// ==========================================
async function buscarTodosTicketsMovidesk(token, personCustomFieldValuesMap, movideskClientSectorCustomFieldId) {
  const mapa = {};
  let skip = 0;
  const take = 100;
  let total = 0;
  const maxRetries = 3;
  const retryDelayMs = 2000;

  console.log('[Movidesk] Iniciando busca paginada de tickets com política de Resiliência...');

  while (true) {
    let resp;
    let retries = 0;
    let success = false;
    let falhaCritica = false;

    while (retries < maxRetries && !success) {
      try {
        resp = await axios.get('https://api.movidesk.com/public/v1/tickets', { // <<< ENDPOINT E PASSAGEM DE TOKEN CORRETOS
          params: {
            token: token,
            '$select': 'id,subject,category,urgency,status,createdDate,resolvedIn,ownerTeam,serviceFirstLevel,slaSolutionDate,createdBy,clients',
            '$expand': 'actions($select=description),createdBy',
            '$filter': 'createdDate gt 2023-01-01T00:00:00Z',
            '$top': take,
            '$skip': skip,
          },
          httpsAgent,
          timeout: 90000,
        });
        success = true;
      } catch (err) {
        retries++;
        console.error(`[Movidesk] Erro na página skip=${skip} (tentativa ${retries}/${maxRetries}): ${err.message}`);
        if (retries < maxRetries) {
          console.log(`[Movidesk] Tentando novamente em ${retryDelayMs / 1000} segundos...`);
          await new Promise(resolve => setTimeout(resolve, retryDelayMs));
        } else {
          console.error(`[Movidesk] Todas as tentativas falharam para skip=${skip}. Retornando dados parciais coletados.`);
          falhaCritica = true;
        }
      }
    }

    if (falhaCritica || !success) break;
    const lista = resp.data && resp.data.value ? resp.data.value : (Array.isArray(resp.data) ? resp.data : []);
    if (lista.length === 0) break;

    for (const t of lista) {
      const idTicket = t.id || t.Id || 'N/A';
      const assunto = t.subject || t.Subject || 'N/A';
      const statusTicket = t.status || t.Status || 'N/A';
      const equipeBruta = t.ownerTeam || t.OwnerTeam || 'N/A';
      const criadorObjeto = t.createdBy || t.CreatedBy || {};

      const equipe = String(equipeBruta).trim();
      if (equipe === 'Suporte e Infraestrutura' || equipe === 'Administradores') continue;

      const prioridadeMovidesk = urgencyMap[t.urgency || t.Urgency] || 'N/A';
      const categoria = t.category || t.Category || 'N/A';
      const tipo = t.serviceFirstLevel || t.ServiceFirstLevel ? String(t.serviceFirstLevel || t.ServiceFirstLevel).trim() : 'N/A';

      const criacaoRaw = t.createdDate || t.CreatedDate || null;
      const resolucaoRaw = t.resolvedIn || t.ResolvedIn || null;
      const createdDate = criacaoRaw ? new Date(criacaoRaw) : null;
      const closedDate = resolucaoRaw ? new Date(resolucaoRaw) : null;

      const acoes = t.actions || t.Actions || [];
      const descRaw = acoes[0]?.description || acoes[0]?.Description || acoes[0]?.body || acoes[0]?.Body || '';
      let descTratada = String(descRaw)
        .replace(/<[^>]*>?/gm, '')
        .replace(/\[cid:[^\]]*\]/g, '')
        .replace(/[\r\n\t]+/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
      const descricao = descTratada.substring(0, 300) || 'N/A';

      const solicitante = criadorObjeto.personName ||
        criadorObjeto.PersonName ||
        criadorObjeto.name ||
        criadorObjeto.Name ||
        criadorObjeto.businessName ||
        criadorObjeto.BusinessName ||
        'N/A';

      let setorCliente = 'N/A';
      const createdById = criadorObjeto.id;

      if (createdById && personCustomFieldValuesMap && movideskClientSectorCustomFieldId && personCustomFieldValuesMap[createdById]) {
        setorCliente = personCustomFieldValuesMap[createdById][movideskClientSectorCustomFieldId] || 'Não informado';
      }

      if (idTicket === 'N/A' || solicitante === 'N/A' || equipe === 'N/A') {
        console.log(`\n============== [AUDITORIA CRÍTICA TICKET INTERNO] ==============`);
        console.log(`ID Detectado: ${idTicket}`);
        console.log(`Assunto: ${assunto}`);
        console.log(`Status do Ticket: ${statusTicket}`);
        console.log(`Estrutura Bruta Recebida da API (Primeiras chaves):`, Object.keys(t).slice(0, 8));
        console.log(`=================================================================\n`);
      }

      mapa[String(idTicket)] = {
        area: setorCliente,
        tipo,
        categoria,
        descricao,
        prioridade: prioridadeMovidesk,
        solicitante: String(solicitante).trim(),
        equipe: equipe,
        abertura: createdDate ? createdDate.toISOString() : null,
        encerramento: closedDate ? closedDate.toISOString() : null,
        status: statusTicket
      };

      total++;
    }

    console.log(`[Movidesk] ${total} tickets legítimos processados na memória (skip=${skip})...`);
    skip += take;

    // Trava de salvaguarda para a rota de chamados não estourar em homologação
    if (total > 5000) break;
  }

  console.log(`[Movidesk] Total final: ${total} tickets carregados.`);
  return mapa;
}

// ==========================================
// 5. INTEGRAÇÃO DA API MONDAY (CORRIGIDA)
// ==========================================
async function buscarTodosItensMonday(boardId, token) {
  let allItems = [];
  let cursor = null;
  let hasNextPage = true;

  console.log(`[Monday] Buscando itens do board ${boardId}...`);

  while (hasNextPage) {
    const query = `
      query {
        boards(ids: ${boardId}) {
          items_page(limit: 500, cursor: ${cursor ? `"${cursor}"` : "null"}) {
            cursor
            items {
              id
              name
              group { title }
              column_values {
                id
                text
              }
            }
          }
        }
      }
    `;

    const response = await axios.post('https://api.monday.com/v2', // <<< ENDPOINT CORRETO
      { query },
      { headers: { 'Authorization': token, 'Content-Type': 'application/json' } }
    );

    if (response.data.errors) {
      throw new Error(`Erro GraphQL Monday: ${JSON.stringify(response.data.errors)}`);
    }

    const boardData = response.data.data.boards[0];
    if (!boardData) break;

    const data = boardData.items_page;
    allItems = allItems.concat(data.items);
    cursor = data.cursor;
    hasNextPage = !!cursor;
  }
  console.log(`[Monday] Total de ${allItems.length} itens encontrados no board ${boardId}.`);
  return allItems;
}

const getColVal = (item, id) => {
  const c = item.column_values?.find(cv => cv.id === id);
  return c?.text?.trim() || '';
};

// ==========================================
// 6. MOTOR DE PROCESSAMENTO DO EXCEL (CORRIGIDO)
// ==========================================
async function gerarExcelParaSprint(
  sprintNome, itensDaSprint, mapaMovidesk, dirPath,
  mondayTextColId, mondayStatusColId, mondayPriorityColId, mondayDescriptionColId, mondayTypeColId,
  mondayCreationDateColId, mondayResolutionDateColId, // <<< NOVOS PARÂMETROS
  listaEstruturaCamposClientes // <<< RENOMEADO PARA CLAREZA E CONSISTÊNCIA
) {
  const workbook = new ExcelJS.Workbook();
  const estiloAzul = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF034C8C' } };
  const fonteBranca = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
  const fonteNegrito = { bold: true };

  let dentroSLACount = 0;
  let foraSLACount = 0;
  let totalTempoResolucao = 0;
  let countResolvidos = 0;
  const contagemArea = {};
  const contagemTipo = {};
  const contagemPrio = {};
  const contagemStatus = {};
  const contagemSemanaBurnDown = {};
  const contagemSemanaFinalizadosMovidesk = {};
  const ticketsForaSLA = [];

  const ticketsSheet = workbook.addWorksheet('Tickets');
  ticketsSheet.columns = [
    { header: 'Ticket', key: 'ticket', width: 10 },
    { header: 'Área', key: 'area', width: 20 },
    { header: 'Tipo', key: 'tipo', width: 18 },
    { header: 'Categoria', key: 'categoria', width: 18 },
    { header: 'Título', key: 'titulo', width: 40 },
    { header: 'Descrição', key: 'descricao', width: 40 },
    { header: 'Prioridade', key: 'prioridade', width: 12 },
    { header: 'Solicitante', key: 'solicitante', width: 25 },
    { header: 'Equipe TI', key: 'equipe', width: 20 },
    { header: 'Abertura', key: 'abertura', width: 14 },
    { header: 'Encerramento', key: 'encerramento', width: 14 },
    { header: 'Status', key: 'status', width: 16 },
    { header: 'Semana', key: 'semana', width: 12 },
    { header: 'Mês', key: 'mes', width: 14 },
  ];

  ticketsSheet.getRow(1).eachCell(cell => {
    cell.fill = estiloAzul;
    cell.font = fonteBranca;
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  });
  ticketsSheet.getRow(1).height = 24;

  for (const item of itensDaSprint) {
    let mondayMovideskId = getColVal(item, mondayTextColId);
    const mov = mondayMovideskId ? (mapaMovidesk[mondayMovideskId] || {}) : {};

    const prioMon = mondayPriorityMap[getColVal(item, mondayPriorityColId)] || 'Não definida';
    const typeMon = getColVal(item, mondayTypeColId);
    const descMon = getColVal(item, mondayDescriptionColId);

    const ticketId = mondayMovideskId || item.id;
    const area = mov.area || item.group?.title || 'N/A';
    const tipo = typeMon || mov.tipo || 'N/A';
    const categoria = mov.categoria || 'N/A';
    const titulo = item.name;
    const descricao = descMon || mov.descricao || item.name || 'N/A';
    const prioridade = prioMon;
    const solicitante = mov.solicitante || 'N/A';
    const equipe = mov.equipe || 'N/A';

    const abertura = mov.abertura ? new Date(mov.abertura).toLocaleDateString('pt-BR') : 'N/A';
    const encerramento = mov.encerramento ? new Date(mov.encerramento).toLocaleDateString('pt-BR') : 'N/A';

    const statusCru = getColVal(item, mondayStatusColId) || 'A fazer';
    let statusFinal = mondayStatusMap[statusCru] || statusCru || 'A fazer';

    const statusUpper = String(statusFinal).trim().toUpperCase();
    if (['CONCLUÍDO', 'CONCLUIDO', 'DONE', 'FEITO'].includes(statusUpper)) statusFinal = 'Feito';
    else if (['A FAZER', 'TO DO'].includes(statusUpper)) statusFinal = 'A fazer';
    else if (['IMPEDIMENTO', 'IMPEDIDO'].includes(statusUpper)) statusFinal = 'Impedimento';
    else if (['WORKING ON IT', 'FAZENDO'].includes(statusUpper)) statusFinal = 'Fazendo';
    else if (['IMPLANTADO', 'IMPLANTAÇÃO', 'IMPLANTACAO'].includes(statusUpper)) statusFinal = 'Implantação';
    else if (['HOMOLOGAÇÃO', 'HOMOLOGACAO'].includes(statusUpper)) statusFinal = 'Homologação';
    else if (['FILA'].includes(statusUpper)) statusFinal = 'Fila';
    else if (['EM ANÁLISE', 'EM ANALISE'].includes(statusUpper)) statusFinal = 'Em Análise';

    const status = statusFinal;
    let tempoResolucao = mov.tempoResolucao ?? 'N/A';

    contagemArea[area] = (contagemArea[area] || 0) + 1;
    contagemTipo[tipo] = (contagemTipo[tipo] || 0) + 1;
    contagemPrio[prioridade] = (contagemPrio[prioridade] || 0) + 1;
    contagemStatus[statusFinal] = (contagemStatus[statusFinal] || 0) + 1;

    const dataAberturaMovideskObj = mov.abertura ? new Date(mov.abertura) : null;
    let semanaLinha = dataAberturaMovideskObj ? calcularSemana(dataAberturaMovideskObj) : 'N/A';
    const mesLinha = dataAberturaMovideskObj ? nomeMes(dataAberturaMovideskObj) : 'N/A';

    if (semanaLinha === 'N/A') {
      // Usando os novos parâmetros para IDs de data do Monday, removendo hardcodes
      const dataMondayCrua = getColVal(item, mondayCreationDateColId);
      if (dataMondayCrua && !isNaN(new Date(dataMondayCrua).getTime())) {
        semanaLinha = calcularSemana(new Date(dataMondayCrua));
      } else {
        semanaLinha = 'Semana 1';
      }
    }

    if (contagemSemanaBurnDown[semanaLinha] !== undefined) {
      contagemSemanaBurnDown[semanaLinha]++;
    } else {
      contagemSemanaBurnDown[semanaLinha] = 1;
    }

    if (['Feito', 'Implantação'].includes(statusFinal)) {
      if (mov.encerramento) {
        const dataEncerramentoMovideskObj = new Date(mov.encerramento);
        if (!isNaN(dataEncerramentoMovideskObj.getTime())) {
          const semanaEncerramentoMovidesk = calcularSemana(dataEncerramentoMovideskObj);
          contagemSemanaFinalizadosMovidesk[semanaEncerramentoMovidesk] = (contagemSemanaFinalizadosMovidesk[semanaEncerramentoMovidesk] || 0) + 1;
        }
      } else if (semanaLinha && semanaLinha !== 'N/A') {
        contagemSemanaFinalizadosMovidesk[semanaLinha] = (contagemSemanaFinalizadosMovidesk[semanaLinha] || 0) + 1;
      }
    }

    let dentroSLA = 'Não';
    let calculoAtraso = 0;

    const dataInicioReal = dataAberturaMovideskObj;
    const dataEncerramentoMovideskObj = mov.encerramento ? new Date(mov.encerramento) : null;
    const dataFimReal = dataEncerramentoMovideskObj;
    if (dataInicioReal && !isNaN(dataInicioReal.getTime())) {
     if (dataFimReal && !isNaN(dataFimReal.getTime()) && ['Feito', 'Implantação', 'Concluído', 'Encerrado'].includes(statusFinal)) {
        const diferencaDias = Math.ceil((dataFimReal - dataInicioReal) / (1000 * 60 * 60 * 24));
        if (diferencaDias <= 15) {
          dentroSLA = 'Sim';
          calculoAtraso = 0;
        } else {
          dentroSLA = 'Não';
          calculoAtraso = diferencaDias - 15;
        }
      } else {
        const agingDiasAtual = Math.ceil((new Date() - dataInicioReal) / (1000 * 60 * 60 * 24));
        if (agingDiasAtual <= 15) {
          dentroSLA = 'Sim';
          calculoAtraso = 0;
        } else {
          dentroSLA = 'Não';
          calculoAtraso = agingDiasAtual - 15;
        }
      }
    }

    if (dentroSLA === 'Sim') {
      dentroSLACount++;
    } else {
      foraSLACount++;
      ticketsForaSLA.push({
        ticket: ticketId,
        area: area,
        titulo: titulo,
        prioridade: prioridade,
        atraso: calculoAtraso
      });
    }

    const isConcluido = ['Feito', 'Implantação'].includes(statusFinal);
    if (isConcluido && typeof tempoResolucao === 'number') {
      totalTempoResolucao += tempoResolucao;
      countResolvidos++;
    } else if (isConcluido && tempoResolucao === 'N/A' && getColVal(item, mondayCreationDateColId) && getColVal(item, mondayResolutionDateColId)) { // <<< Usando os novos parâmetros
      const a = new Date(getColVal(item, mondayCreationDateColId)), e = new Date(getColVal(item, mondayResolutionDateColId));
      if (!isNaN(a.getTime()) && !isNaN(e.getTime())) {
        const mondayTempoResolucao = Math.ceil((e - a) / 86400000);
        totalTempoResolucao += mondayTempoResolucao;
        countResolvidos++;
        tempoResolucao = mondayTempoResolucao;
      }
    }

    const row = ticketsSheet.addRow({
      ticket: ticketId, area: area, tipo: tipo, categoria: categoria, titulo: titulo,
      descricao: descricao, prioridade: prioridade, solicitante: solicitante, equipe: equipe, abertura: abertura, encerramento: encerramento, status: status,
      semana: semanaLinha, mes: mesLinha,
    });

    row.eachCell(cell => { cell.alignment = { vertical: 'middle', wrapText: true }; });
  }

  const totalTicketsDaSprint = itensDaSprint.length;
  const percentualSLA = totalTicketsDaSprint > 0 ? (dentroSLACount / totalTicketsDaSprint) : 0;
  const tempoMedio = countResolvidos > 0 ? (totalTempoResolucao / countResolvidos).toFixed(2) : 0;

  // ============================================================
  // 3.1 NOVA ABA: MAPEAMENTO DE CAMPOS PERSONALIZADOS DE CLIENTES
  //    (Utiliza a listaEstruturaCamposClientes correta)
  // ============================================================
  if (listaEstruturaCamposClientes && listaEstruturaCamposClientes.length > 0) {
    const camposClientesSheet = workbook.addWorksheet('Campos Personalizados Clientes');
    camposClientesSheet.views = [{ showGridLines: true }];

    camposClientesSheet.columns = [
      { header: 'ID do Campo', key: 'customFieldId', width: 15 },
      { header: 'ID da Regra', key: 'customFieldRuleId', width: 15 },
      { header: 'Tipo de Campo', key: 'type', width: 18 },
      { header: 'ID do Item Opção', key: 'customFieldItemId', width: 18 },
      { header: 'Nome/Valor da Opção (customFieldItem)', key: 'customFieldItem', width: 40 }
    ];

    camposClientesSheet.getRow(1).eachCell((cell) => {
      cell.fill = estiloAzul;
      cell.font = fonteBranca;
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });
    camposClientesSheet.getRow(1).height = 24;

    for (const campo of listaEstruturaCamposClientes) {
      if (!campo.items || campo.items.length === 0) {
        const row = camposClientesSheet.addRow({
          customFieldId: campo.customFieldId,
          customFieldRuleId: campo.customFieldRuleId,
          type: campo.type || 'N/A',
          customFieldItemId: 'N/A',
          customFieldItem: '(Campo de preenchimento livre)'
        });
        row.eachCell(cell => { cell.alignment = { vertical: 'middle' }; });
      } else {
        for (const item of campo.items) {
          const row = camposClientesSheet.addRow({
            customFieldId: campo.customFieldId,
            customFieldRuleId: campo.customFieldRuleId,
            type: campo.type || 'N/A',
            customFieldItemId: item.customFieldItemId,
            customFieldItem: item.customFieldItem
          });
          row.eachCell(cell => { cell.alignment = { vertical: 'middle' }; });
        }
      }
    }
    console.log(`[Excel] ${listaEstruturaCamposClientes.length} estruturas de campos adicionadas à planilha da sprint.`);
  }

  // --- ABA DASHBOARD ---
  const dashSheet = workbook.addWorksheet('Dashboard');
  dashSheet.views = [{ showGridLines: true }];

  dashSheet.mergeCells('A1:C1');
  dashSheet.getCell('A1').value = 'Dashboard de Change Log - Governança de TI';
  dashSheet.getCell('A1').font = { bold: true, size: 16, color: { argb: 'FF034C8C' } };

  dashSheet.getCell('A2').value = 'Período';
  dashSheet.getCell('B2').value = sprintNome.toUpperCase();
  dashSheet.getCell('D2').value = 'Base';
  dashSheet.mergeCells('E2:H2');
  dashSheet.getCell('E2').value = 'Tickets corporativos para intranet / Power BI / Excel';

  dashSheet.getCell('A4').value = 'Total de tickets';
  dashSheet.getCell('A5').value = totalTicketsDaSprint;
  dashSheet.getCell('A5').font = fonteNegrito;

  dashSheet.getCell('D4').value = '% dentro do SLA';
  dashSheet.getCell('D5').value = percentualSLA;
  dashSheet.getCell('D5').numFmt = '0.00%';
  dashSheet.getCell('D5').font = fonteNegrito;

  dashSheet.getCell('G4').value = 'Tempo médio (dias)';
  dashSheet.getCell('G5').value = parseFloat(tempoMedio);
  dashSheet.getCell('G5').font = fonteNegrito;

  dashSheet.getCell('J4').value = 'Áreas atendidas';
  dashSheet.getCell('J5').value = Object.keys(contagemArea).length;
  dashSheet.getCell('J5').font = fonteNegrito;

  ['A4', 'D4', 'G4', 'J4'].forEach(pos => {
    dashSheet.getCell(pos).fill = estiloAzul;
    dashSheet.getCell(pos).font = fonteBranca;
    dashSheet.getCell(pos).alignment = { horizontal: 'center', vertical: 'middle' };
  });
  ['A5', 'D5', 'G5', 'J5'].forEach(pos => {
    dashSheet.getCell(pos).alignment = { horizontal: 'center', vertical: 'middle' };
    dashSheet.getCell(pos).border = {
      bottom: { style: 'thin' }, top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }
    };
  });

  dashSheet.getCell('J21').value = 'Status final';
  dashSheet.getCell('J21').font = fonteNegrito;
  ['Feito', 'Implantação', 'Encerrado'].forEach((s, i) => {
    dashSheet.getCell(`J${22 + i}`).value = s;
    dashSheet.getCell(`K${22 + i}`).value = contagemStatus[s] || 0;
  });

  // --- ABA RESUMO ---
  const resumoSheet = workbook.addWorksheet('Resumo');
  resumoSheet.views = [{ showGridLines: true }];

  resumoSheet.getCell('A1').value = 'Resumo Executivo'; resumoSheet.getCell('A1').font = fonteNegrito;
  resumoSheet.getCell('A2').value = 'Total de tickets'; resumoSheet.getCell('B2').value = totalTicketsDaSprint;
  resumoSheet.getCell('A5').value = percentualSLA; resumoSheet.getCell('B5').numFmt = '0.00%';
  resumoSheet.getCell('A6').value = parseFloat(tempoMedio);
  resumoSheet.getCell('A7').value = Object.keys(contagemArea).length;

  // Tickets por Área (CORRIGIDO: ordenação por valor)
  resumoSheet.getCell('A9').value = 'Tickets por Área'; resumoSheet.getCell('A9').font = fonteNegrito;
  Object.entries(contagemArea)
    .sort((a, b) => b[1] - a[1]) // <<< CORRIGIDO
    .forEach(([a, q], i) => {
      resumoSheet.getCell(`A${10 + i}`).value = a;
      resumoSheet.getCell(`B${10 + i}`).value = q;
    });

  // Distribuição por Tipo (CORRIGIDO: ordenação por valor)
  resumoSheet.getCell('D9').value = 'Distribuição por Tipo'; resumoSheet.getCell('D9').font = fonteNegrito;
  Object.entries(contagemTipo)
    .sort((a, b) => b[1] - a[1]) // <<< CORRIGIDO
    .forEach(([t, q], i) => {
      resumoSheet.getCell(`D${10 + i}`).value = t;
      resumoSheet.getCell(`E${10 + i}`).value = q;
    });

  // Prioridades Ordenadas Estruturadas
  resumoSheet.getCell('G9').value = 'Prioridade'; resumoSheet.getCell('G9').font = fonteNegrito;
  const prioridadesOrdenadas = ['Muito Alta', 'Alta', 'Média', 'Baixa', 'Não definida'];
  prioridadesOrdenadas.forEach((p, i) => {
    resumoSheet.getCell(`G${10 + i}`).value = p;
    resumoSheet.getCell(`H${10 + i}`).value = contagemPrio[p] || 0;
  });

  // Cumprimento SLA
  resumoSheet.getCell('J9').value = 'Cumprimento SLA'; resumoSheet.getCell('J9').font = fonteNegrito;
  resumoSheet.getCell('J10').value = 'Sim'; resumoSheet.getCell('K10').value = dentroSLACount;
  resumoSheet.getCell('J11').value = 'Não'; resumoSheet.getCell('K11').value = foraSLACount;

  // Status Final (Movidesk)
  resumoSheet.getCell('M9').value = 'Status Final'; resumoSheet.getCell('M9').font = fonteNegrito;

  const listaStatusOficiais = ['Feito', 'A fazer', 'Impedimento', 'Fazendo', 'Implantação', 'Homologação', 'Fila', 'Em Análise'];
  listaStatusOficiais.forEach((s, i) => {
    resumoSheet.getCell(`M${10 + i}`).value = s;
    resumoSheet.getCell(`N${10 + i}`).value = contagemStatus[s] || 0;
  });

  // Proteção Visual contra sobreposição: Injeta Semanas nas colunas P e Q
  resumoSheet.getCell('P9').value = 'Tickets por Semana'; resumoSheet.getCell('P9').font = fonteNegrito;
  resumoSheet.getCell('Q9').value = 'Volume'; resumoSheet.getCell('Q9').font = fonteNegrito;
  Object.entries(contagemSemanaFinalizadosMovidesk)
    .sort((a, b) => a.localeCompare(b))
    .forEach(([semana, qtd], i) => {
      resumoSheet.getCell(`P${10 + i}`).value = semana;
      resumoSheet.getCell(`Q${10 + i}`).value = qtd;
    });

  // Top 10 Fora do SLA (Ofensores Críticos)
  resumoSheet.getCell('D19').value = 'Top tickets fora do SLA'; resumoSheet.getCell('D19').font = fonteNegrito;
  resumoSheet.getRow(20).height = 18;
  ['D20', 'E20', 'F20', 'G20'].forEach((c, i) => {
    const cell = resumoSheet.getCell(c);
    cell.value = ['Ticket', 'Área', 'Atraso (dias)', 'Título'][i];
    cell.fill = estiloAzul;
    cell.font = fonteBranca;
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  });

  const topForaSLA = ticketsForaSLA
    .filter(t => typeof t.atraso === 'number')
    .sort((a, b) => b.atraso - a.atraso)
    .slice(0, 10);

  topForaSLA.forEach((t, i) => {
    const rIdx = 21 + i;
    resumoSheet.getCell(`D${rIdx}`).value = t.ticket;
    resumoSheet.getCell(`E${rIdx}`).value = t.area;
    resumoSheet.getCell(`F${rIdx}`).value = t.atraso;
    resumoSheet.getCell(`G${rIdx}`).value = t.titulo;
    resumoSheet.getRow(rIdx).alignment = { vertical: 'middle', wrapText: true };
  });

  resumoSheet.getColumn(1).width = 30;
  resumoSheet.getColumn(2).width = 15;
  resumoSheet.getColumn(4).width = 18;
  resumoSheet.getColumn(5).width = 25;
  resumoSheet.getColumn(6).width = 15;
  resumoSheet.getColumn(7).width = 35;

  const nomeArquivo = `monday_sprint_${sprintNome.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, '_').replace(/[^a-z0-9_\-]/g, '')}.xlsx`;

  await fs.mkdir(dirPath, { recursive: true });
  await workbook.xlsx.writeFile(path.join(dirPath, nomeArquivo));
  console.log(`[Monday] Excel gerado com sucesso para sprint "${sprintNome}": ${nomeArquivo}`);

  return {
    nome_exibicao: `MONDAY - ${sprintNome.toUpperCase()}`,
    caminho_arquivo: path.join('utils', nomeArquivo)
  };
}

// =========================================================================
// 7. ORQUESTRADOR PRINCIPAL DA ATUALIZAÇÃO SPRINT (CORRIGIDO)
// =========================================================================
async function sincronizarMondaySprints(
    token,                              // 1. Token do Monday
    boardId,                            // 2. ID do Board do Monday
    mondayTextColId,                    // 3. ID da coluna de texto (ID do Movidesk)
    mondayStatusColId,                  // 4. ID da coluna de status
    mondayPriorityColId,                // 5. ID da coluna de prioridade
    mondayDescriptionColId,             // 6. ID da coluna de descrição
    mondayTypeColId,                    // 7. ID da coluna de tipo
    mondayCreationDateColId,            // 8. ID da coluna de data de criação no Monday (Agora vem do server.js)
    mondayResolutionDateColId,          // 9. ID da coluna de data de resolução no Monday (Agora vem do server.js)
    movideskClientSectorCustomFieldId,  // 10. ID do campo personalizado "Setor"
    dirPath                             // 11. Caminho do diretório de salvamento
) {
    console.log('[Monday] Iniciando sincronização completa de sprints...');

    const movideskToken = process.env.MOVIDESK_TOKEN;

    // 1. Busca os VALORES dos campos personalizados de TODAS as pessoas/clientes do Movidesk.
    const personCustomFieldValuesMap = movideskToken
        ? await buscarValoresCamposPersonalizadosPessoas(movideskToken)
        : {};

    // 2. Busca a ESTRUTURA dos campos personalizados de clientes (para a aba de mapeamento no Excel).
    const listaEstruturaCamposClientes = movideskToken
        ? await buscarEstruturaCamposPersonalizadosClientes(movideskToken)
        : [];

    // 3. Busca todos os tickets do Movidesk, INJETANDO o setor do cliente obtido no passo 1.
    const mapaMovidesk = movideskToken
        ? await buscarTodosTicketsMovidesk(movideskToken, personCustomFieldValuesMap, movideskClientSectorCustomFieldId)
        : {};

    let todosItensMonday;
    try {
        todosItensMonday = await buscarTodosItensMonday(boardId, token);
    } catch (err) {
        throw new Error(`[Monday] Erro crítico na extração de dados do Monday: ${err.message}`);
    }

    const sprints = {};
    for (const item of todosItensMonday) {
        const nomeGrupo = item.group?.title ? item.group.title.trim() : 'Sem Sprint';

        const nomeGrupoUpper = nomeGrupo.toUpperCase();
        if (['BACKLOG', 'PROXIMA SPRINT', 'PRÓXIMA SPRINT', 'SPRINT ATUAL'].includes(nomeGrupoUpper)) {
            continue;
        }

        if (!sprints[nomeGrupo]) sprints[nomeGrupo] = [];
        sprints[nomeGrupo].push(item);
    }

    console.log(`[Monday] ${Object.keys(sprints).length} sprints legítimas mapeadas para processamento.`);

    const resultados = [];
    for (const [sprintNome, itensDaSprint] of Object.entries(sprints)) {
        try {
            console.log(`[Monday] Gerando lote da planilha para: "${sprintNome}" (${itensDaSprint.length} registros).`);

            const res = await gerarExcelParaSprint(
                sprintNome,
                itensDaSprint,
                mapaMovidesk,
                dirPath,
                mondayTextColId,
                mondayStatusColId,
                mondayPriorityColId,
                mondayDescriptionColId,
                mondayTypeColId,
                mondayCreationDateColId, // <<< PARÂMETRO AGORA USADO CORRETAMENTE
                mondayResolutionDateColId, // <<< PARÂMETRO AGORA USADO CORRETAMENTE
                listaEstruturaCamposClientes
            );

            resultados.push(res);
        } catch (err) {
            console.error(`[Monday] Falha ao processar a fatia da sprint "${sprintNome}":`, err.message, err.stack);
        }
    }

    // =========================================================================
    // MOTOR DE MAPA DE ARQUIVOS ESTÁTICOS COMPATÍVEL COM GITHUB PAGES / OFFLINE
    // =========================================================================
    try {
        const arquivosNaPasta = await fs.readdir(dirPath);

        const listaCompletaSprints = arquivosNaPasta
            .filter(f => f.endsWith('.xlsx'))
            .map(f => {
                let nomeExibicao = f.replace('.xlsx', '').toUpperCase().replace(/_/g, ' ');
                if (nomeExibicao.startsWith('TICKETS SPRINT ')) {
                    nomeExibicao = `MOVIDESK - ${nomeExibicao.replace('TICKETS SPRINT ', '')}`;
                } else if (nomeExibicao.startsWith('MONDAY SPRINT ')) {
                    nomeExibicao = `MONDAY - ${nomeExibicao.replace('MONDAY SPRINT ', '')}`;
                }

                return {
                    nome_exibicao: nomeExibicao,
                    caminho_arquivo: `utils/${f}`
                };
            })
            .sort((a, b) => b.nome_exibicao.localeCompare(a.nome_exibicao));

        const caminhoDoIndiceJson = path.join(dirPath, 'sprints.json');
        await fs.writeFile(caminhoDoIndiceJson, JSON.stringify(listaCompletaSprints, null, 2), 'utf8');
        console.log(`[Altona Engine] Índice de arquivos estáticos criado em: utils/sprints.json`);

    } catch (errDiscorv) {
        console.error('[Altona Engine] Erro ao tentar mapear arquivos estáticos da pasta utils:', errDiscorv.message);
    }

    return resultados;
}

module.exports = { sincronizarMondaySprints };

