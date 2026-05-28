const axios = require('axios');
const ExcelJS = require('exceljs');
const path = require('path');
const https = require('https');
const fs = require('fs').promises;

// ==========================================
// 1. CONFIGURAÇÕES E MAPAS DE DE-PARA
// ==========================================
// Mapeamento de status brutos para status padronizados
const mondayStatusMap = {
  "Concluído": "Feito", "Done": "Feito", "Feito": "Feito",
  "Implantado": "Implantação", "Implantação": "Implantação",
  "Homologação": "Homologação", "Em andamento": "Fazendo",
  "Working on it": "Fazendo", "A fazer": "A fazer",
  "To Do": "A fazer", "Impedimento": "Impedimento",
  "Encerrado": "Feito", // Encerrado pode ser considerado como feito para fins de SLA/conclusão
  "Cancelado": "Cancelado",
  "Fila": "Fila", "Em Análise": "Em Análise",
  "": "A fazer" // Default para vazio
};

// Mapeamento de prioridades brutos para prioridades padronizadas
const mondayPriorityMap = {
  "Alta": "Alta", "High": "Alta",
  "Média": "Média", "Medium": "Média",
  "Baixa": "Baixa", "Low": "Baixa",
  "Muito Alta": "Muito Alta", "Critical": "Muito Alta",
  "": "Não definida"
};

// Mapeamento de códigos de urgência Movidesk para prioridades padronizadas
const urgencyMap = {
  53857: "Alta",
  53859: "Baixa",
  53858: "Média",
  53860: "Muito Alta"
};

// Agente HTTPS para permitir requisições sem validação de certificado (útil em ambientes de desenvolvimento/proxy)
const httpsAgent = new https.Agent({ rejectUnauthorized: false });

// ==========================================
// 2. FUNÇÕES AUXILIARES DE DATA
// ==========================================
/**
 * Calcula a semana do mês para uma dada data.
 * @param {Date} dateObj - Objeto Date.
 * @returns {string} - 'Semana 1', 'Semana 2', etc., ou 'Semana 1' se a data for inválida.
 */
function calcularSemana(dateObj) {
  if (!dateObj || isNaN(dateObj.getTime())) return 'Semana 1'; // Default para Semana 1 se inválido
  const dia = dateObj.getDate();
  if (dia <= 7) return 'Semana 1';
  if (dia <= 14) return 'Semana 2';
  if (dia <= 21) return 'Semana 3';
  if (dia <= 28) return 'Semana 4';
  return 'Semana 5';
}

/**
 * Retorna o nome do mês em português para uma dada data.
 * @param {Date} dateObj - Objeto Date.
 * @returns {string} - Nome do mês ou 'Indefinido' se a data for inválida.
 */
function nomeMes(dateObj) {
  if (!dateObj || isNaN(dateObj.getTime())) return 'Indefinido';
  const meses = [
    'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
    'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
  ];
  return meses[dateObj.getMonth()];
}

// ==========================================
// 3. INTEGRAÇÃO DA API MOVIDESK COM RETRY
// ==========================================

// Cache para evitar chamadas repetidas de API para a mesma pessoa/setor
const cacheSetoresPessoas = {}; // O nome do cache pode permanecer genérico

/**
 * Busca o setor de um solicitante (cliente) no Movidesk usando o endpoint /clients.
 * Este é o endpoint correto para IDs GUIDs de solicitantes.
 * @param {string} clientId - ID do cliente no Movidesk (GUID).
 * @param {string} movideskToken - Token de autenticação da API Movidesk.
 * @returns {Promise<string>} - O nome do setor ou 'Área não identificada'.
 */
async function buscarSetorPorClienteId(clientId, movideskToken) {
  if (!clientId || String(clientId).trim().toUpperCase() === 'N/A' || String(clientId).trim() === '') {
    return 'Área não identificada';
  }

  // Se já buscamos essa pessoa antes nesta execução, retorna o setor salvo na memória
  if (cacheSetoresPessoas[clientId]) {
    return cacheSetoresPessoas[clientId];
  }

  try {
    // Consulta a API de clientes do Movidesk para obter campos customizados
    // NOTA: Para IDs GUIDs de solicitantes, o endpoint correto é /clients/{id}
    const urlClient = `https://api.movidesk.com/public/v1/clients/${clientId}`;

    // console.log(`[Movidesk][Setor] Buscando setor para clientId: ${clientId}`); // DEBUG

    const resClient = await axios.get(urlClient, {
      params: {
        'token': movideskToken,
        '$expand': 'customFieldValues', // Crucial para obter campos customizados
      },
      httpsAgent,
      timeout: 15000, // Timeout específico para esta chamada
    });

    const clienteEncontrado = resClient.data;

    if (clienteEncontrado && clienteEncontrado.customFieldValues) {
      // Procura pelo campo customizado 'setor' (case-insensitive)
      const campoSetor = clienteEncontrado.customFieldValues.find(cf =>
        (cf.customFieldName && String(cf.customFieldName).trim().toLowerCase() === 'setor') ||
        (cf.customField && String(cf.customField).trim().toLowerCase() === 'setor')
      );

      if (campoSetor) {
        const valorSetor = String(campoSetor.customFieldItem || campoSetor.value || '').trim();
        if (valorSetor && valorSetor.toUpperCase() !== 'N/A' && valorSetor !== '') {
          cacheSetoresPessoas[clientId] = valorSetor;
          // console.log(`[Movidesk][Setor] Setor encontrado para ${clientId}: ${valorSetor}`); // DEBUG
          return valorSetor;
        }
      }
    }
  } catch (error) {
    console.warn(`[Movidesk][Setor] Falha ao buscar setor para cliente ID ${clientId}: ${error.message}`);
    if (error.response) {
      // console.warn(`[Movidesk][Setor] Resposta de erro: ${JSON.stringify(error.response.data)}`); // DEBUG
      // Um 404 aqui pode significar que o ID do cliente não existe ou não tem permissão
    }
  }

  // Se não encontrou o campo, deu erro, ou o valor é vazio/N/A, armazena e retorna o padrão
  cacheSetoresPessoas[clientId] = 'Área não identificada';
  return 'Área não identificada';
}


/**
 * Busca todos os tickets do Movidesk de forma paginada.
 * @param {string} movideskToken - Token de autenticação da API Movidesk.
 * @returns {Promise<Object>} - Um mapa de tickets Movidesk indexado pelo ID do ticket.
 */
async function buscarTodosTicketsMovidesk(movideskToken) {
  const mapa = {};
  let skip = 0;
  const take = 100;
  let total = 0;
  const maxRetries = 3;
  const retryDelayMs = 2000;
  const urlMovideskTickets = `https://api.movidesk.com/public/v1/tickets`;

  console.log('[Movidesk] Iniciando busca paginada de tickets...');

  while (true) {
    let resp;
    let retries = 0;
    let success = false;

    while (retries < maxRetries && !success) {
      try {
        resp = await axios.get(urlMovideskTickets, {
          params: {
            token: movideskToken,
            '$select': 'id,subject,category,urgency,status,createdDate,resolvedIn,closedIn,ownerTeam,owner,serviceFirstLevel,slaSolutionDate',
            '$expand': 'actions,createdBy',
            '$filter': "createdDate gt 2023-01-01T00:00:00Z",
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
          await new Promise(resolve => setTimeout(resolve, retryDelayMs));
        } else {
          console.error(`[Movidesk] Falha após ${maxRetries} tentativas para skip=${skip}. Abortando.`);
          return mapa;
        }
      }
    }

    if (!success) break;

    const lista = resp.data.value || (Array.isArray(resp.data) ? resp.data : []);
    if (lista.length === 0) break;

    for (const t of lista) {
      const equipe = t.ownerTeam ? String(t.ownerTeam).trim() : 'N/A';
      if (equipe === 'Suporte e Infraestrutura' || equipe === 'Administradores') continue;

      const prioridadeMovidesk = urgencyMap[t.urgency] || 'N/A';
      const tipo = t.serviceFirstLevel ? String(t.serviceFirstLevel).trim() : 'N/A';
      const createdDate = t.createdDate ? new Date(t.createdDate) : null;
      const resolvedInDate = t.resolvedIn ? new Date(t.resolvedIn) : null;
      const closedInDate = t.closedIn ? new Date(t.closedIn) : null;

      const finalizationDate = resolvedInDate || closedInDate;

      let tempoResolucaoDias = 'N/A';
      if (createdDate && finalizationDate && !isNaN(createdDate.getTime()) && !isNaN(finalizationDate.getTime())) {
        tempoResolucaoDias = Math.ceil((finalizationDate - createdDate) / (1000 * 60 * 60 * 24));
      }

      const descRaw = t.actions?.[0]?.description || t.actions?.[0]?.body || '';
      const descricao = String(descRaw)
        .replace(/<[^>]*>?/gm, '')
        .replace(/\[cid:[^\]]*\]/g, '')
        .replace(/[\r\n\t]+/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
        .substring(0, 300) || 'N/A';

      const solicitante = t.createdBy?.personName || t.createdBy?.name || t.createdBy?.businessName || 'N/A';
      const solicitanteId = t.createdBy?.id; // ID do cliente (GUID)
      const setorRealMovidesk = await buscarSetorPorClienteId(solicitanteId, movideskToken);

      mapa[String(t.id)] = {
        area: setorRealMovidesk,
        tipo,
        categoria: t.category || 'N/A',
        descricao,
        prioridade: prioridadeMovidesk,
        solicitante: String(solicitante).trim(),
        equipe: equipe,
        abertura: createdDate ? createdDate.toISOString() : null,
        encerramento: finalizationDate ? finalizationDate.toISOString() : null,
        status: t.status || 'N/A',
        tempoResolucao: tempoResolucaoDias,
        mesAbertura: createdDate ? nomeMes(createdDate) : 'Indefinido'
      };

      total++;
    }

    console.log(`[Movidesk] ${total} tickets carregados (skip=${skip})...`);
    skip += take;
  }

  console.log(`[Movidesk] Total final: ${total} tickets carregados.`);
  return mapa;
}

// ==========================================
// 4. INTEGRAÇÃO DA API MONDAY
// ==========================================
/**
 * Busca todos os itens de um board Monday.com de forma paginada.
 * @param {string} boardId - ID do board Monday.com.
 * @param {string} token - Token de autenticação da API Monday.com.
 * @returns {Promise<Array<Object>>} - Lista de itens do board.
 */
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
                value
              }
            }
          }
        }
      }
    `;

    const response = await axios.post('https://api.monday.com/v2',
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

/**
 * Extrai o valor de texto de uma coluna de item Monday.com.
 * @param {Object} item - Objeto de item Monday.com.
 * @param {string} id - ID da coluna.
 * @returns {string} - Texto da coluna ou string vazia.
 */
const getColVal = (item, id) => {
  const c = item.column_values?.find(cv => cv.id === id);
  return c?.text?.trim() || '';
};

// ==========================================
// 5. MOTOR DE PROCESSAMENTO DO EXCEL
// ==========================================
/**
 * Gera um arquivo Excel com os dados da sprint e dashboards.
 * @param {string} sprintNome - Nome da sprint.
 * @param {Array<Object>} itensDaSprint - Lista de itens Monday.com da sprint.
 * @param {Object} mapaMovidesk - Mapa de tickets Movidesk para enriquecimento de dados.
 * @param {string} dirPath - Caminho do diretório para salvar o Excel.
 * @param {string} mondayTextColId - ID da coluna Monday com o ID do Movidesk.
 * @param {string} mondayStatusColId - ID da coluna Monday de status.
 * @param {string} mondayPriorityColId - ID da coluna Monday de prioridade.
 * @param {string} mondayDescriptionColId - ID da coluna Monday de descrição.
 * @param {string} mondayTypeColId - ID da coluna Monday de tipo.
 * @param {string} mondaySlaDateColId - ID da coluna Monday de data SLA (opcional, não usado no SLA de 15 dias).
 * @returns {Promise<Object>} - Objeto com nome de exibição e caminho do arquivo.
 */
async function gerarExcelParaSprint(sprintNome, itensDaSprint, mapaMovidesk, dirPath,
  mondayTextColId, mondayStatusColId, mondayPriorityColId, mondayDescriptionColId, mondayTypeColId, mondaySlaDateColId) { // mondaySlaDateColId é mantido para compatibilidade, mas não usado diretamente no SLA de 15 dias.

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
  const contagemSemanaBurnDown = {}; // Contagem de tickets por semana de abertura
  const contagemSemanaFinalizadosMovidesk = {}; // Contagem de tickets finalizados por semana de encerramento
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

  // Loop para processamento de linhas e contadores
  for (const item of itensDaSprint) {
    let mondayMovideskId = getColVal(item, mondayTextColId);
    const mov = mondayMovideskId ? (mapaMovidesk[mondayMovideskId] || {}) : {};

    // ============================================================
    // 1. PADRONIZAÇÃO DO STATUS (ANTECIPADA PARA USO NOS FALLBACKS)
    // ============================================================
    const statusCru = getColVal(item, mondayStatusColId) || 'A fazer';
    // Aplica o mapeamento direto primeiro
    let statusFinal = mondayStatusMap[statusCru] || statusCru;

    // Ajustes de status para garantir consistência após o mapeamento
    const statusUpper = String(statusFinal).trim().toUpperCase();
    if (['CONCLUÍDO', 'CONCLUIDO', 'DONE'].includes(statusUpper)) statusFinal = 'Feito';
    else if (['A FAZER', 'TO DO'].includes(statusUpper)) statusFinal = 'A fazer';
    else if (['IMPEDIMENTO', 'IMPEDIDO'].includes(statusUpper)) statusFinal = 'Impedimento';
    else if (['WORKING ON IT', 'FAZENDO', 'EM ANDAMENTO'].includes(statusUpper)) statusFinal = 'Fazendo';
    else if (['IMPLANTADO', 'IMPLANTAÇÃO', 'IMPLANTACAO'].includes(statusUpper)) statusFinal = 'Implantação';
    else if (['HOMOLOGAÇÃO', 'HOMOLOGACAO'].includes(statusUpper)) statusFinal = 'Homologação';
    else if (['FILA'].includes(statusUpper)) statusFinal = 'Fila';
    else if (['EM ANÁLISE', 'EM ANALISE'].includes(statusUpper)) statusFinal = 'Em Análise';
    else if (['CANCELADO'].includes(statusUpper)) statusFinal = 'Cancelado';
    else if (['ENCERRADO'].includes(statusUpper)) statusFinal = 'Feito'; // Encerrado é tratado como Feito
    else statusFinal = 'A fazer'; // Default final para qualquer outro status não mapeado

    const status = statusFinal;

    // ============================================================
    // 2. EXTRAÇÃO E PADRONIZAÇÃO COMPLETA DE CAMPOS ("NÃO IDENTIFICADO")
    // ============================================================
    const prioMon = mondayPriorityMap[getColVal(item, mondayPriorityColId)] || 'Prioridade não definida';
    const typeMon = getColVal(item, mondayTypeColId);
    const descMon = getColVal(item, mondayDescriptionColId);

    const ticketId = mondayMovideskId || item.id || 'Ticket não identificado';

    // Pega o dado dinâmico enviado automaticamente pela API do Movidesk (Parte 1)
    let area = mov.area || 'Área não identificada';
    if (!area || area.trim().toUpperCase() === 'N/A' || area === 'Sem Sprint') {
      area = 'Área não identificada';
    }

    // Tipo (Ajustado para "Tipo não identificado")
    let tipo = typeMon || mov.tipo || 'Tipo não identificado';
    if (tipo.trim().toUpperCase() === 'N/A') tipo = 'Tipo não identificado';

    // Categoria
    let categoria = mov.categoria || 'Categoria não identificada';
    if (categoria.trim().toUpperCase() === 'N/A') categoria = 'Categoria não identificada';

    const titulo = item.name || 'Título não identificado';
    const descricao = descMon || mov.descricao || item.name || 'Descrição não identificada';
    const prioridade = prioMon;

    // Solicitante
    let solicitante = mov.solicitante || 'Solicitante não identificado';
    if (solicitante.trim().toUpperCase() === 'N/A') solicitante = 'Solicitante não identificado';

    // Equipe TI
    let equipe = mov.equipe || 'Equipe não identificada';
    if (equipe.trim().toUpperCase() === 'N/A') equipe = 'Equipe não identificada';

    // Captura e validação das datas nativas da Monday para contingência
    const dataMondayCriacaoRaw = getColVal(item, 'log_de_cria__o3'); // Coluna de criação Monday
    const dataMondayConclusaoRaw = getColVal(item, 'data'); // Coluna de conclusão/resolução Monday

    const dataMondayCriacaoObj = (dataMondayCriacaoRaw && !isNaN(new Date(dataMondayCriacaoRaw).getTime())) ? new Date(dataMondayCriacaoRaw) : null;
    const dataMondayConclusaoObj = (dataMondayConclusaoRaw && !isNaN(new Date(dataMondayConclusaoRaw).getTime())) ? new Date(dataMondayConclusaoRaw) : null;


    // Formatação de Abertura
    let abertura = 'Abertura não identificada';
    let dataReferenciaAberturaObj = null;
    if (mov.abertura) {
      dataReferenciaAberturaObj = new Date(mov.abertura);
      abertura = dataReferenciaAberturaObj.toLocaleDateString('pt-BR');
    } else if (dataMondayCriacaoObj) {
      dataReferenciaAberturaObj = dataMondayCriacaoObj;
      abertura = dataMondayCriacaoObj.toLocaleDateString('pt-BR');
    }

    // Formatação de Encerramento
    let encerramento = 'Encerramento não identificada';
    let dataReferenciaEncerramentoObj = null;
    if (mov.encerramento) {
      dataReferenciaEncerramentoObj = new Date(mov.encerramento);
      encerramento = dataReferenciaEncerramentoObj.toLocaleDateString('pt-BR');
    } else if (['Feito', 'Implantação'].includes(statusFinal) && dataMondayConclusaoObj) {
      dataReferenciaEncerramentoObj = dataMondayConclusaoObj;
      encerramento = dataMondayConclusaoObj.toLocaleDateString('pt-BR');
    }


    // ============================================================
    // 3. CONTEXTO DE SEMANA E MÊS (PADRONIZADOS)
    // ============================================================
    let semanaLinha = dataReferenciaAberturaObj ? calcularSemana(dataReferenciaAberturaObj) : 'Semana não identificada';

    let mesLinha = 'Mês não identificado';
    if (mov.mesAbertura && mov.mesAbertura !== 'Indefinido' && mov.mesAbertura !== 'N/A') {
      mesLinha = mov.mesAbertura;
    } else if (dataReferenciaAberturaObj) {
      mesLinha = nomeMes(dataReferenciaAberturaObj);
    }

    // Alimentação dos acumuladores estatísticos das abas de Dashboard
    contagemArea[area] = (contagemArea[area] || 0) + 1;
    contagemTipo[tipo] = (contagemTipo[tipo] || 0) + 1;
    contagemPrio[prioridade] = (contagemPrio[prioridade] || 0) + 1;
    contagemStatus[statusFinal] = (contagemStatus[statusFinal] || 0) + 1;

    if (semanaLinha !== 'Semana não identificada') {
      if (contagemSemanaBurnDown[semanaLinha] !== undefined) {
        contagemSemanaBurnDown[semanaLinha]++;
      } else {
        contagemSemanaBurnDown[semanaLinha] = 1;
      }
    }

    // ============================================================
    // 4. CONTAGEM CRÍTICA DE ACORDOS SEMANAIS (MÓDULO DE ENTREGA)
    // ============================================================
    if (['Feito', 'Implantação'].includes(statusFinal)) {
      if (dataReferenciaEncerramentoObj && !isNaN(dataReferenciaEncerramentoObj.getTime())) {
        const semanaEncerramento = calcularSemana(dataReferenciaEncerramentoObj);
        contagemSemanaFinalizadosMovidesk[semanaEncerramento] = (contagemSemanaFinalizadosMovidesk[semanaEncerramento] || 0) + 1;
      } else if (semanaLinha && semanaLinha !== 'Semana não identificada') {
        // Fallback para semana de abertura se não houver data de encerramento válida
        contagemSemanaFinalizadosMovidesk[semanaLinha] = (contagemSemanaFinalizadosMovidesk[semanaLinha] || 0) + 1;
      }
    }

    // ============================================================
    // 5. MOTOR DE SLA DA SPRINT: CRITÉRIO DOS 15 DIAS CORRIDOS
    // ============================================================
    let dentroSLA = 'Não';
    let calculoAtraso = 0;

    const dataInicioReal = dataReferenciaAberturaObj;
    const dataFimReal = dataReferenciaEncerramentoObj; // Já padronizado acima

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
      } else if (!['Feito', 'Implantação', 'Concluído', 'Encerrado', 'Cancelado'].includes(statusFinal)) {
        // Se o ticket ainda não foi finalizado e não foi cancelado
        const agingDiasAtual = Math.ceil((new Date() - dataInicioReal) / (1000 * 60 * 60 * 24));
        if (agingDiasAtual <= 15) {
          dentroSLA = 'Sim';
          calculoAtraso = 0;
        } else {
          dentroSLA = 'Não';
          calculoAtraso = agingDiasAtual - 15;
        }
      }
      // Se for cancelado, não entra no cálculo de SLA de atraso
    }

    if (dentroSLA === 'Sim') {
      dentroSLACount++;
    } else if (statusFinal !== 'Cancelado') { // Não conta cancelados como fora do SLA para este relatório
      foraSLACount++;
      ticketsForaSLA.push({
        ticket: ticketId,
        area: area,
        titulo: titulo,
        prioridade: prioridade,
        atraso: calculoAtraso
      });
    }

    // ============================================================
    // 6. CÁLCULO DO TEMPO DE RESOLUÇÃO EM DIAS CORRIDOS
    // ============================================================
    let tempoResolucao = mov.tempoResolucao ?? 'N/A';
    const isConcluido = ['Feito', 'Implantação'].includes(statusFinal);

    if (isConcluido) {
      if (typeof tempoResolucao === 'number') {
        totalTempoResolucao += tempoResolucao;
        countResolvidos++;
      } else if (dataReferenciaAberturaObj && dataReferenciaEncerramentoObj) {
        const mondayTempoResolucao = Math.ceil((dataReferenciaEncerramentoObj - dataReferenciaAberturaObj) / (1000 * 60 * 60 * 24));
        if (!isNaN(mondayTempoResolucao) && mondayTempoResolucao >= 0) { // Garante que é um número válido e não negativo
          totalTempoResolucao += mondayTempoResolucao;
          countResolvidos++;
          tempoResolucao = mondayTempoResolucao;
        }
      }
    }


    // Injeta a linha completamente padronizada no ExcelJS
    const row = ticketsSheet.addRow({
      ticket: ticketId, area: area, tipo: tipo, categoria: categoria, titulo: titulo,
      descricao: descricao, prioridade: prioridade, solicitante: solicitante, equipe: equipe,
      abertura: abertura, encerramento: encerramento, status: status,
      semana: semanaLinha, mes: mesLinha,
    });

    row.eachCell(cell => { cell.alignment = { vertical: 'middle', wrapText: true }; });
  }


  const totalTicketsDaSprint = itensDaSprint.length;
  // O cálculo de percentualSLA deve considerar apenas tickets que não são 'Cancelado'
  const totalTicketsNaoCancelados = itensDaSprint.filter(item => {
    const statusCru = getColVal(item, mondayStatusColId) || 'A fazer';
    let status = mondayStatusMap[statusCru] || statusCru;
    const statusUpper = String(status).trim().toUpperCase();
    if (['CONCLUÍDO', 'CONCLUIDO', 'DONE'].includes(statusUpper)) status = 'Feito';
    else if (['A FAZER', 'TO DO'].includes(statusUpper)) status = 'A fazer';
    else if (['IMPEDIMENTO', 'IMPEDIDO'].includes(statusUpper)) status = 'Impedimento';
    else if (['WORKING ON IT', 'FAZENDO', 'EM ANDAMENTO'].includes(statusUpper)) status = 'Fazendo';
    else if (['IMPLANTADO', 'IMPLANTAÇÃO', 'IMPLANTACAO'].includes(statusUpper)) status = 'Implantação';
    else if (['HOMOLOGAÇÃO', 'HOMOLOGACAO'].includes(statusUpper)) status = 'Homologação';
    else if (['FILA'].includes(statusUpper)) status = 'Fila';
    else if (['EM ANÁLISE', 'EM ANALISE'].includes(statusUpper)) status = 'Em Análise';
    else if (['CANCELADO'].includes(statusUpper)) status = 'Cancelado';
    else if (['ENCERRADO'].includes(statusUpper)) status = 'Feito';
    else status = 'A fazer';
    return status !== 'Cancelado';
  }).length;

  const percentualSLA = totalTicketsNaoCancelados > 0 ? (dentroSLACount / totalTicketsNaoCancelados) : 0;
  const tempoMedio = countResolvidos > 0 ? (totalTempoResolucao / countResolvidos).toFixed(2) : 0;

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

  // Tickets por Área
  resumoSheet.getCell('A9').value = 'Tickets por Área'; resumoSheet.getCell('A9').font = fonteNegrito;
  Object.entries(contagemArea)
    .sort((a, b) => b[1] - a[1])
    .forEach(([a, q], i) => {
      resumoSheet.getCell(`A${10 + i}`).value = a;
      resumoSheet.getCell(`B${10 + i}`).value = q;
    });

  // Distribuição por Tipo
  resumoSheet.getCell('D9').value = 'Distribuição por Tipo'; resumoSheet.getCell('D9').font = fonteNegrito;
  Object.entries(contagemTipo)
    .sort((a, b) => b[1] - a[1])
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

  const listaStatusOficiais = ['Feito', 'A fazer', 'Impedimento', 'Fazendo', 'Implantação', 'Homologação', 'Fila', 'Em Análise', 'Cancelado'];
  listaStatusOficiais.forEach((s, i) => {
    resumoSheet.getCell(`M${10 + i}`).value = s;
    resumoSheet.getCell(`N${10 + i}`).value = contagemStatus[s] || 0;
  });

  // Proteção Visual contra sobreposição: Injeta Semanas nas colunas P e Q
  resumoSheet.getCell('P9').value = 'Tickets por Semana'; resumoSheet.getCell('P9').font = fonteNegrito;
  resumoSheet.getCell('Q9').value = 'Volume'; resumoSheet.getCell('Q9').font = fonteNegrito;
  Object.entries(contagemSemanaFinalizadosMovidesk)
    .sort((a, b) => a[0].localeCompare(b[0]))
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



async function sincronizarMondaySprints(token, boardId, mondayTextColId, mondayStatusColId, mondayPriorityColId, mondayDescriptionColId, mondayTypeColId, dirPath) {
  console.log('[Monday] Iniciando sincronização completa de sprints...');

  const movideskToken = process.env.MOVIDESK_TOKEN;
  const mapaMovidesk = movideskToken
    ? await buscarTodosTicketsMovidesk(movideskToken)
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

    // Ignora os groups administrativos
    const nomeGrupoUpper = nomeGrupo.toUpperCase();
    if (['BACKLOG', 'PROXIMA SPRINT', 'PRÓXIMA SPRINT', 'SPRINT ATUAL'].includes(nomeGrupoUpper)) {
      continue; // Pula o item e nao gera excel
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
        null // Adicionado o 10º parâmetro (mondaySlaDateColId) para evitar desalinhamento de argumentos
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
    // Lê a pasta física utils para descobrir o histórico real de arquivos gerados
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
      // Garante que o select exiba as sprints mais novas primeiro por ordenação reversa
      .sort((a, b) => b.nome_exibicao.localeCompare(a.nome_exibicao));

    // Grava um arquivo JSON estático de índice contendo o array mapeado de planilhas
    const caminhoDoIndiceJson = path.join(dirPath, 'sprints.json');
    await fs.writeFile(caminhoDoIndiceJson, JSON.stringify(listaCompletaSprints, null, 2), 'utf8');
    console.log(`[Altona Engine] Índice de arquivos estáticos criado em: utils/sprints.json`);

  } catch (errDiscorv) {
    console.error('[Altona Engine] Erro ao tentar mapear arquivos estáticos da pasta utils:', errDiscorv.message);
  }

  return resultados;
}

module.exports = { sincronizarMondaySprints };