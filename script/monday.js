const axios = require('axios');
const ExcelJS = require('exceljs');
const path = require('path');
const https = require('https');
const fs = require('fs').promises;

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
  53857: "Alta",
  53858: "Média",
  53859: "Baixa",
  53860: "Muito Alta"
};

const httpsAgent = new https.Agent({ rejectUnauthorized: false });

// SLA baseado no ciclo da sprint
const SLA_DIAS = 15;

// Status que indicam ticket encerrado
const STATUS_ENCERRADO = ['Feito', 'Implantação', 'Concluído', 'Encerrado'];

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

function calcularSLA(dataEncerramentoObj, statusFinal, dataFimSprint) {
  const encerrado = STATUS_ENCERRADO.includes(statusFinal);

  if (!dataFimSprint) {
    return { dentroSLA: 'Aberto', atraso: 0 };
  }

  // Ticket ainda aberto (A fazer, Fazendo, Impedido)
  if (!encerrado) {
    const hoje = new Date();
    // Se a data atual já passou do dia de fechamento daquela Sprint cai como nao
    if (hoje > dataFimSprint) {
      const diasAtraso = Math.ceil((hoje - dataFimSprint) / (1000 * 24 * 60 * 60));
      return { dentroSLA: 'Não', atraso: diasAtraso };
    }
    return { dentroSLA: 'Aberto', atraso: 0 };
  }

  //Ticket encerrado mas sem data válida de fechamento
  if (!dataEncerramentoObj || isNaN(dataEncerramentoObj.getTime())) {
    return { dentroSLA: 'Não', atraso: 0 };
  }

  //Entregou até o último dia planejado daquela Sprint
  if (dataEncerramentoObj <= dataFimSprint) {
    return { dentroSLA: 'Sim', atraso: 0 };
  } else {
    // Se entregou depois que o ciclo fechou, calcula o atraso em relação ao fim da sprint
    const atraso = Math.ceil((dataEncerramentoObj - dataFimSprint) / (1000 * 24 * 60 * 60));
    return { dentroSLA: 'Não', atraso: atraso };
  }
}


async function buscarValoresCamposPersonalizadosPessoas(token) {
  const mapaValoresCamposPessoas = {};
  let skip = 0;
  const take = 100;
  const maxRetries = 3;
  const retryDelayMs = 2000;

  console.log('[Movidesk] Iniciando varredura de valores de campos personalizados de pessoas...');

  while (true) {
    let resp;
    let retries = 0;
    let success = false;
    let falhaCritica = false;

    while (retries < maxRetries && !success) {
      try {
        resp = await axios.get('https://api.movidesk.com/public/v1/persons', {
          params: {
            token,
            '$select': 'id',
            '$expand': 'customFieldValues($expand=items)',
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

  console.log(`[Movidesk] Concluído. ${Object.keys(mapaValoresCamposPessoas).length} pessoas mapeadas.`);
  return mapaValoresCamposPessoas;
}

async function buscarEstruturaCamposPersonalizadosClientes(token) {
  const mapaCamposGlobais = {};
  let skip = 0;
  const take = 100;
  const maxRetries = 3;
  const retryDelayMs = 2000;

  console.log('[Movidesk] Iniciando varredura de estrutura de campos personalizados de clientes...');

  try {
    while (true) {
      let resp;
      let retries = 0;
      let success = false;
      let falhaCritica = false;

      while (retries < maxRetries && !success) {
        try {
          resp = await axios.get('https://api.movidesk.com/public/v1/customFields', {
            params: {
              token,
              '$filter': `entityType eq 'Person'`,
              '$top': take,
              '$skip': skip,
            },
            httpsAgent,
            timeout: 90000,
          });
          success = true;
        } catch (err) {
          retries++;
          if (retries < maxRetries) {
            await new Promise(resolve => setTimeout(resolve, retryDelayMs));
          } else {
            falhaCritica = true;
          }
        }
      }

      if (falhaCritica || !success) break;

      const campos = resp.data.value || (Array.isArray(resp.data) ? resp.data : []);
      if (campos.length === 0) break;

      for (const campo of campos) {
        const idCampo = String(campo.id);
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

    console.log(`[Movidesk] Concluído. ${relatorioFinal.length} estruturas de campos identificadas.`);
    return relatorioFinal;

  } catch (outerErr) {
    console.error(`[Movidesk] Erro crítico ao buscar estrutura de campos: ${outerErr.message}`);
    return [];
  }
}

async function buscarTodosTicketsMovidesk(token, personCustomFieldValuesMap, movideskClientSectorCustomFieldId) {
  const mapa = {};
  let skip = 0;
  const take = 100;
  let total = 0;
  const maxRetries = 3;
  const retryDelayMs = 2000;

  console.log('[Movidesk] Iniciando busca paginada de tickets...');

  while (true) {
    let resp;
    let retries = 0;
    let success = false;
    let falhaCritica = false;

    while (retries < maxRetries && !success) {
      try {
        resp = await axios.get('https://api.movidesk.com/public/v1/tickets', {
          params: {
            token,
            '$select': 'id,subject,category,urgency,status,createdDate,resolvedIn,ownerTeam,serviceFirstLevel,slaSolutionDate,createdBy,clients',
            // 🔥 CORREÇÃO 1: Adicionado ',clients' no expand para garantir que a API retorne os dados se o criador vier nulo
            '$expand': 'actions($select=description),createdBy,clients',
            // Removido o filtro de data para buscar todos os tickets, independentemente da data de criação
            // '$filter': 'createdDate gt 2025-01-01T00:00:00Z',
            '$top': take,
            '$skip': skip,
          },
          httpsAgent,
          timeout: 90000,
        });
        success = true;
      } catch (err) {
        retries++;
        if (retries < maxRetries) {
          await new Promise(resolve => setTimeout(resolve, retryDelayMs));
        } else {
          falhaCritica = true;
        }
      }
    }

    if (falhaCritica || !success) break;

    const lista = resp.data && resp.data.value ? resp.data.value : (Array.isArray(resp.data) ? resp.data : []);
    if (lista.length === 0) break;

    for (const t of lista) {
      // 🔥 Força o ID do ticket do Movidesk a virar String de forma limpa e sem espaços
      if (!t.id) continue;
      const idTicket = String(t.id).trim();
      
      const statusTicket = t.status || 'N/A';
      const equipeBruta = t.ownerTeam || 'N/A';

      const equipe = String(equipeBruta).trim();
      if (equipe === 'Suporte e Infraestrutura' || equipe === 'Administradores') continue;

      const prioridadeMovidesk = urgencyMap[t.urgency] || 'N/A';
      const categoria = t.category || 'N/A';
      const tipo = t.serviceFirstLevel ? String(t.serviceFirstLevel).trim() : 'N/A';

      const createdDate = t.createdDate ? new Date(t.createdDate) : null;
      const closedDate = t.resolvedIn ? new Date(t.resolvedIn) : null;

      const acoes = t.actions || [];
      const descRaw = acoes[0]?.description || ''; // Acessa o primeiro item do array de ações
      const descricao = String(descRaw)
        .replace(/<[^>]*>?/gm, '')
        .replace(/\[cid:[^\]]*\]/g, '')
        .replace(/[\r\n\t]+/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
        .substring(0, 300) || 'N/A';

      // =========================================================================
      // 🔥 CORREÇÃO 2: LÓGICA DE PRIORIDADE DE PESSOA (createdBy -> clients)
      //    Trata usuários inativos e múltiplos clientes
      // =========================================================================
      let solicitante = 'Solicitante não informado';
      let pessoaIdSetor = null;

      const criadorObjeto = t.createdBy || {};
      const listaClientes = Array.isArray(t.clients) ? t.clients : [];

      // Cenário A: O criador existe e está ativo (possuindo nome válido)
      if (criadorObjeto.personName || criadorObjeto.name || criadorObjeto.businessName) {
        solicitante = criadorObjeto.personName || criadorObjeto.name || criadorObjeto.businessName;
        pessoaIdSetor = criadorObjeto.id ? String(criadorObjeto.id) : null;
      } 
      // Cenário B: Criador nulo/inativo, mas temos uma lista de clientes (ex: as 5 pessoas)
      else if (listaClientes.length > 0) {
        // Se houver mais de um cliente, junta os nomes separados por vírgula
        const nomesClientes = listaClientes
          .map(c => c.personName || c.name || c.businessName)
          .filter(nome => nome) // Remove nulos
          .join(', ');

        solicitante = nomesClientes || 'Múltiplos Clientes';
        
        // Pega o ID do primeiro cliente do array para mapear o setor corporativo
        pessoaIdSetor = listaClientes[0]?.id ? String(listaClientes[0].id) : null;
      }

      // Caso especial: Usuário inativo/desabilitado que perdeu o vínculo de nome, mas o ID ainda existe
      if (solicitante === 'Solicitante não informado' && criadorObjeto.id) {
          solicitante = `Usuário Desativado (ID: ${criadorObjeto.id})`;
          pessoaIdSetor = String(criadorObjeto.id);
      }

      // 2. Busca do Setor no seu mapa de pessoas de forma segura
      let setorCliente = 'Setor não informado';
      if (pessoaIdSetor && personCustomFieldValuesMap && movideskClientSectorCustomFieldId) {
        const camposPessoa = personCustomFieldValuesMap[pessoaIdSetor];
        if (camposPessoa && camposPessoa[String(movideskClientSectorCustomFieldId)] !== undefined && camposPessoa[String(movideskClientSectorCustomFieldId)] !== null) {
          setorCliente = String(camposPessoa[String(movideskClientSectorCustomFieldId)]).trim();
        }
      }
      // =========================================================================

      mapa[idTicket] = {
        area: setorCliente,
        tipo,
        categoria,
        descricao,
        prioridade: prioridadeMovidesk,
        solicitante: String(solicitante).trim(),
        equipe,
        abertura: createdDate ? createdDate.toISOString() : null,
        encerramento: closedDate ? closedDate.toISOString() : null,
        status: statusTicket
      };

      total++;
    }

    console.log(`[Movidesk] ${total} tickets processados (skip=${skip})...`);
    skip += take;

    // Condição para encerrar a paginação real
    if (lista.length < take) break;
    if (total > 5000) break;
  }

  console.log(`[Movidesk] Total final: ${total} tickets carregados.`);
  return mapa;
}


async function buscarTodosItensMonday(boardId, token) {
  let allItems = [];
  let cursor = null;
  let hasNextPage = true;

  console.log(`[Monday] Buscando itens do board ${boardId}...`);

  while (hasNextPage) {
    // created_at adicionado para usar como data de entrada na sprint (base do SLA)
    const query = `
      query {
        boards(ids: ${boardId}) {
          items_page(limit: 500, cursor: ${cursor ? `"${cursor}"` : "null"}) {
            cursor
            items {
              id
              name
              created_at
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

    const response = await axios.post(
      'https://api.monday.com/v2',
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

  console.log(`[Monday] Total de ${allItems.length} itens encontrados.`);
  return allItems;
}

const getColVal = (item, id) => {
  const c = item.column_values?.find(cv => cv.id === id);
  return c?.text?.trim() || '';
};

async function gerarExcelParaSprint(
  sprintNome, itensDaSprint, mapaMovidesk, dirPath,
  mondayTextColId, mondayStatusColId, mondayPriorityColId,
  mondayDescriptionColId, mondayTypeColId,
  mondayCreationDateColId, mondayResolutionDateColId,
  listaEstruturaCamposClientes
) {
  const workbook = new ExcelJS.Workbook();
  const estiloAzul = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF034C8C' } };
  const fonteBranca = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
  const fonteNegrito = { bold: true };

  const totalTicketsDaSprint = itensDaSprint.length;

  const ticketsSheet = workbook.addWorksheet('Tickets');
  const dashSheet = workbook.addWorksheet('Dashboard');
  const resumoSheet = workbook.addWorksheet('Resumo');

  dashSheet.views = [{ showGridLines: true }];
  resumoSheet.views = [{ showGridLines: true }];

  let dataInicioSprint = null;
  for (const item of itensDaSprint) {
    if (item.created_at) {
      const dataItem = new Date(item.created_at);
      if (!isNaN(dataItem.getTime())) {
        if (!dataInicioSprint || dataItem < dataInicioSprint) {
          dataInicioSprint = dataItem;
        }
      }
    }
  }

  if (!dataInicioSprint) dataInicioSprint = new Date();
  dataInicioSprint.setHours(0, 0, 0, 0);

  const dataFimSprint = new Date(dataInicioSprint.getTime());
  dataFimSprint.setDate(dataInicioSprint.getDate() + 13);
  dataFimSprint.setHours(23, 59, 59, 999);

  const strInicio = dataInicioSprint.toLocaleDateString('pt-BR');
  const strFim = dataFimSprint.toLocaleDateString('pt-BR');
  const periodoSprintTexto = `${strInicio} a ${strFim}`;

  let dentroSLACount = 0;
  let foraSLACount = 0;
  let totalTempoResolucao = 0;
  let countResolvidos = 0;
  const contagemArea = {};
  const contagemTipo = {};
  const contagemPrio = {};
  const contagemStatus = {};
  const contagemSemanaFinalizados = {};
  const ticketsForaSLA = [];

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
    { header: 'Abertura', key: 'abertura', width: 16 },
    { header: 'Encerramento', key: 'encerramento', width: 14 },
    { header: 'Status', key: 'status', width: 16 },
    { header: 'Dentro da SLA', key: 'sla', width: 14 },
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
    try {
      // 🔥 CORREÇÃO PRINCIPAL: Garante que o ID do Monday seja uma String limpa
      const mondayMovideskIdRaw = typeof getColVal === 'function' ? String(getColVal(item, mondayTextColId)).trim() : null;
      
      // 🔥 NOVA LÓGICA: 'mov' agora só terá dados se o ID do Monday *existir como chave* no mapa do Movidesk
      // Isso garante que itens Monday-only (ou IDs inválidos) resultem em 'mov' vazio.
      const mov = (mondayMovideskIdRaw && mapaMovidesk[mondayMovideskIdRaw]) ? mapaMovidesk[mondayMovideskIdRaw] : {};

      // Determina se este item do Monday *realmente* tem dados correspondentes no Movidesk
      const temDadosMovidesk = Object.keys(mov).length > 0;

      const prioMon = typeof getColVal === 'function' ? (mondayPriorityMap[getColVal(item, mondayPriorityColId)] || 'Não definida') : 'Não definida';
      const typeMon = typeof getColVal === 'function' ? getColVal(item, mondayTypeColId) : null;
      const descMon = typeof getColVal === 'function' ? getColVal(item, mondayDescriptionColId) : null;

      // =========================================================================
      // AJUSTE: Solicitante, Área e Equipe vêm SOMENTE do Movidesk.
      // Se não houver dados do Movidesk, usam o fallback "Não informado".
      // =========================================================================
      // O ticketId será o ID do Movidesk (se temDadosMovidesk for true) ou o ID do Monday
      const ticketId = temDadosMovidesk ? mondayMovideskIdRaw : item.id;
      
      // Área: Estritamente do Movidesk, ou "Setor não informado"
      const area = mov.area || 'Setor não informado';
      
      // Tipo: Pode vir do Monday se não houver Movidesk, ou "Tipo não informado"
      const tipo = typeMon || mov.tipo || 'Tipo não informado';
      
      // Categoria: Estritamente do Movidesk, ou "Categoria não informada"
      const categoria = mov.categoria || 'Categoria não informada';
      
      // Título: Pode vir do Monday se não houver Movidesk, ou "Sem Título"
      const titulo = item.name || 'Sem Título';
      
      // Descrição: Pode vir do Monday se não houver Movidesk, ou "Descrição não informada"
      const descricao = descMon || mov.descricao || item.name || 'Descrição não informada';
      
      // Prioridade: Vem do Monday, ou "Não definida"
      const prioridade = prioMon; 

      // Solicitante: Estritamente do Movidesk, ou "Solicitante não informado"
      const solicitante = mov.solicitante || 'Solicitante não informado';
      
      // Equipe: Estritamente do Movidesk, ou "Equipe não informada"
      const equipe = mov.equipe || 'Equipe não informada';
      // =========================================================================

      let dataEntradaSprint = item.created_at ? new Date(item.created_at) : null;
      if (!dataEntradaSprint && typeof getColVal === 'function') {
        const dataMondayCriacao = getColVal(item, mondayCreationDateColId);
        if (dataMondayCriacao && !isNaN(new Date(dataMondayCriacao).getTime())) {
          dataEntradaSprint = new Date(dataMondayCriacao);
        }
      }

      let dataEncerramentoObj = mov.encerramento ? new Date(mov.encerramento) : null;
      if (!dataEncerramentoObj && typeof getColVal === 'function') {
        const dataMondayResolucao = getColVal(item, mondayResolutionDateColId);
        if (dataMondayResolucao && !isNaN(new Date(dataMondayResolucao).getTime())) {
          dataEncerramentoObj = new Date(dataMondayResolucao);
        }
      }

      const aberturaFormatada = dataEntradaSprint ? dataEntradaSprint.toLocaleDateString('pt-BR') : 'N/A';
      const encerramentoFormatado = dataEncerramentoObj ? dataEncerramentoObj.toLocaleDateString('pt-BR') : 'Não';

      const statusCru = typeof getColVal === 'function' ? (getColVal(item, mondayStatusColId) || 'A fazer') : 'A fazer';
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

      const semanaLinha = dataEntradaSprint ? calcularSemana(dataEntradaSprint) : 'Semana 1';
      const mesLinha = dataEntradaSprint ? nomeMes(dataEntradaSprint) : 'N/A';

      contagemArea[area] = (contagemArea[area] || 0) + 1;
      contagemTipo[tipo] = (contagemTipo[tipo] || 0) + 1;
      contagemPrio[prioridade] = (contagemPrio[prioridade] || 0) + 1;
      contagemStatus[statusFinal] = (contagemStatus[statusFinal] || 0) + 1;

      if (['Feito', 'Implantação'].includes(statusFinal)) {
        const dataRef = dataEncerramentoObj || dataEntradaSprint;
        if (dataRef) {
          const sem = calcularSemana(dataRef);
          contagemSemanaFinalizados[sem] = (contagemSemanaFinalizados[sem] || 0) + 1;
        }
      }

      const { dentroSLA, atraso } = calcularSLA(dataEncerramentoObj, statusFinal, dataFimSprint);

      if (dentroSLA === 'Sim') {
        dentroSLACount++;
      } else if (dentroSLA === 'Não') {
        foraSLACount++;
        ticketsForaSLA.push({ ticket: ticketId, area, titulo, prioridade, atraso });
      }

      let tempoResolucao = mov.tempoResolucao ?? 'N/A';
      if (['Feito', 'Implantação'].includes(statusFinal)) {
        if (typeof tempoResolucao === 'number') {
          totalTempoResolucao += tempoResolucao;
          countResolvidos++;
        } else if (dataEntradaSprint && dataEncerramentoObj) {
          const dias = Math.ceil((dataEncerramentoObj - dataEntradaSprint) / 86400000);
          totalTempoResolucao += dias;
          countResolvidos++;
        }
      }

      const row = ticketsSheet.addRow({
        ticket: ticketId, area, tipo, categoria, titulo,
        descricao, prioridade, solicitante, equipe,
        abertura: aberturaFormatada,
        encerramento: encerramentoFormatado,
        status: statusFinal,
        sla: dentroSLA,
        semana: semanaLinha,
        mes: mesLinha,
      });
      row.eachCell(cell => { cell.alignment = { vertical: 'middle', wrapText: true }; });
    } catch (errItem) {
      console.warn(`[Monday] Ignorando linha com erro na sprint ${sprintNome}:`, errItem.message);
    }
  }

  const totalEncerrados = dentroSLACount + foraSLACount;
  const percentualSLA = totalEncerrados > 0 ? (dentroSLACount / totalEncerrados) : 0;
  const tempoMedio = countResolvidos > 0 ? (totalTempoResolucao / countResolvidos).toFixed(2) : 0;

  // ============================================================
  // ABA: CAMPOS PERSONALIZADOS CLIENTES (Opcional)
  // ============================================================
  if (listaEstruturaCamposClientes && listaEstruturaCamposClientes.length > 0) {
    const camposClientesSheet = workbook.addWorksheet('Campos Personalizados Clientes');
    camposClientesSheet.columns = [
      { header: 'ID do Campo', key: 'customFieldId', width: 15 },
      { header: 'ID da Regra', key: 'customFieldRuleId', width: 15 },
      { header: 'Tipo de Campo', key: 'type', width: 18 },
      { header: 'ID do Item Opção', key: 'customFieldItemId', width: 18 },
      { header: 'Nome/Valor da Opção', key: 'customFieldItem', width: 40 }
    ];
    camposClientesSheet.getRow(1).eachCell(cell => {
      cell.fill = estiloAzul; cell.font = fonteBranca; cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });
    for (const campo of listaEstruturaCamposClientes) {
      if (!campo.items || campo.items.length === 0) {
        camposClientesSheet.addRow({ customFieldId: campo.customFieldId, customFieldRuleId: campo.customFieldRuleId, type: campo.type || 'N/A', customFieldItemId: 'N/A', customFieldItem: '(Livre)' });
      } else {
        for (const item of campo.items) {
          camposClientesSheet.addRow({ customFieldId: campo.customFieldId, customFieldRuleId: campo.customFieldRuleId, type: campo.type || 'N/A', customFieldItemId: item.customFieldItemId, customFieldItem: item.customFieldItem });
        }
      }
    }
  }

  // ============================================================
  // MONTAGEM SEGURO DA ABA 1: DASHBOARD
  // ============================================================
  dashSheet.mergeCells('A1:C1');
  dashSheet.getCell('A1').value = 'Dashboard de Change Log - Governança de TI';
  dashSheet.getCell('A1').font = { bold: true, size: 16, color: { argb: 'FF034C8C' } };

  dashSheet.getCell('A2').value = 'Período';
  dashSheet.getCell('B2').value = periodoSprintTexto;
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
  dashSheet.getCell('G5').value = parseFloat(tempoMedio || 0);
  dashSheet.getCell('G5').font = fonteNegrito;

  dashSheet.getCell('J4').value = 'Áreas atendidas';
  dashSheet.getCell('J5').value = Object.keys(contagemArea).length;
  dashSheet.getCell('J5').font = fonteNegrito;

  ['A4', 'D4', 'G4', 'J4'].forEach(pos => {
    dashSheet.getCell(pos).fill = estiloAzul; dashSheet.getCell(pos).font = fonteBranca; dashSheet.getCell(pos).alignment = { horizontal: 'center', vertical: 'middle' };
  });
  ['A5', 'D5', 'G5', 'J5'].forEach(pos => {
    dashSheet.getCell(pos).alignment = { horizontal: 'center', vertical: 'middle' };
    dashSheet.getCell(pos).border = { bottom: { style: 'thin' }, top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
  });

  dashSheet.getCell('J21').value = 'Status final';
  dashSheet.getCell('J21').font = fonteNegrito;
  ['Feito', 'Implantação', 'Encerrado'].forEach((s, i) => {
    dashSheet.getCell(`J${22 + i}`).value = s;
    dashSheet.getCell(`K${22 + i}`).value = contagemStatus[s] || 0;
  });

  // ============================================================
  // MONTAGEM SEGURO DA ABA 2: RESUMO
  // ============================================================
  resumoSheet.getCell('A1').value = 'Resumo Executivo';
  resumoSheet.getCell('A1').font = fonteNegrito;
  resumoSheet.getCell('A2').value = 'Total de tickets';
  resumoSheet.getCell('B2').value = totalTicketsDaSprint;
  resumoSheet.getCell('A5').value = percentualSLA;
  resumoSheet.getCell('B5').numFmt = '0.00%';
  resumoSheet.getCell('A6').value = parseFloat(tempoMedio || 0);
  resumoSheet.getCell('A7').value = Object.keys(contagemArea).length;

  resumoSheet.getCell('A9').value = 'Tickets por Área';

  resumoSheet.getCell('A9').font = fonteNegrito;
  Object.entries(contagemArea).sort((a, b) => b[1] - a[1]).forEach(([a, q], i) => {
    resumoSheet.getCell(`A${10 + i}`).value = a;
    resumoSheet.getCell(`B${10 + i}`).value = q;
  });

  resumoSheet.getCell('D9').value = 'Distribuição por Tipo';
  resumoSheet.getCell('D9').font = fonteNegrito;
  Object.entries(contagemTipo).sort((a, b) => b[1] - a[1]).forEach(([t, q], i) => {
    resumoSheet.getCell(`D${10 + i}`).value = t;
    resumoSheet.getCell(`E${10 + i}`).value = q;
  });

  resumoSheet.getCell('G9').value = 'Prioridade';
  resumoSheet.getCell('G9').font = fonteNegrito;
  ['Muito Alta', 'Alta', 'Média', 'Baixa', 'Não definida'].forEach((p, i) => {
    resumoSheet.getCell(`G${10 + i}`).value = p;
    resumoSheet.getCell(`H${10 + i}`).value = contagemPrio[p] || 0;
  });

  resumoSheet.getCell('J9').value = 'Cumprimento SLA (15 dias)';
  resumoSheet.getCell('J9').font = fonteNegrito;
  resumoSheet.getCell('J10').value = 'Sim'; resumoSheet.getCell('K10').value = dentroSLACount;
  resumoSheet.getCell('J11').value = 'Não'; resumoSheet.getCell('K11').value = foraSLACount;
  resumoSheet.getCell('J12').value = 'Abertos'; resumoSheet.getCell('K12').value = totalTicketsDaSprint - totalEncerrados;

  resumoSheet.getCell('M9').value = 'Status Final';
  resumoSheet.getCell('M9').font = fonteNegrito;
  ['Feito', 'A fazer', 'Impedimento', 'Fazendo', 'Implantação', 'Homologação', 'Fila', 'Em Análise'].forEach((s, i) => {
    resumoSheet.getCell(`M${10 + i}`).value = s;
    resumoSheet.getCell(`N${10 + i}`).value = contagemStatus[s] || 0;
  });

  resumoSheet.getCell('P9').value = 'Tickets por Semana';
  resumoSheet.getCell('P9').font = fonteNegrito;
  resumoSheet.getCell('Q9').value = 'Volume';
  resumoSheet.getCell('Q9').font = fonteNegrito;
  Object.entries(contagemSemanaFinalizados).sort((a, b) => a[0].localeCompare(b[0])).forEach(([semana, qtd], i) => {
    resumoSheet.getCell(`P${10 + i}`).value = semana;
    resumoSheet.getCell(`Q${10 + i}`).value = qtd;
  });

  resumoSheet.getCell('D19').value = 'Top tickets fora do SLA';
  resumoSheet.getCell('D19').font = fonteNegrito;
  ['D20', 'E20', 'F20', 'G20'].forEach((c, i) => {
    const cell = resumoSheet.getCell(c);
    cell.value = ['Ticket', 'Área', 'Atraso (dias)', 'Título'][i];
    cell.fill = estiloAzul;
    cell.font = fonteBranca;
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  });

  ticketsForaSLA
    .filter(t => typeof t.atraso === 'number')
    .sort((a, b) => b.atraso - a.atraso)
    .slice(0, 10)
    .forEach((t, i) => {
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

  // ============================================================
  // GRAVAÇÃO DO ARQUIVO FÍSICO COM TODAS AS ABAS GARANTIDAS
  // ============================================================
  const nomeArquivo = `monday_sprint_${sprintNome.toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, '_').replace(/[^a-z0-9_\-]/g, '')}.xlsx`;

  await fs.mkdir(dirPath, { recursive: true });
  await workbook.xlsx.writeFile(path.join(dirPath, nomeArquivo));
  console.log(`[Monday] Excel gerado com todas as abas obrigatórias: ${nomeArquivo}`);

  return {
    nome_exibicao: `MONDAY - ${sprintNome.toUpperCase()}`,
    caminho_arquivo: `utils/${nomeArquivo}`
  };
}

async function sincronizarMondaySprints(
  token, boardId,
  mondayTextColId, mondayStatusColId, mondayPriorityColId,
  mondayDescriptionColId, mondayTypeColId,
  mondayCreationDateColId, mondayResolutionDateColId,
  movideskClientSectorCustomFieldId, dirPath
) {
  console.log('[Monday] Iniciando sincronização completa de sprints...');

  const movideskToken = process.env.MOVIDESK_TOKEN;

  const personCustomFieldValuesMap = movideskToken
    ? await buscarValoresCamposPersonalizadosPessoas(movideskToken)
    : {};

  const listaEstruturaCamposClientes = movideskToken
    ? await buscarEstruturaCamposPersonalizadosClientes(movideskToken)
    : [];

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
  const gruposIgnorados = ['BACKLOG', 'PROXIMA SPRINT', 'PRÓXIMA SPRINT', 'SPRINT ATUAL'];

  for (const item of todosItensMonday) {
    const nomeGrupo = item.group?.title ? item.group.title.trim() : 'Sem Sprint';
    if (gruposIgnorados.includes(nomeGrupo.toUpperCase())) continue;
    if (!sprints[nomeGrupo]) sprints[nomeGrupo] = [];
    sprints[nomeGrupo].push(item);
  }

  console.log(`[Monday] ${Object.keys(sprints).length} sprints mapeadas.`);

  const resultados = [];
  for (const [sprintNome, itensDaSprint] of Object.entries(sprints)) {
    try {
      console.log(`[Monday] Gerando Excel para: "${sprintNome}" (${itensDaSprint.length} itens).`);
      const res = await gerarExcelParaSprint(
        sprintNome, itensDaSprint, mapaMovidesk, dirPath,
        mondayTextColId, mondayStatusColId, mondayPriorityColId,
        mondayDescriptionColId, mondayTypeColId,
        mondayCreationDateColId, mondayResolutionDateColId,
        listaEstruturaCamposClientes
      );
      if (res) resultados.push(res);
    } catch (err) {
      console.error(`[Monday] Falha na sprint "${sprintNome}":`, err.message);
    }
  }

  try {
    await fs.mkdir(dirPath, { recursive: true });
    const arquivosNaPasta = await fs.readdir(dirPath);
    const planilhas = arquivosNaPasta.filter(f => f.endsWith('.xlsx'));

    if (planilhas.length === 0 && resultados.length > 0) {
      console.warn(' Nenhuma planilha encontrada no disco, usando dados da memória...');
      await fs.writeFile(
        path.join(dirPath, 'sprints.json'),
        JSON.stringify(resultados, null, 2),
        'utf8'
      );
      return resultados;
    }

    const listaCompletaSprints = planilhas
      .map(f => {
        let nomeExibicao = f.replace('.xlsx', '').toUpperCase().replace(/_/g, ' ');
        if (nomeExibicao.startsWith('TICKETS SPRINT '))
          nomeExibicao = `MOVIDESK - ${nomeExibicao.replace('TICKETS SPRINT ', '')}`;
        else if (nomeExibicao.startsWith('MONDAY SPRINT '))
          nomeExibicao = `MONDAY - ${nomeExibicao.replace('MONDAY SPRINT ', '')}`;
        return { nome_exibicao: nomeExibicao, caminho_arquivo: `utils/${f}` };
      })
      .sort((a, b) => b.nome_exibicao.localeCompare(a.nome_exibicao));

    if (listaCompletaSprints.length > 0) {
      await fs.writeFile(
        path.join(dirPath, 'sprints.json'),
        JSON.stringify(listaCompletaSprints, null, 2),
        'utf8'
      );
      console.log('[Altona] Índice estático utils/sprints.json atualizado.');
    } else {
      console.error(' Abortando atualização do sprints.json: Nenhuma sprint válida foi gerada.');
    }
  } catch (errDiscovery) {
    console.error('[Altona] Erro ao gerar sprints.json:', errDiscovery.message);
  }

  return resultados;
}

module.exports = { sincronizarMondaySprints, buscarEstruturaCamposPersonalizadosClientes };