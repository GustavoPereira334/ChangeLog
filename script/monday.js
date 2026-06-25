const axios = require('axios');
const ExcelJS = require('exceljs');
const path = require('path');
const https = require('https');
const crypto = require('crypto');
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
const MOVIDESK_TICKETS_ENDPOINTS = [
  'https://api.movidesk.com/public/v1/tickets',
  'https://api.movidesk.com/public/v1/tickets/merged',
  'https://api.movidesk.com/public/v1/tickets/past',
];

const MOVIDESK_TICKET_SELECT = 'id,subject,category,urgency,status,createdDate,resolvedIn,ownerTeam,serviceFirstLevel,slaSolutionDate,createdBy,clients';
const MOVIDESK_TICKET_EXPAND = 'actions($select=description),createdBy,clients';

function normalizarTextoComparacao(value) {
  return String(value || '')
    .trim()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase();
}

function valorVazioOuGenerico(value) {
  const texto = normalizarTextoComparacao(value);
  return [
    '',
    'N/A',
    'NA',
    'NAO',
    'NAO INFORMADO',
    'NAO INFORMADA',
    'SETOR NAO INFORMADO',
    'TIPO NAO INFORMADO',
    'CATEGORIA NAO INFORMADA',
    'DESCRICAO NAO INFORMADA',
    'SOLICITANTE NAO DEFINIDO',
    'EQUIPE NAO DEFINIDA',
  ].includes(texto);
}

function ticketPrecisaComplemento(ticket) {
  if (!ticket || ticket.ignorado) return true;
  return ['area', 'tipo', 'categoria', 'descricao', 'solicitante', 'equipe', 'abertura', 'encerramento', 'status']
    .some(key => valorVazioOuGenerico(ticket[key]));
}

function mesclarCampoPreferindoValor(baseValue, incomingValue) {
  if (valorVazioOuGenerico(baseValue) && !valorVazioOuGenerico(incomingValue)) {
    return incomingValue;
  }
  return baseValue;
}

function mesclarDadosTicket(base = {}, incoming = {}) {
  if (!incoming || incoming.ignorado) return base || {};
  const merged = { ...(base && !base.ignorado ? base : {}) };
  for (const key of ['titulo', 'area', 'tipo', 'categoria', 'descricao', 'prioridade', 'solicitante', 'equipe', 'abertura', 'encerramento', 'status']) {
    merged[key] = mesclarCampoPreferindoValor(merged[key], incoming[key]);
  }
  return merged;
}

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

  if (!encerrado) {
    const hoje = new Date();
    if (hoje > dataFimSprint) {
      const diasAtraso = Math.ceil((hoje - dataFimSprint) / (1000 * 24 * 60 * 60));
      return { dentroSLA: 'Não', atraso: diasAtraso };
    }
    return { dentroSLA: 'Aberto', atraso: 0 };
  }

  if (!dataEncerramentoObj || isNaN(dataEncerramentoObj.getTime())) {
    return { dentroSLA: 'Não', atraso: 0 };
  }

  if (dataEncerramentoObj <= dataFimSprint) {
    return { dentroSLA: 'Sim', atraso: 0 };
  } else {
    const atraso = Math.ceil((dataEncerramentoObj - dataFimSprint) / (1000 * 24 * 60 * 60));
    return { dentroSLA: 'Não', atraso: atraso };
  }
}

function limparTextoEncodingBug(texto) {
  if (!texto) return 'N/A';
  return texto
    .replace(/Ã§/g, 'ç')
    .replace(/Ã£/g, 'ã')
    .replace(/Ã³/g, 'ó')
    .replace(/Ã¡/g, 'á')
    .replace(/Ãª/g, 'ê')
    .replace(/Ã/g, 'a') // fallback 
    .replace(/Â/g, ''); // remove regex extra
}


// Função para calcular a assinatura digital de uma sprint para detectar mudanças
function calcularFingerprintSprint(itensDaSprint, mapaMovidesk, mondayTextColId, mondayStatusColId, mondayPriorityColId) {
  const dataToHash = itensDaSprint.map(item => {
    const idRaw = String(getColVal(item, mondayTextColId)).trim().replace(/\D/g, '');
    const mov = idRaw ? mapaMovidesk[idRaw] : {};
    return {
      id: item.id,
      name: item.name,
      updated_at: item.updated_at || '', // Data de modificação do Monday
      status: getColVal(item, mondayStatusColId),
      priority: getColVal(item, mondayPriorityColId),
      mov_status: mov?.status || '',
      mov_encerramento: mov?.encerramento || ''
    };
  });
  return crypto.createHash('sha256').update(JSON.stringify(dataToHash)).digest('hex');
}

async function buscarValoresCamposPersonalizadosPessoas(token) {
  const mapaValoresCamposPessoas = {};
  let skip = 0;
  const take = 300;
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
  const take = 300;
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

// paginação
async function buscarTodosTicketsMovidesk(token, personCustomFieldValuesMap, movideskClientSectorCustomFieldId) {
  const mapa = {};
  let skip = 0;
  const take = 300;
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
            '$expand': 'actions($select=description),createdBy,clients',
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
      if (!t.id) continue;
      const idTicket = String(t.id).trim();
      const statusTicket = t.status || 'N/A';
      const equipeBruta = t.ownerTeam || 'N/A';

      const equipe = t.ownerTeam && t.ownerTeam.trim() !== ''
        ? String(t.ownerTeam).trim()
        : (t.serviceFirstLevel || 'Equipe não definida');

      if (equipe === 'Suporte e Infraestrutura' || equipe === 'Administradores') {
        mapa[idTicket] = { ignorado: true, equipe };
        continue;
      }

      const prioridadeMovidesk = urgencyMap[t.urgency] || 'N/A';
      const categoria = t.category || 'N/A';
      const tipo = t.serviceFirstLevel ? String(t.serviceFirstLevel).trim() : 'N/A';

      const createdDate = t.createdDate ? new Date(t.createdDate) : null;
      const closedDate = t.resolvedIn ? new Date(t.resolvedIn) : null;

      const acoes = t.actions || [];
      const descRaw = acoes[0]?.description || '';
      const descricao = String(descRaw)
        .replace(/<[^>]*>?/gm, '')
        .replace(/\[cid:[^\]]*\]/g, '')
        .replace(/[\r\n\t]+/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
        .substring(0, 300) || 'N/A';

      let solicitante = 'Solicitante não definido';

      if (t.createdBy?.businessName) {
        solicitante = t.createdBy.businessName;
      }
      else if (t.createdBy?.personName || t.createdBy?.name) {
        solicitante = t.createdBy.personName || t.createdBy.name;
      }
      else if (t.clients && t.clients.length > 0) {
        solicitante = t.clients
          .map(c => c.businessName || c.personName || c.name)
          .filter(Boolean)
          .join(', ');
      }
      else if (t.createdBy?.id) {
        solicitante = `Usuário Desativado (ID: ${t.createdBy.id})`;
      }

      let pessoaIdSetor = null;
      const criadorObjeto = t.createdBy || {};
      const listaClientes = Array.isArray(t.clients) ? t.clients : [];

      if (criadorObjeto.personName || criadorObjeto.name || criadorObjeto.businessName) {
        solicitante = criadorObjeto.personName || criadorObjeto.name || criadorObjeto.businessName;
        pessoaIdSetor = criadorObjeto.id ? String(criadorObjeto.id) : null;
      }
      else if (listaClientes.length > 0) {
        const nomesClientes = listaClientes
          .map(c => c.personName || c.name || c.businessName)
          .filter(nome => nome)
          .join(', ');

        solicitante = nomesClientes || 'Múltiplos Clientes';
        pessoaIdSetor = listaClientes[0]?.id ? String(listaClientes[0].id) : null;
      }

      if (solicitante === 'Solicitante não informado' && criadorObjeto.id) {
        solicitante = `Usuário Desativado (ID: ${criadorObjeto.id})`;
        pessoaIdSetor = String(criadorObjeto.id);
      }

      let setorCliente = 'Setor não informado';
      if (pessoaIdSetor && personCustomFieldValuesMap && movideskClientSectorCustomFieldId) {
        const camposPessoa = personCustomFieldValuesMap[pessoaIdSetor];
        if (camposPessoa && camposPessoa[String(movideskClientSectorCustomFieldId)] !== undefined && camposPessoa[String(movideskClientSectorCustomFieldId)] !== null) {
          setorCliente = String(camposPessoa[String(movideskClientSectorCustomFieldId)]).trim();
        }
      }

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

    if (lista.length < take) break;
    if (total > 5000) break;
  }

  console.log(`[Movidesk] Total final: ${total} tickets carregados.`);
  return mapa;
}

// Função auxiliar para realizar a chamada HTTP de um lote específico
function normalizarTicketMovidesk(t, personCustomFieldValuesMap, movideskClientSectorCustomFieldId) {
  if (!t || !t.id) return null;

  const equipe = t.ownerTeam && t.ownerTeam.trim() !== ''
    ? String(t.ownerTeam).trim()
    : (t.serviceFirstLevel || 'Equipe não definida');

  if (equipe === 'Suporte e Infraestrutura' || equipe === 'Administradores') {
    return { id: String(t.id).trim(), dados: { ignorado: true, equipe } };
  }

  const acoes = t.actions || [];
  const descRaw = acoes[0]?.description || '';
  const descricao = String(descRaw)
    .replace(/<[^>]*>?/gm, '')
    .replace(/\[cid:[^\]]*\]/g, '')
    .replace(/[\r\n\t]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .substring(0, 300) || 'N/A';

  let solicitante = 'Solicitante não definido';
  let pessoaIdSetor = null;
  const criadorObjeto = t.createdBy || {};
  const listaClientes = Array.isArray(t.clients) ? t.clients : [];

  if (criadorObjeto.businessName || criadorObjeto.personName || criadorObjeto.name) {
    solicitante = criadorObjeto.businessName || criadorObjeto.personName || criadorObjeto.name;
    pessoaIdSetor = criadorObjeto.id ? String(criadorObjeto.id) : null;
  } else if (listaClientes.length > 0) {
    solicitante = listaClientes
      .map(c => c.businessName || c.personName || c.name)
      .filter(Boolean)
      .join(', ') || 'MÃºltiplos Clientes';
    pessoaIdSetor = listaClientes[0]?.id ? String(listaClientes[0].id) : null;
  } else if (criadorObjeto.id) {
    solicitante = `UsuÃ¡rio Desativado (ID: ${criadorObjeto.id})`;
    pessoaIdSetor = String(criadorObjeto.id);
  }

  let setorCliente = 'Setor não informado';
  if (pessoaIdSetor && personCustomFieldValuesMap && movideskClientSectorCustomFieldId) {
    const camposPessoa = personCustomFieldValuesMap[pessoaIdSetor];
    const valorSetor = camposPessoa?.[String(movideskClientSectorCustomFieldId)];
    if (!valorVazioOuGenerico(valorSetor)) {
      setorCliente = String(valorSetor).trim();
    }
  }

  const createdDate = t.createdDate ? new Date(t.createdDate) : null;
  const closedDate = t.resolvedIn ? new Date(t.resolvedIn) : null;

  return {
    id: String(t.id).trim(),
    dados: {
      area: setorCliente,
      titulo: t.subject || 'Sem TÃ­tulo',
      tipo: t.serviceFirstLevel ? String(t.serviceFirstLevel).trim() : 'N/A',
      categoria: t.category || 'N/A',
      descricao,
      prioridade: urgencyMap[t.urgency] || 'N/A',
      solicitante: String(solicitante).trim(),
      equipe,
      abertura: createdDate && !isNaN(createdDate.getTime()) ? createdDate.toISOString() : null,
      encerramento: closedDate && !isNaN(closedDate.getTime()) ? closedDate.toISOString() : null,
      status: t.status || 'N/A'
    }
  };
}

async function buscarTodosTicketsMovideskConcorrente(token, take = 300, maxParallel = 4) {
  let skip = 0;
  let todosTickets = [];

  while (true) {
    const lote = [];
    for (let i = 0; i < maxParallel; i++) {
      lote.push(
        axios.get('https://api.movidesk.com/public/v1/tickets', {
          params: {
            token,
            '$select': MOVIDESK_TICKET_SELECT,
            '$expand': MOVIDESK_TICKET_EXPAND,
            '$top': take,
            '$skip': skip + i * take,
          },
          httpsAgent,
          timeout: 90000,
        }).then(resp => {
          const dados = resp.data.value || resp.data || [];
          console.log(`[Movidesk] ${dados.length} tickets processados (skip=${skip + i * take})...`);
          return dados;
        }).catch(err => {
          console.error(`[Movidesk] Erro skip=${skip + i * take}: ${err.message}`);
          return [];
        })
      );
    }

    // executa todas em paralelo
    const resultados = await Promise.all(lote);
    const ticketsLote = resultados.flat();

    // adiciona ao acumulado
    todosTickets.push(...ticketsLote);

    // se o último lote veio vazio encerra
    if (ticketsLote.length < take * maxParallel) {
      break;
    }

    // avança o skip para o próximo bloco
    skip += take * maxParallel;
  }

  console.log(`[Movidesk] Total final: ${todosTickets.length} tickets carregados.`);
  return todosTickets;
}

async function realizarChamadaLote(token, batch, endpoint = MOVIDESK_TICKETS_ENDPOINTS[0]) {
  const filterString = batch.map(id => `id eq ${id}`).join(' or ');
  const resp = await axios.get(endpoint, {
    params: {
      token,
      '$select': MOVIDESK_TICKET_SELECT,
      '$expand': MOVIDESK_TICKET_EXPAND,
      '$filter': filterString
    },
    httpsAgent,
    timeout: 90000,
  });
  return resp.data && resp.data.value ? resp.data.value : (Array.isArray(resp.data) ? resp.data : []);
}


async function buscarTicketsEspecificosMovideskConcorrente(token, ids, personCustomFieldValuesMap, movideskClientSectorCustomFieldId) {
  const mapa = {};
  const batchSize = 15;
  const concurrency = 8;
  const maxRetries = 3;
  const retryDelayMs = 1500;

  const batches = [];
  for (let i = 0; i < ids.length; i += batchSize) {
    batches.push(ids.slice(i, i + batchSize));
  }

  console.log(`[Movidesk] Buscando ${ids.length} tickets antigos/faltantes em ${batches.length} lotes concorrentes...`);

  let lotesProcessados = 0;

  async function processarLote(lote) {
    if (lote.length === 0) return;
    let success = false;
    let retries = 0;
    let tickets = [];

    while (retries < maxRetries && !success) {
      try {
        tickets = await realizarChamadaLote(token, lote);
        success = true;
      } catch (err) {
        retries++;
        if (retries >= maxRetries) {
          for (const id of lote) {
            mapa[id] = { ignorado: true };
          }
        } else {
          await new Promise(resolve => setTimeout(resolve, retryDelayMs));
        }
      }
    }

    if (success && tickets && tickets.length > 0) {
      for (const t of tickets) {
        if (!t.id) continue;
        const idTicket = String(t.id).trim();
        const statusTicket = t.status || 'N/A';
        const equipe = t.ownerTeam && t.ownerTeam.trim() !== ''
          ? String(t.ownerTeam).trim()
          : (t.serviceFirstLevel || 'Equipe não definida');

        if (equipe === 'Suporte e Infraestrutura' || equipe === 'Administradores') {
          mapa[idTicket] = { ignorado: true, equipe };
          continue;
        }

        const prioridadeMovidesk = urgencyMap[t.urgency] || 'N/A';
        const categoria = t.category || 'N/A';
        const tipo = t.serviceFirstLevel ? String(t.serviceFirstLevel).trim() : 'N/A';
        const createdDate = t.createdDate ? new Date(t.createdDate) : null;
        const closedDate = t.resolvedIn ? new Date(t.resolvedIn) : null;

        const acoes = t.actions || [];
        const descRaw = acoes[0]?.description || '';
        const descricao = String(descRaw)
          .replace(/<[^>]*>?/gm, '')
          .replace(/\[cid:[^\]]*\]/g, '')
          .replace(/[\r\n\t]+/g, ' ')
          .replace(/\s+/g, ' ')
          .trim()
          .substring(0, 300) || 'N/A';

        let solicitante = 'Solicitante não definido';
        let personIdSetor = null;
        const criadorObjeto = t.createdBy || {};
        const listaClientes = Array.isArray(t.clients) ? t.clients : [];

        if (criadorObjeto.businessName || criadorObjeto.personName || criadorObjeto.name) {
          solicitante = criadorObjeto.businessName || criadorObjeto.personName || criadorObjeto.name;
          personIdSetor = criadorObjeto.id ? String(criadorObjeto.id) : null;
        } else if (listaClientes.length > 0) {
          solicitante = listaClientes.map(c => c.personName || c.name || c.businessName).filter(Boolean).join(', ') || 'Múltiplos Clientes';
          personIdSetor = listaClientes[0]?.id ? String(listaClientes[0].id) : null;
        }

        if ((solicitante === 'Solicitante não definido' || solicitante === 'Solicitante não informado') && criadorObjeto.id) {
          solicitante = `Usuário Desativado (ID: ${criadorObjeto.id})`;
          personIdSetor = String(criadorObjeto.id);
        }

        let setorCliente = 'Setor não informado';
        if (personIdSetor && personCustomFieldValuesMap && movideskClientSectorCustomFieldId) {
          const camposPessoa = personCustomFieldValuesMap[personIdSetor];
          if (camposPessoa && camposPessoa[String(movideskClientSectorCustomFieldId)] !== undefined && camposPessoa[String(movideskClientSectorCustomFieldId)] !== null) {
            setorCliente = String(camposPessoa[String(movideskClientSectorCustomFieldId)]).trim();
          }
        }

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
      }
    }
    lotesProcessados++;
    if (lotesProcessados % 20 === 0 || lotesProcessados === batches.length) {
      console.log(`[Movidesk] ${lotesProcessados}/${batches.length} lotes antigos processados (${Object.keys(mapa).length} tickets mapeados)...`);
    }
  }

  const queue = [...batches];
  async function worker() {
    while (queue.length > 0) {
      const lote = queue.shift();
      await processarLote(lote);
    }
  }

  const workers = Array.from({ length: concurrency }, () => worker());
  await Promise.all(workers);

  console.log(`[Movidesk] Busca concorrente concluída. Mapeados ${Object.keys(mapa).length} tickets antigos.`);
  return mapa;
}

async function buscarTicketsEspecificosMovideskFallback(token, ids, personCustomFieldValuesMap, movideskClientSectorCustomFieldId) {
  const mapa = {};
  const batchSize = 15;
  const concurrency = 8;
  const maxRetries = 3;
  const retryDelayMs = 1500;
  const batches = [];

  for (let i = 0; i < ids.length; i += batchSize) {
    batches.push(ids.slice(i, i + batchSize));
  }

  console.log(`[Movidesk] Complementando ${ids.length} tickets especificos em ${batches.length} lotes pelos endpoints normal, merged e past...`);

  let lotesProcessados = 0;

  async function processarLote(lote) {
    for (const endpoint of MOVIDESK_TICKETS_ENDPOINTS) {
      let tickets = [];
      let success = false;
      let retries = 0;

      while (retries < maxRetries && !success) {
        try {
          tickets = await realizarChamadaLote(token, lote, endpoint);
          success = true;
        } catch (err) {
          retries++;
          if (retries < maxRetries) {
            await new Promise(resolve => setTimeout(resolve, retryDelayMs));
          }
        }
      }

      if (!success || tickets.length === 0) continue;

      for (const ticket of tickets) {
        const normalizado = normalizarTicketMovidesk(ticket, personCustomFieldValuesMap, movideskClientSectorCustomFieldId);
        if (!normalizado) continue;
        mapa[normalizado.id] = mesclarDadosTicket(mapa[normalizado.id], normalizado.dados);
      }
    }

    for (const id of lote) {
      if (!mapa[id]) mapa[id] = { ignorado: true };
    }

    lotesProcessados++;
    if (lotesProcessados % 20 === 0 || lotesProcessados === batches.length) {
      console.log(`[Movidesk] ${lotesProcessados}/${batches.length} lotes complementados (${Object.keys(mapa).length} tickets mapeados)...`);
    }
  }

  const queue = [...batches];
  async function worker() {
    while (queue.length > 0) {
      const lote = queue.shift();
      await processarLote(lote);
    }
  }

  await Promise.all(Array.from({ length: concurrency }, () => worker()));

  console.log(`[Movidesk] Complemento por fallback concluido. Mapeados ${Object.keys(mapa).length} tickets.`);
  return mapa;
}

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
              created_at
              updated_at
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
  dataFimSprint.setDate(dataInicioSprint.getDate() + (SLA_DIAS - 1));
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
  const idsDaSprint = new Set();

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
    { header: 'Origem', key: 'origem', width: 18 },
  ];

  ticketsSheet.getRow(1).eachCell(cell => {
    cell.fill = estiloAzul;
    cell.font = fonteBranca;
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  });
  ticketsSheet.getRow(1).height = 24;

  for (const item of itensDaSprint) {
    try {
      let mondayMovideskIdRaw = String(getColVal(item, mondayTextColId)).trim();
      mondayMovideskIdRaw = mondayMovideskIdRaw.replace(/\D/g, '');
      if (mondayMovideskIdRaw) idsDaSprint.add(mondayMovideskIdRaw);

      const mov = (mondayMovideskIdRaw && mapaMovidesk[mondayMovideskIdRaw]) ? mapaMovidesk[mondayMovideskIdRaw] : {};

      const movFinal = (mov && !mov.ignorado) ? mov : {};
      const temDadosMovidesk = Object.keys(movFinal).length > 0;

      const prioMon = typeof getColVal === 'function' ? (mondayPriorityMap[getColVal(item, mondayPriorityColId)] || 'Não definida') : 'Não definida';
      const typeMon = typeof getColVal === 'function' ? getColVal(item, mondayTypeColId) : null;
      const descMon = typeof getColVal === 'function' ? getColVal(item, mondayDescriptionColId) : null;

      const ticketId = mondayMovideskIdRaw || 'ID não informado';

      const area = movFinal.area || 'Setor não informado';
      const tipo = typeMon || movFinal.tipo || 'Tipo não informado';
      const categoria = movFinal.categoria || 'Categoria não informada';
      const titulo = item.name || 'Sem Título';
      const descricao = movFinal?.descricao || getColVal(item, mondayDescriptionColId) || 'Descrição não informada';
      const prioridade = prioMon;
      const solicitante = movFinal?.solicitante && movFinal.solicitante !== 'Solicitante não definido'
        ? movFinal.solicitante
        : (getColVal(item, 'colunaSolicitanteMonday') || 'Solicitante não definido');
      const equipe = movFinal?.equipe && movFinal.equipe !== 'Equipe não definida'
        ? movFinal.equipe
        : (getColVal(item, 'colunaEquipeMonday') || 'Equipe não definida');

      let dataEntradaSprint = movFinal.abertura ? new Date(movFinal.abertura) : null;
      if (!dataEntradaSprint && t.createdDate) {
        dataEntradaSprint = new Date(t.createdDate);
      }

      let dataEncerramentoObj = null;
      if (movFinal.encerramento) {
        dataEncerramentoObj = new Date(movFinal.encerramento);
      } else if (t.resolvedIn) {
        dataEncerramentoObj = new Date(t.resolvedIn);
      } else if (t.closedIn) {
        dataEncerramentoObj = new Date(t.closedIn);
      } else if (t.reopenedIn) {
        dataEncerramentoObj = new Date(t.reopenedIn);
      }

      if (dataEncerramentoObj && dataEntradaSprint && dataEncerramentoObj < dataEntradaSprint) {
        dataEncerramentoObj = null;
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

      const encerradoNoPeriodo = STATUS_ENCERRADO.includes(statusFinal)
        && dataEncerramentoObj
        && dataEncerramentoObj >= dataInicioSprint
        && dataEncerramentoObj <= dataFimSprint;

      if (encerradoNoPeriodo) {
        dentroSLACount++;
      }

      let tempoResolucao = movFinal.tempoResolucao ?? 'N/A';
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
        origem: 'Planejado',
      });
      row.eachCell(cell => { cell.alignment = { vertical: 'middle', wrapText: true }; });
    } catch (errItem) {
      console.warn(`[Monday] Ignorando linha com erro na sprint ${sprintNome}:`, errItem.message);
    }
  }

  for (const [idMovidesk, mov] of Object.entries(mapaMovidesk)) {
    if (!idMovidesk || idsDaSprint.has(idMovidesk) || !mov || mov.ignorado) continue;

    const equipeForaSprint = normalizarTextoComparacao(mov.equipe || '');
    if (
      !equipeForaSprint.includes('SISTEMAS') &&
      !equipeForaSprint.includes('ADH')
    ) {
      continue;
    }

    const dataEncerramentoForaSprint = mov.encerramento ? new Date(mov.encerramento) : null;
    if (!dataEncerramentoForaSprint || isNaN(dataEncerramentoForaSprint.getTime())) continue;
    if (dataEncerramentoForaSprint < dataInicioSprint || dataEncerramentoForaSprint > dataFimSprint) continue;

    foraSLACount++;
    const area = mov.area || 'Setor nÃo informado';
    const tipo = mov.tipo || 'Tipo nÃo informado';
    const categoria = mov.categoria || 'Categoria nÃo informada';
    const titulo = mov.titulo || `Chamado ${idMovidesk}`;
    const descricao = mov.descricao || 'DescriÃ§Ã£o nÃo informada';
    const prioridade = mov.prioridade || 'NÃo definida';
    const solicitante = mov.solicitante || 'Solicitante nÃo definido';
    const equipe = mov.equipe || 'Equipe nÃo definida';
    const dataAberturaForaSprint = mov.abertura ? new Date(mov.abertura) : null;
    const aberturaFormatada = dataAberturaForaSprint && !isNaN(dataAberturaForaSprint.getTime())
      ? dataAberturaForaSprint.toLocaleDateString('pt-BR')
      : 'N/A';
    const encerramentoFormatado = dataEncerramentoForaSprint.toLocaleDateString('pt-BR');
    const semanaLinha = calcularSemana(dataEncerramentoForaSprint);
    const mesLinha = nomeMes(dataEncerramentoForaSprint);

    ticketsForaSLA.push({ ticket: idMovidesk, area, titulo, prioridade, atraso: 0 });

    const row = ticketsSheet.addRow({
      ticket: idMovidesk, area, tipo, categoria, titulo,
      descricao, prioridade, solicitante, equipe,
      abertura: aberturaFormatada,
      encerramento: encerramentoFormatado,
      status: mov.status || 'Encerrado',
      sla: 'Não',
      semana: semanaLinha,
      mes: mesLinha,
      origem: 'Fora da Sprint',
    });
    row.eachCell(cell => { cell.alignment = { vertical: 'middle', wrapText: true }; });
  }

  const totalEncerrados = dentroSLACount + foraSLACount;
  const percentualSLA = totalEncerrados > 0 ? (dentroSLACount / totalEncerrados) : 0;
  const tempoMedio = countResolvidos > 0 ? (totalTempoResolucao / countResolvidos).toFixed(2) : 0;

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
  // DASHBOARD
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

  dashSheet.getCell('D4').value = '% planejado';
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
  // RESUMO
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

  resumoSheet.getCell('J9').value = 'Planejado x Fora da Sprint';
  resumoSheet.getCell('J9').font = fonteNegrito;
  resumoSheet.getCell('J10').value = 'Planejado';
  resumoSheet.getCell('K10').value = totalTicketsDaSprint;

  resumoSheet.getCell('J11').value = 'Fora da Sprint';
  resumoSheet.getCell('K11').value = foraSLACount;

  resumoSheet.getCell('J12').value = 'Abertos';
  resumoSheet.getCell('K12').value = Math.max(totalTicketsDaSprint - dentroSLACount, 0);

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

  resumoSheet.getCell('D19').value = 'Demandas atendidas fora da sprint';
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

  // 1. TENTA CARREGAR OS CACHES LOCAIS DO DISCO (SE EXISTIREM)
  let mapaMovidesk = {};
  try {
    const data = await fs.readFile(path.join(dirPath, 'movidesk_cache.json'), 'utf8');
    mapaMovidesk = JSON.parse(data);
    console.log(`[Cache] Carregados ${Object.keys(mapaMovidesk).length} tickets do cache local do Movidesk.`);
  } catch {
    console.log('[Cache] Nenhum cache local do Movidesk encontrado. Iniciando do zero.');
  }

  let cacheSprints = {};
  try {
    const data = await fs.readFile(path.join(dirPath, 'sprints_cache.json'), 'utf8');
    cacheSprints = JSON.parse(data);
    console.log(`[Cache] Carregados metadados de ${Object.keys(cacheSprints).length} sprints do cache local.`);
  } catch {
    console.log('[Cache] Nenhum cache local de sprints encontrado.');
  }

  // 2. Busca os dados de pessoas e campos personalizados primeiro
  const personCustomFieldValuesMap = movideskToken
    ? await buscarValoresCamposPersonalizadosPessoas(movideskToken)
    : {};

  const listaEstruturaCamposClientes = movideskToken
    ? await buscarEstruturaCamposPersonalizadosClientes(movideskToken)
    : [];


  if (movideskToken) {
    const novosTickets = await buscarTodosTicketsMovidesk(movideskToken, personCustomFieldValuesMap, movideskClientSectorCustomFieldId);
    Object.assign(mapaMovidesk, novosTickets);
  }

  // 4. Busca todos os itens do Monday
  let todosItensMonday;
  try {
    todosItensMonday = await buscarTodosItensMonday(boardId, token);
  } catch (err) {
    console.error('[Monday] ERRO COMPLETO:', err);
    console.error('[Monday] RESPONSE:', err.response?.data);
    console.error('[Monday] STATUS:', err.response?.status);

    throw new Error(
      `[Monday] Erro crítico na extração de dados do Monday: ${err.response?.data?.errors
        ? JSON.stringify(err.response.data.errors)
        : err.message || JSON.stringify(err)
      }`
    );
  }

  // 5. Extrai todos os IDs únicos de tickets do Movidesk que estão no Monday
  const idsMovideskUnicos = new Set();
  for (const item of todosItensMonday) {
    let idRaw = String(getColVal(item, mondayTextColId)).trim().replace(/\D/g, '');
    if (idRaw) idsMovideskUnicos.add(idRaw);
  }
  const arrayIds = Array.from(idsMovideskUnicos);
  console.log(`[Monday] Identificados ${arrayIds.length} IDs únicos de tickets do Movidesk no board.`);

  // 6. Identifica quais IDs do Monday NÃO foram encontrados no mapa (tickets antigos)
  const idsFaltantes = [];
  for (const id of arrayIds) {
    if (ticketPrecisaComplemento(mapaMovidesk[id])) {
      idsFaltantes.push(id);
    }
  }
  console.log(`[Monday] Dos ${arrayIds.length} tickets do Monday, ${idsFaltantes.length} são antigos/faltantes.`);

  // 7. Busca concorrente paralela (8 requisições simultâneas) apenas para os tickets faltantes do "past"
  if (movideskToken && idsFaltantes.length > 0) {
    const mapaFaltantes = await buscarTicketsEspecificosMovideskFallback(
      movideskToken,
      idsFaltantes,
      personCustomFieldValuesMap,
      movideskClientSectorCustomFieldId
    );
    // Mescla os tickets recuperados sem apagar campos bons ja existentes no cache
    for (const [id, dados] of Object.entries(mapaFaltantes)) {
      mapaMovidesk[id] = mesclarDadosTicket(mapaMovidesk[id], dados);
    }
  }

  // 8. Agrupa os itens do Monday por Sprint
  const sprints = {};
  const gruposIgnorados = ['BACKLOG', 'PROXIMA SPRINT', 'PRÓXIMA SPRINT', 'SPRINT ATUAL'];

  for (const item of todosItensMonday) {
    const nomeGrupo = item.group?.title ? item.group.title.trim() : 'Sem Sprint';
    if (gruposIgnorados.includes(nomeGrupo.toUpperCase())) continue;
    if (!sprints[nomeGrupo]) sprints[nomeGrupo] = [];
    sprints[nomeGrupo].push(item);
  }

  console.log(`[Monday] ${Object.keys(sprints).length} sprints mapeadas.`);

  // 9. Gera os arquivos Excel de forma incremental (Apenas se houver alterações!)
  const resultados = [];
  let totalSprintsIgnoradas = 0;

  for (const [sprintNome, itensDaSprint] of Object.entries(sprints)) {
    try {
      const nomeArquivo = `monday_sprint_${sprintNome.toLowerCase()
        .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
        .replace(/\s+/g, '_').replace(/[^a-z0-9_\-]/g, '')}.xlsx`;
      const caminhoCompletoExcel = path.join(dirPath, nomeArquivo);

      // Verifica se o arquivo Excel já existe fisicamente no disco
      let existeExcel = false;
      try {
        await fs.access(caminhoCompletoExcel);
        existeExcel = true;
      } catch { }

      // Calcula a assinatura digital atual da sprint
      const fingerprintAtual = calcularFingerprintSprint(
        itensDaSprint,
        mapaMovidesk,
        mondayTextColId,
        mondayStatusColId,
        mondayPriorityColId
      );

      // SE O EXCEL EXISTE E A ASSINATURA É IGUAL: Pula a geração e aproveita o arquivo!
      if (existeExcel && cacheSprints[sprintNome] && cacheSprints[sprintNome].fingerprint === fingerprintAtual) {
        totalSprintsIgnoradas++;
        resultados.push({
          nome_exibicao: `MONDAY - ${sprintNome.toUpperCase()}`,
          caminho_arquivo: `utils/${nomeArquivo}`
        });
        continue;
      }

      // Caso contrário, gera o Excel normalmente e atualiza o cache
      console.log(`[Monday] Gerando Excel para: "${sprintNome}" (${itensDaSprint.length} itens).`);
      const res = await gerarExcelParaSprint(
        sprintNome, itensDaSprint, mapaMovidesk, dirPath,
        mondayTextColId, mondayStatusColId, mondayPriorityColId,
        mondayDescriptionColId, mondayTypeColId,
        mondayCreationDateColId, mondayResolutionDateColId,
        listaEstruturaCamposClientes
      );

      if (res) {
        resultados.push(res);
        cacheSprints[sprintNome] = {
          fingerprint: fingerprintAtual,
          arquivo: nomeArquivo
        };
      }
    } catch (err) {
      console.error(`[Monday] Falha na sprint "${sprintNome}":`, err.message);
    }
  }

  if (totalSprintsIgnoradas > 0) {
    console.log(`[Cache] ${totalSprintsIgnoradas} sprints sem alterações foram carregadas do cache local.`);
  }

  // 10. SALVA OS CACHES ATUALIZADOS NO DISCO PARA A PRÓXIMA EXECUÇÃO
  try {
    await fs.mkdir(dirPath, { recursive: true });
    await fs.writeFile(path.join(dirPath, 'movidesk_cache.json'), JSON.stringify(mapaMovidesk, null, 2), 'utf8');
    await fs.writeFile(path.join(dirPath, 'sprints_cache.json'), JSON.stringify(cacheSprints, null, 2), 'utf8');
    console.log('[Cache] Arquivos de cache local atualizados com sucesso.');
  } catch (errCache) {
    console.error('[Cache] Erro ao salvar caches locais:', errCache.message);
  }

  try {
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
      console.log('[Altona] Índice estático utils/sprints.json updated.');
    } else {
      console.error(' Abortando atualização do sprints.json: Nenhuma sprint válida foi gerada.');
    }
  } catch (errDiscovery) {
    console.error('[Altona] Erro ao gerar sprints.json:', errDiscovery.message);
  }

  return resultados;
}

module.exports = { sincronizarMondaySprints, buscarEstruturaCamposPersonalizadosClientes };
