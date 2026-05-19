const axios = require('axios');
const https = require('https');
const path = require('path');
const fs = require('fs').promises;

// Altere para o ID numérico real do campo que você encontrou no Movidesk
const ID_DO_MEU_CAMPO_PERSONALIZADO = 12345; 

async function sincronizarMovidesk(sprintDesejada, token, outputDirPath) {
    const agent = new https.Agent({ rejectUnauthorized: false });

    console.log(`[${new Date().toLocaleTimeString()}] Buscando tickets do Movidesk para JSON...`);

    const response = await axios.get('https://api.movidesk.com/public/v1/tickets', {
        params: {
            token,
            // 1. INCLUÍDO 'customFieldValues' no $select
            '$select': 'id,subject,category,urgency,status,createdDate,closedIn,ownerTeam,serviceFirstLevel,slaSolutionTime,slaSolutionDate,actions,customFieldValues',
            // 2. INCLUÍDO 'customFieldValues' no $expand
            '$expand': 'actions($select=description),createdBy,customFieldValues',
            '$top': 500, 
            '$filter': "createdDate gt 2023-01-01T00:00:00Z"
        },
        httpsAgent: agent,
        timeout: 90000
    });

    const listaBruta = response.data.value || (Array.isArray(response.data) ? response.data : []);
    if (listaBruta.length === 0) throw new Error("Nenhum ticket retornado da API do Movidesk.");

    const ticketsFiltrados = listaBruta.filter(t => {
        const equipe = t.ownerTeam ? String(t.ownerTeam).trim() : "";
        return equipe !== 'Suporte e Infraestrutura' && equipe !== 'Administradores';
    });

    const ticketsProcessados = ticketsFiltrados.map(ticket => {
        const prioridade = 'N/A'; 

        const area = ticket.createdBy?.department || ticket.category || 'N/A';
        const tipo = ticket.serviceFirstLevel || 'N/A';
        const ticketStatus = ticket.status || '';
        const createdDate = ticket.createdDate ? new Date(ticket.createdDate) : null;
        const closedDate = ticket.closedIn ? new Date(ticket.closedIn) : null;
        const slaDate = ticket.slaSolutionDate ? new Date(ticket.slaSolutionDate) : null;

        let dentroSLA = 'N/A';
        let slaDias = null;
        let tempoResolucao = 'N/A';

        const dataLimite = slaDate;
        const dataEntrega = closedDate || new Date();

        if (dataLimite && createdDate) {
            slaDias = Math.ceil((dataLimite - createdDate) / (1000 * 60 * 60 * 24));
            dentroSLA = dataEntrega <= dataLimite ? 'Sim' : 'Não';
        }

        if (createdDate && closedDate) {
            tempoResolucao = Math.ceil((closedDate - createdDate) / (1000 * 60 * 60 * 24));
        }

        const descRaw = ticket.actions?.[0]?.description || 'Sem descrição';
        const descLimpa = String(descRaw).replace(/<[^>]*>?/gm, '').substring(0, 500);

        // 3. LOGICA PARA FILTRAR E PEGAR O VALOR DO SEU CAMPO
        let valorCampoPersonalizado = 'N/A';
        if (ticket.customFieldValues && Array.isArray(ticket.customFieldValues)) {
            const campoEncontrado = ticket.customFieldValues.find(
                f => f.customFieldId === ID_DO_MEU_CAMPO_PERSONALIZADO
            );
            if (campoEncontrado) {
                // Se for campo de texto/número usa .value, se for lista com multiplas opções pode requerer lógica adicional
                valorCampoPersonalizado = campoEncontrado.value || 'N/A'; 
            }
        }

        return {
            id: ticket.id,
            area,
            tipo,
            categoria: ticket.category || 'N/A',
            subject: ticket.subject, 
            description: descLimpa,
            prioridade, 
            solicitante: ticket.createdBy?.personName || 'N/A',
            equipe: ticket.ownerTeam || 'N/A',
            abertura: createdDate ? createdDate.toISOString() : null,
            encerramento: closedDate ? closedDate.toISOString() : null,
            status: ticketStatus,
            slaDias,
            tempoResolucao,
            dentroSLA,
            slaSolutionDate: slaDate ? slaDate.toISOString() : null,
            // 4. ADICIONADO O DADO NO RETORNO DO JSON
            meuCampoPersonalizado: valorCampoPersonalizado 
        };
    });

    const jsonPath = path.join(outputDirPath, 'dadosMovidesk.json');
    await fs.writeFile(jsonPath, JSON.stringify(ticketsProcessados, null, 2));

    console.log(`[${new Date().toLocaleTimeString()}] Arquivo JSON do Movidesk salvo com sucesso em: ${jsonPath}`);

    return { success: true, message: "Dados do Movidesk sincronizados e salvos em JSON." };
}

module.exports = { sincronizarMovidesk };