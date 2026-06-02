const express = require('express');
const cors = require('cors');
const fs = require('fs');
const fsPromises = require('fs/promises');
const path = require('path');
const chokidar = require('chokidar');

const envPath = path.join(__dirname, '.env');
console.log("Tentando carregar .env em:", envPath);

const dotenv = require('dotenv');
const resultDotEnv = dotenv.config({ path: envPath });

if (resultDotEnv.error) {
    console.error("Erro ao carregar .env:", resultDotEnv.error);
} else {
    console.log("Variáveis carregadas com sucesso.");
    console.log("MOVIDESK_TOKEN existe:", !!process.env.MOVIDESK_TOKEN);
    console.log("MONDAY_TOKEN existe:", !!process.env.MONDAY_TOKEN);
    console.log("MONDAY_TEXT_COLUMN_ID:", process.env.MONDAY_TEXT_COLUMN_ID);
    console.log("MONDAY_DESCRIPTION_COLUMN_ID:", process.env.MONDAY_DESCRIPTION_COLUMN_ID);
    console.log("MONDAY_TYPE_COLUMN_ID:", process.env.MONDAY_TYPE_COLUMN_ID);
    console.log("MONDAY_STATUS_COLUMN_ID:", process.env.MONDAY_STATUS_COLUMN_ID);
    console.log("MONDAY_PRIORITY_COLUMN_ID:", process.env.MONDAY_PRIORITY_COLUMN_ID);
    console.log("MOVIDESK_CLIENT_SECTOR_CUSTOM_FIELD_ID:", process.env.MOVIDESK_CLIENT_SECTOR_CUSTOM_FIELD_ID);
    console.log("MONDAY_CREATION_DATE_COLUMN_ID:", process.env.MONDAY_CREATION_DATE_COLUMN_ID);
    console.log("MONDAY_RESOLUTION_DATE_COLUMN_ID:", process.env.MONDAY_RESOLUTION_DATE_COLUMN_ID);
}


// Importa a lógica de sincronização
const { sincronizarMovidesk } = require('./movidesk');
const { sincronizarMondaySprints } = require('./monday');

// CUIDADO: Desabilita a verificação de certificado TLS para requisições HTTPS.
// Reafirmo que esta é uma vulnerabilidade de segurança grave e deve ser usada com extrema cautela,
// preferencialmente APENAS em ambientes de desenvolvimento isolados e NUNCA em produção.
process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0';

const app = express();
const PORT = process.env.PORT || 3000;

// Define o caminho da pasta 'utils' na raiz do projeto (um nível acima)
const dirPath = path.join(__dirname, '..', 'utils');

// Middlewares
app.use(cors());
app.use(express.json());
app.use('/utils', express.static(dirPath));

// Garante que a pasta 'utils' exista na raiz
if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath);
    console.log(`Pasta 'utils' criada em: ${dirPath}`);
}

// Monitoramento de arquivos com Chokidar
const watcher = chokidar.watch(dirPath, { persistent: true, ignoreInitial: true });
watcher.on('add', filePath => {
    console.log(`[${new Date().toLocaleTimeString()}] Novo arquivo detectado: ${path.basename(filePath)}`);
});

// Rota Movidesk 
app.get('/api/sincronizar-movidesk', async (req, res) => {
    console.log("Caminho utilizado para arquivos:", dirPath);

    try {
        const TOKEN = process.env.MOVIDESK_TOKEN;
        if (!TOKEN) throw new Error("MOVIDESK_TOKEN não configurado no .env");

        const sprintDesejada = req.query.sprint || 'Geral';
        console.log(`[${new Date().toLocaleTimeString()}] Iniciando Movidesk (apenas JSON): ${sprintDesejada}...`);

        const result = await sincronizarMovidesk(sprintDesejada, TOKEN, dirPath);
        res.status(200).json(result);
    } catch (error) {
        console.error(`[${new Date().toLocaleTimeString()}] Erro Movidesk:`, error.message);
        res.status(500).json({ success: false, erro: error.message });
    }
});

// Rota Monday
app.get('/api/sincronizar-monday-sprints', async (req, res) => {
    try {
        const TOKEN = process.env.MONDAY_TOKEN;
        if (!TOKEN) throw new Error("MONDAY_TOKEN não configurado no .env");

        // Os IDs das colunas do Monday são lidos do .env
        const mondayTextColId = process.env.MONDAY_TEXT_COLUMN_ID;
        const movideskClientSectorCustomFieldId = process.env.MOVIDESK_CLIENT_SECTOR_CUSTOM_FIELD_ID;
        const mondayStatusColId = process.env.MONDAY_STATUS_COLUMN_ID;
        const mondayPriorityColId = process.env.MONDAY_PRIORITY_COLUMN_ID;
        const mondayCreationDateColId = process.env.MONDAY_CREATION_DATE_COLUMN_ID;
        const mondayResolutionDateColId = process.env.MONDAY_RESOLUTION_DATE_COLUMN_ID;
        const mondayDescriptionColId = process.env.MONDAY_DESCRIPTION_COLUMN_ID;
        const mondayTypeColId = process.env.MONDAY_TYPE_COLUMN_ID;

        // VALIDAÇÃO COMPLETA DE TODOS OS IDs
        if (!mondayTextColId || !mondayStatusColId || !mondayPriorityColId || !mondayDescriptionColId || !mondayTypeColId || !mondayCreationDateColId || !mondayResolutionDateColId || !movideskClientSectorCustomFieldId) {
            throw new Error("Um ou mais IDs de coluna/campo do Monday.com/Movidesk não estão configurados no .env. Verifique MONDAY_TEXT_COLUMN_ID, MONDAY_STATUS_COLUMN_ID, MONDAY_PRIORITY_COLUMN_ID, MONDAY_DESCRIPTION_COLUMN_ID, MONDAY_TYPE_COLUMN_ID, MONDAY_CREATION_DATE_COLUMN_ID, MONDAY_RESOLUTION_DATE_COLUMN_ID, MOVIDESK_CLIENT_SECTOR_CUSTOM_FIELD_ID.");
        }

        console.log(`[${new Date().toLocaleTimeString()}] Iniciando sincronização Monday.com...`);

        // ATRIBUIÇÃO DO RESULTADO À VARIÁVEL 'result'
        const result = await sincronizarMondaySprints(
            TOKEN, // ou mondayToken
            process.env.MONDAY_BOARD_ID,
            mondayTextColId,
            mondayStatusColId,
            mondayPriorityColId,
            mondayDescriptionColId,
            mondayTypeColId,
            mondayCreationDateColId, // <<< PARÂMETRO ADICIONADO AQUI
            mondayResolutionDateColId, // <<< PARÂMETRO ADICIONADO AQUI
            movideskClientSectorCustomFieldId,
            dirPath
        );
        res.status(200).json({ success: true, message: "Monday sincronizado. Arquivos gerados: " + (result.map(r => r.nome_exibicao) || []).join(', '), files: result });
    } catch (error) {
        console.error(`[${new Date().toLocaleTimeString()}] Erro na rota Monday:`, error.message);
        res.status(500).json({ success: false, erro: error.message });
    }
});

// Rota para listar arquivos
app.get('/api/sprints-da-pasta', async (req, res) => {
    try {
        const files = await fsPromises.readdir(dirPath);

        const lista = files
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

        res.json(lista);
    } catch (err) {
        res.status(500).json({ erro: "Erro ao ler pasta de arquivos históricos." });
    }
});


app.get('/api/sincronizar-tudo', async (req, res) => {
    try {
        console.log("Iniciando sincronização completa...");

        const movideskToken = process.env.MOVIDESK_TOKEN;
        const mondayToken = process.env.MONDAY_TOKEN;

        if (!movideskToken) throw new Error("MOVIDESK_TOKEN não configurado no .env para sincronização completa.");
        if (!mondayToken) throw new Error("MONDAY_TOKEN não configurado no .env para sincronização completa.");

        // IDs das colunas do Monday
        const mondayTextColId = process.env.MONDAY_TEXT_COLUMN_ID;
        const movideskClientSectorCustomFieldId = process.env.MOVIDESK_CLIENT_SECTOR_CUSTOM_FIELD_ID;
        const mondayStatusColId = process.env.MONDAY_STATUS_COLUMN_ID;
        const mondayCreationDateColId = process.env.MONDAY_CREATION_DATE_COLUMN_ID;
        const mondayResolutionDateColId = process.env.MONDAY_RESOLUTION_DATE_COLUMN_ID;
        const mondayPriorityColId = process.env.MONDAY_PRIORITY_COLUMN_ID;
        const mondayDescriptionColId = process.env.MONDAY_DESCRIPTION_COLUMN_ID;
        const mondayTypeColId = process.env.MONDAY_TYPE_COLUMN_ID;

        // VALIDAÇÃO COMPLETA DE TODOS OS IDs
        if (!mondayTextColId || !mondayStatusColId || !mondayPriorityColId || !mondayDescriptionColId || !mondayTypeColId || !mondayCreationDateColId || !mondayResolutionDateColId || !movideskClientSectorCustomFieldId) {
            throw new Error("Um ou mais IDs de coluna/campo do Monday.com/Movidesk não estão configurados no .env para sincronização completa.");
        }

        const sprintDesejada = req.query.sprint || 'Geral';
        await sincronizarMovidesk(sprintDesejada, movideskToken, dirPath);
        console.log("Dados do Movidesk sincronizados (JSON) com sucesso.");

        // ATRIBUIÇÃO DO RESULTADO À VARIÁVEL 'result'
        const result = await sincronizarMondaySprints(
            TOKEN, // ou mondayToken
            process.env.MONDAY_BOARD_ID,
            mondayTextColId,
            mondayStatusColId,
            mondayPriorityColId,
            mondayDescriptionColId,
            mondayTypeColId,
            mondayCreationDateColId, // <<< PARÂMETRO ADICIONADO AQUI
            mondayResolutionDateColId, // <<< PARÂMETRO ADICIONADO AQUI
            movideskClientSectorCustomFieldId,
            dirPath
        );


        res.status(200).json({ success: true, message: "Sincronização completa realizada com sucesso!", files: result.map(r => r.nome_exibicao), data: result });
    } catch (error) {
        console.error("Erro na sincronização completa:", error.message);
        res.status(500).json({ success: false, erro: error.message });
    }
});

app.listen(PORT, () => {
    console.log(`Servidor rodando em http://localhost:${PORT}`);
});