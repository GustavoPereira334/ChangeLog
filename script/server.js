const express = require('express');
const cors = require('cors');
const fs = require('fs');
const fsPromises = require('fs/promises');
const path = require('path');
const chokidar = require('chokidar');
const dotenv = require('dotenv');

process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0';

const envPath = path.join(__dirname, '..', '.env');
const resultDotEnv = dotenv.config({ path: envPath });

if (resultDotEnv.error) {
    console.error("Erro ao carregar .env:", resultDotEnv.error);
} else {
    console.log("Variáveis carregadas com sucesso.");
}

const { sincronizarMondaySprints } = require('./monday');
const app = express();
const PORT = process.env.PORT || 3000;
const dirPath = path.join(__dirname, '..', 'utils');


app.use(cors());
app.use(express.json());
app.use('/utils', express.static(dirPath));

if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
    console.log(`Pasta 'utils' criada em: ${dirPath}`);
}

const watcher = chokidar.watch(dirPath, { persistent: true, ignoreInitial: true });
watcher.on('add', filePath => {
    console.log(`[${new Date().toLocaleTimeString()}] Novo arquivo detectado: ${path.basename(filePath)}`);
});

async function executarSincronizacaoCompleta() {
    const movideskToken = process.env.MOVIDESK_TOKEN;
    const mondayToken = process.env.MONDAY_TOKEN;

    if (!movideskToken) throw new Error("MOVIDESK_TOKEN não configurado no .env");
    if (!mondayToken) throw new Error("MONDAY_TOKEN não configurado no .env");

    const mondayTextColId = process.env.MONDAY_TEXT_COLUMN_ID;
    const mondayStatusColId = process.env.MONDAY_STATUS_COLUMN_ID;
    const mondayPriorityColId = process.env.MONDAY_PRIORITY_COLUMN_ID;
    const mondayDescriptionColId = process.env.MONDAY_DESCRIPTION_COLUMN_ID;
    const mondayTypeColId = process.env.MONDAY_TYPE_COLUMN_ID;
    const mondayCreationDateColId = process.env.MONDAY_CREATION_DATE_COLUMN_ID;
    const mondayResolutionDateColId = process.env.MONDAY_RESOLUTION_DATE_COLUMN_ID;
    const movideskClientSectorId = process.env.MOVIDESK_CLIENT_SECTOR_CUSTOM_FIELD_ID;

    if (!mondayTextColId || !mondayStatusColId || !mondayPriorityColId ||
        !mondayDescriptionColId || !mondayTypeColId) {
        throw new Error("Um ou mais IDs de coluna do Monday.com não estão configurados no .env");
    }

    console.log(`[Sync] Iniciando sincronização completa em ${new Date().toLocaleString()}...`);

    const result = await sincronizarMondaySprints(
        mondayToken,
        process.env.MONDAY_BOARD_ID,
        mondayTextColId,
        mondayStatusColId,
        mondayPriorityColId,
        mondayDescriptionColId,
        mondayTypeColId,
        mondayCreationDateColId,
        mondayResolutionDateColId,
        movideskClientSectorId,
        dirPath
    );

    console.log('[Sinc Monday] Monday concluído. Arquivos:', result.map(r => r.nome_exibicao).join(', '));
    return result;
}


// agendamento para Sincronização a cada 15 dias

const QUINZE_DIAS_MS = 15 * 24 * 60 * 60 * 1000;

function agendarProximaSincronizacao() {
    console.log(`[Sincronização] Próxima sincronização agendada em 15 dias (${new Date(Date.now() + QUINZE_DIAS_MS).toLocaleString()}).`);

    setTimeout(async () => {
        console.log('[Sincronização] Disparando sincronização automática...');
        try {
            await executarSincronizacaoCompleta();
            console.log('[Sincronização] Sincronização automática concluída com sucesso.');
        } catch (err) {
            console.error('[Sincronização] Erro na sincronização automática:', err.message);
        } finally {
            agendarProximaSincronizacao(); // reagenda para o próximo ciclo
        }
    }, QUINZE_DIAS_MS);
}

// ============================================================
// ROTAS PÚBLICAS (sem autenticação)
// ============================================================

// Lista os arquivos Excel gerados na pasta utils
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
                return { nome_exibicao: nomeExibicao, caminho_arquivo: `utils/${f}` };
            })
            .sort((a, b) => b.nome_exibicao.localeCompare(a.nome_exibicao));

        res.json(lista);
    } catch (err) {
        res.status(500).json({ erro: "Erro ao ler pasta de arquivos históricos." });
    }
});

// ============================================================
// ROTAS PROTEGIDAS (exigem x-api-key no header)
// ============================================================

// Sincronização manual completa (dev only)
app.get('/api/sincronizar-tudo', async (req, res) => {
    try {
        const result = await executarSincronizacaoCompleta();
        res.status(200).json({
            success: true,
            message: 'Sincronização completa realizada com sucesso!',
            files: result.map(r => r.nome_exibicao),
            data: result
        });
    } catch (error) {
        console.error('[Sync] Erro:', error.message);
        res.status(500).json({ success: false, erro: error.message });
    }
});

// Utilitário: lista campos personalizados de clientes no Movidesk
// Útil para identificar o ID correto do campo Setor
app.get('/api/campos-clientes', async (req, res) => {

    try {
        const TOKEN = process.env.MOVIDESK_TOKEN;
        if (!TOKEN) throw new Error("MOVIDESK_TOKEN não configurado no .env");

        const axios = require('axios');
        const https = require('https');
        const httpsAgent = new https.Agent({ rejectUnauthorized: false });

        console.log(`[${new Date().toLocaleTimeString()}] Buscando campos personalizados de pessoas...`);

        const camposEncontrados = {};
        let totalAnalisadas = 0;
        let skip = 0;
        const take = 300;
        const maxPessoas = 300;

        while (totalAnalisadas < maxPessoas) {
            let resp;
            try {
                resp = await axios.get('https://api.movidesk.com/public/v1/persons', {
                    params: { token: TOKEN, '$expand': 'customFieldValues', '$top': take, '$skip': skip },
                    httpsAgent,
                    timeout: 30000,
                });
            } catch (e) {
                console.warn(`[campos-clientes] Erro skip=${skip}: ${e.message}`);
                break;
            }

            const pagina = resp.data.value || (Array.isArray(resp.data) ? resp.data : []);
            if (pagina.length === 0) break;

            for (const pessoa of pagina) {
                if (totalAnalisadas >= maxPessoas) break;
                totalAnalisadas++;

                const campos = pessoa.customFieldValues || [];
                if (!campos.length) continue;

                for (const campo of campos) {
                    const id = String(campo.customFieldId || '?');

                    if (!camposEncontrados[id]) {
                        camposEncontrados[id] = {
                            customFieldId: id,
                            exemplosDeValor: [],
                            exemplosDeItems: []
                        };
                    }

                    if (campo.value !== null && campo.value !== undefined && campo.value !== '') {
                        const valor = String(campo.value);
                        if (!camposEncontrados[id].exemplosDeValor.includes(valor)) {
                            camposEncontrados[id].exemplosDeValor.push(valor);
                        }
                    }

                    if (Array.isArray(campo.items)) {
                        campo.items.forEach(item => {
                            const nomeItem = item.customFieldItem || item.name || item.label || item.value;
                            if (nomeItem && !camposEncontrados[id].exemplosDeItems.includes(nomeItem)) {
                                camposEncontrados[id].exemplosDeItems.push(nomeItem);
                            }
                        });
                    }
                }
            }
            skip += take;
        }

        const lista = Object.values(camposEncontrados);

        res.status(200).json({
            success: true,
            instrucao: "Identifique o campo 'Setor' pelos exemplosDeValor/exemplosDeItems e copie o customFieldId para o .env como MOVIDESK_CLIENT_SECTOR_CUSTOM_FIELD_ID",
            total_campos: lista.length,
            total_pessoas_analisadas: totalAnalisadas,
            campos: lista
        });

    } catch (error) {
        console.error(`[${new Date().toLocaleTimeString()}] Erro ao buscar campos de clientes:`, error.message);
        res.status(500).json({ success: false, erro: error.message });
    }
});

// ============================================================
// INICIALIZAÇÃO DO SERVIDOR
// ============================================================
app.listen(PORT, () => {
    console.log(`Servidor rodando em http://10.0.9.14:${PORT}`);
    agendarProximaSincronizacao();
});