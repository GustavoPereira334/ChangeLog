const express = require('express');
const cors = require('cors');
const fs = require('fs');
const fsPromises = require('fs/promises');
const path = require('path');
const chokidar = require('chokidar');
const dotenv = require('dotenv');

const envPath = path.join(__dirname, '.env');
console.log("Tentando carregar .env em:", envPath);

const resultDotEnv = dotenv.config({ path: envPath });

if (resultDotEnv.error) {
    console.error("Erro ao carregar .env:", resultDotEnv.error);
} else {
    console.log("Variáveis carregadas com sucesso.");
    console.log("MOVIDESK_TOKEN existe:", !!process.env.MOVIDESK_TOKEN);
    console.log("MONDAY_TOKEN existe:", !!process.env.MONDAY_TOKEN);
    console.log("MONDAY_BOARD_ID:", process.env.MONDAY_BOARD_ID);
    console.log("MONDAY_TEXT_COLUMN_ID:", process.env.MONDAY_TEXT_COLUMN_ID);
    console.log("MONDAY_DESCRIPTION_COLUMN_ID:", process.env.MONDAY_DESCRIPTION_COLUMN_ID);
    console.log("MONDAY_TYPE_COLUMN_ID:", process.env.MONDAY_TYPE_COLUMN_ID);
    console.log("MONDAY_STATUS_COLUMN_ID:", process.env.MONDAY_STATUS_COLUMN_ID);
    console.log("MONDAY_PRIORITY_COLUMN_ID:", process.env.MONDAY_PRIORITY_COLUMN_ID);
    console.log("MONDAY_CREATION_DATE_COLUMN_ID:", process.env.MONDAY_CREATION_DATE_COLUMN_ID);
    console.log("MONDAY_RESOLUTION_DATE_COLUMN_ID:", process.env.MONDAY_RESOLUTION_DATE_COLUMN_ID);
    console.log("MOVIDESK_CLIENT_SECTOR_CUSTOM_FIELD_ID:", process.env.MOVIDESK_CLIENT_SECTOR_CUSTOM_FIELD_ID);
}

const { sincronizarMovidesk } = require('./movidesk');
const { sincronizarMondaySprints, buscarEstruturaCamposPersonalizadosClientes } = require('./monday');

// gambiarra temporária pra ignorar certificado em local/dev
process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0';

const app = express();
const PORT = process.env.PORT || 3000;
const dirPath = path.join(__dirname, '..', 'utils');

app.use(cors());
app.use(express.json());
app.use('/utils', express.static(dirPath));

if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath);
    console.log(`Pasta 'utils' criada em: ${dirPath}`);
}

const watcher = chokidar.watch(dirPath, { persistent: true, ignoreInitial: true });
watcher.on('add', filePath => {
    console.log(`[${new Date().toLocaleTimeString()}] Novo arquivo detectado: ${path.basename(filePath)}`);
});

function resolverSetor(pessoa) {
    const fieldId = String(process.env.MOVIDESK_CLIENT_SECTOR_CUSTOM_FIELD_ID);
    if (!fieldId) return "Sem setor (ID não configurado no .env)";

    // Localiza o campo personalizado do setor 
    const campo = pessoa.customFieldValues?.find(c => String(c.customFieldId) === fieldId);
    if (!campo) return "Sem setor";

    //captura o setor no movidesk (campo personalizado)
    if (campo.value !== null && campo.value !== undefined && String(campo.value).trim() !== '') {
        return String(campo.value).trim();
    }

    return "Sem setor (Texto em branco)";
}

app.get('/api/campos-clientes', async (req, res) => {
    try {
        const TOKEN = process.env.MOVIDESK_TOKEN;
        if (!TOKEN) throw new Error("MOVIDESK_TOKEN não configurado no .env");

        const axios = require('axios');
        const https = require('https');
        const httpsAgent = new https.Agent({ rejectUnauthorized: false });

        console.log(`[${new Date().toLocaleTimeString()}] Buscando campos personalizados direto das pessoas...`);

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
                    httpsAgent, timeout: 30000,
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

                    if (id === '243059') {
                        console.log('\n========== CAMPO SETOR ==========');
                        console.log('Pessoa:', pessoa.businessName || pessoa.personName);
                        console.log(JSON.stringify(campo, null, 2));
                        console.log('=================================\n');
                    }

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

app.get('/api/sincronizar-movidesk', async (req, res) => {
    try {
        const TOKEN = process.env.MOVIDESK_TOKEN;
        if (!TOKEN) throw new Error("MOVIDESK_TOKEN não configurado no .env");

        const sprintDesejada = req.query.sprint || 'Geral';
        console.log(`[${new Date().toLocaleTimeString()}] Iniciando Movidesk: ${sprintDesejada}...`);

        const result = await sincronizarMovidesk(sprintDesejada, TOKEN, dirPath);
        res.status(200).json(result);
    } catch (error) {
        console.error(`[${new Date().toLocaleTimeString()}] Erro Movidesk:`, error.message);
        res.status(500).json({ success: false, erro: error.message });
    }
});

app.get('/api/testar-pessoa', async (req, res) => {
    try {
        const TOKEN = process.env.MOVIDESK_TOKEN;
        const axios = require('axios');
        const https = require('https');
        const httpsAgent = new https.Agent({ rejectUnauthorized: false });

        let rules = [];
        try {
            const rulesResp = await axios.get('https://api.movidesk.com/public/v1/customFieldRules', { params: { token: TOKEN }, httpsAgent });
            rules = rulesResp.data || [];
        } catch (ruleError) {
            console.warn(`⚠️ [Aviso] Falha ao carregar customFieldRules (${ruleError.message}). Continuando sem regras.`);
        }

        const resp = await axios.get('https://api.movidesk.com/public/v1/persons', {
            params: { token: TOKEN, '$expand': 'customFieldValues', '$top': 1000 },
            httpsAgent
        });

        const pessoas = resp.data.value || (Array.isArray(resp.data) ? resp.data : []);
        const termo = (req.query.nome || '').toLowerCase();

        const encontrados = pessoas.filter(p => {
            const nome = (p.personName || p.businessName || p.profileName || p.email || '').toLowerCase();
            return nome.includes(termo);
        });

        res.json({
            total: encontrados.length,
            pessoas: encontrados.map(p => ({
                id: p.id,
                nome: p.personName,
                empresa: p.businessName,
                email: p.email,
                setor_resolvido: resolverSetor(p, rules),
                customFields: p.customFieldValues
            }))
        });

    } catch (err) {
        console.error(err);
        res.status(500).json({ success: false, erro: err.message });
    }
});

// rota pra testar o caso específico do Alan (fix do combobox line)
app.get('/api/testar-alan', async (req, res) => {
    try {
        const TOKEN = process.env.MOVIDESK_TOKEN;
        const axios = require('axios');
        const https = require('https');
        const httpsAgent = new https.Agent({ rejectUnauthorized: false });

        let rules = [];
        try {
            // A rota correta na API do Movidesk para ler os itens de listas/combobox
            console.log(`[Movidesk] Baixando dicionário de opções de listas...`);
            const rulesResp = await axios.get('https://api.movidesk.com/public/v1/customFieldRules', {
                params: { token: TOKEN },
                httpsAgent
            });

            // Armazena e garante que a array de opções venha estruturada por ID de regra
            rules = rulesResp.data || [];
        } catch (ruleError) {
            console.warn(`⚠️ [Aviso] Falha ao carregar customFieldRules (${ruleError.message}). Usando fallback.`);
        }

        const resp = await axios.get('https://api.movidesk.com/public/v1/persons', {
            params: { token: TOKEN, '$expand': 'customFieldValues', '$top': 1000 },
            httpsAgent
        });

        const pessoas = resp.data.value || (Array.isArray(resp.data) ? resp.data : []);
        const alan = pessoas.find(p => (p.businessName || p.personName || '').toLowerCase().includes('alan jonis da silva'));

        if (!alan) {
            return res.json({ success: false, message: 'Alan não encontrado' });
        }

        // Passa os dados brutos e o catálogo de regras baixado para descriptografar o índice 'line'
        const setorDefinido = resolverSetor(alan, rules);

        const respostaFinal = {
            ...alan,
            setor_resolvido: setorDefinido
        };

        console.log('\n========== DEBUG ALAN RE-ALINHADO (COMBOBOX) ==========');
        console.log(JSON.stringify(respostaFinal, null, 2));
        console.log('========================================================\n');

        res.json(respostaFinal);

    } catch (err) {
        res.status(500).json({ erro: err.message });
    }
});

app.get('/api/debug-combobox', async (req, res) => {
    try {
        const axios = require('axios');
        const https = require('https');

        const resp = await axios.get('https://api.movidesk.com/public/v1/customFieldRules', {
            params: { token: process.env.MOVIDESK_TOKEN },
            httpsAgent: new https.Agent({ rejectUnauthorized: false })
        });

        res.json(resp.data);
    } catch (err) {
        res.status(500).json({ erro: err.message, response: err.response?.data });
    }
});

app.get('/api/sincronizar-monday-sprints', async (req, res) => {
    try {
        const mondayToken = process.env.MONDAY_TOKEN;
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
            !mondayDescriptionColId || !mondayTypeColId ||
            !mondayCreationDateColId || !mondayResolutionDateColId) {
            throw new Error("Um ou mais IDs de coluna do Monday.com não estão configurados no .env");
        }

        console.log(`[${new Date().toLocaleTimeString()}] Iniciando sincronização Monday.com...`);

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

        res.status(200).json({
            success: true,
            message: "Monday sincronizado. Arquivos gerados: " + result.map(r => r.nome_exibicao).join(', '),
            files: result
        });
    } catch (error) {
        console.error(`[${new Date().toLocaleTimeString()}] Erro Monday:`, error.message);
        res.status(500).json({ success: false, erro: error.message });
    }
});

// lê os xlsx que já foram gerados na pasta utils
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

app.get('/api/sincronizar-tudo', async (req, res) => {
    try {
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
            !mondayDescriptionColId || !mondayTypeColId ||
            !mondayCreationDateColId || !mondayResolutionDateColId) {
            throw new Error("Um ou mais IDs de coluna do Monday.com não estão configurados no .env");
        }

        const sprintDesejada = req.query.sprint || 'Geral';

        console.log("Iniciando sincronização completa...");
        await sincronizarMovidesk(sprintDesejada, movideskToken, dirPath);
        console.log("Movidesk sincronizado com sucesso.");

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

        res.status(200).json({
            success: true,
            message: "Sincronização completa realizada com sucesso!",
            files: result.map(r => r.nome_exibicao),
            data: result
        });
    } catch (error) {
        console.error("Erro na sincronização completa:", error.message);
        res.status(500).json({ success: false, erro: error.message });
    }
});

app.listen(PORT, () => {
    console.log(`Servidor rodando em http://localhost:${PORT}`);
});