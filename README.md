# ChangeLog
ChangeLog Sistemas - TI
Dashboard interno de governança de TI da EletroAço Altona. Puxa dados do Movidesk e do Monday.com, gera um Excel por Sprint e exibe tudo num painel HTML para a intranet.
Como funciona
O backend em Node/Express expõe algumas rotas. A principal delas sincroniza os dois sistemas: busca tickets do Movidesk, itens do Monday por GraphQL, cruza pelo ID do ticket e salva um .xlsx por Sprint na pasta utils/. O frontend lê esses arquivos direto e monta os gráficos com Google Charts.
A sincronização roda automaticamente a cada 15 dias. Pra disparar manualmente, basta chamar a rota com a API Key no header.
Estrutura
changelog/
├── index.html
├── script/
│   ├── config.js       # URL base do servidor 
│   ├── index.js        # Dashboard: gráficos, filtros, leitura do Excel
│   ├── server.js       # API, sincronizador setor, autenticação
│   ├── monday.js       # Integração Monday + geração do Excel
│   └── movidesk.js     # Integração Movidesk
├── style/
│   └── style.css
├── utils/              # Gerado automaticamente
│   ├── monday_sprint_*.xlsx
│   └── sprints.json
├── .env
└── 