# ChangeLog
ChangeLog Sistemas - TI
Dashboard interno de governança de TI da EletroAço Altona. Puxa dados do Movidesk e do Monday.com, gera um Excel por Sprint e exibe tudo num painel HTML para a intranet.
Como funciona
O backend em Node/Express expõe algumas rotas. A principal delas sincroniza os dois sistemas: busca tickets do Movidesk, itens do Monday por GraphQL, cruza pelo ID do ticket e salva um .xlsx por Sprint na pasta utils/. O frontend lê esses arquivos direto e monta os gráficos com Google Charts.
A sincronização roda automaticamente a cada 15 dias. Pra disparar manualmente, basta chamar a rota com a API Key no header
