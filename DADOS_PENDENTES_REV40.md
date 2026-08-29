# ELEVATTA ERP REV40.1 — Dados necessários para completar a implantação

A arquitetura pode funcionar sem estes dados, mas eles são necessários para a implantação definitiva.

## 1. Ferramentas 01–50 — prioridade alta

Enviar, para cada item disponível:

- número (01 a 50);
- nome da ferramenta/equipamento;
- categoria, se souber;
- marca, se souber;
- modelo, se souber;
- número de série/patrimônio complementar, se houver;
- condição atual;
- observação opcional.

Formato livre: Excel, foto de relação, texto ou lista parcial.

Só depois dessa identificação é seguro pesquisar e associar imagens de referência. A fotografia real do patrimônio deve ter prioridade sobre imagem ilustrativa.

## 2. Perfis do ERP — prioridade alta antes de habilitar bloqueios

Proposta:

- `ADMIN` — acesso total, usuários, conflitos, de/para e cadastros;
- `ENGENHARIA` — produção, FVS/medição, consolidação, validação e consulta de ferramentas;
- `ALMOXARIFADO` — ferramentas, saída, devolução, inventário e consulta necessária de funcionário/obra/pacote;
- `CAMPO` — apontamentos e FVS operacionais;
- `CONSULTA` — somente leitura dos painéis autorizados.

Precisamos confirmar:

- quem pode editar/aprovar FVS;
- quem pode validar e efetivar importação;
- quem pode cadastrar/movimentar ferramenta;
- quem pode resolver conflitos da outbox;
- quem pode alterar de/para e cadastros mestres;
- quem pode acessar Administração.

Até essa definição ser aprovada, a REV40 não força RBAC para evitar bloquear os usuários atuais.

## 3. De/para histórico de funcionários

O próprio módulo `Validação & Consolidação` apresenta os grupos históricos não reconhecidos e permite escolher o funcionário atual.

Regra: **não fazer associação automática apenas pela matrícula antiga**, porque o material histórico contém matrículas antigas reutilizadas com nomes diferentes.

## 4. Tenant Microsoft 365

Para homologação final no Render:

- `M365_TENANT_ID`;
- `M365_CLIENT_ID`;
- `M365_CLIENT_SECRET`;
- `M365_WORKBOOK_URL`;
- `SESSION_SECRET`;
- `APP_BASE_URL`.

Não enviar segredo para ser gravado no código. Configurar diretamente nas variáveis seguras do Render.
