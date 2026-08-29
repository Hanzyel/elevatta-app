# ELEVATTA ERP REV40.1 — Produção, Qualidade, Validação e Almoxarifado

## O que é esta versão

A REV40.1 transforma o Elevatta em um ERP de obra com duas superfícies integradas:

- `/` — PWA operacional mobile/offline para Apontamentos e FVS + Medição;
- `/erp.html` — Control Tower desktop para consulta, consolidação, validação, ferramentas, cadastros e administração.

A fonte central é `BASE_ELEVATTA_SHAREPOINT.xlsx`, hospedada no SharePoint. Fotos, assinaturas e arquivos importados são armazenados como arquivos no SharePoint e o Excel guarda os IDs/URLs. O desktop usa o proxy autenticado `/api/erp/file` para exibir essas evidências dentro do ERP sem expor token Microsoft.

## Módulos

1. Apontamentos
2. Qualidade / FVS + Medição
3. Validação & Consolidação
4. Ferramentas / Almoxarifado
5. Cadastros mestres
6. Administração / auditoria / sincronização

## Princípio central: FRENTE_ID

Apontamento, FVS, Medição e Ferramentas são vinculados pela mesma frente de serviço. A chave é determinística a partir de:

`OBRA + PACOTE + PAVIMENTO + LOCAL/FRENTE`

Se um apontamento legado não possuir pavimento nem local, ele permanece contabilizado na produção, mas fica como **localização pendente** e não recebe vínculo automático de qualidade/ferramentas.

## Offline-first

- shell do PWA em Cache Storage;
- dados operacionais e fotos em IndexedDB;
- outbox persistente;
- sincronização automática ao recuperar conexão;
- Background Sync quando suportado;
- trava compartilhada entre página e Service Worker;
- erros transitórios usam retentativa com backoff;
- conflitos de negócio (ex.: ferramenta já retirada em outro aparelho) ficam bloqueados para revisão na Administração e não são repetidos infinitamente.

## Ferramentas

A base contém `FER-001` a `FER-050` como posições de patrimônio. Até receber o cadastro real, ficam `CADASTRO PENDENTE` e `ATIVO=NAO`.

Uma retirada exige:

- funcionário;
- obra;
- pacote;
- local/frente;
- ferramenta disponível e ativa;
- responsável do almoxarifado;
- assinatura do funcionário.

A devolução registra condição, responsável pelo recebimento e foto quando necessário. Todo movimento permanece no histórico.

## Validação

O importador não depende do nome da aba. Ele procura a estrutura pelos cabeçalhos e valida conteúdo. Entre outras regras:

- data e matrícula;
- fração > 0 e <= 1;
- pessoa-dia <= 1 DH;
- ausência + trabalho;
- pacote/cadastro;
- duplicidade;
- obra/local;
- hash do arquivo;
- quarentena de linhas críticas;
- **qualquer CRÍTICO bloqueia a importação**; alertas não críticos permitem importação;
- de/para controlado de funcionários históricos por `matrícula + nome`, armazenado em `tbConfig`.

## Render

- Build Command: `npm install`
- Start Command: `npm start`
- Health Check Path: `/healthz`
- Node: `>=20 <25`

O endpoint `/healthz` deve retornar `version: 40.1.0`.

## Arquivos principais

- `index.html` — PWA operacional
- `erp.html` / `erp.js` / `erp.css` — ERP desktop
- `erp-core.js` — IndexedDB, outbox, FRENTE_ID e sincronização
- `validator.js` — validação/importação
- `server.mjs` — servidor + Data Gateway + OAuth Microsoft
- `m365.js` — compatibilidade/sincronização dos módulos operacionais existentes
- `sw.js` — cache/offline + Background Sync (cache final `elevatta-erp-v40-r3`)
- `erp-schema.json` — contrato das tabelas do ERP
- `BASE_ELEVATTA_SHAREPOINT.xlsx` — base central para upload no SharePoint
- `consolidacao_rev38_legacy.html` — consolidador legado preservado para conferência

Leia também `ARQUITETURA_ERP_REV40.md`, `CONFIGURAR_MICROSOFT_365.md` e `TESTES_REV40_ERP.md`.

## Pendências de implantação

Consulte `DADOS_PENDENTES_REV40.md`. As principais são o cadastro real FER-001…FER-050 e a aprovação dos perfis/permissões.
