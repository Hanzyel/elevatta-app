# Arquitetura — ELEVATTA ERP REV40.1

## Domínios

### Produção
- APONTAMENTOS
- FUNCIONARIOS
- EQUIPES
- PACOTES

### Qualidade
- FVS_REGISTROS
- FVS_ITENS
- MEDICOES
- FOTOS

### Operação / Recursos
- FERRAMENTAS
- TERMOS_FERRAMENTAS
- MOVIMENTOS_FERRAMENTAS

### Governança de dados
- FRENTES_SERVICO
- IMPORTACOES
- IMPORT_OCORRENCIAS
- AUDITORIA
- USUARIOS
- CONFIG
- OBRAS

## Chaves

- registros transacionais usam IDs próprios e imutáveis;
- matrícula não é a chave técnica do funcionário;
- número da linha do Excel nunca é chave de negócio;
- `FRENTE_ID` é determinístico;
- arquivos importados são deduplicados por SHA-256;
- movimentos de ferramentas nunca são apagados do histórico.

## FRENTE_ID

Uma frente identificável exige:

- obra;
- pacote;
- e pelo menos pavimento ou local/frente.

Sem localização suficiente, o apontamento permanece válido para produção mas fica `LOCALIZAÇÃO PENDENTE`, sem ligação automática com FVS/medição/ferramentas.

## Sincronização

### Operacional existente
`index.html → m365.js/elevatta-m365 IndexedDB → Graph/SharePoint`

### ERP novo
`erp.html/index → erp-core/elevatta-erp IndexedDB → Data Gateway server.mjs → Graph/SharePoint`

A fila ERP possui:

- `createdAt`;
- `attempts`;
- `nextAttemptAt`;
- `lastError`;
- `blocked`;
- anexos em `assets`.

Erros 429/timeout/5xx são transitórios. Erros 4xx permanentes ficam bloqueados para revisão, exceto autenticação e estados reconhecidamente transitórios.

## Concorrência

- Gateway serializa escritas na instância Node;
- Excel usa sessão persistente quando disponível;
- inserts novos são enviados em lotes de até 100 linhas;
- retirada valida disponibilidade no servidor imediatamente antes da gravação;
- devolução valida termo aberto e ferramenta realmente em saída naquele termo;
- página e Service Worker usam `flush-lock` compartilhado no IndexedDB ERP.

## Fotos e assinaturas

Excel guarda URL/ID, não base64. Arquivos ficam em SharePoint. Para visualização no desktop, `/api/erp/file` atua como proxy autenticado same-origin; o token Graph permanece no servidor.

## Evolução futura

A interface trabalha contra funções/API e não contra números de linha. Isso permite trocar o adaptador Excel/SharePoint por SQL/Dataverse no futuro sem reconstruir os módulos de negócio.

## Migração e validação histórica

- CRÍTICO bloqueia importação tanto na interface quanto no Gateway.
- Pacote sem código pode ser reconciliado automaticamente somente quando o nome possui correspondência única no cadastro.
- Funcionário histórico nunca é reconciliado automaticamente apenas pela matrícula. O de/para usa matrícula + nome e é persistido em CONFIG com auditoria.
- O motor mantém a linha original rastreável por `MATRICULA_ORIGEM`, `NOME_ORIGEM` e hash da importação durante a validação.

## Autorização

A tabela `tbUsuarios` já faz parte do contrato. O bloqueio por perfil ficará desativado até a empresa confirmar a matriz de permissões, evitando perda de acesso durante a migração.
