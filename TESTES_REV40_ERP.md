# ELEVATTA ERP REV40.1 — Relatório de auditoria e testes

Data da revisão: 28/08/2026

## 1. Escopo auditado

A REV40.1 foi revisada como um ERP integrado, e não somente como páginas isoladas. Foram auditados:

- PWA operacional `/`;
- Apontamentos;
- FVS + Medição Física;
- câmera/editor preservados da revisão anterior;
- Control Tower desktop `/erp.html`;
- Validação & Consolidação;
- Ferramentas / Almoxarifado;
- IndexedDB e filas offline;
- Service Worker;
- Data Gateway Node;
- autenticação Microsoft 365;
- contrato de tabelas do Excel central;
- 19 arquivos reais de apontamento fornecidos para a auditoria.

## 2. Resultado executivo

**Status técnico da build:** APROVADA PARA HOMOLOGAÇÃO EM AMBIENTE REAL.

Não significa homologação final do tenant Microsoft ou da câmera física. Esses dois testes dependem, respectivamente, das credenciais/arquivo SharePoint da empresa e de um aparelho Android/iPhone real.

## 3. Testes estruturais

Executados com sucesso:

- `node --check server.mjs`;
- `node --check m365.js`;
- `node --check rev32.js`;
- `node --check rev33.js`;
- `node --check sw.js`;
- `node --check erp-core.js`;
- `node --check validator.js`;
- `node --check erp.js`;
- validação dos 10 blocos JavaScript internos do `index.html`;
- verificação de IDs HTML duplicados;
- verificação de referências locais ausentes;
- verificação do shell pré-cacheado do Service Worker;
- validação de `package.json` e scripts npm.

Resultado mais recente:

- `index.html`: 232 IDs, 0 duplicados, 0 referências locais faltantes, 10 scripts internos sem erro;
- `erp.html`: 26 IDs, 0 duplicados, 0 referências locais faltantes, script interno sem erro;
- Service Worker: 17 arquivos essenciais no shell, 0 ausentes.

## 4. Smoke test do servidor

Servidor iniciado localmente com `npm start`/`node server.mjs`.

Verificado:

| Rota | Resultado esperado | Resultado |
|---|---|---|
| `/healthz` | HTTP 200 | OK |
| `/api/erp/schema` | HTTP 200 | OK |
| `/` | HTTP 200 | OK |
| `/erp.html` | HTTP 200 | OK |
| `/erp.js` | HTTP 200 | OK |
| `/sw.js` | HTTP 200 | OK |
| `/api/erp/status` sem variáveis M365 | falha controlada | HTTP 503 correto |
| `/api/erp/file` sem M365 configurado | falha controlada | HTTP 503 correto |

`/healthz` informou versão `40.1.0` e schema ERP `40.0.0`.

## 4.1 Base Excel central

Conferência estrutural final do arquivo `BASE_ELEVATTA_SHAREPOINT.xlsx`:

- 19 abas físicas;
- 18 tabelas estruturadas de banco;
- nomes e cabeçalhos das 18 tabelas exatamente compatíveis com `erp-schema.json`;
- nenhuma tabela de banco faltando;
- nenhuma tabela de banco extra;
- `tbFerramentas` em `A4:O54`, correspondendo a cabeçalho + 50 posições de patrimônio;
- primeira posição `FER-001`;
- última posição `FER-050`;
- nenhum token de erro `#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?` ou `#N/A` encontrado na estrutura XML do workbook;
- `unzip -t BASE_ELEVATTA_SHAREPOINT.xlsx`: sem erro de integridade do pacote XLSX.

## 5. Teste dos 19 apontamentos reais

O motor REV40 foi executado contra os 19 arquivos `.xlsx` fornecidos.

### Totais

- arquivos testados: **19**;
- lançamentos aceitos: **709**;
- linhas em quarentena: **6**;
- críticos: **6**;
- alertas: **139**;
- lançamentos produtivos sem localização suficiente para FRENTE_ID: **31**.

### Críticos reais

O arquivo de 27/08/2026 ficou **BLOQUEADO**:

- 5 lançamentos com trabalho/ausência e matrícula ausente;
- 1 lançamento de trabalho com fração inválida/ausente.

A política final é:

- `CRITICO` => **BLOQUEADO** e não pode ser gravado no ERP;
- erro estrutural/sem lançamento válido => **REJEITADO**;
- somente alertas => **APROVADO_COM_ALERTAS**;
- sem ocorrência relevante => **APROVADO**.

O Gateway repete essa validação na gravação: esconder o botão no navegador não é a única proteção.

### Alertas agregados

- `COLUNA_OPCIONAL_AUSENTE`: 20 ocorrências;
- `FUNCIONARIO_NAO_CADASTRADO`: 88 ocorrências históricas;
- `LOCAL_AUSENTE`: 31 ocorrências.

O desktop agrupa ocorrências repetidas para não exibir dezenas de avisos idênticos.

## 6. Reconciliação de pacotes

Foi detectado que modelos históricos usam `—` no campo de código, embora o nome do serviço esteja correto.

Correção implementada:

1. código válido tem prioridade;
2. se o código está vazio/`—`, o validador procura o nome normalizado;
3. correspondência única usa automaticamente o pacote canônico;
4. nome ambíguo exige de/para;
5. ausência não é validada como pacote de produção.

Após a correção, os falsos alertas de pacote dos 19 arquivos caíram de **192 para 0**.

## 7. Reconciliação de funcionários históricos

Não foi aplicado alias automático de matrícula porque matrículas antigas aparecem ligadas a nomes diferentes em datas diferentes.

A solução REV40 é controlada:

- chave de origem = `matrícula antiga + nome antigo`;
- usuário escolhe o funcionário atual;
- associação é persistida em `tbConfig`;
- arquivo é revalidado imediatamente;
- auditoria `DEPARA_FUNCIONARIO` é criada;
- um de/para que provocar pessoa acima de 1 DH ou ausência + trabalho pode transformar a importação em **BLOQUEADA**.

Foi executado teste sintético confirmando que o alias substitui a matrícula/nome pelo cadastro canônico e remove o alerta específico. O mesmo teste, ao apontar propositalmente para uma pessoa incompatível no dia, fez a validação bloquear por conflito de DH, comprovando a proteção cruzada.

## 8. FRENTE_ID

Regra final:

`obra + pacote + pavimento + local/frente`

Uma frente real só é gerada quando existem:

- obra;
- pacote;
- e pelo menos `pavimento` ou `local/frente`.

Se a localização não for suficiente:

- o apontamento continua nos indicadores de produção;
- aparece como `LOCALIZAÇÃO PENDENTE`;
- não recebe vínculo automático com FVS, Medição ou Ferramentas.

Isso evita cruzamentos falsos.

## 9. Offline-first

### Shell

- Cache Storage para HTML/JS/CSS/manifest/ícones;
- cache final: `elevatta-erp-v40-r3`;
- abertura do shell pelo aparelho;
- atualização em paralelo quando houver rede.

### Dados

- IndexedDB para dados, anexos e outbox;
- solicitação de armazenamento persistente quando o navegador permitir;
- indicação de uso/quota no painel administrativo;
- nenhum registro é considerado protegido antes da conclusão da escrita local.

### Sincronização

- evento `online`;
- reabertura/retorno ao app;
- sincronização manual opcional;
- Background Sync quando suportado;
- Service Worker processa a outbox ERP;
- `flush-lock` compartilhado impede página e Service Worker de enviarem a mesma transação em paralelo;
- retentativa com backoff para falhas transitórias;
- conflitos permanentes ficam bloqueados para revisão, sem descarte silencioso.

## 10. Ferramentas / Almoxarifado

### Cadastro

Existem slots `FER-001` a `FER-050`.

Enquanto a relação real não for informada:

- `NOME = A cadastrar`;
- `STATUS = CADASTRO PENDENTE`;
- `ATIVO = NAO`.

Assim nenhum patrimônio fictício pode ser retirado.

### Saída

Validação obrigatória no cliente e novamente no servidor:

- funcionário;
- obra;
- pacote;
- local/frente;
- responsável do almoxarifado;
- ferramenta ativa e disponível;
- assinatura;
- IDs de ferramenta sem duplicidade.

A assinatura vira PNG no armazenamento de arquivos e o Excel guarda URL + hash.

### Devolução

O Gateway valida:

- termo existente;
- termo ainda aberto;
- ferramenta realmente retirada naquele termo;
- ausência de devolução duplicada.

Registra condição, responsável por receber e foto quando informada.

### Conflito entre celulares offline

Dois celulares completamente offline podem registrar a mesma ferramenta antes de conhecerem o movimento um do outro. Ao retornar a rede, o Gateway aceita a primeira transação válida e rejeita a segunda com conflito. A pendência rejeitada fica visível em Administração.

Para reduzir conflitos de campo, continua recomendável um aparelho principal de almoxarifado por obra durante períodos totalmente sem sinal.

## 11. Livro razão / termo

No desktop:

- histórico respeita filtros globais;
- termo abre dentro do ERP;
- mostra funcionário, obra, pacote/frente, almoxarife e situação;
- mostra assinatura sincronizada;
- se a assinatura ainda estiver apenas offline, informa que está protegida no aparelho;
- mostra todos os eventos de saída/devolução;
- mostra fotos de devolução quando disponíveis.

A consolidação por frente permite abrir o termo da ferramenta diretamente.

## 12. FVS, fotos e consolidação

A mesma `FRENTE_ID` relaciona produção, FVS, Medição e Ferramentas.

Regra de conferência:

- sem FVS => `NÃO`;
- FVS aprovada e sem NC/pendência => `SIM`;
- FVS com NC => `NC / NÃO`;
- FVS pendente => `PENDÊNCIA`.

O detalhe da frente abre na mesma página e apresenta:

- produção;
- pessoas;
- FVS e checklist;
- situação da qualidade;
- fotografias;
- medição física;
- histórico de ferramentas.

## 13. Proxy autenticado de arquivos

Foi corrigido um problema importante: `webUrl` do SharePoint não é necessariamente uma URL de imagem apropriada para `<img>`.

A REV40 usa `/api/erp/file` como proxy autenticado e same-origin:

- navegador não recebe token Graph para buscar o arquivo;
- novas fotos/assinaturas gravam `appUrl` no ERP;
- URLs antigas do SharePoint são convertidas para o proxy na leitura do desktop;
- FVS sincronizada pelo fluxo legado também passa a gravar o endereço interno compatível.

## 14. Concorrência e Microsoft 365

Implementado:

- sessão de workbook reutilizável quando disponível;
- serialização de escrita na instância Node;
- inclusão em lotes de até 100 linhas;
- retentativa para timeout/429/502/503/504;
- upsert antes da remoção de sobras nos fluxos legados sensíveis;
- IDs imutáveis e idempotência por chave;
- importação deduplicada por SHA-256 e assinatura de linha.

## 15. Limitações que exigem homologação externa

### Microsoft 365 real

Não foi possível testar gravação real no tenant da empresa neste ambiente porque as credenciais Microsoft/SharePoint não estão disponíveis aqui.

Testar após deploy:

1. login;
2. leitura das 18 tabelas;
3. criação de uma retirada;
4. upload da assinatura;
5. devolução com foto;
6. importação de um arquivo aprovado;
7. FVS com foto;
8. abrir as imagens pelo desktop;
9. desligar/ligar a internet durante cada fluxo.

### Câmera física

A sintaxe e o ciclo de vida permanecem auditados, mas a estabilidade de câmera precisa ser confirmada em hardware Android/iPhone real com 10+ capturas consecutivas.

### Background Sync

O ERP não depende exclusivamente dele. Navegadores sem suporte sincronizam quando a página reaparece/volta a ficar online.

## 16. Pendências de negócio, não de código-base

1. relação real das ferramentas 01–50;
2. definição final dos perfis/permissões;
3. de/para dos funcionários históricos mostrado pelo módulo de validação;
4. homologação do tenant Microsoft 365 e dos celulares reais.
