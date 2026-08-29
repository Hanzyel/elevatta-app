# Microsoft 365 / SharePoint — ELEVATTA ERP REV40.1

## 1. Arquitetura

`PWA/ERP → armazenamento offline → Data Gateway Node → Microsoft Graph → SharePoint/Excel`

Os módulos operacionais legados mantêm compatibilidade com a fila Microsoft existente, enquanto os novos módulos ERP usam o Data Gateway. Nenhum `client secret` é exposto no navegador.

## 2. App Registration no Microsoft Entra

Crie/reutilize um App Registration e configure:

- plataforma: **Web**;
- Redirect URI: `https://SEU-APP.onrender.com/auth/callback`;
- permissões delegadas: `User.Read` e `Files.ReadWrite`;
- Client Secret;
- `offline_access` já é solicitado no fluxo OAuth.

## 3. Variáveis no Render

Configure:

- `M365_TENANT_ID`
- `M365_CLIENT_ID`
- `M365_CLIENT_SECRET`
- `M365_WORKBOOK_URL`
- `SESSION_SECRET`
- `APP_BASE_URL` = `https://SEU-APP.onrender.com`

Não coloque segredos em GitHub, `index.html`, `erp.js`, `m365.js` ou outro arquivo público.

## 4. Base SharePoint

Faça upload de `BASE_ELEVATTA_SHAREPOINT.xlsx` e preserve os nomes das tabelas estruturadas:

- `tbApontamentos`
- `tbFvsRegistros`
- `tbFvsItens`
- `tbMedicoes`
- `tbFotos`
- `tbPacotes`
- `tbFuncionarios`
- `tbConfig`
- `tbObras`
- `tbEquipes`
- `tbFrentesServico`
- `tbFerramentas`
- `tbTermosFerramentas`
- `tbMovimentosFerramentas`
- `tbImportacoes`
- `tbImportacaoOcorrencias`
- `tbAuditoria`
- `tbUsuarios`

Copie o link completo do arquivo para `M365_WORKBOOK_URL`.

## 5. Pastas de arquivos

O Data Gateway cria a árvore `ELEVATTA_ERP` ao lado da planilha para arquivos dos novos módulos. Os módulos FVS existentes também preservam sua estrutura de fotos compatível.

Exemplos:

- `ELEVATTA_ERP/Ferramentas/Catalogo/`
- `ELEVATTA_ERP/Ferramentas/Devolucoes/<TERMO>/`
- `ELEVATTA_ERP/Assinaturas/Ferramentas/<TERMO>/`
- `ELEVATTA_ERP/Importacoes/<IMPORTACAO>/`

## 6. Primeiro uso

1. Faça o deploy no Render.
2. Abra o Elevatta com internet.
3. Faça o login Microsoft uma vez.
4. Acesse `/erp.html` e confirme o status.
5. Abra o PWA operacional e conclua o cache inicial.
6. Instale na tela inicial dos celulares usados no canteiro.

Depois disso, o trabalho pode continuar offline. Registros aguardam na outbox e são sincronizados quando a conexão e a sessão Microsoft estiverem disponíveis.

## 7. Conflitos

Um erro de rede é tentado novamente automaticamente. Um conflito permanente, como retirada de uma ferramenta que outro aparelho já retirou, fica marcado como **CONFLITO** em `ERP → Administração`.

Fotos e assinaturas exibidas pelo ERP passam pelo endpoint autenticado `/api/erp/file`; o navegador não precisa receber credencial Graph para carregar as evidências.

O administrador pode:

- tentar novamente depois de resolver a causa;
- descartar a pendência conscientemente.

Nada é descartado silenciosamente.

## 8. Observação sobre múltiplos aparelhos offline

Dois aparelhos totalmente offline não têm como saber em tempo real que a mesma ferramenta foi movimentada no outro. A proteção definitiva ocorre no Gateway ao reconectar. Para reduzir conflitos de campo, recomenda-se um dispositivo principal de almoxarifado por obra quando estiver sem rede.

## 9. Validação antes de gravar importações

`BLOQUEADO` e qualquer importação com `CRITICOS > 0` são recusados também pelo servidor. Corrija o arquivo e valide novamente antes de enviar. De/para de funcionário é gravado em `tbConfig` e auditado.
