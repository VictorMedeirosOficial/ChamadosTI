# HelpDesk TI • Policlínica Bernardo Félix da Silva (CPSMS)

Sistema oficial de abertura, acompanhamento e gestão Kanban de chamados técnicos de TI da **Policlínica Bernardo Félix da Silva** — **Consórcio Público de Saúde da Microrregião de Sobral (CPSMS)**.

## Arquitetura 100% Online e Gratuita
- **Front-end:** `index.html` hospedado via GitHub Pages.
- **Banco de Dados Online:** Planilha Google integrada via `Codigo.gs` (Google Apps Script).

## Como Ativar a Planilha Online
1. Crie uma planilha no Google Sheets.
2. Acesse **Extensões > Apps Script** e cole o conteúdo do arquivo `Codigo.gs`.
3. Clique em **Implantar > Nova implantação > App da Web** (Acesso: *Qualquer pessoa*).
4. Copie a URL `/exec` gerada, abra o site e clique no botão **Vincular Planilha Online** no topo direito (ou cole a URL na variável `DEFAULT_APPS_SCRIPT_URL` dentro do `index.html`).
