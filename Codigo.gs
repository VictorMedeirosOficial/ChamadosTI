/**
 * BANCO DE DADOS ONLINE - HELPDESK TI (POLICLÍNICA BERNARDO FÉLIX DA SILVA - CPSMS)
 * Criado por Victor Medeiros - Recepção Geral / Policlínica Sobral
 * Cole este código em: Extensões > Apps Script na sua Planilha Google
 */

const ABA_CHAMADOS = "Chamados";
const ABA_MONITORAMENTO = "Monitoramento";
const ABA_EQUIPE = "Equipe_TI";

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('HelpDesk TI - Policlínica')
    .addItem('Inicializar Estrutura da Planilha', 'setupPlanilha')
    .addToUi();
}

function setupPlanilha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Aba Chamados
  let sheetChamados = ss.getSheetByName(ABA_CHAMADOS);
  if (!sheetChamados) sheetChamados = ss.insertSheet(ABA_CHAMADOS);
  const colsChamados = [
    "id", "protocolo", "data_ocorrencia", "hora_ocorrencia", "solicitante",
    "setor", "ocorrencia", "status", "tecnico_responsavel",
    "solucao_aplicada", "data_conclusao", "imagem_base64"
  ];
  if (sheetChamados.getLastRow() === 0) {
    sheetChamados.appendRow(colsChamados);
    sheetChamados.getRange(1, 1, 1, colsChamados.length)
      .setBackground("#0F6C2B").setFontColor("#FFFFFF").setFontWeight("bold");
    sheetChamados.setFrozenRows(1);
  }

  // 2. Aba Monitoramento (Exclusivo TI)
  let sheetMon = ss.getSheetByName(ABA_MONITORAMENTO);
  if (!sheetMon) sheetMon = ss.insertSheet(ABA_MONITORAMENTO);
  const colsMon = ["id", "data_registro", "titulo", "local_ativo", "observacao", "imagem_base64"];
  if (sheetMon.getLastRow() === 0) {
    sheetMon.appendRow(colsMon);
    sheetMon.getRange(1, 1, 1, colsMon.length)
      .setBackground("#0F6C2B").setFontColor("#FFFFFF").setFontWeight("bold");
    sheetMon.setFrozenRows(1);
  }

  // 3. Aba Equipe_TI (Ajudantes e Técnicos - Exclusivo TI)
  let sheetEq = ss.getSheetByName(ABA_EQUIPE);
  if (!sheetEq) sheetEq = ss.insertSheet(ABA_EQUIPE);
  const colsEq = ["id", "nome", "funcao", "setor_origem", "contato", "status", "foto_base64"];
  if (sheetEq.getLastRow() === 0) {
    sheetEq.appendRow(colsEq);
    sheetEq.getRange(1, 1, 1, colsEq.length)
      .setBackground("#0F6C2B").setFontColor("#FFFFFF").setFontWeight("bold");
    sheetEq.setFrozenRows(1);
    sheetEq.appendRow(["EQ-01", "João Victor", "Analista / Técnico de TI", "TI Geral", "Ramal TI", "Ativo", ""]);
    sheetEq.appendRow(["EQ-02", "Victor Medeiros", "Apoio Técnico / Ouvidoria & Recepção", "Recepção Geral", "Ramal Recepção", "Ativo", ""]);
  }
}

function doGet(e) {
  setupPlanilha();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return createJsonResponse({
    status: "success",
    chamados: lerAba(ss.getSheetByName(ABA_CHAMADOS)),
    monitoramentos: lerAba(ss.getSheetByName(ABA_MONITORAMENTO)),
    equipe: lerAba(ss.getSheetByName(ABA_EQUIPE))
  });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(15000);

  try {
    setupPlanilha();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const body = JSON.parse(e.postData.contents);
    const entity = body.entity || "chamado";
    const action = body.action || "create";

    // === ENTIDADE: CHAMADOS ===
    if (entity === "chamado") {
      const sheet = ss.getSheetByName(ABA_CHAMADOS);
      if (action === "create") {
        const agora = new Date();
        const id = body.id || ("CH-" + agora.getTime().toString().slice(-6));
        const protocolo = body.protocolo || ("POLI-" + agora.getFullYear() + "-" + Math.floor(1000 + Math.random() * 9000));
        const dataStr = body.data_ocorrencia || Utilities.formatDate(agora, "America/Fortaleza", "dd/MM/yyyy");
        const horaStr = body.hora_ocorrencia || Utilities.formatDate(agora, "America/Fortaleza", "HH:mm:ss");

        sheet.appendRow([
          id, protocolo, dataStr, horaStr,
          body.solicitante || "", body.setor || "", body.ocorrencia || "",
          "Aberto", "Aguardando TI", "", "", body.imagem_base64 || ""
        ]);
        return createJsonResponse({ status: "success", id: id, protocolo: protocolo });
      }

      if (action === "update") {
        const data = sheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][0]) === String(body.id)) {
            const row = i + 1;
            if (body.status !== undefined) sheet.getRange(row, 8).setValue(body.status);
            if (body.tecnico_responsavel !== undefined) sheet.getRange(row, 9).setValue(body.tecnico_responsavel);
            if (body.solucao_aplicada !== undefined) sheet.getRange(row, 10).setValue(body.solucao_aplicada);
            if (body.status === "Concluído") {
              sheet.getRange(row, 11).setValue(Utilities.formatDate(new Date(), "America/Fortaleza", "dd/MM/yyyy HH:mm"));
            }
            return createJsonResponse({ status: "success" });
          }
        }
      }

      if (action === "delete") {
        return excluirLinhaPorId(sheet, body.id);
      }
    }

    // === ENTIDADE: MONITORAMENTO (EXCLUSIVO TI) ===
    if (entity === "monitoramento") {
      const sheet = ss.getSheetByName(ABA_MONITORAMENTO);
      if (action === "create") {
        const agora = new Date();
        const id = body.id || ("MON-" + agora.getTime().toString().slice(-6));
        const dataReg = body.data_registro || Utilities.formatDate(agora, "America/Fortaleza", "dd/MM/yyyy HH:mm");
        sheet.appendRow([
          id, dataReg, body.titulo || "", body.local_ativo || "",
          body.observacao || "", body.imagem_base64 || ""
        ]);
        return createJsonResponse({ status: "success", id: id });
      }
      if (action === "delete") {
        return excluirLinhaPorId(sheet, body.id);
      }
    }

    // === ENTIDADE: EQUIPE / AJUDANTES TI (EXCLUSIVO TI) ===
    if (entity === "equipe") {
      const sheet = ss.getSheetByName(ABA_EQUIPE);
      if (action === "create") {
        const id = body.id || ("EQ-" + new Date().getTime().toString().slice(-6));
        sheet.appendRow([
          id, body.nome || "", body.funcao || "", body.setor_origem || "",
          body.contato || "", body.status || "Ativo", body.foto_base64 || ""
        ]);
        return createJsonResponse({ status: "success", id: id });
      }
      if (action === "delete") {
        return excluirLinhaPorId(sheet, body.id);
      }
    }

    return createJsonResponse({ status: "error", message: "Ação inválida" });
  } catch (err) {
    return createJsonResponse({ status: "error", message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

function excluirLinhaPorId(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return createJsonResponse({ status: "success" });
    }
  }
  return createJsonResponse({ status: "error", message: "Registro não encontrado" });
}

function lerAba(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  const headers = values[0];
  return values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, idx) => obj[h] = row[idx]);
    return obj;
  });
}

function createJsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
