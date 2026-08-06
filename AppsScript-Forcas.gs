/**
 * Backend do Google Apps Script para a pesquisa Forças de Caráter.
 *
 * ATIVAÇÃO:
 * 1. Cole este arquivo em um projeto do Google Apps Script.
 * 2. Implante como "Aplicativo da Web".
 * 3. Execute como: você.
 * 4. Quem pode acessar: qualquer pessoa.
 * 5. Se a URL da implantação mudar, atualize WEB_APP_URL no index.html.
 */

const CONFIG = Object.freeze({
  spreadsheetId: "1CrQJx-M-u5F9G6RQ7X6fThwaF5OU2gcRPlkrRRp8WPQ",
  sheetName: "Mapa Impulso",
  apiKey: "2026Impulso$",
  adminEmail: "impulsoflow@gmail.com",
  senderName: "Instituto Impulso Coaching de Liderança",
  maxPdfBase64Length: 8500000
});

const HEADERS = [
  "Recebido em",
  "Respondido em",
  "Consentimento em",
  "Instrumento",
  "Nome",
  "E-mail",
  "Top 5",
  "Ranking completo (JSON)",
  "WhatsApp",
  "Idade",
  "Profissão",
  "Escolaridade",
  "Arquivo PDF",
  "E-mail enviado",
  "Erro",
  "Submission ID"
];

function doGet() {
  return jsonResponse_({
    ok: true,
    service: "Forças de Caráter",
    version: "2026-08-06"
  });
}

function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    const payload = parsePayload_(e);
    validatePayload_(payload);

    const spreadsheet = SpreadsheetApp.openById(CONFIG.spreadsheetId);
    const sheet = spreadsheet.getSheetByName(CONFIG.sheetName) ||
      spreadsheet.insertSheet(CONFIG.sheetName);

    ensureHeaders_(sheet);

    const submissionId = clean_(payload.submissionId) || Utilities.getUuid();
    const top5 = getTop5_(payload);
    const ranking = getRanking_(payload);
    const existingRow = findSubmissionRow_(sheet, submissionId);

    let row = existingRow;
    if (!row) {
      sheet.appendRow([
        new Date(),
        parseDateOrNow_(payload.respondidoEm),
        clean_(payload.consentimentoEm),
        clean_(payload.instrumento) || "Forças de Caráter",
        clean_(payload.nome),
        clean_(payload.email).toLowerCase(),
        top5.join(" | "),
        JSON.stringify(ranking),
        clean_(payload.whatsapp),
        clean_(payload.idade),
        clean_(payload.profissao),
        clean_(payload.escolaridade),
        clean_(payload.pdfFileName),
        "PENDENTE",
        "",
        submissionId
      ]);
      row = sheet.getLastRow();
      SpreadsheetApp.flush();
    }

    if (String(sheet.getRange(row, 14).getValue()).toUpperCase() === "SIM") {
      return jsonResponse_({
        ok: true,
        saved: true,
        emailSent: true,
        duplicate: true,
        submissionId: submissionId,
        row: row
      });
    }

    try {
      sendClientReport_(payload, top5);
      sheet.getRange(row, 14, 1, 2).setValues([["SIM", ""]]);
      SpreadsheetApp.flush();

      notifyAdmin_(payload, top5, submissionId);

      return jsonResponse_({
        ok: true,
        saved: true,
        emailSent: true,
        duplicate: Boolean(existingRow),
        submissionId: submissionId,
        row: row
      });
    } catch (emailError) {
      const message = errorMessage_(emailError);
      sheet.getRange(row, 14, 1, 2).setValues([["NÃO", message]]);
      SpreadsheetApp.flush();

      return jsonResponse_({
        ok: false,
        saved: true,
        emailSent: false,
        submissionId: submissionId,
        row: row,
        message: "O resultado foi salvo, mas o e-mail não pôde ser enviado: " + message
      });
    }
  } catch (error) {
    return jsonResponse_({
      ok: false,
      saved: false,
      emailSent: false,
      message: errorMessage_(error)
    });
  } finally {
    try {
      lock.releaseLock();
    } catch (ignore) {}
  }
}

function parsePayload_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error("Requisição sem conteúdo.");
  }

  try {
    return JSON.parse(e.postData.contents);
  } catch (error) {
    throw new Error("JSON inválido.");
  }
}

function validatePayload_(payload) {
  if (!payload || clean_(payload.apiKey) !== CONFIG.apiKey) {
    throw new Error("Chave de acesso inválida.");
  }

  if (!clean_(payload.nome)) {
    throw new Error("Nome não informado.");
  }

  const email = clean_(payload.email).toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    throw new Error("E-mail inválido.");
  }

  if (payload.sendEmail !== true) {
    throw new Error("O envio por e-mail não foi autorizado pela requisição.");
  }

  if (!clean_(payload.pdfBase64)) {
    throw new Error("PDF não recebido.");
  }

  if (clean_(payload.pdfBase64).length > CONFIG.maxPdfBase64Length) {
    throw new Error("PDF excede o tamanho máximo permitido.");
  }
}

function ensureHeaders_(sheet) {
  const current = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const needsUpdate = HEADERS.some(function(header, index) {
    return String(current[index] || "").trim() !== header;
  });

  if (needsUpdate) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight("bold")
      .setBackground("#0b63ce")
      .setFontColor("#ffffff");
  }
}

function findSubmissionRow_(sheet, submissionId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const match = sheet
    .getRange(2, 16, lastRow - 1, 1)
    .createTextFinder(submissionId)
    .matchEntireCell(true)
    .findNext();

  return match ? match.getRow() : 0;
}

function getTop5_(payload) {
  if (Array.isArray(payload.top5Lista)) {
    return payload.top5Lista.slice(0, 5).map(clean_).filter(Boolean);
  }

  return [
    payload.top1,
    payload.top2,
    payload.top3,
    payload.top4,
    payload.top5
  ].map(clean_).filter(Boolean);
}

function getRanking_(payload) {
  if (Array.isArray(payload.rankingCompleto)) {
    return payload.rankingCompleto;
  }

  if (clean_(payload.rankingCompletoJson)) {
    try {
      const parsed = JSON.parse(payload.rankingCompletoJson);
      if (Array.isArray(parsed)) return parsed;
    } catch (ignore) {}
  }

  const names = [
    payload.top1,
    payload.top2,
    payload.top3,
    payload.top4,
    payload.top5,
    payload.bot1,
    payload.bot2,
    payload.bot3,
    payload.bot4,
    payload.bot5
  ].map(clean_).filter(Boolean);

  return names.map(function(name, index) {
    return { posicao: index + 1, nome: name };
  });
}

function sendClientReport_(payload, top5) {
  const rawBase64 = clean_(payload.pdfBase64).replace(/^data:application\/pdf;base64,/, "");
  const bytes = Utilities.base64Decode(rawBase64);
  const fileName = sanitizeFilename_(payload.pdfFileName || "Forcas_de_Carater.pdf");
  const pdfBlob = Utilities.newBlob(bytes, MimeType.PDF, fileName);

  const clientName = clean_(payload.nome);
  const subject = "Seu resultado — Forças de Caráter";
  const plainText =
    "Olá, " + clientName + "!\n\n" +
    "Seu relatório da pesquisa Forças de Caráter está anexado a este e-mail.\n\n" +
    "Suas cinco forças principais: " + top5.join(", ") + ".\n\n" +
    "Instituto Impulso Coaching de Liderança";

  const htmlBody =
    '<div style="font-family:Arial,sans-serif;color:#0f172a;line-height:1.6">' +
      '<h2 style="color:#063b7a">Olá, ' + escapeHtml_(clientName) + '!</h2>' +
      '<p>Seu relatório da pesquisa <strong>Forças de Caráter</strong> está anexado.</p>' +
      '<p><strong>Suas cinco forças principais:</strong><br>' +
        escapeHtml_(top5.join(" • ")) +
      '</p>' +
      '<p>Use este material como apoio para reconhecer e aplicar seus recursos mais naturais.</p>' +
      '<p style="color:#475569">Instituto Impulso Coaching de Liderança<br>' +
      'A mudança pode acontecer em um instante.</p>' +
    '</div>';

  GmailApp.sendEmail(
    clean_(payload.email).toLowerCase(),
    subject,
    plainText,
    {
      htmlBody: htmlBody,
      attachments: [pdfBlob],
      name: CONFIG.senderName
    }
  );
}

function notifyAdmin_(payload, top5, submissionId) {
  try {
    const subject = "Novo resultado — Forças de Caráter";
    const body =
      "Participante: " + clean_(payload.nome) + "\n" +
      "E-mail: " + clean_(payload.email).toLowerCase() + "\n" +
      "WhatsApp: " + clean_(payload.whatsapp) + "\n" +
      "Top 5: " + top5.join(", ") + "\n" +
      "Submission ID: " + submissionId;

    GmailApp.sendEmail(CONFIG.adminEmail, subject, body, {
      name: CONFIG.senderName
    });
  } catch (ignore) {
    console.log("Aviso administrativo não enviado: " + errorMessage_(ignore));
  }
}

function parseDateOrNow_(value) {
  if (!clean_(value)) return new Date();
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? new Date() : parsed;
}

function sanitizeFilename_(value) {
  const name = clean_(value)
    .replace(/[\\/:*?"<>|]/g, "_")
    .replace(/\s+/g, "_");

  return /\.pdf$/i.test(name) ? name : (name || "Forcas_de_Carater") + ".pdf";
}

function clean_(value) {
  return String(value === null || value === undefined ? "" : value).trim();
}

function escapeHtml_(value) {
  return clean_(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function errorMessage_(error) {
  return error && error.message ? String(error.message) : String(error || "Erro desconhecido.");
}

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
