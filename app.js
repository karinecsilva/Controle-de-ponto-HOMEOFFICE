// Função principal para registrar o ponto
function registrarPonto(acao) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dataAtual = new Date();
  const dataFormatada = Utilities.formatDate(dataAtual, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
  const email = "meuemail@gmail.com"; // Substitua pelo seu email

  let ultimaLinha = sheet.getLastRow();

  // Se a última linha for uma nova data, adiciona uma nova linha
  if (ultimaLinha === 0 || sheet.getRange(ultimaLinha, 1).getValue() !== Utilities.formatDate(dataAtual, Session.getScriptTimeZone(), "dd/MM/yyyy")) {
    sheet.appendRow([Utilities.formatDate(dataAtual, Session.getScriptTimeZone(), "dd/MM/yyyy")]);
    ultimaLinha = sheet.getLastRow();
  }

  // Define as ações: início, pausa para almoço, retorno do almoço, finalizar expediente
  switch (acao) {
    case 'inicio':
      sheet.getRange(ultimaLinha, 2).setValue(dataFormatada);
      enviarEmail(email, "Início do Expediente", "O expediente foi iniciado às " + dataFormatada);
      break;
    case 'almoco':
      sheet.getRange(ultimaLinha, 3).setValue(dataFormatada);
      enviarEmail(email, "Pausa para Almoço", "A pausa para o almoço começou às " + dataFormatada);
      break;
    case 'retorno_almoco':
      sheet.getRange(ultimaLinha, 4).setValue(dataFormatada);
      enviarEmail(email, "Retorno do Almoço", "O retorno do almoço foi às " + dataFormatada);
      break;
    case 'fim':
      sheet.getRange(ultimaLinha, 5).setValue(dataFormatada);
      enviarEmail(email, "Final do Expediente", "O expediente foi finalizado às " + dataFormatada);
      break;
  }
}

// Função para enviar o email de notificação
function enviarEmail(email, assunto, corpo) {
  GmailApp.sendEmail(email, assunto, corpo);
}

// Funções para o frontend (interface HTML)
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle("Controle de Ponto")
      .setWidth(300)
      .setHeight(300);
}

function registrarAcao(acao) {
  registrarPonto(acao);
}
