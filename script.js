const FOLDER_ID = '1W_e52ISxXuxjIFMxTuQHJX0urBgb1L2n';

function doPost(e) {
  // Trata teste local ou erro de requisição
  if (!e || !e.postData || !e.postData.contents) {
    return ContentService.createTextOutput(JSON.stringify({ 
      status: 'error', 
      message: 'Requisição vazia ou teste local detectado.' 
    })).setMimeType(ContentService.MimeType.JSON);
  }

  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (parseError) {
    return ContentService.createTextOutput(JSON.stringify({ 
      status: 'error', 
      message: 'JSON inválido: ' + parseError.message 
    })).setMimeType(ContentService.MimeType.JSON);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Capturas") || ss.insertSheet("Capturas");
  
  // Garante cabeçalhos se a aba for nova
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Nome", "Data/Hora", "Link da Imagem", "Status"]);
  }

  const timestamp = new Date(data.timestamp);
  const formattedDate = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
  
  let driveLink = "";
  let statusMsg = "Sucesso";

  // Processa a imagem se existir
  if (data.imageBase64) {
    try {
      const base64Data = data.imageBase64.split(',')[1];
      const imageBlob = Utilities.newBlob(
        Utilities.base64Decode(base64Data),
        'image/jpeg',
        `captura_${Date.now()}_${data.name.replace(/\s+/g, '_')}.jpg`
      );

      const folder = DriveApp.getFolderById(FOLDER_ID);
      const file = folder.createFile(imageBlob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      driveLink = file.getUrl();
    } catch (driveError) {
      statusMsg = "Erro no Drive: " + driveError.message;
    }
  } else {
    statusMsg = "Sem imagem recebida";
  }

  sheet.appendRow([
    data.name || "Anônimo",
    formattedDate,
    driveLink,
    statusMsg
  ]);

  return ContentService.createTextOutput(JSON.stringify({ 
    status: 'success', 
    message: 'Registro salvo',
    driveLink: driveLink
  })).setMimeType(ContentService.MimeType.JSON);
}
