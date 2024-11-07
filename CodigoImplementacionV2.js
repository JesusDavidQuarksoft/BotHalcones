
var token = "Token";
function setWebhook() {
  var webAppUrl = 'https://script.google.com/macros/s/AKfycbz-pbJNd4nTv9Q1XkCXRyKnKP47_OawXD4b-1YxslqnOd7bDS_6_diu7yN0EM0DTRlO/exec';
  var response = UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/setWebhook?url=${webAppUrl}`);
  Logger.log(response);
}

function sendText(chatId, text) {
  UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/sendMessage`, {
    method: "post",
    payload: {
      chat_id: String(chatId),
      text: text,
      parse_mode: "HTML"
    }
  });
}

function findUserRow(sheet, column, value) {
  var data = sheet.getRange(2, column, sheet.getLastRow() - 1, 1).getValues().flat().map(String);
  var index = data.indexOf(String(value));
  return index !== -1 ? index + 2 : -1;
}

function validateUser(userId, rfc) {
  var sheet = SpreadsheetApp.openById("154X5o9_2X2oV78OibIT20AXh1cZrkNRUbDMBecCAwbc").getSheetByName("Lideres");
  var userRow = findUserRow(sheet, 1, userId);
  if (userRow !== -1) {
    return sheet.getRange(userRow, 5).getValue() === rfc; 
  }
  return false;
}

function doPost(e) {
  var contents = JSON.parse(e.postData.contents);
  if (contents.callback_query) return handleCallbackAction(contents.callback_query);

  var chatId = contents.message.chat.id;
  var messageText = contents.message.text;
  var id_message = contents.message.from.id;
  var firstName = contents.message.from.first_name;
  var username = contents.message.from.username;

  var sessionState = PropertiesService.getScriptProperties().getProperty('sessionState_' + chatId);

  if (messageText === '/start') {
    startSession(chatId);
  } else if (sessionState === 'awaiting_user_id') {
    handleUserId(chatId, messageText, id_message);
  } else if (sessionState === 'awaiting_password') {
    handlePassword(chatId, messageText);
  } else if (sessionState === 'awaiting_rfc') {
    handleRFC(chatId, messageText, id_message, firstName, username);
  } else {
    sendText(chatId, "Para empezar, por favor utiliza el comando /start.");
  }
}

function startSession(chatId) {
  sendText(chatId, "Hola!! 👋👋.\nPor favor ingresa tu ID de usuario ✍️✍️");
  PropertiesService.getScriptProperties().setProperty('sessionState_' + chatId, 'awaiting_user_id');
}

function handleUserId(chatId, userId, id_message) {
  var sheet = SpreadsheetApp.openById("154X5o9_2X2oV78OibIT20AXh1cZrkNRUbDMBecCAwbc").getSheetByName("Lideres");

  var telegramUserRow = findUserRow(sheet, 7, id_message); 
  if (telegramUserRow !== -1) {
    var associatedUserId = sheet.getRange(telegramUserRow, 1).getValue(); 
    if (String(associatedUserId) !== String(userId)) {
      sendText(chatId, "❌ ❌ Ya tienes una cuenta asociada a este Telegram. No puedes usar otro ID de usuario.");
      return;
    }
  }

  var userRow = findUserRow(sheet, 1, userId); 

  if (userRow !== -1) {
    var existingChatId = sheet.getRange(userRow, 7).getValue(); 
    if (existingChatId && existingChatId != chatId) {
      sendText(chatId, "❌ ❌ Este ID ya está asociado a otro usuario de Telegram. No puedes usarlo.");
    } else if (sheet.getRange(userRow, 6).getValue() === "ACTIVO") { 
      PropertiesService.getScriptProperties().setProperty('currentUserId_' + chatId, userId);
      PropertiesService.getScriptProperties().setProperty('expectedPassword_' + chatId, sheet.getRange(userRow, 3).getValue()); 
      sendText(chatId, "Por favor ingresa tu contraseña ✍️✍️.");
      PropertiesService.getScriptProperties().setProperty('sessionState_' + chatId, 'awaiting_password');
    } else {
      sendText(chatId, "❌ ❌ Tu cuenta está inactiva. Por favor contacta al administrador para más información.");
    }
  } else {
    sendText(chatId, "❌ ❌ El ID proporcionado no existe. Por favor, regístrate primero.");
  }
}

function handlePassword(chatId, password) {
  var expectedPassword = PropertiesService.getScriptProperties().getProperty('expectedPassword_' + chatId);
  if (password === expectedPassword) {
    sendText(chatId, "Contraseña validada correctamente ✅✅.\nPor favor ingresa tu RFC ✍️✍️.");
    PropertiesService.getScriptProperties().setProperty('sessionState_' + chatId, 'awaiting_rfc');
  } else {
    sendText(chatId, "❌ ❌ La contraseña proporcionada es incorrecta. Por favor, inténtalo nuevamente.");
  }
}

function handleRFC(chatId, rfc, id_message, firstName, username) {
  var currentUserId = PropertiesService.getScriptProperties().getProperty('currentUserId_' + chatId);
  if (validateUser(currentUserId, rfc)) {
    var sheet = SpreadsheetApp.openById("154X5o9_2X2oV78OibIT20AXh1cZrkNRUbDMBecCAwbc").getSheetByName("Lideres");
    var userRow = findUserRow(sheet, 1, currentUserId); 

    sheet.getRange(userRow, 7).setValue(id_message); 
    sheet.getRange(userRow, 8).setValue(firstName); 
    sheet.getRange(userRow, 9).setValue(username); 

    sendText(chatId, "RFC validado correctamente ✅.\nBienvenido al menú principal 👋👋.");
    sendInitialMenu(chatId);
    PropertiesService.getScriptProperties().deleteAllProperties();
  } else {
    sendText(chatId, "❌ ❌ El RFC proporcionado no coincide con el registrado para este usuario. Por favor, inténtalo nuevamente.");
  }
}

function sendInitialMenu(chatId) {
  var sheet = SpreadsheetApp.openById("154X5o9_2X2oV78OibIT20AXh1cZrkNRUbDMBecCAwbc").getSheetByName("Carpetas");
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  var filteredData = data.filter(row => row[1]);

  var keyboard = {
    inline_keyboard: filteredData.map(row => [{ text: `📂 Descargar último reporte de ${row[0]}`, callback_data: row[0] }])
  };

  UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/sendMessage`, {
    method: "post",
    payload: {
      chat_id: String(chatId),
      text: "Por favor selecciona el tipo de reporte que deseas descargar:",
      parse_mode: "HTML",
      reply_markup: JSON.stringify(keyboard)
    }
  });
}


function handleCallbackAction(callbackQuery) {
  var chatId = callbackQuery.message.chat.id;
  var callbackData = callbackQuery.data;

  var sheet = SpreadsheetApp.openById("154X5o9_2X2oV78OibIT20AXh1cZrkNRUbDMBecCAwbc").getSheetByName("Lideres");
  var telegramUserRow = findUserRow(sheet, 7, chatId); 

  if (telegramUserRow !== -1 && sheet.getRange(telegramUserRow, 6).getValue() === "ACTIVO") { 
    var folderId;
    var option = getOptionByCallbackData(callbackData);

    if (option) {
      folderId = option.folderId;
      sendText(chatId, `🔍 Opción elegida: Descargar último reporte de ${option.name}.\n`);
      PropertiesService.getScriptProperties().setProperty('selectedFolder_' + chatId, folderId);
      sendFileTypeMenu(chatId);
    } else if (callbackData === 'option_pdf' || callbackData === 'option_excel') {
      folderId = PropertiesService.getScriptProperties().getProperty('selectedFolder_' + chatId);
      var userId = sheet.getRange(telegramUserRow, 1).getValue(); 
      if (callbackData === 'option_pdf') {
        sendText(chatId, "🔍 Opción elegida: Descargar el último reporte en PDF.\nDescargando reporte.... 📥📥📥");
        searchAndSendLatestPDF(chatId, folderId, userId);
      } else if (callbackData === 'option_excel') {
        sendText(chatId, "🔍 Opción elegida: Descargar el último reporte en Excel.\nDescargando reporte.... 📥📥📥");
        searchAndSendSheetFile(chatId, folderId, userId);
      }
    }
  } else {
    sendText(chatId, "❌ Tu cuenta está inactiva. Por favor contacta al administrador para más información.");
  }

  UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/answerCallbackQuery?callback_query_id=${callbackQuery.id}`);
}

function getOptionByCallbackData(callbackData) {
  var sheet = SpreadsheetApp.openById("154X5o9_2X2oV78OibIT20AXh1cZrkNRUbDMBecCAwbc").getSheetByName("Carpetas");
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) { 
    var folderName = data[i][0];
    var folderUrl = data[i][1];
    var folderId = extractFolderIdFromUrl(folderUrl);

    if (folderName.toLowerCase() === callbackData.toLowerCase().replace('folder_', '')) {
      return {
        name: folderName,
        folderId: folderId
      };
    }
  }
  return null;
}

function extractFolderIdFromUrl(url) {
  var match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

function sendFileTypeMenu(chatId) {
  var keyboard = {
    inline_keyboard: [
      [{ text: "📄 Descargar el último reporte como PDF", callback_data: "option_pdf" }],
      [{ text: "📊 Descargar el último reporte como Excel", callback_data: "option_excel" }]
    ]
  };

  UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/sendMessage`, {
    method: "post",
    payload: {
      chat_id: String(chatId),
      text: "Por favor selecciona el formato del reporte que desea descargar:",
      parse_mode: "HTML",
      reply_markup: JSON.stringify(keyboard)
    }
  });
}



function convertToPDF(file) {
  var fileId = file.getId();
  var url = 'https://docs.google.com/spreadsheets/d/' + fileId + '/export?format=pdf';

  var response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
    }
  });

  return response.getBlob().setName(file.getName().replace(/\.[^/.]+$/, ".pdf"));
}

function sendDocument(chatId, fileBlob, fileName) {
  var formData = {
    method: "post",
    payload: {
      chat_id: String(chatId),
      caption: "Aquí está el último reporte en PDF ✅✅.",
      document: fileBlob,
      filename: fileName
    }
  };
  UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/sendDocument`, formData);

    Utilities.sleep(7000); // Mandar el menu 7 segundos despues

    sendInitialMenu(chatId);

}


function sendSheetFile(chatId, file) {
  var url = `https://drive.google.com/open?id=${file.getId()}`;

  UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/sendMessage`, {
    method: "post",
    payload: {
      chat_id: String(chatId),
      text: "Aquí está el último reporte de Excel ✅✅.\n[Descargar archivo](" + url + ")",
      parse_mode: "Markdown"
    }
  });
  

    Utilities.sleep(7000); // Mandar el menu 7 segundos despues

    sendInitialMenu(chatId);

  }



function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  if (sheet.getName() === "Carpetas") {
    updateTelegramMenus(); 
  }
}


function updateTelegramMenus() {
  var sheetUsuarios = SpreadsheetApp.openById("154X5o9_2X2oV78OibIT20AXh1cZrkNRUbDMBecCAwbc").getSheetByName("Lideres");
  var sheetCarpetas = SpreadsheetApp.openById("154X5o9_2X2oV78OibIT20AXh1cZrkNRUbDMBecCAwbc").getSheetByName("Carpetas");
  
  var carpetasData = sheetCarpetas.getRange(2, 1, sheetCarpetas.getLastRow() - 1, 2).getValues();
  
  var carpetasValidas = carpetasData.filter(row => row[0] && row[1]);

  if (carpetasValidas.length > 0) {
    var usuariosData = sheetUsuarios.getRange(2, 7, sheetUsuarios.getLastRow() - 1, 1).getValues().flat();

    usuariosData.forEach(function(chatId) {
      if (chatId) {
        sendInitialMenu(chatId);  
      }
    });
  } else {
    Logger.log("No se encontraron carpetas válidas con ambas columnas llenas para actualizar los menús.");
  }
  
}
function searchAndSendSheetFile(chatId, folderId, userId) {
  var files = DriveApp.getFolderById(folderId).searchFiles(`title contains "${userId}"`);
  var latestFile = null;
  var folderName = getFolderNameById(folderId); 
  
  while (files.hasNext()) {
    var file = files.next();
    if (!latestFile || file.getLastUpdated() > latestFile.getLastUpdated()) {
      latestFile = file;
    }
  }

  if (latestFile) {
    sendSheetFile(chatId, latestFile);

    registerDownload(userId, "username_placeholder", latestFile.getName(), "Excel", folderName, folderId);
  } else {
    sendText(chatId, "No se encontraron reportes asociados a tu usuario ❌❌.");
  }


}

function searchAndSendLatestPDF(chatId, folderId, userId) {
  var files = DriveApp.getFolderById(folderId).searchFiles(`title contains "${userId}"`);
  var latestFile = null;
  var folderName = getFolderNameById(folderId);
  
  while (files.hasNext()) {
    var file = files.next();
    if (!latestFile || file.getLastUpdated() > latestFile.getLastUpdated()) {
      latestFile = file;
    }
  }

  if (latestFile) {
    var pdfBlob = convertToPDF(latestFile);
    sendDocument(chatId, pdfBlob, latestFile.getName().replace(/\.[^/.]+$/, ".pdf"));

    registerDownload(userId, "username_placeholder", latestFile.getName(), "PDF", folderName, folderId);
  } else {
    sendText(chatId, "No se encontraron reportes asociados a tu usuario ❌❌.");
  }
}

// registrar datos de la descarga en la hoja
function registerDownload(userId, userName, fileName, folderName, format) {
  var sheetUsuarios = SpreadsheetApp.openById("154X5o9_2X2oV78OibIT20AXh1cZrkNRUbDMBecCAwbc").getSheetByName("Lideres");
  var sheetRegistros = SpreadsheetApp.openById("InsertarLinkParaRegistrarClicks").getSheetByName("Registros_BotHalcones_1");
  var timestamp = new Date();

  var userRow = findUserRow(sheetUsuarios, 1, userId);
  var userName = userRow !== -1 ? sheetUsuarios.getRange(userRow, 2).getValue() : "Sin nombre";

  sheetRegistros.appendRow([userId, userName, fileName, format, folderName, timestamp]);
}

function getFolderNameById(folderId) {
  var sheetCarpetas = SpreadsheetApp.openById("154X5o9_2X2oV78OibIT20AXh1cZrkNRUbDMBecCAwbc").getSheetByName("Carpetas");
  var data = sheetCarpetas.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][1].includes(folderId)) {
      return data[i][0]; 
    }
  }
}



function registerDownload(userId, userName, fileName, format, folderName) {
  const sheetUsuarios = SpreadsheetApp.openById("154X5o9_2X2oV78OibIT20AXh1cZrkNRUbDMBecCAwbc").getSheetByName("Lideres");
  const sheetRegistros = getRegistroSheet(); 
  const timestamp = new Date();
  const date = timestamp.toISOString().slice(0, 10); 
  const time = timestamp.toTimeString().split(' ')[0]; 

  const userRow = findUserRow(sheetUsuarios, 1, userId);
  if (userRow !== -1) {
    userName = sheetUsuarios.getRange(userRow, 2).getValue();
  } else {
    userName = "Sin nombre"; 
  }
  sheetRegistros.appendRow([userId, userName, fileName, folderName, format, date, time]);
}

function getRegistroSheet() {
  const spreadsheet = SpreadsheetApp.openById("154X5o9_2X2oV78OibIT20AXh1cZrkNRUbDMBecCAwbc");
  const sheet = spreadsheet.getSheetByName("Registros_BotHalcones");
    if (!sheet) {
    const newSheet = spreadsheet.insertSheet("Registros_BotHalcones");
    const headers = ["User ID", "User Name", "File Name", "Folder Name", "Format", "Date", "Time"];
    newSheet.appendRow(headers);
    return newSheet;
  }
  
  return sheet;
}

function findUserRow(sheet, column, value) {
  const range = sheet.getRange(1, column, sheet.getLastRow());
  const values = range.getValues();
  
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === value) {
      return i + 1; 
    }
  }
  return -1;
}
