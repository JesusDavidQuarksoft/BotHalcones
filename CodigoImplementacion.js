var token = "7445103715:AAHbJyyNgvZHN1KWiLWVI-8qCwbtCwiFEp0";
var folderPosteoId = '1FBZdfhdH8uj2osZvKXh7Y8jwu8qHDa-k';
var folderDiarioId = '19ZJSj8cjZOrcbg5qSOPM5XegspZ1XF1D';
var folderPYMEId = '1IGx7rSiIZ5zRM6V6g-f7kxkxmoKk0nGw';
var folderCalidadId = '18U8WyEwnpS3PjieWwd9p3drz3Kk3vnOs';
var folderMuestreoId = '1BFUL6JPe1hKVfwE8-Z317Jn_BqpcZJnM'; 

function setWebhook() {
  var webAppUrl = 'https://script.google.com/macros/s/AKfycbwEwTsMVQYS2AHaCiDssC2u5z-cyr5kKJvERxrNoT7KRAw1nlZI0Hv_h8cNzRlf7_Jq/exec';
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
  var keyboard = {
    inline_keyboard: [
      [{ text: "📂 Descargar último reporte de Posteo", callback_data: "folder_posteo" }],
      [{ text: "📂 Descargar último reporte de Diario", callback_data: "folder_diario" }],
      [{ text: "📂 Descargar último reporte de PYME", callback_data: "folder_pyme" }],
      [{ text: "📂 Descargar último reporte de Calidad", callback_data: "folder_calidad" }],
      [{ text: "📂 Descargar último reporte de Muestreo", callback_data: "folder_muestreo" }] 
    ]
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

  Logger.log(`Callback data received: ${callbackData}`);
  
  var sheet = SpreadsheetApp.openById("154X5o9_2X2oV78OibIT20AXh1cZrkNRUbDMBecCAwbc").getSheetByName("Lideres");
  var telegramUserRow = findUserRow(sheet, 7, chatId); 

  if (telegramUserRow !== -1 && sheet.getRange(telegramUserRow, 6).getValue() === "ACTIVO") { 
    var folderId;
    if (callbackData === 'folder_posteo') {
      folderId = folderPosteoId;
    } else if (callbackData === 'folder_diario') {
      folderId = folderDiarioId;
    } else if (callbackData === 'folder_pyme') {
      folderId = folderPYMEId;
    } else if (callbackData === 'folder_calidad') {
      folderId = folderCalidadId;
    } else if (callbackData === 'folder_muestreo') { 
      folderId = folderMuestreoId;
    }

    Logger.log(`Folder ID resolved: ${folderId}`);

    if (folderId) {
      sendText(chatId, `🔍 Opción elegida: ${callbackData.replace('folder_', 'Descargar último reporte de ')}.\n`);
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

function searchAndSendLatestPDF(chatId, folderId, userId) {
  var files = DriveApp.getFolderById(folderId).searchFiles(`title contains "${userId}"`);
  var latestFile = null;
  while (files.hasNext()) {
    var file = files.next();
    if (!latestFile || file.getLastUpdated() > latestFile.getLastUpdated()) {
      latestFile = file;
    }
  }

  if (latestFile) {
    var pdfBlob = convertToPDF(latestFile);
    sendDocument(chatId, pdfBlob, latestFile.getName().replace(/\.[^/.]+$/, ".pdf"));
  } else {
    sendText(chatId, "No se encontraron reportes asociados a tu usuario ❌❌.");
  }

  sendInitialMenu(chatId);
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
}

function searchAndSendSheetFile(chatId, folderId, userId) {
  var files = DriveApp.getFolderById(folderId).searchFiles(`title contains "${userId}"`);
  var latestFile = null;
  while (files.hasNext()) {
    var file = files.next();
    if (!latestFile || file.getLastUpdated() > latestFile.getLastUpdated()) {
      latestFile = file;
    }
  }

  if (latestFile) {
    sendSheetFile(chatId, latestFile);
  } else {
    sendText(chatId, "No se encontraron reportes asociados a tu usuario ❌❌.");
  }

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
}
