function onFormSubmit(e) {
  // Obtener la hoja de respuestas del formulario
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1'); // Cambia el nombre si es diferente
  if (!sheet) {
    Logger.log('No se encontró la hoja de respuestas');
    return;
  }
  
  var responses = sheet.getDataRange().getValues();
  var maxCapacity = 35; // Límite de personas por slot

  // Inicializar contadores para los slots
  var slotCount = {
    "Lunes 26 de Agosto: 10:30 - 11:30.": 0,
    "Lunes 26 de Agosto: 14:00 - 15:00.": 0,
    "Martes 27 de Agosto: 10:30 - 11:30.": 0,
    "Martes 27 de Agosto: 14:00 - 15:00.": 0,
    "Miércoles 28 de Agosto: 12:00 - 13:00": 0
  };

  // Contar el número de personas por slot
  for (var i = 1; i < responses.length; i++) { // Comenzar en 1 para omitir la fila de encabezado
    var slot = responses[i][5]; // Cambia el índice según la columna que corresponda a los slots
    if (slot in slotCount) {
      slotCount[slot]++;
    }
  }

  // Obtener la última respuesta
  var latestResponse = responses[responses.length - 1];
  var selectedSlot = latestResponse[5]; // Cambia el índice según la columna que corresponda a los slots
  var email = latestResponse[1]; // Dirección de correo electrónico está en la columna 2

  // Verificar si la dirección de correo electrónico es válida
  if (!isValidEmail(email)) {
    Logger.log('Dirección de correo electrónico inválida: ' + email);
    return;
  }

  // Verificar si el slot está lleno
  if (slotCount[selectedSlot] >= maxCapacity) {
    // Enviar email notificando al usuario que no se pudo inscribir
    MailApp.sendEmail({
      to: email,
      subject: "Inscripción fallida - Slot lleno",
      body: "Lo sentimos, pero el slot " + selectedSlot + " ya está lleno. Por favor, elige otro slot."
    });
    
    // Eliminar la fila de la respuesta si el slot está lleno
    sheet.deleteRow(responses.length);
  } else {
    // Obtener el enlace para el slot seleccionado
    var link = getMeetingLink(selectedSlot);
    
    // Enviar email con el enlace a la reunión
    MailApp.sendEmail({
      to: email,
      subject: "Confirmación de inscripción al curso - " + selectedSlot,
      body: "Gracias por inscribirte en el curso. Aquí tienes el enlace para unirte a la reunión: " + link
    });
  }

  // Actualizar la información de slots disponibles
  updateSlotAvailability(slotCount, maxCapacity);
}

function updateSlotAvailability(slotCount, maxCapacity) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Slots Disponibles'); // Crea esta hoja en tu documento
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Slots Disponibles');
  }
  
  // Limpiar la hoja
  sheet.clear();

  // Escribir encabezados
  sheet.getRange(1, 1).setValue('Slot');
  sheet.getRange(1, 2).setValue('Ocupados');
  sheet.getRange(1, 3).setValue('Disponibles');

  // Escribir datos de cada slot
  var row = 2;
  for (var slot in slotCount) {
    sheet.getRange(row, 1).setValue(slot);
    sheet.getRange(row, 2).setValue(slotCount[slot]);
    sheet.getRange(row, 3).setValue(maxCapacity - slotCount[slot]);
    row++;
  }
}

function isValidEmail(email) {
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

function getMeetingLink(slot) {
  var linkSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Links'); // Cambia 'Links' por el nombre correcto
  if (!linkSheet) {
    Logger.log('No se encontró la hoja de enlaces');
    return "Enlace no disponible";
  }
  
  var data = linkSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) { // Comenzar en 1 para omitir la fila de encabezado
    if (data[i][0] === slot) {
      Logger.log('Enlace encontrado: ' + data[i][1]);
      return data[i][1]; // Retornar el enlace
    }
  }
  Logger.log('Enlace no encontrado para el slot: ' + slot);
  return "Enlace no disponible";
}
