/**
 * INSTRUCCIONES DE INSTALACIÓN EN GOOGLE SHEETS:
 * 
 * 1. Crea una nueva hoja de cálculo en Google Sheets (https://sheets.new).
 * 2. Ve a la barra de menú superior y selecciona: Extensiones > Apps Script.
 * 3. Borra cualquier código existente en el editor de Apps Script.
 * 4. Pega todo el contenido de este archivo en el editor.
 * 5. Haz clic en el botón de guardar (icono de disco).
 * 6. Haz clic en "Desplegar" (Deploy) > "Nuevo despliegue" (New deployment).
 * 7. Configura las siguientes opciones:
 *    - Tipo de despliegue: Aplicación web (Web app).
 *    - Descripción: Bitacora UAS Sync.
 *    - Ejecutar como (Execute as): Mi usuario (Tu cuenta de Google).
 *    - Quién tiene acceso (Who has access): Cualquiera (Anyone).
 * 8. Haz clic en "Desplegar". Google te pedirá autorizar permisos para acceder a las hojas de cálculo. Concede los permisos necesarios.
 * 9. Copia la "URL de la aplicación web" generada (debe terminar en /exec).
 * 10. Pega esa URL en la interfaz web de Foxtrot Sky Eye (campo superior derecho).
 */

// Nombre de la hoja que se creará para guardar los registros
var SHEET_NAME = "Bitacora_UAS";

// Encabezados de la tabla
var HEADERS = [
  "id", "dateStr", "opName", "opDni", "opLegajo", "flightDate", 
  "startTime", "endTime", "duration", "location", "wind", "altitude", 
  "type", "droneModel", "droneSn", "batInitial", "batFinal", 
  "hasNovelty", "noveltyDesc", "analysis"
];

/**
 * Método GET: Retorna todos los registros de la hoja en formato JSON
 */
function doGet(e) {
  try {
    var sheet = getOrCreateSheet();
    var data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return ContentService.createTextOutput(JSON.stringify([]))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    var headers = data[0];
    var flights = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var flight = {};
      
      headers.forEach(function(header, colIndex) {
        var val = row[colIndex];
        
        // Conversión de tipos de datos al formato original
        if (header === "wind" || header === "duration") {
          flight[header] = parseFloat(val) || 0;
        } else if (header === "batInitial" || header === "batFinal" || header === "altitude") {
          flight[header] = parseInt(val, 10) || 0;
        } else if (header === "hasNovelty") {
          flight[header] = (String(val).toLowerCase() === "true" || val === true);
        } else if (header === "flightDate" && val instanceof Date) {
          // Formatear fecha a YYYY-MM-DD
          flight[header] = val.toISOString().split('T')[0];
        } else {
          flight[header] = val;
        }
      });
      
      flights.push(flight);
    }
    
    return ContentService.createTextOutput(JSON.stringify(flights))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({ error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Método POST: Recibe el listado de vuelos en JSON y sincroniza la hoja de cálculo
 */
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var flights = Array.isArray(payload) ? payload : [payload];
    
    var sheet = getOrCreateSheet();
    
    // Obtener los IDs ya existentes en la hoja de cálculo para actualizar o insertar
    var data = sheet.getDataRange().getValues();
    var idRowMap = {};
    
    // Mapear los IDs existentes con sus respectivos números de fila (1-indexed)
    if (data.length > 1) {
      for (var i = 1; i < data.length; i++) {
        var flightId = data[i][0]; // ID está en la primera columna
        if (flightId) {
          idRowMap[flightId] = i + 1;
        }
      }
    }
    
    // Procesar cada vuelo del payload
    flights.forEach(function(flight) {
      // Crear array de la fila en orden de los headers
      var rowData = HEADERS.map(function(header) {
        var val = flight[header];
        return val !== undefined ? val : "";
      });
      
      var flightId = flight.id;
      
      if (flightId && idRowMap[flightId]) {
        // Actualizar fila existente
        var rowNum = idRowMap[flightId];
        sheet.getRange(rowNum, 1, 1, HEADERS.length).setValues([rowData]);
      } else {
        // Insertar nueva fila al final
        sheet.appendRow(rowData);
        // Si el ID se acaba de registrar, agregarlo al mapa temporal por si viene duplicado en el payload
        if (flightId) {
          idRowMap[flightId] = sheet.getLastRow();
        }
      }
    });
    
    var responseObj = { 
      status: "success", 
      message: "Sincronización completada exitosamente.",
      recordsProcessed: flights.length
    };
    
    return ContentService.createTextOutput(JSON.stringify(responseObj))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch(error) {
    var errObj = { status: "error", message: error.toString() };
    return ContentService.createTextOutput(JSON.stringify(errObj))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Obtiene o crea la hoja de cálculo interna con sus respectivos encabezados
 */
function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // Configurar encabezados
    sheet.appendRow(HEADERS);
    // Dar formato estético básico a los encabezados
    var headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#0d133a");
    headerRange.setFontColor("#00f5d4");
    headerRange.setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
  }
  
  return sheet;
}
