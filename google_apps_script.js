// ============================================================
// FOXTROT SKY EYE - Google Apps Script Webhook
// Versión: 2.0 - 20 columnas completas
// ============================================================
// INSTRUCCIONES DE DESPLIEGUE:
// 1. Abre tu Google Sheet → Extensiones → Apps Script
// 2. Pega este código (reemplaza el contenido anterior)
// 3. Haz clic en "Implementar" → "Gestionar implementaciones"
// 4. Edita la implementación activa (ícono lápiz)
// 5. En "Quién tiene acceso" → selecciona "Cualquiera" (Anyone)
// 6. Copia la URL generada y pégala en el campo "Sheets URL" de tu app
// ============================================================

function doPost(e) {
  try {
    // Abre el primer sheet del Google Spreadsheet por ID
    // El ID está en la URL de tu planilla: docs.google.com/spreadsheets/d/[ESTE_ID]/edit
    var sheet = SpreadsheetApp.openById('1JLuDgU8B14K4B5hsTEl5CmUCWQ8C4YDzF4e5HuvlavA').getSheets()[0];

    // Parsea el cuerpo JSON enviado desde index.html
    var data = JSON.parse(e.postData.contents);

    // Google Sheets tiene un límite de 50,000 caracteres por celda.
    // Si la imagen es un base64 grande, usamos el marcador de posición.
    var imageValue = data.image || "";
    if (imageValue.length > 49000) {
      imageValue = "[Imagen guardada localmente - demasiado grande para Sheets]";
    }

    // Inserta una nueva fila con los 20 campos de la bitácora
    // El orden debe coincidir con las cabeceras de tu Google Sheet:
    // ID | Operador | DNI | Legajo | Fecha | Hora Inicio | Hora Fin | Duración (min) |
    // Ubicación | Viento (km/h) | Altitud (m) | Tipo de Vuelo | Modelo Dron | S/N Dron |
    // Tiene Novedad | Descripción Novedad | Análisis Preliminar |
    // Batería Inicial (%) | Batería Final (%) | Imagen
    sheet.appendRow([
      data.id          || "",
      data.operator    || "",
      data.dni         || "",
      data.legajo      || "",
      data.date        || "",
      data.startTime   || "",
      data.endTime     || "",
      data.duration    || "",
      data.location    || "",
      data.wind        || "",
      data.altitude    || "",
      data.type        || "",
      data.droneModel  || "",
      data.droneSN     || "",
      data.hasNovelty  || "",
      data.noveltyDesc || "",
      data.analysis    || "",
      data.batInitial  || "",
      data.batFinal    || "",
      imageValue
    ]);

    // Respuesta de éxito en JSON
    return ContentService
      .createTextOutput(JSON.stringify({ "status": "success", "rows": 1 }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    // Respuesta de error con el mensaje del error para facilitar depuración
    return ContentService
      .createTextOutput(JSON.stringify({ "status": "error", "message": error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Responde a peticiones GET para verificar que el webhook está activo
// Puedes probarlo abriendo la URL directamente en el navegador
function doGet(e) {
  return ContentService.createTextOutput(
    "✅ Webhook Foxtrot Sky Eye activo. Usa HTTP POST con Content-Type: application/json para registrar vuelos."
  );
}
