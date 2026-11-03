const APP_NAME = "Prueba diagnóstica oral - Personal hotelero";

function doGet() {
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle(APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/*
 * EJECUTAR ESTA FUNCIÓN UNA SOLA VEZ
 * antes de publicar la aplicación.
 */
function configurarProyecto() {
  const props = PropertiesService.getScriptProperties();

  // Carpeta donde se guardarán los audios
  const folder = DriveApp.createFolder(
    "Audios - Prueba oral inglés y portugués",
  );

  // Planilla de resultados
  const ss = SpreadsheetApp.create(
    "Resultados - Prueba oral personal hotelero",
  );

  const sheet = ss.getSheets()[0];
  sheet.setName("Audios");

  sheet
    .getRange(1, 1, 1, 10)
    .setValues([
      [
        "Fecha",
        "Clave",
        "Código participante",
        "Código hotel",
        "Idioma",
        "Pregunta",
        "Situación",
        "Archivo",
        "Enlace Drive",
        "Formato",
      ],
    ]);

  sheet.setFrozenRows(1);

  // Hoja de participantes
  const participants = ss.insertSheet("Participantes");

  participants
    .getRange(1, 1, 1, 5)
    .setValues([
      [
        "Fecha de finalización",
        "Código participante",
        "Código hotel",
        "Estado",
        "Cantidad de audios",
      ],
    ]);

  participants.setFrozenRows(1);

  props.setProperties({
    AUDIO_FOLDER_ID: folder.getId(),
    RESULTS_SPREADSHEET_ID: ss.getId(),
  });

  Logger.log("CARPETA DE AUDIOS:");
  Logger.log(folder.getUrl());

  Logger.log("PLANILLA DE RESULTADOS:");
  Logger.log(ss.getUrl());
}

function doPost(e) {
  try {
    const datos = JSON.parse(e.postData.contents);
    let resultado;

    if (datos.accion === "finalizar") {
      resultado = finalizarParticipante(datos);
    } else {
      resultado = guardarAudio(datos);
    }

    return ContentService.createTextOutput(
      JSON.stringify(resultado),
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({
        ok: false,
        error: String(err.message || err),
      }),
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

/*
 * Guarda cada respuesta oral individualmente.
 */
function guardarAudio(datos) {
  const props = PropertiesService.getScriptProperties();

  const folderId = props.getProperty("AUDIO_FOLDER_ID");
  const spreadsheetId = props.getProperty("RESULTS_SPREADSHEET_ID");

  if (!folderId || !spreadsheetId) {
    throw new Error("Primero debe ejecutar la función configurarProyecto().");
  }

  if (!datos || !datos.codigo || !datos.base64) {
    throw new Error("Datos incompletos.");
  }

  const codigo = limpiarCodigo(datos.codigo);
  const hotel = limpiarCodigo(datos.hotel || "SIN-HOTEL");

  const idioma =
    datos.idioma === "EN" ? "EN" : datos.idioma === "PT" ? "PT" : null;

  if (!idioma) {
    throw new Error("Idioma inválido.");
  }

  const pregunta = Number(datos.pregunta);

  if (!Number.isInteger(pregunta) || pregunta < 1 || pregunta > 10) {
    throw new Error("Número de pregunta inválido.");
  }

  const rootFolder = DriveApp.getFolderById(folderId);

  // Una carpeta por participante
  const participantFolder = obtenerOCrearCarpeta(rootFolder, codigo);

  const mimeType = datos.mimeType || "audio/webm";
  const extension = obtenerExtension(mimeType);

  const numero = String(pregunta).padStart(2, "0");

  const fileName = codigo + "_" + idioma + "_Q" + numero + "." + extension;

  // Si vuelve a grabar la pregunta, reemplazamos el audio anterior.
  const previousFiles = participantFolder.getFilesByName(fileName);

  while (previousFiles.hasNext()) {
    previousFiles.next().setTrashed(true);
  }

  const bytes = Utilities.base64Decode(datos.base64);

  const blob = Utilities.newBlob(bytes, mimeType, fileName);

  const file = participantFolder.createFile(blob);

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName("Audios");

  const key = codigo + "|" + idioma + "|" + pregunta;

  const rowData = [
    new Date(),
    key,
    codigo,
    hotel,
    idioma,
    pregunta,
    datos.situacion || "",
    fileName,
    file.getUrl(),
    mimeType,
  ];

  // Evita duplicados si la persona vuelve a grabar.
  const lastRow = sheet.getLastRow();

  let existingRow = 0;

  if (lastRow > 1) {
    const keys = sheet
      .getRange(2, 2, lastRow - 1, 1)
      .getValues()
      .flat();

    const index = keys.indexOf(key);

    if (index !== -1) {
      existingRow = index + 2;
    }
  }

  if (existingRow) {
    sheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }

  return {
    ok: true,
    fileName: fileName,
    url: file.getUrl(),
  };
}

/*
 * Se ejecuta al terminar las 20 preguntas.
 */
function finalizarParticipante(datos) {
  const props = PropertiesService.getScriptProperties();

  const spreadsheetId = props.getProperty("RESULTS_SPREADSHEET_ID");

  if (!spreadsheetId) {
    throw new Error("Primero debe ejecutar configurarProyecto().");
  }

  const codigo = limpiarCodigo(datos.codigo);
  const hotel = limpiarCodigo(datos.hotel || "SIN-HOTEL");

  const ss = SpreadsheetApp.openById(spreadsheetId);

  const audios = ss.getSheetByName("Audios");

  const participants = ss.getSheetByName("Participantes");

  let cantidad = 0;

  if (audios.getLastRow() > 1) {
    const codigos = audios
      .getRange(2, 3, audios.getLastRow() - 1, 1)
      .getValues()
      .flat();

    cantidad = codigos.filter((c) => c === codigo).length;
  }

  participants.appendRow([
    new Date(),
    codigo,
    hotel,
    cantidad >= 20 ? "COMPLETO" : "INCOMPLETO",
    cantidad,
  ]);

  return {
    ok: true,
    cantidad: cantidad,
  };
}

function obtenerOCrearCarpeta(parent, name) {
  const folders = parent.getFoldersByName(name);

  if (folders.hasNext()) {
    return folders.next();
  }

  return parent.createFolder(name);
}

function limpiarCodigo(value) {
  return String(value || "")
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9_-]/g, "")
    .substring(0, 40);
}

function obtenerExtension(mime) {
  if (mime.includes("mp4")) {
    return "m4a";
  }

  if (mime.includes("ogg")) {
    return "ogg";
  }

  if (mime.includes("wav")) {
    return "wav";
  }

  return "webm";
}
