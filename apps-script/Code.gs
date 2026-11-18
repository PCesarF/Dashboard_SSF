// Configurar la propiedad de la secuencia de comandos SSF_SPREADSHEET_ID.
const SSF_COLUMNS = {
  "PROYECTOS": [
    "nombre",
    "codigo",
    "estado",
    "valor_total",
    "total_ejecutado",
    "ejecutado_vigencia",
    "presupuesto_2026",
    "saldo",
    "v2024",
    "v2025",
    "porc_ejec_vigencia",
    "porc_ejec_general",
    "porc_ejec_obra",
    "fecha_inicio",
    "fecha_fin",
    "dias_info",
    "dias_num",
    "lmi_2026",
    "lmi_ejec",
    "lmi_saldo"
  ],
  "ALCANCES": [
    "proyecto_nombre",
    "alcance_nombre",
    "avance",
    "fecha_inicio",
    "fecha_fin"
  ],
  "MODIFICACIONES": [
    "nombre",
    "codigo",
    "modificacion",
    "categoría_modificacion",
    "fuentes_modificacion",
    "nueva_fecha_modificacion",
    "valor_total_modificacion",
    "valor_modificacion_saldo_para_obras",
    "valor_modificacion_ley 115",
    "valor_modificacion_excedentes 55%",
    "cambio_nombre",
    "primer_año_modificacion",
    "segundo_año_modificacion",
    "tercer_año_modificacion",
    "cuarto_año_modificacion"
  ]
};

function doGet(e) {
  const callback = e && e.parameter && e.parameter.callback;
  if (callback && !/^[A-Za-z_$][A-Za-z0-9_$]{0,63}$/.test(callback)) {
    return ContentService.createTextOutput('Callback inválido')
      .setMimeType(ContentService.MimeType.TEXT);
  }

  const payload = {tipo: 'SSF_DRIVE_BRIDGE', ok: false};
  try {
    const id = PropertiesService.getScriptProperties().getProperty('SSF_SPREADSHEET_ID');
    if (!id) throw new Error('Falta la propiedad SSF_SPREADSHEET_ID');
    const book = SpreadsheetApp.openById(id);
    const timezone = book.getSpreadsheetTimeZone();
    payload.hojas = {};
    for (const [name, columns] of Object.entries(SSF_COLUMNS)) {
      payload.hojas[name] = readDashboardSheet_(book, name, columns, timezone);
    }
    payload.ok = true;
  } catch (error) {
    console.error(error);
    delete payload.hojas;
    payload.error = 'No se pudo leer la plantilla maestra de SSF.';
  }

  const json = JSON.stringify(payload).replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
  return ContentService.createTextOutput(callback ? callback + '(' + json + ');' : json)
    .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

function readDashboardSheet_(book, name, columns, timezone) {
  const sheet = book.getSheetByName(name);
  if (!sheet || sheet.getLastRow() < 1) throw new Error('Falta la hoja ' + name);
  const [headers, ...rows] = sheet.getDataRange().getValues();
  const positions = columns.map(column => {
    const index = headers.findIndex(header => String(header).trim() === column);
    if (index < 0) throw new Error('Falta la columna ' + column + ' en ' + name);
    return index;
  });

  return [columns].concat(rows.map(row => positions.map(index => {
    const value = row[index];
    return value instanceof Date ? Utilities.formatDate(value, timezone, 'yyyy-MM-dd') : value;
  })));
}
