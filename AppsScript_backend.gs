/**
 * ENSAYO DE TRAFOS — Backend
 * -----------------------------------------------------------------
 * 1) Creá una Google Sheet nueva. Va a tener 2 hojas: "Maestro" e "Historial"
 *    (este código las crea solo si no existen, no hace falta armarlas a mano).
 * 2) Extensiones > Apps Script > borrá el contenido default > pegá esto.
 * 3) Implementar > Nueva implementación > tipo "Aplicación web"
 *      - Ejecutar como: Yo
 *      - Quién tiene acceso: Cualquier usuario
 * 4) Copiá la URL (termina en /exec) y pegala en ensayo_de_trafos.html
 *    en la constante APPS_SCRIPT_URL.
 * -----------------------------------------------------------------
 */

const HEADERS = [
  'Fecha', 'N Serie', 'Ubicacion', 'Es Pozo', 'Latitud', 'Longitud', 'Link Maps',
  'Marca', 'Tipo', 'Potencia', 'Tension Primaria', 'Tension Secundaria', 'Estado Visual',
  'Ensayos Realizados', 'Indice Polarizacion', 'Relacion Absorcion',
  'Desviacion Resistencia', 'Estado Conmutador', 'Estado Servicio',
  'Comentarios', 'Foto Placa', 'Foto Ensayo'
];

const COL = {
  FECHA: 0, SERIE: 1, UBICACION: 2, ES_POZO: 3, LAT: 4, LNG: 5, MAPS: 6,
  MARCA: 7, TIPO: 8, POTENCIA: 9, TENSION_P: 10, TENSION_S: 11, ESTADO_VISUAL: 12,
  ENSAYOS: 13, INDICE_POL: 14, RELACION_ABS: 15, DESVIACION: 16, CONMUTADOR: 17,
  ESTADO_SERVICIO: 18, COMENTARIOS: 19, FOTO_PLACA: 20, FOTO_ENSAYO: 21
};

function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(HEADERS);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  } else {
    // Mantenemos el encabezado siempre sincronizado con la estructura actual,
    // para que las columnas no queden desalineadas si el esquema cambió.
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  return sheet;
}

function getOrCreatePhotosFolder() {
  const folderName = 'ENSAYO DE TRAFOS - Fotos';
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

function saveFotoToDrive(fotoData, serie, tipo) {
  if (!fotoData || !fotoData.base64) return '';
  const folder = getOrCreatePhotosFolder();
  const bytes = Utilities.base64Decode(fotoData.base64);
  const blob = Utilities.newBlob(bytes, fotoData.mimeType || 'image/jpeg',
    `${serie}_${tipo}_${new Date().getTime()}`);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function findMaestroRow(maestro, serie) {
  const lastRow = maestro.getLastRow();
  if (lastRow <= 1) return -1;
  const series = maestro.getRange(2, COL.SERIE + 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < series.length; i++) {
    if (String(series[i][0]).trim() === String(serie).trim()) return i + 2;
  }
  return -1;
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);

  // --- Actualización de ubicación (desde "Consultar Trafo") ---
  // Solo toca Ubicacion / Coordenadas / (opcional) Foto Ensayo si es un pozo.
  // El resto de los datos del trafo (marca, potencia, tensiones, estado, etc.)
  // se mantienen intactos.
  if (data.action === 'updateLocation') {
    const maestro = getOrCreateSheet('Maestro');
    const historial = getOrCreateSheet('Historial');
    const mapsLink = (data.lat && data.lng) ? `https://www.google.com/maps?q=${data.lat},${data.lng}` : '';
    const foundRow = findMaestroRow(maestro, data.serie);

    if (foundRow > -1) {
      maestro.getRange(foundRow, COL.FECHA + 1).setValue(data.timestamp || new Date().toISOString());
      maestro.getRange(foundRow, COL.UBICACION + 1).setValue(data.ubicacion || '');
      maestro.getRange(foundRow, COL.LAT + 1).setValue(data.lat || '');
      maestro.getRange(foundRow, COL.LNG + 1).setValue(data.lng || '');
      maestro.getRange(foundRow, COL.MAPS + 1).setValue(mapsLink);

      if (data.fotoEnsayo && data.fotoEnsayo.base64) {
        const nuevaFotoUrl = saveFotoToDrive(data.fotoEnsayo, data.serie, 'ensayo');
        maestro.getRange(foundRow, COL.FOTO_ENSAYO + 1).setValue(nuevaFotoUrl);
      }
      // Si la nueva ubicación es un predio, NO tocamos Foto Ensayo: se mantiene la que ya tenía.

      const fullRow = maestro.getRange(foundRow, 1, 1, HEADERS.length).getValues()[0];
      fullRow[COL.COMENTARIOS] = (fullRow[COL.COMENTARIOS] ? fullRow[COL.COMENTARIOS] + ' | ' : '') + 'Cambio de ubicación';
      historial.appendRow(fullRow);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ result: foundRow > -1 ? 'ok' : 'not_found' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // --- Carga normal de un ensayo completo ---
  const mapsLink = (data.lat && data.lng) ? `https://www.google.com/maps?q=${data.lat},${data.lng}` : '';
  const fotoPlacaUrl = saveFotoToDrive(data.fotoPlaca, data.serie, 'placa');
  const fotoEnsayoUrl = saveFotoToDrive(data.fotoEnsayo, data.serie, 'ensayo');

  const row = [
    data.timestamp || new Date().toISOString(),
    data.serie || '',
    data.ubicacion || '',
    data.esPozo ? 'Si' : 'No',
    data.lat || '',
    data.lng || '',
    mapsLink,
    data.marca || '',
    data.tipo || '',
    data.potencia || '',
    data.tensionP || '',
    data.tensionS || '',
    data.estadoVisual || '',
    data.ensayosRealizados || '',
    data.indicePolarizacion || '',
    data.relacionAbsorcion || '',
    data.desviacionResistencia || '',
    data.estadoConmutador || '',
    data.estadoServicio || '',
    data.comentarios || '',
    fotoPlacaUrl,
    fotoEnsayoUrl
  ];

  const historial = getOrCreateSheet('Historial');
  historial.appendRow(row);

  const maestro = getOrCreateSheet('Maestro');
  const foundRow = findMaestroRow(maestro, data.serie);
  if (foundRow > -1) {
    maestro.getRange(foundRow, 1, 1, row.length).setValues([row]);
  } else {
    maestro.appendRow(row);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ result: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function rowToOutput(r) {
  return {
    found: true,
    serie: r[COL.SERIE],
    fecha: r[COL.FECHA],
    ubicacion: r[COL.UBICACION],
    lat: r[COL.LAT],
    lng: r[COL.LNG],
    marca: r[COL.MARCA],
    tipo: r[COL.TIPO],
    potencia: r[COL.POTENCIA],
    tensionP: r[COL.TENSION_P],
    tensionS: r[COL.TENSION_S],
    estadoVisual: r[COL.ESTADO_VISUAL],
    estadoServicio: r[COL.ESTADO_SERVICIO],
    fotoPlaca: r[COL.FOTO_PLACA],
    fotoEnsayo: r[COL.FOTO_ENSAYO]
  };
}

function doGet(e) {
  const maestro = getOrCreateSheet('Maestro');
  const lastRow = maestro.getLastRow();

  // --- Exportar planilla completa como CSV ---
  if (e.parameter.export === 'maestro') {
    let csv = '';
    if (lastRow >= 1) {
      const values = maestro.getRange(1, 1, lastRow, HEADERS.length).getValues();
      csv = values.map(row =>
        row.map(cell => {
          const s = String(cell === null || cell === undefined ? '' : cell).replace(/"/g, '""');
          return `"${s}"`;
        }).join(',')
      ).join('\n');
    }
    return ContentService.createTextOutput(csv).setMimeType(ContentService.MimeType.CSV);
  }

  // --- Filtrar por Estado para el servicio ---
  if (e.parameter.estado) {
    const output = { results: [] };
    if (lastRow > 1) {
      const values = maestro.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
      for (let i = 0; i < values.length; i++) {
        const r = values[i];
        if (String(r[COL.ESTADO_SERVICIO]).trim() === String(e.parameter.estado).trim() && String(r[COL.SERIE]).trim() !== '') {
          output.results.push(rowToOutput(r));
        }
      }
    }
    return ContentService.createTextOutput(JSON.stringify(output)).setMimeType(ContentService.MimeType.JSON);
  }

  // --- Buscar por N° de serie (comportamiento original) ---
  const serie = e.parameter.serie;
  let output = { found: false };
  if (serie && lastRow > 1) {
    const values = maestro.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
    for (let i = 0; i < values.length; i++) {
      const r = values[i];
      if (String(r[COL.SERIE]).trim() === String(serie).trim()) {
        output = rowToOutput(r);
        break;
      }
    }
  }

  return ContentService.createTextOutput(JSON.stringify(output)).setMimeType(ContentService.MimeType.JSON);
}
