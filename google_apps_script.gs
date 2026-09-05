const SPREADSHEET_ID = '15zqkKz4fA1UU7KWeD7K5dmnEQdzdz3gERt6xkcvTZU8';
const SHEET_NAME = 'Resultados';
const HEADERS = ['id', 'test', 'guardadoEn', 'nombre', 'fecha', 'datoExtra', 'puntaje', 'incorrectas', 'total', 'porcentaje', 'respuestas'];

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || '{}');
    if (payload.action !== 'save' || !payload.test || !payload.result) {
      return jsonResponse({ok: false, error: 'Solicitud incompleta'});
    }
    const sheet = getSheet();
    purgeOldRecords(sheet);
    const result = payload.result;
    const id = result.id || Utilities.getUuid();
    const ids = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat() : [];
    if (ids.indexOf(id) === -1) {
      sheet.appendRow([
        id,
        payload.test,
        result.guardadoEn || new Date().toISOString(),
        result.nombre || '',
        result.fecha || '',
        result.edad || result.wordClass || '',
        result.puntaje || 0,
        result.incorrectas || 0,
        result.total || 0,
        result.porcentaje || 0,
        JSON.stringify(result.respuestas || [])
      ]);
    }
    return jsonResponse({ok: true, id});
  } catch (error) {
    return jsonResponse({ok: false, error: String(error)});
  }
}

function doGet(e) {
  try {
    const test = e.parameter.test || '';
    const sheet = getSheet();
    purgeOldRecords(sheet);
    const values = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues() : [];
    const records = values.filter(row => !test || row[1] === test).map(row => ({
      id: row[0],
      test: row[1],
      guardadoEn: row[2],
      nombre: row[3],
      fecha: row[4],
      datoExtra: row[5],
      edad: row[1] === 'pretest' ? row[5] : '',
      wordClass: row[1] === 'postest' ? row[5] : '',
      puntaje: row[6],
      incorrectas: row[7],
      total: row[8],
      porcentaje: row[9],
      respuestas: parseResponses(row[10])
    })).reverse();
    return jsonResponse({ok: true, records});
  } catch (error) {
    return jsonResponse({ok: false, error: String(error), records: []});
  }
}

function getSheet() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
  return sheet;
}

function parseResponses(value) {
  try { return JSON.parse(value || '[]'); }
  catch (error) { return []; }
}

function purgeOldRecords(sheet) {
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  const dates = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
  for (let index = dates.length - 1; index >= 0; index--) {
    const savedDate = String(dates[index][0] || '').slice(0, 10);
    if (savedDate && savedDate < today) sheet.deleteRow(index + 2);
  }
}

function jsonResponse(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}
