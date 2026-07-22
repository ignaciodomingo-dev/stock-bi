// ============================================================
// SCRIPT 1 — STOCK ACTUAL BSALE
// Descarga snapshot actual de stock por variante y oficina.
// Excluye y marca oficinas según NEGOCIO.* (ver config_local.js).
// Trigger recomendado: cada 1 hora.
// ============================================================

const STOCK_CONFIG = {
  TOKEN_PROPERTY: 'BSALE_TOKEN',
  SHEET_NAME: 'Stock Actual',
  LIMIT: 50
};

const STOCK_HEADERS = [
  'ID Stock',
  'ID Variante',
  'SKU',
  'Descripción',
  'ID Oficina',
  'Nombre Oficina',
  'Tipo Oficina',
  'Stock Real',
  'Stock Comprometido',
  'Stock Disponible',
  'Última Actualización'
];

function descargarStockActual() {
  var token = _getToken();

  SpreadsheetApp.getActiveSpreadsheet()
    .toast('Descargando stock desde Bsale…', 'Stock Actual', -1);

  var filas = [];
  var offset = 0;
  var total = null;
  var ahora = new Date().toLocaleString('es-CL');

  while (true) {
    var url = 'https://api.bsale.io/v1/stocks.json'
      + '?expand=%5Bvariant,office%5D'
      + '&limit=' + STOCK_CONFIG.LIMIT
      + '&offset=' + offset;

    var data = _fetchBsale(url, token);
    var items = data.items || [];

    if (total === null) total = Number(data.count || 0);
    if (!items.length) break;

    items.forEach(function(stock) {
      var variante = stock.variant || {};
      var oficina  = stock.office  || {};
      var nombreOficina = (oficina.name || '').toUpperCase().trim();

      if (NEGOCIO.OFICINAS_EXCLUIR.some(function(ex) {
        return nombreOficina.indexOf(ex.toUpperCase()) !== -1;
      })) return;

      var tipoOficina = NEGOCIO.OFICINAS_SECUNDARIAS.some(function(s) {
        return nombreOficina.indexOf(s.toUpperCase()) !== -1;
      }) ? 'Secundaria' : 'Principal';

      filas.push([
        stock.id              || '',
        variante.id           || '',
        variante.code         || '',
        variante.description  || '',
        oficina.id            || '',
        oficina.name          || '',
        tipoOficina,
        stock.quantity            || 0,
        stock.quantityCommitment  || 0,
        stock.quantityAvailable   || 0,
        ahora
      ]);
    });

    offset += STOCK_CONFIG.LIMIT;
    if (offset >= total) break;
    Utilities.sleep(300);
  }

  // Escritura atómica: solo limpiamos/escribimos si la descarga trajo datos.
  // Si Bsale falló a mitad, _fetchBsale ya lanzó y nunca llegamos acá → la hoja
  // conserva su última versión buena (evita dejarla vacía = "quiebre" falso).
  if (!filas.length) {
    throw new Error('❌ Bsale no devolvió stock. NO se modificó la hoja (se conserva la última versión).');
  }
  var sheet = _prepararHoja(STOCK_CONFIG.SHEET_NAME, STOCK_HEADERS);
  sheet.getRange(2, 1, filas.length, STOCK_HEADERS.length).setValues(filas);
  _aplicarFormato(sheet, filas.length);

  var msg = '✅ Stock actualizado: ' + filas.length + ' registros (' + ahora + ')';
  Logger.log(msg);
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'Stock Actual', 8);
}

function instalarTriggerStock() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'descargarStockActual') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('descargarStockActual')
    .timeBased()
    .everyHours(1)
    .create();
  Logger.log('✅ Trigger horario instalado para descargarStockActual');
}

function _getToken() {
  var token = PropertiesService.getScriptProperties()
    .getProperty(STOCK_CONFIG.TOKEN_PROPERTY);
  if (!token) throw new Error('❌ No existe token guardado.');
  return token;
}

function _prepararHoja(nombre, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(nombre);
  if (!sheet) sheet = ss.insertSheet(nombre);
  sheet.clearContents();
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#F3F3F3')
    .setFontColor('#000000');
  sheet.setFrozenRows(1);
  return sheet;
}

function _fetchBsale(url, token) {
  var res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { access_token: token },
    muteHttpExceptions: true
  });
  var code = res.getResponseCode();
  if (code !== 200) {
    throw new Error('❌ Error API Bsale HTTP ' + code + ': ' + res.getContentText());
  }
  return JSON.parse(res.getContentText());
}

function _aplicarFormato(sheet, numFilas) {
  sheet.getRange(2, 8, numFilas, 3).setNumberFormat('#,##0');
  var regla = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(0)
    .setBackground('#FCE8E6')
    .setRanges([sheet.getRange(2, 10, numFilas, 1)])
    .build();
  sheet.setConditionalFormatRules([regla]);
  sheet.autoResizeColumns(1, STOCK_HEADERS.length);
}
