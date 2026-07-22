// ============================================================
// ARCHIVO DIARIO DE STOCK — memoria histórica
// "Stock Actual" se sobreescribe cada hora y no deja historia.
// Este script guarda un snapshot compacto (Fecha, SKU, Stock Vendible)
// en la pestaña "Stock Histórico" en modo APPEND (no overwrite).
//
// Desbloquea distinguir "cero venta" de "cero stock" (quiebre) y, a
// futuro, calcular ventas perdidas por quiebre (velocidad × días sin
// stock) — el número para gerencia.
//
// Stock Vendible = suma del Stock Disponible de las oficinas "Principal"
// por SKU (misma definición que usa el semáforo; excluye secundarias).
//
// Uso: archivarStockDiario(). Pensado para correr 1 vez al día, después
// del pipeline. Es idempotente: si ya hay snapshot de hoy, no duplica.
// ============================================================

const ARCH_CONFIG = {
  SHEET_STOCK:  'Stock Actual',
  SHEET_SALIDA: 'Stock Histórico',
  // Stock Actual (1-based): SKU=3, Tipo Oficina=7, Stock Disponible=10
  ST_SKU: 3, ST_TIPO: 7, ST_DISP: 10,
  // Aviso cuando la hoja se acerca a mucho volumen (para compactar a mensual).
  AVISO_FILAS: 200000
};

const ARCH_HEADERS = ['Fecha', 'SKU', 'Stock Vendible'];

function archivarStockDiario() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Archivando snapshot de stock…', 'Stock Histórico', -1);

  var hoy = _archHoy0();

  var sheet = _archPrepararHoja();

  // Idempotencia: si la última fila ya es de hoy, no duplicamos.
  if (_archYaArchivadoHoy(sheet, hoy)) {
    var aviso = '↔️ Snapshot de hoy ya existe en "' + ARCH_CONFIG.SHEET_SALIDA + '". No se duplica.';
    Logger.log(aviso); ss.toast(aviso, 'Stock Histórico', 8);
    return;
  }

  // SKU → stock vendible (Principal), misma definición que el semáforo.
  var stock = _archStockVendible();
  var skus = Object.keys(stock).sort();
  if (!skus.length) {
    throw new Error('❌ "' + ARCH_CONFIG.SHEET_STOCK + '" no tiene stock vendible que archivar.');
  }

  var filas = skus.map(function(sku) { return [hoy, sku, stock[sku]]; });

  var desde = sheet.getLastRow() + 1;
  sheet.getRange(desde, 1, filas.length, ARCH_HEADERS.length).setValues(filas);
  sheet.getRange(desde, 1, filas.length, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(desde, 3, filas.length, 1).setNumberFormat('#,##0');

  var total = sheet.getLastRow() - 1;   // sin encabezado
  var msg = '✅ Stock Histórico: +' + filas.length + ' SKUs archivados (' +
            Utilities.formatDate(new Date(hoy), Session.getScriptTimeZone(), 'yyyy-MM-dd') +
            '). Total ' + total.toLocaleString('es-CL') + ' filas.';
  if (total > ARCH_CONFIG.AVISO_FILAS) {
    msg += ' ⚠️ Conviene compactar a mensual (se acerca al límite de celdas).';
  }
  Logger.log(msg); ss.toast(msg, 'Stock Histórico', 12);
}

// ------------------------------------------------------------
// LECTURAS
// ------------------------------------------------------------
// SKU → stock vendible (suma de oficinas "Principal"). Copia deliberada de
// la lógica del semáforo para no crear dependencia entre módulos.
function _archStockVendible() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ARCH_CONFIG.SHEET_STOCK);
  if (!sh || sh.getLastRow() < 2) throw new Error('❌ Falta "' + ARCH_CONFIG.SHEET_STOCK + '".');
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, ARCH_CONFIG.ST_DISP).getValues();
  var map = {};
  v.forEach(function(r) {
    if (String(r[ARCH_CONFIG.ST_TIPO - 1]).trim() !== 'Principal') return;   // excluye oficinas secundarias
    var sku = String(r[ARCH_CONFIG.ST_SKU - 1] || '').trim();
    if (!sku) return;
    map[sku] = (map[sku] || 0) + parseFloat(r[ARCH_CONFIG.ST_DISP - 1] || 0);
  });
  return map;
}

// ------------------------------------------------------------
// HELPERS
// ------------------------------------------------------------
function _archPrepararHoja() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ARCH_CONFIG.SHEET_SALIDA);
  if (!sheet) {
    sheet = ss.insertSheet(ARCH_CONFIG.SHEET_SALIDA);
    sheet.appendRow(ARCH_HEADERS);
    sheet.getRange(1, 1, 1, ARCH_HEADERS.length)
      .setFontWeight('bold').setBackground('#F3F3F3').setFontColor('#000000');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ¿La última fila archivada ya es de hoy? Como siempre agregamos el bloque
// del día de una vez, basta mirar la fecha de la última fila.
function _archYaArchivadoHoy(sheet, hoy) {
  var last = sheet.getLastRow();
  if (last < 2) return false;
  var v = sheet.getRange(last, 1).getValue();
  var ep = (v instanceof Date) ? _archNormaliza(v) : 0;
  return ep === hoy;
}

// Fecha de hoy a las 00:00 (para comparar por día, sin hora).
function _archHoy0() {
  var d = new Date();
  return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
}

function _archNormaliza(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
}
