// ============================================================
// VENTAS MENSUALES — agregado para Looker Studio
// Lee "Ventas Históricas" (~245k filas, demasiado para Looker) y
// escribe "Ventas Mensuales": una fila por Mes × SKU con unidades
// y monto neto. Tabla larga, encabezado en fila 1 — formato que
// Looker Studio consume directo.
//
// Corre a diario dentro del pipeline (ver pipeline.js) o manual:
// calcularVentasMensuales().
// ============================================================

const VM_CONFIG = {
  SHEET_VENTAS: 'Ventas Históricas',
  SHEET_STOCK:  'Stock Actual',
  SHEET_SALIDA: 'Ventas Mensuales',
  // Columnas de Ventas Históricas (1-based): Fecha=1, SKU=5, Cantidad=6, TotalNeto=8
  COL_FECHA: 1, COL_SKU: 5, COL_CANT: 6, COL_TOTAL: 8,
  COL_STOCK_SKU: 3
};

const VM_HEADERS = ['Mes', 'SKU', 'Unidades', 'Monto Neto', 'Vigente'];

function calcularVentasMensuales() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ventas = ss.getSheetByName(VM_CONFIG.SHEET_VENTAS);
  if (!ventas || ventas.getLastRow() < 2) {
    throw new Error('❌ No hay datos en "' + VM_CONFIG.SHEET_VENTAS + '".');
  }
  ss.toast('Agregando ventas mensuales…', 'Ventas Mensuales', -1);

  var n = ventas.getLastRow() - 1;
  var datos = ventas.getRange(2, 1, n, VM_CONFIG.COL_TOTAL).getValues();

  // clave "yyyy-MM|SKU" → {u: unidades, m: monto}
  var agg = {};
  datos.forEach(function(r) {
    var sku = String(r[VM_CONFIG.COL_SKU - 1] || '').trim();
    if (!sku) return;
    var ep = _vmEpoch(r[VM_CONFIG.COL_FECHA - 1]);
    if (!ep) return;
    var d = new Date(ep);
    var mes = d.getFullYear() + '-' + _vm2(d.getMonth() + 1);
    var k = mes + '|' + sku;
    var e = agg[k] || (agg[k] = { u: 0, m: 0 });
    e.u += parseFloat(r[VM_CONFIG.COL_CANT - 1] || 0);
    e.m += parseFloat(r[VM_CONFIG.COL_TOTAL - 1] || 0);
  });

  var vigentes = _vmSkusEnStock();
  var filas = Object.keys(agg).map(function(k) {
    var p = k.split('|');
    return [p[0], p[1], agg[k].u, Math.round(agg[k].m), vigentes[p[1]] ? 'Sí' : 'No'];
  });
  // Orden: mes descendente, luego SKU.
  filas.sort(function(a, b) {
    if (a[0] !== b[0]) return a[0] < b[0] ? 1 : -1;
    return a[1] < b[1] ? -1 : 1;
  });

  // Escritura atómica: hoja temporal → renombrar (si algo falla, la
  // versión anterior queda intacta y Looker no ve una tabla a medias).
  var tmpName = VM_CONFIG.SHEET_SALIDA + '_tmp';
  var tmp = ss.getSheetByName(tmpName);
  if (tmp) ss.deleteSheet(tmp);
  tmp = ss.insertSheet(tmpName);
  tmp.getRange(1, 1, 1, VM_HEADERS.length).setValues([VM_HEADERS])
    .setFontWeight('bold').setBackground('#F3F3F3').setFontColor('#000000');
  if (filas.length) {
    tmp.getRange(2, 1, filas.length, VM_HEADERS.length).setValues(filas);
    tmp.getRange(2, 3, filas.length, 2).setNumberFormat('#,##0');
  }
  tmp.setFrozenRows(1);

  var vieja = ss.getSheetByName(VM_CONFIG.SHEET_SALIDA);
  if (vieja) ss.deleteSheet(vieja);
  tmp.setName(VM_CONFIG.SHEET_SALIDA);

  var msg = '✅ Ventas Mensuales: ' + filas.length + ' filas (mes × SKU).';
  Logger.log(msg); ss.toast(msg, 'Ventas Mensuales', 10);
}

function _vmSkusEnStock() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VM_CONFIG.SHEET_STOCK);
  var set = {};
  if (!sh || sh.getLastRow() < 2) return set;
  sh.getRange(2, VM_CONFIG.COL_STOCK_SKU, sh.getLastRow() - 1, 1).getValues()
    .forEach(function(r) { var s = String(r[0] || '').trim(); if (s) set[s] = true; });
  return set;
}

// ------------------------------------------------------------
// HELPERS
// ------------------------------------------------------------
function _vm2(n) { return (n < 10 ? '0' : '') + n; }

function _vmEpoch(v) {
  if (v instanceof Date) return v.getTime();
  var s = String(v || '').trim();
  if (!s) return 0;
  var p = s.split(/[\/\-.]/);
  if (p.length < 3) return 0;
  var d, m, y;
  if (p[0].length === 4) { y = +p[0]; m = +p[1]; d = +p[2]; }
  else { d = +p[0]; m = +p[1]; y = +p[2]; }
  if (!y || !m || !d) return 0;
  return new Date(y, m - 1, d).getTime();
}
