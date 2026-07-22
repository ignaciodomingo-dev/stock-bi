// ============================================================
// VELOCIDAD DE VENTA por SKU
// Lee "Ventas Históricas" y escribe la pestaña "Velocidad":
// una fila por SKU con unidades, fechas y velocidad diaria en
// tres ventanas (histórica / 180d / 90d) para poder calibrar el
// semáforo sin recalcular. Devoluciones ya vienen negativas.
//
// Uso: calcularVelocidad(). Base para el Semáforo de Quiebre.
// ============================================================

const VEL_CONFIG = {
  SHEET_VENTAS:   'Ventas Históricas',
  SHEET_STOCK:    'Stock Actual',         // fuente de verdad de SKUs vigentes
  SHEET_SALIDA:   'Velocidad',
  // Columnas de Ventas Históricas (1-based): Fecha=1, SKU=5, Cantidad=6
  COL_FECHA: 1, COL_SKU: 5, COL_CANT: 6,
  COL_STOCK_SKU: 3                        // columna SKU en Stock Actual
};

const VEL_HEADERS = [
  'SKU', 'Total Unidades', 'Primera Venta', 'Última Venta',
  'Días sin venta', 'Meses con venta',
  'Vel/día histórica', 'Vel/día 180d', 'Vel/día 90d',
  'Índice Próx 90d'   // estacionalidad: venta histórica de los próximos 3 meses calendario vs promedio (1 = normal)
];

function calcularVelocidad() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ventas = ss.getSheetByName(VEL_CONFIG.SHEET_VENTAS);
  if (!ventas || ventas.getLastRow() < 2) {
    throw new Error('❌ No hay datos en "' + VEL_CONFIG.SHEET_VENTAS + '".');
  }
  ss.toast('Calculando velocidad…', 'Velocidad', -1);

  var n = ventas.getLastRow() - 1;
  // Leemos solo las 6 primeras columnas (hasta Cantidad) para ahorrar memoria.
  var datos = ventas.getRange(2, 1, n, VEL_CONFIG.COL_CANT).getValues();

  var hoy   = _velHoy0();                 // hoy a las 00:00 (epoch ms)
  var c90   = hoy - 90  * 86400000;
  var c180  = hoy - 180 * 86400000;
  var DIA   = 86400000;

  var mapa = {};   // SKU → agregado
  datos.forEach(function(r) {
    var sku = String(r[VEL_CONFIG.COL_SKU - 1] || '').trim();
    if (!sku) return;                                   // líneas sin SKU (despachos/cargos)
    var cant = parseFloat(r[VEL_CONFIG.COL_CANT - 1] || 0);
    var ep   = _velEpoch(r[VEL_CONFIG.COL_FECHA - 1]);
    if (!ep) return;                                    // fecha ilegible

    var e = mapa[sku];
    if (!e) e = mapa[sku] = { total: 0, prim: ep, ult: ep, meses: {}, u90: 0, u180: 0,
                              umes: [0,0,0,0,0,0,0,0,0,0,0,0] };
    e.total += cant;
    if (ep < e.prim) e.prim = ep;
    if (ep > e.ult)  e.ult  = ep;
    e.meses[_velMes(ep)] = true;
    e.umes[new Date(ep).getMonth()] += cant;    // acumulado por mes calendario (estacionalidad)
    if (ep >= c180) e.u180 += cant;
    if (ep >= c90)  e.u90  += cant;
  });

  // Solo SKUs vigentes (los que Bsale lista en Stock Actual). El resto es
  // basura descontinuada que se cargó alguna vez y ya no existe.
  var vigentes = _velSkusEnStock();
  var omitidos = 0;
  var filas = Object.keys(mapa).filter(function(sku) {
    if (vigentes[sku]) return true;
    omitidos++; return false;
  }).map(function(sku) {
    var e = mapa[sku];
    var diasHist = Math.max(1, Math.round((hoy - e.prim) / DIA) + 1);

    // Índice de temporada próxima: proporción de la venta histórica que cae
    // en los próximos 3 meses calendario, normalizada (1 = venta uniforme;
    // 2 = ese trimestre vende el doble de lo normal). Solo significativo con
    // ≥12 meses de historia; si no, se deja vacío.
    var indice90 = '';
    if (e.total > 0 && diasHist >= 365) {
      var mesActual = new Date(hoy).getMonth();
      var uProx = 0;
      for (var k = 1; k <= 3; k++) uProx += e.umes[(mesActual + k) % 12];
      indice90 = Number(((uProx / e.total) * 4).toFixed(2));
    }

    return [
      sku,
      e.total,
      _velFecha(e.prim),
      _velFecha(e.ult),
      Math.round((hoy - e.ult) / DIA),
      Object.keys(e.meses).length,
      e.total / diasHist,
      e.u180 / 180,
      e.u90 / 90,
      indice90
    ];
  });
  // Más vendidos arriba.
  filas.sort(function(a, b) { return b[1] - a[1]; });

  var sheet = _velPrepararHoja();
  if (filas.length) {
    sheet.getRange(2, 1, filas.length, VEL_HEADERS.length).setValues(filas);
    sheet.getRange(2, 2, filas.length, 1).setNumberFormat('#,##0');           // total
    sheet.getRange(2, 7, filas.length, 4).setNumberFormat('#,##0.00');        // velocidades + índice
  }

  var msg = '✅ Velocidad: ' + filas.length + ' SKUs vigentes · ' + omitidos
          + ' descartados (no están en Stock Actual).';
  Logger.log(msg); ss.toast(msg, 'Velocidad', 12);
}

// SKUs vigentes según Stock Actual (lo que Bsale maneja hoy).
function _velSkusEnStock() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VEL_CONFIG.SHEET_STOCK);
  if (!sh || sh.getLastRow() < 2) {
    throw new Error('❌ Falta "' + VEL_CONFIG.SHEET_STOCK + '". Corre descargarStockActual() primero.');
  }
  var col = sh.getRange(2, VEL_CONFIG.COL_STOCK_SKU, sh.getLastRow() - 1, 1).getValues();
  var set = {};
  col.forEach(function(r) { var s = String(r[0] || '').trim(); if (s) set[s] = true; });
  return set;
}

// ------------------------------------------------------------
// HELPERS
// ------------------------------------------------------------
function _velPrepararHoja() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(VEL_CONFIG.SHEET_SALIDA);
  if (!sheet) sheet = ss.insertSheet(VEL_CONFIG.SHEET_SALIDA);
  sheet.clearContents();
  sheet.appendRow(VEL_HEADERS);
  sheet.getRange(1, 1, 1, VEL_HEADERS.length)
    .setFontWeight('bold').setBackground('#F3F3F3').setFontColor('#000000');
  sheet.setFrozenRows(1);
  return sheet;
}

function _velHoy0() {
  var d = new Date();
  return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
}

// Epoch (ms) desde una celda de fecha (Date o texto dd-mm-yyyy / yyyy-mm-dd).
function _velEpoch(v) {
  if (v instanceof Date) return v.getTime();
  var s = String(v || '').trim();
  if (!s) return 0;
  var p = s.split(/[\/\-.]/);
  if (p.length < 3) return 0;
  var d, m, y;
  if (p[0].length === 4) { y = +p[0]; m = +p[1]; d = +p[2]; }   // yyyy-mm-dd
  else { d = +p[0]; m = +p[1]; y = +p[2]; }                      // dd-mm-yyyy (es-CL)
  if (!y || !m || !d) return 0;
  return new Date(y, m - 1, d).getTime();
}

function _velFecha(ep) {
  var d = new Date(ep);
  return d.getFullYear() + '-' + _vel2(d.getMonth() + 1) + '-' + _vel2(d.getDate());
}

function _velMes(ep) {
  var d = new Date(ep);
  return d.getFullYear() + '-' + _vel2(d.getMonth() + 1);
}

function _vel2(n) { return (n < 10 ? '0' : '') + n; }
