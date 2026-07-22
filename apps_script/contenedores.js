// ============================================================
// CONTENEDORES EN CAMINO — despivote de "Resumen contenedores"
// La pestaña viene en formato ancho: pares de columnas SKU/Unidades
// por contenedor, con el código del contenedor y su fecha de llegada
// (ETA) en las filas sobre el encabezado.
//
// obtenerEnCamino() → { sku: { unidades, etaMs } } considerando SOLO
// contenedores con ETA futura (los llegados ya están en el stock de
// Bsale; sumarlos sería contar doble).
//
// [Supuesto] Contenedor sin fecha legible → se ignora (no se puede
// saber si llegó). Se registra en el log para completar la pestaña.
// ============================================================

const CONT_CONFIG = {
  SHEET: 'Resumen contenedores'
};

function obtenerEnCamino() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONT_CONFIG.SHEET);
  var map = {};
  if (!sh || sh.getLastRow() < 2) return map;
  var vals = sh.getDataRange().getValues();
  var hoy = _contHoy0();

  // Fila de encabezado: la que tiene celdas == 'SKU'.
  var hdr = -1;
  for (var i = 0; i < vals.length; i++) {
    if (vals[i].some(function(c) { return String(c).trim().toUpperCase() === 'SKU'; })) { hdr = i; break; }
  }
  if (hdr < 0) return map;

  var sinFecha = [];
  for (var c = 0; c < vals[hdr].length; c++) {
    if (String(vals[hdr][c]).trim().toUpperCase() !== 'SKU') continue;

    // Código y ETA: en las filas sobre el encabezado, columnas c y c+1.
    var code = '', etaMs = 0;
    for (var rr = 0; rr < hdr; rr++) {
      [c, c + 1].forEach(function(cc) {
        var cell = (vals[rr] || [])[cc];
        if (!code && /^[A-Z0-9]{2,6}[-\s]?\d/i.test(String(cell || '').trim())) code = String(cell).trim();
        if (!etaMs) { var e = _contEpoch(cell); if (e) etaMs = e; }
      });
    }

    if (!etaMs) { if (code) sinFecha.push(code); continue; }   // sin ETA → no se puede usar
    if (etaMs <= hoy) continue;                                 // ya llegó → está en Bsale

    // Filas de datos: SKU en col c, unidades en col c+1.
    for (var r2 = hdr + 1; r2 < vals.length; r2++) {
      var sku = String((vals[r2] || [])[c] || '').trim();
      if (!sku) continue;
      var uds = parseFloat((vals[r2] || [])[c + 1] || 0);
      if (!(uds > 0)) continue;
      var e2 = map[sku] || (map[sku] = { unidades: 0, etaMs: etaMs });
      e2.unidades += uds;
      if (etaMs < e2.etaMs) e2.etaMs = etaMs;   // la ETA más próxima
    }
  }

  if (sinFecha.length) {
    Logger.log('⚠️ Contenedores sin fecha de llegada legible (ignorados como tránsito): '
               + sinFecha.join(', ') + '. Completar la fecha en la pestaña para que cuenten.');
  }
  Logger.log('🚢 En camino: ' + Object.keys(map).length + ' SKUs con unidades en tránsito.');
  return map;
}

// ------------------------------------------------------------
// HELPERS
// ------------------------------------------------------------
function _contHoy0() {
  var d = new Date();
  return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
}

// Fecha desde celda (Date o texto dd-mm-yyyy / yyyy-mm-dd). 0 si no parsea.
function _contEpoch(v) {
  if (v instanceof Date) return v.getTime();
  var s = String(v || '').trim();
  if (!/\d{1,4}[\/\-.]\d{1,2}[\/\-.]\d{1,4}/.test(s)) return 0;
  var p = s.split(/[\/\-.]/);
  if (p.length < 3) return 0;
  var d, m, y;
  if (p[0].length === 4) { y = +p[0]; m = +p[1]; d = +p[2]; }
  else { d = +p[0]; m = +p[1]; y = +p[2]; }
  if (!y || !m || !d || y < 2020 || y > 2100 || m > 12 || d > 31) return 0;
  return new Date(y, m - 1, d).getTime();
}
