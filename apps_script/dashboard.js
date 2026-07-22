// ============================================================
// DASHBOARD WEB — BI de stock y ventas como web app
// Sirve dashboard.html (doGet) y le entrega los datos agregados
// (getDatosDashboard). Todo se calcula server-side para que el
// navegador reciba un payload chico.
//
// Deploy (una vez): Implementar → Nueva implementación → Aplicación
// web → Ejecutar como: yo · Acceso: cualquiera con el vínculo.
// Actualizaciones de código posteriores: clasp push basta si se usa
// "Administrar implementaciones → editar → nueva versión".
// ============================================================

const DASH_CONFIG = {
  SHEET_SEM: 'Semáforo',
  SHEET_VM:  'Ventas Mensuales',
  SHEET_VEL: 'Velocidad',
  TOP_N: 20,
  MESES_TOP: 3,        // ventana del ranking de productos
  MESES_TENDENCIA: 36  // meses de historia en el gráfico
};

function doGet() {
  var tpl = HtmlService.createTemplateFromFile('dashboard_ui');
  tpl.marca = (typeof NEGOCIO !== 'undefined' && NEGOCIO.MARCA) ? NEGOCIO.MARCA : 'Stock BI';
  return tpl.evaluate()
    .setTitle('Stock BI — ' + tpl.marca)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Payload completo para el front. Una sola llamada por carga de página.
function getDatosDashboard() {
  var sem = _dashLeerSemaforo();
  var vm  = _dashLeerVentasMensuales();

  // KPIs
  var kpis = { rojos: 0, amarillos: 0, enCamino: 0, verdes: 0, sinRotacion: 0, total: sem.length };
  sem.forEach(function(r) {
    var e = r.estado;
    if (e.indexOf('🔴') === 0)      kpis.rojos++;
    else if (e.indexOf('🟡') === 0) kpis.amarillos++;
    else if (e.indexOf('🔵') === 0) kpis.enCamino++;
    else if (e.indexOf('🟢') === 0) kpis.verdes++;
    else if (e.indexOf('⚫') === 0) kpis.sinRotacion++;
  });

  // Tendencia mensual (todos los SKUs vigentes) — últimos N meses.
  var porMes = {};
  vm.forEach(function(r) {
    if (r.vigente !== 'Sí') return;
    var e = porMes[r.mes] || (porMes[r.mes] = { u: 0, m: 0 });
    e.u += r.unidades; e.m += r.monto;
  });
  var meses = Object.keys(porMes).sort().slice(-DASH_CONFIG.MESES_TENDENCIA);
  var tendencia = meses.map(function(m) {
    return { mes: m, unidades: porMes[m].u, monto: porMes[m].m };
  });

  // Top SKUs últimos N meses.
  var corte = meses.slice(-DASH_CONFIG.MESES_TOP);
  var porSku = {};
  vm.forEach(function(r) {
    if (corte.indexOf(r.mes) === -1) return;
    var e = porSku[r.sku] || (porSku[r.sku] = { u: 0, m: 0 });
    e.u += r.unidades; e.m += r.monto;
  });
  var top = Object.keys(porSku).map(function(sku) {
    return { sku: sku, unidades: porSku[sku].u, monto: porSku[sku].m };
  }).sort(function(a, b) { return b.unidades - a.unidades; })
    .slice(0, DASH_CONFIG.TOP_N);

  return {
    actualizado: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm'),
    kpis: kpis,
    semaforo: sem,          // tabla completa (~430 filas, liviana)
    tendencia: tendencia,
    top: top,
    ventanaTop: DASH_CONFIG.MESES_TOP
  };
}

// ------------------------------------------------------------
// LECTURAS
// ------------------------------------------------------------
function _dashLeerSemaforo() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DASH_CONFIG.SHEET_SEM);
  if (!sh || sh.getLastRow() < 2) return [];
  // Cols (1-based): 1 SKU · 2 Desc · 3 Stock · 4 Vel · 5 Cobertura · 6 EnCamino ·
  // 7 ETA · 8 CobTránsito · 9 Índice90 · 10 Origen · 11 Lead · 12 Estado
  return sh.getRange(2, 1, sh.getLastRow() - 1, 12).getValues().map(function(r) {
    return {
      sku: String(r[0]), desc: String(r[1]), stock: Number(r[2]) || 0,
      vel: Number(r[3]) || 0, cobertura: r[4] === '' ? null : Number(r[4]),
      enCamino: Number(r[5]) || 0, eta: r[6] === '' ? null : Number(r[6]),
      cobTransito: r[7] === '' ? null : Number(r[7]),
      idx90: r[8] === '' ? null : Number(r[8]),
      origen: String(r[9]), lead: Number(r[10]) || null, estado: String(r[11])
    };
  });
}

function _dashLeerVentasMensuales() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DASH_CONFIG.SHEET_VM);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().map(function(r) {
    return {
      mes: String(r[0]), sku: String(r[1]),
      unidades: Number(r[2]) || 0, monto: Number(r[3]) || 0, vigente: String(r[4])
    };
  });
}
