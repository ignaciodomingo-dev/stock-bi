// ============================================================
// SEMÁFORO DE QUIEBRE
// Cruza Stock Actual (stock vendible) ÷ Velocidad (180d) = días de
// cobertura, y compara contra el lead time del origen del SKU (según
// NEGOCIO.ORIGEN_LEAD) para pintar rojo/amarillo/verde.
//
// Origen: catálogo → prefijo de contenedor → prefijo/keyword de SKU →
// defecto. Los mapeos de proveedor viven en NEGOCIO.* (config_local.js).
//
// Requiere: Stock Actual y Velocidad ya calculados.
// Uso: calcularSemaforo().
// ============================================================

const SEM_CONFIG = {
  SHEET_STOCK:    'Stock Actual',
  SHEET_VEL:      'Velocidad',
  SHEET_CONT:     'Resumen contenedores',
  SHEET_CATALOGO: 'Config SKU',
  SHEET_DESCONT:  'Descontinuados',   // lista manual de SKUs a excluir (col A)
  SHEET_SALIDA:   'Semáforo',

  // Config SKU (1-based): SKU=1, Origen=6  (cruce directo con el SKU de Bsale)
  CAT_SKU: 1, CAT_ORIGEN: 6,
  // Stock Actual (1-based): SKU=3, Descripción=4, Tipo Oficina=7, Stock Disponible=10
  ST_SKU: 3, ST_DESC: 4, ST_TIPO: 7, ST_DISP: 10,
  // Velocidad (1-based): SKU=1, Primera=3, DíasSinVenta=5, Meses=6, Vel180=8, Índice90=10
  VL_SKU: 1, VL_PRIM: 3, VL_DSV: 5, VL_MESES: 6, VL_V180: 8, VL_IDX90: 10,

  VENTANA_DIAS: 180,
  COLCHON:      30,        // amarillo = lead .. lead+colchón
  NUEVO_MESES:  2,         // ≤ 2 meses con venta → marcar "(nuevo)"
  DESCONT_DIAS: 365,       // > 12 meses sin venta → "posible descontinuado"

  // Alertas por email. El destinatario NO va en el código (dato de empresa):
  // se toma de la Script Property PROP_EMAIL o, si no existe, del dueño del
  // script (Session.getEffectiveUser). PROP_ROJOS guarda el set de rojos de
  // la corrida anterior para avisar SOLO los cambios, sin spam.
  PROP_EMAIL:  'EMAIL_ALERTAS',
  PROP_ROJOS:  'SEM_ROJOS_PREV'
};

// Origen, lead times, prefijos y keywords viven en NEGOCIO.* (config_local.js),
// se leen en runtime dentro de las funciones (no como literales en el repo).

const SEM_HEADERS = [
  'SKU', 'Descripción', 'Stock Vendible', 'Vel/día (180d aj.)', 'Días Cobertura',
  'En Camino', 'ETA (días)', 'Cobertura c/tránsito', 'Índice Próx 90d',
  'Origen', 'Lead (días)', 'Estado', 'Días sin venta', 'Meses con venta', 'Origen (fuente)'
];
// Índices 0-based de la fila: 0 SKU · 1 Desc · 2 Stock · 3 Vel · 4 Cobertura ·
// 5 EnCamino · 6 ETA · 7 CobTránsito · 8 Índice90 · 9 Origen · 10 Lead ·
// 11 Estado · 12 DSV · 13 Meses · 14 Fuente
const SEM_COL_ESTADO = 11;

function calcularSemaforo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Calculando semáforo…', 'Semáforo', -1);

  var stock = _semStockVendible();     // SKU → {disp, desc}
  var vel   = _semVelocidad();         // SKU → {prim, dsv, meses, v180}
  var cat   = _semOrigenCatalogo();    // SKU → país (Config SKU col F) — autoritativo
  var catKeys = Object.keys(cat).sort(function(a, b) { return b.length - a.length; }); // largo→corto
  var cont  = _semOrigenContenedor();  // SKU → país (respaldo, prefijo contenedor)
  var descont = _semDescontinuados();  // SKUs confirmados muertos → fuera del semáforo
  var enCamino = obtenerEnCamino();    // SKU → {unidades, etaMs} (contenedores.js)
  var hoy   = _semHoy0();
  var DIA   = 86400000;
  var totalStock = Object.keys(stock).length;

  var filas = Object.keys(stock).filter(function(sku) {
    return !descont[sku];
  }).map(function(sku) {
    var s = stock[sku];
    var v = vel[sku] || null;

    // Origen + lead. Prioridad: catálogo exacto → catálogo por prefijo →
    // contenedor → palabra clave → defecto China.
    var origen, fuente, pref;
    if (cat[sku])                            { origen = cat[sku]; fuente = 'catálogo'; }
    else if ((pref = _semPrefijoOrigen(sku, cat, catKeys))) { origen = pref; fuente = 'catálogo (prefijo)'; }
    else if (cont[sku])                      { origen = cont[sku]; fuente = 'contenedor'; }
    else if (_semEsIndiaPorPrefijo(sku))     { origen = 'India'; fuente = 'prefijo SKU'; }
    else if (_semEsIndiaPorTexto(s.desc) || _semEsIndiaPorTexto(sku)) { origen = 'India'; fuente = 'palabra clave'; }
    else { origen = 'China'; fuente = 'defecto'; }
    var lead = NEGOCIO.ORIGEN_LEAD[origen];

    // Velocidad ajustada (para SKU nuevo: divide por su edad real, no por 180)
    var velAj = 0;
    if (v && v.v180 > 0) {
      var u180 = v.v180 * SEM_CONFIG.VENTANA_DIAS;                 // unidades en la ventana
      var edad = v.prim ? Math.max(1, Math.round((hoy - v.prim) / DIA) + 1) : SEM_CONFIG.VENTANA_DIAS;
      velAj = u180 / Math.min(SEM_CONFIG.VENTANA_DIAS, edad);
    }

    var dsv   = v ? v.dsv : '';
    var meses = v ? v.meses : 0;
    var idx90 = v ? v.idx90 : '';

    // Tránsito: unidades en contenedores con ETA futura.
    var ec = enCamino[sku] || null;
    var ecUds  = ec ? ec.unidades : 0;
    var etaDias = ec ? Math.max(0, Math.round((ec.etaMs - hoy) / DIA)) : '';

    // Estado. La urgencia de PEDIR usa cobertura c/tránsito (stock + lo que
    // ya viene); "Quiebre hoy" sigue siendo operativo (no hay stock que vender).
    var estado, cobertura = '', cobTransito = '';
    if (s.disp <= 0) {
      estado = ecUds > 0 ? '🔴 Quiebre hoy (llega en ' + etaDias + 'd)' : '🔴 Quiebre hoy';
    } else if (lead == null) {
      estado = '⚠️ Falta lead (' + origen + ')';   // país sin lead time configurado
    } else if (dsv !== '' && dsv > SEM_CONFIG.DESCONT_DIAS) {
      estado = '🚫 Posible descontinuado';          // >12 meses sin venta
    } else if (velAj <= 0) {
      estado = '⚫ Sin rotación';
    } else {
      cobertura   = Math.round(s.disp / velAj);
      cobTransito = Math.round((s.disp + ecUds) / velAj);
      if (cobTransito < lead)                             estado = '🔴 Rojo';
      else if (cobTransito <= lead + SEM_CONFIG.COLCHON)  estado = '🟡 Amarillo';
      else if (cobertura < lead && ecUds > 0)             estado = '🔵 En camino';  // salvado por el contenedor
      else                                                estado = '🟢 Verde';
      if (meses && meses <= SEM_CONFIG.NUEVO_MESES) estado += ' (nuevo)';
    }

    return [sku, s.desc, s.disp, velAj ? Number(velAj.toFixed(2)) : 0, cobertura,
            ecUds || '', etaDias, cobTransito, idx90,
            origen, lead, estado, dsv, meses, fuente];
  });

  // Orden por urgencia (match por prefijo: tolera sufijos como "(nuevo)" o la ETA).
  function _prio(estado) {
    var e = String(estado);
    if (e.indexOf('🔴 Quiebre') === 0) return 0;
    if (e.indexOf('🔴 Rojo') === 0)    return 1;
    if (e.indexOf('🟡') === 0)         return 2;
    if (e.indexOf('🔵') === 0)         return 3;
    if (e.indexOf('🟢') === 0)         return 4;
    if (e.indexOf('⚫') === 0)         return 5;
    if (e.indexOf('🚫') === 0)         return 6;
    return 9;
  }
  filas.sort(function(a, b) {
    var pa = _prio(a[SEM_COL_ESTADO]), pb = _prio(b[SEM_COL_ESTADO]);
    if (pa !== pb) return pa - pb;
    // menor cobertura c/tránsito primero (la urgencia real de pedir)
    return (a[7] === '' ? 1e9 : a[7]) - (b[7] === '' ? 1e9 : b[7]);
  });

  var sheet = _semPrepararHoja();
  if (filas.length) {
    sheet.getRange(2, 1, filas.length, SEM_HEADERS.length).setValues(filas);
    sheet.getRange(2, 3, filas.length, 1).setNumberFormat('#,##0');   // stock
    sheet.getRange(2, 6, filas.length, 1).setNumberFormat('#,##0');   // en camino
    _semFormatoColor(sheet, filas.length);
  }

  // Alerta por email de SKUs que ENTRAN nuevos en rojo (no rompe el semáforo
  // si el correo falla).
  try { _semAlertarNuevosRojos(filas); }
  catch (e) { Logger.log('⚠️ Falló alerta por email (semáforo OK igual): ' + e.message); }

  var msg = '✅ Semáforo: ' + filas.length + ' SKUs activos · '
          + (totalStock - filas.length) + ' descontinuados excluidos.';
  Logger.log(msg); ss.toast(msg, 'Semáforo', 12);
}

// ------------------------------------------------------------
// ALERTAS POR EMAIL
// ------------------------------------------------------------
// Avisa por correo los SKUs que entran NUEVOS en rojo (🔴) respecto de la
// corrida anterior. Sin spam: solo los cambios. El set de rojos previo se
// guarda en Script Properties. La primera corrida solo siembra el estado
// (no manda correo, para evitar un mail gigante inicial).
function _semAlertarNuevosRojos(filas) {
  var props = PropertiesService.getScriptProperties();

  // Rojos de ahora: estado con 🔴 (Quiebre hoy o Rojo). SKU → fila.
  var rojosAhora = {};
  filas.forEach(function(f) {
    if (String(f[SEM_COL_ESTADO]).indexOf('🔴') !== -1) rojosAhora[f[0]] = f;
  });
  var listaAhora = Object.keys(rojosAhora);

  var prevRaw = props.getProperty(SEM_CONFIG.PROP_ROJOS);
  if (prevRaw == null) {   // primera corrida: sembrar y salir
    props.setProperty(SEM_CONFIG.PROP_ROJOS, JSON.stringify(listaAhora));
    Logger.log('ℹ️ Alertas: estado inicial sembrado (' + listaAhora.length + ' rojos). Sin correo.');
    return;
  }

  var prev = {};
  try { JSON.parse(prevRaw).forEach(function(s) { prev[s] = true; }); } catch (e) {}
  var nuevos = listaAhora.filter(function(sku) { return !prev[sku]; });

  // Guardar el estado actual SIEMPRE (haya o no nuevos rojos).
  props.setProperty(SEM_CONFIG.PROP_ROJOS, JSON.stringify(listaAhora));

  if (!nuevos.length) { Logger.log('ℹ️ Alertas: sin nuevos rojos.'); return; }

  // Destinatario en runtime (no en el repo): property > dueño del script.
  var dest = props.getProperty(SEM_CONFIG.PROP_EMAIL) || Session.getEffectiveUser().getEmail();
  if (!dest) { Logger.log('⚠️ Alertas: sin destinatario (define la property ' + SEM_CONFIG.PROP_EMAIL + ').'); return; }

  var hoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy');
  var celda = 'border:1px solid #ccc;padding:4px 8px';
  var rows = nuevos.map(function(sku) {
    var f = rojosAhora[sku];
    // SKU | Descripción | Stock | Cobertura(días) | Origen | Estado
    return '<tr>'
      + '<td style="' + celda + '"><b>' + f[0] + '</b></td>'
      + '<td style="' + celda + '">' + f[1] + '</td>'
      + '<td style="' + celda + ';text-align:right">' + f[2] + '</td>'
      + '<td style="' + celda + ';text-align:right">' + (f[4] === '' ? '–' : f[4]) + '</td>'
      + '<td style="' + celda + '">' + f[9] + '</td>'
      + '<td style="' + celda + '">' + f[SEM_COL_ESTADO] + '</td>'
      + '</tr>';
  }).join('');

  var html = '<p>' + nuevos.length + ' SKU(s) entraron en <b>rojo</b> hoy (' + hoy + '):</p>'
    + '<table style="border-collapse:collapse;font-family:Arial;font-size:13px">'
    + '<tr style="background:#F3F3F3;color:#000">'
    + '<th style="' + celda + '">SKU</th><th style="' + celda + '">Descripción</th>'
    + '<th style="' + celda + '">Stock</th><th style="' + celda + '">Cobertura (d)</th>'
    + '<th style="' + celda + '">Origen</th><th style="' + celda + '">Estado</th></tr>'
    + rows + '</table>'
    + '<p style="color:#666;font-size:12px">Semáforo de Quiebre · ' + NEGOCIO.MARCA + '. '
    + 'Solo se avisan cambios (SKUs que entran nuevos en rojo).</p>';

  MailApp.sendEmail({
    to: dest,
    subject: '🔴 Semáforo ' + NEGOCIO.MARCA + ': ' + nuevos.length + ' SKU(s) nuevos en quiebre (' + hoy + ')',
    htmlBody: html
  });
  Logger.log('📧 Alerta enviada: ' + nuevos.length + ' nuevos rojos → ' + dest);
}

// Lista manual de SKUs descontinuados (pestaña "Descontinuados", col A).
// Si no existe, la crea vacía para que el usuario pegue ahí los confirmados.
function _semDescontinuados() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SEM_CONFIG.SHEET_DESCONT);
  if (!sh) {
    sh = ss.insertSheet(SEM_CONFIG.SHEET_DESCONT);
    sh.appendRow(['SKU (descontinuado)']);
    sh.getRange(1, 1).setFontWeight('bold').setBackground('#F3F3F3').setFontColor('#000000');
    sh.setFrozenRows(1);
    return {};
  }
  var set = {};
  if (sh.getLastRow() < 2) return set;
  sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function(r) {
    var s = String(r[0] || '').trim();
    if (s) set[s] = true;
  });
  return set;
}

// ------------------------------------------------------------
// LECTURAS
// ------------------------------------------------------------
function _semStockVendible() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SEM_CONFIG.SHEET_STOCK);
  if (!sh || sh.getLastRow() < 2) throw new Error('❌ Falta "' + SEM_CONFIG.SHEET_STOCK + '".');
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, SEM_CONFIG.ST_DISP).getValues();
  var map = {};
  v.forEach(function(r) {
    if (String(r[SEM_CONFIG.ST_TIPO - 1]).trim() !== 'Principal') return;   // excluye oficinas secundarias
    var sku = String(r[SEM_CONFIG.ST_SKU - 1] || '').trim();
    if (!sku) return;
    var e = map[sku] || (map[sku] = { disp: 0, desc: '' });
    e.disp += parseFloat(r[SEM_CONFIG.ST_DISP - 1] || 0);
    if (!e.desc) e.desc = String(r[SEM_CONFIG.ST_DESC - 1] || '');
  });
  return map;
}

function _semVelocidad() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SEM_CONFIG.SHEET_VEL);
  if (!sh || sh.getLastRow() < 2) throw new Error('❌ Falta "' + SEM_CONFIG.SHEET_VEL + '". Corre calcularVelocidad().');
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, SEM_CONFIG.VL_IDX90).getValues();
  var map = {};
  v.forEach(function(r) {
    var sku = String(r[SEM_CONFIG.VL_SKU - 1] || '').trim();
    if (!sku) return;
    map[sku] = {
      prim:  _semEpoch(r[SEM_CONFIG.VL_PRIM - 1]),
      dsv:   Number(r[SEM_CONFIG.VL_DSV - 1] || 0),
      meses: Number(r[SEM_CONFIG.VL_MESES - 1] || 0),
      v180:  parseFloat(r[SEM_CONFIG.VL_V180 - 1] || 0),
      idx90: r[SEM_CONFIG.VL_IDX90 - 1] === '' ? '' : Number(r[SEM_CONFIG.VL_IDX90 - 1] || '')
    };
  });
  return map;
}

// Origen autoritativo desde Config SKU (col A = SKU, col F = país).
// Solo filas con país (saltea encabezados de categoría y la fila "Intermediario").
function _semOrigenCatalogo() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SEM_CONFIG.SHEET_CATALOGO);
  var map = {};
  if (!sh || sh.getLastRow() < 2) return map;
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, SEM_CONFIG.CAT_ORIGEN).getValues();
  v.forEach(function(r) {
    var sku  = String(r[SEM_CONFIG.CAT_SKU - 1] || '').trim();
    var pais = String(r[SEM_CONFIG.CAT_ORIGEN - 1] || '').trim();
    if (sku && pais) map[sku] = _semTitle(pais);
  });
  return map;
}

function _semTitle(s) { return s.charAt(0).toUpperCase() + s.slice(1).toLowerCase(); }

// Match por prefijo: si el SKU empieza con un SKU del catálogo (seguido de
// espacio o '-'), hereda su origen. catKeys viene ordenado largo→corto para
// tomar el match más específico. Evita falsos al exigir borde de palabra.
function _semPrefijoOrigen(sku, cat, catKeys) {
  for (var i = 0; i < catKeys.length; i++) {
    var k = catKeys[i];
    if (k.length >= 3 && sku.length > k.length && sku.indexOf(k) === 0) {
      var nxt = sku.charAt(k.length);
      if (nxt === ' ' || nxt === '-') return cat[k];
    }
  }
  return null;
}

// Parsea el formato ancho de "Resumen contenedores" → SKU → origen.
function _semOrigenContenedor() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SEM_CONFIG.SHEET_CONT);
  var map = {};
  if (!sh || sh.getLastRow() < 2) return map;
  var vals = sh.getDataRange().getValues();

  // Fila de encabezado: la que tiene celdas == 'SKU'.
  var hdr = -1;
  for (var i = 0; i < vals.length; i++) {
    if (vals[i].some(function(c) { return String(c).trim().toUpperCase() === 'SKU'; })) { hdr = i; break; }
  }
  if (hdr < 0) return map;

  for (var c = 0; c < vals[hdr].length; c++) {
    if (String(vals[hdr][c]).trim().toUpperCase() !== 'SKU') continue;
    // Código del contenedor: arriba del encabezado, en la columna de al lado (c+1).
    var code = '';
    for (var rr = 0; rr < hdr; rr++) {
      var cell = String((vals[rr] || [])[c + 1] || '').trim();
      if (/^[A-Z]{2,4}-?\d/i.test(cell)) { code = cell; break; }   // código tipo "AAA-123" / "AAA123"
    }
    var prefijo = code.split(/[-\s]/)[0].toUpperCase();
    var origen = NEGOCIO.PREFIJO_ORIGEN[prefijo];
    if (!origen) continue;                       // prefijo desconocido → al fallback
    for (var rr2 = hdr + 1; rr2 < vals.length; rr2++) {
      var sku = String((vals[rr2] || [])[c] || '').trim();
      if (sku && !map[sku]) map[sku] = origen;   // primero gana
    }
  }
  return map;
}

// ------------------------------------------------------------
// HELPERS
// ------------------------------------------------------------
function _semEsIndiaPorPrefijo(sku) {
  var s = String(sku || '').trim().toUpperCase();
  return NEGOCIO.SKU_PREFIJOS_INDIA.some(function(p) {
    return s.indexOf(p) === 0 && /\d/.test(s.charAt(p.length));   // prefijo + dígito
  });
}

function _semEsIndiaPorTexto(txt) {
  var t = String(txt || '').toUpperCase();
  return NEGOCIO.KEYWORDS_INDIA.some(function(k) { return t.indexOf(k) !== -1; });
}

function _semPrepararHoja() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SEM_CONFIG.SHEET_SALIDA);
  if (!sheet) sheet = ss.insertSheet(SEM_CONFIG.SHEET_SALIDA);
  sheet.clearContents();
  sheet.clearConditionalFormatRules();
  sheet.appendRow(SEM_HEADERS);
  sheet.getRange(1, 1, 1, SEM_HEADERS.length)
    .setFontWeight('bold').setBackground('#F3F3F3').setFontColor('#000000');
  sheet.setFrozenRows(1);
  return sheet;
}

function _semFormatoColor(sheet, n) {
  var col = sheet.getRange(2, SEM_COL_ESTADO + 1, n, 1);   // columna Estado (1-based)
  function regla(texto, bg) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains(texto).setBackground(bg).setRanges([col]).build();
  }
  sheet.setConditionalFormatRules([
    regla('Quiebre', '#F4C7C3'),
    regla('Rojo',    '#F4C7C3'),
    regla('Amarillo','#FCE8B2'),
    regla('En camino', '#CFE2F3'),
    regla('Verde',   '#B7E1CD'),
    regla('Sin rotación', '#D9D9D9'),
    regla('descontinuado', '#B7B7B7')
  ]);
  sheet.autoResizeColumns(1, SEM_HEADERS.length);
}

function _semHoy0() {
  var d = new Date();
  return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
}

function _semEpoch(v) {
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
