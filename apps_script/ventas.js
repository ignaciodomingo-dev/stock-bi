// ============================================================
// DESCARGA RÁPIDA DE VENTAS — de lo más nuevo a lo más antiguo
// Reescritura del Script 2 para terminar en minutos, no días:
//   • Descarga en PARALELO (UrlFetchApp.fetchAll, ~20 páginas/ronda).
//   • Recorre del documento más NUEVO al más viejo.
//   • Llena "todo lo que quepa": se detiene al acercarse al tope de
//     celdas (MAX_FILAS) o al llegar a 2021. Así te quedas siempre
//     con la venta más reciente, que es la que pesa para velocidad.
//   • Estructura cruda completa (8 columnas).
//   • Resumible: corta a los 5 min y se reejecuta (o vía trigger).
//
// Requisito: BSALE_TOKEN en Script Properties.
// Uso: descargarVentasRapido() (la 1ª vez LIMPIA la hoja y arranca
// desde lo más nuevo). Reejecuta o instala el trigger hasta "✅ …".
// ============================================================

const VR_CONFIG = {
  TOKEN_PROPERTY: 'BSALE_TOKEN',
  SHEET:          'Ventas Históricas',
  LIMIT:          50,                 // tope de página de la API Bsale
  BATCH_PAGES:    20,                 // páginas en paralelo por ronda → 1.000 docs/ronda
  MAX_MS:         5 * 60 * 1000,
  MAX_FILAS:      1100000,            // tope de seguridad (~8,8M celdas con 8 cols)
  FECHA_DESDE:    1609459200,         // 2021-01-01 (corte inferior)
  TIPOS_VENTA:     [1, 5, 11, 16, 22],
  TIPOS_DEVOLUCION:[2],

  // claves de progreso
  COUNT_KEY: 'VR_COUNT',
  DIR_KEY:   'VR_DIR',      // 'asc' (más viejo primero) | 'desc'
  IDX_KEY:   'VR_IDX',      // índice de página en el recorrido nuevo→viejo
  MAXID_KEY: 'VR_MAX_ID',   // id del doc más nuevo ya guardado (para actualización incremental)
  DONE_KEY:  'VR_DONE'      // marca de backfill terminado (evita que el trigger lo reinicie)
};

const VR_TIPO_NOMBRES = {
  1: 'BOLETA ELECTRÓNICA T', 2: 'NOTA DE CRÉDITO', 5: 'FACTURA ELECTRÓNICA',
  11: 'FACTURA MANUAL', 16: 'FACTURA EXPORTACIÓN', 22: 'BOLETA ELECTRÓNICA'
};

const VR_HEADERS = [
  'Fecha Emisión', 'N° Documento', 'Tipo Documento', 'Tipo Movimiento',
  'SKU', 'Cantidad', 'Precio Unit Neto', 'Total Neto'
];

// ------------------------------------------------------------
// PRINCIPAL — ejecutar (o vía trigger) hasta "✅ Descarga completa"
// ------------------------------------------------------------
function descargarVentasRapido() {
  var token = PropertiesService.getScriptProperties().getProperty(VR_CONFIG.TOKEN_PROPERTY);
  if (!token) throw new Error('❌ No existe BSALE_TOKEN en Script Properties.');

  var props = PropertiesService.getScriptProperties();
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var idx   = parseInt(props.getProperty(VR_CONFIG.IDX_KEY) || '-1');

  var count, dir, sheet;
  if (idx < 0) {
    // Ya completado: NO reiniciar (esto evita que el trigger vuelva a vaciar la hoja).
    if (props.getProperty(VR_CONFIG.DONE_KEY)) {
      var ya = 'ℹ️ Backfill ya completo. Para rehacerlo desde cero: reiniciarVentasRapido(). '
             + 'Lo que debe quedar activo es la actualización diaria (actualizarVentasRapido).';
      Logger.log(ya); SpreadsheetApp.getActiveSpreadsheet().toast(ya, 'Ventas', 12);
      return;
    }
    // Arranque limpio: detectar total + orden, y vaciar la hoja.
    var info = _vrInicializar(token);
    count = info.count; dir = info.dir;
    sheet = _vrPrepararHoja();
    idx = 0;
    props.setProperty(VR_CONFIG.COUNT_KEY, String(count));
    props.setProperty(VR_CONFIG.DIR_KEY, dir);
    props.setProperty(VR_CONFIG.IDX_KEY, '0');
    Logger.log('▶️ Inicio. Total documentos: ' + count + ' · orden API: ' + dir);
  } else {
    count = parseInt(props.getProperty(VR_CONFIG.COUNT_KEY) || '0');
    dir   = props.getProperty(VR_CONFIG.DIR_KEY) || 'asc';
    sheet = ss.getSheetByName(VR_CONFIG.SHEET) || _vrPrepararHoja();
  }

  var totalPages = Math.ceil(count / VR_CONFIG.LIMIT);
  var inicio = Date.now();
  var fallidas = 0;
  ss.toast('Descargando (nuevo→viejo) ' + idx + '/' + totalPages + ' páginas…', 'Ventas', -1);

  while (idx < totalPages) {
    // ¿Hoja llena? Terminamos con lo que cabe.
    if (sheet.getLastRow() >= VR_CONFIG.MAX_FILAS) {
      _vrFinalizar(props, sheet, '🟡 Hoja llena (tope de seguridad). Quedó la venta más reciente.');
      return;
    }
    // ¿Se acabó el tiempo del tramo?
    if (Date.now() - inicio > VR_CONFIG.MAX_MS) {
      props.setProperty(VR_CONFIG.IDX_KEY, String(idx));
      var pausa = '⏸️ Pausa en página ' + idx + ' / ' + totalPages + '. Reejecuta descargarVentasRapido().';
      Logger.log(pausa); ss.toast(pausa, 'Ventas', 12);
      return;
    }

    // Armar la ronda de páginas (en orden nuevo→viejo) y pedirlas en paralelo.
    var reqs = [];
    for (var b = 0; b < VR_CONFIG.BATCH_PAGES && (idx + b) < totalPages; b++) {
      var off = _vrOffset(idx + b, totalPages, dir);
      reqs.push({
        url: 'https://api.bsale.io/v1/documents.json?expand=%5Bdetails%5D&limit='
             + VR_CONFIG.LIMIT + '&offset=' + off,
        method: 'get', headers: { access_token: token }, muteHttpExceptions: true
      });
    }

    var resps = UrlFetchApp.fetchAll(reqs);
    var filas = [];
    var algunaEnRango = false;
    var maxIdRonda = 0;

    resps.forEach(function(res) {
      if (res.getResponseCode() !== 200) { fallidas++; return; }
      var data = JSON.parse(res.getContentText());
      (data.items || []).forEach(function(doc) {
        var did = Number(doc.id || 0); if (did > maxIdRonda) maxIdRonda = did;  // para incremental
        if (doc.state === 1) return;                                  // anulado
        var tipoId = parseInt((doc.document_type || {}).id || 0);
        if (VR_CONFIG.TIPOS_VENTA.concat(VR_CONFIG.TIPOS_DEVOLUCION).indexOf(tipoId) === -1) return;
        var emision = doc.emissionDate || 0;
        if (emision < VR_CONFIG.FECHA_DESDE) return;                  // anterior a 2021
        algunaEnRango = true;

        var esDev = VR_CONFIG.TIPOS_DEVOLUCION.indexOf(tipoId) !== -1;
        var fecha = new Date(emision * 1000).toLocaleDateString('es-CL');
        var dets  = ((doc.details || {}).items) || [];
        dets.forEach(function(det) {
          var v = det.variant || {};
          var cant = parseFloat(det.quantity || 0);
          if (esDev) cant = -Math.abs(cant);
          filas.push([
            fecha, doc.number || '',
            VR_TIPO_NOMBRES[tipoId] || ('Tipo ' + tipoId),
            esDev ? 'Devolución' : 'Venta',
            v.code || '', cant,
            det.netUnitValue || 0, det.netAmount || 0
          ]);
        });
      });
    });

    if (filas.length) {
      try {
        sheet.getRange(sheet.getLastRow() + 1, 1, filas.length, VR_HEADERS.length).setValues(filas);
      } catch (e) {
        // Backstop por si el tope de celdas llega antes que MAX_FILAS.
        _vrFinalizar(props, sheet, '🟡 Tope de celdas alcanzado al escribir. Quedó la venta más reciente.');
        return;
      }
    }

    idx += reqs.length;
    props.setProperty(VR_CONFIG.IDX_KEY, String(idx));
    _vrBumpMaxId(props, maxIdRonda);   // recuerda el doc más nuevo visto

    // Yendo de nuevo→viejo: si una ronda completa ya no trae nada en rango,
    // lo que sigue es aún más viejo → terminamos.
    if (!algunaEnRango && filas.length === 0) {
      _vrFinalizar(props, sheet, '✅ Llegamos al corte de fecha (2021).');
      return;
    }
    Utilities.sleep(100);
  }

  _vrFinalizar(props, sheet, '✅ Descarga completa (recorrido todos los documentos).');
  if (fallidas) Logger.log('⚠️ ' + fallidas + ' páginas fallaron (reintenta para rellenar huecos menores).');
}

// ============================================================
// ACTUALIZACIÓN INCREMENTAL — trae solo ventas NUEVAS
// Pensada para correr a diario (trigger). No re-escanea todo:
// recorre de lo más nuevo hacia atrás hasta toparse con un
// documento que ya teníamos (por id), inserta lo nuevo arriba y,
// si la hoja supera el tope, suelta las filas más antiguas
// (ventana móvil). Atómica: si la corrida se corta antes de
// alcanzar el borde, no escribe nada y reintenta limpio.
// ============================================================
function actualizarVentasRapido() {
  var token = PropertiesService.getScriptProperties().getProperty(VR_CONFIG.TOKEN_PROPERTY);
  if (!token) throw new Error('❌ No existe BSALE_TOKEN.');

  var props = PropertiesService.getScriptProperties();
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(VR_CONFIG.SHEET);
  if (!sheet || sheet.getLastRow() < 2) {
    throw new Error('❌ Aún no hay histórico. Corre primero descargarVentasRapido() hasta el final.');
  }
  var maxId = parseInt(props.getProperty(VR_CONFIG.MAXID_KEY) || '0');
  if (!maxId) {
    throw new Error('❌ Falta el marcador del último doc (VR_MAX_ID). Termina el backfill con descargarVentasRapido().');
  }

  var info = _vrInicializar(token);          // total + orden (solo lectura)
  var totalPages = Math.ceil(info.count / VR_CONFIG.LIMIT);
  var tipos = VR_CONFIG.TIPOS_VENTA.concat(VR_CONFIG.TIPOS_DEVOLUCION);

  var nuevas = [];
  var nuevoMax = maxId;
  var k = 0, alcanzado = false;

  while (k < totalPages && !alcanzado) {
    var reqs = [];
    for (var b = 0; b < VR_CONFIG.BATCH_PAGES && (k + b) < totalPages; b++) {
      var off = _vrOffset(k + b, totalPages, info.dir);
      reqs.push({
        url: 'https://api.bsale.io/v1/documents.json?expand=%5Bdetails%5D&limit='
             + VR_CONFIG.LIMIT + '&offset=' + off,
        method: 'get', headers: { access_token: token }, muteHttpExceptions: true
      });
    }
    var resps = UrlFetchApp.fetchAll(reqs);
    var nuevoEnRonda = false;

    resps.forEach(function(res) {
      if (res.getResponseCode() !== 200) return;
      var data = JSON.parse(res.getContentText());
      (data.items || []).forEach(function(doc) {
        var did = Number(doc.id || 0); if (did > nuevoMax) nuevoMax = did;
        if (did <= maxId) return;                       // ya lo teníamos
        nuevoEnRonda = true;
        if (doc.state === 1) return;
        var tipoId = parseInt((doc.document_type || {}).id || 0);
        if (tipos.indexOf(tipoId) === -1) return;
        var emision = doc.emissionDate || 0;
        if (emision < VR_CONFIG.FECHA_DESDE) return;

        var esDev = VR_CONFIG.TIPOS_DEVOLUCION.indexOf(tipoId) !== -1;
        var fecha = new Date(emision * 1000).toLocaleDateString('es-CL');
        (((doc.details || {}).items) || []).forEach(function(det) {
          var v = det.variant || {};
          var cant = parseFloat(det.quantity || 0);
          if (esDev) cant = -Math.abs(cant);
          nuevas.push([
            fecha, doc.number || '', VR_TIPO_NOMBRES[tipoId] || ('Tipo ' + tipoId),
            esDev ? 'Devolución' : 'Venta', v.code || '', cant,
            det.netUnitValue || 0, det.netAmount || 0
          ]);
        });
      });
    });

    k += reqs.length;
    if (!nuevoEnRonda) alcanzado = true;     // ronda completa ya conocida → al día
  }

  // Inserción atómica al final: arriba van las nuevas (más recientes primero).
  if (nuevas.length) {
    sheet.insertRowsAfter(1, nuevas.length);
    sheet.getRange(2, 1, nuevas.length, VR_HEADERS.length).setValues(nuevas);
    var exceso = (sheet.getLastRow() - 1) - VR_CONFIG.MAX_FILAS;   // ventana móvil
    if (exceso > 0) sheet.deleteRows(sheet.getLastRow() - exceso + 1, exceso);
  }
  props.setProperty(VR_CONFIG.MAXID_KEY, String(nuevoMax));

  var msg = '✅ Actualización: ' + nuevas.length + ' líneas nuevas'
          + (nuevas.length ? ' (insertadas arriba).' : ' — ya estaba al día.');
  Logger.log(msg); ss.toast(msg, 'Ventas', 10);
}

function instalarTriggerActualizacion() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'actualizarVentasRapido') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('actualizarVentasRapido').timeBased().everyDays(1).atHour(7).create();
  Logger.log('✅ Trigger diario (07:00) instalado para actualizarVentasRapido.');
}

// ------------------------------------------------------------
// RECUPERACIÓN SIN MARCADOR — para cuando se perdieron las Script
// Properties (VR_MAX_ID/VR_DONE), p. ej. al migrar de proyecto.
// Usa la ÚLTIMA FECHA presente en la hoja como corte: borra ese día
// (por si quedó incompleto) y re-descarga todo desde esa fecha hasta
// hoy. Al terminar siembra VR_MAX_ID y VR_DONE para que la
// actualización diaria vuelva a funcionar sola.
// Idempotente: se puede correr de nuevo sin duplicar.
// ------------------------------------------------------------
function recuperarVentasDesdeUltimaFecha() {
  var token = PropertiesService.getScriptProperties().getProperty(VR_CONFIG.TOKEN_PROPERTY);
  if (!token) throw new Error('❌ No existe BSALE_TOKEN.');
  var props = PropertiesService.getScriptProperties();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(VR_CONFIG.SHEET);
  if (!sheet || sheet.getLastRow() < 2) {
    throw new Error('❌ No hay histórico. Para descarga desde cero usa descargarVentasRapido().');
  }

  // 1) Fecha de corte = la más nueva de la hoja (fila 2, hoja ordenada nueva→vieja).
  var corteMs = _vrEpochCelda(sheet.getRange(2, 1).getValue());
  if (!corteMs) throw new Error('❌ No pude leer la fecha de la fila 2.');
  var corteSeg = Math.floor(corteMs / 1000);   // emissionDate de Bsale es epoch en segundos
  ss.toast('Recuperando ventas desde ' + new Date(corteMs).toLocaleDateString('es-CL') + '…', 'Ventas', -1);

  // 2) Borrar las filas de esa fecha (bloque contiguo arriba) para no duplicar el día de corte.
  var nBorrar = 0;
  var maxScan = Math.min(sheet.getLastRow() - 1, 5000);
  var fechas = sheet.getRange(2, 1, maxScan, 1).getValues();
  for (var i = 0; i < fechas.length; i++) {
    if (_vrEpochCelda(fechas[i][0]) === corteMs) nBorrar++; else break;
  }
  if (nBorrar > 0) sheet.deleteRows(2, nBorrar);

  // 3) Recorrer la API de lo más nuevo a lo más viejo hasta pasar el corte.
  var info = _vrInicializar(token);
  var totalPages = Math.ceil(info.count / VR_CONFIG.LIMIT);
  var tipos = VR_CONFIG.TIPOS_VENTA.concat(VR_CONFIG.TIPOS_DEVOLUCION);
  var nuevas = [], maxId = 0, k = 0, pasado = false;

  while (k < totalPages && !pasado) {
    var reqs = [];
    for (var b = 0; b < VR_CONFIG.BATCH_PAGES && (k + b) < totalPages; b++) {
      reqs.push({
        url: 'https://api.bsale.io/v1/documents.json?expand=%5Bdetails%5D&limit='
             + VR_CONFIG.LIMIT + '&offset=' + _vrOffset(k + b, totalPages, info.dir),
        method: 'get', headers: { access_token: token }, muteHttpExceptions: true
      });
    }
    var resps = UrlFetchApp.fetchAll(reqs);
    var todosViejos = true;

    resps.forEach(function(res) {
      if (res.getResponseCode() !== 200) return;
      (JSON.parse(res.getContentText()).items || []).forEach(function(doc) {
        var did = Number(doc.id || 0); if (did > maxId) maxId = did;
        var emision = doc.emissionDate || 0;
        if (emision >= corteSeg) todosViejos = false; else return;   // anterior al corte → fuera
        if (doc.state === 1) return;
        var tipoId = parseInt((doc.document_type || {}).id || 0);
        if (tipos.indexOf(tipoId) === -1) return;
        var esDev = VR_CONFIG.TIPOS_DEVOLUCION.indexOf(tipoId) !== -1;
        var fecha = new Date(emision * 1000).toLocaleDateString('es-CL');
        (((doc.details || {}).items) || []).forEach(function(det) {
          var v = det.variant || {};
          var cant = parseFloat(det.quantity || 0);
          if (esDev) cant = -Math.abs(cant);
          nuevas.push([fecha, doc.number || '', VR_TIPO_NOMBRES[tipoId] || ('Tipo ' + tipoId),
                       esDev ? 'Devolución' : 'Venta', v.code || '', cant,
                       det.netUnitValue || 0, det.netAmount || 0]);
        });
      });
    });

    k += reqs.length;
    if (todosViejos) pasado = true;   // ronda completa anterior al corte → listo
  }

  // 4) Insertar arriba y sembrar los marcadores.
  if (nuevas.length) {
    sheet.insertRowsAfter(1, nuevas.length);
    sheet.getRange(2, 1, nuevas.length, VR_HEADERS.length).setValues(nuevas);
  }
  props.setProperty(VR_CONFIG.MAXID_KEY, String(maxId));
  props.setProperty(VR_CONFIG.DONE_KEY, '1');

  var msg = '✅ Recuperación: ' + nBorrar + ' filas del día de corte reemplazadas, '
          + nuevas.length + ' líneas insertadas. VR_MAX_ID sembrado (' + maxId + ').';
  Logger.log(msg); ss.toast(msg, 'Ventas', 15);
}

// Epoch ms desde celda de fecha (Date o texto dd-mm-yyyy).
function _vrEpochCelda(v) {
  if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate()).getTime();
  var s = String(v || '').trim();
  var p = s.split(/[\/\-.]/);
  if (p.length < 3) return 0;
  var d, m, y;
  if (p[0].length === 4) { y = +p[0]; m = +p[1]; d = +p[2]; }
  else { d = +p[0]; m = +p[1]; y = +p[2]; }
  if (!y || !m || !d) return 0;
  return new Date(y, m - 1, d).getTime();
}

// Reinicia el progreso (la próxima corrida limpia la hoja y arranca de cero).
function reiniciarVentasRapido() {
  var p = PropertiesService.getScriptProperties();
  [VR_CONFIG.COUNT_KEY, VR_CONFIG.DIR_KEY, VR_CONFIG.IDX_KEY, VR_CONFIG.DONE_KEY].forEach(function(k) { p.deleteProperty(k); });
  Logger.log('🗑️ Progreso reiniciado (la próxima corrida vacía la hoja y descarga de nuevo).');
}

// ------------------------------------------------------------
// HELPERS
// ------------------------------------------------------------
// Guarda el id más alto visto (marcador para la actualización incremental).
function _vrBumpMaxId(props, id) {
  if (!id) return;
  var actual = parseInt(props.getProperty(VR_CONFIG.MAXID_KEY) || '0');
  if (id > actual) props.setProperty(VR_CONFIG.MAXID_KEY, String(id));
}

// Offset de la k-ésima página en orden nuevo→viejo.
function _vrOffset(k, totalPages, dir) {
  // dir 'asc'  = la API entrega del más viejo al más nuevo → lo nuevo está al final
  // dir 'desc' = lo nuevo está al principio
  var page = (dir === 'asc') ? (totalPages - 1 - k) : k;
  return page * VR_CONFIG.LIMIT;
}

// Obtiene total de documentos y detecta el orden de la API (asc/desc).
function _vrInicializar(token) {
  var base = 'https://api.bsale.io/v1/documents.json?limit=' + VR_CONFIG.LIMIT + '&offset=';
  var prim = _vrFetch(base + '0', token);
  var count = Number(prim.count || 0);
  if (count === 0) throw new Error('❌ La API no devolvió documentos.');

  var totalPages = Math.ceil(count / VR_CONFIG.LIMIT);
  var emisionPrimera = ((prim.items || [])[0] || {}).emissionDate || 0;

  var ultOffset = (totalPages - 1) * VR_CONFIG.LIMIT;
  var ult = _vrFetch(base + ultOffset, token);
  var itemsUlt = ult.items || [];
  var emisionUltima = (itemsUlt[itemsUlt.length - 1] || {}).emissionDate || 0;

  // Si el final es más nuevo que el principio → la API ordena ascendente.
  var dir = (emisionUltima >= emisionPrimera) ? 'asc' : 'desc';
  return { count: count, dir: dir };
}

function _vrPrepararHoja() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(VR_CONFIG.SHEET);
  if (!sheet) sheet = ss.insertSheet(VR_CONFIG.SHEET);
  sheet.clearContents();
  sheet.appendRow(VR_HEADERS);
  sheet.getRange(1, 1, 1, VR_HEADERS.length)
    .setFontWeight('bold').setBackground('#F3F3F3').setFontColor('#000000');
  sheet.setFrozenRows(1);
  return sheet;
}

function _vrFinalizar(props, sheet, motivo) {
  [VR_CONFIG.COUNT_KEY, VR_CONFIG.DIR_KEY, VR_CONFIG.IDX_KEY].forEach(function(k) { props.deleteProperty(k); });
  props.setProperty(VR_CONFIG.DONE_KEY, '1');   // marca terminado: el trigger ya no lo reinicia
  var filas = Math.max(0, sheet.getLastRow() - 1);
  var nueva = filas ? sheet.getRange(2, 1, 1, 1).getValue() : '—';
  var vieja = filas ? sheet.getRange(sheet.getLastRow(), 1, 1, 1).getValue() : '—';
  var msg = motivo + ' Filas: ' + filas + ' · más nueva: ' + nueva + ' · más antigua: ' + vieja;
  Logger.log(msg);
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'Ventas', 15);
}

function _vrFetch(url, token) {
  var res = UrlFetchApp.fetch(url, { method: 'get', headers: { access_token: token }, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) {
    throw new Error('❌ Error API Bsale HTTP ' + res.getResponseCode() + ': ' + res.getContentText());
  }
  return JSON.parse(res.getContentText());
}
