// ============================================================
// ESTACIONALIDAD por categoría — matriz Año × Mes
// Lee "Ventas Históricas" y escribe la pestaña "Estacionalidad":
// para cada SKU de una CATEGORÍA de "Config SKU", una tabla con las
// unidades vendidas en cada mes de cada año, más el total y el peso de
// la ventana jun–oct (temporada del "18" en Chile). Sirve para decidir
// compra mirando el patrón real, no la intuición.
//
// OJO (riesgo §9 del brief): un mes en cero puede ser "no hubo demanda"
// O "no hubo stock" (quiebre). Los ceros dentro de un patrón que
// normalmente vende son sospechosos de quiebre. Se marcan en rojo suave.
//
// Uso: analizarEstacionalidad()            → categoría por defecto (NEGOCIO.EST_CATEGORIA)
//      analizarEstacionalidad('Bolsas …')  → categoría específica de "Config SKU"
// ============================================================

const EST_CONFIG = {
  SHEET_VENTAS:   'Ventas Históricas',
  SHEET_CATALOGO: 'Config SKU',
  SHEET_SALIDA:   'Estacionalidad',
  // Ventas Históricas (1-based): Fecha=1, SKU=5, Cantidad=6
  COL_FECHA: 1, COL_SKU: 5, COL_CANT: 6,
  // Config SKU (1-based): SKU=col A(1), Origen=col F(6). Las filas de
  // encabezado de categoría tienen texto en A y la col Origen vacía.
  CAT_SKU: 1, CAT_ORIGEN: 6,
  // Ventana de temporada a resaltar (meses 1-12). Jun–Oct = 6..10.
  TEMP_INI: 6, TEMP_FIN: 10
};

const EST_MESES = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];

function analizarEstacionalidad(categoria) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  categoria = String(categoria || NEGOCIO.EST_CATEGORIA || '').trim();
  if (!categoria) throw new Error('❌ No hay categoría (define NEGOCIO.EST_CATEGORIA o pásala como argumento).');

  var ventas = ss.getSheetByName(EST_CONFIG.SHEET_VENTAS);
  if (!ventas || ventas.getLastRow() < 2) {
    throw new Error('❌ No hay datos en "' + EST_CONFIG.SHEET_VENTAS + '".');
  }
  ss.toast('Analizando estacionalidad · ' + categoria + '…', 'Estacionalidad', -1);

  // SKUs de la categoría (desde "Config SKU"), normalizados a MAYÚSCULAS.
  var skus = _estSkusDeCategoria(categoria);
  if (!skus.length) {
    throw new Error('❌ La categoría "' + categoria + '" no tiene SKUs en "' +
      EST_CONFIG.SHEET_CATALOGO + '". Categorías detectadas: ' +
      _estCategorias().join(' · '));
  }
  var objetivo = {};
  skus.forEach(function(s) { objetivo[s] = true; });

  var n = ventas.getLastRow() - 1;
  var datos = ventas.getRange(2, 1, n, EST_CONFIG.COL_CANT).getValues();

  // agg[sku][año][mes 0-11] = unidades
  var agg = {};
  var anios = {};
  datos.forEach(function(r) {
    var sku = String(r[EST_CONFIG.COL_SKU - 1] || '').trim().toUpperCase();
    if (!objetivo[sku]) return;
    var ep = _estEpoch(r[EST_CONFIG.COL_FECHA - 1]);
    if (!ep) return;
    var cant = parseFloat(r[EST_CONFIG.COL_CANT - 1] || 0);
    var d = new Date(ep);
    var y = d.getFullYear(), m = d.getMonth();
    if (!agg[sku]) agg[sku] = {};
    if (!agg[sku][y]) agg[sku][y] = new Array(12).fill(0);
    agg[sku][y][m] += cant;
    anios[y] = true;
  });

  var listaAnios = Object.keys(anios).map(Number).sort(function(a, b) { return a - b; });

  var sheet = _estPrepararHoja();

  // Panel-resumen arriba: unidades jun–oct por año, cada SKU + familia.
  var fila = _estResumen(sheet, 1, categoria, skus, agg, listaAnios);
  fila += 1; // separación antes del detalle

  // Un bloque por SKU (en el orden del catálogo) + bloque "FAMILIA" con la suma.
  var familia = {};   // año → array 12
  skus.forEach(function(sku) {
    fila = _estBloque(sheet, fila, sku, agg[sku] || {}, listaAnios);
    if (agg[sku]) listaAnios.forEach(function(y) {
      if (!agg[sku][y]) return;
      if (!familia[y]) familia[y] = new Array(12).fill(0);
      for (var m = 0; m < 12; m++) familia[y][m] += agg[sku][y][m];
    });
    fila += 1; // línea en blanco entre bloques
  });
  _estBloque(sheet, fila, 'FAMILIA (suma de ' + skus.length + ')', familia, listaAnios);

  sheet.autoResizeColumns(1, 16);
  var msg = '✅ Estacionalidad · ' + categoria + ': ' + skus.length + ' SKUs · '
          + (listaAnios[0] || '?') + '–' + (listaAnios[listaAnios.length - 1] || '?');
  Logger.log(msg); ss.toast(msg, 'Estacionalidad', 12);
}

// ------------------------------------------------------------
// SELECCIÓN DE SKUs POR CATEGORÍA (Config SKU)
// ------------------------------------------------------------
// El catálogo está organizado por bloques: una fila de encabezado con el
// nombre de la categoría (texto en col A, col Origen vacía) y debajo sus
// SKUs (col A + col Origen con país). Devuelve los SKUs (MAYÚS) de la
// categoría pedida, en el orden del catálogo.
function _estSkusDeCategoria(categoria) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EST_CONFIG.SHEET_CATALOGO);
  if (!sh || sh.getLastRow() < 2) return [];
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, EST_CONFIG.CAT_ORIGEN).getValues();
  var objetivo = categoria.toUpperCase();
  var out = [];
  var catActual = '';
  v.forEach(function(r) {
    var a    = String(r[EST_CONFIG.CAT_SKU - 1] || '').trim();
    var pais = String(r[EST_CONFIG.CAT_ORIGEN - 1] || '').trim();
    if (!a) return;
    if (!pais) { catActual = a.toUpperCase(); return; }    // fila de encabezado de categoría
    if (catActual === objetivo) out.push(a.toUpperCase()); // SKU dentro de la categoría
  });
  return out;
}

// Lista de categorías detectadas en el catálogo (para mensajes de ayuda).
function _estCategorias() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EST_CONFIG.SHEET_CATALOGO);
  if (!sh || sh.getLastRow() < 2) return [];
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, EST_CONFIG.CAT_ORIGEN).getValues();
  var out = [];
  v.forEach(function(r) {
    var a    = String(r[EST_CONFIG.CAT_SKU - 1] || '').trim();
    var pais = String(r[EST_CONFIG.CAT_ORIGEN - 1] || '').trim();
    if (a && !pais) out.push(a);   // encabezado de categoría
  });
  return out;
}

// ------------------------------------------------------------
// ESCRITURA DEL REPORTE
// ------------------------------------------------------------
// Panel-resumen: título + nota + matriz Año × (SKUs + Familia) con
// unidades vendidas en jun–oct. Devuelve la siguiente fila libre.
function _estResumen(sheet, fila, categoria, skus, agg, listaAnios) {
  var anioActual = new Date().getFullYear();

  // Título
  sheet.getRange(fila, 1).setValue('ESTACIONALIDAD ' + categoria.toUpperCase() +
      ' — unidades vendidas Jun–Oct por año')
    .setFontWeight('bold').setFontSize(13).setFontColor('#000000');
  fila++;
  // Nota de lectura (neutral, sin conclusiones prefijadas)
  sheet.getRange(fila, 1).setValue('Ventana Jun–Oct resaltada (temporada del 18). '
      + 'Un cero dentro de la temporada puede ser quiebre (sin stock), no falta de demanda.')
    .setFontStyle('italic').setFontColor('#666666');
  fila += 2;

  // Cabecera
  var cab = ['Año'].concat(skus).concat(['FAMILIA']);
  sheet.getRange(fila, 1, 1, cab.length).setValues([cab])
    .setFontWeight('bold').setBackground('#F3F3F3').setFontColor('#000000')
    .setHorizontalAlignment('center');
  fila++;

  var ini = fila;
  listaAnios.forEach(function(y) {
    var etiqueta = (y === anioActual) ? (y + ' (parcial)') : String(y);
    var fam = 0;
    var vals = skus.map(function(sku) {
      var arr = (agg[sku] && agg[sku][y]) ? agg[sku][y] : null;
      var temp = 0;
      if (arr) for (var m = EST_CONFIG.TEMP_INI - 1; m <= EST_CONFIG.TEMP_FIN - 1; m++) temp += arr[m];
      fam += temp;
      return temp;
    });
    var out = [etiqueta].concat(vals).concat([fam]);
    sheet.getRange(fila, 1, 1, out.length).setValues([out]);
    if (y === anioActual) {
      sheet.getRange(fila, 1, 1, out.length).setFontColor('#999999');   // año incompleto, atenuado
    }
    fila++;
  });

  // Formato: números con miles, familia en negrita, centrado
  sheet.getRange(ini, 2, fila - ini, skus.length + 1).setNumberFormat('#,##0')
    .setHorizontalAlignment('center');
  sheet.getRange(ini, skus.length + 2, fila - ini, 1).setFontWeight('bold').setBackground('#E8E8E8');
  sheet.getRange(ini, 1, fila - ini, 1).setHorizontalAlignment('center');

  return fila;
}

// Escribe un bloque (título + matriz Año×Mes + Total + Jun–Oct + %) y
// devuelve la siguiente fila libre.
function _estBloque(sheet, fila, titulo, porAnio, listaAnios) {
  // Título del SKU
  sheet.getRange(fila, 1).setValue(titulo)
    .setFontWeight('bold').setFontColor('#000000').setFontSize(11);
  fila++;

  // Cabecera: Año | Ene..Dic | Total | Jun-Oct | %
  var cab = ['Año'].concat(EST_MESES).concat(['Total', 'Jun–Oct', '% Jun–Oct']);
  sheet.getRange(fila, 1, 1, cab.length).setValues([cab])
    .setFontWeight('bold').setBackground('#EEEEEE');
  fila++;

  var iniFila = fila;
  var promTemp = 0, nAniosConVenta = 0;
  listaAnios.forEach(function(y) {
    var arr = porAnio[y] || new Array(12).fill(0);
    var total = 0, temp = 0;
    for (var m = 0; m < 12; m++) {
      total += arr[m];
      if ((m + 1) >= EST_CONFIG.TEMP_INI && (m + 1) <= EST_CONFIG.TEMP_FIN) temp += arr[m];
    }
    var pct = total > 0 ? Math.round(temp / total * 100) : 0;
    var out = [y].concat(arr).concat([total, temp, pct + '%']);
    sheet.getRange(fila, 1, 1, out.length).setValues([out]);

    // Resaltar ceros dentro de la ventana de temporada (posible quiebre).
    for (var m = 0; m < 12; m++) {
      if ((m + 1) >= EST_CONFIG.TEMP_INI && (m + 1) <= EST_CONFIG.TEMP_FIN && arr[m] === 0) {
        sheet.getRange(fila, 2 + m).setBackground('#F4CCCC');   // rojo suave = sospecha de quiebre
      }
    }
    // Resaltar la columna Jun–Oct
    sheet.getRange(fila, 14).setBackground('#E8E8E8');
    if (total > 0) { promTemp += temp; nAniosConVenta++; }
    fila++;
  });

  // Formato numérico de la matriz
  if (fila > iniFila) sheet.getRange(iniFila, 2, fila - iniFila, 14).setNumberFormat('#,##0');

  // Línea resumen: promedio de unidades jun–oct por año con venta.
  var prom = nAniosConVenta ? Math.round(promTemp / nAniosConVenta) : 0;
  sheet.getRange(fila, 1).setValue('Prom. Jun–Oct/año: ' + prom.toLocaleString('es-CL') + ' uds')
    .setFontStyle('italic').setFontColor('#666666');
  fila++;

  return fila;
}

// ------------------------------------------------------------
// HELPERS
// ------------------------------------------------------------
function _estPrepararHoja() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(EST_CONFIG.SHEET_SALIDA);
  if (!sheet) sheet = ss.insertSheet(EST_CONFIG.SHEET_SALIDA);
  sheet.clear();
  return sheet;
}

// Epoch (ms) desde una celda de fecha (Date o texto dd-mm-yyyy / yyyy-mm-dd).
function _estEpoch(v) {
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
