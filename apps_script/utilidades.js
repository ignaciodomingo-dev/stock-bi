// ============================================================
// UTILIDADES — mantenimiento puntual
// aplicarEstiloGris(): reformatea las hojas generadas al estilo
// escala de grises (fondo blanco, letra negra, encabezado gris
// claro). Correr UNA VEZ tras el cambio de estilo; las corridas
// futuras ya escriben con este formato. No toca los colores
// funcionales del semáforo (rojo/amarillo/verde/azul por estado),
// que son formato condicional y transmiten información.
// ============================================================

function aplicarEstiloGris() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojas = ['Ventas Históricas', 'Stock Actual', 'Velocidad', 'Semáforo',
               'Ventas Mensuales', 'Stock Histórico', 'Descontinuados', 'Estacionalidad'];
  var hechas = [];

  hojas.forEach(function(nombre) {
    var sh = ss.getSheetByName(nombre);
    if (!sh || sh.getLastRow() < 1) return;

    // Cuerpo: fondo por defecto (blanco), letra negra.
    sh.getDataRange().setBackground(null).setFontColor('#000000');
    // Encabezado: gris claro, negrita.
    sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn()))
      .setFontWeight('bold').setBackground('#F3F3F3').setFontColor('#000000');
    hechas.push(nombre);
  });

  var msg = '✅ Estilo gris aplicado a: ' + hechas.join(', ');
  Logger.log(msg); ss.toast(msg, 'Utilidades', 12);
}
