// ============================================================
// PIPELINE DIARIO — Ventas → Velocidad → Semáforo → Archivo stock
// Reemplaza el trigger diario que solo actualizaba ventas: ahora
// la cadena completa corre sola cada mañana. Cada paso va en
// try/catch para que un fallo (p. ej. Bsale caído) no impida
// recalcular el semáforo con el stock horario ya disponible.
//
// Instalación (una vez): instalarTriggerPipeline()
// (desinstala el trigger antiguo de actualizarVentasRapido).
// ============================================================

function actualizacionDiariaCompleta() {
  var errores = [];

  try {
    actualizarVentasRapido();
  } catch (e) {
    errores.push('Ventas: ' + e.message);
    Logger.log('⚠️ Falló actualización de ventas (sigo con velocidad/semáforo): ' + e.message);
  }

  try {
    calcularVelocidad();
  } catch (e) {
    errores.push('Velocidad: ' + e.message);
    Logger.log('⚠️ Falló velocidad (sigo con semáforo sobre la velocidad anterior): ' + e.message);
  }

  try {
    calcularSemaforo();
  } catch (e) {
    errores.push('Semáforo: ' + e.message);
    Logger.log('❌ Falló semáforo: ' + e.message);
  }

  try {
    calcularVentasMensuales();   // agregado para Looker Studio
  } catch (e) {
    errores.push('Ventas Mensuales: ' + e.message);
    Logger.log('⚠️ Falló agregado mensual (Looker verá la versión anterior): ' + e.message);
  }

  // Post-pipeline: guardar snapshot histórico del stock del día (append).
  try {
    archivarStockDiario();
  } catch (e) {
    errores.push('Archivo stock: ' + e.message);
    Logger.log('⚠️ Falló archivo diario de stock: ' + e.message);
  }

  var msg = errores.length
    ? '⚠️ Pipeline con errores → ' + errores.join(' | ')
    : '✅ Pipeline diario completo: ventas + velocidad + semáforo.';
  Logger.log(msg);
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'Pipeline', 15);
}

// Instala el trigger diario del pipeline (07:00) y elimina el antiguo
// de actualizarVentasRapido para no duplicar la descarga de ventas.
function instalarTriggerPipeline() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var fn = t.getHandlerFunction();
    if (fn === 'actualizarVentasRapido' || fn === 'actualizacionDiariaCompleta') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('actualizacionDiariaCompleta').timeBased().everyDays(1).atHour(7).create();
  Logger.log('✅ Trigger diario (07:00) instalado: actualizacionDiariaCompleta '
           + '(ventas → velocidad → semáforo). Trigger antiguo eliminado.');
}
