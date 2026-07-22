// ============================================================
// PLANTILLA de config_local.js — datos del negocio (SÍ va a git).
//
// Cómo usar:
//   1. Copiar este archivo a config_local.js (mismo directorio).
//   2. Reemplazar los valores de ejemplo por los reales.
//   3. config_local.js está en .gitignore → nunca se commitea, pero clasp
//      lo sube igual a Apps Script (no está en .claspignore).
//
// El resto del código lee de NEGOCIO.* en runtime; si falta config_local.js,
// las funciones fallarán con "NEGOCIO is not defined" → crear el archivo.
// ============================================================

var NEGOCIO = {
  // Rótulo de marca para asuntos de correo y títulos de reportes.
  MARCA: 'MiEmpresa',

  // Oficinas (hoja "Stock Actual"): por nombre.
  OFICINAS_EXCLUIR:     ['BODEGA A EXCLUIR'],   // no es stock real
  OFICINAS_SECUNDARIAS: ['BODEGA SECUNDARIA'],  // no es stock vendible

  // Origen de importación → lead time (días).
  ORIGEN_LEAD: { 'PaisA': 75, 'PaisB': 120 },

  // Prefijo del código de contenedor → país de origen.
  PREFIJO_ORIGEN: {
    'AAA': 'PaisB',
    'BBB': 'PaisA'
  },

  // Prefijos de SKU (prefijo + dígito) → país B.
  SKU_PREFIJOS_INDIA: ['XX', 'YY'],

  // Palabras clave de material en la descripción → país B.
  KEYWORDS_INDIA: ['MATERIAL1', 'MATERIAL2'],

  // Estacionalidad: categoría de "Config SKU" a analizar por defecto.
  EST_CATEGORIA: 'Categoría Ejemplo'
};
