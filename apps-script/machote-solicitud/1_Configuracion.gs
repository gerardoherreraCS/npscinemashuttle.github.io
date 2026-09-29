/**
 * MACHOTE MAESTRO — SOLICITUD DE TRANSPORTE (CINEMA SHUTTLE)
 * ------------------------------------------------------------
 * Vive 100% en Google Sheets + Apps Script. Diseñado para:
 *  1) Duplicarse por proyecto (Archivo > Hacer una copia) sin perder
 *     la configuración (menús, fórmulas, validaciones, formato).
 *  2) Permitir que producción/cliente use cualquiera de las 200 filas
 *     ya preparadas sin romper nada (no hace falta insertar filas a mano).
 *  3) Alimentar sin fricción a la BASE DE DATOS MAESTRA (Tablero Analítico):
 *     la pestaña se llama "Solicitudes Transporte" (igual que en los
 *     proyectos anteriores) y usa los mismos encabezados que el consolidador
 *     sabe leer (VEHICULO, PROVEEDOR, NOMBRE CHOFER, PLACAS DE LA UNIDAD...).
 *
 * Archivos del proyecto de Apps Script:
 *   - 1_Configuracion.gs, 2_Catalogo_y_Cascada.gs, 3_Hoja_Solicitud.gs, 4_Formulario_Servidor.gs
 *   - Formulario.html (formulario de captura: panel lateral y web app)
 *
 * Este script NUNCA escribe en el catálogo maestro (solo lo lee).
 */

// ====================== CONFIGURACIÓN ======================
const CATALOG_ID = '1BGtrUrcbjCZzFU3xUInleVbesmmgVFJY4rixDYAWkF4';
const TAB_UNIDADES = 'CatalogoTranspo';
const TAB_DRIVERS = 'CATALOGO DRIVERS';

// Nombre OBLIGATORIO de la pestaña de captura. El consolidador busca
// exactamente esta pestaña en cada archivo de la carpeta de proyectos.
const HOJA_SOLICITUD = 'Solicitudes Transporte';
// Nombres anteriores: si el archivo todavía tiene alguno, se renombra solo.
const NOMBRES_ANTERIORES = ['Solicitud'];

const N_FILAS = 200;   // filas de captura ya configuradas
const N_DIAS = 31;     // días del calendario, generado desde "Fecha de Inicio"
const MAX_UNIDADES_POR_ENVIO = 20; // tope para el formulario (evita envíos por error)

// columnas de la hoja (fijas; si cambias esto, todo el resto se ajusta solo)
const COL = {
  CONCEPTO: 1, PAX_SOLICITADO: 2, BILINGUE: 3,
  VEHICULO: 4, PROVEEDOR: 5, PLACAS: 6, ANIO: 7, PAX: 8,
  CHOFER: 9, CEL: 10, VALID_UNIDAD: 11, VALID_DRIVER: 12,
  DIA_START: 13, get DIA_END() { return this.DIA_START + N_DIAS - 1; },
};
COL.D_FULL = COL.DIA_END + 1;
COL.D_FORANEO = COL.DIA_END + 2;
COL.D_TRANSFER = COL.DIA_END + 3;
COL.D_CANCELADO = COL.DIA_END + 4;
COL.T_DIA = COL.DIA_END + 5;
COL.T_FORANEA = COL.DIA_END + 6;
COL.T_TRANSFER = COL.DIA_END + 7;
COL.HORAS_EXTRA = COL.DIA_END + 8;
COL.T_HORA = COL.DIA_END + 9;
COL.VIATICOS = COL.DIA_END + 10;
COL.TOTAL = COL.DIA_END + 11;
COL.INCIDENCIA = COL.DIA_END + 12;
COL.OBS = COL.DIA_END + 13;
const LAST_COL = COL.OBS;

const GROUP_ROW = 7, HDR_ROW = 8, DATA_START = 9;
const DATA_END = DATA_START + N_FILAS - 1;

// Celdas del encabezado del proyecto (fila, columna)
const CELDA = {
  PROYECTO: [3, 2], CLIENTE: [3, 7],
  PRODUCTOR: [4, 2], CAPITAN: [4, 7],
  ESPECIALISTA: [5, 2], FECHA_INICIO: [5, 7],
};

// Correos de las personas que SÍ pueden tocar fórmulas, encabezados y estructura.
const EDITORES_EQUIPO = [
  // 'xochitl@cinemashuttle.mx',
  // 'samanta@cinemashuttle.mx',
];

// ====================== UTILIDADES ======================
// Devuelve la hoja de captura. Si el archivo aún tiene el nombre viejo
// ("Solicitud"), la renombra a "Solicitudes Transporte" para que el
// consolidador la encuentre.
function hojaSolicitud(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(HOJA_SOLICITUD);
  if (sh) return sh;
  for (var i = 0; i < NOMBRES_ANTERIORES.length; i++) {
    var vieja = ss.getSheetByName(NOMBRES_ANTERIORES[i]);
    if (vieja) { vieja.setName(HOJA_SOLICITUD); return vieja; }
  }
  return null;
}

// Alertas que no truenan cuando el código corre desde un trigger o la web app
function avisar(msg) {
  try { SpreadsheetApp.getUi().alert(msg); }
  catch (e) { Logger.log(msg); }
}

function celdaFechaInicio(sol) {
  return sol.getRange(CELDA.FECHA_INICIO[0], CELDA.FECHA_INICIO[1]);
}

// ====================== MENÚ ======================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Machote Transporte')
    .addItem('Abrir formulario de captura (panel lateral)', 'abrirFormularioLateral')
    .addSeparator()
    .addItem('1) Actualizar catálogo (proveedores / placas / choferes)', 'actualizarCatalogo')
    .addItem('2) Construir / reparar hoja "' + HOJA_SOLICITUD + '"', 'construirSolicitud')
    .addItem('3) Preparar para nuevo proyecto (limpiar datos, conservar configuración)', 'limpiarParaNuevoProyecto')
    .addSeparator()
    .addItem('4) Proteger hoja (solo el equipo puede editar fórmulas/estructura)', 'protegerHoja')
    .addItem('Agregarme como editor del equipo', 'agregarmeComoEditor')
    .addItem('Quitar protección (temporal, para reparar)', 'quitarProteccion')
    .addSeparator()
    .addItem('Configuración inicial completa (1 + 2)', 'configuracionInicial')
    .addItem('Diagnóstico: ¿está cargado el catálogo?', 'verEstadoCatalogo')
    .addSeparator()
    .addItem('5) Publicar formulario para clientes (ver instrucciones)', 'instruccionesFormulario')
    .addToUi();
}

function abrirFormularioLateral() {
  var html = HtmlService.createHtmlOutputFromFile('Formulario').setTitle('Solicitud de transporte');
  SpreadsheetApp.getUi().showSidebar(html);
}

function verEstadoCatalogo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var bu = ss.getSheetByName('BDD Unidades');
  var bd = ss.getSheetByName('BDD Choferes');
  var catRange = ss.getRangeByName('CAT_LIST');
  if (!bu || !bd || !catRange || bu.getLastRow() < 2) {
    avisar('ERROR: El catálogo NO está cargado todavía.\n\nCorre "1) Actualizar catálogo" (o "Configuración inicial completa") antes de llenar la solicitud.');
    return;
  }
  var nUnidades = bu.getLastRow() - 1;
  var nChoferes = bd.getLastRow() - 1;
  var nCats = catRange.getValues().filter(function (r) { return r[0]; }).length;
  avisar('OK - Catálogo cargado:\n\n' + nCats + ' categorías de vehículo\n' + nUnidades + ' unidades disponibles\n' + nChoferes + ' choferes\n\nSi elegiste un Vehículo y no ves opciones de Proveedor, revisa que estés escribiendo en una fila entre la ' + DATA_START + ' y la ' + DATA_END + ', y que ese Vehículo sí tenga unidades disponibles (mira la hoja "BDD Unidades").');
}

function configuracionInicial() {
  actualizarCatalogo();
  construirSolicitud();
  avisar('Listo. Catálogo cargado y hoja "' + HOJA_SOLICITUD + '" construida. Ya puedes empezar a llenar.');
}

