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
 *   - Codigo.gs       (este archivo)
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
    .createMenu('🚐 Machote Transporte')
    .addItem('📝 Abrir formulario de captura (panel lateral)', 'abrirFormularioLateral')
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
    avisar('❌ El catálogo NO está cargado todavía.\n\nCorre "1) Actualizar catálogo" (o "Configuración inicial completa") antes de llenar la solicitud.');
    return;
  }
  var nUnidades = bu.getLastRow() - 1;
  var nChoferes = bd.getLastRow() - 1;
  var nCats = catRange.getValues().filter(function (r) { return r[0]; }).length;
  avisar('✅ Catálogo cargado:\n\n' + nCats + ' categorías de vehículo\n' + nUnidades + ' unidades disponibles\n' + nChoferes + ' choferes\n\nSi elegiste un Vehículo y no ves opciones de Proveedor, revisa que estés escribiendo en una fila entre la ' + DATA_START + ' y la ' + DATA_END + ', y que ese Vehículo sí tenga unidades disponibles (mira la hoja "BDD Unidades").');
}

function configuracionInicial() {
  actualizarCatalogo();
  construirSolicitud();
  avisar('Listo. Catálogo cargado y hoja "' + HOJA_SOLICITUD + '" construida. Ya puedes empezar a llenar.');
}

// ====================== 1) ACTUALIZAR CATÁLOGO ======================
function actualizarCatalogo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const catalogo = SpreadsheetApp.openById(CATALOG_ID);

  const shU = catalogo.getSheetByName(TAB_UNIDADES);
  const shD = catalogo.getSheetByName(TAB_DRIVERS);
  if (!shU || !shD) {
    avisar('No encontré las pestañas "' + TAB_UNIDADES + '" o "' + TAB_DRIVERS + '" en el catálogo. Revisa CATALOG_ID / nombres de pestaña en el código.');
    return;
  }

  const datosU = shU.getDataRange().getValues();
  const hdrU = datosU[0].map(function (h) { return String(h).trim().toUpperCase(); });
  const idxU = {
    proveedor: hdrU.indexOf('PROVEEDOR'),
    categoria: idxTolerante(hdrU, ['CATEGORIA COMUN']),
    pax: idxTolerante(hdrU, ['CANTIDAD PAX']),
    placa: hdrU.indexOf('PLACAS'),
    anio: hdrU.indexOf('AÑO'),
    costoTotal: idxTolerante(hdrU, ['COSTO TOTAL']),
    tarifaForanea: idxTolerante(hdrU, ['TARIFA FORANEA']),
    transfer: idxTolerante(hdrU, ['TRASFER', 'TRANSFER']),
    validacion: idxTolerante(hdrU, ['VALIDACIÓN', 'VALIDACION']),
    estatus: idxTolerante(hdrU, ['ESTATUS DE UNIDAD', 'ESTATUS'])
  };

  const unidades = [];
  for (var i = 1; i < datosU.length; i++) {
    const row = datosU[i];
    const prov = row[idxU.proveedor], cat = row[idxU.categoria], placa = row[idxU.placa];
    if (!prov || !placa || !cat || String(cat).trim().toUpperCase() === 'X') continue;
    const estatus = idxU.estatus > -1 ? String(row[idxU.estatus] || '').trim() : '';
    unidades.push({
      proveedor: normProv(prov), categoria: String(cat).trim().toUpperCase(),
      placa: String(placa).trim(), anio: row[idxU.anio], pax: idxU.pax > -1 ? row[idxU.pax] : '',
      tarifaDiaria: numOr0(row[idxU.costoTotal]), tarifaForanea: numOr0(row[idxU.tarifaForanea]),
      tarifaTransfer: numOr0(row[idxU.transfer]),
      validacion: (row[idxU.validacion] === true || String(row[idxU.validacion]).toUpperCase() === 'TRUE'),
      estatus: estatus
    });
  }

  const datosD = shD.getDataRange().getValues();
  const hdrD = datosD[0].map(function (h) { return String(h).trim().toUpperCase(); });
  const idxD = { proveedor: hdrD.indexOf('PROVEEDOR'), nombre: hdrD.indexOf('DRIVER'), validacion: idxTolerante(hdrD, ['VALIDACION']) };
  const choferes = [];
  for (var j = 1; j < datosD.length; j++) {
    const row = datosD[j];
    if (!row[idxD.proveedor] || !row[idxD.nombre]) continue;
    choferes.push({
      proveedor: normProv(row[idxD.proveedor]), nombre: String(row[idxD.nombre]).trim(),
      validacion: (row[idxD.validacion] === true || String(row[idxD.validacion]).toUpperCase() === 'TRUE')
    });
  }

  escribirBDD(ss, unidades, choferes);
  reconstruirListas(ss, unidades);
  refrescarValidacionesFilasLlenas(ss);

  // Si la hoja ya existe, reaplicamos la lista de Vehículo (CAT_LIST pudo cambiar de tamaño)
  var sol = hojaSolicitud(ss);
  if (sol) aplicarValidacionesEstaticas(sol);

  avisar('Catálogo actualizado: ' + unidades.length + ' unidades, ' + choferes.length + ' choferes.');
}

function normProv(p) { return String(p).replace(/\s+/g, ' ').trim(); }
function numOr0(v) { return (typeof v === 'number') ? v : 0; }
function idxTolerante(headerUpper, candidatos) {
  for (var i = 0; i < candidatos.length; i++) {
    for (var j = 0; j < headerUpper.length; j++) {
      if (headerUpper[j].indexOf(candidatos[i]) !== -1) return j;
    }
  }
  return -1;
}

function escribirBDD(ss, unidades, choferes) {
  var disponibles = unidades.filter(function (u) { return u.estatus.toUpperCase() !== 'NO DISPONIBLE'; });

  var bu = ss.getSheetByName('BDD Unidades') || ss.insertSheet('BDD Unidades');
  bu.clearContents();
  var buHdr = ['Proveedor', 'Categoria', 'Placa', 'Año', 'Pax', 'Tarifa Diaria', 'Tarifa Foránea', 'Tarifa Transfer', 'Validación', 'Estatus'];
  bu.getRange(1, 1, 1, buHdr.length).setValues([buHdr]).setFontWeight('bold');
  if (disponibles.length) {
    var rowsU = disponibles.map(function (u) {
      return [u.proveedor, u.categoria, u.placa, u.anio, u.pax, u.tarifaDiaria, u.tarifaForanea, u.tarifaTransfer, u.validacion, u.estatus];
    });
    bu.getRange(2, 1, rowsU.length, buHdr.length).setValues(rowsU);
  }

  var bd = ss.getSheetByName('BDD Choferes') || ss.insertSheet('BDD Choferes');
  bd.clearContents();
  var bdHdr = ['Proveedor', 'Chofer', 'Validación'];
  bd.getRange(1, 1, 1, bdHdr.length).setValues([bdHdr]).setFontWeight('bold');
  if (choferes.length) {
    var rowsD = choferes.map(function (d) { return [d.proveedor, d.nombre, d.validacion]; });
    bd.getRange(2, 1, rowsD.length, bdHdr.length).setValues(rowsD);
  }

  bu.autoResizeColumns(1, buHdr.length);
  bd.autoResizeColumns(1, bdHdr.length);
}

// ====================== LISTAS EN CASCADA ======================
// Solo el Vehículo usa un rango con nombre (CAT_LIST) porque es una lista fija.
// Proveedor, Placas y Chofer se calculan al vuelo leyendo 'BDD Unidades' /
// 'BDD Choferes' en cada edición (ver CASCADA DINÁMICA más abajo).
function reconstruirListas(ss, unidades) {
  ss.getNamedRanges().forEach(function (nr) { if (nr.getName() === 'CAT_LIST') nr.remove(); });
  var ls = ss.getSheetByName('Listas') || ss.insertSheet('Listas');
  ls.clear();

  var disponibles = unidades.filter(function (u) { return u.estatus.toUpperCase() !== 'NO DISPONIBLE'; });
  var cats = uniqueSorted(disponibles.map(function (u) { return u.categoria; }));
  ls.getRange(1, 1).setValue('Categoria');
  if (cats.length) ls.getRange(2, 1, cats.length, 1).setValues(cats.map(function (c) { return [c]; }));
  ss.setNamedRange('CAT_LIST', ls.getRange(2, 1, Math.max(1, cats.length), 1));

  ls.hideSheet();
}

function uniqueSorted(arr) {
  var seen = {};
  var out = [];
  arr.forEach(function (x) { if (x && !seen[x]) { seen[x] = true; out.push(x); } });
  out.sort();
  return out;
}

// ====================== CASCADA DINÁMICA (onEdit) ======================
// Trigger simple: se activa solo. Soporta pegar varias filas a la vez.
function onEdit(e) {
  try {
    var sh = e.range.getSheet();
    if (sh.getName() !== HOJA_SOLICITUD) return;
    var r1 = Math.max(e.range.getRow(), DATA_START);
    var r2 = Math.min(e.range.getLastRow(), DATA_END);
    if (r1 > r2) return;
    var c1 = e.range.getColumn(), c2 = e.range.getLastColumn();
    var tocaVehiculo = c1 <= COL.VEHICULO && COL.VEHICULO <= c2;
    var tocaProveedor = c1 <= COL.PROVEEDOR && COL.PROVEEDOR <= c2;
    if (!tocaVehiculo && !tocaProveedor) return;

    for (var row = r1; row <= r2; row++) {
      if (tocaVehiculo) {
        // Si se pegó vehículo Y proveedor juntos, no borramos el proveedor pegado
        if (!tocaProveedor) {
          sh.getRange(row, COL.PROVEEDOR).clearContent().clearDataValidations();
          sh.getRange(row, COL.PLACAS).clearContent().clearDataValidations();
          sh.getRange(row, COL.CHOFER).clearContent().clearDataValidations();
        }
        aplicarValidacionProveedor(sh, row);
      }
      if (tocaProveedor) {
        if (!tocaVehiculo) {
          sh.getRange(row, COL.PLACAS).clearContent().clearDataValidations();
          sh.getRange(row, COL.CHOFER).clearContent().clearDataValidations();
        }
        aplicarValidacionPlacas(sh, row);
        aplicarValidacionChofer(sh, row);
      }
    }
  } catch (err) {
    Logger.log('Error en onEdit: ' + err);
  }
}

// Lee 'BDD Unidades' y filtra por categoría -> lista de proveedores
function aplicarValidacionProveedor(sh, row) {
  var ss = sh.getParent();
  var vehiculo = sh.getRange(row, COL.VEHICULO).getValue();
  var cell = sh.getRange(row, COL.PROVEEDOR);
  if (!vehiculo) { cell.clearDataValidations(); return; }

  var bu = ss.getSheetByName('BDD Unidades');
  if (!bu || bu.getLastRow() < 2) { cell.clearDataValidations(); return; }
  var datos = bu.getRange(2, 1, bu.getLastRow() - 1, 2).getValues(); // A Proveedor, B Categoria

  var provs = [];
  var visto = {};
  datos.forEach(function (r) {
    if (r[1] === vehiculo && r[0] && !visto[r[0]]) { visto[r[0]] = true; provs.push(r[0]); }
  });
  if (provs.length === 0) { cell.clearDataValidations(); return; }
  cell.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(provs, true).setAllowInvalid(false).build());
}

// Lee 'BDD Unidades' y filtra por categoría + proveedor -> lista de placas
function aplicarValidacionPlacas(sh, row) {
  var ss = sh.getParent();
  var vehiculo = sh.getRange(row, COL.VEHICULO).getValue();
  var proveedor = sh.getRange(row, COL.PROVEEDOR).getValue();
  var cell = sh.getRange(row, COL.PLACAS);
  if (!vehiculo || !proveedor) { cell.clearDataValidations(); return; }

  var bu = ss.getSheetByName('BDD Unidades');
  if (!bu || bu.getLastRow() < 2) { cell.clearDataValidations(); return; }
  var datos = bu.getRange(2, 1, bu.getLastRow() - 1, 3).getValues(); // A Proveedor, B Categoria, C Placa

  var placas = [];
  datos.forEach(function (r) {
    if (r[1] === vehiculo && r[0] === proveedor && r[2]) placas.push(r[2]);
  });
  if (placas.length === 0) { cell.clearDataValidations(); return; }
  cell.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(placas, true).setAllowInvalid(false).build());
}

// Lee 'BDD Choferes' y filtra por proveedor -> lista de choferes
function aplicarValidacionChofer(sh, row) {
  var ss = sh.getParent();
  var proveedor = sh.getRange(row, COL.PROVEEDOR).getValue();
  var cell = sh.getRange(row, COL.CHOFER);
  if (!proveedor) { cell.clearDataValidations(); return; }

  var bd = ss.getSheetByName('BDD Choferes');
  if (!bd || bd.getLastRow() < 2) { cell.clearDataValidations(); return; }
  var datos = bd.getRange(2, 1, bd.getLastRow() - 1, 2).getValues(); // A Proveedor, B Chofer

  var nombres = [];
  datos.forEach(function (r) { if (r[0] === proveedor && r[1]) nombres.push(r[1]); });
  if (nombres.length === 0) { cell.clearDataValidations(); return; }
  cell.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(nombres, true).setAllowInvalid(false).build());
}

// Reaplica las validaciones dinámicas en filas que YA tienen datos capturados
function refrescarValidacionesFilasLlenas(ss) {
  var sh = hojaSolicitud(ss);
  if (!sh) return;
  var vehiculos = sh.getRange(DATA_START, COL.VEHICULO, N_FILAS, 1).getValues();
  var proveedores = sh.getRange(DATA_START, COL.PROVEEDOR, N_FILAS, 1).getValues();
  for (var i = 0; i < N_FILAS; i++) {
    var row = DATA_START + i;
    if (vehiculos[i][0]) aplicarValidacionProveedor(sh, row);
    if (proveedores[i][0]) { aplicarValidacionPlacas(sh, row); aplicarValidacionChofer(sh, row); }
  }
}

// ====================== 2) CONSTRUIR / REPARAR SOLICITUD ======================
// Seguro de re-ejecutar: solo toca encabezados, columnas de fórmula, validaciones
// y formato. NUNCA borra las columnas de captura manual si ya tienen datos.
function construirSolicitud() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sol = hojaSolicitud(ss);
  const esNueva = !sol;
  if (!sol) sol = ss.insertSheet(HOJA_SOLICITUD, 0);
  ss.setActiveSheet(sol);
  sol.showSheet();

  if (sol.getMaxColumns() < LAST_COL) sol.insertColumnsAfter(sol.getMaxColumns(), LAST_COL - sol.getMaxColumns());
  sol.setColumnWidths(1, LAST_COL, 90);
  sol.getRange(1, 1, 1, LAST_COL).setFontFamily('Arial');

  // ---- Título ----
  sol.getRange('A1').setValue('SOLICITUD DE TRANSPORTE — CINEMA SHUTTLE')
    .setFontWeight('bold').setFontSize(16).setFontColor('#191C16');
  sol.getRange('A2').setValue('Llena solo las celdas VERDES. Placas, año, validación y tarifas se llenan solos al elegir vehículo → proveedor → placa. '
    + 'En el calendario, cada día lleva un código de 2 letras: 1ª = modalidad (1 Full day, F Foráneo, T Transfer), 2ª = estatus (S Solicitado, C Confirmado, P Tentativo, X Cancelado). Ej: "1S" = full day solicitado, "FC" = foráneo confirmado.')
    .setFontStyle('italic').setFontSize(9).setFontColor('#555555');
  try { sol.getRange(2, 1, 1, 20).merge(); } catch (e) {}

  // ---- Encabezado del proyecto ----
  setLabel(sol, 3, 1, 'PROYECTO:');
  setLabel(sol, 3, 6, 'CLIENTE / USUARIO:');
  setLabel(sol, 4, 1, 'PRODUCTOR:');
  setLabel(sol, 4, 6, 'CAPITÁN DE TRANSPORTACIÓN:');
  setLabel(sol, 5, 1, 'ESPECIALISTA DE TRANSPORTE:');
  setLabel(sol, 5, 6, 'FECHA DE INICIO:');
  mergeInput(sol, 3, 2, 3);
  mergeInput(sol, 3, 7, 3);
  mergeInput(sol, 4, 2, 3);
  mergeInput(sol, 4, 7, 3);
  mergeInput(sol, 5, 2, 3);
  var fechaCell = celdaFechaInicio(sol);
  fechaCell.setNumberFormat('dd/mm/yyyy').setFontColor('#6E8241');
  fechaCell.setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false)
    .setHelpText('Escribe una fecha válida (dd/mm/aaaa)').build());
  ss.setNamedRange('FECHA_INICIO', fechaCell);
  try { sol.getRange(5, 8, 1, 5).merge(); } catch (e) {}
  sol.getRange(5, 8).setValue('(el calendario de abajo se genera solo a partir de esta fecha, ' + N_DIAS + ' días)')
    .setFontStyle('italic').setFontSize(8).setFontColor('#777777');

  // ---- Grupos y encabezados ----
  group(sol, GROUP_ROW, COL.CONCEPTO, COL.PAX, 'DATOS DE LA UNIDAD (Pax/Bilingüe = lo pedido por el cliente)');
  group(sol, GROUP_ROW, COL.CHOFER, COL.CEL, 'CHOFER');
  group(sol, GROUP_ROW, COL.VALID_UNIDAD, COL.VALID_DRIVER, 'VALIDACIÓN DOCUMENTAL');
  group(sol, GROUP_ROW, COL.DIA_START, COL.DIA_END, 'CALENDARIO DE USO   (1=Full · F=Foráneo · T=Transfer)  +  (S=Solicitado · C=Confirmado · P=Tentativo · X=Cancelado)');
  group(sol, GROUP_ROW, COL.D_FULL, COL.D_CANCELADO, 'DÍAS POR TIPO (automático)');
  group(sol, GROUP_ROW, COL.T_DIA, COL.T_TRANSFER, 'TARIFAS (automático)');
  group(sol, GROUP_ROW, COL.HORAS_EXTRA, COL.VIATICOS, 'EXTRAS (manual)');
  group(sol, GROUP_ROW, COL.TOTAL, COL.OBS, ' ');

  // Estos encabezados son los que lee el consolidador: no los cambies.
  var headers = {};
  headers[COL.CONCEPTO] = 'CONCEPTO'; headers[COL.PAX_SOLICITADO] = 'PAX SOLICITADO'; headers[COL.BILINGUE] = 'CHOFER BILINGÜE';
  headers[COL.VEHICULO] = 'VEHICULO'; headers[COL.PROVEEDOR] = 'PROVEEDOR';
  headers[COL.PLACAS] = 'PLACAS DE LA UNIDAD'; headers[COL.ANIO] = 'AÑO DE LA UNIDAD'; headers[COL.PAX] = 'PAX';
  headers[COL.CHOFER] = 'NOMBRE CHOFER'; headers[COL.CEL] = 'CEL CHOFER';
  headers[COL.VALID_UNIDAD] = 'VALIDACION UNIDAD'; headers[COL.VALID_DRIVER] = 'VALIDACION DRIVER';
  headers[COL.D_FULL] = 'DIAS FULL'; headers[COL.D_FORANEO] = 'DIAS FORANEO'; headers[COL.D_TRANSFER] = 'DIAS TRANSFER'; headers[COL.D_CANCELADO] = 'DIAS CANCELADOS';
  headers[COL.T_DIA] = 'TARIFA DIARIA'; headers[COL.T_FORANEA] = 'TARIFA FORANEA'; headers[COL.T_TRANSFER] = 'TARIFA TRANSFER';
  headers[COL.HORAS_EXTRA] = 'HORAS EXTRA'; headers[COL.T_HORA] = '$/HORA EXTRA'; headers[COL.VIATICOS] = 'VIATICOS EXTRAS';
  headers[COL.TOTAL] = 'COSTO TOTAL ESTIMADO'; headers[COL.INCIDENCIA] = 'INCIDENCIA'; headers[COL.OBS] = 'OBSERVACIONES TRANSPORTACIÓN';
  Object.keys(headers).forEach(function (c) {
    sol.getRange(HDR_ROW, Number(c)).setValue(headers[c]).setFontWeight('bold').setFontColor('#FFFFFF')
      .setBackground('#191C16').setHorizontalAlignment('center').setWrap(true);
  });

  // encabezados de días (fórmulas). Si no hay fecha de inicio muestran "DIA n".
  var fechaA1 = '$' + colLetra(CELDA.FECHA_INICIO[1]) + '$' + CELDA.FECHA_INICIO[0];
  var formulasDias = [];
  for (var i = 0; i < N_DIAS; i++) {
    formulasDias.push('=IF(ISNUMBER(' + fechaA1 + '),' + fechaA1 + '+' + i + ',"DIA ' + (i + 1) + '")');
    sol.setColumnWidth(COL.DIA_START + i, 46);
  }
  sol.getRange(HDR_ROW, COL.DIA_START, 1, N_DIAS).setFormulas([formulasDias]).setNumberFormat('dd-mmm')
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#191C16').setHorizontalAlignment('center');

  construirFilasDeCaptura(sol);
  aplicarFormulasYFormato(sol);
  aplicarValidacionesEstaticas(sol);
  aplicarFormatoCondicional(sol);

  // ---- fila de totales ----
  var totRow = DATA_END + 2;
  if (sol.getMaxRows() < totRow + 2) sol.insertRowsAfter(sol.getMaxRows(), totRow + 2 - sol.getMaxRows());
  try { sol.getRange(totRow, COL.CONCEPTO, 1, COL.TOTAL - COL.CONCEPTO).merge(); } catch (e) {}
  sol.getRange(totRow, COL.CONCEPTO).setValue('COSTO TOTAL DEL PROYECTO').setFontWeight('bold').setFontSize(12);
  var gt = sol.getRange(totRow, COL.TOTAL);
  gt.setFormula('=SUM(' + sol.getRange(DATA_START, COL.TOTAL, N_FILAS, 1).getA1Notation() + ')');
  gt.setFontWeight('bold').setFontSize(12).setNumberFormat('$#,##0').setBackground('#E9EEDD');

  sol.setFrozenRows(HDR_ROW);
  sol.setHiddenGridlines(true);

  if (esNueva) avisar('Hoja "' + HOJA_SOLICITUD + '" construida. Corre ahora "1) Actualizar catálogo" si no lo has hecho.');
  else avisar('Hoja "' + HOJA_SOLICITUD + '" verificada/reparada. Tus datos capturados no se tocaron.');
}

function setLabel(sh, r, c, text) {
  sh.getRange(r, c).setValue(text).setFontWeight('bold').setBackground('#E9EEDD');
}
function mergeInput(sh, r, c, span) {
  var rng = sh.getRange(r, c, 1, span);
  try { rng.merge(); } catch (e) {}
  rng.setFontColor('#6E8241');
}
function group(sh, r, c1, c2, text) {
  var rng = sh.getRange(r, c1, 1, c2 - c1 + 1);
  try { rng.merge(); } catch (e) {}
  rng.getCell(1, 1).setValue(text);
  rng.setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#6E8241')
    .setHorizontalAlignment('center').setWrap(true);
}
function colLetra(n) {
  var s = '';
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// ====================== FILAS DE CAPTURA ======================
function construirFilasDeCaptura(sol) {
  var maxRows = sol.getMaxRows();
  if (maxRows < DATA_END) sol.insertRowsAfter(maxRows, DATA_END - maxRows);
  sol.getRange(DATA_START, 1, N_FILAS, LAST_COL).setBorder(true, true, true, true, true, true, '#D9D9D9', SpreadsheetApp.BorderStyle.SOLID);
}

function aplicarFormulasYFormato(sol) {
  var U = '\'BDD Unidades\'!', D = '\'BDD Choferes\'!';
  function col2array(fn) {
    var out = [];
    for (var i = 0; i < N_FILAS; i++) out.push([fn(DATA_START + i)]);
    return out;
  }
  function a1(r, c) { return colLetra(c) + r; }
  function buscarUnidad(colRes, fallback) {
    return function (r) {
      return '=IF(' + a1(r, COL.PLACAS) + '="",' + fallback + ',IFERROR(INDEX(' + U + '$' + colRes + ':$' + colRes + ',MATCH(' + a1(r, COL.PLACAS) + ',' + U + '$C:$C,0)),' + fallback + '))';
    };
  }
  function diaRange(r) { return a1(r, COL.DIA_START) + ':' + a1(r, COL.DIA_END); }

  sol.getRange(DATA_START, COL.ANIO, N_FILAS, 1).setFormulas(col2array(buscarUnidad('D', '""')));
  sol.getRange(DATA_START, COL.PAX, N_FILAS, 1).setFormulas(col2array(buscarUnidad('E', '""')));
  sol.getRange(DATA_START, COL.VALID_UNIDAD, N_FILAS, 1).setFormulas(col2array(buscarUnidad('I', '""')));
  sol.getRange(DATA_START, COL.VALID_DRIVER, N_FILAS, 1).setFormulas(col2array(function (r) {
    return '=IF(' + a1(r, COL.CHOFER) + '="","",IFERROR(INDEX(' + D + '$C:$C,MATCH(' + a1(r, COL.CHOFER) + ',' + D + '$B:$B,0)),""))';
  }));
  // Días por tipo (excluyen cancelados: "1X" no cuenta como Full)
  sol.getRange(DATA_START, COL.D_FULL, N_FILAS, 1).setFormulas(col2array(function (r) { return '=COUNTIFS(' + diaRange(r) + ',"1*",' + diaRange(r) + ',"<>1X")'; }));
  sol.getRange(DATA_START, COL.D_FORANEO, N_FILAS, 1).setFormulas(col2array(function (r) { return '=COUNTIFS(' + diaRange(r) + ',"F*",' + diaRange(r) + ',"<>FX")'; }));
  sol.getRange(DATA_START, COL.D_TRANSFER, N_FILAS, 1).setFormulas(col2array(function (r) { return '=COUNTIFS(' + diaRange(r) + ',"T*",' + diaRange(r) + ',"<>TX")'; }));
  sol.getRange(DATA_START, COL.D_CANCELADO, N_FILAS, 1).setFormulas(col2array(function (r) { return '=COUNTIF(' + diaRange(r) + ',"?X")'; }));
  // Tarifas (fallback 0)
  sol.getRange(DATA_START, COL.T_DIA, N_FILAS, 1).setFormulas(col2array(buscarUnidad('F', '0')));
  sol.getRange(DATA_START, COL.T_FORANEA, N_FILAS, 1).setFormulas(col2array(buscarUnidad('G', '0')));
  sol.getRange(DATA_START, COL.T_TRANSFER, N_FILAS, 1).setFormulas(col2array(buscarUnidad('H', '0')));
  // Costo total (días cancelados no se cobran)
  sol.getRange(DATA_START, COL.TOTAL, N_FILAS, 1).setFormulas(col2array(function (r) {
    return '=IFERROR(' + a1(r, COL.D_FULL) + '*' + a1(r, COL.T_DIA)
      + '+' + a1(r, COL.D_FORANEO) + '*' + a1(r, COL.T_FORANEA)
      + '+' + a1(r, COL.D_TRANSFER) + '*' + a1(r, COL.T_TRANSFER)
      + '+N(' + a1(r, COL.HORAS_EXTRA) + ')*N(' + a1(r, COL.T_HORA) + ')'
      + '+N(' + a1(r, COL.VIATICOS) + '),0)';
  }));

  // Formatos numéricos
  sol.getRange(DATA_START, COL.T_DIA, N_FILAS, 3).setNumberFormat('$#,##0');
  sol.getRange(DATA_START, COL.VIATICOS, N_FILAS, 1).setNumberFormat('$#,##0');
  sol.getRange(DATA_START, COL.T_HORA, N_FILAS, 1).setNumberFormat('$#,##0');
  sol.getRange(DATA_START, COL.TOTAL, N_FILAS, 1).setNumberFormat('$#,##0').setFontWeight('bold');
  sol.getRange(DATA_START, COL.D_FULL, N_FILAS, 4).setHorizontalAlignment('center');
  sol.getRange(DATA_START, COL.DIA_START, N_FILAS, N_DIAS).setHorizontalAlignment('center');

  // color de fuente: verde = manual, negro = fórmula
  [COL.CONCEPTO, COL.PAX_SOLICITADO, COL.BILINGUE, COL.VEHICULO, COL.PROVEEDOR, COL.PLACAS, COL.CEL, COL.CHOFER, COL.HORAS_EXTRA, COL.VIATICOS, COL.INCIDENCIA, COL.OBS]
    .forEach(function (c) { sol.getRange(DATA_START, c, N_FILAS, 1).setFontColor('#6E8241'); });
  [COL.ANIO, COL.PAX, COL.VALID_UNIDAD, COL.VALID_DRIVER, COL.D_FULL, COL.D_FORANEO, COL.D_TRANSFER, COL.D_CANCELADO, COL.T_DIA, COL.T_FORANEA, COL.T_TRANSFER, COL.TOTAL]
    .forEach(function (c) { sol.getRange(DATA_START, c, N_FILAS, 1).setFontColor('#000000'); });
  sol.getRange(DATA_START, COL.DIA_START, N_FILAS, N_DIAS).setFontColor('#6E8241');

  // $/hora extra: default 300 solo donde esté vacío
  var actuales = sol.getRange(DATA_START, COL.T_HORA, N_FILAS, 1).getValues();
  var nuevos = actuales.map(function (v) { return [v[0] === '' ? 300 : v[0]]; });
  sol.getRange(DATA_START, COL.T_HORA, N_FILAS, 1).setValues(nuevos).setBackground('#FCF1DC');
}

// ====================== VALIDACIONES ESTÁTICAS ======================
function codigosDia() {
  var out = [];
  ['1', 'F', 'T'].forEach(function (m) { ['S', 'C', 'P', 'X'].forEach(function (e) { out.push(m + e); }); });
  return out;
}

function aplicarValidacionesEstaticas(sol) {
  var ss = sol.getParent();

  var catRange = ss.getRangeByName('CAT_LIST');
  if (catRange) {
    var ruleVeh = SpreadsheetApp.newDataValidation().requireValueInRange(catRange, true).setAllowInvalid(false).build();
    sol.getRange(DATA_START, COL.VEHICULO, N_FILAS, 1).setDataValidation(ruleVeh);
  }

  var ruleSiNo = SpreadsheetApp.newDataValidation().requireValueInList(['Sí', 'No'], true).setAllowInvalid(true).build();
  sol.getRange(DATA_START, COL.BILINGUE, N_FILAS, 1).setDataValidation(ruleSiNo);
  sol.getRange(DATA_START, COL.INCIDENCIA, N_FILAS, 1).setDataValidation(ruleSiNo);

  var ruleDia = SpreadsheetApp.newDataValidation().requireValueInList(codigosDia(), true).setAllowInvalid(false).build();
  sol.getRange(DATA_START, COL.DIA_START, N_FILAS, N_DIAS).setDataValidation(ruleDia);

  var ruleNum = SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false).build();
  sol.getRange(DATA_START, COL.HORAS_EXTRA, N_FILAS, 1).setDataValidation(ruleNum);
  sol.getRange(DATA_START, COL.VIATICOS, N_FILAS, 1).setDataValidation(ruleNum);
  // Proveedor / Placas / Chofer: las pone onEdit() según la cascada.
}

// ====================== FORMATO CONDICIONAL ======================
function aplicarFormatoCondicional(sol) {
  sol.clearConditionalFormatRules();
  var rules = [];

  var vuRange = sol.getRange(DATA_START, COL.VALID_UNIDAD, N_FILAS, 1);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=' + colLetra(COL.VALID_UNIDAD) + DATA_START + '=FALSE')
    .setBackground('#F3D6C9').setRanges([vuRange]).build());

  var vdRange = sol.getRange(DATA_START, COL.VALID_DRIVER, N_FILAS, 1);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=' + colLetra(COL.VALID_DRIVER) + DATA_START + '=FALSE')
    .setBackground('#F3D6C9').setRanges([vdRange]).build());

  // doble reserva de la misma placa el mismo día (ignora cancelados)
  var diaGrid = sol.getRange(DATA_START, COL.DIA_START, N_FILAS, N_DIAS);
  var pCol = colLetra(COL.PLACAS), dCol = colLetra(COL.DIA_START);
  var topLeftDia = dCol + DATA_START;
  var placasAbs = '$' + pCol + '$' + DATA_START + ':$' + pCol + '$' + DATA_END;
  var diaColRel = dCol + '$' + DATA_START + ':' + dCol + '$' + DATA_END;
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($' + pCol + DATA_START + '<>"",' + topLeftDia + '<>"",RIGHT(' + topLeftDia + ',1)<>"X",'
      + 'SUMPRODUCT((' + placasAbs + '=$' + pCol + DATA_START + ')*(' + diaColRel + '<>"")*(RIGHT(' + diaColRel + ',1)<>"X"))>1)')
    .setBackground('#F5DCA8').setRanges([diaGrid]).build());

  // Color por ESTATUS (2º carácter del código)
  [
    { letra: 'S', fondo: '#4B5320', texto: '#FFFFFF' }, // Solicitado - verde militar
    { letra: 'C', fondo: '#AACB3D', texto: '#1B1E17' }, // Confirmado - verde limón
    { letra: 'P', fondo: '#F5C947', texto: '#1B1E17' }, // Tentativo - amarillo
    { letra: 'X', fondo: '#C0392B', texto: '#FFFFFF' }  // Cancelado - rojo
  ].forEach(function (e) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=RIGHT(' + topLeftDia + ',1)="' + e.letra + '"')
      .setBackground(e.fondo).setFontColor(e.texto)
      .setRanges([diaGrid]).build());
  });

  sol.setConditionalFormatRules(rules);
}

// ====================== 3) PREPARAR PARA NUEVO PROYECTO ======================
function limpiarParaNuevoProyecto() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert('Esto borra el proyecto/cliente/productor y todas las filas capturadas de ESTE archivo (no las fórmulas ni el catálogo). ¿Continuar?', ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  var sol = hojaSolicitud();
  if (!sol) { ui.alert('No existe la hoja "' + HOJA_SOLICITUD + '" todavía. Corre primero "2) Construir / reparar".'); return; }

  Object.keys(CELDA).forEach(function (k) { sol.getRange(CELDA[k][0], CELDA[k][1]).clearContent(); });

  var manualCols = [COL.CONCEPTO, COL.PAX_SOLICITADO, COL.BILINGUE, COL.VEHICULO, COL.PROVEEDOR, COL.PLACAS, COL.CHOFER, COL.CEL, COL.HORAS_EXTRA, COL.VIATICOS, COL.INCIDENCIA, COL.OBS];
  manualCols.forEach(function (c) { sol.getRange(DATA_START, c, N_FILAS, 1).clearContent(); });
  sol.getRange(DATA_START, COL.DIA_START, N_FILAS, N_DIAS).clearContent();
  [COL.PROVEEDOR, COL.PLACAS, COL.CHOFER].forEach(function (c) { sol.getRange(DATA_START, c, N_FILAS, 1).clearDataValidations(); });
  sol.getRange(DATA_START, COL.T_HORA, N_FILAS, 1).setValue(300);

  ui.alert('Listo. Archivo en blanco y listo para el nuevo proyecto.\n\nRecuerda: el nombre del archivo debe llevar el nombre del proyecto y guardarse en la carpeta de proyectos para que la base maestra lo lea.');
}

// ====================== 4) PROTECCIÓN ======================
function protegerHoja() {
  var sol = hojaSolicitud();
  var ui = SpreadsheetApp.getUi();
  if (!sol) { ui.alert('Construye primero la hoja (opción 2 del menú).'); return; }

  sol.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) { p.remove(); });
  sol.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) { p.remove(); });

  var protection = sol.protect().setDescription('Estructura y fórmulas — solo equipo de transporte');
  try { if (protection.canDomainEdit()) protection.setDomainEdit(false); } catch (e) {}

  var miCorreo = '';
  try { miCorreo = Session.getActiveUser().getEmail(); } catch (e) {}
  var editores = EDITORES_EQUIPO.slice();
  if (miCorreo && editores.indexOf(miCorreo) === -1) editores.push(miCorreo);

  protection.removeEditors(protection.getEditors());
  if (editores.length) protection.addEditors(editores);

  var abiertos = Object.keys(CELDA).map(function (k) { return sol.getRange(CELDA[k][0], CELDA[k][1]); });
  [COL.CONCEPTO, COL.PAX_SOLICITADO, COL.BILINGUE, COL.VEHICULO, COL.PROVEEDOR, COL.PLACAS, COL.CHOFER, COL.CEL,
   COL.HORAS_EXTRA, COL.T_HORA, COL.VIATICOS, COL.INCIDENCIA, COL.OBS].forEach(function (c) {
    abiertos.push(sol.getRange(DATA_START, c, N_FILAS, 1));
  });
  abiertos.push(sol.getRange(DATA_START, COL.DIA_START, N_FILAS, N_DIAS));
  protection.setUnprotectedRanges(abiertos);

  ui.alert('Hoja protegida ✅\n\nCualquiera con el archivo puede llenar la solicitud. Solo estas personas pueden tocar fórmulas, encabezados o insertar/borrar filas y columnas:\n\n' + editores.join('\n') + '\n\nPara agregar a alguien, edita EDITORES_EQUIPO en el código y vuelve a correr esta opción.');
}

function agregarmeComoEditor() {
  var correo = Session.getActiveUser().getEmail();
  if (!correo) { avisar('No pude leer tu correo automáticamente. Agrégalo a mano en EDITORES_EQUIPO, dentro del código.'); return; }
  avisar('Tu correo (' + correo + ') se incluye automáticamente cada vez que corras "4) Proteger hoja", junto con los de EDITORES_EQUIPO.');
}

function quitarProteccion() {
  var sol = hojaSolicitud();
  var ui = SpreadsheetApp.getUi();
  if (!sol) return;
  var resp = ui.alert('Esto quita TODA la protección de la hoja. Úsalo solo temporalmente para reparar algo. ¿Continuar?', ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  sol.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) { p.remove(); });
  sol.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) { p.remove(); });
  ui.alert('Protección quitada. No olvides volver a correr "4) Proteger hoja" cuando termines.');
}

// ====================== 5) FORMULARIO (PANEL LATERAL + WEB APP) ======================
// El MISMO archivo Formulario.html sirve para:
//  - el panel lateral dentro de la hoja (menú "📝 Abrir formulario de captura")
//  - la página web para clientes (Implementar → Aplicación web)
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Formulario')
    .setTitle('Cinema Shuttle — Solicitud de Transporte')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function instruccionesFormulario() {
  avisar(
    'Para publicar el formulario de clientes:\n\n' +
    '1. En el editor de Apps Script: Implementar → Nueva implementación.\n' +
    '2. Tipo: Aplicación web.\n' +
    '3. Ejecutar como: Yo (tu cuenta).\n' +
    '4. Quién tiene acceso: Cualquier usuario.\n' +
    '5. Implementar → copia la URL (termina en /exec) y compártela.\n\n' +
    'IMPORTANTE: cada vez que cambies el código, ve a Implementar → Administrar implementaciones → ✏️ → Versión: "Nueva versión" → Implementar. ' +
    'Si no lo haces, la URL sigue mostrando la versión vieja.\n\n' +
    'Cada copia del machote (cada proyecto) tiene su propia URL: hay que implementarla en esa copia.'
  );
}

// Datos para pintar el formulario. OJO: google.script.run NO puede devolver
// objetos Date (el formulario recibe null y se queda "Cargando..."), por eso
// todo se manda como texto.
function obtenerOpcionesFormulario() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sol = hojaSolicitud(ss);
  var catRange = ss.getRangeByName('CAT_LIST');
  var categorias = catRange ? catRange.getValues().map(function (r) { return String(r[0]); }).filter(function (v) { return v; }) : [];

  var inicio = sol ? leerFechaInicio(sol) : null;
  var fin = inicio ? sumarDias(inicio, N_DIAS - 1) : null;
  return {
    categorias: categorias,
    fechaInicio: inicio,   // 'yyyy-MM-dd' o null
    fechaFin: fin,
    nDias: N_DIAS,
    proyecto: sol ? String(sol.getRange(CELDA.PROYECTO[0], CELDA.PROYECTO[1]).getDisplayValue()) : '',
    hojaLista: !!sol
  };
}

// Lee la fecha de inicio como texto 'yyyy-MM-dd' en la zona horaria de la hoja
function leerFechaInicio(sol) {
  var v = celdaFechaInicio(sol).getValue();
  if (!(v instanceof Date) || isNaN(v.getTime())) return null;
  return Utilities.formatDate(v, sol.getParent().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
}

// Aritmética de fechas en "número de día" (sin horas ni zonas horarias)
function diaNum(iso) {
  var p = String(iso).split('-');
  return Math.round(Date.UTC(Number(p[0]), Number(p[1]) - 1, Number(p[2])) / 86400000);
}
function sumarDias(iso, n) {
  var d = new Date((diaNum(iso) + n) * 86400000);
  return d.getUTCFullYear() + '-' + ('0' + (d.getUTCMonth() + 1)).slice(-2) + '-' + ('0' + d.getUTCDate()).slice(-2);
}
function esISO(s) { return /^\d{4}-\d{2}-\d{2}$/.test(String(s || '')); }

// Escribe la solicitud en la(s) siguiente(s) fila(s) vacía(s). NUNCA asigna
// proveedor/placa/chofer — eso lo hace el especialista de transporte después.
function crearSolicitudCliente(datos) {
  datos = datos || {};
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(20000)) return { ok: false, mensaje: 'Hay otra solicitud guardándose en este momento. Intenta de nuevo en unos segundos.' };
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sol = hojaSolicitud(ss);
    if (!sol) return { ok: false, mensaje: 'La hoja "' + HOJA_SOLICITUD + '" todavía no existe. Avisa a tu contacto en Cinema Shuttle.' };

    var vehiculo = String(datos.vehiculo || '').trim();
    var concepto = String(datos.concepto || '').trim();
    if (!vehiculo) return { ok: false, mensaje: 'Elige el tipo de vehículo.' };
    if (!concepto) return { ok: false, mensaje: 'Escribe para qué es la unidad (concepto).' };
    if (!esISO(datos.fechaDesde) || !esISO(datos.fechaHasta)) return { ok: false, mensaje: 'Elige las fechas "Desde" y "Hasta".' };
    if (diaNum(datos.fechaHasta) < diaNum(datos.fechaDesde)) return { ok: false, mensaje: 'La fecha "Hasta" no puede ser antes de "Desde".' };

    var inicio = leerFechaInicio(sol);
    if (!inicio) return { ok: false, mensaje: 'El proyecto todavía no tiene "Fecha de inicio". Avisa a tu contacto en Cinema Shuttle.' };

    // columnas del calendario que caen dentro del rango pedido
    var base = diaNum(inicio);
    var desdeIdx = diaNum(datos.fechaDesde) - base;
    var hastaIdx = diaNum(datos.fechaHasta) - base;
    var i1 = Math.max(0, desdeIdx), i2 = Math.min(N_DIAS - 1, hastaIdx);
    if (i1 > i2) {
      return { ok: false, mensaje: 'Las fechas quedan fuera del calendario del proyecto (' + inicio + ' a ' + sumarDias(inicio, N_DIAS - 1) + ').' };
    }
    var fueraDeRango = (desdeIdx < 0 || hastaIdx > N_DIAS - 1);

    var modalidad = ['1', 'F', 'T'].indexOf(String(datos.modalidad)) > -1 ? String(datos.modalidad) : '1';
    var codigo = modalidad + 'S'; // siempre entra como "Solicitado"

    var cantidad = Math.floor(Number(datos.cantidad) || 1);
    cantidad = Math.max(1, Math.min(MAX_UNIDADES_POR_ENVIO, cantidad));

    var filas = filasVacias(sol, cantidad);
    if (filas.length < cantidad) {
      return { ok: false, mensaje: 'Solo quedan ' + filas.length + ' filas libres. Avisa a tu contacto en Cinema Shuttle.' };
    }

    var obs = [];
    if (datos.solicitante) obs.push('Solicitó: ' + String(datos.solicitante).trim());
    if (datos.horario) obs.push('Horario: ' + String(datos.horario).trim());
    if (datos.observaciones) obs.push(String(datos.observaciones).trim());
    if (fueraDeRango) obs.push('(Fechas pedidas: ' + datos.fechaDesde + ' a ' + datos.fechaHasta + ')');

    var dias = [];
    for (var d = 0; d < N_DIAS; d++) dias.push(d >= i1 && d <= i2 ? codigo : '');

    filas.forEach(function (fila, n) {
      var etiqueta = cantidad > 1 ? concepto + ' ' + (n + 1) : concepto;
      sol.getRange(fila, COL.CONCEPTO, 1, 4).setValues([[etiqueta, datos.paxSolicitado || '', datos.bilingue ? 'Sí' : 'No', vehiculo]]);
      sol.getRange(fila, COL.DIA_START, 1, N_DIAS).setValues([dias]);
      if (obs.length) sol.getRange(fila, COL.OBS).setValue(obs.join(' · '));
      // setValue desde código NO dispara onEdit: aplicamos la lista de proveedores a mano
      aplicarValidacionProveedor(sol, fila);
    });
    SpreadsheetApp.flush();

    var msg = (cantidad > 1 ? cantidad + ' unidades registradas' : 'Solicitud registrada')
      + ' (' + (i2 - i1 + 1) + ' día(s)). Tu equipo de transporte la va a revisar y asignar la unidad.';
    if (fueraDeRango) msg += ' Ojo: algunos días quedaron fuera del calendario del proyecto y se anotaron en observaciones.';
    return { ok: true, mensaje: msg, filas: filas };
  } catch (err) {
    return { ok: false, mensaje: 'Error al guardar: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

// Filas libres = sin concepto, sin vehículo y sin días capturados
function filasVacias(sol, cuantas) {
  var datos = sol.getRange(DATA_START, 1, N_FILAS, COL.DIA_END).getValues();
  var out = [];
  for (var i = 0; i < N_FILAS && out.length < cuantas; i++) {
    var r = datos[i];
    if (r[COL.CONCEPTO - 1] !== '' || r[COL.VEHICULO - 1] !== '') continue;
    var hayDias = false;
    for (var c = COL.DIA_START - 1; c < COL.DIA_END; c++) { if (r[c] !== '') { hayDias = true; break; } }
    if (!hayDias) out.push(DATA_START + i);
  }
  return out;
}

// Guarda la calificación de un servicio ya realizado (chofer y unidad, 1 a 5 estrellas)
function registrarCalificacion(datos) {
  datos = datos || {};
  var lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var cal = ss.getSheetByName('Calificaciones');
    if (!cal) {
      cal = ss.insertSheet('Calificaciones');
      cal.getRange(1, 1, 1, 8).setValues([[
        'Fecha de registro', 'Fecha del servicio', 'Proyecto', 'Chofer (como lo escribió el cliente)',
        'Placas (como las escribió el cliente)', 'Calificación Chofer (1-5)', 'Calificación Unidad (1-5)', 'Comentarios'
      ]]).setFontWeight('bold').setBackground('#191C16').setFontColor('#FFFFFF');
      cal.setFrozenRows(1);
    }
    cal.appendRow([
      new Date(), datos.fechaServicio || '', datos.proyecto || '', datos.chofer || '',
      datos.placas || '', Number(datos.calificacionChofer) || '', Number(datos.calificacionUnidad) || '', datos.comentarios || ''
    ]);
    return { ok: true, mensaje: '¡Gracias por tu calificación!' };
  } catch (err) {
    return { ok: false, mensaje: 'Error al guardar: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}
