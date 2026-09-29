/**
 * BASE DE DATOS MAESTRA DE TRANSPORTACIÓN — CINEMA SHUTTLE
 * ---------------------------------------------------------
 * Pega este código en el Apps Script del archivo "Tablero Analítico Transportación".
 *
 * Qué hace:
 *  - Recorre las carpetas de proyectos (FOLDER_IDS) y abre cada Google Sheet.
 *  - De cada archivo lee SOLO la pestaña "Solicitudes Transporte" (ignora
 *    Catálogos, Crew List, Estimate, Cierre, tablas dinámicas, etc.).
 *  - De esa pestaña toma SOLO las filas que de verdad tienen una solicitud
 *    (vehículo + algún dato: concepto, proveedor, chofer, placas, días o costo).
 *  - Funciona con el machote nuevo Y con los formatos anteriores
 *    (DIA 13, COSTO FINAL, PAGO PROVEEDOR, etc.): busca las columnas por nombre.
 *  - Es INCREMENTAL: solo vuelve a leer los archivos que cambiaron desde la
 *    última vez. Para agregar proyectos anteriores basta con meterlos a la
 *    carpeta y presionar "Consolidar ahora".
 *  - Si hay muchos archivos, trabaja por lotes y se reprograma solo hasta
 *    terminar (límite de 6 minutos de Apps Script).
 *
 * Pestañas que crea en ESTE archivo:
 *  - Panel               → botón/instrucciones y estado de la última corrida
 *  - BD Solicitudes      → una fila por unidad solicitada (la base para reportes)
 *  - BD Uso Diario       → una fila por unidad por día de uso (código del calendario)
 *  - Registro Archivos   → qué archivos se leyeron, cuándo y con qué resultado
 *  - Alias               → (opcional) unifica nombres de proveedores/choferes
 *  - Resumen por Proyecto, KPI … → tablas resumen (las usa el Dashboard)
 *  - Log                 → historial de corridas
 *
 * NUNCA escribe en los archivos de los proyectos: solo los lee.
 */

// ====================== CONFIGURACIÓN ======================
// Carpetas donde están los archivos de los proyectos. Puedes poner varias.
const FOLDER_IDS = [
  '1AZ3XSitDHx5Mc3JiUnID157gADXytLx8', // PROYECTOS TRANSPO 2026 (nuevos, con el machote)
  '1HYmV5WFMfHg-gjP6Qob739hXB3H1C8Fm', // PROYECTOS 2026 (proyectos anteriores)
];
const INCLUIR_SUBCARPETAS = true;

// Pestaña que se lee en cada archivo (sin importar mayúsculas/acentos/espacios extra)
const HOJA_ORIGEN = 'Solicitudes Transporte';

// Deja vacío '' para leer cualquier Google Sheet de la carpeta, o pon un texto
// (ej. 'TRANSPO REQUEST') para leer solo archivos cuyo nombre lo contenga.
const FILTRO_NOMBRE = '';
const EXCLUIR_COPIAS = true; // ignora "Copia de…" / "Copy of…"

const MAX_MS_POR_LOTE = 4.5 * 60 * 1000; // margen bajo el límite de 6 min

const H = {
  PANEL: 'Panel',
  BD: 'BD Solicitudes',
  DIARIO: 'BD Uso Diario',
  REGISTRO: 'Registro Archivos',
  ALIAS: 'Alias',
  RESUMEN_PROY: 'Resumen por Proyecto',
  LOG: 'Log',
};

// ====================== CAMPOS DE LA BASE ======================
// key = nombre interno · titulo = encabezado en "BD Solicitudes"
// alias = cómo puede venir el encabezado en los archivos (ya normalizado:
// MAYÚSCULAS, sin acentos, sin signos). Se busca coincidencia EXACTA.
const CAMPOS = [
  { key: 'vehiculo', titulo: 'Vehículo', alias: ['VEHICULO'] },
  { key: 'concepto', titulo: 'Concepto', alias: ['CONCEPTO', 'CONCEPTO USO'] },
  { key: 'idioma', titulo: 'Idioma / Bilingüe', alias: ['IDIOMA CHOFER', 'CHOFER BILINGUE', 'BILINGUE'] },
  { key: 'requisitos', titulo: 'Requisitos unidad', alias: ['REQUISITOS UNIDAD', 'REQUISITOS'] },
  { key: 'paxSolicitado', titulo: 'Pax solicitado', alias: ['CANTIDAD DE PASAJEROS', 'PAX SOLICITADO'] },
  { key: 'proveedor', titulo: 'Proveedor (original)', alias: ['PROVEEDOR'] },
  { key: 'chofer', titulo: 'Chofer (original)', alias: ['NOMBRE CHOFER', 'NOMBRE DEL CHOFER', 'CHOFER', 'DRIVER'] },
  { key: 'cel', titulo: 'Cel. chofer', alias: ['CEL CHOFER', 'CELULAR CHOFER', 'TELEFONO CHOFER'] },
  { key: 'placas', titulo: 'Placas', alias: ['PLACAS DE LA UNIDAD', 'PLACAS'] },
  { key: 'anio', titulo: 'Año unidad', alias: ['ANO DE LA UNIDAD', 'ANO'] },
  { key: 'paxUnidad', titulo: 'Pax unidad', alias: ['PAX'] },
  { key: 'tarifaUnidad', titulo: 'Tarifa unidad', alias: ['TARIFA UNIDAD', 'TARIFA DIARIA'], dinero: true },
  { key: 'cinemaDia', titulo: 'Cinema por día', alias: [], dinero: true },            // 1er "CINEMA"
  { key: 'ajusteBilingue', titulo: 'Ajuste bilingüe', alias: ['AJUSTE DE COSTO BILINGUE'], dinero: true },
  { key: 'costoFinalDia', titulo: 'Costo final por día', alias: ['COSTO FINAL'], dinero: true },
  { key: 'tarifaForanea', titulo: 'Tarifa foránea', alias: ['TARIFA FORANEA POR DIA', 'TARIFA FORANEA'], dinero: true },
  { key: 'tarifaTransfer', titulo: 'Tarifa transfer', alias: ['TARIFA TRANSFER'], dinero: true },
  { key: 'diasFull', titulo: 'Días full', alias: ['DIAS COMPLETOS', 'DIAS FULL'], numero: true },
  { key: 'diasForaneo', titulo: 'Días foráneo', alias: ['DIAS FORANEO', 'DIAS FORANEOS'], numero: true },
  { key: 'diasTransfer', titulo: 'Días transfer', alias: ['DIAS TRANSFER'], numero: true },
  { key: 'diasCancelados', titulo: 'Días cancelados', alias: ['DIAS CANCELADOS'], numero: true },
  { key: 'costoCdmx', titulo: 'Costo total CDMX', alias: ['COSTO TOTAL CDMX'], dinero: true },
  { key: 'costoForaneo', titulo: 'Costo total foráneo', alias: ['COSTO TOTAL FORANEO'], dinero: true },
  { key: 'costoTransfer', titulo: 'Costo total transfer', alias: ['COSTO TOTAL TRANSFER'], dinero: true },
  { key: 'horasExtra', titulo: 'Horas extra', alias: ['HORAS EXTRA', 'HORAS EXTRAS'], numero: true },
  { key: 'tarifaHoraExtra', titulo: '$ por hora extra', alias: ['HORA EXTRA'], dinero: true },
  { key: 'costoHorasExtra', titulo: 'Costo horas extra', alias: ['COSTO DE HORAS EXTRAS', 'COSTO HORAS EXTRA', 'COSTO HORAS EXTRAS'], dinero: true },
  { key: 'costoDiasDobles', titulo: 'Costo días dobles', alias: ['COSTO DIAS DOBLES'], dinero: true },
  { key: 'viaticos', titulo: 'Viáticos extra', alias: ['VIATICOS EXTRAS', 'VIATICOS EXTRA', 'VIATICOS'], dinero: true },
  { key: 'cinemaTotal', titulo: 'Cinema total', alias: [], dinero: true },            // 2º "CINEMA"
  { key: 'pagoProveedor', titulo: 'Pago proveedor', alias: ['PAGO PROVEEDOR'], dinero: true },
  { key: 'costoTotal', titulo: 'Costo total estimado', alias: ['COSTO TOTAL ESTIMADO', 'COSTO TOTAL'], dinero: true },
  { key: 'linea', titulo: 'Línea', alias: ['LINEA'] },
  { key: 'razonSocial', titulo: 'Razón social', alias: ['RAZON SOCIAL CONSOLIDADA', 'RAZON SOCIAL'] },
  { key: 'validUnidad', titulo: 'Validación unidad', alias: ['VALIDACION UNIDAD'] },
  { key: 'validDriver', titulo: 'Validación driver', alias: ['VALIDACION DRIVER'] },
  { key: 'estatus', titulo: 'Estatus', alias: ['ESTATUS'] },
  { key: 'incidencia', titulo: 'Incidencia', alias: ['INCIDENCIA', 'INCIDENCIAS'] },
  { key: 'locaciones', titulo: 'Locaciones', alias: ['LOCACIONES'] },
  { key: 'observaciones', titulo: 'Observaciones', alias: ['OBSERVACIONES'] },
  { key: 'obsTranspo', titulo: 'Observaciones transportación', alias: ['OBSERVACIONES TRANSPORTACION'] },
];

// Columnas de "BD Solicitudes" (en este orden)
const ENC_META_INICIO = ['ID Renglón', 'Proyecto', 'Productor', 'Cliente', 'Capitán transportación', 'Especialista', 'Mes', 'Fecha inicio'];
const ENC_DERIVADOS = ['Proveedor', 'Chofer', 'Días de servicio', 'Primer día', 'Último día'];
const ENC_META_FIN = ['Archivo', 'Fila origen', 'URL archivo', 'ID Archivo', 'Consolidado el'];
const ENC_BD = ENC_META_INICIO.concat(ENC_DERIVADOS, CAMPOS.map(function (c) { return c.titulo; }), ENC_META_FIN);
const IDX_BD_ID_ARCHIVO = ENC_BD.indexOf('ID Archivo');

const ENC_DIARIO = ['ID Renglón', 'Proyecto', 'Fecha', 'Día (encabezado)', 'Código', 'Modalidad', 'Estatus día', 'Vehículo', 'Proveedor', 'Chofer', 'Placas', 'ID Archivo'];
const IDX_DIARIO_ID_ARCHIVO = ENC_DIARIO.indexOf('ID Archivo');

const ENC_REGISTRO = ['ID Archivo', 'Archivo', 'Proyecto detectado', 'Estado', 'Filas', 'Días de uso', 'Modificado en Drive', 'Leído el', 'URL'];

// Etiquetas del encabezado del proyecto (arriba de la tabla)
const ETIQUETAS_META = {
  'PROYECTO': 'proyecto', 'NOMBRE DEL PROYECTO': 'proyecto',
  'PRODUCTOR': 'productor', 'PRODUCTORA': 'productor',
  'CLIENTE': 'cliente', 'CLIENTE USUARIO': 'cliente', 'USUARIO': 'cliente',
  'CAPITAN DE TRANSPORTACION': 'capitan', 'CAPITAN DE TRANSPORTE': 'capitan',
  'ESPECIALISTA DE TRANSPORTE': 'especialista', 'ESPECIALISTA': 'especialista',
  'MES': 'mes', 'FECHA DE INICIO': 'fechaInicio', 'FECHA INICIO': 'fechaInicio',
};

// Si en una fila aparece alguno de estos textos, ahí termina la tabla de solicitudes
// (debajo suele haber resúmenes / tablas dinámicas que NO son solicitudes).
const MARCAS_FIN = ['PROVEEDOR', 'SUMA TOTAL', 'GRAN TOTAL', 'COSTO TOTAL DEL PROYECTO', 'GASTO TOTAL DEL PROYECTO', 'PROCESO DE CIERRE', 'VEHICULOS', 'MATERIALES'];

// ====================== MENÚ ======================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Base Maestra Transporte')
    .addItem('▶ Consolidar ahora (solo lo nuevo o modificado)', 'consolidarAhora')
    .addItem('↻ Reprocesar TODO desde cero', 'reprocesarTodo')
    .addSeparator()
    .addItem('Activar actualización automática (cada hora)', 'activarAutomatico')
    .addItem('Desactivar actualización automática', 'desactivarAutomatico')
    .addSeparator()
    .addItem('Crear / reparar Panel y hoja de Alias', 'prepararPanel')
    .addToUi();
}

// ====================== PUNTOS DE ENTRADA ======================
// Botón principal (asígnalo a un dibujo en la hoja Panel: "consolidarAhora")
function consolidarAhora() {
  ejecutar_(false, true);
}

function reprocesarTodo() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.alert('Se volverán a leer TODOS los archivos de la carpeta (puede tardar varios minutos y seguir en segundo plano). ¿Continuar?', ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;
  ejecutar_(true, true);
}

// La llama el trigger de continuación cuando quedaron archivos pendientes
function continuarConsolidacion() {
  borrarTriggers_('continuarConsolidacion');
  ejecutar_(false, false);
}

// La llama el trigger horario
function consolidacionAutomatica() {
  ejecutar_(false, false);
}

function activarAutomatico() {
  borrarTriggers_('consolidacionAutomatica');
  ScriptApp.newTrigger('consolidacionAutomatica').timeBased().everyHours(1).create();
  avisar_('Listo. La base se actualizará sola cada hora (solo lee archivos que cambiaron).');
}

function desactivarAutomatico() {
  borrarTriggers_('consolidacionAutomatica');
  borrarTriggers_('continuarConsolidacion');
  avisar_('Actualización automática desactivada.');
}

function borrarTriggers_(handler) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === handler) ScriptApp.deleteTrigger(t);
  });
}

// ====================== MOTOR PRINCIPAL ======================
function ejecutar_(forzarTodo, interactivo) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    if (interactivo) avisar_('Ya hay una consolidación en curso. Espera a que termine (mira la hoja Panel).');
    return;
  }
  var inicio = Date.now();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    prepararHojas_(ss);
    estadoPanel_(ss, '⏳ Consolidando…', '');
    var tz = ss.getSpreadsheetTimeZone();
    var alias = leerAlias_(ss);

    // 1) Qué hay en las carpetas
    var archivos = listarArchivos_(ss.getId());
    var porId = {};
    archivos.forEach(function (a) { porId[a.id] = a; });

    // 2) Qué ya habíamos leído
    var registro = leerRegistro_(ss); // id -> fila del registro (array)
    if (forzarTodo) Object.keys(registro).forEach(function (id) { registro[id][6] = ''; });

    // 3) Datos actuales de la base agrupados por archivo
    var bdPorArchivo = agruparPorArchivo_(ss.getSheetByName(H.BD), ENC_BD.length, IDX_BD_ID_ARCHIVO);
    var diarioPorArchivo = agruparPorArchivo_(ss.getSheetByName(H.DIARIO), ENC_DIARIO.length, IDX_DIARIO_ID_ARCHIVO);

    // 4) Archivos que ya no están en la carpeta → se quitan de la base
    var quitados = 0;
    Object.keys(registro).forEach(function (id) {
      if (!porId[id]) { delete registro[id]; delete bdPorArchivo[id]; delete diarioPorArchivo[id]; quitados++; }
    });

    // 5) Pendientes = nuevos o modificados
    var pendientes = archivos.filter(function (a) {
      var reg = registro[a.id];
      return !reg || String(reg[6]) !== a.modificado;
    });
    var totalPendientes = pendientes.length;

    var procesados = 0, errores = [];
    var ahora = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
    for (var i = 0; i < pendientes.length; i++) {
      if (Date.now() - inicio > MAX_MS_POR_LOTE) break;
      var a = pendientes[i];
      var res;
      try {
        res = leerArchivo_(a, tz, alias, ahora);
      } catch (e) {
        res = { estado: 'ERROR: ' + e.message, filas: [], dias: [], proyecto: '' };
        errores.push(a.nombre + ' (' + e.message + ')');
      }
      if (res.estado.indexOf('ERROR') === 0) {
        // No tocamos los datos anteriores de ese archivo y lo dejamos
        // "sin fecha" para que se reintente en la siguiente corrida.
        registro[a.id] = [a.id, a.nombre, (registro[a.id] || [])[2] || '', res.estado, (registro[a.id] || [])[4] || 0, (registro[a.id] || [])[5] || 0, '', ahora, a.url];
      } else {
        bdPorArchivo[a.id] = res.filas;
        diarioPorArchivo[a.id] = res.dias;
        registro[a.id] = [a.id, a.nombre, res.proyecto, res.estado, res.filas.length, res.dias.length, a.modificado, ahora, a.url];
      }
      procesados++;
    }
    var faltan = totalPendientes - procesados;

    // 6) Escribir todo (orden: archivos por nombre)
    var orden = archivos.map(function (a) { return a.id; });
    var filasBD = aplanar_(bdPorArchivo, orden);
    var filasDiario = aplanar_(diarioPorArchivo, orden);
    escribirTabla_(ss.getSheetByName(H.BD), ENC_BD, filasBD);
    escribirTabla_(ss.getSheetByName(H.DIARIO), ENC_DIARIO, filasDiario);
    escribirRegistro_(ss, registro, orden);
    formatearBD_(ss);

    // 7) Resúmenes / KPIs (siempre sobre la base completa)
    calcularResumenes_(ss, filasBD, ahora);

    // 8) ¿Faltó tiempo? → nos reprogramamos en 1 minuto
    var estado;
    if (faltan > 0) {
      borrarTriggers_('continuarConsolidacion');
      ScriptApp.newTrigger('continuarConsolidacion').timeBased().after(60 * 1000).create();
      estado = '⏳ En proceso: faltan ' + faltan + ' archivo(s). Continúa solo en ~1 minuto.';
    } else {
      estado = '✅ Al día';
    }
    var detalle = 'Archivos en carpeta: ' + archivos.length + ' · Leídos en esta corrida: ' + procesados
      + (quitados ? ' · Quitados: ' + quitados : '')
      + ' · Filas en la base: ' + filasBD.length
      + (errores.length ? ' · Con error: ' + errores.length : '');
    estadoPanel_(ss, estado, detalle);
    ss.getSheetByName(H.LOG).appendRow([new Date(), estado, detalle, errores.join(' | ')]);

    if (interactivo) {
      ss.toast(detalle, estado, 10);
      if (faltan > 0) avisar_(estado + '\n\n' + detalle + '\n\nPuedes cerrar el archivo: el proceso sigue en segundo plano. Revisa la hoja Panel.');
    }
  } catch (e) {
    estadoPanel_(ss, '❌ Error', e.message);
    ss.getSheetByName(H.LOG).appendRow([new Date(), '❌ Error', e.message, e.stack || '']);
    if (interactivo) avisar_('Ocurrió un error: ' + e.message);
    else throw e;
  } finally {
    lock.releaseLock();
  }
}

// ====================== LECTURA DE CARPETAS ======================
function listarArchivos_(idPropio) {
  var out = [], vistos = {};
  function recorrer(folder, ruta) {
    var it = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    while (it.hasNext()) {
      var f = it.next();
      var id = f.getId();
      if (id === idPropio || vistos[id]) continue;
      var nombre = f.getName();
      if (FILTRO_NOMBRE && nombre.toUpperCase().indexOf(FILTRO_NOMBRE.toUpperCase()) === -1) continue;
      if (EXCLUIR_COPIAS && esCopia_(nombre)) continue;
      vistos[id] = true;
      out.push({ id: id, nombre: nombre, url: f.getUrl(), carpeta: ruta, modificado: String(f.getLastUpdated().getTime()) });
    }
    if (INCLUIR_SUBCARPETAS) {
      var sub = folder.getFolders();
      while (sub.hasNext()) { var s = sub.next(); recorrer(s, ruta + ' / ' + s.getName()); }
    }
  }
  FOLDER_IDS.forEach(function (fid) {
    var folder = DriveApp.getFolderById(fid);
    recorrer(folder, folder.getName());
  });
  out.sort(function (a, b) { return a.nombre.localeCompare(b.nombre); });
  return out;
}

function esCopia_(nombre) {
  var n = nombre.trim().toUpperCase();
  return n.indexOf('COPY OF') === 0 || n.indexOf('COPIA DE') === 0;
}

// ====================== LECTURA DE UN ARCHIVO ======================
function leerArchivo_(archivo, tz, alias, ahora) {
  var libro = abrirConReintento_(archivo.id);
  var hoja = buscarHoja_(libro, HOJA_ORIGEN);
  if (!hoja) return { estado: 'SIN PESTAÑA "' + HOJA_ORIGEN + '"', filas: [], dias: [], proyecto: '' };

  var valores = hoja.getDataRange().getValues();
  var r = extraerSolicitudes(valores, { nombreArchivo: archivo.nombre, idArchivo: archivo.id, tz: tz, alias: alias });
  if (r.error) return { estado: r.error, filas: [], dias: [], proyecto: r.meta ? r.meta.proyecto : '' };

  var filas = r.filas.map(function (f) { return filaBD_(f, r.meta, archivo, ahora); });
  var dias = [];
  r.filas.forEach(function (f) {
    f.dias.forEach(function (d) {
      dias.push([f.idRenglon, r.meta.proyecto, d.fecha || '', d.encabezado, d.codigo, d.modalidad, d.estatus,
                 f.valores.vehiculo, f.proveedorNorm, f.choferNorm, f.valores.placas, archivo.id]);
    });
  });
  return { estado: filas.length ? 'OK' : 'OK (sin solicitudes capturadas)', filas: filas, dias: dias, proyecto: r.meta.proyecto };
}

function abrirConReintento_(id) {
  for (var intento = 1; ; intento++) {
    try { return SpreadsheetApp.openById(id); }
    catch (e) {
      if (intento >= 3) throw e;
      Utilities.sleep(1500 * intento);
    }
  }
}

function buscarHoja_(libro, nombre) {
  var exacta = libro.getSheetByName(nombre);
  if (exacta) return exacta;
  var objetivo = normalizar(nombre);
  var hojas = libro.getSheets();
  for (var i = 0; i < hojas.length; i++) {
    if (normalizar(hojas[i].getName()) === objetivo) return hojas[i];
  }
  return null;
}

function filaBD_(f, meta, archivo, ahora) {
  var v = f.valores;
  var fila = [f.idRenglon, meta.proyecto, meta.productor, meta.cliente, meta.capitan, meta.especialista, meta.mes, meta.fechaInicio || '',
              f.proveedorNorm, f.choferNorm, f.diasServicio, f.primerDia || '', f.ultimoDia || ''];
  CAMPOS.forEach(function (c) { fila.push(v[c.key] === undefined ? '' : v[c.key]); });
  fila.push(archivo.nombre, f.filaOrigen, archivo.url, archivo.id, ahora);
  return fila;
}

// =====================================================================
//  PARSER (funciones puras: no usan servicios de Google, se pueden probar)
// =====================================================================

// Normaliza textos para comparar: MAYÚSCULAS, sin acentos, sin signos, espacios simples
function normalizar(v) {
  return String(v === null || v === undefined ? '' : v)
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, ' ')
    .trim();
}

function esFecha_(v) { return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime()); }

function esVacio_(v) {
  if (v === '' || v === null || v === undefined) return true;
  var s = String(v).trim();
  return s === '' || s.charAt(0) === '#'; // #N/A, #REF!, #VALUE!…
}

function limpio_(v) {
  if (esVacio_(v)) return '';
  if (typeof v === 'string') return v.replace(/\s+/g, ' ').trim();
  return v;
}

function numero_(v) {
  if (typeof v === 'number') return v;
  if (esVacio_(v)) return 0;
  var n = parseFloat(String(v).replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}

function booleano_(v) {
  if (v === true || v === false) return v;
  var s = normalizar(v);
  if (s === 'TRUE' || s === 'VERDADERO' || s === 'SI') return true;
  if (s === 'FALSE' || s === 'FALSO' || s === 'NO') return false;
  return '';
}

function esEncabezadoDia_(v) {
  if (esFecha_(v)) return true;
  return /^DIA \d{1,2}$/.test(normalizar(v));
}

const MESES_ = { ENE: 1, FEB: 2, MAR: 3, ABR: 4, MAY: 5, JUN: 6, JUL: 7, AGO: 8, SEP: 9, SET: 9, OCT: 10, NOV: 11, DIC: 12,
                 JAN: 1, APR: 4, AUG: 8, DEC: 12 };

// Convierte a 'yyyy-MM-dd' un encabezado de día: Date, o texto tipo "LUNES 28 sept"
function fechaDeDia_(valor, anio, tz) {
  if (esFecha_(valor)) {
    if (valor.getFullYear() < 2000) return '';
    return typeof Utilities !== 'undefined' && tz
      ? Utilities.formatDate(valor, tz, 'yyyy-MM-dd')
      : valor.getFullYear() + '-' + ('0' + (valor.getMonth() + 1)).slice(-2) + '-' + ('0' + valor.getDate()).slice(-2);
  }
  var m = normalizar(valor).match(/(\d{1,2}) (?:DE )?([A-Z]{3})/);
  if (!m || !anio) return '';
  var mes = MESES_[m[2]];
  if (!mes) return '';
  return anio + '-' + ('0' + mes).slice(-2) + '-' + ('0' + Number(m[1])).slice(-2);
}

// Nombre de proveedor legible y comparable:
// "PROV -  CHRISTIAN ALEXIS ANGEL" / "CHRISTIAN ANGEL: Christian Alexis… RFC…" → "CHRISTIAN ALEXIS ANGEL" / "CHRISTIAN ANGEL"
function normalizarProveedor(v, alias) {
  if (esVacio_(v)) return '';
  var s = String(v).replace(/\s+/g, ' ').trim();
  s = s.replace(/^PROV(EEDOR)?\s*[A-Z]?\s*[-–:]\s*/i, '');
  s = s.split(/[:,;]| RFC| R\.F\.C/i)[0];
  var tokens = s.split(' ');
  for (var i = 0; i < tokens.length; i++) {
    if (tokens[i].indexOf('@') > -1) { tokens = tokens.slice(0, i); break; } // corta en el correo
  }
  s = tokens.join(' ');
  var n = normalizar(s);
  return (alias && alias.PROVEEDOR && alias.PROVEEDOR[n]) || n;
}

function normalizarChofer(v, alias) {
  if (esVacio_(v)) return '';
  var s = String(v).replace(/^\s*(DRIVER|CHOFER)\s*:\s*/i, '');
  s = s.split(/\n| PARAM[EÉ]DICO/i)[0];
  var n = normalizar(s);
  if (/^\d+$/.test(n)) return ''; // números sueltos no son nombres
  return (alias && alias.CHOFER && alias.CHOFER[n]) || n;
}

function modalidadDe_(codigo) {
  var c = normalizar(codigo).charAt(0);
  return { '1': 'FULL', 'F': 'FORANEO', 'T': 'TRANSFER', 'D': 'DOBLE', 'M': 'MEDIO DIA' }[c] || 'OTRO';
}
function estatusDe_(codigo) {
  var s = normalizar(codigo);
  if (s.length < 2) return '';
  return { 'S': 'SOLICITADO', 'C': 'CONFIRMADO', 'P': 'TENTATIVO', 'X': 'CANCELADO' }[s.charAt(s.length - 1)] || '';
}

function limpiarNombreProyecto(nombreArchivo) {
  return String(nombreArchivo)
    .replace(/TRANSPO\s*REQUEST\s*(20\d\d)?/i, '')
    .replace(/^[\d.\-\s]+/, '')
    .replace(/[\s\-–_]+$/, '')
    .replace(/^[\s\-–_]+/, '')
    .trim();
}

/**
 * Recibe los valores de la pestaña (getValues) y devuelve:
 *   { meta: {proyecto, productor, …}, filas: [{valores, dias, …}], error }
 */
function extraerSolicitudes(valores, opts) {
  opts = opts || {};
  var norm = valores.map(function (fila) { return fila.map(normalizar); });

  // 1) fila de encabezados = primera que tenga VEHICULO y PROVEEDOR
  var hdr = -1;
  for (var i = 0; i < norm.length; i++) {
    if (norm[i].indexOf('VEHICULO') > -1 && norm[i].indexOf('PROVEEDOR') > -1) { hdr = i; break; }
  }
  var meta = leerMeta_(valores, norm, hdr === -1 ? Math.min(norm.length, 10) : hdr, opts);
  if (hdr === -1) return { meta: meta, filas: [], error: 'SIN ENCABEZADOS (no encontré VEHICULO + PROVEEDOR)' };

  // 2) mapa de columnas
  var enc = norm[hdr];
  var col = {};
  CAMPOS.forEach(function (c) {
    for (var k = 0; k < c.alias.length; k++) {
      var p = enc.indexOf(c.alias[k]);
      if (p > -1) { col[c.key] = p; return; }
    }
  });
  var cinemas = [];
  enc.forEach(function (h, j) { if (h === 'CINEMA') cinemas.push(j); });
  if (cinemas.length) col.cinemaDia = cinemas[0];
  if (cinemas.length > 1) col.cinemaTotal = cinemas[cinemas.length - 1];

  // columnas del calendario
  var anio = anioReferencia_(meta, opts.nombreArchivo);
  var colsDia = [];
  valores[hdr].forEach(function (h, j) {
    if (!esEncabezadoDia_(h)) return;
    var fecha = fechaDeDia_(h, anio, opts.tz);
    // formato viejo: "DIA 13" abajo y "LUNES 28 sept" (o una fecha) en la fila de arriba
    if (!fecha && hdr > 0) fecha = fechaDeDia_(valores[hdr - 1][j], anio, opts.tz);
    colsDia.push({ j: j, encabezado: esFecha_(h) ? (fecha || 'DIA ' + (colsDia.length + 1)) : String(h).trim(), fecha: fecha });
  });

  // 3) filas de datos
  var filas = [];
  for (var r = hdr + 1; r < valores.length; r++) {
    var fn = norm[r];
    if (fn.indexOf('VEHICULO') > -1 && fn.indexOf('PROVEEDOR') > -1) break;               // otra tabla
    if (MARCAS_FIN.some(function (m) { return fn.indexOf(m) > -1; })) break;                // resúmenes

    var fila = valores[r];
    var vehiculo = limpio_(fila[col.vehiculo]);
    if (!vehiculo || typeof vehiculo !== 'string' || vehiculo.length > 40) continue;

    var v = {};
    CAMPOS.forEach(function (c) {
      if (col[c.key] === undefined) { v[c.key] = ''; return; }
      var x = fila[col[c.key]];
      if (c.dinero || c.numero) v[c.key] = esVacio_(x) ? '' : numero_(x);
      else if (c.key === 'validUnidad' || c.key === 'validDriver') { var b = booleano_(x); v[c.key] = b === '' ? limpio_(x) : b; }
      else v[c.key] = esFecha_(x) ? x : limpio_(x);
    });

    // calendario de esta fila
    var dias = [];
    colsDia.forEach(function (cd) {
      var cod = limpio_(fila[cd.j]);
      if (cod === '' || cod === 0) return;
      cod = String(cod).toUpperCase();
      dias.push({ encabezado: cd.encabezado, fecha: cd.fecha, codigo: cod, modalidad: modalidadDe_(cod), estatus: estatusDe_(cod) });
    });
    var activos = dias.filter(function (d) { return d.estatus !== 'CANCELADO'; });

    // ¿es una solicitud real? vehículo + al menos un dato más
    var costo = numero_(v.costoTotal) || numero_(v.pagoProveedor) || numero_(v.costoFinalDia);
    var tieneAlgo = v.concepto || v.proveedor || v.chofer || v.placas || dias.length || costo > 0;
    if (!tieneAlgo) continue;

    // días por tipo: si el archivo no trae esas columnas, se calculan del calendario
    function contar(mod) { return activos.filter(function (d) { return d.modalidad === mod; }).length; }
    if (col.diasFull === undefined) v.diasFull = contar('FULL');
    if (col.diasForaneo === undefined) v.diasForaneo = contar('FORANEO');
    if (col.diasTransfer === undefined) v.diasTransfer = contar('TRANSFER');
    if (col.diasCancelados === undefined) v.diasCancelados = dias.length - activos.length;
    if (col.costoHorasExtra === undefined && v.horasExtra !== '' && v.tarifaHoraExtra !== '') {
      v.costoHorasExtra = numero_(v.horasExtra) * numero_(v.tarifaHoraExtra);
    }

    var fechas = activos.map(function (d) { return d.fecha; }).filter(function (x) { return x; }).sort();
    var diasServicio = numero_(v.diasFull) + numero_(v.diasForaneo) + numero_(v.diasTransfer);
    if (!diasServicio) diasServicio = activos.length;

    filas.push({
      idRenglon: (opts.idArchivo || 'archivo') + '#' + (r + 1),
      filaOrigen: r + 1,
      valores: v,
      dias: dias,
      diasServicio: diasServicio,
      primerDia: fechas[0] || '',
      ultimoDia: fechas[fechas.length - 1] || '',
      proveedorNorm: normalizarProveedor(v.proveedor, opts.alias),
      choferNorm: normalizarChofer(v.chofer, opts.alias),
    });
  }
  return { meta: meta, filas: filas };
}

function anioReferencia_(meta, nombreArchivo) {
  var m = String(meta.fechaInicio || '').match(/^(\d{4})/);
  if (m) return Number(m[1]);
  m = String(nombreArchivo || '').match(/20\d\d/);
  if (m) return Number(m[0]);
  return new Date().getFullYear();
}

// Busca etiquetas (Proyecto, Productor, Mes…) arriba de la tabla y toma
// el/los valor(es) a su derecha hasta la siguiente etiqueta.
function leerMeta_(valores, norm, hastaFila, opts) {
  var meta = { proyecto: '', productor: '', cliente: '', capitan: '', especialista: '', mes: '', fechaInicio: '' };
  for (var r = 0; r < hastaFila; r++) {
    for (var c = 0; c < norm[r].length; c++) {
      var campo = ETIQUETAS_META[norm[r][c]];
      if (!campo || meta[campo]) continue;
      var partes = [];
      for (var k = c + 1; k < norm[r].length; k++) {
        if (ETIQUETAS_META[norm[r][k]]) break;
        var x = valores[r][k];
        if (esVacio_(x)) continue;
        if (typeof x === 'string' && x.trim().charAt(0) === '(') continue; // notas de ayuda
        if (campo === 'fechaInicio') {
          partes.push(esFecha_(x) ? fechaDeDia_(x, null, opts.tz) : String(x).trim());
          break;
        }
        partes.push(esFecha_(x) ? fechaDeDia_(x, null, opts.tz) : String(x).replace(/\s+/g, ' ').trim());
        if (partes.length >= 3) break;
      }
      meta[campo] = partes.join(' / ');
    }
  }
  var proyectoArchivo = limpiarNombreProyecto(opts.nombreArchivo || '');
  if (!meta.proyecto || normalizar(meta.proyecto) === 'NOMBRE PROYECTO') meta.proyecto = proyectoArchivo;
  if (!meta.mes && meta.fechaInicio) {
    var mm = String(meta.fechaInicio).match(/^\d{4}-(\d{2})/);
    if (mm) meta.mes = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'][Number(mm[1]) - 1];
  }
  if (meta.mes) meta.mes = String(meta.mes).toUpperCase();
  return meta;
}

// =====================================================================
//  ESCRITURA EN ESTE ARCHIVO
// =====================================================================
function prepararHojas_(ss) {
  function asegurar(nombre, encabezado) {
    var h = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
    if (encabezado && h.getLastRow() === 0) {
      h.getRange(1, 1, 1, encabezado.length).setValues([encabezado]);
    }
    return h;
  }
  if (!ss.getSheetByName(H.PANEL)) prepararPanel();
  asegurar(H.BD, ENC_BD);
  asegurar(H.DIARIO, ENC_DIARIO);
  asegurar(H.REGISTRO, ENC_REGISTRO);
  asegurar(H.LOG, ['Fecha', 'Estado', 'Detalle', 'Errores']);
  if (!ss.getSheetByName(H.ALIAS)) crearHojaAlias_(ss);

  // Si cambió la estructura de columnas (nueva versión del script), reprocesa todo
  var bd = ss.getSheetByName(H.BD);
  var encActual = bd.getRange(1, 1, 1, Math.max(1, bd.getLastColumn())).getValues()[0].join('|');
  if (encActual !== ENC_BD.join('|')) {
    bd.clear();
    bd.getRange(1, 1, 1, ENC_BD.length).setValues([ENC_BD]);
    ss.getSheetByName(H.DIARIO).clear();
    ss.getSheetByName(H.DIARIO).getRange(1, 1, 1, ENC_DIARIO.length).setValues([ENC_DIARIO]);
    var reg = ss.getSheetByName(H.REGISTRO);
    reg.clear();
    reg.getRange(1, 1, 1, ENC_REGISTRO.length).setValues([ENC_REGISTRO]);
  }
}

function agruparPorArchivo_(hoja, ancho, idxId) {
  var out = {};
  if (!hoja || hoja.getLastRow() < 2) return out;
  var datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, ancho).getValues();
  datos.forEach(function (f) {
    var id = f[idxId];
    if (!id) return;
    (out[id] = out[id] || []).push(f);
  });
  return out;
}

function aplanar_(porArchivo, orden) {
  var out = [];
  orden.forEach(function (id) { if (porArchivo[id]) out = out.concat(porArchivo[id]); });
  return out;
}

function escribirTabla_(hoja, encabezado, filas) {
  hoja.clearContents();
  hoja.getRange(1, 1, 1, encabezado.length).setValues([encabezado])
    .setFontWeight('bold').setBackground('#191C16').setFontColor('#FFFFFF');
  hoja.setFrozenRows(1);
  if (filas.length) {
    // asegurar espacio suficiente
    var necesarias = filas.length + 1;
    if (hoja.getMaxRows() < necesarias) hoja.insertRowsAfter(hoja.getMaxRows(), necesarias - hoja.getMaxRows());
    if (hoja.getMaxColumns() < encabezado.length) hoja.insertColumnsAfter(hoja.getMaxColumns(), encabezado.length - hoja.getMaxColumns());
    hoja.getRange(2, 1, filas.length, encabezado.length).setValues(filas);
  }
  // quita filas sobrantes muy grandes para no inflar el archivo
  var sobra = hoja.getMaxRows() - (filas.length + 1);
  if (sobra > 500) hoja.deleteRows(filas.length + 2, sobra - 100);
}

function leerRegistro_(ss) {
  var h = ss.getSheetByName(H.REGISTRO);
  var out = {};
  if (!h || h.getLastRow() < 2) return out;
  h.getRange(2, 1, h.getLastRow() - 1, ENC_REGISTRO.length).getValues().forEach(function (f) {
    if (f[0]) out[f[0]] = f.map(function (x, i) { return i === 6 ? String(x) : x; });
  });
  return out;
}

function escribirRegistro_(ss, registro, orden) {
  var filas = orden.filter(function (id) { return registro[id]; }).map(function (id) { return registro[id]; });
  var h = ss.getSheetByName(H.REGISTRO);
  escribirTabla_(h, ENC_REGISTRO, filas); // "Modificado en Drive" = milisegundos (sirve para detectar cambios)
}

function formatearBD_(ss) {
  var bd = ss.getSheetByName(H.BD);
  var n = Math.max(1, bd.getLastRow() - 1);
  CAMPOS.forEach(function (c) {
    if (!c.dinero) return;
    bd.getRange(2, ENC_BD.indexOf(c.titulo) + 1, n, 1).setNumberFormat('$#,##0.00');
  });
}

// ====================== ALIAS (unificar nombres) ======================
function crearHojaAlias_(ss) {
  var h = ss.insertSheet(H.ALIAS);
  h.getRange(1, 1, 1, 3).setValues([['Tipo (PROVEEDOR o CHOFER)', 'Como aparece en la BD (columna Proveedor / Chofer)', 'Nombre oficial']])
    .setFontWeight('bold').setBackground('#6E8241').setFontColor('#FFFFFF');
  h.getRange(2, 1, 2, 3).setValues([
    ['PROVEEDOR', 'CHRISTIAN ANGEL', 'CHRISTIAN ALEXIS ANGEL CHAVEZ'],
    ['CHOFER', 'MODESTO', 'MODESTO SALAZAR'],
  ]);
  h.getRange(2, 1, 100, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['PROVEEDOR', 'CHOFER'], true).build());
  h.setColumnWidths(1, 3, 260);
  h.setFrozenRows(1);
}

function leerAlias_(ss) {
  var out = { PROVEEDOR: {}, CHOFER: {} };
  var h = ss.getSheetByName(H.ALIAS);
  if (!h || h.getLastRow() < 2) return out;
  h.getRange(2, 1, h.getLastRow() - 1, 3).getValues().forEach(function (f) {
    var tipo = normalizar(f[0]);
    if (out[tipo] && f[1] && f[2]) out[tipo][normalizar(f[1])] = normalizar(f[2]);
  });
  return out;
}

// ====================== RESÚMENES / KPIs ======================
function calcularResumenes_(ss, filasBD, ahora) {
  var I = {};
  ENC_BD.forEach(function (h, i) { I[h] = i; });
  var proyectos = {}, choferes = {}, proveedores = {};
  var validables = 0, validadas = 0, incidencias = 0, gasto = 0;

  filasBD.forEach(function (f) {
    var proyecto = f[I['Proyecto']] || '(sin proyecto)';
    var prov = f[I['Proveedor']], chofer = f[I['Chofer']];
    var costo = numero_(f[I['Costo total estimado']]) || numero_(f[I['Pago proveedor']]);
    var dias = numero_(f[I['Días de servicio']]);
    var vu = f[I['Validación unidad']], vd = f[I['Validación driver']];
    gasto += costo;

    var p = proyectos[proyecto] = proyectos[proyecto] || { productor: f[I['Productor']], mes: f[I['Mes']], inicio: f[I['Fecha inicio']] || f[I['Primer día']], unidades: 0, dias: 0, costo: 0, pago: 0, cinema: 0 };
    p.unidades++; p.dias += dias; p.costo += costo;
    p.pago += numero_(f[I['Pago proveedor']]); p.cinema += numero_(f[I['Cinema total']]);

    if (chofer) { var c = choferes[chofer] = choferes[chofer] || { n: 0, dias: 0 }; c.n++; c.dias += dias; }
    if (prov) { var q = proveedores[prov] = proveedores[prov] || { gasto: 0, n: 0 }; q.gasto += costo; q.n++; }

    if (vu !== '' || vd !== '') {
      validables++;
      if (vu === true && vd === true) validadas++;
    }
    if (normalizar(f[I['Incidencia']]).indexOf('S') === 0) incidencias++;
  });

  // Estas 3 pestañas mantienen 2 columnas para que el Dashboard siga funcionando
  escribirKPI_(ss, 'KPI Unidades por Proyecto', ['Proyecto', 'Unidades'],
    Object.keys(proyectos).map(function (k) { return [k, proyectos[k].unidades]; }).sort(function (a, b) { return b[1] - a[1]; }));
  escribirKPI_(ss, 'KPI Ranking Choferes', ['Chofer', 'Asignaciones'],
    Object.keys(choferes).map(function (k) { return [k, choferes[k].n]; }).sort(function (a, b) { return b[1] - a[1]; }));
  escribirKPI_(ss, 'KPI Gasto por Proveedor', ['Proveedor', 'Gasto Total'],
    Object.keys(proveedores).map(function (k) { return [k, proveedores[k].gasto]; }).sort(function (a, b) { return b[1] - a[1]; }));

  escribirKPI_(ss, H.RESUMEN_PROY, ['Proyecto', 'Productor', 'Mes', 'Inicio', 'Unidades', 'Días de servicio', 'Costo total', 'Pago proveedor', 'Cinema'],
    Object.keys(proyectos).map(function (k) {
      var p = proyectos[k];
      return [k, p.productor, p.mes, p.inicio, p.unidades, p.dias, p.costo, p.pago, p.cinema];
    }).sort(function (a, b) { return String(b[3]).localeCompare(String(a[3])); }));
  var rp = ss.getSheetByName(H.RESUMEN_PROY);
  if (rp.getLastRow() > 1) rp.getRange(2, 7, rp.getLastRow() - 1, 3).setNumberFormat('$#,##0');
  var kg = ss.getSheetByName('KPI Gasto por Proveedor');
  if (kg.getLastRow() > 1) kg.getRange(2, 2, kg.getLastRow() - 1, 1).setNumberFormat('$#,##0');

  // Resumen: el Dashboard lee la fila 4 como "Última actualización" (no cambiar el orden)
  var res = ss.getSheetByName('KPI Resumen') || ss.insertSheet('KPI Resumen');
  res.clearContents();
  res.getRange(1, 1, 7, 2).setValues([
    ['Indicador', 'Valor'],
    ['% Documentación validada', validables ? (validadas / validables * 100).toFixed(1) + '%' : 'N/A'],
    ['Incidencias registradas', incidencias],
    ['Última actualización', new Date()],
    ['Proyectos', Object.keys(proyectos).length],
    ['Unidades solicitadas', filasBD.length],
    ['Gasto total estimado', '$' + Math.round(gasto).toLocaleString('en-US')],
  ]);
}

function escribirKPI_(ss, nombre, encabezado, filas) {
  var h = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
  escribirTabla_(h, encabezado, filas);
}

// ====================== PANEL ======================
function prepararPanel() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var p = ss.getSheetByName(H.PANEL) || ss.insertSheet(H.PANEL, 0);
  p.clear();
  p.setHiddenGridlines(true);
  p.setColumnWidth(1, 30); p.setColumnWidth(2, 260); p.setColumnWidth(3, 520);
  p.getRange('B2').setValue('BASE MAESTRA DE TRANSPORTACIÓN — CINEMA SHUTTLE').setFontSize(16).setFontWeight('bold');
  p.getRange('B4').setValue('Estado:').setFontWeight('bold');
  p.getRange('B5').setValue('Detalle:').setFontWeight('bold');
  p.getRange('B6').setValue('Última corrida:').setFontWeight('bold');
  p.getRange('B8').setValue('CÓMO USARLO').setFontWeight('bold').setFontColor('#6E8241');
  p.getRange('B9:C15').setValues([
    ['1. Guarda los archivos', 'Mete cada archivo de proyecto (nuevo o anterior) en la carpeta de proyectos. Debe tener la pestaña "' + HOJA_ORIGEN + '".'],
    ['2. Consolida', 'Menú 📊 Base Maestra Transporte → ▶ Consolidar ahora (o el botón verde de esta hoja).'],
    ['3. Revisa', '"Registro Archivos" dice qué se leyó y si algún archivo no tiene la pestaña o dio error.'],
    ['4. Reportes', 'Usa "BD Solicitudes" (una fila por unidad) y "BD Uso Diario" (una fila por unidad por día) para tablas dinámicas, Looker Studio o el Dashboard.'],
    ['5. Nombres repetidos', 'Si un proveedor/chofer aparece escrito de varias formas, agrégalo en la hoja "Alias" y vuelve a correr "Reprocesar TODO".'],
    ['6. Automático', 'Menú → "Activar actualización automática" para que se actualice sola cada hora.'],
    ['Botón', 'Insertar → Dibujo → crea un rectángulo "CONSOLIDAR" → Guardar. Clic en el dibujo → ⋮ → Asignar secuencia de comandos → escribe: consolidarAhora'],
  ]);
  p.getRange('B9:B15').setFontWeight('bold').setVerticalAlignment('top');
  p.getRange('C9:C15').setWrap(true);
  p.getRange('B4:C6').setBackground('#F5F7F0');
}

function estadoPanel_(ss, estado, detalle) {
  var p = ss.getSheetByName(H.PANEL);
  if (!p) return;
  p.getRange('C4').setValue(estado).setFontWeight('bold');
  p.getRange('C5').setValue(detalle).setWrap(true);
  p.getRange('C6').setValue(new Date()).setNumberFormat('dd/mm/yyyy HH:mm');
  SpreadsheetApp.flush();
}

function avisar_(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

// ====================== DASHBOARD (WEB APP) ======================
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Dashboard')
    .setTitle('Tablero de Transportación')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// google.script.run no acepta objetos Date: se mandan números/texto
function obtenerDatosDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  function leer(nombre) {
    var hoja = ss.getSheetByName(nombre);
    if (!hoja || hoja.getLastRow() === 0) return [];
    return hoja.getDataRange().getValues().map(function (f) {
      return f.map(function (x) { return esFecha_(x) ? x.toISOString() : x; });
    });
  }
  return {
    unidadesPorProyecto: leer('KPI Unidades por Proyecto'),
    rankingChoferes: leer('KPI Ranking Choferes'),
    gastoPorProveedor: leer('KPI Gasto por Proveedor'),
    resumen: leer('KPI Resumen')
  };
}

// Para pruebas locales con Node (no afecta a Apps Script)
if (typeof module !== 'undefined') {
  module.exports = { extraerSolicitudes: extraerSolicitudes, normalizar: normalizar, normalizarProveedor: normalizarProveedor, normalizarChofer: normalizarChofer, limpiarNombreProyecto: limpiarNombreProyecto, ENC_BD: ENC_BD };
}
