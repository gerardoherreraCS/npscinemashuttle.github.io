// ====================== 5) FORMULARIO (PANEL LATERAL + WEB APP) ======================
// El MISMO archivo Formulario.html sirve para:
//  - el panel lateral dentro de la hoja (menú "Abrir formulario de captura")
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
    'IMPORTANTE: cada vez que cambies el código, ve a Implementar → Administrar implementaciones → (lápiz) → Versión: "Nueva versión" → Implementar. ' +
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
