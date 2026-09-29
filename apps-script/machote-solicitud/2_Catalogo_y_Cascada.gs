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

