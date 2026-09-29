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

  ui.alert('Hoja protegida (OK)\n\nCualquiera con el archivo puede llenar la solicitud. Solo estas personas pueden tocar fórmulas, encabezados o insertar/borrar filas y columnas:\n\n' + editores.join('\n') + '\n\nPara agregar a alguien, edita EDITORES_EQUIPO en el código y vuelve a correr esta opción.');
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

