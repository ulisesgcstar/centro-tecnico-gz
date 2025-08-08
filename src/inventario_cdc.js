/* exported onEdit, afterSubmitInventario */
/******************************************************************
 * 1) Captura manual — escribir o arrastrar códigos en Inventario 
 * respaldo 29/07/25. Este codigo funciona con la alerta de sello duplicado.
 ******************************************************************/
// Este sí se disparará como simple trigger
function onEdit(e) {
  onEditInventario(e);
}

function onEditInventario(e){
  const sh = e.range.getSheet();
  if (sh.getName() !== 'Inventario_CdC' || e.range.getColumn() !== 1) return;

  const numRows = e.range.getNumRows();
  const codigos = e.range.getValues().map(r => String(r[0]).trim());

  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const hist = ss.getSheetByName('Historial_Sellos');
  const idxH = headerMap(hist);                       // índice 1-based

  // Construir un mapa {código → datos}
  const datosHist = hist.getRange(2, 1,
                      hist.getLastRow()-1, hist.getLastColumn()).getValues();
  const mapa = {};
  datosHist.forEach(r => {
    const c = String(r[idxH['Sello Nuevo']-1]).trim();
    if (c){
      mapa[c] = {
        pdv   : r[idxH['PdV']               -1],
        ubic  : r[idxH['Ubicación']         -1],
        fecha : r[idxH['Fecha Instalación'] -1],
        motivo: r[idxH['Motivo de revisión']-1]
      };
    }
  });

  // Preparar salida D-H
  const salida = codigos.map(c => {
    if (!c) return ['', '', '', '', ''];
    const i = mapa[c];
    return i
      ? ['Instalado', i.pdv, i.ubic, i.fecha, i.motivo]
      : ['Disponible', '',    '',    '',     ''];
  });

  sh.getRange(e.range.getRow(), 4, numRows, 5).setValues(salida);
}


/******************************************************************
 * 2) Captura automática — se ejecuta cada vez que el formulario
 *    agrega filas a Historial_Sellos
 ******************************************************************/
function afterSubmitInventario(e){
  /*--------------------------------------------------------------
   *  Datos básicos del formulario (ajusta los títulos si cambian)
   *------------------------------------------------------------*/
  const v      = e.namedValues;                          // {pregunta:[valor]}

// DEBUG: confirmar ejecución y ver qué llega
Logger.log('>>> afterSubmitInventario disparado');
Logger.log('>>> namedValues keys: ' + JSON.stringify(Object.keys(v)));
Logger.log('>>> namedValues ejemplo PdV: ' + JSON.stringify(v['Selecciona un punto de venta.']));

  const pdv    = (v['Selecciona un punto de venta.'] || [''])[0].trim();
  const motivo = (v['Motivo de revisión']             || [''])[0].trim();
  const fecha  = (v['Fecha de la visita']             || [''])[0];

  // Ubicaciones posibles en el formulario
  const UBIC = [
    'Back Cámara de medición',
    'Front Cámara de medición',
    'Interconexión 1',
    'Interconexión 2',
    'Gaspar',
    'Preciador'
  ];

  // Hoja Inventario + mapa de columnas
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const inv  = ss.getSheetByName('Inventario_CdC');
  const idxI = headerMap(inv);                           // 1-based

 /*--------------------------------------------------------------
 *  Recorrer cada ubicación marcada en la respuesta
 *------------------------------------------------------------*/
UBIC.forEach(u => {
  const resp = (v[u] || [''])[0].toString().trim();
  if (!resp) return;                                   // esa ubicación no se tocó
  const selloNuevo = `C-${resp}`;                      // formato del código

  // 1) Ejecuta la actualización en Inventario_CdC y guarda el resultado
  const fueInstalado = actualizarInventarioFila_(
    selloNuevo,
    pdv,
    u,
    fecha,
    motivo,
    inv,
    idxI
  );

  // 2) Si no estaba en inventario, registra la alerta
  if (!fueInstalado) {
    registrarAlerta(selloNuevo, pdv, u, fecha);
  }
});
}


/******************************************************************
 * 3) Helper: marca una fila como Instalado si el código existe
 ******************************************************************/
function actualizarInventarioFila_(codigo, pdv, ubic, fecha, motivo, invSheet, idxInv){
  const datos = invSheet.getRange(2, 1, invSheet.getLastRow()-1, 1).getValues();
  const pos   = datos.findIndex(r => String(r[0]).trim() === codigo);
  if (pos === -1) return false;                                // devuelve false si no lo encuentra

  const fila = pos + 2;                                  // compensar cabecera
  invSheet.getRange(fila, idxInv['Estado']           ).setValue('Instalado');
  invSheet.getRange(fila, idxInv['PdV']              ).setValue(pdv);
  invSheet.getRange(fila, idxInv['Ubicación']        ).setValue(ubic);
  invSheet.getRange(fila, idxInv['Fecha Instalación']).setValue(fecha);
  invSheet.getRange(fila, idxInv['Observaciones']    ).setValue(motivo);
  return true;  // indica que sí se actualizó
}

/**
 * Inserta una fila en ALERTAS con tipo de alerta y datos correspondientes
 */
function registrarAlerta(codigo, pdv, ubic, fecha) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const hist    = ss.getSheetByName('Historial_Sellos');
  const alertas = ss.getSheetByName('ALERTAS');
  const hmH     = headerMap(hist);
  const hmA     = headerMap(alertas);

    Logger.log('>>> registrarAlerta – código recibido: ' + codigo);
  Logger.log('>>> sellosHist raw: ' + JSON.stringify(
    hist.getRange(2, hmH['Sello Nuevo'], hist.getLastRow() - 1, 1)
        .getValues().flat()
  ));


    // 1) Determinar tipo de alerta usando Historial_Sellos
  const sellosHist = hist
    .getRange(2, hmH['Sello Nuevo'], hist.getLastRow()-1, 1)
    .getValues().flat().filter(String);
  const ocurrencias = sellosHist.filter(v => v === codigo).length;
  const tipo = (ocurrencias > 1)
    ? 'Sello duplicado en Historial_Sellos'
    : 'no existe en Inventario_CdC';


  // 2) Construir fila
  const row = [];
  row[ hmA['PdV']               -1 ] = pdv;
  row[ hmA['Ubicación']         -1 ] = ubic;
  row[ hmA['Sello reportado']   -1 ] = codigo;
  row[ hmA['Fecha Instalación'] -1 ] = fecha;
  row[ hmA['Tipo de alerta']    -1 ] = tipo;
  row[ hmA['Causa']             -1 ] = '';      // selección manual
  row[ hmA['Resuelto']          -1 ] = false;   // casilla vacía
  row[ hmA['Notas']             -1 ] = tipo;    // misma descripción

  alertas.appendRow(row);
}

