/* exported onFormSubmit, getUltimoSelloNuevo, SH_RESPUESTAS */
/** Prueba de push Ulises 07/08/2025
 * Apps Script – Historial y Estado de Sellos + envío de correos
 * Autor: Nova (ChatGPT) · Proyecto Ulises – 1 ago 2025
 */

// ID de tu Spreadsheet
const SHEET_ID      = '1AgIVVj6OzdAh3QH8w2EIVxSIl175CIfOuHQWQyEddGU';

// Nombres de las pestañas
const SH_RESPUESTAS = 'Respuestas de formulario 1';
const SH_HISTORIAL  = 'Historial_Sellos';
const SH_ESTADO     = 'Estado_Actual';

// Nombre del sistema para asuntos y firma
const SYSTEM_NAME   = 'Centro Técnico GZ';

// Títulos EXACTOS de las preguntas en el formulario
const Q = {
  PDV          : 'Selecciona un punto de venta.',
  FECHA_VISITA : 'Fecha de la visita',
  MOTIVO       : 'Motivo de revisión',
  UBICACIONES  : [
    'Back Cámara de medición',
    'Front Cámara de medición',
    'Interconexión 1',
    'Interconexión 2',
    'Gaspar',
    'Preciador'
  ]
};

/** Lectura segura de namedValues */
function safe(nv, key) {
  return (nv[key] || [''])[0].toString().trim();
}

/** Plantillas de correo */
const TEMPLATES = {
  'Reparación': {
    subject: pdv => `Reparación en ${pdv} – ${SYSTEM_NAME}`,
    buildBody: ({ pdv, fechaVisita, nv, sellos }) =>
      `Se ha realizado una intervención técnica por motivo de Reparación.\n\n` +
      `________________________________________\n` +
      `\t Punto de Venta \n\n` +
      `➤ ${pdv}\n` +
      `________________________________________\n` +
      `\t Incidencia reportada \n` +
      `• ${safe(nv, 'Rep. Incidencia reportada')}\n\n` +
      `\t Diagnóstico de la falla \n` +
      `• ${safe(nv, 'Rep. Diagnóstico de la falla')}\n\n` +
      `\t Solución o acción realizada \n` + 
      `• ${safe(nv, 'Rep. Solución o acción realizada')}\n` +
      `________________________________________\n` +
      `\t Sellos nuevos instalados \n` +
      sellos.map(s => `➤ ${s.ubicacion}: ${s.sello}`).join('\n') +
      `\n\n— Centro Técnico GZ · Asistente: Nova`
  },

  'Cambio de pieza': {
    subject: pdv => `Cambio de pieza en ${pdv} – ${SYSTEM_NAME}`,
    buildBody: ({ pdv, fechaVisita, nv, sellos }) =>
      `Se ha realizado una intervención técnica por motivo de Cambio de pieza.\n\n` +
      `________________________________________\n` +
      `\t Punto de Venta \n\n` +
      `➤ ${pdv}\n` +
      `________________________________________\n` +
      `\t ID del componente dañado \n` +
      `• ${safe(nv, 'ID del componente Dañado')}\n\n` +
      `\t Acción sobre el componente dañado \n` +
      `• ${safe(nv, 'Acción sobre el componente dañado.')}\n\n` +
      `\t Componente instalado \n` +
      `• ${safe(nv, 'Selecciona un componente')}\n\n` +
      `\t Estado del componente \n` +
      `• ${safe(nv, '¿Cuál es el estado del componente?')}\n\n` +
      `\t Resultado final \n` +
      `• ${safe(nv, 'Cdp. Resultado final')}\n` +
      `________________________________________\n` +
      `\t Sellos nuevos instalados \n` +
      sellos.map(s => `➤ ${s.ubicacion}: ${s.sello}`).join('\n') +
      `\n\n— Centro Técnico GZ · Asistente: Nova`
  },

  'Chequeo de Impresora': {
    subject: pdv => `Chequeo de impresora en ${pdv} – ${SYSTEM_NAME}`,
    buildBody: ({ pdv, fechaVisita, nv, sellos }) =>
      `Se ha realizado una intervención técnica por motivo de Chequeo de impresora.\n\n` +
      `________________________________________\n` +
      `\t Punto de Venta \n\n` +
      `➤ ${pdv}\n` +
      `________________________________________\n` +
      `\t Incidencia reportada \n` +
      `• ${safe(nv, 'Imp. Incidencia reportada')}\n\n` +
      `\t Diagnóstico de la falla \n` +
      `• ${safe(nv, 'Imp. Diagnóstico de la falla')}\n\n` +
      `\t Solución o acción realizada \n` +
      `• ${safe(nv, 'Imp. Solución o acción realizada')}\n` +
      `________________________________________\n` +
      `\t Sellos nuevos instalados \n` +
      sellos.map(s => `➤ ${s.ubicacion}: ${s.sello}`).join('\n') +
      `\n\n— Centro Técnico GZ · Asistente: Nova`
  },

  'Revisión': {
    subject: pdv => `Revisión técnica en ${pdv} – ${SYSTEM_NAME}`,
    buildBody: ({ pdv, fechaVisita, nv, sellos }) =>
      `Se ha realizado una intervención técnica por motivo de Revisión.\n\n` +
      `________________________________________\n` +
      `\t Punto de Venta \n\n` +
      `➤ ${pdv}\n` +
      `________________________________________\n` +
      `\t Incidencia reportada \n` +
      `• ${safe(nv, 'Rev. Incidencia reportada')}\n\n` +
      `\t Diagnóstico \n` +
      `• ${safe(nv, 'Rev. Diagnóstico')}\n\n` +
      `\t Solución o acción realizada \n` +
      `• ${safe(nv, 'Rev. Solución o acción realizada')}\n` +
      `________________________________________\n` +
      `\t Sellos nuevos instalados \n` +
      sellos.map(s => `➤ ${s.ubicacion}: ${s.sello}`).join('\n') +
      `\n\n— Centro Técnico GZ · Asistente: Nova`
  },

  'Nueva Estación': {
    subject: pdv => `Nueva Estación en ${pdv} – ${SYSTEM_NAME}`,
    buildBody: ({ pdv, fechaVisita, nv, sellos }) =>
      `Se ha realizado una intervención técnica por motivo de Nueva Estación.\n\n` +
      `________________________________________\n` +
      `\t Punto de Venta \n\n` +
      `➤ ${pdv}\n` +
      `________________________________________\n` +
      `\t Descripción de la acción realizada \n` +
      `• ${safe(nv, 'Nva. Describe la acción realizada')}\n` +
      `________________________________________\n` +
      `\t Sellos nuevos instalados \n` +
      sellos.map(s => `➤ ${s.ubicacion}: ${s.sello}`).join('\n') +
      `\n\n— Centro Técnico GZ · Asistente: Nova`
  },

  'Accesos y configuración': {
    subject: pdv => `Accesos y configuración en ${pdv} – ${SYSTEM_NAME}`,
    buildBody: ({ pdv, fechaVisita, nv, sellos }) =>
      `Se ha realizado una intervención técnica por motivo de Accesos y configuración.\n\n` +
      `________________________________________\n` +
      `\t Punto de Venta \n\n` +
      `➤ ${pdv}\n` +
      `________________________________________\n` +
      `\t Tipo de acceso \n` +
      `• ${safe(nv, 'Acc. Tipo de Acceso')}\n\n` +
      `\t Motivo del acceso \n` +
      `• ${safe(nv, 'Acc. Motivo de los accesos')}\n\n` +
      `\t Configuración en Programador \n` +
      `• ${safe(nv, 'Acc. Configuración realizada en Programador')}\n\n` +
      `\t Configuración en Supervisor \n` +
      `• ${safe(nv, 'Acc. Configuración realizada en Supervisor')}\n\n` +
      `\t ¿Realiza reimpresión? \n` +
      `• ${safe(nv, 'Acc. ¿Realiza Reimpresión?')}\n` +
      `________________________________________\n` +
      `\t Sellos nuevos instalados \n` +
      sellos.map(s => `➤ ${s.ubicacion}: ${s.sello}`).join('\n') +
      `\n\n— Centro Técnico GZ · Asistente: Nova`
  }
};

function onFormSubmit(e) {
  if (!e || !e.namedValues) return;

  const nv          = e.namedValues;
  const pdv         = safe(nv, Q.PDV);
  const fechaVisita = parseFecha(safe(nv, Q.FECHA_VISITA));
  const motivo      = safe(nv, Q.MOTIVO);
  if (!pdv || !fechaVisita) return;

  // Abre hojas y mapea encabezados
  const ss       = SpreadsheetApp.openById(SHEET_ID);
  const shHist   = ss.getSheetByName(SH_HISTORIAL);
  const shEstado = ss.getSheetByName(SH_ESTADO);
  const idxHist   = headerMap(shHist);
  const idxEstado = headerMap(shEstado);

  // Asegura fila en Estado_Actual
  const rowEstado = findOrCreatePdVRow(shEstado, idxEstado, pdv);

  // Recolecta sellos nuevos
  const nuevosSellos = [];
  Q.UBICACIONES.forEach(u => {
    const r = safe(nv, u);
    if (!r) return;
    nuevosSellos.push({ ubicacion: u, sello: `C-${r}` });
    // Registro en Historial_Sellos
    const selloPrev = idxEstado[u]
      ? shEstado.getRange(rowEstado, idxEstado[u]).getValue()
      : '';
    shHist.appendRow(buildHistRow(idxHist, {
      PdV:              pdv,
      Ubicacion:        u,
      SelloNuevo:       `C-${r}`,
      FechaInstalacion: fechaVisita,
      Estado:           'Instalado',
      Motivo:           motivo
    }));
    if (selloPrev) {
      shHist.appendRow(buildHistRow(idxHist, {
        PdV:            pdv,
        Ubicacion:      u,
        SelloRemovido:  selloPrev,
        FechaRemocion:  fechaVisita,
        Estado:         'Removido',
        Motivo:         motivo
      }));
    }
    // Actualiza Estado_Actual
    if (idxEstado[u]) {
      shEstado.getRange(rowEstado, idxEstado[u]).setValue(`C-${r}`);
    }
  });

  // Prepara y envía correo
  const DESTINATARIO = 'jguadarrama.c@tomza.com,Facturacion.Colotlan@tomza.com';
  Logger.log(`🕵️‍♂️ Enviando correo para motivo "${motivo}" a: ${DESTINATARIO}`);

  const template = TEMPLATES[motivo];
  if (template) {
    const asunto = template.subject(pdv);
    const cuerpo = template.buildBody({ pdv, fechaVisita, nv, sellos: nuevosSellos });

    GmailApp.sendEmail(
      DESTINATARIO,
      asunto,
      cuerpo,
      {
        from:    'jguadarrama.c@tomza.com',
        name:    SYSTEM_NAME,
        replyTo: 'jguadarrama.c@tomza.com'
      }
    );
    Logger.log(`✅ Enviado a ${DESTINATARIO} con asunto: ${asunto}`);
  } else {
    Logger.log(`⚠️ No existe plantilla para motivo "${motivo}"`);
  }
}

/* ===================== HELPERS ===================== */

function headerMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.reduce((m, h, i) => {
    m[String(h).trim()] = i + 1;
    return m;
  }, {});
}

function findOrCreatePdVRow(sheet, idx, pdv) {
  const colPdV = idx['PdV'] || 1;
  const data   = sheet.getRange(2, colPdV, sheet.getLastRow() - 1, 1).getValues();
  for (let i=0; i<data.length; i++) {
    if (String(data[i][0]).trim().toUpperCase() === pdv.toUpperCase()) {
      return i + 2;
    }
  }
  const row = Array(sheet.getLastColumn()).fill('');
  row[colPdV - 1] = pdv;
  sheet.appendRow(row);
  return sheet.getLastRow();
}

function buildHistRow(idx, obj) {
  const fila = Array(Object.keys(idx).length).fill('');
  if (obj.PdV)              fila[idx['PdV']               - 1] = obj.PdV;
  if (obj.Ubicacion)        fila[idx['Ubicación']         - 1] = obj.Ubicacion;
  if (obj.SelloNuevo)       fila[idx['Sello Nuevo']       - 1] = obj.SelloNuevo;
  if (obj.SelloRemovido)    fila[idx['Sello Removido']    - 1] = obj.SelloRemovido;
  if (obj.FechaInstalacion) fila[idx['Fecha Instalación'] - 1] = obj.FechaInstalacion;
  if (obj.FechaRemocion)    fila[idx['Fecha Remoción']    - 1] = obj.FechaRemocion;
  if (obj.Estado)           fila[idx['Estado']            - 1] = obj.Estado;
  if (obj.Motivo)           fila[idx['Motivo de revisión'] - 1] = obj.Motivo;
  return fila;
}

function parseFecha(valor) {
  if (valor instanceof Date) return valor;
  const d = new Date(valor);
  return isNaN(d) ? new Date() : d;
}

function getUltimoSelloNuevo(pdv, ubicacion) {
  const data = SpreadsheetApp
    .openById(SHEET_ID)
    .getSheetByName(SH_HISTORIAL)
    .getDataRange()
    .getValues();
  const hdr = data.shift();
  const iPdV = hdr.indexOf('PdV');
  const iUb  = hdr.indexOf('Ubicación');
  const iS   = hdr.indexOf('Sello Nuevo');
  const iF   = hdr.indexOf('Fecha Instalación');
  let best   = '', maxD = new Date(0);
  data.forEach(r => {
    if (r[iPdV] === pdv && r[iUb] === ubicacion) {
      const d = new Date(r[iF]);
      if (d > maxD) { maxD = d; best = r[iS]; }
    }
  });
  return best;
}
//agrego texto para verificar un pull.