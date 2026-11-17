/**
 * Rechazos desde la guardia · Instituto Lanari
 * Corre en la cuenta institucional (franzosi.ricardo@lanari.uba.ar).
 *
 * - Lee todos los días los mails de guardia@lanari.uba.ar que dicen "rechazo".
 * - Agrega una fila por rechazo a la hoja "Rechazos guardia Lanari" (pestaña RECHAZOS).
 * - Si llega una anulación, marca el rechazo original como "Anulado".
 * - Lo que no puede interpretar con seguridad queda con Estado "A revisar".
 * - No modifica la bandeja: no marca como leído ni etiqueta. Registra los mails procesados en PROCESADOS.
 *
 * Puesta en marcha (una vez): ejecutar `instalar`. Crea la hoja (si no existe), carga la base
 * histórica de SEED, la comparte en modo lectura con el lector del panel y programa el disparador diario.
 */

const RCFG = {
  NOMBRE_HOJA: 'Rechazos guardia Lanari',
  REMITENTE: 'guardia@lanari.uba.ar',
  LECTOR_PANEL: 'ricardofranzosi@gmail.com', // cuenta que publica el panel (solo lectura)
  HORA_DIARIA: 7,
  VENTANA_DIAS: 10,
  ENCABEZADO: ['N°', 'Fecha', 'Hora', 'Mes', 'Paciente', 'DNI', 'N° beneficio PAMI', 'Tipo', 'Categoría de motivo',
    'Motivo del rechazo', 'Diagnóstico', 'Dónde aguardaba', 'Estado', 'Observación', 'Mail'],
};

// Base histórica: archivo depurado en Drive (xlsx, pestaña "Rechazos"). Se importa una sola vez en `instalar`.
const SEED_FILE_ID = '1vwS1VJoa3dH5l49wsS85-ZqXTXrXgObe';

/* ================= Instalación ================= */
function instalar() {
  const props = PropertiesService.getScriptProperties();
  let ss = props.getProperty('SHEET_ID') ? SpreadsheetApp.openById(props.getProperty('SHEET_ID')) : null;
  if (!ss) {
    ss = SpreadsheetApp.create(RCFG.NOMBRE_HOJA);
    props.setProperty('SHEET_ID', ss.getId());
    const sh = ss.getSheets()[0];
    sh.setName('RECHAZOS');
    sh.getRange(1, 1, 1, RCFG.ENCABEZADO.length).setValues([RCFG.ENCABEZADO]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('dd/mm/yyyy');
    sh.getRange('C:C').setNumberFormat('@');
    const pr = ss.insertSheet('PROCESADOS');
    pr.getRange(1, 1, 1, 4).setValues([['ID mail', 'Fecha mail', 'Resultado', 'Detalle']]).setFontWeight('bold');
    sh.getRange('F:G').setNumberFormat('@');
    if (SEED_FILE_ID) importarBase_(sh, props);
  }
  try { ss.addViewer(RCFG.LECTOR_PANEL); } catch (e) { console.warn('No se pudo compartir: ' + e); }
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'procesar').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('procesar').timeBased().everyDays(1).atHour(RCFG.HORA_DIARIA).create();
  console.log('Hoja: ' + ss.getUrl());
  return ss.getUrl();
}

/** Importa la base histórica desde el xlsx depurado (convierte una copia temporal a Hojas de cálculo). */
function importarBase_(sh, props) {
  DriveApp.getFileById(SEED_FILE_ID); // verifica acceso (y pide el permiso de Drive)
  const r = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + SEED_FILE_ID + '/copy', {
    method: 'post', contentType: 'application/json', headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: JSON.stringify({ name: 'tmp rechazos base', mimeType: 'application/vnd.google-apps.spreadsheet' }) });
  const tmpId = JSON.parse(r.getContentText()).id;
  try {
    const src = SpreadsheetApp.openById(tmpId).getSheetByName('Rechazos');
    const n = src.getLastRow() - 1;
    if (n > 0) {
      const vals = src.getRange(2, 1, n, 15).getValues();
      const disp = src.getRange(2, 1, n, 15).getDisplayValues();
      const links = src.getRange(2, 15, n, 1).getRichTextValues().map(x => x[0].getLinkUrl() || '');
      const tz = Session.getScriptTimeZone();
      let corte = '';
      const filas = vals.map((v, i) => ({ v, d: disp[i], l: links[i] })).filter(x => x.v[0] !== '' && x.v[0] !== null).map(({ v, d, l }) => {
        const hm = String(d[2] || '').match(/(\d{1,2}):(\d{2})/);
        const hora = hm ? hm[1].padStart(2, '0') + ':' + hm[2] : '';
        const f = v[1] instanceof Date ? v[1] : null;
        const clave = f ? Utilities.formatDate(f, tz, 'yyyy-MM-dd') + 'T' + hora : '';
        if (clave > corte) corte = clave;
        return [v[0], f || v[1], hora, v[3], v[4], d[5], d[6], v[7], v[8], v[9], v[10], v[11], v[12], v[13], l || ''];
      });
      sh.getRange(2, 1, filas.length, 15).setValues(filas);
      props.setProperty('CORTE', corte);
    }
  } finally { DriveApp.getFileById(tmpId).setTrashed(true); }
}

/** Vuelve a cargar la base histórica (borra las filas actuales de RECHAZOS y PROCESADOS). */
function reimportarBase() {
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.openById(props.getProperty('SHEET_ID'));
  const sh = ss.getSheetByName('RECHAZOS'), pr = ss.getSheetByName('PROCESADOS');
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 15).clearContent();
  if (pr.getLastRow() > 1) pr.getRange(2, 1, pr.getLastRow() - 1, 4).clearContent();
  sh.getRange('C:C').setNumberFormat('@'); sh.getRange('F:G').setNumberFormat('@');
  importarBase_(sh, props);
  console.log('Corte: ' + props.getProperty('CORTE'));
}

/* ================= Proceso diario ================= */
function procesar() {
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.openById(props.getProperty('SHEET_ID'));
  const sh = ss.getSheetByName('RECHAZOS'), pr = ss.getSheetByName('PROCESADOS');
  const tz = Session.getScriptTimeZone() || 'America/Argentina/Buenos_Aires';
  const corte = props.getProperty('CORTE') || '';
  const hechos = new Set(pr.getLastRow() > 1 ? pr.getRange(2, 1, pr.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : []);
  const q = 'from:' + RCFG.REMITENTE + ' (rechazo OR rchazo OR rc OR anula OR anulacion OR anulación) newer_than:' + RCFG.VENTANA_DIAS + 'd';
  const msgs = [];
  GmailApp.search(q, 0, 200).forEach(th => th.getMessages().forEach(m => {
    if (m.getFrom().toLowerCase().indexOf(RCFG.REMITENTE) < 0) return;
    if (hechos.has(m.getId())) return;
    const clave = Utilities.formatDate(m.getDate(), tz, "yyyy-MM-dd'T'HH:mm");
    if (corte && clave <= corte) return;
    msgs.push({ m, th, clave });
  }));
  msgs.sort((a, b) => a.clave < b.clave ? -1 : 1);
  const log = [];
  msgs.forEach(({ m, th, clave }) => {
    let res;
    try { res = procesarMensaje_(m, th, sh, tz); } catch (e) { res = ['Error', String(e)]; }
    log.push([m.getId(), clave, res[0], res[1] || '']);
  });
  if (log.length) pr.getRange(pr.getLastRow() + 1, 1, log.length, 4).setValues(log);
  return log.length;
}

function procesarMensaje_(m, th, sh, tz) {
  const asunto = m.getSubject() || th.getFirstMessageSubject() || '';
  const cuerpo = limpiarCuerpo_(m.getPlainBody() || '');
  const texto = (asunto + ' \n ' + cuerpo);
  const low = sinAcentos_(texto.toLowerCase());
  const fecha = m.getDate();
  const url = 'https://mail.google.com/mail/u/0/#all/' + m.getId();

  // Anulación
  if (/anul\w*\s+(?:el\s+|del\s+)?(?:rechazo|rc)\b|(?:rechazo|rc)\s+anulad|^\s*anulacion/.test(low)) {
    const ap = apellidoAsunto_(asunto) || apellidoAsunto_(th.getFirstMessageSubject() || '');
    const fila = buscarVigente_(sh, ap, fecha);
    if (fila) {
      sh.getRange(fila, 13).setValue('Anulado');
      sh.getRange(fila, 14).setValue('Anulado por mail del ' + Utilities.formatDate(fecha, tz, 'dd/MM/yyyy') + ': ' + cuerpo.slice(0, 160));
      return ['Anulación aplicada', 'fila ' + fila + ' (' + ap + ')'];
    }
    agregar_(sh, fecha, tz, { paciente: ap || '', tipo: 'Anulación', cat: 'No especificado', motivo: 'Anulación sin rechazo previo encontrado', estado: 'A revisar', obs: cuerpo.slice(0, 200), url });
    return ['Anulación a revisar', ap];
  }
  if (!/rechaz|rchaz|\brc\b/.test(low)) return ['Ignorado', 'no es rechazo'];

  const d = extraer_(asunto, cuerpo);
  // Duplicado: mismo paciente el mismo día
  if (esDuplicado_(sh, d, fecha, tz)) return ['Duplicado', d.paciente];
  const dudoso = !d.paciente || d.cat === 'No especificado';
  agregar_(sh, fecha, tz, { ...d, estado: dudoso ? 'A revisar' : 'Vigente', obs: dudoso ? cuerpo.slice(0, 200) : '', url });
  return [dudoso ? 'A revisar' : 'Agregado', d.paciente];
}

/**
 * Ejecutar a mano: completa "Dónde aguardaba" en las filas que lo tienen vacío,
 * releyendo el mail original (columna Mail) con la extracción actual. No toca nada más.
 */
function completarDonde() {
  const sh = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID')).getSheetByName('RECHAZOS');
  const v = sh.getDataRange().getValues();
  const h = v.findIndex(r => r.indexOf('Dónde aguardaba') >= 0);
  const cD = v[h].indexOf('Dónde aguardaba'), cM = v[h].indexOf('Mail');
  let n = 0;
  for (let i = h + 1; i < v.length; i++) {
    if (String(v[i][cD]).trim() || !v[i][cM]) continue;
    const id = String(v[i][cM]).split('/').pop();
    try {
      const msgs = GmailApp.getThreadById(id).getMessages();
      for (const m of msgs) {
        const d = extraer_(m.getSubject(), limpiarCuerpo_(m.getPlainBody()));
        if (d.donde) { sh.getRange(i + 1, cD + 1).setValue(d.donde); n++; console.log(v[i][4] + ' → ' + d.donde); break; }
      }
    } catch (e) { console.warn('Fila ' + (i + 1) + ': ' + e); }
  }
  console.log('Completadas: ' + n);
}

/* ================= Extracción ================= */
function extraer_(asunto, cuerpo) {
  const t = cuerpo.replace(/\s+/g, ' ').replace(/\b(Dr|Dra|Sr|Sra|Nro|N)\.\s*/g, '$1 ');
  const low = sinAcentos_(t.toLowerCase());
  const lowAs = sinAcentos_(asunto.toLowerCase());
  let paciente = '';
  let m = t.match(/(?:paciente|para el paciente|para la paciente)\s*:?\s+([A-Za-zÁÉÍÓÚÜÑáéíóúüñ' ,.-]+?)(?=\s*(?:,|\.|;|\bcon\b|\bDNI\b|\bdni\b|\bafiliad|\bbenef|\bBENEF|\bNro\b|\bN°|\bpor\b|\bel mismo\b|\bla misma\b|\bque\b|\bdx\b|\bDX\b|\bquien\b|$))/);
  if (m) {
    let nom = m[1].trim();
    const resto = t.slice(m.index + m[0].length).match(/^,\s*([A-ZÁÉÍÓÚÑ][a-záéíóúñ]+(?:\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+){0,2})(?=\s*(?:,|\.|;|\bcon\b|\bDNI\b|\bpor\b|$))/);
    if (resto) nom += ', ' + resto[1];
    if (!/^(que|el|se|en|para|con)\b/i.test(nom)) paciente = titulo_(nom);
  }
  if (!paciente) paciente = titulo_(apellidoAsunto_(asunto));
  const dni = (t.match(/\bDNI\s*(?:N[°º]?|nro\.?)?\s*:?\s*([\d.]{6,11})/i) || [])[1];
  const ben = (t.match(/\b(\d{9,14}\/\d{2,3})\b/) || t.match(/(?:benef\w*|afiliad\w*|beneficiari\w*)\s*(?:N[°º]?|nro)?\s*:?\s*(\d{9,16})/i) || [])[1];
  const dx = (t.match(/\b(?:dx|diagn[oó]stico(?:\s+de)?)\s*:?\s*([^.;,\n]{3,80})/i) || [])[1];
  // Lugar donde aguarda: "aguarda en", "espera en", "se encuentra en", "permanece en", "aguarda en: …" (hasta el punto)
  let donde = (t.match(/\b(?:aguarda(?:ndo)?|espera(?:ndo)?|se encuentra|permanece|esta|está|ubicad[oa])\s+en\s*:?\s+(?:la |el |su )?([^.;\n]{3,120}?)(?=\s+(?:dx|con dx|por|con diagn)\b|\s+--|\s+secretar[ií]a de guardia|[.;\n]|$)/i) || [])[1];
  if (donde) donde = donde.replace(/,\s*(?:dx|diagn).*$/i, '').replace(/\s+(?:saludos|atte\.?|atentamente|gracias|muchas gracias)\b.*$/i, '');
  // Tipo
  const tipo = /ambulancia/.test(lowAs) && !/derivacion/.test(lowAs) || /rechazo de ambulancia/.test(low) ? 'Rechazo de ambulancia'
    : /derivacion/.test(lowAs) || /derivacion/.test(low) ? 'Rechazo + pedido de derivación' : 'Rechazo conformado';
  // Motivo
  const cama = /(falta|faltas|sin|no contar con|no cuenta con|no hay|no tener|no disponer de)[^.]{0,40}cama/.test(low) || /camas? disponibles?/.test(low) && /no /.test(low);
  const areas = [];
  if (/(uti|terapia intensiva|unidad de terapia)/.test(low)) areas.push('UTI');
  if (/sala general/.test(low)) areas.push('sala general');
  if (/guardia externa/.test(low)) areas.push('guardia externa');
  const tomo = /tomograf/.test(low);
  const esp = low.match(/(?:falta de|sin|no contar con|no hay)\s+(?:medico\s+)?(traumatolog\w*|otorrino\w*|psiquiatr\w*|ginecolog\w*|cirujan\w*|urolog\w*|neurolog\w*|neurocirujan\w*|oftalmolog\w*|especialista\w*|cardiolog\w*|medico\s+\w+)/);
  const insumo = /(insumo|instrumental|material|equipamiento|protesis|caja quirurgica|elemento)/.test(low);
  const quirofano = /quirofano/.test(low);
  let cat = 'No especificado', motivo = '';
  if (cama && !(tomo && low.indexOf('tomograf') < low.search(/cama/))) {
    cat = 'Falta de camas';
    motivo = areas.length === 0 ? 'Falta de camas en la institución' : 'Falta de camas en ' + areas.join(' y ');
    if (tomo) motivo = 'Falta de camas y sin tomógrafo';
  } else if (tomo) { cat = 'Sin tomógrafo'; motivo = cama ? 'Sin tomógrafo y sin camas' + (areas.length ? ' en ' + areas.join(' y ') : '') : 'Sin tomógrafo disponible'; }
  else if (esp) { cat = 'Falta de especialista de guardia'; motivo = 'Sin ' + esp[1] + ' de guardia'; }
  else if (quirofano) { cat = 'Instrumental / insumos / equipamiento'; motivo = 'Sin quirófano disponible'; }
  else if (insumo) { cat = 'Instrumental / insumos / equipamiento'; motivo = 'Falta de instrumental / insumos / equipamiento'; }
  else if (/ascensor/.test(low)) { cat = 'Ascensor inhabilitado'; motivo = 'Ascensor inhabilitado'; }
  else if (/complejidad/.test(low)) { cat = 'Falta de complejidad'; motivo = 'Falta de complejidad'; }
  else if (/(sin|falta de|no contar con|no cuenta con|no hay) ambulancia/.test(low)) { cat = 'Sin ambulancia'; motivo = 'Sin ambulancia'; }
  else motivo = 'No especificado en el mail';
  if (donde) {
    donde = donde.trim();
    if (/^(ambulancia|movil|móvil)/i.test(donde)) donde = 'Ambulancia ' + donde.replace(/^(la |el )?(ambulancia|movil|móvil)\s*/i, '').trim();
    donde = donde.charAt(0).toUpperCase() + donde.slice(1);
  }
  return { paciente, dni: dni ? dni.replace(/\D/g, '') : '', ben: ben || '', tipo, cat, motivo, dx: dx ? capital_(dx.trim()) : 'No consignado', donde: donde || '' };
}

/* ================= Hoja ================= */
function agregar_(sh, fecha, tz, d) {
  const last = sh.getLastRow();
  const n = last > 1 ? (Number(sh.getRange(last, 1).getValue()) || last - 1) + 1 : 1;
  const dia = new Date(Utilities.formatDate(fecha, tz, 'yyyy/MM/dd'));
  sh.appendRow([n, dia, Utilities.formatDate(fecha, tz, 'HH:mm'), Number(Utilities.formatDate(fecha, tz, 'M')), d.paciente || '', d.dni || '', d.ben || '',
    d.tipo || '', d.cat || '', d.motivo || '', d.dx || '', d.donde || '', d.estado || 'Vigente', d.obs || '', d.url || '']);
}
function buscarVigente_(sh, apellido, fecha) {
  if (!apellido) return 0;
  const ap = sinAcentos_(apellido.toLowerCase());
  const n = sh.getLastRow(); if (n < 2) return 0;
  const vals = sh.getRange(2, 1, n - 1, 13).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    const pac = sinAcentos_(String(vals[i][4]).toLowerCase()), est = String(vals[i][12]);
    const f = vals[i][1] instanceof Date ? vals[i][1] : null;
    if (est === 'Vigente' && pac.indexOf(ap) >= 0 && (!f || (fecha - f) / 864e5 <= 15)) return i + 2;
  }
  return 0;
}
function esDuplicado_(sh, d, fecha, tz) {
  const n = sh.getLastRow(); if (n < 2 || !d.paciente) return false;
  const hoy = Utilities.formatDate(fecha, tz, 'yyyy-MM-dd');
  const ap = sinAcentos_(d.paciente.toLowerCase().split(/[ ,]+/)[0]);
  const desde = Math.max(2, n - 40);
  const vals = sh.getRange(desde, 1, n - desde + 1, 13).getValues();
  return vals.some(r => r[1] instanceof Date && Utilities.formatDate(r[1], tz, 'yyyy-MM-dd') === hoy && r[12] !== 'Anulado' &&
    ((d.dni && String(r[5]) === d.dni) || sinAcentos_(String(r[4]).toLowerCase()).indexOf(ap) === 0));
}

/* ================= Utilidades ================= */
function limpiarCuerpo_(s) {
  const corte = s.search(/\n\s*(El .{5,80} escribi[oó]:|On .{5,80} wrote:|-{2,}\s*Forwarded|De:\s.+\nEnviado)/i);
  if (corte > 0) s = s.slice(0, corte);
  return s.replace(/\n--\s*\n[\s\S]*$/, '').replace(/\n\s*--\s*$[\s\S]*/m, '').replace(/\n\s*Secretar[ií]a de guardia[\s\S]*$/i, '').replace(/^>.*$/gm, '').trim();
}
function apellidoAsunto_(asunto) {
  const partes = String(asunto || '').toUpperCase().replace(/^(RE|RV|FW|FWD):\s*/g, '').split(/[\/|-]/).map(x => x.trim()).filter(Boolean);
  const ruido = /^(LANARI|LANRI|LANARO|CABA|RC|RECHAZO|RECHAZOS|DERIVACION|DERIVACIÓN|AMBULANCIA|ANULACION|ANULACIÓN|ANULACION RECHAZO|CONFORMADO)$/;
  const cand = partes.filter(p => !ruido.test(p));
  let ap = cand.length ? cand[cand.length - 1] : '';
  ap = ap.replace(/^(ANULACI[OÓ]N\s+)?(DE\s+)?RECHAZO\s+/, '').trim();
  return ap;
}
function sinAcentos_(s) { return String(s).normalize('NFD').replace(/[̀-ͯ]/g, ''); }
function titulo_(s) { return String(s || '').toLowerCase().replace(/(^|[\s,'-])([a-záéíóúüñ])/g, (a, b, c) => b + c.toUpperCase()).trim(); }
function capital_(s) { return s.charAt(0).toUpperCase() + s.slice(1); }

/** Prueba sin escribir nada: muestra cómo se interpretarían los últimos mails. */
function probar() {
  const tz = Session.getScriptTimeZone();
  GmailApp.search('from:' + RCFG.REMITENTE + ' rechazo newer_than:10d', 0, 10).forEach(th => th.getMessages().forEach(m => {
    if (m.getFrom().toLowerCase().indexOf(RCFG.REMITENTE) < 0) return;
    console.log(Utilities.formatDate(m.getDate(), tz, 'dd/MM HH:mm') + ' | ' + JSON.stringify(extraer_(m.getSubject(), limpiarCuerpo_(m.getPlainBody()))));
  }));
}
