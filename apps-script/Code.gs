/**
 * Auditoría Médica y Gestión Hospitalaria Lanari
 * Backend de acceso para el panel publicado en GitHub Pages.
 *
 * - Solo acepta mails @lanari.uba.ar (y, si existe la pestaña "ACCESO", solo los mails listados ahí).
 * - Envía un código de 6 dígitos al mail (vence en 10 minutos, 5 intentos).
 * - Con el código correcto entrega un token de sesión (6 horas).
 * - Con el token devuelve el contenido ACTUAL de la planilla (valores tal como se ven).
 *
 * Se publica como "Aplicación web": Ejecutar como = Yo · Quién tiene acceso = Cualquier usuario.
 * La URL es pública, pero sin código válido no entrega datos.
 */

// Los identificadores de planillas, carpetas y archivos NO están en este archivo (el repositorio es público):
// se guardan en Propiedades del script (clave IDS, JSON) dentro del proyecto de Apps Script.
const IDS = (() => { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('IDS') || '{}'); } catch (e) { return {}; } })();

const CONFIG = {
  SHEET_ID: IDS.SHEET_ID,          // planilla de Admisión (SOLICITUD DE CAMAS PARA INTERNACIÓN)
  RECHAZOS_ID: IDS.RECHAZOS_ID,    // hoja "Rechazos guardia Lanari"
  // Estadísticas: carpetas "Est. especialidad 2024 / 2025 / 2026" (Corina Páez) y la planilla base 2024.
  ESTAD_CARPETAS: IDS.ESTAD_CARPETAS || [],
  ESTAD_EXTRA: IDS.ESTAD_EXTRA || [],
  ESTAD_HORA: 6,
  // Facturación PAMI: hoja privada "Panel Lanari - datos Facturación PAMI" y carpeta "Facturación PAMI Lanari (privado)".
  // Pestañas: FAC_LIQ, FAC_FACT, FAC_MOD, REC_LAB, REC_OPS, LIBRO, REC_RECH, META.
  FACT_ID: IDS.FACT_ID,
  FACT_CARPETA: IDS.FACT_CARPETA,  // donde la verificación mensual deja portal_pami_AAAAMM.xlsx
  GESTION_ID: IDS.GESTION_ID,      // planilla "gestion" de Cobranzas (facturas emitidas)
  FOTOS_ID: IDS.FOTOS_ID,          // fotos de las responsables (archivo privado)
  DOMINIO: 'lanari.uba.ar',
  HOJA_ACCESO: 'ACCESO',          // opcional: lista blanca de mails en la columna A
  HOJA_LOG: 'LOG ACCESOS',        // opcional: si existe, se registran ingresos y consultas
  HOJAS_EXCLUIDAS: ['ACCESO', 'LOG ACCESOS'],
  CODIGO_TTL_S: 600,              // 10 minutos
  SESION_TTL_S: 21600,            // 6 horas (máximo de CacheService)
  MAX_INTENTOS: 5,
  MAX_CODIGOS_POR_HORA: 4,
  NOMBRE_REMITENTE: 'Auditoría Médica y Gestión Hospitalaria Lanari',
  // Registro de usuarios y copias diarias: hoja y carpeta privadas en el Drive del dueño del script (se crean con instalarRegistro).
  ADMINS: ['franzosi.ricardo@lanari.uba.ar'],   // únicos que ven la sección Cambios del panel
  COPIA_HORA: 7, COPIA_MINUTO: 0,               // copia diaria del tablero (camas + rechazos)
  COPIA_HORAS: 24,                              // cada copia se conserva 24 h (queda solo la última)
  INFORME_A: 'franzosi.ricardo@lanari.uba.ar',    // recibe a diario (7:00) los cambios del tablero y los ingresos de usuarios
};

function doGet() {
  return json_({ ok: true, servicio: 'Auditoría Médica y Gestión Hospitalaria Lanari', mensaje: 'Usar POST desde el panel.' });
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { return json_({ ok: false, error: 'solicitud_invalida' }); }
  try {
    switch (body.action) {
      case 'request': return json_(pedirCodigo_(body.email));
      case 'verify': return json_(verificarCodigo_(body.email, body.code));
      case 'data': return json_(datos_(body.token));
      case 'logout': return json_(salir_(body.token));
      case 'estadisticas': return json_(estadisticas_(body.token));
      case 'facturacion': return json_(facturacion_(body.token));
      case 'vista': return json_(vista_(body.token, body.vista, body.disp));
      case 'copias': return json_(copias_(body.token));
      case 'copia': return json_(copia_(body.token, body.fecha));
      default: return json_({ ok: false, error: 'accion_desconocida' });
    }
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'error_interno' });
  }
}

/* ---------- Acceso ---------- */
function normEmail_(email) { return String(email || '').trim().toLowerCase(); }

function emailPermitido_(email) {
  const re = new RegExp('^[^@\\s]+@' + CONFIG.DOMINIO.replace(/\./g, '\\.') + '$');
  if (!re.test(email)) return false;
  const hoja = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.HOJA_ACCESO);
  if (!hoja) return true; // sin lista blanca: cualquier mail del dominio
  const lista = hoja.getRange(1, 1, Math.max(1, hoja.getLastRow()), 1).getDisplayValues().map(r => normEmail_(r[0]));
  return lista.indexOf(email) >= 0;
}

function pedirCodigo_(emailRaw) {
  const email = normEmail_(emailRaw);
  const cache = CacheService.getScriptCache();
  if (!emailPermitido_(email)) {
    log_(email, 'codigo_rechazado');
    return { ok: false, error: 'mail_no_autorizado' };
  }
  const rlKey = 'rl:' + email;
  const n = Number(cache.get(rlKey) || 0);
  if (n >= CONFIG.MAX_CODIGOS_POR_HORA) return { ok: false, error: 'demasiados_codigos' };
  cache.put(rlKey, String(n + 1), 3600);

  const codigo = codigoAleatorio_();
  cache.put('otp:' + email, JSON.stringify({ h: hash_(codigo + email), i: 0 }), CONFIG.CODIGO_TTL_S);

  MailApp.sendEmail({
    to: email,
    subject: 'Código de acceso: ' + codigo + ' · Auditoría Médica y Gestión Hospitalaria Lanari',
    name: CONFIG.NOMBRE_REMITENTE,
    body: 'Tu código para ingresar al panel de Auditoría Médica y Gestión Hospitalaria Lanari es: ' + codigo +
      '\n\nVence en 10 minutos. Si no lo pediste, ignorá este mensaje.',
    htmlBody: '<p>Tu código para ingresar al panel de Auditoría Médica y Gestión Hospitalaria Lanari es:</p>' +
      '<p style="font-size:28px;font-weight:bold;letter-spacing:6px;font-family:monospace">' + codigo + '</p>' +
      '<p style="color:#666">Vence en 10 minutos. Si no lo pediste, ignorá este mensaje.</p>',
  });
  log_(email, 'codigo_enviado');
  return { ok: true };
}

function verificarCodigo_(emailRaw, codigo) {
  const email = normEmail_(emailRaw);
  const cache = CacheService.getScriptCache();
  const key = 'otp:' + email;
  const raw = cache.get(key);
  if (!raw) return { ok: false, error: 'codigo_vencido' };
  const otp = JSON.parse(raw);
  if (otp.i >= CONFIG.MAX_INTENTOS) { cache.remove(key); return { ok: false, error: 'demasiados_intentos' }; }
  if (hash_(String(codigo || '').trim() + email) !== otp.h) {
    otp.i++;
    cache.put(key, JSON.stringify(otp), CONFIG.CODIGO_TTL_S);
    return { ok: false, error: 'codigo_incorrecto', restantes: CONFIG.MAX_INTENTOS - otp.i };
  }
  cache.remove(key);
  const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
  cache.put('tok:' + token, email, CONFIG.SESION_TTL_S);
  log_(email, 'ingreso');
  return { ok: true, token: token, email: email, vence: Date.now() + CONFIG.SESION_TTL_S * 1000 };
}

function sesion_(token) {
  if (!token || typeof token !== 'string') return null;
  return CacheService.getScriptCache().get('tok:' + token);
}

function datos_(token) {
  const email = sesion_(token);
  if (!email) return { ok: false, error: 'sesion_vencida' };
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const hojas = {};
  ss.getSheets().forEach(sh => {
    const nombre = sh.getName();
    if (CONFIG.HOJAS_EXCLUIDAS.indexOf(nombre) >= 0) return;
    hojas[nombre] = sh.getDataRange().getDisplayValues();
  });
  if (CONFIG.RECHAZOS_ID) {
    try { hojas['RECHAZOS GUARDIA'] = SpreadsheetApp.openById(CONFIG.RECHAZOS_ID).getSheetByName('RECHAZOS').getDataRange().getDisplayValues(); }
    catch (err) { console.warn('Rechazos no disponibles: ' + err); }
  }
  return { ok: true, email: email, admin: esAdmin_(email), generado: new Date().toISOString(), hojas: hojas, fotos: fotos_() };
}

function fotos_() {
  if (!CONFIG.FOTOS_ID) return {};
  const c = CacheService.getScriptCache(), k = 'fotos_' + CONFIG.FOTOS_ID; const h = c.get(k); // la clave cambia si cambia el archivo
  if (h) return JSON.parse(h);
  try { const t = DriveApp.getFileById(CONFIG.FOTOS_ID).getBlob().getDataAsString(); c.put(k, t, 21600); return JSON.parse(t); }
  catch (e) { console.warn('Fotos: ' + e); return {}; }
}

function salir_(token) {
  if (token) CacheService.getScriptCache().remove('tok:' + token);
  return { ok: true };
}

/* ---------- Estadísticas (copia automática para el panel) ---------- */
function estadisticas_(token) {
  const email = sesion_(token);
  if (!email) return { ok: false, error: 'sesion_vencida' };
  const id = PropertiesService.getScriptProperties().getProperty('EST_CACHE_ID');
  if (!id) return { ok: false, error: 'estadisticas_no_configuradas' };
  const hojas = {};
  SpreadsheetApp.openById(id).getSheets().forEach(sh => { hojas[sh.getName()] = sh.getDataRange().getDisplayValues(); });
  return { ok: true, hojas: hojas };
}

/* ---------- Facturación PAMI (hoja privada) ---------- */
function facturacion_(token) {
  const email = sesion_(token);
  if (!email) return { ok: false, error: 'sesion_vencida' };
  if (!CONFIG.FACT_ID) return { ok: false, error: 'facturacion_no_configurada' };
  const tz = Session.getScriptTimeZone();
  const hojas = {};
  SpreadsheetApp.openById(CONFIG.FACT_ID).getSheets().forEach(sh => {
    hojas[sh.getName()] = sh.getDataRange().getValues().map(r => r.map(v => v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : v));
  });
  return { ok: true, hojas: hojas };
}

/** Ejecutar una vez: crea la copia y el disparador diario. */
function instalarEstadisticas() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'actualizarEstadisticas').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('actualizarEstadisticas').timeBased().everyDays(1).atHour(CONFIG.ESTAD_HORA).create();
  return actualizarEstadisticas(true);
}

/** Convierte las planillas de Estadística a texto plano en una hoja propia que el panel lee rápido. */
function actualizarEstadisticas(forzar) {
  forzar = forzar === true; // desde un disparador llega un objeto de evento
  const t0 = Date.now(); let pendientes = 0;
  const props = PropertiesService.getScriptProperties();
  let cache = props.getProperty('EST_CACHE_ID') ? SpreadsheetApp.openById(props.getProperty('EST_CACHE_ID')) : null;
  if (!cache) { cache = SpreadsheetApp.create('Panel camas Lanari - estadísticas (copia automática)'); props.setProperty('EST_CACHE_ID', cache.getId()); }
  const tz = Session.getScriptTimeZone();
  const ahora = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
  const vistos = JSON.parse(props.getProperty('EST_VISTOS') || '{}');
  // 1) Archivos de la autora (generales y por especialidad)
  const archivos = [];
  CONFIG.ESTAD_CARPETAS.forEach(cid => {
    const it = DriveApp.getFolderById(cid).getFiles();
    while (it.hasNext()) {
      const f = it.next(); const n = f.getName();
      if (!/\.xlsx$/i.test(n) || !/^est/i.test(n)) continue;
      archivos.push({ id: f.getId(), titulo: n, mod: f.getLastUpdated(), tipo: /^estad[ií]sticas 20\d\d\.xlsx$/i.test(n) ? 'general' : 'especialidad' });
    }
  });
  // Por título, quedarse con el más reciente (hay copias en varias carpetas)
  const porTitulo = {};
  archivos.forEach(a => { const k = a.titulo.toLowerCase(); if (!porTitulo[k] || a.mod > porTitulo[k].mod) porTitulo[k] = a; });
  const lista = Object.values(porTitulo).sort((a, b) => a.titulo.localeCompare(b.titulo));
  CONFIG.ESTAD_EXTRA.forEach(x => lista.unshift({ id: x.id, titulo: x.titulo, mod: DriveApp.getFileById(x.id).getLastUpdated(), tipo: 'general', gsheet: true, anio: x.anio }));
  const meta = [['Pestaña', 'Tipo', 'Año', 'Archivo', 'Hoja', 'Modificado', 'Sincronizado']];
  const usadas = new Set(['EST_META']);
  lista.forEach(a => {
    const anioGen = a.anio || (a.titulo.match(/(20\d\d)/) || [])[1] || '';
    // Prefijo estable por archivo (no cambia si se agregan archivos nuevos)
    if (a.tipo !== 'general' && !vistos[a.id + '#p']) { vistos.__n = (vistos.__n || 0) + 1; vistos[a.id + '#p'] = 'ESP ' + vistos.__n; }
    const prefijo = a.tipo === 'general' ? 'EST ' + anioGen : vistos[a.id + '#p'];
    const firma = a.id + '|' + a.mod.getTime();
    let hojas = null;
    let cambiado = forzar || vistos[a.id] !== firma;
    if (cambiado && Date.now() - t0 > 4.5 * 60 * 1000) { cambiado = false; pendientes++; } // límite de 6 min por ejecución
    if (cambiado) {
      let tmpId = null;
      try {
        let ss;
        if (a.gsheet) ss = SpreadsheetApp.openById(a.id);
        else {
          const r = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + a.id + '/copy', { method: 'post', contentType: 'application/json',
            headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, payload: JSON.stringify({ name: 'tmp estadistica', mimeType: 'application/vnd.google-apps.spreadsheet' }), muteHttpExceptions: true });
          tmpId = JSON.parse(r.getContentText()).id; ss = SpreadsheetApp.openById(tmpId);
        }
        hojas = ss.getSheets().map(sh => ({ nombre: sh.getName(), vals: (a.tipo === 'especialidad' ? sanitizar_ : x => x)(sh.getDataRange().getValues().map(row => row.map(v => v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : (v === null ? '' : String(v))))) }));
      } catch (e) { console.warn(a.titulo + ': ' + e); }
      finally { if (tmpId) try { DriveApp.getFileById(tmpId).setTrashed(true); } catch (e) {} }
    }
    const n = hojas ? hojas.length : Number(vistos[a.id + '#n'] || 0);
    for (let i = 0; i < n; i++) {
      const tab = prefijo + ' ' + (i + 1); usadas.add(tab);
      let sh = cache.getSheetByName(tab);
      if (hojas) {
        if (!sh) sh = cache.insertSheet(tab); sh.clear();
        const v = hojas[i].vals; const w = Math.max(1, ...v.map(r => r.length));
        const norm = v.map(r => r.concat(Array(w - r.length).fill('')));
        if (norm.length) { sh.getRange(1, 1, norm.length, w).setNumberFormat('@').setValues(norm); }
      }
      meta.push([tab, a.tipo, anioGen, a.titulo, hojas ? hojas[i].nombre : (vistos[a.id + '#h' + i] || ''), Utilities.formatDate(a.mod, tz, 'yyyy-MM-dd'), ahora]);
      if (hojas) vistos[a.id + '#h' + i] = hojas[i].nombre;
    }
    if (hojas) { vistos[a.id] = firma; vistos[a.id + '#n'] = hojas.length; }
  });
  // Pestañas que ya no corresponden
  cache.getSheets().forEach(sh => { if (!usadas.has(sh.getName()) && cache.getSheets().length > 1) cache.deleteSheet(sh); });
  let m = cache.getSheetByName('EST_META'); if (!m) m = cache.insertSheet('EST_META'); m.clear();
  m.getRange(1, 1, meta.length, meta[0].length).setNumberFormat('@').setValues(meta);
  const def = cache.getSheetByName('Hoja 1') || cache.getSheetByName('Sheet1'); if (def && cache.getSheets().length > 1) cache.deleteSheet(def);
  props.setProperty('EST_VISTOS', JSON.stringify(vistos));
  if (pendientes) {
    ScriptApp.newTrigger('actualizarEstadisticas').timeBased().after(60 * 1000).create();
    console.log('Quedan ' + pendientes + ' archivos; se continúa en 1 minuto.');
  }
  limpiarDisparadoresUnicos_();
  console.log('Archivos: ' + lista.length + ' · pestañas: ' + (meta.length - 1));
  return lista.length;
}

/** Borra los disparadores de continuación ya usados (los que no son el diario). */
function limpiarDisparadoresUnicos_() {
  const ts = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'actualizarEstadisticas');
  // El diario se crea primero; los de continuación se agregan después. Si hay más de 2, borrar los más viejos de continuación.
  if (ts.length > 2) ts.slice(1, ts.length - 1).forEach(t => ScriptApp.deleteTrigger(t));
}

/**
 * Quita filas con texto en las columnas de datos (B a M): son listados nominales de pacientes
 * (p. ej. Neurocirugía: nombre, estadía, mes, cirujano) que el panel no debe recibir.
 * Se conservan los encabezados de tablas con meses en filas (Servicio Social).
 */
function sanitizar_(vals) {
  const MESES = /^(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|setiembre|octubre|noviembre|diciembre)$/i;
  const esMesA = r => r && MESES.test(String(r[0]).trim());
  const esNum = x => x !== '' && !isNaN(Number(String(x).replace(',', '.')));
  const textos = r => r.slice(1, 13).filter(c => { const x = String(c).trim(); return x && !esNum(x) && !/^\d{4}-\d{2}-\d{2}$/.test(x) && !/^[a-z]{2,4}\.?-\d{2}$/i.test(x) && !/^(total|tot|%)$/i.test(x); }).length;
  return vals.map((r, i) => textos(r) >= 2 && !esMesA(vals[i + 1]) && !esMesA(vals[i + 2]) ? ['(fila omitida)'] : r);
}

/* ---------- Utilidades ---------- */
function codigoAleatorio_() {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Date.now());
  let n = 0;
  for (let i = 0; i < 4; i++) n = (n * 256 + (bytes[i] & 0xff)) % 1000000007;
  return String(n % 1000000).padStart(6, '0');
}
function hash_(s) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s));
}
/* ---------- Registro de usuarios (hoja privada) ---------- */
function registro_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('REG_ID'), ss = null;
  if (id) try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; }
  if (!ss) {
    ss = SpreadsheetApp.create('Panel Lanari - registro de usuarios (privado)');
    props.setProperty('REG_ID', ss.getId());
    const sh = ss.getSheets()[0]; sh.setName('INGRESOS');
    sh.appendRow(['Fecha y hora', 'Mail', 'Acción', 'Pestaña', 'Dispositivo']); sh.setFrozenRows(1);
  }
  return ss;
}
function log_(email, accion, pestana, disp) {
  try {
    const sh = registro_().getSheetByName('INGRESOS');
    sh.appendRow([Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd HH:mm:ss'), email, accion, pestana || '', disp || '']);
  } catch (err) { console.warn(err); }
}
function esAdmin_(email) { return CONFIG.ADMINS.indexOf(String(email || '').toLowerCase()) >= 0; }

/** El panel avisa qué pestaña abre el usuario (una vez por pestaña cada 10 minutos). */
function vista_(token, vista, disp) {
  const email = sesion_(token);
  if (!email) return { ok: false, error: 'sesion_vencida' };
  const v = String(vista || '').replace(/[^a-z_]/gi, '').slice(0, 30);
  if (!v) return { ok: true };
  const c = CacheService.getScriptCache(), k = 'vis:' + token + ':' + v;
  if (c.get(k)) return { ok: true };
  c.put(k, '1', 600);
  log_(email, 'vista', v, String(disp || '').slice(0, 20));
  return { ok: true };
}

/** Arma la pestaña RESUMEN de la hoja de registro (se ejecuta con la copia diaria). */
function resumenUsuarios() {
  const ss = registro_(), v = ss.getSheetByName('INGRESOS').getDataRange().getDisplayValues().slice(1);
  const hoy = Date.now(), d30 = Utilities.formatDate(new Date(hoy - 30 * 864e5), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd');
  const NOM = { mapa: 'Mapa de camas', solicitud: 'Solicitud', movimientos: 'Movimientos', espera: 'Espera', indicadores: 'Indicadores', rechazos: 'Rechazos', estadisticas: 'Estadísticas', facturacion: 'Facturación PAMI', obrassociales: 'Facturación OS', cambios: 'Cambios' };
  const U = {};
  v.forEach(r => {
    const [f, mail, acc, pes, disp] = r; if (!mail) return;
    const u = U[mail] = U[mail] || { mail, primero: f, ultimo: '', ing: 0, ing30: 0, dias: new Set(), vistas: {}, disp: {} };
    if (f > u.ultimo) u.ultimo = f; if (f < u.primero) u.primero = f;
    if (acc === 'ingreso') { u.ing++; if (f >= d30) u.ing30++; }
    if (acc === 'vista' || acc === 'ingreso') u.dias.add(f.slice(0, 10));
    if (acc === 'vista' && pes) u.vistas[pes] = (u.vistas[pes] || 0) + 1;
    if (disp) u.disp[disp] = (u.disp[disp] || 0) + 1;
  });
  const filas = Object.values(U).sort((a, b) => b.ultimo.localeCompare(a.ultimo)).map(u => [
    u.mail, u.primero, u.ultimo, u.ing, u.ing30, u.dias.size,
    Object.entries(u.vistas).sort((a, b) => b[1] - a[1]).map(([k, n]) => (NOM[k] || k) + ' (' + n + ')').join(', '),
    Object.entries(u.disp).sort((a, b) => b[1] - a[1]).map(([k, n]) => k + ' (' + n + ')').join(', ')]);
  let sh = ss.getSheetByName('RESUMEN'); if (!sh) sh = ss.insertSheet('RESUMEN', 0);
  sh.clear();
  const enc = ['Mail', 'Primer registro', 'Último movimiento', 'Ingresos (total)', 'Ingresos últimos 30 días', 'Días con uso', 'Pestañas vistas (veces)', 'Dispositivo'];
  sh.getRange(1, 1, 1, enc.length).setValues([enc]).setFontWeight('bold'); sh.setFrozenRows(1);
  if (filas.length) sh.getRange(2, 1, filas.length, enc.length).setValues(filas);
  sh.getRange(filas.length + 3, 1).setValue('Actualizado: ' + Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd HH:mm') + ' · el detalle de cada movimiento está en la pestaña INGRESOS');
}

/* ---------- Copias diarias del tablero (camas + rechazos) ---------- */
function carpetaCopias_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('COPIAS_ID');
  if (id) try { return DriveApp.getFolderById(id); } catch (e) {}
  const f = DriveApp.createFolder('Panel Lanari - copias diarias del tablero (privado)');
  props.setProperty('COPIAS_ID', f.getId());
  return f;
}
/** Disparador diario: guarda el estado de la planilla de camas y de rechazos como JSON. */
function copiaDiaria() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID), hojas = {};
  ss.getSheets().forEach(sh => { if (CONFIG.HOJAS_EXCLUIDAS.indexOf(sh.getName()) < 0) hojas[sh.getName()] = sh.getDataRange().getDisplayValues(); });
  try { hojas['RECHAZOS GUARDIA'] = SpreadsheetApp.openById(CONFIG.RECHAZOS_ID).getSheetByName('RECHAZOS').getDataRange().getDisplayValues(); } catch (e) { console.warn(e); }
  const fecha = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd');
  const carpeta = carpetaCopias_(), nombre = 'copia_' + fecha + '.json';
  // Copia anterior (para el informe por mail)
  let prev = null; const itp = carpeta.getFiles(); let fp = null;
  while (itp.hasNext()) { const f = itp.next(); if (f.getName() !== nombre && /^copia_.*\.json$/.test(f.getName()) && (!fp || f.getDateCreated() > fp.getDateCreated())) fp = f; }
  if (fp) try { prev = JSON.parse(fp.getBlob().getDataAsString()); } catch (e) { console.warn(e); }
  const mismas = carpeta.getFilesByName(nombre); while (mismas.hasNext()) mismas.next().setTrashed(true);
  carpeta.createFile(nombre, JSON.stringify({ fecha: fecha, generado: new Date().toISOString(), hojas: hojas }), 'application/json');
  // Retención: se borran las copias con más de COPIA_HORAS (la nueva reemplaza a la del día anterior)
  const limite = Date.now() - (CONFIG.COPIA_HORAS - 1) * 36e5, it = carpeta.getFiles();
  while (it.hasNext()) { const f = it.next(); if (f.getName() !== nombre && f.getDateCreated().getTime() < limite) f.setTrashed(true); }
  console.log('Copia guardada: ' + nombre);
  try { resumenUsuarios(); } catch (e) { console.warn('Resumen usuarios: ' + e); }
  try { enviarInforme_(prev, { fecha: fecha, hojas: hojas }); } catch (e) { console.warn('Informe: ' + e); }
}

/* ---------- Informe diario por mail ---------- */
function enviarInforme_(prev, actual) {
  const tz = 'America/Argentina/Buenos_Aires', e = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;');
  const U = s => String(s || '').trim().toUpperCase();
  const camasDe = h => { const m = {}; const rows = h['LISTADO DIARIO'] || []; let ok = false;
    rows.forEach(r => { const a = U(r[0]); if (a === 'CAMAS') { ok = true; return; } if (!ok || !a || !U(r[2])) return; if (/PROGRAMADOS/.test(U(r[1]))) { ok = false; return; } if (!m[a]) m[a] = { pac: String(r[1] || '').trim(), est: U(r[2]), ais: U(r[3]) }; }); return m; };
  const tabla = (tit, enc, filas) => `<h3 style="margin:18px 0 6px;font:600 15px Arial">${tit} (${filas.length})</h3>` + (filas.length
    ? `<table style="border-collapse:collapse;font:13px Arial">${'<tr>' + enc.map(x => `<th style="text-align:left;border-bottom:1px solid #ccc;padding:4px 8px">${x}</th>`).join('') + '</tr>'}${filas.map(f => '<tr>' + f.map(x => `<td style="border-bottom:1px solid #eee;padding:4px 8px">${x}</td>`).join('') + '</tr>').join('')}</table>`
    : '<p style="font:13px Arial;color:#777">Sin cambios.</p>');
  let html = `<div style="font:14px Arial;color:#222"><h2 style="font:600 18px Arial;margin:0 0 4px">Auditoría Médica y Gestión Hospitalaria Lanari · informe diario</h2>
    <p style="color:#666;margin:0">${Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm')}${prev ? ' · cambios desde la copia del ' + prev.fecha.split('-').reverse().join('/') : ' · primera copia: todavía no hay con qué comparar'}</p>`;
  if (prev) {
    const A = camasDe(prev.hojas), B = camasDe(actual.hojas), d = x => x ? e(x.est.toLowerCase()) + (x.pac ? ' · ' + e(x.pac) : '') + (x.ais && x.ais !== 'SIN AISLAMIENTO' && x.ais !== 'NO' ? ' · aisl.' : '') : '—';
    const camas = [...new Set(Object.keys(A).concat(Object.keys(B)))].filter(k => !A[k] || !B[k] || A[k].est !== B[k].est || A[k].pac !== B[k].pac || A[k].ais !== B[k].ais).map(k => [`<b>${e(k)}</b>`, d(A[k]), d(B[k])]);
    html += tabla('Camas con cambios', ['Cama', 'Antes', 'Ahora'], camas);
    const ped = h => { const r = h['Pedidos'] || []; const H = (r[0] || []).map(U); const ix = n => H.findIndex(x => x.indexOf(n) >= 0);
      const iT = ix('MARCA TEMPORAL') >= 0 ? ix('MARCA TEMPORAL') : ix('TIMESTAMP'), iN = ix('NOMBRE'), iO = ix('OBRA SOCIAL'), iI = ix('FECHA DE INGRESO'), iM = ix('MOTIVO');
      return r.slice(1).filter(x => x.some(v => String(v).trim())).map(x => ({ k: x[iT] + '|' + x[iN], n: x[iN], os: x[iO], ing: x[iI], mot: x[iM] })); };
    const pa = new Set(ped(prev.hojas).map(x => x.k));
    html += tabla('Pedidos de internación nuevos', ['Paciente', 'Obra social', 'Ingreso', 'Motivo'], ped(actual.hojas).filter(x => !pa.has(x.k)).map(x => [e(x.n), e(x.os), e(x.ing), e(x.mot)]));
    const rech = h => { const r = h['RECHAZOS GUARDIA'] || []; const H = (r[0] || []).map(U); const ix = n => H.indexOf(n);
      return r.slice(1).map(x => ({ n: x[ix('N°')], f: x[ix('FECHA')] + ' ' + x[ix('HORA')], p: x[ix('PACIENTE')], m: x[ix('MOTIVO DEL RECHAZO')], d: x[ix('DÓNDE AGUARDABA')], est: x[ix('ESTADO')] })).filter(x => x.n); };
    const ra = {}; rech(prev.hojas).forEach(x => ra[x.n] = x);
    const rb = rech(actual.hojas);
    html += tabla('Rechazos de guardia nuevos', ['N°', 'Fecha', 'Paciente', 'Motivo', 'Dónde aguardaba'], rb.filter(x => !ra[x.n]).map(x => [e(x.n), e(x.f), e(x.p), e(x.m), e(x.d || '—')]));
    const rc = rb.filter(x => ra[x.n] && (ra[x.n].est !== x.est || ra[x.n].d !== x.d)).map(x => [e(x.n), e(x.p), e(ra[x.n].est + (ra[x.n].d ? ' · ' + ra[x.n].d : '')), e(x.est + (x.d ? ' · ' + x.d : ''))]);
    if (rc.length) html += tabla('Rechazos modificados', ['N°', 'Paciente', 'Antes', 'Ahora'], rc);
  }
  // Ingresos de usuarios desde el informe anterior (sin repetir lo ya informado)
  const props = PropertiesService.getScriptProperties(), ahora = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  const desde = props.getProperty('INFORME_DESDE') || Utilities.formatDate(new Date(Date.now() - 864e5), tz, 'yyyy-MM-dd HH:mm:ss');
  const v = registro_().getSheetByName('INGRESOS').getDataRange().getDisplayValues().slice(1).filter(r => r[0] > desde && r[0] <= ahora && r[1]);
  const NOM = { mapa: 'Mapa', solicitud: 'Solicitud', movimientos: 'Movimientos', espera: 'Espera', indicadores: 'Indicadores', rechazos: 'Rechazos', estadisticas: 'Estadísticas', facturacion: 'Fact. PAMI', obrassociales: 'Fact. OS', cambios: 'Cambios' };
  const U2 = {};
  v.forEach(r => { const u = U2[r[1]] = U2[r[1]] || { ing: 0, prim: r[0], ult: r[0], vis: {}, disp: r[4] || '' }; if (r[2] === 'ingreso') u.ing++; if (r[0] < u.prim) u.prim = r[0]; if (r[0] > u.ult) u.ult = r[0]; if (r[2] === 'vista' && r[3]) u.vis[r[3]] = 1; if (r[4]) u.disp = r[4]; });
  html += tabla('Usuarios que usaron el panel desde el informe anterior', ['Mail', 'Ingresos', 'Primer / último movimiento', 'Pestañas vistas', 'Dispositivo'],
    Object.entries(U2).sort((a, b) => b[1].ult.localeCompare(a[1].ult)).map(([m, u]) => [e(m), u.ing, e(u.prim.slice(11, 16) + ' – ' + u.ult.slice(11, 16) + (u.prim.slice(0, 10) !== u.ult.slice(0, 10) ? ' (' + u.prim.slice(8, 10) + '/' + u.prim.slice(5, 7) + '→' + u.ult.slice(8, 10) + '/' + u.ult.slice(5, 7) + ')' : '')), e(Object.keys(u.vis).map(k => NOM[k] || k).join(', ') || '—'), e(u.disp || '—')]));
  html += `<p style="font:12px Arial;color:#777;margin-top:18px">Registro completo: <a href="${registro_().getUrl()}">Panel Lanari - registro de usuarios</a> · Panel: https://ricardo-franzosi.github.io/lanari-gestion-camas/</p></div>`;
  MailApp.sendEmail({ to: CONFIG.INFORME_A, subject: 'Informe diario del tablero Lanari · ' + Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy'), htmlBody: html, name: CONFIG.NOMBRE_REMITENTE });
  props.setProperty('INFORME_DESDE', ahora);
}
function copias_(token) {
  const email = sesion_(token);
  if (!email) return { ok: false, error: 'sesion_vencida' };
  if (!esAdmin_(email)) return { ok: false, error: 'sin_permiso' };
  const it = carpetaCopias_().getFiles(), fechas = [];
  while (it.hasNext()) { const m = it.next().getName().match(/^copia_(\d{4}-\d{2}-\d{2})\.json$/); if (m) fechas.push(m[1]); }
  return { ok: true, fechas: fechas.sort().reverse() };
}
function copia_(token, fecha) {
  const email = sesion_(token);
  if (!email) return { ok: false, error: 'sesion_vencida' };
  if (!esAdmin_(email)) return { ok: false, error: 'sin_permiso' };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(fecha))) return { ok: false, error: 'fecha_invalida' };
  const it = carpetaCopias_().getFilesByName('copia_' + fecha + '.json');
  if (!it.hasNext()) return { ok: false, error: 'sin_copia' };
  log_(email, 'cambios', fecha);
  return Object.assign({ ok: true }, JSON.parse(it.next().getBlob().getDataAsString()));
}
/** Ejecutar una vez: crea la hoja de registro, la carpeta de copias, el disparador diario y la primera copia. */
function instalarRegistro() {
  registro_(); carpetaCopias_();
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'copiaDiaria').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('copiaDiaria').timeBased().everyDays(1).atHour(CONFIG.COPIA_HORA).nearMinute(CONFIG.COPIA_MINUTO).inTimezone('America/Argentina/Buenos_Aires').create();
  copiaDiaria();
  console.log('Registro: ' + registro_().getUrl());
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Ejecutar una vez desde el editor para autorizar los permisos (planilla y envío de mails). */
function autorizar() {
  SpreadsheetApp.openById(CONFIG.SHEET_ID).getName();
  MailApp.getRemainingDailyQuota();
}

/* ---------- Facturación PAMI: actualización automática ----------
 * Carpeta privada en Drive (CONFIG.FACT_CARPETA) con:
 *  - la hoja de datos del panel (CONFIG.FACT_ID),
 *  - los archivos "portal_pami_AAAAMM.xlsx" que deja la verificación mensual del portal PAMI
 *    (pestañas FAC_LIQ, FAC_MOD, REC_LAB y REC_OPS).
 * actualizarFacturacion() corre todos los días a las 7:30 (después de la carga de rechazos):
 *  1. si hay un archivo del portal más nuevo, copia sus pestañas a la hoja de datos;
 *  2. rearma FAC_FACT con las facturas PAMI de la planilla "gestion" de Cobranzas;
 *  3. recalcula REC_RECH: valor en $ de cada rechazo de guardia (mails y libro) cruzando con las OP de recupero.
 */
function instalarFacturacion() {
  PropertiesService.getScriptProperties().deleteProperty('FACT_PORTAL'); // fuerza reimportar el último archivo del portal
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'actualizarFacturacion').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('actualizarFacturacion').timeBased().everyDays(1).atHour(7).nearMinute(30).create();
  return actualizarFacturacion();
}

function actualizarFacturacion() {
  const tz = Session.getScriptTimeZone();
  const ss = SpreadsheetApp.openById(CONFIG.FACT_ID);
  const props = PropertiesService.getScriptProperties();
  const log = [];
  // 1) Importar el último archivo del portal
  if (CONFIG.FACT_CARPETA) {
    let ult = null;
    const it = DriveApp.getFolderById(CONFIG.FACT_CARPETA).getFiles();
    while (it.hasNext()) { const f = it.next(); if (/^portal_pami_\d{6}.*\.xlsx$/i.test(f.getName()) && (!ult || f.getLastUpdated() > ult.getLastUpdated())) ult = f; }
    const firma = ult ? ult.getId() + '|' + ult.getLastUpdated().getTime() : '';
    if (ult && props.getProperty('FACT_PORTAL') !== firma) {
      let tmp = null;
      try {
        const r = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + ult.getId() + '/copy', { method: 'post', contentType: 'application/json',
          headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, payload: JSON.stringify({ name: 'tmp portal pami', mimeType: 'application/vnd.google-apps.spreadsheet' }), muteHttpExceptions: true });
        tmp = JSON.parse(r.getContentText()).id;
        const src = SpreadsheetApp.openById(tmp);
        ['FAC_LIQ', 'FAC_MOD', 'REC_LAB', 'REC_OPS'].forEach(n => { const sh = src.getSheetByName(n); if (sh) escribirHoja_(ss, n, sh.getDataRange().getValues().map(r => r.map(v => v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM') : v))); }); // períodos como texto AAAA-MM
        props.setProperty('FACT_PORTAL', firma);
        log.push('portal: ' + ult.getName());
      } finally { if (tmp) try { DriveApp.getFileById(tmp).setTrashed(true); } catch (e) {} }
    }
  }
  // 2) Facturas PAMI de la planilla "gestion" (Cobranzas)
  if (CONFIG.GESTION_ID) {
    try {
      const prev = {}; const fx = ss.getSheetByName('FAC_FACT');
      if (fx) fx.getDataRange().getValues().slice(1).forEach(r => { if (r[6] === 'SI') prev[String(r[2])] = true; });
      const v = SpreadsheetApp.openById(CONFIG.GESTION_ID).getSheets()[0].getDataRange().getValues();
      const out = [['fecha_envio', 'tipo', 'numero', 'importe', 'fecha_pago', 'debito', 'es_capita_dc']];
      v.forEach(r => {
        const os = String(r[1] || '').toUpperCase().replace(/\s/g, '');
        if (!(/^INS.*JP/.test(os) || /PAMI|UGL/.test(os))) return;
        const f = r[0] instanceof Date ? r[0] : null; if (!f || f.getFullYear() < 2025) return;
        const tipo = /^N\s*[\/.]?\s*C|NOTA/i.test(String(r[2]).trim()) ? 'NC' : 'FACTURA'; // NC, N/C, N. CREDITO
        out.push([Utilities.formatDate(f, tz, 'yyyy-MM-dd'), tipo, String(r[3] || ''), typeof r[4] === 'number' ? r[4] : '', r[7] instanceof Date ? Utilities.formatDate(r[7], tz, 'yyyy-MM-dd') : String(r[7] || ''),
          typeof r[8] === 'number' ? r[8] : '', prev[String(r[3] || '')] ? 'SI' : '']);
      });
      escribirHoja_(ss, 'FAC_FACT', out); log.push('facturas: ' + (out.length - 1));
      // Todas las obras sociales desde 2025 (pestaña Facturación OS). Se copian tal cual; el panel agrupa las variantes de nombre.
      const os = [['fecha_envio', 'obra_social', 'tipo', 'numero', 'importe', 'forma_envio', 'fecha_recibido', 'fecha_pago', 'debito', 'saldo']];
      const fd = x => x instanceof Date ? Utilities.formatDate(x, tz, 'yyyy-MM-dd') : String(x == null ? '' : x).trim();
      v.forEach(r => {
        const f = r[0] instanceof Date ? r[0] : null; if (!f || f.getFullYear() < 2025 || !String(r[1] || '').trim()) return;
        const tipo = /^N\s*[\/.]?\s*C|NOTA/i.test(String(r[2]).trim()) ? 'NC' : 'FACTURA';
        os.push([fd(f), String(r[1]).trim(), tipo, String(r[3] || ''), typeof r[4] === 'number' ? r[4] : '', String(r[5] || '').trim(), fd(r[6]), fd(r[7]),
          typeof r[8] === 'number' ? r[8] : '', typeof r[9] === 'number' ? r[9] : '']);
      });
      escribirHoja_(ss, 'FAC_OS', os); log.push('comprobantes OS: ' + (os.length - 1));
    } catch (e) { console.warn('gestion: ' + e); log.push('gestion: error ' + e); }
  }
  // 3) Valor de cada rechazo
  const ops = (ss.getSheetByName('REC_OPS') || { getDataRange: () => ({ getValues: () => [[]] }) }).getDataRange().getValues();
  const libro = ss.getSheetByName('LIBRO') ? ss.getSheetByName('LIBRO').getDataRange().getValues() : [[]];
  let guardia = [[]];
  try { guardia = SpreadsheetApp.openById(CONFIG.RECHAZOS_ID).getSheetByName('RECHAZOS').getDataRange().getValues(); } catch (e) { console.warn('rechazos: ' + e); }
  const fmt = d => d instanceof Date ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : d;
  const rech = construirRecRech_(ops.map(r => r.map(fmt)), guardia.map(r => r.map(fmt)), libro.map(r => r.map(fmt)));
  escribirHoja_(ss, 'REC_RECH', rech); log.push('rechazos/episodios: ' + (rech.length - 1));
  escribirHoja_(ss, 'META', [['clave', 'valor'], ['generado', Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm')],
    ['fuente_portal', 'efectores.pami.org.ar' + (props.getProperty('FACT_PORTAL') ? ' · ' + DriveApp.getFileById(props.getProperty('FACT_PORTAL').split('|')[0]).getName() : '')],
    ['fuente_facturas', 'Planilla "gestion" (Cobranzas)'], ['responsable', 'Griselda Rios']]);
  console.log(log.join(' · '));
  return log;
}

function escribirHoja_(ss, nombre, vals) {
  let sh = ss.getSheetByName(nombre); if (!sh) sh = ss.insertSheet(nombre);
  sh.clearContents();
  if (!vals.length) return;
  const w = Math.max(...vals.map(r => r.length));
  const rg = sh.getRange(1, 1, vals.length, w);
  // Columnas de texto (períodos "2026-01", N° de OP o beneficio) en formato texto para que Sheets no las convierta en fechas o números
  const txt = vals[0].map((_, j) => vals.slice(1).some(r => typeof r[j] === 'string' && r[j] !== ''));
  txt.forEach((t, j) => { if (t) sh.getRange(1, j + 1, vals.length, 1).setNumberFormat('@'); });
  rg.setValues(vals.map(r => r.concat(Array(w - r.length).fill(''))));
}

/** Cruce rechazos ↔ OP de recupero (módulo 69). Función pura: recibe y devuelve matrices de valores. */
function construirRecRech_(opsV, guardiaV, libroV) {
  // Las celdas con fecha llegan como Date: pasarlas a texto AAAA-MM-DD antes de comparar períodos
  const p2 = n => String(n).padStart(2, '0');
  const sinFechas = M => (M || []).map(r => r.map(v => v instanceof Date ? v.getFullYear() + '-' + p2(v.getMonth() + 1) + '-' + p2(v.getDate()) : v));
  opsV = sinFechas(opsV); guardiaV = sinFechas(guardiaV); libroV = sinFechas(libroV);
  const z = s => String(s == null ? '' : s).replace(/^0+/, '');
  const ymI = p => +p.slice(0, 4) * 12 + +p.slice(5, 7);
  const perTxt = p => p.slice(5, 7) + '/' + p.slice(0, 4);
  const num = v => typeof v === 'number' ? v : Number(String(v || '0').replace(',', '.')) || 0;
  const ed1 = (a, b) => { if (a === b) return true; if (Math.abs(a.length - b.length) > 1) return false; let i = 0, j = 0, d = 0;
    while (i < a.length && j < b.length) { if (a[i] === b[j]) { i++; j++; continue; } if (++d > 1) return false; if (a.length > b.length) i++; else if (b.length > a.length) j++; else { i++; j++; } }
    return d + (a.length - i) + (b.length - j) <= 1; };
  const H = opsV[0] || []; const ix = k => H.indexOf(k);
  const ops = opsV.slice(1).filter(r => r[0] !== '' && r[0] != null).map(r => ({ op: String(r[ix('op')]), ben: z(r[ix('beneficio')]), per: String(r[ix('periodo_prest')]).slice(0, 7), prest: String(r[ix('prestador')]),
    mt: num(r[ix('monto_recupero')]), cg: num(r[ix('cargo_adm')]), lq: String(r[ix('liquidaciones')] || ''), pazNom: String(r[ix('paciente_planilla_paez')] || ''), pazNota: String(r[ix('nota_planilla_paez')] || '') }));
  const maxPer = ops.map(o => o.per).sort().pop() || '';
  const byB = {}; ops.forEach(o => (byB[o.ben] = byB[o.ben] || []).push(o)); const keys = Object.keys(byB);
  const buscar = (nb, ym) => { // OP del beneficiario en el mes del rechazo ±1
    if (!nb) return []; let c = byB[nb];
    if (!c) { const ks = keys.filter(k => ed1(k, nb)); if (ks.length === 1) c = byB[ks[0]]; }
    if (!c || !ym) return []; const m = ymI(ym); return c.filter(o => Math.abs(ymI(o.per) - m) <= 1);
  };
  // Hoja de rechazos de guardia
  const hi = guardiaV.findIndex(r => r.some(v => /categor[ií]a de motivo/i.test(String(v))));
  const gh = hi >= 0 ? guardiaV[hi].map(String) : []; const gc = re => gh.findIndex(h => re.test(h));
  const G = hi < 0 ? [] : guardiaV.slice(hi + 1).filter(r => r[gc(/^N°/)] !== '' && r[gc(/^N°/)] != null).map(r => {
    let f = String(r[gc(/^Fecha/i)] || ''); const m = f.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/); if (m) f = m[3] + '-' + m[2].padStart(2, '0') + '-' + m[1].padStart(2, '0');
    const bRaw = String(r[gc(/benef/i)] || '').replace(/[^0-9/]/g, '');
    const nb = bRaw.includes('/') ? bRaw.split('/')[0] : (bRaw.length >= 13 ? bRaw.slice(0, -2) : bRaw);
    return { n: String(r[gc(/^N°/)]), fecha: f.slice(0, 10), pac: String(r[gc(/^Paciente/i)] || ''), ben: z(nb), motivo: String(r[gc(/^Motivo del/i)] || ''), estado: String(r[gc(/^Estado/i)] || ''), vig: /^vigente$/i.test(String(r[gc(/^Estado/i)] || '').trim()) };
  });
  const LBh = (libroV[0] || []).map(String); const lc = k => LBh.indexOf(k);
  const LB = libroV.slice(1).filter(r => r[lc('mes')]).map(r => ({ fecha: String(r[lc('fecha')]).slice(0, 10), mes: String(r[lc('mes')]).slice(0, 7), pac: String(r[lc('paciente')] || ''),
    ben: z(String(r[lc('beneficio')] || '').replace(/[^0-9/]/g, '').split('/')[0]), motivo: String(r[lc('motivo')] || '') }));
  // OP → rechazos que la explican
  const gOp = {}, lOp = {};
  G.forEach((g, i) => buscar(g.ben, g.fecha.slice(0, 7)).forEach(o => (gOp[o.op] = gOp[o.op] || []).push(i)));
  LB.forEach((l, i) => buscar(l.ben, l.mes).forEach(o => (lOp[o.op] = lOp[o.op] || []).push(i)));
  // Episodios: mismo beneficiario y prestador en meses consecutivos (PAMI emite una OP por mes de internación)
  const grupos = {}; ops.forEach(o => (grupos[o.ben + '|' + o.prest] = grupos[o.ben + '|' + o.prest] || []).push(o));
  const epis = [];
  Object.values(grupos).forEach(lst => { lst.sort((a, b) => ymI(a.per) - ymI(b.per)); let cur = [lst[0]];
    lst.slice(1).forEach(o => { if (ymI(o.per) - ymI(cur[cur.length - 1].per) <= 1) cur.push(o); else { epis.push(cur); cur = [o]; } }); epis.push(cur); });
  const out = [['tipo_fila', 'clasificacion', 'n_rechazo_guardia', 'fecha_rechazo', 'paciente', 'beneficio', 'motivo', 'prestador_que_interno', 'meses_internado', 'prestacion_desde', 'prestacion_hasta', 'op_inicial', 'ops', 'monto_recupero', 'cargo_adm', 'costo_total', 'liquidaciones', 'nota']];
  const usedG = new Set(), usedL = new Set();
  epis.forEach(e => {
    const ngs = [...new Set(e.flatMap(o => (gOp[o.op] || []).filter(i => G[i].vig)))].sort((a, b) => a - b);
    const nls = [...new Set(e.flatMap(o => lOp[o.op] || []))].sort((a, b) => a - b);
    ngs.forEach(i => usedG.add(i)); nls.forEach(i => usedL.add(i));
    const clas = ngs.length ? 'Rechazo conformado (guardia 2026)' : nls.length ? 'Rechazo registrado en libro de guardia' : 'Sin rechazo conformado (de oficio u otro)';
    let r0 = null; for (const o of e) { const g = (gOp[o.op] || []).filter(i => G[i].vig); if (g.length) { r0 = G[Math.min(...g)]; break; } const l = lOp[o.op] || []; if (l.length) { r0 = LB[Math.min(...l)]; break; } }
    const mt = e.reduce((s, o) => s + o.mt, 0), cg = e.reduce((s, o) => s + o.cg, 0);
    const nota = [];
    if (!r0 && mt < 70000) nota.push('Monto menor (< $70.000)');
    if (!r0 && /^no se envio/i.test(e[0].pazNota)) nota.push('Planilla C. Páez: no se envió rechazo desde el Instituto');
    out.push(['Con recupero', clas, ngs.map(i => G[i].n).join(', '), r0 ? r0.fecha : '', r0 ? r0.pac : e[0].pazNom, e[0].ben, r0 ? r0.motivo : '', e[0].prest, e.length, perTxt(e[0].per), perTxt(e[e.length - 1].per),
      e[0].op, e.map(o => o.op).join(' '), Math.round(mt * 100) / 100, Math.round(cg * 100) / 100, Math.round((mt + cg) * 100) / 100, e.map(o => o.lq).join(' '), nota.join('; ')]);
  });
  G.forEach((g, i) => { if (usedG.has(i)) return;
    const nota = maxPer && g.fecha.slice(0, 7) > maxPer ? 'Aún no liquidable (portal hasta prestación ' + perTxt(maxPer) + ')' : !g.ben ? 'Sin nº de beneficio' : 'Sin OP de recupero en el portal';
    out.push(['Sin recupero', 'Rechazo guardia 2026 (' + g.estado + ')', g.n, g.fecha, g.pac, g.ben, g.motivo, '', 0, '', '', '', '', 0, 0, 0, '', nota]); });
  LB.forEach((l, i) => { if (usedL.has(i) || l.mes >= '2026') return;
    out.push(['Sin recupero', 'Rechazo libro de guardia 2025', '', l.fecha, l.pac, l.ben, l.motivo, '', 0, '', '', '', '', 0, 0, 0, '', 'Sin OP de recupero en el portal']); });
  return out;
}
