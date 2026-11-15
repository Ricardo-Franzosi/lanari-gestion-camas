/**
 * Gestión de camas · Instituto Lanari
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

const CONFIG = {
  SHEET_ID: '1L19hyfd_YLGOj_i8JP092XHVeFHASh973aiVOFMozew',
  RECHAZOS_ID: '1qLReNN5bl_zUT3ul8DCuk-BKhg7jiGihm4h56LjlxwM', // hoja "Rechazos guardia Lanari" (cuenta del Lanari, compartida en lectura)
  // Estadísticas: planillas del Departamento de Estadística (recopiladas por Corina Páez) compartidas por Drive.
  // Solo se leen estas carpetas (no se toca ningún otro archivo de Drive).
  ESTAD_CARPETAS: ['1Y7Q5ZOh1NXoGuuYXa_bHGCmNSQpRxc2S', '1XEHbFk3d0JQHLgMr-wigKKwhSipOugrY', '1IsmuQLEc8Re3LoM30PFHKvvhX0NNNNvR'], // Est. especialidad 2024 / 2025 / 2026
  ESTAD_EXTRA: [{ anio: 2024, id: '1R_xKkRv2sf24LNLkEFC-lG3WiGnC4IIyZVXUhljhHbQ', titulo: 'Estadisticas 2024 (base para panel)' }],
  ESTAD_HORA: 6,
  DOMINIO: 'lanari.uba.ar',
  HOJA_ACCESO: 'ACCESO',          // opcional: lista blanca de mails en la columna A
  HOJA_LOG: 'LOG ACCESOS',        // opcional: si existe, se registran ingresos y consultas
  HOJAS_EXCLUIDAS: ['ACCESO', 'LOG ACCESOS'],
  CODIGO_TTL_S: 600,              // 10 minutos
  SESION_TTL_S: 21600,            // 6 horas (máximo de CacheService)
  MAX_INTENTOS: 5,
  MAX_CODIGOS_POR_HORA: 4,
  NOMBRE_REMITENTE: 'Gestión de camas · Instituto Lanari',
};

function doGet() {
  return json_({ ok: true, servicio: 'Gestión de camas Lanari', mensaje: 'Usar POST desde el panel.' });
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
    subject: 'Código de acceso: ' + codigo + ' · Gestión de camas Lanari',
    name: CONFIG.NOMBRE_REMITENTE,
    body: 'Tu código para ingresar al panel de gestión de camas es: ' + codigo +
      '\n\nVence en 10 minutos. Si no lo pediste, ignorá este mensaje.',
    htmlBody: '<p>Tu código para ingresar al panel de gestión de camas es:</p>' +
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
  log_(email, 'consulta');
  return { ok: true, email: email, generado: new Date().toISOString(), hojas: hojas };
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
  log_(email, 'estadisticas');
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
  let iEsp = 0;
  lista.forEach(a => {
    const anioGen = a.anio || (a.titulo.match(/(20\d\d)/) || [])[1] || '';
    const prefijo = a.tipo === 'general' ? 'EST ' + anioGen : 'ESP ' + (++iEsp);
    const firma = a.id + '|' + a.mod.getTime();
    let hojas = null;
    const cambiado = forzar || vistos[a.id] !== firma;
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
        hojas = ss.getSheets().map(sh => ({ nombre: sh.getName(), vals: sh.getDataRange().getValues().map(row => row.map(v => v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : (v === null ? '' : String(v)))) }));
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
  console.log('Archivos: ' + lista.length + ' · pestañas: ' + (meta.length - 1));
  return lista.length;
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
function log_(email, accion) {
  try {
    const hoja = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.HOJA_LOG);
    if (hoja) hoja.appendRow([new Date(), email, accion]);
  } catch (err) { console.warn(err); }
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Ejecutar una vez desde el editor para autorizar los permisos (planilla y envío de mails). */
function autorizar() {
  SpreadsheetApp.openById(CONFIG.SHEET_ID).getName();
  MailApp.getRemainingDailyQuota();
}
