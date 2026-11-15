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
  log_(email, 'consulta');
  return { ok: true, email: email, generado: new Date().toISOString(), hojas: hojas };
}

function salir_(token) {
  if (token) CacheService.getScriptCache().remove('tok:' + token);
  return { ok: true };
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
