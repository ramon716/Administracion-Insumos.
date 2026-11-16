/**
 * ADMINISTRACIÓN DE INSUMOS – Rural Bioenergía S.A. (Ypoti)
 * Bajas de insumos (stock) y recepción de compras, conectada a Estansys (API v1).
 *
 * SEGURIDAD: el token de Estansys NO está en este código ni en el HTML. Se guarda solo en
 * Propiedades de la secuencia de comandos (ESTANSYS_TOKEN), que no se publican ni se suben a GitHub.
 * Las llamadas a Estansys las hace siempre el servidor de Google; el celular nunca ve el token.
 *
 * PRIMERA VEZ:
 *  1) En la planilla: Extensiones > Apps Script. Pegá este código en Código.gs y creá el archivo HTML "Index".
 *  2) ⚙️ Configuración del proyecto > Propiedades de la secuencia de comandos > Agregar:
 *       ESTANSYS_TOKEN = (el token que te pasó Estansys)
 *       ESTANSYS_URL   = https://estansys.eon.com.py/api/v1   (opcional, ya viene por defecto)
 *  3) Arriba elegí la función "setup" y tocá ▶ Ejecutar (aceptá los permisos).
 *     Crea las hojas, arma la tabla Almacenes (estancia → almacenes), un activador cada 15 minutos
 *     (reenvíos y estados) y otro diario a las 23 h (copia .xlsx de respaldo en la carpeta "Respaldos Administración de Insumos").
 *  4) Implementar > Nueva implementación > Aplicación web (Ejecutar como: Yo · Acceso: Cualquier usuario).
 *
 * Hojas:
 *  Usuarios     : Usuario | Clave | Nombre | Estancias (TODAS, o códigos/nombres separados por coma) | Rol (admin / capataz / lectura) | Activo
 *  Almacenes    : Estancia código | Estancia | Almacén código | Almacén | Usar (SI/NO)
 *  Bajas        : una fila por producto. Referencia | Fecha | Estancia código | Estancia | Almacén | Producto | Nombre | Unidad | Cantidad | Observación | Usuario | Registrada | Estado | Nº SAP | Motivo rechazo | Actualizado
 *  Recepciones  : una fila por línea. Referencia | Fecha | Orden | Nº OC | Proveedor | Línea | Producto | Nombre | Unidad | Almacén | Pedida | Recibida | Pendiente | Observación línea | Observación | Usuario | Registrada | Estado
 *
 * Estados de una baja: ENVIANDO / SIN ENVIAR (se reintenta sola) / PENDIENTE / APROBADA / EN_SAP / ERROR_SAP / RECHAZADA / ERROR (revisar)
 */

var SH = { U: 'Usuarios', A: 'Almacenes', B: 'Bajas', R: 'Recepciones' };
var HEAD = {
  Usuarios: ['Usuario', 'Clave', 'Nombre', 'Estancias', 'Rol', 'Activo'],
  Almacenes: ['Estancia código', 'Estancia', 'Almacén código', 'Almacén', 'Usar'],
  Bajas: ['Referencia', 'Fecha', 'Estancia código', 'Estancia', 'Almacén', 'Producto', 'Nombre', 'Unidad', 'Cantidad', 'Observación', 'Usuario', 'Registrada', 'Estado', 'Nº SAP', 'Motivo rechazo', 'Actualizado'],
  Recepciones: ['Referencia', 'Fecha', 'Orden', 'Nº OC', 'Proveedor', 'Línea', 'Producto', 'Nombre', 'Unidad', 'Almacén', 'Pedida', 'Recibida', 'Pendiente', 'Observación línea', 'Observación', 'Usuario', 'Registrada', 'Estado']
};
var TZ = 'America/Asuncion', SESSION_SECONDS = 21600;
var ABIERTO = { 'ENVIANDO': 1, 'SIN ENVIAR': 1, 'PENDIENTE': 1, 'APROBADA': 1, 'ERROR_SAP': 1 };

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Administración de Insumos')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------------- utilidades ---------------- */
var _ss = null;
function ss_() { if (!_ss) _ss = SpreadsheetApp.getActiveSpreadsheet(); if (!_ss) throw new Error('Abrí el código desde la planilla (Extensiones > Apps Script)'); return _ss; }
function sheet_(n) {
  var s = ss_().getSheetByName(n);
  if (!s) {
    s = ss_().insertSheet(n);
    s.getRange(1, 1, 1, HEAD[n].length).setValues([HEAD[n]]).setFontWeight('bold').setBackground('#0f3d24').setFontColor('#ffffff');
    s.setFrozenRows(1);
    if (n === SH.B) s.getRange('A:A').setNumberFormat('@');
    if (n === SH.R) s.getRange('A:A').setNumberFormat('@');
  }
  return s;
}
function str_(v) { return v === null || v === undefined ? '' : String(v).trim(); }
function up_(v) { return str_(v).toUpperCase(); }
function num_(v) { var n = Number(String(v).replace(',', '.')); return isNaN(n) ? 0 : n; }
function nrm_(v) { return up_(v).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^A-Z0-9]+/g, ' ').trim(); }
function hoy_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function ahora_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function fmtDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  var s = str_(v), m;
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0, 10);
  if ((m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(s))) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return '';
}
function rows_(name) {
  var s = sheet_(name), last = s.getLastRow();
  if (last < 2) return [];
  return s.getRange(2, 1, last - 1, HEAD[name].length).getValues();
}
function lock_(fn) { var l = LockService.getScriptLock(); l.waitLock(25000); try { return fn(); } finally { l.releaseLock(); } }
function cache_() { return CacheService.getScriptCache(); }
function cj_(k) { var v = cache_().get(k); if (!v) return null; try { return JSON.parse(v); } catch (e) { return null; } }
function cput_(k, o, sec) { try { var s = JSON.stringify(o); if (s.length < 95000) cache_().put(k, s, sec); } catch (e) {} }

/* ---------------- API Estansys ---------------- */
function cfg_() {
  var p = PropertiesService.getScriptProperties();
  var tok = str_(p.getProperty('ESTANSYS_TOKEN'));
  if (!tok) throw new Error('Falta el token de Estansys: cargalo en Configuración del proyecto > Propiedades (ESTANSYS_TOKEN).');
  return { url: (str_(p.getProperty('ESTANSYS_URL')) || 'https://estansys.eon.com.py/api/v1').replace(/\/+$/, ''), tok: tok };
}
/* devuelve {code, body}; code 0 = sin conexión */
function api_(method, path, body) {
  var c = cfg_(), opt = { method: method, muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + c.tok, Accept: 'application/json' } };
  if (body) { opt.contentType = 'application/json'; opt.payload = JSON.stringify(body); }
  var r;
  for (var i = 0; i < 2; i++) {
    try { r = UrlFetchApp.fetch(c.url + path, opt); } catch (e) { return { code: 0, body: null, err: String(e.message || e) }; }
    if (r.getResponseCode() !== 429 || i) break;
    var w = Number((r.getHeaders() || {})['Retry-After']) || 5; Utilities.sleep(Math.min(w, 20) * 1000);
  }
  var t = r.getContentText(), b = null; try { b = JSON.parse(t); } catch (e) {}
  return { code: r.getResponseCode(), body: b };
}
function apiErr_(res) {
  if (!res.code) return 'Sin conexión con Estansys' + (res.err ? ' (' + res.err + ')' : '');
  var e = res.body && res.body.error;
  if (res.code === 401) return 'Estansys rechazó el token (NO_AUTORIZADO). Avisá al administrador.';
  if (!e) return 'Estansys respondió ' + res.code;
  return e.mensaje + (e.detalles && e.detalles.length ? ': ' + e.detalles.map(function (d) { return d.mensaje; }).join(' · ') : '');
}
function get_(path) { var r = api_('get', path); if (r.code !== 200) throw new Error(apiErr_(r)); return r.body; }

/* maestros con caché */
function estancias_() { var k = 'm_est', c = cj_(k); if (c) return c; c = get_('/estancias').datos || []; cput_(k, c, 21600); return c; }
function almacenesApi_() { var k = 'm_alm', c = cj_(k); if (c) return c; c = get_('/almacenes').datos || []; cput_(k, c, 21600); return c; }

/* ---------------- tabla estancia → almacenes ---------------- */
/* sugiere la estancia de cada almacén por el nombre (luego se corrige en la hoja Almacenes) */
var CLAVES_ = [
  ['YBY PORA', ['YBY PORA', 'YBYPOR', 'YBYPO']], ['YBY PYTA', ['YBY PYTA', 'YBYPYT']], ['CERRO MEMBY', ['CERRO MEMBY', 'CMEMB', 'CERROM', 'AGR CM']],
  ['CIELO AZUL', ['CIELO AZUL', 'CAZUL']], ['LA ESPERANZA', ['LA ESPERANZA', 'LAESP']], ['LUSIPAR', ['LUSIPAR', 'LSP', 'FARMACIA']],
  ['ORO VERDE', ['ORO VERDE', 'OROV']], ['SANTA CLARA', ['SANTA CLARA', 'RC SC']], ['SANTA MARIA DA SERRA', ['SANTA MARIA', 'SMSER']],
  ['BELLA VISTA', ['BELA VISTA', 'BELLA VISTA', 'BELAV']], ['YPOTI', ['YPOTI']], ['CONFINAMIENTO', ['CONFINAMIENTO', 'PCONF']]
];
function sugerir_(alm, ests) {
  var t = ' ' + nrm_(alm.codigo) + ' ' + nrm_(alm.nombre) + ' ';
  for (var i = 0; i < CLAVES_.length; i++) {
    var hit = CLAVES_[i][1].some(function (k) { return t.indexOf(' ' + k + ' ') >= 0 || t.indexOf(k) >= 0 && k.length > 4; });
    if (hit) { var e = ests.filter(function (x) { return nrm_(x.nombre).indexOf(CLAVES_[i][0]) >= 0; })[0]; if (e) return e; }
  }
  return null;
}
function armarAlmacenes() {
  var ests = estancias_(), alms = almacenesApi_(), s = sheet_(SH.A), ya = {};
  rows_(SH.A).forEach(function (r) { if (str_(r[2])) ya[up_(r[0]) + '|' + up_(r[2])] = 1; ya['*|' + up_(r[2])] = 1; });
  var nuevos = [];
  alms.forEach(function (a) {
    if (ya['*|' + up_(a.codigo)]) return;
    var e = sugerir_(a, ests);
    nuevos.push(e ? [e.codigo, e.nombre, a.codigo, a.nombre, 'SI'] : ['', '(sin asignar)', a.codigo, a.nombre, 'NO']);
  });
  nuevos.sort(function (x, y) { return x[1] < y[1] ? -1 : x[1] > y[1] ? 1 : 0; });
  if (nuevos.length) s.getRange(s.getLastRow() + 1, 1, nuevos.length, 5).setValues(nuevos);
  s.getRange('A:A').setNumberFormat('@'); s.getRange('C:C').setNumberFormat('@');
  return nuevos.length + ' almacenes agregados';
}
/* {estanciaCodigo: [{codigo,nombre}]} */
function mapa_() {
  var m = {};
  rows_(SH.A).forEach(function (r) {
    if (!str_(r[0]) || !str_(r[2]) || up_(r[4]) === 'NO') return;
    (m[up_(r[0])] = m[up_(r[0])] || []).push({ codigo: str_(r[2]), nombre: str_(r[3]) || str_(r[2]) });
  });
  return m;
}

/* ---------------- sesión ---------------- */
function login(usuario, clave) {
  usuario = str_(usuario).toLowerCase(); clave = str_(clave);
  var u = rows_(SH.U).filter(function (r) { return str_(r[0]).toLowerCase() === usuario && str_(r[1]) === clave && up_(r[5]) !== 'NO'; })[0];
  if (!u) throw new Error('Usuario o contraseña incorrectos');
  var rol = str_(u[4]).toLowerCase();
  var user = {
    usuario: str_(u[0]), nombre: str_(u[2]) || str_(u[0]),
    estancias: up_(u[3]) === 'TODAS' || !str_(u[3]) ? ['*'] : str_(u[3]).split(',').map(nrm_).filter(String),
    rol: rol === 'admin' ? 'admin' : rol === 'lectura' ? 'lectura' : 'capataz'
  };
  var token = Utilities.getUuid();
  cache_().put('t_' + token, JSON.stringify(user), SESSION_SECONDS);
  var out = getData(token); out.token = token; return out;
}
function logout(token) { cache_().remove('t_' + token); return true; }
function auth_(token) {
  var c = cache_(), j = token && c.get('t_' + token);
  if (!j) throw new Error('SESION_VENCIDA');
  c.put('t_' + token, j, SESSION_SECONDS);
  return JSON.parse(j);
}
/* la estancia {codigo,nombre} está asignada al usuario (por código o por nombre, sin acentos) */
function veEst_(user, e) {
  if (user.estancias[0] === '*') return true;
  var c = nrm_(e.codigo), n = nrm_(e.nombre);
  return user.estancias.some(function (x) { return x === c || n === x || n.indexOf(x) >= 0; });
}
function misEst_(user) { return estancias_().filter(function (e) { return veEst_(user, e); }); }
function misAlm_(user) {
  var m = mapa_(), out = {};
  misEst_(user).forEach(function (e) { (m[up_(e.codigo)] || []).forEach(function (a) { out[up_(a.codigo)] = a; }); });
  return out;
}

/* ---------------- lectura ---------------- */
function getData(token) {
  var user = auth_(token), m = mapa_();
  var ests = misEst_(user).map(function (e) { return { codigo: e.codigo, nombre: e.nombre, almacenes: m[up_(e.codigo)] || [] }; });
  var veo = {}; ests.forEach(function (e) { veo[up_(e.codigo)] = 1; });
  var B = {}, ordB = [];
  rows_(SH.B).forEach(function (r) {
    var ref = str_(r[0]); if (!ref || !veo[up_(r[2])]) return;
    if (!B[ref]) { B[ref] = { ref: ref, fecha: fmtDate_(r[1]), est: str_(r[2]), estN: str_(r[3]), alm: str_(r[4]), obs: str_(r[9]), usuario: str_(r[10]), reg: str_(r[11]), estado: str_(r[12]), sap: str_(r[13]), motivo: str_(r[14]), lin: [] }; ordB.push(ref); }
    B[ref].lin.push({ p: str_(r[5]), n: str_(r[6]), u: str_(r[7]), c: num_(r[8]) });
  });
  var almVeo = misAlm_(user), R = {}, ordR = [];
  rows_(SH.R).forEach(function (r) {
    var ref = str_(r[0]); if (!ref) return;
    if (user.rol !== 'admin' && !(str_(r[9]) && almVeo[up_(r[9])])) return;
    if (!R[ref]) { R[ref] = { ref: ref, fecha: fmtDate_(r[1]), orden: str_(r[2]), oc: str_(r[3]), prov: str_(r[4]), obs: str_(r[14]), usuario: str_(r[15]), reg: str_(r[16]), estado: str_(r[17]), lin: [] }; ordR.push(ref); }
    R[ref].lin.push({ id: str_(r[5]), p: str_(r[6]), n: str_(r[7]), u: str_(r[8]), alm: str_(r[9]), ped: num_(r[10]), rec: num_(r[11]), pend: num_(r[12]), obs: str_(r[13]) });
  });
  return {
    user: user, hoy: hoy_(), estancias: ests,
    bajas: ordB.reverse().slice(0, 300).map(function (k) { return B[k]; }),
    recepciones: ordR.reverse().slice(0, 200).map(function (k) { return R[k]; })
  };
}
/* productos con stock en un almacén (todas las páginas, caché 3 min) */
function getProductos(token, almacen) {
  var user = auth_(token), a = misAlm_(user)[up_(almacen)];
  if (!a) throw new Error('No tenés acceso al almacén ' + almacen);
  var k = 'p_' + up_(almacen), c = cj_(k); if (c) return c;
  var out = [], pag = 1, b;
  do {
    b = get_('/productos?almacen=' + encodeURIComponent(a.codigo) + '&pagina=' + pag);
    (b.datos || []).forEach(function (p) { out.push([p.codigo, p.nombre, p.unidad || '', Number(p.stock) || 0]); });
    pag++;
  } while (b.hay_mas && pag <= 30);
  cput_(k, out, 180);
  return out;
}
/* órdenes de compra pendientes, solo las líneas de mis almacenes (admin: todas) */
function ordenes_() { var k = 'oc_pend', c = cj_(k); if (c) return c; c = get_('/ordenes-compra/pendientes').datos || []; cput_(k, c, 60); return c; }
function getOrdenes(token) {
  var user = auth_(token), alm = misAlm_(user), adm = user.rol === 'admin';
  return ordenes_().map(function (o) {
    var L = (o.lineas || []).filter(function (l) { return l.cantidad_pendiente > 0 && (adm || (l.almacen && alm[up_(l.almacen)])); });
    return L.length ? { orden: o.orden, oc: o.numero_oc, prov: o.proveedor, fecha: o.fecha, lineas: L } : null;
  }).filter(Boolean);
}

/* ---------------- BAJAS ---------------- */
function filasDe_(name, ref) {
  var s = sheet_(name), last = s.getLastRow(); if (last < 2) return [];
  var col = s.getRange(2, 1, last - 1, 1).getValues(), out = [];
  for (var i = 0; i < col.length; i++) if (str_(col[i][0]) === ref) out.push(i + 2);
  return out;
}
function setEstado_(name, ref, vals) { /* vals: {col: valor} (columnas 1-based) */
  var s = sheet_(name);
  filasDe_(name, ref).forEach(function (row) { Object.keys(vals).forEach(function (c) { s.getRange(row, Number(c)).setValue(vals[c]); }); });
}
function borrar_(name, ref) { var s = sheet_(name); filasDe_(name, ref).reverse().forEach(function (row) { s.deleteRow(row); }); }

function enviarBaja(token, b) {
  var user = auth_(token);
  if (user.rol === 'lectura') throw new Error('Tu usuario es solo de lectura');
  var ref = str_(b.referencia); if (!/^[A-Za-z0-9._-]{6,100}$/.test(ref)) throw new Error('Referencia inválida');
  var est = misEst_(user).filter(function (e) { return up_(e.codigo) === up_(b.estancia); })[0];
  if (!est) throw new Error('No tenés acceso a esa estancia');
  var alm = (mapa_()[up_(est.codigo)] || []).filter(function (a) { return up_(a.codigo) === up_(b.almacen); })[0];
  if (!alm) throw new Error('El almacén ' + b.almacen + ' no es de ' + est.nombre);
  var fecha = str_(b.fecha); if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha)) throw new Error('Fecha inválida');
  if (fecha > hoy_()) throw new Error('La fecha no puede ser posterior a hoy');
  var lin = (b.lineas || []).map(function (l) { return { producto: str_(l.producto), nombre: str_(l.nombre), unidad: str_(l.unidad), cantidad: Math.round(num_(l.cantidad) * 1e6) / 1e6 }; })
    .filter(function (l) { return l.producto && l.cantidad > 0; });
  if (!lin.length) throw new Error('Agregá al menos un producto con cantidad');
  if (lin.length > 50) throw new Error('Máximo 50 productos por baja');
  var obs = str_(b.observacion).slice(0, 254), quien = (user.nombre + ' (' + user.usuario + ')').slice(0, 100);
  var payload = { referencia: ref, estancia: est.codigo, almacen: alm.codigo, fecha: fecha, usuario: quien, lineas: lin.map(function (l) { return { producto: l.producto, cantidad: l.cantidad }; }) };
  if (obs) payload.observacion = obs;
  lock_(function () {
    if (filasDe_(SH.B, ref).length) return; /* ya estaba guardada: solo se reenvía */
    var now = ahora_(), s = sheet_(SH.B), r0 = s.getLastRow() + 1;
    s.getRange(r0, 1, lin.length, HEAD.Bajas.length).setValues(lin.map(function (l) {
      return [ref, fecha, est.codigo, est.nombre, alm.codigo, l.producto, l.nombre, l.unidad, l.cantidad, obs, quien, now, 'ENVIANDO', '', '', now];
    }));
  });
  return mandarBaja_(ref, payload, true);
}
function mandarBaja_(ref, payload, interactivo) {
  var res = api_('post', '/bajas', payload);
  if (res.code === 200 || res.code === 201) {
    var x = res.body || {};
    setEstado_(SH.B, ref, { 13: x.estado || 'PENDIENTE', 14: x.numero_sap || '', 15: x.motivo_rechazo || '', 16: ahora_() });
    cache_().remove('p_' + up_(payload.almacen));
    return { ok: true, ref: ref, estado: x.estado || 'PENDIENTE', duplicada: !!x.duplicada };
  }
  if (res.code === 400 || res.code === 422) {
    var msg = apiErr_(res);
    if (interactivo) { borrar_(SH.B, ref); throw new Error(msg); }
    setEstado_(SH.B, ref, { 13: 'ERROR', 15: msg.slice(0, 400), 16: ahora_() });
    return { ok: false, ref: ref, error: msg };
  }
  /* sin conexión, límite, conflicto, token o error del servidor: queda guardada y se reintenta */
  setEstado_(SH.B, ref, { 13: 'SIN ENVIAR', 15: apiErr_(res).slice(0, 400), 16: ahora_() });
  return { ok: false, pendiente: true, ref: ref, error: apiErr_(res) };
}
function payloadBaja_(ref) {
  var R = rows_(SH.B).filter(function (r) { return str_(r[0]) === ref; }); if (!R.length) return null;
  var p = { referencia: ref, estancia: str_(R[0][2]), almacen: str_(R[0][4]), fecha: fmtDate_(R[0][1]), usuario: str_(R[0][10]), lineas: R.map(function (r) { return { producto: str_(r[5]), cantidad: num_(r[8]) }; }) };
  if (str_(R[0][9])) p.observacion = str_(R[0][9]);
  return p;
}

/* ---------------- RECEPCIONES ---------------- */
function enviarRecepcion(token, rc) {
  var user = auth_(token);
  if (user.rol === 'lectura') throw new Error('Tu usuario es solo de lectura');
  var ref = str_(rc.referencia); if (!/^[A-Za-z0-9._-]{6,100}$/.test(ref)) throw new Error('Referencia inválida');
  var fecha = str_(rc.fecha); if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha) || fecha > hoy_()) throw new Error('La fecha no puede ser posterior a hoy');
  cache_().remove('oc_pend');
  var yaGuardada = filasDe_(SH.R, ref).length > 0, o = null, lin = [];
  if (!yaGuardada) {
    o = getOrdenes(token).filter(function (x) { return String(x.orden) === String(rc.orden); })[0];
    if (!o) throw new Error('La orden ya no tiene pendientes o no es de tus almacenes. Tocá ↻ para actualizar.');
    var byId = {}; o.lineas.forEach(function (l) { byId[String(l.id_linea)] = l; });
    var vistos = {};
    (rc.lineas || []).forEach(function (l) {
      var k = String(l.id_linea), ol = byId[k]; if (!ol || vistos[k]) return; vistos[k] = 1;
      var c = Math.round(num_(l.cantidad_recibida) * 1e6) / 1e6;
      if (c < 0) throw new Error('Cantidad negativa en ' + ol.nombre);
      if (c > ol.cantidad_pendiente + 1e-9) throw new Error(ol.nombre + ': recibiste ' + c + ' pero solo quedan ' + ol.cantidad_pendiente + ' pendientes');
      lin.push({ ol: ol, c: c, obs: str_(l.observacion).slice(0, 254) });
    });
    if (!lin.length) throw new Error('No hay líneas para registrar');
  }
  var obs = str_(rc.observacion).slice(0, 254), quien = (user.nombre + ' (' + user.usuario + ')').slice(0, 100);
  lock_(function () {
    if (filasDe_(SH.R, ref).length) return;
    var now = ahora_(), s = sheet_(SH.R), r0 = s.getLastRow() + 1;
    s.getRange(r0, 1, lin.length, HEAD.Recepciones.length).setValues(lin.map(function (x) {
      return [ref, fecha, o.orden, o.oc || '', o.prov || '', x.ol.id_linea, x.ol.producto, x.ol.nombre, x.ol.unidad || '', x.ol.almacen || '', x.ol.cantidad_pedida, x.c, Math.round((x.ol.cantidad_pendiente - x.c) * 1e6) / 1e6, x.obs, obs, quien, now, 'ENVIANDO'];
    }));
  });
  return mandarRecep_(ref, payloadRecep_(ref), true);
}
function payloadRecep_(ref) {
  var R = rows_(SH.R).filter(function (r) { return str_(r[0]) === ref; }); if (!R.length) return null;
  var p = { referencia: ref, orden: Number(R[0][2]), fecha: fmtDate_(R[0][1]), usuario: str_(R[0][15]), lineas: R.map(function (r) { var l = { id_linea: Number(r[5]), cantidad_recibida: num_(r[11]) }; if (str_(r[13])) l.observacion = str_(r[13]); return l; }) };
  if (str_(R[0][14])) p.observacion = str_(R[0][14]);
  return p;
}
function mandarRecep_(ref, payload, interactivo) {
  var res = api_('post', '/recepciones', payload);
  if (res.code === 200 || res.code === 201) { setEstado_(SH.R, ref, { 18: 'REGISTRADA' }); cache_().remove('oc_pend'); return { ok: true, ref: ref, duplicada: !!(res.body && res.body.duplicada) }; }
  if (res.code === 400 || res.code === 422) {
    var msg = apiErr_(res);
    if (interactivo) { borrar_(SH.R, ref); throw new Error(msg); }
    setEstado_(SH.R, ref, { 18: ('ERROR: ' + msg).slice(0, 400) });
    return { ok: false, ref: ref, error: msg };
  }
  setEstado_(SH.R, ref, { 18: 'SIN ENVIAR' });
  return { ok: false, pendiente: true, ref: ref, error: apiErr_(res) };
}

/* ---------------- tareas automáticas (activador cada 15 min) ----------------
   1) reenvía lo que quedó SIN ENVIAR   2) trae de Estansys el estado de las bajas abiertas */
function tarea() {
  var hechos = 0;
  var B = {}; rows_(SH.B).forEach(function (r) { var ref = str_(r[0]); if (ref) B[ref] = up_(r[12]); });
  Object.keys(B).forEach(function (ref) {
    if (hechos >= 40) return;
    if (B[ref] === 'SIN ENVIAR' || B[ref] === 'ENVIANDO') { hechos++; mandarBaja_(ref, payloadBaja_(ref), false); }
  });
  var R = {}; rows_(SH.R).forEach(function (r) { var ref = str_(r[0]); if (ref) R[ref] = up_(r[17]); });
  Object.keys(R).forEach(function (ref) {
    if (hechos >= 45) return;
    if (R[ref] === 'SIN ENVIAR' || R[ref] === 'ENVIANDO') { hechos++; mandarRecep_(ref, payloadRecep_(ref), false); }
  });
  Object.keys(B).forEach(function (ref) {
    if (hechos >= 50 || !{ 'PENDIENTE': 1, 'APROBADA': 1, 'ERROR_SAP': 1 }[B[ref]]) return;
    hechos++;
    var res = api_('get', '/bajas/' + encodeURIComponent(ref));
    if (res.code === 200 && res.body && res.body.estado && res.body.estado !== B[ref])
      setEstado_(SH.B, ref, { 13: res.body.estado, 14: res.body.numero_sap || '', 15: res.body.motivo_rechazo || '', 16: ahora_() });
  });
  return hechos;
}
function actualizar(token) { auth_(token); lock_(function () { tarea(); }); return getData(token); }

/* ---------------- instalación ---------------- */
function setup() {
  Object.keys(HEAD).forEach(function (n) { sheet_(n); });
  ss_().setSpreadsheetTimeZone(TZ);
  var U = sheet_(SH.U);
  if (U.getLastRow() < 2) U.getRange(2, 1, 1, 6).setValues([['ramon', 'CAMBIAR123', 'Ramon Sosa', 'TODAS', 'admin', 'SI']]);
  var r = armarAlmacenes();
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'tarea') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('tarea').timeBased().everyMinutes(15).create();
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'respaldoDiario') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('respaldoDiario').timeBased().everyDays(1).atHour(23).create();
  return 'Listo · ' + r + ' · conexión con Estansys OK · respaldo diario a las 23 h';
}

/* ---------------- respaldo diario .xlsx en Drive ----------------
   Copia Bajas, Recepciones y Almacenes (NO la hoja Usuarios, que tiene las claves) a un .xlsx con fecha.
   Guarda los últimos 90 días; los más viejos van a la papelera. */
function carpetaRespaldo_() {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty('RESPALDO_CARPETA'), f = null;
  if (id) { try { f = DriveApp.getFolderById(id); if (f.isTrashed()) f = null; } catch (e) { f = null; } }
  if (!f) { f = DriveApp.createFolder('Respaldos Administración de Insumos'); p.setProperty('RESPALDO_CARPETA', f.getId()); }
  return f;
}
function respaldoDiario() {
  var fecha = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), tmp = SpreadsheetApp.create('tmp_respaldo_insumos_' + fecha);
  try {
    [SH.B, SH.R, SH.A].forEach(function (n) { var c = sheet_(n).copyTo(tmp); c.setName(n); });
    tmp.deleteSheet(tmp.getSheets()[0]);
    SpreadsheetApp.flush();
    var url = 'https://docs.google.com/spreadsheets/d/' + tmp.getId() + '/export?format=xlsx';
    var blob = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob().setName('Insumos_respaldo_' + fecha + '.xlsx');
    var dir = carpetaRespaldo_(), viejo = dir.getFilesByName(blob.getName());
    while (viejo.hasNext()) viejo.next().setTrashed(true);
    dir.createFile(blob);
    var lim = new Date(new Date().getTime() - 90 * 864e5), it = dir.getFiles();
    while (it.hasNext()) { var f = it.next(); if (f.getDateCreated() < lim) f.setTrashed(true); }
  } finally { DriveApp.getFileById(tmp.getId()).setTrashed(true); }
  return 'Respaldo ' + fecha + ' guardado';
}
