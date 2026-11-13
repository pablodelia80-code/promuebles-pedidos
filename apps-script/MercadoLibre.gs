/**
 * ProMuebles · Pedidos — conexión de solo lectura con Mercado Libre (API oficial).
 * Se pega en la planilla como archivo nuevo: Extensiones → Apps Script → + → Secuencia de comandos → "MercadoLibre".
 * La Secret Key y los tokens quedan en las propiedades del script, nunca en la planilla ni en el repositorio.
 */

const ML_CLIENT_ID = '4172653903189737';
const ML_REDIRECT = 'https://pablodelia80-code.github.io/promuebles-pedidos/';
const ML_API = 'https://api.mercadolibre.com';
// Solo se pueden consultar estas rutas, y solo con GET.
const ML_RUTAS_OK = /^\/(users\/me|orders\/|messages\/|marketplace\/messages\/|shipments\/|packs\/)/;

// ---------- Menú: conectar ----------
function mlGuardarSecret() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Secret Key de Mercado Libre', 'Pegá la Secret Key de la aplicación "ProMuebles Pedidos" (la ves en developers.mercadolibre.com.ar → Mis aplicaciones).', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const s = r.getResponseText().trim();
  if (s.length < 10) { ui.alert('Eso no parece una Secret Key.'); return; }
  PropertiesService.getScriptProperties().setProperty('ML_SECRET', s);
  ui.alert('Secret Key guardada.');
}

function mlConectar() {
  if (!PropertiesService.getScriptProperties().getProperty('ML_SECRET')) {
    SpreadsheetApp.getUi().alert('Primero guardá la Secret Key (menú ProMuebles → Mercado Libre: guardar Secret Key).');
    return;
  }
  const url = 'https://auth.mercadolibre.com.ar/authorization?response_type=code&client_id=' + ML_CLIENT_ID +
    '&redirect_uri=' + encodeURIComponent(ML_REDIRECT);
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.5">' +
    '<p><b>1.</b> Hacé clic acá y aceptá el permiso (con la cuenta PROMUEBLES1):</p>' +
    '<p><a href="' + url + '" target="_blank">Autorizar en Mercado Libre</a></p>' +
    '<p><b>2.</b> Te va a llevar a la app de pedidos. Mirá la dirección de arriba del navegador: tiene <code>?code=TG-...</code>. Copiá todo lo que sigue a <code>code=</code> (hasta el final o hasta un &amp;).</p>' +
    '<p><b>3.</b> Pegalo acá:</p>' +
    '<input id="c" style="width:100%;padding:6px" placeholder="TG-..."/>' +
    '<p><button onclick="var c=document.getElementById(\'c\').value;document.getElementById(\'m\').innerText=\'Conectando...\';' +
    'google.script.run.withSuccessHandler(function(t){document.getElementById(\'m\').innerText=t}).withFailureHandler(function(e){document.getElementById(\'m\').innerText=\'Error: \'+e.message}).mlCanjear(c)">Conectar</button></p>' +
    '<p id="m" style="color:#127a3b"></p></div>'
  ).setWidth(520).setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, 'Conectar Mercado Libre');
}

function mlCanjear(code) {
  code = String(code || '').trim();
  if (!code) throw new Error('Pegá el código.');
  const j = mlPedirToken_({ grant_type: 'authorization_code', code: code, redirect_uri: ML_REDIRECT });
  mlGuardarToken_(j);
  const me = mlGet_('/users/me');
  return 'Conectado como ' + (me.datos && me.datos.nickname ? me.datos.nickname : '(sin nombre)') + '. Ya podés cerrar esta ventana.';
}

// ---------- Tokens ----------
function mlPedirToken_(extra) {
  const secret = PropertiesService.getScriptProperties().getProperty('ML_SECRET');
  if (!secret) throw new Error('Falta la Secret Key.');
  const r = UrlFetchApp.fetch(ML_API + '/oauth/token', {
    method: 'post', contentType: 'application/x-www-form-urlencoded', headers: { accept: 'application/json' },
    payload: Object.assign({ client_id: ML_CLIENT_ID, client_secret: secret }, extra), muteHttpExceptions: true
  });
  let j = {}; try { j = JSON.parse(r.getContentText()); } catch (e) { }
  if (r.getResponseCode() !== 200) throw new Error('Mercado Libre respondió ' + r.getResponseCode() + ': ' + (j.message || j.error || 'sin detalle'));
  return j;
}

function mlGuardarToken_(j) {
  const p = PropertiesService.getScriptProperties();
  p.setProperty('ML_ACCESS', j.access_token);
  p.setProperty('ML_EXPIRA', String(Date.now() + (Number(j.expires_in) || 21600) * 1000));
  if (j.refresh_token) p.setProperty('ML_REFRESH', j.refresh_token); // es de un solo uso: siempre guardar el nuevo
}

function mlToken_() {
  const p = PropertiesService.getScriptProperties();
  const refresh = p.getProperty('ML_REFRESH');
  if (!refresh) throw new Error('Mercado Libre no está conectado.');
  if (p.getProperty('ML_ACCESS') && Date.now() < (Number(p.getProperty('ML_EXPIRA')) || 0) - 300000) return p.getProperty('ML_ACCESS');
  mlGuardarToken_(mlPedirToken_({ grant_type: 'refresh_token', refresh_token: refresh }));
  return p.getProperty('ML_ACCESS');
}

// ---------- Consulta (solo lectura) ----------
function mlGet_(ruta) {
  if (!ML_RUTAS_OK.test(String(ruta || ''))) throw new Error('Ruta no permitida.');
  const r = UrlFetchApp.fetch(ML_API + ruta, { method: 'get', headers: { Authorization: 'Bearer ' + mlToken_() }, muteHttpExceptions: true });
  let d = null; try { d = JSON.parse(r.getContentText()); } catch (e) { d = r.getContentText().slice(0, 500); }
  return { codigo: r.getResponseCode(), datos: d };
}

// ---------- Ventas de Mercado Libre → pedidos ----------
const ML_DESDE = '2026-10-11'; // solo ventas desde esta fecha (las anteriores ya se cargaron a mano)
const ML_COLORES = ['Blanco', 'Wengue', 'Roble Kendall', 'Nogal Pacífico', 'Negro', 'Gris Perla', 'Gris Sombra', 'Kentoky'];
const mlSinTilde_ = s => String(s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
const mlNombre_ = s => String(s || '').replace(/\s+/g, ' ').trim().toLowerCase().replace(/(^|[\s\-])\S/g, c => c.toUpperCase());

// Lo que escribió el comprador en el chat de la venta (texto completo) y el celular, si lo dejó.
function mlChat_(msgs, seller) {
  const textos = (msgs || []).filter(m => m.from && String(m.from.user_id) !== String(seller))
    .sort((a, b) => String((a.message_date || {}).created).localeCompare(String((b.message_date || {}).created)))
    .map(m => String(m.text || '').replace(/\s+/g, ' ').trim()).filter(Boolean);
  const todo = textos.join(' ');
  const m = todo.match(/(?:\+?54\s?9?\s?)?\b(?:11|15)[\s\-]?\d{4}[\s\-]?\d{4}\b/);
  return { chat: textos.join(' | ').slice(0, 800), tel: m ? m[0].replace(/\D/g, '') : '' };
}

function mlBloque_(id, chat, revisar) {
  return '[ML] Venta #' + id + '. ATENCIÓN: falta confirmar la dirección de entrega. ' +
    (chat ? 'Chat del comprador: ' + chat : 'El comprador todavía no escribió.') +
    (revisar.length ? ' Revisar: ' + revisar.join(', ') + '.' : '') + ' [/ML]';
}

function mlArmarPedido_(o, seller, msgs) {
  const it = (o.order_items || [])[0] || {}, item = it.item || {}, tit = String(item.title || '');
  const attrs = {}; (item.variation_attributes || []).forEach(a => { attrs[mlSinTilde_(a.name)] = a.value_name; });
  const caj = (tit.match(/(\d+)\s*cajones/i) || [])[1];
  let tam = /king/i.test(tit) ? 'King 180' : /queen/i.test(tit) ? 'Queen' : /plaza\s*y\s*media/i.test(tit) ? 'Plaza y Media' : '2 Plazas';
  let modelo = '';
  if (/estantes/i.test(tit)) { modelo = 'Estantes Vip'; tam = '2 Plazas'; }
  else if (caj) modelo = (caj === '6' && tam === 'Plaza y Media') ? 'Esquinero' : caj + ' Vip';
  const buscar = v => ML_COLORES.find(c => mlSinTilde_(c) === mlSinTilde_(v));
  const color = buscar(attrs['nombre del diseno']) || buscar(attrs['color']) || attrs['color'] || '';
  const revisar = [];
  if (!modelo) revisar.push('modelo');
  if (!color) revisar.push('color');
  const creado = Utilities.formatDate(new Date(o.date_created), TZ, 'yyyy-MM-dd');
  const entrega = aFecha_(creado); entrega.setDate(entrega.getDate() + 7);
  const cobrado = (o.payments || []).filter(p => p.status === 'approved').reduce((s, p) => s + (Number(p.total_paid_amount) || 0), 0) || Number(o.paid_amount) || 0;
  const b = o.buyer || {}, c = mlChat_(msgs, seller);
  return {
    uid: 'ml' + o.id, creado: creado, nombre: mlNombre_((b.first_name || '') + ' ' + (b.last_name || '')) || b.nickname || 'Comprador de ML',
    tel: c.tel, dir: '', barrio: '', zona: '', nota: mlBloque_(o.id, c.chat, revisar), origen: 'Mercado Libre',
    modelo: modelo, tam: tam, color: color, cierre: false, precio: cobrado, pagado: true, extra: 0, concepto: '',
    fecha: fmt_(entrega), estado: 'pedido', entregado: ''
  };
}

function mlMensajes_(o, seller) {
  const m = mlGet_('/messages/packs/' + (o.pack_id || o.id) + '/sellers/' + seller + '?tag=post_sale');
  return (m.codigo === 200 && m.datos && m.datos.messages) || [];
}

// Carga las ventas nuevas y, en los pedidos de ML que siguen sin dirección, actualiza lo que escribió el comprador.
// soloVista = true: no escribe nada, solo devuelve lo que haría.
function mlSincronizar_(soloVista) {
  const desde = PropertiesService.getScriptProperties().getProperty('ML_DESDE') || ML_DESDE;
  const seller = mlGet_('/users/me').datos.id;
  const r = mlGet_('/orders/search?seller=' + seller + '&sort=date_desc&limit=50');
  if (r.codigo !== 200) throw new Error('Mercado Libre respondió ' + r.codigo + ' al pedir las ventas.');
  const existentes = leerPedidos_(), ya = {};
  existentes.forEach(p => { ya[p.uid] = 1; });
  const nuevos = [], actualizados = [];
  (r.datos.results || []).slice().reverse().forEach(o => {
    if (o.status !== 'paid') return;
    if (Utilities.formatDate(new Date(o.date_created), TZ, 'yyyy-MM-dd') < desde) return;
    if (ya['ml' + o.id]) return;
    const det = mlGet_('/orders/' + o.id); // el listado trae solo el usuario; el detalle trae nombre y apellido
    const ped = mlArmarPedido_(det.codigo === 200 && det.datos ? det.datos : o, seller, mlMensajes_(o, seller));
    nuevos.push(ped);
    if (!soloVista) guardarPedido_(ped);
  });
  existentes.forEach(p => {
    if (!/^ml\d+$/.test(p.uid) || p.estado === 'entregada' || p.dir) return;
    const bloque = /\[ML\][\s\S]*?\[\/ML\]/;
    if (!bloque.test(p.nota)) return;
    const id = p.uid.slice(2), c = mlChat_(mlMensajes_({ id: id }, seller), seller);
    const nuevo = mlBloque_(id, c.chat, []);
    if (nuevo === (p.nota.match(bloque) || [''])[0] && (!c.tel || p.tel)) return;
    actualizados.push({ uid: p.uid, nombre: p.nombre });
    if (!soloVista) { p.nota = p.nota.replace(bloque, nuevo); if (!p.tel && c.tel) p.tel = c.tel; guardarPedido_(p); }
  });
  if (!soloVista && (nuevos.length || actualizados.length)) actualizarResumenes();
  return { nuevos: nuevos, actualizados: actualizados };
}

function mlRevisionAutomatica() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { migrar_(); mlSincronizar_(false); } catch (e) { console.error(e); } finally { lock.releaseLock(); }
}

function mlRevisarAhora() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    migrar_();
    const r = mlSincronizar_(false);
    SpreadsheetApp.getUi().alert('Listo. Ventas nuevas cargadas: ' + r.nuevos.length + '. Pedidos actualizados con el chat: ' + r.actualizados.length + '.');
  } finally { lock.releaseLock(); }
}

function mlActivarRevision() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'mlRevisionAutomatica').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('mlRevisionAutomatica').timeBased().everyHours(1).create();
  SpreadsheetApp.getUi().alert('Listo. La planilla va a revisar Mercado Libre una vez por hora.');
}

// Acciones extra de la web app (se enganchan desde doPost en Codigo.gs).
function accionExtra_(req) {
  try {
    switch (req.accion) {
      case 'ml_vista': return json_({ ok: true, resultado: mlSincronizar_(true) });
      case 'bot_estado': case 'bot_cambiar': return accionBot_(req);
      case 'ml_sync': { const r = mlSincronizar_(false); return json_({ ok: true, nuevos: r.nuevos.length, actualizados: r.actualizados.length, pedidos: leerPedidos_() }); }
    }
  } catch (err) { return json_({ ok: false, error: String(err.message || err) }); }
  return json_({ ok: false, error: 'Acción desconocida' });
}
