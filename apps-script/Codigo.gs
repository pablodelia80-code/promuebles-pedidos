/**
 * ProMuebles · Pedidos — la planilla en la nube.
 * Se pega en la planilla de Google: Extensiones → Apps Script.
 * Guarda los pedidos de la app y arma los resúmenes diario, semanal y mensual.
 */

const TZ = 'America/Argentina/Buenos_Aires';
const H_PEDIDOS = 'Pedidos';
const H_DIARIO = 'Ventas por día';
const H_SEMANAL = 'Ventas por semana';
const H_MENSUAL = 'Ventas por mes';
const H_MODELOS = 'Modelos más vendidos';
const H_SUSC = 'Avisos (no tocar)';

const MARRON = '#773e13', MARRON_CLARO = '#c8984f', DORADO = '#fdb813', PAPEL = '#fffdf9', SUAVE = '#f6ecdc';
const ORIGENES = ['WhatsApp', 'Mercado Libre', 'Instagram', 'Web'];
const ESTADOS = { pedido: 'Pedido', armado: 'En armado', lista: 'Lista', entregada: 'Entregada' };

// Columnas de la hoja Pedidos (el orden importa)
const COLS = [
  ['id', 'N.º', 70], ['creado', 'Fecha de venta', 110], ['nombre', 'Cliente', 170], ['tel', 'Teléfono', 120],
  ['dir', 'Dirección', 200], ['barrio', 'Barrio', 130], ['zona', 'Zona', 130], ['origen', 'Origen', 115],
  ['modelo', 'Modelo', 100], ['tam', 'Tamaño', 105], ['color', 'Color', 110], ['cierre', 'Cierre suave', 95],
  ['precio', 'Precio', 115], ['pagado', 'Ya pagado', 85], ['extra', 'Extra a cobrar', 110], ['concepto', 'Concepto del extra', 150],
  ['fecha', 'Fecha de entrega', 120], ['estado', 'Estado', 105],
  ['entregado', 'Entregada el', 110], ['uid', 'Código interno', 10]
];
const C = {}; COLS.forEach((c, i) => C[c[0]] = i);

// ---------- Menú ----------
function onOpen() {
  SpreadsheetApp.getUi().createMenu('ProMuebles')
    .addItem('1. Preparar la planilla', 'configurar')
    .addItem('Cambiar la clave de la app', 'cambiarClave')
    .addItem('Actualizar resúmenes', 'actualizarResumenes')
    .addToUi();
}

function configurar() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(TZ);
  ss.setSpreadsheetLocale('es_AR');
  const p = hoja_(H_PEDIDOS);
  if (p.getLastRow() === 0) {
    p.getRange(1, 1, 1, COLS.length).setValues([COLS.map(c => c[1])]);
  }
  migrar_();
  formatoPedidos_(p);
  [H_DIARIO, H_SEMANAL, H_MENSUAL, H_MODELOS].forEach(hoja_);
  const s = hoja_(H_SUSC);
  if (s.getLastRow() === 0) s.getRange(1, 1, 1, 4).setValues([['Dispositivo', 'Alta', 'Endpoint', 'Datos']]);
  s.hideSheet();
  ['Hoja 1', 'Sheet1', 'Hoja1'].forEach(n => { const h = ss.getSheetByName(n); if (h && ss.getSheets().length > 1 && h.getLastRow() === 0) ss.deleteSheet(h); });
  ss.setActiveSheet(p); ss.moveActiveSheet(1);
  actualizarResumenes();
  if (!PropertiesService.getScriptProperties().getProperty('CLAVE')) cambiarClave();
  SpreadsheetApp.getUi().alert('Listo. La planilla quedó preparada.');
}

function cambiarClave() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Clave de la app', 'Escribí la clave que van a usar Pedro y vos para entrar a la app (mínimo 4 caracteres).', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const clave = r.getResponseText().trim();
  if (clave.length < 4) { ui.alert('La clave tiene que tener al menos 4 caracteres.'); return; }
  PropertiesService.getScriptProperties().setProperty('CLAVE', clave);
  ui.alert('Clave guardada.');
}

// ---------- Web app ----------
function doGet() {
  return json_({ ok: true, app: 'ProMuebles Pedidos' });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'Pedido mal armado' }); }
  const clave = PropertiesService.getScriptProperties().getProperty('CLAVE');
  if (!clave || req.clave !== clave) return json_({ ok: false, error: 'clave' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    migrar_();
    switch (req.accion) {
      case 'listar': return json_({ ok: true, pedidos: leerPedidos_() });
      case 'guardar': { const p = guardarPedido_(req.pedido); actualizarResumenes(); return json_({ ok: true, pedido: p, pedidos: leerPedidos_() }); }
      case 'borrar': { borrarPedido_(req.uid); actualizarResumenes(); return json_({ ok: true, pedidos: leerPedidos_() }); }
      case 'suscribir': suscribir_(req.dispositivo, req.suscripcion); return json_({ ok: true });
      case 'avisos': return json_({ ok: true, avisos: avisosDelDia_(), suscripciones: leerSuscripciones_() });
      case 'baja': bajaSuscripcion_(req.endpoint); return json_({ ok: true });
      default: return json_({ ok: false, error: 'Acción desconocida' });
    }
  } finally { lock.releaseLock(); }
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// ---------- Pedidos ----------
function hoja_(n) { const ss = SpreadsheetApp.getActive(); return ss.getSheetByName(n) || ss.insertSheet(n); }
const fmt_ = d => d instanceof Date ? Utilities.formatDate(d, TZ, 'yyyy-MM-dd') : (d ? String(d).slice(0, 10) : '');
const aFecha_ = s => { if (!s) return ''; const [y, m, d] = String(s).split('-').map(Number); return new Date(y, m - 1, d); };
const hoyISO_ = () => Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');

// Planillas armadas antes de las columnas de cobro: las agrega sin perder datos.
function migrar_() {
  const h = hoja_(H_PEDIDOS);
  if (h.getLastRow() === 0 || h.getLastColumn() < 1) return;
  const enc = h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0];
  if (enc.indexOf('Ya pagado') !== -1 || enc.indexOf('Fecha de entrega') !== 13) return;
  h.insertColumnsBefore(14, 3);
  h.getRange(1, 1, 1, COLS.length).setValues([COLS.map(c => c[1])]);
  const n = h.getLastRow() - 1;
  if (n > 0) {
    const org = h.getRange(2, C.origen + 1, n, 1).getValues();
    h.getRange(2, C.pagado + 1, n, 1).setValues(org.map(r => [r[0] === 'Mercado Libre' ? 'Sí' : 'No']));
  }
  formatoPedidos_(h);
}

function leerPedidos_() {
  const h = hoja_(H_PEDIDOS), n = h.getLastRow() - 1;
  if (n < 1) return [];
  return h.getRange(2, 1, n, COLS.length).getValues().filter(r => r[C.uid]).map(r => ({
    uid: String(r[C.uid]), num: r[C.id], creado: fmt_(r[C.creado]), nombre: r[C.nombre], tel: String(r[C.tel]), dir: r[C.dir],
    barrio: r[C.barrio], origen: r[C.origen], modelo: r[C.modelo], tam: r[C.tam], color: r[C.color],
    cierre: r[C.cierre] === 'Sí', precio: Number(r[C.precio]) || 0,
    pagado: r[C.pagado] === 'Sí', extra: Number(r[C.extra]) || 0, concepto: String(r[C.concepto] || ''), fecha: fmt_(r[C.fecha]),
    estado: Object.keys(ESTADOS).find(k => ESTADOS[k] === r[C.estado]) || 'pedido', entregado: fmt_(r[C.entregado])
  }));
}

function filaDe_(p, num, creado, entregadoPrevio) {
  const entregado = p.estado === 'entregada' ? (entregadoPrevio || hoyISO_()) : '';
  const f = [];
  f[C.id] = num; f[C.creado] = aFecha_(creado); f[C.nombre] = p.nombre; f[C.tel] = "'" + p.tel; f[C.dir] = p.dir;
  f[C.barrio] = p.barrio; f[C.zona] = p.zona || ''; f[C.origen] = p.origen; f[C.modelo] = p.modelo; f[C.tam] = p.tam;
  f[C.color] = p.color; f[C.cierre] = p.cierre ? 'Sí' : 'No'; f[C.precio] = Number(p.precio) || 0;
  f[C.pagado] = p.pagado ? 'Sí' : 'No'; f[C.extra] = Number(p.extra) || 0; f[C.concepto] = p.concepto || '';
  f[C.fecha] = aFecha_(p.fecha); f[C.estado] = ESTADOS[p.estado] || 'Pedido'; f[C.entregado] = aFecha_(entregado); f[C.uid] = p.uid;
  return f;
}

function guardarPedido_(p) {
  const h = hoja_(H_PEDIDOS), n = h.getLastRow() - 1;
  const filas = n > 0 ? h.getRange(2, 1, n, COLS.length).getValues() : [];
  const i = filas.findIndex(r => String(r[C.uid]) === String(p.uid));
  if (i >= 0) {
    const r = filas[i];
    h.getRange(i + 2, 1, 1, COLS.length).setValues([filaDe_(p, r[C.id], fmt_(r[C.creado]), fmt_(r[C.entregado]))]);
    p.num = r[C.id];
  } else {
    const num = filas.reduce((m, r) => Math.max(m, Number(r[C.id]) || 0), 0) + 1;
    p.num = num; p.creado = p.creado || hoyISO_();
    h.getRange(n + 2, 1, 1, COLS.length).setValues([filaDe_(p, num, p.creado, '')]);
    formatoPedidos_(h);
  }
  return p;
}

function borrarPedido_(uid) {
  const h = hoja_(H_PEDIDOS), n = h.getLastRow() - 1;
  if (n < 1) return;
  const ids = h.getRange(2, C.uid + 1, n, 1).getValues();
  const i = ids.findIndex(r => String(r[0]) === String(uid));
  if (i >= 0) h.deleteRow(i + 2);
}

function formatoPedidos_(h) {
  const filas = Math.max(h.getMaxRows(), 2);
  h.setFrozenRows(1);
  h.getRange(1, 1, 1, COLS.length).setBackground(MARRON).setFontColor('#ffffff').setFontWeight('bold')
    .setFontFamily('Montserrat').setVerticalAlignment('middle').setWrap(true);
  h.setRowHeight(1, 36);
  COLS.forEach((c, i) => h.setColumnWidth(i + 1, c[2]));
  const datos = h.getRange(2, 1, filas - 1, COLS.length);
  datos.setFontFamily('Poppins').setFontSize(10).setVerticalAlignment('middle');
  h.getRange(2, C.creado + 1, filas - 1, 1).setNumberFormat('dd/mm/yyyy');
  h.getRange(2, C.fecha + 1, filas - 1, 1).setNumberFormat('dd/mm/yyyy');
  h.getRange(2, C.entregado + 1, filas - 1, 1).setNumberFormat('dd/mm/yyyy');
  h.getRange(2, C.precio + 1, filas - 1, 1).setNumberFormat('"$ "#,##0');
  h.getRange(2, C.extra + 1, filas - 1, 1).setNumberFormat('"$ "#,##0');
  h.getRange(2, C.id + 1, filas - 1, 1).setHorizontalAlignment('center');
  h.showColumns(1, COLS.length);
  h.hideColumns(C.uid + 1);
  // Bandas alternadas
  h.getBandings().forEach(b => b.remove());
  h.getRange(1, 1, filas, COLS.length).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
    .setHeaderRowColor(MARRON).setFirstRowColor('#ffffff').setSecondRowColor(SUAVE);
  // Colores del estado
  const est = h.getRange(2, C.estado + 1, filas - 1, 1);
  const reglas = [
    ['Pedido', SUAVE, MARRON], ['En armado', DORADO, MARRON], ['Lista', '#d3f5df', '#127a3b'], ['Entregada', '#ececec', '#6b6b6b']
  ].map(([t, bg, fg]) => SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(t).setBackground(bg).setFontColor(fg).setBold(true).setRanges([est]).build());
  h.setConditionalFormatRules(reglas);
  if (!h.getFilter()) h.getRange(1, 1, filas, COLS.length - 1).createFilter();
}

// ---------- Resúmenes ----------
function actualizarResumenes() {
  const ped = leerPedidos_();
  const semanaDe = iso => { const d = aFecha_(iso); const dow = (d.getDay() + 6) % 7; d.setDate(d.getDate() - dow); return fmt_(d); };
  const mesDe = iso => iso.slice(0, 7);

  const agrupar = keyFn => {
    const g = {};
    ped.forEach(p => {
      if (!p.creado) return;
      const k = keyFn(p.creado);
      const x = g[k] = g[k] || { camas: 0, total: 0, origen: {}, cierre: 0 };
      x.camas++; x.total += p.precio; x.origen[p.origen] = (x.origen[p.origen] || 0) + 1; if (p.cierre) x.cierre++;
    });
    return Object.keys(g).sort().reverse().map(k => [k, g[k]]);
  };
  const filasResumen = (lista, etiqueta) => lista.map(([k, x]) =>
    [etiqueta(k), x.camas, x.total, x.camas ? Math.round(x.total / x.camas) : 0].concat(ORIGENES.map(o => x.origen[o] || 0)));
  const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];

  escribirResumen_(H_DIARIO, 'Ventas por día', 'Día',
    filasResumen(agrupar(k => k), k => aFecha_(k)), 'dddd dd/mm/yyyy');
  escribirResumen_(H_SEMANAL, 'Ventas por semana', 'Semana (desde el lunes)',
    filasResumen(agrupar(semanaDe), k => aFecha_(k)), '"Semana del "dd/mm/yyyy');
  escribirResumen_(H_MENSUAL, 'Ventas por mes', 'Mes',
    filasResumen(agrupar(mesDe), k => { const [y, m] = k.split('-'); return MESES[Number(m) - 1].replace(/^./, c => c.toUpperCase()) + ' ' + y; }), null);

  // Modelos más vendidos, por mes
  const h = hoja_(H_MODELOS); h.clear(); h.getBandings().forEach(b => b.remove());
  const porMes = {};
  ped.forEach(p => { if (!p.creado) return; const m = mesDe(p.creado), k = p.modelo + ' ' + p.tam; porMes[m] = porMes[m] || {}; porMes[m][k] = (porMes[m][k] || 0) + 1; });
  const filas = [];
  Object.keys(porMes).sort().reverse().forEach(m => {
    const [y, mm] = m.split('-'); const nom = MESES[Number(mm) - 1].replace(/^./, c => c.toUpperCase()) + ' ' + y;
    Object.entries(porMes[m]).sort((a, b) => b[1] - a[1]).forEach(([k, n]) => filas.push([nom, k, n]));
  });
  titulo_(h, 'Modelos más vendidos', 3);
  h.getRange(3, 1, 1, 3).setValues([['Mes', 'Modelo y tamaño', 'Camas']]);
  cabecera_(h.getRange(3, 1, 1, 3));
  if (filas.length) {
    h.getRange(4, 1, filas.length, 3).setValues(filas).setFontFamily('Poppins');
    h.getRange(3, 1, filas.length + 1, 3).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
      .setHeaderRowColor(MARRON).setFirstRowColor('#ffffff').setSecondRowColor(SUAVE);
  }
  [130, 200, 80].forEach((w, i) => h.setColumnWidth(i + 1, w));
  h.setFrozenRows(3);
}

function titulo_(h, txt, ancho) {
  h.getRange(1, 1, 1, ancho).merge().setValue(txt).setFontFamily('Montserrat').setFontSize(16).setFontWeight('bold')
    .setFontColor(MARRON).setBackground(PAPEL).setVerticalAlignment('middle');
  h.setRowHeight(1, 40);
  h.getRange(2, 1).setValue('Se actualiza sola con cada pedido que se carga en la app. Las ventas se cuentan por la fecha en que se cargó el pedido.')
    .setFontFamily('Poppins').setFontSize(9).setFontColor(MARRON_CLARO);
}
function cabecera_(r) {
  r.setBackground(MARRON).setFontColor('#ffffff').setFontWeight('bold').setFontFamily('Montserrat').setWrap(true).setVerticalAlignment('middle');
}

function escribirResumen_(nombre, tit, colFecha, filas, formatoFecha) {
  const h = hoja_(nombre); h.clear(); h.getBandings().forEach(b => b.remove());
  const cab = [colFecha, 'Camas vendidas', 'Total vendido', 'Promedio por cama'].concat(ORIGENES);
  titulo_(h, tit, cab.length);
  h.getRange(3, 1, 1, cab.length).setValues([cab]);
  cabecera_(h.getRange(3, 1, 1, cab.length));
  h.setRowHeight(3, 34);
  if (filas.length) {
    h.getRange(4, 1, filas.length, cab.length).setValues(filas).setFontFamily('Poppins').setFontSize(10);
    if (formatoFecha) h.getRange(4, 1, filas.length, 1).setNumberFormat(formatoFecha);
    h.getRange(4, 3, filas.length, 2).setNumberFormat('"$ "#,##0');
    h.getRange(4, 2, filas.length, cab.length - 1).setHorizontalAlignment('center');
    h.getRange(4, 3, filas.length, 2).setHorizontalAlignment('right');
    h.getRange(3, 1, filas.length + 1, cab.length).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
      .setHeaderRowColor(MARRON).setFirstRowColor('#ffffff').setSecondRowColor(SUAVE);
    // Fila de totales
    const t = filas.length + 4;
    h.getRange(t, 1, 1, cab.length).setValues([['Total', '=SUM(B4:B' + (t - 1) + ')', '=SUM(C4:C' + (t - 1) + ')',
      '=IFERROR(C' + t + '/B' + t + ',0)'].concat(ORIGENES.map((o, i) => { const col = String.fromCharCode(69 + i); return '=SUM(' + col + '4:' + col + (t - 1) + ')'; }))]);
    h.getRange(t, 1, 1, cab.length).setFontWeight('bold').setBackground(DORADO).setFontColor(MARRON).setFontFamily('Montserrat');
    h.getRange(t, 3, 1, 2).setNumberFormat('"$ "#,##0').setHorizontalAlignment('right');
    h.getRange(t, 2, 1, 1).setHorizontalAlignment('center');
    h.getRange(t, 5, 1, ORIGENES.length).setHorizontalAlignment('center');
  } else {
    h.getRange(4, 1).setValue('Todavía no hay ventas cargadas.').setFontFamily('Poppins').setFontColor(MARRON_CLARO);
  }
  [190, 95, 120, 120, 95, 110, 95, 80].forEach((w, i) => h.setColumnWidth(i + 1, w));
  h.setFrozenRows(3);
}

// ---------- Avisos al celular ----------
function suscribir_(dispositivo, sub) {
  if (!sub || !sub.endpoint) return;
  const h = hoja_(H_SUSC), n = h.getLastRow() - 1;
  const eps = n > 0 ? h.getRange(2, 3, n, 1).getValues().map(r => r[0]) : [];
  const i = eps.indexOf(sub.endpoint);
  const fila = [dispositivo || 'Celular', new Date(), sub.endpoint, JSON.stringify(sub)];
  if (i >= 0) h.getRange(i + 2, 1, 1, 4).setValues([fila]); else h.appendRow(fila);
}
function leerSuscripciones_() {
  const h = hoja_(H_SUSC), n = h.getLastRow() - 1;
  if (n < 1) return [];
  return h.getRange(2, 1, n, 4).getValues().map(r => { try { return JSON.parse(r[3]); } catch (e) { return null; } }).filter(Boolean);
}
function bajaSuscripcion_(endpoint) {
  const h = hoja_(H_SUSC), n = h.getLastRow() - 1;
  if (n < 1) return;
  const eps = h.getRange(2, 3, n, 1).getValues().map(r => r[0]);
  const i = eps.indexOf(endpoint);
  if (i >= 0) h.deleteRow(i + 2);
}

// Los avisos que salen cada mañana: hoy, mañana y en 3 días.
function avisosDelDia_() {
  const hoy = aFecha_(hoyISO_());
  const enDias = n => { const d = new Date(hoy); d.setDate(d.getDate() + n); return fmt_(d); };
  const cuando = { [enDias(0)]: 'Hoy entregás', [enDias(1)]: 'Mañana entregás', [enDias(3)]: 'En 3 días entregás' };
  const avisos = [];
  leerPedidos_().forEach(p => {
    if (p.estado === 'entregada' || !cuando[p.fecha]) return;
    avisos.push({
      title: cuando[p.fecha],
      body: p.nombre + ' · ' + p.modelo + ' ' + p.tam + ' · entrega en ' + p.barrio + (p.estado !== 'lista' ? ' · todavía figura ' + ESTADOS[p.estado].toLowerCase() : ''),
      uid: p.uid
    });
  });
  return avisos;
}
