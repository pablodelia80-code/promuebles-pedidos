// Corre todas las mañanas (GitHub Actions): arma los avisos de entrega y los manda a los celulares.
//  - 4 días antes: aviso común.
//  - 48 hs antes: aviso con más énfasis (queda en pantalla hasta que se toca).
import webpush from 'web-push';

const { API_URL, APP_CLAVE, VAPID_PUBLIC, VAPID_PRIVATE } = process.env;
if (!API_URL || !APP_CLAVE || !VAPID_PUBLIC || !VAPID_PRIVATE) { console.error('Faltan los secretos del repositorio.'); process.exit(1); }
webpush.setVapidDetails('mailto:pablodelia80@gmail.com', VAPID_PUBLIC, VAPID_PRIVATE);

const llamar = async (accion, extra = {}) => {
  const r = await fetch(API_URL, { method: 'POST', body: JSON.stringify({ clave: APP_CLAVE, accion, ...extra }) });
  const j = await r.json();
  if (!j.ok) throw new Error(j.error || 'Error de la planilla');
  return j;
};

const hoy = new Date().toLocaleDateString('en-CA', { timeZone: 'America/Argentina/Buenos_Aires' });
const dias = iso => Math.round((Date.parse(iso + 'T00:00:00Z') - Date.parse(hoy + 'T00:00:00Z')) / 864e5);
const aCobrar = p => ((p.pagado === undefined ? p.origen === 'Mercado Libre' : p.pagado) ? 0 : Number(p.precio) || 0) + (Number(p.extra) || 0);
const plata = n => '$ ' + n.toLocaleString('es-AR');
const ESTADO = { pedido: 'todavía en pedido', armado: 'todavía en armado' };

const { pedidos } = await llamar('listar');
const { suscripciones } = await llamar('avisos');

const avisos = [];
for (const p of pedidos) {
  if (p.estado === 'entregada' || !p.fecha) continue;
  const d = dias(p.fecha);
  if (d !== 4 && d !== 2) continue;
  const cama = `${p.modelo} ${p.tam}`;
  const aparte = (ESTADO[p.estado] ? ` · ${ESTADO[p.estado]}` : '') + (aCobrar(p) > 0 ? ` · COBRAR ${plata(aCobrar(p))}` : '');
  avisos.push(d === 2
    ? { title: 'ENTREGA EN 48 HS', body: `${p.nombre} · ${cama} · entrega en ${p.barrio}${aparte}`, uid: p.uid, urgente: true }
    : { title: 'Entrega en 4 días', body: `Quedaste con ${p.nombre} · ${cama} · entrega en ${p.barrio}${aparte}`, uid: p.uid });
}

console.log(`${avisos.length} avisos para ${suscripciones.length} celulares`);
for (const sub of suscripciones) {
  for (const a of avisos) {
    try { await webpush.sendNotification(sub, JSON.stringify(a), { TTL: 6 * 3600, urgency: a.urgente ? 'high' : 'normal' }); }
    catch (err) {
      if (err.statusCode === 404 || err.statusCode === 410) { await llamar('baja', { endpoint: sub.endpoint }); console.log('Celular dado de baja (ya no recibe)'); break; }
      console.error('No se pudo enviar:', err.statusCode || err.message);
    }
  }
}
