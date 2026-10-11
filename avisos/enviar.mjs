// Corre todas las mañanas (GitHub Actions): pide a la planilla los avisos del día y los manda a los celulares.
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

const { avisos, suscripciones } = await llamar('avisos');
console.log(`${avisos.length} avisos para ${suscripciones.length} celulares`);
for (const sub of suscripciones) {
  for (const a of avisos) {
    try { await webpush.sendNotification(sub, JSON.stringify(a), { TTL: 6 * 3600 }); }
    catch (err) {
      if (err.statusCode === 404 || err.statusCode === 410) { await llamar('baja', { endpoint: sub.endpoint }); console.log('Celular dado de baja (ya no recibe)'); break; }
      console.error('No se pudo enviar:', err.statusCode || err.message);
    }
  }
}
