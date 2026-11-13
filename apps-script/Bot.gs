/**
 * Bot de WhatsApp: prender y apagar desde la app.
 * La app le pide a la planilla y la planilla le pide al bot (Cloudflare).
 * Así la clave del bot queda guardada acá, nunca en el celular.
 */

const BOT_URL = 'https://whatsapp-agent.pronux-estudio.workers.dev/admin/bot';

function botGuardarClave() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const n = ui.prompt('Bot de WhatsApp (1 de 2)', 'Número de WhatsApp del negocio, solo números (ej. 5491112345678).', ui.ButtonSet.OK_CANCEL);
  if (n.getSelectedButton() !== ui.Button.OK) return;
  const numero = n.getResponseText().replace(/\D/g, '');
  if (numero.length < 10) { ui.alert('Ese número no parece completo.'); return; }
  const t = ui.prompt('Bot de WhatsApp (2 de 2)', 'Clave del bot (te la pasa Pablo).', ui.ButtonSet.OK_CANCEL);
  if (t.getSelectedButton() !== ui.Button.OK) return;
  const token = t.getResponseText().trim();
  if (token.length < 16) { ui.alert('Esa clave es muy corta.'); return; }
  props.setProperties({ BOT_NUMERO: numero, BOT_TOKEN: token });
  try {
    const r = botLlamar_();
    ui.alert('Listo. El bot está ' + (r.on ? 'PRENDIDO' : 'APAGADO') + '.');
  } catch (err) {
    ui.alert('Guardé los datos, pero el bot no respondió: ' + (err.message || err));
  }
}

// on: true / false para cambiar; sin on, solo lee el estado.
function botLlamar_(on) {
  const props = PropertiesService.getScriptProperties();
  const numero = props.getProperty('BOT_NUMERO'), token = props.getProperty('BOT_TOKEN');
  if (!numero || !token) throw new Error('bot_sin_configurar');
  const opt = { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true };
  let url = BOT_URL + '?biz=' + numero;
  if (typeof on === 'boolean') {
    url = BOT_URL;
    opt.method = 'post';
    opt.contentType = 'application/json';
    opt.payload = JSON.stringify({ biz: numero, on: on });
  }
  const res = UrlFetchApp.fetch(url, opt);
  let j = {};
  try { j = JSON.parse(res.getContentText()); } catch (e) {}
  if (!j.ok) throw new Error(res.getResponseCode() === 401 ? 'La clave del bot no es correcta' : (j.error || 'El bot no respondió'));
  return j;
}

// Se engancha desde accionExtra_ (MercadoLibre.gs).
function accionBot_(req) {
  try {
    const r = req.accion === 'bot_cambiar' ? botLlamar_(req.on === true) : botLlamar_();
    return json_({ ok: true, on: r.on });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}
