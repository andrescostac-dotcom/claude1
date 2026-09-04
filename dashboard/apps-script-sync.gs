/**
 * CosCor — Sync Meta Ads + Kommo CRM into this Sheet, daily.
 *
 * SETUP (hacer una sola vez):
 * 1. En este Sheet: Extensiones > Apps Script.
 * 2. Borrar el contenido de Code.gs y pegar este archivo completo.
 * 3. Ir a "Configuración del proyecto" (ícono de engranaje, panel izquierdo) >
 *    "Propiedades del script" > agregar estas 4 propiedades:
 *      META_ACCESS_TOKEN     -> el token de Meta
 *      META_AD_ACCOUNT_ID    -> act_2636529573412003
 *      KOMMO_SUBDOMAIN       -> infocoscorlife.kommo.com
 *      KOMMO_ACCESS_TOKEN    -> el token de Kommo
 * 4. En el editor, seleccionar la función "setup" en el dropdown de arriba y
 *    apretar "Ejecutar". Va a pedir autorización (Google avisa que es un script
 *    no verificado — es normal para scripts propios, click en "Avanzado" >
 *    "Ir a [nombre del proyecto] (no seguro)" > Permitir).
 *    Esto corre la primera sincronización Y deja instalado el disparador diario.
 * 5. Listo. Todos los días a la hora elegida (ver DAILY_HOUR abajo) se va a
 *    actualizar solo. Para forzar una actualización manual, ejecutar "syncAll".
 */

const DAILY_HOUR = 8; // hora (0-23, horario de Argentina) a la que corre la sincronización diaria

const PIPELINE_ID = 13684663;
const WON_ID = 142;
const LOST_ID = 143;
const QUALIFIED_IDS = [105609867, 109532768, 105671691, 105671695, WON_ID]; // visita-reunion, reunion realizada, 2da reunion, negociacion, + ganados
const EXCLUDE_TAGS = ['Difusión', 'Follow-up 1', 'JN', 'WA', 'Interes Futuro', 'Apta Credito', 'Presu menos 300', 'Barrio Cerrado', 'Inmobiliaria'];
const TAG_ALIASES = { 'Difusion Misiones': 'Misiones' };
const ADSET_TO_DEV = { 'FAMILIA_CH': 'Chubut', 'FAMILIA_SI': 'Simón Iriondo', 'FAMILIA_MIS': 'Misiones' };
// Meta no deja pedir insights de más de 37 meses atrás (error #3018) — 36 es el máximo
// histórico posible, calculado siempre relativo a "hoy" para que nunca se pase del límite
// (a diferencia de una fecha fija, que eventualmente lo superaría con el paso del tiempo).
const META_MONTHS_BACK = 36;

function props() { return PropertiesService.getScriptProperties(); }

// Reintenta un UrlFetchApp.fetch() ante fallos transitorios -- ej. "Address unavailable"
// (un hiccup momentáneo de conexión de Google hacia el servidor de destino, no un error de
// nuestro código ni de los datos) o un 429/5xx de la API. Sin esto, cualquier request de las
// ~decenas que hace un sync se cortaba y tiraba abajo syncAll() entero para ese día (visto en
// vivo el 3/9: "Exception: Address unavailable" en pleno fetchContactsById). El sheet no se
// pierde en ese caso -- sheet.clear() corre recién al final, después de todos los fetches -- pero
// el sync de ese día directamente no llegaba a correr. Reintenta con backoff (1s, 2s, 4s) antes
// de rendirse y dejar que la excepción se propague (así una falla real y persistente sigue
// avisando por mail, como siempre hace Apps Script con un trigger que tira excepción).
function fetchWithRetry(url, options) {
  const MAX_ATTEMPTS = 4;
  let lastError;
  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    try {
      const resp = options ? UrlFetchApp.fetch(url, options) : UrlFetchApp.fetch(url);
      const code = resp.getResponseCode();
      // Un 429/5xx solo llega hasta acá si el caller pidió muteHttpExceptions:true (si no,
      // UrlFetchApp ya tiró la excepción que atajamos abajo) -- son errores transitorios de la
      // API, vale la pena reintentar en vez de devolver la respuesta de error tal cual.
      if (code === 429 || code >= 500) {
        lastError = new Error(`HTTP ${code} de ${url.split('?')[0]}`);
      } else {
        return resp;
      }
    } catch (e) {
      lastError = e;
    }
    if (attempt < MAX_ATTEMPTS) Utilities.sleep(1000 * Math.pow(2, attempt - 1)); // 1s, 2s, 4s
  }
  throw lastError;
}

function setup() {
  syncAll();
  installDailyTrigger();
}

function installDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncAll') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncAll')
    .timeBased()
    .atHour(DAILY_HOUR)
    .everyDays(1)
    .inTimezone('America/Argentina/Buenos_Aires')
    .create();
}

function syncAll() {
  syncKommoLeads();
  syncMetaDaily();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const meta = ss.getSheetByName('Meta') || ss.insertSheet('Meta');
  meta.getRange('A1').setValue('Última actualización');
  meta.getRange('A2').setValue(new Date());
}

function canonicalTag(t) { return TAG_ALIASES[t] || t; }

// --- DEBUG: pega el ID de un lead que en Kommo muestre "Información rastreada"
// (lo ves en la URL cuando abrís ese lead, ej. .../leads/detail/1243840) y corré esta
// función una vez desde el editor. Mirá el resultado en Ver > Registros (o Ctrl+Enter).
// Pegame acá lo que imprima para que ajuste extractCustomField con el campo real.
function debugLead() {
  const LEAD_ID = 22689017;
  const subdomain = props().getProperty('KOMMO_SUBDOMAIN');
  const token = props().getProperty('KOMMO_ACCESS_TOKEN');
  const headers = { Authorization: 'Bearer ' + token };
  const resp = fetchWithRetry(
    `https://${subdomain}/api/v4/leads/${LEAD_ID}?with=source_id,catalog_elements`,
    { headers, muteHttpExceptions: true });
  Logger.log('HTTP ' + resp.getResponseCode());
  Logger.log(resp.getContentText());
}

// --- DEBUG: "Casa Visitada" (2444391) y "Fecha Visita" (576940) salieron vacíos en TODOS
// los leads (incluso en uno "Logrado con éxito"), a pesar de que el usuario confirmó esos IDs
// como correctos. Esta función imprime:
// 1) el custom_fields_values completo del lead (con field_id de cada uno, para ver si 2444391/
//    576940 aparecen ahí con o sin valor).
// 2) lo mismo para cada CONTACTO vinculado al lead — por si esos 2 campos viven en el contacto
//    y no en el lead (posible explicación de por qué nunca aparecen del lado del lead).
// Correr una vez desde el editor y pegarme el resultado (Ver > Registros o Ctrl+Enter).
function debugCustomFields() {
  const LEAD_ID = 6441028; // el usuario confirmó que este lead SÍ tiene Casa Visitada / Fecha Visita cargadas
  const subdomain = props().getProperty('KOMMO_SUBDOMAIN');
  const token = props().getProperty('KOMMO_ACCESS_TOKEN');
  const headers = { Authorization: 'Bearer ' + token };

  const resp = fetchWithRetry(
    `https://${subdomain}/api/v4/leads/${LEAD_ID}?with=contacts`,
    { headers, muteHttpExceptions: true });
  Logger.log('HTTP lead: ' + resp.getResponseCode());
  const lead = JSON.parse(resp.getContentText());
  Logger.log('--- LEAD custom_fields_values ---');
  Logger.log(JSON.stringify(lead.custom_fields_values, null, 2));

  const contacts = (lead._embedded && lead._embedded.contacts) || [];
  Logger.log(`--- ${contacts.length} contacto(s) vinculado(s) ---`);
  for (const c of contacts) {
    const cResp = fetchWithRetry(
      `https://${subdomain}/api/v4/contacts/${c.id}`,
      { headers, muteHttpExceptions: true });
    Logger.log(`HTTP contacto ${c.id}: ` + cResp.getResponseCode());
    const contact = JSON.parse(cResp.getContentText());
    Logger.log(`--- CONTACTO ${c.id} custom_fields_values ---`);
    Logger.log(JSON.stringify(contact.custom_fields_values, null, 2));
  }
}

// --- DEBUG: confirmar qué action_type de Meta corresponde a "Conversaciones por mensajes
// iniciados" (la métrica que se ve en Meta Ads Manager). Trae los últimos 7 días de la campaña
// WHATSAPP con TODOS los action_type que devuelve Meta (no solo los que usa el dashboard) y
// suma cada uno por separado. Correr una vez, mirar el log, y comparar la suma de
// "onsite_conversion.messaging_conversation_started_7d" (la que ahora usa el dashboard) contra
// lo que Meta Ads Manager muestra como "Conversaciones por mensajes iniciados" para la campaña
// WHATSAPP en los últimos 7 días -- si no coincide, pegarme el log completo para ajustar la
// clave correcta en syncMetaDaily().
function debugMetaActions() {
  const token = props().getProperty('META_ACCESS_TOKEN');
  const adAccount = props().getProperty('META_AD_ACCOUNT_ID');
  const until = new Date();
  const since = new Date();
  since.setDate(since.getDate() - 7);
  const fmt = d => Utilities.formatDate(d, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd');
  const filtering = JSON.stringify([{ field: 'campaign.name', operator: 'CONTAIN', value: 'WHATSAPP' }]);
  const url = `https://graph.facebook.com/v19.0/${adAccount}/insights` +
    `?level=adset&time_range=${encodeURIComponent(JSON.stringify({ since: fmt(since), until: fmt(until) }))}` +
    `&filtering=${encodeURIComponent(filtering)}` +
    `&fields=campaign_name,adset_name,spend,actions&limit=100&access_token=${encodeURIComponent(token)}`;
  const resp = JSON.parse(fetchWithRetry(url).getContentText());
  const data = resp.data || [];
  Logger.log(`--- últimos 7 días, campaña WHATSAPP, ${data.length} fila(s) (1 por conjunto de anuncios) ---`);
  const totals = {};
  for (const r of data) {
    Logger.log(`${r.adset_name}: ${JSON.stringify(r.actions)}`);
    (r.actions || []).forEach(a => { totals[a.action_type] = (totals[a.action_type] || 0) + parseFloat(a.value); });
  }
  Logger.log('--- TOTALES por action_type (últimos 7 días, toda la campaña WHATSAPP) ---');
  Logger.log(JSON.stringify(totals, null, 2));
}

// Busca en custom_fields_values del lead un campo cuyo nombre o código contenga alguno de
// los patrones dados (comparación insensible a mayúsculas/espacios/guiones), y devuelve su
// primer valor. Kommo suele guardar los UTM de un lead (los que trajo el clic del anuncio)
// como campos personalizados llamados "UTM Campaign" / "UTM Content" (o utm_campaign /
// utm_content según cómo esté configurada la cuenta) — por eso el match es flexible.
function extractCustomField(customFields, patterns) {
  if (!customFields) return '';
  for (const f of customFields) {
    const name = String(f.field_name || '').toLowerCase().replace(/[\s_-]/g, '');
    const code = String(f.field_code || '').toLowerCase().replace(/[\s_-]/g, '');
    for (const p of patterns) {
      if (name.indexOf(p) !== -1 || code.indexOf(p) !== -1) {
        const v = f.values && f.values[0] && f.values[0].value;
        if (v) return String(v);
      }
    }
  }
  return '';
}

// IDs de campos personalizados de Kommo (Ajustes > Campos personalizados > Leads).
// A diferencia de extractCustomField (matching por nombre, flexible pero ambiguo), acá el ID
// es exacto — se pidieron directamente al usuario.
const FIELD_CASA_VISITADA = 2444391;
const FIELD_FUENTE = 578180;
const FIELD_FECHA_VISITA = 576940;
const FIELD_PHONE = 559132; // campo "Phone" del CONTACTO (visto en debugCustomFields)

// Devuelve el/los valor(es) de un campo personalizado por field_id, como string. Sirve tanto
// para campos de texto/selección simple (values[0].value) como multiselección (values.length > 1
// -> se unen con ", ", igual que se hace con los tags de desarrollo).
function extractCustomFieldById(customFields, fieldId) {
  if (!customFields) return '';
  // Number(...) por las dudas de que la API devuelva field_id como string en algún caso —
  // === estricto entre number y string nunca matchea y da falso "campo vacío".
  const f = customFields.find(cf => Number(cf.field_id) === fieldId);
  if (!f || !f.values) return '';
  return f.values
    .map(v => v.value)
    .filter(v => v !== undefined && v !== null && v !== '')
    .join(', ');
}

// Igual que extractCustomFieldById pero para campos tipo fecha: Kommo guarda esos campos como
// timestamp unix (segundos). Devuelve un objeto Date de Apps Script, o '' si no hay valor.
function extractCustomFieldDateById(customFields, fieldId) {
  if (!customFields) return '';
  const f = customFields.find(cf => Number(cf.field_id) === fieldId);
  if (!f || !f.values || !f.values[0]) return '';
  const raw = f.values[0].value;
  if (raw === undefined || raw === null || raw === '') return '';
  const n = Number(raw);
  if (!isNaN(n) && n > 1000000000) return new Date(n * 1000);
  const parsed = new Date(raw);
  return isNaN(parsed.getTime()) ? '' : parsed;
}

// El campo Phone puede traer varios números (trabajo, celular, etc.), cada uno con enum_code.
// Preferimos el marcado como celular/móvil si existe; si no, el primero que haya.
function extractPhone(customFields) {
  if (!customFields) return '';
  const f = customFields.find(cf => Number(cf.field_id) === FIELD_PHONE);
  if (!f || !f.values || !f.values.length) return '';
  const mobile = f.values.find(v => /mob|cel/i.test(v.enum_code || ''));
  return (mobile || f.values[0]).value || '';
}

// Trae custom_fields_values de varios contactos de una — Kommo permite filtrar por múltiples
// id con filter[id][]=... — en tandas (para no armar URLs demasiado largas) de a BATCH.
// Confirmado con el usuario (23/8): "Casa Visitada" y "Fecha Visita" viven en el CONTACTO
// vinculado al lead, no en el lead — por eso hace falta este segundo fetch (a diferencia de
// Fuente y los UTM, que sí están en custom_fields_values del lead).
function fetchContactsById(ids) {
  const subdomain = props().getProperty('KOMMO_SUBDOMAIN');
  const token = props().getProperty('KOMMO_ACCESS_TOKEN');
  const headers = { Authorization: 'Bearer ' + token };
  const byId = {};
  // 150 tiraba "Limit Exceeded: URLFetch URL Length" (cada filter[id][]=NNNNNNN suma ~13-22
  // caracteres a la URL) -- 50 deja margen de sobra bajo el límite de UrlFetchApp.
  const BATCH = 50;
  const idList = Array.from(ids);
  for (let i = 0; i < idList.length; i += BATCH) {
    const batch = idList.slice(i, i + BATCH);
    const filterParams = batch.map(id => `filter[id][]=${id}`).join('&');
    let page = 1;
    while (true) {
      const resp = fetchWithRetry(
        `https://${subdomain}/api/v4/contacts?${filterParams}&limit=250&page=${page}`,
        { headers, muteHttpExceptions: true });
      if (resp.getResponseCode() === 204) break;
      const data = JSON.parse(resp.getContentText());
      const contacts = (data._embedded && data._embedded.contacts) || [];
      if (!contacts.length) break;
      for (const c of contacts) byId[c.id] = { name: c.name || '', customFields: c.custom_fields_values };
      if (!data._links || !data._links.next) break;
      page++;
    }
  }
  return byId;
}

function syncKommoLeads() {
  const subdomain = props().getProperty('KOMMO_SUBDOMAIN');
  const token = props().getProperty('KOMMO_ACCESS_TOKEN');
  const headers = { Authorization: 'Bearer ' + token };

  // pipeline statuses (for names + sort order)
  const pipeline = JSON.parse(fetchWithRetry(
    `https://${subdomain}/api/v4/leads/pipelines/${PIPELINE_ID}`, { headers }).getContentText());
  const statuses = {};
  pipeline._embedded.statuses.forEach(s => { statuses[s.id] = s.name; });

  // Fase 1: traer todos los leads (con su contacto principal embebido vía with=contacts) y
  // armar filas parciales -- todavía sin casa_visitada/fecha_visita, que dependen del contacto.
  const partial = [];
  const contactIds = new Set();
  let page = 1;
  while (true) {
    const resp = fetchWithRetry(
      `https://${subdomain}/api/v4/leads?with=contacts&limit=250&page=${page}`,
      { headers, muteHttpExceptions: true });
    if (resp.getResponseCode() === 204) break;
    const data = JSON.parse(resp.getContentText());
    const leads = (data._embedded && data._embedded.leads) || [];
    if (!leads.length) break;
    for (const l of leads) {
      const rawTags = (l._embedded && l._embedded.tags || []).map(t => t.name);
      const devTags = [...new Set(rawTags.filter(t => !EXCLUDE_TAGS.includes(t)).map(canonicalTag))];
      const utmCampaign = extractCustomField(l.custom_fields_values, ['utmcampaign']);
      const utmContent = extractCustomField(l.custom_fields_values, ['utmcontent']);
      const fuente = extractCustomFieldById(l.custom_fields_values, FIELD_FUENTE);
      const leadContacts = (l._embedded && l._embedded.contacts) || [];
      const mainContact = leadContacts.find(c => c.is_main) || leadContacts[0] || null;
      if (mainContact) contactIds.add(mainContact.id);
      partial.push({
        id: l.id, created_at: l.created_at, status_id: l.status_id,
        status_name: statuses[l.status_id] || 'Desconocido',
        devTags, price: l.price || 0, utmCampaign, utmContent, fuente,
        contactId: mainContact ? mainContact.id : null,
      });
    }
    if (!data._links || !data._links.next) break;
    page++;
  }

  // Fase 2: traer los custom_fields_values de todos los contactos referenciados, de una.
  const contactFields = fetchContactsById(contactIds);

  // Fase 3: resolver casa_visitada / fecha_visita / nombre / teléfono por contacto y armar
  // las filas finales.
  const rows = partial.map(p => {
    const contact = p.contactId ? contactFields[p.contactId] : null;
    const cf = contact ? contact.customFields : null;
    const casaVisitada = extractCustomFieldById(cf, FIELD_CASA_VISITADA);
    const fechaVisita = extractCustomFieldDateById(cf, FIELD_FECHA_VISITA);
    const contactName = contact ? contact.name : '';
    const phone = extractPhone(cf);
    return [
      p.id,
      new Date(p.created_at * 1000),
      p.status_id,
      p.status_name,
      QUALIFIED_IDS.includes(p.status_id) ? 1 : 0,
      p.status_id === WON_ID ? 1 : 0,
      p.status_id === LOST_ID ? 1 : 0,
      p.devTags.join(', '),
      p.price,
      p.utmCampaign,
      p.utmContent,
      casaVisitada,
      p.fuente,
      fechaVisita,
      contactName,
      phone,
    ];
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Leads') || ss.insertSheet('Leads');
  sheet.clear();
  sheet.appendRow(['id', 'created_at', 'status_id', 'status_name', 'calificado', 'ganado', 'perdido', 'desarrollos', 'price', 'utm_campaign', 'utm_content', 'casa_visitada', 'fuente', 'fecha_visita', 'contact_name', 'phone']);
  if (rows.length) sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(2, 2, Math.max(rows.length, 1), 1).setNumberFormat('yyyy-mm-dd hh:mm');
  sheet.getRange(2, 14, Math.max(rows.length, 1), 1).setNumberFormat('yyyy-mm-dd hh:mm');
}

function syncMetaDaily() {
  const token = props().getProperty('META_ACCESS_TOKEN');
  const adAccount = props().getProperty('META_AD_ACCOUNT_ID');

  const campResp = fetchWithRetry(
    `https://graph.facebook.com/v19.0/${adAccount}/campaigns?fields=name,objective,status&limit=100&access_token=${encodeURIComponent(token)}`);
  const campaigns = JSON.parse(campResp.getContentText()).data || [];
  const objectives = {}; const statuses = {};
  campaigns.forEach(c => { objectives[c.name] = c.objective; statuses[c.name] = c.status; });

  const since = new Date();
  since.setMonth(since.getMonth() - META_MONTHS_BACK);
  const until = new Date();
  const fmt = d => Utilities.formatDate(d, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd');
  const fields = 'campaign_name,adset_name,spend,impressions,reach,clicks,actions';
  let url = `https://graph.facebook.com/v19.0/${adAccount}/insights` +
    `?level=adset&time_increment=1&time_range=${encodeURIComponent(JSON.stringify({ since: fmt(since), until: fmt(until) }))}` +
    `&fields=${fields}&limit=1000&access_token=${encodeURIComponent(token)}`;

  const rows = [];
  while (url) {
    const resp = JSON.parse(fetchWithRetry(url).getContentText());
    for (const r of resp.data || []) {
      const actions = {};
      (r.actions || []).forEach(a => { actions[a.action_type] = parseFloat(a.value); });
      // "Conversaciones" = Meta "Conversaciones por mensajes iniciados". El nombre estándar de
      // Meta para esa métrica es messaging_conversation_started_7d -- se prioriza ese; si no
      // viene (cuentas viejas a veces solo tienen el genérico), cae a total_messaging_connection
      // como antes. Ver debugMetaActions() para confirmar contra lo que muestra Meta Ads Manager.
      const conversations = actions['onsite_conversion.messaging_conversation_started_7d']
        ?? actions['onsite_conversion.total_messaging_connection'] ?? 0;
      rows.push([
        r.date_start,
        r.campaign_name,
        r.adset_name,
        objectives[r.campaign_name] || '',
        statuses[r.campaign_name] || '',
        parseFloat(r.spend || 0),
        parseInt(r.impressions || 0, 10),
        parseInt(r.reach || 0, 10),
        parseInt(r.clicks || 0, 10),
        actions['link_click'] || 0,
        conversations,
        actions['post_engagement'] || 0,
      ]);
    }
    url = resp.paging && resp.paging.next;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('MetaDaily') || ss.insertSheet('MetaDaily');
  sheet.clear();
  sheet.appendRow(['date', 'campaign_name', 'adset_name', 'objective', 'status', 'spend', 'impressions', 'reach', 'clicks', 'link_clicks', 'conversations', 'post_engagement']);
  if (rows.length) sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}
