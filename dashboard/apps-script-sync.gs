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
const META_DAYS_BACK = 120; // solo se usa si META_SINCE_DATE está vacío
const META_SINCE_DATE = '2020-01-01'; // trae todo el histórico disponible en la cuenta de Meta

function props() { return PropertiesService.getScriptProperties(); }

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

function syncKommoLeads() {
  const subdomain = props().getProperty('KOMMO_SUBDOMAIN');
  const token = props().getProperty('KOMMO_ACCESS_TOKEN');
  const headers = { Authorization: 'Bearer ' + token };

  // pipeline statuses (for names + sort order)
  const pipeline = JSON.parse(UrlFetchApp.fetch(
    `https://${subdomain}/api/v4/leads/pipelines/${PIPELINE_ID}`, { headers }).getContentText());
  const statuses = {};
  pipeline._embedded.statuses.forEach(s => { statuses[s.id] = s.name; });

  const rows = [];
  let page = 1;
  while (true) {
    const resp = UrlFetchApp.fetch(
      `https://${subdomain}/api/v4/leads?limit=250&page=${page}`,
      { headers, muteHttpExceptions: true });
    if (resp.getResponseCode() === 204) break;
    const data = JSON.parse(resp.getContentText());
    const leads = (data._embedded && data._embedded.leads) || [];
    if (!leads.length) break;
    for (const l of leads) {
      const rawTags = (l._embedded && l._embedded.tags || []).map(t => t.name);
      const devTags = [...new Set(rawTags.filter(t => !EXCLUDE_TAGS.includes(t)).map(canonicalTag))];
      rows.push([
        l.id,
        new Date(l.created_at * 1000),
        l.status_id,
        statuses[l.status_id] || 'Desconocido',
        QUALIFIED_IDS.includes(l.status_id) ? 1 : 0,
        l.status_id === WON_ID ? 1 : 0,
        l.status_id === LOST_ID ? 1 : 0,
        devTags.join(', '),
        l.price || 0,
      ]);
    }
    if (!data._links || !data._links.next) break;
    page++;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Leads') || ss.insertSheet('Leads');
  sheet.clear();
  sheet.appendRow(['id', 'created_at', 'status_id', 'status_name', 'calificado', 'ganado', 'perdido', 'desarrollos', 'price']);
  if (rows.length) sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(2, 2, Math.max(rows.length, 1), 1).setNumberFormat('yyyy-mm-dd hh:mm');
}

function syncMetaDaily() {
  const token = props().getProperty('META_ACCESS_TOKEN');
  const adAccount = props().getProperty('META_AD_ACCOUNT_ID');

  const campResp = UrlFetchApp.fetch(
    `https://graph.facebook.com/v19.0/${adAccount}/campaigns?fields=name,objective,status&limit=100&access_token=${encodeURIComponent(token)}`);
  const campaigns = JSON.parse(campResp.getContentText()).data || [];
  const objectives = {}; const statuses = {};
  campaigns.forEach(c => { objectives[c.name] = c.objective; statuses[c.name] = c.status; });

  const since = META_SINCE_DATE ? new Date(META_SINCE_DATE + 'T00:00:00Z') : new Date(Date.now() - META_DAYS_BACK * 86400000);
  const until = new Date();
  const fmt = d => Utilities.formatDate(d, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd');
  const fields = 'campaign_name,adset_name,spend,impressions,reach,clicks,actions';
  let url = `https://graph.facebook.com/v19.0/${adAccount}/insights` +
    `?level=adset&time_increment=1&time_range=${encodeURIComponent(JSON.stringify({ since: fmt(since), until: fmt(until) }))}` +
    `&fields=${fields}&limit=1000&access_token=${encodeURIComponent(token)}`;

  const rows = [];
  while (url) {
    const resp = JSON.parse(UrlFetchApp.fetch(url).getContentText());
    for (const r of resp.data || []) {
      const actions = {};
      (r.actions || []).forEach(a => { actions[a.action_type] = parseFloat(a.value); });
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
        actions['onsite_conversion.total_messaging_connection'] || 0,
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
