import json, os, datetime, calendar
import openpyxl

BASE = os.path.dirname(os.path.abspath(__file__))
XLSX_PATH = os.path.join(BASE, "claude_dashboard.xlsx")

WON_ID, LOST_ID = 142, 143
QUALIFIED_IDS = [105609867, 109532768, 105671691, 105671695, WON_ID]  # visita-reunion, reunion realizada, 2da reunion, negociacion, + ganados

ADSET_TO_DEV = {"FAMILIA_CH": "Chubut", "FAMILIA_SI": "Simón Iriondo", "FAMILIA_MIS": "Misiones"}
TAG_ALIASES = {"Difusion Misiones": "Misiones"}
EXCLUDE_TAGS = ["Difusión", "Follow-up 1", "JN", "WA", "Interes Futuro", "Apta Credito",
                "Presu menos 300", "Barrio Cerrado", "Inmobiliaria"]

RESULT_LABEL = {
    "OUTCOME_LEADS": "Conversaciones de WhatsApp",
    "OUTCOME_AWARENESS": "Alcance",
    "OUTCOME_TRAFFIC": "Clics al enlace",
    "OUTCOME_ENGAGEMENT": "Interacciones",
}
RESULT_FIELD = {
    "OUTCOME_LEADS": "conversations",
    "OUTCOME_AWARENESS": "reach",
    "OUTCOME_TRAFFIC": "link_clicks",
    "OUTCOME_ENGAGEMENT": "post_engagement",
}

# Curated funnel order (Kommo doesn't expose "sort" through the Sheet sync yet,
# so this is inferred from the pipeline stage names / the qualified-stage sequence
# already encoded in QUALIFIED_IDS). Anything not listed here is appended at the end
# in first-seen order, so a future/renamed status never disappears from the funnel.
STATUS_ORDER_HINT = [
    105609855,  # Contacto inicial
    105609863,  # potenciales
    105609859,  # Follow UP
    105609867,  # visita-reunion
    109532768,  # reunion realizada
    105671691,  # 2da reunion
    105671695,  # negociacion
]

wb = openpyxl.load_workbook(XLSX_PATH, data_only=True)

# ---- Leads ----
ws = wb["Leads"]
rows = list(ws.iter_rows(values_only=True))
header, lead_rows = rows[0], [r for r in rows[1:] if r[0] is not None]

leads_out = []
statuses = {}
first_seen_status_order = []
for r in lead_rows:
    lid, created_at, status_id, status_name, calificado, ganado, perdido, desarrollos, price = r
    status_id = int(status_id)
    if status_id not in statuses:
        statuses[status_id] = status_name
        first_seen_status_order.append(status_id)
    if isinstance(created_at, datetime.datetime):
        created_epoch = calendar.timegm(created_at.timetuple())
    else:
        created_epoch = None
    tags = [t.strip() for t in desarrollos.split(",")] if desarrollos else []
    leads_out.append({
        "id": int(lid),
        "status_id": status_id,
        "price": price or 0,
        "created_at": created_epoch,
        "tags": tags,
    })

status_order = {}
for i, sid in enumerate(STATUS_ORDER_HINT):
    status_order[sid] = i
next_i = len(STATUS_ORDER_HINT)
for sid in first_seen_status_order:
    if sid not in status_order:
        status_order[sid] = next_i
        next_i += 1

# ---- MetaDaily ----
ws2 = wb["MetaDaily"]
rows2 = list(ws2.iter_rows(values_only=True))
h2, meta_rows = rows2[0], [r for r in rows2[1:] if r[0] is not None]

meta_daily = []
campaign_objectives = {}
campaign_status = {}
for r in meta_rows:
    (date, campaign_name, adset_name, objective, status, spend, impressions,
     reach, clicks, link_clicks, conversations, post_engagement) = r
    date_str = date.strftime("%Y-%m-%d") if isinstance(date, datetime.datetime) else str(date)
    campaign_objectives.setdefault(campaign_name, objective)
    campaign_status[campaign_name] = status  # last-seen status (most recent day) wins
    meta_daily.append({
        "date": date_str,
        "campaign_name": campaign_name,
        "adset_name": adset_name,
        "spend": float(spend or 0),
        "impressions": int(impressions or 0),
        "reach": int(reach or 0),
        "clicks": int(clicks or 0),
        "link_clicks": float(link_clicks or 0),
        "conversations": float(conversations or 0),
        "leads_meta": 0,
        "post_engagement": float(post_engagement or 0),
    })

# ---- "Meta" tab: last sync timestamp ----
ws3 = wb["Meta"]
last_sync_raw = ws3["A2"].value
if isinstance(last_sync_raw, datetime.datetime):
    generated_at = last_sync_raw.isoformat() + "Z"
else:
    generated_at = datetime.datetime.now(datetime.timezone.utc).isoformat().replace("+00:00", "Z")

d = {"generated_at": generated_at}

payload = {
    "generated_at": d["generated_at"],
    "statuses": {str(k): v for k, v in statuses.items()},
    "status_order": {str(k): v for k, v in status_order.items()},
    "leads": leads_out,
    "meta_daily": meta_daily,
    "campaign_objectives": campaign_objectives,
    "campaign_status": campaign_status,
    "won_id": WON_ID,
    "lost_id": LOST_ID,
    "qualified_ids": QUALIFIED_IDS,
    "adset_to_dev": ADSET_TO_DEV,
    "tag_aliases": TAG_ALIASES,
    "exclude_tags": EXCLUDE_TAGS,
    "result_label": RESULT_LABEL,
    "result_field": RESULT_FIELD,
}

print(f"OK: {len(leads_out)} leads, {len(meta_daily)} filas de meta, generated_at={d['generated_at']}")


data_json = json.dumps(payload, ensure_ascii=False)

gen_utc = datetime.datetime.fromisoformat(d["generated_at"].replace("Z", "+00:00"))
gen_local = gen_utc - datetime.timedelta(hours=3)
MESES = ["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"]
fecha_str = f"{gen_local.day} de {MESES[gen_local.month-1]} de {gen_local.year}, {gen_local.hour:02d}:{gen_local.minute:02d} hs (ARG)"

CSS = r"""
:root {
  --bg: #F0EEE6;
  --surface: #FFFFFF;
  --surface-2: #E9E5DA;
  --ink: #1B1812;
  --ink-muted: #6E6858;
  --border: #DED9C9;
  --accent-teal: #24A965;
  --accent-blue: #154AA0;
  --good: #B98A2E;
  --critical: #B23B34;
  --funnel-0: #6CD79D; --funnel-1: #5AC68D; --funnel-2: #48B57D; --funnel-3: #34A46E;
  --funnel-4: #1C945F; --funnel-5: #138252; --funnel-6: #0B7146; --funnel-7: #03603B; --funnel-8: #00502F;
  --shadow: 0 1px 2px rgba(27,24,18,0.06), 0 8px 24px -12px rgba(27,24,18,0.18);
}
@media (prefers-color-scheme: dark) {
  :root:not([data-theme="light"]) {
    --bg: #15130F; --surface: #1E1B16; --surface-2: #262219; --ink: #F3EFE6; --ink-muted: #A39C8C;
    --border: #332E24; --accent-teal: #1EA662; --accent-blue: #3065BE; --good: #D9A94F; --critical: #E2635C;
    --funnel-0: #4AB67F; --funnel-1: #3CAB74; --funnel-2: #2D9F69; --funnel-3: #1A935E;
    --funnel-4: #1A8656; --funnel-5: #007B4B; --funnel-6: #036E43; --funnel-7: #06623C; --funnel-8: #085634;
    --shadow: 0 1px 2px rgba(0,0,0,0.3), 0 8px 24px -12px rgba(0,0,0,0.5);
  }
}
:root[data-theme="dark"] {
  --bg: #15130F; --surface: #1E1B16; --surface-2: #262219; --ink: #F3EFE6; --ink-muted: #A39C8C;
  --border: #332E24; --accent-teal: #1EA662; --accent-blue: #3065BE; --good: #D9A94F; --critical: #E2635C;
  --funnel-0: #4AB67F; --funnel-1: #3CAB74; --funnel-2: #2D9F69; --funnel-3: #1A935E;
  --funnel-4: #1A8656; --funnel-5: #007B4B; --funnel-6: #036E43; --funnel-7: #06623C; --funnel-8: #085634;
  --shadow: 0 1px 2px rgba(0,0,0,0.3), 0 8px 24px -12px rgba(0,0,0,0.5);
}
* { box-sizing: border-box; }
html, body { margin: 0; padding: 0; }
body {
  background: var(--bg); color: var(--ink);
  font-family: "Work Sans", ui-sans-serif, system-ui, sans-serif;
  font-size: 15px; line-height: 1.45; -webkit-font-smoothing: antialiased;
}
.wrap { max-width: 1220px; margin: 0 auto; padding: 28px 24px 64px; }
header.top { display: flex; align-items: center; justify-content: space-between; gap: 20px; margin-bottom: 20px; flex-wrap: wrap; }
.brand { display: flex; align-items: center; gap: 14px; }
.brand-mark {
  width: 52px; height: 52px; border-radius: 50%; background: #15130F; color: #F3EFE6;
  border: 1.5px solid var(--border); display: flex; align-items: center; justify-content: center;
  font-family: "Bricolage Grotesque", sans-serif; font-weight: 700; font-size: 12px;
  letter-spacing: 0.02em; text-align: center; line-height: 1.05; flex-shrink: 0; box-shadow: var(--shadow);
}
.brand-text h1 { font-family: "Bricolage Grotesque", sans-serif; font-weight: 700; font-size: 22px; margin: 0; letter-spacing: -0.01em; }
.updated-badge {
  display: flex; align-items: center; gap: 8px; background: var(--surface); border: 1px solid var(--border);
  border-radius: 100px; padding: 8px 14px; font-size: 12.5px; color: var(--ink-muted);
  font-family: "IBM Plex Mono", monospace; box-shadow: var(--shadow); white-space: nowrap;
}
.updated-badge .pulse {
  width: 7px; height: 7px; border-radius: 50%; background: var(--accent-teal);
  box-shadow: 0 0 0 0 var(--accent-teal); animation: pulse 2.4s infinite;
}
@media (prefers-reduced-motion: reduce) { .updated-badge .pulse { animation: none; } }
@keyframes pulse {
  0% { box-shadow: 0 0 0 0 color-mix(in srgb, var(--accent-teal) 55%, transparent); }
  70% { box-shadow: 0 0 0 7px transparent; }
  100% { box-shadow: 0 0 0 0 transparent; }
}

.filter-bar { display: flex; align-items: center; gap: 8px; margin-bottom: 8px; flex-wrap: wrap; }
.filter-btn {
  font-family: "Work Sans", sans-serif; font-size: 13px; font-weight: 600; color: var(--ink-muted);
  background: var(--surface); border: 1px solid var(--border); border-radius: 100px;
  padding: 7px 14px; cursor: pointer; transition: background-color .15s, color .15s, border-color .15s;
}
.filter-btn:hover { border-color: var(--accent-teal); }
.filter-btn:focus-visible { outline: 2px solid var(--accent-blue); outline-offset: 1px; }
.filter-btn.active { background: var(--ink); color: var(--bg); border-color: var(--ink); }
.range-caption { font-size: 12px; color: var(--ink-muted); margin: 0 0 24px; font-family: "IBM Plex Mono", monospace; }

.kpi-grid { display: grid; grid-template-columns: repeat(5, 1fr); gap: 14px; margin-bottom: 28px; }
@media (max-width: 1080px) { .kpi-grid { grid-template-columns: repeat(3, 1fr); } }
@media (max-width: 640px) { .kpi-grid { grid-template-columns: repeat(2, 1fr); } }
.kpi-card { background: var(--surface); border: 1px solid var(--border); border-radius: 16px; padding: 18px 18px 16px; box-shadow: var(--shadow); }
.kpi-label { font-size: 11.5px; text-transform: uppercase; letter-spacing: 0.06em; color: var(--ink-muted); font-weight: 600; margin: 0 0 10px; }
.kpi-value-row { display: flex; align-items: baseline; gap: 8px; flex-wrap: wrap; }
.kpi-value { font-family: "Bricolage Grotesque", sans-serif; font-weight: 700; font-size: 28px; letter-spacing: -0.01em; font-variant-numeric: tabular-nums; line-height: 1.05; }
.kpi-delta { font-size: 12px; font-weight: 600; padding: 2px 7px; border-radius: 100px; font-variant-numeric: tabular-nums; white-space: nowrap; }
.kpi-delta.good { color: var(--good); background: color-mix(in srgb, var(--good) 16%, transparent); }
.kpi-delta.bad { color: var(--critical); background: color-mix(in srgb, var(--critical) 14%, transparent); }
.kpi-delta.neutral { color: var(--ink-muted); background: var(--surface-2); }
.kpi-sub { margin: 8px 0 0; font-size: 12.5px; color: var(--ink-muted); }
.kpi-sub b { color: var(--ink); font-weight: 600; font-variant-numeric: tabular-nums; }

.panels { display: grid; grid-template-columns: 1.3fr 1fr; gap: 16px; margin-bottom: 16px; align-items: start; }
@media (max-width: 900px) { .panels { grid-template-columns: 1fr; } }
.panel { background: var(--surface); border: 1px solid var(--border); border-radius: 16px; padding: 20px 20px 18px; box-shadow: var(--shadow); }
.panel h2 { font-family: "Bricolage Grotesque", sans-serif; font-size: 16px; font-weight: 700; margin: 0 0 4px; text-wrap: balance; }
.panel .panel-sub { margin: 0 0 16px; color: var(--ink-muted); font-size: 12.5px; }

.bar-row { display: grid; grid-template-columns: 132px 1fr 52px; align-items: center; gap: 10px; padding: 6px 0; }
.bar-label { font-size: 12.5px; color: var(--ink-muted); display: flex; align-items: center; gap: 6px; }
.bar-label .dot { width: 8px; height: 8px; border-radius: 50%; background: var(--accent-teal); flex-shrink: 0; }
.bar-track { background: var(--surface-2); border-radius: 6px; height: 16px; overflow: hidden; }
.bar-fill { height: 100%; width: var(--pct); border-radius: 6px; min-width: 3px; }
.bar-value { text-align: right; font-family: "IBM Plex Mono", monospace; font-size: 12.5px; font-variant-numeric: tabular-nums; }
.empty-note { font-size: 12.5px; color: var(--ink-muted); padding: 8px 0; }

.won-lost { display: flex; gap: 10px; margin-top: 14px; padding-top: 14px; border-top: 1px dashed var(--border); flex-wrap: wrap; }
.pill { display: inline-flex; align-items: center; gap: 6px; font-size: 12px; font-weight: 600; padding: 5px 10px; border-radius: 100px; }
.pill.won { background: color-mix(in srgb, var(--good) 18%, var(--surface)); color: var(--good); }
.pill.lost { background: color-mix(in srgb, var(--critical) 16%, var(--surface)); color: var(--critical); }
.pill .dotp { width: 6px; height: 6px; border-radius: 50%; background: currentColor; }

.camp-row { padding: 12px 0; border-bottom: 1px solid var(--border); }
.camp-row:last-child { border-bottom: none; padding-bottom: 0; }
.camp-head { display: flex; justify-content: space-between; align-items: baseline; margin-bottom: 6px; gap: 8px; }
.camp-name { font-weight: 600; font-size: 13.5px; }
.camp-objective { font-size: 11px; color: var(--ink-muted); text-transform: uppercase; letter-spacing: 0.04em; }
.camp-fill-primary { background: var(--accent-blue); }
.camp-fill-muted { background: var(--ink-muted); opacity: 0.45; }
.camp-foot { display: flex; justify-content: space-between; margin-top: 6px; font-family: "IBM Plex Mono", monospace; font-size: 12.5px; }
.camp-foot .dim { color: var(--ink-muted); font-family: "Work Sans", sans-serif; }
.is-primary .camp-name { color: var(--accent-blue); }

.wide-panel { margin-bottom: 16px; }
.table-scroll { overflow-x: auto; }
table.data-table { width: 100%; border-collapse: collapse; font-size: 13px; min-width: 640px; }
table.data-table th { text-align: left; font-size: 11px; text-transform: uppercase; letter-spacing: 0.04em; color: var(--ink-muted); font-weight: 600; padding: 0 10px 8px; border-bottom: 1px solid var(--border); white-space: nowrap; }
table.data-table td { padding: 10px 10px; border-bottom: 1px solid var(--border); font-variant-numeric: tabular-nums; white-space: nowrap; }
table.data-table td.label-cell { font-weight: 600; white-space: normal; font-variant-numeric: initial; }
table.data-table tr:last-child td { border-bottom: none; }
table.data-table tr.total-row td { font-weight: 700; border-top: 1px solid var(--border); }
.status-pill { font-size: 10.5px; font-weight: 700; text-transform: uppercase; letter-spacing: .03em; padding: 2px 7px; border-radius: 100px; }
.status-pill.active { color: var(--good); background: color-mix(in srgb, var(--good) 16%, transparent); }
.status-pill.paused { color: var(--ink-muted); background: var(--surface-2); }
.result-caption { display: block; font-size: 10.5px; color: var(--ink-muted); font-family: "Work Sans", sans-serif; font-weight: 400; }
.no-data { color: var(--ink-muted); }

footer { margin-top: 24px; padding-top: 18px; border-top: 1px solid var(--border); display: flex; justify-content: space-between; flex-wrap: wrap; gap: 10px; font-size: 12px; color: var(--ink-muted); }
"""

HTML_TEMPLATE = """<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>Dashboard CosCor</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Bricolage+Grotesque:opsz,wght@12..96,400..800&family=Work+Sans:ital,wght@0,400;0,500;0,600;0,700;1,400&family=IBM+Plex+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
__CSS__
</style>
</head>
<body>
<div class="wrap">

  <header class="top">
    <div class="brand">
      <div class="brand-mark">CosCor</div>
      <div class="brand-text"><h1>Dashboard CosCor</h1></div>
    </div>
    <div class="updated-badge"><span class="pulse"></span> Actualizado el __FECHA__</div>
  </header>

  <div class="filter-bar" id="filterBar" role="group" aria-label="Rango de fechas"></div>
  <p class="range-caption" id="rangeCaption"></p>

  <div class="kpi-grid" id="kpiGrid"></div>

  <div class="panels">
    <div class="panel">
      <h2>Embudo de ventas</h2>
      <p class="panel-sub">Leads generados en el período, por etapa actual en Kommo</p>
      <div id="funnelRows"></div>
      <div class="won-lost" id="wonLost"></div>
    </div>
    <div class="panel">
      <h2>Inversión por campaña</h2>
      <p class="panel-sub">Gasto en Meta Ads en el período · ordenado de mayor a menor</p>
      <div id="campaignRows"></div>
    </div>
  </div>

  <div class="panel wide-panel">
    <h2>Leads con visita por desarrollo × inversión en Meta</h2>
    <p class="panel-sub">Cruce entre los conjuntos de anuncios de Meta y los desarrollos etiquetados en Kommo · "con visita" incluye visita, reunión realizada, 2da reunión, negociación y ganados</p>
    <div class="table-scroll"><table class="data-table" id="devTable"></table></div>
  </div>

  <div class="panel wide-panel">
    <h2>Conversaciones de WhatsApp por grupo de anuncios</h2>
    <p class="panel-sub">Conversaciones de WhatsApp iniciadas en Meta Ads (campañas con objetivo de conversión), por conjunto de anuncios</p>
    <div class="table-scroll"><table class="data-table" id="whatsappTable"></table></div>
  </div>

  <div class="panel wide-panel">
    <h2>Meta Ads — panel completo</h2>
    <p class="panel-sub">Todas las campañas activas y pausadas · el "resultado" depende del objetivo de cada campaña</p>
    <div class="table-scroll"><table class="data-table" id="metaTable"></table></div>
  </div>

  <div class="panel wide-panel">
    <h2>Leads por desarrollo</h2>
    <p class="panel-sub">Volumen de leads generados en el período según la propiedad o zona etiquetada en Kommo</p>
    <div id="projectRows"></div>
  </div>

  <footer>
    <span>Fuentes: Meta Ads (cuenta CosCor) · Kommo CRM (infocoscorlife)</span>
    <span>Se actualiza automáticamente todos los días</span>
  </footer>

</div>
<script>
const DATA = __DATA_JSON__;
__JS__
</script>
</body>
</html>
"""

JS = r"""
const ARG_OFFSET_MS = 3 * 3600 * 1000;
const fmtInt = n => Math.round(n).toLocaleString('es-AR');
const fmtARS = n => '$' + Math.round(n).toLocaleString('es-AR');
const fmtPct = (n, d=1) => (n === null || n === undefined || isNaN(n)) ? '—' : n.toFixed(d) + '%';

function argDateFromMs(ms) {
  const shifted = new Date(ms + ARG_OFFSET_MS);
  return { y: shifted.getUTCFullYear(), m: shifted.getUTCMonth(), d: shifted.getUTCDate() };
}
function dateKey(y, m, d) { return new Date(Date.UTC(y, m, d)).getTime(); }
function addDaysKey(key, n) { return key + n * 86400000; }
function keyToStr(key) {
  const dt = new Date(key);
  return `${dt.getUTCFullYear()}-${String(dt.getUTCMonth()+1).padStart(2,'0')}-${String(dt.getUTCDate()).padStart(2,'0')}`;
}
function keyToLabel(key) {
  const dt = new Date(key);
  const meses = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
  return `${dt.getUTCDate()} ${meses[dt.getUTCMonth()]}`;
}
function keyToArgUtcStart(key) { return key - ARG_OFFSET_MS; } // ms epoch (UTC) of ARG midnight for that calendar date
function daysInMonth(y, m) { return new Date(Date.UTC(y, m+1, 0)).getUTCDate(); }

const now = new Date();
const today = argDateFromMs(now.getTime());
const todayKey = dateKey(today.y, today.m, today.d);

function minDataKey() {
  // Earliest date across all leads + all Meta rows, so "Todo el período" always
  // covers the full history in the Sheet, however far back it goes.
  let minKey = todayKey;
  for (const l of DATA.leads) {
    if (l.created_at == null) continue;
    const ad = argDateFromMs(l.created_at * 1000);
    const k = dateKey(ad.y, ad.m, ad.d);
    if (k < minKey) minKey = k;
  }
  for (const r of DATA.meta_daily) {
    const [y, m, d] = r.date.split('-').map(Number);
    const k = dateKey(y, m - 1, d);
    if (k < minKey) minKey = k;
  }
  return minKey;
}

function buildRanges() {
  const yStart = addDaysKey(todayKey, -1), yEnd = yStart;
  const l7Start = addDaysKey(todayKey, -7), l7End = addDaysKey(todayKey, -1);
  const l30Start = addDaysKey(todayKey, -30), l30End = addDaysKey(todayKey, -1);
  const thisMonthStart = dateKey(today.y, today.m, 1), thisMonthEnd = todayKey;
  const lastMonthY = today.m === 0 ? today.y - 1 : today.y;
  const lastMonthM = today.m === 0 ? 11 : today.m - 1;
  const lastMonthStart = dateKey(lastMonthY, lastMonthM, 1);
  const lastMonthEnd = dateKey(lastMonthY, lastMonthM, daysInMonth(lastMonthY, lastMonthM));

  const prevLastMonthY = lastMonthM === 0 ? lastMonthY - 1 : lastMonthY;
  const prevLastMonthM = lastMonthM === 0 ? 11 : lastMonthM - 1;
  const mtdDayCount = today.d;
  const prevMtdEndDay = Math.min(mtdDayCount, daysInMonth(prevLastMonthY, prevLastMonthM));

  return {
    allTime: { label: 'Todo el período', start: minDataKey(), end: todayKey,
      prevStart: null, prevEnd: null },
    yesterday: { label: 'Ayer', start: yStart, end: yEnd,
      prevStart: addDaysKey(yStart, -1), prevEnd: addDaysKey(yEnd, -1) },
    last7: { label: 'Últimos 7 días', start: l7Start, end: l7End,
      prevStart: addDaysKey(l7Start, -7), prevEnd: addDaysKey(l7End, -7) },
    last30: { label: 'Últimos 30 días', start: l30Start, end: l30End,
      prevStart: addDaysKey(l30Start, -30), prevEnd: addDaysKey(l30End, -30) },
    lastMonth: { label: 'El mes pasado', start: lastMonthStart, end: lastMonthEnd,
      prevStart: dateKey(prevLastMonthY, prevLastMonthM, 1), prevEnd: dateKey(prevLastMonthY, prevLastMonthM, daysInMonth(prevLastMonthY, prevLastMonthM)) },
    mtd: { label: 'Este mes hasta la fecha', start: thisMonthStart, end: thisMonthEnd,
      prevStart: dateKey(lastMonthY, lastMonthM, 1), prevEnd: dateKey(lastMonthY, lastMonthM, prevMtdEndDay) },
  };
}
const RANGES = buildRanges();

function filterMetaDaily(startKey, endKey) {
  const s = keyToStr(startKey), e = keyToStr(endKey);
  return DATA.meta_daily.filter(r => r.date >= s && r.date <= e);
}
function filterLeads(startKey, endKey) {
  const s = keyToArgUtcStart(startKey) / 1000;
  const e = keyToArgUtcStart(addDaysKey(endKey, 1)) / 1000;
  return DATA.leads.filter(l => l.created_at >= s && l.created_at < e);
}

function canonicalTag(t) { return DATA.tag_aliases[t] || t; }

function aggregateWhatsappByAdset(rows) {
  // Conversaciones de WhatsApp iniciadas en Meta, por conjunto de anuncios —
  // solo campañas con objetivo de conversión (OUTCOME_LEADS = WhatsApp).
  const byAdset = {};
  for (const r of rows) {
    if (DATA.campaign_objectives[r.campaign_name] !== 'OUTCOME_LEADS') continue;
    const a = byAdset[r.adset_name] || { spend: 0, conversations: 0 };
    a.spend += r.spend;
    a.conversations += r.conversations;
    byAdset[r.adset_name] = a;
  }
  return byAdset;
}

function aggregateMeta(rows) {
  const agg = { spend: 0, impressions: 0, reach: 0, clicks: 0, link_clicks: 0, conversations: 0, leads_meta: 0, post_engagement: 0 };
  const byCampaign = {};
  const byAdset = {};
  for (const r of rows) {
    for (const k of Object.keys(agg)) agg[k] += r[k] || 0;
    const c = byCampaign[r.campaign_name] || { spend: 0, impressions: 0, reach: 0, clicks: 0, link_clicks: 0, conversations: 0, post_engagement: 0 };
    c.spend += r.spend; c.impressions += r.impressions; c.reach += r.reach; c.clicks += r.clicks;
    c.link_clicks += r.link_clicks; c.conversations += r.conversations; c.post_engagement += r.post_engagement;
    byCampaign[r.campaign_name] = c;
    const a = byAdset[r.adset_name] || { spend: 0, impressions: 0, clicks: 0, conversations: 0 };
    a.spend += r.spend; a.impressions += r.impressions; a.clicks += r.clicks; a.conversations += r.conversations;
    byAdset[r.adset_name] = a;
  }
  return { ...agg, byCampaign, byAdset };
}

function aggregateLeads(leads) {
  const byStatus = {}; const byTag = {};
  let won = 0, lost = 0, qualified = 0;
  for (const l of leads) {
    byStatus[l.status_id] = (byStatus[l.status_id] || 0) + 1;
    if (l.status_id === DATA.won_id) won++;
    if (l.status_id === DATA.lost_id) lost++;
    if (DATA.qualified_ids.includes(l.status_id)) qualified++;
    const seen = new Set();
    for (const t of l.tags) {
      if (DATA.exclude_tags.includes(t)) continue;
      const c = canonicalTag(t);
      if (seen.has(c)) continue;
      seen.add(c);
      byTag[c] = (byTag[c] || 0) + 1;
    }
  }
  return { total: leads.length, byStatus, byTag, won, lost, qualified };
}

function devLeadsByTagAndQualified(leads) {
  const byTag = {}; // tag -> {leads, qualified, spend}
  for (const l of leads) {
    const qualified = DATA.qualified_ids.includes(l.status_id);
    const seen = new Set();
    for (const t of l.tags) {
      if (DATA.exclude_tags.includes(t)) continue;
      const c = canonicalTag(t);
      if (seen.has(c)) continue;
      seen.add(c);
      if (!byTag[c]) byTag[c] = { leads: 0, qualified: 0 };
      byTag[c].leads++;
      if (qualified) byTag[c].qualified++;
    }
  }
  return byTag;
}

function delta(curr, prev) {
  if (prev === 0 || prev === null || prev === undefined) {
    if (curr === 0) return { pct: null, dir: 0 };
    return { pct: null, dir: 1, isNew: true };
  }
  const pct = (curr - prev) / prev * 100;
  return { pct, dir: pct > 0 ? 1 : (pct < 0 ? -1 : 0) };
}

function deltaBadge(curr, prev, higherIsGood) {
  const d = delta(curr, prev);
  if (d.pct === null && !d.isNew) return '<span class="kpi-delta neutral">sin cambios</span>';
  if (d.isNew) return '<span class="kpi-delta neutral">nuevo</span>';
  const arrow = d.dir >= 0 ? '▲' : '▼';
  let cls = 'neutral';
  if (higherIsGood !== null) {
    const isGood = higherIsGood ? d.dir >= 0 : d.dir <= 0;
    cls = isGood ? 'good' : 'bad';
  }
  return `<span class="kpi-delta ${cls}">${arrow} ${Math.abs(d.pct).toFixed(0)}%</span>`;
}

const FUNNEL_ORDER = Object.keys(DATA.status_order)
  .filter(id => Number(id) !== DATA.won_id && Number(id) !== DATA.lost_id)
  .sort((a, b) => DATA.status_order[a] - DATA.status_order[b]);

function render(rangeKey) {
  const r = RANGES[rangeKey];
  const noPrev = r.prevStart === null;
  document.querySelectorAll('.filter-btn').forEach(b => b.classList.toggle('active', b.dataset.range === rangeKey));

  const curLeads = filterLeads(r.start, r.end);
  const prevLeads = noPrev ? [] : filterLeads(r.prevStart, r.prevEnd);
  const curMeta = aggregateMeta(filterMetaDaily(r.start, r.end));
  const prevMeta = aggregateMeta(noPrev ? [] : filterMetaDaily(r.prevStart, r.prevEnd));
  const curAgg = aggregateLeads(curLeads);
  const prevAgg = aggregateLeads(prevLeads);
  const badge = (curVal, prevVal, higherIsGood) => noPrev ? '' : deltaBadge(curVal, prevVal ?? 0, higherIsGood);

  document.getElementById('rangeCaption').textContent = noPrev
    ? `${keyToLabel(r.start)} – ${keyToLabel(r.end)} · todo el historial disponible`
    : `${keyToLabel(r.start)} – ${keyToLabel(r.end)} · vs. ${keyToLabel(r.prevStart)} – ${keyToLabel(r.prevEnd)}`;

  // ---- KPIs ----
  const convRate = (curAgg.won + curAgg.lost) ? curAgg.won / (curAgg.won + curAgg.lost) * 100 : null;
  const prevConvRate = (prevAgg.won + prevAgg.lost) ? prevAgg.won / (prevAgg.won + prevAgg.lost) * 100 : null;
  const qualRate = curAgg.total ? curAgg.qualified / curAgg.total * 100 : null;
  const whatsappCur = curMeta.byCampaign['WHATSAPP'] || { spend: 0, conversations: 0 };
  const whatsappPrev = prevMeta.byCampaign['WHATSAPP'] || { spend: 0, conversations: 0 };
  const cpcCur = whatsappCur.conversations ? whatsappCur.spend / whatsappCur.conversations : null;

  const kpis = [
    { label: 'Leads generados', value: fmtInt(curAgg.total), badge: badge(curAgg.total, prevAgg.total, true),
      sub: `${fmtInt(curAgg.won)} ganados · ${fmtInt(curAgg.lost)} perdidos` },
    { label: 'Leads con visita', value: fmtInt(curAgg.qualified), badge: badge(curAgg.qualified, prevAgg.qualified, true),
      sub: `${fmtPct(qualRate)} del total · visita, 2da reunión, negociación o ganado` },
    { label: 'Tasa de conversión', value: fmtPct(convRate), badge: convRate === null ? '' : badge(convRate, prevConvRate, true),
      sub: 'ganados / (ganados + perdidos)' },
    { label: 'Inversión en Meta Ads', value: fmtARS(curMeta.spend), badge: badge(curMeta.spend, prevMeta.spend, null),
      sub: `${fmtInt(curMeta.impressions)} impresiones` },
    { label: 'Conversaciones de WhatsApp', value: fmtInt(whatsappCur.conversations), badge: badge(whatsappCur.conversations, whatsappPrev.conversations, true),
      sub: cpcCur === null ? 'Meta Ads · WhatsApp' : `${fmtARS(cpcCur)} costo por conversación` },
  ];
  document.getElementById('kpiGrid').innerHTML = kpis.map(k => `
    <div class="kpi-card">
      <p class="kpi-label">${k.label}</p>
      <div class="kpi-value-row"><span class="kpi-value">${k.value}</span>${k.badge}</div>
      <p class="kpi-sub">${k.sub}</p>
    </div>`).join('');

  // ---- Funnel (skip zero) ----
  const maxFunnel = Math.max(1, ...FUNNEL_ORDER.map(id => curAgg.byStatus[id] || 0));
  const funnelHtml = FUNNEL_ORDER.map((id, i) => {
    const count = curAgg.byStatus[id] || 0;
    if (!count) return '';
    const pct = count / maxFunnel * 100;
    return `<div class="bar-row">
      <div class="bar-label">${DATA.statuses[id]}</div>
      <div class="bar-track"><div class="bar-fill" style="width:${pct}%; background:var(--funnel-${i % 9})"></div></div>
      <div class="bar-value">${fmtInt(count)}</div>
    </div>`;
  }).join('');
  document.getElementById('funnelRows').innerHTML = funnelHtml || '<p class="empty-note">Sin leads en este período.</p>';
  document.getElementById('wonLost').innerHTML = `
    <span class="pill won"><span class="dotp"></span>${fmtInt(curAgg.won)} logrados con éxito</span>
    <span class="pill lost"><span class="dotp"></span>${fmtInt(curAgg.lost)} venta perdida</span>`;

  // ---- Campaigns (sorted by spend desc) ----
  const OBJ_LABEL = { OUTCOME_LEADS: 'Conversión · WhatsApp', OUTCOME_AWARENESS: 'Alcance / awareness', OUTCOME_TRAFFIC: 'Tráfico', OUTCOME_ENGAGEMENT: 'Interacción' };
  const campEntries = Object.entries(curMeta.byCampaign).sort((a, b) => b[1].spend - a[1].spend);
  const maxSpend = Math.max(1, ...campEntries.map(([, c]) => c.spend));
  document.getElementById('campaignRows').innerHTML = campEntries.length ? campEntries.map(([name, c]) => {
    const isPrimary = name === 'WHATSAPP';
    const pct = c.spend / maxSpend * 100;
    const extra = c.conversations ? `${fmtInt(c.conversations)} conversaci${c.conversations === 1 ? 'ón' : 'ones'}` : `${fmtInt(c.clicks)} clics`;
    return `<div class="camp-row ${isPrimary ? 'is-primary' : ''}">
      <div class="camp-head"><span class="camp-name">${name}</span><span class="camp-objective">${OBJ_LABEL[DATA.campaign_objectives[name]] || ''}</span></div>
      <div class="bar-track"><div class="bar-fill ${isPrimary ? 'camp-fill-primary' : 'camp-fill-muted'}" style="width:${pct}%"></div></div>
      <div class="camp-foot"><span>${fmtARS(c.spend)}</span><span class="dim">${extra}</span></div>
    </div>`;
  }).join('') : '<p class="empty-note">Sin inversión en este período.</p>';

  // ---- Dev x spend cross table ----
  const devLeads = devLeadsByTagAndQualified(curLeads);
  const allDevNames = new Set([...Object.values(DATA.adset_to_dev), ...Object.keys(devLeads)]);
  const devRows = [...allDevNames].map(name => {
    const matchingAdset = Object.entries(DATA.adset_to_dev).find(([, dev]) => dev === name)?.[0];
    const spend = matchingAdset ? (curMeta.byAdset[matchingAdset]?.spend || 0) : 0;
    const hasSpend = !!matchingAdset;
    const l = devLeads[name] || { leads: 0, qualified: 0 };
    return { name, spend, hasSpend, leads: l.leads, qualified: l.qualified };
  }).sort((a, b) => b.spend - a.spend);
  const totalDev = devRows.reduce((acc, r) => ({ spend: acc.spend + r.spend, leads: acc.leads + r.leads, qualified: acc.qualified + r.qualified }), { spend: 0, leads: 0, qualified: 0 });
  document.getElementById('devTable').innerHTML = `
    <thead><tr><th>Desarrollo</th><th>Leads</th><th>Con visita</th><th>Tasa de visita</th><th>Inversión Meta</th><th>Costo / lead</th><th>Costo / lead con visita</th></tr></thead>
    <tbody>
      ${devRows.map(r => `<tr>
        <td class="label-cell">${r.name}</td>
        <td>${fmtInt(r.leads)}</td>
        <td>${fmtInt(r.qualified)}</td>
        <td>${r.leads ? fmtPct(r.qualified / r.leads * 100, 0) : '—'}</td>
        <td>${r.hasSpend ? fmtARS(r.spend) : '<span class="no-data">sin campaña propia</span>'}</td>
        <td>${r.hasSpend && r.leads ? fmtARS(r.spend / r.leads) : '—'}</td>
        <td>${r.hasSpend && r.qualified ? fmtARS(r.spend / r.qualified) : '—'}</td>
      </tr>`).join('')}
      <tr class="total-row"><td class="label-cell">Total</td><td>${fmtInt(totalDev.leads)}</td><td>${fmtInt(totalDev.qualified)}</td>
        <td>${totalDev.leads ? fmtPct(totalDev.qualified / totalDev.leads * 100, 0) : '—'}</td>
        <td>${fmtARS(totalDev.spend)}</td><td>${totalDev.leads ? fmtARS(totalDev.spend / totalDev.leads) : '—'}</td>
        <td>${totalDev.qualified ? fmtARS(totalDev.spend / totalDev.qualified) : '—'}</td></tr>
    </tbody>`;

  // ---- Conversaciones de WhatsApp por grupo de anuncios ----
  const waAdsets = aggregateWhatsappByAdset(filterMetaDaily(r.start, r.end));
  const waEntries = Object.entries(waAdsets).sort((a, b) => b[1].conversations - a[1].conversations);
  const waTotalConv = waEntries.reduce((s, [, a]) => s + a.conversations, 0);
  const waTotalSpend = waEntries.reduce((s, [, a]) => s + a.spend, 0);
  document.getElementById('whatsappTable').innerHTML = waEntries.length ? `
    <thead><tr><th>Grupo de anuncios</th><th>Conversaciones</th><th>% del total</th><th>Inversión</th><th>Costo / conversación</th></tr></thead>
    <tbody>
      ${waEntries.map(([name, a]) => `<tr>
        <td class="label-cell">${name}</td>
        <td>${fmtInt(a.conversations)}</td>
        <td>${waTotalConv ? fmtPct(a.conversations / waTotalConv * 100, 0) : '—'}</td>
        <td>${fmtARS(a.spend)}</td>
        <td>${a.conversations ? fmtARS(a.spend / a.conversations) : '—'}</td>
      </tr>`).join('')}
      <tr class="total-row"><td class="label-cell">Total</td><td>${fmtInt(waTotalConv)}</td><td>100%</td>
        <td>${fmtARS(waTotalSpend)}</td><td>${waTotalConv ? fmtARS(waTotalSpend / waTotalConv) : '—'}</td></tr>
    </tbody>` : `<tbody><tr><td class="empty-note" style="border-bottom:none;">Sin conversaciones en este período.</td></tr></tbody>`;

  // ---- Meta full panel ----
  const metaRows = Object.entries(curMeta.byCampaign).sort((a, b) => b[1].spend - a[1].spend);
  document.getElementById('metaTable').innerHTML = `
    <thead><tr><th>Campaña</th><th>Estado</th><th>Resultado</th><th>Gasto</th><th>CPA</th><th>Impresiones</th><th>Alcance</th><th>CTR</th><th>CPM</th></tr></thead>
    <tbody>
      ${metaRows.map(([name, c]) => {
        const obj = DATA.campaign_objectives[name];
        const field = DATA.result_field[obj];
        const label = DATA.result_label[obj] || 'Resultado';
        const resultVal = field === 'reach' ? c.reach : (field === 'link_clicks' ? c.link_clicks : (field === 'post_engagement' ? c.post_engagement : c.conversations));
        const cpa = resultVal ? c.spend / resultVal : null;
        const ctr = c.impressions ? c.clicks / c.impressions * 100 : null;
        const cpm = c.impressions ? c.spend / c.impressions * 1000 : null;
        const status = (DATA.campaign_status[name] || '').toLowerCase();
        return `<tr>
          <td class="label-cell">${name}</td>
          <td><span class="status-pill ${status === 'active' ? 'active' : 'paused'}">${status === 'active' ? 'Activa' : 'Pausada'}</span></td>
          <td>${fmtInt(resultVal || 0)}<span class="result-caption">${label}</span></td>
          <td>${fmtARS(c.spend)}</td>
          <td>${cpa === null ? '—' : fmtARS(cpa)}</td>
          <td>${fmtInt(c.impressions)}</td>
          <td>${fmtInt(c.reach)}</td>
          <td>${fmtPct(ctr, 2)}</td>
          <td>${cpm === null ? '—' : fmtARS(cpm)}</td>
        </tr>`;
      }).join('')}
    </tbody>`;

  // ---- Leads por desarrollo (list) ----
  const projEntries = Object.entries(curAgg.byTag).sort((a, b) => b[1] - a[1]).slice(0, 8);
  const maxProj = Math.max(1, ...projEntries.map(([, v]) => v));
  document.getElementById('projectRows').innerHTML = projEntries.length ? projEntries.map(([name, count]) => `
    <div class="bar-row">
      <div class="bar-label"><span class="dot"></span>${name}</div>
      <div class="bar-track"><div class="bar-fill" style="width:${count / maxProj * 100}%; background:var(--accent-teal)"></div></div>
      <div class="bar-value">${fmtInt(count)}</div>
    </div>`).join('') : '<p class="empty-note">Sin leads en este período.</p>';
}

const FILTERS = [
  ['allTime', 'Todo el período'], ['yesterday', 'Ayer'], ['last7', 'Últimos 7 días'], ['last30', 'Últimos 30 días'],
  ['lastMonth', 'El mes pasado'], ['mtd', 'Este mes hasta la fecha'],
];
document.getElementById('filterBar').innerHTML = FILTERS.map(([key, label]) =>
  `<button class="filter-btn" data-range="${key}">${label}</button>`).join('');
document.querySelectorAll('.filter-btn').forEach(b => b.addEventListener('click', () => render(b.dataset.range)));
render('allTime');
"""

html = (HTML_TEMPLATE
        .replace("__CSS__", CSS)
        .replace("__FECHA__", fecha_str)
        .replace("__DATA_JSON__", data_json)
        .replace("__JS__", JS))

out_path = os.path.join(BASE, "dashboard.html")
with open(out_path, "w", encoding="utf-8") as f:
    f.write(html)
print("Written:", out_path, len(html), "bytes")

# ---- Artifact-ready variant: strip <!doctype>/<html>/<head>/<body> wrapper tags,
# since the Artifact tool supplies its own skeleton and rejects a full document. ----
artifact_html = html.replace('<!doctype html>\n<html lang="es">\n<head>\n', "")
artifact_html = artifact_html.replace("</head>\n<body>\n", "")
artifact_html = artifact_html.rstrip()
assert artifact_html.endswith("</body>\n</html>"), "unexpected trailing markup, check manually"
artifact_html = artifact_html[: -len("</body>\n</html>")].rstrip()
artifact_path = os.path.join(BASE, "dashboard_artifact.html")
with open(artifact_path, "w", encoding="utf-8") as f:
    f.write(artifact_html)
print("Written:", artifact_path, len(artifact_html), "bytes (Artifact-ready)")
