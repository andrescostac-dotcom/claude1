import json, os, datetime, calendar
import openpyxl

BASE = os.path.dirname(os.path.abspath(__file__))
XLSX_PATH = os.path.join(BASE, "claude_dashboard.xlsx")

WON_ID, LOST_ID = 142, 143
QUALIFIED_IDS = [105609867, 109532768, 105671691, 105671695, WON_ID]  # visita-reunion, reunion realizada, 2da reunion, negociacion, + ganados
# "Visita calificada" (definido con el usuario, 22/8): un sub-conjunto más estricto de
# QUALIFIED_IDS — solo las etapas avanzadas, sin contar la primera visita/reunión.
QUALIFIED_VISIT_IDS = [105671691, 105671695, WON_ID]  # 2da reunion, negociacion, + ganados

ADSET_TO_DEV = {"FAMILIA_CH": "Chubut", "FAMILIA_SI": "Simón Iriondo", "FAMILIA_MIS": "Misiones",
                "FAMILIA_3FEB": "3 de Febrero Lomas"}

# Confirmado con el usuario (22/8): el UTM Content que guarda Kommo no es el mismo string que
# el nombre del adset en Meta, pero corresponde 1 a 1 (o varias variantes de creatividad -> 1
# adset). Mapeo manual para poder cruzar leads/visitas con el gasto real de ese adset.
# "{{adset.name}}" queda deliberadamente afuera: es un parámetro UTM mal configurado en Meta
# Ads Manager (el macro no se resolvió), no hay forma de saber a qué adset correspondía.
UTM_CONTENT_TO_ADSET = {
    "Chubut": "FAMILIA_CH",
    "Chubut_REEL": "FAMILIA_CH",
    "Simon de Iriondo": "FAMILIA_SI",
    "Simon de Iriondo_REEL": "FAMILIA_SI",
    "REEL_MISIONES": "FAMILIA_MIS",
    "REEL_MISIONES_2": "FAMILIA_MIS",
    "FAMILIA": "Prueba: FAMILIA",
    "3 de Febrero_corto": "FAMILIA_3FEB",
}
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
# Lookup por nombre de columna (no por posición) para que agregar/reordenar columnas en el
# Sheet (p.ej. utm_campaign / utm_content) no rompa el parser.
ws = wb["Leads"]
rows = list(ws.iter_rows(values_only=True))
header, lead_rows = rows[0], [r for r in rows[1:] if r[0] is not None]
col = {name: i for i, name in enumerate(header)}


def get(r, name, default=None):
    i = col.get(name)
    return r[i] if i is not None and i < len(r) else default


leads_out = []
statuses = {}
first_seen_status_order = []
for r in lead_rows:
    lid = get(r, "id")
    created_at = get(r, "created_at")
    status_id = int(get(r, "status_id"))
    status_name = get(r, "status_name")
    desarrollos = get(r, "desarrollos")
    price = get(r, "price")
    utm_campaign = (get(r, "utm_campaign") or "").strip()
    utm_content = (get(r, "utm_content") or "").strip()
    casa_visitada = (get(r, "casa_visitada") or "").strip()
    fuente = (get(r, "fuente") or "").strip()
    fecha_visita = get(r, "fecha_visita")
    if status_id not in statuses:
        statuses[status_id] = status_name
        first_seen_status_order.append(status_id)
    if isinstance(created_at, datetime.datetime):
        created_epoch = calendar.timegm(created_at.timetuple())
    else:
        created_epoch = None
    if isinstance(fecha_visita, datetime.datetime):
        fecha_visita_epoch = calendar.timegm(fecha_visita.timetuple())
    else:
        fecha_visita_epoch = None
    tags = [t.strip() for t in desarrollos.split(",")] if desarrollos else []
    leads_out.append({
        "id": int(lid),
        "status_id": status_id,
        "price": price or 0,
        "created_at": created_epoch,
        "tags": tags,
        "utm_campaign": utm_campaign,
        "utm_content": utm_content,
        "casa_visitada": casa_visitada,
        "fuente": fuente,
        "fecha_visita": fecha_visita_epoch,
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
    "qualified_visit_ids": QUALIFIED_VISIT_IDS,
    "adset_to_dev": ADSET_TO_DEV,
    "utm_content_to_adset": UTM_CONTENT_TO_ADSET,
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
  overflow-x: hidden; /* nada debe poder desbordar el ancho de la página, ni en mobile */
}
.wrap { max-width: 1220px; margin: 0 auto; padding: 28px 24px 64px; }
@media (max-width: 640px) { .wrap { padding: 18px 14px 48px; } }
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
  display: flex; align-items: center; gap: 7px; background: var(--surface); border: 1px solid var(--border);
  border-radius: 100px; padding: 6px 12px; font-size: 10.5px; color: var(--ink-muted); max-width: 100%;
  font-family: "IBM Plex Mono", monospace; box-shadow: var(--shadow); white-space: nowrap; box-sizing: border-box;
}
@media (max-width: 480px) {
  .updated-badge { white-space: normal; font-size: 10px; line-height: 1.4; }
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

.kpi-groups { margin-bottom: 14px; }
.kpi-section { margin-bottom: 22px; }
.kpi-section-head { display: flex; align-items: center; gap: 10px; margin: 0 0 10px; }
.kpi-section-title { font-family: "IBM Plex Mono", monospace; font-size: 11px; text-transform: uppercase; letter-spacing: 0.07em; color: var(--ink-muted); font-weight: 500; white-space: nowrap; }
.kpi-section-rule { flex: 1; height: 1px; background: var(--border); }
.kpi-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(210px, 1fr)); gap: 14px; }
.kpi-card { background: var(--surface); border: 1px solid var(--border); border-radius: 16px; padding: 18px 18px 16px; box-shadow: var(--shadow); border-top: 3px solid var(--kpi-accent, var(--border)); min-width: 0; }
@media (max-width: 520px) {
  .kpi-grid { grid-template-columns: repeat(2, 1fr); gap: 10px; }
  .kpi-card { padding: 14px 14px 12px; border-radius: 13px; }
  .kpi-value { font-size: 22px; }
  .kpi-label { font-size: 10.5px; margin-bottom: 6px; }
  .kpi-sub { font-size: 11.5px; }
}
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
/* Sombras de scroll: en mobile una tabla ancha se ve "completa" aunque le falten columnas a la
   derecha — este truco (100% CSS, sin JS) muestra un degradé en el borde que tiene más contenido
   para scrollear, y desaparece solo cuando ya no queda nada para ese lado. */
.table-scroll {
  overflow-x: auto;
  background-color: var(--surface);
  background-image:
    linear-gradient(to right, var(--surface) 60%, transparent),
    linear-gradient(to left, var(--surface) 60%, transparent),
    linear-gradient(to right, rgba(0,0,0,.12), transparent),
    linear-gradient(to left, rgba(0,0,0,.12), transparent);
  background-repeat: no-repeat;
  background-size: 24px 100%, 24px 100%, 10px 100%, 10px 100%;
  background-position: left center, right center, left center, right center;
  background-attachment: local, local, scroll, scroll;
}
table.data-table { width: 100%; border-collapse: collapse; font-size: 13px; min-width: 640px; }
table.data-table th { text-align: left; font-size: 11px; text-transform: uppercase; letter-spacing: 0.04em; color: var(--ink-muted); font-weight: 600; padding: 8px 10px; border-bottom: 1px solid var(--border); white-space: nowrap; }
th.sortable-th { cursor: pointer; user-select: none; min-height: 44px; } /* padding real, no solo bottom, para que el toque no falle */
th.sortable-th:hover { color: var(--ink); }
th.sortable-th::after { content: '⇅'; margin-left: 5px; opacity: .35; font-size: 9px; }
th.sortable-th.sort-asc::after { content: '▲'; opacity: 1; color: var(--accent-blue); }
th.sortable-th.sort-desc::after { content: '▼'; opacity: 1; color: var(--accent-blue); }
table.data-table td { padding: 10px 10px; border-bottom: 1px solid var(--border); font-variant-numeric: tabular-nums; white-space: nowrap; }
table.data-table td.label-cell { font-weight: 600; white-space: normal; font-variant-numeric: initial; }
table.data-table tr:last-child td { border-bottom: none; }
table.data-table tr.total-row td { font-weight: 700; border-top: 1px solid var(--border); }
table.data-table tbody tr:hover td { background: var(--surface-2); }
.status-pill { font-size: 10.5px; font-weight: 700; text-transform: uppercase; letter-spacing: .03em; padding: 2px 7px; border-radius: 100px; }
.status-pill.active { color: var(--good); background: color-mix(in srgb, var(--good) 16%, transparent); }
.status-pill.paused { color: var(--ink-muted); background: var(--surface-2); }
.result-caption { display: block; font-size: 10.5px; color: var(--ink-muted); font-family: "Work Sans", sans-serif; font-weight: 400; }
.no-data { color: var(--ink-muted); }
.resolved-pill { font-size: 11.5px; font-weight: 600; padding: 2px 8px; border-radius: 100px;
  color: var(--accent-blue); background: color-mix(in srgb, var(--accent-blue) 14%, transparent); white-space: nowrap; }
.resolved-pill.unresolved { color: var(--ink-muted); background: var(--surface-2); font-weight: 500;
  font-style: italic; border: 1px dashed var(--border); }

/* Selector de vista tipo "segmented control" — mismo lenguaje visual que los botones de fecha,
   para que se lea de entrada como algo para tocar/clickear, no como una etiqueta de sección. */
.tab-bar {
  display: flex; gap: 4px; margin: 0 0 22px; padding: 4px; width: fit-content; max-width: 100%;
  overflow-x: auto; background: var(--surface-2); border: 1px solid var(--border); border-radius: 100px;
}
.tab-btn {
  font-family: "Work Sans", sans-serif; font-size: 13.5px; font-weight: 600; color: var(--ink-muted);
  background: none; border: none; border-radius: 100px; padding: 10px 18px; min-height: 44px;
  cursor: pointer; white-space: nowrap; transition: background-color .15s, color .15s, box-shadow .15s;
}
.tab-btn:hover { color: var(--ink); }
.tab-btn.active { color: #fff; background: #146F42; box-shadow: var(--shadow); } /* verde fijo (no var(--accent-teal)): así el contraste con blanco queda garantizado en los dos temas */
.tab-panel[hidden] { display: none; }
@media (max-width: 480px) { .tab-btn { font-size: 12.5px; padding: 10px 14px; } }

.chart-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 16px; }
.chart-card { background: var(--surface-2); border: 1px solid var(--border); border-radius: 14px; padding: 16px 16px 12px; }
.chart-card h3 { font-family: "Bricolage Grotesque", sans-serif; font-size: 16px; font-weight: 700; margin: 0 0 3px; }
.chart-caption { font-size: 12.5px; color: var(--ink-muted); margin: 0 0 14px; line-height: 1.45; }
.chart-wrap { position: relative; }
.chart-svg { width: 100%; height: auto; aspect-ratio: 480 / 150; display: block; overflow: visible; }
.chart-line { fill: none; stroke: var(--accent-blue); stroke-width: 2.8; stroke-linejoin: round; stroke-linecap: round; }
.chart-gridline { stroke: var(--border); stroke-width: 1; }
/* Ojo: este texto vive DENTRO del <svg viewBox="0 0 480 150">, que se achica bastante al
   dibujarse en pantalla (una chart-card real mide ~270-380px de ancho, no 480) — por eso el
   font-size acá tiene que ser bastante más grande que el resto de la página para que en
   pantalla termine leyéndose a un tamaño normal, no diminuto. */
.chart-axis-label { font-family: "IBM Plex Mono", monospace; font-size: 20px; font-weight: 500; fill: var(--ink-muted); }
.chart-hover-line { stroke: var(--ink-muted); stroke-width: 1.4; stroke-dasharray: 3 3; }
.chart-hover-dot { fill: var(--accent-blue); stroke: var(--surface-2); stroke-width: 2.5; }
.chart-tooltip { position: absolute; pointer-events: none; background: var(--ink); color: var(--bg);
  font-family: "IBM Plex Mono", monospace; font-size: 13px; font-weight: 600; padding: 6px 11px; border-radius: 7px;
  white-space: nowrap; transform: translate(-50%, -130%); box-shadow: var(--shadow); z-index: 2; }
.chart-empty-note { font-size: 12px; color: var(--ink-muted); padding: 20px 0; text-align: center; }

.reco-list { display: flex; flex-direction: column; gap: 10px; }
.reco-item { display: flex; gap: 10px; align-items: flex-start; padding: 12px 14px; border-radius: 12px;
  background: var(--surface-2); border: 1px solid var(--border); font-size: 13px; line-height: 1.5; }
.reco-icon { font-size: 16px; flex-shrink: 0; line-height: 1.4; }

footer { margin-top: 24px; padding-top: 18px; border-top: 1px solid var(--border); display: flex; justify-content: space-between; flex-wrap: wrap; gap: 10px; font-size: 12px; color: var(--ink-muted); }

/* ---- Pulido visual: los elementos con datos "vivos" (KPIs, panels, chart cards) responden
   un poco al hover para sentirse interactivos, no solo estáticos. ---- */
.kpi-card, .panel, .chart-card { transition: box-shadow .15s, transform .15s; }
.kpi-card:hover, .chart-card:hover { transform: translateY(-1px); box-shadow: 0 2px 4px rgba(27,24,18,0.08), 0 12px 28px -12px rgba(27,24,18,0.22); }
@media (prefers-color-scheme: dark) {
  :root:not([data-theme="light"]) .kpi-card:hover, :root:not([data-theme="light"]) .chart-card:hover { box-shadow: 0 2px 4px rgba(0,0,0,0.35), 0 12px 28px -12px rgba(0,0,0,0.6); }
}
:root[data-theme="dark"] .kpi-card:hover, :root[data-theme="dark"] .chart-card:hover { box-shadow: 0 2px 4px rgba(0,0,0,0.35), 0 12px 28px -12px rgba(0,0,0,0.6); }
"""

HTML_TEMPLATE = """<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex, nofollow, noarchive, nosnippet, noimageindex">
<meta name="googlebot" content="noindex, nofollow">
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

  <div class="tab-bar" id="tabBar" role="tablist">
    <button class="tab-btn active" data-tab="resumen" role="tab" aria-selected="true">Resumen</button>
    <button class="tab-btn" data-tab="costos" role="tab" aria-selected="false">Costos y recomendaciones</button>
  </div>

  <div class="filter-bar" id="filterBar" role="group" aria-label="Rango de fechas"></div>
  <p class="range-caption" id="rangeCaption"></p>

  <div class="tab-panel" id="tab-resumen" role="tabpanel">

    <div class="kpi-groups" id="kpiGrid"></div>

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
      <p class="panel-sub">Conversaciones de WhatsApp iniciadas en Meta Ads (campañas con objetivo de conversión), por conjunto de anuncios · *las columnas de visita resuelven el UTM Content de cada lead en Kommo al conjunto de anuncios real (mapeo manual confirmado) — quedan afuera los leads con UTM sin resolver</p>
      <div class="table-scroll"><table class="data-table" id="whatsappTable"></table></div>
    </div>

    <div class="panel wide-panel">
      <h2>Leads por UTM (Campaign × Content)</h2>
      <p class="panel-sub">Los UTM se guardan en Kommo por lead, tal como llegaron del clic en el anuncio — permiten ver el detalle real de origen incluso para desarrollos sin conjunto de anuncios propio (p. ej. 3 de Febrero Lomas). *Los costos son a nivel del conjunto de anuncios que resuelve el UTM Content — Meta no da spend por creatividad, así que se repiten en las filas que comparten adset.</p>
      <div class="table-scroll"><table class="data-table" id="utmTable"></table></div>
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

  </div>

  <div class="tab-panel" id="tab-costos" role="tabpanel" hidden>

    <div class="panel wide-panel">
      <h2>Recomendaciones del período</h2>
      <p class="panel-sub">Se generan solas según el rango de fechas elegido arriba — funciona como reporte semanal o mensual según qué filtro uses</p>
      <div id="recommendations"></div>
    </div>

    <div class="panel wide-panel">
      <h2>Costo por resultado en el tiempo</h2>
      <p class="panel-sub">Evolución dentro del rango elegido arriba · pasá el mouse por un punto para ver el valor exacto</p>
      <div class="chart-grid" id="costCharts"></div>
    </div>

    <div class="panel wide-panel">
      <h2>Inversión por resultado</h2>
      <p class="panel-sub">Mismos períodos que los gráficos de arriba, con la inversión y las cantidades detrás de cada costo. *Costo/conversación usa solo la inversión de la campaña de conversión por WhatsApp, no el total de Meta Ads.</p>
      <div class="table-scroll"><table class="data-table" id="investmentTable"></table></div>
    </div>

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

// Ordenamiento de tablas al hacer click en el header — genérico, se re-llama cada vez que se
// repuebla una tabla (el listener viejo se descarta solo junto con el <thead> reemplazado).
// Lee el valor real de cada celda desde data-sort (no el texto formateado con $/% ya
// redondeado), así que el orden numérico es siempre correcto. Las filas .total-row quedan
// siempre pegadas abajo.
function makeSortable(tableId) {
  const table = document.getElementById(tableId);
  const thead = table && table.querySelector('thead');
  if (!thead) return;
  const ths = [...thead.querySelectorAll('th')];
  ths.forEach((th, colIdx) => {
    th.classList.add('sortable-th');
    th.addEventListener('click', () => {
      const tbody = table.querySelector('tbody');
      const totalRow = tbody.querySelector('tr.total-row');
      const rows = [...tbody.querySelectorAll('tr')].filter(r => r !== totalRow);
      if (!rows.length) return;
      const asc = th.dataset.sortDir !== 'asc';
      ths.forEach(t => { delete t.dataset.sortDir; t.classList.remove('sort-asc', 'sort-desc'); });
      th.dataset.sortDir = asc ? 'asc' : 'desc';
      th.classList.add(asc ? 'sort-asc' : 'sort-desc');
      rows.sort((rowA, rowB) => {
        const cellA = rowA.children[colIdx], cellB = rowB.children[colIdx];
        const rawA = cellA && cellA.dataset.sort, rawB = cellB && cellB.dataset.sort;
        const numA = parseFloat(rawA), numB = parseFloat(rawB);
        const numeric = rawA !== undefined && rawB !== undefined && (rawA === '' || !isNaN(numA)) && (rawB === '' || !isNaN(numB));
        if (numeric) {
          const aNaN = isNaN(numA), bNaN = isNaN(numB);
          if (aNaN && bNaN) return 0;
          if (aNaN) return 1; // sin dato siempre al final, en cualquier dirección
          if (bNaN) return -1;
          return asc ? numA - numB : numB - numA;
        }
        const textA = (rawA ?? cellA?.textContent ?? '').trim();
        const textB = (rawB ?? cellB?.textContent ?? '').trim();
        return asc ? textA.localeCompare(textB, 'es') : textB.localeCompare(textA, 'es');
      });
      rows.forEach(r => tbody.appendChild(r));
      if (totalRow) tbody.appendChild(totalRow);
    });
  });
}

function aggregateUtm(leads) {
  const byUtm = {};
  for (const l of leads) {
    const campaign = l.utm_campaign || '(sin UTM)';
    const content = l.utm_content || '(sin UTM)';
    const key = campaign + '::' + content;
    if (!byUtm[key]) byUtm[key] = { campaign, content, leads: 0, qualified: 0, qualifiedVisit: 0 };
    byUtm[key].leads++;
    if (DATA.qualified_ids.includes(l.status_id)) byUtm[key].qualified++;
    if (DATA.qualified_visit_ids.includes(l.status_id)) byUtm[key].qualifiedVisit++;
  }
  return Object.values(byUtm);
}

function resolveAdset(utmContent) {
  // El UTM Content que guarda Kommo no es literalmente el nombre del adset en Meta (distintas
  // variantes de creatividad comparten un mismo adset) — DATA.utm_content_to_adset es el mapeo
  // manual confirmado para resolverlo. Sin mapeo, se intenta el nombre tal cual (por si alguna
  // vez coincide directo).
  return DATA.utm_content_to_adset[utmContent] || utmContent;
}

function leadsByUtmContent(leads) {
  // Leads (visita, visita calificada) agrupados por conjunto de anuncios REAL de Meta
  // (resolviendo el UTM Content de cada lead vía resolveAdset) — así "Chubut" + "Chubut_REEL"
  // suman juntos sobre FAMILIA_CH en vez de partir el gasto de ese adset en dos.
  const byContent = {};
  for (const l of leads) {
    if (!l.utm_content) continue;
    const key = resolveAdset(l.utm_content);
    const c = byContent[key] || { leads: 0, visit: 0, qualifiedVisit: 0 };
    c.leads++;
    if (DATA.qualified_ids.includes(l.status_id)) c.visit++;
    if (DATA.qualified_visit_ids.includes(l.status_id)) c.qualifiedVisit++;
    byContent[key] = c;
  }
  return byContent;
}

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
  let won = 0, lost = 0, qualified = 0, qualifiedVisit = 0;
  for (const l of leads) {
    byStatus[l.status_id] = (byStatus[l.status_id] || 0) + 1;
    if (l.status_id === DATA.won_id) won++;
    if (l.status_id === DATA.lost_id) lost++;
    if (DATA.qualified_ids.includes(l.status_id)) qualified++;
    if (DATA.qualified_visit_ids.includes(l.status_id)) qualifiedVisit++;
    const seen = new Set();
    for (const t of l.tags) {
      if (DATA.exclude_tags.includes(t)) continue;
      const c = canonicalTag(t);
      if (seen.has(c)) continue;
      seen.add(c);
      byTag[c] = (byTag[c] || 0) + 1;
    }
  }
  return { total: leads.length, byStatus, byTag, won, lost, qualified, qualifiedVisit };
}

function devLeadsByTagAndQualified(leads) {
  const byTag = {}; // tag -> {leads, qualified, qualifiedVisit}
  for (const l of leads) {
    const qualified = DATA.qualified_ids.includes(l.status_id);
    const qualifiedVisit = DATA.qualified_visit_ids.includes(l.status_id);
    const seen = new Set();
    for (const t of l.tags) {
      if (DATA.exclude_tags.includes(t)) continue;
      const c = canonicalTag(t);
      if (seen.has(c)) continue;
      seen.add(c);
      if (!byTag[c]) byTag[c] = { leads: 0, qualified: 0, qualifiedVisit: 0 };
      byTag[c].leads++;
      if (qualified) byTag[c].qualified++;
      if (qualifiedVisit) byTag[c].qualifiedVisit++;
    }
  }
  return byTag;
}

// ============ Tab "Costos y recomendaciones" ============

function buildBuckets(startKey, endKey) {
  // Diario si el rango elegido es corto (<=21 días), semanal si es más largo — así el
  // gráfico no queda ni vacío (1 punto) ni saturado (100+ puntos) según el filtro activo.
  const totalDays = Math.round((endKey - startKey) / 86400000) + 1;
  const bucketDays = totalDays > 21 ? 7 : 1;
  const buckets = [];
  let cur = startKey;
  while (cur <= endKey) {
    const bEnd = Math.min(addDaysKey(cur, bucketDays - 1), endKey);
    buckets.push({ start: cur, end: bEnd });
    cur = addDaysKey(bEnd, 1);
  }
  return buckets;
}

function bucketLabel(b) {
  return b.start === b.end ? keyToLabel(b.start) : `${keyToLabel(b.start)}–${keyToLabel(b.end)}`;
}

function computeBucketMetrics(b) {
  const leads = filterLeads(b.start, b.end);
  const metaRows = filterMetaDaily(b.start, b.end);
  let spend = 0, waSpend = 0, conversations = 0;
  for (const row of metaRows) {
    spend += row.spend;
    if (row.campaign_name === 'WHATSAPP') { waSpend += row.spend; conversations += row.conversations; }
  }
  let visit = 0, qualifiedVisit = 0;
  for (const l of leads) {
    if (DATA.qualified_ids.includes(l.status_id)) visit++;
    if (DATA.qualified_visit_ids.includes(l.status_id)) qualifiedVisit++;
  }
  return { ...b, leads: leads.length, spend, waSpend, conversations, visit, qualifiedVisit };
}

// Gráfico de línea minimalista en SVG a mano (sin librerías externas) con hover: crosshair +
// tooltip por punto. buckets ya vienen con las métricas calculadas (computeBucketMetrics).
function renderLineChart(containerId, buckets, getValue, fmt) {
  const el = document.getElementById(containerId);
  const points = buckets.map(b => ({ x: (b.start + b.end) / 2, label: bucketLabel(b), value: getValue(b) }));
  const valid = points.filter(p => p.value != null && !isNaN(p.value) && isFinite(p.value));
  if (valid.length < 2) {
    el.innerHTML = '<p class="chart-empty-note">No hay suficientes datos en este período para graficar una tendencia.</p>';
    return;
  }
  const W = 480, H = 150, padL = 70, padR = 10, padT = 24, padB = 26;
  const xs = points.map(p => p.x);
  const minX = Math.min(...xs), maxX = Math.max(...xs);
  const vals = valid.map(p => p.value);
  const minY = 0, maxY = Math.max(...vals) * 1.15 || 1;
  const xScale = x => padL + (x - minX) / ((maxX - minX) || 1) * (W - padL - padR);
  const yScale = y => H - padB - (y - minY) / ((maxY - minY) || 1) * (H - padT - padB);

  let pathD = '';
  points.forEach((p, i) => {
    if (p.value == null) return;
    const cmd = (i === 0 || points[i - 1].value == null) ? 'M' : 'L';
    pathD += `${cmd} ${xScale(p.x).toFixed(1)} ${yScale(p.value).toFixed(1)} `;
  });

  // Eje Y: versión compacta (ej. "$125k") — con el font-size grande que necesita el gráfico
  // para leerse bien, el valor completo ($125.395) no entra en el margen izquierdo.
  const fmtAxis = y => {
    const abs = Math.abs(y);
    if (abs >= 1000000) return '$' + (y / 1000000).toFixed(1).replace(/\.0$/, '') + 'M';
    if (abs >= 1000) return '$' + Math.round(y / 1000) + 'k';
    return fmt(y);
  };
  const gridLines = [0, 0.5, 1].map(f => {
    const y = minY + (maxY - minY) * f;
    const yy = yScale(y).toFixed(1);
    return `<line class="chart-gridline" x1="${padL}" x2="${W - padR}" y1="${yy}" y2="${yy}"/>
      <text class="chart-axis-label" x="${padL - 10}" y="${(+yy + 6).toFixed(1)}" text-anchor="end">${fmtAxis(y)}</text>`;
  }).join('');

  const xLabels = `
    <text class="chart-axis-label" x="${padL}" y="${H - 4}" text-anchor="start">${points[0].label}</text>
    <text class="chart-axis-label" x="${W - padR}" y="${H - 4}" text-anchor="end">${points[points.length - 1].label}</text>`;

  const dotsHtml = points.map((p, i) => p.value == null ? '' :
    `<circle class="chart-pt" data-i="${i}" cx="${xScale(p.x).toFixed(1)}" cy="${yScale(p.value).toFixed(1)}" r="10" fill="transparent"/>`
  ).join('');

  el.innerHTML = `<div class="chart-wrap">
    <svg class="chart-svg" viewBox="0 0 ${W} ${H}">
      ${gridLines}
      <path class="chart-line" d="${pathD.trim()}"/>
      ${xLabels}
      <line class="chart-hover-line" id="${containerId}-hoverline" x1="0" x2="0" y1="${padT}" y2="${H - padB}" style="display:none"/>
      <circle class="chart-hover-dot" id="${containerId}-hoverdot" r="4" style="display:none"/>
      ${dotsHtml}
    </svg>
    <div class="chart-tooltip" id="${containerId}-tooltip" style="display:none"></div>
  </div>`;

  const wrapEl = el.querySelector('.chart-wrap');
  const svgEl = el.querySelector('svg');
  const hoverLine = el.querySelector(`#${containerId}-hoverline`);
  const hoverDot = el.querySelector(`#${containerId}-hoverdot`);
  const tooltip = el.querySelector(`#${containerId}-tooltip`);
  el.querySelectorAll('.chart-pt').forEach(c => {
    const show = () => {
      const i = +c.dataset.i;
      const p = points[i];
      if (p.value == null) return;
      const cx = c.getAttribute('cx'), cy = c.getAttribute('cy');
      hoverLine.setAttribute('x1', cx); hoverLine.setAttribute('x2', cx); hoverLine.style.display = '';
      hoverDot.setAttribute('cx', cx); hoverDot.setAttribute('cy', cy); hoverDot.style.display = '';
      const wrapRect = wrapEl.getBoundingClientRect();
      const svgRect = svgEl.getBoundingClientRect();
      const scaleX = svgRect.width / W, scaleY = svgRect.height / H;
      tooltip.style.left = ((svgRect.left - wrapRect.left) + cx * scaleX) + 'px';
      tooltip.style.top = ((svgRect.top - wrapRect.top) + cy * scaleY) + 'px';
      tooltip.style.display = '';
      tooltip.textContent = `${p.label} · ${fmt(p.value)}`;
    };
    c.addEventListener('mouseenter', show);
    c.addEventListener('mousemove', show);
    c.addEventListener('mouseleave', () => { hoverLine.style.display = 'none'; hoverDot.style.display = 'none'; tooltip.style.display = 'none'; });
  });
}

function renderInvestmentTable(buckets) {
  const el = document.getElementById('investmentTable');
  if (!buckets.length) {
    el.innerHTML = '<tbody><tr><td class="empty-note" style="border-bottom:none;">Sin datos en este período.</td></tr></tbody>';
    return;
  }
  el.innerHTML = `
    <thead><tr>
      <th>Período</th><th>Inversión</th>
      <th>Leads nuevos</th><th>Costo / lead</th>
      <th>Visitas</th><th>Costo / visita</th>
      <th>Visita calificada</th><th>Costo / visita calif.</th>
      <th>Conversaciones</th><th>Costo / conversación*</th>
    </tr></thead>
    <tbody>
      ${buckets.map(b => {
        const costLead = b.leads ? b.spend / b.leads : null;
        const costVisit = b.visit ? b.spend / b.visit : null;
        const costQualVisit = b.qualifiedVisit ? b.spend / b.qualifiedVisit : null;
        const costConv = b.conversations ? b.waSpend / b.conversations : null;
        return `<tr>
        <td class="label-cell" data-sort="${b.start}">${bucketLabel(b)}</td>
        <td data-sort="${b.spend}">${fmtARS(b.spend)}</td>
        <td data-sort="${b.leads}">${fmtInt(b.leads)}</td>
        <td data-sort="${costLead ?? ''}">${costLead != null ? fmtARS(costLead) : '—'}</td>
        <td data-sort="${b.visit}">${fmtInt(b.visit)}</td>
        <td data-sort="${costVisit ?? ''}">${costVisit != null ? fmtARS(costVisit) : '—'}</td>
        <td data-sort="${b.qualifiedVisit}">${fmtInt(b.qualifiedVisit)}</td>
        <td data-sort="${costQualVisit ?? ''}">${costQualVisit != null ? fmtARS(costQualVisit) : '—'}</td>
        <td data-sort="${b.conversations}">${fmtInt(b.conversations)}</td>
        <td data-sort="${costConv ?? ''}">${costConv != null ? fmtARS(costConv) : '—'}</td>
      </tr>`;
      }).join('')}
    </tbody>`;
  makeSortable('investmentTable');
}

// Reglas simples, recalculadas en cada render — funcionan como un mini reporte automático
// que se ajusta solo al rango de fechas elegido (semanal si filtrás "últimos 7 días",
// mensual si filtrás "el mes pasado", etc.).
function computeRecommendations(r, curLeads, curMeta, curAgg, prevAgg, prevMeta) {
  const items = [];
  const MIN_VISITS = 2, MIN_CONVERSATIONS = 5;

  // Mejor / peor desarrollo por costo por visita — exige un piso de inversión además de
  // visitas mínimas, para no comparar un desarrollo casi sin presupuesto de prueba contra
  // uno con inversión real (el "más barato" de un canal apenas testeado no es un ganador
  // confiable, solo tiene pocos datos).
  const MIN_SPEND_FOR_RECO = 15000;
  const devLeads = devLeadsByTagAndQualified(curLeads);
  const devRows = Object.entries(DATA.adset_to_dev).map(([adset, dev]) => {
    const spend = curMeta.byAdset[adset]?.spend || 0;
    const l = devLeads[dev] || { leads: 0, qualified: 0 };
    return { dev, spend, qualified: l.qualified, costVisit: l.qualified ? spend / l.qualified : null };
  }).filter(d => d.spend >= MIN_SPEND_FOR_RECO && d.qualified >= MIN_VISITS);
  if (devRows.length >= 2) {
    const sorted = [...devRows].sort((a, b) => a.costVisit - b.costVisit);
    const best = sorted[0], worst = sorted[sorted.length - 1];
    if (best.dev !== worst.dev && worst.costVisit > best.costVisit) {
      const mult = worst.costVisit / best.costVisit;
      const multLabel = mult >= 20 ? 'varias veces' : `${mult.toFixed(1)}×`;
      items.push({ icon: '🟢', html: `<b>${best.dev}</b> tiene el mejor costo por visita del período: ${fmtARS(best.costVisit)} (${fmtInt(best.qualified)} visitas, ${fmtARS(best.spend)} invertidos). Buen candidato para reforzar presupuesto.` });
      items.push({ icon: '🔴', html: `<b>${worst.dev}</b> tiene el peor costo por visita: ${fmtARS(worst.costVisit)} — ${multLabel} más caro que ${best.dev}. Vale la pena revisar creatividad, segmentación o bajarle presupuesto.` });
    }
  }

  // Mejor / peor grupo de anuncios por costo por conversación de WhatsApp
  const waAdsets = aggregateWhatsappByAdset(filterMetaDaily(r.start, r.end));
  const waRows = Object.entries(waAdsets).filter(([, a]) => a.conversations >= MIN_CONVERSATIONS);
  if (waRows.length >= 2) {
    const sorted = [...waRows].sort((a, b) => (a[1].spend / a[1].conversations) - (b[1].spend / b[1].conversations));
    const [bestName, bestA] = sorted[0], [worstName, worstA] = sorted[sorted.length - 1];
    if (bestName !== worstName) {
      items.push({ icon: '💬', html: `En WhatsApp, <b>${bestName}</b> tiene el costo por conversación más bajo (${fmtARS(bestA.spend / bestA.conversations)}) y <b>${worstName}</b> el más alto (${fmtARS(worstA.spend / worstA.conversations)}).` });
    }
  }

  // Tendencia del costo por visita vs. el período anterior
  if (r.prevStart !== null) {
    const curCV = curAgg.qualified ? curMeta.spend / curAgg.qualified : null;
    const prevCV = prevAgg.qualified ? prevMeta.spend / prevAgg.qualified : null;
    if (curCV != null && prevCV) {
      const change = (curCV - prevCV) / prevCV * 100;
      if (Math.abs(change) >= 15) {
        items.push({ icon: change > 0 ? '📈' : '📉',
          html: `El costo por visita general ${change > 0 ? 'subió' : 'bajó'} ${Math.abs(change).toFixed(0)}% vs. el período anterior (${fmtARS(prevCV)} → ${fmtARS(curCV)}).` });
      }
    }
  }

  if (curAgg.qualified < 5) {
    items.push({ icon: 'ℹ️', html: `Muestra chica en este período (${fmtInt(curAgg.qualified)} visitas) — las recomendaciones son más confiables con más volumen. Probá con un rango más amplio (ej. "Todo el período").` });
  }

  return items;
}

function renderRecommendations(items) {
  const el = document.getElementById('recommendations');
  el.innerHTML = items.length
    ? `<div class="reco-list">${items.map(it => `<div class="reco-item"><span class="reco-icon">${it.icon}</span><span>${it.html}</span></div>`).join('')}</div>`
    : '<p class="empty-note">No hay suficientes datos todavía para generar recomendaciones en este período.</p>';
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
  // "Con visita" / "sin visita" reemplaza al ganado/perdido literal de Kommo (casi no se usan
  // esos 2 estados) — acá "ganado" = llegó a hacer una visita, "perdido" = el resto.
  const noVisit = curAgg.total - curAgg.qualified;
  const qualRate = curAgg.total ? curAgg.qualified / curAgg.total * 100 : null;
  const prevQualRate = prevAgg.total ? prevAgg.qualified / prevAgg.total * 100 : null;
  // "Visita calificada" (22/8): sub-conjunto más estricto de "con visita" — solo 2da reunión,
  // negociación o ganado (sin contar la primera visita/reunión).
  const qualVisitRate = curAgg.total ? curAgg.qualifiedVisit / curAgg.total * 100 : null;
  const prevQualVisitRate = prevAgg.total ? prevAgg.qualifiedVisit / prevAgg.total * 100 : null;
  const whatsappCur = curMeta.byCampaign['WHATSAPP'] || { spend: 0, conversations: 0 };
  const whatsappPrev = prevMeta.byCampaign['WHATSAPP'] || { spend: 0, conversations: 0 };
  const cpcCur = whatsappCur.conversations ? whatsappCur.spend / whatsappCur.conversations : null;
  const cpcPrev = whatsappPrev.conversations ? whatsappPrev.spend / whatsappPrev.conversations : null;
  // Tasa de conversión sobre conversaciones de WhatsApp: de las conversaciones que arrancó
  // Meta, cuántas terminaron siendo un lead con visita en Kommo.
  const waConvRate = whatsappCur.conversations ? curAgg.qualified / whatsappCur.conversations * 100 : null;
  const prevWaConvRate = whatsappPrev.conversations ? prevAgg.qualified / whatsappPrev.conversations * 100 : null;
  // Costo por visita / visita calificada: toda la inversión de Meta del período (no solo
  // WhatsApp) dividida por los leads con visita — cuánto sale, en total, conseguir una.
  const costPerVisitCur = curAgg.qualified ? curMeta.spend / curAgg.qualified : null;
  const costPerVisitPrev = prevAgg.qualified ? prevMeta.spend / prevAgg.qualified : null;
  const costPerQualVisitCur = curAgg.qualifiedVisit ? curMeta.spend / curAgg.qualifiedVisit : null;
  const costPerQualVisitPrev = prevAgg.qualifiedVisit ? prevMeta.spend / prevAgg.qualifiedVisit : null;

  const kpiGroups = [
    { title: 'Leads', accent: 'var(--accent-teal)', items: [
      { label: 'Leads generados', value: fmtInt(curAgg.total), badge: badge(curAgg.total, prevAgg.total, true),
        sub: `${fmtInt(curAgg.qualified)} con visita · ${fmtInt(noVisit)} sin visita` },
      { label: 'Leads con visita', value: fmtInt(curAgg.qualified), badge: badge(curAgg.qualified, prevAgg.qualified, true),
        sub: `${fmtPct(qualRate)} del total · visita, reunión, negociación o ganado` },
      { label: 'Leads con visita calificada', value: fmtInt(curAgg.qualifiedVisit), badge: badge(curAgg.qualifiedVisit, prevAgg.qualifiedVisit, true),
        sub: `${fmtPct(qualVisitRate)} del total · 2da reunión, negociación o ganado` },
    ]},
    { title: 'Tasas de conversión', accent: 'var(--accent-blue)', sortDir: 'desc', items: [
      { label: 'Tasa de conversión (visita)', value: fmtPct(qualRate), badge: qualRate === null ? '' : badge(qualRate, prevQualRate, true),
        sub: 'leads con visita / total de leads', sortValue: qualRate },
      { label: 'Tasa de conversión (visita calificada)', value: fmtPct(qualVisitRate), badge: qualVisitRate === null ? '' : badge(qualVisitRate, prevQualVisitRate, true),
        sub: 'leads con visita calificada / total de leads', sortValue: qualVisitRate },
      { label: 'Tasa de conversión (WhatsApp)', value: fmtPct(waConvRate), badge: waConvRate === null ? '' : badge(waConvRate, prevWaConvRate, true),
        sub: waConvRate === null ? 'sin conversaciones en el período' : `${fmtInt(curAgg.qualified)} con visita / ${fmtInt(whatsappCur.conversations)} conversaciones`, sortValue: waConvRate },
    ]},
    { title: 'Meta Ads', accent: 'var(--good)', items: [
      { label: 'Inversión en Meta Ads', value: fmtARS(curMeta.spend), badge: badge(curMeta.spend, prevMeta.spend, null),
        sub: `${fmtInt(curMeta.impressions)} impresiones` },
      { label: 'Conversaciones de WhatsApp', value: fmtInt(whatsappCur.conversations), badge: badge(whatsappCur.conversations, whatsappPrev.conversations, true),
        sub: 'Meta Ads · campaña de conversión' },
    ]},
    { title: 'Costo por resultado', accent: 'var(--ink-muted)', sortDir: 'asc', items: [
      { label: 'Costo por visita', value: costPerVisitCur === null ? '—' : fmtARS(costPerVisitCur), badge: costPerVisitCur === null ? '' : badge(costPerVisitCur, costPerVisitPrev, false),
        sub: 'inversión total en Meta Ads / leads con visita', sortValue: costPerVisitCur },
      { label: 'Costo por visita calificada', value: costPerQualVisitCur === null ? '—' : fmtARS(costPerQualVisitCur), badge: costPerQualVisitCur === null ? '' : badge(costPerQualVisitCur, costPerQualVisitPrev, false),
        sub: 'inversión total en Meta Ads / leads con visita calificada', sortValue: costPerQualVisitCur },
      { label: 'Costo por conversación', value: cpcCur === null ? '—' : fmtARS(cpcCur), badge: cpcCur === null ? '' : badge(cpcCur, cpcPrev, false),
        sub: 'inversión en WhatsApp / conversaciones', sortValue: cpcCur },
    ]},
  ];
  // "Tasas de conversión" (mayor a menor) y "Costo por resultado" (menor a mayor) se ordenan
  // según el valor real del período (no un orden fijo) — los sin dato (—) siempre al final.
  kpiGroups.forEach(g => {
    if (!g.sortDir) return;
    const sign = g.sortDir === 'desc' ? -1 : 1;
    g.items.sort((a, b) => {
      const av = a.sortValue, bv = b.sortValue;
      if (av == null && bv == null) return 0;
      if (av == null) return 1;
      if (bv == null) return -1;
      return (av - bv) * sign;
    });
  });
  document.getElementById('kpiGrid').innerHTML = kpiGroups.map(g => `
    <div class="kpi-section">
      <div class="kpi-section-head"><span class="kpi-section-title">${g.title}</span><span class="kpi-section-rule"></span></div>
      <div class="kpi-grid" style="--kpi-accent:${g.accent}">
        ${g.items.map(k => `
          <div class="kpi-card">
            <p class="kpi-label">${k.label}</p>
            <div class="kpi-value-row"><span class="kpi-value">${k.value}</span>${k.badge}</div>
            <p class="kpi-sub">${k.sub}</p>
          </div>`).join('')}
      </div>
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
    <span class="pill won"><span class="dotp"></span>${fmtInt(curAgg.qualified)} con visita</span>
    <span class="pill lost"><span class="dotp"></span>${fmtInt(noVisit)} sin visita</span>`;

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
    const l = devLeads[name] || { leads: 0, qualified: 0, qualifiedVisit: 0 };
    return { name, spend, hasSpend, leads: l.leads, qualified: l.qualified, qualifiedVisit: l.qualifiedVisit };
  }).sort((a, b) => b.spend - a.spend);
  const totalDev = devRows.reduce((acc, r) => ({
    spend: acc.spend + r.spend, leads: acc.leads + r.leads, qualified: acc.qualified + r.qualified, qualifiedVisit: acc.qualifiedVisit + r.qualifiedVisit,
  }), { spend: 0, leads: 0, qualified: 0, qualifiedVisit: 0 });
  document.getElementById('devTable').innerHTML = `
    <thead><tr><th>Desarrollo</th><th>Leads</th><th>Con visita</th><th>Tasa de visita</th><th>Inversión Meta</th><th>Costo / lead</th><th>Costo / lead con visita</th><th>Costo / lead visita calificada</th></tr></thead>
    <tbody>
      ${devRows.map(r => {
        const rate = r.leads ? r.qualified / r.leads * 100 : null;
        const costLead = r.hasSpend && r.leads ? r.spend / r.leads : null;
        const costQual = r.hasSpend && r.qualified ? r.spend / r.qualified : null;
        const costQualVisit = r.hasSpend && r.qualifiedVisit ? r.spend / r.qualifiedVisit : null;
        return `<tr>
        <td class="label-cell" data-sort="${r.name}">${r.name}</td>
        <td data-sort="${r.leads}">${fmtInt(r.leads)}</td>
        <td data-sort="${r.qualified}">${fmtInt(r.qualified)}</td>
        <td data-sort="${rate ?? ''}">${rate != null ? fmtPct(rate, 0) : '—'}</td>
        <td data-sort="${r.hasSpend ? r.spend : ''}">${r.hasSpend ? fmtARS(r.spend) : '<span class="no-data">sin campaña propia</span>'}</td>
        <td data-sort="${costLead ?? ''}">${costLead != null ? fmtARS(costLead) : '—'}</td>
        <td data-sort="${costQual ?? ''}">${costQual != null ? fmtARS(costQual) : '—'}</td>
        <td data-sort="${costQualVisit ?? ''}">${costQualVisit != null ? fmtARS(costQualVisit) : '—'}</td>
      </tr>`;
      }).join('')}
      <tr class="total-row"><td class="label-cell">Total</td><td>${fmtInt(totalDev.leads)}</td><td>${fmtInt(totalDev.qualified)}</td>
        <td>${totalDev.leads ? fmtPct(totalDev.qualified / totalDev.leads * 100, 0) : '—'}</td>
        <td>${fmtARS(totalDev.spend)}</td><td>${totalDev.leads ? fmtARS(totalDev.spend / totalDev.leads) : '—'}</td>
        <td>${totalDev.qualified ? fmtARS(totalDev.spend / totalDev.qualified) : '—'}</td>
        <td>${totalDev.qualifiedVisit ? fmtARS(totalDev.spend / totalDev.qualifiedVisit) : '—'}</td></tr>
    </tbody>`;
  makeSortable('devTable');

  // ---- Conversaciones de WhatsApp por grupo de anuncios ----
  const waAdsets = aggregateWhatsappByAdset(filterMetaDaily(r.start, r.end));
  const waVisits = leadsByUtmContent(curLeads); // cruce por utm_content === adset_name
  const waEntries = Object.entries(waAdsets).sort((a, b) => b[1].conversations - a[1].conversations);
  const waTotalConv = waEntries.reduce((s, [, a]) => s + a.conversations, 0);
  const waTotalSpend = waEntries.reduce((s, [, a]) => s + a.spend, 0);
  const waTotalVisits = waEntries.reduce((s, [name]) => s + (waVisits[name]?.visit || 0), 0);
  const waTotalQualVisits = waEntries.reduce((s, [name]) => s + (waVisits[name]?.qualifiedVisit || 0), 0);
  document.getElementById('whatsappTable').innerHTML = waEntries.length ? `
    <thead><tr>
      <th>Grupo de anuncios</th><th>Conversaciones</th><th>% del total</th>
      <th>Inversión</th><th>Costo / conversación</th>
      <th>Con visita*</th><th>Costo / visita*</th>
      <th>Visita calificada*</th><th>Costo / visita calificada*</th>
    </tr></thead>
    <tbody>
      ${waEntries.map(([name, a]) => {
        const v = waVisits[name];
        const pct = waTotalConv ? a.conversations / waTotalConv * 100 : null;
        const cpc = a.conversations ? a.spend / a.conversations : null;
        const costVisit = v && v.visit ? a.spend / v.visit : null;
        const costQualVisit = v && v.qualifiedVisit ? a.spend / v.qualifiedVisit : null;
        return `<tr>
        <td class="label-cell" data-sort="${name}">${name}</td>
        <td data-sort="${a.conversations}">${fmtInt(a.conversations)}</td>
        <td data-sort="${pct ?? ''}">${pct != null ? fmtPct(pct, 0) : '—'}</td>
        <td data-sort="${a.spend}">${fmtARS(a.spend)}</td>
        <td data-sort="${cpc ?? ''}">${cpc != null ? fmtARS(cpc) : '—'}</td>
        <td data-sort="${v ? v.visit : ''}">${v ? fmtInt(v.visit) : '<span class="no-data">sin UTM</span>'}</td>
        <td data-sort="${costVisit ?? ''}">${costVisit != null ? fmtARS(costVisit) : '—'}</td>
        <td data-sort="${v ? v.qualifiedVisit : ''}">${v ? fmtInt(v.qualifiedVisit) : '<span class="no-data">sin UTM</span>'}</td>
        <td data-sort="${costQualVisit ?? ''}">${costQualVisit != null ? fmtARS(costQualVisit) : '—'}</td>
      </tr>`;
      }).join('')}
      <tr class="total-row"><td class="label-cell">Total</td><td>${fmtInt(waTotalConv)}</td><td>100%</td>
        <td>${fmtARS(waTotalSpend)}</td><td>${waTotalConv ? fmtARS(waTotalSpend / waTotalConv) : '—'}</td>
        <td>${fmtInt(waTotalVisits)}</td><td>${waTotalVisits ? fmtARS(waTotalSpend / waTotalVisits) : '—'}</td>
        <td>${fmtInt(waTotalQualVisits)}</td><td>${waTotalQualVisits ? fmtARS(waTotalSpend / waTotalQualVisits) : '—'}</td></tr>
    </tbody>` : `<tbody><tr><td class="empty-note" style="border-bottom:none;">Sin conversaciones en este período.</td></tr></tbody>`;
  makeSortable('whatsappTable');

  // ---- Leads por UTM (campaign x content) ----
  // Leads/visitas: nivel de detalle real por UTM Content (creatividad). Costo/visita, costo/lead
  // calificado y costo/iniciar conversación: Meta solo da spend y conversaciones por ADSET, no
  // por creatividad — así que esas 3 columnas se calculan a nivel del adset resuelto (mismo
  // valor repetido en todas las filas que comparten adset) para no partir su gasto entre ellas.
  const utmRows = aggregateUtm(curLeads).sort((a, b) => b.leads - a.leads);
  const utmAdsetAgg = leadsByUtmContent(curLeads); // { [adset resuelto]: {leads, visit, qualifiedVisit} }
  document.getElementById('utmTable').innerHTML = utmRows.length ? `
    <thead><tr>
      <th>UTM Content</th><th>Conjunto de anuncios (Meta)</th>
      <th>Leads</th><th>Con visita</th><th>Visita calificada</th>
      <th>Costo / visita*</th><th>Costo / visita calificada*</th><th>Costo / conversación*</th>
    </tr></thead>
    <tbody>
      ${utmRows.map(u => {
        const resolved = u.content === '(sin UTM)' ? null : resolveAdset(u.content);
        const metaAdset = resolved ? curMeta.byAdset[resolved] : null;
        const agg = resolved ? utmAdsetAgg[resolved] : null;
        const costVisit = metaAdset && agg && agg.visit ? metaAdset.spend / agg.visit : null;
        const costQualVisit = metaAdset && agg && agg.qualifiedVisit ? metaAdset.spend / agg.qualifiedVisit : null;
        const costConv = metaAdset && metaAdset.conversations ? metaAdset.spend / metaAdset.conversations : null;
        return `<tr>
          <td class="label-cell" data-sort="${u.content}">${u.content}</td>
          <td data-sort="${resolved || ''}">${metaAdset ? `<span class="resolved-pill">${resolved}</span>` : '<span class="resolved-pill unresolved">sin resolver</span>'}</td>
          <td data-sort="${u.leads}">${fmtInt(u.leads)}</td>
          <td data-sort="${u.qualified}">${fmtInt(u.qualified)}</td>
          <td data-sort="${u.qualifiedVisit}">${fmtInt(u.qualifiedVisit)}</td>
          <td data-sort="${costVisit ?? ''}">${costVisit != null ? fmtARS(costVisit) : '—'}</td>
          <td data-sort="${costQualVisit ?? ''}">${costQualVisit != null ? fmtARS(costQualVisit) : '—'}</td>
          <td data-sort="${costConv ?? ''}">${costConv != null ? fmtARS(costConv) : '—'}</td>
        </tr>`;
      }).join('')}
      ${(() => {
        // Fila de totales: suma simple para leads/visitas (son por-lead, no se duplican), pero
        // para los $ hay que sumar el spend/conversaciones de cada adset UNA sola vez (si no,
        // un adset compartido por 2 filas de UTM se contaría el doble).
        const totLeads = utmRows.reduce((s, u) => s + u.leads, 0);
        const totVisit = utmRows.reduce((s, u) => s + u.qualified, 0);
        const totQualVisit = utmRows.reduce((s, u) => s + u.qualifiedVisit, 0);
        const distinctAdsets = [...new Set(utmRows
          .map(u => u.content === '(sin UTM)' ? null : resolveAdset(u.content))
          .filter(a => a && curMeta.byAdset[a]))];
        const totSpend = distinctAdsets.reduce((s, a) => s + curMeta.byAdset[a].spend, 0);
        const totConv = distinctAdsets.reduce((s, a) => s + curMeta.byAdset[a].conversations, 0);
        const totVisitMatched = distinctAdsets.reduce((s, a) => s + (utmAdsetAgg[a]?.visit || 0), 0);
        const totQualVisitMatched = distinctAdsets.reduce((s, a) => s + (utmAdsetAgg[a]?.qualifiedVisit || 0), 0);
        return `<tr class="total-row"><td class="label-cell">Total</td><td></td>
          <td>${fmtInt(totLeads)}</td><td>${fmtInt(totVisit)}</td><td>${fmtInt(totQualVisit)}</td>
          <td>${totVisitMatched ? fmtARS(totSpend / totVisitMatched) : '—'}</td>
          <td>${totQualVisitMatched ? fmtARS(totSpend / totQualVisitMatched) : '—'}</td>
          <td>${totConv ? fmtARS(totSpend / totConv) : '—'}</td></tr>`;
      })()}
    </tbody>` : `<tbody><tr><td class="empty-note" style="border-bottom:none;">Sin leads en este período.</td></tr></tbody>`;
  makeSortable('utmTable');

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
          <td class="label-cell" data-sort="${name}">${name}</td>
          <td data-sort="${status}"><span class="status-pill ${status === 'active' ? 'active' : 'paused'}">${status === 'active' ? 'Activa' : 'Pausada'}</span></td>
          <td data-sort="${resultVal || 0}">${fmtInt(resultVal || 0)}<span class="result-caption">${label}</span></td>
          <td data-sort="${c.spend}">${fmtARS(c.spend)}</td>
          <td data-sort="${cpa ?? ''}">${cpa === null ? '—' : fmtARS(cpa)}</td>
          <td data-sort="${c.impressions}">${fmtInt(c.impressions)}</td>
          <td data-sort="${c.reach}">${fmtInt(c.reach)}</td>
          <td data-sort="${ctr ?? ''}">${fmtPct(ctr, 2)}</td>
          <td data-sort="${cpm ?? ''}">${cpm === null ? '—' : fmtARS(cpm)}</td>
        </tr>`;
      }).join('')}
    </tbody>`;
  makeSortable('metaTable');

  // ---- Leads por desarrollo (list) ----
  const projEntries = Object.entries(curAgg.byTag).sort((a, b) => b[1] - a[1]).slice(0, 8);
  const maxProj = Math.max(1, ...projEntries.map(([, v]) => v));
  document.getElementById('projectRows').innerHTML = projEntries.length ? projEntries.map(([name, count]) => `
    <div class="bar-row">
      <div class="bar-label"><span class="dot"></span>${name}</div>
      <div class="bar-track"><div class="bar-fill" style="width:${count / maxProj * 100}%; background:var(--accent-teal)"></div></div>
      <div class="bar-value">${fmtInt(count)}</div>
    </div>`).join('') : '<p class="empty-note">Sin leads en este período.</p>';

  // ---- Tab "Costos y recomendaciones" ----
  const buckets = buildBuckets(r.start, r.end).map(computeBucketMetrics);
  document.getElementById('costCharts').innerHTML = `
    <div class="chart-card"><h3>Costo por lead nuevo</h3><p class="chart-caption">Inversión total en Meta Ads / leads generados</p><div id="chartLead"></div></div>
    <div class="chart-card"><h3>Costo por visita</h3><p class="chart-caption">⚠️ los últimos períodos quedan subestimados: un lead recién creado todavía no tuvo tiempo de llegar a "visita"</p><div id="chartVisit"></div></div>
    <div class="chart-card"><h3>Costo por visita calificada</h3><p class="chart-caption">⚠️ mismo efecto de rezago que "costo por visita", más marcado por ser una etapa más avanzada</p><div id="chartQualVisit"></div></div>
    <div class="chart-card"><h3>Costo por conversación iniciada</h3><p class="chart-caption">Inversión en la campaña de WhatsApp / conversaciones iniciadas</p><div id="chartConv"></div></div>`;
  renderLineChart('chartLead', buckets, b => b.leads ? b.spend / b.leads : null, fmtARS);
  renderLineChart('chartVisit', buckets, b => b.visit ? b.spend / b.visit : null, fmtARS);
  renderLineChart('chartQualVisit', buckets, b => b.qualifiedVisit ? b.spend / b.qualifiedVisit : null, fmtARS);
  renderLineChart('chartConv', buckets, b => b.conversations ? b.waSpend / b.conversations : null, fmtARS);
  renderInvestmentTable(buckets);
  renderRecommendations(computeRecommendations(r, curLeads, curMeta, curAgg, prevAgg, prevMeta));
}

const FILTERS = [
  ['allTime', 'Todo el período'], ['yesterday', 'Ayer'], ['last7', 'Últimos 7 días'], ['last30', 'Últimos 30 días'],
  ['lastMonth', 'El mes pasado'], ['mtd', 'Este mes hasta la fecha'],
];
document.getElementById('filterBar').innerHTML = FILTERS.map(([key, label]) =>
  `<button class="filter-btn" data-range="${key}">${label}</button>`).join('');
document.querySelectorAll('.filter-btn').forEach(b => b.addEventListener('click', () => render(b.dataset.range)));

document.querySelectorAll('.tab-btn').forEach(btn => btn.addEventListener('click', () => {
  document.querySelectorAll('.tab-btn').forEach(b => { b.classList.toggle('active', b === btn); b.setAttribute('aria-selected', b === btn ? 'true' : 'false'); });
  document.querySelectorAll('.tab-panel').forEach(p => { p.hidden = p.id !== `tab-${btn.dataset.tab}`; });
}));

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
