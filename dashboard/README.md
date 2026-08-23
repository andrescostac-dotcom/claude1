# CosCor — Dashboard de marketing (Kommo + Meta Ads)

## Cómo está armado

1. **Google Apps Script** (`apps-script-sync.gs`, instalado por el usuario en su
   Google Sheet) corre todos los días a las 8am ARG, pega las credenciales de
   Kommo CRM y Meta Ads (guardadas en Propiedades del Script de Google, nunca
   visibles para Claude) y escribe 3 pestañas en la Sheet: `Leads`,
   `MetaDaily`, `Meta` (timestamp del último sync).

2. **La Google Sheet** ("Claude Dashboard", propiedad de info@coscor.life):
   - fileId: `1WzajqMMHfH8YidFIkZJI328_Rf2FPl6f7-zzDf-WnEs`

3. **`build_dashboard.py`** en esta carpeta lee un `.xlsx` exportado de esa
   Sheet (`claude_dashboard.xlsx`, no versionado) y genera:
   - `dashboard.html` — página completa standalone (para abrir directo en un navegador)
   - `dashboard_artifact.html` — la misma página sin los tags `<!doctype>/<html>/<head>/<body>`, lista para publicar con la tool `Artifact`

4. **El Artifact publicado** (el link que ve el usuario):
   - URL: `https://claude.ai/code/artifact/3e96904c-f359-4169-aef1-103e0ad90659`

## Procedimiento para refrescar el dashboard (correr esto en cada firing de la Routine diaria)

1. Descargar la Sheet como xlsx vía el connector de Google Drive:
   `mcp__Google_Drive__download_file_content` con
   `fileId=1WzajqMMHfH8YidFIkZJI328_Rf2FPl6f7-zzDf-WnEs` y
   `exportMimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`.
   La respuesta es JSON `{content: <base64>, ...}` — puede exceder el límite
   de tokens de un solo tool result; si eso pasa, el resultado queda guardado
   en un archivo local (el error lo indica) — leer ese archivo con Python,
   decodificar el campo `content` en base64 y escribirlo en
   `dashboard/claude_dashboard.xlsx`.
2. `pip install -q openpyxl` si no está instalado.
3. `cd dashboard && python3 build_dashboard.py` — imprime cuántos leads/filas
   de Meta procesó y el `generated_at` (debería ser el de la corrida más
   reciente del Apps Script, no uno viejo — si no cambió respecto a la vez
   anterior, el Apps Script no corrió y hay que avisarle al usuario).
4. Publicar `dashboard/dashboard_artifact.html` con la tool `Artifact`,
   pasando `url` = la URL del Artifact de arriba (así actualiza el mismo
   link en vez de crear uno nuevo), y el mismo `favicon` (📈) y `title`
   (Dashboard CosCor) que ya tiene.
5. No hace falta comentar nada en el chat si todo salió bien — es una
   actualización silenciosa. Si algo falla (Sheet sin cambios hace varios
   días, error de Drive, etc.) avisar al usuario.

## Métricas clave (agregado 22/8)

- **Leads con visita**: KPI + tabla "Leads con visita por desarrollo × inversión
  en Meta". Definición: `status_id` en `visita - reunion`, `Reunion realizada`,
  `2da Reunion`, `Negociacion` o `Logrado con éxito` (constante
  `QUALIFIED_IDS` en `build_dashboard.py`, la misma que ya existía como
  "leads calificados"). Se muestra el conteo, el % sobre el total de leads del
  período, y el desglose por desarrollo/grupo de anuncios (vía el tag de
  Kommo, cruzado con `ADSET_TO_DEV`).
- **Conversaciones de WhatsApp por grupo de anuncios**: nueva tabla, filtra
  `meta_daily` a filas cuya campaña tiene `objective = OUTCOME_LEADS` (la
  campaña de conversión por WhatsApp) y agrupa por `adset_name`.
- **Filtro "Todo el período"**: nuevo, y ahora es el default al abrir el
  dashboard. Calcula el rango dinámicamente como la fecha más antigua entre
  todos los leads y todas las filas de Meta hasta hoy — no hay comparación
  "vs. período anterior" en esta vista (no tiene sentido para un acumulado
  histórico), los badges de variación se ocultan.
- **Historial completo**: Kommo ya traía todos los leads sin filtro de fecha
  (la paginación de `syncKommoLeads` no tiene límite). Meta sí estaba
  limitado a `META_DAYS_BACK = 120` días rodantes — se cambió a
  `META_SINCE_DATE = '2020-01-01'` en `apps-script-sync.gs` para traer todo
  el histórico que la cuenta de Meta tenga disponible, de una vez y para
  siempre (no vuelve a truncar con el paso del tiempo). Requiere que el
  usuario vuelva a pegar el `.gs` actualizado en su Apps Script y corra
  `syncAll` una vez.

## Cambios (22/8, ronda 2)

- **KPIs**: "Leads generados" ahora muestra con/sin visita en vez de
  ganado/perdido literal de Kommo. Se agregaron "Tasa de conversión
  (visitas)" (= leads con visita / total) y "Tasa de conversión (WhatsApp)"
  (= leads con visita / conversaciones de WhatsApp de Meta) — 6 KPIs en
  total.
- **Embudo**: los pills "logrado con éxito" / "venta perdida" (que casi no
  se usan como estados literales en Kommo) se reemplazaron por "con visita"
  / "sin visita", calculados con la misma definición que el resto del
  dashboard (`QUALIFIED_IDS`).
- **UTM**: `apps-script-sync.gs` ahora extrae `utm_campaign` y `utm_content`
  de `custom_fields_values` de cada lead en Kommo (matching flexible por
  nombre/código de campo, insensible a mayúsculas/espacios). Nueva tabla
  "Leads por UTM (Campaign × Content)": leads, con visita, tasa de visita, y
  cruce con inversión de Meta cuando el `UTM Content` coincide exactamente
  con el nombre de un conjunto de anuncios (`adset_name`). Objetivo:
  detectar desarrollos como "3 de Febrero Lomas" que no tienen conjunto de
  anuncios propio en `ADSET_TO_DEV` pero sí quedan identificados por UTM.
  **Requiere volver a correr el Apps Script** — hasta que el usuario lo
  haga, todos los leads existentes van a aparecer como "(sin UTM)" porque
  esa columna no existía antes en la Sheet.
  - Riesgo conocido: no verifiqué en vivo que Kommo efectivamente tenga
    campos personalizados llamados "UTM Campaign"/"UTM Content" en esta
    cuenta (no hay acceso a la API desde este entorno) — el matching es
    flexible pero si el nombre real difiere del patrón `utmcampaign`/
    `utmcontent` (sin espacios/guiones/mayúsculas) no va a encontrarlos. Si
    después de correr el script la tabla sigue en "(sin UTM)", hay que
    revisar el nombre exacto del campo en Kommo (Ajustes → Campos
    personalizados → Leads) y ajustar `extractCustomField`.

## Tab "Costos y recomendaciones" (agregado 23/8)

- **Gráficos de costo en el tiempo** (Costo por lead nuevo / por visita / por visita
  calificada / por conversación iniciada): se buildean con `buildBuckets` (diario si
  el rango ≤21 días, semanal si es más largo) + `computeBucketMetrics`, y usan los
  mismos filtros de fecha que el resto del dashboard (comparten `render()`). Son SVG
  a mano, sin librerías — línea + hover con crosshair y tooltip.
  ⚠️ "Costo por visita" y "costo por visita calificada" tienen un sesgo de rezago
  real: agrupan por `created_at` del lead, no por la fecha en que llegó a esa etapa
  (ese dato no existe), así que los períodos más recientes siempre van a verse
  artificialmente baratos/vacíos — un lead de ayer todavía no tuvo tiempo de llegar
  a "2da reunión". Hay un aviso (⚠️) en la tarjeta de cada uno de estos 2 gráficos.
- **Tabla "Inversión por resultado"**: mismos buckets que los gráficos, con inversión
  + cantidad + costo unitario por columna. Costo/conversación usa el spend de la
  campaña WHATSAPP únicamente (no el total de Meta Ads), igual que en el resto del
  dashboard.
- **Recomendaciones**: reglas simples recalculadas en cada render (no hay IA de por
  medio) — mejor/peor desarrollo por costo por visita (con piso de $15.000 de
  inversión y mínimo 2 visitas, para no comparar un desarrollo casi sin testear
  contra uno con presupuesto real), mejor/peor grupo de anuncios por costo de
  conversación de WhatsApp (mínimo 5 conversaciones), y variación del costo por
  visita general vs. el período anterior si supera ±15%. Al cambiar el filtro de
  fecha arriba, se recalculan solas — "Últimos 7 días" da un reporte semanal,
  "El mes pasado" uno mensual, etc.

## Campos nuevos de Kommo (agregado 23/8, ronda 2)

Se agregaron 3 campos personalizados de Kommo, identificados por `field_id` (no por
nombre, para evitar la ambigüedad que tuvo el matching flexible de UTM):

- **Casa Visitada** (`field_id` 2444391) → columna `casa_visitada`
- **Fuente** (`field_id` 578180) → columna `fuente`
- **Fecha Visita** (`field_id` 576940, campo de tipo fecha) → columna `fecha_visita`

`apps-script-sync.gs` ahora los extrae con `extractCustomFieldById` /
`extractCustomFieldDateById` y los escribe como 3 columnas nuevas en la pestaña `Leads`.
`build_dashboard.py` ya los lee y los suma a cada lead en el payload (`casa_visitada`,
`fuente`, `fecha_visita` — este último como epoch, `null` si el lead no tiene fecha de
visita cargada). Todavía no arman ninguna tabla/gráfico nuevo en el HTML — falta ver
datos reales (no hay acceso a la API de Kommo desde este entorno, así que no pude
confirmar en vivo qué valores devuelve cada campo: si "Fuente" es de selección simple o
múltiple, qué opciones tiene "Casa Visitada", etc.).

**Para activarlo**: pegar el `.gs` actualizado en el editor de Apps Script (reemplaza el
archivo entero, igual que la vez pasada) y correr `syncAll` una vez. Los leads que no
tengan estos campos cargados en Kommo van a aparecer con valor vacío — no rompe nada.

Ideas ya conversadas para una vez que haya datos reales:
- **Fuente**: tabla de leads/visita/visita calificada por fuente (igual estructura que
  la tabla de UTM).
- **Casa Visitada**: breakdown de qué casa/modelo se visita más.
- **Fecha Visita**: resuelve el sesgo de rezago documentado abajo en "Costo por visita" —
  se podría usar la fecha real de la visita en vez de `created_at` del lead para esos 2
  gráficos, cuando el lead la tenga cargada (fallback a `created_at` si no).

## Nota sobre el orden del funnel

`build_dashboard.py` define `STATUS_ORDER_HINT`: un orden inferido de los
estados del pipeline de Kommo, porque el Apps Script actual no exporta el
campo `sort` real de cada estado (solo lo tenía la versión Python original,
`fetch_data.py`, que ya no corre). Es una aproximación razonable, no exacta.
Si en algún momento se quiere el orden exacto, hay que sumar al
`apps-script-sync.gs` una pestaña `Statuses` con `id, name, sort` (se obtiene
del mismo pipeline que ya se pide en `syncKommoLeads`) y leerla acá en vez de
usar el hint.

## Credenciales

Los tokens de Kommo y Meta Ads **nunca pasan por este repo, por Claude, ni
por este chat** — viven únicamente en las Propiedades del Script de Google
Apps Script, del lado del usuario. Este flujo entero (repo + Routine) solo
toca la Google Sheet ya sincronizada, vía el connector de Google Drive de la
cuenta del usuario.
