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
