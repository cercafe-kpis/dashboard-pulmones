# Score Pulmonar — Dashboard de Análisis

Aplicación web de una sola página (`index.html`, sin backend propio, sin proceso de build) para visualizar y analizar los resultados de inspección pulmonar porcina capturados con el método **Madec**. Los datos se capturan en campo con la app complementaria **"Inspección Pulmonar"** y se almacenan en una lista de SharePoint; este dashboard los lee, calcula los indicadores y los presenta en gráficas, tablas, PDF y Excel.

Hace parte del ecosistema de herramientas de Cercafe para monitoreo de producción porcina (Cercafe KPIs).

---

## Índice

1. [Arquitectura general](#arquitectura-general)
2. [Configuración](#configuración)
3. [Autenticación y origen de datos](#autenticación-y-origen-de-datos)
4. [Control de acceso por asociado](#control-de-acceso-por-asociado)
5. [Estructura de la interfaz](#estructura-de-la-interfaz)
6. [Metodología y cálculos](#metodología-y-cálculos)
7. [Exportación a PDF](#exportación-a-pdf)
8. [Exportación a Excel](#exportación-a-excel)
9. [Campos esperados en la lista de SharePoint](#campos-esperados-en-la-lista-de-sharepoint)
10. [Despliegue](#despliegue)
11. [Decisiones de diseño y notas para mantenimiento](#decisiones-de-diseño-y-notas-para-mantenimiento)
12. [Mapa de funciones clave](#mapa-de-funciones-clave)

---

## Arquitectura general

Todo el proyecto vive en **un único archivo `index.html`**: HTML, CSS y JavaScript en el mismo documento. No hay paso de compilación ni framework — es HTML/CSS/JS plano, pensado para desplegarse tal cual en GitHub Pages (o cualquier hosting estático).

Librerías externas cargadas por CDN (`unpkg.com`):
- **`@azure/msal-browser`** — autenticación contra Microsoft Entra ID (Azure AD).
- **`xlsx` (SheetJS)** — generación de archivos `.xlsx` para la exportación a Excel.

No hay base de datos propia: los datos viven en una **lista de SharePoint** y se consultan en vivo, desde el navegador del usuario, vía **Microsoft Graph API**.

```
┌─────────────────┐        ┌──────────────────┐        ┌───────────────────────┐
│  App de captura   │ ---> │  Lista SharePoint  │ <--- │  Este dashboard        │
│ "Inspección        │      │ "InspeccionPulmones"│      │ (index.html)          │
│  Pulmonar"          │      │                    │      │ lee vía Graph API,     │
│  (PWA en celular)  │      │                    │      │ calcula y visualiza    │
└─────────────────┘        └──────────────────┘        └───────────────────────┘
```

---

## Configuración

Toda la configuración específica del despliegue está centralizada al inicio del `<script>` principal:

```javascript
const CONFIG = {
  CLIENT_ID: '5d4e5dd1-82f1-48d0-a0b0-f104d938c770',   // App Registration en Azure AD
  TENANT_ID: 'f2ea1671-e87d-4f02-9941-b721acbdbc01',   // Tenant de Microsoft 365 de Cercafe
  SP_SITE_URL: 'https://cercafe.sharepoint.com/sites/Inspeccion_pulmones',
  LIST_NAME: 'InspeccionPulmones'
};
```

| Campo | Qué es | Cuándo cambiarlo |
|---|---|---|
| `CLIENT_ID` | ID de la App Registration de Azure AD que autoriza el login | Solo si se crea una nueva App Registration |
| `TENANT_ID` | ID del tenant de Microsoft 365 | Solo si cambia de organización |
| `SP_SITE_URL` | URL del sitio de SharePoint donde vive la lista | Si se mueve la lista a otro sitio |
| `LIST_NAME` | Nombre visible de la lista dentro de ese sitio | Si se renombra la lista |

**Importante:** el `redirectUri` de MSAL se calcula dinámicamente a partir de `location.origin + location.pathname` — no hay que configurarlo a mano, pero sí hay que registrar la URL exacta donde se publique el dashboard como *Redirect URI* en la App Registration de Azure AD (portal de Azure → App registrations → Authentication).

---

## Autenticación y origen de datos

1. **Login** (`doLogin()`): abre un popup de Microsoft (MSAL.js) pidiendo los scopes `User.Read` y `Sites.ReadWrite.All`.
2. **Resolución del sitio/lista** (`resolverSP()`): traduce la URL del sitio (`CONFIG.SP_SITE_URL`) a un `siteId` de Graph, y busca el `listId` de la lista por su nombre (`CONFIG.LIST_NAME`). Se cachea en memoria (`spSiteId` / `spListId`) para no repetir la búsqueda.
3. **Carga de datos** (`cargarDatos()`): trae **todos** los ítems de la lista vía `GET /sites/{id}/lists/{id}/items?expand=fields&$top=5000`, paginando con `@odata.nextLink` si hay más de 5000 registros.
4. **Deduplicación**: cada registro se identifica por su `_spId` (ID interno de SharePoint, siempre único). Si por algún motivo faltara, se usa como respaldo la combinación `granja_nombre + consecutivo + orden` como clave — esta combinación es, en teoría, única por cerdo.
5. Una vez cargado, `allData` (variable global) contiene todos los registros; `filteredData` es el subconjunto después de aplicar los filtros de la pestaña activa.

> **Nota:** este dashboard **no escribe** en SharePoint — es de solo lectura. La escritura la hace exclusivamente la app de captura.

---

## Control de acceso por asociado

El dashboard sirve tanto a Cercafe (que ve todas las granjas) como a cada asociado individual (que solo debe ver las suyas). Esto se resuelve con un parámetro en la URL:

```
https://.../dashboard-pulmones/?site=GAM
```

```javascript
const SITE_ASOCIADO = {
  'HBM': 'HBM', 'GAM': 'GAM', 'CER': 'CER',
  'CAMPEON': 'CAMPEON', 'AGROJABAR': 'AGRO JABAR', 'CDO': 'CERDOS DEL OTUN',
};
```

- Si la URL trae `?site=GAM`, se guarda en `sessionStorage` (persiste mientras dure la pestaña/sesión del navegador) y se asigna `userAsociado = 'GAM'`.
- Con `userAsociado` seteado: el filtro de "Asociado" se oculta, y todos los datos (`getCompFilteredData()`, `poblarFiltros()`, etc.) se restringen automáticamente a las granjas de ese asociado (vía el mapa `GRANJA_ASOCIADO`).
- Sin el parámetro `?site=` (o con un valor no reconocido): `userAsociado = null` → modo administrador, ve todas las granjas y puede elegir el asociado desde el filtro.
- El botón **"Abrir Tu.HUB"** (portal interno del asociado) también depende de este parámetro, usando el mapa `TUHUB_URL`.

**Para dar acceso a un nuevo asociado:** agregar su clave a `SITE_ASOCIADO` y, si aplica, su URL de Tu.HUB en `TUHUB_URL`, y compartirle el link con `?site=SUCLAVE`.

---

## Estructura de la interfaz

### Pestaña 1 — "Análisis por granja"
Vista de una sola granja/lote a la vez (filtrada por Asociado → Granja → Consecutivo → rango de fechas):
- Tarjetas KPI: IDN, % Consolidación, Frecuencia neumonía, Pérdidas GDP, Frecuencia pleuritis (Leve/Severa), Poliserositis — cada una con su referencia de red Cercafe.
- Sección **"Análisis de resultados"**: tabla interpretativa + cuadro de recomendación automática según la escala del IDN.
- Gráficas: evolución IDN por lote (cada punto = un consecutivo) e IDN promedio mensual.
- Tarjetas de "Otras lesiones" (neumonía intersticial, absceso, nódulo, pericarditis, manchas de leche) y "Artefactos en pulmón sano" (agua, sangre, petequias, ambos — medidos solo sobre cerdos sin ninguna lesión).
- Botón **Exportar PDF** (`exportarPDF()` → `imprimirPDF()`).

### Pestaña 2 — "Comparativo entre granjas"
Compara varias granjas a la vez:
- Checklist de granjas a incluir + selector de indicadores (agrupados en "Indicadores principales", "Otros hallazgos", "Artefactos").
- Gráficas de barras, una por indicador seleccionado, con orden Ascendente / Descendente / Por granja (`_ordenComparativo`), y línea de referencia Cercafe. La gráfica de IDN, además, colorea cada barra según la escala de severidad (`colorEscalaIdn`).
- **Distribución de severidad**: una dona (gráfica de torta) por cada granja seleccionada.
- **Resultados por tabla**: sub-pestañas "Por granja" / "Por consecutivo", con botón **Exportar a Excel** (siempre exporta los 15 indicadores completos, sin importar cuáles estén marcados arriba, respetando las granjas filtradas).
- Botón **Exportar PDF** (`exportarPDFComparativo()` → `imprimirPDFComparativo()`).

---

## Metodología y cálculos

### El dato base: `consolidacion_total` y `categoria`
Cada registro (un cerdo inspeccionado) trae, calculado por la app de captura a partir del % de consolidación de cada uno de los 6 lóbulos pulmonares (ponderados según el método Madec):

- **`consolidacion_total`**: % de consolidación pulmonar de ese cerdo (0–100%, continuo).
- **`categoria`**: ese mismo valor agrupado en bandas de 10 puntos:

  | Consolidación | Categoría |
  |---|---|
  | = 0% | 0 |
  | 0%–10% | 1 |
  | 10%–20% | 2 |
  | 20%–30% | 3 |
  | 30%–40% | 4 |
  | 40%–50% | 5 |
  | > 50% | 6 |

### IDN (Índice de Neumonía)
```
IDN = CT / TCE
CT  = suma de "categoria" de TODOS los cerdos del grupo (granja, lote, mes, etc.)
TCE = TOTAL de cerdos inspeccionados en ese mismo grupo
```
Es decir, el IDN es un **promedio ponderado directo sobre el total de cerdos**, no un promedio de promedios por lote. Esto se calcula en `calcIDN()`, `computeStatsByFarm()`, `computeStatsByLote()` y `dibujarMensual()` — las cuatro funciones usan la misma lógica de suma-total/conteo-total para que el resultado sea consistente sin importar cuántos lotes distintos entren en el filtro aplicado.

**Escala de interpretación del IDN** (a nivel de grupo/promedio):
`Sin lesión = 0 · Leve <0.56 · Moderado 0.56–0.89 · Severo >0.9`

### % Consolidación
```
% Consolidación = Σ(consolidacion_total de cada cerdo) / número de cerdos
```
Promedio simple y directo — igual metodología que el IDN.

### Distribución de severidad (la "dona")
Clasifica a **cada cerdo individualmente** según su `consolidacion_total` crudo (no la categoría 0–6):

| Consolidación | Bucket |
|---|---|
| = 0% | Sin |
| 0%–10% | Leve |
| 10%–30% | Moderada |
| > 30% | Severa |

> ⚠️ **Punto de confusión frecuente:** el IDN (promedio poblacional 0–6) y la dona de severidad (clasificación individual con 4 bandas) parten del mismo dato pero cuentan historias distintas. Un lote puede salir "Severo" en el IDN aun con 0% de cerdos individualmente "Severa" en la dona — basta con que la mayoría tenga algo de lesión leve/moderada para que el promedio suba por encima de 0.9. **No es un error de datos**; son dos métricas complementarias, no la misma escala aplicada dos veces. Por este motivo se agregó una leyenda explícita bajo cada título (`Escala IDN: ...` / `Escala severidad: ...`) para dejarlo claro en la interfaz.

### Otros indicadores (15 en total, ver `INDICADORES_COMP`)
| Grupo | Indicadores |
|---|---|
| Principales | IDN, % Consolidación, % Neumonía, % Pleuritis, % Poliserositis, Pérdidas GDP (kg) |
| Otros hallazgos | Neumonía intersticial, Absceso pulmón, Nódulo pulmón, Pericarditis, Manchas de leche |
| Artefactos | Artefacto agua, Artefacto sangre, Petequias, Ambos artefactos |

- **Poliserositis** = pleuritis (leve o severa) **Y** pericarditis presentes en el mismo cerdo.
- Los **artefactos** (agua, sangre, petequias) solo se calculan sobre el subconjunto de cerdos **sin ninguna lesión** (pulmón sano) — de ahí el subtítulo "Medición sobre cerdos sin ninguna lesión" en esa sección.
- **Pérdidas GDP**: suma de `perdida_gdp_kg` (pérdida de ganancia diaria de peso estimada, en kg) del grupo — no es un promedio, es un total acumulado.

---

## Exportación a PDF

Se genera con `window.print()` del navegador (el usuario elige "Guardar como PDF" en el diálogo de impresión) — no hay generación de PDF en servidor.

- El nombre sugerido del archivo se controla temporalmente sobreescribiendo `document.title` justo antes de imprimir, y se restaura con el evento `afterprint` (no con un `setTimeout` fijo, para no cortar el contenido si el usuario se demora eligiendo dónde guardar).
- Las gráficas de barras y las donas se regeneran en **canvases temporales fuera de pantalla** con un ancho fijo (`generarBarraTemp()`, `generarDonaTemp()`), para que la imagen resultante tenga siempre la misma proporción sin importar el tamaño que tuvieran en pantalla al momento de exportar.
- **Todo el layout de las plantillas de impresión usa `display:grid`, nunca `display:flex` + `flex-wrap`** — ver la nota en [Decisiones de diseño](#decisiones-de-diseño-y-notas-para-mantenimiento).

---

## Exportación a Excel

Botón "Exportar a Excel" en "Resultados por tabla" (`exportarExcelResultados()`), usando SheetJS (`XLSX.utils.json_to_sheet` + `XLSX.writeFile`):
- Detecta automáticamente la sub-pestaña activa ("Por granja" / "Por consecutivo") y exporta ese nivel de agregación.
- Exporta **siempre los 15 indicadores completos**, sin importar cuáles estén marcados en los chips de la sección de gráficas (esos chips solo controlan qué se grafica, no qué se exporta).
- Sí respeta las granjas marcadas en el checklist de filtro — exporta exactamente ese subconjunto.

---

## Campos esperados en la lista de SharePoint

La lista `InspeccionPulmones` debe tener (entre otros) estos campos, uno por cada cerdo inspeccionado:

| Campo | Tipo | Uso |
|---|---|---|
| `granja_nombre` | Texto | Nombre interno de la granja (mayúsculas) |
| `granja_id` | Número | ID numérico de la granja |
| `consecutivo` | Texto | Número de lote |
| `orden` | Número | Número de cerdo dentro del lote (clave de deduplicación) |
| `fecha_inspeccion` | Fecha | Fecha de la inspección |
| `consolidacion_total` | Número | % de consolidación pulmonar del cerdo (0–100) |
| `categoria` | Número | Categoría Madec (0–6), derivada de `consolidacion_total` |
| `pleuritis_leve`, `pleuritis_severa` | Sí/No | Hallazgos |
| `pericarditis`, `neumonia_intersticial`, `absceso_pulmon`, `nodulo_pulmon`, `manchas_leche` | Sí/No | Otras lesiones |
| `artefacto_agua`, `artefacto_sangre`, `petequias` | Sí/No | Artefactos (solo relevantes si el cerdo no tiene ninguna lesión) |
| `perdida_gdp_kg` | Número | Pérdida de ganancia diaria de peso estimada, en kg |

El mapeo `granja_nombre` (mayúsculas, como se guarda) → nombre de presentación se hace con `GRANJA_VISTA`, y granja → asociado con `GRANJA_ASOCIADO`. **Si se agrega una granja nueva a la operación, hay que agregarla en `GRANJAS`, `GRANJA_VISTA` y `GRANJA_ASOCIADO`** dentro del script — no hay una fuente de verdad externa para esto todavía.

---

## Despliegue

1. El archivo `index.html` se sirve tal cual (por ejemplo, GitHub Pages, repo `cercafe-kpis/dashboard-pulmones`).
2. Registrar la URL final exacta como *Redirect URI* (tipo SPA) en la App Registration de Azure AD (`CONFIG.CLIENT_ID`).
3. Confirmar que la cuenta que inicia sesión tenga permisos de lectura sobre el sitio/lista de SharePoint configurados en `CONFIG`.
4. No requiere variables de entorno ni build — subir el archivo y listo.

---

## Decisiones de diseño y notas para mantenimiento

- **CSS Grid, no Flexbox, en las plantillas de impresión.** Chrome tiene un bug conocido: los contenedores `display:flex` con `flex-wrap` se rompen al paginar un PDF (una tarjeta puede estirarse y ocupar una página entera en blanco). `display:grid` con `grid-template-columns` sí pagina correctamente. Si se agrega una sección nueva al PDF, usar grid.
- **IDN calculado sobre el total, no por promedio de lotes.** Ver la sección de [Metodología](#idn-índice-de-neumonía) — este fue un bug corregido: antes se promediaba el IDN de cada lote por separado (dándole igual peso a un lote de 5 cerdos que a uno de 500). Ahora se suma el total de categorías y se divide entre el total de cerdos, de forma consistente en las 4 funciones que lo calculan.
- **La dona de severidad y el IDN usan escalas distintas a propósito** — no unificar los colores/rangos sin agregar antes una leyenda clara, para no repetir la confusión que motivó agregar las leyendas "Escala IDN" / "Escala severidad" bajo cada título.
- **Deduplicación por `_spId` primero, por `granja+consecutivo+orden` como respaldo.** Si en algún momento se detectan registros duplicados con `_spId` distintos, es un problema de la app de captura (reintentos de sincronización o recaptura de un lote ya guardado), no de este dashboard — revisar `sincronizarPendientes()` e `iniciarSesion()` en `index_app.html`.
- **`ns` en `drawBarChartComparativo` / `ordenarIndicador`**: quedó un parámetro de conteo de cerdos por granja que se probó mostrar sobre las barras (opción descartada visualmente) y luego se dejó de dibujar, pero el parámetro se sigue pasando por la cadena de llamadas. Es inofensivo (no se usa), se puede limpiar en una refactorización futura si se desea.
- **`resolverSP()` toma la primera lista que coincida por nombre** (`ld.value[0]`). Si alguna vez existieran dos listas con el mismo nombre en el sitio, tomaría la que la API devuelva primero (no garantizado). No ha sido un problema hasta ahora, pero vale la pena tenerlo presente.

---

## Mapa de funciones clave

| Función | Rol |
|---|---|
| `doLogin()` / `getToken()` / `resolverSP()` | Autenticación MSAL y resolución del sitio/lista de SharePoint |
| `cargarDatos()` | Trae todos los registros de la lista (con paginación) y deduplica |
| `poblarFiltros()` / `filtrarPorAsociado()` / `aplicarFiltros()` | Manejo de los filtros de la pestaña "Análisis por granja" |
| `calcIDN(data)` | IDN pooled (CT/TCE) sobre un conjunto de datos |
| `actualizarKPIs()` | Rellena las tarjetas KPI de "Análisis por granja" |
| `generarAnalisisDashboard(...)` | Genera el HTML de "Análisis de resultados" + recomendación |
| `dibujarLinea()` / `dibujarMensual()` | Gráficas de evolución IDN por lote / por mes |
| `dibujarComparativo()` | Orquesta donas + tabla + gráficas del tab "Comparativo" |
| `dibujarDonasPorGranja()` | Una dona de severidad por cada granja seleccionada |
| `computeStatsByFarm(data)` / `computeStatsByLote(data)` | Agregación de los 15 indicadores por granja / por lote |
| `renderComparativo()` / `drawBarChartComparativo()` | Gráficas de barras comparativas (canvas nativo) |
| `actualizarTablas(data)` | Tablas "Por granja" / "Por consecutivo" |
| `exportarExcelResultados()` / `descargarExcel()` | Exportación a `.xlsx` (SheetJS) |
| `exportarPDF()` / `imprimirPDF()` | PDF de "Análisis por granja" |
| `exportarPDFComparativo()` / `imprimirPDFComparativo()` | PDF de "Comparativo entre granjas" |
| `generarBarraTemp()` / `generarDonaTemp()` | Renderizan gráficas en canvas temporal con ancho fijo, para las imágenes del PDF |

---

*Este README fue generado a partir de una revisión directa del código fuente de `index.html`. Si el código cambia, actualizar este documento en consecuencia.*
