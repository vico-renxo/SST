# CLAUDE.md — Sistema SST (SSOMA) · Guía Completa para Claude Code

> Archivo autocontenido. Cualquier instancia de Claude Code que lo lea puede trabajar
> en el proyecto sin preguntas básicas de contexto.
> **Regla:** actualizar la sección correspondiente después de cada cambio que modifique
> columnas de Sheets, funciones públicas o patrones nuevos.

---

## 1. IDENTIDAD DEL PROYECTO

**Nombre:** Sistema de Gestión SSOMA — Adecco Perú
**Propósito:** Plataforma web para Seguridad, Salud Ocupacional y Medio Ambiente:
inspecciones, EPP, capacitaciones, desvíos, mapa de riesgos, rol de turnos,
eventos/accidentes e IPERC.

**Stack tecnológico detectado en código:**
- Backend: Google Apps Script (GAS) — archivos .js del proyecto
- Frontend: HTML/CSS/JS servido por GAS doGet() → HtmlService
- Base de datos: Google Sheets (múltiples spreadsheets independientes)
- Almacenamiento: Google Drive
- Push notifications: Cloudflare Worker (viczul.com) + Web Push API
- Mensajería: Telegram Bot API
- IA: Google Gemini 2.5 Flash (gemini-2.5-flash)
- UI libs: Bootstrap 5.3, SweetAlert2, DataTables, SignaturePad 4.1.7, Chart.js

**URL producción GAS:** https://script.google.com/macros/s/AKfycbwJrer0KO6jEd9HFso-AKzARyzlVdRrblJzm1H2i2ylWCbsCS9XzLGAfuQio2EPMzg/exec
**Cloudflare Worker:** https://viczul.com

---

### 1.1 Spreadsheets (IDs reales del código)

Todos los IDs están centralizados en `SPREADSHEET_IDS` de `Code.js` (clave camelCase). Usar siempre esa constante, no hardcodear.

| Alias / clave SPREADSHEET_IDS | Descripción | ID |
|---|---|---|
| `personal` | Gestión personal / login | 1NDDHlTfWxmObgm8JZu5WAnCECB3gU6e_k7o_sFcMrkw |
| CHECK | Inspecciones checklist (main) | 12KkPwl_gfQCkqS9ZHsp4hS2fFkebgNbszvTDtZELObU |
| CHECK_V2 | Inspecciones checklist (test) | 1NR4VtBUqO6DkM_rSjNqC8m19-QPjrd_IW1aEmmsUD6U |
| DESVIOS | Desvíos y observaciones | 1eIJfA7dAlkQ1rXcRGC2qSFnvZ-jYIPn8cA_TbUZcWZE |
| GRAFICOS | KPIs y gráficos | 1J_v47ohrGj8S1XfWUdneH0l7mMTB8auSOEscHZwsM0g |
| EPP | Equipos de protección personal | 1Mxy5SkDdLy1Ihct844uLq5ZALe-RFarDfWo9j65kBcE |
| CAPACITACIONES | Capacitaciones y charlas | 1Ev5_B3jMtjy_xXt13NYBXYwFA-maFAeLSKfiCFIsMQo |
| ROL_EMPLEADOS | Rol turnos (hoja empleados SS) | 1SrkbAD8aoLGCCr8oMh0yRp3iiRl0Du4WEpUU88zOCOc |
| ROL_ALERTAS | Rol turnos (BD alertas) | 12h2yVs0NlD3h3zMYl_93o7ohOKzurxcPZXifoTyVigE |
| EVENTOS | Accidentes y eventos | 1Xo5HgaHfskg_mkguGTuuR_AcKeQUoRoG--ch-V8KTpw |
| MAPA_RIESGOS | Mapa de riesgos | 1EfQvY59m1l1SB_GD__CzL-qJQFdbtYzM9Y2q1u2L3cI |
| IPERC | Matriz IPERC (override via ScriptProperties IPERC_SS_ID) | 1ANw0WcZiDYfZhDTmEazUeBtx7_MyhG9UOnVTJNx-Xjo |

### 1.2 Carpetas Drive (IDs reales del código)

| Constante en código | Uso | ID |
|---|---|---|
| folderIdFirma | Firmas perfil usuario | 1TzV9UlPupxeRyo7l2Vn2nO9mh64WG_Kv |
| FOLDER_IDEPP | Firmas EPP | 1rzdA2M0abUuT83i9YFF7-W8eLHWrs84s |
| foldercharlas | Archivos de charlas | 1IWmNW4wMZbC43QHrdivlRtAwnTHa3c2v |
| foldefirmascap | Firmas capacitaciones | 1GcIoeFFtpZ6EISt0w5R1byNkDy5I3ugi |
| folderimgcheck | Imágenes inspecciones main | 13qGGx2VJRlbcPSw9b2Ldn10HlJQJg0zd |
| folderpdfcheck | PDFs inspecciones main | 1Be7s5TlJRS6sj0f6NxhGPK9EqFJcVM6N |
| folderimgcheck_v2 | Imágenes inspecciones v2 | 1AnA7M-M7NXVuduuL9ispiLS5mXYvkrb8 |
| folderpdfcheck_v2 | PDFs inspecciones v2 | 15Ofbz86NUhZRuFlWjLUlR6kg1YAY7SXo |
| FOLDER_DB_ID | JSON rol_turnos.json | 17tKcRGZtUjE0HwosxlGrycFWIJ20aaS8 |
| COMUNICADOS_FOLDER_ID | Imágenes comunicados | 1_a0rg1PK13NtkDLQ-y-tqKslo9aOc8DV |

### 1.3 Estructura de carpetas del repositorio

```
/home/user/SST/
├── Code.js              # GLOBAL: login, usuarios, SPREADSHEET_IDS, Gemini helper, Telegram
├── Graficos.js          # KPIs, gráficos, pronóstico IA, caché Map en memoria
├── CheckCode.js         # Inspecciones checklist (main) — hojas CHECK SS
├── TestCode.js          # Inspecciones checklist v2 (alternativo) — hojas CHECK_V2 SS
├── EppCode.js           # EPP: stock, movimientos, registro, firma trabajador
├── CapaciCode.js        # Capacitaciones, charlas, evaluaciones
├── DesvioscCode.js      # Desvíos y observaciones de seguridad
├── EventosCode.js       # Eventos, accidentes, mapa de coordenadas
├── IpercCode.js         # Matriz IPERC + análisis con Gemini
├── PassoCode.js         # PASSO: inspecciones + reuniones de seguridad
├── RolCode.js           # Rol de turnos (JSON en Drive)
├── HhtCode.js           # HHT (Horas Hombre Trabajadas)
├── NotificacionesCode.js# Push notifications via Cloudflare Worker
├── Telegram.js          # Notificaciones Telegram Bot
├── AlertasCode.js       # Alertas vencimientos EPP (usa EppCode)
├── CodeMapa.js          # Mapa de Riesgos CRUD
├── ComunicadosCode.js   # Comunicados internos (usa PERSONAL SS)
├── Homecode.js          # Avisos ERP (hoja AVISOS en PERSONAL SS)
├── index.html           # SPA: login + router de módulos + tema Neo Brutalism toggle
├── home.html            # Dashboard post-login: avisos ERP + marcador de asistencia (iframe externo)
├── css.html             # Estilos GLOBALES: Bootstrap 5.3, sidebar, navbar, tema Neo Brutalism, panel IA
├── css-modulos.html     # Estilos ESPECÍFICOS de cada módulo (home, Check, Rol, EPP, etc.) — cargado en index.html tras css.html
│
│   ── CHECKLIST ──────────────────────────────────────────────────────────────
├── Check.html           # Formulario de inspección checklist (main v1)
├── Test.html            # Formulario de inspección checklist (v2/test con IA)
├── IndexCheck.html      # Vista de tarjetas de registros + PDF Masivo
├── ActualCheck.html     # Estado actual de checklist (resumen rápido)
├── EditCheck.html       # Edición de registro de checklist existente
├── ListaCheck.html      # Lista tabular de checklists (sin tarjetas)
├── TestCheck.html       # Variante de Test en tabla
├── CheckTest.html       # Vista alternativa de Test
├── InventarioCheck.html # Gestión del inventario de equipos (main)
├── InventarioTest.html  # Gestión del inventario de equipos (v2)
│
│   ── EPP ────────────────────────────────────────────────────────────────────
├── MovimEpp.html        # UI movimientos EPP (ingresos/salidas de almacén)
├── Asignaciones.html    # UI firma EPP por trabajador (confirmar entrega)
├── MatrizEpp.html       # Configuración de reglas EPP por cargo/producto
├── EPPMaestro.html      # Gestión maestra de productos y almacenes
│
│   ── CAPACITACIONES ─────────────────────────────────────────────────────────
├── Capacitaciones.html  # Gestión de capacitaciones y charlas (módulo principal)
├── RegistrosCap.html    # Registros de capacitación con PDF descargable (SSOMA-FR006)
├── BuscadorCap.html     # Cumplimiento por trabajador (una fila/trabajador + historial)
├── BuscadorCharlas.html # Buscador de charlas registradas
├── Cursos.html          # Catálogo de cursos/temas de capacitación
├── Evaluacion.html      # Módulo de evaluación/examen online
├── Examen.html          # Formulario de examen individual
├── Matriz.html          # Vista de la matriz de capacitación por cargo
├── ReportesLaboral.html # Reportes laborales de capacitación (PDF/Excel)
│
│   ── OTROS MÓDULOS ───────────────────────────────────────────────────────────
├── IndexDesvios.html    # UI desvíos y observaciones
├── MapaRiesgos.html     # UI mapa de riesgos (con mapa Leaflet/Google)
├── Inspecciones.html    # UI inspecciones PASSO
├── PASSO.html           # UI programa PASSO (inspecciones + reuniones)
├── Usuarios.html        # UI administración usuarios (solo admin)
├── Rol.html             # UI rol de turnos (planner semanal)
├── Graficosindex.html   # UI KPIs y gráficos (Chart.js + Gemini pronóstico)
├── Eventos.html         # UI eventos y accidentes
├── Iperc.html           # UI matriz IPERC
├── Hht.html             # UI Horas Hombre Trabajadas
├── Listas.html          # Gestión de listas maestras (dropdowns del sistema)
├── Comunicados.html     # Comunicados internos con imagen
└── Pagina web/
    ├── cloudflare-worker.js  # Worker push notifications
    └── worker.js
```
---

## 1.4 Script Properties requeridas (GAS → Configuración → Propiedades de script)

| Clave | Usado por | Descripción |
|---|---|---|
| `GEMINI_KEY` | Code.js `API_KEY` | API Key de Google AI Studio |
| `TELEGRAM_BOT_TOKEN` | Telegram.js `TELEGRAM_CONFIG.botToken` | Token del bot de Telegram |
| `TELEGRAM_CHAT_ID` | Telegram.js `TELEGRAM_CONFIG.chatId` | ID del chat/grupo de Telegram |
| `PUSH_AUTH_TOKEN` | NotificacionesCode.js `PUSH_AUTH_TOKEN` | Token de autenticación Cloudflare Worker |
| `ADMIN_EMAIL` | AlertasCode.js | Email del administrador (sin fallback — ver L8) |
| `IPERC_SS_ID` | IpercCode.js | Spreadsheet IPERC (override del default) |
| `IPERC_CONFIG_EMPRESA` | IpercCode.js | Config empresa IPERC |
| `IPERC_GEMINI_MODELO` | IpercCode.js | Modelo Gemini para análisis IPERC |
| `DB_FILE_ID_V4` | RolCode.js | Cache del fileId del JSON de turnos en Drive |

**IMPORTANTE:** Mientras las Script Properties no estén configuradas, el sistema usa los valores hardcodeados como fallback. Una vez configuradas las propiedades, los valores hardcodeados se vuelven letra muerta pero no rompen nada.

---

## 2. ARQUITECTURA Y MÓDULOS

### Code.js — Núcleo global
**Responsabilidad:** Login, gestión de usuarios, constantes globales, helper Gemini, helper Telegram.
**Spreadsheet:** PERSONAL (1NDDHlTfWxmObgm8JZu5WAnCECB3gU6e_k7o_sFcMrkw)
**Constantes globales (disponibles en TODO el proyecto):**
- `SPREADSHEET_IDS` — objeto con alias check/desvios/graficos
- `API_KEY` — Gemini API Key (HARDCODEADA — ver sección 9)
- `geminiUrl` — endpoint Gemini 2.5 Flash
- `folderIdFirma` — Drive folder para firmas de perfil

**Funciones públicas:**
- `doGet(e)` — router HTTP: retorna HTML o JSON (asistente voz)
- `include(filename)` — incluye partials HTML
- `loginData(obj)` — verifica credenciales, retorna datos de sesión
- `actualizarUsuariologin(datos)` — actualiza perfil + sube firma a Drive
- `obtenerUsuariosPaginado(offset, limit, filtro)` — lista paginada usuarios
- `agregarUsuario(data)` — crea usuario + email bienvenida + Telegram
- `actualizarUsuario(data)` — actualiza usuario + Telegram
- `actualizarCondicionUsuario(id, condicion)` — actualiza col L (condición laboral); valida contra CONDICIONES_VALIDAS
- `eliminarUsuarioPorUsuario(usuario)` — elimina + Telegram
- `buscarDatosPorNumero(numero)` — busca por DNI
- `getTodasLasListas()` — listas maestras con caché 5 min (CacheService, key: listas_globales_v5)
- `getRecordsList()` / `saveRecordsList(records)` — hoja LISTAS
- `getColor()` / `saveColor(color)` — color personalizado usuario (celda J1 RESUMEN)
- `enviarTelegram(mensaje)` — envía mensaje a Telegram
- `invalidarTodasLasCaches()` — invalida caché backend completa: CacheService keys `listas_globales_v5`, `TEMAS_CACHE`, `lista_temas` + llama `_invalidateDesviosCache()`, `_invalidateStockCache_()`, `limpiarCache()`. Retorna `JSON.stringify({ok,ts})`. Llamada desde `forzarRefrescoListas()` (index.html) para sincronizar junto con la limpieza de localStorage frontend.
- `getIncompatibilidadData(...)` — incompatibilidades paginadas
- `agregarIncompatibilidad(data)` / `actualizarIncompatibilidad(id, data)` / `eliminarIncompatibilidad(id)`

**Funciones privadas:**
- `_gasAsistenteHandler(e)` — router de asistente voz
- `_normText(s)` — normaliza texto (minúsculas, sin acentos)
- `_callGemini(prompt, generationConfig)` — llamada a Gemini API (ÚNICO punto de llamada, reutilizar siempre)

**Dependencias:** Telegram.js (notificarNuevoUsuario, etc.), getSpreadsheetCapacitaciones

---

### CheckCode.js — Inspecciones Checklist
**Responsabilidad:** CRUD de inspecciones, PDF desde HTML, filtro de inventario por cargo.
**Spreadsheet:** CHECK (12KkPwl_gfQCkqS9ZHsp4hS2fFkebgNbszvTDtZELObU)
**Hojas usadas:** B DATOS, CHECK LIST, INVENTARIO, HISTORIAL, ACTUAL, MENÚ, Acceso, FORMATO

**Funciones públicas:**
- `getDropDownarray(cargo)` — inventario filtrado por cargo + disponibilidad por período
- `searchData(obj)` — busca ítems del checklist por equipo
- `saveDataCheck(obj)` — guarda inspección (cols A-O + fotos BQ-CT + firma CU + firma supervisor CV)
- `getDatosRegistroCheck(offset, limit, filtroMes, filtroEquipo)` — historial paginado
- `generarPDFdesdeHTML(recordId)` — genera PDF e sube a Drive
- `getPdfUrl(columnAValue)` — alias de generarPDFdesdeHTML
- `deleteData(ID)` — elimina fila + fotos Drive
- `getEquiposCheck()` — lista de equipos únicos
- `getHeadersCheck()` — cabeceras B DATOS
- `setStatusCheck()` — contadores celdas T1/U1
- `getResponsablesPersonal()` — lista personal activo para selector ad6
- `getSupervisoresOperaciones()` — supervisores activos para adSup
- `saveDataCheckSeguimiento(objData)` — guarda fotos de subsanación por ítem
- `setGestionResponsable(objData)` — marca ítem como EN_GESTION
- `obtenerDatosChecklist()` — ítems CHECK LIST
- `agregarChecklist(data)` / `eliminarChecklist(rowIndex)` / `actualizarChecklistPorEquipo(...)`
- `obtenerInventarioServerSide(...)` / `agregarInventario(data)` / `actualizarInventario(data)` / `eliminarInventarioPorNum(num)`
- `generarItemsConGemini(base64DataUrl, textoBase, numItems)` — genera ítems con IA
- `generarPDFsMasivosCheck(filtroMes, filtroEquipo)` — genera PDFs de todos los registros filtrados por mes (1-12 o "Todos") y equipo, los copia a una subcarpeta de `folderpdfcheck`, la comparte como pública y retorna `JSON.stringify({url, total, exitosos, fallidos})`. Límite: 60 registros por lote. Usada por el botón "PDF Masivo" en IndexCheck.html.

**Funciones privadas:**
- `_parseFechaCheck(txt)` — parsea fecha texto a timestamp
- `_estaDisponibleInsp(freq, ultimaDate, hoyDate)` — disponibilidad por período
- `_freqEsCalendario(freq)` / `_mismoPeriodoCalInsp(...)` / `_inicioPeriodoActualInsp(...)`
- `_diasVencidoInsp(...)` / `_calcularCumplimientoInsp(...)`
- `parseObsCell(cellValue)` — parsea formato num::url~~num::url
- `agregarFotosSeccion(obj, fila, folder)` — sube fotos por sección a Drive
- `calcularEstadoCheck(checkId, sheet, rowIndex)` — estado automático
- `convertirUrlParaPDF(url)` — URL Drive → data-URI base64

**Dependencias:** getSpreadsheetPersonal() (Code.js), MailApp

---

### EppCode.js — EPP (Equipos de Protección Personal)
**Responsabilidad:** Stock, movimientos, registro de entregas/devoluciones, firmas.
**Spreadsheet:** EPP (1Mxy5SkDdLy1Ihct844uLq5ZALe-RFarDfWo9j65kBcE)
**Hojas:** STOCK, MOVIMIENTOS, REGISTRO, MATRIZ, ALMACENES

**Constantes de columnas (IDX — 1-based):**
- IDX.STOCK: ID(1) ALMACEN(2) PRODUCTO(3) VARIANTE(4) CATEGORIA(5) STOCK(6) STOCK_MINIMO(7) PRECIO(8) IMG(9) FILE(10)
- IDX.MOV: ID_MOV(1) FECHA(2) OPERACION(3) ALMACEN(4) PRODUCTO(5) VARIANTE(6) CANTIDAD(7) ID_PROV(8) MARCA(9) COSTO_UNITARIO(10) MONEDA(11) IMPORTE(12) USUARIO(13) OBS(14) DNI(15) CARGO(16) FIRMA_URL(17) ESTADO(18) FECHA_CONFIRMACION(19)
- IDX.REG: ID_REG(1) FECHA(2) OPERACION(3) ALMACEN(4) PRODUCTO(5) VARIANTE(6) DNI(7) NOMBRES(8) EMPRESA(9) CARGO(10) CANTIDAD(11) COSTO_UNITARIO(12) MONEDA(13) IMPORTE(14) USUARIO(15) OBS(16) DEVOLVIBLE(17) VIDA_UTIL_DIAS(18) FREC_INSP(19) REQ_CAP_TEMA(20) FECHA_VENCIMIENTO(21) PROX_INSPECCION(22) FIRMA_URL(23) REF_ID(24) ESTADO(25) FECHA_CONFIRMACION(26)

**Funciones públicas:**
- `getInit(almacenId)` — catalogo inicial (almacenes, productos, categorías, matriz)
- `searchStockPaged(payload)` — stock paginado server-side
- `searchHistorialPaged(payload)` — historial REGISTRO o MOV paginado
- `getVariantesPorProducto(almacenId, producto)` — variantes disponibles
- `getPersonaByDni(dni)` — datos persona desde PERSONAL
- `buscarPersonaFlexible(query)` — búsqueda por DNI o nombre parcial
- `getReglasCargo(productoBase, cargo)` — reglas MATRIZ por producto/cargo
- `getHistorialByDni(dni, limit)` — historial EPP de un trabajador
- `confirmarEntregaEPP(regId, firmaBase64, dni)` — trabajador confirma con firma
- `obtenerEntregasPendientes(dniLogin)` — entregas pendientes de confirmación
- `generarRegistroEPP(dni)` — genera PDF "Registro de Entrega de EPP" (RE-SSOMA-011), sube a FOLDER_IDEPP y retorna URL pública

**Funciones privadas:**
- `_sh(name)` — atajo getSpreadsheetEPP().getSheetByName(name)
- `_str(x)` / `_num(x)` / `_today()` / `_fmtDateOut(v)` — utils de tipo
- `_readRows(name)` — lee hoja completa sin cabecera
- `_appendArray(name, arr)` — append fila con preservación formato DNI
- `_readStockCache_()` / `_readRegistroCache_()` / `_readMovCache_()` — caché CacheService 120s
- `_invalidateStockCache_()` — invalida claves stock:all:v1, registro:all:v1, mov:all:v1
- `_buildEmailToNameMap()` — mapa email→nombre desde PERSONAL
- `_resolveUsuarioNombre()` / `_resolveUsuarioDni()` — usuario GAS activo

**Dependencias:** getSpreadsheetPersonal() (Code.js), NotificacionesCode.js

---

### CapaciCode.js — Capacitaciones
**Responsabilidad:** Gestión de capacitaciones, charlas, evaluaciones y firmas.
**Spreadsheet:** CAPACITACIONES (1Ev5_B3jMtjy_xXt13NYBXYwFA-maFAeLSKfiCFIsMQo)
**Hojas:** Matriz, B DATOS, LIST
**Carpetas Drive:** foldercharlas (charlas), foldefirmascap (firmas cap)
**Dependencias:** getSpreadsheetPersonal() (Code.js)
**Funciones públicas relevantes:**
- `obtenerDatosPorDNI(dni)` — cursos asignados al trabajador según matriz (estado actual)
- `obtenerDashboardLaboral()` — KPIs + `cumplimientoPorTrabajador` (todos activos) + incumplidores + vacunas
- `getHistorialCapacitacionesTrabajador(dni)` — todos los intentos de evaluación agrupados por tema, newest-first. Retorna `JSON.stringify({historial:[{tema, intentos:[{puntaje,fecha,estado}]}]})`. Excluye filas ACTIVACION. Usada por el botón "Acciones" en BuscadorCap.html.
- `getCumplimientoPorTrabajador(search, fechaDesde, fechaHasta)` — una fila por trabajador activo (DNI, Nombre, Cargo, Empresa, Aprobados, Previstos, %). Filtra B DATOS por rango de fecha de evaluación; calcula aprobados vigentes vs previstos de la Matriz. Retorna `JSON.stringify({headers, data:[{dni,nombre,cargo,empresa,aprobados,previstos,porcentaje}]})`. Usada por BuscadorCap.html.
- `generarRegistroCap(codigo)` — genera PDF "REGISTRO DE INDUCCIÓN, CAPACITACIÓN, ENTRENAMIENTO Y SIMULACROS DE EMERGENCIA" (SSOMA-FR006). Lee TEMAS por código → lee REGISTRO FIRMAS por Código Registro (col 11) → obtiene info empresa → busca firma capacitador en PERSONAL. Mapea área+tema a checkboxes TIPO (actividad/materia). Convierte firmas Drive a base64. Genera HTML → PDF via `Utilities.newBlob().getAs(MimeType.PDF)` → guarda en foldercharlas. Retorna `JSON.stringify({url})`. Usada por RegistrosCap.html botón PDF.

---

### DesvioscCode.js — Desvíos
**Responsabilidad:** CRUD desvíos/observaciones, análisis IA, alertas email.
**Spreadsheet:** DESVIOS (via SPREADSHEET_IDS.desvios)
**Hojas:** B DATOS, INSPECCIÓN, ANALISIS, FICHA RAC T1, MENÚ, Acceso
**Dependencias:** SPREADSHEET_IDS (Code.js), getSpreadsheetPersonal(), GmailApp

---

### EventosCode.js — Eventos / Accidentes
**Responsabilidad:** Registro de eventos/accidentes, mapa de coordenadas, stock.
**Spreadsheet:** EVENTOS (1Xo5HgaHfskg_mkguGTuuR_AcKeQUoRoG--ch-V8KTpw)
**Hojas:** B DATOS, Listas, Stock
**Dependencias:** getSpreadsheetPersonal() (Code.js)
**Funciones públicas:**
- `searchByCoordinates(x, y)` — búsqueda por coordenadas GPS
- `getAllPoints()` — todos los puntos del mapa
- `deleteByIDAccindentes(id)` — eliminar registro por ID
- `getNombreEmpleado(idEnfermo)` — nombre completo desde PERSONAL
- `getParte()` — listas Parte/Atención desde hoja Listas
- `getStockData()` — stock de medicamentos (hoja Stock)
- `salvar(pedido)` — guardar despacho de medicamentos
- `saveFormDataAccidentesV2(...)` — guardar/editar evento accidente (función principal)
- `generateUniqueID()` — genera ID correlativo vía PropertiesService

---

### IpercCode.js — Matriz IPERC
**Responsabilidad:** CRUD matriz IPERC + análisis con Gemini por fila.
**Spreadsheet:** configurable via PropertiesService key IPERC_SS_ID (default: 1ANw0WcZiDYfZhDTmEazUeBtx7_MyhG9UOnVTJNx-Xjo)
**Hojas:** DATOS, Configuracion Matriz
**PropertiesService keys:** IPERC_SS_ID, IPERC_CONFIG_EMPRESA, IPERC_GEMINI_MODELO
**Funciones públicas:**
- `ipercLeerMatrizConfig()` / `ipercLeerMatrizConfigJSON()` — config de la matriz
- `ipercGuardarFilasAlSheet(filasJSON)` / `ipercLeerFilasDelSheet()` — CRUD filas
- `ipercGuardarConfigEmpresa(cfgJSON)` / `ipercLeerConfigEmpresa()` — config empresa
- `ipercAnalizarFilaConIA(filaJSON)` — análisis de riesgo con Gemini

---

### PassoCode.js — PASSO
**Responsabilidad:** Programa de inspecciones + reuniones de seguridad.
**Dependencias:** getCheckSpreadsheet(), getSpreadsheetCapacitaciones(), getSpreadsheetPersonal()
**Funciones públicas:**
- `obtenerDatosPASSOInspecciones()` — cruce INVENTARIO × B DATOS para cumplimiento
- `getPassoDesdeMatriz()` — capacitaciones desde hoja Matriz
- `obtenerReuniones()` / `agregarReunion(data)` / `actualizarReunion(data)` — reuniones (hoja REUNIONES en CAPACITACIONES SS)

---

### RolCode.js — Rol de Turnos
**Responsabilidad:** Gestión de turnos (JSON en Drive), empleados, MOF.
**Spreadsheets:** `SPREADSHEET_IDS.rolEmpleados` (ROL_EMPLEADOS), `SPREADSHEET_IDS.rolAlertas` (ROL_ALERTAS) — leídos desde Code.js, ya no hardcodeados localmente.
**Drive:** FOLDER_DB_ID (17tKcRGZtUjE0HwosxlGrycFWIJ20aaS8) — archivos rol_turnos.json, department_config.json
**PropertiesService:** DB_FILE_ID_V4 — caché del fileId del JSON en Drive
**Funciones públicas:**
- `getEmployeesFromDB()` — lista empleados desde PERSONAL (JSON string)
- `getRolInitialData()` — empleados + MOF combinados
- `savePlannerToDrive(jsonString)` / `loadPlannerFromDrive()` — persistencia JSON
- `saveDepartmentConfig(config)` / `getDepartmentConfig()` — config departamentos
- `saveFullReport(payload)` — guarda reporte semanal en BD_Detalle / BD_Resumen_Semanal
- `getMOFConfigData()` — cargos/áreas desde hoja MOF
- `generateShortName(fullName)` — nombre corto (Inicial. Apellido)

---

### Graficos.js — KPIs y Gráficos
**Responsabilidad:** Datos para dashboards, índices IF/IG/IA, pronóstico con Gemini.
**Spreadsheets:** usa SPREADSHEET_IDS (check/desvios/graficos) + CAPACITACIONES
**Caché:** Map en memoria (cacheGraficos) con TTL 3 min + PronosticoCache en hoja IA (TTL 7 días)
**Funciones públicas:**
- `obtenerDatosNiveles(fechaInicio, fechaFin, empresa)` — índices acumulados
- `obtenerDatosMensuales(fechaFin, empresa)` — series mensuales IF/IG/IA
- `obtenerEmpresasUnicas()` — empresas para filtros
- `obtenerResumenDesviosPorMes(...)` / `obtenerResumenEstadosPorMes(...)` — resúmenes
- `limpiarCache()` / `obtenerEstadoCache()` — gestión caché
- `generarPronosticoConGemini(fechaInicioStr, fechaFinStr)` — pronóstico IA

---

### NotificacionesCode.js — Push Notifications
**Responsabilidad:** Envío de push notifications via Cloudflare Worker.
**Constantes:** PUSH_WORKER_URL = https://viczul.com, PUSH_AUTH_TOKEN (hardcodeado — ver sección 9)
**ROL_SS_ID_ALERTAS:** 12h2yVs0NlD3h3zMYl_93o7ohOKzurxcPZXifoTyVigE
**Funciones públicas:**
- `enviarPushNotification(dni, title, body, tag)` — push a un trabajador
- `enviarPushBulk(dnis, title, body, tag)` — push a varios
- `notificarEntregaEPP(...)` / `notificarConfirmacionEPP(...)` / `notificarCapacitacionPush(dnis, tema, fecha)`
- `notificarATodos(titulo, mensaje)` — push broadcast
- `verificarYEnviarAlertasInspeccion()` — trigger cada 4h (inspecciones vencidas)
- `simularAlertasInspeccion()` — diagnóstico sin enviar
- `testPushConnection()` — test conexión Worker

---

### Telegram.js — Notificaciones Telegram
**Responsabilidad:** Envío de mensajes, documentos e imágenes al chat de Telegram.
**Config:** TELEGRAM_CONFIG.botToken (hardcodeado), TELEGRAM_CONFIG.chatId
**Funciones públicas:**
- `enviarTelegram(mensaje, opciones)` — mensaje con opciones (botones, parseMode)
- `notificarNuevoUsuario(datos)` / `notificarUsuarioActualizado(datos)` / `notificarUsuarioEliminado(usuario)`
- `notificarDesvio(datos)` / `notificarChecklistCompletado(datos)` / `notificarEvento(datos)`
- `notificarEPP(datos)` / `notificarLogin(nombre)` / `notificarMapaRiesgos(datos)`
- `notificarCapacitacionTelegram(datos)` — capacitación programada (renombrada desde `notificarCapacitacion` para evitar colisión con la versión Push de NotificacionesCode.js)
- `enviarResumenDiario()` — trigger automático 8 AM
- `testTelegram()` — test de conexión

---

### AlertasCode.js — Alertas EPP
**Responsabilidad:** Detectar EPP vencidos o próximos a vencer.
**Dependencias:** getSpreadsheetEPP() (EppCode.js), IDX.REG, SHEPP, _readMatrizGrid_() (EppCode.js)
**PropertiesService:** `ADMIN_EMAIL` — email del administrador para acceso total (SIN fallback — si no está configurado, ningún usuario es admin por email)
**Algoritmo (entrega más reciente = estado actual):**
1. Agrupa por clave `dni|producto` (sin variante — una entrega nueva de cualquier variante cancela la alerta de variantes anteriores).
2. Por cada producto: toma la entrega MÁS RECIENTE por `fechaEntrega` (col FECHA).
3. Si esa entrega no tiene `fechaVenc` → cubierto indefinidamente, sin alerta.
4. Si `fechaVenc > hoy` → cubierto, sin alerta.
5. `diffDias <= 0` → VENCIDO (bg-rojo); `0 < diffDias <= 15` → VENCE EN N DÍAS (bg-naranja).
**IMPORTANTE — ADMIN_EMAIL:** usar solo `PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL')`, sin fallback a `Session.getActiveUser().getEmail()`. El fallback haría `esAdmin=true` para todos, bypasseando el filtro por DNI y mostrando EPPs de otros trabajadores.
**Filtro MATRIZ:** usa `_readMatrizGrid_().byProduct[pb].previstoCargos` para mostrar solo EPPs asignados al cargo del trabajador (col CARGO de REGISTRO). Admin (ADMIN_EMAIL) omite el filtro.
**Funciones públicas:**
- `obtenerAlertasVencimientos(dniLogin)` — alertas vencidas/próximas filtradas por MATRIZ del cargo
- `verificarAlertasCompletas(dniLogin)` — resumen tieneVencimientos + tienePendientes

---

### CodeMapa.js — Mapa de Riesgos
**Spreadsheet:** MAPA_RIESGOS (1EfQvY59m1l1SB_GD__CzL-qJQFdbtYzM9Y2q1u2L3cI)
**Hojas:** MAPAS, ICONOS
**Funciones públicas:**
- `getMapas()` / `getMapasPage(search, offset, limit)` — listado con paginación
- `saveMapa(mapa)` / `deleteMapa(id)` — CRUD
- `getIconos()` / `subirIcono(tipo, nombre, base64)` / `eliminarIcono(iconoId)`

---

### HhtCode.js — HHT
**Responsabilidad:** Horas Hombre Trabajadas — hoja HHT en PERSONAL SS.
**Hojas:** HHT (9 columnas, en PERSONAL SS)

---

### ComunicadosCode.js — Comunicados
**Responsabilidad:** Comunicados internos con imagen, estado activo/inactivo.
**Hoja:** COMUNICADOS (en PERSONAL SS, creada automáticamente si no existe)
**Drive:** COMUNICADOS_FOLDER_ID (1_a0rg1PK13NtkDLQ-y-tqKslo9aOc8DV)

---

### Homecode.js — Avisos ERP
**Responsabilidad:** Avisos del sistema ERP (hoja AVISOS en PERSONAL SS).

---
## 3. ESTRUCTURA REAL DE SPREADSHEETS

### 3.1 PERSONAL SS (1NDDHlTfWxmObgm8JZu5WAnCECB3gU6e_k7o_sFcMrkw)

**Hoja: PERSONAL** — hoja principal de usuarios del sistema
| Col | Índice (0-based) | Dato | Tipo |
|---|---|---|---|
| A | 0 | ID numérico | Número |
| B | 1 | DNI / Usuario login | Texto |
| C | 2 | Nombre completo | Texto |
| D | 3 | (no mapeado) | — |
| E | 4 | Empresa | Texto |
| F | 5 | (no mapeado) | — |
| G | 6 | Cargo | Texto |
| H-K | 7-10 | (no mapeados) | — |
| L | 11 | Condición laboral | ACTIVO / LICENCIA / SUSPENSIÓN DE LABORES / POSTULANTE / LIQUIDADO / TRASPASO / VISITANTE |
| M | 12 | Email | Texto |
| N | 13 | Contraseña | Texto |
| O | 14 | URL Foto | Texto |
| P | 15 | Autorizado login (SI/NO) | Texto |
| Q | 16 | Accesos (módulos separados por coma) | Texto |
| R | 17 | URL Firma | Texto |
| S | 18 | Fecha de Cese | Fecha |

**loginData() retorna (índice → campo sesión):**
row[0]→id, row[1]→usuario, row[2]→nombre, row[6]→cargo, row[4]→empresa,
row[12]→email, row[13]→password, row[14]→foto, row[16]→accesos, row[17]→firma

**localStorage 'userName' (JSON en frontend):**
```
{ id, usuario, nombre, cargo, empresa, email, password, foto, accesos, firma, token }
```

**Otras hojas en PERSONAL SS:**
- Log — registro de logins (fecha, nombre)
- LISTAS — listas maestras del sistema
- RESUMEN — celda J1 = color personalizado usuario
- MOF — áreas (col B) y cargos (col C) para Rol de Turnos
- Accesos — contraseñas login (col B) y borrado (col C)
- INCOMPATIBILIDADES — datos de incompatibilidades
- HHT — Horas Hombre Trabajadas (9 columnas)
- AVISOS — avisos ERP
- COMUNICADOS — comunicados internos (creada automáticamente)
- LISTAS (en Code.js) — col B-P: listas desplegables

---

### 3.2 CHECK SS (12KkPwl_gfQCkqS9ZHsp4hS2fFkebgNbszvTDtZELObU)

**Hoja: B DATOS** — registros de inspecciones (hasta col 100 / CV)
| Col | N° (1-based) | Dato | Funciones que la usan |
|---|---|---|---|
| A | 1 | ID/Timestamp Unix | saveDataCheck, getDatosRegistroCheck |
| B | 2 | Empresa | saveDataCheck |
| C | 3 | Equipo | saveDataCheck |
| D | 4 | Código/Placa | saveDataCheck |
| E | 5 | Área | saveDataCheck |
| F | 6 | Proceso | saveDataCheck |
| G | 7 | Supervisor de Operaciones | saveDataCheck |
| H | 8 | Lugar | saveDataCheck |
| I | 9 | Plan de Acción | saveDataCheck |
| J | 10 | Fecha (Date) | saveDataCheck, getDatosRegistroCheck |
| K | 11 | Imagen principal URL | saveDataCheck |
| L | 12 | Responsable del Área | saveDataCheck, generarPDFdesdeHTML |
| M | 13 | Estado (Conforme/Abierto/En Proceso/Cerrado) | calcularEstadoCheck |
| N | 14 | Tipo Inspección (Planeada/No Planeada) | saveDataCheck |
| O | 15 | Ítems CSV (Si,No,NA,...) | saveDataCheck |
| P-AQ | 16-43 | Fórmulas calculadas (setFormula copia desde fila anterior) | setFormula |
| BQ-BS | 69-71 | Sección 1: foto obs, foto sub, comentario | agregarFotosSeccion |
| BT-BV | 72-74 | Sección 2: foto obs, foto sub, comentario | agregarFotosSeccion |
| ... | ... | Patrón repite cada 3 cols, hasta sección 10 | |
| CQ-CS | 97-99 — | Sección 10: foto obs, foto sub, comentario | agregarFotosSeccion |
| CU | 99 | Firma trabajador URL (o texto 'Area inspeccionada sin personal presente') | saveDataCheck |
| CV | 100 | Firma supervisor URL | saveDataCheck |

**IMPORTANTE — Formato fotos obs (cols 69-98):**
Formato codificado: num::url~~num2::url2 (parseObsCell lo parsea)
Subsanación EN_GESTION: la celda contiene literalmente 'EN_GESTION' como valor del ítem

**Hoja: INVENTARIO** — equipos para dropdown checklist
| Col | Índice frontend (r[N]) | Dato |
|---|---|---|
| B | r[0] | Empresa |
| C | r[1] | Área |
| D | r[2] | Equipo (nombre) |
| E | r[3] | Código/Placa |
| F | r[4] | Cargo responsable (filtro por login) |
| H | r[6] | Mensaje 1 |
| L | r[10] | Lugar(es) separados por coma |
| M | r[11] | Frecuencia de inspección |
| O | r[13] | Estado (RETIRADO = excluir) |
| R | r[17] | Flag disponible (calculado en getDropDownarray) |
| S | r[18] | Lugares disponibles en período (calculado) |

**Filtro cargo col F en getDropDownarray (backend) y aftersecondDropDownChange (frontend):**
- Col F vacía = sin restricción → visible a todos
- Col F con valor = solo visible si cargo del usuario coincide (comparación includes)
- Supervisores (cargo.includes('supervisor')) = ven todo sin filtro

**Otras hojas en CHECK SS:**
- CHECK LIST — ítems de verificación por equipo
- HISTORIAL — historial de checklists
- ACTUAL — celda J2 valor actual, celdas T1/U1 contadores
- MENÚ — celda B24 email destinatario alertas
- Acceso — contraseñas módulo check
- FORMATO — hoja legacy para PDF (ya no se usa en flujo principal)

---

### 3.3 EPP SS (1Mxy5SkDdLy1Ihct844uLq5ZALe-RFarDfWo9j65kBcE)

**Hojas:** STOCK, MOVIMIENTOS, REGISTRO, MATRIZ, ALMACENES
Ver IDX en sección 2 (EppCode.js) para columnas exactas 1-based.

**REGISTRO col 25 (ESTADO):** Pendiente → Confirmado (cuando trabajador firma)
**REGISTRO col 26 (FECHA_CONFIRMACION):** Fecha de confirmación por firma
**REGISTRO col 23 (FIRMA_URL):** URL firma trabajador subida a Drive (FOLDER_IDEPP)

---

### 3.4 DESVIOS SS (1eIJfA7dAlkQ1rXcRGC2qSFnvZ-jYIPn8cA_TbUZcWZE)
**Hojas:** B DATOS (registros desvíos), INSPECCIÓN, ANALISIS, FICHA RAC T1, MENÚ, Acceso
Usa SPREADSHEET_IDS.desvios — definido en Code.js.

---

### 3.5 CAPACITACIONES SS (1Ev5_B3jMtjy_xXt13NYBXYwFA-maFAeLSKfiCFIsMQo)
**Hojas:** Matriz (configuración cursos/cargos), B DATOS (registros), LIST
- Matriz fila 17 (índice 16) = nombres de cursos (desde col 4)
- Matriz fila 18+ = cargos (col 3) y si aplica (col 4+)
- B DATOS cols A-P: registro de capacitaciones
- REUNIONES (en este SS): reuniones de seguridad (PassoCode.js)

---

### 3.6 ROL SS (12h2yVs0NlD3h3zMYl_93o7ohOKzurxcPZXifoTyVigE)
**Hojas:** BD_Detalle, BD_Resumen_Semanal (creadas automáticamente por saveFullReport)
- BD_Detalle: FECHA, LUNES_SEMANA, ID_EMP, DNI, NOMBRE, ZONA, TURNO, HORAS, TIPO, CODIGO, REGISTRADO_EL
- BD_Resumen_Semanal: LUNES_SEMANA, ID_EMP, DNI, NOMBRE, HH_TOTAL, HH_REGULAR, HH_EXTRA, NOCHES, DIAS_TRAB, DIAS_DESC, TIENE_AUS, DETALLE_AUS, ESTADO, REGISTRADO_EL

---
---

## 4. PATRONES DE CÓDIGO

### 4.1 Llamada al backend desde frontend
```javascript
// PATRÓN CORRECTO — siempre incluir withFailureHandler
google.script.run
  .withSuccessHandler(function(result) { /* usar result */ })
  .withFailureHandler(function(err) { console.error('Error:', err.message); })
  .miFuncionBackend(param1, param2);
```
Actualmente todas las 208 cadenas `google.script.run` tienen `withFailureHandler` ✅ — verificado con parser de balance de llaves. Siempre incluir al crear nuevas.

### 4.2 Caché en backend GAS
```javascript
// PATRÓN ESTÁNDAR: CacheService (TTL máx 21600 seg = 6h)
function getDatos() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('clave');
  if (cached) return cached;
  var result = /* leer Sheets */;
  cache.put('clave', JSON.stringify(result), 600); // 10 min
  return JSON.stringify(result);
}
```
NO usar: variables globales como cache (se resetean entre invocaciones).
NO usar: cache manual con timestamp — usar TTL de CacheService.

### 4.3 Llamada a Gemini API
```javascript
// FUNCIÓN CENTRALIZADA en Code.js — firma actual:
// prompt:          texto del prompt (null si se usa partsOverride)
// generationConfig: objeto opcional { temperature, maxOutputTokens, responseMimeType }
// partsOverride:   array de parts para enviar imágenes/archivos junto con texto
// modelOverride:   nombre del modelo (por defecto 'gemini-2.5-flash')
function _callGemini(prompt, generationConfig, partsOverride, modelOverride) { ... }

// Solo texto:
_callGemini(prompt)
// Texto + config:
_callGemini(prompt, { temperature: 0.1, maxOutputTokens: 2048 })
// Con imagen (inlineData):
const parts = [{ text: prompt }, { inlineData: { mimeType, data: base64 } }];
_callGemini(null, null, parts)
// Modelo custom:
_callGemini(prompt, null, null, 'gemini-pro')
```
Todos los módulos deben usar _callGemini() — NO reimplementar la llamada HTTP.
API_KEY se lee desde ScriptProperties clave `GEMINI_KEY` con fallback al valor hardcodeado.

### 4.4 Envío de correo
```javascript
// ESTÁNDAR: MailApp (menor scope requerido)
MailApp.sendEmail({ to: email, subject: asunto, htmlBody: html });
// NO usar GmailApp salvo que se necesite hilo/etiquetas
```

### 4.5 Normalización de texto
```javascript
// FUNCIÓN GLOBAL en Code.js
function _norm(s) {
  return String(s||'').trim().toUpperCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'');
}
```
NO redefinir en módulos — ya definida en Code.js, NotificacionesCode.js, CodeMapa.js.

### 4.6 Lectura segura de celda
```javascript
// Helpers en EppCode.js — usar en módulo EPP
function _str(v) { return String(v == null ? '' : v).trim(); }
function _num(v) { return parseFloat(v) || 0; }
function _today() { var d = new Date(); d.setHours(0,0,0,0); return d; }
```

### 4.7 Guardar fila nueva en Sheet
```javascript
// Patrón BatchWrite — evitar appendRow en loop
var rows = data.map(function(item) { return [item.a, item.b, item.c]; });
sheet.getRange(sheet.getLastRow()+1, 1, rows.length, rows[0].length).setValues(rows);
```

### 4.8 Gestión de sesión en frontend
```javascript
// Leer sesión
var user = JSON.parse(localStorage.getItem('userName') || '{}');
// Campos disponibles: id, usuario, nombre, cargo, empresa, email, password, foto, accesos, firma, token
// Verificar acceso a módulo:
var accesos = (user.accesos || '').split(',').map(s => s.trim());
if (!accesos.includes('CHECK')) { /* sin acceso */ }
```

---

## 5. ROLES Y ACCESOS

**Campo 'accesos' en localStorage** — lista separada por comas, ej: `CHECK,EPP,DESVIOS`

| Módulo token | Función que lo verifica | Descripción |
|---|---|---|
| CHECK | index.html router | Inspecciones checklist |
| EPP | index.html router | EPP: stock, movimientos, asignaciones |
| DESVIOS | index.html router | Desvíos y observaciones |
| CAPACITACIONES | index.html router | Capacitaciones y charlas |
| EVENTOS | index.html router | Accidentes y eventos |
| MAPA | index.html router | Mapa de riesgos |
| ROL | index.html router | Rol de turnos |
| IPERC | index.html router | Matriz IPERC |
| GRAFICOS | index.html router | KPIs y gráficos |
| USUARIOS | index.html router | Solo admin: gestión usuarios |
| HHT | index.html router | Horas Hombre Trabajadas |

**Supervisores** — `cargo.includes('SUPERVISOR')` → ven todos los equipos sin filtro de cargo en Check.
**Condición col L — acceso y notificaciones:**
- `CONDICIONES_ACCESO = ['ACTIVO', 'LICENCIA']` — definida en Code.js; única fuente de verdad
- Login bloqueado si col L no está en CONDICIONES_ACCESO O si col P = 'NO'
- **Excepción "Todo"**: usuarios con col Q (accesos) = `'todo'` (case-insensitive) siempre pueden hacer login y recibir notificaciones, independientemente de col L
- Login bloqueado si col L no está en CONDICIONES_ACCESO O si col P = 'NO'
- Notificaciones (bulk/broadcast) solo se envían a ACTIVO y LICENCIA — sin excepciones
- `verificarYEnviarAlertasInspeccion()`: filtra `dnisEnTurno` contra `_getDnisActivos()` antes de enviar push — LIQUIDADO/SUSPENDIDO no reciben aunque estén en el ROL
- Col P (AUTORIZADO SI/NO) se mantiene como segunda barrera de seguridad
- `actualizarCondicionUsuario(id, condicion)` — escribe col L (col 12, 1-based) desde Usuarios.html
- `obtenerUsuariosPaginado` retorna col L en index 6 del array de columnas (header "CONDICIÓN")
**Autorizado 'SI'** — col P de PERSONAL debe ser 'SI' para permitir login.

---

## 6. PROTOCOLO DE TRABAJO

### 6.1 Antes de modificar una función
1. Leer el archivo completo donde está definida
2. Buscar TODOS los callers: `grep -r 'nombreFuncion' /home/user/SST/`
3. Verificar si hay versión _v2 del mismo nombre
4. Revisar qué columnas de Sheets lee (ver Sección 3 para índices exactos)

### 6.2 Al agregar nueva columna a un Sheet
1. Actualizar el mapa de columnas en CLAUDE.md Sección 3
2. Si es EPP: actualizar objeto IDX en EppCode.js
3. Si es CHECK: actualizar comentarios en saveDataCheck()
4. Buscar funciones que lean el mismo sheet y actualizar

### 6.3 Al crear nueva función backend
1. Nombre en camelCase, idioma español o inglés consistente con el módulo
2. Retornar siempre JSON.stringify() — nunca objetos directos
3. Agregar Logger.log() al inicio y al final con resultado
4. Envolver en try/catch con return JSON.stringify({error: e.message})

### 6.4 Flujo de firma EPP (REGISTRO col 25)
```
Estado inicial:  col 25 = 'Pendiente'
Trabajador firma → sube imagen a FOLDER_IDEPP → URL en col 23
                → col 25 = 'Confirmado'
                → col 26 = fecha confirmación
```

### 6.5 Flujo de fotos en inspecciones CHECK
```
Foto nueva: subida a Drive folderimgcheck → URL guardada en cell
Formato obs: 'num::url~~num2::url2' — parseObsCell() lo desempaqueta
Subsanación EN_GESTION: la celda contiene literalmente 'EN_GESTION'
Col 69-98 (BQ-CS): grupos de 3 cols → [foto_obs, foto_sub, comentario]
Col 99 (CU): firma trabajador | Col 100 (CV): firma supervisor
```

---

## 7. COMANDOS ÚTILES

```bash
# Ver funciones exportadas de un archivo GAS
grep -n '^function ' /home/user/SST/CheckCode.js

# Buscar todos los callers de una función
grep -rn 'saveDataCheck' /home/user/SST/

# Ver todas las constantes de IDs
grep -n 'SPREADSHEET_ID\|FOLDER_ID\|openById' /home/user/SST/Code.js

# Buscar columnas usadas en un sheet
grep -n 'row\[\|\[0\]\|\[1\]\|\[2\]' /home/user/SST/CheckCode.js | head -40

# Ver qué archivos usan SPREADSHEET_IDS
grep -rn 'SPREADSHEET_IDS\.' /home/user/SST/

# Buscar funciones sin withFailureHandler
grep -rn 'google.script.run' /home/user/SST/*.html | grep -v 'withFailureHandler' | grep -v 'withSuccessHandler'

# Ver estructura del repositorio
ls -la /home/user/SST/*.js /home/user/SST/*.html
```

---

## 8. REGLAS CRÍTICAS

1. **SPREADSHEET_IDS** está definido en Code.js — NO redefinir en otros archivos
2. **IDX** y **SHEPP** son objetos de EppCode.js — solo para módulo EPP
3. **_callGemini()** en Code.js — todos los módulos deben usarla
4. **_norm()** en Code.js — no duplicar en módulos
5. Las columnas de CHECK B DATOS son **1-based** (notación Sheet); acceso JS es `row[col-1]`
6. Las columnas de PERSONAL son **0-based** (acceso JS directo `row[idx]`)
7. NUNCA usar `appendRow()` en loop — siempre batch `setValues()`
8. SIEMPRE agregar `withFailureHandler` en `google.script.run`
9. Los IDs de Spreadsheet y Drive NO deben hardcodearse fuera de Code.js
10. `MailApp` es el estándar para correos — no usar GmailApp salvo necesidad específica

---

## 9. SISTEMA DE TEMAS (Design Themes)

### 9.1 Tema Normal (por defecto)
- Fuente: Inter, system-ui, sans-serif
- Colores: Bootstrap 5.3 + azul primario #0d6efd
- Bordes redondeados: 10-12px en botones e inputs
- Sombras suaves (box-shadow blur)
- Sidebar: fondo blanco, links azul en hover

### 9.2 Tema Neo Brutalism
- Activado con clase CSS `body.neo-brutalism`
- Fuente: Space Grotesk (cargada via Google Fonts en css.html)
- Paleta: amarillo #FFDD00, negro #000, naranja #FF5F1F, fondo crema #F5F0E8
- Bordes: 2.5-3px sólidos #000, `border-radius: 0`
- Sombras: offset sólido sin blur `3px 3px 0 #000`
- Sidebar: fondo amarillo #FFDD00, links negros → hover negro/amarillo
- Navbar: fondo negro, borde inferior amarillo
- SweetAlert2: mismo estilo brutalist (popup square + box-shadow)

### 9.3 Toggle del tema
- **Botón**: `#btn-theme-toggle` en la navbar (ícono ⚡ normal / 🎨 neo)
- **Persistencia**: `localStorage('sst-theme')` → valores: `'default'` | `'neo-brutalism'`
- **Init**: `_initTheme()` se ejecuta en `DOMContentLoaded` (index.html)
- **Función toggle**: `toggleNeoTheme()` en index.html
- **CSS override**: todos los overrides en `css.html` bajo el selector `body.neo-brutalism`

### 9.5 Escala tipográfica del sistema (definida en css.html :root)

```css
--font-xs:   0.72rem;   /* 11.5px — badges, estado-badge, helper text */
--font-sm:   0.76rem;   /* 12.2px — chips de filtro (check-chip, chip-mes, epp-chip, maestro-chip, chipepp, filter-tab, badge-frec) */
--font-base: 0.84rem;   /* 13.4px — form-control, form-select, table td */
--font-md:   0.875rem;  /* 14px   — form-label, subtítulos de módulo */
--font-lg:   0.95rem;   /* 15.2px — títulos principales */
```

**Reglas obligatorias:**
- **Chips de filtro** → siempre `var(--font-sm)` — NO usar px ni rem hardcodeados
- **Labels de formulario** → `var(--font-md)` vía `.form-label`
- **Inputs/selects** → `var(--font-base)` vía `.form-control` / `.form-select`
- **Encabezados de tabla** (`th`) → `var(--font-sm)` con `font-weight: 700`
- **Celdas de tabla** (`td`) → `var(--font-base)`
- **Badges** → `var(--font-xs)` con `!important`
- **NO usar** `style="font-size:12px"` en contenedores de chips — CSS ya lo maneja

### 9.6 Dónde agregar CSS nuevo
- **Estilos globales** (sidebar, navbar, componentes compartidos, tema): → `css.html`
- **Estilos de un módulo específico** (Check, Rol, EPP, Eventos, etc.): → `css-modulos.html` en la sección del módulo correspondiente
- **NUNCA** agregar `<style>` dentro de un archivo `.html` de módulo — todos los módulos HTML deben estar libres de bloques `<style>`

### 9.4 Cómo agregar estilos Neo Brutalism a un módulo nuevo
```css
/* En css.html, dentro del bloque body.neo-brutalism */
body.neo-brutalism .mi-nuevo-componente {
  border: 2.5px solid #000 !important;
  border-radius: 0 !important;
  box-shadow: 4px 4px 0 #000 !important;
}
```

---

## 10. ANTI-PATRONES DETECTADOS (no reproducir)

| Anti-patrón | Dónde aparece | Estado |
|---|---|---|
| `getDropDownarray_v2()` — duplicar con sufijo _v2 | CheckCode.js | Pendiente |
| `console.log` en backend GAS | DesvioscCode.js | Pendiente |
| `GmailApp.sendEmail()` mezclado con MailApp | DesvioscCode.js | Pendiente |
| Llamadas directas a Gemini API | ~~DesvioscCode, CheckCode, CapaciCode — CORREGIDO~~ | ✅ Resuelto |
| `enviarTelegram()` duplicada en Code.js | ~~Code.js — ELIMINADA~~ | ✅ Resuelto |
| Variables globales como cache | Varios archivos | Pendiente |
| IDs hardcodeados en múltiples archivos | ~~Code.js, RolCode.js (TODOS), CodeMapa.js — CORREGIDO~~ | ✅ Parcial (ver nota) |
| `setValue()` individual en loop (`agregarUsuario`, `actualizarUsuario`) | ~~Code.js — CORREGIDO~~ | ✅ Resuelto |
| `appendRow()` en loop | ~~CapaciCode.js guardarPreguntasMultiples — CORREGIDO~~ | ✅ Resuelto |
| `google.script.run` sin withFailureHandler | ~~40 archivos~~ — 208 cadenas, TODAS con handler ✅ | ✅ Resuelto |
| Redefine _norm() local | NotificacionesCode.js, CodeMapa.js | Pendiente |
| `include()` redefinida en RolCode.js | ~~RolCode.js — ELIMINADA~~ | ✅ Resuelto |
| Mezcla Bootstrap 4/5.1/5.3 | Bootstrap 5.3.3 consistente en todo el repo | ✅ Resuelto |
| Mezcla Chart.js 3.9.1 + 4.4.0 + 4.4.4 | ~~Graficosindex, Test — CORREGIDO~~ | ✅ Resuelto (todos en 4.4.4) |
| Mezcla Font Awesome 6.4.0 vs 6.5.0 | ~~index, Examen, TestCheck — CORREGIDO~~ | ✅ Resuelto (todos en 6.5.0) |
| SweetAlert2 `@11` flotante sin versión fija | ~~index.html, Examen.html — CORREGIDO~~ | ✅ Resuelto (`@11.14.0` fijo) |
| `notificarCapacitacion` duplicada con firmas distintas en Telegram.js y NotificacionesCode.js | ~~RENOMBRADAS: `notificarCapacitacionTelegram` y `notificarCapacitacionPush`~~ | ✅ Resuelto |
| `describirImagen()` muerta (comentario explícito "NO ES USADA") | ~~DesvioscCode.js — ELIMINADA~~ | ✅ Resuelto |
| Credenciales hardcodeadas (Gemini, Telegram, Push) | ~~Code.js, Telegram.js, NotificacionesCode.js — CORREGIDO~~ | ✅ Parcial (fallback temporal) |

**Nota IDs parcial:** AlertasCode.js aún tiene `ROL_SS_ID_ALERTAS` hardcodeado; DesvioscCode.js tiene folder IDs de imágenes/PDFs. Migrar cuando se toque esos módulos.

---

## 11. CHECKEOS DE VALIDACIÓN

Antes de hacer commit, verificar:

```bash
# 1. No hay google.script.run sin withFailureHandler nuevo
grep -rn 'google.script.run' /home/user/SST/*.html | grep -v 'withFailureHandler' | grep -v '//'

# 2. No hay console.log en archivos .js backend
grep -rn 'console\.log' /home/user/SST/*.js

# 3. No hay IDs de Sheets hardcodeados fuera de Code.js (excepción documentada: AlertasCode.ROL_SS_ID_ALERTAS, DesvioscCode folder IDs)
grep -rn '"1[A-Za-z0-9_-]\{40,\}"' /home/user/SST/*.js | grep -v Code.js

# 4. No hay appendRow en loop
grep -B5 'appendRow' /home/user/SST/*.js | grep -E 'for|forEach|map|while'

# 5. Toda función nueva retorna JSON.stringify
# (revisión manual del diff)
```

---

## 12. LECCIONES APRENDIDAS — DECISIONES CRÍTICAS DE ARQUITECTURA

### L1 · HTML parciales NO procesan template directives GAS

**Problema:** Se intentó dividir `ReportesLaboral.html` en `rlCss.html` + `rlScript.html` usando `<?!= include('rlCss') ?>` para evitar timeouts al subir archivos grandes.

**Lo que pasó:** Las directivas `<?!=...?>` solo se procesan en el archivo raíz cargado via `HtmlService.createTemplateFromFile()` (actualmente `index.html`). Los parciales incluidos con `include()` se sirven como HTML estático — las directivas aparecen literalmente en el DOM.

**Regla:** Los archivos HTML de módulos (Check.html, MovimEpp.html, ReportesLaboral.html, etc.) son **parciales estáticos**. No usar `<?!=...?>` en ellos. El JS de cada módulo permanece en su archivo HTML. El CSS de los módulos vive en `css-modulos.html` (NO en el HTML del módulo).

---

### L2 · Push a GitHub: el proxy CCR es de solo lectura por sesión

**Problema:** El proxy local `127.0.0.1:XXXXX/git/...` cambia de puerto en cada invocación y solo permite `git fetch` (GET). `git push` devuelve 403 siempre. El MCP `create_or_update_file` también devuelve 403 (la integración de Anthropic tiene acceso de solo lectura al repo).

**Solución permanente:** Configurar el remote con un PAT de GitHub al inicio de cada sesión:
```bash
PAT="github_pat_..."
git remote set-url origin "https://x-access-token:${PAT}@github.com/vico-renxo/SST.git"
```
**IMPORTANTE:** El proxy puede sobrescribir el remote al cambiar de puerto. Ejecutar `git remote set-url` justo antes de cada push si el primero falla con 403.

---

### L3 · Archivos HTML grandes (>600 líneas) — estrategia de edición

**Problema:** Pasar contenido de archivos grandes como parámetro de herramientas MCP causa "stream idle timeout — partial response received". El timeout ocurre durante la generación del parámetro, no durante la llamada HTTP.

**Decisiones correctas:**
1. Usar `Edit` (diff) en lugar de `Write` (archivo completo) siempre que sea posible.
2. Nunca leer un archivo completo y pasarlo íntegro a una herramienta MCP en el mismo turno.
3. Si se necesita reescribir un archivo grande, usar `Bash` con heredoc (`cat > archivo << 'EOF'`) — es una operación local que no genera timeout.
4. Hacer push después de cada commit individual, no acumular commits.

**Decisiones incorrectas a evitar:**
- Intentar pasar 40KB+ como parámetro `content` a `mcp__github__create_or_update_file`.
- Crear archivos de "split" para evitar el timeout — complica la arquitectura sin resolver el problema raíz.
- Minificar el CSS/JS para reducir líneas — no soluciona el timeout del stream, solo oscurece el código.

---

### L4 · Chart.js en módulos GAS — patrón correcto

**Para agregar Chart.js** a un módulo HTML de GAS:
1. Agregar CDN al inicio del archivo: `<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.min.js"></script>`
2. Guardar instancias de charts en variables de módulo: `let _chDonut = null`
3. Siempre destruir antes de recrear: `function _destroyChart(ref) { if(ref) { try { ref.destroy(); } catch(e){} } return null; }`
4. Para charts de tamaño dinámico (basado en número de filas), ajustar `canvas.style.height` antes de crear el chart y usar `maintainAspectRatio: false`.

---

### L5 · `aspect-ratio` falla en flex containers cuando el CSS está en el `<head>`

**Problema:** Los módulos HTML originalmente tenían `<style>` blocks inline (en el body del SPA). Al centralizarlos en `css-modulos.html` (cargado en `<head>`), las cards tiktok-style (`.tiktok-card`) dejaron de tener la altura correcta a pesar de tener `aspect-ratio: 9/16`.

**Causa raíz:** `aspect-ratio` en un elemento con `display: flex` puede fallar silenciosamente cuando el CSS está en el `<head>` y el elemento existe en un contexto de layout complejo (tab-panes, containment). Los inline `<style>` del body tenían mayor prioridad en cascada (posición tardía en el documento), enmascarando el problema.

**Solución definitiva — truco `padding-bottom`:**
```css
/* ✅ Funciona en TODOS los contextos: flex, grid, contain, head o body */
.tiktok-card {
  position: relative;
  overflow: hidden;
  width: 100%;
  height: 0;
  padding-bottom: 177.78%; /* = (16/9) × 100% para ratio 9:16 */
}
/* El contenido que iría al fondo con justify-content:flex-end
   debe ser position:absolute bottom:0 en su lugar */
.tiktok-text {
  position: absolute;
  bottom: 0;
  left: 0;
  right: 0;
}
```

**Regla:** Para cualquier card con proporción fija (tiktok-style, portrait, landscape), siempre usar `padding-bottom` en lugar de `aspect-ratio`. Es universalmente compatible y no depende del contexto de layout.

**Fórmulas comunes:**
- 9:16 portrait → `padding-bottom: 177.78%`
- 16:9 landscape → `padding-bottom: 56.25%`
- 1:1 cuadrado → `padding-bottom: 100%`
- 4:3 → `padding-bottom: 75%`

---

### L6 · css-modulos.html — estructura y errores a evitar

**Estado actual:** `css-modulos.html` tiene UNA sola `<style>` tag que envuelve todo el archivo (línea 1: `<style>`, línea final: `</style>`). Se carga en el `<head>` de index.html via `<?!= HtmlService.createHtmlOutputFromFile('css-modulos').getContent() ?>`.

**Errores detectados y corregidos:**
- `//comment` (comentario estilo JS) dentro de CSS → inválido, usar `/* comment */`
- `*/` suelto sin `/*` previo → dangling closer, produce error de parseo silencioso
- `contain: layout style paint` en un contenedor de cards → puede interferir con el cálculo de tamaños de hijos; evitar salvo necesidad probada

**Verificar balance de comentarios antes de editar css-modulos.html:**
```bash
python3 -c "
content = open('/home/user/SST/css-modulos.html').read()
import re
pos = 0; in_c = False; issues = []
while pos < len(content):
    if not in_c:
        idx = content.find('/*', pos)
        if idx == -1: break
        in_c = True; cs = idx; pos = idx + 2
    else:
        idx = content.find('*/', pos)
        if idx == -1:
            issues.append(f'UNCLOSED comment at line {content[:cs].count(chr(10))+1}')
            break
        in_c = False; pos = idx + 2
[print(i) for i in issues] or print('OK')
"
```

---

### L7 · Neo Brutalism — cobertura completa de módulos en css-modulos.html

**Estado:** El bloque `body.neo-brutalism` en `css-modulos.html` (sección "NEO BRUTALISM — Overrides de módulos específicos", ~línea 5246 en adelante) cubre todos los módulos del proyecto.

**Arquitectura del bloque de overrides:**
- **css.html** — 56 reglas globales: navbar, sidebar, .btn, .card, .modal-content, table, .badge, .check-chip, .tiktok-card, SweetAlert2
- **css-modulos.html** — ~650 reglas modulares organizadas por módulo

**Namespaces CSS por módulo (para agregar overrides nuevos):**

| Módulo | Namespace / selector raíz | Clases clave con override |
|---|---|---|
| IndexDesvios/Check | `.estado-badge`, `.badge-potencial` | Sí |
| Asignaciones EPP | `.epp-unified-card`, `.epp-chip`, `.epp-badge-*`, `.firma-modal`, `.panel-pendientes`, `.alerta-modal-grande` | Sí |
| MovimEpp | `.epp-scope .card-prod`, `.tag-previsto`, `.badge-stocklow`, `.variant-pill` | Sí |
| EPPMaestro | `.kpi-card`, `.maestro-chip`, `.celda-epp`, `.st-OK/.NOTIF/.ENTR`, `.tooltip-epp` | Sí |
| MatrizApp | `.chipepp`, `.chip-input` | Sí |
| EditCheck | `.obs-card-sub` variants, `.obs-photo-box` | Sí |
| Check/Test | `.obs-thumb-wrapper`, `.ia-section`, `.ia-fortaleza/.oportunidad/.objetivo`, `.ev-score-*` | Sí |
| PASSO | `.badge-frec`, `.celda-cumple/.falta/.parcial`, `.celda-porcentaje-*`, `.celda-mes/.prog/.vacia`, `.fila-actividad/.gerencia`, `.col-*` | Sí |
| Eventos | `.ev-card-face`, `.ev-front-content`, `.ev-card-back`, `.ev-more-panel`, `.piramide-container`, `.nivel`, `.bueno/.malo` | Sí |
| MapaRiesgos | `#riskmaps-module .tiktok-card`, `.map-title-overlay`, `.type-badge`, `.tb-preventivo/.informativo/.restrictivo/.obligatorio`, `.legend-chip/.sidebar`, `.zoom-controls`, `.rm-item` | Sí |
| Rol | `#rol-module-wrapper .modal-box`, `.shift-card`, `.chip-work/.rest/.vac/.fal/.med/.otr/.per`, `.filter-bar-container`, `.filter-tab`, `.compliance-panel`, `.heatmap-wrapper`, `.tile`, `.t-ok/.warn/.danger/.empty`, `.tab-count` | Sí |
| Graficosindex | `.chart-container`, `.pronostico-ia-container .card/.header/.btn-*/.status-indicator` | Sí |
| Capacitaciones | `.rl-kpi-card`, `.rl-topic-bar`, `.radio-label`, `.choice`, `.radio-group-label`, `.codigo-*` | Sí |
| ReportesLaboral | `.rl-card`, `.rl-panel`, `.rl-table th`, `.rl-tab-btn`, `.rl-badge`, `.rl-search-box`, `.rl-vac-card`, `.rl-worker-card`, `.rl-rank-*`, `.rl-seg-*`, `.rl-progress-*` | Sí |
| BuscadorCharlas | `.card-charla`, `.charla-header`, `.chip-mes-charla` | Sí |
| BuscadorCap | `.badge-pct-ok/.med/.low`, `#modalDetalleBox`, `.det-tema-hdr` | Sí |
| Comunicados | `.com-card`, `.com-badge-tipo`, `.status-pill.activo/.inactivo`, `.com-preview-box` | Sí |
| IPERC | `#iperc-module-wrapper .card/.badge`, `th` | Sí |
| Evaluacion | `.star-rating`, `.star`, `.star.selected` | Sí |
| Asignaciones/Global | `.estado-confirmado/.pendiente/.rechazado`, `.avatar-circle`, `.staff-item` | Sí |
| Listas/HHT | `#miVista table/th/td/input` | Sí |

**Paleta Neo Brutalism (referencia rápida):**
- Amarillo activo: `#FFDD00` (texto #000)
- Naranja alerta: `#FF5F1F` (texto #fff)
- Negro primary: `#000` (texto #FFDD00)
- Crema fondo: `#F5F0E8`
- Borde: `2.5px solid #000`, shadow: `3px 3px 0 #000` (sin blur)

**Regla al agregar nuevo componente:**
```css
/* En css-modulos.html, al final del bloque NEO BRUTALISM */
body.neo-brutalism .mi-nuevo-componente {
  border-radius: 0 !important;
  border: 2.5px solid #000 !important;
  box-shadow: 3px 3px 0 #000 !important;
}
body.neo-brutalism .mi-nuevo-componente.active {
  background: #FFDD00 !important;
  color: #000 !important;
}
```

**Verificar cobertura:**
```bash
# Contar clases con override vs total
grep -c 'body\.neo-brutalism' /home/user/SST/css-modulos.html
# Buscar clase específica
grep 'body\.neo-brutalism.*\.mi-clase' /home/user/SST/css-modulos.html
```

---

---

### L8 · AlertasCode — ADMIN_EMAIL nunca debe caer en Session.getActiveUser()

**Bug introducido:** Al "mejorar" `ADMIN_EMAIL = "tu_correo_admin@gmail.com"` (placeholder) por
`PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL') || Session.getActiveUser().getEmail()`
se hizo que `esAdmin = true` para **todos** los usuarios (su correo siempre iguala al fallback de su propia sesión).

**Consecuencias:**
1. Filtro `if (!esAdmin && fila[COL_DNI] !== dniLogin) continue` se saltaba → se cargaban EPPs de TODOS los trabajadores.
2. Bloque `if (!esAdmin && cargoDelDni)` no ejecutaba → `productosEnMatriz = null` → sin filtro MATRIZ → EPPs ajenos al cargo aparecían.
3. El panel de alertas del trabajador mostraba EPPs vencidos de OTROS compañeros como si fueran suyos.

**Patrón correcto (invariable):**
```javascript
const adminEmail = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL') || '';
const esAdmin = (adminEmail !== '' && correoActual === adminEmail) || !dniLogin;
```

**Regla:** En cualquier función que use `esAdmin` para decidir si filtrar por DNI,
NUNCA hacer fallback del email admin al correo de la sesión activa.

---

### L9 · `grep` para withFailureHandler da falsos positivos — usar parser con balance de llaves

**Problema:** El comando estándar de validación:
```bash
grep -rn 'google.script.run' /home/user/SST/*.html | grep -v 'withFailureHandler'
```
reportó ~180 instancias "sin withFailureHandler". Al investigar con un parser de balance de llaves, se encontró que **todas las 208 cadenas ya tenían el handler**.

**Causa raíz:** El grep detectaba `google.script.run` dentro de un `withSuccessHandler` anidado de otra cadena. El handler de la cadena exterior ya existía pero quedaba fuera del fragmento que grep capturaba, haciendo parecer que faltaba.

**Regla:** Para auditar cobertura de `withFailureHandler`, NO usar grep simple. Usar el parser de balance de llaves o revisar manualmente cadena por cadena. El comando de validación en Sección 11 puede dar falsos positivos con handlers anidados.

**Estado actual (mayo 2026):** 208 cadenas `google.script.run` en 39 archivos HTML — todas con `withFailureHandler` ✅.

---

### L10 · Auditoría de seguridad — cambios implementados (mayo 2026)

**Credenciales movidas a PropertiesService (con fallback temporal):**
- `Code.js`: `API_KEY` → clave `GEMINI_KEY`
- `Telegram.js`: `botToken` → `TELEGRAM_BOT_TOKEN`, `chatId` → `TELEGRAM_CHAT_ID`
- `NotificacionesCode.js`: `PUSH_AUTH_TOKEN` → clave `PUSH_AUTH_TOKEN`

**Pendiente de resolver (decisión de arquitectura):**
- Contraseñas en texto plano en col N de PERSONAL — migración requiere plan de hashing sin romper logins
- CORS wildcard `*` en cloudflare-worker.js — cambiar al dominio real de la app
- Funciones `_asst_*` del asistente de voz — router llama 14 funciones que no existen en ningún archivo

**Deuda técnica conocida no bloqueante:**
- `AlertasCode.js`: `ROL_SS_ID_ALERTAS` aún hardcodeado
- `DesvioscCode.js`: folder IDs de imágenes/PDFs aún hardcodeados
- `_norm()`: aún redefinida en NotificacionesCode.js y CodeMapa.js
- Caché: módulos EPP, CHECK, Graficos usan variables en memoria en vez de CacheService

---

### L11 · Auditoría de consistencia frontend + dead code (mayo 2026)

**Versiones de librerías CDN estandarizadas:**
- Chart.js: 3.9.1 (Graficosindex) y 4.4.0 (Test) → **4.4.4 unificado**. Verificado sin breaking changes (sin `xAxes`/`yAxes` plural, sin `tooltips:` deprecado).
- Font Awesome: 6.4.0 (index, Examen, TestCheck) → **6.5.0 unificado** (alineado con Asignaciones).
- SweetAlert2: `@11` flotante → **`@11.14.0` fijo** (evita updates breaking silenciosos).
- Bootstrap: 5.3.3 ya consistente en todo el repo ✅.

**Dead code eliminado:**
- `describirImagen()` en DesvioscCode.js — comentario explícito "FUNCIONA, PERO NO ES USADA EN ESTA APLICACIÓN" + cero callers.
- `include()` redundante en RolCode.js — ya estaba definida globalmente en Code.js.

**Colisión silenciosa resuelta:**
- `notificarCapacitacion(datos)` (Telegram.js) y `notificarCapacitacion(dnis, tema, fecha)` (NotificacionesCode.js) tenían el mismo nombre con firmas distintas — en GAS la última definición cargada gana, generando bug latente.
- Renombradas a `notificarCapacitacionTelegram` y `notificarCapacitacionPush` para que el caller elija explícitamente.

**IDs centralizados completos:**
- RolCode.js eliminó `EMPLOYEES_SS_ID` y `SPREADSHEET_ID` locales — ahora usa `SPREADSHEET_IDS.rolEmpleados` y `SPREADSHEET_IDS.rolAlertas`.

**Falsos positivos detectados en la auditoría (no requieren acción):**
- TestCode.js balance de llaves: el conteo crudo da 261/260 pero un parser que excluye strings/comentarios da 208/208 ✅.
- IpercCode.js: agente reportó "extra braces"; parser real confirma balanceado 92/92 ✅.
- HTML "huérfanos" (Check.html, EditCheck.html, Evaluacion.html, etc.): no se cargan vía `include()` pero SÍ están referenciados desde el router SPA de index.html (34 referencias para Check.html). Carga dinámica, no son huérfanos.
- Funciones `_asst_*`: las 14 mencionadas como "no existen" en L10 SÍ existen — están en CapaciCode.js líneas 2735-3213.

**Regla:** los `grep` simples para detectar braces, callers o funciones huérfanas dan muchos falsos positivos. Siempre verificar con parser que excluya strings/comments antes de actuar.

---

*Fin de CLAUDE.md — Actualizar después de cada cambio estructural.*
