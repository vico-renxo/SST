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

| Alias | Descripción | ID |
|---|---|---|
| PERSONAL | Gestión personal / login | 1NDDHlTfWxmObgm8JZu5WAnCECB3gU6e_k7o_sFcMrkw |
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
├── index.html           # SPA: login + router de módulos
├── home.html            # Dashboard post-login
├── css.html             # Estilos globales Bootstrap 5.3 + FA icons
├── Check.html           # UI checklist inspecciones
├── MovimEpp.html        # UI movimientos EPP
├── Asignaciones.html    # UI firma EPP por trabajador
├── Capacitaciones.html  # UI capacitaciones
├── IndexDesvios.html    # UI desvíos
├── MapaRiesgos.html     # UI mapa de riesgos
├── Usuarios.html        # UI administración usuarios
├── Rol.html             # UI rol de turnos
└── Pagina web/
    ├── cloudflare-worker.js  # Worker push notifications
    └── worker.js
```
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
- `eliminarUsuarioPorUsuario(usuario)` — elimina + Telegram
- `buscarDatosPorNumero(numero)` — busca por DNI
- `getTodasLasListas()` — listas maestras con caché 5 min (CacheService, key: listas_globales_v5)
- `getRecordsList()` / `saveRecordsList(records)` — hoja LISTAS
- `getColor()` / `saveColor(color)` — color personalizado usuario (celda J1 RESUMEN)
- `enviarTelegram(mensaje)` — envía mensaje a Telegram
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
**Spreadsheets:** ROL_EMPLEADOS (1SrkbAD8aoLGCCr8oMh0yRp3iiRl0Du4WEpUU88zOCOc), ROL_ALERTAS (12h2yVs0NlD3h3zMYl_93o7ohOKzurxcPZXifoTyVigE)
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
- `notificarEntregaEPP(...)` / `notificarConfirmacionEPP(...)` / `notificarCapacitacion(...)`
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
- `enviarResumenDiario()` — trigger automático 8 AM
- `testTelegram()` — test de conexión

---

### AlertasCode.js — Alertas EPP
**Responsabilidad:** Detectar EPP vencidos o próximos a vencer.
**Dependencias:** getSpreadsheetEPP() (EppCode.js), IDX.REG, SHEPP (EppCode.js)
**Funciones públicas:**
- `obtenerAlertasVencimientos(dniLogin)` — alertas para un trabajador
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
| L | 11 | Estado (SI/ACTIVO/CESADO/BAJA) | Texto |
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
Actualmente 180+ llamadas sin withFailureHandler — agregar siempre al crear nuevas.

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
// FUNCIÓN CENTRALIZADA en Code.js
function _callGemini(prompt, modelOverride) {
  var model = modelOverride || 'gemini-2.5-flash';
  var API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_KEY') || API_KEY_DEFAULT;
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent?key=' + API_KEY;
  var payload = JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] });
  var resp = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', payload: payload, muteHttpExceptions: true });
  var data = JSON.parse(resp.getContentText());
  return data.candidates[0].content.parts[0].text;
}
```
Todos los módulos deben usar _callGemini() — NO reimplementar la llamada HTTP.

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
**Estado 'ACTIVO' / 'SI'** — ambos valores válidos en col L de PERSONAL para usuario activo.
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

## 9. ANTI-PATRONES DETECTADOS (no reproducir)

| Anti-patrón | Dónde aparece | Correcto |
|---|---|---|
| `getDropDownarray_v2()` — duplicar con sufijo _v2 | CheckCode.js | Modificar función original |
| `console.log` en backend GAS | DesvioscCode.js (11 instancias) | Usar Logger.log() |
| `GmailApp.sendEmail()` mezclado con MailApp | DesvioscCode.js | Usar MailApp |
| 3 patrones distintos de llamada a Gemini | Varios archivos | Usar _callGemini() |
| Variables globales como cache | Varios archivos | Usar CacheService |
| IDs hardcodeados en múltiples archivos | 9 archivos .js | Solo en Code.js |
| `appendRow()` en loop | Varios | Batch setValues() |
| `google.script.run` sin withFailureHandler | 40 archivos HTML | Siempre incluir |
| Redefine _norm() local | NotificacionesCode.js, CodeMapa.js | Usar global de Code.js |
| Mezcla Bootstrap 4/5.1/5.3 | Varios HTML | Usar solo 5.3.x |

---

## 10. CHECKEOS DE VALIDACIÓN

Antes de hacer commit, verificar:

```bash
# 1. No hay google.script.run sin withFailureHandler nuevo
grep -rn 'google.script.run' /home/user/SST/*.html | grep -v 'withFailureHandler' | grep -v '//'

# 2. No hay console.log en archivos .js backend
grep -rn 'console\.log' /home/user/SST/*.js

# 3. No hay IDs de Sheets hardcodeados fuera de Code.js
grep -rn '"1[A-Za-z0-9_-]\{40,\}"' /home/user/SST/*.js | grep -v Code.js | grep -v RolCode.js

# 4. No hay appendRow en loop
grep -B5 'appendRow' /home/user/SST/*.js | grep -E 'for|forEach|map|while'

# 5. Toda función nueva retorna JSON.stringify
# (revisión manual del diff)
```

---

## 11. LECCIONES APRENDIDAS — DECISIONES CRÍTICAS DE ARQUITECTURA

### L1 · HTML parciales NO procesan template directives GAS

**Problema:** Se intentó dividir `ReportesLaboral.html` en `rlCss.html` + `rlScript.html` usando `<?!= include('rlCss') ?>` para evitar timeouts al subir archivos grandes.

**Lo que pasó:** Las directivas `<?!=...?>` solo se procesan en el archivo raíz cargado via `HtmlService.createTemplateFromFile()` (actualmente `index.html`). Los parciales incluidos con `include()` se sirven como HTML estático — las directivas aparecen literalmente en el DOM.

**Regla:** Los archivos HTML de módulos (Check.html, MovimEpp.html, ReportesLaboral.html, etc.) son **parciales estáticos**. No usar `<?!=...?>` en ellos. Todo CSS y JS debe estar autocontenido en el mismo archivo.

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

*Fin de CLAUDE.md — Actualizar después de cada cambio estructural.*
