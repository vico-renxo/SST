// ============================================================
//  IpercCode.js — Modulo IPERC Linea Base (backend)
//
//  Puente entre Iperc.html y el Google Sheet + Gemini IA.
//  TODAS las funciones estan namespaced con prefijo "iperc" para
//  evitar colisiones con otros modulos del sistema (Check, PASSO,
//  Rol, etc.).
//
//  NO define doGet — eso ya vive en Code.js.
// ============================================================

// ─────────────────────────────────────────
//  CONFIGURACION — puedes cambiar el IPERC_SS_ID_DEFAULT
//  o configurarlo via Script Properties como "IPERC_SS_ID".
// ─────────────────────────────────────────
const IPERC_SS_ID_DEFAULT = "1ANw0WcZiDYfZhDTmEazUeBtx7_MyhG9UOnVTJNx-Xjo"; // opcional: ID del Spreadsheet donde vive la hoja DATOS

const IPERC_SHEET_DATOS   = "DATOS";
const IPERC_FILA_INICIO   = 6;

const IPERC_COL_PROCESO   = 2;   // B
const IPERC_COL_AREA      = 3;   // C
const IPERC_COL_TAREA     = 4;   // D
const IPERC_COL_PUESTO    = 5;   // E
const IPERC_COL_RUTINARIO = 6;   // F
const IPERC_COL_PELIGROS  = 7;   // G
const IPERC_COL_RIESGO    = 8;   // H
const IPERC_COL_PROB      = 9;   // I
const IPERC_COL_SEV       = 10;  // J
// K(11) Score inicial (formula), L(12) Nivel inicial (formula)
const IPERC_COL_ELIMINAC  = 13;  // M
const IPERC_COL_SUSTIT    = 14;  // N
const IPERC_COL_ING       = 15;  // O
const IPERC_COL_ADM       = 16;  // P
const IPERC_COL_EPP       = 17;  // Q

// ─────────────────────────────────────────
//  SPREADSHEET ACCESS
// ─────────────────────────────────────────
function ipercGetSpreadsheet_() {
  try {
    const props = PropertiesService.getScriptProperties();
    const cfgId = props.getProperty("IPERC_SS_ID") || IPERC_SS_ID_DEFAULT;
    if (cfgId) return SpreadsheetApp.openById(cfgId);
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    Logger.log("ipercGetSpreadsheet_: " + e);
    return null;
  }
}

// ─────────────────────────────────────────
//  MATRIZ CONFIG POR DEFECTO (FCX 4x4)
// ─────────────────────────────────────────
function ipercMatrizFCXDefault_() {
  return {
    nombre: "Matriz FCX 4x4", size: 4, maxProb: 4, maxSev: 4,
    probs: [
      { v:1, l:"Improbable",   d:"Muy improbable durante la vida de la operacion" },
      { v:2, l:"Posible",      d:"Puede ocurrir durante la vida de la operacion" },
      { v:3, l:"Probable",     d:"Puede ocurrir menos de una vez al anio" },
      { v:4, l:"Casi Seguro",  d:"Evento recurrente o mas de una vez al anio" }
    ],
    sevs: [
      { v:1, l:"Menor",         d:"Lesion minima o primeros auxilios" },
      { v:2, l:"Moderado",      d:"Tratamiento medico o labores restringidas" },
      { v:3, l:"Significativo", d:"Fatalidades o discapacidades permanentes" },
      { v:4, l:"Catastrofico",  d:"Fatalidades multiples" }
    ],
    nivs: [
      { l:"BAJO",    de:1,  ha:2,  c:"#3d9e3d", t:"#ffffff" },
      { l:"MEDIO",   de:3,  ha:4,  c:"#d4a017", t:"#ffffff" },
      { l:"ALTO",    de:6,  ha:8,  c:"#e05c00", t:"#ffffff" },
      { l:"CRITICO", de:9,  ha:16, c:"#cc0000", t:"#ffffff" }
    ]
  };
}

function ipercLeerMatrizConfig() {
  const ss = ipercGetSpreadsheet_();
  if (!ss) return ipercMatrizFCXDefault_();
  const sheet = ss.getSheetByName("Configuracion Matriz");
  if (!sheet) return ipercMatrizFCXDefault_();
  try {
    const maxProb = Number(sheet.getRange("D3").getValue()) || 4;
    const maxSev  = Number(sheet.getRange("F3").getValue()) || 4;
    const probs = [];
    for (let i = 0; i < maxProb; i++) {
      const r = sheet.getRange(5 + i, 1, 1, 3).getValues()[0];
      probs.push({ v: Number(r[0]) || (i+1), l: String(r[1] || "Nivel "+(i+1)), d: String(r[2] || "") });
    }
    const sevs = [];
    for (let j = 0; j < maxSev; j++) {
      const r2 = sheet.getRange(5 + j, 5, 1, 3).getValues()[0];
      sevs.push({ v: Number(r2[0]) || (j+1), l: String(r2[1] || "Nivel "+(j+1)), d: String(r2[2] || "") });
    }
    const nivs = [];
    for (let k = 5; k <= 12; k++) {
      const r3 = sheet.getRange(k, 9, 1, 5).getValues()[0];
      if (!r3[0]) break;
      nivs.push({ l: String(r3[0]), de: Number(r3[1]) || 1, ha: Number(r3[2]) || 4, c: String(r3[3] || "#3d9e3d"), t: String(r3[4] || "#ffffff") });
    }
    const cfg = ipercMatrizFCXDefault_();
    return {
      nombre:  String(sheet.getRange("B3").getValue() || cfg.nombre),
      size:    Math.max(maxProb, maxSev),
      maxProb: maxProb,
      maxSev:  maxSev,
      probs: probs.length > 0 ? probs : cfg.probs,
      sevs:  sevs.length  > 0 ? sevs  : cfg.sevs,
      nivs:  nivs.length  > 0 ? nivs  : cfg.nivs
    };
  } catch(e) {
    Logger.log("ipercLeerMatrizConfig: " + e);
    return ipercMatrizFCXDefault_();
  }
}

function ipercLeerMatrizConfigJSON() {
  return JSON.stringify(ipercLeerMatrizConfig());
}

// ─────────────────────────────────────────
//  HELPERS DE NIVEL / FORMULAS
// ─────────────────────────────────────────
function ipercCalcularNivelByScore_(score, cfg) {
  const nivs = (cfg && cfg.nivs) || [];
  for (let i = 0; i < nivs.length; i++) {
    const n = nivs[i];
    if (score >= n.de && score <= n.ha) {
      return { label: n.l, color: n.c, colorTexto: n.t };
    }
  }
  if (nivs.length) {
    const last = nivs[nivs.length - 1];
    return { label: last.l, color: last.c, colorTexto: last.t };
  }
  return null;
}

function ipercBuildFormulaNivel_(fila, col, cfg) {
  const nivs = (cfg && cfg.nivs) || [];
  if (!nivs.length) return '=""';
  const cell = col + fila;
  let formula = '=IF(' + cell + '="","",';
  let closures = 1;
  for (let i = 0; i < nivs.length; i++) {
    const n = nivs[i];
    if (i === nivs.length - 1) {
      formula += '"' + n.l + '"';
    } else {
      formula += 'IF(AND(' + cell + '>=' + n.de + ',' + cell + '<=' + n.ha + '),"' + n.l + '",';
      closures++;
    }
  }
  for (let k = 0; k < closures; k++) formula += ')';
  return formula;
}

// ─────────────────────────────────────────
//  PUENTE HTML <-> SHEET
// ─────────────────────────────────────────

/** Guarda todas las filas IPERC al Sheet */
function ipercGuardarFilasAlSheet(filasJSON) {
  const ss = ipercGetSpreadsheet_();
  if (!ss) return { ok: false, error: "Configura IPERC_SS_ID en Script Properties" };
  const sheet = ss.getSheetByName(IPERC_SHEET_DATOS);
  if (!sheet) return { ok: false, error: "Hoja " + IPERC_SHEET_DATOS + " no encontrada" };

  const filas = JSON.parse(filasJSON);
  const cfg   = ipercLeerMatrizConfig();

  const ultimaFila = Math.max(sheet.getLastRow(), IPERC_FILA_INICIO + filas.length - 1);
  if (ultimaFila >= IPERC_FILA_INICIO) {
    sheet.getRange(IPERC_FILA_INICIO, 1, ultimaFila - IPERC_FILA_INICIO + 1, 24).clearContent();
    sheet.getRange(IPERC_FILA_INICIO, 1, ultimaFila - IPERC_FILA_INICIO + 1, 24).setBackground("#FFFFFF");
  }

  filas.forEach((f, i) => {
    if (!f.proceso) return;
    const fila = IPERC_FILA_INICIO + i;

    sheet.getRange(fila, 1).setValue(i + 1);
    sheet.getRange(fila, IPERC_COL_PROCESO  ).setValue(f.proceso   || "");
    sheet.getRange(fila, IPERC_COL_AREA     ).setValue(f.area      || "");
    sheet.getRange(fila, IPERC_COL_TAREA    ).setValue(f.tarea     || "");
    sheet.getRange(fila, IPERC_COL_PUESTO   ).setValue(f.puesto    || "");
    sheet.getRange(fila, IPERC_COL_RUTINARIO).setValue(f.rnr       || "R");

    if (f.peligros) {
      sheet.getRange(fila, IPERC_COL_PELIGROS).setValue(f.peligros    || "");
      sheet.getRange(fila, IPERC_COL_RIESGO  ).setValue(f.riesgo      || "");
      sheet.getRange(fila, IPERC_COL_PROB    ).setValue(Number(f.prob) || "");
      sheet.getRange(fila, IPERC_COL_SEV     ).setValue(Number(f.sev)  || "");
      sheet.getRange(fila, IPERC_COL_ELIMINAC).setValue(f.eliminacion  || "");
      sheet.getRange(fila, IPERC_COL_SUSTIT  ).setValue(f.sustitucion  || "");
      sheet.getRange(fila, IPERC_COL_ING     ).setValue(f.ing          || "");
      sheet.getRange(fila, IPERC_COL_ADM     ).setValue(f.adm          || "");
      sheet.getRange(fila, IPERC_COL_EPP     ).setValue(f.epp          || "");
      sheet.getRange(fila, IPERC_COL_PELIGROS, 1, IPERC_COL_EPP - IPERC_COL_PELIGROS + 1).setBackground("#FFFACD");

      sheet.getRange(fila, 11).setFormula(`=IF(I${fila}*J${fila}=0,"",I${fila}*J${fila})`);
      sheet.getRange(fila, 12).setFormula(ipercBuildFormulaNivel_(fila, "K", cfg));
      if (f.prob && f.sev) {
        const nv = ipercCalcularNivelByScore_(Number(f.prob) * Number(f.sev), cfg);
        if (nv) sheet.getRange(fila, 12).setBackground(nv.color).setFontColor(nv.colorTexto).setFontWeight("bold").setHorizontalAlignment("center");
      }
    }

    if (f.probRes) sheet.getRange(fila, 18).setValue(Number(f.probRes) || "");
    if (f.sevRes)  sheet.getRange(fila, 19).setValue(Number(f.sevRes)  || "");
    sheet.getRange(fila, 20).setFormula(`=IF(R${fila}*S${fila}=0,"",R${fila}*S${fila})`);
    sheet.getRange(fila, 21).setFormula(ipercBuildFormulaNivel_(fila, "T", cfg));
    if (f.accion)      sheet.getRange(fila, 22).setValue(f.accion      || "");
    if (f.responsable) sheet.getRange(fila, 23).setValue(f.responsable || "");

    sheet.getRange(fila, 1, 1, 23).setWrap(true).setVerticalAlignment("top");
    sheet.setRowHeight(fila, 85);
  });

  return { ok: true, filas: filas.filter(f => f.proceso).length };
}

/** Lee filas del Sheet y las devuelve al HTML */
function ipercLeerFilasDelSheet() {
  const ss = ipercGetSpreadsheet_();
  if (!ss) return JSON.stringify([]);
  const sheet = ss.getSheetByName(IPERC_SHEET_DATOS);
  if (!sheet) return JSON.stringify([]);

  const uf = sheet.getLastRow();
  if (uf < IPERC_FILA_INICIO) return JSON.stringify([]);

  const datos = sheet.getRange(IPERC_FILA_INICIO, 1, uf - IPERC_FILA_INICIO + 1, 24).getValues();
  const filas = datos
    .filter(r => r[IPERC_COL_PROCESO - 1])
    .map(r => ({
      proceso:    r[IPERC_COL_PROCESO - 1],
      area:       r[IPERC_COL_AREA    - 1],
      tarea:      r[IPERC_COL_TAREA   - 1],
      puesto:     r[IPERC_COL_PUESTO  - 1],
      rnr:        r[IPERC_COL_RUTINARIO - 1],
      peligros:   r[IPERC_COL_PELIGROS - 1],
      riesgo:     r[IPERC_COL_RIESGO  - 1],
      prob:       r[IPERC_COL_PROB    - 1],
      sev:        r[IPERC_COL_SEV     - 1],
      eliminacion:r[IPERC_COL_ELIMINAC - 1],
      sustitucion:r[IPERC_COL_SUSTIT  - 1],
      ing:        r[IPERC_COL_ING     - 1],
      adm:        r[IPERC_COL_ADM     - 1],
      epp:        r[IPERC_COL_EPP     - 1],
      probRes:    r[17],
      sevRes:     r[18],
      accion:     r[21],
      responsable:r[22]
    }));

  return JSON.stringify(filas);
}

// ─────────────────────────────────────────
//  CONFIG EMPRESA + API KEY (Script Properties)
// ─────────────────────────────────────────
function ipercGuardarConfigEmpresa(cfgJSON) {
  PropertiesService.getScriptProperties().setProperty("IPERC_CONFIG_EMPRESA", cfgJSON);
  return { ok: true };
}

function ipercLeerConfigEmpresa() {
  return PropertiesService.getScriptProperties().getProperty("IPERC_CONFIG_EMPRESA") || "{}";
}

function ipercObtenerAPIKey_() {
  // Usa la clave global del sistema definida en Code.js
  if (typeof API_KEY !== 'undefined' && API_KEY && API_KEY.trim().length > 20) {
    return API_KEY.trim();
  }
  return null;
}

// ─────────────────────────────────────────
//  SANITIZADOR JSON ROBUSTO
//  Recupera JSON de respuestas Gemini que pueden venir:
//   - con fences markdown (```json ... ```)
//   - truncadas (hit maxOutputTokens) con string/objeto sin cerrar
//   - con saltos de linea / tabs literales dentro de strings (ilegal en JSON)
//   - con caracteres de control (< 0x20)
//  Parser char-by-char que normaliza y cierra lo que falta.
// ─────────────────────────────────────────
function ipercSanitizarJSON_(raw) {
  if (!raw) return "{}";
  var s = String(raw).replace(/```json/gi, "").replace(/```/g, "").trim();

  var start = s.indexOf("{");
  if (start === -1) return "{}";
  s = s.substring(start);

  var out = "";
  var inStr = false;
  var esc = false;
  var depth = 0;
  var finished = false;

  for (var i = 0; i < s.length && !finished; i++) {
    var c = s.charAt(i);
    var code = s.charCodeAt(i);

    if (esc) {
      // Char despues de backslash: siempre se conserva
      out += c;
      esc = false;
      continue;
    }

    if (c === "\\") {
      out += c;
      esc = true;
      continue;
    }

    if (c === "\"") {
      out += c;
      inStr = !inStr;
      continue;
    }

    if (inStr) {
      // Dentro de string: escapar/normalizar chars de control invalidos
      if (code < 0x20) {
        if (c === "\n" || c === "\r" || c === "\t") {
          out += " ";
        }
        // otros control chars: omitir
        continue;
      }
      out += c;
    } else {
      // Fuera de string: contar llaves para cerrar el objeto cuando toque
      if (c === "{") {
        depth++;
        out += c;
      } else if (c === "}") {
        depth--;
        out += c;
        if (depth === 0) {
          finished = true;
        }
      } else {
        out += c;
      }
    }
  }

  // Si la string quedo abierta (truncado), cerrarla
  if (inStr) {
    if (esc) out += "\\"; // backslash colgante
    out += "\"";
  }

  // Si hay llaves abiertas (truncado), cerrarlas.
  // Quitar coma final colgante antes de cerrar.
  while (depth > 0) {
    out = out.replace(/,\s*$/, "");
    out += "}";
    depth--;
  }

  return out;
}

// ─────────────────────────────────────────
//  ANALISIS CON GEMINI — AUTOCONTENIDO
// ─────────────────────────────────────────
function ipercAnalizarFilaConIA(filaJSON) {
  const apiKey = ipercObtenerAPIKey_();
  if (!apiKey) return JSON.stringify({ error: "API Key no configurada en el sistema (Code.js)." });

  const f = JSON.parse(filaJSON);
  const modelo = PropertiesService.getScriptProperties().getProperty("IPERC_GEMINI_MODELO") || "gemini-2.5-flash";
  const url = "https://generativelanguage.googleapis.com/v1beta/models/" + modelo + ":generateContent?key=" + apiKey;

  const tipo = (f.rnr === "NR") ? "No Rutinario" : "Rutinario";
  const escP = f.escalaProb || "1=Improbable, 2=Posible, 3=Probable, 4=Casi Seguro";
  const escS = f.escalaSev  || "1=Menor, 2=Moderado, 3=Significativo, 4=Catastrofico";
  const escN = f.escalaNivs || "1-2=BAJO, 3-4=MEDIO, 6-8=ALTO, 9-16=CRITICO";
  const mP   = f.maxProb || 4;
  const mS   = f.maxSev  || 4;

  const prompt =
    "Eres un experto SSOMA en mineria en Peru. Normas: DS-024-2016-EM, ISO 45001:2018, NIOSH.\n" +
    "TAREA:\n- Proceso: " + f.proceso + "\n- Area: " + f.area +
    "\n- Tarea: " + f.tarea + "\n- Puesto: " + f.puesto + "\n- Tipo: " + tipo + "\n\n" +
    "PROBABILIDAD (entero 1 al " + mP + "): " + escP + "\n" +
    "CONSECUENCIA (entero 1 al " + mS + "): " + escS + "\n" +
    "NIVELES (ProbxCons): " + escN + "\n\n" +
    "REGLAS ESTRICTAS:\n" +
    "1. Responde SOLO con JSON valido. Sin texto extra, sin backticks, sin markdown.\n" +
    "2. En listas usa ' | ' como separador (NO saltos de linea).\n" +
    "3. probabilidad = entero 1-" + mP + ", severidad = entero 1-" + mS + ".\n" +
    "4. prob_residual y sev_residual deben dar producto MENOR al inicial.\n" +
    "5. Terminologia tecnica minera peruana.\n" +
    "JSON a devolver:\n" +
    '{"peligros":"1. Peligro A | 2. Peligro B | 3. Peligro C",' +
    '"riesgo":"1. Riesgo A | 2. Riesgo B",' +
    '"probabilidad":2,"prob_justificacion":"razon breve",' +
    '"severidad":3,"sev_justificacion":"razon breve",' +
    '"eliminacion":"medida o N/A","sustitucion":"medida o N/A",' +
    '"ing_controles":"1. Control 1 | 2. Control 2",' +
    '"adm_controles":"1. PETS | 2. Capacitacion | 3. Supervision",' +
    '"epp":"1. Casco | 2. Lentes | 3. Guantes | 4. Botas punta acero",' +
    '"prob_residual":1,"sev_residual":2,' +
    '"accion_mejora":"accion concreta y medible"}';

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.1,
      maxOutputTokens: 2048,
      responseMimeType: "application/json"
    }
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    const body = response.getContentText();

    if (code !== 200) {
      let errMsg = "HTTP " + code;
      try {
        const errObj = JSON.parse(body);
        errMsg = errObj.error && errObj.error.message ? errObj.error.message.substring(0, 120) : errMsg;
      } catch(_) {}
      return JSON.stringify({ error: errMsg });
    }

    const data = JSON.parse(body);
    const content = data.candidates &&
                    data.candidates[0] &&
                    data.candidates[0].content &&
                    data.candidates[0].content.parts &&
                    data.candidates[0].content.parts[0] &&
                    data.candidates[0].content.parts[0].text;

    if (!content) {
      const reason = (data.candidates && data.candidates[0] && data.candidates[0].finishReason) || "desconocida";
      return JSON.stringify({ error: "Sin contenido. Razon: " + reason });
    }

    const limpio = ipercSanitizarJSON_(content);
    try {
      const resultado = JSON.parse(limpio);
      return JSON.stringify(resultado);
    } catch (ePrimario) {
      // Segundo intento: quitar el ultimo par clave-valor incompleto y reparar
      const limpio2 = ipercRepararJSONFinal_(limpio);
      try {
        const resultado2 = JSON.parse(limpio2);
        return JSON.stringify(resultado2);
      } catch (eSecundario) {
        Logger.log("ipercAnalizarFilaConIA: parse fail. Raw: " + String(content).substring(0, 500));
        Logger.log("ipercAnalizarFilaConIA: limpio: " + limpio.substring(0, 500));
        return JSON.stringify({ error: "IA devolvio JSON invalido. Intenta de nuevo o reduce la fila." });
      }
    }

  } catch(e) {
    return JSON.stringify({ error: "Error servidor: " + e.toString().substring(0, 100) });
  }
}

// Reparador agresivo: remueve pares clave-valor incompletos al final
function ipercRepararJSONFinal_(s) {
  if (!s) return "{}";
  var t = String(s);

  // Quitar la llave final para trabajar sobre el contenido
  var cerradas = 0;
  while (t.length && t.charAt(t.length - 1) === "}") {
    t = t.substring(0, t.length - 1);
    cerradas++;
  }

  // Remover patrones colgantes al final, iterativamente
  var prev;
  do {
    prev = t;
    t = t.replace(/[,\s]+$/, "");                        // coma/whitespace final
    t = t.replace(/"[^"]*"\s*:\s*[^,}\]"]*$/, "");       // clave:valor_incompleto
    t = t.replace(/"[^"]*"\s*:\s*$/, "");                // clave: (sin valor)
    t = t.replace(/,\s*"[^"\n]*"\s*$/, "");              // ,"clave" (sin colon)
    t = t.replace(/[,\s]+$/, "");                        // coma/whitespace final
  } while (t !== prev);

  // Restaurar llaves
  for (var i = 0; i < Math.max(cerradas, 1); i++) t += "}";
  return t;
}