/**
 * ASISTENTE SSOMA-OP
 * Cloudflare Worker con Workers AI (Llama 3.1) — acceso a todos los módulos del sistema
 *
 * Variables de entorno requeridas (wrangler secret put):
 *   GAS_URL  → URL de tu GAS Web App
 *
 * Bindings requeridos en wrangler.toml:
 *   [[ai]]            binding = "AI"
 *   [[kv_namespaces]] binding = "CACHE"  id = "..."
 */

const SYSTEM_PROMPT = `Eres el Asistente SSOMA-OP, asistente inteligente del sistema integral de gestión de seguridad y salud ocupacional (SSOMA).

MÓDULOS QUE CONOCES Y PUEDES CONSULTAR:

1. CAPACITACIONES (B DATOS, Matriz, TEMAS, REGISTRO FIRMAS)
   - Registros por trabajador: DNI, Nombre, Cargo, Empresa, Tema, Área, Puntaje, Fecha, Estado (Aprobado/Reprobado/Pendiente), Capacitador
   - Temas activos/vigentes, puntaje mínimo, ciclos anuales
   - REGISTRO FIRMAS: solo aprobados con fecha de habilitación y firma URL

2. PERSONAL (hoja PERSONAL + INFO ADICIONAL)
   - Todos los trabajadores: DNI, Nombre, Empresa, Cargo, Condición (activo/inactivo)
   - INFO ADICIONAL: EMO, vacunas (Tétano, Hepatitis B, Neumococo, Influenza, COVID-19, Carnet de Sanidad)

3. EPP — Equipos de Protección Personal
   - STOCK: inventario por almacén/producto/variante, stock actual vs mínimo, alertas
   - REGISTRO: entregas a trabajadores (DNI, producto, fechas, vencimiento, estado)
   - MATRIZ: requerimientos de EPP por cargo

4. EVENTOS / ACCIDENTES
   - Incidentes y accidentes: tipo, trabajador, empresa, lugar, días de descanso, estado

5. DESVÍOS / INSPECCIONES
   - Observaciones de seguridad: equipo, clasificación, responsable, estado, proceso

6. CHECK LIST / EVALUACIONES
   - Evaluaciones con ítems 1-5 o Si/No
   - Clasificación automática: Bueno (≥70%), Regular (50-69%), Deficiente (<50%)

7. IPERC — Identificación de Peligros y Evaluación de Riesgos
   - Matriz de riesgos por proceso/área/tarea/puesto, nivel inicial y residual, controles

REGLAS:
- Responde SIEMPRE en español, claro y conciso
- Usa los datos reales del sistema, nunca inventes cifras
- Máximo 3-4 oraciones (será leído en voz alta)
- Si hay un DNI de 8 dígitos en la pregunta, úsalo en la herramienta correspondiente`;

async function llamarGAS(toolName, args, gasUrl) {
  const url  = `${gasUrl}?action=${toolName}&params=${encodeURIComponent(JSON.stringify(args))}`;
  const resp = await fetch(url, { cf: { cacheTtl: 30 } });
  if (!resp.ok) throw new Error(`GAS error ${resp.status}`);
  return await resp.json();
}

async function resolverTool(toolName, args, env) {
  const cacheKey = `${toolName}:${JSON.stringify(args)}`;

  if (env.CACHE) {
    const cached = await env.CACHE.get(cacheKey);
    if (cached) return JSON.parse(cached);
  }

  const resultado = await llamarGAS(toolName, args, env.GAS_URL);

  if (env.CACHE) {
    await env.CACHE.put(cacheKey, JSON.stringify(resultado), { expirationTtl: 60 });
  }

  return resultado;
}

function _extraerPalabra(texto, regex) {
  const m = texto.match(regex);
  return m ? String(m[1] || "").trim() : "";
}

function _decidirHerramienta(pregunta) {
  const q = pregunta.toLowerCase();
  const dniMatch = pregunta.match(/\b\d{8}\b/);

  // ── EPP Stock ─────────────────────────────────────────────────────────────
  if (/stock|inventario epp|reposici[oó]n|almac[eé]n.*epp|epp.*stock|epp.*bajo|epp.*m[ií]nimo|cuántos? (cascos?|guantes?|chalecos?|lentes?|botas?)/.test(q))
    return { tool: "consultar_epp_stock", args: {
      soloAlerta: /bajo|m[ií]nimo|alerta|reponer|reposici[oó]n/.test(q) ? "true" : "false",
      producto: _extraerPalabra(q, /(?:de|del)\s+(casco|guante|chaleco|lente|bota|epp[\w\s]*?)(?:\s+en|\s+del?|$)/i)
    }};

  // ── EPP Registro/Entregas ─────────────────────────────────────────────────
  if (/entrega.*epp|epp.*entrega|recibió.*epp|epp.*trabajador|vencimiento.*epp|epp.*vencer|próximo.*vencer|caducar|devoluci[oó]n/.test(q))
    return { tool: "consultar_epp_registro", args: {
      dni: dniMatch ? dniMatch[0] : "",
      proxVencer: /vencer|vencimiento|pr[oó]ximo|caducar/.test(q) ? "true" : "false"
    }};

  // ── Vacunas / EMO ─────────────────────────────────────────────────────────
  if (/vacuna|emo|examen m[eé]dico|hepatitis|t[eé]tano|influenza|neumococo|covid|carnet.*sanidad|sanidad/.test(q))
    return { tool: "consultar_vacunas_emo", args: {
      dni:    dniMatch ? dniMatch[0] : "",
      vacuna: _extraerPalabra(q, /(emo|hepatitis|t[eé]tano|influenza|neumococo|covid|carnet)/i),
      soloSin: /no tiene|sin vacuna|falta|pendiente|sin emo|no vacunado/.test(q) ? "true" : "false"
    }};

  // ── Accidentes / Eventos ──────────────────────────────────────────────────
  if (/accidente|incidente|evento.*seguridad|lesi[oó]n|herido|d[ií]as? de descanso|descanso m[eé]dico|incapacidad/.test(q))
    return { tool: "consultar_eventos", args: {
      dni:  dniMatch ? dniMatch[0] : "",
      tipo: _extraerPalabra(q, /(?:accidente|incidente|evento)\s+(?:de\s+)?(\w+)/i)
    }};

  // ── Desvíos / Inspecciones ────────────────────────────────────────────────
  if (/desv[ií]o|desviaci[oó]n|inspecci[oó]n|observaci[oó]n|hallazgo|amonest|no conform|incumplimiento/.test(q))
    return { tool: "consultar_desvios", args: {
      estado:        /pendiente|abierto/.test(q) ? "pendiente" : "",
      clasificacion: /cr[ií]tico|mayor|menor/.test(q) ? _extraerPalabra(q, /(cr[ií]tico|mayor|menor)/i) : ""
    }};

  // ── Checklist / Evaluaciones ──────────────────────────────────────────────
  if (/checklist|check list|evaluaci[oó]n|evaluado|puntaje checklist|rendimiento|desempe[ñn]o|equipo evaluad/.test(q))
    return { tool: "consultar_checklist", args: {
      equipo:   _extraerPalabra(q, /(?:checklist|evaluaci[oó]n|equipo)\s+(?:de\s+)?([a-záéíóúñ\s]+?)(?:\s+de|\s+del|\s+en|$)/i),
      evaluado: ""
    }};

  // ── IPERC ─────────────────────────────────────────────────────────────────
  if (/iperc|peligro|nivel de riesgo|matriz de riesgo|riesgo.*[aá]rea|[aá]rea.*riesgo|severidad.*riesgo/.test(q))
    return { tool: "consultar_iperc", args: {
      area:  _extraerPalabra(q, /(?:[aá]rea|zona|proceso)\s+(?:de\s+)?([a-záéíóúñ\s]+?)(?:\s+|$)/i),
      nivel: _extraerPalabra(q, /(cr[ií]tico|alto|medio|bajo)/i)
    }};

  // ── Personal ──────────────────────────────────────────────────────────────
  if (/personal|trabajador(?:es)?|empleado(?:s)?|activos?|inactivos?|cu[aá]ntos? trabaj|n[oó]mina/.test(q)) {
    if (dniMatch) return { tool: "consultar_personal", args: { dni: dniMatch[0] } };
    return { tool: "consultar_personal", args: {
      estado:  /activ/.test(q) ? "activo" : /inactiv/.test(q) ? "inactivo" : "",
      empresa: _extraerPalabra(q, /empresa\s+([a-záéíóúñ\s]+?)(?:\s+|$)/i),
      cargo:   _extraerPalabra(q, /cargo\s+([a-záéíóúñ\s]+?)(?:\s+|$)/i)
    }};
  }

  // ── Capacitaciones ───────────────────────────────────────────────────────
  if (dniMatch)
    return { tool: "estado_trabajador", args: { dni: dniMatch[0] } };

  if (/activ|vigente|habilit|disponible|ahora mismo|en este momento/.test(q))
    return { tool: "temas_activos", args: {} };

  if (/cumplimiento|estad[ií]stica|porcentaje|cu[aá]ntos? aprobaron|resumen.*capacit|cu[aá]ntos? hay/.test(q))
    return { tool: "resumen_cumplimiento", args: { tema: "" } };

  if (/registro firmas|firma url|pdf virtual|firmas/.test(q))
    return { tool: "registro_firmas", args: { tema: "" } };

  if (/existe.*tema|tiene examen|puntaje m[ií]nimo|capacitador del tema/.test(q)) {
    const temaM = q.match(/(?:tema|curso)\s+["']?([a-záéíóúñ\s]+)["']?/i);
    return { tool: "verificar_tema", args: { tema: temaM ? temaM[1].trim() : "" } };
  }

  if (/qu[eé] temas|cu[aá]les temas|qu[eé] cursos|cu[aá]les cursos/.test(q))
    return { tool: "temas_activos", args: {} };

  if (/registros|capacitaciones|aprobado|reprobado|pendiente/.test(q))
    return { tool: "buscar_registros", args: { query: pregunta } };

  return { tool: null, args: {} };
}

export default {
  async fetch(request, env) {
    const corsHeaders = {
      "Access-Control-Allow-Origin":  "*",
      "Access-Control-Allow-Methods": "POST, OPTIONS, GET",
      "Access-Control-Allow-Headers": "Content-Type"
    };

    if (request.method === "OPTIONS")
      return new Response(null, { headers: corsHeaders });

    const pathname = new URL(request.url).pathname;

    if (pathname === "/debug") {
      return Response.json({
        ai:     typeof env.AI    !== "undefined" ? "✅ OK" : "❌ FALTA — agrega binding Workers AI con nombre 'AI'",
        cache:  typeof env.CACHE !== "undefined" ? "✅ OK" : "⚠️ Sin caché (opcional)",
        gasUrl: env.GAS_URL ? "✅ OK" : "❌ FALTA — agrega secret GAS_URL"
      }, { headers: corsHeaders });
    }

    if (pathname === "/ping") {
      if (env.GAS_URL) {
        try { await fetch(`${env.GAS_URL}?action=ping&params={}`); } catch {}
      }
      return Response.json({ ok: true }, { headers: corsHeaders });
    }

    if (pathname === "/test-gas") {
      try {
        const url   = `${env.GAS_URL}?action=temas_activos&params=%7B%7D`;
        const resp  = await fetch(url);
        const texto = await resp.text();
        let json = null;
        try { json = JSON.parse(texto); } catch {}
        return Response.json({ status: resp.status, ok: resp.ok, textoRaw: texto.substring(0, 500), jsonParsed: json }, { headers: corsHeaders });
      } catch(e) {
        return Response.json({ error: e.message }, { headers: corsHeaders });
      }
    }

    if (request.method !== "POST")
      return Response.json({ error: "Método no permitido" }, { status: 405, headers: corsHeaders });

    if (!env.AI)
      return Response.json({ error: "Workers AI no configurado. En Cloudflare: Bindings → Add binding → Workers AI → nombre 'AI'" }, { status: 500, headers: corsHeaders });

    if (!env.GAS_URL)
      return Response.json({ error: "GAS_URL no configurado. En Cloudflare: Settings → Variables and Secrets → Add Secret → GAS_URL" }, { status: 500, headers: corsHeaders });

    let pregunta = "";
    try {
      const body = await request.json();
      pregunta   = String(body.pregunta || "").trim();
    } catch {
      return Response.json({ error: "Body JSON inválido" }, { status: 400, headers: corsHeaders });
    }

    if (!pregunta)
      return Response.json({ error: "Pregunta vacía" }, { status: 400, headers: corsHeaders });

    try {
      const { tool, args } = _decidirHerramienta(pregunta);

      let datosReales = null;
      if (tool) {
        try {
          datosReales = await resolverTool(tool, args, env);
        } catch (e) {
          datosReales = { error: e.message };
        }
      }

      const contexto = datosReales
        ? `\n\nDATOS REALES DEL SISTEMA:\n${JSON.stringify(datosReales, null, 2)}\n\nResponde basándote en estos datos. Sé conciso (máximo 3-4 oraciones).`
        : "";

      const respuesta = await env.AI.run("@cf/meta/llama-3.1-8b-instruct", {
        messages: [
          { role: "system", content: SYSTEM_PROMPT },
          { role: "user",   content: pregunta + contexto }
        ],
        max_tokens: 450
      });

      return Response.json({ respuesta: respuesta.response || "No pude generar una respuesta." }, { headers: corsHeaders });

    } catch (err) {
      return Response.json({ error: "Error interno: " + err.message }, { status: 500, headers: corsHeaders });
    }
  }
};
