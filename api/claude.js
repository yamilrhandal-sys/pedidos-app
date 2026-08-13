// ============================================================
// PROXY A LA API DE CLAUDE
// ------------------------------------------------------------
// La clave de Anthropic NUNCA debe viajar al navegador: cualquiera
// podría verla en las herramientas de desarrollo y gastar la cuota.
// Esta función corre en el servidor de Vercel, guarda la clave como
// variable de entorno y reenvía la petición.
//
// Variables de entorno necesarias en Vercel:
//   ANTHROPIC_API_KEY       — clave de console.anthropic.com (secreta)
//   VITE_SUPABASE_URL       — la misma que ya usa la app
//   VITE_SUPABASE_ANON_KEY  — la misma que ya usa la app
// ============================================================

// Modelos permitidos. Cualquier otro se rechaza, para que nadie
// pueda usar el proxy con un modelo más caro del previsto.
const MODELOS_PERMITIDOS = new Set([
  'claude-haiku-4-5-20251001',
  'claude-sonnet-4-6',
]);

const MAX_TOKENS_TOPE = 4000;
const TAMANO_MAXIMO_PETICION = 12 * 1024 * 1024; // 12 MB (las proformas van en base64)

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Solo se acepta POST' });
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    console.error('Falta ANTHROPIC_API_KEY en las variables de entorno');
    return res.status(500).json({ error: 'El servidor no tiene configurada la clave de Claude.' });
  }

  // ---------- 1. Verificar que quien llama tenga sesión en Supabase ----------
  // Sin esto el proxy quedaría abierto a cualquiera que descubra la URL.
  const authHeader = req.headers.authorization || '';
  const token = authHeader.startsWith('Bearer ') ? authHeader.slice(7) : '';
  if (!token) {
    return res.status(401).json({ error: 'Falta la sesión. Inicia sesión de nuevo.' });
  }

  const supabaseUrl = process.env.VITE_SUPABASE_URL;
  const supabaseAnonKey = process.env.VITE_SUPABASE_ANON_KEY;
  if (!supabaseUrl || !supabaseAnonKey) {
    console.error('Faltan las variables de Supabase en el servidor');
    return res.status(500).json({ error: 'El servidor no puede verificar la sesión.' });
  }

  try {
    const verificacion = await fetch(`${supabaseUrl}/auth/v1/user`, {
      headers: { Authorization: `Bearer ${token}`, apikey: supabaseAnonKey },
    });
    if (!verificacion.ok) {
      return res.status(401).json({ error: 'Sesión inválida o expirada. Inicia sesión de nuevo.' });
    }
  } catch (err) {
    console.error('No se pudo verificar la sesión:', err);
    return res.status(503).json({ error: 'No se pudo verificar la sesión. Intenta de nuevo.' });
  }

  // ---------- 2. Revisar lo que se pide ----------
  const cuerpo = req.body;
  if (!cuerpo || typeof cuerpo !== 'object') {
    return res.status(400).json({ error: 'Petición mal formada.' });
  }

  if (JSON.stringify(cuerpo).length > TAMANO_MAXIMO_PETICION) {
    return res.status(413).json({ error: 'El archivo es demasiado grande. Usa uno más liviano.' });
  }

  if (!MODELOS_PERMITIDOS.has(cuerpo.model)) {
    return res.status(400).json({ error: `Modelo no permitido: ${cuerpo.model}` });
  }

  if (!Array.isArray(cuerpo.messages) || cuerpo.messages.length === 0) {
    return res.status(400).json({ error: 'Faltan los mensajes de la petición.' });
  }

  // ---------- 3. Reenviar a Anthropic ----------
  try {
    const respuesta = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
      },
      body: JSON.stringify({
        model: cuerpo.model,
        max_tokens: Math.min(Number(cuerpo.max_tokens) || 1000, MAX_TOKENS_TOPE),
        messages: cuerpo.messages,
        ...(cuerpo.system ? { system: cuerpo.system } : {}),
      }),
    });

    const datos = await respuesta.json();

    if (!respuesta.ok) {
      // El detalle del error se registra en el servidor; al usuario le llega
      // un mensaje corto para no exponer información interna.
      console.error('Error de la API de Claude:', respuesta.status, datos);
      const mensaje =
        respuesta.status === 429
          ? 'Demasiadas peticiones seguidas. Espera un momento e intenta de nuevo.'
          : respuesta.status === 401
            ? 'La clave de Claude configurada en el servidor no es válida.'
            : 'Claude no pudo procesar la petición. Intenta de nuevo.';
      return res.status(respuesta.status).json({ error: mensaje });
    }

    return res.status(200).json(datos);
  } catch (err) {
    console.error('Fallo al contactar a Claude:', err);
    return res.status(503).json({ error: 'No se pudo contactar a Claude. Revisa tu conexión.' });
  }
}

// Permitir cuerpos grandes: las proformas viajan como imagen/PDF en base64
export const config = {
  api: {
    bodyParser: { sizeLimit: '12mb' },
  },
};
