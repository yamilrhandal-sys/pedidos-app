import React, { useState, useEffect, useCallback, useRef } from 'react';
import {
  Plus, Trash2, FileDown, Package, Truck, ClipboardList, X, ChevronRight,
  ChevronDown, ChevronLeft, Search, Camera, ImageOff, AlertCircle, Settings, Layers, Tag, Pencil, Globe, Upload, Wallet, BarChart3, Copy,
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { createClient } from '@supabase/supabase-js';

// ============================================================
// SHIM DE ALMACENAMIENTO — Puente window.storage -> localStorage
// ------------------------------------------------------------
// Dentro de los artifacts de Claude existe window.storage.
// Fuera (Vercel, navegador normal) NO existe, y la app crashea.
// Este bloque crea window.storage usando localStorage del navegador.
//
// PENDIENTE: al migrar a Supabase, reemplazar SOLO este bloque
// por llamadas a supabase.from('...') — el resto del código
// no necesita ningún cambio.
// ============================================================
// ------------------------------------------------------------
// Storage HÍBRIDO:
//   • compartido = true  -> Supabase, tabla app_state (todo el equipo ve lo mismo)
//   • compartido = false -> localStorage (sesión y preferencias de ESTE dispositivo)
// La API pública (get/set/delete/list) es idéntica a la anterior, así que
// el resto del código no cambia.
// ------------------------------------------------------------
if (typeof window !== 'undefined' && !window.storage) {
  // Usa el MISMO cliente que el resto de la app (definido más abajo como `supabase`).
  // Se lee de forma perezosa: estas funciones sólo corren después de que el
  // módulo terminó de evaluarse, así que `supabase` ya existe.
  const cliente = () => { try { return supabase; } catch { return null; } };

  const prefijoLocal = 'personal:';

  window.storage = {
    async get(clave, compartido = false) {
      const _sb = cliente();
      if (compartido && _sb) {
        const { data, error } = await _sb
          .from('app_state').select('value').eq('key', clave).maybeSingle();
        if (error) { console.error('storage.get', error); return null; }
        if (!data) return null;
        // value es JSONB; lo devolvemos como string para que JSON.parse funcione igual
        return { key: clave, value: JSON.stringify(data.value), shared: true };
      }
      const bruto = localStorage.getItem(prefijoLocal + clave);
      if (bruto === null) return null;
      return { key: clave, value: bruto, shared: false };
    },
    async set(clave, valor, compartido = false) {
      const _sb = cliente();
      if (compartido && _sb) {
        // valor llega como string JSON; lo parseamos para guardarlo como JSONB
        let parsed;
        try { parsed = JSON.parse(valor); } catch { parsed = valor; }
        const { error } = await _sb
          .from('app_state')
          .upsert({ key: clave, value: parsed, updated_at: new Date().toISOString() });
        // IMPORTANTE: lanzar el error para que quien llama pueda reintentar.
        // Si solo se registrara en consola, el dato se perdería en silencio.
        if (error) {
          console.error('storage.set', clave, error);
          throw error;
        }
        return { key: clave, value: valor, shared: true };
      }
      localStorage.setItem(prefijoLocal + clave, valor);
      return { key: clave, value: valor, shared: false };
    },
    async delete(clave, compartido = false) {
      const _sb = cliente();
      if (compartido && _sb) {
        const { error } = await _sb.from('app_state').delete().eq('key', clave);
        if (error) { console.error('storage.delete', clave, error); throw error; }
        return { key: clave, deleted: true, shared: true };
      }
      localStorage.removeItem(prefijoLocal + clave);
      return { key: clave, deleted: true, shared: false };
    },
    async list(inicio = '', compartido = false) {
      const _sb = cliente();
      if (compartido && _sb) {
        const { data, error } = await _sb
          .from('app_state').select('key').like('key', inicio + '%');
        if (error) { console.error('storage.list', error); return { keys: [], prefix: inicio, shared: true }; }
        return { keys: (data || []).map((r) => r.key), prefix: inicio, shared: true };
      }
      const claves = [];
      for (let i = 0; i < localStorage.length; i++) {
        const k = localStorage.key(i);
        if (k && k.startsWith(prefijoLocal + inicio)) claves.push(k.slice(prefijoLocal.length));
      }
      return { keys: claves, prefix: inicio, shared: false };
    },
  };
}


// ---------- Storage helpers ----------
// Descarga un Excel de forma explícita (Blob + enlace). Es más confiable dentro del
// iframe de la app que XLSX.writeFile, que puede fallar en silencio.
// Devuelve true si logró disparar la descarga.
function descargarLibro(wb, nombreArchivo) {
  try {
    const binario = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([binario], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = nombreArchivo;
    a.rel = 'noopener';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 2000);
    return true;
  } catch (e) {
    // Último recurso: el método propio de la librería
    try {
      XLSX.writeFile(wb, nombreArchivo);
      return true;
    } catch (e2) {
      return false;
    }
  }
}

// Atajo: arma el libro desde un arreglo de filas y lo descarga
function descargarExcel(filas, nombreHoja, nombreArchivo, cols) {
  const ws = XLSX.utils.json_to_sheet(filas);
  if (cols) ws['!cols'] = cols;
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, nombreHoja);
  return descargarLibro(wb, nombreArchivo);
}

// Botón para descargar una plantilla de Excel
function BotonesPlantilla({ filas, nombreHoja, nombreArchivo, cols }) {
  return (
    <button
      onClick={() => descargarExcel(filas, nombreHoja, nombreArchivo, cols)}
      className="w-full py-2.5 rounded-lg border border-app-line text-xs text-app-light flex items-center justify-center gap-1.5 active:bg-app-active"
    >
      <FileDown size={14} /> Descargar plantilla
    </button>
  );
}

const KEYS = {
  products: 'pedidos:products',
  suppliers: 'pedidos:suppliers',
  orders: 'pedidos:orders',
  departamentos: 'pedidos:departamentos',
  tipos: 'pedidos:tipos',
  marcas: 'pedidos:marcas',
  empresa: 'pedidos:empresa',   // { logo: dataURL }
  embarcadores: 'pedidos:embarcadores', // [{ id, nombre, contacto, telefono, email }]
  marcasProveedores: 'pedidos:marcasProveedores', // { "Polo Club": [idProveedor, ...] }
  tasaCambio: 'pedidos:tasaCambio',
  factores: 'pedidos:factoresImportacion',
  ciudades: 'pedidos:ciudades',
  fabricas: 'pedidos:fabricas',
  presupuestos: 'pedidos:presupuestos',
  usuarios: 'pedidos:usuarios',
  borradores: 'pedidos:borradores',
};
const SESION_KEY = 'pedidos:sesion'; // storage PERSONAL: recuerda el usuario de este dispositivo
const ULTIMO_USUARIO_KEY = 'pedidos:ultimo'; // storage PERSONAL: último usuario que entró (sobrevive a logout)
const MODO_KEY = 'pedidos:modo';     // storage PERSONAL: 'auto' | 'movil' — preferencia de vista

async function loadShared(key, fallback) {
  try {
    const res = await window.storage.get(key, true);
    if (!res) return fallback;
    const valor = JSON.parse(res.value);
    // Si lo guardado es null/undefined, o cambió de forma respecto al
    // respaldo (p. ej. se esperaba una lista y hay un objeto), usar el respaldo.
    if (valor === null || valor === undefined) return fallback;
    if (Array.isArray(fallback) && !Array.isArray(valor)) return fallback;
    return valor;
  } catch {
    return fallback;
  }
}
async function saveShared(key, value) {
  try {
    await window.storage.set(key, JSON.stringify(value), true);
  } catch (e) {
    console.error('Storage error', e);
  }
}
// Storage personal (por dispositivo/cuenta, no compartido con el equipo)
async function loadPersonal(key, fallback) {
  try {
    const res = await window.storage.get(key, false);
    return res ? JSON.parse(res.value) : fallback;
  } catch {
    return fallback;
  }
}
async function savePersonal(key, value) {
  try {
    await window.storage.set(key, JSON.stringify(value), false);
  } catch (e) {
    console.error('Storage error', e);
  }
}
async function deletePersonal(key) {
  try {
    await window.storage.delete(key, false);
  } catch (e) { /* puede no existir */ }
}

// Hash del PIN (SHA-256) — nunca se guarda el PIN en texto plano
async function hashPin(pin) {
  const data = new TextEncoder().encode('pedidos-app-salt::' + pin);
  const buf = await crypto.subtle.digest('SHA-256', data);
  return Array.from(new Uint8Array(buf)).map((b) => b.toString(16).padStart(2, '0')).join('');
}

const MONEDAS_COSTO = ['RMB', 'USD', 'HNL'];
const SIMBOLO = { RMB: '¥', USD: '$', HNL: 'L' };
const MULTIPLICADORES = { RMB: [12, 13, 14, 15, 16, 17, 18], USD: [72, 76, 79, 83] };

// ---------- Orígenes de compra ----------
const ORIGENES = [
  { id: 'china',    label: 'China',            emoji: '🇨🇳', moneda: 'RMB', color: '#e85d5d' },
  { id: 'usa',      label: 'Estados Unidos',   emoji: '🇺🇸', moneda: 'USD', color: '#5db8e8' },
  { id: 'panama',   label: 'Panamá',           emoji: '🇵🇦', moneda: 'USD', color: '#5de89a' },
  { id: 'honduras', label: 'Honduras',         emoji: '🇭🇳', moneda: 'HNL', color: '#d4a574' },
];

// ---------- Destinos del pedido ----------
const DESTINOS = [
  { id: 'H', label: 'Honduras', emoji: '🇭🇳' },
  { id: 'G', label: 'Afiliada', emoji: '🇬🇹' },
];
const destinoInfo = (id) => DESTINOS.find((d) => d.id === id) || DESTINOS[0];

// Muestra el código con la letra de destino. La terminación H/G es EXCLUSIVA de China
// (porque en China creamos el código nosotros y necesitamos diferenciar destinos).
// Para USA, Panamá o cualquier item SIN origen definido: el código queda tal cual.
const codigoConDestino = (it) => {
  if (it.origen !== 'china') return it.codigo;
  if (!it.destino) return it.codigo;
  const partes = (it.codigo || '').split('-');
  const base = partes[0];
  if (/\d[HG]$/.test(base)) {
    partes[0] = base.replace(/[HG]$/, it.destino);
    return partes.join('-');
  }
  return `${it.codigo}-${it.destino}`;
};

// Normaliza un código: sin espacios ni guiones repetidos
const sanitizarCodigo = (str) => (str || '').trim().replace(/\s+/g, '-').replace(/-+/g, '-');

// ---------- Presupuestos ----------
// Clave de un presupuesto: "H|ROPA DE CABALLEROS|Camisa"
const presuKey = (destino, depto, tipo) => `${destino}|${depto}|${tipo}`;

function calcularGastoPresupuestos(orders, rate) {
  const gasto = {}; // key -> { usd, piezas }
  (orders || []).forEach((o) => {
    (o.items || []).forEach((it) => {
      const destino = it.destino || 'H';
      const depto = it.departamento || '(Sin depto)';
      const tipo = it.tipo || '(Sin tipo)';
      const k = presuKey(destino, depto, tipo);
      const piezas = sumVariantes(it.variantes);
      const montoUsd = (it.costoMoneda === 'RMB' ? (piezas * it.costoMonto) / rate : piezas * it.costoMonto);
      if (!gasto[k]) gasto[k] = { usd: 0, piezas: 0 };
      gasto[k].usd += montoUsd;
      gasto[k].piezas += piezas;
    });
  });
  return gasto;
}

function Presupuestos({ presupuestos, setPresupuestos, departamentos = [], tipos = [], orders = [], tasaCambio, onBack }) {
  const [destino, setDestino] = useState('H');
  const [deptoSel, setDeptoSel] = useState('');
  const [tipoSel, setTipoSel] = useState('');
  const [openId, setOpenId] = useState(null);
  const [importResult, setImportResult] = useState(null);

  const rate = parseFloat(tasaCambio?.rmbUsd) || 7.25;
  const gasto = calcularGastoPresupuestos(orders, rate);

  const setLimite = (campo, valor) => {
    if (!deptoSel || !tipoSel) return;
    const k = presuKey(destino, deptoSel, tipoSel);
    const actual = presupuestos[k] || { monto: 0, piezas: 0 };
    setPresupuestos({ ...presupuestos, [k]: { ...actual, [campo]: parseFloat(valor) || 0 } });
  };

  const eliminarPresu = (k) => {
    const next = { ...presupuestos };
    delete next[k];
    setPresupuestos(next);
  };

  // Filas de ejemplo de la plantilla de presupuestos
  const FILAS_PLANTILLA = [
    { Destino: 'Honduras', Departamento: departamentos[0] || 'ROPA DE CABALLEROS', Tipo: (tipos[0]?.nombre) || 'Jeans', 'Monto USD': 10000, Piezas: 2000 },
    { Destino: 'Afiliada', Departamento: departamentos[0] || 'ROPA DE CABALLEROS', Tipo: (tipos[0]?.nombre) || 'Jeans', 'Monto USD': 5000, Piezas: 1200 },
  ];

  // Importar presupuestos desde Excel
  const importarExcel = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setImportResult(null);
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws);
        const next = { ...presupuestos };
        let cargados = 0, errores = 0;

        rows.forEach((row) => {
          // Aceptar variaciones de encabezados
          const destinoRaw = String(row.Destino || row.destino || row.País || row.Pais || '').trim().toLowerCase();
          const depto = String(row.Departamento || row.departamento || row.Depto || '').trim();
          const tipo = String(row.Tipo || row.tipo || '').trim();
          const monto = parseFloat(row['Monto USD'] || row.Monto || row.monto || row.USD || 0) || 0;
          const piezas = parseInt(row.Piezas || row.piezas || row.Cantidad || 0) || 0;

          // Normalizar destino a H / G
          let destId = '';
          if (destinoRaw.startsWith('h') || destinoRaw.includes('hondur')) destId = 'H';
          else if (destinoRaw.startsWith('g') || destinoRaw.includes('guatemal')) destId = 'G';

          if (!destId || !depto || !tipo || (monto === 0 && piezas === 0)) {
            errores++;
            return;
          }
          next[presuKey(destId, depto, tipo)] = { monto, piezas };
          cargados++;
        });

        setPresupuestos(next);
        setImportResult({ cargados, errores, total: rows.length });
      } catch (err) {
        setImportResult({ error: 'No se pudo leer el archivo. Verifica que sea un Excel válido.' });
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  // Presupuestos existentes del destino activo
  const delDestino = Object.entries(presupuestos || {})
    .filter(([k]) => k.startsWith(`${destino}|`))
    .map(([k, v]) => {
      const [, depto, tipo] = k.split('|');
      return { k, depto, tipo, ...v, gastado: gasto[k] || { usd: 0, piezas: 0 } };
    });

  // Agrupar por departamento
  const porDepto = {};
  delDestino.forEach((p) => {
    if (!porDepto[p.depto]) porDepto[p.depto] = [];
    porDepto[p.depto].push(p);
  });

  const kActual = deptoSel && tipoSel ? presuKey(destino, deptoSel, tipoSel) : null;
  const presuActual = kActual ? presupuestos[kActual] : null;

  return (
    <div className="space-y-5">
      <button onClick={onBack} className="text-sm text-app-dim2 flex items-center gap-1">← Volver</button>

      <div>
        <h1 className="text-lg font-bold text-app-white mb-1">Presupuestos</h1>
        <p className="text-xs text-app-dim2">Define límites de dinero y piezas por Destino → Departamento → Tipo. El gasto se acumula de todos los pedidos.</p>
      </div>

      {/* Importar desde Excel */}
      <div className="bg-app-blue border border-app-line2 rounded-xl p-3 space-y-2">
        <p className="text-sm font-medium flex items-center gap-2"><Upload size={15} className="text-app-sky" /> Cargar presupuestos desde Excel</p>
        <p className="text-xs text-app-dim2">Columnas: Destino (Honduras/Afiliada), Departamento, Tipo, Monto USD, Piezas.</p>
        <BotonesPlantilla
          filas={FILAS_PLANTILLA}
          nombreHoja="Plantilla"
          nombreArchivo="plantilla-presupuestos.xlsx"
          cols={[{ wch: 14 }, { wch: 26 }, { wch: 20 }, { wch: 12 }, { wch: 10 }]}
        />
        <div className="flex gap-2">
          <label className="w-full py-2 rounded-lg bg-app-sky text-app-bg text-xs font-semibold flex items-center justify-center gap-1.5 cursor-pointer active:opacity-80">
            <Upload size={13} /> Subir Excel
            <input type="file" accept=".xlsx,.xls" onChange={importarExcel} className="hidden" />
          </label>
        </div>
        {importResult && (
          importResult.error ? (
            <div className="flex items-center gap-2 bg-app-redbg text-app-red2 text-xs rounded-lg px-3 py-2">
              <AlertCircle size={14} /> {importResult.error}
            </div>
          ) : (
            <div className="bg-app-bg border border-app-line rounded-lg px-3 py-2 text-xs">
              <span className="text-app-green font-semibold">✓ {importResult.cargados} cargados</span>
              {importResult.errores > 0 && <span className="text-app-dim2"> · {importResult.errores} con datos incompletos (omitidos)</span>}
            </div>
          )
        )}
      </div>

      {/* Selector de destino */}
      <div className="flex gap-2">
        {DESTINOS.map((d) => (
          <button
            key={d.id}
            onClick={() => setDestino(d.id)}
            className={`flex-1 py-2.5 rounded-xl text-sm font-medium border transition flex items-center justify-center gap-1.5 ${
              destino === d.id ? 'bg-app-gold text-app-bg border-app-gold' : 'bg-app-panel border-app-line text-app-dim2'
            }`}
          >
            <span>{d.emoji}</span> {d.label}
          </button>
        ))}
      </div>

      {/* Formulario para definir un presupuesto */}
      <div className="bg-app-panel border border-app-line rounded-xl shadow-app p-4 space-y-3">
        <p className="text-sm font-medium">Definir presupuesto para {destinoInfo(destino).emoji} {destinoInfo(destino).label}</p>
        <SearchableSelect id="presu-depto" openId={openId} setOpenId={setOpenId}
          value={deptoSel} onChange={(v) => { setDeptoSel(v); setTipoSel(''); }}
          options={departamentos} placeholder="Elige departamento…" recentKey="departamento" />
        {deptoSel && (
          <SearchableSelect id="presu-tipo" openId={openId} setOpenId={setOpenId}
            value={tipoSel} onChange={setTipoSel}
            options={tipos.map((t) => t.nombre)} placeholder="Elige tipo de producto…" recentKey="tipo" />
        )}
        {deptoSel && tipoSel && (
          <div className="flex gap-2">
            <div className="flex-1">
              <label className="text-xs text-app-dim2 uppercase tracking-wide mb-1 block">Límite USD</label>
              <input type="number" min={0} placeholder="0"
                value={presuActual?.monto || ''}
                onChange={(e) => setLimite('monto', e.target.value)}
                className="w-full bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm" />
            </div>
            <div className="flex-1">
              <label className="text-xs text-app-dim2 uppercase tracking-wide mb-1 block">Límite piezas</label>
              <input type="number" min={0} placeholder="0"
                value={presuActual?.piezas || ''}
                onChange={(e) => setLimite('piezas', e.target.value)}
                className="w-full bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm" />
            </div>
          </div>
        )}
      </div>

      {/* Lista jerárquica de presupuestos del destino */}
      {Object.keys(porDepto).length === 0 ? (
        <p className="text-center text-xs text-app-dim py-6">Sin presupuestos definidos para {destinoInfo(destino).label}.</p>
      ) : (
        <div className="space-y-4">
          {Object.entries(porDepto).sort((a, b) => a[0].localeCompare(b[0])).map(([depto, lista]) => (
            <div key={depto}>
              <div className="flex items-center gap-2 mb-2">
                <Layers size={14} className="text-app-gold shrink-0" />
                <span className="text-sm font-semibold text-app-white">{depto}</span>
              </div>
              <div className="space-y-2">
                {lista.sort((a, b) => a.tipo.localeCompare(b.tipo)).map((p) => {
                  const pctUsd = p.monto > 0 ? (p.gastado.usd / p.monto) * 100 : 0;
                  const pctPz = p.piezas > 0 ? (p.gastado.piezas / p.piezas) * 100 : 0;
                  const colorUsd = pctUsd > 100 ? '#e85d5d' : pctUsd >= 80 ? '#e8a33d' : '#5de89a';
                  const colorPz = pctPz > 100 ? '#e85d5d' : pctPz >= 80 ? '#e8a33d' : '#5de89a';
                  return (
                    <div key={p.k} className="bg-app-panel border border-app-line rounded-xl p-3">
                      <div className="flex items-center justify-between mb-2">
                        <span className="text-sm text-app-white flex items-center gap-1.5"><Tag size={12} className="text-app-dim2" /> {p.tipo}</span>
                        <BotonBorrar onConfirm={() => eliminarPresu(p.k)} size={14} />
                      </div>
                      {p.monto > 0 && (
                        <div className="mb-2">
                          <div className="flex justify-between text-xs mb-1">
                            <span className="text-app-dim2">Dinero</span>
                            <span className={p.gastado.usd > p.monto ? 'text-app-red2 font-semibold' : 'text-app-light'}>
                              {fmtMoneda(p.gastado.usd, 'USD')} / {fmtMoneda(p.monto, 'USD')}
                            </span>
                          </div>
                          <div className="h-1.5 rounded-full bg-app-bg overflow-hidden">
                            <div className="h-full rounded-full" style={{ width: `${Math.min(pctUsd, 100)}%`, backgroundColor: colorUsd }} />
                          </div>
                        </div>
                      )}
                      {p.piezas > 0 && (
                        <div>
                          <div className="flex justify-between text-xs mb-1">
                            <span className="text-app-dim2">Piezas</span>
                            <span className={p.gastado.piezas > p.piezas ? 'text-app-red2 font-semibold' : 'text-app-light'}>
                              {p.gastado.piezas.toLocaleString()} / {p.piezas.toLocaleString()}
                            </span>
                          </div>
                          <div className="h-1.5 rounded-full bg-app-bg overflow-hidden">
                            <div className="h-full rounded-full" style={{ width: `${Math.min(pctPz, 100)}%`, backgroundColor: colorPz }} />
                          </div>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}


// ---------- Login: usuario + PIN ----------
// ---------- Borrado seguro: pide confirmación en dos toques ----------
// Evita perder datos por un toque accidental. Si no se confirma en 5 s, se cancela solo.
function BotonBorrar({ onConfirm, size = 14, texto = null, icono = 'trash', aviso = '' }) {
  const [confirmando, setConfirmando] = useState(false);

  useEffect(() => {
    if (!confirmando) return;
    const t = setTimeout(() => setConfirmando(false), 5000);
    return () => clearTimeout(t);
  }, [confirmando]);

  if (confirmando) {
    return (
      <span className="flex items-center gap-1.5 shrink-0" onClick={(e) => e.stopPropagation()}>
        {aviso && <span className="text-xs text-app-gold whitespace-nowrap">{aviso}</span>}
        <button
          onClick={(e) => { e.stopPropagation(); setConfirmando(false); }}
          className="text-xs text-app-dim2 border border-app-line rounded-lg px-2 py-1"
        >
          Cancelar
        </button>
        <button
          onClick={(e) => { e.stopPropagation(); setConfirmando(false); onConfirm(); }}
          className="text-xs font-semibold text-app-red2 bg-app-redbg border border-app-line rounded-lg px-2 py-1"
        >
          Borrar
        </button>
      </span>
    );
  }

  const Icono = icono === 'x' ? X : Trash2;
  return (
    <button
      onClick={(e) => { e.stopPropagation(); setConfirmando(true); }}
      className={texto
        ? 'text-xs text-app-red2 flex items-center gap-1 shrink-0'
        : 'text-app-dim active:text-app-red shrink-0'}
      title="Borrar"
    >
      <Icono size={size} />{texto && <span>{texto}</span>}
    </button>
  );
}

// ============================================================
// CLIENTE DE SUPABASE
// ------------------------------------------------------------
// Vite solo expone variables que empiezan con VITE_.
// En Vercel deben llamarse: VITE_SUPABASE_URL y VITE_SUPABASE_ANON_KEY
// ============================================================
const SUPABASE_URL = import.meta.env.VITE_SUPABASE_URL;
const SUPABASE_ANON_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY;
const SUPABASE_CONFIGURADO = !!(SUPABASE_URL && SUPABASE_ANON_KEY);

const supabase = SUPABASE_CONFIGURADO
  ? createClient(SUPABASE_URL, SUPABASE_ANON_KEY)
  : null;

// ------------------------------------------------------------
// FOTOS — Supabase Storage (bucket "productos")
// De cada foto se guardan dos versiones:
//   • original  -> tal cual se tomó, sin recomprimir
//   • miniatura -> 400px de ancho, para listas y tarjetas
// En el producto queda: { url, thumb, path, pathThumb }
// ------------------------------------------------------------
const BUCKET_FOTOS = 'productos';
const ANCHO_MINIATURA = 400;

// Genera una miniatura JPEG a partir del archivo original
function crearMiniatura(file, ancho = ANCHO_MINIATURA) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => {
      URL.revokeObjectURL(url);
      const escala = Math.min(1, ancho / img.width);
      const w = Math.round(img.width * escala);
      const h = Math.round(img.height * escala);
      const canvas = document.createElement('canvas');
      canvas.width = w; canvas.height = h;
      canvas.getContext('2d').drawImage(img, 0, 0, w, h);
      canvas.toBlob(
        (blob) => (blob ? resolve(blob) : reject(new Error('No se pudo crear la miniatura'))),
        'image/jpeg',
        0.8
      );
    };
    img.onerror = () => { URL.revokeObjectURL(url); reject(new Error('Imagen no válida')); };
    img.src = url;
  });
}

// Sube una foto (original + miniatura) y devuelve sus URLs
async function subirFoto(file, codigoProducto = 'sin-codigo') {
  if (!supabase) throw new Error('Supabase no está configurado');

  const ext = (file.name.split('.').pop() || 'jpg').toLowerCase();
  const sello = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
  const carpeta = `${codigoProducto}`.replace(/[^a-zA-Z0-9_-]/g, '_') || 'sin-codigo';
  const pathOriginal = `${carpeta}/${sello}.${ext}`;
  const pathThumb = `${carpeta}/${sello}_thumb.jpg`;

  const { error: errOrig } = await supabase.storage
    .from(BUCKET_FOTOS)
    .upload(pathOriginal, file, { contentType: file.type, upsert: false });
  if (errOrig) throw errOrig;

  let urlThumb = null;
  try {
    const thumb = await crearMiniatura(file);
    const { error: errThumb } = await supabase.storage
      .from(BUCKET_FOTOS)
      .upload(pathThumb, thumb, { contentType: 'image/jpeg', upsert: false });
    if (!errThumb) {
      urlThumb = supabase.storage.from(BUCKET_FOTOS).getPublicUrl(pathThumb).data.publicUrl;
    }
  } catch { /* si falla la miniatura, se usa la original */ }

  const urlOriginal = supabase.storage.from(BUCKET_FOTOS).getPublicUrl(pathOriginal).data.publicUrl;

  return {
    url: urlOriginal,
    thumb: urlThumb || urlOriginal,
    path: pathOriginal,
    pathThumb: urlThumb ? pathThumb : null,
  };
}

// Borra una foto y su miniatura del almacenamiento
async function borrarFoto(foto) {
  if (!supabase || !foto?.path) return;
  const rutas = [foto.path, foto.pathThumb].filter(Boolean);
  try { await supabase.storage.from(BUCKET_FOTOS).remove(rutas); } catch { /* ignorar */ }
}

// ------------------------------------------------------------
// AUDITORÍA — registra quién creó, editó o borró qué y cuándo.
// Compara la lista anterior contra la nueva (por id) y escribe
// una fila por cada elemento agregado, cambiado o quitado.
// No hace falta tocar cada pantalla: basta con llamarla desde
// las funciones persist* con el antes y el después.
// ------------------------------------------------------------
async function registrarAuditoria(tabla, anterior, nuevo, actor, etiqueta) {
  if (!supabase) return;
  try {
    const antes = new Map((anterior || []).map((x) => [x.id, x]));
    const despues = new Map((nuevo || []).map((x) => [x.id, x]));
    const filas = [];

    for (const [id, item] of despues) {
      if (!antes.has(id)) {
        filas.push({ accion: 'crear', resumen: etiqueta(item), detalle: { despues: item } });
      } else if (JSON.stringify(antes.get(id)) !== JSON.stringify(item)) {
        filas.push({ accion: 'editar', resumen: etiqueta(item), detalle: { antes: antes.get(id), despues: item } });
      }
    }
    for (const [id, item] of antes) {
      if (!despues.has(id)) {
        filas.push({ accion: 'borrar', resumen: etiqueta(item), detalle: { antes: item } });
      }
    }
    if (!filas.length || filas.length > 200) return; // evita ruido en cargas masivas de catálogos

    await supabase.from('auditoria').insert(
      filas.map((f) => ({
        tabla,
        accion: f.accion,
        resumen: f.resumen,
        detalle: f.detalle,
        usuario_email: actor?.email || null,
        usuario_nombre: actor?.nombre || null,
      }))
    );
  } catch (e) {
    console.error('auditoria', e); // nunca debe romper el guardado por esto
  }
}

// Variante para valores que no son listas (p. ej. factores de costo, datos de empresa)
async function registrarAuditoriaObjeto(tabla, anterior, nuevo, actor, resumen) {
  if (!supabase) return;
  if (JSON.stringify(anterior) === JSON.stringify(nuevo)) return;
  try {
    await supabase.from('auditoria').insert([{
      tabla, accion: 'editar', resumen,
      detalle: { antes: anterior, despues: nuevo },
      usuario_email: actor?.email || null,
      usuario_nombre: actor?.nombre || null,
    }]);
  } catch (e) {
    console.error('auditoria', e);
  }
}
function urlFoto(foto, preferirMiniatura = true) {
  if (!foto) return null;
  if (typeof foto === 'string') return foto;
  return (preferirMiniatura ? foto.thumb : foto.url) || foto.url || null;
}

// Normaliza el campo de fotos de un producto a un arreglo
function listaFotos(producto) {
  if (!producto) return [];
  if (Array.isArray(producto.fotos)) return producto.fotos;
  if (producto.foto) return [producto.foto];
  return [];
}

// Pantalla de aviso si faltan las variables de entorno
function PantallaSinConfigurar() {
  return (
    <div className="min-h-screen bg-app-bg flex flex-col items-center justify-center px-6 font-sans">
      <style>{APP_STYLES}</style>
      <div className="w-full max-w-md bg-app-panel border border-app-line rounded-app-lg p-8 space-y-3 text-center">
        <p className="text-4xl">⚙️</p>
        <h1 className="text-lg font-bold text-app-white">Falta configurar Supabase</h1>
        <p className="text-sm text-app-dim2">
          En Vercel → Settings → Environment Variables hay que agregar:
        </p>
        <div className="bg-app-bg rounded-app p-3 text-left text-xs text-app-gold font-mono">
          VITE_SUPABASE_URL<br />VITE_SUPABASE_ANON_KEY
        </div>
        <p className="text-xs text-app-dim">
          Después hay que hacer Redeploy para que tomen efecto.
        </p>
      </div>
    </div>
  );
}

// ============================================================
// PANTALLA DE LOGIN — Supabase Auth (correo + contraseña)
// ------------------------------------------------------------
// Reemplaza el login por PIN. La contraseña vive cifrada en el
// servidor de Supabase; la app nunca la ve ni la guarda.
// ============================================================
function PantallaLogin() {
  const [correo, setCorreo] = useState('');
  const [clave, setClave] = useState('');
  const [verClave, setVerClave] = useState(false);
  const [error, setError] = useState('');
  const [aviso, setAviso] = useState('');
  const [ocupado, setOcupado] = useState(false);
  const [modoReset, setModoReset] = useState(false);

  const correoValido = (v) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test((v || '').trim());

  const entrar = async () => {
    setError(''); setAviso('');
    if (!correoValido(correo)) return setError('Escribe un correo válido.');
    if (!clave) return setError('Escribe tu contraseña.');
    setOcupado(true);
    try {
      const { error: err } = await supabase.auth.signInWithPassword({
        email: correo.trim().toLowerCase(),
        password: clave,
      });
      if (err) {
        const m = (err.message || '').toLowerCase();
        if (m.includes('invalid login')) setError('Correo o contraseña incorrectos.');
        else if (m.includes('not confirmed')) setError('Tu correo aún no está confirmado. Revisa tu bandeja de entrada.');
        else setError(err.message);
      }
    } catch (e) {
      setError('No se pudo conectar. Revisa tu conexión.');
    } finally {
      setOcupado(false);
    }
  };

  const enviarReset = async () => {
    setError(''); setAviso('');
    if (!correoValido(correo)) return setError('Escribe un correo válido.');
    setOcupado(true);
    try {
      const { error: err } = await supabase.auth.resetPasswordForEmail(correo.trim().toLowerCase());
      if (err) setError(err.message);
      else setAviso('Te enviamos un correo para restablecer tu contraseña.');
    } catch (e) {
      setError('No se pudo enviar el correo.');
    } finally {
      setOcupado(false);
    }
  };

  return (
    <div className="min-h-screen bg-app-bg flex flex-col items-center justify-center px-6 font-sans">
      <style>{APP_STYLES}</style>
      <div className="w-full max-w-sm bg-app-panel border border-app-line rounded-app-lg p-8 space-y-5">
        <div className="text-center">
          <p className="text-4xl mb-2">🛍️</p>
          <h1 className="text-2xl font-bold text-app-white">Pedidos</h1>
          <p className="text-xs text-app-dim mt-1">
            {modoReset ? 'Te enviaremos un correo para restablecerla' : 'Ingresa con tu correo y contraseña'}
          </p>
        </div>

        {error && (
          <div className="flex items-start gap-2 bg-app-redbg text-app-red2 text-xs rounded-app px-3 py-2">
            <AlertCircle size={14} className="shrink-0 mt-0.5" /> <span>{error}</span>
          </div>
        )}
        {aviso && (
          <div className="bg-app-green text-app-green text-xs rounded-app px-3 py-2">{aviso}</div>
        )}

        <div className="space-y-3">
          <input
            type="email"
            inputMode="email"
            autoComplete="username"
            placeholder="tu.correo@carrion.hn"
            value={correo}
            onChange={(e) => setCorreo(e.target.value)}
            disabled={ocupado}
            className="w-full bg-app-bg border border-app-line rounded-app px-4 py-3 text-sm text-app-white"
          />

          {!modoReset && (
            <div className="relative">
              <input
                type={verClave ? 'text' : 'password'}
                autoComplete="current-password"
                placeholder="Contraseña"
                value={clave}
                onChange={(e) => setClave(e.target.value)}
                onKeyDown={(e) => { if (e.key === 'Enter') entrar(); }}
                disabled={ocupado}
                className="w-full bg-app-bg border border-app-line rounded-app px-4 py-3 pr-16 text-sm text-app-white"
              />
              <button
                type="button"
                onClick={() => setVerClave((v) => !v)}
                className="absolute right-3 top-1/2 -translate-y-1/2 text-xs text-app-dim2"
              >
                {verClave ? 'Ocultar' : 'Ver'}
              </button>
            </div>
          )}

          <button
            onClick={modoReset ? enviarReset : entrar}
            disabled={ocupado}
            className="w-full py-3 rounded-app bg-app-gold text-app-bg font-semibold text-sm disabled:opacity-50"
          >
            {ocupado ? 'Un momento…' : (modoReset ? 'Enviar correo' : 'Entrar')}
          </button>
        </div>

        <button
          onClick={() => { setModoReset((v) => !v); setError(''); setAviso(''); }}
          className="w-full text-xs text-app-dim2 text-center py-1"
        >
          {modoReset ? '← Volver' : '¿Olvidaste tu contraseña?'}
        </button>

        <p className="text-xs text-app-dim text-center border-t border-app-line pt-4">
          Si no tienes cuenta, pídele al administrador que te dé de alta.
        </p>
      </div>
    </div>
  );
}

// ============================================================
// PANTALLA VINCULAR — la cuenta entró pero no es ningún comprador
// ============================================================
function PantallaVincular({ correo, usuarios = [], onVincular, onSalir }) {
  const sinCorreo = usuarios.filter(u => !u.email);
  const [seleccionado, setSeleccionado] = useState('');

  return (
    <div className="min-h-screen bg-app-bg flex flex-col items-center justify-center px-6 font-sans">
      <style>{APP_STYLES}</style>
      <div className="w-full max-w-sm bg-app-panel border border-app-line rounded-app-lg p-8 space-y-4 text-center">
        <p className="text-4xl">🔗</p>
        <h1 className="text-lg font-bold text-app-white">Cuenta sin vincular</h1>
        <p className="text-sm text-app-dim2">
          Entraste como <span className="text-app-gold">{correo}</span>, pero ese correo no está
          asignado a ningún comprador.
        </p>

        {sinCorreo.length > 0 ? (
          <div className="space-y-3 text-left">
            <p className="text-sm text-app-white font-medium text-center">¿Cuál de estos compradores eres?</p>
            <select
              value={seleccionado}
              onChange={e => setSeleccionado(e.target.value)}
              className="w-full px-3 py-2 rounded-app bg-app-bg border border-app-line text-app-white text-sm"
            >
              <option value="">— Elige tu nombre —</option>
              {sinCorreo.map(u => (
                <option key={u.id} value={u.id}>#{u.numero} — {u.nombre}</option>
              ))}
            </select>
            <button
              disabled={!seleccionado}
              onClick={() => onVincular(seleccionado)}
              className="w-full py-3 rounded-app bg-app-gold text-app-bg font-semibold text-sm disabled:opacity-40"
            >
              Soy este comprador
            </button>
          </div>
        ) : (
          <p className="text-xs text-app-dim">
            Pídele al administrador que agregue tu correo en Administración → Compradores.
          </p>
        )}

        <button
          onClick={onSalir}
          className="w-full text-xs text-app-dim2 text-center py-1"
        >
          Salir
        </button>
      </div>
    </div>
  );
}

const APP_STYLES = `
  .bg-app-bg { background-color: #13151a; }
  .bg-app-panel { background-color: #1a1d24; }
  .bg-app-bg-95 { background-color: rgba(19,21,26,0.95); }
  .bg-app-bg-80 { background-color: rgba(19,21,26,0.80); }
  .bg-app-blue { background-color: #1a2e3a; }
  .bg-app-green { background-color: #1a3a24; }
  .bg-app-active { background-color: #22252d; }
  .bg-app-redbg { background-color: #3a1a1a; }
  .bg-app-goldbg { background-color: #3a2e1a; }
  .bg-app-sky { background-color: #5db8e8; }
  .bg-app-gold { background-color: #e8a33d; }
  .border-app-line { border-color: #2a2d35; }
  .border-app-line2 { border-color: #2a4a5a; }
  .border-app-line3 { border-color: #3a3d45; }
  .border-app-gold { border-color: #e8a33d; }
  .text-app-bg { color: #13151a; }
  .text-app-dim3 { color: #3a3d45; }
  .text-app-dim4 { color: #4a4d55; }
  .text-app-sky { color: #5db8e8; }
  .text-app-green { color: #5de89a; }
  .text-app-dim { color: #6b6f7a; }
  .text-app-dim2 { color: #9a9da5; }
  .text-app-light { color: #c9ccd3; }
  .text-app-red { color: #e85d5d; }
  .text-app-red2 { color: #e87d7d; }
  .text-app-gold { color: #e8a33d; }
  .text-app-white { color: #f0ede4; }
  .divide-app-line > * + * { border-color: #2a2d35; }
  /* Sombras sutiles */
  .shadow-app { box-shadow: 0 1px 3px rgba(0,0,0,0.2); }
  .shadow-app-lg { box-shadow: 0 4px 16px rgba(0,0,0,0.3); }
  /* Radio de 18px para contenedores grandes */
  .rounded-app-lg { border-radius: 18px; }
`;

function PantallaOrigen({ empresa, onSelect, paisesPermitidos }) {
  const origenes = paisesPermitidos
    ? ORIGENES.filter((o) => paisesPermitidos.includes(o.id))
    : ORIGENES;
  return (
    <div className="min-h-screen bg-app-bg flex flex-col items-center justify-center px-6 gap-4 font-sans">
      <style>{APP_STYLES}</style>
      <div className="text-center mb-4">
        {empresa?.logo ? (
          <img src={empresa.logo} alt="Logo" className="h-16 w-auto mx-auto mb-4" />
        ) : (
          <Globe size={36} className="text-app-gold mx-auto mb-3" />
        )}
        <h1 className="text-2xl font-bold text-app-white">Pedidos</h1>
        <p className="text-sm text-app-dim2 mt-1">¿Desde dónde vas a hacer el pedido?</p>
      </div>
      {origenes.length === 0 && (
        <p className="text-sm text-app-dim text-center max-w-sm">
          Tu cuenta no tiene países asignados todavía. Pídele al administrador que te
          habilite acceso en Administración → Compradores.
        </p>
      )}
      {origenes.map((o) => (
        <button
          key={o.id}
          onClick={() => onSelect(o)}
          className="w-full max-w-sm bg-app-panel border border-app-line rounded-app-lg p-5 flex items-center gap-4 active:scale-[0.98] transition shadow-app"
        >
          <span className="text-4xl">{o.emoji}</span>
          <div className="text-left">
            <p className="text-lg font-semibold text-app-white">{o.label}</p>
            <p className="text-xs text-app-dim2">Moneda: {o.moneda}</p>
          </div>
          <ChevronRight size={20} className="text-app-dim ml-auto shrink-0" />
        </button>
      ))}
    </div>
  );
}



function roundToNext9(x) {
  if (x <= 9) return 9;
  const n = Math.ceil((x - 9) / 10);
  return n * 10 + 9;
}

// Calcula el costo puesto en bodega en Lempiras.
// - Origen China: costo RMB → USD (÷ tasa RMB/USD) → + % Niki → × factor China = Lempiras
// - Origen USA/Panamá: costo USD × factor del país = Lempiras
// Niki es un recargo en % que aplica SOLO a China, tras convertir a USD, antes del factor.
function costoBodegaHNL(costoMonto, costoMoneda, origenId, factores, tasaRmbUsd, nikiPct) {
  const monto = parseFloat(costoMonto);
  if (!monto || monto <= 0) return 0;
  const factor = parseFloat(factores?.[origenId]) || 0;
  if (!factor) return 0;
  let usd = monto;
  if (costoMoneda === 'RMB') {
    const tasa = parseFloat(tasaRmbUsd) || 7.25;
    usd = monto / tasa;
  }
  if (origenId === 'china') {
    const pct = parseFloat(nikiPct) || 0;
    usd = usd * (1 + pct / 100);
  }
  return usd * factor;
}

// Devuelve el valor Niki en USD (solo aplica a China). Null si no aplica o faltan datos.
function calcularNikiUSD(costoMonto, costoMoneda, origenId, tasaRmbUsd, nikiPct) {
  if (origenId !== 'china') return null;
  const monto = parseFloat(costoMonto);
  if (!monto || monto <= 0) return null;
  const tasa = parseFloat(tasaRmbUsd) || 7.25;
  const usd = costoMoneda === 'RMB' ? monto / tasa : monto;
  const pct = parseFloat(nikiPct) || 0;
  return usd * (1 + pct / 100);
}

// Sugiere precios de venta a partir del costo bodega en Lempiras.
// Fórmula real: precio = costo × (1 + margen%) × 1.15 (IVA), redondeado a terminación 9.
// Cada origen tiene su propio set de opciones:
// - China: 15 opciones con etiqueta doble "margen% = etiqueta%"
// - USA/Panamá: 9 opciones (100% a 180% en pasos de 10%) con etiqueta simple
function sugerirPreciosVentaHNL(costoBodega, origenId) {
  if (!costoBodega || costoBodega <= 0) return [];
  const IVA = 1.15;
  const opciones = origenId === 'china'
    ? [
        { margen: 15, etiqueta: 120 },
        { margen: 20, etiqueta: 128 },
        { margen: 25, etiqueta: 138 },
        { margen: 30, etiqueta: 147 },
        { margen: 35, etiqueta: 157 },
        { margen: 40, etiqueta: 166 },
        { margen: 45, etiqueta: 176 },
        { margen: 50, etiqueta: 185 },
        { margen: 55, etiqueta: 195 },
        { margen: 60, etiqueta: 204 },
        { margen: 65, etiqueta: 214 },
        { margen: 70, etiqueta: 223 },
        { margen: 80, etiqueta: 242 },
        { margen: 90, etiqueta: 261 },
        { margen: 100, etiqueta: 280 },
      ]
    : origenId === 'honduras'
    ? [
        // Honduras: márgenes 40% a 100% en pasos variados
        { margen: 40 }, { margen: 45 }, { margen: 50 },
        { margen: 55 }, { margen: 60 }, { margen: 70 },
        { margen: 80 }, { margen: 90 }, { margen: 100 },
      ]
    : [
        // USA y Panamá: márgenes 100% a 180% en pasos de 10%
        { margen: 100 }, { margen: 110 }, { margen: 120 },
        { margen: 130 }, { margen: 140 }, { margen: 150 },
        { margen: 160 }, { margen: 170 }, { margen: 180 },
      ];
  return opciones.map(({ margen, etiqueta }) => ({
    margen,
    etiqueta,
    precio: roundToNext9(costoBodega * (1 + margen / 100) * IVA),
  }));
}

// Función anterior (por si algo la sigue usando)
function sugerirPreciosVenta(costoMonto, costoMoneda) {
  const monto = parseFloat(costoMonto);
  if (!monto || monto <= 0) return [];
  const mults = MULTIPLICADORES[costoMoneda] || [];
  return mults.map((m) => ({ multiplicador: m, precio: roundToNext9(monto * m) }));
}

const uid = () => Math.random().toString(36).slice(2, 10);

// Correlativo de pedido por origen y año: CN-26-0001, US-26-0001, PA-26-0001
// Reinicia numeración cada año.
const PREFIJO_ORIGEN = { china: 'CN', usa: 'US', panama: 'PA', honduras: 'HO' };
function generarNumeroPedido(orders, origenId) {
  const anio = String(new Date().getFullYear()).slice(-2);
  const prefijo = `${PREFIJO_ORIGEN[origenId] || 'XX'}-${anio}-`;
  let max = 0;
  (orders || []).forEach((o) => {
    if (o.numero && o.numero.startsWith(prefijo)) {
      const num = parseInt(o.numero.slice(prefijo.length), 10);
      if (!isNaN(num) && num > max) max = num;
    }
  });
  return `${prefijo}${String(max + 1).padStart(4, '0')}`;
}
const variantKey = (talla, color) => `${talla}__${color}`;

const fmtMoneda = (n, moneda) =>
  `${SIMBOLO[moneda] || ''} ${new Intl.NumberFormat('es-HN', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(n || 0)} ${moneda}`;

const fmtLempiras = (n) => fmtMoneda(n, 'HNL');
const todayISO = () => new Date().toISOString().slice(0, 10);

const sumVariantes = (variantes) => (variantes || []).reduce((s, v) => s + (v.cantidad || 0), 0);
const varLabel = (v) => (v.cintura ? `Cintura ${v.cintura} / Largo ${v.largo}` : v.talla);

// ---------- Main App ----------
// ------------------------------------------------------------
// Visor de fotos a pantalla completa.
// Muestra la imagen en resolución original y permite pasar
// entre las fotos del producto con flechas o deslizando.
// ------------------------------------------------------------
function VisorFotos({ fotos = [], indiceInicial = 0, titulo = '', onCerrar }) {
  const [i, setI] = useState(indiceInicial);
  const total = fotos.length;
  const tactoX = useRef(null);

  const anterior = useCallback(() => setI((n) => (n - 1 + total) % total), [total]);
  const siguiente = useCallback(() => setI((n) => (n + 1) % total), [total]);

  // Teclado: flechas para navegar, Esc para salir
  useEffect(() => {
    const alPulsar = (e) => {
      if (e.key === 'Escape') onCerrar();
      else if (e.key === 'ArrowLeft' && total > 1) anterior();
      else if (e.key === 'ArrowRight' && total > 1) siguiente();
    };
    window.addEventListener('keydown', alPulsar);
    return () => window.removeEventListener('keydown', alPulsar);
  }, [anterior, siguiente, onCerrar, total]);

  // Bloquear el desplazamiento del fondo mientras el visor está abierto
  useEffect(() => {
    const previo = document.body.style.overflow;
    document.body.style.overflow = 'hidden';
    return () => { document.body.style.overflow = previo; };
  }, []);

  if (!total) return null;

  const alTocarInicio = (e) => { tactoX.current = e.touches[0].clientX; };
  const alTocarFin = (e) => {
    if (tactoX.current === null || total < 2) return;
    const dif = e.changedTouches[0].clientX - tactoX.current;
    if (Math.abs(dif) > 50) (dif > 0 ? anterior() : siguiente());
    tactoX.current = null;
  };

  const actual = fotos[i];

  return (
    <div
      className="fixed inset-0 z-50 bg-black flex flex-col"
      style={{
        paddingTop: 'env(safe-area-inset-top)',
        paddingBottom: 'env(safe-area-inset-bottom)',
      }}
      onTouchStart={alTocarInicio}
      onTouchEnd={alTocarFin}
    >
      {/* Barra superior */}
      <div className="flex items-center justify-between px-4 py-3 text-white shrink-0">
        <span className="text-sm truncate pr-3">{titulo}</span>
        <div className="flex items-center gap-3 shrink-0">
          {total > 1 && <span className="text-xs opacity-70">{i + 1} / {total}</span>}
          <button
            onClick={onCerrar}
            className="bg-white/10 rounded-full p-2 -mr-1"
            aria-label="Cerrar"
          >
            <X size={22} />
          </button>
        </div>
      </div>

      {/* Imagen */}
      <div
        className="flex-1 flex items-center justify-center relative min-h-0 px-2"
        onClick={(e) => { if (e.target === e.currentTarget) onCerrar(); }}
      >
        <img
          src={urlFoto(actual, false)}
          alt={titulo}
          className="max-w-full max-h-full object-contain"
        />

        {total > 1 && (
          <>
            <button
              onClick={anterior}
              className="absolute left-2 top-1/2 -translate-y-1/2 bg-white/10 hover:bg-white/20 text-white rounded-full p-2"
              aria-label="Anterior"
            >
              <ChevronLeft size={24} />
            </button>
            <button
              onClick={siguiente}
              className="absolute right-2 top-1/2 -translate-y-1/2 bg-white/10 hover:bg-white/20 text-white rounded-full p-2"
              aria-label="Siguiente"
            >
              <ChevronRight size={24} />
            </button>
          </>
        )}
      </div>

      {/* Miniaturas */}
      {total > 1 && (
        <div className="flex gap-2 overflow-x-auto px-4 py-3 shrink-0">
          {fotos.map((f, n) => (
            <button
              key={f.path || n}
              onClick={() => setI(n)}
              className={`w-14 h-14 rounded-md overflow-hidden shrink-0 border-2 ${
                n === i ? 'border-white' : 'border-transparent opacity-50'
              }`}
            >
              <img src={urlFoto(f)} alt="" className="w-full h-full object-cover" />
            </button>
          ))}
        </div>
      )}

      {/* Descargar la original */}
      <div className="px-4 pb-4 shrink-0">
        <a
          href={urlFoto(actual, false)}
          target="_blank"
          rel="noopener noreferrer"
          className="block text-center text-xs text-white/60 hover:text-white py-1"
        >
          Abrir original en pestaña nueva
        </a>
      </div>
    </div>
  );
}

// ------------------------------------------------------------
// Red de seguridad: si algo falla al dibujar la pantalla, en vez de
// quedar en negro se muestra el error y un botón para recuperarse.
// ------------------------------------------------------------
class LimiteDeError extends React.Component {
  constructor(props) {
    super(props);
    this.state = { error: null };
  }
  static getDerivedStateFromError(error) {
    return { error };
  }
  componentDidCatch(error, info) {
    console.error('Fallo en la interfaz:', error, info);
  }
  render() {
    if (!this.state.error) return this.props.children;
    return (
      <div className="min-h-screen bg-app-bg text-app-white font-sans flex items-center justify-center px-6">
        <style>{APP_STYLES}</style>
        <div className="w-full max-w-md bg-app-panel border border-app-line rounded-app-lg p-6 space-y-4">
          <p className="text-3xl text-center">⚠️</p>
          <h1 className="text-lg font-bold text-center">Algo salió mal</h1>
          <p className="text-sm text-app-dim2 text-center">
            La pantalla no se pudo mostrar. Tus datos guardados están a salvo.
          </p>
          <pre className="text-xs text-app-dim3 bg-app-bg border border-app-line rounded-app p-3 overflow-auto max-h-40 whitespace-pre-wrap">
            {String(this.state.error?.message || this.state.error)}
          </pre>
          <div className="flex gap-2">
            <button
              onClick={() => this.setState({ error: null })}
              className="flex-1 py-2.5 rounded-app border border-app-line text-sm"
            >
              Intentar de nuevo
            </button>
            <button
              onClick={() => window.location.reload()}
              className="flex-1 py-2.5 rounded-app bg-app-gold text-app-bg text-sm font-semibold"
            >
              Recargar
            </button>
          </div>
        </div>
      </div>
    );
  }
}

// Estado para listas: garantiza que el valor SIEMPRE sea un arreglo.
// Si algo intenta guardar un objeto, null o un valor suelto (por datos
// antiguos o una respuesta inesperada del servidor), se ignora y queda [].
// Así ningún .filter/.map del resto de la app puede reventar.
function useLista(inicial = []) {
  const [valor, setValor] = useState(Array.isArray(inicial) ? inicial : []);
  const setSeguro = useCallback((nuevo) => {
    setValor((prev) => {
      const resuelto = typeof nuevo === 'function' ? nuevo(prev) : nuevo;
      if (!Array.isArray(resuelto)) {
        console.error('useLista: se intentó guardar algo que no es lista:', resuelto);
        return prev;
      }
      return resuelto;
    });
  }, []);
  return [valor, setSeguro];
}

function PedidosAppInterno() {
  const [view, setView] = useState('pedidos');
  const [loading, setLoading] = useState(true);
  const [origenActivo, setOrigenActivo] = useState(null); // null = pantalla de selección

  const [products, setProducts] = useLista([]);
  const [suppliers, setSuppliers] = useLista([]);
  const [orders, setOrders] = useLista([]);
  const [departamentos, setDepartamentos] = useLista([]);
  const [tipos, setTipos] = useLista([]);
  const [marcas, setMarcas] = useLista([]);
  const [empresa, setEmpresa] = useState({});
  const [embarcadores, setEmbarcadores] = useLista([]);
  const [marcasProveedores, setMarcasProveedores] = useState({}); // marca -> [supplierId]
  const [ciudades, setCiudades] = useLista([]);
  const [fabricas, setFabricas] = useLista([]);
  const [presupuestos, setPresupuestos] = useState({});
  const [usuarios, setUsuarios] = useLista([]);
  const [factores, setFactores] = useState({ china: 32, usa: 22, panama: 18, honduras: 1 });
  const [borradores, setBorradores] = useState({}); // { china: {...}, usa: {...}, panama: {...} }
  const [usuarioActivo, setUsuarioActivo] = useState(null);
  const [sesionAuth, setSesionAuth] = useState(null);   // sesion de Supabase
  const [authCargando, setAuthCargando] = useState(true);
  const [ultimoNombreLogin, setUltimoNombreLogin] = useState('');
  const sesionAuthRef = useRef(null);   // recuerda el correo activo, para poder registrar el cierre de sesión
  // Modo de visualización: 'auto' (responsive) o 'movil' (forzar vista compacta en PC)
  const [modoVista, setModoVista] = useState('auto');
  // Ancho de la ventana (para responsive)
  const [anchoVentana, setAnchoVentana] = useState(typeof window !== 'undefined' ? window.innerWidth : 400);
  useEffect(() => {
    const onResize = () => setAnchoVentana(window.innerWidth);
    window.addEventListener('resize', onResize);
    return () => window.removeEventListener('resize', onResize);
  }, []);
  // Escritorio real = pantalla ancha Y modo automático (no forzado a móvil)
  const esEscritorio = anchoVentana >= 1024 && modoVista === 'auto';
  const [tasaCambio, setTasaCambio] = useState({ rmbUsd: 7.25 });

  const [activeOrderId, setActiveOrderId] = useState(null);

  useEffect(() => {
    // Los datos compartidos viven en Supabase y requieren sesión iniciada.
    // Si aún no hay sesión, esperamos: este efecto se repite cuando llega.
    if (SUPABASE_CONFIGURADO && !sesionAuth) return;
    (async () => {
      const [p, s, o, d, t, mc, emp, embs, mcp, tc, ci, fa, pr, us, fac, bo, sesion, modoGuardado, ultimoUsuario] = await Promise.all([
        loadShared(KEYS.products, seedProducts()),
        loadShared(KEYS.suppliers, seedSuppliers()),
        loadShared(KEYS.orders, []),
        loadShared(KEYS.departamentos, seedDepartamentos()),
        loadShared(KEYS.tipos, seedTipos()),
        loadShared(KEYS.marcas, seedMarcas()),
        loadShared(KEYS.empresa, {}),
        loadShared(KEYS.embarcadores, []),
        loadShared(KEYS.marcasProveedores, {}),
        loadShared(KEYS.tasaCambio, { rmbUsd: 7.25 }),
        loadShared(KEYS.ciudades, seedCiudades()),
        loadShared(KEYS.fabricas, seedFabricas()),
        loadShared(KEYS.presupuestos, {}),
        loadShared(KEYS.usuarios, []),
        loadShared(KEYS.factores, { china: 32, usa: 22, panama: 18, honduras: 1 }),
        loadShared(KEYS.borradores, {}),
        loadPersonal(SESION_KEY, null),
        loadPersonal(MODO_KEY, 'auto'),
        loadPersonal(ULTIMO_USUARIO_KEY, null),
      ]);
      // Salvaguarda: garantiza que las listas nunca queden en un tipo incorrecto,
      // aunque el shared storage devuelva algo inesperado por residuos históricos.
      const arr = (v) => (Array.isArray(v) ? v : []);
      setProducts(arr(p));
      setSuppliers(arr(s));
      setOrders(arr(o));
      setDepartamentos(arr(d));
      setTipos(arr(t));
      setMarcas(arr(mc));
      if (emp && typeof emp === 'object') setEmpresa(emp);
      if (Array.isArray(embs)) setEmbarcadores(embs);
      if (mcp && typeof mcp === 'object') setMarcasProveedores(mcp);
      setTasaCambio(tc);
      setCiudades(ci);
      setFabricas(fa);
      setPresupuestos(pr && typeof pr === 'object' ? pr : {});
      const listaUsuarios = Array.isArray(us) ? us : [];
      setUsuarios(listaUsuarios);
      if (fac && typeof fac === 'object') setFactores({ china: 32, usa: 22, panama: 18, honduras: 1, ...fac });
      if (bo && typeof bo === 'object') setBorradores(bo);
      if (modoGuardado === 'movil' || modoGuardado === 'auto') setModoVista(modoGuardado);
      if (ultimoUsuario && ultimoUsuario.nombre) setUltimoNombreLogin(ultimoUsuario.nombre);
      // La sesión ahora la maneja Supabase Auth (ver efecto más abajo).
      // Ya no se restaura por usuarioId guardado en el dispositivo.
      setLoading(false);
    })();
  }, [sesionAuth]);

  // ---------- Modo offline: cola de sincronización ----------
  const [online, setOnline] = useState(typeof navigator !== 'undefined' ? navigator.onLine : true);
  const [pendienteSync, setPendienteSync] = useState(false);
  const [sincronizando, setSincronizando] = useState(false);
  const pendingRef = useRef(new Map()); // key -> último valor sin guardar

  // Guarda con tolerancia a fallos: si no hay conexión (o falla), encola y reintenta después
  const guardarConCola = useCallback(async (key, value) => {
    // Salvaguarda: nunca escribir null/undefined, borraría los datos guardados.
    if (value === null || value === undefined) {
      console.error('guardarConCola: intento de guardar vacío en', key, '— ignorado');
      return;
    }
    // Sin conexión: va directo a la cola y se avisa al usuario
    if (typeof navigator !== 'undefined' && !navigator.onLine) {
      pendingRef.current.set(key, value);
      setPendienteSync(true);
      return;
    }
    try {
      await window.storage.set(key, JSON.stringify(value), true);
      // Guardado correcto: si esta clave estaba pendiente de un intento previo, se libera
      pendingRef.current.delete(key);
      if (pendingRef.current.size === 0) setPendienteSync(false);
    } catch (e) {
      // Solo aquí hay algo que avisar: el guardado falló y queda en cola para reintentar
      console.error('No se pudo guardar', key, e);
      pendingRef.current.set(key, value);
      setPendienteSync(true);
    }
  }, []);

  // Reintenta guardar todo lo pendiente
  const sincronizar = useCallback(async () => {
    if (pendingRef.current.size === 0) { setPendienteSync(false); return; }
    setSincronizando(true);
    const entradas = [...pendingRef.current.entries()];
    for (const [key, value] of entradas) {
      try {
        await window.storage.set(key, JSON.stringify(value), true);
        pendingRef.current.delete(key);
      } catch (e) { /* sigue pendiente */ }
    }
    setSincronizando(false);
    setPendienteSync(pendingRef.current.size > 0);
  }, []);

  // Detectar cambios de conexión y reintentar periódicamente
  useEffect(() => {
    const onOnline = () => { setOnline(true); sincronizar(); };
    const onOffline = () => setOnline(false);
    window.addEventListener('online', onOnline);
    window.addEventListener('offline', onOffline);
    const intervalo = setInterval(() => {
      if (navigator.onLine && pendingRef.current.size > 0) sincronizar();
    }, 30000);
    return () => {
      window.removeEventListener('online', onOnline);
      window.removeEventListener('offline', onOffline);
      clearInterval(intervalo);
    };
  }, [sincronizar]);

  const actorAuditoria = useCallback(
    () => ({ email: usuarioActivo?.email || sesionAuth?.user?.email || null, nombre: usuarioActivo?.nombre || null }),
    [usuarioActivo, sesionAuth]
  );
  useEffect(() => { sesionAuthRef.current = sesionAuth; }, [sesionAuth]);

  const persistProducts = useCallback((next) => {
    setProducts((prev) => {
      registrarAuditoria('productos', prev, next, actorAuditoria(), (p) => `${p.codigo || ''} — ${p.descripcion || ''}`);
      return next;
    });
    guardarConCola(KEYS.products, next);
  }, [guardarConCola, actorAuditoria]);
  const persistSuppliers = useCallback((next) => {
    setSuppliers((prev) => {
      registrarAuditoria('proveedores', prev, next, actorAuditoria(), (s) => s.nombre || '');
      return next;
    });
    guardarConCola(KEYS.suppliers, next);
  }, [guardarConCola, actorAuditoria]);
  const persistOrders = useCallback((next) => {
    setOrders((prev) => {
      registrarAuditoria('pedidos', prev, next, actorAuditoria(), (o) => `Pedido ${o.numero || o.id}`);
      return next;
    });
    guardarConCola(KEYS.orders, next);
  }, [guardarConCola, actorAuditoria]);
  const persistTasaCambio = useCallback((next) => { setTasaCambio(next); guardarConCola(KEYS.tasaCambio, next); }, [guardarConCola]);
  const persistDepartamentos = useCallback((next) => { setDepartamentos(next); guardarConCola(KEYS.departamentos, next); }, [guardarConCola]);
  const persistTipos = useCallback((next) => { setTipos(next); guardarConCola(KEYS.tipos, next); }, [guardarConCola]);
  const persistMarcas = useCallback((next) => { setMarcas(next); guardarConCola(KEYS.marcas, next); }, [guardarConCola]);
  const persistEmpresa = useCallback((next) => { setEmpresa(next); guardarConCola(KEYS.empresa, next); }, [guardarConCola]);
  const persistEmbarcadores = useCallback((next) => { setEmbarcadores(next); guardarConCola(KEYS.embarcadores, next); }, [guardarConCola]);
  const persistMarcasProveedores = useCallback((next) => { setMarcasProveedores(next); guardarConCola(KEYS.marcasProveedores, next); }, [guardarConCola]);
  const persistCiudades = useCallback((next) => { setCiudades(next); guardarConCola(KEYS.ciudades, next); }, [guardarConCola]);
  const persistFabricas = useCallback((next) => { setFabricas(next); guardarConCola(KEYS.fabricas, next); }, [guardarConCola]);
  const persistPresupuestos = useCallback((next) => { setPresupuestos(next); guardarConCola(KEYS.presupuestos, next); }, [guardarConCola]);
  const persistUsuarios = useCallback((next) => {
    setUsuarios((prev) => {
      registrarAuditoria('compradores', prev, next, actorAuditoria(), (u) => `#${u.numero ?? ''} ${u.nombre || ''}`);
      return next;
    });
    guardarConCola(KEYS.usuarios, next);
  }, [guardarConCola, actorAuditoria]);
  const persistFactores = useCallback((next) => {
    setFactores((prev) => {
      registrarAuditoriaObjeto('factores_costo', prev, next, actorAuditoria(), 'Cambio en factores de costo puesto en bodega');
      return next;
    });
    guardarConCola(KEYS.factores, next);
  }, [guardarConCola, actorAuditoria]);

  // Guarda/actualiza el borrador de un origen (compartido con el equipo)
  const guardarBorrador = useCallback((origenId, datos) => {
    setBorradores((prev) => {
      const anterior = prev?.[origenId];
      // Comparar solo los campos de contenido (excluye ultimaEdicion) para evitar
      // re-guardar cuando el JSON efectivo es igual. Esto evita loops.
      const camposContenido = ['supplierId', 'items', 'notas', 'codigoInicio', 'embarcadorId'];
      const iguales = anterior && camposContenido.every(
        (k) => JSON.stringify(anterior[k]) === JSON.stringify(datos[k])
      );
      if (iguales) return prev; // sin cambio, no re-renderizar
      // Quién lo creó: se conserva el creador original si ya existía; si es nuevo, el usuario actual.
      const creadoPor = anterior?.creadoPor || usuarioActivo?.nombre || '';
      const next = { ...prev, [origenId]: { ...datos, creadoPor, ultimaEdicion: new Date().toISOString() } };
      guardarConCola(KEYS.borradores, next);
      return next;
    });
  }, [guardarConCola, usuarioActivo]);
  const limpiarBorrador = useCallback((origenId) => {
    setBorradores((prev) => {
      const next = { ...prev };
      delete next[origenId];
      guardarConCola(KEYS.borradores, next);
      return next;
    });
  }, [guardarConCola]);

  // Cierra la sesión en Supabase. El efecto de autenticación limpia el resto.
  const handleLogout = useCallback(async () => {
    try { if (supabase) await supabase.auth.signOut(); } catch (e) { /* ignorar */ }
    setUsuarioActivo(null);
    setSesionAuth(null);
    deletePersonal(SESION_KEY);
  }, []);

  // Asigna el correo de la sesión actual al comprador elegido.
  // Si es el primero en vincularse, lo marca como admin.
  const handleVincular = useCallback((compradorId) => {
    const correo = (sesionAuth?.user?.email || '').trim().toLowerCase();
    if (!correo) return;
    const ningunoVinculado = usuarios.every(u => !u.email);
    const next = usuarios.map(u =>
      u.id === compradorId
        ? { ...u, email: correo, esAdmin: ningunoVinculado ? true : u.esAdmin }
        : u
    );
    persistUsuarios(next);
  }, [sesionAuth, usuarios, persistUsuarios]);

  // ---------- Autenticación con Supabase ----------
  // 1) Al abrir la app: leer la sesión guardada y quedar atento a cambios
  useEffect(() => {
    if (!supabase) { setAuthCargando(false); return; }
    let vivo = true;
    (async () => {
      try {
        const { data } = await supabase.auth.getSession();
        if (vivo) setSesionAuth(data?.session || null);
      } catch (e) {
        console.error('Error leyendo sesión', e);
      } finally {
        if (vivo) setAuthCargando(false);
      }
    })();
    const { data: sub } = supabase.auth.onAuthStateChange((evento, sesion) => {
      setSesionAuth(sesion || null);
      if (!sesion) setUsuarioActivo(null);
      if (evento === 'SIGNED_IN' && sesion?.user?.email) {
        supabase.from('auditoria').insert([{
          tabla: 'sesion', accion: 'crear', resumen: 'Inicio de sesión',
          usuario_email: sesion.user.email, usuario_nombre: null,
        }]).then(() => {}, () => {});
      } else if (evento === 'SIGNED_OUT') {
        supabase.from('auditoria').insert([{
          tabla: 'sesion', accion: 'borrar', resumen: 'Cierre de sesión',
          usuario_email: sesionAuthRef.current?.user?.email || null, usuario_nombre: null,
        }]).then(() => {}, () => {});
      }
    });
    return () => { vivo = false; sub?.subscription?.unsubscribe(); };
  }, []);

  // 2) Vincular la sesión con el comprador que tenga ese correo
  useEffect(() => {
    if (loading || !sesionAuth) return;
    const correo = (sesionAuth.user?.email || '').trim().toLowerCase();
    if (!correo) return;

    const encontrado = usuarios.find(
      (u) => (u.email || '').trim().toLowerCase() === correo
    );
    if (encontrado) {
      // Si nadie tiene esAdmin aún, promover al primero que entre
      const hayAdmin = usuarios.some(u => u.esAdmin);
      if (!encontrado.esAdmin && !hayAdmin) {
        const next = usuarios.map(u => u.id === encontrado.id ? { ...u, esAdmin: true } : u);
        setUsuarios(next);
        saveShared(KEYS.usuarios, next);
        return; // el efecto se re-ejecuta con el usuario actualizado
      }
      if (usuarioActivo?.id !== encontrado.id) setUsuarioActivo(encontrado);
      return;
    }

    // Primer arranque: si no hay ningún comprador, este correo se vuelve el administrador
    if (usuarios.length === 0 && !usuarioActivo) {
      const primero = {
        id: uid(),
        numero: 1,
        nombre: correo.split('@')[0],
        prefijo: '',
        email: correo,
        esAdmin: true,
      };
      const next = [primero];
      setUsuarios(next);
      saveShared(KEYS.usuarios, next);
      setUsuarioActivo(primero);
    }
  }, [sesionAuth, usuarios, loading, usuarioActivo]);

  const toggleModoVista = useCallback(() => {
    setModoVista((prev) => {
      const next = prev === 'auto' ? 'movil' : 'auto';
      savePersonal(MODO_KEY, next);
      return next;
    });
  }, []);

  // Creación rápida desde los selectores (no duplica si ya existe)
  const crearMarca = useCallback((nombre) => {
    setMarcas((prev) => {
      if (prev.some((m) => m.toLowerCase() === nombre.toLowerCase())) return prev;
      const next = [...prev, nombre].sort((a, b) => a.localeCompare(b));
      guardarConCola(KEYS.marcas, next);
      return next;
    });
  }, [guardarConCola]);
  const crearFabrica = useCallback((nombre) => {
    setFabricas((prev) => {
      if (prev.some((f) => f.toLowerCase() === nombre.toLowerCase())) return prev;
      const next = [...prev, nombre].sort((a, b) => a.localeCompare(b));
      guardarConCola(KEYS.fabricas, next);
      return next;
    });
  }, [guardarConCola]);
  const crearCiudad = useCallback((nombre) => {
    setCiudades((prev) => {
      if (prev.some((c) => c.toLowerCase() === nombre.toLowerCase())) return prev;
      const next = [...prev, nombre];
      guardarConCola(KEYS.ciudades, next);
      return next;
    });
  }, [guardarConCola]);

  // Gate de acceso
  // 1) Faltan las variables de entorno de Supabase
  if (!SUPABASE_CONFIGURADO) return <PantallaSinConfigurar />;

  // 2) Todavía verificando si hay sesión guardada
  if (authCargando) {
    return (
      <div className="min-h-screen bg-app-bg flex items-center justify-center font-sans">
        <style>{APP_STYLES}</style>
        <p className="text-app-dim2 text-sm">Cargando…</p>
      </div>
    );
  }

  // 3) Sin sesión: pedir correo y contraseña
  if (!sesionAuth) return <PantallaLogin />;

  // 3b) Ya con sesión: traer los datos compartidos desde Supabase
  if (loading) {
    return (
      <div className="min-h-screen bg-app-bg flex items-center justify-center">
        <div className="text-app-gold text-sm tracking-widest uppercase animate-pulse">
          Cargando inventario…
        </div>
      </div>
    );
  }

  // 4) Con sesión pero el correo no corresponde a ningún comprador
  if (!usuarioActivo) {
    return (
      <PantallaVincular
        correo={sesionAuth.user?.email || ''}
        usuarios={usuarios}
        onVincular={handleVincular}
        onSalir={handleLogout}
      />
    );
  }

  // ¿El usuario activo es administrador? (mismo criterio que en Administración)
  // Regla: u.esAdmin === true. Respaldo: si nadie es admin explícito, Yamil Handal lo es.
  const hayAdminExplicito = (usuarios || []).some((x) => x.esAdmin);
  const soyAdmin = !!usuarioActivo && (
    usuarioActivo.esAdmin ||
    (!hayAdminExplicito && (usuarioActivo.nombre || '').trim().toLowerCase() === 'yamil handal')
  );
  // Rol: admin siempre tiene acceso total. Si no es admin, se usa el campo "rol"
  // ('supervisor' o 'comprador'); si no está definido, se trata como comprador.
  const miRol = soyAdmin ? 'admin' : (usuarioActivo?.rol === 'supervisor' ? 'supervisor' : 'comprador');
  const soyPrivilegiado = miRol === 'admin' || miRol === 'supervisor';
  // Países que el usuario puede ver. Admin/supervisor: todos. Comprador: los que
  // tenga asignados; si el campo no existe (cuentas antiguas), ve todos por defecto
  // para no dejar a nadie fuera al activar esta función.
  const paisesPermitidos = soyPrivilegiado
    ? ORIGENES.map((o) => o.id)
    : (Array.isArray(usuarioActivo?.paises) ? usuarioActivo.paises : ORIGENES.map((o) => o.id));

  // Pantalla de selección de origen (o si el país activo ya no está permitido para este usuario)
  if (!origenActivo || !paisesPermitidos.includes(origenActivo.id)) {
    return (
      <PantallaOrigen
        empresa={empresa}
        paisesPermitidos={paisesPermitidos}
        onSelect={(o) => { setOrigenActivo(o); setView('pedidos'); }}
      />
    );
  }

  // Filtrar datos por origen activo
  // Proveedores del origen activo. El filtro es ESTRICTO: cada proveedor pertenece a un
  // solo país de compra. Los que quedaron sin país (creados antes de esta regla) se
  // muestran aparte en la pantalla de Proveedores para asignarles el suyo.
  const suppliersOrigen = (Array.isArray(suppliers) ? suppliers : []).filter((s) => s.origen === origenActivo.id);
  const suppliersSinOrigen = (Array.isArray(suppliers) ? suppliers : []).filter((s) => !s.origen);
  const productsOrigen = (Array.isArray(products) ? products : []).filter((p) => !p.origen || p.origen === origenActivo.id);

  const ordersOrigenTodas = (Array.isArray(orders) ? orders : []).filter((o) => !o.origen || o.origen === origenActivo.id);
  // Un comprador solo ve sus propios pedidos; supervisor y admin ven todos.
  const ordersOrigen = soyPrivilegiado
    ? ordersOrigenTodas
    : ordersOrigenTodas.filter((o) => !o.creadoPor || o.creadoPor === usuarioActivo?.nombre);

  // El pedido "sin terminar" (borrador) de este país es compartido en la base de datos,
  // pero solo debe verlo quien lo está armando; el administrador ve todos.
  const borradorDelPais = borradores?.[origenActivo.id] || null;
  const borradorEsMio = !!borradorDelPais && (
    soyAdmin || !borradorDelPais.creadoPor || borradorDelPais.creadoPor === usuarioActivo?.nombre
  );
  const borradorVisible = borradorEsMio ? borradorDelPais : null;
  // Si hay un borrador de otra persona, no dejamos empezar uno nuevo (se perdería el suyo)
  const borradorAjenoBloquea = !!borradorDelPais && !borradorEsMio;

  const activeOrder = ordersOrigen.find((o) => o.id === activeOrderId)
    || (soyPrivilegiado ? orders.find((o) => o.id === activeOrderId) : null); // fallback para pedidos sin origen (solo si puede verlos)

  return (
    <div
      className={`min-h-screen bg-app-bg text-app-white font-sans ${esEscritorio ? '' : 'pb-24'}`}
      style={esEscritorio ? undefined : { paddingBottom: 'calc(6rem + env(safe-area-inset-bottom))' }}
    >
      <style>{APP_STYLES + `
        .placeholder-app-dim::placeholder { color: #6b6f7a; }
      `}</style>
      <Header
        view={view}
        setView={setView}
        origen={origenActivo}
        onCambiarOrigen={() => setOrigenActivo(null)}
        modoVista={modoVista}
        onToggleModoVista={toggleModoVista}
        esEscritorio={esEscritorio}
        borradorAjenoBloquea={borradorAjenoBloquea}
        usuarioActivo={usuarioActivo}
        onLogout={handleLogout}
      />

      {/* Banner de estado de conexión */}
      {!online ? (
        <div className="bg-app-goldbg border-b border-app-line px-4 py-2 text-center">
          <p className="text-xs text-app-gold font-medium">📡 Sin conexión — trabajando localmente</p>
          <p className="text-xs text-app-dim2 mt-0.5">No cierres la app. Tus cambios se guardarán solos al volver el internet.</p>
        </div>
      ) : pendienteSync ? (
        <div className="bg-app-blue border-b border-app-line2 px-4 py-2 flex items-center justify-center gap-2">
          <p className="text-xs text-app-sky">
            {sincronizando ? 'Guardando cambios pendientes…' : 'Hay cambios sin guardar.'}
          </p>
          {!sincronizando && (
            <button onClick={sincronizar} className="text-xs text-app-sky underline font-medium">Guardar ahora</button>
          )}
        </div>
      ) : null}

      {esEscritorio && (
        <SideNav view={view} setView={setView} soyAdmin={soyAdmin} soyPrivilegiado={soyPrivilegiado} />
      )}

      <main className={esEscritorio
        ? "ml-56 max-w-6xl px-6 pt-6"
        : "max-w-md mx-auto px-4 pt-4"
      }>
        {view === 'pedidos' && (
          <PedidosList
            orders={ordersOrigen}
            suppliers={suppliersOrigen}
            borrador={borradorVisible}
            borradorAjenoBloquea={borradorAjenoBloquea}
            onOpen={(id) => { setActiveOrderId(id); setView('detalle'); }}
            onNew={() => setView('nuevo')}
            onContinuarBorrador={() => setView('nuevo')}
            onDescartarBorrador={() => limpiarBorrador(origenActivo.id)}
          />
        )}

        {view === 'nuevo' && (
          <NuevoPedido
            embarcadores={embarcadores}
            products={productsOrigen}
            setProducts={persistProducts}
            departamentos={departamentos}
            tipos={tipos}
            setTipos={persistTipos}
            marcas={marcas}
            marcasProveedores={marcasProveedores}
            ciudades={ciudades}
            fabricas={fabricas}
            factores={factores}
            suppliers={suppliersOrigen}
            tasaCambio={tasaCambio}
            setTasaCambio={persistTasaCambio}
            origen={origenActivo}
            borrador={borradorVisible}
            onGuardarBorrador={(datos) => guardarBorrador(origenActivo.id, datos)}
            usuarioActivoNombre={usuarioActivo?.nombre || ''}
            usuarioActivoPrefijo={usuarioActivo?.prefijo || ''}
            onCancel={() => setView('pedidos')}
            onCreate={(pedidosNuevos) => {
              // pedidosNuevos puede ser un solo pedido o un array (cuando se separan por destino)
              const arr = Array.isArray(pedidosNuevos) ? pedidosNuevos : [pedidosNuevos];
              // Asignar correlativo secuencial por origen y año (ej. CN-26-0001)
              let acumulado = [...orders];
              const marcados = arr.map((o) => {
                const numero = generarNumeroPedido(acumulado, origenActivo.id);
                const pedido = { ...o, origen: origenActivo.id, numero, creadoPor: usuarioActivo?.nombre || '' };
                acumulado = [pedido, ...acumulado]; // para que el siguiente pedido del batch tome el siguiente número
                return pedido;
              });
              const next = [...marcados, ...orders];
              persistOrders(next);
              // Limpiar el borrador de este origen (se guardó como pedido definitivo)
              limpiarBorrador(origenActivo.id);
              setActiveOrderId(marcados[0].id);
              setView('detalle');
            }}
            onCreateMarca={crearMarca}
            onCreateFabrica={crearFabrica}
            onCreateCiudad={crearCiudad}
          />
        )}

        {view === 'detalle' && activeOrder && (
          <DetallePedido
            empresa={empresa}
            embarcadores={embarcadores}
            order={activeOrder}
            supplier={suppliers.find((s) => s.id === activeOrder.supplierId)}
            tasaCambio={tasaCambio}
            setTasaCambio={persistTasaCambio}
            onBack={() => setView('pedidos')}
            onUpdateStatus={(status) => {
              const next = orders.map((o) => (o.id === activeOrder.id ? { ...o, status } : o));
              persistOrders(next);
            }}
            hayBorradorPendiente={!!(borradorVisible?.items?.length > 0)}
            borradorAjenoBloquea={borradorAjenoBloquea}
            onDuplicar={() => {
              // Copia los artículos y el proveedor a un pedido nuevo (borrador) y lo abre
              guardarBorrador(origenActivo.id, {
                supplierId: activeOrder.supplierId,
                items: activeOrder.items,
                notas: '',
                codigoInicio: '',
              });
              setView('nuevo');
            }}
          />
        )}

        {view === 'productos' && (
          <Catalogo
            products={productsOrigen}
            setProducts={(next) => {
              // Mezclar con productos de otros orígenes
              const otrosOrigenes = products.filter((p) => p.origen && p.origen !== origenActivo.id);
              persistProducts([...otrosOrigenes, ...next]);
            }}
            departamentos={departamentos}
            tipos={tipos}
            setTipos={persistTipos}
            marcas={marcas}
            ciudades={ciudades}
            fabricas={fabricas}
            factores={factores}
            tasaCambio={tasaCambio}
            origen={origenActivo}
            orders={ordersOrigen}
            suppliers={suppliersOrigen}
            usuarioActivoNombre={usuarioActivo?.nombre || ''}
            usuarioActivoPrefijo={usuarioActivo?.prefijo || ''}
            puedoBorrar={soyPrivilegiado}
            onOpenConfig={() => setView('config')}
            onCreateMarca={crearMarca}
            onCreateFabrica={crearFabrica}
            onCreateCiudad={crearCiudad}
          />
        )}

        {view === 'proveedores' && (
          <Proveedores
            suppliers={suppliers}
            setSuppliers={persistSuppliers}
            origen={origenActivo}
            puedoBorrar={soyPrivilegiado}
          />
        )}

        {view === 'config' && (
          <Config
            departamentos={departamentos}
            setDepartamentos={persistDepartamentos}
            tipos={tipos}
            setTipos={persistTipos}
            marcas={marcas}
            setMarcas={persistMarcas}
            marcasProveedores={marcasProveedores}
            setMarcasProveedores={persistMarcasProveedores}
            suppliers={suppliersOrigen}
            origen={origenActivo}
            ciudades={ciudades}
            setCiudades={persistCiudades}
            fabricas={fabricas}
            setFabricas={persistFabricas}
            onBack={() => setView('productos')}
            onPresupuestos={() => setView('presupuestos')}
          />
        )}

        {view === 'administracion' && soyAdmin && (
          <Administracion
            empresa={empresa}
            setEmpresa={persistEmpresa}
            embarcadores={embarcadores}
            setEmbarcadores={persistEmbarcadores}
            usuarios={usuarios}
            setUsuarios={persistUsuarios}
            usuarioActivo={usuarioActivo}
            onLogout={handleLogout}
            factores={factores}
            setFactores={persistFactores}
            tasaCambio={tasaCambio}
            setTasaCambio={persistTasaCambio}
          />
        )}
        {view === 'administracion' && !soyAdmin && (
          <div className="text-center py-16 text-app-dim2">
            <Settings size={28} className="mx-auto mb-2 text-app-dim3" />
            <p className="text-sm">Esta sección es solo para el administrador.</p>
          </div>
        )}

        {view === 'presupuestos' && (
          <Presupuestos
            presupuestos={presupuestos}
            setPresupuestos={persistPresupuestos}
            departamentos={departamentos}
            tipos={tipos}
            orders={orders}
            tasaCambio={tasaCambio}
            onBack={() => setView('pedidos')}
          />
        )}

        {view === 'reportes' && soyPrivilegiado && (
          <Reportes
            orders={orders}
            suppliers={suppliers}
            tasaCambio={tasaCambio}
            factores={factores}
          />
        )}
      </main>

      {!esEscritorio && (
        <BottomNav view={view} setView={setView} soyAdmin={soyAdmin} soyPrivilegiado={soyPrivilegiado} />
      )}
    </div>
  );
}

// Punto de entrada: la app envuelta en la red de seguridad
export default function PedidosApp() {
  return (
    <LimiteDeError>
      <PedidosAppInterno />
    </LimiteDeError>
  );
}

// ---------- Seed data ----------
function seedDepartamentos() {
  return ['Caballeros', 'Niños'];
}
function seedMarcas() {
  return ['Collezione'];
}
function seedFabricas() {
  return [];
}
// Ciudades base agrupadas por origen (para filtrar el selector según el país activo)
const CIUDADES_POR_ORIGEN = {
  china: [
    'Guangzhou', 'Shenzhen', 'Shanghai', 'Yiwu', 'Dongguan',
    'Foshan', 'Hangzhou', 'Ningbo', 'Wenzhou', 'Zhongshan',
    'Huizhou', 'Quanzhou', 'Xiamen', 'Jinjiang', 'Suzhou',
  ],
  usa: [
    'Los Ángeles', 'Miami', 'New York', 'Dallas', 'Houston',
    'Chicago', 'Atlanta', 'Las Vegas', 'San Francisco', 'Seattle',
  ],
  panama: ['Ciudad de Panamá', 'Colón'],
  honduras: ['Tegucigalpa', 'San Pedro Sula'],
};

// Devuelve las ciudades a mostrar para un origen: las de ese país + cualquier
// ciudad personalizada que el usuario haya agregado (no listada en ningún origen)
function ciudadesDeOrigen(todasCiudades, origenId) {
  if (!origenId) return todasCiudades;
  const base = CIUDADES_POR_ORIGEN[origenId] || [];
  const todasBase = Object.values(CIUDADES_POR_ORIGEN).flat();
  const personalizadas = (todasCiudades || []).filter((c) => !todasBase.includes(c));
  // Mantener solo las base de este origen que existan en la lista global, + personalizadas
  const deEsteOrigen = (todasCiudades || []).filter((c) => base.includes(c));
  return [...deEsteOrigen, ...personalizadas];
}

function seedCiudades() {
  return [
    ...CIUDADES_POR_ORIGEN.china,
    ...CIUDADES_POR_ORIGEN.usa,
    ...CIUDADES_POR_ORIGEN.panama,
    ...CIUDADES_POR_ORIGEN.honduras,
  ];
}
function seedTipos() {
  return [
    { id: uid(), nombre: 'Playera', medida: 'simple', tallas: ['XS', 'S', 'M', 'L', 'XL', '2XL', '3XL', '4XL'] },
    { id: uid(), nombre: 'Pantalón', medida: 'cintura_largo', cinturas: ['28', '30', '32', '34', '36', '38'], largos: ['30', '32', '34'] },
    { id: uid(), nombre: 'Short', medida: 'simple', tallas: ['XS', 'S', 'M', 'L', 'XL'] },
    { id: uid(), nombre: 'Ropa niño', medida: 'simple', tallas: ['2', '4', '6', '8'] },
    { id: uid(), nombre: 'Ropa Juvenil', medida: 'simple', tallas: ['10', '12', '14', '16', '18'] },
  ];
}
function getTallasDeTipo(tipo) {
  if (!tipo) return [];
  if (tipo.medida === 'cintura_largo') {
    const cinturas = tipo.cinturas || [];
    const largos = tipo.largos || [];
    const combos = [];
    cinturas.forEach((c) => largos.forEach((l) => combos.push(`${c}/${l}`)));
    return combos;
  }
  return tipo.tallas || [];
}
function seedProducts() {
  return [
    {
      id: uid(), codigo: 'PH-001',
      descripcion: 'Playera cuello redondo 100% algodón', departamento: 'Caballeros',
      tipo: 'Playera', medida: 'simple', colores: ['Blanco', 'Negro'],
      variantes: [
        { talla: 'M', color: 'Blanco', cantidad: 30 },
        { talla: 'L', color: 'Blanco', cantidad: 25 },
        { talla: 'M', color: 'Negro', cantidad: 20 },
        { talla: 'L', color: 'Negro', cantidad: 15 },
      ],
      costoMonto: 28, costoMoneda: 'RMB', ventaLempiras: 295, foto: null,
    },
    {
      id: uid(), codigo: 'PA-027',
      descripcion: 'Corte recto, lavado índigo', departamento: 'Caballeros',
      tipo: 'Pantalón', medida: 'cintura_largo', colores: ['Índigo'],
      variantes: [
        { cintura: '32', largo: '32', color: 'Índigo', talla: 'Cintura 32 / Largo 32', cantidad: 20 },
        { cintura: '34', largo: '32', color: 'Índigo', talla: 'Cintura 34 / Largo 32', cantidad: 25 },
        { cintura: '36', largo: '34', color: 'Índigo', talla: 'Cintura 36 / Largo 34', cantidad: 15 },
      ],
      costoMonto: 9.5, costoMoneda: 'USD', ventaLempiras: 590, foto: null,
    },
  ];
}
function seedSuppliers() {
  return [
    { id: uid(), nombre: 'Textiles del Norte SA', contacto: 'Laura Méndez', email: 'ventas@textilesnorte.mx', telefono: '81 1234 5678' },
    { id: uid(), nombre: 'Confecciones Bajío', contacto: 'Rubén Salas', email: 'pedidos@confbajio.mx', telefono: '477 222 3344' },
  ];
}

// ---------- Header / Nav ----------
function Header({ view, setView, origen, onCambiarOrigen, modoVista, onToggleModoVista, esEscritorio, borradorAjenoBloquea = false, usuarioActivo = null, onLogout }) {
  const [menuUsuario, setMenuUsuario] = useState(false);
  const titles = {
    pedidos: 'Pedidos', nuevo: 'Nuevo pedido', productos: 'Catálogo',
    proveedores: 'Proveedores', detalle: 'Detalle de pedido', config: 'Departamentos y tipos',
    presupuestos: 'Presupuesto', reportes: 'Reportes', administracion: 'Administración',
  };
  // El botón de modo solo tiene sentido si la pantalla es ancha (≥1024 px).
  // Si el usuario está en modo móvil forzado en pantalla ancha, se lo mostramos para que pueda volver a auto.
  const mostrarBotonModo = typeof window !== 'undefined' && window.innerWidth >= 1024;
  return (
    <header
      className="sticky top-0 z-20 bg-app-bg-95 backdrop-blur border-b border-app-line shadow-app"
      style={{ paddingTop: 'env(safe-area-inset-top)' }}
    >
      <div className={esEscritorio
        ? "max-w-6xl mx-auto px-6 py-3 flex items-center justify-between"
        : "max-w-md mx-auto px-4 py-3 flex items-center justify-between"
      }>
        <div>
          <button
            onClick={onCambiarOrigen}
            className="flex items-center gap-1.5 text-xs text-app-dim2 mb-0.5 active:opacity-70"
          >
            <span>{origen?.emoji}</span>
            <span>{origen?.label}</span>
            <Globe size={11} />
          </button>
          <h1 className="text-xl font-semibold tracking-tight text-app-white">{titles[view]}</h1>
        </div>
        <div className="flex items-center gap-2">
          {mostrarBotonModo && (
            <button
              onClick={onToggleModoVista}
              className="flex items-center gap-1.5 text-xs bg-app-panel border border-app-line rounded-lg px-2.5 py-2 text-app-light active:bg-app-active"
              title={modoVista === 'auto' ? 'Cambiar a vista móvil compacta' : 'Volver a vista automática'}
            >
              <span className="text-sm">{modoVista === 'auto' ? '💻' : '📱'}</span>
              <span className="hidden sm:inline">{modoVista === 'auto' ? 'Escritorio' : 'Móvil'}</span>
            </button>
          )}
          {view !== 'nuevo' && view !== 'config' && (
            <button
              onClick={() => { if (!borradorAjenoBloquea) setView('nuevo'); }}
              disabled={borradorAjenoBloquea}
              className="bg-app-gold text-app-bg rounded-full p-2.5 active:scale-95 transition disabled:opacity-40"
              aria-label="Nuevo pedido"
              title={borradorAjenoBloquea ? 'Hay un pedido en progreso de otro comprador en este país' : 'Nuevo pedido'}
            >
              <Plus size={20} strokeWidth={2.5} />
            </button>
          )}

          {/* Menú de usuario: disponible en todas las pantallas, para cualquier rol */}
          {usuarioActivo && (
            <div className="relative">
              <button
                onClick={() => setMenuUsuario((v) => !v)}
                className="w-9 h-9 rounded-full bg-app-panel border border-app-line text-app-gold font-bold text-sm flex items-center justify-center active:bg-app-active"
                aria-label="Menú de usuario"
              >
                {(usuarioActivo.nombre || '?').charAt(0).toUpperCase()}
              </button>
              {menuUsuario && (
                <>
                  {/* Capa para cerrar al tocar fuera */}
                  <div className="fixed inset-0 z-30" onClick={() => setMenuUsuario(false)} />
                  <div className="absolute right-0 mt-2 w-56 bg-app-panel border border-app-line rounded-xl shadow-app-lg z-40 overflow-hidden">
                    <div className="px-3 py-2.5 border-b border-app-line">
                      <p className="text-sm text-app-white truncate">{usuarioActivo.nombre}</p>
                      {usuarioActivo.email && (
                        <p className="text-xs text-app-dim2 truncate">{usuarioActivo.email}</p>
                      )}
                      <p className="text-xs text-app-dim3 mt-0.5">
                        {usuarioActivo.esAdmin ? '👑 Administrador'
                          : usuarioActivo.rol === 'supervisor' ? '🔎 Supervisor'
                          : '🛒 Comprador'}
                      </p>
                    </div>
                    <button
                      onClick={() => { setMenuUsuario(false); onLogout && onLogout(); }}
                      className="w-full text-left px-3 py-2.5 text-sm text-app-red2 active:bg-app-active"
                    >
                      Cerrar sesión
                    </button>
                  </div>
                </>
              )}
            </div>
          )}
        </div>
      </div>
    </header>
  );
}

// ---------- Reportes: jerarquía Origen → Comprador → Proveedor → Marca → Depto → Tipo ----------
function Reportes({ orders = [], suppliers = [], tasaCambio, factores }) {
  const [rangoFecha, setRangoFecha] = useState('all'); // all | 30d | 90d | 12m | custom
  const [fechaDesde, setFechaDesde] = useState('');
  const [fechaHasta, setFechaHasta] = useState('');
  const [expandidos, setExpandidos] = useState(new Set()); // paths tipo "china|Yamil|Proveedor"

  const toggle = (path) => {
    setExpandidos((prev) => {
      const next = new Set(prev);
      if (next.has(path)) next.delete(path); else next.add(path);
      return next;
    });
  };
  const isExp = (path) => expandidos.has(path);

  // Filtrar pedidos por rango de fechas
  const pedidosFiltrados = (() => {
    if (rangoFecha === 'all') return orders;
    const ahora = new Date();
    let desde;
    if (rangoFecha === '30d') { desde = new Date(ahora); desde.setDate(desde.getDate() - 30); }
    else if (rangoFecha === '90d') { desde = new Date(ahora); desde.setDate(desde.getDate() - 90); }
    else if (rangoFecha === '12m') { desde = new Date(ahora); desde.setMonth(desde.getMonth() - 12); }
    else if (rangoFecha === 'custom') {
      const d = fechaDesde ? new Date(fechaDesde) : null;
      const h = fechaHasta ? new Date(fechaHasta) : null;
      return orders.filter((o) => {
        if (!o.fecha) return true;              // pedidos sin fecha: no se descartan
        const f = new Date(o.fecha);
        if (isNaN(f)) return true;              // fecha ilegible: tampoco se descarta
        if (d && f < d) return false;
        if (h && f > h) return false;
        return true;
      });
    }
    return orders.filter((o) => {
      if (!o.fecha) return true;
      const f = new Date(o.fecha);
      return isNaN(f) ? true : f >= desde;
    });
  })();

  // Calcular totales de un item en Lempiras (costo bodega) y USD
  const totalesItem = (it) => {
    const bodega = costoBodegaHNL(it.costoMonto, it.costoMoneda, it.origen, factores, tasaCambio?.rmbUsd, factores?.nikiPct);
    const pzs = sumVariantes(it.variantes);
    let usd = it.costoMonto || 0;
    if (it.costoMoneda === 'RMB') usd = usd / (parseFloat(tasaCambio?.rmbUsd) || 7.25);
    return { pzs, hnl: bodega * pzs, usd: usd * pzs };
  };

  // Construir jerarquía anidada
  const arbol = {}; // { origen: { comprador: { proveedor: { marca: { depto: { tipo: {pzs, hnl, usd} } } } } } }
  pedidosFiltrados.forEach((o) => {
    const origenId = o.origen || 'china';
    const comprador = o.creadoPor || 'Sin comprador';
    const supplier = suppliers.find((s) => s.id === o.supplierId);
    const nombreProveedor = supplier?.nombre || 'Proveedor eliminado';
    (o.items || []).forEach((it) => {
      const marca = it.marca || 'Sin marca';
      const depto = it.departamento || 'Sin departamento';
      const tipo = it.tipo || 'Sin tipo';
      const { pzs, hnl, usd } = totalesItem(it);
      arbol[origenId] = arbol[origenId] || {};
      arbol[origenId][comprador] = arbol[origenId][comprador] || {};
      arbol[origenId][comprador][nombreProveedor] = arbol[origenId][comprador][nombreProveedor] || {};
      arbol[origenId][comprador][nombreProveedor][marca] = arbol[origenId][comprador][nombreProveedor][marca] || {};
      arbol[origenId][comprador][nombreProveedor][marca][depto] = arbol[origenId][comprador][nombreProveedor][marca][depto] || {};
      const bucket = arbol[origenId][comprador][nombreProveedor][marca][depto];
      bucket[tipo] = bucket[tipo] || { pzs: 0, hnl: 0, usd: 0 };
      bucket[tipo].pzs += pzs;
      bucket[tipo].hnl += hnl;
      bucket[tipo].usd += usd;
    });
  });

  // Sumar totales de un nivel dado (recursivo)
  const sumarNivel = (obj) => {
    let pzs = 0, hnl = 0, usd = 0;
    const walk = (n) => {
      if (n?.pzs !== undefined && n?.hnl !== undefined) {
        pzs += n.pzs; hnl += n.hnl; usd += n.usd; return;
      }
      Object.values(n || {}).forEach(walk);
    };
    walk(obj);
    return { pzs, hnl, usd };
  };

  const origenesConDatos = Object.keys(arbol);

  const exportarExcel = () => {
    const rows = [];
    Object.entries(arbol).forEach(([origenId, porComprador]) => {
      Object.entries(porComprador).forEach(([comprador, porProveedor]) => {
        Object.entries(porProveedor).forEach(([proveedor, porMarca]) => {
          Object.entries(porMarca).forEach(([marca, porDepto]) => {
            Object.entries(porDepto).forEach(([depto, porTipo]) => {
              Object.entries(porTipo).forEach(([tipo, tot]) => {
                rows.push({
                  Origen: origenId,
                  Comprador: comprador,
                  Proveedor: proveedor,
                  Marca: marca,
                  Departamento: depto,
                  Tipo: tipo,
                  Piezas: tot.pzs,
                  'Total USD': parseFloat((tot.usd || 0).toFixed(2)),
                });
              });
            });
          });
        });
      });
    });
    if (rows.length === 0) rows.push({ Origen: 'Sin datos' });
    const ws = XLSX.utils.json_to_sheet(rows);
    ws['!cols'] = [{ wch: 10 }, { wch: 16 }, { wch: 22 }, { wch: 18 }, { wch: 22 }, { wch: 18 }, { wch: 10 }, { wch: 14 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Reporte');
    descargarLibro(wb, `reporte_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const Total = ({ pzs, usd }) => (
    <span className="text-xs text-app-dim2 whitespace-nowrap ml-2">
      {pzs} pzs · <span className="text-app-gold">{fmtMoneda(usd, 'USD')}</span>
    </span>
  );

  return (
    <div className="space-y-3">
      <h2 className="text-lg font-semibold">📊 Reportes</h2>

      {/* Filtro de fechas */}
      <div className="bg-app-panel border border-app-line rounded-xl p-3 space-y-2">
        <label className="text-xs uppercase tracking-wide text-app-dim2">Rango de fechas</label>
        <div className="flex flex-wrap gap-1.5">
          {[
            { id: 'all', label: 'Todo' },
            { id: '30d', label: 'Último mes' },
            { id: '90d', label: 'Último trimestre' },
            { id: '12m', label: 'Último año' },
            { id: 'custom', label: 'Personalizado' },
          ].map((r) => (
            <button
              key={r.id}
              onClick={() => setRangoFecha(r.id)}
              className={`text-xs rounded-lg border px-2.5 py-1.5 ${rangoFecha === r.id ? 'bg-app-gold text-app-bg border-app-gold font-semibold' : 'bg-app-bg border-app-line text-app-light'}`}
            >
              {r.label}
            </button>
          ))}
        </div>
        {rangoFecha === 'custom' && (
          <div className="flex gap-2 mt-1">
            <input type="date" value={fechaDesde} onChange={(e) => setFechaDesde(e.target.value)}
              className="flex-1 bg-app-bg border border-app-line rounded-lg px-2 py-1.5 text-xs" />
            <input type="date" value={fechaHasta} onChange={(e) => setFechaHasta(e.target.value)}
              className="flex-1 bg-app-bg border border-app-line rounded-lg px-2 py-1.5 text-xs" />
          </div>
        )}
        <button onClick={exportarExcel}
          className="w-full mt-1 py-2 rounded-lg bg-app-panel border border-app-line text-app-sky text-xs font-medium flex items-center justify-center gap-1.5">
          <FileDown size={13} /> Exportar reporte a Excel
        </button>
      </div>

      {/* Árbol jerárquico */}
      {origenesConDatos.length === 0 ? (
        <div className="text-center py-12 text-app-dim2">
          <BarChart3 size={28} className="mx-auto mb-2 text-app-dim3" />
          <p className="text-sm">Sin pedidos en el rango seleccionado.</p>
        </div>
      ) : (
        <div className="space-y-1.5">
          {origenesConDatos.map((origenId) => {
            const origenObj = arbol[origenId];
            const totalOrigen = sumarNivel(origenObj);
            const pathOrigen = origenId;
            const origenInfo = ORIGENES.find((o) => o.id === origenId) || { emoji: '🏳', label: origenId };
            return (
              <div key={origenId} className="bg-app-panel border border-app-line rounded-xl overflow-hidden">
                {/* Origen */}
                <button onClick={() => toggle(pathOrigen)} className="w-full flex items-center justify-between px-3 py-2.5 active:bg-app-active">
                  <div className="flex items-start gap-2 min-w-0 flex-1 pr-2">
                    <ChevronDown size={14} className={`text-app-dim shrink-0 transition-transform ${isExp(pathOrigen) ? '' : '-rotate-90'}`} />
                    <span className="text-base">{origenInfo.emoji}</span>
                    <span className="text-sm font-semibold">{origenInfo.label}</span>
                  </div>
                  <Total {...totalOrigen} />
                </button>

                {/* Compradores */}
                {isExp(pathOrigen) && Object.entries(origenObj).map(([comprador, compObj]) => {
                  const pathComp = `${pathOrigen}|${comprador}`;
                  const totalComp = sumarNivel(compObj);
                  return (
                    <div key={comprador} className="border-t border-app-line">
                      <button onClick={() => toggle(pathComp)} className="w-full flex items-center justify-between px-3 py-2 pl-8 active:bg-app-active">
                        <div className="flex items-start gap-2 min-w-0 flex-1 pr-2">
                          <ChevronDown size={12} className={`text-app-dim shrink-0 transition-transform ${isExp(pathComp) ? '' : '-rotate-90'}`} />
                          <span className="text-xs font-medium text-app-light">👤 {comprador}</span>
                        </div>
                        <Total {...totalComp} />
                      </button>

                      {/* Proveedores */}
                      {isExp(pathComp) && Object.entries(compObj).map(([prov, provObj]) => {
                        const pathProv = `${pathComp}|${prov}`;
                        const totalProv = sumarNivel(provObj);
                        return (
                          <div key={prov} className="border-t border-app-line bg-app-bg">
                            <button onClick={() => toggle(pathProv)} className="w-full flex items-center justify-between px-3 py-2 pl-14 active:bg-app-active">
                              <div className="flex items-start gap-2 min-w-0 flex-1 pr-2">
                                <ChevronDown size={12} className={`text-app-dim shrink-0 transition-transform ${isExp(pathProv) ? '' : '-rotate-90'}`} />
                                <span className="text-xs text-app-light break-words">🚚 {prov}</span>
                              </div>
                              <Total {...totalProv} />
                            </button>

                            {/* Marcas */}
                            {isExp(pathProv) && Object.entries(provObj).map(([marca, marcaObj]) => {
                              const pathMarca = `${pathProv}|${marca}`;
                              const totalMarca = sumarNivel(marcaObj);
                              return (
                                <div key={marca} className="border-t border-app-line">
                                  <button onClick={() => toggle(pathMarca)} className="w-full flex items-center justify-between px-3 py-1.5 pl-20 active:bg-app-active">
                                    <div className="flex items-start gap-2 min-w-0 flex-1 pr-2">
                                      <ChevronDown size={11} className={`text-app-dim shrink-0 transition-transform ${isExp(pathMarca) ? '' : '-rotate-90'}`} />
                                      <span className="text-xs text-app-dim2 break-words">🏷 {marca}</span>
                                    </div>
                                    <Total {...totalMarca} />
                                  </button>

                                  {/* Departamentos */}
                                  {isExp(pathMarca) && Object.entries(marcaObj).map(([depto, deptoObj]) => {
                                    const pathDepto = `${pathMarca}|${depto}`;
                                    const totalDepto = sumarNivel(deptoObj);
                                    return (
                                      <div key={depto} className="border-t border-app-line">
                                        <button onClick={() => toggle(pathDepto)} className="w-full flex items-center justify-between px-3 py-1.5 pl-24 active:bg-app-active">
                                          <div className="flex items-start gap-2 min-w-0 flex-1 pr-2">
                                            <ChevronDown size={11} className={`text-app-dim shrink-0 transition-transform ${isExp(pathDepto) ? '' : '-rotate-90'}`} />
                                            <span className="text-xs text-app-dim2 break-words">📁 {depto}</span>
                                          </div>
                                          <Total {...totalDepto} />
                                        </button>

                                        {/* Tipos (hoja del árbol) */}
                                        {isExp(pathDepto) && Object.entries(deptoObj).map(([tipo, tot]) => (
                                          <div key={tipo} className="border-t border-app-line flex items-start justify-between px-3 py-1.5 pl-28 bg-app-panel gap-2">
                                            <span className="text-xs text-app-dim break-words flex-1 pr-2">• {tipo}</span>
                                            <Total {...tot} />
                                          </div>
                                        ))}
                                      </div>
                                    );
                                  })}
                                </div>
                              );
                            })}
                          </div>
                        );
                      })}
                    </div>
                  );
                })}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

// ---------- SideNav: barra lateral izquierda (solo escritorio) ----------
function SideNav({ view, setView, soyAdmin, soyPrivilegiado }) {
  const items = [
    { key: 'pedidos', label: 'Pedidos', icon: ClipboardList },
    { key: 'productos', label: 'Catálogo', icon: Package },
    { key: 'proveedores', label: 'Proveedores', icon: Truck },
    { key: 'presupuestos', label: 'Presupuesto', icon: Wallet },
    ...(soyPrivilegiado ? [{ key: 'reportes', label: 'Reportes', icon: BarChart3 }] : []),
    ...(soyAdmin ? [{ key: 'administracion', label: 'Administración', icon: Settings }] : []),
  ];
  return (
    <nav className="fixed left-0 top-14 bottom-0 w-56 bg-app-panel border-r border-app-line z-10 overflow-y-auto shadow-app">
      <div className="p-3 space-y-1">
        {items.map(({ key, label, icon: Icon }) => {
          const active = view === key
            || (view === 'detalle' && key === 'pedidos')
            || (view === 'nuevo' && key === 'pedidos')
            || (view === 'config' && key === 'productos')

          return (
            <button
              key={key}
              onClick={() => setView(key)}
              className={`w-full flex items-center gap-2.5 px-3 py-2.5 rounded-lg text-sm font-medium transition ${
                active
                  ? 'bg-app-goldbg text-app-gold'
                  : 'text-app-light hover:bg-app-active'
              }`}
            >
              <Icon size={16} className="shrink-0" />
              <span>{label}</span>
            </button>
          );
        })}
      </div>
    </nav>
  );
}

function BottomNav({ view, setView, soyAdmin, soyPrivilegiado }) {
  const items = [
    { key: 'pedidos', label: 'Pedidos', icon: ClipboardList },
    { key: 'productos', label: 'Catálogo', icon: Package },
    { key: 'proveedores', label: 'Proveed.', icon: Truck },
    { key: 'presupuestos', label: 'Presup.', icon: Wallet },
    ...(soyPrivilegiado ? [{ key: 'reportes', label: 'Reportes', icon: BarChart3 }] : []),
    ...(soyAdmin ? [{ key: 'administracion', label: 'Admin', icon: Settings }] : []),
  ];
  return (
    <nav
      className="fixed bottom-0 inset-x-0 bg-app-panel border-t border-app-line z-10 shadow-app-lg"
      style={{ paddingBottom: 'env(safe-area-inset-bottom)' }}
    >
      <div className="max-w-md mx-auto flex">
        {items.map(({ key, label, icon: Icon }) => {
          const active = view === key
            || (view === 'detalle' && key === 'pedidos')
            || (view === 'nuevo' && key === 'pedidos')
            || (view === 'config' && key === 'productos')

          return (
            <button
              key={key}
              onClick={() => setView(key)}
              className={`flex-1 flex flex-col items-center gap-1 py-3 transition ${active ? 'text-app-gold' : 'text-app-dim'}`}
            >
              <Icon size={19} strokeWidth={active ? 2.5 : 2} />
              <span className="text-[10px] font-medium">{label}</span>
            </button>
          );
        })}
      </div>
    </nav>
  );
}

// ---------- Pedidos list ----------
const ESTADOS_PEDIDO = {
  pendiente: { bg: 'bg-app-goldbg', text: 'text-app-gold', label: 'Pendiente' },
  enviado: { bg: 'bg-app-blue', text: 'text-app-sky', label: 'Enviado' },
  recibido: { bg: 'bg-app-green', text: 'text-app-green', label: 'Recibido' },
};

function StatusPill({ status }) {
  const map = ESTADOS_PEDIDO;
  const s = map[status] || map.pendiente;
  return <span className={`${s.bg} ${s.text} text-xs font-medium px-2.5 py-1 rounded-full`}>{s.label}</span>;
}

function itemTotal(item) {
  return sumVariantes(item.variantes) * item.costoMonto;
}
function orderTotalsByCurrency(items) {
  const totals = {};
  items.forEach((it) => {
    const m = it.costoMoneda || 'USD';
    totals[m] = (totals[m] || 0) + itemTotal(it);
  });
  return totals;
}

function PedidosList({ orders = [], suppliers = [], borrador, borradorAjenoBloquea = false, onOpen, onNew, onContinuarBorrador, onDescartarBorrador }) {
  const [query, setQuery] = useState('');
  const [filtroEstado, setFiltroEstado] = useState('todos');
  const [filtroDestino, setFiltroDestino] = useState('todos');
  const [gruposAbiertos, setGruposAbiertos] = useState(new Set()); // proveedores expandidos

  // Filtrado: texto (número, proveedor, comprador), estado y destino
  const ordersFiltrados = orders.filter((o) => {
    if (filtroEstado !== 'todos' && o.status !== filtroEstado) return false;
    if (filtroDestino !== 'todos' && (o.destinoPedido || '') !== filtroDestino) return false;
    const q = query.trim().toLowerCase();
    if (q) {
      const supplier = suppliers.find((s) => s.id === o.supplierId);
      const texto = `${o.numero || ''} ${supplier?.nombre || ''} ${o.creadoPor || ''} ${o.fecha || ''} ${o.notas || ''}`.toLowerCase();
      if (!texto.includes(q)) return false;
    }
    return true;
  });

  const hayFiltro = !!(query.trim() || filtroEstado !== 'todos' || filtroDestino !== 'todos');
  const limpiarFiltros = () => { setQuery(''); setFiltroEstado('todos'); setFiltroDestino('todos'); };
  const contarEstado = (st) => (st === 'todos' ? orders.length : orders.filter((o) => o.status === st).length);
  const hayDestinos = orders.some((o) => o.destinoPedido);

  const barraFiltros = (
    <div className="space-y-2 mb-3">
      <div className="relative">
        <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-app-dim" />
        <input
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          placeholder="Buscar por número, proveedor o comprador…"
          className="w-full bg-app-panel border border-app-line rounded-xl pl-9 pr-3 py-2.5 text-sm"
        />
      </div>
      <div className="flex gap-1.5 overflow-x-auto pb-0.5">
        {[
          { id: 'todos', label: 'Todos' },
          { id: 'pendiente', label: 'Pendientes' },
          { id: 'enviado', label: 'Enviados' },
          { id: 'recibido', label: 'Recibidos' },
        ].map((e) => (
          <button
            key={e.id}
            onClick={() => setFiltroEstado(e.id)}
            className={`text-xs rounded-lg border px-2.5 py-1.5 whitespace-nowrap shrink-0 ${
              filtroEstado === e.id
                ? 'bg-app-gold text-app-bg border-app-gold font-semibold'
                : 'bg-app-panel border-app-line text-app-light'
            }`}
          >
            {e.label} <span className="opacity-70">{contarEstado(e.id)}</span>
          </button>
        ))}
      </div>
      {hayDestinos && (
        <div className="flex gap-1.5">
          {[{ id: 'todos', label: 'Ambos países' }, ...DESTINOS.map((d) => ({ id: d.id, label: `${d.emoji} ${d.label}` }))].map((d) => (
            <button
              key={d.id}
              onClick={() => setFiltroDestino(d.id)}
              className={`text-xs rounded-lg border px-2.5 py-1.5 whitespace-nowrap ${
                filtroDestino === d.id
                  ? 'bg-app-gold text-app-bg border-app-gold font-semibold'
                  : 'bg-app-panel border-app-line text-app-light'
              }`}
            >
              {d.label}
            </button>
          ))}
        </div>
      )}
      {hayFiltro && (
        <div className="flex items-center justify-between">
          <span className="text-xs text-app-dim2">
            {ordersFiltrados.length} de {orders.length} pedido{orders.length !== 1 ? 's' : ''}
          </span>
          <button onClick={limpiarFiltros} className="text-xs text-app-red2 active:opacity-70">Limpiar filtros</button>
        </div>
      )}
    </div>
  );

  const borradorPiezas = borrador?.items ? borrador.items.reduce((s, it) => s + sumVariantes(it.variantes), 0) : 0;
  const supplierBorrador = borrador?.supplierId ? suppliers.find((s) => s.id === borrador.supplierId) : null;
  const bannerBorrador = borrador && borrador.items?.length > 0 ? (
    <div className="bg-app-blue border border-app-line2 rounded-xl p-3 mb-3">
      <div className="flex items-start justify-between gap-2">
        <div className="min-w-0">
          <p className="text-sm font-semibold text-app-sky">📝 Pedido sin terminar</p>
          <p className="text-xs text-app-light mt-0.5">
            {supplierBorrador ? supplierBorrador.nombre : 'Sin proveedor'} · {borrador.items?.length || 0} art · {borradorPiezas} pzs
          </p>
        </div>
      </div>
      <div className="flex gap-2 mt-2.5">
        <button onClick={onDescartarBorrador} className="flex-1 py-2 rounded-lg border border-app-line text-xs text-app-dim2 active:bg-app-active">
          Descartar
        </button>
        <button onClick={onContinuarBorrador} className="flex-1 py-2 rounded-lg bg-app-sky text-app-bg text-xs font-semibold active:opacity-80">
          Continuar
        </button>
      </div>
    </div>
  ) : null;

  const avisoBorradorAjeno = borradorAjenoBloquea ? (
    <div className="bg-app-panel border border-app-line rounded-xl p-3 mb-3">
      <p className="text-sm font-semibold text-app-gold">⏳ Hay un pedido en progreso</p>
      <p className="text-xs text-app-dim2 mt-0.5">
        Otro comprador está armando un pedido en este país todavía. Cuando lo termine o lo descarte, podrás crear el tuyo.
      </p>
    </div>
  ) : null;

  const manejarNuevo = () => { if (!borradorAjenoBloquea) onNew(); };

  if (orders.length === 0) {
    return (
      <div>
        {bannerBorrador}
        {avisoBorradorAjeno}
        <div className="flex flex-col items-center text-center mt-16 gap-3">
          <ClipboardList size={36} className="text-app-dim3" />
          <p className="text-app-dim2 text-sm max-w-xs">Aún no hay pedidos. Crea el primero para tu equipo.</p>
          <button
            onClick={manejarNuevo}
            disabled={borradorAjenoBloquea}
            className="mt-2 bg-app-gold text-app-bg font-semibold text-sm px-5 py-2.5 rounded-full active:scale-95 transition disabled:opacity-40"
          >
            Crear pedido
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="space-y-2.5">
      {bannerBorrador}
      {barraFiltros}
      {ordersFiltrados.length === 0 && (
        <div className="text-center py-10 text-app-dim2">
          <Search size={26} className="mx-auto mb-2 text-app-dim3" />
          <p className="text-sm">Ningún pedido coincide con esos filtros.</p>
          <button onClick={limpiarFiltros} className="text-xs text-app-sky mt-2 active:opacity-70">Limpiar filtros</button>
        </div>
      )}

      {(() => {
        if (ordersFiltrados.length === 0) return null;
        // Agrupar por proveedor
        const grupos = {};
        ordersFiltrados.forEach((o) => {
          const sup = suppliers.find((s) => s.id === o.supplierId);
          const nombre = sup ? sup.nombre : 'Proveedor eliminado';
          if (!grupos[nombre]) grupos[nombre] = [];
          grupos[nombre].push(o);
        });
        const nombresOrdenados = Object.keys(grupos).sort((a, b) => a.localeCompare(b));

        const toggle = (nombre) => {
          setGruposAbiertos((prev) => {
            const next = new Set(prev);
            next.has(nombre) ? next.delete(nombre) : next.add(nombre);
            return next;
          });
        };

        const tarjetaPedido = (o) => {
          const supplier = suppliers.find((s) => s.id === o.supplierId);
          const totals = orderTotalsByCurrency(o.items);
          const piezas = (o.items || []).reduce((s, it) => s + sumVariantes(it.variantes), 0);
          return (
            <button
              key={o.id}
              onClick={() => onOpen(o.id)}
              className="w-full bg-app-panel rounded-xl p-3.5 text-left border border-app-line active:scale-95 transition flex items-center justify-between"
            >
              <div className="min-w-0">
                <div className="flex items-center gap-2 mb-1">
                  <StatusPill status={o.status} />
                  <span className="text-xs text-app-dim">{o.fecha}</span>
                  {o.destinoPedido && (
                    <span className="text-xs bg-app-bg border border-app-line rounded-full px-2 py-0.5 text-app-light">
                      {destinoInfo(o.destinoPedido).emoji} {destinoInfo(o.destinoPedido).label}
                    </span>
                  )}
                </div>
                {o.numero && <p className="text-xs font-mono text-app-gold">{o.numero}</p>}
                {o.creadoPor && <p className="text-xs text-app-dim">por {o.creadoPor}</p>}
                <p className="text-xs text-app-dim2">
                  {(o.items || []).length} producto{(o.items || []).length !== 1 ? 's' : ''} · {piezas} pzs ·{' '}
                  {Object.entries(totals).map(([m, v], i) => (
                    <span key={m}>{i > 0 ? ' + ' : ''}{fmtMoneda(v, m)}</span>
                  ))}
                </p>
              </div>
              <ChevronRight size={18} className="text-app-dim4 shrink-0" />
            </button>
          );
        };

        return (
          <div className="space-y-2">
            {nombresOrdenados.map((nombre) => {
              const pedidos = grupos[nombre];
              const abierto = gruposAbiertos.has(nombre);
              return (
                <div key={nombre} className="bg-app-panel border border-app-line rounded-2xl overflow-hidden">
                  <button
                    onClick={() => toggle(nombre)}
                    className="w-full flex items-center gap-2 px-4 py-3.5 text-left active:bg-app-active"
                  >
                    <ChevronDown size={16} className={`text-app-dim shrink-0 transition-transform ${abierto ? '' : '-rotate-90'}`} />
                    <div className="min-w-0 flex-1">
                      <p className="font-medium truncate">{nombre}</p>
                      <p className="text-xs text-app-dim2">
                        {pedidos.length} pedido{pedidos.length !== 1 ? 's' : ''}
                      </p>
                    </div>
                  </button>
                  {abierto && (
                    <div className="px-2 pb-2 space-y-2">
                      {pedidos.map(tarjetaPedido)}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        );
      })()}
    </div>
  );
}

// ---------- Variant matrix (shared between Catalogo y Pedido) ----------
function VariantMatrix({ tallas = [], colores = [], values, onChange, accent = '#e8a33d' }) {
  // values: { "talla__color": cantidad }
  if (tallas.length === 0 || colores.length === 0) {
    return (
      <p className="text-xs text-app-dim italic">
        Selecciona un tipo de producto y agrega al menos un color para capturar cantidades.
      </p>
    );
  }
  const setCell = (talla, color, val) => {
    const next = { ...values, [variantKey(talla, color)]: val };
    onChange(next);
  };
  const total = Object.values(values).reduce((s, v) => s + (parseInt(v) || 0), 0);

  return (
    <div>
      <div className="overflow-x-auto rounded-lg border border-app-line">
        <table className="w-full text-xs">
          <thead>
            <tr>
              <th className="text-left px-2 py-1.5 bg-app-bg text-app-dim2 sticky left-0">Talla \ Color</th>
              {colores.map((c) => (
                <th key={c} className="px-2 py-1.5 bg-app-bg text-app-dim2 font-medium whitespace-nowrap">{c}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {tallas.map((t) => (
              <tr key={t} className="border-t border-app-line">
                <td className="px-2 py-1.5 text-app-light sticky left-0 bg-app-panel font-medium">{t}</td>
                {colores.map((c) => (
                  <td key={c} className="px-1 py-1">
                    <input
                      type="number"
                      min={0}
                      value={values[variantKey(t, c)] ?? ''}
                      onChange={(e) => setCell(t, c, e.target.value)}
                      placeholder="0"
                      className="w-12 bg-app-bg border border-app-line rounded px-1 py-1 text-center text-xs"
                    />
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <p className="text-xs text-app-dim2 mt-1.5">
        Total piezas: <span style={{ color: accent }} className="font-semibold">{total}</span>
      </p>
    </div>
  );
}

function matrixToVariantes(values) {
  return Object.entries(values)
    .map(([k, v]) => {
      const [talla, color] = k.split('__');
      return { talla, color, cantidad: parseInt(v) || 0 };
    })
    .filter((v) => v.cantidad > 0);
}
function variantesToMatrix(variantes) {
  const map = {};
  (variantes || []).forEach((v) => { map[variantKey(v.talla, v.color)] = v.cantidad; });
  return map;
}

// ---------- Matriz DUAL: cantidades para H y G en la misma tabla ----------
// Claves: "talla__color__H" y "talla__color__G"
const dualKey = (talla, color, destino) => `${talla}__${color}__${destino}`;

function VariantMatrixDual({ tallas = [], colores = [], values, onChange, accent = '#e8a33d' }) {
  if (tallas.length === 0 || colores.length === 0) {
    return (
      <p className="text-xs text-app-dim italic">
        Selecciona un tipo de producto y agrega al menos un color para capturar cantidades.
      </p>
    );
  }
  const setCell = (talla, color, destino, val) => {
    onChange({ ...values, [dualKey(talla, color, destino)]: val });
  };
  const totalH = Object.entries(values).reduce((s, [k, v]) => s + (k.endsWith('__H') ? (parseInt(v) || 0) : 0), 0);
  const totalG = Object.entries(values).reduce((s, [k, v]) => s + (k.endsWith('__G') ? (parseInt(v) || 0) : 0), 0);

  return (
    <div>
      <div className="overflow-x-auto rounded-lg border border-app-line">
        <table className="w-full text-xs">
          <thead>
            <tr>
              <th rowSpan={2} className="text-left px-2 py-1.5 bg-app-bg text-app-dim2 sticky left-0 align-bottom">Talla</th>
              {colores.map((c) => (
                <th key={c} colSpan={2} className="px-2 py-1 bg-app-bg text-app-dim2 font-medium whitespace-nowrap text-center border-l border-app-line">{c}</th>
              ))}
            </tr>
            <tr>
              {colores.map((c) => (
                <React.Fragment key={c}>
                  <th className="px-1 py-1 bg-app-bg text-app-red2 font-medium text-center border-l border-app-line">🇭🇳 H</th>
                  <th className="px-1 py-1 bg-app-bg text-app-green font-medium text-center">🇬🇹 G</th>
                </React.Fragment>
              ))}
            </tr>
          </thead>
          <tbody>
            {tallas.map((t) => (
              <tr key={t} className="border-t border-app-line">
                <td className="px-2 py-1.5 text-app-light sticky left-0 bg-app-panel font-medium">{t}</td>
                {colores.map((c) => (
                  <React.Fragment key={c}>
                    <td className="px-1 py-1 border-l border-app-line">
                      <input type="number" min={0} placeholder="0"
                        value={values[dualKey(t, c, 'H')] ?? ''}
                        onChange={(e) => setCell(t, c, 'H', e.target.value)}
                        className="w-11 bg-app-bg border border-app-line rounded px-1 py-1 text-center text-xs" />
                    </td>
                    <td className="px-1 py-1">
                      <input type="number" min={0} placeholder="0"
                        value={values[dualKey(t, c, 'G')] ?? ''}
                        onChange={(e) => setCell(t, c, 'G', e.target.value)}
                        className="w-11 bg-app-bg border border-app-line rounded px-1 py-1 text-center text-xs" />
                    </td>
                  </React.Fragment>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div className="flex gap-4 mt-1.5">
        <p className="text-xs text-app-dim2">🇭🇳 Honduras: <span className="text-app-red2 font-semibold">{totalH}</span></p>
        <p className="text-xs text-app-dim2">🇬🇹 Afiliada: <span className="text-app-green font-semibold">{totalG}</span></p>
      </div>
    </div>
  );
}

// Extrae variantes de un destino específico de la matriz dual
function dualMatrixToVariantes(values, destino) {
  return Object.entries(values)
    .filter(([k]) => k.endsWith(`__${destino}`))
    .map(([k, v]) => {
      const [talla, color] = k.split('__');
      return { talla, color, cantidad: parseInt(v) || 0 };
    })
    .filter((v) => v.cantidad > 0);
}

// ---------- Matriz Cintura x Largo (por color) ----------
const clKey = (color, cintura, largo) => `${color}__${cintura}__${largo}`;

function matrixCLToVariantes(values) {
  return Object.entries(values)
    .map(([k, v]) => {
      const [color, cintura, largo] = k.split('__');
      return { color, cintura, largo, talla: `Cintura ${cintura} / Largo ${largo}`, cantidad: parseInt(v) || 0 };
    })
    .filter((v) => v.cantidad > 0);
}
function variantesToMatrixCL(variantes) {
  const map = {};
  (variantes || []).forEach((v) => {
    if (v.cintura && v.largo) map[clKey(v.color, v.cintura, v.largo)] = v.cantidad;
  });
  return map;
}

function CinturaLargoMatrix({ cinturas, largos, colores = [], values, onChange, accent = '#e8a33d' }) {
  if (cinturas.length === 0 || largos.length === 0 || colores.length === 0) {
    return (
      <p className="text-xs text-app-dim italic">
        Selecciona un tipo de producto (con cinturas y largos definidos) y agrega al menos un color.
      </p>
    );
  }
  const setCell = (color, cintura, largo, val) => onChange({ ...values, [clKey(color, cintura, largo)]: val });
  const total = Object.values(values).reduce((s, v) => s + (parseInt(v) || 0), 0);

  return (
    <div className="space-y-3">
      {colores.map((color) => {
        const colorTotal = cinturas.reduce(
          (s, c) => s + largos.reduce((s2, l) => s2 + (parseInt(values[clKey(color, c, l)]) || 0), 0),
          0
        );
        return (
          <div key={color}>
            <p className="text-xs font-medium text-app-light mb-1">{color} <span className="text-app-dim font-normal">· {colorTotal} pzs</span></p>
            <div className="overflow-x-auto rounded-lg border border-app-line">
              <table className="w-full text-xs">
                <thead>
                  <tr>
                    <th className="text-left px-2 py-1.5 bg-app-bg text-app-dim2 sticky left-0">Cintura \ Largo</th>
                    {largos.map((l) => (
                      <th key={l} className="px-2 py-1.5 bg-app-bg text-app-dim2 font-medium whitespace-nowrap">{l}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {cinturas.map((c) => (
                    <tr key={c} className="border-t border-app-line">
                      <td className="px-2 py-1.5 text-app-light sticky left-0 bg-app-panel font-medium">{c}</td>
                      {largos.map((l) => (
                        <td key={l} className="px-1 py-1">
                          <input
                            type="number"
                            min={0}
                            value={values[clKey(color, c, l)] ?? ''}
                            onChange={(e) => setCell(color, c, l, e.target.value)}
                            placeholder="0"
                            className="w-12 bg-app-bg border border-app-line rounded px-1 py-1 text-center text-xs"
                          />
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        );
      })}
      <p className="text-xs text-app-dim2">
        Total piezas: <span style={{ color: accent }} className="font-semibold">{total}</span>
      </p>
    </div>
  );
}


// ---------- Bloque de totales con conversión RMB→USD ----------
function TotalesConversion({ totals, tasaCambio, setTasaCambio, label = 'Total estimado (costo)' }) {
  const rmbTotal = totals['RMB'] || 0;
  const usdTotal = totals['USD'] || 0;
  const rate = parseFloat(tasaCambio?.rmbUsd) || 0;
  const rmbEnUsd = rate > 0 ? rmbTotal / rate : null;
  const grandTotalUsd = rmbEnUsd !== null ? rmbEnUsd + usdTotal : null;

  return (
    <div className="bg-app-panel border border-app-line rounded-xl px-4 py-3 space-y-2">
      <span className="text-sm text-app-dim2 block">{label}</span>

      {/* RMB */}
      {rmbTotal > 0 && (
        <div className="flex justify-between items-center">
          <span className="text-xs text-app-dim">RMB</span>
          <span className="text-base font-semibold text-app-gold">{fmtMoneda(rmbTotal, 'RMB')}</span>
        </div>
      )}

      {/* USD directo */}
      {usdTotal > 0 && (
        <div className="flex justify-between items-center">
          <span className="text-xs text-app-dim">USD</span>
          <span className="text-base font-semibold text-app-gold">{fmtMoneda(usdTotal, 'USD')}</span>
        </div>
      )}

      {/* Separador y conversión */}
      {rmbTotal > 0 && (
        <>
          <div className="border-t border-app-line pt-2">
            <div className="flex items-center gap-2 mb-1.5">
              <span className="text-xs text-app-dim2">Tasa RMB → USD</span>
              <div className="flex items-center gap-1 ml-auto">
                <span className="text-xs text-app-dim">1 USD =</span>
                <input
                  type="number"
                  step="0.01"
                  min="0"
                  value={tasaCambio?.rmbUsd ?? ''}
                  onChange={(e) => setTasaCambio({ ...tasaCambio, rmbUsd: e.target.value })}
                  className="w-20 bg-app-bg border border-app-line rounded-lg px-2 py-1 text-sm text-right"
                />
                <span className="text-xs text-app-dim">RMB</span>
              </div>
            </div>
            {rmbEnUsd !== null && (
              <div className="flex justify-between items-center">
                <span className="text-xs text-app-dim2">RMB en USD</span>
                <span className="text-sm text-app-light">{fmtMoneda(rmbEnUsd, 'USD')}</span>
              </div>
            )}
            {grandTotalUsd !== null && (usdTotal > 0 || rmbEnUsd !== null) && (
              <div className="flex justify-between items-center mt-1 pt-1 border-t border-app-line">
                <span className="text-xs font-semibold text-app-dim2">Total en USD</span>
                <span className="text-lg font-bold text-app-gold">{fmtMoneda(grandTotalUsd, 'USD')}</span>
              </div>
            )}
          </div>
        </>
      )}
    </div>
  );
}

function NuevoPedido({ products = [], setProducts, departamentos = [], tipos = [], setTipos, marcas = [], marcasProveedores = {}, ciudades = [], fabricas = [], factores, suppliers = [], embarcadores = [], tasaCambio, setTasaCambio, origen, borrador, onGuardarBorrador, usuarioActivoNombre, usuarioActivoPrefijo, onCancel, onCreate, onCreateMarca, onCreateFabrica, onCreateCiudad }) {
  const [visor, setVisor] = useState(null);   // fotos a mostrar en pantalla completa
  const [supplierId, setSupplierId] = useState(borrador?.supplierId || suppliers[0]?.id || '');
  const [items, setItems] = useState(borrador?.items || []);
  const [query, setQuery] = useState('');
  const [notas, setNotas] = useState(borrador?.notas || '');
  const [expandedId, setExpandedId] = useState(null);
  const [draftMatrix, setDraftMatrix] = useState({});
  const [draftDestino, setDraftDestino] = useState('H');
  const [showNewProduct, setShowNewProduct] = useState(false);
  const [codigoInicio, setCodigoInicio] = useState(borrador?.codigoInicio || ''); // número de inicio del correlativo (solo China)
  const [embarcadorId, setEmbarcadorId] = useState(borrador?.embarcadorId || ''); // compañía de embarque (USA/Panamá)
  const [ciudadPorDefecto, setCiudadPorDefecto] = useState(borrador?.ciudadPorDefecto || ''); // ciudad elegida en el primer artículo
  const setTiposExterno = (t) => { if (setTipos) setTipos(t); };

  // Marcas del proveedor elegido primero; el resto queda disponible más abajo
  const marcasPriorizadas = (() => {
    if (!supplierId) return marcas;
    const esDelProveedor = (m) => (marcasProveedores[m] || []).includes(supplierId);
    const propias = marcas.filter(esDelProveedor);
    if (propias.length === 0) return marcas;
    return [...propias, ...marcas.filter((m) => !esDelProveedor(m))];
  })();

  // Guardar borrador automáticamente cuando cambian los datos clave del pedido.
  // Uso ref para no depender de la identidad de onGuardarBorrador (evita loop infinito).
  const onGuardarBorradorRef = useRef(onGuardarBorrador);
  useEffect(() => { onGuardarBorradorRef.current = onGuardarBorrador; }, [onGuardarBorrador]);

  useEffect(() => {
    if (!onGuardarBorradorRef.current) return;
    // Solo guardar como borrador si hay artículos capturados (el resto puede tener valores
    // por defecto como supplierId inicial, no es progreso real todavía)
    if (items.length === 0) return;
    const timer = setTimeout(() => {
      onGuardarBorradorRef.current({ supplierId, items, notas, codigoInicio, embarcadorId, ciudadPorDefecto });
    }, 500);
    return () => clearTimeout(timer);
  }, [supplierId, items, notas, codigoInicio, embarcadorId, ciudadPorDefecto]);


  // Solo mostrar resultados cuando el usuario escribe algo en el buscador
  const filtered = query.trim().length === 0 ? [] : products.filter((p) =>
    `${p.codigo} ${p.descripcion} ${p.tipo} ${p.subtipo || ''} ${(p.colores || []).join(' ')}`.toLowerCase().includes(query.toLowerCase())
  );

  const openProduct = (p) => {
    if (expandedId === p.id) { setExpandedId(null); return; }
    setExpandedId(p.id);
    const esCL = p.medida === 'cintura_largo';
    if (esCL) {
      // Cintura/largo se maneja por país (un item por destino)
      const existing = items.find((it) => it.productId === p.id);
      setDraftMatrix(existing ? variantesToMatrixCL(existing.variantes) : {});
      setDraftDestino(existing?.destino || 'H');
    } else {
      // Reconstruir matriz dual desde los items H y G existentes de este producto
      const itemH = items.find((it) => it.productId === p.id && (it.destino || 'H') === 'H');
      const itemG = items.find((it) => it.productId === p.id && it.destino === 'G');
      const dual = {};
      (itemH?.variantes || []).forEach((v) => { dual[dualKey(v.talla, v.color, 'H')] = v.cantidad; });
      (itemG?.variantes || []).forEach((v) => { dual[dualKey(v.talla, v.color, 'G')] = v.cantidad; });
      setDraftMatrix(dual);
    }
  };

  const confirmAdd = (product) => {
    const esCL = product.medida === 'cintura_largo';
    const variantes = esCL ? matrixCLToVariantes(draftMatrix) : matrixToVariantes(draftMatrix);
    if (variantes.length === 0) return;
    setItems((prev) => {
      // Clave única producto+destino
      const filteredPrev = prev.filter((it) => !(it.productId === product.id && (it.destino || 'H') === draftDestino));
      return [
          ...filteredPrev,
          {
            productId: product.id, codigo: product.codigo, descripcion: product.descripcion || '',
            destino: draftDestino, origen: product.origen || origen?.id || '',
            tipo: product.tipo, subtipo: product.subtipo, departamento: product.departamento, marca: product.marca,
            ciudad: product.ciudad || '', fabrica: product.fabrica || '',
            costoMonto: product.costoMonto, costoMoneda: product.costoMoneda,
            ventaLempiras: product.ventaLempiras ?? null,
            foto: product.foto || null, fotos: listaFotos(product), variantes,
          },
        ];
    });
    setExpandedId(null);
  };

  // Agregar capturando ambos países a la vez desde la matriz dual
  const confirmAddDual = (product) => {
    const variantesH = dualMatrixToVariantes(draftMatrix, 'H');
    const variantesG = dualMatrixToVariantes(draftMatrix, 'G');
    if (variantesH.length === 0 && variantesG.length === 0) return;

    setItems((prev) => {
      // Quitar los items previos de este producto (ambos países) para reemplazarlos
      let next = prev.filter((it) => it.productId !== product.id);
      const base = {
        productId: product.id, codigo: product.codigo, descripcion: product.descripcion || '',
        origen: product.origen || origen?.id || '',
        tipo: product.tipo, subtipo: product.subtipo, departamento: product.departamento, marca: product.marca,
        ciudad: product.ciudad || '', fabrica: product.fabrica || '',
        costoMonto: product.costoMonto, costoMoneda: product.costoMoneda,
        ventaLempiras: product.ventaLempiras ?? null, foto: product.foto || null, fotos: listaFotos(product),
      };
      if (variantesH.length > 0) next.push({ ...base, destino: 'H', variantes: variantesH });
      if (variantesG.length > 0) next.push({ ...base, destino: 'G', variantes: variantesG });
      return next;
    });
    setExpandedId(null);
  };

  const removeItem = (productId, destino) => setItems((prev) => prev.filter((it) => !(it.productId === productId && (it.destino || 'H') === (destino || 'H'))));

  const handleNewProductSaved = (nuevos) => {
    // Actualizar catálogo: si un producto con ese código YA existe (segundo pase para el otro
    // país), reusar el existente en vez de duplicarlo.
    setProducts((prev) => {
      const map = new Map(prev.map((p) => [(p.codigo || '').trim().toLowerCase(), p]));
      const merged = [...prev];
      nuevos.forEach((n) => {
        if (!map.has((n.codigo || '').trim().toLowerCase())) merged.push(n);
      });
      return merged;
    });

    // Resolver los IDs finales: si el código ya existía, usamos el ID del producto existente
    const resolverProducto = (n) => {
      const existente = products.find((p) => (p.codigo || '').trim().toLowerCase() === (n.codigo || '').trim().toLowerCase());
      return existente || n;
    };

    // Agregar cada producto nuevo al pedido con sus variantes ya capturadas.
    // La clave es productId+destino, para que podamos tener el mismo artículo con 🇭🇳 y 🇬🇹.
    const itemsNuevos = nuevos
      .filter((p) => p.variantes && p.variantes.length > 0)
      .map((n) => {
        const prod = resolverProducto(n);
        return {
          productId: prod.id,
          codigo: prod.codigo,
          descripcion: prod.descripcion || '',
          destino: n.destinoPedido || 'H',
          origen: prod.origen || origen?.id || '',
          tipo: prod.tipo,
          subtipo: prod.subtipo,
          departamento: prod.departamento,
          marca: prod.marca,
          ciudad: prod.ciudad || '',
          fabrica: prod.fabrica || '',
          costoMonto: prod.costoMonto,
          costoMoneda: prod.costoMoneda,
          ventaLempiras: prod.ventaLempiras ?? null,
          foto: prod.foto || null,
          fotos: listaFotos(prod),
          variantes: n.variantes,
        };
      });

    setItems((prev) => {
      // Evitar duplicados por productId+destino (permite un item por país)
      const key = (it) => `${it.productId}__${it.destino || 'H'}`;
      const nuevosKeys = new Set(itemsNuevos.map(key));
      const sinDuplicados = prev.filter((it) => !nuevosKeys.has(key(it)));
      return [...sinDuplicados, ...itemsNuevos];
    });

    // NO cerrar la ventana: el usuario decide cuándo con "Finalizar" dentro del ProductForm
  };

  const totals = orderTotalsByCurrency(items);
  const canCreate = supplierId && items.length > 0 && (
    !(origen?.id === 'usa' || origen?.id === 'panama') || embarcadorId.trim()
  );

  const handleCreate = () => {
    if (!canCreate) return;
    const esChina = origen?.id === 'china';
    const esUSAoPanama = origen?.id === 'usa' || origen?.id === 'panama';
    const esHonduras = origen?.id === 'honduras';
    
    // Validar que haya embarcador en USA/Panamá
    if (esUSAoPanama && !embarcadorId.trim()) {
      alert('Debes elegir una compañía de embarque para crear el pedido.');
      return;
    }

    const base = {
      supplierId, notas, fecha: todayISO(), status: 'pendiente',
      embarcadorId: esUSAoPanama ? embarcadorId : undefined,
    };

    if (esChina) {
      // China: un solo pedido con items H y G mezclados (el código lleva la letra)
      onCreate({ id: uid(), items, ...base });
      return;
    }

    // USA / Panamá / Honduras: separar en un pedido por destino para tener control
    const itemsH = items.filter((it) => (it.destino || 'H') === 'H');
    const itemsG = items.filter((it) => it.destino === 'G');
    const pedidos = [];
    if (itemsH.length > 0) pedidos.push({ id: uid(), items: itemsH, destinoPedido: 'H', ...base });
    if (itemsG.length > 0) pedidos.push({ id: uid(), items: itemsG, destinoPedido: 'G', ...base });
    onCreate(pedidos);
  };

  if (suppliers.length === 0) {
    return (
      <div className="text-center mt-16 text-sm text-app-dim2">
        Primero agrega un proveedor en la pestaña "Proveedores".
        <div className="mt-4"><button onClick={onCancel} className="text-app-gold font-medium">Volver</button></div>
      </div>
    );
  }

  return (
    <div className="space-y-4">
      <div>
        <label className="text-xs uppercase tracking-wide text-app-dim2 mb-1.5 block">Proveedor</label>
        <select
          value={supplierId}
          onChange={(e) => setSupplierId(e.target.value)}
          className="w-full bg-app-panel border border-app-line rounded-xl px-3 py-3 text-sm focus:outline-none focus:ring-2 focus:ring-amber-400"
        >
          {suppliers.map((s) => <option key={s.id} value={s.id}>{s.nombre}</option>)}
        </select>
      </div>

      {/* Compañía de embarque — solo USA y Panamá */}
      {(origen?.id === 'usa' || origen?.id === 'panama') && (
        <div>
          <label className="text-xs uppercase tracking-wide text-app-dim2 mb-1.5 block">
            Compañía de embarque <span className="text-app-red">*</span>
          </label>
          {embarcadores.length === 0 ? (
            <p className="text-xs text-app-dim2 bg-app-panel border border-app-line rounded-xl px-3 py-3">
              No hay compañías de embarque registradas. El administrador puede agregarlas en Administración 🚢.
            </p>
          ) : (
            <>
              <select
                value={embarcadorId}
                onChange={(e) => setEmbarcadorId(e.target.value)}
                className="w-full bg-app-panel border border-app-line rounded-xl px-3 py-3 text-sm appearance-none"
              >
                <option value="">— Sin compañía de embarque —</option>
                {embarcadores.map((e) => <option key={e.id} value={e.id}>{e.nombre}</option>)}
              </select>
              {embarcadorId && (() => {
                const emb = embarcadores.find((e) => e.id === embarcadorId);
                if (!emb) return null;
                const datos = [emb.contacto, emb.telefono, emb.email].filter(Boolean).join(' · ');
                return datos ? <p className="text-xs text-app-dim2 mt-1">{datos}</p> : null;
              })()}
            </>
          )}
        </div>
      )}

      {/* Código inicial de la compra — solo China */}
      {origen?.id === 'china' && (() => {
        const prefijo = prefijoCorrelativo(usuarioActivoNombre, usuarioActivoPrefijo);
        const numInicio = parseInt(codigoInicio, 10);
        // ¿El número escrito ya existe en otro producto?
        let yaExiste = false;
        if (!isNaN(numInicio) && numInicio > 0) {
          const codigoCompleto = `${prefijo}${String(numInicio).padStart(5, '0')}`;
          yaExiste = products.some((p) => {
            const base = (p.codigoBase || p.codigo || '').trim();
            // Comparar la parte base (sin color ni destino)
            return base.startsWith(codigoCompleto);
          });
        }
        const preview = !isNaN(numInicio) && numInicio > 0
          ? `${prefijo}${String(numInicio).padStart(5, '0')}`
          : null;
        return (
          <div>
            <label className="text-xs uppercase tracking-wide text-app-dim2 mb-1.5 block">
              Código inicial de esta compra <span className="normal-case text-app-dim">(opcional)</span>
            </label>
            <div className="flex items-center gap-2">
              <span className="bg-app-bg border border-app-line rounded-xl px-3 py-3 text-sm font-mono text-app-gold shrink-0">
                {prefijo}
              </span>
              <input
                type="number"
                inputMode="numeric"
                min={1}
                value={codigoInicio}
                onChange={(e) => setCodigoInicio(e.target.value.replace(/\D/g, ''))}
                placeholder="00001"
                className="flex-1 bg-app-panel border border-app-line rounded-xl px-3 py-3 text-sm font-mono focus:outline-none focus:ring-2 focus:ring-amber-400"
              />
            </div>
            {preview && (
              <p className={`text-xs mt-1 ${yaExiste ? 'text-app-red2' : 'text-app-dim'}`}>
                {yaExiste
                  ? `⚠ El código ${preview} ya existe. Se usará el siguiente número libre.`
                  : `El primer artículo iniciará en ${preview} y seguirá el correlativo.`}
              </p>
            )}
            {!preview && (
              <p className="text-xs text-app-dim mt-1">
                Deja vacío para continuar automáticamente desde el último número usado.
              </p>
            )}
          </div>
        );
      })()}

      <div>
        <label className="text-xs uppercase tracking-wide text-app-dim2 mb-1.5 block">Buscar producto</label>
        <div className="flex gap-2">
          <div className="relative flex-1">
            <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-app-dim" />
            <input
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              placeholder="Descripción, código, tipo…"
              className="w-full bg-app-panel border border-app-line rounded-xl pl-9 pr-3 py-3 text-sm focus:outline-none focus:ring-2 focus:ring-amber-400"
            />
          </div>
        </div>

        {!showNewProduct ? (
          <button
            onClick={() => setShowNewProduct(true)}
            className="mt-2 w-full py-2.5 rounded-xl border border-dashed border-app-line3 text-sm text-app-dim2 flex items-center justify-center gap-2 active:bg-app-panel"
          >
            <Plus size={16} /> ¿No existe en el catálogo? Crear producto nuevo
          </button>
        ) : (
          <div className="mt-2">
            <ProductForm
              products={products}
              departamentos={departamentos}
              tipos={tipos}
              marcas={marcasPriorizadas}
              ciudades={ciudades}
              ciudadPorDefecto={ciudadPorDefecto}
              onSetCiudadPorDefecto={setCiudadPorDefecto}
              fabricas={fabricas}
              factores={factores}
              tasaCambio={tasaCambio}
              origen={origen}
              pedidoMode={true}
              usuarioActivoNombre={usuarioActivoNombre}
              usuarioActivoPrefijo={usuarioActivoPrefijo}
              numeroInicioCorrelativo={codigoInicio}
              title="Crear artículo y agregar al pedido"
              onCancel={() => setShowNewProduct(false)}
              onSave={(nuevos) => handleNewProductSaved(nuevos)}
              onUpdateTipos={setTiposExterno}
              onCreateMarca={onCreateMarca}
              onCreateFabrica={onCreateFabrica}
              onCreateCiudad={onCreateCiudad}
            />
          </div>
        )}

        {query.trim().length > 0 && (
          <div className="mt-2 rounded-xl border border-app-line divide-y divide-app-line overflow-hidden">
            {filtered.length === 0 && <p className="text-xs text-app-dim p-3">Sin resultados para "{query}".</p>}
            {filtered.map((p) => {
            const inOrderItems = items.filter((it) => it.productId === p.id);
            const inOrder = inOrderItems.length > 0;
            const inOrderPzs = inOrderItems.reduce((s, it) => s + sumVariantes(it.variantes), 0);
            const esCL = p.medida === 'cintura_largo';
            const tallas = p.tallas?.length > 0 ? p.tallas : [...new Set((p.variantes || []).map((v) => v.talla))];
            const cinturas = p.cinturas?.length > 0 ? p.cinturas : [...new Set((p.variantes || []).map((v) => v.cintura).filter(Boolean))];
            const largos = p.largos?.length > 0 ? p.largos : [...new Set((p.variantes || []).map((v) => v.largo).filter(Boolean))];
            return (
              <div key={p.id}>
                <button onClick={() => openProduct(p)} className="w-full flex items-center gap-3 px-3 py-2.5 text-left active:bg-app-active">
                  {listaFotos(p).length ? (
                    <img
                      src={urlFoto(listaFotos(p)[0])}
                      alt={p.descripcion}
                      className="w-10 h-10 object-cover rounded-md shrink-0 cursor-zoom-in"
                      onClick={(e) => { e.stopPropagation(); setVisor({ fotos: listaFotos(p), titulo: p.descripcion || p.codigo }); }}
                    />
                  ) : (
                    <div className="w-10 h-10 rounded-md bg-app-bg border border-app-line flex items-center justify-center shrink-0">
                      <ImageOff size={13} className="text-app-dim3" />
                    </div>
                  )}
                  <div className="min-w-0 flex-1">
                    <p className="text-sm truncate">{p.descripcion}</p>
                    <p className="text-xs text-app-dim truncate">
                      {p.codigo} · {p.tipo}{p.subtipo ? ` (${p.subtipo})` : ''} · {fmtMoneda(p.costoMonto, p.costoMoneda)}
                      {inOrder ? ` · ${inOrderPzs} pzs en pedido` : ''}
                    </p>
                  </div>
                  <ChevronDown size={16} className={`text-app-dim shrink-0 transition-transform ${expandedId === p.id ? 'rotate-180' : ''}`} />
                </button>
                {expandedId === p.id && (
                  <div className="px-3 pb-3 bg-app-panel">
                    {esCL ? (
                      <>
                        <div className="mb-2">
                          <label className="text-xs uppercase tracking-wide text-app-dim2 mb-1.5 block">Destino</label>
                          <div className="flex gap-2">
                            {DESTINOS.map((d) => (
                              <button
                                key={d.id}
                                type="button"
                                onClick={() => setDraftDestino(d.id)}
                                className={`flex-1 py-2 rounded-lg text-sm font-medium border transition flex items-center justify-center gap-1.5 ${
                                  draftDestino === d.id
                                    ? 'bg-app-gold text-app-bg border-app-gold'
                                    : 'bg-app-bg border-app-line text-app-dim2'
                                }`}
                              >
                                <span>{d.emoji}</span> {d.label} ({d.id})
                              </button>
                            ))}
                          </div>
                        </div>
                        <CinturaLargoMatrix
                          cinturas={cinturas}
                          largos={largos}
                          colores={p.colores || []}
                          values={draftMatrix}
                          onChange={setDraftMatrix}
                        />
                        <button
                          onClick={() => confirmAdd(p)}
                          className="mt-2 w-full py-2 rounded-lg bg-app-gold text-app-bg text-sm font-semibold"
                        >
                          Agregar al pedido → {DESTINOS.find((d) => d.id === draftDestino)?.emoji}
                        </button>
                      </>
                    ) : (
                      <>
                        <p className="text-xs text-app-dim2 mb-2">Captura las cantidades para cada país. Se agregarán al pedido solo los que tengan cantidad.</p>
                        <VariantMatrixDual
                          tallas={tallas}
                          colores={p.colores || []}
                          values={draftMatrix}
                          onChange={setDraftMatrix}
                        />
                        <button
                          onClick={() => confirmAddDual(p)}
                          className="mt-2 w-full py-2 rounded-lg bg-app-gold text-app-bg text-sm font-semibold"
                        >
                          Agregar al pedido 🇭🇳 🇬🇹
                        </button>
                      </>
                    )}
                  </div>
                )}
              </div>
            );
          })}
          </div>
        )}
      </div>

      {items.length > 0 && (
        <div>
          <label className="text-xs uppercase tracking-wide text-app-dim2 mb-1.5 block">Productos en el pedido</label>
          <div className="space-y-2">
            {items.map((it) => (
              <div key={`${it.productId}-${it.destino || 'H'}`} className="bg-app-panel border border-app-line rounded-xl p-3">
                <div className="flex items-center gap-3">
                  <div className="min-w-0 flex-1">
                    <p className="text-sm truncate">{it.descripcion}</p>
                    <p className="text-xs text-app-dim">
                      <span className="font-mono text-app-gold">{codigoConDestino(it)}</span>
                      {it.destino ? ` ${destinoInfo(it.destino).emoji}` : ''}{it.subtipo ? ` (${it.subtipo})` : ''} · {sumVariantes(it.variantes)} pzs
                    </p>
                  </div>
                  <span className="text-xs text-app-dim2 shrink-0">{fmtMoneda(itemTotal(it), it.costoMoneda)}</span>
                  <BotonBorrar onConfirm={() => removeItem(it.productId, it.destino)} size={16} />
                </div>
                <p className="text-xs text-app-dim mt-1.5">
                  {(it.variantes || []).map((v) => `${varLabel(v)} / ${v.color}: ${v.cantidad}`).join(' · ')}
                </p>
              </div>
            ))}
          </div>
        </div>
      )}

      <div>
        <label className="text-xs uppercase tracking-wide text-app-dim2 mb-1.5 block">Notas (opcional)</label>
        <textarea
          value={notas}
          onChange={(e) => setNotas(e.target.value)}
          rows={2}
          placeholder="Fecha de entrega deseada, instrucciones especiales…"
          className="w-full bg-app-panel border border-app-line rounded-xl px-3 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-amber-400 resize-none"
        />
      </div>

      <TotalesConversion
        totals={totals}
        tasaCambio={tasaCambio}
        setTasaCambio={setTasaCambio}
        label="Total estimado (costo)"
      />

      <div className="flex gap-3 pt-1">
        <button onClick={onCancel} className="flex-1 py-3 rounded-xl border border-app-line text-sm font-medium text-app-dim2 active:bg-app-panel">Cancelar</button>
        <button
          onClick={handleCreate}
          disabled={!canCreate}
          className="flex-1 py-3 rounded-xl bg-app-gold text-app-bg text-sm font-semibold disabled:opacity-40 active:scale-95 transition"
        >
          Crear pedido
        </button>
      </div>

      {visor && (
        <VisorFotos
          fotos={visor.fotos}
          titulo={visor.titulo}
          onCerrar={() => setVisor(null)}
        />
      )}
    </div>
  );
}

// ---------- Detalle pedido ----------
function DetallePedido({ order, supplier, onBack, onUpdateStatus, tasaCambio, setTasaCambio, onDuplicar, hayBorradorPendiente, borradorAjenoBloquea = false, empresa, embarcadores = [] }) {
  const [visor, setVisor] = useState(null);   // fotos a mostrar en pantalla completa
  const [confirmDuplicar, setConfirmDuplicar] = useState(false);
  const [pdfHtml, setPdfHtml] = useState(null);   // documento a mostrar
  const [pdfModo, setPdfModo] = useState('proveedor');
  const pdfRef = useRef(null);
  const totals = orderTotalsByCurrency(order.items);

  const exportExcel = () => {
    const rows = [];
    const enlacesFoto = [];   // URL de la foto de cada fila, para el hipervínculo
    const itemsOrdenados = [...order.items].sort((a, b) => (a.destino || 'H').localeCompare(b.destino || 'H'));
    // Nombre del destino: prioriza el destino del item; si no, usa el destinoPedido del pedido
    const nombreDestino = (it) => {
      const d = it.destino || order.destinoPedido;
      if (d === 'G') return 'Afiliada';
      if (d === 'H') return 'Honduras';
      return '';
    };
    const nombreOrigen = (oid) => (ORIGENES.find((o) => o.id === oid)?.label || oid || '');
    itemsOrdenados.forEach((it) => {
      const enlaceFoto = urlFoto(listaFotos(it)[0], false) || '';
      (it.variantes || []).forEach((v) => {
        rows.push({
          Código: codigoConDestino(it), Destino: nombreDestino(it), Origen: nombreOrigen(it.origen || order.origen),
          Descripción: it.descripcion, Marca: it.marca || '', Departamento: it.departamento, Tipo: it.tipo, Subtipo: it.subtipo || '',
          Talla: v.cintura ? '' : v.talla, Cintura: v.cintura || '', Largo: v.largo || '',
          Color: v.color, Cantidad: v.cantidad,
          'Costo unitario': it.costoMonto, Moneda: it.costoMoneda, Subtotal: v.cantidad * it.costoMonto,
          'Venta HNL': it.ventaLempiras ?? '',
          Ciudad: it.ciudad || '', Fábrica: it.fabrica || '',
          Foto: enlaceFoto ? 'Ver foto' : '',
        });
        enlacesFoto.push(enlaceFoto);   // misma posición que la fila
      });
    });
    rows.push({});
    Object.entries(totals).forEach(([m, v]) => rows.push({ Código: 'TOTAL', Moneda: m, Subtotal: v }));

    // Cabecera con los datos generales del pedido
    const etiquetaEstado = ESTADOS_PEDIDO[order.status]?.label || order.status || '';
    const piezasTotales = (order.items || []).reduce((s, it) => s + sumVariantes(it.variantes), 0);
    const ws = XLSX.utils.aoa_to_sheet([
      ['PEDIDO', order.numero || ''],
      ['Proveedor', supplier ? supplier.nombre : 'Proveedor eliminado'],
      ['Origen', nombreOrigen(order.origen || order.items[0]?.origen)],
      ['Fecha', order.fecha || ''],
      ['Comprador', order.creadoPor || ''],
      ...(order.embarcadorId ? [['Embarque', (embarcadores.find((e) => e.id === order.embarcadorId)?.nombre) || '']] : []),
      ...(order.embarcadorId && embarcadores.find((e) => e.id === order.embarcadorId)?.direccion
        ? [['Dirección embarque', embarcadores.find((e) => e.id === order.embarcadorId).direccion]] : []),
      ['Estado', etiquetaEstado],
      ['Destino', order.destinoPedido ? destinoInfo(order.destinoPedido).label : 'Mixto'],
      ['Piezas', piezasTotales],
      ['Notas', order.notas || ''],
      [],
    ]);
    XLSX.utils.sheet_add_json(ws, rows, { origin: -1 });

    // Convertir la columna "Foto" en enlaces clicables
    const rango = XLSX.utils.decode_range(ws['!ref']);
    let colFoto = -1, filaEncabezado = -1;
    for (let R = rango.s.r; R <= rango.e.r && colFoto === -1; R++) {
      for (let C = rango.s.c; C <= rango.e.c; C++) {
        const celda = ws[XLSX.utils.encode_cell({ r: R, c: C })];
        if (celda && celda.v === 'Foto') { colFoto = C; filaEncabezado = R; break; }
      }
    }
    if (colFoto !== -1) {
      enlacesFoto.forEach((url, i) => {
        if (!url) return;
        const ref = XLSX.utils.encode_cell({ r: filaEncabezado + 1 + i, c: colFoto });
        const celda = ws[ref];
        if (!celda) return;
        celda.l = { Target: url, Tooltip: 'Abrir foto del producto' };
        celda.s = { font: { color: { rgb: '0563C1' }, underline: true } };
      });
    }

    ws['!cols'] = [
      { wch: 16 }, { wch: 11 }, { wch: 10 }, { wch: 28 }, { wch: 14 }, { wch: 20 }, { wch: 14 }, { wch: 14 },
      { wch: 10 }, { wch: 8 }, { wch: 8 }, { wch: 14 }, { wch: 9 }, { wch: 13 }, { wch: 8 }, { wch: 12 },
      { wch: 11 }, { wch: 16 }, { wch: 18 }, { wch: 10 },
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Pedido');
    descargarLibro(wb, `${order.numero || 'pedido'}_${supplier ? (supplier.nombre || '').replace(/\s+/g, '_') : order.id}_${order.fecha}.xlsx`);
  };

  // ---------- PDF del pedido: agrupado por artículo con matriz de tallas × colores ----------
  // modo 'proveedor' → sin datos internos (departamento, ciudad, venta, margen)
  // modo 'interno'   → todo
  const exportPDF = (modo = 'proveedor') => {
    const esInterno = modo === 'interno';
    const esc = (t) => String(t ?? '').replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));

    // Construye la tabla de cantidades de un artículo
    // Tabla de líneas: una fila por artículo con estilo, descripción, cantidad, precio y total
    const filaDe = (it) => {
      const pzs = sumVariantes(it.variantes);
      const total = pzs * (it.costoMonto || 0);
      // Detalle breve de colores/tallas debajo de la descripción (sin ocupar mucho)
      const colores = [...new Set((it.variantes || []).map((v) => v.color))];
      const detalle = colores.length > 0 ? `<div class="lin-det">${esc(colores.join(', '))}</div>` : '';
      const foto = urlFoto(listaFotos(it)[0], false);
      const celdaFoto = foto
        ? `<td class="foto"><img src="${esc(foto)}" alt=""></td>`
        : `<td class="foto"><div class="sin-foto"></div></td>`;
      return `
        <tr>
          ${celdaFoto}
          <td class="cod">${esc(codigoConDestino(it))}</td>
          <td class="desc">${esc(it.descripcion || '')}${detalle}</td>
          <td class="num">${pzs}</td>
          <td class="num">${fmtMoneda(it.costoMonto, it.costoMoneda)}</td>
          <td class="num">${fmtMoneda(total, it.costoMoneda)}</td>
        </tr>`;
    };

    // Agrupar por destino cuando el pedido mezcla países
    const destinos = [...new Set((order.items || []).map((it) => it.destino || order.destinoPedido || 'H'))];
    const mezcla = destinos.length > 1;

    const tablaDe = (items) => `
      <table class="lineas">
        <thead>
          <tr>
            <th class="foto">Foto</th>
            <th class="cod">Estilo</th>
            <th class="desc">Descripción</th>
            <th class="num">Cant.</th>
            <th class="num">Precio unit.</th>
            <th class="num">Total</th>
          </tr>
        </thead>
        <tbody>${items.map(filaDe).join('')}</tbody>
      </table>`;

    let cuerpo = '';
    if (mezcla) {
      destinos.forEach((d) => {
        const items = (order.items || []).filter((it) => (it.destino || order.destinoPedido || 'H') === d);
        if (items.length === 0) return;
        const pzsD = items.reduce((s, it) => s + sumVariantes(it.variantes), 0);
        cuerpo += `<h2 class="dest">${destinoInfo(d).label} — ${items.length} artículos · ${pzsD} pzs</h2>`;
        cuerpo += tablaDe(items);
      });
    } else {
      cuerpo = tablaDe(order.items);
    }

    const piezasTot = (order.items || []).reduce((s, it) => s + sumVariantes(it.variantes), 0);
    const totalesTxt = Object.entries(totals)
      .map(([m, v]) => `<div class="tot-linea"><span>Total ${m}</span><strong>${fmtMoneda(v, m)}</strong></div>`)
      .join('');

    const contacto = supplier
      ? [supplier.contacto, supplier.email, supplier.telefono].filter(Boolean).join(' · ')
      : '';

    const html = `
      <html>
        <head>
          <meta charset="utf-8" />
          <title>${esc(order.numero || 'Pedido')} - ${esc(supplier ? supplier.nombre : '')}</title>
          <style>
            @page { size: A4 portrait; margin: 14mm; }
            /* Conservar colores e imágenes al imprimir o guardar como PDF */
            * { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
            table.lineas tr { break-inside: avoid; page-break-inside: avoid; }
            * { box-sizing: border-box; }
            body { font-family: -apple-system, "Helvetica Neue", Arial, sans-serif; color: #1a1a1a; font-size: 11px; margin: 0; }
            .cab { border-bottom: 2px solid #1a1a1a; padding-bottom: 10px; margin-bottom: 14px; }
            .cab-top { display: flex; justify-content: space-between; align-items: flex-start; gap: 14px; }
            .cab-izq { display: flex; align-items: flex-start; gap: 12px; }
            .logo { max-height: 46px; max-width: 210px; object-fit: contain; flex: none; }
            h1 { font-size: 17px; margin: 0 0 2px; letter-spacing: .5px; }
            .num-pedido { font-size: 15px; font-weight: bold; font-family: monospace; }
            .tipo-doc { font-size: 9px; color: #777; text-transform: uppercase; letter-spacing: 1px; }
            .datos { display: flex; gap: 28px; margin-top: 10px; font-size: 11px; }
            .datos div { line-height: 1.6; }
            .etq { color: #777; display: inline-block; min-width: 62px; }
            .notas { margin-top: 8px; padding: 6px 8px; background: #f6f6f6; border-left: 3px solid #999; font-size: 11px; }
            h2.dest { font-size: 12px; margin: 16px 0 8px; padding: 5px 8px; background: #eee; border-radius: 3px; }
            table.lineas { border-collapse: collapse; width: 100%; margin: 4px 0 10px; font-size: 11px; }
            table.lineas thead th { background: #1a1a1a; color: #fff; font-weight: 600; padding: 7px 8px; text-align: left; }
            table.lineas tbody td { border-bottom: 1px solid #e5e5e5; padding: 7px 8px; vertical-align: top; }
            table.lineas tbody tr:nth-child(even) td { background: #fafafa; }
            table.lineas .cod { font-family: monospace; font-weight: bold; white-space: nowrap; }
            table.lineas .desc { color: #333; }
            table.lineas .num { text-align: right; white-space: nowrap; }
            table.lineas th.num { text-align: right; }
            table.lineas .foto { width: 54px; padding: 4px 6px; }
            table.lineas .foto img {
              width: 46px; height: 46px; object-fit: cover;
              border-radius: 4px; border: 1px solid #ddd; display: block;
            }
            table.lineas .sin-foto {
              width: 46px; height: 46px; border-radius: 4px;
              border: 1px dashed #ddd; background: #f5f5f5;
            }
            .lin-det { color: #888; font-size: 9px; margin-top: 2px; }
            .resumen { margin-top: 16px; border-top: 2px solid #1a1a1a; padding-top: 9px; display: flex; justify-content: space-between; align-items: flex-start; }
            .resumen .izq { font-size: 11px; color: #555; line-height: 1.7; }
            .tot-linea { display: flex; justify-content: space-between; gap: 22px; font-size: 13px; padding: 2px 0; }
            .firma { margin-top: 34px; display: flex; gap: 50px; font-size: 10px; color: #777; }
            .docs { margin-top: 18px; padding: 11px 13px; border: 1px solid #ddd; border-radius: 4px; background: #fafafa; page-break-inside: avoid; }
            .docs-titulo { font-size: 11px; font-weight: bold; text-transform: uppercase; letter-spacing: .5px; margin-bottom: 5px; }
            .docs-intro { font-size: 10px; margin: 0 0 4px; color: #444; }
            .docs-lista { font-size: 10px; margin: 0 0 6px; padding-left: 18px; color: #333; line-height: 1.5; }
            .docs-lista li { margin-bottom: 2px; }
            .docs-nota { font-size: 10px; margin: 0; color: #444; border-top: 1px dashed #ccc; padding-top: 6px; }
            .firma div { border-top: 1px solid #999; padding-top: 4px; width: 190px; text-align: center; }
            .pie { margin-top: 20px; font-size: 9px; color: #999; text-align: center; }
          </style>
        </head>
        <body>
          <div class="cab">
            <div class="cab-top">
              <div class="cab-izq">
                ${empresa?.logo ? `<img src="${empresa.logo}" class="logo" alt="" />` : ''}
                <div>
                <div class="tipo-doc">${esInterno ? 'Hoja de pedido — copia interna' : 'Hoja de pedido'}</div>
                <h1>${esc(supplier ? supplier.nombre : 'Hoja de pedido')}</h1>
                ${contacto ? `<div style="color:#666">${esc(contacto)}</div>` : ''}
                </div>
              </div>
              <div style="text-align:right">
                <div class="tipo-doc">Pedido</div>
                <div class="num-pedido">${esc(order.numero || '—')}</div>
              </div>
            </div>
            <div class="datos">
              <div>
                <div><span class="etq">Fecha:</span> ${esc(order.fecha || '')}</div>
                <div><span class="etq">Origen:</span> ${esc(ORIGENES.find((o) => o.id === (order.origen || order.items[0]?.origen))?.label || '')}</div>
                <div><span class="etq">Destino:</span> ${mezcla ? 'Honduras y Afiliada' : esc(destinoInfo(destinos[0]).label)}</div>
              </div>
              <div>
                <div><span class="etq">Comprador:</span> ${esc(order.creadoPor || '—')}</div>
                <div><span class="etq">Estado:</span> ${esc(ESTADOS_PEDIDO[order.status]?.label || '')}</div>
                <div><span class="etq">Artículos:</span> ${(order.items || []).length} &nbsp;·&nbsp; <strong>${piezasTot} pzs</strong></div>
                ${(() => {
                  const emb = embarcadores.find((e) => e.id === order.embarcadorId);
                  if (!emb) return '';
                  const cont = [emb.contacto, emb.telefono, emb.email].filter(Boolean).join(' · ');
                  return `<div><span class="etq">Embarque:</span> ${esc(emb.nombre)}${cont ? ` <span style="color:#999">(${esc(cont)})</span>` : ''}${emb.direccion ? `<div style="margin-left:62px;color:#999">${esc(emb.direccion)}</div>` : ''}</div>`;
                })()}
              </div>
            </div>
            ${order.notas ? `<div class="notas"><strong>Notas:</strong> ${esc(order.notas)}</div>` : ''}
          </div>

          ${cuerpo}

          <div class="resumen">
            <div class="izq">
              ${(order.items || []).length} artículos<br/>
              <strong>${piezasTot} piezas en total</strong>
            </div>
            <div>${totalesTxt}</div>
          </div>

          ${(() => {
            const emb = embarcadores.find((e) => e.id === order.embarcadorId);
            if (!emb) return ''; // solo cuando hay compañía de embarque
            const nombreEmb = emb.nombre || 'la compañía de embarque';
            return `
            <div class="docs">
              <div class="docs-titulo">Documentación requerida</div>
              <p class="docs-intro">Para revisión previa, el proveedor deberá enviar:</p>
              <ol class="docs-lista">
                <li>Factura original con detalle de modelo, origen, marca, término de compra, días de crédito y tejido (si aplica), firmada y sellada.</li>
                <li>Copia de la factura (sin firma ni sello).</li>
                <li>Instructivo.</li>
                <li>Lista de empaque (Packing List).</li>
                <li>Lista de precios y carta de crédito indicando término de compra y días de crédito, firmadas y selladas.</li>
                <li>Orden de compra.</li>
              </ol>
              <p class="docs-nota">Al momento de la entrega de la mercancía, presentar la documentación anterior en un folder identificado a nombre de <strong>${esc(nombreEmb)}</strong>.</p>
            </div>`;
          })()}

          ${esInterno ? '' : `
          <div class="firma">
            <div>Comprador</div>
            <div>Proveedor</div>
          </div>`}

          <div class="pie">${esInterno ? 'Documento interno — no enviar al proveedor' : `Pedido ${esc(order.numero || '')} · ${esc(order.fecha || '')}`}</div>
        </body>
      </html>`;

    // Se muestra dentro de la app: abrir ventanas nuevas está bloqueado en este entorno
    setPdfHtml(html);
    setPdfModo(modo);
  };

  // Imprime solo el documento (el iframe), no la pantalla de la app
  const imprimirPdf = () => {
    const marco = pdfRef.current;
    if (!marco || !marco.contentWindow) return;
    marco.contentWindow.focus();
    marco.contentWindow.print();
  };

  return (
    <div className="space-y-4">
      <button onClick={onBack} className="text-sm text-app-dim2 flex items-center gap-1">← Volver a pedidos</button>

      <div className="bg-app-panel border border-app-line rounded-2xl p-4">
        <div className="flex items-center justify-between mb-2">
          <div className="flex items-center gap-2">
            <StatusPill status={order.status} />
            {order.numero && <span className="text-xs font-mono text-app-gold bg-app-bg border border-app-line rounded-full px-2 py-0.5">{order.numero}</span>}
          </div>
          <span className="text-xs text-app-dim">{order.fecha}</span>
        </div>
        <p className="font-semibold text-base">{supplier ? supplier.nombre : 'Proveedor eliminado'}</p>
        {order.destinoPedido && (
          <p className="text-sm text-app-gold mt-1">
            {destinoInfo(order.destinoPedido).emoji} Destino: {destinoInfo(order.destinoPedido).label}
          </p>
        )}
        {supplier && <p className="text-xs text-app-dim2 mt-0.5">{supplier.contacto} · {supplier.telefono}</p>}
        {order.creadoPor && <p className="text-xs text-app-dim mt-0.5">Creado por {order.creadoPor}</p>}
        {order.notas && <p className="text-sm text-app-light mt-2">{order.notas}</p>}
      </div>

      <div className="space-y-4">
        {DESTINOS.filter((d) => (order.items || []).some((it) => (it.destino || 'H') === d.id)).map((d) => {
          const itemsDestino = (order.items || []).filter((it) => (it.destino || 'H') === d.id);
          const totalPzs = itemsDestino.reduce((sum, it) => sum + sumVariantes(it.variantes), 0);
          return (
            <div key={d.id}>
              <div className="flex items-center gap-2 mb-2">
                <span className="text-base">{d.emoji}</span>
                <span className="text-sm font-semibold text-app-white">{d.label}</span>
                <span className="text-xs text-app-dim2 bg-app-panel border border-app-line rounded-full px-2 py-0.5">
                  {itemsDestino.length} art · {totalPzs} pzs
                </span>
              </div>
              <div className="space-y-2">
                {itemsDestino.map((it) => (
                  <div key={it.productId} className="bg-app-panel border border-app-line rounded-xl p-3">
                    <div className="flex items-center gap-3">
                      {listaFotos(it).length ? (
                        <img
                          src={urlFoto(listaFotos(it)[0])}
                          alt={it.descripcion}
                          className="w-11 h-11 object-cover rounded-lg shrink-0 cursor-zoom-in"
                          onClick={(e) => { e.stopPropagation(); setVisor({ fotos: listaFotos(it), titulo: it.descripcion || it.codigo }); }}
                        />
                      ) : (
                        <div className="w-11 h-11 rounded-lg bg-app-bg border border-app-line flex items-center justify-center shrink-0">
                          <ImageOff size={14} className="text-app-dim3" />
                        </div>
                      )}
                      <div className="min-w-0 flex-1">
                        <p className="text-sm font-mono truncate">{it.descripcion}</p>
                        <p className="text-xs text-app-dim">
                          <span className="font-mono text-app-gold">{codigoConDestino(it)}</span>
                          {it.subtipo ? ` (${it.subtipo})` : ''} · {sumVariantes(it.variantes)} pzs{it.marca ? ` · ${it.marca}` : ''}
                        </p>
                      </div>
                      <p className="text-xs text-app-gold shrink-0">{fmtMoneda(itemTotal(it), it.costoMoneda)}</p>
                    </div>
                    <div className="mt-2 pt-2 border-t border-app-line space-y-0.5">
                      {(it.variantes || []).map((v) => (
                        <div key={variantKey(v.talla, v.color)} className="flex justify-between text-xs">
                          <span className="text-app-dim2">{varLabel(v)} · {v.color}</span>
                          <span>{v.cantidad} × {fmtMoneda(it.costoMonto, it.costoMoneda)}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                ))}
              </div>
            </div>
          );
        })}
      </div>

      <TotalesConversion
        totals={totals}
        tasaCambio={tasaCambio}
        setTasaCambio={setTasaCambio}
        label="Total (costo)"
      />

      <div>
        <label className="text-xs uppercase tracking-wide text-app-dim2 mb-1.5 block">Estado</label>
        <div className="flex gap-2">
          {['pendiente', 'enviado', 'recibido'].map((s) => (
            <button
              key={s}
              onClick={() => onUpdateStatus(s)}
              className={`flex-1 py-2 rounded-lg text-xs font-medium border transition ${order.status === s ? 'bg-app-gold text-app-bg border-app-gold' : 'border-app-line text-app-dim2'}`}
            >
              {s[0].toUpperCase() + s.slice(1)}
            </button>
          ))}
        </div>
      </div>

      <div className="flex gap-3 pt-1">
        <button onClick={exportExcel} className="flex-1 py-3 rounded-xl border border-app-line text-sm font-medium flex items-center justify-center gap-2 active:bg-app-panel">
          <FileDown size={16} /> Excel
        </button>
        <button onClick={() => exportPDF('proveedor')} className="flex-1 py-3 rounded-xl bg-app-gold text-app-bg text-sm font-semibold flex items-center justify-center gap-2 active:scale-95 transition">
          <FileDown size={16} /> Hoja proveedor
        </button>
      </div>
      <button
        onClick={() => exportPDF('interno')}
        className="w-full py-2.5 rounded-xl border border-app-line text-xs text-app-dim2 flex items-center justify-center gap-1.5 active:bg-app-panel"
      >
        <FileDown size={14} /> Hoja interna (con precios de venta)
      </button>

      {/* Vista previa del documento — se muestra dentro de la app */}
      {pdfHtml && (
        <div className="fixed inset-0 z-50 flex flex-col" style={{ backgroundColor: 'rgba(0,0,0,0.85)' }}>
          <div className="flex items-center justify-between gap-2 px-4 py-3 bg-app-panel border-b border-app-line">
            <div className="min-w-0">
              <p className="text-sm font-semibold truncate">
                {pdfModo === 'interno' ? 'Hoja de pedido — interna' : 'Hoja de pedido'}
              </p>
              <p className="text-xs text-app-dim2 truncate">{order.numero || 'Pedido'}</p>
            </div>
            <div className="flex items-center gap-2 shrink-0">
              <button
                onClick={imprimirPdf}
                className="px-4 py-2 rounded-lg bg-app-gold text-app-bg text-sm font-semibold"
              >
                Imprimir / Guardar PDF
              </button>
              <button
                onClick={() => setPdfHtml(null)}
                className="p-2 rounded-lg border border-app-line text-app-dim2"
                aria-label="Cerrar"
              >
                <X size={16} />
              </button>
            </div>
          </div>
          <iframe
            ref={pdfRef}
            srcDoc={pdfHtml}
            title="Documento del pedido"
            className="flex-1 w-full bg-white"
          />
          <p className="text-xs text-center text-app-dim2 py-2 bg-app-panel border-t border-app-line">
            En el diálogo de impresión elige <strong>Guardar como PDF</strong> para tener el archivo.
          </p>
        </div>
      )}

      {/* Duplicar pedido — útil para reposiciones al mismo proveedor */}
      {onDuplicar && (
        <div className="pt-1">
          {borradorAjenoBloquea ? (
            <div className="bg-app-panel border border-app-line rounded-xl p-3">
              <p className="text-xs text-app-dim2">
                ⏳ No puedes duplicar ahora: otro comprador tiene un pedido sin terminar en este país.
              </p>
            </div>
          ) : !confirmDuplicar ? (
            <button
              onClick={() => setConfirmDuplicar(true)}
              className="w-full py-3 rounded-xl border border-app-line text-sm font-medium text-app-sky flex items-center justify-center gap-2 active:bg-app-panel"
            >
              <Copy size={16} /> Duplicar este pedido
            </button>
          ) : (
            <div className="bg-app-blue border border-app-line2 rounded-xl p-3 space-y-2">
              <p className="text-xs text-app-light">
                Se abrirá un pedido nuevo con los mismos artículos y cantidades, listo para que ajustes lo que necesites.
                {hayBorradorPendiente && (
                  <span className="block text-app-gold mt-1">⚠ Tienes un pedido sin terminar en este origen. Se reemplazará.</span>
                )}
              </p>
              <div className="flex gap-2">
                <button onClick={() => setConfirmDuplicar(false)} className="flex-1 py-2 rounded-lg border border-app-line text-xs text-app-dim2">
                  Cancelar
                </button>
                <button onClick={() => { setConfirmDuplicar(false); onDuplicar(); }} className="flex-1 py-2 rounded-lg bg-app-sky text-app-bg text-xs font-semibold">
                  Duplicar
                </button>
              </div>
            </div>
          )}
        </div>
      )}

      {visor && (
        <VisorFotos
          fotos={visor.fotos}
          titulo={visor.titulo}
          onCerrar={() => setVisor(null)}
        />
      )}
    </div>
  );
}

// ---------- Catálogo ----------
const emptyForm = () => ({
  codigo: '', descripcion: '', departamento: '', tipo: '', subtipo: '', marca: '', genero: '', origen: '',
  ciudad: '', fabrica: '',
  colores: [], colorInput: '',
  costoMonto: '', costoMoneda: 'USD', ventaLempiras: '', foto: null, fotos: [],
});

// Devuelve el prefijo de correlativo (Letra + Año + Mes), ej. "Y266"
function prefijoCorrelativo(nombreComprador, prefijoComprador) {
  const now = new Date();
  const anio = String(now.getFullYear()).slice(-2);   // '26'
  const mes = String(now.getMonth());                  // julio = '6'
  let inicial;
  if (prefijoComprador && prefijoComprador.trim()) {
    inicial = prefijoComprador.trim().toUpperCase();
  } else if (nombreComprador && nombreComprador.trim()) {
    inicial = nombreComprador.trim().charAt(0).toUpperCase();
  } else {
    inicial = 'Y';
  }
  return `${inicial}${anio}${mes}`;
}

// Código correlativo: Y (comprador) + año 2 dígitos + mes + correlativo de 5 dígitos
// Ejemplo en julio 2026: Y26600001, Y26600002…
// Si se pasa `numeroInicio`, el correlativo arranca desde ahí (si está libre); de lo
// contrario sigue automático desde el último número usado con ese prefijo.
function generarCodigoCorrelativo(products, nombreComprador, prefijoComprador, numeroInicio) {
  const prefijo = prefijoCorrelativo(nombreComprador, prefijoComprador);
  let max = 0;
  const usados = new Set();
  (products || []).forEach((p) => {
    const base = (p.codigoBase || p.codigo || '').trim();
    if (base.startsWith(prefijo)) {
      const num = parseInt(base.slice(prefijo.length), 10);
      if (!isNaN(num)) {
        if (num > max) max = num;
        usados.add(num);
      }
    }
  });
  // Si el usuario fijó un número de inicio, buscar desde ahí el primero libre
  const inicio = parseInt(numeroInicio, 10);
  if (!isNaN(inicio) && inicio > 0) {
    let n = inicio;
    while (usados.has(n)) n++; // saltar los ya ocupados
    return `${prefijo}${String(n).padStart(5, '0')}`;
  }
  return `${prefijo}${String(max + 1).padStart(5, '0')}`;
}

const GENEROS = ['Ho.', 'Da.', 'No.', 'Na.', 'Be.'];

// Mapa de departamento → género sugerido
const DEPTO_GENERO = {
  'ROPA DE CABALLEROS': 'Ho.',
  'ROPA DEPORTIVA DE CABALLEROS': 'Ho.',
  'ROPA INT. DE CABALLERO': 'Ho.',
  'ROPA INTERIOR DE CABALLERO': 'Ho.',
  'ACCESORIOS DE CABALLERO': 'Ho.',
  'ZAPATO DE CABALLERO': 'Ho.',
  'ROPA DAMA': 'Da.',
  'ROPA AMERICANA DAMA': 'Da.',
  'ROPA DEPORTIVA DE DAMAS': 'Da.',
  'ROPA INT. DAMA': 'Da.',
  'ROPA INTERIOR DAMA': 'Da.',
  'ACCESORIOS DAMA': 'Da.',
  'ZAPATO DE DAMA': 'Da.',
  'ROPA DE NIÑOS': 'No.',
  'ROPA INTERIOR DE NINOS': 'No.',
  'ROPA JUVENIL NINOS': 'No.',
  'ZAPATO DE NIÑO': 'No.',
  'ROPA DE NINAS': 'Na.',
  'ROPA INTERIOR DE NINAS': 'Na.',
  'ROPA JUVENIL NINAS': 'Na.',
  'ZAPATO DE NIÑA': 'Na.',
  'BEBE': 'Be.',
};

function abreviarMarca(marca) {
  if (!marca) return '';
  const palabras = marca.trim().split(/\s+/).filter(Boolean);
  // Artículos que se omiten al inicio (ej. "The Simpsons" → "Simpsons")
  const articulos = ['the', 'la', 'el', 'los', 'las'];
  const filtradas = palabras.filter((p) => !articulos.includes(p.toLowerCase()));
  if (filtradas.length === 0) return marca;
  if (filtradas.length === 1) return filtradas[0]; // una sola palabra, sin punto
  // Dos o más palabras: primera completa + inicial de la segunda (sin punto final;
  // el punto separador lo agrega generarDescripcion al unir las partes)
  return `${filtradas[0]}${filtradas[1][0].toUpperCase()}`;
}

function generarDescripcion(marca, tipo, subtipo, genero) {
  const partes = [];
  const marcaAbr = abreviarMarca(marca);
  if (marcaAbr) partes.push(marcaAbr);
  const tipoDesc = subtipo ? `${tipo}${subtipo}` : tipo;
  if (tipoDesc) partes.push(tipoDesc);
  if (genero) partes.push(genero);
  return partes.join('.');
}

// ---------- Selector con buscador ----------
function SearchableSelect({ id, value, onChange, options, placeholder, openId, setOpenId, recentKey, onCreate }) {
  const [query, setQuery] = useState('');
  const [recientes, setRecientes] = useState([]);
  const wrapRef = useRef(null);
  const open = openId === id;

  // Cargar recientes de este campo (marca, departamento, etc.)
  useEffect(() => {
    if (!recentKey) return;
    (async () => {
      const r = await loadShared(`pedidos:recientes:${recentKey}`, []);
      setRecientes(Array.isArray(r) ? r : []);
    })();
  }, [recentKey]);

  useEffect(() => {
    const handleClick = (e) => {
      if (open && wrapRef.current && !wrapRef.current.contains(e.target)) setOpenId(null);
    };
    document.addEventListener('mousedown', handleClick);
    return () => document.removeEventListener('mousedown', handleClick);
  }, [open, setOpenId]);

  const registrarReciente = (val) => {
    if (!recentKey || !val) return;
    const next = [val, ...recientes.filter((x) => x !== val)].slice(0, 8);
    setRecientes(next);
    saveShared(`pedidos:recientes:${recentKey}`, next);
  };

  const selectValue = (val) => {
    registrarReciente(val);
    onChange(val);
    setQuery('');
    setOpenId(null);
  };

  const crearValor = () => {
    const nuevo = query.trim();
    if (!nuevo || !onCreate) return;
    onCreate(nuevo);      // agrega a la lista maestra
    registrarReciente(nuevo);
    onChange(nuevo);      // lo selecciona
    setQuery('');
    setOpenId(null);
  };

  // Ordenar: recientes primero (en su orden), luego el resto alfabético
  const ordenar = (arr) => {
    if (!recentKey || recientes.length === 0) return arr;
    const enRecientes = recientes.filter((r) => arr.includes(r));
    const resto = arr.filter((o) => !recientes.includes(o));
    return [...enRecientes, ...resto];
  };

  const filtered = ordenar(options.filter((o) => o.toLowerCase().includes(query.toLowerCase())));
  const hayRecientes = recentKey && recientes.length > 0 && !query;
  // Mostrar opción de crear si hay texto escrito que no coincide exactamente con una opción existente
  const puedeCrear = onCreate && query.trim().length > 0 &&
    !options.some((o) => o.toLowerCase() === query.trim().toLowerCase());

  return (
    <div className={`relative ${open ? 'z-50' : 'z-0'}`} ref={wrapRef}>
      <input
        value={open ? query : value}
        onFocus={() => { setOpenId(id); setQuery(''); }}
        onChange={(e) => setQuery(e.target.value)}
        onKeyDown={(e) => { if (e.key === 'Enter' && puedeCrear) { e.preventDefault(); crearValor(); } }}
        placeholder={placeholder}
        className="w-full bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm relative"
      />
      {open && (
        <div
          className="absolute z-50 mt-1 w-full max-h-56 overflow-y-auto rounded-xl border border-app-line bg-app-panel shadow-app-lg"
          style={{ backgroundColor: '#1a1d24' }}
        >
          {value && (
            <button
              onMouseDown={(e) => { e.preventDefault(); onChange(''); setOpenId(null); }}
              className="w-full text-left px-3 py-2.5 text-xs text-app-red2 border-b border-app-line active:bg-app-active flex items-center justify-between"
            >
              Quitar selección
              <X size={13} />
            </button>
          )}
          {puedeCrear && (
            <button
              onMouseDown={(e) => { e.preventDefault(); crearValor(); }}
              className="w-full text-left px-3 py-2.5 text-sm font-medium text-app-gold active:bg-app-active flex items-center gap-2 border-b border-app-line"
            >
              <Plus size={14} className="shrink-0" />
              <span className="truncate">Crear "{query.trim()}"</span>
            </button>
          )}
          {filtered.length === 0 && !puedeCrear && (
            <p className="px-3 py-2.5 text-xs text-app-dim">Sin resultados.</p>
          )}
          {filtered.slice(0, 50).map((o, i) => {
            const esReciente = hayRecientes && recientes.includes(o);
            const esUltimoReciente = hayRecientes && esReciente && (i === recientes.filter((r) => options.includes(r)).length - 1);
            return (
              <button
                key={o}
                onMouseDown={(e) => { e.preventDefault(); selectValue(o); }}
                className={`w-full text-left px-3 py-2.5 text-sm active:bg-app-active flex items-center gap-2 ${esUltimoReciente ? 'border-b border-app-line' : ''} ${o === value ? 'text-app-gold font-semibold bg-app-goldbg' : 'text-app-white'}`}
              >
                {esReciente && <span className="text-xs text-app-dim shrink-0">🕘</span>}
                <span className="truncate">{o}</span>
              </button>
            );
          })}
          {filtered.length > 50 && (
            <p className="px-3 py-2.5 text-xs text-app-dim">Sigue escribiendo para acotar ({filtered.length} resultados)…</p>
          )}
        </div>
      )}
    </div>
  );
}

// ---------- Formulario de producto (reutilizable: catálogo y pedido) ----------
function ProductForm({ products = [], departamentos = [], tipos = [], marcas = [], ciudades = [], ciudadPorDefecto = '', onSetCiudadPorDefecto, fabricas = [], factores, tasaCambio, onSave, onCancel, onUpdateTipos, onCreateMarca, onCreateFabrica, onCreateCiudad, origen, pedidoMode = false, initialCodigo = '', usuarioActivoNombre = '', usuarioActivoPrefijo = '', numeroInicioCorrelativo = '', title = 'Nuevo producto' }) {
  const monedaOrigen = origen?.id === 'china' ? 'RMB' : (origen?.id === 'honduras' ? 'HNL' : 'USD');
  const [form, setForm] = useState({
    ...emptyForm(),
    ciudad: ciudadPorDefecto, // usar la ciudad elegida en el primer artículo del pedido
    costoMoneda: monedaOrigen,
    codigo: origen?.id === 'china'
      ? generarCodigoCorrelativo(products, usuarioActivoNombre, usuarioActivoPrefijo, numeroInicioCorrelativo) + (pedidoMode ? 'H' : '')
      : (initialCodigo || ''),
  });
  const [error, setError] = useState('');
  const [subiendoFotos, setSubiendoFotos] = useState(false);
  const [errorFoto, setErrorFoto] = useState('');
  const [codigoEditado, setCodigoEditado] = useState(false); // el usuario escribió el código a mano
  const [openFieldId, setOpenFieldId] = useState(null);
  const [descEditada, setDescEditada] = useState(false);
  const [sinTalla, setSinTalla] = useState(false);
  const [matrix, setMatrix] = useState({});       // cantidades del pedido (solo en pedidoMode)
  const [qtySinTalla, setQtySinTalla] = useState({}); // cantidades sin talla por color
  const [destinoPedido, setDestinoPedido] = useState('H'); // destino del pedido (solo pedidoMode)
  const [productoYaCreado, setProductoYaCreado] = useState(false); // en pedidoMode, tras guardar el producto no se recrea
  const [feedbackPedidoMode, setFeedbackPedidoMode] = useState(''); // mensaje temporal de confirmación

  // Limpiar feedback tras 2 segundos
  useEffect(() => {
    if (!feedbackPedidoMode) return;
    const t = setTimeout(() => setFeedbackPedidoMode(''), 2000);
    return () => clearTimeout(t);
  }, [feedbackPedidoMode]);

  // Si el usuario cambia el "Código inicial de esta compra" con el formulario ya abierto,
  // recalcular el código de referencia. No pisa el código si ya lo escribió a mano ni si
  // el producto ya fue creado (segundo pase para el otro país).
  useEffect(() => {
    if (origen?.id !== 'china') return;
    if (codigoEditado || productoYaCreado) return;
    const nuevo = generarCodigoCorrelativo(products, usuarioActivoNombre, usuarioActivoPrefijo, numeroInicioCorrelativo)
      + (pedidoMode ? destinoPedido : '');
    setForm((f) => (f.codigo === nuevo ? f : { ...f, codigo: nuevo }));
  }, [numeroInicioCorrelativo]);

  // Actualiza las tallas de un tipo directamente desde el formulario
  const setTiposInline = (tipoId, tallas) => {
    if (onUpdateTipos) onUpdateTipos(tipos.map((t) => t.id === tipoId ? { ...t, tallas } : t));
  };

  const tipoSeleccionado = tipos.find((t) => t.nombre === form.tipo);
  const esCinturaLargo = tipoSeleccionado?.medida === 'cintura_largo';
  const tallasDisponibles = getTallasDeTipo(tipoSeleccionado);

  const updateField = (changes) => {
    setForm((prev) => {
      const next = { ...prev, ...changes };
      if (!descEditada) {
        next.descripcion = generarDescripcion(next.marca, next.tipo, next.subtipo, next.genero);
      }
      return next;
    });
  };

  const handleDeptChange = (v) => {
    const generoSugerido = DEPTO_GENERO[v] || '';
    updateField({ departamento: v, genero: generoSugerido });
  };

  const handlePhoto = async (e) => {
    const archivos = Array.from(e.target.files || []);
    e.target.value = '';
    if (!archivos.length) return;
    setSubiendoFotos(true);
    setErrorFoto('');
    try {
      const subidas = [];
      for (const file of archivos) {
        try {
          subidas.push(await subirFoto(file, form.codigo));
        } catch (err) {
          setErrorFoto(`No se pudo subir "${file.name}": ${err.message || err}`);
        }
      }
      if (subidas.length) {
        setForm((f) => ({ ...f, fotos: [...(f.fotos || []), ...subidas] }));
      }
    } finally {
      setSubiendoFotos(false);
    }
  };

  const quitarFoto = async (indice) => {
    const foto = (form.fotos || [])[indice];
    setForm((f) => ({ ...f, fotos: (f.fotos || []).filter((_, i) => i !== indice) }));
    await borrarFoto(foto);
  };

  const addColor = () => {
    const c = form.colorInput.trim();
    if (!c) return;
    if ((form.colores || []).includes(c)) { setForm((f) => ({ ...f, colorInput: '' })); return; }
    setForm((f) => ({ ...f, colores: [...f.colores, c], colorInput: '' }));
  };
  const removeColor = (c) => setForm((f) => ({ ...f, colores: (f.colores || []).filter((x) => x !== c) }));

  const handleSave = () => {
    setError('');
    const codigoBase = sanitizarCodigo(form.codigo);
    if (!codigoBase) return setError('El código de referencia es obligatorio.');
    if (!form.departamento) return setError('Selecciona un departamento.');
    // Marca: obligatoria para USA/Panamá/Honduras, opcional para China
    if ((origen?.id === 'usa' || origen?.id === 'panama' || origen?.id === 'honduras') && !form.marca) return setError('Selecciona una marca.');
    if (!form.tipo && !sinTalla) return setError('Selecciona un tipo de producto.');
    if ((form.colores || []).length === 0) return setError('Agrega al menos un color.');
    if (!form.costoMonto) return setError('Captura el precio de costo.');
    if (!form.ventaLempiras) return setError('Captura el precio de venta en Lempiras.');

    // Verificar que cada código color no exista ya
    // (se salta si el producto ya fue creado en esta misma sesión — 2do pase para otro país)
    const esChina = origen?.id === 'china';
    
    if (!productoYaCreado) {
      if (esChina) {
        // China: un producto por color con código BASE-Color
        const conflictos = (form.colores || []).filter((color) => {
          const codigoColor = sanitizarCodigo(`${codigoBase}-${color}`);
          return products.some((p) => (p.codigo || '').trim().toLowerCase() === codigoColor.toLowerCase());
        });
        if (conflictos.length > 0) return setError(`Ya existen: ${conflictos.map((c) => `${codigoBase}-${c}`).join(', ')}`);
      } else {
        // USA/Panamá/Honduras: un solo producto con todos los colores, código único
        if (products.some((p) => (p.codigo || '').trim().toLowerCase() === codigoBase.toLowerCase())) {
          return setError(`Ya existe un producto con el código ${codigoBase}.`);
        }
      }
    }

    // En modo pedido validar que haya al menos alguna cantidad
    if (pedidoMode) {
      const totalCantidad = sinTalla
        ? Object.values(qtySinTalla).reduce((s, v) => s + (parseInt(v) || 0), 0)
        : (esCinturaLargo ? matrixCLToVariantes(matrix) : matrixToVariantes(matrix)).reduce((s, v) => s + v.cantidad, 0);
      if (totalCantidad === 0) return setError('Captura al menos una cantidad para el pedido.');
    }

    // Estructura de un producto (China: uno por color; USA/Panamá/Honduras: uno con todos)
    const construirProducto = (colorOrColores, variantes) => {
      const esUno = typeof colorOrColores === 'string';
      const codigo = esUno
        ? sanitizarCodigo(`${codigoBase}-${colorOrColores}`)
        : codigoBase; // USA/Panamá/Honduras: código tal cual
      return {
        id: uid(),
        codigo,
        codigoBase,
        color: esUno ? colorOrColores : undefined,
        origen: origen?.id || '',
        descripcion: (form.descripcion || '').trim(),
        departamento: form.departamento,
        tipo: form.tipo,
        subtipo: form.subtipo,
        medida: sinTalla ? 'sin_talla' : (tipoSeleccionado?.medida || 'simple'),
        tallas: sinTalla ? [] : tallasDisponibles,
        cinturas: esCinturaLargo ? (tipoSeleccionado?.cinturas || []) : [],
        largos: esCinturaLargo ? (tipoSeleccionado?.largos || []) : [],
        marca: form.marca,
        ciudad: (form.ciudad || '').trim(),
        fabrica: (form.fabrica || '').trim(),
        colores: esUno ? [colorOrColores] : colorOrColores,
        variantes,
        costoMonto: parseFloat(form.costoMonto),
        costoMoneda: form.costoMoneda,
        ventaLempiras: parseFloat(form.ventaLempiras),
        foto: form.fotos?.[0] ? urlFoto(form.fotos[0], false) : null,
        fotos: form.fotos || [],
        destinoPedido: pedidoMode ? destinoPedido : undefined,
      };
    };

    let nuevos;
    if (esChina) {
      // China: un producto por color
      nuevos = (form.colores || []).map((color) => {
        let variantes = [];
        if (pedidoMode) {
          if (sinTalla) {
            variantes = [{ talla: 'Sin talla', color, cantidad: parseInt(qtySinTalla[color] || 0) }].filter((v) => v.cantidad > 0);
          } else {
            const todas = esCinturaLargo ? matrixCLToVariantes(matrix) : matrixToVariantes(matrix);
            variantes = todas.filter((v) => v.color === color);
          }
        }
        return construirProducto(color, variantes);
      });
    } else {
      // USA/Panamá: UN solo producto con todos los colores dentro
      let variantes = [];
      if (pedidoMode) {
        if (sinTalla) {
          variantes = form.colores
            .map((c) => ({ talla: 'Sin talla', color: c, cantidad: parseInt(qtySinTalla[c] || 0) }))
            .filter((v) => v.cantidad > 0);
        } else {
          variantes = esCinturaLargo ? matrixCLToVariantes(matrix) : matrixToVariantes(matrix);
        }
      }
      nuevos = [construirProducto(form.colores, variantes)];
    }

    onSave(nuevos);

    // En modo pedido, NO cerrar la ventana: mantener info del artículo, limpiar solo
    // las cantidades para poder capturar el otro país en seguida.
    if (pedidoMode) {
      setProductoYaCreado(true);
      setMatrix({});
      setQtySinTalla({});
      // Cambiar automáticamente al otro destino para agilizar la captura
      const otroDestino = destinoPedido === 'H' ? 'G' : 'H';
      setDestinoPedido(otroDestino);
      const totalPzs = nuevos.reduce((s, p) => s + (p.variantes || []).reduce((a, v) => a + (v.cantidad || 0), 0), 0);
      const emojiPrev = destinoPedido === 'H' ? '🇭🇳' : '🇬🇹';
      setFeedbackPedidoMode(`✓ ${emojiPrev} Agregado (${totalPzs} pzs). Ahora captura ${otroDestino === 'H' ? '🇭🇳' : '🇬🇹'}`);
      if (navigator.vibrate) navigator.vibrate(40);
    }
  };

  return (
    <div className="bg-app-panel border border-app-line rounded-xl shadow-app p-4 space-y-3">
      <div className="flex justify-between items-center">
        <span className="text-sm font-medium">{title}</span>
        <button onClick={onCancel}><X size={16} className="text-app-dim" /></button>
      </div>

      {error && (
        <div className="flex items-center gap-2 bg-app-redbg text-app-red2 text-xs rounded-lg px-3 py-2">
          <AlertCircle size={14} className="shrink-0" /> {error}
        </div>
      )}

      <div className="flex gap-2">
        <div className="flex-1">
          <label className="text-xs text-app-dim2 uppercase tracking-wide mb-1 flex items-center justify-between">
            <span>Ciudad de compra</span>
            <span className="normal-case tracking-normal text-app-dim">(opcional)</span>
          </label>
          <SearchableSelect
            id="ciudad"
            openId={openFieldId}
            setOpenId={setOpenFieldId}
            value={form.ciudad}
            onChange={(v) => {
              setForm({ ...form, ciudad: v });
              // Primera vez que se elige ciudad en este pedido: guardarla como default
              if (!ciudadPorDefecto && v && onSetCiudadPorDefecto) {
                onSetCiudadPorDefecto(v);
              }
            }}
            options={ciudadesDeOrigen(ciudades, origen?.id)}
            placeholder="Buscar ciudad…"
            recentKey="ciudad"
            onCreate={onCreateCiudad}
          />
        </div>
        {origen?.id === 'china' && (
        <div className="flex-1">
          <label className="text-xs text-app-dim2 uppercase tracking-wide mb-1 flex items-center justify-between">
            <span>Fábrica</span>
            <span className="normal-case tracking-normal text-app-dim">(opcional)</span>
          </label>
          <div className="relative">
            <SearchableSelect
              id="fabrica"
              openId={openFieldId}
              setOpenId={setOpenFieldId}
              value={form.fabrica}
              onChange={(v) => setForm({ ...form, fabrica: v })}
              options={fabricas || []}
              placeholder="Buscar o dejar en blanco…"
              recentKey="fabrica"
              onCreate={onCreateFabrica}
            />
            {form.fabrica && (
              <button
                type="button"
                onClick={() => setForm({ ...form, fabrica: '' })}
                className="absolute right-2 top-1/2 -translate-y-1/2 text-app-dim active:text-app-red p-1"
                title="Dejar sin fábrica"
              >
                <X size={14} />
              </button>
            )}
          </div>
        </div>
        )}
      </div>

      <div className="space-y-2">
        {(form.fotos || []).length > 0 && (
          <div className="grid grid-cols-3 gap-2">
            {(form.fotos || []).map((foto, i) => (
              <div key={foto.path || i} className="relative aspect-square">
                <img
                  src={urlFoto(foto)}
                  alt={`Foto ${i + 1}`}
                  className="w-full h-full object-cover rounded-lg border border-app-line"
                />
                <button
                  onClick={() => quitarFoto(i)}
                  className="absolute top-1 right-1 bg-app-bg-80 rounded-full p-1"
                  aria-label="Quitar foto"
                >
                  <X size={12} />
                </button>
              </div>
            ))}
          </div>
        )}

        <label
          className={`w-full rounded-lg border border-dashed border-app-line3 flex flex-col items-center
                      justify-center gap-1 py-4 px-3 text-center cursor-pointer
                      ${subiendoFotos ? 'opacity-60 pointer-events-none' : 'hover:border-app-gold'}`}
        >
          <Camera size={20} className="text-app-dim3" />
          <span className="text-xs text-app-dim2">
            {subiendoFotos
              ? 'Subiendo…'
              : (form.fotos || []).length
                ? 'Agregar más fotos'
                : 'Tomar foto o elegir del dispositivo'}
          </span>
          <span className="text-xs text-app-dim3">Se guardan en resolución original</span>
          <input
            type="file"
            accept="image/*"
            multiple
            className="hidden"
            onChange={handlePhoto}
            disabled={subiendoFotos}
          />
        </label>

        {errorFoto && <p className="text-xs text-app-red">{errorFoto}</p>}
      </div>

      <div>
        <label className="text-xs text-app-dim2 uppercase tracking-wide mb-1 flex items-center justify-between">
          <span>Código de referencia <span className="text-app-red">*</span></span>
        </label>
        <div className="flex gap-2">
          <input
            placeholder="Código único del producto"
            value={form.codigo}
            onChange={(e) => { setCodigoEditado(true); setForm({ ...form, codigo: e.target.value.replace(/\s+/g, '-') }); }}
            className="flex-1 bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm font-mono"
          />
          {origen?.id === 'china' && (
            <button
              type="button"
              onClick={() => setForm((f) => ({ ...f, codigo: generarCodigoCorrelativo(products, usuarioActivoNombre, usuarioActivoPrefijo) + (pedidoMode ? destinoPedido : '') }))}
              className="px-3 rounded-lg border border-app-line text-app-gold text-sm active:bg-app-active"
              title="Generar siguiente correlativo"
            >
              ↺
            </button>
          )}
        </div>
        {origen?.id === 'china' && (
          <p className="text-xs text-app-dim mt-1">Correlativo automático: Y + año + mes + número</p>
        )}
      </div>

      <div>
        <label className="text-xs text-app-dim2 uppercase tracking-wide mb-1 flex items-center justify-between">
          <span>Marca {origen?.id === 'china'
            ? <span className="normal-case tracking-normal text-app-dim">(opcional)</span>
            : <span className="text-app-red">*</span>}
          </span>
        </label>
        <SearchableSelect
          id="marca"
          openId={openFieldId}
          setOpenId={setOpenFieldId}
          value={form.marca}
          onChange={(v) => updateField({ marca: v })}
          options={marcas}
          placeholder={origen?.id === 'china' ? 'Buscar o dejar en blanco…' : 'Buscar marca…'}
          recentKey="marca"
          onCreate={onCreateMarca}
        />
      </div>
      <div>
        <label className="text-xs text-app-dim2 uppercase tracking-wide mb-1 flex items-center justify-between">
          <span>Departamento <span className="text-app-red">*</span></span>
        </label>
        <SearchableSelect
          id="departamento"
          openId={openFieldId}
          setOpenId={setOpenFieldId}
          value={form.departamento}
          onChange={handleDeptChange}
          options={departamentos}
          placeholder="Buscar departamento…"
          recentKey="departamento"
        />
      </div>

      {marcas.length === 0 && (
        <p className="text-xs text-app-gold">No tienes marcas creadas. Toca el ícono de ajustes en el catálogo para agregar una.</p>
      )}
      {departamentos.length === 0 && (
        <p className="text-xs text-app-gold">No tienes departamentos creados. Toca el ícono de ajustes en el catálogo para agregar uno.</p>
      )}

      <div>
        <label className="text-xs text-app-dim2 uppercase tracking-wide mb-1 flex items-center justify-between">
          <span>Tipo de producto <span className="text-app-red">*</span></span>
        </label>
        <SearchableSelect
          id="tipo"
          openId={openFieldId}
          setOpenId={setOpenFieldId}
          value={form.tipo}
          onChange={(v) => updateField({ tipo: v, subtipo: '' })}
          options={tipos.map((t) => t.nombre)}
          placeholder="Buscar tipo… (define las tallas)"
          recentKey="tipo"
        />
      </div>
      {tipos.length === 0 && (
        <p className="text-xs text-app-gold">No tienes tipos creados. Toca el ícono de ajustes en el catálogo para agregar uno.</p>
      )}

      {tipoSeleccionado?.subtipos && tipoSeleccionado.subtipos.length > 0 && (
        <div>
          <label className="text-xs text-app-dim2 uppercase tracking-wide mb-1 flex items-center justify-between">
            <span>Subtipo</span>
            <span className="normal-case tracking-normal text-app-dim">(opcional)</span>
          </label>
          <SearchableSelect
            id="subtipo"
            openId={openFieldId}
            setOpenId={setOpenFieldId}
            value={form.subtipo}
            onChange={(v) => updateField({ subtipo: v })}
            options={tipoSeleccionado.subtipos}
            placeholder={`Buscar subtipo de ${tipoSeleccionado.nombre}…`}
          />
        </div>
      )}

      <div>
        <div className="flex items-center justify-between mb-1.5">
          <label className="text-xs uppercase tracking-wide text-app-dim2">Descripción del sistema</label>
          {descEditada && (
            <button
              type="button"
              onClick={() => {
                const desc = generarDescripcion(form.marca, form.tipo, form.subtipo, form.genero);
                setForm((f) => ({ ...f, descripcion: desc }));
                setDescEditada(false);
              }}
              className="text-xs text-app-gold underline decoration-dotted"
            >
              ↺ Regenerar
            </button>
          )}
        </div>
        <input
          value={form.descripcion}
          onChange={(e) => { setForm((f) => ({ ...f, descripcion: e.target.value })); setDescEditada(true); }}
          placeholder="Se genera al completar Marca + Tipo + Departamento"
          className="w-full bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm font-mono"
        />
        {form.descripcion && (
          <p className="text-xs text-app-dim mt-1">
            Preview: <span className="text-app-gold font-mono">{form.descripcion}</span>
            {descEditada && <span className="text-app-dim2 ml-1">(editada manualmente)</span>}
          </p>
        )}
      </div>

      <div>
        <label className="text-xs text-app-dim2 uppercase tracking-wide mb-1 flex items-center justify-between">
          <span>Colores <span className="text-app-red">*</span></span>
        </label>
        <div className="flex gap-2">
          <input
            placeholder="Agregar color y Enter"
            value={form.colorInput}
            onChange={(e) => setForm({ ...form, colorInput: e.target.value })}
            onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); addColor(); } }}
            className="flex-1 bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm"
          />
          <button onClick={addColor} className="px-3 rounded-lg border border-app-line text-sm text-app-dim2">Añadir</button>
        </div>
        {/* Atajo: agregar "Surtido" con un toque */}
        {!(form.colores || []).includes('Surtido') && (
          <button
            type="button"
            onClick={() => setForm((f) => ({ ...f, colores: [...f.colores, 'Surtido'], colorInput: '' }))}
            className="mt-1.5 text-xs text-app-gold border border-app-line rounded-lg px-2.5 py-1.5 active:bg-app-active"
          >
            + Surtido
          </button>
        )}
        {(form.colores || []).length > 0 && (
          <div className="mt-2 space-y-1.5">
            {(form.colores || []).map((c) => (
              <div key={c} className="flex items-center justify-between bg-app-bg border border-app-line rounded-lg px-3 py-2">
                <div className="min-w-0">
                  <p className="text-sm">{c}</p>
                  {form.codigo && (
                    <p className="text-xs font-mono text-app-gold">
                      {origen?.id === 'china' ? sanitizarCodigo(`${form.codigo}-${c}`) : form.codigo}
                    </p>
                  )}
                </div>
                <button onClick={() => removeColor(c)} className="text-app-dim active:text-app-red shrink-0 ml-2">
                  <X size={14} />
                </button>
              </div>
            ))}
          </div>
        )}
      </div>

      {/* Tallas disponibles — solo informativo, cantidades en el pedido */}
      {form.tipo && tipoSeleccionado && (
        <div className="bg-app-bg border border-app-line rounded-xl px-3 py-2.5">
          <div className="flex items-center justify-between mb-1.5">
            <label className="text-xs uppercase tracking-wide text-app-dim2">
              {pedidoMode ? 'Cantidad del pedido por talla' : 'Tallas disponibles'}
            </label>
            <button
              type="button"
              onClick={() => { setSinTalla(!sinTalla); setMatrix({}); setQtySinTalla({}); }}
              className={`text-xs px-2.5 py-1 rounded-lg border ${sinTalla ? 'bg-app-gold text-app-bg border-app-gold font-semibold' : 'bg-app-bg border-app-line text-app-dim2'}`}
            >
              Sin talla
            </button>
          </div>
          {sinTalla ? (
            pedidoMode && (form.colores || []).length > 0 ? (
              <div className="space-y-2">
                {(form.colores || []).map((color) => (
                  <div key={color} className="flex items-center gap-3">
                    <p className="text-sm flex-1">{color}</p>
                    <input type="number" min={0} placeholder="0"
                      value={qtySinTalla[color] ?? ''}
                      onChange={(e) => setQtySinTalla((prev) => ({ ...prev, [color]: e.target.value }))}
                      className="w-20 bg-app-panel border border-app-line rounded-lg px-2 py-1.5 text-sm text-center"
                    />
                    <span className="text-xs text-app-dim2">pzs</span>
                  </div>
                ))}
              </div>
            ) : (
              <p className="text-xs text-app-dim2 italic">Este producto no maneja tallas.</p>
            )
          ) : tallasDisponibles.length === 0 ? (
            <div className="space-y-1.5">
              <p className="text-xs text-app-gold">El tipo no tiene tallas. Elige un grupo:</p>
              <div className="flex flex-wrap gap-1.5">
                {GRUPOS_TALLAS.filter((g) => !['Cintura pantalón', 'Largo pantalón'].includes(g.label)).map((g) => (
                  <button key={g.label} type="button"
                    onClick={() => setTiposInline(tipoSeleccionado.id, g.tallas)}
                    className="text-xs bg-app-panel border border-app-line rounded-lg px-2.5 py-1.5 text-app-dim2">
                    {g.label}
                  </button>
                ))}
              </div>
            </div>
          ) : pedidoMode ? (
            // En modo pedido: mostrar tabla editable de cantidades
            esCinturaLargo ? (
              <CinturaLargoMatrix
                cinturas={tipoSeleccionado?.cinturas || []}
                largos={tipoSeleccionado?.largos || []}
                colores={(form.colores || []).length > 0 ? form.colores : ['_']}
                values={matrix}
                onChange={setMatrix}
              />
            ) : (
              <VariantMatrix
                tallas={tallasDisponibles}
                colores={(form.colores || []).length > 0 ? form.colores : ['_']}
                values={matrix}
                onChange={setMatrix}
              />
            )
          ) : (
            // En modo catálogo: solo mostrar chips de tallas
            <div className="flex flex-wrap gap-1.5">
              {tallasDisponibles.map((t) => (
                <span key={t} className="text-xs bg-app-panel border border-app-line rounded-full px-2.5 py-1 text-app-light">{t}</span>
              ))}
            </div>
          )}
        </div>
      )}

      <div>
        <label className="text-xs text-app-dim2 uppercase tracking-wide mb-1 flex items-center justify-between">
          <span>Precio de costo <span className="text-app-red">*</span></span>
        </label>
        <div className="flex gap-2">
          <input
            placeholder="Precio de costo"
            type="number"
            value={form.costoMonto}
            onChange={(e) => setForm({ ...form, costoMonto: e.target.value })}
            className="flex-1 bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm"
          />
          <select
            value={form.costoMoneda}
            onChange={(e) => setForm({ ...form, costoMoneda: e.target.value })}
            className="w-24 bg-app-bg border border-app-line rounded-lg px-2 py-2 text-sm"
          >
            {MONEDAS_COSTO.map((m) => <option key={m}>{m}</option>)}
          </select>
        </div>
      </div>

      {/* Niki en USD (solo China) — mostrar arriba del costo bodega */}
      {(() => {
        const niki = calcularNikiUSD(form.costoMonto, form.costoMoneda, origen?.id, tasaCambio?.rmbUsd, factores?.nikiPct);
        if (niki === null) return null;
        return (
          <div className="bg-app-panel border border-app-line rounded-xl px-3 py-2">
            <p className="text-xs uppercase tracking-wide text-app-dim2">Niki</p>
            <p className="text-base font-semibold text-app-gold mt-0.5">$ {niki.toFixed(2)} USD</p>
          </div>
        );
      })()}

      {/* Costo puesto en bodega (Lempiras) */}
      {(() => {
        const bodega = costoBodegaHNL(form.costoMonto, form.costoMoneda, origen?.id, factores, tasaCambio?.rmbUsd, factores?.nikiPct);
        if (!bodega) return null;
        const factorPais = factores?.[origen?.id];
        return (
          <div className="bg-app-blue border border-app-line2 rounded-xl px-3 py-2.5">
            <div className="flex items-center justify-between">
              <div>
                <p className="text-xs uppercase tracking-wide text-app-dim2">Costo puesto en bodega</p>
                <p className="text-lg font-bold text-app-sky mt-0.5">{fmtLempiras(bodega)}</p>
              </div>
              <div className="text-right">
                <p className="text-xs text-app-dim2">Factor {origen?.label}</p>
                <p className="text-sm font-semibold text-app-white">× {factorPais}</p>
              </div>
            </div>
            <p className="text-xs text-app-dim2 mt-1">
              Incluye flete, seguro, impuestos y tasa USD→HNL.
            </p>
          </div>
        );
      })()}

      {(() => {
        const bodega = costoBodegaHNL(form.costoMonto, form.costoMoneda, origen?.id, factores, tasaCambio?.rmbUsd, factores?.nikiPct);
        const sugeridos = sugerirPreciosVentaHNL(bodega, origen?.id);
        if (sugeridos.length === 0) return null;
        return (
          <div>
            <p className="text-xs uppercase tracking-wide text-app-dim2 mb-1.5">
              Venta sugerida (margen + IVA)
            </p>
            <div className="flex flex-wrap gap-1.5">
              {sugeridos.map(({ margen, etiqueta, precio }) => (
                <button
                  key={margen}
                  onClick={() => setForm((f) => ({ ...f, ventaLempiras: String(precio) }))}
                  className={`text-xs rounded-lg border px-2.5 py-1.5 transition ${
                    String(precio) === form.ventaLempiras
                      ? 'bg-app-gold text-app-bg border-app-gold font-semibold'
                      : 'bg-app-bg border-app-line text-app-light'
                  }`}
                >
                  <span className="font-mono">
                    {etiqueta ? `${margen}% = ${etiqueta}%` : `${margen}%`}
                  </span>
                  <span className="mx-1 text-app-dim2">→</span>
                  {fmtLempiras(precio)}
                </button>
              ))}
            </div>
          </div>
        );
      })()}

      <div>
        <label className="text-xs text-app-dim2 uppercase tracking-wide mb-1 flex items-center justify-between">
          <span>Precio de venta (HNL) <span className="text-app-red">*</span></span>
        </label>
        <input
          placeholder="Precio de venta en Lempiras"
          type="number"
          value={form.ventaLempiras}
          onChange={(e) => setForm({ ...form, ventaLempiras: e.target.value })}
          className="w-full bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm"
        />
      </div>

      {pedidoMode && (
        <div>
          <label className="text-xs uppercase tracking-wide text-app-dim2 mb-1.5 block">Destino del pedido</label>
          <div className="flex gap-2">
            {DESTINOS.map((d) => (
              <button
                key={d.id}
                type="button"
                onClick={() => {
                  setDestinoPedido(d.id);
                  // La terminación H/G del código es EXCLUSIVA de China (donde el código
                  // se genera aquí). Para USA/Panamá el código del proveedor no se toca.
                  if (origen?.id === 'china') {
                    setForm((f) => ({
                      ...f,
                      codigo: f.codigo ? f.codigo.replace(/[HG]$/, '') + d.id : f.codigo,
                    }));
                  }
                }}
                className={`flex-1 py-2 rounded-lg text-sm font-medium border transition flex items-center justify-center gap-1.5 ${
                  destinoPedido === d.id
                    ? 'bg-app-gold text-app-bg border-app-gold'
                    : 'bg-app-bg border-app-line text-app-dim2'
                }`}
              >
                <span>{d.emoji}</span> {d.label} ({d.id})
              </button>
            ))}
          </div>
        </div>
      )}

      {feedbackPedidoMode && (
        <p className="text-xs text-app-green text-center font-medium">{feedbackPedidoMode}</p>
      )}

      <div className="flex gap-2">
        <button onClick={handleSave} className="flex-1 py-2.5 rounded-lg bg-app-gold text-app-bg text-sm font-semibold">
          {pedidoMode
            ? productoYaCreado
              ? `Agregar ${destinoInfo(destinoPedido).emoji} al pedido`
              : (form.colores || []).length > 1
                ? `Guardar ${(form.colores || []).length} artículos y agregar al pedido`
                : 'Guardar artículo y agregar al pedido'
            : (form.colores || []).length > 1
              ? `Guardar ${(form.colores || []).length} productos (uno por color)`
              : 'Guardar producto'}
        </button>
        {pedidoMode && productoYaCreado && (
          <button
            onClick={onCancel}
            className="px-4 py-2.5 rounded-lg border border-app-line text-app-dim2 text-sm"
          >
            Finalizar
          </button>
        )}
      </div>
    </div>
  );
}

// ---------- Formulario de edición de producto ----------
function EditProductForm({ product, products = [], departamentos = [], tipos = [], marcas = [], ciudades = [], fabricas = [], onSave, onCancel }) {
  const tipoObj = tipos.find((t) => t.nombre === product.tipo);
  const esCLInit = tipoObj?.medida === 'cintura_largo';

  const [form, setForm] = useState({
    codigo: product.codigo,
    descripcion: product.descripcion || '',
    departamento: product.departamento || '',
    tipo: product.tipo || '',
    subtipo: product.subtipo || '',
    marca: product.marca || '',
    genero: product.genero || '',
    color: (product.colores || [])[0] || '',
    ciudad: product.ciudad || '',
    fabrica: product.fabrica || '',
    costoMonto: String(product.costoMonto || ''),
    costoMoneda: product.costoMoneda || 'USD',
    ventaLempiras: String(product.ventaLempiras || ''),
    foto: product.foto || null,
    fotos: listaFotos(product),
  });
  const [matrix, setMatrix] = useState(
    esCLInit ? variantesToMatrixCL(product.variantes) : variantesToMatrix(product.variantes)
  );
  const [error, setError] = useState('');
  const [subiendoFotos, setSubiendoFotos] = useState(false);
  const [errorFoto, setErrorFoto] = useState('');
  const [openFieldId, setOpenFieldId] = useState(null);
  const [descEditada, setDescEditada] = useState(true); // en edición siempre manual por defecto

  const tipoSeleccionado = tipos.find((t) => t.nombre === form.tipo);
  const esCinturaLargo = tipoSeleccionado?.medida === 'cintura_largo';
  const tallasDisponibles = getTallasDeTipo(tipoSeleccionado);

  const updateField = (changes) => {
    setForm((prev) => {
      const next = { ...prev, ...changes };
      if (!descEditada) {
        next.descripcion = generarDescripcion(next.marca, next.tipo, next.subtipo, next.genero);
      }
      return next;
    });
  };

  const handlePhoto = async (e) => {
    const archivos = Array.from(e.target.files || []);
    e.target.value = '';
    if (!archivos.length) return;
    setSubiendoFotos(true);
    setErrorFoto('');
    try {
      const subidas = [];
      for (const file of archivos) {
        try {
          subidas.push(await subirFoto(file, form.codigo));
        } catch (err) {
          setErrorFoto(`No se pudo subir "${file.name}": ${err.message || err}`);
        }
      }
      if (subidas.length) {
        setForm((f) => ({ ...f, fotos: [...(f.fotos || []), ...subidas] }));
      }
    } finally {
      setSubiendoFotos(false);
    }
  };

  const quitarFoto = async (indice) => {
    const foto = (form.fotos || [])[indice];
    setForm((f) => ({ ...f, fotos: (f.fotos || []).filter((_, i) => i !== indice) }));
    await borrarFoto(foto);
  };

  const handleSave = () => {
    setError('');
    if (!(form.codigo || '').trim()) return setError('El código es obligatorio.');
    // Verificar unicidad excluyendo el propio producto
    if (products.some((p) => p.id !== product.id && (p.codigo || '').trim().toLowerCase() === (form.codigo || '').trim().toLowerCase())) {
      return setError('Ese código ya lo usa otro producto.');
    }
    if (!form.departamento) return setError('Selecciona un departamento.');
    if (!form.tipo) return setError('Selecciona un tipo de producto.');
    if (!form.costoMonto) return setError('Captura el precio de costo.');
    if (!form.ventaLempiras) return setError('Captura el precio de venta en Lempiras.');

    const variantes = esCinturaLargo
      ? matrixCLToVariantes(matrix).filter((v) => v.color === form.color)
      : matrixToVariantes(matrix).filter((v) => v.color === form.color);

    onSave({
      ...product,
      codigo: sanitizarCodigo(form.codigo),
      descripcion: (form.descripcion || '').trim(),
      departamento: form.departamento,
      tipo: form.tipo,
      subtipo: form.subtipo,
      medida: tipoSeleccionado?.medida || 'simple',
      marca: form.marca,
      genero: form.genero,
      ciudad: (form.ciudad || '').trim(),
      fabrica: (form.fabrica || '').trim(),
      colores: [form.color],
      color: form.color,
      variantes,
      costoMonto: parseFloat(form.costoMonto),
      costoMoneda: form.costoMoneda,
      ventaLempiras: parseFloat(form.ventaLempiras),
      foto: form.fotos?.[0] ? urlFoto(form.fotos[0], false) : null,
      fotos: form.fotos || [],
    });
  };

  // Reconstruir matrix cuando el color del form cambia
  const matrixConColor = {};
  Object.entries(matrix).forEach(([k, v]) => {
    if (esCinturaLargo) {
      const [, cintura, largo] = k.split('__');
      matrixConColor[clKey(form.color, cintura, largo)] = v;
    } else {
      const [talla] = k.split('__');
      matrixConColor[variantKey(talla, form.color)] = v;
    }
  });

  return (
    <div className="mt-3 pt-3 border-t border-app-line space-y-2.5">
      {error && (
        <div className="flex items-center gap-2 bg-app-redbg text-app-red2 text-xs rounded-lg px-3 py-2">
          <AlertCircle size={14} className="shrink-0" /> {error}
        </div>
      )}

      <div className="flex gap-2">
        <div className="flex-1">
          <label className="text-xs text-app-dim2 uppercase tracking-wide mb-1 flex items-center justify-between">
            <span>Ciudad de compra</span>
            <span className="normal-case tracking-normal text-app-dim">(opcional)</span>
          </label>
          <SearchableSelect id="e-ciudad2" openId={openFieldId} setOpenId={setOpenFieldId}
            value={form.ciudad} onChange={(v) => setForm({ ...form, ciudad: v })}
            options={ciudadesDeOrigen(ciudades, product.origen)} placeholder="Buscar ciudad…" recentKey="ciudad" />
        </div>
        {product.origen === 'china' && (
        <div className="flex-1">
          <label className="text-xs text-app-dim2 uppercase tracking-wide mb-1 flex items-center justify-between">
            <span>Fábrica</span>
            <span className="normal-case tracking-normal text-app-dim">(opcional)</span>
          </label>
          <div className="relative">
            <SearchableSelect id="e-fabrica2" openId={openFieldId} setOpenId={setOpenFieldId}
              value={form.fabrica} onChange={(v) => setForm({ ...form, fabrica: v })}
              options={fabricas || []} placeholder="Buscar o dejar en blanco…" recentKey="fabrica" />
            {form.fabrica && (
              <button
                type="button"
                onClick={() => setForm({ ...form, fabrica: '' })}
                className="absolute right-2 top-1/2 -translate-y-1/2 text-app-dim active:text-app-red p-1"
                title="Dejar sin fábrica"
              >
                <X size={14} />
              </button>
            )}
          </div>
        </div>
        )}
      </div>

      <div className="space-y-2">
        {(form.fotos || []).length > 0 && (
          <div className="grid grid-cols-3 gap-2">
            {(form.fotos || []).map((foto, i) => (
              <div key={foto.path || i} className="relative aspect-square">
                <img
                  src={urlFoto(foto)}
                  alt={`Foto ${i + 1}`}
                  className="w-full h-full object-cover rounded-lg border border-app-line"
                />
                <button
                  onClick={() => quitarFoto(i)}
                  className="absolute top-1 right-1 bg-app-bg-80 rounded-full p-1"
                  aria-label="Quitar foto"
                >
                  <X size={12} />
                </button>
              </div>
            ))}
          </div>
        )}

        <label
          className={`w-full rounded-lg border border-dashed border-app-line3 flex flex-col items-center
                      justify-center gap-1 py-4 px-3 text-center cursor-pointer
                      ${subiendoFotos ? 'opacity-60 pointer-events-none' : 'hover:border-app-gold'}`}
        >
          <Camera size={18} className="text-app-dim3" />
          <span className="text-xs text-app-dim2">
            {subiendoFotos
              ? 'Subiendo…'
              : (form.fotos || []).length ? 'Agregar más fotos' : 'Tomar foto o elegir del dispositivo'}
          </span>
          <input
            type="file"
            accept="image/*"
            multiple
            className="hidden"
            onChange={handlePhoto}
            disabled={subiendoFotos}
          />
        </label>

        {errorFoto && <p className="text-xs text-app-red">{errorFoto}</p>}
      </div>

      <input
        placeholder="Código"
        value={form.codigo}
        onChange={(e) => setForm({ ...form, codigo: e.target.value.replace(/\s+/g, '-') })}
        className="w-full bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm font-mono"
      />

      <div>
        <div className="flex items-center justify-between mb-1">
          <label className="text-xs text-app-dim2 uppercase tracking-wide">Descripción del sistema</label>
          <button
            type="button"
            onClick={() => {
              const genero = DEPTO_GENERO[form.departamento] || '';
              const desc = generarDescripcion(form.marca, form.tipo, form.subtipo, genero);
              setForm((f) => ({ ...f, descripcion: desc }));
            }}
            className="text-xs text-app-gold underline decoration-dotted"
          >
            ↺ Regenerar
          </button>
        </div>
        <input
          value={form.descripcion}
          onChange={(e) => setForm((f) => ({ ...f, descripcion: e.target.value }))}
          className="w-full bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm font-mono"
        />
      </div>

      <SearchableSelect id="e-marca" openId={openFieldId} setOpenId={setOpenFieldId}
        value={form.marca} onChange={(v) => updateField({ marca: v })}
        options={marcas} placeholder="Buscar marca…" recentKey="marca" />

      <SearchableSelect id="e-depto" openId={openFieldId} setOpenId={setOpenFieldId}
        value={form.departamento} onChange={(v) => {
          const genero = DEPTO_GENERO[v] || '';
          updateField({ departamento: v, genero });
        }}
        options={departamentos} placeholder="Buscar departamento…" recentKey="departamento" />

      <SearchableSelect id="e-tipo" openId={openFieldId} setOpenId={setOpenFieldId}
        value={form.tipo} onChange={(v) => { updateField({ tipo: v, subtipo: '' }); setMatrix({}); }}
        options={tipos.map((t) => t.nombre)} placeholder="Buscar tipo de producto…" recentKey="tipo" />

      {tipoSeleccionado?.subtipos?.length > 0 && (
        <SearchableSelect id="e-subtipo" openId={openFieldId} setOpenId={setOpenFieldId}
          value={form.subtipo} onChange={(v) => updateField({ subtipo: v })}
          options={tipoSeleccionado.subtipos} placeholder={`Subtipo de ${tipoSeleccionado.nombre}…`} />
      )}

      <div>
        <label className="text-xs text-app-dim2 uppercase tracking-wide mb-1.5 block">Color del producto</label>
        <input
          value={form.color}
          onChange={(e) => setForm({ ...form, color: e.target.value })}
          placeholder="Color"
          className="w-full bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm"
        />
        {form.color && form.codigo && (
          <p className="text-xs font-mono text-app-gold mt-1">{form.codigo}</p>
        )}
      </div>

      <div>
        <label className="text-xs text-app-dim2 uppercase tracking-wide mb-1.5 block">
          {esCinturaLargo ? 'Cantidad por cintura, largo y color' : 'Cantidad por talla'}
        </label>
        {esCinturaLargo ? (
          <CinturaLargoMatrix
            cinturas={tipoSeleccionado?.cinturas || []}
            largos={tipoSeleccionado?.largos || []}
            colores={[form.color]}
            values={matrixConColor}
            onChange={(newMatrix) => {
              const simplified = {};
              Object.entries(newMatrix).forEach(([k, v]) => {
                const [, cintura, largo] = k.split('__');
                simplified[`__${cintura}__${largo}`] = v;
              });
              setMatrix(newMatrix);
            }}
          />
        ) : (
          <VariantMatrix
            tallas={tallasDisponibles}
            colores={[form.color]}
            values={matrixConColor}
            onChange={(newMatrix) => setMatrix(newMatrix)}
          />
        )}
      </div>

      <div className="flex gap-2">
        <input placeholder="Precio costo" type="number" value={form.costoMonto}
          onChange={(e) => setForm({ ...form, costoMonto: e.target.value })}
          className="flex-1 bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm" />
        <select value={form.costoMoneda} onChange={(e) => setForm({ ...form, costoMoneda: e.target.value })}
          className="w-24 bg-app-bg border border-app-line rounded-lg px-2 py-2 text-sm">
          {MONEDAS_COSTO.map((m) => <option key={m}>{m}</option>)}
        </select>
      </div>

      {sugerirPreciosVenta(form.costoMonto, form.costoMoneda).length > 0 && (
        <div className="flex flex-wrap gap-1.5">
          {sugerirPreciosVenta(form.costoMonto, form.costoMoneda).map(({ multiplicador, precio }) => (
            <button key={multiplicador}
              onClick={() => setForm((f) => ({ ...f, ventaLempiras: String(precio) }))}
              className={`text-xs rounded-lg border px-2 py-1.5 ${String(precio) === form.ventaLempiras ? 'bg-app-gold text-app-bg border-app-gold font-semibold' : 'bg-app-bg border-app-line text-app-light'}`}>
              ×{multiplicador} → {fmtLempiras(precio)}
            </button>
          ))}
        </div>
      )}

      <input placeholder="Precio de venta en Lempiras (HNL)" type="number" value={form.ventaLempiras}
        onChange={(e) => setForm({ ...form, ventaLempiras: e.target.value })}
        className="w-full bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm" />

      <div className="flex gap-2 pt-1">
        <button onClick={onCancel} className="flex-1 py-2.5 rounded-lg border border-app-line text-sm text-app-dim2">
          Cancelar
        </button>
        <button onClick={handleSave} className="flex-1 py-2.5 rounded-lg bg-app-gold text-app-bg text-sm font-semibold">
          Guardar cambios
        </button>
      </div>
    </div>
  );
}

function Catalogo({ products = [], setProducts, departamentos = [], tipos = [], setTipos, marcas = [], ciudades = [], fabricas = [], factores, tasaCambio, origen, orders = [], suppliers = [], usuarioActivoNombre, usuarioActivoPrefijo, puedoBorrar = true, onOpenConfig, onCreateMarca, onCreateFabrica, onCreateCiudad }) {
  const [visor, setVisor] = useState(null);   // fotos a mostrar en pantalla completa
  const [showForm, setShowForm] = useState(false);
  const [query, setQuery] = useState('');
  const [expandedProductId, setExpandedProductId] = useState(null);
  const [editingProductId, setEditingProductId] = useState(null);
  // Filtros
  const [filtroProveedor, setFiltroProveedor] = useState('');
  const [filtroDepartamento, setFiltroDepartamento] = useState('');
  const [filtroMarca, setFiltroMarca] = useState('');
  const [openFiltroId, setOpenFiltroId] = useState(null);

  const hayFiltroActivo = !!(filtroProveedor || filtroDepartamento || filtroMarca || query.trim());

  // Productos que aparecen en pedidos del proveedor seleccionado
  const productIdsDelProveedor = (() => {
    if (!filtroProveedor) return null;
    const proveedorObj = (suppliers || []).find((s) => s.nombre === filtroProveedor);
    if (!proveedorObj) return new Set();
    const ids = new Set();
    (orders || []).forEach((o) => {
      if (o.supplierId === proveedorObj.id) {
        (o.items || []).forEach((it) => ids.add(it.productId));
      }
    });
    return ids;
  })();

  const filtered = !hayFiltroActivo ? [] : products.filter((p) => {
    if (filtroDepartamento && p.departamento !== filtroDepartamento) return false;
    if (filtroMarca && p.marca !== filtroMarca) return false;
    if (productIdsDelProveedor && !productIdsDelProveedor.has(p.id)) return false;
    if (query.trim()) {
      const texto = `${p.codigo} ${p.descripcion} ${p.tipo} ${p.subtipo || ''} ${p.departamento} ${p.marca || ''} ${(p.colores || []).join(' ')}`.toLowerCase();
      if (!texto.includes(query.toLowerCase())) return false;
    }
    return true;
  });

  // Agrupar por marca
  const grupos = filtered.reduce((acc, p) => {
    const marca = p.marca || '(Sin marca)';
    if (!acc[marca]) acc[marca] = [];
    acc[marca].push(p);
    return acc;
  }, {});
  const marcasOrdenadas = Object.keys(grupos).sort((a, b) => a.localeCompare(b));

  const removeProduct = (id) => setProducts(products.filter((p) => p.id !== id));

  const saveEdit = (updated) => {
    setProducts(products.map((p) => (p.id === updated.id ? updated : p)));
    setEditingProductId(null);
  };

  const [marcasColapsadas, setMarcasColapsadas] = useState({});
  const toggleMarca = (m) => setMarcasColapsadas((prev) => ({ ...prev, [m]: !prev[m] }));

  // Opciones únicas para los selectores (solo proveedores de este origen)
  const proveedoresUnicos = [...new Set((suppliers || []).map((s) => s.nombre))].sort();
  const limpiarFiltros = () => {
    setFiltroProveedor('');
    setFiltroDepartamento('');
    setFiltroMarca('');
    setQuery('');
  };

  return (
    <div className="space-y-4">
      <div className="flex gap-2">
        <div className="relative flex-1">
          <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-app-dim" />
          <input
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="Buscar por descripción, código, tipo…"
            className="w-full bg-app-panel border border-app-line rounded-xl pl-9 pr-3 py-3 text-sm focus:outline-none focus:ring-2 focus:ring-amber-400"
          />
        </div>
        <button
          onClick={onOpenConfig}
          className="shrink-0 w-12 bg-app-panel border border-app-line rounded-xl flex items-center justify-center text-app-dim2 active:bg-app-active"
          aria-label="Departamentos y tipos"
        >
          <Settings size={18} />
        </button>
      </div>

      {/* Filtros: Proveedor, Departamento, Marca */}
      <div className="bg-app-panel border border-app-line rounded-xl p-3 space-y-2">
        <div className="flex items-center justify-between mb-1">
          <label className="text-xs uppercase tracking-wide text-app-dim2">Filtros</label>
          {hayFiltroActivo && (
            <button onClick={limpiarFiltros} className="text-xs text-app-red2 active:opacity-70">
              Limpiar
            </button>
          )}
        </div>
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
          <SearchableSelect
            id="filtro-proveedor"
            openId={openFiltroId}
            setOpenId={setOpenFiltroId}
            value={filtroProveedor}
            onChange={setFiltroProveedor}
            options={proveedoresUnicos}
            placeholder="Proveedor"
          />
          <SearchableSelect
            id="filtro-departamento"
            openId={openFiltroId}
            setOpenId={setOpenFiltroId}
            value={filtroDepartamento}
            onChange={setFiltroDepartamento}
            options={departamentos}
            placeholder="Departamento"
          />
          <SearchableSelect
            id="filtro-marca"
            openId={openFiltroId}
            setOpenId={setOpenFiltroId}
            value={filtroMarca}
            onChange={setFiltroMarca}
            options={marcas}
            placeholder="Marca"
          />
        </div>
      </div>

      {!showForm ? (
        <button
          onClick={() => setShowForm(true)}
          className="w-full py-3 rounded-xl border border-dashed border-app-line3 text-sm text-app-dim2 flex items-center justify-center gap-2 active:bg-app-panel"
        >
          <Plus size={16} /> Agregar producto
        </button>
      ) : (
        <ProductForm
          products={products}
          departamentos={departamentos}
          tipos={tipos}
          marcas={marcas}
          ciudades={ciudades}
          fabricas={fabricas}
          factores={factores}
          tasaCambio={tasaCambio}
          origen={origen}
          usuarioActivoNombre={usuarioActivoNombre}
          usuarioActivoPrefijo={usuarioActivoPrefijo}
          onUpdateTipos={setTipos}
          onCancel={() => setShowForm(false)}
          onSave={(nuevos) => { setProducts([...products, ...nuevos]); setShowForm(false); }}
          onCreateMarca={onCreateMarca}
          onCreateFabrica={onCreateFabrica}
          onCreateCiudad={onCreateCiudad}
        />
      )}

      {/* Sin filtros: mensaje limpio */}
      {!hayFiltroActivo && !showForm && (
        <div className="text-center py-12 text-app-dim2">
          <Search size={28} className="mx-auto mb-2 text-app-dim3" />
          <p className="text-sm">Elige un filtro o escribe en el buscador para ver productos.</p>
          <p className="text-xs text-app-dim mt-1">{products.length} productos en el catálogo.</p>
        </div>
      )}

      {/* Con filtros pero sin resultados */}
      {hayFiltroActivo && filtered.length === 0 && !showForm && (
        <p className="text-center text-xs text-app-dim py-6">Sin productos para esos filtros.</p>
      )}

      <div className="space-y-3">
        {marcasOrdenadas.map((marca) => {
          const colapsada = marcasColapsadas[marca];
          const grupo = grupos[marca];
          const totalPzs = 0;
          return (
            <div key={marca} className="rounded-2xl border border-app-line overflow-hidden">
              {/* Cabecera de marca */}
              <button
                onClick={() => toggleMarca(marca)}
                className="w-full flex items-center justify-between px-4 py-3 bg-app-active"
              >
                <div className="flex items-center gap-2">
                  <Tag size={14} className="text-app-gold shrink-0" />
                  <span className="text-sm font-semibold text-app-white">{marca}</span>
                  <span className="text-xs text-app-dim2 bg-app-bg rounded-full px-2 py-0.5">
                    {grupo.length} artículo{grupo.length !== 1 ? 's' : ''}
                  </span>
                </div>
                <ChevronDown size={16} className={`text-app-dim transition-transform ${colapsada ? '-rotate-90' : ''}`} />
              </button>

              {/* Productos del grupo */}
              {!colapsada && (
                <div className="divide-y divide-app-line">
                  {grupo.map((p) => (
                    <div key={p.id} className="bg-app-panel">
                      <button
                        onClick={() => setExpandedProductId(expandedProductId === p.id ? null : p.id)}
                        className="w-full flex items-center gap-3 px-3 py-3 text-left"
                      >
                        {listaFotos(p).length ? (
                          <img
                            src={urlFoto(listaFotos(p)[0])}
                            alt={p.descripcion}
                            className="w-11 h-11 object-cover rounded-lg shrink-0 cursor-zoom-in"
                            onClick={(e) => { e.stopPropagation(); setVisor({ fotos: listaFotos(p), titulo: p.descripcion || p.codigo }); }}
                          />
                        ) : (
                          <div className="w-11 h-11 rounded-lg bg-app-bg border border-app-line flex items-center justify-center shrink-0">
                            <ImageOff size={15} className="text-app-dim3" />
                          </div>
                        )}
                        <div className="min-w-0 flex-1">
                          <div className="flex items-center gap-1.5">
                            <span className="text-xs font-mono text-app-gold bg-app-goldbg rounded px-1.5 py-0.5 shrink-0">{p.codigo}</span>
                            <p className="text-sm font-mono truncate">{p.descripcion}</p>
                          </div>
                          <p className="text-xs text-app-dim truncate">{p.departamento} · {p.tipo}{p.subtipo ? ` (${p.subtipo})` : ''}</p>
                          <p className="text-xs text-app-dim2">
                            {fmtMoneda(p.costoMonto, p.costoMoneda)} · Venta {fmtLempiras(p.ventaLempiras)}
                          </p>
                        </div>
                        <ChevronDown size={15} className={`text-app-dim shrink-0 transition-transform ${expandedProductId === p.id ? 'rotate-180' : ''}`} />
                      </button>

                      {expandedProductId === p.id && (
                        editingProductId === p.id ? (
                          <div className="px-3 pb-3">
                            <EditProductForm
                              product={p}
                              products={products}
                              departamentos={departamentos}
                              tipos={tipos}
                              marcas={marcas}
                              ciudades={ciudades}
                              fabricas={fabricas}
                              onSave={saveEdit}
                              onCancel={() => setEditingProductId(null)}
                            />
                          </div>
                        ) : (
                          <div className="px-3 pb-3 border-t border-app-line mt-0 pt-3">
                            {(p.ciudad || p.fabrica) && (
                              <div className="flex gap-3 mb-2">
                                {p.ciudad && (
                                  <span className="text-xs text-app-dim2">🏙 <span className="text-app-light">{p.ciudad}</span></span>
                                )}
                                {p.fabrica && (
                                  <span className="text-xs text-app-dim2">🏭 <span className="text-app-light">{p.fabrica}</span></span>
                                )}
                              </div>
                            )}
                            <div className="flex flex-wrap gap-1.5 mb-3">
                              {(p.variantes || []).map((v) => (
                                <span key={variantKey(v.talla, v.color)} className="text-xs bg-app-bg border border-app-line rounded-full px-2 py-1">
                                  {varLabel(v)}: <span className="text-app-gold font-medium">{v.cantidad}</span>
                                </span>
                              ))}
                            </div>
                            <div className="flex items-center gap-3">
                              <button onClick={() => setEditingProductId(p.id)} className="text-xs text-app-sky flex items-center gap-1">
                                <Pencil size={13} /> Editar
                              </button>
                              {puedoBorrar && (
                                <BotonBorrar
                                  onConfirm={() => removeProduct(p.id)}
                                  size={13}
                                  texto="Eliminar"
                                  aviso={(() => {
                                    const n = (orders || []).filter((o) => o.items.some((it) => it.productId === p.id)).length;
                                    return n > 0 ? `⚠ Está en ${n} pedido${n !== 1 ? 's' : ''}` : '';
                                  })()}
                                />
                              )}
                            </div>
                          </div>
                        )
                      )}
                    </div>
                  ))}
                </div>
              )}
            </div>
          );
        })}
      </div>

      {visor && (
        <VisorFotos
          fotos={visor.fotos}
          titulo={visor.titulo}
          onCerrar={() => setVisor(null)}
        />
      )}
    </div>
  );
}

// ---------- Config: Departamentos y Tipos ----------
// ---------- Datos para importar desde listado de tienda ----------
const IMPORT_DEPARTAMENTOS = [
  "ABARROTERIA",
  "ACCESORIOS DAMA",
  "ACCESORIOS DE CABALLERO",
  "ACCESORIOS DE NINAS",
  "ACCESORIOS DE VEHICULOS",
  "ACCESORIOS MEDICOS",
  "ARTICULOS DE NAVIDAD",
  "ARTICULOS DE VERANO",
  "ARTICULOS DEL HOGAR",
  "ARTICULOS ESCOLARES",
  "ARTICULOS PROMOCIONALES",
  "BEBE",
  "CARTERAS",
  "COSMETICOS Y FRAGANCIAS IMP",
  "DEPORTES",
  "ELECTRONICA",
  "HIGIENE PERSONAL",
  "JOYERIA",
  "JUGUETERIA",
  "MALETAS",
  "MASCOTAS",
  "OFERTAS",
  "OPTICA",
  "PERFUMERIA",
  "REGALOS",
  "ROPA AMERICANA DAMA",
  "ROPA DAMA",
  "ROPA DE CABALLEROS",
  "ROPA DE NINAS",
  "ROPA DE NIÑOS",
  "ROPA DEPORTIVA DE CABALLEROS",
  "ROPA DEPORTIVA DE DAMAS",
  "ROPA INT. DAMA",
  "ROPA INT. DE CABALLERO",
  "ROPA INTERIOR DE NINAS",
  "ROPA INTERIOR DE NINOS",
  "ROPA JUVENIL NINAS",
  "ROPA JUVENIL NINOS",
  "TRAJES FORMALES",
  "ZAPATO DE CABALLERO",
  "ZAPATO DE DAMA",
  "ZAPATO DE NIÑA",
  "ZAPATO DE NIÑO",
  "MUEBLERIA",
];
const IMPORT_TIPOS = [
  { nombre: "Abanico", medida: 'simple', tallas: [] },
  { nombre: "Abre Latas", medida: 'simple', tallas: [] },
  { nombre: "Abrigo", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Accesorio Bebe", medida: 'simple', tallas: [] },
  { nombre: "Accesorio Cocina", medida: 'simple', tallas: [] },
  { nombre: "Accesorio Crocs", medida: 'simple', tallas: [] },
  { nombre: "Accesorios", medida: 'simple', tallas: [] },
  { nombre: "Accesorios Celulares", medida: 'simple', tallas: [] },
  { nombre: "Accesorios Dama", medida: 'simple', tallas: [] },
  { nombre: "Accesorios Zapato", medida: 'simple', tallas: [] },
  { nombre: "Aceite", medida: 'simple', tallas: [] },
  { nombre: "Aceitera", medida: 'simple', tallas: [] },
  { nombre: "Acondicionador", medida: 'simple', tallas: [] },
  { nombre: "Adaptador", medida: 'simple', tallas: [] },
  { nombre: "Adorno", medida: 'simple', tallas: [] },
  { nombre: "Adorno Nav", medida: 'simple', tallas: [] },
  { nombre: "Afeitadora", medida: 'simple', tallas: [] },
  { nombre: "Agenda", medida: 'simple', tallas: [] },
  { nombre: "Agua de Rosa", medida: 'simple', tallas: [] },
  { nombre: "Agua en bote", medida: 'simple', tallas: [] },
  { nombre: "Agua Oxigenada", medida: 'simple', tallas: [] },
  { nombre: "Ajedrez", medida: 'simple', tallas: [] },
  { nombre: "Album", medida: 'simple', tallas: [] },
  { nombre: "Alcancia", medida: 'simple', tallas: [] },
  { nombre: "Alcohol", medida: 'simple', tallas: [] },
  { nombre: "Aletas", medida: 'simple', tallas: [] },
  { nombre: "Alfombra", medida: 'simple', tallas: [] },
  { nombre: "Algodón", medida: 'simple', tallas: [] },
  { nombre: "Alicate", medida: 'simple', tallas: [] },
  { nombre: "Almohada", medida: 'simple', tallas: [] },
  { nombre: "Almohadilla Asiento Carro", medida: 'simple', tallas: [] },
  { nombre: "Almohadilla Facial", medida: 'simple', tallas: [] },
  { nombre: "Andadera", medida: 'simple', tallas: [] },
  { nombre: "Andador", medida: 'simple', tallas: [] },
  { nombre: "Anillo", medida: 'simple', tallas: [] },
  { nombre: "Anillo Servilletero", medida: 'simple', tallas: [] },
  { nombre: "Anti Ronquidos", medida: 'simple', tallas: [] },
  { nombre: "AntiFace", medida: 'simple', tallas: [] },
  { nombre: "Aplicador Maquillaje", medida: 'simple', tallas: [] },
  { nombre: "Aplicadores", medida: 'simple', tallas: [] },
  { nombre: "Arbol Navideño", medida: 'simple', tallas: [] },
  { nombre: "Arenero", medida: 'simple', tallas: [] },
  { nombre: "Aretes", medida: 'simple', tallas: [] },
  { nombre: "Argollas", medida: 'simple', tallas: [] },
  { nombre: "Aro de Luz", medida: 'simple', tallas: [] },
  { nombre: "Aro para Graduar", medida: 'simple', tallas: [] },
  { nombre: "Aromatizante", medida: 'simple', tallas: [] },
  { nombre: "Arreglo Floral", medida: 'simple', tallas: [] },
  { nombre: "Arreglo Regalo", medida: 'simple', tallas: [] },
  { nombre: "Arreglo Sastre", medida: 'simple', tallas: [] },
  { nombre: "Articulos Hogar", medida: 'simple', tallas: [] },
  { nombre: "Asiento Carro Bebe", medida: 'simple', tallas: [] },
  { nombre: "Asiento Inodoro", medida: 'simple', tallas: [] },
  { nombre: "Aspirador", medida: 'simple', tallas: [] },
  { nombre: "Aspiradora", medida: 'simple', tallas: [] },
  { nombre: "Atrapa Sueño", medida: 'simple', tallas: [] },
  { nombre: "Audifonos", medida: 'simple', tallas: [] },
  { nombre: "Auriculares", medida: 'simple', tallas: [] },
  { nombre: "Avion", medida: 'simple', tallas: [] },
  { nombre: "Azucarera", medida: 'simple', tallas: [] },
  { nombre: "Babero", medida: 'simple', tallas: [] },
  { nombre: "Baby Doll", medida: 'simple', tallas: [] },
  { nombre: "Balon", medida: 'simple', tallas: [] },
  { nombre: "Balsamo", medida: 'simple', tallas: [] },
  { nombre: "Banda", medida: 'simple', tallas: [] },
  { nombre: "Bandeja", medida: 'simple', tallas: [] },
  { nombre: "Banner", medida: 'simple', tallas: [] },
  { nombre: "Banquito", medida: 'simple', tallas: [] },
  { nombre: "Bañera", medida: 'simple', tallas: [] },
  { nombre: "Bar", medida: 'simple', tallas: [] },
  { nombre: "Barajas", medida: 'simple', tallas: [] },
  { nombre: "Bascula", medida: 'simple', tallas: [] },
  { nombre: "Base", medida: 'simple', tallas: [] },
  { nombre: "Base Arbol", medida: 'simple', tallas: [] },
  { nombre: "Base Cama", medida: 'simple', tallas: [] },
  { nombre: "Baston", medida: 'simple', tallas: [] },
  { nombre: "Basurero", medida: 'simple', tallas: [] },
  { nombre: "Bata", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Baterias", medida: 'simple', tallas: [] },
  { nombre: "Batido", medida: 'simple', tallas: [] },
  { nombre: "Batidora", medida: 'simple', tallas: [] },
  { nombre: "Bebidas", medida: 'simple', tallas: [] },
  { nombre: "Bermuda", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Betun", medida: 'simple', tallas: [] },
  { nombre: "Biberon", medida: 'simple', tallas: [] },
  { nombre: "Biberon Juguetes", medida: 'simple', tallas: [] },
  { nombre: "Bicicleta", medida: 'simple', tallas: [] },
  { nombre: "Biker", medida: 'simple', tallas: [] },
  { nombre: "Bikini", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Billetera", medida: 'simple', tallas: [] },
  { nombre: "Bingo", medida: 'simple', tallas: [] },
  { nombre: "Blazer", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Blusa", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Blush", medida: 'simple', tallas: [] },
  { nombre: "Bluson", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Bocina", medida: 'simple', tallas: [] },
  { nombre: "Body Bebe", medida: 'simple', tallas: [] },
  { nombre: "Body Suit", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Boina", medida: 'simple', tallas: [] },
  { nombre: "Bola Decorativa", medida: 'simple', tallas: [] },
  { nombre: "Bolero", medida: 'simple', tallas: [] },
  { nombre: "Boligrafo", medida: 'simple', tallas: [] },
  { nombre: "Bolsa", medida: 'simple', tallas: [] },
  { nombre: "Bolsa Decorada", medida: 'simple', tallas: [] },
  { nombre: "Bolsa Impermiable", medida: 'simple', tallas: [] },
  { nombre: "Bolsa Lentes", medida: 'simple', tallas: [] },
  { nombre: "Bolsa Mochila", medida: 'simple', tallas: [] },
  { nombre: "Bolsa Orina", medida: 'simple', tallas: [] },
  { nombre: "Bolsa Regalo", medida: 'simple', tallas: [] },
  { nombre: "Bolsa Termica", medida: 'simple', tallas: [] },
  { nombre: "Bolso", medida: 'simple', tallas: [] },
  { nombre: "Bolso Termico", medida: 'simple', tallas: [] },
  { nombre: "Bomba", medida: 'simple', tallas: [] },
  { nombre: "Bomba Aire", medida: 'simple', tallas: [] },
  { nombre: "Bombonera", medida: 'simple', tallas: [] },
  { nombre: "Bomper", medida: 'simple', tallas: [] },
  { nombre: "Bota", medida: 'simple', tallas: [] },
  { nombre: "Botellas", medida: 'simple', tallas: [] },
  { nombre: "Botin", medida: 'simple', tallas: [] },
  { nombre: "Botiquin", medida: 'simple', tallas: [] },
  { nombre: "Boxer", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Bralette", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Brasalete", medida: 'simple', tallas: [] },
  { nombre: "Brassier", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Brillo Labial", medida: 'simple', tallas: [] },
  { nombre: "Brocha", medida: 'simple', tallas: [] },
  { nombre: "Broche", medida: 'simple', tallas: [] },
  { nombre: "Bufanda", medida: 'simple', tallas: [] },
  { nombre: "Burbuja", medida: 'simple', tallas: [] },
  { nombre: "Buzo", medida: 'simple', tallas: [] },
  { nombre: "Cable", medida: 'simple', tallas: [] },
  { nombre: "Cabri", medida: 'simple', tallas: [] },
  { nombre: "Cachetero", medida: 'simple', tallas: [] },
  { nombre: "Cadena", medida: 'simple', tallas: [] },
  { nombre: "Cagador Bebe", medida: 'simple', tallas: [] },
  { nombre: "Caja", medida: 'simple', tallas: [] },
  { nombre: "Caja Para Joyeria", medida: 'simple', tallas: [] },
  { nombre: "Caja Para Regalo", medida: 'simple', tallas: [] },
  { nombre: "Caja Registradora", medida: 'simple', tallas: [] },
  { nombre: "Calcamonias", medida: 'simple', tallas: [] },
  { nombre: "Calceta", medida: 'simple', tallas: [] },
  { nombre: "Calcetines", medida: 'simple', tallas: [] },
  { nombre: "Calzador", medida: 'simple', tallas: [] },
  { nombre: "Calzoncillo", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Cama", medida: 'simple', tallas: [] },
  { nombre: "Cama Mascota", medida: 'simple', tallas: [] },
  { nombre: "Cama Portatil", medida: 'simple', tallas: [] },
  { nombre: "Camara", medida: 'simple', tallas: [] },
  { nombre: "Camara Carro", medida: 'simple', tallas: [] },
  { nombre: "Cambiador", medida: 'simple', tallas: [] },
  { nombre: "Camino Mesa", medida: 'simple', tallas: [] },
  { nombre: "Camion", medida: 'simple', tallas: [] },
  { nombre: "Camisa", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Camisa Polo", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Camiseta", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Camison", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Campana", medida: 'simple', tallas: [] },
  { nombre: "Canasta", medida: 'simple', tallas: [] },
  { nombre: "Candado", medida: 'simple', tallas: [] },
  { nombre: "Candelabro", medida: 'simple', tallas: [] },
  { nombre: "Cangurera", medida: 'simple', tallas: [] },
  { nombre: "Capote", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Capri", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Capucha", medida: 'simple', tallas: [] },
  { nombre: "Cardigan", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Careta", medida: 'simple', tallas: [] },
  { nombre: "Cargador Celular", medida: 'simple', tallas: [] },
  { nombre: "Carro", medida: 'simple', tallas: [] },
  { nombre: "Cartera", medida: 'simple', tallas: [] },
  { nombre: "Casa Juguete", medida: 'simple', tallas: [] },
  { nombre: "Casa Para Mascotas", medida: 'simple', tallas: [] },
  { nombre: "Cascanueces", medida: 'simple', tallas: [] },
  { nombre: "Casco", medida: 'simple', tallas: [] },
  { nombre: "Castillo", medida: 'simple', tallas: [] },
  { nombre: "Cejas", medida: 'simple', tallas: [] },
  { nombre: "Centro", medida: 'simple', tallas: [] },
  { nombre: "Centro de Mesa", medida: 'simple', tallas: [] },
  { nombre: "Cepillo", medida: 'simple', tallas: [] },
  { nombre: "Cepillo Baño", medida: 'simple', tallas: [] },
  { nombre: "Cepillo Cabello", medida: 'simple', tallas: [] },
  { nombre: "Cepillo Cejas", medida: 'simple', tallas: [] },
  { nombre: "Cepillo Dental", medida: 'simple', tallas: [] },
  { nombre: "Cepillo Facial", medida: 'simple', tallas: [] },
  { nombre: "Cepillo Mascota", medida: 'simple', tallas: [] },
  { nombre: "Cepillo p/Biberon", medida: 'simple', tallas: [] },
  { nombre: "Cepillo Pestañas", medida: 'simple', tallas: [] },
  { nombre: "Cepillo Pie", medida: 'simple', tallas: [] },
  { nombre: "Cera", medida: 'simple', tallas: [] },
  { nombre: "Cesta", medida: 'simple', tallas: [] },
  { nombre: "Chaleco", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Chaqueta", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Charm", medida: 'simple', tallas: [] },
  { nombre: "Chicles", medida: 'simple', tallas: [] },
  { nombre: "Chimenea", medida: 'simple', tallas: [] },
  { nombre: "Chocolates", medida: 'simple', tallas: [] },
  { nombre: "Choker", medida: 'simple', tallas: [] },
  { nombre: "Chongo Regalo", medida: 'simple', tallas: [] },
  { nombre: "Chumpa", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Chupete", medida: 'simple', tallas: [] },
  { nombre: "Chupon", medida: 'simple', tallas: [] },
  { nombre: "Churros", medida: 'simple', tallas: [] },
  { nombre: "Cinta", medida: 'simple', tallas: [] },
  { nombre: "Cinta Cabello", medida: 'simple', tallas: [] },
  { nombre: "Cinta Decorativa", medida: 'simple', tallas: [] },
  { nombre: "Cinta Regalo", medida: 'simple', tallas: [] },
  { nombre: "Cinturon", medida: 'simple', tallas: [] },
  { nombre: "Cobertor", medida: 'simple', tallas: [] },
  { nombre: "Cobertor Maleta", medida: 'simple', tallas: [] },
  { nombre: "Cobertor Traje", medida: 'simple', tallas: [] },
  { nombre: "Coche", medida: 'simple', tallas: [] },
  { nombre: "Codera", medida: 'simple', tallas: [] },
  { nombre: "Cofre", medida: 'simple', tallas: [] },
  { nombre: "Cohete", medida: 'simple', tallas: [] },
  { nombre: "Cojin", medida: 'simple', tallas: [] },
  { nombre: "Cola", medida: 'simple', tallas: [] },
  { nombre: "Colcha", medida: 'simple', tallas: [] },
  { nombre: "Colchon", medida: 'simple', tallas: [] },
  { nombre: "Colgante", medida: 'simple', tallas: [] },
  { nombre: "Collar", medida: 'simple', tallas: [] },
  { nombre: "Collar Para Lentes", medida: 'simple', tallas: [] },
  { nombre: "Collar Para Mascarilla", medida: 'simple', tallas: [] },
  { nombre: "Collar Para Mascotas", medida: 'simple', tallas: [] },
  { nombre: "Colonia", medida: 'simple', tallas: [] },
  { nombre: "Condimentero", medida: 'simple', tallas: [] },
  { nombre: "Conjunto", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Copa", medida: 'simple', tallas: [] },
  { nombre: "Corbata", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Corbatin", medida: 'simple', tallas: [] },
  { nombre: "Corona", medida: 'simple', tallas: [] },
  { nombre: "Corral", medida: 'simple', tallas: [] },
  { nombre: "Correa Mascota", medida: 'simple', tallas: [] },
  { nombre: "Corrector", medida: 'simple', tallas: [] },
  { nombre: "Corset", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Corta Cuticula", medida: 'simple', tallas: [] },
  { nombre: "Corta Uñas", medida: 'simple', tallas: [] },
  { nombre: "Corta Uñas Mascota", medida: 'simple', tallas: [] },
  { nombre: "Cortador", medida: 'simple', tallas: [] },
  { nombre: "Cortina", medida: 'simple', tallas: [] },
  { nombre: "Cosmetiquera", medida: 'simple', tallas: [] },
  { nombre: "Crayones", medida: 'simple', tallas: [] },
  { nombre: "Crema", medida: 'simple', tallas: [] },
  { nombre: "Crema Afeitar", medida: 'simple', tallas: [] },
  { nombre: "Crema Depiladora", medida: 'simple', tallas: [] },
  { nombre: "Crema Para Cabello", medida: 'simple', tallas: [] },
  { nombre: "Crema Zapatos", medida: 'simple', tallas: [] },
  { nombre: "Crocs", medida: 'simple', tallas: [] },
  { nombre: "Crop Top", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Cuadro Decorativo", medida: 'simple', tallas: [] },
  { nombre: "Cubayera", medida: 'simple', tallas: [] },
  { nombre: "Cubeta", medida: 'simple', tallas: [] },
  { nombre: "Cubo Rubik", medida: 'simple', tallas: [] },
  { nombre: "Cubre Pezon", medida: 'simple', tallas: [] },
  { nombre: "Cucharas", medida: 'simple', tallas: [] },
  { nombre: "Cucharones", medida: 'simple', tallas: [] },
  { nombre: "Cuchillo", medida: 'simple', tallas: [] },
  { nombre: "Cuerda", medida: 'simple', tallas: [] },
  { nombre: "Cuna", medida: 'simple', tallas: [] },
  { nombre: "Curita", medida: 'simple', tallas: [] },
  { nombre: "Decoracion", medida: 'simple', tallas: [] },
  { nombre: "Delantal", medida: 'simple', tallas: [] },
  { nombre: "Delineador", medida: 'simple', tallas: [] },
  { nombre: "Depiladora", medida: 'simple', tallas: [] },
  { nombre: "Desenredante", medida: 'simple', tallas: [] },
  { nombre: "Desinfectante", medida: 'simple', tallas: [] },
  { nombre: "Desmaquillante", medida: 'simple', tallas: [] },
  { nombre: "Desodorante", medida: 'simple', tallas: [] },
  { nombre: "Desodorante Ambiental", medida: 'simple', tallas: [] },
  { nombre: "Desodorante Carro", medida: 'simple', tallas: [] },
  { nombre: "Destapador", medida: 'simple', tallas: [] },
  { nombre: "Diadema", medida: 'simple', tallas: [] },
  { nombre: "Difuminador", medida: 'simple', tallas: [] },
  { nombre: "Difusor", medida: 'simple', tallas: [] },
  { nombre: "Dije", medida: 'simple', tallas: [] },
  { nombre: "Dinosaurio", medida: 'simple', tallas: [] },
  { nombre: "Disfraz", medida: 'simple', tallas: [] },
  { nombre: "Dispensador", medida: 'simple', tallas: [] },
  { nombre: "Dispensador Jabon", medida: 'simple', tallas: [] },
  { nombre: "Display", medida: 'simple', tallas: [] },
  { nombre: "Domino", medida: 'simple', tallas: [] },
  { nombre: "Dona", medida: 'simple', tallas: [] },
  { nombre: "Drone", medida: 'simple', tallas: [] },
  { nombre: "Dulces", medida: 'simple', tallas: [] },
  { nombre: "Edredon", medida: 'simple', tallas: [] },
  { nombre: "Ejercitador", medida: 'simple', tallas: [] },
  { nombre: "Empaque Regalo", medida: 'simple', tallas: [] },
  { nombre: "Encendedor", medida: 'simple', tallas: [] },
  { nombre: "Encrespador", medida: 'simple', tallas: [] },
  { nombre: "Enjuague Bucal", medida: 'simple', tallas: [] },
  { nombre: "Ensaladera", medida: 'simple', tallas: [] },
  { nombre: "Enterizo", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Entretenedor", medida: 'simple', tallas: [] },
  { nombre: "Envase", medida: 'simple', tallas: [] },
  { nombre: "Escalera", medida: 'simple', tallas: [] },
  { nombre: "Escapulario", medida: 'simple', tallas: [] },
  { nombre: "Escurridor", medida: 'simple', tallas: [] },
  { nombre: "Esmalte", medida: 'simple', tallas: [] },
  { nombre: "Espada", medida: 'simple', tallas: [] },
  { nombre: "Espatula", medida: 'simple', tallas: [] },
  { nombre: "Espejo", medida: 'simple', tallas: [] },
  { nombre: "Esponja", medida: 'simple', tallas: [] },
  { nombre: "Espuma Limpiadora", medida: 'simple', tallas: [] },
  { nombre: "Espumador", medida: 'simple', tallas: [] },
  { nombre: "Esquinera", medida: 'simple', tallas: [] },
  { nombre: "Esterilizador", medida: 'simple', tallas: [] },
  { nombre: "Estrella", medida: 'simple', tallas: [] },
  { nombre: "Estuche", medida: 'simple', tallas: [] },
  { nombre: "Estufa", medida: 'simple', tallas: [] },
  { nombre: "Etiqueta Equipaje", medida: 'simple', tallas: [] },
  { nombre: "Etiquetas", medida: 'simple', tallas: [] },
  { nombre: "Exfoliante", medida: 'simple', tallas: [] },
  { nombre: "Exhibidor", medida: 'simple', tallas: [] },
  { nombre: "Exprimidor", medida: 'simple', tallas: [] },
  { nombre: "Extenciones Cabello", medida: 'simple', tallas: [] },
  { nombre: "Extendedor Brassier", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Extension Botones", medida: 'simple', tallas: [] },
  { nombre: "Extractor", medida: 'simple', tallas: [] },
  { nombre: "Faja", medida: 'simple', tallas: [] },
  { nombre: "Faja Talladora", medida: 'simple', tallas: [] },
  { nombre: "Fajuelos", medida: 'simple', tallas: [] },
  { nombre: "Falda", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Falda Arbol de Navidad", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Falda Para Cama", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Figura Decorativa", medida: 'simple', tallas: [] },
  { nombre: "Figuritas", medida: 'simple', tallas: [] },
  { nombre: "Fijador", medida: 'simple', tallas: [] },
  { nombre: "Florero", medida: 'simple', tallas: [] },
  { nombre: "Flores", medida: 'simple', tallas: [] },
  { nombre: "Flotador", medida: 'simple', tallas: [] },
  { nombre: "Forro", medida: 'simple', tallas: [] },
  { nombre: "Forro Asiento Carro", medida: 'simple', tallas: [] },
  { nombre: "Franela", medida: 'simple', tallas: [] },
  { nombre: "Frasco", medida: 'simple', tallas: [] },
  { nombre: "Frazada", medida: 'simple', tallas: [] },
  { nombre: "Freidora de Aire", medida: 'simple', tallas: [] },
  { nombre: "Frutero", medida: 'simple', tallas: [] },
  { nombre: "Fuente Decorativa", medida: 'simple', tallas: [] },
  { nombre: "Fundas", medida: 'simple', tallas: [] },
  { nombre: "Gafete", medida: 'simple', tallas: [] },
  { nombre: "Galletas", medida: 'simple', tallas: [] },
  { nombre: "Gancho", medida: 'simple', tallas: [] },
  { nombre: "Gancho Pañal", medida: 'simple', tallas: [] },
  { nombre: "Gargantia", medida: 'simple', tallas: [] },
  { nombre: "Gavetero", medida: 'simple', tallas: [] },
  { nombre: "Gel", medida: 'simple', tallas: [] },
  { nombre: "Gel Analgesico", medida: 'simple', tallas: [] },
  { nombre: "Gel Baño", medida: 'simple', tallas: [] },
  { nombre: "Gel Para Cabello", medida: 'simple', tallas: [] },
  { nombre: "Gel Para Cejas", medida: 'simple', tallas: [] },
  { nombre: "Gel Para Manos", medida: 'simple', tallas: [] },
  { nombre: "Gelatina", medida: 'simple', tallas: [] },
  { nombre: "Gimnasio", medida: 'simple', tallas: [] },
  { nombre: "Globo", medida: 'simple', tallas: [] },
  { nombre: "Gloss", medida: 'simple', tallas: [] },
  { nombre: "Gorra", medida: 'simple', tallas: [] },
  { nombre: "Gorro", medida: 'simple', tallas: [] },
  { nombre: "Grifo", medida: 'simple', tallas: [] },
  { nombre: "Guantes", medida: 'simple', tallas: [] },
  { nombre: "Guayabera", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Guirnalda", medida: 'simple', tallas: [] },
  { nombre: "Guitarra", medida: 'simple', tallas: [] },
  { nombre: "Gusano", medida: 'simple', tallas: [] },
  { nombre: "Helados", medida: 'simple', tallas: [] },
  { nombre: "Helicoptero", medida: 'simple', tallas: [] },
  { nombre: "Hervidor", medida: 'simple', tallas: [] },
  { nombre: "Hidratante", medida: 'simple', tallas: [] },
  { nombre: "Hielera", medida: 'simple', tallas: [] },
  { nombre: "Hilo Dental", medida: 'simple', tallas: [] },
  { nombre: "Hisopos", medida: 'simple', tallas: [] },
  { nombre: "Horma Zapato", medida: 'simple', tallas: [] },
  { nombre: "Horno", medida: 'simple', tallas: [] },
  { nombre: "Humidificador", medida: 'simple', tallas: [] },
  { nombre: "Iman", medida: 'simple', tallas: [] },
  { nombre: "Impermeable", medida: 'simple', tallas: [] },
  { nombre: "Individuales", medida: 'simple', tallas: [] },
  { nombre: "Jabon", medida: 'simple', tallas: [] },
  { nombre: "Jabon Intimo", medida: 'simple', tallas: [] },
  { nombre: "Jabon Liquido", medida: 'simple', tallas: [] },
  { nombre: "Jarra", medida: 'simple', tallas: [] },
  { nombre: "Jarron", medida: 'simple', tallas: [] },
  { nombre: "Jeans", medida: 'cintura_largo', cinturas: ["28", "30", "32", "34", "36", "38"], largos: ["30", "32", "34"] },
  { nombre: "Jegging", medida: 'simple', tallas: [] },
  { nombre: "Jenga", medida: 'simple', tallas: [] },
  { nombre: "Jogger", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Joyero", medida: 'simple', tallas: [] },
  { nombre: "Juego", medida: 'simple', tallas: [] },
  { nombre: "Juego Abecedario", medida: 'simple', tallas: [] },
  { nombre: "Juego Arte", medida: 'simple', tallas: [] },
  { nombre: "Juego Baloncesto", medida: 'simple', tallas: [] },
  { nombre: "Juego Baño", medida: 'simple', tallas: [] },
  { nombre: "Juego Beisbol", medida: 'simple', tallas: [] },
  { nombre: "Juego Bloques", medida: 'simple', tallas: [] },
  { nombre: "Juego Boliche", medida: 'simple', tallas: [] },
  { nombre: "Juego Bolos", medida: 'simple', tallas: [] },
  { nombre: "Juego Boxeo", medida: 'simple', tallas: [] },
  { nombre: "Juego Cocina", medida: 'simple', tallas: [] },
  { nombre: "Juego Comedor", medida: 'simple', tallas: [] },
  { nombre: "Juego de Mesa", medida: 'simple', tallas: [] },
  { nombre: "Juego Doctor", medida: 'simple', tallas: [] },
  { nombre: "Juego Educativo", medida: 'simple', tallas: [] },
  { nombre: "Juego Futbol", medida: 'simple', tallas: [] },
  { nombre: "Juego Hockey", medida: 'simple', tallas: [] },
  { nombre: "Juego Memoria", medida: 'simple', tallas: [] },
  { nombre: "Juego Parchis", medida: 'simple', tallas: [] },
  { nombre: "Juego Pesca", medida: 'simple', tallas: [] },
  { nombre: "Juego Playa", medida: 'simple', tallas: [] },
  { nombre: "Juego Serpientes", medida: 'simple', tallas: [] },
  { nombre: "Juego Sofa", medida: 'simple', tallas: [] },
  { nombre: "Juguete", medida: 'simple', tallas: [] },
  { nombre: "Juguete Accion", medida: 'simple', tallas: [] },
  { nombre: "Juguete Aprendizaje", medida: 'simple', tallas: [] },
  { nombre: "Juguete Araña", medida: 'simple', tallas: [] },
  { nombre: "Juguete Aspiradora", medida: 'simple', tallas: [] },
  { nombre: "Juguete Bebe", medida: 'simple', tallas: [] },
  { nombre: "Juguete Bocina", medida: 'simple', tallas: [] },
  { nombre: "Juguete Cocina", medida: 'simple', tallas: [] },
  { nombre: "Juguete Educativo", medida: 'simple', tallas: [] },
  { nombre: "Juguete Mascota", medida: 'simple', tallas: [] },
  { nombre: "Juguete Musical", medida: 'simple', tallas: [] },
  { nombre: "Juguete Sorpresa", medida: 'simple', tallas: [] },
  { nombre: "Jumper", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Keratina", medida: 'simple', tallas: [] },
  { nombre: "Kimono", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Kit", medida: 'simple', tallas: [] },
  { nombre: "Kit Cuidado Dental", medida: 'simple', tallas: [] },
  { nombre: "Kit Emergencia Carro", medida: 'simple', tallas: [] },
  { nombre: "Kit Lavado Carro", medida: 'simple', tallas: [] },
  { nombre: "Kit Maquillaje", medida: 'simple', tallas: [] },
  { nombre: "Kit Viaje", medida: 'simple', tallas: [] },
  { nombre: "Kleenex", medida: 'simple', tallas: [] },
  { nombre: "Lampara", medida: 'simple', tallas: [] },
  { nombre: "Lanzador", medida: 'simple', tallas: [] },
  { nombre: "Lapicera", medida: 'simple', tallas: [] },
  { nombre: "Lapiz Cejas", medida: 'simple', tallas: [] },
  { nombre: "Lapiz Labial", medida: 'simple', tallas: [] },
  { nombre: "Lapiz Ojo", medida: 'simple', tallas: [] },
  { nombre: "Lavador", medida: 'simple', tallas: [] },
  { nombre: "Leggins", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Legos", medida: 'simple', tallas: [] },
  { nombre: "Lente Tecnologico", medida: 'simple', tallas: [] },
  { nombre: "Lentes Natacion", medida: 'simple', tallas: [] },
  { nombre: "Lentes Para Fiesta", medida: 'simple', tallas: [] },
  { nombre: "Lentes Para Leer", medida: 'simple', tallas: [] },
  { nombre: "Lentes Para Sol", medida: 'simple', tallas: [] },
  { nombre: "Lentes Proteccion", medida: 'simple', tallas: [] },
  { nombre: "Leotardo", medida: 'simple', tallas: [] },
  { nombre: "Letras", medida: 'simple', tallas: [] },
  { nombre: "Levanta pompa", medida: 'simple', tallas: [] },
  { nombre: "Libreta", medida: 'simple', tallas: [] },
  { nombre: "Libro", medida: 'simple', tallas: [] },
  { nombre: "Licra", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Licuadora", medida: 'simple', tallas: [] },
  { nombre: "Liguero", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Lima", medida: 'simple', tallas: [] },
  { nombre: "Limpia Grifos", medida: 'simple', tallas: [] },
  { nombre: "Limpiador Poro", medida: 'simple', tallas: [] },
  { nombre: "Linterna", medida: 'simple', tallas: [] },
  { nombre: "Lip Gloss", medida: 'simple', tallas: [] },
  { nombre: "Lipstick", medida: 'simple', tallas: [] },
  { nombre: "Liston", medida: 'simple', tallas: [] },
  { nombre: "Llavero", medida: 'simple', tallas: [] },
  { nombre: "Lonchera", medida: 'simple', tallas: [] },
  { nombre: "Loteria Nacional", medida: 'simple', tallas: [] },
  { nombre: "Luces", medida: 'simple', tallas: [] },
  { nombre: "Macetera", medida: 'simple', tallas: [] },
  { nombre: "Maleta", medida: 'simple', tallas: [] },
  { nombre: "Maletin", medida: 'simple', tallas: [] },
  { nombre: "Malla", medida: 'simple', tallas: [] },
  { nombre: "Mamadera", medida: 'simple', tallas: [] },
  { nombre: "Mameluco", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Mancuernas", medida: 'simple', tallas: [] },
  { nombre: "Mando Bluetooth", medida: 'simple', tallas: [] },
  { nombre: "Manga", medida: 'simple', tallas: [] },
  { nombre: "Manicure", medida: 'simple', tallas: [] },
  { nombre: "Maniquera", medida: 'simple', tallas: [] },
  { nombre: "Mantas", medida: 'simple', tallas: [] },
  { nombre: "Mantel", medida: 'simple', tallas: [] },
  { nombre: "Maquillaje", medida: 'simple', tallas: [] },
  { nombre: "Maquina de Afeitar", medida: 'simple', tallas: [] },
  { nombre: "Maquina de Palomintas", medida: 'simple', tallas: [] },
  { nombre: "Marcadores", medida: 'simple', tallas: [] },
  { nombre: "Marco", medida: 'simple', tallas: [] },
  { nombre: "Mariquera", medida: 'simple', tallas: [] },
  { nombre: "Masajeador", medida: 'simple', tallas: [] },
  { nombre: "Mascara", medida: 'simple', tallas: [] },
  { nombre: "Mascarilla", medida: 'simple', tallas: [] },
  { nombre: "Mascarilla Facial", medida: 'simple', tallas: [] },
  { nombre: "Mascarilla Quirurgica", medida: 'simple', tallas: [] },
  { nombre: "Material", medida: 'simple', tallas: [] },
  { nombre: "Medias", medida: 'simple', tallas: [] },
  { nombre: "Medidor Presion", medida: 'simple', tallas: [] },
  { nombre: "Mesa", medida: 'simple', tallas: [] },
  { nombre: "Mesa Bebe", medida: 'simple', tallas: [] },
  { nombre: "Mochila", medida: 'simple', tallas: [] },
  { nombre: "Moises", medida: 'simple', tallas: [] },
  { nombre: "Molde", medida: 'simple', tallas: [] },
  { nombre: "Monedero", medida: 'simple', tallas: [] },
  { nombre: "Mono", medida: 'simple', tallas: [] },
  { nombre: "Monopoly", medida: 'simple', tallas: [] },
  { nombre: "Montura", medida: 'simple', tallas: [] },
  { nombre: "Mordedor", medida: 'simple', tallas: [] },
  { nombre: "Mosquitero", medida: 'simple', tallas: [] },
  { nombre: "Moto", medida: 'simple', tallas: [] },
  { nombre: "Mueble", medida: 'simple', tallas: [] },
  { nombre: "Muñeca", medida: 'simple', tallas: [] },
  { nombre: "Muñequera", medida: 'simple', tallas: [] },
  { nombre: "Nacimiento", medida: 'simple', tallas: [] },
  { nombre: "Nave", medida: 'simple', tallas: [] },
  { nombre: "Nebulizador", medida: 'simple', tallas: [] },
  { nombre: "Neceser", medida: 'simple', tallas: [] },
  { nombre: "Nevera", medida: 'simple', tallas: [] },
  { nombre: "Oferta", medida: 'simple', tallas: [] },
  { nombre: "Orejera", medida: 'simple', tallas: [] },
  { nombre: "Organizador", medida: 'simple', tallas: [] },
  { nombre: "Organizador  papel Toalla", medida: 'simple', tallas: [] },
  { nombre: "Organizador Baño", medida: 'simple', tallas: [] },
  { nombre: "Organizador Zapatos", medida: 'simple', tallas: [] },
  { nombre: "Overol", medida: 'simple', tallas: [] },
  { nombre: "Oximetro", medida: 'simple', tallas: [] },
  { nombre: "Pabellon", medida: 'simple', tallas: [] },
  { nombre: "Pacificador", medida: 'simple', tallas: [] },
  { nombre: "Pajaro", medida: 'simple', tallas: [] },
  { nombre: "Palette", medida: 'simple', tallas: [] },
  { nombre: "Palillos", medida: 'simple', tallas: [] },
  { nombre: "Palomitas", medida: 'simple', tallas: [] },
  { nombre: "Pantalla", medida: 'simple', tallas: [] },
  { nombre: "Pantalon", medida: 'cintura_largo', cinturas: ["28", "30", "32", "34", "36", "38"], largos: ["30", "32", "34"] },
  { nombre: "Pantaloneta", medida: 'cintura_largo', cinturas: ["28", "30", "32", "34", "36", "38"], largos: ["30", "32", "34"] },
  { nombre: "Pantimedia", medida: 'simple', tallas: [] },
  { nombre: "Pantuflas", medida: 'simple', tallas: [] },
  { nombre: "Panty", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Panty Hilo", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Pañal", medida: 'simple', tallas: [] },
  { nombre: "Pañalera", medida: 'simple', tallas: [] },
  { nombre: "Pañales Desechables", medida: 'simple', tallas: [] },
  { nombre: "Pañales Ecologicos", medida: 'simple', tallas: [] },
  { nombre: "Pañuelo", medida: 'simple', tallas: [] },
  { nombre: "Papel China", medida: 'simple', tallas: [] },
  { nombre: "Papel de Regalo", medida: 'simple', tallas: [] },
  { nombre: "Paragua", medida: 'simple', tallas: [] },
  { nombre: "Parcho", medida: 'simple', tallas: [] },
  { nombre: "Parlante", medida: 'simple', tallas: [] },
  { nombre: "Pasaportera", medida: 'simple', tallas: [] },
  { nombre: "Pascuas", medida: 'simple', tallas: [] },
  { nombre: "Pashmina", medida: 'simple', tallas: [] },
  { nombre: "Pasta Dental", medida: 'simple', tallas: [] },
  { nombre: "Paste", medida: 'simple', tallas: [] },
  { nombre: "Patineta", medida: 'simple', tallas: [] },
  { nombre: "Pechera", medida: 'simple', tallas: [] },
  { nombre: "Pedicure", medida: 'simple', tallas: [] },
  { nombre: "Pegamento Pestañas", medida: 'simple', tallas: [] },
  { nombre: "Pegamento Uñas", medida: 'simple', tallas: [] },
  { nombre: "Peine", medida: 'simple', tallas: [] },
  { nombre: "Peineta", medida: 'simple', tallas: [] },
  { nombre: "Pela Papas", medida: 'simple', tallas: [] },
  { nombre: "Pelota", medida: 'simple', tallas: [] },
  { nombre: "Peluche", medida: 'simple', tallas: [] },
  { nombre: "Percoladora", medida: 'simple', tallas: [] },
  { nombre: "Perfiladora", medida: 'simple', tallas: [] },
  { nombre: "Perfume", medida: 'simple', tallas: [] },
  { nombre: "Perro", medida: 'simple', tallas: [] },
  { nombre: "Pesa", medida: 'simple', tallas: [] },
  { nombre: "Pestañas", medida: 'simple', tallas: [] },
  { nombre: "Peticote", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Piedra Poma", medida: 'simple', tallas: [] },
  { nombre: "Piercing", medida: 'simple', tallas: [] },
  { nombre: "Pijama", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Pimentero", medida: 'simple', tallas: [] },
  { nombre: "Pinza", medida: 'simple', tallas: [] },
  { nombre: "Piscina", medida: 'simple', tallas: [] },
  { nombre: "Pista", medida: 'simple', tallas: [] },
  { nombre: "Pista Carro", medida: 'simple', tallas: [] },
  { nombre: "Pistola", medida: 'simple', tallas: [] },
  { nombre: "Pizarra", medida: 'simple', tallas: [] },
  { nombre: "Placa", medida: 'simple', tallas: [] },
  { nombre: "Plancha", medida: 'simple', tallas: [] },
  { nombre: "Plancha Cabello", medida: 'simple', tallas: [] },
  { nombre: "Planta", medida: 'simple', tallas: [] },
  { nombre: "Plastilina", medida: 'simple', tallas: [] },
  { nombre: "Plato", medida: 'simple', tallas: [] },
  { nombre: "Pluma", medida: 'simple', tallas: [] },
  { nombre: "Pokemon", medida: 'simple', tallas: [] },
  { nombre: "Polvos", medida: 'simple', tallas: [] },
  { nombre: "Poncho", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Ponsetia", medida: 'simple', tallas: [] },
  { nombre: "Pony", medida: 'simple', tallas: [] },
  { nombre: "Popit", medida: 'simple', tallas: [] },
  { nombre: "Porta Carnet", medida: 'simple', tallas: [] },
  { nombre: "Porta Pasaporte", medida: 'simple', tallas: [] },
  { nombre: "Porta Plato", medida: 'simple', tallas: [] },
  { nombre: "Porta Toallitas", medida: 'simple', tallas: [] },
  { nombre: "Porta Vasos", medida: 'simple', tallas: [] },
  { nombre: "Porta Velas", medida: 'simple', tallas: [] },
  { nombre: "Portador Mascota", medida: 'simple', tallas: [] },
  { nombre: "Prendedor", medida: 'simple', tallas: [] },
  { nombre: "Prensa Corbata", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Protector", medida: 'simple', tallas: [] },
  { nombre: "Protector Cabello", medida: 'simple', tallas: [] },
  { nombre: "Protector de Pezon", medida: 'simple', tallas: [] },
  { nombre: "Protector Pie", medida: 'simple', tallas: [] },
  { nombre: "Protector Solar", medida: 'simple', tallas: [] },
  { nombre: "Protector Solar Carro", medida: 'simple', tallas: [] },
  { nombre: "Publicidad", medida: 'simple', tallas: [] },
  { nombre: "Pulsera", medida: 'simple', tallas: [] },
  { nombre: "Quequera", medida: 'simple', tallas: [] },
  { nombre: "Quita Esmalte", medida: 'simple', tallas: [] },
  { nombre: "Radio", medida: 'simple', tallas: [] },
  { nombre: "Rallador", medida: 'simple', tallas: [] },
  { nombre: "Rama", medida: 'simple', tallas: [] },
  { nombre: "Raqueta", medida: 'simple', tallas: [] },
  { nombre: "Rascador", medida: 'simple', tallas: [] },
  { nombre: "Rasuradora", medida: 'simple', tallas: [] },
  { nombre: "Rebanador", medida: 'simple', tallas: [] },
  { nombre: "Recargas", medida: 'simple', tallas: [] },
  { nombre: "Recipiente", medida: 'simple', tallas: [] },
  { nombre: "Recipiente Mascota", medida: 'simple', tallas: [] },
  { nombre: "Regalia", medida: 'simple', tallas: [] },
  { nombre: "REGALOS", medida: 'simple', tallas: [] },
  { nombre: "Registradora", medida: 'simple', tallas: [] },
  { nombre: "Regla", medida: 'simple', tallas: [] },
  { nombre: "Relleno", medida: 'simple', tallas: [] },
  { nombre: "Reloj", medida: 'simple', tallas: [] },
  { nombre: "Removedor de Callos", medida: 'simple', tallas: [] },
  { nombre: "Removedor de Cuticula", medida: 'simple', tallas: [] },
  { nombre: "Removedor de Pelusa", medida: 'simple', tallas: [] },
  { nombre: "Removedor Para Dientes", medida: 'simple', tallas: [] },
  { nombre: "Reno", medida: 'simple', tallas: [] },
  { nombre: "Repelente", medida: 'simple', tallas: [] },
  { nombre: "Repuesto", medida: 'simple', tallas: [] },
  { nombre: "Retratera", medida: 'simple', tallas: [] },
  { nombre: "Rimel", medida: 'simple', tallas: [] },
  { nombre: "Rizador", medida: 'simple', tallas: [] },
  { nombre: "Robot", medida: 'simple', tallas: [] },
  { nombre: "Rodillera", medida: 'simple', tallas: [] },
  { nombre: "Rodillo", medida: 'simple', tallas: [] },
  { nombre: "Rompecabezas", medida: 'simple', tallas: [] },
  { nombre: "Ropa Oferta", medida: 'simple', tallas: [] },
  { nombre: "Ropero", medida: 'simple', tallas: [] },
  { nombre: "Rosario", medida: 'simple', tallas: [] },
  { nombre: "Rosquillas", medida: 'simple', tallas: [] },
  { nombre: "Rubor", medida: 'simple', tallas: [] },
  { nombre: "Rulos", medida: 'simple', tallas: [] },
  { nombre: "Sabanas", medida: 'simple', tallas: [] },
  { nombre: "Sabanilla", medida: 'simple', tallas: [] },
  { nombre: "Sacacorcho", medida: 'simple', tallas: [] },
  { nombre: "Sacapuntas", medida: 'simple', tallas: [] },
  { nombre: "Saco", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Sagrada Familia", medida: 'simple', tallas: [] },
  { nombre: "Salero", medida: 'simple', tallas: [] },
  { nombre: "Salida de Baño", medida: 'simple', tallas: [] },
  { nombre: "Sandalias", medida: 'simple', tallas: [] },
  { nombre: "Sandwichera", medida: 'simple', tallas: [] },
  { nombre: "Sanitizante", medida: 'simple', tallas: [] },
  { nombre: "Sarten", medida: 'simple', tallas: [] },
  { nombre: "Saya", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Secadora Cabello", medida: 'simple', tallas: [] },
  { nombre: "Semillas", medida: 'simple', tallas: [] },
  { nombre: "Servilletero", medida: 'simple', tallas: [] },
  { nombre: "Set", medida: 'simple', tallas: [] },
  { nombre: "Set Baño", medida: 'simple', tallas: [] },
  { nombre: "Set Barbero", medida: 'simple', tallas: [] },
  { nombre: "Set Bebe", medida: 'simple', tallas: [] },
  { nombre: "Set Belleza", medida: 'simple', tallas: [] },
  { nombre: "Set Buceo", medida: 'simple', tallas: [] },
  { nombre: "Set Caballero", medida: 'simple', tallas: [] },
  { nombre: "Set Escolar", medida: 'simple', tallas: [] },
  { nombre: "Set Estacionamiento", medida: 'simple', tallas: [] },
  { nombre: "Set Manicura", medida: 'simple', tallas: [] },
  { nombre: "Set Maquillaje", medida: 'simple', tallas: [] },
  { nombre: "Set Pedicure", medida: 'simple', tallas: [] },
  { nombre: "Set Recien Nacido", medida: 'simple', tallas: [] },
  { nombre: "Set Terraza", medida: 'simple', tallas: [] },
  { nombre: "Shampoo", medida: 'simple', tallas: [] },
  { nombre: "Short", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Silla", medida: 'simple', tallas: [] },
  { nombre: "Sillon", medida: 'simple', tallas: [] },
  { nombre: "Slime", medida: 'simple', tallas: [] },
  { nombre: "Smarwhatch", medida: 'simple', tallas: [] },
  { nombre: "Snack", medida: 'simple', tallas: [] },
  { nombre: "Sobres", medida: 'simple', tallas: [] },
  { nombre: "Sofa", medida: 'simple', tallas: [] },
  { nombre: "Sombras", medida: 'simple', tallas: [] },
  { nombre: "Sombrero", medida: 'simple', tallas: [] },
  { nombre: "Sombrilla", medida: 'simple', tallas: [] },
  { nombre: "Sonaja", medida: 'simple', tallas: [] },
  { nombre: "Soporte Bebe", medida: 'simple', tallas: [] },
  { nombre: "Soporte Celular", medida: 'simple', tallas: [] },
  { nombre: "Soporte Codo", medida: 'simple', tallas: [] },
  { nombre: "Soporte Cuello", medida: 'simple', tallas: [] },
  { nombre: "Soporte Mano", medida: 'simple', tallas: [] },
  { nombre: "Soporte Pantorrillas", medida: 'simple', tallas: [] },
  { nombre: "Soporte Rodilla", medida: 'simple', tallas: [] },
  { nombre: "Soporte Tobillo", medida: 'simple', tallas: [] },
  { nombre: "Splash", medida: 'simple', tallas: [] },
  { nombre: "Stickers", medida: 'simple', tallas: [] },
  { nombre: "Straple", medida: 'simple', tallas: [] },
  { nombre: "Sudadera", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Sueter", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Sujetador", medida: 'simple', tallas: [] },
  { nombre: "Tabla", medida: 'simple', tallas: [] },
  { nombre: "Talco", medida: 'simple', tallas: [] },
  { nombre: "Tallador", medida: 'simple', tallas: [] },
  { nombre: "Talquera", medida: 'simple', tallas: [] },
  { nombre: "Tambor", medida: 'simple', tallas: [] },
  { nombre: "Tanga", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Tanque", medida: 'simple', tallas: [] },
  { nombre: "Tape", medida: 'simple', tallas: [] },
  { nombre: "Tarjetas", medida: 'simple', tallas: [] },
  { nombre: "Tarjetero", medida: 'simple', tallas: [] },
  { nombre: "Tatuajes", medida: 'simple', tallas: [] },
  { nombre: "Taza", medida: 'simple', tallas: [] },
  { nombre: "Tazon", medida: 'simple', tallas: [] },
  { nombre: "Telefono", medida: 'simple', tallas: [] },
  { nombre: "Tennis", medida: 'simple', tallas: [] },
  { nombre: "Termo", medida: 'simple', tallas: [] },
  { nombre: "Termometro", medida: 'simple', tallas: [] },
  { nombre: "Tester", medida: 'simple', tallas: [] },
  { nombre: "Tetera", medida: 'simple', tallas: [] },
  { nombre: "Tetina", medida: 'simple', tallas: [] },
  { nombre: "Thermo", medida: 'simple', tallas: [] },
  { nombre: "Tiara", medida: 'simple', tallas: [] },
  { nombre: "Tienda Campaña", medida: 'simple', tallas: [] },
  { nombre: "Tijera", medida: 'simple', tallas: [] },
  { nombre: "Timbre", medida: 'simple', tallas: [] },
  { nombre: "Tinta", medida: 'simple', tallas: [] },
  { nombre: "Tinte", medida: 'simple', tallas: [] },
  { nombre: "Tirantes", medida: 'simple', tallas: [] },
  { nombre: "Toalla", medida: 'simple', tallas: [] },
  { nombre: "Toallas bebe", medida: 'simple', tallas: [] },
  { nombre: "Toallas Desmaquillantes", medida: 'simple', tallas: [] },
  { nombre: "Toallas Humedas", medida: 'simple', tallas: [] },
  { nombre: "Tobillera", medida: 'simple', tallas: [] },
  { nombre: "Tobimedias", medida: 'simple', tallas: [] },
  { nombre: "Tonico", medida: 'simple', tallas: [] },
  { nombre: "Top", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Torero", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Tortillera", medida: 'simple', tallas: [] },
  { nombre: "Tostadora", medida: 'simple', tallas: [] },
  { nombre: "Tractor", medida: 'simple', tallas: [] },
  { nombre: "Traje", medida: 'simple', tallas: [] },
  { nombre: "Traje de Baño", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Traje Formal", medida: 'simple', tallas: [] },
  { nombre: "Transformers", medida: 'simple', tallas: [] },
  { nombre: "Tratamiento", medida: 'simple', tallas: [] },
  { nombre: "Tratamiento Uñas", medida: 'simple', tallas: [] },
  { nombre: "Tren", medida: 'simple', tallas: [] },
  { nombre: "Triciclo", medida: 'simple', tallas: [] },
  { nombre: "Trineo", medida: 'simple', tallas: [] },
  { nombre: "Tripode", medida: 'simple', tallas: [] },
  { nombre: "Turbante", medida: 'simple', tallas: [] },
  { nombre: "Tutu", medida: 'simple', tallas: [] },
  { nombre: "Uniformes", medida: 'simple', tallas: [] },
  { nombre: "Uñas", medida: 'simple', tallas: [] },
  { nombre: "Utencilio Cocina", medida: 'simple', tallas: [] },
  { nombre: "Valvula", medida: 'simple', tallas: [] },
  { nombre: "Vaporizador", medida: 'simple', tallas: [] },
  { nombre: "Varios", medida: 'simple', tallas: [] },
  { nombre: "Vaselina", medida: 'simple', tallas: [] },
  { nombre: "Vaso", medida: 'simple', tallas: [] },
  { nombre: "Vela", medida: 'simple', tallas: [] },
  { nombre: "Velas Cumpleaños", medida: 'simple', tallas: [] },
  { nombre: "Venda", medida: 'simple', tallas: [] },
  { nombre: "Ventilador", medida: 'simple', tallas: [] },
  { nombre: "Vestido", medida: 'simple', tallas: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"] },
  { nombre: "Vicera", medida: 'simple', tallas: [] },
  { nombre: "VICKS", medida: 'simple', tallas: [] },
  { nombre: "Villa Navideña", medida: 'simple', tallas: [] },
  { nombre: "Vincha", medida: 'simple', tallas: [] },
  { nombre: "Visera", medida: 'simple', tallas: [] },
  { nombre: "Wafflera", medida: 'simple', tallas: [] },
  { nombre: "Walkie Talkie", medida: 'simple', tallas: [] },
  { nombre: "Zapatilla", medida: 'simple', tallas: [] },
  { nombre: "Zapatos", medida: 'simple', tallas: [] },
];

// ---------- Grupos predefinidos de tallas y subtipos ----------
const GRUPOS_TALLAS = [
  { label: 'Ropa adulto', tallas: ['XS', 'S', 'M', 'L', 'XL', '2XL', '3XL', '4XL'] },
  { label: 'Ropa niño', tallas: ['2', '4', '6', '8'] },
  { label: 'Ropa Juvenil', tallas: ['10', '12', '14', '16', '18'] },
  { label: 'Ropa bebé', tallas: ['0-3m', '3-6m', '6-12m', '12-18m', '18-24m'] },
  { label: 'Zapato adulto H', tallas: [
    '35/US3', '36/US4', '37/US5', '38/US6', '39/US7',
    '40/US8', '41/US9', '42/US10', '43/US11', '44/US12',
    '45/US13', '46/US14', '47/US15',
  ]},
  { label: 'Zapato adulto D', tallas: [
    '33/US2', '34/US3', '35/US5', '36/US6', '37/US7',
    '38/US8', '39/US9', '40/US10', '41/US11', '42/US12',
  ]},
  { label: 'Zapato niño', tallas: [
    '18/US3', '19/US4', '20/US4.5', '21/US5', '22/US6',
    '23/US7', '24/US8', '25/US9', '26/US9.5', '27/US10',
    '28/US11', '29/US12', '30/US12.5', '31/US13', '32/US1',
    '33/US2', '34/US3',
  ]},
  { label: 'Zapato bebé', tallas: [
    '14/US1', '15/US1.5', '16/US1.5', '17/US2', '18/US3',
  ]},
  { label: 'Cintura pantalón', tallas: ['28', '30', '32', '34', '36', '38', '40'] },
  { label: 'Largo pantalón', tallas: ['30', '32', '34'] },
  { label: 'Talla única', tallas: ['Única'] },
];

const GRUPOS_SUBTIPOS = [
  { label: 'Corte', items: ['Slim', 'Straight', 'Skinny', 'Regular', 'Relaxed', 'Bootcut', 'Wide Leg'] },
  { label: 'Cuello', items: ['Redondo', 'V', 'Polo', 'Henley', 'Turtleneck', 'Cuadrado'] },
  { label: 'Manga', items: ['Corta', 'Larga', 'Sin manga', '3/4', 'Ranglan'] },
  { label: 'Zapato', items: ['Oxford', 'Derby', 'Mocasín', 'Bota', 'Botín', 'Tenis', 'Sandalia', 'Chancla', 'Loafer'] },
  { label: 'Bolso', items: ['Tote', 'Clutch', 'Crossbody', 'Mochila', 'Cartera', 'Maleta'] },
  { label: 'Estilo', items: ['Casual', 'Formal', 'Deportivo', 'Clásico', 'Vintage', 'Oversized'] },
];

function PresetButtons({ grupos, onSelect, colorClass = 'bg-app-panel border-app-line3 text-app-dim2' }) {
  return (
    <div className="flex flex-wrap gap-1.5">
      {grupos.map((g) => (
        <button
          key={g.label}
          type="button"
          onClick={() => onSelect(g)}
          className={`text-xs border rounded-lg px-2.5 py-1.5 active:opacity-70 ${colorClass}`}
        >
          {g.label}
        </button>
      ))}
    </div>
  );
}

// ---------- Lista maestra de compradores para cargar de un tirón ----------
const EQUIPO_COMPRADORES = [
  { numero: 1, nombre: 'Vicente Carrion' },
  { numero: 2, nombre: 'Francisca Carrion' },
  { numero: 3, nombre: 'Dario Carrion' },
  { numero: 5, nombre: 'Mirna Carrion' },
  { numero: 6, nombre: 'Jenny Carrion' },
  { numero: 7, nombre: 'Tatiana Carrion' },
  { numero: 10, nombre: 'Hernan Alejandro' },
  { numero: 11, nombre: 'Antonio Cruz' },
  { numero: 12, nombre: 'Lilian Zuniga' },
  { numero: 13, nombre: 'Samyr Handal' },
  { numero: 15, nombre: 'Francis Carrion' },
  { numero: 16, nombre: 'Yamil Handal' },
  { numero: 17, nombre: 'Gustavo Barahona' },
  { numero: 18, nombre: 'Josue Argueta' },
  { numero: 19, nombre: 'Sulma Amaya' },
  { numero: 22, nombre: 'Diego Carrion' },
  { numero: 23, nombre: 'Leticia Gallegos' },
  { numero: 24, nombre: 'Johana Espinoza' },
  { numero: 25, nombre: 'Sandra Amaya' },
  { numero: 26, nombre: 'Dulce Galo' },
];

// ---------- Gestión de usuarios (Config) ----------
function UsuariosConfig({ usuarios = [], setUsuarios, usuarioActivo, onLogout }) {
  const [numero, setNumero] = useState('');
  const [nombre, setNombre] = useState('');
  const [prefijo, setPrefijo] = useState('');
  const [correo, setCorreo] = useState('');
  const [pin, setPin] = useState('');
  const [error, setError] = useState('');
  const [editandoCorreoId, setEditandoCorreoId] = useState(null);
  const [editCorreo, setEditCorreo] = useState('');
  const [guardando, setGuardando] = useState(false);
  const [confirmCarga, setConfirmCarga] = useState(false);
  const [editandoId, setEditandoId] = useState(null);
  const [editPrefijo, setEditPrefijo] = useState('');
  const [pinCambioUsuarioId, setPinCambioUsuarioId] = useState(null);
  const [pinCambioValor, setPinCambioValor] = useState('');
  const [pinCambioConfirm, setPinCambioConfirm] = useState('');

  // ¿Quién es administrador?
  // Regla: cualquier usuario con u.esAdmin === true. Como respaldo, Yamil Handal es admin
  // por defecto (para arrancar sin bloqueos si nadie está marcado).
  const esAdmin = (u) => {
    if (!u) return false;
    if (u.esAdmin) return true;
    // Respaldo: si nadie tiene esAdmin, Yamil Handal es admin
    const hayAdminExplicito = (usuarios || []).some((x) => x.esAdmin);
    if (!hayAdminExplicito && u.nombre && u.nombre.trim().toLowerCase() === 'yamil handal') return true;
    return false;
  };
  const soyAdmin = esAdmin(usuarioActivo);

  // Devuelve el prefijo efectivo de un usuario (el que use su código): manual o inicial del nombre
  const prefijoEfectivo = (u) => {
    if (u.prefijo && u.prefijo.trim()) return u.prefijo.trim().toUpperCase();
    return u.nombre ? u.nombre.trim().charAt(0).toUpperCase() : '?';
  };

  // Contar cuántos usuarios comparten cada prefijo (para el aviso de conflicto)
  const conteoPrefijos = usuarios.reduce((acc, u) => {
    const p = prefijoEfectivo(u);
    acc[p] = (acc[p] || 0) + 1;
    return acc;
  }, {});

  const agregar = async () => {
    setError('');
    const n = nombre.trim();
    const num = numero.trim();
    const pref = prefijo.trim().toUpperCase();
    if (!n) return setError('Escribe el nombre.');
    if (usuarios.some((u) => (u.nombre || '').toLowerCase() === n.toLowerCase())) return setError('Ya existe un usuario con ese nombre.');
    if (num && usuarios.some((u) => String(u.numero) === num)) return setError(`Ya existe un usuario con el número ${num}.`);
    if (pref && !/^[A-Z]{1,3}$/.test(pref)) return setError('El prefijo debe ser 1 a 3 letras (ej. F, Fr, Fra).');
    // PIN es OPCIONAL — si no se pone, el usuario queda con PIN pendiente
    let pinHash = null;
    if (pin) {
      if (!/^\d{4,6}$/.test(pin)) return setError('El PIN debe ser de 4 a 6 dígitos (o déjalo vacío).');
      setGuardando(true);
      pinHash = await hashPin(pin);
    }
    const mail = correo.trim().toLowerCase();
    if (mail && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(mail)) return setError('El correo no es válido.');
    if (mail && usuarios.some((u) => (u.email || '').toLowerCase() === mail)) return setError('Ya hay un comprador con ese correo.');
    setUsuarios([...usuarios, { id: uid(), numero: num ? parseInt(num, 10) : null, nombre: n, prefijo: pref || null, email: mail || null, pinHash }]);
    setNumero('');
    setNombre('');
    setPrefijo('');
    setCorreo('');
    setPin('');
    setGuardando(false);
  };

  const eliminar = (u) => {
    if (u.id === usuarioActivo?.id) return; // no puedes eliminarte a ti mismo
    setUsuarios(usuarios.filter((x) => x.id !== u.id));
  };

  const abrirEdicionCorreo = (u) => {
    setEditandoCorreoId(u.id);
    setEditCorreo(u.email || '');
    setError('');
  };

  const guardarCorreo = (u) => {
    const mail = editCorreo.trim().toLowerCase();
    if (mail && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(mail)) return setError('El correo no es válido.');
    if (mail && usuarios.some((x) => x.id !== u.id && (x.email || '').toLowerCase() === mail)) {
      return setError('Ese correo ya está asignado a otro comprador.');
    }
    setUsuarios(usuarios.map((x) => (x.id === u.id ? { ...x, email: mail || null } : x)));
    setEditandoCorreoId(null);
    setEditCorreo('');
    setError('');
  };

  const abrirEdicionPrefijo = (u) => {
    setEditandoId(u.id);
    setEditPrefijo(u.prefijo || '');
    setError('');
  };

  const guardarPrefijo = (u) => {
    const p = editPrefijo.trim().toUpperCase();
    if (p && !/^[A-Z]{1,3}$/.test(p)) {
      setError('El prefijo debe ser 1 a 3 letras.');
      return;
    }
    setUsuarios(usuarios.map((x) => x.id === u.id ? { ...x, prefijo: p || null } : x));
    setEditandoId(null);
    setEditPrefijo('');
  };

  // Cambia el rol de un comprador. Al pasar a supervisor, se le da acceso a
  // todos los países (paises se ignora cuando el rol no es 'comprador').
  const cambiarRol = (u, rol) => {
    setUsuarios(usuarios.map((x) => x.id === u.id ? { ...x, rol } : x));
  };

  // Activa/desactiva un país para un comprador
  const alternarPais = (u, paisId) => {
    const actuales = Array.isArray(u.paises) ? u.paises : ORIGENES.map((o) => o.id);
    const next = actuales.includes(paisId)
      ? actuales.filter((p) => p !== paisId)
      : [...actuales, paisId];
    setUsuarios(usuarios.map((x) => x.id === u.id ? { ...x, paises: next } : x));
  };

  // Cambiar/asignar PIN (solo administrador desde la lista)
  const abrirCambioPin = (u) => {
    setPinCambioUsuarioId(u.id);
    setPinCambioValor('');
    setPinCambioConfirm('');
    setError('');
  };

  const guardarNuevoPin = async (u) => {
    setError('');
    // Si dejas ambos vacíos: quitar el PIN (dejar en "Sin PIN")
    if (!pinCambioValor && !pinCambioConfirm) {
      setUsuarios(usuarios.map((x) => x.id === u.id ? { ...x, pinHash: null } : x));
      setPinCambioUsuarioId(null);
      return;
    }
    if (!/^\d{4,6}$/.test(pinCambioValor)) return setError('El PIN debe ser de 4 a 6 dígitos.');
    if (pinCambioValor !== pinCambioConfirm) return setError('Los PIN no coinciden.');
    const pinHash = await hashPin(pinCambioValor);
    setUsuarios(usuarios.map((x) => x.id === u.id ? { ...x, pinHash } : x));
    setPinCambioUsuarioId(null);
    setPinCambioValor('');
    setPinCambioConfirm('');
  };

  const cargarEquipo = () => {
    // Agregar solo los que no existan ya (por número o nombre)
    const nuevos = EQUIPO_COMPRADORES.filter((c) => {
      return !usuarios.some((u) =>
        String(u.numero) === String(c.numero) ||
        (u.nombre || '').toLowerCase() === (c.nombre || '').toLowerCase()
      );
    }).map((c) => ({ id: uid(), numero: c.numero, nombre: c.nombre, prefijo: null, pinHash: null }));
    setUsuarios([...usuarios, ...nuevos]);
    setConfirmCarga(false);
  };

  // Ordenar por número (los sin número al final)
  const usuariosOrdenados = [...usuarios].sort((a, b) => {
    const na = a.numero ? parseInt(a.numero, 10) : 9999;
    const nb = b.numero ? parseInt(b.numero, 10) : 9999;
    return na - nb;
  });

  const cuantosFaltan = EQUIPO_COMPRADORES.filter((c) => {
    return !usuarios.some((u) =>
      String(u.numero) === String(c.numero) ||
      (u.nombre || '').toLowerCase() === (c.nombre || '').toLowerCase()
    );
  }).length;

  return (
    <div className="space-y-3">
      {/* Sesión actual */}
      <div className="flex items-center justify-between bg-app-panel border border-app-line rounded-xl px-3 py-2.5">
        <div className="flex items-center gap-2.5">
          <span className="w-8 h-8 rounded-full bg-app-goldbg text-app-gold flex items-center justify-center font-bold text-sm">
            {usuarioActivo?.nombre?.charAt(0).toUpperCase()}
          </span>
          <div>
            <p className="text-sm font-medium">
              {usuarioActivo?.numero && <span className="text-app-dim2 mr-1">#{usuarioActivo.numero}</span>}
              {usuarioActivo?.nombre}
            </p>
            <p className="text-xs text-app-dim">Sesión activa en este dispositivo</p>
          </div>
        </div>
        <button onClick={onLogout} className="text-xs text-app-red2 border border-app-line rounded-lg px-2.5 py-1.5 active:bg-app-active">
          Cerrar sesión
        </button>
      </div>

      {/* Botón carga rápida del equipo */}
      {cuantosFaltan > 0 && (
        <div className="bg-app-blue border border-app-line2 rounded-xl p-3">
          {!confirmCarga ? (
            <button
              onClick={() => setConfirmCarga(true)}
              className="w-full text-sm text-app-sky font-medium py-1"
            >
              📋 Cargar equipo de compradores ({cuantosFaltan} pendientes)
            </button>
          ) : (
            <div className="space-y-2">
              <p className="text-xs text-app-light">
                Se agregarán {cuantosFaltan} compradores sin PIN. Podrán entrar tocando su nombre. Los PIN y prefijos personalizados se asignan después.
              </p>
              <div className="flex gap-2">
                <button onClick={() => setConfirmCarga(false)} className="flex-1 py-1.5 rounded-lg border border-app-line text-xs text-app-dim2">
                  Cancelar
                </button>
                <button onClick={cargarEquipo} className="flex-1 py-1.5 rounded-lg bg-app-sky text-app-bg text-xs font-semibold">
                  Sí, agregar {cuantosFaltan}
                </button>
              </div>
            </div>
          )}
        </div>
      )}

      {error && (
        <div className="flex items-center gap-2 bg-app-redbg text-app-red2 text-xs rounded-lg px-3 py-2">
          <AlertCircle size={14} className="shrink-0" /> {error}
        </div>
      )}

      {/* Alta manual de usuario */}
      <div className="bg-app-panel border border-app-line rounded-xl p-3 space-y-2">
        <p className="text-xs text-app-dim2 uppercase tracking-wide">Agregar comprador</p>
        <div className="flex gap-2">
          <input
            placeholder="#"
            type="number"
            inputMode="numeric"
            value={numero}
            onChange={(e) => setNumero(e.target.value.replace(/\D/g, ''))}
            className="w-14 bg-app-bg border border-app-line rounded-lg px-2 py-2 text-sm text-center"
          />
          <input
            placeholder="Nombre del comprador"
            value={nombre}
            onChange={(e) => setNombre(e.target.value)}
            className="flex-1 bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm"
          />
          <input
            placeholder="Pref."
            value={prefijo}
            onChange={(e) => setPrefijo(e.target.value.replace(/[^A-Za-z]/g, '').slice(0, 3))}
            className="w-16 bg-app-bg border border-app-line rounded-lg px-2 py-2 text-sm text-center uppercase font-mono"
          />
        </div>
        <p className="text-xs text-app-dim">Prefijo opcional: 1-3 letras para diferenciar del resto del equipo (ej. F, Fr, Fra). Si lo dejas vacío, se usa la primera letra del nombre.</p>
        <div className="flex gap-2">
          <input
            placeholder="Correo para entrar a la app"
            type="email"
            inputMode="email"
            value={correo}
            onChange={(e) => setCorreo(e.target.value)}
            className="flex-1 bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm"
          />
          <button onClick={agregar} disabled={guardando} className="px-4 rounded-lg bg-app-gold text-app-bg text-sm font-semibold disabled:opacity-50">
            {guardando ? '…' : 'Alta'}
          </button>
        </div>
        <p className="text-xs text-app-dim">
          El correo debe ser el mismo que registres en Supabase → Authentication. Sin correo, el
          comprador aparece en las listas pero no puede entrar a la app.
        </p>
      </div>

      {/* Lista de usuarios */}
      <div className="space-y-1.5">
        {usuariosOrdenados.map((u) => {
          const pref = prefijoEfectivo(u);
          const enConflicto = conteoPrefijos[pref] > 1;
          const uEsAdmin = esAdmin(u);
          const cambiandoPin = pinCambioUsuarioId === u.id;
          return (
            <div key={u.id} className={`bg-app-panel border rounded-lg px-3 py-2 ${enConflicto ? 'border-red-500/40' : 'border-app-line'}`}>
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-2.5 min-w-0 flex-1">
                  <span className="w-7 h-7 rounded-full bg-app-bg text-app-gold flex items-center justify-center font-bold text-xs border border-app-line shrink-0">
                    {(u.nombre || '?').charAt(0).toUpperCase()}
                  </span>
                  <div className="min-w-0 flex-1">
                    <p className="text-sm">
                      {u.numero && <span className="text-app-dim2 mr-1.5">#{u.numero}</span>}
                      {u.nombre}
                      {u.id === usuarioActivo?.id && <span className="text-xs text-app-gold ml-2">(tú)</span>}
                      {uEsAdmin && (
                        <span className="text-xs bg-app-goldbg border border-app-gold/40 text-app-gold rounded px-1.5 py-0.5 ml-2 font-semibold">
                          👑 Admin
                        </span>
                      )}
                    </p>
                    <div className="flex items-center gap-2 mt-0.5 flex-wrap">
                      {editandoId === u.id ? (
                        <div className="flex items-center gap-1.5">
                          <input
                            autoFocus
                            value={editPrefijo}
                            onChange={(e) => setEditPrefijo(e.target.value.replace(/[^A-Za-z]/g, '').slice(0, 3))}
                            onKeyDown={(e) => { if (e.key === 'Enter') guardarPrefijo(u); if (e.key === 'Escape') setEditandoId(null); }}
                            className="w-14 bg-app-bg border border-app-gold rounded px-1.5 py-0.5 text-xs text-center uppercase font-mono"
                            placeholder={(u.nombre || '?').charAt(0).toUpperCase()}
                          />
                          <button onClick={() => guardarPrefijo(u)} className="text-xs text-app-green font-semibold">✓</button>
                          <button onClick={() => setEditandoId(null)} className="text-xs text-app-dim">✕</button>
                        </div>
                      ) : (
                        <button
                          onClick={() => abrirEdicionPrefijo(u)}
                          className={`text-xs rounded px-1.5 py-0.5 font-mono border ${
                            enConflicto
                              ? 'bg-app-redbg border-red-500/50 text-app-red2'
                              : 'bg-app-goldbg border-app-gold/40 text-app-gold'
                          }`}
                          title="Editar prefijo"
                        >
                          {pref}
                        </button>
                      )}
                      {enConflicto && editandoId !== u.id && (
                        <span className="text-xs text-app-red2 flex items-center gap-1">
                          <AlertCircle size={11} /> Prefijo repetido
                        </span>
                      )}
                    </div>
                    {/* Correo de acceso */}
                    <div className="mt-1">
                      {editandoCorreoId === u.id ? (
                        <div className="flex items-center gap-1.5">
                          <input
                            autoFocus
                            type="email"
                            value={editCorreo}
                            onChange={(e) => setEditCorreo(e.target.value)}
                            onKeyDown={(e) => { if (e.key === 'Enter') guardarCorreo(u); if (e.key === 'Escape') setEditandoCorreoId(null); }}
                            placeholder="correo@carrion.hn"
                            className="flex-1 min-w-0 bg-app-bg border border-app-gold rounded px-2 py-0.5 text-xs"
                          />
                          <button onClick={() => guardarCorreo(u)} className="text-xs text-app-green font-semibold">✓</button>
                          <button onClick={() => setEditandoCorreoId(null)} className="text-xs text-app-dim">✕</button>
                        </div>
                      ) : (
                        <button
                          onClick={() => soyAdmin && abrirEdicionCorreo(u)}
                          className="text-xs text-left"
                          title={soyAdmin ? 'Editar correo' : ''}
                        >
                          {u.email
                            ? <span className="text-app-dim2">✉️ {u.email}</span>
                            : <span className="text-app-gold">✉️ Sin correo — no puede entrar</span>}
                        </button>
                      )}
                    </div>

                    {/* Rol y países — solo el admin los edita; a sí mismo no se los quita */}
                    {soyAdmin && !uEsAdmin && (
                      <div className="mt-2 pt-2 border-t border-app-line space-y-1.5">
                        <div className="flex items-center gap-1.5">
                          <span className="text-xs text-app-dim3 w-14 shrink-0">Rol</span>
                          <select
                            value={u.rol === 'supervisor' ? 'supervisor' : 'comprador'}
                            onChange={(e) => cambiarRol(u, e.target.value)}
                            className="bg-app-bg border border-app-line rounded px-2 py-1 text-xs"
                          >
                            <option value="comprador">Comprador</option>
                            <option value="supervisor">Supervisor</option>
                          </select>
                        </div>
                        {(u.rol !== 'supervisor') && (
                          <div className="flex items-start gap-1.5">
                            <span className="text-xs text-app-dim3 w-14 shrink-0 mt-1">Países</span>
                            <div className="flex flex-wrap gap-1">
                              {ORIGENES.map((o) => {
                                const activos = Array.isArray(u.paises) ? u.paises : ORIGENES.map((x) => x.id);
                                const activo = activos.includes(o.id);
                                return (
                                  <button
                                    key={o.id}
                                    onClick={() => alternarPais(u, o.id)}
                                    className={`text-xs rounded-full px-2 py-0.5 border ${
                                      activo
                                        ? 'bg-app-goldbg border-app-gold/40 text-app-gold'
                                        : 'bg-app-bg border-app-line text-app-dim3'
                                    }`}
                                  >
                                    {o.emoji} {o.label}
                                  </button>
                                );
                              })}
                            </div>
                          </div>
                        )}
                      </div>
                    )}
                    {!soyAdmin && !uEsAdmin && (
                      <p className="text-xs text-app-dim3 mt-1">
                        {u.rol === 'supervisor' ? '🔎 Supervisor' : '🛒 Comprador'}
                        {u.rol !== 'supervisor' && Array.isArray(u.paises) && u.paises.length < ORIGENES.length && (
                          <> · {u.paises.map((id) => ORIGENES.find((o) => o.id === id)?.emoji).join(' ')}</>
                        )}
                      </p>
                    )}
                  </div>
                </div>
                <div className="flex items-center gap-1.5 shrink-0">
                  {/* El PIN se retiró: la contraseña ahora la maneja Supabase Auth.
                      Cada comprador la restablece desde "¿Olvidaste tu contraseña?". */}
                  {u.id !== usuarioActivo?.id && (
                    <BotonBorrar onConfirm={() => eliminar(u)} size={14} />
                  )}
                </div>
              </div>

              {/* Panel de cambio de PIN (visible solo cuando se abre para este usuario) */}
              {cambiandoPin && (
                <div className="mt-2 pt-2 border-t border-app-line space-y-2">
                  <p className="text-xs text-app-sky font-medium">
                    {u.pinHash ? 'Cambiar PIN' : 'Asignar PIN'} para {u.nombre}
                  </p>
                  <div className="flex gap-2">
                    <input
                      autoFocus
                      placeholder="Nuevo PIN (4-6 dígitos)"
                      type="password"
                      inputMode="numeric"
                      maxLength={6}
                      value={pinCambioValor}
                      onChange={(e) => setPinCambioValor(e.target.value.replace(/\D/g, ''))}
                      className="flex-1 bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm text-center"
                    />
                    <input
                      placeholder="Confirmar"
                      type="password"
                      inputMode="numeric"
                      maxLength={6}
                      value={pinCambioConfirm}
                      onChange={(e) => setPinCambioConfirm(e.target.value.replace(/\D/g, ''))}
                      className="flex-1 bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm text-center"
                    />
                  </div>
                  <div className="flex gap-2">
                    <button
                      onClick={() => setPinCambioUsuarioId(null)}
                      className="flex-1 py-1.5 rounded-lg border border-app-line text-xs text-app-dim2"
                    >
                      Cancelar
                    </button>
                    {u.pinHash && (
                      <button
                        onClick={() => guardarNuevoPin(u)}
                        className="flex-1 py-1.5 rounded-lg border border-app-line text-xs text-app-red2"
                        title="Dejar en blanco = quitar PIN"
                      >
                        Quitar PIN
                      </button>
                    )}
                    <button
                      onClick={() => guardarNuevoPin(u)}
                      className="flex-1 py-1.5 rounded-lg bg-app-gold text-app-bg text-xs font-semibold"
                    >
                      Guardar
                    </button>
                  </div>
                </div>
              )}
            </div>
          );
        })}
      </div>
      {!soyAdmin && (
        <p className="text-xs text-app-dim2 bg-app-panel border border-app-line rounded-lg px-3 py-2">
          🔐 Solo el administrador puede asignar o cambiar los PIN. Si necesitas cambiar el tuyo, pídeselo al administrador.
        </p>
      )}
      <p className="text-xs text-app-dim">Cada quien entra con su nombre y PIN opcional. Los pedidos quedan firmados con el nombre de quien los crea. El prefijo aparece en el correlativo del código del producto (ej. Yamil = Y26600001).</p>
    </div>
  );
}

function FabricasList({ fabricas = [], setFabricas }) {
  const [newFab, setNewFab] = useState('');
  const add = () => {
    const f = newFab.trim();
    if (!f || fabricas.includes(f)) return;
    setFabricas([...fabricas, f].sort((a, b) => a.localeCompare(b)));
    setNewFab('');
  };
  const remove = (f) => setFabricas(fabricas.filter((x) => x !== f));
  return (
    <div>
      <div className="flex gap-2 mb-2">
        <input
          value={newFab}
          onChange={(e) => setNewFab(e.target.value)}
          onKeyDown={(e) => { if (e.key === 'Enter') add(); }}
          placeholder="Nombre de la fábrica o proveedor…"
          className="flex-1 bg-app-panel border border-app-line rounded-lg px-3 py-2 text-sm"
        />
        <button onClick={add} className="px-3 rounded-lg bg-app-gold text-app-bg text-sm font-semibold">Agregar</button>
      </div>
      <div className="flex flex-wrap gap-1.5">
        {fabricas.map((f) => (
          <span key={f} className="bg-app-panel border border-app-line text-sm rounded-full pl-3 pr-1.5 py-1.5 flex items-center gap-1.5">
            {f}
            <BotonBorrar onConfirm={() => remove(f)} size={13} icono="x" />
          </span>
        ))}
        {fabricas.length === 0 && <p className="text-xs text-app-dim">Sin fábricas aún. Agrégalas aquí para buscarlas al crear productos.</p>}
      </div>
    </div>
  );
}

// ---------- Módulo de Administración (solo admin): Usuarios + Costo puesto en bodega ----------
// ------------------------------------------------------------
// Historial de auditoría — quién hizo qué y cuándo.
// Solo lectura; carga los últimos registros desde Supabase.
// ------------------------------------------------------------
const ETIQUETA_ACCION = {
  crear: { texto: 'Creó', color: 'text-green-400' },
  editar: { texto: 'Editó', color: 'text-app-gold' },
  borrar: { texto: 'Borró', color: 'text-red-400' },
};
const ETIQUETA_TABLA = {
  productos: 'Producto', pedidos: 'Pedido', proveedores: 'Proveedor',
  compradores: 'Comprador', factores_costo: 'Factores de costo', sesion: 'Sesión',
};

function PanelAuditoria({ soyAdmin }) {
  const [filas, setFilas] = useState([]);
  const [cargando, setCargando] = useState(true);
  const [error, setError] = useState('');
  const [filtroTabla, setFiltroTabla] = useState('');

  const cargar = useCallback(async () => {
    if (!supabase) { setCargando(false); return; }
    setCargando(true);
    setError('');
    try {
      let q = supabase.from('auditoria').select('*').order('fecha', { ascending: false }).limit(150);
      if (filtroTabla) q = q.eq('tabla', filtroTabla);
      const { data, error: err } = await q;
      if (err) throw err;
      setFilas(data || []);
    } catch (e) {
      setError('No se pudo cargar el historial.');
    } finally {
      setCargando(false);
    }
  }, [filtroTabla]);

  useEffect(() => { cargar(); }, [cargar]);

  if (!soyAdmin) return null;

  return (
    <section className="pt-2">
      <div className="flex items-center justify-between mb-2">
        <h2 className="text-sm font-semibold text-app-white flex items-center gap-1.5">
          📜 Historial de actividad
        </h2>
        <button onClick={cargar} className="text-xs text-app-dim2 underline">Actualizar</button>
      </div>

      <select
        value={filtroTabla}
        onChange={(e) => setFiltroTabla(e.target.value)}
        className="w-full bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm mb-2"
      >
        <option value="">Todo</option>
        {Object.entries(ETIQUETA_TABLA).map(([k, v]) => (
          <option key={k} value={k}>{v}</option>
        ))}
      </select>

      {cargando ? (
        <p className="text-xs text-app-dim py-3">Cargando…</p>
      ) : error ? (
        <p className="text-xs text-app-red py-3">{error}</p>
      ) : filas.length === 0 ? (
        <p className="text-xs text-app-dim py-3">Sin actividad registrada todavía.</p>
      ) : (
        <div className="space-y-1.5 max-h-96 overflow-y-auto">
          {filas.map((f) => {
            const acc = ETIQUETA_ACCION[f.accion] || { texto: f.accion, color: 'text-app-dim2' };
            return (
              <div key={f.id} className="bg-app-panel border border-app-line rounded-lg px-3 py-2 text-xs">
                <div className="flex items-center justify-between gap-2">
                  <span>
                    <span className={`font-semibold ${acc.color}`}>{acc.texto}</span>
                    {' '}
                    <span className="text-app-dim2">{ETIQUETA_TABLA[f.tabla] || f.tabla}</span>
                  </span>
                  <span className="text-app-dim3 shrink-0">
                    {new Date(f.fecha).toLocaleString('es-HN', { day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit' })}
                  </span>
                </div>
                {f.resumen && <p className="text-app-white mt-0.5 truncate">{f.resumen}</p>}
                <p className="text-app-dim3 mt-0.5">{f.usuario_nombre || f.usuario_email || 'Desconocido'}</p>
              </div>
            );
          })}
        </div>
      )}
    </section>
  );
}

function Administracion({ usuarios = [], setUsuarios, usuarioActivo, onLogout, factores, setFactores, tasaCambio, setTasaCambio, empresa, setEmpresa, embarcadores = [], setEmbarcadores }) {
  const [errorLogo, setErrorLogo] = useState('');
  const [showEmbForm, setShowEmbForm] = useState(false);
  const [embForm, setEmbForm] = useState({ nombre: '', contacto: '', telefono: '', email: '', direccion: '' });
  const [editEmbId, setEditEmbId] = useState(null);

  const addEmbarcador = () => {
    if (!(embForm.nombre || '').trim()) return;
    if (editEmbId) {
      setEmbarcadores((embarcadores || []).map((e) => e.id === editEmbId ? { ...e, ...embForm } : e));
    } else {
      setEmbarcadores([...(embarcadores || []), { ...embForm, id: uid() }]);
    }
    setEmbForm({ nombre: '', contacto: '', telefono: '', email: '', direccion: '' });
    setEditEmbId(null);
    setShowEmbForm(false);
  };
  const editarEmbarcador = (e) => {
    setEditEmbId(e.id);
    setEmbForm({ nombre: e.nombre, contacto: e.contacto || '', telefono: e.telefono || '', email: e.email || '', direccion: e.direccion || '' });
    setShowEmbForm(true);
  };
  const quitarEmbarcador = (id) => setEmbarcadores((embarcadores || []).filter((e) => e.id !== id));

  // Carga el logo reduciéndolo antes de guardar (evita ocupar el almacenamiento)
  const cargarLogo = (e) => {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file) return;
    setErrorLogo('');
    const reader = new FileReader();
    reader.onload = () => {
      const img = new Image();
      img.onload = () => {
        const MAX = 400; // ancho máximo en píxeles
        const escala = Math.min(1, MAX / img.width);
        const w = Math.round(img.width * escala);
        const h = Math.round(img.height * escala);
        const canvas = document.createElement('canvas');
        canvas.width = w; canvas.height = h;
        const ctx = canvas.getContext('2d');
        ctx.drawImage(img, 0, 0, w, h);
        
        // Eliminar fondo blanco reemplazando pixeles blancos con transparencia
        const imageData = ctx.getImageData(0, 0, w, h);
        const data = imageData.data;
        const umbral = 250; // pixeles >= este valor se consideran "blancos"
        for (let i = 0; i < data.length; i += 4) {
          const r = data[i], g = data[i+1], b = data[i+2];
          // Si el pixel es casi blanco (R, G, B > umbral)
          if (r > umbral && g > umbral && b > umbral) {
            data[i+3] = 0; // alfa = 0 (transparente)
          }
        }
        ctx.putImageData(imageData, 0, 0);
        
        const dataUrl = canvas.toDataURL('image/png');
        if (dataUrl.length > 400000) {
          setErrorLogo('La imagen es muy pesada. Usa una versión más pequeña o en PNG.');
          return;
        }
        setEmpresa({ ...(empresa || {}), logo: dataUrl });
      };
      img.onerror = () => setErrorLogo('No pude leer esa imagen.');
      img.src = reader.result;
    };
    reader.onerror = () => setErrorLogo('No pude leer el archivo.');
    reader.readAsDataURL(file);
  };

  const quitarLogo = () => {
    const next = { ...(empresa || {}) };
    delete next.logo;
    setEmpresa(next);
  };

  return (
    <div className="space-y-6">
      <h2 className="text-lg font-semibold">🔐 Administración</h2>

      <section>
        <h3 className="text-sm font-semibold mb-2 flex items-center gap-2">🏢 Logo de la empresa</h3>
        <p className="text-xs text-app-dim2 mb-2">
          Aparece en la hoja de pedido que se imprime. Se reduce automáticamente para no ocupar espacio.
        </p>
        <div className="bg-app-panel border border-app-line rounded-xl p-3 space-y-2.5">
          {empresa?.logo ? (
            <div className="flex items-center gap-3">
              <div className="bg-white rounded-lg p-2 border border-app-line shrink-0">
                <img src={empresa.logo} alt="Logo" className="h-14 w-auto object-contain" />
              </div>
              <div className="flex-1 min-w-0">
                <p className="text-xs text-app-green">✓ Logo cargado</p>
                <p className="text-xs text-app-dim mt-0.5">Así se verá en el documento.</p>
              </div>
              <BotonBorrar onConfirm={quitarLogo} size={14} />
            </div>
          ) : (
            <label className="flex flex-col items-center justify-center gap-1.5 py-6 rounded-xl border border-dashed border-app-line3 text-app-dim2 cursor-pointer active:bg-app-active">
              <Upload size={20} className="text-app-gold" />
              <span className="text-sm">Subir logo</span>
              <span className="text-xs text-app-dim">PNG o JPG</span>
              <input type="file" accept="image/*" onChange={cargarLogo} className="hidden" />
            </label>
          )}
          {empresa?.logo && (
            <label className="w-full py-2 rounded-lg border border-app-line text-xs text-app-sky flex items-center justify-center gap-1.5 cursor-pointer active:bg-app-active">
              <Upload size={13} /> Cambiar logo
              <input type="file" accept="image/*" onChange={cargarLogo} className="hidden" />
            </label>
          )}
          {errorLogo && (
            <div className="flex items-start gap-2 bg-app-redbg text-app-red2 text-xs rounded-lg px-3 py-2">
              <AlertCircle size={14} className="shrink-0 mt-0.5" /> {errorLogo}
            </div>
          )}
        </div>
      </section>

      <section>
        <h3 className="text-sm font-semibold mb-2 flex items-center gap-2">🚢 Compañías de embarque</h3>
        <p className="text-xs text-app-dim2 mb-2">
          Pre-registra tus embarcadores. Al crear un pedido de 🇺🇸 USA o 🇵🇦 Panamá podrás elegir uno, y aparecerá en la hoja de pedido.
        </p>
        <div className="space-y-1.5 mb-2">
          {(embarcadores || []).map((e) => (
            <div key={e.id} className="bg-app-panel border border-app-line rounded-lg px-3 py-2 flex items-start justify-between gap-2">
              <div className="min-w-0 flex-1">
                <p className="text-sm font-medium truncate">{e.nombre}</p>
                <p className="text-xs text-app-dim2">
                  {[e.contacto, e.telefono, e.email].filter(Boolean).join(' · ') || 'Sin datos de contacto'}
                </p>
                {e.direccion && <p className="text-xs text-app-dim mt-0.5">{e.direccion}</p>}
              </div>
              <div className="flex items-center gap-1 shrink-0">
                <button onClick={() => editarEmbarcador(e)} className="text-app-dim2 active:text-app-sky p-1" title="Editar">
                  <Pencil size={13} />
                </button>
                <BotonBorrar onConfirm={() => quitarEmbarcador(e.id)} size={13} />
              </div>
            </div>
          ))}
          {(embarcadores || []).length === 0 && !showEmbForm && (
            <p className="text-center text-xs text-app-dim py-3">Aún no tienes compañías de embarque.</p>
          )}
        </div>

        {showEmbForm ? (
          <div className="bg-app-panel border border-app-line rounded-xl p-3 space-y-2">
            <div className="flex justify-between items-center">
              <span className="text-sm font-medium">{editEmbId ? 'Editar embarcador' : 'Nuevo embarcador'}</span>
              <button onClick={() => { setShowEmbForm(false); setEditEmbId(null); setEmbForm({ nombre: '', contacto: '', telefono: '', email: '', direccion: '' }); }}>
                <X size={16} className="text-app-dim" />
              </button>
            </div>
            <input value={embForm.nombre} onChange={(e) => setEmbForm({ ...embForm, nombre: e.target.value })}
              placeholder="Nombre de la compañía *" className="w-full bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm" />
            <input value={embForm.contacto} onChange={(e) => setEmbForm({ ...embForm, contacto: e.target.value })}
              placeholder="Contacto (persona)" className="w-full bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm" />
            <input value={embForm.telefono} onChange={(e) => setEmbForm({ ...embForm, telefono: e.target.value })}
              placeholder="Teléfono" className="w-full bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm" />
            <input value={embForm.email} onChange={(e) => setEmbForm({ ...embForm, email: e.target.value })}
              placeholder="Correo" className="w-full bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm" />
            <textarea value={embForm.direccion} onChange={(e) => setEmbForm({ ...embForm, direccion: e.target.value })}
              placeholder="Dirección" rows={2} className="w-full bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm resize-none" />
            <button onClick={addEmbarcador} disabled={!(embForm.nombre || '').trim()}
              className="w-full py-2.5 rounded-lg bg-app-gold text-app-bg text-sm font-semibold disabled:opacity-50">
              {editEmbId ? 'Guardar cambios' : 'Agregar embarcador'}
            </button>
          </div>
        ) : (
          <button onClick={() => setShowEmbForm(true)}
            className="w-full py-3 rounded-xl border border-dashed border-app-line3 text-sm text-app-dim2 flex items-center justify-center gap-2 active:bg-app-panel">
            <Plus size={16} /> Agregar compañía de embarque
          </button>
        )}
      </section>

      <section>
        <UsuariosConfig
          usuarios={usuarios || []}
          setUsuarios={setUsuarios}
          usuarioActivo={usuarioActivo}
          onLogout={onLogout}
        />
      </section>

      <section>
        <h3 className="text-sm font-semibold mb-2 flex items-center gap-2">💵 Costo puesto en bodega</h3>
        <p className="text-xs text-app-dim2 mb-2">
          El factor por país convierte el costo USD del proveedor a Lempiras puestos en tu bodega. Incluye flete, seguro, impuestos y tasa USD→HNL.
        </p>
        <div className="bg-app-panel border border-app-line rounded-xl p-3 space-y-2.5 shadow-app">
          <div>
            <label className="text-xs text-app-dim2 uppercase tracking-wide mb-1 block">Tasa RMB → USD (para China)</label>
            <input
              type="number" step="0.01" min={0}
              value={tasaCambio?.rmbUsd ?? 7.25}
              onChange={(e) => setTasaCambio({ ...(tasaCambio || {}), rmbUsd: parseFloat(e.target.value) || 0 })}
              className="w-full bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm"
            />
            <p className="text-xs text-app-dim mt-1">Ej: 7.25 → 1 USD equivale a 7.25 RMB</p>
          </div>

          {ORIGENES.map((o) => (
            <div key={o.id}>
              <label className="text-xs text-app-dim2 uppercase tracking-wide mb-1 flex items-center gap-1.5">
                <span>{o.emoji}</span> Factor {o.label}
              </label>
              <div className="flex items-center gap-2">
                <span className="text-xs text-app-dim2">USD ×</span>
                <input
                  type="number" step="0.1" min={0}
                  value={factores?.[o.id] ?? ''}
                  onChange={(e) => setFactores({ ...(factores || {}), [o.id]: parseFloat(e.target.value) || 0 })}
                  className="flex-1 bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm"
                />
                <span className="text-xs text-app-dim2">= HNL</span>
              </div>
            </div>
          ))}

          {/* % Niki — solo para China: recargo en USD tras la conversión */}
          <div className="border-t border-app-line pt-3 mt-2">
            <label className="text-xs text-app-dim2 uppercase tracking-wide mb-1 flex items-center gap-1.5">
              <span>🇨🇳</span> Niki (% recargo en USD, solo China)
            </label>
            <div className="flex items-center gap-2">
              <span className="text-xs text-app-dim2">USD ×</span>
              <input
                type="number" step="0.1" min={0}
                value={factores?.nikiPct ?? ''}
                onChange={(e) => setFactores({ ...(factores || {}), nikiPct: parseFloat(e.target.value) || 0 })}
                className="flex-1 bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm"
                placeholder="Ej: 15"
              />
              <span className="text-xs text-app-dim2">%</span>
            </div>
            <p className="text-xs text-app-dim mt-1">Se suma al USD del proveedor antes de multiplicar por el factor de China.</p>
          </div>
        </div>
      </section>

      <PanelAuditoria soyAdmin={!!usuarioActivo?.esAdmin} />
    </div>
  );
}

// ---------- Marcas con sus proveedores (relación aparte, no toca el modelo de producto) ----------
function MarcasConfig({ marcas = [], setMarcas, marcasProveedores, setMarcasProveedores, suppliers = [], origen }) {
  const [nueva, setNueva] = useState('');
  const [editandoMarca, setEditandoMarca] = useState(null);
  const [importPreview, setImportPreview] = useState(null);
  const [importError, setImportError] = useState('');

  const norm = (t) => String(t ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toLowerCase();

  const provsDeMarca = (m) => (marcasProveedores[m] || []);
  const nombreProv = (id) => suppliers.find((s) => s.id === id)?.nombre || '(eliminado)';
  const emojiProv = (id) => ORIGENES.find((o) => o.id === suppliers.find((s) => s.id === id)?.origen)?.emoji || '';

  const addMarca = () => {
    const m = nueva.trim();
    if (!m || marcas.some((x) => norm(x) === norm(m))) { setNueva(''); return; }
    setMarcas([...marcas, m].sort((a, b) => a.localeCompare(b)));
    setNueva('');
  };

  const removeMarca = (m) => {
    setMarcas(marcas.filter((x) => x !== m));
    const next = { ...marcasProveedores };
    delete next[m];
    setMarcasProveedores(next);
  };

  const toggleProveedor = (marca, supplierId) => {
    const actuales = provsDeMarca(marca);
    const next = actuales.includes(supplierId)
      ? actuales.filter((id) => id !== supplierId)
      : [...actuales, supplierId];
    setMarcasProveedores({ ...marcasProveedores, [marca]: next });
  };

  // ----- Plantilla -----
  const FILAS_PLANTILLA = [
    { Marca: 'Polo Club', Proveedores: 'Guangzhou Trading Co, Yiwu Textiles' },
    { Marca: 'Collezione', Proveedores: 'Guangzhou Trading Co' },
    { Marca: 'MOD', Proveedores: '' },
  ];

  // ----- Importación -----
  const leerArchivo = (e) => {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file) return;
    setImportError(''); setImportPreview(null);
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: 'array' });
        const filas = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' });
        if (filas.length === 0) return setImportError('El archivo no tiene filas.');

        const claves = [...new Set(filas.flatMap((f) => Object.keys(f)))];
        const buscarCol = (...alias) => claves.find((k) => alias.includes(norm(k)));
        const kMarca = buscarCol('marca', 'marcas', 'brand');
        const kProv = buscarCol('proveedores', 'proveedor', 'supplier', 'suppliers');
        if (!kMarca) return setImportError('No encontré la columna "Marca". Descarga la plantilla para ver el formato.');

        // Acumula: acepta proveedores por coma y también varias filas de la misma marca
        const acumulado = {};   // marcaTexto -> Set(supplierId)
        const noEncontrados = new Set();
        filas.forEach((f) => {
          const marca = String(f[kMarca] ?? '').trim();
          if (!marca) return;
          if (!acumulado[marca]) acumulado[marca] = new Set();
          if (!kProv) return;
          String(f[kProv] ?? '').split(/[,;]/).forEach((txt) => {
            const nombre = txt.trim();
            if (!nombre) return;
            const prov = suppliers.find((s) => norm(s.nombre) === norm(nombre));
            if (prov) acumulado[marca].add(prov.id);
            else noEncontrados.add(nombre);
          });
        });

        const marcasNuevas = [];
        const relaciones = [];
        Object.entries(acumulado).forEach(([marca, setIds]) => {
          const existente = marcas.find((x) => norm(x) === norm(marca));
          const nombreFinal = existente || marca;
          if (!existente) marcasNuevas.push(nombreFinal);
          const yaTiene = new Set(marcasProveedores[nombreFinal] || []);
          const agregar = [...setIds].filter((id) => !yaTiene.has(id));
          if (agregar.length > 0) relaciones.push({ marca: nombreFinal, agregar });
        });

        if (marcasNuevas.length === 0 && relaciones.length === 0) {
          return setImportError('Todo lo del archivo ya está registrado.');
        }
        setImportPreview({ marcasNuevas, relaciones, noEncontrados: [...noEncontrados] });
      } catch (err) {
        setImportError('No pude leer el archivo. Debe ser .xlsx o .csv');
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const confirmarImportacion = () => {
    const { marcasNuevas, relaciones } = importPreview;
    if (marcasNuevas.length > 0) {
      setMarcas([...marcas, ...marcasNuevas].sort((a, b) => a.localeCompare(b)));
    }
    if (relaciones.length > 0) {
      const next = { ...marcasProveedores };
      relaciones.forEach(({ marca, agregar }) => {
        next[marca] = [...new Set([...(next[marca] || []), ...agregar])];
      });
      setMarcasProveedores(next);
    }
    setImportPreview(null);
  };

  return (
    <div className="space-y-3">
      {/* Carga masiva */}
      <div className="bg-app-panel border border-app-line rounded-xl p-3 space-y-2">
        <p className="text-xs uppercase tracking-wide text-app-dim2">Cargar varias de una vez</p>
        <div className="flex gap-2">
          <label className="flex-1 py-2.5 rounded-lg border border-app-line text-xs text-app-sky flex items-center justify-center gap-1.5 active:bg-app-active cursor-pointer">
            <Upload size={14} /> Importar Excel
            <input type="file" accept=".xlsx,.xls,.csv" onChange={leerArchivo} className="hidden" />
          </label>
        </div>
        <BotonesPlantilla
          filas={FILAS_PLANTILLA}
          nombreHoja="Plantilla"
          nombreArchivo="plantilla-marcas.xlsx"
          cols={[{ wch: 24 }, { wch: 52 }]}
        />
        <p className="text-xs text-app-dim">
          Columnas: <span className="font-mono text-app-dim2">Marca</span> (obligatoria) y <span className="font-mono text-app-dim2">Proveedores</span>.
          Puedes poner varios proveedores separados por coma, o repetir la marca en varias filas.
        </p>

        {importError && (
          <div className="flex items-start gap-2 bg-app-redbg text-app-red2 text-xs rounded-lg px-3 py-2">
            <AlertCircle size={14} className="shrink-0 mt-0.5" /> {importError}
          </div>
        )}

        {importPreview && (
          <div className="bg-app-blue border border-app-line2 rounded-lg p-3 space-y-2">
            {importPreview.marcasNuevas.length > 0 && (
              <p className="text-xs text-app-sky font-medium">
                Marcas nuevas: {importPreview.marcasNuevas.length} — {importPreview.marcasNuevas.slice(0, 6).join(', ')}
                {importPreview.marcasNuevas.length > 6 ? '…' : ''}
              </p>
            )}
            {importPreview.relaciones.length > 0 && (
              <p className="text-xs text-app-sky">
                Se enlazarán proveedores a {importPreview.relaciones.length} marca{importPreview.relaciones.length !== 1 ? 's' : ''}.
              </p>
            )}
            {importPreview.noEncontrados.length > 0 && (
              <p className="text-xs text-app-gold">
                ⚠ Estos proveedores no existen y se omiten: {importPreview.noEncontrados.join(', ')}
              </p>
            )}
            <div className="flex gap-2">
              <button onClick={() => setImportPreview(null)} className="flex-1 py-2 rounded-lg border border-app-line text-xs text-app-dim2">Cancelar</button>
              <button onClick={confirmarImportacion} className="flex-1 py-2 rounded-lg bg-app-sky text-app-bg text-xs font-semibold">Importar</button>
            </div>
          </div>
        )}
      </div>

      {/* Alta manual */}
      <div className="flex gap-2">
        <input
          value={nueva}
          onChange={(e) => setNueva(e.target.value)}
          onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); addMarca(); } }}
          placeholder="Nueva marca"
          className="flex-1 bg-app-panel border border-app-line rounded-lg px-3 py-2 text-sm"
        />
        <button onClick={addMarca} className="px-4 rounded-lg bg-app-gold text-app-bg text-sm font-semibold">Agregar</button>
      </div>

      {/* Lista de marcas con sus proveedores */}
      <div className="space-y-1.5">
        {marcas.map((m) => {
          const ids = provsDeMarca(m);
          const abierto = editandoMarca === m;
          return (
            <div key={m} className="bg-app-panel border border-app-line rounded-lg px-3 py-2">
              <div className="flex items-center justify-between gap-2">
                <button onClick={() => setEditandoMarca(abierto ? null : m)} className="flex items-center gap-2 min-w-0 flex-1 text-left">
                  <ChevronDown size={13} className={`text-app-dim shrink-0 transition-transform ${abierto ? '' : '-rotate-90'}`} />
                  <span className="text-sm truncate">{m}</span>
                  <span className="text-xs text-app-dim2 shrink-0">
                    {ids.length > 0 ? `${ids.length} proveedor${ids.length !== 1 ? 'es' : ''}` : 'sin proveedor'}
                  </span>
                </button>
                <BotonBorrar onConfirm={() => removeMarca(m)} size={13} icono="x" />
              </div>

              {!abierto && ids.length > 0 && (
                <div className="flex flex-wrap gap-1 mt-1 pl-5">
                  {ids.map((id) => (
                    <span key={id} className="text-xs bg-app-bg border border-app-line rounded px-1.5 py-0.5 text-app-dim2">
                      {emojiProv(id)} {nombreProv(id)}
                    </span>
                  ))}
                </div>
              )}

              {abierto && (
                <div className="mt-2 pt-2 border-t border-app-line space-y-1">
                  <p className="text-xs text-app-dim2 mb-1">Marca los proveedores que manejan esta marca:</p>
                  {suppliers.length === 0 && (
                    <p className="text-xs text-app-dim">No hay proveedores registrados todavía.</p>
                  )}
                  {suppliers.map((s) => {
                    const activo = ids.includes(s.id);
                    return (
                      <button
                        key={s.id}
                        onClick={() => toggleProveedor(m, s.id)}
                        className={`w-full flex items-center gap-2 text-left text-xs rounded-lg px-2.5 py-2 border ${
                          activo ? 'bg-app-goldbg border-app-gold text-app-gold' : 'bg-app-bg border-app-line text-app-light'
                        }`}
                      >
                        <span className="shrink-0">{activo ? '✓' : '○'}</span>
                        <span className="shrink-0">{ORIGENES.find((o) => o.id === s.origen)?.emoji || '—'}</span>
                        <span className="truncate">{s.nombre}</span>
                      </button>
                    );
                  })}
                </div>
              )}
            </div>
          );
        })}
        {marcas.length === 0 && <p className="text-center text-xs text-app-dim py-4">Sin marcas todavía.</p>}
      </div>
    </div>
  );
}

function Config({ departamentos = [], setDepartamentos, tipos = [], setTipos, marcas = [], setMarcas, marcasProveedores, setMarcasProveedores, suppliers = [], origen, ciudades = [], setCiudades, fabricas = [], setFabricas, onBack, onPresupuestos }) {
  const [newDept, setNewDept] = useState('');
  const [newCiudad, setNewCiudad] = useState('');
  const [newTipoNombre, setNewTipoNombre] = useState('');
  const [newMedida, setNewMedida] = useState('simple');
  const [newTipoTallas, setNewTipoTallas] = useState('');
  const [newCinturas, setNewCinturas] = useState('');
  const [newLargos, setNewLargos] = useState('');
  const [newSubtipos, setNewSubtipos] = useState('');
  const [editingTipoId, setEditingTipoId] = useState(null);
  const [editNombre, setEditNombre] = useState('');
  const [queryTipos, setQueryTipos] = useState(''); // búsqueda de tipos
  const [editTallas, setEditTallas] = useState('');
  const [editCinturas, setEditCinturas] = useState('');
  const [editLargos, setEditLargos] = useState('');
  const [editingSubtiposId, setEditingSubtiposId] = useState(null);
  const [editSubtipos, setEditSubtipos] = useState('');

  const addDept = () => {
    const d = newDept.trim();
    if (!d || departamentos.includes(d)) return;
    setDepartamentos([...departamentos, d]);
    setNewDept('');
  };
  const removeDept = (d) => setDepartamentos(departamentos.filter((x) => x !== d));


  const addCiudad = () => {
    const c = newCiudad.trim();
    if (!c || ciudades.includes(c)) return;
    setCiudades([...ciudades, c].sort((a, b) => a.localeCompare(b)));
    setNewCiudad('');
  };
  const removeCiudad = (c) => setCiudades(ciudades.filter((x) => x !== c));

  const parseList = (s) => s.split(',').map((t) => t.trim()).filter(Boolean);

  const addTipo = () => {
    const nombre = newTipoNombre.trim();
    if (!nombre) return;
    const subtipos = parseList(newSubtipos);
    if (newMedida === 'cintura_largo') {
      const cinturas = parseList(newCinturas);
      const largos = parseList(newLargos);
      if (cinturas.length === 0 || largos.length === 0) return;
      setTipos([...tipos, { id: uid(), nombre, medida: 'cintura_largo', cinturas, largos, subtipos }]);
      setNewCinturas('');
      setNewLargos('');
    } else {
      const tallas = parseList(newTipoTallas);
      if (tallas.length === 0) return;
      setTipos([...tipos, { id: uid(), nombre, medida: 'simple', tallas, subtipos }]);
      setNewTipoTallas('');
    }
    setNewTipoNombre('');
    setNewSubtipos('');
  };
  const removeTipo = (id) => setTipos(tipos.filter((t) => t.id !== id));

  const startEditSubtipos = (t) => { setEditingSubtiposId(t.id); setEditSubtipos((t.subtipos || []).join(', ')); };
  const saveEditSubtipos = (id) => {
    setTipos(tipos.map((t) => (t.id === id ? { ...t, subtipos: parseList(editSubtipos) } : t)));
    setEditingSubtiposId(null);
  };

  const [imported, setImported] = useState(false);
  const importarListado = () => {
    const existingDeptLower = new Set(departamentos.map((d) => d.toLowerCase()));
    const nuevosDept = IMPORT_DEPARTAMENTOS.filter((d) => !existingDeptLower.has(d.toLowerCase()));
    if (nuevosDept.length > 0) setDepartamentos([...departamentos, ...nuevosDept]);

    const existingTipoLower = new Set(tipos.map((t) => (t.nombre || '').toLowerCase()));
    const nuevosTipos = IMPORT_TIPOS.filter((t) => !existingTipoLower.has((t.nombre || '').toLowerCase())).map((t) => ({ id: uid(), ...t }));
    if (nuevosTipos.length > 0) setTipos([...tipos, ...nuevosTipos]);

    setImported(true);
    setTimeout(() => setImported(false), 4000);
  };

  const startEdit = (t) => {
    setEditingTipoId(t.id);
    setEditNombre(t.nombre);
    if (t.medida === 'cintura_largo') {
      setEditCinturas((t.cinturas || []).join(', '));
      setEditLargos((t.largos || []).join(', '));
    } else {
      setEditTallas((t.tallas || []).join(', '));
    }
  };
  const saveEdit = (t) => {
    if (t.medida === 'cintura_largo') {
      setTipos(tipos.map((x) => (x.id === t.id ? { ...x, nombre: editNombre, cinturas: parseList(editCinturas), largos: parseList(editLargos) } : x)));
    } else {
      setTipos(tipos.map((x) => (x.id === t.id ? { ...x, nombre: editNombre, tallas: parseList(editTallas) } : x)));
    }
    setEditingTipoId(null);
    setEditNombre('');
  };

  return (
    <div className="space-y-6">
      <button onClick={onBack} className="text-sm text-app-dim2 flex items-center gap-1">← Volver</button>

      <section>
        <h2 className="text-sm font-semibold mb-2 flex items-center gap-2"><Layers size={15} className="text-app-gold" /> Departamentos</h2>
        <div className="flex gap-2 mb-2">
          <input
            value={newDept}
            onChange={(e) => setNewDept(e.target.value)}
            onKeyDown={(e) => { if (e.key === 'Enter') addDept(); }}
            placeholder="Ej. Damas, Calzado, Accesorios…"
            className="flex-1 bg-app-panel border border-app-line rounded-lg px-3 py-2 text-sm"
          />
          <button onClick={addDept} className="px-3 rounded-lg bg-app-gold text-app-bg text-sm font-semibold">Agregar</button>
        </div>
        <details className="bg-app-panel border border-app-line rounded-lg group">
          <summary className="px-3 py-2.5 cursor-pointer font-medium text-sm list-none flex items-center justify-between hover:bg-app-active">
            <span>Departamentos ({departamentos.length})</span>
            <ChevronDown size={16} className="group-open:rotate-180 transition-transform" />
          </summary>
          <div className="px-3 pb-2.5 pt-1 border-t border-app-line">
            <div className="flex flex-wrap gap-1.5">
              {departamentos.map((d) => (
                <span key={d} className="bg-app-bg border border-app-line text-sm rounded-full pl-3 pr-1.5 py-1.5 flex items-center gap-1.5">
                  {d}
                  <BotonBorrar onConfirm={() => removeDept(d)} size={13} icono="x" />
                </span>
              ))}
              {departamentos.length === 0 && <p className="text-xs text-app-dim">Sin departamentos aún.</p>}
            </div>
          </div>
        </details>
      </section>

      <section>
        <h2 className="text-sm font-semibold mb-2 flex items-center gap-2"><Tag size={15} className="text-app-gold" /> Marcas y sus proveedores</h2>
        <MarcasConfig
          marcas={marcas}
          setMarcas={setMarcas}
          marcasProveedores={marcasProveedores}
          setMarcasProveedores={setMarcasProveedores}
          suppliers={suppliers}
          origen={origen}
        />
      </section>

      <section>
        <h2 className="text-sm font-semibold mb-2 flex items-center gap-2">🏙 Ciudades de compra</h2>
        <div className="flex gap-2 mb-2">
          <input
            value={newCiudad}
            onChange={(e) => setNewCiudad(e.target.value)}
            onKeyDown={(e) => { if (e.key === 'Enter') addCiudad(); }}
            placeholder="Ej. Guangzhou, Miami, Ciudad de Panamá…"
            className="flex-1 bg-app-panel border border-app-line rounded-lg px-3 py-2 text-sm"
          />
          <button onClick={addCiudad} className="px-3 rounded-lg bg-app-gold text-app-bg text-sm font-semibold">Agregar</button>
        </div>
        <div className="flex flex-wrap gap-1.5">
          {ciudades.map((c) => (
            <span key={c} className="bg-app-panel border border-app-line text-sm rounded-full pl-3 pr-1.5 py-1.5 flex items-center gap-1.5">
              {c}
              <BotonBorrar onConfirm={() => removeCiudad(c)} size={13} icono="x" />
            </span>
          ))}
          {ciudades.length === 0 && <p className="text-xs text-app-dim">Sin ciudades aún.</p>}
        </div>
      </section>

      <section>
        <h2 className="text-sm font-semibold mb-2 flex items-center gap-2">🏭 Fábricas / Proveedores de fábrica</h2>
        <FabricasList fabricas={fabricas} setFabricas={setFabricas} />
      </section>

      <section>
        <h2 className="text-sm font-semibold mb-2 flex items-center gap-2"><Package size={15} className="text-app-gold" /> Tipos de producto y sus tallas</h2>
        <div className="bg-app-panel border border-app-line rounded-xl p-3 space-y-2 mb-3">
          <input
            value={newTipoNombre}
            onChange={(e) => setNewTipoNombre(e.target.value)}
            placeholder="Nombre del tipo (ej. Camisa, Sudadera…)"
            className="w-full bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm"
          />

          <div className="flex gap-2">
            <button
              onClick={() => setNewMedida('simple')}
              className={`flex-1 py-2 rounded-lg text-xs font-medium border ${newMedida === 'simple' ? 'bg-app-gold text-app-bg border-app-gold' : 'border-app-line text-app-dim2'}`}
            >
              Talla única
            </button>
            <button
              onClick={() => setNewMedida('cintura_largo')}
              className={`flex-1 py-2 rounded-lg text-xs font-medium border ${newMedida === 'cintura_largo' ? 'bg-app-gold text-app-bg border-app-gold' : 'border-app-line text-app-dim2'}`}
            >
              Cintura + Largo
            </button>
          </div>

          {newMedida === 'simple' ? (
            <div className="space-y-1.5">
              <p className="text-xs text-app-dim2">Grupos predefinidos (toca para usar):</p>
              <PresetButtons
                grupos={GRUPOS_TALLAS.filter((g) => !['Cintura pantalón', 'Largo pantalón'].includes(g.label))}
                onSelect={(g) => setNewTipoTallas(g.tallas.join(', '))}
              />
              <input
                value={newTipoTallas}
                onChange={(e) => setNewTipoTallas(e.target.value)}
                placeholder="O escribe tallas separadas por coma: XS, S, M…"
                className="w-full bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm"
              />
            </div>
          ) : (
            <>
              <div className="space-y-1.5">
                <p className="text-xs text-app-dim2">Cinturas predefinidas:</p>
                <PresetButtons
                  grupos={GRUPOS_TALLAS.filter((g) => g.label === 'Cintura pantalón')}
                  onSelect={(g) => setNewCinturas(g.tallas.join(', '))}
                />
                <input
                  value={newCinturas}
                  onChange={(e) => setNewCinturas(e.target.value)}
                  placeholder="O escribe cinturas: 28, 30, 32, 34"
                  className="w-full bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm"
                />
              </div>
              <div className="space-y-1.5">
                <p className="text-xs text-app-dim2">Largos predefinidos:</p>
                <PresetButtons
                  grupos={GRUPOS_TALLAS.filter((g) => g.label === 'Largo pantalón')}
                  onSelect={(g) => setNewLargos(g.tallas.join(', '))}
                />
                <input
                  value={newLargos}
                  onChange={(e) => setNewLargos(e.target.value)}
                  placeholder="O escribe largos: 30, 32, 34"
                  className="w-full bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm"
                />
              </div>
              <p className="text-xs text-app-dim">Se generará una combinación por cada cintura × largo, ej. 32/34.</p>
            </>
          )}

          <div className="space-y-1.5">
            <p className="text-xs text-app-dim2">Subtipos predefinidos (opcional):</p>
            <PresetButtons
              grupos={GRUPOS_SUBTIPOS}
              onSelect={(g) => {
                const actual = newSubtipos ? newSubtipos.split(',').map((s) => s.trim()).filter(Boolean) : [];
                const nuevos = g.items.filter((i) => !actual.includes(i));
                setNewSubtipos([...actual, ...nuevos].join(', '));
              }}
            />
            <input
              value={newSubtipos}
              onChange={(e) => setNewSubtipos(e.target.value)}
              placeholder="O escribe subtipos: Slim, Straight, Skinny"
              className="w-full bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm"
            />
          </div>

          <button onClick={addTipo} className="w-full py-2 rounded-lg bg-app-gold text-app-bg text-sm font-semibold">Agregar tipo</button>
        </div>

        {tipos.length > 0 && (
          <div className="mt-3 mb-2">
            <input
              type="text"
              placeholder="🔍 Buscar tipo..."
              value={queryTipos}
              onChange={(e) => setQueryTipos(e.target.value)}
              className="w-full bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm"
            />
          </div>
        )}

        <div className="space-y-2">
          {tipos.filter((t) => (t.nombre || '').toLowerCase().includes(queryTipos.toLowerCase())).map((t) => (
            <div key={t.id} className="bg-app-panel border border-app-line rounded-xl p-3">
              <div className="flex items-center justify-between">
                {editingTipoId === t.id ? (
                  <input
                    value={editNombre}
                    onChange={(e) => setEditNombre(e.target.value)}
                    placeholder="Nombre del tipo"
                    className="flex-1 bg-app-bg border border-app-line rounded-lg px-2.5 py-1.5 text-sm font-medium"
                  />
                ) : (
                  <p className="text-sm font-medium">{t.nombre}</p>
                )}
                <BotonBorrar onConfirm={() => removeTipo(t.id)} size={15} />
              </div>
              {editingTipoId === t.id ? (
                <div className="space-y-2 mt-2">
                  {t.medida === 'cintura_largo' ? (
                    <>
                      <PresetButtons
                        grupos={GRUPOS_TALLAS.filter((g) => g.label === 'Cintura pantalón')}
                        onSelect={(g) => setEditCinturas(g.tallas.join(', '))}
                      />
                      <input value={editCinturas} onChange={(e) => setEditCinturas(e.target.value)}
                        placeholder="Cinturas" className="w-full bg-app-bg border border-app-line rounded-lg px-3 py-1.5 text-sm" />
                      <PresetButtons
                        grupos={GRUPOS_TALLAS.filter((g) => g.label === 'Largo pantalón')}
                        onSelect={(g) => setEditLargos(g.tallas.join(', '))}
                      />
                      <input value={editLargos} onChange={(e) => setEditLargos(e.target.value)}
                        placeholder="Largos" className="w-full bg-app-bg border border-app-line rounded-lg px-3 py-1.5 text-sm" />
                    </>
                  ) : (
                    <>
                      <PresetButtons
                        grupos={GRUPOS_TALLAS.filter((g) => !['Cintura pantalón', 'Largo pantalón'].includes(g.label))}
                        onSelect={(g) => setEditTallas(g.tallas.join(', '))}
                      />
                      <input value={editTallas} onChange={(e) => setEditTallas(e.target.value)}
                        className="w-full bg-app-bg border border-app-line rounded-lg px-3 py-1.5 text-sm" />
                    </>
                  )}
                  <button onClick={() => saveEdit(t)} className="px-3 py-1.5 rounded-lg bg-app-gold text-app-bg text-xs font-semibold">Guardar</button>
                </div>
              ) : (
                <button onClick={() => startEdit(t)} className="text-xs text-app-dim2 mt-1 underline decoration-dotted block">
                  {t.medida === 'cintura_largo'
                    ? `Cintura: ${(t.cinturas || []).join(', ')} · Largo: ${(t.largos || []).join(', ')}`
                    : `Tallas: ${(t.tallas || []).join(', ')}`}
                </button>
              )}

              {editingSubtiposId === t.id ? (
                <div className="space-y-2 mt-2">
                  <PresetButtons
                    grupos={GRUPOS_SUBTIPOS}
                    onSelect={(g) => {
                      const actual = editSubtipos ? editSubtipos.split(',').map((s) => s.trim()).filter(Boolean) : [];
                      const nuevos = g.items.filter((i) => !actual.includes(i));
                      setEditSubtipos([...actual, ...nuevos].join(', '));
                    }}
                  />
                  <div className="flex gap-2">
                    <input value={editSubtipos} onChange={(e) => setEditSubtipos(e.target.value)}
                      placeholder="Slim, Straight, Skinny"
                      className="flex-1 bg-app-bg border border-app-line rounded-lg px-3 py-1.5 text-sm" />
                    <button onClick={() => saveEditSubtipos(t.id)} className="px-3 rounded-lg bg-app-gold text-app-bg text-xs font-semibold">Guardar</button>
                  </div>
                </div>
              ) : (
                <button onClick={() => startEditSubtipos(t)} className="text-xs text-app-dim2 mt-1 underline decoration-dotted block">
                  Subtipos: {(t.subtipos || []).length > 0 ? t.subtipos.join(', ') : 'ninguno (toca para agregar)'}
                </button>
              )}
            </div>
          ))}
          {tipos.length === 0 && <p className="text-xs text-app-dim">Sin tipos aún.</p>}
        </div>
      </section>
    </div>
  );
}

// ---------- Proveedores ----------
function Proveedores({ suppliers = [], setSuppliers, origen, puedoBorrar = true }) {
  const [showForm, setShowForm] = useState(false);
  const [form, setForm] = useState({ nombre: '', contacto: '', email: '', telefono: '' });
  const [editingId, setEditingId] = useState(null);
  const [editForm, setEditForm] = useState({});

  // Cada proveedor pertenece a UN país de compra. Aquí solo se ven los del origen activo.
  const delOrigen = suppliers.filter((s) => s.origen === origen?.id);
  // Proveedores creados antes de esta regla: no tienen país asignado
  const sinPais = suppliers.filter((s) => !s.origen);

  const asignarPais = (id) => {
    setSuppliers(suppliers.map((s) => (s.id === id ? { ...s, origen: origen.id } : s)));
  };

  const addSupplier = () => {
    if (!form.nombre) return;
    setSuppliers([...suppliers, { ...form, id: uid(), origen: origen?.id || '' }]);
    setForm({ nombre: '', contacto: '', email: '', telefono: '' });
    setShowForm(false);
  };

  const startEdit = (s) => {
    setEditingId(s.id);
    setEditForm({ nombre: s.nombre, contacto: s.contacto, email: s.email, telefono: s.telefono });
  };

  const saveEdit = (id) => {
    if (!editForm.nombre) return;
    setSuppliers(suppliers.map((s) => s.id === id ? { ...s, ...editForm } : s));
    setEditingId(null);
  };

  const removeSupplier = (id) => setSuppliers(suppliers.filter((s) => s.id !== id));

  // ---------- Importar proveedores desde Excel ----------
  const [importPreview, setImportPreview] = useState(null);
  const [importError, setImportError] = useState('');

  const normalizar = (t) => String(t ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toLowerCase();
  const MAPA_PAIS = {
    china: 'china', cn: 'china',
    usa: 'usa', 'estados unidos': 'usa', eeuu: 'usa', 'ee.uu.': 'usa', us: 'usa',
    panama: 'panama', pa: 'panama',
  };

  const FILAS_PLANTILLA = [
    { Nombre: 'Guangzhou Trading Co', Contacto: 'Li Wei', Correo: 'ventas@gztrading.cn', Teléfono: '+86 20 1234 5678', País: 'China' },
    { Nombre: 'Miami Wholesale Inc', Contacto: 'John Smith', Correo: 'sales@miamiws.com', Teléfono: '+1 305 555 0100', País: 'USA' },
    { Nombre: 'Zona Libre Corp', Contacto: 'Ana Ruiz', Correo: 'info@zonalibre.pa', Teléfono: '+507 430 1122', País: 'Panamá' },
  ];

  const leerArchivoProveedores = (e) => {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file) return;
    setImportError('');
    setImportPreview(null);
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const filas = XLSX.utils.sheet_to_json(ws, { defval: '' });
        if (filas.length === 0) return setImportError('El archivo no tiene filas.');

        // Reunir los encabezados de todas las filas (no solo la primera, por si viene incompleta)
        const claves = [...new Set(filas.flatMap((f) => Object.keys(f)))];
        const buscarCol = (...alias) => claves.find((k) => alias.includes(normalizar(k)));
        const kNombre = buscarCol('nombre', 'empresa', 'proveedor', 'nombre de la empresa', 'razon social');
        const kContacto = buscarCol('contacto', 'persona de contacto', 'vendedor');
        const kCorreo = buscarCol('correo', 'email', 'e-mail', 'mail');
        const kTel = buscarCol('telefono', 'tel', 'celular', 'whatsapp', 'movil');
        const kPais = buscarCol('pais', 'origen', 'country');

        if (!kNombre) {
          return setImportError('No encontré la columna "Nombre". Descarga la plantilla para ver el formato exacto.');
        }

        const nuevos = [];
        let duplicados = 0;
        let sinNombre = 0;

        filas.forEach((f) => {
          const nombre = String(f[kNombre] ?? '').trim();
          if (!nombre) { sinNombre++; return; }
          const paisTxt = kPais ? normalizar(f[kPais]) : '';
          const origenId = MAPA_PAIS[paisTxt] || origen?.id || '';
          const yaExiste =
            suppliers.some((s) => normalizar(s.nombre) === normalizar(nombre) && s.origen === origenId) ||
            nuevos.some((n) => normalizar(n.nombre) === normalizar(nombre) && n.origen === origenId);
          if (yaExiste) { duplicados++; return; }
          nuevos.push({
            id: uid(),
            nombre,
            contacto: kContacto ? String(f[kContacto] ?? '').trim() : '',
            email: kCorreo ? String(f[kCorreo] ?? '').trim() : '',
            telefono: kTel ? String(f[kTel] ?? '').trim() : '',
            origen: origenId,
          });
        });

        if (nuevos.length === 0) {
          return setImportError(
            duplicados > 0
              ? `Todos los proveedores del archivo (${duplicados}) ya existen.`
              : 'No encontré proveedores válidos en el archivo.'
          );
        }
        setImportPreview({ nuevos, duplicados, sinNombre });
      } catch (err) {
        setImportError('No pude leer el archivo. Debe ser .xlsx o .csv');
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const confirmarImportacion = () => {
    setSuppliers([...suppliers, ...importPreview.nuevos]);
    setImportPreview(null);
  };

  return (
    <div className="space-y-4">
      {origen && (
        <div className="flex items-center gap-2 bg-app-panel border border-app-line rounded-xl px-3 py-2">
          <span className="text-lg">{origen.emoji}</span>
          <span className="text-xs text-app-dim2">Proveedores de <span className="text-app-white font-medium">{origen.label}</span></span>
        </div>
      )}

      {sinPais.length > 0 && (
        <div className="bg-app-goldbg border border-app-gold rounded-xl p-3 space-y-2">
          <p className="text-xs text-app-gold font-medium">
            ⚠ {sinPais.length} proveedor{sinPais.length !== 1 ? 'es' : ''} sin país de compra
          </p>
          <p className="text-xs text-app-dim2">
            Cada proveedor pertenece a un solo país. Asígnalos para que aparezcan donde corresponde, o bórralos.
          </p>
          {sinPais.map((s) => (
            <div key={s.id} className="flex items-center justify-between gap-2 bg-app-panel border border-app-line rounded-lg px-3 py-2">
              <span className="text-sm truncate">{s.nombre}</span>
              <div className="flex items-center gap-1.5 shrink-0">
                <button
                  onClick={() => asignarPais(s.id)}
                  className="text-xs bg-app-gold text-app-bg rounded-lg px-2.5 py-1.5 font-semibold whitespace-nowrap"
                >
                  Asignar a {origen?.emoji} {origen?.label}
                </button>
                {puedoBorrar && <BotonBorrar onConfirm={() => removeSupplier(s.id)} size={14} />}
              </div>
            </div>
          ))}
        </div>
      )}

      {/* Importar proveedores desde Excel */}
      <div className="bg-app-panel border border-app-line rounded-xl p-3 space-y-2">
        <p className="text-xs uppercase tracking-wide text-app-dim2">Cargar varios de una vez</p>
        <div className="flex gap-2">
          <label className="w-full py-2.5 rounded-lg border border-app-line text-xs text-app-sky flex items-center justify-center gap-1.5 active:bg-app-active cursor-pointer">
            <Upload size={14} /> Importar Excel
            <input type="file" accept=".xlsx,.xls,.csv" onChange={leerArchivoProveedores} className="hidden" />
          </label>
        </div>
        <BotonesPlantilla
          filas={FILAS_PLANTILLA}
          nombreHoja="Plantilla"
          nombreArchivo="plantilla-proveedores.xlsx"
          cols={[{ wch: 28 }, { wch: 20 }, { wch: 28 }, { wch: 20 }, { wch: 12 }]}
        />
        <p className="text-xs text-app-dim">
          Columnas: <span className="font-mono text-app-dim2">Nombre</span> (obligatoria), Contacto, Correo, Teléfono y País.
          Si dejas País vacío, se asigna a {origen?.emoji} {origen?.label}.
        </p>

        {importError && (
          <div className="flex items-start gap-2 bg-app-redbg text-app-red2 text-xs rounded-lg px-3 py-2">
            <AlertCircle size={14} className="shrink-0 mt-0.5" /> {importError}
          </div>
        )}

        {importPreview && (
          <div className="bg-app-blue border border-app-line2 rounded-lg p-3 space-y-2">
            <p className="text-xs text-app-sky font-medium">
              Listos para importar: {importPreview.nuevos.length} proveedor{importPreview.nuevos.length !== 1 ? 'es' : ''}
            </p>
            <div className="max-h-40 overflow-y-auto space-y-1">
              {importPreview.nuevos.map((n) => (
                <div key={n.id} className="flex items-center justify-between gap-2 text-xs">
                  <span className="text-app-light truncate">{n.nombre}</span>
                  <span className="text-app-dim2 shrink-0">
                    {ORIGENES.find((o) => o.id === n.origen)?.emoji || '—'}
                  </span>
                </div>
              ))}
            </div>
            {(importPreview.duplicados > 0 || importPreview.sinNombre > 0) && (
              <p className="text-xs text-app-gold">
                {importPreview.duplicados > 0 && `Se omiten ${importPreview.duplicados} que ya existen. `}
                {importPreview.sinNombre > 0 && `Se omiten ${importPreview.sinNombre} filas sin nombre.`}
              </p>
            )}
            <div className="flex gap-2">
              <button onClick={() => setImportPreview(null)} className="flex-1 py-2 rounded-lg border border-app-line text-xs text-app-dim2">
                Cancelar
              </button>
              <button onClick={confirmarImportacion} className="flex-1 py-2 rounded-lg bg-app-sky text-app-bg text-xs font-semibold">
                Importar {importPreview.nuevos.length}
              </button>
            </div>
          </div>
        )}
      </div>

      {!showForm ? (
        <button onClick={() => setShowForm(true)} className="w-full py-3 rounded-xl border border-dashed border-app-line3 text-sm text-app-dim2 flex items-center justify-center gap-2 active:bg-app-panel">
          <Plus size={16} /> Agregar proveedor
        </button>
      ) : (        <div className="bg-app-panel border border-app-line rounded-xl shadow-app p-4 space-y-3">
          <div className="flex justify-between items-center">
            <span className="text-sm font-medium">Nuevo proveedor</span>
            <button onClick={() => setShowForm(false)}><X size={16} className="text-app-dim" /></button>
          </div>
          <input placeholder="Nombre de la empresa" value={form.nombre} onChange={(e) => setForm({ ...form, nombre: e.target.value })} className="w-full bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm" />
          <input placeholder="Persona de contacto" value={form.contacto} onChange={(e) => setForm({ ...form, contacto: e.target.value })} className="w-full bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm" />
          <input placeholder="Correo" value={form.email} onChange={(e) => setForm({ ...form, email: e.target.value })} className="w-full bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm" />
          <input placeholder="Teléfono" value={form.telefono} onChange={(e) => setForm({ ...form, telefono: e.target.value })} className="w-full bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm" />
          <button onClick={addSupplier} className="w-full py-2.5 rounded-lg bg-app-gold text-app-bg text-sm font-semibold">Guardar proveedor</button>
        </div>
      )}

      <div className="space-y-2">
        {delOrigen.length === 0 && (
          <p className="text-center text-xs text-app-dim py-6">
            Aún no tienes proveedores de {origen?.label}.
          </p>
        )}
        {delOrigen.map((s) => (
          <div key={s.id} className="bg-app-panel border border-app-line rounded-xl p-3">
            {editingId === s.id ? (
              <div className="space-y-2">
                <div className="flex justify-between items-center mb-1">
                  <span className="text-xs font-medium text-app-dim2 uppercase tracking-wide">Editando proveedor</span>
                  <button onClick={() => setEditingId(null)}><X size={15} className="text-app-dim" /></button>
                </div>
                <input placeholder="Nombre de la empresa" value={editForm.nombre} onChange={(e) => setEditForm({ ...editForm, nombre: e.target.value })} className="w-full bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm" />
                <input placeholder="Persona de contacto" value={editForm.contacto} onChange={(e) => setEditForm({ ...editForm, contacto: e.target.value })} className="w-full bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm" />
                <input placeholder="Correo" value={editForm.email} onChange={(e) => setEditForm({ ...editForm, email: e.target.value })} className="w-full bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm" />
                <input placeholder="Teléfono" value={editForm.telefono} onChange={(e) => setEditForm({ ...editForm, telefono: e.target.value })} className="w-full bg-app-bg border border-app-line rounded-lg px-3 py-2 text-sm" />
                <div className="flex gap-2 pt-1">
                  <button onClick={() => setEditingId(null)} className="flex-1 py-2 rounded-lg border border-app-line text-sm text-app-dim2">Cancelar</button>
                  <button onClick={() => saveEdit(s.id)} className="flex-1 py-2 rounded-lg bg-app-gold text-app-bg text-sm font-semibold">Guardar</button>
                </div>
              </div>
            ) : (
              <div className="flex items-center justify-between gap-2">
                <div className="min-w-0">
                  <p className="text-sm font-medium truncate">{s.nombre}</p>
                  {s.contacto && <p className="text-xs text-app-dim truncate">{s.contacto}</p>}
                  {s.email && <p className="text-xs text-app-dim truncate">{s.email}</p>}
                  {s.telefono && <p className="text-xs text-app-dim">{s.telefono}</p>}
                </div>
                <div className="flex items-center gap-3 shrink-0">
                  <button onClick={() => startEdit(s)} className="text-app-sky active:opacity-70">
                    <Pencil size={15} />
                  </button>
                  {puedoBorrar && <BotonBorrar onConfirm={() => removeSupplier(s.id)} size={15} />}
                </div>
              </div>
            )}
          </div>
        ))}
        {suppliers.length === 0 && <p className="text-center text-xs text-app-dim py-6">Sin proveedores.</p>}
      </div>
    </div>
  );
}
