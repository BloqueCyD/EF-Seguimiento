/* =========================================================================
   CONFIGURACIÓN
   ========================================================================= */

// Pegá aquí la URL de tu Apps Script NUEVO (la que termina en /exec).
// Se genera al desplegar el Code.gs de este mismo paquete. Ver INSTRUCCIONES.md.
const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzDTgbG6cNP5JA5Py8Q4ZWiWC6w1v-hGgb5sfnSTNoUIzfw2NeTzw7mbi25F6tqj4xE4g/exec";

// Cada cuánto se vuelve a sincronizar automáticamente (ms). 20 minutos:
// para este tipo de datos no hace falta chequear más seguido, y así se
// minimizan las interrupciones mientras estás mirando la página. El botón
// "Actualizar" de arriba siempre funciona al toque, sin esperar este plazo.
const AUTO_REFRESH_MS = 20 * 60 * 1000;

// Nombre de la hoja/pestaña dentro del Google Sheet que contiene los datos.
const NOMBRE_HOJA = "2026";

// Nombres de columnas tal cual están en tu Excel. Si en algún momento
// cambiás el nombre de una columna en la planilla, solo hay que
// actualizarlo acá.
const COL = {
  expediente: "Nro Expediente",
  comision: "Comisión",
  bloque: "Bloque",
  autor: "Autor/a",
  coautores: "Coautor/es",
  tipo: "Tipo de proyecto",
  sumario: "Sumario",
  resumen: "Resumen",
  fechaIngreso: "Fecha de Ingreso",
  otrosGiros: "¿Tiene otros giros?",
  reunionAsesores: "¿Reunión Asesores?",
  fechaReunionAsesores: "Fecha Reunión Asesores",
  pasoDiputados: "¿Pasó a diputados?",
  reunionDiputados: "¿Reunión Diputados?",
  fechaReunionDiputados: "Fecha Reunión Diputados",
  despachado: "¿Se despachó?",
  firmanDictamen: "¿Quienes firman dictamen?",
  noFirmanDictamen: "¿Quienes no firman dictamen?",
  numeroDespacho: "Número de despacho",
  aprobadoRecinto: "¿Se aprobó en el recinto?",
  fechaAprobacion: "¿Fecha de aprobación?",
  observaciones: "Observaciones",
};

// Las 4 fechas que existen a lo largo del expediente, usadas en conjunto
// por el filtro de fecha (un expediente matchea si CUALQUIERA de estas
// cuatro fechas cae dentro del rango elegido).
const COLUMNAS_FECHA = [COL.fechaIngreso, COL.fechaReunionAsesores, COL.fechaReunionDiputados, COL.fechaAprobacion];

/* =========================================================================
   ETAPAS (Kanban)
   ========================================================================= */

// Las 5 etapas del seguimiento, en orden. "match" se evalúa de la última
// a la primera, así una etapa más avanzada gana si se cumplen varias
// condiciones a la vez (ej. un expediente ya sancionado también podría
// tener "reunión diputados" en Sí, pero nos interesa que aparezca en
// Sancionado, no en Diputados).
const ETAPAS = [
  {
    id: "ingresado",
    nombre: "Ingresado",
    color: "#9C8B5E",
    match: (row) => true, // etapa por defecto: aún no pasó por asesores ni etapas siguientes
  },
  {
    id: "asesores",
    nombre: "Asesores",
    color: "#1E3A5F",
    match: (row) => esAfirmativo(row[COL.reunionAsesores]),
  },
  {
    id: "diputados",
    nombre: "Diputados",
    color: "#C08A2E",
    match: (row) => esAfirmativo(row[COL.reunionDiputados]),
  },
  {
    id: "despacho",
    nombre: "Despacho",
    color: "#6E5A87",
    match: (row) => String(row[COL.numeroDespacho] || "").trim() !== "",
  },
  {
    id: "sancionado",
    nombre: "Sancionado",
    color: "#6B7A4F",
    match: (row) => esAfirmativo(row[COL.aprobadoRecinto]),
  },
];

function getEtapa(row) {
  // de la última a la primera, así la etapa más avanzada gana
  for (let i = ETAPAS.length - 1; i >= 0; i--) {
    if (ETAPAS[i].match(row)) return ETAPAS[i];
  }
  return ETAPAS[0];
}

// Valor especial para agrupar expedientes sin ninguna fecha detectable.
const SIN_FECHA = "Sin fecha";

function parsearFecha(valorCelda) {
  if (valorCelda === null || valorCelda === undefined || valorCelda === "") return null;
  if (valorCelda instanceof Date) {
    return isNaN(valorCelda.getTime()) ? null : valorCelda;
  }
  const str = String(valorCelda).trim();
  if (!str) return null;

  // Formato "AAAA-MM-DD" (el que manda el Apps Script cuando la celda es
  // una fecha real de Google Sheets). La armamos como fecha LOCAL directo,
  // sin pasar por new Date(str), porque eso la interpreta como UTC
  // medianoche y en Argentina (UTC-3) cae un día antes.
  let m = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) {
    const fecha = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    return isNaN(fecha.getTime()) ? null : fecha;
  }

  // Formato "DD/MM/AAAA" o "DD-MM-AAAA" (fecha tipeada como texto en la
  // planilla, con el orden día/mes/año habitual en Argentina).
  m = str.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if (m) {
    let dia = Number(m[1]), mes = Number(m[2]), anio = Number(m[3]);
    if (anio < 100) anio += 2000;
    if (mes >= 1 && mes <= 12 && dia >= 1 && dia <= 31) {
      const fecha = new Date(anio, mes - 1, dia);
      return isNaN(fecha.getTime()) ? null : fecha;
    }
  }

  // Último recurso: dejar que el navegador intente interpretarla.
  const fecha = new Date(str);
  return isNaN(fecha.getTime()) ? null : fecha;
}

function parsearAnio(valorCelda) {
  const fecha = parsearFecha(valorCelda);
  if (!fecha) return null;
  const anio = fecha.getFullYear();
  if (anio <= 1990 || anio >= 2100) return null;
  return String(anio);
}

// Año del expediente: prioriza Fecha de Ingreso; si no hay, busca el año
// dentro del número de expediente (ej. "1234-D-2026"); si no, "Sin fecha".
function getAnio(row) {
  const anioIngreso = parsearAnio(row[COL.fechaIngreso]);
  if (anioIngreso) return anioIngreso;
  const match = String(row[COL.expediente] || "").match(/\d+-[A-Za-z]-(\d{4})/);
  if (match) return match[1];
  return SIN_FECHA;
}

// Devuelve todas las fechas válidas (objetos Date) que tiene un expediente,
// usadas por el filtro de rango de fechas combinado.
function getFechasDelItem(row) {
  return COLUMNAS_FECHA.map((c) => parsearFecha(row[c])).filter(Boolean);
}

/* =========================================================================
   ESTADO GLOBAL
   ========================================================================= */

let RAW_DATA = null;
let FILTROS = {
  comisiones: [],     // array vacío = todas
  bloques: [],         // array vacío = todos
  tipos: [],            // array vacío = todos
  autores: [],           // array vacío = todos
  fechasSancion: [],     // array vacío = todas (fechas de aprobación en el recinto)
  soloTablas: false,     // true = solo sancionados "por tablas"
  soloConDespacho: false, // true = solo sancionados que NO fueron por tablas
  etapas: [],           // array vacío = todas
  fechaDesde: "",       // "YYYY-MM-DD" o ""
  fechaHasta: "",       // "YYYY-MM-DD" o ""
  busqueda: "",
};
let autoRefreshTimer = null;
let EXPORTAR_SOLO_FILTRADO = true;

/* =========================================================================
   HELPERS
   ========================================================================= */

function normaliza(v) {
  if (v === null || v === undefined) return "";
  return String(v).trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

// Las columnas "¿...?" de la planilla pueden estar cargadas de formas muy
// distintas según cómo las armaste: casilla de verificación real de Google
// Sheets (que llega como booleano true/false), texto "Sí"/"No", o palabras
// como "Aprobado"/"Hecho"/"X". Esta función reconoce todas esas variantes
// como "sí" sin que haya que tocar la planilla.
function esAfirmativo(valor) {
  if (valor === true) return true;
  if (valor === false || valor === null || valor === undefined) return false;
  const v = normaliza(valor);
  if (v === "") return false;
  const afirmativos = ["si", "x", "hecho", "ok", "1", "true", "aprobado", "aprobada", "realizada", "realizado", "sancionado", "sancionada"];
  return afirmativos.includes(v) || v.startsWith("aprobado") || v.startsWith("si ");
}

function escapeHtml(str) {
  if (str === null || str === undefined) return "";
  return String(str).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// Recorta espacios al principio/final y colapsa espacios dobles internos
// (una causa muy común de que "Salud" y "Salud " —con un espacio de más
// al final, invisible en la planilla— aparezcan como dos filtros distintos).
function limpiarEspacios(v) {
  return String(v || "").replace(/\s+/g, " ").trim();
}

// A veces la misma comisión/bloque/autor está cargado con variaciones
// menores (espacios de más, mayúsculas/minúsculas distintas) en filas
// distintas del Sheet. Esta función arma un mapa "valor normalizado ->
// mejor variante" para que todas esas filas se agrupen bajo un solo
// filtro en vez de aparecer duplicadas.
function construirMapaCanonico(valoresCrudos) {
  const porClave = new Map(); // clave normalizada -> Map(variante -> cantidad)
  valoresCrudos.forEach((raw) => {
    const val = limpiarEspacios(raw);
    if (!val) return;
    const clave = normaliza(val);
    if (!clave) return;
    if (!porClave.has(clave)) porClave.set(clave, new Map());
    const sub = porClave.get(clave);
    sub.set(val, (sub.get(val) || 0) + 1);
  });
  const canonico = new Map(); // clave normalizada -> variante elegida
  porClave.forEach((sub, clave) => {
    let mejorLabel = null, mejorCantidad = -1;
    sub.forEach((cantidad, label) => {
      if (cantidad > mejorCantidad || (cantidad === mejorCantidad && (mejorLabel === null || label < mejorLabel))) {
        mejorLabel = label; mejorCantidad = cantidad;
      }
    });
    canonico.set(clave, mejorLabel);
  });
  return canonico;
}

function canonicalizar(mapa, raw, fallback) {
  const val = limpiarEspacios(raw);
  if (!val) return fallback;
  const clave = normaliza(val);
  return mapa.get(clave) || val;
}

function showToast(msg, isError = false) {
  const t = document.getElementById("toast");
  t.textContent = msg;
  t.className = "toast show" + (isError ? " error" : "");
  setTimeout(() => { t.className = "toast"; }, 3200);
}

function setSyncStatus(state, text) {
  const dot = document.getElementById("sync-dot");
  const txt = document.getElementById("sync-text");
  dot.className = "sync-dot " + (state === "ok" ? "" : state);
  txt.textContent = text;
}

// Formatea una fecha "YYYY-MM-DD" (como la manda el Apps Script) a "DD/MM/YYYY"
// para mostrar. Parsea el string directamente, sin pasar por new Date(), para
// evitar que el huso horario corra la fecha un día para atrás o adelante.
function formatFecha(valorCelda) {
  const fecha = parsearFecha(valorCelda);
  if (!fecha) return null;
  return `${String(fecha.getDate()).padStart(2, "0")}/${String(fecha.getMonth() + 1).padStart(2, "0")}/${fecha.getFullYear()}`;
}

function animarNumero(el, destino, duracionMs = 700) {
  if (!el) return;
  const t0 = performance.now();
  function paso(ahora) {
    const p = Math.min(1, (ahora - t0) / duracionMs);
    const eased = 1 - Math.pow(1 - p, 3);
    el.textContent = Math.round(destino * eased);
    if (p < 1) requestAnimationFrame(paso);
  }
  requestAnimationFrame(paso);
}

/* =========================================================================
   CARGA DE DATOS
   ========================================================================= */

// Guarda dónde estaba parado el usuario (scroll de la página, de la tabla
// y de cada columna del Kanban) para restaurarlo después de volver a
// construir el HTML. Sin esto, cada actualización automática "reiniciaba"
// la vista al tope de la página.
function capturarScroll() {
  const estado = {
    pagina: window.scrollY,
    tabla: document.getElementById("table-scroll")?.scrollTop || 0,
    kanban: {},
  };
  document.querySelectorAll(".kcol-body").forEach((col) => {
    const etapaId = col.closest(".kcol")?.dataset.etapaId;
    if (etapaId) estado.kanban[etapaId] = col.scrollTop;
  });
  return estado;
}

function restaurarScroll(estado) {
  if (!estado) return;
  requestAnimationFrame(() => {
    window.scrollTo(0, estado.pagina);
    const tabla = document.getElementById("table-scroll");
    if (tabla) tabla.scrollTop = estado.tabla;
    document.querySelectorAll(".kcol-body").forEach((col) => {
      const etapaId = col.closest(".kcol")?.dataset.etapaId;
      if (etapaId && estado.kanban[etapaId] !== undefined) col.scrollTop = estado.kanban[etapaId];
    });
  });
}

async function cargarDatos(manual = false) {
  const btn = document.getElementById("refresh-btn");
  btn.classList.add("spinning");
  setSyncStatus("loading", manual ? "Actualizando…" : "Sincronizando…");
  const scrollPrevio = capturarScroll();

  try {
    const res = await fetch(APPS_SCRIPT_URL + "?action=getAll&_=" + Date.now());
    if (!res.ok) throw new Error("Error HTTP " + res.status);
    const json = await res.json();
    if (!json.success) throw new Error(json.error || "Error desconocido del script");

    RAW_DATA = json.data;
    document.getElementById("error-zone").innerHTML = "";
    document.getElementById("loading-screen").style.display = "none";
    document.getElementById("app-content").style.display = "block";

    setSyncStatus("ok", "Conectado");
    document.getElementById("footer-time").textContent = new Date().toLocaleString("es-AR");

    render();
    restaurarScroll(scrollPrevio);
  } catch (err) {
    console.error(err);
    setSyncStatus("error", "Sin conexión");
    document.getElementById("loading-screen").style.display = "none";
    document.getElementById("app-content").style.display = "block";
    document.getElementById("error-zone").innerHTML = `
      <div class="error-banner">
        <span>⚠️</span>
        <div>
          <b>No se pudo conectar con la planilla.</b><br>
          ${err.message}. Verificá que la URL del Apps Script en comisiones.js sea correcta y que esté
          implementada como "Cualquier usuario" puede acceder, y probá actualizar de nuevo.
        </div>
      </div>`;
    if (RAW_DATA) { render(); restaurarScroll(scrollPrevio); } // mantenemos lo último que funcionó
  } finally {
    btn.classList.remove("spinning");
  }
}

function reiniciarAutoRefresh() {
  if (autoRefreshTimer) clearInterval(autoRefreshTimer);
  autoRefreshTimer = setInterval(() => cargarDatos(false), AUTO_REFRESH_MS);
}

/* =========================================================================
   TRANSFORMACIÓN: filas del sheet -> items enriquecidos
   ========================================================================= */

function getAllItems() {
  if (!RAW_DATA || !RAW_DATA[NOMBRE_HOJA]) return [];
  const rows = RAW_DATA[NOMBRE_HOJA].rows;

  // Un solo mapa de "variante -> canónica" por columna, calculado sobre
  // todas las filas, para que los filtros no muestren duplicados casi
  // idénticos (ver limpiarEspacios / construirMapaCanonico más arriba).
  const mapaComision = construirMapaCanonico(rows.map((r) => r[COL.comision]));
  const mapaBloque = construirMapaCanonico(rows.map((r) => r[COL.bloque]));
  const mapaTipo = construirMapaCanonico(rows.map((r) => r[COL.tipo]));
  const mapaAutor = construirMapaCanonico(rows.map((r) => r[COL.autor]));

  return rows.map((row) => ({
    expediente: String(row[COL.expediente] || "(sin número)"),
    sumario: String(row[COL.sumario] || ""),
    comision: canonicalizar(mapaComision, row[COL.comision], "Sin comisión"),
    bloque: canonicalizar(mapaBloque, row[COL.bloque], "Sin bloque"),
    tipo: canonicalizar(mapaTipo, row[COL.tipo], "Sin tipo"),
    autor: canonicalizar(mapaAutor, row[COL.autor], "Sin autor/a"),
    fechaSancion: formatFecha(row[COL.fechaAprobacion]),
    porTablas: esPorTablas(row[COL.observaciones]),
    etapa: getEtapa(row),
    anio: getAnio(row),
    fechas: getFechasDelItem(row),
    row: row,
  }));
}

function fechaEnRango(fechasItem, desde, hasta) {
  if (!desde && !hasta) return true;
  if (!fechasItem.length) return false;
  const desdeDate = desde ? new Date(desde + "T00:00:00") : null;
  const hastaDate = hasta ? new Date(hasta + "T23:59:59") : null;
  return fechasItem.some((f) => {
    if (desdeDate && f < desdeDate) return false;
    if (hastaDate && f > hastaDate) return false;
    return true;
  });
}

// Todos los filtros se combinan con AND entre sí (comisión Y bloque Y tipo
// Y etapa Y fecha Y búsqueda), y cada filtro de tipo "pill" admite varios
// valores seleccionados a la vez (OR dentro del mismo filtro).
function getItemsFiltrados() {
  let items = getAllItems();
  if (FILTROS.comisiones.length > 0) items = items.filter((i) => FILTROS.comisiones.includes(i.comision));
  if (FILTROS.bloques.length > 0) items = items.filter((i) => FILTROS.bloques.includes(i.bloque));
  if (FILTROS.tipos.length > 0) items = items.filter((i) => FILTROS.tipos.includes(i.tipo));
  if (FILTROS.autores.length > 0) items = items.filter((i) => FILTROS.autores.includes(i.autor));
  if (FILTROS.fechasSancion.length > 0) items = items.filter((i) => i.fechaSancion && FILTROS.fechasSancion.includes(i.fechaSancion));
  if (FILTROS.soloTablas) items = items.filter((i) => i.etapa.id === "sancionado" && i.porTablas);
  if (FILTROS.soloConDespacho) items = items.filter((i) => i.etapa.id === "sancionado" && !i.porTablas);
  if (FILTROS.etapas.length > 0) items = items.filter((i) => FILTROS.etapas.includes(i.etapa.id));
  if (FILTROS.fechaDesde || FILTROS.fechaHasta) {
    items = items.filter((i) => fechaEnRango(i.fechas, FILTROS.fechaDesde, FILTROS.fechaHasta));
  }
  if (FILTROS.busqueda.trim() !== "") {
    const q = normaliza(FILTROS.busqueda);
    items = items.filter((i) =>
      normaliza(i.expediente).includes(q) ||
      normaliza(i.sumario).includes(q) ||
      normaliza(i.row[COL.autor]).includes(q)
    );
  }
  return items;
}

// Detecta si la columna Observaciones indica que se aprobó "sobre tablas".
// Busca la palabra suelta "tablas" (o la variante "tables") en cualquier
// parte del texto, sin importar mayúsculas/minúsculas ni tildes, y aunque
// la celda tenga más contenido alrededor (ej: "Sancionado sobre TABLAS.",
// "s/tablas", "Aprobado sobre tablas en la sesión del 12/05").
function esPorTablas(observaciones) {
  const texto = normaliza(observaciones);
  if (!texto) return false;
  return /\btabl(?:as|es)\b/.test(texto);
}

// Fechas de sanción distintas que existen hoy en el Sheet, de la más
// reciente a la más antigua. Se recalcula con cada actualización de datos,
// así que las fechas nuevas aparecen solas en el desplegable.
function fechasSancionUnicas() {
  const mapa = new Map(); // "dd/mm/aaaa" -> timestamp (para ordenar)
  getAllItems().forEach((i) => {
    if (!i.fechaSancion || mapa.has(i.fechaSancion)) return;
    const f = parsearFecha(i.row[COL.fechaAprobacion]);
    if (f) mapa.set(i.fechaSancion, f.getTime());
  });
  return Array.from(mapa.entries()).sort((a, b) => b[1] - a[1]).map(([label]) => label);
}

function valoresUnicos(campo) {
  const set = new Set();
  getAllItems().forEach((i) => set.add(i[campo]));
  return Array.from(set).filter(Boolean).sort();
}

/* =========================================================================
   RENDER: overview stats
   ========================================================================= */

function renderOverview() {
  const items = getItemsFiltrados();
  const total = items.length;
  const conteoPorEtapa = {};
  ETAPAS.forEach((e) => { conteoPorEtapa[e.id] = items.filter((i) => i.etapa.id === e.id).length; });

  const porComision = {};
  items.forEach((i) => { porComision[i.comision] = (porComision[i.comision] || 0) + 1; });
  const comisionTop = Object.entries(porComision).sort((a, b) => b[1] - a[1])[0];

  const html = `
    <div class="stat-card">
      <div class="label">Total de expedientes</div>
      <div class="value" data-count="${total}">0</div>
      <div class="sub">con los filtros actuales</div>
    </div>
    <div class="stat-card">
      <div class="label">Sancionados</div>
      <div class="value" style="color:var(--olive)" data-count="${conteoPorEtapa.sancionado}">0</div>
      <div class="sub">aprobados en el recinto</div>
    </div>
    <div class="stat-card">
      <div class="label">Con despacho</div>
      <div class="value" style="color:var(--violet)" data-count="${conteoPorEtapa.despacho}">0</div>
      <div class="sub">esperando tratamiento en el recinto</div>
    </div>
    <div class="stat-card">
      <div class="label">Comisión con más actividad</div>
      <div class="value" style="font-size:19px; color:var(--ink);">${comisionTop ? comisionTop[0] : "—"}</div>
      <div class="sub">${comisionTop ? comisionTop[1] + " expedientes — ver el detalle abajo, en Gráficos" : ""}</div>
    </div>
  `;
  document.getElementById("overview-stats").innerHTML = html;
  document.querySelectorAll("#overview-stats .value[data-count]").forEach((el) => {
    animarNumero(el, Number(el.dataset.count));
  });

  renderGraficos(items);
}

/* =========================================================================
   RENDER: gráficos
   ========================================================================= */

function renderDonutHTML(data) {
  const total = data.reduce((s, d) => s + d.value, 0);
  let gradient = "var(--paper-dim)";
  if (total > 0) {
    let acumulado = 0;
    const stops = data.filter((d) => d.value > 0).map((d) => {
      const desde = (acumulado / total) * 100;
      acumulado += d.value;
      const hasta = (acumulado / total) * 100;
      return `${d.color} ${desde.toFixed(2)}% ${hasta.toFixed(2)}%`;
    });
    gradient = `conic-gradient(${stops.join(", ")})`;
  }
  const legend = data.map((d) => `
    <button type="button" class="legend-row ${d.active ? "active" : ""}" data-legend-id="${d.id}">
      <span class="legend-dot" style="background:${d.color}"></span>
      <span class="legend-label">${escapeHtml(d.label)}</span>
      <span class="legend-value">${d.value}</span>
    </button>
  `).join("");
  return `
    <div class="donut-wrap">
      <div class="donut-ring" style="background:${gradient}">
        <div class="donut-center">
          <div class="donut-total">${total}</div>
          <div class="donut-total-label">TOTAL</div>
        </div>
      </div>
      <div class="donut-legend">${legend || `<div class="empty-state">Sin datos.</div>`}</div>
    </div>`;
}

function renderBarrasHTML(data) {
  const max = Math.max(1, ...data.map((d) => d.value));
  const cols = data.map((d) => `
    <button type="button" class="bar-col ${d.active ? "active" : ""}" data-bar-id="${escapeHtml(d.id)}" title="${escapeHtml(d.label)}: ${d.value}">
      <span class="bar-val">${d.value}</span>
      <span class="bar-fill-wrap"><span class="bar-fill" style="height:${Math.round((d.value / max) * 100)}%"></span></span>
      <span class="bar-lbl">${escapeHtml(d.label)}</span>
    </button>
  `).join("");
  return `<div class="bar-chart">${cols}</div>`;
}

function renderRankingHTML(data) {
  const max = Math.max(1, ...data.map((d) => d.value));
  const filas = data.map((d) => `
    <button type="button" class="hbar-row ${d.active ? "active" : ""}" data-hbar-id="${escapeHtml(d.id).replace(/"/g, "&quot;")}" data-hbar-group="${d.group}">
      <span class="hbar-label">${escapeHtml(d.label)}</span>
      <span class="hbar-track"><span class="hbar-fill" style="width:${Math.round((d.value / max) * 100)}%"></span></span>
      <span class="hbar-value">${d.value}</span>
    </button>
  `).join("");
  return `<div class="hbar-list">${filas}</div>`;
}

function renderGraficos(items) {
  const cont = document.getElementById("charts-zone");
  if (!cont) return;

  const porEtapa = ETAPAS.map((e) => ({
    id: e.id, label: e.nombre, value: items.filter((i) => i.etapa.id === e.id).length,
    color: e.color, active: FILTROS.etapas.includes(e.id),
  }));

  const anios = Array.from(new Set(items.map((i) => i.anio))).filter(Boolean).sort((a, b) => {
    if (a === SIN_FECHA) return 1; if (b === SIN_FECHA) return -1; return Number(a) - Number(b);
  });
  const porAnio = anios.map((a) => ({
    id: a, label: a === SIN_FECHA ? "S/F" : a, value: items.filter((i) => i.anio === a).length, active: false,
  }));

  const porComisionMap = {};
  items.forEach((i) => { porComisionMap[i.comision] = (porComisionMap[i.comision] || 0) + 1; });
  const porComision = Object.entries(porComisionMap).sort((a, b) => b[1] - a[1])
    .map(([c, count]) => ({ id: c, group: "comision", label: c, value: count, active: FILTROS.comisiones.includes(c) }));

  const porBloqueMap = {};
  items.forEach((i) => { porBloqueMap[i.bloque] = (porBloqueMap[i.bloque] || 0) + 1; });
  const porBloque = Object.entries(porBloqueMap).sort((a, b) => b[1] - a[1])
    .map(([b, count]) => ({ id: b, group: "bloque", label: b, value: count, active: FILTROS.bloques.includes(b) }));

  const porAutorMap = {};
  items.forEach((i) => { porAutorMap[i.autor] = (porAutorMap[i.autor] || 0) + 1; });
  const porAutor = Object.entries(porAutorMap).sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], "es"))
    .map(([a, count]) => ({ id: a, group: "autor", label: a, value: count, active: FILTROS.autores.includes(a) }));

  cont.innerHTML = `
    <div class="viz-card">
      <div class="viz-card-title">Distribución por etapa <span class="viz-hint">(clic para filtrar)</span></div>
      <div class="viz-card-body">${renderDonutHTML(porEtapa)}</div>
    </div>
    <div class="viz-card">
      <div class="viz-card-title">Expedientes por año de ingreso</div>
      <div class="viz-card-body">${porAnio.length ? renderBarrasHTML(porAnio) : `<div class="empty-state">Sin datos de años con estos filtros.</div>`}</div>
    </div>
    <div class="viz-card">
      <div class="viz-card-title">Comisiones con más actividad <span class="viz-hint">(clic para filtrar)</span></div>
      <div class="viz-card-body">${porComision.length ? renderRankingHTML(porComision) : `<div class="empty-state">Sin datos con estos filtros.</div>`}</div>
    </div>
    <div class="viz-card">
      <div class="viz-card-title">Bloques con más actividad <span class="viz-hint">(clic para filtrar)</span></div>
      <div class="viz-card-body">${porBloque.length ? renderRankingHTML(porBloque) : `<div class="empty-state">Sin datos con estos filtros.</div>`}</div>
    </div>
    <div class="viz-card viz-card-wide">
      <div class="viz-card-title">Autores/as con más proyectos <span class="viz-hint">(según los filtros aplicados · clic para filtrar)</span></div>
      <div class="viz-card-body" style="align-items:flex-start;">${porAutor.length ? `<div style="width:100%; max-height:340px; overflow-y:auto;">${renderRankingHTML(porAutor)}</div>` : `<div class="empty-state">Sin datos con estos filtros.</div>`}</div>
    </div>
  `;

  cont.querySelectorAll("[data-legend-id]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const id = btn.dataset.legendId;
      const idx = FILTROS.etapas.indexOf(id);
      if (idx === -1) FILTROS.etapas.push(id); else FILTROS.etapas.splice(idx, 1);
      renderAll();
    });
  });
  cont.querySelectorAll("[data-hbar-id]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const val = btn.dataset.hbarId;
      const group = btn.dataset.hbarGroup;
      const lista = group === "comision" ? FILTROS.comisiones : group === "autor" ? FILTROS.autores : FILTROS.bloques;
      const idx = lista.indexOf(val);
      if (idx === -1) lista.push(val); else lista.splice(idx, 1);
      renderAll();
    });
  });
}

/* =========================================================================
   RENDER: filtros
   ========================================================================= */

function labelMultiselect(seleccionados, todosLabel) {
  const n = seleccionados.length;
  if (n === 0) return todosLabel;
  if (n === 1) return seleccionados[0];
  return `${n} seleccionados`;
}

function htmlMultiselect(idPrefix, valores, seleccionados, todosLabel, sinDatosLabel) {
  return `
    <div class="multiselect" id="ms-${idPrefix}">
      <button type="button" class="multiselect-toggle" id="ms-${idPrefix}-toggle" data-todos-label="${escapeHtml(todosLabel)}">
        <span>${escapeHtml(labelMultiselect(seleccionados, todosLabel))}</span>
        <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M6 9l6 6 6-6"/></svg>
      </button>
      <div class="multiselect-menu" id="ms-${idPrefix}-menu">
        ${valores.length ? valores.map((v) => `
          <label class="multiselect-option">
            <input type="checkbox" value="${escapeHtml(v)}" ${seleccionados.includes(v) ? "checked" : ""}>
            <span>${escapeHtml(v)}</span>
          </label>
        `).join("") : `<div class="empty-state" style="padding:14px 8px;">${escapeHtml(sinDatosLabel)}</div>`}
      </div>
    </div>
  `;
}

function renderFiltros() {
  const comisiones = valoresUnicos("comision");
  const bloques = valoresUnicos("bloque");
  const tipos = valoresUnicos("tipo");
  const autores = valoresUnicos("autor");
  const fechasSancion = fechasSancionUnicas();

  const html = `
    <div class="filter-group" style="grid-column: span 2;">
      <label>Buscar (expediente, sumario o autor/a)</label>
      <input type="text" id="f-busqueda" placeholder="Escribí un término…" value="${FILTROS.busqueda.replace(/"/g, "&quot;")}">
    </div>
    <div class="filter-group" style="grid-column: span 2;">
      <label>Rango de fechas <span style="text-transform:none; font-weight:400; opacity:0.7;">(ingreso, asesores, diputados o aprobación)</span></label>
      <div class="date-range">
        <input type="date" id="f-fecha-desde" value="${FILTROS.fechaDesde}">
        <span>a</span>
        <input type="date" id="f-fecha-hasta" value="${FILTROS.fechaHasta}">
      </div>
    </div>
    <div class="filter-group">
      <label>Comisión</label>
      ${htmlMultiselect("comisiones", comisiones, FILTROS.comisiones, "Todas las comisiones", "Sin comisiones cargadas.")}
    </div>
    <div class="filter-group">
      <label>Autor/a</label>
      ${htmlMultiselect("autores", autores, FILTROS.autores, "Todos los autores/as", "Sin autores/as cargados.")}
    </div>
    <div class="filter-group" style="grid-column: span 2;">
      <label>Bloque <span style="text-transform:none; font-weight:400; opacity:0.7;">(elegí uno o varios)</span></label>
      <div class="filter-pills" id="f-bloques">
        ${bloques.map((b) => `<span class="pill ${FILTROS.bloques.includes(b) ? "active" : ""}" data-bloque="${escapeHtml(b)}">${escapeHtml(b)}</span>`).join("")}
      </div>
    </div>
    <div class="filter-group">
      <label>Tipo de proyecto <span style="text-transform:none; font-weight:400; opacity:0.7;">(elegí uno o varios)</span></label>
      <div class="filter-pills" id="f-tipos">
        ${tipos.map((t) => `<span class="pill ${FILTROS.tipos.includes(t) ? "active" : ""}" data-tipo="${escapeHtml(t)}">${escapeHtml(t)}</span>`).join("")}
      </div>
    </div>
    <div class="filter-group" style="grid-column: span 2;">
      <label>Etapa <span style="text-transform:none; font-weight:400; opacity:0.7;">(elegí una o varias)</span></label>
      <div class="filter-pills" id="f-etapas">
        ${ETAPAS.map((e) => {
          const activa = FILTROS.etapas.includes(e.id);
          return `<span class="pill ${activa ? "active" : ""}" data-etapa="${e.id}" style="${activa ? `background:${e.color};border-color:${e.color};` : ""}">${e.nombre}</span>`;
        }).join("")}
      </div>
    </div>
    <div class="filter-group">
      <label>Fecha de sanción</label>
      ${htmlMultiselect("fechas-sancion", fechasSancion, FILTROS.fechasSancion, "Todas las fechas", "Todavía no hay proyectos sancionados.")}
    </div>
    <div class="filter-group">
      <label>Sancionados</label>
      <div class="filter-pills">
        <span class="pill ${FILTROS.soloTablas ? "active" : ""}" id="f-tablas">Tablas</span>
        <span class="pill ${FILTROS.soloConDespacho ? "active" : ""}" id="f-con-despacho">Sancionado con despacho</span>
      </div>
    </div>
    <div class="filter-group" style="justify-content:flex-end;">
      <button class="clear-filters" id="f-clear">Limpiar filtros</button>
    </div>
  `;
  document.getElementById("filters-zone").innerHTML = html;

  document.getElementById("f-busqueda").addEventListener("input", (e) => {
    FILTROS.busqueda = e.target.value;
    renderOverview(); renderKanban(); renderTabla();
  });
  document.getElementById("f-fecha-desde").addEventListener("change", (e) => {
    FILTROS.fechaDesde = e.target.value;
    renderAll();
  });
  document.getElementById("f-fecha-hasta").addEventListener("change", (e) => {
    FILTROS.fechaHasta = e.target.value;
    renderAll();
  });
  attachMultiselect("comisiones", (seleccion) => { FILTROS.comisiones = seleccion; });
  attachMultiselect("autores", (seleccion) => { FILTROS.autores = seleccion; });
  attachMultiselect("fechas-sancion", (seleccion) => { FILTROS.fechasSancion = seleccion; });
  // "Tablas" y "Sancionado con despacho" son opuestos entre sí (uno es
  // justo lo contrario del otro), así que al prender uno se apaga el otro.
  document.getElementById("f-tablas").addEventListener("click", () => {
    FILTROS.soloTablas = !FILTROS.soloTablas;
    if (FILTROS.soloTablas) FILTROS.soloConDespacho = false;
    renderAll();
  });
  document.getElementById("f-con-despacho").addEventListener("click", () => {
    FILTROS.soloConDespacho = !FILTROS.soloConDespacho;
    if (FILTROS.soloConDespacho) FILTROS.soloTablas = false;
    renderAll();
  });
  document.querySelectorAll("#f-bloques .pill").forEach((pill) => {
    pill.addEventListener("click", () => togglePill(FILTROS.bloques, pill.dataset.bloque));
  });
  document.querySelectorAll("#f-tipos .pill").forEach((pill) => {
    pill.addEventListener("click", () => togglePill(FILTROS.tipos, pill.dataset.tipo));
  });
  document.querySelectorAll("#f-etapas .pill").forEach((pill) => {
    pill.addEventListener("click", () => togglePill(FILTROS.etapas, pill.dataset.etapa));
  });
  document.getElementById("f-clear").addEventListener("click", () => {
    FILTROS = { comisiones: [], bloques: [], tipos: [], autores: [], fechasSancion: [], soloTablas: false, soloConDespacho: false, etapas: [], fechaDesde: "", fechaHasta: "", busqueda: "" };
    renderAll();
  });
}

// Los desplegables multi-selección no se re-renderizan enteros en cada
// click (a diferencia de los filtros "pill"), para que el menú quede
// abierto mientras se van tildando varias opciones seguidas.
function attachMultiselect(idPrefix, onChange) {
  const wrap = document.getElementById(`ms-${idPrefix}`);
  const toggle = document.getElementById(`ms-${idPrefix}-toggle`);
  const menu = document.getElementById(`ms-${idPrefix}-menu`);
  if (!wrap || !toggle || !menu) return;

  toggle.addEventListener("click", (e) => {
    e.stopPropagation();
    wrap.classList.toggle("open");
  });

  menu.querySelectorAll('input[type="checkbox"]').forEach((cb) => {
    cb.addEventListener("change", () => {
      const seleccion = Array.from(menu.querySelectorAll('input[type="checkbox"]:checked')).map((c) => c.value);
      onChange(seleccion);
      const labelEl = toggle.querySelector("span");
      const todosLabel = toggle.dataset.todosLabel || "Todos";
      if (labelEl) labelEl.textContent = labelMultiselect(seleccion, todosLabel);
      renderOverview(); renderKanban(); renderTabla();
    });
  });
}

// Cierra cualquier desplegable abierto si se clickea afuera de él.
document.addEventListener("click", (e) => {
  document.querySelectorAll(".multiselect.open").forEach((el) => {
    if (!el.contains(e.target)) el.classList.remove("open");
  });
});

function togglePill(lista, valor) {
  const idx = lista.indexOf(valor);
  if (idx === -1) lista.push(valor); else lista.splice(idx, 1);
  renderAll();
}

/* =========================================================================
   RENDER: Kanban (de solo lectura)
   ========================================================================= */

function renderKanban() {
  const items = getItemsFiltrados();
  const board = document.getElementById("kanban-board");

  board.innerHTML = ETAPAS.map((etapa) => {
    const itemsEtapa = items.filter((i) => i.etapa.id === etapa.id);
    return `
      <div class="kcol state-${etapa.id}" data-etapa-id="${etapa.id}" style="--kcolor:${etapa.color}">
        <div class="kcol-head">
          <span class="title">${etapa.nombre}</span>
          <span class="count">${itemsEtapa.length}</span>
        </div>
        <div class="kcol-body">
          ${itemsEtapa.length ? itemsEtapa.map((item) => renderCard(item)).join("") : `<div class="empty-state">Sin expedientes</div>`}
        </div>
      </div>
    `;
  }).join("");

  board.querySelectorAll(".kcard").forEach((card) => {
    card.addEventListener("click", () => abrirModal(card.dataset.row));
  });
}

function renderCard(item) {
  // Dato extra según la etapa: si tiene despacho, el número; si está
  // sancionado, la fecha en que se aprobó en el recinto.
  let tagExtra = "";
  if (item.etapa.id === "despacho" || item.etapa.id === "sancionado") {
    const numDespacho = String(item.row[COL.numeroDespacho] || "").trim();
    if (numDespacho) tagExtra += `<span class="tag despacho-tag">Despacho N.º ${escapeHtml(numDespacho)}</span>`;
  }
  if (item.etapa.id === "sancionado") {
    const fechaSancion = formatFecha(item.row[COL.fechaAprobacion]);
    if (fechaSancion) tagExtra += `<span class="tag sancion-tag">Sancionado ${escapeHtml(fechaSancion)}</span>`;
  }

  return `
    <div class="kcard" data-row="${escapeHtml(item.row._rowNumber)}" title="Clic para ver el resumen completo">
      <span class="exp">${escapeHtml(item.expediente)}</span>
      <span class="title">${escapeHtml(item.sumario || "(sin sumario)")}</span>
      <div class="meta">
        <span class="tag com-tag">${escapeHtml(item.comision)}</span>
        <span class="tag">${escapeHtml(item.bloque)}</span>
        ${tagExtra}
      </div>
    </div>
  `;
}

/* =========================================================================
   MODAL: detalle del expediente (incluye el Resumen completo)
   ========================================================================= */

function abrirModal(rowNumber) {
  const item = getAllItems().find((i) => String(i.row._rowNumber) === String(rowNumber));
  if (!item) return;

  const overlay = document.getElementById("modal-overlay");
  const panel = document.getElementById("modal-panel");

  const resumen = String(item.row[COL.resumen] || "").trim();
  const observaciones = String(item.row[COL.observaciones] || "").trim();
  const otrosGiros = String(item.row[COL.otrosGiros] || "").trim();

  let infoExtra = "";
  const numDespacho = String(item.row[COL.numeroDespacho] || "").trim();
  if (numDespacho) infoExtra += `<span class="tag despacho-tag">Despacho N.º ${escapeHtml(numDespacho)}</span>`;
  const fechaSancion = formatFecha(item.row[COL.fechaAprobacion]);
  if (item.etapa.id === "sancionado" && fechaSancion) infoExtra += `<span class="tag sancion-tag">Sancionado ${escapeHtml(fechaSancion)}</span>`;

  panel.innerHTML = `
    <button type="button" class="modal-close" id="modal-close-btn" aria-label="Cerrar">&times;</button>
    <div class="modal-header">
      <span class="modal-exp">${escapeHtml(item.expediente)}</span>
      ${statusChip(item.etapa)}
    </div>
    <h3 class="modal-sumario">${escapeHtml(item.sumario || "(sin sumario)")}</h3>
    <div class="modal-tags">
      <span class="tag com-tag">${escapeHtml(item.comision)}</span>
      <span class="tag">${escapeHtml(item.bloque)}</span>
      <span class="tag">${escapeHtml(item.tipo)}</span>
      ${infoExtra}
    </div>

    <div class="modal-section">
      <div class="modal-label">Resumen</div>
      <div class="modal-text">${resumen ? escapeHtml(resumen).replace(/\n/g, "<br>") : `<span class="modal-empty">Todavía no se cargó un resumen para este expediente en el Sheet.</span>`}</div>
    </div>

    <div class="modal-grid">
      <div><div class="modal-label">Autor/a</div><div class="modal-text">${escapeHtml(item.row[COL.autor] || "—")}</div></div>
      <div><div class="modal-label">Coautor/es</div><div class="modal-text">${escapeHtml(item.row[COL.coautores] || "—")}</div></div>
      <div><div class="modal-label">Fecha de ingreso</div><div class="modal-text">${escapeHtml(formatFecha(item.row[COL.fechaIngreso]) || "—")}</div></div>
      <div><div class="modal-label">¿Quienes firman dictamen?</div><div class="modal-text">${escapeHtml(item.row[COL.firmanDictamen] || "—")}</div></div>
      <div><div class="modal-label">¿Quienes no firman dictamen?</div><div class="modal-text">${escapeHtml(item.row[COL.noFirmanDictamen] || "—")}</div></div>
    </div>

    ${otrosGiros ? `
      <div class="modal-section">
        <div class="modal-label">¿Tiene otros giros?</div>
        <div class="modal-text">${escapeHtml(otrosGiros).replace(/\n/g, "<br>")}</div>
      </div>
    ` : ""}

    ${observaciones ? `
      <div class="modal-section">
        <div class="modal-label">Observaciones</div>
        <div class="modal-text">${escapeHtml(observaciones).replace(/\n/g, "<br>")}</div>
      </div>
    ` : ""}
  `;

  document.getElementById("modal-close-btn").addEventListener("click", cerrarModal);
  overlay.classList.add("show");
  document.body.style.overflow = "hidden";
}

function cerrarModal() {
  const overlay = document.getElementById("modal-overlay");
  overlay.classList.remove("show");
  document.body.style.overflow = "";
}

document.addEventListener("click", (e) => {
  if (e.target.id === "modal-overlay") cerrarModal();
});
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") cerrarModal();
});

/* =========================================================================
   RENDER: tabla de detalle
   ========================================================================= */

function renderTabla() {
  const items = getItemsFiltrados();
  const head = document.getElementById("table-head");
  const body = document.getElementById("table-body");
  const headers = RAW_DATA && RAW_DATA[NOMBRE_HOJA] ? RAW_DATA[NOMBRE_HOJA].headers : [];

  head.innerHTML = "<tr>" + headers.map((h) => `<th>${escapeHtml(h)}</th>`).join("") + "<th>Etapa</th></tr>";
  body.innerHTML = items.map((item) => {
    return `<tr class="clickable-row" data-row="${escapeHtml(item.row._rowNumber)}">` + headers.map((h) => {
      const cls = h === COL.sumario ? ' class="proyecto-cell"' : "";
      return `<td${cls}>${escapeHtml(item.row[h] ?? "")}</td>`;
    }).join("") + `<td>${statusChip(item.etapa)}</td></tr>`;
  }).join("") || `<tr><td colspan="${headers.length + 1}"><div class="empty-state">No hay resultados con estos filtros.</div></td></tr>`;

  body.querySelectorAll(".clickable-row").forEach((tr) => {
    tr.addEventListener("click", () => abrirModal(tr.dataset.row));
  });
}

function statusChip(etapa) {
  return `<span class="status-chip state-${etapa.id}">${etapa.nombre}</span>`;
}

/* =========================================================================
   EXPORTAR: Excel y PDF
   ========================================================================= */

function renderExportBar() {
  const el = document.getElementById("export-bar");
  if (!el) return;
  el.innerHTML = `
    <label class="export-checkbox">
      <input type="checkbox" id="exp-filtrado" ${EXPORTAR_SOLO_FILTRADO ? "checked" : ""}>
      Exportar solo lo filtrado (si no, exporta todo)
    </label>
    <button class="export-btn excel" id="exp-excel">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/></svg>
      Exportar a Excel
    </button>
    <button class="export-btn pdf" id="exp-pdf">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/></svg>
      Exportar a PDF
    </button>
  `;
  document.getElementById("exp-filtrado").addEventListener("change", (e) => {
    EXPORTAR_SOLO_FILTRADO = e.target.checked;
  });
  document.getElementById("exp-excel").addEventListener("click", () => exportarExcel());
  document.getElementById("exp-pdf").addEventListener("click", () => exportarPDF());
}

function getDatosExport() {
  const items = EXPORTAR_SOLO_FILTRADO ? getItemsFiltrados() : getAllItems();
  const headers = RAW_DATA && RAW_DATA[NOMBRE_HOJA] ? RAW_DATA[NOMBRE_HOJA].headers : [];
  const rows = items.map((item) => [...headers.map((h) => item.row[h] ?? ""), item.etapa.nombre]);
  return { headers: [...headers, "Etapa"], rows };
}

function descargarExcel(headers, rows, nombreArchivo) {
  if (typeof XLSX === "undefined") {
    showToast("No se pudo cargar la librería de Excel. Revisá tu conexión a internet.", true);
    return;
  }
  const data = [headers, ...rows];
  const ws = XLSX.utils.aoa_to_sheet(data);
  ws["!cols"] = headers.map(() => ({ wch: 26 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Datos");
  XLSX.writeFile(wb, nombreArchivo);
  showToast("Excel descargado ✓");
}

function exportarExcel() {
  const { headers, rows } = getDatosExport();
  if (!rows.length) { showToast("No hay datos para exportar con estos filtros.", true); return; }
  descargarExcel(headers, rows, "seguimiento_comisiones.xlsx");
}

function imprimirComoPDF(titulo, headers, rows) {
  const printArea = document.getElementById("print-area");
  const fecha = new Date().toLocaleString("es-AR");
  printArea.innerHTML = `
    <h1>${escapeHtml(titulo)}</h1>
    <div class="print-meta">Generado el ${fecha} · ${rows.length} resultado${rows.length === 1 ? "" : "s"}</div>
    <table>
      <thead><tr>${headers.map((h) => `<th>${escapeHtml(h)}</th>`).join("")}</tr></thead>
      <tbody>${rows.map((r) => `<tr>${r.map((c) => `<td>${escapeHtml(c)}</td>`).join("")}</tr>`).join("")}</tbody>
    </table>
  `;
  window.print();
}

function exportarPDF() {
  const { headers, rows } = getDatosExport();
  if (!rows.length) { showToast("No hay datos para exportar con estos filtros.", true); return; }
  imprimirComoPDF("Seguimiento Integral de Comisiones", headers, rows);
}

/* =========================================================================
   RENDER ALL
   ========================================================================= */

function renderAll() {
  renderOverview();
  renderFiltros();
  renderKanban();
  renderTabla();
  renderExportBar();
}

function render() {
  renderAll();
}

/* =========================================================================
   INIT
   ========================================================================= */

// Le pedimos al navegador que NO restaure el scroll por su cuenta (algunos
// navegadores lo hacen automáticamente al reabrir una pestaña o volver con
// "atrás"), y forzamos que la página arranque siempre arriba de todo.
if ("scrollRestoration" in history) {
  history.scrollRestoration = "manual";
}
window.scrollTo(0, 0);

// Si el navegador restaura esta página desde su caché de "atrás/adelante"
// (bfcache) en vez de cargarla de nuevo de cero, FILTROS seguiría teniendo
// el valor que tenía en ese momento. Lo reiniciamos a mano para que, sin
// importar cómo se vuelva a esta pestaña, siempre se vea "en cero": sin
// filtros aplicados y arriba de todo.
window.addEventListener("pageshow", (event) => {
  if (event.persisted) {
    FILTROS = { comisiones: [], bloques: [], tipos: [], autores: [], fechasSancion: [], soloTablas: false, soloConDespacho: false, etapas: [], fechaDesde: "", fechaHasta: "", busqueda: "" };
    window.scrollTo(0, 0);
    if (RAW_DATA) renderAll();
  }
});

cargarDatos(false);
reiniciarAutoRefresh();
