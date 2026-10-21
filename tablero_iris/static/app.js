"use strict";

/* Sin animación de entrada: además de ser más ágil para un panel que se
   recarga seguido, evita un problema real al exportar a PDF/imprimir:
   Chart.js anima el primer dibujo con requestAnimationFrame, y en el modo
   de impresión de Chromium esos frames pueden no llegar a completarse,
   dejando los gráficos a medio dibujar (barras invisibles, líneas cortadas)
   en el PDF resultante. Sin animación, cada gráfico queda completo desde
   el primer pintado. */
if (typeof Chart !== "undefined") Chart.defaults.animation = false;

/* ---------------------------------------------------------------- estado --- */
const els = {
  cargando: document.getElementById("cargando"),
  errorCarga: document.getElementById("error-carga"),
  errorCargaDetalle: document.getElementById("error-carga-detalle"),
  btnReintentar: document.getElementById("btn-reintentar"),
  vacio: document.getElementById("vacio"),
  sinResultados: document.getElementById("sin-resultados"),
  tablero: document.getElementById("tablero"),
  modoCampanas: document.getElementById("modo-campanas"),
  modoGenerico: document.getElementById("modo-generico"),
  estadoArchivo: document.getElementById("estado-archivo"),
  estadoPunto: document.getElementById("estado-punto"),
  aviso: document.getElementById("aviso"),
  inputArchivo: document.getElementById("input-archivo"),
  btnRecargar: document.getElementById("btn-recargar"),
  btnExportar: document.getElementById("btn-exportar"),
  btnDescargarPdf: document.getElementById("btn-descargar-pdf"),
  btnTema: document.getElementById("btn-tema"),
  iconoTema: document.getElementById("icono-tema"),
  btnLimpiarFiltros: document.getElementById("btn-limpiar-filtros"),
  btnMenu: document.getElementById("btn-menu"),
  sidebar: document.getElementById("sidebar"),
  fondoSidebar: document.getElementById("fondo-sidebar"),
  fBuscar: document.getElementById("f-buscar"),
  fTipo: document.getElementById("f-tipo"),
  sidebarFiltros: document.getElementById("sidebar-filtros"),
  kpis: document.getElementById("seccion-resumen"),
  statsResumen: document.getElementById("seccion-resumen-stats"),
  lecturaRapida: document.getElementById("lectura-rapida"),
  selSemanalMetrica: document.getElementById("sel-semanal-metrica"),
  selSemanalAgrupacion: document.getElementById("sel-semanal-agrupacion"),
  tituloSemanal: document.getElementById("titulo-semanal"),
  listaEfectividadEntidad: document.getElementById("lista-efectividad-entidad"),
  listaMejores: document.getElementById("lista-mejores"),
  listaPeores: document.getElementById("lista-peores"),
  notaTabla: document.getElementById("nota-tabla"),
  tablaCampanas: document.getElementById("tabla-campanas"),
  paginacion: document.getElementById("paginacion"),
  selMetrica: document.getElementById("sel-metrica"),
  selAgrupar: document.getElementById("sel-agrupar"),
  kpisGenerico: document.getElementById("kpis-generico"),
  notaTablaGenerico: document.getElementById("nota-tabla-generico"),
  tablaGenerica: document.getElementById("tabla-generica"),
  encabezadoImpresion: document.getElementById("encabezado-impresion"),
};

const graficos = {};
let marcaConocida = null;
let ultimaBusquedaTimeout = null;
let ultimoDashboard = null;

/* Estado de la tabla de campañas: orden y paginación (todo en el cliente,
   los datos ya vienen filtrados del servidor). */
const TABLA_FILAS_POR_PAGINA = 12;
const tablaEstado = { columna: "fecha_evento", direccion: "desc", pagina: 1 };

const FORMATO_NUM = new Intl.NumberFormat("es-CO");
const FORMATO_MONEDA = new Intl.NumberFormat("es-CO", { style: "currency", currency: "COP", maximumFractionDigits: 0 });

/* Paleta institucional Alcaldía de Bogotá: el rojo manda, apoyado por
   blanco y amarillo; el teal queda como único apoyo neutro (semáforo de
   efectividad) y el azul casi no se usa. Debe coincidir con las variables
   CSS en :root (static/style.css). */
const COLOR = {
  violeta: "#e4032e", violetaOsc: "#a10f1f", // rojo institucional (uso digital), color de marca principal
  ambar: "#f9d248", ambarOsc: "#d9a61e", // amarillo institucional
  teal: "#1fac8f", tealOsc: "#0d8a72",
  rosa: "#e3272a", rosaOsc: "#b3021f", // rojo institucional del escudo, tono de apoyo
  grafito: "#4a4a56", grafitoOsc: "#26262e",
  tinta: "#736d8a",
};

/* ------------------------------------------------------------- utilidades --- */
function mostrarAviso(texto, esError = false) {
  els.aviso.textContent = texto;
  els.aviso.classList.toggle("error", esError);
  els.aviso.classList.add("mostrar");
  clearTimeout(mostrarAviso._t);
  mostrarAviso._t = setTimeout(() => els.aviso.classList.remove("mostrar"), 3800);
}

function num(v) { return v === null || v === undefined ? "—" : FORMATO_NUM.format(v); }
function moneda(v) { return v === null || v === undefined ? "—" : FORMATO_MONEDA.format(v); }
function pct(v) { return v === null || v === undefined ? "—" : `${FORMATO_NUM.format(v)} %`; }
function fechaCorta(v) { return v ? new Date(v + "T00:00:00").toLocaleDateString("es-CO") : "—"; }

function claseEfectividad(v) {
  if (v === null || v === undefined) return "";
  if (v >= 85) return "alto";
  if (v >= 60) return "medio";
  return "bajo";
}

function llenarSelect(select, opciones, valorActual) {
  select.innerHTML = "";
  for (const op of opciones) {
    const o = document.createElement("option");
    o.value = op;
    o.textContent = op;
    if (op === valorActual) o.selected = true;
    select.appendChild(o);
  }
}

/* --------------------------------------------- filtros de selección múltiple --- */
/* Entidad, Mes y Año se eligen con checkboxes (en vez de un <select> de una
   sola opción) para poder combinar varios a la vez -p.ej. dos meses, o tres
   entidades-, algo que un <select> normal no permite. `msEstado` guarda, por
   cada filtro, el conjunto de valores marcados. */
const MULTISELECTS = ["entidad", "mes", "anio"];
const msEstado = { entidad: new Set(), mes: new Set(), anio: new Set() };
let multiselectAbierto = null;

function _msElementos(clave) {
  return {
    raiz: document.getElementById(`ms-${clave}`),
    boton: document.getElementById(`ms-${clave}`).querySelector(".multiselect-boton"),
    panel: document.getElementById(`ms-${clave}-panel`),
  };
}

function _sinTildes(txt) {
  // Para que buscar "secretaria" también encuentre "Secretaría" (y
  // similares): quita tildes/diacríticos antes de comparar.
  return txt.normalize("NFD").replace(/[̀-ͯ]/g, "");
}

function _msEtiqueta(clave) {
  const { raiz } = _msElementos(clave);
  const sel = msEstado[clave];
  const vacia = raiz.dataset.etiquetaVacia;
  if (sel.size === 0) return vacia;
  if (sel.size === 1) return [...sel][0];
  return `${sel.size} seleccionadas`;
}

function _msActualizarEtiqueta(clave) {
  const { boton } = _msElementos(clave);
  boton.querySelector(".multiselect-etiqueta").textContent = _msEtiqueta(clave);
}

/* Reconstruye las opciones (checkboxes) de un filtro con la lista vigente
   -los valores disponibles pueden cambiar al recargar los datos-,
   conservando la selección previa en los valores que sigan existiendo. */
function actualizarMultiselect(clave, opciones) {
  const { raiz, panel } = _msElementos(clave);
  const vacia = raiz.dataset.etiquetaVacia;
  const valores = opciones.filter((o) => o !== vacia);
  msEstado[clave] = new Set([...msEstado[clave]].filter((v) => valores.includes(v)));

  const cont = panel.querySelector(".multiselect-opciones");
  cont.innerHTML = "";
  if (!valores.length) {
    const vacio = document.createElement("p");
    vacio.className = "multiselect-sin-opciones";
    vacio.textContent = "No hay opciones disponibles.";
    cont.appendChild(vacio);
  }
  for (const valor of valores) {
    const etiqueta = document.createElement("label");
    etiqueta.className = "multiselect-opcion";
    etiqueta.dataset.texto = _sinTildes(valor.toLowerCase());
    const input = document.createElement("input");
    input.type = "checkbox";
    input.value = valor;
    input.checked = msEstado[clave].has(valor);
    input.addEventListener("change", () => {
      if (input.checked) msEstado[clave].add(valor);
      else msEstado[clave].delete(valor);
      _msActualizarEtiqueta(clave);
      cargarTablero();
    });
    const span = document.createElement("span");
    span.textContent = valor;
    span.title = valor;
    etiqueta.append(input, span);
    cont.appendChild(etiqueta);
  }
  _msActualizarEtiqueta(clave);
}

function _msPosicionar(clave) {
  const { boton, panel } = _msElementos(clave);
  const r = boton.getBoundingClientRect();
  // El botón mide lo que el sidebar permite (~220px), pero nombres como los
  // de las entidades distritales no caben ahí: el panel se abre más ancho
  // que el botón (hasta 360px) para que el texto se lea completo, sin
  // salirse de la pantalla.
  const anchoMin = Math.max(r.width, 230);
  const anchoDeseado = clave === "entidad" ? 360 : anchoMin;
  const ancho = Math.min(anchoDeseado, window.innerWidth - 16);
  panel.style.left = `${Math.max(8, Math.min(r.left, window.innerWidth - ancho - 8))}px`;
  panel.style.width = `${ancho}px`;
  const alturaEstim = Math.min(panel.scrollHeight || 320, 320);
  if (r.bottom + alturaEstim + 10 > window.innerHeight) {
    panel.style.top = `${Math.max(8, r.top - alturaEstim - 6)}px`;
  } else {
    panel.style.top = `${r.bottom + 6}px`;
  }
}

function cerrarMultiselects() {
  if (!multiselectAbierto) return;
  const { raiz, boton, panel } = _msElementos(multiselectAbierto);
  panel.hidden = true;
  raiz.classList.remove("abierto");
  boton.setAttribute("aria-expanded", "false");
  multiselectAbierto = null;
}

function abrirMultiselect(clave) {
  if (multiselectAbierto === clave) { cerrarMultiselects(); return; }
  cerrarMultiselects();
  const { raiz, boton, panel } = _msElementos(clave);
  panel.hidden = false;
  raiz.classList.add("abierto");
  boton.setAttribute("aria-expanded", "true");
  _msPosicionar(clave);
  multiselectAbierto = clave;
  panel.querySelector(".multiselect-buscador input")?.focus();
}

function _inicializarMultiselects() {
  // Los paneles se cuelgan directo de <body> (position: fixed calculado por
  // JS) para que nunca queden recortados por el overflow-y del sidebar ni,
  // en móvil, por el transform que usa el sidebar para deslizarse como
  // overlay (un transform crea un nuevo contenedor para position:fixed).
  for (const clave of MULTISELECTS) {
    const { panel } = _msElementos(clave);
    document.body.appendChild(panel);
    panel.addEventListener("click", (e) => e.stopPropagation());

    const { boton } = _msElementos(clave);
    boton.addEventListener("click", (e) => { e.stopPropagation(); abrirMultiselect(clave); });

    panel.querySelector('[data-accion="todas"]')?.addEventListener("click", () => {
      const casillas = [...panel.querySelectorAll('input[type="checkbox"]')];
      casillas.forEach((c) => { c.checked = true; });
      msEstado[clave] = new Set(casillas.map((c) => c.value));
      _msActualizarEtiqueta(clave);
      cargarTablero();
    });
    panel.querySelector('[data-accion="ninguna"]')?.addEventListener("click", () => {
      panel.querySelectorAll('input[type="checkbox"]').forEach((c) => { c.checked = false; });
      msEstado[clave] = new Set();
      _msActualizarEtiqueta(clave);
      cargarTablero();
    });
    const buscador = panel.querySelector(".multiselect-buscador input");
    buscador?.addEventListener("input", () => {
      const q = _sinTildes(buscador.value.trim().toLowerCase());
      panel.querySelectorAll(".multiselect-opcion").forEach((op) => {
        op.classList.toggle("oculta", Boolean(q) && !op.dataset.texto.includes(q));
      });
    });
  }
  document.addEventListener("click", cerrarMultiselects);
  window.addEventListener("resize", () => { if (multiselectAbierto) _msPosicionar(multiselectAbierto); });
  document.addEventListener("keydown", (e) => { if (e.key === "Escape") cerrarMultiselects(); });
}
_inicializarMultiselects();

function construirQuery() {
  const p = new URLSearchParams();
  if (els.fTipo.value && els.fTipo.value !== "Todos") p.set("tipo", els.fTipo.value);
  for (const v of msEstado.entidad) p.append("entidad", v);
  for (const v of msEstado.mes) p.append("mes", v);
  for (const v of msEstado.anio) p.append("anio", v);
  if (els.fBuscar.value.trim()) p.set("buscar", els.fBuscar.value.trim());
  return p;
}

function destruirGrafico(id) {
  if (graficos[id]) { graficos[id].destroy(); delete graficos[id]; }
}

/* -------------------------------------------------------------- modo oscuro --- */
function aplicarTema(tema) {
  if (tema === "oscuro") {
    document.documentElement.setAttribute("data-tema", "oscuro");
    els.iconoTema.className = "fa-solid fa-sun";
  } else {
    document.documentElement.removeAttribute("data-tema");
    els.iconoTema.className = "fa-solid fa-moon";
  }
  try { localStorage.setItem("iris-tema", tema); } catch (e) { /* almacenamiento no disponible */ }
  // Los gráficos de Chart.js fijan sus colores al crearse: hay que
  // reconstruirlos para que tomen los tonos del nuevo tema.
  if (ultimoDashboard) {
    if (ultimoDashboard.modo === "campanas") renderizarCampanas(ultimoDashboard);
    else renderizarGenerico(ultimoDashboard);
  }
}

els.btnTema?.addEventListener("click", () => {
  const actual = document.documentElement.getAttribute("data-tema") === "oscuro" ? "oscuro" : "claro";
  aplicarTema(actual === "oscuro" ? "claro" : "oscuro");
});

// Sincroniza el icono del botón con el tema ya aplicado por el script inline del <head>.
aplicarTemaInicialIcono();
function aplicarTemaInicialIcono() {
  const activo = document.documentElement.getAttribute("data-tema") === "oscuro";
  if (els.iconoTema) els.iconoTema.className = activo ? "fa-solid fa-sun" : "fa-solid fa-moon";
}

/* ------------------------------------------------------------------ carga --- */
async function cargarTablero() {
  const query = construirQuery();
  const esPrimeraCarga = els.tablero.hidden && els.vacio.hidden && els.sinResultados.hidden;
  if (esPrimeraCarga) {
    els.cargando.hidden = false;
    els.errorCarga.hidden = true;
  }

  let datos;
  try {
    const resp = await fetch(`/api/dashboard?${query.toString()}`);
    if (resp.status === 401) {
      window.location.href = "/login";
      return;
    }
    if (!resp.ok) throw new Error(`El servidor respondió con error ${resp.status}.`);
    datos = await resp.json();
    marcarConexion(true);
  } catch (e) {
    marcarConexion(false);
    if (esPrimeraCarga) {
      els.cargando.hidden = true;
      els.errorCarga.hidden = false;
      els.errorCargaDetalle.textContent = e.message || "No se pudo contactar con el servidor.";
    } else {
      mostrarAviso("No se pudo contactar al servidor.", true);
    }
    return;
  }
  els.cargando.hidden = true;
  els.errorCarga.hidden = true;

  if (!datos.hay_datos) {
    els.vacio.hidden = false;
    els.sinResultados.hidden = true;
    els.tablero.hidden = true;
    els.sidebarFiltros.hidden = true;
    els.estadoArchivo.textContent = "Todavía no hay datos";
    return;
  }
  els.vacio.hidden = true;

  if (datos.sin_resultados) {
    els.sinResultados.hidden = false;
    els.tablero.hidden = true;
    els.sidebarFiltros.hidden = false;
    actualizarFiltrosDisponibles(datos.filtros_disponibles);
    els.estadoArchivo.textContent = `${datos.meta.nombre} · sin coincidencias`;
    return;
  }
  els.sinResultados.hidden = true;
  els.tablero.hidden = false;

  els.estadoArchivo.textContent = `${datos.meta.nombre} · actualizado ${new Date(datos.meta.actualizado).toLocaleString("es-CO")}`;

  ultimoDashboard = datos;
  tablaEstado.pagina = 1;

  if (datos.modo === "campanas") {
    els.modoCampanas.hidden = false;
    els.modoGenerico.hidden = true;
    els.sidebarFiltros.hidden = false;
    renderizarCampanas(datos);
  } else {
    els.modoCampanas.hidden = true;
    els.modoGenerico.hidden = false;
    els.sidebarFiltros.hidden = true;
    renderizarGenerico(datos);
  }
}

function marcarConexion(ok) {
  els.estadoPunto?.classList.toggle("punto-error", !ok);
}

/* -------------------------------------------------------- modo campañas --- */
function actualizarFiltrosDisponibles(f) {
  if (!f) return;
  llenarSelect(els.fTipo, f.tipos, els.fTipo.value || "Todos");
  actualizarMultiselect("entidad", f.entidades);
  actualizarMultiselect("mes", f.meses);
  actualizarMultiselect("anio", f.anios);
}

let datosSemanales = null;

function renderizarCampanas(datos) {
  actualizarFiltrosDisponibles(datos.filtros_disponibles);
  datosSemanales = datos.evolucion_periodos || null;

  const pasos = [
    () => renderizarKPIs(datos.kpis),
    () => renderizarGraficoEvolucion(datos.evolucion),
    () => renderizarGraficoCanal(datos.comparativo_canal),
    () => renderizarGraficoResultado(datos.resultado_envios),
    () => renderizarGraficoSemanal(),
    () => renderizarGraficoEntidades(datos.top_entidades),
    () => renderizarEfectividadEntidad(datos.efectividad_entidades),
    () => renderizarLecturaRapida(datos),
    () => renderizarRankings(datos.ranking),
    () => renderizarTablaCampanas(datos.tabla),
    () => renderizarEncabezadoImpresion(datos),
  ];
  for (const paso of pasos) {
    try { paso(); } catch (e) { console.error(e); }
  }
  actualizarImagenesDeImpresionPronto();
}

function renderizarKPIs(k) {
  const tarjetas = [
    { icono: "fa-bullhorn", bg: "bg-violeta", etiqueta: "Campañas", valor: num(k.total_campanas), nota: `${num(k.total_envios)} envíos registrados` },
    { icono: "fa-users", bg: "bg-teal", etiqueta: "Destinatarios alcanzados", valor: num(k.total_destinatarios), nota: `${num(k.total_efectivos)} mensajes efectivos` },
    { icono: "fa-bolt", bg: "bg-ambar", etiqueta: "Efectividad promedio", valor: pct(k.efectividad_promedio), nota: "sobre envíos con resultado reportado" },
    { icono: "fa-sack-dollar", bg: "bg-grafito", etiqueta: "Inversión total", valor: moneda(k.inversion_total), nota: "suma de campañas filtradas" },
    { icono: "fa-tag", bg: "bg-rosa", etiqueta: "Costo por destinatario", valor: k.costo_por_destinatario !== null ? moneda(k.costo_por_destinatario) : "—", nota: "inversión ÷ destinatarios" },
  ];
  els.kpis.innerHTML = tarjetas.map(t => `
    <div class="kpi">
      <span class="kpi-icono ${t.bg}"><i class="fa-solid ${t.icono}"></i></span>
      <div>
        <p class="etiqueta">${t.etiqueta}</p>
        <p class="valor">${t.valor}</p>
        <p class="nota">${t.nota}</p>
      </div>
    </div>`).join("");
  renderizarStatsResumen(k);
}

function renderizarStatsResumen(k) {
  if (!els.statsResumen) return;
  const chips = [
    { icono: "fa-comment-sms", etiqueta: "Envíos por SMS", valor: num(k.envios_sms) },
    { icono: "fa-envelope", etiqueta: "Envíos por mailing", valor: num(k.envios_mailing) },
    { icono: "fa-building-columns", etiqueta: "Entidades solicitantes", valor: num(k.entidades_solicitantes) },
    { icono: "fa-calendar-check", etiqueta: "Último envío medido", valor: fechaCorta(k.ultimo_envio) },
  ];
  els.statsResumen.innerHTML = chips.map(c => `
    <div class="stat-chip">
      <span class="stat-chip-icono"><i class="fa-solid ${c.icono}"></i></span>
      <div>
        <p class="etiqueta">${c.etiqueta}</p>
        <p class="valor">${c.valor}</p>
      </div>
    </div>`).join("");
}

function esTemaOscuro() { return document.documentElement.getAttribute("data-tema") === "oscuro"; }

/* Estas tarjetas ya no tienen fondo de color (ver .tarjeta-color en
   style.css): ahora son blancas con un borde superior de acento, así que
   sus gráficos se dibujan siempre con texto oscuro sobre blanco, en
   pantalla, al imprimir y al exportar. */
function temaGraficoColor() {
  return { color: "#2b2540", grid: "#e7e3f5", eje: "#4b4460" };
}
function temaGraficoBlanco() {
  const osc = esTemaOscuro();
  return {
    texto: osc ? "#cdc7de" : "#2b2540",
    textoSuave: osc ? "#9992b3" : "#736d8a",
    grid: osc ? "#2f2850" : "#efedf7",
  };
}

function renderizarGraficoEvolucion(evolucion) {
  destruirGrafico("evolucion");
  const ctx = document.getElementById("g-evolucion");
  if (!evolucion || evolucion.length === 0) return;
  const T = temaGraficoColor();
  graficos.evolucion = new Chart(ctx, {
    data: {
      labels: evolucion.map(e => e.periodo),
      datasets: [
        {
          type: "bar", label: "Campañas", data: evolucion.map(e => e.campanas),
          backgroundColor: "rgba(228,3,46,.18)",
          borderRadius: 5, yAxisID: "y", order: 2,
        },
        {
          type: "line", label: "Efectividad %", data: evolucion.map(e => e.efectividad),
          borderColor: COLOR.violeta,
          backgroundColor: COLOR.violeta,
          pointBackgroundColor: COLOR.ambar, pointBorderColor: COLOR.violeta,
          borderWidth: 2.5, tension: .35, spanGaps: true, yAxisID: "y1", order: 1,
        },
      ],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { mode: "index", intersect: false },
      scales: {
        y: { beginAtZero: true, ticks: { color: T.eje }, grid: { color: T.grid } },
        y1: { beginAtZero: true, max: 100, position: "right", ticks: { color: T.eje }, grid: { display: false } },
        x: { ticks: { color: T.eje }, grid: { display: false } },
      },
      plugins: { legend: { position: "bottom", labels: { color: T.color, boxWidth: 12, usePointStyle: true } } },
    },
  });
}

function renderizarGraficoCanal(comparativo) {
  destruirGrafico("canal");
  const ctx = document.getElementById("g-canal");
  if (!comparativo || comparativo.length === 0) return;
  const T = temaGraficoColor();
  graficos.canal = new Chart(ctx, {
    type: "bar",
    data: {
      labels: comparativo.map(c => c.tipo),
      datasets: [
        { label: "Campañas", data: comparativo.map(c => c.campanas), backgroundColor: "rgba(217,166,30,.3)", borderRadius: 6 },
        { label: "Efectividad %", data: comparativo.map(c => c.efectividad), backgroundColor: COLOR.ambarOsc, borderRadius: 6 },
      ],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      scales: {
        y: { beginAtZero: true, ticks: { color: T.eje }, grid: { color: T.grid } },
        x: { ticks: { color: T.color, font: { weight: "600" } }, grid: { display: false } },
      },
      plugins: { legend: { position: "bottom", labels: { color: T.color, boxWidth: 12, usePointStyle: true } } },
    },
  });
}

function renderizarGraficoResultado(r) {
  destruirGrafico("resultado");
  const ctx = document.getElementById("g-resultado");
  if (!r) return;
  const T = temaGraficoColor();
  const datos = [
    { l: "Efectivos", v: r.efectivos },
    { l: "No entregados", v: r.no_entregados },
    { l: "Rechazados", v: r.rechazados },
    { l: "Expirados", v: r.expirados },
    { l: "Inválidos", v: r.invalidos },
  ].filter(d => d.v);
  const tonos = [COLOR.tealOsc, "#4fc3ab", "#8ad9c5", "#bfeadf", "#e3f6f0"];
  graficos.resultado = new Chart(ctx, {
    type: "doughnut",
    data: {
      labels: datos.map(d => d.l),
      datasets: [{ data: datos.map(d => d.v), backgroundColor: tonos, borderWidth: 2, borderColor: "#fff" }],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position: "bottom", labels: { color: T.color, boxWidth: 12, usePointStyle: true, font: { size: 11 } } } },
    },
  });
}

function renderizarGraficoEntidades(top) {
  destruirGrafico("entidades");
  const ctx = document.getElementById("g-entidades");
  if (!top || top.length === 0) return;
  const T = temaGraficoBlanco();
  const ordenado = [...top].reverse();
  graficos.entidades = new Chart(ctx, {
    type: "bar",
    data: {
      labels: ordenado.map(e => e.entidad.length > 30 ? e.entidad.slice(0, 28) + "…" : e.entidad),
      datasets: [{ label: "Campañas", data: ordenado.map(e => e.campanas), backgroundColor: COLOR.violeta, borderRadius: 5, barThickness: 14 }],
    },
    options: {
      indexAxis: "y",
      responsive: true, maintainAspectRatio: false,
      scales: {
        x: { beginAtZero: true, grid: { color: T.grid }, ticks: { color: T.textoSuave } },
        y: { grid: { display: false }, ticks: { color: T.texto, font: { size: 11.5 } } },
      },
      plugins: { legend: { display: false } },
    },
  });
}

/* El canal se dibuja en este orden para que, en el área apilada, SMS quede
   encima de Mailing (coincide con cómo se ve: la zona de Mailing es más
   amplia y queda de fondo, y SMS resalta por encima). */
const ORDEN_CANAL_SEMANAL = ["Mailing", "SMS"];
const COLORES_CANAL_SEMANAL = {
  SMS: { linea: COLOR.violeta, relleno: "rgba(228, 3, 46, .3)" },
  Mailing: { linea: COLOR.tinta, relleno: "rgba(115, 109, 138, .32)" },
};
const ETIQUETAS_METRICA_SEMANAL = { efectivos: "Mensajes efectivos", envios: "Envíos", valor: "Valor" };

function renderizarGraficoSemanal() {
  destruirGrafico("semanal");
  const ctx = document.getElementById("g-semanal");
  if (!datosSemanales || !els.selSemanalAgrupacion || !els.selSemanalMetrica) return;

  const agrupacion = els.selSemanalAgrupacion.value || "semana";
  const metrica = els.selSemanalMetrica.value || "efectivos";
  const bloque = datosSemanales[agrupacion] || { etiquetas: [], series: {} };
  if (els.tituloSemanal) {
    const porTexto = agrupacion === "mes" ? "por mes" : "por semana";
    els.tituloSemanal.textContent = `${ETIQUETAS_METRICA_SEMANAL[metrica] || metrica} ${porTexto}`;
  }
  if (!bloque.etiquetas.length) return;

  const T = temaGraficoBlanco();
  const formatear = (v) => (metrica === "valor" ? moneda(v) : num(v));

  const datasets = ORDEN_CANAL_SEMANAL
    .filter((canal) => bloque.series[canal])
    .map((canal) => {
      const c = COLORES_CANAL_SEMANAL[canal];
      return {
        label: canal,
        data: bloque.series[canal][metrica] || [],
        borderColor: c.linea,
        backgroundColor: c.relleno,
        fill: true,
        tension: .4,
        pointRadius: 0,
        pointHoverRadius: 4,
        borderWidth: 2,
      };
    });

  graficos.semanal = new Chart(ctx, {
    type: "line",
    data: { labels: bloque.etiquetas, datasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { mode: "index", intersect: false },
      plugins: {
        legend: { position: "bottom", labels: { color: T.texto, boxWidth: 10, usePointStyle: true, padding: 14 } },
        tooltip: { callbacks: { label: (c) => `${c.dataset.label}: ${formatear(c.parsed.y)}` } },
      },
      scales: {
        x: {
          grid: { display: false },
          ticks: { color: T.textoSuave, autoSkip: true, maxTicksLimit: 9 },
        },
        y: {
          beginAtZero: true,
          grid: { color: T.grid },
          ticks: { color: T.textoSuave, callback: (v) => formatear(v) },
        },
      },
    },
  });
}

function renderizarEfectividadEntidad(lista) {
  if (!els.listaEfectividadEntidad) return;
  if (!lista || !lista.length) {
    els.listaEfectividadEntidad.innerHTML = "<li>Sin datos suficientes todavía.</li>";
    return;
  }
  els.listaEfectividadEntidad.innerHTML = lista.map((e) => {
    const clase = claseEfectividad(e.efectividad);
    const ancho = e.efectividad === null || e.efectividad === undefined ? 0 : Math.max(0, Math.min(100, e.efectividad));
    const etiquetaEnvios = e.envios === 1 ? "envío" : "envíos";
    return `
      <li>
        <div class="item-efectividad-cabecera">
          <span class="item-efectividad-nombre">${e.entidad} <span class="item-efectividad-envios">(${num(e.envios)} ${etiquetaEnvios})</span></span>
          <span class="item-efectividad-pct ${clase}">${pct(e.efectividad)}</span>
        </div>
        <div class="barra-efectividad"><div class="barra-efectividad-relleno ${clase}" style="width:${ancho}%"></div></div>
      </li>`;
  }).join("");
}

function renderizarLecturaRapida(datos) {
  const k = datos.kpis;
  const mejorCanal = [...datos.comparativo_canal].sort((a, b) => (b.efectividad || 0) - (a.efectividad || 0))[0];
  const topEntidad = datos.top_entidades[0];
  const mejorMes = [...datos.evolucion].filter(e => e.efectividad !== null).sort((a, b) => b.efectividad - a.efectividad)[0];

  const puntos = [];
  if (mejorCanal) puntos.push(`<b>${mejorCanal.tipo}</b> tiene la mayor efectividad entre canales, con ${pct(mejorCanal.efectividad)}.`);
  if (topEntidad) puntos.push(`<b>${topEntidad.entidad}</b> es la entidad con más campañas (${num(topEntidad.campanas)}).`);
  if (mejorMes) puntos.push(`<b>${mejorMes.periodo}</b> fue el mes con mejor efectividad (${pct(mejorMes.efectividad)}).`);
  if (k.costo_por_destinatario !== null) puntos.push(`Cada destinatario alcanzado cuesta en promedio <b>${moneda(k.costo_por_destinatario)}</b>.`);

  els.lecturaRapida.innerHTML = puntos.map(p => `<li><i class="fa-solid fa-circle-check"></i><span>${p}</span></li>`).join("") || "<li>Sin datos suficientes todavía.</li>";
}

function itemRanking(f, i) {
  return `
    <li>
      <span class="puesto">${i + 1}</span>
      <span class="info">
        <span class="nombre" title="${f.campana}">${f.campana || "Sin nombre"}</span>
        <span class="meta">${f.entidad || "—"} · ${f.tipo} · ${num(f.destinatarios)} destinatarios</span>
      </span>
      <span class="pct ${claseEfectividad(f.efectividad)}">${pct(f.efectividad)}</span>
    </li>`;
}

function renderizarRankings(r) {
  els.listaMejores.innerHTML = (r.mejores || []).map(itemRanking).join("") || "<li>Sin datos suficientes.</li>";
  els.listaPeores.innerHTML = (r.peores || []).map(itemRanking).join("") || "<li>Sin datos suficientes.</li>";
}

/* ------------------------------------------------------ tabla: orden + paginación --- */
const COLUMNAS_TABLA = [
  { clave: "campana", etiqueta: "Campaña", tipo: "texto" },
  { clave: "entidad", etiqueta: "Entidad", tipo: "texto" },
  { clave: "tipo", etiqueta: "Canal", tipo: "texto" },
  { clave: "mes", etiqueta: "Mes", tipo: "texto" },
  { clave: "destinatarios", etiqueta: "Destinatarios", tipo: "numero" },
  { clave: "efectividad", etiqueta: "Efectividad %", tipo: "numero" },
  { clave: "valor", etiqueta: "Valor", tipo: "numero" },
  { clave: "fecha_evento", etiqueta: "Fecha", tipo: "fecha" },
];

function filasComoObjetos(t) {
  return t.filas.map(fila => {
    const o = {};
    t.columnas.forEach((c, i) => { o[c] = fila[i]; });
    return o;
  });
}

function ordenarFilas(filas, columna, direccion) {
  const def = COLUMNAS_TABLA.find(c => c.clave === columna);
  const signo = direccion === "asc" ? 1 : -1;
  return [...filas].sort((a, b) => {
    let va = a[columna], vb = b[columna];
    if (va === null || va === undefined) return 1;
    if (vb === null || vb === undefined) return -1;
    if (def && def.tipo === "texto") {
      va = String(va).toLowerCase(); vb = String(vb).toLowerCase();
      return va < vb ? -signo : va > vb ? signo : 0;
    }
    if (def && def.tipo === "fecha") {
      va = va ? new Date(va).getTime() : -Infinity;
      vb = vb ? new Date(vb).getTime() : -Infinity;
    }
    return (va - vb) * signo;
  });
}

function renderizarTablaCampanas(t) {
  const todasLasFilas = filasComoObjetos(t);
  const filasOrdenadas = ordenarFilas(todasLasFilas, tablaEstado.columna, tablaEstado.direccion);

  const totalPaginas = Math.max(1, Math.ceil(filasOrdenadas.length / TABLA_FILAS_POR_PAGINA));
  if (tablaEstado.pagina > totalPaginas) tablaEstado.pagina = totalPaginas;
  const inicio = (tablaEstado.pagina - 1) * TABLA_FILAS_POR_PAGINA;
  const filasPagina = filasOrdenadas.slice(inicio, inicio + TABLA_FILAS_POR_PAGINA);

  els.notaTabla.textContent = `${t.total_filas} campaña(s) — una fila de la base de datos equivale a una campaña.`;

  const encabezado = `<thead><tr>${COLUMNAS_TABLA.map(c => {
    const activa = c.clave === tablaEstado.columna;
    const flecha = activa ? (tablaEstado.direccion === "asc" ? "▲" : "▼") : "↕";
    return `<th data-col="${c.clave}" class="${activa ? "orden-activo" : ""}">${c.etiqueta}<span class="flecha-orden">${flecha}</span></th>`;
  }).join("")}</tr></thead>`;

  const cuerpo = filasPagina.map(f => {
    const chip = f.tipo === "SMS" ? "chip-sms" : "chip-mailing";
    return `<tr>
      <td>${f.campana || "<span class=\"sin-dato\">Sin nombre</span>"}</td>
      <td>${f.entidad || "—"}</td>
      <td><span class="chip ${chip}">${f.tipo}</span></td>
      <td>${f.mes || "—"}</td>
      <td>${num(f.destinatarios)}</td>
      <td>${f.efectividad === null ? '<span class="sin-dato">Pendiente</span>' : pct(f.efectividad)}</td>
      <td>${moneda(f.valor)}</td>
      <td>${f.fecha_evento || "—"}</td>
    </tr>`;
  }).join("");

  els.tablaCampanas.innerHTML = encabezado + `<tbody>${cuerpo}</tbody>`;

  els.tablaCampanas.querySelectorAll("thead th").forEach(th => {
    th.addEventListener("click", () => {
      const col = th.dataset.col;
      if (tablaEstado.columna === col) {
        tablaEstado.direccion = tablaEstado.direccion === "asc" ? "desc" : "asc";
      } else {
        tablaEstado.columna = col;
        tablaEstado.direccion = "asc";
      }
      tablaEstado.pagina = 1;
      renderizarTablaCampanas(t);
    });
  });

  renderizarPaginacion(totalPaginas, t);
}

function renderizarPaginacion(totalPaginas, t) {
  if (totalPaginas <= 1) { els.paginacion.innerHTML = ""; return; }
  const p = tablaEstado.pagina;
  const botones = [];

  botones.push(`<button class="pagina-btn" data-pagina="${p - 1}" ${p === 1 ? "disabled" : ""} aria-label="Página anterior"><i class="fa-solid fa-chevron-left"></i></button>`);

  const paginasAMostrar = new Set([1, totalPaginas, p, p - 1, p + 1]);
  let anterior = 0;
  for (let i = 1; i <= totalPaginas; i++) {
    if (!paginasAMostrar.has(i)) continue;
    if (i - anterior > 1) botones.push(`<span class="pagina-puntos">…</span>`);
    botones.push(`<button class="pagina-btn ${i === p ? "activa" : ""}" data-pagina="${i}">${i}</button>`);
    anterior = i;
  }

  botones.push(`<button class="pagina-btn" data-pagina="${p + 1}" ${p === totalPaginas ? "disabled" : ""} aria-label="Página siguiente"><i class="fa-solid fa-chevron-right"></i></button>`);

  els.paginacion.innerHTML = botones.join("");
  els.paginacion.querySelectorAll("[data-pagina]").forEach(btn => {
    btn.addEventListener("click", () => {
      const destino = parseInt(btn.dataset.pagina, 10);
      if (destino < 1 || destino > totalPaginas) return;
      tablaEstado.pagina = destino;
      renderizarTablaCampanas(t);
      document.getElementById("seccion-tabla").scrollIntoView({ behavior: "smooth", block: "start" });
    });
  });
}

/* ------------------------------------------------------------- impresión --- */
function renderizarEncabezadoImpresion(datos) {
  const filtros = [];
  if (els.fTipo.value && els.fTipo.value !== "Todos") filtros.push(`Canal: ${els.fTipo.value}`);
  if (msEstado.entidad.size) filtros.push(`Entidad: ${[...msEstado.entidad].join(", ")}`);
  if (msEstado.mes.size) filtros.push(`Mes: ${[...msEstado.mes].join(", ")}`);
  if (msEstado.anio.size) filtros.push(`Año: ${[...msEstado.anio].join(", ")}`);
  if (els.fBuscar.value.trim()) filtros.push(`Búsqueda: "${els.fBuscar.value.trim()}"`);

  els.encabezadoImpresion.classList.add("con-contenido");
  els.encabezadoImpresion.innerHTML = `
    <h1 style="font-family:Fraunces,serif;margin:0 0 .2rem;font-size:1.4rem;">Línea 195 · Panel de campañas</h1>
    <p style="margin:0;color:#555;font-size:.85rem;">
      Generado el ${new Date().toLocaleString("es-CO")} · Fuente: ${datos.meta.nombre}
      ${filtros.length ? " · Filtros: " + filtros.join(", ") : " · Sin filtros aplicados"}
    </p>`;
}

/* Chromium no siempre repinta bien el <canvas> de Chart.js al generar el
   PDF cuando la tarjeta cambia de tamaño (queda parcialmente en blanco).
   Por eso mantenemos, junto a cada gráfico, una imagen estática ya
   congelada -el CSS de impresión oculta el canvas y muestra la imagen en
   su lugar-, que sí se imprime siempre de forma fiable.
   En vez de generar esa imagen solo al vuelo justo antes de imprimir (lo
   que dependía de que el evento "beforeprint" llegara a tiempo y a veces
   dejaba algún gráfico en blanco), la actualizamos cada vez que un
   gráfico se (re)dibuja, así que para cuando el usuario pulsa "Imprimir"
   la imagen de respaldo ya está lista de antemano; "beforeprint" queda
   solo como último refresco de seguridad. */
function prepararGraficosParaImprimir() {
  Object.values(graficos).forEach((chart) => {
    try {
      const lienzo = chart.canvas?.closest(".lienzo");
      if (!lienzo) return;
      let img = lienzo.querySelector("img.imagen-impresion");
      if (!img) {
        img = document.createElement("img");
        img.className = "imagen-impresion";
        img.alt = "";
        lienzo.appendChild(img);
      }
      img.src = chart.toBase64Image("image/png", 1);
    } catch (e) { /* si un gráfico falla, se imprime sin su imagen de respaldo */ }
  });
}
function limpiarImagenesDeImpresion() {
  document.querySelectorAll(".lienzo img.imagen-impresion").forEach((img) => img.remove());
}
/* Doble requestAnimationFrame: espera a que el navegador realmente pinte
   el canvas (un solo rAF puede ejecutarse antes de ese pintado) antes de
   congelarlo como imagen. */
function actualizarImagenesDeImpresionPronto() {
  requestAnimationFrame(() => requestAnimationFrame(prepararGraficosParaImprimir));
}
window.addEventListener("beforeprint", prepararGraficosParaImprimir);
window.addEventListener("afterprint", limpiarImagenesDeImpresion);

/* --------------------------------------------------------- modo genérico --- */
function renderizarGenerico(datos) {
  llenarSelect(els.selMetrica, datos.columnas.numericas, datos.seleccion.metrica);
  llenarSelect(els.selAgrupar, datos.columnas.categoricas, datos.seleccion.agrupar);

  const c = datos.cifras;
  const iconos = [
    { icono: "fa-table-list", bg: "bg-violeta", etiqueta: "Filas", valor: num(c.filas) },
    c.total !== undefined ? { icono: "fa-sack-dollar", bg: "bg-teal", etiqueta: "Total", valor: num(c.total) } : null,
    c.promedio !== undefined ? { icono: "fa-chart-simple", bg: "bg-ambar", etiqueta: "Promedio", valor: num(Math.round(c.promedio * 100) / 100) } : null,
    c.maximo !== undefined ? { icono: "fa-arrow-up-wide-short", bg: "bg-rosa", etiqueta: "Máximo", valor: num(c.maximo) } : null,
  ].filter(Boolean);
  els.kpisGenerico.innerHTML = iconos.map(t => `
    <div class="kpi">
      <span class="kpi-icono ${t.bg}"><i class="fa-solid ${t.icono}"></i></span>
      <div><p class="etiqueta">${t.etiqueta}</p><p class="valor">${t.valor}</p></div>
    </div>`).join("");

  destruirGrafico("generico");
  const ctx = document.getElementById("g-generico");
  if (datos.por_categoria && datos.por_categoria.length) {
    const T = temaGraficoColor();
    graficos.generico = new Chart(ctx, {
      type: "bar",
      data: {
        labels: datos.por_categoria.map(p => p.etiqueta),
        datasets: [{ data: datos.por_categoria.map(p => p.valor), backgroundColor: "rgba(255,255,255,.6)", borderRadius: 6 }],
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        scales: {
          y: { beginAtZero: true, ticks: { color: T.eje }, grid: { color: T.grid } },
          x: { ticks: { color: T.color }, grid: { display: false } },
        },
        plugins: { legend: { display: false } },
      },
    });
  }

  els.notaTablaGenerico.textContent = `Mostrando ${datos.tabla.filas.length} de ${datos.tabla.total_filas} filas.`;
  const encabezado = `<thead><tr>${datos.tabla.columnas.map(c2 => `<th>${c2}</th>`).join("")}</tr></thead>`;
  const cuerpo = datos.tabla.filas.map(f => `<tr>${f.map(v => `<td>${v === null ? "" : v}</td>`).join("")}</tr>`).join("");
  els.tablaGenerica.innerHTML = encabezado + `<tbody>${cuerpo}</tbody>`;
  actualizarImagenesDeImpresionPronto();
}

/* ------------------------------------------------------------------ subida --- */
async function subirArchivo(archivo) {
  const form = new FormData();
  form.append("archivo", archivo);
  els.estadoArchivo.textContent = "Subiendo archivo…";
  try {
    const resp = await fetch("/api/subir", { method: "POST", body: form });
    if (resp.status === 401) { window.location.href = "/login"; return; }
    const datos = await resp.json();
    if (!resp.ok) throw new Error(datos.error || "Error al subir el archivo.");
    mostrarAviso("Datos actualizados correctamente.");
    marcaConocida = datos.marca;
    await cargarTablero();
  } catch (e) {
    mostrarAviso(e.message, true);
  }
}

/* -------------------------------------------------------- auto-actualización --- */
async function revisarCambios() {
  try {
    const resp = await fetch("/api/estado");
    if (resp.status === 401) return;
    const datos = await resp.json();
    if (marcaConocida !== null && datos.marca !== marcaConocida) {
      mostrarAviso("El archivo cambió en disco — actualizando…");
      await cargarTablero();
    }
    marcaConocida = datos.marca;
    marcarConexion(true);
  } catch (e) {
    marcarConexion(false);
  }
}

/* --------------------------------------------------------------- eventos --- */
els.inputArchivo?.addEventListener("change", (e) => {
  if (e.target.files[0]) subirArchivo(e.target.files[0]);
  e.target.value = "";
});

els.btnRecargar?.addEventListener("click", () => { cargarTablero(); mostrarAviso("Datos recargados."); });
els.btnReintentar?.addEventListener("click", () => cargarTablero());
els.selSemanalMetrica?.addEventListener("change", () => { renderizarGraficoSemanal(); actualizarImagenesDeImpresionPronto(); });
els.selSemanalAgrupacion?.addEventListener("change", () => { renderizarGraficoSemanal(); actualizarImagenesDeImpresionPronto(); });

// Los gráficos ya se dibujan siempre con texto oscuro sobre fondo blanco
// (ver temaGraficoColor), igual que se necesita en el Excel, así que para
// exportar basta con capturar el canvas tal cual se ve en pantalla.
async function _capturarGraficosParaExportar() {
  const imagenes = {};
  for (const clave of ["evolucion", "canal", "resultado", "entidades", "semanal"]) {
    if (graficos[clave]) {
      try { imagenes[clave] = graficos[clave].toBase64Image("image/png", 1.0); } catch (e) { /* si un gráfico no se puede capturar, se omite */ }
    }
  }
  return imagenes;
}

async function _descargarInforme(ruta, nombreArchivo, mensajeError) {
  const query = construirQuery();
  // Captura los gráficos con la paleta clara (fondo blanco, texto y líneas
  // oscuras) para que el informe los incruste nítidos, en vez del degradado
  // de color que se ve en pantalla.
  const imagenes = await _capturarGraficosParaExportar();
  try {
    const resp = await fetch(`${ruta}?${query.toString()}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ imagenes }),
    });
    if (resp.status === 401) { window.location.href = "/login"; return; }
    if (!resp.ok) {
      const datos = await resp.json().catch(() => ({}));
      throw new Error(datos.error || mensajeError);
    }
    const blob = await resp.blob();
    const url = URL.createObjectURL(blob);
    const enlace = document.createElement("a");
    enlace.href = url;
    enlace.download = nombreArchivo;
    document.body.appendChild(enlace);
    enlace.click();
    enlace.remove();
    URL.revokeObjectURL(url);
  } catch (e) {
    mostrarAviso(e.message, true);
  }
}

els.btnExportar?.addEventListener("click", () => _descargarInforme(
  "/api/exportar", "informe_campanas.xlsx", "No se pudo generar el informe."));

/* Solo descarga directa de PDF, sin opción de "Imprimir": el diálogo de
   impresión del navegador depende de que Chrome alcance a "congelar" cada
   gráfico a tiempo justo antes de imprimir, y a veces -de forma
   impredecible- deja alguno en blanco o sin colores. Este PDF, en cambio,
   se genera siempre en el servidor (con reportlab, a partir de los mismos
   datos y gráficos que se ven en pantalla), así que sale completo y con
   los mismos colores cada vez. */
els.btnDescargarPdf?.addEventListener("click", () => _descargarInforme(
  "/api/exportar-pdf", "informe_campanas.pdf", "No se pudo generar el PDF."));

els.btnLimpiarFiltros?.addEventListener("click", () => {
  els.fBuscar.value = "";
  els.fTipo.value = "Todos";
  for (const clave of MULTISELECTS) {
    msEstado[clave] = new Set();
    const { panel } = _msElementos(clave);
    panel.querySelectorAll('input[type="checkbox"]').forEach((c) => { c.checked = false; });
    _msActualizarEtiqueta(clave);
  }
  cargarTablero();
});

els.fTipo?.addEventListener("change", cargarTablero);
els.fBuscar?.addEventListener("input", () => {
  clearTimeout(ultimaBusquedaTimeout);
  ultimaBusquedaTimeout = setTimeout(cargarTablero, 350);
});

els.selMetrica?.addEventListener("change", cargarTablero);
els.selAgrupar?.addEventListener("change", cargarTablero);

/* ------------------------------------------------------------ sidebar móvil --- */
function abrirSidebar() { els.sidebar.classList.add("abierto"); }
function cerrarSidebar() { els.sidebar.classList.remove("abierto"); }
els.btnMenu?.addEventListener("click", abrirSidebar);
els.fondoSidebar?.addEventListener("click", cerrarSidebar);

const enlacesNav = document.querySelectorAll("[data-nav]");
enlacesNav.forEach(a => a.addEventListener("click", () => {
  enlacesNav.forEach(x => x.classList.remove("activo"));
  a.classList.add("activo");
  cerrarSidebar();
}));

if (enlacesNav.length) {
  const objetivos = [...enlacesNav].map(a => document.querySelector(a.getAttribute("href"))).filter(Boolean);
  const observador = new IntersectionObserver((entradas) => {
    const visible = entradas.find(e => e.isIntersecting);
    if (!visible) return;
    const idx = objetivos.indexOf(visible.target);
    if (idx === -1) return;
    enlacesNav.forEach(x => x.classList.remove("activo"));
    enlacesNav[idx].classList.add("activo");
  }, { rootMargin: "-45% 0px -50% 0px" });
  objetivos.forEach(o => observador.observe(o));
}

/* ------------------------------------------------------------------- drag&drop --- */
let dropZoneListo = false;
function habilitarArrastrar() {
  if (dropZoneListo) return;
  dropZoneListo = true;
  ["dragover", "drop"].forEach(ev => window.addEventListener(ev, (e) => e.preventDefault()));
  window.addEventListener("drop", (e) => {
    const f = e.dataTransfer.files[0];
    if (f && /\.(xlsx|xlsm)$/i.test(f.name)) subirArchivo(f);
  });
}
habilitarArrastrar();

/* ------------------------------------------------------------------- init --- */
cargarTablero();
setInterval(revisarCambios, 10000);
