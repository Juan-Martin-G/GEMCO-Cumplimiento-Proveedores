/*
 * Cumplimiento de proveedores · GEMCO — panel
 * Lee los datos precalculados por generar_datos.py (datos_cumplimiento.json).
 *
 * Estructura:
 *   CONFIG → parámetros de presentación (semáforo, formato)
 *   Util   → utilidades de meses y celdas
 *   Datos  → carga del JSON y adaptación al modelo que usa la interfaz
 *   UI     → vistas, gráficos y tablas
 */
'use strict';

const CONFIG = {
  // Semáforo de cumplimiento (fracciones: 0.8 = 80 %).
  //   cumplimiento < amarilloDesde              → rojo
  //   amarilloDesde ≤ cumplimiento < verdeDesde → amarillo
  //   cumplimiento ≥ verdeDesde                 → verde
  //   sin forecast                              → gris
  semaforo: { amarilloDesde: 0.80, verdeDesde: 1.00 },
  archivoDatos: 'datos_cumplimiento.json',
  locale: 'es-CL',
};

/* ===================================================================== */
/*  Util                                                                 */
/* ===================================================================== */

const Util = (() => {
  const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto',
    'septiembre', 'octubre', 'noviembre', 'diciembre'];
  const MESES_CORTOS = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
  // Desplazamientos relativos a la columna de etiquetas ("Item") de cada bloque.
  const LAYOUT = { meses: [1, 2, 3, 5, 6, 7, 9, 10, 11, 13, 14, 15], trimestres: [4, 8, 12, 16], total: 17 };

  function colLetra(c) {
    let s = '';
    for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
    return s;
  }
  const refHoja = h => (/^[A-Za-z0-9_]+$/.test(h) ? h : `'${h.replace(/'/g, "''")}'`);
  const mesCorto = (anio, mes) => `${MESES_CORTOS[mes]}-${String(anio).slice(2)}`;
  const soloDia = d => new Date(d.getFullYear(), d.getMonth(), d.getDate());
  // 'YYYY-MM-DD' → Date local (sin desfase de zona horaria)
  const fecha = s => {
    if (!s) return null;
    const [y, m, d] = s.split('-').map(Number);
    return new Date(y, m - 1, d);
  };

  return { MESES, MESES_CORTOS, LAYOUT, colLetra, refHoja, mesCorto, soloDia, fecha };
})();

/* ===================================================================== */
/*  Datos                                                                */
/* ===================================================================== */

const Datos = (() => {
  // Convierte el JSON al modelo que usa la interfaz (fechas como Date, filas de detalle por bloque).
  function adaptar(j) {
    const m = j.meta;
    const filaF = r => ({ ...r, fecha: Util.fecha(r.fecha) });
    const marcas = j.marcas.map(mc => {
      const facturas = mc.detalle ? mc.detalle.facturas.map(filaF) : [];
      const bo = mc.detalle ? mc.detalle.bo.map(filaF) : [];
      const bloques = mc.bloques.map(b => {
        if (b.calc) {
          const periodo = new Set(b.meses.map(x => `${x.anio}|${x.mes}`));
          const enPeriodo = r => periodo.has(`${r.anio}|${r.mes}`);
          b.calc.filasF = facturas.filter(enPeriodo);
          b.calc.filasB = bo.filter(enPeriodo);
        }
        return b;
      });
      return { ...mc, bloques, boTodas: bo };
    });
    return {
      meta: m,
      archivo: m.archivo,
      generado: m.generado ? new Date(m.generado) : null,
      incluirDetalle: !!m.incluirDetalle,
      hojas: m.hojas,
      tolerancia: m.tolerancia,
      categorias: m.categorias,
      sinCategoria: m.sinCategoria,
      fact: { nombre: m.fact.hoja, filas: m.fact.filas, corte: Util.fecha(m.fact.corte) },
      bo: { nombre: m.bo.hoja, filas: m.bo.filas, corte: Util.fecha(m.bo.corte) },
      corte: Util.fecha(m.corte),
      parametros: j.parametros,
      problemas: j.problemas,
      resumen: j.resumen,
      ocultas: j.ocultas,
      marcas,
    };
  }

  async function cargar() {
    let resp;
    try {
      resp = await fetch(CONFIG.archivoDatos, { cache: 'no-store' });
    } catch (e) {
      throw new Error(location.protocol === 'file:'
        ? 'Abriste el panel como archivo local: el navegador no permite leer el JSON así. Usa un servidor local (python -m http.server -d docs) o la página publicada.'
        : `No se pudo descargar ${CONFIG.archivoDatos} (${e.message}).`);
    }
    if (!resp.ok) throw new Error(`No se pudo descargar ${CONFIG.archivoDatos} (HTTP ${resp.status}). ¿Se corrió generar_datos.py y se hizo push?`);
    return adaptar(await resp.json());
  }

  return { cargar, adaptar };
})();


/* ===================================================================== */
/*  UI                                                                   */
/* ===================================================================== */

const UI = (() => {
  let modelo = null;
  const estado = {
    vista: 'resumen', marca: null, bloque: 0,
    soloForecast: false, orden: { k: 'ytd', dir: -1 },
    soloDif: false, boFuera: false,
  };
  const charts = {};

  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => [...r.querySelectorAll(s)];
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));

  /* ---------- Formatos es-CL ---------- */
  const nf = new Intl.NumberFormat(CONFIG.locale, { maximumFractionDigits: 2 });
  const nf0 = new Intl.NumberFormat(CONFIG.locale, { maximumFractionDigits: 0 });
  const nf1 = new Intl.NumberFormat(CONFIG.locale, { minimumFractionDigits: 1, maximumFractionDigits: 1 });
  const nfTC = new Intl.NumberFormat(CONFIG.locale, { maximumFractionDigits: 6 });
  const nfCompacto = new Intl.NumberFormat(CONFIG.locale, { notation: 'compact', maximumFractionDigits: 1 });
  const fmt = v => (v == null || !Number.isFinite(v) ? '—' : nf.format(v));
  const fmtPct = v => (v == null || !Number.isFinite(v) ? '—' : `${nf1.format(v * 100)}%`);
  const fmtTC = v => (v == null ? '—' : nfTC.format(v));
  const pad = n => String(n).padStart(2, '0');
  const fmtFecha = d => (d ? `${pad(d.getDate())}-${pad(d.getMonth() + 1)}-${d.getFullYear()}` : '—');
  const mesLargo = (m) => `${Util.MESES[m.mes]} ${m.anio}`;

  /* ---------- Semáforo ---------- */
  const ICONO = { verde: '✓', amarillo: '!', rojo: '✕', gris: '–' };
  function semaforo(v) {
    if (v == null || !Number.isFinite(v)) return 'gris';
    if (v < CONFIG.semaforo.amarilloDesde) return 'rojo';
    if (v < CONFIG.semaforo.verdeDesde) return 'amarillo';
    return 'verde';
  }
  const etiquetaSemaforo = e => ({
    verde: `≥ ${fmtPct(CONFIG.semaforo.verdeDesde)}`,
    amarillo: `${fmtPct(CONFIG.semaforo.amarilloDesde)} – ${fmtPct(CONFIG.semaforo.verdeDesde)}`,
    rojo: `< ${fmtPct(CONFIG.semaforo.amarilloDesde)}`,
    gris: 'Sin forecast',
  }[e]);
  const chip = (v, tip) => {
    const e = semaforo(v);
    const contenido = v == null ? '<span class="sr">Sin forecast</span><span aria-hidden="true">—</span>' : `<span aria-hidden="true">${ICONO[e]}</span>${fmtPct(v)}`;
    return `<span class="chip chip-${e}"${tip ? ` data-tip="${esc(tip)}" tabindex="0"` : ''}>${contenido}</span>`;
  };
  const info = tip => `<button type="button" class="info" data-tip="${esc(tip)}" aria-label="Origen del dato">i</button>`;

  const css = v => getComputedStyle(document.documentElement).getPropertyValue(v).trim();

  /* ---------- Textos de origen (tooltips) ---------- */
  function descPeriodo(b) {
    const ini = b.meses[0], fin = b.meses[11];
    return b.tipo === 'calendario' ? `Año=${ini.anio}` : `${Util.mesCorto(ini.anio, ini.mes)} a ${Util.mesCorto(fin.anio, fin.mes)} (pares Año/Mes de las filas 1–2)`;
  }
  function refTotal(marca, b, fila) {
    return `${Util.refHoja(marca.hoja)}!${Util.colLetra(b.c0 + Util.LAYOUT.total)}${fila}`;
  }
  function textoConversion(marca) {
    if (!marca.convierte) return '';
    const pares = marca.paresConversion.map(p => `${p.orig}>${p.dest} × ${fmtTC(p.tc)}`);
    return `\nIncluye documentos convertidos a ${marca.moneda} (${pares.join(', ')}, TC de "${modelo.hojas.moneda}"). Ver recuadro "Conversión de moneda".`;
  }
  function origen(marca, b, campo) {
    const p = descPeriodo(b);
    const conv = textoConversion(marca);
    const T = Util.colLetra(b.c0 + Util.LAYOUT.total);
    switch (campo) {
      case 'fact': return `Σ ${modelo.hojas.facturacion} col 'Total moneda marca' | Proveedor=${marca.nombre}, ${p} · equivale a ${refTotal(marca, b, 3)}${conv}`;
      case 'bo': return `Σ ${modelo.hojas.bo} col 'Total moneda marca' | Marca=${marca.nombre}, Año OC y Mes de la fecha OC, ${p} · equivale a ${refTotal(marca, b, 4)}\nSolo OC del período: BO de OC de otros años no entra, igual que en el Excel. Filas con Cantidad pdte = 0 suman 0.${conv}`;
      case 'factbo': return `Facturación + Backorder · equivale a ${refTotal(marca, b, 5)}${conv}`;
      case 'fc': return `Celda Total del Forecast (fila 6, input manual) · ${refTotal(marca, b, 6)}`;
      case 'cumpl': return `Fact+BO total / Forecast total · equivale a ${refTotal(marca, b, 7)} = IFERROR(${T}5/${T}6,0)\nCompara lo acumulado contra los 12 meses de forecast.`;
      case 'ytd': {
        const y = b.calc.ytd;
        if (!y.hasta) return 'Calculada por la app: el período aún no comienza a la fecha de corte.';
        return `Calculada por la app (no existe en el Excel).\nFact+BO de ${mesLargo(y.desde)} a ${mesLargo(y.hasta)} / Σ Forecast mensual del mismo rango.\nIncluye el mes de corte completo (corte ${fmtFecha(modelo.corte)}), aunque el mes no haya terminado.`;
      }
      default: return '';
    }
  }

  /* ---------- Tooltip propio (hover, foco y toque) ---------- */
  function initTooltips() {
    const tip = $('#tooltip');
    let actual = null;
    const mostrar = el => {
      actual = el;
      tip.textContent = el.dataset.tip;
      tip.hidden = false;
      const r = el.getBoundingClientRect();
      const w = tip.offsetWidth, h = tip.offsetHeight;
      let x = Math.min(Math.max(8, r.left + r.width / 2 - w / 2), window.innerWidth - w - 8);
      let y = r.bottom + 8;
      if (y + h > window.innerHeight - 8) y = Math.max(8, r.top - h - 8);
      tip.style.left = `${x}px`;
      tip.style.top = `${y}px`;
    };
    const ocultar = () => { actual = null; tip.hidden = true; };
    document.addEventListener('mouseover', e => {
      const el = e.target.closest('[data-tip]');
      if (el && el !== actual) mostrar(el);
      else if (!el && actual && !actual.matches(':focus')) ocultar();
    });
    document.addEventListener('focusin', e => { const el = e.target.closest('[data-tip]'); if (el) mostrar(el); });
    document.addEventListener('focusout', ocultar);
    document.addEventListener('click', e => {
      const el = e.target.closest('[data-tip]');
      if (!el) ocultar();
      else if (el.matches('.info')) { e.preventDefault(); if (actual === el && !tip.hidden) ocultar(); else mostrar(el); }
    });
    window.addEventListener('scroll', ocultar, { passive: true });
    document.addEventListener('keydown', e => { if (e.key === 'Escape') ocultar(); });
  }

  /* ---------- Tema claro / oscuro ---------- */
  function temaActual() {
    const t = document.documentElement.dataset.theme;
    if (t) return t;
    return window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
  }
  function initTema() {
    $('#btn-tema').addEventListener('click', () => {
      const nuevo = temaActual() === 'dark' ? 'light' : 'dark';
      document.documentElement.dataset.theme = nuevo;
      try { localStorage.setItem('tema', nuevo); } catch (e) { /* sin almacenamiento */ }
      actualizarBotonTema();
      if (modelo) renderVista();
    });
    if (window.matchMedia) {
      window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => { if (modelo && !document.documentElement.dataset.theme) renderVista(); });
    }
    actualizarBotonTema();
  }
  function actualizarBotonTema() {
    const b = $('#btn-tema');
    const oscuro = temaActual() === 'dark';
    b.setAttribute('aria-label', oscuro ? 'Cambiar a modo claro' : 'Cambiar a modo oscuro');
    b.textContent = oscuro ? '☀' : '☾';
  }

  /* ---------- Carga de datos (JSON precalculado) ---------- */
  async function cargarDatos() {
    const st = $('#estado-carga');
    try {
      modelo = await Datos.cargar();
      estado.marca = modelo.marcas[0].hoja;
      $('#carga').hidden = true;
      $('#app').hidden = false;
      $('#tabs').hidden = false;
      renderMeta();
      irA(estado.vista);
    } catch (err) {
      console.error(err);
      st.className = 'estado-carga error';
      st.textContent = err.message;
    }
  }

  /* ---------- Cabecera de datos + banner ---------- */
  function renderMeta() {
    const m = modelo;
    $('#meta').innerHTML = `
      <div class="meta-item"><span class="meta-k">Fuente</span><span class="meta-v">${esc(m.archivo)}</span></div>
      <div class="meta-item"><span class="meta-k">Generado</span><span class="meta-v">${m.generado ? `${fmtFecha(m.generado)} ${pad(m.generado.getHours())}:${pad(m.generado.getMinutes())}` : '—'}</span></div>
      <div class="meta-item"><span class="meta-k">Facturación</span><span class="meta-v">hoja <code>"${esc(m.fact.nombre)}"</code> · ${nf0.format(m.fact.filas)} filas · corte <strong>${fmtFecha(m.fact.corte)}</strong></span></div>
      <div class="meta-item"><span class="meta-k">Backorder</span><span class="meta-v">hoja <code>"${esc(m.bo.nombre)}"</code> · ${nf0.format(m.bo.filas)} filas · corte <strong>${fmtFecha(m.bo.corte)}</strong></span></div>
      <div class="meta-item"><span class="meta-k">Conciliación</span><span class="meta-v"><button type="button" class="link" data-ir="conciliacion">${m.resumen.diferencias ? `✕ ${nf0.format(m.resumen.diferencias)} diferencias` : `✔ ${nf0.format(m.resumen.ok)}/${nf0.format(m.resumen.comparaciones)} cuadran`}</button></span></div>`;
    const r = m.resumen;
    const banner = $('#banner');
    if (r.diferencias || r.errores) {
      const partes = [];
      if (r.diferencias) partes.push(`${r.diferencias} ${r.diferencias === 1 ? 'valor no cuadra' : 'valores no cuadran'} con el Excel`);
      if (r.errores) partes.push(`${r.errores} ${r.errores === 1 ? 'error' : 'errores'} de datos`);
      banner.innerHTML = `<span aria-hidden="true">⚠</span> Conciliación: ${partes.join(' y ')}. <button type="button" class="link" data-ir="conciliacion">Ver detalle</button>`;
      banner.hidden = false;
    } else {
      banner.hidden = true;
    }
  }

  /* ---------- Navegación ---------- */
  function initTabs() {
    document.addEventListener('click', e => {
      const t = e.target.closest('[data-ir]');
      if (t) { e.preventDefault(); irA(t.dataset.ir); }
    });
  }
  function irA(vista) {
    estado.vista = vista;
    $$('#tabs [role=tab]').forEach(b => b.setAttribute('aria-selected', String(b.dataset.ir === vista)));
    $$('.vista').forEach(v => { v.hidden = v.id !== `vista-${vista}`; });
    renderVista();
    window.scrollTo({ top: 0 });
  }
  function renderVista() {
    Object.keys(charts).forEach(k => { charts[k].destroy(); delete charts[k]; });
    if (!modelo) return;
    if (typeof Chart !== 'undefined') {
      Chart.defaults.font.family = css('--font');
      Chart.defaults.color = css('--ink-2');
      Chart.defaults.borderColor = css('--grid');
    }
    ({ resumen: renderResumen, ficha: renderFicha, conciliacion: renderConciliacion, metodo: renderMetodo })[estado.vista]();
  }

  /* ---------- Plugins Chart.js propios ---------- */
  const pluginMeta = {
    id: 'meta',
    afterDatasetsDraw(chart, args, o) {
      if (!o || o.valor == null) return;
      const esc_ = chart.scales[o.eje];
      const a = chart.chartArea;
      const p = esc_.getPixelForValue(o.valor);
      const { ctx } = chart;
      ctx.save();
      ctx.strokeStyle = o.color; ctx.lineWidth = 1.5; ctx.setLineDash([5, 4]);
      ctx.beginPath();
      if (o.eje === 'x') { ctx.moveTo(p, a.top); ctx.lineTo(p, a.bottom); } else { ctx.moveTo(a.left, p); ctx.lineTo(a.right, p); }
      ctx.stroke();
      ctx.setLineDash([]);
      ctx.fillStyle = o.color; ctx.font = `600 11px ${css('--font')}`;
      if (o.eje === 'x') { ctx.textAlign = 'left'; ctx.textBaseline = 'top'; ctx.fillText(o.texto, p + 4, a.top); } else { ctx.textAlign = 'left'; ctx.textBaseline = 'bottom'; ctx.fillText(o.texto, a.left + 4, p - 3); }
      ctx.restore();
    },
  };
  const pluginEtiquetas = {
    id: 'etiquetas',
    afterDatasetsDraw(chart, args, o) {
      if (!o || !o.fmt) return;
      const { ctx } = chart;
      ctx.save();
      ctx.font = `600 11px ${css('--font')}`; ctx.fillStyle = o.color;
      chart.data.datasets.forEach((ds, i) => {
        if (o.datasets && !o.datasets.includes(i)) return;
        const meta = chart.getDatasetMeta(i);
        if (meta.hidden) return;
        meta.data.forEach((el, j) => {
          const v = ds.data[j];
          if (v == null) return;
          if (chart.options.indexAxis === 'y') { ctx.textAlign = 'left'; ctx.textBaseline = 'middle'; ctx.fillText(o.fmt(v, j, i), el.x + 5, el.y); } else { ctx.textAlign = 'center'; ctx.textBaseline = 'bottom'; ctx.fillText(o.fmt(v, j, i), el.x, el.y - 4); }
        });
      });
      ctx.restore();
    },
  };
  const colorEstado = e => css(`--st-${e}`);
  const sinChart = () => typeof Chart === 'undefined';

  /* ---------- Vista Resumen ---------- */
  function filasResumen() {
    return modelo.marcas.map(m => {
      const b = m.bloques[0];
      const c = b.calc;
      if (!c) return { m, nombre: m.nombre, moneda: m.moneda, error: true };
      return { m, nombre: m.nombre, moneda: m.moneda, fact: c.total.fact, bo: c.total.bo, factbo: c.total.factbo,
        fc: c.tieneForecast ? c.total.fc : null, cumpl: c.total.cumpl, ytd: c.ytd.cumpl, trims: c.trims, b };
    });
  }

  function miniBarras(r) {
    if (!r.trims) return '';
    return `<div class="minibarras" role="img" aria-label="Cumplimiento por trimestre: ${r.trims.map(t => `${t.nombre} ${fmtPct(t.cumpl)}`).join(', ')}">${r.trims.map(t => {
      const e = semaforo(t.cumpl);
      const h = t.cumpl == null ? 6 : Math.max(4, Math.min(t.cumpl, 2) / 2 * 100);
      const parcial = t.estado === 'parcial';
      const tip = `${t.nombre}${parcial ? ' (parcial: el trimestre no ha terminado)' : t.estado === 'futuro' ? ' (futuro)' : ''}: ${t.cumpl == null ? 'sin forecast' : fmtPct(t.cumpl)}\nFact+BO ${fmt(t.factbo)} / Forecast ${fmt(t.fc)} ${r.moneda}`;
      return `<span class="mb${parcial ? ' parcial' : ''}" data-tip="${esc(tip)}"><span class="mb-col"><span class="mb-bar mb-${e}" style="height:${h}%"></span></span><span class="mb-l">${t.nombre}${parcial ? '<sup>p</sup>' : ''}</span></span>`;
    }).join('')}</div>`;
  }

  function renderResumen() {
    const v = $('#vista-resumen');
    let filas = filasResumen();
    const total = filas.length;
    const conFc = filas.filter(r => r.fc != null);
    const cuenta = e => conFc.filter(r => semaforo(r.ytd) === e).length;
    const corte = modelo.corte;
    if (estado.soloForecast) filas = conFc;
    const { k, dir } = estado.orden;
    filas = [...filas].sort((a, b) => {
      const x = a[k], y = b[k];
      if (typeof x === 'string' || typeof y === 'string') return dir * String(x || '').localeCompare(String(y || ''), 'es');
      if (x == null && y == null) return a.nombre.localeCompare(b.nombre, 'es');
      if (x == null) return 1;
      if (y == null) return -1;
      return dir * (x - y);
    });
    const mesCorte = corte ? Util.MESES[corte.getMonth()] : '';
    const cols = [
      ['nombre', 'Marca', 'text'], ['moneda', 'Moneda', 'text'], ['fact', 'Facturación', 'num'], ['bo', 'BO abierto', 'num'],
      ['factbo', 'Fact + BO', 'num'], ['fc', 'Forecast anual', 'num'], ['cumpl', 'Cumpl. anual (Excel)', 'num'],
      ['ytd', 'Cumpl. YTD (app)', 'num'],
    ];
    const anio = modelo.marcas[0].bloques[0].meses[0].anio;
    v.innerHTML = `
      <div class="encabezado-vista">
        <div>
          <h1>Resumen ${esc(anio)}</h1>
          <p class="sub">Año calendario por marca, cada una en su moneda. Entre marcas solo se comparan porcentajes y conteos.</p>
        </div>
      </div>
      <div class="stats">
        <div class="stat"><span class="stat-v">${total}</span><span class="stat-k">marcas visibles</span></div>
        <div class="stat"><span class="stat-v">${conFc.length}</span><span class="stat-k">con forecast</span></div>
        <div class="stat"><span class="stat-v"><span class="dot dot-verde" aria-hidden="true">${ICONO.verde}</span>${cuenta('verde')}</span><span class="stat-k">YTD ${esc(etiquetaSemaforo('verde'))}</span></div>
        <div class="stat"><span class="stat-v"><span class="dot dot-amarillo" aria-hidden="true">${ICONO.amarillo}</span>${cuenta('amarillo')}</span><span class="stat-k">YTD ${esc(etiquetaSemaforo('amarillo'))}</span></div>
        <div class="stat"><span class="stat-v"><span class="dot dot-rojo" aria-hidden="true">${ICONO.rojo}</span>${cuenta('rojo')}</span><span class="stat-k">YTD ${esc(etiquetaSemaforo('rojo'))}</span></div>
      </div>
      <section class="panel">
        <div class="panel-h">
          <h2>Cumplimiento por marca</h2>
          <p class="nota">YTD (calculado por la app, enero a ${esc(mesCorte)} incluido) vs. anual del Excel (acumulado / 12 meses de forecast). Clic en una barra para abrir la ficha.</p>
        </div>
        ${conFc.length ? `<div class="chart-box" style="height:${Math.max(180, conFc.length * 64 + 60)}px"><canvas id="ch-resumen" aria-label="Gráfico de cumplimiento por marca"></canvas></div>` : '<p class="vacio">Ninguna marca tiene forecast.</p>'}
      </section>
      <section class="panel">
        <div class="panel-h panel-h-fila">
          <h2>Detalle por marca</h2>
          <div class="controles">
            <label class="check"><input type="checkbox" id="solo-fc" ${estado.soloForecast ? 'checked' : ''}> Solo con forecast</label>
            <label class="orden-movil">Ordenar
              <select id="orden-sel">${cols.map(([ck, cl]) => `<option value="${ck}" ${ck === k ? 'selected' : ''}>${esc(cl)}</option>`).join('')}</select>
              <button type="button" class="btn-sec btn-mini" id="orden-dir" aria-label="Invertir orden">${dir > 0 ? '↑' : '↓'}</button>
            </label>
          </div>
        </div>
        <div class="leyenda-semaforo">${['verde', 'amarillo', 'rojo', 'gris'].map(e => `<span class="chip chip-${e}"><span aria-hidden="true">${ICONO[e]}</span>${esc(etiquetaSemaforo(e))}</span>`).join('')}<span class="nota">Semáforo configurable en <code>CONFIG.semaforo</code> (app.js). <sup>p</sup> = trimestre parcial.</span></div>
        <div class="tabla-wrap">
          <table class="tabla tabla-resumen">
            <thead><tr>${cols.map(([ck, cl, t]) => `<th scope="col" class="${t === 'num' ? 'n' : ''}" aria-sort="${ck === k ? (dir > 0 ? 'ascending' : 'descending') : 'none'}"><button type="button" data-orden="${ck}">${esc(cl)}${ck === k ? `<span aria-hidden="true">${dir > 0 ? ' ↑' : ' ↓'}</span>` : ''}</button></th>`).join('')}<th scope="col">Q1 – Q4</th></tr></thead>
            <tbody>${filas.map(r => r.error ? `<tr><td data-l="Marca" class="marca-cel">${esc(r.nombre)}</td><td colspan="8" class="error-txt">Estructura de hoja inesperada. Ver Conciliación.</td></tr>` : `
              <tr class="fila-click" data-marca="${esc(r.m.hoja)}" tabindex="0">
                <td data-l="Marca" class="marca-cel"><span class="marca-n">${esc(r.nombre)}</span>${r.m.convierte ? `<span class="tag" data-tip="${esc(`Incluye documentos convertidos a ${r.moneda}.${textoConversion(r.m)}`)}">conv.</span>` : ''}</td>
                <td data-l="Moneda">${esc(r.moneda)}</td>
                <td data-l="Facturación" class="n" data-tip="${esc(origen(r.m, r.b, 'fact'))}">${fmt(r.fact)}</td>
                <td data-l="BO abierto" class="n" data-tip="${esc(origen(r.m, r.b, 'bo'))}">${fmt(r.bo)}</td>
                <td data-l="Fact + BO" class="n" data-tip="${esc(origen(r.m, r.b, 'factbo'))}">${fmt(r.factbo)}</td>
                <td data-l="Forecast anual" class="n" data-tip="${esc(origen(r.m, r.b, 'fc'))}">${r.fc == null ? '<span class="muted">sin forecast</span>' : fmt(r.fc)}</td>
                <td data-l="Cumpl. anual (Excel)" class="n">${chip(r.cumpl, origen(r.m, r.b, 'cumpl'))}</td>
                <td data-l="Cumpl. YTD (app)" class="n">${chip(r.ytd, origen(r.m, r.b, 'ytd'))}</td>
                <td data-l="Q1 – Q4">${r.fc == null ? '<span class="muted">—</span>' : miniBarras(r)}</td>
              </tr>`).join('')}
            </tbody>
          </table>
        </div>
      </section>`;

    $('#solo-fc').addEventListener('change', e => { estado.soloForecast = e.target.checked; renderResumen(); });
    $('#orden-sel').addEventListener('change', e => { estado.orden = { k: e.target.value, dir: estado.orden.dir }; renderResumen(); });
    $('#orden-dir').addEventListener('click', () => { estado.orden.dir *= -1; renderResumen(); });
    $$('[data-orden]', v).forEach(b => b.addEventListener('click', () => {
      const ck = b.dataset.orden;
      estado.orden = { k: ck, dir: estado.orden.k === ck ? -estado.orden.dir : (ck === 'nombre' || ck === 'moneda' ? 1 : -1) };
      renderResumen();
    }));
    $$('.fila-click', v).forEach(tr => {
      const abrir = () => { estado.marca = tr.dataset.marca; estado.bloque = 0; irA('ficha'); };
      tr.addEventListener('click', e => { if (!e.target.closest('.info')) abrir(); });
      tr.addEventListener('keydown', e => { if (e.key === 'Enter') abrir(); });
    });

    if (conFc.length && !sinChart()) {
      const orden = [...conFc].sort((a, b) => (b.ytd ?? -1) - (a.ytd ?? -1));
      charts.resumen = new Chart($('#ch-resumen'), {
        type: 'bar',
        data: {
          labels: orden.map(r => r.nombre),
          datasets: [
            { label: 'Cumplimiento YTD (app)', data: orden.map(r => r.ytd == null ? null : r.ytd * 100), backgroundColor: css('--s1'), borderRadius: 4, maxBarThickness: 18, categoryPercentage: 0.75, barPercentage: 0.9 },
            { label: 'Cumplimiento anual (Excel)', data: orden.map(r => r.cumpl == null ? null : r.cumpl * 100), backgroundColor: css('--ref'), borderRadius: 4, maxBarThickness: 18, categoryPercentage: 0.75, barPercentage: 0.9 },
          ],
        },
        options: {
          indexAxis: 'y', responsive: true, maintainAspectRatio: false,
          layout: { padding: { right: 48 } },
          scales: {
            x: { beginAtZero: true, grid: { color: css('--grid') }, ticks: { callback: x => `${x}%` } },
            y: { grid: { display: false }, ticks: { color: css('--ink-1'), font: { weight: 600 } } },
          },
          plugins: {
            legend: { position: 'top', align: 'start', labels: { boxWidth: 12, boxHeight: 12, useBorderRadius: true, borderRadius: 3 } },
            tooltip: { callbacks: { label: c => `${c.dataset.label}: ${fmtPct(c.parsed.x / 100)}`, afterBody: items => { const r = orden[items[0].dataIndex]; return `Semáforo YTD: ${etiquetaSemaforo(semaforo(r.ytd))}`; } } },
            meta: { valor: 100, eje: 'x', color: css('--ink-2'), texto: '100%' },
            etiquetas: { color: css('--ink-2'), fmt: v => fmtPct(v / 100) },
          },
          onClick: (e, els) => { if (els.length) { estado.marca = orden[els[0].index].m.hoja; estado.bloque = 0; irA('ficha'); } },
          onHover: (e, els) => { e.native.target.style.cursor = els.length ? 'pointer' : 'default'; },
        },
        plugins: [pluginMeta, pluginEtiquetas],
      });
    }
  }

  /* ---------- Vista Ficha ---------- */
  function renderFicha() {
    const v = $('#vista-ficha');
    const marca = modelo.marcas.find(m => m.hoja === estado.marca) || modelo.marcas[0];
    estado.marca = marca.hoja;
    if (estado.bloque >= marca.bloques.length) estado.bloque = 0;
    const b = marca.bloques[estado.bloque];
    const c = b.calc;
    const mon = marca.moneda;

    const selector = `
      <div class="ficha-tools">
        <label class="campo"><span>Marca</span>
          <select id="sel-marca">${modelo.marcas.map(m => `<option value="${esc(m.hoja)}" ${m === marca ? 'selected' : ''}>${esc(m.nombre)}</option>`).join('')}</select>
        </label>
        ${marca.tieneFiscal ? `<div class="segmentado" role="group" aria-label="Período">${marca.bloques.map((bl, i) => `<button type="button" data-bloque="${i}" aria-pressed="${i === estado.bloque}">${bl.tipo === 'calendario' ? 'Año calendario' : 'Año fiscal USA'}</button>`).join('')}</div>` : ''}
      </div>`;

    if (!c) {
      v.innerHTML = `${selector}<div class="panel error-txt">La hoja "${esc(marca.hoja)}" (${esc(b.tipo)}) tiene una estructura inesperada: ${esc(b.errores.join(' '))}</div>`;
      bindFicha(v);
      return;
    }

    const y = c.ytd;
    const ytdTitulo = b.tipo === 'calendario'
      ? `Cumplimiento YTD${y.hasta ? ` (${Util.MESES_CORTOS[y.desde.mes].toLowerCase()}–${Util.MESES_CORTOS[y.hasta.mes].toLowerCase()})` : ''}`
      : `Cumplimiento acumulado${y.completo ? ' (período completo)' : ''}`;
    const kpi = (titulo, valorHtml, tip, extra = '') => `
      <div class="kpi"><div class="kpi-k">${esc(titulo)} ${info(tip)}</div><div class="kpi-v">${valorHtml}</div>${extra ? `<div class="kpi-x">${extra}</div>` : ''}</div>`;
    const monto = v_ => `${fmt(v_)} <span class="mon">${esc(mon)}</span>`;
    const trimParcial = c.trims.find(t => t.estado === 'parcial');

    v.innerHTML = `
      ${selector}
      <div class="encabezado-vista">
        <div>
          <h1>${esc(marca.nombre)} <span class="moneda-chip">${esc(mon)}</span></h1>
          <p class="sub">${esc(b.titulo)} · hoja <code>"${esc(marca.hoja)}"</code>${b.tipo === 'fiscal' ? ` · columnas ${Util.colLetra(b.c0)}–${Util.colLetra(b.c0 + 17)}` : ''}${c.tieneForecast ? '' : ' · <strong>sin forecast</strong>: solo Facturación y BO'}</p>
        </div>
      </div>
      <div class="kpis">
        ${kpi('Facturación', monto(c.total.fact), origen(marca, b, 'fact'))}
        ${kpi('Backorder abierto', monto(c.total.bo), origen(marca, b, 'bo'))}
        ${kpi('Fact + BO', monto(c.total.factbo), origen(marca, b, 'factbo'))}
        ${c.tieneForecast ? `
          ${kpi('Forecast anual', monto(c.total.fc), origen(marca, b, 'fc'))}
          ${kpi('Cumplimiento anual (Excel)', chip(c.total.cumpl), origen(marca, b, 'cumpl'), 'acumulado / 12 meses de forecast')}
          ${kpi(ytdTitulo, chip(y.cumpl), origen(marca, b, 'ytd'), `<span class="tag tag-app">calculada por la app</span>${y.hasta ? ` Fact+BO ${fmt(y.factbo)} / Fc ${fmt(y.fc)}` : ''}`)}
        ` : kpi('Forecast', '<span class="muted">sin forecast</span>', `Fila 6 vacía o en 0 (${refTotal(marca, b, 6)}). No se calcula cumplimiento: se muestra "—", nunca 0%.`)}
      </div>
      ${renderConversion(marca, b)}
      <div class="grid-2">
        <section class="panel panel-ancho">
          <div class="panel-h"><h2>Facturación + Backorder mensual${c.tieneForecast ? ' vs. Forecast' : ''}</h2>
            <p class="nota">Montos en ${esc(mon)}. ${c.hayDerivado ? 'Forecast punteado = derivado (trimestre / 3) porque las celdas mensuales están vacías.' : ''}</p></div>
          <div class="chart-box" style="height:300px"><canvas id="ch-mensual" aria-label="Gráfico mensual"></canvas></div>
          ${tablaMensualAccesible(c, mon)}
        </section>
        <section class="panel">
          <div class="panel-h"><h2>Cumplimiento trimestral</h2>
            <p class="nota">Fact+BO del trimestre / Forecast del trimestre (fila 7).${trimParcial ? ` ${trimParcial.nombre} es <strong>parcial</strong>: el corte (${fmtFecha(modelo.corte)}) es anterior al fin del trimestre.` : ''}</p></div>
          ${c.trims.some(t => t.cumpl != null) ? '<div class="chart-box" style="height:260px"><canvas id="ch-trim" aria-label="Gráfico de cumplimiento trimestral"></canvas></div>' : '<p class="vacio">Sin forecast: no hay cumplimiento trimestral.</p>'}
          <table class="tabla tabla-mini"><thead><tr><th>Trim.</th><th class="n">Fact+BO</th><th class="n">Forecast</th><th class="n">Cumpl.</th></tr></thead>
          <tbody>${c.trims.map((t, q) => `<tr><td>${t.nombre}${t.estado === 'parcial' ? ' <span class="tag">parcial</span>' : t.estado === 'futuro' ? ' <span class="tag">futuro</span>' : ''}</td><td class="n">${fmt(t.factbo)}</td><td class="n">${t.fc > 0 ? fmt(t.fc) : '—'}</td><td class="n">${chip(t.cumpl, `Equivale a ${Util.refHoja(marca.hoja)}!${Util.colLetra(b.c0 + Util.LAYOUT.trimestres[q])}7 = Fact+BO ${t.nombre} / Forecast ${t.nombre}`)}</td></tr>`).join('')}</tbody></table>
        </section>
      </div>
      ${renderCategorias(marca, b, c)}
      ${modelo.incluirDetalle ? `<section class="panel">
        <div class="panel-h panel-h-fila"><div><h2>Backorder abierto</h2><p class="nota">Líneas con Cantidad pdte &gt; 0. Antigüedad en días a la fecha de corte de BO (${fmtFecha(modelo.bo.corte)}).</p></div>
          <label class="check"><input type="checkbox" id="bo-fuera" ${estado.boFuera ? 'checked' : ''}> Incluir OC fuera del período</label></div>
        <div id="tabla-bo"></div>
      </section>
      <section class="panel">
        <div class="panel-h"><h2>Facturas del período</h2><p class="nota">${esc(b.titulo)}. Monto = "Total moneda marca" (lo que suma la hoja).</p></div>
        <div id="tabla-fact"></div>
      </section>` : '<p class="nota detalle-omitido">El detalle de BO abiertas y facturas no se publica (generar_datos.py con INCLUIR_DETALLE = False).</p>'}`;

    bindFicha(v);
    if (modelo.incluirDetalle) {
      $('#bo-fuera').addEventListener('change', e => { estado.boFuera = e.target.checked; renderTablaBO(marca, c); });
      renderTablaBO(marca, c);
      renderTablaFact(marca, c);
    }
    chartsFicha(marca, b, c);
  }

  function bindFicha(v) {
    $('#sel-marca', v).addEventListener('change', e => { estado.marca = e.target.value; estado.bloque = 0; renderVista(); });
    $$('[data-bloque]', v).forEach(btn => btn.addEventListener('click', () => { estado.bloque = +btn.dataset.bloque; renderVista(); }));
  }

  function tablaMensualAccesible(c, mon) {
    return `<details class="tabla-acc"><summary>Ver datos del gráfico en tabla</summary><div class="tabla-wrap"><table class="tabla tabla-mini">
      <thead><tr><th>Mes</th><th class="n">Facturación</th><th class="n">Backorder</th><th class="n">Fact+BO</th><th class="n">Forecast</th></tr></thead>
      <tbody>${c.meses.map(m => `<tr><td>${esc(m.etiqueta)}</td><td class="n">${fmt(m.fact)}</td><td class="n">${fmt(m.bo)}</td><td class="n">${fmt(m.factbo)}</td><td class="n">${m.fc ? fmt(m.fc) + (m.fcDerivado ? ' <span class="tag">derivado</span>' : '') : '—'}</td></tr>`).join('')}</tbody>
    </table></div><p class="nota">Montos en ${esc(mon)}.</p></details>`;
  }

  function renderConversion(marca, b) {
    if (!marca.convierte) return '';
    const filas = b.conversion.map(f => ({ ...f, etiqueta: f.base }));
    const hayEnPeriodo = filas.some(f => f.convertido);
    const pares = Object.entries(modelo.parametros.pares);
    return `<section class="panel panel-conv">
      <div class="panel-h"><h2><span aria-hidden="true">⇄</span> Conversión de moneda</h2>
        <p class="nota">${esc(marca.nombre)} se reporta en <strong>${esc(marca.moneda)}</strong>. Los documentos en otra moneda se multiplican por el TC de la hoja "${esc(modelo.hojas.moneda)}" (columnas D:E). ${hayEnPeriodo ? '' : '<strong>En este período no hay documentos convertidos.</strong>'}</p></div>
      <div class="tabla-wrap"><table class="tabla tabla-mini">
        <thead><tr><th>Base</th><th>Moneda original</th><th class="n">Docs</th><th class="n">Monto original</th><th class="n">TC usado</th><th class="n">Monto convertido (${esc(marca.moneda)})</th></tr></thead>
        <tbody>${filas.map(f => `<tr class="${f.convertido ? 'resaltar' : ''}"><td>${esc(f.etiqueta)}</td><td>${esc(f.moneda)}${f.convertido ? ` <span class="tag">→ ${esc(marca.moneda)}</span>` : ''}</td><td class="n">${f.n}</td><td class="n">${fmt(f.orig)}</td><td class="n">${f.tcs.map(fmtTC).join(' / ')}</td><td class="n">${fmt(f.conv)}</td></tr>`).join('') || '<tr><td colspan="6" class="muted">Sin documentos en el período.</td></tr>'}</tbody>
      </table></div>
      <p class="nota">Tipos de cambio en "${esc(modelo.hojas.moneda)}": ${pares.map(([p, x]) => `${esc(p)} = ${fmtTC(x.tc)} (${esc(x.celda)})`).join(' · ') || '—'}. En BO, el monto original es "Total Neto ME" = Precio × Cantidad pdte.</p>
    </section>`;
  }

  function renderCategorias(marca, b, c) {
    const cats = [...modelo.categorias, modelo.sinCategoria];
    const totF = c.total.fact, totB = c.total.bo;
    const max = Math.max(1, ...cats.map(k => Math.max(Math.abs(c.catF[k]), Math.abs(c.catB[k]))));
    const barra = (v_, clase) => `<span class="cat-bar"><span class="${clase}" style="width:${Math.abs(v_) / max * 100}%"></span></span>`;
    const raw = obj => Object.entries(obj).map(([k, v_]) => `${k}: ${fmt(v_)}`).join(' · ');
    const T = Util.colLetra(b.c0 + Util.LAYOUT.total);
    const filaCat = (k, i) => {
      const sin = k === modelo.sinCategoria;
      const tipF = sin ? `Valores de "Categoria" fuera de las 4 estándar: ${raw(c.rawF) || 'ninguno'}.\nNo existe en la hoja (no suma en la fila 23).` : `Equivale a ${Util.refHoja(marca.hoja)}!${T}${19 + i}`;
      const tipB = sin ? `Valores de "Grupo" fuera de las 4 estándar: ${raw(c.rawB) || 'ninguno'}.\nNo existe en la hoja (no suma en la fila 15).` : `Equivale a ${Util.refHoja(marca.hoja)}!${T}${11 + i}`;
      return `<tr class="${sin ? 'sin-cat' : ''}"><th scope="row">${esc(k)}</th>
        <td class="n" data-tip="${esc(tipF)}">${fmt(c.catF[k])}</td><td class="n muted">${totF ? fmtPct(c.catF[k] / totF) : '—'}</td><td class="bar-cel">${barra(c.catF[k], 'cb-f')}</td>
        <td class="n" data-tip="${esc(tipB)}">${fmt(c.catB[k])}</td><td class="n muted">${totB ? fmtPct(c.catB[k] / totB) : '—'}</td><td class="bar-cel">${barra(c.catB[k], 'cb-b')}</td></tr>`;
    };
    const s4 = cat => modelo.categorias.reduce((s, k) => s + cat[k], 0);
    return `<section class="panel">
      <div class="panel-h"><h2>Desglose por categoría</h2><p class="nota">Facturación por "Categoria"; BO por "Grupo". Montos en ${esc(marca.moneda)}. % sobre el total de la fila 3 / 4 (como la columna S).</p></div>
      <div class="tabla-wrap"><table class="tabla tabla-cat">
        <thead><tr><th>Categoría</th><th class="n">Facturación</th><th class="n">%</th><th></th><th class="n">Backorder</th><th class="n">%</th><th></th></tr></thead>
        <tbody>${cats.map(filaCat).join('')}</tbody>
        <tfoot>
          <tr><th scope="row">Total 4 categorías</th><td class="n" data-tip="${esc(`Equivale a ${Util.refHoja(marca.hoja)}!${T}23`)}">${fmt(s4(c.catF))}</td><td></td><td></td><td class="n" data-tip="${esc(`Equivale a ${Util.refHoja(marca.hoja)}!${T}15`)}">${fmt(s4(c.catB))}</td><td></td><td></td></tr>
          <tr><th scope="row">Total general</th><td class="n">${fmt(totF)}</td><td></td><td></td><td class="n">${fmt(totB)}</td><td></td><td></td></tr>
        </tfoot>
      </table></div></section>`;
  }

  function chartsFicha(marca, b, c) {
    if (sinChart()) return;
    const mon = marca.moneda;
    const surface = css('--surface');
    const datasets = [
      { label: 'Facturación', data: c.meses.map(m => m.fact), backgroundColor: css('--s1'), stack: 's', borderColor: surface, borderWidth: { top: 0, bottom: 0, left: 0, right: 0 }, maxBarThickness: 40 },
      { label: 'Backorder', data: c.meses.map(m => m.bo), backgroundColor: css('--s2'), stack: 's', borderColor: surface, borderWidth: { bottom: 2, top: 0, left: 0, right: 0 }, borderSkipped: false, borderRadius: { topLeft: 4, topRight: 4 }, maxBarThickness: 40 },
    ];
    if (c.tieneForecast) {
      const der = c.meses.map(m => m.fcDerivado);
      datasets.push({
        type: 'line', label: 'Forecast', data: c.meses.map(m => m.fc || null), stack: 'fc',
        borderColor: css('--ink-1'), backgroundColor: css('--ink-1'), borderWidth: 2, pointRadius: 4, pointHoverRadius: 6,
        pointBackgroundColor: c.meses.map(m => (m.fcDerivado ? surface : css('--ink-1'))), pointBorderWidth: 2,
        segment: { borderDash: s => (der[s.p0DataIndex] || der[s.p1DataIndex] ? [6, 4] : undefined) },
        spanGaps: false, order: -1,
      });
    }
    charts.mensual = new Chart($('#ch-mensual'), {
      type: 'bar',
      data: { labels: c.meses.map(m => m.etiqueta), datasets },
      options: {
        responsive: true, maintainAspectRatio: false,
        interaction: { mode: 'index', intersect: false },
        scales: {
          x: { stacked: true, grid: { display: false } },
          y: { stacked: true, beginAtZero: true, grid: { color: css('--grid') }, ticks: { callback: x => nfCompacto.format(x) } },
        },
        plugins: {
          legend: { position: 'top', align: 'start', labels: { boxWidth: 12, boxHeight: 12, useBorderRadius: true, borderRadius: 3 } },
          tooltip: {
            callbacks: {
              label: t => `${t.dataset.label}: ${fmt(t.parsed.y)} ${mon}${t.dataset.label === 'Forecast' && c.meses[t.dataIndex].fcDerivado ? ' (derivado: trimestre/3)' : ''}`,
              footer: items => `Fact+BO: ${fmt(c.meses[items[0].dataIndex].factbo)} ${mon}`,
            },
          },
        },
      },
    });

    if (!$('#ch-trim')) return;
    const etiquetas = c.trims.map(t => (t.estado === 'parcial' ? `${t.nombre} (parcial)` : t.nombre));
    const valores = c.trims.map(t => (t.cumpl == null ? null : t.cumpl * 100));
    charts.trim = new Chart($('#ch-trim'), {
      type: 'bar',
      data: { labels: etiquetas, datasets: [{ label: 'Cumplimiento', data: valores, backgroundColor: c.trims.map(t => colorEstado(semaforo(t.cumpl))), borderRadius: 4, maxBarThickness: 56 }] },
      options: {
        responsive: true, maintainAspectRatio: false,
        layout: { padding: { top: 18 } },
        scales: {
          y: { beginAtZero: true, suggestedMax: 120, grid: { color: css('--grid') }, ticks: { callback: x => `${x}%` } },
          x: { grid: { display: false } },
        },
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: t => { const q = c.trims[t.dataIndex]; return [`Cumplimiento: ${fmtPct(q.cumpl)}`, `Fact+BO ${fmt(q.factbo)} / Forecast ${fmt(q.fc)} ${mon}`, `Semáforo: ${etiquetaSemaforo(semaforo(q.cumpl))}`]; } } },
          meta: { valor: 100, eje: 'y', color: css('--ink-2'), texto: '100%' },
          etiquetas: { color: css('--ink-1'), fmt: x => fmtPct(x / 100) },
        },
      },
      plugins: [pluginMeta, pluginEtiquetas],
    });
  }

  /* ---------- Tablas con búsqueda, orden y CSV ---------- */
  function tablaDatos(cont, def) {
    const st = { q: '', k: def.orden.k, dir: def.orden.dir };
    cont.innerHTML = `<div class="tabla-tools">
        <input type="search" placeholder="Buscar…" aria-label="Buscar en ${esc(def.titulo)}">
        <span class="cuenta muted"></span>
        <button type="button" class="btn-sec">Exportar CSV</button>
      </div>
      <div class="tabla-wrap tabla-scroll"><table class="tabla tabla-datos"><thead></thead><tbody></tbody><tfoot></tfoot></table></div>`;
    const tbody = $('tbody', cont), thead = $('thead', cont), tfoot = $('tfoot', cont);
    const textoFila = r => def.columnas.map(c => {
      const x = c.v(r);
      return x instanceof Date ? fmtFecha(x) : typeof x === 'number' ? `${x} ${fmt(x)}` : String(x == null ? '' : x);
    }).join(' ').toLowerCase();
    const indices = def.filas.map(r => ({ r, t: textoFila(r) }));
    function filtradas() {
      const q = st.q.trim().toLowerCase();
      const col = def.columnas.find(c => c.k === st.k);
      const res = indices.filter(x => !q || q.split(/\s+/).every(p => x.t.includes(p))).map(x => x.r);
      if (col) {
        res.sort((a, b) => {
          const x = col.v(a), y = col.v(b);
          if (x instanceof Date || typeof x === 'number') return st.dir * ((x || 0) - (y || 0));
          return st.dir * String(x || '').localeCompare(String(y || ''), 'es', { numeric: true });
        });
      }
      return res;
    }
    function render() {
      const filas = filtradas();
      thead.innerHTML = `<tr>${def.columnas.map(c => `<th scope="col" class="${c.tipo === 'num' ? 'n' : ''}" aria-sort="${c.k === st.k ? (st.dir > 0 ? 'ascending' : 'descending') : 'none'}"><button type="button" data-k="${c.k}">${esc(c.l)}${c.k === st.k ? `<span aria-hidden="true">${st.dir > 0 ? ' ↑' : ' ↓'}</span>` : ''}</button></th>`).join('')}</tr>`;
      tbody.innerHTML = filas.length ? filas.map(r => `<tr>${def.columnas.map(c => `<td class="${c.tipo === 'num' ? 'n' : ''}${c.clase ? ` ${c.clase}` : ''}" data-l="${esc(c.l)}">${c.f ? c.f(r) : esc(c.v(r))}</td>`).join('')}</tr>`).join('')
        : `<tr><td colspan="${def.columnas.length}" class="vacio">Sin filas${st.q ? ' que coincidan con la búsqueda' : ''}.</td></tr>`;
      const sumables = def.columnas.filter(c => c.suma);
      tfoot.innerHTML = sumables.length && filas.length ? `<tr>${def.columnas.map((c, i) => c.suma ? `<td class="n"><strong>${fmt(filas.reduce((s, r) => s + (c.v(r) || 0), 0))}</strong></td>` : `<td>${i === 0 ? '<strong>Total</strong>' : ''}</td>`).join('')}</tr>` : '';
      $('.cuenta', cont).textContent = `${nf0.format(filas.length)} de ${nf0.format(def.filas.length)} filas`;
      $$('[data-k]', thead).forEach(btn => btn.addEventListener('click', () => {
        const k = btn.dataset.k;
        st.dir = st.k === k ? -st.dir : 1; st.k = k; render();
      }));
    }
    $('input', cont).addEventListener('input', e => { st.q = e.target.value; render(); });
    $('.btn-sec', cont).addEventListener('click', () => exportarCSV(def.csv, def.columnas, filtradas()));
    render();
  }

  function exportarCSV(nombre, columnas, filas) {
    const celdaCSV = x => {
      let s = x instanceof Date ? fmtFecha(x) : typeof x === 'number' ? String(x).replace('.', ',') : String(x == null ? '' : x);
      if (/[";\n\r]/.test(s)) s = `"${s.replace(/"/g, '""')}"`;
      return s;
    };
    const lineas = [columnas.map(c => celdaCSV(c.l)).join(';'), ...filas.map(r => columnas.map(c => celdaCSV(c.v(r))).join(';'))];
    const blob = new Blob(['﻿' + lineas.join('\r\n')], { type: 'text/csv;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `${nombre}.csv`;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
  }

  const archivoSeguro = s => s.replace(/[^\p{L}\p{N}_-]+/gu, '_').replace(/^_|_$/g, '');

  function renderTablaBO(marca, c) {
    const corte = modelo.bo.corte ? Util.soloDia(modelo.bo.corte) : null;
    const base = estado.boFuera ? marca.boTodas : c.filasB;
    const filas = base.filter(r => r.cantidad > 0);
    const antig = r => (corte ? Math.round((corte - Util.soloDia(r.fecha)) / 864e5) : null);
    tablaDatos($('#tabla-bo'), {
      titulo: 'backorder', csv: archivoSeguro(`BO_abiertas_${marca.nombre}`), filas,
      orden: { k: 'dias', dir: -1 },
      columnas: [
        { k: 'oc', l: 'OC', v: r => r.referencia },
        { k: 'fecha', l: 'Fecha OC', v: r => r.fecha, f: r => fmtFecha(r.fecha) },
        { k: 'prod', l: 'Producto', v: r => r.producto },
        { k: 'desc', l: 'Descripción', v: r => r.descripcion, clase: 'desc' },
        { k: 'cat', l: 'Grupo', v: r => r.cat },
        { k: 'cant', l: 'Cant. pdte', tipo: 'num', v: r => r.cantidad, f: r => fmt(r.cantidad) },
        { k: 'mon', l: 'Moneda orig.', v: r => r.monedaOrig, f: r => esc(r.monedaOrig) + (r.convertido ? ' <span class="tag">conv.</span>' : '') },
        { k: 'precio', l: 'Precio', tipo: 'num', v: r => r.precio, f: r => fmt(r.precio) },
        { k: 'orig', l: 'Total orig.', tipo: 'num', v: r => r.montoOrig, f: r => fmt(r.montoOrig) },
        { k: 'tc', l: 'TC', tipo: 'num', v: r => r.tc, f: r => fmtTC(r.tc) },
        { k: 'total', l: `Total ${marca.moneda}`, tipo: 'num', suma: true, v: r => r.total, f: r => fmt(r.total) },
        { k: 'dias', l: 'Antigüedad (días)', tipo: 'num', v: antig, f: r => { const d = antig(r); return `<span class="${d > 90 ? 'dias-alto' : ''}">${fmt(d)}</span>`; } },
      ],
    });
  }

  function renderTablaFact(marca, c) {
    tablaDatos($('#tabla-fact'), {
      titulo: 'facturas', csv: archivoSeguro(`Facturas_${marca.nombre}`), filas: c.filasF,
      orden: { k: 'fecha', dir: -1 },
      columnas: [
        { k: 'fecha', l: 'Fecha', v: r => r.fecha, f: r => fmtFecha(r.fecha) },
        { k: 'folio', l: 'Folio', v: r => r.folio },
        { k: 'com', l: 'Comentarios', v: r => r.comentarios, clase: 'desc' },
        { k: 'oc', l: 'OC', v: r => r.oc },
        { k: 'cat', l: 'Categoría', v: r => r.cat, f: r => esc(r.cat) + (r.cat === modelo.sinCategoria && r.catRaw ? ` <span class="muted">(${esc(r.catRaw)})</span>` : '') },
        { k: 'mon', l: 'Moneda orig.', v: r => r.monedaOrig, f: r => esc(r.monedaOrig) + (r.convertido ? ' <span class="tag">conv.</span>' : '') },
        { k: 'orig', l: 'Monto orig.', tipo: 'num', v: r => r.montoOrig, f: r => fmt(r.montoOrig) },
        { k: 'tc', l: 'TC', tipo: 'num', v: r => r.tc, f: r => fmtTC(r.tc) },
        { k: 'total', l: `Monto ${marca.moneda}`, tipo: 'num', suma: true, v: r => r.total, f: r => fmt(r.total) },
      ],
    });
  }

  /* ---------- Vista Conciliación ---------- */
  function renderConciliacion() {
    const v = $('#vista-conciliacion');
    const r = modelo.resumen;
    const icono = { ok: '✔', info: 'ℹ', aviso: '!', error: '✕' };
    const valorCel = it => (it.excelNoNumerico ? '<span class="error-txt">no numérica</span>' : it.excelVacia ? '<span class="muted">vacía</span>' : it.tipo === 'pct' ? fmtPct(it.excel) : fmt(it.excel));
    const fmtDif = (it) => {
      if (it.ok) return '<span class="ok-txt">✔</span>';
      if (it.dif == null || !Number.isFinite(it.dif)) return '<span class="error-txt">✕ celda no numérica</span>';
      return `<span class="error-txt">✕ ${it.tipo === 'pct' ? `${it.dif > 0 ? '+' : ''}${nf.format(it.dif * 100)} pp` : `${it.dif > 0 ? '+' : ''}${fmt(it.dif)}`}</span>`;
    };
    const bloquesHtml = modelo.marcas.map(m => {
      const items = m.bloques.flatMap(b => b.conc);
      const dif = items.filter(i => !i.ok).length;
      const erroresMarca = modelo.problemas.filter(p => p.marca === m.nombre && p.nivel !== 'ok');
      if (estado.soloDif && !dif && !erroresMarca.length && m.bloques.every(b => b.calc)) return '';
      return `<section class="panel conc-marca">
        <div class="panel-h panel-h-fila"><h2>${esc(m.nombre)} <span class="muted">· hoja "${esc(m.hoja)}" · ${esc(m.moneda)}</span></h2>
          <span class="chip ${dif ? 'chip-rojo' : 'chip-verde'}">${dif ? `✕ ${dif} con diferencia` : `✔ ${items.length}/${items.length}`}</span></div>
        ${erroresMarca.map(p => `<p class="problema problema-${p.nivel}"><span aria-hidden="true">${icono[p.nivel]}</span> ${esc(p.titulo)}${p.detalle ? `<br><span class="muted">${esc(p.detalle)}</span>` : ''}</p>`).join('')}
        ${m.bloques.map(b => !b.calc ? `<p class="error-txt">Bloque ${esc(b.tipo)}: ${esc(b.errores.join(' '))}</p>` : `
          <h3>${esc(b.titulo)}</h3>
          <div class="tabla-wrap"><table class="tabla tabla-conc">
            <thead><tr><th>Concepto</th><th>Celda</th><th class="n">Excel (cache)</th><th class="n">Recalculado app</th><th class="n">Resultado</th></tr></thead>
            <tbody>${b.conc.filter(it => !estado.soloDif || !it.ok).map(it => `<tr class="${it.ok ? '' : 'fila-dif'}">
              <td data-l="Concepto">${esc(it.concepto)}${it.nota ? ` ${info(it.nota)}` : ''}</td>
              <td data-l="Celda"><code>${esc(it.celda)}</code></td>
              <td data-l="Excel" class="n">${valorCel(it)}</td>
              <td data-l="App" class="n">${it.app == null ? '—' : it.tipo === 'pct' ? fmtPct(it.app) : fmt(it.app)}</td>
              <td data-l="Resultado" class="n">${fmtDif(it)}</td></tr>`).join('')}</tbody>
          </table></div>`).join('')}
      </section>`;
    }).join('');
    const globales = modelo.problemas.filter(p => !p.marca || p.nivel === 'ok');
    v.innerHTML = `
      <div class="encabezado-vista"><div>
        <h1>Conciliación</h1>
        <p class="sub">Resultado de generar_datos.py: lo recalculado desde las BBDD contra los valores cacheados de cada hoja de marca visible. Si algo no cuadra, el script no publica el JSON. Tolerancia: ${fmt(modelo.tolerancia.monto)} en montos y ${nfTC.format(modelo.tolerancia.pct)} en % (${nf.format(modelo.tolerancia.pct * 100)} pp).</p>
      </div></div>
      <div class="stats">
        <div class="stat"><span class="stat-v">${nf0.format(r.comparaciones)}</span><span class="stat-k">comparaciones</span></div>
        <div class="stat"><span class="stat-v ok-txt">✔ ${nf0.format(r.ok)}</span><span class="stat-k">cuadran</span></div>
        <div class="stat"><span class="stat-v ${r.diferencias ? 'error-txt' : ''}">${r.diferencias ? '✕ ' : ''}${nf0.format(r.diferencias)}</span><span class="stat-k">con diferencia</span></div>
        <div class="stat"><span class="stat-v ${r.errores ? 'error-txt' : ''}">${nf0.format(r.errores)}</span><span class="stat-k">errores de datos</span></div>
        <div class="stat"><span class="stat-v">${nf0.format(r.avisos)}</span><span class="stat-k">avisos</span></div>
      </div>
      <section class="panel">
        <div class="panel-h"><h2>Validaciones de datos</h2></div>
        <ul class="problemas">${[...globales, ...modelo.problemas.filter(p => p.marca && p.nivel !== 'ok')].map(p => `<li class="problema problema-${p.nivel}"><span class="pi" aria-hidden="true">${icono[p.nivel]}</span><div>${esc(p.titulo)}${p.detalle ? `<div class="muted det">${esc(p.detalle)}</div>` : ''}</div></li>`).join('')}</ul>
      </section>
      <div class="panel-h-fila filtros"><label class="check"><input type="checkbox" id="solo-dif" ${estado.soloDif ? 'checked' : ''}> Mostrar solo diferencias</label></div>
      <div class="conc-grid">${bloquesHtml || '<p class="vacio">No hay diferencias. Todo cuadra con el Excel. ✔</p>'}</div>`;
    $('#solo-dif').addEventListener('change', e => { estado.soloDif = e.target.checked; renderConciliacion(); });
  }

  /* ---------- Vista Cómo se calcula ---------- */
  function renderMetodo() {
    $$('[data-cfg]').forEach(el => {
      const k = el.dataset.cfg;
      el.textContent = k === 'verde' ? fmtPct(CONFIG.semaforo.verdeDesde) : k === 'amarillo' ? fmtPct(CONFIG.semaforo.amarilloDesde)
        : k === 'tolMonto' ? fmt(modelo.tolerancia.monto) : k === 'tolPct' ? nfTC.format(modelo.tolerancia.pct) : '';
    });
    const pares = Object.entries(modelo.parametros.pares).map(([p, x]) => `${p} × ${fmtTC(x.tc)}`).join(' · ');
    const marcasMon = modelo.parametros.marcas.map(m => `${m.marca.trim()} → ${m.moneda}`).join(' · ');
    $('#metodo-moneda').textContent = `Hoy en "${modelo.hojas.moneda}": ${marcasMon || '—'}. Pares: ${pares || '—'}.`;
  }

  function init() {
    initTema();
    initTooltips();
    cargarDatos();
    initTabs();
  }

  return { init };
})();

/* Acceso: sin sesión iniciada en index.html se vuelve al login (barrera básica, ver README). */
const SESION = { auth: 'cumpl_auth', email: 'cumpl_auth_email' };
function sesionActiva() {
  try { return sessionStorage.getItem(SESION.auth) === 'true'; } catch (e) { return false; }
}

/* Tema guardado: se aplica antes de pintar para evitar parpadeo. */
if (typeof document !== 'undefined') {
  if (!sesionActiva()) location.replace('index.html');
  try {
    const t = localStorage.getItem('tema');
    if (t === 'light' || t === 'dark') document.documentElement.dataset.theme = t;
  } catch (e) { /* sin almacenamiento */ }
  document.addEventListener('DOMContentLoaded', () => {
    UI.init();
    const salir = document.getElementById('btn-salir');
    try { salir.title = `Sesión: ${sessionStorage.getItem(SESION.email) || ''}`; } catch (e) { /* sin almacenamiento */ }
    salir.addEventListener('click', () => {
      try { Object.values(SESION).forEach(k => sessionStorage.removeItem(k)); } catch (e) { /* sin almacenamiento */ }
      location.replace('index.html');
    });
  });
}

