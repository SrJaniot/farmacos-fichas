// Fichas de dilución: buscador estático que lee datos/medicamentos.json
// (ese JSON se genera desde el Excel con herramientas/convertir.py).

const estado = {
  datos: null,
  porId: {},
  titulos: {},
  busqueda: '',
  filtro: 'todos',
};

const FILTROS = [
  { clave: 'todos', texto: 'Todos' },
  { clave: 'antibiotico', texto: 'Antibióticos' },
  { clave: 'otro', texto: 'Otros medicamentos' },
  { clave: 'electrolito', texto: 'Electrolitos' },
  { clave: 'periferica', texto: '⚠ Vía periférica con precaución' },
  { clave: 'subcutanea', texto: 'Uso subcutáneo' },
];

// --GRUPOS DE CAMPOS DE LA FICHA (LAS CLAVES VIENEN DEL CONVERTIDOR)
const GRUPOS = [
  { titulo: 'Reconstitución', campos: ['reconstitucion', 'estab_recon_amb', 'estab_recon_frio'] },
  { titulo: 'Dilución', campos: ['vol_normal', 'conc_estandar', 'vol_restringido', 'conc_restringido'] },
  { titulo: 'Estabilidad de la solución', campos: ['estab_sol_amb', 'estab_sol_frio'] },
  { titulo: 'Administración y dosis', campos: ['tiempo_infusion', 'dosis_mantenimiento', 'bolos', 'dosis_maxima'] },
  { titulo: 'Características fisicoquímicas', campos: ['osmolaridad', 'ph'] },
];

const GUIAS = [
  { id: 'soluciones', titulo: 'Soluciones y mezcla de potasio 1:1', texto: 'SSN 1/4, 1/2, 3 %, 7,5 %, solución polarizante y soluciones base.' },
  { id: 'insulina', titulo: 'Escala de insulina suplementaria', texto: 'Dosis según glucemia y sensibilidad a la insulina.' },
  { id: 'periferica', titulo: 'Vía periférica con precaución', texto: 'Normas y medicamentos que, si es posible, van por vía central.' },
  { id: 'subcutanea', titulo: 'Vía subcutánea (paliativos)', texto: 'Normas, medicamentos aptos y posología por vía subcutánea.' },
];

// --AUTORAS DE LA HERRAMIENTA, EN EL ORDEN EN QUE SE MUESTRAN
const AUTORAS = [
  { nombre: 'Mariana Sandoval Padilla', cargo: 'Enfermera PEP de Urgencias' },
  { nombre: 'Luna Mercedes Sandoval Silva', cargo: 'Enfermera PEP de Medicina Interna' },
];

// "<strong>Mariana…</strong>, Enfermera PEP de Urgencias; y <strong>Luna…</strong>, Enfermera PEP…"
function autorasHtml() {
  const partes = AUTORAS.map(a => `<strong>${esc(a.nombre)}</strong>${a.cargo ? `, ${esc(a.cargo)}` : ''}`);
  return partes.length > 1 ? `${partes.slice(0, -1).join('; ')}; y ${partes[partes.length - 1]}` : partes[0];
}

const app = document.getElementById('app');

// --UTILIDADES

function esc(texto) {
  return String(texto ?? '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function normalizar(texto) {
  return String(texto ?? '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
}

// Convierte las URL de un texto en enlaces (el texto ya viene escapado).
function enlazar(textoEscapado) {
  return textoEscapado.replace(/https?:\/\/[^\s<]+/g, url =>
    `<a href="${url}" target="_blank" rel="noopener" class="text-teal-700 underline break-all dark:text-teal-300">${url}</a>`);
}

function textoExcel(texto) {
  return `<div class="texto-excel">${enlazar(esc(texto))}</div>`;
}

function etiquetasDe(m) {
  const et = [];
  if (m.categoria === 'Antibiótico') et.push('<span class="etiqueta et-azul">Antibiótico</span>');
  if (m.categoria === 'Electrolito') et.push('<span class="etiqueta et-verde">Electrolito</span>');
  et.push(`<span class="etiqueta et-gris">${esc(m.forma)}</span>`);
  if (m.periferica_precaucion) et.push('<span class="etiqueta et-naranja" title="Si es posible, administrar por vía central">⚠ Periférica con precaución</span>');
  if (m.subcutanea) et.push('<span class="etiqueta et-violeta">Uso subcutáneo</span>');
  return et.join(' ');
}

function textoBuscable(m) {
  return normalizar([m.nombre, m.hoja, ...m.presentaciones.map(p => `${p.titulo} ${p.marca}`)].join(' '));
}

function filtrar() {
  const palabras = normalizar(estado.busqueda).split(/\s+/).filter(Boolean);
  return estado.datos.medicamentos.filter(m => {
    switch (estado.filtro) {
      case 'antibiotico': if (m.categoria !== 'Antibiótico') return false; break;
      case 'otro': if (m.categoria !== 'Otro') return false; break;
      case 'electrolito': if (m.categoria !== 'Electrolito') return false; break;
      case 'periferica': if (!m.periferica_precaucion) return false; break;
      case 'subcutanea': if (!m.subcutanea) return false; break;
    }
    return palabras.every(p => m._busqueda.includes(p));
  });
}

// --VISTA: INICIO (BUSCADOR)

function vistaInicio() {
  document.title = 'Fichas de dilución';
  app.innerHTML = `
    <section class="mb-6">
      <p class="text-xs font-bold uppercase tracking-widest text-teal-700 dark:text-teal-400">Consulta rápida</p>
      <h1 class="mt-1 text-2xl font-bold text-slate-900 sm:text-3xl dark:text-white">¿Cómo se diluye y administra?</h1>
      <p class="mt-1 text-slate-600 dark:text-slate-400">Busque por nombre del medicamento, concentración o laboratorio.</p>
      <p class="mt-2 text-sm text-slate-500 dark:text-slate-400">Herramienta elaborada por <span class="text-slate-700 dark:text-slate-200">${autorasHtml()}</span></p>
    </section>

    <div class="relative">
      <svg class="pointer-events-none absolute left-4 top-1/2 h-5 w-5 -translate-y-1/2 text-slate-400" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/></svg>
      <input id="buscar" type="search" autocomplete="off" spellcheck="false" placeholder="Ej: Noradrenalina, Vancomicina, Vitalis…"
        class="w-full rounded-2xl border border-slate-300 bg-white py-4 pl-12 pr-4 text-lg shadow-sm outline-none focus:border-teal-600 focus:ring-4 focus:ring-teal-600/15 dark:border-slate-700 dark:bg-slate-900"
        value="${esc(estado.busqueda)}" aria-label="Buscar medicamento">
    </div>

    <div id="filtros" class="mt-4 flex flex-wrap gap-2" role="group" aria-label="Filtros">
      ${FILTROS.map(f => `<button type="button" data-filtro="${f.clave}" class="chip ${f.clave === estado.filtro ? 'chip-activo' : ''}">${esc(f.texto)}</button>`).join('')}
    </div>

    <p id="contador" class="mt-4 text-sm text-slate-500 dark:text-slate-400" aria-live="polite"></p>
    <div id="resultados" class="mt-2 grid gap-3 sm:grid-cols-2"></div>

    ${avisoImportante()}

    <section class="mt-10">
      <h2 class="mb-3 text-lg font-bold text-slate-900 dark:text-white">Guías y tablas</h2>
      <div class="grid gap-3 sm:grid-cols-2">
        ${GUIAS.map(g => `
          <a href="#/guia/${g.id}" class="tarjeta block p-4 hover:border-teal-600 dark:hover:border-teal-500">
            <span class="font-semibold text-slate-900 dark:text-white">${esc(g.titulo)}</span>
            <span class="mt-1 block text-sm text-slate-500 dark:text-slate-400">${esc(g.texto)}</span>
          </a>`).join('')}
      </div>
    </section>`;

  const input = document.getElementById('buscar');
  input.addEventListener('input', () => { estado.busqueda = input.value; pintarResultados(); });
  input.addEventListener('keydown', e => {
    if (e.key === 'Enter') {
      const primero = document.querySelector('#resultados a');
      if (primero) primero.click();
    }
  });
  document.getElementById('filtros').addEventListener('click', e => {
    const btn = e.target.closest('[data-filtro]');
    if (!btn) return;
    estado.filtro = btn.dataset.filtro;
    document.querySelectorAll('#filtros .chip').forEach(b => b.classList.toggle('chip-activo', b === btn));
    pintarResultados();
  });
  pintarResultados();
  if (window.matchMedia('(min-width: 640px)').matches) input.focus();
}

function pintarResultados() {
  const lista = filtrar();
  const total = estado.datos.medicamentos.length;
  document.getElementById('contador').textContent =
    lista.length === total ? `${total} medicamentos` : `${lista.length} de ${total} medicamentos`;

  const cont = document.getElementById('resultados');
  if (!lista.length) {
    cont.innerHTML = `
      <div class="tarjeta col-span-full p-8 text-center">
        <p class="font-semibold text-slate-900 dark:text-white">No encontramos coincidencias</p>
        <p class="mt-1 text-sm text-slate-500 dark:text-slate-400">Pruebe con otro nombre, revise la ortografía o quite el filtro.</p>
      </div>`;
    return;
  }
  cont.innerHTML = lista.map(m => {
    const marcas = [...new Set(m.presentaciones.map(p => p.marca).filter(Boolean))];
    return `
      <a href="#/m/${m.id}" class="tarjeta group block p-4 hover:border-teal-600 dark:hover:border-teal-500">
        <span class="flex items-start justify-between gap-2">
          <span class="text-base font-bold text-slate-900 group-hover:text-teal-700 dark:text-white dark:group-hover:text-teal-300">${esc(m.nombre)}</span>
          <span class="text-slate-400 group-hover:text-teal-600" aria-hidden="true">›</span>
        </span>
        <span class="mt-2 flex flex-wrap gap-1">${etiquetasDe(m)}</span>
        ${m.presentaciones.length > 1 || marcas.length ? `<span class="mt-2 block text-xs text-slate-500 dark:text-slate-400">${m.presentaciones.length > 1 ? `${m.presentaciones.length} presentaciones` : 'Laboratorio'}${marcas.length ? ': ' + esc(marcas.join(', ')) : ''}</span>` : ''}
      </a>`;
  }).join('');
}

function avisoImportante() {
  return `
    <aside class="mt-8 rounded-2xl border border-amber-300 bg-amber-50 p-4 text-sm text-amber-900 dark:border-amber-800 dark:bg-amber-950/40 dark:text-amber-200">
      <strong>Importante:</strong> esta herramienta es de apoyo y transcribe la tabla de dilución institucional.
      No reemplaza la prescripción médica ni la validación del químico farmacéutico. Ante cualquier duda,
      confirme con el servicio farmacéutico y con la ficha técnica del laboratorio disponible.
    </aside>`;
}

// --VISTA: FICHA DE UN MEDICAMENTO

function vistaFicha(id, indicePres) {
  const m = estado.porId[id];
  if (!m) { vistaNoEncontrada(); return; }
  document.title = `${m.nombre} · Fichas de dilución`;
  const i = Math.min(Math.max(parseInt(indicePres, 10) || 0, 0), m.presentaciones.length - 1);
  const p = m.presentaciones[i];
  const d = p.datos;

  const pestanas = m.presentaciones.length > 1 ? `
    <div class="mt-5">
      <p class="mb-2 text-xs font-semibold uppercase tracking-wider text-slate-500">Presentación / laboratorio</p>
      <div class="flex flex-wrap gap-2" role="tablist">
        ${m.presentaciones.map((x, j) => `
          <a href="#/m/${m.id}/${j}" role="tab" aria-selected="${j === i}" class="chip ${j === i ? 'chip-activo' : ''}">${esc(nombrePestana(m, j))}</a>`).join('')}
      </div>
    </div>` : '';

  const alertas = [];
  if (d.no_compatibles) alertas.push(`
    <div class="rounded-2xl border border-red-300 bg-red-50 p-4 dark:border-red-900 dark:bg-red-950/40">
      <h3 class="text-sm font-bold uppercase tracking-wide text-red-800 dark:text-red-300">✕ No compatibles para mezcla o infusión en Y</h3>
      <div class="mt-2 text-red-950 dark:text-red-100">${textoIncompatibles(d.no_compatibles)}</div>
    </div>`);
  if (d.soluciones_compatibles) alertas.push(`
    <div class="rounded-2xl border border-emerald-300 bg-emerald-50 p-4 dark:border-emerald-900 dark:bg-emerald-950/40">
      <h3 class="text-sm font-bold uppercase tracking-wide text-emerald-800 dark:text-emerald-300">✓ Soluciones compatibles para dilución</h3>
      <div class="mt-2 text-emerald-950 dark:text-emerald-100">${textoExcel(d.soluciones_compatibles)}</div>
    </div>`);

  const grupos = GRUPOS.map(g => {
    const filas = g.campos.filter(c => d[c]);
    if (!filas.length) return '';
    return `
      <section class="tarjeta overflow-hidden">
        <h3 class="border-b border-slate-200 bg-slate-50 px-4 py-2 text-sm font-bold text-slate-700 dark:border-slate-800 dark:bg-slate-800/50 dark:text-slate-200">${esc(g.titulo)}</h3>
        <dl class="divide-y divide-slate-100 dark:divide-slate-800">
          ${filas.map(c => `
            <div class="grid gap-1 px-4 py-3 sm:grid-cols-[minmax(0,2fr)_minmax(0,3fr)] sm:gap-4">
              <dt class="text-sm text-slate-500 dark:text-slate-400">${esc(estado.titulos[c])}</dt>
              <dd class="font-medium text-slate-900 dark:text-slate-100">${textoExcel(d[c])}</dd>
            </div>`).join('')}
        </dl>
      </section>`;
  }).join('');

  const bloque = (campo, clases) => d[campo] ? `
    <section class="rounded-2xl border p-4 ${clases}">
      <h3 class="text-sm font-bold">${esc(estado.titulos[campo])}</h3>
      <div class="mt-2">${textoExcel(d[campo])}</div>
    </section>` : '';

  const notaPeriferica = m.periferica_precaucion ? `
    <div class="rounded-2xl border border-orange-300 bg-orange-50 p-4 text-sm text-orange-900 dark:border-orange-900 dark:bg-orange-950/40 dark:text-orange-200">
      <strong>⚠ Vía periférica con precaución.</strong> De ser posible, use una vía central. Si no, monitoree el acceso venoso periférico por la alta probabilidad de flebitis.
      <a href="#/guia/periferica" class="underline">Ver normas</a>.
    </div>` : '';

  const notaSubcutanea = m.subcutanea ? `
    <div class="rounded-2xl border border-violet-300 bg-violet-50 p-4 text-sm text-violet-900 dark:border-violet-900 dark:bg-violet-950/40 dark:text-violet-200">
      Este medicamento aparece en la lista de uso por <strong>vía subcutánea (paliativos)</strong>.
      <a href="#/guia/subcutanea" class="underline">Ver posología y normas</a>.
    </div>` : '';

  const notas = m.notas.length ? `
    <section class="tarjeta p-4">
      <h3 class="text-sm font-bold text-slate-700 dark:text-slate-200">Notas de la hoja</h3>
      <ul class="mt-2 list-disc space-y-1 pl-5 text-sm text-slate-600 dark:text-slate-300">
        ${m.notas.map(n => `<li>${textoExcel(n)}</li>`).join('')}
      </ul>
    </section>` : '';

  app.innerHTML = `
    <a href="#/" class="btn-suave mb-4">‹ Volver al buscador</a>
    <article>
      <header>
        <h1 class="text-2xl font-bold text-slate-900 sm:text-3xl dark:text-white">${esc(m.nombre)}</h1>
        <p class="mt-1 text-slate-600 dark:text-slate-400">${esc(p.titulo)}</p>
        <div class="mt-3 flex flex-wrap gap-1">${etiquetasDe(m)}</div>
      </header>
      ${pestanas}
      <div class="mt-6 grid gap-4">
        ${alertas.join('')}
        ${notaPeriferica}
        <div class="grid gap-4 lg:grid-cols-2">${grupos}</div>
        ${bloque('precauciones', 'border-amber-300 bg-amber-50 text-amber-950 dark:border-amber-900 dark:bg-amber-950/40 dark:text-amber-100')}
        ${bloque('efectos', 'border-slate-200 bg-white text-slate-800 dark:border-slate-800 dark:bg-slate-900 dark:text-slate-100')}
        ${bloque('observaciones', 'border-slate-200 bg-white text-slate-800 dark:border-slate-800 dark:bg-slate-900 dark:text-slate-100')}
        ${notaSubcutanea}
        ${notas}
      </div>
    </article>
    ${avisoImportante()}`;
  window.scrollTo(0, 0);
}

function nombrePestana(m, j) {
  const marca = m.presentaciones[j].marca;
  if (!marca) return `Presentación ${j + 1}`;
  const repetidas = m.presentaciones.filter(x => x.marca === marca).length;
  if (repetidas === 1) return marca;
  const n = m.presentaciones.slice(0, j + 1).filter(x => x.marca === marca).length;
  return `${marca} (${n})`;
}

// Cada medicamento de la lista de incompatibles que tenga ficha queda como enlace a la búsqueda.
function textoIncompatibles(texto) {
  return `<div class="texto-excel">${texto.split(/(,|\n)/).map(parte => {
    const limpio = parte.trim();
    if (!limpio || parte === ',' || parte === '\n') return esc(parte);
    const palabra = normalizar(limpio).split(/[\s/(]+/)[0];
    const existe = palabra.length >= 5 && estado.datos.medicamentos.some(m => normalizar(m.nombre).startsWith(palabra.slice(0, 6)));
    if (!existe) return esc(parte);
    const espacio = parte.match(/^\s*/)[0];
    return `${esc(espacio)}<a href="#/?q=${encodeURIComponent(limpio.split(/[\s/(]+/)[0])}" class="underline decoration-red-300 underline-offset-2 hover:decoration-red-700">${esc(limpio)}</a>`;
  }).join('')}</div>`;
}

// --VISTAS: GUIAS

function vistaGuia(id) {
  const vistas = { soluciones: guiaSoluciones, insulina: guiaInsulina, periferica: guiaPeriferica, subcutanea: guiaSubcutanea, uso: guiaUso };
  const fn = vistas[id];
  if (!fn) { vistaNoEncontrada(); return; }
  const { titulo, html } = fn();
  document.title = `${titulo} · Fichas de dilución`;
  app.innerHTML = `
    <a href="#/" class="btn-suave mb-4">‹ Volver al buscador</a>
    <h1 class="text-2xl font-bold text-slate-900 sm:text-3xl dark:text-white">${esc(titulo)}</h1>
    <div class="mt-6 grid gap-4">${html}</div>
    ${avisoImportante()}`;
  window.scrollTo(0, 0);
}

function tabla(encabezados, filas) {
  return `
    <div class="tarjeta overflow-x-auto">
      <table class="tabla">
        <thead><tr>${encabezados.map(h => `<th>${esc(h)}</th>`).join('')}</tr></thead>
        <tbody>${filas.map(f => `<tr>${f.map(c => `<td>${textoExcel(c)}</td>`).join('')}</tr>`).join('')}</tbody>
      </table>
    </div>`;
}

function guiaSoluciones() {
  const s = estado.datos.soluciones;
  return {
    titulo: 'Soluciones',
    html: `
      <section>
        <h2 class="mb-2 font-bold">Soluciones salinas</h2>
        ${tabla(['Solución', 'Dilución', 'Osmolaridad (mOsmol/L)'], s.salinas.map(x => [x.solucion, x.dilucion, x.osmolaridad]))}
        ${s.precaucion_salinas ? `<p class="mt-2 rounded-xl bg-amber-50 p-3 text-sm text-amber-900 dark:bg-amber-950/40 dark:text-amber-200"><strong>Precauciones:</strong> ${esc(s.precaucion_salinas)}</p>` : ''}
      </section>
      <section>
        <h2 class="mb-2 font-bold">Solución polarizante</h2>
        ${tabla(['Campo', 'Valor'], s.polarizante.map(x => [x.campo, x.valor]))}
      </section>
      <section>
        <h2 class="mb-2 font-bold">Preparación mezcla de potasio 1:1</h2>
        ${tabla(s.potasio.encabezados, s.potasio.filas)}
        ${s.potasio.nota ? `<p class="mt-2 rounded-xl bg-red-50 p-3 text-sm font-semibold text-red-900 dark:bg-red-950/40 dark:text-red-200">⚠ ${esc(s.potasio.nota)} (catéter venoso central).</p>` : ''}
      </section>
      <section>
        <h2 class="mb-2 font-bold">Soluciones base</h2>
        ${tabla(['Solución', 'Osmolaridad', 'pH'], s.base.map(x => [x.solucion, x.osmolaridad, x.ph]))}
      </section>`,
  };
}

function guiaInsulina() {
  const s = estado.datos.insulina;
  return {
    titulo: 'Escala de insulina suplementaria',
    html: `
      <p class="text-slate-600 dark:text-slate-400">Dosis en unidades según la glucemia.</p>
      ${tabla(s.encabezados, s.filas)}
      ${s.nota ? `<p class="rounded-xl bg-sky-50 p-3 text-sm text-sky-900 dark:bg-sky-950/40 dark:text-sky-200">${esc(s.nota)}</p>` : ''}`,
  };
}

function enlaceSiExiste(nombre) {
  const palabra = normalizar(nombre).replace(/[^a-z ]/g, ' ').trim().split(/\s+/)[0] || '';
  const m = palabra.length >= 5 && estado.datos.medicamentos.find(x => normalizar(x.nombre).startsWith(palabra.slice(0, 6)));
  return m ? `<a href="#/?q=${encodeURIComponent(palabra)}" class="text-teal-700 underline dark:text-teal-300">${esc(nombre)}</a>` : esc(nombre);
}

function guiaPeriferica() {
  const s = estado.datos.periferica;
  return {
    titulo: 'Vía periférica con precaución',
    html: `
      <section class="tarjeta p-4">
        <h2 class="font-bold">Normas</h2>
        <p class="mt-1 text-sm text-slate-500 dark:text-slate-400">Se administran con precaución por vía periférica los medicamentos que cumplen una o varias de estas condiciones:</p>
        <ol class="mt-2 list-decimal space-y-1 pl-5">${s.normas.map(n => `<li>${esc(n)}</li>`).join('')}</ol>
      </section>
      ${s.nota ? `<section class="rounded-2xl border border-orange-300 bg-orange-50 p-4 text-orange-950 dark:border-orange-900 dark:bg-orange-950/40 dark:text-orange-100">${textoExcel(s.nota.replace('se identificarán con un sombreado naranja', 'se identifican en las fichas con la etiqueta naranja ⚠').replace('Este sombreado', 'Esta etiqueta'))}</section>` : ''}
      <section class="tarjeta p-4">
        <h2 class="font-bold">Medicamentos</h2>
        <ul class="mt-2 grid gap-1 sm:grid-cols-2">${s.medicamentos.map(n => `<li>⚠ ${enlaceSiExiste(n)}</li>`).join('')}</ul>
      </section>`,
  };
}

function guiaSubcutanea() {
  const s = estado.datos.subcutanea;
  return {
    titulo: 'Vía subcutánea (paliativos)',
    html: `
      <section class="tarjeta p-4">
        <h2 class="font-bold">Normas</h2>
        <ol class="mt-2 list-decimal space-y-2 pl-5">${s.normas.map(n => `<li>${textoExcel(n)}</li>`).join('')}</ol>
      </section>
      <section>
        <h2 class="mb-2 font-bold">Medicamentos utilizados por vía subcutánea</h2>
        ${tabla(['Fármaco', 'Indicación', 'Posología'], s.tabla.map(x => [x.farmaco, x.indicacion, x.posologia]))}
      </section>
      <section class="tarjeta p-4">
        <h2 class="font-bold">Lista de medicamentos aptos</h2>
        <ul class="mt-2 space-y-1">${s.aptos.map(a => `
          <li>${esc(a.nombre)}${a.nota ? ` <span class="etiqueta et-gris">${esc(a.nota)}</span>` : ''}${a.enlace ? ` · <a href="${esc(a.enlace)}" target="_blank" rel="noopener" class="text-sm text-teal-700 underline dark:text-teal-300">ficha técnica</a>` : ''}</li>`).join('')}
        </ul>
      </section>`,
  };
}

function guiaUso() {
  return {
    titulo: '¿Cómo usarlo?',
    html: `
      <section class="tarjeta space-y-3 p-5 leading-relaxed">
        <p><strong>1. Busque</strong> el medicamento por nombre, concentración (ej. <em>1 g</em>) o laboratorio (ej. <em>Vitalis</em>). No importan tildes ni mayúsculas.</p>
        <p><strong>2. Filtre</strong> si quiere ver solo antibióticos, electrolitos, los que requieren precaución por vía periférica o los de uso subcutáneo.</p>
        <p><strong>3. Abra la ficha.</strong> Arriba verá en rojo los medicamentos <strong>no compatibles</strong> y en verde las <strong>soluciones compatibles</strong>. Luego la reconstitución, la dilución, la estabilidad, la administración y las dosis.</p>
        <p><strong>4. Elija el laboratorio.</strong> Si hay varias marcas, cada una tiene su pestaña: la estabilidad y el tiempo de infusión pueden cambiar entre marcas.</p>
        <p><strong>Etiquetas:</strong>
          <span class="etiqueta et-naranja">⚠ Periférica con precaución</span> de ser posible, use vía central;
          <span class="etiqueta et-violeta">Uso subcutáneo</span> aparece en la lista de paliativos.</p>
        <p><strong>Abreviaturas frecuentes:</strong> SSN = solución salina normal 0,9 %; DAD = dextrosa en agua destilada; LR = lactato de Ringer; API / A.E.P.I. = agua estéril para inyección; T amb = temperatura ambiente; CVC = catéter venoso central.</p>
        <p><strong>Botón A+</strong>: agranda la letra (se repite el ciclo al pulsarlo varias veces).</p>
        <p>Puede guardar el enlace de una ficha o agregar la página a la pantalla de inicio del celular.</p>
        <p class="border-t border-slate-200 pt-3 text-sm text-slate-500 dark:border-slate-800 dark:text-slate-400">Herramienta elaborada por ${autorasHtml()}, con base en la tabla de dilución de medicamentos de uso hospitalario.</p>
      </section>`,
  };
}

function vistaNoEncontrada() {
  app.innerHTML = `
    <div class="tarjeta p-8 text-center">
      <p class="font-semibold">No encontramos esa página.</p>
      <a href="#/" class="btn-suave mt-4">‹ Volver al buscador</a>
    </div>`;
}

// --PIE DE PAGINA CON LA FUENTE

function pintarPie() {
  const f = estado.datos.fuente || {};
  document.getElementById('pie').innerHTML = `
    <p class="text-base text-slate-700 dark:text-slate-200">Herramienta elaborada por ${autorasHtml()}</p>
    <p><strong class="text-slate-700 dark:text-slate-200">Fuente:</strong> ${esc(f.titulo || 'Tabla de dilución medicamentos de uso hospitalario')}${f.aprobado ? ` · ${esc(f.aprobado)}` : ''}.</p>
    ${f.autores?.length ? `<p>Tabla elaborada por ${esc(f.autores.join(', ').replace(/\.?$/, '.'))}</p>` : ''}
    ${f.referencias?.length ? `<p>Referencias: ${f.referencias.map(r => enlazar(esc(r))).join(' · ')}</p>` : ''}
    <p>Datos tomados de «${esc(estado.datos.archivo_fuente)}», convertidos el ${esc(estado.datos.generado)}.</p>`;
}

// --TAMAÑO DE LETRA (SE RECUERDA EN ESTE NAVEGADOR)

const TAMANOS = ['100%', '112.5%', '125%'];
function aplicarLetra(i) {
  document.documentElement.style.fontSize = TAMANOS[i];
  try { localStorage.setItem('fichas-letra', String(i)); } catch (e) { /* sin almacenamiento */ }
}
let letra = 0;
try { letra = Math.min(parseInt(localStorage.getItem('fichas-letra') || '0', 10) || 0, TAMANOS.length - 1); } catch (e) { letra = 0; }
aplicarLetra(letra);
document.getElementById('btn-letra').addEventListener('click', () => { letra = (letra + 1) % TAMANOS.length; aplicarLetra(letra); });

// --RUTAS (#/  #/?q=texto  #/m/<id>[/<presentacion>]  #/guia/<id>)

function enrutar() {
  const hash = location.hash.replace(/^#/, '') || '/';
  const [ruta, consulta] = hash.split('?');
  const partes = ruta.split('/').filter(Boolean);
  if (partes[0] === 'm') return vistaFicha(decodeURIComponent(partes[1] || ''), partes[2]);
  if (partes[0] === 'guia') return vistaGuia(partes[1]);
  if (consulta) {
    const q = new URLSearchParams(consulta).get('q');
    if (q !== null) { estado.busqueda = q; estado.filtro = 'todos'; }
  }
  vistaInicio();
}

async function iniciar() {
  try {
    const resp = await fetch('datos/medicamentos.json', { cache: 'no-cache' });
    if (!resp.ok) throw new Error(resp.status);
    estado.datos = await resp.json();
  } catch (e) {
    app.innerHTML = `<div class="tarjeta p-8 text-center"><p class="font-semibold">No se pudieron cargar las fichas.</p><p class="mt-1 text-sm text-slate-500">Revise su conexión y recargue la página.</p></div>`;
    return;
  }
  estado.datos.medicamentos.forEach(m => { m._busqueda = textoBuscable(m); estado.porId[m.id] = m; });
  estado.datos.campos.forEach(c => { estado.titulos[c.clave] = c.titulo; });
  pintarPie();
  window.addEventListener('hashchange', enrutar);
  enrutar();
}

iniciar();
