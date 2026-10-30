// Dashboard: mismas vistas que las tablas dinámicas del Excel (Prom-Max, Info semanal,
// Rend x persona, Revisión x días) más tendencia y avance por bloque.
const Dashboard = (() => {
  const $ = id => document.getElementById(id);
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const fmt = (x, d = 0) => Number(x || 0).toLocaleString('es-CO', { maximumFractionDigits: d, minimumFractionDigits: d });
  const fechaStr = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  const corto = (s, n) => (s.length > n ? `${s.slice(0, n - 1)}…` : s);
  const numSort = (a, b) => parseFloat(String(a).replace(/[^\d.]/g, '')) - parseFloat(String(b).replace(/[^\d.]/g, '')) || String(a).localeCompare(String(b));

  let ctx = null, servidor = null, enlazado = false, laborPersona = '', laborTend = '', laborBloque = '';

  // ---------- Datos ----------
  function rango(tipo) {
    const hoy = new Date();
    const h = fechaStr(hoy);
    if (tipo === '7') { const d = new Date(hoy); d.setDate(d.getDate() - 6); return [fechaStr(d), h]; }
    if (tipo === '30') { const d = new Date(hoy); d.setDate(d.getDate() - 29); return [fechaStr(d), h]; }
    if (tipo === 'mes') return [fechaStr(new Date(hoy.getFullYear(), hoy.getMonth(), 1)), h];
    const d = new Date(hoy); d.setDate(d.getDate() - ((d.getDay() + 6) % 7));
    return [fechaStr(d), h];
  }

  // ---------- Filtro por año, mes y semana (ISO, como la columna SEM) ----------
  const MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
  const utc = s => { const [y, m, d] = s.split('-').map(Number); return new Date(Date.UTC(y, m - 1, d)); };
  const isoStr = d => d.toISOString().slice(0, 10);
  const masDias = (s, n) => { const d = utc(s); d.setUTCDate(d.getUTCDate() + n); return isoStr(d); };
  function lunesSemana1(y) { const d = new Date(Date.UTC(y, 0, 4)); d.setUTCDate(d.getUTCDate() - ((d.getUTCDay() + 6) % 7)); return isoStr(d); }
  const semanasDelAnio = y => Calc.semanaISO(`${y}-12-28`);
  const diaCorto = s => `${+s.slice(8)} ${MESES[+s.slice(5, 7) - 1].slice(0, 3).toLowerCase()}`;

  function pintarSemanas() {
    const y = +$('fAnio').value, m = $('fMes').value, cur = $('fSemana').value;
    let ops = [];
    if (y) {
      const l1 = lunesSemana1(y);
      for (let w = 1; w <= semanasDelAnio(y); w++) {
        const ini = masDias(l1, 7 * (w - 1)), fin = masDias(ini, 6);
        const mesIni = `${y}-${m.padStart(2, '0')}-01`, mesFin = isoStr(new Date(Date.UTC(y, +m, 0)));
        if (m && (fin < mesIni || ini > mesFin)) continue;
        ops.push([String(w), `Sem ${w} · ${diaCorto(ini)} a ${diaCorto(fin)}`]);
      }
    }
    opciones($('fSemana'), ops, 'Todas');
    if (!ops.some(([v]) => v === cur)) $('fSemana').value = '';
  }

  function pintarPeriodo() {
    const y = new Date().getFullYear();
    opciones($('fAnio'), [y + 1, y, y - 1, y - 2, y - 3].map(a => [String(a), String(a)]), 'Elegir');
    opciones($('fMes'), MESES.map((n, i) => [String(i + 1), n]), 'Todos');
    pintarSemanas();
  }

  // Convierte la selección de año, mes y semana en el rango Desde–Hasta.
  function rangoPeriodo() {
    if (!$('fAnio').value && ($('fMes').value || $('fSemana').value)) $('fAnio').value = String(new Date().getFullYear());
    const y = +$('fAnio').value, m = +$('fMes').value, w = +$('fSemana').value;
    if (!y) return null;
    if (w) { const ini = masDias(lunesSemana1(y), 7 * (w - 1)); return [ini, masDias(ini, 6)]; }
    if (m) return [`${y}-${String(m).padStart(2, '0')}-01`, isoStr(new Date(Date.UTC(y, m, 0)))];
    return [`${y}-01-01`, `${y}-12-31`];
  }

  const limpiarPeriodo = () => { $('fAnio').value = ''; $('fMes').value = ''; pintarSemanas(); };

  async function filasCombinadas() {
    const locales = await ctx.filasLocales();
    const borrados = new Set(await DB.getMeta('borrados', []));
    const porId = new Map();
    (servidor || []).forEach(f => { if (!borrados.has(f['ID REGISTRO'])) porId.set(f['ID FILA'], f); });
    // Los registros de este celular reemplazan los del servidor: son la versión más reciente aquí.
    const regsLocales = new Set(locales.map(f => f['ID REGISTRO']));
    for (const [k, f] of porId) if (regsLocales.has(f['ID REGISTRO'])) porId.delete(k);
    locales.forEach(f => porId.set(f['ID FILA'], f));
    return [...porId.values()].map(f => ({ ...f, FECHA: String(f.FECHA).slice(0, 10) }));
  }

  async function traerServidor() {
    const s = await DB.getMeta('sync', { url: '' });
    if (!s.url) { $('dashSource').textContent = 'Mostrando datos de este celular (sin servidor configurado).'; return; }
    if (!navigator.onLine) { $('dashSource').textContent = 'Sin señal: mostrando datos de este celular.'; return; }
    $('dashSource').textContent = 'Consultando servidor…';
    try {
      servidor = await Sync.filas($('fDesde').value, $('fHasta').value);
      $('dashSource').textContent = `Servidor y celular combinados · ${new Date().toLocaleTimeString('es-CO', { hour: '2-digit', minute: '2-digit' })}`;
    } catch (e) {
      $('dashSource').textContent = `No se pudo consultar el servidor (${e.message}); mostrando datos de este celular.`;
    }
  }

  function filtrar(filas) {
    const d = $('fDesde').value, h = $('fHasta').value, l = $('fLabor').value, p = $('fPersona').value, b = $('fBloque').value, g = $('fLugar').value;
    const a = $('fArea').value, li = $('fLider').value;
    return filas.filter(f => (!d || f.FECHA >= d) && (!h || f.FECHA <= h) && (!l || f.LABOR === l) && (!p || String(f.CEDULA) === p) && (!b || f.BLOQUE === b) && (!g || lugar(f) === g)
      && (!a || areaDe(f) === a) && (!li || liderDe(f) === li));
  }

  const tipo = f => Calc.tipoMedida(f.MEDIDA);
  // Las filas de antes del módulo de áreas eran todas del cultivo San Jacinto.
  const areaDe = f => Calc.AREA_ANTES[f.AREA] || f.AREA || Calc.AREA_BASE;
  const liderDe = f => String(f.LIDER || '').trim().toUpperCase();
  // La columna se llamó LUGARDE TRABAJO en la primera versión de la hoja.
  const lugar = f => f['LUGAR DE TRABAJO'] ?? f['LUGARDE TRABAJO'] ?? '';
  const unidad = t => `${t}/h`;

  function opciones(sel, valores, todos) {
    const cur = sel.value;
    sel.innerHTML = `<option value="">${todos}</option>` + valores.map(([v, t]) => `<option value="${esc(v)}">${esc(t)}</option>`).join('');
    if (valores.some(([v]) => v === cur)) sel.value = cur;
  }

  // ---------- Piezas de gráficos ----------
  function encabezado(titulo, sub, extra = '') {
    return `<div style="display:flex;gap:12px;align-items:flex-start;justify-content:space-between;flex-wrap:wrap"><div><h3 class="chart-title">${esc(titulo)}</h3><p class="chart-sub">${sub}</p></div>${extra}</div>`;
  }

  function selectorLabor(id, labores, actual) {
    return `<select id="${id}" style="min-height:38px;font-size:14px;padding:6px 8px;width:auto;max-width:240px">${labores.map(l => `<option ${l === actual ? 'selected' : ''}>${esc(l)}</option>`).join('')}</select>`;
  }

  function ejeTicks(max) {
    if (max <= 0) return [0];
    const paso0 = max / 4, mag = 10 ** Math.floor(Math.log10(paso0));
    const paso = [1, 2, 2.5, 5, 10].map(m => m * mag).find(s => s >= paso0);
    const out = [];
    for (let v = 0; v <= max + paso * 0.001; v += paso) out.push(v);
    if (out.at(-1) < max) out.push(out.at(-1) + paso);
    return out;
  }

  // Barras horizontales. filas: [{label, valor, marca?, tip, texto}]
  function barrasH(filas, { ancho = 640, etiqueta = 200, color = 'var(--series-1)', dec = 0, maxFijo = null, margen = 110 } = {}) {
    if (!filas.length) return '<div class="empty">Sin datos en este filtro.</div>';
    etiqueta = Math.min(etiqueta, Math.round(ancho * 0.38));
    const alto = 26, gap = 8, top = 22, x0 = etiqueta + 8, xMax = ancho - Math.min(margen, ancho * 0.3);
    const maxCar = Math.floor(etiqueta / 7.8);
    const max = maxFijo ?? Math.max(...filas.map(f => Math.max(f.valor, f.marca || 0)));
    const ticks = ejeTicks(max);
    const escala = v => x0 + (v / ticks.at(-1)) * (xMax - x0);
    const H = top + filas.length * (alto + gap) + 4;
    const grid = ticks.map(t => `<line class="grid" x1="${escala(t)}" x2="${escala(t)}" y1="${top - 6}" y2="${H}"/><text x="${escala(t)}" y="12" text-anchor="middle">${fmt(t, t < 10 && t % 1 ? 1 : 0)}</text>`).join('');
    const barras = filas.map((f, i) => {
      const y = top + i * (alto + gap);
      const w = Math.max(2, escala(f.valor) - x0);
      const marca = f.marca != null ? `<line x1="${escala(f.marca)}" x2="${escala(f.marca)}" y1="${y - 3}" y2="${y + alto + 3}" stroke="var(--text)" stroke-width="2.5" stroke-linecap="round"/>` : '';
      const fin = Math.max(escala(f.valor), f.marca != null ? escala(f.marca) : 0);
      return `<g data-tip="${esc(f.tip)}">
        <rect x="0" y="${y - gap / 2}" width="${ancho}" height="${alto + gap}" fill="transparent"/>
        <text class="lbl" x="${etiqueta}" y="${y + alto / 2 + 4}" text-anchor="end">${esc(corto(f.label, maxCar))}</text>
        ${w > 8
    ? `<path d="M${x0} ${y} h${w - 4} a4 4 0 0 1 4 4 v${alto - 8} a4 4 0 0 1 -4 4 h-${w - 4} z" fill="${f.color || color}"/>`
    : `<rect x="${x0}" y="${y}" width="${w}" height="${alto}" fill="${f.color || color}"/>`}
        ${marca}
        <text class="lbl" x="${fin + 6}" y="${y + alto / 2 + 4}">${f.texto ?? fmt(f.valor, dec)}</text>
      </g>`;
    }).join('');
    return `<svg class="chart" viewBox="0 0 ${ancho} ${H}" role="img">${grid}<line class="axis" x1="${x0}" x2="${x0}" y1="${top - 6}" y2="${H}"/>${barras}</svg>`;
  }

  function linea(puntos, { ancho = 900, alto = 240, unidadTxt = '' } = {}) {
    if (!puntos.length) return '<div class="empty">Sin datos en este filtro.</div>';
    const l = 48, r = 16, t = 16, b = 32;
    const max = Math.max(...puntos.map(p => p.valor));
    const ticks = ejeTicks(max);
    const X = i => (puntos.length === 1 ? (l + ancho - r) / 2 : l + (i / (puntos.length - 1)) * (ancho - l - r));
    const Y = v => t + (1 - v / ticks.at(-1)) * (alto - t - b);
    const grid = ticks.map(v => `<line class="grid" x1="${l}" x2="${ancho - r}" y1="${Y(v)}" y2="${Y(v)}"/><text x="${l - 6}" y="${Y(v) + 4}" text-anchor="end">${fmt(v, v % 1 ? 1 : 0)}</text>`).join('');
    const salto = Math.ceil(puntos.length / Math.max(3, Math.floor(ancho / 80)));
    const xs = puntos.map((p, i) => (i % salto === 0 || i === puntos.length - 1 ? `<text x="${X(i)}" y="${alto - 10}" text-anchor="middle">${esc(p.etiqueta)}</text>` : '')).join('');
    const path = puntos.map((p, i) => `${i ? 'L' : 'M'}${X(i)} ${Y(p.valor)}`).join(' ');
    const dots = puntos.map((p, i) => `<g data-tip="${esc(p.tip)}"><rect x="${X(i) - 14}" y="${t}" width="28" height="${alto - t - b}" fill="transparent"/><circle cx="${X(i)}" cy="${Y(p.valor)}" r="4.5" fill="var(--series-1)" stroke="var(--surface)" stroke-width="2"/></g>`).join('');
    return `<svg class="chart" viewBox="0 0 ${ancho} ${alto}" role="img" aria-label="${esc(unidadTxt)}">${grid}<path d="${path}" fill="none" stroke="var(--series-1)" stroke-width="2" stroke-linejoin="round"/>${dots}${xs}</svg>`;
  }

  // ---------- Vistas ----------
  function kpis(filas) {
    const horas = filas.reduce((s, f) => s + Number(f['TIEMPO LAB'] || 0), 0);
    const personas = new Set(filas.map(f => f.CEDULA)).size;
    const plantas = filas.filter(f => tipo(f) === 'plantas').reduce((s, f) => s + Number(f.CANTIDAD || 0), 0);
    const tallos = filas.filter(f => tipo(f) === 'tallos').reduce((s, f) => s + Number(f.CANTIDAD || 0), 0);
    const unidades = filas.filter(f => tipo(f) === 'unidades').reduce((s, f) => s + Number(f.CANTIDAD || 0), 0);
    const hMedidas = filas.filter(f => tipo(f) !== 'tiempo').reduce((s, f) => s + Number(f['TIEMPO LAB'] || 0), 0);
    const dias = new Set(filas.map(f => f.FECHA)).size;
    $('kpis').innerHTML = [
      ['Horas laboradas', fmt(horas, 1), `${dias} días con registros`],
      ['Personas', fmt(personas), `${filas.length} filas de rendimiento`],
      ['Plantas trabajadas', fmt(plantas), 'labores medidas en plantas'],
      ['Tallos', fmt(tallos), 'labores medidas en tallos'],
      unidades ? ['Unidades', fmt(unidades), 'labores medidas en unidades'] : null,
      ['Tiempo en labores medidas', horas ? `${fmt((hMedidas / horas) * 100)}%` : '—', 'plantas, tallos y unidades sobre el total'],
    ].filter(Boolean).map(([l, v, s]) => `<div class="kpi"><div class="k-l">${l}</div><div class="k-v">${v}</div><div class="k-s">${s}</div></div>`).join('');
  }

  function promMax(filas) {
    const grupos = {};
    filas.filter(f => tipo(f) !== 'tiempo' && Number(f['TIEMPO LAB']) > 0).forEach(f => {
      const g = grupos[f.LABOR] ||= { labor: f.LABOR, t: tipo(f), suma: 0, n: 0, max: 0, cant: 0, horas: 0 };
      const r = Number(f.RENDIMIENTO || 0);
      g.suma += r; g.n++; g.max = Math.max(g.max, r); g.cant += Number(f.CANTIDAD || 0); g.horas += Number(f['TIEMPO LAB'] || 0);
    });
    const porUnidad = t => Object.values(grupos).filter(g => g.t === t).sort((a, b) => b.suma / b.n - a.suma / a.n).map(g => ({
      label: g.labor, valor: g.suma / g.n, marca: g.max,
      texto: `${fmt(g.suma / g.n)} · máx ${fmt(g.max)}`,
      tip: `${g.labor}\nPromedio ${fmt(g.suma / g.n, 1)} ${unidad(g.t)}\nMáximo ${fmt(g.max, 1)} ${unidad(g.t)}\n${g.n} registros · ${fmt(g.cant)} ${g.t} en ${fmt(g.horas, 1)} h`,
    }));
    const pl = porUnidad('plantas'), ta = porUnidad('tallos'), un = porUnidad('unidades');
    $('cPromMax').innerHTML = encabezado('Rendimiento promedio y máximo por labor', 'Promedio de los rendimientos por persona en el período; la marca negra es el máximo.') +
      `<div class="legend"><span><i style="background:var(--series-1)"></i>Promedio</span><span><i class="tick"></i>Máximo</span></div>` +
      (pl.length || ta.length || un.length
        ? [[pl, 'Plantas por hora'], [ta, 'Tallos por hora'], [un, 'Unidades por hora']].filter(([x]) => x.length)
          .map(([x, t], i) => `<p class="chart-sub" style="margin:${i ? 12 : 8}px 0 4px"><b>${t}</b></p>${barrasH(x, { ancho: anchoDe('cPromMax'), etiqueta: 240, margen: 130 })}`).join('')
        : '<div class="empty">No hay labores medidas en plantas, tallos o unidades en este filtro.</div>');
  }

  function horasLabor(filas) {
    const g = {};
    filas.forEach(f => { g[f.LABOR] = (g[f.LABOR] || 0) + Number(f['TIEMPO LAB'] || 0); });
    const total = Object.values(g).reduce((a, b) => a + b, 0);
    let lista = Object.entries(g).sort((a, b) => b[1] - a[1]);
    if (lista.length > 12) {
      const resto = lista.slice(11).reduce((s, x) => s + x[1], 0);
      lista = [...lista.slice(0, 11), [`Otras (${lista.length - 11})`, resto]];
    }
    $('cHorasLabor').innerHTML = encabezado('Horas por labor', 'Tiempo laborado en el período y su porcentaje del total.') +
      barrasH(lista.map(([l, h]) => ({ label: l, valor: h, texto: `${fmt(h, 1)} h · ${fmt((h / total) * 100)}%`, tip: `${l}\n${fmt(h, 2)} horas\n${fmt((h / total) * 100, 1)}% del tiempo` })), { ancho: anchoDe('cHorasLabor'), etiqueta: 200 });
  }

  function laboresMedidas(filas) {
    const c = {};
    filas.filter(f => tipo(f) !== 'tiempo').forEach(f => { c[f.LABOR] = (c[f.LABOR] || 0) + 1; });
    return Object.entries(c).sort((a, b) => b[1] - a[1]).map(x => x[0]);
  }

  function porPersona(filas) {
    const labores = laboresMedidas(filas);
    if (!labores.includes(laborPersona)) laborPersona = labores[0] || '';
    const fl = filas.filter(f => f.LABOR === laborPersona && Number(f['TIEMPO LAB']) > 0);
    const g = {};
    fl.forEach(f => {
      const p = g[f.CEDULA] ||= { nombre: f.NOMBRE, cant: 0, horas: 0, suma: 0, n: 0 };
      p.cant += Number(f.CANTIDAD || 0); p.horas += Number(f['TIEMPO LAB'] || 0); p.suma += Number(f.RENDIMIENTO || 0); p.n++;
    });
    const t = fl[0] ? tipo(fl[0]) : 'plantas';
    const lista = Object.values(g).map(p => ({ ...p, rend: p.horas ? p.cant / p.horas : 0 })).sort((a, b) => b.rend - a.rend);
    const prom = lista.length ? lista.reduce((s, p) => s + p.rend, 0) / lista.length : 0;
    $('cPersona').innerHTML = encabezado('Rendimiento por persona', labores.length ? `${esc(unidad(t))} en el período (cantidad total ÷ horas). Promedio del grupo: ${fmt(prom, 1)}.` : 'Sin labores medidas.',
      labores.length ? selectorLabor('selPersona', labores, laborPersona) : '') +
      barrasH(lista.map(p => ({
        label: p.nombre, valor: p.rend, texto: fmt(p.rend, 1),
        color: p.rend < prom * 0.8 ? 'var(--warn)' : 'var(--series-1)',
        tip: `${p.nombre}\n${fmt(p.rend, 1)} ${unidad(t)}\n${fmt(p.cant)} ${t} en ${fmt(p.horas, 1)} h · ${p.n} registros${p.rend < prom * 0.8 ? '\nMás de 20% por debajo del promedio' : ''}`,
      })), { ancho: anchoDe('cPersona'), etiqueta: 210, margen: 50 }) +
      (lista.some(p => p.rend < prom * 0.8) ? '<div class="legend" style="margin-top:8px"><span><i style="background:var(--warn)"></i>Más de 20% por debajo del promedio del grupo</span></div>' : '');
    $('selPersona')?.addEventListener('change', e => { laborPersona = e.target.value; render(); });
  }

  function tendencia(filas) {
    const labores = laboresMedidas(filas);
    if (!labores.includes(laborTend)) laborTend = labores[0] || '';
    const fl = filas.filter(f => f.LABOR === laborTend && Number(f['TIEMPO LAB']) > 0);
    const dias = new Set(fl.map(f => f.FECHA));
    const semanal = ($('fHasta').value && $('fDesde').value && (Date.parse($('fHasta').value) - Date.parse($('fDesde').value)) / 864e5 > 45);
    const clave = f => (semanal ? `${f['AÑO']}-S${String(f.SEM).padStart(2, '0')}` : f.FECHA);
    const g = {};
    fl.forEach(f => { const k = clave(f); const x = g[k] ||= { cant: 0, horas: 0 }; x.cant += Number(f.CANTIDAD || 0); x.horas += Number(f['TIEMPO LAB'] || 0); });
    const t = fl[0] ? tipo(fl[0]) : 'plantas';
    const puntos = Object.keys(g).sort().map(k => ({
      etiqueta: semanal ? k.slice(5) : `${k.slice(8)}/${k.slice(5, 7)}`,
      valor: g[k].horas ? g[k].cant / g[k].horas : 0,
      tip: `${semanal ? `Semana ${k.slice(6)}` : k}\n${fmt(g[k].horas ? g[k].cant / g[k].horas : 0, 1)} ${unidad(t)}\n${fmt(g[k].cant)} ${t} en ${fmt(g[k].horas, 1)} h`,
    }));
    $('cTendencia').innerHTML = encabezado(`Tendencia de rendimiento${laborTend ? ` · ${laborTend}` : ''}`, labores.length ? `${esc(unidad(t))} por ${semanal ? 'semana' : 'día'}, todo el equipo (cantidad ÷ horas). ${dias.size} días con registros.` : 'Sin labores medidas.',
      labores.length ? selectorLabor('selTend', labores, laborTend) : '') + linea(puntos, { ancho: anchoDe('cTendencia'), alto: 240, unidadTxt: unidad(t) });
    $('selTend')?.addEventListener('change', e => { laborTend = e.target.value; render(); });
  }

  function revision(filas) {
    const dias = [...new Set(filas.map(f => f.FECHA))].sort();
    const g = {};
    filas.forEach(f => {
      const p = g[f.CEDULA] ||= { nombre: f.NOMBRE, d: {}, a: {} };
      p.a[f.FECHA] ||= areaDe(f);
      p.d[f.FECHA] = (p.d[f.FECHA] || 0) + Number(f['TIEMPO LAB'] || 0);
    });
    const personas = Object.values(g).sort((a, b) => a.nombre.localeCompare(b.nombre));
    if (!personas.length || dias.length > 31) {
      $('cRevision').innerHTML = encabezado('Revisión de horas por día', `Horas laboradas por persona y día frente a la jornada de su área.`) +
        `<div class="empty">${personas.length ? 'Elige un rango de 31 días o menos para ver la revisión diaria.' : 'Sin datos en este filtro.'}</div>`;
      return;
    }
    // La jornada esperada depende del horario del área y del día de la semana.
    const celda = (h, area, fecha) => {
      if (!h) return '<td class="c" style="color:var(--text-3)">·</td>';
      const jornada = ctx.jornada(area, fecha);
      const bajo = jornada > 0 && h < jornada - 0.02, alto = jornada > 0 && h > jornada + 1;
      const st = bajo ? 'background:var(--warn-soft);color:var(--warn);font-weight:600' : alto ? 'background:var(--bad-soft);color:var(--bad);font-weight:600' : '';
      return `<td class="c" style="${st}" title="${fmt(h, 2)} h${bajo ? ' · incompleto' : alto ? ' · excede la jornada' : ''}">${fmt(h, 1)}${bajo ? '↓' : alto ? '↑' : ''}</td>`;
    };
    const dsem = ['D', 'L', 'M', 'X', 'J', 'V', 'S'];
    $('cRevision').innerHTML = encabezado('Revisión de horas por día', `Horas laboradas por persona y día. La jornada esperada es la del horario de cada área para ese día.`) +
      `<div class="legend"><span><i style="background:var(--warn-soft);border:1px solid var(--warn)"></i>↓ Menos de la jornada</span><span><i style="background:var(--bad-soft);border:1px solid var(--bad)"></i>↑ Más de una hora extra</span></div>
      <div class="table-scroll"><table class="hours-table heat"><thead><tr><th>Persona</th>${dias.map(d => `<th class="n" style="text-align:center">${dsem[new Date(`${d}T12:00`).getDay()]} ${d.slice(8)}</th>`).join('')}<th class="n">Total</th></tr></thead>
      <tbody>${personas.map(p => `<tr><td style="white-space:nowrap">${esc(corto(p.nombre, 28))}</td>${dias.map(d => celda(p.d[d], p.a[d], d)).join('')}<td class="n"><b>${fmt(Object.values(p.d).reduce((a, b) => a + b, 0), 1)}</b></td></tr>`).join('')}</tbody></table></div>`;
  }

  function bloques(filas) {
    // Cada cultivo tiene su plano: se usa el del área filtrada o, si no hay filtro, el del área de este celular.
    const area = Calc.tienePlano($('fArea').value) ? $('fArea').value : Calc.tienePlano(ctx.area()) ? ctx.area() : Calc.AREA_BASE;
    const sede = Calc.areaInfo(area).sede;
    const camas = ctx.camas().filter(c => Calc.sedeCama(c) === sede);
    const fl = filas.filter(f => tipo(f) === 'plantas' && f.BLOQUE && areaDe(f) === area);
    const labores = [...new Set(fl.map(f => f.LABOR))].sort();
    if (!labores.includes(laborBloque)) laborBloque = labores[0] || '';
    const inv = {};
    camas.forEach(c => { inv[c[0]] = (inv[c[0]] || 0) + c[3]; });
    const hecho = {};
    fl.filter(f => f.LABOR === laborBloque).forEach(f => { hecho[f.BLOQUE] = (hecho[f.BLOQUE] || 0) + Number(f.CANTIDAD || 0); });
    const lista = Object.keys(inv).sort(numSort).map(b => ({
      label: b, valor: inv[b] ? Math.min(100, ((hecho[b] || 0) / inv[b]) * 100) : 0,
      texto: `${fmt(((hecho[b] || 0) / inv[b]) * 100, ((hecho[b] || 0) / inv[b]) * 100 < 10 ? 1 : 0)}% · ${fmt(hecho[b] || 0)} de ${fmt(inv[b])}`,
      tip: `${b}\n${fmt(hecho[b] || 0)} plantas trabajadas\n${fmt(inv[b])} plantas en el plano`,
    }));
    $('cBloques').innerHTML = encabezado(`Avance por bloque · ${sede[0]}${sede.slice(1).toLowerCase().replace(/ (.)/g, (m, c) => ` ${c.toUpperCase()}`)}`, labores.length ? 'Plantas trabajadas en el período sobre el inventario del plano de producción. Si una cama se repasó, puede superar el 100%.' : 'Sin labores medidas en plantas en este filtro.',
      labores.length ? selectorLabor('selBloque', labores, laborBloque) : '') +
      (labores.length ? barrasH(lista, { ancho: anchoDe('cBloques'), etiqueta: 44, maxFijo: 100, margen: 160 }) : '');
    $('selBloque')?.addEventListener('change', e => { laborBloque = e.target.value; render(); });
  }

  // Ancho real de la tarjeta, para que el texto de los gráficos se lea a 12 px en celular y PC.
  const anchoDe = id => Math.max(300, ($(id).clientWidth || 640) - 34);

  let ultimas = [];
  async function render() {
    const todas = await filasCombinadas();
    // Opciones de filtros con todo el rango para no perder selecciones.
    opciones($('fLabor'), [...new Set(todas.map(f => f.LABOR))].sort().map(l => [l, l]), 'Todas');
    const pers = new Map(todas.map(f => [String(f.CEDULA), f.NOMBRE]));
    opciones($('fPersona'), [...pers].sort((a, b) => a[1].localeCompare(b[1])), 'Todas');
    opciones($('fBloque'), [...new Set(todas.map(f => f.BLOQUE).filter(Boolean))].sort(numSort).map(b => [b, b]), 'Todos');
    opciones($('fArea'), Calc.AREAS.map(a => [a.id, a.id]), 'Todas');
    opciones($('fLider'), [...new Set(todas.map(liderDe).filter(Boolean))].sort().map(x => [x, x]), 'Todos');
    opciones($('fLugar'), [...new Set(todas.map(lugar).filter(Boolean))].sort().map(b => [b, b]), 'Todos');
    const filas = filtrar(todas);
    ultimas = filas;
    pintarResumen();
    kpis(filas); promMax(filas); horasLabor(filas); porPersona(filas); tendencia(filas); revision(filas); bloques(filas);
  }

  // Resumen de filtros (celular) y número de filtros activos en el botón flotante.
  const FILTROS = [['fArea', 'Área'], ['fLider', 'Líder'], ['fLabor', 'Labor'], ['fPersona', 'Persona'], ['fBloque', 'Bloque'], ['fLugar', 'Lugar']];
  function pintarResumen() {
    const corta = f => (f ? `${f.slice(8)}/${f.slice(5, 7)}` : '…');
    const sem = $('fSemana').value ? `Sem ${$('fSemana').value} · ` : '';
    const activos = FILTROS.filter(([id]) => $(id).value).map(([id]) => $(id).selectedOptions[0]?.textContent || $(id).value);
    $('filtrosResumen').innerHTML = `<span class="fecha">${esc(sem)}${corta($('fDesde').value)} – ${corta($('fHasta').value)}</span>` +
      activos.map(t => `<span>${esc(corto(t, 24))}</span>`).join('');
    $('fabN').textContent = activos.length;
    $('fabN').classList.toggle('hidden', !activos.length);
  }

  function abrirFiltros(abrir) {
    $('dashFiltros').classList.toggle('abierto', abrir);
    $('filtrosFondo').classList.toggle('abierto', abrir);
  }

  function tooltips() {
    const tip = $('tip');
    document.addEventListener('pointermove', ev => {
      const g = ev.target.closest?.('[data-tip]');
      if (!g || !document.getElementById('view-dashboard').contains(g)) { tip.classList.add('hidden'); return; }
      const [t, ...rest] = g.dataset.tip.split('\n');
      tip.innerHTML = `<b>${esc(t)}</b>${rest.map(esc).join('<br>')}`;
      tip.classList.remove('hidden');
      const w = tip.offsetWidth, h = tip.offsetHeight;
      let x = ev.clientX + 14, y = ev.clientY + 14;
      if (x + w > innerWidth - 8) x = ev.clientX - w - 14;
      if (y + h > innerHeight - 8) y = ev.clientY - h - 14;
      tip.style.left = `${Math.max(8, x)}px`; tip.style.top = `${Math.max(8, y)}px`;
    });
    document.addEventListener('scroll', () => tip.classList.add('hidden'), { passive: true });
  }

  function enlazar() {
    enlazado = true;
    tooltips();
    $('rangeSeg').addEventListener('click', async ev => {
      const b = ev.target.closest('button[data-r]');
      if (!b) return;
      document.querySelectorAll('#rangeSeg button').forEach(x => x.classList.toggle('on', x === b));
      [$('fDesde').value, $('fHasta').value] = rango(b.dataset.r);
      limpiarPeriodo();
      await traerServidor(); render();
    });
    ['fDesde', 'fHasta'].forEach(id => $(id).addEventListener('change', async () => {
      document.querySelectorAll('#rangeSeg button').forEach(x => x.classList.remove('on'));
      limpiarPeriodo();
      await traerServidor(); render();
    }));
    ['fAnio', 'fMes', 'fSemana'].forEach(id => $(id).addEventListener('change', async () => {
      if (id !== 'fSemana') pintarSemanas();
      const r = rangoPeriodo();
      if (!r) return;
      if (id !== 'fSemana') pintarSemanas();
      [$('fDesde').value, $('fHasta').value] = r;
      document.querySelectorAll('#rangeSeg button').forEach(x => x.classList.remove('on'));
      await traerServidor(); render();
    }));
    ['fArea', 'fLider', 'fLabor', 'fPersona', 'fBloque', 'fLugar'].forEach(id => $(id).addEventListener('change', render));
    $('fabFiltros').addEventListener('click', () => abrirFiltros(true));
    $('filtrosListo').addEventListener('click', () => abrirFiltros(false));
    $('filtrosFondo').addEventListener('click', () => abrirFiltros(false));
    $('filtrosLimpiar').addEventListener('click', () => { FILTROS.forEach(([id]) => { $(id).value = ''; }); render(); });
    $('dashRefresh').addEventListener('click', async () => { await traerServidor(); render(); });
    let t;
    window.addEventListener('resize', () => { clearTimeout(t); t = setTimeout(() => { if (!$('view-dashboard').classList.contains('hidden')) render(); }, 250); });
    $('dashExport').addEventListener('click', () => {
      if (!ultimas.length) { ctx.toast('No hay filas para descargar con estos filtros.'); return; }
      const orden = [...ultimas].sort((a, b) => a.FECHA.localeCompare(b.FECHA) || String(a['HRA INICIO']).localeCompare(String(b['HRA INICIO'])));
      ctx.descargarExcel(orden, `rendimientos-${$('fDesde').value}-a-${$('fHasta').value}.xlsx`);
    });
  }

  async function abrir(c) {
    ctx = c;
    if (!enlazado) {
      enlazar();
      pintarPeriodo();
      [$('fDesde').value, $('fHasta').value] = rango('sem');
    }
    await render();
    traerServidor().then(render);
  }

  return { abrir };
})();
