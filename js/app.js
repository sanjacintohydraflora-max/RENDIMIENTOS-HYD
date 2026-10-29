// Interfaz de la app: registro en campo, registros del día, ajustes y arranque.
(() => {
  const $ = id => document.getElementById(id);
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const hoyStr = () => { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`; };
  const ahoraStr = () => { const d = new Date(); return `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`; };
  const fmt = (x, d = 0) => Number(x || 0).toLocaleString('es-CO', { maximumFractionDigits: d, minimumFractionDigits: d });
  const uid = () => (crypto.randomUUID ? crypto.randomUUID() : `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`);
  const numSort = (a, b) => parseFloat(String(a).replace(/[^\d.]/g, '')) - parseFloat(String(b).replace(/[^\d.]/g, '')) || String(a).localeCompare(String(b));

  const S = {
    empleados: [], empMap: {}, labores: [], camas: [], cfg: { ...Calc.CONFIG_BASE },
    sel: [], labor: null, camasSel: new Set(), tramos: [], editId: null, cantManual: false, editOriginal: null,
  };

  function toast(msg, ms = 2600) {
    const t = $('toast');
    t.textContent = msg;
    t.classList.remove('hidden');
    clearTimeout(toast._t);
    toast._t = setTimeout(() => t.classList.add('hidden'), ms);
  }

  // ---------- Carga de catálogos ----------
  async function cargarJSON(url) {
    try { const r = await fetch(url, { cache: 'no-cache' }); return r.ok ? r.json() : null; } catch { return null; }
  }

  async function cargarCatalogos() {
    S.cfg = { ...Calc.CONFIG_BASE, ...(await DB.getMeta('config', {})) };
    S.labores = await DB.getMeta('labores', []);
    if (!S.labores.length) {
      S.labores = (await cargarJSON('data/labores.json')) || [];
      if (S.labores.length) await DB.setMeta('labores', S.labores);
    }
    S.camas = await DB.getMeta('camas', []);
    if (!S.camas.length) {
      const c = await cargarJSON('data/camas.json');
      S.camas = c?.camas || [];
      if (S.camas.length) await DB.setMeta('camas', S.camas);
    }
    S.empleados = await DB.getMeta('empleados', []);
    if (!S.empleados.length) {
      // Solo existe en copias de prueba; en producción la lista llega del servidor o del Excel.
      const e = await cargarJSON('local/empleados.json');
      if (e?.length) { S.empleados = e; await DB.setMeta('empleados', e); }
    }
    S.empMap = Object.fromEntries(S.empleados.map(e => [String(e.cedula), e]));
    llenarBloques();
  }

  // ---------- Autocompletado ----------
  function marcar(texto, q) {
    let out = esc(texto);
    Calc.norm(q).split(/\s+/).filter(Boolean).forEach(t => {
      const n = Calc.norm(texto);
      const i = n.indexOf(t);
      if (i >= 0) {
        const orig = texto.slice(i, i + t.length);
        out = out.replace(esc(orig), `<mark>${esc(orig)}</mark>`);
      }
    });
    return out;
  }

  function autocompletar(input, box, buscarFn, pintarFn, elegirFn) {
    let items = [], hl = 0;
    const pintar = () => {
      if (!items.length) { box.classList.add('hidden'); box.innerHTML = ''; return; }
      box.innerHTML = items.map((it, i) => `<button type="button" data-i="${i}" class="${i === hl ? 'hl' : ''}">${pintarFn(it, input.value)}</button>`).join('');
      box.classList.remove('hidden');
    };
    const actualizar = () => { items = buscarFn(input.value); hl = 0; pintar(); };
    input.addEventListener('input', actualizar);
    input.addEventListener('focus', actualizar);
    input.addEventListener('keydown', ev => {
      if (box.classList.contains('hidden')) return;
      if (ev.key === 'ArrowDown') { hl = Math.min(items.length - 1, hl + 1); pintar(); ev.preventDefault(); }
      else if (ev.key === 'ArrowUp') { hl = Math.max(0, hl - 1); pintar(); ev.preventDefault(); }
      else if (ev.key === 'Enter') { if (items[hl]) { elegirFn(items[hl]); items = []; pintar(); } ev.preventDefault(); }
      else if (ev.key === 'Escape') { items = []; pintar(); }
    });
    // pointerdown para elegir antes de que el input pierda el foco
    box.addEventListener('pointerdown', ev => {
      const b = ev.target.closest('button[data-i]');
      if (!b) return;
      ev.preventDefault();
      elegirFn(items[+b.dataset.i]);
      items = []; pintar();
    });
    input.addEventListener('blur', () => setTimeout(() => { items = []; pintar(); }, 150));
  }

  function buscarEmpleados(q) {
    const activos = S.empleados.filter(e => Calc.norm(e.estado) !== 'RETIRADO' && !S.sel.includes(String(e.cedula)));
    if (!Calc.norm(q)) return [];
    const res = Calc.buscar(activos, q, 'nombre', ['cedula']);
    const campo = e => /CULTIVO|MIPE PRODUCCION/.test(Calc.norm(e.area)) ? 0 : 1;
    return res.sort((a, b) => campo(a) - campo(b) || a.nombre.localeCompare(b.nombre)).slice(0, 8);
  }

  function buscarLabores(q) {
    if (!Calc.norm(q)) return S.labores.slice();
    return Calc.buscar(S.labores, q, 'nombre');
  }

  function badgeMedida(medida) {
    const t = Calc.tipoMedida(medida);
    const cls = t === 'plantas' ? 'plantas' : t === 'tallos' ? 'tallos' : Calc.esGrupalPorMedida(medida) ? 'gr' : '';
    return `<span class="badge ${cls}">${esc(medida)}</span>`;
  }

  // ---------- Formulario ----------
  function elegirLabor(l) {
    S.labor = l;
    $('laborInput').value = l.nombre;
    actualizarLabor();
  }

  function actualizarLabor() {
    const l = S.labor;
    const tipo = l ? Calc.tipoMedida(l.medida) : 'tiempo';
    $('laborInfo').innerHTML = l
      ? `${badgeMedida(l.medida)} ${tipo === 'plantas' ? 'Requiere bloque, lote, camas y plantas.' : tipo === 'tallos' ? 'Requiere número de tallos.' : 'Se mide en tiempo; camas opcionales.'}${Calc.esGrupalPorMedida(l.medida) ? ' Labor grupal: agrega a todos los participantes.' : ''}`
      : '';
    const reqPl = tipo === 'plantas';
    ['bloqueLbl', 'loteLbl', 'camasLbl'].forEach(id => $(id).classList.toggle('req', reqPl));
    $('cantField').classList.toggle('hidden', tipo === 'tiempo');
    $('cantLbl').textContent = tipo === 'tallos' ? 'Tallos' : 'Plantas';
    $('cantLbl').classList.toggle('req', tipo !== 'tiempo');
    if (tipo === 'plantas' && !S.cantManual) autoPlantas();
    actualizarCantHint();
    recalcular();
  }

  function pintarChips() {
    $('empChips').innerHTML = S.sel.map(c => {
      const e = S.empMap[c];
      return `<span class="chip">${esc(e ? e.nombre : c)}<button type="button" data-quitar="${esc(c)}" aria-label="Quitar">×</button></span>`;
    }).join('');
    $('empHint').textContent = S.sel.length > 1
      ? `Labor grupal: ${S.sel.length} personas. La cantidad se reparte en partes iguales.`
      : 'Si la labor es grupal, agrega a todas las personas que participaron.';
    recalcular();
  }

  function elegirEmpleado(e) {
    const c = String(e.cedula);
    if (!S.sel.includes(c)) S.sel.push(c);
    $('empInput').value = '';
    pintarChips();
    $('empInput').focus();
  }

  function llenarBloques() {
    const bloques = [...new Set(S.camas.map(c => c[0]))].sort(numSort);
    const cur = $('bloque').value;
    $('bloque').innerHTML = '<option value="">—</option>' + bloques.map(b => `<option>${esc(b)}</option>`).join('');
    if (bloques.includes(cur)) $('bloque').value = cur;
  }

  function llenarLotes() {
    const b = $('bloque').value;
    const lotes = [...new Set(S.camas.filter(c => c[0] === b).map(c => c[1]))].sort(numSort);
    $('lote').innerHTML = '<option value="">—</option>' + lotes.map(l => `<option value="${esc(l)}">Lote ${esc(l)}</option>`).join('');
  }

  const camasDelLote = () => S.camas.filter(c => c[0] === $('bloque').value && c[1] === $('lote').value).sort((a, b) => a[2] - b[2]);

  function pintarCamas() {
    const lista = camasDelLote();
    const grid = $('camasGrid');
    if (!lista.length) {
      grid.innerHTML = '';
      $('camasHint').textContent = $('bloque').value ? 'Elige el lote para ver sus camas.' : 'Elige bloque y lote para ver las camas.';
      return;
    }
    grid.innerHTML = lista.map(c => `<button type="button" class="cama ${S.camasSel.has(c[2]) ? 'on' : ''}" data-cama="${c[2]}" title="${esc(c[4])}">${c[2]}<small>${fmt(c[3])}</small></button>`).join('');
    const total = lista.reduce((s, c) => s + c[3], 0);
    const variedades = [...new Set(lista.map(c => c[4]))].join(', ');
    $('camasHint').textContent = `${lista.length} camas, ${fmt(total)} plantas (${variedades}). Toca las camas realizadas; el número pequeño son las plantas.`;
  }

  function camasCambiaron(desdeTexto = false) {
    if (!desdeTexto) $('camasText').value = Calc.formatCamas([...S.camasSel]);
    document.querySelectorAll('#camasGrid .cama').forEach(b => b.classList.toggle('on', S.camasSel.has(+b.dataset.cama)));
    if (S.labor && Calc.tipoMedida(S.labor.medida) === 'plantas' && !S.cantManual) autoPlantas();
    actualizarCantHint();
    recalcular();
  }

  // Bloque/lote/camas que están marcados ahora en el selector (aún no añadidos a la lista).
  function tramoActual() {
    const bloque = $('bloque').value, lote = $('lote').value;
    if (!bloque || !lote) return null;
    const lista = camasDelLote();
    const camas = [...S.camasSel].sort((a, b) => a - b);
    const plantas = lista.filter(c => S.camasSel.has(c[2])).reduce((s, c) => s + c[3], 0);
    return { bloque, lote, camas, plantas, faltan: camas.filter(n => !lista.some(c => c[2] === n)) };
  }

  function tramosFormulario() {
    const t = tramoActual();
    const extra = t && (t.camas.length || !S.tramos.length) ? [{ bloque: t.bloque, lote: t.lote, camas: t.camas, plantas: t.plantas }] : [];
    return [...S.tramos, ...extra];
  }

  function plantasPlano() {
    const t = tramoActual();
    const todos = [...S.tramos, ...(t ? [t] : [])];
    return {
      total: todos.reduce((s, x) => s + (Number(x.plantas) || 0), 0),
      n: todos.reduce((s, x) => s + x.camas.length, 0),
      faltan: t ? t.faltan : [],
    };
  }

  function pintarTramos() {
    $('tramosBox').classList.toggle('hidden', !S.tramos.length);
    $('tramosChips').innerHTML = S.tramos.map((t, i) => `<span class="chip"><button type="button" data-tramo="${i}" style="font-size:14px;padding:0;color:var(--text)">${esc(t.bloque)} · Lote ${esc(t.lote)} · camas ${esc(Calc.formatCamas(t.camas) || '—')}${t.plantas ? ` · ${fmt(t.plantas)} pl.` : ''}</button><button type="button" data-quitar-tramo="${i}" aria-label="Quitar">×</button></span>`).join('');
    $('tramoAdd').textContent = S.tramos.length ? '+ Añadir este bloque y seguir con otro' : '+ Añadir otro bloque o lote';
  }

  function anadirTramo() {
    const t = tramoActual();
    if (!t || !t.camas.length) { toast('Elige bloque, lote y camas antes de añadir.'); return false; }
    S.tramos.push({ bloque: t.bloque, lote: t.lote, camas: t.camas, plantas: t.plantas });
    S.camasSel = new Set();
    $('lote').value = '';
    pintarTramos(); pintarCamas(); camasCambiaron();
    return true;
  }

  function editarTramo(i) {
    const actual = tramoActual();
    if (actual && actual.camas.length) S.tramos.push({ bloque: actual.bloque, lote: actual.lote, camas: actual.camas, plantas: actual.plantas });
    const [t] = S.tramos.splice(i, 1);
    $('bloque').value = t.bloque; llenarLotes(); $('lote').value = t.lote;
    S.camasSel = new Set(t.camas);
    pintarTramos(); pintarCamas(); camasCambiaron();
  }

  function autoPlantas() {
    const p = plantasPlano();
    $('cantidad').value = p.total || '';
  }

  function actualizarCantHint() {
    const tipo = S.labor ? Calc.tipoMedida(S.labor.medida) : 'tiempo';
    if (tipo !== 'plantas') { $('cantHint').innerHTML = ''; return; }
    const p = plantasPlano();
    let h = p.n ? `Plano: ${fmt(p.total)} plantas en ${p.n} camas.` : 'Marca camas para calcular las plantas del plano.';
    if (S.cantManual && p.n) h += ` <a href="#" id="usarPlano">Usar plano</a>`;
    if (p.faltan.length) h += ` Camas fuera del plano: ${p.faltan.join(', ')}.`;
    $('cantHint').innerHTML = h;
  }

  function aplicarDescuentosAuto() {
    const d = Calc.descuentosAuto($('hInicio').value, $('hFin').value, S.cfg);
    $('dDesayuno').checked = d.desayuno;
    $('dAlmuerzo').checked = d.almuerzo;
    $('dPausa').checked = d.pausa;
  }

  function pintarMinutosDescuento() {
    $('dDesayunoMin').textContent = `${S.cfg.desayunoMin} min`;
    $('dAlmuerzoMin').textContent = `${S.cfg.almuerzoMin} min`;
    $('dPausaMin').textContent = `${S.cfg.pausaMin} min`;
    $('dPausa').closest('.check').classList.toggle('hidden', !(S.cfg.pausaMin > 0));
  }

  function leerFormulario() {
    return {
      id: S.editId || uid(),
      fecha: $('fecha').value,
      lider: $('lider').value.trim(),
      labor: S.labor?.nombre || '',
      medida: S.labor?.medida || '',
      empleados: S.sel.slice(),
      nombres: Object.fromEntries(S.sel.map(c => [c, S.empMap[c]?.nombre || c])),
      horaInicio: $('hInicio').value,
      horaFin: $('hFin').value,
      desayuno: $('dDesayuno').checked,
      almuerzo: $('dAlmuerzo').checked,
      pausa: $('dPausa').checked,
      tramos: tramosFormulario(),
      cantidad: $('cantidad').value === '' ? null : Number($('cantidad').value),
      obs: $('obs').value.trim(),
    };
  }

  function recalcular() {
    const reg = leerFormulario();
    const box = $('calc');
    if (!reg.labor || !reg.empleados.length || !reg.horaInicio || !reg.horaFin) {
      box.innerHTML = `<div class="empty" style="padding:8px">${reg.horaInicio && !reg.horaFin ? 'Sin hora de fin: puedes guardarlo sin terminar y completarlo después.' : 'Completa la labor, las personas y el horario.'}</div>`;
      return;
    }
    const c = Calc.calcular(reg, S.cfg, S.empMap);
    const unidad = c.tipo === 'tallos' ? 'tallos' : 'plantas';
    const desc = [c.desayuno && `desayuno ${fmt(c.desayuno, 2)}`, c.almuerzo && `almuerzo ${fmt(c.almuerzo, 2)}`, c.pausa && `pausa ${fmt(c.pausa, 2)}`].filter(Boolean).join(', ');
    box.innerHTML = `<table>
      <tr><td>Horas totales</td><td>${fmt(c.horasTotales, 2)} h</td></tr>
      <tr><td>Descuentos</td><td>${desc || 'ninguno'}</td></tr>
      <tr><td>Tiempo laborado por persona</td><td>${fmt(c.tiempoLab, 2)} h</td></tr>
      ${c.tipo !== 'tiempo' ? `
      <tr><td>Total ${unidad}${c.n > 1 ? ` (grupo de ${c.n})` : ''}</td><td>${fmt(c.total)}</td></tr>
      ${c.n > 1 ? `<tr><td>${unidad[0].toUpperCase() + unidad.slice(1)} por persona</td><td>${fmt(c.porPersona, 1)}</td></tr>` : ''}
      <tr><td class="big">Rendimiento por persona</td><td class="big">${fmt(c.rendimiento, 1)} ${unidad}/h</td></tr>` : ''}
    </table>`;
  }

  function limpiarFormulario(mantener = true) {
    const fin = $('hFin').value;
    S.sel = []; S.labor = null; S.camasSel = new Set(); S.tramos = []; S.editId = null; S.cantManual = false; S.editOriginal = null;
    $('laborInput').value = ''; $('empInput').value = ''; $('cantidad').value = ''; $('obs').value = '';
    $('camasText').value = '';
    $('bloque').value = ''; llenarLotes();
    $('hInicio').value = mantener && fin ? fin : '';
    $('hFin').value = '';
    $('editBanner').classList.add('hidden');
    $('saveBtn').textContent = 'Guardar registro';
    $('savePartial').classList.remove('hidden'); $('partialHint').classList.remove('hidden');
    $('formErr').textContent = '';
    pintarChips(); pintarTramos(); actualizarLabor(); pintarCamas(); aplicarDescuentosAuto();
    pintarAbiertos();
  }

  async function cargarEnFormulario(id) {
    const r = await DB.getRegistro(id);
    if (!r) return;
    S.editId = r.id; S.editOriginal = r;
    $('fecha').value = r.fecha; $('lider').value = r.lider || '';
    S.labor = S.labores.find(l => l.nombre === r.labor) || { nombre: r.labor, medida: r.medida };
    $('laborInput').value = r.labor;
    S.sel = r.empleados.slice();
    r.empleados.forEach(c => { if (!S.empMap[c]) S.empMap[c] = { nombre: r.nombres?.[c] || c, cedula: c }; });
    $('hInicio').value = r.horaInicio; $('hFin').value = r.horaFin || '';
    S.tramos = Calc.tramosDe(r).map(t => ({ ...t, camas: [...(t.camas || [])] }));
    $('bloque').value = ''; llenarLotes();
    S.camasSel = new Set();
    pintarTramos();
    S.cantManual = r.cantidad != null && r.estado !== 'abierto';
    $('cantidad').value = r.cantidad ?? '';
    $('obs').value = r.obs || '';
    pintarCamas(); camasCambiaron();
    $('dDesayuno').checked = !!r.desayuno; $('dAlmuerzo').checked = !!r.almuerzo; $('dPausa').checked = !!r.pausa;
    pintarChips(); actualizarLabor();
    const abierto = r.estado === 'abierto';
    $('editBanner').classList.remove('hidden');
    $('editBanner').querySelector('.badge').textContent = abierto ? 'Terminando registro' : 'Editando registro';
    $('saveBtn').textContent = abierto ? 'Terminar registro' : 'Guardar cambios';
    $('savePartial').classList.toggle('hidden', !abierto); $('partialHint').classList.toggle('hidden', !abierto);
    if (abierto && !r.horaFin) setTimeout(() => $('hFin').focus(), 50);
    pintarAbiertos();
    mostrar('registrar');
    window.scrollTo(0, 0);
  }

  async function guardar(ev, parcial = false) {
    ev?.preventDefault();
    // Si escribió una labor exacta sin elegirla de la lista, se acepta.
    if (!S.labor && $('laborInput').value) {
      const l = S.labores.find(x => Calc.norm(x.nombre) === Calc.norm($('laborInput').value));
      if (l) S.labor = l;
    }
    const reg = leerFormulario();
    reg.estado = parcial ? 'abierto' : 'completo';
    if (parcial) reg.horaFin = '';
    const errores = Calc.validar(reg, S.cfg, { parcial });
    if ($('laborInput').value && !S.labor) errores.unshift('La labor escrita no está en la lista; elígela de las sugerencias.');
    const todos = await DB.allRegistros();
    // No se empieza otro día mientras queden registros sin terminar de días anteriores.
    const pendientes = todos.filter(r => r.estado === 'abierto' && r.id !== reg.id && r.fecha < reg.fecha);
    if (pendientes.length) {
      const dias = [...new Set(pendientes.map(r => r.fecha))].join(', ');
      errores.unshift(`Tienes ${pendientes.length} registro(s) sin terminar del ${dias}. Termínalos antes de registrar el ${reg.fecha}.`);
    }
    errores.push(...Calc.cruces(reg, todos, c => S.empMap[c]?.nombre || reg.nombres?.[c] || c));
    if (errores.length) { $('formErr').innerHTML = errores.map(esc).join('<br>'); return; }
    reg.creado = S.editOriginal?.creado || new Date().toISOString();
    reg.actualizado = new Date().toISOString();
    reg.sincronizado = false;
    await DB.putRegistro(reg);
    if (reg.lider) await DB.setMeta('lider', reg.lider);
    const editando = !!S.editId;
    toast(parcial ? 'Guardado sin terminar. Lo completas desde el aviso de arriba o en Registros.'
      : editando ? (S.editOriginal?.estado === 'abierto' ? 'Registro terminado' : 'Cambios guardados')
        : `Registro guardado${reg.empleados.length > 1 ? ` para ${reg.empleados.length} personas` : ''}`, 3500);
    limpiarFormulario(!editando && !parcial);
    actualizarSyncPill();
    Sync.sincronizar({ catalogos: false });
  }

  // ---------- Registros sin terminar ----------
  async function pintarAbiertos() {
    const abiertos = (await DB.allRegistros()).filter(r => r.estado === 'abierto' && r.id !== S.editId)
      .sort((a, b) => a.fecha.localeCompare(b.fecha) || a.horaInicio.localeCompare(b.horaInicio));
    $('abiertosCard').classList.toggle('hidden', !abiertos.length);
    $('abiertosLista').innerHTML = abiertos.map(r => {
      const nombres = r.empleados.map(x => S.empMap[x]?.nombre || r.nombres?.[x] || x).join(', ');
      const t = Calc.tramosDe(r).map(x => `${x.bloque} L${x.lote}${x.camas?.length ? ` camas ${Calc.formatCamas(x.camas)}` : ''}`).join(' · ');
      return `<div class="reg"><div class="r-main"><div class="r-t">${esc(r.labor)}</div><div class="r-s">${esc(nombres)}</div>
        <div class="r-s">${esc(r.fecha)} desde las ${esc(r.horaInicio)}${t ? ` · ${esc(t)}` : ''}</div></div>
        <div class="r-a"><button type="button" class="btn small primary" data-editar="${r.id}">Completar</button></div></div>`;
    }).join('');
  }

  // ---------- Registros del día ----------
  async function pintarHoy() {
    const f = $('hoyFecha').value || hoyStr();
    const regs = (await DB.registrosPorFecha(f)).sort((a, b) => (a.horaInicio || '').localeCompare(b.horaInicio || ''));
    const lista = $('hoyLista');
    if (!regs.length) {
      lista.innerHTML = '<div class="empty">No hay registros en este día.</div>';
      $('hoyHoras').innerHTML = '<div class="empty">Sin datos.</div>';
      return;
    }
    lista.innerHTML = regs.map(r => {
      const c = Calc.calcular(r, S.cfg, S.empMap);
      const nombres = r.empleados.map(x => S.empMap[x]?.nombre || r.nombres?.[x] || x).join(', ');
      const unidad = c.tipo === 'tallos' ? 'tallos' : 'plantas';
      const abierto = r.estado === 'abierto';
      const det = [
        abierto ? `desde las ${r.horaInicio}, sin terminar` : `${r.horaInicio}–${r.horaFin} (${fmt(c.tiempoLab, 2)} h)`,
        ...Calc.tramosDe(r).map(t => `${t.bloque} lote ${t.lote}${t.camas?.length ? ` camas ${Calc.formatCamas(t.camas)}` : ''}`),
        !abierto && c.tipo !== 'tiempo' && `${fmt(c.total)} ${unidad} · ${fmt(c.rendimiento, 1)} ${unidad}/h por persona`,
      ].filter(Boolean).join(' · ');
      return `<div class="reg">
        <div class="r-main">
          <div class="r-t">${esc(r.labor)} ${r.empleados.length > 1 ? `<span class="badge gr">Grupo ${r.empleados.length}</span>` : ''}</div>
          <div class="r-s">${esc(nombres)}</div>
          <div class="r-s">${esc(det)}</div>
          ${r.obs ? `<div class="r-s">${esc(r.obs)}</div>` : ''}
        </div>
        <div class="r-a">
          ${abierto ? '<span class="st pend">Sin terminar</span>' : `<span class="st ${r.sincronizado ? 'ok' : 'pend'}">${r.sincronizado ? 'Sincronizado' : 'Pendiente'}</span>`}
          <button type="button" class="btn small ${abierto ? 'primary' : ''}" data-editar="${r.id}">${abierto ? 'Completar' : 'Editar'}</button>
          <button type="button" class="btn small danger" data-borrar="${r.id}">Borrar</button>
        </div>
      </div>`;
    }).join('');

    // Horas por persona y cruces de horario
    const porPersona = {};
    regs.forEach(r => {
      const c = Calc.calcular(r, S.cfg, S.empMap);
      r.empleados.forEach(ced => {
        const p = porPersona[ced] ||= { nombre: S.empMap[ced]?.nombre || r.nombres?.[ced] || ced, horas: 0, tramos: [] };
        p.horas += c.tiempoLab;
        if (r.estado === 'abierto') p.abierto = true;
        else p.tramos.push([Calc.minutos(r.horaInicio), Calc.minutos(r.horaFin)]);
      });
    });
    const filas = Object.values(porPersona).sort((a, b) => a.nombre.localeCompare(b.nombre)).map(p => {
      p.tramos.sort((a, b) => a[0] - b[0]);
      const cruce = p.tramos.some((t, i) => i > 0 && t[0] < p.tramos[i - 1][1]);
      const bajo = p.horas < S.cfg.jornada - 0.25;
      const nota = p.abierto ? '<span class="flag-low">Labor sin terminar</span>' : cruce ? '<span class="flag-high">Horarios cruzados</span>' : bajo ? `<span class="flag-low">Faltan ${fmt(S.cfg.jornada - p.horas, 2)} h</span>` : 'Completo';
      return `<tr><td>${esc(p.nombre)}</td><td class="n">${fmt(p.horas, 2)}</td><td>${nota}</td></tr>`;
    }).join('');
    $('hoyHoras').innerHTML = `<div class="table-scroll"><table class="hours-table"><thead><tr><th>Persona</th><th class="n">Horas lab.</th><th>Revisión (jornada ${fmt(S.cfg.jornada, 2)} h)</th></tr></thead><tbody>${filas}</tbody></table></div>`;
  }

  async function borrar(id) {
    if (!confirm('¿Borrar este registro? Se borra también del servidor al sincronizar.')) return;
    await DB.deleteRegistro(id);
    const b = await DB.getMeta('borrados', []);
    b.push(id);
    await DB.setMeta('borrados', b);
    toast('Registro borrado');
    pintarHoy(); actualizarSyncPill();
    Sync.sincronizar({ catalogos: false });
  }

  // ---------- Estado de sincronización ----------
  async function actualizarSyncPill(st) {
    st = st || await Sync.estado();
    const pill = $('syncPill');
    pill.classList.remove('pend', 'off');
    let txt;
    if (st.tipo === 'sincronizando') txt = 'Sincronizando…';
    else if (st.tipo === 'sin-servidor') { txt = st.pendientes ? `${st.pendientes} en el celular` : 'Solo local'; pill.classList.add('pend'); }
    else if (st.tipo === 'sin-senal') { txt = st.pendientes ? `Sin señal · ${st.pendientes} pend.` : 'Sin señal'; pill.classList.add('off'); }
    else if (st.error) { txt = `Error · ${st.pendientes} pend.`; pill.classList.add('off'); }
    else if (st.pendientes) { txt = `${st.pendientes} pendientes`; pill.classList.add('pend'); }
    else txt = 'Al día';
    $('syncText').textContent = txt;
    pill.title = st.error ? `Último error: ${st.error}` : st.ultima ? `Última sincronización: ${new Date(st.ultima).toLocaleString('es-CO')}` : 'Sincronizar ahora';
  }

  // ---------- Importación de Excel ----------
  function leerExcel(file) {
    return new Promise((resolve, reject) => {
      const fr = new FileReader();
      fr.onload = () => { try { resolve(XLSX.read(fr.result, { type: 'array', cellDates: true })); } catch (e) { reject(e); } };
      fr.onerror = () => reject(fr.error);
      fr.readAsArrayBuffer(file);
    });
  }

  // Busca la hoja y la fila de encabezados que contienen las columnas pedidas.
  function tabla(wb, preferida, requeridas) {
    const hojas = [...wb.SheetNames].sort((a, b) => (Calc.norm(b).includes(preferida) ? 1 : 0) - (Calc.norm(a).includes(preferida) ? 1 : 0));
    for (const nombre of hojas) {
      const filas = XLSX.utils.sheet_to_json(wb.Sheets[nombre], { header: 1, defval: '', raw: true });
      for (let i = 0; i < Math.min(15, filas.length); i++) {
        const enc = filas[i].map(h => Calc.norm(h).replace(/\s+/g, ' '));
        const idx = {};
        for (const [k, pats] of Object.entries(requeridas)) idx[k] = enc.findIndex(h => pats.some(p => h.startsWith(p)));
        if (Object.values(idx).every(v => v >= 0)) return { filas: filas.slice(i + 1), idx, enc };
      }
    }
    return null;
  }

  async function importarEmpleados(file) {
    const wb = await leerExcel(file);
    const t = tabla(wb, 'EMPLEADO', { nombre: ['NOMBRE'], cedula: ['CEDULA'], estado: ['ESTADO'] });
    if (!t) throw new Error('No encontré una hoja con columnas NOMBRE, ESTADO y CEDULA.');
    const col = p => t.enc.findIndex(h => h.startsWith(p));
    const iL = col('LUGAR'), iC = col('CARGO'), iA = col('AREA');
    const lista = t.filas.filter(f => String(f[t.idx.nombre]).trim() && String(f[t.idx.cedula]).trim()).map(f => ({
      nombre: String(f[t.idx.nombre]).trim().replace(/\s+/g, ' '),
      estado: String(f[t.idx.estado]).trim().toUpperCase(),
      cedula: String(f[t.idx.cedula]).trim().replace(/\.0$/, ''),
      lugar: iL >= 0 ? String(f[iL]).trim() : '',
      cargo: iC >= 0 ? String(f[iC]).trim() : '',
      area: iA >= 0 ? String(f[iA]).trim() : '',
    }));
    if (!lista.length) throw new Error('La hoja de empleados está vacía.');
    // Conserva a quienes ya no están en el Excel como RETIRADO para no perder su historial.
    const nuevos = new Set(lista.map(e => e.cedula));
    const salieron = S.empleados.filter(e => !nuevos.has(String(e.cedula))).map(e => ({ ...e, estado: 'RETIRADO' }));
    const todos = [...lista, ...salieron];
    await DB.setMeta('empleados', todos);
    S.empleados = todos;
    S.empMap = Object.fromEntries(todos.map(e => [String(e.cedula), e]));
    return { activos: lista.filter(e => e.estado !== 'RETIRADO').length, retirados: todos.length - lista.filter(e => e.estado !== 'RETIRADO').length, todos };
  }

  async function importarLabores(file) {
    const wb = await leerExcel(file);
    const t = tabla(wb, 'LABOR', { nombre: ['LABOR'], medida: ['MEDIDA'] });
    if (!t) throw new Error('No encontré una hoja con columnas LABORES y MEDIDA.');
    const lista = t.filas.filter(f => String(f[t.idx.nombre]).trim()).map(f => ({ nombre: String(f[t.idx.nombre]).trim(), medida: String(f[t.idx.medida]).trim() || 'Tiempo' }));
    await DB.setMeta('labores', lista);
    S.labores = lista;
    return lista;
  }

  async function importarPlano(file) {
    const wb = await leerExcel(file);
    const t = tabla(wb, 'DATA', { bloque: ['BLOQUE'], lote: ['LOTE'], era: ['ERA', 'CAMA'], plantas: ['CANTIDAD DE PLANTA', 'PLANTAS'] });
    if (!t) throw new Error('No encontré una hoja con BLOQUE, LOTE, ERA y CANTIDAD DE PLANTAS.');
    const iV = t.enc.findIndex(h => h.startsWith('VARIEDAD'));
    const lista = t.filas.filter(f => String(f[t.idx.bloque]).trim() && f[t.idx.era] !== '').map(f => {
      const lote = f[t.idx.lote];
      return [String(f[t.idx.bloque]).trim(), typeof lote === 'number' && Number.isInteger(lote) ? String(lote) : String(lote).trim(), Number(f[t.idx.era]), Number(f[t.idx.plantas]) || 0, iV >= 0 ? String(f[iV]).trim() : ''];
    }).filter(c => Number.isFinite(c[2]));
    await DB.setMeta('camas', lista);
    S.camas = lista;
    llenarBloques();
    return lista;
  }

  async function pintarAjustes() {
    const act = S.empleados.filter(e => Calc.norm(e.estado) !== 'RETIRADO').length;
    $('empStatus').textContent = `${act} empleados activos y ${S.empleados.length - act} retirados en la lista.`;
    const bloques = new Set(S.camas.map(c => c[0])).size;
    const plantas = S.camas.reduce((s, c) => s + c[3], 0);
    $('catStatus').textContent = `${S.labores.length} labores. Plano: ${bloques} bloques, ${fmt(S.camas.length)} camas, ${fmt(plantas)} plantas.`;
    const c = S.cfg;
    $('cDesH').value = c.desayunoHora; $('cDesM').value = c.desayunoMin;
    $('cAlmH').value = c.almuerzoHora; $('cAlmM').value = c.almuerzoMin;
    $('cPauH').value = c.pausaHora; $('cPauM').value = c.pausaMin;
    $('cJornada').value = c.jornada;
    const s = await DB.getMeta('sync', { url: '', key: '' });
    $('cUrl').value = s.url; $('cKey').value = s.key;
    const regs = await DB.allRegistros();
    const pend = regs.filter(r => !r.sincronizado).length;
    $('localInfo').textContent = `${regs.length} registros guardados en este celular, ${pend} sin sincronizar.`;
    const ult = await DB.getMeta('ultimaSync', null);
    $('syncInfo').textContent = ult ? `Última sincronización: ${new Date(ult).toLocaleString('es-CO')}` : '';
  }

  async function filasLocales() {
    const regs = (await DB.allRegistros()).filter(r => r.estado !== 'abierto');
    return regs.flatMap(r => Calc.calcular(r, S.cfg, S.empMap).filas);
  }

  function descargarExcel(filas, nombre) {
    const ws = XLSX.utils.json_to_sheet(filas, { header: Calc.COLUMNAS });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'RENDIMIENTOS');
    XLSX.writeFile(wb, nombre);
  }

  // ---------- Compartir la conexión ----------
  // El enlace lleva la URL y la clave después del #, que el navegador nunca envía a GitHub.
  const codificar = o => btoa(unescape(encodeURIComponent(JSON.stringify(o))));
  const decodificar = t => JSON.parse(decodeURIComponent(escape(atob(t))));

  async function enlaceConexion() {
    const s = await DB.getMeta('sync', { url: '', key: '' });
    if (!s.url) return null;
    return `${location.origin}${location.pathname}#conectar=${encodeURIComponent(codificar({ u: s.url, k: s.key }))}`;
  }

  async function mostrarCompartir() {
    const link = await enlaceConexion();
    if (!link) { toast('Primero conecta este equipo al servidor (Guardar y probar).'); return; }
    const qr = qrcode(0, 'M');
    qr.addData(link);
    qr.make();
    $('shareQr').innerHTML = qr.createSvgTag({ cellSize: 4, margin: 2, scalable: true });
    const svg = $('shareQr').querySelector('svg');
    if (svg) { svg.style.width = '240px'; svg.style.maxWidth = '100%'; svg.style.height = 'auto'; svg.style.display = 'block'; }
    $('shareLink').value = link;
    $('shareSend').classList.toggle('hidden', !navigator.share);
    $('shareBox').classList.remove('hidden');
  }

  async function conectarDesdeEnlace() {
    const m = location.hash.match(/^#conectar=(.+)$/);
    if (!m) return false;
    history.replaceState(null, '', location.pathname);
    let datos;
    try { datos = decodificar(decodeURIComponent(m[1])); } catch { toast('El enlace de conexión está incompleto. Pide uno nuevo.', 5000); return false; }
    if (!datos?.u) return false;
    if (!confirm('¿Conectar este celular al servidor de rendimientos Hydraflora?')) return false;
    await DB.setMeta('sync', { url: datos.u, key: datos.k || '' });
    const r = await Sync.sincronizar();
    await cargarCatalogos();
    toast(r?.error ? `No se pudo conectar: ${r.error}` : 'Celular conectado. Ya bajaron empleados, labores y plano.', 5000);
    return true;
  }

  // ---------- Navegación ----------
  function mostrar(vista) {
    document.querySelectorAll('main[id^="view-"]').forEach(m => m.classList.toggle('hidden', m.id !== `view-${vista}`));
    document.querySelectorAll('#tabs button').forEach(b => b.classList.toggle('active', b.dataset.view === vista));
    if (vista === 'hoy') pintarHoy();
    if (vista === 'ajustes') pintarAjustes();
    if (vista === 'dashboard') Dashboard.abrir({ filasLocales, empMap: () => S.empMap, camas: () => S.camas, cfg: () => S.cfg, descargarExcel, toast });
    try { localStorage.setItem('vista', vista); } catch { /* sin almacenamiento */ }
  }

  // ---------- Eventos ----------
  function enlazar() {
    document.querySelectorAll('#tabs button').forEach(b => b.addEventListener('click', () => mostrar(b.dataset.view)));

    autocompletar($('laborInput'), $('laborSuggest'), buscarLabores,
      (l, q) => `<div class="s-main">${marcar(l.nombre, q)}</div><div class="s-sub">${esc(l.medida)}</div>`, elegirLabor);
    $('laborInput').addEventListener('input', () => { if (S.labor && S.labor.nombre !== $('laborInput').value) { S.labor = null; actualizarLabor(); } });

    autocompletar($('empInput'), $('empSuggest'), buscarEmpleados,
      (e, q) => `<div class="s-main">${marcar(e.nombre, q)}</div><div class="s-sub">${marcar(e.cedula, q)} · ${esc(e.cargo || '')} · ${esc(e.area || e.lugar || '')}</div>`, elegirEmpleado);
    $('empChips').addEventListener('click', ev => {
      const b = ev.target.closest('[data-quitar]');
      if (b) { S.sel = S.sel.filter(c => c !== b.dataset.quitar); pintarChips(); }
    });

    document.querySelectorAll('[data-now]').forEach(b => b.addEventListener('click', () => {
      $(b.dataset.now).value = ahoraStr(); aplicarDescuentosAuto(); recalcular();
    }));
    ['hInicio', 'hFin'].forEach(id => $(id).addEventListener('change', () => { aplicarDescuentosAuto(); recalcular(); }));
    ['dDesayuno', 'dAlmuerzo', 'dPausa'].forEach(id => $(id).addEventListener('change', recalcular));

    $('bloque').addEventListener('change', () => { llenarLotes(); S.camasSel = new Set(); pintarCamas(); camasCambiaron(); });
    $('lote').addEventListener('change', () => { S.camasSel = new Set(); pintarCamas(); camasCambiaron(); });
    $('camasGrid').addEventListener('click', ev => {
      const b = ev.target.closest('.cama');
      if (!b) return;
      const n = +b.dataset.cama;
      S.camasSel.has(n) ? S.camasSel.delete(n) : S.camasSel.add(n);
      camasCambiaron();
    });
    $('camasText').addEventListener('input', () => { S.camasSel = new Set(Calc.parseCamas($('camasText').value)); camasCambiaron(true); });
    $('camasAll').addEventListener('click', () => { S.camasSel = new Set(camasDelLote().map(c => c[2])); camasCambiaron(); });
    $('camasNone').addEventListener('click', () => { S.camasSel = new Set(); camasCambiaron(); });
    $('cantidad').addEventListener('input', () => { S.cantManual = $('cantidad').value !== ''; actualizarCantHint(); recalcular(); });
    $('cantHint').addEventListener('click', ev => {
      if (ev.target.id === 'usarPlano') { ev.preventDefault(); S.cantManual = false; autoPlantas(); actualizarCantHint(); recalcular(); }
    });
    $('obs').addEventListener('input', recalcular);
    $('fecha').addEventListener('change', recalcular);
    $('form').addEventListener('submit', ev => guardar(ev, false));
    $('savePartial').addEventListener('click', () => guardar(null, true));
    $('tramoAdd').addEventListener('click', anadirTramo);
    $('tramosChips').addEventListener('click', ev => {
      const q = ev.target.closest('[data-quitar-tramo]');
      if (q) { S.tramos.splice(+q.dataset.quitarTramo, 1); pintarTramos(); camasCambiaron(); return; }
      const e = ev.target.closest('[data-tramo]');
      if (e) editarTramo(+e.dataset.tramo);
    });
    $('abiertosLista').addEventListener('click', ev => {
      const e = ev.target.closest('[data-editar]'); if (e) cargarEnFormulario(e.dataset.editar);
    });
    $('cancelEdit').addEventListener('click', () => limpiarFormulario(false));

    $('hoyFecha').addEventListener('change', pintarHoy);
    $('hoySync').addEventListener('click', async () => { const r = await Sync.sincronizar(); toast(r?.error ? `No se pudo sincronizar: ${r.error}` : 'Sincronización completa'); pintarHoy(); });
    $('hoyLista').addEventListener('click', ev => {
      const e = ev.target.closest('[data-editar]'); if (e) cargarEnFormulario(e.dataset.editar);
      const d = ev.target.closest('[data-borrar]'); if (d) borrar(d.dataset.borrar);
    });
    $('syncPill').addEventListener('click', async () => {
      const st = await Sync.estado();
      if (st.tipo === 'sin-servidor') { toast('Configura el servidor en Ajustes para sincronizar.'); return; }
      const r = await Sync.sincronizar();
      if (r?.error) toast(`No se pudo sincronizar: ${r.error}`);
    });

    const subir = (btn, input, fn, ok) => {
      $(btn).addEventListener('click', () => $(input).click());
      $(input).addEventListener('change', async () => {
        const f = $(input).files[0];
        $(input).value = '';
        if (!f) return;
        try {
          const res = await fn(f);
          toast(ok(res), 4000);
          pintarAjustes();
        } catch (e) { toast(`Error: ${e.message}`, 5000); }
      });
    };
    const alServidor = async (tipo, datos) => {
      const s = await DB.getMeta('sync', { url: '' });
      if (!s.url || !navigator.onLine) return ' Se subirá a la hoja de Google al conectar la sincronización.';
      try { await Sync.subirCatalogo(tipo, datos); return ' También quedó en el servidor.'; } catch (e) { return ` No se pudo enviar al servidor: ${e.message}`; }
    };
    subir('btnEmp', 'fileEmp', async f => { const r = await importarEmpleados(f); r.nota = await alServidor('empleados', r.todos); return r; },
      r => `Empleados cargados: ${r.activos} activos, ${r.retirados} retirados.${r.nota}`);
    subir('btnLab', 'fileLab', async f => { const r = await importarLabores(f); return { n: r.length, nota: await alServidor('labores', r) }; },
      r => `${r.n} labores cargadas.${r.nota}`);
    subir('btnPlano', 'filePlano', async f => { const r = await importarPlano(f); return { n: r.length, nota: await alServidor('camas', r) }; },
      r => `${r.n} camas cargadas del plano.${r.nota}`);

    $('saveCfg').addEventListener('click', async () => {
      S.cfg = {
        desayunoHora: $('cDesH').value || '08:30', desayunoMin: Number($('cDesM').value) || 0,
        almuerzoHora: $('cAlmH').value || '12:00', almuerzoMin: Number($('cAlmM').value) || 0,
        pausaHora: $('cPauH').value || '10:00', pausaMin: Number($('cPauM').value) || 0,
        jornada: Number($('cJornada').value) || 8.4,
      };
      await DB.setMeta('config', S.cfg);
      pintarMinutosDescuento();
      toast('Horario guardado. Aplica a los registros nuevos y a los que edites.');
    });
    $('saveSync').addEventListener('click', async () => {
      await DB.setMeta('sync', { url: $('cUrl').value.trim(), key: $('cKey').value.trim() });
      $('syncInfo').textContent = 'Probando conexión…';
      const r = await Sync.sincronizar();
      $('syncInfo').textContent = r?.error ? `No se pudo conectar: ${r.error}` : `Conectado. ${r?.enviados || 0} registros enviados y catálogos actualizados.`;
      await cargarCatalogos();
      pintarAjustes();
    });
    $('shareSync').addEventListener('click', mostrarCompartir);
    $('shareHide').addEventListener('click', () => $('shareBox').classList.add('hidden'));
    $('shareCopy').addEventListener('click', async () => {
      try { await navigator.clipboard.writeText($('shareLink').value); toast('Enlace copiado'); }
      catch { $('shareLink').select(); document.execCommand('copy'); toast('Enlace copiado'); }
    });
    $('shareSend').addEventListener('click', () => navigator.share({ title: 'Rendimientos Hydraflora', text: 'Abre este enlace para conectar la app de rendimientos:', url: $('shareLink').value }).catch(() => {}));
    $('exportLocal').addEventListener('click', async () => descargarExcel(await filasLocales(), `rendimientos-celular-${hoyStr()}.xlsx`));

    Sync.onCambio(st => { actualizarSyncPill(st); if (st.tipo !== 'sincronizando') cargarCatalogos(); });
  }

  async function iniciar() {
    if ('serviceWorker' in navigator) {
      // Cuando llega una versión nueva de la app, recarga una vez para mostrarla.
      const habiaControlador = !!navigator.serviceWorker.controller;
      let recargado = false;
      navigator.serviceWorker.addEventListener('controllerchange', () => {
        if (!habiaControlador || recargado) return;
        // No recarga si hay un registro a medio llenar: se perdería lo escrito.
        if (S.sel.length || S.labor || $('laborInput').value || $('empInput').value) { toast('Hay una versión nueva de la app. Se verá la próxima vez que la abras.', 4000); return; }
        recargado = true;
        location.reload();
      });
      navigator.serviceWorker.register('sw.js', { updateViaCache: 'none' })
        .then(reg => { document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') reg.update().catch(() => {}); }); })
        .catch(() => {});
    }
    await cargarCatalogos();
    enlazar();
    $('fecha').value = hoyStr();
    $('hoyFecha').value = hoyStr();
    $('lider').value = await DB.getMeta('lider', '');
    pintarMinutosDescuento();
    limpiarFormulario(false);
    actualizarSyncPill();
    let vista = 'registrar';
    try { vista = localStorage.getItem('vista') || vista; } catch { /* sin almacenamiento */ }
    if (location.hash === '#dashboard') vista = 'dashboard';
    const conectado = await conectarDesdeEnlace();
    mostrar(conectado ? 'registrar' : vista);
    if (!conectado) Sync.sincronizar();
  }

  document.addEventListener('DOMContentLoaded', iniciar);
})();
