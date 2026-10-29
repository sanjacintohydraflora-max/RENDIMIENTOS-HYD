// Sincronización con la hoja de Google (Apps Script). Todo se guarda primero en el celular;
// cuando hay señal se envían los registros pendientes y se bajan los catálogos actualizados.
const Sync = (() => {
  let enCurso = false;
  const oyentes = new Set();
  const emitir = estado => oyentes.forEach(fn => fn(estado));

  async function conf() {
    return DB.getMeta('sync', { url: '', key: '' });
  }

  async function llamar(params, body) {
    const { url, key } = await conf();
    if (!url) throw new Error('Falta configurar la URL del servidor en Ajustes.');
    const qs = new URLSearchParams({ ...params, key }).toString();
    const opts = body
      ? { method: 'POST', body: JSON.stringify({ ...body, key }), headers: { 'Content-Type': 'text/plain;charset=utf-8' } }
      : { method: 'GET' };
    if (!/\/exec(\?|$)/.test(url)) throw new Error('La URL debe terminar en /exec (la de "Implementar > Gestionar implementaciones").');
    let res;
    try {
      res = await fetch(`${url}${url.includes('?') ? '&' : '?'}${qs}`, { ...opts, redirect: 'follow' });
    } catch {
      throw new Error('el servidor no respondió. Revisa la URL y que la implementación tenga acceso "Cualquier persona" y se ejecute como "Yo".');
    }
    if (!res.ok) throw new Error(`el servidor respondió ${res.status}. Revisa que la implementación tenga acceso "Cualquier persona".`);
    let data;
    try { data = await res.json(); } catch {
      throw new Error('la URL abrió una página de Google en vez del script. Revisa que el acceso sea "Cualquier persona".');
    }
    if (!data.ok) throw new Error(data.error === 'Clave incorrecta' ? 'clave incorrecta. Si ejecutaste configurar antes de cambiar la clave, prueba con cambia-esta-clave.' : data.error || 'error del servidor');
    return data;
  }

  async function pendientes() {
    const regs = await DB.allRegistros();
    const borrados = await DB.getMeta('borrados', []);
    // Los registros sin terminar se quedan en el celular hasta completarlos.
    return { regs: regs.filter(r => !r.sincronizado && r.estado !== 'abierto'), borrados };
  }

  async function empleadosMapa() {
    const lista = await DB.getMeta('empleados', []);
    return Object.fromEntries(lista.map(e => [String(e.cedula), e]));
  }

  // Envía pendientes en lotes de 100 registros.
  async function subir() {
    const { regs, borrados } = await pendientes();
    if (!regs.length && !borrados.length) return 0;
    const cfg = { ...Calc.CONFIG_BASE, ...(await DB.getMeta('config', {})) };
    const emp = await empleadosMapa();
    for (let i = 0; i < Math.max(regs.length, 1); i += 100) {
      const lote = regs.slice(i, i + 100);
      const payload = lote.map(r => ({ id: r.id, filas: Calc.calcular(r, cfg, emp).filas }));
      const enviados = i === 0 ? borrados : [];
      await llamar({ action: 'guardar' }, { action: 'guardar', registros: payload, borrados: enviados });
      // Marca como sincronizados solo los que no se editaron mientras se enviaban.
      for (const r of lote) {
        const actual = await DB.getRegistro(r.id);
        if (actual && actual.actualizado === r.actualizado) {
          actual.sincronizado = true;
          await DB.putRegistro(actual);
        }
      }
      if (enviados.length) {
        const resto = (await DB.getMeta('borrados', [])).filter(id => !enviados.includes(id));
        await DB.setMeta('borrados', resto);
      }
    }
    return regs.length;
  }

  async function bajarCatalogos() {
    const data = await llamar({ action: 'catalogos' });
    // Si la hoja aún no tiene un catálogo y este celular sí, lo sube; si lo tiene, el servidor manda.
    for (const tipo of ['empleados', 'labores', 'camas']) {
      if (data[tipo]?.length) await DB.setMeta(tipo, data[tipo]);
      else {
        const local = await DB.getMeta(tipo, []);
        if (local.length) await subirCatalogo(tipo, local);
      }
    }
    await DB.setMeta('catalogosFecha', new Date().toISOString());
    return data;
  }

  async function subirCatalogo(tipo, datos) {
    return llamar({ action: 'catalogo' }, { action: 'catalogo', tipo, datos });
  }

  async function filas(desde, hasta) {
    const data = await llamar({ action: 'filas', desde, hasta });
    return data.filas || [];
  }

  async function sincronizar({ catalogos = true } = {}) {
    if (enCurso) return;
    const { url } = await conf();
    if (!url) { emitir(await estado()); return; }
    if (!navigator.onLine) { emitir(await estado()); return; }
    enCurso = true;
    emitir({ tipo: 'sincronizando' });
    let error = null, enviados = 0;
    try {
      enviados = await subir();
      if (catalogos) await bajarCatalogos();
      await DB.setMeta('ultimaSync', new Date().toISOString());
    } catch (e) {
      error = e.message || String(e);
    } finally {
      enCurso = false;
    }
    emitir({ ...(await estado()), error, enviados });
    return { error, enviados };
  }

  async function estado() {
    const { url } = await conf();
    const { regs, borrados } = await pendientes();
    return {
      tipo: !url ? 'sin-servidor' : !navigator.onLine ? 'sin-senal' : 'listo',
      pendientes: regs.length + borrados.length,
      ultima: await DB.getMeta('ultimaSync', null),
    };
  }

  window.addEventListener('online', () => sincronizar());
  window.addEventListener('offline', async () => emitir(await estado()));
  setInterval(() => { if (document.visibilityState === 'visible') sincronizar({ catalogos: false }); }, 90000);

  return { sincronizar, estado, subirCatalogo, bajarCatalogos, filas, llamar, onCambio: fn => oyentes.add(fn) };
})();
