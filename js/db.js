// Almacenamiento local en IndexedDB: funciona sin señal y sobrevive al cierre de la app.
// Tiendas: "registros" (uno por labor anotada) y "meta" (catálogos, configuración, borrados pendientes).
const DB = (() => {
  const NAME = 'rendimientos-hydraflora';
  let dbp = null;

  function open() {
    if (dbp) return dbp;
    dbp = new Promise((resolve, reject) => {
      const req = indexedDB.open(NAME, 1);
      req.onupgradeneeded = () => {
        const db = req.result;
        if (!db.objectStoreNames.contains('registros')) {
          const s = db.createObjectStore('registros', { keyPath: 'id' });
          s.createIndex('fecha', 'fecha');
        }
        if (!db.objectStoreNames.contains('meta')) db.createObjectStore('meta');
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
    return dbp;
  }

  async function tx(store, mode, fn) {
    const db = await open();
    return new Promise((resolve, reject) => {
      const t = db.transaction(store, mode);
      const s = t.objectStore(store);
      let out;
      Promise.resolve(fn(s)).then(v => { out = v; });
      t.oncomplete = () => resolve(out);
      t.onerror = () => reject(t.error);
      t.onabort = () => reject(t.error);
    });
  }

  const wrap = req => new Promise((res, rej) => { req.onsuccess = () => res(req.result); req.onerror = () => rej(req.error); });

  return {
    getMeta: (k, def = null) => tx('meta', 'readonly', s => wrap(s.get(k))).then(v => v === undefined ? def : v),
    setMeta: (k, v) => tx('meta', 'readwrite', s => { s.put(v, k); }),
    putRegistro: r => tx('registros', 'readwrite', s => { s.put(r); }),
    putRegistros: rs => tx('registros', 'readwrite', s => { rs.forEach(r => s.put(r)); }),
    deleteRegistro: id => tx('registros', 'readwrite', s => { s.delete(id); }),
    getRegistro: id => tx('registros', 'readonly', s => wrap(s.get(id))),
    allRegistros: () => tx('registros', 'readonly', s => wrap(s.getAll())),
    registrosPorFecha: f => tx('registros', 'readonly', s => wrap(s.index('fecha').getAll(f))),
  };
})();
