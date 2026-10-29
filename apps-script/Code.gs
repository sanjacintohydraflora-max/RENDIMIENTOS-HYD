/**
 * Servidor de la app de rendimientos Hydraflora (Google Apps Script).
 *
 * Instalación (una sola vez):
 *  1. Crea una Google Sheet vacía. Menú Extensiones > Apps Script.
 *  2. Pega este archivo completo en Code.gs y guarda.
 *  3. Cambia CLAVE_INICIAL abajo por una clave tuya y ejecuta la función `configurar`
 *     (acepta los permisos). Crea las pestañas RENDIMIENTOS, EMPLEADOS, LABORES y CAMAS.
 *  4. Implementar > Nueva implementación > Tipo: Aplicación web.
 *     Ejecutar como: Yo. Quién tiene acceso: Cualquier persona.
 *  5. Copia la URL que termina en /exec y pégala en la app (Ajustes > Sincronización) con la clave.
 *
 * La hoja RENDIMIENTOS tiene las mismas columnas del Excel de rendimientos, así que puedes
 * hacer tablas dinámicas o descargarla a Excel cuando quieras.
 */

const CLAVE_INICIAL = 'cambia-esta-clave';

const HOJAS = {
  RENDIMIENTOS: [
    'AÑO', 'MES', 'SEM', 'FECHA', 'LUGAR DE TRABAJO', 'ESTADO', 'NOMBRE', 'CEDULA', 'LABOR',
    'HRA INICIO', 'HRA FIN', 'HORAS TOTALES', 'DESAYUNO', 'PAUSAS', 'ALMUERZO', 'TIEMPO LAB',
    'CANTIDAD', 'RENDIMIENTO', 'OBSERVACIONES', 'MEDIDA', 'BLOQUE', 'LOTE', 'CAMAS', 'N CAMAS',
    'GRUPAL', 'N PERSONAS', 'CANTIDAD GRUPO', 'LIDER', 'ID REGISTRO', 'ID FILA', 'ACTUALIZADO',
  ],
  EMPLEADOS: ['NOMBRE', 'ESTADO', 'CEDULA', 'LUGARDE TRABAJO', 'CARGO', 'AREA'],
  LABORES: ['LABORES', 'MEDIDA'],
  CAMAS: ['BLOQUE', 'LOTE', 'ERA', 'CANTIDAD DE PLANTAS', 'VARIEDAD'],
};
// Columnas que se guardan como texto para que Sheets no las convierta en fecha o número.
const COLUMNAS_TEXTO = ['CEDULA', 'HRA INICIO', 'HRA FIN', 'LOTE', 'CAMAS', 'ID REGISTRO', 'ID FILA', 'ACTUALIZADO'];

function configurar() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('CLAVE')) props.setProperty('CLAVE', CLAVE_INICIAL);
  const ss = SpreadsheetApp.getActive();
  Object.keys(HOJAS).forEach(nombre => {
    const h = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
    const enc = HOJAS[nombre];
    h.getRange(1, 1, 1, enc.length).setValues([enc]).setFontWeight('bold');
    h.setFrozenRows(1);
    enc.forEach((c, i) => {
      if (COLUMNAS_TEXTO.indexOf(c) >= 0) h.getRange(2, i + 1, h.getMaxRows() - 1, 1).setNumberFormat('@');
    });
  });
  const hoja1 = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (hoja1 && ss.getSheets().length > 1) ss.deleteSheet(hoja1);
}

/** Cambia la clave sin tocar el código: ejecuta cambiarClave('nueva') desde el editor. */
function cambiarClave(nueva) {
  PropertiesService.getScriptProperties().setProperty('CLAVE', nueva);
}

function responder(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function autorizado(key) {
  const clave = PropertiesService.getScriptProperties().getProperty('CLAVE') || CLAVE_INICIAL;
  return key && key === clave;
}

function doGet(e) {
  try {
    const p = e.parameter || {};
    if (!autorizado(p.key)) return responder({ ok: false, error: 'Clave incorrecta' });
    if (p.action === 'catalogos') return responder({ ok: true, ...catalogos() });
    if (p.action === 'filas') return responder({ ok: true, filas: filas(p.desde, p.hasta) });
    return responder({ ok: true, ping: new Date().toISOString() });
  } catch (err) {
    return responder({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const body = JSON.parse(e.postData.contents || '{}');
    if (!autorizado(body.key)) return responder({ ok: false, error: 'Clave incorrecta' });
    if (body.action === 'guardar') return responder({ ok: true, ...guardar(body.registros || [], body.borrados || []) });
    if (body.action === 'catalogo') return responder({ ok: true, n: guardarCatalogo(body.tipo, body.datos || []) });
    return responder({ ok: false, error: 'Acción desconocida' });
  } catch (err) {
    return responder({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function hoja(nombre) {
  const h = SpreadsheetApp.getActive().getSheetByName(nombre);
  if (!h) throw new Error('Falta la pestaña ' + nombre + '. Ejecuta configurar().');
  return h;
}

function leer(nombre) {
  const h = hoja(nombre);
  const n = h.getLastRow() - 1;
  if (n < 1) return { h: h, enc: HOJAS[nombre], datos: [] };
  const v = h.getRange(1, 1, n + 1, h.getLastColumn()).getValues();
  return { h: h, enc: v[0].map(String), datos: v.slice(1) };
}

/** Reemplaza las filas de cada registro recibido (así una edición no duplica) y borra los eliminados. */
function guardar(registros, borrados) {
  const t = leer('RENDIMIENTOS');
  const enc = HOJAS.RENDIMIENTOS;
  const iReg = t.enc.indexOf('ID REGISTRO');
  const quitar = {};
  registros.forEach(r => { quitar[r.id] = true; });
  borrados.forEach(id => { quitar[id] = true; });

  // Borra de abajo hacia arriba agrupando filas seguidas.
  const aBorrar = [];
  t.datos.forEach((fila, i) => { if (quitar[fila[iReg]]) aBorrar.push(i + 2); });
  for (let k = aBorrar.length - 1; k >= 0;) {
    let fin = aBorrar[k], ini = fin;
    while (k - 1 >= 0 && aBorrar[k - 1] === ini - 1) { ini--; k--; }
    t.h.deleteRows(ini, fin - ini + 1);
    k--;
  }

  const nuevas = [];
  registros.forEach(r => (r.filas || []).forEach(f => {
    nuevas.push(enc.map(c => {
      // Celulares con la versión anterior mandan la columna como LUGARDE TRABAJO.
      const v = c === 'LUGAR DE TRABAJO' && f[c] === undefined ? f['LUGARDE TRABAJO'] : f[c];
      if (c === 'FECHA' && v) { const p = String(v).split('-'); return new Date(+p[0], +p[1] - 1, +p[2]); }
      return v === undefined || v === null ? '' : v;
    }));
  }));
  if (nuevas.length) t.h.getRange(t.h.getLastRow() + 1, 1, nuevas.length, enc.length).setValues(nuevas);
  return { reemplazadas: aBorrar.length, agregadas: nuevas.length };
}

function texto(v, col) {
  if (v instanceof Date) {
    const tz = Session.getScriptTimeZone();
    if (col === 'HRA INICIO' || col === 'HRA FIN') return Utilities.formatDate(v, tz, 'HH:mm');
    return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  }
  return v;
}

function filas(desde, hasta) {
  const t = leer('RENDIMIENTOS');
  return t.datos
    .map(fila => {
      const o = {};
      t.enc.forEach((c, i) => { o[c] = texto(fila[i], c); });
      o.CEDULA = String(o.CEDULA);
      o.LOTE = String(o.LOTE);
      return o;
    })
    .filter(o => o['ID FILA'] !== '' && (!desde || String(o.FECHA) >= desde) && (!hasta || String(o.FECHA) <= hasta));
}

function catalogos() {
  const e = leer('EMPLEADOS'), l = leer('LABORES'), c = leer('CAMAS');
  const col = (t, n) => t.enc.indexOf(n);
  return {
    empleados: e.datos.filter(f => f[col(e, 'NOMBRE')] !== '').map(f => ({
      nombre: String(f[col(e, 'NOMBRE')]).trim(), estado: String(f[col(e, 'ESTADO')]).trim(),
      cedula: String(f[col(e, 'CEDULA')]).trim(), lugar: String(f[col(e, 'LUGARDE TRABAJO')]).trim(),
      cargo: String(f[col(e, 'CARGO')]).trim(), area: String(f[col(e, 'AREA')]).trim(),
    })),
    labores: l.datos.filter(f => f[0] !== '').map(f => ({ nombre: String(f[0]).trim(), medida: String(f[1]).trim() })),
    camas: c.datos.filter(f => f[0] !== '').map(f => [String(f[0]).trim(), String(f[1]).trim(), Number(f[2]), Number(f[3]) || 0, String(f[4]).trim()]),
  };
}

function guardarCatalogo(tipo, datos) {
  const nombre = { empleados: 'EMPLEADOS', labores: 'LABORES', camas: 'CAMAS' }[tipo];
  if (!nombre) throw new Error('Catálogo desconocido');
  const h = hoja(nombre);
  const enc = HOJAS[nombre];
  const filasNuevas = datos.map(d => {
    if (tipo === 'empleados') return [d.nombre, d.estado, String(d.cedula), d.lugar || '', d.cargo || '', d.area || ''];
    if (tipo === 'labores') return [d.nombre, d.medida];
    return [d[0], String(d[1]), d[2], d[3], d[4]];
  });
  if (h.getLastRow() > 1) h.getRange(2, 1, h.getLastRow() - 1, enc.length).clearContent();
  if (filasNuevas.length) h.getRange(2, 1, filasNuevas.length, enc.length).setValues(filasNuevas);
  return filasNuevas.length;
}
