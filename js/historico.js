// Carga de rendimientos históricos: los Excel que se llenaban antes de la app (hoja RENDIMIENTOS de
// "03 RENDIMIENTOS POSCOSECHAS" o "04 RENDIMIENTOS COMERCIALIZADORA") se convierten a las columnas de la
// app y se envían a la hoja de Google, para que el dashboard los muestre.
const Historico = (() => {
  const norm = s => Calc.norm(s).replace(/\s+/g, ' ').trim();
  const dos = n => String(n).padStart(2, '0');

  // Fecha de Excel (número de serie o texto) a AAAA-MM-DD.
  function fechaDe(v) {
    if (typeof v === 'number') { const d = XLSX.SSF.parse_date_code(v); return d ? `${d.y}-${dos(d.m)}-${dos(d.d)}` : ''; }
    const s = String(v || '').trim();
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return `${m[1]}-${dos(m[2])}-${dos(m[3])}`;
    m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    return m ? `${m[3]}-${dos(m[2])}-${dos(m[1])}` : '';
  }

  // Hora de Excel (fracción del día o texto "7:00", "07:00:00", "2:30 PM") a HH:MM.
  function horaDe(v) {
    if (typeof v === 'number') { const min = Math.round((v % 1) * 1440); return `${dos(Math.floor(min / 60) % 24)}:${dos(min % 60)}`; }
    const m = String(v || '').trim().match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*([AaPp])?/);
    if (!m) return '';
    let h = +m[1];
    if (m[3] && /p/i.test(m[3]) && h < 12) h += 12;
    if (m[3] && /a/i.test(m[3]) && h === 12) h = 0;
    return `${dos(h)}:${m[2]}`;
  }

  // Identificador estable: si se carga dos veces el mismo archivo, la hoja reemplaza las filas en vez de duplicarlas.
  function hash(s) {
    let h1 = 0x811c9dc5, h2 = 0x1234567;
    for (let i = 0; i < s.length; i++) { const c = s.charCodeAt(i); h1 = Math.imul(h1 ^ c, 16777619); h2 = Math.imul(h2 ^ c, 2246822519); }
    return (h1 >>> 0).toString(36) + (h2 >>> 0).toString(36);
  }

  const num = v => { const n = Number(String(v ?? '').replace(',', '.')); return Number.isFinite(n) ? n : 0; };

  // Área de cada fila: el archivo de comercializadora es todo de Comercializadora. En el de poscosechas,
  // las labores que solo existen en Tintas van a Tintas; el resto, según el lugar (Sonsón = San Jacinto).
  function areaDe(tipoArchivo, lugar, labor, labores) {
    if (tipoArchivo === 'COMERCIALIZADORA') return 'COMERCIALIZADORA';
    const areas = labores.filter(l => norm(l.nombre) === norm(labor)).map(l => norm(l.area));
    if (areas.length && areas.every(a => a === 'TINTAS')) return 'TINTAS';
    return /SONSON|SAN JACINTO|FSJ/.test(norm(lugar)) ? 'POSCOSECHA SAN JACINTO' : 'POSCOSECHA LA CEJA';
  }

  function convertir(wb, nombreArchivo, labores) {
    const hoja = wb.SheetNames.find(n => norm(n) === 'RENDIMIENTOS') || wb.SheetNames.find(n => norm(n).includes('RENDIMIENTO'));
    if (!hoja) throw new Error('El archivo no tiene una hoja RENDIMIENTOS.');
    const filas = XLSX.utils.sheet_to_json(wb.Sheets[hoja], { header: 1, defval: '', raw: true });
    const iEnc = filas.findIndex(f => f.some(c => norm(c) === 'CEDULA') && f.some(c => norm(c) === 'FECHA'));
    if (iEnc < 0) throw new Error('No encontré los encabezados FECHA y CEDULA en la hoja RENDIMIENTOS.');
    const enc = filas[iEnc].map(c => norm(c).replace(/ /g, ''));
    const col = (...nombres) => enc.findIndex(h => nombres.some(n => h === n.replace(/ /g, '')));
    const c = {
      fecha: col('FECHA'), lugar: col('LUGARDE TRABAJO', 'LUGAR DE TRABAJO'), estado: col('ESTADO'), nombre: col('NOMBRE'),
      cedula: col('CEDULA'), labor: col('LABOR', 'LABORES'), ini: col('HRA INICIO', 'HORA INICIO'), fin: col('HRA FIN', 'HORA FIN'),
      total: col('HORAS TOTALES'), des: col('DESAYUNO'), pau: col('PAUSAS', 'PAUSA'), alm: col('ALMUERZO'),
      lab: col('TIEMPO LAB'), cant: col('TALLO', 'TALLOS', 'CANTIDAD'), rend: col('RENDIMIENTO'), obs: col('OBSERVACIONES'),
    };
    if ([c.fecha, c.cedula, c.labor, c.lab].some(i => i < 0)) throw new Error('A la hoja RENDIMIENTOS le faltan columnas (FECHA, CEDULA, LABOR o TIEMPO LAB).');
    const tipoArchivo = /COMERCIALIZ/.test(norm(nombreArchivo)) ? 'COMERCIALIZADORA' : 'POSCOSECHA';
    const vistos = new Map();
    const salida = [];
    let omitidas = 0;
    filas.slice(iEnc + 1).forEach(f => {
      const fecha = fechaDe(f[c.fecha]), ced = String(f[c.cedula] ?? '').trim().replace(/\.0$/, '');
      const labor = String(f[c.labor] ?? '').trim().replace(/\s+/g, ' ');
      if (!fecha || !ced || !labor || labor === 'None') { if (fecha || ced) omitidas++; return; }
      const area = areaDe(tipoArchivo, f[c.lugar], labor, labores);
      const cat = labores.find(l => norm(l.nombre) === norm(labor) && (norm(l.area) === Calc.areaInfo(area).grupo || norm(l.area) === area.split(' ')[0]));
      const cantidad = c.cant >= 0 ? num(f[c.cant]) : 0;
      const medida = cat?.medida || (cantidad > 0 ? 'Tallos' : 'Tiempo');
      const ini = c.ini >= 0 ? horaDe(f[c.ini]) : '', fin = c.fin >= 0 ? horaDe(f[c.fin]) : '';
      const base = `H|${fecha}|${ced}|${norm(labor)}|${ini}|${fin}|${cantidad}`;
      const rep = (vistos.get(base) || 0) + 1;
      vistos.set(base, rep);
      const id = `H${hash(`${base}|${rep}`)}`;
      const [y, m] = fecha.split('-').map(Number);
      salida.push({
        id,
        filas: [{
          'AÑO': y, 'MES': m, 'SEM': Calc.semanaISO(fecha), 'FECHA': fecha, 'AREA': area,
          'LUGAR DE TRABAJO': String(f[c.lugar] ?? '').trim(), 'LUGARDE TRABAJO': String(f[c.lugar] ?? '').trim(),
          'ESTADO': String(f[c.estado] ?? '').trim(), 'NOMBRE': String(f[c.nombre] ?? '').trim().replace(/\s+/g, ' '), 'CEDULA': ced,
          'LABOR': labor, 'HRA INICIO': ini, 'HRA FIN': fin,
          'HORAS TOTALES': Calc.r2(num(f[c.total])), 'DESAYUNO': Calc.r2(num(f[c.des])), 'PAUSAS': Calc.r2(num(f[c.pau])), 'ALMUERZO': Calc.r2(num(f[c.alm])),
          'TIEMPO LAB': Calc.r2(num(f[c.lab])), 'CANTIDAD': cantidad, 'RENDIMIENTO': Calc.r2(num(f[c.rend])),
          'OBSERVACIONES': String(f[c.obs] ?? '').trim(), 'MEDIDA': medida, 'BLOQUE': '', 'LOTE': '', 'CAMAS': '', 'N CAMAS': 0,
          'GRUPAL': 'NO', 'N PERSONAS': 1, 'CANTIDAD GRUPO': cantidad, 'LIDER': 'HISTORICO EXCEL',
          'ID REGISTRO': id, 'ID FILA': `${id}-${ced}`, 'ACTUALIZADO': new Date().toISOString(), 'HORAS EXTRAS': '', 'EXTRAS CONFIRMADAS': '',
        }],
      });
    });
    const fechas = salida.map(r => r.filas[0].FECHA).sort();
    const porArea = {};
    salida.forEach(r => { porArea[r.filas[0].AREA] = (porArea[r.filas[0].AREA] || 0) + 1; });
    return { registros: salida, omitidas, desde: fechas[0], hasta: fechas.at(-1), porArea };
  }

  // Envía a la hoja en lotes. `avance(n, total)` informa el progreso.
  async function enviar(registros, avance) {
    const lote = 800;
    for (let i = 0; i < registros.length; i += lote) {
      await Sync.llamar({ action: 'guardar' }, { action: 'guardar', registros: registros.slice(i, i + lote), borrados: [] });
      avance(Math.min(i + lote, registros.length), registros.length);
    }
  }

  return { convertir, enviar, fechaDe, horaDe };
})();
