// Reglas de cálculo. Replican las fórmulas del Excel de rendimientos:
//   TIEMPO LAB  = (HRA FIN - HRA INICIO) - DESAYUNO - PAUSAS - ALMUERZO   (en horas)
//   RENDIMIENTO = CANTIDAD / TIEMPO LAB
// En una labor grupal la cantidad se reparte en partes iguales (todos comparten el mismo horario),
// y cada persona queda con su propia fila, como en el Excel. Si el registro tiene varios
// bloques/lotes ("tramos"), cada tramo es una fila y el tiempo se reparte según sus plantas.
const Calc = (() => {
  const COLUMNAS = [
    'AÑO', 'MES', 'SEM', 'FECHA', 'AREA', 'LUGAR DE TRABAJO', 'ESTADO', 'NOMBRE', 'CEDULA', 'LABOR',
    'HRA INICIO', 'HRA FIN', 'HORAS TOTALES', 'DESAYUNO', 'PAUSAS', 'ALMUERZO', 'TIEMPO LAB',
    'CANTIDAD', 'RENDIMIENTO', 'OBSERVACIONES', 'MEDIDA', 'BLOQUE', 'LOTE', 'CAMAS', 'N CAMAS',
    'GRUPAL', 'N PERSONAS', 'CANTIDAD GRUPO', 'LIDER', 'ID REGISTRO', 'ID FILA', 'ACTUALIZADO', 'HORAS EXTRAS', 'EXTRAS CONFIRMADAS',
  ];

  // Áreas de la empresa. Los dos cultivos comparten labores pero cada uno tiene su plano de producción.
  // Manzanares no tiene plano: sus labores en plantas se registran solo por tiempo.
  const AREAS = [
    { id: 'CULTIVO SAN JACINTO', grupo: 'CULTIVO', sede: 'SAN JACINTO', horario: 'GENERAL' },
    { id: 'CULTIVO LA CEJA', grupo: 'CULTIVO', sede: 'LA CEJA', horario: 'GENERAL' },
    { id: 'CULTIVO MANZANARES', grupo: 'CULTIVO', horario: 'GENERAL' },
    { id: 'POSCOSECHA SAN JACINTO', grupo: 'POSCOSECHA', horario: 'GENERAL' },
    { id: 'POSCOSECHA LA CEJA', grupo: 'POSCOSECHA', horario: 'CORTO' },
    { id: 'TINTAS', grupo: 'TINTAS', horario: 'CORTO' },
    { id: 'COMERCIALIZADORA', grupo: 'COMERCIALIZADORA', horario: 'CORTO' },
  ];
  const AREA_BASE = 'CULTIVO SAN JACINTO';
  // Nombres de la primera versión del módulo de áreas.
  const AREA_ANTES = { POSCOSECHA: 'POSCOSECHA SAN JACINTO' };
  const areaInfo = id => AREAS.find(a => a.id === (AREA_ANTES[id] || id)) || { id, grupo: id, horario: 'GENERAL' };

  // Horarios por área (días 1 = lunes … 6 = sábado). La jornada esperada de cada día se calcula
  // con la entrada, la salida y los descuentos que caen dentro.
  const HORARIOS = {
    GENERAL: {
      nombre: 'Cultivos y Poscosecha San Jacinto',
      desayunoHora: '08:30', desayunoMin: 26, almuerzoHora: '12:00', almuerzoMin: 25, pausaHora: '10:00', pausaMin: 10,
      dias: { 1: ['06:15', '15:30'], 2: ['06:15', '15:30'], 3: ['06:15', '15:30'], 4: ['06:15', '15:30'], 5: ['06:15', '15:30'] },
    },
    CORTO: {
      nombre: 'Comercializadora, Tintas y Poscosecha La Ceja',
      desayunoHora: '09:00', desayunoMin: 20, almuerzoHora: '12:00', almuerzoMin: 20, pausaHora: '10:00', pausaMin: 10,
      dias: { 1: ['07:00', '15:00'], 2: ['07:00', '16:00'], 3: ['07:00', '16:20'], 4: ['07:00', '16:30'], 5: ['07:00', '16:30'] },
    },
  };
  // Configuración con el horario del área (los cambios hechos en Ajustes van en cfg.horarios).
  function cfgDe(cfg, area) {
    const k = areaInfo(area).horario || 'GENERAL';
    return { ...cfg, ...HORARIOS[k], ...(cfg.horarios?.[k] || {}), horarioId: k };
  }
  const diaSemana = fecha => { const [y, m, d] = fecha.split('-').map(Number); return new Date(Date.UTC(y, m - 1, d)).getUTCDay(); };
  // Festivos de Colombia: fijos, los que se corren al lunes (Ley Emiliani) y los que dependen de la Pascua.
  const festivosCache = {};
  function festivos(y) {
    if (festivosCache[y]) return festivosCache[y];
    const f = d => d.toISOString().slice(0, 10);
    const dia = (m, d) => new Date(Date.UTC(y, m - 1, d));
    const lunes = d => { const x = new Date(d); while (x.getUTCDay() !== 1) x.setUTCDate(x.getUTCDate() + 1); return x; };
    const mas = (d, n) => { const x = new Date(d); x.setUTCDate(x.getUTCDate() + n); return x; };
    // Domingo de Pascua (algoritmo de Meeus/Butcher)
    const a = y % 19, b = Math.floor(y / 100), c = y % 100, dd = Math.floor(b / 4), e = b % 4, g = Math.floor((8 * b + 13) / 25);
    const h = (19 * a + b - dd - g + 15) % 30, i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7;
    const mm = Math.floor((a + 11 * h + 19 * l) / 433), mes = Math.floor((h + l - 7 * mm + 90) / 25), d = (h + l - 7 * mm + 33 * mes + 19) % 32;
    const pascua = dia(mes, d);
    const lista = [
      dia(1, 1), dia(5, 1), dia(7, 20), dia(8, 7), dia(12, 8), dia(12, 25),
      ...[[1, 6], [3, 19], [6, 29], [8, 15], [10, 12], [11, 1], [11, 11]].map(([m2, d2]) => lunes(dia(m2, d2))),
      mas(pascua, -3), mas(pascua, -2), mas(pascua, 43), mas(pascua, 64), mas(pascua, 71),
    ];
    return (festivosCache[y] = new Set(lista.map(f)));
  }
  const esFestivo = fecha => !!fecha && festivos(Number(fecha.slice(0, 4))).has(fecha);
  // Sábado, domingo o festivo: todo el tiempo trabajado es extra.
  const motivoExtra = fecha => !fecha ? '' : esFestivo(fecha) ? 'festivo' : ({ 0: 'domingo', 6: 'sábado' })[diaSemana(fecha)] || '';
  const horarioDia = (cfgArea, fecha) => (esFestivo(fecha) ? null : cfgArea.dias?.[diaSemana(fecha)] || null);
  // Horas productivas esperadas ese día; 0 si el área no trabaja ese día (no se revisa).
  function jornadaDia(cfgArea, fecha) {
    const h = horarioDia(cfgArea, fecha);
    if (!h) return 0;
    return r2(calcular({ horaInicio: h[0], horaFin: h[1], ...descuentosAuto(h[0], h[1], cfgArea), empleados: [] }, cfgArea, null).tiempoLab);
  }
  // Minutos que la hora de fin pasa la hora de salida del horario del área (0 si no la pasa o ese día no tiene horario).
  function minutosExtra(reg, cfgArea) {
    const a = minutos(reg.horaInicio), b = minutos(reg.horaFin);
    if (!reg.fecha || b == null || esPorDias(reg.labor)) return 0;
    if (motivoExtra(reg.fecha)) return a != null && b > a ? b - a : 0;
    const h = horarioDia(cfgArea, reg.fecha);
    return h ? Math.max(0, b - minutos(h[1])) : 0;
  }
  // Horas de esta labor que caen después de la salida (solo el tramo propio de la labor).
  // En sábado, domingo o festivo es todo el tiempo laborado de la labor.
  function horasExtra(reg, cfgArea, tiempoLab = 0) {
    const a = minutos(reg.horaInicio), b = minutos(reg.horaFin);
    if (!reg.fecha || a == null || b == null || b <= a || esPorDias(reg.labor)) return 0;
    if (motivoExtra(reg.fecha)) return tiempoLab;
    const h = cfgArea && horarioDia(cfgArea, reg.fecha);
    return h ? Math.max(0, b - Math.max(a, minutos(h[1]))) / 60 : 0;
  }
  const esCultivo = id => areaInfo(id).grupo === 'CULTIVO';
  const tienePlano = id => !!areaInfo(id).sede;
  // Labores sin área (catálogo anterior) son de cultivo; TODAS aplica a cualquier área.
  const laborEnArea = (l, area) => { const a = norm(l.area) || 'CULTIVO'; return a === 'TODAS' || a === areaInfo(area).grupo; };
  // Camas del plano: [bloque, lote, cama, plantas, variedad, sede]; las del plano anterior son de San Jacinto.
  const sedeCama = c => c[5] || 'SAN JACINTO';

  // Registros antiguos guardaban un solo bloque/lote; los nuevos traen la lista `tramos`.
  function tramosDe(reg) {
    if (reg.tramos?.length) return reg.tramos;
    if (reg.bloque || reg.camas?.length) return [{ bloque: reg.bloque || '', lote: reg.lote || '', camas: reg.camas || [], plantas: null }];
    return [];
  }

  const CONFIG_BASE = {
    desayunoHora: '08:30', desayunoMin: 26,
    almuerzoHora: '12:00', almuerzoMin: 25,
    pausaHora: '10:00', pausaMin: 10,
    jornada: 8.23,
    jornadaInicio: '06:15', jornadaFin: '15:30',
    jornadaSabado: 8.23, pausaDiasMin: 10,
  };

  const norm = s => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().trim();

  const minutos = hhmm => {
    if (!hhmm || !/^\d{1,2}:\d{2}/.test(hhmm)) return null;
    const [h, m] = hhmm.split(':').map(Number);
    return h * 60 + m;
  };

  const r2 = x => Math.round(x * 100) / 100;

  // Tipo de medida según la columna MEDIDA del catálogo de labores.
  function tipoMedida(medida) {
    const m = norm(medida);
    if (m.startsWith('PLANTA')) return 'plantas';
    if (m.startsWith('TALLO')) return 'tallos';
    if (m.startsWith('UNIDAD')) return 'unidades';
    return 'tiempo';
  }
  // Labores que cubren días completos y se pueden registrar por rango de fechas.
  const esPorDias = labor => /VACACION|INCAPACIDAD|AUSENCIAS? POR DIA/.test(norm(labor));

  // Días entre dos fechas (YYYY-MM-DD), sin domingos y con sábados opcionales.
  function diasDelRango(desde, hasta, sabados = false) {
    const out = [];
    const [y, m, d] = desde.split('-').map(Number);
    const dt = new Date(Date.UTC(y, m - 1, d));
    for (let i = 0; i < 400; i++) {
      const s = dt.toISOString().slice(0, 10);
      if (s > hasta) break;
      const dow = dt.getUTCDay();
      if (dow !== 0 && (dow !== 6 || sabados) && !esFestivo(s)) out.push(s);
      dt.setUTCDate(dt.getUTCDate() + 1);
    }
    return out;
  }

  const esGrupalPorMedida = medida => /\bGR\b/.test(norm(medida));

  // Un descuento aplica cuando su hora cae dentro del intervalo trabajado.
  function descuentosAuto(inicio, fin, cfg) {
    const a = minutos(inicio), b = minutos(fin);
    const dentro = h => { const x = minutos(h); return a != null && b != null && x != null && a <= x && x < b; };
    return {
      desayuno: cfg.desayunoMin > 0 && dentro(cfg.desayunoHora),
      almuerzo: cfg.almuerzoMin > 0 && dentro(cfg.almuerzoHora),
      pausa: cfg.pausaMin > 0 && dentro(cfg.pausaHora),
    };
  }

  // "1-10, 15 18" -> [1..10, 15, 18]
  function parseCamas(txt) {
    const out = new Set();
    String(txt || '').split(/[,;\s]+/).forEach(p => {
      if (!p) return;
      const m = p.match(/^(\d+)\s*-\s*(\d+)$/);
      if (m) {
        let a = +m[1], b = +m[2];
        if (a > b) [a, b] = [b, a];
        if (b - a <= 1000) for (let i = a; i <= b; i++) out.add(i);
      } else if (/^\d+$/.test(p)) out.add(+p);
    });
    return [...out].sort((x, y) => x - y);
  }

  function formatCamas(list) {
    const xs = [...new Set(list)].sort((a, b) => a - b);
    const parts = [];
    for (let i = 0; i < xs.length; i++) {
      let j = i;
      while (j + 1 < xs.length && xs[j + 1] === xs[j] + 1) j++;
      parts.push(j > i ? `${xs[i]}-${xs[j]}` : `${xs[i]}`);
      i = j;
    }
    return parts.join(', ');
  }

  // Semana ISO (equivale a WEEKNUM(fecha; 21) de Excel).
  function semanaISO(fechaStr) {
    const [y, m, d] = fechaStr.split('-').map(Number);
    const dt = new Date(Date.UTC(y, m - 1, d));
    const dia = dt.getUTCDay() || 7;
    dt.setUTCDate(dt.getUTCDate() + 4 - dia);
    const ini = new Date(Date.UTC(dt.getUTCFullYear(), 0, 1));
    return Math.ceil(((dt - ini) / 864e5 + 1) / 7);
  }

  // Calcula horas y reparto de un registro. `empleados` = mapa cédula -> empleado.
  function calcular(reg, cfg, empleados) {
    const a = minutos(reg.horaInicio), b = minutos(reg.horaFin);
    const horasTotales = a != null && b != null && b > a ? (b - a) / 60 : 0;
    const desayuno = reg.desayuno ? r2(cfg.desayunoMin / 60) : 0;
    const almuerzo = reg.almuerzo ? r2(cfg.almuerzoMin / 60) : 0;
    const pausa = reg.pausa ? r2((reg.pausaMin ?? cfg.pausaMin) / 60) : 0;
    const tiempoLab = Math.max(0, horasTotales - desayuno - almuerzo - pausa);
    const extras = horasExtra(reg, cfg, tiempoLab);
    const personas = reg.empleados || [];
    const n = personas.length;
    const tipo = tipoMedida(reg.medida);
    const total = tipo === 'tiempo' ? 0 : Number(reg.cantidad) || 0;
    const porPersona = n ? total / n : 0;
    const rendimiento = tiempoLab > 0 && porPersona > 0 ? porPersona / tiempoLab : 0;

    const [y, m] = (reg.fecha || '0000-00').split('-').map(Number);
    const tramos = tramosDe(reg);
    const partes = tramos.length ? tramos : [{ bloque: '', lote: '', camas: [] }];
    const sumaPl = partes.reduce((s, t) => s + (Number(t.plantas) || 0), 0);
    const sumaCa = partes.reduce((s, t) => s + (t.camas?.length || 0), 0);
    const peso = t => (sumaPl > 0 ? (Number(t.plantas) || 0) / sumaPl : sumaCa > 0 ? (t.camas?.length || 0) / sumaCa : 1 / partes.length);
    const filas = personas.flatMap(ced => partes.map((t, i) => {
      const e = (empleados && empleados[ced]) || { nombre: reg.nombres?.[ced] || ced, cedula: ced };
      const w = peso(t);
      return {
        'AÑO': y, 'MES': m, 'SEM': reg.fecha ? semanaISO(reg.fecha) : '', 'FECHA': reg.fecha, 'AREA': reg.area || '',
        'LUGAR DE TRABAJO': e.lugar || '', 'ESTADO': e.estado || '', 'NOMBRE': e.nombre, 'CEDULA': String(ced),
        'LABOR': reg.labor, 'HRA INICIO': reg.horaInicio, 'HRA FIN': reg.horaFin,
        'HORAS TOTALES': r2(horasTotales * w), 'DESAYUNO': r2(desayuno * w), 'PAUSAS': r2(pausa * w), 'ALMUERZO': r2(almuerzo * w),
        'TIEMPO LAB': r2(tiempoLab * w), 'CANTIDAD': r2(porPersona * w), 'RENDIMIENTO': r2(rendimiento),
        'OBSERVACIONES': reg.obs || '', 'MEDIDA': reg.medida || '', 'BLOQUE': t.bloque || '',
        'LOTE': t.lote || '', 'CAMAS': formatCamas(t.camas || []), 'N CAMAS': (t.camas || []).length,
        'GRUPAL': n > 1 ? 'SI' : 'NO', 'N PERSONAS': n, 'CANTIDAD GRUPO': total, 'LIDER': reg.lider || '',
        'ID REGISTRO': reg.id, 'ID FILA': partes.length > 1 ? `${reg.id}-${ced}-${i + 1}` : `${reg.id}-${ced}`, 'ACTUALIZADO': reg.actualizado || '',
        'HORAS EXTRAS': r2(extras * w), 'EXTRAS CONFIRMADAS': reg.extrasOk?.[ced] ? 'SI' : '',
      };
    }));
    return { horasTotales, desayuno, almuerzo, pausa, tiempoLab, extras, n, tipo, total, porPersona, rendimiento, filas };
  }

  // Devuelve una lista de errores en español; vacía si el registro se puede guardar.
  // Un registro parcial (sin terminar) solo necesita fecha, labor, personas y hora de inicio.
  function validar(reg, cfg, { parcial = false } = {}) {
    const e = [];
    if (!reg.fecha) e.push('Falta la fecha.');
    if (!reg.labor) e.push('Elige una labor de la lista.');
    if (!reg.empleados?.length) e.push('Agrega al menos un empleado.');
    const a = minutos(reg.horaInicio), b = minutos(reg.horaFin);
    if (parcial) {
      if (a == null) e.push('Falta la hora de inicio.');
      if (a != null && b != null && b <= a) e.push('La hora de fin debe ser posterior a la de inicio.');
      return e;
    }
    if (a == null || b == null) e.push('Faltan la hora de inicio y la de fin.');
    else if (b <= a) e.push('La hora de fin debe ser posterior a la de inicio.');
    const tipo = tipoMedida(reg.medida);
    const tramos = tramosDe(reg);
    if (tipo === 'plantas') {
      if (!tramos.length || tramos.some(t => !t.bloque || !t.lote)) e.push('Esta labor se mide en plantas: elige bloque y lote.');
      if (!tramos.some(t => t.camas?.length)) e.push('Esta labor se mide en plantas: marca las camas realizadas.');
      if (!(Number(reg.cantidad) > 0)) e.push('Esta labor se mide en plantas: el número de plantas es obligatorio.');
    }
    if (tipo === 'tallos' && !(Number(reg.cantidad) > 0)) e.push('Esta labor se mide en tallos: digita los tallos.');
    if (tipo === 'unidades' && !(Number(reg.cantidad) > 0)) e.push('Esta labor se mide en unidades: digita las unidades.');
    const extra = cfg ? minutosExtra(reg, cfg) : 0;
    if (extra > 0 && !reg.extras) e.push(motivoExtra(reg.fecha)
      ? `El ${reg.fecha.split('-').reverse().join('/')} es ${motivoExtra(reg.fecha)}: todo el tiempo es extra, marca "Horas extras" para guardarlo.`
      : `La hora de fin pasa ${extra} min la salida de las ${horarioDia(cfg, reg.fecha)[1]}: marca "Horas extras" para guardarlo.`);
    if (cfg && a != null && b != null && b > a) {
      const c = calcular(reg, cfg, null);
      if (c.tiempoLab <= 0) e.push('Con los descuentos marcados el tiempo laborado queda en cero.');
    }
    return e;
  }

  // Horarios cruzados: una persona no puede estar en dos labores a la vez el mismo día.
  // Un registro sin terminar ocupa desde su hora de inicio en adelante.
  function cruces(reg, otros, nombre = c => c) {
    const fin = r => (r.estado === 'abierto' || minutos(r.horaFin) == null ? Infinity : minutos(r.horaFin));
    const a = minutos(reg.horaInicio), b = fin(reg);
    if (a == null) return [];
    const msgs = [];
    otros.filter(o => o.id !== reg.id && o.fecha === reg.fecha).forEach(o => {
      const oa = minutos(o.horaInicio), ob = fin(o);
      if (oa == null || !(a < ob && oa < b)) return;
      reg.empleados.filter(c => o.empleados.includes(c)).forEach(c => {
        msgs.push(o.estado === 'abierto'
          ? `${nombre(c)} tiene ${o.labor} sin terminar desde las ${o.horaInicio}. Termínalo antes de asignarle otra labor.`
          : `${nombre(c)} ya está en ${o.labor} de ${o.horaInicio} a ${o.horaFin}; el horario se cruza.`);
      });
    });
    return [...new Set(msgs)];
  }

  // Antes de empezar un día, el último día trabajado de cada persona debe tener la jornada completa.
  // Se compara el tiempo laborado (ya sin desayuno ni almuerzo) con las horas esperadas.
  function diaIncompleto(reg, otros, cfg, nombre = c => c) {
    const msgs = [];
    const previos = otros.filter(o => o.id !== reg.id && o.estado !== 'abierto' && o.fecha < reg.fecha);
    reg.empleados.forEach(c => {
      const suyos = previos.filter(o => o.empleados.includes(c));
      if (!suyos.length) return;
      const ultimo = suyos.reduce((m, o) => (o.fecha > m ? o.fecha : m), '');
      // Un día de vacaciones, incapacidad o ausencia cuenta como completo.
      if (suyos.some(o => o.fecha === ultimo && esPorDias(o.labor))) return;
      const horas = suyos.filter(o => o.fecha === ultimo).reduce((s, o) => s + calcular(o, cfg, null).tiempoLab, 0);
      const [y, m, d] = ultimo.split('-').map(Number);
      const sabado = new Date(Date.UTC(y, m - 1, d)).getUTCDay() === 6;
      const meta = Number(sabado ? cfg.jornadaSabado ?? cfg.jornada : cfg.jornada) || 0;
      if (meta > 0 && horas < meta - 0.02) msgs.push(`${nombre(c)} tiene ${r2(horas)} h de ${meta} h el ${ultimo}. Completa ese día antes de registrar el ${reg.fecha}.`);
    });
    return msgs;
  }

  // Días anteriores a `fecha` en que alguien trabajó más que la jornada (horas extras) y el líder
  // aún no lo confirmó. Se agrupa por día y persona.
  function extrasPorConfirmar(todos, fecha, cfg) {
    const g = new Map();
    todos.filter(r => r.fecha && r.fecha < fecha && r.estado !== 'abierto' && !esPorDias(r.labor)).forEach(r => {
      const c = calcular(r, cfg, null);
      r.empleados.forEach(ced => {
        const k = `${r.fecha}|${ced}`;
        const p = g.get(k) || g.set(k, { fecha: r.fecha, ced, nombre: r.nombres?.[ced] || ced, horas: 0, extras: 0, regs: [] }).get(k);
        p.horas += c.tiempoLab; p.extras += c.extras; p.regs.push(r);
      });
    });
    return [...g.values()].map(p => {
      const meta = jornadaDia(cfg, p.fecha);
      const extra = Math.max(p.extras, meta > 0 ? p.horas - meta : 0);
      return { ...p, horas: r2(p.horas), meta, extra: r2(extra) };
    }).filter(p => p.extra > 0.02 && !p.regs.every(r => r.extrasOk?.[p.ced]))
      .sort((a, b) => a.fecha.localeCompare(b.fecha) || a.nombre.localeCompare(b.nombre));
  }

  // Filas del servidor de la misma persona y el mismo día que registros de este celular, hechas en otro
  // celular u otro líder (otro ID de registro).
  function crucesServidor(locales, filasServidor) {
    const ids = new Set(locales.map(r => r.id));
    const mias = new Map();
    locales.forEach(r => r.empleados.forEach(c => { const k = `${r.fecha}|${c}`; (mias.get(k) || mias.set(k, []).get(k)).push(r); }));
    return filasServidor.filter(f => !ids.has(String(f['ID REGISTRO'])) && mias.has(`${String(f.FECHA).slice(0, 10)}|${String(f.CEDULA)}`))
      .map(f => {
        const fecha = String(f.FECHA).slice(0, 10), mis = mias.get(`${fecha}|${String(f.CEDULA)}`);
        const a = minutos(String(f['HRA INICIO'])), b = minutos(String(f['HRA FIN']));
        const cruza = a != null && b != null && mis.some(r => { const x = minutos(r.horaInicio), y = minutos(r.horaFin) ?? 24 * 60; return x != null && a < y && x < b; });
        return { id: String(f['ID FILA']), fecha, ced: String(f.CEDULA), nombre: f.NOMBRE, lider: f.LIDER || 'sin nombre', area: f.AREA || '',
          labor: f.LABOR, inicio: String(f['HRA INICIO']), fin: String(f['HRA FIN']), cruza };
      });
  }

  // Personas que no completaron la jornada de un día o que dejaron labores sin terminar.
  function incompletosDelDia(delDia, fecha, cfg) {
    const meta = jornadaDia(cfg, fecha);
    const ceds = [...new Set(delDia.flatMap(o => o.empleados))];
    return ceds.map(c => {
      const suyos = delDia.filter(o => o.empleados.includes(c));
      const abiertos = suyos.filter(o => o.estado === 'abierto');
      const horas = suyos.filter(o => o.estado !== 'abierto').reduce((s, o) => s + calcular(o, cfg, null).tiempoLab, 0);
      const porDias = suyos.some(o => esPorDias(o.labor));
      const nombre = suyos.find(o => o.nombres?.[c])?.nombres[c] || c;
      return { ced: c, nombre, horas: r2(horas), meta, abiertos, completo: !abiertos.length && (porDias || meta <= 0 || horas >= meta - 0.02) };
    }).filter(p => !p.completo);
  }

  // Días que impiden escribir en `fecha`: se trabaja un día a la vez. Solo se puede escribir en el
  // día incompleto más antiguo; si no hay ninguno, en cualquier fecha.
  function diasQueBloquean(todos, fecha, cfg, excluir = null) {
    const porFecha = new Map();
    todos.forEach(o => { if (o.id !== excluir && o.fecha) (porFecha.get(o.fecha) || porFecha.set(o.fecha, []).get(o.fecha)).push(o); });
    const inc = [...porFecha].map(([f, regs]) => ({ fecha: f, personas: incompletosDelDia(regs, f, cfg) }))
      .filter(d => d.personas.length).sort((a, b) => a.fecha.localeCompare(b.fecha));
    if (!inc.length || inc[0].fecha === fecha) return [];
    return inc.filter(d => d.fecha !== fecha);
  }

  // Calculadora: evalúa sumas, restas, multiplicaciones, divisiones y paréntesis sin usar eval.
  // Devuelve null si la cuenta está incompleta o no es válida.
  function evaluar(expr) {
    const toks = String(expr).replace(/,/g, '.').match(/\d+(?:\.\d+)?|\.\d+|[-+*/()]/g);
    if (!toks || toks.join('') !== String(expr).replace(/,/g, '.').replace(/\s+/g, '')) return null;
    let i = 0;
    const num = () => {
      const t = toks[i++];
      if (t === '(') { const v = suma(); if (toks[i++] !== ')') throw 0; return v; }
      if (t === '-') return -num();
      if (t === undefined || !/^[\d.]/.test(t)) throw 0;
      return Number(t);
    };
    const prod = () => { let v = num(); while (toks[i] === '*' || toks[i] === '/') { const o = toks[i++], b = num(); v = o === '*' ? v * b : v / b; } return v; };
    const suma = () => { let v = prod(); while (toks[i] === '+' || toks[i] === '-') { const o = toks[i++], b = prod(); v = o === '+' ? v + b : v - b; } return v; };
    try { const v = suma(); return i === toks.length && Number.isFinite(v) ? Math.round(v * 100) / 100 : null; } catch { return null; }
  }

  // Busca empleados por cualquier parte del nombre o cédula. Todas las palabras deben coincidir.
  function buscar(lista, q, campo = 'nombre', extra = []) {
    const toks = norm(q).split(/\s+/).filter(Boolean);
    if (!toks.length) return [];
    return lista.filter(it => {
      const hay = norm([it[campo], ...extra.map(k => it[k])].join(' '));
      return toks.every(t => hay.includes(t));
    });
  }

  return { evaluar, COLUMNAS, AREAS, AREA_BASE, AREA_ANTES, HORARIOS, cfgDe, horarioDia, jornadaDia, areaInfo, esCultivo, tienePlano, laborEnArea, sedeCama, CONFIG_BASE, tramosDe, cruces, diaIncompleto, diasQueBloquean, incompletosDelDia, minutosExtra, horasExtra, extrasPorConfirmar, crucesServidor, esFestivo, motivoExtra, esPorDias, diasDelRango, norm, minutos, tipoMedida, esGrupalPorMedida, descuentosAuto, parseCamas, formatCamas, semanaISO, calcular, validar, buscar, r2 };
})();
if (typeof module !== 'undefined') module.exports = Calc;
