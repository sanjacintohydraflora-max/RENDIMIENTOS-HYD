// Reglas de cálculo. Replican las fórmulas del Excel de rendimientos:
//   TIEMPO LAB  = (HRA FIN - HRA INICIO) - DESAYUNO - PAUSAS - ALMUERZO   (en horas)
//   RENDIMIENTO = CANTIDAD / TIEMPO LAB
// En una labor grupal la cantidad se reparte en partes iguales (todos comparten el mismo horario),
// y cada persona queda con su propia fila, como en el Excel. Si el registro tiene varios
// bloques/lotes ("tramos"), cada tramo es una fila y el tiempo se reparte según sus plantas.
const Calc = (() => {
  const COLUMNAS = [
    'AÑO', 'MES', 'SEM', 'FECHA', 'LUGAR DE TRABAJO', 'ESTADO', 'NOMBRE', 'CEDULA', 'LABOR',
    'HRA INICIO', 'HRA FIN', 'HORAS TOTALES', 'DESAYUNO', 'PAUSAS', 'ALMUERZO', 'TIEMPO LAB',
    'CANTIDAD', 'RENDIMIENTO', 'OBSERVACIONES', 'MEDIDA', 'BLOQUE', 'LOTE', 'CAMAS', 'N CAMAS',
    'GRUPAL', 'N PERSONAS', 'CANTIDAD GRUPO', 'LIDER', 'ID REGISTRO', 'ID FILA', 'ACTUALIZADO',
  ];

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
      if (dow !== 0 && (dow !== 6 || sabados)) out.push(s);
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
        'AÑO': y, 'MES': m, 'SEM': reg.fecha ? semanaISO(reg.fecha) : '', 'FECHA': reg.fecha,
        'LUGAR DE TRABAJO': e.lugar || '', 'ESTADO': e.estado || '', 'NOMBRE': e.nombre, 'CEDULA': String(ced),
        'LABOR': reg.labor, 'HRA INICIO': reg.horaInicio, 'HRA FIN': reg.horaFin,
        'HORAS TOTALES': r2(horasTotales * w), 'DESAYUNO': r2(desayuno * w), 'PAUSAS': r2(pausa * w), 'ALMUERZO': r2(almuerzo * w),
        'TIEMPO LAB': r2(tiempoLab * w), 'CANTIDAD': r2(porPersona * w), 'RENDIMIENTO': r2(rendimiento),
        'OBSERVACIONES': reg.obs || '', 'MEDIDA': reg.medida || '', 'BLOQUE': t.bloque || '',
        'LOTE': t.lote || '', 'CAMAS': formatCamas(t.camas || []), 'N CAMAS': (t.camas || []).length,
        'GRUPAL': n > 1 ? 'SI' : 'NO', 'N PERSONAS': n, 'CANTIDAD GRUPO': total, 'LIDER': reg.lider || '',
        'ID REGISTRO': reg.id, 'ID FILA': partes.length > 1 ? `${reg.id}-${ced}-${i + 1}` : `${reg.id}-${ced}`, 'ACTUALIZADO': reg.actualizado || '',
      };
    }));
    return { horasTotales, desayuno, almuerzo, pausa, tiempoLab, n, tipo, total, porPersona, rendimiento, filas };
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

  // Busca empleados por cualquier parte del nombre o cédula. Todas las palabras deben coincidir.
  function buscar(lista, q, campo = 'nombre', extra = []) {
    const toks = norm(q).split(/\s+/).filter(Boolean);
    if (!toks.length) return [];
    return lista.filter(it => {
      const hay = norm([it[campo], ...extra.map(k => it[k])].join(' '));
      return toks.every(t => hay.includes(t));
    });
  }

  return { COLUMNAS, CONFIG_BASE, tramosDe, cruces, diaIncompleto, esPorDias, diasDelRango, norm, minutos, tipoMedida, esGrupalPorMedida, descuentosAuto, parseCamas, formatCamas, semanaISO, calcular, validar, buscar, r2 };
})();
if (typeof module !== 'undefined') module.exports = Calc;
