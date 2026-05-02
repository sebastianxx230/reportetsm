(() => {
  const $ = s => document.querySelector(s);

  const MAESTRO_PERSONAL = {
    "lamper torres jose ali": { empresa: "Top Solution Metal SAC", contrato: "Planilla" },
    "munoz garcia miguel angel": { empresa: "Top Solution Metal SAC", contrato: "Planilla" },
    "rincon ramirez giovanni jose": { empresa: "Top Solution Metal SAC", contrato: "Planilla" },
    "rubina reynoso nilton aquiles": { empresa: "Top Solution Metal SAC", contrato: "Planilla" },
    "saldana alanya carlos alberto": { empresa: "Top Solution Metal SAC", contrato: "Planilla" },
    "valderrama alva sergio angel": { empresa: "Top Solution Metal SAC", contrato: "Planilla" },
    "valderrama mozombite gustavo": { empresa: "Top Solution Metal SAC", contrato: "Planilla" },
    "cabrera carlos david": { empresa: "Top Solution Metal SAC", contrato: "Recibos por Honorarios" },
    "capcha rojas genaro vicente": { empresa: "Top Solution Metal SAC", contrato: "Recibos por Honorarios" },
    "cordova mori adriel": { empresa: "Top Solution Metal SAC", contrato: "Recibos por Honorarios" },
    "cotos alba alex luis": { empresa: "Top Solution Metal SAC", contrato: "Recibos por Honorarios" },
    "culqui aspajo pedro augusto": { empresa: "Top Solution Metal SAC", contrato: "Recibos por Honorarios" },
    "curitima murayari jorge miguel": { empresa: "Top Solution Metal SAC", contrato: "Recibos por Honorarios" },
    "gonzales chavez jesus": { empresa: "Top Solution Metal SAC", contrato: "Recibos por Honorarios" },
    "mendoza chauran luis enrique": { empresa: "Top Solution Metal SAC", contrato: "Recibos por Honorarios" },
    "montes bullon christopher": { empresa: "Top Solution Metal SAC", contrato: "Recibos por Honorarios" },
    "prada del pino jose luis": { empresa: "Top Solution Metal SAC", contrato: "Recibos por Honorarios" },
    "roncal quinto luis alberto": { empresa: "Top Solution Metal SAC", contrato: "Recibos por Honorarios" },
    "tamaris angeles jose antonio": { empresa: "Top Solution Metal SAC", contrato: "Recibos por Honorarios" },
    "andrade riobueno marco antonio": { empresa: "Top Solution Metal Group", contrato: "Planilla" },
    "bonifacio rodriguez yanfranco": { empresa: "Top Solution Metal Group", contrato: "Planilla" },
    "galindo cabezas joel serafin": { empresa: "Top Solution Metal Group", contrato: "Planilla" },
    "garcia cornelio roberto carlos": { empresa: "Top Solution Metal Group", contrato: "Planilla" },
    "gonzales llumpo cristian esteban": { empresa: "Top Solution Metal Group", contrato: "Planilla" },
    "lazaro salas jhon javier": { empresa: "Top Solution Metal Group", contrato: "Planilla" },
    "mejias perez jose alberto": { empresa: "Top Solution Metal Group", contrato: "Planilla" },
    "pezo fatama tono": { empresa: "Top Solution Metal Group", contrato: "Planilla" },
    "rivas sotomayor yorman rafael": { empresa: "Top Solution Metal Group", contrato: "Planilla" },
    "sequeiros prudencio jose luis": { empresa: "Top Solution Metal Group", contrato: "Planilla" },
    "suarez torrelles jehison omar": { empresa: "Top Solution Metal Group", contrato: "Planilla" },
    "aguilar mesia jose cirilo": { empresa: "Top Solution Metal Group", contrato: "Recibos por Honorarios" },
    "chahua villa efrain": { empresa: "Top Solution Metal Group", contrato: "Recibos por Honorarios" },
    "cipiran ramirez abraham moises": { empresa: "Top Solution Metal Group", contrato: "Recibos por Honorarios" },
    "collazos marquez eliot karl": { empresa: "Top Solution Metal Group", contrato: "Recibos por Honorarios" },
    "curitima murayari neuber": { empresa: "Top Solution Metal Group", contrato: "Recibos por Honorarios" },
    "delgado aguinaga anderson eduardo": { empresa: "Top Solution Metal Group", contrato: "Recibos por Honorarios" },
    "elescano ortiz leandro federico": { empresa: "Top Solution Metal Group", contrato: "Recibos por Honorarios" },
    "elescano ortiz leonardo antonio": { empresa: "Top Solution Metal Group", contrato: "Recibos por Honorarios" },
    "flores tuanama geiner": { empresa: "Top Solution Metal Group", contrato: "Recibos por Honorarios" },
    "lopez quispe karlo francisco": { empresa: "Top Solution Metal Group", contrato: "Recibos por Honorarios" },
    "pina gordon daniel andres": { empresa: "Top Solution Metal Group", contrato: "Recibos por Honorarios" },
    "ramirez arismendi manuel oscar": { empresa: "Top Solution Metal Group", contrato: "Recibos por Honorarios" },
    "urquia guevara orlando": { empresa: "Top Solution Metal Group", contrato: "Recibos por Honorarios" },
    "zuta vilca edinson": { empresa: "Top Solution Metal Group", contrato: "Recibos por Honorarios" }
  };

  const HORARIO_GENERAL = { 
    inH: 7, inM: 0, outH: 16, outM: 0, horasJornal: 8, diasLaborables: [1, 2, 3, 4, 5, 6] 
  };

  const SETTINGS = { dedupeMinutes: 3, weekdayLunchHours: 1, topeHsDiario: 6 };

  const state = {
    workbook: null, rawRows: [], cleanRows: [], summaryRows: [], anomalyRows: [],
    workers: [], tareo: { dates: [], rows: [] }, fileName: 'tareo_tsm.xlsx', map: null
  };

  const els = {
    fileInput: $('#fileInput'), fileName: $('#fileName'), dropzone: $('#dropzone'),
    themeBtn: $('#themeToggleBtn'), errorBox: $('#errorBox'), statusBadge: $('#statusBadge'),
    startDate: $('#startDate'), endDate: $('#endDate'), processBtn: $('#processBtn'),
    downloadXlsxBtn: $('#downloadXlsxBtn'), downloadCsvBtn: $('#downloadCsvBtn'),
    sampleBtn: $('#sampleBtn'), statRows: $('#statRows'), statWorkers: $('#statWorkers'),
    statDays: $('#statDays'), statSummary: $('#statSummary'), filterBar: $('#filterBar'),
    filterEmpresa: $('#filterEmpresa'), filterContrato: $('#filterContrato'),
    clearFiltersBtn: $('#clearFiltersBtn'), resumen: $('#tab-resumen'),
    tareo: $('#tab-tareo'), bd: $('#tab-bd'), anomalias: $('#tab-anomalias')
  };

  const DAYS = ['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'];
  const HOLIDAY_OVERRIDES = new Set(['02/04/2026','03/04/2026']);

  const pad = n => String(n).padStart(2, '0');
  const normalize = v => String(v ?? '').trim().toLowerCase();
  
  const normalizeSimple = v => normalize(v).normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9 ]+/g, ' ').replace(/\s+/g, ' ').trim();

  const fmtDate = d => `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()}`;
  const fmtInput = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const fmtTime = d => `${pad(d.getHours())}:${pad(d.getMinutes())}`;
  const weekdayName = d => DAYS[d.getDay()];
  const dayOnly = d => new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const stamp = d => dayOnly(d).getTime();
  const isSunday = d => d.getDay() === 0;
  const roundDownHalf = n => Math.floor((n || 0) * 2) / 2;
  const hoursBetween = (a, b) => Math.max(0, (b - a) / 36e5);
  const diffMinutes = (a, b) => Math.abs((b - a) / 60000);

  function displayNum(v) {
    if (v == null || v === '' || v === 0) return v === 0 ? '0' : '';
    const n = Number(v);
    return isFinite(n) ? n.toFixed(1).replace(/\.0$/, '') : String(v);
  }

  function escapeHtml(v) {
    return String(v ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  function isHoliday(d) { return HOLIDAY_OVERRIDES.has(fmtDate(d)); }
  function showError(msg) { els.errorBox.textContent = msg; els.errorBox.classList.remove('hidden'); }
  function clearError() { els.errorBox.classList.add('hidden'); els.errorBox.textContent = ''; }
  function setStatus(text, tone = '') { els.statusBadge.className = 'pill ' + tone; els.statusBadge.textContent = text; }
  function setFileName(text) { els.fileName.textContent = text || 'Ningún archivo seleccionado'; }

  function getEventType(estado) {
    const k = normalizeSimple(estado);
    if (k.includes('entrada') || k.includes('ingreso') || k.includes('check in') || k.includes('checkin') || k === 'in') return 'IN';
    if (k.includes('salida') || k.includes('egreso') || k.includes('check out') || k.includes('checkout') || k === 'out') return 'OUT';
    return '';
  }

  function parseExcelDate(value) {
    if (value instanceof Date && !isNaN(value)) return value;
    if (typeof value === 'number') {
      const p = XLSX.SSF.parse_date_code(value);
      if (p) return new Date(p.y, p.m - 1, p.d, p.H || 0, p.M || 0, p.S || 0);
    }
    const t = String(value ?? '').trim();
    if (!t) return null;
    const isoLike = t.match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
    if (isoLike) return new Date(Number(isoLike[1]), Number(isoLike[2]) - 1, Number(isoLike[3]), Number(isoLike[4] || 0), Number(isoLike[5] || 0), Number(isoLike[6] || 0));
    const dmy = t.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
    if (dmy) {
      const y = dmy[3].length === 2 ? Number('20' + dmy[3]) : Number(dmy[3]);
      return new Date(y, Number(dmy[2]) - 1, Number(dmy[1]), Number(dmy[4] || 0), Number(dmy[5] || 0), Number(dmy[6] || 0));
    }
    const direct = new Date(t);
    return isNaN(direct) ? null : direct;
  }

  function detectColumns(headers) {
    const found = { idPersona: '', nombre: '', hora: '', estado: '' };
    headers.forEach(h => {
      const k = normalizeSimple(h);
      if (!found.idPersona && ['id de persona','id persona','codigo de persona','codigo persona','id trabajador','id empleado','legajo','codigo','id'].includes(k)) found.idPersona = h;
      if (!found.nombre && ['nombre','colaborador','trabajador','empleado','personal','name','nombre completo'].includes(k)) found.nombre = h;
      if (!found.hora && ['hora','fecha y hora','fecha hora','datetime','marcacion','marcacion completa','fecha','time'].includes(k)) found.hora = h;
      if (!found.estado && ['estado de asistencia','estado','tipo','movimiento','attendance status'].includes(k)) found.estado = h;
    });
    return found;
  }

  function findWorkerInMaestro(excelName) {
    const normExcel = normalizeSimple(excelName);
    if (MAESTRO_PERSONAL[normExcel]) return MAESTRO_PERSONAL[normExcel];
    
    const excelWords = normExcel.split(' ');
    let bestMatch = null;
    let bestScore = 0;
    
    for (const dictKey in MAESTRO_PERSONAL) {
      const dictWords = dictKey.split(' ');
      let matches = 0;
      for (const w of excelWords) {
        if (dictWords.includes(w)) matches++;
      }
      if (matches > bestScore) {
        bestScore = matches;
        bestMatch = MAESTRO_PERSONAL[dictKey];
      }
    }
    
    if (bestScore >= 2) {
      return bestMatch;
    }
    return null;
  }

  function getWorkerKey(row) {
    return String(row.idPersona || '').trim() || normalizeSimple(row.nombre);
  }

  function buildRows(rows, map) {
    if (!map.nombre || !map.hora || !map.estado) {
      throw new Error('No se detectaron bien las columnas clave. Revisa tu Excel.');
    }
    
    return rows.map((row, i) => {
      const dt = parseExcelDate(row[map.hora]);
      const nombre = String(row[map.nombre] ?? '').trim();
      const estado = String(row[map.estado] ?? '').trim();
      const idPersona = map.idPersona ? String(row[map.idPersona] ?? '').trim() : '';
      const tipo = getEventType(estado);

      if (!dt || !nombre || !estado) return null;

      const infoPersonal = findWorkerInMaestro(nombre);
      if (!infoPersonal) return null;

      const workerKey = String(idPersona || '').trim() || normalizeSimple(nombre);

      return { 
        workerKey, idPersona, nombre, hora: dt, fecha: dayOnly(dt), estado, tipo, 
        empresa: infoPersonal.empresa, contrato: infoPersonal.contrato, fila: i + 2 
      };
    }).filter(Boolean).sort((a, b) => a.hora - b.hora || a.fila - b.fila);
  }

  function buildWorkerCatalog(rows) {
    const catalog = new Map();
    [...rows].sort((a, b) => a.fila - b.fila).forEach(r => {
      const key = getWorkerKey(r);
      if (!catalog.has(key)) {
        catalog.set(key, { 
          workerKey: key, idPersona: r.idPersona, nombre: r.nombre, 
          empresa: r.empresa, contrato: r.contrato, order: r.fila 
        });
      }
    });
    return [...catalog.values()].sort((a, b) => {
      if (a.empresa !== b.empresa) return a.empresa.localeCompare(b.empresa);
      if (a.contrato !== b.contrato) return a.contrato.localeCompare(b.contrato);
      return a.order - b.order;
    });
  }

  function analyzeDayMarks(records, fecha) {
    const anomalies = [];
    const typed = [...records].sort((a, b) => a.hora - b.hora || a.fila - b.fila).filter(r => {
      if (r.tipo) return true;
      anomalies.push(`Estado no reconocido: ${r.estado} ${fmtTime(r.hora)}`);
      return false;
    });

    const cleaned = [];
    for (const r of typed) {
      const prev = cleaned[cleaned.length - 1];
      if (prev && prev.tipo === r.tipo && diffMinutes(prev.hora, r.hora) <= SETTINGS.dedupeMinutes) {
        anomalies.push(`${r.tipo === 'IN' ? 'Entrada' : 'Salida'} duplicada ${fmtTime(r.hora)}`);
        continue;
      }
      cleaned.push(r);
    }

    const sessions = [];
    let openEntry = null;

    for (const r of cleaned) {
      if (r.tipo === 'IN') {
        if (!openEntry) openEntry = r;
        else anomalies.push(`Entrada repetida ${fmtTime(r.hora)}`);
      }
      if (r.tipo === 'OUT') {
        if (!openEntry) {
          anomalies.push(`Salida sin entrada ${fmtTime(r.hora)}`);
          continue;
        }
        if (r.hora <= openEntry.hora) {
          anomalies.push(`Salida antes de entrada ${fmtTime(r.hora)}`);
          continue;
        }
        sessions.push({ entrada: openEntry.hora, salida: r.hora, horas: hoursBetween(openEntry.hora, r.hora) });
        openEntry = null;
      }
    }

    if (openEntry) anomalies.push(`Entrada sin salida ${fmtTime(openEntry.hora)}`);

    const totalBruto = sessions.reduce((acc, s) => acc + s.horas, 0);
    return {
      rawCount: records.length, cleanedCount: cleaned.length, sessions, anomalies,
      primeraEntrada: sessions[0]?.entrada || cleaned.find(x => x.tipo === 'IN')?.hora || null,
      ultimaSalida: sessions[sessions.length - 1]?.salida || [...cleaned].reverse().find(x => x.tipo === 'OUT')?.hora || null,
      totalBruto: sessions.length ? totalBruto : null, totalNeto: sessions.length ? totalBruto : null
    };
  }

  function calcularHorasConTolerancia(minutosTotales) {
    const horasEnteras = Math.floor(minutosTotales / 60);
    const minutosSueltos = minutosTotales % 60;
    let hsToleradas = horasEnteras;
    if (minutosSueltos >= 51) hsToleradas += 1.0; 
    else if (minutosSueltos >= 45) hsToleradas += 0.5;
    return hsToleradas;
  }

  function calcWorkedHoursFromAnalysis(analysis, fecha) {
    if (!analysis.sessions.length) {
      return { brutas: null, netas: null, decimal: null, jornal: 0, hs: 0, dominical: 0 };
    }

    const horario = HORARIO_GENERAL;
    const isDomingoFeriado = isSunday(fecha) || isHoliday(fecha);

    if (isDomingoFeriado) {
      const netMinutes = Math.round(analysis.totalNeto * 60);
      let hsDomingo = calcularHorasConTolerancia(netMinutes);
      return { brutas: analysis.totalBruto, netas: analysis.totalNeto, decimal: analysis.totalNeto, jornal: 0, hs: 0, dominical: hsDomingo };
    }

    const horaEntradaOficial = new Date(fecha); horaEntradaOficial.setHours(horario.inH, horario.inM, 0, 0);
    const horaSalidaOficial = new Date(fecha); horaSalidaOficial.setHours(horario.outH, horario.outM, 0, 0);

    let minInWindow = 0; 
    let minOutWindow = 0; 

    analysis.sessions.forEach(s => {
      let inStart = Math.max(s.entrada.getTime(), horaEntradaOficial.getTime());
      let inEnd = Math.min(s.salida.getTime(), horaSalidaOficial.getTime());
      if (inEnd > inStart) minInWindow += (inEnd - inStart) / 60000;

      let outStart = Math.max(s.entrada.getTime(), horaSalidaOficial.getTime());
      let outEnd = s.salida.getTime();
      if (outEnd > outStart) minOutWindow += (outEnd - outStart) / 60000;
    });

    if (minInWindow >= 240) {
      minInWindow -= 60;
      if (minInWindow < 0) minInWindow = 0;
    }

    const minutosObjetivoJornal = horario.horasJornal * 60; 
    let jornalFinal = 0;
    
    if (minInWindow >= minutosObjetivoJornal - 15) {
      jornalFinal = horario.horasJornal;
    } else {
      jornalFinal = roundDownHalf(minInWindow / 60);
    }

    let hsFinal = 0;
    if (minOutWindow > 0) {
      hsFinal = calcularHorasConTolerancia(minOutWindow);
      if (hsFinal > SETTINGS.topeHsDiario) hsFinal = SETTINGS.topeHsDiario;
    }

    return { brutas: analysis.totalBruto, netas: analysis.totalNeto, decimal: analysis.totalNeto, jornal: jornalFinal, hs: hsFinal, dominical: 0 };
  }

  function buildSummary(cleanRows, startDate, endDate) {
    const effectiveRows = cleanRows.filter(r => stamp(r.fecha) >= stamp(startDate) && stamp(r.fecha) <= stamp(endDate));
    const group = new Map();

    effectiveRows.forEach(r => {
      const workerKey = getWorkerKey(r);
      const key = `${workerKey}__${fmtDate(r.fecha)}`;
      if (!group.has(key)) {
        group.set(key, { 
          workerKey, idPersona: r.idPersona, nombre: r.nombre, fecha: r.fecha, 
          empresa: r.empresa, contrato: r.contrato,
          records: [], order: r.fila 
        });
      }
      const g = group.get(key);
      g.records.push(r);
      if (r.fila < g.order) g.order = r.fila;
    });

    const workers = buildWorkerCatalog(effectiveRows);

    const summaryRows = [...group.values()].map(g => {
      const analysis = analyzeDayMarks(g.records, g.fecha);
      const calc = calcWorkedHoursFromAnalysis(analysis, g.fecha);

      let alerta = '';
      const sinMarcas = analysis.rawCount === 0;
      const marcaIncompleta = analysis.rawCount > 0 && analysis.sessions.length === 0;
      const isLaborable = HORARIO_GENERAL.diasLaborables.includes(g.fecha.getDay());

      if (sinMarcas && isLaborable && !isHoliday(g.fecha)) { alerta = 'F'; } 
      else if (marcaIncompleta) { alerta = 'I'; } 
      else if (analysis.anomalies.length) { alerta = 'REV'; }

      return {
        workerKey: g.workerKey, idPersona: g.idPersona, nombre: g.nombre, fecha: g.fecha, 
        empresa: g.empresa, contrato: g.contrato,
        primeraEntrada: analysis.primeraEntrada, ultimaSalida: analysis.ultimaSalida,
        horasBrutas: calc.brutas, horasNetas: calc.netas, horasDecimal: calc.decimal,
        jornal: calc.jornal, hs: calc.hs, dominical: calc.dominical,
        anomalias: analysis.anomalies.length, detalle: analysis.anomalies.join(' | '),
        order: g.order, feriado: isHoliday(g.fecha), alerta
      };
    }).sort((a, b) => {
      if (a.empresa !== b.empresa) return a.empresa.localeCompare(b.empresa);
      if (a.order !== b.order) return a.order - b.order;
      return a.fecha - b.fecha;
    });

    return { summaryRows, effectiveRows, workers };
  }

  function buildTareo(summaryRows, workers, startDate, endDate) {
    const dates = [];
    for (let d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 1)) { dates.push(new Date(d)); }
    const idx = new Map(summaryRows.map(r => [`${r.workerKey}__${fmtDate(r.fecha)}`, r]));

    const rows = workers.map(w => {
      const out = { 
        workerKey: w.workerKey, idPersona: w.idPersona, nombre: w.nombre, 
        empresa: w.empresa, contrato: w.contrato,
        dias: {}, hNormal: 0, hExtras: 0, dominical: 0 
      };

      dates.forEach(d => {
        const rec = idx.get(`${w.workerKey}__${fmtDate(d)}`);
        const feriado = isHoliday(d);
        const isLaborable = HORARIO_GENERAL.diasLaborables.includes(d.getDay());
        
        const sinMarcaDiaLaborable = !rec && isLaborable && !feriado;
        let alerta = rec?.alerta || (sinMarcaDiaLaborable ? 'F' : '');
        const jornal = rec?.jornal ?? 0;
        const hs = rec?.hs ?? 0;
        const dominical = rec?.dominical ?? 0;

        out.dias[fmtDate(d)] = { jornal, hs, dominical, feriado, alerta, detalle: rec?.detalle || '' };
        out.hNormal += jornal; out.hExtras += hs; out.dominical += dominical;
      });
      return out;
    });
    return { dates, rows };
  }

  function getFilteredData() {
    const emp = els.filterEmpresa.value;
    const con = els.filterContrato.value;

    const fWorkers = state.workers.filter(w => 
      (emp === 'ALL' || w.empresa === emp) &&
      (con === 'ALL' || w.contrato === con)
    );
    
    const validKeys = new Set(fWorkers.map(w => w.workerKey));

    return {
      summaryRows: state.summaryRows.filter(r => validKeys.has(r.workerKey)),
      cleanRows: state.cleanRows.filter(r => validKeys.has(r.workerKey)),
      anomalyRows: state.anomalyRows.filter(r => validKeys.has(r.workerKey)),
      tareoRows: state.tareo.rows.filter(r => validKeys.has(r.workerKey)),
      workers: fWorkers
    };
  }

  function populateFilters() {
    const empresas = [...new Set(state.workers.map(w => w.empresa))].sort();
    const contratos = [...new Set(state.workers.map(w => w.contrato))].sort();

    els.filterEmpresa.innerHTML = `<option value="ALL">Todas las Empresas</option>` + empresas.map(v => `<option value="${v}">${v}</option>`).join('');
    els.filterContrato.innerHTML = `<option value="ALL">Tipo de Contratación</option>` + contratos.map(v => `<option value="${v}">${v}</option>`).join('');
    
    els.filterBar.style.display = 'flex';
  }

  function handleRowsLoaded(rows) {
    state.rawRows = rows;
    state.map = detectColumns(Object.keys(rows[0] || {}));
    state.cleanRows = buildRows(rows, state.map);

    if (!state.cleanRows.length) throw new Error('No se encontraron coincidencias entre el Excel y la lista de obreros configurada.');

    const ds = state.cleanRows.map(r => stamp(r.fecha));
    els.startDate.value = fmtInput(new Date(Math.min(...ds)));
    els.endDate.value = fmtInput(new Date(Math.max(...ds)));

    processNow();
  }

  function renderTable(container, headers, rows, options = {}) {
    const nameColIndex = Number.isInteger(options.nameColIndex) ? options.nameColIndex : 0;
    container.innerHTML =
      '<table><thead><tr>' +
      headers.map((h, i) => `<th class="${i === 0 ? 'name-col' : ''}">${escapeHtml(h)}</th>`).join('') +
      '</tr></thead><tbody>' +
      rows.map(r =>
        '<tr>' +
        r.map((v, i) => {
          const value = v == null ? '' : v;
          const isNum = typeof value === 'number';
          let className = [ i === 0 ? 'name-col' : '', isNum ? 'num' : '' ];
          
          if (i === nameColIndex) {
            const contratoStr = String(r[options.contratoStrIndex] ?? '').toUpperCase();
            className.push(contratoStr.includes('PLANILLA') ? 'planilla-name' : 'rrhh-name');
          }
          
          if (options.renderTags && i === options.tagIndex && value !== '') {
             return `<td class="${className.join(' ').trim()}">${value}</td>`;
          }

          return `<td class="${className.join(' ').trim()}">${escapeHtml(isNum ? displayNum(value) : value)}</td>`;
        }).join('') +
        '</tr>'
      ).join('') +
      '</tbody></table>';
  }

  function renderAll() {
    const fd = getFilteredData();

    els.statWorkers.textContent = fd.workers.length;
    els.statSummary.textContent = fd.summaryRows.length;

    renderTable(els.resumen,
      ['Nombre','ID','Fecha','Día','Asignación','Primera entrada','Última salida','Jornal','H.S.','Alerta','Detalle'],
      fd.summaryRows.slice(0, 2000).map(r => [
        r.nombre, r.idPersona || '', fmtDate(r.fecha), weekdayName(r.fecha),
        `<span class="meta-tag">${r.empresa}</span><span class="meta-tag">${r.contrato}</span>`,
        r.primeraEntrada ? fmtTime(r.primeraEntrada) : '', r.ultimaSalida ? fmtTime(r.ultimaSalida) : '',
        displayNum(r.jornal), displayNum(r.hs), r.alerta, r.detalle
      ]),
      { nameColIndex: 0, contratoStrIndex: 4, renderTags: true, tagIndex: 4 }
    );

    renderTable(els.bd,
      ['Nombre','ID','Fecha','Día','Hora','Estado','Tipo','Empresa','Contrato'],
      fd.cleanRows.slice(0, 3000).map(r => [
        r.nombre, r.idPersona || '', fmtDate(r.fecha), weekdayName(r.fecha), fmtTime(r.hora), r.estado, r.tipo || '', r.empresa, r.contrato
      ]),
      { nameColIndex: 0, contratoStrIndex: 8 }
    );

    renderTable(els.anomalias,
      ['Nombre','ID','Fecha','Día','Asignación','Alerta','Detalle','Jornal','H.S.'],
      fd.anomalyRows.slice(0, 3000).map(r => [
        r.nombre, r.idPersona || '', fmtDate(r.fecha), weekdayName(r.fecha), `${r.empresa} - ${r.contrato}`, r.alerta, r.detalle, displayNum(r.jornal), displayNum(r.hs)
      ]),
      { nameColIndex: 0, contratoStrIndex: 4 }
    );

    let html = '<table><thead><tr>';
    html += '<th class="name-col" rowspan="2">Colaboradores</th>';
    html += state.tareo.dates.map(d => `<th colspan="2" class="day-head day-start">${escapeHtml(fmtDate(d))}<span class="day-name">${escapeHtml(weekdayName(d))}</span></th>`).join('');
    html += '<th rowspan="2" class="totals-head totals-start">H.NORMAL</th><th rowspan="2" class="totals-head">H.EXTRAS</th><th rowspan="2" class="totals-head">DOMINICAL</th></tr><tr>';
    html += state.tareo.dates.map((d, idx) => `<th class="sub-day-start ${idx === 0 ? 'day-start' : ''}">JORNAL</th><th>H.S.</th>`).join('');
    html += '</tr></thead><tbody>';

    fd.tareoRows.forEach(r => {
      const isPlanilla = r.contrato.toUpperCase().includes('PLANILLA');
      const nameClass = isPlanilla ? 'planilla-name' : 'rrhh-name';
      html += `<tr><td class="name-col ${nameClass}"><div>${escapeHtml(r.nombre)}</div><div style="font-size:0.7rem;color:var(--muted);font-weight:600;">${r.empresa} / ${r.contrato}</div></td>`;
      
      state.tareo.dates.forEach((d, idx) => {
        const k = fmtDate(d);
        const v = r.dias[k] || { jornal: 0, hs: 0, feriado: false, alerta: '', detalle: '' };
        
        let journalText = displayNum(v.jornal);
        if (v.alerta === 'F') journalText = 'F';
        if (v.alerta === 'I') journalText = 'I';

        const hsText = (v.alerta === 'F' || v.alerta === 'I' || v.hs === 0) ? '' : displayNum(v.hs);
        
        let cellClass = '';
        if (v.alerta === 'F') cellClass = 'alerta-f';
        else if (v.alerta === 'I') cellClass = 'alerta-i';
        else if (v.alerta) cellClass = 'alert-cell';

        const title = escapeHtml(v.detalle || (v.alerta === 'F' ? 'Sin marcaciones (Falta)' : (v.alerta === 'I' ? 'Marcación incompleta' : '')));
        html += `<td class="num ${idx === 0 ? 'day-start sub-day-start' : ''} ${cellClass}" title="${title}">${escapeHtml(journalText)}</td>`;
        html += `<td class="num ${cellClass}" title="${title}">${escapeHtml(hsText)}</td>`;
      });
      html += `<td class="num totals-cell totals-start">${escapeHtml(displayNum(r.hNormal))}</td>`;
      html += `<td class="num totals-cell">${escapeHtml(displayNum(r.hExtras))}</td>`;
      html += `<td class="num totals-cell">${escapeHtml(displayNum(r.dominical))}</td></tr>`;
    });

    html += '</tbody></table>';
    els.tareo.innerHTML = html;

    requestAnimationFrame(() => {
      const tareoThead = els.tareo.querySelector('thead tr:first-child');
      if (tareoThead) {
        const offset = tareoThead.offsetHeight;
        els.tareo.querySelectorAll('thead tr:nth-child(2) th').forEach(th => {
          th.style.top = (offset - 1) + 'px'; 
          th.style.zIndex = '9'; 
        });
      }
    });
  }

  function updateStats(sourceRows) {
    els.statRows.textContent = sourceRows.length;
    els.statDays.textContent = state.tareo.dates.length;
  }

  function buildTareoSheetData(filteredTareoRows) {
    const row1 = ['Colaboradores'];
    const row2 = [''];
    state.tareo.dates.forEach(d => { row1.push(`${fmtDate(d)} ${weekdayName(d)}`, ''); row2.push('JORNAL', 'H.S.'); });
    row1.push('H.NORMAL', 'H.EXTRAS', 'DOMINICAL');
    row2.push('', '', '');

    const data = [row1, row2];
    filteredTareoRows.forEach(r => {
      const row = [`${r.nombre} [${r.empresa} - ${r.contrato}]`];
      state.tareo.dates.forEach(d => {
        const k = fmtDate(d);
        const cell = r.dias[k] || { jornal: 0, hs: 0, alerta: '' };
        if (cell.alerta === 'F') { row.push('F', ''); }
        else if (cell.alerta === 'I') { row.push('I', ''); }
        else {
          row.push(cell.jornal ?? 0);
          row.push(cell.hs ?? 0);
        }
      });
      row.push(r.hNormal, r.hExtras, r.dominical);
      data.push(row);
    });
    return data;
  }

  function workbookOut() {
    const fd = getFilteredData(); 
    const wb = XLSX.utils.book_new();
    const tareoData = buildTareoSheetData(fd.tareoRows);
    const wsTareo = XLSX.utils.aoa_to_sheet(tareoData);
    wsTareo['!merges'] = [];

    let c = 1;
    state.tareo.dates.forEach(() => { wsTareo['!merges'].push({ s: { r: 0, c }, e: { r: 0, c: c + 1 } }); c += 2; });
    wsTareo['!merges'].push({ s: { r: 0, c: 0 }, e: { r: 1, c: 0 } });

    const start = 1 + state.tareo.dates.length * 2;
    wsTareo['!merges'].push({ s: { r: 0, c: start }, e: { r: 1, c: start } });
    wsTareo['!merges'].push({ s: { r: 0, c: start + 1 }, e: { r: 1, c: start + 1 } });
    wsTareo['!merges'].push({ s: { r: 0, c: start + 2 }, e: { r: 1, c: start + 2 } });

    for (let r = 2; r < tareoData.length; r++) {
      for (let col = 1; col < tareoData[r].length; col++) {
        const ref = XLSX.utils.encode_cell({ r, c: col });
        if (wsTareo[ref] && typeof wsTareo[ref].v === 'number') wsTareo[ref].z = '0.##';
      }
    }

    wsTareo['!cols'] = [{ wch: 45 }, ...Array(tareoData[0].length - 1).fill({ wch: 10 })];
    XLSX.utils.book_append_sheet(wb, wsTareo, 'TAREO');

    const resumenData = [
      ['Empresa','Contrato','Nombre','ID','Fecha','Día','Primera entrada','Última salida','Jornal','H.S.','Alerta','Detalle'],
      ...fd.summaryRows.map(r => [
        r.empresa, r.contrato, r.nombre, r.idPersona || '', fmtDate(r.fecha), weekdayName(r.fecha),
        r.primeraEntrada ? fmtTime(r.primeraEntrada) : '', r.ultimaSalida ? fmtTime(r.ultimaSalida) : '',
        r.jornal ?? '', r.hs ?? '', r.alerta, r.detalle
      ])
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(resumenData), 'RESUMEN');

    const anomaliasData = [
      ['Empresa','Contrato','Nombre','ID','Fecha','Día','Alerta','Detalle','Jornal','H.S.'],
      ...fd.anomalyRows.map(r => [
        r.empresa, r.contrato, r.nombre, r.idPersona || '', fmtDate(r.fecha), weekdayName(r.fecha), r.alerta, r.detalle, r.jornal ?? '', r.hs ?? ''
      ])
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(anomaliasData), 'ANOMALIAS');

    return wb;
  }

  function processNow() {
    clearError();
    try {
      if (!state.rawRows.length) throw new Error('Primero carga un archivo o un ejemplo.');
      const startDate = els.startDate.value ? new Date(els.startDate.value + 'T00:00:00') : new Date();
      const endDate = els.endDate.value ? new Date(els.endDate.value + 'T00:00:00') : new Date();
      if (startDate > endDate) throw new Error('La fecha de inicio no puede ser mayor que la fecha fin.');

      const { summaryRows, effectiveRows, workers } = buildSummary(state.cleanRows, startDate, endDate);
      state.summaryRows = summaryRows;
      state.anomalyRows = summaryRows.filter(r => r.alerta || r.anomalias > 0);
      state.workers = workers;
      state.tareo = buildTareo(summaryRows, workers, startDate, endDate);

      populateFilters();
      renderAll();
      updateStats(effectiveRows);

      setStatus(`Procesado correctamente`, 'ok');
      els.downloadXlsxBtn.disabled = !summaryRows.length;
      els.downloadCsvBtn.disabled = !summaryRows.length;
    } catch (err) {
      setStatus('Ocurrió un error', 'error');
      showError(err.message || 'No se pudo procesar el archivo.');
    }
  }

  function handleWorkbook(file) {
    clearError();
    setFileName(file?.name || 'Ningún archivo seleccionado');
    const reader = new FileReader();
    reader.onload = e => {
      try {
        state.workbook = XLSX.read(e.target.result, { type: 'array', cellDates: true, raw: true });
        state.fileName = (file.name || 'tareo').replace(/\.(xlsx|xls)$/i, '') + '_filtrado.xlsx';
        const name = state.workbook.SheetNames[0];
        const ws = state.workbook.Sheets[name];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: true, cellDates: true });
        if (!rows.length) throw new Error('La hoja seleccionada no tiene datos.');
        handleRowsLoaded(rows);
      } catch (err) {
        setStatus('No se pudo leer el Excel', 'error');
        showError(err.message || 'Revisa si el archivo está dañado o no es un Excel válido.');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function loadSample() {
    const rows = [
      { 'ID de persona': '101', Nombre: 'Jose Lamper Torres', Hora: '01/04/2026 06:50', 'Estado de asistencia': 'Registro de entrada' },
      { 'ID de persona': '101', Nombre: 'Jose Lamper Torres', Hora: '01/04/2026 16:00', 'Estado de asistencia': 'Registro de salida' },
      { 'ID de persona': '102', Nombre: 'Abraham Cipiran Ramirez', Hora: '01/04/2026 07:15', 'Estado de asistencia': 'Registro de entrada' },
      { 'ID de persona': '102', Nombre: 'Abraham Cipiran Ramirez', Hora: '01/04/2026 16:00', 'Estado de asistencia': 'Registro de salida' },
      { 'ID de persona': '999', Nombre: 'Persona Externa', Hora: '01/04/2026 08:00', 'Estado de asistencia': 'Registro de entrada' }
    ];
    state.fileName = 'tareo_tsm_test.xlsx';
    setFileName('Ejemplo Cargado');
    handleRowsLoaded(rows);
  }

  function downloadXlsx() {
    try { XLSX.writeFile(workbookOut(), state.fileName); } 
    catch (err) { showError('No se pudo generar el Excel de salida.'); }
  }

  function downloadCsv() {
    const fd = getFilteredData();
    const blob = new Blob([XLSX.utils.sheet_to_csv(XLSX.utils.aoa_to_sheet(buildTareoSheetData(fd.tareoRows)))], { type: 'text/csv;charset=utf-8;' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'tareo.csv';
    a.click();
    URL.revokeObjectURL(a.href);
  }

  els.filterEmpresa.addEventListener('change', renderAll);
  els.filterContrato.addEventListener('change', renderAll);
  
  els.clearFiltersBtn.addEventListener('click', () => {
    els.filterEmpresa.value = 'ALL';
    els.filterContrato.value = 'ALL';
    renderAll();
  });

  els.themeBtn.addEventListener('change', e => { 
    document.documentElement.setAttribute('data-theme', e.target.checked ? 'dark' : 'light'); 
  });
  els.fileInput.addEventListener('change', e => { const f = e.target.files?.[0]; if (f) handleWorkbook(f); });
  els.processBtn.addEventListener('click', processNow);
  els.downloadXlsxBtn.addEventListener('click', downloadXlsx);
  els.downloadCsvBtn.addEventListener('click', downloadCsv);
  els.sampleBtn.addEventListener('click', loadSample);

  ['dragenter','dragover'].forEach(ev => els.dropzone.addEventListener(ev, e => { e.preventDefault(); els.dropzone.classList.add('drag'); }));
  ['dragleave','drop'].forEach(ev => els.dropzone.addEventListener(ev, e => { e.preventDefault(); els.dropzone.classList.remove('drag'); }));
  els.dropzone.addEventListener('drop', e => { const f = e.dataTransfer.files?.[0]; if (f) handleWorkbook(f); });

  document.querySelectorAll('.tab').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab').forEach(x => x.classList.remove('active'));
      btn.classList.add('active');
      ['resumen','tareo','bd','anomalias'].forEach(k => document.getElementById('tab-' + k).classList.add('hidden'));
      document.getElementById('tab-' + btn.dataset.tab).classList.remove('hidden');
    });
  });
})();
