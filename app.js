/* =========================================================================
   CONCILIACIÓN BANCO/CLIENTES EN DÓLARES
   Lógica de cruce 1:1 con el proceso manual: ALTO / BAJO / SIN MATCH / REPETIDO
   Todo corre en el navegador. Ningún archivo sale de la máquina del usuario.
   ========================================================================= */

const COLORS = {
  ALTO:      { fill: 'FFC6EFCE' },
  MEDIO:     { fill: 'FFFFEB9C' },
  BAJO:      { fill: 'FFFFB6C1' },
  'SIN MATCH': { fill: 'FFFFC7CE' },
  REPETIDO:  { fill: 'FFBDD7EE' },
};
const PURPLE_ARGB = 'FF7030A0';

let state = {
  bancoFile: null,
  clientesFile: null,
  bancoWb: null,       // raw ArrayBuffer, kept for output generation
  clientesWb: null,
  bancoRows: [],       // parsed banco data
  cliRows: [],         // parsed clientes data
  bancoHeaderMap: {},  // column index by logical name
  cliHeaderMap: {},
  bancoDataStartRow: 2,
  cliDataStartRow: 2,
  ratesByDate: {},     // 'YYYY-MM-DD' -> number|null
  aplicado: false,     // true once user confirms rates via the Aplicar button
  matched: false,
};

/* ---------------- Helpers ---------------- */

function normCuit(val) {
  if (val === null || val === undefined) return '';
  return String(val).trim().replace(/[\s\-.]/g, '');
}

function normMonto(val) {
  if (val === null || val === undefined || val === '') return null;
  let s = String(val).toLowerCase();
  s = s.replace(/dolares|usd|u\$d|\$/g, '');
  s = s.replace(/[^0-9.,]/g, '').trim();
  s = s.replace(',', '.');
  const f = parseFloat(s);
  return isNaN(f) ? null : f;
}

function extractCuitFromConcepto(concepto) {
  if (!concepto) return '';
  const m = String(concepto).match(/\d{10,11}/);
  return m ? m[0] : '';
}

function toDateOnly(d) {
  if (!d) return null;
  const dt = (d instanceof Date) ? d : new Date(d);
  if (isNaN(dt.getTime())) return null;
  return new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
}

function dateKey(d) {
  if (!d) return '';
  const p = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

function dateLabel(d) {
  if (!d) return '';
  const p = n => String(n).padStart(2, '0');
  return `${p(d.getDate())}/${p(d.getMonth() + 1)}/${d.getFullYear()}`;
}

function findColumn(headerRow, candidates) {
  // headerRow: array of {text, col}
  for (const cell of headerRow) {
    const t = (cell.text || '').toLowerCase();
    for (const cand of candidates) {
      if (t.includes(cand.toLowerCase())) return cell.col;
    }
  }
  return null;
}

/* ---------------- File loading ---------------- */

async function loadWorkbook(file) {
  const buf = await file.arrayBuffer();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  return { wb, buf };
}

function getHeaderCells(ws) {
  const row = ws.getRow(1);
  const cells = [];
  row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    cells.push({ text: String(cell.value ?? ''), col: colNumber });
  });
  return cells;
}

async function parseClientes(file) {
  const { wb, buf } = await loadWorkbook(file);
  const ws = wb.worksheets[0];
  const headers = getHeaderCells(ws);

  const colCuit   = findColumn(headers, ['cuil/cuit', 'cuil', 'cuit']);
  const colMonto  = findColumn(headers, ['monto']);
  const colNro    = findColumn(headers, ['número de cliente', 'numero de cliente']);
  const colNombre = findColumn(headers, ['nombre y apellido del pasajero', 'nombre y apellido']);
  const colDni    = findColumn(headers, ['dni']);
  const colMarca  = findColumn(headers, ['marca temporal']);
  const colObs    = findColumn(headers, ['observaciones']);
  const colRefBco = findColumn(headers, ['ref bco', 'ref. bco', 'referencia bco']);

  const rows = [];
  const lastRow = ws.actualRowCount;
  for (let r = 2; r <= lastRow; r++) {
    const row = ws.getRow(r);
    const nro = colNro ? row.getCell(colNro).value : null;
    const nombre = colNombre ? row.getCell(colNombre).value : null;
    // Stop-detection: consider row empty if nro and nombre both blank
    if ((nro === null || nro === undefined || nro === '') &&
        (nombre === null || nombre === undefined || nombre === '')) {
      continue;
    }
    const cuitRaw = colCuit ? normCuit(row.getCell(colCuit).value) : '';
    const montoRaw = colMonto ? normMonto(row.getCell(colMonto).value) : null;
    const dni = colDni ? row.getCell(colDni).value : '';
    rows.push({
      excelRow: r,
      nro: nro,
      nombre: nombre,
      dni: dni,
      cuitRaw,
      monto: montoRaw,
      nivel: null,
      bidx: null,
      cuitOk: false,
      montoOk: false,
      fechaOk: false,
      bancoRef: null,
      bancoFecha: null,
    });
  }

  return {
    wb, buf, ws,
    headerMap: { colCuit, colMonto, colNro, colNombre, colDni, colMarca, colObs, colRefBco },
    rows,
  };
}

async function parseBanco(file) {
  const { wb, buf } = await loadWorkbook(file);
  const ws = wb.worksheets[0];
  const headers = getHeaderCells(ws);

  const colFecha  = findColumn(headers, ['fecha pago']);
  const colRef    = findColumn(headers, ['referencia']);
  const colConc   = findColumn(headers, ['concepto']);
  // "DÉBITO" appears possibly twice (DÉBITO and DÉBITO.1); take the first occurrence only
  let colDebito = null;
  for (const cell of headers) {
    if (cell.text.toLowerCase().startsWith('débito') || cell.text.toLowerCase().startsWith('debito')) {
      colDebito = cell.col;
      break;
    }
  }
  const colObs   = findColumn(headers, ['observaciones']);
  const colTipo  = findColumn(headers, ['tipo']);
  const colTC    = findColumn(headers, ['t.c.', 't.c', 'tc']);

  const rows = [];
  const lastRow = ws.actualRowCount;
  for (let r = 2; r <= lastRow; r++) {
    const row = ws.getRow(r);
    const fechaRaw = colFecha ? row.getCell(colFecha).value : null;
    const ref = colRef ? row.getCell(colRef).value : null;
    const concepto = colConc ? row.getCell(colConc).value : '';
    const debito = colDebito ? row.getCell(colDebito).value : null;
    if ((ref === null || ref === undefined || ref === '') &&
        (fechaRaw === null || fechaRaw === undefined)) {
      continue;
    }
    const fecha = toDateOnly(fechaRaw);
    rows.push({
      excelRow: r,
      fecha,
      referencia: ref,
      concepto,
      monto: normMonto(debito),
      cuit: extractCuitFromConcepto(concepto),
      used: false,
    });
  }

  return {
    wb, buf, ws,
    headerMap: { colFecha, colRef, colConc, colDebito, colObs, colTipo, colTC },
    rows,
  };
}

/* ---------------- Matching algorithm ---------------- */

function findMatch(cuitRaw, monto, bancoRows, usedSet) {
  const cc = cuitRaw;

  const byCuit = bancoRows.filter(b => b.cuit === cc && cc !== '');
  if (byCuit.length > 0) {
    const byCuitMonto = byCuit.filter(b => b.monto === monto);
    const availableMonto = byCuitMonto.filter(b => !usedSet.has(b.excelRow));
    if (availableMonto.length > 0) {
      return { nivel: 'ALTO', b: availableMonto[0], cuitOk: true, montoOk: true, fechaOk: true };
    }
    const availableCuit = byCuit.filter(b => !usedSet.has(b.excelRow));
    if (availableCuit.length > 0) {
      return { nivel: 'BAJO', b: availableCuit[0], cuitOk: true, montoOk: false, fechaOk: false };
    }
    if (byCuitMonto.length > 0) {
      return { nivel: 'ALTO', b: byCuitMonto[0], cuitOk: true, montoOk: true, fechaOk: true };
    }
    return { nivel: 'BAJO', b: byCuit[0], cuitOk: true, montoOk: false, fechaOk: false };
  }

  // Partial CUIT (7-10 digits): search last 8 digits inside CONCEPTO
  if (cc.length < 11 && cc.length >= 7) {
    const partial = cc.length >= 8 ? cc.slice(-8) : cc;
    for (const b of bancoRows) {
      if (usedSet.has(b.excelRow)) continue;
      if (String(b.concepto).includes(partial)) {
        const montoOk = b.monto === monto;
        return { nivel: 'BAJO', b, cuitOk: false, montoOk, fechaOk: montoOk };
      }
    }
  }

  // 10-digit CUIT missing one trailing digit: brute-force append 0-9
  if (cc.length === 10) {
    for (let s = 0; s <= 9; s++) {
      const candidate = cc + s;
      const found = bancoRows.filter(b => b.cuit === candidate);
      if (found.length > 0) {
        const avail = found.filter(b => !usedSet.has(b.excelRow));
        const availMonto = avail.filter(b => b.monto === monto);
        if (availMonto.length > 0) {
          return { nivel: 'BAJO', b: availMonto[0], cuitOk: false, montoOk: true, fechaOk: true };
        }
        if (avail.length > 0) {
          return { nivel: 'BAJO', b: avail[0], cuitOk: false, montoOk: false, fechaOk: false };
        }
      }
    }
  }

  // CUIT with 1-2 digit typo, same monto
  const sameMontoAvail = bancoRows.filter(b => b.monto === monto && !usedSet.has(b.excelRow));
  for (const b of sameMontoAvail) {
    const bc = b.cuit;
    if (bc.length === 11 && cc.length >= 9) {
      const ccPad = cc.padStart(11, '0');
      let diffs = 0;
      for (let i = 0; i < 11; i++) if (ccPad[i] !== bc[i]) diffs++;
      if (diffs <= 2) {
        return { nivel: 'BAJO', b, cuitOk: false, montoOk: true, fechaOk: true };
      }
    }
  }

  return { nivel: 'SIN MATCH', b: null, cuitOk: false, montoOk: false, fechaOk: false };
}

function runMatching() {
  const bancoRows = state.bancoRows;
  const cliRows = state.cliRows;

  // Pass 1: tentative match for every row (to compute REPETIDO keys)
  const used1 = new Set();
  const pre = cliRows.map(row => {
    const r = findMatch(row.cuitRaw, row.monto, bancoRows, used1);
    if (r.b) used1.add(r.b.excelRow);
    return r;
  });

  const seen = new Map();
  const isDup = new Array(cliRows.length).fill(false);
  pre.forEach((r, i) => {
    const row = cliRows[i];
    const ref = r.b ? r.b.referencia : null;
    const key = `${row.nro}|${row.cuitRaw}|${row.monto}|${ref}`;
    if (seen.has(key)) {
      isDup[i] = true;
    } else {
      seen.set(key, i);
    }
  });

  // Pass 2: final match, skipping duplicates in consumption order
  const used2 = new Set();
  cliRows.forEach((row, i) => {
    if (isDup[i]) {
      row.nivel = 'REPETIDO';
      row.bidx = null;
      row.cuitOk = row.montoOk = row.fechaOk = false;
      return;
    }
    const r = findMatch(row.cuitRaw, row.monto, bancoRows, used2);
    row.nivel = r.nivel;
    row.bidx = r.b;
    row.cuitOk = r.cuitOk;
    row.montoOk = r.montoOk;
    row.fechaOk = r.fechaOk;
    if (r.b) {
      used2.add(r.b.excelRow);
      row.bancoRef = r.b.referencia;
      row.bancoFecha = r.b.fecha;
    }
  });

  state.matched = true;
}

/* ---------------- Rates (cotización BNA) ---------------- */

async function fetchOfficialRate() {
  try {
    const res = await fetch('https://dolarapi.com/v1/dolares/oficial');
    if (!res.ok) throw new Error('bad response');
    const data = await res.json();
    if (data && typeof data.compra === 'number') return data.compra;
    return null;
  } catch (e) {
    return null;
  }
}

function collectBancoDates() {
  const dates = new Set();
  state.bancoRows.forEach(b => { if (b.fecha) dates.add(dateKey(b.fecha)); });
  const today = toDateOnly(new Date());
  const parsed = Array.from(dates).map(k => {
    const [y, m, d] = k.split('-').map(Number);
    return new Date(y, m - 1, d);
  });
  // Sort by proximity to today first (closest first); ties broken chronologically
  parsed.sort((a, b) => {
    const diffA = Math.abs(a.getTime() - today.getTime());
    const diffB = Math.abs(b.getTime() - today.getTime());
    if (diffA !== diffB) return diffA - diffB;
    return a.getTime() - b.getTime();
  });
  return parsed;
}

async function renderRatesStep() {
  const dates = collectBancoDates();
  const today = toDateOnly(new Date());
  const list = document.getElementById('rates-list');
  list.innerHTML = '';
  state.aplicado = false;
  let todayFetchFailed = false;

  for (const d of dates) {
    const key = dateKey(d);
    const isToday = key === dateKey(today);
    const row = document.createElement('div');
    row.className = 'rate-row';
    row.innerHTML = `
      <div class="date">${dateLabel(d)}${isToday ? ' (hoy)' : ''}</div>
      <div class="status pending" id="status-${key}">Pendiente</div>
      <div>
        <span class="rate-prefix">$</span>
        <input type="text" class="mono rate-input" id="rate-${key}" placeholder="Ej: 1460">
      </div>
    `;
    list.appendChild(row);

    const input = row.querySelector(`#rate-${key}`);
    input.addEventListener('input', () => {
      const val = parseFloat(input.value.replace(',', '.'));
      state.ratesByDate[key] = isNaN(val) ? null : val;
      state.aplicado = false;
      updateRateStatus(key);
      refreshAplicarAvailability();
      refreshDownloadAvailability();
    });

    if (isToday) {
      const val = await fetchOfficialRate();
      if (val) {
        input.value = val;
        state.ratesByDate[key] = val;
      } else {
        todayFetchFailed = true;
      }
    }
    updateRateStatus(key);
  }

  document.getElementById('btn-refetch').style.display = todayFetchFailed ? 'inline-block' : 'none';
  refreshAplicarAvailability();
  refreshDownloadAvailability();
}

function updateRateStatus(key) {
  const el = document.getElementById(`status-${key}`);
  if (!el) return;
  const has = state.ratesByDate[key] !== undefined && state.ratesByDate[key] !== null;
  el.textContent = has ? 'Cargada' : 'Pendiente';
  el.className = 'status ' + (has ? 'ok' : 'pending');
}

function allRatesFilled() {
  const dates = collectBancoDates().map(dateKey);
  return dates.length > 0 && dates.every(k => state.ratesByDate[k] !== undefined && state.ratesByDate[k] !== null);
}

function refreshAplicarAvailability() {
  const btn = document.getElementById('btn-aplicar');
  if (btn) btn.disabled = !allRatesFilled();
}

function refreshDownloadAvailability() {
  const btnCli = document.getElementById('btn-download-clientes');
  const btnBco = document.getElementById('btn-download-banco');
  const warn = document.getElementById('rates-missing-warn');
  const ok = allRatesFilled() && state.aplicado;
  if (btnCli) btnCli.disabled = !ok;
  if (btnBco) btnBco.disabled = !ok;
  warn.style.display = ok ? 'none' : 'block';
}

/* ---------------- Rendering results ---------------- */

function niveltoClass(n) {
  return n.replace(/\s+/g, '');
}

function renderSummary() {
  const counts = { ALTO: 0, MEDIO: 0, BAJO: 0, 'SIN MATCH': 0, REPETIDO: 0 };
  state.cliRows.forEach(r => counts[r.nivel] = (counts[r.nivel] || 0) + 1);
  const total = state.cliRows.length;

  const grid = document.getElementById('summary-grid');
  grid.innerHTML = '';
  const order = ['ALTO', 'BAJO', 'SIN MATCH', 'REPETIDO'];
  const labels = { ALTO: 'Alto', BAJO: 'Bajo', 'SIN MATCH': 'Sin match', REPETIDO: 'Repetido' };
  const classes = { ALTO: 'alto', BAJO: 'bajo', 'SIN MATCH': 'sinmatch', REPETIDO: 'repetido' };

  order.forEach(k => {
    const c = counts[k] || 0;
    if (c === 0 && k !== 'ALTO') return;
    const pct = total ? ((c / total) * 100).toFixed(1) : '0.0';
    const div = document.createElement('div');
    div.className = `stat ${classes[k]}`;
    div.innerHTML = `<div class="count">${c}</div><div class="lbl">${labels[k]}</div><div class="pct">${pct}%</div>`;
    grid.appendChild(div);
  });

  const totalDiv = document.createElement('div');
  totalDiv.className = 'stat total';
  totalDiv.innerHTML = `<div class="count">${total}</div><div class="lbl">Total</div><div class="pct">100%</div>`;
  grid.appendChild(totalDiv);
}

function renderTable() {
  const tbody = document.getElementById('results-tbody');
  tbody.innerHTML = '';
  state.cliRows.forEach((row, i) => {
    const tr = document.createElement('tr');
    const check = v => v ? '<span class="check-yes">✔</span>' : '<span class="check-no">✖</span>';
    tr.innerHTML = `
      <td>${i + 1}</td>
      <td>${row.nro ?? ''}</td>
      <td>${row.nombre ?? ''}</td>
      <td class="mono">${row.cuitRaw}</td>
      <td class="mono">${row.monto ?? ''}</td>
      <td><span class="badge ${niveltoClass(row.nivel)}">${row.nivel}</span></td>
      <td>${check(row.cuitOk)}</td>
      <td>${check(row.montoOk)}</td>
      <td>${check(row.fechaOk)}</td>
      <td class="mono">${row.bancoRef ?? '-'}</td>
      <td class="mono">${row.bancoFecha ? dateLabel(row.bancoFecha) : '-'}</td>
    `;
    tbody.appendChild(tr);
  });
}

/* ---------------- Excel generation ---------------- */

function argbFromString(hexNoAlpha) {
  return { argb: 'FF' + hexNoAlpha.replace('#', '') };
}

async function generateClientesOutput() {
  const src = state.clientesParsed;
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(src.buf);
  const ws = wb.worksheets[0];

  const maxCol = ws.actualColumnCount;
  const colNivel = maxCol + 1, colCuitV = maxCol + 2, colMontoV = maxCol + 3, colFechaV = maxCol + 4;

  ws.getCell(1, colNivel).value = 'NIVEL DE MATCH';
  ws.getCell(1, colCuitV).value = 'CUIT/CUIL';
  ws.getCell(1, colMontoV).value = 'MONTO';
  ws.getCell(1, colFechaV).value = 'FECHA';
  [colNivel, colCuitV, colMontoV, colFechaV].forEach(c => {
    ws.getCell(1, c).font = { bold: true };
  });

  const { colObs, colRefBco } = src.headerMap;

  state.cliRows.forEach(row => {
    const r = row.excelRow;
    const nivelCell = ws.getCell(r, colNivel);
    nivelCell.value = row.nivel;
    nivelCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS[row.nivel].fill } };
    nivelCell.font = { color: { argb: 'FF000000' } };

    ws.getCell(r, colCuitV).value = row.cuitOk ? '✔' : '✖';
    ws.getCell(r, colMontoV).value = row.montoOk ? '✔' : '✖';
    ws.getCell(r, colFechaV).value = row.fechaOk ? '✔' : '✖';

    if (row.nivel === 'REPETIDO') {
      if (colObs) {
        const cell = ws.getCell(r, colObs);
        const existing = cell.value;
        cell.value = existing ? `REPETIDO - ${existing}` : 'REPETIDO';
      }
      return;
    }

    if (row.bidx) {
      const bFecha = row.bancoFecha;
      const key = bFecha ? dateKey(bFecha) : null;
      const tc = key ? state.ratesByDate[key] : null;
      if (colRefBco) ws.getCell(r, colRefBco).value = row.bancoRef;
      if (colObs) {
        const cell = ws.getCell(r, colObs);
        if (!cell.value) {
          const fechaStr = bFecha ? dateLabel(bFecha) : '';
          cell.value = tc
            ? `Acreditado el ${fechaStr} - Cotización USD BNA compra: $${tc.toLocaleString('es-AR')}`
            : `Acreditado el ${fechaStr}`;
        }
      }
    }
  });

  // RESUMEN sheet
  const existing = wb.getWorksheet('RESUMEN');
  if (existing) wb.removeWorksheet(existing.id);
  const resumen = wb.addWorksheet('RESUMEN');
  resumen.getCell('A1').value = 'RESUMEN DE CONCILIACIÓN';
  resumen.getCell('A1').font = { bold: true, size: 14 };

  const niveles = ['ALTO', 'MEDIO', 'BAJO', 'SIN MATCH', 'REPETIDO'];
  const total = state.cliRows.length;
  const counts = {};
  niveles.forEach(n => counts[n] = state.cliRows.filter(r => r.nivel === n).length);

  resumen.getCell('A3').value = 'NIVEL'; resumen.getCell('B3').value = 'CANTIDAD'; resumen.getCell('C3').value = '%';
  ['A3', 'B3', 'C3'].forEach(a => resumen.getCell(a).font = { bold: true });

  let rr = 4;
  niveles.forEach(n => {
    const c = counts[n];
    resumen.getCell(rr, 1).value = n;
    resumen.getCell(rr, 2).value = c;
    resumen.getCell(rr, 3).value = total ? `${((c / total) * 100).toFixed(1)}%` : '0%';
    [1, 2, 3].forEach(col => {
      resumen.getCell(rr, col).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS[n].fill } };
    });
    rr++;
  });
  resumen.getCell(rr, 1).value = 'TOTAL';
  resumen.getCell(rr, 2).value = total;
  resumen.getCell(rr, 1).font = { bold: true };
  resumen.getCell(rr, 2).font = { bold: true };
  rr += 2;
  resumen.getCell(rr, 1).value = 'Cotizaciones BNA compra utilizadas:';
  resumen.getCell(rr, 1).font = { bold: true };
  rr++;
  Object.keys(state.ratesByDate).sort().forEach(k => {
    const val = state.ratesByDate[k];
    const [y, m, d] = k.split('-').map(Number);
    resumen.getCell(rr, 1).value = dateLabel(new Date(y, m - 1, d));
    resumen.getCell(rr, 2).value = val ? `$${val.toLocaleString('es-AR')}` : '';
    rr++;
  });

  return wb;
}

async function generateBancoOutput() {
  const src = state.bancoParsed;
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(src.buf);
  const ws = wb.worksheets[0];
  const { colObs, colTipo, colTC } = src.headerMap;

  const matchedByRow = new Map();
  state.cliRows.forEach(row => {
    if (row.bidx) matchedByRow.set(row.bidx.excelRow, row);
  });

  matchedByRow.forEach((cliRow, bancoExcelRow) => {
    if (colObs) {
      const cell = ws.getCell(bancoExcelRow, colObs);
      cell.value = `PAX ${cliRow.nro ?? ''} ${cliRow.nombre ?? ''} ${cliRow.dni ?? ''}`;
      cell.font = { color: { argb: PURPLE_ARGB } };
    }
    if (colTipo) {
      const cell = ws.getCell(bancoExcelRow, colTipo);
      if (!cell.value) cell.value = 'Transf Pax';
    }
    const bancoRow = state.bancoRows.find(b => b.excelRow === bancoExcelRow);
    if (colTC && bancoRow && bancoRow.cuit) {
      const key = bancoRow.fecha ? dateKey(bancoRow.fecha) : null;
      const tc = key ? state.ratesByDate[key] : null;
      if (tc) {
        const cell = ws.getCell(bancoExcelRow, colTC);
        cell.value = tc;
        cell.font = { color: { argb: PURPLE_ARGB } };
      }
    }
  });

  return wb;
}

function downloadWorkbook(wb, filename) {
  wb.xlsx.writeBuffer().then(buffer => {
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  });
}

function todayFilenameSuffix() {
  const d = new Date();
  const p = n => String(n).padStart(2, '0');
  return `${p(d.getDate())}-${p(d.getMonth() + 1)}-${d.getFullYear()}`;
}

/* ---------------- UI wiring ---------------- */

function setupDropzone(zoneId, inputId, filenameId, onFile) {
  const zone = document.getElementById(zoneId);
  const input = document.getElementById(inputId);
  const fnEl = document.getElementById(filenameId);

  zone.addEventListener('click', () => input.click());
  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('drag'));
  zone.addEventListener('drop', e => {
    e.preventDefault();
    zone.classList.remove('drag');
    if (e.dataTransfer.files.length) {
      input.files = e.dataTransfer.files;
      handleFile(input.files[0]);
    }
  });
  input.addEventListener('change', () => {
    if (input.files.length) handleFile(input.files[0]);
  });

  function handleFile(file) {
    fnEl.textContent = file.name;
    onFile(file);
  }
}

function checkReadyToAnalyze() {
  const btn = document.getElementById('btn-analyze');
  btn.disabled = !(state.bancoFile && state.clientesFile);
}

async function onAnalyze() {
  const btn = document.getElementById('btn-analyze');
  btn.innerHTML = '<span class="spinner"></span> Analizando…';
  btn.disabled = true;

  const clientesParsed = await parseClientes(state.clientesFile);
  const bancoParsed = await parseBanco(state.bancoFile);

  state.clientesParsed = clientesParsed;
  state.bancoParsed = bancoParsed;
  state.cliRows = clientesParsed.rows;
  state.bancoRows = bancoParsed.rows;

  runMatching();

  document.getElementById('step-empty').style.display = 'none';
  document.getElementById('step-rates').style.display = 'block';
  document.getElementById('step-results').style.display = 'block';

  await renderRatesStep();
  renderSummary();
  renderTable();

  btn.textContent = 'Analizar';
  btn.disabled = false;
}

document.addEventListener('DOMContentLoaded', () => {
  setupDropzone('dz-banco', 'file-banco', 'fn-banco', file => {
    state.bancoFile = file;
    checkReadyToAnalyze();
  });
  setupDropzone('dz-clientes', 'file-clientes', 'fn-clientes', file => {
    state.clientesFile = file;
    checkReadyToAnalyze();
  });

  document.getElementById('btn-analyze').addEventListener('click', onAnalyze);

  document.getElementById('btn-aplicar').addEventListener('click', () => {
    if (!allRatesFilled()) return;
    state.aplicado = true;
    refreshDownloadAvailability();
  });

  document.getElementById('btn-refetch').addEventListener('click', async () => {
    const today = toDateOnly(new Date());
    const key = dateKey(today);
    const input = document.getElementById(`rate-${key}`);
    if (!input) return;
    const val = await fetchOfficialRate();
    if (val) {
      input.value = val;
      state.ratesByDate[key] = val;
      state.aplicado = false;
      updateRateStatus(key);
      refreshAplicarAvailability();
      refreshDownloadAvailability();
      document.getElementById('btn-refetch').style.display = 'none';
    }
  });

  document.getElementById('btn-download-clientes').addEventListener('click', async () => {
    const btn = document.getElementById('btn-download-clientes');
    const original = btn.textContent;
    btn.innerHTML = '<span class="spinner"></span> Generando…';
    btn.disabled = true;

    const cliWb = await generateClientesOutput();
    const suffix = todayFilenameSuffix();
    downloadWorkbook(cliWb, `CLIENTES_TRANSFERENCIAS_EN_USD_${suffix}.xlsx`);

    btn.textContent = original;
    refreshDownloadAvailability();
  });

  document.getElementById('btn-download-banco').addEventListener('click', async () => {
    const btn = document.getElementById('btn-download-banco');
    const original = btn.textContent;
    btn.innerHTML = '<span class="spinner"></span> Generando…';
    btn.disabled = true;

    const bancoWb = await generateBancoOutput();
    const suffix = todayFilenameSuffix();
    downloadWorkbook(bancoWb, `BANCO_MACRO_INGRESO_EN_USD_${suffix}.xlsx`);

    btn.textContent = original;
    refreshDownloadAvailability();
  });
});
