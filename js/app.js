// ===== State =====
let rawData = [];
let cycles = {};
let cycleSummary = [];
let allCycleNums = [];
let firstCycleSkipped = false;

// ===== Column name mapping =====
const COL_MAP = {
  'Index':'Index','Step':'Step','Loop1':'Loop1','Loop2':'Loop2',
  'Action':'Action','Model':'Model','Step Time(ms)':'StepTime_ms',
  'Total Time(ms)':'TotalTime_ms','Voltage(mV)':'Voltage_mV',
  'Current(mA)':'Current_mA','Power(W)':'Power_W',
  'Temperature(C)':'Temp_C','mAh':'mAh','Wh':'Wh',
  'DCIR':'DCIR','Ch Status':'ChStatus'
};

// Unit conversion
const TO_A = 1/1000;
const TO_AH = 1/1000;
const TO_V = 1/1000;
const TO_S = 1/1000;

// ===== Excel parsing =====
function parseExcel(ab) {
  const wb = XLSX.read(ab, {type:'array'});
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, {defval:null});

  const headers = Object.keys(rows[0]);
  const colIdx = {};
  headers.forEach(h => {
    const key = COL_MAP[h] || h;
    colIdx[key] = h;
  });

  rawData = rows.map((r) => {
    const get = (key) => {
      const orig = colIdx[key];
      return orig ? r[orig] : undefined;
    };
    const loop1 = get('Loop1');
    const cn = parseCycleNum(loop1);
    return {
      Loop1: loop1,
      Action: (get('Action') || '').trim().toUpperCase(),
      Model: (get('Model') || '').trim(),
      StepTime_s: (+get('StepTime_ms') || 0) * TO_S,
      Voltage_V: (+get('Voltage_mV') || 0) * TO_V,
      Current_A: (+get('Current_mA') || 0) * TO_A,
      Temp_C: +get('Temp_C') || 0,
      Ah: (+get('mAh') || 0) * TO_AH,
      CycleNum: cn,
    };
  }).filter(r => r.CycleNum !== null);
}

function parseCycleNum(v) {
  if (v === null || v === undefined) return null;
  const m = String(v).trim().match(/L(\d+)/);
  return m ? parseInt(m[1]) : null;
}

function parseHms(s) {
  if (!s) return 0;
  const parts = String(s).trim().split(':');
  if (parts.length === 3) {
    return (+parts[0]) * 3600 + (+parts[1]) * 60 + (+parts[2]);
  } else if (parts.length === 2) {
    return (+parts[0]) * 60 + (+parts[1]);
  }
  return +(s) || 0;
}

// ===== FUD parsing =====
function parseFud(text) {
  const lines = text.split(/\r?\n/);
  let headerIdx = -1;
  for (let i = 0; i < lines.length; i++) {
    if (lines[i].includes('System time') && lines[i].includes('Action')) {
      headerIdx = i;
      break;
    }
  }
  if (headerIdx === -1) throw new Error('找不到 FUD 檔案表頭');

  const headers = lines[headerIdx].split(',').map(h => h.trim());
  const colIdx = {};
  headers.forEach((h, i) => { colIdx[h] = i; });

  const rows = [];
  for (let i = headerIdx + 1; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;
    const fields = line.split(',');
    const get = (key) => {
      const idx = colIdx[key];
      return (idx !== undefined && idx < fields.length) ? fields[idx].trim() : '';
    };

    const action = (get('Action') || '').replace(/[\\"]/g, '').trim().toUpperCase();
    const loop1 = get('Loop1');
    const cn = parseCycleNum(loop1);

    rows.push({
      Loop1: loop1,
      Action: action,
      Model: action.startsWith('CC') ? 'CC-CV' : '',
      StepTime_s: parseHms(get('Step time')),
      Voltage_V: +(get('V')) || 0,
      Current_A: +(get('I')) || 0,
      Temp_C: +(get('T')) || 0,
      Ah: (+(get('mAh')) || 0) / 1000,
      CycleNum: cn,
    });
  }

  rawData = rows;

  // Auto-detect cycle numbers
  let cycleNum = 1;
  let prevAction = '';
  for (let i = 0; i < rawData.length; i++) {
    if (rawData[i].CycleNum === null) {
      const a = rawData[i].Action;
      if (prevAction === 'REST' && (a === 'DISCHARGE' || a === 'CHARGE')) {
        cycleNum++;
      }
      rawData[i].CycleNum = cycleNum;
      rawData[i].Loop1 = 'L' + cycleNum;
    } else {
      cycleNum = rawData[i].CycleNum;
    }
    prevAction = rawData[i].Action;
  }
}

// ===== Data processing =====
function processData() {
  cycles = {};
  rawData.forEach(r => {
    const cn = r.CycleNum;
    if (!cycles[cn]) cycles[cn] = [];
    cycles[cn].push(r);
  });

  allCycleNums = Object.keys(cycles).map(Number).sort((a,b)=>a-b);
  if (allCycleNums.length === 0) return;

  const minCycle = allCycleNums[0];
  firstCycleSkipped = false;

  // Check if first cycle has CHARGE before DISCHARGE → skip its CHARGE
  const firstRows = cycles[minCycle];
  const firstAction = firstRows.length > 0 ? firstRows[0].Action : '';
  const hasChargeFirst = firstAction === 'CHARGE';
  if (hasChargeFirst) {
    cycles[minCycle] = firstRows.filter(r => r.Action !== 'CHARGE');
    firstCycleSkipped = true;
  }

  cycleSummary = allCycleNums.map(cn => {
    const rows = cycles[cn];
    const discharge = rows.filter(r => r.Action === 'DISCHARGE');
    const charge = cn === minCycle && hasChargeFirst ? [] : rows.filter(r => r.Action === 'CHARGE');

    const qDisAh = discharge.length > 0 ? Math.max(...discharge.map(r => r.Ah)) : 0;
    const qChgAh = charge.length > 0 ? Math.max(...charge.map(r => r.Ah)) : 0;
    const ce = qChgAh > 0 ? qDisAh / qChgAh * 100 : NaN;

    let ccRatio = NaN;
    if (charge.length > 0) {
      const currents = charge.map(r => r.Current_A);
      const maxI = Math.max(...currents);
      if (maxI > 0) {
        const threshold = maxI * 0.95;
        for (let i = 0; i < currents.length; i++) {
          if (currents[i] < threshold) {
            const ccAh = i > 0 ? charge[i].Ah : 0;
            ccRatio = qChgAh > 0 ? ccAh / qChgAh * 100 : NaN;
            break;
          }
        }
      }
    }

    return { Cycle: cn, Q_dis_Ah: qDisAh, Q_chg_Ah: qChgAh, CE_pct: ce, CC_ratio_pct: ccRatio };
  });

  const maxQ = Math.max(...cycleSummary.map(s => s.Q_dis_Ah));
  cycleSummary.forEach(s => s.Retention_pct = maxQ > 0 ? s.Q_dis_Ah / maxQ * 100 : 0);
}

function zeroRefCapacity(rows) {
  const result = rows.map(r => ({...r, Ah_rel: r.Ah}));
  const actions = [...new Set(result.map(r => r.Action))];
  actions.forEach(action => {
    const indices = result.map((r,i) => ({r,i})).filter(x => x.r.Action === action);
    if (indices.length > 0) {
      const startAh = result[indices[0].i].Ah;
      indices.forEach(x => { result[x.i].Ah_rel = result[x.i].Ah - startAh; });
    }
  });
  return result;
}

function computeDQDV(action, cycleRange, sgWin, sgOrd) {
  const results = [];
  const a = action.toUpperCase();
  allCycleNums.forEach(cn => {
    if (cycleRange && (cn < cycleRange[0] || cn > cycleRange[1])) return;
    const rows = cycles[cn];
    let seg = cn === allCycleNums[0] && firstCycleSkipped && a === 'CHARGE' ? [] : zeroRefCapacity(rows).filter(r => r.Action === a);
    // Exclude CV phase in CHARGE
    if (a === 'CHARGE' && seg.length > 0) {
      const maxI = Math.max(...seg.map(r => Math.abs(r.Current_A)));
      const cvThreshold = maxI * 0.15;
      seg = seg.filter(r => Math.abs(r.Current_A) >= cvThreshold);
    }
    // Boundary trim
    const trimN = seg.length >= 20 ? 6 : seg.length >= 10 ? 4 : seg.length >= 6 ? 2 : 0;
    const segTrimmed = seg.slice(trimN, -trimN);
    if (segTrimmed.length < 4) return;
    seg = segTrimmed;

    const v = seg.map(r => r.Voltage_V);
    const q = seg.map(r => r.Ah_rel);
    const n = v.length;
    const dqdv = new Array(n).fill(NaN);
    for (let i = 1; i < n - 1; i++) {
      const dv = v[i+1] - v[i-1];
      if (Math.abs(dv) > 1e-3) dqdv[i] = (q[i+1] - q[i-1]) / dv;
    }

    const valid = dqdv.map((x,i) => ({x,i})).filter(d => !isNaN(d.x));
    const validValues = valid.map(d => d.x);
    const validIndices = valid.map(d => d.i);
    if (validValues.length >= sgWin && sgWin >= 3 && sgWin % 2 === 1) {
      try { const sm = savgolay(validValues, sgWin, sgOrd); validIndices.forEach((idx,k) => { dqdv[idx]=sm[k]; }); } catch(e) {}
    }
    const pts = [];
    for (let i = 0; i < n; i++) {
      if (!isNaN(dqdv[i]) && Math.abs(dqdv[i]) < 100) {
        pts.push({ v: v[i], d: dqdv[i] });
      }
    }
    pts.sort((p, q) => p.v - q.v);
    pts.forEach(p => results.push({ Cycle: cn, Voltage_V: p.v, dQdV: p.d }));
  });
  return results;
}

function savgolay(y, w, order) {
  const pad = Math.floor(w / 2);
  const n = y.length;
  if (n < w || w < 3 || w % 2 !== 1) return y.slice();
  const XtX = [];
  for (let r = 0; r <= order; r++) XtX.push(new Array(order + 1).fill(0));
  for (let k = 0; k < w; k++) {
    const j = k - pad;
    let p = 1;
    for (let r = 0; r <= order; r++) {
      let q = 1;
      for (let c = 0; c <= order; c++) {
        XtX[r][c] += p * q;
        q *= j;
      }
      p *= j;
    }
  }
  const res = new Array(n);
  for (let i = 0; i < n; i++) {
    const XtY = new Array(order + 1).fill(0);
    for (let k = 0; k < w; k++) {
      const idx = Math.max(0, Math.min(n - 1, i + k - pad));
      const j = k - pad;
      let p = 1;
      for (let r = 0; r <= order; r++) {
        XtY[r] += p * y[idx];
        p *= j;
      }
    }
    const a = solveLinear(XtX.slice().map(row => row.slice()), XtY.slice());
    res[i] = a[0];
  }
  return res;
}

function solveLinear(A, b) {
  const n = b.length;
  for (let i = 0; i < n; i++) A[i].push(b[i]);
  for (let i = 0; i < n; i++) {
    let maxRow = i;
    for (let k = i + 1; k < n; k++) {
      if (Math.abs(A[k][i]) > Math.abs(A[maxRow][i])) maxRow = k;
    }
    [A[i], A[maxRow]] = [A[maxRow], A[i]];
    if (Math.abs(A[i][i]) < 1e-12) continue;
    for (let k = i + 1; k < n; k++) {
      const factor = A[k][i] / A[i][i];
      for (let j = i; j <= n; j++) A[k][j] -= factor * A[i][j];
    }
  }
  const x = new Array(n).fill(0);
  for (let i = n - 1; i >= 0; i--) {
    let s = A[i][n];
    for (let j = i + 1; j < n; j++) s -= A[i][j] * x[j];
    x[i] = Math.abs(A[i][i]) < 1e-12 ? 0 : s / A[i][i];
  }
  return x;
}

// ===== Common axis layout =====
function baseLayout() {
  const fs = parseInt(document.getElementById('fontSize').value) || 12;
  const padL = Math.max(65, 36 + fs * 3);
  const padB = Math.max(55, 36 + fs * 2.5);
  return {
    font: { color: '#000000', size: fs, family: 'gg sans, sans-serif' },
    margin: { l: padL, r: 20, t: Math.round(fs * 4 + 22), b: padB },
    paper_bgcolor: '#FAF9F5',
    plot_bgcolor: '#FAF9F5',
    hovermode: 'x unified',
    xaxis: { title: { font: { color: '#000000' }, standoff: Math.max(14, fs * 1.0) }, tickfont: { color: '#000000' }, color: '#000000', showline: true, linecolor: '#000000', linewidth: 1, mirror: true, ticks: 'outside' },
    yaxis: { title: { font: { color: '#000000' }, standoff: Math.max(14, fs * 1.0) }, tickfont: { color: '#000000' }, color: '#000000', showline: true, linecolor: '#000000', linewidth: 1, mirror: true, ticks: 'outside' },
  };
}

function applyAxisRange(layout, prefix) {
  const xmn = document.getElementById(prefix + 'Xmin').value;
  const xmx = document.getElementById(prefix + 'Xmax').value;
  const ymn = document.getElementById(prefix + 'Ymin').value;
  const ymx = document.getElementById(prefix + 'Ymax').value;
  if (xmn !== '' && xmx !== '') layout.xaxis.range = [parseFloat(xmn), parseFloat(xmx)];
  else if (xmn !== '') layout.xaxis.range = [parseFloat(xmn), layout.xaxis.range ? layout.xaxis.range[1] : undefined];
  else if (xmx !== '') layout.xaxis.range = [layout.xaxis.range ? layout.xaxis.range[0] : undefined, parseFloat(xmx)];
  if (ymn !== '' && ymx !== '') layout.yaxis.range = [parseFloat(ymn), parseFloat(ymx)];
  else if (ymn !== '') layout.yaxis.range = [parseFloat(ymn), layout.yaxis.range ? layout.yaxis.range[1] : undefined];
  else if (ymx !== '') layout.yaxis.range = [layout.yaxis.range ? layout.yaxis.range[0] : undefined, parseFloat(ymx)];
}

// ===== Get cycles to display based on user selection =====
function getCyclesToDisplay() {
  const mode = document.getElementById('curveDisplay').value;
  const [rMin, rMax] = getRange();
  const shown = allCycleNums.filter(cn => cn >= rMin && cn <= rMax);
  
  if (mode === 'all') return shown;
  if (mode === 'first_last') {
    if (shown.length <= 2) return shown;
    return [shown[0], shown[shown.length - 1]];
  }
  if (mode === 'interval_5') {
    return shown.filter((_, i) => i % 5 === 0);
  }
  if (mode === 'interval_10') {
    return shown.filter((_, i) => i % 10 === 0);
  }
  return shown;
}

// ===== Charting =====
function renderCycleLife() {
  const yKey = document.getElementById('cycleY').value;
  const refMode = document.getElementById('refMode').value;
  const refManual = parseFloat(document.getElementById('refManual').value) || 0;
  const [rMin, rMax] = getRange();

  const filtered = cycleSummary.filter(s => s.Cycle >= rMin && s.Cycle <= rMax);
  if (filtered.length === 0) return showError('所選範圍內無數據');

  let yVals, yLabel;
  const xVals = filtered.map(s => s.Cycle);

  if (yKey === 'retention') {
    const ref = refMode === 'manual' && refManual > 0 ? refManual : Math.max(...cycleSummary.map(s => s.Q_dis_Ah));
    yVals = filtered.map(s => ref > 0 ? s.Q_dis_Ah / ref * 100 : 0);
    yLabel = 'Retention (%)';
  } else if (yKey === 'ccratio') {
    yVals = filtered.map(s => s.CC_ratio_pct);
    yLabel = 'CC Ratio (%)';
  } else {
    yVals = filtered.map(s => s.CE_pct);
    yLabel = 'Coulombic Efficiency (%)';
  }

  const customdata = filtered.map(s => [s.Q_dis_Ah.toFixed(4), s.Q_chg_Ah.toFixed(4)]);

  const trace = {
    x: xVals, y: yVals,
    mode: 'markers+lines', type: 'scatter',
    marker: { size: 6, color: '#339af0', line: { width: 1, color: '#1864ab' } },
    line: { width: 1.5, color: '#339af0' },
    hovertemplate: 'Cycle %{x}<br>%{y:.2f}%<br>Q<sub>dis</sub>: %{customdata[0]} Ah<br>Q<sub>chg</sub>: %{customdata[1]} Ah<extra></extra>',
    customdata: customdata,
  };

  const fsTitle = parseInt(document.getElementById('fontSize').value) || 12;
  const layout = baseLayout();
  layout.title = { text: 'Cycle Life', x: 0.5, xanchor: 'center', y: 0.98, yanchor: 'top', font: { size: Math.round(fsTitle * 2.0), color: '#000000', family: "'gg sans Bold', 'gg sans', sans-serif" } };
  layout.xaxis.title = 'Cycle Number';
  layout.xaxis.dtick = 100;
  layout.xaxis.minor = { dtick: 50, gridcolor: '#e0ddd5' };
  layout.yaxis.title = yLabel;
  layout.yaxis.dtick = 20;
  layout.yaxis.minor = { dtick: 10, gridcolor: '#e0ddd5' };
  layout.yaxis.rangemode = 'tozero';
  applyAxisRange(layout, 'cl');

  Plotly.newPlot('chart-cycleLife', [trace], layout, {responsive: true, displayModeBar: true, scrollZoom: true, displaylogo: false, modeBarButtonsToRemove: ['sendDataToCloud', 'hoverClosestCartesian', 'hoverCompareCartesian']});
}

function renderCDCurves() {
  const mode = document.querySelector('input[name="cdMode"]:checked').value;
  const cyclesToShow = getCyclesToDisplay();
  const traces = [];
  const total = cyclesToShow.length;

  // Color gradients for distinct curves
  const chargeColor = (t) => `rgb(${Math.min(255,Math.round(255*(0.8+0.2*(1-t))))},${Math.min(255,Math.round(255*(0.4+0.4*t)))},${Math.min(255,Math.round(255*(0.2+0.3*(1-t))))})`;
  const dischargeColor = (t) => `rgb(${Math.min(255,Math.round(255*(0.2+0.3*t)))},${Math.min(255,Math.round(255*(0.5+0.3*(1-t))))},${Math.min(255,Math.round(255*(0.8+0.2*(1-t))))})`;

  cyclesToShow.forEach((cn, i) => {
    const t = total > 1 ? i / (total - 1) : 0;
    const isFirstCycle = cn === allCycleNums[0];
    const z = zeroRefCapacity(cycles[cn]);

    if (mode === 'charge' || mode === 'overlay') {
      if (isFirstCycle && firstCycleSkipped) { /* skip first charge */ }
      else {
        const seg = z.filter(r => r.Action === 'CHARGE' && Number.isFinite(r.Ah_rel) && Number.isFinite(r.Voltage_V));
        if (seg.length > 0) {
          const showLegend = (i === 0 || i === total - 1);
          const vVals = seg.map(r => r.Voltage_V);
          const vMin = Math.min(...vVals), vMax = Math.max(...vVals);
          const vRange = vMax - vMin;
          if (vRange > 0.01) {
            const vLo = vMin + vRange * 0.15;
            const vHi = vMax - vRange * 0.10;
            let segMain = seg.filter(r => r.Voltage_V >= vLo && r.Voltage_V <= vHi);
            segMain = segMain.slice().sort((a, b) => a.Ah_rel - b.Ah_rel);
            if (segMain.length >= 2) {
              const minAh = segMain[0].Ah_rel;
              const maxAh = segMain[segMain.length - 1].Ah_rel;
              const ahRange = maxAh - minAh;
              if (ahRange > 1e-6) {
                traces.push({
                  x: segMain.map(r => (r.Ah_rel - minAh) / ahRange * 100),
                  y: segMain.map(r => r.Voltage_V),
                  mode: 'lines', type: 'scatter',
                  line: { shape: 'linear', width: 1.2, color: chargeColor(t) },
                  name: showLegend ? (i===0 ? `C${cn} Charge (first)` : `C${cn} Charge (last)`) : `C${cn}`,
                  legendgroup: 'charge',
                  showlegend: showLegend,
                  hovertemplate: `Cycle ${cn}<br>V: %{y:.4f}V<br>SOC: %{x:.1f}%<extra></extra>`,
                  opacity: 0.85,
                });
              }
            }
          }
        }
      }
    }

    if (mode === 'discharge' || mode === 'overlay') {
      const seg = z.filter(r => r.Action === 'DISCHARGE' && Number.isFinite(r.Ah_rel) && Number.isFinite(r.Voltage_V));
      if (seg.length > 0) {
        const showLegend = (i === 0 || i === total - 1);
        const vVals = seg.map(r => r.Voltage_V);
        const vMin = Math.min(...vVals), vMax = Math.max(...vVals);
        const vRange = vMax - vMin;
        if (vRange > 0.01) {
          const vLo = vMin + vRange * 0.10;
          const vHi = vMax - vRange * 0.15;
          let segMain = seg.filter(r => r.Voltage_V >= vLo && r.Voltage_V <= vHi);
          segMain = segMain.slice().sort((a, b) => a.Ah_rel - b.Ah_rel);
          if (segMain.length >= 2) {
            const minAh = segMain[0].Ah_rel;
            const maxAh = segMain[segMain.length - 1].Ah_rel;
            const ahRange = maxAh - minAh;
            if (ahRange > 1e-6) {
              traces.push({
                x: segMain.map(r => (r.Ah_rel - minAh) / ahRange * 100),
                y: segMain.map(r => r.Voltage_V),
                mode: 'lines', type: 'scatter',
                line: { shape: 'linear', width: 1.2, color: dischargeColor(t) },
                name: showLegend ? (i===0 ? `C${cn} Discharge (first)` : `C${cn} Discharge (last)`) : `C${cn}`,
                legendgroup: 'discharge',
                showlegend: showLegend,
                hovertemplate: `Cycle ${cn}<br>V: %{y:.4f}V<br>SOC: %{x:.1f}%<extra></extra>`,
                opacity: 0.85,
              });
            }
          }
        }
      }
    }
  });

  const fsTitle = parseInt(document.getElementById('fontSize').value) || 12;
  const layout = baseLayout();
  layout.title = { text: 'Charge / Discharge Curves', x: 0.5, xanchor: 'center', y: 0.98, yanchor: 'top', font: { size: Math.round(fsTitle * 2.0), color: '#000000', family: "'gg sans Bold', 'gg sans', sans-serif" } };
  layout.xaxis.title = 'SOC (%)';
  layout.yaxis.title = 'Voltage (V)';
  applyAxisRange(layout, 'cd');

  Plotly.newPlot('chart-cdCurves', traces, layout, {responsive: true, displayModeBar: true, scrollZoom: true, displaylogo: false, modeBarButtonsToRemove: ['sendDataToCloud', 'hoverClosestCartesian', 'hoverCompareCartesian']});
}

function renderDQDV() {
  const rawAction = document.querySelector('input[name="dqMode"]:checked').value;
  const [rMin, rMax] = getRange();
  const sgWin = parseInt(document.getElementById('sgWindow').value);
  const sgOrd = parseInt(document.getElementById('sgOrder').value);
  const cyclesToShow = getCyclesToDisplay();

  const actions = rawAction === 'BOTH' ? ['CHARGE', 'DISCHARGE'] : [rawAction];
  const traces = [];

  actions.forEach((action, ai) => {
    const data = computeDQDV(action, [rMin, rMax], sgWin, sgOrd);
    if (data.length === 0) return;

    const grouped = {};
    data.forEach(d => { if (!grouped[d.Cycle]) grouped[d.Cycle] = []; grouped[d.Cycle].push(d); });

    const cycles_list = Object.keys(grouped).map(Number).sort((a,b)=>a-b).filter(cn => cyclesToShow.includes(cn));
    const total = cycles_list.length;

    cycles_list.forEach((cn, i) => {
      const pts = grouped[cn].filter(p => !isNaN(p.dQdV));
      if (pts.length === 0) return;
      const t = total > 1 ? i / (total - 1) : 0;
      const isCharge = action === 'CHARGE';
      const r = Math.min(255, Math.round(255 * (isCharge ? (0.8+0.2*(1-t)) : (0.2+0.3*t))));
      const g = Math.min(255, Math.round(255 * (isCharge ? (0.4+0.4*t) : (0.5+0.3*(1-t)))));
      const b = Math.min(255, Math.round(255 * (isCharge ? (0.2+0.3*(1-t)) : (0.8+0.2*(1-t)))));
      const color = `rgb(${r},${g},${b})`;

      traces.push({
        x: pts.map(p => p.Voltage_V), y: pts.map(p => p.dQdV),
        mode: 'lines', type: 'scatter',
        line: { width: 1.2, color },
        name: `C${cn} ${action}`,
        legendgroup: action,
        showlegend: i === 0 || i === total - 1,
        hovertemplate: `Cycle ${cn}<br>V: %{x:.4f}V<br>dQ/dV: %{y:.2f}<extra></extra>`,
        opacity: 0.85,
      });
    });
  });

  if (traces.length === 0) return showError('所選範圍內無足夠數據計算 dQ/dV');

  const fsTitle = parseInt(document.getElementById('fontSize').value) || 12;
  const layout = baseLayout();
  layout.title = { text: 'dQ/dV vs Voltage', x: 0.5, xanchor: 'center', y: 0.98, yanchor: 'top', font: { size: Math.round(fsTitle * 2.0), color: '#000000', family: "'gg sans Bold', 'gg sans', sans-serif" } };
  layout.xaxis.title = 'Voltage (V)';
  layout.yaxis.title = 'dQ/dV';
  applyAxisRange(layout, 'dq');

  Plotly.newPlot('chart-dqdV', traces, layout, {responsive: true, displayModeBar: true, scrollZoom: true, displaylogo: false, modeBarButtonsToRemove: ['sendDataToCloud', 'hoverClosestCartesian', 'hoverCompareCartesian']});
}

// ===== UI helpers =====
function getRange() {
  return [parseInt(document.getElementById('rangeMin').value), parseInt(document.getElementById('rangeMax').value)];
}

function updateRangeLabels() {
  document.getElementById('rangeLabelMin').textContent = document.getElementById('rangeMin').value;
  document.getElementById('rangeLabelMax').textContent = document.getElementById('rangeMax').value;
}

function showError(msg) {
  const el = document.getElementById('errorMsg');
  el.textContent = msg;
  el.style.display = 'block';
  setTimeout(() => { el.style.display = 'none'; }, 3000);
}

function getActiveChartId() {
  const active = document.querySelector('.tab-btn.active');
  return active ? 'chart-' + active.dataset.tab : null;
}

function resizeActiveChart() {
  const id = getActiveChartId();
  const el = document.getElementById(id);
  if (el && el.data) Plotly.Plots.resize(el);
}

function updateSummary() {
  if (cycleSummary.length === 0) return;
  const s = cycleSummary;
  const valid = s.filter(x => x.Q_dis_Ah > 0);
  const first = valid.length > 0 ? valid[0] : s[0];
  const last = valid.length > 0 ? valid[valid.length-1] : s[s.length-1];
  const retention = first.Q_dis_Ah > 0 ? last.Q_dis_Ah / first.Q_dis_Ah * 100 : 0;
  const ceVals = s.map(x => x.CE_pct).filter(x => !isNaN(x));
  const avgCE = ceVals.length > 0 ? ceVals.reduce((a,b) => a+b, 0) / ceVals.length : NaN;

  document.getElementById('summaryMetrics').innerHTML = `
    <div class="metric"><span class="val">${s.length}</span><span class="lbl">總循環數</span></div>
    <div class="metric"><span class="val">${first.Q_dis_Ah.toFixed(4)}</span><span class="lbl">首圈 Q<sub>dis</sub> (Ah)</span></div>
    <div class="metric"><span class="val">${last.Q_dis_Ah.toFixed(4)}</span><span class="lbl">末圈 Q<sub>dis</sub> (Ah)</span></div>
    <div class="metric"><span class="val">${retention.toFixed(1)}%</span><span class="lbl">Retention</span></div>
    <div class="metric"><span class="val">${isNaN(avgCE) ? 'N/A' : avgCE.toFixed(2) + '%'}</span><span class="lbl">平均 CE</span></div>
  `;
  document.getElementById('summaryBar').style.display = 'block';

  let html = '<table><thead><tr><th>Cycle</th><th>Q_dis (Ah)</th><th>Q_chg (Ah)</th><th>CE (%)</th><th>CC Ratio (%)</th><th>Retention (%)</th></tr></thead><tbody>';
  s.forEach(r => {
    html += `<tr><td>${r.Cycle}</td><td>${r.Q_dis_Ah.toFixed(4)}</td><td>${r.Q_chg_Ah.toFixed(4)}</td><td>${isNaN(r.CE_pct) ? 'N/A' : r.CE_pct.toFixed(2)}</td><td>${isNaN(r.CC_ratio_pct) ? 'N/A' : r.CC_ratio_pct.toFixed(2)}</td><td>${r.Retention_pct.toFixed(2)}</td></tr>`;
  });
  html += '</tbody></table>';
  document.getElementById('summaryTable').innerHTML = html;
}

let _chartObserver = null;
function setupResizeObserver() {
  if (_chartObserver) _chartObserver.disconnect();
  _chartObserver = new ResizeObserver(() => {
    ['chart-cycleLife','chart-cdCurves','chart-dqdV'].forEach(id => {
      const el = document.getElementById(id);
      if (el && el.data) Plotly.Plots.resize(el);
    });
  });
  const summary = document.getElementById('summaryBar');
  if (summary) _chartObserver.observe(summary);
}

function refreshAllCharts() {
  renderCycleLife();
  renderCDCurves();
  renderDQDV();
  updateSummary();
}

// ===== File handling =====
function handleFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  if (ext !== 'xlsx' && ext !== 'csv' && ext !== 'fud') { showError('僅支援 .xlsx / .csv / .fud 格式'); return; }
  document.getElementById('fileName').textContent = file.name;
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      if (ext === 'fud') {
        parseFud(e.target.result);
      } else {
        parseExcel(e.target.result);
      }
      processData();
      const min = allCycleNums[0], max = allCycleNums[allCycleNums.length-1];
      document.getElementById('rangeMin').min = min; document.getElementById('rangeMin').max = max; document.getElementById('rangeMin').value = min;
      document.getElementById('rangeMax').min = min; document.getElementById('rangeMax').max = max; document.getElementById('rangeMax').value = max;
      document.getElementById('rangeLabelMin').textContent = min;
      document.getElementById('rangeLabelMax').textContent = max;
      document.getElementById('placeholder').style.display = 'none';
      document.getElementById('summaryBar').style.display = 'none';
      refreshAllCharts();
      setupResizeObserver();
    } catch (err) { showError('解析錯誤：' + err.message); }
  };
  if (ext === 'fud') {
    reader.readAsText(file);
  } else {
    reader.readAsArrayBuffer(file);
  }
}

// ===== Initialization =====
document.addEventListener('DOMContentLoaded', () => {
  const uploadArea = document.getElementById('uploadArea');
  const fileInput = document.getElementById('fileInput');

  uploadArea.addEventListener('click', () => fileInput.click());
  uploadArea.addEventListener('dragover', (e) => { e.preventDefault(); uploadArea.style.borderColor = '#4299e1'; });
  uploadArea.addEventListener('dragleave', () => { uploadArea.style.borderColor = '#cbd5e0'; });
  uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.style.borderColor = '#cbd5e0';
    if (e.dataTransfer.files.length > 0) { fileInput.files = e.dataTransfer.files; handleFile(e.dataTransfer.files[0]); }
  });

  fileInput.addEventListener('change', (e) => { if (e.target.files.length > 0) handleFile(e.target.files[0]); });

  // Tabs
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.chart-pane').forEach(p => p.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById('pane-' + btn.dataset.tab).classList.add('active');
      setTimeout(resizeActiveChart, 50);
    });
  });

  // Controls
  document.getElementById('cycleY').addEventListener('change', renderCycleLife);
  document.getElementById('refMode').addEventListener('change', () => {
    document.getElementById('refManual').style.display = document.getElementById('refMode').value === 'manual' ? 'block' : 'none';
    renderCycleLife();
  });
  document.getElementById('refManual').addEventListener('input', renderCycleLife);
  document.getElementById('rangeMin').addEventListener('input', () => { updateRangeLabels(); refreshAllCharts(); });
  document.getElementById('rangeMax').addEventListener('input', () => { updateRangeLabels(); refreshAllCharts(); });
  document.querySelectorAll('input[name="cdMode"]').forEach(el => el.addEventListener('change', renderCDCurves));
  document.querySelectorAll('input[name="dqMode"]').forEach(el => el.addEventListener('change', renderDQDV));
  document.getElementById('sgWindow').addEventListener('change', renderDQDV);
  document.getElementById('sgOrder').addEventListener('change', renderDQDV);
  document.getElementById('fontSize').addEventListener('change', refreshAllCharts);
  document.getElementById('curveDisplay').addEventListener('change', () => { renderCDCurves(); renderDQDV(); });
  
  ['cl','cd','dq'].forEach(p => {
    [p+'Xmin', p+'Xmax', p+'Ymin', p+'Ymax'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.addEventListener('change', () => {
        if (p === 'cl') renderCycleLife();
        else if (p === 'cd') renderCDCurves();
        else renderDQDV();
      });
    });
  });

  // Summary toggle
  document.getElementById('summaryToggle').addEventListener('click', () => {
    const tbl = document.getElementById('summaryTable');
    const btn = document.getElementById('summaryToggle');
    tbl.classList.toggle('show');
    btn.textContent = tbl.classList.contains('show') ? '隱藏詳細資料 ▲' : '顯示詳細資料 ▼';
    if (_chartObserver) _chartObserver.disconnect();
    setupResizeObserver();
  });
});
