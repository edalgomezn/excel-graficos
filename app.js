(function(){
  'use strict';

  // ===== Config / constantes
  const BLUE = '#1f77b4';
  const NF = new Intl.NumberFormat('es-ES');
  const RX_HHMM = /(\d{1,2}):(\d{2})(?::\d{2})?/;

  // Inserción en Excel: colocar gráfico desde F2 (arriba a la derecha)
  const EMBED_START_COL = 'F';
  const EMBED_START_ROW = 2;
  const EMBED_WIDTH_COLS = 10;   // F..(F + 9)
  const EMBED_HEIGHT_ROWS = 20;  // alto aprox en filas

  // ===== DOM
  const $file   = document.getElementById('fileInput');
  const $btn    = document.getElementById('btnProcess');
  const $btnEmb = document.getElementById('btnEmbed');
  const $btnClr = document.getElementById('btnClear');
  const $spin   = document.getElementById('spin');
  const $info   = document.getElementById('fileInfo');
  const $alert  = document.getElementById('alertBox');
  const $grid   = document.getElementById('chartsGrid');

  // ===== Estado
  let fileName = '';
  let validatedBuffer = null;       // ArrayBuffer validado por validator.js
  const charts = Object.create(null);
  let processedSheets = [];         // [{sheetName, canvasId}]

  // ===== Chart.js options
  const chartOptions = Object.freeze({
    responsive: true,
    interaction: { mode: 'index', intersect: false },
    plugins: {
      legend: { display: true, position: 'bottom' },
      tooltip: {
        callbacks: {
          title: items => items[0]?.label ?? '',
          label: ctx => `${ctx.dataset.label}: ${NF.format(ctx.parsed.y)}`
        }
      }
    },
    scales: {
      x: {
        title: { display: true, text: 'Hora (hh:mm)' },
        ticks: { minRotation: 90, maxRotation: 90, autoSkip: true, autoSkipPadding: 4, font: { size: 10 } },
        grid: { color: 'rgba(0,0,0,0.08)' }
      },
      y: {
        title: { display: true, text: 'Columna B (Y)' },
        beginAtZero: true,
        ticks: { callback: v => NF.format(v) },
        grid: { color: 'rgba(0,0,0,0.08)' }
      }
    }
  });

  // ===== Eventos UI
  $btn.addEventListener('click', onProcess);
  $btnEmb.addEventListener('click', embedChartsIntoExcel);
  $btnClr.addEventListener('click', resetAll);

  // Escuchar cuando la validación previa pasó
  document.addEventListener('validation:passed', (ev) => {
    fileName = ev.detail.fileName || 'Libro';
    validatedBuffer = ev.detail.arrayBuffer || null;
    // Habilitar limpiar; "Procesar" ya fue habilitado por validator.js
    $btnClr.disabled = false;
    // Mensaje informativo
    console.debug('[validator] OK', { fileName, okSheets: ev.detail.okSheets });
  });

  function onProcess(){
    if (!validatedBuffer){
      showAlert('Primero selecciona un archivo válido (pasa validación).', 'danger');
      return;
    }
    processWorkbook(validatedBuffer);
  }

  // ===== Lógica principal: leer Excel desde buffer validado, graficar todas las hojas
  function processWorkbook(arrayBuffer){
    setBusy(true); clearAlert(); destroyCharts(); hideCharts(); clearGrid(); processedSheets = [];

    try{
      const wb = XLSX.read(new Uint8Array(arrayBuffer), { type:'array' });
      let rendered = 0;

      wb.SheetNames.forEach((sheetName, idx) => {
        const ws = wb.Sheets[sheetName];
        if (!ws || !ws['!ref']) return; // ya validado, pero por seguridad

        const { labels, values } = extractSorted(ws);
        if (!labels.length) return;

        const canvasId = createChartCard(sheetName, idx);
        renderChart(canvasId, sheetName.toUpperCase(), labels, values);
        processedSheets.push({ sheetName, canvasId });
        rendered++;
      });

      if (!rendered){
        showAlert('No se generaron gráficos. Revisa datos en columnas B (Y) y C (X).', 'danger');
      }else{
        showAlert(`Listo: ${rendered} gráfico(s) generados.`, 'success');
        showCharts();
        const canEmbed = processedSheets.length && processedSheets.every(p => document.getElementById(p.canvasId));
        $btnEmb.disabled = !canEmbed;
      }
    }catch(err){
      console.error(err);
      showAlert('Error procesando el archivo: ' + err.message, 'danger');
    }finally{
      setBusy(false);
    }
  }

  // Crear tarjeta+canvas para una hoja
  function createChartCard(sheetName, idx){
    const safe = slug(`${sheetName}-${idx}`);
    const canvasId = `chart-${safe}`;

    const col = document.createElement('div');
    col.className = 'col-12';
    col.id = `col-${canvasId}`;
    col.innerHTML = `
      <div class="card chart-card shadow-sm h-100">
        <div class="card-body">
          <div class="d-flex align-items-center justify-content-between mb-2">
            <h2 class="h5 m-0">${escapeHtml(sheetName.toUpperCase())}</h2>
            <span class="badge text-bg-secondary fw-mono">B→Y · C→X (hh:mm)</span>
          </div>
          <canvas id="${canvasId}" aria-label="Gráfico ${escapeHtml(sheetName)}" role="img"></canvas>
        </div>
      </div>`;
    $grid.appendChild(col);
    return canvasId;
  }

  function clearGrid(){ $grid.innerHTML = ''; }

  // Extrae y ordena por hora asc
  function extractSorted(ws){
    const range = XLSX.utils.decode_range(ws['!ref']);
    const rows = [];
    for (let r = range.s.r + 1; r <= range.e.r; r++){
      const y = toNumber(ws[XLSX.utils.encode_cell({ r, c: 1 })]?.v);   // B
      const t = toHHMM (ws[XLSX.utils.encode_cell({ r, c: 2 })]?.v);   // C
      if (y !== null && Number.isFinite(y) && t){ rows.push({ key: t.key, hhmm: t.hhmm, y, idx:r }); }
    }
    rows.sort((a,b)=> (a.key - b.key) || (a.idx - b.idx));
    return { labels: rows.map(r=>r.hhmm), values: rows.map(r=>r.y) };
  }

  // Parsers
  function toNumber(v){
    if (v === null || v === undefined) return null;
    if (typeof v === 'number') return v;
    if (typeof v === 'string'){
      let s = v.trim(), isPct = /%$/.test(s);
      s = s.replace(/\s+/g,'').replace(/%$/,'');
      const neg = /^\(.*\)$/.test(s); if (neg) s = s.replace(/^\(|\)$/g,'');
      s = s.replace(/[^0-9eE\.,\-+]/g,'');
      const c = s.lastIndexOf(','), d = s.lastIndexOf('.');
      if (c !== -1 && d !== -1){ s = (c > d) ? s.replace(/\./g,'').replace(',', '.') : s.replace(/,/g,''); }
      else if (c !== -1){ s = s.replace(',', '.'); }
      let n = Number(s); if (!Number.isFinite(n)) return null;
      if (neg) n = -Math.abs(n); if (isPct) n /= 100;
      return n;
    }
    return null;
  }

  function toHHMM(v){
    if (v === null || v === undefined) return null;

    if (typeof v === 'number'){
      try{
        const d = XLSX.SSF.parse_date_code(v);
        if (d && (d.H !== undefined || d.M !== undefined)){
          const hh = d.H || 0, mm = d.M || 0;
          return { hhmm: `${String(hh).padStart(2,'0')}:${String(mm).padStart(2,'0')}`, key: hh*60+mm };
        }
      }catch{}
      v = String(v);
    }

    if (typeof v === 'string'){
      const t = v.trim();
      const m = t.match(RX_HHMM);
      if (m){
        let hh = parseInt(m[1],10), mm = parseInt(m[2],10);
        const ampm = t.match(/\b([AP]\.?M\.?|AM|PM)\b/i);
        if (ampm){ const am = /^A/i.test(ampm[1]); if (am && hh===12) hh=0; if (!am && hh<12) hh+=12; }
        if (hh<0||hh>23||mm<0||mm>59) return null;
        return { hhmm: `${String(hh).padStart(2,'0')}:${String(mm).padStart(2,'0')}`, key: hh*60+mm };
      }
      const dot = t.match(/^(\d{1,2})\.(\d{2})$/);
      if (dot){
        const hh = +dot[1], mm = +dot[2];
        if (hh<0||hh>23||mm<0||mm>59) return null;
        return { hhmm: `${String(hh).padStart(2,'0')}:${String(mm).padStart(2,'0')}`, key: hh*60+mm };
      }
      // En app mantenemos política laxa: (si llegara aquí) lo descartamos (ya fue validado antes)
      return null;
    }
    return null;
  }

  // Render chart
  function renderChart(canvasId, title, labels, values){
    if (charts[canvasId]) { try{ charts[canvasId].destroy(); }catch{} }
    const ctx = document.getElementById(canvasId);
    charts[canvasId] = new Chart(ctx, {
      type: 'line',
      data: {
        labels,
        datasets: [{
          label: title,
          data: values,
          fill: false,
          borderColor: BLUE,
          pointBorderColor: BLUE,
          pointBackgroundColor: '#ffffff',
          borderWidth: 3,
          pointRadius: 4,
          pointHoverRadius: 5,
          tension: 0
        }]
      },
      options: { ...chartOptions, plugins: { ...chartOptions.plugins, title: { display:true, text:title } } }
    });
  }

  // Incrustar en Excel (todas las hojas procesadas)
  async function embedChartsIntoExcel(){
    try{
      if (!validatedBuffer){ showAlert('No hay archivo validado en memoria.', 'danger'); return; }
      if (!window.ExcelJS){ showAlert('ExcelJS no está cargado.', 'danger'); return; }
      if (!processedSheets.length){ showAlert('Primero genera los gráficos (Procesar).', 'warning'); return; }

      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(validatedBuffer);

      for (const {sheetName, canvasId} of processedSheets){
        const ws = wb.getWorksheet(sheetName);
        if (!ws){ showAlert(`(Aviso) No se encontró la hoja <b>${escapeHtml(sheetName)}</b> para incrustar.`, 'warning'); continue; }

        const canvas = document.getElementById(canvasId);
        if (!canvas){ showAlert(`(Aviso) Falta el canvas del gráfico ${escapeHtml(sheetName)}.`, 'warning'); continue; }

        const dataURL = canvas.toDataURL('image/png');
        const imageId = wb.addImage({ base64: dataURL, extension: 'png' });

        const startCol = EMBED_START_COL;
        const startRow = EMBED_START_ROW;
        const endCol   = shiftCol(startCol, EMBED_WIDTH_COLS - 1);
        const endRow   = startRow + EMBED_HEIGHT_ROWS - 1;
        const range    = `${startCol}${startRow}:${endCol}${endRow}`;

        ws.addImage(imageId, range);
      }

      const out = await wb.xlsx.writeBuffer();
      const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const a = document.createElement('a');
      const url = URL.createObjectURL(blob);
      a.href = url;
      const base = (fileName || 'Libro').replace(/\.(xlsx|xls)$/i,'');
      a.download = `${base}-con-graficos.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);

      showAlert(`Descarga lista: Excel con ${processedSheets.length} gráfico(s) incrustado(s) desde F2.`, 'success');
    }catch(err){
      console.error(err);
      showAlert('No se pudo incrustar y descargar: ' + err.message, 'danger');
    }
  }

  // ===== Utilidades Excel (columnas) =====
  function shiftCol(colLetter, delta){
    const colNum = colToNum(colLetter) + delta;
    return numToCol(colNum);
  }
  function colToNum(col){
    let n = 0;
    for (let i=0; i<col.length; i++){ n = n*26 + (col.charCodeAt(i)-64); }
    return n;
  }
  function numToCol(n){
    let s = '';
    while (n > 0){ const r = (n-1)%26; s = String.fromCharCode(65+r) + s; n = Math.floor((n-1)/26); }
    return s;
  }

  // ===== Utilidades UI
  function showCharts(){ $grid.classList.remove('d-none'); }
  function hideCharts(){ $grid.classList.add('d-none'); }
  function destroyCharts(){ for (const id in charts){ try{ charts[id].destroy(); }catch{} delete charts[id]; } }
  function showAlert(html, type='info'){ $alert.innerHTML = `<div class="alert alert-${type}" role="alert">${html}</div>`; }
  function clearAlert(){ $alert.innerHTML = ''; }
  function setBusy(b){
    $btn.disabled = b || !validatedBuffer;
    const canEmbed = !!validatedBuffer && processedSheets.length && processedSheets.every(p => document.getElementById(p.canvasId));
    $btnEmb.disabled = b || !canEmbed;
    $btnClr.disabled = b || (!$file.files?.length);
    $spin.classList.toggle('d-none', !b);
  }
  function resetAll(){
    window.__validated = null;
    validatedBuffer = null;
    fileName = '';
    processedSheets = [];
    destroyCharts(); hideCharts(); clearAlert(); clearGrid();
    $info.textContent = 'Sin archivos seleccionados';
    $file.value = '';
    $btn.disabled = true;
    $btnEmb.disabled = true;
    $btnClr.disabled = true;
  }

  function slug(s){
    return String(s)
      .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
      .replace(/[^a-zA-Z0-9\-_. ]+/g,'')
      .trim().replace(/\s+/g,'-').toLowerCase();
  }
  function escapeHtml(s){
    return String(s).replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
  }
})();
