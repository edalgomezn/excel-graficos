(function(){
  'use strict';

  // ====== Parámetros de validación (ajustables) ======
  const MAX_FILE_SIZE_MB = 50;     // tamaño máximo permitido
  const MAX_SHEETS_WARN  = 50;     // advertir si hay más de N hojas
  const MAX_POINTS_WARN  = 10000;  // advertir si una hoja supera N puntos válidos
  const ACCEPTED_EXT     = /\.(xlsx|xls)$/i;

  // ====== DOM compartido con la app ======
  const $file  = document.getElementById('fileInput');
  const $info  = document.getElementById('fileInfo');
  const $alert = document.getElementById('alertBox');
  const $btnProcess = document.getElementById('btnProcess');
  const $btnEmbed   = document.getElementById('btnEmbed');
  const $btnClear   = document.getElementById('btnClear');

  // Estado validado compartido
  window.__validated = null; // { arrayBuffer, fileName, sheetsReport:[], okSheets:[name], totalValidPoints }

  // Eventos
  $file.addEventListener('change', onFileChange);

  function onFileChange(){
    resetState();
    const file = $file.files?.[0] || null;
    if (!file){
      showInfo('Sin archivos seleccionados');
      disableAll();
      return;
    }
    showInfo(`Seleccionado: ${file.name}`);
    validateFileAndWorkbook(file);
  }

  // ====== Validación principal ======
  async function validateFileAndWorkbook(file){
    // 0) Librerías
    const libsMissing = [];
    if (!window.XLSX)   libsMissing.push('SheetJS (XLSX)');
    if (!window.Chart)  libsMissing.push('Chart.js');
    if (!window.ExcelJS)libsMissing.push('ExcelJS');
    if (libsMissing.length){
      showAlert(`Faltan librerías: <b>${libsMissing.join(', ')}</b>. Descarga los JS y referencia localmente.`, 'danger');
      disableAll();
      return;
    }

    // 1) Tipo y tamaño
    if (!ACCEPTED_EXT.test(file.name)){
      showAlert('Tipo de archivo no permitido. Solo .xlsx o .xls.', 'danger');
      disableAll();
      return;
    }
    const sizeMB = file.size / (1024*1024);
    if (sizeMB > MAX_FILE_SIZE_MB){
      showAlert(`El archivo supera el máximo (${sizeMB.toFixed(1)} MB > ${MAX_FILE_SIZE_MB} MB).`, 'danger');
      disableAll();
      return;
    }

    // 2) Leer buffer
    let arrayBuffer = null;
    try{
      arrayBuffer = await file.arrayBuffer();
    }catch{
      showAlert('No se pudo leer el archivo (permiso/corrupción).', 'danger');
      disableAll();
      return;
    }

    // 3) Abrir workbook c/SheetJS
    let wb;
    try{
      wb = XLSX.read(new Uint8Array(arrayBuffer), { type:'array' });
    }catch(err){
      showAlert('Archivo dañado, protegido o formato no válido.', 'danger');
      disableAll();
      return;
    }
    if (!wb.SheetNames?.length){
      showAlert('El archivo no contiene hojas.', 'danger');
      disableAll();
      return;
    }

    // 4) Advertencia por muchas hojas
    const warnings = [];
    if (wb.SheetNames.length > MAX_SHEETS_WARN){
      warnings.push(`El libro tiene <b>${wb.SheetNames.length}</b> hojas (podría demorar).`);
    }

    // 5) Validación por hoja (estructura y datos)
    const report = [];
    const okSheets = [];
    let totalValidPoints = 0;

    for (const sheetName of wb.SheetNames){
      const ws = wb.Sheets[sheetName];
      const sheetRes = validateSheet(ws, sheetName);
      report.push(sheetRes);
      if (sheetRes.validPoints > 0){
        okSheets.push(sheetName);
        totalValidPoints += sheetRes.validPoints;
      }
    }

    if (!okSheets.length){
      const msg = [
        'No hay hojas con datos suficientes en B (Y) y C (X).',
        'Revisa que existan valores numéricos en B y horas válidas en C (hh:mm, hh:mm:ss, AM/PM o serial de Excel).'
      ].join(' ');
      showAlert(msg, 'danger');
      disableAll();
      return;
    }

    if (totalValidPoints > MAX_POINTS_WARN){
      warnings.push(`Se detectaron <b>${totalValidPoints}</b> puntos válidos en total (podría afectar el rendimiento).`);
    }

    // 6) Mostrar resumen amigable
    renderValidationSummary(report, warnings);

    // 7) Guardar estado validado y habilitar UI
    window.__validated = {
      arrayBuffer,
      fileName: file.name,
      sheetsReport: report,
      okSheets,
      totalValidPoints
    };

    $btnProcess.disabled = false;
    $btnClear.disabled = false;
    // El botón "Descargar" queda deshabilitado hasta que la app genere gráficos.
    // Emitimos un evento para que app.js se entere:
    document.dispatchEvent(new CustomEvent('validation:passed', {
      detail: {
        fileName: file.name,
        arrayBuffer,
        okSheets
      }
    }));
  }

  // ====== Validación de una hoja ======
  function validateSheet(ws, sheetName){
    const res = {
      sheetName,
      hasRef: !!ws && !!ws['!ref'],
      totalRows: 0,          // datos (excluyendo header)
      validPoints: 0,
      invalidB: 0,
      invalidC: 0,
      bothInvalid: 0,
      timeDuplicates: 0,     // cantidad de hh:mm duplicadas (no agregadas)
      notes: []
    };
    if (!res.hasRef){
      res.notes.push('Hoja vacía o sin rango.');
      return res;
    }

    const range = XLSX.utils.decode_range(ws['!ref']);
    const seenTimes = new Map(); // hh:mm -> count
    for (let r = range.s.r + 1; r <= range.e.r; r++){
      const cellB = ws[XLSX.utils.encode_cell({ r, c: 1 })]; // B
      const cellC = ws[XLSX.utils.encode_cell({ r, c: 2 })]; // C

      const numY = toNumber(cellB?.v);
      const t    = toHHMM(cellC?.v);

      // Contabilizamos fila si hay al menos un valor no vacío en B o C
      const hasSomething = (cellB?.v !== undefined) || (cellC?.v !== undefined);
      if (hasSomething) res.totalRows++;

      const okB = (numY !== null) && Number.isFinite(numY);
      const okC = !!t;

      if (okB && okC){
        res.validPoints++;
        const cnt = (seenTimes.get(t.hhmm) || 0) + 1;
        seenTimes.set(t.hhmm, cnt);
        if (cnt > 1) res.timeDuplicates++;
      } else if (!okB && okC){
        res.invalidB++;
      } else if (okB && !okC){
        res.invalidC++;
      } else if (hasSomething){
        res.bothInvalid++;
      }
    }

    if (res.validPoints === 0){
      res.notes.push('Sin pares válidos B↔C.');
    }
    if (res.timeDuplicates > 0){
      res.notes.push(`Se encontraron ${res.timeDuplicates} duplicados de hora (se mantendrá orden estable).`);
    }
    return res;
  }

  // ====== Render de resumen de validación ======
  function renderValidationSummary(report, warnings){
    const lines = [];

    if (warnings.length){
      lines.push(`<div class="alert alert-warning mb-2"><ul class="mb-0">${warnings.map(w=>`<li>${w}</li>`).join('')}</ul></div>`);
    }

    lines.push('<div class="table-responsive"><table class="table table-sm align-middle mb-0">');
    lines.push('<thead><tr><th>Hoja</th><th class="text-end">Filas</th><th class="text-end text-success">Válidos</th><th class="text-end">B invál.</th><th class="text-end">C invál.</th><th class="text-end">Ambos invál.</th><th class="text-end">Duplic. hora</th><th>Notas</th></tr></thead><tbody>');

    for (const s of report){
      lines.push(
        `<tr>
          <td>${escapeHtml(s.sheetName)}</td>
          <td class="text-end">${s.totalRows}</td>
          <td class="text-end text-success">${s.validPoints}</td>
          <td class="text-end">${s.invalidB}</td>
          <td class="text-end">${s.invalidC}</td>
          <td class="text-end">${s.bothInvalid}</td>
          <td class="text-end">${s.timeDuplicates}</td>
          <td>${s.notes.map(escapeHtml).join(' · ')}</td>
        </tr>`
      );
    }
    lines.push('</tbody></table></div>');

    showAlert(lines.join(''), 'info');
  }

  // ====== Helpers de parseo (idénticos a la app para coherencia) ======
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
    const RX_HHMM = /(\d{1,2}):(\d{2})(?::\d{2})?/;

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
      // Política estricta en validación: si no parece hora, la marcamos como inválida (no +∞)
      return null;
    }
    return null;
  }

  // ====== Utilidades UI ======
  function showAlert(html, type='info'){ $alert.innerHTML = `<div class="alert alert-${type}" role="alert">${html}</div>`; }
  function showInfo(text){ $info.textContent = text; }
  function disableAll(){
    $btnProcess.disabled = true;
    $btnEmbed.disabled   = true;
    $btnClear.disabled   = !$file.files?.length;
  }
  function resetState(){
    window.__validated = null;
    $btnProcess.disabled = true;
    $btnEmbed.disabled   = true;
    $btnClear.disabled   = true;
    $alert.innerHTML = '';
  }
  function escapeHtml(s){
    return String(s).replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
  }
})();
