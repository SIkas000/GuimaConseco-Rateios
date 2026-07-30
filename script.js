let pivotHeaders = [], pivotRows = [];
let linhasExport = [];

function showAlert(msg,type){ const e=document.getElementById('alertBox'); e.className='alert '+type; e.innerHTML=msg; e.style.display='flex'; }
function hideAlert(){ document.getElementById('alertBox').style.display='none'; }
function setProgress(p){ const b=document.getElementById('progressBar'); b.style.display='block'; document.getElementById('progressFill').style.width=p+'%'; if(p>=100) setTimeout(()=>b.style.display='none',700); }
function tick(){ return new Promise(r=>setTimeout(r,30)); }

function parseBRNumber(v){
  if(typeof v === 'number') return v;
  if(v===null || v===undefined) return NaN;
  let s = String(v).trim();
  s = s.replace(/[^\d,.\-]/g,'');
  if(s==='') return NaN;
  if(s.includes(',')){ s = s.replace(/\./g,'').replace(',', '.'); }
  return parseFloat(s);
}
function fmtBR(n){ return n.toLocaleString('pt-BR', {minimumFractionDigits:2, maximumFractionDigits:2}); }
function fmtMoney(n){ return 'R$ ' + fmtBR(n); }

function bestMatch(headers, keywords){
  return headers.find(h => keywords.some(k => h.toLowerCase().includes(k))) || headers[0];
}

// Procura a aba cujo nome seja "DI" (qualquer capitalização).
// Se não encontrar, usa a primeira aba e avisa o usuário.
function findDiSheet(wb) {
  const diName = wb.SheetNames.find(n => n.trim().toLowerCase() === 'di');
  if (diName) return { sheetName: diName, found: true };
  return { sheetName: wb.SheetNames[0], found: false };
}

function loadPivot(input){
  const file = input.files[0];
  if(!file) return;
  const ext = file.name.split('.').pop().toLowerCase();
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      let wb;
      if(ext==='csv') wb = XLSX.read(e.target.result,{type:'string'});
      else wb = XLSX.read(e.target.result,{type:'array'});

      // ── AJUSTE: procura aba "DI" (case-insensitive) ──
      const { sheetName, found } = findDiSheet(wb);
      if (!found) {
        showAlert(`⚠️ Aba "DI" não encontrada. Usando a primeira aba disponível: "${sheetName}".`, 'warn');
      }
      const ws = wb.Sheets[sheetName];
      // ─────────────────────────────────────────────────

      const json = XLSX.utils.sheet_to_json(ws,{defval:''});
      if(!json.length){ showAlert('⚠️ Arquivo vazio ou sem dados reconhecíveis.','error'); return; }
      pivotRows = json;
      pivotHeaders = Object.keys(json[0]);
      document.getElementById('pivotCard').classList.add('loaded');
      document.getElementById('pivotInfo').textContent = `✓ ${file.name} — aba "${sheetName}" — ${json.length.toLocaleString('pt-BR')} linhas`;
      document.getElementById('pivotIcon').textContent = '✅';
      populatePivotSelects();
      document.getElementById('step1').classList.replace('active','done');
      document.getElementById('step2').classList.add('active');
      if(found) hideAlert();
    } catch(err){ showAlert('⚠️ Erro ao ler o arquivo. Verifique se é um Excel ou CSV válido.','error'); }
  };
  if(ext==='csv') reader.readAsText(file,'UTF-8');
  else reader.readAsArrayBuffer(file);
}

function populatePivotSelects(){
  const ccSel = document.getElementById('pivotCcCol');
  const valSel = document.getElementById('pivotValCol');
  ccSel.innerHTML = pivotHeaders.map(h=>`<option value="${h}">${h}</option>`).join('');
  valSel.innerHTML = pivotHeaders.map(h=>`<option value="${h}">${h}</option>`).join('');
  const ccGuess = bestMatch(pivotHeaders, ['centro de custo','centro_custo','c. custo','c.custo','cc','custo','código','cod']);
  const valGuess = bestMatch(pivotHeaders, ['valor','frete','cobrado','total']);
  ccSel.value = ccGuess;
  valSel.value = (valGuess && valGuess!==ccGuess) ? valGuess : (pivotHeaders.find(h=>h!==ccGuess) || pivotHeaders[0]);
  document.getElementById('pivotColsRow').style.display = 'grid';
}

function getPivotData(){
  const ccCol = document.getElementById('pivotCcCol').value;
  const valCol = document.getElementById('pivotValCol').value;
  const list = [];
  pivotRows.forEach(row => {
    const cc = String(row[ccCol] ?? '').trim();
    if(!cc) return;
    // Ignora linhas de totais/subtotais (ex: "Total Geral", "Total", "(em branco)" etc.)
    if(/total|subtotal|grand|em branco/i.test(cc)) return;
    const val = parseBRNumber(row[valCol]);
    if(isNaN(val)) return;
    list.push({cc, valor: val});
  });
  return list;
}

function totalFreteAtual(){
  return getPivotData().reduce((a,b)=>a+b.valor,0);
}

function calcularPisCofins(){
  if(!pivotRows.length){ showAlert('⚠️ Envie a tabela dinâmica antes de calcular o PIS/COFINS.','warn'); return; }
  const total = totalFreteAtual();
  if(!total){ showAlert('⚠️ Não consegui somar a tabela dinâmica. Confira as colunas selecionadas.','error'); return; }
  const pisPct = parseBRNumber(document.getElementById('pisPct').value)/100;
  const cofinsPct = parseBRNumber(document.getElementById('cofinsPct').value)/100;
  document.getElementById('pisValor').value = fmtBR(round2(total*pisPct));
  document.getElementById('cofinsValor').value = fmtBR(round2(total*cofinsPct));
  hideAlert();
}

function round2(n){ return Math.round((n + Number.EPSILON)*100)/100; }

async function processData(){
  hideAlert();
  if(!pivotRows.length){ showAlert('⚠️ Envie a tabela dinâmica com os valores por Centro de Custo.','error'); return; }

  const dataLancRaw = document.getElementById('dataLanc').value;
  const nfNum = document.getElementById('nfNum').value.trim();
  const fornecedor = document.getElementById('fornecedor').value.trim();
  const ccPai = document.getElementById('ccPai').value.trim();
  const contaFrete = document.getElementById('contaFrete').value.trim();
  const contaPis = document.getElementById('contaPis').value.trim();
  const contaCofins = document.getElementById('contaCofins').value.trim();
  const contaTransitoria = document.getElementById('contaTransitoria').value.trim();
  const codigoPosicao = document.getElementById('codigoPosicao').value.trim();
  const totalPis = parseBRNumber(document.getElementById('pisValor').value);
  const totalCofins = parseBRNumber(document.getElementById('cofinsValor').value);

  if(!dataLancRaw || !nfNum || !fornecedor){ showAlert('⚠️ Preencha data, número da NF e fornecedor.','error'); return; }
  if(isNaN(totalPis) || isNaN(totalCofins)){ showAlert('⚠️ Informe o valor do PIS e do COFINS (ou use o botão de cálculo automático).','error'); return; }

  const pivotData = getPivotData();
  if(!pivotData.length){ showAlert('⚠️ Não encontrei linhas válidas na tabela dinâmica. Confira as colunas selecionadas.','error'); return; }

  document.getElementById('btnProcess').disabled = true;
  document.getElementById('btnProcess').textContent = '⏳ Processando…';
  setProgress(20);
  await tick();

  const [y,m,d] = dataLancRaw.split('-').map(Number);
  const dataLanc = new Date(y, m-1, d);
  const totalFrete = round2(pivotData.reduce((a,b)=>a+b.valor,0));

  const histNf = `Recl. NF. ${nfNum} de ${fornecedor}`;
  const histPis = `Recl. PIS S/ NF. ${nfNum} de ${fornecedor}`;
  const histCofins = `Recl. COFINS S/ NF. ${nfNum} de ${fornecedor}`;

  const pcts = pivotData.map(r => round2(r.valor/totalFrete*100));
  let idxMenor = 0;
  pivotData.forEach((r,i)=>{ if(r.valor < pivotData[idxMenor].valor) idxMenor = i; });
  pcts[idxMenor] = round2(pcts[idxMenor] + round2(100 - pcts.reduce((a,b)=>a+b,0)));

  const pisPorCc = pcts.map(p => round2(p*totalPis/100));
  pisPorCc[idxMenor] = round2(pisPorCc[idxMenor] + round2(totalPis - pisPorCc.reduce((a,b)=>a+b,0)));

  const cofinsPorCc = pcts.map(p => round2(p*totalCofins/100));
  cofinsPorCc[idxMenor] = round2(cofinsPorCc[idxMenor] + round2(totalCofins - cofinsPorCc.reduce((a,b)=>a+b,0)));

  setProgress(55);
  await tick();

  const linhas = [], preview = [];

  function addHeader(contaA, contaB, valor, hist){
    linhas.push([1, {__type:'date', v: dataLanc.getFullYear()+'-'+String(dataLanc.getMonth()+1).padStart(2,'0')+'-'+String(dataLanc.getDate()).padStart(2,'0')}, Number(contaA), Number(contaB), valor, hist, Number(codigoPosicao)]);
    preview.push({tipo:'header', data:dataLanc, contaA, contaB, valor, extra:hist});
  }
  function addDetail(cc, pct, valor, dc){
    linhas.push([2, Number(contaTransitoria), Number(cc), pct, valor, dc, null]);
    preview.push({tipo:'detail', contaA:contaTransitoria, contaB:cc, pct, valor, dc});
  }

  addHeader(contaFrete, contaTransitoria, totalFrete, histNf);
  addDetail(ccPai, 100, totalFrete, 'C');

  addHeader(contaTransitoria, contaPis, totalPis, histPis);
  addDetail(ccPai, 100, totalPis, 'C');

  addHeader(contaTransitoria, contaCofins, totalCofins, histCofins);
  addDetail(ccPai, 100, totalCofins, 'C');

  addHeader(contaTransitoria, contaFrete, totalFrete, histNf);
  pivotData.forEach((r,i)=> addDetail(r.cc, pcts[i], r.valor, 'D'));

  addHeader(contaPis, contaTransitoria, totalPis, histPis);
  pivotData.forEach((r,i)=> addDetail(r.cc, pcts[i], pisPorCc[i], 'C'));

  addHeader(contaCofins, contaTransitoria, totalCofins, histCofins);
  pivotData.forEach((r,i)=> addDetail(r.cc, pcts[i], cofinsPorCc[i], 'C'));

  linhasExport = linhas;
  setProgress(85);
  await tick();

  document.getElementById('statCcs').textContent = pivotData.length.toLocaleString('pt-BR');
  document.getElementById('statFrete').textContent = fmtMoney(totalFrete);
  document.getElementById('statPis').textContent = fmtMoney(totalPis);
  document.getElementById('statCofins').textContent = fmtMoney(totalCofins);

  const body = document.getElementById('previewBody');
  const dataFmt = dataLanc.toLocaleDateString('pt-BR');
  body.innerHTML = preview.map(p=>{
    if(p.tipo==='header'){
      return `<tr class="header-row"><td>1</td><td>${dataFmt}</td><td>${p.contaA}</td><td>${p.contaB}</td><td>${fmtBR(p.valor)}</td><td>${p.extra}</td><td></td></tr>`;
    } else {
      const tag = p.dc==='D' ? `<span class="tag-D">D</span>` : `<span class="tag-C">C</span>`;
      return `<tr><td>2</td><td></td><td>${p.contaA}</td><td>${p.contaB}</td><td>${fmtBR(p.valor)} <span style="color:#8aa88a">(${fmtBR(p.pct)}%)</span></td><td>${tag}</td><td></td></tr>`;
    }
  }).join('');

  document.getElementById('previewNote').textContent = `${linhas.length} linhas geradas (6 blocos) para ${pivotData.length} centros de custo. Diferença de arredondamento absorvida pelo CC de menor valor (${pivotData[idxMenor].cc}).`;

  setProgress(100);
  document.getElementById('resultCard').classList.add('visible');
  document.getElementById('step2').classList.replace('active','done');
  document.getElementById('step3').classList.add('active');
  document.getElementById('btnProcess').disabled = false;
  document.getElementById('btnProcess').textContent = '⚙️ Reprocessar';

  showAlert(`✅ Lançamento calculado para ${pivotData.length} CCs. Confira o preview e baixe o Excel.`,'info');
  document.getElementById('resultCard').scrollIntoView({behavior:'smooth'});
}

function crc32(buf){
  // CRC32 table
  if(!crc32.table){
    crc32.table = new Uint32Array(256);
    for(let i=0;i<256;i++){
      let c=i;
      for(let j=0;j<8;j++) c = (c&1) ? (0xEDB88320^(c>>>1)) : (c>>>1);
      crc32.table[i]=c;
    }
  }
  let crc=0xFFFFFFFF;
  for(let i=0;i<buf.length;i++) crc = crc32.table[(crc^buf[i])&0xFF]^(crc>>>8);
  return (crc^0xFFFFFFFF)>>>0;
}

function buildXlsx(linhasExport, sheetName){
  // Converte data {__type:'date', v:'YYYY-MM-DD'} para serial Excel
  function dateToSerial(v){
    const [y,m,d] = v.split('-').map(Number);
    // Base: 30 de Dezembro de 1899 (como o Excel conta)
    const base   = Date.UTC(1899,11,30);
    const target = Date.UTC(y, m-1, d);
    return Math.round((target - base) / 86400000);
  }

  const enc = new TextEncoder();

  // ── Styles XML com Aptos Narrow 11, date format DD/MM/YYYY ──
  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="DD/MM/YYYY"/></numFmts><fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Aptos Narrow"/><family val="2"/><scheme val="minor"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;

  // ── Sheet XML ──
  const cols = ['A','B','C','D','E','F','G'];
  let sheetRows = '';
  linhasExport.forEach((row, ri) => {
    let cells = '';
    row.forEach((val, ci) => {
      const addr = cols[ci] + (ri+1);
      if(val === null || val === undefined){
        // célula vazia, não emite
      } else if(typeof val === 'object' && val.__type === 'date'){
        const serial = dateToSerial(val.v);
        cells += `<c r="${addr}" s="1" t="n"><v>${serial}</v></c>`;
      } else if(typeof val === 'number'){
        cells += `<c r="${addr}" t="n"><v>${val}</v></c>`;
      } else {
        const escaped = String(val).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');
        cells += `<c r="${addr}" t="inlineStr"><is><t>${escaped}</t></is></c>`;
      }
    });
    sheetRows += `<row r="${ri+1}">${cells}</row>`;
  });
  const lastRow = linhasExport.length;
  const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${sheetRows}</sheetData></worksheet>`;

  // ── Workbook XML ──
  const safeSheet = sheetName.replace(/&/g,'&amp;').replace(/'/g,'&apos;');
  const wbXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="${safeSheet}" sheetId="1" r:id="rId1"/></sheets></workbook>`;

  const wbRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;

  const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;

  const ctXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`;

  // ── Monta ZIP ──
  const entries = [
    { name: '[Content_Types].xml',         data: enc.encode(ctXml) },
    { name: '_rels/.rels',                 data: enc.encode(relsXml) },
    { name: 'xl/workbook.xml',             data: enc.encode(wbXml) },
    { name: 'xl/_rels/workbook.xml.rels',  data: enc.encode(wbRelsXml) },
    { name: 'xl/styles.xml',              data: enc.encode(stylesXml) },
    { name: 'xl/worksheets/sheet1.xml',   data: enc.encode(sheetXml) },
  ];

  // Data fixa para reproducibilidade (2026-01-01 00:00:00)
  const dosDate = (2026-1980)<<9 | 1<<5 | 1; // 0x2C21
  const dosTime = 0;

  const localHeaders = [];
  const centralDir   = [];
  let offset = 0;

  entries.forEach(({name, data}) => {
    const nameBytes = enc.encode(name);
    const crc   = crc32(data);
    const size  = data.length;

    // Local file header
    const lfh = new Uint8Array(30 + nameBytes.length);
    const v = new DataView(lfh.buffer);
    v.setUint32(0,  0x04034B50, true); // signature
    v.setUint16(4,  20, true);          // version needed
    v.setUint16(6,  0,  true);          // flags
    v.setUint16(8,  0,  true);          // compression: stored
    v.setUint16(10, dosTime, true);
    v.setUint16(12, dosDate, true);
    v.setUint32(14, crc,  true);
    v.setUint32(18, size, true);        // compressed size
    v.setUint32(22, size, true);        // uncompressed size
    v.setUint16(26, nameBytes.length, true);
    v.setUint16(28, 0, true);           // extra length
    lfh.set(nameBytes, 30);

    localHeaders.push(lfh);
    localHeaders.push(data);

    // Central directory header
    const cdh = new Uint8Array(46 + nameBytes.length);
    const cv = new DataView(cdh.buffer);
    cv.setUint32(0,  0x02014B50, true);
    cv.setUint16(4,  20, true);
    cv.setUint16(6,  20, true);
    cv.setUint16(8,  0,  true);
    cv.setUint16(10, 0,  true);
    cv.setUint16(12, dosTime, true);
    cv.setUint16(14, dosDate, true);
    cv.setUint32(16, crc,  true);
    cv.setUint32(20, size, true);
    cv.setUint32(24, size, true);
    cv.setUint16(28, nameBytes.length, true);
    cv.setUint16(30, 0, true);
    cv.setUint16(32, 0, true);
    cv.setUint16(34, 0, true);
    cv.setUint16(36, 0, true);
    cv.setUint32(38, 0, true);
    cv.setUint32(42, offset, true);
    cdh.set(nameBytes, 46);

    centralDir.push(cdh);
    offset += lfh.length + data.length;
  });

  const cdOffset = offset;
  const cdSize   = centralDir.reduce((a,b) => a+b.length, 0);

  // End of central directory
  const eocd = new Uint8Array(22);
  const ev = new DataView(eocd.buffer);
  ev.setUint32(0, 0x06054B50, true);
  ev.setUint16(4, 0, true);
  ev.setUint16(6, 0, true);
  ev.setUint16(8, entries.length, true);
  ev.setUint16(10, entries.length, true);
  ev.setUint32(12, cdSize, true);
  ev.setUint32(16, cdOffset, true);
  ev.setUint16(20, 0, true);

  // Concatena tudo
  const totalSize = offset + cdSize + eocd.length;
  const zip = new Uint8Array(totalSize);
  let pos = 0;
  [...localHeaders, ...centralDir, [eocd]].flat().forEach(chunk => {
    zip.set(chunk, pos); pos += chunk.length;
  });

  return zip;
}

function downloadResult(){
  if(!linhasExport.length){ showAlert('⚠️ Calcule o lançamento antes de baixar.','error'); return; }

  const nfNum      = document.getElementById('nfNum').value.trim();
  const fornecedor = document.getElementById('fornecedor').value.trim();
  const sheetName  = ('Lançamento reclassificação ' + fornecedor).substring(0, 31);
  const filename   = 'Lancamento_reclassificacao_' + fornecedor.replace(/[^\w]+/g,'_') + '_NF_' + nfNum + '.xlsx';

  try {
    const zip  = buildXlsx(linhasExport, sheetName);
    const blob = new Blob([zip], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    const a = Object.assign(document.createElement('a'),{ href:URL.createObjectURL(blob), download:filename });
    a.click();
    URL.revokeObjectURL(a.href);
  } catch(err){
    console.error(err);
    showAlert('⚠️ Erro ao gerar o arquivo: ' + err.message, 'error');
  }
}

function resetAll(){
  pivotHeaders=[]; pivotRows=[]; linhasExport=[];
  document.getElementById('pivotCard').classList.remove('loaded');
  document.getElementById('pivotInfo').style.display='none';
  document.getElementById('pivotIcon').textContent = '📊';
  document.getElementById('pivotFile').value = '';
  document.getElementById('pivotColsRow').style.display = 'none';
  document.getElementById('dataLanc').value='';
  document.getElementById('nfNum').value='';
  document.getElementById('fornecedor').value='';
  document.getElementById('ccPai').value='1100112';
  document.getElementById('pisPct').value='1,65';
  document.getElementById('cofinsPct').value='7,60';
  document.getElementById('pisValor').value='';
  document.getElementById('cofinsValor').value='';
  document.getElementById('contaFrete').value='11905';
  document.getElementById('contaPis').value='11813';
  document.getElementById('contaCofins').value='11814';
  document.getElementById('contaTransitoria').value='51112';
  document.getElementById('codigoPosicao').value='8';
  document.getElementById('resultCard').classList.remove('visible');
  hideAlert();
  ['step1','step2','step3'].forEach(id=>{ document.getElementById(id).classList.remove('active','done'); });
  document.getElementById('step1').classList.add('active');
  document.getElementById('btnProcess').textContent = '⚙️ Calcular e montar lançamento';
}