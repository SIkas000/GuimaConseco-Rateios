let data1=[], data2=[], headers1=[], headers2=[], resultData=[];

function showAlert(msg,type){ const e=document.getElementById('alertBox'); e.className='alert '+type; e.innerHTML=msg; e.style.display='flex'; }
function hideAlert(){ document.getElementById('alertBox').style.display='none'; }
function setProgress(p){ const b=document.getElementById('progressBar'); b.style.display='block'; document.getElementById('progressFill').style.width=p+'%'; if(p>=100) setTimeout(()=>b.style.display='none',700); }

function loadFile(input, num){
  const file = input.files[0];
  if(!file) return;
  const ext = file.name.split('.').pop().toLowerCase();
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      let wb;
      if(ext==='csv') wb = XLSX.read(e.target.result,{type:'string'});
      else wb = XLSX.read(e.target.result,{type:'array'});
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(ws,{defval:''});
      if(!json.length){ showAlert('⚠️ Arquivo vazio ou sem dados reconhecíveis.','error'); return; }
      const headers = Object.keys(json[0]);
      if(num===1){ data1=json; headers1=headers; }
      else { data2=json; headers2=headers; }
      document.getElementById('card'+num).classList.add('loaded');
      document.getElementById('info'+num).textContent = `✓ ${file.name}  —  ${json.length.toLocaleString('pt-BR')} linhas`;
      document.getElementById('icon'+num).textContent = '✅';
      hideAlert();
      checkBothLoaded();
    } catch(err){ showAlert('⚠️ Erro ao ler o arquivo. Verifique se é um Excel ou CSV válido.','error'); }
  };
  if(ext==='csv') reader.readAsText(file,'UTF-8');
  else reader.readAsArrayBuffer(file);
}

function checkBothLoaded(){
  if(!data1.length || !data2.length) return;
  populateSelects();
  document.getElementById('configCard').classList.add('visible');
  document.getElementById('processRow').style.display='flex';
  document.getElementById('step1').classList.replace('active','done');
  document.getElementById('step2').classList.add('active');
}

function bestMatch(headers, keywords){
  return headers.find(h => keywords.some(k => h.toLowerCase().includes(k))) || headers[0];
}

function populateSelects(){
  const fill = (id, opts, def) => {
    const s = document.getElementById(id);
    s.innerHTML = opts.map(h=>`<option value="${h}">${h}</option>`).join('');
    if(def && opts.includes(def)) s.value = def;
    else { const m = bestMatch(opts, typeof def === 'object' ? def : [def]); if(m) s.value=m; }
  };
  // Planilha 1: chave = "centro de custo" (col J normalmente)
  fill('col1Key', headers1, bestMatch(headers1,['centro de custo','centro_custo','centro custo','centrodecusto']));
  // Planilha 1: destino = "descrição centro"
  fill('col1Dest', headers1, bestMatch(headers1,['descrição centro','descricao centro','desc centro','descrição c']));
  // Planilha 2: referência = col A (C. Custo)
  fill('col2Key', headers2, bestMatch(headers2,['c. custo','c.custo','custo','código','cod','key']));
  // Planilha 2: valor a trazer = Texto C.Custo (col E)
  fill('col2Val', headers2, bestMatch(headers2,['texto','text c','texto c','textocc']));
}

async function processData(){
  hideAlert();
  const c1Key  = document.getElementById('col1Key').value;
  const c1Dest = document.getElementById('col1Dest').value;
  const c2Key  = document.getElementById('col2Key').value;
  const c2Val  = document.getElementById('col2Val').value;

  document.getElementById('btnProcess').disabled=true;
  document.getElementById('btnProcess').textContent='⏳ Processando…';
  setProgress(15);
  await tick();

  // Monta lookup da planilha 2: c2Key → c2Val
  const lookup = new Map();
  data2.forEach(row => {
    const k = String(row[c2Key]??'').trim();
    if(k) lookup.set(k, row[c2Val]??'');
  });

  setProgress(45);
  await tick();

  let found=0, notFound=0;
  resultData = data1.map(row => {
    const keyVal = String(row[c1Key]??'').trim();
    const match  = lookup.has(keyVal) ? lookup.get(keyVal) : null;
    if(match!==null) found++; else notFound++;
    const out = {...row};
    out[c1Dest] = match !== null ? match : (row[c1Dest]??'');
    out['__status'] = match !== null ? 'ok' : 'nao_encontrado';
    return out;
  });

  setProgress(80);
  await tick();

  const total = resultData.length;
  document.getElementById('statTotal').textContent = total.toLocaleString('pt-BR');
  document.getElementById('statOk').textContent    = found.toLocaleString('pt-BR');
  document.getElementById('statErr').textContent   = notFound.toLocaleString('pt-BR');
  document.getElementById('statPct').textContent   = Math.round(found/total*100)+'%';
  document.getElementById('btnErrors').style.display = notFound > 0 ? 'flex' : 'none';

  // Preview — mostrar só colunas relevantes: col1Key, col1Dest, status + mais 3 primeiras
  const otherCols = headers1.filter(h => h!==c1Key && h!==c1Dest).slice(0,4);
  const previewCols = [c1Key, c1Dest, ...otherCols];

  const head = document.getElementById('previewHead');
  const body = document.getElementById('previewBody');
  head.innerHTML = '<tr>'
    + previewCols.map(c => `<th class="${c===c1Dest?'highlight':''}">${c}</th>`).join('')
    + '<th>Status</th></tr>';

  const preview = resultData.slice(0, 150);
  body.innerHTML = preview.map(row => {
    const ok = row['__status']==='ok';
    const cells = previewCols.map(c =>
      `<td class="${c===c1Dest?'highlight':''}">${row[c]??''}</td>`
    ).join('');
    const tag = ok
      ? `<span class="tag-ok">✓ ok</span>`
      : `<span class="tag-err">⚠ não encontrado</span>`;
    return `<tr>${cells}<td>${tag}</td></tr>`;
  }).join('');

  const note = total > 150 ? `Mostrando 150 de ${total.toLocaleString('pt-BR')} linhas. O arquivo exportado conterá todas as linhas.` : `${total.toLocaleString('pt-BR')} linhas processadas.`;
  document.getElementById('previewNote').textContent = note;

  setProgress(100);
  document.getElementById('resultCard').classList.add('visible');
  document.getElementById('step2').classList.replace('active','done');
  document.getElementById('step3').classList.add('active');
  document.getElementById('btnProcess').disabled=false;
  document.getElementById('btnProcess').textContent='⚙️ Reprocessar';

  if(notFound>0){
    showAlert(`⚠️ ${notFound.toLocaleString('pt-BR')} linhas sem correspondência. Baixe o arquivo de erros para verificar os códigos não encontrados.`,'warn');
  } else {
    showAlert(`✅ Todos os ${total.toLocaleString('pt-BR')} registros foram preenchidos com sucesso!`,'info');
  }
  document.getElementById('resultCard').scrollIntoView({behavior:'smooth'});
}

function buildExport(rows){
  const c1Dest = document.getElementById('col1Dest').value;
  // Exporta todas as colunas originais (na ordem original), com col1Dest atualizado, sem __status
  return rows.map(row => {
    const out = {};
    headers1.forEach(h => out[h] = h===c1Dest ? (row[h]??'') : (row[h]??''));
    return out;
  });
}

function downloadResult(fmt){
  const rows = buildExport(resultData);
  const ws = XLSX.utils.json_to_sheet(rows, {header: headers1});
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Resultado');
  if(fmt==='csv'){
    const csv = XLSX.utils.sheet_to_csv(ws);
    const blob = new Blob(['\uFEFF'+csv],{type:'text/csv;charset=utf-8;'});
    const a = Object.assign(document.createElement('a'),{href:URL.createObjectURL(blob),download:'centro_custo_resultado.csv'});
    a.click(); URL.revokeObjectURL(a.href);
  } else {
    XLSX.writeFile(wb,'centro_custo_resultado.xlsx');
  }
}

function downloadErrors(){
  const erros = resultData.filter(r=>r['__status']==='nao_encontrado');
  if(!erros.length){ showAlert('ℹ️ Nenhum erro para exportar.','info'); return; }
  const rows = buildExport(erros);
  const ws = XLSX.utils.json_to_sheet(rows,{header:headers1});
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Erros');
  XLSX.writeFile(wb,'centro_custo_erros.xlsx');
}

function resetAll(){
  data1=[]; data2=[]; headers1=[]; headers2=[]; resultData=[];
  [1,2].forEach(n=>{
    document.getElementById('card'+n).classList.remove('loaded');
    document.getElementById('info'+n).style.display='none';
    document.getElementById('icon'+n).textContent = n===1?'📋':'🗂️';
    document.getElementById('file'+n).value='';
  });
  document.getElementById('configCard').classList.remove('visible');
  document.getElementById('processRow').style.display='none';
  document.getElementById('resultCard').classList.remove('visible');
  hideAlert();
  ['step1','step2','step3'].forEach(id=>{
    document.getElementById(id).classList.remove('active','done');
  });
  document.getElementById('step1').classList.add('active');
  document.getElementById('btnProcess').textContent='⚙️ Processar cruzamento';
}

function tick(){ return new Promise(r=>setTimeout(r,30)); }