const DB_NAME='controle_medidores_db', DB_VERSION=1;
let db, records=[], imports=[], chartDay, chartPhase;

function esc(v){return String(v??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[m]))}
function norm(v){return String(v??'').trim()}
function normalizeHeader(v){return norm(v).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/\s+/g,' ')}
function fmtDateBR(v){
  let s=norm(v); if(!s) return '';
  if(/^\d{2}\/\d{2}\/\d{2}$/.test(s)){let [d,m,y]=s.split('/'); return `${d}/${m}/20${y}`}
  if(/^\d{2}\/\d{2}\/\d{4}/.test(s)) return s.slice(0,10);
  if(/^\d{4}-\d{2}-\d{2}/.test(s)){let [y,m,d]=s.slice(0,10).split('-');return `${d}/${m}/${y}`}
  return s;
}
function dateKey(v){let s=fmtDateBR(v),p=s.split('/');return p.length===3?`${p[2]}-${p[1]}-${p[0]}`:s}
function download(name,text,type='application/json'){let a=document.createElement('a');a.href=URL.createObjectURL(new Blob([text],{type}));a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000)}
function hashString(s){let h=2166136261;for(let i=0;i<s.length;i++){h^=s.charCodeAt(i);h=Math.imul(h,16777619)}return (h>>>0).toString(36)}

function openDB(){return new Promise((resolve,reject)=>{let r=indexedDB.open(DB_NAME,DB_VERSION);
r.onupgradeneeded=e=>{let d=e.target.result;if(!d.objectStoreNames.contains('records'))d.createObjectStore('records',{keyPath:'uid'});if(!d.objectStoreNames.contains('imports'))d.createObjectStore('imports',{keyPath:'id',autoIncrement:true})};
r.onsuccess=e=>{db=e.target.result;resolve()};r.onerror=()=>reject(r.error)})}
function storeGetAll(name){return new Promise((res,rej)=>{let r=db.transaction(name,'readonly').objectStore(name).getAll();r.onsuccess=()=>res(r.result);r.onerror=()=>rej(r.error)})}
function storePutMany(name,items){return new Promise((res,rej)=>{let tx=db.transaction(name,'readwrite'),s=tx.objectStore(name);items.forEach(x=>s.put(x));tx.oncomplete=res;tx.onerror=()=>rej(tx.error)})}
function storeAdd(name,item){return new Promise((res,rej)=>{let r=db.transaction(name,'readwrite').objectStore(name).add(item);r.onsuccess=()=>res(r.result);r.onerror=()=>rej(r.error)})}
function storeDelete(name,key){return new Promise((res,rej)=>{let r=db.transaction(name,'readwrite').objectStore(name).delete(key);r.onsuccess=res;r.onerror=()=>rej(r.error)})}
function deleteRecordsWhere(predicate){return new Promise((res,rej)=>{let tx=db.transaction('records','readwrite'),st=tx.objectStore('records'),req=st.openCursor(),n=0;req.onsuccess=e=>{let c=e.target.result;if(c){if(predicate(c.value)){c.delete();n++}c.continue()}};tx.oncomplete=()=>res(n);tx.onerror=()=>rej(tx.error)})}
function storeClear(name){return new Promise((res,rej)=>{let r=db.transaction(name,'readwrite').objectStore(name).clear();r.onsuccess=res;r.onerror=()=>rej(r.error)})}

async function reloadData(){records=await storeGetAll('records');imports=await storeGetAll('imports');imports.sort((a,b)=>new Date(b.importedAt)-new Date(a.importedAt));renderAll()}
document.querySelectorAll('.tab').forEach(t=>t.onclick=()=>{document.querySelectorAll('.tab,.page').forEach(x=>x.classList.remove('active'));t.classList.add('active');document.getElementById(t.dataset.page).classList.add('active');if(t.dataset.page==='mov')renderMovTable();if(t.dataset.page==='alertas')renderInvalidMeters();if(t.dataset.page==='imports')renderImports()});

function chooseHeader(headers, target, occurrence=1){
  let n=0,t=normalizeHeader(target);for(let i=0;i<headers.length;i++){if(normalizeHeader(headers[i])===t){n++;if(n===occurrence)return i}}return -1
}
function getVal(row,headers,name,occ=1){let i=chooseHeader(headers,name,occ);return i>=0?row[i]:''}
function firstNonEmpty(...xs){return xs.find(x=>norm(x))??''}

async function handleImport(file){
  if(!file)return;
  try{
    document.getElementById('progressBar').style.width='20%';document.getElementById('importMsg').textContent='Lendo planilha...';
    let data=await file.arrayBuffer(), book=XLSX.read(data,{type:'array',cellDates:false}), sh=book.Sheets[book.SheetNames[0]];
    let matrix=XLSX.utils.sheet_to_json(sh,{header:1,defval:'',raw:false});
    if(matrix.length<2)throw new Error('A planilha não possui dados.');
    let headers=matrix[0].map(x=>norm(x));
    if(chooseHeader(headers,'Medidor Instalado')<0 || chooseHeader(headers,'Medidor Retirado')<0)throw new Error('Não encontrei as colunas "Medidor Instalado" e "Medidor Retirado".');
    document.getElementById('progressBar').style.width='45%';document.getElementById('importMsg').textContent='Analisando movimentações...';

    let batch=[], skippedEmpty=0, importStamp=new Date().toISOString(), batchId='imp_'+Date.now()+'_'+Math.random().toString(36).slice(2,8);
    for(let r=1;r<matrix.length;r++){
      let row=matrix[r];
      let installed=norm(getVal(row,headers,'Medidor Instalado')), removed=norm(getVal(row,headers,'Medidor Retirado')), meter=norm(getVal(row,headers,'Medidor'));
      if(!installed && !removed){skippedEmpty++;continue}
      let type1=getVal(row,headers,'Tipo de Atividade',1), type2=getVal(row,headers,'Tipo de Atividade',2);
      let noteType=norm(type2)||norm(type1);
      let observations=[
        getVal(row,headers,'Observação',1),getVal(row,headers,'Observação',2),
        getVal(row,headers,'Observações',1)
      ].map(norm).filter(Boolean).filter((v,i,a)=>a.indexOf(v)===i).join(' | ');
      let rec={
        data:fmtDateBR(getVal(row,headers,'Data')), recurso:norm(getVal(row,headers,'Recurso')),
        status:norm(getVal(row,headers,'Status da Atividade')), cidade:norm(getVal(row,headers,'Cidade')), bairro:norm(getVal(row,headers,'Bairro')),
        inicio:norm(getVal(row,headers,'Início')), fim:norm(getVal(row,headers,'Fim')),
        ordemServico:norm(getVal(row,headers,'Ordem de Serviço')), numeroNota:norm(getVal(row,headers,'Número da Nota')),
        idAtividade:norm(getVal(row,headers,'ID da Atividade')), medidor:meter, retirado:removed, instalado:installed,
        tipoFase:norm(getVal(row,headers,'Tipo Fase')), tipoNota:noteType, observacoes:observations,
        importedFile:file.name, importedAt:importStamp, importBatchId:batchId
      };
      rec.movimento=installed&&removed?'Instalação + Retirada':installed?'Instalação':'Retirada';
      let identity=[rec.idAtividade,rec.data,rec.numeroNota,rec.ordemServico,rec.instalado,rec.retirado,rec.medidor,rec.recurso].join('|');
      rec.uid=hashString(identity);
      batch.push(rec);
    }
    let existing=new Set(records.map(x=>x.uid)), unique=batch.filter(x=>!existing.has(x.uid));
    document.getElementById('progressBar').style.width='75%';document.getElementById('importMsg').textContent='Salvando na base acumulativa...';
    await storePutMany('records',unique);
    await storeAdd('imports',{fileName:file.name,importedAt:importStamp,batchId,rowsTotal:matrix.length-1,movementsFound:batch.length,newRecords:unique.length,duplicates:batch.length-unique.length,ignoredNoMeter:skippedEmpty});
    document.getElementById('progressBar').style.width='100%';
    document.getElementById('importMsg').innerHTML=`Importação concluída: <b>${unique.length}</b> novos registros; <b>${batch.length-unique.length}</b> duplicados ignorados.`;
    await reloadData();setTimeout(closeImport,1200);
  }catch(e){document.getElementById('importMsg').textContent='Erro: '+e.message;document.getElementById('progressBar').style.width='0'}
}
function openImport(){document.getElementById('importModal').style.display='flex'}
function closeImport(){document.getElementById('importModal').style.display='none';document.getElementById('excelFile').value='';document.getElementById('progressBar').style.width='0';document.getElementById('importMsg').textContent=''}
let dz=document.getElementById('dropzone');['dragenter','dragover'].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.add('drag')}));['dragleave','drop'].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.remove('drag')}));dz.addEventListener('drop',e=>handleImport(e.dataTransfer.files[0]));

function monthKey(v){
  let s=fmtDateBR(v),p=s.split('/');
  return p.length===3?`${p[2]}-${p[1]}`:''
}
function monthLabel(k){
  if(!k||!/^[0-9]{4}-[0-9]{2}$/.test(k))return k;
  const nomes=['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  let [y,m]=k.split('-');
  return `${nomes[Number(m)-1]} / ${y}`
}
function filtered(){
  let m=document.getElementById('fMes')?.value||'',d=document.getElementById('fData')?.value||'',f=document.getElementById('fFase')?.value||'',t=document.getElementById('fTipo')?.value||'';
  return records.filter(r=>(!m||monthKey(r.data)===m)&&(!d||r.data===d)&&(!f||r.tipoFase===f)&&(!t||r.tipoNota===t))
}
function populateFilters(){
  let mesEl=document.getElementById('fMes');
  if(mesEl){
    let currentMes=mesEl.value;
    let meses=[...new Set(records.map(r=>monthKey(r.data)).filter(Boolean))].sort().reverse();
    mesEl.innerHTML='<option value="">Todos</option>'+meses.map(v=>`<option value="${esc(v)}" ${v===currentMes?'selected':''}>${esc(monthLabel(v))}</option>`).join('');
  }
  [['fData','data'],['fFase','tipoFase'],['fTipo','tipoNota']].forEach(([id,key])=>{
    let el=document.getElementById(id),current=el.value;
    let base=records;
    let mesSelecionado=document.getElementById('fMes')?.value||'';
    if(key==='data' && mesSelecionado) base=records.filter(r=>monthKey(r.data)===mesSelecionado);
    let vals=[...new Set(base.map(r=>r[key]).filter(Boolean))].sort((a,b)=>key==='data'?dateKey(b).localeCompare(dateKey(a)):a.localeCompare(b,'pt-BR'));
    el.innerHTML=`<option value="">${key==='tipoNota'?'Todos':'Todas'}</option>`+vals.map(v=>`<option ${v===current?'selected':''}>${esc(v)}</option>`).join('');
    if(!vals.includes(current)) el.value='';
  })
}
function renderAll(){
  populateFilters();let a=filtered(),inst=a.filter(r=>r.instalado).length,ret=a.filter(r=>r.retirado).length;
  kInst.textContent=inst.toLocaleString('pt-BR');kRet.textContent=ret.toLocaleString('pt-BR');kSaldo.textContent=(inst-ret).toLocaleString('pt-BR');kMov.textContent=a.length.toLocaleString('pt-BR');kBases.textContent=imports.length.toLocaleString('pt-BR');
  renderLast(a);renderCharts(a);renderImports();renderInvalidMeterBadge();
}
function rowHtml(r){return `<tr><td>${esc(r.data)}</td><td>${r.movimento==='Instalação'?'<span class="tag inst">INSTALAÇÃO</span>':r.movimento==='Retirada'?'<span class="tag ret">RETIRADA</span>':'<span class="tag both">INST. + RET.</span>'}</td><td>${esc(r.instalado)}</td><td>${esc(r.retirado)}</td><td>${esc(r.medidor)}</td><td>${esc(r.tipoFase)}</td><td>${esc(r.tipoNota)}</td><td>${esc(r.numeroNota)}</td><td>${esc(r.recurso)}</td><td>${esc(r.cidade)}</td><td>${esc(r.bairro)}</td><td>${esc(r.observacoes)}</td></tr>`}
const th=`<thead><tr><th>Data</th><th>Movimento</th><th>Instalado</th><th>Retirado</th><th>Medidor</th><th>Fase</th><th>Tipo de Nota</th><th>Nº Nota</th><th>Equipe</th><th>Cidade</th><th>Bairro</th><th>Observações</th></tr></thead>`;
function renderLast(a){let list=[...a].sort((x,y)=>(dateKey(y.data)+y.importedAt).localeCompare(dateKey(x.data)+x.importedAt)).slice(0,30);lastTable.innerHTML=th+'<tbody>'+list.map(rowHtml).join('')+(list.length?'':'<tr><td colspan="12" class="muted">Nenhuma movimentação encontrada.</td></tr>')+'</tbody>'}
function renderMovTable(){
  let q=(qMov.value||'').toLowerCase(),mt=mTipo.value,limit=+mLimit.value;
  let list=records.filter(r=>(!mt||r.movimento===mt)&&(!q||Object.values(r).join(' ').toLowerCase().includes(q))).sort((a,b)=>dateKey(b.data).localeCompare(dateKey(a.data))).slice(0,limit);
  movTable.innerHTML=th+'<tbody>'+list.map(rowHtml).join('')+(list.length?'':'<tr><td colspan="12" class="muted">Nenhum resultado.</td></tr>')+'</tbody>'
}
function renderCharts(a){
  let byDay={};a.forEach(r=>{let k=r.data||'Sem data';byDay[k]??={i:0,r:0};if(r.instalado)byDay[k].i++;if(r.retirado)byDay[k].r++});
  let days=Object.keys(byDay).sort((x,y)=>dateKey(x).localeCompare(dateKey(y)));
  if(chartDay)chartDay.destroy(); chartDay=new Chart(document.getElementById('chartDay'),{type:'bar',data:{labels:days,datasets:[{label:'Instalados',data:days.map(d=>byDay[d].i),backgroundColor:'#35d07f'},{label:'Retirados',data:days.map(d=>byDay[d].r),backgroundColor:'#ffae42'}]},options:{maintainAspectRatio:false,responsive:true,plugins:{legend:{labels:{color:'#355d7b'}}},scales:{x:{ticks:{color:'#6b7f93'},grid:{color:'#e6edf4'}},y:{beginAtZero:true,ticks:{color:'#8fa6bd',precision:0},grid:{color:'#e6edf4'}}}}});
  let phases={};a.forEach(r=>{let k=r.tipoFase||'Não informado';phases[k]=(phases[k]||0)+1});
  if(chartPhase)chartPhase.destroy();chartPhase=new Chart(document.getElementById('chartPhase'),{type:'doughnut',data:{labels:Object.keys(phases),datasets:[{data:Object.values(phases),backgroundColor:['#21d4fd','#35d07f','#ffd84d','#ff6b6b','#7a78ff','#ff8ce8','#6fd3a8']}]},options:{maintainAspectRatio:false,responsive:true,plugins:{legend:{position:'bottom',labels:{color:'#355d7b'}}}}});
}
function parseMeterQueries(){
  return [...new Set(String(meterSearch.value||'').split(/[\s,;]+/).map(norm).filter(Boolean))];
}
function updateMeterQueryCount(){
  let n=parseMeterQueries().length;
  let el=document.getElementById('meterQueryCount');if(el)el.textContent=`${n.toLocaleString('pt-BR')} medidor(es) informado(s)`;
}
function clearMeterSearch(){
  meterSearch.value='';updateMeterQueryCount();meterResult.innerHTML='<span class="muted">Nenhuma consulta realizada.</span>';
}
function searchMeter(){
  let queries=parseMeterQueries();
  if(!queries.length){meterResult.innerHTML='<span class="muted">Informe pelo menos um número de medidor.</span>';return}
  let found=[],notFound=[];
  for(let q of queries){
    let exact=records.filter(r=>[r.medidor,r.instalado,r.retirado].some(v=>norm(v)===q)).sort((a,b)=>dateKey(a.data).localeCompare(dateKey(b.data)));
    if(exact.length) found.push({q,rows:exact}); else notFound.push(q);
  }
  let totalRows=found.reduce((s,x)=>s+x.rows.length,0);
  let summary=`<div class="grid-kpi" style="grid-template-columns:repeat(3,minmax(150px,1fr));margin-bottom:14px">
    <div class="kpi"><div class="label">Consultados</div><div class="value">${queries.length.toLocaleString('pt-BR')}</div></div>
    <div class="kpi inst"><div class="label">Encontrados</div><div class="value">${found.length.toLocaleString('pt-BR')}</div></div>
    <div class="kpi"><div class="label">Não encontrados</div><div class="value" style="color:var(--red)">${notFound.length.toLocaleString('pt-BR')}</div></div>
  </div>`;
  let foundHtml=found.length?`<h3>Medidores encontrados (${found.length})</h3>`+found.map(x=>`<div style="margin:14px 0 6px"><b>Medidor: ${esc(x.q)}</b> <span class="muted small">• ${x.rows.length} registro(s)</span></div><div class="table-wrap" style="max-height:300px"><table>${th}<tbody>${x.rows.map(rowHtml).join('')}</tbody></table></div>`).join(''):`<div class="notice">Nenhum dos medidores informados foi encontrado.</div>`;
  let missingHtml=notFound.length?`<div class="panel" style="margin-top:15px;background:#fff8f8"><h3>Não encontrados (${notFound.length})</h3><div style="display:flex;gap:7px;flex-wrap:wrap">${notFound.map(x=>`<span class="tag" style="background:#fdecee;color:#b02a37">${esc(x)}</span>`).join('')}</div></div>`:'';
  meterResult.innerHTML=summary+`<div class="small muted" style="margin-bottom:10px">${totalRows.toLocaleString('pt-BR')} registro(s) localizado(s) no histórico.</div>`+foundHtml+missingHtml;
}

function meterDigits(v){return String(v??'').replace(/\D/g,'')}
// Regra da planilha CALCULAR DÍGITO VERIFICADOR: pesos 5,4,3,2,7,6,5,4,3,2; módulo 11.
// Só sugere correção quando o número original possui EXATAMENTE 10 dígitos.
function calcularDigitoMedidor(v){
  let d=meterDigits(v);if(d.length!==10)return null;
  const pesos=[5,4,3,2,7,6,5,4,3,2];
  let soma=d.split('').reduce((acc,n,i)=>acc+(+n)*pesos[i],0), resto=soma%11, dv=11-resto;
  if(dv===10||dv===11)dv=0;
  return {original:d,dv:String(dv),corrigido:d+dv};
}
function invalidInstalledMeters(){return records.filter(r=>norm(r.instalado) && meterDigits(r.instalado).length<11)}
function renderInvalidMeterBadge(){
  let bad=invalidInstalledMeters(), installed=records.filter(r=>norm(r.instalado));
  let badge=document.getElementById('invalidMeterBadge');if(badge){badge.textContent=bad.length.toLocaleString('pt-BR');badge.classList.toggle('show',bad.length>0)}
  let a=document.getElementById('kInstalledChecked'),o=document.getElementById('kInstalledOk'),b=document.getElementById('kInstalledInvalid');
  if(a)a.textContent=installed.length.toLocaleString('pt-BR');if(o)o.textContent=(installed.length-bad.length).toLocaleString('pt-BR');if(b)b.textContent=bad.length.toLocaleString('pt-BR');
}
function renderInvalidMeters(){
  renderInvalidMeterBadge();
  let table=document.getElementById('invalidMeterTable');if(!table)return;
  let q=(document.getElementById('qInvalidMeter')?.value||'').toLowerCase(),limit=+(document.getElementById('invalidLimit')?.value||100);
  let list=invalidInstalledMeters().filter(r=>!q||Object.values(r).join(' ').toLowerCase().includes(q)).sort((a,b)=>dateKey(b.data).localeCompare(dateKey(a.data))).slice(0,limit);
  let head='<thead><tr><th>Alerta</th><th>Original da Base</th><th>Dígitos</th><th>Dígito Calculado</th><th>Sugestão Corrigida</th><th>Ação</th><th>Data</th><th>Nº Nota</th><th>Ordem de Serviço</th><th>Equipe</th><th>Tipo de Nota</th><th>Fase</th><th>Cidade</th><th>Bairro</th><th>Observações</th><th>Arquivo</th></tr></thead>';
  let body=list.map(r=>{let calc=calcularDigitoMedidor(r.instalado);return `<tr class="invalid-meter-row"><td><span class="tag alert">${calc?'DÍGITO CALCULADO — CONFERIR':'CONFERIR'}</span></td><td><b>${esc(r.instalado)}</b></td><td><b style="color:var(--red)">${meterDigits(r.instalado).length}</b></td><td>${calc?'<b>'+esc(calc.dv)+'</b>':'—'}</td><td>${calc?'<b style="color:var(--green)">'+esc(calc.corrigido)+'</b>':'Não calculado'}</td><td>${calc?'<button onclick="navigator.clipboard.writeText(\''+esc(calc.corrigido)+'\');this.textContent=\'Copiado ✓\';setTimeout(()=>this.textContent=\'Copiar sugestão\',1200)">Copiar sugestão</button>':'—'}</td><td>${esc(r.data)}</td><td>${esc(r.numeroNota)}</td><td>${esc(r.ordemServico)}</td><td>${esc(r.recurso)}</td><td>${esc(r.tipoNota)}</td><td>${esc(r.tipoFase)}</td><td>${esc(r.cidade)}</td><td>${esc(r.bairro)}</td><td>${esc(r.observacoes)}</td><td>${esc(r.importedFile)}</td></tr>`}).join('');
  table.innerHTML=head+'<tbody>'+body+(list.length?'':'<tr><td colspan="16" class="muted">Nenhum medidor instalado com menos de 11 dígitos. ✓</td></tr>')+'</tbody>';
}

function exportInvalidMeters(){
  let bad=invalidInstalledMeters();if(!bad.length){alert('Não há medidores instalados com menos de 11 dígitos.');return}
  let rows=bad.map(r=>{let calc=calcularDigitoMedidor(r.instalado);return {'Medidor Original':r.instalado,'Quantidade de Dígitos':meterDigits(r.instalado).length,'Dígito Calculado':calc?calc.dv:'','Sugestão Corrigida':calc?calc.corrigido:'','Situação':calc?'DÍGITO CALCULADO - CONFERIR':'CONFERIR - NÃO CALCULADO','Data':r.data,'Número da Nota':r.numeroNota,'Ordem de Serviço':r.ordemServico,'Equipe':r.recurso,'Tipo de Nota':r.tipoNota,'Tipo Fase':r.tipoFase,'Cidade':r.cidade,'Bairro':r.bairro,'Observações':r.observacoes,'Arquivo Importado':r.importedFile}});
  let ws=XLSX.utils.json_to_sheet(rows),book=XLSX.utils.book_new();XLSX.utils.book_append_sheet(book,ws,'Digitos_Conferir');XLSX.writeFile(book,'Medidores_Digito_Calculado_CONFERIR_'+new Date().toISOString().slice(0,10)+'.xlsx');
}

function renderImports(){
  importList.innerHTML=imports.length?imports.map(x=>`<div class="import-card" style="grid-template-columns:1.3fr .55fr .55fr .55fr .55fr auto"><div><b>${esc(x.fileName)}</b><div class="small muted">${new Date(x.importedAt).toLocaleString('pt-BR')}</div></div><div><span class="muted small">Movimentações</span><br><b>${x.movementsFound??0}</b></div><div><span class="muted small">Novos</span><br><b style="color:#198754">${x.newRecords??0}</b></div><div><span class="muted small">Duplicados</span><br><b style="color:#e8871e">${x.duplicates??0}</b></div><div><span class="muted small">Linhas lidas</span><br><b>${x.rowsTotal??0}</b></div><div><button class="danger" onclick="deleteImportBase(${x.id})">Excluir base</button></div></div>`).join(''):'<div class="muted">Ainda não há bases importadas.</div>'
}

async function deleteImportBase(id){
  let imp=imports.find(x=>x.id===id);if(!imp)return;
  let msg=`Excluir a base \"${imp.fileName}\" importada em ${new Date(imp.importedAt).toLocaleString('pt-BR')}?\n\nOs registros pertencentes a essa importação também serão excluídos.`;
  if(!confirm(msg))return;
  let removed=0;
  if(imp.batchId){removed=await deleteRecordsWhere(r=>r.importBatchId===imp.batchId)}
  else{
    // Compatibilidade com bases antigas criadas antes da identificação por lote.
    removed=await deleteRecordsWhere(r=>r.importedFile===imp.fileName);
  }
  await storeDelete('imports',id);
  await reloadData();
  alert(`Base excluída com sucesso. ${removed} registro(s) removido(s).`);
}
function parseBulkMeters(){
  let raw=norm(document.getElementById('bulkMeters')?.value||'');if(!raw)return [];
  return [...new Set(raw.split(/[\s,;|]+/).map(x=>x.trim()).filter(Boolean))];
}
function updateBulkCount(){let el=document.getElementById('bulkCount');if(el)el.textContent=parseBulkMeters().length.toLocaleString('pt-BR')}
function isoToBR(v){if(!v)return '';let [y,m,d]=v.split('-');return `${d}/${m}/${y}`}
function clearBulkForm(){['bulkDate','bulkPhase','bulkNoteType','bulkNoteNumber','bulkResource','bulkCity','bulkNeighborhood','bulkObs','bulkMeters'].forEach(id=>{let e=document.getElementById(id);if(e)e.value=''});document.getElementById('bulkMovType').value='Instalação';updateBulkCount()}
async function saveBulkMeters(){
  let meters=parseBulkMeters();if(!meters.length){alert('Cole ou digite pelo menos um número de medidor.');return}
  let mov=document.getElementById('bulkMovType').value,date=isoToBR(document.getElementById('bulkDate').value),stamp=new Date().toISOString(),batchId='manual_'+Date.now()+'_'+Math.random().toString(36).slice(2,8);
  if(!date){alert('Informe a data da movimentação.');return}
  let common={data:date,recurso:norm(bulkResource.value),status:'',cidade:norm(bulkCity.value),bairro:norm(bulkNeighborhood.value),inicio:'',fim:'',ordemServico:'',numeroNota:norm(bulkNoteNumber.value),idAtividade:'',tipoFase:norm(bulkPhase.value),tipoNota:norm(bulkNoteType.value),observacoes:norm(bulkObs.value),importedFile:'LANÇAMENTO MANUAL EM MASSA',importedAt:stamp,importBatchId:batchId};
  let batch=meters.map((m,i)=>{let r={...common,medidor:m,instalado:mov==='Instalação'?m:'',retirado:mov==='Retirada'?m:'',movimento:mov};let identity=['manual',r.data,r.movimento,m,r.numeroNota,r.recurso,i].join('|');r.uid=hashString(identity);return r});
  let existing=new Set(records.map(x=>x.uid)),unique=batch.filter(x=>!existing.has(x.uid));
  await storePutMany('records',unique);
  await storeAdd('imports',{fileName:`Lançamento manual em massa - ${mov}`,importedAt:stamp,batchId,rowsTotal:meters.length,movementsFound:meters.length,newRecords:unique.length,duplicates:meters.length-unique.length,ignoredNoMeter:0,manual:true});
  await reloadData();
  alert(`${unique.length} medidor(es) salvo(s) com sucesso.`);clearBulkForm();
}

function backupData(){let payload={app:'Controle de Medidores',version:1,createdAt:new Date().toISOString(),records,imports};download('Backup_Medidores_'+new Date().toISOString().slice(0,10)+'.json',JSON.stringify(payload,null,2))}
async function restoreBackup(file){
  if(!file)return;try{let obj=JSON.parse(await file.text());if(!Array.isArray(obj.records)||!Array.isArray(obj.imports))throw new Error('Backup inválido.');
  if(!confirm(`Restaurar ${obj.records.length} registros? Os dados atuais serão substituídos.`))return;
  await storeClear('records');await storeClear('imports');await storePutMany('records',obj.records);for(let x of obj.imports){let y={...x};delete y.id;await storeAdd('imports',y)}await reloadData();alert('Backup restaurado com sucesso.')}catch(e){alert('Erro ao restaurar: '+e.message)}
}
async function clearDatabase(){if(!confirm('ATENÇÃO: deseja apagar TODOS os registros e históricos deste navegador?'))return;if(!confirm('Confirme novamente. Essa ação não pode ser desfeita sem um backup.'))return;await storeClear('records');await storeClear('imports');await reloadData()}
function exportExcel(){
  if(!records.length){alert('Não há dados para exportar.');return}
  let rows=records.map(r=>({'Data':r.data,'Movimento':r.movimento,'Medidor':r.medidor,'Medidor Retirado':r.retirado,'Medidor Instalado':r.instalado,'Tipo Fase':r.tipoFase,'Tipo de Nota':r.tipoNota,'Número da Nota':r.numeroNota,'Ordem de Serviço':r.ordemServico,'ID da Atividade':r.idAtividade,'Recurso':r.recurso,'Status':r.status,'Cidade':r.cidade,'Bairro':r.bairro,'Início':r.inicio,'Fim':r.fim,'Observações':r.observacoes,'Arquivo Importado':r.importedFile,'Data da Importação':new Date(r.importedAt).toLocaleString('pt-BR')}));
  let ws=XLSX.utils.json_to_sheet(rows), book=XLSX.utils.book_new();XLSX.utils.book_append_sheet(book,ws,'Medidores');XLSX.writeFile(book,'Controle_Medidores_Consolidado_'+new Date().toISOString().slice(0,10)+'.xlsx')
}
openDB().then(()=>{let d=document.getElementById('bulkDate');if(d&&!d.value)d.value=new Date().toISOString().slice(0,10);return reloadData()}).catch(e=>alert('Erro ao iniciar o banco local: '+e.message));
