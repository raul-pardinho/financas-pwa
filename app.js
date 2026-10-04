(()=>{
const APP_VERSION='2.0.0-beta.1',DB='financas_dashboard_v2',STORE='workbook',THEME='financas_theme_v2';
const state={data:null,month:null,page:'overview',theme:'light'};
const $=s=>document.querySelector(s),$$=s=>[...document.querySelectorAll(s)];
const money=new Intl.NumberFormat('pt-BR',{style:'currency',currency:'BRL'});
const pctFmt=new Intl.NumberFormat('pt-BR',{style:'percent',minimumFractionDigits:0,maximumFractionDigits:1});
const monthFmt=new Intl.DateTimeFormat('pt-BR',{month:'long',year:'numeric',timeZone:'UTC'});
const shortMonthFmt=new Intl.DateTimeFormat('pt-BR',{month:'short',year:'2-digit',timeZone:'UTC'});
const nonOperational=new Set(['TRANSITORIO_TERCEIROS','TRANSFERENCIA_INTERNA','FINANCIAMENTO_RECEBIDO']);
const labels={
  REMUNERACAO_RAUL:'Remuneração Raul',REMUNERACAO_LETICIA:'Remuneração Letícia',ALUGUEL_RECEBIDO:'Aluguel recebido',
  CONTRIBUICAO_ESCOLA:'Contribuição escola',DECIMO_TERCEIRO:'13º Letícia',MORADIA_ATUAL:'Moradia',
  IMOVEL_ALUGADO:'Imóvel próprio',VEICULO:'Veículo',VITOR:'Vitor',CASA:'Casa',TRANSPORTE:'Transporte',
  ALIMENTACAO:'Alimentação',SAUDE:'Saúde',PESSOAL_RAUL:'Pessoal Raul',PESSOAL_LETICIA:'Pessoal Letícia',
  ASSINATURAS:'Assinaturas',IA_TECNOLOGIA:'IA e tecnologia',PRESENTES_EVENTOS:'Presentes e eventos',
  PETS:'Pets',VIAGEM:'Viagem',DIVIDA_PESSOAL:'Empréstimo pessoal',A_CLASSIFICAR:'A classificar'
};

function n(v){
  if(typeof v==='number')return Number.isFinite(v)?v:0;
  if(v==null||v==='')return 0;
  let s=String(v).trim().replace(/R\$/g,'').replace(/\s/g,'');
  if(/^-?\d+(\.\d+)?$/.test(s))return +s;
  s=s.replace(/\./g,'').replace(',','.');
  return Number.isFinite(+s)?+s:0;
}
function pm(v){
  if(!v)return null;
  const s=String(v).trim();
  if(/^\d{4}-\d{2}$/.test(s))return s;
  if(/^\d{4}-\d{2}-\d{2}/.test(s))return s.slice(0,7);
  const d=new Date(s);
  return Number.isNaN(d)?null:d.toISOString().slice(0,7);
}
function mdate(m){
  const[y,mo]=m.split('-').map(Number);
  return new Date(Date.UTC(y,mo-1,1));
}
function monthLabel(m){
  if(!m)return'—';
  let s=monthFmt.format(mdate(m)).replace('.','');
  return s[0].toUpperCase()+s.slice(1);
}
function shortMonth(m){return shortMonthFmt.format(mdate(m)).replace('.','');}
function addM(m,k){
  const d=mdate(m);d.setUTCMonth(d.getUTCMonth()+k);return d.toISOString().slice(0,7);
}
function rows(matrix){
  if(!matrix?.length)return[];
  const h=matrix[0].map(x=>String(x??'').trim());
  return matrix.slice(1).filter(r=>r.some(x=>x!==''&&x!=null))
    .map(r=>Object.fromEntries(h.map((x,i)=>[x,r[i]??'']).filter(x=>x[0])));
}
function wideBudgetRows(matrix){
  if(!matrix?.length)return[];
  const hi=matrix.findIndex(r=>String(r?.[0]||'').trim().toUpperCase()==='NATUREZA');
  if(hi<0)return[];
  const h=matrix[hi].map(x=>String(x??'').trim());
  return matrix.slice(hi+1).filter(r=>['RECEITA','DESPESA'].includes(String(r?.[0]||'').toUpperCase()))
    .map(r=>Object.fromEntries(h.map((x,i)=>[x,r[i]??'']).filter(x=>x)));
}
function toast(t){
  const e=$('#toast');e.textContent=t;e.classList.remove('hidden');clearTimeout(toast.t);
  toast.t=setTimeout(()=>e.classList.add('hidden'),3200);
}
function dbOpen(){
  return new Promise((ok,no)=>{
    const r=indexedDB.open(DB,1);
    r.onupgradeneeded=()=>{if(!r.result.objectStoreNames.contains(STORE))r.result.createObjectStore(STORE)};
    r.onsuccess=()=>ok(r.result);r.onerror=()=>no(r.error);
  });
}
async function dbSet(v){const d=await dbOpen();return new Promise((ok,no)=>{const tx=d.transaction(STORE,'readwrite');tx.objectStore(STORE).put(v,'data');tx.oncomplete=ok;tx.onerror=()=>no(tx.error)})}
async function dbGet(){const d=await dbOpen();return new Promise((ok,no)=>{const r=d.transaction(STORE,'readonly').objectStore(STORE).get('data');r.onsuccess=()=>ok(r.result||null);r.onerror=()=>no(r.error)})}
async function dbClear(){const d=await dbOpen();return new Promise((ok,no)=>{const tx=d.transaction(STORE,'readwrite');tx.objectStore(STORE).clear();tx.oncomplete=ok;tx.onerror=()=>no(tx.error)})}

function parseWorkbook(buf,name){
  if(!window.XLSX)throw Error('Biblioteca XLSX não carregou. Atualize a página.');
  const wb=XLSX.read(buf,{type:'array',raw:false}),sheets={};
  wb.SheetNames.forEach(s=>{
    const matrix=XLSX.utils.sheet_to_json(wb.Sheets[s],{header:1,defval:'',raw:false});
    sheets[s]=s.startsWith('orcamento_v')?wideBudgetRows(matrix):rows(matrix);
  });
  for(const req of ['movimentacoes','parcelamentos','patrimonio','controle'])if(!sheets[req])throw Error('Aba obrigatória ausente: '+req);
  return{fileName:name,importedAt:new Date().toISOString(),sheets};
}
function controls(){return Object.fromEntries((state.data?.sheets?.controle||[]).map(r=>[String(r.chave),r.valor]))}
function planning(){return state.data?.sheets?.planejamento||[]}
function movements(){return state.data?.sheets?.movimentacoes||[]}
function installments(){return state.data?.sheets?.parcelamentos||[]}
function patrimony(){return state.data?.sheets?.patrimonio||[]}

function months(){
  const s=new Set();
  movements().forEach(r=>[pm(r.competencia||r.data_transacao),pm(r.data_caixa)].forEach(x=>x&&s.add(x)));
  planning().forEach(r=>{const x=pm(r.competencia);if(x)s.add(x)});
  const wb=state.data?.sheets?.orcamento_v2||state.data?.sheets?.orcamento_v1||[];
  wb.forEach(r=>Object.keys(r).filter(k=>/^\d{4}-\d{2}$/.test(k)&&n(r[k])!==0).forEach(k=>s.add(k)));
  return[...s].sort();
}
function todayIso(){return new Date().toISOString().slice(0,10)}
function selectedRelation(m=state.month){
  const cur=todayIso().slice(0,7);
  return m<cur?'past':m>cur?'future':'current';
}
function daysInMonth(m){const[y,mo]=m.split('-').map(Number);return new Date(Date.UTC(y,mo,0)).getUTCDate()}
function calendarRatio(m){
  const rel=selectedRelation(m);if(rel==='past')return 1;if(rel==='future')return 0;
  return Math.min(1,new Date().getDate()/daysInMonth(m));
}
function workdayCounts(m){
  const[y,mo]=m.split('-').map(Number),last=daysInMonth(m),today=selectedRelation(m)==='current'?new Date().getDate():last;
  let total=0,elapsed=0;
  for(let d=1;d<=last;d++){const dow=new Date(Date.UTC(y,mo-1,d)).getUTCDay();if(dow>=1&&dow<=5){total++;if(d<=today)elapsed++;}}
  if(selectedRelation(m)==='future')elapsed=0;
  return{total,elapsed,ratio:total?elapsed/total:0};
}

function rawMov(month=state.month,view='COMPETENCIA'){
  const today=todayIso();
  return movements().filter(r=>{
    if(String(r.status_movimento||'REALIZADO').toUpperCase()==='CANCELADO')return false;
    if(view==='CAIXA'){
      const dc=String(r.data_caixa||'');
      return pm(dc)===month&&dc&&dc<=today;
    }
    if(String(r.categoria||'').toUpperCase()==='FATURA_CARTAO_PAGA_PENDENTE')return false;
    return pm(r.competencia||r.data_transacao)===month;
  });
}
function competenceMov(month=state.month){
  const a=rawMov(month,'COMPETENCIA'),out=[],seen=new Set();
  for(const r of a){
    if(String(r.tipo||'').toUpperCase()==='PARCELA'&&Math.abs(n(r.valor_total_compra))>0){
      const key=String(r.id_parcelamento||[r.titular,r.data_transacao,r.estabelecimento,r.valor_total_compra,r.parcelas_total].join('|'));
      if(seen.has(key))continue;seen.add(key);
      out.push({...r,valor:(n(r.valor)<0?-1:1)*Math.abs(n(r.valor_total_compra))});
    }else out.push(r);
  }
  return out;
}
function isOperational(r){return !nonOperational.has(String(r.grupo||'').toUpperCase())}
function confirmedIncome(r){
  return isOperational(r)&&n(r.valor)>0&&String(r.grupo||'').toUpperCase()==='RECEITA'&&String(r.categoria||'').toUpperCase()!=='A_CLASSIFICAR';
}
function operationalExpense(r){
  return isOperational(r)&&n(r.valor)<0&&String(r.grupo||'').toUpperCase()!=='RECEITA'&&String(r.categoria||'').toUpperCase()!=='RESERVA_FAMILIAR';
}
function canonSub(category,sub='',desc=''){
  const c=String(category||'').toUpperCase(),s=(String(sub||'')+' '+String(desc||'')).toLowerCase();
  if(c==='ALIMENTACAO'){
    if(/supermerc/.test(s))return'Supermercado';
    if(/conveni|padaria|oxxo/.test(s))return'Conveniência';
    if(/restaurante|delivery|ifood|fast.?food|mcdonald/.test(s))return'Restaurantes';
    if(/trabalho|almo[cç]o/.test(s))return'Alimentação dia a dia — trabalho Raul';
  }
  if(c==='TRANSPORTE'&&/trabalho|ônibus|onibus|autopass|transurc/.test(s))return'Deslocamento trabalho Raul';
  return String(sub||'').trim()||labels[c]||c||'Outros';
}
function budgetKey(category,sub){return String(category||'').toUpperCase()+'|'+canonSub(category,sub)}
function movementKey(r){return budgetKey(r.categoria,canonSub(r.categoria,r.subcategoria,r.descricao_original))}

function wideBudgetSheet(){
  if(state.data?.sheets?.orcamento_v2?.length)return{name:'orcamento_v2',label:'Orçamento V2 · rascunho',rows:state.data.sheets.orcamento_v2};
  if(state.data?.sheets?.orcamento_v1?.length)return{name:'orcamento_v1',label:'Orçamento V1 · rascunho',rows:state.data.sheets.orcamento_v1};
  return null;
}
function budgetLines(month=state.month){
  const wide=wideBudgetSheet();
  if(wide){
    return wide.rows.filter(r=>n(r[month])!==0).map(r=>({
      nature:String(r.Natureza||'').toUpperCase(),group:String(r.Grupo||''),category:String(r.Categoria||''),
      sub:String(r.Subcategoria||''),holder:String(r.Titular||''),method:String(r['Método forecast']||'').toUpperCase(),
      value:Math.abs(n(r[month])),source:wide.label
    }));
  }
  return planning().filter(r=>pm(r.competencia)===month&&String(r.cenario||'').toUpperCase()==='BUDGET'&&String(r.visao||'').toUpperCase()==='COMPETENCIA')
    .map(r=>({nature:n(r.valor)>=0?'RECEITA':'DESPESA',group:String(r.grupo||''),category:String(r.categoria||''),sub:String(r.subcategoria||''),holder:String(r.titular||''),method:'',value:Math.abs(n(r.valor)),source:'Planejamento oficial'}));
}
function metaLines(month=state.month){
  return planning().filter(r=>pm(r.competencia)===month&&String(r.cenario||'').toUpperCase()==='META'&&String(r.visao||'').toUpperCase()==='COMPETENCIA')
    .map(r=>({nature:n(r.valor)>=0?'RECEITA':'DESPESA',category:String(r.categoria||''),sub:String(r.subcategoria||''),value:Math.abs(n(r.valor))}));
}
function budgetSource(){return budgetLines()[0]?.source||'Sem orçamento para o período'}

function realizedByKey(month=state.month){
  const map=new Map();
  competenceMov(month).filter(operationalExpense).forEach(r=>{
    const key=movementKey(r),x=map.get(key)||{value:0,category:String(r.categoria||''),sub:canonSub(r.categoria,r.subcategoria,r.descricao_original)};
    x.value+=Math.abs(n(r.valor));map.set(key,x);
  });
  return map;
}
function metaByKey(month=state.month){
  const map=new Map();metaLines(month).filter(r=>r.nature==='DESPESA').forEach(r=>map.set(budgetKey(r.category,r.sub),r.value));return map;
}
function forecastValue(method,budget,realized,month=state.month){
  const rel=selectedRelation(month);if(rel==='past')return realized;if(rel==='future')return budget;
  const m=String(method||'').toUpperCase();
  if(m==='PACING_CALENDARIO'){
    const ratio=calendarRatio(month);if(ratio<.15)return Math.max(realized,budget);
    return Math.max(realized,realized/Math.max(ratio,.01));
  }
  if(m==='PACING_DIAS_UTEIS'){
    const ratio=workdayCounts(month).ratio;if(ratio<.15)return Math.max(realized,budget);
    return Math.max(realized,realized/Math.max(ratio,.01));
  }
  return Math.max(realized,budget);
}
function categoryRows(month=state.month){
  const b=budgetLines(month).filter(r=>r.nature==='DESPESA'),real=realizedByKey(month),meta=metaByKey(month),out=new Map();
  b.forEach(r=>{
    const key=budgetKey(r.category,r.sub),x=out.get(key)||{key,category:r.category,sub:canonSub(r.category,r.sub),budget:0,meta:null,realized:0,forecast:0,method:r.method};
    x.budget+=r.value;x.method=r.method||x.method;out.set(key,x);
  });
  for(const[key,x]of real){
    const y=out.get(key)||{key,category:x.category,sub:x.sub,budget:0,meta:null,realized:0,forecast:0,method:'SEM_ORCAMENTO'};
    y.realized=x.value;out.set(key,y);
  }
  for(const[key,x]of out){
    x.realized=real.get(key)?.value||0;
    x.meta=meta.has(key)?meta.get(key):null;
    x.forecast=forecastValue(x.method,x.budget,x.realized,month);
    x.deltaBudget=x.forecast-x.budget;
    x.deltaMeta=x.meta==null?null:x.forecast-x.meta;
  }
  return[...out.values()].sort((a,b)=>Math.max(b.budget,b.forecast,b.realized)-Math.max(a.budget,a.forecast,a.realized));
}
function monthlySummary(month=state.month){
  const b=budgetLines(month),cats=categoryRows(month),cm=competenceMov(month);
  const budgetIncome=b.filter(r=>r.nature==='RECEITA').reduce((s,r)=>s+r.value,0);
  const budgetExpense=cats.reduce((s,r)=>s+r.budget,0);
  const metaVals=cats.filter(r=>r.meta!=null),metaExpense=metaVals.length?metaVals.reduce((s,r)=>s+r.meta,0):null;
  const realizedIncome=cm.filter(confirmedIncome).reduce((s,r)=>s+n(r.valor),0);
  const realizedExpense=cats.reduce((s,r)=>s+r.realized,0);
  const forecastExpense=cats.reduce((s,r)=>s+r.forecast,0);
  const rel=selectedRelation(month);
  const forecastIncome=rel==='past'?realizedIncome:Math.max(realizedIncome,budgetIncome);
  return{budgetIncome,budgetExpense,metaExpense,realizedIncome,realizedExpense,forecastIncome,forecastExpense,
    forecastResult:forecastIncome-forecastExpense,budgetResult:budgetIncome-budgetExpense};
}
function statusFor(row){
  if(!row.budget)return'neutral';
  const ratio=row.forecast/row.budget;
  if(row.meta!=null&&row.forecast<=row.meta)return'good';
  if(ratio<=1)return'good';
  if(ratio<=1.08)return'warn';
  return'bad';
}
function statusText(s){return s==='good'?'No caminho':s==='warn'?'Atenção':s==='bad'?'Acima':'Sem orçamento'}

function latestPatrimony(){
  const snaps=patrimonySnapshots();return snaps.at(-1)||{key:null,assets:0,liabilities:0,net:0,liquid:0,familyReserve:0,items:[]};
}
function patrimonySnapshots(){
  const by=new Map();
  patrimony().forEach(r=>{
    const key=pm(r.data_referencia||r.competencia||r.data)||'Atual';
    if(!by.has(key))by.set(key,{key,assets:0,liabilities:0,net:0,liquid:0,familyReserve:0,items:[]});
    const x=by.get(key),cls=String(r.classe||'').toUpperCase(),g=String(r.grupo||'').toUpperCase(),v=Math.abs(n(r.valor_base));
    if(cls==='ATIVO')x.assets+=v;else if(cls==='PASSIVO')x.liabilities+=v;
    if(cls==='ATIVO'&&['RESERVA','RESERVA_PESSOAL'].includes(g))x.liquid+=v;
    if(cls==='ATIVO'&&g==='RESERVA')x.familyReserve+=v;
    x.items.push({...r,value:v});
  });
  return[...by.values()].map(x=>({...x,net:x.assets-x.liabilities})).sort((a,b)=>String(a.key).localeCompare(String(b.key)));
}
function activeInstallmentRows(){
  return installments().filter(r=>String(r.status||'ATIVO').toUpperCase()==='ATIVO');
}
function remainingInstallments(r){
  const given=n(r.parcelas_restantes);if(given)return given;
  return Math.max(0,n(r.numero_parcelas)-n(r.parcelas_pagas));
}
function thirdPartyRows(){
  return activeInstallmentRows().filter(r=>String(r.terceiro||'').toLowerCase()==='true'||r.terceiro===true||String(r.responsavel_terceiro||'').trim());
}
function thirdPartySummary(){
  const a=thirdPartyRows(),exposure=a.reduce((s,r)=>s+remainingInstallments(r)*Math.abs(n(r.valor_parcela)),0);
  const current=a.reduce((s,r)=>{const st=pm(r.competencia_inicio),en=pm(r.competencia_fim);return st&&en&&state.month>=st&&state.month<=en?s+Math.abs(n(r.valor_parcela)):s},0);
  const reimb=rawMov(state.month,'CAIXA').filter(r=>n(r.valor)>0&&(String(r.responsavel_terceiro||'').trim()||String(r.categoria||'').toUpperCase()==='TERCEIROS_HEITOR')).reduce((s,r)=>s+n(r.valor),0);
  return{rows:a,exposure,current,reimb};
}
function debtRows(){
  const p=latestPatrimony(),b=budgetLines();
  return p.items.filter(r=>String(r.classe||'').toUpperCase()==='PASSIVO').map(r=>{
    const group=String(r.grupo||'').toUpperCase();
    let monthly=0;
    const par=activeInstallmentRows().find(x=>String(x.categoria||'').toUpperCase()===group&&!x.terceiro);
    if(par)monthly=Math.abs(n(par.valor_parcela));
    if(!monthly){
      const match=b.find(x=>x.nature==='DESPESA'&&String(x.category||'').toUpperCase()===(group==='IMOVEL'?'IMOVEL_ALUGADO':group)&&/financ/i.test(x.sub));
      if(match)monthly=match.value;
    }
    return{name:r.item||labels[group]||group,group,balance:Math.abs(n(r.valor_base)),monthly,note:r.observacao||'',origin:r.origem||''};
  });
}
function reserveCurrent(){return latestPatrimony().familyReserve}

function cashSummary(){
  const a=rawMov(state.month,'CAIXA'),external=a.filter(r=>String(r.grupo||'').toUpperCase()!=='TRANSFERENCIA_INTERNA'&&String(r.grupo||'').toUpperCase()!=='FINANCIAMENTO_RECEBIDO');
  const inflow=external.filter(r=>n(r.valor)>0).reduce((s,r)=>s+n(r.valor),0),outflow=Math.abs(external.filter(r=>n(r.valor)<0).reduce((s,r)=>s+n(r.valor),0));
  const familyIn=external.filter(r=>n(r.valor)>0&&String(r.grupo||'').toUpperCase()!=='TRANSITORIO_TERCEIROS').reduce((s,r)=>s+n(r.valor),0);
  const familyOut=Math.abs(external.filter(r=>n(r.valor)<0&&String(r.grupo||'').toUpperCase()!=='TRANSITORIO_TERCEIROS').reduce((s,r)=>s+n(r.valor),0));
  return{inflow,outflow,result:inflow-outflow,familyIn,familyOut,familyResult:familyIn-familyOut};
}
function cardSummary(){
  const map=new Map();
  movements().filter(r=>String(r.conta_cartao||'').toUpperCase().includes('CARTAO')&&pm(r.data_caixa)===state.month&&n(r.valor)<0).forEach(r=>{
    const name=String(r.conta_cartao),x=map.get(name)||{name,total:0,family:0,third:0,count:0};
    const v=Math.abs(n(r.valor));x.total+=v;x.count++;
    if(String(r.responsavel_terceiro||'').trim()||String(r.natureza||'').toUpperCase()==='TERCEIRO')x.third+=v;else x.family+=v;
    map.set(name,x);
  });
  return[...map.values()].sort((a,b)=>b.total-a.total);
}
function currentBalances(){
  const c=controls(),out=[];
  Object.entries(c).forEach(([k,v])=>{
    if(/^saldo_/.test(k)&&/2026_\d{2}/.test(k))out.push({key:k,value:n(v)});
  });
  return out;
}

function kpi(label,value,sub='',tone=''){
  const display=typeof value==='number'?money.format(value):value;
  return'<article class="kpi '+tone+'"><span class="label">'+label+'</span><div class="value">'+display+'</div><div class="sub">'+sub+'</div></article>';
}
function renderOverview(){
  const s=monthlySummary(),delta=s.forecastExpense-s.budgetExpense,meta=s.metaExpense;
  const health=!s.budgetExpense?'neutral':delta<=0?'good':delta/s.budgetExpense<=.05?'warn':'bad';
  $('#overviewTitle').textContent='Como estamos em '+monthLabel(state.month);
  $('#monthHealth').className='status-pill '+health;$('#monthHealth').textContent=health==='good'?'Dentro do orçamento':health==='warn'?'Atenção ao fechamento':health==='bad'?'Pressão no orçamento':'Sem orçamento';
  $('#overviewKpis').innerHTML=[
    kpi('Receita realizada',s.realizedIncome,'Receita confirmada por competência'),
    kpi('Orçamento',s.budgetExpense,'Teto realista do mês'),
    kpi('Meta',meta==null?'A definir':money.format(meta),'Stretch: desafio abaixo do orçamento'),
    kpi('Realizado MTD',s.realizedExpense,pctFmt.format(s.budgetExpense?s.realizedExpense/s.budgetExpense:0)+' do orçamento'),
    kpi('Forecast',s.forecastExpense,(delta>=0?'+':'')+money.format(delta)+' vs orçamento',health),
    kpi('Resultado projetado',s.forecastResult,'Receita forecast menos despesas',s.forecastResult>=0?'good':'bad')
  ].join('');
  const vals=[['Orçamento',s.budgetExpense,'budget'],...(meta==null?[]:[['Meta',meta,'target']]),['Realizado',s.realizedExpense,'realized'],['Forecast',s.forecastExpense,'forecast']],mx=Math.max(...vals.map(x=>x[1]),1);
  $('#monthlyProgress').innerHTML=vals.map(x=>'<div class="progress-line"><span class="progress-name">'+x[0]+'</span><div class="progress-track"><div class="progress-fill '+x[2]+'" style="width:'+Math.min(100,x[1]/mx*100)+'%"></div></div><span class="progress-value">'+money.format(x[1])+'</span></div>').join('')+
    '<div class="progress-meta"><span>'+Math.round(calendarRatio(state.month)*100)+'% do mês corrido</span><span>Forecast '+(delta>0?'acima':'abaixo')+' do orçamento em '+money.format(Math.abs(delta))+'</span></div>';
  renderAlerts();
  renderOverviewBars();
  const p=latestPatrimony(),d=debtRows().reduce((sum,x)=>sum+x.balance,0),t=thirdPartySummary();
  $('#overviewReserve').innerHTML='<div class="metric-big">'+money.format(p.familyReserve)+'</div><div class="metric-caption">Reserva familiar formal. Liquidez total registrada: '+money.format(p.liquid)+'.</div>';
  $('#overviewDebt').innerHTML='<div class="metric-big">'+money.format(d)+'</div><div class="metric-caption">'+debtRows().length+' obrigações patrimoniais registradas.</div>';
  $('#overviewThirdParty').innerHTML='<div class="metric-big">'+money.format(t.exposure)+'</div><div class="metric-caption">Exposição remanescente em parcelamentos de terceiros.</div>';
}
function renderAlerts(){
  const cats=categoryRows(),alerts=[];
  cats.filter(x=>x.budget>0&&x.forecast>x.budget*1.08&&x.deltaBudget>100).sort((a,b)=>b.deltaBudget-a.deltaBudget).slice(0,4).forEach(x=>alerts.push({tone:'bad',title:x.sub,text:'Forecast '+money.format(x.forecast)+' · '+money.format(x.deltaBudget)+' acima do orçamento.'}));
  cats.filter(x=>x.budget>0&&x.realized>x.budget).sort((a,b)=>b.realized-a.realized).slice(0,2).forEach(x=>alerts.push({tone:'warn',title:x.sub,text:'O realizado já ultrapassou o orçamento antes do fechamento.'}));
  if(!metaLines().length)alerts.push({tone:'neutral',title:'Meta ainda não cadastrada',text:'O código já suporta META por categoria. Vamos defini-la depois de aprovar o orçamento 2027.'});
  if(!alerts.length)alerts.push({tone:'good',title:'Sem alertas relevantes',text:'Nenhuma categoria apresenta desvio material no momento.'});
  $('#alertsPanel').innerHTML=alerts.slice(0,5).map(a=>'<div class="alert '+a.tone+'"><span class="icon">'+(a.tone==='bad'?'!':a.tone==='warn'?'△':a.tone==='good'?'✓':'i')+'</span><div><strong>'+a.title+'</strong><span>'+a.text+'</span></div></div>').join('');
}
function renderOverviewBars(){
  const a=categoryRows().filter(x=>Math.max(x.budget,x.forecast)>0).sort((x,y)=>Math.abs(y.deltaBudget)-Math.abs(x.deltaBudget)).slice(0,8),mx=Math.max(...a.map(x=>Math.max(x.budget,x.forecast)),1);
  $('#overviewCategoryBars').innerHTML=a.map(x=>'<div class="category-bar-row"><span class="category-label">'+x.sub+'</span><div class="dual-track"><span class="budget-mark" style="left:'+Math.min(99,x.budget/mx*100)+'%"></span><div class="forecast-bar" style="width:'+Math.min(100,x.forecast/mx*100)+'%"></div></div><span class="bar-caption '+(x.deltaBudget>0?'bad':'good')+'">'+(x.deltaBudget>0?'+':'')+money.format(x.deltaBudget)+'</span></div>').join('')||'<p class="metric-caption">Sem categorias orçadas para o período.</p>';
}
function renderBudget(){
  const s=monthlySummary(),cats=categoryRows(),meta=s.metaExpense,delta=s.forecastExpense-s.budgetExpense;
  $('#budgetKpis').innerHTML=[
    kpi('Orçado',s.budgetExpense,budgetSource()),
    kpi('Meta',meta==null?'A definir':money.format(meta),'Desafio de eficiência'),
    kpi('Realizado MTD',s.realizedExpense,'Competência'),
    kpi('Forecast',s.forecastExpense,(delta>=0?'+':'')+money.format(delta)+' vs orçamento',delta<=0?'good':delta/s.budgetExpense<=.05?'warn':'bad'),
    kpi('Resultado projetado',s.forecastResult,'Receita menos forecast',s.forecastResult>=0?'good':'bad')
  ].join('');
  $('#pacingNote').textContent='Forecast: fixos/eventos usam valor conhecido; variáveis usam pacing '+(selectedRelation()==='current'?'MTD':'do período')+' por dias corridos ou úteis.';
  $('#budgetTable').innerHTML=cats.map(x=>{
    const st=statusFor(x),db=x.deltaBudget,dm=x.deltaMeta;
    return'<tr><td class="row-title">'+x.sub+'</td><td>'+money.format(x.budget)+'</td><td>'+(x.meta==null?'—':money.format(x.meta))+'</td><td>'+money.format(x.realized)+'</td><td>'+money.format(x.forecast)+'</td><td class="delta '+(db>0?'bad':'good')+'">'+(db>0?'+':'')+money.format(db)+'</td><td class="delta '+(dm==null?'':dm>0?'bad':'good')+'">'+(dm==null?'—':(dm>0?'+':'')+money.format(dm))+'</td><td><span class="table-status '+st+'">'+statusText(st)+'</span></td></tr>';
  }).join('')||'<tr><td colspan="8">Sem orçamento para o mês.</td></tr>';
  const top=cats.filter(x=>Math.max(x.budget,x.forecast,x.realized)>0).slice(0,10),mx=Math.max(...top.map(x=>Math.max(x.budget,x.meta||0,x.realized,x.forecast)),1);
  $('#budgetChart').innerHTML=top.map(x=>'<div class="budget-chart-row"><span class="category-label">'+x.sub+'</span><div class="bars"><div class="tiny-track"><div class="tiny-fill budget" style="width:'+x.budget/mx*100+'%"></div></div>'+(x.meta==null?'':'<div class="tiny-track"><div class="tiny-fill target" style="width:'+x.meta/mx*100+'%"></div></div>')+'<div class="tiny-track"><div class="tiny-fill realized" style="width:'+x.realized/mx*100+'%"></div></div><div class="tiny-track"><div class="tiny-fill forecast" style="width:'+x.forecast/mx*100+'%"></div></div></div><span class="bar-caption">'+money.format(x.forecast)+'</span></div>').join('');
}
function renderPatrimony(){
  const s=patrimonySnapshots(),p=s.at(-1)||{assets:0,liabilities:0,net:0,liquid:0,familyReserve:0,items:[]};
  $('#patrimonyKpis').innerHTML=[
    kpi('Ativos',p.assets,'Valor-base registrado'),
    kpi('Passivos',p.liabilities,'Saldos devedores patrimoniais'),
    kpi('Patrimônio líquido',p.net,'Ativos menos passivos',p.net>=0?'good':'bad'),
    kpi('Liquidez',p.liquid,'Reserva familiar: '+money.format(p.familyReserve))
  ].join('');
  if(s.length<=1)$('#patrimonyEvolution').innerHTML='<div class="metric-caption">Temos apenas um snapshot patrimonial. A evolução surgirá conforme registrarmos fechamentos mensais.</div>';
  else{
    const mx=Math.max(...s.map(x=>Math.abs(x.net)),1);
    $('#patrimonyEvolution').innerHTML=s.map(x=>'<div class="evo-col"><div class="evo-track"><div class="evo-fill" style="height:'+Math.max(4,Math.abs(x.net)/mx*100)+'%"></div></div><div class="evo-value">'+money.format(x.net)+'</div><div class="evo-label">'+shortMonth(x.key)+'</div></div>').join('');
  }
  const assets=p.items.filter(r=>String(r.classe||'').toUpperCase()==='ATIVO'),liab=p.items.filter(r=>String(r.classe||'').toUpperCase()==='PASSIVO');
  $('#patrimonyComposition').innerHTML='<div class="split-box"><h4>Ativos</h4><ul class="item-list">'+assets.map(x=>'<li><span>'+x.item+'</span><strong>'+money.format(x.value)+'</strong></li>').join('')+'</ul></div><div class="split-box"><h4>Passivos</h4><ul class="item-list">'+liab.map(x=>'<li><span>'+x.item+'</span><strong>'+money.format(x.value)+'</strong></li>').join('')+'</ul></div>';
}
function renderDebts(){
  const d=debtRows(),total=d.reduce((s,x)=>s+x.balance,0),monthly=d.reduce((s,x)=>s+x.monthly,0),income=monthlySummary().budgetIncome;
  $('#debtKpis').innerHTML=[kpi('Saldo devedor',total,'Patrimônio registrado'),kpi('Parcelas/mês',monthly,'Obrigações conhecidas'),kpi('Peso na renda',income?pctFmt.format(monthly/income):'—','Parcelas ÷ renda orçada'),kpi('Dívidas registradas',String(d.length),'Imóvel, veículo e empréstimos')].join('');
  $('#debtList').innerHTML=d.map(x=>'<article class="debt-card"><div class="debt-head"><div><p class="eyebrow">'+(labels[x.group]||x.group)+'</p><h3>'+x.name+'</h3></div><span class="debt-balance">'+money.format(x.balance)+'</span></div><div class="debt-meta"><div class="meta-box"><span>Parcela mensal</span><strong>'+(x.monthly?money.format(x.monthly):'A confirmar')+'</strong></div><div class="meta-box"><span>Origem</span><strong>'+(x.origin||'Base Mestre')+'</strong></div><div class="meta-box"><span>Status</span><strong>Ativa</strong></div></div>'+(x.note?'<p class="metric-caption">'+x.note+'</p>':'')+'</article>').join('')||'<p class="metric-caption">Nenhuma dívida registrada.</p>';
}
function renderThirdParty(){
  const t=thirdPartySummary();
  $('#thirdPartyKpis').innerHTML=[kpi('Exposição remanescente',t.exposure,'Obrigações ainda no seu crédito'),kpi('Parcela do mês',t.current,'Compromissos de terceiros no período'),kpi('Reembolsos recebidos',t.reimb,'Caixa do mês'),kpi('Parcelamentos ativos',String(t.rows.length),'Separados do consumo familiar')].join('');
  $('#thirdPartyList').innerHTML=t.rows.map(r=>'<div class="third-row"><div><span class="name">'+(r.descricao||'Parcelamento')+'</span><span class="small-label">'+(r.responsavel_terceiro||'Terceiro')+'</span></div><div><span class="small-label">Parcela</span><span class="small-value">'+money.format(Math.abs(n(r.valor_parcela)))+'</span></div><div><span class="small-label">Restantes</span><span class="small-value">'+remainingInstallments(r)+'</span></div><div><span class="small-label">Exposição</span><span class="small-value">'+money.format(remainingInstallments(r)*Math.abs(n(r.valor_parcela)))+'</span></div></div>').join('')||'<p class="metric-caption">Nenhuma exposição de terceiros registrada.</p>';
}
function renderCash(){
  const c=cashSummary(),cards=cardSummary(),third=thirdPartySummary();
  $('#cashKpis').innerHTML=[kpi('Entradas efetivas',c.inflow,'Inclui reembolsos e outros ingressos'),kpi('Saídas efetivas',c.outflow,'Movimentos externos no caixa'),kpi('Resultado de caixa',c.result,'Entradas menos saídas',c.result>=0?'good':'bad'),kpi('Faturas do período',cards.reduce((s,x)=>s+x.total,0),'Compras com data de caixa no mês'),kpi('Terceiros nas faturas',cards.reduce((s,x)=>s+x.third,0),'Não é consumo familiar')].join('');
  $('#cardsPanel').innerHTML=cards.map(x=>'<div class="card-row"><div><span class="name">'+x.name.replaceAll('_',' ')+'</span><span class="small-label">'+x.count+' lançamentos</span></div><div><span class="small-label">Fatura</span><span class="small-value">'+money.format(x.total)+'</span></div><div><span class="small-label">Família</span><span class="small-value">'+money.format(x.family)+'</span></div><div><span class="small-label">Terceiros</span><span class="small-value">'+money.format(x.third)+'</span></div></div>').join('')||'<p class="metric-caption">Nenhuma fatura detalhada com caixa neste mês.</p>';
  $('#cashPanel').innerHTML='<div class="cash-flow"><div class="cash-box in"><span>Entradas</span><strong>'+money.format(c.inflow)+'</strong></div><div class="cash-box out"><span>Saídas</span><strong>'+money.format(c.outflow)+'</strong></div></div><div class="cash-result"><span class="small-label">Resultado do caixa</span><strong>'+money.format(c.result)+'</strong><div class="metric-caption">Visão operacional. Transferências internas são excluídas; terceiros continuam visíveis porque afetam o dinheiro em conta.</div></div>';
}
function renderMeta(){
  const c=controls(),when=state.data?.importedAt?new Date(state.data.importedAt).toLocaleString('pt-BR'):'—';
  $('#dataStatus').textContent='App '+APP_VERSION+' · Base '+(c.schema_version||'—')+' · '+when;
  $('#budgetSource').textContent=budgetSource();
}
function render(){
  if(!state.data){$('#emptyState').classList.remove('hidden');$('#appShell').classList.add('hidden');return}
  $('#emptyState').classList.add('hidden');$('#appShell').classList.remove('hidden');
  renderMeta();renderOverview();renderBudget();renderPatrimony();renderDebts();renderThirdParty();renderCash();
}
function populateMonths(){
  const a=months(),now=todayIso().slice(0,7);
  if(!state.month||!a.includes(state.month))state.month=a.includes(now)?now:(a.find(x=>x>now)||a.at(-1));
  $('#monthSelect').innerHTML=a.map(m=>'<option value="'+m+'" '+(m===state.month?'selected':'')+'>'+monthLabel(m)+'</option>').join('');
}
function selectPage(page){
  state.page=page;$$('.page').forEach(x=>x.classList.toggle('active',x.dataset.page===page));$$('.tab').forEach(x=>x.classList.toggle('active',x.dataset.page===page));
}
function applyTheme(t){
  state.theme=t;document.documentElement.setAttribute('data-theme',t);localStorage.setItem(THEME,t);
  const b=$('#themeToggle');if(b)b.textContent=t==='dark'?'☀︎ Light':'☾ Dark';
}
async function importFile(file){
  try{toast('Importando Base Mestre…');const d=parseWorkbook(await file.arrayBuffer(),file.name);await dbSet(d);state.data=d;populateMonths();render();toast('Base atualizada.')}
  catch(e){console.error(e);toast(e.message||'Falha ao importar a planilha.')}
}
function bind(){
  $('#fileInput').addEventListener('change',e=>e.target.files?.[0]&&importFile(e.target.files[0]));
  $$('.file-input-mirror').forEach(el=>el.addEventListener('change',e=>e.target.files?.[0]&&importFile(e.target.files[0])));
  $('#monthSelect').addEventListener('change',e=>{state.month=e.target.value;render()});
  $('.tabs').addEventListener('click',e=>{const b=e.target.closest('[data-page]');if(b)selectPage(b.dataset.page)});
  $('#themeToggle').addEventListener('click',()=>applyTheme(state.theme==='dark'?'light':'dark'));
  $('#clearDataBtn').addEventListener('click',async()=>{if(confirm('Remover a cópia local da Base Mestre deste dispositivo?')){await dbClear();state.data=null;render();toast('Base local removida.')}});
}
async function init(){
  applyTheme(localStorage.getItem(THEME)||(matchMedia('(prefers-color-scheme: dark)').matches?'dark':'light'));
  bind();try{state.data=await dbGet()}catch(e){}
  if(state.data)populateMonths();selectPage('overview');render();
  if('serviceWorker'in navigator&&location.protocol.startsWith('http'))navigator.serviceWorker.register('./sw.js').catch(()=>{});
}
document.addEventListener('DOMContentLoaded',init);
})();