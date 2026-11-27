const $=id=>document.getElementById(id);
const fmt=x=>Number(x).toLocaleString('ko-KR',{maximumFractionDigits:3});
let worker, nextId=0, requests=new Map(), sectors=[], rowChoices=[], cells=[], basket=[], result=null;
// ▼ 분류별 엑셀 파일 경로 (data 폴더의 실제 파일명과 같아야 함)
const P='./data/(표)(2023연장)투입산출표_생산자가격_';
const DATASETS={
 basic:{label:'기본부문',file:'./data/io-2023.xlsx'},
 small:{label:'소분류',file:P+'통합소분류.xlsx'},
 middle:{label:'중분류',file:P+'통합중분류.xlsx'},
 large:{label:'대분류',file:P+'통합대분류.xlsx'}
};
let currentKey='basic';
function text(tag,value){const e=document.createElement(tag);e.textContent=value;return e;}
function call(type,payload){return new Promise((resolve,reject)=>{const id=++nextId;requests.set(id,{resolve,reject});worker.postMessage({id,type,payload});});}
function startWorker(){worker=new Worker('./worker.js');worker.onmessage=({data:d})=>{if(d.type==='status'){$('status').textContent=d.text;return;}const p=requests.get(d.id);if(!p)return;requests.delete(d.id);d.ok?p.resolve(d.result):p.reject(Error(d.error));};worker.onerror=e=>{for(const p of requests.values())p.reject(Error(e.message||'Worker 실행 오류'));requests.clear();};}
function options(el,list){el.replaceChildren(...list.map(x=>{const o=document.createElement('option');o.value=x.code;o.textContent=x.code+' '+x.name;return o;}));}
function invalidate(){result=null;$('results').hidden=true;}
async function load(buffer){
 invalidate();basket=[];renderBasket();
 for(const id of ['sheet','row','sector','add','run','row-search','col-search','sector-search','amount'])$(id).disabled=true;
 if(worker)worker.terminate();for(const p of requests.values())p.reject(Error('데이터 다시 로드'));requests.clear();startWorker();
 try{const info=await call('init',buffer);sectors=info.sectors;options($('sector'),sectors);$('sector-search').value='';$('row-search').value='';$('col-search').value='';$('sheet').replaceChildren(...info.sheets.map(s=>{const o=text('option',s);o.value=s;return o;}));
 $('validation').textContent=JSON.stringify(info.residuals,null,2)+'\n최대 절대 잔차: 허용값 1e-8 미만';
 for(const id of ['sheet','row','sector','add','run','row-search','col-search','sector-search','amount'])$(id).disabled=false;
 await explore();loadLinkage();loadDemand();initAnalysis();$('status').textContent='준비 완료 · '+DATASETS[currentKey].label+' · '+sectors.length+'개 부문 · '+info.sheets.length+'개 시트 · 계수 검증 통과';
 }catch(e){$('status').textContent='준비 실패: '+e.message+' — 새로고침 후 재시도하세요.';}
}
let exploreSeq=0;
async function explore(){const seq=++exploreSeq;try{const d=await call('explore',{sheet:$('sheet').value,row:$('row').value});if(seq!==exploreSeq)return;rowChoices=d.rows;filterRows(d.selected);cells=d.cells;$('unit').textContent=d.unit;$('explore-caption').textContent=$('sheet').value+' / 행: '+d.selected+' '+d.name;renderCells();}catch(e){$('status').textContent=e.message;}}
function filterRows(selected){const q=$('row-search').value.trim();options($('row'),rowChoices.filter(r=>(r.code+' '+r.name).includes(q)));if([...$('row').options].some(o=>o.value===selected))$('row').value=selected;}
function renderCells(){const q=$('col-search').value.trim();$('cells').replaceChildren(...cells.filter(c=>(c.code+' '+c.name).includes(q)).map(c=>{const tr=document.createElement('tr');for(const v of [c.code,c.name,c.value===null?'빈 셀':typeof c.value==='number'?fmt(c.value):c.value])tr.append(text('td',v));return tr;}));}
$('sheet').onchange=()=>{$('row-search').value='';$('row').replaceChildren();explore();};$('row').onchange=explore;
$('row-search').oninput=()=>{filterRows();if($('row').value)explore();else{$('cells').replaceChildren();$('explore-caption').textContent='일치하는 행이 없습니다.';}};
$('col-search').oninput=renderCells;
$('sector-search').oninput=()=>{const q=$('sector-search').value.trim();options($('sector'),sectors.filter(s=>(s.code+' '+s.name).includes(q)));};
function renderBasket(){$('basket').replaceChildren(...basket.map((s,i)=>{const li=text('li',s.code+' '+s.name+' / '+fmt(s.amount)+'백만원 ');const b=text('button','삭제');b.type='button';b.onclick=()=>{basket.splice(i,1);invalidate();renderBasket();};li.append(b);return li;}));$('sim-status').textContent=basket.length?basket.length+'개 입력 준비. 계산 버튼을 누르세요.':'시나리오를 추가하세요.';}
$('scenario').onsubmit=e=>{e.preventDefault();const amount=Number($('amount').value),s=sectors.find(s=>s.code===$('sector').value);if(!s||!Number.isFinite(amount)||amount<=0){$('sim-status').textContent='산업과 양의 증가액을 확인하세요.';return;}basket.push({...s,amount});invalidate();renderBasket();};
$('clear').onclick=()=>{basket=[];invalidate();renderBasket();};
$('run').onclick=async()=>{if(!basket.length){$('sim-status').textContent='시나리오를 먼저 추가하세요.';return;}const snapshot=basket.map(s=>({...s}));$('run').disabled=true;$('sim-status').textContent='계산 중…';try{const r=await call('simulate',snapshot);if(JSON.stringify(snapshot)!==JSON.stringify(basket)){$('sim-status').textContent='입력이 변경되었습니다. 다시 계산하세요.';return;}result=r;renderResult();$('sim-status').textContent='계산 완료. 아래 결과는 현재 시나리오에 대응합니다.';}catch(e){invalidate();$('sim-status').textContent='계산 실패: '+e.message;}finally{$('run').disabled=false;}};
function renderResult(){$('results').hidden=false;$('cards').replaceChildren(...[['수요 증가액',result.demand],['생산유발액',result.production],['수입유발액',result.imports],['부가가치유발액',result.valueAdded],...(result.income?[['└ 임금(피용자보수)',result.income.compensation],['└ 영업잉여',result.income.surplus],['└ 고정자본소모',result.income.depreciation],['└ 생산세',result.income.tax]]:[])].map(([title,v])=>{const d=text('div',title);d.className='card';d.append(text('strong',fmt(v)),text('span','백만원'));return d;}));const max=Math.max(...result.rows.map(r=>r.production),1);$('bars').replaceChildren(...result.rows.slice(0,15).map(r=>{const d=text('div',r.name+' · '+fmt(r.production)+'백만원');d.className='bar';const t=document.createElement('div');t.className='track';const f=document.createElement('div');f.className='fill';f.style.width=Math.max(0,r.production/max*100)+'%';t.append(f);d.append(t);return d;}));$('result-rows').replaceChildren(...result.rows.map(r=>{const tr=document.createElement('tr');[r.code,r.name,fmt(r.production),fmt(r.imports),fmt(r.valueAdded)].forEach(v=>tr.append(text('td',v)));return tr;}));}
function csvCell(v){let s=String(v);if(/^[=+@-]/.test(s)&&typeof v!=='number')s="'"+s;return '"'+s.replaceAll('"','""')+'"';}
$('download').onclick=()=>{if(!result)return;const rows=[['산업코드','산업명','생산유발액_백만원','수입유발액_백만원','부가가치유발액_백만원'],...result.rows.map(r=>[r.code,r.name,r.production,r.imports,r.valueAdded])];const url=URL.createObjectURL(new Blob(['\ufeff'+rows.map(r=>r.map(csvCell).join(',')).join('\r\n')],{type:'text/csv;charset=utf-8'}));const a=document.createElement('a');a.href=url;a.download='io-2023-'+currentKey+'-results.csv';a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);};

// ===================== 분석 페이지 3~7 =====================
const esc=s=>String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const f1=x=>Number(x).toLocaleString('ko-KR',{minimumFractionDigits:1,maximumFractionDigits:1});
const f3=x=>Number(x).toLocaleString('ko-KR',{minimumFractionDigits:3,maximumFractionDigits:3});
const host=document.getElementById('simulation')||document.querySelector('main')||document.body;
const WRAP='overflow-x:auto;max-height:480px;overflow-y:auto';
const CARDS='display:flex;flex-wrap:wrap;gap:10px;margin:10px 0';
function mkSection(id,html,after){const s=document.createElement('section');s.id=id;if(host.className)s.className=host.className;s.innerHTML=html;after.insertAdjacentElement('afterend',s);return s;}
function mkTable(headers,rows){const t=document.createElement('table'),th=document.createElement('thead'),tr=document.createElement('tr');headers.forEach(h=>tr.append(text('th',h)));th.append(tr);const tb=document.createElement('tbody');rows.forEach(r=>{const x=document.createElement('tr');r.forEach(v=>x.append(text('td',v)));tb.append(x);});t.append(th,tb);return t;}
function barRow(label,value,max,unit){const d=text('div',label+' · '+fmt(value)+unit);d.className='bar';const t=document.createElement('div');t.className='track';const f=document.createElement('div');f.className='fill';f.style.width=Math.max(0,Math.min(100,max>0?value/max*100:0))+'%';t.append(f);d.append(t);return d;}
function cardEl(title,value,unit){const d=text('div',title);d.className='card';d.append(text('strong',value),text('span',unit));return d;}
const secName=c=>{const s=sectors.find(x=>x.code===c);return s?s.name:c;};
const pickers={};
function pickerHTML(p){return '<input id="'+p+'-search" type="search" placeholder="산업 검색" disabled> <select id="'+p+'-sector" disabled></select>';}
function wirePicker(p,extra){const get=()=>{let e=typeof extra==='function'?extra():extra;if(e===true)e={value:'ALL',label:'전체 산업 (전 부문 일괄)'};return e||null;};const fill=()=>{const q=$(p+'-search').value.trim();options($(p+'-sector'),sectors.filter(s=>(s.code+' '+s.name).includes(q)));const e=get();if(e){const o=text('option',e.label);o.value=e.value;$(p+'-sector').prepend(o);$(p+'-sector').selectedIndex=0;}};pickers[p]=fill;$(p+'-search').oninput=fill;}
function enablePicker(p){$(p+'-search').value='';pickers[p]();$(p+'-search').disabled=false;$(p+'-sector').disabled=false;}

// 페이지 이동 메뉴
const nav=document.createElement('nav');nav.style.cssText='margin:16px 0;display:flex;flex-wrap:wrap;gap:12px';
nav.innerHTML='<strong>추가 분석</strong><a href="#linkage">3 연쇄효과</a><a href="#demand">4 수요 구조</a><a href="#public">5 공공투자 평가</a><a href="#supply">6 공급망·수입대체</a><a href="#paths">7 파급경로</a><a href="#price">8 가격 파급</a><a href="#extract">9 가상 제거법</a><a href="#scorecard">10 스코어카드</a>';
host.insertAdjacentElement('afterend',nav);

// ---------- 3. 연쇄효과 ----------
let linkRows=[];
const linkSec=mkSection('linkage',`
<h2>3. 산업 연쇄효과 분석</h2>
<p>생산유발계수 역행렬의 열합·행합으로 산출합니다. 영향력계수(후방연쇄)와 감응도계수(전방연쇄)는 전 산업 평균을 1로 한 상대값입니다(단위 없음). 변동계수는 파급효과가 특정 산업에 쏠린 정도로, 클수록 몇몇 산업에만 효과가 집중됩니다. 유발계수 합은 최종수요 1원당 유발되는 금액(원/원)입니다.</p>
<div id="link-plot"></div>
<div>
 <input id="link-search" type="search" placeholder="산업 검색" disabled>
 <select id="link-sort" disabled>
  <option value="influence">영향력계수 높은 순</option><option value="sensitivity">감응도계수 높은 순</option>
  <option value="production">생산유발계수 합 높은 순</option><option value="valueAdded">부가가치유발계수 합 높은 순</option>
  <option value="imports">수입유발계수 합 높은 순</option><option value="importRatio">직접 수입의존도 높은 순</option>
  <option value="vaRatio">직접 부가가치율 높은 순</option><option value="influenceCV">후방 변동계수 높은 순</option>
 </select>
 <select id="link-type" disabled><option value="">전체 유형</option><option>중추 산업</option><option>최종수요 지향</option><option>중간수요 지향</option><option>독립 산업</option></select>
</div>
<p id="link-status">분류를 불러오면 표시됩니다.</p>
<div style="${WRAP}"><table><thead><tr><th>코드</th><th>산업</th><th>영향력계수</th><th>감응도계수</th><th>후방 변동계수</th><th>전방 변동계수</th><th>생산유발계수 합</th><th>부가가치유발계수 합</th><th>수입유발계수 합</th><th>직접 부가가치율</th><th>직접 수입의존도</th><th>유형</th></tr></thead><tbody id="link-rows"></tbody></table></div>`,nav);
const TYPE_COLOR={'중추':'#c0392b','최종':'#2e86c1','중간':'#27ae60','독립':'#8e8e93'};
function renderScatter(){
 if(!linkRows.length){$('link-plot').replaceChildren();return;}
 const W=640,H=420,P=48,xs=linkRows.map(r=>r.influence),ys=linkRows.map(r=>r.sensitivity);
 const x0=Math.min(0.7,...xs)-0.03,x1=Math.max(1.3,...xs)+0.03,y0=Math.min(0.2,...ys)-0.03,y1=Math.max(1.3,...ys)+0.03;
 const sx=v=>P+(v-x0)/(x1-x0)*(W-2*P),sy=v=>H-P-(v-y0)/(y1-y0)*(H-2*P);
 const lab=linkRows.length<=60;
 let s='<svg viewBox="0 0 '+W+' '+H+'" style="width:100%;max-width:680px;height:auto" role="img" aria-label="영향력계수와 감응도계수 산점도">';
 s+='<line x1="'+sx(1)+'" y1="'+P/2+'" x2="'+sx(1)+'" y2="'+(H-P)+'" stroke="currentColor" opacity=".35" stroke-dasharray="4 3"/>';
 s+='<line x1="'+P+'" y1="'+sy(1)+'" x2="'+(W-P/2)+'" y2="'+sy(1)+'" stroke="currentColor" opacity=".35" stroke-dasharray="4 3"/>';
 s+='<rect x="'+P+'" y="'+P/2+'" width="'+(W-1.5*P)+'" height="'+(H-1.5*P)+'" fill="none" stroke="currentColor" opacity=".25"/>';
 s+='<text x="'+(W-P-4)+'" y="'+(P/2+14)+'" text-anchor="end" font-size="12" fill="currentColor" opacity=".7">중추 산업</text>';
 s+='<text x="'+(P+6)+'" y="'+(P/2+14)+'" font-size="12" fill="currentColor" opacity=".7">중간수요 지향</text>';
 s+='<text x="'+(W-P-4)+'" y="'+(H-P-6)+'" text-anchor="end" font-size="12" fill="currentColor" opacity=".7">최종수요 지향</text>';
 s+='<text x="'+(P+6)+'" y="'+(H-P-6)+'" font-size="12" fill="currentColor" opacity=".7">독립 산업</text>';
 s+='<text x="'+(W/2)+'" y="'+(H-10)+'" text-anchor="middle" font-size="12" fill="currentColor">영향력계수 (후방연쇄) →</text>';
 s+='<text transform="translate(14 '+(H/2)+') rotate(-90)" text-anchor="middle" font-size="12" fill="currentColor">감응도계수 (전방연쇄) →</text>';
 for(const r of linkRows){const c=TYPE_COLOR[r.type.slice(0,2)]||'#8e8e93';s+='<circle cx="'+sx(r.influence).toFixed(1)+'" cy="'+sy(r.sensitivity).toFixed(1)+'" r="'+(lab?5:3.5)+'" fill="'+c+'" fill-opacity=".75"><title>'+esc(r.code+' '+r.name+' · 영향력 '+f3(r.influence)+' / 감응도 '+f3(r.sensitivity))+'</title></circle>';
  if(lab)s+='<text x="'+(sx(r.influence)+7).toFixed(1)+'" y="'+(sy(r.sensitivity)+4).toFixed(1)+'" font-size="9" fill="currentColor" opacity=".75">'+esc(r.code)+'</text>';}
 s+='</svg><p>점에 마우스를 올리면 산업명이 보입니다. 점선은 평균(1)입니다.</p>';
 $('link-plot').innerHTML=s;
}
function renderLinkage(){
 const q=$('link-search').value.trim(),t=$('link-type').value,k=$('link-sort').value;
 const list=linkRows.filter(r=>(r.code+' '+r.name).includes(q)&&(!t||r.type.startsWith(t))).sort((x,y)=>y[k]-x[k]);
 $('link-rows').replaceChildren(...list.map(r=>{const tr=document.createElement('tr');[r.code,r.name,f3(r.influence),f3(r.sensitivity),f3(r.influenceCV),f3(r.sensitivityCV),f3(r.production),f3(r.valueAdded),f3(r.imports),f3(r.vaRatio),f3(r.importRatio),r.type].forEach(v=>tr.append(text('td',v)));return tr;}));
 $('link-status').textContent=list.length+'개 산업 표시 (전체 '+linkRows.length+'개). 영향력·감응도계수 1 이상이면 평균 이상입니다.';
}
async function loadLinkage(){
 for(const id of ['link-search','link-sort','link-type'])$(id).disabled=true;
 $('link-status').textContent='계산 중…';
 try{linkRows=(await call('linkage')).rows;renderScatter();renderLinkage();renderImportRank();renderExtRank();for(const id of ['link-search','link-sort','link-type'])$(id).disabled=false;}
 catch(e){$('link-status').textContent='연쇄효과 계산 실패: '+e.message;}
}
$('link-search').oninput=renderLinkage;$('link-sort').onchange=renderLinkage;$('link-type').onchange=renderLinkage;

// ---------- 4. 최종수요 구조 ----------
let demData=null,demAll=[],demSel=null;
const SHORT={'9111':'민간소비','9112':'정부소비','9121':'민간투자','9122':'정부투자','9131':'재고증감','9132':'귀중품','9140':'수출'};
const DEMLAB={'9111':'민간소비 주도','9112':'정부소비 주도','9121':'민간투자 주도','9122':'정부투자 주도','9140':'수출 주도',other:'재고·기타 주도'};
const itemName=it=>SHORT[it.code]||it.name||it.code;
const demSec=mkSection('demand',`
<h2>4. 최종수요 구조 분석 — 무엇이 이 산업을 먹여 살리는가</h2>
<p>각 산업의 생산이 민간소비·정부소비·민간투자·정부투자·수출 중 무엇에 의해 유발되는지 보여 줍니다(원자료 '최종수요 항목별 생산유발액' 시트, 단위 백만원). 정부소비와 정부투자를 합친 <b>공공수요 비중</b>은 공공부문 지출에 대한 민감도를 가늠하는 지표입니다.</p>
<h3>전체 경제: 수요 항목별 유발 비중</h3>
<p id="dem-insight"></p>
<div id="dem-eco" style="${WRAP}"></div>
<h3>산업별 수요 구조</h3>
<div id="dem-chips">공기업 관련 산업 바로가기: </div>
<div>
 <input id="dem-search" type="search" placeholder="산업 검색" disabled>
 <select id="dem-sort" disabled><option value="total">생산유발액 큰 순</option><option value="export">수출 비중 높은 순</option><option value="public">공공수요 비중 높은 순</option><option value="cons">민간소비 비중 높은 순</option><option value="invest">투자 비중 높은 순</option></select>
 <select id="dem-type" disabled><option value="">전체 유형</option>${Object.values(DEMLAB).map(v=>'<option>'+v+'</option>').join('')}</select>
</div>
<p id="dem-status">분류를 불러오면 표시됩니다.</p>
<div style="${WRAP}"><table><thead><tr><th>코드</th><th>산업</th><th>생산유발액(백만원)</th><th>민간소비 %</th><th>정부소비 %</th><th>민간투자 %</th><th>정부투자 %</th><th>수출 %</th><th>공공수요 %</th><th>주도 수요</th></tr></thead><tbody id="dem-rows"></tbody></table></div>
<h3 id="dem-detail-title">행을 클릭하면 상세 구성이 표시됩니다</h3><div id="dem-detail"></div>`,linkSec);
for(const [lab,kw] of [['전력','전력'],['가스','가스'],['건설','건설'],['운송','운송'],['수도','수도'],['보건','보건'],['통신','통신']]){const b=text('button',lab);b.type='button';b.onclick=()=>{$('dem-search').value=kw;renderDemand();};$('dem-chips').append(b,' ');}
const demClear=text('button','전체');demClear.type='button';demClear.onclick=()=>{$('dem-search').value='';renderDemand();};$('dem-chips').append(demClear);
function demCompute(){return demData.rows.map(r=>{const t=r.total>0?r.total:1,sh=c=>(r.by[c]||0)/t*100;const g={'9111':sh('9111'),'9112':sh('9112'),'9121':sh('9121'),'9122':sh('9122'),'9140':sh('9140'),other:((r.by['9131']||0)+(r.by['9132']||0))/t*100};let lead='other',mx=-1e18;for(const k in g)if(g[k]>mx){mx=g[k];lead=k;}return {...r,sh:g,pub:g['9112']+g['9122'],inv:g['9121']+g['9122'],type:DEMLAB[lead]};});}
function renderEco(){
 const eco=demData.economy||{},p=eco.production;
 if(!p){$('dem-eco').textContent='전체 경제 구성 자료가 없습니다.';$('dem-insight').textContent='';return;}
 const sum=o=>Object.values(o||{}).reduce((a,b)=>a+b,0),sp=sum(p),si=sum(eco.imports),sv=sum(eco.valueAdded);
 const rows=demData.items.filter(it=>p[it.code]!==undefined).map(it=>[itemName(it),fmt(p[it.code]),f1(p[it.code]/sp*100)+'%',si?f1((eco.imports[it.code]||0)/si*100)+'%':'-',sv?f1((eco.valueAdded[it.code]||0)/sv*100)+'%':'-']);
 $('dem-eco').replaceChildren(mkTable(['수요 항목','생산유발액(백만원)','생산유발 비중','수입유발 비중','부가가치유발 비중'],rows));
 const ex=demData.items.find(it=>it.code==='9140');
 $('dem-insight').textContent=(ex&&si&&sv)?'수출은 전체 생산유발의 '+f1(p['9140']/sp*100)+'%를 차지하지만, 수입유발에서는 '+f1((eco.imports['9140']||0)/si*100)+'%, 부가가치유발에서는 '+f1((eco.valueAdded['9140']||0)/sv*100)+'%를 차지합니다. 비중이 어긋나는 항목일수록 수입 의존도나 부가가치율이 다르다는 뜻입니다.':'';
}
function renderDemand(){
 if(!demAll.length)return;
 const q=$('dem-search').value.trim(),t=$('dem-type').value,k=$('dem-sort').value;
 const key={total:r=>r.total,export:r=>r.sh['9140'],public:r=>r.pub,cons:r=>r.sh['9111'],invest:r=>r.inv}[k];
 const list=demAll.filter(r=>(r.code+' '+r.name).includes(q)&&(!t||r.type===t)).sort((a,b)=>key(b)-key(a));
 $('dem-rows').replaceChildren(...list.map(r=>{const tr=document.createElement('tr');tr.style.cursor='pointer';if(r.code===demSel)tr.style.background='rgba(127,127,127,.2)';[r.code,r.name,fmt(r.total),f1(r.sh['9111']),f1(r.sh['9112']),f1(r.sh['9121']),f1(r.sh['9122']),f1(r.sh['9140']),f1(r.pub),r.type].forEach(v=>tr.append(text('td',v)));tr.onclick=()=>{demSel=r.code;renderDemand();renderDemDetail();};return tr;}));
 $('dem-status').textContent=list.length+'개 산업 표시 (전체 '+demAll.length+'개). 행을 클릭하면 아래에 상세 구성이 나옵니다.';
}
function renderDemDetail(){
 const r=demAll.find(x=>x.code===demSel);if(!r)return;
 $('dem-detail-title').textContent=r.code+' '+r.name+' — 수요 항목별 생산유발액 (백만원)';
 const mx=Math.max(...demData.items.map(it=>r.by[it.code]||0),1);
 $('dem-detail').replaceChildren(...demData.items.filter(it=>(r.by[it.code]||0)!==0).map(it=>barRow(itemName(it)+' ('+f1((r.by[it.code]||0)/(r.total>0?r.total:1)*100)+'%)',r.by[it.code]||0,mx,'백만원')),text('p','공공수요(정부소비+정부투자) 비중 '+f1(r.pub)+'% · '+r.type));
}
async function loadDemand(){
 demData=null;demAll=[];demSel=null;$('dem-status').textContent='계산 중…';$('dem-rows').replaceChildren();$('dem-detail').replaceChildren();
 try{demData=await call('demand');demAll=demCompute();renderEco();renderDemand();for(const id of ['dem-search','dem-sort','dem-type'])$(id).disabled=false;}
 catch(e){$('dem-status').textContent='수요 구조 계산 실패: '+e.message;}
}
$('dem-search').oninput=renderDemand;$('dem-sort').onchange=renderDemand;$('dem-type').onchange=renderDemand;

// ---------- 5. 공공투자 평가 ----------
let pubBasket=[],pubRes=null;
const pubSec=mkSection('public',`
<h2>5. 공공투자 사업 경제효과 비교</h2>
<p>사업을 해당 산업의 최종수요 증가로 보고 생산·부가가치·수입 유발액을 구합니다. 부가가치는 피용자보수(임금) 등으로 나눠 볼 수 있어 <b>임금으로 돌아가는 몫</b>까지 비교합니다(고용 자료가 없어 일자리 수 대신 노동소득을 사용). 예산이 달라도 공정하게 비교하도록 <b>예산 1조원(=1,000,000백만원)당</b> 값으로 환산합니다.</p>
<div id="pub-chips">바로 추가(예산 1조원): </div>
<form id="pub-form">${pickerHTML('pub')} <input id="pub-amount" type="number" min="1" step="1" value="1000000" disabled> 백만원 <button id="pub-add" type="submit" disabled>사업 추가</button> <button id="pub-clear" type="button">비우기</button></form>
<ul id="pub-list"></ul>
<button id="pub-run" type="button" disabled>경제효과 비교 계산</button>
<p id="pub-status">사업을 추가하세요.</p>
<div id="pub-out"></div>`,demSec);
wirePicker('pub',false);
for(const [lab,kw] of [['SOC·건설','건설'],['전력','전력'],['가스','가스'],['운송·물류','운송'],['수도·환경','수도'],['보건·의료','보건'],['정보통신','통신']]){const b=text('button',lab);b.type='button';b.onclick=()=>{const s=sectors.find(x=>x.name.includes(kw));if(!s){$('pub-status').textContent='"'+kw+'"가 들어간 산업이 이 분류에 없습니다. 직접 검색해 추가하세요.';return;}pubBasket.push({...s,amount:1000000});pubRes=null;$('pub-out').replaceChildren();renderPubList();};$('pub-chips').append(b,' ');}
function renderPubList(){
 $('pub-list').replaceChildren(...pubBasket.map((s,i)=>{const li=text('li',s.code+' '+s.name+' / '+fmt(s.amount)+'백만원 ');const b=text('button','삭제');b.type='button';b.onclick=()=>{pubBasket.splice(i,1);pubRes=null;$('pub-out').replaceChildren();renderPubList();};li.append(b);return li;}));
 $('pub-status').textContent=pubBasket.length?pubBasket.length+'개 사업. 계산 버튼을 누르세요.':'사업을 추가하세요.';
}
$('pub-form').onsubmit=e=>{e.preventDefault();const a=Number($('pub-amount').value),s=sectors.find(x=>x.code===$('pub-sector').value);if(!s||!Number.isFinite(a)||a<=0){$('pub-status').textContent='산업과 양의 예산을 확인하세요.';return;}pubBasket.push({...s,amount:a});pubRes=null;$('pub-out').replaceChildren();renderPubList();};
$('pub-clear').onclick=()=>{pubBasket=[];pubRes=null;$('pub-out').replaceChildren();renderPubList();};
$('pub-run').onclick=async()=>{
 if(!pubBasket.length){$('pub-status').textContent='사업을 먼저 추가하세요.';return;}
 $('pub-run').disabled=true;$('pub-status').textContent='계산 중…';
 try{const r=await call('evaluate',pubBasket.map(s=>({code:s.code,name:s.name,amount:s.amount})));pubRes=r.rows;renderPubResult();$('pub-status').textContent='계산 완료.';}
 catch(e){$('pub-status').textContent='계산 실패: '+e.message;}finally{$('pub-run').disabled=false;}
};
function renderPubResult(){
 const rows=pubRes,hasW=rows.every(r=>r.compensation!==null),o=$('pub-out'),nz=v=>v===null?'-':fmt(v);
 const tbl=mkTable(['사업','예산(백만원)','생산유발액','부가가치유발액','임금(피용자보수)유발액','수입유발액','[1조원당] 생산','[1조원당] 부가가치','[1조원당] 임금','[1조원당] 수입'],
  rows.map(r=>[r.code+' '+r.name,fmt(r.amount),fmt(r.production),fmt(r.valueAdded),nz(r.compensation),fmt(r.imports),fmt(r.per.production),fmt(r.per.valueAdded),nz(r.per.compensation),fmt(r.per.imports)]));
 const w=document.createElement('div');w.style.cssText=WRAP;w.append(tbl);
 const out=[w];
 const mxV=Math.max(...rows.map(r=>r.per.valueAdded),1);
 out.push(text('h3','예산 1조원당 부가가치유발액 (백만원)'),...rows.map(r=>barRow(r.code+' '+r.name,r.per.valueAdded,mxV,'백만원')));
 if(hasW){const mxW=Math.max(...rows.map(r=>r.per.compensation),1);out.push(text('h3','예산 1조원당 임금(피용자보수)유발액 (백만원)'),...rows.map(r=>barRow(r.code+' '+r.name,r.per.compensation,mxW,'백만원')));}
 const bv=[...rows].sort((a,b)=>b.per.valueAdded-a.per.valueAdded)[0],bi=[...rows].sort((a,b)=>a.per.imports-b.per.imports)[0];
 out.push(text('p','1조원당 부가가치유발액이 가장 큰 사업은 '+bv.name+'이고, 수입유발액이 가장 작은(국내 환류가 큰) 사업은 '+bi.name+'입니다. 승수는 투입계수가 고정된 선형 모형의 결과이며 공급 제약과 가격 변화는 반영하지 않습니다.'));
 o.replaceChildren(...out);
}

// ---------- 6. 공급망·수입대체 ----------
const supSec=mkSection('supply',`
<h2>6. 공급망 안정성: 수입대체 시뮬레이션</h2>
<p>선택한 산업(또는 전 산업)이 쓰는 <b>수입 중간재의 일부를 같은 산업의 국산품으로 바꿨을 때</b>, 현재 국산 최종수요(국산거래표 최종수요계)를 충족하기 위한 생산·부가가치·수입이 어떻게 변하는지 계산합니다. 가격·품질이 같고 국내 공급여력이 충분하다고 가정하므로, 증가율이 매우 큰 산업은 현실적 공급 가능성을 따로 확인해야 합니다.</p>
<form id="sup-form">${pickerHTML('sup')} 대체 비율 <input id="sup-ratio" type="number" min="1" max="100" step="1" value="10" disabled>% <button id="sup-run" type="submit" disabled>계산</button></form>
<p id="sup-status">분류를 불러오면 사용할 수 있습니다.</p>
<div id="sup-out"></div>
<h3>수입 의존도가 높은 산업 TOP 10 (직접 수입의존도 = 투입 중 수입 중간재 비율)</h3>
<div id="sup-rank" style="${WRAP}"></div>`,pubSec);
wirePicker('sup',true);
function renderImportRank(){
 if(!linkRows.length)return;
 const top=[...linkRows].sort((a,b)=>b.importRatio-a.importRatio).slice(0,10);
 $('sup-rank').replaceChildren(mkTable(['코드','산업','직접 수입의존도','수입유발계수 합','부가가치유발계수 합'],top.map(r=>[r.code,r.name,f3(r.importRatio),f3(r.imports),f3(r.valueAdded)])));
}
$('sup-form').onsubmit=async e=>{
 e.preventDefault();const code=$('sup-sector').value,ratio=Number($('sup-ratio').value);
 if(!code||!Number.isFinite(ratio)||ratio<=0||ratio>100){$('sup-status').textContent='산업과 1~100 사이 대체 비율을 확인하세요.';return;}
 $('sup-run').disabled=true;$('sup-status').textContent='계산 중…';
 try{
  const r=await call('substitute',{code,ratio});
  const dP=r.production1-r.production0,dV=r.valueAdded1-r.valueAdded0,dI=r.imports0-r.imports1;
  const w=document.createElement('div');w.style.cssText=WRAP;
  w.append(mkTable(['산업','기준 생산(백만원)','증가액(백만원)','증가율(%)','비고'],r.rows.map(x=>{const g=x.base>0?x.delta/x.base*100:0;return [x.code+' '+x.name,fmt(x.base),fmt(x.delta),f1(g),g>=50?'공급여력 확인 필요':''];})));
  $('sup-out').replaceChildren(
   (()=>{const d=document.createElement('div');d.style.cssText=CARDS;d.append(cardEl('대체 대상 수입 중간재',fmt(r.replaceable),'백만원'),cardEl('생산 증가',fmt(dP),'백만원 (+'+f1(dP/r.production0*100)+'%)'),cardEl('부가가치 증가',fmt(dV),'백만원 (+'+f1(dV/r.valueAdded0*100)+'%)'),cardEl('수입 감소',fmt(dI),'백만원 (-'+f1(dI/r.imports0*100)+'%)'));return d;})(),
   text('h3','생산이 가장 크게 늘어나는 산업'),w);
  $('sup-status').textContent=(code==='ALL'?'전 산업':secName(code))+' 수입 중간재 '+r.ratio+'% 국산 대체 결과입니다.';
 }catch(err){$('sup-status').textContent='계산 실패: '+err.message;}finally{$('sup-run').disabled=false;}
};

// ---------- 7. 파급경로 추적 ----------
const pathSec=mkSection('paths',`
<h2>7. 파급경로 추적: 수요가 퍼지는 단계</h2>
<p>한 산업에 수요가 생기면 먼저 그 산업이 생산하고(0차), 필요한 중간재를 만드는 산업이 생산하며(1차), 그 산업들의 중간재를 만드는 산업이 다시 생산하는(2차) 식으로 퍼집니다. 총 생산유발액 = 0차 + 1차 + 2차 + … (역행렬의 급수 전개)로, 단계별 크기와 주요 산업을 보여 줍니다.</p>
<form id="path-form">${pickerHTML('path')} <input id="path-amount" type="number" min="1" step="1" value="1000" disabled> 백만원 <button id="path-run" type="submit" disabled>경로 분해</button></form>
<p id="path-status">분류를 불러오면 사용할 수 있습니다.</p>
<div id="path-out"></div>`,supSec);
wirePicker('path',false);
$('path-form').onsubmit=async e=>{
 e.preventDefault();const code=$('path-sector').value,amount=Number($('path-amount').value);
 if(!code||!Number.isFinite(amount)||amount<=0){$('path-status').textContent='산업과 양의 수요 증가액을 확인하세요.';return;}
 $('path-run').disabled=true;$('path-status').textContent='계산 중…';
 try{
  const r=await call('paths',{code,amount});
  const mx=Math.max(...r.rounds.map(x=>x.amount),1);
  const lst=(arr)=>{const w=document.createElement('div');w.style.cssText=WRAP;w.append(mkTable(['코드','산업','생산유발액(백만원)'],arr.map(x=>[x.code,x.name,fmt(x.amount)])));return w;};
  const d=document.createElement('div');d.style.cssText=CARDS;d.append(cardEl('수요 증가액',fmt(r.demand),'백만원'),cardEl('총 생산유발액',fmt(r.total),'백만원'),cardEl('생산승수',f3(r.total/r.demand),'배'));
  $('path-out').replaceChildren(d,text('h3','단계별 생산유발액'),...r.rounds.map(x=>barRow(x.label+' ('+f1(x.amount/r.total*100)+'%)',x.amount,mx,'백만원')),
   text('h3','1차: 직접 투입을 공급하는 상위 산업'),lst(r.top1),text('h3','2차: 투입의 투입을 공급하는 상위 산업'),lst(r.top2),text('h3','전체 단계 합산 상위 산업'),lst(r.topAll));
  $('path-status').textContent=secName(code)+' 수요 '+fmt(amount)+'백만원의 파급경로입니다.';
 }catch(err){$('path-status').textContent='계산 실패: '+err.message;}finally{$('path-run').disabled=false;}
};

// ---------- 8. 요금·원자재 가격 파급 ----------
const priceSec=mkSection('price',`
<h2>8. 요금·원자재 가격 파급 시뮬레이터</h2>
<p>전기·가스·수도 요금이 오르거나 환율·국제 원자재 가격 때문에 수입재 가격이 오르면, 산업 비용과 물가가 얼마나 오를까요? 투입산출표의 <b>가격 모형(Leontief price model)</b>으로 계산합니다. 오른 비용이 하류 산업의 가격에 <b>모두 전가</b>된다고 가정한 상한 추정이며, 수요 변화는 반영하지 않습니다.</p>
<div id="pr-chips">바로 실행: </div>
<form id="pr-form">
 <select id="pr-mode" disabled><option value="domestic">국내 산업 가격(요금) 인상</option><option value="import">수입재 가격 상승 (환율·원자재)</option></select>
 ${pickerHTML('pr')} <input id="pr-pct" type="number" min="0.1" step="0.1" value="10" disabled>% <button id="pr-run" type="submit" disabled>계산</button>
</form>
<p id="pr-status">분류를 불러오면 사용할 수 있습니다.</p>
<div id="pr-out"></div>`,pathSec);
wirePicker('pr',()=>$('pr-mode').value==='import');
$('pr-mode').onchange=()=>pickers.pr();
async function runPrice(){
 const mode=$('pr-mode').value,code=$('pr-sector').value,pct=Number($('pr-pct').value);
 if(!code||!Number.isFinite(pct)||pct<=0){$('pr-status').textContent='산업과 양의 상승률을 확인하세요.';return;}
 $('pr-run').disabled=true;$('pr-status').textContent='계산 중…';
 try{
  const r=await call('price',{mode,code,pct});
  const what=mode==='domestic'?secName(code)+' 가격 '+r.pct+'% 인상':(code==='ALL'?'모든 수입재':secName(code)+' 수입가')+' '+r.pct+'% 상승';
  const d=document.createElement('div');d.style.cssText=CARDS;
  d.append(cardEl('국내 최종수요 물가 상승',r.index===null?'-':f3(r.index),'%'),cardEl('산업 평균 가격(비용) 상승',f3(r.mean),'%'),cardEl('가장 큰 산업 상승',f3(r.max),'%'));
  const hasSq=r.rows.some(x=>x.squeeze!==null);
  const w=document.createElement('div');w.style.cssText=WRAP;
  w.append(mkTable(['산업','가격 상승률(%)','1차 비용상승률(%)'].concat(hasSq?['가격 못 올리면 영업잉여 감소율(%)']:[]),r.rows.map(x=>[x.code+' '+x.name,f3(x.dp),f3(x.direct)].concat(hasSq?[x.squeeze===null?'-':f1(x.squeeze)]:[]))));
  const mx=Math.max(...r.rows.map(x=>x.dp),0.0001);
  $('pr-out').replaceChildren(d,text('h3','가격이 가장 크게 오르는 산업 (상위 15개)'),...r.rows.map(x=>barRow(x.code+' '+x.name,x.dp,mx,'%')),text('h3','상세'),w,
   text('p','해석: 1차 비용상승률은 충격이 처음 닿는 직접 비용 증가분이고, 가격 상승률은 하류로 전가가 끝난 뒤의 값입니다. 마지막 열은 해당 산업이 가격을 올리지 못할 때 1차 비용 증가가 영업잉여(이윤)를 몇 % 줄이는지의 근사입니다. 값이 큰 산업은 요금·원자재 충격에 수익성이 취약합니다.'));
  $('pr-status').textContent=what+' 시나리오 결과입니다.';
 }catch(e){$('pr-status').textContent='계산 실패: '+e.message;}finally{$('pr-run').disabled=false;}
}
$('pr-form').onsubmit=e=>{e.preventDefault();runPrice();};
for(const [lab,mode,kw,pct] of [['전기요금 +10%','domestic','전력',10],['도시가스 +10%','domestic','가스',10],['수도요금 +10%','domestic','수도',10],['원유·광산품 수입가 +20%','import','광산',20],['환율 +10% (모든 수입재)','import','ALL',10]]){
 const b=text('button',lab);b.type='button';
 b.onclick=()=>{$('pr-mode').value=mode;$('pr-search').value='';pickers.pr();if(kw==='ALL'){$('pr-sector').value='ALL';}else{const s=sectors.find(x=>x.name.includes(kw))||(kw==='광산'?sectors.find(x=>x.name.includes('원유')):null);if(!s){$('pr-status').textContent='"'+kw+'"가 들어간 산업이 이 분류에 없습니다. 직접 검색하세요.';return;}$('pr-sector').value=s.code;}$('pr-pct').value=pct;runPrice();};
 $('pr-chips').append(b,' ');
}

// ---------- 9. 가상 제거법 ----------
const extSec=mkSection('extract',`
<h2>9. 가상 제거법 — 이 산업이 사라진다면?</h2>
<p>한 산업이 완전히 사라져 생산도 투입도 할 수 없다고 가정하고, 나머지 산업이 남은 국산 최종수요를 맞추도록 다시 균형을 잡았을 때 사라지는 총산출·부가가치의 비중입니다. 영향력계수(후방)와 감응도계수(전방)가 각각 한쪽 방향만 본다면, 이 지표는 <b>양쪽을 한꺼번에 반영한 구조적 중요도</b>입니다. 손실에는 그 산업 자신의 산출도 포함됩니다.</p>
<div><input id="ext-search" type="search" placeholder="산업 검색" disabled></div>
<p id="ext-status">분류를 불러오면 표시됩니다. 행을 클릭하면 어떤 산업이 가장 크게 타격받는지 나옵니다.</p>
<div style="${WRAP}"><table><thead><tr><th>순위</th><th>코드</th><th>산업</th><th>총산출 손실(%)</th><th>부가가치 손실(%)</th><th>영향력계수</th><th>감응도계수</th><th>유형</th></tr></thead><tbody id="ext-rows"></tbody></table></div>
<h3 id="ext-title"></h3><div id="ext-out"></div>`,priceSec);
let extSel=null;
function renderExtRank(){
 const base=linkRows.filter(r=>r.extOutput!==null&&r.extOutput!==undefined).sort((a,b)=>b.extOutput-a.extOutput);
 if(!base.length){$('ext-status').textContent='이 파일에서 국산 최종수요를 읽지 못해 가상 제거법을 계산할 수 없습니다.';$('ext-rows').replaceChildren();return;}
 const q=$('ext-search').value.trim();
 const rank=new Map(base.map((r,i)=>[r.code,i+1]));
 const list=base.filter(r=>(r.code+' '+r.name).includes(q));
 $('ext-rows').replaceChildren(...list.map(r=>{const tr=document.createElement('tr');tr.style.cursor='pointer';if(r.code===extSel)tr.style.background='rgba(127,127,127,.2)';[rank.get(r.code),r.code,r.name,f3(r.extOutput),f3(r.extVA),f3(r.influence),f3(r.sensitivity),r.type].forEach(v=>tr.append(text('td',v)));tr.onclick=()=>{extSel=r.code;renderExtRank();loadExtraction(r.code);};return tr;}));
 $('ext-status').textContent=list.length+'개 산업 표시. 행을 클릭하면 아래에 상세 영향이 나옵니다.';
 $('ext-search').disabled=false;
}
$('ext-search').oninput=renderExtRank;
async function loadExtraction(code){
 $('ext-title').textContent=secName(code)+' 제거 시 영향';$('ext-out').textContent='계산 중…';
 try{
  const r=await call('extraction',{code});
  const d=document.createElement('div');d.style.cssText=CARDS;
  d.append(cardEl('총산출 손실',fmt(r.outputLoss),'백만원 ('+f1(r.outputShare)+'%)'),cardEl('부가가치 손실',fmt(r.vaLoss),'백만원 ('+f1(r.vaShare)+'%)'),cardEl('자체 산출 비중',f1(r.own/r.outputLoss*100),'% (손실 중)'));
  const mx=Math.max(...r.rows.map(x=>x.delta),1);
  $('ext-out').replaceChildren(d,text('h3','가장 크게 타격받는 산업 (자체 산출 감소분, 백만원)'),...r.rows.map(x=>barRow(x.code+' '+x.name+' (-'+f1(x.share)+'%)',x.delta,mx,'백만원')),text('p','나머지 산업의 감소분은 그 산업이 이 산업의 산출에 의존하고 이 산업이 그 산업의 산출을 사 주던 관계가 끊기는 효과를 합친 값입니다.'));
 }catch(e){$('ext-out').textContent='계산 실패: '+e.message;}
}

// ---------- 10. 산업 스코어카드 ----------
const scSec=mkSection('scorecard',`
<h2>10. 산업 스코어카드 — 한 장으로 보는 산업 프로필</h2>
<p>산업 하나(또는 두 산업의 비교)를 여섯 가지 지표의 <b>전체 산업 내 백분위</b>로 보여 줍니다. 100에 가까울수록 해당 지표가 모든 산업 중 상위입니다.</p>
<ul>
<li><b>후방 영향력</b>: 영향력계수 (이 산업의 수요가 다른 산업을 끌어올리는 힘)</li>
<li><b>전방 감응도</b>: 감응도계수 (다른 산업의 수요 증가를 받아내는 정도)</li>
<li><b>수입 의존도</b>: 투입 중 수입 중간재 비율</li>
<li><b>공공수요 비중</b>: 정부소비+정부투자가 유발한 생산의 비중</li>
<li><b>제거 시 손실</b>: 가상 제거법의 총산출 손실</li>
<li><b>물가 파급력</b>: 이 산업 가격이 오를 때 국내 최종수요 물가에 미치는 영향</li>
</ul>
<form id="sc-form">산업 A ${pickerHTML('sc-a')} 비교 산업 B ${pickerHTML('sc-b')} <button id="sc-run" type="submit" disabled>스코어카드 보기</button></form>
<p id="sc-status">분류를 불러오면 사용할 수 있습니다.</p>
<div id="sc-out"></div>`,extSec);
wirePicker('sc-a',null);wirePicker('sc-b',{value:'',label:'(비교 안 함)'});
function renderScore(){
 const a=$('sc-a-sector').value,b=$('sc-b-sector').value;
 if(!a||!linkRows.length){$('sc-status').textContent='연쇄효과 자료가 아직 준비되지 않았습니다.';return;}
 const pubMap=new Map(demAll.map(r=>[r.code,r.pub]));
 const AX=[['후방 영향력',r=>r.influence,f3],['전방 감응도',r=>r.sensitivity,f3],['수입 의존도',r=>r.importRatio,f3],['공공수요 비중',r=>pubMap.has(r.code)?pubMap.get(r.code):0,v=>f1(v)+'%'],['제거 시 손실',r=>r.extOutput===null||r.extOutput===undefined?0:r.extOutput,v=>f3(v)+'%'],['물가 파급력',r=>r.priceReach===null||r.priceReach===undefined?0:r.priceReach,f3]];
 const vals=AX.map(([n,f])=>linkRows.map(f));
 const prof=code=>{const r=linkRows.find(x=>x.code===code);return r?AX.map(([n,f],i)=>{const v=f(r);return {name:n,raw:v,pct:vals[i].filter(x=>x<=v).length/vals[i].length*100,fmt:AX[i][2]};}):null;};
 const A=prof(a),B=b?prof(b):null;if(!A)return;
 const W=460,H=420,cx=W/2,cy=H/2+4,R=140,N=AX.length;
 const pt=(i,p)=>{const ang=-Math.PI/2+i*2*Math.PI/N,rr=R*p/100;return [cx+rr*Math.cos(ang),cy+rr*Math.sin(ang)];};
 let s='<svg viewBox="0 0 '+W+' '+H+'" style="width:100%;max-width:520px;height:auto" role="img" aria-label="산업 스코어카드 레이더 차트">';
 for(const g of [20,40,60,80,100])s+='<polygon points="'+AX.map((_,i)=>pt(i,g).map(v=>v.toFixed(1)).join(',')).join(' ')+'" fill="none" stroke="currentColor" opacity="'+(g===100?.4:.18)+'"/>';
 AX.forEach((ax,i)=>{const [x,y]=pt(i,100),[lx,ly]=pt(i,118);s+='<line x1="'+cx+'" y1="'+cy+'" x2="'+x.toFixed(1)+'" y2="'+y.toFixed(1)+'" stroke="currentColor" opacity=".2"/><text x="'+lx.toFixed(1)+'" y="'+(ly+4).toFixed(1)+'" text-anchor="middle" font-size="12" fill="currentColor">'+esc(ax[0])+'</text>';});
 const poly=(P,c)=>'<polygon points="'+P.map((p,i)=>pt(i,p.pct).map(v=>v.toFixed(1)).join(',')).join(' ')+'" fill="'+c+'" fill-opacity=".22" stroke="'+c+'" stroke-width="2"/>'+P.map((p,i)=>{const [x,y]=pt(i,p.pct);return '<circle cx="'+x.toFixed(1)+'" cy="'+y.toFixed(1)+'" r="3.5" fill="'+c+'"><title>'+esc(p.name+' '+f1(p.pct)+'백분위')+'</title></circle>';}).join('');
 s+=poly(A,'#c0392b');if(B)s+=poly(B,'#2e86c1');
 s+='</svg>';
 const plot=document.createElement('div');plot.innerHTML=s;
 const legend=text('p','■ A: '+a+' '+secName(a)+(B?'   ■ B: '+b+' '+secName(b):''));
 const tbl=mkTable(['지표','A 값','A 백분위'].concat(B?['B 값','B 백분위']:[]),AX.map((ax,i)=>[ax[0],A[i].fmt(A[i].raw),f1(A[i].pct)].concat(B?[B[i].fmt(B[i].raw),f1(B[i].pct)]:[])));
 const topPct=p=>Math.max(1,Math.round(100-p+100/linkRows.length));
 const srt=[...A].sort((x,y)=>y.pct-x.pct);
 const note=text('p',secName(a)+'은(는) '+srt[0].name+'(상위 약 '+topPct(srt[0].pct)+'%)과 '+srt[1].name+'(상위 약 '+topPct(srt[1].pct)+'%)에서 두드러지고, '+srt[srt.length-1].name+'은(는) 상대적으로 낮습니다(상위 약 '+topPct(srt[srt.length-1].pct)+'%).');
 const tw=document.createElement('div');tw.style.cssText=WRAP;tw.append(tbl);
 $('sc-out').replaceChildren(plot,legend,note,tw);
 $('sc-status').textContent='백분위는 현재 불러온 분류의 '+linkRows.length+'개 산업 기준입니다.';
}
$('sc-form').onsubmit=e=>{e.preventDefault();renderScore();};

// ---------- 분류를 불러온 뒤 새 페이지 활성화 ----------
function initAnalysis(){
 pubBasket=[];pubRes=null;renderPubList();$('pub-out').replaceChildren();$('sup-out').replaceChildren();$('path-out').replaceChildren();
 for(const p of ['pub','sup','path','pr','sc-a','sc-b'])enablePicker(p);$('pr-out').replaceChildren();$('ext-out').replaceChildren();$('ext-title').textContent='';$('sc-out').replaceChildren();
 for(const id of ['pub-amount','pub-add','pub-run','sup-ratio','sup-run','path-amount','path-run','pr-mode','pr-pct','pr-run','sc-run'])$(id).disabled=false;$('pr-status').textContent='시나리오를 고르거나 직접 입력하세요.';$('sc-status').textContent='산업을 고르고 스코어카드를 보세요.';
 $('sup-status').textContent='산업과 대체 비율을 정하고 계산하세요.';$('path-status').textContent='산업과 수요 증가액을 정하고 계산하세요.';
}

// ===== 분류 선택 + 불러오기 =====
const loadBtn=$('load-basic');
const picker=document.createElement('select');
picker.id='dataset';
picker.setAttribute('aria-label','부문 분류 선택');
for(const [k,v] of Object.entries(DATASETS)){const o=document.createElement('option');o.value=k;o.textContent=v.label;picker.append(o);}
loadBtn.before(picker);
function idleLabel(){return DATASETS[picker.value].label+' 불러오기';}
loadBtn.textContent=idleLabel();
picker.onchange=()=>{loadBtn.disabled=false;loadBtn.textContent=idleLabel();};

async function loadWorkbook(key){
 currentKey=key;
 try{
  $('status').textContent=DATASETS[key].label+' Excel 내려받는 중…';
  const response=await fetch(encodeURI(DATASETS[key].file));
  if(!response.ok)throw Error(DATASETS[key].file+' 다운로드 실패 (HTTP '+response.status+')');
  await load(await response.arrayBuffer());
 }catch(error){
  $('status').textContent='준비 실패: '+error.message+' — 파일 경로와 연결 상태를 확인하세요.';
 }
}
loadBtn.onclick=async()=>{
 const key=picker.value;
 loadBtn.disabled=true;picker.disabled=true;
 loadBtn.textContent='불러오는 중…';
 try{
  await loadWorkbook(key);
  const ready=!$('sheet').disabled;
  loadBtn.textContent=ready?DATASETS[key].label+' 불러오기 완료':DATASETS[key].label+' 다시 불러오기';
  loadBtn.disabled=ready;
 }finally{
  picker.disabled=false;
  if($('sheet').disabled){loadBtn.disabled=false;loadBtn.textContent=DATASETS[key].label+' 다시 불러오기';}
 }
};
