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
 await explore();loadLinkage();$('status').textContent='준비 완료 · '+DATASETS[currentKey].label+' · '+sectors.length+'개 부문 · '+info.sheets.length+'개 시트 · 계수 검증 통과';
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
function renderResult(){$('results').hidden=false;$('cards').replaceChildren(...[['수요 증가액',result.demand],['생산유발액',result.production],['수입유발액',result.imports],['부가가치유발액',result.valueAdded]].map(([title,v])=>{const d=text('div',title);d.className='card';d.append(text('strong',fmt(v)),text('span','백만원'));return d;}));const max=Math.max(...result.rows.map(r=>r.production),1);$('bars').replaceChildren(...result.rows.slice(0,15).map(r=>{const d=text('div',r.name+' · '+fmt(r.production)+'백만원');d.className='bar';const t=document.createElement('div');t.className='track';const f=document.createElement('div');f.className='fill';f.style.width=Math.max(0,r.production/max*100)+'%';t.append(f);d.append(t);return d;}));$('result-rows').replaceChildren(...result.rows.map(r=>{const tr=document.createElement('tr');[r.code,r.name,fmt(r.production),fmt(r.imports),fmt(r.valueAdded)].forEach(v=>tr.append(text('td',v)));return tr;}));}
function csvCell(v){let s=String(v);if(/^[=+@-]/.test(s)&&typeof v!=='number')s="'"+s;return '"'+s.replaceAll('"','""')+'"';}
$('download').onclick=()=>{if(!result)return;const rows=[['산업코드','산업명','생산유발액_백만원','수입유발액_백만원','부가가치유발액_백만원'],...result.rows.map(r=>[r.code,r.name,r.production,r.imports,r.valueAdded])];const url=URL.createObjectURL(new Blob(['\ufeff'+rows.map(r=>r.map(csvCell).join(',')).join('\r\n')],{type:'text/csv;charset=utf-8'}));const a=document.createElement('a');a.href=url;a.download='io-2023-'+currentKey+'-results.csv';a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);};

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


// ===== 3. 산업 연쇄효과 분석 (영향력계수·감응도계수) =====
let linkRows=[];
const host=document.getElementById('simulation')||document.querySelector('main')||document.body;
const linkSec=document.createElement('section');
linkSec.id='linkage';
if(host.className)linkSec.className=host.className;
linkSec.innerHTML=`
<h2>3. 산업 연쇄효과 분석</h2>
<p>생산유발계수 역행렬의 열합·행합으로 산출합니다. 영향력계수(후방연쇄)와 감응도계수(전방연쇄)는 전 산업 평균을 1로 한 상대값이며 단위가 없습니다. 생산·부가가치·수입 유발계수 합은 최종수요 1단위(원/원)당 유발되는 크기입니다.</p>
<div>
 <input id="link-search" type="search" placeholder="산업 검색" disabled>
 <select id="link-sort" disabled>
  <option value="influence">영향력계수 높은 순</option>
  <option value="sensitivity">감응도계수 높은 순</option>
  <option value="production">생산유발계수 합 높은 순</option>
  <option value="valueAdded">부가가치유발계수 합 높은 순</option>
  <option value="imports">수입유발계수 합 높은 순</option>
 </select>
 <select id="link-type" disabled>
  <option value="">전체 유형</option>
  <option>중추 산업</option><option>최종수요 지향</option><option>중간수요 지향</option><option>독립 산업</option>
 </select>
</div>
<p id="link-status">분류를 불러오면 표시됩니다.</p>
<div style="overflow-x:auto;max-height:480px;overflow-y:auto">
<table><thead><tr><th>코드</th><th>산업</th><th>영향력계수</th><th>감응도계수</th><th>생산유발계수 합</th><th>부가가치유발계수 합</th><th>수입유발계수 합</th><th>유형</th></tr></thead>
<tbody id="link-rows"></tbody></table></div>`;
host.insertAdjacentElement('afterend',linkSec);
const f3=x=>Number(x).toLocaleString('ko-KR',{minimumFractionDigits:3,maximumFractionDigits:3});
function renderLinkage(){
 const q=$('link-search').value.trim(),t=$('link-type').value,k=$('link-sort').value;
 const list=linkRows.filter(r=>(r.code+' '+r.name).includes(q)&&(!t||r.type.startsWith(t))).sort((x,y)=>y[k]-x[k]);
 $('link-rows').replaceChildren(...list.map(r=>{const tr=document.createElement('tr');[r.code,r.name,f3(r.influence),f3(r.sensitivity),f3(r.production),f3(r.valueAdded),f3(r.imports),r.type].forEach(v=>tr.append(text('td',v)));return tr;}));
 $('link-status').textContent=list.length+'개 산업 표시 (전체 '+linkRows.length+'개). 영향력·감응도계수 1 이상이면 평균 이상입니다.';
}
async function loadLinkage(){
 for(const id of ['link-search','link-sort','link-type'])$(id).disabled=true;
 $('link-status').textContent='계산 중…';
 try{linkRows=(await call('linkage')).rows;renderLinkage();for(const id of ['link-search','link-sort','link-type'])$(id).disabled=false;}
 catch(e){$('link-status').textContent='연쇄효과 계산 실패: '+e.message;}
}
$('link-search').oninput=renderLinkage;$('link-sort').onchange=renderLinkage;$('link-type').onchange=renderLinkage;
