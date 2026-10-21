
(()=>{const p=window.DATA_PAYLOAD,rows=p.rows,sc=p.scenarios,$=s=>document.querySelector(s),fmt=(n,d=1)=>Number(n||0).toLocaleString('fr-FR',{minimumFractionDigits:d,maximumFractionDigits:d}),esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
let scope='all',key='r_'+sc[0],dir=-1;const chk=new Set();
const cols=[['name','Amphi','t'],['building','Bât.','t'],['city','Ville','t'],['site','Site','t'],['capacity','Places','n'],...sc.flatMap(n=>[['r_'+n,'Taux '+n+' (%)','n'],['h_'+n,'Heures '+n,'n']]),...sc.slice(1).map(n=>['d_'+n,'Écart '+n+' vs Réf. (pts)','n'])];
const days={AN:138,EX:10,HX:128},pname={AN:"l'année complète",EX:"les périodes d'examen (05/01/2026 – 17/01/2026 ; 18/05/2026 – 30/06/2026)",HX:"les jours hors périodes d'examen"};let per='AN';
function setPeriod(q){per=q;rows.forEach(r=>{Object.assign(r,r.P[q]);sc.slice(1).forEach(n=>r['d_'+n]=Math.round((r['r_'+n]-r['r_'+sc[0]])*100)/100)});$('#period-note').textContent='Période affichée : '+pname[q]+' — '+days[q]+' jour(s) pédagogique(s).'+(days[q]===0?' Aucune donnée sur cette période dans l’export.':'');document.querySelectorAll('.pseg button').forEach(b=>b.setAttribute('aria-pressed',String(b.dataset.period===q)))}
setPeriod('AN');
const sel={city:new Set(),site:new Set(),building:new Set()},labels={city:['Toutes les villes','ville(s)'],site:['Tous les sites','site(s)'],building:['Tous les bâtiments','bâtiment(s)']},ids={city:'#m-city',site:'#m-site',building:'#m-bld'};
function base(){return rows.filter(r=>scope==='all'||(scope==='let'&&r.removed)||(scope==='deg'&&r.host))}
function pool(field){let p=base();if(field!=='city'&&sel.city.size)p=p.filter(r=>sel.city.has(r.city));if(field==='building'&&sel.site.size)p=p.filter(r=>sel.site.has(r.site));return p}
function msel(field){const root=$(ids[field]),open=root.dataset.open==='1',counts={};pool(field).forEach(r=>{if(r[field])counts[r[field]]=(counts[r[field]]||0)+1});[...sel[field]].forEach(v=>{if(!(v in counts))sel[field].delete(v)});
const n=sel[field].size,label=n?n+' '+labels[field][1]+' sélectionné(s)':labels[field][0];
root.innerHTML='<button type="button" class="msel-btn" aria-expanded="'+open+'">'+esc(label)+' ▾</button><div class="msel-panel"'+(open?'':' hidden')+'><div class="msel-act"><a href="#" data-a="all">Tout cocher</a> · <a href="#" data-a="none">Tout décocher</a></div>'+Object.keys(counts).sort().map(v=>'<label><input type="checkbox" value="'+esc(v)+'"'+(sel[field].has(v)?' checked':'')+'> '+esc(v)+' <small>('+counts[v]+')</small></label>').join('')+'</div>';
root.querySelector('.msel-btn').onclick=e=>{e.stopPropagation();const o=root.dataset.open==='1';document.querySelectorAll('.msel').forEach(m=>{m.dataset.open='0';m.querySelector('.msel-panel').hidden=true});root.dataset.open=o?'0':'1';root.querySelector('.msel-panel').hidden=o};
root.querySelector('.msel-panel').onclick=e=>e.stopPropagation();
root.querySelectorAll('input').forEach(c=>c.onchange=()=>{c.checked?sel[field].add(c.value):sel[field].delete(c.value);render()});
root.querySelectorAll('[data-a]').forEach(a=>a.onclick=e=>{e.preventDefault();sel[field].clear();if(a.dataset.a==='all')Object.keys(counts).forEach(v=>sel[field].add(v));render()});}
function visible(){const q=$('#f-q').value.trim().toLowerCase();return base().filter(r=>(!sel.city.size||sel.city.has(r.city))&&(!sel.site.size||sel.site.has(r.site))&&(!sel.building.size||sel.building.has(r.building))&&(!q||(r.name+' '+r.code+' '+r.building).toLowerCase().includes(q)))}
function render(){['city','site','building'].forEach(msel);
const v=visible().sort((a,b)=>{const x=a[key],y=b[key];return(typeof x==='string'?x.localeCompare(y,'fr'):(x-y))*dir||a.name.localeCompare(b.name,'fr')});
$('#t thead').innerHTML='<tr><th><input type="checkbox" id="chk-all" title="Cocher / décocher les amphis affichés" '+(v.length&&v.every(r=>chk.has(r.code))?'checked':'')+'></th>'+cols.map(([k,l,t])=>'<th data-k="'+k+'" class="'+(t==='n'?'num':'')+'">'+esc(l)+(k===key?(dir>0?' ▲':' ▼'):' ↕')+'</th>').join('')+'</tr>';
$('#t tbody').innerHTML=v.map(r=>'<tr'+(r.removed?' style="background:#e6f3ef"':r.host?' style="background:#fff5e0"':'')+'><td><input type="checkbox" class="rc" value="'+esc(r.code)+'"'+(chk.has(r.code)?' checked':'')+'></td>'+cols.map(([k,l,t])=>t==='n'?'<td class="num">'+(k==='capacity'?r[k]:fmt(r[k],k[0]==='h'?1:2))+'</td>':'<td>'+esc(r[k])+'</td>').join('')+'</tr>').join('');
$('#count').textContent=v.length+' amphi(s) affiché(s) sur '+rows.length+'.';
document.querySelectorAll('#t th[data-k]').forEach(th=>th.onclick=()=>{dir=key===th.dataset.k?-dir:(th.classList.contains('num')?-1:1);key=th.dataset.k;render()});
document.querySelectorAll('#t .rc').forEach(c=>c.onchange=()=>{c.checked?chk.add(c.value):chk.delete(c.value);render()});
const all=$('#chk-all');if(all)all.onchange=()=>{v.forEach(r=>all.checked?chk.add(r.code):chk.delete(r.code));render()};
average(v);}
function average(v){const picked=rows.filter(r=>chk.has(r.code)),set=picked.length?picked:v,n=set.length,own=picked.length>0;
const mean=f=>n?set.reduce((a,r)=>a+r[f],0)/n:0,sum=f=>set.reduce((a,r)=>a+r[f],0),ref=mean('r_'+sc[0]);
$('#avg').innerHTML='<h2>Moyenne générale des amphis '+(own?'sélectionnés ('+n+' coché'+(n>1?'s':'')+')':'affichés ('+n+') — cochez des amphis pour choisir')+'</h2>'+(n?'<div class="tablewrap"><table><thead><tr><th>Cas</th><th class="num">Taux moyen (%)</th><th class="num">Écart vs Référence (pts)</th><th class="num">Heures totales</th><th class="num">Heures moyennes / amphi</th></tr></thead><tbody>'+sc.map(s=>{const m=mean('r_'+s),d=m-ref;return '<tr><td>'+esc(s)+'</td><td class="num"><b>'+fmt(m,2)+'</b></td><td class="num">'+(s===sc[0]?'—':(d>0?'+':'')+fmt(d,2))+'</td><td class="num">'+fmt(sum('h_'+s),1)+'</td><td class="num">'+fmt(sum('h_'+s)/n,1)+'</td></tr>'}).join('')+'</tbody></table></div><p class="muted">Moyenne arithmétique des taux des amphis retenus (tous ont le même nombre de jours théoriques, donc identique à la moyenne pondérée par les heures). '+(own?'<a href="#" id="clr">Décocher tous</a>':'')+'</p>':'<p class="muted">Aucun amphi affiché.</p>');
const c=$('#clr');if(c)c.onclick=e=>{e.preventDefault();chk.clear();render()};}
document.addEventListener('click',()=>document.querySelectorAll('.msel').forEach(m=>{if(m.dataset.open==='1'){m.dataset.open='0';m.querySelector('.msel-panel').hidden=true}}));
$('#f-q').addEventListener('input',render);
$('#f-reset').addEventListener('click',()=>{Object.values(sel).forEach(x=>x.clear());chk.clear();$('#f-q').value='';scope='all';setPeriod('AN');document.querySelectorAll('.seg button').forEach(x=>x.setAttribute('aria-pressed',String(x.dataset.scope==='all')));render()});
document.querySelectorAll('.pseg button').forEach(b=>b.addEventListener('click',()=>{setPeriod(b.dataset.period);render()}));
document.querySelectorAll('.seg button').forEach(b=>b.addEventListener('click',()=>{scope=b.dataset.scope;document.querySelectorAll('.seg button').forEach(x=>x.setAttribute('aria-pressed',String(x===b)));render()}));
render();})();
