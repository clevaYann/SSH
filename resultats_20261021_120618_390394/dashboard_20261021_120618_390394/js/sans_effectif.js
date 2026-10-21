
(()=>{const p=window.DATA_PAYLOAD,$=s=>document.querySelector(s),esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
p.cats.forEach(c=>$('#c').add(new Option(c,c)));let k='DATE',d=1;
function render(){const q=$('#q').value.toLowerCase(),c=$('#c').value,v=p.rows.filter(r=>(!c||r.CATEGORIE===c)&&(!q||Object.values(r).join(' ').toLowerCase().includes(q))).sort((a,b)=>{const x=a[k],y=b[k];return(typeof x==='string'?x.localeCompare(y,'fr'):x-y)*d});
$('#t thead').innerHTML='<tr>'+p.cols.map(([f,l])=>'<th data-k="'+f+'">'+esc(l)+(f===k?(d>0?' ▲':' ▼'):' ↕')+'</th>').join('')+'</tr>';
$('#t tbody').innerHTML=v.map(r=>'<tr>'+p.cols.map(([f])=>'<td>'+esc(r[f])+'</td>').join('')+'</tr>').join('');$('#n').textContent=v.length+' séance(s) sur '+p.rows.length+'.';
document.querySelectorAll('#t th').forEach(th=>th.onclick=()=>{d=k===th.dataset.k?-d:1;k=th.dataset.k;render()})}
$('#q').addEventListener('input',render);$('#c').addEventListener('change',render);render();})();
