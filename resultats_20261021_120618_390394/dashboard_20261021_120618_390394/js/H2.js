
(async function(){
  
  
  
  
  const rawConflicts = []
;
  
  
  
  
  
  const EXAM_WEEKS = new Set(["2026-S02", "2026-S03"]);
  const EXAM_DATES = new Set(["05/01/2026", "06/01/2026", "07/01/2026", "08/01/2026", "09/01/2026", "12/01/2026", "13/01/2026", "14/01/2026", "15/01/2026", "16/01/2026"]);
  const EXAM_LABEL = "les périodes d'examen (05/01/2026 – 17/01/2026 ; 18/05/2026 – 30/06/2026)";



  // Semaine -> jours ouvrés disponibles / date du lundi, dérivés du bucket 'TOUS' (non filtré)
  const nbJoursSemaineLookup = {};
  const lundiLookup = {};
  (weeklyData['TOUS']||[]).forEach(w => { nbJoursSemaineLookup[w.label] = w.nbJoursOuvresSemaine; lundiLookup[w.label] = w.lundi; });

  const conflicts = rawConflicts.map(c => {
    const [dd,mm,yyyy] = c.date.split('/');
    return {
      ...c,
      dateSort: `${yyyy}${mm}${dd}`,
      annee: parseInt(yyyy,10),
      sameMatiere: c.matiere1 === c.matiere2 && c.hdebut1 === c.hdebut2 && c.hfin1 === c.hfin2
    };
  });

  const data = rawData.map(r => ({
    salle: r.salle || '(sans nom)', batiment: r.batiment || '', ville: r.ville || '', site: r.site || '', etage: r.etage || '',
    capa: r.capa, surface: r.surface, ratio: r.ratio, heures: r.heures ?? 0,
    nbCours: r.nbCours ?? 0, taux: r.taux,
    heuresS1: r.heuresS1 ?? 0, tauxS1: r.tauxS1, heuresS2: r.heuresS2 ?? 0, tauxS2: r.tauxS2,
    heuresEX: r.heuresEX ?? 0, tauxEX: r.tauxEX, heuresHX: r.heuresHX ?? 0, tauxHX: r.tauxHX
  }));
  const fmt = (n, d=1) => (n===null||n===undefined) ? '—' : n.toLocaleString('fr-FR', {minimumFractionDigits:d, maximumFractionDigits:d});

  // Ville/site par bâtiment, pour appliquer les filtres globaux aux chevauchements (qui n'ont pas ville/site en propre)
  const batInfo = {};
  data.forEach(d => { if (d.batiment && !batInfo[d.batiment]) batInfo[d.batiment] = { ville: d.ville, site: d.site }; });

  // ================= Semestre =================
  let semestre = '';  // '' = année complète, 'S1', 'S2'
  function eTaux(d){ return semestre==='S1' ? d.tauxS1 : semestre==='S2' ? d.tauxS2 : semestre==='EX' ? d.tauxEX : semestre==='HX' ? d.tauxHX : d.taux; }
  function eHeures(d){ return semestre==='S1' ? (d.heuresS1||0) : semestre==='S2' ? (d.heuresS2||0) : semestre==='EX' ? (d.heuresEX||0) : semestre==='HX' ? (d.heuresHX||0) : (d.heures||0); }
  function semLabel(){ return semestre==='EX' ? EXAM_LABEL : semestre==='HX' ? 'les jours hors périodes d\'examen' : semestre==='S1' ? '1er semestre (S37–S53)' : semestre==='S2' ? '2nd semestre (S1–S19)' : "l'année complète"; }
  document.querySelectorAll('.seg button[data-sem]').forEach(btn => {
    btn.addEventListener('click', () => {
      semestre = btn.dataset.sem;
      document.querySelectorAll('.seg button[data-sem]').forEach(b => b.setAttribute('aria-pressed', b===btn ? 'true':'false'));
      page=1; gridPage=1; renderAll();
    });
  });

  // ================= Filtres globaux =================
  const gSearch = document.getElementById('g-search');
  const gReset = document.getElementById('g-reset');
  const gNote = document.getElementById('g-note');

  function titleCase(s){ return s.charAt(0)+s.slice(1).toLowerCase(); }

  // ---- Fabrique générique de menu à cases à cocher ----
  function makeMsel(name, { allLabel, formatLabel, onChange }) {
    const root = document.getElementById(`g-${name}-msel`);
    const btn = document.getElementById(`g-${name}-btn`);
    const panel = document.getElementById(`g-${name}-panel`);
    const list = document.getElementById(`g-${name}-list`);
    const selected = new Set();  // vide = tous

    function updateLabel(){
      if (selected.size === 0) btn.textContent = allLabel;
      else if (selected.size === 1) btn.textContent = formatLabel([...selected][0]);
      else btn.textContent = `${selected.size} sélectionnés`;
    }
    function setOptions(values, counts){
      [...selected].forEach(v => { if (!values.includes(v)) selected.delete(v); });
      list.innerHTML = values.map(v => `
        <label><input type="checkbox" value="${v}" ${selected.has(v)?'checked':''}> ${formatLabel(v)}${counts ? ` <span class="msel-count">${counts[v] ?? 0}</span>` : ''}</label>
      `).join('');
      updateLabel();
    }
    btn.addEventListener('click', () => {
      const open = panel.hidden;
      document.querySelectorAll('.msel-panel').forEach(p => { p.hidden = true; });
      document.querySelectorAll('.msel-btn').forEach(b => b.setAttribute('aria-expanded','false'));
      panel.hidden = !open;
      btn.setAttribute('aria-expanded', open ? 'true' : 'false');
    });
    document.addEventListener('click', (e) => {
      if (!root.contains(e.target)) { panel.hidden = true; btn.setAttribute('aria-expanded', 'false'); }
    });
    list.addEventListener('change', (e) => {
      if (e.target.type !== 'checkbox') return;
      if (e.target.checked) selected.add(e.target.value); else selected.delete(e.target.value);
      updateLabel();
      onChange();
    });
    const allBtn = root.querySelector('[data-action="all"]');
    const noneBtn = root.querySelector('[data-action="none"]');
    if (allBtn) allBtn.addEventListener('click', () => {
      list.querySelectorAll('input[type="checkbox"]:not(.msel-group-cb)').forEach(cb => { cb.checked = true; selected.add(cb.value); });
      updateLabel(); onChange();
    });
    if (noneBtn) noneBtn.addEventListener('click', () => {
      list.querySelectorAll('input[type="checkbox"]:not(.msel-group-cb)').forEach(cb => { cb.checked = false; });
      selected.clear();
      updateLabel(); onChange();
    });

    return { selected, setOptions, updateLabel };
  }

  const onFilterChange = () => { page=1; gridPage=1; renderAll(); };

  const villeMsel = makeMsel('ville', {
    allLabel: 'Toutes les villes', formatLabel: titleCase,
    onChange: () => { refreshSiteOptions(); refreshBatimentOptions(); onFilterChange(); }
  });
  const siteMsel = makeMsel('site', {
    allLabel: 'Tous les sites', formatLabel: titleCase,
    onChange: () => { refreshBatimentOptions(); onFilterChange(); }
  });
  const batMsel = makeMsel('batiment', {
    allLabel: 'Tous les bâtiments', formatLabel: (b) => b,
    onChange: () => { refreshBatimentOptions(); onFilterChange(); }
  });
  const tauxLabels = { low:'Sous-utilisée (< 10%)', mid:'Modérée (10–30%)', high:'Forte (≥ 30%)' };
  const tauxMsel = makeMsel('taux', {
    allLabel: 'Tous les taux', formatLabel: (t) => tauxLabels[t] || t,
    onChange: onFilterChange
  });

  villeMsel.setOptions([...new Set(data.map(d=>d.ville).filter(Boolean))].sort());

  function refreshSiteOptions(){
    const pool = villeMsel.selected.size ? data.filter(d => villeMsel.selected.has(d.ville)) : data;
    const sites = [...new Set(pool.map(d=>d.site).filter(Boolean))].sort();
    const counts = {}; sites.forEach(s => { counts[s] = pool.filter(d=>d.site===s).length; });
    siteMsel.setOptions(sites, counts);
  }

  // ================= Groupement des bâtiments (STEE / SSH / EEI / Autres) =================
  const buildingGroups = {
    STEE: ['BA1','BA2','ISA','IUT','CNC','CAN','DUB','GTE','IBE','IPR1','IPR2','PAL','SCI_R','STI','DAL_R','FER','LAV','IPM','IPM2'],
    SSH:  ['DEG','ICL','LET'],
    EEI:  ['CNA','CNB','B3','B1','B2','EC9','EC3','EC6']
  };
  const groupNames = ['STEE','SSH','EEI','Autres'];
  function groupOf(bat){
    for (const g of ['STEE','SSH','EEI']) { if (buildingGroups[g].includes(bat)) return g; }
    return 'Autres';
  }

  function refreshBatimentOptions(){
    let pool = data;
    if (villeMsel.selected.size) pool = pool.filter(d => villeMsel.selected.has(d.ville));
    if (siteMsel.selected.size) pool = pool.filter(d => siteMsel.selected.has(d.site));
    const bats = [...new Set(pool.map(d=>d.batiment).filter(Boolean))].sort();
    const counts = {}; bats.forEach(b => { counts[b] = pool.filter(d=>d.batiment===b).length; });

    [...batMsel.selected].forEach(v => { if (!bats.includes(v)) batMsel.selected.delete(v); });

    const list = document.getElementById('g-batiment-list');
    list.innerHTML = groupNames.map(g => {
      const codes = bats.filter(b => groupOf(b) === g);
      if (!codes.length) return '';
      const allChecked = codes.every(b => batMsel.selected.has(b));
      const groupTotal = codes.reduce((s,b)=>s+counts[b],0);
      return `
        <div class="msel-group">
          <label class="msel-group-label"><input type="checkbox" class="msel-group-cb" data-group="${g}" ${allChecked?'checked':''}> ${g} <span class="msel-count">${groupTotal}</span></label>
          <div class="msel-group-children">
            ${codes.map(b => `<label><input type="checkbox" value="${b}" ${batMsel.selected.has(b)?'checked':''}> ${b} <span class="msel-count">${counts[b]}</span></label>`).join('')}
          </div>
        </div>`;
    }).join('');

    list.querySelectorAll('.msel-group-cb').forEach(cb => {
      const codes = bats.filter(b => groupOf(b) === cb.dataset.group);
      const checkedCount = codes.filter(b => batMsel.selected.has(b)).length;
      cb.indeterminate = checkedCount > 0 && checkedCount < codes.length;
    });

    batMsel.updateLabel();
  }
  // Case à cocher de groupe : sélectionne/désélectionne tous les bâtiments du groupe.
  // Capture phase pour intercepter l'événement avant l'écouteur générique de makeMsel
  // (qui traiterait sinon "STEE" comme un code bâtiment invalide).
  document.getElementById('g-batiment-list').addEventListener('change', (e) => {
    if (!e.target.classList.contains('msel-group-cb')) return;
    e.stopPropagation();
    const codes = buildingGroups[e.target.dataset.group] || [];
    let pool = data;
    if (villeMsel.selected.size) pool = pool.filter(d => villeMsel.selected.has(d.ville));
    if (siteMsel.selected.size) pool = pool.filter(d => siteMsel.selected.has(d.site));
    const available = new Set(pool.map(d=>d.batiment));
    const codesPresent = e.target.dataset.group === 'Autres'
      ? [...available].filter(b => groupOf(b) === 'Autres')
      : codes.filter(b => available.has(b));
    if (e.target.checked) codesPresent.forEach(c => batMsel.selected.add(c));
    else codesPresent.forEach(c => batMsel.selected.delete(c));
    refreshBatimentOptions();
    onFilterChange();
  }, true);

  refreshSiteOptions();
  refreshBatimentOptions();

  function taxTier(t){ if (t===null||t===undefined) return ''; if (t<10) return 'low'; if (t<30) return 'mid'; return 'high'; }
  function matchesGlobal(d){
    const q = gSearch.value.trim().toLowerCase();
    if (q && !(d.salle.toLowerCase().includes(q) || d.batiment.toLowerCase().includes(q))) return false;
    if (villeMsel.selected.size && !villeMsel.selected.has(d.ville)) return false;
    if (siteMsel.selected.size && !siteMsel.selected.has(d.site)) return false;
    if (batMsel.selected.size && !batMsel.selected.has(d.batiment)) return false;
    if (tauxMsel.selected.size && !tauxMsel.selected.has(taxTier(eTaux(d)))) return false;
    return true;
  }
  function updateNote(){
    const n = data.filter(matchesGlobal).length;
    const parts = [];
    if (villeMsel.selected.size) parts.push([...villeMsel.selected].map(titleCase).join('+'));
    if (siteMsel.selected.size) parts.push([...siteMsel.selected].map(titleCase).join('+'));
    if (batMsel.selected.size) parts.push([...batMsel.selected].join('+'));
    if (gSearch.value.trim()) parts.push(`"${gSearch.value.trim()}"`);
    if (tauxMsel.selected.size) parts.push([...tauxMsel.selected].map(t=>({low:'sous-utilisée',mid:'modérée',high:'forte'}[t])).join('+'));
    gNote.textContent = parts.length ? `${n} salle${n>1?'s':''} — ${parts.join(' · ')}` : '';
  }

  // ================= Tabs =================
  const tabs = document.querySelectorAll('.tab');
  const panels = { overview: document.getElementById('tab-overview'), grid: document.getElementById('tab-grid'), weekly: document.getElementById('tab-weekly'), exec: document.getElementById('tab-exec'), conflits: document.getElementById('tab-conflits') };

  tabs.forEach(t => t.addEventListener('click', () => {
    tabs.forEach(x => x.setAttribute('aria-selected','false'));
    t.setAttribute('aria-selected','true');
    Object.values(panels).forEach(p => p.classList.remove('active'));
    panels[t.dataset.tab].classList.add('active');
    renderAll();
  }));

  // ================= Bar list helper =================
  function renderBarlist(el, items, opts={}) {
    if (!items.length){ el.innerHTML = '<p class="opp-note">Aucune donnée pour ce filtre.</p>'; return; }
    const max = Math.max(...items.map(i=>i.value), 1);
    el.innerHTML = items.map(it => `
      <div class="barrow" tabindex="0" aria-label="${it.label}, ${fmt(it.value, opts.decimals ?? 2)}${opts.suffix||''}">
        <span class="rlabel" title="${it.fullLabel || it.label}">${it.label}</span>
        <span class="track"><span class="fill ${opts.cls||''}" data-w="${(it.value/max*100).toFixed(1)}"></span></span>
        <span class="rval">${fmt(it.value, opts.decimals ?? 2)}${opts.suffix||''}</span>
      </div>`).join('');
    requestAnimationFrame(() => el.querySelectorAll('.fill').forEach(f => { f.style.width = f.dataset.w + '%'; }));
  }

  function heatClass(pct){
    if (pct > 100) return 'heat-x';
    if (pct <= 0) return 'heat-0';
    if (pct < 10) return 'heat-1';
    if (pct < 25) return 'heat-2';
    if (pct < 40) return 'heat-3';
    if (pct < 60) return 'heat-4';
    return 'heat-5';
  }

  // ================= TAB 1 : Vue d'ensemble =================
  let sortKey = 'taux', sortDir = 'desc', page = 1;
  function renderOverview(){
    const scopedAll = data.filter(matchesGlobal);
    const filtered = scopedAll.filter(d => eTaux(d) !== null && eTaux(d) !== undefined);

    const tauxMoyen = filtered.length ? filtered.reduce((s,d)=>s+eTaux(d),0) / filtered.length : 0;
    const sousUtil = filtered.filter(d => eTaux(d) < 10).length;
    const surUtil = filtered.filter(d => eTaux(d) >= 30).length;
    const totalHeures = scopedAll.reduce((s,d)=>s+eHeures(d),0);

    document.getElementById('kpi-total').textContent = scopedAll.length;
    document.getElementById('kpi-taux').innerHTML = fmt(tauxMoyen,2) + '<small>%</small>';
    document.getElementById('kpi-sous').textContent = sousUtil;
    document.getElementById('kpi-sur').textContent = surUtil;
    document.getElementById('kpi-heures').innerHTML = fmt(totalHeures,0) + '<small> h</small>';
    document.querySelectorAll('.kpi .label')[1].textContent = `Taux d'occupation moyen (${semLabel()})`;

    const byBat = {};
    filtered.forEach(d => { if (!d.batiment) return; (byBat[d.batiment] ??= []).push(eTaux(d)); });
    const batRanking = Object.entries(byBat).map(([b, arr]) => ({ label:b, value: arr.reduce((s,v)=>s+v,0)/arr.length })).sort((a,b) => b.value - a.value);
    document.getElementById('batiment-count-sub').textContent = `${batRanking.length} bâtiments, trié décroissant — ${semLabel()}`;
    renderBarlist(document.getElementById('chart-batiment'), batRanking, { suffix:'%', decimals:2 });

    const buckets = [{label:'0–5',min:0,max:5},{label:'5–10',min:5,max:10},{label:'10–15',min:10,max:15},{label:'15–20',min:15,max:20},{label:'20–25',min:20,max:25},{label:'25–30',min:25,max:30},{label:'30–40',min:30,max:40},{label:'40+',min:40,max:Infinity}];
    const counts = buckets.map(b => filtered.filter(d => eTaux(d) >= b.min && eTaux(d) < b.max).length);
    const maxCount = Math.max(...counts, 1);
    const totalCompte = filtered.length || 1;
    const histoEl = document.getElementById('chart-distrib');
    histoEl.innerHTML = buckets.map((b,i) => {
      const pct = counts[i] / totalCompte * 100;
      return `<div class="col"><span class="count">${counts[i]}<span class="pct">${fmt(pct,0)}%</span></span><span class="bar" data-h="${(counts[i]/maxCount*100).toFixed(1)}" style="height:2px" title="${b.label}% d'occupation : ${counts[i]} salles (${fmt(pct,1)}% du total)"></span><span class="tick">${b.label}%</span></div>`;
    }).join('');
    requestAnimationFrame(() => histoEl.querySelectorAll('.bar').forEach(b => { b.style.height = b.dataset.h + '%'; }));

    const opp = scopedAll.filter(d => (d.capa||0) >= 40 && eTaux(d) !== null && eTaux(d) !== undefined && eTaux(d) < 10)
      .sort((a,b) => (b.capa||0) - (a.capa||0)).slice(0, 10)
      .map(d => ({ label:`${d.salle} (${d.batiment})`, value:d.capa, fullLabel:`${d.salle} — ${d.batiment} — ${fmt(eTaux(d),2)}% d'usage` }));
    renderBarlist(document.getElementById('chart-opp'), opp, { suffix:' places', decimals:0, cls:'crit' });

    // table
    let rows = scopedAll.slice();
    const keyFn = (d) => sortKey==='taux' ? eTaux(d) : sortKey==='heures' ? eHeures(d) : d[sortKey];
    rows.sort((a,b) => {
      let av=keyFn(a), bv=keyFn(b);
      if (av===null||av===undefined) av = (sortKey==='salle'||sortKey==='batiment'||sortKey==='etage') ? '' : -Infinity;
      if (bv===null||bv===undefined) bv = (sortKey==='salle'||sortKey==='batiment'||sortKey==='etage') ? '' : -Infinity;
      if (typeof av === 'string'){ av=av.toLowerCase(); bv=bv.toLowerCase(); }
      const cmp = av<bv?-1:av>bv?1:0; return sortDir==='asc'?cmp:-cmp;
    });
    const size = parseInt(document.getElementById('pg-size').value,10);
    const totalPages = Math.max(1, Math.ceil(rows.length/size));
    page = Math.min(page, totalPages);
    const pageRows = rows.slice((page-1)*size, (page-1)*size+size);
    document.getElementById('pg-info').textContent = `Page ${page} / ${totalPages}`;
    document.getElementById('pg-prev').disabled = page<=1;
    document.getElementById('pg-next').disabled = page>=totalPages;
    document.getElementById('table-body').innerHTML = pageRows.map(d => {
      const t = eTaux(d);
      const tier = taxTier(t); const pillCls = tier==='low'?'crit':tier==='high'?'ok':'warn';
      return `<tr><td class="salle">${d.salle}</td><td>${d.batiment||'—'}</td><td>${d.etage||'—'}</td>
        <td class="num">${fmt(d.capa,0)}</td><td class="num">${fmt(d.surface,1)}</td><td class="num">${fmt(d.ratio,2)}</td>
        <td class="num">${fmt(eHeures(d),1)}</td><td class="num">${fmt(d.nbCours,0)}</td>
        <td class="num">${t!==null&&t!==undefined?`<span class="pill ${pillCls}">${fmt(t,2)}%</span>`:'—'}</td></tr>`;
    }).join('');
  }
  document.querySelectorAll('#tab-overview thead th[data-key]').forEach(th => {
    th.addEventListener('click', () => {
      const key = th.dataset.key;
      if (sortKey===key) sortDir = sortDir==='asc'?'desc':'asc'; else { sortKey=key; sortDir = (key==='salle'||key==='batiment'||key==='etage')?'asc':'desc'; }
      document.querySelectorAll('#tab-overview thead th').forEach(h => { delete h.dataset.active; h.querySelector('.arrow').textContent='↕'; });
      th.dataset.active=''; th.querySelector('.arrow').textContent = sortDir==='asc'?'↑':'↓';
      page = 1; renderOverview();
    });
  });
  document.getElementById('pg-size').addEventListener('change', () => { page=1; renderOverview(); });
  document.getElementById('pg-prev').addEventListener('click', () => { page--; renderOverview(); });
  document.getElementById('pg-next').addEventListener('click', () => { page++; renderOverview(); });

  // ================= TAB 2 : Grille horaire =================
  let gridSortKey = 'taux', gridSortDir = 'desc', gridPage = 1;
  const HOURS = [8,9,10,11,12,13,14,15,16,17];
  const gridHead = document.getElementById('grid-head');
  HOURS.forEach(h => { const th = document.createElement('th'); th.className='num'; th.textContent = `${h}h-${h+1}h`; gridHead.appendChild(th); });

  function renderGrid(){
    let rows = data.filter(matchesGlobal).slice();
    const keyFn = (d) => gridSortKey==='taux' ? eTaux(d) : d[gridSortKey];
    rows.sort((a,b) => {
      let av=keyFn(a), bv=keyFn(b);
      if (av===null||av===undefined) av = (gridSortKey==='salle'||gridSortKey==='batiment'||gridSortKey==='etage') ? '' : -Infinity;
      if (bv===null||bv===undefined) bv = (gridSortKey==='salle'||gridSortKey==='batiment'||gridSortKey==='etage') ? '' : -Infinity;
      if (typeof av === 'string'){ av=av.toLowerCase(); bv=bv.toLowerCase(); }
      const cmp = av<bv?-1:av>bv?1:0; return gridSortDir==='asc'?cmp:-cmp;
    });
    const size = 30;
    const totalPages = Math.max(1, Math.ceil(rows.length/size));
    gridPage = Math.min(gridPage, totalPages);
    const pageRows = rows.slice((gridPage-1)*size, (gridPage-1)*size+size);
    document.getElementById('grid-info').textContent = `Page ${gridPage} / ${totalPages}`;
    document.getElementById('grid-prev').disabled = gridPage<=1;
    document.getElementById('grid-next').disabled = gridPage>=totalPages;
    document.getElementById('grid-count').textContent = `${rows.length} salles — taux sur ${semLabel()}${semestre ? ' · grille horaire calculée sur l\'année complète' : ''}`;

    document.getElementById('grid-body').innerHTML = pageRows.map(d => {
      const key = `${d.salle}|${d.batiment}`;
      const hm = heatmapSalle[key];
      let hmCells = '';
      if (hm) {
        hmCells = HOURS.map(h => {
          let sum=0, n=0;
          ['lundi','mardi','mercredi','jeudi','vendredi'].forEach(j => {
            if (hm[j] && hm[j][`${h}h-${h+1}h`] !== undefined) { sum += hm[j][`${h}h-${h+1}h`]; n++; }
          });
          const avg = n ? sum/n : 0;
          const title = avg > 100 ? `Conflit potentiel : réservations qui se chevauchent certaines semaines (${avg.toFixed(0)}%)` : `${avg.toFixed(0)}%`;
          return `<td class="heatcell ${heatClass(avg)}" title="${title}">${avg.toFixed(0)}</td>`;
        }).join('');
      } else {
        hmCells = HOURS.map(()=>`<td class="heatcell heat-0">—</td>`).join('');
      }
      const t = eTaux(d);
      return `<tr><td class="salle">${d.salle}</td><td>${d.batiment||'—'}</td><td>${d.etage||'—'}</td>
        <td class="num">${fmt(d.surface,0)}</td><td class="num">${fmt(d.capa,0)}</td><td class="num">${fmt(d.ratio,2)}</td>
        <td class="num">${t!==null&&t!==undefined?fmt(t,1)+'%':'—'}</td>${hmCells}</tr>`;
    }).join('');
    document.querySelectorAll('.heatcell').forEach(c => {
      c.style.background = `var(--${c.className.split(' ')[1]})`;
    });
  }
  document.querySelectorAll('#grid-head th[data-key]').forEach(th => {
    th.addEventListener('click', () => {
      const key = th.dataset.key;
      if (gridSortKey===key) gridSortDir = gridSortDir==='asc'?'desc':'asc'; else { gridSortKey=key; gridSortDir = (key==='salle'||key==='batiment'||key==='etage')?'asc':'desc'; }
      document.querySelectorAll('#grid-head th').forEach(h => { delete h.dataset.active; if(h.querySelector('.arrow')) h.querySelector('.arrow').textContent='↕'; });
      th.dataset.active=''; if(th.querySelector('.arrow')) th.querySelector('.arrow').textContent = gridSortDir==='asc'?'↑':'↓';
      gridPage = 1; renderGrid();
    });
  });
  document.getElementById('grid-prev').addEventListener('click', () => { gridPage--; renderGrid(); });
  document.getElementById('grid-next').addEventListener('click', () => { gridPage++; renderGrid(); });

  // ================= TAB 3 : Analyse hebdomadaire =================
  function renderWeekly(){
    // Tout est agrégé dynamiquement sur l'ensemble des salles qui passent matchesGlobal
    // (ville/site/bâtiment/recherche/taux), comme les autres onglets — plus de restriction
    // à "un seul bâtiment à la fois".
    const scopedAll = data.filter(matchesGlobal);
    const nbSallesFiltre = scopedAll.length;

    const parts = [];
    if (villeMsel.selected.size) parts.push([...villeMsel.selected].map(titleCase).join('+'));
    if (siteMsel.selected.size) parts.push([...siteMsel.selected].map(titleCase).join('+'));
    if (batMsel.selected.size) parts.push([...batMsel.selected].join('+'));
    if (gSearch.value.trim()) parts.push(`"${gSearch.value.trim()}"`);
    if (tauxMsel.selected.size) parts.push([...tauxMsel.selected].map(t=>({low:'sous-utilisée',mid:'modérée',high:'forte'}[t])).join('+'));
    const scopeNote = `${nbSallesFiltre} salle${nbSallesFiltre>1?'s':''}` + (parts.length ? ` — ${parts.join(' · ')}` : ' — tous bâtiments');
    document.getElementById('weekly-scope').textContent = scopeNote + (semestre ? ` — ${semLabel()}` : '');
    document.getElementById('heatmap-scope').textContent = scopeNote + ` — ${semLabel()}`;
    document.getElementById('peakday-scope').textContent = scopeNote + ` — ${semLabel()}`;

    // Agrège les heures brutes (weekly_par_salle.json) sur les salles filtrées, semaine par semaine
    const parSemaine = {};
    scopedAll.forEach(d => {
      const raw = weeklySalleData[`${d.salle}|${d.batiment}`];
      if (!raw) return;
      // Filet de sécurité : ConvertTo-Json sérialise un objet seul (pas [objet]) quand une
      // salle n'a de données que sur une seule semaine — on normalise toujours en tableau.
      const arr = Array.isArray(raw) ? raw : [raw];
      arr.forEach(w => {
        if (!parSemaine[w.label]) parSemaine[w.label] = { label:w.label, annee:w.annee, semaine:w.semaine, heures:0 };
        parSemaine[w.label].heures += w.heures;
      });
    });
    let weeks = Object.values(parSemaine).map(w => {
      const njs = nbJoursSemaineLookup[w.label] || 5;
      const dispo = Math.max(1,nbSallesFiltre) * Math.max(1,njs) * 10;
      return { ...w, tauxPct: dispo ? w.heures/dispo*100 : 0, lundi: lundiLookup[w.label] || '' };
    }).sort((a,b)=> a.annee-b.annee || a.semaine-b.semaine);
    if (semestre === 'S1') weeks = weeks.filter(w => w.annee === 2025);
    if (semestre === 'S2') weeks = weeks.filter(w => w.annee === 2026);
    if (semestre === 'EX') weeks = weeks.filter(w => EXAM_WEEKS.has(w.label));
    if (semestre === 'HX') weeks = weeks.filter(w => !EXAM_WEEKS.has(w.label));

    const chartEl = document.getElementById('weekly-chart');
    if (!weeks.length){ chartEl.innerHTML = '<p class="opp-note">Aucune donnée pour ce filtre.</p>'; }
    else {
      // Échelle absolue (0-100%) : la hauteur d'une barre est directement proportionnelle
      // à son taux réel, pas relative au maximum du jeu filtré (sinon deux semaines proches
      // en % pourraient paraître très différentes).
      chartEl.innerHTML = weeks.map(w => {
        const sem = w.semaine <= 25 ? 's2' : 's1';
        const h = Math.min(100, Math.max(1, w.tauxPct));
        return `<div class="wcol" tabindex="0" aria-label="Semaine ${w.label}, ${fmt(w.tauxPct,1)}%">
          <span class="wval">${fmt(w.tauxPct,0)}%</span>
          <span class="wbar ${sem}" style="height:${h}%" title="${w.label} (lundi ${w.lundi}) : ${fmt(w.tauxPct,1)}%"></span>
          <span class="wtick">S${w.semaine}</span>
        </div>`;
      }).join('');
    }

    // Heatmap jour×heure = moyenne des % par salle sur l'ensemble filtré. C'est exact (pas
    // une approximation) : dans heatmap_par_salle*.json chaque salle est déjà divisée par le
    // même dénominateur (jours ouvrés du jour concerné, sur la période choisie), donc la
    // moyenne des % équivaut à la vraie agrégation minutes-occupées / minutes-disponibles.
    // La source change selon le sélecteur Année/1er sem./2nd sem.
    const heatmapSource = semestre === 'S1' ? heatmapSalleS1 : semestre === 'S2' ? heatmapSalleS2 : semestre === 'EX' ? heatmapSalleEX : semestre === 'HX' ? heatmapSalleHX : heatmapSalle;
    const keysFiltre = scopedAll.map(d => `${d.salle}|${d.batiment}`);
    function moyenneHeatmap(jour, label){
      if (!keysFiltre.length) return 0;
      let sum = 0;
      keysFiltre.forEach(k => { sum += (heatmapSource[k] && heatmapSource[k][jour] && heatmapSource[k][jour][label]) || 0; });
      return sum / keysFiltre.length;
    }
    const grid = document.getElementById('hm-grid');
    let html = '<div></div>' + HOURS.map(h=>`<div class="hm-h">${h}h-${h+1}h</div>`).join('');
    ['lundi','mardi','mercredi','jeudi','vendredi'].forEach(jour => {
      html += `<div class="hm-day">${jour.charAt(0).toUpperCase()+jour.slice(1)}</div>`;
      HOURS.forEach(h => {
        const label = `${h}h-${h+1}h`;
        const pct = moyenneHeatmap(jour, label);
        html += `<div class="hm-cell ${heatClass(pct)}" title="${jour} ${label} : ${fmt(pct,1)}%">${pct.toFixed(0)}</div>`;
      });
    });
    grid.innerHTML = html;
    grid.querySelectorAll('.hm-cell').forEach(c => { c.style.background = `var(--${c.className.split(' ')[1]})`; });

    const peak = weeks.slice().sort((a,b)=>b.tauxPct-a.tauxPct).slice(0,5).map(w => ({ label: w.label, value: w.tauxPct, fullLabel:`Semaine ${w.label} (lundi ${w.lundi}) : ${fmt(w.tauxPct,1)}%` }));
    renderBarlist(document.getElementById('chart-peakweeks'), peak, { suffix:'%', decimals:1 });

    const joursLabels = { lundi:'Lundi', mardi:'Mardi', mercredi:'Mercredi', jeudi:'Jeudi', vendredi:'Vendredi' };
    const peakDays = Object.entries(joursLabels).map(([jour,label]) => {
      const vals = HOURS.map(h => moyenneHeatmap(jour, `${h}h-${h+1}h`));
      const moy = vals.reduce((s,v)=>s+v,0) / vals.length;
      return { label, value: moy, fullLabel: `${label} : ${fmt(moy,1)}% d'occupation moyenne (8h-18h)` };
    }).sort((a,b) => b.value - a.value);
    renderBarlist(document.getElementById('chart-peakdays'), peakDays, { suffix:'%', decimals:1 });
  }

  // ================= TAB 4 : Synthèse exécutive =================
  function renderExec(){
    const scopedAll = data.filter(matchesGlobal);
    const scoped = scopedAll.filter(d => eTaux(d) !== null && eTaux(d) !== undefined);
    const tauxMoyen = scoped.length ? scoped.reduce((s,d)=>s+eTaux(d),0) / scoped.length : 0;
    const totalSeances = scopedAll.reduce((s,d)=>s+(d.nbCours||0),0);

    document.getElementById('exec-jours').textContent = semestre==='S1' ? '69' : semestre==='S2' ? '69' : semestre==='EX' ? '10' : semestre==='HX' ? '128' : '138';
    document.getElementById('exec-creneaux').textContent = semestre==='S1' ? String(69*10) : semestre==='S2' ? String(69*10) : semestre==='EX' ? String(10*10) : semestre==='HX' ? String(128*10) : String(138*10);
    document.getElementById('exec-creneaux-l').textContent = semestre ? `créneaux disponibles / salle (h), ${semLabel()}` : 'créneaux disponibles / salle (h)';
    document.getElementById('exec-seances').textContent = totalSeances.toLocaleString('fr-FR');
    document.getElementById('exec-taux').innerHTML = fmt(tauxMoyen,1) + '%';
    document.querySelectorAll('.exec-num .l')[4].textContent = `taux d'occupation moyen (${semLabel()})`;

    const scopedKeys = new Set(scopedAll.map(d => `${d.salle}|${d.batiment}`));
    let conflits = 0, salleAvecConflitMax = null, conflitMax = 0;
    Object.entries(heatmapSalle).forEach(([key, jours]) => {
      if (!scopedKeys.has(key)) return;
      Object.values(jours).forEach(buckets => {
        Object.values(buckets).forEach(pct => {
          if (pct > 100) { conflits++; if (pct > conflitMax) { conflitMax = pct; salleAvecConflitMax = key; } }
        });
      });
    });
    document.getElementById('exec-conflits').textContent = conflits.toLocaleString('fr-FR');

    const byBat = {};
    scoped.forEach(d => { if (!d.batiment) return; (byBat[d.batiment] ??= []).push(eTaux(d)); });
    const ranking = Object.entries(byBat).map(([b,arr]) => ({ b, taux: arr.reduce((s,v)=>s+v,0)/arr.length, n: arr.length }))
      .filter(r => r.n >= 3).sort((a,b) => b.taux - a.taux);
    const top5 = ranking.slice(0,5), bottom5 = ranking.slice(-5).reverse();

    document.getElementById('exec-top5').innerHTML = top5.map(r => `<tr><td>${r.b}</td><td class="num">${fmt(r.taux,1)}%</td></tr>`).join('');
    document.getElementById('exec-bottom5').innerHTML = bottom5.map(r => `<tr><td>${r.b}</td><td class="num">${fmt(r.taux,1)}%</td></tr>`).join('');

    const opp5 = scopedAll.filter(d => (d.capa||0)>=40 && eTaux(d)!==null && eTaux(d)<10).sort((a,b)=>(b.capa||0)-(a.capa||0)).slice(0,5);
    document.getElementById('exec-opp5').innerHTML = opp5.map(d => `<tr><td>${d.salle} (${d.batiment})</td><td class="num">${fmt(d.capa,0)}</td></tr>`).join('');

    const maxSalle = scoped.slice().sort((a,b)=>eTaux(b)-eTaux(a))[0];
    const nbGrandesFaibles = scopedAll.filter(d => (d.capa||0)>=40 && eTaux(d)!==null && eTaux(d)<10).length;
    const placesGrandesFaibles = scopedAll.filter(d => (d.capa||0)>=40 && eTaux(d)!==null && eTaux(d)<10).reduce((s,d)=>s+(d.capa||0),0);
    const top2 = ranking.slice(0,2);

    document.getElementById('insight-row').innerHTML = `
      <div class="insight-card">
        <b>1.</b> Les bâtiments <b>${top2[0]?.b}</b> et <b>${top2[1]?.b}</b> affichent les taux moyens les plus élevés (<b>${fmt(top2[0]?.taux,1)}%</b> et <b>${fmt(top2[1]?.taux,1)}%</b>) — l'essentiel du parc reste en-dessous.
      </div>
      <div class="insight-card">
        <b>2.</b> ${maxSalle && eTaux(maxSalle) > 100
          ? `La salle la plus sollicitée sur ${semLabel()}, <b>${maxSalle.salle}</b> (${maxSalle.batiment}), dépasse les 100% (<b>${fmt(eTaux(maxSalle),1)}%</b>) — signe de réservations qui se chevauchent, pas d'une salle libre en permanence.`
          : `Aucune salle n'atteint 100% d'occupation sur ${semLabel()} : la plus sollicitée, <b>${maxSalle?.salle}</b> (${maxSalle?.batiment}), culmine à <b>${fmt(maxSalle?eTaux(maxSalle):null,1)}%</b>.`}
      </div>
      <div class="insight-card warn">
        <b>3.</b> <b>${nbGrandesFaibles} salles</b> de grande capacité (≥40 places, ${placesGrandesFaibles.toLocaleString('fr-FR')} places cumulées) restent sous les 10% d'occupation — un vrai potentiel de réorganisation avant toute extension.
      </div>
      <div class="insight-card crit">
        <b>4.</b> <b>${conflits.toLocaleString('fr-FR')} créneaux</b> salle/jour/heure affichent des réservations qui se chevauchent sur certaines semaines (ex. ${salleAvecConflitMax ? salleAvecConflitMax.replace('|',' — ') : '—'}) — l'occupation \"pleine\" n'exclut pas des conflits ponctuels à résoudre.
      </div>`;

    const steps = [
      ['1','Extraction CC'], ['2','Reconstitution PERIODE → dates ISO'], ['3','Contrôle multi-salles'],
      ['4','Exclusion blocages admin (≥8h)'], ['5','Dédoublonnage chevauchements (≤10h/j)'], ['6','Taux base 10h'], ['7','Heatmaps jour×heure'], ['8','Publication interactive']
    ];
    document.getElementById('pipeline').innerHTML = steps.map(([n,t],i) => `
      <div class="pstep"><div class="circle">${n}</div><div class="ptxt">${t}</div></div>
      ${i<steps.length-1 ? '<div class="pconn"></div>' : ''}
    `).join('');
  }

  // ================= TAB 5 : Chevauchements =================
  let cfSortKey = 'overlapMin', cfSortDir = 'desc', cfPage = 1;
  const cfSearch = document.getElementById('cf-search');
  const cfShowFusion = document.getElementById('cf-show-fusion');

  function cfInGlobalScope(c){
    if (semestre === 'S1' && c.annee !== 2025) return false;
    if (semestre === 'S2' && c.annee !== 2026) return false;
    if (semestre === 'EX' && !EXAM_DATES.has(c.date)) return false;
    if (semestre === 'HX' && EXAM_DATES.has(c.date)) return false;
    const info = batInfo[c.batiment];
    if (villeMsel.selected.size && !(info && villeMsel.selected.has(info.ville))) return false;
    if (siteMsel.selected.size && !(info && siteMsel.selected.has(info.site))) return false;
    if (batMsel.selected.size && !batMsel.selected.has(c.batiment)) return false;
    const gq = gSearch.value.trim().toLowerCase();
    if (gq && !(c.salle.toLowerCase().includes(gq) || (c.batiment||'').toLowerCase().includes(gq))) return false;
    return true;
  }
  function cfMatches(c){
    if (!cfInGlobalScope(c)) return false;
    if (!cfShowFusion.checked && c.sameMatiere) return false;
    const q = cfSearch.value.trim().toLowerCase();
    if (q && !(c.salle.toLowerCase().includes(q) || (c.batiment||'').toLowerCase().includes(q) ||
        c.matiere1.toLowerCase().includes(q) || c.matiere2.toLowerCase().includes(q))) return false;
    return true;
  }

  function renderConflits(){
    const scopedAll = conflicts.filter(cfInGlobalScope);
    const fusionCount = scopedAll.filter(c => c.sameMatiere).length;
    const filtered = conflicts.filter(cfMatches);

    document.getElementById('cf-scope').textContent = semLabel().charAt(0).toUpperCase() + semLabel().slice(1);
    document.getElementById('cf-kpi-total').textContent = scopedAll.length.toLocaleString('fr-FR');
    document.getElementById('cf-kpi-salles').textContent = new Set(scopedAll.map(c=>`${c.salle}|${c.batiment}`)).size;
    document.getElementById('cf-kpi-motifs').textContent = new Set(scopedAll.map(c=>`${c.salle}|${c.batiment}|${c.jour}|${c.hdebut1}|${c.hfin1}|${c.matiere1}|${c.hdebut2}|${c.hfin2}|${c.matiere2}`)).size.toLocaleString('fr-FR');
    document.getElementById('cf-kpi-fusion').textContent = fusionCount.toLocaleString('fr-FR');

    const n = filtered.length;
    const parts = [];
    if (cfSearch.value.trim()) parts.push(`"${cfSearch.value.trim()}"`);
    if (!cfShowFusion.checked) parts.push('groupes fusionnés masqués');
    document.getElementById('cf-note').textContent = `${n} chevauchement${n>1?'s':''}` + (parts.length ? ` — ${parts.join(' · ')}` : '');

    let rows = filtered.slice();
    rows.sort((a,b) => {
      let av = a[cfSortKey], bv = b[cfSortKey];
      if (typeof av === 'string'){ av=av.toLowerCase(); bv=bv.toLowerCase(); }
      const cmp = av<bv?-1:av>bv?1:0; return cfSortDir==='asc'?cmp:-cmp;
    });
    const size = parseInt(document.getElementById('cf-size').value,10);
    const totalPages = Math.max(1, Math.ceil(rows.length/size));
    cfPage = Math.min(cfPage, totalPages);
    const pageRows = rows.slice((cfPage-1)*size, (cfPage-1)*size+size);
    document.getElementById('cf-info').textContent = `Page ${cfPage} / ${totalPages}`;
    document.getElementById('cf-prev').disabled = cfPage<=1;
    document.getElementById('cf-next').disabled = cfPage>=totalPages;

    document.getElementById('cf-table-body').innerHTML = pageRows.map(c => `
      <tr>
        <td class="salle">${c.salle}</td><td>${c.batiment||'—'}</td><td>${c.date}</td><td>${c.jour}</td>
        <td class="salle">${c.hdebut1}–${c.hfin1}<br><span style="color:var(--ink-muted);font-size:.82em">${c.matiere1}${c.diplome1?` · ${c.diplome1}`:''}</span></td>
        <td class="salle">${c.hdebut2}–${c.hfin2}<br><span style="color:var(--ink-muted);font-size:.82em">${c.matiere2}${c.diplome2?` · ${c.diplome2}`:''}</span></td>
        <td class="num">${c.overlapMin} min</td>
      </tr>`).join('');
  }
  document.querySelectorAll('#tab-conflits thead th[data-key]').forEach(th => {
    th.addEventListener('click', () => {
      const key = th.dataset.key;
      if (cfSortKey===key) cfSortDir = cfSortDir==='asc'?'desc':'asc'; else { cfSortKey=key; cfSortDir = (key==='overlapMin')?'desc':'asc'; }
      document.querySelectorAll('#tab-conflits thead th').forEach(h => { delete h.dataset.active; if(h.querySelector('.arrow')) h.querySelector('.arrow').textContent='↕'; });
      th.dataset.active=''; if(th.querySelector('.arrow')) th.querySelector('.arrow').textContent = cfSortDir==='asc'?'↑':'↓';
      cfPage = 1; renderConflits();
    });
  });
  document.getElementById('cf-size').addEventListener('change', () => { cfPage=1; renderConflits(); });
  document.getElementById('cf-prev').addEventListener('click', () => { cfPage--; renderConflits(); });
  document.getElementById('cf-next').addEventListener('click', () => { cfPage++; renderConflits(); });
  cfSearch.addEventListener('input', () => { cfPage=1; renderConflits(); });
  cfShowFusion.addEventListener('change', () => { cfPage=1; renderConflits(); });

  // ================= Orchestration =================
  function renderAll(){
    updateNote();
    renderOverview();
    renderGrid();
    renderWeekly();
    renderExec();
    renderConflits();
  }
  gSearch.addEventListener('input', () => { page=1; gridPage=1; renderAll(); });
  gReset.addEventListener('click', () => {
    gSearch.value='';
    [villeMsel, siteMsel, batMsel, tauxMsel].forEach(m => m.selected.clear());
    document.querySelectorAll('.msel-list input[type="checkbox"]').forEach(cb => cb.checked = false);
    [villeMsel, siteMsel, batMsel, tauxMsel].forEach(m => m.updateLabel());
    refreshSiteOptions(); refreshBatimentOptions(); page=1; gridPage=1; renderAll();
  });

  renderAll();
})();
