let meta = null;
let filters = {};
let filterValues = {};
let msComponents = [];

function positionPopover(btnEl, panelEl){
  const r = btnEl.getBoundingClientRect();
  panelEl.style.position = 'fixed';
  const vw = Math.max(document.documentElement.clientWidth, window.innerWidth || 0);
  const vh = Math.max(document.documentElement.clientHeight, window.innerHeight || 0);
  // Default: match trigger width, but allow custom popover width (e.g., filter picker)
  const desired = parseInt(panelEl.dataset.popWidth || '', 10);
  let width = (Number.isFinite(desired) && desired > 0) ? desired : Math.round(r.width);

  // Special-case: keep the "filter picker" popover anchored to the LEFT of the trigger.
  // If it doesn't fit, shrink its width instead of shifting it to the right (prevents "jump" between clicks).
  const lockLeft = panelEl.classList.contains('chooser') || panelEl.dataset.lockLeft === '1';
  let left = Math.round(r.left);
  if(lockLeft){
    const maxW = Math.max(260, Math.round(vw - left - 8));
    width = Math.min(width, maxW);
    left = Math.max(8, left);
  } else {
    // Clamp inside viewport so popover doesn't "jump" to strange places
    left = Math.min(Math.max(8, left), Math.max(8, vw - width - 8));
  }
  const top = Math.round(r.bottom) + 6;
  panelEl.style.left = `${left}px`;
  panelEl.style.top = `${top}px`;
  // Keep a stable width for special popovers (e.g. "Выбрать фильтры")
  panelEl.style.width = `${width}px`;
  if(Number.isFinite(desired) && desired > 0){
    panelEl.style.minWidth = `${width}px`;
    panelEl.style.maxWidth = `${width}px`;
  } else {
    panelEl.style.minWidth = '';
    panelEl.style.maxWidth = '';
  }
  panelEl.style.right = 'auto';

  // Height: fit within viewport
  const maxH = Math.max(180, vh - top - 12);
  panelEl.style.maxHeight = `${maxH}px`;
}

function closeAllPopovers(exceptEl){
  document.querySelectorAll('.ms.open').forEach(el=>{
    if(exceptEl && el === exceptEl) return;
    el.classList.remove('open');
  });
}

function qs(id){ return document.getElementById(id); }


function setTableWrapHeight(){
  const tw = qs('tableWrap');
  if(!tw) return;
  const r = tw.getBoundingClientRect();
  const h = Math.max(220, Math.round(window.innerHeight - r.top - 8));
  tw.style.height = h + 'px';
}

function collapseFilters(force){
  const panel = qs('filtersPanel');
  if(!panel) return;
  panel.classList.add('collapsed');
  const btn = qs('toggleFiltersBtn');
  if(btn) btn.textContent = 'Показать фильтры';
  // close any open popovers so they don't hang over the grid
  closeAllPopovers();
  setTableWrapHeight();
}

function expandFilters(){
  const panel = qs('filtersPanel');
  if(!panel) return;
  panel.classList.remove('collapsed');
  const btn = qs('toggleFiltersBtn');
  if(btn) btn.textContent = 'Скрыть фильтры';
  setTableWrapHeight();
}

function setupFiltersCollapse(){
  const panel = qs('filtersPanel');
  const btn = qs('toggleFiltersBtn');
  const tw = qs('tableWrap');
  if(btn && panel){
    btn.addEventListener('click', ()=>{
      if(panel.classList.contains('collapsed')) expandFilters();
      else collapseFilters();
    });
  }
  if(tw){
    let collapseT = null;
    tw.addEventListener('scroll', ()=>{
      // Hide filters smoothly when user starts working with the table
      if(tw.scrollTop > 6){
        if(collapseT) clearTimeout(collapseT);
        collapseT = setTimeout(()=>collapseFilters(), 90);
      }
      // Close popovers on any scroll so dropdowns don't hang over the grid
      closeAllPopovers();
    }, {passive:true});
  }

  // If the page itself is scrolled (e.g. on smaller screens), hide filters too
  window.addEventListener('scroll', ()=>{
    if(window.scrollY > 10) collapseFilters();
    closeAllPopovers();
  }, {passive:true});
  window.addEventListener('resize', ()=>setTableWrapHeight());
  // initial
  setTimeout(setTableWrapHeight, 0);
}

// Snap horizontal scrolling to one time-column (30 min) at a time.
// This makes the grid "step" by columns instead of stopping mid-cell.
function setupHorizontalSnap(){
  const tw = qs('tableWrap');
  const grid = qs('grid');
  if(!tw || !grid) return;

  if(tw.dataset.snapInit === '1') return;
  tw.dataset.snapInit = '1';

  let t = null;
  function slotWidth(){
    const first = grid.querySelector('col[data-col-key^="t_"]');
    if(!first) return 0;
    const w = parseInt(String(first.style.width || '').replace('px',''), 10);
    if(Number.isFinite(w) && w > 10) return w;
    // fallback
    const th = grid.querySelector('thead th[data-col-key^="t_"]');
    return th ? Math.round(th.getBoundingClientRect().width) : 0;
  }

  function snap(){
    const w = slotWidth();
    if(!w) return;
    const x = tw.scrollLeft;
    const snapped = Math.round(x / w) * w;
    if(Math.abs(snapped - x) >= 1){
      tw.scrollLeft = snapped;
    }
  }

  tw.addEventListener('scroll', ()=>{
    // debounce snap so the user can scroll freely, then we align to the nearest column
    if(t) clearTimeout(t);
    t = setTimeout(snap, 120);
  }, {passive:true});
}

async function loadMeta(){
  await loadPalette();
  const r = await fetch('/api/meta');
  meta = await r.json();
  if(!meta.has_data){
    qs('status').textContent = 'Файл ещё не загружен.';
    return;
  }
  filterValues = meta.values || {};
  buildDaySelect(meta.dates);
  buildFilters(meta.columns, filterValues);
}

function buildDaySelect(dates){
  const sel = qs('daySelect');
  sel.innerHTML = '';
  (dates || []).forEach(d=>{
    const opt = document.createElement('option');
    opt.value = d;
    opt.textContent = formatDateRu(d);
    sel.appendChild(opt);
  });
  if(dates && dates.length){ sel.value = dates[0]; }
}

function buildFilters(columns, valuesMap){
  const wrap = qs('filtersWrap');
  wrap.innerHTML = '';
  filters = {};
  msComponents = [];

  (columns || []).forEach(col=>{
    const div = document.createElement('div');
    div.className='filterItem';
    div.dataset.col = col;

    const label = document.createElement('label');
    label.textContent = col;

    const values = (valuesMap && valuesMap[col]) ? valuesMap[col] : [];
    const ms = makeMultiSelect(col, values, (selected)=>{
      filters[col] = selected; // empty => "all"
    });
    filters[col] = [];
    msComponents.push(ms);

    div.appendChild(label);
    div.appendChild(ms);
    wrap.appendChild(div);
  });

  buildFilterPicker(columns || []);
}

function defaultVisibleFilters(columns){
  // По умолчанию показываем ровно то, что нужно диспетчеру
  const preferred = [
    'Номер заявки',
    'Категория',
    'Подкатегория',
    'Статус',
    'ТС',
    'Гос номер ТС',
    'Класс назначенного ТС',
    'Водитель',
    'Подразделение',
    'Аварийность'
  ];
  const out = [];
  preferred.forEach(p=>{ if(columns.includes(p)) out.push(p); });
  if(out.length) return out;
  return columns.slice(0, Math.min(8, columns.length));
}

function buildFilterPicker(columns){
  const mount = qs('filterPicker');
  if(!mount) return;
  mount.innerHTML = '';

  // Use a versioned key so old UI state doesn't break defaults after updates
  const STORAGE_KEY = 'visibleFilters_v2';
  let selected = [];
  try{
    selected = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
  }catch(e){ selected = []; }

  // Always start from the dispatcher-friendly default set on a fresh install
  // (and also when saved config is too broad / invalid).
  const fallback = defaultVisibleFilters(columns);
  if(!Array.isArray(selected) || selected.length === 0){
    selected = fallback;
    localStorage.setItem(STORAGE_KEY, JSON.stringify(selected));
  } else {
    // sanitize (remove unknown columns)
    selected = selected.filter(c=>columns.includes(c));
    // if user previously had "everything" selected, reset to sane defaults
    if(selected.length > 14){
      selected = fallback;
    }
    localStorage.setItem(STORAGE_KEY, JSON.stringify(selected));
  }

  const ms = document.createElement('div');
  ms.className = 'ms';

  const btn = document.createElement('div');
  btn.className = 'ms-btn';
  const label = document.createElement('div');
  label.textContent = 'Выбрать фильтры';
  const right = document.createElement('div');
  right.style.display='flex';
  right.style.alignItems='center';
  right.style.gap='8px';
  const pill = document.createElement('span');
  pill.className='ms-pill';
  const caret = document.createElement('span');
  caret.textContent='▾';
  caret.style.color='var(--muted)';
  right.appendChild(pill);
  right.appendChild(caret);
  btn.appendChild(label);
  btn.appendChild(right);

  const panel = document.createElement('div');
  panel.className='ms-panel chooser';
  panel.dataset.popWidth = '460';

  const actions = document.createElement('div');
  actions.className='ms-actions';
  const leftBtns = document.createElement('div');
  leftBtns.style.display='flex';
  leftBtns.style.gap='8px';
  const bAll = document.createElement('button');
  bAll.type='button';
  bAll.textContent='Все';
  const bNone = document.createElement('button');
  bNone.type='button';
  bNone.textContent='Скрыть все';
  leftBtns.appendChild(bAll);
  leftBtns.appendChild(bNone);
  actions.appendChild(leftBtns);

  const search = document.createElement('input');
  search.type='text';
  search.placeholder='Поиск фильтра...';
  search.className='ms-search';
  actions.appendChild(search);

  const list = document.createElement('div');
  list.className='ms-list';

  function apply(){
    localStorage.setItem(STORAGE_KEY, JSON.stringify(selected));
    pill.textContent = String(selected.length);
    // show/hide controls
    const items = document.querySelectorAll('#filtersWrap .filterItem');
    items.forEach(it=>{
      const c = it.dataset.col;
      if(!c) return;
      it.classList.toggle('hidden', !selected.includes(c));
    });
  }

  function renderList(){
    list.innerHTML='';
    const q = (search.value || '').toLowerCase();
    columns
      .filter(c=> !q || String(c).toLowerCase().includes(q))
      .forEach(c=>{
        const row = document.createElement('label');
        row.className='ms-option';
        const cb = document.createElement('input');
        cb.type='checkbox';
        cb.checked = selected.includes(c);
        cb.addEventListener('change', ()=>{
          if(cb.checked){
            if(!selected.includes(c)) selected.push(c);
          }else{
            selected = selected.filter(x=>x!==c);
          }
          apply();
        });
        const text = document.createElement('span');
        text.textContent = c;
        row.appendChild(cb);
        row.appendChild(text);
        list.appendChild(row);
      });
  }

  bAll.addEventListener('click', (e)=>{ e.preventDefault(); selected = columns.slice(); renderList(); apply(); });
  bNone.addEventListener('click', (e)=>{ e.preventDefault(); selected = []; renderList(); apply(); });
  search.addEventListener('input', renderList);

  panel.appendChild(actions);
  panel.appendChild(list);

  btn.addEventListener('click', (e)=>{
    e.stopPropagation();
    const willOpen = !ms.classList.contains('open');
    closeAllPopovers(willOpen ? ms : null);
    ms.classList.toggle('open');
    if(ms.classList.contains('open')){
      positionPopover(btn, panel);
      search.focus();
    }
  });
  document.addEventListener('click', (e)=>{
    if(!ms.contains(e.target)) ms.classList.remove('open');
  });
  const rep = ()=>{ if(ms.classList.contains('open')) positionPopover(btn, panel); };
  window.addEventListener('resize', rep);
  window.addEventListener('scroll', rep, true);

  ms.appendChild(btn);
  ms.appendChild(panel);
  mount.appendChild(ms);

  renderList();
  apply();
}

async function loadPalette(){
  try{
    const r = await fetch('/api/palette');
    if(!r.ok) return;
    const p = await r.json();
    applyPalette(p);
  }catch(e){ /* ignore */ }
}

function applyPalette(p){
  const root = document.documentElement;
  if(p.accent) root.style.setProperty('--accent', p.accent);
  if(p.schedule_fill) root.style.setProperty('--fillSchedule', p.schedule_fill);
  if(p.plan_fill) root.style.setProperty('--fillPlan', p.plan_fill);
  if(p.fact_fill) root.style.setProperty('--fillFact', p.fact_fill);
  if(p.border) root.style.setProperty('--gridBorder', p.border);
}

async function upload(){
  const f = qs('file').files[0];
  if(!f){ qs('status').textContent='Выберите файл .xlsx'; return; }
  const fd = new FormData();
  fd.append('file', f);
  qs('status').textContent='Загружаю...';
  const r = await fetch('/api/upload', {method:'POST', body: fd});
  const j = await r.json();
  if(!j.ok){ qs('status').textContent='Ошибка загрузки'; return; }
  qs('status').textContent='Загружено.';
  filterValues = j.values || {};
  buildDaySelect(j.dates);
  buildFilters(j.columns, filterValues);
  await refresh();
}

async function refresh(){
  const day = qs('daySelect').value;
  if(!day){ qs('status').textContent='Выберите дату'; return; }
  const params = new URLSearchParams();
  params.set('day', day);
  params.set('filters', JSON.stringify(filters));
  qs('status').textContent='Считаю...';
  const r = await fetch('/api/schedule?' + params.toString());
  const data = await r.json();
  qs('status').textContent = `Показано заявок: ${data.filtered_count} / ${data.total_count}`;
  renderGrid(data);
}

function renderGrid(data){
  const grid = qs('grid');
  grid.innerHTML = '';

  // Use <colgroup> so that resizing affects BOTH header and body (fixes sticky overlap & "can't resize" issues)
  const colgroup = document.createElement('colgroup');
  function addCol(key, fallback){
    const col = document.createElement('col');
    col.dataset.colKey = key;
    const w = getSavedWidth(key, fallback);
    col.style.width = w + 'px';
    colgroup.appendChild(col);
  }

  const thead = document.createElement('thead');
  const hr = document.createElement('tr');

  const leftHeaders = [
    // стартовые ширины максимально компактные (диспетчер сможет расширить)
    {t:'ТС', cls:'st1', w:110},
    {t:'Гос номер ТС', cls:'st2', w:80},
    {t:'Класс ТС', cls:'st3', w:160},
    {t:'График работы', cls:'st4', w:105},
    {t:'Режим работы', cls:'st5', w:105},
    {t:'', cls:'st6', w:82},
  ];

  // cols for sticky block
  leftHeaders.forEach(h=> addCol(h.cls, h.w));
  // cols for time slots
  data.slots.forEach(slot=> addCol('t_' + slot, 44));

  grid.appendChild(colgroup);

  leftHeaders.forEach(h=>{
    const th = document.createElement('th');
    th.textContent = h.t;
    th.className = h.cls + ' resizable';
    th.dataset.colKey = h.cls;
    hr.appendChild(th);
  });

  data.slots.forEach(slot=>{
    const th = document.createElement('th');
    th.textContent = slot;
    th.classList.add('resizable');
    th.dataset.colKey = 't_' + slot;
    hr.appendChild(th);
  });
  thead.appendChild(hr);
  grid.appendChild(thead);

  const tbody = document.createElement('tbody');

  data.vehicles.forEach((v, vidx)=>{
    const plate = v.vehicle_plate;
    const regime = (v.regime_start && v.regime_end) ? `${v.regime_start} - ${v.regime_end}` : '';

    const layers = [
      {label:'График работы', kind:'schedule', get:(slot)=> (data.schedule?.[plate]?.[slot] ? 1 : 0)},
      {label:'План', kind:'plan', get:(slot)=> (data.plan?.[plate]?.[slot] || [])},
      {label:'Факт', kind:'fact', get:(slot)=> (data.fact?.[plate]?.[slot] ? 1 : 0)},
    ];

    layers.forEach((layer, idx)=>{
      const tr = document.createElement('tr');
      if(vidx % 2 === 1) tr.classList.add('groupAlt');

      if(idx===0){
        tr.appendChild(makeSpanCell(v.vehicle_name || '', 3, 'st1'));
        tr.appendChild(makeSpanCell(plate, 3, 'st2'));
        tr.appendChild(makeSpanCell(v.vehicle_class || '', 3, 'st3'));
        tr.appendChild(makeSpanCell(v.schedule_text || '', 3, 'st4'));
        tr.appendChild(makeSpanCell(regime, 3, 'st5'));
      }

      const layerTd = document.createElement('td');
      layerTd.textContent = layer.label;
      layerTd.className = 'st6 layer';
      tr.appendChild(layerTd);

      data.slots.forEach(slot=>{
        const td = document.createElement('td');
        td.className = 'slot';
        const val = layer.get(slot);

        const hasVal = Array.isArray(val) ? (val.length>0) : !!val;
        if(layer.kind==='schedule' && hasVal){ td.classList.add('schedule'); }
        if(layer.kind==='plan' && hasVal){
          td.classList.add('plan');
          // Show request number(s) inside filled plan cells
          const nums = val;
          const first = String(nums[0]);
          const extra = nums.length>1 ? ` +${nums.length-1}` : '';
          const label = `№${first}${extra}`;
          const span = document.createElement('span');
          span.className='planTag';
          span.textContent = label;
          td.appendChild(span);
          if(nums.length>1){ td.title = nums.map(n=>`№${n}`).join(', '); }
        }
        if(layer.kind==='fact' && hasVal){ td.classList.add('fact'); }

        if(layer.kind==='schedule' || layer.kind==='fact'){
          td.style.cursor='pointer';
          td.dataset.day = data.day;
          td.dataset.plate = plate;
          td.dataset.kind = layer.kind;
          td.dataset.slot = slot;
          td.dataset.current = String(hasVal ? 1 : 0);
          td.addEventListener('click', ()=>toggle(td));
        } else {
          td.style.cursor='not-allowed';
        }
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
  });

  grid.appendChild(tbody);

  setupColumnResizers(grid);
  requestAnimationFrame(()=>requestAnimationFrame(()=>updateStickyVars(grid)));
  // Enable step-by-step horizontal scrolling (snap to one time column)
  setupHorizontalSnap();
}

function getSavedWidth(key, fallback){
  const saved = key ? localStorage.getItem('colw:' + key) : null;
  const w = saved ? parseInt(saved, 10) : fallback;
  return (Number.isFinite(w) && w > 10) ? w : fallback;
}

function setupColumnResizers(table){
  const ths = table.querySelectorAll('thead th.resizable');
  ths.forEach(th=>{
    if(th.querySelector('.col-resizer')) return;
    const grip = document.createElement('div');
    grip.className='col-resizer';
    th.appendChild(grip);

    let startX = 0;
    let startW = 0;
    const key = th.dataset.colKey || th.textContent;

    function onMove(e){
      const dx = e.clientX - startX;
      const minW = (String(key).startsWith('t_')) ? 26 : 60;
      const newW = Math.max(minW, startW + dx);
      // Update both header and body via colgroup
      const col = table.querySelector(`col[data-col-key="${CSS.escape(String(key))}"]`);
      if(col) col.style.width = newW + 'px';
      if(key) localStorage.setItem('colw:' + key, String(newW));
      updateStickyVars(table);
    }
    function onUp(){
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
    }
    grip.addEventListener('mousedown', (e)=>{
      e.preventDefault();
      startX = e.clientX;
      startW = th.getBoundingClientRect().width;
      document.addEventListener('mousemove', onMove);
      document.addEventListener('mouseup', onUp);
    });
  });
}

function updateStickyVars(table){
  const colW = (key)=>{
    const col = table.querySelector(`col[data-col-key="${CSS.escape(String(key))}"]`);
    if(!col) return null;
    const w = parseInt(String(col.style.width || '').replace('px',''), 10);
    if(Number.isFinite(w) && w > 0) return w;
    // fallback to computed width
    const cw = Math.round(col.getBoundingClientRect().width);
    return cw || null;
  }
  const w1 = colW('st1') || 0;
  const w2 = colW('st2') || 0;
  const w3 = colW('st3') || 0;
  const w4 = colW('st4') || 0;
  const w5 = colW('st5') || 0;
  const w6 = colW('st6') || 0;

  // Update CSS vars (used by fallback CSS) AND also apply inline left offsets
  // Inline offsets are more robust across browser zoom/scaling and prevent overlap glitches.
  const root = document.documentElement;
  root.style.setProperty('--st1w', w1 + 'px');
  root.style.setProperty('--st2w', w2 + 'px');
  root.style.setProperty('--st3w', w3 + 'px');
  root.style.setProperty('--st4w', w4 + 'px');
  root.style.setProperty('--st5w', w5 + 'px');
  root.style.setProperty('--st6w', w6 + 'px');

  const lefts = {
    st1: 0,
    st2: w1,
    st3: w1 + w2,
    st4: w1 + w2 + w3,
    st5: w1 + w2 + w3 + w4,
    st6: w1 + w2 + w3 + w4 + w5,
  };
  Object.keys(lefts).forEach(cls=>{
    const x = Math.round(lefts[cls]);
    table.querySelectorAll('th.' + cls + ', td.' + cls).forEach(cell=>{
      cell.style.left = x + 'px';
    });
  });
}

function makeSpanCell(text, rowspan, cls){
  const td = document.createElement('td');
  td.textContent = text;
  td.rowSpan = rowspan;
  td.className = cls;
  return td;
}

async function toggle(td){
  const current = (td.dataset.current === '1');
  const payload = {
    day: td.dataset.day,
    plate: td.dataset.plate,
    kind: td.dataset.kind,
    slot: td.dataset.slot,
    value: current ? 0 : 1
  };
  const r = await fetch('/api/mark', {
    method:'POST',
    headers: {'Content-Type':'application/json'},
    body: JSON.stringify(payload)
  });
  if(r.ok){ await refresh(); }
}

function exportXlsx(){
  const day = qs('daySelect').value;
  const params = new URLSearchParams();
  params.set('day', day);
  params.set('filters', JSON.stringify(filters));
  window.location.href = '/export?' + params.toString();
}

function formatDateRu(iso){
  if(!iso) return '';
  const m = String(iso).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if(!m) return String(iso);
  return `${m[3]}.${m[2]}.${m[1]}`;
}

// --- Multi-select ---
function makeMultiSelect(colName, values, onChange){
  const root = document.createElement('div');
  root.className='ms';

  const btn = document.createElement('div');
  btn.className='ms-btn';
  btn.tabIndex = 0;
  btn.innerHTML = `<span class="ms-label">Все</span><span class="ms-pill" style="display:none"></span>`;

  const panel = document.createElement('div');
  panel.className='ms-panel';

  const actions = document.createElement('div');
  actions.className='ms-actions';

  const leftActions = document.createElement('div');
  leftActions.style.display='flex';
  leftActions.style.gap='8px';

  const allBtn = document.createElement('button');
  allBtn.type='button';
  allBtn.textContent='Все';
  const clearBtn = document.createElement('button');
  clearBtn.type='button';
  clearBtn.textContent='Очистить';

  leftActions.appendChild(allBtn);
  leftActions.appendChild(clearBtn);

  const count = document.createElement('div');
  count.className='small';
  count.textContent='0 выбрано';

  actions.appendChild(leftActions);
  actions.appendChild(count);

  const search = document.createElement('input');
  search.type='text';
  search.className='ms-search';
  search.placeholder='Поиск...';

  const list = document.createElement('div');
  list.className='ms-list';

  panel.appendChild(actions);
  panel.appendChild(search);
  panel.appendChild(list);

  const state = {
    selected: new Set(),
    filtered: values || []
  };

  function renderList(){
    list.innerHTML='';
    const q = (search.value || '').trim().toLowerCase();
    const shown = (values || []).filter(v => !q || String(v).toLowerCase().includes(q));
    shown.forEach(v=>{
      const row = document.createElement('label');
      row.className='ms-item';
      const cb = document.createElement('input');
      cb.type='checkbox';
      cb.checked = state.selected.has(v);
      cb.addEventListener('change', ()=>{
        if(cb.checked) state.selected.add(v); else state.selected.delete(v);
        sync();
      });
      const sp = document.createElement('span');
      sp.textContent = v;
      row.appendChild(cb);
      row.appendChild(sp);
      list.appendChild(row);
    });
  }

  function sync(){
    const n = state.selected.size;
    const label = btn.querySelector('.ms-label');
    const pill = btn.querySelector('.ms-pill');
    if(n===0){
      label.textContent='Все';
      pill.style.display='none';
    } else {
      label.textContent='Выбрано';
      pill.style.display='inline-block';
      pill.textContent=String(n);
    }
    count.textContent = `${n} выбрано`;
    onChange(Array.from(state.selected));
  }

  btn.addEventListener('click', (e)=>{
    e.stopPropagation();
    const willOpen = !root.classList.contains('open');
    closeAllPopovers(willOpen ? root : null);
    root.classList.toggle('open');
    if(root.classList.contains('open')) positionPopover(btn, panel);
  });
  document.addEventListener('click', (e)=>{
    if(!root.contains(e.target)) root.classList.remove('open');
  });
  const rep = ()=>{ if(root.classList.contains('open')) positionPopover(btn, panel); };
  window.addEventListener('resize', rep);
  window.addEventListener('scroll', rep, true);
  allBtn.addEventListener('click', ()=>{
    state.selected = new Set(values || []);
    renderList();
    sync();
  });
  clearBtn.addEventListener('click', ()=>{
    state.selected = new Set();
    renderList();
    sync();
  });
  search.addEventListener('input', renderList);

  renderList();
  sync();

  // external controls
  root._clear = () => {
    state.selected = new Set();
    renderList();
    sync();
  };
  root._selectAll = () => {
    state.selected = new Set(values || []);
    renderList();
    sync();
  };

  root.appendChild(btn);
  root.appendChild(panel);
  return root;
}

// wire
qs('uploadBtn').addEventListener('click', upload);
qs('refreshBtn').addEventListener('click', refresh);
qs('exportBtn').addEventListener('click', exportXlsx);
qs('clearAllFiltersBtn').addEventListener('click', ()=>{
  // empty selections => means "all"
  msComponents.forEach(ms=>{
    if(typeof ms._clear === 'function') ms._clear();
  });
});
setupFiltersCollapse();
loadMeta().then(()=>refresh());
