function qs(id){ return document.getElementById(id); }

async function loadPalette(){
  try{
    const r = await fetch('/api/palette');
    if(!r.ok) return;
    const p = await r.json();
    const root = document.documentElement;
    if(p.accent) root.style.setProperty('--accent', p.accent);
    if(p.schedule_fill) root.style.setProperty('--accentSoft', p.schedule_fill);
    if(p.plan_fill) root.style.setProperty('--accentSoft2', p.plan_fill);
    if(p.fact_fill) root.style.setProperty('--blueSoft', p.fact_fill);
    if(p.border) root.style.setProperty('--border', p.border);
  }catch(e){ /* ignore */ }
}

async function load(){
  await loadPalette();
  qs('dirStatus').textContent = 'Загружаю...';
  const r = await fetch('/api/directory');
  const j = await r.json();
  render(j.items || []);
  qs('dirStatus').textContent = `Строк: ${(j.items || []).length}`;
}

function render(items){
  const tbl = qs('dirTable');
  tbl.innerHTML = '';

  const thead = document.createElement('thead');
  const tr = document.createElement('tr');
  ['Гос номер','График работы','Режим начало','Режим конец',''].forEach(t=>{
    const th = document.createElement('th');
    th.textContent = t;
    tr.appendChild(th);
  });
  thead.appendChild(tr);
  tbl.appendChild(thead);

  const tbody = document.createElement('tbody');
  (items || []).forEach(it=> tbody.appendChild(row(it)));
  tbl.appendChild(tbody);
}

function row(it){
  const tr = document.createElement('tr');

  const plate = inputCell(it.vehicle_plate || '', 'AA000AA00');
  const schedule = inputCell(it.schedule_text || '', '5/2 8 часов');
  const rs = inputCell(it.regime_start || '', '07:00');
  const re = inputCell(it.regime_end || '', '18:00');

  tr.appendChild(plate.td);
  tr.appendChild(schedule.td);
  tr.appendChild(rs.td);
  tr.appendChild(re.td);

  const tdAct = document.createElement('td');
  tdAct.style.textAlign='right';

  const saveBtn = document.createElement('button');
  saveBtn.textContent = 'Сохранить';
  saveBtn.className = 'primary';
  saveBtn.addEventListener('click', async ()=>{
    await upsert({
      vehicle_plate: plate.input.value.trim(),
      schedule_text: schedule.input.value.trim(),
      regime_start: rs.input.value.trim(),
      regime_end: re.input.value.trim(),
    });
  });

  const delBtn = document.createElement('button');
  delBtn.textContent = 'Удалить';
  delBtn.style.marginLeft = '8px';
  delBtn.addEventListener('click', async ()=>{
    const p = plate.input.value.trim();
    if(!p) return;
    await fetch('/api/directory/delete', {
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body: JSON.stringify({vehicle_plate: p})
    });
    await load();
  });

  tdAct.appendChild(saveBtn);
  tdAct.appendChild(delBtn);
  tr.appendChild(tdAct);

  return tr;
}

function inputCell(value, placeholder){
  const td = document.createElement('td');
  const input = document.createElement('input');
  input.type='text';
  input.value=value;
  input.placeholder=placeholder;
  input.style.width='100%';
  td.appendChild(input);
  return {td, input};
}

async function upsert(payload){
  qs('dirStatus').textContent='Сохраняю...';
  const r = await fetch('/api/directory/upsert', {
    method:'POST',
    headers:{'Content-Type':'application/json'},
    body: JSON.stringify(payload)
  });
  if(!r.ok){
    const t = await r.text();
    qs('dirStatus').textContent = 'Ошибка: ' + t;
    return;
  }
  await load();
}

qs('addRowBtn').addEventListener('click', ()=>{
  const tbody = qs('dirTable').querySelector('tbody');
  if(!tbody) return;
  tbody.prepend(row({vehicle_plate:'', schedule_text:'', regime_start:'', regime_end:''}));
});

load();
