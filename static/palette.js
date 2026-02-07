function qs(id){ return document.getElementById(id); }

function normalizeHex(s){
  if(!s) return null;
  s = String(s).trim();
  if(!s) return null;
  if(!s.startsWith('#')) s = '#' + s;
  if(/^#[0-9a-fA-F]{3}$/.test(s)){
    // expand #abc -> #aabbcc
    s = '#' + s[1] + s[1] + s[2] + s[2] + s[3] + s[3];
  }
  if(/^#[0-9a-fA-F]{6}$/.test(s)) return s.toLowerCase();
  return null;
}

function applyPalette(p){
  const root = document.documentElement;
  if(p.accent) root.style.setProperty('--accent', p.accent);
  if(p.schedule_fill) root.style.setProperty('--fillSchedule', p.schedule_fill);
  if(p.plan_fill) root.style.setProperty('--fillPlan', p.plan_fill);
  if(p.fact_fill) root.style.setProperty('--fillFact', p.fact_fill);
  if(p.border) root.style.setProperty('--gridBorder', p.border);
}

async function load(){
  const r = await fetch('/api/palette');
  const p = await r.json();
  fillInputs(p);
  applyPalette(p);
}

function fillInputs(p){
  const map = {
    accent: p.accent,
    schedule: p.schedule_fill,
    plan: p.plan_fill,
    fact: p.fact_fill,
    border: p.border,
  };
  for(const k of Object.keys(map)){
    const v = map[k] || '#000000';
    qs(k).value = normalizeHex(v) || '#000000';
    qs(k + 'Text').value = normalizeHex(v) || v;
  }
}

function readInputs(){
  const get = (k)=> normalizeHex(qs(k + 'Text').value) || normalizeHex(qs(k).value);
  return {
    accent: get('accent'),
    schedule_fill: get('schedule'),
    plan_fill: get('plan'),
    fact_fill: get('fact'),
    border: get('border')
  };
}

function wireSync(k){
  const c = qs(k);
  const t = qs(k + 'Text');
  c.addEventListener('input', ()=>{
    t.value = c.value;
    applyPalette(readInputs());
  });
  t.addEventListener('input', ()=>{
    const v = normalizeHex(t.value);
    if(v) c.value = v;
    applyPalette(readInputs());
  });
}

async function save(){
  const p = readInputs();
  qs('pStatus').textContent = 'Сохраняю...';
  const r = await fetch('/api/palette', {
    method: 'POST',
    headers: {'Content-Type':'application/json'},
    body: JSON.stringify(p)
  });
  if(r.ok){
    qs('pStatus').textContent = 'Сохранено';
  } else {
    qs('pStatus').textContent = 'Ошибка сохранения';
  }
}

async function reset(){
  qs('pStatus').textContent = 'Сбрасываю...';
  const r = await fetch('/api/palette/reset', {method:'POST'});
  if(r.ok){
    const p = await r.json();
    fillInputs(p);
    applyPalette(p);
    qs('pStatus').textContent = 'Сброшено';
  } else {
    qs('pStatus').textContent = 'Ошибка сброса';
  }
}

['accent','schedule','plan','fact','border'].forEach(wireSync);
qs('pSave').addEventListener('click', save);
qs('pReset').addEventListener('click', reset);
load();
