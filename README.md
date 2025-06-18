<KOKZHAL_logistics>
<html lang="ru">
<head><meta charset="UTF-8"/><title>Сканер онлайн</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
 body { font-family: sans-serif; padding:20px; max-width:400px; margin:auto; }
 input,button { font-size:18px; padding:8px; width:100%; margin:10px 0; }
 .status { padding:10px; border:2px solid #ccc; border-radius:8px; margin:10px 0; }
 .ok { border-color:green; color:green; }
 .error { border-color:red; color:red; }
 .dup { border-color:orange; color:orange; }
</style>
</head>
<body>
<h2>📦 Онлайн Сканер</h2>
<input type="file" id="fileInput" accept=".xlsx"/>
<div class="status" id="status">Загрузите Excel</div>
<input id="scanner" type="text" placeholder="Сканируйте трек-код..." autofocus/>
<button id="exportBtn">📥 Скачать результат</button>
<audio id="sound-ok" src="https://actions.google.com/sounds/v1/cartoon/pop.ogg"></audio>
<audio id="sound-error" src="https://actions.google.com/sounds/v1/alarms/beep_short.ogg"></audio>
<script>
let expected=[], scanned=[], buf='', timer;
const fileInput = document.getElementById('fileInput');
const status = document.getElementById('status');
const scanner = document.getElementById('scanner');
const exportBtn = document.getElementById('exportBtn');
const soundOk = document.getElementById('sound-ok');
const soundError = document.getElementById('sound-error');

fileInput.onchange = e => {
  const f = e.target.files[0];
  const reader = new FileReader();
  reader.onload = ev => {
    const wb = XLSX.read(new Uint8Array(ev.target.result), {type:'array'});
    const sh = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sh, {defval: "", raw: false});
expected = data.map(r => ({
  code: (r["Трек-код"] || "").toString().trim(),
  client: r["Клиент"] || "—",
  address: r["Адрес"] || "—"
}));

    expected = data.map(r=>({
      code:r[0]?.toString().trim(),
      client:r[1]?.toString().trim()||'—',
      address:r[2]?.toString().trim()||'—'
    }));
    status.textContent = `✅ Загрузка успешна (${expected.length} записей)`;
    status.className='status ok';
    scanner.focus();
  };
  reader.readAsArrayBuffer(f);
};

document.addEventListener('keydown', e=>{
  if (timer) clearTimeout(timer);
  if (e.key==='Shift'||e.key==='Tab') return;
  if (e.key==='Enter') {
    process(buf.trim());
    buf = '';
    return;
  }
  buf += e.key;
  timer = setTimeout(()=>{ process(buf.trim()); buf=''; }, 150);
});

function process(code){
  if (!code) return;
 code = code.trim().toLowerCase();
const dup = scanned.find(x=>x['Трек-код'].toLowerCase() === code);
const found = expected.find(x=>x.code.toLowerCase() === code);
  let state, cl='ok';
  if (dup) { state='🔁 Дубликат'; cl='dup'; soundError.play(); }
  else if (!found) { state='❌ Не найден'; cl='error'; soundError.play(); }
  else {
    state='✅ Найден'; soundOk.play();
    scanned.push({
      'Дата': new Date().toLocaleString(),
      'Трек-код':code,
      'Клиент':found.client,
      'Адрес':found.address,
      'Статус':state
    });
  }
  status.textContent = `${state} — ${code}`;
  status.className = `status ${cl}`;
}

exportBtn.onclick = () => {
  if (!scanned.length) return alert('Нет данных для экспорта');
  const ws = XLSX.utils.json_to_sheet(scanned);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Сканирование');
  XLSX.writeFile(wb, 'scan_result.xlsx');
};
</script>
</body>
</html>
