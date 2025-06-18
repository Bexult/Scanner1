<KOKZHAL_LOGISTICS>
<html lang="ru">
<head>
  <meta charset="UTF-8" />
  <title>Сканер трек-кодов</title>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
  <style>
    body { font-family: sans-serif; max-width: 500px; margin: auto; padding: 20px; }
    h2 { text-align: center; }
    input, button { width: 100%; padding: 10px; font-size: 18px; margin: 10px 0; }
    .status { border: 2px solid #ccc; padding: 15px; border-radius: 8px; font-size: 18px; }
    .ok { border-color: green; color: green; }
    .error { border-color: red; color: red; }
    .dup { border-color: orange; color: orange; }
  </style>
</head>
<body>

<h2>📦 Онлайн Сканер</h2>

<input type="file" id="fileInput" accept=".xlsx" />
<div class="status" id="status">Загрузите Excel-файл</div>
<input id="scanner" type="text" placeholder="Сканируйте трек-код..." autofocus />
<button id="exportBtn">📥 Скачать результат</button>

<audio id="sound-ok" src="https://actions.google.com/sounds/v1/cartoon/pop.ogg"></audio>
<audio id="sound-error" src="https://actions.google.com/sounds/v1/alarms/beep_short.ogg"></audio>

<script>
let expected = [], scanned = [];
const fileInput = document.getElementById('fileInput');
const status = document.getElementById('status');
const scanner = document.getElementById('scanner');
const exportBtn = document.getElementById('exportBtn');
const soundOk = document.getElementById('sound-ok');
const soundError = document.getElementById('sound-error');

fileInput.onchange = (e) => {
  const file = e.target.files[0];
  const reader = new FileReader();
  reader.onload = (ev) => {
    const workbook = XLSX.read(new Uint8Array(ev.target.result), { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
    expected = data.map(r => ({
      code: (r["Трек-код"] || "").toString().trim(),
      client: r["Клиент"] || "—",
      address: r["Адрес"] || "—"
    }));
    status.textContent = `✅ Загружено ${expected.length} записей`;
    status.className = 'status ok';
    scanner.focus();
  };
  reader.readAsArrayBuffer(file);
};

scanner.addEventListener('input', () => {
  clearTimeout(window.inputTimer);
  window.inputTimer = setTimeout(() => {
    const code = scanner.value.trim();
    if (code) {
      process(code);
      scanner.value = '';
    }
  }, 200);
});

function process(code) {
  code = code.toLowerCase();
  const dup = scanned.find(x => x['Трек-код'].toLowerCase() === code);
  const found = expected.find(x => x.code.toLowerCase() === code);
  let cl = 'ok', msg = '';

  if (dup) {
    msg = '🔁 Дубликат: ' + code;
    cl = 'dup';
    soundError.play();
  } else if (!found) {
    msg = '❌ Не найден: ' + code;
    cl = 'error';
    soundError.play();
  } else {
    msg = '✅ Найден: ' + found.client + " — " + found.address;
    scanned.push({
      'Дата': new Date().toLocaleString(),
      'Трек-код': code,
      'Клиент': found.client,
      'Адрес': found.address,
      'Статус': 'Найден'
    });
    soundOk.play();
  }

  status.textContent = msg;
  status.className = 'status ' + cl;
}

exportBtn.onclick = () => {
  if (!scanned.length) return alert("Нет данных для выгрузки");
  const ws = XLSX.utils.json_to_sheet(scanned);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Результат');
  XLSX.writeFile(wb, 'результат_сканирования.xlsx');
};
</script>

</body>
</html>
