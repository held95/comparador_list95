const inputExcel = document.getElementById('inputExcel');
const btnProcessar = document.getElementById('btnProcessar');
const downloadLink = document.getElementById('downloadLink');
const LIMIAR = 0.9;

let workbook;

inputExcel.addEventListener('change', e => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = evt => {
    const data = new Uint8Array(evt.target.result);
    workbook = XLSX.read(data, { type: 'array' });
    btnProcessar.disabled = false;
  };
  reader.readAsArrayBuffer(file);
});

btnProcessar.addEventListener('click', () => {
  if (!workbook) return;

  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const json = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false });
  const colA = json.map(r => (r[0]||'').toString().trim()).filter(v => v);
  const colB = json.map(r => (r[1]||'').toString().trim()).filter(v => v);

  function normalizar(nome) {
    let s = nome.toLowerCase();
    s = unidecode(s);
    s = s.replace(/[:]/g, ' ');
    s = s.replace(/[^a-z\s]/g, '');
    s = s.replace(/\s+/g, ' ');
    s = s.replace(/^(rn|rn i+|desconhecido|de|do|da|dos|das)\s+/g, '');
    return s.trim();
  }

  const normA = colA.map(normalizar);
  const normB = colB.map(normalizar);

  const results = [];

  for (let i=0; i<colA.length; i++) {
    const na = colA[i], nna = normA[i];
    let bestB = '', bestScore = 0;

    for (let j=0; j<colB.length; j++) {
      const nb = colB[j], nnb = normB[j];
      const score = fuseScore(nna, nnb);
      if (score > bestScore || (score === bestScore && /rn|desconhecido/.test(nb.toLowerCase()))) {
        bestScore = score;
        bestB = nb;
      }
    }
    results.push([na, '=', bestScore >= LIMIAR ? bestB : '']);
  }

  const usadosB = new Set(results.filter(r=>r[2]).map(r=>r[2]));
  colB.forEach(nb => {
    if (!usadosB.has(nb)) results.push(['', '=', nb]);
  });

  const ws = XLSX.utils.aoa_to_sheet([["Coluna A","=","Coluna B"], ...results]);
  const wbout = XLSX.write({ Sheets: { 'Resultado': ws }, SheetNames:['Resultado'] }, { bookType:'xlsx', type:'array' });
  const blob = new Blob([wbout], { type:'application/octet-stream' });
  const url = URL.createObjectURL(blob);
  downloadLink.href = url;
  downloadLink.download = 'Resultado_Cross_X_Tasy.xlsx';
  downloadLink.textContent = '⬇️ Baixar resultado';
  downloadLink.style.display = 'block';
});

// Score usando Fuse.js tokenSortRatio aproximado
function fuseScore(a, b) {
  const fuse = new Fuse([b], { includeScore: true, useExtendedSearch: false, isCaseSensitive:false, tokenize: true });
  const res = fuse.search(a);
  return res.length ? (1 - res[0].score) : 0;
}
