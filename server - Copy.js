const express = require('express');
const xlsx = require('xlsx');
const app = express();
const port = 3000;

app.set('view engine', 'ejs');
app.set('views', __dirname + '/views');

let erogazioni = [];
let tipologia = [];
let erogazioniTipologia = [];

// Utility base
function readXLSX(filename) {
  const workbook = xlsx.readFile(filename);
  const sheetName = workbook.SheetNames[0];
  return xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
}
function formattaData(data) {
  if (!data) return "";
  if (typeof data === "number") {
    const epoch = new Date(1899, 11, 30);
    const date = new Date(epoch.getTime() + data * 86400000);
    return `${String(date.getDate()).padStart(2, '0')}-${String(date.getMonth() + 1).padStart(2, '0')}-${date.getFullYear()}`;
  }
  if (data instanceof Date) {
    return `${String(data.getDate()).padStart(2, '0')}-${String(data.getMonth() + 1).padStart(2, '0')}-${data.getFullYear()}`;
  }
  data = String(data);
  let parts = data.split(/[-\/]/);
  if (parts.length === 3) {
    let [a, b, c] = parts;
    if (a.length === 4) return `${c.padStart(2, '0')}-${b.padStart(2, '0')}-${a}`;
    if (c.length === 4) return `${a.padStart(2, '0')}-${b.padStart(2, '0')}-${c}`;
  }
  return data;
}
function estraiAnno(data) {
  if (!data) return null;
  let p = String(data).split(/[-\/]/);
  if (p.length === 3) {
    if (p[2].length === 4) return parseInt(p[2]);
    if (p[0].length === 4) return parseInt(p[0]);
  }
  return null;
}
function anniDaArray(arr, campo) {
  return Array.from(new Set(arr.map(r => estraiAnno(r[campo])).filter(x => !isNaN(x) && x != null))).sort((a, b) => a - b);
}
function soggettiDaArray(arr) {
  return Array.from(new Set(arr.map(r => r.Soggetto).filter(x => typeof x === 'string' && x.length))).sort();
}
function uniqueColValues(arr, col) {
  return Array.from(new Set(arr.map(r => r[col]).filter(v => typeof v === 'string' && v.trim().length))).sort();
}
function arrSelected(param) {
  if (!param) return [];
  if (Array.isArray(param)) return param;
  return [param];
}
function parseDataForSort(str) {
  if (!str) return 0;
  if (!isNaN(str)) {
    const epoch = new Date(1899, 11, 30);
    return new Date(epoch.getTime() + Number(str) * 86400000).getTime();
  }
  let p = String(str).split(/[-\/]/);
  if (p.length === 3) {
    if (p[0].length === 4) return new Date(p[0], p[1] - 1, p[2]).getTime();
    if (p[2].length === 4) return new Date(p[2], p[1] - 1, p[0]).getTime();
  }
  return Date.parse(str);
}
function calcSommaImporti(ris) {
  return ris.reduce((acc, r) => {
    let v = r.Importo || r.importo || r.IMPORTO || 0;
    if (typeof v === 'string') v = v.replace(',', '.');
    const n = parseFloat(v);
    return acc + (isNaN(n) ? 0 : n);
  }, 0);
}
// Per SI/NO
function prettifyBool(val) {
  if (val === true || val === "true" || val === 1 || val === "1") return "SI";
  if (val === false || val === "false" || val === 0 || val === "0") return "NO";
  if (val === '' || val === null || val === undefined) return "";
  return val;
}

function initData() {
  erogazioni = readXLSX('SCN_Erogazioni.xlsx');
  tipologia = readXLSX('TipologiaSoggetti.xlsx');
  erogazioniTipologia = readXLSX('SCN_Erogazioni_Tipologia.xlsx');
}

app.get('/', (req, res) => { res.send('Server TaoMau attivo'); });

// --- ERGAZIONI PRINCIPALE ---
app.get('/tabella/erogazioni', (req, res) => {
  let risultati = erogazioni.map(riga => ({
    ...riga,
    DataErogazione: formattaData(riga.DataErogazione),
    DataTrasmissione: formattaData(riga.DataTrasmissione)
  }));
  let anniLista = anniDaArray(risultati, "DataErogazione");
  let annoStartDefault = anniLista.length ? anniLista[0] : "";
  let annoEndDefault = anniLista.length ? anniLista[anniLista.length - 1] : "";
  let soggettiLista = soggettiDaArray(risultati);

  const annoStart = req.query.annoStart ? parseInt(req.query.annoStart) : annoStartDefault;
  const annoEnd   = req.query.annoEnd ? parseInt(req.query.annoEnd) : annoEndDefault;

  let risultatoFiltrato = risultati.filter(riga => {
    let anno = estraiAnno(riga.DataErogazione);
    if (annoStart && annoEnd && anno) return anno >= annoStart && anno <= annoEnd;
    return true;
  });
  if (req.query.soggetto && req.query.soggetto !== 'Tutti') {
    risultatoFiltrato = risultatoFiltrato.filter(r => r.Soggetto === req.query.soggetto);
  }
  if (req.query.soggetto_parziale && req.query.soggetto_parziale.trim() !== "") {
    let part = req.query.soggetto_parziale.trim().toLowerCase();
    risultatoFiltrato = risultatoFiltrato.filter(r => r.Soggetto && r.Soggetto.toLowerCase().includes(part));
  }
  const sortField = req.query.sortField || null;
  const sortDir = req.query.sortDir || 'asc';
  if (sortField) {
    risultatoFiltrato.sort((a, b) => {
      let va = a[sortField];
      let vb = b[sortField];
      let cmp = 0;
      if (['DataErogazione', 'DataTrasmissione', 'DataInizio', 'DataFine'].includes(sortField)) {
        cmp = parseDataForSort(va) - parseDataForSort(vb);
      } else {
        let asnum = Number(va), bsnum = Number(vb);
        cmp = (!isNaN(asnum) && !isNaN(bsnum)) ? (asnum - bsnum) : String(va).localeCompare(String(vb), 'it', { numeric: true });
      }
      return sortDir === 'asc' ? cmp : -cmp;
    });
  }
  const sommaImporti = calcSommaImporti(risultatoFiltrato);
  res.render('tabella', {
    dati: risultatoFiltrato, titolo: 'SCN_Erogazioni', query: req.query,
    anni: anniLista, annoStartDefault, annoEndDefault, soggetti: soggettiLista,
    sommaImporti, sortField, sortDir, filterOptions: {}
  });
});

// --- TIPOLOGIA ---
app.get('/tabella/tipologia', (req, res) => {
  res.render('tabella', {
    dati: tipologia, titolo: 'TipologiaSoggetti', query: req.query,
    anni: [], annoStartDefault: null, annoEndDefault: null,
    soggetti: [], sommaImporti: null, sortField: null, sortDir: null, filterOptions: {}
  });
});

// --- EROGAZIONI_TIPOLOGIA CON CHECKBOX MULTIPLI e SI/NO ---
app.get('/tabella/erogazioni_tipologia', (req, res) => {
  let risultati = erogazioniTipologia.map(riga => ({
    Soggetto: riga.Soggetto,
    Importo: riga.Importo || riga.importo || riga.IMPORTO,
    DataErogazione: formattaData(riga.DataErogazione),
    AnnoErogazione: estraiAnno(formattaData(riga.DataErogazione)),
    Azienda: prettifyBool(riga.Azienda),
    Politico: prettifyBool(riga.Politico),
    Taormina: prettifyBool(riga.Taormina),
    Camera_Senato: prettifyBool(riga.Camera_Senato),
    ARS: prettifyBool(riga.ARS),
    FENAPI: prettifyBool(riga.FENAPI)
  }));

  let anniLista = anniDaArray(risultati, "DataErogazione");
  let annoStartDefault = anniLista.length ? anniLista[0] : "";
  let annoEndDefault = anniLista.length ? anniLista[anniLista.length - 1] : "";
  let soggettiLista = soggettiDaArray(risultati);

  let filterOptions = {
    Azienda: uniqueColValues(risultati, "Azienda"),
    Politico: uniqueColValues(risultati, "Politico"),
    Taormina: uniqueColValues(risultati, "Taormina"),
    Camera_Senato: uniqueColValues(risultati, "Camera_Senato"),
    ARS: uniqueColValues(risultati, "ARS"),
    FENAPI: uniqueColValues(risultati, "FENAPI")
  };

  const annoStart = req.query.annoStart ? parseInt(req.query.annoStart) : annoStartDefault;
  const annoEnd = req.query.annoEnd ? parseInt(req.query.annoEnd) : annoEndDefault;
  let risultatiFiltrati = risultati.filter(riga => {
    let anno = riga.AnnoErogazione;
    let checkAnno = (!annoStart || !annoEnd || !anno) ? true : (anno >= annoStart && anno <= annoEnd);
    let checkSogg = req.query.soggetto && req.query.soggetto !== 'Tutti' ? riga.Soggetto === req.query.soggetto : true;
    let checkAzienda = req.query.Azienda ? arrSelected(req.query.Azienda).includes(riga.Azienda) : true;
    let checkPolitico = req.query.Politico ? arrSelected(req.query.Politico).includes(riga.Politico) : true;
    let checkTaormina = req.query.Taormina ? arrSelected(req.query.Taormina).includes(riga.Taormina) : true;
    let checkCameraSenato = req.query.Camera_Senato ? arrSelected(req.query.Camera_Senato).includes(riga.Camera_Senato) : true;
    let checkARS = req.query.ARS ? arrSelected(req.query.ARS).includes(riga.ARS) : true;
    let checkFENAPI = req.query.FENAPI ? arrSelected(req.query.FENAPI).includes(riga.FENAPI) : true;
    return checkAnno && checkSogg && checkAzienda && checkPolitico && checkTaormina && checkCameraSenato && checkARS && checkFENAPI;
  });
  if (req.query.soggetto_parziale && req.query.soggetto_parziale.trim() !== "") {
    let part = req.query.soggetto_parziale.trim().toLowerCase();
    risultatiFiltrati = risultatiFiltrati.filter(r => r.Soggetto && r.Soggetto.toLowerCase().includes(part));
  }
  const sortField = req.query.sortField || null;
  const sortDir = req.query.sortDir || 'asc';
  if (sortField) {
    risultatiFiltrati.sort((a, b) => {
      let va = a[sortField];
      let vb = b[sortField];
      let cmp = 0;
      if (sortField === 'DataErogazione') cmp = parseDataForSort(va) - parseDataForSort(vb);
      else if (sortField === 'Importo') {
        let asnum = Number(va), bsnum = Number(vb);
        cmp = (!isNaN(asnum) && !isNaN(bsnum)) ? (asnum - bsnum) : String(va).localeCompare(String(vb), 'it', { numeric: true });
      } else if (sortField === 'AnnoErogazione') cmp = Number(va) - Number(vb);
      else cmp = String(va).localeCompare(String(vb), 'it', { numeric: true });
      return sortDir === 'asc' ? cmp : -cmp;
    });
  }
  const sommaImporti = risultatiFiltrati.reduce((acc, r) => {
    let valore = r.Importo || 0;
    if (typeof valore === 'string') valore = valore.replace(',', '.');
    const num = parseFloat(valore);
    return acc + (isNaN(num) ? 0 : num);
  }, 0);

  res.render('tabella', {
    dati: risultatiFiltrati,
    titolo: 'SCN_Erogazioni_Tipologia',
    query: req.query,
    anni: anniLista,
    annoStartDefault,
    annoEndDefault,
    soggetti: soggettiLista,
    sommaImporti,
    sortField,
    sortDir,
    filterOptions
  });
});

app.listen(port, () => {
  initData();
  console.log(`Server attivo su http://localhost:${port}`);
});
