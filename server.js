const express = require('express');
const xlsx = require('xlsx');
const app = express();
const port = 3000;

app.set('view engine', 'ejs');
app.set('views', __dirname + '/views');

let erogazioniSoggettiGruppi = [];

function readXLSX(filename) {
  const workbook = xlsx.readFile(filename);
  const sheetName = workbook.SheetNames[0];
  return xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
}
function mapGroupBool(val) {
  if (val == null) return null;
  return String(val).trim().toUpperCase();
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
  return Array.from(new Set(arr.map(r => mapGroupBool(r[col])).filter(v => typeof v === 'string' && v.trim().length)));
}
function arrSelected(param) {
  if (!param) return [];
  if (Array.isArray(param)) return param;
  return [param];
}
function calcSommaImporti(arr) {
  return arr.reduce((acc, r) => {
    let v = r.Importo;
    if (typeof v === 'string') v = v.replace(',', '.');
    let num = parseFloat(v);
    return acc + (isNaN(num) ? 0 : num);
  }, 0);
}
function parseDataForSort(str) {
  if (!str) return 0;
  let p = String(str).split(/[-\/]/);
  if (p.length === 3) {
    if (p[0].length === 4) return new Date(p[0], p[1] - 1, p[2]).getTime();
    if (p[2].length === 4) return new Date(p[2], p[1] - 1, p[0]).getTime();
  }
  return Date.parse(str);
}

function initData() {
  erogazioniSoggettiGruppi = readXLSX('SV_SCN_EROGAZIONI_SOGGETTI_GRUPPI_Importi.xlsx');
}

app.get('/', (req, res) => { res.send('Server SCN attivo'); });

app.get('/tabella/erogazioni_soggetti_gruppi', (req, res) => {
  let gruppi = [
    "G_Aziende", "G_FENAPI", "G_Politici", "G_Taormina",
    "G_Messina", "G_REG", "G_NAZ", "G_EspTitGrat"
  ];

  let dati = erogazioniSoggettiGruppi.map(riga => ({
    Anno: riga.Anno,
    Soggetto: riga.Soggetto,
    Importo: riga.Importo,
    Partito: riga.Partito,
    DataErogazione: formattaData(riga.DataErogazione),
    DataTrasmissione: formattaData(riga.DataTrasmissione),
    G_Aziende: mapGroupBool(riga.G_Aziende),
    G_FENAPI: mapGroupBool(riga.G_FENAPI),
    G_Politici: mapGroupBool(riga.G_Politici),
    G_Taormina: mapGroupBool(riga.G_Taormina),
    G_Messina: mapGroupBool(riga.G_Messina),
    G_REG: mapGroupBool(riga.G_REG),
    G_NAZ: mapGroupBool(riga.G_NAZ),
    G_EspTitGrat: mapGroupBool(riga.G_EspTitGrat)
  }));

  let anniLista = anniDaArray(dati, "DataErogazione");
  let annoStartDefault = anniLista.length ? anniLista[0] : "";
  let annoEndDefault = anniLista.length ? anniLista[anniLista.length - 1] : "";
  let soggettiLista = soggettiDaArray(dati);
  let partitiLista = uniqueColValues(dati, "Partito");

  let gruppoOptions = {};
  gruppi.forEach(g => {
    gruppoOptions[g] = uniqueColValues(dati, g);
  });

  const annoStart = req.query.annoStart ? parseInt(req.query.annoStart) : annoStartDefault;
  const annoEnd = req.query.annoEnd ? parseInt(req.query.annoEnd) : annoEndDefault;

  let datiFiltrati = dati.filter(riga => {
    let anno = estraiAnno(riga.DataErogazione);
    let checkAnno = (!annoStart || !annoEnd || !anno) ? true : (anno >= annoStart && anno <= annoEnd);
    let checkSogg = req.query.soggetto && req.query.soggetto !== 'Tutti' ? riga.Soggetto === req.query.soggetto : true;
    let checkPartito = req.query.partito && arrSelected(req.query.partito).length > 0
      ? arrSelected(req.query.partito).includes(riga.Partito)
      : true;
    let checkGruppi = gruppi.every(gr => {
      if (req.query[gr] && arrSelected(req.query[gr]).length > 0) {
        return arrSelected(req.query[gr]).includes(riga[gr]);
      }
      return true;
    });
    return checkAnno && checkSogg && checkPartito && checkGruppi;
  });

  if (req.query.soggetto_parziale && req.query.soggetto_parziale.trim() !== "") {
    let part = req.query.soggetto_parziale.trim().toLowerCase();
    datiFiltrati = datiFiltrati.filter(r => r.Soggetto && r.Soggetto.toLowerCase().includes(part));
  }

  const sortField = req.query.sortField || null;
  const sortDir = req.query.sortDir || 'asc';
  if (sortField) {
    datiFiltrati.sort((a, b) => {
      let va = a[sortField], vb = b[sortField], cmp = 0;
      if (sortField === 'DataErogazione' || sortField === 'DataTrasmissione') {
        cmp = parseDataForSort(va) - parseDataForSort(vb);
      } else if (sortField === 'Importo') {
        let asnum = Number(va), bsnum = Number(vb);
        cmp = (!isNaN(asnum) && !isNaN(bsnum)) ? (asnum - bsnum) : String(va).localeCompare(String(vb), 'it', { numeric: true });
      } else {
        cmp = String(va).localeCompare(String(vb), 'it', { numeric: true });
      }
      return sortDir === 'asc' ? cmp : -cmp;
    });
  }

  const sommaImporti = calcSommaImporti(datiFiltrati);

  res.render('tabella', {
    dati: datiFiltrati,
    titolo: 'SV_SCN_EROGAZIONI_SOGGETTI_GRUPPI_Importi',
    query: req.query,
    anni: anniLista,
    annoStartDefault,
    annoEndDefault,
    soggetti: soggettiLista,
    partiti: partitiLista,
    gruppi: gruppi,
    gruppoOptions: gruppoOptions,
    sommaImporti,
    sortField,
    sortDir,
    filterOptions: {}
  });
});

app.listen(port, () => {
  initData();
  console.log(`Server attivo su http://localhost:${port}`);
});
