// Seconda cassa: tab "MESE" = Cassa 1, "MESE C2" = Cassa 2, "MESE TOT" = somma (solo formule).
const test   = require('node:test');
const assert = require('node:assert');
const fs     = require('node:fs');
const path   = require('node:path');

const code = fs.readFileSync(path.join(__dirname, '..', 'Codice.gs'), 'utf8');

// Foglio finto: righe 1-3 intestazioni, dalla riga 4 un giorno per riga (13 colonne).
function fakeSheet(dayLabels) {
  const cells = [['NOVEMBRE 2026'], [''], ['Data']].concat(dayLabels.map(l => [l]));
  cells.forEach(r => { while (r.length < 13) r.push(''); });
  return {
    cells,
    getLastRow: () => cells.length,
    getRange: (row, col, nRows = 1, nCols = 1) => ({
      getValue:    () => cells[row - 1][col - 1],
      getValues:   () => cells.slice(row - 1, row - 1 + nRows).map(r => r.slice(col - 1, col - 1 + nCols)),
      setValue:    v  => { cells[row - 1][col - 1] = v; },
      setFormula:  f  => { cells[row - 1][col - 1] = f; },
      setValues:   vs => vs.forEach((r, i) => r.forEach((v, j) => { cells[row - 1 + i][col - 1 + j] = v; })),
    }),
  };
}

// Cartella finta con più tab per nome.
function load(tabs) {
  const ss = { getSheetByName: name => tabs[name] || null, getUrl: () => 'https://example/foglio' };
  const SpreadsheetApp = { openById: () => ss };
  return new Function('SpreadsheetApp', code + '\nreturn { writeToSheet, writeFattura };')(SpreadsheetApp);
}

const SCONTRINO = { asl: 0, farmaco: 2000, parafarmaco: 500, varie: 0, servizi0: 0,
                    serviziIva: 6, shopper: 0, dpi: 0, fatture: 0, scontrini: 0 };

test('scontrino della Cassa 2 va nel tab "MESE C2" e non tocca la Cassa 1', () => {
  const c1 = fakeSheet(['ven 13/11']);
  const c2 = fakeSheet(['ven 13/11']);
  const app = load({ 'NOVEMBRE': c1, 'NOVEMBRE C2': c2 });

  const res = app.writeToSheet(SCONTRINO, '2026-11-13', '', 2);

  assert.strictEqual(c2.cells[3][2], 2000);
  assert.strictEqual(c1.cells[3][2], '');
  assert.strictEqual(res.tab, 'NOVEMBRE C2');
});

test('rileggere lo scontrino della Cassa 2 sovrascrive solo la Cassa 2', () => {
  const c1 = fakeSheet(['ven 13/11']);
  const c2 = fakeSheet(['ven 13/11']);
  const app = load({ 'NOVEMBRE': c1, 'NOVEMBRE C2': c2 });

  app.writeToSheet({ ...SCONTRINO, farmaco: 1000 }, '2026-11-13', '', 1);
  app.writeToSheet({ ...SCONTRINO, farmaco: 300 },  '2026-11-13', '', 2);
  app.writeToSheet({ ...SCONTRINO, farmaco: 350 },  '2026-11-13', '', 2);

  assert.strictEqual(c1.cells[3][2], 1000);
  assert.strictEqual(c2.cells[3][2], 350);
});

test('senza indicare la cassa si scrive nella Cassa 1, come prima', () => {
  const c1 = fakeSheet(['ven 13/11']);
  const c2 = fakeSheet(['ven 13/11']);
  const app = load({ 'NOVEMBRE': c1, 'NOVEMBRE C2': c2 });

  const res = app.writeToSheet(SCONTRINO, '2026-11-13', '');

  assert.strictEqual(c1.cells[3][2], 2000);
  assert.strictEqual(c2.cells[3][2], '');
  assert.strictEqual(res.tab, 'NOVEMBRE');
});

test('Cassa 2 richiesta ma tab C2 assente: errore, nessuna scrittura nella Cassa 1', () => {
  const c1 = fakeSheet(['ven 13/11']);
  const app = load({ 'NOVEMBRE': c1 });

  assert.throws(() => app.writeToSheet(SCONTRINO, '2026-11-13', '', 2), /NOVEMBRE C2/);
  assert.strictEqual(c1.cells[3][2], '');
});

test('doPost passa la cassa a writeToSheet', () => {
  const c1 = fakeSheet(['ven 13/11']);
  const c2 = fakeSheet(['ven 13/11']);
  const ss = { getSheetByName: name => ({ 'NOVEMBRE': c1, 'NOVEMBRE C2': c2 })[name] || null };
  const SpreadsheetApp  = { openById: () => ss };
  const ContentService  = { createTextOutput: s => ({ setMimeType: () => s }), MimeType: { JSON: 'json' } };
  const doPost = new Function('SpreadsheetApp', 'ContentService', code + '\nreturn doPost;')(SpreadsheetApp, ContentService);

  const out = JSON.parse(doPost({ postData: { contents: JSON.stringify({ fields: SCONTRINO, date: '2026-11-13', note: '', cassa: 2 }) } }));

  assert.strictEqual(out.tab, 'NOVEMBRE C2');
  assert.strictEqual(c2.cells[3][2], 2000);
});

// --- Report mensile ---------------------------------------------------------

function loadReport(tabs) {
  const sent = [];
  const ss = { getSheetByName: name => tabs[name] || null, getUrl: () => 'https://example/foglio' };
  const SpreadsheetApp = { openById: () => ss };
  const MailApp   = { sendEmail: m => sent.push(m) };
  const Utilities = { newBlob: (data, type, name) => ({ data, type, name }) };
  const fn = new Function('SpreadsheetApp', 'MailApp', 'Utilities', code + '\nreturn sendMonthReport;')(SpreadsheetApp, MailApp, Utilities);
  return { sendMonthReport: fn, sent };
}

test('report: se esiste il tab TOT, la commercialista riceve i totali delle due casse', () => {
  const c1  = fakeSheet(['ven 13/11']);
  const tot = fakeSheet(['ven 13/11']);
  c1.cells[3][2]  = 1000;
  tot.cells[3][2] = 1300;
  const { sendMonthReport, sent } = loadReport({ 'NOVEMBRE': c1, 'NOVEMBRE TOT': tot });

  sendMonthReport(2026, 11);

  assert.match(sent[0].attachments[0].data, /ven 13\/11;;1300;/);
  assert.doesNotMatch(sent[0].attachments[0].data, /;1000;/);
});

test('report: senza tab TOT (una sola cassa) si usa il tab del mese, come prima', () => {
  const c1 = fakeSheet(['ven 13/11']);
  c1.cells[3][2] = 1000;
  const { sendMonthReport, sent } = loadReport({ 'NOVEMBRE': c1 });

  sendMonthReport(2026, 11);

  assert.match(sent[0].attachments[0].data, /ven 13\/11;;1000;/);
});

// --- preparaCassa2: crea "MESE C2" e "MESE TOT" -------------------------------

// Tab finto di un mese "come quello vero": titolo, "Data", un giorno per riga
// (importi + formula totale in L), poi TOTALI e PERCENTUALI.
function monthSheet(name, year, m) {
  const DN = ['dom','lun','mar','mer','gio','ven','sab'];
  const days = new Date(year, m + 1, 0).getDate();
  const cells = [[name + ' ' + year], [''], ['Data']];
  for (let d = 1; d <= days; d++) {
    const dt = new Date(year, m, d, 12);
    const r  = [DN[dt.getDay()] + ' ' + String(d).padStart(2, '0') + '/' + String(m + 1).padStart(2, '0')];
    r[2] = 100 + d; r[11] = `=SUM(B${d + 3}:J${d + 3})-K${d + 3}`; r[12] = 'nota ' + d;
    cells.push(r);
  }
  cells.push(['TOTALI', '=SUM(B4:B' + (days + 3) + ')'], ['PERCENTUALI']);
  cells.forEach(r => { for (let i = 0; i < 13; i++) if (r[i] === undefined) r[i] = ''; });
  return sheetFrom(name, cells);
}

function sheetFrom(name, cells) {
  const sh = {
    name, cells,
    getName: () => sh.name,
    setName: n => {
      if (sh._ss && sh._ss._failOn === n) throw new Error('errore simulato su ' + n);
      sh.name = n; return sh;
    },
    getLastRow: () => cells.length,
    copyTo: ss => ss._add(sheetFrom('Copia di ' + sh.name, cells.map(r => r.slice()))),
    getRange: (row, col, nRows = 1, nCols = 1) => {
      const each = f => { for (let i = 0; i < nRows; i++) for (let j = 0; j < nCols; j++) f(row - 1 + i, col - 1 + j, i, j); };
      return {
        getValues:   () => cells.slice(row - 1, row - 1 + nRows).map(r => r.slice(col - 1, col - 1 + nCols)),
        clearContent: () => each((r, c) => { cells[r][c] = ''; }),
        setFormulas: fs => each((r, c, i, j) => { cells[r][c] = fs[i][j]; }),
      };
    },
  };
  return sh;
}

function loadPrepara(year, months) {
  const sheets = [];
  const ss = {
    getSheetByName: n => sheets.find(s => s.name === n) || null,
    _add: s => { s._ss = ss; sheets.push(s); return s; },
    deleteSheet: s => { const i = sheets.indexOf(s); if (i >= 0) sheets.splice(i, 1); },
  };
  months.forEach(m => ss._add(monthSheet(['GENNAIO','FEBBRAIO','MARZO','APRILE','MAGGIO','GIUGNO','LUGLIO',
    'AGOSTO','SETTEMBRE','OTTOBRE','NOVEMBRE','DICEMBRE'][m], year, m)));
  const SpreadsheetApp = { openById: () => ss };
  const Logger = { log() {} };
  const fn = new Function('SpreadsheetApp', 'Logger', code + '\nreturn preparaCassa2;')(SpreadsheetApp, Logger);
  return { preparaCassa2: fn, ss, sheets };
}

const ALL = [0,1,2,3,4,5,6,7,8,9,10,11];

test('preparaCassa2 crea C2 vuoto e TOT con le somme, dal mese indicato in poi', () => {
  const { preparaCassa2, ss } = loadPrepara(2026, ALL);

  preparaCassa2(2026, 11);

  assert.strictEqual(ss.getSheetByName('OTTOBRE C2'), null);
  assert.strictEqual(ss.getSheetByName('OTTOBRE TOT'), null);
  const c2  = ss.getSheetByName('NOVEMBRE C2');
  const tot = ss.getSheetByName('NOVEMBRE TOT');
  const dic = ss.getSheetByName('DICEMBRE TOT');
  assert.ok(c2 && tot && dic && ss.getSheetByName('DICEMBRE C2'));

  // C2: stesse date, importi e note vuoti, formula del totale e righe TOTALI intatte
  const r13 = 3 + 13;                       // ven 13/11 è alla riga 16 (indice 15)
  assert.strictEqual(c2.cells[r13 - 1][0], 'ven 13/11');
  assert.deepStrictEqual(c2.cells[r13 - 1].slice(1, 11), ['', '', '', '', '', '', '', '', '', '']);
  assert.strictEqual(c2.cells[r13 - 1][12], '');
  assert.strictEqual(c2.cells[r13 - 1][11], '=SUM(B16:J16)-K16');
  assert.strictEqual(c2.cells[33][0], 'TOTALI');

  // TOT: somma cella per cella delle due casse; vuoto se entrambe vuote (giorno a cassa singola ok)
  assert.strictEqual(tot.cells[r13 - 1][0], 'ven 13/11');
  assert.strictEqual(tot.cells[r13 - 1][2],
    "=IF(COUNT('NOVEMBRE'!C16,'NOVEMBRE C2'!C16),ROUND(SUM('NOVEMBRE'!C16,'NOVEMBRE C2'!C16),2),\"\")");
  assert.strictEqual(tot.cells[r13 - 1][10],
    "=IF(COUNT('NOVEMBRE'!K16,'NOVEMBRE C2'!K16),ROUND(SUM('NOVEMBRE'!K16,'NOVEMBRE C2'!K16),2),\"\")");
  assert.strictEqual(tot.cells[r13 - 1][11], '=SUM(B16:J16)-K16');
  assert.strictEqual(tot.cells[r13 - 1][12], "=TEXTJOIN(\" · \",TRUE,'NOVEMBRE'!M16,'NOVEMBRE C2'!M16)");
  assert.strictEqual(tot.cells[33][0], 'TOTALI');

  // la Cassa 1 non viene toccata
  const c1 = ss.getSheetByName('NOVEMBRE');
  assert.strictEqual(c1.cells[r13 - 1][2], 113);
  assert.strictEqual(c1.cells[r13 - 1][12], 'nota 13');
});

test('preparaCassa2 non sovrascrive tab C2 già esistenti (nessuna modifica)', () => {
  const { preparaCassa2, ss, sheets } = loadPrepara(2026, ALL);
  preparaCassa2(2026, 12);
  ss.getSheetByName('DICEMBRE C2').cells[5][2] = 999;
  const before = sheets.length;

  assert.throws(() => preparaCassa2(2026, 11), /DICEMBRE C2/);
  assert.strictEqual(sheets.length, before);
  assert.strictEqual(ss.getSheetByName('DICEMBRE C2').cells[5][2], 999);
});

test('preparaCassa2 si ferma senza creare nulla se le date di un mese non tornano', () => {
  const { preparaCassa2, ss, sheets } = loadPrepara(2026, ALL);
  ss.getSheetByName('DICEMBRE').cells[3 + 5][0] = 'mar 06/12';   // 06/12/2026 è domenica
  const before = sheets.length;

  assert.throws(() => preparaCassa2(2026, 11), /DICEMBRE/);
  assert.strictEqual(sheets.length, before);
});

test('preparaCassa2 rimette la formula del totale (L) in C2 e TOT anche se nella Cassa 1 c\'è un valore fisso', () => {
  const { preparaCassa2, ss } = loadPrepara(2026, ALL);
  ss.getSheetByName('DICEMBRE').cells[3 + 4][11] = 1234.5;      // sab 05/12 → riga 8, L scritto a mano

  preparaCassa2(2026, 12);

  assert.strictEqual(ss.getSheetByName('DICEMBRE C2').cells[7][11],  '=SUM(B8:J8)-K8');
  assert.strictEqual(ss.getSheetByName('DICEMBRE TOT').cells[7][11], '=SUM(B8:J8)-K8');
  assert.strictEqual(ss.getSheetByName('DICEMBRE').cells[7][11], 1234.5);   // Cassa 1 intatta
});

test('preparaCassa2: se si blocca a metà cancella i tab appena creati e si può rilanciare', () => {
  const { preparaCassa2, ss, sheets } = loadPrepara(2026, ALL);
  const before = sheets.map(s => s.name);
  ss._failOn = 'DICEMBRE TOT';

  assert.throws(() => preparaCassa2(2026, 11), /DICEMBRE TOT/);
  assert.deepStrictEqual(sheets.map(s => s.name), before);

  ss._failOn = null;
  preparaCassa2(2026, 11);
  assert.ok(ss.getSheetByName('DICEMBRE TOT'));
});
