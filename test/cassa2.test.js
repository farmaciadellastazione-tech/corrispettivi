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
