// Bug dei martedì: Google Fogli (locale it) legge "mar 27/10" come 27 marzo 2010.
// Il foglio finto qui sotto imita quel comportamento quando la cella non è in formato testo.
const test   = require('node:test');
const assert = require('node:assert');
const fs     = require('node:fs');
const path   = require('node:path');

const code = fs.readFileSync(path.join(__dirname, '..', 'Codice.gs'), 'utf8');

// Conversione automatica di Google Fogli: "mar dd/yy" → Date(20yy, marzo, dd).
function autoParse(v, fmt) {
  if (fmt === '@' || typeof v !== 'string') return v;
  const m = v.match(/^mar (\d{2})\/(\d{2})$/);
  return m ? new Date(2000 + +m[2], 2, +m[1]) : v;
}

// Foglio finto di un mese: riga 1 titolo, riga 3 "Data", poi un giorno per riga, poi TOTALI.
function fakeSheet(title, dayCells) {
  const cells   = [[title], [''], ['Data']].concat(dayCells.map(v => [v]), [['TOTALI']]);
  cells.forEach(r => { while (r.length < 13) r.push(''); });
  const formats = {};
  const key = (r, c) => r + ',' + c;
  const set = (r, c, v) => { cells[r - 1][c - 1] = autoParse(v, formats[key(r, c)]); };
  return {
    cells, formats,
    getLastRow: () => cells.length,
    getRange: (row, col, nRows = 1, nCols = 1) => ({
      getValue:        () => cells[row - 1][col - 1],
      getValues:       () => cells.slice(row - 1, row - 1 + nRows).map(r => r.slice(col - 1, col - 1 + nCols)),
      setValue:        v  => set(row, col, v),
      setFormula:      f  => { cells[row - 1][col - 1] = f; },
      setValues:       vs => vs.forEach((r, i) => r.forEach((v, j) => set(row + i, col + j, v))),
      setNumberFormat: f  => { for (let i = 0; i < nRows; i++) for (let j = 0; j < nCols; j++) formats[key(row + i, col + j)] = f; },
    }),
  };
}

// Etichette come le ha il foglio vero: alcune Date (corrette), altre testo.
function ottobre() {
  const DN = ['dom','lun','mar','mer','gio','ven','sab'];
  return Array.from({ length: 31 }, (_, i) => {
    const d = new Date(2026, 9, i + 1);
    return i % 2 ? d : DN[d.getDay()] + ' ' + String(i + 1).padStart(2, '0') + '/10';
  });
}

function load(sheets) {
  const sent = [];
  const ss = { getSheetByName: n => sheets[n] || null, getUrl: () => 'https://example.invalid/foglio' };
  const SpreadsheetApp = { openById: () => ss };
  const MailApp   = { sendEmail: m => sent.push(m) };
  const Utilities = { newBlob: (data, type, name) => ({ data, type, name }) };
  const Logger    = { log() {} };
  const api = new Function('SpreadsheetApp', 'MailApp', 'Utilities', 'Logger',
    code + '\nreturn { writeToSheet, writeFattura, sendMonthReport, riparaDate: typeof riparaDate === "function" ? riparaDate : null };')(SpreadsheetApp, MailApp, Utilities, Logger);
  return { api, sent };
}

const SCONTRINO = { asl: 0, farmaco: 2000, parafarmaco: 500, varie: 0, servizi0: 0,
                    serviziIva: 6, shopper: 0, dpi: 0, fatture: 0, scontrini: 0 };

test('martedì: dopo lo scontrino la fattura dello stesso giorno trova ancora la riga', () => {
  const sheet = fakeSheet('OTTOBRE 2026', ottobre());
  const { api } = load({ OTTOBRE: sheet });

  api.writeToSheet(SCONTRINO, '2026-10-27', '');
  const res = api.writeFattura({ data: '2026-10-27', importo: 150, note: '' });

  assert.strictEqual(res.data, 'mar 27/10');
  assert.strictEqual(sheet.cells[29][9], 150);
});

test('martedì: ri-registrare lo scontrino non dà "Data non trovata"', () => {
  const sheet = fakeSheet('OTTOBRE 2026', ottobre());
  const { api } = load({ OTTOBRE: sheet });

  api.writeToSheet(SCONTRINO, '2026-10-27', '');
  assert.doesNotThrow(() => api.writeToSheet(SCONTRINO, '2026-10-27', ''));
});

test('martedì: il report mensile mostra "mar 27/10" e non una data di marzo', () => {
  const sheet = fakeSheet('OTTOBRE 2026', ottobre());
  const { api, sent } = load({ OTTOBRE: sheet });

  api.writeToSheet(SCONTRINO, '2026-10-27', '');
  api.sendMonthReport(2026, 10);

  const csv = sent[0].attachments[0].data;
  assert.match(csv, /\nmar 27\/10;/);
  assert.doesNotMatch(csv, /\/03;/);
});

test('riparaDate riscrive come testo le date-martedì corrotte senza toccare gli importi', () => {
  const days = ottobre();
  days[26] = new Date(2010, 2, 27);          // 27/10 corrotto in 27/03/2010
  const sheet = fakeSheet('OTTOBRE 2026', days);
  sheet.cells[29][2] = 2122.84;              // importo Farmaco della riga 27/10
  const others = {};
  ['GENNAIO','FEBBRAIO','MARZO','APRILE','MAGGIO','GIUGNO','LUGLIO','AGOSTO','SETTEMBRE','NOVEMBRE','DICEMBRE']
    .forEach(n => { others[n] = null; });
  const { api } = load({ ...others, OTTOBRE: sheet });

  assert.ok(api.riparaDate, 'funzione riparaDate mancante');
  api.riparaDate(2026);

  assert.strictEqual(sheet.cells[29][0], 'mar 27/10');
  assert.strictEqual(sheet.cells[3][0],  'gio 01/10');
  assert.strictEqual(sheet.cells[4][0],  'ven 02/10');
  assert.strictEqual(sheet.cells[29][2], 2122.84);
  assert.strictEqual(sheet.cells[34][0], 'TOTALI');
});

function allTabs(sheets) {
  const all = {};
  ['GENNAIO','FEBBRAIO','MARZO','APRILE','MAGGIO','GIUGNO','LUGLIO','AGOSTO','SETTEMBRE','OTTOBRE','NOVEMBRE','DICEMBRE']
    .forEach(n => { all[n] = sheets[n] || null; });
  return all;
}

test('riparaDate si ferma senza scrivere nulla se un giorno non corrisponde alla sua riga', () => {
  const set = Array.from({ length: 30 }, (_, i) => new Date(2026, 8, i + 1));
  set[7] = new Date(2009, 2, 8);               // martedì 08/09 corrotto
  const settembre = fakeSheet('SETTEMBRE 2026', set);
  const days = ottobre();
  days[3] = 'mer 05/10';                       // riga del 4/10 con l'etichetta di un altro giorno
  const ott = fakeSheet('OTTOBRE 2026', days);
  const { api } = load(allTabs({ SETTEMBRE: settembre, OTTOBRE: ott }));

  assert.throws(() => api.riparaDate(2026), /OTTOBRE riga 7/);
  assert.ok(settembre.cells[10][0] instanceof Date, 'SETTEMBRE non doveva essere toccato');
  assert.deepStrictEqual(Object.keys(settembre.formats), []);
});

test('riparaDate conta solo le date davvero corrette, non le Date già giuste', () => {
  const days = ottobre();
  days[26] = new Date(2010, 2, 27);
  const { api } = load(allTabs({ OTTOBRE: fakeSheet('OTTOBRE 2026', days) }));

  const report = api.riparaDate(2026);
  assert.ok(report.some(r => r.startsWith('OTTOBRE: 1 date corrette')), report.join(' / '));
});
