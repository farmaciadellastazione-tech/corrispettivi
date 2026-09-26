// Test di Codice.gs con un foglio Google finto in memoria (node --test).
const test   = require('node:test');
const assert = require('node:assert');
const fs     = require('node:fs');
const path   = require('node:path');

const code = fs.readFileSync(path.join(__dirname, '..', 'Codice.gs'), 'utf8');

// Foglio finto: righe 1-3 intestazioni, dalla riga 4 un giorno per riga (13 colonne).
function fakeSheet(dayLabels) {
  const cells = [['SETTEMBRE 2026'], [''], ['Data']].concat(dayLabels.map(l => [l]));
  cells.forEach(r => { while (r.length < 13) r.push(''); });
  const sheet = {
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
  return sheet;
}

function load(sheet) {
  const SpreadsheetApp = { openById: () => ({ getSheetByName: () => sheet }) };
  return new Function('SpreadsheetApp', code + '\nreturn { writeToSheet, writeFattura };')(SpreadsheetApp);
}

const SCONTRINO = { asl: 0, farmaco: 2000, parafarmaco: 500, varie: 0, servizi0: 0,
                    serviziIva: 6, shopper: 0, dpi: 0, fatture: 0, scontrini: 0 };

test('fattura inserita prima dello scontrino resta anche se la data in colonna A è testo', () => {
  const sheet = fakeSheet(['ven 25/09', 'sab 26/09']);
  const app   = load(sheet);

  app.writeFattura({ data: '2026-09-26', importo: 150, note: '20/PA' });
  app.writeToSheet(SCONTRINO, '2026-09-26', '');

  assert.strictEqual(sheet.cells[4][9], 150);
});

test('fattura inserita prima dello scontrino resta se la data in colonna A è un oggetto Date', () => {
  const sheet = fakeSheet([new Date('2026-09-25T00:00:00'), new Date('2026-09-26T00:00:00')]);
  const app   = load(sheet);

  app.writeFattura({ data: '2026-09-26', importo: 150, note: '' });
  const res = app.writeToSheet(SCONTRINO, '2026-09-26', '');

  assert.strictEqual(sheet.cells[4][9], 150);
  assert.strictEqual(res.total, 2656);
});

test('fattura letta dallo scontrino si somma a quella manuale al primo inserimento', () => {
  const sheet = fakeSheet(['sab 26/09']);
  const app   = load(sheet);

  app.writeFattura({ data: '2026-09-26', importo: 150, note: '' });
  app.writeToSheet({ ...SCONTRINO, fatture: 40 }, '2026-09-26', '');

  assert.strictEqual(sheet.cells[3][9], 190);
});

test('ri-registrare lo scontrino dello stesso giorno sovrascrive (non raddoppia) i valori', () => {
  const sheet = fakeSheet(['sab 26/09']);
  const app   = load(sheet);

  app.writeToSheet({ ...SCONTRINO, fatture: 40 }, '2026-09-26', '');
  app.writeToSheet({ ...SCONTRINO, fatture: 40 }, '2026-09-26', '');

  assert.strictEqual(sheet.cells[3][9], 40);
  assert.strictEqual(sheet.cells[3][2], 2000);
});
