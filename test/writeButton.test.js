// Test del pulsante "Scrivi nel foglio" di index.html con un DOM finto (node --test).
const test   = require('node:test');
const assert = require('node:assert');
const fs     = require('node:fs');
const path   = require('node:path');

const html   = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
const script = html.match(/<script>([\s\S]*?)<\/script>/)[1];

function fakeEl() {
  return {
    style: {}, value: '', disabled: false, textContent: '', innerHTML: '', src: '', className: '',
    classList: { add() {}, remove() {}, toggle() {} },
    addEventListener() {}, scrollIntoView() {}, appendChild() {}, remove() {},
  };
}

function loadApp(fetchImpl) {
  const els = {};
  const document = {
    getElementById: id => (els[id] ||= fakeEl()),
    querySelectorAll: () => [],
    createElement: () => fakeEl(),
    body: fakeEl(),
  };
  const window       = { addEventListener() {} };
  const navigator    = {};
  const localStorage = { getItem: () => null, setItem() {} };
  const api = new Function('document', 'window', 'navigator', 'localStorage', 'fetch',
    script + '\nreturn { writeToSheet, resetAll, renderResult };')(document, window, navigator, localStorage, fetchImpl);
  return { api, els: document.getElementById };
}

const okFetch = async () => ({ text: async () => JSON.stringify({ date: 'ven 30/10', tab: 'OTTOBRE', total: 100 }) });

test('dopo una scrittura riuscita e "Nuovo scontrino" il pulsante Scrivi nel foglio torna attivo', async () => {
  const { api, els } = loadApp(okFetch);
  els('date-input').value = '2026-10-30';
  await api.writeToSheet();
  api.resetAll();
  api.renderResult({ data: '2026-10-31', farmaco: 2554.82 });
  assert.strictEqual(els('write-btn').disabled, false);
});

test('un nuovo risultato riattiva il pulsante anche senza passare da "Nuovo scontrino"', async () => {
  const { api, els } = loadApp(okFetch);
  els('date-input').value = '2026-10-30';
  await api.writeToSheet();
  api.renderResult({ data: '2026-10-31', farmaco: 2554.82 });
  assert.strictEqual(els('write-btn').disabled, false);
});

test('mentre una scrittura è in corso, un secondo tocco non invia un altro salvataggio', async () => {
  const pending = [];
  const slowFetch = () => new Promise(r => pending.push(r));
  const { api, els } = loadApp(slowFetch);
  els('date-input').value = '2026-10-30';
  const first = api.writeToSheet();
  api.resetAll();
  api.renderResult({ data: '2026-10-31', farmaco: 2554.82 });
  els('date-input').value = '2026-10-31';
  const second = api.writeToSheet();
  const calls = pending.length;
  const ok = { text: async () => JSON.stringify({ date: 'ven 30/10', tab: 'OTTOBRE', total: 100 }) };
  pending.forEach(r => r(ok));
  await Promise.all([first, second]);
  assert.strictEqual(calls, 1);
});
