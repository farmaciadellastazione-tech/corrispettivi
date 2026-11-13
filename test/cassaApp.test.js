// Scelta della cassa in index.html (node --test, DOM finto).
const test   = require('node:test');
const assert = require('node:assert');
const fs     = require('node:fs');
const path   = require('node:path');

const html   = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
const script = html.match(/<script>([\s\S]*?)<\/script>/)[1];

function fakeEl() {
  return {
    style: {}, value: '', checked: false, disabled: false, textContent: '', innerHTML: '', src: '', className: '',
    classList: { add() {}, remove() {}, toggle() {} },
    addEventListener() {}, scrollIntoView() {}, appendChild() {}, remove() {},
  };
}

function loadApp(store = {}) {
  const els = {};
  const calls = [];
  const fetch = async (url, opts) => {
    calls.push(JSON.parse(opts.body));
    return { text: async () => JSON.stringify({ date: 'ven 13/11', tab: 'NOVEMBRE', total: 100 }) };
  };
  const document = {
    getElementById: id => (els[id] ||= fakeEl()),
    querySelectorAll: () => [],
    createElement: () => fakeEl(),
    body: fakeEl(),
  };
  const window       = { addEventListener() {} };
  const navigator    = {};
  const localStorage = { getItem: k => (k in store ? store[k] : null), setItem: (k, v) => { store[k] = String(v); } };
  const api = new Function('document', 'window', 'navigator', 'localStorage', 'fetch',
    script + '\nreturn { writeToSheet, renderResult, resetAll, saveSettings, loadSettings };')(document, window, navigator, localStorage, fetch);
  return { api, el: document.getElementById, calls, store };
}

test('una sola cassa (impostazione di default): si scrive senza scegliere, nella Cassa 1', async () => {
  const { api, el, calls } = loadApp();
  api.loadSettings();
  api.renderResult({ data: '2026-11-13', farmaco: 100 });

  assert.strictEqual(el('cassa-row').style.display, 'none');
  await api.writeToSheet();

  assert.strictEqual(calls.length, 1);
  assert.strictEqual(calls[0].cassa, undefined);
});

test('due casse attive: senza scegliere la cassa non si scrive nulla', async () => {
  const { api, el, calls } = loadApp({ due_casse: '1' });
  api.loadSettings();
  api.renderResult({ data: '2026-11-13', farmaco: 100 });

  assert.notStrictEqual(el('cassa-row').style.display, 'none');
  await api.writeToSheet();

  assert.strictEqual(calls.length, 0);
  assert.match(el('error-box').textContent, /cassa/i);
});

test('due casse attive: la cassa scelta viene inviata e mostrata nella conferma', async () => {
  const { api, el, calls } = loadApp({ due_casse: '1' });
  api.loadSettings();
  api.renderResult({ data: '2026-11-13', farmaco: 100 });
  el('cassa-input').value = '2';

  await api.writeToSheet();

  assert.strictEqual(calls[0].cassa, 2);
  assert.match(el('success-detail').textContent, /Cassa 2/);
});

test('un nuovo scontrino letto azzera la scelta della cassa', async () => {
  const { api, el, calls } = loadApp({ due_casse: '1' });
  api.loadSettings();
  api.renderResult({ data: '2026-11-13', farmaco: 100 });
  el('cassa-input').value = '2';
  await api.writeToSheet();
  api.resetAll();

  api.renderResult({ data: '2026-11-13', farmaco: 200 });
  await api.writeToSheet();

  assert.strictEqual(el('cassa-input').value, '');
  assert.strictEqual(calls.length, 1);
});

test('Impostazioni: la spunta "Seconda cassa" si salva e si può togliere', () => {
  const { api, el, store } = loadApp();
  el('due-casse-input').checked = true;
  api.saveSettings();
  assert.strictEqual(store.due_casse, '1');

  el('due-casse-input').checked = false;
  api.saveSettings();
  assert.strictEqual(store.due_casse, '0');
});
