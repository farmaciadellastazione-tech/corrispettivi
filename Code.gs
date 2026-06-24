// Google Apps Script — Corrispettivi (solo scrittura nel foglio)
// L'analisi AI avviene nel browser — nessuna configurazione necessaria qui

const MONTH_NAMES = ['GENNAIO','FEBBRAIO','MARZO','APRILE','MAGGIO','GIUGNO',
                     'LUGLIO','AGOSTO','SETTEMBRE','OTTOBRE','NOVEMBRE','DICEMBRE'];
const DAY_NAMES   = ['dom','lun','mar','mer','gio','ven','sab'];
const FOGLI       = { 2026: '1sXz2Mh3jTeTElGy0UAIw8UgO9RSuHaV5QOXNPJVMjo4' };

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    return out(writeToSheet(req.fields, req.date, req.note || ''));
  } catch(err) {
    return out({ error: err.message });
  }
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
         .setMimeType(ContentService.MimeType.JSON);
}

function writeToSheet(fields, dateStr, note) {
  const date      = new Date(dateStr + 'T12:00:00');
  const year      = date.getFullYear();
  const ssId      = FOGLI[year];
  if (!ssId) throw new Error('Foglio ' + year + ' non configurato nella costante FOGLI');

  const ss        = SpreadsheetApp.openById(ssId);
  const monthName = MONTH_NAMES[date.getMonth()];
  const sheet     = ss.getSheetByName(monthName);
  if (!sheet) throw new Error('Tab "' + monthName + '" non trovato');

  const label = formatLabel(date);
  const rows  = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
  let targetRow = -1;
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === label) { targetRow = i + 1; break; }
  }
  if (targetRow < 0) throw new Error('Data "' + label + '" non trovata nel tab ' + monthName);

  const F     = fields;
  const vals  = [F.asl, F.farmaco, F.parafarmaco, F.varie, F.servizi0, F.serviziIva, F.shopper, F.dpi, F.fatture, F.scontrini];
  const total = vals.reduce((s, v) => s + (parseFloat(v) || 0), 0);

  sheet.getRange(targetRow, 1, 1, 13).setValues([[
    label,
    n(F.asl), n(F.farmaco), n(F.parafarmaco), n(F.varie),
    n(F.servizi0), n(F.serviziIva), n(F.shopper), n(F.dpi), n(F.fatture), n(F.scontrini),
    total, note
  ]]);

  return { date: label, tab: monthName, total: total };
}

function n(x) { const v = parseFloat(x); return v ? v : ''; }

function formatLabel(date) {
  return DAY_NAMES[date.getDay()] + ' ' +
         String(date.getDate()).padStart(2, '0') + '/' +
         String(date.getMonth() + 1).padStart(2, '0');
}
