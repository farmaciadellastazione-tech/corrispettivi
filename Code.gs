// Google Apps Script — Corrispettivi (solo scrittura nel foglio)
// L'analisi AI avviene nel browser — nessuna configurazione necessaria qui

const MONTH_NAMES = ['GENNAIO','FEBBRAIO','MARZO','APRILE','MAGGIO','GIUGNO',
                     'LUGLIO','AGOSTO','SETTEMBRE','OTTOBRE','NOVEMBRE','DICEMBRE'];
const DAY_NAMES   = ['dom','lun','mar','mer','gio','ven','sab'];
const FOGLI          = { 2026: '1sXz2Mh3jTeTElGy0UAIw8UgO9RSuHaV5QOXNPJVMjo4' };
const RECIPIENT_EMAIL = 'antonella.nordi@studioassociatobc.com';
const COL_HEADERS     = ['Data','ASL','Farmaco','Parafarmaco','Varie','Servizi 0%','Servizi IVAti','Shopper','DPI','Fatture','Scont. annullati','Totale','Note'];

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    if (req.action === 'sendReport') return out(sendMonthReport(req.year, req.month));
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
  const total = [F.asl, F.farmaco, F.parafarmaco, F.varie, F.servizi0, F.serviziIva, F.shopper, F.dpi, F.fatture]
                .reduce((s, v) => s + (parseFloat(v) || 0), 0) - (parseFloat(F.scontrini) || 0);

  sheet.getRange(targetRow, 1, 1, 11).setValues([[
    label,
    n(F.asl), n(F.farmaco), n(F.parafarmaco), n(F.varie),
    n(F.servizi0), n(F.serviziIva), n(F.shopper), n(F.dpi), n(F.fatture), n(F.scontrini)
  ]]);
  sheet.getRange(targetRow, 12).setFormula(`=SUM(B${targetRow}:J${targetRow})-K${targetRow}`);
  sheet.getRange(targetRow, 13).setValue(note);

  return { date: label, tab: monthName, total: total };
}

function n(x) { const v = parseFloat(x); return v ? v : ''; }

function sendMonthReport(year, month) {
  const ssId = FOGLI[year];
  if (!ssId) throw new Error('Foglio ' + year + ' non configurato nella costante FOGLI');
  const ss        = SpreadsheetApp.openById(ssId);
  const monthName = MONTH_NAMES[month - 1];
  const sheet     = ss.getSheetByName(monthName);
  if (!sheet) throw new Error('Tab "' + monthName + '" non trovato');
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) throw new Error('Nessun dato nel tab ' + monthName);
  const data = sheet.getRange(1, 1, lastRow, 13).getValues();
  const rows = data.filter(r => /^[a-z]{3}\s\d{2}\/\d{2}$/.test(String(r[0]).trim()));
  if (rows.length === 0) throw new Error('Nessuna riga con dati trovata per ' + monthName + ' ' + year);

  // CSV con BOM per Excel
  const csvLines = [COL_HEADERS.join(';')];
  rows.forEach(r => csvLines.push(r.map(v => String(v).replace(/;/g, ',')).join(';')));
  const csv = '﻿' + csvLines.join('\r\n');

  // Tabella HTML
  let tbl = '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:sans-serif;font-size:13px">';
  tbl += '<thead><tr>' + COL_HEADERS.map(h => `<th style="background:#4f8ef7;color:#fff;padding:6px 10px;text-align:left">${h}</th>`).join('') + '</tr></thead><tbody>';
  rows.forEach((r, i) => {
    tbl += `<tr style="background:${i % 2 === 0 ? '#f5f5f5' : '#fff'}">` +
      r.map(v => `<td style="padding:5px 10px">${v !== '' ? v : ''}</td>`).join('') + '</tr>';
  });
  tbl += '</tbody></table>';

  const subject = `Corrispettivi ${monthName} ${year} — Farmacia della Stazione`;
  const body = `<p>Buongiorno,</p>
<p>in allegato i corrispettivi di <strong>${monthName} ${year}</strong> della Farmacia della Stazione.</p>
<p>Il foglio Google è consultabile al link:<br><a href="${ss.getUrl()}">${ss.getUrl()}</a></p>
<br>${tbl}<br>
<p style="color:#999;font-size:11px">Email generata automaticamente da Scontrino Z.</p>`;

  MailApp.sendEmail({
    to: RECIPIENT_EMAIL,
    subject: subject,
    htmlBody: body,
    attachments: [Utilities.newBlob(csv, 'text/csv', `corrispettivi_${monthName.toLowerCase()}_${year}.csv`)]
  });
  return { sent: true, month: monthName, year: year, rows: rows.length };
}

// Esegui questa funzione UNA VOLTA dall'editor Apps Script per attivare l'invio automatico il 1° di ogni mese
function setupEndOfMonthTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'triggerEndOfMonth') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('triggerEndOfMonth').timeBased().onMonthDay(1).atHour(9).create();
}

function triggerEndOfMonth() {
  const now   = new Date();
  const year  = now.getMonth() === 0 ? now.getFullYear() - 1 : now.getFullYear();
  const month = now.getMonth() === 0 ? 12 : now.getMonth(); // getMonth() è 0-based: jan=0 → report dic anno prec.
  try {
    sendMonthReport(year, month);
  } catch(err) {
    MailApp.sendEmail(RECIPIENT_EMAIL, 'Errore invio corrispettivi automatico ' + year, 'Errore: ' + err.message);
  }
}

function formatLabel(date) {
  return DAY_NAMES[date.getDay()] + ' ' +
         String(date.getDate()).padStart(2, '0') + '/' +
         String(date.getMonth() + 1).padStart(2, '0');
}
