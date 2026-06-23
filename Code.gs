// Google Apps Script — Scontrino Z → Corrispettivi
//
// Script Properties da configurare (Progetto → Impostazioni → Proprietà script):
//   ANTHROPIC_API_KEY  → chiave API Anthropic
//   SPREADSHEET_2026   → 1sXz2Mh3jTeTElGy0UAIw8UgO9RSuHaV5QOXNPJVMjo4
//
// Deployment → Distribuisci → Nuova distribuzione → App web
//   Esegui come: Me  |  Accesso: Chiunque

const MONTH_NAMES = [
  'GENNAIO','FEBBRAIO','MARZO','APRILE','MAGGIO','GIUGNO',
  'LUGLIO','AGOSTO','SETTEMBRE','OTTOBRE','NOVEMBRE','DICEMBRE'
];
const DAY_NAMES = ['dom','lun','mar','mer','gio','ven','sab'];

// ---------------------------------------------------------------------------
// ENTRY POINT
// ---------------------------------------------------------------------------

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    let result;
    if      (req.action === 'analyze') result = analyzeImage(req.imageBase64, req.mediaType || 'image/jpeg');
    else if (req.action === 'write')   result = writeToSheet(req.fields, req.date, req.note || '');
    else throw new Error('Azione non riconosciuta: ' + req.action);
    return jsonOut(result);
  } catch (err) {
    return jsonOut({ error: err.message });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------------------
// ANALISI IMMAGINE CON CLAUDE HAIKU
// ---------------------------------------------------------------------------

function analyzeImage(imageBase64, mediaType) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) throw new Error('ANTHROPIC_API_KEY non configurata nelle Script Properties');

  const prompt =
    'Sei un assistente per una farmacia italiana. Analizza questo scontrino Z ' +
    '(Rapporto Reparti Z) e restituisci SOLO un JSON valido.\n\n' +
    'Reparti del registratore di cassa:\n' +
    '01-USL/ASL: ricette SSN — può essere negativo per resi\n' +
    '02-FARMACO: usa "AL NETTO DI SCONTI/MAGG.", NON il lordo\n' +
    '03-PARAFARMACO: usa "AL NETTO DI SCONTI/MAGG.", NON il lordo\n' +
    '04-VARIE: varie\n' +
    '05-SERVIZI IVA 0: servizi esenti IVA (0 se assente)\n' +
    '06-SERVIZI 22%/IVAti: servizi con IVA\n' +
    'SHOPPER: sacchetti (0 se assente)\n' +
    'DPI: dispositivi protezione individuale (0 se assente)\n' +
    'FATTURE: fatture/note credito (0 se assente)\n' +
    'Scontrini annullati: totale scontrini annullati (0 se assente)\n\n' +
    'La DATA è in fondo allo scontrino, vicino a "DOC. GEST." o "CASSA" ' +
    '(formato gg-mm-aaaa). NON usare la "DATA ULTIMO AZZERAMENTO".\n\n' +
    'Restituisci SOLO questo JSON (valori numerici con punto decimale, 0 se assente):\n' +
    '{"data":"YYYY-MM-DD","asl":0,"farmaco":0,"parafarmaco":0,"varie":0,' +
    '"servizi0":0,"serviziIva":0,"shopper":0,"dpi":0,"fatture":0,"scontrini":0}';

  const payload = {
    model: 'claude-haiku-4-5-20251001',
    max_tokens: 512,
    messages: [{
      role: 'user',
      content: [
        { type: 'image', source: { type: 'base64', media_type: mediaType, data: imageBase64 } },
        { type: 'text', text: prompt }
      ]
    }]
  };

  const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const rd = JSON.parse(resp.getContentText());
  if (rd.error) throw new Error('Claude API: ' + rd.error.message);

  const text = rd.content[0].text.trim();
  const match = text.match(/\{[\s\S]*\}/);
  if (!match) throw new Error('Risposta AI non valida: ' + text.substring(0, 200));

  return JSON.parse(match[0]);
}

// ---------------------------------------------------------------------------
// SCRITTURA SUL FOGLIO
// ---------------------------------------------------------------------------

function writeToSheet(fields, dateStr, note) {
  const date      = new Date(dateStr + 'T12:00:00'); // mezzogiorno → evita problemi fuso
  const year      = date.getFullYear();
  const monthIdx  = date.getMonth();
  const monthName = MONTH_NAMES[monthIdx];
  const dateLabel = formatDateLabel(date);

  const ss    = getOrCreateSpreadsheet(year);
  const sheet = ss.getSheetByName(monthName);
  if (!sheet) throw new Error('Tab "' + monthName + '" non trovato nel foglio ' + year);

  // Cerca la riga con questa data (colonna A)
  const lastRow = sheet.getLastRow();
  const colA    = sheet.getRange(1, 1, lastRow, 1).getValues();
  let targetRow = -1;
  let totaliRow = -1;

  for (let i = 0; i < colA.length; i++) {
    const cell = String(colA[i][0]).trim();
    if (cell === dateLabel)  targetRow = i + 1;
    if (cell === 'TOTALI')   totaliRow = i + 1;
  }

  if (targetRow < 0) throw new Error('Riga "' + dateLabel + '" non trovata nel tab ' + monthName);

  const total = sumFields(fields);

  // Colonne: A=Data B=ASL C=Farmaco D=Parafarmaco E=Varie F=Serv0 G=ServIva
  //          H=Shopper I=DPI J=Fatture K=ScontrAnn L=Totale M=Note
  const rowData = [
    dateLabel,
    v(fields.asl),        v(fields.farmaco),     v(fields.parafarmaco),
    v(fields.varie),      v(fields.servizi0),     v(fields.serviziIva),
    v(fields.shopper),    v(fields.dpi),          v(fields.fatture),
    v(fields.scontrini),  total,                  note
  ];

  sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);

  return { date: dateLabel, tab: monthName, total: total };
}

// Restituisce il valore numerico solo se > 0 (celle vuote per i reparti assenti)
function v(x) { return (x && x !== 0) ? x : ''; }

function sumFields(f) {
  return (f.asl||0) + (f.farmaco||0) + (f.parafarmaco||0) + (f.varie||0) +
         (f.servizi0||0) + (f.serviziIva||0) + (f.shopper||0) +
         (f.dpi||0) + (f.fatture||0) + (f.scontrini||0);
}

function formatDateLabel(date) {
  const day = DAY_NAMES[date.getDay()];
  const dd  = String(date.getDate()).padStart(2, '0');
  const mm  = String(date.getMonth() + 1).padStart(2, '0');
  return day + ' ' + dd + '/' + mm;
}

// ---------------------------------------------------------------------------
// GESTIONE FOGLIO ANNUALE — crea automaticamente il foglio del nuovo anno
// ---------------------------------------------------------------------------

function getOrCreateSpreadsheet(year) {
  const props   = PropertiesService.getScriptProperties();
  const propKey = 'SPREADSHEET_' + year;
  const ssId    = props.getProperty(propKey);

  if (ssId) {
    try { return SpreadsheetApp.openById(ssId); } catch (e) { /* ricrea sotto */ }
  }

  // Nuovo anno: crea un foglio con tutti i 12 tab mensili pre-popolati
  const ss = SpreadsheetApp.create(year + '_Corrispettivi_Unito');
  props.setProperty(propKey, ss.getId());

  // Rinomina il foglio di default e popola tutti i mesi
  const defaultSheet = ss.getSheets()[0];
  defaultSheet.setName(MONTH_NAMES[0]);
  populateMonthSheet(defaultSheet, MONTH_NAMES[0], year, 0);

  for (let m = 1; m < 12; m++) {
    const sh = ss.insertSheet(MONTH_NAMES[m]);
    populateMonthSheet(sh, MONTH_NAMES[m], year, m);
  }

  return ss;
}

function populateMonthSheet(sheet, monthName, year, monthIdx) {
  const headers = [
    'Data', 'ASL', 'Farmaco', 'Parafarmaco', 'Varie',
    'Servizi IVA 0', 'Servizi IVAti', 'Shopper', 'DPI', 'Fatture',
    'Scont. annullati', 'Totale giornaliero', 'Note'
  ];

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#e8eaf0');

  const daysInMonth = new Date(year, monthIdx + 1, 0).getDate();
  const rows = [];
  for (let d = 1; d <= daysInMonth; d++) {
    const date = new Date(year, monthIdx, d);
    rows.push([formatDateLabel(date), '', '', '', '', '', '', '', '', '', '', '', '']);
  }
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  // Riga TOTALI con formule SUM
  const totRow   = daysInMonth + 2;
  const totData  = ['TOTALI'];
  for (let col = 2; col <= 12; col++) {
    totData.push('=SUM(' + colLetter(col) + '2:' + colLetter(col) + (totRow - 1) + ')');
  }
  totData.push('');
  sheet.getRange(totRow, 1, 1, headers.length).setValues([totData]);
  sheet.getRange(totRow, 1, 1, headers.length).setFontWeight('bold');

  // Formato valuta
  sheet.getRange(2, 2, daysInMonth + 1, 11).setNumberFormat('€ #,##0.00');
  sheet.setColumnWidth(1, 80);
  for (let c = 2; c <= 12; c++) sheet.setColumnWidth(c, 100);
  sheet.setColumnWidth(13, 180);
}

function colLetter(n) {
  let r = '';
  while (n > 0) { const rem = (n - 1) % 26; r = String.fromCharCode(65 + rem) + r; n = Math.floor((n - 1) / 26); }
  return r;
}
