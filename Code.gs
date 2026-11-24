/**
 * FINANZE PERSONALI — backend Google Apps Script
 * ------------------------------------------------------------
 * Database: il Google Sheet a cui è collegato questo script
 * (un foglio per entità, creati automaticamente da setup()).
 *
 * INSTALLAZIONE (una tantum, ~5 minuti) — usa un account Google PERSONALE
 * 1. Crea un nuovo Google Sheet vuoto (es. "Finanze Personali - DB").
 * 2. Menu Estensioni > Apps Script.
 * 3. Incolla questo file in Code.gs; crea un file HTML chiamato "Index"
 *    (File > + > HTML) e incolla il contenuto di index.html.
 * 4. Seleziona la funzione "setup" e premi Esegui (autorizza l'accesso).
 *    Nel Log di esecuzione compare il TOKEN API: copialo.
 *    (Lo ritrovi in Impostazioni progetto > Proprietà script > API_TOKEN.)
 * 5. (Facoltativo) Esegui "installaTrigger" per quotazioni giornaliere
 *    e backup settimanale automatico.
 * 6. Esegui il deploy > Nuovo deployment > App web
 *    Esegui come: Me — Chi ha accesso: Chiunque.
 *    Ogni operazione sui dati richiede il token: senza token non si legge nulla.
 * 7. Apri l'app (index.html ospitata, es. GitHub Pages) da qualsiasi browser,
 *    incolla URL /exec + token. Per il telefono usa il QR in Impostazioni.
 */

const APP_TITLE = 'Finanze Personali';

const SCHEMA = {
  Conti:         ['id','nome','tipo','istituto','riferimento','saldoIniziale','dataSaldo','colore','contoAddebitoId','plafond','attivo','ordine'],
  Movimenti:     ['id','data','contoId','importo','categoriaId','descrizione','note','tag','trasferimentoId','hash','origine','creato'],
  Categorie:     ['id','nome','tipo','parentId','colore','icona'],
  Budget:        ['id','categoriaId','periodo','importo'],
  Regole:        ['id','pattern','categoriaId','priorita'],
  Operazioni:    ['id','data','contoId','tipo','ticker','quantita','prezzo','commissioni','tasse','contoRegolamentoId','movimentoId','note'],
  Titoli:        ['ticker','nome','isin','classe','valuta','simboloGF','fattore','prezzo','cambio','aggiornato','fonte'],
  Obiettivi:     ['id','nome','target','scadenza','contoId','versato','colore','icona'],
  Ricorrenze:    ['id','descrizione','contoId','importo','categoriaId','frequenza','prossimaData','attiva'],
  ProfiliImport: ['id','nome','config'],
  Impostazioni:  ['chiave','valore']
};
const NUMERIC = ['saldoIniziale','plafond','ordine','importo','priorita','quantita','prezzo','commissioni','tasse','target','versato','fattore','cambio'];
const SETUP_VERSION = '4';

/* ===================== Web app ===================== */

function doGet() {
  ensureSetup_(false);
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/* ===================== Setup ===================== */

function ss_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SS_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Database non configurato: dall\'editor Apps Script esegui la funzione setup().');
  props.setProperty('SS_ID', ss.getId());
  return ss;
}

function setup() {
  ensureSetup_(true);
  const props = PropertiesService.getScriptProperties();
  let tok = props.getProperty('API_TOKEN');
  if (!tok) {
    tok = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
    props.setProperty('API_TOKEN', tok);
  }
  Logger.log('Setup completato. Database: ' + ss_().getUrl());
  Logger.log('TOKEN API (copialo nell\'app): ' + tok);
}

/* ===================== API (token obbligatorio) ===================== */

function apiMap_() {
  return {
    loadAll: loadAll_, saveRows: saveRows_, deleteRows: deleteRows_,
    refreshQuotes: refreshQuotes_, exportFile: exportFile_, restoreBackup: restoreBackup_
  };
}

function safeEq_(a, b) {
  if (a.length !== b.length) return false;
  let r = 0;
  for (let i = 0; i < a.length; i++) r |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return r === 0;
}

function handle_(payload) {
  let req;
  try { req = JSON.parse(payload || ''); } catch (e) { return { ok: false, error: 'Richiesta non valida' }; }
  const tok = PropertiesService.getScriptProperties().getProperty('API_TOKEN');
  if (!tok) return { ok: false, code: 'AUTH', error: 'Token API non configurato: esegui setup() dall\'editor Apps Script.' };
  if (typeof req.token !== 'string' || !safeEq_(req.token.trim(), tok)) {
    Utilities.sleep(700);
    return { ok: false, code: 'AUTH', error: 'Token non valido.' };
  }
  const fn = apiMap_()[req.fn];
  if (!fn) return { ok: false, error: 'Funzione non consentita: ' + req.fn };
  try {
    return { ok: true, result: fn.apply(null, Array.isArray(req.args) ? req.args : []) };
  } catch (e) {
    return { ok: false, error: String(e && e.message || e) };
  }
}

/** Chiamata dall'interfaccia servita da Apps Script (google.script.run). */
function api(payload) { return JSON.stringify(handle_(payload)); }

/** Chiamata dall'app ospitata altrove (fetch POST, Content-Type text/plain). */
function doPost(e) {
  const out = handle_(e && e.postData ? e.postData.contents : '');
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function ensureSetup_(force) {
  const props = PropertiesService.getScriptProperties();
  if (!force && props.getProperty('SETUP_V') === SETUP_VERSION) return;
  const ss = ss_();
  Object.keys(SCHEMA).forEach(function (name) {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    let hdr = sh.getLastColumn() ? header_(sh).filter(function (h) { return h !== ''; }) : [];
    const missing = SCHEMA[name].filter(function (h) { return hdr.indexOf(h) < 0; });
    if (missing.length) sh.getRange(1, hdr.length + 1, 1, missing.length).setValues([missing]);
    hdr = hdr.concat(missing);
    formatRows_(sh, hdr, 2, sh.getMaxRows() - 1);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, hdr.length).setFontWeight('bold').setBackground('#1F3D2E').setFontColor('#FFFFFF');
  });
  const cat = ss.getSheetByName('Categorie');
  if (cat.getLastRow() < 2) writeObjects_(cat, defaultCategories_());
  ['Foglio1', 'Foglio 1', 'Sheet1'].forEach(function (n) {
    const s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
  props.setProperty('SETUP_V', SETUP_VERSION);
}

function defaultCategories_() {
  const rows = [
    ['casa','Casa','uscita','','#6B8F71','🏠'],
    ['casa_affitto','Affitto / Mutuo','uscita','casa','',''],
    ['casa_utenze','Luce e gas','uscita','casa','',''],
    ['casa_acqua','Acqua e rifiuti','uscita','casa','',''],
    ['casa_tel','Internet e telefono','uscita','casa','',''],
    ['casa_cond','Condominio','uscita','casa','',''],
    ['casa_manut','Manutenzione','uscita','casa','',''],
    ['alim','Alimentari','uscita','','#9CC34F','🛒'],
    ['alim_super','Supermercato','uscita','alim','',''],
    ['alim_merc','Mercato e negozi','uscita','alim','',''],
    ['trasp','Trasporti','uscita','','#5B8DEF','🚗'],
    ['trasp_carb','Carburante','uscita','trasp','',''],
    ['trasp_pub','Mezzi pubblici','uscita','trasp','',''],
    ['trasp_assic','Assicurazione e bollo','uscita','trasp','',''],
    ['trasp_manut','Manutenzione auto','uscita','trasp','',''],
    ['trasp_park','Parcheggi e pedaggi','uscita','trasp','',''],
    ['rist','Ristoranti e bar','uscita','','#F08A4B','🍽️'],
    ['salute','Salute','uscita','','#E25C6B','💊'],
    ['salute_farm','Farmacia','uscita','salute','',''],
    ['salute_visite','Visite mediche','uscita','salute','',''],
    ['fam','Famiglia e figli','uscita','','#B07CC6','👨‍👩‍👧'],
    ['shop','Shopping','uscita','','#E8A33D','🛍️'],
    ['shop_abb','Abbigliamento','uscita','shop','',''],
    ['shop_elett','Elettronica','uscita','shop','',''],
    ['shop_casa','Casa e arredo','uscita','shop','',''],
    ['svago','Tempo libero','uscita','','#3FB6B2','🎬'],
    ['svago_abbon','Abbonamenti','uscita','svago','',''],
    ['svago_sport','Sport','uscita','svago','',''],
    ['svago_viaggi','Viaggi','uscita','svago','',''],
    ['tasse','Tasse e imposte','uscita','','#8C7B6B','🧾'],
    ['banca','Spese bancarie','uscita','','#9AA0A6','🏦'],
    ['regali','Regali e donazioni','uscita','','#E89AB0','🎁'],
    ['altro_u','Altre uscite','uscita','','#B0AEA8','📦'],
    ['stip','Stipendio','entrata','','#2E9E62','💼'],
    ['rimb','Rimborsi','entrata','','#7CC4A0','↩️'],
    ['rendite','Rendite finanziarie','entrata','','#3E7CB1','📈'],
    ['altro_e','Altre entrate','entrata','','#A3C9A8','💰'],
    ['giroconto','Giroconto','neutro','','#9AA0A6','🔁'],
    ['investimenti','Investimenti','neutro','','#3E7CB1','📊']
  ];
  return rows.map(function (r) { return { id: r[0], nome: r[1], tipo: r[2], parentId: r[3], colore: r[4], icona: r[5] }; });
}

/* ===================== Helper fogli ===================== */

function header_(sh) {
  return sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
}

function formatRows_(sh, hdr, start, n) {
  if (n <= 0) return;
  hdr.forEach(function (h, i) {
    sh.getRange(start, i + 1, n, 1).setNumberFormat(NUMERIC.indexOf(h) >= 0 ? '#,##0.00########' : '@');
  });
}

function ensureCapacity_(sh, hdr, lastNeeded) {
  const max = sh.getMaxRows();
  if (lastNeeded > max) {
    const add = lastNeeded - max + 500;
    sh.insertRowsAfter(max, add);
    formatRows_(sh, hdr, max + 1, add);
  }
}

function toRow_(o, hdr) {
  return hdr.map(function (h) {
    const v = o[h];
    if (v === undefined || v === null) return '';
    if (NUMERIC.indexOf(h) >= 0) {
      if (v === '') return '';
      const n = Number(v);
      return isNaN(n) ? '' : n;
    }
    let s = String(v);
    if (s.charAt(0) === '=') s = "'" + s; // evita interpretazione come formula
    return s;
  });
}

function writeObjects_(sh, objs) {
  if (!objs.length) return;
  const hdr = header_(sh);
  const start = sh.getLastRow() + 1;
  ensureCapacity_(sh, hdr, start + objs.length - 1);
  sh.getRange(start, 1, objs.length, hdr.length).setValues(objs.map(function (o) { return toRow_(o, hdr); }));
}

function readSheet_(sh) {
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return [];
  const v = sh.getRange(1, 1, lr, lc).getValues();
  const hdr = v[0].map(String);
  const tz = Session.getScriptTimeZone();
  const out = [];
  for (let r = 1; r < v.length; r++) {
    if (v[r][0] === '' || v[r][0] === null) continue;
    const o = {};
    hdr.forEach(function (h, i) {
      if (!h) return;
      let x = v[r][i];
      if (x instanceof Date) x = Utilities.formatDate(x, tz, (h === 'aggiornato' || h === 'creato') ? 'yyyy-MM-dd HH:mm' : 'yyyy-MM-dd');
      o[h] = x;
    });
    out.push(o);
  }
  return out;
}

function sheet_(name) {
  if (!SCHEMA[name]) throw new Error('Entità sconosciuta: ' + name);
  return ss_().getSheetByName(name);
}

/* ===================== API chiamate dal client ===================== */

function loadAll_() {
  ensureSetup_(false);
  const ss = ss_();
  const out = {};
  Object.keys(SCHEMA).forEach(function (n) { out[n] = readSheet_(ss.getSheetByName(n)); });
  out._meta = { ssUrl: ss.getUrl(), tz: Session.getScriptTimeZone() };
  return JSON.stringify(out);
}

function saveRows_(name, json) {
  const rows = JSON.parse(json);
  if (!rows.length) return JSON.stringify({ ok: true, n: 0 });
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = sheet_(name);
    const key = SCHEMA[name][0];
    const lr = sh.getLastRow();
    const hdr = header_(sh);
    const ki = hdr.indexOf(key);
    const data = lr >= 2 ? sh.getRange(2, 1, lr - 1, hdr.length).getValues() : [];
    const index = {};
    data.forEach(function (row, i) { index[String(row[ki])] = i; });
    const appends = [], updates = [];
    rows.forEach(function (o) {
      const line = toRow_(o, hdr);
      const at = index[String(o[key])];
      if (at === undefined) { appends.push(line); index[String(o[key])] = -1; }
      else if (at >= 0) { data[at] = line; updates.push(at); }
    });
    if (updates.length > 15) {
      sh.getRange(2, 1, data.length, hdr.length).setValues(data);
    } else {
      updates.forEach(function (i) { sh.getRange(i + 2, 1, 1, hdr.length).setValues([data[i]]); });
    }
    if (appends.length) {
      const start = Math.max(sh.getLastRow(), 1) + 1;
      ensureCapacity_(sh, hdr, start + appends.length - 1);
      sh.getRange(start, 1, appends.length, hdr.length).setValues(appends);
    }
    return JSON.stringify({ ok: true, n: rows.length });
  } finally {
    lock.releaseLock();
  }
}

function deleteRows_(name, keysJson) {
  const keys = JSON.parse(keysJson).map(String);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = sheet_(name);
    const lr = sh.getLastRow();
    if (lr < 2) return JSON.stringify({ ok: true, n: 0 });
    const col = sh.getRange(2, 1, lr - 1, 1).getValues();
    const set = {};
    keys.forEach(function (k) { set[k] = true; });
    let n = 0;
    for (let i = col.length - 1; i >= 0; i--) {
      if (set[String(col[i][0])]) { sh.deleteRow(i + 2); n++; }
    }
    return JSON.stringify({ ok: true, n: n });
  } finally {
    lock.releaseLock();
  }
}

/** Aggiorna i prezzi dei titoli con simbolo GOOGLEFINANCE e fonte != manuale. */
function refreshQuotes_() {
  const ss = ss_();
  const sh = ss.getSheetByName('Titoli');
  const lr = sh.getLastRow();
  if (lr < 2) return JSON.stringify({ ok: true, updated: 0, failed: [] });
  const hdr = header_(sh);
  const c = function (n) { return hdr.indexOf(n); };
  const data = sh.getRange(2, 1, lr - 1, hdr.length).getValues();
  const reqs = [];
  data.forEach(function (r, i) {
    const sym = String(r[c('simboloGF')] || '').trim();
    if (!sym || String(r[c('fonte')]) === 'manuale') return;
    reqs.push({ i: i, sym: sym.replace(/"/g, ''), val: String(r[c('valuta')] || 'EUR').toUpperCase().trim() });
  });
  if (!reqs.length) return JSON.stringify({ ok: true, updated: 0, failed: [] });

  let tmp = ss.getSheetByName('_quote');
  if (!tmp) { tmp = ss.insertSheet('_quote'); tmp.hideSheet(); }
  tmp.clear();
  const f = reqs.map(function (q) {
    return [
      '=IFERROR(GOOGLEFINANCE("' + q.sym + '","price"),"")',
      q.val === 'EUR' ? '=1' : '=IFERROR(GOOGLEFINANCE("CURRENCY:' + q.val + 'EUR"),"")'
    ];
  });
  tmp.getRange(1, 1, f.length, 2).setFormulas(f);
  SpreadsheetApp.flush();
  let vals = tmp.getRange(1, 1, f.length, 2).getValues();
  for (let t = 0; t < 3 && vals.some(function (v) { return typeof v[0] !== 'number'; }); t++) {
    Utilities.sleep(1500);
    SpreadsheetApp.flush();
    vals = tmp.getRange(1, 1, f.length, 2).getValues();
  }
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  let updated = 0;
  const failed = [];
  reqs.forEach(function (q, k) {
    const p = vals[k][0], fx = vals[k][1];
    if (typeof p === 'number' && p > 0) {
      data[q.i][c('prezzo')] = p;
      data[q.i][c('cambio')] = (typeof fx === 'number' && fx > 0) ? fx : 1;
      data[q.i][c('aggiornato')] = now;
      updated++;
    } else failed.push(q.sym);
  });
  ['prezzo', 'cambio', 'aggiornato'].forEach(function (n) {
    sh.getRange(2, c(n) + 1, data.length, 1).setValues(data.map(function (r) { return [r[c(n)]]; }));
  });
  tmp.clear();
  return JSON.stringify({ ok: true, updated: updated, failed: failed });
}

/** Salva un file (CSV/JSON) nella cartella Drive "Finanze Personali - Export". */
function exportFile_(name, content, mime) {
  const folder = folder_('Finanze Personali - Export');
  const file = folder.createFile(name, content, mime || 'text/plain');
  return JSON.stringify({ ok: true, url: file.getUrl(), name: file.getName() });
}

/** Ripristino completo da backup JSON (sovrascrive tutti i dati). */
function restoreBackup_(json) {
  const db = JSON.parse(json);
  const lock = LockService.getScriptLock();
  lock.waitLock(60000);
  try {
    const ss = ss_();
    Object.keys(SCHEMA).forEach(function (n) {
      if (!Array.isArray(db[n])) return;
      const sh = ss.getSheetByName(n);
      const lr = sh.getLastRow();
      if (lr >= 2) sh.getRange(2, 1, lr - 1, sh.getLastColumn()).clearContent();
      writeObjects_(sh, db[n]);
    });
    return JSON.stringify({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

/* ===================== Automazioni ===================== */

function installaTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (['triggerQuotazioni', 'backupSettimanale'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('triggerQuotazioni').timeBased().everyDays(1).atHour(19).create();
  ScriptApp.newTrigger('backupSettimanale').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(6).create();
  Logger.log('Trigger installati: quotazioni giornaliere (19:00) e backup settimanale (lunedì 6:00).');
}

function triggerQuotazioni() { refreshQuotes_(); }

function backupSettimanale() {
  const ss = ss_();
  const folder = folder_('Finanze Personali - Backup');
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  DriveApp.getFileById(ss.getId()).makeCopy(ss.getName() + ' - backup ' + stamp, folder);
  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) files.push(it.next());
  files.sort(function (a, b) { return b.getDateCreated() - a.getDateCreated(); });
  files.slice(8).forEach(function (f) { f.setTrashed(true); }); // conserva gli ultimi 8
}

function folder_(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}
