// Server-Test: Code.gs in vm mit Apps-Script-Mocks (Sheet, Drive, Mail, Slack, Trigger)
const fs = require('fs'), vm = require('vm');
let pass = 0, fail = 0;
function ok(c, label, extra) { if (c) { pass++; console.log('  ✔', label); } else { fail++; console.log('  ✖', label, extra === undefined ? '' : JSON.stringify(extra)); } }

function colNum(s) { let n = 0; for (const ch of s) n = n * 26 + ch.charCodeAt(0) - 64; return n; }
function colName(n) { let s = ''; while (n > 0) { s = String.fromCharCode(65 + (n - 1) % 26) + s; n = Math.floor((n - 1) / 26); } return s; }

// ---------- Mock-Blatt ----------
function makeSheet(name, id) {
  const cells = {};
  const sh = {
    _name: name, cells,
    getName: () => name, getSheetId: () => id,
    getLastRow: () => Math.max(0, ...Object.keys(cells).map(k => +k.match(/\d+/)[0])),
    getRange(a, b, c, d) {
      let r1, c1, r2, c2, addr = null;
      if (typeof a === 'string') { addr = a.toUpperCase(); const m = /^([A-Z]+)(\d+)$/.exec(addr); if (!m) throw new Error('Bad A1 ' + a); c1 = c2 = colNum(m[1]); r1 = r2 = +m[2]; }
      else { r1 = a; c1 = b; r2 = r1 + (c || 1) - 1; c2 = c1 + (d || 1) - 1; }
      const key = (r, cc) => colName(cc) + r;
      const rng = {
        getValue: () => sh.read(key(r1, c1)),
        setValue: (v) => { sh.write(key(r1, c1), v); return rng; },
        getValues: () => { const out = []; for (let r = r1; r <= r2; r++) { const row = []; for (let cc = c1; cc <= c2; cc++) row.push(sh.read(key(r, cc))); out.push(row); } return out; },
        setValues: (vs) => { vs.forEach((row, i) => row.forEach((v, j) => sh.write(key(r1 + i, c1 + j), v))); return rng; },
        getSheet: () => sh, getA1Notation: () => addr,
        insertCheckboxes: () => rng,
        clearContent: () => { for (let r = r1; r <= r2; r++) for (let cc = c1; cc <= c2; cc++) sh.write(key(r, cc), ''); return rng; },
      };
      ['setFontWeight', 'setBackground', 'setFontColor', 'setNumberFormat', 'setDataValidation'].forEach(m => rng[m] = () => rng);
      return rng;
    },
    getDataRange() { const lr = sh.getLastRow(); return sh.getRange(1, 1, lr, 4); },
    setColumnWidth() {}, setFrozenRows() {},
    read(k) { return k in cells ? cells[k] : ''; },
    write(k, v) { cells[k] = v; },
  };
  return sh;
}

// Protokoll-Blatt: Formeln unabhängig von Code.gs nachgebaut (Vorlage, Wertespalte K)
function makeProtokoll() {
  const sh = makeSheet('Zählprotokoll', 1);
  const N = (k) => Number(sh.cells[k] || 0);
  const calc = {};
  const rows = { B: 5, C: 10, D: 20, E: 50, F: 100 };
  const sum = (row) => Object.keys(rows).reduce((s, c) => s + rows[c] * N(c + row), 0);
  const f = {
    K7: () => N('G7') + N('I7'),
    K11: () => N('G7') + N('I7') + N('K9'),
    K17: () => sum(16),
    K20: () => 'BCDEFGHI'.split('').reduce((s, c) => s + N(c + '20'), 0),
    K29: () => sum(28) + N('K30'),
    K34: () => sum(33) + N('K35'),
    K37: () => f.K17() + f.K20() + N('K24') + f.K29() + f.K34(),
    K39: () => f.K29() + f.K34(),
    K41: () => f.K37() - f.K39(),
  };
  const origRead = sh.read;
  sh.read = (k) => (k in f ? Math.round(f[k]() * 100) / 100 : origRead(k));
  sh.cells.K24 = 337;
  [5, 10, 20, 50, 100].forEach((b, i) => { sh.cells['BCDEF'[i] + '15'] = [8, 7, 4, 1, 0][i]; });
  return sh;
}

// ---------- Mock-Drive ----------
let idSeq = 0;
function makeFolder(name, access) {
  const f = { id: 'F' + (++idSeq), name, folders: [], files: [], access: access || 'OWNER' };
  f.getName = () => name; f.getId = () => f.id;
  f.getFoldersByName = (n) => { const l = f.folders.filter(x => x.name === n); let i = 0; return { hasNext: () => i < l.length, next: () => l[i++] }; };
  f.createFolder = (n) => { if (f.access === 'VIEW') throw new Error('Keine Berechtigung'); const c = makeFolder(n); c.parent = f; f.folders.push(c); drive.byId[c.id] = c; return c; };
  f.getFilesByName = (n) => { const l = f.files.filter(x => x.name === n && !x.trashed); let i = 0; return { hasNext: () => i < l.length, next: () => l[i++] }; };
  f.createFile = (blob) => { if (f.access === 'VIEW') throw new Error('Keine Berechtigung'); return addFile(f, blob.getName(), blob); };
  f.getAccess = () => f.access;
  return f;
}
function addFile(folder, name, blob) {
  const file = { name, blob, trashed: false, id: 'Dfile' + String(++idSeq).padStart(27, '0') };
  file.getName = () => file.name; file.getId = () => file.id; file.getUrl = () => 'https://drive/' + file.id; file.setTrashed = (t) => { file.trashed = t; };
  folder.files.push(file); drive.byId[file.id] = file;
  return file;
}
const drive = { byId: {} };

function build(opts) {
  opts = opts || {};
  idSeq = 0; drive.byId = {};
  const root = makeFolder('01_Kasse'); drive.byId[root.id] = root;
  const fallback = makeFolder('Ausweich'); drive.byId[fallback.id] = fallback;
  const proto = makeProtokoll();
  const sheets = [proto];
  const props = {};
  const log = { mails: [], fetches: [], toasts: [], triggers: [] };
  const ssFile = { id: 'SS1', makeCopy: (name, folder) => addFile(folder, name, { copyOf: 'SS1', snapshot: JSON.parse(JSON.stringify(proto.cells)) }), getId: () => 'SS1' };
  drive.byId['SS1'] = ssFile;

  const ss = {
    getId: () => 'SS1', getSpreadsheetTimeZone: () => 'Europe/Berlin',
    getSheetByName: (n) => sheets.find(s => s.getName() === n) || null,
    getSheets: () => sheets,
    insertSheet: (n) => { const s = makeSheet(n, sheets.length + 1); sheets.push(s); return s; },
    toast: (m) => log.toasts.push(m),
  };
  const ctx = vm.createContext({
    console,
    Utilities: {
      formatDate: (d, tz, pat) => {
        const p = {}; new Intl.DateTimeFormat('en-GB', { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hour12: false, weekday: 'short' })
          .formatToParts(d).forEach(x => p[x.type] = x.value);
        const wd = { Mon: 1, Tue: 2, Wed: 3, Thu: 4, Fri: 5, Sat: 6, Sun: 7 }[p.weekday];
        return pat.replace('yyyy', p.year).replace('MM', p.month).replace('dd', p.day).replace('HH', p.hour).replace('mm', p.minute).replace(/^u$/, String(wd));
      },
      base64Decode: (s) => Array.from(Buffer.from(s, 'base64')).map(b => b > 127 ? b - 256 : b),
      newBlob: (bytes, type, name) => ({ _bytes: bytes, name, getName() { return this.name; }, setName(n) { this.name = n; return this; }, getDataAsString: () => Buffer.from(bytes.map(b => b & 255)).toString('latin1') }),
    },
    LockService: { getScriptLock: () => ({ tryLock: () => !opts.locked, releaseLock() {} }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => props[k] === undefined ? null : props[k], setProperty: (k, v) => { props[k] = v; } }) },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => (opts.noActive ? null : ss), openById: () => ss, flush() {},
      newDataValidation: () => ({ requireValueInList() { return this; }, build() { return {}; } }),
    },
    DriveApp: {
      getFolderById: (id) => { if (opts.driveDown) throw new Error('Kein Zugriff'); const f = drive.byId[id]; if (!f || !f.folders) throw new Error('Element nicht gefunden'); return f; },
      getFileById: (id) => drive.byId[id],
      Permission: { OWNER: 'OWNER', EDIT: 'EDIT', ORGANIZER: 'ORGANIZER', FILE_ORGANIZER: 'FILE_ORGANIZER', VIEW: 'VIEW' },
    },
    Session: { getEffectiveUser: () => 'u' },
    ScriptApp: {
      getOAuthToken: () => 'tok', getProjectTriggers: () => log.triggers,
      deleteTrigger: (t) => { log.triggers = log.triggers.filter(x => x !== t); },
    },
    MailApp: { sendEmail: (m) => { if (opts.mailFails) throw new Error('Mail-Kontingent'); log.mails.push(m); } },
    UrlFetchApp: {
      fetch: (url, o) => {
        log.fetches.push({ url, o });
        if (/hooks\.slack\.com/.test(url)) return { getResponseCode: () => opts.slackCode || 200 };
        if (opts.exportCode) return { getResponseCode: () => opts.exportCode, getBlob: () => null };
        const pages = opts.pages || 1; let txt = '%PDF-1.4\n';
        for (let i = 0; i < pages; i++) txt += '<< /Type /Page /Parent 1 0 R >>\n';
        const bytes = Array.from(Buffer.from(txt, 'latin1')).map(b => b > 127 ? b - 256 : b);
        const blob = { name: 'x.pdf', _bytes: bytes, getName() { return this.name; }, setName(n) { this.name = n; return this; }, getDataAsString: () => txt };
        return { getResponseCode: () => 200, getBlob: () => blob };
      },
    },
    HtmlService: { createHtmlOutputFromFile: () => ({ setTitle() { return this; }, addMetaTag() { return this; } }) },
  });
  vm.runInContext(fs.readFileSync(require('path').join(__dirname, '..', 'src', 'Code.gs'), 'utf8'), ctx);
  const run = (s) => vm.runInContext(s, ctx);
  const setCfg = (k, v) => { ctx.__k = k; ctx.__v = v; run('setConfigValue_(getSS_(), __k, __v)'); };
  // Config anlegen und Root setzen
  run('createConfigSheet()');
  setCfg('Root-Ordner-ID', root.id);
  return { ctx, run, ss, proto, root, fallback, drive, log, props, setCfg, opts };
}

const basePayload = (over) => Object.assign({
  date: '2026-11-12', name: 'Laura', diff: 0,
  v: { ec_links: 100.5, ec_rechts: 20, rechnungen: 0, ist_5: 8, ist_10: 6, ist_20: 4, ist_50: 0, ist_100: 0,
    klein_200: 116, klein_100: 53, klein_50: 14.5, klein_20: 9.2, klein_10: 1.6, klein_5: 2.05, klein_2: 0.34, klein_1: 0.14,
    ent1_5: 0, ent1_10: 0, ent1_20: 7, ent1_50: 1, ent1_100: 1 } }, over || {});
function save(env, payload) { env.ctx.__p = payload; return env.run('saveProtokoll(__p)'); }
function find(folder, path) { let c = folder; for (const p of path.split('/')) { c = c.folders.find(x => x.name === p); if (!c) return null; } return c; }
const names = (f) => f ? f.files.map(x => x.name) : null;

// ============ 1. Config-Aufbau & Upgrade ============
console.log('1. Config');
{
  const e = build();
  const cfgSheet = e.ss.getSheetByName('Config');
  const keys = cfgSheet.getDataRange().getValues().map(r => r[0]);
  ['Differenz-Prüfung', 'Differenz-Toleranz', 'Differenz-Pfad', 'Kopie-Name', 'Belege-Pfad', 'Belege-Dateiname', 'Belege Pflicht', 'Ausweich-Ordner-ID', 'E-Mail-Empfänger', 'Slack-Webhook', 'Erinnerung-Uhrzeit', 'Erinnerung-Tage']
    .forEach(k => ok(keys.indexOf(k) > -1, 'Schlüssel vorhanden: ' + k));
  // alter Config-Reiter nur mit 3 Schlüsseln wird ergänzt, bestehende Werte bleiben
  const e2 = build({ noActive: false });
  const sh = e2.ss.getSheetByName('Config');
  Object.keys(sh.cells).forEach(k => delete sh.cells[k]);
  sh.getRange(1, 1, 4, 3).setValues([['Schlüssel', 'Wert', 'Hinweis'], ['Zielpfad', 'Alt/Pfad', ''], ['Dateiname', 'X.pdf', ''], ['Root-Ordner-ID', 'abc', '']]);
  const cfg = e2.run('readConfig_(getSS_())');
  ok(cfg['Zielpfad'] === 'Alt/Pfad' && cfg['Dateiname'] === 'X.pdf', 'bestehende Werte bleiben beim Ergänzen');
  ok(cfg['Differenz-Prüfung'] === 'ja' && cfg['Belege-Pfad'].indexOf('EC-Belege') > -1, 'neue Schlüssel mit Standard ergänzt');
  const fld = e2.run('readFieldCells_(getSS_())');
  ok(fld['differenz'] === undefined, 'Felder-Reiter hat keinen Differenz-Eintrag mehr');
}

// ============ 2. Normaler Abschluss → PDF ============
console.log('2. Abschluss geht auf → PDF');
{
  const e = build();
  const init = e.run('getInitData()');
  ok(init.today.match(/^\d{4}-\d\d-\d\d$/) && init.soll.join(',') === '8,7,4,1,0' && init.tresor === 337, 'getInitData liefert Soll/Tresor', init);
  ok(init.diffCheck === true && init.tol === 0 && init.notify === false && init.belegePflicht === false, 'getInitData Schalter', init);
  const r = save(e, basePayload());
  ok(r.ok && r.kind === 'pdf', 'Ergebnis ok/pdf', r);
  const folder = find(e.root, 'Abschlüsse_2026/11/Zählprotokolle');
  ok(folder && names(folder).join() === '2026-11-12_Zählprotokoll.pdf', 'PDF im Monatsordner', names(folder));
  ok(!find(e.root, 'Abschlüsse_2026/11/Differenzen'), 'kein Differenzen-Ordner angelegt');
  ok(e.proto.cells.K45 === 'Laura' && e.proto.cells.K46 === undefined, 'Mitarbeiter im Blatt, kein Differenztext', [e.proto.cells.K45, e.proto.cells.K46]);
  ok(e.log.mails.length === 0 && e.log.fetches.length === 1, 'keine Meldung, nur der PDF-Export');
  const last = JSON.parse(e.props.LAST_EXPORT);
  ok(last.day === '2026-11-12' && last.kind === 'pdf', 'LAST_EXPORT gemerkt', last);
  ok(/Letzter Export/.test(JSON.stringify(e.ss.getSheetByName('Config').cells)), 'Status in Config');
  const r2 = save(e, basePayload());
  ok(names(folder).sort().join() === '2026-11-12_Zählprotokoll.pdf,2026-11-12_Zählprotokoll_2.pdf', 'zweites Speichern nummeriert', names(folder));
  e.setCfg('Wenn Datei existiert', 'ersetzen');
  save(e, basePayload());
  ok(folder.files.filter(f => !f.trashed).length === 2 && folder.files.some(f => f.trashed), 'ersetzen: Original im Papierkorb');
}

// ============ 3. Differenz → Kopie ============
{
  // Meldung bei jedem Abschluss (auch ohne Differenz)
  const em = build(); em.setCfg('E-Mail-Empfänger', 'a@x.de'); em.setCfg('Slack-Webhook', 'https://hooks.slack.com/services/T/B/X');
  const rm = save(em, basePayload());
  ok(rm.ok && rm.kind === 'pdf' && em.log.mails.length === 1 && /gespeichert/.test(em.log.mails[0].subject) && !/Differenz 0/.test(em.log.mails[0].subject), 'PDF-Fall: Mail bei jedem Abschluss', em.log.mails.map(m => m.subject));
  { const bd = em.log.mails[0].body;
    ok(/Datum: /.test(bd) && !/Mitarbeiter|Laura/.test(bd) && /Barbestand SOLL/.test(bd) && /Barbestand IST/.test(bd) && /Differenz: /.test(bd) && /PDF: https:\/\/drive\//.test(bd) && /EC-Belege \(Scan\): keine/.test(bd) && !/Unbarer Umsatz|Barentnahme/.test(bd), 'PDF-Mail: Kerndaten + Links', bd); }
  ok(em.log.fetches.some(f => /slack/.test(f.url)) && /Meldung gesendet: E-Mail, Slack/.test(rm.message), 'PDF-Fall: auch Slack + Hinweis in der Antwort', rm.message);
  const en = build(); en.setCfg('E-Mail-Empfänger', 'a@x.de'); en.setCfg('Meldung bei jedem Abschluss', 'nein');
  save(en, basePayload());
  ok(en.log.mails.length === 0, 'Schalter „nein“: PDF-Fall ohne Mail');
}
console.log('3. Differenz → Kopie statt PDF');
{
  const e = build();
  const r = save(e, basePayload({ diff: -12.5 }));
  ok(r.ok && r.kind === 'kopie', 'kind=kopie', r);
  const d = find(e.root, 'Abschlüsse_2026/11/Differenzen');
  ok(d && names(d).join() === '2026-11-12_Zählprotokoll_DIFFERENZ_-12,50', 'Kopie benannt', names(d));
  ok(!find(e.root, 'Abschlüsse_2026/11/Zählprotokolle'), 'kein PDF-Ordner/PDF');
  ok(e.log.fetches.length === 0, 'kein PDF-Export aufgerufen');
  ok(d.files[0].blob.snapshot.K46 === undefined && d.files[0].blob.snapshot.K45 === 'Laura', 'Kopie enthält Eingaben, aber keinen Differenztext');
  ok(r.message.indexOf('KEIN PDF') > -1 && e.log.mails.length === 0, 'ohne Empfänger keine Mail');
  ok(JSON.parse(e.props.LAST_EXPORT).kind === 'kopie', 'LAST_EXPORT kind=kopie');
  // Überschuss + Kopie mit gleichem Namen → _2
  save(e, basePayload({ diff: -12.5 }));
  ok(names(d).length === 2 && names(d)[1].endsWith('_2'), 'Kopie wird nicht überschrieben', names(d));

  // Mail + Slack
  const e2 = build();
  e2.setCfg('E-Mail-Empfänger', 'a@x.de, b@y.de; kaputt');
  e2.setCfg('Slack-Webhook', 'https://hooks.slack.com/services/T/B/X');
  const r2 = save(e2, basePayload({ diff: 7.3 }));
  ok(r2.ok && e2.log.mails.length === 1 && e2.log.mails[0].to === 'a@x.de,b@y.de', 'Mail an gültige Empfänger', e2.log.mails.map(m => m.to));
  ok(/Differenz 7,30 €/.test(e2.log.mails[0].subject) && /Kopie des Sheets: https:\/\/drive\//.test(e2.log.mails[0].body) && !/Laura|Mitarbeiter/.test(e2.log.mails[0].body), 'Mail-Inhalt (Betrag, Link, kein Mitarbeiter)');
  ok(e2.log.mails[0].body.indexOf('Eingaben:') === -1 && e2.log.mails[0].body.indexOf('EC links') === -1, 'Differenz-Mail ohne Eingabenliste', e2.log.mails[0].body);
  { const bd = e2.log.mails[0].body;
    ok(/Datum: \d\d\.\d\d\.\d{4}/.test(bd) && !/Mitarbeiter|Laura/.test(bd) && /Barbestand SOLL: /.test(bd) && /Barbestand IST: /.test(bd) && /Differenz: -?7,30|Differenz: 7,30/.test(bd.replace(/−/g, '-')), 'Mail zeigt Datum, SOLL, IST, Differenz (ohne Mitarbeiter)', bd);
    ok(!/Unbarer Umsatz|Barentnahme|Tillhub:/.test(bd), 'keine weiteren Zahlen in der Mail', bd);
    ok((bd.match(/Differenz:/g) || []).length === 1, 'Differenz nur einmal in der Mail'); }
  ok(/EC-Belege \(Scan\): keine/.test(e2.log.mails[0].body), 'ohne Beleg: "keine"');
  // mit Beleg: Link in Mail und Slack
  const e2b = build(); e2b.setCfg('E-Mail-Empfänger', 'a@x.de'); e2b.setCfg('Slack-Webhook', 'https://hooks.slack.com/services/T/B/X');
  e2b.ctx.__b = { date: '2026-11-12', pdf: Buffer.concat([Buffer.from('%PDF-1.4\n'), Buffer.alloc(300, 65)]).toString('base64') };
  const br = e2b.run('saveBelege(__b)');
  ok(br.ok && /^Dfile\d+$/.test(br.id) && /^https:\/\/drive\//.test(br.url), 'saveBelege liefert id + url', br);
  save(e2b, basePayload({ diff: -15, belegeId: br.id }));
  ok(e2b.log.mails[0].body.indexOf('EC-Belege (Scan): ' + br.url) > -1, 'Mail enthält Link zum Scan');
  ok(JSON.parse(e2b.log.fetches.find(f => /slack/.test(f.url)).o.payload).text.indexOf(br.url) > -1, 'Slack enthält Link zum Scan');
  // erfundene/ungültige ID wird nicht übernommen
  const e2c = build(); e2c.setCfg('E-Mail-Empfänger', 'a@x.de');
  save(e2c, basePayload({ diff: -15, belegeId: 'https://evil.example/x' }));
  save(e2c, basePayload({ diff: -15, belegeId: 'NICHTVORHANDEN123' }));
  ok(e2c.log.mails.every(m => /EC-Belege \(Scan\): keine/.test(m.body)), 'ungültige/fremde ID → kein Link');
  const sl = e2.log.fetches.find(f => /slack/.test(f.url));
  ok(sl && /Differenz 7,30/.test(JSON.parse(sl.o.payload).text), 'Slack-Nachricht');
  ok(/Meldung gesendet: E-Mail, Slack/.test(r2.message), 'Rückmeldung nennt Kanäle', r2.message);

  // Mail kaputt → trotzdem Kopie, Hinweis
  const e3 = build({ mailFails: true });
  e3.setCfg('E-Mail-Empfänger', 'a@x.de');
  const r3 = save(e3, basePayload({ diff: -1 }));
  ok(r3.ok && r3.kind === 'kopie' && /Meldung fehlgeschlagen/.test(r3.message), 'Mail-Fehler bricht Kopie nicht ab', r3.message);
  // Slack-URL falsch / HTTP-Fehler
  const e4 = build(); e4.setCfg('Slack-Webhook', 'https://evil.example/hook');
  const r4 = save(e4, basePayload({ diff: -1 }));
  ok(r4.ok && /muss mit https:\/\/hooks\.slack\.com/.test(r4.message) && !e4.log.fetches.some(f => /evil/.test(f.url)), 'fremde Webhook-URL wird nicht angefragt');
  const e5 = build({ slackCode: 404 }); e5.setCfg('Slack-Webhook', 'https://hooks.slack.com/services/T/B/X');
  ok(/Slack: HTTP 404/.test(save(e5, basePayload({ diff: -1 })).message), 'Slack HTTP-Fehler gemeldet');
  // Slack-Escaping
  const e6 = build(); e6.setCfg('Slack-Webhook', 'https://hooks.slack.com/services/T/B/X');
  save(e6, basePayload({ name: 'A <!channel> & B', diff: 2 }));
  e6.run("notify_(readConfig_(getSS_()), 'Test <!channel>', 'a & b')");
  const t6 = JSON.parse(e6.log.fetches.filter(f => /slack/.test(f.url)).pop().o.payload).text;
  ok(t6.indexOf('<!channel>') === -1 && t6.indexOf('&lt;!channel&gt;') > -1, 'Slack: < > & escaped');
}

// ============ 4. Toleranz & Schalter ============
console.log('4. Toleranz / Schalter');
{
  const e = build(); e.setCfg('Differenz-Toleranz', '0,50');
  ok(save(e, basePayload({ diff: 0.5 })).kind === 'pdf', 'Differenz 0,50 bei Toleranz 0,50 → PDF');
  ok(save(e, basePayload({ diff: -0.5 })).kind === 'pdf', '−0,50 → PDF');
  ok(save(e, basePayload({ diff: 0.51 })).kind === 'kopie', '0,51 → Kopie');
  const e2 = build(); e2.setCfg('Differenz-Prüfung', 'nein');
  const r = save(e2, basePayload({ diff: undefined }));
  ok(r.ok && r.kind === 'pdf' && r.diff === null, 'Prüfung aus: Differenz nicht nötig → PDF', r);
  ok(e2.proto.cells.K46 === undefined, 'Prüfung aus: kein Differenztext im Blatt', e2.proto.cells.K46);
  { // alter Felder-Reiter mit differenz → K46: Altes wird geleert, nichts neu geschrieben
    const el = build(); const fs_ = el.ss.getSheetByName('Felder');
    const lr = fs_.getLastRow(); fs_.getRange(lr + 1, 1, 1, 4).setValues([['differenz', 'K46', 'alt', 'Text']]);
    el.proto.cells.K46 = 'Differenz laut Tillhub: 9,99 €';
    save(el, basePayload());
    ok(el.proto.cells.K46 === '', 'alter Differenztext wird entfernt', el.proto.cells.K46);
  }
  const e3 = build();
  const r3 = save(e3, basePayload({ diff: undefined }));
  ok(!r3.ok && /Differenz laut Tillhub/.test(r3.message), 'Prüfung an, Differenz fehlt → abgelehnt', r3.message);
  ok(!save(e3, basePayload({ diff: 'abc' })).ok, 'Differenz Text → abgelehnt');
  ok(!save(e3, basePayload({ diff: 1e9 })).ok, 'absurde Differenz → abgelehnt');
  ok(e3.log.fetches.length === 0 && Object.keys(e3.proto.cells).indexOf('G7') === -1, 'abgelehnte Eingaben schreiben nichts ins Blatt');
  const e4 = build(); e4.setCfg('Differenz-Pfad', '');
  save(e4, basePayload({ diff: 3 }));
  ok(names(find(e4.root, 'Abschlüsse_2026/11/Zählprotokolle')).join().indexOf('DIFFERENZ_3,00') > -1, 'Differenz-Pfad leer → Zielordner');
  const e5 = build(); e5.setCfg('Kopie-Name', 'Diff {DIFF} / {TT}.{MM}.');
  save(e5, basePayload({ diff: -2 }));
  ok(names(find(e5.root, 'Abschlüsse_2026/11/Differenzen')).join() === 'Diff -2,00 - 12.11.', 'Kopie-Name: verbotene Zeichen ersetzt', names(find(e5.root, 'Abschlüsse_2026/11/Differenzen')));
}

// ============ 5. Eingabe-Validierung ============
console.log('5. Validierung');
{
  const e = build();
  const bad = (label, mut) => { const p = basePayload(); mut(p); const r = save(e, p); ok(r.ok === false, label + ' → abgelehnt', r.message); };
  bad('negativ', p => p.v.ec_links = -5);
  bad('Anzahl mit Komma', p => p.v.ist_5 = 2.5);
  bad('Text', p => p.v.rechnungen = 'abc');
  bad('Münzbetrag kein Vielfaches (0,30 bei 50 ct)', p => p.v.klein_50 = 0.3);
  bad('Datum falsch', p => p.date = '10.11.2026');
  bad('31.02.', p => p.date = '2026-02-31');
  bad('Name leer', p => p.name = '  ');
  ok(e.log.fetches.length === 0, 'nichts exportiert');
  const r = save(e, basePayload({ name: '=HYPERLINK("x")' }));
  ok(e.proto.cells.K45 === 'HYPERLINK("x")', 'Formel-Präfix im Namen entfernt', e.proto.cells.K45);
  // Blatt weicht von App ab
  const e2 = build();
  const orig = e2.proto.read;
  e2.proto.read = (k) => k === 'K37' ? 999 : orig(k);
  const r2 = save(e2, basePayload());
  ok(!r2.ok && /weicht von der App ab/.test(r2.message) && e2.log.fetches.length === 0, 'Rechenabweichung → nichts gespeichert', r2.message);
  // Sperre
  const e3 = build({ locked: true });
  ok(/anderer Export/.test(save(e3, basePayload()).message), 'Sperre belegt → freundliche Meldung');
}

// ============ 6. Ordner / Jahreswechsel ============
console.log('6. Ordnerstruktur');
{
  const e = build();
  save(e, basePayload({ date: '2026-12-31' }));
  save(e, basePayload({ date: '2027-01-01' }));
  ok(names(find(e.root, 'Abschlüsse_2026/12/Zählprotokolle')).join() === '2026-12-31_Zählprotokoll.pdf', 'Dezember 2026');
  ok(names(find(e.root, 'Abschlüsse_2027/01/Zählprotokolle')).join() === '2027-01-01_Zählprotokoll.pdf', 'Januar 2027 (neuer Jahresordner)');
  const e2 = build(); e2.setCfg('Fehlende Ordner anlegen', 'nein');
  const r = save(e2, basePayload());
  ok(!r.ok && /nicht gefunden in/.test(r.message), 'ohne Anlegen: klare Meldung', r.message);
  // Windows-Schreibweise mit \ und vorangestelltem Root-Namen
  const e3 = build();
  e3.setCfg('Zielpfad', '01_Kasse\\Abschlüsse_{JJJJ}\\{MM}\\Zählprotokolle');
  e3.setCfg('Belege-Pfad', '01_Kasse\\Abschlüsse_{JJJJ}\\{MM}\\EC-Belege');
  const r3 = save(e3, basePayload());
  ok(r3.ok && names(find(e3.root, 'Abschlüsse_2026/11/Zählprotokolle')).join() === '2026-11-12_Zählprotokoll.pdf', 'Backslash-Pfad mit Root-Namen: richtig verschachtelt', r3.message);
  ok(!e3.root.folders.some(f => /\\/.test(f.name)) && e3.root.folders.length === 1 && e3.root.folders[0].name === 'Abschlüsse_2026', 'kein Ordner mit Backslash im Namen, kein doppelter Root-Ordner', e3.root.folders.map(f => f.name));
  const e4 = build(); e4.setCfg('Zielpfad', 'Abschlüsse_{JJJJ}\\{MM}\\Zählprotokolle');
  save(e4, basePayload());
  ok(find(e4.root, 'Abschlüsse_2026/11/Zählprotokolle') && e4.root.folders.length === 1, 'Backslash ohne Root-Namen');
}

// ============ 7. EC-Belege ============
console.log('7. EC-Belege');
{
  const e = build();
  const pdf = Buffer.concat([Buffer.from('%PDF-1.4\n'), Buffer.alloc(300, 65), Buffer.from('\n%%EOF\n')]).toString('base64');
  e.ctx.__b = { date: '2026-11-12', pdf };
  const r = e.run('saveBelege(__b)');
  ok(r.ok && r.name === '2026-11-12_EC-Belege.pdf', 'Beleg-PDF gespeichert', r);
  ok(names(find(e.root, 'Abschlüsse_2026/11/EC-Belege')).join() === '2026-11-12_EC-Belege.pdf', 'im EC-Belege-Ordner');
  e.ctx.__b = { date: '2026-11-12', pdf };
  e.run('saveBelege(__b)');
  ok(names(find(e.root, 'Abschlüsse_2026/11/EC-Belege')).length === 2, 'zweites Mal nummeriert');
  const badCases = [
    ['kein PDF', { date: '2026-11-12', pdf: Buffer.alloc(500, 66).toString('base64') }],
    ['leer', { date: '2026-11-12', pdf: '' }],
    ['Datum falsch', { date: 'x', pdf }],
    ['zu groß', { date: '2026-11-12', pdf: 'A'.repeat(14000001) }],
  ];
  badCases.forEach(([l, p]) => { e.ctx.__b = p; const rr = e.run('saveBelege(__b)'); ok(rr.ok === false, l + ' → abgelehnt', rr.message); });
  e.setCfg('Belege-Dateiname', '../Belege:{TT}');
  e.ctx.__b = { date: '2026-11-12', pdf };
  const r3 = e.run('saveBelege(__b)');
  ok(r3.ok && r3.name === '..-Belege-12.pdf', 'Dateiname bereinigt, .pdf ergänzt', r3.name);
}

// ============ 8. Kein Zugriff / Ausweichordner ============
console.log('8. Kein Drive-Zugriff');
{
  const e = build({ driveDown: true });
  const r = save(e, basePayload());
  ok(!r.ok && /Kein Zugriff|nicht gefunden|Kein Ordner/.test(r.message), 'Fehlermeldung statt Absturz', r.message);
  const e1 = build({ driveDown: true }); e1.setCfg('E-Mail-Empfänger', 'chef@x.de');
  const r1 = save(e1, basePayload());
  ok(!r1.ok && e1.log.mails.length === 1 && /Speichern fehlgeschlagen/.test(e1.log.mails[0].subject) && /Barbestand IST/.test(e1.log.mails[0].body) && !/Laura/.test(e1.log.mails[0].body) && !/EC links/.test(e1.log.mails[0].body),
    'Fehlschlag: Mail mit Wertzusammenfassung', e1.log.mails[0] && e1.log.mails[0].body.slice(0, 200));
  ok(/Meldung gesendet/.test(r1.message), 'Rückmeldung nennt Mail');

  // Root kaputt, Ausweichordner ok
  const e2 = build(); e2.setCfg('Root-Ordner-ID', 'GIBTSNICHT'); e2.setCfg('Ausweich-Ordner-ID', e2.fallback.id);
  const r2 = save(e2, basePayload());
  ok(r2.ok && e2.fallback.files.length === 1 && /Ausweichordner/.test(r2.message), 'Ausweichordner genutzt + Hinweis', r2.message);
  const r2b = save(e2, basePayload({ diff: -3 }));
  ok(r2b.ok && r2b.kind === 'kopie' && e2.fallback.files.length === 2, 'Kopie landet im Ausweichordner');
  e2.ctx.__b = { date: '2026-11-12', pdf: Buffer.concat([Buffer.from('%PDF-1.4\n'), Buffer.alloc(300, 65)]).toString('base64') };
  ok(e2.run('saveBelege(__b)').ok && e2.fallback.files.length === 3, 'Belege ebenfalls im Ausweichordner');
  // beides kaputt
  const e3 = build(); e3.setCfg('Root-Ordner-ID', 'X'); e3.setCfg('Ausweich-Ordner-ID', 'Y');
  ok(/Ausweichordner ebenfalls nicht erreichbar/.test(save(e3, basePayload()).message), 'beide kaputt → beide genannt');
  // Export-HTTP-Fehler → Fehlschlag + Wertsicherung
  const e4 = build({ exportCode: 500 }); e4.setCfg('E-Mail-Empfänger', 'a@x.de');
  const r4 = save(e4, basePayload());
  ok(!r4.ok && /HTTP 500/.test(r4.message) && e4.log.mails.length === 1, 'PDF-Export HTTP 500 → Meldung + Mail');
  // Mehr als eine Seite → Warnung
  const e5 = build({ pages: 2 });
  ok(/2 Seiten statt 1/.test(save(e5, basePayload()).message), 'Mehrseiten-PDF wird gemeldet');

  // checkAccess
  const c1 = build(); ok(c1.run('checkAccess()').level === 'ok', 'checkAccess ok (Zielpfad existiert noch nicht → trotzdem ok)', c1.run('checkAccess()'));
  const c2 = build({ driveDown: true }); ok(c2.run('checkAccess()').level === 'error', 'checkAccess error ohne Zugriff', c2.run('checkAccess()'));
  const c3 = build({ driveDown: true }); c3.setCfg('Ausweich-Ordner-ID', 'F2');
  c3.opts.driveDown = false; c3.setCfg('Root-Ordner-ID', 'nix');
  ok(c3.run('checkAccess()').level === 'warn', 'checkAccess warn mit Ausweichordner', c3.run('checkAccess()'));
  const c4 = build(); c4.root.access = 'VIEW';
  ok(c4.run('checkAccess()').level === 'warn' && /Leserechte/.test(c4.run('checkAccess()').message), 'checkAccess warn bei Leserechten', c4.run('checkAccess()'));
  const c5 = build(); c5.setCfg('Root-Ordner-ID', '');
  ok(c5.run('checkAccess()').level === 'error', 'checkAccess error ohne Root-ID');
  // Leserechte → Speichern scheitert mit Meldung
  const c6 = build(); c6.root.access = 'VIEW';
  ok(!save(c6, basePayload()).ok, 'Leserechte: Speichern liefert Fehler statt Absturz');
}

// ============ 9. Abend-Erinnerung ============
console.log('9. Abend-Erinnerung');
{
  const e = build(); e.setCfg('E-Mail-Empfänger', 'a@x.de');
  e.run('abendCheck()');
  ok(e.log.mails.length === 1 && /noch nicht gespeichert/.test(e.log.mails[0].subject), 'nichts gespeichert → Erinnerung');
  save(e, basePayload({ date: new Date().toISOString().slice(0, 10) }));
  const before = e.log.mails.length;
  e.run('abendCheck()');
  ok(e.log.mails.length === before, 'heute gespeichert → keine Erinnerung');
  const e2 = build(); e2.setCfg('E-Mail-Empfänger', 'a@x.de'); e2.setCfg('Erinnerung-Tage', 'Mo');
  const wd = Number(e2.run("Utilities.formatDate(new Date(), 'Europe/Berlin', 'u')"));
  e2.setCfg('Erinnerung-Tage', ['Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa', 'So'].filter((_, i) => i + 1 !== wd).join(','));
  e2.run('abendCheck()');
  ok(e2.log.mails.length === 0, 'heutiger Wochentag nicht gelistet → keine Erinnerung');
  const e3 = build(); // Trigger-Einrichtung
  e3.run("(function(){ var ui = { alert: function(){ globalThis.__alert = [].slice.call(arguments).join('|'); }, ButtonSet: { OK: 1 } }; SpreadsheetApp.getUi = function(){ return ui; }; })()");
  e3.ctx.ScriptApp.newTrigger = (fn) => { const t = { fn, getHandlerFunction: () => fn, timeBased() { return t; }, atHour(h) { t.h = h; return t; }, nearMinute(m) { t.m = m; return t; }, everyDays() { return t; }, inTimezone(z) { t.tz = z; return t; }, create() { e3.log.triggers.push(t); return t; } }; return t; };
  e3.run('erinnerungEinrichten()');
  ok(e3.log.triggers.length === 0 && /keine gültige Uhrzeit/.test(e3.ctx.__alert), 'ohne Uhrzeit: kein Trigger');
  e3.setCfg('Erinnerung-Uhrzeit', '22:30'); e3.run('erinnerungEinrichten()');
  ok(e3.log.triggers.length === 0 && /Kein Empf/.test(e3.ctx.__alert), 'ohne Empfänger: kein Trigger');
  e3.setCfg('E-Mail-Empfänger', 'a@x.de'); e3.run('erinnerungEinrichten()'); e3.run('erinnerungEinrichten()');
  ok(e3.log.triggers.length === 1 && e3.log.triggers[0].h === 22 && e3.log.triggers[0].m === 30 && e3.log.triggers[0].tz === 'Europe/Berlin', 'genau ein Trigger 22:30 (kein Duplikat)', e3.log.triggers.map(t => [t.h, t.m]));
  ok(e3.run("parseDays_('Mo, Di;Sa so mo')").join() === '1,2,6,7' && e3.run("parseTime_('7.05')").h === 7 && e3.run("parseTime_('25:00')") === null, 'parseDays_/parseTime_');
}

// ============ 10. Web-App ohne aktives Sheet ============
console.log('10. Web-App-Kontext');
{
  const e = build();           // einmal im Sheet → SS_ID gemerkt
  e.opts.noActive = true;
  const init = e.run('getInitData()');
  ok(!!init.today && init.soll.length === 5, 'Web-App findet Tabelle über gemerkte ID');
  const e2 = build(); e2.opts.noActive = true; Object.keys(e2.props).forEach(k => delete e2.props[k]);
  ok(/Tabelle nicht gefunden/.test(String((() => { try { e2.run('getSS_()'); } catch (x) { return x.message; } })() || '')), 'ohne ID klare Meldung');
  ok(e.run("doGet()") !== undefined, 'doGet liefert Seite');
}

console.log('\n' + pass + ' ok, ' + fail + ' Fehler');
process.exit(fail ? 1 : 0);
