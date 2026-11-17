/**
 * Zählprotokoll → PDF → Drive-Ordner (Einstellungen im Reiter "Config", Zellen im Reiter "Felder")
 *
 * Einbinden: Erweiterungen → Apps Script → Code.gs + Index.html einfügen → speichern.
 * Auslösen:
 *  - Web-App (Tablet): Formular mit allen Feldern (doGet + Index.html)
 *    → trägt die Werte ins Blatt ein und speichert das PDF
 *    → geht der Abschluss nicht auf (Differenz laut Tillhub), wird eine Kopie des Sheets gespeichert
 *    → EC-Belege werden fotografiert und als PDF abgelegt
 *  - Desktop: Menü "Zählprotokoll → Als PDF speichern" (exportiert das Blatt, wie es ist)
 *  - Optional: Checkbox in der "Auslöser-Zelle" (Trigger, einmalig einrichten() ausführen)
 *
 * Benötigte Berechtigungen (appsscript.json → oauthScopes):
 *   spreadsheets, drive, script.external_request, script.container.ui,
 *   script.scriptapp (Abend-Erinnerung), script.send_mail (E-Mail), userinfo.email (Rechte-Prüfung)
 */

const CONFIG_SHEET = 'Config';
const FIELDS_SHEET = 'Felder';
const PROP_LAST = 'LAST_EXPORT';

// [Schlüssel, Standardwert, Hinweis]
const DEFAULTS = [
  ['Protokoll-Blatt', '', 'Name des Blatts, das exportiert wird. Leer = erstes Blatt (ohne Config/Felder)'],
  ['Datumszelle', '', 'z. B. K43. Leer = heutiges Datum (nur für Menü/Checkbox; die Web-App nutzt ihr Datumsfeld)'],
  ['Druckbereich', '', 'z. B. A1:K45. Leer = ganzes Blatt'],
  ['Ausrichtung', 'Hochformat', 'Hochformat | Querformat'],
  ['Root-Ordner-ID', '', 'ID des Startordners (aus der Drive-URL nach /folders/). Die Pfade unten beginnen unterhalb dieses Ordners'],
  ['Zielpfad', 'Abschlüsse_{JJJJ}/{MM}/Zählprotokolle', 'Relativ zum Root, Ordner mit / trennen (\\ geht auch). Platzhalter: {JJJJ} {MM} {TT}'],
  ['Zielordner-ID', '', 'Optional: direkte Ordner-ID, überschreibt Root-ID + Zielpfad'],
  ['Dateiname', '{JJJJ}-{MM}-{TT}_Zählprotokoll.pdf', 'Platzhalter: {JJJJ} {MM} {TT}'],
  ['Wenn Datei existiert', 'nummerieren', 'nummerieren (_2, _3 …) | ersetzen (alte Datei → Papierkorb)'],
  ['Fehlende Ordner anlegen', 'ja', 'ja | nein'],
  ['Auslöser-Zelle', '', 'Optional (Checkbox-Trigger): Zelle für die Checkbox, z. B. A47 (außerhalb des Druckbereichs!)'],
  ['Status-Zelle', '', 'Optional: Zelle für die Statusmeldung nach dem Export (außerhalb des Druckbereichs)'],
  ['Letzter Export', '', 'Wird vom Skript gesetzt – nicht ändern'],
  // --- Abschluss geht nicht auf ---
  ['Differenz-Prüfung', 'ja', 'ja | nein – Feld "Differenz laut Tillhub" im Formular. Bei Differenz: Kopie des Sheets statt PDF'],
  ['Differenz-Toleranz', '0', 'Erlaubte Abweichung in € (z. B. 0,50). 0 = jede Differenz zählt'],
  ['Differenz-Pfad', 'Abschlüsse_{JJJJ}/{MM}/Differenzen', 'Ordner für die Kopie (relativ zum Root). Leer = Zielordner'],
  ['Kopie-Name', '{JJJJ}-{MM}-{TT}_Zählprotokoll_DIFFERENZ_{DIFF}', 'Name der Kopie. {DIFF} = Differenz, z. B. -12,50'],
  // --- EC-Belege ---
  ['Belege-Pfad', 'Abschlüsse_{JJJJ}/{MM}/EC-Belege', 'Ordner für das Beleg-PDF (relativ zum Root)'],
  ['Belege-Dateiname', '{JJJJ}-{MM}-{TT}_EC-Belege.pdf', 'Platzhalter: {JJJJ} {MM} {TT}'],
  ['Belege Pflicht', 'nein', 'ja | nein – ja = Speichern nur mit mindestens einem Beleg-Foto'],
  ['Beleg-Zuschnitt', 'ja', 'ja | nein – Beleg im Foto automatisch erkennen, begradigen und zuschneiden (in der App pro Foto umschaltbar)'],
  // --- Zugriff / Meldungen ---
  ['Meldung bei jedem Abschluss', 'ja', 'ja | nein – ja = Mail/Slack bei jedem gespeicherten Abschluss (nicht nur bei Differenz). Empfänger siehe unten'],
  ['Ausweich-Ordner-ID', '', 'Optional: Ordner, in den gespeichert wird, wenn der Zielordner nicht erreichbar ist'],
  ['E-Mail-Empfänger', '', 'Komma-getrennt. Leer = keine Mails. Gilt für Differenz, Fehler und Erinnerung'],
  ['Slack-Webhook', '', 'URL eines Incoming-Webhooks (https://hooks.slack.com/services/…). Leer = kein Slack. Nicht weitergeben!'],
  ['Erinnerung-Uhrzeit', '', 'z. B. 22:30. Leer = keine Erinnerung. Danach im Menü "Abend-Erinnerung einrichten" ausführen'],
  ['Erinnerung-Tage', 'Mo,Di,Mi,Do,Fr,Sa,So', 'Tage, an denen erinnert wird, wenn noch nichts gespeichert wurde'],
];

const DROPDOWNS = {
  'Ausrichtung': ['Hochformat', 'Querformat'],
  'Wenn Datei existiert': ['nummerieren', 'ersetzen'],
  'Fehlende Ordner anlegen': ['ja', 'nein'],
  'Differenz-Prüfung': ['ja', 'nein'],
  'Belege Pflicht': ['ja', 'nein'],
  'Beleg-Zuschnitt': ['ja', 'nein'],
  'Meldung bei jedem Abschluss': ['ja', 'nein'],
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Zählprotokoll')
    .addItem('Als PDF speichern', 'exportZaehlprotokoll')
    .addItem('Config- und Felder-Reiter anlegen/ergänzen', 'createConfigSheet')
    .addItem('Zugriff auf Zielordner prüfen', 'zugriffPruefenMenu')
    .addItem('Abend-Erinnerung einrichten', 'erinnerungEinrichten')
    .addItem('Checkbox-Auslöser einrichten', 'einrichten')
    .addToUi();
}

// ---------- Auslöser ----------

/** Desktop: Menü/Button – exportiert das Blatt, wie es gerade ist */
function exportZaehlprotokoll() {
  const ss = getSS_();
  const r = runExport_(ss);
  report_(ss, r);
  if (r.ok) {
    ss.toast(r.message, 'PDF gespeichert', 8);
  } else {
    SpreadsheetApp.getUi().alert('Export fehlgeschlagen', r.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/** Menü: zeigt das Ergebnis der Zugriffsprüfung */
function zugriffPruefenMenu() {
  const a = checkAccess();
  SpreadsheetApp.getUi().alert('Zugriff auf den Zielordner',
    (a.level === 'ok' ? 'OK. ' : a.level === 'warn' ? 'Hinweis: ' : 'Problem: ') + a.message,
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/** Optional: installierbarer Bearbeiten-Trigger – reagiert auf die Checkbox */
function onEditTrigger(e) {
  if (!e || !e.range) return;
  const ss = e.source;

  let cfg, sheet;
  try {
    cfg = readConfig_(ss);
    sheet = getProtokollSheet_(ss, cfg);
  } catch (err) {
    return;
  }
  const cell = (cfg['Auslöser-Zelle'] || '').replace(/\$/g, '').toUpperCase();
  if (!cell) return;
  if (e.range.getSheet().getSheetId() !== sheet.getSheetId()) return;
  if (e.range.getA1Notation() !== cell) return;
  if (e.range.getValue() !== true) return;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;      // Doppel-Tipp abfangen
  try {
    e.range.setValue(false);            // Checkbox zurücksetzen (löst keinen neuen Trigger aus)
    ss.toast('PDF wird erstellt …', 'Zählprotokoll', 5);
    const r = runExport_(ss);
    report_(ss, r);
    ss.toast(r.message, r.ok ? 'PDF gespeichert' : 'Export fehlgeschlagen', 15);
  } finally {
    lock.releaseLock();
  }
}

/** Optional, einmalig: Checkbox in "Auslöser-Zelle" anlegen + Trigger installieren */
function einrichten() {
  const ss = getSS_();
  const ui = SpreadsheetApp.getUi();
  const cfg = readConfig_(ss);
  const cell = (cfg['Auslöser-Zelle'] || '').replace(/\$/g, '').toUpperCase();
  if (!cell) {
    ui.alert('Bitte zuerst im Reiter Config bei "Auslöser-Zelle" eine Zelle eintragen (z. B. A47, außerhalb des Druckbereichs) und dann erneut ausführen.');
    return;
  }

  const sheet = getProtokollSheet_(ss, cfg);
  sheet.getRange(cell).insertCheckboxes().setValue(false);

  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onEditTrigger') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onEditTrigger').forSpreadsheet(ss).onEdit().create();

  ui.alert('Fertig',
    'Checkbox in ' + sheet.getName() + '!' + cell + ' angelegt und Trigger installiert.\n' +
    'Haken setzen = PDF speichern.', ui.ButtonSet.OK);
}

// ---------- Abend-Erinnerung ----------

const DAY_TOKENS = ['mo', 'di', 'mi', 'do', 'fr', 'sa', 'so'];

function parseDays_(s) {
  const out = [];
  String(s || '').toLowerCase().split(/[\s,;]+/).forEach(function (t) {
    const i = DAY_TOKENS.indexOf(t.slice(0, 2));
    if (i > -1 && out.indexOf(i + 1) === -1) out.push(i + 1);      // 1 = Montag … 7 = Sonntag
  });
  return out;
}

function parseTime_(s) {
  const m = /^(\d{1,2})[:.](\d{2})$/.exec(String(s || '').trim());
  if (!m || +m[1] > 23 || +m[2] > 59) return null;
  return { h: +m[1], m: +m[2] };
}

/** Menü, einmalig: legt den täglichen Erinnerungs-Auslöser an (oder entfernt ihn, wenn keine Uhrzeit gesetzt ist) */
function erinnerungEinrichten() {
  const ss = getSS_();
  const ui = SpreadsheetApp.getUi();
  const cfg = readConfig_(ss);

  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'abendCheck') ScriptApp.deleteTrigger(t);
  });

  const t = parseTime_(cfg['Erinnerung-Uhrzeit']);
  if (!t) {
    ui.alert('Erinnerung aus', 'In Config bei "Erinnerung-Uhrzeit" ist keine gültige Uhrzeit eingetragen (z. B. 22:30). Es gibt keine Erinnerung.', ui.ButtonSet.OK);
    return;
  }
  if (!hasNotifyTarget_(cfg)) {
    ui.alert('Kein Empfänger', 'Trage in Config bei "E-Mail-Empfänger" oder "Slack-Webhook" ein, wohin die Erinnerung gehen soll, und führe das dann erneut aus.', ui.ButtonSet.OK);
    return;
  }
  ScriptApp.newTrigger('abendCheck').timeBased().atHour(t.h).nearMinute(t.m).everyDays(1)
    .inTimezone(ss.getSpreadsheetTimeZone()).create();
  ui.alert('Erinnerung eingerichtet',
    'Täglich gegen ' + cfg['Erinnerung-Uhrzeit'] + ' Uhr (Abweichung bis ca. 15 Minuten) an den Tagen: ' + cfg['Erinnerung-Tage'] + '.', ui.ButtonSet.OK);
}

/** Läuft per Zeit-Auslöser: meldet, wenn heute noch kein Abschluss gespeichert wurde */
function abendCheck() {
  const ss = getSS_();
  const cfg = readConfig_(ss);
  const tz = ss.getSpreadsheetTimeZone();
  const now = new Date();
  const wd = Number(Utilities.formatDate(now, tz, 'u'));
  if (parseDays_(cfg['Erinnerung-Tage']).indexOf(wd) === -1) return;

  const today = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const last = getLastExport_();
  if (last && last.day === today) return;

  const label = Utilities.formatDate(now, tz, 'dd.MM.yyyy');
  notify_(cfg,
    'Zählprotokoll ' + label + ': noch nicht gespeichert',
    'Für heute (' + label + ') wurde noch kein Kassenabschluss gespeichert.\n' +
    'Bitte Zählprotokoll in der App ausfüllen und speichern.\n' +
    (cfg['Letzter Export'] ? 'Letzter Eintrag: ' + cfg['Letzter Export'] : ''));
}

function markExported_(ss, date, kind) {
  try {
    PropertiesService.getScriptProperties().setProperty(PROP_LAST, JSON.stringify({
      day: Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd'),
      kind: kind,
      at: new Date().toISOString(),
    }));
  } catch (e) { /* nur Komfort */ }
}

function getLastExport_() {
  try {
    return JSON.parse(PropertiesService.getScriptProperties().getProperty(PROP_LAST) || 'null');
  } catch (e) {
    return null;
  }
}

// ---------- Meldungen (E-Mail / Slack) ----------

const MAIL_RE = /^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/;
const SLACK_RE = /^https:\/\/hooks\.slack\.com\/services\//;

function parseRecipients_(s) {
  return String(s || '').split(/[\s,;]+/).filter(function (x) { return MAIL_RE.test(x); });
}

function hasNotifyTarget_(cfg) {
  return parseRecipients_(cfg['E-Mail-Empfänger']).length > 0 || SLACK_RE.test(String(cfg['Slack-Webhook'] || ''));
}

function slackEscape_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** Sendet an alle eingetragenen Kanäle; wirft nie. Liefert {sent: [Kanäle], problems: [Texte]} */
function notify_(cfg, subject, body) {
  const out = { sent: [], problems: [] };

  const to = parseRecipients_(cfg['E-Mail-Empfänger']);
  if (to.length) {
    try {
      MailApp.sendEmail({ to: to.join(','), subject: subject, body: body });
      out.sent.push('E-Mail');
    } catch (e) {
      out.problems.push('E-Mail: ' + String(e.message || e));
    }
  }

  const hook = String(cfg['Slack-Webhook'] || '').trim();
  if (hook) {
    if (!SLACK_RE.test(hook)) {
      out.problems.push('Slack: Die Webhook-URL muss mit https://hooks.slack.com/services/ beginnen.');
    } else {
      try {
        const res = UrlFetchApp.fetch(hook, {
          method: 'post',
          contentType: 'application/json',
          payload: JSON.stringify({ text: '*' + slackEscape_(subject) + '*\n' + slackEscape_(body) }),
          muteHttpExceptions: true,
        });
        if (res.getResponseCode() >= 300) out.problems.push('Slack: HTTP ' + res.getResponseCode());
        else out.sent.push('Slack');
      } catch (e) {
        out.problems.push('Slack: ' + String(e.message || e));
      }
    }
  }
  return out;
}

function notifyNote_(n) {
  const lines = [];
  if (n.sent.length) lines.push('Meldung gesendet: ' + n.sent.join(', ') + '.');
  n.problems.forEach(function (p) { lines.push('⚠ Meldung fehlgeschlagen – ' + p); });
  return lines.join('\n');
}

// ---------- Zahlen- und Namenshelfer ----------

function num_(s, fallback) {
  const t = String(s === undefined || s === null ? '' : s).trim().replace(',', '.');
  if (t === '') return fallback;
  const x = Number(t);
  return isFinite(x) ? x : fallback;
}

/** -12.5 → "-12,50" */
function fmtNum_(x) {
  return (Math.round(x * 100) / 100).toFixed(2).replace('.', ',');
}

/** 1003.83 → "1.003,83 €" */
function fmtEur_(x) {
  const s = Math.abs(Math.round(x * 100) / 100).toFixed(2).split('.');
  return (x < 0 ? '-' : '') + s[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.') + ',' + s[1] + ' €';
}

function safeName_(s) {
  return String(s).replace(/[\\\/:*?"<>|\u0000-\u001f]/g, '-').trim().slice(0, 180);
}

function uniqueName_(folder, name) {
  if (!folder.getFilesByName(name).hasNext()) return name;
  let n = 2;
  while (folder.getFilesByName(name + '_' + n).hasNext()) n++;
  return name + '_' + n;
}

function diffCheckOn_(cfg) {
  return String(cfg['Differenz-Prüfung']).toLowerCase() !== 'nein';
}

// ---------- Web-App: Formular → Blatt → PDF / Kopie ----------

const BILLS = [5, 10, 20, 50, 100];
const COIN_CENTS = [200, 100, 50, 20, 10, 5, 2, 1];

/** Liefert die Seite mit dem Formular */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Zählprotokoll')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Startdaten für das Formular */
function getInitData() {
  const ss = getSS_();
  const cfg = readConfig_(ss);
  const sheet = getProtokollSheet_(ss, cfg);
  const cells = readFieldCells_(ss);
  return {
    today: Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd'),
    soll: BILLS.map(function (b) { return readNum_(sheet, cells['soll_' + b], null); }),
    tresor: readNum_(sheet, cells['tresor'], 337),
    last: cfg['Letzter Export'] || '',
    diffCheck: diffCheckOn_(cfg),
    tol: num_(cfg['Differenz-Toleranz'], 0),
    notify: hasNotifyTarget_(cfg),
    belegePflicht: String(cfg['Belege Pflicht']).toLowerCase() === 'ja',
    belegeCrop: String(cfg['Beleg-Zuschnitt']).toLowerCase() !== 'nein',
  };
}

/** Speichern-Button: Werte ins Blatt schreiben, gegenrechnen, PDF (oder bei Differenz eine Kopie) ablegen */
function saveProtokoll(d) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { ok: false, message: 'Ein anderer Export läuft gerade – bitte kurz warten und erneut versuchen.' };
  }
  let ss = null, cfg = null, summary = null;
  try {
    ss = getSS_();
    cfg = readConfig_(ss);
    const sheet = getProtokollSheet_(ss, cfg);
    const cells = readFieldCells_(ss);

    const clean = cleanInput_(d, cfg);
    const tresor = readNum_(sheet, cells['tresor'], 337);
    const totals = calcTotals_(clean.v, tresor);
    const beleg = belegLink_(d.belegeId);
    summary = { clean: clean, totals: totals, beleg: beleg };      // ab hier gelten Fehler als "Speichern fehlgeschlagen"

    writeInputs_(sheet, cells, clean);
    SpreadsheetApp.flush();

    const problems = verifyTotals_(sheet, cells, totals);
    if (problems.length) {
      const msg = 'Die Berechnung im Blatt weicht von der App ab – nichts wurde gespeichert.\n' +
        problems.slice(0, 4).join('\n') +
        '\nBitte Reiter "Felder" (Zellen-Zuordnung) prüfen.';
      report_(ss, { ok: false, message: msg.split('\n')[0] });
      return { ok: false, message: msg };
    }

    const tol = num_(cfg['Differenz-Toleranz'], 0);
    const mismatch = clean.diff !== null && Math.abs(clean.diff) > tol + 0.0049;
    const r = mismatch ? exportCopy_(ss, cfg, clean) : runExport_(ss, clean.date);
    report_(ss, r);

    let note = '';
    if (r.ok && r.kind === 'kopie') {
      note = notifyNote_(notify_(cfg,
        'Kassenabschluss ' + dateLabel_(ss, clean.date) + ': Differenz ' + fmtEur_(clean.diff),
        'Der Abschluss geht nicht auf. Es wurde kein PDF gespeichert, sondern eine Kopie des Sheets.\n\n' +
        summaryText_(ss, clean, totals) + '\n\n' +
        'Kopie des Sheets: ' + r.url + '\n' +
        'EC-Belege (Scan): ' + (beleg || 'keine')));
    } else if (r.ok && String(cfg['Meldung bei jedem Abschluss']).toLowerCase() !== 'nein') {
      note = notifyNote_(notify_(cfg,
        'Kassenabschluss ' + dateLabel_(ss, clean.date) + ': gespeichert' +
          (clean.diff !== null && Math.abs(clean.diff) > 0.0049 ? ' (Differenz ' + fmtEur_(clean.diff) + ', innerhalb der Toleranz)' : ''),
        'Der Kassenabschluss wurde als PDF gespeichert.\n\n' +
        summaryText_(ss, clean, totals) + '\n\n' +
        'PDF: ' + r.url + '\n' +
        'EC-Belege (Scan): ' + (beleg || 'keine')));
    } else if (!r.ok) {
      note = failureNote_(ss, cfg, summary, r.message);
    }
    return {
      ok: r.ok, kind: r.kind, url: r.url, totals: totals, diff: clean.diff,
      message: r.message + (note ? '\n' + note : ''),
    };
  } catch (e) {
    const msg = String(e.message || e);
    const note = summary ? failureNote_(ss, cfg, summary, msg) : '';
    return { ok: false, message: msg + (note ? '\n' + note : '') };
  } finally {
    lock.releaseLock();
  }
}

/** Schickt bei einem Speicherfehler die Werte per Mail/Slack, damit nichts verloren geht */
function failureNote_(ss, cfg, summary, errorText) {
  try {
    const n = notify_(cfg,
      'Zählprotokoll ' + dateLabel_(ss, summary.clean.date) + ': Speichern fehlgeschlagen',
      'Das Zählprotokoll konnte nicht gespeichert werden.\nFehler: ' + errorText + '\n\n' +
      'Die Eingaben liegen noch in der App; dort erneut speichern.\n\n' +
      summaryText_(ss, summary.clean, summary.totals) +
      (summary.beleg ? '\n\nEC-Belege (Scan): ' + summary.beleg : ''));
    return notifyNote_(n);
  } catch (e) {
    return '';
  }
}

function dateLabel_(ss, date) {
  return Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), 'dd.MM.yyyy');
}

/** Alle Eingaben als Text (für Mails, Slack und als Sicherung) */
function summaryText_(ss, clean, totals) {
  return [
    'Datum: ' + dateLabel_(ss, clean.date),
    'Barbestand SOLL: ' + fmtEur_(totals.soll),
    'Barbestand IST: ' + fmtEur_(totals.ist),
    'Differenz: ' + (clean.diff !== null ? fmtEur_(clean.diff) : 'nicht angegeben'),
  ].join('\n');
}

/** Link zum gespeicherten Beleg-PDF (die App schickt nur die Datei-ID; die URL kommt aus Drive) */
function belegLink_(id) {
  if (!id || !/^[A-Za-z0-9_-]{10,}$/.test(String(id))) return '';
  try {
    return DriveApp.getFileById(String(id)).getUrl();
  } catch (e) {
    return '';
  }
}

/** Eingaben prüfen und bereinigen (der Client wird nicht blind vertraut) */
function cleanInput_(d, cfg) {
  if (!d || typeof d !== 'object') throw new Error('Keine Daten erhalten.');

  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(d.date || ''));
  if (!m) throw new Error('Datum fehlt oder ist ungültig.');
  const date = new Date(+m[1], +m[2] - 1, +m[3], 12);
  if (date.getMonth() !== +m[2] - 1) throw new Error('Datum ist ungültig.');

  const name = String(d.name || '').replace(/^[=+\-@\s]+/, '').trim().slice(0, 60);
  if (!name) throw new Error('Bitte den Mitarbeiter eintragen.');

  const src = d.v || {};
  const v = {};
  fieldDefs_().forEach(function (def) {
    const key = def[0];
    if (def[3] !== 'Eingabe' || key === 'datum' || key === 'mitarbeiter') return;
    const isCount = /^(ist|ent[12])_\d+$/.test(key);
    let x = Number(src[key] === undefined || src[key] === null || src[key] === '' ? 0 : src[key]);
    if (!isFinite(x) || x < 0) throw new Error('Ungültiger Wert bei "' + def[2] + '".');
    if (isCount) {
      if (x !== Math.floor(x) || x > 10000) throw new Error('Ungültige Anzahl bei "' + def[2] + '".');
    } else {
      if (x > 1000000) throw new Error('Ungültiger Betrag bei "' + def[2] + '".');
      x = Math.round(x * 100) / 100;
      const coin = /^klein_(\d+)$/.exec(key);
      if (coin && Math.round(x * 100) % Number(coin[1]) !== 0) {
        throw new Error('Kleingeld ' + coinLabel_(Number(coin[1])) + ': ' + x + ' € ist kein Vielfaches der Münze.');
      }
    }
    v[key] = x;
  });

  let diff = null;
  if (cfg && diffCheckOn_(cfg)) {
    if (d.diff === undefined || d.diff === null || d.diff === '') {
      throw new Error('Bitte die Differenz laut Tillhub angeben (0, wenn der Abschluss aufgeht).');
    }
    const x = Number(d.diff);
    if (!isFinite(x) || Math.abs(x) > 1000000) throw new Error('Ungültige Differenz laut Tillhub.');
    diff = Math.round(x * 100) / 100;
  }
  return { date: date, name: name, v: v, diff: diff };
}

/** Gleiche Rechenlogik wie die Formeln im Blatt (J7, J11, J17, J20, J29, J34, J37, J39, J41) */
function calcTotals_(v, tresor) {
  const r2 = function (x) { return Math.round(x * 100) / 100; };
  const bills = function (p) {
    return BILLS.reduce(function (s, b) { return s + b * v[p + '_' + b]; }, 0);
  };
  const sumEC = v.ec_links + v.ec_rechts;
  const unbar = sumEC + v.rechnungen;
  const scheine = bills('ist');
  const klein = COIN_CENTS.reduce(function (s, c) { return s + v['klein_' + c]; }, 0);
  const ent1 = bills('ent1');
  const ent2 = 0;                                     // Barentnahme 2 entfällt
  const entnahme = ent1;
  const soll = scheine + klein + tresor + entnahme;   // J17+J20+J24+J29+J30+J34+J35
  const ist = soll - entnahme;                        // J37 − J39
  return {
    sumEC: r2(sumEC), unbar: r2(unbar), scheine: r2(scheine), klein: r2(klein),
    ent1: r2(ent1), ent2: r2(ent2), soll: r2(soll), entnahme: r2(entnahme), ist: r2(ist),
  };
}

const CHECKS = [
  ['pruef_ec', 'sumEC'], ['pruef_unbar', 'unbar'], ['pruef_scheine', 'scheine'],
  ['pruef_klein', 'klein'], ['pruef_ent1', 'ent1'], ['pruef_ent2', 'ent2'],
  ['pruef_soll', 'soll'], ['pruef_entnahme', 'entnahme'], ['pruef_ist', 'ist'],
];

/** Vergleicht die vom Blatt berechneten Summen mit der App-Berechnung */
function verifyTotals_(sheet, cells, totals) {
  const problems = [];
  CHECKS.forEach(function (c) {
    const cell = cells[c[0]];
    if (!cell) return;
    const raw = sheet.getRange(cell).getValue();
    const x = Number(raw);
    if (raw === '' || !isFinite(x) || Math.abs(x - totals[c[1]]) > 0.005) {
      problems.push(cell + ': Blatt ' + (raw === '' ? 'leer' : raw) + ' ≠ App ' + totals[c[1]]);
    }
  });
  return problems;
}

function writeInputs_(sheet, cells, clean) {
  Object.keys(clean.v).forEach(function (k) {
    if (cells[k]) sheet.getRange(cells[k]).setValue(clean.v[k]);
  });
  fieldDefs_().forEach(function (def) {              // nicht mehr genutzte Felder (Barentnahme 2) zurücksetzen
    if (def[3] === 'Nullen' && cells[def[0]]) sheet.getRange(cells[def[0]]).setValue(0);
  });
  if (cells['datum']) sheet.getRange(cells['datum']).setValue(clean.date);
  if (cells['mitarbeiter']) sheet.getRange(cells['mitarbeiter']).setValue(clean.name);
  if (cells['differenz']) {                          // alter Eintrag im Reiter Felder: Text nicht mehr schreiben, Altes entfernen
    try { sheet.getRange(cells['differenz']).clearContent(); } catch (e) { /* nur Komfort */ }
  }
}

function readNum_(sheet, cell, fallback) {
  if (!cell) return fallback;
  const x = Number(sheet.getRange(cell).getValue());
  return isFinite(x) ? x : fallback;
}

// ---------- EC-Belege (PDF aus der App) ----------

/** Nimmt das in der App erzeugte Beleg-PDF (Base64) entgegen und legt es ab */
function saveBelege(p) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { ok: false, message: 'Ein anderer Vorgang läuft gerade – bitte kurz warten und erneut versuchen.' };
  }
  try {
    const ss = getSS_();
    const cfg = readConfig_(ss);

    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(p && p.date || ''));
    if (!m) throw new Error('Datum fehlt oder ist ungültig.');
    const date = new Date(+m[1], +m[2] - 1, +m[3], 12);
    if (date.getMonth() !== +m[2] - 1) throw new Error('Datum ist ungültig.');

    const b64 = String(p.pdf || '');
    if (!b64) throw new Error('Keine Beleg-Daten erhalten.');
    if (b64.length > 14000000) throw new Error('Die Belege sind zu groß (über 10 MB). Bitte weniger Fotos auf einmal senden.');
    const bytes = Utilities.base64Decode(b64);
    if (bytes.length < 100 || bytes[0] !== 37 || bytes[1] !== 80 || bytes[2] !== 68 || bytes[3] !== 70) {
      throw new Error('Die Beleg-Datei ist kein gültiges PDF.');
    }

    let name = safeName_(fill_(cfg['Belege-Dateiname'], ss, date));
    if (!/\.pdf$/i.test(name)) name += '.pdf';
    const blob = Utilities.newBlob(bytes, 'application/pdf', name);

    const tgt = getTargetFolder_(cfg, ss, date, 'belege');
    const file = saveFile_(tgt.folder, blob, cfg);
    return {
      ok: true,
      name: file.getName(),
      id: file.getId(),
      url: file.getUrl(),
      message: file.getName() + ' → ' + tgt.folder.getName() + (tgt.note ? '\n' + tgt.note : ''),
    };
  } catch (e) {
    return { ok: false, message: String(e.message || e) };
  } finally {
    lock.releaseLock();
  }
}

// ---------- Felder-Zuordnung (Reiter "Felder") ----------

function coinLabel_(c) {
  return (c / 100).toFixed(2).replace('.', ',') + ' €';
}

/** [Schlüssel, Standardzelle, Beschreibung, Art] – Standardzellen für das Blatt mit Wertespalte K */
function fieldDefs_() {
  const d = [];
  d.push(['ec_links', 'G7', 'EC links (2481), Betrag in €', 'Eingabe']);
  d.push(['ec_rechts', 'I7', 'EC rechts (2325), Betrag in €', 'Eingabe']);
  d.push(['rechnungen', 'K9', 'Rechnungen, Betrag in €', 'Eingabe']);
  BILLS.forEach(function (b, i) {
    d.push(['ist_' + b, 'BCDEF'.charAt(i) + '16', 'Scheine in Kasse (Ist), Anzahl ' + b + ' €', 'Eingabe']);
  });
  COIN_CENTS.forEach(function (c, i) {
    d.push(['klein_' + c, 'BCDEFGHI'.charAt(i) + '20', 'Kleingeld ' + coinLabel_(c) + ', Summe in €', 'Eingabe']);
  });
  BILLS.forEach(function (b, i) {
    d.push(['ent1_' + b, 'BCDEF'.charAt(i) + '28', 'Barentnahme 1, Anzahl ' + b + ' €', 'Eingabe']);
  });
  d.push(['ent1_wert', 'K30', 'Barentnahme 1, "oder Wert" (wird auf 0 gesetzt)', 'Nullen']);
  // Barentnahme 2 gibt es in der App nicht mehr; die Zellen im Blatt werden bei jedem Speichern auf 0 gesetzt
  BILLS.forEach(function (b, i) {
    d.push(['ent2_' + b, 'BCDEF'.charAt(i) + '33', 'Barentnahme 2, Anzahl ' + b + ' € (wird auf 0 gesetzt)', 'Nullen']);
  });
  d.push(['ent2_wert', 'K35', 'Barentnahme 2, "oder Wert" (wird auf 0 gesetzt)', 'Nullen']);
  d.push(['datum', 'K43', 'Datum (überschreibt die Formel HEUTE)', 'Eingabe']);
  d.push(['mitarbeiter', 'K45', 'Mitarbeiter Theke', 'Eingabe']);
  BILLS.forEach(function (b, i) {
    d.push(['soll_' + b, 'BCDEF'.charAt(i) + '15', 'Scheine Soll, Anzahl ' + b + ' € (nur lesen)', 'Lesen']);
  });
  d.push(['tresor', 'K24', 'Tresor-Sollbetrag (nur lesen)', 'Lesen']);
  d.push(['pruef_ec', 'K7', 'Summe EC', 'Prüfen']);
  d.push(['pruef_unbar', 'K11', 'Unbarer Umsatz', 'Prüfen']);
  d.push(['pruef_scheine', 'K17', 'Summe Scheine in Kasse', 'Prüfen']);
  d.push(['pruef_klein', 'K20', 'Summe Kleingeld', 'Prüfen']);
  d.push(['pruef_ent1', 'K29', 'Barentnahme 1, Summe Scheine', 'Prüfen']);
  d.push(['pruef_ent2', 'K34', 'Barentnahme 2, Summe Scheine (muss 0 sein)', 'Prüfen']);
  d.push(['pruef_soll', 'K37', 'Barbestand Soll (für Tillhub)', 'Prüfen']);
  d.push(['pruef_entnahme', 'K39', 'Barentnahme gesamt (für Tillhub)', 'Prüfen']);
  d.push(['pruef_ist', 'K41', 'Barbestand IST', 'Prüfen']);
  return d;
}

function ensureFields_(ss) {
  const defs = fieldDefs_();
  let sh = ss.getSheetByName(FIELDS_SHEET);

  if (!sh) {
    sh = ss.insertSheet(FIELDS_SHEET);
    const rows = [['Schlüssel', 'Zelle im Protokoll-Blatt', 'Beschreibung', 'Art']].concat(defs);
    sh.getRange(1, 1, rows.length, 4).setValues(rows);
    sh.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#e8eaed');
    sh.getRange(2, 1, defs.length, 1).setFontColor('#5f6368');
    sh.getRange(2, 2, defs.length, 1).setNumberFormat('@').setFontWeight('bold').setBackground('#eef6e8');
    sh.getRange(2, 3, defs.length, 2).setFontColor('#5f6368');
    sh.setColumnWidth(1, 150);
    sh.setColumnWidth(2, 190);
    sh.setColumnWidth(3, 380);
    sh.setColumnWidth(4, 80);
    sh.setFrozenRows(1);
    return sh;
  }

  // fehlende Schlüssel ergänzen (bestehende Zellen bleiben)
  const have = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(function (r) { return String(r[0]).trim(); });
  const missing = defs.filter(function (d) { return have.indexOf(d[0]) === -1; });
  if (missing.length) {
    const start = sh.getLastRow() + 1;
    sh.getRange(start, 1, missing.length, 4).setValues(missing);
    sh.getRange(start, 2, missing.length, 1).setNumberFormat('@');
  }
  return sh;
}

function readFieldCells_(ss) {
  const sh = ensureFields_(ss);
  const map = {};
  fieldDefs_().forEach(function (d) { map[d[0]] = d[1]; });   // Fallback
  sh.getDataRange().getValues().slice(1).forEach(function (r) {
    const key = String(r[0]).trim();
    if (key && (key in map || key === 'differenz')) map[key] = String(r[1]).replace(/\$/g, '').trim().toUpperCase();
  });
  return map;
}

// ---------- Kern ----------

/** Führt den PDF-Export aus, wirft nie – liefert {ok, kind, url, message}. dateOverride: Datum aus der Web-App */
function runExport_(ss, dateOverride) {
  try {
    const cfg = readConfig_(ss);
    const sheet = getProtokollSheet_(ss, cfg);
    const date = dateOverride || getDatum_(ss, sheet, cfg);
    const tgt = getTargetFolder_(cfg, ss, date, 'main');
    const name = fill_(cfg['Dateiname'], ss, date);

    SpreadsheetApp.flush();
    const blob = exportPdf_(ss, sheet, cfg).setName(name);
    const pages = countPdfPages_(blob);
    const file = saveFile_(tgt.folder, blob, cfg);
    markExported_(ss, date, 'pdf');

    return {
      ok: true,
      kind: 'pdf',
      url: file.getUrl(),
      message: file.getName() + ' → ' + tgt.folder.getName() +
        (pages > 1 ? '  ⚠ ' + pages + ' Seiten statt 1!' : '') +
        (tgt.note ? '\n⚠ ' + tgt.note : ''),
    };
  } catch (e) {
    return { ok: false, kind: 'pdf', message: String(e.message || e) };
  }
}

/** Abschluss geht nicht auf: Kopie des Sheets statt PDF. Wirft nie. */
function exportCopy_(ss, cfg, clean) {
  try {
    const tgt = getTargetFolder_(cfg, ss, clean.date, 'diff');
    const base = safeName_(
      fill_(cfg['Kopie-Name'] || '{JJJJ}-{MM}-{TT}_Zählprotokoll_DIFFERENZ_{DIFF}', ss, clean.date)
        .replace(/\{DIFF\}/g, fmtNum_(clean.diff)));
    const name = uniqueName_(tgt.folder, base);

    SpreadsheetApp.flush();
    const file = DriveApp.getFileById(ss.getId()).makeCopy(name, tgt.folder);
    markExported_(ss, clean.date, 'kopie');

    return {
      ok: true,
      kind: 'kopie',
      url: file.getUrl(),
      message: 'Differenz ' + fmtEur_(clean.diff) + ' – es wurde KEIN PDF gespeichert, sondern eine Kopie des Sheets:\n' +
        file.getName() + ' → ' + tgt.folder.getName() + (tgt.note ? '\n⚠ ' + tgt.note : ''),
    };
  } catch (e) {
    return { ok: false, kind: 'kopie', message: String(e.message || e) };
  }
}

/** Status in Config ("Letzter Export") und optional in "Status-Zelle" festhalten */
function report_(ss, r) {
  try {
    const stamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'dd.MM.yyyy HH:mm');
    const text = stamp + ' – ' + (r.ok ? (r.kind === 'kopie' ? '⚠ ' : '✔ ') : '✖ ') + String(r.message).replace(/\s*\n\s*/g, ' | ');
    setConfigValue_(ss, 'Letzter Export', text);

    const cfg = readConfig_(ss);
    const statusCell = (cfg['Status-Zelle'] || '').replace(/\$/g, '');
    if (statusCell) getProtokollSheet_(ss, cfg).getRange(statusCell).setValue(text);
  } catch (e) {
    // Statusmeldung ist nur Komfort – Fehler hier ignorieren
  }
}

/** Tabelle ermitteln: im Sheet-Kontext direkt, in der Web-App über die gemerkte ID */
function getSS_() {
  const props = PropertiesService.getScriptProperties();
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    props.setProperty('SS_ID', active.getId());
    return active;
  }
  const id = props.getProperty('SS_ID');
  if (!id) {
    throw new Error('Tabelle nicht gefunden. Bitte einmal im Sheet das Menü "Zählprotokoll → Als PDF speichern" ausführen.');
  }
  return SpreadsheetApp.openById(id);
}

// ---------- Config ----------

/** Config- und Felder-Reiter anlegen bzw. um fehlende Zeilen ergänzen */
function createConfigSheet() {
  const ss = getSS_();
  ensureConfig_(ss);
  ensureFields_(ss);
}

function ensureConfig_(ss) {
  let sh = ss.getSheetByName(CONFIG_SHEET);
  if (sh) {
    addMissingKeys_(sh);
    return sh;
  }

  sh = ss.insertSheet(CONFIG_SHEET);
  const rows = [['Schlüssel', 'Wert', 'Hinweis']].concat(DEFAULTS);
  sh.getRange(1, 1, rows.length, 3).setValues(rows);
  sh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#e8eaed');
  sh.getRange(2, 1, rows.length - 1, 1).setFontWeight('bold');
  sh.getRange(2, 3, rows.length - 1, 1).setFontColor('#5f6368');
  sh.getRange(2, 2, rows.length - 1, 1).setNumberFormat('@'); // alles als Text
  sh.setColumnWidth(1, 200);
  sh.setColumnWidth(2, 420);
  sh.setColumnWidth(3, 460);
  sh.setFrozenRows(1);

  DEFAULTS.forEach(function (d, i) {
    const opts = DROPDOWNS[d[0]];
    if (opts) {
      sh.getRange(i + 2, 2).setDataValidation(
        SpreadsheetApp.newDataValidation().requireValueInList(opts, true).build()
      );
    }
  });
  return sh;
}

/** Ergänzt Schlüssel, die in einem älteren Config-Reiter noch fehlen (bestehende Werte bleiben) */
function addMissingKeys_(sh) {
  const have = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(function (r) { return String(r[0]).trim(); });
  const missing = DEFAULTS.filter(function (d) { return have.indexOf(d[0]) === -1; });
  if (!missing.length) return;

  const start = sh.getLastRow() + 1;
  sh.getRange(start, 1, missing.length, 3).setValues(missing);
  sh.getRange(start, 1, missing.length, 1).setFontWeight('bold');
  sh.getRange(start, 3, missing.length, 1).setFontColor('#5f6368');
  sh.getRange(start, 2, missing.length, 1).setNumberFormat('@');
  missing.forEach(function (d, i) {
    const opts = DROPDOWNS[d[0]];
    if (opts) {
      sh.getRange(start + i, 2).setDataValidation(
        SpreadsheetApp.newDataValidation().requireValueInList(opts, true).build()
      );
    }
  });
}

function readConfig_(ss) {
  const sh = ensureConfig_(ss);
  const cfg = {};
  DEFAULTS.forEach(function (d) { cfg[d[0]] = d[1]; }); // Fallback
  sh.getDataRange().getValues().slice(1).forEach(function (r) {
    const key = String(r[0]).trim();
    if (key && key in cfg) cfg[key] = String(r[1]).trim();
  });
  return cfg;
}

function setConfigValue_(ss, key, value) {
  const sh = ensureConfig_(ss);
  const keys = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (let i = 0; i < keys.length; i++) {
    if (String(keys[i][0]).trim() === key) {
      sh.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
}

// ---------- Bausteine ----------

function getProtokollSheet_(ss, cfg) {
  const wanted = cfg['Protokoll-Blatt'];
  if (wanted) {
    const sh = ss.getSheetByName(wanted);
    if (!sh) throw new Error('Blatt "' + wanted + '" nicht gefunden (Config → Protokoll-Blatt).');
    return sh;
  }
  const first = ss.getSheets().filter(function (s) {
    return s.getName() !== CONFIG_SHEET && s.getName() !== FIELDS_SHEET;
  })[0];
  if (!first) throw new Error('Kein Protokoll-Blatt gefunden.');
  return first;
}

function getDatum_(ss, sheet, cfg) {
  const cell = cfg['Datumszelle'];
  if (!cell) return new Date();

  const v = sheet.getRange(cell).getValue();
  if (v instanceof Date) return v;

  const s = String(v).trim();
  let m = s.match(/(\d{1,2})\.(\d{1,2})\.(\d{4})/);   // 10.11.2026
  if (m) return new Date(+m[3], +m[2] - 1, +m[1], 12);
  m = s.match(/(\d{4})-(\d{2})-(\d{2})/);             // 2026-11-10
  if (m) return new Date(+m[1], +m[2] - 1, +m[3], 12);

  throw new Error('Kein gültiges Datum in ' + sheet.getName() + '!' + cell + ' (Wert: "' + s + '").');
}

function fill_(text, ss, date) {
  const tz = ss.getSpreadsheetTimeZone();
  return text
    .replace(/\{JJJJ\}/g, Utilities.formatDate(date, tz, 'yyyy'))
    .replace(/\{MM\}/g, Utilities.formatDate(date, tz, 'MM'))
    .replace(/\{TT\}/g, Utilities.formatDate(date, tz, 'dd'));
}

function getFolder_(id, label) {
  try {
    return DriveApp.getFolderById(id);
  } catch (e) {
    throw new Error(
      'Config "' + label + '": Kein Ordner mit der ID "' + id + '" gefunden oder kein Zugriff.\n' +
      'ID direkt aus der Drive-URL (nach /folders/) kopieren und Schreibrechte prüfen.'
    );
  }
}

const PATH_KEYS = { main: 'Zielpfad', diff: 'Differenz-Pfad', belege: 'Belege-Pfad' };

/** Pfad in Ordnernamen zerlegen: „/“ und „\“ (Windows) gelten beide als Trenner */
function pathParts_(tpl, ss, date) {
  return fill_(tpl, ss, date).split(/[\/\\]/).map(function (p) { return p.trim(); }).filter(Boolean);
}

/** Pfadteile relativ zum Root. Beginnt der Pfad mit dem Namen des Root-Ordners selbst (z. B. aus dem Explorer kopiert), wird dieser Teil übersprungen. */
function relParts_(root, tpl, ss, date) {
  const parts = pathParts_(tpl, ss, date);
  if (parts.length && parts[0] === root.getName() && !root.getFoldersByName(parts[0]).hasNext()) parts.shift();
  return parts;
}

/** Löst den Ordner auf (legt fehlende Ordner an, wenn erlaubt). kind: main | diff | belege */
function resolveFolder_(cfg, ss, date, kind) {
  kind = kind || 'main';
  if (kind === 'main' && cfg['Zielordner-ID']) return getFolder_(cfg['Zielordner-ID'], 'Zielordner-ID');

  const key = PATH_KEYS[kind];
  if (!cfg['Root-Ordner-ID']) {
    throw new Error(kind === 'main'
      ? 'Weder Zielordner-ID noch Root-Ordner-ID in der Config.'
      : 'Root-Ordner-ID fehlt in der Config (wird für "' + key + '" gebraucht).');
  }
  let current = getFolder_(cfg['Root-Ordner-ID'], 'Root-Ordner-ID');
  const create = String(cfg['Fehlende Ordner anlegen']).toLowerCase() === 'ja';

  relParts_(current, cfg[key], ss, date).forEach(function (part) {
    const it = current.getFoldersByName(part);
    if (it.hasNext()) {
      current = it.next();
    } else if (create) {
      current = current.createFolder(part);
    } else {
      throw new Error('Ordner "' + part + '" nicht gefunden in "' + current.getName() + '".');
    }
  });
  return current;
}

/**
 * Liefert {folder, note}. Ist der Zielordner nicht erreichbar und ein Ausweichordner eingetragen,
 * wird dort gespeichert und note erklärt das.
 */
function getTargetFolder_(cfg, ss, date, kind) {
  if (kind === 'diff' && (!cfg['Differenz-Pfad'] || !cfg['Root-Ordner-ID'])) kind = 'main';   // leer = Zielordner
  try {
    return { folder: resolveFolder_(cfg, ss, date, kind), note: '' };
  } catch (e) {
    const fb = cfg['Ausweich-Ordner-ID'];
    if (!fb) throw e;
    let folder;
    try {
      folder = getFolder_(fb, 'Ausweich-Ordner-ID');
    } catch (e2) {
      throw new Error(String(e.message || e) + '\nAusweichordner ebenfalls nicht erreichbar: ' + String(e2.message || e2));
    }
    return {
      folder: folder,
      note: 'Zielordner nicht erreichbar (' + String(e.message || e).split('\n')[0] + ') – im Ausweichordner "' +
        folder.getName() + '" gespeichert. Bitte später in den richtigen Ordner verschieben.',
    };
  }
}

// ---------- Zugriffsprüfung ----------

/** Findet den tiefsten vorhandenen Ordner des Zielpfads, ohne etwas anzulegen */
function probeTarget_(cfg, ss, date) {
  if (cfg['Zielordner-ID']) return { folder: getFolder_(cfg['Zielordner-ID'], 'Zielordner-ID'), missing: [] };
  if (!cfg['Root-Ordner-ID']) throw new Error('Weder Zielordner-ID noch Root-Ordner-ID in der Config.');
  let current = getFolder_(cfg['Root-Ordner-ID'], 'Root-Ordner-ID');
  const parts = relParts_(current, cfg['Zielpfad'], ss, date);
  for (let i = 0; i < parts.length; i++) {
    const it = current.getFoldersByName(parts[i]);
    if (!it.hasNext()) return { folder: current, missing: parts.slice(i) };
    current = it.next();
  }
  return { folder: current, missing: [] };
}

/** true = darf schreiben, false = nur lesen, null = nicht feststellbar */
function canWrite_(folder) {
  try {
    const p = folder.getAccess(Session.getEffectiveUser());
    const P = DriveApp.Permission;
    return [P.OWNER, P.EDIT, P.ORGANIZER, P.FILE_ORGANIZER].indexOf(p) > -1;
  } catch (e) {
    return null;
  }
}

/** Wird von der Web-App beim Öffnen aufgerufen: {level: ok|warn|error, message} */
function checkAccess() {
  try {
    const ss = getSS_();
    const cfg = readConfig_(ss);
    const fb = cfg['Ausweich-Ordner-ID'];
    const fbNote = fb ? ' Es wird im Ausweichordner gespeichert.' : ' Ein Ausweichordner ist in der Config nicht eingetragen.';

    let probe;
    try {
      probe = probeTarget_(cfg, ss, new Date());
    } catch (e) {
      let fbOk = false;
      if (fb) { try { getFolder_(fb, 'Ausweich-Ordner-ID'); fbOk = true; } catch (e2) { /* unten */ } }
      return fbOk
        ? { level: 'warn', message: 'Zielordner nicht erreichbar (' + String(e.message || e).split('\n')[0] + ').' + fbNote }
        : { level: 'error', message: 'Kein Zugriff auf den Zielordner: ' + String(e.message || e).split('\n')[0] +
            '\nDu kannst ausfüllen, Speichern klappt erst wieder mit Zugriff. Bitte Freigabe prüfen' +
            (fb ? ' (auch der Ausweichordner ist nicht erreichbar).' : ' oder einen Ausweichordner in der Config eintragen.') };
    }

    if (canWrite_(probe.folder) === false) {
      return { level: 'warn', message: 'Dein Konto hat vermutlich nur Leserechte auf "' + probe.folder.getName() +
        '". Speichern kann scheitern.' + fbNote };
    }
    return { level: 'ok', message: 'Zielordner "' + probe.folder.getName() + '" erreichbar.' };
  } catch (e) {
    return { level: 'error', message: String(e.message || e) };
  }
}

// ---------- PDF ----------

function exportPdf_(ss, sheet, cfg) {
  const landscape = cfg['Ausrichtung'].toLowerCase().indexOf('quer') === 0;
  const params = {
    format: 'pdf',
    gid: sheet.getSheetId(),
    size: 'A4',
    portrait: !landscape,
    scale: 4,                        // 4 = "An Seite anpassen" → immer genau 1 Seite (Breite UND Höhe)
    top_margin: 0.4,
    bottom_margin: 0.4,
    left_margin: 0.4,
    right_margin: 0.4,
    horizontal_alignment: 'CENTER',
    gridlines: false,
    printtitle: false,
    sheetnames: false,
    pagenumbers: false,
    attachment: false,
  };
  if (cfg['Druckbereich']) params.range = cfg['Druckbereich'];

  const qs = Object.keys(params).map(function (k) { return k + '=' + encodeURIComponent(params[k]); }).join('&');
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?' + qs;
  const res = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('PDF-Export fehlgeschlagen (HTTP ' + res.getResponseCode() + ').');
  }
  return res.getBlob();
}

/** Zählt Seiten-Objekte im PDF (reine Sicherheitsprüfung; 0 = nicht ermittelbar) */
function countPdfPages_(blob) {
  const m = blob.getDataAsString('ISO-8859-1').match(/\/Type\s*\/Page(?![A-Za-z])/g);
  return m ? m.length : 0;
}

function saveFile_(folder, blob, cfg) {
  const name = blob.getName();
  const existing = folder.getFilesByName(name);
  if (existing.hasNext()) {
    if (cfg['Wenn Datei existiert'].toLowerCase() === 'ersetzen') {
      while (existing.hasNext()) existing.next().setTrashed(true);
    } else {
      const base = name.replace(/\.pdf$/i, '');
      let n = 2;
      while (folder.getFilesByName(base + '_' + n + '.pdf').hasNext()) n++;
      blob.setName(base + '_' + n + '.pdf');
    }
  }
  return folder.createFile(blob);
}
