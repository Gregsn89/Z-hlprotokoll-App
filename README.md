# Zählprotokoll (Kassenabschluss)

Google-Apps-Script-Web-App (Tablet) + Container-Skript im Google Sheet „Zählprotokoll“ (studiobloc.de).

## Was sie tut
- Formular für alle grünen Felder → schreibt ins Sheet → rechnet nach (`verifyTotals_`)
- Geht der Abschluss auf: **ein A4-PDF** in Drive (Pfad aus `Config`)
- „Differenz laut Tillhub“ ≠ 0 (über Toleranz): **Kopie des Sheets** statt PDF
- EC-Belege: Fotos → automatischer Zuschnitt/Drehung (Vorschau, manuell korrigierbar) → **ein PDF**
- Mail/Slack bei jedem Abschluss (Datum, SOLL, IST, Differenz, Links), Fehler-Mail, Abend-Erinnerung
- Zugriffsprüfung + Ausweichordner; Entwurf im Browser (localStorage)

## Struktur
| Pfad | Inhalt |
|---|---|
| `src/Code.gs` | Server (Export, Config, Mail/Slack, Trigger) |
| `src/Index.html` | Web-App (Formular, Foto-Pipeline, PDF-Erzeugung im Browser) |
| `src/appsscript.json` | Manifest (Scopes, Web-App-Einstellungen) |
| `tests/server.test.js` | Server-Tests (vm mit Apps-Script-Mocks): `npm test` |
| `docs/` | Screenshot „Rechte zulassen“ |

## Einrichten
1. Sheet öffnen → Erweiterungen → Apps Script
2. `Code.gs`, `Index.html` und `appsscript.json` einfügen (oder per `clasp`, siehe unten)
3. Reiter `Config` und `Felder` werden beim ersten Start angelegt
4. Bereitstellen → Web-App: *Ausführen als* „Nutzer, der auf die Web-App zugreift“, *Zugriff* „Jeder mit Google-Konto“
5. Rechte einmalig zulassen (`docs/Zaehlprotokoll_Rechte_zulassen.png`)

## Config (Reiter `Config`)
- Ordner = `Root-Ordner-ID` + Pfad; Platzhalter `{JJJJ}` `{MM}` `{TT}` (Datum des Abschlusses)
  - `Zielpfad` (PDF), `Differenz-Pfad` (Kopie; leer = Zielordner), `Belege-Pfad` (EC-Belege)
  - `Zielordner-ID` überschreibt Root + Zielpfad
- `Ausweich-Ordner-ID`: Ordner in Meine Ablage des ausführenden Kontos, falls der Zielordner nicht erreichbar ist
- `E-Mail-Empfänger`, `Slack-Webhook`, `Meldung bei jedem Abschluss` (ja/nein)
- Config-Änderungen gelten sofort; **nur bei Code-Änderungen** neue Version bereitstellen

## Neue Version bereitstellen
Bereitstellen → Bereitstellungen verwalten → Stift → *Neue Version* → Bereitstellen (URL bleibt gleich).

## clasp (optional)
```
npm i -g @google/clasp && clasp login
cp .clasp.json.example .clasp.json   # scriptId eintragen (.clasp.json ist git-ignoriert)
clasp push
```

## Sicherheit
- **Slack-Webhook-URL ist ein Geheimnis**: steht im Reiter `Config` (auch in Sheet-Kopien!) → Reiter schützen/ausblenden, nie ins Repo committen
- Keine Tokens/Webhooks im Code oder in Tests
