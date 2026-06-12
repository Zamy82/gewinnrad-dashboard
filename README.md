# Gewinnspiel KPI Dashboard

Eigenständiges Dashboard für das Sportstech **Gewinnrad-Gewinnspiel** (All-in-One Flyer).
Dunkles, self-contained HTML-Dashboard (kein Build nötig) mit zwei Sichten:

- **Kampagne gesamt** — Aggregate über die ganze Kampagne (seit 12.03.2026): Teilnehmer,
  zu Ende gespielt, Fotos, Ø Produktbewertung, Sterne-/Plattform-/Land-/Gewinn-Verteilung.
- **Monats-Tabs** — die 4 Kern-KPIs pro Monat: **Gesamt Nutzer · Rezis erzeugt · Conversion ·
  Beste Conversion**, plus Tagesverlauf und Tagestabelle.

> Komplett **unabhängig vom PMO-Dashboard** und **PII-frei** (nur aggregierte Zahlen).

## Zwei Datenquellen — bewusst getrennt

| Sicht | Quelle | Warum |
|---|---|---|
| **Monats-Tabs** (Tageswerte) | 9-Uhr-Mail-Tagesberichte | Das Live-System zählt pro Tag korrekt. |
| **Kampagne gesamt** | All-in-One-Flyer-Sheet (`Gewinnrad_nutzer` + `Gewinnrad_Rezi`) | Authentische Gesamtsummen. |

⚠️ Die Tagesreihe **nicht** aus dem Sheet bauen: die Spalte `Gewinnrad_Rezi.created` ist ein
Batch-Import-Zeitstempel (alle Fotos auf einen Tag geklumpt), taugt nicht für eine echte Tagesreihe.

## Dateien
- `index.html` — das komplette Dashboard (dunkles Theme, Chart.js via CDN).
- `data.js` — **die geladene Datenquelle.** Setzt `window.DASHBOARD_DATA`. Wird per `<script>`-Tag
  eingebunden → funktioniert **lokal (file://) und online** identisch (kein `fetch`-/CORS-Problem).
- `data.json` — lesbare Quell-Kopie desselben Inhalts (für Diffs/Doku). `data.js` ist maßgeblich.

## Datenschema
```js
window.DASHBOARD_DATA = {
  meta: { source:"live", updated:"2026-06-10", hinweis:"…" },
  campaign: {
    teilnehmer:906, zuEnde:267, completed:314, conversionPct:29.5,
    fotosHochgeladen:164, fotosValidiert:140, bewertungen:807, avgSterne:4.57,
    sterne:{ "5":523,"4":232,"3":44,"2":3,"1":5 },
    plattform:{…}, land:{…}, gewinn:{…}
  },
  days: [ { date:"2026-05-01", gesamtNutzer:15, reziProdukte:15, ohneRezi:0,
            abgeschlossen:6, fotoHochgeladen:4, conversion:26.7, rezisBild:4 }, … ]
};
```

## KPI-Definitionen (transparent)
- **Teilnehmer** = Sheet-Zeile mit Zeitstempel (`Created`). Hinweis: das PMO filtert ~8 Test-/Admin-Zeilen → 898.
- **Zu Ende gespielt** = letzte Seite „Zu Ende gespielt" (267). *Alternative:* Sheet-Spalte `Completed=TRUE` (314).
- **Conversion** (Monats-KPI) = Rezis mit Bild / Gesamt Nutzer.
- **Beste Conversion** = bester Einzeltag im Monat.

## Deployment (Vercel via GitHub)
Statische Seite, kein Node-Build. Repo zu Vercel verbinden → jeder Commit löst Auto-Deploy aus.
Git-Identität (Sportstech-Standard, lokal pro Repo): `user.name=PMO-SP`, `user.email=pmo@sportstech.de`.

## Ziel-Architektur (voll automatisch)
```
9-Uhr-Mail ─► Power Automate (Flow 1, täglich) ─┐
                                                 ├─► GitHub (data.js) ─► Vercel Auto-Deploy ─► Link
Flyer-Sheet ─► Power Automate (Flow 2, wöchentl.)┘
```
Power Automate schreibt die neuen Zahlen per GitHub-Connector ins Repo; Vercel baut automatisch neu.

## Daten manuell aktualisieren (Fallback)
Frischen Sheet-Export (`.xlsx`) ablegen und `update-data.ps1` ausführen — regeneriert die
Kampagnen-Aggregate in `data.js`/`data.json`. Tageswerte kommen aus den Mails (siehe Flow 1).

## Backlog / Geplante Erweiterungen
- [x] **Welche Produkte haben eine Rezension erhalten?** ✅ umgesetzt 11.06.2026 — Sicht „Rezensierte Produkte"
      (Produkt × Plattform-Matrix) in Kampagne-gesamt + Monats-Tabs. Quelle: `Gewinnrad_Rezi.csv`,
      Definition = Rezension mit Bild (`Bewertungs_Foto` befüllt). Plattform `Sportstech.de` wird mit
      `Trustpilot` zusammengefasst (so wie in den Tagesberichten). Berechnung steckt in `update-data.ps1`.
- **Phase 3 — Sheet-Anbindung:** Voll-Automatik via Pipedream+Google-OAuth wurde **bewusst verworfen**
      (11.06.2026) — Pipedreams Google-Connector verlangt breiten „alle Drive-Dateien"-Zugriff auf das
      (Firmen-)Google-Konto, unverhältnismäßig fürs Lesen eines Sheets. **Entscheidung: semi-automatisch.**
      Kampagne/Verteilungen/Produkte werden bei Bedarf per `update-data.ps1` aktualisiert (Sheet als `.xlsx`
      exportieren → Skript → commit/push). Voll-Auto bliebe nur möglich, wenn der Kollege **PII-freie
      Auswertungs-Tabs** als CSV „im Web veröffentlicht" — dann könnten wir die ohne Google-Login ziehen.
