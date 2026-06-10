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
- [ ] **Welche Produkte haben eine Rezension erhalten?** (Wunsch Zamy, 10.06.2026 — umsetzen NACH der
      Automatisierung). Neue Sicht „Rezensierte Produkte": pro Produkt die Anzahl Rezensionen, idealerweise
      aufgeschlüsselt nach Plattform (Amazon / Otto / Sportstech.de / Trustpilot …).
      **Datenquellen liegen bereits vor:**
      - Pro Tag: Tabelle *„Rezi-Summary (Produkt × Plattform)"* in jedem 9-Uhr-Bericht.
      - Kampagne gesamt: Sheet-Tab *„Rezi Conversion"* (`Rezi Conversion.csv`, Produkt × Plattform-Kreuztabelle)
        bzw. `Gewinnrad_Rezi.csv` (Spalten `Order`/Produkt + `platform`).
      Einbau analog zu den bestehenden Verteilungs-Karten (Balkenliste je Monat + Kampagne).
