# Gewinnrad Dashboard

Automatisiertes Dashboard für das Sportstech **Gewinnrad-Gewinnspiel**. Liest die täglichen
9:00-Uhr-Tagesberichte aus Outlook (Ordner *Gewinnrad Bewertungen*) und zeigt pro Monat:
Gesamt Nutzer, Rezi-Produkte, Rezis mit Bild, Conversion-Rate, Tagesverlauf.

## Aufbau (3 Schichten)

```
Outlook  ──Power Automate──►  Excel (OneDrive)
                              gewinnrad.json (OneDrive, anonymer Link)
                                        │
                              Vercel-Dashboard (index.html) liest data.json
```

## Dateien
- `index.html` — das komplette Dashboard (eigenständig, kein Build nötig).
- `data.json` — Datenquelle. Wird von Power Automate befüllt. Solange leer → Beispieldaten.

## Datenschema (`data.json`)
```json
{
  "meta": { "source": "live", "updated": "2025-06-05" },
  "days": [
    { "date":"2025-05-01","gesamtNutzer":19,"reziProdukte":18,"ohneRezi":1,
      "abgeschlossen":9,"fotoHochgeladen":8,"conversion":44.4,"rezisBild":8 }
  ]
}
```

## Power-Automate-Flow (ohne Admin / ohne Premium)
1. **Trigger:** *Bei Eingang einer neuen E-Mail (V3)* — Ordner `Gewinnrad Bewertungen`,
   Betrefffilter `Gewinnrad Tagesbericht` (nur die 9:00-Mail, nicht die 16:00-Zwischenberichte).
2. **HTML in Text** — Mail-Body in sauberen Text wandeln.
3. **Verfassen** pro Kennzahl — Wert vor dem Label ausschneiden, z. B. *Gesamt Nutzer*:
   `trim(last(split(first(split(<HTMLzuText>, 'Gesamt Nutzer')), decodeUriComponent('%0A'))))`
4. **Zeile zu Tabelle hinzufügen** (Excel Online) → `Gewinnrad_Daten.xlsx` / Tabelle `Daten`.
5. **(Brücke zum Dashboard)** Excel-Inhalt als JSON zusammensetzen und per
   *Datei erstellen/aktualisieren* nach OneDrive als `gewinnrad.json` schreiben →
   anonymen Freigabe-Link erzeugen → diesen Link als `data.json`-Quelle im Dashboard nutzen.

> **Robuster Endausbau:** Sobald der Kollege (Report-Ersteller) die Zahlen als **CSV-Anhang**
> mitschickt, ersetzen wir Schritt 2–3 durch *CSV einlesen* — dann ist die Pipeline unabhängig
> vom Mail-Layout.

## Deployment (Vercel, wie Wishbeat)
Statische Seite — einfach den Ordner zu Vercel deployen. Kein Node-Build nötig.

## May-Auswertung nachtragen
Sobald die echten Mai-Zahlen vorliegen (Parsing-Backfill oder CSV vom Kollegen),
in `data.json` unter `days` eintragen — das Dashboard rechnet Summen & Conversion automatisch.
