# Investitionsrechner

Web-Version des Investitionsrechners für Wirtschaftlichkeits- und Investitionsrechnungen.
Läuft komplett im Browser – keine Installation, keine Server, keine externen Bibliotheken.
Eingaben werden nicht übertragen.

## Verfahren
- **Statisch:** Kostenvergleich (inkl. kritischer Menge), Gewinnvergleich (inkl. Break-even), Rentabilität (ROI), Amortisation
- **Dynamisch:** Kapitalwert, interner Zinsfuß (IRR/MIRR, Fisher-Zins), Annuität, dynamische Amortisation, Vermögensendwert (Soll-/Habenzins)
- **Wertorientiert:** Economic Value Added / Market Value Added
- **Risiko:** Sensitivitätsanalyse mit Tornado- und Spinnendiagramm, kritische Werte

## Export
- **PDF / Drucken:** Bericht mit allen Diagrammen → im Druckdialog „Als PDF speichern“ wählen
- **TXT** (Bericht) und **CSV** (Excel-kompatibel, Semikolon/Dezimalkomma)
- **Eingaben speichern/laden** als JSON-Datei

## Veröffentlichen über GitHub Pages
1. Neues Repository anlegen und `index.html` (und diese README) hochladen.
2. *Settings → Pages → Build and deployment → Source: „Deploy from a branch“*, Branch `main`, Ordner `/ (root)` → *Save*.
3. Nach ca. 1 Minute ist der Rechner unter `https://<benutzername>.github.io/<repository>/` erreichbar.

Hinweis: Modellrechnungen vor Steuern (außer EVA); kein Ersatz für steuerliche oder betriebswirtschaftliche Beratung.
