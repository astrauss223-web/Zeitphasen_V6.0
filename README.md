# Zeitphasenmodell V6.0 – Robustes Portfolio

**Zeitphasenorientierte Allokation, Managementansätze, delistete Fonds & intelligente Depotsteuerung**

---

## 🎯 Highlights & Neuerungen in Version 6.0

Version 6.0 erweitert das bewährte Zeitphasenmodell um die vollständige Unterstützung **delisteter Fonds**, ein klares **Farbleitsystem für Fondszustände**, eine **optimierte Sortierlogik** sowie eine **robuste Excel-/CSV-Import-Engine**, die auch komplexe mehrzeilige Depotauszüge auf den Cent genau einliest.

---

### 1. 🛡️ Vollständige Integration & Kennzeichnung delisteter Fonds
* **Aufnahme historischer & delisteter Fonds:**
  * Delistete Fonds (gekennzeichnet durch das interne Flag `_isDelistet: true`) stammen aus der offiziellen VEM-Erweiterungsliste und können weiterhin im Bestand geführt und analysiert werden.
* **Intelligente Investitionssperre (`disabled` / `.delisted-input-locked`):**
  * Für delistete Fonds ohne Vorabbestand (Einmalanlage = `0,00 €` und monatliche Sparrate = `0,00 €`) sind die Eingabefelder im Auswahl-Dialog **gesperrt** und optisch ausgegraut.
  * Ein Hinweistooltip informiert: *„Erstinvestition in delistete Fonds nicht möglich“*.
  * **Bestandsschutz:** Bereits investierte Bestände (> `0,00 €`) bleiben voll editierbar, um bestehende Positionen anpassen, umschichten oder abbauen zu können.
* **Dezentes, gedimmtes Kartendesign:**
  * In der rechten Seitenleiste (*„Auswahl operativer Manager“*) heben sich delistete Fonds durch ein unaufdringliches, helles Design (`.panel-fund-delisted-card`: Hintergrund `#f8fafc`, dezent gestrichelter Rahmen `#cbd5e1`, gedimmte Typografie `#64748b` / `#94a3b8`, `opacity: 0.82`) ab.
  * Dadurch wird jede optische Verwechslung mit farbigen Schichten oder Ländergewichtungs-Balken vermieden.

---

### 2. 🔤 Alphabetische Sortierung & Sektions-Überschriften
* **Strukturierte Sortierung in allen Ansichten:**
  * **Auswahl-Dialog (Modal):** Pro Managementansatz erscheinen zuerst alle **aktiven Fonds (A → Z)** und danach alle **delisteten Fonds (A → Z)**.
  * **Rechtes Panel (*„Auswahl operativer Manager“*):** Dieselbe A–Z-Reihenfolge innerhalb jedes Layers; eine feine, gestrichelte Trennlinie trennt aktive von delisteten Fonds.
* **Prominente farbliche Sektions-Banner im Modal:**
  * **`AKTIV`**: Frischer hellgrüner Hintergrund (`#dcfce7`) mit dunkelgrünem Rand (`#16a34a`) und dunkelgrünem Text – sofort beim Scrollen erfassbar.
  * **`DELISTET`**: Auffälliger hellroter Hintergrund (`#fee2e2`) mit rotem Akzentrand (`#dc2626`) und rotem Text – klare visuelle Zäsur.

---

### 3. 📥 Hochpräzise Depot-Import Engine (CSV / XLSX / PDF / JSON)
* **Behebung des WKN-False-Positive-Filters (`getRowMonetaryAmount`):**
  * Deutsche Währungsbeträge mit Tausendertrennpunkten (z. B. `9.773,02 €`, `6.386,93 €`, `5.042,79 €`, `7.371,50 €` etc.) ergaben nach dem Entfernen von Sonderzeichen exakt 6 Ziffern (`977302`, `638693` usw.) und wurden von älteren Parsern fälschlicherweise als WKN interpretiert und übersprungen.
  * Die Erkennung prüft nun dediziert auf Währungsmerkmale (Komma, `€`, geschützte Leerzeichen) und liest alle Positionen centgenau ein (z. B. Testdatei `AW21_ursprüngliche Formatierung beibehalten.xlsx` mit exakt **643.458,11 €** statt zuvor unvollständigen 599.758,33 €).
* **Mehrzeilige Tabellenerkennung:**
  * Erkennt Summenzeilen und Gruppenköpfe (z. B. wenn der Betrag in Zeile 1 steht und die WKN in Zeile 2).
* **Empfehlungslisten-Automatik:**
  * Fonds ohne direkte WKN-Übereinstimmung in der Datenbank werden anhand ihres Anlageschwerpunkts automatisch der passenden Schicht zugeordnet und als aggregierte Position (*„Empfehlungsliste“*) gebucht.

---

### 4. ⏱️ Das Zeitphasenmodell im Überblick

| Zeitphase | Horizont | Primärer Managementansatz | Schicht-ID |
|---|---|---|---|
| **Phase 1** | `< 1 Jahr` | Geldmarkt | `block-kasse` |
| **Phase 2** | `> 2 Jahre` | Zielrendite / WB Anleihen / EB Anleihen | `block-defensiv` |
| **Phase 3** | `> 4 Jahre` | Zielrendite / WB Anleihen / EB Anleihen | `block-ausgewogen` |
| **Phase 4** | `> 6 Jahre` | WB Aktien/Anleihen | `block-dynamisch` |
| **Phase 5** | `> 8 Jahre` | WB Aktien | `block-maerkte-weit` |
| **Phase 6** | `> 10 Jahre` | EB Aktien | `block-maerkte-eng` |
| **Losgelöst** | Flexibel | Tagesgeld (Bankeinlagen) | `block-tagesgeld` |
| **Losgelöst** | Opportunistisch | Echte / unechte Anlageklassen mit unternehmerischen Risiken | `block-spezial` |

---

### 5. 🛠️ Weitere Komfort- & Analysefunktionen
* **👁️ Sichtbarkeits-Umschalter für operative Manager:**
  * Mit einem Klick auf das Auge-Icon im rechten Panel lassen sich alle Fondskarten temporär ausblenden (ideal für kundenorientierte Strategiegespräche). Schicht-Zylinder und Summen bleiben dabei aktiv.
* **🔍 Sofort-Suche mit Sprungmarke & Flash-Highlight:**
  * Schnelle Fonds-Suche nach Name, WKN oder ISIN mit automatischer Öffnung des Topfes und goldener 2,5-Sekunden-Hervorhebung.
* **💾 Depot-Bibliothek:**
  * Beliebig viele Portfolios lokal im Browser sichern, benennen, laden oder als `.json`-Datei teilen.
* **📄 Anlagevorschlag als PDF:**
  * Generierung eines druckfertigen, tabellarischen PDF-Reports mit Gesamtsummen, Einzelsummen und Sparraten.
* **🌍 Ländergewichtungen:**
  * Detaillierte Top-5-Länderallokationen mit Clusterung (Nordamerika, Europa, Asien, Schwellenländer, Sonstige) per Hover-Tooltip.

---

## 🧭 Bedienungsanleitung

### Depot importieren (CSV, Excel oder PDF)
1. In der oberen Menüleiste auf **„Depot laden (CSV / Excel)“** klicken (oder das Upload-Icon `[ ⬆ ]` in der rechten Seitenleiste nutzen).
2. Eine `.xlsx`-, `.xls`-, `.csv`- oder `.pdf`-Datei auswählen.
3. Die Engine analysiert WKNs, ISINs, Fondsnamen, Anlageschwerpunkte und Beträge.
4. Das Import-Modal zeigt die Trefferquote und den Gesamtwert an. Mit einem Klick wird das Portfolio direkt in die Zeitphasen überführt.

### Fonds manuell zuordnen
1. Auf einen der 3D-Zylinder unter den Zeitphasen 1–6 oder die Spezial-Zylinder klicken.
2. Im geöffneten Dialog nach dem gewünschten Fonds suchen oder filtern.
3. Über den Button **`+ Auswählen`** / **`✓ Ausgewählt`** den Fonds hinzufügen.
4. Einmalbetrag und/oder monatliche Sparrate eintragen (bei delisteten Neuanlagen gesperrt).

### Portfolio sichern & exportieren
* **In Bibliothek:** Auf **„Bibliothek“** klicken → Setup benennen und speichern.
* **JSON-Export:** Auf das Download-Symbol `[ ⬇ ]` im rechten Panel klicken, um eine(`.json`)-Sicherungsdatei herunterzuladen.
* **PDF-Bericht:** Unten im rechten Panel auf **„Anlagevorschlag PDF“** klicken.

---

## 🗂️ Dateistruktur Version 6.0

```text
V6.0/
├── index.html               # Hauptanwendung (Layout, Zeitphasen-Grid, Modals & Toolbar)
├── styles_v6.0.css          # Styling (MLP-Design, Zylinder, delistete Karten, Banner)
├── app_v6.0.js              # Gesamte App-Logik (Import-Engine, Berechnungen, State, Events)
├── fund_data.js             # Fondsdatenbank (inkl. delisteter Fonds, WKNs, ISINs & Länderdaten)
├── crawler_server.py        # Lokaler Python-Server zum Crawlen aktueller Fondsdaten
├── update_country_data.py   # Skript zur automatischen Aktualisierung der Ländergewichtungen
└── README.md                # Diese Dokumentation
```

---

## 💻 Systemvoraussetzungen & Browser

Vollständig clientseitige Single-Page-Applikation. Keine Serverinstallation oder Datenbank erforderlich.

* **Unterstützte Browser:**
  * Google Chrome / Chromium (empfohlen)
  * Apple Safari (macOS & iPadOS)
  * Mozilla Firefox
  * Microsoft Edge

---

*Zeitphasenmodell V6.0 · Stand: September 2026*
