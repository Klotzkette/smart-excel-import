# Smart Excel Import

Ein Excel-Add-in, das Dokumente (PDF, Word, Excel, Bilder/Screenshots) per Drag & Drop importiert und als **dynamische Tabelle mit echten Excel-Formeln** einfügt.

![Excel Add-in](https://img.shields.io/badge/Excel-Add--in-217346?logo=microsoft-excel&logoColor=white)

## Was macht dieses Add-in?

Statt toter Zahlen aus PDFs oder Screenshots erzeugt Smart Excel Import **lebendige Tabellen**:

- **Summen** werden als `=SUM(B2:B10)` eingefügt, nicht als feste Zahl
- **Berechnungen** (Prozent, Differenzen, Durchschnitte) als echte Formeln
- **Professionelle Formatierung** mit Überschriften, Farben, Rahmen und Zahlenformaten
- **Metadaten** (Quelle, Datum) werden automatisch erkannt und angezeigt
- **Inhaltsverzeichnisse** aus anwaltlichen Akten und Gerichtsakten werden strukturiert in Excel abgebildet

## Unterstützte Dateiformate

| Format | Verarbeitung |
|--------|-------------|
| **PDF** | Text-Extraktion via pdf.js |
| **Word (.docx)** | Text-Extraktion via mammoth.js |
| **Excel (.xlsx/.xls)** | Daten-Extraktion via SheetJS |
| **Bilder (JPG, PNG, etc.)** | Vision-API (GPT-4o / Claude) analysiert Screenshots |

## KI-Anbieter

Wahlweise:
- **OpenAI** (GPT-4o) — mit Vision für Bilder
- **Anthropic** (Claude Sonnet) — mit Vision für Bilder

Du brauchst einen eigenen API-Schlüssel, den du in den Einstellungen des Add-ins eingibst.

## Installation (Sideloading)

### Voraussetzungen

- Microsoft Excel (Desktop oder Web)
- Ein API-Schlüssel von [OpenAI](https://platform.openai.com/api-keys) oder [Anthropic](https://console.anthropic.com/)
- Ein lokaler Webserver (z.B. `http-server`, `live-server`, oder Python)

### Schritt 1: Repository klonen

```bash
git clone https://github.com/Klotzkette/smart-excel-import.git
cd smart-excel-import
```

### Schritt 2: Lokalen Server starten

```bash
# Option A: Mit npx (kein npm install nötig)
npx http-server ./src -p 3000 -c-1 --cors

# Option B: Mit Python
cd src && python -m http.server 3000
```

### Schritt 3: Manifest-URL anpassen

In `manifest.xml` alle URLs von `https://localhost:3000` auf `http://localhost:3000` ändern (falls kein HTTPS).

Für HTTPS (empfohlen für Excel Desktop):
```bash
# Zertifikate generieren
npx office-addin-dev-certs install
# Server mit HTTPS starten
npx http-server ./src -p 3000 --ssl --cert ~/.office-addin-dev-certs/localhost.crt --key ~/.office-addin-dev-certs/localhost.key -c-1
```

### Schritt 4: Add-in in Excel laden

#### Excel Desktop (Windows)

1. Öffne Excel
2. Gehe zu **Einfügen** → **Add-ins** → **Meine Add-ins**
3. Klicke auf **Benutzerdefinierte Add-ins hochladen**
4. Wähle die `manifest.xml` Datei aus
5. Das Add-in erscheint im **Home**-Tab als "Smart Import"

#### Excel Desktop (Mac)

1. Öffne Excel
2. Gehe zu **Einfügen** → **Add-ins** → **Meine Add-ins**
3. Klicke auf das Zahnrad-Symbol → **Benutzerdefinierte Add-ins hochladen**
4. Wähle die `manifest.xml` Datei aus

#### Excel Online

1. Öffne eine Excel-Datei in Office 365
2. Gehe zu **Einfügen** → **Office-Add-ins**
3. Klicke auf **Mein Add-in hochladen**
4. Wähle die `manifest.xml` Datei aus

### Schritt 5: API-Schlüssel eingeben

1. Öffne das Add-in (Button "Smart Import" im Home-Tab)
2. Klicke auf **Einstellungen**
3. Wähle deinen KI-Anbieter (OpenAI oder Anthropic)
4. Gib deinen API-Schlüssel ein und klicke **Speichern**

## Benutzung

1. **Datei importieren**: Ziehe eine Datei per Drag & Drop in die Drop-Zone, oder klicke auf "Datei auswählen"
2. **Optionen wählen**:
   - **Startposition**: Wo die Tabelle beginnen soll (Standard: A1)
   - **Professionell formatieren**: Farben, Rahmen, Zahlenformate
   - **Formeln statt Festwerte**: Berechnungen als echte Excel-Formeln
3. **Importieren**: Klicke auf den Button und warte auf die Verarbeitung

## Wie funktioniert die Formel-Erkennung?

Die KI analysiert das Dokument und erkennt:

- **Summenzeilen**: Wenn eine Zeile offensichtlich die Summe der darüber liegenden Werte ist → `=SUM(...)`
- **Prozentberechnungen**: Anteile und Margen → `=B2/B$10` oder `=(B2-C2)/B2`
- **Differenzen**: Veränderungen → `=B2-C2`
- **Durchschnitte**: Mittelwerte → `=AVERAGE(...)`

Alle Formeln verwenden englische Funktionsnamen (die Excel automatisch in die lokale Sprache übersetzt).

## Projektstruktur

```
smart-excel-import/
├── manifest.xml              # Office Add-in Manifest
├── package.json
├── README.md
├── LICENSE
├── .gitignore
├── scripts/
│   └── generate-icons.py     # Icon-Generator
└── src/
    ├── assets/               # Add-in Icons
    │   ├── icon-16.png
    │   ├── icon-32.png
    │   ├── icon-64.png
    │   └── icon-80.png
    └── taskpane/
        ├── taskpane.html     # Sidebar HTML
        ├── taskpane.css      # Styles
        └── taskpane.js       # Kernlogik
```

## Technologie

- **Office.js** — Excel JavaScript API
- **pdf.js** — PDF-Textextraktion (CDN)
- **mammoth.js** — Word-Dokumentverarbeitung (CDN)
- **SheetJS** — Excel-Datei-Parsing (CDN)
- **OpenAI API** / **Anthropic API** — KI-Analyse mit Vision-Support

## Einschränkungen

- Die Qualität der Formel-Erkennung hängt von der Klarheit des Quelldokuments ab
- Sehr komplexe verschachtelte Tabellen können vereinfacht werden
- Screenshots müssen lesbar sein (ausreichende Auflösung)
- API-Kosten fallen pro Import an (typisch: $0.01–0.10 pro Dokument, Bilder etwas mehr)

## Lizenz

MIT — siehe [LICENSE](LICENSE).
