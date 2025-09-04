# Qualitätsmonitor für Materialstammdaten

Ein Tool zur Prüfung von Vollständigkeit & Plausibilität von Materialstammdaten und Vergleich mit Siemens-Webdaten (MyMobase)

## ⚠️ Wichtiger Hinweis: Technische Limitationen

### 🚀 Deployment-Umgebung
Die Anwendung wird über **Render** bereitgestellt und nutzt derzeit die **kostenlose Version** mit folgenden Einschränkungen:

- **CPU**: 0.1 CPU (sehr begrenzt)
- **RAM**: 512 MB (minimal)
- **Performance**: Langsamere Verarbeitung großer Dateien

### 📊 Auswirkungen auf die Nutzung

**Was bedeutet das für Sie?**

1. **Dateigröße**: Maximale Dateigröße auf **10 MB** begrenzt
2. **Verarbeitungszeit**: 
   - Kleine Dateien (< 50 Produkte): ~2-5 Minuten
   - Mittlere Dateien (50-200 Produkte): ~5-15 Minuten
   - Große Dateien (200+ Produkte): ~15-30 Minuten
3. **Web-Suche**: Langsamere Abfrage von MyMobase-Daten
4. **Gleichzeitige Nutzer**: Nur ein Nutzer kann die Anwendung gleichzeitig verwenden

**Empfehlung**: Für größere Datenmengen oder häufige Nutzung empfehlen wir ein Upgrade auf eine kostenpflichtige Version.

## 🆕 Neue Features (Version 2.1)

### Neues Layout: Spaltenblöcke
- **Jeder Fachbegriff bildet einen Spaltenblock** mit DB-Wert (links) und Web-Wert (rechts)
- **Keine zusätzlichen Zeilen** - Web-Daten werden in die gleichen Zeilen geschrieben
- **Übersichtliche Struktur** für bessere Vergleichbarkeit

### Exakte Vergleiche ohne Toleranz
- **Stringfelder**: Exakte Gleichheit nach Trim (case-sensitive)
- **Gewichte**: Exakte Gleichheit der Zahlen in kg
- **Maße**: Exakte Integer-Gleichheit in mm

### Farbkodierung
- 🟢 **Grün**: Exakte Übereinstimmung zwischen DB und Web
- 🔴 **Rot**: Beide Werte vorhanden, aber ungleich
- 🟠 **Orange**: Mindestens ein Wert fehlt

## 📊 Tabellenstruktur

### Eingangstabelle
- **Header in Zeile 3**: Materialkurztext, Her.-Artikelnummer, Fert./Prüfhinweis, Werkstoff, Nettogewicht, Länge, Breite, Höhe
- **Daten ab Zeile 4**: Produkt-ID (A2V) in der entsprechenden Spalte
- **Dynamische Spaltenerkennung** basierend auf Header-Text

### Ausgangstabelle
- **Zeile 3**: Hauptüberschriften als zusammengefasste Blöcke
- **Zeile 4**: Unterüberschriften "DB-Wert" und "Web-Wert"
- **Daten ab Zeile 5**: DB-Werte (links) und Web-Werte (rechts) in den entsprechenden Spalten

## 🚀 Verwendung

1. **Excel-Datei hochladen** über die Web-Oberfläche
2. **Verarbeiten** klicken - das System:
   - Erkennt A2V-Nummern automatisch
   - Scraped Web-Daten von MyMobase
   - Erstellt neue Tabelle mit dem gewünschten Layout
   - Führt exakte Vergleiche durch
   - Markiert Web-Zellen entsprechend der Farbkodierung
3. **Herunterladen** der verarbeiteten Excel-Datei

## 🔧 Technische Details

### Spaltenblöcke
```
A: Produkt-ID (A2V)
C-D: Materialkurztext (DB | Web)
E-F: Her.-Artikelnummer (DB | Web)
G-H: Fert./Prüfhinweis (DB | Web)
I-J: Werkstoff (DB | Web)
K-L: Nettogewicht (DB | Web)
M-N: Länge (DB | Web)
O-P: Breite (DB | Web)
Q-R: Höhe (DB | Web)
```

### Vergleichslogik
- **Materialkurztext**: Exakte String-Gleichheit
- **Her.-Artikelnummer**: Normalisierte Artikelnummer-Vergleiche
- **Werkstoff**: Exakte String-Gleichheit
- **Nettogewicht**: Exakte Zahlen-Gleichheit in kg
- **Abmessungen**: Exakte Integer-Gleichheit in mm

### Web-Daten-Extraktion
- **Produkttitel** → Materialkurztext (Web)
- **Weitere Artikelnummer** → Her.-Artikelnummer (Web)
- **Werkstoff** → Werkstoff (Web)
- **Gewicht** → Nettogewicht (Web) in kg
- **Abmessungen** → Länge/Breite/Höhe (Web) in mm

## 📋 Anforderungen

- Node.js >= 18
- Excel-Dateien mit A2V-Nummern
- Internetverbindung für MyMobase-Scraping

## 🛠️ Lokale Installation

```bash
git clone <repository-url>
cd qmp-siemens-produktcheck-main
npm install
npm start
```

Das Tool läuft dann unter `http://localhost:3000`

## 🚀 Deployment auf Render

### 1. GitHub Repository vorbereiten
```bash
git add .
git commit -m "QMP Siemens Produktcheck v2.1 - Neues Layout mit Spaltenblöcken"
git push origin main
```

### 2. Render Service erstellen
1. Gehen Sie zu [render.com](https://render.com)
2. Klicken Sie auf "New +" → "Web Service"
3. Verbinden Sie Ihr GitHub Repository
4. Konfigurieren Sie den Service:
   - **Name**: `qmp-siemens-produktcheck`
   - **Environment**: `Node`
   - **Build Command**: `npm install && npm run install-browsers`
   - **Start Command**: `node server.js`
   - **Plan**: `Starter` (oder höher)

### 3. Umgebungsvariablen setzen
- `SCRAPE_CONCURRENCY`: `4`
- `NODE_VERSION`: `18`
- `DISABLE_PLAYWRIGHT`: `0`

### 4. Deploy
- Klicken Sie auf "Create Web Service"
- Render baut und deployed automatisch
- Die URL wird nach dem Build angezeigt

## 🧪 Testen

### Test-Excel erstellen
```bash
node test-excel.js
```

### Test-Datei verwenden
1. Öffnen Sie die Web-Oberfläche
2. Laden Sie `test-input.xlsx` hoch
3. Klicken Sie auf "Verarbeiten"
4. Laden Sie das Ergebnis herunter

## 📝 Changelog

### Version 2.1
- ✅ Neues Layout mit Spaltenblöcken
- ✅ Keine zusätzlichen Zeilen mehr
- ✅ Exakte Vergleiche ohne Toleranz
- ✅ Verbesserte Farbkodierung (Grün/Rot/Orange)
- ✅ Dynamische Spaltenerkennung
- ✅ Optimierte Web-Daten-Extraktion
- ✅ Render-Deployment optimiert
- ✅ Verbessertes Error-Handling

### Version 2.0
- Grundlegende Funktionalität
- Excel-Verarbeitung
- MyMobase-Scraping
- Einfache Vergleiche

## 🔍 Troubleshooting

### HTTP 500 Fehler
- Überprüfen Sie die Render-Logs
- Stellen Sie sicher, dass alle Abhängigkeiten installiert sind
- Überprüfen Sie die Excel-Datei auf korrekte Struktur

### Spaltenblöcke werden nicht erstellt
- Überprüfen Sie die Header-Zeile (Zeile 3)
- Stellen Sie sicher, dass A2V-Nummern in der Produkt-ID-Spalte stehen
- Überprüfen Sie die Render-Logs für Details
