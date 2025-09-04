# Nutzeranleitung: Qualitätsmonitor für Materialstammdaten

## 📋 Übersicht

Der **Qualitätsmonitor für Materialstammdaten** ist eine Webanwendung zur Prüfung von Vollständigkeit & Plausibilität von Produktdaten und zum Vergleich mit Siemens-Webdaten aus MyMobase.

---

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

---

## 🎯 Funktionsweise

### Schritt 1: Datei Upload 📁
- **Unterstützte Formate**: `.xlsx` und `.xls`
- **Maximale Größe**: 10 MB
- **Erforderliche Spalten**: 
  - Spalte C: Material-Kurztext
  - Spalte E: Herstellartikelnummer
  - Spalte N: Fert./Prüfhinweis
  - Spalte P: Werkstoff
  - Spalte S: Nettogewicht
  - Spalte U: Länge
  - Spalte V: Breite
  - Spalte W: Höhe
  - Spalte Z: A2V-Nummer (für Siemens-Produkte)

### Schritt 2: Qualitätsprüfung ✅
**Was wird geprüft?**

1. **Vollständigkeit**: Sind alle erforderlichen Felder ausgefüllt?
2. **Plausibilität**: Sind die Werte logisch und realistisch?
3. **Formatierung**: Stimmen die Datenformate?

**Ergebnis**: 
- **Grüne Zellen**: Vollständig und plausibel
- **Rote Zellen**: Unvollständig oder unplausibel
- **Ampelbewertung**: Gesamtstatus pro Zeile (grün/rot)

### Schritt 3: Web-Suche 🔍
**Nur für Siemens-Produkte** (A2V-Nummern)

**Was wird gesucht?**
- Materialkurztext
- Herstellartikelnummer
- Fert./Prüfhinweis
- Werkstoff
- Nettogewicht
- Länge, Breite, Höhe

**Vergleich**: DB-Werte vs. MyMobase-Webdaten

---

## 📊 Ergebnisinterpretation

### Qualitätsprüfung-Charts
1. **Vollständig & richtig**: Prozentsatz der korrekten Datensätze
2. **Unvollständig/unplausibel**: Prozentsatz der problematischen Datensätze

### Web-Suche-Charts
1. **Schritt 1: Produktanalyse**
   - Gesamtdatensätze vs. Siemens-Produkte
   - Zeigt: Wie viele Produkte sind von Siemens?

2. **Schritt 2: Websuche**
   - Gefundene vs. fehlende Web-Werte
   - Zeigt: Wie erfolgreich war die MyMobase-Suche?

3. **Schritt 3: Vergleich**
   - Übereinstimmungen vs. Abweichungen
   - Zeigt: Wie gut stimmen DB- und Web-Daten überein?

4. **Schritt 4: Ampelbewertung**
   - Qualität OK vs. Qualität fehlerhaft
   - Zeigt: Gesamtqualitätsstatus der Siemens-Produkte

---

## 🔧 Technische Details

### Backend-Technologie
- **Node.js** mit **Express.js**
- **ExcelJS** für Excel-Verarbeitung
- **Cheerio** für Web-Scraping
- **Chart.js** für Diagramme

### Datenquellen
- **Eingabe**: Excel-Dateien (.xlsx/.xls)
- **Web-Daten**: MyMobase (Siemens Mobility)
- **Vergleich**: Automatische Text- und Zahlenvergleiche

### Sicherheit
- **CORS** aktiviert für sichere Kommunikation
- **Helmet** für zusätzliche Sicherheitsheader
- **Dateivalidierung** vor Verarbeitung

---

## 📈 Performance-Optimierungen

### Für bessere Geschwindigkeit:
1. **Dateigröße reduzieren**: Nur notwendige Spalten behalten
2. **Siemens-Produkte filtern**: Nur relevante Zeilen hochladen
3. **Geduld haben**: Bei großen Dateien kann die Verarbeitung länger dauern

### Batch-Verarbeitung:
- **Qualitätsprüfung**: 20 Zeilen pro Batch
- **Web-Suche**: 10 Zeilen pro Batch
- **Automatische Pausen**: Zur Entlastung der CPU

---

## 🚨 Häufige Probleme & Lösungen

### Problem: "Datei zu groß"
**Lösung**: Datei auf unter 10 MB reduzieren oder in kleinere Teile aufteilen

### Problem: "Web-Suche dauert zu lange"
**Lösung**: 
- Weniger Siemens-Produkte in der Datei
- Warten Sie ab (0.1 CPU ist sehr langsam)
- Versuchen Sie es zu einer anderen Tageszeit

### Problem: "Qualitätsprüfung hängt sich auf"
**Lösung**:
- Seite neu laden
- Kleinere Datei verwenden
- Browser-Cache leeren

### Problem: "Keine Siemens-Produkte gefunden"
**Lösung**:
- Prüfen Sie Spalte Z (A2V-Nummern)
- Stellen Sie sicher, dass A2V-Nummern vorhanden sind
- Format: A2V gefolgt von Zahlen

---

## 📞 Support & Kontakt

Bei technischen Problemen oder Fragen:
- **GitHub Repository**: Onssrasra / qmp-produktcheck
- **Service URL**: https://qmp-produktcheck.onrender.com

---

## 🔄 Updates & Verbesserungen

Die Anwendung wird kontinuierlich verbessert:
- **Performance-Optimierungen** für 0.1 CPU
- **Bessere Fehlerbehandlung**
- **Erweiterte Vergleichsalgorithmen**
- **Neue Visualisierungsoptionen**

---

*Letzte Aktualisierung: Dezember 2024*
