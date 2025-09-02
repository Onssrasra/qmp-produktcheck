// 4) Verarbeite die Siemens-Datei
    for (const ws of siemensWb.worksheets) {
      // 4.1 Spaltenstruktur berechnen
      const structure = calculateNewColumnStructure(ws);

      // 4.2 Spalten einfügen (von rechts nach links)
      for (const pair of [...structure.pairs].reverse()) {
        const insertPos = getColumnIndex(pair.original) + 1; // rechts neben der Originalspalte
        ws.spliceColumns(insertPos, 0, [null]);
      }

      // 4.3 Zeile 4 (Labels) einfügen
      ws.spliceRows(LABEL_ROW, 0, [null]);

      // 4.4 Zeilen 2 & 3 Inhalte in Web-Spalten spiegeln + Labels schreiben
      for (const pair of structure.pairs) {
        // Inhalte 2/3 spiegeln
        const dbTech = ws.getCell(`${pair.dbCol}2`).value;
        const dbName = ws.getCell(`${pair.dbCol}3`).value;
        ws.getCell(`${pair.webCol}2`).value = dbTech;
        ws.getCell(`${pair.webCol}3`).value = dbName;
        copyColumnFormatting(ws, pair.dbCol, pair.webCol, 1, 3);

        // Labels Zeile 4
        ws.getCell(`${pair.dbCol}${LABEL_ROW}`).value  = 'DB-Wert';
        ws.getCell(`${pair.webCol}${LABEL_ROW}`).value = 'Web-Wert';
        applyLabelCellFormatting(ws, `${pair.dbCol}${LABEL_ROW}`, false);
        applyLabelCellFormatting(ws, `${pair.webCol}${LABEL_ROW}`, true);
      }

      // 4.5 Top-Header (Zeile 1) setzen
      applyTopHeader(ws);

      // 4.6 Header in Zeile 2 und 3 pro Paar zusammenfassen
      mergePairHeaders(ws, structure.pairs);

      // === Ampel-Spalte anhängen ===
      const ampInsertIndex = ws.columnCount + 1;
      ws.spliceColumns(ampInsertIndex, 0, [null]);
      const ampCol = getColumnLetter(ampInsertIndex);
      ws.getCell(`${ampCol}2`).value = 'AMP';
      ws.getCell(`${ampCol}3`).value = 'Ampelbewertung';
      ws.getCell(`${ampCol}${LABEL_ROW}`).value = 'Status';
      applyLabelCellFormatting(ws, `${ampCol}${LABEL_ROW}`, false);
      ws.getCell(`${ampCol}2`).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      ws.getCell(`${ampCol}3`).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

      // 4.7 Web-Daten eintragen / vergleichen
      const siemensRows = siemensRowsPerSheet.get(wb.worksheets.find(w => w.name === ws.name)) || [];
      for (let i = 0; i < siemensRows.length; i++) {
        const currentRow = 5 + i; // Start ab Zeile 5 (nach Labels)

        // Ampel-Accumulator für diese Zeile: Status per Spalte
        const columnStatus = { E:'none', N:'none', S:'none', U:'none', V:'none', W:'none' };

        // neue Z-Spalte (A2V) bestimmen
        let zCol = ORIGINAL_COLS.Z;
        if (structure.otherCols.has(ORIGINAL_COLS.Z)) zCol = structure.otherCols.get(ORIGINAL_COLS.Z);
        const a2v = (ws.getCell(`${zCol}${currentRow}`).value || '').toString().trim().toUpperCase();
        const web = resultsMap.get(a2v) || {};

        // je Paar
        for (const pair of structure.pairs) {
          const dbValue = ws.getCell(`${pair.dbCol}${currentRow}`).value;
          let webValue = null;
          let isEqual = false;

          switch (pair.original) {
            case 'C': // Material-Kurztext
              webValue = (web.Produkttitel && web.Produkttitel !== 'Nicht gefunden') ? web.Produkttitel : null;
              isEqual  = webValue ? eqText(dbValue || '', webValue) : false;
              break;
            case 'E': // Herstellartikelnummer - ANGEPASSTE LOGIK
              // Wenn DB-Wert mit A2V anfängt → Web-Wert = DB-Wert
              if (String(dbValue || '').trim().toUpperCase().startsWith('A2V')) {
                webValue = dbValue;
                isEqual = true; // DB-Wert = Web-Wert, also Übereinstimmung
              } else {
                // Sonst „Weitere Artikelnummer" oder Fallback A2V
                webValue = (web['Weitere Artikelnummer'] && web['Weitere Artikelnummer'] !== 'Nicht gefunden')
                          ? web['Weitere Artikelnummer']
                          : a2v;
                isEqual = eqPart(dbValue || a2v, webValue);
              }
              break;
            case 'N': // Fert./Prüfhinweis
              if (web.Materialklassifizierung && web.Materialklassifizierung !== 'Nicht gefunden') {
                const code = normalizeNCode(mapMaterialClassificationToExcel(web.Materialklassifizierung));
                if (code) { webValue = code; isEqual = eqN(dbValue || '', code); }
              }
              break;
            case 'P': // Werkstoff
              webValue = (web.Werkstoff && web.Werkstoff !== 'Nicht gefunden') ? web.Werkstoff : null;
              isEqual  = webValue ? eqText(dbValue || '', webValue) : false;
              break;
            case 'S': // Nettogewicht
              if (web.Gewicht && web.Gewicht !== 'Nicht gefunden') {
                const { value } = parseWeight(web.Gewicht);
                if (value != null) { webValue = value; isEqual = eqWeight(dbValue, web.Gewicht); }
              }
              break;
            case 'U': // Länge
              if (web.Abmessung && web.Abmessung !== 'Nicht gefunden') {
                const d = parseDimensionsToLBH(web.Abmessung);
                if (d.L != null) { webValue = d.L; isEqual = eqDimension(dbValue, web.Abmessung, 'L'); }
              }
              break;
            case 'V': // Breite
              if (web.Abmessung && web.Abmessung !== 'Nicht gefunden') {
                const d = parseDimensionsToLBH(web.Abmessung);
                if (d.B != null) { webValue = d.B; isEqual = eqDimension(dbValue, web.Abmessung, 'B'); }
              }
              break;
            case 'W': // Höhe
              if (web.Abmessung && web.Abmessung !== 'Nicht gefunden') {
                const d = parseDimensionsToLBH(web.Abmessung);
                if (d.H != null) { webValue = d.H; isEqual = eqDimension(dbValue, web.Abmessung, 'H'); }
              }
              break;
          }

          const hasDb = hasValue(dbValue);
          const hasWeb = webValue !== null;

          if (hasWeb) {
            ws.getCell(`${pair.webCol}${currentRow}`).value = webValue;
            // Nur markieren wenn DB-Wert vorhanden ist
            if (hasDb) {
              fillColor(ws, `${pair.webCol}${currentRow}`, isEqual ? 'green' : 'red');
            }
          } else {
            // Web-Wert fehlt → orange
            fillColor(ws, `${pair.webCol}${currentRow}`, 'orange');
          }

          // Für Ampelzustand relevante Spalten tracken
          if (AMPEL_REQUIRED_ORIGINALS.includes(pair.original)) {
            if (hasWeb && isEqual) {
              columnStatus[pair.original] = 'green';
            } else if (hasWeb && !isEqual) {
              columnStatus[pair.original] = 'red';
            } else {
              // Orange (Web-Wert fehlt) zählt als OK für Ampel
              columnStatus[pair.original] = 'orange';
            }
          }
        }

        // Ampelwert setzen: ROT wenn mindestens eine Pflichtspalte rot, sonst GRÜN
        // (Orange zählt als OK)
        const hasRedColumn = Object.values(columnStatus).includes('red');
        const ampelColor = hasRedColumn ? 'red' : 'green';
        
        // ANPASSUNG: Keine Texte mehr in den Datenzeilen → nur farbliche Markierung
        ws.getCell(`${ampCol}${currentRow}`).value = ''; // Kein Text
        fillColor(ws, `${ampCol}${currentRow}`, ampelColor);
        ws.getCell(`${ampCol}${currentRow}`).alignment = { horizontal: 'center', vertical: 'middle' };
      }
    }

    const out = await siemensWb.xlsx.writeBuffer();
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition','attachment; filename="Web_Vergleich_Siemens.xlsx"');
    res.send(Buffer.from(out));

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => console.log(`Server running at http://0.0.0.0:${PORT}`));
