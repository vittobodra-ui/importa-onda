/* ANAGRAFE -> ONDE: copia le colonne "gialle" dell'anagrafe nelle colonne corrispondenti del file onda,
   con i capi nell'ordine dell'anagrafe. Lavora su oggetti cella di SheetJS ({t, v, w, z}). */
(function (root) {
  'use strict';

  // Colonne gialle del foglio BOVINI dell'anagrafe -> nome della colonna nel file onda
  var GIALLE_BOVINI = [
    ['CODICE CAPO', 'CODICE CAPO'],
    ['CODICE_ELETTRONICO', 'CODICE_ELETTRONICO'],
    ['RAZZA CODICE', 'RAZZA CODICE'],
    ['SESSO', 'SESSO'],
    ['CODICE MADRE', 'CODICE MADRE'],
    ['MOTIVO INGRESSO', 'MOTIVO INGRESSO'],
    ['DATA NASCITA', 'DATA NASCITA'],
    ['DATA INGRESSO', 'DATA INGRESSO'],
    ['CODICE AZIENDA PROVENIENZA', 'CODICE AZIENDA PROVENIENZA'],
    ['CODICE FISCALE PROVENIENZA', 'CODICE FISCALE PROVENIENZA'],
    ['SPECIE AZIENDA PROVENIENZA', 'SPECIE AZIENDA PROVENIENZA'],
    ['ESTREMI MODELLO INGRESSO', 'NUMERO MODELLO INGRESSO'],
    ['DATA MODELLO INGRESSO', 'DATA MODELLO INGRESSO'],
    ['MOTIVO USCITA', 'MOTIVO USCITA'],
    ['DATA USCITA', 'DATA USCITA'],
    ['CODICE AZIENDA DESTINAZIONE', 'CODICE AZIENDA DESTINAZIONE'],
    ['NUMERO MODELLO USCITA', 'NUMERO MODELLO USCITA'],
    ['DATA MODELLO USCITA', 'DATA MODELLO USCITA'],
    ['INTERVALLO DATE', 'INTERVALLO DATE'],
    ['DATA MACELLAZIONE', 'INTERVALLO DATE'], // stessa colonna (AC) nell'export originale della BDN
    ['DATA ESTRAZIONE DATI', 'DATA ESTRAZIONE DATI']
  ];
  // Colonne gialle del foglio DATI DELL'ATTIVITA (stesso valore per tutti i capi)
  var GIALLE_ATTIVITA = ['AZIENDA CODICE', 'CODICE FISCALE', 'CODICE SPECIE'];

  function norm(s) { return String(s == null ? '' : s).toUpperCase().replace(/\s+/g, ' ').trim(); }
  function cellText(c) {
    if (!c || c.v == null) return '';
    if (c.v instanceof Date) return c.v.toISOString().slice(0, 10);
    return String(c.v).trim();
  }
  function cloneCell(c) {
    if (!c || c.v == null || cellText(c) === '') return null;
    var n = { t: c.t, v: c.v };
    if (c.z) n.z = c.z;
    if (c.w != null) n.w = c.w;
    return n;
  }
  function addr(XLSX, r, c) { return XLSX.utils.encode_cell({ r: r, c: c }); }

  // Cerca in un foglio la riga di intestazione che contiene la colonna "chiave"
  function findHeader(ws, XLSX, key) {
    if (!ws || !ws['!ref']) return null;
    var rg = XLSX.utils.decode_range(ws['!ref']);
    for (var r = rg.s.r; r <= Math.min(rg.e.r, rg.s.r + 15); r++) {
      var cols = [], hit = false;
      for (var c = rg.s.c; c <= rg.e.c; c++) {
        var h = norm(cellText(ws[addr(XLSX, r, c)]));
        cols.push({ c: c, name: h });
        if (h === key) hit = true;
      }
      if (hit) return { ws: ws, headerRow: r, cols: cols, range: rg };
    }
    return null;
  }
  function findSheet(wb, XLSX, key, prefer) {
    var names = wb.SheetNames.slice();
    names.sort(function (a, b) { return (norm(b) === prefer) - (norm(a) === prefer); });
    for (var i = 0; i < names.length; i++) {
      var h = findHeader(wb.Sheets[names[i]], XLSX, key);
      if (h) { h.name = names[i]; return h; }
    }
    return null;
  }
  function colOf(h, name) {
    for (var i = 0; i < h.cols.length; i++) if (h.cols[i].name === name) return h.cols[i].c;
    return -1;
  }

  var COLONNA_A = 'CAR-IT';

  // Colonne del file onda da eliminare (U-Z), con il loro contenuto
  var SCARTA = ['CODICE FIERA DESTINAZIONE', 'CODICE PAESE DESTINAZIONE', 'CODICE STATO DESTINAZIONE',
    'REGIONE MACELLO DESTINAZIONE', 'CODICE MACELLO DESTINAZIONE', 'CODICE PASCOLO DESTINAZIONE',
    'CODICE ASSEGNATO IN PRECEDENZA'];

  // Elimina una colonna spostando a sinistra quelle successive
  function deleteCol(ws, XLSX, col) {
    var rg = XLSX.utils.decode_range(ws['!ref']);
    for (var r = rg.s.r; r <= rg.e.r; r++) {
      for (var c = col; c <= rg.e.c; c++) {
        var next = ws[addr(XLSX, r, c + 1)];
        if (next) ws[addr(XLSX, r, c)] = next; else delete ws[addr(XLSX, r, c)];
      }
      delete ws[addr(XLSX, r, rg.e.c)];
    }
    rg.e.c -= 1;
    ws['!ref'] = XLSX.utils.encode_range(rg);
    if (ws['!cols']) ws['!cols'].splice(col, 1);
    delete ws['!merges'];
  }

  // Sposta tutte le celle del foglio a destra di una colonna (libera la colonna A)
  function shiftRight(ws, XLSX) {
    var rg = XLSX.utils.decode_range(ws['!ref']);
    for (var r = rg.s.r; r <= rg.e.r; r++) {
      for (var c = rg.e.c; c >= rg.s.c; c--) {
        var a = addr(XLSX, r, c);
        if (ws[a]) { ws[addr(XLSX, r, c + 1)] = ws[a]; delete ws[a]; }
      }
    }
    rg.e.c += 1;
    ws['!ref'] = XLSX.utils.encode_range(rg);
    if (ws['!cols']) ws['!cols'].unshift({});
    delete ws['!merges'];
  }

  function anagToOnde(XLSX, anagWb, ondeWb) {
    var A = findSheet(anagWb, XLSX, 'CODICE CAPO', 'BOVINI');
    if (!A) throw new Error('Nel file ANAGRAFE non trovo il foglio BOVINI con la colonna "CODICE CAPO".');
    var O = findSheet(ondeWb, XLSX, 'CODICE CAPO', 'BOVINI');
    if (!O) throw new Error('Nel file ONDA non trovo la riga di intestazione con "CODICE CAPO".');
    if (O.cols[0].name) { shiftRight(O.ws, XLSX); O = findSheet(ondeWb, XLSX, 'CODICE CAPO', 'BOVINI'); }
    var ATT = findSheet(anagWb, XLSX, 'AZIENDA CODICE', "DATI DELL'ATTIVITA");

    // Abbinamento colonne: colonna anagrafe -> colonna onde
    var map = [], usate = {};
    GIALLE_BOVINI.forEach(function (p) {
      var ca = colOf(A, p[0]), co = colOf(O, p[1]);
      if (ca >= 0 && co >= 0 && !usate[co]) { map.push({ from: ca, to: co, label: p[0] + (p[0] !== p[1] ? ' → ' + p[1] : '') }); usate[co] = true; }
    });

    // Valori fissi dal foglio DATI DELL'ATTIVITA (prima riga dati; per CODICE FISCALE la prima colonna con quel nome)
    var fissi = [];
    if (ATT) {
      GIALLE_ATTIVITA.forEach(function (name) {
        var ca = colOf(ATT, name), co = colOf(O, name);
        if (ca < 0 || co < 0 || usate[co]) return;
        var cell = null;
        for (var r = ATT.headerRow + 1; r <= ATT.range.e.r && !cell; r++) cell = cloneCell(ATT.ws[addr(XLSX, r, ca)]);
        if (cell) { fissi.push({ to: co, cell: cell, label: name + " (dati dell'attività)" }); usate[co] = true; }
      });
    }

    // Colonna A: sempre "CAR-IT" su ogni riga compilata
    var costanti = [{ to: O.range.s.c, cell: { t: 's', v: COLONNA_A, w: COLONNA_A }, label: 'colonna A = ' + COLONNA_A }];
    usate[O.range.s.c] = true;

    // Righe dell'anagrafe nell'ordine del file
    var keyA = colOf(A, 'CODICE CAPO');
    var rows = [];
    for (var r = A.headerRow + 1; r <= A.range.e.r; r++) {
      if (cellText(A.ws[addr(XLSX, r, keyA)])) rows.push(r);
    }

    // Svuoto le righe dati del file onda e scrivo quelle nuove
    var ws = O.ws, lastCol = O.range.e.c;
    for (var rr = O.headerRow + 1; rr <= O.range.e.r; rr++)
      for (var c = O.range.s.c; c <= lastCol; c++) delete ws[addr(XLSX, rr, c)];

    rows.forEach(function (ar, i) {
      var orow = O.headerRow + 1 + i;
      costanti.concat(fissi).forEach(function (f) { ws[addr(XLSX, orow, f.to)] = cloneCell(f.cell); });
      map.forEach(function (m) {
        var cell = cloneCell(A.ws[addr(XLSX, ar, m.from)]);
        if (cell) ws[addr(XLSX, orow, m.to)] = cell;
      });
    });
    var lastRow = O.headerRow + rows.length;
    ws['!ref'] = XLSX.utils.encode_range({ s: O.range.s, e: { r: Math.max(lastRow, O.headerRow), c: lastCol } });
    if (ws['!autofilter']) ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: O.headerRow, c: O.range.s.c }, e: { r: lastRow, c: lastCol } }) };

    var vuote = O.cols.filter(function (col) { return col.name && !usate[col.c] && SCARTA.indexOf(col.name) < 0; }).map(function (col) { return col.name; });

    // Elimino le colonne da scartare (da destra verso sinistra per non sballare gli indici)
    var scartate = [];
    O.cols.slice().reverse().forEach(function (col) {
      if (SCARTA.indexOf(col.name) >= 0) { deleteCol(ws, XLSX, col.c); scartate.unshift(col.name); }
    });
    if (ws['!autofilter']) delete ws['!autofilter'];
    O = findHeader(ws, XLSX, 'CODICE CAPO');

    var header = O.cols.map(function (col) { return col.name || ''; });
    var preview = rows.map(function (ar, i) {
      return O.cols.map(function (col) {
        var cell = ws[addr(XLSX, O.headerRow + 1 + i, col.c)];
        return cell ? (cell.w != null ? cell.w : cellText(cell)) : '';
      });
    });
    return {
      wb: ondeWb,
      report: {
        capi: rows.length, sheet: O.name,
        abbinate: map.map(function (m) { return m.label; }).concat(fissi.map(function (f) { return f.label; })),
        costanti: costanti.map(function (k) { return k.label; }),
        vuote: vuote, scartate: scartate, header: header, rows: preview
      }
    };
  }

  // Intestazione di un file onda vuoto (come l'export BDN; le colonne U-Z vengono poi eliminate)
  var INTESTAZIONE_ONDA = ['', 'AZIENDA CODICE', 'CODICE FISCALE', 'CODICE SPECIE', 'CODICE CAPO', 'CODICE_ELETTRONICO',
    'RAZZA CODICE', 'SESSO', 'CODICE MADRE', 'MOTIVO INGRESSO', 'DATA NASCITA', 'DATA INGRESSO',
    'CODICE AZIENDA PROVENIENZA', 'CODICE FISCALE PROVENIENZA', 'SPECIE AZIENDA PROVENIENZA', 'NUMERO MODELLO INGRESSO',
    'DATA MODELLO INGRESSO', 'MOTIVO USCITA', 'DATA USCITA', 'CODICE AZIENDA DESTINAZIONE', 'NUMERO MODELLO USCITA',
    'DATA MODELLO USCITA', 'INTERVALLO DATE', 'DATA ESTRAZIONE DATI'];

  function nuovoOnda(XLSX) {
    var wb = XLSX.utils.book_new();
    var ws = XLSX.utils.aoa_to_sheet([INTESTAZIONE_ONDA]);
    ws['!cols'] = INTESTAZIONE_ONDA.map(function (h) { return { wch: Math.max(8, h.length + 2) }; });
    XLSX.utils.book_append_sheet(wb, ws, 'BOVINI');
    return wb;
  }

  // ---------- CONFRONTO TRA DUE FILE ----------
  // Ogni campo: nomi possibili della colonna nei vari file (VET INFO / BDN e gestionale ONDA)
  var CAMPI = [
    { id: 'ingresso', label: 'Data ingresso', nomi: ['DATA INGRESSO', 'DATAARRIVO', 'DATA ARRIVO'], data: true },
    { id: 'nascita', label: 'Data nascita', nomi: ['DATA NASCITA', 'DATANASCITA'], data: true },
    { id: 'razza', label: 'Razza', nomi: ['RAZZA CODICE', 'CODRAZZA'] },
    { id: 'sesso', label: 'Sesso', nomi: ['SESSO', 'CODSESSO'] },
    { id: 'prov', label: 'Azienda provenienza', nomi: ['CODICE AZIENDA PROVENIENZA', 'CODAZPROVENIENZA'] },
    { id: 'uscita', label: 'Data uscita', nomi: ['DATA USCITA', 'DATAUSCITA'], data: true },
    { id: 'dest', label: 'Azienda destinazione', nomi: ['CODICE AZIENDA DESTINAZIONE', 'CODAZDESTINAZIONE'] }
  ];
  var CHIAVI = ['CODICE CAPO', 'IDENTIFICATIVO', 'CODICE_CAPO', 'MATRICOLA'];
  var EXTRA = ['NUMLAVORO', 'LOCAZIONE'];

  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function dataISO(c, XLSX) {
    if (!c || c.v == null || c.v === '') return '';
    if (c.v instanceof Date) return c.v.getFullYear() + '-' + pad(c.v.getMonth() + 1) + '-' + pad(c.v.getDate());
    if (c.t === 'n') {
      var d = XLSX.SSF.parse_date_code(c.v);
      return d ? d.y + '-' + pad(d.m) + '-' + pad(d.d) : String(c.v);
    }
    var t = String(c.v).trim(), m;
    if ((m = t.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})/))) return m[3] + '-' + pad(+m[2]) + '-' + pad(+m[1]);
    if ((m = t.match(/^(\d{4})-(\d{2})-(\d{2})/))) return m[1] + '-' + m[2] + '-' + m[3];
    return t;
  }
  function mostraData(iso) { var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso); return m ? m[3] + '/' + m[2] + '/' + m[1] : iso; }

  function leggiCapi(wb, XLSX, nomeFile) {
    var h = null, key = null;
    for (var i = 0; i < wb.SheetNames.length && !h; i++) {
      for (var k = 0; k < CHIAVI.length && !h; k++) {
        var f = findHeader(wb.Sheets[wb.SheetNames[i]], XLSX, CHIAVI[k]);
        if (f) { h = f; h.name = wb.SheetNames[i]; key = CHIAVI[k]; }
      }
    }
    if (!h) throw new Error('Nel file "' + nomeFile + '" non trovo la colonna con il codice del capo (CODICE CAPO o Identificativo).');
    var cKey = colOf(h, key), cols = {};
    CAMPI.concat(EXTRA.map(function (e) { return { id: e, nomi: [e] }; })).forEach(function (f) {
      for (var j = 0; j < f.nomi.length; j++) { var c = colOf(h, f.nomi[j]); if (c >= 0) { cols[f.id] = c; break; } }
    });
    var capi = {}, ordine = [];
    for (var r = h.headerRow + 1; r <= h.range.e.r; r++) {
      var cod = cellText(h.ws[addr(XLSX, r, cKey)]).toUpperCase().replace(/\s+/g, '');
      if (!cod) continue;
      var rec = { codice: cod, riga: r + 1, valori: {} };
      Object.keys(cols).forEach(function (id) {
        var cell = h.ws[addr(XLSX, r, cols[id])];
        var campo = CAMPI.filter(function (f) { return f.id === id; })[0];
        rec.valori[id] = campo && campo.data ? dataISO(cell, XLSX) : cellText(cell).toUpperCase();
      });
      if (!capi[cod]) ordine.push(cod);
      capi[cod] = capi[cod] || rec;
    }
    return { file: nomeFile, foglio: h.name, capi: capi, ordine: ordine, campi: cols };
  }

  function confronta(XLSX, wbA, nomeA, wbB, nomeB) {
    var A = leggiCapi(wbA, XLSX, nomeA), B = leggiCapi(wbB, XLSX, nomeB);
    function info(rec, src) {
      var parti = [];
      if (rec.valori.ingresso) parti.push('ingresso ' + mostraData(rec.valori.ingresso));
      if (rec.valori.prov) parti.push('da ' + rec.valori.prov);
      if (rec.valori.razza) parti.push(rec.valori.razza);
      if (rec.valori.sesso) parti.push(rec.valori.sesso);
      if (rec.valori.nascita) parti.push('nato ' + mostraData(rec.valori.nascita));
      if (rec.valori.uscita) parti.push('uscito ' + mostraData(rec.valori.uscita));
      if (rec.valori.NUMLAVORO) parti.push('n. lavoro ' + rec.valori.NUMLAVORO);
      if (rec.valori.LOCAZIONE) parti.push('locazione ' + rec.valori.LOCAZIONE);
      return parti.join(' · ');
    }
    var soloA = A.ordine.filter(function (c) { return !B.capi[c]; }).map(function (c) { var r = A.capi[c]; return { codice: c, file: A.file, foglio: A.foglio, riga: r.riga, info: info(r) }; });
    var soloB = B.ordine.filter(function (c) { return !A.capi[c]; }).map(function (c) { var r = B.capi[c]; return { codice: c, file: B.file, foglio: B.foglio, riga: r.riga, info: info(r) }; });
    var diversi = [], mancanti = [];
    A.ordine.forEach(function (c) {
      var a = A.capi[c], b = B.capi[c];
      if (!b) return;
      CAMPI.forEach(function (f) {
        if (!(f.id in A.campi) || !(f.id in B.campi)) return;
        var va = a.valori[f.id] || '', vb = b.valori[f.id] || '';
        if (va === vb) return;
        var riga = { codice: c, campo: f.label, valA: f.data ? mostraData(va) : va, valB: f.data ? mostraData(vb) : vb, rigaA: a.riga, rigaB: b.riga };
        (va && vb ? diversi : mancanti).push(riga);
      });
    });
    return { A: { file: A.file, foglio: A.foglio, n: A.ordine.length }, B: { file: B.file, foglio: B.foglio, n: B.ordine.length },
      soloA: soloA, soloB: soloB, diversi: diversi, mancanti: mancanti };
  }

  var api = { anagToOnde: anagToOnde, nuovoOnda: nuovoOnda, confronta: confronta, GIALLE_BOVINI: GIALLE_BOVINI, GIALLE_ATTIVITA: GIALLE_ATTIVITA };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.OndeMerge = api;
})(this);
