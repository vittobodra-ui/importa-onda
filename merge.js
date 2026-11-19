/* Logica di unione ONDE -> ANAGRAFE. Lavora su oggetti cella di SheetJS ({t, v, w, z}). */
(function (root) {
  'use strict';

  // Nomi di colonna diversi tra i due file che indicano lo stesso dato
  var ALIAS = {
    'NUMERO MODELLO INGRESSO': 'ESTREMI MODELLO INGRESSO',
    'CODICE PAESE DESTINAZIONE': 'CODICE STATO DESTINAZIONE'
  };

  // Razze note (si arricchisce da sola con le coppie trovate nell'anagrafe)
  var RAZZE_BASE = {
    MTT: 'METICCIO / INCROCIO',
    PMT: 'PIEMONTESE',
    FRS: 'FRISONA',
    VPR: 'VALDOSTANA PEZZATA ROSSA',
    PRO: "PEZZATA ROSSA D'OROPA"
  };

  // Colonne che non contano per decidere se un capo è "aggiornato"
  var IGNORA_CONFRONTO = { 'DATA ESTRAZIONE DATI': true };

  function norm(s) {
    return String(s == null ? '' : s).toUpperCase().replace(/\s+/g, ' ').trim();
  }
  function cellText(c) {
    if (!c || c.v == null) return '';
    if (c.t === 'd') return c.v instanceof Date ? c.v.toISOString().slice(0, 10) : String(c.v);
    return String(c.v).trim();
  }
  function cloneCell(c) {
    if (!c || c.v == null || cellText(c) === '') return null;
    var n = { t: c.t, v: c.v };
    if (c.z) n.z = c.z;
    if (c.w != null) n.w = c.w;
    return n;
  }
  function strCell(s) { return { t: 's', v: s, w: s }; }

  // Trova il foglio e la riga di intestazione che contiene "CODICE CAPO"
  function locate(wb, XLSX, preferName) {
    var names = wb.SheetNames.slice();
    if (preferName) names.sort(function (a, b) { return (norm(b) === preferName) - (norm(a) === preferName); });
    for (var i = 0; i < names.length; i++) {
      var ws = wb.Sheets[names[i]];
      if (!ws || !ws['!ref']) continue;
      var rg = XLSX.utils.decode_range(ws['!ref']);
      for (var r = rg.s.r; r <= Math.min(rg.e.r, rg.s.r + 15); r++) {
        var cols = [];
        var found = -1;
        for (var c = rg.s.c; c <= rg.e.c; c++) {
          var h = norm(cellText(ws[XLSX.utils.encode_cell({ r: r, c: c })]));
          cols.push({ c: c, name: ALIAS[h] || h });
          if (h === 'CODICE CAPO') found = c;
        }
        if (found >= 0) return { name: names[i], ws: ws, headerRow: r, cols: cols, keyCol: found, range: rg };
      }
    }
    return null;
  }

  // Legge le righe dati (array di celle indicizzate per colonna)
  function readRows(loc, XLSX) {
    var rows = [];
    for (var r = loc.headerRow + 1; r <= loc.range.e.r; r++) {
      var key = cellText(loc.ws[XLSX.utils.encode_cell({ r: r, c: loc.keyCol })]);
      if (!key) continue;
      var byName = {};
      loc.cols.forEach(function (col) {
        if (col.name) byName[col.name] = loc.ws[XLSX.utils.encode_cell({ r: r, c: col.c })] || null;
      });
      rows.push({ key: key, cells: byName });
    }
    return rows;
  }

  function parseDate(c) {
    if (!c || c.v == null) return Infinity;
    if (c.t === 'n') return c.v; // seriale Excel
    if (c.v instanceof Date) return c.v.getTime() / 86400000 + 25569;
    var m = String(c.v).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (!m) return Infinity;
    return Date.UTC(+m[3], +m[2] - 1, +m[1]) / 86400000 + 25569;
  }

  /*
   * opts: { sort: bool, fillDefaults: {colName: value} | null, razze: {} }
   * Restituisce { wb, report }
   */
  function merge(XLSX, anagWb, ondeWbs, opts) {
    opts = opts || {};
    var dest = locate(anagWb, XLSX, 'BOVINI');
    if (!dest) throw new Error('Nel file ANAGRAFE non trovo un foglio con la colonna "CODICE CAPO".');

    var destNames = dest.cols.map(function (c) { return c.name; });
    var existing = readRows(dest, XLSX);
    var index = {};
    existing.forEach(function (row, i) { index[row.key] = i; existing[i].status = 'invariato'; });

    // Dizionario razze: base + quanto già presente in anagrafe + extra passati
    var razze = {};
    Object.keys(RAZZE_BASE).forEach(function (k) { razze[k] = RAZZE_BASE[k]; });
    Object.keys(opts.razze || {}).forEach(function (k) { razze[k] = opts.razze[k]; });
    existing.forEach(function (row) {
      var cod = cellText(row.cells['RAZZA CODICE']);
      var den = cellText(row.cells['RAZZA DENOMINAZIONE']);
      if (cod && den) razze[cod] = den;
    });

    var report = { nuovi: 0, aggiornati: 0, invariati: 0, totaleOnde: 0, razzeSconosciute: {}, colonneNonUsate: {}, colonneVuote: [] };

    ondeWbs.forEach(function (wb) {
      var src = locate(wb, XLSX, 'BOVINI');
      if (!src) throw new Error('In uno dei file ONDE non trovo la colonna "CODICE CAPO".');
      var common = src.cols.filter(function (c) { return c.name && destNames.indexOf(c.name) >= 0; }).map(function (c) { return c.name; });
      src.cols.forEach(function (c) { if (c.name && destNames.indexOf(c.name) < 0) report.colonneNonUsate[c.name] = true; });
      var srcRows = readRows(src, XLSX);

      srcRows.forEach(function (sr) {
        report.totaleOnde++;
        if (index.hasOwnProperty(sr.key)) {
          var row = existing[index[sr.key]];
          var changed = false;
          common.forEach(function (name) {
            if (IGNORA_CONFRONTO[name]) return;
            var nc = cloneCell(sr.cells[name]);
            if (!nc) return; // non cancello dati già presenti
            if (cellText(nc) !== cellText(row.cells[name])) { row.cells[name] = nc; changed = true; }
          });
          if (changed) {
            row.status = row.status === 'nuovo' ? 'nuovo' : 'aggiornato';
            if (sr.cells['DATA ESTRAZIONE DATI']) row.cells['DATA ESTRAZIONE DATI'] = cloneCell(sr.cells['DATA ESTRAZIONE DATI']);
          }
        } else {
          var cells = {};
          common.forEach(function (name) { cells[name] = cloneCell(sr.cells[name]); });
          var cod = cellText(cells['RAZZA CODICE']);
          if (destNames.indexOf('RAZZA DENOMINAZIONE') >= 0 && cod) {
            if (razze[cod]) cells['RAZZA DENOMINAZIONE'] = strCell(razze[cod]);
            else report.razzeSconosciute[cod] = true;
          }
          if (opts.fillDefaults) {
            Object.keys(opts.fillDefaults).forEach(function (k) {
              if (destNames.indexOf(k) >= 0 && !cells[k] && opts.fillDefaults[k] !== '') cells[k] = strCell(opts.fillDefaults[k]);
            });
          }
          index[sr.key] = existing.length;
          existing.push({ key: sr.key, cells: cells, status: 'nuovo' });
        }
      });
    });

    existing.forEach(function (r) {
      if (r.status === 'nuovo') report.nuovi++;
      else if (r.status === 'aggiornato') report.aggiornati++;
      else report.invariati++;
    });

    if (opts.sort !== false) {
      existing.sort(function (a, b) {
        var da = parseDate(a.cells['DATA INGRESSO']), db = parseDate(b.cells['DATA INGRESSO']);
        if (da !== db) return da - db;
        return a.key < b.key ? -1 : a.key > b.key ? 1 : 0;
      });
    }

    // Riscrivo il foglio: cancello le vecchie righe dati e scrivo le nuove
    var ws = dest.ws;
    for (var r = dest.headerRow + 1; r <= dest.range.e.r; r++) {
      for (var c = dest.range.s.c; c <= dest.range.e.c; c++) delete ws[XLSX.utils.encode_cell({ r: r, c: c })];
    }
    existing.forEach(function (row, i) {
      var rr = dest.headerRow + 1 + i;
      dest.cols.forEach(function (col) {
        var cell = col.name ? row.cells[col.name] : null;
        if (cell && cell.v != null) ws[XLSX.utils.encode_cell({ r: rr, c: col.c })] = cell;
      });
    });
    var lastRow = Math.max(dest.headerRow, dest.headerRow + existing.length);
    ws['!ref'] = XLSX.utils.encode_range({ s: dest.range.s, e: { r: lastRow, c: dest.range.e.c } });
    if (ws['!autofilter']) ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: dest.headerRow, c: dest.range.s.c }, e: { r: lastRow, c: dest.range.e.c } }) };

    report.razzeSconosciute = Object.keys(report.razzeSconosciute);
    report.colonneNonUsate = Object.keys(report.colonneNonUsate);
    report.razze = razze;
    report.sheet = dest.name;
    report.header = destNames;
    report.rows = existing.map(function (row) {
      return { status: row.status, values: destNames.map(function (n) { return n ? (row.cells[n] ? (row.cells[n].w != null ? row.cells[n].w : cellText(row.cells[n])) : '') : ''; }) };
    });
    return { wb: anagWb, report: report };
  }

  var api = { merge: merge, locate: locate, ALIAS: ALIAS, RAZZE_BASE: RAZZE_BASE };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.OndeMerge = api;
})(this);
