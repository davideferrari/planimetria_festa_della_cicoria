// ============================================================
// TEST AUTOMATICI (in memoria, dati statici 62 tavoli)
// Esegui da Editor: eseguiTestFestaAutomatici_()
// oppure dal menu Gestione Fiera → Esegui test automatici (Logger)
// ============================================================

function assertVero_(condizione, messaggio) {
  if (!condizione) throw new Error(messaggio || 'assert fallita');
}

function sommaPersoneAttive_(prenotazioni) {
  var s = 0;
  for (var i = 0; i < prenotazioni.length; i++) {
    if (prenotazioni[i][8] !== 'Cancellata') s += prenotazioni[i][3];
  }
  return s;
}

function tuttiTavoliPieni_(tavoli) {
  for (var i = 0; i < tavoli.length; i++) {
    if (tavoli[i][6] !== 'Pieno') return false;
  }
  return true;
}

/** Grafo non diretto: ogni arco presente in entrambe le direzioni. */
function testAdiacenzaGrafoSimmetrico_() {
  var n;
  for (n = 1; n <= TOTALE_TAVOLI; n++) {
    var vicini = ADIACENZA_TAVOLI[n] || [];
    for (var j = 0; j < vicini.length; j++) {
      var v = vicini[j];
      var back = ADIACENZA_TAVOLI[v] || [];
      assertVero_(back.indexOf(n) !== -1, 'Arco ' + n + '->' + v + ' senza inverso');
    }
  }
}

/** Sala Chiosco: fila 34-38 non collegata ai tavoli 39-48. */
function testChioscoFilaBassaIsolata_() {
  var bassi = [34, 35, 36, 37, 38];
  var alti = [];
  for (var a = 39; a <= 48; a++) alti.push(a);
  var bi;
  for (bi = 0; bi < bassi.length; bi++) {
    var b = bassi[bi];
    var adj = ADIACENZA_TAVOLI[b] || [];
    for (var ai = 0; ai < alti.length; ai++) {
      assertVero_(adj.indexOf(alti[ai]) === -1, 'Chiosco: ' + b + ' non deve essere adiacente a ' + alti[ai]);
    }
  }
}

/** 62 prenotazioni da 8 persone → tutti i tavoli pieni, 496 posti occupati. */
function testRiempimentoSoloInserimenti62x8_() {
  var stato = creaStatoFestaVuoto_();
  var k;
  for (k = 0; k < 62; k++) {
    aggiungiPrenotazioneInMemoria_(stato, 'Gruppo ' + k, '3' + k, 8, 'No', '');
  }
  assertVero_(stato.prenotazioni.length === 62, '62 prenotazioni');
  assertVero_(sommaPersoneAttive_(stato.prenotazioni) === 496, '496 persone');
  assertVero_(tuttiTavoliPieni_(stato.tavoli), 'tutti Pieno');
  var occ = calcolaOccupazioneMappa_(stato.prenotazioni);
  var chk = 0;
  for (var t = 1; t <= 62; t++) {
    chk += occ[t] || 0;
  }
  assertVero_(chk === 496, 'occupazione map 496');
}

/** Dopo riempimento completo, la successiva fallisce. */
function testRifiutoOltreCapacita_() {
  var stato = creaStatoFestaVuoto_();
  var k;
  for (k = 0; k < 62; k++) {
    aggiungiPrenotazioneInMemoria_(stato, 'G' + k, 'x', 8, 'No', '');
  }
  var errore = false;
  try {
    aggiungiPrenotazioneInMemoria_(stato, 'Extra', 'x', 1, 'No', '');
  } catch (e) {
    errore = true;
  }
  assertVero_(errore, 'atteso errore dopo capienza esaurita');
}

/** Prima prenotazione: zona Sala Ballo, tavolo singolo (best-fit). */
function testPrimaPrenotazioneSalaBallo_() {
  var stato = creaStatoFestaVuoto_();
  aggiungiPrenotazioneInMemoria_(stato, 'A', '1', 2, 'No', '');
  assertVero_(stato.prenotazioni[0][6] === ZONA_SALA_BALLO, 'zona Ballo');
  var lista = parseTavoli_(stato.prenotazioni[0][5]);
  assertVero_(lista.length === 1, 'singolo tavolo');
  assertVero_(lista[0] >= 1 && lista[0] <= 33, 'tavolo in Ballo');
}

/** Prenotazione con disabili → solo tavoli accessibili (39-41). */
function testDisabiliUsaTavoliAccessibili_() {
  var stato = creaStatoFestaVuoto_();
  aggiungiPrenotazioneInMemoria_(stato, 'D', '1', 2, 'Si', '');
  var lista = parseTavoli_(stato.prenotazioni[0][5]);
  assertVero_(lista.length === 1, 'un tavolo');
  assertVero_(lista[0] >= 39 && lista[0] <= 41, 'tavolo accessibile Chiosco');
}

/**
 * Inserimenti + modifiche: più prenotazioni da 2, poi aumento persone su una riga
 * (deve restare coerente e nelle regole).
 */
function testInserimentiEModifiche_() {
  var stato = creaStatoFestaVuoto_();
  var i;
  for (i = 0; i < 20; i++) {
    aggiungiPrenotazioneInMemoria_(stato, 'P' + i, 'n' + i, 2, 'No', '');
  }
  modificaPrenotazioneInMemoria_(stato, 1, 8, 'No', 'note');
  var occ = calcolaOccupazioneMappa_(stato.prenotazioni);
  assertVero_(occupazioneCoerente_(stato.tavoli, occ), 'occupazione coerente dopo modifica');
  var row = stato.prenotazioni[0];
  assertVero_(row[3] === 8, '8 persone');
  var tavoliPren = parseTavoli_(row[5]);
  assertVero_(tavoliPren.length >= 1, 'tavolo assegnato');
}

function eseguiTestFestaAutomatici() {
  var suite = [
    ['Grafo adiacenza simmetrico', testAdiacenzaGrafoSimmetrico_],
    ['Chiosco fila bassa isolata da 39-48', testChioscoFilaBassaIsolata_],
    ['Prima prenotazione Sala Ballo', testPrimaPrenotazioneSalaBallo_],
    ['Disabili → tavoli accessibili', testDisabiliUsaTavoliAccessibili_],
    ['Inserimenti + modifica', testInserimentiEModifiche_],
    ['Riempimento 62×8 persone', testRiempimentoSoloInserimenti62x8_],
    ['Rifiuto oltre capacità', testRifiutoOltreCapacita_]
  ];

  var log = [];
  var ok = 0;
  var fail = 0;
  var si;
  for (si = 0; si < suite.length; si++) {
    var nome = suite[si][0];
    var fn = suite[si][1];
    try {
      fn();
      log.push('OK  ' + nome);
      ok++;
    } catch (e) {
      log.push('FAIL ' + nome + ': ' + e.message);
      fail++;
    }
  }

  var report = log.join('\n');
  Logger.log(report);
  Logger.log('--- Risultato: ' + ok + ' ok, ' + fail + ' falliti ---');
  return { ok: ok, fail: fail, report: report };
}

/** Voce menu: mostra alert compatto + dettagli in Logger. */
function eseguiTestFestaDaMenu() {
  var r = eseguiTestFestaAutomatici_();
  var ui = SpreadsheetApp.getUi();
  if (r.fail === 0) {
    ui.alert('Test automatici', 'Tutti superati (' + r.ok + '). Vedi Logger (Visualizza → Log) per il dettaglio.', ui.ButtonSet.OK);
  } else {
    ui.alert('Test automatici', r.fail + ' test falliti. Apri Logger (Visualizza → Log) per i messaggi.', ui.ButtonSet.OK);
  }
}
