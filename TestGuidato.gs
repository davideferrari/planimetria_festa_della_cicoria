// ============================================================
// TEST GUIDATO — sidebar: scenari multipli, dati reali sul foglio
// ============================================================

var TG_PREFISSO_NOTE_ = 'Test guidato';

/** Inserisce ripetutamente prenotazioni da 8 persone fino a errore (capienza / non riorganizzabile). */
function testGuidatoBatchRiempiFinoAPieno_() {
  var aggiunte = 0;
  var ultimoOk = '';
  var k;
  for (k = 0; k < 400; k++) {
    try {
      ultimoOk = aggiungiPrenotazione('TG Fill ' + (aggiunte + 1), String(7770000 + aggiunte), 8, 'No', TG_PREFISSO_NOTE_ + ' batch');
      aggiunte++;
    } catch (e) {
      focusFoglioPlanimetria_();
      return {
        messaggio:
          'Inserite ' +
          aggiunte +
          ' prenotazioni da 8 persone. Stop: ' +
          e.message,
        nuovoId: null
      };
    }
  }
  focusFoglioPlanimetria_();
  return {
    messaggio: 'Limite sicurezza iterazioni (' + aggiunte + ' inserimenti). Ultimo OK: ' + ultimoOk,
    nuovoId: null
  };
}

function focusFoglioPlanimetria_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_PLANIMETRIA);
  if (sh) ss.setActiveSheet(sh);
}

/**
 * Elenco scenari per la schermata iniziale della sidebar.
 */
function getElencoScenariGuidato() {
  return [
    {
      id: 'inserimenti_completo',
      titolo: '1 — Solo inserimenti fino al “completo”',
      desc: 'Pochi passi guidati, poi un unico passaggio che inserisce molte prenotazioni da 8 persone fino a esaurire i posti.'
    },
    {
      id: 'inserimenti_modifiche',
      titolo: '2 — Inserimenti e modifiche (casi tipici)',
      desc: 'Alcune prenotazioni, poi modifiche di persone e di esigenza disabili per vedere riassegnazioni.'
    },
    {
      id: 'inserimenti_modifica_cancella',
      titolo: '3 — Inserimenti, modifiche e cancellazioni',
      desc: 'Prenotazioni multiple, una modifica e una cancellazione; osserva come cambiano tavoli e planimetria.'
    },
    {
      id: 'inserimenti_modifica_cancella_completo',
      titolo: '4 — Inserimenti + modifiche + cancellazioni, poi riempimento totale',
      desc: 'Scenario misto su pochi tavoli, poi riempimento batch come nello scenario 1.'
    },
    {
      id: 'finale_libero',
      titolo: '5 — FINALE: azioni libere (random)',
      desc:
        'Cinque azioni su richiesta: inserimento 1–8 o oltre 8 persone (senza disabili), modifica o cancellazione casuali su prenotazioni create qui; Fine per chiudere. Etichetta dopo ogni click.'
    }
  ];
}

function getTestGuidatoPassiScenario_(scenarioId) {
  var n = TG_PREFISSO_NOTE_;

  if (scenarioId === 'inserimenti_completo') {
    return [
      {
        tipo: 'info',
        titolo: 'Preparazione',
        testo:
          'Consigliato: <b>copia del foglio</b> o <b>Inizializza Sistema</b> prima, così parti da tavoli vuoti. Apri il foglio <b>PLANIMETRIA</b>.',
        verifica: 'Planimetria leggibile; idealmente nessuna prenotazione attiva.'
      },
      {
        tipo: 'info',
        titolo: 'Riempimento massivo (in arrivo)',
        testo:
          'Il passo successivo può richiedere <b>decine di secondi o 1–2 minuti</b>: il sistema aggiungerà in automatico prenotazioni da <b>8 persone</b> finché non restano più posti utilizzabili.',
        verifica: 'Non chiudere la sidebar durante l’operazione.'
      },
      {
        tipo: 'batchRiempi',
        titolo: 'Esecuzione — riempimento fino a esaurimento',
        testo: 'Vengono create in sequenza prenotazioni <b>TG Fill N</b> da 8 persone ciascuna.',
        verifica: 'I tavoli passano a giallo/rosso; a fine messaggio compare il motivo di stop (capienza).'
      },
      {
        tipo: 'info',
        titolo: 'Verifica finale',
        testo: 'Controlla che non restino (o restino pochissimi) posti liberi e che la planimetria sia coerente con la saturazione.',
        verifica: 'Tavoli prevalentemente Pieno; messaggio di stop del batch comprensibile.'
      }
    ];
  }

  if (scenarioId === 'inserimenti_modifiche') {
    return [
      {
        tipo: 'info',
        titolo: 'Scenario modifiche',
        testo: 'Tre inserimenti, poi due modifiche: prima aumento persone, poi attivazione <b>disabili Sì</b> sull’ultima prenotazione (possibile cambio tavolo).',
        verifica: 'Foglio PLANIMETRIA aperto.'
      },
      {
        tipo: 'aggiungi',
        titolo: 'Inserimento A',
        testo: '<b>Demo Mod A</b> — 2 persone, senza disabili.',
        nome: 'Demo Mod A',
        telefono: '991000001',
        persone: 2,
        disabili: 'No',
        note: n,
        verifica: 'Sala Ballo: tavolo parziale 2/8.'
      },
      {
        tipo: 'aggiungi',
        titolo: 'Inserimento B',
        testo: '<b>Demo Mod B</b> — 3 persone, senza disabili.',
        nome: 'Demo Mod B',
        telefono: '991000002',
        persone: 3,
        disabili: 'No',
        note: n,
        verifica: 'Secondo tavolo occupato in Ballo (ordine di priorità zona).'
      },
      {
        tipo: 'aggiungi',
        titolo: 'Inserimento C',
        testo: '<b>Demo Mod C</b> — 2 persone.',
        nome: 'Demo Mod C',
        telefono: '991000003',
        persone: 2,
        disabili: 'No',
        note: n,
        verifica: 'Tre prenotazioni attive; planimetria con più celle colorate.'
      },
      {
        tipo: 'modifica',
        titolo: 'Modifica — più persone sulla prima',
        testo: 'Porta la <b>prima</b> prenotazione a <b>5 persone</b> (stesso tavolo se possibile).',
        idFonte: 'primo',
        nuovePersone: 5,
        disabili: 'No',
        note: n,
        verifica: 'Conteggio 5/8 sul tavolo della prima prenotazione o messaggio di spostamento.'
      },
      {
        tipo: 'modifica',
        titolo: 'Modifica — disabili sull’ultima',
        testo: 'Ultima prenotazione: imposta <b>disabili Sì</b> (2 persone). Atteso: tavolo accessibile 39–41 se possibile.',
        idFonte: 'ultimo',
        nuovePersone: 2,
        disabili: 'Si',
        note: n,
        verifica: 'Tavolo 39, 40 o 41 con simbolo accessibilità; eventuale spostamento da conferma.'
      },
      {
        tipo: 'info',
        titolo: 'Fine scenario 2',
        testo: 'Puoi confrontare PRENOTAZIONI e PLANIMETRIA. Per ripulire: cancella le righe demo o Inizializza su una copia.',
        verifica: 'Comportamento coerente con le regole disabili e zona.'
      }
    ];
  }

  if (scenarioId === 'inserimenti_modifica_cancella') {
    return [
      {
        tipo: 'info',
        titolo: 'Scenario con cancellazione',
        testo: 'Tre inserimenti, modifica sulla prima, <b>cancellazione dell’ultima</b> prenotazione inserita.',
        verifica: 'PLANIMETRIA visibile.'
      },
      {
        tipo: 'aggiungi',
        titolo: 'Inserimento X',
        testo: '<b>Demo Can X</b> — 2 persone.',
        nome: 'Demo Can X',
        telefono: '992000001',
        persone: 2,
        disabili: 'No',
        note: n,
        verifica: 'Primo tavolo occupato.'
      },
      {
        tipo: 'aggiungi',
        titolo: 'Inserimento Y',
        testo: '<b>Demo Can Y</b> — 4 persone.',
        nome: 'Demo Can Y',
        telefono: '992000002',
        persone: 4,
        disabili: 'No',
        note: n,
        verifica: 'Secondo blocco occupato.'
      },
      {
        tipo: 'aggiungi',
        titolo: 'Inserimento Z',
        testo: '<b>Demo Can Z</b> — 2 persone.',
        nome: 'Demo Can Z',
        telefono: '992000003',
        persone: 2,
        disabili: 'No',
        note: n,
        verifica: 'Tre prenotazioni attive.'
      },
      {
        tipo: 'modifica',
        titolo: 'Modifica prima prenotazione',
        testo: 'Prima prenotazione → <b>4 persone</b>.',
        idFonte: 'primo',
        nuovePersone: 4,
        disabili: 'No',
        note: n,
        verifica: 'Aggiornamento occupazione sul relativo tavolo.'
      },
      {
        tipo: 'cancella',
        titolo: 'Cancellazione ultima (Z)',
        testo: 'Viene <b>cancellata</b> l’ultima prenotazione inserita (Demo Can Z).',
        idFonte: 'ultimo',
        verifica: 'Un tavolo si libera o si riduce l’occupazione; Z non è più Confermata in PRENOTAZIONI.'
      },
      {
        tipo: 'info',
        titolo: 'Fine scenario 3',
        testo: 'Restano attive X e Y; Z risulta cancellata. Controlla il foglio PRENOTAZIONI.',
        verifica: 'Stato Cancellata sulla riga Z se la colonna è visibile.'
      }
    ];
  }

  if (scenarioId === 'inserimenti_modifica_cancella_completo') {
    return [
      {
        tipo: 'info',
        titolo: 'Scenario completo (misto + batch)',
        testo: 'Due inserimenti, una modifica, una cancellazione, poi <b>riempimento batch</b> come nello scenario 1.',
        verifica: 'Ideale foglio vuoto o copia; PLANIMETRIA aperta.'
      },
      {
        tipo: 'aggiungi',
        titolo: 'Inserimento M1',
        testo: '<b>Demo Mix 1</b> — 2 persone.',
        nome: 'Demo Mix 1',
        telefono: '993000001',
        persone: 2,
        disabili: 'No',
        note: n,
        verifica: 'Primo tavolo in Ballo.'
      },
      {
        tipo: 'aggiungi',
        titolo: 'Inserimento M2',
        testo: '<b>Demo Mix 2</b> — 3 persone.',
        nome: 'Demo Mix 2',
        telefono: '993000002',
        persone: 3,
        disabili: 'No',
        note: n,
        verifica: 'Secondo tavolo occupato.'
      },
      {
        tipo: 'modifica',
        titolo: 'Modifica Mix 1',
        testo: 'Prima prenotazione → <b>6 persone</b>.',
        idFonte: 'primo',
        nuovePersone: 6,
        disabili: 'No',
        note: n,
        verifica: 'Occupazione aggiornata su un tavolo.'
      },
      {
        tipo: 'cancella',
        titolo: 'Cancella Mix 2',
        testo: 'Cancellazione dell’<b>ultima</b> prenotazione (Mix 2) per liberare posti prima del riempimento globale.',
        idFonte: 'ultimo',
        verifica: 'Almeno un tavolo si alleggerisce.'
      },
      {
        tipo: 'info',
        titolo: 'Prima del riempimento totale',
        testo: 'Il passo successivo esegue il <b>batch</b> (molte aggiunte da 8 persone). Può richiedere più di un minuto.',
        verifica: 'Salva il foglio se necessario; non chiudere la sidebar durante il batch.'
      },
      {
        tipo: 'batchRiempi',
        titolo: 'Riempimento fino a esaurimento',
        testo: 'Prenotazioni <b>TG Fill N</b> da 8 persone fino a stop per capienza.',
        verifica: 'Planimetria sempre più satura; messaggio finale con conteggio inserimenti.'
      },
      {
        tipo: 'info',
        titolo: 'Fine scenario 4',
        testo: 'Hai attraversato modifica, cancellazione e saturazione. Ripristina con copia o Inizializza.',
        verifica: 'Sistema coerente con aspettative sui posti.'
      }
    ];
  }

  if (scenarioId === 'finale_libero') {
    return [
      {
        tipo: 'finaleLibero',
        titolo: 'Scenario FINALE — azioni libere',
        testo:
          'Ogni tasto esegue un’azione con valori <b>casuali</b> (nessun inserimento con disabili). ' +
          'Le modifiche e le cancellazioni usano le prenotazioni create in questa sessione (ID tracciati). ' +
          'Dopo ogni azione compare un riepilogo nella casella sotto.',
        verifica: 'PLANIMETRIA e PRENOTAZIONI aggiornati; ripeti le azioni che ti servono per coprire tutti i casi.'
      }
    ];
  }

  throw new Error('Scenario sconosciuto: ' + scenarioId);
}

/**
 * Prenotazioni Confermate tra gli ID passati dal client (sessione test FINALE).
 * @param {number[]} ids
 * @return {{ id: number, persone: number }[]}
 */
function tgPrenotazioniConfermateDaIds_(ids) {
  var dati = leggiTuttiIDati_();
  var out = [];
  var i;
  var j;
  for (i = 0; i < ids.length; i++) {
    var want = ids[i];
    for (j = 0; j < dati.prenotazioni.length; j++) {
      var r = dati.prenotazioni[j];
      if (r[0] === want && r[8] === 'Confermata') {
        out.push({ id: want, persone: r[3] });
        break;
      }
    }
  }
  return out;
}

/**
 * @param {string} azione 'add1_8' | 'add9plus' | 'modifica' | 'cancella'
 * @param {{ ids: number[], scenarioId: string }} contesto
 */
function testGuidatoEseguiAzioneFinale(azione, contesto) {
  contesto = contesto || {};
  var ids = contesto.ids || [];
  var note = TG_PREFISSO_NOTE_ + ' FINALE';

  if (azione === 'add1_8') {
    var p1 = 1 + Math.floor(Math.random() * 8);
    var nome1 = 'TG Fin ' + Utilities.getUuid().slice(0, 8);
    var tel1 = '997' + ('0000000' + Math.floor(Math.random() * 10000000)).slice(-7);
    var msg1 = aggiungiPrenotazione(nome1, tel1, p1, 'No', note);
    var m1 = msg1.match(/#(\d+)/);
    var nuovo1 = m1 ? parseInt(m1[1], 10) : null;
    focusFoglioPlanimetria_();
    return {
      messaggio: msg1,
      messaggioEtichetta: 'Inserimento tavolo da ' + p1 + ' persone (random 1–8)',
      nuovoId: nuovo1
    };
  }

  if (azione === 'add9plus') {
    var p2 = 9 + Math.floor(Math.random() * 16);
    var nome2 = 'TG Fin+ ' + Utilities.getUuid().slice(0, 8);
    var tel2 = '998' + ('0000000' + Math.floor(Math.random() * 10000000)).slice(-7);
    var msg2 = aggiungiPrenotazione(nome2, tel2, p2, 'No', note);
    var m2 = msg2.match(/#(\d+)/);
    var nuovo2 = m2 ? parseInt(m2[1], 10) : null;
    focusFoglioPlanimetria_();
    return {
      messaggio: msg2,
      messaggioEtichetta: 'Inserimento tavolo da ' + p2 + ' persone (>8, random)',
      nuovoId: nuovo2
    };
  }

  if (azione === 'modifica') {
    var valideM = tgPrenotazioniConfermateDaIds_(ids);
    if (valideM.length === 0) {
      throw new Error('Nessuna prenotazione attiva tracciata: inserisci prima con i tasti sopra.');
    }
    var sceltaM = valideM[Math.floor(Math.random() * valideM.length)];
    var attuale = sceltaM.persone;
    var nuoveP;
    var tent;
    for (tent = 0; tent < 12; tent++) {
      nuoveP = 1 + Math.floor(Math.random() * 24);
      if (nuoveP !== attuale) break;
    }
    var msgM = modificaPrenotazione(sceltaM.id, nuoveP, 'No', note);
    focusFoglioPlanimetria_();
    return {
      messaggio: msgM,
      messaggioEtichetta:
        'Modifica prenotazione #' + sceltaM.id + ': da ' + attuale + ' a ' + nuoveP + ' persone (random)'
    };
  }

  if (azione === 'cancella') {
    var valideC = tgPrenotazioniConfermateDaIds_(ids);
    if (valideC.length === 0) {
      throw new Error('Nessuna prenotazione da cancellare: inserisci prima con i tasti sopra.');
    }
    var sceltaC = valideC[Math.floor(Math.random() * valideC.length)];
    var msgC = cancellaPrenotazione(sceltaC.id);
    focusFoglioPlanimetria_();
    return {
      messaggio: msgC,
      messaggioEtichetta: 'Cancellazione prenotazione #' + sceltaC.id + ' (scelta casuale tra le sessioni)',
      idRimosso: sceltaC.id
    };
  }

  throw new Error('Azione FINALE sconosciuta: ' + azione);
}

/**
 * @param {number} indice
 * @param {{ ids: number[], scenarioId: string }} contesto
 */
function testGuidatoEseguiPasso(indice, contesto) {
  contesto = contesto || {};
  var ids = contesto.ids || [];
  var scenarioId = contesto.scenarioId;
  if (!scenarioId) throw new Error('Scegli uno scenario.');

  var passi = getTestGuidatoPassiScenario_(scenarioId);
  if (indice < 0 || indice >= passi.length) {
    throw new Error('Passo non valido.');
  }
  var p = passi[indice];
  var out = {
    tipo: p.tipo,
    messaggio: '',
    nuovoId: null,
    idsConsigliati: ids.slice()
  };

  if (p.tipo === 'info') {
    focusFoglioPlanimetria_();
    out.messaggio = 'Leggi le indicazioni e usa Continua.';
    return out;
  }

  if (p.tipo === 'batchRiempi') {
    var br = testGuidatoBatchRiempiFinoAPieno_();
    out.messaggio = br.messaggio;
    return out;
  }

  if (p.tipo === 'aggiungi') {
    var msg = aggiungiPrenotazione(p.nome, p.telefono, p.persone, p.disabili, p.note);
    var m = msg.match(/#(\d+)/);
    if (m) out.nuovoId = parseInt(m[1], 10);
    out.messaggio = msg;
    focusFoglioPlanimetria_();
    return out;
  }

  if (p.tipo === 'modifica') {
    var mid = p.id;
    if (p.idFonte === 'primo') {
      mid = ids[0];
      if (!mid) throw new Error('Manca ID prima prenotazione: esegui gli inserimenti precedenti.');
    } else if (p.idFonte === 'ultimo') {
      mid = ids[ids.length - 1];
      if (!mid) throw new Error('Nessuna prenotazione precedente.');
    }
    out.messaggio = modificaPrenotazione(mid, p.nuovePersone, p.disabili, p.note || '');
    focusFoglioPlanimetria_();
    return out;
  }

  if (p.tipo === 'cancella') {
    var cid = p.id;
    if (p.idFonte === 'primo') {
      cid = ids[0];
      if (!cid) throw new Error('Manca ID per la cancellazione.');
    } else if (p.idFonte === 'ultimo') {
      cid = ids[ids.length - 1];
      if (!cid) throw new Error('Nessun ID ultima prenotazione.');
    }
    out.messaggio = cancellaPrenotazione(cid);
    focusFoglioPlanimetria_();
    return out;
  }

  throw new Error('Tipo passo sconosciuto: ' + p.tipo);
}

function getTestGuidatoPassiPerClient(scenarioId) {
  var passi = getTestGuidatoPassiScenario_(scenarioId);
  return passi.map(function(p, i) {
    return {
      indice: i,
      tipo: p.tipo,
      titolo: p.titolo,
      testo: p.testo,
      verifica: p.verifica
    };
  });
}

function mostraTestGuidatoPlanimetria() {
  var html = HtmlService.createHtmlOutput(getHtmlTestGuidato_())
    .setTitle('Test guidato planimetria')
    .setWidth(500);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getHtmlTestGuidato_() {
  return (
    '<!DOCTYPE html><html><head><base target="_top">'
    + '<style>'
    + 'body { font-family: system-ui, Segoe UI, sans-serif; padding: 12px; margin: 0; font-size: 14px; line-height: 1.45; }'
    + 'h2 { font-size: 16px; margin: 0 0 8px; }'
    + 'h3 { font-size: 13px; margin: 10px 0 6px; color: #333; }'
    + '.meta { color: #555; font-size: 12px; margin-bottom: 10px; }'
    + '.sc { background: #fff; border: 1px solid #dadce0; border-radius: 10px; padding: 10px; margin-bottom: 10px; cursor: pointer; text-align: left; width: 100%; box-sizing: border-box; }'
    + '.sc:hover { background: #f8f9fa; }'
    + '.sc b { display: block; margin-bottom: 4px; color: #1565c0; }'
    + '.sc span { font-size: 12px; color: #555; }'
    + '.box { background: #f5f5f8; border-radius: 8px; padding: 12px; margin-bottom: 12px; border: 1px solid #e0e0e8; }'
    + '.verifica { background: #e8f5e9; border: 1px solid #c8e6c9; padding: 10px; border-radius: 8px; margin-top: 10px; font-size: 13px; }'
    + '.verifica strong { display: block; margin-bottom: 4px; }'
    + '.esito { font-size: 12px; color: #1b5e20; margin-top: 8px; white-space: pre-wrap; max-height: 120px; overflow-y: auto; }'
    + '.err { color: #b71c1c; }'
    + 'button.btn-go { width: 100%; padding: 12px; font-size: 15px; border: none; border-radius: 8px; cursor: pointer; margin-top: 8px; background: #1976d2; color: white; }'
    + 'button.btn-go:disabled { background: #90caf9; cursor: wait; }'
    + 'button.btn-back { width: 100%; padding: 8px; font-size: 13px; margin-top: 8px; background: #eceff1; border: none; border-radius: 8px; cursor: pointer; }'
    + '</style></head><body>'
    + '<div id="root"></div>'
    + '<script>'
    + 'var passi = []; var idx = 0; var ids = []; var scenarioId = "";'
    + 'function el(id) { return document.getElementById(id); }'
    + 'function renderPicker(scenari) {'
    + '  var h = "<h2>Scegli uno scenario</h2><p class=\\"meta\\">Dati reali sul foglio: usa una <b>copia</b> o <b>Inizializza</b> prima.</p>";'
    + '  for (var i = 0; i < scenari.length; i++) {'
    + '    var s = scenari[i];'
    + '    h += "<button type=\\"button\\" class=\\"sc\\" data-id=\\"" + s.id + "\\"><b>" + s.titolo + "</b><span>" + s.desc + "</span></button>";'
    + '  }'
    + '  el("root").innerHTML = h;'
    + '  var btns = document.querySelectorAll(".sc");'
    + '  for (var j = 0; j < btns.length; j++) {'
    + '    btns[j].onclick = function() { avviaScenario(this.getAttribute("data-id")); };'
    + '  }'
    + '}'
    + 'function avviaScenario(sid) {'
    + '  scenarioId = sid; idx = 0; ids = [];'
    + '  google.script.run.withSuccessHandler(function(list) {'
    + '    passi = list; renderStep();'
    + '  }).withFailureHandler(fail).getTestGuidatoPassiPerClient(sid);'
    + '}'
    + 'function renderFinaleLibero(p) {'
    + '  var h = "<h2>Test guidato</h2><p class=\\"meta\\">Scenario FINALE — azioni libere</p>";'
    + '  h += "<div class=\\"box\\"><strong>" + p.titolo + "</strong>";'
    + '  h += "<p>" + p.testo + "</p>";'
    + '  h += "<div class=\\"verifica\\"><strong>Cosa controllare sulla planimetria</strong>" + p.verifica + "</div>";'
    + '  h += "<p class=\\"meta\\" style=\\"margin-top:10px\\"><b>Ultima azione</b></p>";'
    + '  h += "<div id=\\"lblFinale\\" class=\\"esito\\" style=\\"max-height:200px;border:1px solid #c8e6c9;padding:8px;border-radius:6px;background:#fafafa\\"></div></div>";'
    + '  h += "<button type=\\"button\\" class=\\"btn-go\\" id=\\"tgFinAdd18\\">Inserimento tavolo 1–8 persone (random)</button>";'
    + '  h += "<button type=\\"button\\" class=\\"btn-go\\" id=\\"tgFinAdd9\\">Inserimento tavolo oltre 8 persone (random)</button>";'
    + '  h += "<button type=\\"button\\" class=\\"btn-go\\" id=\\"tgFinMod\\">Modifica prenotazione (numeri random)</button>";'
    + '  h += "<button type=\\"button\\" class=\\"btn-go\\" id=\\"tgFinCan\\">Cancellazione prenotazione (random)</button>";'
    + '  h += "<button type=\\"button\\" class=\\"btn-back\\" id=\\"tgFinEnd\\">Fine</button>";'
    + '  h += "<button type=\\"button\\" class=\\"btn-back\\" id=\\"tgFinBack\\">Cambia scenario</button>";'
    + '  el("root").innerHTML = h;'
    + '  el("tgFinAdd18").onclick = function() { eseguiFinale("add1_8"); };'
    + '  el("tgFinAdd9").onclick = function() { eseguiFinale("add9plus"); };'
    + '  el("tgFinMod").onclick = function() { eseguiFinale("modifica"); };'
    + '  el("tgFinCan").onclick = function() { eseguiFinale("cancella"); };'
    + '  el("tgFinEnd").onclick = function() { fineFinale(); };'
    + '  el("tgFinBack").onclick = function() { tornaScelta(); };'
    + '}'
    + 'function eseguiFinale(azione) {'
    + '  var ctx = { ids: ids, scenarioId: scenarioId };'
    + '  var lb = el("lblFinale"); if (lb) lb.textContent = "Attendere…";'
    + '  google.script.run.withSuccessHandler(function(r) {'
    + '    var l2 = el("lblFinale");'
    + '    if (l2) {'
    + '      var et = r.messaggioEtichetta || "";'
    + '      var det = r.messaggio || "";'
    + '      l2.textContent = et + (det && det !== et ? ("\\n\\n" + det) : (det && !et ? det : ""));'
    + '    }'
    + '    if (r.nuovoId) ids.push(r.nuovoId);'
    + '    if (r.idRimosso) ids = ids.filter(function(x) { return x !== r.idRimosso; });'
    + '  }).withFailureHandler(fail).testGuidatoEseguiAzioneFinale(azione, ctx);'
    + '}'
    + 'function fineFinale() { idx++; renderStep(); }'
    + 'function renderStep() {'
    + '  if (idx >= passi.length) {'
    + '    el("root").innerHTML = "<div class=\\"box\\"><b>Scenario completato.</b></div><button class=\\"btn-back\\" onclick=\\"tornaScelta()\\">Scegli un altro scenario</button>";'
    + '    return;'
    + '  }'
    + '  var p = passi[idx];'
    + '  if (p.tipo === "finaleLibero") { renderFinaleLibero(p); return; }'
    + '  var btnLabel = "Continua";'
    + '  if (p.tipo === "batchRiempi") btnLabel = "Esegui riempimento (attendi…)";'
    + '  else if (p.tipo !== "info") btnLabel = "Esegui questo passo";'
    + '  var h = "<h2>Test guidato</h2><p class=\\"meta\\">Scenario attivo · passo " + (idx+1) + " di " + passi.length + "</p>";'
    + '  h += "<div class=\\"box\\"><strong>" + p.titolo + "</strong>";'
    + '  h += "<p>" + p.testo + "</p>";'
    + '  h += "<div class=\\"verifica\\"><strong>Cosa controllare sulla planimetria</strong>" + p.verifica + "</div>";'
    + '  h += "<div id=\\"esito\\" class=\\"esito\\"></div></div>";'
    + '  h += "<button class=\\"btn-go\\" id=\\"btnGo\\" onclick=\\"avanti()\\">" + btnLabel + "</button>";'
    + '  h += "<button class=\\"btn-back\\" type=\\"button\\" onclick=\\"tornaScelta()\\">Cambia scenario</button>";'
    + '  el("root").innerHTML = h;'
    + '}'
    + 'function tornaScelta() {'
    + '  passi = []; idx = 0; ids = []; scenarioId = "";'
    + '  google.script.run.withSuccessHandler(renderPicker).withFailureHandler(fail).getElencoScenariGuidato();'
    + '}'
    + 'function avanti() {'
    + '  var p = passi[idx];'
    + '  var b = el("btnGo"); b.disabled = true; b.textContent = "Attendere…";'
    + '  var ctx = { ids: ids, scenarioId: scenarioId };'
    + '  if (p.tipo === "info") {'
    + '    google.script.run.withSuccessHandler(function() { idx++; renderStep(); }).withFailureHandler(fail).testGuidatoEseguiPasso(idx, ctx);'
    + '    return;'
    + '  }'
    + '  google.script.run.withSuccessHandler(function(r) {'
    + '    if (r.nuovoId) { ids.push(r.nuovoId); }'
    + '    var es = el("esito"); if (es) es.textContent = r.messaggio || "";'
    + '    idx++; renderStep();'
    + '  }).withFailureHandler(fail).testGuidatoEseguiPasso(idx, ctx);'
    + '}'
    + 'function fail(e) {'
    + '  var b = el("btnGo"); if (b) { b.disabled = false; b.textContent = "Riprova"; }'
    + '  var root = el("root");'
    + '  if (root) root.innerHTML += "<p class=\\"err\\"><b>Errore:</b> " + (e.message || e) + "</p>";'
    + '}'
    + 'google.script.run.withSuccessHandler(renderPicker).withFailureHandler(fail).getElencoScenariGuidato();'
    + '</script></body></html>'
  );
}
