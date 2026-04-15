// ============================================================
// TEST GUIDATO — sidebar: un passo alla volta, dati reali sul foglio
// (aggiorna PLANIMETRIA dopo ogni azione). Usa foglio di prova o backup.
// ============================================================

/**
 * Passi dello scenario. tipo: info | aggiungi | modifica
 * Per modifica: idFonte 'primo' | 'ultimo' usa contesto.ids dal client.
 */
function getTestGuidatoPassi_() {
  return [
    {
      tipo: 'info',
      titolo: 'Preparazione',
      testo: 'Apri il foglio <b>PLANIMETRIA</b> (puoi lasciare questa sidebar aperta). Dopo ogni passo il foglio si aggiorna e puoi confrontare con le indicazioni sotto.',
      verifica: 'Vedi la planimetria (anche vuota se non hai ancora prenotazioni).'
    },
    {
      tipo: 'aggiungi',
      titolo: 'Passo 1 — Due persone, senza disabili',
      testo: 'Verrà creata una prenotazione <b>Demo Guida 1</b>, 2 persone, disabili No. Atteso: zona <b>Sala Ballo</b> (priorità di riempimento).',
      nome: 'Demo Guida 1',
      telefono: '999000001',
      persone: 2,
      disabili: 'No',
      note: 'Test guidato planimetria',
      verifica: 'In Sala Ballo compare un tavolo con 2/8 occupati (colore giallo se parziale).'
    },
    {
      tipo: 'aggiungi',
      titolo: 'Passo 2 — Due persone con disabili',
      testo: 'Prenotazione <b>Demo Guida 2</b> con disabili Sì. Atteso: uno dei tavoli accessibili <b>39–41</b> in Sala Chiosco.',
      nome: 'Demo Guida 2',
      telefono: '999000002',
      persone: 2,
      disabili: 'Si',
      note: 'Test guidato planimetria',
      verifica: 'Icona sedia e numero tra 39 e 41; simbolo accessibilità sul tavolo.'
    },
    {
      tipo: 'modifica',
      titolo: 'Passo 3 — Modifica la prima prenotazione',
      testo: 'Si portano a <b>8 persone</b> la prenotazione creata al passo 1 (stesso tavolo se c\'è posto, altrimenti riassegnazione).',
      idFonte: 'primo',
      nuovePersone: 8,
      disabili: 'No',
      note: 'Test guidato',
      verifica: 'Il tavolo della prima demo mostra 8/8 o la prenotazione è stata spostata; controlla messaggio di conferma.'
    },
    {
      tipo: 'info',
      titolo: 'Fine demo',
      testo: 'Puoi cancellare le prenotazioni di prova dal foglio PRENOTAZIONI o usare <b>Inizializza Sistema</b> su una copia del file.',
      verifica: 'Hai visto la planimetria aggiornarsi passo dopo passo.'
    }
  ];
}

function focusFoglioPlanimetria_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_PLANIMETRIA);
  if (sh) ss.setActiveSheet(sh);
}

/**
 * @param {number} indice - indice passo
 * @param {Object} contesto - { ids: number[] } ID delle aggiunte in ordine (dal client)
 */
function testGuidatoEseguiPasso(indice, contesto) {
  contesto = contesto || {};
  var ids = contesto.ids || [];
  var passi = getTestGuidatoPassi_();
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
    out.messaggio = 'Leggi le indicazioni e clicca Continua quando hai verificato.';
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
      if (!mid) throw new Error('Manca l\'ID della prima prenotazione: esegui i passi 1 e 2 prima.');
    } else if (p.idFonte === 'ultimo') {
      mid = ids[ids.length - 1];
      if (!mid) throw new Error('Nessuna prenotazione precedente.');
    }
    out.messaggio = modificaPrenotazione(mid, p.nuovePersone, p.disabili, p.note || '');
    focusFoglioPlanimetria_();
    return out;
  }

  throw new Error('Tipo passo sconosciuto.');
}

function getTestGuidatoPassiPerClient() {
  var passi = getTestGuidatoPassi_();
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
    .setWidth(480);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getHtmlTestGuidato_() {
  return '<!DOCTYPE html><html><head><base target="_top">'
    + '<style>'
    + 'body { font-family: system-ui, Segoe UI, sans-serif; padding: 12px; margin: 0; font-size: 14px; line-height: 1.45; }'
    + 'h2 { font-size: 16px; margin: 0 0 10px; }'
    + '.meta { color: #555; font-size: 12px; margin-bottom: 12px; }'
    + '.box { background: #f5f5f8; border-radius: 8px; padding: 12px; margin-bottom: 12px; border: 1px solid #e0e0e8; }'
    + '.verifica { background: #e8f5e9; border: 1px solid #c8e6c9; padding: 10px; border-radius: 8px; margin-top: 10px; font-size: 13px; }'
    + '.verifica strong { display: block; margin-bottom: 4px; }'
    + '.esito { font-size: 13px; color: #1b5e20; margin-top: 8px; white-space: pre-wrap; }'
    + '.err { color: #b71c1c; }'
    + 'button { width: 100%; padding: 12px; font-size: 15px; border: none; border-radius: 8px; cursor: pointer; margin-top: 8px; }'
    + '.btn-go { background: #1976d2; color: white; }'
    + '.btn-go:disabled { background: #90caf9; cursor: wait; }'
    + '</style></head><body>'
    + '<h2>Test guidato planimetria</h2>'
    + '<p class="meta">Ogni passo modifica i dati sul foglio (come le prenotazioni vere). Dopo l\'azione si apre il foglio PLANIMETRIA.</p>'
    + '<div id="step"></div>'
    + '<button class="btn-go" id="btn" onclick="avanti()">Continua</button>'
    + '<p class="meta" id="hint">Nei passi solo testo, Continua passa al successivo. Negli altri, esegue la prenotazione sul foglio.</p>'
    + '<script>'
    + 'var passi = []; var idx = 0; var ids = [];'
    + 'function render() {'
    + '  if (idx >= passi.length) {'
    + '    document.getElementById("step").innerHTML = "<div class=\\"box\\"><b>Scenario completato.</b></div>";'
    + '    document.getElementById("btn").style.display = "none";'
    + '    document.getElementById("hint").style.display = "none";'
    + '    return;'
    + '  }'
    + '  var p = passi[idx];'
    + '  var h = "<div class=\\"box\\"><strong>" + (idx+1) + " / " + passi.length + " — " + p.titolo + "</strong>";'
    + '  h += "<p>" + p.testo + "</p>";'
    + '  h += "<div class=\\"verifica\\"><strong>Cosa controllare sulla planimetria</strong>" + p.verifica + "</div>";'
    + '  h += "<div id=\\"esito\\" class=\\"esito\\"></div></div>";'
    + '  document.getElementById("step").innerHTML = h;'
    + '  document.getElementById("btn").disabled = false;'
    + '  document.getElementById("btn").textContent = (p.tipo === "info") ? "Continua" : "Esegui questo passo";'
    + '}'
    + 'function avanti() {'
    + '  var p = passi[idx];'
    + '  var b = document.getElementById("btn"); b.disabled = true; b.textContent = "Attendere...";'
    + '  if (p.tipo === "info") {'
    + '    google.script.run.withSuccessHandler(function() { idx++; render(); }).withFailureHandler(fail)'
    + '      .testGuidatoEseguiPasso(idx, { ids: ids });'
    + '    return;'
    + '  }'
    + '  google.script.run.withSuccessHandler(function(r) {'
    + '    if (r.nuovoId) { ids.push(r.nuovoId); }'
    + '    var el = document.getElementById("esito"); if (!el) el = document.querySelector(".esito");'
    + '    if (el) el.textContent = r.messaggio || "";'
    + '    idx++; render();'
    + '  }).withFailureHandler(fail).testGuidatoEseguiPasso(idx, { ids: ids });'
    + '}'
    + 'function fail(e) { document.getElementById("btn").disabled = false; document.getElementById("btn").textContent = "Riprova";'
    + '  var d = document.getElementById("step"); d.innerHTML += "<p class=\\"err\\"><b>Errore:</b> " + (e.message || e) + "</p>"; }'
    + 'google.script.run.withSuccessHandler(function(list) { passi = list; render(); }).withFailureHandler(fail).getTestGuidatoPassiPerClient();'
    + '</script></body></html>';
}
