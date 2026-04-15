// ============================================================
// CONFIGURAZIONE - PLANIMETRIA FESTA (62 tavoli, 3 sale)
// Geometria: 33 (Sala Ballo) + 15 (Sala Chiosco) + 14 (Sala Esterna) = 62
// ============================================================
const SHEET_TAVOLI = 'TAVOLI';
const SHEET_PRENOTAZIONI = 'PRENOTAZIONI';
const SHEET_DASHBOARD = 'DASHBOARD';
const SHEET_PLANIMETRIA = 'PLANIMETRIA';
const SHEET_ISTRUZIONI = 'ISTRUZIONI';
const POSTI_PER_TAVOLO = 8;

const ZONA_SALA_BALLO = 'Sala Ballo';
const ZONA_SALA_CHIOSCO = 'Sala Chiosco';
const ZONA_SALA_ESTERNA = 'Sala Esterna';

/** Tavoli 39-41: fila Chiosco sinistra (accesso disabili). */
const TAVOLI_ACCESSIBILI = [39, 40, 41];

/** Priorità riempimento: 0 = Ballo, 1 = Chiosco, 2 = Esterna (solo se stessa scelta altrimenti). */
function prioritaRiempimentoNumeroTavolo_(num) {
  if (num >= 1 && num <= 33) return 0;
  if (num >= 34 && num <= 48) return 1;
  if (num >= 49 && num <= 62) return 2;
  return 99;
}

/**
 * Adiacenza = tavoli attaccati in orizzontale nella stessa fila o in verticale tra file
 * (blocchi separati da corridoi / muri non sono collegati; le tre sale non sono tra loro adiacenti).
 */
function buildAdiacenzaFesta_() {
  var adj = {};
  function edge(a, b) {
    if (!adj[a]) adj[a] = [];
    if (!adj[b]) adj[b] = [];
    if (adj[a].indexOf(b) === -1) adj[a].push(b);
    if (adj[b].indexOf(a) === -1) adj[b].push(a);
  }

  var r;
  // --- Sala Ballo sinistra: 6 file x 3 tavoli (1-18) ---
  for (r = 0; r < 6; r++) {
    var b = r * 3 + 1;
    edge(b, b + 1);
    edge(b + 1, b + 2);
    if (r < 5) {
      edge(b, b + 3);
      edge(b + 1, b + 4);
      edge(b + 2, b + 5);
    }
  }

  // --- Sala Ballo destra (19-33) ---
  edge(19, 20);
  edge(20, 21);
  edge(22, 23);
  edge(23, 24);
  edge(25, 26);
  edge(27, 28);
  edge(29, 30);
  edge(31, 32);
  edge(32, 33);
  edge(19, 22);
  edge(20, 23);
  edge(21, 24);
  edge(22, 25);
  edge(23, 26);
  edge(25, 27);
  edge(26, 28);
  edge(27, 29);
  edge(28, 30);
  edge(29, 31);
  edge(30, 32);

  // --- Sala Chiosco (34-48): fila 5 in basso (solo tra loro) + tre file superiori (collegate tra loro), senza archi basso↔alto ---
  edge(34, 35);
  edge(35, 36);
  edge(36, 37);
  edge(37, 38);
  edge(39, 40);
  edge(40, 41);
  edge(42, 43);
  edge(43, 44);
  edge(45, 46);
  edge(46, 47);
  edge(47, 48);
  edge(41, 42);
  edge(44, 45);

  // --- Sala Esterna (49-62): 7 file x 2 tavoli ---
  for (r = 0; r < 7; r++) {
    var a = 49 + r * 2;
    edge(a, a + 1);
    if (r < 6) {
      edge(a, a + 2);
      edge(a + 1, a + 3);
    }
  }

  return adj;
}

var ADIACENZA_TAVOLI = buildAdiacenzaFesta_();

function buildMappaPlanimetriaFesta_() {
  var m = [];
  var lane;
  var rightNums = [
    [19, 20, 21],
    [22, 23, 24],
    [25, 26],
    [27, 28],
    [29, 30],
    [31, 32, 33]
  ];
  /** Blocco dx dopo un solo corridoio in G (col 7): tavoli da col 8 in poi (H, J, L, …). */
  var rightCols = [
    [8, 10, 12],
    [8, 10, 12],
    [8, 10],
    [8, 10],
    [8, 10],
    [8, 10, 12]
  ];
  for (var lane = 0; lane < 6; lane++) {
    var top = 4 + lane * 3;
    var base = 1 + lane * 3;
    m.push({ tavolo: base, row: top, col: 1, sezione: ZONA_SALA_BALLO });
    m.push({ tavolo: base + 1, row: top, col: 3, sezione: ZONA_SALA_BALLO });
    m.push({ tavolo: base + 2, row: top, col: 5, sezione: ZONA_SALA_BALLO });
    var rn = rightNums[lane];
    var rc = rightCols[lane];
    for (var i = 0; i < rn.length; i++) {
      m.push({ tavolo: rn[i], row: top, col: rc[i], sezione: ZONA_SALA_BALLO });
    }
  }

  /**
   * Sala Chiosco: tre file verticali (3 | col sep | 3 | col sep | 4), riga separatrice,
   * sotto fila orizzontale di 5 tavoli adiacenti (34-38).
   */
  var chVertTop = 24;
  var sepCol1 = 3;
  var sepCol2 = 6;
  var u1 = [39, 40, 41];
  var u2 = [42, 43, 44];
  var u3 = [45, 46, 47, 48];
  var j2;
  for (j2 = 0; j2 < 3; j2++) {
    m.push({ tavolo: u1[j2], row: chVertTop + j2 * 3, col: 1, sezione: ZONA_SALA_CHIOSCO });
  }
  for (j2 = 0; j2 < 3; j2++) {
    m.push({ tavolo: u2[j2], row: chVertTop + j2 * 3, col: sepCol1 + 1, sezione: ZONA_SALA_CHIOSCO });
  }
  for (j2 = 0; j2 < 4; j2++) {
    m.push({ tavolo: u3[j2], row: chVertTop + j2 * 3, col: sepCol2 + 1, sezione: ZONA_SALA_CHIOSCO });
  }

  var chHoriTop = 37;
  var chBot = [34, 35, 36, 37, 38];
  var chBotC = [1, 3, 5, 7, 9];
  for (var j = 0; j < 5; j++) {
    m.push({ tavolo: chBot[j], row: chHoriTop, col: chBotC[j], sezione: ZONA_SALA_CHIOSCO });
  }

  /** Sala Esterna: prime colonne (stessa geometria 7x2), sotto il Chiosco. */
  var topE = 42;
  for (lane = 0; lane < 7; lane++) {
    top = topE + lane * 3;
    var t0 = 49 + lane * 2;
    m.push({ tavolo: t0, row: top, col: 1, sezione: ZONA_SALA_ESTERNA });
    m.push({ tavolo: t0 + 1, row: top, col: 3, sezione: ZONA_SALA_ESTERNA });
  }
  return m;
}

var MAPPA_TAVOLI = buildMappaPlanimetriaFesta_();

var TOTALE_TAVOLI = 62;

function generaRigheTavoliFesta_() {
  var rows = [];
  var n;
  for (n = 1; n <= 33; n++) {
    rows.push([ZONA_SALA_BALLO, n, POSTI_PER_TAVOLO, 0, POSTI_PER_TAVOLO, 'No', 'Libero']);
  }
  for (n = 34; n <= 48; n++) {
    var acc = (n >= 39 && n <= 41) ? 'Si' : 'No';
    rows.push([ZONA_SALA_CHIOSCO, n, POSTI_PER_TAVOLO, 0, POSTI_PER_TAVOLO, acc, 'Libero']);
  }
  for (n = 49; n <= 62; n++) {
    rows.push([ZONA_SALA_ESTERNA, n, POSTI_PER_TAVOLO, 0, POSTI_PER_TAVOLO, 'No', 'Libero']);
  }
  return rows;
}

/** Colonna 1-based -> A, B, ..., Z, AA, ... */
function colIndexToLetter_(col) {
  var s = '';
  var n = col;
  while (n > 0) {
    var rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** Rettangolo in notazione A1 da angolo sup.-sin. e dimensioni (evita getRange a 4 argomenti ambigui). */
function rangeA1FromDims_(row, col, numRows, numCols) {
  var r2 = row + numRows - 1;
  var c2 = col + numCols - 1;
  return colIndexToLetter_(col) + row + ':' + colIndexToLetter_(c2) + r2;
}

// ============================================================
// MENU PERSONALIZZATO
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('\u{1F3AA} Gestione Fiera')
    .addItem('Nuova Prenotazione', 'mostraFormPrenotazione')
    .addItem('Modifica Prenotazione', 'mostraFormModifica')
    .addItem('Trova Prenotazione', 'mostraFormRicerca')
    .addItem('Cancella Prenotazione', 'mostraFormCancellazione')
    .addItem('Sposta Prenotazione', 'mostraFormSpostamento')
    .addSeparator()
    .addItem('Consolida tavoli parziali (manuale)', 'ottimizzaTavoli')
    .addItem('Aggiorna Dashboard', 'aggiornaDashboard')
    .addSeparator()
    .addItem('Inizializza Sistema', 'inizializzaSistema')
    .addSeparator()
    .addItem('Esegui test automatici (Logger)', 'eseguiTestFestaDaMenu')
    .addItem('Test guidato planimetria (sidebar)', 'mostraTestGuidatoPlanimetria')
    .addToUi();
}

// ============================================================
// INIZIALIZZAZIONE
// ============================================================
function inizializzaSistema() {
  const ui = SpreadsheetApp.getUi();
  const risposta = ui.alert(
    'Inizializzazione',
    'Questo creerà ' + TOTALE_TAVOLI + ' tavoli (planimetria festa). I dati esistenti verranno sovrascritti. Continuare?',
    ui.ButtonSet.YES_NO
  );
  if (risposta !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- FOGLIO TAVOLI ---
  let sheetTavoli = ss.getSheetByName(SHEET_TAVOLI);
  if (!sheetTavoli) sheetTavoli = ss.insertSheet(SHEET_TAVOLI);
  sheetTavoli.clear();

  const headers = ['Zona', 'N. Tavolo', 'Posti Totali', 'Posti Occupati', 'Posti Liberi', 'Accessibile Disabili', 'Stato'];
  sheetTavoli.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');

  const tavoli = generaRigheTavoliFesta_();

  sheetTavoli.getRange(2, 1, tavoli.length, tavoli[0].length).setValues(tavoli);
  applicaFormattazioneCondizionale_(sheetTavoli);

  // --- FOGLIO PRENOTAZIONI ---
  let sheetPren = ss.getSheetByName(SHEET_PRENOTAZIONI);
  if (!sheetPren) sheetPren = ss.insertSheet(SHEET_PRENOTAZIONI);
  sheetPren.clear();
  const headersPren = ['ID', 'Nome', 'Telefono', 'N. Persone', 'Disabili', 'Tavolo', 'Zona', 'Data Prenotazione', 'Stato', 'Note'];
  sheetPren.getRange(1, 1, 1, headersPren.length).setValues([headersPren]).setFontWeight('bold');
  formattaColonnaTavoloComeTesto_(sheetPren);
  PropertiesService.getDocumentProperties().setProperty('FMT_TAVOLO_TXT_V1', '1');

  // --- DASHBOARD, PLANIMETRIA, ISTRUZIONI ---
  aggiornaDashboard();
  aggiornaPlanimetria();
  creaFoglioIstruzioni_();

  ui.alert('Sistema inizializzato con ' + TOTALE_TAVOLI + ' tavoli.\n\n' +
           'Sala Ballo: 1-33 | Chiosco: 34-48 (39-41 accessibili) | Esterna: 49-62\n\n' +
           'Totale posti: ' + (TOTALE_TAVOLI * POSTI_PER_TAVOLO));
}

// ============================================================
// FOGLIO ISTRUZIONI
// ============================================================
function creaFoglioIstruzioni_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_ISTRUZIONI);
  if (!sheet) sheet = ss.insertSheet(SHEET_ISTRUZIONI);
  sheet.clear();

  const righe = [
    ['ISTRUZIONI - COME FUNZIONA IL SISTEMA DI PRENOTAZIONI'],
    [''],
    ['PLANIMETRIA E NUMERAZIONE TAVOLI (62 tavoli totali)'],
    ['Totale geometrico dalla planimetria: 33 (Sala Ballo) + 15 (Sala Chiosco) + 14 (Sala Esterna) = 62.'],
    ['Sala Ballo = tavoli 1-33: blocco sinistro 6 file x 3 tavoli orizzontali (1-18); passaggio; blocco destro 6 file con 3+3+2+2+2+3 tavoli (19-33).'],
    ['Sala Chiosco = tavoli 34-48: sopra tre file verticali (39-41 accessibili; 42-44; 45-48) collegate tra loro; sotto fila orizzontale 34-38 (adiacenti solo in fila). Per gruppi multi-tavolo la fila in basso non è considerata adiacente alle file superiori.'],
    ['Sala Esterna = tavoli 49-62: 7 file, ciascuna con 2 tavoli affiancati orizzontalmente.'],
    ['Le tre sale non sono collegate tra loro nella mappa adiacenza: i gruppi multi-tavolo restano dentro la stessa sala.'],
    [''],
    ['PRIORITA DI RIEMPIMENTO (per tempi e ordine operativo)'],
    ['Il sistema assegna sempre prima i posti in Sala Ballo (1-33), poi Sala Chiosco (34-48), infine Sala Esterna (49-62).'],
    ['All\'interno della stessa zona, resta valida la logica "best-fit": minor spreco di posti (residuo minimo dopo l\'inserimento).'],
    ['La ricerca scandisce i tavoli in ordine di priorita di zona per ridurre i passaggi e allinearsi al flusso della festa.'],
    [''],
    ['REGOLA 1: CAPACITA TAVOLI'],
    ['Ogni tavolo ha 8 posti a sedere.'],
    ['Di norma fino a 8 persone su un tavolo; se nessun tavolo ha abbastanza posti, su 2+ adiacenti.'],
    ['Oltre 8 persone: sempre piu tavoli adiacenti (multi-tavolo).'],
    ['Esempio: 12 persone = 2 tavoli vicini (8 + 4 posti).'],
    [''],
    ['REGOLA 2: TAVOLI ACCESSIBILI DISABILI'],
    ['I tavoli 39, 40, 41 (Sala Chiosco, fila sinistra) sono contrassegnati come accessibili.'],
    ['Se una prenotazione NON ha persone disabili, il sistema NON usa i tavoli accessibili,'],
    ['a meno che tutti gli altri tavoli siano pieni (ultima risorsa).'],
    ['Se una prenotazione HA persone disabili, il sistema assegna SOLO tavoli accessibili.'],
    ['Le persone disabili assegnate a tavoli accessibili NON vengono MAI spostate'],
    ['in tavoli non accessibili, nemmeno durante la riorganizzazione automatica.'],
    [''],
    ['REGOLA 3: ASSEGNAZIONE BEST-FIT E PRIORITA ZONA'],
    ['Dopo aver rispettato la priorita Sala Ballo > Chiosco > Esterna, il sistema sceglie il tavolo con il minor numero'],
    ['di posti residui dopo l\'inserimento. Questo riempie prima i tavoli gia parzialmente'],
    ['occupati, evitando di sprecare tavoli vuoti per gruppi piccoli.'],
    ['Esempio: se un tavolo ha 3 posti liberi e arriva un gruppo di 3, quello e il tavolo ideale.'],
    [''],
    ['REGOLA 4: PRENOTAZIONI MULTI-TAVOLO'],
    ['Oltre 8 persone servono sempre piu tavoli adiacenti. Anche gruppi fino a 8 possono essere'],
    ['messi su 2 (o piu) tavoli vicini se nessun tavolo singolo ha abbastanza posti liberi.'],
    ['Quando il gruppo e piu grande di 8, il sistema cerca un blocco di tavoli adiacenti'],
    ['(vicini tra loro) con abbastanza posti liberi totali.'],
    ['I tavoli devono essere connessi nella mappa di adiacenza per garantire che il gruppo'],
    ['sia seduto vicino. Il gruppo riempie i tavoli nell\'ordine scritto (1;2;3): prima fino a 8 sul primo,'],
    ['poi sul secondo, ecc. Non si salta un tavolo pieno per sedere solo sui successivi (non sarebbero un blocco unico).'],
    ['Se due prenotazioni risultano incompatibili, i conteggi mostrano overflow (es. 9/8): serve spostare una prenotazione.'],
    ['Nel foglio PRENOTAZIONI, piu tavoli si scrivono con punto e virgola (es. 2;3), mai con virgola.'],
    ['Con locale italiano, "1,2" nel foglio diventa il numero 1,2 e i posti occupati risultano sballati (es. 9/8 su un tavolo).'],
    [''],
    ['REGOLA 5: AUTO-RIORGANIZZAZIONE'],
    ['Se nessun tavolo (o blocco di tavoli) ha abbastanza posti liberi, il sistema tenta'],
    ['automaticamente di spostare prenotazioni esistenti per liberare spazio.'],
    ['Funziona sia per prenotazioni singole che multi-tavolo:'],
    ['- Singolo tavolo: sposta prenotazioni piccole per liberare un tavolo intero.'],
    ['- Multi-tavolo: sposta prenotazioni (anche multi-tavolo) verso altri tavoli/blocchi'],
    ['  per liberare un blocco adiacente adatto al nuovo gruppo.'],
    ['- Riordino: per prenotazioni gia su piu tavoli (es. 3;4), puo cambiare l\'ordine (4;3)'],
    ['  per ridistribuire le persone sui tavoli senza cambiare il gruppo; a volte serve un piccolo'],
    ['  spostamento (es. 2 persone) per evitare di superare 8 posti su un tavolo.'],
    ['Esempio: se due tavoli adiacenti sono occupati da un gruppo grande e serve liberare uno per disabili,'],
    ['il sistema puo spostare il gruppo su altri tavoli vicini nella stessa sala.'],
    ['Lo spostamento avviene SOLO se:'],
    ['- Le persone da spostare trovano posto in un altro tavolo (o blocco adiacente)'],
    ['- Le persone disabili NON vengono spostate fuori da tavoli accessibili'],
    ['- Lo spostamento libera abbastanza posti per il nuovo gruppo'],
    ['L\'operatore viene informato nel messaggio di conferma degli spostamenti effettuati.'],
    [''],
    ['REGOLA 6: OTTIMIZZAZIONE MANUALE'],
    ['Dal menu "Ottimizza Tavoli", il sistema analizza tutti i tavoli parzialmente occupati'],
    ['e propone di consolidare le prenotazioni per liberare tavoli interi.'],
    ['L\'operatore puo accettare o rifiutare ogni proposta di spostamento.'],
    [''],
    ['REGOLA 7: VERIFICA DISPONIBILITA'],
    ['Prima di prenotare, l\'operatore puo verificare se c\'e posto. Il sistema risponde con:'],
    ['- VERDE: posto disponibile direttamente'],
    ['- GIALLO: posto disponibile ottimizzando i tavoli (con spostamenti automatici)'],
    ['- ROSSO: impossibile accettare la prenotazione'],
    [''],
    ['REGOLA 8: MODIFICA PRENOTAZIONE'],
    ['Una prenotazione puo essere modificata (numero persone, disabili, note).'],
    ['Se il numero di persone aumenta e il tavolo attuale non basta, il sistema'],
    ['riassegna automaticamente un tavolo adeguato (o multi-tavolo se necessario).'],
    ['Se si aggiunge la necessita di accesso disabili, il sistema sposta la prenotazione'],
    ['in un tavolo accessibile.'],
    [''],
    ['NOTA: MAPPA ADIACENZA'],
    ['L\'adiacenza segue in genere file e colonne (orizzontale nella fila, verticale tra file consecutive).'],
    ['Eccezione Sala Chiosco: la fila in basso (34-38) è adiacente solo tra loro, non alle tre file superiori (corridoio).'],
    ['I corridoi tra blocchi non collegano tavoli; le tre sale sono grafi separati per i blocchi multi-tavolo.'],
  ];

  var dati = righe.map(function(r) { return [r[0] || '']; });
  sheet.getRange(1, 1, dati.length, 1).setValues(dati);

  sheet.getRange(1, 1).setFontSize(16).setFontWeight('bold');
  sheet.setColumnWidth(1, 700);

  var titoli = [];
  for (var ti = 0; ti < righe.length; ti++) {
    var s = righe[ti][0] || '';
    if (s.indexOf('REGOLA') === 0 || s.indexOf('ISTRUZIONI') === 0 || s.indexOf('PLANIMETRIA') === 0 ||
        s.indexOf('PRIORITA') === 0 || s.indexOf('NOTA:') === 0) {
      titoli.push(ti + 1);
    }
  }
  for (var tj = 0; tj < titoli.length; tj++) {
    sheet.getRange(titoli[tj], 1).setFontSize(13).setFontWeight('bold').setBackground('#e8eaf6');
  }
}

// ============================================================
// FORMATTAZIONE CONDIZIONALE
// ============================================================
function applicaFormattazioneCondizionale_(sheet) {
  const range = sheet.getRange('G2:G200');
  const regole = [];
  regole.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Libero').setBackground('#b7e1cd').setRanges([range]).build());
  regole.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Parziale').setBackground('#fce8b2').setRanges([range]).build());
  regole.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Pieno').setBackground('#f4c7c3').setRanges([range]).build());
  sheet.setConditionalFormatRules(regole);
}

// ============================================================
// UTILITA: parsing tavolo (singolo o multi)
// Il campo Tavolo: numero singolo oppure piu tavoli separati da ; (es: "2;3").
// NON usare la virgola nel foglio: in locale italiano "1,2" diventa il numero 1,2
// e le 9 persone finiscono tutte conteggiate sul tavolo 1 (9/8).
// ============================================================
function parseDueTavoliDaNumeroDecimale_(x) {
  if (!(x > 0) || x === Math.floor(x)) return null;
  var str = x.toFixed(10).replace(/\.?0+$/, '');
  var dot = str.indexOf('.');
  if (dot === -1) return null;
  var a = parseInt(str.slice(0, dot), 10);
  var d = str.slice(dot + 1).replace(/^0+/, '').replace(/0+$/, '');
  if (!d || isNaN(a) || a < 1) return null;
  var b = parseInt(d, 10);
  if (isNaN(b) || b < 1) return null;
  return [a, b];
}

function parseTavoli_(valore) {
  if (valore === '' || valore == null) return [];
  if (typeof valore === 'number' && !isNaN(valore)) {
    if (valore >= 1 && valore === Math.floor(valore)) return [valore];
    var duo = parseDueTavoliDaNumeroDecimale_(valore);
    if (duo) return duo;
    var flo = Math.floor(valore);
    return flo >= 1 ? [flo] : [];
  }
  var s = String(valore).trim();
  if (!s) return [];
  if (s.indexOf(';') !== -1) {
    return s.split(';').map(function(x) { return parseInt(x.trim(), 10); }).filter(function(x) { return !isNaN(x) && x >= 1; });
  }
  if (s.indexOf(',') !== -1) {
    return s.split(',').map(function(x) { return parseInt(x.trim(), 10); }).filter(function(x) { return !isNaN(x) && x >= 1; });
  }
  var n = parseInt(s, 10);
  return isNaN(n) || n < 1 ? [] : [n];
}

function tavoliToString_(arr) {
  return arr.join(';');
}

/** Etichetta per messaggi all'utente (virgole ok, non scritta nel foglio). */
function etichettaTavoliVisiva_(valoreOArr) {
  var arr = Array.isArray(valoreOArr) ? valoreOArr : parseTavoli_(valoreOArr);
  return arr.join(', ');
}

function formattaColonnaTavoloComeTesto_(sheetPren) {
  if (!sheetPren) return;
  var lr = sheetPren.getLastRow();
  var end = Math.max(lr, 500);
  sheetPren.getRange(2, 6, end, 6).setNumberFormat('@');
}

/**
 * Corregge celle gia salvate come numero decimale (es. 1,2 -> 1.2) e imposta testo.
 */
function riparaColonnaTavoloLocale_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetPren = ss.getSheetByName(SHEET_PRENOTAZIONI);
  if (!sheetPren) {
    SpreadsheetApp.getUi().alert('Foglio PRENOTAZIONI non trovato.');
    return;
  }
  formattaColonnaTavoloComeTesto_(sheetPren);
  var lr = sheetPren.getLastRow();
  if (lr < 2) {
    SpreadsheetApp.getUi().alert('Colonna Tavolo (F) impostata come testo. Nessuna prenotazione da controllare.');
    return;
  }
  var numCols = 10;
  var prenotazioni = sheetPren.getRange(2, 1, lr, numCols).getValues();
  var modifiche = 0;
  for (var i = 0; i < prenotazioni.length; i++) {
    var c = prenotazioni[i][5];
    if (typeof c === 'number' && c !== Math.floor(c)) {
      var p = parseTavoli_(c);
      if (p.length >= 2) {
        prenotazioni[i][5] = tavoliToString_(p);
        modifiche++;
      }
    } else if (typeof c === 'string' && c.indexOf(',') !== -1 && c.indexOf(';') === -1) {
      var p2 = parseTavoli_(c);
      if (p2.length >= 2) {
        prenotazioni[i][5] = tavoliToString_(p2);
        modifiche++;
      }
    }
  }
  if (modifiche) {
    sheetPren.getRange(2, 1, lr, numCols).setValues(prenotazioni);
  }
  PropertiesService.getDocumentProperties().setProperty('FMT_TAVOLO_TXT_V1', '1');
  SpreadsheetApp.getUi().alert('Colonna Tavolo impostata come testo.' + (modifiche ? ' Corrette ' + modifiche + ' celle multi-tavolo salvate per errore come decimali.' : ''));
}

// ============================================================
// LETTURA BATCH
// ============================================================
function leggiTuttiIDati_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetTavoli = ss.getSheetByName(SHEET_TAVOLI);
  const sheetPren = ss.getSheetByName(SHEET_PRENOTAZIONI);

  if (sheetPren) {
    var props = PropertiesService.getDocumentProperties();
    if (props.getProperty('FMT_TAVOLO_TXT_V1') !== '1') {
      formattaColonnaTavoloComeTesto_(sheetPren);
      props.setProperty('FMT_TAVOLO_TXT_V1', '1');
    }
  }

  const numTavoli = sheetTavoli.getLastRow() - 1;
  const tavoli = numTavoli > 0 ? sheetTavoli.getRange(2, 1, numTavoli, 7).getValues() : [];

  const numPren = sheetPren.getLastRow() - 1;
  const prenotazioni = numPren > 0 ? sheetPren.getRange(2, 1, numPren, 10).getValues() : [];

  return { ss: ss, sheetTavoli: sheetTavoli, sheetPren: sheetPren, tavoli: tavoli, prenotazioni: prenotazioni };
}

// ============================================================
// CALCOLO OCCUPAZIONE IN MEMORIA
// Ordine = ordine delle RIGHE nel foglio (ID crescente): prima arrivata, prima servita.
// Multi-tavolo: percorso sequenziale nella lista salvata (es. 1;2;3). Ogni tavolo
// riceve fino a 8 persone del gruppo in ordine; non si "salta" un tavolo pieno per
// mettere gli altri dopo (romperebbe l'adiacenza). Se un tavolo e gia pieno ma
// servono ancora posti, l'overflow va su quel tavolo (es. 9/8) cosi si vede il conflitto.
// ============================================================
function calcolaOccupazioneMappa_(prenotazioni) {
  return calcolaOccupazioneMappaCore_(prenotazioni, -1);
}

function distribuisciMultiRigidoConOverflow_(listaTavoli, persone, occupazione) {
  var r = persone;
  for (var t = 0; t < listaTavoli.length; t++) {
    var tid = listaTavoli[t];
    if (r <= 0) break;
    var gia = occupazione[tid] || 0;
    var free = POSTI_PER_TAVOLO - gia;
    var want = Math.min(r, POSTI_PER_TAVOLO);
    if (free <= 0) {
      occupazione[tid] = gia + r;
      r = 0;
      break;
    }
    var take = Math.min(want, free);
    occupazione[tid] = gia + take;
    r -= take;
  }
  if (r > 0) {
    var ult = listaTavoli[listaTavoli.length - 1];
    occupazione[ult] = (occupazione[ult] || 0) + r;
  }
}

/**
 * Verifica che numPersone persone possano sedersi in sequenza lungo listaTavoliOrdinata.
 * OGNI tavolo nella lista deve ricevere almeno 1 persona (altrimenti il blocco e troppo grande).
 * Se un tavolo intermedio e pieno (free=0), la catena e rotta → false.
 */
function puoCollocareMultitavolo_(numPersone, listaTavoliOrdinata, occupazione) {
  var r = numPersone;
  for (var t = 0; t < listaTavoliOrdinata.length; t++) {
    var tid = listaTavoliOrdinata[t];
    var gia = occupazione[tid] || 0;
    var free = POSTI_PER_TAVOLO - gia;
    if (free <= 0) return false;
    if (r <= 0) return false;
    var take = Math.min(r, free);
    r -= take;
  }
  return r === 0;
}

function puoCollocareBloccoAdiacente_(numPersone, percorsoBFS, occupazione) {
  if (puoCollocareMultitavolo_(numPersone, percorsoBFS, occupazione)) return percorsoBFS.slice();
  var rev = percorsoBFS.slice().reverse();
  if (puoCollocareMultitavolo_(numPersone, rev, occupazione)) return rev;
  return null;
}

function calcolaOccupazioneMappaCore_(prenotazioni, indiceEscluso) {
  var occupazione = {};
  for (var i = 0; i < prenotazioni.length; i++) {
    if (indiceEscluso >= 0 && i === indiceEscluso) continue;
    var riga = prenotazioni[i];
    if (riga[8] === 'Cancellata') continue;
    var persone = riga[3];
    if (!persone) continue;
    var listaTavoli = parseTavoli_(riga[5]);
    if (listaTavoli.length === 0) continue;
    if (listaTavoli.length === 1) {
      var tid = listaTavoli[0];
      occupazione[tid] = (occupazione[tid] || 0) + persone;
    } else {
      distribuisciMultiRigidoConOverflow_(listaTavoli, persone, occupazione);
    }
  }
  return occupazione;
}

// ============================================================
// AGGIORNA TAVOLI IN BATCH
// ============================================================
/** Aggiorna colonne occupazione/liberi/stato sulle righe tavoli (senza scrivere sul foglio). */
function aggiornaRigheTavoliDaOccupazione_(tavoli, occupazione) {
  for (var i = 0; i < tavoli.length; i++) {
    var numTavolo = tavoli[i][1];
    var postiTotali = tavoli[i][2];
    var occ = occupazione[numTavolo] || 0;
    var liberi = postiTotali - occ;
    tavoli[i][3] = occ;
    tavoli[i][4] = liberi;
    tavoli[i][6] = occ >= postiTotali ? 'Pieno' : (occ > 0 ? 'Parziale' : 'Libero');
  }
}

function aggiornaStatoTavoli_(sheetTavoli, tavoli, occupazione) {
  aggiornaRigheTavoliDaOccupazione_(tavoli, occupazione);
  sheetTavoli.getRange(2, 1, tavoli.length, tavoli[0].length).setValues(tavoli);
}

// ============================================================
// TROVA MIGLIOR TAVOLO SINGOLO - Best-fit (per <= 8 persone)
// ============================================================
function trovaMigliorTavolo_(tavoli, occupazione, numPersone, necessitaAccessibile) {
  var candidati = [];
  for (const riga of tavoli) {
    var numTavolo = riga[1];
    var postiTotali = riga[2];
    var accessibile = riga[5] === 'Si';
    var occ = occupazione[numTavolo] || 0;
    var liberi = postiTotali - occ;

    if (liberi < numPersone) continue;
    if (necessitaAccessibile && !accessibile) continue;

    candidati.push({
      tavolo: numTavolo, zona: riga[0], postiLiberi: liberi,
      residuo: liberi - numPersone, accessibile: accessibile
    });
  }
  if (candidati.length === 0) return null;

  candidati.sort(function(a, b) {
    if (!necessitaAccessibile) {
      if (a.accessibile && !b.accessibile) return 1;
      if (!a.accessibile && b.accessibile) return -1;
    }
    var pa = prioritaRiempimentoNumeroTavolo_(a.tavolo);
    var pb = prioritaRiempimentoNumeroTavolo_(b.tavolo);
    if (pa !== pb) return pa - pb;
    return a.residuo - b.residuo;
  });
  return candidati[0];
}

// ============================================================
// TROVA TAVOLI MULTIPLI ADIACENTI
//
// Cerca un blocco connesso con posti liberi >= numPersone.
// Prova piu dimensioni k: da ceil(N/8) fino a min(nTavoli, N), cosi
// gruppi <=8 possono usare 2+ tavoli vicini se nessuno ha abbastanza posti.
// ============================================================
function trovaTavoliMultipli_(tavoli, occupazione, numPersone, necessitaAccessibile) {
  var kMin = Math.max(1, Math.ceil(numPersone / POSTI_PER_TAVOLO));
  var kMax = Math.min(tavoli.length, numPersone);
  var blocchi = [];
  for (var k = kMin; k <= kMax; k++) {
    var sub = trovaTavoliMultipliPerDimensione_(tavoli, occupazione, numPersone, necessitaAccessibile, k);
    for (var s = 0; s < sub.length; s++) blocchi.push(sub[s]);
  }
  if (blocchi.length === 0) return null;

  blocchi.sort(function(a, b) {
    if (a.tavoli.length !== b.tavoli.length) return a.tavoli.length - b.tavoli.length;
    if (!necessitaAccessibile) {
      if (a.haAccessibile && !b.haAccessibile) return 1;
      if (!a.haAccessibile && b.haAccessibile) return -1;
    }
    var pa = prioritaRiempimentoNumeroTavolo_(a.tavoli[0]);
    var pb = prioritaRiempimentoNumeroTavolo_(b.tavoli[0]);
    if (pa !== pb) return pa - pb;
    return a.residuo - b.residuo;
  });
  return blocchi[0];
}

function trovaTavoliMultipliPerDimensione_(tavoli, occupazione, numPersone, necessitaAccessibile, tavoliServono) {
  var infoTavoli = {};
  for (const riga of tavoli) {
    var num = riga[1];
    var occ = occupazione[num] || 0;
    var liberi = riga[2] - occ;
    infoTavoli[num] = {
      numero: num, zona: riga[0], liberi: liberi, totali: riga[2],
      accessibile: riga[5] === 'Si'
    };
  }

  var blocchi = [];
  var numeriTavoli = Object.keys(infoTavoli).map(Number);
  numeriTavoli.sort(function(a, b) {
    var pa = prioritaRiempimentoNumeroTavolo_(a);
    var pb = prioritaRiempimentoNumeroTavolo_(b);
    if (pa !== pb) return pa - pb;
    return a - b;
  });

  for (var i = 0; i < numeriTavoli.length; i++) {
    var start = numeriTavoli[i];
    var risultatiBFS = [];
    trovaBlocchi_(start, [], infoTavoli, tavoliServono, risultatiBFS);

    for (var b = 0; b < risultatiBFS.length; b++) {
      var blocco = risultatiBFS[b];
      var postiLiberiTotali = 0;
      var haAccessibile = false;
      for (var t = 0; t < blocco.length; t++) {
        postiLiberiTotali += infoTavoli[blocco[t]].liberi;
        if (infoTavoli[blocco[t]].accessibile) haAccessibile = true;
      }

      if (postiLiberiTotali < numPersone) continue;
      if (necessitaAccessibile && !haAccessibile) continue;

      var ordineCollocazione = puoCollocareBloccoAdiacente_(numPersone, blocco, occupazione);
      if (!ordineCollocazione) continue;

      var bloccoOrdinato = blocco.slice().sort(function(a, b) { return a - b; });
      var chiave = bloccoOrdinato.join(',');
      var duplicato = false;
      for (var x = 0; x < blocchi.length; x++) {
        if (blocchi[x].chiave === chiave) { duplicato = true; break; }
      }
      if (duplicato) continue;

      blocchi.push({
        tavoli: ordineCollocazione,
        chiave: chiave,
        zona: infoTavoli[ordineCollocazione[0]].zona,
        postiLiberi: postiLiberiTotali,
        residuo: postiLiberiTotali - numPersone,
        haAccessibile: haAccessibile
      });
    }
  }
  return blocchi;
}

// Ricorsione per trovare blocchi connessi di dimensione target
function trovaBlocchi_(nodo, percorso, infoTavoli, target, risultati) {
  var nuovoPercorso = percorso.concat([nodo]);
  if (nuovoPercorso.length === target) {
    risultati.push(nuovoPercorso);
    return;
  }

  var adiacenti = ADIACENZA_TAVOLI[nodo] || [];
  for (var i = 0; i < adiacenti.length; i++) {
    var vicino = adiacenti[i];
    if (!infoTavoli[vicino]) continue;
    if (nuovoPercorso.indexOf(vicino) !== -1) continue;
    trovaBlocchi_(vicino, nuovoPercorso, infoTavoli, target, risultati);
  }
}

var MAX_PERSONE_MICRO_SPOSTA_ = 6;

function clonaPrenotazioni_(pren) {
  return pren.map(function(row) { return row.slice(); });
}

function zonaPerNumeroTavolo_(num, tavoli) {
  for (var i = 0; i < tavoli.length; i++) {
    if (tavoli[i][1] === num) return tavoli[i][0];
  }
  return '';
}

function occupazioneCoerente_(tavoli, occ) {
  for (var i = 0; i < tavoli.length; i++) {
    var n = tavoli[i][1];
    if ((occ[n] || 0) > tavoli[i][2]) return false;
  }
  return true;
}

function permutazioniListeTavoli_(arr) {
  if (arr.length <= 1) return [arr.slice()];
  var ris = [];
  for (var i = 0; i < arr.length; i++) {
    var el = arr[i];
    var rest = arr.slice(0, i).concat(arr.slice(i + 1));
    var sub = permutazioniListeTavoli_(rest);
    for (var s = 0; s < sub.length; s++) {
      ris.push([el].concat(sub[s]));
    }
  }
  return ris;
}

function estraiSpostamentiPerModifiche_(preOriginale, copy) {
  var sp = [];
  for (var i = 0; i < copy.length; i++) {
    if (copy[i][8] === 'Cancellata') continue;
    if (preOriginale[i][5] === copy[i][5] && preOriginale[i][6] === copy[i][6]) continue;
    sp.push({
      prenotazioneId: copy[i][0],
      nome: copy[i][1],
      persone: copy[i][3],
      daTavolo: preOriginale[i][5],
      aTavolo: copy[i][5],
      aZona: copy[i][6],
      rigaIndice: i
    });
  }
  return sp;
}

/**
 * Prova permutazioni dell'ordine tavoli nelle prenotazioni multi-tavolo (es. 4;3 al posto di 3;4)
 * e al massimo uno spostamento di un gruppo piccolo (<=6 persone) su un altro tavolo.
 */
function tentaRiorganizzazioneAvanzata_(tavoli, prenotazioni, occupazione, personeRichieste, necessitaAccessibile) {
  var preOriginale = clonaPrenotazioni_(prenotazioni);

  var multiRows = [];
  for (var i = 0; i < prenotazioni.length; i++) {
    if (prenotazioni[i][8] === 'Cancellata') continue;
    var L = parseTavoli_(prenotazioni[i][5]);
    if (L.length >= 2) multiRows.push({ idx: i, perms: permutazioniListeTavoli_(L) });
  }

  function provaMicroSuCopia(baseCopy) {
    for (var si = 0; si < baseCopy.length; si++) {
      if (baseCopy[si][8] === 'Cancellata') continue;
      var lista = parseTavoli_(baseCopy[si][5]);
      if (lista.length !== 1) continue;
      var ps = baseCopy[si][3];
      if (!ps || ps > MAX_PERSONE_MICRO_SPOSTA_) continue;
      var dis = baseCopy[si][4];
      var da = lista[0];
      for (var ti = 0; ti < tavoli.length; ti++) {
        var dest = tavoli[ti][1];
        if (dest === da) continue;
        if (dis === 'Si' && tavoli[ti][5] !== 'Si') continue;
        var tryCopy = clonaPrenotazioni_(baseCopy);
        tryCopy[si][5] = dest;
        tryCopy[si][6] = tavoli[ti][0];
        var occ = calcolaOccupazioneMappa_(tryCopy);
        if (!occupazioneCoerente_(tavoli, occ)) continue;
        var blocco = trovaTavoliMultipli_(tavoli, occ, personeRichieste, necessitaAccessibile);
        if (blocco) {
          return {
            blocco: blocco,
            spostamenti: estraiSpostamentiPerModifiche_(preOriginale, tryCopy)
          };
        }
      }
    }
    return null;
  }

  function provaSuCopia(copy) {
    var occ = calcolaOccupazioneMappa_(copy);
    if (!occupazioneCoerente_(tavoli, occ)) return null;
    var blocco = trovaTavoliMultipli_(tavoli, occ, personeRichieste, necessitaAccessibile);
    if (!blocco) return null;
    return {
      blocco: blocco,
      spostamenti: estraiSpostamentiPerModifiche_(preOriginale, copy)
    };
  }

  function dfsMulti(mi, copy) {
    if (mi === multiRows.length) {
      var r = provaSuCopia(copy);
      if (r) return r;
      return provaMicroSuCopia(copy);
    }
    var mr = multiRows[mi];
    for (var p = 0; p < mr.perms.length; p++) {
      var next = clonaPrenotazioni_(copy);
      var ord = mr.perms[p];
      next[mr.idx][5] = tavoliToString_(ord);
      next[mr.idx][6] = zonaPerNumeroTavolo_(ord[0], tavoli);
      var hit = dfsMulti(mi + 1, next);
      if (hit) return hit;
    }
    return null;
  }

  return dfsMulti(0, clonaPrenotazioni_(prenotazioni));
}

// ============================================================
// AUTO-RIORGANIZZAZIONE (per singolo tavolo, <= 8 persone)
// ============================================================
function tentaRiorganizzazione_(tavoli, prenotazioni, occupazione, personeRichieste, necessitaAccessibile) {
  var prenotazioniPerTavolo = {};
  for (var i = 0; i < prenotazioni.length; i++) {
    var p = prenotazioni[i];
    if (p[8] === 'Cancellata') continue;
    var listaTav = parseTavoli_(p[5]);
    for (var lt = 0; lt < listaTav.length; lt++) {
      var tav = listaTav[lt];
      if (!prenotazioniPerTavolo[tav]) prenotazioniPerTavolo[tav] = [];
      prenotazioniPerTavolo[tav].push({
        rigaIndice: i, id: p[0], nome: p[1], persone: p[3], disabili: p[4]
      });
    }
  }

  var candidatiSvuotamento = [];
  for (const riga of tavoli) {
    var numTavolo = riga[1];
    var postiTotali = riga[2];
    var accessibile = riga[5] === 'Si';
    var occ = occupazione[numTavolo] || 0;

    if (occ === 0) continue;
    if (necessitaAccessibile && !accessibile) continue;
    if (postiTotali < personeRichieste) continue;

    candidatiSvuotamento.push({
      tavolo: numTavolo, zona: riga[0], occupati: occ,
      daLiberare: personeRichieste - (postiTotali - occ),
      accessibile: accessibile,
      prenotazioni: prenotazioniPerTavolo[numTavolo] || []
    });
  }

  candidatiSvuotamento.sort(function(a, b) { return a.daLiberare - b.daLiberare; });

  for (var c = 0; c < candidatiSvuotamento.length; c++) {
    var candidato = candidatiSvuotamento[c];
    if (candidato.daLiberare <= 0) {
      return {
        tavolo: { tavolo: candidato.tavolo, zona: candidato.zona, postiLiberi: POSTI_PER_TAVOLO - candidato.occupati },
        spostamenti: []
      };
    }

    var prenDaSpostare = [].concat(candidato.prenotazioni).sort(function(a, b) { return a.persone - b.persone; });

    var personeSpostate = 0;
    var spostamentiProposti = [];
    var occupazioneSimulata = {};
    for (var k in occupazione) { occupazioneSimulata[k] = occupazione[k]; }

    for (var pi = 0; pi < prenDaSpostare.length; pi++) {
      var pren = prenDaSpostare[pi];
      if (personeSpostate >= candidato.daLiberare) break;

      var destinazione = null;
      var migliorResiduo = Infinity;
      var migliorPri = 999;

      for (const riga of tavoli) {
        var numTav = riga[1];
        if (numTav === candidato.tavolo) continue;
        var postiTot = riga[2];
        var acc = riga[5] === 'Si';
        var occSim = occupazioneSimulata[numTav] || 0;
        var libSim = postiTot - occSim;

        if (libSim < pren.persone) continue;
        if (pren.disabili === 'Si' && !acc) continue;

        var residuo = libSim - pren.persone;
        var pri = prioritaRiempimentoNumeroTavolo_(numTav);
        if (pri < migliorPri || (pri === migliorPri && residuo < migliorResiduo)) {
          migliorPri = pri;
          migliorResiduo = residuo;
          destinazione = { tavolo: numTav, zona: riga[0] };
        }
      }

      if (destinazione) {
        spostamentiProposti.push({
          prenotazioneId: pren.id, nome: pren.nome, persone: pren.persone,
          daTavolo: candidato.tavolo, aTavolo: destinazione.tavolo, aZona: destinazione.zona,
          rigaIndice: pren.rigaIndice
        });
        personeSpostate += pren.persone;
        occupazioneSimulata[candidato.tavolo] -= pren.persone;
        occupazioneSimulata[destinazione.tavolo] = (occupazioneSimulata[destinazione.tavolo] || 0) + pren.persone;
      }
    }

    if (personeSpostate >= candidato.daLiberare) {
      return {
        tavolo: { tavolo: candidato.tavolo, zona: candidato.zona, postiLiberi: personeRichieste },
        spostamenti: spostamentiProposti
      };
    }
  }

  return null;
}

// ============================================================
// AUTO-RIORGANIZZAZIONE MULTI-TAVOLO
//
// Quando trovaTavoliMultipli_ fallisce, tenta di spostare
// prenotazioni esistenti verso altri tavoli/blocchi per
// liberare un blocco adiacente adatto.
// ============================================================
function tentaRiorganizzazioneMulti_(tavoli, prenotazioni, occupazione, personeRichieste, necessitaAccessibile) {
  var kMin = Math.max(1, Math.ceil(personeRichieste / POSTI_PER_TAVOLO));
  var infoTavoli = {};
  for (var i = 0; i < tavoli.length; i++) {
    infoTavoli[tavoli[i][1]] = {
      numero: tavoli[i][1], zona: tavoli[i][0], totali: tavoli[i][2], accessibile: tavoli[i][5] === 'Si'
    };
  }

  var numeriTavoli = Object.keys(infoTavoli).map(Number);
  numeriTavoli.sort(function(a, b) {
    var pa = prioritaRiempimentoNumeroTavolo_(a);
    var pb = prioritaRiempimentoNumeroTavolo_(b);
    if (pa !== pb) return pa - pb;
    return a - b;
  });
  var kMax = Math.min(numeriTavoli.length, personeRichieste);

  var blocchiPossibili = [];
  for (var k = kMin; k <= kMax; k++) {
    for (var i = 0; i < numeriTavoli.length; i++) {
      var risultatiBFS = [];
      trovaBlocchi_(numeriTavoli[i], [], infoTavoli, k, risultatiBFS);
      for (var b = 0; b < risultatiBFS.length; b++) {
        var blocco = risultatiBFS[b].slice().sort(function(a, b) { return a - b; });
        var chiave = blocco.join(',');
        var haAcc = false;
        var capTotale = 0;
        for (var t = 0; t < blocco.length; t++) {
          if (infoTavoli[blocco[t]].accessibile) haAcc = true;
          capTotale += infoTavoli[blocco[t]].totali;
        }
        if (necessitaAccessibile && !haAcc) continue;
        if (capTotale < personeRichieste) continue;

        var dup = false;
        for (var x = 0; x < blocchiPossibili.length; x++) {
          if (blocchiPossibili[x].chiave === chiave) { dup = true; break; }
        }
        if (!dup) blocchiPossibili.push({ tavoli: blocco, chiave: chiave, capacita: capTotale });
      }
    }
  }

  blocchiPossibili.sort(function(a, b) {
    var pa = prioritaRiempimentoNumeroTavolo_(a.tavoli[0]);
    var pb = prioritaRiempimentoNumeroTavolo_(b.tavoli[0]);
    if (pa !== pb) return pa - pb;
    return a.tavoli.length - b.tavoli.length;
  });

  // Costruisci lista prenotazioni attive come unita intere
  var prenotazioniAttive = [];
  for (var i = 0; i < prenotazioni.length; i++) {
    if (prenotazioni[i][8] === 'Cancellata') continue;
    var lista = parseTavoli_(prenotazioni[i][5]);
    if (lista.length === 0) continue;
    prenotazioniAttive.push({
      rigaIndice: i, id: prenotazioni[i][0], nome: prenotazioni[i][1],
      persone: prenotazioni[i][3], disabili: prenotazioni[i][4], tavoli: lista
    });
  }

  // Per ogni blocco candidato, prova a spostare le prenotazioni che lo occupano
  // Ordina: blocchi che richiedono meno spostamenti prima
  for (var bi = 0; bi < blocchiPossibili.length; bi++) {
    var blocco = blocchiPossibili[bi];

    // Trova prenotazioni che si sovrappongono a questo blocco
    var daRicollocare = [];
    for (var pi = 0; pi < prenotazioniAttive.length; pi++) {
      var pren = prenotazioniAttive[pi];
      for (var t = 0; t < pren.tavoli.length; t++) {
        if (blocco.tavoli.indexOf(pren.tavoli[t]) !== -1) {
          daRicollocare.push(pren);
          break;
        }
      }
    }

    if (daRicollocare.length === 0) continue;

    // Simula la rimozione delle prenotazioni sovrapposte
    var occSimulata = {};
    for (var k in occupazione) occSimulata[k] = occupazione[k];

    for (var ri = 0; ri < daRicollocare.length; ri++) {
      var restanti = daRicollocare[ri].persone;
      for (var t = 0; t < daRicollocare[ri].tavoli.length; t++) {
        var rimossi = Math.min(restanti, POSTI_PER_TAVOLO);
        occSimulata[daRicollocare[ri].tavoli[t]] = Math.max(0, (occSimulata[daRicollocare[ri].tavoli[t]] || 0) - rimossi);
        restanti -= rimossi;
      }
    }

    // Verifica che il blocco ora abbia posto
    var postiLibBloc = 0;
    for (var t = 0; t < blocco.tavoli.length; t++) {
      postiLibBloc += infoTavoli[blocco.tavoli[t]].totali - (occSimulata[blocco.tavoli[t]] || 0);
    }
    if (postiLibBloc < personeRichieste) continue;

    // Riserva il blocco marcandolo come pieno cosi le prenotazioni
    // spostate non possono finire su tavoli del blocco destinazione
    var occDopoNuova = {};
    for (var k in occSimulata) occDopoNuova[k] = occSimulata[k];
    for (var t = 0; t < blocco.tavoli.length; t++) {
      occDopoNuova[blocco.tavoli[t]] = infoTavoli[blocco.tavoli[t]].totali;
    }

    // Prova a ricollocare ogni prenotazione spostata
    daRicollocare.sort(function(a, b) { return a.persone - b.persone; });
    var spostamentiProposti = [];
    var tuttiRicollocati = true;

    for (var ri = 0; ri < daRicollocare.length; ri++) {
      var pren = daRicollocare[ri];
      var prenAcc = pren.disabili === 'Si';
      var trovato = false;
      var labelDa = etichettaTavoliVisiva_(pren.tavoli);

      if (pren.persone <= POSTI_PER_TAVOLO) {
        // Cerca un singolo tavolo
        var migliorTav = null;
        var migliorRes = Infinity;
        var migliorPriR = 999;
        for (var ti = 0; ti < tavoli.length; ti++) {
          var numT = tavoli[ti][1];
          var lib = tavoli[ti][2] - (occDopoNuova[numT] || 0);
          if (lib < pren.persone) continue;
          if (prenAcc && tavoli[ti][5] !== 'Si') continue;
          var res = lib - pren.persone;
          var priR = prioritaRiempimentoNumeroTavolo_(numT);
          if (priR < migliorPriR || (priR === migliorPriR && res < migliorRes)) {
            migliorPriR = priR;
            migliorRes = res;
            migliorTav = { tavolo: numT, zona: tavoli[ti][0] };
          }
        }
        if (migliorTav) {
          spostamentiProposti.push({
            prenotazioneId: pren.id, nome: pren.nome, persone: pren.persone,
            daTavolo: labelDa, aTavolo: migliorTav.tavolo, aZona: migliorTav.zona,
            rigaIndice: pren.rigaIndice
          });
          occDopoNuova[migliorTav.tavolo] = (occDopoNuova[migliorTav.tavolo] || 0) + pren.persone;
          trovato = true;
        }
      } else {
        // Cerca un blocco multi-tavolo
        var nuovoBlocco = trovaTavoliMultipli_(tavoli, occDopoNuova, pren.persone, prenAcc);
        if (nuovoBlocco) {
          spostamentiProposti.push({
            prenotazioneId: pren.id, nome: pren.nome, persone: pren.persone,
            daTavolo: labelDa, aTavolo: tavoliToString_(nuovoBlocco.tavoli), aZona: nuovoBlocco.zona,
            rigaIndice: pren.rigaIndice
          });
          var restP = pren.persone;
          for (var t = 0; t < nuovoBlocco.tavoli.length; t++) {
            var spazio = POSTI_PER_TAVOLO - (occDopoNuova[nuovoBlocco.tavoli[t]] || 0);
            var assP = Math.min(restP, spazio);
            occDopoNuova[nuovoBlocco.tavoli[t]] = (occDopoNuova[nuovoBlocco.tavoli[t]] || 0) + assP;
            restP -= assP;
          }
          trovato = true;
        }
      }

      if (!trovato) { tuttiRicollocati = false; break; }
    }

    if (tuttiRicollocati) {
      return {
        blocco: { tavoli: blocco.tavoli, zona: infoTavoli[blocco.tavoli[0]].zona, postiLiberi: personeRichieste },
        spostamenti: spostamentiProposti
      };
    }
  }

  return null;
}

// ============================================================
// APPLICARE SPOSTAMENTI
// ============================================================
function applicaSpostamentiSoloDati_(prenotazioni, spostamenti) {
  for (var i = 0; i < spostamenti.length; i++) {
    var s = spostamenti[i];
    prenotazioni[s.rigaIndice][5] = s.aTavolo;
    prenotazioni[s.rigaIndice][6] = s.aZona;
  }
}

function applicaSpostamenti_(sheetPren, prenotazioni, spostamenti) {
  applicaSpostamentiSoloDati_(prenotazioni, spostamenti);
  sheetPren.getRange(2, 1, prenotazioni.length, prenotazioni[0].length).setValues(prenotazioni);
}

/**
 * Stato festa in memoria per test / tooling: stessi array del foglio TAVOLI e PRENOTAZIONI.
 */
function creaStatoFestaVuoto_() {
  var tavoli = generaRigheTavoliFesta_().map(function(r) { return r.slice(); });
  return { tavoli: tavoli, prenotazioni: [] };
}

/**
 * Risolve assegnazione tavolo/i per una nuova prenotazione (muta prenotazioni in caso di riorganizzazione).
 * @return {{ tavoloAssegnato: number|string, zonaAssegnata: string, messaggioExtra: string }}
 */
function risolviAssegnazioneNuovaPrenotazione_(tavoli, prenotazioni, persone, disabili) {
  var occupazione = calcolaOccupazioneMappa_(prenotazioni);
  var necessitaAccessibile = (disabili === 'Si');
  var messaggioExtra = '';
  var tavoloAssegnato = null;
  var zonaAssegnata = '';

  function dopoRiorg_() {
    occupazione = calcolaOccupazioneMappa_(prenotazioni);
    aggiornaRigheTavoliDaOccupazione_(tavoli, occupazione);
  }

  if (persone <= POSTI_PER_TAVOLO) {
    var tavolo = trovaMigliorTavolo_(tavoli, occupazione, persone, necessitaAccessibile);
    var bloccoPiccolo = null;

    if (!tavolo) {
      var riorg = tentaRiorganizzazione_(tavoli, prenotazioni, occupazione, persone, necessitaAccessibile);
      if (riorg) {
        applicaSpostamentiSoloDati_(prenotazioni, riorg.spostamenti);
        tavolo = riorg.tavolo;
        var nomiSpostati = riorg.spostamenti.map(function(s) { return s.nome + ' (da T.' + etichettaTavoliVisiva_(s.daTavolo) + ' a T.' + etichettaTavoliVisiva_(s.aTavolo) + ')'; });
        messaggioExtra = ' [Spostati: ' + nomiSpostati.join(', ') + ']';
        dopoRiorg_();
      }
    }

    if (!tavolo) {
      bloccoPiccolo = trovaTavoliMultipli_(tavoli, occupazione, persone, necessitaAccessibile);
    }

    if (!tavolo && !bloccoPiccolo) {
      var riorgMulti = tentaRiorganizzazioneMulti_(tavoli, prenotazioni, occupazione, persone, necessitaAccessibile);
      if (riorgMulti) {
        applicaSpostamentiSoloDati_(prenotazioni, riorgMulti.spostamenti);
        bloccoPiccolo = riorgMulti.blocco;
        var nomiM = riorgMulti.spostamenti.map(function(s) { return s.nome + ' (da T.' + etichettaTavoliVisiva_(s.daTavolo) + ' a T.' + etichettaTavoliVisiva_(s.aTavolo) + ')'; });
        messaggioExtra = (messaggioExtra || '') + ' [Spostati: ' + nomiM.join(', ') + ']';
        dopoRiorg_();
      }
    }

    if (!tavolo && !bloccoPiccolo) {
      var avanzata = tentaRiorganizzazioneAvanzata_(tavoli, prenotazioni, occupazione, persone, necessitaAccessibile);
      if (avanzata) {
        applicaSpostamentiSoloDati_(prenotazioni, avanzata.spostamenti);
        bloccoPiccolo = avanzata.blocco;
        var nomiA = avanzata.spostamenti.map(function(s) { return s.nome + ' (da T.' + etichettaTavoliVisiva_(s.daTavolo) + ' a T.' + etichettaTavoliVisiva_(s.aTavolo) + ')'; });
        messaggioExtra = (messaggioExtra || '') + ' [Spostati: ' + nomiA.join(', ') + ']';
        dopoRiorg_();
      }
    }

    if (tavolo) {
      tavoloAssegnato = tavolo.tavolo;
      zonaAssegnata = tavolo.zona;
    } else if (bloccoPiccolo) {
      tavoloAssegnato = tavoliToString_(bloccoPiccolo.tavoli);
      zonaAssegnata = bloccoPiccolo.zona;
    } else {
      var postiLib = tavoli.reduce(function(s, t) { return s + t[2] - (occupazione[t[1]] || 0); }, 0);
      throw new Error('Nessun tavolo (o blocco adiacente) per ' + persone + ' persone' + (necessitaAccessibile ? ' (accessibile)' : '') + '. Posti liberi: ' + postiLib);
    }
  } else {
    var blocco = trovaTavoliMultipli_(tavoli, occupazione, persone, necessitaAccessibile);

    if (!blocco) {
      var riorgMulti2 = tentaRiorganizzazioneMulti_(tavoli, prenotazioni, occupazione, persone, necessitaAccessibile);
      if (riorgMulti2) {
        applicaSpostamentiSoloDati_(prenotazioni, riorgMulti2.spostamenti);
        blocco = riorgMulti2.blocco;
        var nomiSpostati2 = riorgMulti2.spostamenti.map(function(s) { return s.nome + ' (da T.' + etichettaTavoliVisiva_(s.daTavolo) + ' a T.' + etichettaTavoliVisiva_(s.aTavolo) + ')'; });
        messaggioExtra = ' [Spostati: ' + nomiSpostati2.join(', ') + ']';
        dopoRiorg_();
      }
    }

    if (!blocco) {
      var postiLib2 = tavoli.reduce(function(s, t) { return s + t[2] - (occupazione[t[1]] || 0); }, 0);
      throw new Error('Nessun blocco di tavoli adiacenti per ' + persone + ' persone' + (necessitaAccessibile ? ' (accessibile)' : '') + '. Posti liberi totali: ' + postiLib2 + ', ma non riorganizzabili.');
    }

    tavoloAssegnato = tavoliToString_(blocco.tavoli);
    zonaAssegnata = blocco.zona;
  }

  return { tavoloAssegnato: tavoloAssegnato, zonaAssegnata: zonaAssegnata, messaggioExtra: messaggioExtra };
}

/**
 * Aggiunge una prenotazione in memoria (stesso algoritmo del menu). Aggiorna stato.tavoli.
 * @return {number} nuovo ID
 */
function aggiungiPrenotazioneInMemoria_(stato, nome, telefono, persone, disabili, note) {
  var r = risolviAssegnazioneNuovaPrenotazione_(stato.tavoli, stato.prenotazioni, persone, disabili);
  var nuovoId = 1;
  if (stato.prenotazioni.length > 0) {
    var ids = stato.prenotazioni.map(function(row) { return row[0]; }).filter(function(v) { return v; });
    if (ids.length > 0) nuovoId = Math.max.apply(null, ids) + 1;
  }
  var nuovaRiga = [nuovoId, nome, telefono, persone, disabili, r.tavoloAssegnato, r.zonaAssegnata, new Date(), 'Confermata', note || ''];
  stato.prenotazioni.push(nuovaRiga);
  var nuovaOcc = calcolaOccupazioneMappa_(stato.prenotazioni);
  aggiornaRigheTavoliDaOccupazione_(stato.tavoli, nuovaOcc);
  return nuovoId;
}

// ============================================================
// FORM NUOVA PRENOTAZIONE
// ============================================================
function mostraFormPrenotazione() {
  var html = HtmlService.createHtmlOutput(getHtmlPrenotazione_())
    .setWidth(420).setHeight(650).setTitle('Nuova Prenotazione');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getHtmlPrenotazione_() {
  return '<style>'
    + 'body { font-family: Arial, sans-serif; padding: 15px; margin: 0; }'
    + 'label { display: block; margin-top: 12px; font-weight: bold; font-size: 14px; }'
    + 'input, select { width: 100%; padding: 10px; margin-top: 4px; font-size: 14px;'
    + '  border: 1px solid #ccc; border-radius: 6px; box-sizing: border-box; }'
    + '.btn { margin-top: 12px; padding: 14px 24px; font-size: 16px; color: white;'
    + '  border: none; border-radius: 8px; cursor: pointer; width: 100%; transition: background-color 0.2s; }'
    + '.btn-ok { background-color: #4CAF50; }'
    + '.btn-ok:hover:not(:disabled) { background-color: #45a049; }'
    + '.btn-ok:disabled { background-color: #999; cursor: wait; }'
    + '.btn-verifica { background-color: #007bff; margin-top: 20px; }'
    + '.btn-verifica:hover:not(:disabled) { background-color: #0069d9; }'
    + '.btn-verifica:disabled { background-color: #999; cursor: wait; }'
    + '#risultato { margin-top: 15px; padding: 12px; border-radius: 6px; font-size: 14px; display: none; line-height: 1.5; }'
    + '.successo { background-color: #d4edda; color: #155724; border: 1px solid #c3e6cb; }'
    + '.errore { background-color: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }'
    + '.avviso { background-color: #fff3cd; color: #856404; border: 1px solid #ffeeba; }'
    + '.spinner { display: inline-block; width: 16px; height: 16px; border: 3px solid #fff; border-top-color: transparent;'
    + '  border-radius: 50%; animation: spin 0.8s linear infinite; vertical-align: middle; margin-right: 8px; }'
    + '@keyframes spin { to { transform: rotate(360deg); } }'
    + '</style>'
    + '<label>Nome e Cognome:</label>'
    + '<input type="text" id="nome" placeholder="Mario Rossi">'
    + '<label>Telefono:</label>'
    + '<input type="tel" id="telefono" placeholder="333 1234567">'
    + '<label>Numero Persone:</label>'
    + '<input type="number" id="persone" min="1" value="2">'
    + '<label>Presenza Disabili:</label>'
    + '<select id="disabili"><option value="No">No</option><option value="Si">Si - servono tavoli accessibili</option></select>'
    + '<label>Note:</label>'
    + '<input type="text" id="note" placeholder="Note aggiuntive...">'
    + '<button class="btn btn-verifica" id="btnVerifica" onclick="verifica()">VERIFICA DISPONIBILITA</button>'
    + '<button class="btn btn-ok" id="btnPrenota" onclick="prenota()">PRENOTA</button>'
    + '<div id="risultato"></div>'
    + '<script>'
    + 'function verifica() {'
    + '  var persone = parseInt(document.getElementById("persone").value);'
    + '  var disabili = document.getElementById("disabili").value;'
    + '  if (!persone || persone < 1) { mostraRisultato("Inserisci un numero di persone valido.", "errore"); return; }'
    + '  var b = document.getElementById("btnVerifica"); b.disabled = true;'
    + '  b.innerHTML = \'<span class="spinner"></span>Verifica in corso...\';'
    + '  google.script.run'
    + '    .withSuccessHandler(function(r) { b.disabled = false; b.innerHTML = "VERIFICA DISPONIBILITA";'
    + '      if (r.esito === "diretto") mostraRisultato(r.messaggio, "successo");'
    + '      else if (r.esito === "ottimizzazione") mostraRisultato(r.messaggio, "avviso");'
    + '      else mostraRisultato(r.messaggio, "errore");'
    + '    }).withFailureHandler(function(err) { b.disabled = false; b.innerHTML = "VERIFICA DISPONIBILITA"; mostraRisultato(err.message, "errore"); })'
    + '    .verificaDisponibilita(persone, disabili);'
    + '}'
    + 'function prenota() {'
    + '  var nome = document.getElementById("nome").value.trim();'
    + '  var telefono = document.getElementById("telefono").value.trim();'
    + '  var persone = parseInt(document.getElementById("persone").value);'
    + '  var disabili = document.getElementById("disabili").value;'
    + '  var note = document.getElementById("note").value.trim();'
    + '  if (!nome || !telefono) { mostraRisultato("Compila Nome e Telefono.", "errore"); return; }'
    + '  if (!persone || persone < 1) { mostraRisultato("Inserisci un numero di persone valido.", "errore"); return; }'
    + '  setLoading(true);'
    + '  google.script.run'
    + '    .withSuccessHandler(function(msg) { mostraRisultato(msg, "successo"); setLoading(false);'
    + '      document.getElementById("nome").value = ""; document.getElementById("telefono").value = "";'
    + '      document.getElementById("persone").value = "2"; document.getElementById("disabili").value = "No";'
    + '      document.getElementById("note").value = "";'
    + '    }).withFailureHandler(function(err) { mostraRisultato(err.message, "errore"); setLoading(false); })'
    + '    .aggiungiPrenotazione(nome, telefono, persone, disabili, note);'
    + '}'
    + 'function setLoading(on) { var b = document.getElementById("btnPrenota"); b.disabled = on;'
    + '  b.innerHTML = on ? \'<span class="spinner"></span>Prenotazione in corso...\' : "PRENOTA"; }'
    + 'function mostraRisultato(msg, tipo) { var d = document.getElementById("risultato");'
    + '  d.textContent = msg; d.className = tipo; d.style.display = "block"; }'
    + '</script>';
}

// ============================================================
// AGGIUNGI PRENOTAZIONE (singola o multi-tavolo)
// ============================================================
function aggiungiPrenotazione(nome, telefono, persone, disabili, note) {
  var dati = leggiTuttiIDati_();
  var assegn = risolviAssegnazioneNuovaPrenotazione_(dati.tavoli, dati.prenotazioni, persone, disabili);
  var messaggioExtra = assegn.messaggioExtra;
  var tavoloAssegnato = assegn.tavoloAssegnato;
  var zonaAssegnata = assegn.zonaAssegnata;

  var nuovoId = 1;
  if (dati.prenotazioni.length > 0) {
    var ids = dati.prenotazioni.map(function(r) { return r[0]; }).filter(function(v) { return v; });
    if (ids.length > 0) nuovoId = Math.max.apply(null, ids) + 1;
  }

  var nuovaRiga = [nuovoId, nome, telefono, persone, disabili, tavoloAssegnato, zonaAssegnata, new Date(), 'Confermata', note || ''];
  dati.sheetPren.getRange(dati.prenotazioni.length + 2, 1, 1, nuovaRiga.length).setValues([nuovaRiga]);

  var nuovePrenotazioni = dati.prenotazioni.concat([nuovaRiga]);
  var nuovaOccupazione = calcolaOccupazioneMappa_(nuovePrenotazioni);
  aggiornaStatoTavoli_(dati.sheetTavoli, dati.tavoli, nuovaOccupazione);

  aggiornaDashboard();
  aggiornaPlanimetria();

  var listaEt = parseTavoli_(tavoloAssegnato);
  var labelTavolo = listaEt.length > 1 ? 'Tavoli ' + etichettaTavoliVisiva_(listaEt) : 'Tavolo ' + etichettaTavoliVisiva_(tavoloAssegnato);
  return 'Prenotazione #' + nuovoId + ' confermata! ' + labelTavolo + ' (' + zonaAssegnata + ') - ' + persone + ' persone.' + messaggioExtra;
}

// ============================================================
// FORM MODIFICA PRENOTAZIONE
// ============================================================
function mostraFormModifica() {
  var html = HtmlService.createHtmlOutput(getHtmlModifica_())
    .setWidth(420).setHeight(550).setTitle('Modifica Prenotazione');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getHtmlModifica_() {
  return '<style>'
    + 'body { font-family: Arial, sans-serif; padding: 15px; margin: 0; }'
    + 'label { display: block; margin-top: 12px; font-weight: bold; font-size: 14px; }'
    + 'input, select { width: 100%; padding: 10px; margin-top: 4px; font-size: 14px;'
    + '  border: 1px solid #ccc; border-radius: 6px; box-sizing: border-box; }'
    + '.btn { margin-top: 16px; padding: 14px 24px; font-size: 16px; color: white;'
    + '  border: none; border-radius: 8px; cursor: pointer; width: 100%; }'
    + '.btn-modifica { background-color: #ff9800; }'
    + '.btn-modifica:hover:not(:disabled) { background-color: #e68a00; }'
    + '.btn-modifica:disabled { background-color: #999; cursor: wait; }'
    + '#risultato { margin-top: 15px; padding: 12px; border-radius: 6px; font-size: 14px; display: none; line-height: 1.5; }'
    + '.successo { background-color: #d4edda; color: #155724; border: 1px solid #c3e6cb; }'
    + '.errore { background-color: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }'
    + '.avviso { background-color: #fff3cd; color: #856404; border: 1px solid #ffeeba; }'
    + '.spinner { display: inline-block; width: 16px; height: 16px; border: 3px solid #fff; border-top-color: transparent;'
    + '  border-radius: 50%; animation: spin 0.8s linear infinite; vertical-align: middle; margin-right: 8px; }'
    + '@keyframes spin { to { transform: rotate(360deg); } }'
    + '</style>'
    + '<label>ID Prenotazione da modificare:</label>'
    + '<input type="number" id="idPren" min="1" placeholder="Es: 3">'
    + '<label>Nuovo Numero Persone:</label>'
    + '<input type="number" id="persone" min="1" placeholder="Es: 12">'
    + '<label>Presenza Disabili:</label>'
    + '<select id="disabili"><option value="No">No</option><option value="Si">Si - servono tavoli accessibili</option></select>'
    + '<label>Note:</label>'
    + '<input type="text" id="note" placeholder="Note aggiuntive...">'
    + '<button class="btn btn-modifica" id="btnModifica" onclick="modifica()">MODIFICA PRENOTAZIONE</button>'
    + '<div id="risultato"></div>'
    + '<script>'
    + 'function modifica() {'
    + '  var id = parseInt(document.getElementById("idPren").value);'
    + '  var persone = parseInt(document.getElementById("persone").value);'
    + '  var disabili = document.getElementById("disabili").value;'
    + '  var note = document.getElementById("note").value.trim();'
    + '  if (!id) { mostraRisultato("Inserisci un ID valido.", "errore"); return; }'
    + '  if (!persone || persone < 1) { mostraRisultato("Inserisci un numero di persone valido.", "errore"); return; }'
    + '  var b = document.getElementById("btnModifica"); b.disabled = true;'
    + '  b.innerHTML = \'<span class="spinner"></span>Modifica in corso...\';'
    + '  google.script.run'
    + '    .withSuccessHandler(function(msg) { mostraRisultato(msg, "successo"); b.disabled = false; b.innerHTML = "MODIFICA PRENOTAZIONE"; })'
    + '    .withFailureHandler(function(err) { mostraRisultato(err.message, "errore"); b.disabled = false; b.innerHTML = "MODIFICA PRENOTAZIONE"; })'
    + '    .modificaPrenotazione(id, persone, disabili, note);'
    + '}'
    + 'function mostraRisultato(msg, tipo) { var d = document.getElementById("risultato");'
    + '  d.textContent = msg; d.className = tipo; d.style.display = "block"; }'
    + '</script>';
}

/**
 * Modifica prenotazione in memoria (stesso algoritmo del form). Aggiorna stato.tavoli.
 * @return {string} messaggio come modificaPrenotazione
 */
function modificaPrenotazioneInMemoria_(stato, id, nuovePersone, disabili, note) {
  var rigaPren = -1;
  var prenotazione = null;
  var i;
  for (i = 0; i < stato.prenotazioni.length; i++) {
    if (stato.prenotazioni[i][0] === id && stato.prenotazioni[i][8] === 'Confermata') {
      prenotazione = stato.prenotazioni[i];
      rigaPren = i;
      break;
    }
  }

  if (!prenotazione) throw new Error('Prenotazione #' + id + ' non trovata o non attiva.');

  var vecchioTavolo = prenotazione[5];
  var necessitaAccessibile = (disabili === 'Si');

  stato.prenotazioni[rigaPren][3] = nuovePersone;
  stato.prenotazioni[rigaPren][4] = disabili;
  stato.prenotazioni[rigaPren][9] = note;

  var serveCambioTavolo = false;
  var vecchiTavoli = parseTavoli_(vecchioTavolo);
  var occSenza = calcolaOccupazioneSenza_(stato.prenotazioni, rigaPren);

  if (nuovePersone <= POSTI_PER_TAVOLO && vecchiTavoli.length === 1) {
    var liberiVecchio = stato.tavoli.reduce(function(lib, t) {
      if (t[1] === vecchiTavoli[0]) return t[2] - (occSenza[t[1]] || 0);
      return lib;
    }, 0);

    var vecchioAccessibile = false;
    for (var t = 0; t < stato.tavoli.length; t++) {
      if (stato.tavoli[t][1] === vecchiTavoli[0]) { vecchioAccessibile = stato.tavoli[t][5] === 'Si'; break; }
    }

    if (liberiVecchio < nuovePersone) serveCambioTavolo = true;
    if (necessitaAccessibile && !vecchioAccessibile) serveCambioTavolo = true;
  } else if (nuovePersone > POSTI_PER_TAVOLO || vecchiTavoli.length > 1) {
    serveCambioTavolo = true;
  }

  var messaggioCambio = '';

  if (serveCambioTavolo) {
    if (nuovePersone <= POSTI_PER_TAVOLO) {
      var nuovoTavolo = trovaMigliorTavolo_(stato.tavoli, occSenza, nuovePersone, necessitaAccessibile);
      var bloccoPiccoloM = null;
      if (!nuovoTavolo) {
        var riorg = tentaRiorganizzazione_(stato.tavoli, stato.prenotazioni, occSenza, nuovePersone, necessitaAccessibile);
        if (riorg) {
          applicaSpostamentiSoloDati_(stato.prenotazioni, riorg.spostamenti);
          nuovoTavolo = riorg.tavolo;
          var nomiSp = riorg.spostamenti.map(function(s) { return s.nome + ' (da T.' + etichettaTavoliVisiva_(s.daTavolo) + ' a T.' + etichettaTavoliVisiva_(s.aTavolo) + ')'; });
          messaggioCambio += ' [Spostati: ' + nomiSp.join(', ') + ']';
        }
      }
      if (!nuovoTavolo) {
        bloccoPiccoloM = trovaTavoliMultipli_(stato.tavoli, occSenza, nuovePersone, necessitaAccessibile);
      }
      if (!nuovoTavolo && !bloccoPiccoloM) {
        var riorgMultiM = tentaRiorganizzazioneMulti_(stato.tavoli, stato.prenotazioni, occSenza, nuovePersone, necessitaAccessibile);
        if (riorgMultiM) {
          applicaSpostamentiSoloDati_(stato.prenotazioni, riorgMultiM.spostamenti);
          bloccoPiccoloM = riorgMultiM.blocco;
          var nomiM = riorgMultiM.spostamenti.map(function(s) { return s.nome + ' (da T.' + etichettaTavoliVisiva_(s.daTavolo) + ' a T.' + etichettaTavoliVisiva_(s.aTavolo) + ')'; });
          messaggioCambio += ' [Spostati: ' + nomiM.join(', ') + ']';
        }
      }
      if (!nuovoTavolo && !bloccoPiccoloM) {
        var avanzataM = tentaRiorganizzazioneAvanzata_(stato.tavoli, stato.prenotazioni, occSenza, nuovePersone, necessitaAccessibile);
        if (avanzataM) {
          applicaSpostamentiSoloDati_(stato.prenotazioni, avanzataM.spostamenti);
          bloccoPiccoloM = avanzataM.blocco;
          var nomiA = avanzataM.spostamenti.map(function(s) { return s.nome + ' (da T.' + etichettaTavoliVisiva_(s.daTavolo) + ' a T.' + etichettaTavoliVisiva_(s.aTavolo) + ')'; });
          messaggioCambio += ' [Spostati: ' + nomiA.join(', ') + ']';
        }
      }
      if (nuovoTavolo) {
        stato.prenotazioni[rigaPren][5] = nuovoTavolo.tavolo;
        stato.prenotazioni[rigaPren][6] = nuovoTavolo.zona;
        messaggioCambio = ' Tavolo cambiato: T.' + etichettaTavoliVisiva_(vecchioTavolo) + ' -> T.' + etichettaTavoliVisiva_(nuovoTavolo.tavolo) + '.' + messaggioCambio;
      } else if (bloccoPiccoloM) {
        stato.prenotazioni[rigaPren][5] = tavoliToString_(bloccoPiccoloM.tavoli);
        stato.prenotazioni[rigaPren][6] = bloccoPiccoloM.zona;
        messaggioCambio = ' Tavoli cambiati: T.' + etichettaTavoliVisiva_(vecchioTavolo) + ' -> T.' + etichettaTavoliVisiva_(bloccoPiccoloM.tavoli) + '.' + messaggioCambio;
      } else {
        throw new Error('Nessun tavolo (o blocco adiacente) disponibile per ' + nuovePersone + ' persone.');
      }
    } else {
      var blocco = trovaTavoliMultipli_(stato.tavoli, occSenza, nuovePersone, necessitaAccessibile);
      if (!blocco) {
        var riorgMulti = tentaRiorganizzazioneMulti_(stato.tavoli, stato.prenotazioni, occSenza, nuovePersone, necessitaAccessibile);
        if (riorgMulti) {
          applicaSpostamentiSoloDati_(stato.prenotazioni, riorgMulti.spostamenti);
          blocco = riorgMulti.blocco;
          var nomiSp2 = riorgMulti.spostamenti.map(function(s) { return s.nome + ' (da T.' + etichettaTavoliVisiva_(s.daTavolo) + ' a T.' + etichettaTavoliVisiva_(s.aTavolo) + ')'; });
          messaggioCambio += ' [Spostati: ' + nomiSp2.join(', ') + ']';
        }
      }
      if (!blocco) throw new Error('Nessun blocco di tavoli adiacenti per ' + nuovePersone + ' persone.');
      stato.prenotazioni[rigaPren][5] = tavoliToString_(blocco.tavoli);
      stato.prenotazioni[rigaPren][6] = blocco.zona;
      messaggioCambio = ' Tavoli cambiati: T.' + etichettaTavoliVisiva_(vecchioTavolo) + ' -> T.' + etichettaTavoliVisiva_(blocco.tavoli) + '.' + messaggioCambio;
    }
  }

  aggiornaRigheTavoliDaOccupazione_(stato.tavoli, calcolaOccupazioneMappa_(stato.prenotazioni));
  return 'Prenotazione #' + id + ' modificata: ' + nuovePersone + ' persone, disabili: ' + disabili + '.' + messaggioCambio;
}

function modificaPrenotazione(id, nuovePersone, disabili, note) {
  var dati = leggiTuttiIDati_();
  var msg = modificaPrenotazioneInMemoria_({ tavoli: dati.tavoli, prenotazioni: dati.prenotazioni }, id, nuovePersone, disabili, note);
  dati.sheetPren.getRange(2, 1, dati.prenotazioni.length, dati.prenotazioni[0].length).setValues(dati.prenotazioni);
  var nuovaOcc = calcolaOccupazioneMappa_(dati.prenotazioni);
  aggiornaStatoTavoli_(dati.sheetTavoli, dati.tavoli, nuovaOcc);
  aggiornaDashboard();
  aggiornaPlanimetria();
  return msg;
}

// Calcola occupazione escludendo una prenotazione specifica
function calcolaOccupazioneSenza_(prenotazioni, indiceDaEscludere) {
  return calcolaOccupazioneMappaCore_(prenotazioni, indiceDaEscludere);
}

// ============================================================
// FORM CANCELLAZIONE
// ============================================================
function mostraFormCancellazione() {
  var html = HtmlService.createHtmlOutput(getHtmlCancellazione_())
    .setWidth(420).setHeight(350).setTitle('Cancella Prenotazione');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getHtmlCancellazione_() {
  return '<style>'
    + 'body { font-family: Arial, sans-serif; padding: 15px; margin: 0; }'
    + 'label { display: block; margin-top: 12px; font-weight: bold; font-size: 14px; }'
    + 'input { width: 100%; padding: 10px; margin-top: 4px; font-size: 14px; border: 1px solid #ccc; border-radius: 6px; box-sizing: border-box; }'
    + '.btn { margin-top: 20px; padding: 14px 24px; font-size: 16px; color: white; border: none; border-radius: 8px; cursor: pointer; width: 100%; }'
    + '.btn-canc { background-color: #dc3545; }'
    + '.btn-canc:hover:not(:disabled) { background-color: #c82333; }'
    + '.btn-canc:disabled { background-color: #999; cursor: wait; }'
    + '#risultato { margin-top: 15px; padding: 12px; border-radius: 6px; font-size: 14px; display: none; line-height: 1.5; }'
    + '.successo { background-color: #d4edda; color: #155724; border: 1px solid #c3e6cb; }'
    + '.errore { background-color: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }'
    + '.spinner { display: inline-block; width: 16px; height: 16px; border: 3px solid #fff; border-top-color: transparent;'
    + '  border-radius: 50%; animation: spin 0.8s linear infinite; vertical-align: middle; margin-right: 8px; }'
    + '@keyframes spin { to { transform: rotate(360deg); } }'
    + '</style>'
    + '<label>ID Prenotazione da cancellare:</label>'
    + '<input type="number" id="idPren" min="1" placeholder="Es: 3">'
    + '<button class="btn btn-canc" id="btnCancella" onclick="cancella()">CANCELLA PRENOTAZIONE</button>'
    + '<div id="risultato"></div>'
    + '<script>'
    + 'function cancella() {'
    + '  var id = parseInt(document.getElementById("idPren").value);'
    + '  if (!id) { mostraRisultato("Inserisci un ID valido.", "errore"); return; }'
    + '  setLoading(true);'
    + '  google.script.run'
    + '    .withSuccessHandler(function(msg) { mostraRisultato(msg, "successo"); setLoading(false); document.getElementById("idPren").value = ""; })'
    + '    .withFailureHandler(function(err) { mostraRisultato(err.message, "errore"); setLoading(false); })'
    + '    .cancellaPrenotazione(id);'
    + '}'
    + 'function setLoading(on) { var b = document.getElementById("btnCancella"); b.disabled = on;'
    + '  b.innerHTML = on ? \'<span class="spinner"></span>Cancellazione...\' : "CANCELLA PRENOTAZIONE"; }'
    + 'function mostraRisultato(msg, tipo) { var d = document.getElementById("risultato");'
    + '  d.textContent = msg; d.className = tipo; d.style.display = "block"; }'
    + '</script>';
}

function cancellaPrenotazione(id) {
  var dati = leggiTuttiIDati_();
  var trovata = false;
  var nomePren = '';
  for (var i = 0; i < dati.prenotazioni.length; i++) {
    if (dati.prenotazioni[i][0] === id && dati.prenotazioni[i][8] !== 'Cancellata') {
      dati.prenotazioni[i][8] = 'Cancellata';
      nomePren = dati.prenotazioni[i][1];
      trovata = true;
      break;
    }
  }
  if (!trovata) throw new Error('Prenotazione #' + id + ' non trovata o gia cancellata.');

  dati.sheetPren.getRange(2, 1, dati.prenotazioni.length, dati.prenotazioni[0].length).setValues(dati.prenotazioni);
  var occupazione = calcolaOccupazioneMappa_(dati.prenotazioni);
  aggiornaStatoTavoli_(dati.sheetTavoli, dati.tavoli, occupazione);
  aggiornaDashboard();
  aggiornaPlanimetria();

  return 'Prenotazione #' + id + ' (' + nomePren + ') cancellata.';
}

// ============================================================
// FORM SPOSTAMENTO
// ============================================================
function mostraFormSpostamento() {
  var html = HtmlService.createHtmlOutput(getHtmlSpostamento_())
    .setWidth(420).setHeight(400).setTitle('Sposta Prenotazione');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getHtmlSpostamento_() {
  return '<style>'
    + 'body { font-family: Arial, sans-serif; padding: 15px; margin: 0; }'
    + 'label { display: block; margin-top: 12px; font-weight: bold; font-size: 14px; }'
    + 'input { width: 100%; padding: 10px; margin-top: 4px; font-size: 14px; border: 1px solid #ccc; border-radius: 6px; box-sizing: border-box; }'
    + '.btn { margin-top: 20px; padding: 14px 24px; font-size: 16px; color: white; border: none; border-radius: 8px; cursor: pointer; width: 100%; }'
    + '.btn-sposta { background-color: #007bff; }'
    + '.btn-sposta:hover:not(:disabled) { background-color: #0069d9; }'
    + '.btn-sposta:disabled { background-color: #999; cursor: wait; }'
    + '#risultato { margin-top: 15px; padding: 12px; border-radius: 6px; font-size: 14px; display: none; line-height: 1.5; }'
    + '.successo { background-color: #d4edda; color: #155724; border: 1px solid #c3e6cb; }'
    + '.errore { background-color: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }'
    + '.spinner { display: inline-block; width: 16px; height: 16px; border: 3px solid #fff; border-top-color: transparent;'
    + '  border-radius: 50%; animation: spin 0.8s linear infinite; vertical-align: middle; margin-right: 8px; }'
    + '@keyframes spin { to { transform: rotate(360deg); } }'
    + '</style>'
    + '<label>ID Prenotazione da spostare:</label>'
    + '<input type="number" id="idPren" min="1" placeholder="Es: 2">'
    + '<label>Nuovo Tavolo (es: 3 oppure 2;3 per piu tavoli — usa ; non la virgola):</label>'
    + '<input type="text" id="nuovoTavolo" placeholder="Es: 3 oppure 2;3">'
    + '<button class="btn btn-sposta" id="btnSposta" onclick="sposta()">SPOSTA PRENOTAZIONE</button>'
    + '<div id="risultato"></div>'
    + '<script>'
    + 'function sposta() {'
    + '  var id = parseInt(document.getElementById("idPren").value);'
    + '  var tavolo = document.getElementById("nuovoTavolo").value.trim();'
    + '  if (!id || !tavolo) { mostraRisultato("Compila entrambi i campi.", "errore"); return; }'
    + '  setLoading(true);'
    + '  google.script.run'
    + '    .withSuccessHandler(function(msg) { mostraRisultato(msg, "successo"); setLoading(false); document.getElementById("idPren").value=""; document.getElementById("nuovoTavolo").value=""; })'
    + '    .withFailureHandler(function(err) { mostraRisultato(err.message, "errore"); setLoading(false); })'
    + '    .spostaPrenotazione(id, tavolo);'
    + '}'
    + 'function setLoading(on) { var b = document.getElementById("btnSposta"); b.disabled = on;'
    + '  b.innerHTML = on ? \'<span class="spinner"></span>Spostamento...\' : "SPOSTA PRENOTAZIONE"; }'
    + 'function mostraRisultato(msg, tipo) { var d = document.getElementById("risultato");'
    + '  d.textContent = msg; d.className = tipo; d.style.display = "block"; }'
    + '</script>';
}

function spostaPrenotazione(id, nuovoTavoloStr) {
  var dati = leggiTuttiIDati_();
  var occupazione = calcolaOccupazioneMappa_(dati.prenotazioni);

  var prenotazione = null;
  var rigaPren = -1;
  for (var i = 0; i < dati.prenotazioni.length; i++) {
    if (dati.prenotazioni[i][0] === id && dati.prenotazioni[i][8] === 'Confermata') {
      prenotazione = dati.prenotazioni[i];
      rigaPren = i;
      break;
    }
  }
  if (!prenotazione) throw new Error('Prenotazione #' + id + ' non trovata o non attiva.');

  var numPersone = prenotazione[3];
  var vecchioTavolo = prenotazione[5];
  var nuoviTavoli = parseTavoli_(nuovoTavoloStr);

  if (nuoviTavoli.length === 0) throw new Error('Tavolo non valido: ' + nuovoTavoloStr);

  // Ricalcola occupazione senza questa prenotazione
  var occSenza = calcolaOccupazioneSenza_(dati.prenotazioni, rigaPren);

  // Verifica che i tavoli esistano e abbiano posto
  var postiLiberiTotali = 0;
  var zona = '';
  for (var t = 0; t < nuoviTavoli.length; t++) {
    var trovato = false;
    for (var j = 0; j < dati.tavoli.length; j++) {
      if (dati.tavoli[j][1] === nuoviTavoli[t]) {
        trovato = true;
        zona = dati.tavoli[j][0];
        postiLiberiTotali += dati.tavoli[j][2] - (occSenza[nuoviTavoli[t]] || 0);
        break;
      }
    }
    if (!trovato) throw new Error('Tavolo ' + nuoviTavoli[t] + ' non esiste.');
  }

  if (postiLiberiTotali < numPersone) {
    throw new Error('I tavoli ' + nuovoTavoloStr + ' hanno solo ' + postiLiberiTotali + ' posti liberi, ne servono ' + numPersone + '.');
  }

  var valTavolo;
  if (nuoviTavoli.length === 1) {
    valTavolo = nuoviTavoli[0];
  } else {
    var ordSposta = null;
    if (puoCollocareMultitavolo_(numPersone, nuoviTavoli, occSenza)) ordSposta = nuoviTavoli.slice();
    else if (puoCollocareMultitavolo_(numPersone, nuoviTavoli.slice().reverse(), occSenza)) ordSposta = nuoviTavoli.slice().reverse();
    if (!ordSposta) {
      throw new Error('Il gruppo non puo sedersi in sequenza su questi tavoli: un tavolo intermedio e gia pieno o mancano posti nel percorso. Scegli altri tavoli o libera un tavolo nel mezzo.');
    }
    valTavolo = tavoliToString_(ordSposta);
    for (var z = 0; z < dati.tavoli.length; z++) {
      if (dati.tavoli[z][1] === ordSposta[0]) { zona = dati.tavoli[z][0]; break; }
    }
  }

  dati.prenotazioni[rigaPren][5] = valTavolo;
  dati.prenotazioni[rigaPren][6] = zona;

  dati.sheetPren.getRange(2, 1, dati.prenotazioni.length, dati.prenotazioni[0].length).setValues(dati.prenotazioni);

  var nuovaOcc = calcolaOccupazioneMappa_(dati.prenotazioni);
  aggiornaStatoTavoli_(dati.sheetTavoli, dati.tavoli, nuovaOcc);
  aggiornaDashboard();
  aggiornaPlanimetria();

  return 'Prenotazione #' + id + ' spostata: T.' + etichettaTavoliVisiva_(vecchioTavolo) + ' -> T.' + etichettaTavoliVisiva_(valTavolo) + '.';
}

// ============================================================
// OTTIMIZZA TAVOLI
// ============================================================
function ottimizzaTavoli() {
  var ui = SpreadsheetApp.getUi();
  var dati = leggiTuttiIDati_();
  var occupazione = calcolaOccupazioneMappa_(dati.prenotazioni);

  for (var i = 0; i < dati.tavoli.length; i++) {
    var numTavolo = dati.tavoli[i][1];
    var postiTotali = dati.tavoli[i][2];
    var occ = occupazione[numTavolo] || 0;
    dati.tavoli[i][3] = occ;
    dati.tavoli[i][4] = postiTotali - occ;
    dati.tavoli[i][6] = occ >= postiTotali ? 'Pieno' : (occ > 0 ? 'Parziale' : 'Libero');
  }

  var prenotazioniPerTavolo = {};
  for (var i = 0; i < dati.prenotazioni.length; i++) {
    var p = dati.prenotazioni[i];
    if (p[8] === 'Cancellata') continue;
    var listaTav = parseTavoli_(p[5]);
    // Solo prenotazioni singolo-tavolo possono essere ottimizzate
    if (listaTav.length !== 1) continue;
    var tav = listaTav[0];
    if (!prenotazioniPerTavolo[tav]) prenotazioniPerTavolo[tav] = [];
    prenotazioniPerTavolo[tav].push({ rigaIndice: i, persone: p[3], nome: p[1] });
  }

  var parziali = [];
  for (var i = 0; i < dati.tavoli.length; i++) {
    var t = dati.tavoli[i];
    if (t[6] === 'Parziale') {
      parziali.push({
        numero: t[1], zona: t[0], occupati: t[3], liberi: t[4],
        accessibile: t[5] === 'Si', prenotazioni: prenotazioniPerTavolo[t[1]] || []
      });
    }
  }

  if (parziali.length < 2) {
    ui.alert('Ottimizzazione', 'Non ci sono abbastanza tavoli parziali da ottimizzare.', ui.ButtonSet.OK);
    return;
  }

  parziali.sort(function(a, b) { return a.occupati - b.occupati; });

  var spostamenti = [];
  var occupazioneSim = {};
  for (var k in occupazione) { occupazioneSim[k] = occupazione[k]; }

  for (var i = 0; i < parziali.length; i++) {
    var sorgente = parziali[i];
    if ((occupazioneSim[sorgente.numero] || 0) === 0) continue;

    for (var j = parziali.length - 1; j > i; j--) {
      var dest = parziali[j];
      var liberiDest = POSTI_PER_TAVOLO - (occupazioneSim[dest.numero] || 0);
      var occSorgente = occupazioneSim[sorgente.numero] || 0;

      if (liberiDest >= occSorgente) {
        var prenDaSpostare = (prenotazioniPerTavolo[sorgente.numero] || []).filter(function(p) {
          return dati.prenotazioni[p.rigaIndice][5] == sorgente.numero && dati.prenotazioni[p.rigaIndice][8] !== 'Cancellata';
        });
        for (var pi = 0; pi < prenDaSpostare.length; pi++) {
          spostamenti.push({
            nome: prenDaSpostare[pi].nome, persone: prenDaSpostare[pi].persone,
            daTavolo: sorgente.numero, aTavolo: dest.numero, aZona: dest.zona,
            rigaIndice: prenDaSpostare[pi].rigaIndice
          });
        }
        occupazioneSim[dest.numero] = (occupazioneSim[dest.numero] || 0) + occSorgente;
        occupazioneSim[sorgente.numero] = 0;
        break;
      }
    }
  }

  if (spostamenti.length === 0) {
    ui.alert('Ottimizzazione', 'Nessuno spostamento possibile al momento.', ui.ButtonSet.OK);
    return;
  }

  var msg = 'Spostamenti proposti:\n\n';
  for (var i = 0; i < spostamenti.length; i++) {
    msg += '\u2022 ' + spostamenti[i].nome + ' (' + spostamenti[i].persone + ' pers.) - T.' + etichettaTavoliVisiva_(spostamenti[i].daTavolo) + ' -> T.' + etichettaTavoliVisiva_(spostamenti[i].aTavolo) + '\n';
  }
  msg += '\nApplicare?';

  if (ui.alert('Ottimizzazione Tavoli', msg, ui.ButtonSet.YES_NO) === ui.Button.YES) {
    for (var i = 0; i < spostamenti.length; i++) {
      dati.prenotazioni[spostamenti[i].rigaIndice][5] = spostamenti[i].aTavolo;
      dati.prenotazioni[spostamenti[i].rigaIndice][6] = spostamenti[i].aZona;
    }
    dati.sheetPren.getRange(2, 1, dati.prenotazioni.length, dati.prenotazioni[0].length).setValues(dati.prenotazioni);
    var nuovaOcc = calcolaOccupazioneMappa_(dati.prenotazioni);
    aggiornaStatoTavoli_(dati.sheetTavoli, dati.tavoli, nuovaOcc);
    aggiornaDashboard();
    aggiornaPlanimetria();
    ui.alert('Fatto!', spostamenti.length + ' spostamenti applicati.', ui.ButtonSet.OK);
  }
}

// ============================================================
// AGGIORNA DASHBOARD
// ============================================================
function aggiornaDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDash = ss.getSheetByName(SHEET_DASHBOARD);
  if (!sheetDash) sheetDash = ss.insertSheet(SHEET_DASHBOARD, 0);

  var sheetTavoli = ss.getSheetByName(SHEET_TAVOLI);
  var numTavoli = sheetTavoli.getLastRow() - 1;
  if (numTavoli <= 0) return;

  var datiTavoli = sheetTavoli.getRange(2, 1, numTavoli, 7).getValues();

  var postiTotali = 0, postiOccupati = 0, postiLiberi = 0;
  var tavoliLiberi = 0, tavoliParziali = 0, tavoliPieni = 0;
  var distribuzionePostiLiberi = {};

  for (var i = 0; i < datiTavoli.length; i++) {
    var t = datiTavoli[i];
    postiTotali += t[2]; postiOccupati += t[3]; postiLiberi += t[4];
    if (t[6] === 'Libero') tavoliLiberi++;
    else if (t[6] === 'Parziale') tavoliParziali++;
    else if (t[6] === 'Pieno') tavoliPieni++;
    if (t[4] > 0) distribuzionePostiLiberi[t[4]] = (distribuzionePostiLiberi[t[4]] || 0) + 1;
  }

  var contenuto = [];
  var formattazione = [];

  contenuto.push(['GESTIONE PRENOTAZIONI - FESTA (62 tavoli)', '', '', '']);
  formattazione.push({ riga: 0, col: 0, size: 18, bold: true, bg: null, color: null });
  contenuto.push(['Aggiornato: ' + Utilities.formatDate(new Date(), 'Europe/Rome', 'dd/MM/yyyy HH:mm'), '', '', '']);
  formattazione.push({ riga: 1, col: 0, size: 10, bold: false, bg: null, color: '#666666' });
  contenuto.push(['', '', '', '']);
  contenuto.push(['RIEPILOGO', '', '', '']);
  formattazione.push({ riga: 3, col: 0, size: 14, bold: true, bg: null, color: null });
  contenuto.push(['Posti Totali', postiTotali, '', '']);
  contenuto.push(['Posti Occupati', postiOccupati, '', '']);
  contenuto.push(['POSTI DISPONIBILI', postiLiberi, '', '']);
  formattazione.push({ riga: 6, col: 1, size: 18, bold: true, bg: '#b7e1cd', color: null });
  contenuto.push(['', '', '', '']);
  contenuto.push(['Tavoli Liberi', tavoliLiberi, '', '']);
  formattazione.push({ riga: 8, col: 1, size: null, bold: true, bg: '#b7e1cd', color: null });
  contenuto.push(['Tavoli Parziali', tavoliParziali, '', '']);
  formattazione.push({ riga: 9, col: 1, size: null, bold: true, bg: '#fce8b2', color: null });
  contenuto.push(['Tavoli Pieni', tavoliPieni, '', '']);
  formattazione.push({ riga: 10, col: 1, size: null, bold: true, bg: '#f4c7c3', color: null });
  contenuto.push(['', '', '', '']);
  contenuto.push(['PRENOTAZIONI ACCETTABILI', '', '', '']);
  formattazione.push({ riga: 12, col: 0, size: 14, bold: true, bg: null, color: null });

  var chiavi = Object.keys(distribuzionePostiLiberi).map(Number).sort(function(a, b) { return b - a; });
  if (chiavi.length === 0) {
    contenuto.push(['TUTTO ESAURITO!', '', '', '']);
    formattazione.push({ riga: 13, col: 0, size: 14, bold: true, bg: '#f4c7c3', color: '#dc3545' });
  } else {
    contenuto.push(['Gruppo da...', 'Tavoli disponibili', '', '']);
    formattazione.push({ riga: 13, col: 0, size: 12, bold: true, bg: '#e8eaf6', color: null });
    formattazione.push({ riga: 13, col: 1, size: 12, bold: true, bg: '#e8eaf6', color: null });
    for (var ki = 0; ki < chiavi.length; ki++) {
      var conteggio = 0;
      for (var kj = 0; kj < chiavi.length; kj++) { if (chiavi[kj] >= chiavi[ki]) conteggio += distribuzionePostiLiberi[chiavi[kj]]; }
      contenuto.push([chiavi[ki] + ' persone', conteggio + ' tavoli', '', '']);
    }
  }

  sheetDash.clear();
  sheetDash.getRange(1, 1, contenuto.length, 4).setValues(contenuto).setFontSize(13);
  for (var fi = 0; fi < formattazione.length; fi++) {
    var f = formattazione[fi];
    var cella = sheetDash.getRange(f.riga + 1, f.col + 1);
    if (f.size) cella.setFontSize(f.size);
    if (f.bold) cella.setFontWeight('bold');
    if (f.bg) cella.setBackground(f.bg);
    if (f.color) cella.setFontColor(f.color);
  }
  sheetDash.setColumnWidth(1, 280);
  sheetDash.setColumnWidth(2, 180);
  sheetDash.setColumnWidth(3, 150);
  sheetDash.setColumnWidth(4, 150);
}

// ============================================================
// PLANIMETRIA VISIVA (layout: Ballo sx | corridoio | Ballo dx; Chiosco; Esterna)
// ============================================================
function aggiornaPlanimetria() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_PLANIMETRIA);
  if (!sheet) sheet = ss.insertSheet(SHEET_PLANIMETRIA);

  var sheetTavoli = ss.getSheetByName(SHEET_TAVOLI);
  var sheetPren = ss.getSheetByName(SHEET_PRENOTAZIONI);

  var numTavoli = sheetTavoli.getLastRow() - 1;
  if (numTavoli <= 0) return;

  var datiTavoli = sheetTavoli.getRange(2, 1, numTavoli, 7).getValues();
  var numPren = sheetPren.getLastRow() - 1;
  var prenotazioni = numPren > 0 ? sheetPren.getRange(2, 1, numPren, 10).getValues() : [];

  var nomiPerTavolo = {};
  for (var i = 0; i < prenotazioni.length; i++) {
    var p = prenotazioni[i];
    if (p[8] === 'Cancellata') continue;
    var lista = parseTavoli_(p[5]);
    for (var t = 0; t < lista.length; t++) {
      if (!nomiPerTavolo[lista[t]]) nomiPerTavolo[lista[t]] = [];
      nomiPerTavolo[lista[t]].push(p[1] + ' (' + p[3] + ')');
    }
  }

  var statoTavoli = {};
  for (var ir = 0; ir < datiTavoli.length; ir++) {
    var rowT = datiTavoli[ir];
    statoTavoli[rowT[1]] = { occupati: rowT[3], liberi: rowT[4], totali: rowT[2], stato: rowT[6], accessibile: rowT[5] === 'Si' };
  }

  sheet.clear();

  // Nessun loop setColumnWidth: in alcuni fogli/ambienti puo generare "Queste colonne sono fuori dai limiti".
  // Larghezze opzionali solo prime colonne (100-1000 px).
  try {
    for (var cw = 1; cw <= 22; cw++) sheet.setColumnWidth(cw, 110);
  } catch (eW) {}

  sheet.getRange('A1').setValue('PLANIMETRIA TAVOLI - FESTA').setFontSize(14).setFontWeight('bold');
  sheet.getRange('A2:D2').setValues([['Legenda:', 'Libero', 'Parziale', 'Pieno']]).setFontWeight('bold');
  sheet.getRange('B2').setBackground('#b7e1cd');
  sheet.getRange('C2').setBackground('#fce8b2');
  sheet.getRange('D2').setBackground('#f4c7c3');

  sheet.getRange('A3').setValue('Sala Ballo (blocco sx: 6x3 tavoli | corridoio col. G | blocco dx: 3+3+2+2+2+3)').setFontWeight('bold');
  sheet.getRange('A22').setValue('Sala Chiosco (tre file verticali 3 | · | 3 | · | 4; riga vuota; fila 5 orizzontale sotto)').setFontWeight('bold');
  sheet.getRange('A40').setValue('Sala Esterna (7 file x 2 tavoli, prime colonne)').setFontWeight('bold');

  var maxRow = 65;
  for (var rr = 1; rr <= maxRow; rr++) {
    try {
      sheet.setRowHeight(rr, 22);
    } catch (eR) {}
  }

  for (var mi = 0; mi < MAPPA_TAVOLI.length; mi++) {
    var m = MAPPA_TAVOLI[mi];
    var st = statoTavoli[m.tavolo];
    if (!st) continue;
    var rng = sheet.getRange(rangeA1FromDims_(m.row, m.col, 3, 2));
    rng.merge();
    var label = 'T.' + m.tavolo;
    if (st.accessibile) label += ' \u267F';
    var nomi = nomiPerTavolo[m.tavolo] || [];
    var testo = label + '\n' + st.occupati + '/' + st.totali + ' occ. · ' + st.liberi + ' liberi\n' + (nomi.length ? nomi.join(', ') : '(vuoto)');
    rng.setValue(testo).setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('center').setFontSize(9);
    var bg = '#b7e1cd';
    if (st.stato === 'Parziale') bg = '#fce8b2';
    if (st.stato === 'Pieno') bg = '#f4c7c3';
    rng.setBackground(bg).setBorder(true, true, true, true, false, false, '#333333', SpreadsheetApp.BorderStyle.SOLID);
  }

  sheet.getRange('A' + maxRow).setValue('\u267F = tavolo accessibile (Chiosco: fila verticale sinistra 39-41)').setFontSize(9).setFontColor('#666');

  SpreadsheetApp.flush();
}

// ============================================================
// VERIFICA DISPONIBILITA
// ============================================================
function verificaDisponibilita(persone, disabili) {
  var dati = leggiTuttiIDati_();
  var occupazione = calcolaOccupazioneMappa_(dati.prenotazioni);
  var necessitaAccessibile = (disabili === 'Si');

  if (persone <= POSTI_PER_TAVOLO) {
    var tavolo = trovaMigliorTavolo_(dati.tavoli, occupazione, persone, necessitaAccessibile);
    if (tavolo) {
      return { esito: 'diretto', messaggio: 'Posto disponibile! Tavolo ' + tavolo.tavolo + ' (' + tavolo.zona + '), ' + tavolo.postiLiberi + ' posti liberi.' };
    }

    var riorg = tentaRiorganizzazione_(dati.tavoli, dati.prenotazioni, occupazione, persone, necessitaAccessibile);
    if (riorg) {
      var dettagli = riorg.spostamenti.map(function(s) { return s.nome + ' da T.' + etichettaTavoliVisiva_(s.daTavolo) + ' a T.' + etichettaTavoliVisiva_(s.aTavolo); }).join(', ');
      return { esito: 'ottimizzazione', messaggio: 'Posto disponibile ottimizzando! Servono ' + riorg.spostamenti.length + ' spostamenti (' + dettagli + ') per liberare il Tavolo ' + riorg.tavolo.tavolo + '.' };
    }

    var bloccoP = trovaTavoliMultipli_(dati.tavoli, occupazione, persone, necessitaAccessibile);
    if (bloccoP) {
      return { esito: 'diretto', messaggio: 'Posto disponibile su tavoli adiacenti! Tavoli ' + etichettaTavoliVisiva_(bloccoP.tavoli) + ' (' + bloccoP.zona + '), ' + bloccoP.postiLiberi + ' posti liberi nel blocco.' };
    }

    var riorgMulti = tentaRiorganizzazioneMulti_(dati.tavoli, dati.prenotazioni, occupazione, persone, necessitaAccessibile);
    if (riorgMulti) {
      var detM = riorgMulti.spostamenti.map(function(s) { return s.nome + ' da T.' + etichettaTavoliVisiva_(s.daTavolo) + ' a T.' + etichettaTavoliVisiva_(s.aTavolo); }).join(', ');
      return { esito: 'ottimizzazione', messaggio: 'Posto disponibile ottimizzando! Servono ' + riorgMulti.spostamenti.length + ' spostamenti (' + detM + ') per liberare i Tavoli ' + etichettaTavoliVisiva_(riorgMulti.blocco.tavoli) + '.' };
    }

    var avanzata = tentaRiorganizzazioneAvanzata_(dati.tavoli, dati.prenotazioni, occupazione, persone, necessitaAccessibile);
    if (avanzata) {
      var detA = avanzata.spostamenti.map(function(s) { return s.nome + ' da T.' + etichettaTavoliVisiva_(s.daTavolo) + ' a T.' + etichettaTavoliVisiva_(s.aTavolo); }).join(', ');
      return { esito: 'ottimizzazione', messaggio: 'Posto disponibile ottimizzando (riordino tavoli / spostamento piccolo gruppo)! ' + detA + ' → Tavoli ' + etichettaTavoliVisiva_(avanzata.blocco.tavoli) + '.' };
    }
  } else {
    var blocco = trovaTavoliMultipli_(dati.tavoli, occupazione, persone, necessitaAccessibile);
    if (blocco) {
      return { esito: 'diretto', messaggio: 'Posto disponibile! Tavoli ' + etichettaTavoliVisiva_(blocco.tavoli) + ' (' + blocco.zona + '), ' + blocco.postiLiberi + ' posti liberi totali.' };
    }

    var riorgMulti = tentaRiorganizzazioneMulti_(dati.tavoli, dati.prenotazioni, occupazione, persone, necessitaAccessibile);
    if (riorgMulti) {
      var dettagli = riorgMulti.spostamenti.map(function(s) { return s.nome + ' da T.' + etichettaTavoliVisiva_(s.daTavolo) + ' a T.' + etichettaTavoliVisiva_(s.aTavolo); }).join(', ');
      return { esito: 'ottimizzazione', messaggio: 'Posto disponibile ottimizzando! Servono ' + riorgMulti.spostamenti.length + ' spostamenti (' + dettagli + ') per liberare i Tavoli ' + etichettaTavoliVisiva_(riorgMulti.blocco.tavoli) + '.' };
    }
  }

  var postiLib = dati.tavoli.reduce(function(s, t) { return s + t[2] - (occupazione[t[1]] || 0); }, 0);
  return { esito: 'no', messaggio: 'Impossibile accettare ' + persone + ' persone' + (necessitaAccessibile ? ' (accessibile)' : '') + '. Posti liberi totali: ' + postiLib + '.' };
}

// ============================================================
// FORM TROVA PRENOTAZIONE
// ============================================================
function mostraFormRicerca() {
  var html = HtmlService.createHtmlOutput(getHtmlRicerca_())
    .setWidth(420).setHeight(500).setTitle('Trova Prenotazione');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getHtmlRicerca_() {
  return '<style>'
    + 'body { font-family: Arial, sans-serif; padding: 15px; margin: 0; }'
    + 'label { display: block; margin-top: 12px; font-weight: bold; font-size: 14px; }'
    + 'input { width: 100%; padding: 10px; margin-top: 4px; font-size: 14px; border: 1px solid #ccc; border-radius: 6px; box-sizing: border-box; }'
    + '.btn { margin-top: 20px; padding: 14px 24px; font-size: 16px; color: white; border: none; border-radius: 8px; cursor: pointer; width: 100%; }'
    + '.btn-cerca { background-color: #6f42c1; }'
    + '.btn-cerca:hover:not(:disabled) { background-color: #5a32a3; }'
    + '.btn-cerca:disabled { background-color: #999; cursor: wait; }'
    + '#risultati { margin-top: 15px; }'
    + '.trovato { background-color: #d4edda; color: #155724; border: 1px solid #c3e6cb; padding: 14px; border-radius: 8px; margin-bottom: 10px; font-size: 15px; line-height: 1.6; }'
    + '.trovato b { font-size: 16px; }'
    + '.non-trovato { background-color: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; padding: 14px; border-radius: 8px; font-size: 14px; }'
    + '.spinner { display: inline-block; width: 16px; height: 16px; border: 3px solid #fff; border-top-color: transparent;'
    + '  border-radius: 50%; animation: spin 0.8s linear infinite; vertical-align: middle; margin-right: 8px; }'
    + '@keyframes spin { to { transform: rotate(360deg); } }'
    + '</style>'
    + '<label>Nome e Cognome (anche parziale):</label>'
    + '<input type="text" id="nome" placeholder="Es: Mario">'
    + '<label>Numero Persone (facoltativo):</label>'
    + '<input type="number" id="persone" min="1" placeholder="Lascia vuoto per cercare tutti">'
    + '<button class="btn btn-cerca" id="btnCerca" onclick="cerca()">CERCA PRENOTAZIONE</button>'
    + '<div id="risultati"></div>'
    + '<script>'
    + 'function cerca() {'
    + '  var nome = document.getElementById("nome").value.trim();'
    + '  var persone = document.getElementById("persone").value;'
    + '  persone = persone ? parseInt(persone) : 0;'
    + '  if (!nome) { document.getElementById("risultati").innerHTML = \'<div class="non-trovato">Inserisci un nome da cercare.</div>\'; return; }'
    + '  var b = document.getElementById("btnCerca"); b.disabled = true;'
    + '  b.innerHTML = \'<span class="spinner"></span>Ricerca...\';'
    + '  google.script.run'
    + '    .withSuccessHandler(function(risultati) {'
    + '      b.disabled = false; b.innerHTML = "CERCA PRENOTAZIONE";'
    + '      var div = document.getElementById("risultati");'
    + '      if (!risultati || risultati.length === 0) { div.innerHTML = \'<div class="non-trovato">Nessuna prenotazione trovata.</div>\'; return; }'
    + '      var html = "";'
    + '      for (var i = 0; i < risultati.length; i++) { var r = risultati[i];'
    + '        html += \'<div class="trovato"><b>\' + r.nome + \'</b>, al tavolo <b>T.\' + r.tavolo + \'</b>\';'
    + '        html += \', prenotazione per <b>\' + r.persone + \' persone</b> con telefono <b>\' + r.telefono + \'</b>\';'
    + '        if (r.note) html += \'<br>Note: \' + r.note;'
    + '        html += \'<br><small>Prenotazione #\' + r.id + \' - \' + r.data + \'</small></div>\'; }'
    + '      div.innerHTML = html;'
    + '    }).withFailureHandler(function(err) { b.disabled = false; b.innerHTML = "CERCA PRENOTAZIONE";'
    + '      document.getElementById("risultati").innerHTML = \'<div class="non-trovato">\' + err.message + \'</div>\'; })'
    + '    .cercaPrenotazione(nome, persone);'
    + '}'
    + '</script>';
}

function cercaPrenotazione(nome, persone) {
  var dati = leggiTuttiIDati_();
  var risultati = [];
  var nomeLower = nome.toLowerCase();

  for (var i = 0; i < dati.prenotazioni.length; i++) {
    var p = dati.prenotazioni[i];
    if (p[8] === 'Cancellata') continue;
    if (String(p[1]).toLowerCase().indexOf(nomeLower) === -1) continue;
    if (persone > 0 && p[3] !== persone) continue;

    var dataStr = '';
    if (p[7] instanceof Date) { dataStr = Utilities.formatDate(p[7], 'Europe/Rome', 'dd/MM/yyyy HH:mm'); }
    else { dataStr = String(p[7]); }

    risultati.push({ id: p[0], nome: p[1], telefono: p[2], persone: p[3], tavolo: p[5], zona: p[6], data: dataStr, note: p[9] || '' });
  }
  return risultati;
}
