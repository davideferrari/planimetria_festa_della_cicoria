# Gestione prenotazioni — Festa della cicoria

Sistema di gestione prenotazioni tavoli per la **Festa della cicoria**, basato su **Google Sheets** e **Google Apps Script**. La disposizione dei tavoli e delle sale è modellata sulla planimetria dell’evento (vedi immagine di riferimento sotto).

![Planimetria Festa della cicoria](Festa_della_cicoria.jpeg)

*Planimetria tecnica di riferimento per l’edizione (aree cucina, chiosco, zona ballo, servizi, date operative, ecc.).*

## Struttura del progetto

| File | Ruolo |
|------|--------|
| `Prenotazioni.gs` | Logica principale: menu, fogli, form, assegnazione tavoli, dashboard, planimetria |
| `Tests.gs` | Test automatici in memoria (esecuzione da menu o da editor) |
| `TestGuidato.gs` | Test guidato con sidebar e scenari sulla planimetria reale del foglio |
| `Festa_della_cicoria.jpeg` | Planimetria dell’area festa (riferimento visivo e logistico) |

## Caratteristiche

- **62 tavoli** in **3 zone**: **Sala Ballo** (1–33), **Sala Chiosco** (34–48), **Sala Esterna** (49–62)
- **8 posti per tavolo**, **496 posti** totali se tutti i tavoli sono pieni
- **Tavoli accessibili**: 39, 40, 41 (fila Chiosco a sinistra, vicino all’accesso disabili)
- Priorità di riempimento: prima Sala Ballo, poi Chiosco, poi Esterna (con best-fit dentro la zona)
- Adiacenza tra tavoli definita per gruppi multi-tavolo; le tre sale **non** sono collegate tra loro nella mappa
- Foglio **PLANIMETRIA** aggiornato con lo stato delle prenotazioni
- Foglio **ISTRUZIONI** generato in fase di inizializzazione
- **Test automatici** (Logger) e **test guidato** (sidebar) per verifiche e dimostrazioni

## Installazione

### 1. Creare il Google Sheet

1. Apri [Google Fogli](https://sheets.google.com) e crea un nuovo foglio di calcolo
2. Rinominalo come preferisci (es. «Prenotazioni Festa della cicoria 2026»)

### 2. Inserire il codice Apps Script

1. Dal foglio: **Estensioni** → **Apps Script**
2. Crea o rinomina i file `.gs` in modo da avere **tre file** con il contenuto di:
   - `Prenotazioni.gs`
   - `Tests.gs`
   - `TestGuidato.gs`  
   (In un unico progetto Apps Script possono convivere più file; le funzioni sono condivise nello stesso contesto.)
3. Salva (**Ctrl+S**)

### 3. Inizializzare il sistema

1. Ricarica il foglio (**F5**)
2. Compare il menu **Gestione Fiera**
3. **Gestione Fiera** → **Inizializza Sistema**
4. Alla prima esecuzione autorizza lo script quando richiesto
5. Vengono creati/aggiornati i fogli **TAVOLI**, **PRENOTAZIONI**, **DASHBOARD**, **PLANIMETRIA**, **ISTRUZIONI**

## Utilizzo (menu Gestione Fiera)

| Azione | Voce di menu |
|--------|----------------|
| Nuova prenotazione | Nuova Prenotazione |
| Modifica dati prenotazione | Modifica Prenotazione |
| Ricerca | Trova Prenotazione |
| Cancellazione | Cancella Prenotazione |
| Spostamento su altro tavolo | Sposta Prenotazione |
| Ottimizzazione manuale (consolidamento) | Consolida tavoli parziali (manuale) |
| Aggiornare riepilogo | Aggiorna Dashboard |
| Reset completo dati tavoli/prenotazioni (conferma) | Inizializza Sistema |
| Test in memoria | Esegui test automatici (Logger) |
| Scenari sulla planimetria | Test guidato planimetria (sidebar) |

## Come funziona l’ottimizzazione

### Assegnazione (best-fit)

In inserimento, il sistema sceglie il tavolo (o la combinazione di tavoli adiacenti) che **minimizza lo spreco di posti** rispettando zone, accessibilità e capacità.

### Auto-riorganizzazione

Se non c’è posto su un singolo tavolo, lo script può **spostare** prenotazioni più piccole tra tavoli adiacenti per liberare capacità. Gli spostamenti compaiono nel messaggio di feedback.

### Consolidamento manuale

**Consolida tavoli parziali (manuale)** analizza i tavoli parzialmente occupati e propone spostamenti per liberare interi tavoli; l’operatore può accettare o rifiutare.

## Personalizzazione

### Tavoli accessibili

In `Prenotazioni.gs`, array `TAVOLI_ACCESSIBILI` (attualmente `[39, 40, 41]`). Allineare i numeri alla planimetria reale.

### Posti diversi da 8

Dopo l’inizializzazione, modificare la colonna **Posti Totali** nel foglio **TAVOLI** per i tavoli che non hanno 8 posti.
