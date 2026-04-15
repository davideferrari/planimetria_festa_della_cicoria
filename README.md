# Gestione Prenotazioni Fiera

Sistema di gestione prenotazioni tavoli per fiere ed eventi basato su Google Sheets + Apps Script.

## Caratteristiche

- **73 tavoli** divisi in 4 zone (Sala Principale, Blocco Alto SX, Blocco Alto DX, Blocco Centrale)
- **8 posti per tavolo**, circa 584 posti totali
- Gestione prenotazioni con form laterali semplici
- Supporto disabili (tavoli accessibili vicino alle uscite grandi)
- Dashboard in tempo reale con riepilogo posti e prenotazioni accettabili
- Ottimizzazione passiva (best-fit) e attiva (auto-riorganizzazione)

## Installazione

### 1. Creare il Google Sheet

1. Vai su [Google Sheets](https://sheets.google.com) e crea un nuovo foglio
2. Rinominalo (es. "Prenotazioni Fiera 2026")

### 2. Inserire il codice Apps Script

1. Nel foglio, vai su **Estensioni** > **Apps Script**
2. Cancella tutto il contenuto nel file `Codice.gs`
3. Copia e incolla l'intero contenuto del file `Code.gs` di questo progetto
4. Salva con **Ctrl+S**
5. Chiudi l'editor Apps Script

### 3. Inizializzare il sistema

1. Ricarica il Google Sheet (F5)
2. Apparirà il menu **🎪 Gestione Fiera** nella barra dei menu
3. Clicca su **Gestione Fiera** > **Inizializza Sistema**
4. La prima volta Google chiederà di autorizzare lo script: accetta
5. Il sistema creerà i fogli TAVOLI, PRENOTAZIONI e DASHBOARD

## Utilizzo

| Azione | Menu |
|---|---|
| Nuova prenotazione | Gestione Fiera > Nuova Prenotazione |
| Cancellare una prenotazione | Gestione Fiera > Cancella Prenotazione |
| Spostare una prenotazione | Gestione Fiera > Sposta Prenotazione |
| Ottimizzare i tavoli | Gestione Fiera > Ottimizza Tavoli |
| Aggiornare il riepilogo | Gestione Fiera > Aggiorna Dashboard |

## Come funziona l'ottimizzazione

### Ottimizzazione passiva (automatica)
Quando si inserisce una prenotazione, il sistema sceglie il tavolo con il **minor numero di posti liberi** che può contenere il gruppo. Questo minimizza lo spreco di spazio.

### Auto-riorganizzazione (automatica)
Se nessun tavolo ha abbastanza posti liberi per un nuovo gruppo, il sistema **tenta automaticamente** di spostare prenotazioni più piccole da un tavolo a un altro per liberare spazio sufficiente. L'operatore viene informato degli spostamenti nel messaggio di conferma.

### Ottimizzazione manuale
Dal menu "Ottimizza Tavoli", il sistema analizza tutti i tavoli parzialmente occupati e propone spostamenti per consolidare le prenotazioni e liberare tavoli interi. L'operatore può accettare o rifiutare.

## Personalizzazione

### Tavoli accessibili (disabili)
Nel file `Code.gs`, modifica l'array `TAVOLI_ACCESSIBILI` con i numeri dei tavoli realmente vicini alle uscite grandi:

```javascript
const TAVOLI_ACCESSIBILI = [1, 2, 3, 8, 9, 10, ...];
```

### Posti per tavolo
Se alcuni tavoli hanno un numero di posti diverso da 8, dopo l'inizializzazione modifica direttamente la colonna "Posti Totali" nel foglio TAVOLI.
