# Torneo Ass 'e Mazz 2026 — sito GitHub Pages

Sito statico responsive per il monitoraggio del torneo di Ass 'e Mazz in Piazza Vittoria a Saviano.

## Cosa mostra il sito

- panoramica automatica del torneo;
- classifica generale dopo ciascun turno;
- composizione e risultati di tutti i tavoli;
- riepilogo e PDF del regolamento ufficiale;
- PDF con gli accoppiamenti;
- loghi di AGM Production e Rione Capocaccia.

La locandina, la data e l'orario non vengono visualizzati nel sito.

## Origine dei dati

Il sito prova a leggere i dati in questo ordine:

1. Web app Google Apps Script collegata al Google Sheet;
2. file locale `data/torneo.xlsx`;
3. copia di emergenza `data/torneo.json`.

Per il collegamento automatico segui `APPS_SCRIPT_SETUP.md`.

## Aggiornare il torneo

Una volta configurato Apps Script basta modificare il Google Sheet:

1. nel foglio `Config`, imposta il numero effettivo di giocatori;
2. nel foglio `Giocatori`, correggi o aggiungi i nomi mantenendo gli ID;
3. apri il foglio risultati corrispondente, per esempio `Risultati_65`;
4. compila `Punti partita`, `Posizione`, `Sóle` e `Note`.

Il sito richiede nuovamente i dati ogni 60 secondi. Un tavolo viene considerato concluso soltanto quando sono presenti tutte le posizioni da 1 a 5.

## Calcoli automatici

Il sito applica:

- 9, 6, 4, 2 e 1 punto in base al piazzamento;
- 2 punti bonus per ogni sóla;
- spareggio per sóle, miglior piazzamento, minor punteggio singolo e minor punteggio totale;
- segnalazione delle perfette parità da risolvere con sorteggio;
- evidenziazione della zona dei primi 25.

## Pubblicazione su GitHub Pages

1. Crea un repository GitHub.
2. Carica nella root tutto il contenuto di questa cartella.
3. Apri `Settings → Pages`.
4. Seleziona il branch principale e la cartella `/ (root)`.
5. Salva e attendi la pubblicazione.

## File principali

- `index.html` — pagina principale;
- `styles.css` — grafica responsive;
- `app.js` — caricamento dati e calcolo classifica;
- `site-config.js` — URL della Web app Apps Script e intervallo di aggiornamento;
- `apps-script/Code.gs` — codice da incollare in Apps Script;
- `APPS_SCRIPT_SETUP.md` — procedura completa;
- `data/torneo.xlsx` — riserva locale;
- `docs/` — regolamento e accoppiamenti;
- `assets/` — loghi e asso di bastoni.
