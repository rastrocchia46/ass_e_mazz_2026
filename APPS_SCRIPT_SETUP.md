# Collegamento automatico con Google Sheets tramite Apps Script

1. Carica `data/torneo.xlsx` su Google Drive e aprilo come Google Fogli.
2. Copia l'ID del foglio dall'indirizzo tra `/d/` e `/edit`.
3. Nel Google Sheet apri **Estensioni > Apps Script**.
4. Sostituisci il contenuto di `Code.gs` con il file `apps-script/Code.gs`.
5. Nel codice sostituisci `INCOLLA_QUI_ID_GOOGLE_SHEET` con l'ID copiato.
6. Salva, seleziona `testConnection` e premi **Esegui** una volta, autorizzando lo script.
7. Seleziona **Esegui il deployment > Nuovo deployment > Applicazione web**.
8. Imposta **Esegui come: Me** e **Chi ha accesso: Chiunque**.
9. Copia l'URL che termina con `/exec`.
10. Apri `site-config.js` e sostituisci `INCOLLA_QUI_URL_APPS_SCRIPT` con l'URL `/exec`.
11. Carica tutti i file del sito su GitHub Pages.

Il browser richiede nuovamente i dati ogni 60 secondi. Per modificare la frequenza cambia `refreshMs` in `site-config.js`.
