visual studio code

Per far funzionare tutta la tua applicazione Node.js/Express/EJS con filtri multipli SI/NO esclusivi e tabelle su server o locale, servono questi elementi chiave:

1. File e struttura cartelle
Assicurati di avere questa struttura (esempio):

text
tuo-progetto/
│
├── server.js
├── package.json
├── SCN_Erogazioni.xlsx
├── SCN_Erogazioni_Tipologia.xlsx
├── TipologiaSoggetti.xlsx
└── views/
    └── tabella.ejs
2. Node.js e NPM installati
Sul server (o PC):

Versione Node.js da riga di comando:

text
node -v
npm -v
3. Installare le dipendenze Node
Dentro la cartella del progetto, crea un file package.json (se non ce l’hai):

json
{
  "name": "taomau-app",
  "version": "1.0.0",
  "main": "server.js",
  "dependencies": {
    "express": "^4.18.2",
    "ejs": "^3.1.9",
    "xlsx": "^0.18.5"
  }
}
Poi in terminale:

text
npm install
(Installerà express, ejs, xlsx…)

4. I File Excel devono essere presenti
I file SCN_Erogazioni.xlsx, SCN_Erogazioni_Tipologia.xlsx, TipologiaSoggetti.xlsx devono stare nella stessa cartella dove hai server.js

Verifica che i loro nomi corrispondano ESATTAMENTE (case sensitive!) nel codice.

5. Codice giusto su server.js e views/tabella.ejs
Usa esattamente i file che ti ho fornito sopra (sia server.js che tabella.ejs).

Nella cartella /views deve esserci il file tabella.ejs aggiornato.

I dati Excel devono contenere colonne coerenti con i nomi usati (Soggetto, Azienda, Politico, ecc.).

6. Avvio dell’applicazione
Nel terminale (sempre nella cartella), lancia:

text
node server.js
Se Ok: browser su http://localhost:3000/tabella/erogazioni_tipologia mostra la tabella coi filtri.

Se vuoi tenerla sempre attiva su un server:

text
npm install -g pm2
pm2 start server.js
pm2 save
pm2 startup
7. (OPZIONALE) Hosting/Server pubblico
Se vuoi la app su Internet:

Serve un server VPS/Cloud che supporti Node.js (Netsons CloudNode/VPS, Hetzner, Aruba etc).

Trasferisci tutta la cartella via SFTP/git.

Ripeti installazione dipendenze (npm install) e avvio.

Configura DNS/dominio/se vuoi nginx come proxy (vedi passaggi dedicati).

8. Permessi file
Se usi Linux, assicurati che i file abbiano permessi di lettura per l’utente che lancia node.

Ricorda:
Se cambi il file Excel, riavvia il server (CTRL+C, poi di nuovo node server.js)

Per ogni errore in console, copia qui il messaggio preciso: ti aiuto subito a risolvere!

In sintesi:

Una cartella con TUTTI i file dati/JS/EJS

Dipendenze installate (npm install)

Lancia node

Naviga dal browser

I checkbox SI/NO saranno esclusivi per ogni filtro

-----------------------------------------

http://localhost:3000/tabella/erogazioni

http://localhost:3000/tabella/tipologia

http://localhost:3000/tabella/erogazioni_tipologia

Esempio: http://localhost:3000/tabella/erogazioni?AnnoErogazione=2024