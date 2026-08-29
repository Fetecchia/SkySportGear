# SkySportGear

Prototipo per la gestione e assegnazione di materiale audio/video (videocamere,
microfoni, luci, batterie, cavalletti, ronin, obiettivi, vario) ai cameraman.

## Come pubblicarlo su GitHub Pages (nessun IT richiesto)

1. **Crea un nuovo repository su GitHub**
   - Vai su https://github.com/new
   - Dagli un nome, es. `skysportgear`
   - Lascialo **Public** (necessario per GitHub Pages gratuito) oppure Private se hai GitHub Pro/Team
   - NON aggiungere README/gitignore da GitHub (li abbiamo già)

2. **Carica tutti i file di questa cartella nel repository**
   - Opzione più semplice: nella pagina del repo appena creato, clicca
     "uploading an existing file" e trascina dentro **tutto il contenuto**
     di questa cartella (compresi i file nascosti `.gitignore` e la cartella
     `.github`).
   - Se il tuo file manager non mostra/trascina la cartella `.github`
     (è "nascosta" perché inizia con un punto), crea a mano il file su
     GitHub: "Add file" → "Create new file" → come nome scrivi
     `.github/workflows/deploy.yml` (GitHub crea le cartelle da solo) e
     incolla il contenuto del file `deploy.yml` che trovi qui.
   - In alternativa, se hai Git installato:
     ```
     git init
     git add .
     git commit -m "Prima versione SkySportGear"
     git branch -M main
     git remote add origin https://github.com/TUO-UTENTE/skysportgear.git
     git push -u origin main
     ```

3. **Attiva GitHub Pages con build automatica**
   - Nel repository vai su **Settings → Pages**
   - Alla voce "Source" scegli **GitHub Actions** (non "Deploy from a branch")
   - Salva

4. **Aspetta la build**
   - Vai sulla tab **Actions** del repository: vedrai il workflow
     "Deploy to GitHub Pages" partire automaticamente dopo il push.
   - Quando diventa verde ✅ (1-2 minuti), torna su Settings → Pages: troverai
     l'URL pubblico dell'app, tipo:
     `https://TUO-UTENTE.github.io/skysportgear/`

5. **Aggiungilo come scheda su Teams**
   - Nel canale Teams che userete per il test, clicca **+** in alto
   - Scegli **Website**
   - Incolla l'URL di GitHub Pages
   - Dai un nome alla scheda, es. "SkySportGear"

Da questo momento chiunque nel canale può aprire la scheda e usare l'app.

## Aggiornare l'app in futuro

Ogni volta che il codice viene modificato (es. per collegare SharePoint),
basta caricare/pushare i nuovi file: GitHub Actions ricostruisce e
ripubblica automaticamente il sito nel giro di 1-2 minuti, senza dover
rifare nulla su Teams.

## Nota importante sui dati

Questa versione salva i dati **in memoria nel browser**: si azzerano
ricaricando la pagina e non sono condivisi tra utenti diversi. È pensata
per un primo test dell'interfaccia. Il prossimo passo per un test reale con
più persone è collegare l'app a delle Liste di SharePoint (vedi
conversazione con Claude per i dettagli).
