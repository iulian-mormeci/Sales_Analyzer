# Sales Analyzer

Dashboard desktop per l'analisi delle vendite. Carica un file Excel con i tuoi dati e ottieni grafici interattivi con filtri per periodo e canale.

---

## Requisiti di sistema

- Python 3.11 o superiore
- Windows 10/11, macOS 12+, o Linux (Ubuntu 20.04+)

---

## Installazione

```bash
# Clona il repository o scarica i file
cd sales_analyzer

# Crea un ambiente virtuale (consigliato)
python -m venv venv
# Windows:
venv\Scripts\activate
# macOS/Linux:
source venv/bin/activate

# Installa le dipendenze
pip install -r requirements.txt
```

---

## Avvio

```bash
python main.py
```

---

## Generare dati di esempio

Per testare l'app senza avere dati reali, genera un file Excel di esempio con 500 righe:

```bash
python generate_sample_data.py
```

Viene creato il file `sample_data.xlsx` con dati italiani realistici (prodotti, città, canali di vendita, periodo 2022-2024).

---

## Formato Excel supportato

L'app rileva automaticamente le colonne cercando nomi in italiano e inglese. Colonne supportate:

| Campo            | Nomi riconosciuti                                          |
|------------------|------------------------------------------------------------|
| **Data** *(obbligatorio)*     | `data`, `date`, `periodo`, `mese`, `data_vendita`     |
| **Prodotto** *(obbligatorio)* | `prodotto`, `product`, `categoria`, `category`, `articolo` |
| **Ricavo** *(obbligatorio)*   | `ricavo`, `revenue`, `fatturato`, `importo`, `totale` |
| Quantità         | `quantita`, `quantity`, `qty`, `pezzi`, `quantità`        |
| Canale           | `canale`, `channel`, `canale_vendita`, `sales_channel`    |
| Città / Regione  | `città`, `city`, `regione`, `region`, `provincia`          |

Se la rilevazione automatica è ambigua, appare una finestra di dialogo per associare manualmente le colonne.

---

## Viste disponibili

### 📈 Trend
Grafico a linee dei ricavi nel tempo, con selettore di granularità (Giorno / Settimana / Mese / Trimestre) e opzione per sovrapporre le quantità vendute.

### 🏆 Top Prodotti
Grafico a barre orizzontale dei prodotti/categorie per ricavo totale. Selezionabile Top 5 / 10 / 20 / Tutti.

### 📊 Confronto Periodi
Grafico a barre raggruppate con confronto Anno su Anno (YoY) o Mese su Mese (MoM), con etichette di variazione percentuale e riepilogo statistico.

### 🗺 Mappa Geografica
Mappa a bolle delle città italiane dimensionate per ricavo. Richiede una colonna `città` o `regione` nell'Excel.

---

## Build — Creare il file .exe

### Windows

```batch
build.bat
```

Il file `dist\SalesAnalyzer.exe` è l'eseguibile standalone.

### macOS / Linux

```bash
chmod +x build.sh
./build.sh
```

Il file `dist/SalesAnalyzer` è l'eseguibile standalone.

### Build manuale

```bash
pip install pyinstaller
python create_icon.py   # genera icon.ico
pyinstaller build.spec --clean --noconfirm
```

---

## Struttura file

```
sales_analyzer/
├── main.py                  # Applicazione principale
├── updater.py               # Modulo auto-aggiornamento
├── generate_sample_data.py  # Generatore dati di esempio
├── create_icon.py           # Generatore icona
├── build.spec               # Configurazione PyInstaller
├── build.bat                # Script build Windows
├── build.sh                 # Script build macOS/Linux
├── requirements.txt         # Dipendenze Python
├── VERSION                  # Versione corrente (es. "1.0.0")
└── README.md                # Questo file
```

---

## Configurazione auto-aggiornamento

Il file `updater.py` controlla automaticamente gli aggiornamenti all'avvio consultando l'API GitHub Releases.

### Configurare il repository

Apri `updater.py` e modifica la riga:

```python
GITHUB_REPO = "YOUR_USERNAME/YOUR_REPO"
```

sostituendo `YOUR_USERNAME/YOUR_REPO` con il tuo repository GitHub, ad esempio:

```python
GITHUB_REPO = "mario_rossi/sales-analyzer"
```

---

## Come pubblicare un aggiornamento

1. **Aggiorna la versione** nel file `VERSION`:
   ```
   1.1.0
   ```

2. **Esegui il build** per creare il nuovo `SalesAnalyzer.exe`.

3. **Crea una Release su GitHub**:
   - Vai su `https://github.com/TUO_USERNAME/TUO_REPO/releases/new`
   - Imposta il tag: `v1.1.0` (deve corrispondere al contenuto del file `VERSION`)
   - Titolo: `Sales Analyzer v1.1.0`
   - Carica come asset il file `dist/SalesAnalyzer.exe`
   - Pubblica la release

4. **La prossima volta che gli utenti avviano l'app**, questa rileva la nuova versione e mostra il banner di aggiornamento con il pulsante "Scarica e aggiorna".

### Come funziona l'aggiornamento automatico

- All'avvio, l'app verifica silenziosamente la versione più recente tramite l'API GitHub
- Se disponibile, mostra un banner giallo non bloccante in cima alla finestra
- L'utente clicca "Scarica e aggiorna": si apre una finestra con barra di avanzamento
- L'app scarica il nuovo `.exe` e lo rinomina `SalesAnalyzer_new.exe`
- Viene creato uno script `apply_update.bat` che attende la chiusura dell'app, rimpiazza il vecchio eseguibile e riavvia
- Tutti gli eventi di aggiornamento vengono registrati in `app.log`

---

## Log e debug

Il file `app.log` nella stessa cartella del programma registra:
- Caricamento file
- Errori
- Controlli aggiornamento
- Download aggiornamenti

---

## Licenza

MIT License — libero utilizzo, modifica e distribuzione.
