# Generatore Automatico di Attestati PDF

Questo script Python automatizza la creazione in serie di attestati di partecipazione in formato PDF. Partendo da fogli Excel contenenti i dati dei corsisti, genera documenti ad alta qualità sovrapponendo dinamicamente testi (nomi, ore, date) e loghi su uno sfondo vettoriale, ricalcando l'esatta formattazione grafica richiesta.

## 🛠 Prerequisiti e Installazione

Per utilizzare il programma è necessario avere Python installato sul computer. Prima della prima esecuzione, installa le librerie esterne richieste aprendo il Terminale o il Prompt dei comandi e digitando:

```bash
pip install pandas openpyxl fpdf2

```

## 📁 Struttura dei File Necessari

Affinché lo script funzioni correttamente, i seguenti file devono essere posizionati nella stessa cartella del programma:

* **`sfondo_pulito.svg`**: L'immagine di sfondo vettoriale del certificato, rigorosamente priva di testi variabili.
* **`logo.png`**: (Opzionale) Il logo aziendale o istituzionale che verrà stampato in alto a sinistra.
* **`font_custom.ttf`** e **`font_custom_bold.ttf`**: (Opzionali) I file del font per riprodurre lo stile tipografico esatto. Se non vengono trovati, lo script ripiegherà automaticamente sul font standard di sistema (Helvetica).
* **File Excel**: I fogli di calcolo con i registri dei partecipanti (es. `PARTECIPANTI_PRESENZA_16ORE.xlsx`). È obbligatorio che la prima riga contenga le intestazioni **NOME** e **COGNOME** scritte in maiuscolo.

## ⚙️ Configurazione dei Corsi

La gestione dei corsi e delle date avviene direttamente nel codice sorgente modificando il dizionario `FILE_DATI`. Qui puoi associare il nome esatto del file Excel alle ore e alle giornate specifiche del corso:

```python
FILE_DATI = {
    "PARTECIPANTI_LIVE_9ORE.xlsx": {
        "ore": "9",
        "date": "11, 20 maggio e 3 giugno 2026"
    },
    "PARTECIPANTI_PRESENZA_16ORE.xlsx": {
        "ore": "16",
        "date": "12, 13 e 14 ottobre 2026" 
    }
}

```

Puoi aggiungere infiniti file Excel a questa lista seguendo la medesima struttura.

## 📐 Calibrazione Grafica

Il layout è impostato per calcolare l'ingombro del testo e centrarlo matematicamente su un foglio A4 orizzontale. Se hai bisogno di aggiustare visivamente il risultato, puoi intervenire su due livelli all'interno dello script:

1. **Spostamento Globale:** Modifica la variabile `SPOSTAMENTO_VERTICALE`. Inserendo un numero positivo (es. `5`) tutto il blocco di scritte scenderà di 5 millimetri. Un numero negativo (es. `-5`) alzerà l'intero blocco.
2. **Margini Logo:** Modifica le variabili `MARGINE_SUPERIORE` e `MARGINE_LATERALE` per calibrare i rientri del logo in alto a sinistra rispetto al bordo del foglio.

## 🚀 Logica di Esecuzione

Avviando lo script, il programma seguirà automaticamente questo iter:

1. **Pulizia dell'ambiente:** Elimina completamente la cartella `Attestati_PDF` (se presente) e ne crea una nuova, vuota e pulita. *Attenzione: affinché questa operazione vada a buon fine, assicurati di non avere PDF aperti nel visualizzatore.*
2. **Lettura ed elaborazione:** Scansiona i file Excel dichiarati, ignorando le righe vuote e formattando i nomi in stile "Title Case" (iniziale maiuscola, es. "Mario Rossi").
3. **Generazione PDF:** Per ogni corsista, costruisce il certificato sovrapponendo sfondo, testi centrati e logo.
4. **Salvataggio e Omonimie:** Salva il file nominandolo in modo sicuro (es. `Attestato_Mario_Rossi_16.pdf`). Se individua un omonimo all'interno dello stesso corso, protegge i file aggiungendo un suffisso numerico progressivo finale (es. `_1.pdf`).