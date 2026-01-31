<div align="center">

# 🇮🇹 QE Zero
### Gestione Quadri Economici Opere Pubbliche

![Python](https://img.shields.io/badge/Python-3.x-blue?style=for-the-badge&logo=python)
![GUI](https://img.shields.io/badge/Interface-Tkinter-orange?style=for-the-badge)
![License](https://img.shields.io/badge/License-GPLv3-green?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Portable_&_Stable-purple?style=for-the-badge)

**Dimentica i fogli di calcolo sparsi e gli errori di arrotondamento.** QE Zero è lo strumento open-source progettato per Ingegneri, Architetti e RUP per redigere, gestire e stampare i Quadri Economici dei lavori pubblici con precisione e velocità.

[Caratteristiche](#-caratteristiche-principali) • [Architettura](#-struttura-e-dati) • [Installazione](#-installazione) • [Come Usare](#-come-usare)

</div>

---

## 🏗 Cos'è QE Zero?

**QE Zero** nasce per semplificare la redazione del Quadro Economico, il cuore finanziario di ogni progetto pubblico.

Invece di lottare con formule Excel che saltano o totali che non quadrano tra "Lavori" e "Somme a disposizione", questo software gestisce la logica contabile in automatico. Inserisci le voci, assegna le categorie e il software calcola Imponibili, IVA, Oneri, Incentivi tecnici e totali generali, garantendo sempre la quadratura del bilancio rispetto allo stanziamento.

## ✨ Caratteristiche Principali

* 💼 **Nativamente Portable:** Il software è progettato per funzionare da chiavetta USB o cartella locale senza installazione. Tutto ciò che serve viaggia con te.
* 🗂 **Separazione Intelligente:** Mantiene rigorosamente separati i dati (`QE_DATI`) dai documenti generati (`QE_STAMPE`) per una gestione pulita e sicura.
* 📐 **Logica Lavori Pubblici:** Gestisce automaticamente la distinzione tra **Quadro A** (Lavori, Oneri Sicurezza) e **Quadro B** (Somme a disposizione, IVA, Spese tecniche).
* 🖨 **Reportistica HTML:** Genera stampe professionali e dettagliate visualizzabili in qualsiasi browser e stampabili in PDF, con header dell'Ente e riepiloghi finanziari.
* 📊 **Controllo Economie:** Calcola in tempo reale la differenza tra l'importo stanziato e il totale del QE, evidenziando economie (verde) o fabbisogni aggiuntivi (rosso).
* 💾 **Database SQLite:** I dati sono salvati in locale su un database relazionale leggero e veloce.

## 📂 Struttura e Dati

Grazie all'ultimo aggiornamento, il progetto adotta un'architettura **"Clean Tree"** che protegge i tuoi dati:

```text
QE_ZERO/
├── qe_zero.exe (o .py)   # Il programma principale
├── QE_DATI/              # 🔒 Qui risiede il Database (NON toccare o cancellare)
│   └── qe_zero.db
└── QE_STAMPE/            # 📄 Qui finiscono i tuoi Report HTML/PDF
    ├── Stampa_QE_1.html
    └── Stampa_QE_2.html
```

*Questa struttura permette di svuotare la cartella delle stampe quando vuoi, senza mai rischiare di perdere il database dei progetti.*

## 🚀 Installazione

### Prerequisiti
* Python 3.x installato sul sistema (se si utilizza la versione sorgente).
* Nessun prerequisito se si utilizza l'eseguibile compilato.

### Passaggi (Versione Sorgente)

1.  **Clona il repository** (o scarica lo zip):
    ```bash
    git clone https://github.com/piano-zero/QE_Zero.git
    ```

2.  **Librerie:**
    QE Zero è leggero e utilizza le librerie standard di Python (`tkinter`, `sqlite3`, `os`, `webbrowser`). Non sono richieste installazioni di pacchetti pesanti.

3.  **Avvia l'applicazione:**
    ```bash
    python qe_zero.py
    ```
    *Al primo avvio, il software creerà automaticamente le cartelle `QE_DATI` e `QE_STAMPE`.*

## 📖 Come Usare

1.  **Configurazione Ente:** Imposta i dati dell'Amministrazione (Comune, Ente, Indirizzo) per le intestazioni delle stampe.
2.  **Nuovo Progetto:** Crea un progetto inserendo Oggetto, CUP e Importo Stanziato.
3.  **Gestione QE:** All'interno del progetto, crea una revisione del Quadro Economico (es. "Progetto Esecutivo").
4.  **Inserimento Voci:** Aggiungi le voci di spesa specificando:
    * Descrizione e Importo.
    * Categoria (Lavori o Somme a disposizione).
    * Aliquote (IVA, Oneri previdenziali, etc.).
5.  **Stampa:** Clicca su "Genera Report". Il file verrà salvato nella cartella `QE_STAMPE` e aperto automaticamente nel tuo browser predefinito.

## 🤝 Contribuire

Il progetto è aperto a suggerimenti! Se sei un tecnico o uno sviluppatore:

1.  Fai un **Fork** del progetto.
2.  Crea un branch (`git checkout -b feature/MiglioramentoGrafico`).
3.  Fai **Commit** (`git commit -m 'Migliorato layout di stampa'`).
4.  Fai **Push** (`git push origin feature/MiglioramentoGrafico`).
5.  Apri una **Pull Request**.

## 📄 Licenza

Distribuito sotto licenza **GNU General Public License v3.0**.

---

<div align="center">
  
  Created with ❤️ by [pianozero](https://github.com/piano-zero)
  
  *Se questo progetto ti è stato utile, lascia una ⭐️ al repository!*


</div>
