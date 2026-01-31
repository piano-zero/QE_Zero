<div align="center">

# 🇮🇹 QE Zero
### Gestione Quadri Economici Opere Pubbliche

![Python](https://img.shields.io/badge/Python-3.x-blue?style=for-the-badge&logo=python)
![GUI](https://img.shields.io/badge/Interface-Tkinter-orange?style=for-the-badge)
![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)
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
