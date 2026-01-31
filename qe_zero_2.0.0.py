# QE Zero - Gestione Quadri Economici Opere Pubbliche
# Versione 2.0.0 per la compilazione
#
# Copyright (C) 2025 Rodolfo Sabelli
#
# Questo programma è software libero: puoi ridistribuirlo e/o modificarlo
# secondo i termini della GNU General Public License versione 3 o della
# European Union Public License versione 1.2 (a tua scelta).
#
# Questo programma è distribuito nella speranza che sia utile,
# ma SENZA ALCUNA GARANZIA; senza neppure la garanzia implicita di
# COMMERCIABILITÀ o IDONEITÀ PER UN PARTICOLARE SCOPO.
#
# Vedi LICENSE.txt per il testo completo delle licenze.
#
# Contatti: rodolfo.sabelli@gmail.com
# Repository: [URL GITHUB/GITLAB]

import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
from tkinter import simpledialog
import sqlite3
import datetime
import os
import webbrowser
import csv
import urllib.request
import shutil
import subprocess
import platform
from itertools import groupby

# =============================================================================
# 1. DATABASE MANAGER
# =============================================================================
class DatabaseManager:
    def __init__(self, db_name="qe_zero.db"):
        """Inizializza il database manager con percorsi ottimizzati"""
        # PATH LOGIC OTTIMIZZATA
        program_dir = os.path.dirname(os.path.abspath(__file__))
        local_path = os.path.join(program_dir, "QE_DATI")

        # Modalità Portable: cartella locale ha priorità
        if os.path.exists(local_path):
            self.documents_path = local_path
            # NUOVO: Se siamo in portable, la cartella stampe sta nella directory del programma
            # Risultato: .../QE_ZERO/QE_STAMPE
            self.stampe_path = os.path.join(program_dir, "QE_STAMPE")
        else:
            # Modalità Standard: cartella in Documenti utente
            base_docs = os.path.expanduser("~/Documents")
            self.documents_path = os.path.join(base_docs, "QE_DATI")
            
            # NUOVO: Se siamo in standard, la cartella stampe sta in Documenti
            # Risultato: ~/Documents/QE_STAMPE (accanto a QE_DATI)
            self.stampe_path = os.path.join(base_docs, "QE_STAMPE")
            
            # Creazione cartella DATI (Logica esistente)
            if not os.path.exists(self.documents_path):
                try:
                    os.makedirs(self.documents_path)
                except OSError as e:
                    messagebox.showerror(
                        "Errore", 
                        f"Impossibile creare la cartella {self.documents_path}\n{e}"
                    )

        # --- NUOVO CODICE AGGIUNTO: CREAZIONE CARTELLA STAMPE ---
        # Questo blocco viene eseguito a prescindere dalla modalità (Portable o Standard)
        # Verifica se QE_STAMPE esiste, altrimenti la crea
        if not os.path.exists(self.stampe_path):
            try:
                os.makedirs(self.stampe_path)
            except OSError as e:
                # Gestione errore non bloccante (puoi cambiare in messagebox se preferisci)
                print(f"Errore nella creazione della cartella stampe: {e}")
        # --------------------------------------------------------

        self.db_path = os.path.join(self.documents_path, db_name)   
        self.conn = sqlite3.connect(self.db_path)
        self.conn.execute("PRAGMA foreign_keys = 1")
        
        # Sequenza inizializzazione ottimizzata
        self.crea_tabelle()
        self.check_aggiornamento_db_allegati()
        self.migra_db_1_3()
        self.popola_dati_base()
        self.popola_demo_se_vuoto()

    def crea_tabelle(self):
        """Crea tutte le tabelle del database con schema ottimizzato"""
        c = self.conn.cursor()
        
        # Tabella normative
        c.execute('''CREATE TABLE IF NOT EXISTS normative (
            id INTEGER PRIMARY KEY AUTOINCREMENT, 
            nome TEXT UNIQUE, 
            descrizione TEXT
        )''')
        
        # Tabella progetti
        c.execute('''CREATE TABLE IF NOT EXISTS progetti (
            id INTEGER PRIMARY KEY AUTOINCREMENT, 
            normativa_id INTEGER, 
            cup TEXT, 
            anno INTEGER, 
            titolo TEXT, 
            importo REAL, 
            FOREIGN KEY (normativa_id) REFERENCES normative (id)
        )''')
        
        # Tabella quadri economici
        c.execute('''CREATE TABLE IF NOT EXISTS quadri_economici (
            id INTEGER PRIMARY KEY AUTOINCREMENT, 
            progetto_id INTEGER, 
            nome_versione TEXT, 
            data_creazione TEXT, 
            note TEXT, 
            FOREIGN KEY (progetto_id) REFERENCES progetti (id) ON DELETE CASCADE
        )''')
        
        # Tabella voci (include flag_calcolo_montante)
        c.execute('''CREATE TABLE IF NOT EXISTS voci (
            id INTEGER PRIMARY KEY AUTOINCREMENT, 
            qe_id INTEGER, 
            codice_padre TEXT, 
            codice_completo TEXT, 
            descrizione TEXT, 
            tipo TEXT, 
            valore_imponibile REAL, 
            is_percentuale INTEGER, 
            perc_oneri REAL, 
            includi_oneri_in_iva INTEGER, 
            perc_iva REAL, 
            flag_base_asta INTEGER, 
            flag_soggetto_ribasso INTEGER, 
            macro_base_calcolo TEXT, 
            flag_calcolo_montante INTEGER DEFAULT 0, 
            FOREIGN KEY (qe_id) REFERENCES quadri_economici (id) ON DELETE CASCADE
        )''')
        
        # Tabella catalogo voci
        c.execute('''CREATE TABLE IF NOT EXISTS catalogo_voci (
            id INTEGER PRIMARY KEY AUTOINCREMENT, 
            normativa_id INTEGER, 
            codice TEXT, 
            macro_gruppo INTEGER, 
            descrizione TEXT, 
            UNIQUE(normativa_id, codice), 
            FOREIGN KEY (normativa_id) REFERENCES normative (id) ON DELETE CASCADE
        )''')
        
        # Tabella configurazione
        c.execute('''CREATE TABLE IF NOT EXISTS configurazione (
            chiave TEXT PRIMARY KEY, 
            valore TEXT
        )''')
        
        # Tabella allegati (include descrizione)
        c.execute('''CREATE TABLE IF NOT EXISTS allegati_qe (
            id INTEGER PRIMARY KEY AUTOINCREMENT, 
            qe_id INTEGER, 
            nome_file TEXT, 
            tipo_file TEXT, 
            dati BLOB, 
            data_caricamento TEXT,
            descrizione TEXT DEFAULT '',
            FOREIGN KEY (qe_id) REFERENCES quadri_economici (id) ON DELETE CASCADE
        )''')
        
        self.conn.commit()

    def check_aggiornamento_db_allegati(self):
        """Migrazione: aggiunge colonna descrizione a tabella allegati se mancante"""
        try:
            self.conn.execute("SELECT descrizione FROM allegati_qe LIMIT 1")
        except sqlite3.OperationalError:
            try:
                self.conn.execute("ALTER TABLE allegati_qe ADD COLUMN descrizione TEXT DEFAULT ''")
                self.conn.commit()
                print("✓ Migrazione allegati: colonna 'descrizione' aggiunta")
            except sqlite3.OperationalError:
                pass  # Colonna già esistente
    
    def migra_db_1_3(self):
        """Migrazione v1.3: aggiunge flag_calcolo_montante per DB legacy"""
        try:
            self.conn.execute("SELECT flag_calcolo_montante FROM voci LIMIT 1")
        except sqlite3.OperationalError:
            try:
                self.conn.execute("ALTER TABLE voci ADD COLUMN flag_calcolo_montante INTEGER DEFAULT 0")
                self.conn.commit()
                print("✓ Migrazione v1.3: colonna 'flag_calcolo_montante' aggiunta")
            except sqlite3.OperationalError:
                pass  # Colonna già esistente

    def popola_dati_base(self):
        """Popola dati iniziali: configurazione e normative standard"""
        # Configurazione base
        self.conn.execute(
            "INSERT OR IGNORE INTO configurazione (chiave, valore) VALUES (?, ?)", 
            ("admin_password", "admin")
        )
        
        # Normative standard
        normative = [
            ("D.Lgs. 36/2023 - Opere", "Realizzazione di opere pubbliche"),
            ("D.Lgs. 36/2023 - Beni e Servizi", "Acquisizione di beni e servizi"),
            ("DPR 207/2010", "Ex Regolamento"),
            ("FESR 2021-2027", "Realizzazione di opere pubbliche")
        ]
        
        for n, d in normative: 
            self.conn.execute(
                "INSERT OR IGNORE INTO normative (nome, descrizione) VALUES (?, ?)", 
                (n, d)
            )
        
        self.conn.commit()
        
        # Popolamento catalogo D.Lgs. 36/2023 - Opere
        try:
            id_36_opere = self.conn.execute("SELECT id FROM normative WHERE nome='D.Lgs. 36/2023 - Opere'").fetchone()[0]
            if self.conn.execute("SELECT count(*) FROM catalogo_voci WHERE normativa_id=?", (id_36_opere,)).fetchone()[0] == 0:
                dati = [(id_36_opere, "A", 1, "IMPORTO DEI LAVORI"), (id_36_opere, "B", 1, "IMPORTO DEI COSTI DELLA SICUREZZA"), (id_36_opere, "C", 1, "IMPORTO PER IL CONTRASTO ALLA CRIMINALITÀ"), (id_36_opere, "D", 1, "OPERE DI MITIGAZIONE E COSTI AMBIENTALI"),
                        (id_36_opere, "E.01", 2, "Lavori in amministrazione diretta"), (id_36_opere, "E.02", 2, "Rilievi e indagini (a cura della S.A.)"), (id_36_opere, "E.03", 2, "Rilievi e indagini (a cura del Progettista)"), (id_36_opere, "E.04", 2, "Allacciamenti pubblici servizi"), (id_36_opere, "E.05", 2, "Imprevisti"), (id_36_opere, "E.06", 2, "Accantonamenti"), (id_36_opere, "E.07", 2, "Acquisizione aree/espropri"), (id_36_opere, "E.08", 2, "Spese tecniche"), (id_36_opere, "E.09", 2, "Spese attività tecnico-amministrative"), (id_36_opere, "E.10", 2, "Spese Art. 45"), (id_36_opere, "E.11", 2, "Commissioni"), (id_36_opere, "E.12", 2, "Pubblicità"), (id_36_opere, "E.13", 2, "Prove laboratorio"), (id_36_opere, "E.14", 2, "Collaudi"), (id_36_opere, "E.15", 2, "Verifica archeologica"), (id_36_opere, "E.16", 2, "Tutela giurisdizionale"), (id_36_opere, "E.17", 2, "Opere artistiche")]
                self.conn.executemany("INSERT OR IGNORE INTO catalogo_voci (normativa_id, codice, macro_gruppo, descrizione) VALUES (?, ?, ?, ?)", dati)
        except: pass
        try:
            id_36_serv = self.conn.execute("SELECT id FROM normative WHERE nome='D.Lgs. 36/2023 - Beni e Servizi'").fetchone()[0]
            if self.conn.execute("SELECT count(*) FROM catalogo_voci WHERE normativa_id=?", (id_36_serv,)).fetchone()[0] == 0:
                dati = [(id_36_serv, "1.01", 1, "Importo servizi/forniture"), (id_36_serv, "1.02", 1, "Oneri sicurezza interferenziali"), (id_36_serv, "2.01", 2, "Lavori in economia"), (id_36_serv, "2.02", 2, "Imprevisti"), (id_36_serv, "2.03", 2, "Spese tecniche"), (id_36_serv, "2.04", 2, "Pubblicità/Gara"), (id_36_serv, "2.05", 2, "Revisione prezzi")]
                self.conn.executemany("INSERT OR IGNORE INTO catalogo_voci (normativa_id, codice, macro_gruppo, descrizione) VALUES (?, ?, ?, ?)", dati)
        except: pass
        try:
            id_207 = self.conn.execute("SELECT id FROM normative WHERE nome='DPR 207/2010'").fetchone()[0]
            if self.conn.execute("SELECT count(*) FROM catalogo_voci WHERE normativa_id=?", (id_207,)).fetchone()[0] == 0:
                dati = [(id_207, "1.01", 1, "Lavori"), (id_207, "1.02", 1, "Sicurezza"), (id_207, "2.01", 2, "Lavori in economia"), (id_207, "2.02", 2, "Rilievi, accertamenti e indagini"), (id_207, "2.03", 2, "Allacciamenti"), (id_207, "2.04", 2, "Imprevisti"), (id_207, "2.05", 2, "Acquisizione aree"), (id_207, "2.06", 2, "Revisione prezzi"), (id_207, "2.07", 2, "Spese tecniche"), (id_207, "2.08", 2, "Spese per attività amministrative"), (id_207, "2.09", 2, "Commissioni di gara"), (id_207, "2.10", 2, "Pubblicità"), (id_207, "2.11", 2, "Collaudi")]
                self.conn.executemany("INSERT OR IGNORE INTO catalogo_voci (normativa_id, codice, macro_gruppo, descrizione) VALUES (?, ?, ?, ?)", dati)
        except: pass
        try:
            id_fesr = self.conn.execute("SELECT id FROM normative WHERE nome='FESR 2021-2027'").fetchone()[0]
            if self.conn.execute("SELECT count(*) FROM catalogo_voci WHERE normativa_id=?", (id_fesr,)).fetchone()[0] == 0:
                dati = [(id_fesr, "A.01", 1, "Lavori"), (id_fesr, "A.02", 1, "Sicurezza"), (id_fesr, "B.01", 2, "Economia"), (id_fesr, "B.02", 2, "Rilievi"), (id_fesr, "B.03", 2, "Allacciamenti"), (id_fesr, "B.04", 2, "Imprevisti"), (id_fesr, "B.05", 2, "Aree"), (id_fesr, "B.06", 2, "Accantonamenti"), (id_fesr, "B.07", 2, "Spese tecniche"), (id_fesr, "B.08", 2, "Consulenza"), (id_fesr, "B.09", 2, "Commissioni"), (id_fesr, "B.10", 2, "Pubblicità"), (id_fesr, "B.11", 2, "Collaudi"), (id_fesr, "B.12", 2, "Forniture")]
                self.conn.executemany("INSERT OR IGNORE INTO catalogo_voci (normativa_id, codice, macro_gruppo, descrizione) VALUES (?, ?, ?, ?)", dati)
        except: pass
        self.conn.commit()
        
    def popola_demo_se_vuoto(self):
        """Crea progetto demo solo se il database è completamente vuoto"""
        count = self.conn.execute("SELECT count(*) FROM progetti").fetchone()[0]
        
        if count == 0:
            try:
                id_36 = self.conn.execute(
                    "SELECT id FROM normative WHERE nome='D.Lgs. 36/2023 - Opere'"
                ).fetchone()
                
                if not id_36:
                    return
                
                id_36 = id_36[0]
                
                # Inserisci progetto demo
                self.conn.execute(
                    """INSERT INTO progetti 
                    (normativa_id, cup, anno, titolo, importo) 
                    VALUES (?,?,?,?,?)""", 
                    (id_36, "H61I24000130006", 2025, "Rete di attracchi via mare", 5623000.0)
                )
                
                pid = self.conn.execute("SELECT last_insert_rowid()").fetchone()[0]
                
                # Inserisci QE demo
                self.conn.execute(
                    """INSERT INTO quadri_economici 
                    (progetto_id, nome_versione, data_creazione, note) 
                    VALUES (?,?,?,?)""", 
                    (pid, "1. PFTE", datetime.date.today().strftime("%d/%m/%Y"), "Versione 1.0")
                )
                
                qid1 = self.conn.execute("SELECT last_insert_rowid()").fetchone()[0]
                
                # Inserisci voce demo
                self.conn.execute(
                    """INSERT INTO voci 
                    (qe_id, codice_padre, codice_completo, descrizione, tipo, 
                    valore_imponibile, is_percentuale, perc_oneri, includi_oneri_in_iva, 
                    perc_iva, flag_base_asta, flag_soggetto_ribasso, macro_base_calcolo, 
                    flag_calcolo_montante) 
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", 
                    (qid1, "A", "A.01", "Lavori Edili", "fisso", 100000.0, 
                     0, 0, 0, 10.0, 1, 1, "", 1)
                )
                
                self.conn.commit()
                print("✓ Progetto demo creato con successo")
                
            except Exception as e:
                print(f"Errore creazione progetto demo: {e}")

    # --- CRUD OPERATIONS: CONFIGURAZIONE E NORMATIVE ---
    
    def get_normative(self):
        """Recupera tutte le normative ordinate per ID"""
        return self.conn.execute(
            "SELECT id, nome, descrizione FROM normative ORDER BY id"
        ).fetchall()
    
    def get_config(self, k):
        """Recupera valore di configurazione per chiave"""
        r = self.conn.execute(
            "SELECT valore FROM configurazione WHERE chiave=?", (k,)
        ).fetchone()
        return r[0] if r else ""
    
    def set_config(self, k, v):
        """Salva valore di configurazione"""
        self.conn.execute(
            "INSERT OR REPLACE INTO configurazione (chiave, valore) VALUES (?, ?)", 
            (k, v)
        )
        self.conn.commit()
    
    def inserisci_normativa(self, n, d):
        """Inserisce nuova normativa"""
        try:
            self.conn.execute(
                "INSERT INTO normative (nome, descrizione) VALUES (?, ?)", 
                (n, d)
            )
            self.conn.commit()
            return True
        except sqlite3.IntegrityError:
            return False
    
    def aggiorna_normativa(self, nid, n, d):
        """Aggiorna normativa esistente"""
        try:
            self.conn.execute(
                "UPDATE normative SET nome=?, descrizione=? WHERE id=?", 
                (n, d, nid)
            )
            self.conn.commit()
            return True
        except sqlite3.IntegrityError:
            return False
    
    def elimina_normativa(self, nid):
        """Elimina normativa (e progetti collegati in cascade)"""
        self.conn.execute("DELETE FROM normative WHERE id=?", (nid,))
        self.conn.commit()
    
    def duplica_normativa(self, old_id, new_name, new_desc):
        """Duplica normativa e il suo catalogo voci"""
        try:
            self.conn.execute(
                "INSERT INTO normative (nome, descrizione) VALUES (?, ?)", 
                (new_name, new_desc)
            )
            new_id = self.conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            
            self.conn.execute(
                """INSERT INTO catalogo_voci 
                (normativa_id, codice, macro_gruppo, descrizione) 
                SELECT ?, codice, macro_gruppo, descrizione 
                FROM catalogo_voci WHERE normativa_id = ?""", 
                (new_id, old_id)
            )
            
            self.conn.commit()
            return True
        except Exception as e:
            print(f"Errore duplicazione normativa: {e}")
            return False

    # --- CRUD OPERATIONS: CATALOGO VOCI ---
    
    def get_catalogo(self, normativa_id, mid=None):
        """Recupera catalogo voci per normativa (opzionalmente filtrato per macro)"""
        sql = """SELECT id, codice, macro_gruppo, descrizione 
                 FROM catalogo_voci WHERE normativa_id=?"""
        params = [normativa_id]
        
        if mid is not None:
            sql += " AND macro_gruppo=?"
            params.append(mid)
        
        sql += " ORDER BY codice"
        
        return self.conn.execute(sql, params).fetchall()
    
    def aggiorna_catalogo(self, nid, c, m, d):
        """Aggiorna o inserisce voce di catalogo"""
        r = self.conn.execute(
            "SELECT id FROM catalogo_voci WHERE normativa_id=? AND codice=?", 
            (nid, c)
        ).fetchone()
        
        if not r:
            self.conn.execute(
                """INSERT INTO catalogo_voci 
                (normativa_id, codice, macro_gruppo, descrizione) 
                VALUES (?,?,?,?)""", 
                (nid, c, m, d)
            )
        else:
            self.conn.execute(
                """UPDATE catalogo_voci 
                SET macro_gruppo=?, descrizione=? WHERE id=?""", 
                (m, d, r[0])
            )
        
        self.conn.commit()
    
    def aggiorna_voce_catalogo_id(self, cat_id, nid, c, m, d):
        """Aggiorna voce catalogo per ID (o inserisce se None)"""
        if cat_id:
            self.conn.execute(
                """UPDATE catalogo_voci 
                SET codice=?, macro_gruppo=?, descrizione=? WHERE id=?""", 
                (c, m, d, cat_id)
            )
        else:
            self.conn.execute(
                """INSERT INTO catalogo_voci 
                (normativa_id, codice, macro_gruppo, descrizione) 
                VALUES (?,?,?,?)""", 
                (nid, c, m, d)
            )
        self.conn.commit()
    
    def elimina_voce_catalogo(self, cat_id):
        """Elimina voce dal catalogo"""
        self.conn.execute("DELETE FROM catalogo_voci WHERE id=?", (cat_id,))
        self.conn.commit()

    # --- CRUD OPERATIONS: PROGETTI ---
    
    def get_prossimo_codice(self, qe_id, codice_padre):
        """Calcola il prossimo codice disponibile per una categoria"""
        rows = self.conn.execute(
            "SELECT codice_completo FROM voci WHERE qe_id=? AND codice_padre=?", 
            (qe_id, codice_padre)
        ).fetchall()
        
        numeri = [
            int(r[0].split('.')[-1]) 
            for r in rows 
            if r[0].split('.')[-1].isdigit()
        ]
        
        prossimo = max(numeri) + 1 if numeri else 1
        return f"{codice_padre}.{prossimo:02d}"
    
    def inserisci_progetto(self, nid, c, a, t, i):
        """Inserisce nuovo progetto"""
        self.conn.execute(
            """INSERT INTO progetti 
            (normativa_id, cup, anno, titolo, importo) 
            VALUES (?,?,?,?,?)""", 
            (nid, c, a, t, i)
        )
        self.conn.commit()
    
    def aggiorna_progetto_dati(self, pid, c, a, t, i):
        """Aggiorna dati progetto esistente"""
        self.conn.execute(
            """UPDATE progetti 
            SET cup=?, anno=?, titolo=?, importo=? 
            WHERE id=?""", 
            (c, a, t, i, pid)
        )
        self.conn.commit()
    
    def elimina_progetto(self, pid):
        """Elimina progetto (QE in cascade)"""
        self.conn.execute("DELETE FROM progetti WHERE id=?", (pid,))
        self.conn.commit()
    
    def get_tutti_progetti(self):
        """Recupera tutti i progetti con info normativa"""
        return self.conn.execute(
            """SELECT p.id, p.cup, p.anno, p.titolo, p.importo, n.nome 
            FROM progetti p 
            JOIN normative n ON p.normativa_id = n.id 
            ORDER BY p.id DESC"""
        ).fetchall()
    
    def get_progetto_by_id(self, pid):
        """Recupera singolo progetto per ID"""
        return self.conn.execute(
            "SELECT * FROM progetti WHERE id=?", (pid,)
        ).fetchone()

    # --- CRUD OPERATIONS: QUADRI ECONOMICI ---
    
    def inserisci_qe(self, pid, n, nt):
        """Inserisce nuovo QE"""
        self.conn.execute(
            """INSERT INTO quadri_economici 
            (progetto_id, nome_versione, data_creazione, note) 
            VALUES (?,?,?,?)""", 
            (pid, n, datetime.date.today().strftime("%d/%m/%Y"), nt)
        )
        self.conn.commit()
    
    def aggiorna_qe(self, qid, n, nt):
        """Aggiorna QE esistente"""
        self.conn.execute(
            """UPDATE quadri_economici 
            SET nome_versione=?, note=? 
            WHERE id=?""", 
            (n, nt, qid)
        )
        self.conn.commit()
    
    def elimina_qe(self, qid):
        """Elimina QE (voci in cascade)"""
        self.conn.execute("DELETE FROM quadri_economici WHERE id=?", (qid,))
        self.conn.commit()
    
    def duplica_qe(self, qid, n):
        """Duplica QE con tutte le sue voci"""
        r = self.conn.execute(
            "SELECT * FROM quadri_economici WHERE id=?", (qid,)
        ).fetchone()
        
        if not r:
            return
        
        # Inserisci nuovo QE
        self.conn.execute(
            """INSERT INTO quadri_economici 
            (progetto_id, nome_versione, data_creazione, note) 
            VALUES (?,?,?,?)""", 
            (r[1], n, datetime.date.today().strftime("%d/%m/%Y"), f"Copia di {r[2]}")
        )
        new_qid = self.conn.execute("SELECT last_insert_rowid()").fetchone()[0]
        
        # Copia tutte le voci
        voci = self.conn.execute("SELECT * FROM voci WHERE qe_id=?", (qid,)).fetchall()
        
        for v in voci:
            # Gestione compatibilità: flag_calcolo_montante potrebbe non esistere
            f_mont = v[14] if len(v) > 14 else 0
            
            self.conn.execute(
                """INSERT INTO voci 
                (qe_id, codice_padre, codice_completo, descrizione, tipo, 
                valore_imponibile, is_percentuale, perc_oneri, includi_oneri_in_iva, 
                perc_iva, flag_base_asta, flag_soggetto_ribasso, macro_base_calcolo, 
                flag_calcolo_montante) 
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", 
                (new_qid, v[2], v[3], v[4], v[5], v[6], v[7], v[8], v[9], 
                 v[10], v[11], v[12], v[13], f_mont)
            )
        
        self.conn.commit()
    
    def get_qe_by_progetto(self, pid):
        """Recupera tutti i QE di un progetto"""
        return self.conn.execute(
            "SELECT * FROM quadri_economici WHERE progetto_id=? ORDER BY id DESC", 
            (pid,)
        ).fetchall()
    
    def get_qe_by_id(self, qid):
        """Recupera singolo QE per ID"""
        return self.conn.execute(
            "SELECT * FROM quadri_economici WHERE id=?", (qid,)
        ).fetchone()

    # --- CRUD OPERATIONS: VOCI ---
    
    def get_voci_by_qe(self, qid):
        """Recupera tutte le voci di un QE"""
        return self.conn.execute(
            "SELECT * FROM voci WHERE qe_id=? ORDER BY codice_completo ASC", 
            (qid,)
        ).fetchall()
    
    def inserisci_voce(self, qe_id, cp, cf, desc, tipo, val, isp, po, inc, pi, 
                       f_base, f_rib, m_base, f_mont):
        """Inserisce nuova voce"""
        self.conn.execute(
            """INSERT INTO voci 
            (qe_id, codice_padre, codice_completo, descrizione, tipo, 
            valore_imponibile, is_percentuale, perc_oneri, includi_oneri_in_iva, 
            perc_iva, flag_base_asta, flag_soggetto_ribasso, macro_base_calcolo, 
            flag_calcolo_montante) 
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", 
            (qe_id, cp, cf, desc, tipo, val, isp, po, inc, pi, f_base, f_rib, m_base, f_mont)
        )
        self.conn.commit()
    
    def aggiorna_voce(self, vid, desc, val, isp, po, inc, pi, f_base, f_rib, 
                      m_base, tipo_str, f_mont):
        """Aggiorna voce esistente"""
        self.conn.execute(
            """UPDATE voci SET 
            descrizione=?, valore_imponibile=?, is_percentuale=?, perc_oneri=?, 
            includi_oneri_in_iva=?, perc_iva=?, flag_base_asta=?, 
            flag_soggetto_ribasso=?, macro_base_calcolo=?, tipo=?, 
            flag_calcolo_montante=? 
            WHERE id=?""", 
            (desc, val, isp, po, inc, pi, f_base, f_rib, m_base, tipo_str, f_mont, vid)
        )
        self.conn.commit()
    
    def elimina_voce(self, vid):
        """Elimina voce"""
        self.conn.execute("DELETE FROM voci WHERE id=?", (vid,))
        self.conn.commit()

    # --- CRUD OPERATIONS: ALLEGATI ---
    
    def inserisci_allegato(self, qe_id, nome, tipo, blob_data):
        """Inserisce nuovo allegato"""
        self.conn.execute(
            """INSERT INTO allegati_qe 
            (qe_id, nome_file, tipo_file, dati, data_caricamento) 
            VALUES (?, ?, ?, ?, ?)""", 
            (qe_id, nome, tipo, blob_data, 
             datetime.datetime.now().strftime("%d/%m/%Y %H:%M"))
        )
        self.conn.commit()
    
    def get_allegati_headers_by_qe(self, qe_id):
        """Recupera lista allegati (senza blob) per un QE"""
        return self.conn.execute(
            """SELECT id, nome_file, data_caricamento 
            FROM allegati_qe WHERE qe_id=? ORDER BY id DESC""", 
            (qe_id,)
        ).fetchall()
    
    def get_allegato_blob(self, all_id):
        """Recupera blob allegato per ID"""
        return self.conn.execute(
            "SELECT nome_file, dati FROM allegati_qe WHERE id=?", 
            (all_id,)
        ).fetchone()
    
    def elimina_allegato(self, all_id):
        """Elimina allegato"""
        self.conn.execute("DELETE FROM allegati_qe WHERE id=?", (all_id,))
        self.conn.commit()


# =============================================================================
# 2. APP GESTIONALE
# =============================================================================
class AppGestionale(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("QE Zero 2.x")
        self.geometry("1400x900")
        
        # Configurazione stili
        self.setup_styles()
        
        # Database
        self.db = DatabaseManager()
        
        # Variabili di stato
        self.init_state_variables()
        
        # Setup interfaccia
        self.setup_menu()
        self.setup_notebook()
        self.setup_all_tabs()

    def setup_styles(self):
        """Configura gli stili dell'interfaccia"""
        self.option_add('*background', '#f0f0f0')
        self.option_add('*foreground', 'black')
        self.option_add('*Entry.background', 'white')
        self.option_add('*Entry.foreground', 'black')
        self.option_add('*TCombobox.background', 'white')
        
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        self.bg_std = "#f0f0f0"
        self.bg_edit = "#fff9c4"
        self.fg_std = "black"
        
        # Stili base
        self.style.configure(".", background=self.bg_std, foreground=self.fg_std, 
                            font=("Segoe UI", 10))
        self.style.configure("TLabel", background=self.bg_std, foreground=self.fg_std)
        self.style.configure("TFrame", background=self.bg_std)
        self.style.configure("TLabelframe", background=self.bg_std, foreground=self.fg_std)
        self.style.configure("TLabelframe.Label", background=self.bg_std, 
                            foreground="#003366", font=("Segoe UI", 10, "bold"))
        self.style.configure("TButton", padding=5)
        
        # Stili speciali
        self.style.configure("Header.TLabel", font=("Segoe UI", 14, "bold"), 
                            foreground="#003366")
        self.style.configure("Totali.TLabel", font=("Segoe UI", 12, "bold"), 
                            foreground="#2e7d32")
        self.style.configure("RedInfo.TLabel", font=("Segoe UI", 9), foreground="red")
        self.style.configure("RedValue.TLabel", font=("Segoe UI", 12, "bold"), 
                            foreground="red")
        self.style.configure("Danger.TButton", foreground="red")
        self.style.configure("Discrete.TLabel", font=("Segoe UI", 9, "italic"), 
                            foreground="#555")

    def init_state_variables(self):
        """Inizializza tutte le variabili di stato dell'applicazione"""
        # ID correnti
        self.progetto_corrente_id = None
        self.progetto_base_asta = 0.0
        self.progetto_normativa_id = None
        self.qe_corrente_id = None
        self.id_modifica_proj = None
        self.id_modifica_qe = None
        self.voce_modifica_id = None
        
        # Variabili Tkinter
        self.tipo_voce_var = tk.StringVar(value="fisso")
        self.valore_tipo_var = tk.StringVar(value="fisso")
        self.check_iva_oneri_var = tk.IntVar()
        self.macro_area_var = tk.StringVar()
        self.codice_padre_var = tk.StringVar()
        self.flag_base_asta_var = tk.IntVar()
        self.flag_soggetto_ribasso_var = tk.IntVar()
        self.flag_calcolo_montante_var = tk.IntVar()
        self.normativa_var = tk.StringVar()
        self.inv_inc_var = tk.IntVar()
        
        # Valori calcolati
        self.tot_base_asta_per_calcoli = 0.0
        self.inv_res_val = 0.0

    def setup_menu(self):
        """Crea il menu dell'applicazione"""
        menubar = tk.Menu(self)
        self.config(menu=menubar)
        
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="?", menu=help_menu)
        help_menu.add_command(label="Info", command=self.show_info_window)

    def setup_notebook(self):
        """Crea il notebook con le tab principali"""
        self.nb = ttk.Notebook(self)
        self.nb.pack(expand=True, fill='both', padx=5, pady=5)
        
        # Tab 1: Progetti
        self.t1 = ttk.Frame(self.nb)
        self.nb.add(self.t1, text=' 1. Progetti ')
        
        # Tab 2: Versioni QE
        self.t2 = ttk.Frame(self.nb)
        self.nb.add(self.t2, text=' 2. Versioni QE ')
        
        # Tab 3: Voci (Editor)
        self.t3 = ttk.Frame(self.nb)
        self.nb.add(self.t3, text=' 3. Voci (Editor) ')
        
        # Tab 4: Confronto QE
        self.t6 = ttk.Frame(self.nb)
        self.nb.add(self.t6, text=' 4. Confronto QE ')
        
        # Tab 5: Amministrazione
        self.t5 = ttk.Frame(self.nb)
        self.nb.add(self.t5, text=' 5. Amministrazione ')
        
        self.nb.bind("<<NotebookTabChanged>>", self.on_tab_change)

    def setup_all_tabs(self):
        """Configura tutte le tab"""
        self.setup_tab_interventi()
        self.setup_tab_qe()
        self.setup_tab_voci()
        self.setup_tab_confronto()
        self.setup_tab_admin()

    # --- UTILITY METHODS ---
    
    def show_info_window(self):
        """Mostra finestra info applicazione"""
        messagebox.showinfo(
            "QE Zero 2.x", 
            "Gestione Quadri Economici\nVersione 2.0.0 Ottimizzata\nLogica Montante Esplicita"
        )

    def fmt(self, v):
        """Formatta un numero in valuta italiana (es: 1.234,56)"""
        try:
            if v is None:
                return "0,00"
            return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except:
            return "0,00"

    def parse(self, s):
        """Converte stringa formato italiano in float"""
        if not s:
            return 0.0
        try:
            # Rimuove €, spazi, converte separatori
            clean = str(s).replace('€', '').strip().replace('.', '').replace(',', '.')
            return float(clean)
        except:
            return 0.0

    def on_tab_change(self, event):
        """Gestisce il cambio di tab per refresh automatici"""
        idx = event.widget.index(event.widget.select())
        
        if idx == 0:  # Tab Progetti
            self.refresh_progetti()
        elif idx == 1:  # Tab QE
            self.refresh_qe()
        elif idx == 3:  # Tab Confronto
            self.refresh_confronto_combo()

    # --- TAB 1: GESTIONE PROGETTI ---
    
    def setup_tab_interventi(self):
        """Configura la tab di gestione progetti"""
        # Header
        self.lbl_h_proj = ttk.Label(
            self.t1, 
            text="GESTIONE PROGETTI", 
            style="Header.TLabel", 
            padding=10
        )
        self.lbl_h_proj.pack(fill='x')
        
        # Form nuovo progetto
        self.f_in_proj = tk.LabelFrame(
            self.t1, 
            text="Nuovo Progetto", 
            padx=15, 
            pady=15, 
            bg=self.bg_std
        )
        self.f_in_proj.pack(fill='x', padx=10, pady=5)
        
        # Configurazione colonne
        for i in range(8):
            self.f_in_proj.columnconfigure(i, weight=0)
        self.f_in_proj.columnconfigure(7, weight=1)
        
        # Campo Titolo (riga 0, span completo)
        tk.Label(self.f_in_proj, text="Titolo:", bg=self.bg_std).grid(
            row=0, column=0, sticky='w'
        )
        self.e_tit = ttk.Entry(self.f_in_proj)
        self.e_tit.grid(row=0, column=1, columnspan=7, sticky='ew', padx=5, pady=(5, 10))
        
        # Riga 1: Anno, CUP, Budget, Catalogo
        tk.Label(self.f_in_proj, text="Anno:", bg=self.bg_std).grid(
            row=1, column=0, sticky='e'
        )
        self.e_anno = ttk.Entry(self.f_in_proj, width=6)
        self.e_anno.grid(row=1, column=1, sticky='w', padx=5)
        self.e_anno.insert(0, str(datetime.date.today().year))
        
        tk.Label(self.f_in_proj, text="CUP:", bg=self.bg_std).grid(
            row=1, column=2, sticky='e'
        )
        self.e_cup = ttk.Entry(self.f_in_proj, width=20)
        self.e_cup.grid(row=1, column=3, sticky='w', padx=5)
        
        tk.Label(self.f_in_proj, text="Budget (€):", bg=self.bg_std, fg="green").grid(
            row=1, column=4, sticky='e'
        )
        self.e_imp = ttk.Entry(self.f_in_proj, width=15)
        self.e_imp.grid(row=1, column=5, sticky='w', padx=5)
        
        tk.Label(self.f_in_proj, text="Catalogo:", bg=self.bg_std).grid(
            row=1, column=6, sticky='e'
        )
        self.cb_norm = ttk.Combobox(
            self.f_in_proj, 
            textvariable=self.normativa_var, 
            state="readonly"
        )
        self.cb_norm.grid(row=1, column=7, sticky='ew', padx=5)
        
        # Container principale: lista + pulsanti
        f_main = ttk.Frame(self.t1)
        f_main.pack(fill='both', expand=True, padx=10)
        
        # Lista progetti
        c_list = ttk.LabelFrame(f_main, text="Archivio", padding=10)
        c_list.pack(side='left', fill='both', expand=True)
        
        self.tr_p = ttk.Treeview(
            c_list, 
            columns=("ID", "Norm", "CUP", "Anno", "Tit", "Imp"), 
            show='headings', 
            selectmode='browse'
        )
        
        # Configurazione colonne
        columns_config = [
            ("ID", "ID", 30),
            ("Norm", "Normativa", 180),
            ("CUP", "CUP", 100),
            ("Anno", "Anno", 50),
            ("Tit", "Titolo", 400),
            ("Imp", "Budget", 120)
        ]
        
        for col, text, width in columns_config:
            self.tr_p.heading(col, text=text)
            self.tr_p.column(col, width=width, anchor='e' if col == "Imp" else 'w')
        
        # Scrollbar
        sb = ttk.Scrollbar(c_list, orient="vertical", command=self.tr_p.yview)
        self.tr_p.configure(yscrollcommand=sb.set)
        self.tr_p.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')
        
        self.tr_p.bind("<Double-1>", self.seleziona_progetto)
        
        # Pulsanti laterali
        f_side = ttk.Frame(f_main)
        f_side.pack(side='right', fill='y', padx=10)
        
        self.btn_salva_p = ttk.Button(f_side, text="+ Crea", command=self.salva_progetto)
        self.btn_salva_p.pack(fill='x', pady=5)
        
        ttk.Button(f_side, text="Modifica", command=self.carica_modifica_progetto).pack(
            fill='x', pady=5
        )
        
        ttk.Button(
            f_side, 
            text="Elimina", 
            command=self.elimina_progetto, 
            style="Danger.TButton"
        ).pack(fill='x', pady=5)
        
        self.btn_annulla_p = ttk.Button(f_side, text="Annulla", command=self.reset_form_p)
        self.btn_annulla_p.pack(fill='x', pady=5)
        self.btn_annulla_p.pack_forget()

    # --- METODI TAB 1: PROGETTI ---
    
    def refresh_normative_combo(self):
        """Aggiorna combobox normative"""
        norms = self.db.get_normative()
        vals = [f"{n[0]} - {n[1]}" for n in norms]
        self.cb_norm['values'] = vals
        
        if vals and not self.cb_norm.get():
            self.cb_norm.current(0)
    
    def salva_progetto(self):
        """Salva o aggiorna progetto"""
        if not self.cb_norm.get():
            messagebox.showwarning("Attenzione", "Seleziona una Normativa")
            return
        
        # Estrai ID normativa
        nid = int(self.cb_norm.get().split(" - ")[0])
        
        # Valori form
        c = self.e_cup.get().strip()
        a = self.e_anno.get().strip()
        t = self.e_tit.get().strip()
        imp = self.parse(self.e_imp.get())
        
        if not t:
            messagebox.showwarning("Attenzione", "Inserisci un titolo")
            return
        
        # Modalità modifica o inserimento
        if self.id_modifica_proj:
            self.db.aggiorna_progetto_dati(self.id_modifica_proj, c, a, t, imp)
            self.reset_form_p()
        else:
            self.db.inserisci_progetto(nid, c, a, t, imp)
            self.pulisci_p()
        
        self.refresh_progetti()
    
    def refresh_progetti(self):
        """Aggiorna lista progetti"""
        self.refresh_normative_combo()
        self.tr_p.delete(*self.tr_p.get_children())
        
        for r in self.db.get_tutti_progetti():
            self.tr_p.insert(
                "", "end", 
                values=(r[0], r[5], r[1], r[2], r[3], self.fmt(r[4]))
            )
    
    def seleziona_progetto(self, e):
        """Doppio click su progetto: apre tab QE"""
        s = self.tr_p.selection()
        if not s:
            return
        
        it = self.tr_p.item(s)['values']
        self.progetto_corrente_id = it[0]
        
        # Recupera dati progetto
        p_data = self.db.get_progetto_by_id(it[0])
        self.progetto_normativa_id = p_data[1]
        
        # Aggiorna header nelle altre tab
        txt = f"Progetto: {it[4]} (CUP: {it[2]})"
        for lbl in [self.lbl_p_header_2, self.lbl_p_header_3, self.lbl_p_header_4]:
            lbl.config(text=txt)
        
        self.refresh_qe()
        self.nb.select(1)  # Passa a tab QE
    
    def carica_modifica_progetto(self):
        """Carica progetto in modalità modifica"""
        s = self.tr_p.selection()
        if not s:
            return
        
        pid = self.tr_p.item(s)['values'][0]
        v = self.db.get_progetto_by_id(pid)
        
        if not v:
            return
        
        self.id_modifica_proj = pid
        
        # Imposta normativa nel combo
        norms = self.db.get_normative()
        for n in norms:
            if n[0] == v[1]:
                self.cb_norm.set(f"{n[0]} - {n[1]}")
                break
        
        # Popola campi
        self.e_cup.delete(0, tk.END)
        self.e_cup.insert(0, v[2])
        
        self.e_anno.delete(0, tk.END)
        self.e_anno.insert(0, v[3])
        
        self.e_tit.delete(0, tk.END)
        self.e_tit.insert(0, v[4])
        
        self.e_imp.delete(0, tk.END)
        self.e_imp.insert(0, self.fmt(v[5]))
        
        # Modalità modifica visiva
        self.f_in_proj.config(bg=self.bg_edit)
        self.btn_salva_p.config(text="Salva")
        self.btn_annulla_p.pack(fill='x')
    
    def reset_form_p(self):
        """Reset form progetto"""
        self.id_modifica_proj = None
        self.pulisci_p()
        self.f_in_proj.config(bg=self.bg_std)
        self.btn_salva_p.config(text="+ Crea")
        self.btn_annulla_p.pack_forget()
    
    def pulisci_p(self):
        """Pulisce campi form progetto"""
        self.e_cup.delete(0, tk.END)
        self.e_anno.delete(0, tk.END)
        self.e_tit.delete(0, tk.END)
        self.e_imp.delete(0, tk.END)
    
    def elimina_progetto(self):
        """Elimina progetto selezionato"""
        s = self.tr_p.selection()
        if not s:
            return
        
        if messagebox.askyesno("Conferma", "Eliminare il progetto selezionato?"):
            self.db.elimina_progetto(self.tr_p.item(s)['values'][0])
            self.refresh_progetti()

    # --- TAB 2: GESTIONE VERSIONI QE ---
    
    def setup_tab_qe(self):
        """Configura la tab delle versioni QE"""
        # Info progetto corrente
        self.lbl_p_header_2 = ttk.Label(
            self.t2, 
            text="", 
            style="Discrete.TLabel", 
            padding=5
        )
        self.lbl_p_header_2.pack(fill='x')
        
        # Header
        self.lbl_hq = ttk.Label(
            self.t2, 
            text="NESSUN PROGETTO SELEZIONATO", 
            style="Header.TLabel", 
            padding=10
        )
        self.lbl_hq.pack(fill='x')
        
        # Form nuova versione QE
        self.f_qe_in = tk.LabelFrame(
            self.t2, 
            text="Versione QE", 
            padx=15, 
            pady=15, 
            bg=self.bg_std
        )
        self.f_qe_in.pack(fill='x', padx=10, pady=5)
        
        tk.Label(self.f_qe_in, text="Versione:", bg=self.bg_std).pack(side='left')
        self.e_qn = ttk.Entry(self.f_qe_in, width=30)
        self.e_qn.pack(side='left', padx=5)
        
        tk.Label(self.f_qe_in, text="Note:", bg=self.bg_std).pack(side='left', padx=(10, 0))
        self.e_qt = ttk.Entry(self.f_qe_in, width=50)
        self.e_qt.pack(side='left', padx=5)
        
        # Container principale
        f_main = ttk.Frame(self.t2)
        f_main.pack(fill='both', expand=True, padx=10)
        
        # Lista versioni QE
        c_list = ttk.LabelFrame(f_main, text="Elenco Versioni", padding=10)
        c_list.pack(side='left', fill='both', expand=True)
        
        self.tr_q = ttk.Treeview(
            c_list, 
            columns=("ID", "Nom", "Dat", "Tot", "Not"), 
            show='headings', 
            selectmode='browse'
        )
        
        # Configurazione colonne
        qe_columns = [
            ("ID", "ID", 40),
            ("Nom", "Nome Versione", 200),
            ("Dat", "Data", 100),
            ("Tot", "Totale Intervento", 120),
            ("Not", "Note", 300)
        ]
        
        for k, t, w in qe_columns:
            self.tr_q.heading(k, text=t)
            self.tr_q.column(k, width=w, anchor='e' if k == "Tot" else 'w')
        
        self.tr_q.pack(side='left', fill='both', expand=True)
        self.tr_q.bind("<<TreeviewSelect>>", self.ui_seleziona_qe)
        self.tr_q.bind("<Double-1>", self.ui_apri_qe)
        
        # Pulsanti laterali
        f_side = ttk.Frame(f_main)
        f_side.pack(side='right', fill='y', padx=10)
        
        self.btn_sq = ttk.Button(f_side, text="+ Crea", command=self.save_q)
        self.btn_sq.pack(fill='x', pady=5)
        
        ttk.Button(f_side, text="Modifica", command=self.mod_q).pack(fill='x', pady=5)
        ttk.Button(f_side, text="Duplica", command=self.dup_q).pack(fill='x', pady=5)
        ttk.Button(
            f_side, 
            text="Elimina", 
            command=self.del_q, 
            style="Danger.TButton"
        ).pack(fill='x', pady=5)
        
        ttk.Button(f_side, text="Stampa QE", command=self.genera_report_html).pack(
            fill='x', pady=5
        )
        ttk.Button(f_side, text="Esporta Excel/CSV", command=self.esporta_qe_csv).pack(
            fill='x', pady=5
        )
        
        self.btn_aq = ttk.Button(f_side, text="Annulla", command=self.rst_q)
        self.btn_aq.pack(fill='x', pady=5)
        self.btn_aq.pack_forget()

    def ui_seleziona_qe(self, e):
        """Selezione QE dalla lista"""
        s = self.tr_q.selection()
        if not s:
            return
        
        it = self.tr_q.item(s)['values']
        self.qe_corrente_id = it[0]
        self.lbl_hq.config(text=f"QE Selezionato: {it[1]}")
    
    def ui_apri_qe(self, e):
        """Doppio click: apre editor voci"""
        s = self.tr_q.selection()
        if not s:
            return
        
        it = self.tr_q.item(s)['values']
        self.qe_corrente_id = it[0]
        self.lbl_hq.config(text=f"QE Selezionato: {it[1]}")
        self.lbl_editor_title.config(text=f"{it[1]}")
        
        self.refresh_v()
        self.nb.select(2)  # Passa a tab Voci
    
    def save_q(self):
        """Salva o aggiorna QE"""
        if not self.progetto_corrente_id:
            return
        
        if self.id_modifica_qe:
            self.db.aggiorna_qe(self.id_modifica_qe, self.e_qn.get(), self.e_qt.get())
            self.rst_q()
        else:
            self.db.inserisci_qe(self.progetto_corrente_id, self.e_qn.get(), self.e_qt.get())
            self.e_qn.delete(0, tk.END)
            self.e_qt.delete(0, tk.END)
        
        self.refresh_qe()
    
    def mod_q(self):
        """Carica QE in modalità modifica"""
        s = self.tr_q.selection()
        if not s:
            return
        
        v = self.tr_q.item(s)['values']
        self.id_modifica_qe = v[0]
        
        self.e_qn.delete(0, tk.END)
        self.e_qn.insert(0, v[1])
        
        self.e_qt.delete(0, tk.END)
        self.e_qt.insert(0, v[4])
        
        self.f_qe_in.config(bg=self.bg_edit)
        self.btn_sq.config(text="Salva")
        self.btn_aq.pack(fill='x', pady=5)
    
    def rst_q(self):
        """Reset form QE"""
        self.id_modifica_qe = None
        self.e_qn.delete(0, tk.END)
        self.e_qt.delete(0, tk.END)
        self.f_qe_in.config(bg=self.bg_std)
        self.btn_sq.config(text="+ Crea")
        self.btn_aq.pack_forget()
    
    def refresh_qe(self):
        """Aggiorna lista QE con calcolo totali"""
        self.tr_q.delete(*self.tr_q.get_children())
        
        if not self.progetto_corrente_id:
            return
        
        for q in self.db.get_qe_by_progetto(self.progetto_corrente_id):
            voci = self.db.get_voci_by_qe(q[0])
            
            # Calcolo montante (solo voci con flag=1 e importo fisso)
            montante = 0.0
            for r in voci:
                if r[7] == 0:  # Solo importi fissi
                    f_mont = r[14] if len(r) > 14 else 0
                    if f_mont == 1:
                        montante += r[6]
            
            # Calcolo totale QE
            tot_qe = 0.0
            for r in voci:
                # Imponibile (fisso o percentuale su montante)
                imp = r[6] if r[7] == 0 else (montante * r[6] / 100)
                
                # Oneri
                one = imp * r[8] / 100
                
                # IVA (su imp+oneri se flag=1, altrimenti solo su imp)
                base_iva = (imp + one) if r[9] else imp
                iva = base_iva * r[10] / 100
                
                tot_qe += (imp + one + iva)
            
            self.tr_q.insert(
                "", "end", 
                values=(q[0], q[2], q[3], self.fmt(tot_qe), q[4])
            )
    
    def dup_q(self):
        """Duplica QE selezionato"""
        s = self.tr_q.selection()
        if not s:
            return
        
        qid = self.tr_q.item(s)['values'][0]
        nome_orig = self.tr_q.item(s)['values'][1]
        
        self.db.duplica_qe(qid, f"Copia {nome_orig}")
        self.refresh_qe()
    
    def del_q(self):
        """Elimina QE selezionato"""
        s = self.tr_q.selection()
        if not s:
            return
        
        if messagebox.askyesno("Conferma", "Eliminare questa versione QE?"):
            self.db.elimina_qe(self.tr_q.item(s)['values'][0])
            self.refresh_qe()

    # --- TAB 3: EDITOR VOCI (PARTE 1: LAYOUT UI) ---
    
    def setup_tab_voci(self):
        """Configura la tab dell'editor voci"""
        # Info progetto
        self.lbl_p_header_3 = ttk.Label(
            self.t3, 
            text="", 
            style="Discrete.TLabel", 
            padding=5
        )
        self.lbl_p_header_3.pack(fill='x')
        
        # Header con pulsante allegati
        f_top = ttk.Frame(self.t3)
        f_top.pack(fill='x')
        
        self.lbl_editor_title = ttk.Label(
            f_top, 
            text="Editor Voce", 
            style="Header.TLabel", 
            padding=5
        )
        self.lbl_editor_title.pack(side='left', fill='x', expand=True)
        
        ttk.Button(
            f_top, 
            text="Gestione Allegati", 
            command=self.apri_gestione_allegati
        ).pack(side='right', padx=10, pady=5)
        
        # PanedWindow: treeview | form editor
        paned = tk.PanedWindow(self.t3, orient=tk.HORIZONTAL, sashwidth=5)
        paned.pack(fill='both', expand=True, padx=2, pady=2)
        
        # --- LATO SINISTRO: TREEVIEW VOCI ---
        f_tree = ttk.Frame(paned)
        paned.add(f_tree, minsize=750)
        
        self.tr_v = ttk.Treeview(
            f_tree, 
            columns=("Cod", "Desc", "Imp", "One", "IVA", "Tot", "Note"), 
            show='headings', 
            selectmode='browse'
        )
        
        # Tag per stili
        self.tr_v.tag_configure('group', background='black', foreground='white', 
                                font=('Segoe UI', 10, 'bold'))
        self.tr_v.tag_configure('e18', background='#e6f7ff', 
                                font=('Arial', 10, 'bold'))
        self.tr_v.tag_configure('category', background='#d9d9d9', 
                                font=('Segoe UI', 10, 'bold'))
        
        # Configurazione colonne
        voci_cols = [
            ("Cod", "Cod", 60),
            ("Desc", "Descrizione", 280),
            ("Imp", "Imponibile", 80),
            ("One", "Oneri", 80),
            ("IVA", "IVA", 80),
            ("Tot", "Totale", 90),
            ("Note", "Note", 60)
        ]
        
        for c, name, w in voci_cols:
            self.tr_v.heading(c, text=name)
            self.tr_v.column(c, width=w, anchor='w' if c == "Desc" else 'e')
        
        self.tr_v.pack(fill='both', expand=True)
        self.tr_v.bind("<<TreeviewSelect>>", self.carica_edit_v)
        
        # --- LATO DESTRO: FORM EDITOR ---
        f_edit = ttk.Frame(paned)
        paned.add(f_edit, minsize=400)
        
        # Riepilogo totali (in basso)
        lf_riepilogo = ttk.LabelFrame(f_edit, text="Riepilogo Quadro Economico", padding=5)
        lf_riepilogo.pack(side='bottom', fill='x', padx=5, pady=5)
        lf_riepilogo.columnconfigure(0, weight=1)
        lf_riepilogo.columnconfigure(1, weight=0)
        
        ttk.Label(
            lf_riepilogo, 
            text="TOTALE INTERVENTO (1+2):", 
            font=("Segoe UI", 11), 
            foreground="blue"
        ).grid(row=0, column=0, sticky='w')
        
        self.lbl_val_tot = ttk.Label(
            lf_riepilogo, 
            text="€ 0,00", 
            font=("Segoe UI", 11), 
            foreground="blue"
        )
        self.lbl_val_tot.grid(row=0, column=1, sticky='e')
        
        ttk.Label(
            lf_riepilogo, 
            text="Importo Stanziato:", 
            font=("Segoe UI", 11)
        ).grid(row=1, column=0, sticky='w')
        
        self.lbl_val_stanz = ttk.Label(
            lf_riepilogo, 
            text="€ 0,00", 
            font=("Segoe UI", 11)
        )
        self.lbl_val_stanz.grid(row=1, column=1, sticky='e')
        
        # Label Economie/Fabbisogni con colori
        f_lbl_eco = tk.Frame(lf_riepilogo, bg=self.bg_std)
        f_lbl_eco.grid(row=2, column=0, sticky='w')
        
        tk.Label(
            f_lbl_eco, 
            text="Economie", 
            fg="green", 
            bg=self.bg_std, 
            font=("Segoe UI", 11)
        ).pack(side='left')
        
        tk.Label(
            f_lbl_eco, 
            text=" / ", 
            bg=self.bg_std, 
            font=("Segoe UI", 11)
        ).pack(side='left')
        
        tk.Label(
            f_lbl_eco, 
            text="Fabbisogni", 
            fg="red", 
            bg=self.bg_std, 
            font=("Segoe UI", 11)
        ).pack(side='left')
        
        self.lbl_val_eco = ttk.Label(
            lf_riepilogo, 
            text="€ 0,00", 
            font=("Segoe UI", 11)
        )
        self.lbl_val_eco.grid(row=2, column=1, sticky='e')
        
        # --- SEZIONE 1: CLASSIFICAZIONE ---
        lf_class = ttk.LabelFrame(f_edit, text="1. Classificazione", padding=5)
        lf_class.pack(fill='x', padx=5, pady=2)
        lf_class.columnconfigure(1, weight=1)
        
        ttk.Label(lf_class, text="Macro Area:").grid(row=0, column=0, sticky='w')
        self.cb_macro = ttk.Combobox(
            lf_class, 
            values=[
                "1. Spese per l'esecuzione dell'intervento", 
                "2. Somme a disposizione della S.A."
            ], 
            state="readonly", 
            textvariable=self.macro_area_var
        )
        self.cb_macro.grid(row=0, column=1, sticky='ew', padx=5, pady=2)
        self.cb_macro.bind("<<ComboboxSelected>>", self.upd_cat)
        
        ttk.Label(lf_class, text="Categoria:").grid(row=1, column=0, sticky='w')
        self.cb_cat = ttk.Combobox(
            lf_class, 
            state="readonly", 
            textvariable=self.codice_padre_var
        )
        self.cb_cat.grid(row=1, column=1, sticky='ew', padx=5, pady=2)
        self.cb_cat.bind("<<ComboboxSelected>>", self.calc_code)
        
        self.lbl_code = ttk.Label(
            lf_class, 
            text="Codice: -", 
            font=("Segoe UI", 10, "bold"), 
            foreground="blue"
        )
        self.lbl_code.grid(row=2, column=1, sticky='w', padx=5, pady=2)
        
        # --- SEZIONE 2: DESCRIZIONE ---
        lf_desc = ttk.LabelFrame(f_edit, text="2. Descrizione", padding=5)
        lf_desc.pack(fill='x', padx=5, pady=2)
        
        self.e_desc = ttk.Entry(lf_desc)
        self.e_desc.pack(fill='x')
        
        # --- SEZIONE 3: DETTAGLI ECONOMICI ---
        lf_econ = ttk.LabelFrame(f_edit, text="3. Dettagli Economici", padding=5)
        lf_econ.pack(fill='x', padx=5, pady=2)
        lf_econ.columnconfigure(0, weight=0)
        lf_econ.columnconfigure(1, weight=0)
        lf_econ.columnconfigure(2, weight=0)
        lf_econ.columnconfigure(3, weight=1)
        
        # Riga 1: Radio importo fisso/percentuale
        ttk.Radiobutton(
            lf_econ, 
            text="Importo Fisso (€)", 
            variable=self.valore_tipo_var, 
            value="fisso", 
            command=self.toggle_input_type
        ).grid(row=0, column=0, sticky='w')
        
        self.e_val = ttk.Entry(lf_econ, width=15)
        self.e_val.grid(row=0, column=1, sticky='w', padx=5)
        
        ttk.Radiobutton(
            lf_econ, 
            text="Calcolo Percentuale (%)", 
            variable=self.valore_tipo_var, 
            value="perc", 
            command=self.toggle_input_type
        ).grid(row=0, column=2, sticky='w', padx=10)
        
        self.lbl_info_perc = ttk.Label(
            lf_econ, 
            text="", 
            foreground="gray", 
            font=("Segoe UI", 8)
        )
        self.lbl_info_perc.grid(row=0, column=3, sticky='w', padx=5)
        
        # Riga 2: Oneri, IVA
        f_fiscal = ttk.Frame(lf_econ)
        f_fiscal.grid(row=1, column=0, columnspan=4, sticky='w', pady=(5, 0))
        
        ttk.Label(f_fiscal, text="Oneri (%):").pack(side='left')
        self.e_one = ttk.Entry(f_fiscal, width=8)
        self.e_one.pack(side='left', padx=5)
        
        self.chk_iva = ttk.Checkbutton(
            f_fiscal, 
            text="Includi Oneri nel calcolo IVA", 
            variable=self.check_iva_oneri_var
        )
        self.chk_iva.pack(side='left', padx=20)
        
        ttk.Label(f_fiscal, text="IVA (%):").pack(side='left')
        self.e_iva = ttk.Entry(f_fiscal, width=8)
        self.e_iva.pack(side='left', padx=5)
        
        # --- SEZIONE 4: OPZIONI ---
        lf_opt = ttk.LabelFrame(f_edit, text="4. Opzioni", padding=5)
        lf_opt.pack(fill='x', padx=5, pady=2)
        lf_opt.columnconfigure(0, weight=1)
        lf_opt.columnconfigure(1, weight=1)
        lf_opt.columnconfigure(2, weight=1)
        
        self.chk_base = ttk.Checkbutton(
            lf_opt, 
            text="Voce Base d'Asta", 
            variable=self.flag_base_asta_var
        )
        self.chk_base.grid(row=0, column=0, sticky='w')
        
        self.chk_rib = ttk.Checkbutton(
            lf_opt, 
            text="Soggetto a Ribasso", 
            variable=self.flag_soggetto_ribasso_var
        )
        self.chk_rib.grid(row=0, column=1, sticky='w')
        
        self.chk_mont = ttk.Checkbutton(
            lf_opt, 
            text="Calcolo Montante", 
            variable=self.flag_calcolo_montante_var
        )
        self.chk_mont.grid(row=0, column=2, sticky='w')
        
        # --- PULSANTI AZIONE ---
        f_btns = ttk.Frame(f_edit, padding=5)
        f_btns.pack(fill='x', pady=2)
        
        ttk.Button(f_btns, text="SALVA VOCE", command=self.save_v).pack(fill='x', pady=2)
        
        f_act = ttk.Frame(f_btns)
        f_act.pack(fill='x', pady=2)
        
        ttk.Button(f_act, text="Nuova", command=self.rst_v).pack(
            side='left', expand=True, fill='x', padx=(0, 2)
        )
        
        ttk.Button(
            f_act, 
            text="Elimina", 
            command=self.del_v, 
            style="Danger.TButton"
        ).pack(side='right', expand=True, fill='x', padx=(2, 0))
        
        # --- CALCOLO INVERSO IMPONIBILE ---
        lf_inv = ttk.LabelFrame(f_edit, text="Calcolo inverso dell'imponibile", padding=5)
        lf_inv.pack(fill='x', padx=5, pady=5)
        
        f_inv_left = ttk.Frame(lf_inv)
        f_inv_left.pack(side='left', fill='both', expand=True)
        
        f_inv_right = ttk.Frame(lf_inv)
        f_inv_right.pack(side='right', fill='both', padx=5)
        
        # Campi input calcolo inverso
        ttk.Label(f_inv_left, text="Totale (€):").grid(row=0, column=0, sticky='w')
        self.e_inv_tot = ttk.Entry(f_inv_left, width=12)
        self.e_inv_tot.grid(row=0, column=1, padx=5, pady=2)
        
        ttk.Label(f_inv_left, text="Oneri (%):").grid(row=1, column=0, sticky='w')
        self.e_inv_one = ttk.Entry(f_inv_left, width=8)
        self.e_inv_one.insert(0, "0,0")
        self.e_inv_one.grid(row=1, column=1, padx=5, pady=2)
        
        self.chk_inv_inc = ttk.Checkbutton(
            f_inv_left, 
            text="Includi Oneri nel calcolo IVA", 
            variable=self.inv_inc_var
        )
        self.chk_inv_inc.grid(row=1, column=2, sticky='w', padx=5)
        
        ttk.Label(f_inv_left, text="IVA (%):").grid(row=2, column=0, sticky='w')
        self.e_inv_iva = ttk.Entry(f_inv_left, width=8)
        self.e_inv_iva.insert(0, "22,0")
        self.e_inv_iva.grid(row=2, column=1, padx=5, pady=2)
        
        # Pulsanti e risultato
        ttk.Button(f_inv_right, text="Calcola", command=self.do_calcolo_inverso).pack(
            fill='x', pady=2
        )
        
        self.lbl_inv_res = ttk.Label(
            f_inv_right, 
            text="Imp: € 0,00", 
            font=("Segoe UI", 9, "bold"), 
            foreground="blue"
        )
        self.lbl_inv_res.pack(pady=5)
        
        ttk.Button(f_inv_right, text="Usa nel QE", command=self.usa_risultato_inverso).pack(
            fill='x', pady=2
        )

    def do_calcolo_inverso(self):
        """Calcola imponibile partendo dal totale"""
        try:
            T = self.parse(self.e_inv_tot.get())
            O_perc = self.parse(self.e_inv_one.get()) / 100.0
            V_perc = self.parse(self.e_inv_iva.get()) / 100.0
            Inc = self.inv_inc_var.get()
            
            # Formula inversa
            if Inc:
                denom = (1 + O_perc) * (1 + V_perc)
            else:
                denom = (1 + O_perc + V_perc)
            
            I = T / denom if denom != 0 else 0
            self.inv_res_val = I
            self.lbl_inv_res.config(text=f"Imp: € {self.fmt(I)}")
            
        except Exception as e:
            messagebox.showerror("Errore Calcolo", f"Controlla i valori inseriti\n{e}")

    def usa_risultato_inverso(self):
        """Trasferisce risultato calcolo inverso nel form principale"""
        if self.inv_res_val <= 0:
            messagebox.showinfo("Info", "Effettua prima il calcolo.")
            return
        
        # Imposta importo fisso
        self.valore_tipo_var.set("fisso")
        self.toggle_input_type()
        
        self.e_val.delete(0, tk.END)
        self.e_val.insert(0, self.fmt(self.inv_res_val))
        
        # Copia oneri e IVA
        self.e_one.delete(0, tk.END)
        self.e_one.insert(0, self.e_inv_one.get())
        
        self.e_iva.delete(0, tk.END)
        self.e_iva.insert(0, self.e_inv_iva.get())
        
        self.check_iva_oneri_var.set(self.inv_inc_var.get())
        
    def toggle_input_type(self):
        """Mostra info quando si seleziona calcolo percentuale"""
        t = self.valore_tipo_var.get()
        
        if t == 'perc':
            self.lbl_info_perc.config(
                text=f"Calcolato su: € {self.fmt(self.tot_base_asta_per_calcoli)}"
            )
        else:
            self.lbl_info_perc.config(text="")
        
    def upd_cat(self, e):
        """Aggiorna combobox categorie in base a macro area"""
        if not self.progetto_normativa_id:
            return
        
        m_sel = self.macro_area_var.get()
        m = int(m_sel[0]) if m_sel and m_sel[0] in ('1', '2') else 1
        
        cats = self.db.get_catalogo(self.progetto_normativa_id, m)
        self.cb_cat['values'] = [f"{i[1]} - {i[3]}" for i in cats]
        self.cb_cat.set("")
    
    def calc_code(self, e):
        """Calcola e mostra prossimo codice disponibile"""
        if not self.voce_modifica_id and self.codice_padre_var.get():
            codice_padre = self.codice_padre_var.get().split(' - ')[0]
            prossimo = self.db.get_prossimo_codice(self.qe_corrente_id, codice_padre)
            self.lbl_code.config(text=f"Cod: {prossimo}")
    
    def rst_v(self):
        """Reset form voce"""
        self.voce_modifica_id = None
        
        # Pulisci campi
        self.e_desc.delete(0, tk.END)
        self.e_val.delete(0, tk.END)
        self.e_one.delete(0, tk.END)
        self.e_iva.delete(0, tk.END)
        
        # Valori default
        self.e_one.insert(0, "0,0")
        self.e_iva.insert(0, "0,0")
        
        self.check_iva_oneri_var.set(0)
        self.calc_code(None)
        
        self.valore_tipo_var.set("fisso")
        self.toggle_input_type()
        
        # Reset checkbox
        self.flag_base_asta_var.set(0)
        self.flag_soggetto_ribasso_var.set(0)
        self.flag_calcolo_montante_var.set(0)

    def save_v(self):
        """Salva o aggiorna voce"""
        if not self.qe_corrente_id:
            return
        
        # Valori form
        v = self.parse(self.e_val.get())
        po = self.parse(self.e_one.get())
        pi = self.parse(self.e_iva.get())
        inc = self.check_iva_oneri_var.get()
        
        # Checkbox flags
        f_base = self.flag_base_asta_var.get()
        f_rib = self.flag_soggetto_ribasso_var.get()
        f_mont = self.flag_calcolo_montante_var.get()
        
        tipo_str = self.valore_tipo_var.get()
        m_str = ""
        
        if self.voce_modifica_id:
            # Aggiorna voce esistente
            self.db.aggiorna_voce(
                self.voce_modifica_id, self.e_desc.get(), v, 
                1 if tipo_str == 'perc' else 0, 
                po, inc, pi, f_base, f_rib, m_str, tipo_str, f_mont
            )
        else:
            # Nuova voce
            cp = self.codice_padre_var.get().split(" - ")[0] if self.codice_padre_var.get() else ""
            if not cp:
                messagebox.showwarning("Attenzione", "Seleziona una categoria")
                return
            
            cf = self.db.get_prossimo_codice(self.qe_corrente_id, cp)
            
            self.db.inserisci_voce(
                self.qe_corrente_id, cp, cf, self.e_desc.get(), tipo_str, v, 
                1 if tipo_str == 'perc' else 0, 
                po, inc, pi, f_base, f_rib, m_str, f_mont
            )
        
        self.refresh_v()
        self.rst_v()
    
    def del_v(self):
        """Elimina voce selezionata"""
        if self.voce_modifica_id:
            if messagebox.askyesno("Conferma", "Eliminare questa voce?"):
                self.db.elimina_voce(self.voce_modifica_id)
                self.refresh_v()
                self.rst_v()
    
    def carica_edit_v(self, e):
        """Carica voce selezionata nel form per modifica"""
        s = self.tr_v.selection()
        if not s or not s[0].isdigit():
            return
        
        r = self.db.conn.execute(
            "SELECT * FROM voci WHERE id=?", (int(s[0]),)
        ).fetchone()
        
        if not r:
            return
        
        self.voce_modifica_id = r[0]
        
        # Mostra codice
        self.lbl_code.config(text=f"Mod: {r[3]}")
        
        # Descrizione
        self.e_desc.delete(0, tk.END)
        self.e_desc.insert(0, r[4])
        
        # Valore
        self.e_val.delete(0, tk.END)
        self.e_val.insert(0, self.fmt(r[6]))
        
        # Tipo (fisso/perc)
        self.valore_tipo_var.set('perc' if r[7] else 'fisso')
        
        # Oneri
        self.e_one.delete(0, tk.END)
        self.e_one.insert(0, self.fmt(r[8]))
        
        # Flag includi oneri in IVA
        self.check_iva_oneri_var.set(r[9])
        
        # IVA
        self.e_iva.delete(0, tk.END)
        self.e_iva.insert(0, self.fmt(r[10]))
        
        # Checkbox flags
        self.flag_base_asta_var.set(r[11])
        self.flag_soggetto_ribasso_var.set(r[12])
        
        # flag_calcolo_montante (index 14)
        f_mont = r[14] if len(r) > 14 else 0
        self.flag_calcolo_montante_var.set(f_mont)
        
        self.toggle_input_type()
        
        # Imposta categoria nel combo
        cod_padre = r[2]
        cat_row = self.db.conn.execute(
            """SELECT macro_gruppo, descrizione 
            FROM catalogo_voci 
            WHERE normativa_id=? AND codice=?""", 
            (self.progetto_normativa_id, cod_padre)
        ).fetchone()
        
        if cat_row:
            macro_id = cat_row[0]
            desc_cat = cat_row[1]
            
            if macro_id == 1:
                macro_str = "1. Spese per l'esecuzione dell'intervento"
            elif macro_id == 2:
                macro_str = "2. Somme a disposizione della S.A."
            else:
                macro_str = ""
            
            self.cb_macro.set(macro_str)
            self.upd_cat(None)
            self.cb_cat.set(f"{cod_padre} - {desc_cat}")

    def refresh_v(self):
        """Aggiorna treeview voci con calcolo totali e raggruppamenti"""
        self.tr_v.delete(*self.tr_v.get_children())
        
        if not self.qe_corrente_id:
            return
        
        voci = self.db.get_voci_by_qe(self.qe_corrente_id)
        
        # CALCOLO MONTANTE: somma imponibili con flag_calcolo_montante=1
        self.tot_base_asta_per_calcoli = 0.0
        for r in voci:
            if r[7] == 0:  # Solo importi fissi
                f_mont = r[14] if len(r) > 14 else 0
                if f_mont == 1:
                    self.tot_base_asta_per_calcoli += r[6]
        
        # Aggiorna label info percentuale se attivo
        if self.valore_tipo_var.get() == 'perc':
            self.toggle_input_type()
        
        # Totali globali
        tot_base_asta_imp = 0.0
        tot_oneri_globali = 0.0
        tot_iva_globali = 0.0
        tot_tasse_globali = 0.0
        
        # Separa voci in sezioni
        sec1 = []  # Base asta
        sec2 = []  # Somme a disposizione
        tot_sec2_imp = 0.0
        
        for r in voci:
            # Calcolo importi
            imp = r[6] if r[7] == 0 else (self.tot_base_asta_per_calcoli * r[6] / 100)
            one = imp * r[8] / 100
            base_iva = (imp + one) if r[9] else imp
            iva = base_iva * r[10] / 100
            tot = imp + one + iva
            
            tot_oneri_globali += one
            tot_iva_globali += iva
            tot_tasse_globali += (one + iva)
            
            item = {'r': r, 'imp': imp, 'one': one, 'iva': iva, 'tot': tot}
            
            if r[11] == 1:  # flag_base_asta
                sec1.append(item)
                tot_base_asta_imp += imp
            else:
                sec2.append(item)
                tot_sec2_imp += imp
        
        # Mappa categorie
        cat_map = {}
        cats = self.db.get_catalogo(self.progetto_normativa_id)
        for c in cats:
            cat_map[c[1]] = c[3]
        
        def render_section(items_list, title, is_sec1):
            """Renderizza una sezione con raggruppamenti per categoria"""
            if not items_list and is_sec1:
                return
            
            group_tot = tot_base_asta_imp if is_sec1 else (tot_sec2_imp + tot_tasse_globali)
            
            # Header sezione
            self.tr_v.insert(
                "", "end", 
                text=title, 
                values=("", title, self.fmt(group_tot), "", "", "", ""), 
                tags=('group',)
            )
            
            # Ordina per codice padre
            items_list.sort(key=lambda x: x['r'][2])
            
            # Raggruppa per categoria
            for key, group in groupby(items_list, key=lambda x: x['r'][2]):
                g_list = list(group)
                
                # Totali categoria
                s_imp = sum(x['imp'] for x in g_list)
                s_one = sum(x['one'] for x in g_list)
                s_iva = sum(x['iva'] for x in g_list)
                s_tot = sum(x['tot'] for x in g_list)
                
                cat_desc = cat_map.get(key, f"Categoria {key}")
                
                # Riga categoria
                self.tr_v.insert(
                    "", "end", 
                    values=(
                        key, cat_desc, 
                        self.fmt(s_imp), self.fmt(s_one), 
                        self.fmt(s_iva), self.fmt(s_tot), ""
                    ), 
                    tags=('category',)
                )
                
                # Voci della categoria
                for i in g_list:
                    r = i['r']
                    
                    # Flag info (Ribasso, Montante)
                    f_mont = r[14] if len(r) > 14 else 0
                    info_tags = []
                    if r[12]:  # flag_soggetto_ribasso
                        info_tags.append("Rib")
                    if f_mont:
                        info_tags.append("Mont")
                    note_str = " ".join(info_tags)
                    
                    self.tr_v.insert(
                        "", "end", 
                        iid=str(r[0]), 
                        values=(
                            "  " + r[3], r[4], 
                            self.fmt(i['imp']), self.fmt(i['one']), 
                            self.fmt(i['iva']), self.fmt(i['tot']), 
                            note_str
                        )
                    )
            
            # Riga IVA e imposte (solo sezione 2)
            if not is_sec1:
                self.tr_v.insert(
                    "", "end", 
                    iid="tax", 
                    values=(
                        "", "IVA e altre imposte (Totale)", 
                        self.fmt(tot_tasse_globali), 
                        self.fmt(tot_oneri_globali), 
                        self.fmt(tot_iva_globali), "", ""
                    ), 
                    tags=('e18',)
                )
        
        # Renderizza sezioni
        render_section(sec1, "1. SPESE PER L'ESECUZIONE DELL'INTERVENTO", True)
        render_section(sec2, "2. SOMME A DISPOSIZIONE", False)
        
        # Aggiorna totali in UI
        tot_gen = tot_base_asta_imp + tot_sec2_imp + tot_tasse_globali
        self.lbl_val_tot.config(text=f"€ {self.fmt(tot_gen)}")
        
        # Economie/Fabbisogni
        proj = self.db.get_progetto_by_id(self.progetto_corrente_id)
        budget = proj[5] if proj else 0.0
        diff = budget - tot_gen
        
        self.lbl_val_stanz.config(text=f"€ {self.fmt(budget)}")
        self.lbl_val_eco.config(
            text=f"€ {self.fmt(diff)}", 
            foreground="green" if diff >= 0 else "red"
        )

    def apri_gestione_allegati(self):
        """Finestra gestione allegati PDF con descrizione"""
        if not self.qe_corrente_id:
            return
        
        d = tk.Toplevel(self)
        d.title("Gestione Allegati QE")
        d.geometry("800x500")
        
        # Frame principale con layout orizzontale
        f_main = ttk.Frame(d)
        f_main.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Frame sinistro: Treeview
        f_left = ttk.Frame(f_main)
        f_left.pack(side='left', fill='both', expand=True)
        
        # Treeview allegati (aggiunta colonna Descrizione)
        tr = ttk.Treeview(f_left, columns=("ID", "Nome", "Data", "Desc"), show='headings')
        tr.heading("ID", text="ID")
        tr.column("ID", width=40)
        tr.heading("Nome", text="Nome File")
        tr.column("Nome", width=250)
        tr.heading("Data", text="Data")
        tr.column("Data", width=120)
        tr.heading("Desc", text="Descrizione")
        tr.column("Desc", width=200)
        
        # Scrollbar
        sb = ttk.Scrollbar(f_left, orient="vertical", command=tr.yview)
        tr.configure(yscrollcommand=sb.set)
        tr.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')
        
        # Frame destro: Pulsanti impilati
        f_right = ttk.Frame(f_main)
        f_right.pack(side='right', fill='y', padx=(10, 0))
        
        def refresh():
            """Aggiorna lista allegati"""
            tr.delete(*tr.get_children())
            # Query aggiornata per includere descrizione
            rows = self.db.conn.execute(
                "SELECT id, nome_file, data_caricamento, descrizione FROM allegati_qe WHERE qe_id=? ORDER BY id DESC",
                (self.qe_corrente_id,)
            ).fetchall()
            for r in rows:
                desc = r[3] if len(r) > 3 and r[3] else ""
                tr.insert("", "end", values=(r[0], r[1], r[2], desc))
        
        def carica():
            """Carica nuovo PDF con descrizione"""
            fp = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
            if not fp:
                return
            
            # Dialog per descrizione
            desc = simpledialog.askstring(
                "Descrizione Allegato",
                "Inserisci una descrizione (opzionale):",
                parent=d
            )
            
            try:
                with open(fp, 'rb') as f:
                    blob = f.read()
                
                # Inserimento con descrizione
                self.db.conn.execute(
                    """INSERT INTO allegati_qe 
                    (qe_id, nome_file, tipo_file, dati, data_caricamento, descrizione) 
                    VALUES (?, ?, ?, ?, ?, ?)""",
                    (
                        self.qe_corrente_id,
                        os.path.basename(fp),
                        "pdf",
                        blob,
                        datetime.datetime.now().strftime("%d/%m/%Y %H:%M"),
                        desc if desc else ""
                    )
                )
                self.db.conn.commit()
                
                refresh()
                messagebox.showinfo("OK", "File caricato con successo.", parent=d)
                
            except Exception as e:
                messagebox.showerror("Errore", f"Errore caricamento file:\n{e}", parent=d)
        
        def modifica_descrizione():
            """Modifica descrizione allegato selezionato"""
            s = tr.selection()
            if not s:
                messagebox.showwarning("Attenzione", "Seleziona un allegato.", parent=d)
                return
            
            vals = tr.item(s)['values']
            aid = vals[0]
            desc_attuale = vals[3] if len(vals) > 3 else ""
            
            # Dialog per nuova descrizione
            nuova_desc = simpledialog.askstring(
                "Modifica Descrizione",
                "Inserisci nuova descrizione:",
                initialvalue=desc_attuale,
                parent=d
            )
            
            if nuova_desc is not None:  # Anche stringa vuota è valida
                try:
                    self.db.conn.execute(
                        "UPDATE allegati_qe SET descrizione=? WHERE id=?",
                        (nuova_desc, aid)
                    )
                    self.db.conn.commit()
                    refresh()
                    messagebox.showinfo("OK", "Descrizione aggiornata.", parent=d)
                except Exception as e:
                    messagebox.showerror("Errore", f"Errore aggiornamento:\n{e}", parent=d)
        
        def scarica():
            """Scarica/Apri allegato"""
            s = tr.selection()
            if not s:
                messagebox.showwarning("Attenzione", "Seleziona un allegato.", parent=d)
                return
            
            aid = tr.item(s)['values'][0]
            r = self.db.get_allegato_blob(aid)
            
            if r:
                fn = filedialog.asksaveasfilename(
                    initialfile=r[0],
                    filetypes=[("PDF Files", "*.pdf")],
                    parent=d
                )
                if fn:
                    try:
                        with open(fn, 'wb') as f:
                            f.write(r[1])
                        
                        # Apri file con applicazione predefinita
                        if platform.system() == 'Darwin':  # macOS
                            subprocess.call(('open', fn))
                        elif platform.system() == 'Windows':
                            os.startfile(fn)
                        else:  # Linux
                            subprocess.call(('xdg-open', fn))
                            
                    except Exception as e:
                        messagebox.showerror("Errore", f"Errore salvataggio:\n{e}", parent=d)
        
        def elimina():
            """Elimina allegato"""
            s = tr.selection()
            if not s:
                messagebox.showwarning("Attenzione", "Seleziona un allegato.", parent=d)
                return
            
            if messagebox.askyesno("Conferma", "Rimuovere questo file?", parent=d):
                self.db.elimina_allegato(tr.item(s)['values'][0])
                refresh()
        
        # Pulsanti impilati verticalmente a destra
        ttk.Button(
            f_right,
            text="➕ Carica PDF",
            command=carica,
            width=20
        ).pack(fill='x', pady=(0, 5))
        
        ttk.Button(
            f_right,
            text="✏️ Modifica Descrizione",
            command=modifica_descrizione,
            width=20
        ).pack(fill='x', pady=5)
        
        ttk.Button(
            f_right,
            text="⬇️ Scarica/Apri",
            command=scarica,
            width=20
        ).pack(fill='x', pady=5)
        
        ttk.Button(
            f_right,
            text="🗑️ Elimina",
            command=elimina,
            style="Danger.TButton",
            width=20
        ).pack(fill='x', pady=5)
        
        ttk.Button(
            f_right,
            text="✖ Chiudi",
            command=d.destroy,
            width=20
        ).pack(fill='x', pady=(20, 0))
        
        # Carica dati iniziali
        refresh()

    def genera_report_html(self):
        """Genera e apre report HTML del QE"""
        if not self.qe_corrente_id:
            return
        
        qe = self.db.get_qe_by_id(self.qe_corrente_id)
        proj = self.db.get_progetto_by_id(qe[1])
        voci = self.db.get_voci_by_qe(self.qe_corrente_id)
        
        # Configurazione ente
        ente_nome = self.db.get_config("ente_nome")
        ente_dettagli = (
            f"{self.db.get_config('ente_indirizzo')} - "
            f"{self.db.get_config('ente_citta')}<br>"
            f"Tel: {self.db.get_config('ente_tel')}"
        )
        
        # Calcolo montante
        montante = 0.0
        for r in voci:
            if r[7] == 0:
                f_mont = r[14] if len(r) > 14 else 0
                if f_mont == 1:
                    montante += r[6]
        
        # Separa voci
        l1 = []
        l2 = []
        t_oneri = 0.0
        t_iva = 0.0
        t_tasse = 0.0
        
        for r in voci:
            imp = r[6] if r[7] == 0 else (montante * r[6] / 100)
            one = imp * r[8] / 100
            base_iva = (imp + one) if r[9] else imp
            iva = base_iva * r[10] / 100
            tot = imp + one + iva
            
            t_oneri += one
            t_iva += iva
            t_tasse += (one + iva)
            
            item = {
                'r': r, 'code': r[3], 'desc': r[4], 
                'imp': imp, 'one': one, 'iva': iva, 'tot': tot
            }
            
            if r[11] == 1:
                l1.append(item)
            else:
                l2.append(item)
        
        # Mappa categorie
        cat_map = {}
        cats = self.db.get_catalogo(proj[1])
        for c in cats:
            cat_map[c[1]] = c[3]
        
        def build_table_rows(items):
            """Costruisce righe HTML per una sezione"""
            h = ""
            t_s = 0.0
            
            items.sort(key=lambda x: x['r'][2])
            
            for key, group in groupby(items, key=lambda x: x['r'][2]):
                g_list = list(group)
                
                s_imp = sum(x['imp'] for x in g_list)
                s_one = sum(x['one'] for x in g_list)
                s_iva = sum(x['iva'] for x in g_list)
                s_tot = sum(x['tot'] for x in g_list)
                t_s += s_tot
                
                cat_desc = cat_map.get(key, f"Categoria {key}")
                
                # Riga categoria
                h += (
                    f"<tr class='cat-row'><td>{key}</td><td>{cat_desc}</td>"
                    f"<td align='right'>{self.fmt(s_imp)}</td>"
                    f"<td align='right'>{self.fmt(s_one)}</td>"
                    f"<td align='right'>{self.fmt(s_iva)}</td>"
                    f"<td align='right'>{self.fmt(s_tot)}</td></tr>"
                )
                
                # Righe voci
                for i in g_list:
                    h += (
                        f"<tr><td style='padding-left:20px;'>{i['code']}</td>"
                        f"<td>{i['desc']}</td>"
                        f"<td align='right'>{self.fmt(i['imp'])}</td>"
                        f"<td align='right'>{self.fmt(i['one'])}</td>"
                        f"<td align='right'>{self.fmt(i['iva'])}</td>"
                        f"<td align='right'>{self.fmt(i['tot'])}</td></tr>"
                    )
            
            return h, t_s
        
        r1, tot1 = build_table_rows(l1)
        r2, tot2 = build_table_rows(l2)
        
        # Riga IVA
        r2 += (
            f"<tr style='background-color:#e6f7ff; font-weight:bold;'>"
            f"<td></td><td>Riepilogo IVA e Imposte</td>"
            f"<td align='right'>{self.fmt(t_tasse)}</td>"
            f"<td align='right'>{self.fmt(t_oneri)}</td>"
            f"<td align='right'>{self.fmt(t_iva)}</td>"
            f"<td align='right'></td></tr>"
        )
        
        # Calcolo totali
        t1_imp = sum([x['imp'] for x in l1])
        t2_imp = sum([x['imp'] for x in l2])
        tot2_full = t2_imp + t_tasse
        tot_qe = t1_imp + tot2_full
        
        imp_stanziato = proj[5]
        economie = imp_stanziato - tot_qe
        col_eco = "green" if economie >= 0 else "red"
        
        # Header HTML
        header_html = (
            f"<div class='h-ente'><h2>{ente_nome}</h2>"
            f"<p>{ente_dettagli}</p><hr>"
            f"<h3>Progetto: {proj[4]} (CUP: {proj[2]})</h3>"
            f"<p><b>QE:</b> {qe[2]}<br><b>Note:</b> {qe[4]}</p></div>"
        )
        
        table_start = (
            "<table style='width:100%; table-layout: fixed; border-collapse: collapse;'>"
            "<thead><tr>"
            "<th width='8%'>Cod</th><th width='32%'>Descrizione</th>"
            "<th width='15%' align='right'>Imponibile</th>"
            "<th width='15%' align='right'>Oneri</th>"
            "<th width='15%' align='right'>IVA</th>"
            "<th width='15%' align='right'>Totale</th>"
            "</tr></thead><tbody>"
        )
        
        # HTML completo
        html = f"""<html>
<head>
<style>
body {{ font-family: Arial; padding: 30px; }}
table {{ margin-bottom: 20px; font-size: 12px; width:100%; border-collapse:collapse; }}
td, th {{ border: 1px solid #ccc; padding: 5px; }}
th {{ background: #ddd; }}
.cat-row {{ background-color: #d9d9d9; font-weight: bold; }}
.tot-row {{ background: #ccc; font-weight: bold; }}
.sec-title {{ background-color: #000; color: #fff; padding: 5px; font-weight: bold; margin-top: 20px; }}
</style>
</head>
<body>
{header_html}
<div class="sec-title">1. SPESE PER L'ESECUZIONE DELL'INTERVENTO</div>
{table_start}{r1}
<tr class="tot-row">
<td colspan="2" align="right">Totale (1):</td>
<td align="right">{self.fmt(t1_imp)}</td>
<td></td><td></td><td></td>
</tr></tbody></table>

<div class="sec-title">2. SOMME A DISPOSIZIONE</div>
{table_start}{r2}
<tr class="tot-row">
<td colspan="2" align="right">Totale (2):</td>
<td align="right">{self.fmt(tot2_full)}</td>
<td></td><td></td><td></td>
</tr></tbody></table>

<br>
<table style="width:100%; border: 2px solid #000;">
<tr>
<td width="70%" align="right"><b>TOTALE INTERVENTO (1+2):</b></td>
<td width="30%" align="right"><b>{self.fmt(tot_qe)} €</b></td>
</tr>
<tr>
<td align="right">Importo Stanziato:</td>
<td align="right">{self.fmt(imp_stanziato)} €</td>
</tr>
<tr>
<td align="right"><b>Economie / (Fabbisogni):</b></td>
<td align="right" style="color:{col_eco}"><b>{self.fmt(economie)} €</b></td>
</tr>
</table>
</body>
</html>"""
        
        try:
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            fn = os.path.join(
                self.db.stampe_path, 
                f"Stampa_QE_{self.qe_corrente_id}_{ts}.html"
            )
            
            with open(fn, "w", encoding="utf-8") as f:
                f.write(html)
            
            url = 'file://' + urllib.request.pathname2url(os.path.abspath(fn))
            webbrowser.open(url)
            
        except Exception as e:
            messagebox.showerror(
                "Errore Stampa", 
                f"Impossibile creare il file di stampa:\n{e}"
            )

    def esporta_qe_csv(self):
        """Esporta QE in formato CSV (Excel)"""
        if not self.qe_corrente_id:
            return
        
        fn = filedialog.asksaveasfilename(
            defaultextension=".csv", 
            filetypes=[("CSV (Excel)", "*.csv")]
        )
        
        if not fn:
            return
        
        try:
            voci = self.db.get_voci_by_qe(self.qe_corrente_id)
            proj = self.db.get_progetto_by_id(self.progetto_corrente_id)
            
            # Calcolo montante
            montante = 0.0
            for r in voci:
                if r[7] == 0:
                    f_mont = r[14] if len(r) > 14 else 0
                    if f_mont == 1:
                        montante += r[6]
            
            # Mappa categorie
            cat_map = {}
            cats = self.db.get_catalogo(proj[1])
            for c in cats:
                cat_map[c[1]] = c[3]
            
            # Separa voci
            l1 = []
            l2 = []
            t_oneri = 0.0
            t_iva = 0.0
            t_tasse = 0.0
            
            for r in voci:
                imp = r[6] if r[7] == 0 else (montante * r[6] / 100)
                one = imp * r[8] / 100
                base_iva = (imp + one) if r[9] else imp
                iva = base_iva * r[10] / 100
                tot = imp + one + iva
                
                t_oneri += one
                t_iva += iva
                t_tasse += (one + iva)
                
                item = {
                    'r': r, 'code': r[3], 'desc': r[4], 
                    'imp': imp, 'one': one, 'iva': iva, 'tot': tot
                }
                
                if r[11] == 1:
                    l1.append(item)
                else:
                    l2.append(item)
            
            t1_imp = sum([x['imp'] for x in l1])
            t2_imp = sum([x['imp'] for x in l2])
            
            # Scrivi CSV
            with open(fn, 'w', newline='', encoding='utf-8-sig') as f:
                w = csv.writer(f, delimiter=';')
                w.writerow(["Codice", "Descrizione", "Imponibile", "Oneri", "IVA", "Totale", "Note"])
                
                def write_group_section(items_list):
                    items_list.sort(key=lambda x: x['r'][2])
                    
                    for key, group in groupby(items_list, key=lambda x: x['r'][2]):
                        g_list = list(group)
                        
                        s_imp = sum(x['imp'] for x in g_list)
                        s_one = sum(x['one'] for x in g_list)
                        s_iva = sum(x['iva'] for x in g_list)
                        s_tot = sum(x['tot'] for x in g_list)
                        
                        cat_desc = cat_map.get(key, f"Categoria {key}")
                        
                        w.writerow([
                            key, cat_desc.upper(), 
                            self.fmt(s_imp), self.fmt(s_one), 
                            self.fmt(s_iva), self.fmt(s_tot), 
                            "Riepilogo Categoria"
                        ])
                        
                        for i in g_list:
                            w.writerow([
                                i['code'], i['desc'], 
                                self.fmt(i['imp']), self.fmt(i['one']), 
                                self.fmt(i['iva']), self.fmt(i['tot']), ""
                            ])
                
                w.writerow(["", "1. SPESE PER L'ESECUZIONE DELL'INTERVENTO", "", "", "", "", ""])
                write_group_section(l1)
                w.writerow(["", "Totale (1)", self.fmt(t1_imp), "", "", "", ""])
                
                w.writerow([])
                w.writerow(["", "2. SOMME A DISPOSIZIONE", "", "", "", "", ""])
                write_group_section(l2)
                w.writerow([
                    "", "Riepilogo IVA e Imposte", 
                    self.fmt(t_tasse), self.fmt(t_oneri), 
                    self.fmt(t_iva), "", ""
                ])
                w.writerow(["", "Totale (2)", self.fmt(t2_imp + t_tasse), "", "", "", ""])
                
                w.writerow([])
                tot_qe = t1_imp + t2_imp + t_tasse
                w.writerow(["", "TOTALE COMPLESSIVO", self.fmt(tot_qe), "", "", "", ""])
            
            messagebox.showinfo("Export", "Esportazione completata con successo!")
            
        except Exception as e:
            messagebox.showerror("Errore", f"Errore durante l'esportazione:\n{e}")

    # --- TAB 4: CONFRONTO QE ---
    
    def setup_tab_confronto(self):
        """Configura la tab di confronto tra versioni QE"""
        self.lbl_p_header_4 = ttk.Label(
            self.t6, 
            text="", 
            style="Discrete.TLabel", 
            padding=5
        )
        self.lbl_p_header_4.pack(fill='x')
        
        lbl = ttk.Label(
            self.t6, 
            text="Confronto tra Versioni", 
            style="Header.TLabel", 
            padding=15
        )
        lbl.pack()
        
        # Form selezione QE
        f_sel = ttk.Frame(self.t6)
        f_sel.pack(pady=10)
        
        ttk.Label(f_sel, text="QE A (Rif.):").pack(side='left')
        self.cb_qe1 = ttk.Combobox(f_sel, state="readonly", width=30)
        self.cb_qe1.pack(side='left', padx=5)
        
        ttk.Label(f_sel, text="QE B (Var.):").pack(side='left')
        self.cb_qe2 = ttk.Combobox(f_sel, state="readonly", width=30)
        self.cb_qe2.pack(side='left', padx=5)
        
        ttk.Button(f_sel, text="CONFRONTA", command=self.effettua_confronto).pack(
            side='left', padx=10
        )
        ttk.Button(f_sel, text="Stampa Confronto", command=self.stampa_confronto).pack(
            side='left', padx=10
        )
        ttk.Button(f_sel, text="Esporta Excel/CSV", command=self.esporta_confronto_csv).pack(
            side='left', padx=10
        )
        
        # Treeview risultati
        c_ids = ("Cod", "Desc", "A", "B", "Diff", "Perc")
        self.tr_diff = ttk.Treeview(
            self.t6, 
            columns=c_ids, 
            show='headings', 
            height=20
        )
        
        self.tr_diff.tag_configure('up', foreground='green')
        self.tr_diff.tag_configure('down', foreground='red')
        
        cols_config = [
            ("Cod", "Codice", 60),
            ("Desc", "Descrizione", 300),
            ("A", "Importo A", 100),
            ("B", "Importo B", 100),
            ("Diff", "Differenza", 100),
            ("Perc", "Var %", 80)
        ]
        
        for c, h, w in cols_config:
            self.tr_diff.heading(c, text=h)
            self.tr_diff.column(c, width=w, anchor='w' if c == "Desc" else 'e')
        
        self.tr_diff.pack(expand=True, fill='both', padx=20, pady=10)
        
        self.lbl_diff_tot = ttk.Label(
            self.t6, 
            text="Delta Totale: 0", 
            font=("Arial", 12, "bold")
        )
        self.lbl_diff_tot.pack(pady=10)
    
    def refresh_confronto_combo(self):
        """Aggiorna combobox confronto"""
        if not self.progetto_corrente_id:
            return
        
        qes = self.db.get_qe_by_progetto(self.progetto_corrente_id)
        vals = [f"{q[0]} - {q[2]}" for q in qes]
        
        self.cb_qe1['values'] = vals
        self.cb_qe2['values'] = vals
    
    def effettua_confronto(self):
        """Confronta due versioni QE"""
        s1, s2 = self.cb_qe1.get(), self.cb_qe2.get()
        
        if not s1 or not s2:
            messagebox.showwarning("Attenzione", "Seleziona entrambi i QE da confrontare")
            return
        
        id1, id2 = int(s1.split(' - ')[0]), int(s2.split(' - ')[0])
        v1 = self.db.get_voci_by_qe(id1)
        v2 = self.db.get_voci_by_qe(id2)
        
        def get_calc_imponibili(vl):
            """Calcola imponibili per lista voci"""
            # Calcolo montante
            m = 0.0
            for r in vl:
                if r[7] == 0:
                    f_mont = r[14] if len(r) > 14 else 0
                    if f_mont == 1:
                        m += r[6]
            
            res = {}
            tx = 0.0
            
            for r in vl:
                i = r[6] if r[7] == 0 else (m * r[6] / 100)
                res[r[3]] = {'desc': r[4], 'imp': i}
                
                o = i * r[8] / 100
                base_iva = (i + o) if r[9] else i
                iv = base_iva * r[10] / 100
                tx += (o + iv)
            
            return res, tx
        
        d1, t1 = get_calc_imponibili(v1)
        d2, t2 = get_calc_imponibili(v2)
        
        codes = sorted(list(set(d1.keys()) | set(d2.keys())))
        
        self.tr_diff.delete(*self.tr_diff.get_children())
        
        sa = 0.0
        sb = 0.0
        grand_delta = 0.0
        
        for c in codes:
            i1 = d1.get(c, {'desc': '', 'imp': 0.0})['imp']
            desc = d2.get(c, {'desc': d1.get(c, {'desc': ''})['desc']})['desc']
            i2 = d2.get(c, {'desc': '', 'imp': 0.0})['imp']
            d = i2 - i1
            
            sa += i1
            sb += i2
            grand_delta += d
            
            perc = ((i2 - i1) / i1 * 100) if i1 != 0 else (0.0 if i2 == 0 else 100.0)
            
            tag = 'up' if d > 0.01 else ('down' if d < -0.01 else '')
            
            self.tr_diff.insert(
                "", "end", 
                values=(
                    c, desc, 
                    self.fmt(i1), self.fmt(i2), 
                    self.fmt(d), f"{perc:+.2f}%"
                ), 
                tags=(tag,)
            )
        
        # Riga IVA
        dt = t2 - t1
        grand_delta += dt
        
        perc_t = ((t2 - t1) / t1 * 100) if t1 != 0 else 0.0
        tagt = 'up' if dt > 0.01 else ('down' if dt < -0.01 else '')
        
        self.tr_diff.insert(
            "", "end", 
            values=(
                "", "IVA e altre imposte", 
                self.fmt(t1), self.fmt(t2), 
                self.fmt(dt), f"{perc_t:+.2f}%"
            ), 
            tags=(tagt,)
        )
        
        col_tot = "green" if grand_delta >= 0 else "red"
        self.lbl_diff_tot.config(
            text=f"Variazione Totale: {self.fmt(grand_delta)} €", 
            foreground=col_tot
        )
    
    def stampa_confronto(self):
        """Genera report HTML di confronto"""
        items = self.tr_diff.get_children()
        if not items:
            messagebox.showwarning("Attenzione", "Effettua prima il confronto.")
            return
        
        id1 = int(self.cb_qe1.get().split(' - ')[0])
        id2 = int(self.cb_qe2.get().split(' - ')[0])
        
        v1 = self.db.get_voci_by_qe(id1)
        v2 = self.db.get_voci_by_qe(id2)
        
        def get_data_calc(voci):
            montante = 0.0
            for r in voci:
                if r[7] == 0:
                    f_mont = r[14] if len(r) > 14 else 0
                    if f_mont == 1:
                        montante += r[6]
            
            data = {}
            taxes = 0.0
            
            for r in voci:
                imp = r[6] if r[7] == 0 else (montante * r[6] / 100)
                one = imp * r[8] / 100
                base_iva = (imp + one) if r[9] else imp
                iva = base_iva * r[10] / 100
                taxes += (one + iva)
                data[r[3]] = {'desc': r[4], 'imp': imp, 'flag': r[11]}
            
            return data, taxes
        
        d1, t1 = get_data_calc(v1)
        d2, t2 = get_data_calc(v2)
        
        all_codes = sorted(list(set(d1.keys()) | set(d2.keys())))
        
        sec1 = []
        sec2 = []
        
        for c in all_codes:
            ref = d2.get(c, d1.get(c))
            flag = ref['flag'] if ref else d1.get(c, {}).get('flag', 0)
            
            if flag == 1:
                sec1.append(c)
            else:
                sec2.append(c)
        
        def build_html_rows(cod_list):
            h = ""
            sum_a = 0.0
            sum_b = 0.0
            
            for c in cod_list:
                i1 = d1.get(c, {'imp': 0.0})['imp']
                i2 = d2.get(c, {'imp': 0.0})['imp']
                diff = i2 - i1
                sum_a += i1
                sum_b += i2
                
                desc = d2.get(c, {'desc': d1.get(c, {'desc': ''})['desc']})['desc']
                perc = ((i2 - i1) / i1 * 100) if i1 != 0 else (0.0 if i2 == 0 else 100.0)
                col = "green" if diff > 0.01 else ("red" if diff < -0.01 else "black")
                
                h += (
                    f"<tr><td>{c}</td><td>{desc}</td>"
                    f"<td style='text-align: right;'>{self.fmt(i1)}</td>"
                    f"<td style='text-align: right;'>{self.fmt(i2)}</td>"
                    f"<td style='text-align: right; color:{col};'>{self.fmt(diff)}</td>"
                    f"<td style='text-align: right; color:{col}; font-weight:bold;'>{perc:+.2f}%</td></tr>"
                )
            
            return h, sum_a, sum_b
        
        h1, a1, b1 = build_html_rows(sec1)
        h2, a2, b2 = build_html_rows(sec2)
        
        d1_tot = b1 - a1
        p1_tot = ((b1 - a1) / a1 * 100) if a1 != 0 else 0.0
        c1 = "green" if d1_tot > 0.01 else ("red" if d1_tot < -0.01 else "black")
        
        h1 += (
            f"<tr class='tot-row'><td colspan='2'>Totale (1)</td>"
            f"<td style='text-align: right;'>{self.fmt(a1)}</td>"
            f"<td style='text-align: right;'>{self.fmt(b1)}</td>"
            f"<td style='text-align: right; color:{c1}'>{self.fmt(d1_tot)}</td>"
            f"<td style='text-align: right; color:{c1}'>{p1_tot:+.2f}%</td></tr>"
        )
        
        d_t = t2 - t1
        col_t = "green" if d_t > 0.01 else ("red" if d_t < -0.01 else "black")
        pt = ((t2 - t1) / t1 * 100) if t1 != 0 else 0.0
        
        h2 += (
            f"<tr style='background-color:#e6f7ff; font-weight:bold;'>"
            f"<td></td><td>IVA e altre imposte (Totale)</td>"
            f"<td style='text-align: right;'>{self.fmt(t1)}</td>"
            f"<td style='text-align: right;'>{self.fmt(t2)}</td>"
            f"<td style='text-align: right; color:{col_t}'>{self.fmt(d_t)}</td>"
            f"<td style='text-align: right; color:{col_t}'>{pt:+.2f}%</td></tr>"
        )
        
        tot2_a = a2 + t1
        tot2_b = b2 + t2
        diff2 = tot2_b - tot2_a
        pt2 = ((tot2_b - tot2_a) / tot2_a * 100) if tot2_a != 0 else 0.0
        c2 = "green" if diff2 > 0.01 else ("red" if diff2 < -0.01 else "black")
        
        h2 += (
            f"<tr class='tot-row'><td colspan='2'>Totale (2)</td>"
            f"<td style='text-align: right;'>{self.fmt(tot2_a)}</td>"
            f"<td style='text-align: right;'>{self.fmt(tot2_b)}</td>"
            f"<td style='text-align: right; color:{c2}'>{self.fmt(diff2)}</td>"
            f"<td style='text-align: right; color:{c2}'>{pt2:+.2f}%</td></tr>"
        )
        
        tot1 = a1 + tot2_a
        tot2 = b1 + tot2_b
        d_tot = tot2 - tot1
        col_g = "green" if d_tot > 0.01 else ("red" if d_tot < -0.01 else "black")
        ptot = ((tot2 - tot1) / tot1 * 100) if tot1 != 0 else 0.0
        
        ente = self.db.get_config("ente_nome")
        dett = f"{self.db.get_config('ente_indirizzo')} - {self.db.get_config('ente_citta')}"
        proj = self.db.get_progetto_by_id(self.progetto_corrente_id)
        qe_a = self.cb_qe1.get().split(' - ')[1]
        qe_b = self.cb_qe2.get().split(' - ')[1]
        
        html = f"""<html>
<head>
<style>
body {{ font-family: 'Segoe UI', Arial, sans-serif; padding: 40px; }}
table {{ width: 100%; border-collapse: collapse; margin-top: 0; font-size: 12px; }}
th, td {{ border: 1px solid #ddd; padding: 6px; text-align: left; }}
th {{ background-color: #ddd; color: #000; font-weight: bold; }}
h1, h2, h3 {{ color: #003366; text-align: center; }}
.meta {{ text-align: center; color: #555; margin-bottom: 30px; }}
.sec-title {{ background-color: #000; color: #fff; padding: 5px; font-weight: bold; margin-top: 20px; }}
.tot-row {{ background-color: #ccc; font-weight: bold; font-size: 14px; }}
</style>
</head>
<body>
<h1>{ente}</h1>
<p class="meta">{dett}</p>
<hr>
<h2>CONFRONTO QUADRI ECONOMICI</h2>
<p class="meta">
<b>Progetto:</b> {proj[4]} (CUP: {proj[2]})<br>
Confronto: {qe_a} (A) vs {qe_b} (B)
</p>

<div class="sec-title">1. SPESE PER L'ESECUZIONE DELL'INTERVENTO</div>
<table style="width:100%; table-layout: fixed; border-collapse: collapse;">
<thead>
<tr>
<th width="8%">Cod</th><th width="28%">Desc</th>
<th width="16%" style="text-align: right;">Imp. A</th>
<th width="16%" style="text-align: right;">Imp. B</th>
<th width="16%" style="text-align: right;">Diff</th>
<th width="16%" style="text-align: right;">Var %</th>
</tr>
</thead>
<tbody>{h1}</tbody>
</table>

<div class="sec-title">2. SOMME A DISPOSIZIONE</div>
<table style="width:100%; table-layout: fixed; border-collapse: collapse;">
<thead>
<tr>
<th width="8%">Cod</th><th width="28%">Desc</th>
<th width="16%" style="text-align: right;">Imp. A</th>
<th width="16%" style="text-align: right;">Imp. B</th>
<th width="16%" style="text-align: right;">Diff</th>
<th width="16%" style="text-align: right;">Var %</th>
</tr>
</thead>
<tbody>{h2}</tbody>
</table>

<br>
<table style="width:100%; table-layout: fixed; border-collapse: collapse; border: 2px solid #000;">
<tr class="tot-row">
<th width="36%" style="text-align: right; padding: 6px;">TOTALE COMPLESSIVO (1+2):</th>
<th width="16%" style="text-align: right; padding: 6px;">{self.fmt(tot1)}</th>
<th width="16%" style="text-align: right; padding: 6px;">{self.fmt(tot2)}</th>
<th width="16%" style="text-align: right; color:{col_g}; padding: 6px;">{self.fmt(d_tot)}</th>
<th width="16%" style="text-align: right; color:{col_g}; padding: 6px;">{ptot:+.2f}%</th>
</tr>
</table>

<p style="font-size:10px; color:gray; margin-top:30px;">
Generato il {datetime.datetime.now().strftime("%d/%m/%Y")}
</p>
</body>
</html>"""
        
        try:
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            fn = os.path.join(
                self.db.stampe_path, 
                f"Report_Confronto_{ts}.html")
            
            with open(fn, "w", encoding="utf-8") as f:
                f.write(html)
            
            url = 'file://' + urllib.request.pathname2url(os.path.abspath(fn))
            webbrowser.open(url)
            
        except Exception as e:
            messagebox.showerror("Errore Stampa", f"Impossibile creare il report:\n{e}")

    def esporta_confronto_csv(self):
        """Esporta confronto in CSV"""
        if not self.cb_qe1.get() or not self.cb_qe2.get():
            messagebox.showwarning("Attenzione", "Effettua prima il confronto")
            return
        
        fn = filedialog.asksaveasfilename(
            defaultextension=".csv", 
            filetypes=[("CSV (Excel)", "*.csv")]
        )
        
        if not fn:
            return
        
        id1 = int(self.cb_qe1.get().split(' - ')[0])
        id2 = int(self.cb_qe2.get().split(' - ')[0])
        
        v1 = self.db.get_voci_by_qe(id1)
        v2 = self.db.get_voci_by_qe(id2)
        
        def gd(vl):
            m = 0.0
            d = {}
            tx = 0.0
            
            for r in vl:
                if r[7] == 0:
                    f_mont = r[14] if len(r) > 14 else 0
                    if f_mont == 1:
                        m += r[6]
            
            for r in vl:
                i = r[6] if r[7] == 0 else (m * r[6] / 100)
                o = i * r[8] / 100
                base_iva = (i + o) if r[9] else i
                iv = base_iva * r[10] / 100
                tx += (o + iv)
                d[r[3]] = {'desc': r[4], 'imp': i, 'flag': r[11]}
            
            return d, tx
        
        d1, t1 = gd(v1)
        d2, t2 = gd(v2)
        
        codes = sorted(list(set(d1.keys()) | set(d2.keys())))
        
        try:
            with open(fn, 'w', newline='', encoding='utf-8-sig') as f:
                w = csv.writer(f, delimiter=';')
                w.writerow(["Cod", "Desc", "Imp A", "Imp B", "Diff", "Var %"])
                
                s1 = []
                s2 = []
                
                for c in codes:
                    f = d2.get(c, d1.get(c))['flag']
                    if f == 1:
                        s1.append(c)
                    else:
                        s2.append(c)
                
                def ws(cl):
                    sa = 0.0
                    sb = 0.0
                    
                    for c in cl:
                        i1 = d1.get(c, {'imp': 0.0})['imp']
                        i2 = d2.get(c, {'imp': 0.0})['imp']
                        d = i2 - i1
                        perc = ((i2 - i1) / i1 * 100) if i1 != 0 else (0.0 if i2 == 0 else 100.0)
                        sa += i1
                        sb += i2
                        de = d2.get(c, {'desc': d1.get(c, {'desc': ''})['desc']})['desc']
                        w.writerow([c, de, self.fmt(i1), self.fmt(i2), self.fmt(d), f"{perc:+.2f}%"])
                    
                    return sa, sb
                
                w.writerow(["1. BASE ASTA", "", "", "", "", ""])
                a1, b1 = ws(s1)
                d1_tot = b1 - a1
                p1_tot = ((b1 - a1) / a1 * 100) if a1 != 0 else 0.0
                w.writerow(["Totale 1", "", self.fmt(a1), self.fmt(b1), self.fmt(d1_tot), f"{p1_tot:+.2f}%"])
                
                w.writerow([])
                w.writerow(["2. SOMME DISP", "", "", "", "", ""])
                a2, b2 = ws(s2)
                
                dt = t2 - t1
                pt = ((t2 - t1) / t1 * 100) if t1 != 0 else 0.0
                w.writerow(["", "IVA Tot", self.fmt(t1), self.fmt(t2), self.fmt(dt), f"{pt:+.2f}%"])
                
                t2a = a2 + t1
                t2b = b2 + t2
                dt2 = t2b - t2a
                pt2 = ((t2b - t2a) / t2a * 100) if t2a != 0 else 0.0
                w.writerow(["Totale 2", "", self.fmt(t2a), self.fmt(t2b), self.fmt(dt2), f"{pt2:+.2f}%"])
                
                w.writerow([])
                ta = a1 + t2a
                tb = b1 + t2b
                dgen = tb - ta
                pgen = ((tb - ta) / ta * 100) if ta != 0 else 0.0
                w.writerow(["TOTALE", "", self.fmt(ta), self.fmt(tb), self.fmt(dgen), f"{pgen:+.2f}%"])
            
            messagebox.showinfo("Export", "Esportazione completata con successo!")
            
        except Exception as e:
            messagebox.showerror("Errore", f"Errore durante l'esportazione:\n{e}")

    # --- TAB 5: AMMINISTRAZIONE ---
    
    def setup_tab_admin(self):
        """Configura la tab amministrazione"""
        # Form login
        self.f_log = ttk.Frame(self.t5)
        self.f_log.place(relx=0.5, rely=0.5, anchor='center')
        
        ttk.Label(self.f_log, text="Password Admin:").pack(pady=5)
        self.e_pwd = ttk.Entry(self.f_log, show="*")
        self.e_pwd.pack(pady=5)
        ttk.Button(self.f_log, text="Accedi", command=self.adm_log).pack(pady=10)
        
        # Frame admin (nascosto inizialmente)
        self.f_adm = ttk.Frame(self.t5)
        
        # Container top: 3 sezioni
        f_top_container = ttk.Frame(self.f_adm)
        f_top_container.pack(fill='x', padx=10, pady=5)
        
        # 1. Configurazione SA
        lf_config = ttk.LabelFrame(
            f_top_container, 
            text="1. Configurazione Stazione Appaltante", 
            padding=10
        )
        lf_config.pack(side='left', fill='both', expand=True, padx=(0, 5))
        
        self.entries_cfg = {}
        config_fields = [
            ("ente_nome", "Nome Ente"),
            ("ente_indirizzo", "Indirizzo"),
            ("ente_citta", "Città"),
            ("ente_tel", "Telefono"),
            ("ente_email", "Email"),
            ("ente_pec", "PEC")
        ]
        
        for i, (k, lbl) in enumerate(config_fields):
            r, c = divmod(i, 2)
            ttk.Label(lf_config, text=lbl).grid(row=r, column=c*2, sticky='e', padx=5, pady=2)
            e = ttk.Entry(lf_config, width=25)
            e.grid(row=r, column=c*2+1, sticky='w', padx=5, pady=2)
            self.entries_cfg[k] = e
        
        ttk.Button(
            lf_config, 
            text="Salva Configurazione", 
            command=self.save_config
        ).grid(row=3, column=0, columnspan=4, pady=10)
        
        # 2. Backup
        lf_backup = ttk.LabelFrame(
            f_top_container, 
            text="2. Backup e Ripristino", 
            padding=10
        )
        lf_backup.pack(side='left', fill='both', padx=5)
        
        ttk.Button(lf_backup, text="Backup Database", command=self.backup_db).pack(
            fill='x', pady=5
        )
        ttk.Button(lf_backup, text="Importa da Backup", command=self.importa_backup_dialog).pack(
            fill='x', pady=5
        )
        
        # 3. Sicurezza
        lf_security = ttk.LabelFrame(f_top_container, text="3. Sicurezza", padding=10)
        lf_security.pack(side='left', fill='both', padx=(5, 0))
        
        ttk.Label(lf_security, text="Nuova Pwd:").pack(anchor='w')
        self.e_new_pwd = ttk.Entry(lf_security, show="*", width=15)
        self.e_new_pwd.pack(fill='x', pady=2)
        ttk.Button(lf_security, text="Aggiorna", command=self.update_admin_pwd).pack(
            fill='x', pady=5
        )
        
        # PanedWindow: Normative | Catalogo
        paned = tk.PanedWindow(self.f_adm, orient=tk.HORIZONTAL, bg="#ccc")
        paned.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Normative
        f_norm = ttk.LabelFrame(paned, text="Gestione Cataloghi / Normative", padding=5)
        paned.add(f_norm, width=400)
        
        self.tr_norm = ttk.Treeview(f_norm, columns=("ID", "Nome"), show='headings')
        self.tr_norm.heading("ID", text="ID")
        self.tr_norm.column("ID", width=30)
        self.tr_norm.heading("Nome", text="Descrizione Normativa")
        self.tr_norm.pack(side='left', fill='both', expand=True)
        self.tr_norm.bind("<<TreeviewSelect>>", self.sel_norm_admin)
        
        fn_btns = ttk.Frame(f_norm)
        fn_btns.pack(side='right', fill='y')
        
        ttk.Button(fn_btns, text="Nuova", command=self.new_norm).pack(fill='x', pady=2)
        ttk.Button(fn_btns, text="Modifica", command=self.edit_norm).pack(fill='x', pady=2)
        ttk.Button(fn_btns, text="Duplica", command=self.dup_norm).pack(fill='x', pady=2)
        ttk.Button(
            fn_btns, 
            text="Elimina", 
            command=self.del_norm, 
            style="Danger.TButton"
        ).pack(fill='x', pady=2)
        
        # Catalogo
        f_cat = ttk.LabelFrame(paned, text="Voci del Catalogo Selezionato", padding=5)
        paned.add(f_cat)
        
        self.lbl_sel_norm = ttk.Label(
            f_cat, 
            text="Seleziona una normativa a sinistra", 
            foreground="blue"
        )
        self.lbl_sel_norm.pack(pady=5)
        
        self.tr_cat = ttk.Treeview(
            f_cat, 
            columns=("ID", "Cod", "Macro", "Desc"), 
            show='headings'
        )
        self.tr_cat.heading("ID", text="ID")
        self.tr_cat.column("ID", width=30)
        self.tr_cat.heading("Cod", text="Codice")
        self.tr_cat.column("Cod", width=60)
        self.tr_cat.heading("Macro", text="M")
        self.tr_cat.column("Macro", width=30)
        self.tr_cat.heading("Desc", text="Descrizione Completa")
        self.tr_cat.column("Desc", width=400)
        self.tr_cat.pack(side='left', fill='both', expand=True)
        self.tr_cat.bind("<Double-1>", self.edit_cat_item_dialog)
        
        fc_btns = ttk.Frame(f_cat)
        fc_btns.pack(side='right', fill='y')
        
        ttk.Button(fc_btns, text="Aggiungi Voce", command=self.new_cat_item_dialog).pack(
            fill='x', pady=2
        )
        ttk.Button(
            fc_btns, 
            text="Elimina Voce", 
            command=self.del_cat_item, 
            style="Danger.TButton"
        ).pack(fill='x', pady=2)

    def backup_db(self):
        """Crea backup database"""
        try:
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            src = self.db.db_path
            dst = os.path.join(self.db.documents_path, f"qezero_BACKUP_{ts}.db")
            
            shutil.copy2(src, dst)
            messagebox.showinfo("Backup", f"Backup creato con successo:\n{os.path.basename(dst)}")
            
        except Exception as e:
            messagebox.showerror("Errore Backup", f"Errore durante il backup:\n{e}")

    def importa_backup_dialog(self):
        """Dialog importazione progetti da backup"""
        file_path = filedialog.askopenfilename(
            title="Seleziona file di Backup (.db)",
            filetypes=[("Database SQLite", "*.db"), ("Tutti i file", "*.*")],
            initialdir=self.db.documents_path
        )
        
        if not file_path:
            return
        
        try:
            conn_backup = sqlite3.connect(file_path)
            rows = conn_backup.execute(
                "SELECT id, titolo, cup, importo, normativa_id FROM progetti ORDER BY id DESC"
            ).fetchall()
            
            # Finestra selezione progetti
            imp_win = tk.Toplevel(self)
            imp_win.title("Importa Progetti da Backup")
            imp_win.geometry("800x500")
            
            lbl_info = ttk.Label(
                imp_win,
                text=f"Backup: {os.path.basename(file_path)}\nSeleziona i progetti da importare:",
                padding=10
            )
            lbl_info.pack(fill='x')
            
            cols = ("ID_OLD", "Titolo", "CUP", "Importo", "Norm_ID_OLD")
            tree = ttk.Treeview(imp_win, columns=cols, show='headings', selectmode='extended')
            
            tree.heading("ID_OLD", text="ID Originale")
            tree.column("ID_OLD", width=50)
            tree.heading("Titolo", text="Titolo Progetto")
            tree.column("Titolo", width=400)
            tree.heading("CUP", text="CUP")
            tree.column("CUP", width=100)
            tree.heading("Importo", text="Budget")
            tree.column("Importo", width=100, anchor='e')
            tree.heading("Norm_ID_OLD", text="NID")
            tree.column("Norm_ID_OLD", width=0, stretch=False)
            
            for r in rows:
                tree.insert("", "end", values=(r[0], r[1], r[2], self.fmt(r[3]), r[4]))
            
            tree.pack(fill='both', expand=True, padx=10, pady=10)
            
            def esegui_import():
                selected_items = tree.selection()
                if not selected_items:
                    messagebox.showwarning("Attenzione", "Nessun progetto selezionato.")
                    return
                
                count_ok = 0
                
                for item in selected_items:
                    vals = tree.item(item)['values']
                    old_pid = vals[0]
                    old_nid = vals[4]
                    
                    try:
                        # Recupera dati progetto
                        p_row = conn_backup.execute(
                            "SELECT cup, anno, titolo, importo FROM progetti WHERE id=?",
                            (old_pid,)
                        ).fetchone()
                        
                        # Recupera normativa
                        norm_data = conn_backup.execute(
                            "SELECT nome, descrizione FROM normative WHERE id=?",
                            (old_nid,)
                        ).fetchone()
                        
                        cat_rows = []
                        if norm_data:
                            cat_rows = conn_backup.execute(
                                "SELECT codice, macro_gruppo, descrizione FROM catalogo_voci WHERE normativa_id=?",
                                (old_nid,)
                            ).fetchall()
                        
                        # Recupera QE e voci
                        qes_to_import = []
                        qe_rows = conn_backup.execute(
                            "SELECT id, nome_versione, data_creazione, note FROM quadri_economici WHERE progetto_id=?",
                            (old_pid,)
                        ).fetchall()
                        
                        for qe_row in qe_rows:
                            old_qid = qe_row[0]
                            
                            # COMPATIBILITÀ: gestione flag_calcolo_montante mancante
                            try:
                                v_rows = conn_backup.execute(
                                    """SELECT codice_padre, codice_completo, descrizione, tipo,
                                    valore_imponibile, is_percentuale, perc_oneri, includi_oneri_in_iva,
                                    perc_iva, flag_base_asta, flag_soggetto_ribasso, macro_base_calcolo,
                                    flag_calcolo_montante FROM voci WHERE qe_id=?""",
                                    (old_qid,)
                                ).fetchall()
                            except sqlite3.OperationalError:
                                # Fallback per vecchi DB senza flag_calcolo_montante
                                tmp_v = conn_backup.execute(
                                    """SELECT codice_padre, codice_completo, descrizione, tipo,
                                    valore_imponibile, is_percentuale, perc_oneri, includi_oneri_in_iva,
                                    perc_iva, flag_base_asta, flag_soggetto_ribasso, macro_base_calcolo
                                    FROM voci WHERE qe_id=?""",
                                    (old_qid,)
                                ).fetchall()
                                v_rows = [list(r) + [0] for r in tmp_v]
                            
                            a_rows = conn_backup.execute(
                                "SELECT nome_file, tipo_file, dati, data_caricamento FROM allegati_qe WHERE qe_id=?",
                                (old_qid,)
                            ).fetchall()
                            
                            qes_to_import.append({
                                'meta': qe_row,
                                'voci': v_rows,
                                'allegati': a_rows
                            })
                        
                        # Inserimento nel DB corrente
                        new_nid = 1
                        if norm_data:
                            norm_name, norm_desc = norm_data
                            curr_norm = self.db.conn.execute(
                                "SELECT id FROM normative WHERE nome=?",
                                (norm_name,)
                            ).fetchone()
                            
                            if curr_norm:
                                new_nid = curr_norm[0]
                            else:
                                self.db.conn.execute(
                                    "INSERT INTO normative (nome, descrizione) VALUES (?, ?)",
                                    (norm_name, norm_desc)
                                )
                                new_nid = self.db.conn.execute(
                                    "SELECT last_insert_rowid()"
                                ).fetchone()[0]
                                
                                for c_row in cat_rows:
                                    self.db.conn.execute(
                                        """INSERT INTO catalogo_voci
                                        (normativa_id, codice, macro_gruppo, descrizione)
                                        VALUES (?, ?, ?, ?)""",
                                        (new_nid, c_row[0], c_row[1], c_row[2])
                                    )
                        
                        # Inserisci progetto
                        self.db.conn.execute(
                            """INSERT INTO progetti
                            (normativa_id, cup, anno, titolo, importo)
                            VALUES (?, ?, ?, ?, ?)""",
                            (new_nid, p_row[0], p_row[1], p_row[2], p_row[3])
                        )
                        new_pid = self.db.conn.execute("SELECT last_insert_rowid()").fetchone()[0]
                        
                        # Inserisci QE e voci
                        for qe_obj in qes_to_import:
                            qm = qe_obj['meta']
                            self.db.conn.execute(
                                """INSERT INTO quadri_economici
                                (progetto_id, nome_versione, data_creazione, note)
                                VALUES (?, ?, ?, ?)""",
                                (new_pid, qm[1], qm[2], qm[3])
                            )
                            new_qid = self.db.conn.execute("SELECT last_insert_rowid()").fetchone()[0]
                            
                            for v in qe_obj['voci']:
                                self.db.conn.execute(
                                    """INSERT INTO voci
                                    (qe_id, codice_padre, codice_completo, descrizione, tipo,
                                    valore_imponibile, is_percentuale, perc_oneri, includi_oneri_in_iva,
                                    perc_iva, flag_base_asta, flag_soggetto_ribasso, macro_base_calcolo,
                                    flag_calcolo_montante)
                                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                                    (new_qid, v[0], v[1], v[2], v[3], v[4], v[5], v[6],
                                     v[7], v[8], v[9], v[10], v[11], v[12])
                                )
                            
                            for a in qe_obj['allegati']:
                                self.db.conn.execute(
                                    """INSERT INTO allegati_qe
                                    (qe_id, nome_file, tipo_file, dati, data_caricamento)
                                    VALUES (?, ?, ?, ?, ?)""",
                                    (new_qid, a[0], a[1], a[2], a[3])
                                )
                        
                        count_ok += 1
                        
                    except Exception as e:
                        messagebox.showerror(
                            "Errore Import",
                            f"Errore durante l'importazione del progetto ID {old_pid}:\n{e}"
                        )
                
                self.db.conn.commit()
                conn_backup.close()
                messagebox.showinfo("Fatto", f"Importati correttamente {count_ok} progetti.")
                imp_win.destroy()
                self.refresh_progetti()
            
            ttk.Button(imp_win, text="IMPORTA SELEZIONATI", command=esegui_import).pack(pady=20)
            
        except Exception as e:
            messagebox.showerror("Errore Apertura Backup", str(e))

    def adm_log(self):
        """Login amministrazione"""
        pwd_db = self.db.get_config("admin_password")
        if not pwd_db:
            pwd_db = "admin"
        
        if self.e_pwd.get() == pwd_db:
            self.f_log.place_forget()
            self.f_adm.pack(fill='both', expand=True)
            self.load_cfg()
            self.refresh_norm_list()
        else:
            messagebox.showerror("Errore", "Password errata")
    
    def update_admin_pwd(self):
        """Aggiorna password admin"""
        new_p = self.e_new_pwd.get().strip()
        if not new_p:
            return
        
        self.db.set_config("admin_password", new_p)
        messagebox.showinfo("Successo", "Password aggiornata!")
        self.e_new_pwd.delete(0, tk.END)
    
    def load_cfg(self):
        """Carica configurazione"""
        for k, e in self.entries_cfg.items():
            e.delete(0, tk.END)
            e.insert(0, self.db.get_config(k))
    
    def save_config(self):
        """Salva configurazione"""
        for k, e in self.entries_cfg.items():
            self.db.set_config(k, e.get())
        messagebox.showinfo("OK", "Configurazione Salvata!")
    
    def refresh_norm_list(self):
        """Aggiorna lista normative"""
        self.tr_norm.delete(*self.tr_norm.get_children())
        for n in self.db.get_normative():
            self.tr_norm.insert("", "end", values=(n[0], n[1]))
    
    def sel_norm_admin(self, e):
        """Selezione normativa per gestione catalogo"""
        s = self.tr_norm.selection()
        if not s:
            return
        
        nid = self.tr_norm.item(s)['values'][0]
        nname = self.tr_norm.item(s)['values'][1]
        
        self.active_admin_norm_id = nid
        self.lbl_sel_norm.config(text=f"Catalogo: {nname} (ID: {nid})")
        self.refresh_cat_admin()
    
    def refresh_cat_admin(self):
        """Aggiorna catalogo voci"""
        if not hasattr(self, 'active_admin_norm_id'):
            return
        
        self.tr_cat.delete(*self.tr_cat.get_children())
        for r in self.db.get_catalogo(self.active_admin_norm_id):
            self.tr_cat.insert("", "end", values=r)
    
    def new_norm(self):
        """Nuova normativa"""
        n = simpledialog.askstring("Nuova Normativa", "Nome:")
        d = simpledialog.askstring("Nuova Normativa", "Descrizione:")
        
        if n:
            self.db.inserisci_normativa(n, d if d else "")
            self.refresh_norm_list()
    
    def edit_norm(self):
        """Modifica normativa"""
        s = self.tr_norm.selection()
        if not s:
            return
        
        nid, name = self.tr_norm.item(s)['values']
        new_n = simpledialog.askstring("Modifica Normativa", "Nome:", initialvalue=name)
        
        if new_n:
            self.db.aggiorna_normativa(nid, new_n, "")
            self.refresh_norm_list()
    
    def del_norm(self):
        """Elimina normativa"""
        s = self.tr_norm.selection()
        if s and messagebox.askyesno(
            "Attenzione",
            "Eliminare normativa e TUTTI i progetti collegati?"
        ):
            nid = self.tr_norm.item(s)['values'][0]
            self.db.elimina_normativa(nid)
            self.refresh_norm_list()
            self.tr_cat.delete(*self.tr_cat.get_children())
    
    def dup_norm(self):
        """Duplica normativa"""
        s = self.tr_norm.selection()
        if not s:
            return
        
        nid, name = self.tr_norm.item(s)['values']
        new_n = simpledialog.askstring(
            "Duplica Normativa",
            f"Nome nuova:",
            initialvalue=f"Copia di {name}"
        )
        
        if new_n:
            self.db.duplica_normativa(nid, new_n, f"Copia di {name}")
            self.refresh_norm_list()
    
    def open_cat_dialog(self, title, code="", macro=1, desc="", callback=None):
        """Dialog generico per voci catalogo"""
        d = tk.Toplevel(self)
        d.title(title)
        d.geometry("500x250")
        
        tk.Label(d, text="Codice:").pack(anchor='w', padx=10)
        e_cod = tk.Entry(d)
        e_cod.pack(fill='x', padx=10)
        e_cod.insert(0, code)
        
        tk.Label(d, text="Macro:").pack(anchor='w', padx=10)
        cb_mac = ttk.Combobox(
            d,
            values=[
                "1. Spese per l'esecuzione dell'intervento",
                "2. Somme a disposizione della S.A."
            ],
            state="readonly"
        )
        cb_mac.pack(fill='x', padx=10)
        cb_mac.current(0 if macro == 1 else 1)
        
        tk.Label(d, text="Descrizione:").pack(anchor='w', padx=10)
        e_desc = tk.Entry(d)
        e_desc.pack(fill='x', padx=10)
        e_desc.insert(0, desc)
        
        def on_ok():
            m_val = cb_mac.get()
            m_int = 1 if m_val.startswith("1") else 2
            callback(e_cod.get(), m_int, e_desc.get())
            d.destroy()
        
        tk.Button(d, text="Salva", command=on_ok).pack(pady=20)
    
    def new_cat_item_dialog(self):
        """Nuova voce catalogo"""
        if not hasattr(self, 'active_admin_norm_id'):
            return
        
        def cb(c, m, d):
            self.db.aggiorna_voce_catalogo_id(None, self.active_admin_norm_id, c, m, d)
            self.refresh_cat_admin()
        
        self.open_cat_dialog("Nuova Voce", callback=cb)
    
    def edit_cat_item_dialog(self, e):
        """Modifica voce catalogo"""
        s = self.tr_cat.selection()
        if not s:
            return
        
        vid, cod, mac, desc = self.tr_cat.item(s)['values']
        
        def cb(c, m, d):
            self.db.aggiorna_voce_catalogo_id(vid, self.active_admin_norm_id, c, m, d)
            self.refresh_cat_admin()
        
        self.open_cat_dialog("Modifica Voce", code=cod, macro=mac, desc=desc, callback=cb)
    
    def del_cat_item(self):
        """Elimina voce catalogo"""
        s = self.tr_cat.selection()
        if s and messagebox.askyesno("Conferma", "Eliminare voce dal catalogo?"):
            vid = self.tr_cat.item(s)['values'][0]
            self.db.elimina_voce_catalogo(vid)
            self.refresh_cat_admin()


# =============================================================================
# MAIN ENTRY POINT
# =============================================================================
if __name__ == "__main__":
    app = AppGestionale()
    app.mainloop()
# =============================================================================
