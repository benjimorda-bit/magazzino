"""
Gestionale Magazzino Tappeti
Applicazione desktop per la gestione di un magazzino tappeti.
Sviluppato con customtkinter, pandas, openpyxl e tksheet.
"""

import os
import re
from datetime import datetime, timedelta
from tkinter import messagebox, filedialog
import customtkinter as ctk
import pandas as pd
from tksheet import Sheet

# Configurazione del tema
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

# Costanti
FILE_DATABASE = "magazzino_tappeti.xlsx"
COLONNE_DATABASE = [
    "Stato", "Nr", "Nome", "Provenienza", "Misura", "MQ", "UM",
    "Epoca", "Qualita", "Disegno", "Colore", "Fornitore",
    "Costo", "Listino", "Prezzo Vendita Effettivo", "Data Vendita", "Note"
]
COLONNE_NUMERICHE = ["MQ", "Costo", "Listino", "Prezzo Vendita Effettivo"]
OPZIONI_STATO = ["Disponibile", "Venduto"]
OPZIONI_EPOCA = ["Nuovo", "Vecchio", "Antico"]
OPZIONI_QUALITA = ["Medio", "Fine", "Commerciale", "Extra"]
OPZIONI_DISEGNO = ["Etnico", "Classico", "Decorativo", "Moderno"]


class GestoreDatabase:
    """Classe per la gestione del database Excel."""
    
    def __init__(self, percorso_file: str):
        self.percorso_file = percorso_file
        self.df = self._carica_database()
    
    def _carica_database(self) -> pd.DataFrame:
        """Carica il database da file Excel o ne crea uno nuovo."""
        if os.path.exists(self.percorso_file):
            try:
                df = pd.read_excel(self.percorso_file, engine="openpyxl")
                # Verifica che tutte le colonne esistano
                for col in COLONNE_DATABASE:
                    if col not in df.columns:
                        df[col] = ""
                # Riordina le colonne
                df = df[COLONNE_DATABASE]
            except Exception as e:
                messagebox.showerror("Errore", f"Errore nel caricamento del database:\n{e}")
                df = pd.DataFrame(columns=COLONNE_DATABASE)
        else:
            df = pd.DataFrame(columns=COLONNE_DATABASE)
        
        # Forza le colonne numeriche
        df = self._forza_colonne_numeriche(df)
        return df
    
    def _forza_colonne_numeriche(self, df: pd.DataFrame) -> pd.DataFrame:
        """Converte le colonne monetarie e MQ in valori numerici."""
        for col in COLONNE_NUMERICHE:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        return df
    
    def salva_database(self) -> bool:
        """Salva il dataframe nel file Excel."""
        try:
            self.df.to_excel(self.percorso_file, index=False, engine="openpyxl")
            return True
        except Exception as e:
            messagebox.showerror("Errore", f"Errore nel salvataggio:\n{e}")
            return False
    
    def aggiungi_tappeto(self, dati: dict) -> bool:
        """Aggiunge un nuovo tappeto al database."""
        # Verifica codice univoco
        if dati["Nr"] in self.df["Nr"].values:
            messagebox.showwarning("Attenzione", "Il codice inserito esiste gia nel database.")
            return False
        
        # Crea nuova riga
        nuova_riga = pd.DataFrame([dati])
        nuova_riga = self._forza_colonne_numeriche(nuova_riga)
        self.df = pd.concat([self.df, nuova_riga], ignore_index=True)
        return self.salva_database()
    
    def elimina_tappeto(self, codice: str) -> bool:
        """Elimina un tappeto dal database tramite codice."""
        self.df = self.df[self.df["Nr"] != codice]
        return self.salva_database()
    
    def registra_vendita(self, codice: str, prezzo_effettivo: float) -> bool:
        """Registra la vendita di un tappeto."""
        idx = self.df[self.df["Nr"] == codice].index
        if len(idx) == 0:
            return False
        
        self.df.loc[idx, "Stato"] = "Venduto"
        self.df.loc[idx, "Prezzo Vendita Effettivo"] = prezzo_effettivo
        self.df.loc[idx, "Data Vendita"] = datetime.now().strftime("%Y-%m-%d %H:%M")
        return self.salva_database()
    
    def azzera_storico_vendite(self) -> bool:
        """Elimina tutti i tappeti venduti dal database."""
        self.df = self.df[self.df["Stato"] != "Venduto"]
        return self.salva_database()
    
    def filtra_dati(self, filtri: dict) -> pd.DataFrame:
        """Filtra i dati in base ai criteri specificati."""
        df_filtrato = self.df.copy()
        
        # Filtro per nome/codice
        if filtri.get("nome_codice"):
            termine = filtri["nome_codice"].lower()
            mask = (
                df_filtrato["Nr"].astype(str).str.lower().str.contains(termine, na=False) |
                df_filtrato["Nome"].astype(str).str.lower().str.contains(termine, na=False)
            )
            df_filtrato = df_filtrato[mask]
        
        # Filtro per provenienza
        if filtri.get("provenienza"):
            termine = filtri["provenienza"].lower()
            df_filtrato = df_filtrato[
                df_filtrato["Provenienza"].astype(str).str.lower().str.contains(termine, na=False)
            ]
        
        # Filtro per fornitore
        if filtri.get("fornitore"):
            termine = filtri["fornitore"].lower()
            df_filtrato = df_filtrato[
                df_filtrato["Fornitore"].astype(str).str.lower().str.contains(termine, na=False)
            ]
        
        # Filtro per UM
        if filtri.get("um"):
            termine = filtri["um"].lower()
            df_filtrato = df_filtrato[
                df_filtrato["UM"].astype(str).str.lower().str.contains(termine, na=False)
            ]
        
        # Filtro per stato
        if filtri.get("stato") and filtri["stato"] != "Tutti":
            df_filtrato = df_filtrato[df_filtrato["Stato"] == filtri["stato"]]
        
        return df_filtrato
    
    def ottieni_disponibili(self) -> pd.DataFrame:
        """Restituisce solo i tappeti disponibili."""
        return self.df[self.df["Stato"] == "Disponibile"]
    
    def ottieni_venduti_periodo(self, giorni: int = None) -> pd.DataFrame:
        """Restituisce i tappeti venduti in un determinato periodo."""
        venduti = self.df[self.df["Stato"] == "Venduto"].copy()
        
        if giorni is None or venduti.empty:
            return venduti
        
        # Converti la colonna Data Vendita
        venduti["Data Vendita"] = pd.to_datetime(venduti["Data Vendita"], errors="coerce")
        data_limite = datetime.now() - timedelta(days=giorni)
        
        return venduti[venduti["Data Vendita"] >= data_limite]


def calcola_mq_da_misura(misura: str) -> float:
    """Calcola i metri quadri dalla stringa misura (es. '299*74' o '299x74')."""
    if not misura:
        return 0.0
    
    # Cerca pattern numerico separato da x, X, * o spazi
    pattern = r"(\d+(?:[.,]\d+)?)\s*[xX*]\s*(\d+(?:[.,]\d+)?)"
    match = re.search(pattern, str(misura))
    
    if match:
        try:
            lunghezza = float(match.group(1).replace(",", "."))
            larghezza = float(match.group(2).replace(",", "."))
            return round((lunghezza * larghezza) / 10000, 2)
        except ValueError:
            return 0.0
    return 0.0


class FrameDashboard(ctk.CTkFrame):
    """Frame per la sezione Dashboard."""
    
    def __init__(self, master, gestore_db: GestoreDatabase):
        super().__init__(master, fg_color="transparent")
        self.gestore_db = gestore_db
        self._crea_interfaccia()
    
    def _crea_interfaccia(self):
        """Crea tutti gli elementi dell'interfaccia dashboard."""
        # Titolo
        titolo = ctk.CTkLabel(
            self, text="Dashboard", font=ctk.CTkFont(size=24, weight="bold")
        )
        titolo.pack(pady=(20, 10))
        
        # Frame resoconto magazzino
        frame_resoconto = ctk.CTkFrame(self)
        frame_resoconto.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(
            frame_resoconto, text="Resoconto Magazzino Attuale",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(pady=10)
        
        self.label_pezzi = ctk.CTkLabel(frame_resoconto, text="Pezzi Disponibili: 0")
        self.label_pezzi.pack(pady=5)
        
        self.label_costo = ctk.CTkLabel(frame_resoconto, text="Costo Totale: 0.00 EUR")
        self.label_costo.pack(pady=5)
        
        self.label_listino = ctk.CTkLabel(frame_resoconto, text="Listino Totale: 0.00 EUR")
        self.label_listino.pack(pady=(5, 15))
        
        # Pulsante azzeramento storico
        btn_azzera = ctk.CTkButton(
            frame_resoconto,
            text="Azzera Storico Vendite (Fine Anno)",
            fg_color="#dc3545",
            hover_color="#c82333",
            command=self._azzera_storico
        )
        btn_azzera.pack(pady=(0, 15))
        
        # Tabview statistiche temporali
        ctk.CTkLabel(
            self, text="Statistiche Vendite",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(pady=(20, 10))
        
        self.tabview = ctk.CTkTabview(self, width=500, height=200)
        self.tabview.pack(padx=20, pady=10, fill="x")
        
        self.tabs = {
            "Ultimo Mese": 30,
            "Ultimi 6 Mesi": 180,
            "Ultimo Anno": 365,
            "Sempre (Totale)": None
        }
        
        self.labels_stats = {}
        for nome_tab in self.tabs.keys():
            tab = self.tabview.add(nome_tab)
            
            frame_stats = ctk.CTkFrame(tab, fg_color="transparent")
            frame_stats.pack(fill="both", expand=True, padx=10, pady=10)
            
            self.labels_stats[nome_tab] = {
                "pezzi": ctk.CTkLabel(frame_stats, text="Pezzi Venduti: 0"),
                "incasso": ctk.CTkLabel(frame_stats, text="Incasso Totale: 0.00 EUR"),
                "costo": ctk.CTkLabel(frame_stats, text="Costo Merci: 0.00 EUR"),
                "utile": ctk.CTkLabel(frame_stats, text="Utile Netto: 0.00 EUR")
            }
            
            for label in self.labels_stats[nome_tab].values():
                label.pack(pady=3)
        
        self.aggiorna_dati()
    
    def _azzera_storico(self):
        """Azzera lo storico delle vendite dopo conferma."""
        risposta = messagebox.askyesno(
            "Conferma Azzeramento",
            "Sei sicuro di voler eliminare definitivamente tutti i tappeti venduti?\n\n"
            "Questa operazione non puo essere annullata."
        )
        if risposta:
            if self.gestore_db.azzera_storico_vendite():
                messagebox.showinfo("Completato", "Storico vendite azzerato con successo.")
                self.aggiorna_dati()
    
    def aggiorna_dati(self):
        """Aggiorna tutti i dati della dashboard."""
        # Resoconto magazzino
        disponibili = self.gestore_db.ottieni_disponibili()
        self.label_pezzi.configure(text=f"Pezzi Disponibili: {len(disponibili)}")
        self.label_costo.configure(text=f"Costo Totale: {disponibili['Costo'].sum():.2f} EUR")
        self.label_listino.configure(text=f"Listino Totale: {disponibili['Listino'].sum():.2f} EUR")
        
        # Statistiche temporali
        for nome_tab, giorni in self.tabs.items():
            venduti = self.gestore_db.ottieni_venduti_periodo(giorni)
            
            pezzi = len(venduti)
            incasso = venduti["Prezzo Vendita Effettivo"].sum() if not venduti.empty else 0
            costo = venduti["Costo"].sum() if not venduti.empty else 0
            utile = incasso - costo
            
            self.labels_stats[nome_tab]["pezzi"].configure(text=f"Pezzi Venduti: {pezzi}")
            self.labels_stats[nome_tab]["incasso"].configure(text=f"Incasso Totale: {incasso:.2f} EUR")
            self.labels_stats[nome_tab]["costo"].configure(text=f"Costo Merci: {costo:.2f} EUR")
            self.labels_stats[nome_tab]["utile"].configure(text=f"Utile Netto: {utile:.2f} EUR")


class FrameCercaElenco(ctk.CTkFrame):
    """Frame per la sezione Cerca/Elenco."""
    
    def __init__(self, master, gestore_db: GestoreDatabase):
        super().__init__(master, fg_color="transparent")
        self.gestore_db = gestore_db
        self.df_visualizzato = pd.DataFrame()
        self._crea_interfaccia()
    
    def _crea_interfaccia(self):
        """Crea tutti gli elementi dell'interfaccia cerca/elenco."""
        # Titolo
        ctk.CTkLabel(
            self, text="Cerca / Elenco Tappeti",
            font=ctk.CTkFont(size=24, weight="bold")
        ).pack(pady=(20, 10))
        
        # Frame filtri
        frame_filtri = ctk.CTkFrame(self)
        frame_filtri.pack(fill="x", padx=20, pady=10)
        
        # Prima riga filtri
        frame_riga1 = ctk.CTkFrame(frame_filtri, fg_color="transparent")
        frame_riga1.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(frame_riga1, text="Nome/Codice:").pack(side="left", padx=5)
        self.entry_nome_codice = ctk.CTkEntry(frame_riga1, width=150)
        self.entry_nome_codice.pack(side="left", padx=5)
        
        ctk.CTkLabel(frame_riga1, text="Provenienza:").pack(side="left", padx=5)
        self.entry_provenienza = ctk.CTkEntry(frame_riga1, width=120)
        self.entry_provenienza.pack(side="left", padx=5)
        
        ctk.CTkLabel(frame_riga1, text="Fornitore:").pack(side="left", padx=5)
        self.entry_fornitore = ctk.CTkEntry(frame_riga1, width=120)
        self.entry_fornitore.pack(side="left", padx=5)
        
        # Seconda riga filtri
        frame_riga2 = ctk.CTkFrame(frame_filtri, fg_color="transparent")
        frame_riga2.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(frame_riga2, text="UM:").pack(side="left", padx=5)
        self.entry_um = ctk.CTkEntry(frame_riga2, width=100)
        self.entry_um.pack(side="left", padx=5)
        
        ctk.CTkLabel(frame_riga2, text="Stato:").pack(side="left", padx=5)
        self.combo_stato = ctk.CTkComboBox(
            frame_riga2, values=["Tutti"] + OPZIONI_STATO, width=120
        )
        self.combo_stato.set("Tutti")
        self.combo_stato.pack(side="left", padx=5)
        
        ctk.CTkButton(
            frame_riga2, text="Applica Filtri", command=self._applica_filtri, width=120
        ).pack(side="left", padx=15)
        
        ctk.CTkButton(
            frame_riga2, text="Azzera", command=self._azzera_filtri, width=80
        ).pack(side="left", padx=5)
        
        # Frame per la tabella (tksheet)
        frame_tabella = ctk.CTkFrame(self)
        frame_tabella.pack(fill="both", expand=True, padx=20, pady=10)
        
        self.sheet = Sheet(
            frame_tabella,
            headers=COLONNE_DATABASE,
            show_x_scrollbar=True,
            show_y_scrollbar=True,
            width=900,
            height=400
        )
        self.sheet.pack(fill="both", expand=True)
        self.sheet.enable_bindings("single_select", "row_select", "column_width_resize")
        
        # Frame barra inferiore
        frame_inferiore = ctk.CTkFrame(self)
        frame_inferiore.pack(fill="x", padx=20, pady=10)
        
        # Pulsanti a sinistra
        frame_pulsanti = ctk.CTkFrame(frame_inferiore, fg_color="transparent")
        frame_pulsanti.pack(side="left")
        
        ctk.CTkButton(
            frame_pulsanti, text="Elimina Selezionato",
            fg_color="#dc3545", hover_color="#c82333",
            command=self._elimina_selezionato
        ).pack(side="left", padx=5)
        
        ctk.CTkButton(
            frame_pulsanti, text="Esporta per Stampa",
            command=self._esporta_stampa
        ).pack(side="left", padx=5)
        
        # Contatori a destra
        frame_contatori = ctk.CTkFrame(frame_inferiore, fg_color="transparent")
        frame_contatori.pack(side="right")
        
        self.label_totale_pezzi = ctk.CTkLabel(frame_contatori, text="Pezzi: 0")
        self.label_totale_pezzi.pack(side="left", padx=15)
        
        self.label_totale_costo = ctk.CTkLabel(frame_contatori, text="Costo: 0.00 EUR")
        self.label_totale_costo.pack(side="left", padx=15)
        
        self.label_totale_listino = ctk.CTkLabel(frame_contatori, text="Listino: 0.00 EUR")
        self.label_totale_listino.pack(side="left", padx=15)
        
        # Carica dati iniziali
        self.aggiorna_dati()
    
    def _applica_filtri(self):
        """Applica i filtri e aggiorna la visualizzazione."""
        filtri = {
            "nome_codice": self.entry_nome_codice.get().strip(),
            "provenienza": self.entry_provenienza.get().strip(),
            "fornitore": self.entry_fornitore.get().strip(),
            "um": self.entry_um.get().strip(),
            "stato": self.combo_stato.get()
        }
        self.df_visualizzato = self.gestore_db.filtra_dati(filtri)
        self._aggiorna_tabella()
    
    def _azzera_filtri(self):
        """Azzera tutti i filtri."""
        self.entry_nome_codice.delete(0, "end")
        self.entry_provenienza.delete(0, "end")
        self.entry_fornitore.delete(0, "end")
        self.entry_um.delete(0, "end")
        self.combo_stato.set("Tutti")
        self.aggiorna_dati()
    
    def _aggiorna_tabella(self):
        """Aggiorna la visualizzazione della tabella."""
        # Prepara i dati formattati
        df_display = self.df_visualizzato.copy()
        
        # Formatta i valori monetari
        for col in ["Costo", "Listino", "Prezzo Vendita Effettivo"]:
            if col in df_display.columns:
                df_display[col] = df_display[col].apply(
                    lambda x: f"{x:.2f} EUR" if pd.notna(x) and x != 0 else ""
                )
        
        # Formatta MQ
        if "MQ" in df_display.columns:
            df_display["MQ"] = df_display["MQ"].apply(
                lambda x: f"{x:.2f}" if pd.notna(x) and x != 0 else ""
            )
        
        # Converti in lista di liste
        dati = df_display.values.tolist()
        
        # Aggiorna sheet
        self.sheet.set_sheet_data(dati)
        self.sheet.set_all_column_widths(100)
        
        # Aggiorna contatori
        pezzi = len(self.df_visualizzato)
        costo_tot = self.df_visualizzato["Costo"].sum()
        listino_tot = self.df_visualizzato["Listino"].sum()
        
        self.label_totale_pezzi.configure(text=f"Pezzi: {pezzi}")
        self.label_totale_costo.configure(text=f"Costo: {costo_tot:.2f} EUR")
        self.label_totale_listino.configure(text=f"Listino: {listino_tot:.2f} EUR")
    
    def _elimina_selezionato(self):
        """Elimina il tappeto selezionato."""
        selezione = self.sheet.get_currently_selected()
        if not selezione:
            messagebox.showwarning("Attenzione", "Seleziona un tappeto da eliminare.")
            return
        
        # Ottieni l'indice della riga
        riga_idx = selezione.row if hasattr(selezione, 'row') else selezione[0]
        
        if riga_idx < 0 or riga_idx >= len(self.df_visualizzato):
            messagebox.showwarning("Attenzione", "Selezione non valida.")
            return
        
        codice = self.df_visualizzato.iloc[riga_idx]["Nr"]
        nome = self.df_visualizzato.iloc[riga_idx]["Nome"]
        
        risposta = messagebox.askyesno(
            "Conferma Eliminazione",
            f"Sei sicuro di voler eliminare definitivamente il tappeto:\n\n"
            f"Codice: {codice}\nNome: {nome}"
        )
        
        if risposta:
            if self.gestore_db.elimina_tappeto(codice):
                messagebox.showinfo("Completato", "Tappeto eliminato con successo.")
                self.aggiorna_dati()
    
    def _esporta_stampa(self):
        """Esporta il dataframe filtrato per la stampa."""
        if self.df_visualizzato.empty:
            messagebox.showwarning("Attenzione", "Nessun dato da esportare.")
            return
        
        percorso = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("File Excel", "*.xlsx")],
            title="Salva elenco per stampa"
        )
        
        if percorso:
            try:
                self.df_visualizzato.to_excel(percorso, index=False, engine="openpyxl")
                messagebox.showinfo("Completato", f"File esportato con successo:\n{percorso}")
            except Exception as e:
                messagebox.showerror("Errore", f"Errore durante l'esportazione:\n{e}")
    
    def aggiorna_dati(self):
        """Aggiorna i dati dal database."""
        self.df_visualizzato = self.gestore_db.df.copy()
        self._aggiorna_tabella()


class FrameAggiungiTappeto(ctk.CTkFrame):
    """Frame per la sezione Aggiungi Tappeto."""
    
    def __init__(self, master, gestore_db: GestoreDatabase, callback_aggiornamento):
        super().__init__(master, fg_color="transparent")
        self.gestore_db = gestore_db
        self.callback_aggiornamento = callback_aggiornamento
        self._crea_interfaccia()
    
    def _crea_interfaccia(self):
        """Crea tutti gli elementi dell'interfaccia aggiungi tappeto."""
        # Titolo
        ctk.CTkLabel(
            self, text="Aggiungi Nuovo Tappeto",
            font=ctk.CTkFont(size=24, weight="bold")
        ).pack(pady=(20, 20))
        
        # Frame form
        frame_form = ctk.CTkFrame(self)
        frame_form.pack(fill="both", expand=True, padx=40, pady=10)
        
        # Dizionario per memorizzare i widget di input
        self.inputs = {}
        
        # Configurazione campi
        campi = [
            ("Nr", "entry", True),
            ("Nome", "entry", True),
            ("Provenienza", "entry", False),
            ("Misura", "entry", True),
            ("UM", "entry", False),
            ("Epoca", "combo", False),
            ("Qualita", "combo", False),
            ("Disegno", "combo", False),
            ("Colore", "entry", False),
            ("Fornitore", "entry", False),
            ("Costo", "entry", False),
            ("Listino", "entry", False),
            ("Note", "textbox", False),
        ]
        
        # Crea i campi in griglia
        for i, (nome, tipo, obbligatorio) in enumerate(campi):
            riga = i // 2
            colonna = i % 2
            
            frame_campo = ctk.CTkFrame(frame_form, fg_color="transparent")
            frame_campo.grid(row=riga, column=colonna, padx=20, pady=10, sticky="ew")
            
            label_text = f"{nome}*:" if obbligatorio else f"{nome}:"
            ctk.CTkLabel(frame_campo, text=label_text, width=100, anchor="w").pack(anchor="w")
            
            if tipo == "entry":
                widget = ctk.CTkEntry(frame_campo, width=250)
            elif tipo == "combo":
                if nome == "Epoca":
                    valori = OPZIONI_EPOCA
                elif nome == "Qualita":
                    valori = OPZIONI_QUALITA
                else:  # Disegno
                    valori = OPZIONI_DISEGNO
                widget = ctk.CTkComboBox(frame_campo, values=valori, width=250)
                widget.set("")
            else:  # textbox
                widget = ctk.CTkTextbox(frame_campo, width=250, height=80)
            
            widget.pack(anchor="w")
            self.inputs[nome] = widget
        
        frame_form.columnconfigure(0, weight=1)
        frame_form.columnconfigure(1, weight=1)
        
        # Pulsante salva
        ctk.CTkButton(
            self, text="Salva Tappeto", command=self._salva_tappeto,
            width=200, height=40
        ).pack(pady=30)
    
    def _salva_tappeto(self):
        """Salva il nuovo tappeto nel database."""
        # Ottieni valori
        nr = self.inputs["Nr"].get().strip()
        nome = self.inputs["Nome"].get().strip()
        misura = self.inputs["Misura"].get().strip()
        
        # Validazione campi obbligatori
        if not nr or not nome or not misura:
            messagebox.showwarning(
                "Campi Obbligatori",
                "I campi Nr, Nome e Misura sono obbligatori."
            )
            return
        
        # Validazione campi numerici
        costo_str = self.inputs["Costo"].get().strip()
        listino_str = self.inputs["Listino"].get().strip()
        
        try:
            costo = float(costo_str.replace(",", ".")) if costo_str else 0
        except ValueError:
            messagebox.showwarning("Errore", "Il campo Costo deve contenere un valore numerico.")
            return
        
        try:
            listino = float(listino_str.replace(",", ".")) if listino_str else 0
        except ValueError:
            messagebox.showwarning("Errore", "Il campo Listino deve contenere un valore numerico.")
            return
        
        # Calcola MQ
        mq = calcola_mq_da_misura(misura)
        
        # Ottieni note
        note = self.inputs["Note"].get("1.0", "end-1c").strip()
        
        # Prepara dati
        dati = {
            "Stato": "Disponibile",
            "Nr": nr,
            "Nome": nome,
            "Provenienza": self.inputs["Provenienza"].get().strip(),
            "Misura": misura,
            "MQ": mq,
            "UM": self.inputs["UM"].get().strip(),
            "Epoca": self.inputs["Epoca"].get(),
            "Qualita": self.inputs["Qualita"].get(),
            "Disegno": self.inputs["Disegno"].get(),
            "Colore": self.inputs["Colore"].get().strip(),
            "Fornitore": self.inputs["Fornitore"].get().strip(),
            "Costo": costo,
            "Listino": listino,
            "Prezzo Vendita Effettivo": 0,
            "Data Vendita": "",
            "Note": note
        }
        
        if self.gestore_db.aggiungi_tappeto(dati):
            messagebox.showinfo("Completato", "Tappeto aggiunto con successo.")
            self._pulisci_form()
            self.callback_aggiornamento()
    
    def _pulisci_form(self):
        """Pulisce tutti i campi del form."""
        for nome, widget in self.inputs.items():
            if isinstance(widget, ctk.CTkEntry):
                widget.delete(0, "end")
            elif isinstance(widget, ctk.CTkComboBox):
                widget.set("")
            elif isinstance(widget, ctk.CTkTextbox):
                widget.delete("1.0", "end")


class FrameRegistraVendita(ctk.CTkFrame):
    """Frame per la sezione Registra Vendita."""
    
    def __init__(self, master, gestore_db: GestoreDatabase, callback_aggiornamento):
        super().__init__(master, fg_color="transparent")
        self.gestore_db = gestore_db
        self.callback_aggiornamento = callback_aggiornamento
        self._crea_interfaccia()
    
    def _crea_interfaccia(self):
        """Crea tutti gli elementi dell'interfaccia registra vendita."""
        # Titolo
        ctk.CTkLabel(
            self, text="Registra Vendita",
            font=ctk.CTkFont(size=24, weight="bold")
        ).pack(pady=(20, 30))
        
        # Frame form
        frame_form = ctk.CTkFrame(self)
        frame_form.pack(padx=40, pady=20)
        
        # Selezione tappeto
        ctk.CTkLabel(
            frame_form, text="Seleziona Tappeto Disponibile:",
            font=ctk.CTkFont(size=14)
        ).pack(pady=(20, 5))
        
        self.combo_tappeti = ctk.CTkComboBox(frame_form, values=[], width=400)
        self.combo_tappeti.pack(pady=10)
        
        # Prezzo vendita effettivo
        ctk.CTkLabel(
            frame_form, text="Prezzo Vendita Effettivo (EUR):",
            font=ctk.CTkFont(size=14)
        ).pack(pady=(20, 5))
        
        self.entry_prezzo = ctk.CTkEntry(frame_form, width=200)
        self.entry_prezzo.pack(pady=10)
        
        # Pulsante conferma
        ctk.CTkButton(
            frame_form, text="Conferma Vendita", command=self._conferma_vendita,
            width=200, height=40
        ).pack(pady=30)
        
        # Aggiorna lista tappeti
        self.aggiorna_lista_tappeti()
    
    def aggiorna_lista_tappeti(self):
        """Aggiorna la lista dei tappeti disponibili."""
        disponibili = self.gestore_db.ottieni_disponibili()
        
        if disponibili.empty:
            self.combo_tappeti.configure(values=["Nessun tappeto disponibile"])
            self.combo_tappeti.set("Nessun tappeto disponibile")
        else:
            opzioni = [
                f"{row['Nr']} - {row['Nome']}"
                for _, row in disponibili.iterrows()
            ]
            self.combo_tappeti.configure(values=opzioni)
            if opzioni:
                self.combo_tappeti.set(opzioni[0])
    
    def _conferma_vendita(self):
        """Conferma e registra la vendita."""
        selezione = self.combo_tappeti.get()
        
        if not selezione or selezione == "Nessun tappeto disponibile":
            messagebox.showwarning("Attenzione", "Seleziona un tappeto da vendere.")
            return
        
        # Estrai il codice dalla selezione
        codice = selezione.split(" - ")[0].strip()
        
        # Valida il prezzo
        prezzo_str = self.entry_prezzo.get().strip()
        if not prezzo_str:
            messagebox.showwarning("Attenzione", "Inserisci il prezzo di vendita effettivo.")
            return
        
        try:
            prezzo = float(prezzo_str.replace(",", "."))
        except ValueError:
            messagebox.showwarning("Errore", "Il prezzo deve essere un valore numerico.")
            return
        
        if prezzo < 0:
            messagebox.showwarning("Errore", "Il prezzo non puo essere negativo.")
            return
        
        # Conferma
        risposta = messagebox.askyesno(
            "Conferma Vendita",
            f"Confermi la vendita del tappeto:\n\n"
            f"Codice: {codice}\n"
            f"Prezzo: {prezzo:.2f} EUR"
        )
        
        if risposta:
            if self.gestore_db.registra_vendita(codice, prezzo):
                messagebox.showinfo("Completato", "Vendita registrata con successo.")
                self.entry_prezzo.delete(0, "end")
                self.aggiorna_lista_tappeti()
                self.callback_aggiornamento()


class AppMagazzino(ctk.CTk):
    """Classe principale dell'applicazione."""
    
    def __init__(self):
        super().__init__()
        
        # Configurazione finestra
        self.title("Gestionale Magazzino Tappeti")
        self.geometry("1200x700")
        self.minsize(1000, 600)
        
        # Inizializza il gestore database
        self.gestore_db = GestoreDatabase(FILE_DATABASE)
        
        # Crea interfaccia
        self._crea_layout()
    
    def _crea_layout(self):
        """Crea il layout principale con sidebar e area contenuto."""
        # Configura griglia principale
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        # Sidebar
        self.sidebar = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_rowconfigure(5, weight=1)
        
        # Logo/Titolo sidebar
        ctk.CTkLabel(
            self.sidebar, text="Magazzino\nTappeti",
            font=ctk.CTkFont(size=20, weight="bold")
        ).grid(row=0, column=0, padx=20, pady=(20, 30))
        
        # Pulsanti navigazione
        self.btn_dashboard = ctk.CTkButton(
            self.sidebar, text="Dashboard",
            command=lambda: self._mostra_sezione("dashboard")
        )
        self.btn_dashboard.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        
        self.btn_cerca = ctk.CTkButton(
            self.sidebar, text="Cerca / Elenco",
            command=lambda: self._mostra_sezione("cerca")
        )
        self.btn_cerca.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        
        self.btn_aggiungi = ctk.CTkButton(
            self.sidebar, text="Aggiungi Tappeto",
            command=lambda: self._mostra_sezione("aggiungi")
        )
        self.btn_aggiungi.grid(row=3, column=0, padx=20, pady=10, sticky="ew")
        
        self.btn_vendita = ctk.CTkButton(
            self.sidebar, text="Registra Vendita",
            command=lambda: self._mostra_sezione("vendita")
        )
        self.btn_vendita.grid(row=4, column=0, padx=20, pady=10, sticky="ew")
        
        # Area contenuto principale
        self.frame_contenuto = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_contenuto.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        self.frame_contenuto.grid_columnconfigure(0, weight=1)
        self.frame_contenuto.grid_rowconfigure(0, weight=1)
        
        # Crea le sezioni
        self.sezioni = {}
        
        self.sezioni["dashboard"] = FrameDashboard(self.frame_contenuto, self.gestore_db)
        self.sezioni["cerca"] = FrameCercaElenco(self.frame_contenuto, self.gestore_db)
        self.sezioni["aggiungi"] = FrameAggiungiTappeto(
            self.frame_contenuto, self.gestore_db, self._aggiorna_tutte_sezioni
        )
        self.sezioni["vendita"] = FrameRegistraVendita(
            self.frame_contenuto, self.gestore_db, self._aggiorna_tutte_sezioni
        )
        
        # Posiziona tutte le sezioni (nascoste inizialmente)
        for sezione in self.sezioni.values():
            sezione.grid(row=0, column=0, sticky="nsew")
        
        # Mostra dashboard di default
        self._mostra_sezione("dashboard")
    
    def _mostra_sezione(self, nome_sezione: str):
        """Mostra la sezione specificata e nasconde le altre."""
        for nome, sezione in self.sezioni.items():
            if nome == nome_sezione:
                sezione.tkraise()
                # Aggiorna i dati della sezione visualizzata
                if hasattr(sezione, "aggiorna_dati"):
                    sezione.aggiorna_dati()
                if hasattr(sezione, "aggiorna_lista_tappeti"):
                    sezione.aggiorna_lista_tappeti()
    
    def _aggiorna_tutte_sezioni(self):
        """Callback per aggiornare tutte le sezioni dopo modifiche."""
        for sezione in self.sezioni.values():
            if hasattr(sezione, "aggiorna_dati"):
                sezione.aggiorna_dati()
            if hasattr(sezione, "aggiorna_lista_tappeti"):
                sezione.aggiorna_lista_tappeti()


def main():
    """Funzione principale di avvio."""
    app = AppMagazzino()
    app.mainloop()


if __name__ == "__main__":
    main()
